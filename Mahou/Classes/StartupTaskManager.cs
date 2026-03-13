using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;

namespace Mahou {
	internal static class StartupTaskManager {
		internal const string TaskName = "MahouAutoStart+";

		const int TASK_TRIGGER_LOGON = 9;
		const int TASK_ACTION_EXEC = 0;
		const int TASK_CREATE_OR_UPDATE = 6;
		const int TASK_LOGON_INTERACTIVE_TOKEN = 3;
		const int TASK_RUNLEVEL_HIGHEST = 1;
		const int TASK_INSTANCES_IGNORE_NEW = 2;
		const int ERROR_CANCELLED = 1223;
		const int HRESULT_FILE_NOT_FOUND = unchecked((int)0x80070002);

		public static string GetCurrentExecutablePath() {
			var executablePath = Environment.ProcessPath;
			if (string.IsNullOrWhiteSpace(executablePath))
				executablePath = Application.ExecutablePath;
			if (string.IsNullOrWhiteSpace(executablePath)) {
				using (var process = Process.GetCurrentProcess()) {
					if (process.MainModule != null)
						executablePath = process.MainModule.FileName;
				}
			}
			if (string.IsNullOrWhiteSpace(executablePath))
				throw new InvalidOperationException("Unable to resolve the current executable path.");
			return Path.GetFullPath(executablePath);
		}

		public static bool EnsureCurrentUserLogonTask() {
			try {
				var executablePath = GetCurrentExecutablePath();
				Logging.Log("Ensuring startup task '" + TaskName + "' for executable: " + executablePath);
				if (!File.Exists(executablePath)) {
					Logging.Log("Startup task was not created because the executable does not exist: " + executablePath, 1);
					return false;
				}
				if (IsCurrentUserLogonTaskRegistered()) {
					Logging.Log("Startup task already exists and matches the current executable.");
					return true;
				}

				string error;
				if (TryEnsureWithCom(executablePath, out error)) {
					Logging.Log("Startup task created via Task Scheduler COM API.");
					return true;
				}
				Logging.Log("Task Scheduler COM registration failed: " + error, 2);

				if (TryEnsureWithSchtasks(executablePath, out error)) {
					Logging.Log("Startup task created via schtasks.exe fallback.");
					return true;
				}
				Logging.Log("Failed to create startup task: " + error, 1);
				return false;
			} catch (Exception ex) {
				Logging.Log("Unexpected error while creating the startup task: " + ex.Message + "\n" + ex.StackTrace, 1);
				return false;
			}
		}

		public static bool RemoveCurrentUserLogonTask() {
			try {
				string error;
				if (TryRemoveWithCom(out error)) {
					Logging.Log("Startup task removed via Task Scheduler COM API.");
					return true;
				}
				Logging.Log("Task Scheduler COM removal failed: " + error, 2);

				if (TryRemoveWithSchtasks(out error)) {
					Logging.Log("Startup task removed via schtasks.exe fallback.");
					return true;
				}
				Logging.Log("Failed to remove startup task: " + error, 1);
				return false;
			} catch (Exception ex) {
				Logging.Log("Unexpected error while removing the startup task: " + ex.Message + "\n" + ex.StackTrace, 1);
				return false;
			}
		}

		public static bool IsCurrentUserLogonTaskRegistered() {
			try {
				var executablePath = GetCurrentExecutablePath();
				object serviceObject = null;
				object rootFolderObject = null;
				object registeredTaskObject = null;
				try {
					Connect(out serviceObject, out rootFolderObject);
					dynamic rootFolder = rootFolderObject;
					try {
						registeredTaskObject = rootFolder.GetTask(TaskName);
					} catch (COMException ex) {
						if (ex.ErrorCode == HRESULT_FILE_NOT_FOUND)
							return false;
						throw;
					}
					return IsExpectedTaskConfiguration(registeredTaskObject, executablePath);
				} finally {
					ReleaseComObject(registeredTaskObject);
					ReleaseComObject(rootFolderObject);
					ReleaseComObject(serviceObject);
				}
			} catch (Exception ex) {
				Logging.Log("Failed to inspect the startup task: " + ex.Message + "\n" + ex.StackTrace, 1);
				return false;
			}
		}

		static bool TryEnsureWithCom(string executablePath, out string error) {
			error = string.Empty;
			object serviceObject = null;
			object rootFolderObject = null;
			object taskDefinitionObject = null;
			object logonTriggerObject = null;
			object execActionObject = null;
			object registeredTaskObject = null;
			try {
				Connect(out serviceObject, out rootFolderObject);

				dynamic service = serviceObject;
				dynamic rootFolder = rootFolderObject;
				dynamic taskDefinition = service.NewTask(0);
				taskDefinitionObject = taskDefinition;

				var userSid = GetCurrentUserSid();
				taskDefinition.RegistrationInfo.Description = "Starts Mahou with highest privileges at user logon.";
				taskDefinition.Settings.Enabled = true;
				taskDefinition.Settings.StartWhenAvailable = true;
				taskDefinition.Settings.AllowDemandStart = true;
				taskDefinition.Settings.MultipleInstances = TASK_INSTANCES_IGNORE_NEW;
				taskDefinition.Settings.DisallowStartIfOnBatteries = false;
				taskDefinition.Settings.StopIfGoingOnBatteries = false;
				taskDefinition.Settings.ExecutionTimeLimit = "PT0S";
				taskDefinition.Settings.Hidden = false;
				taskDefinition.Settings.RunOnlyIfNetworkAvailable = false;
				taskDefinition.Settings.Priority = 7;
				taskDefinition.Principal.UserId = userSid;
				taskDefinition.Principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN;
				taskDefinition.Principal.RunLevel = TASK_RUNLEVEL_HIGHEST;

				dynamic logonTrigger = taskDefinition.Triggers.Create(TASK_TRIGGER_LOGON);
				logonTriggerObject = logonTrigger;
				logonTrigger.Enabled = true;
				logonTrigger.UserId = userSid;

				dynamic execAction = taskDefinition.Actions.Create(TASK_ACTION_EXEC);
				execActionObject = execAction;
				execAction.Path = executablePath;
				execAction.WorkingDirectory = GetWorkingDirectory(executablePath);

				registeredTaskObject = rootFolder.RegisterTaskDefinition(
					TaskName,
					taskDefinition,
					TASK_CREATE_OR_UPDATE,
					null,
					null,
					TASK_LOGON_INTERACTIVE_TOKEN,
					null
				);

				if (!IsExpectedTaskConfiguration(registeredTaskObject, executablePath)) {
					error = "The task was registered but validation failed.";
					return false;
				}
				return true;
			} catch (Exception ex) {
				error = ex.Message;
				return false;
			} finally {
				ReleaseComObject(registeredTaskObject);
				ReleaseComObject(execActionObject);
				ReleaseComObject(logonTriggerObject);
				ReleaseComObject(taskDefinitionObject);
				ReleaseComObject(rootFolderObject);
				ReleaseComObject(serviceObject);
			}
		}

		static bool TryRemoveWithCom(out string error) {
			error = string.Empty;
			object serviceObject = null;
			object rootFolderObject = null;
			try {
				Connect(out serviceObject, out rootFolderObject);
				dynamic rootFolder = rootFolderObject;
				try {
					rootFolder.DeleteTask(TaskName, 0);
				} catch (COMException ex) {
					if (ex.ErrorCode == HRESULT_FILE_NOT_FOUND)
						return true;
					throw;
				}
				return !TaskExists();
			} catch (Exception ex) {
				error = ex.Message;
				return false;
			} finally {
				ReleaseComObject(rootFolderObject);
				ReleaseComObject(serviceObject);
			}
		}

		static bool TryEnsureWithSchtasks(string executablePath, out string error) {
			error = string.Empty;
			var xmlPath = string.Empty;
			try {
				xmlPath = WriteTaskXml(executablePath);
				int exitCode;
				string output;
				if (!RunSchtasks("/Create /TN \"" + TaskName + "\" /XML \"" + xmlPath + "\" /F", out exitCode, out output, out error))
					return false;
				if (exitCode != 0) {
					error = "schtasks.exe returned exit code " + exitCode + ". " + output.Trim();
					return false;
				}
				if (!IsCurrentUserLogonTaskRegistered()) {
					error = "schtasks.exe returned success but task validation failed.";
					return false;
				}
				return true;
			} finally {
				try {
					if (!string.IsNullOrWhiteSpace(xmlPath) && File.Exists(xmlPath))
						File.Delete(xmlPath);
				} catch (Exception ex) {
					Logging.Log("Failed to delete temporary startup task XML: " + ex.Message, 2);
				}
			}
		}

		static bool TryRemoveWithSchtasks(out string error) {
			error = string.Empty;
			int exitCode;
			string output;
			if (!RunSchtasks("/Delete /TN \"" + TaskName + "\" /F", out exitCode, out output, out error))
				return false;
			if (exitCode != 0 && TaskExists()) {
				error = "schtasks.exe returned exit code " + exitCode + ". " + output.Trim();
				return false;
			}
			return !TaskExists();
		}

		static void Connect(out object serviceObject, out object rootFolderObject) {
			serviceObject = null;
			rootFolderObject = null;

			var serviceType = Type.GetTypeFromProgID("Schedule.Service");
			if (serviceType == null)
				throw new NotSupportedException("The Windows Task Scheduler COM service is not available.");

			serviceObject = Activator.CreateInstance(serviceType);
			dynamic service = serviceObject;
			service.Connect();
			rootFolderObject = service.GetFolder("\\");
		}

		static bool IsExpectedTaskConfiguration(object registeredTaskObject, string executablePath) {
			object definitionObject = null;
			object actionObject = null;
			try {
				dynamic registeredTask = registeredTaskObject;
				dynamic definition = registeredTask.Definition;
				definitionObject = definition;

				if (Convert.ToInt32(definition.Principal.RunLevel) != TASK_RUNLEVEL_HIGHEST)
					return false;
				if (Convert.ToInt32(definition.Principal.LogonType) != TASK_LOGON_INTERACTIVE_TOKEN)
					return false;
				if (!Convert.ToBoolean(definition.Settings.Enabled))
					return false;
				if (Convert.ToInt32(definition.Actions.Count) < 1)
					return false;

				dynamic action = definition.Actions.Item(1);
				actionObject = action;
				if (!PathEquals(Convert.ToString(action.Path), executablePath))
					return false;

				for (var i = 1; i <= Convert.ToInt32(definition.Triggers.Count); i++) {
					object triggerObject = null;
					try {
						triggerObject = definition.Triggers.Item(i);
						dynamic trigger = triggerObject;
						if (Convert.ToInt32(trigger.Type) != TASK_TRIGGER_LOGON)
							continue;
						if (!Convert.ToBoolean(trigger.Enabled))
							continue;
						var triggerUserId = Convert.ToString(trigger.UserId);
						if (string.IsNullOrWhiteSpace(triggerUserId) || StringComparer.OrdinalIgnoreCase.Equals(triggerUserId, GetCurrentUserSid()))
							return true;
					} finally {
						ReleaseComObject(triggerObject);
					}
				}

				return false;
			} finally {
				ReleaseComObject(actionObject);
				ReleaseComObject(definitionObject);
			}
		}

		static string WriteTaskXml(string executablePath) {
			var userSid = GetCurrentUserSid();
			var builder = new StringBuilder();
			builder.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-16""?>");
			builder.AppendLine(@"<Task version=""1.4"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">");
			builder.AppendLine(@"  <RegistrationInfo>");
			builder.AppendLine(@"    <Description>Starts Mahou with highest privileges at user logon.</Description>");
			builder.AppendLine(@"    <Author>Mahou</Author>");
			builder.AppendLine(@"  </RegistrationInfo>");
			builder.AppendLine(@"  <Triggers>");
			builder.AppendLine(@"    <LogonTrigger>");
			builder.AppendLine(@"      <Enabled>true</Enabled>");
			builder.AppendLine(@"      <UserId>" + SecurityElement.Escape(userSid) + @"</UserId>");
			builder.AppendLine(@"    </LogonTrigger>");
			builder.AppendLine(@"  </Triggers>");
			builder.AppendLine(@"  <Principals>");
			builder.AppendLine(@"    <Principal id=""Author"">");
			builder.AppendLine(@"      <UserId>" + SecurityElement.Escape(userSid) + @"</UserId>");
			builder.AppendLine(@"      <LogonType>InteractiveToken</LogonType>");
			builder.AppendLine(@"      <RunLevel>HighestAvailable</RunLevel>");
			builder.AppendLine(@"    </Principal>");
			builder.AppendLine(@"  </Principals>");
			builder.AppendLine(@"  <Settings>");
			builder.AppendLine(@"    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>");
			builder.AppendLine(@"    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>");
			builder.AppendLine(@"    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>");
			builder.AppendLine(@"    <AllowHardTerminate>true</AllowHardTerminate>");
			builder.AppendLine(@"    <StartWhenAvailable>true</StartWhenAvailable>");
			builder.AppendLine(@"    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>");
			builder.AppendLine(@"    <AllowStartOnDemand>true</AllowStartOnDemand>");
			builder.AppendLine(@"    <Enabled>true</Enabled>");
			builder.AppendLine(@"    <Hidden>false</Hidden>");
			builder.AppendLine(@"    <WakeToRun>false</WakeToRun>");
			builder.AppendLine(@"    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>");
			builder.AppendLine(@"    <Priority>7</Priority>");
			builder.AppendLine(@"  </Settings>");
			builder.AppendLine(@"  <Actions Context=""Author"">");
			builder.AppendLine(@"    <Exec>");
			builder.AppendLine(@"      <Command>" + SecurityElement.Escape(executablePath) + @"</Command>");
			builder.AppendLine(@"      <WorkingDirectory>" + SecurityElement.Escape(GetWorkingDirectory(executablePath)) + @"</WorkingDirectory>");
			builder.AppendLine(@"    </Exec>");
			builder.AppendLine(@"  </Actions>");
			builder.AppendLine(@"</Task>");

			var xmlPath = Path.Combine(Path.GetTempPath(), "MahouStartupTask-" + Guid.NewGuid().ToString("N") + ".xml");
			File.WriteAllText(xmlPath, builder.ToString(), Encoding.Unicode);
			return xmlPath;
		}

		static bool RunSchtasks(string arguments, out int exitCode, out string output, out string error) {
			exitCode = -1;
			output = string.Empty;
			error = string.Empty;

			var startInfo = new ProcessStartInfo {
				FileName = "schtasks.exe",
				Arguments = arguments,
				CreateNoWindow = true,
				WindowStyle = ProcessWindowStyle.Hidden
			};

			var needsElevation = !IsProcessElevated();
			if (needsElevation) {
				startInfo.UseShellExecute = true;
				startInfo.Verb = "runas";
			} else {
				startInfo.UseShellExecute = false;
				startInfo.RedirectStandardOutput = true;
				startInfo.RedirectStandardError = true;
			}

			try {
				using (var process = new Process()) {
					process.StartInfo = startInfo;
					if (!process.Start()) {
						error = "Failed to start schtasks.exe.";
						return false;
					}
					if (!startInfo.UseShellExecute)
						output = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
					process.WaitForExit();
					exitCode = process.ExitCode;
					return true;
				}
			} catch (Win32Exception ex) {
				if (ex.NativeErrorCode == ERROR_CANCELLED) {
					error = "The UAC prompt was cancelled.";
					return false;
				}
				error = ex.Message;
				return false;
			} catch (Exception ex) {
				error = ex.Message;
				return false;
			}
		}

		static bool TaskExists() {
			object serviceObject = null;
			object rootFolderObject = null;
			object registeredTaskObject = null;
			try {
				Connect(out serviceObject, out rootFolderObject);
				dynamic rootFolder = rootFolderObject;
				try {
					registeredTaskObject = rootFolder.GetTask(TaskName);
					return true;
				} catch (COMException ex) {
					if (ex.ErrorCode == HRESULT_FILE_NOT_FOUND)
						return false;
					throw;
				}
			} finally {
				ReleaseComObject(registeredTaskObject);
				ReleaseComObject(rootFolderObject);
				ReleaseComObject(serviceObject);
			}
		}

		static bool PathEquals(string left, string right) {
			if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
				return false;
			try {
				return StringComparer.OrdinalIgnoreCase.Equals(
					Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
					Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
				);
			} catch {
				return StringComparer.OrdinalIgnoreCase.Equals(left, right);
			}
		}

		static string GetCurrentUserSid() {
			using (var identity = WindowsIdentity.GetCurrent()) {
				if (identity.User != null && !string.IsNullOrWhiteSpace(identity.User.Value))
					return identity.User.Value;
				if (!string.IsNullOrWhiteSpace(identity.Name))
					return identity.Name;
			}
			return Environment.UserDomainName + "\\" + Environment.UserName;
		}

		static bool IsProcessElevated() {
			using (var identity = WindowsIdentity.GetCurrent()) {
				var principal = new WindowsPrincipal(identity);
				return principal.IsInRole(WindowsBuiltInRole.Administrator);
			}
		}

		static string GetWorkingDirectory(string executablePath) {
			var workingDirectory = Path.GetDirectoryName(executablePath);
			if (string.IsNullOrWhiteSpace(workingDirectory))
				workingDirectory = AppDomain.CurrentDomain.BaseDirectory;
			return workingDirectory;
		}

		static void ReleaseComObject(object comObject) {
			if (comObject != null && Marshal.IsComObject(comObject))
				Marshal.FinalReleaseComObject(comObject);
		}
	}
}
