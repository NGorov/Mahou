using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using NLog;

namespace Mahou {
	class MMain {

		private static readonly Logger log = LogManager.GetCurrentClassLogger();
		static System.Windows.Forms.Timer hookWatchdog;
		#region DLLs
		[DllImport("user32.dll")]
		public static extern uint RegisterWindowMessage(string message);
		#endregion
		#region Prevent another instance variables
		public const string appGUid = "ec511418-1d57-4dbe-a0c3-c6022b33735b";
		public static uint ao = RegisterWindowMessage("AlderyOpenedMahou!");
		#endregion
		#region All Main variables, arrays etc.
		public static List<KMHook.YuKey> c_word = new List<KMHook.YuKey>();
		public static List<List<KMHook.YuKey>> c_words = new List<List<KMHook.YuKey>>();
		public static IntPtr _hookID = IntPtr.Zero;
		public static IntPtr _mouse_hookID = IntPtr.Zero;
		public static KMHook.LowLevelProc _proc = KMHook.HookCallback;
		public static KMHook.LowLevelProc _mouse_proc = KMHook.MouseHookCallback;
		public static Locales.Locale[] locales = Locales.AllList();
		public static Configs MyConfs = new Configs();
		public static MahouForm mahou;
		public static List<string> lcnmid = new List<string>();
		public static string[] UI = { };
		public static string[] TTips = { };
		public static string[] Msgs = { };
		#endregion
		[STAThread] //DO NOT REMOVE THIS
		public static void Main(string[] args) {
			LogHelper.ConfigureNlog();
			log.Info("Program start");

			Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
			Application.ThreadException += (sender, e) => {
				log.Fatal(e.Exception, "Unhandled UI thread exception");
			};
			AppDomain.CurrentDomain.UnhandledException += (sender, e) => {
				log.Fatal(e.ExceptionObject as Exception, "Unhandled AppDomain exception (isTerminating={0})", e.IsTerminating);
			};
			TaskScheduler.UnobservedTaskException += (sender, e) => {
				log.Error(e.Exception, "Unobserved Task exception");
				e.SetObserved();
			};

			using(var mutex = new Mutex(false, "Global\\" + appGUid)) {
				log.Info("Mutex created");
				if(!mutex.WaitOne(0, false)) {
					KMHook.PostMessage((IntPtr)0xffff, ao, 0, 0);
					return;
				}
			if(locales.Length < 2) {
				Locales.IfLessThan2(exitOnFail: true);
			} else {
					mahou = new MahouForm();
					InitLanguage();
					//Refreshes icon text language at startup
					mahou.icon.RefreshText(MMain.UI[44], MMain.UI[42], MMain.UI[43]);
					KMHook.ReInitSnippets();
					Application.EnableVisualStyles();
					if(args.Length != 0)
						if(args[0] == "_!_updated_!_") {
							mahou.ToggleVisibility();
							MessageBox.Show(Msgs[0], Msgs[1], MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					StartHook();
					StartHookWatchdog();
					//for first run, add your locale 1 & locale 2 to settings
					if(MyConfs.Read("Locales", "locale1Lang") == "" && MyConfs.Read("Locales", "locale2Lang") == "") {
						MyConfs.Write("Locales", "locale1uId", locales[0].uId.ToString());
						MyConfs.Write("Locales", "locale2uId", locales[1].uId.ToString());
						MyConfs.Write("Locales", "locale1Lang", locales[0].Lang);
						MyConfs.Write("Locales", "locale2Lang", locales[1].Lang);
					}
					try {
						Application.Run();
					} catch(Exception ex) {
						log.Fatal(ex, "Global error handler caught the exception in app");
					}
					if(hookWatchdog != null) {
						hookWatchdog.Stop();
						hookWatchdog.Dispose();
					}
					StopHook();
					log.Info("Program shutdown");
				}
			}
		}
		public static void InitLanguage() {
			if(MyConfs.Read("Locales", "LANGUAGE") == "RU") {
				UI = Translation.UIRU;
				TTips = Translation.ToolTipsRU;
				Msgs = Translation.MessagesRU;
			} else if(MyConfs.Read("Locales", "LANGUAGE") == "EN") {
				UI = Translation.UIEN;
				TTips = Translation.ToolTipsEN;
				Msgs = Translation.MessagesEN;
			}
		}
		#region Actions with hooks
		public static void StartHook() {
			if(!CheckHook()) {
				return;
			}
			log.Info("Installing hooks");
			_mouse_hookID = KMHook.SetHook(_mouse_proc, (int)KMHook.KMMessages.WH_MOUSE_LL);
			_hookID = KMHook.SetHook(_proc, (int)KMHook.KMMessages.WH_KEYBOARD_LL);
			if(_hookID == IntPtr.Zero || _mouse_hookID == IntPtr.Zero) {
				log.Warn("Hook installation returned null handle (kbd={0}, mouse={1})", _hookID, _mouse_hookID);
			}
			Thread.Sleep(10);
		}
		public static void StopHook() {
			if(CheckHook()) {
				return;
			}
			log.Info("Removing hooks");
			KMHook.UnhookWindowsHookEx(_hookID);
			KMHook.UnhookWindowsHookEx(_mouse_hookID);
			_hookID = _mouse_hookID = IntPtr.Zero;
			Thread.Sleep(10);
		}
		public static void ReinstallHooks() {
			log.Warn("Reinstalling hooks (watchdog detected dead hooks)");
			StopHook();
			StartHook();
		}
		public static bool CheckHook() {
			return _hookID == IntPtr.Zero;
		}
		static void StartHookWatchdog() {
			hookWatchdog = new System.Windows.Forms.Timer();
			hookWatchdog.Interval = 30000;
			hookWatchdog.Tick += (_, __) => {
				try {
					if(_hookID == IntPtr.Zero || _mouse_hookID == IntPtr.Zero) {
						log.Warn("Hook watchdog: hook handle is zero, reinstalling");
						ReinstallHooks();
					}
				} catch(Exception ex) {
					log.Error(ex, "Hook watchdog error");
				}
			};
			hookWatchdog.Start();
		}
		#endregion
	}
}