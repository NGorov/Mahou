# How to upload a release for Chocolatey

The Chocolatey package downloads `Release_x86_x64.zip` from GitHub **Assets**
(not from the release description text). The URL used in the install script is:

```
https://github.com/NGorov/Mahou3/releases/download/v3.0/Release_x86_x64.zip
```

## Steps to attach the binary as a release asset

1. Go to https://github.com/NGorov/Mahou3/releases/tag/v3.0
2. Click **Edit release**
3. Scroll down to the **"Attach binaries by dropping them here or selecting them"** area
4. Drag-and-drop (or select) `Release_x86_x64.zip` — do **not** paste it into the description text field
5. Click **Update release**

After this the file will appear in the **Assets** section and the direct
download link will work.

## After uploading a new ZIP

If you rebuild the application and upload a different `Release_x86_x64.zip`:

1. Recalculate SHA256:
   ```powershell
   Get-FileHash Release_x86_x64.zip -Algorithm SHA256
   ```
2. Update the `checksum` value in `Chocolatey/tools/chocolateyInstall.ps1`
3. Update the checksum in `Chocolatey/tools/VERIFICATION.txt`
4. Repack the Chocolatey package:
   ```powershell
   cd Chocolatey
   choco pack
   ```

## Packaging note

The `Mahou.nuspec` file lists only `tools\**` in `<files>` so that no archive
or other stray files beside `tools` are included in the `.nupkg`. The installer
is always fetched from GitHub at install time.
