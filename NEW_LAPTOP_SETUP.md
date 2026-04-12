# New Laptop Setup

Use this checklist on a completely fresh laptop to continue development on `AJ Tools`.

## 1. Install core software

1. Install all Windows updates.
2. Install Microsoft Project desktop.
3. Install Visual Studio 2022.
4. In the Visual Studio Installer, make sure these workloads are installed:
   - `.NET desktop development`
   - `Office/SharePoint development`
5. Install Git.

## 2. Install required runtimes and tools

1. Install the VSTO runtime from Microsoft:
   - https://www.microsoft.com/download/details.aspx?id=105522
2. Verify one of these exists:
   - `C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe`
   - `C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe`
3. Install the .NET SDK:
   - https://dotnet.microsoft.com/en-us/download
4. Open PowerShell and verify:
   ```powershell
   dotnet --info
   ```
5. Install WiX v4:
   ```powershell
   dotnet tool install --global wix --version 4.0.5
   ```
6. Verify WiX:
   ```powershell
   wix --version
   ```

## 3. Clone the repo

1. Clone the repo to a normal source path, for example:
   ```powershell
   C:\Users\<your-user>\source\repos\Arian Jahandarfards MS Project Add-in
   ```
2. Open the solution:
   - `Arian Jahandarfards MS Project Add-in.slnx`

## 4. First build check

1. In Visual Studio, build `Debug`.
2. In Visual Studio, build `Release`.
3. If either build fails, check that:
   - Microsoft Project is installed
   - the VSTO runtime is installed
   - the Office/SharePoint development workload is installed

## 5. Debug mode workflow

Use this when you want Microsoft Project to load the latest code from the repo instead of the installed Program Files version.

1. Close Microsoft Project.
2. Open PowerShell in the repo root.
3. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Register-DebugAddIn.ps1 -Configuration Debug -Build
   ```
4. Open Microsoft Project.
5. Test the latest local code.

## 6. Installed mode workflow

Use this when you want to test the real installed add-in or the updater flow.

1. Close Microsoft Project.
2. Open PowerShell in the repo root.
3. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Use-InstalledAddIn.ps1
   ```
4. Open Microsoft Project normally.

## 7. Build a fresh local installer

Use this when testing the local installer package.

1. Open PowerShell in the repo root.
2. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Build-LocalInstaller.ps1
   ```
3. This produces fresh local installer files in:
   - `AJSetup\bin\Release\AJSetup.exe`
   - `AJSetup\bin\Release\AJAddIn.msi`

## 8. GitHub and release notes

1. Sign into Visual Studio with the GitHub-connected account you use for this repo.
2. Pull the latest `master`.
3. Before any real release:
   - keep `AssemblyVersion`
   - `AssemblyFileVersion`
   - and `ApplicationVersion`
     all in sync

## 9. Important project rules

1. Do not click `Publish Now`.
2. GitHub Actions handles publishing.
3. Use `Build-LocalInstaller.ps1` for local MSI tests.
4. Use `Register-DebugAddIn.ps1` for feature work.
5. Use `Use-InstalledAddIn.ps1` before testing the real updater flow.

## 10. Quick sanity test

1. Build `Debug`.
2. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Register-DebugAddIn.ps1 -Configuration Debug -Build
   ```
3. Open Microsoft Project and confirm the AJ ribbon appears.
4. Build `Release`.
5. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Build-LocalInstaller.ps1
   ```
6. Confirm the installer outputs are created.
