# Full Project Handoff

This document is a detailed handoff for the `AJ Tools` / `Arian Jahandarfard's Tools` Microsoft Project VSTO add-in project.

It is intended to let another chat, another laptop, or another developer continue the work without re-discovering the architecture and the recent debugging history.

## 1. Project identity

- Product name: `AJ Tools`
- Developer: `Arian Jahandarfard`
- Main VSTO namespace: `Arian_Jahandarfards_MS_Project_Add_in`
- COM/registry add-in name in some places: `ArianJahandarfardsAddIn`
- Main solution:
  - `Arian Jahandarfards MS Project Add-in.slnx`

## 2. Solution structure

- `Arian Jahandarfards MS Project Add-in\`
  - Main Microsoft Project VSTO add-in project
  - Produces:
    - `Arian Jahandarfards MS Project Add-in.dll`
    - `.vsto`
    - manifest files

- `AJSetup\`
  - Standalone installer/updater EXE
  - Handles:
    - waiting for Project to close
    - MSI install/uninstall
    - VSTO cleanup
    - VSTO registration
    - success/failure UI

- `AJToolsInstaller\`
  - WiX v4 MSI packaging project
  - Main file:
    - `AJToolsInstaller\Package.wxs`

- `scripts\`
  - Local dev and installer workflow scripts

## 3. Key files

- Main add-in:
  - `Arian Jahandarfards MS Project Add-in\AJRibbon.cs`
  - `Arian Jahandarfards MS Project Add-in\AJRibbon.Designer.cs`
  - `Arian Jahandarfards MS Project Add-in\ThisAddIn.cs`
  - `Arian Jahandarfards MS Project Add-in\AJUpdater.cs`
  - `Arian Jahandarfards MS Project Add-in\AJProjectLinker.cs`
  - `Arian Jahandarfards MS Project Add-in\AJProjectLinkerForm.cs`
  - `Arian Jahandarfards MS Project Add-in\Properties\AssemblyInfo.cs`
  - `Arian Jahandarfards MS Project Add-in\Arian Jahandarfards MS Project Add-in.csproj`

- Installer:
  - `AJSetup\Form1.cs`
  - `AJSetup\Program.cs`
  - `AJSetup\AJSetup.csproj`

- Packaging:
  - `AJToolsInstaller\Package.wxs`

- Local scripts:
  - `scripts\Build-LocalInstaller.ps1`
  - `scripts\Register-DebugAddIn.ps1`
  - `scripts\Use-InstalledAddIn.ps1`
  - `scripts\Start-DebugProject.ps1`

- Setup docs:
  - `NEW_LAPTOP_SETUP.md`
  - this file: `FULL_PROJECT_HANDOFF.md`

## 4. Current version state

At the time of this handoff:

- `AssemblyVersion`: `2.4.0.5`
- `AssemblyFileVersion`: `2.4.0.5`
- `ApplicationVersion` in `.csproj`: `2.4.0.5`

These values should continue to stay in sync.

## 5. Release/update architecture

The intended shipping architecture is:

1. Code is changed locally in Visual Studio.
2. Developer commits and pushes to GitHub.
3. GitHub Actions builds Release output.
4. GitHub Actions builds MSI using WiX.
5. GitHub Actions uploads release artifacts to Azure Blob Storage.
6. GitHub Actions updates hosted `version.json`.
7. Installed clients click `Update for Check` in Microsoft Project.
8. Add-in checks hosted `version.json`.
9. If newer version exists:
   - Project shows a simple update prompt
   - launches `AJSetup.exe`
   - Project closes
   - `AJSetup.exe` downloads/installs update

Important design decision:

- The live update UI inside Microsoft Project was simplified.
- The complex update experience now belongs to `AJSetup.exe`, not to a custom in-Project form.

## 6. Important changes made during this session

### 6.1 Updater flow was simplified

`AJUpdater.cs` was changed so the active updater flow now:

- checks hosted version info
- shows simple Windows message boxes in Project
- launches external `AJSetup.exe`
- schedules Project shutdown safely

This avoided the previous freeze-prone architecture.

### 6.2 The fake "Thread was aborted" popup was addressed

Earlier, Project was being closed too aggressively inside the updater path, which caused:

- update succeeded
- but user still saw a misleading `Thread was aborted` error

That path was changed so Project shutdown is delayed slightly after launching the installer.

### 6.3 Stale local MSI issue was found and fixed

This was one of the most important discoveries.

Problem:

- Rebuilding the add-in project did not automatically refresh the local MSI next to `AJSetup.exe`
- so local reinstalls kept deploying old code
- this made it look like code changes were "not taking effect"

Fixes:

- `scripts\Build-LocalInstaller.ps1` was added
- `AJSetup\Form1.cs` now warns if the local MSI beside `AJSetup.exe` is older than the EXE

### 6.4 Visual Studio VSTO auto-registration was disabled

In:

- `Arian Jahandarfards MS Project Add-in\Arian Jahandarfards MS Project Add-in.csproj`

the VSTO auto register/unregister targets were no-op'd:

- `RegisterOfficeAddin`
- `UnregisterOfficeAddin`
- `RemoveOfficeAddInSecurity`

Why:

- Visual Studio was reintroducing stale dev registrations
- that caused build/clean confusion
- and Project sometimes loaded the wrong copy of the add-in

### 6.5 Debug-mode vs installed-mode workflow was introduced

This was added because the project needs 2 intentionally different behaviors:

- `Debug mode`
  - Project should load the latest local build from the repo

- `Installed mode`
  - Project should load the installed Program Files version

Scripts added:

- `scripts\Register-DebugAddIn.ps1`
- `scripts\Use-InstalledAddIn.ps1`
- `scripts\Start-DebugProject.ps1`

### 6.6 Installer cleanup was expanded

`AJSetup\Form1.cs` was updated to clean more aggressively:

- VSTO metadata
- VSTO inclusion entries
- current-user Project add-in keys
- current-user uninstall entries

Later in the session, machine-wide MSI cleanup was improved too.

### 6.7 Old AJ Tools MSI entries were found in Apps & features

It was discovered that multiple stale `AJ Tools` entries existed:

- `2.4.0.2`
- `2.4.0.3`
- `2.4.0.4`
- `2.4.0.5`

These came from stale MSI uninstall records.

Fix:

- `AJSetup\Form1.cs` was changed to enumerate installed AJ Tools MSI product codes and uninstall existing AJ Tools packages by product code before reinstalling
- it also now cleans machine-wide uninstall entries

### 6.8 Progress bar animation changed

The progress bar in the installer was changed from an all-over shimmer style to a more obvious left-to-right moving segment.

## 7. Files currently changed in the working tree

At the time this handoff was written, `git status --short` showed:

- modified:
  - `AJSetup/Form1.cs`
  - `Arian Jahandarfards MS Project Add-in/AJRibbon.Designer.cs`
  - `Arian Jahandarfards MS Project Add-in/AJRibbon.cs`
  - `Arian Jahandarfards MS Project Add-in/Arian Jahandarfards MS Project Add-in.csproj`
  - `Arian Jahandarfards MS Project Add-in/ThisAddIn.cs`

- untracked/new:
  - `Arian Jahandarfards MS Project Add-in/AJProjectLinker.cs`
  - `Arian Jahandarfards MS Project Add-in/AJProjectLinkerForm.cs`
  - `NEW_LAPTOP_SETUP.md`
  - `scripts/Register-DebugAddIn.ps1`
  - `scripts/Start-DebugProject.ps1`
  - `scripts/Use-InstalledAddIn.ps1`

This means there is still uncommitted work in progress.

## 8. Current Project Linker work

This is the major active feature still in progress.

### Goal

User wants a `Project Linker` feature between Microsoft Excel and Microsoft Project.

Desired behavior:

- `Excel` mode:
  - click a row in Excel
  - add-in finds matching task in Project
  - prefer `Unique ID`
  - fall back to task `Name`

- `Excel + Project` mode:
  - Excel -> Project still works
  - clicking a task in Project also finds/selects matching row in Excel

### Earlier approach

Originally, a floating panel with 2 checkboxes was attempted:

- `Microsoft Excel`
- `Microsoft Project + Excel`

This repeatedly failed visually:

- first checkbox did not reliably render
- layout kept looking wrong on the user’s machine

### Current approach

The design was changed to a ribbon dropdown model.

Now the intent is:

- ribbon dropdown:
  - `Excel`
  - `Excel + Project`
  - `Off`

- floating panel:
  - no checkboxes
  - only status display:
    - title
    - current mode
    - green `On` / red `Off`
    - live status line

This is a better fit for Office UI and avoids the repeated checkbox rendering problem.

### Current code state for Project Linker

`AJProjectLinker.cs`:

- now has internal mode enum:
  - `Off`
  - `Excel`
  - `ExcelAndProject`

- adds:
  - `ActivateMode(AJProjectLinkerMode mode)`
  - status text improvements like:
    - `Excel row 6 is linked to UID 10.`

`AJRibbon.Designer.cs` and `AJRibbon.cs`:

- `Project Linker` is being converted from a single ribbon button to a dropdown menu
- menu choices currently coded:
  - `Excel`
  - `Excel + Project`
  - `Off`

`AJProjectLinkerForm.cs`:

- now rewritten again as a tiny status-only panel
- should no longer depend on checkboxes at all

### Important current status

This Project Linker dropdown/status-panel refactor has been coded and compiled successfully in Debug, but it still needs user validation inside Microsoft Project.

It was not fully visually validated by the user before this handoff was written.

## 9. Debug workflow on the current machine

For feature development:

1. Close Microsoft Project.
2. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Register-DebugAddIn.ps1 -Configuration Debug -Build
   ```
3. Open Microsoft Project.
4. Project should load the local Debug build.

For installed/update testing:

1. Close Microsoft Project.
2. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Use-InstalledAddIn.ps1
   ```
3. Open Project normally.

## 10. Release/local installer workflow

For a fresh local installer:

1. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Build-LocalInstaller.ps1
   ```

That script:

- rebuilds the solution
- copies `AJSetup.exe` and logo into add-in release packaging inputs
- builds MSI with WiX
- outputs:
  - `AJSetup\bin\Release\AJSetup.exe`
  - `AJSetup\bin\Release\AJAddIn.msi`

## 11. Important technical gotchas

### 11.1 Do not trust local reinstall tests unless MSI is fresh

This was a major source of confusion earlier.

If the local MSI is stale, code changes can appear to do nothing.

### 11.2 Do not rely on Visual Studio VSTO auto-registration

That behavior was intentionally disabled because it caused stale registration chaos.

### 11.3 Debug mode and installed mode must be treated as separate

If both are active or mixed, Project may load the wrong thing or duplicate behavior.

### 11.4 Do not click Publish Now

GitHub Actions handles release publishing.

## 12. What was already proven to work

These things were successfully verified during the session:

- updater can detect a newer hosted version
- updater can close Project
- `AJSetup.exe` can install the new version
- reopened Project reports the new version correctly
- latest-version check works after update

Versions tested through this process included:

- `2.4.0.3`
- `2.4.0.4`
- `2.4.0.5`

## 13. What still needs validation

### High priority

1. Validate the new Project Linker ribbon dropdown UI in Project.
2. Validate the small status-only Project Linker panel.
3. Validate:
   - `Excel` mode
   - `Excel + Project` mode
   - `Off` mode

### Medium priority

4. Confirm stale `AJ Tools` entries stop accumulating after the new installer cleanup logic.
5. Confirm debug/install switching works cleanly on a fresh machine.

## 14. Recommended next steps

1. Move to the new laptop using `NEW_LAPTOP_SETUP.md`.
2. Build the solution there.
3. Use debug mode to validate the new Project Linker dropdown and panel.
4. Once the UI is confirmed, commit the Project Linker changes.
5. Then continue feature development or release testing from that cleaner baseline.

## 15. One-line summary

The updater and installer architecture were stabilized, stale local/install registration problems were largely solved, the project now has explicit debug-vs-installed workflows, and the main unfinished work is validating the newly redesigned Project Linker dropdown + status panel UI.
