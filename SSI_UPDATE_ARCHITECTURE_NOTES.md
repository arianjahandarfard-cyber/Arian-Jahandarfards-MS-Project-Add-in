## SSI update architecture notes

Date captured: 2026-04-23

### AJ Tools current architecture

- Installer is per-machine MSI.
- Install target is `C:\Program Files (x86)\AJTools\`.
- Project add-in registration is written to `HKLM`.
- Updater launches `AJSetup.exe` with elevation (`runas`).
- `AJSetup.exe` itself requests `requireAdministrator`.

Repo evidence:

- `AJToolsInstaller\Package.wxs`
- `AJSetup\app.manifest`
- `AJSetup\Form1.cs`
- `Arian Jahandarfards MS Project Add-in\AJUpdater.cs`

### SSI Tools architecture observed on this machine

SSI Project add-in is not loading from `Program Files`.

Observed Project add-in registration:

- Registry key:
  `HKCU\Software\Microsoft\Office\MS Project\Addins\SSIToolsForMSProject`
- Manifest value:
  `file:///C:/Users/arian/AppData/Local/SSI_Tools/Application Files/SSIToolsForMSProject_14_0_19_0/SSIToolsForMSProject.vsto|vstolocal`

Observed install root:

- `C:\Users\arian\AppData\Local\SSI_Tools\`

Observed versioned runtime folder:

- `C:\Users\arian\AppData\Local\SSI_Tools\Application Files\SSIToolsForMSProject_14_0_19_0\`

Observed uninstall entry:

- `HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{8128EB31-5BC6-4647-9A16-FE65C179EC54}_is1`
- Display name:
  `SSI Tools Bundled Package version 14.0.19`
- Install location:
  `C:\Users\arian\AppData\Local\SSI_Tools\`

Important interpretation:

- The runtime used by Microsoft Project is clearly per-user (`HKCU` + `LocalAppData`).
- The initial bundled installer may still have written an uninstall entry under machine uninstall records.
- No SSI Project add-in registration was found under `HKLM`.
- No SSI Windows service was found.
- No SSI scheduled task was found.

### SSI local script evidence

File:

- `C:\Users\arian\AppData\Local\SSI_Tools\AddSSIToolsForMSProjectRibbon.ps1`

What it does:

- Treats the base path as user-specific Local AppData.
- Looks under `Application Files`.
- Finds the newest folder matching `SSIToolsForMSProject_*`.
- Finds the `.vsto` inside that folder.
- Writes the Project add-in registration to:
  `HKCU:\Software\Microsoft\Office\MS Project\Addins\SSIToolsForMSProject`

This is strong evidence that SSI updates by switching Project to the newest versioned folder in a user-writable location.

### SSI update feed evidence

Observed SSI log:

- `C:\Users\arian\AppData\Local\SSI_Tools\Logs\Log_2026-04-23-11-34-57.log`

Important lines:

- Update URL:
  `https://ssitools.com/ssitoolsupdates/SSIToolsForMSProject/NoClickOnceSSIToolsForMSProjectUpdateInfo.json`
- Returned version:
  `14.1.5`
- Build date:
  `04-22-2026`
- Download artifact:
  `https://ssitools.com/ssitoolsupdates/SSIToolsForMSProject/ssiproject14-1-5.zip`

Important interpretation:

- SSI appears to use a custom updater.
- The updater is checking a JSON manifest.
- The payload appears to be a ZIP, not an MSI.
- The JSON name includes `NoClickOnce`, which suggests they are intentionally not relying on the standard ClickOnce update path for this flow.

### Working theory

SSI likely does something close to this:

1. Keep the Project add-in files under `LocalAppData`.
2. Download a ZIP for the new version.
3. Extract it into a new versioned folder under `Application Files`.
4. Repoint or refresh the `HKCU` add-in manifest path to the newest `.vsto`.
5. Leave `Program Files` and `HKLM` out of the normal update path.

That design would normally avoid UAC during updates because the updater is only touching user-writable files and current-user registry keys.

### Live update experiment to run

If SSI prompts for an update again while Project is open, a controlled before/after capture should answer the remaining questions:

1. Snapshot `HKCU\Software\Microsoft\Office\MS Project\Addins\SSIToolsForMSProject`.
2. Snapshot `C:\Users\arian\AppData\Local\SSI_Tools\` file timestamps and version folders.
3. Click `Update`.
4. Watch whether:
   - a ZIP is downloaded
   - a new `SSIToolsForMSProject_*` folder appears
   - the manifest path changes
   - only selected files like `DECM.json` are patched via `FileMap`
5. Re-check the SSI log for exact steps.

### Goal for AJ Tools

To behave more like SSI, AJ Tools would likely need to move away from:

- per-machine MSI install
- `Program Files (x86)\AJTools`
- `HKLM` Project add-in registration
- an always-elevated updater

And toward:

- user-writable install root
- versioned application folder layout
- `HKCU` registration
- ZIP or file-copy based updates

### High-value unanswered questions

These are the most important questions to answer during the live SSI update run:

1. Which process actually performs the update?
   - The Project add-in itself
   - a helper EXE
   - PowerShell
   - Windows Script Host
   - another bundled updater

2. Where is the ZIP staged?
   - `%TEMP%`
   - under `C:\Users\arian\AppData\Local\SSI_Tools`
   - another hidden cache path

3. Is the update full side-by-side deployment or partial patching?
   - The ZIP contains a full new runtime folder
   - It also contains top-level support files like `DECM.json`, `SSI_Tools.mpp`, and `SSI_EVTools.mpp`
   - Need to confirm whether the live updater copies both sets or only some of them

4. How does the switchover happen?
   - Does it rewrite `HKCU\Software\Microsoft\Office\MS Project\Addins\SSIToolsForMSProject\Manifest`
   - Does it also update `VSTO\SolutionMetadata`
   - Does it call a local ribbon-registration script

5. What is the rollback behavior?
   - If extraction fails
   - if Project is still holding a file lock
   - if trust/registration fails
   - if the new version crashes on load

6. What gets preserved versus replaced?
   - User configuration
   - mutable content
   - support templates
   - logs
   - custom data files

7. How does cleanup work?
   - Are old version folders retained
   - are they deleted later
   - does uninstall remove all versions or only the current layout

### Important SSI-specific observations

These matter because they point to the architecture SSI is really using:

- The update feed is custom JSON, not MSI.
- The payload is a ZIP, not an MSI.
- The VSTO manifest still exists and is signed.
- The deployment manifest has `install="false"`.
- The application manifest requests `asInvoker`, not admin.
- The runtime folder is versioned (`SSIToolsForMSProject_14_1_5_0`).
- The Project add-in registration is in `HKCU`, not `HKLM`.
- The installed bundle was created with Inno Setup.

Important interpretation:

- SSI appears to be using VSTO manifests and code signing, but not standard ClickOnce updating.
- The JSON name `NoClickOnceSSIToolsForMSProjectUpdateInfo.json` strongly suggests they intentionally replaced the standard ClickOnce updater with a custom ZIP-based updater.
- They kept the good parts of the VSTO deployment/trust model while moving update delivery to a user-writable side-by-side folder system.

### Why this matters for AJ Tools

It may not require a total architecture rewrite in one shot.

A likely migration path for AJ Tools could be:

1. Keep a small bootstrapper/installer for first-time setup.
2. Move the runtime add-in files to a user-writable location.
3. Register the add-in under `HKCU`.
4. Use versioned folders for each release.
5. Update by downloading a ZIP and extracting a new version folder.
6. Atomically repoint the manifest path to the new version.
7. Leave mutable user/config data outside the versioned runtime folder.

That would let the initial install stay enterprise-friendly while removing UAC from normal updates.

### Outside-the-box design questions for AJ Tools

These are easy to miss but matter a lot in corporate environments:

1. Where should mutable data live?
   - Anything the user edits should not live inside the versioned runtime folder.
   - Otherwise updates will overwrite it or force merge logic.

2. What is the true source of truth for "active version"?
   - Registry manifest path
   - local version file
   - installer metadata
   - support bundle manifest

3. How do you make cutover atomic?
   - Extract into a new folder first
   - validate signatures and hashes
   - only then update the registry pointer

4. What is the rollback story?
   - Keep the previous version folder
   - revert the manifest path if launch verification fails

5. What does "update complete" mean?
   - Files copied
   - registry updated
   - VSTO trust valid
   - ribbon loads in Project

6. How do you survive Office resiliency?
   - Office can disable add-ins after crashes
   - `LoadBehavior` and `AddinsData` need to be observed, not assumed

7. What happens when certificates change?
   - If you rotate signing certs, the trust and update story may break unless you plan for publisher continuity and timestamping

8. What happens in locked-down environments?
   - AppLocker / WDAC
   - TLS interception
   - proxy auth
   - blocked execution from `AppData`
   - Group Policy restrictions on VSTO add-ins

9. What happens on multi-user or profile-reset systems?
   - Roaming profiles
   - FSLogix / VDI
   - profile recreation
   - non-persistent desktops

10. How do you separate runtime, content, and licensing?
   - Runtime binaries
   - shipped templates/examples
   - user settings
   - license/cache data
   - logs

11. How do you support repair without admin?
   - Re-register `HKCU` keys
   - rebuild VSTO metadata
   - verify active folder contents

12. How do you make support easier?
   - Log every update step
   - log current/next version
   - log source URL
   - log extracted folder
   - log registry cutover
   - log rollback reason if failure occurs

### Elevated-session wish list

If the shell is relaunched as Administrator, the next capture should add:

1. `netsh trace` or `pktmon` packet/ETW capture during update
2. deeper process and network correlation
3. Windows event log review for:
   - AppLocker
   - Code Integrity
   - VSTO/runtime load failures
4. optional Procmon capture if Sysinternals Procmon is installed

This is not strictly required to copy SSI's architecture, but it would reduce guesswork around the exact live update transaction.

### Final offline-update confirmation

After building a local offline update folder and testing it in SSI's `Check for Updates` dialog, the results were:

1. The path field expects a folder, not the zip file itself.
2. SSI appends `NoClickOnceSSIToolsForMSProjectUpdateInfo.json` inside that folder and reads it.
3. The folder `C:\Users\arian\AppData\Local\Temp\ssi-offline-update-14.1.5` was accepted.
4. SSI successfully read the local update manifest and recognized version `14.1.5`.
5. The update was still blocked during `InstallUpdates()` by SSI's Maintenance and Support entitlement check.

This means SSI's offline mode is not a true "install anything locally with no server validation" path. It is:

1. Offline update discovery from a local manifest folder
2. Offline/local zip source for the payload
3. Separate online M&S/license validation before install proceeds

Relevant local log evidence:

- `Attempting to check for updates locally from: file:///C:/Users/arian/AppData/Local/Temp/ssi-offline-update-14.1.5/NoClickOnceSSIToolsForMSProjectUpdateInfo.json`
- `Successfully able to check for updates locally`
- `M&S validation ... unsuccessful. User not enrolled in M&S`

Architecturally, this is a very useful distinction for AJ Tools:

1. Update source selection and update entitlement are separate concerns.
2. A product can support local/offline payload delivery while still enforcing licensing online.
3. If AJ Tools should support true offline enterprise updates, we should decide explicitly whether install authorization is:
   - never required
   - locally signed/verified
   - policy-driven
   - or validated against a server
