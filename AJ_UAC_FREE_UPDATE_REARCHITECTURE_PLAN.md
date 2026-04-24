# AJ Tools UAC-Free Update Re-Architecture Plan

Date: 2026-04-24

## Goal

Redesign AJ Tools so that normal updates do not require UAC, while preserving the actual Microsoft Project add-in functionality.

The core requirement is simple:

- First install may be allowed to use a bootstrapper.
- Normal updates must not depend on `Program Files`, `HKLM`, `msiexec`, or `runas`.
- The add-in runtime must become user-scoped so it can be replaced without elevation.

## Executive summary

AJ Tools cannot get SSI-style no-UAC updates by changing only the downloader.

The updater currently performs a full machine-wide reinstall:

- install root is `C:\Program Files (x86)\AJTools`
- add-in registration is written under `HKLM`
- updates launch `AJSetup.exe`
- `AJSetup.exe` requires admin
- `AJSetup.exe` runs `msiexec`

As long as those are true, updates will keep requiring UAC.

To behave like SSI, AJ Tools needs a split architecture:

1. A small bootstrap/install layer
2. A user-scoped runtime layout
3. A side-by-side updater that swaps the active version by changing current-user registration

## What must change

These are non-negotiable if the goal is no-UAC updates.

### 1. Stop treating the add-in runtime as a machine install

Current AJ behavior:

- files live in `Program Files`
- updates replace protected files in place

Target behavior:

- runtime lives under a user-writable root such as:
  - `%LocalAppData%\AJTools\`
  - or another IT-approved user-writable folder if `LocalAppData` is blocked in some environments

### 2. Stop registering the Project add-in under `HKLM`

Current AJ behavior:

- Project add-in registration is machine-wide
- manifest path points to `C:\Program Files (x86)\AJTools\...vsto|vstolocal`

Target behavior:

- register under:
  `HKCU\Software\Microsoft\Office\MS Project\Addins\ArianJahandarfardsAddIn`
- `Manifest` points to the active user-scoped `.vsto|vstolocal`

### 3. Stop using MSI as the normal update mechanism

Current AJ behavior:

- update flow downloads an MSI
- launches `AJSetup.exe`
- `AJSetup.exe` uninstalls and reinstalls MSI packages

Target behavior:

- update flow downloads a ZIP
- extracts the next version into a new version folder
- validates the new payload
- repoints the `HKCU` manifest to the new `.vsto`

### 4. Stop requiring admin in the updater executable

Current AJ behavior:

- `AJSetup.exe` is marked `requireAdministrator`
- updater launches it with `Verb = "runas"`

Target behavior:

- the runtime updater runs as the current user
- all files it touches must be writable by that user
- all registry writes must stay in `HKCU`

### 5. Stop hardcoding `Program Files` paths inside the add-in

Current AJ behavior:

- update code and some UI assets assume `C:\Program Files (x86)\AJTools`

Target behavior:

- all runtime paths resolve from a single install-layout service
- the code should never care whether the base root is `LocalAppData`, a network-approved folder, or a future enterprise path

## What can stay

Most of the business logic can stay.

These do not need a conceptual rewrite:

- the actual Microsoft Project add-in functionality
- ribbon code
- milestone logic
- dynamic status logic
- project-linker logic
- user-level config/log patterns that already live under `AppData` or Documents

The big rewrite is not the add-in behavior. It is the deployment model.

## Target architecture

## 1. New runtime layout

Use a structure like this:

```text
%LocalAppData%\AJTools\
  Application Files\
    AJTools_3_1_0_0\
      Arian Jahandarfards MS Project Add-in.vsto
      Arian Jahandarfards MS Project Add-in.dll
      Arian Jahandarfards MS Project Add-in.dll.manifest
      Icons\
      Dependencies...
    AJTools_3_2_0_0\
      ...
  Shared\
    Assets\
    Templates\
  Data\
    Settings\
    Cache\
    Logs\
  Downloads\
    ajtools-3.2.0.zip
  Staging\
    AJTools_3_2_0_0\
  state.json
```

Important rule:

- versioned runtime files go in `Application Files`
- mutable data does not go in versioned folders

That prevents updates from overwriting user data and makes rollback easy.

## 2. Active-version model

The active version should be a pointer, not a folder that gets modified in place.

Recommended source of truth:

- current Project registration in `HKCU`
- optional `state.json` that records:
  - current version
  - previous version
  - install root
  - last successful update
  - last failed update reason

Why this matters:

- updates become side-by-side deployments
- rollback is a pointer reversal, not a reinstall

## 3. Update feed model

Adopt an SSI-style manifest + ZIP payload.

Recommended manifest shape:

```json
{
  "version": "3.2.0",
  "buildDate": "2026-04-24",
  "downloadZipFile": "https://your-host/ajtools/ajtools-3.2.0.zip",
  "releaseNotesUrl": "https://your-host/ajtools/release-notes-3.2.0.html",
  "minimumUpdaterVersion": "1.0.0",
  "sha256": "..."
}
```

Optional later additions:

- `fileMap`
- `minimumSupportedProjectVersion`
- `isMandatory`
- `rollbackBlocked`

## 4. Normal update flow

The new flow should be:

1. Check manifest JSON.
2. Compare installed version to feed version.
3. Download ZIP into `Downloads`.
4. Verify hash/signature.
5. Extract into `Staging\AJTools_x_y_z_w`.
6. Verify required files exist:
   - `.vsto`
   - add-in DLL
   - manifest
   - dependencies
7. Move staged folder into `Application Files\AJTools_x_y_z_w`.
8. Close Project if required.
9. Update `HKCU\...\Addins\ArianJahandarfardsAddIn\Manifest` to point to the new `.vsto|vstolocal`.
10. Refresh any VSTO metadata only if needed.
11. Keep prior version folder for rollback.
12. On next startup, validate load success.
13. If validation fails, revert pointer to prior version.

The critical change is this:

- do not replace the current version in place
- install a new version beside it, then switch

## 5. Offline update flow

AJ Tools should copy SSI's strongest idea here:

- update source and update authorization are separate concerns

Recommended local mode:

1. User or IT selects a folder path.
2. AJ reads `AJToolsUpdateInfo.json` from that folder.
3. Manifest points to either:
   - a ZIP in the same folder
   - or an absolute local/network path
4. AJ installs from that ZIP using the same side-by-side logic as online mode.

This gives you:

- local folder updates
- shared drive updates
- USB/offline updates
- SCCM/Intune-distributed offline payloads

without changing the runtime installer logic.

## 6. Bootstrapper model

Keep the bootstrapper separate from the runtime updater.

### Bootstrapper responsibilities

- first-time setup
- creating the AJ root folder
- copying the initial runtime payload
- writing `HKCU` add-in registration
- validating prerequisites such as VSTO runtime
- optionally migrating old installs

### Runtime updater responsibilities

- manifest check
- ZIP download or local ZIP intake
- extraction
- validation
- registration cutover
- rollback
- cleanup

This separation matters because your current `AJSetup.exe` is trying to be:

- installer
- uninstaller
- update orchestrator
- MSI wrapper
- registry cleaner

That is exactly the kind of design that keeps dragging admin requirements into the update path.

## 7. Registration model

The new primary registration path should be:

`HKCU\Software\Microsoft\Office\MS Project\Addins\ArianJahandarfardsAddIn`

Values:

- `Description`
- `FriendlyName`
- `LoadBehavior = 3`
- `Manifest = file:///.../Arian Jahandarfards MS Project Add-in.vsto|vstolocal`

Do not depend on `HKLM` registration for the normal product path.

If 32-bit and 64-bit registry-view quirks matter for Project, handle them explicitly, but still keep them current-user scoped where possible.

## 8. VSTO and signing strategy

Do not throw away VSTO manifests just because standard ClickOnce is not working for this scenario.

SSI's model strongly suggests the right balance:

- keep signed VSTO manifests
- keep the normal VSTO load model
- replace only the update-delivery mechanism

That means AJ should likely become:

- VSTO runtime + signed manifests
- custom JSON + ZIP updater
- custom `HKCU` registration cutover

Important requirement:

- signing and certificate continuity must be planned now
- if the publishing cert changes later, update trust can break even if the files copy correctly

## 9. Rollback strategy

Rollback must be a first-class feature, not an afterthought.

Recommended design:

- keep the previous version folder
- write `currentVersion` and `previousVersion` into `state.json`
- on Project startup, detect whether the add-in loaded successfully
- if the new version fails validation:
  - repoint `Manifest` back to the previous version
  - mark the failed version as bad
  - do not retry automatically forever

Potential validation signals:

- add-in startup completed
- ribbon loaded
- no fatal startup exception
- optional health heartbeat written by the add-in

## 10. Cleanup strategy

Do not aggressively delete old versions immediately.

Recommended rule:

- keep current version
- keep previous version
- delete anything older than the last two successful versions after a safe grace period

Why:

- rollback stays cheap
- support becomes easier
- disk growth stays controlled

## 11. Corporate-environment strategy

This is where teams often get surprised.

Even if the updater is technically correct, corporate policy can still break it.

You need a plan for:

- AppLocker / WDAC blocking execution from `AppData`
- TLS inspection breaking downloads
- authenticated proxies
- locked-down Office add-in policy
- profile resets / roaming profiles
- security tooling quarantining self-updating binaries

So the design should not hardcode `LocalAppData` as the only option.

Instead:

- create an install-root abstraction
- default to `%LocalAppData%\AJTools`
- allow IT-approved override via config or bootstrapper parameter

That gives you SSI-style updates without forcing one path forever.

## What needs to change in this repo

## 1. Replace the installation model

Current projects involved:

- `AJToolsInstaller`
- `AJSetup`
- `Arian Jahandarfards MS Project Add-in`

Required change:

- stop treating `AJToolsInstaller` and `AJSetup` as the primary runtime update mechanism

Recommended future project split:

- `Arian Jahandarfards MS Project Add-in`
  - actual Project add-in
- `AJBootstrapper`
  - first install + migration + prerequisites + HKCU registration
- `AJRuntimeUpdater`
  - manifest check + ZIP install + rollback + cleanup

You could fold the bootstrapper and updater together later, but they should be logically separate even if they share utilities.

## 2. Remove hardcoded machine paths

These need to be refactored behind a layout service:

- `AJUpdater.cs`
- `AJUpdatePrompt.cs`
- dynamic status form logo loading
- anything else that assumes `C:\Program Files (x86)\AJTools`

Introduce a central class like:

- `AJInstallLayout`

Example responsibilities:

- `GetRootPath()`
- `GetApplicationFilesPath()`
- `GetVersionFolder(version)`
- `GetSharedAssetsPath()`
- `GetStateFilePath()`
- `GetDownloadedZipPath(version)`
- `GetCurrentManifestPath(version)`

Once that exists, the rest of the code stops caring where the files live.

## 3. Replace `AJUpdater` with a real runtime updater

`AJUpdater` should stop launching `AJSetup.exe`.

It should instead:

- read manifest JSON
- compare versions
- support online and offline sources
- download ZIP
- verify integrity
- stage and extract
- repoint `HKCU` registration
- schedule Project restart if needed

It should not:

- invoke `runas`
- invoke `msiexec`
- touch `HKLM`
- assume `Program Files`

## 4. Replace MSI cleanup code with migration logic

Current `AJSetup\Form1.cs` contains lots of machine-install cleanup.

That should be replaced with a one-time migration path:

1. Detect old machine install.
2. Offer migration.
3. Copy or download the new user-scoped runtime.
4. Register `HKCU` version.
5. Optionally unregister old machine version.
6. Leave a clean uninstall/recovery path.

Migration is different from normal updates and should be treated as such.

## 5. Add a registry helper for Project registration

Create a dedicated helper that owns:

- add-in key creation
- manifest path updates
- load behavior repair
- optional metadata cleanup if a stale VSTO registration exists

This must be reusable by:

- bootstrapper
- runtime updater
- repair tool

## 6. Add persistent update state and logging

Create a durable update journal.

Recommended data:

- current version
- previous version
- available version
- source URL or source folder
- zip hash
- install start and finish timestamps
- validation status
- rollback result

This matters a lot for enterprise support.

## 7. Separate runtime assets from mutable data

Anything the user might modify must not live in the versioned runtime folder.

Review each of these categories:

- templates
- exported files
- logs
- caches
- settings
- user-selected workbooks or data

If it needs to survive updates, it should live under `Data` or another non-versioned location.

## Migration plan

## Phase 1: Make the codebase path-agnostic

Goal:

- eliminate `Program Files` assumptions

Work:

- create `AJInstallLayout`
- refactor asset lookups
- refactor updater path logic
- refactor manifest path construction

This phase is required before the new updater can be clean.

## Phase 2: Move registration to `HKCU`

Goal:

- support current-user add-in registration

Work:

- create Project registration helper
- register `.vsto|vstolocal` from a non-machine location
- verify Project loads correctly from `HKCU`

This is one of the key gates for UAC-free updates.

## Phase 3: Introduce side-by-side version folders

Goal:

- stop in-place replacement

Work:

- define version folder naming
- define `state.json`
- define staging and rollback behavior

## Phase 4: Replace MSI-based updates with ZIP-based updates

Goal:

- remove `msiexec` and `runas` from the update path

Work:

- manifest parser
- ZIP download/extract
- hash validation
- registration cutover
- cleanup

## Phase 5: Build migration from the old architecture

Goal:

- move existing AJ users from machine-wide install to the new model

Work:

- detect old MSI install
- guide user through migration
- preserve settings/data
- retire old registration

## Phase 6: Add repair and support tooling

Goal:

- make enterprise deployment survivable

Work:

- repair registration
- re-point manifest
- rebuild VSTO metadata if needed
- show current version, active folder, and last update result

## Most important design decisions to make before coding

These decisions affect everything downstream:

1. What is the default user-scoped install root?
2. Do we support an IT-approved override path from day one?
3. Will runtime updates always be ZIP-based?
4. Where do templates and other shared content live?
5. What is the rollback trigger and rollback policy?
6. How do we validate a successful first launch after cutover?
7. How will certificate rotation be handled?
8. What is the migration experience for current `Program Files` installs?

## Bottom line

If the goal is to replicate SSI's no-UAC update behavior, AJ Tools must stop updating a machine-wide installation.

The real architecture change is:

- from machine-scoped install to user-scoped runtime
- from in-place MSI reinstall to side-by-side ZIP deployment
- from `HKLM` registration to `HKCU` registration
- from elevated setup to current-user cutover

Once those four changes happen, the updater can behave like SSI's.

Without those four changes, AJ Tools will keep asking for UAC no matter how polished the downloader becomes.
