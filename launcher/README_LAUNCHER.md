# PSD Export Launcher (EXE shell + UXP core)

This launcher gives you an `.exe` workflow while keeping the export logic inside the UXP plugin.

## What it does

- Reads plugin id/version from `manifest.json`
- Installs plugin files into:
  - `%APPDATA%\Adobe\UXP\PluginsStorage\PHSP\<version>\Developer\<plugin-id>`
  - `%APPDATA%\Adobe\UXP\PluginsStorage\PHSPBETA\<version>\Developer\<plugin-id>`
- Optionally starts Photoshop
- Supports uninstall/status commands

## Build EXE

From `launcher` folder:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_launcher.ps1
```

Output:

```text
launcher\dist\PSDExportLauncher.exe
```

No .NET SDK is required (PowerShell compiles the C# source directly).

## Quick one-click

Double-click:

```text
launcher\run_install_and_open.cmd
```

This will:
1. Build the launcher EXE
2. Install plugin from repo root
3. Launch Photoshop

## CLI usage

```text
PSDExportLauncher.exe install-run [--plugin-root <path>] [--beta]
PSDExportLauncher.exe install     [--plugin-root <path>]
PSDExportLauncher.exe run         [--beta]
PSDExportLauncher.exe status      [--plugin-root <path>]
PSDExportLauncher.exe uninstall   [--plugin-root <path>]
```

Default action is `install-run`.

## Build single installer package

To create a single-file installer `.exe` that already embeds plugin payload:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_single_installer.ps1
```

Output example:

```text
launcher\dist\PSDExportPipelineInstaller_1.1.77.exe
```

This file can be sent directly to teammates. They run it once to install.

Important: this EXE copies plugin files into Adobe UXP storage folders. Recent Photoshop/Creative Cloud builds may not register sideloaded files copied this way, so the plugin can be present on disk but still not appear in Photoshop. For teammate installs, prefer the CCX route below.

## Build CCX package for teammates

Photoshop UXP plugins are distributed as `.ccx` packages. Build one with:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_ccx_package.ps1
```

Output example:

```text
launcher\dist\PSDExportPipeline_1.1.81.ccx
launcher\dist\Install_PSDExportPipeline_1.1.81.cmd
```

Send both files to teammates in the same folder. They run the `Install_*.cmd` helper, which calls Adobe `UnifiedPluginInstallerAgent.exe` from Creative Cloud Desktop. After install, restart Photoshop and check the Plugins menu/panel list.

If Adobe rejects the generated `.ccx`, use UXP Developer Tools > Actions > Package on the plugin folder. That creates Adobe's packaged CCX flow and is the supported non-UDT install path.

## Optional Git updater

Use this when teammates should choose when to update from Git:

```text
launcher\PSDExportUpdater.cmd
```

To create a small updater-only zip:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_updater_package.ps1
```

Output example:

```text
launcher\dist\PSDExportUpdater_1.1.81.zip
```

The updater stores its settings in:

```text
%APPDATA%\PSDExportPipeline\updater.json
```

Menu flow:

1. Set Git repo URL
2. Select branch/tag
3. Update from Git
4. Build CCX package
5. Install latest CCX
6. Update + build CCX + install

Command-line automation is also available:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\update_from_git.ps1 -RepoUrl "https://example.com/your/repo.git" -Ref "main" -BuildCcx -Install -NonInteractive
```

Notes:
- The Photoshop panel itself does not run `git pull`; UXP should not overwrite its own plugin files.
- If the current plugin folder is already a Git worktree, the updater uses it directly.
- If not, it clones the configured repo to `%APPDATA%\PSDExportPipeline\repo`.
- Installing still uses CCX/Adobe UnifiedPluginInstallerAgent so Photoshop can register the plugin.

Default Git repo:

```text
https://github.com/missingfish1234/PSD_OutPut.git
```

## Publish Current Tool To GitHub

The helper below initializes this plugin folder as a Git repo if needed, commits the current files, and pushes to the default GitHub repo:

```text
launcher\PublishToGitHub.cmd
```

Equivalent command:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\publish_to_github.ps1 -RepoUrl "https://github.com/missingfish1234/PSD_OutPut.git" -Branch "main"
```

The repo should contain `manifest.json` at the root. Generated packages under `launcher/dist/` and local `_inspect_*` folders are ignored by `.gitignore`.

## If plugin does not appear in Photoshop

1. Prefer installing a `.ccx` package with Creative Cloud/UnifiedPluginInstallerAgent.
2. Restart Photoshop completely after install.
3. Ensure Photoshop version is 23+.
4. In Preferences > Plugins, enable Developer Mode only for UDT/debug loading.
5. Use installer status command to confirm copied-file locations:

```text
PSDExportPipelineInstaller_<version>.exe status
```

Expected path includes at least one of:
- `%APPDATA%\Adobe\UXP\PluginsStorage\PHSP\<n>\Developer\<plugin-id>`
- `%APPDATA%\Adobe\UXP\PluginsStorage\PHSP\<n>\External\<plugin-id>`

These copied-file paths are diagnostic only. They do not guarantee Photoshop registered the plugin.

## UXP Developer Tools package

Do not use UDT `Load` on folders under `%APPDATA%\Adobe\UXP\Plugins` or `PluginsStorage`.
Photoshop blocks loading/debugging installed plugins.

Create a clean UDT package:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_udt_debug_package.ps1
```

Output example:

```text
launcher\dist\PSDExportPipeline_UDT_Debug_1.1.80.zip
```

Unzip it, then in UXP Developer Tools use `Add Plugin` on the unzipped plugin folder.
