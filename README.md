# Vivit Control Center

A modular, kiosk-friendly Windows shell and launcher built with WPF (.NET Framework 4.7.2). It provides a unified hub with multiple modules (browser, media player, notes, office host, terminal, scripting, SSH/SFTP, Steam, and more), optional shell replacement mode, and full localization.

## Features
- Modular sidebar with quick access to many tools and web modules
- Splash screen with module loading progress
- Full localization (EN/DE/ES/FR/RU/JA/ZH/EO)
- Optional Windows shell replacement mode (set/restore)
- Settings to enable/disable modules and configure defaults
- Programs launcher (manage custom apps)
- Custom URLs per web module and Fediverse URL
- Email module with IMAP/SMTP, paging, HTML preview, and optional OAuth2 (Google/Microsoft/Custom)
- SSH enhancements: built-in log viewer with Load/Follow and configurable macros

### Bundled modules
- Explorer: Simple file manager
- Office: Word/Excel host (create/open/save) – requires Microsoft Office
- Notes: Rich-text notes with topics tree, autosave, and formatting
- Media Player: Playlist (add/remove/clear), play/pause/stop, previous/next, seek bar, volume, Now Playing and play time
- Steam: Simple Steam library view (Steam must be installed)
- Webbrowser + Social: Embedded browser (WebView2), Social can point to Fediverse
- News: RSS mode (feeds and max count) or Web news mode
- Terminal: Console runner
- Scripting: Simple script runner with language detection (Run via F5/Ctrl+Enter)
- Email: IMAP inbox viewer with background sync/cache, paging (Load More), HTML preview, and SMTP send; supports Password or OAuth2 auth (Google/Microsoft/Custom)
- SSH: Interactive SSH console plus log viewer (tail) with Follow and user-defined macros
- SFTP: File transfer over SSH – requires SSH.NET (Renci.SshNet)
- Launch/Programs: Launcher for custom apps
- Settings: Language, module visibility, paths, news mode, custom URLs, SSH logs/macros, shell mode, and more

## Requirements
- Windows 10/11
- .NET Framework 4.7.2
- WebView2 Runtime (for `Webbrowser`, `News` (Web), and `Social`)
- Microsoft Office (for `Office` module)
- Steam client (for `Steam` module)
- SSH.NET (Renci.SshNet) assemblies available for SFTP functionality
- Optional: SSH.NET (Renci.SshNet) for integrated SSH client; falls back to `ssh.exe` if not present

## Install
1. Download a build or clone and build the solution in Visual Studio.
2. Ensure WebView2 Runtime is installed.
3. If you use the `Office` module, install Microsoft Office (Word/Excel).
4. For SFTP, add SSH.NET by referencing NuGet package `Renci.SshNet` or placing its assemblies next to the executable.

## Build (developers)
- Open `Vivit Control Center/Vivit Control Center.csproj` in Visual Studio.
- Target framework: .NET Framework 4.7.2.
- Build Any CPU or x64 (Debug/Release).

## Usage
### General
- Start the application. A splash shows module loading progress.
- Use the left sidebar to switch modules. The window title shows the active module tag.

### Settings
Open the `Settings` module to configure:
- Language (applies after restart)
- Enable/disable modules (sidebar visibility)
- News mode (RSS vs Web), edit RSS feeds and max items
- Default paths for `Explorer` and `Scripting`
- Fediverse URL (used by `Social`)
- Custom URLs for web modules
- SSH: manage favorite remote log file paths and define macros (name + command)
- Manage Programs (add entries to the launcher)
- Windows Shell: Set this app as shell or restore the standard shell

Saving settings may recommend restarting to apply changes.

### Launch/Programs
- Manage your launchable apps in `Settings` ? Manage Programs.
- Use the `Programs`/`Launch` module to quickly start them.

### Media Player
- Add: choose audio files to add to the playlist
- Remove/Clear: manage playlist items
- Double-click a track to play; use Previous/Next, Play/Pause, Stop
- Seek bar and volume are available; Now Playing and current/total time are displayed

### Notes
- Create a topic (tree on the left), write rich text, formatting toolbar buttons
- Autosave triggers after edits; rename/delete topics as needed

### Office
- Choose app (Word/Excel), create or open documents
- Save/Save As supported; Office must be installed

### Email
- Click `Accounts` to add or edit accounts (IMAP/SMTP). Choose Password or OAuth2.
- For OAuth2, select provider (Google/Microsoft/Custom) and set Client ID/Secret, Tenant, and Scope as needed.
- Select an account and folder (Inbox by default) to load messages.
- Use `Load More` to page older messages. Messages are cached by a background sync service.
- Select a message to preview (HTML supported). Click `New` to compose and `Send` via SMTP.

### SSH
- Enter address (`user@host` or `host[:port]`), click `Connect`
- Provide credentials when prompted; type in the input box and click `Send`
- Log viewer: pick a favorite log from the dropdown and click `Load` to fetch recent lines; enable `Follow` to tail in real time
- Macros: buttons on the right run your predefined commands; configure them in `Settings`

### SFTP
- Requires SSH.NET (Renci.SshNet)
- Browse local/remote, upload/download files

### Browser, Social, News (Web)
- Embedded WebView2; customize URLs in `Settings`

## Shell mode (optional)
- In `Settings`, use Set as Shell to replace `explorer.exe` for the current user.
- Admin rights are required (regedit import). A `Restore Standard Shell` button is provided to revert.
- Warning: When in shell mode, `Explorer` does not auto-start. Use the app’s power options or restore shell.

## Localization
- Language can be selected in `Settings`. On next start, the app loads the corresponding resource dictionary.

## Troubleshooting
- WebView2 init error: Install the Microsoft Edge WebView2 Runtime.
- SFTP: If you see “SSH.NET not found”…, add the `Renci.SshNet` package or copy its assemblies next to the executable.
- SSH: If SSH.NET is missing, the app will use `ssh.exe` automatically.
- Email: Auth failures usually indicate wrong credentials, OAuth scopes/tenant, or server settings.
- Email: If HTML doesn’t render, the message may be plain text only; try viewing another folder.
- Office: If hosting fails, ensure Office is installed and registered.
- Shell mode: If you’re stuck in shell mode, run the app as admin and choose `Restore Standard Shell` in `Settings`.

## License
See the repository’s license file.
