# Contributing to Tailscale Control

The Windows tray app is [`tailscale-control.ps1`](tailscale-control.ps1) and its installer is [`install.ps1`](install.ps1). Start with a PowerShell syntax check. For app changes, use a test Windows account and the Tailscale client version you want to support.

## Review a change

- Check opening and closing the tray menu, startup, and the affected controls. For hotkey changes, check registration, conflicts, and cleanup on exit.
- For account or connection changes, verify disconnected and signed-out states and how failures from the Tailscale client appear. Use test accounts for real switching.
- For installer changes, compare first install and update, then check state after uninstall. Keep Windows startup entries and personal settings out of the fixture.

Report the Windows and Tailscale versions, the tray action or hotkey, and the expected and observed state. Sanitize account names, device names, and diagnostics before sharing.
