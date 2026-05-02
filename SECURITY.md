# Security Policy

## Supported Versions

Security updates are provided for the latest public release only.

| Version | Supported |
|---|---|
| Latest release | Yes |
| Older releases | No |
| Development builds | No |

## Reporting a Vulnerability

Do not report security vulnerabilities through public GitHub issues.

Use GitHub Security Advisories for private reporting:

1. Open the repository on GitHub.
2. Go to **Security**.
3. Open **Advisories**.
4. Choose **Report a vulnerability**.

If GitHub Security Advisories are not available, contact the maintainer privately before publishing details.

Please include:

- A clear description of the issue.
- Steps to reproduce.
- Affected version.
- Windows version.
- PowerShell version.
- Whether the issue requires admin rights.
- Whether the issue affects startup, update, scheduled tasks, exported diagnostics, or local files.
- Screenshots or logs if useful.

Do not include secrets, auth keys, private keys, session tokens, or non-redacted diagnostic exports in reports.

## Scope

This security policy covers vulnerabilities in **Tailscale Control** itself, including:

- Installer behavior.
- Update behavior.
- Scheduled task creation.
- Startup behavior.
- Local file permissions.
- Diagnostic export redaction.
- Tray menu actions.
- PowerShell execution flow.
- Handling of local paths and app assets.

This policy does not cover vulnerabilities in:

- Tailscale itself.
- Windows.
- PowerShell.
- .NET.
- GitHub.
- Third-party services opened by the app.

If the issue is in Tailscale, report it to the Tailscale project or Tailscale security team.

## Diagnostic Exports

Tailscale Control can export diagnostic JSON files.

The default export mode should be **redacted**. Redacted exports are intended for GitHub issues and support requests.

Non-redacted exports may contain sensitive information such as:

- Device names.
- Tailnet names.
- Tailscale IPs.
- Local network IPs.
- Public IPs.
- DNS configuration.
- Local usernames.
- Local paths.
- Command history.

Do not upload non-redacted exports to public issues unless you have reviewed the file and accept the risk.

## Installer and Updates

The installer should validate downloaded files using SHA256 hashes.

Release assets should be treated as the source of truth for public installs and updates. The app should not update from the main branch because the branch may contain unfinished or untested changes.

Expected release assets include:

- `install.ps1`
- `tailscale-control.ps1`
- `tailscale-control.ico`
- `tailscale.ico`
- `tailscale-mtu.ico`

If hashes do not match, installation or update should fail.

## Scheduled Tasks

Tailscale Control may use Windows Task Scheduler for startup or update flows.

Startup tasks should run as the current user and should not require admin rights.

Tasks that update the Tailscale client may require elevated privileges. These tasks should be created intentionally and should not silently run arbitrary user-writable scripts with admin rights.

Any elevated task should use a controlled script path and predictable behavior.

## Local File Permissions

Tailscale Control stores app files under the user profile by default.

Files that do not require admin rights should not be placed in protected system locations.

If any future component runs with elevated privileges, it must not execute scripts from user-writable locations unless the security model is reviewed first.

## Responsible Disclosure

Please allow reasonable time for investigation and patching before public disclosure.

The maintainer will aim to:

- Confirm receipt of the report.
- Reproduce and assess the issue.
- Fix confirmed vulnerabilities.
- Credit the reporter if requested and appropriate.
- Publish a patched release when ready.

## Security Best Practices for Users

- Install only from official GitHub releases.
- Review non-redacted exports before sharing them.
- Keep Tailscale updated.
- Keep Windows updated.
- Avoid running modified scripts from untrusted sources.
- Do not share exported logs that contain private network details.
