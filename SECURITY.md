# Security Policy

## Supported versions

| Version | Supported |
|---|---|
| 2.0.x (Community) | Critical correctness and security fixes |
| Earlier releases | Not supported |

WolfXL Commercial releases are maintained separately; commercial customers
should use the support channel included with their plan.

## Reporting a vulnerability

Report vulnerabilities privately through
[GitHub private vulnerability reporting](https://github.com/SynthGL/wolfxl-oss/security/advisories/new).
Do not open a public issue for a security report.

Include the wolfxl version, a minimal reproducing workbook or script, and the
observed behavior. Untrusted-input parsing paths (workbook loading and XML
handling) are the highest-priority surface.

You will receive an acknowledgement within seven days. Fixes for accepted
reports ship as a patch release on the current Community line, and the
advisory is published after the fix is available.
