# Security policy

[SynthGL/wolfxl-oss](https://github.com/SynthGL/wolfxl-oss) is the canonical
source repository for WolfXL Community, the long-term MIT-licensed open I/O
kernel. Security fixes for shared functionality ship to Community and
Commercial in the same release window. A Commercial subscription is not
required for Community security fixes.

## Supported versions

| Version | Security support | Upgrade guidance |
| --- | --- | --- |
| Community 2.0.x, latest stable patch | Through 2028-10-31, inclusive | Install the latest 2.0.x patch for all available fixes |
| Earlier Community 2.0.x patches | Superseded by the latest patch | Upgrade; fixes are not backported to each old patch |
| Community versions before 2.0 | Unsupported | Upgrade to the latest supported stable release |
| Release candidates (`rc`) | Evaluation only | Use the latest stable release in production |

The support date applies to Community security maintenance, independently of
Commercial feature development. It is not a promise that every historical
patch remains safe to use.

## Report privately

Use [GitHub private vulnerability reporting](https://github.com/SynthGL/wolfxl-oss/security/advisories/new)
or email `security@wolfie.gg`. Do not post an exploit or sensitive workbook in
a public issue. Include the package version, platform, API call, expected and
observed behavior, and a minimal reproducer where possible. Use a synthetic or
redacted workbook rather than confidential customer data.

We acknowledge reports within 7 calendar days. We triage severity and
reachability with the reporter, then work to these remediation targets,
measured from receipt of an actionable report:

| Severity | Target for a fixed release or effective mitigation |
| --- | --- |
| Critical | 7 calendar days |
| High | 30 calendar days |
| Medium | 90 calendar days |

These are remediation targets, not just acknowledgement times. If a target
cannot be met, we tell the reporter why, provide available mitigations and a
revised date, and record the deviation in the advisory. Low-severity reports
are triaged with a documented remediation plan. We update reporters at least
weekly while an accepted vulnerability remains unresolved.

## Scope

Report code execution, unintended network access, unauthorized file access or
writes, parser memory-safety defects, and resource exhaustion from untrusted
files. ZIP decompression, XML parsing, shared strings, workbook dimensions,
and formula processing are in scope. Include reachable vulnerabilities in
shipped dependencies, even when the underlying dependency has its own CVE.
See the [threat model](THREAT_MODEL.md) for the security contract, default
limits, deployment responsibilities, and non-guarantees.

## Advisories, CVEs, and coordinated disclosure

For every accepted vulnerability we publish a GitHub Security Advisory and
obtain a CVE through GitHub's advisory process. For a dependency vulnerability
with an existing CVE, we reference that CVE rather than request a duplicate,
and document WolfXL reachability, affected versions, fixed versions, and
mitigations. Acceptance is based on impact and reachability, not on whether
the defect originated in WolfXL or a dependency.

We use a 90-calendar-day coordinated disclosure window starting with the
initial private report. We coordinate an earlier publication when a fix is
available or exploitation requires urgent warning. Any extension requires an
explicit agreement with the reporter, with a reason and revised disclosure
date. We do not silently extend the window when a fix is late. Advisories
include severity, impact, affected and fixed versions, mitigations, and
reporter credit unless the reporter asks to remain anonymous.

Retrospective advisory coverage includes quick-xml RUSTSEC-2026-0194 in
wolfxl versions before 2.0.11 and the shared-strings resource-exhaustion
("shared-strings bomb") vulnerability. See the
[public advisory list](https://github.com/SynthGL/wolfxl-oss/security/advisories)
for publication status and verified affected and fixed versions. This policy
does not serve as a published advisory or assign an unissued CVE identifier.

## Release cadence

We maintain a weekly stable release train and publish `rc` pre-releases for
compatibility validation. We ship security fixes immediately when ready,
outside the weekly train when necessary, to Community and Commercial in the
same release window for shared vulnerabilities. Security fixes do not wait
for an `rc` evaluation period or a Commercial feature release. Each security
release links its advisories and upgrade guidance.
