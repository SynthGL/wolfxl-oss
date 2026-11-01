# WolfXL Community threat model

WolfXL Community reads, creates, and edits `.xlsx` and `.xlsm` workbooks;
`load_workbook()` also reads `.xlsb` and `.xls`. It detects format from content,
so a filename ending in `.xlsx` does not establish the input format. Treat
workbook bytes, ZIP names and metadata, XML, relationships, dimensions, shared
strings, formulas, and embedded content as untrusted. The security contract
covers an attacker supplying a workbook to a supported, unmodified Community
release. The resource-admission guarantees below apply to XLSX-family input
with default limits. Report security defects in every supported input format
under the [security policy](SECURITY.md).

## Guarantees and trust boundaries

- **No code execution from workbook content.** During workbook I/O and supported
  calculation, we treat formulas and embedded payloads as data, not Python,
  shell commands, native code, or instructions to launch an external program.
- **No macro execution.** With `keep_vba=True`, preserve VBA content as workbook
  data; do not execute it. Macro preservation is not macro inspection or removal.
- **No workbook-directed network access.** During parsing, editing, saving, and
  supported calculation, do not fetch external relationship targets, linked
  workbooks, images, or URLs specified by the workbook. Preserve or inspect
  references as data. Package installation and application-controlled I/O are
  outside this workbook-content boundary.
- **Bounded resource admission for XLSX-family input under default limits.**
  Apply finite package limits to reject excessive ZIP entry counts, declared
  decompressed sizes, and compression ratios. Resource exhaustion that bypasses
  these protections, including shared-string amplification, is a security
  defect. These input bounds are not a process memory or execution-time cap.[^resources]
- **Caller-controlled filesystem access.** Read the input and write the output
  selected by the caller. Do not treat workbook package paths or relationship
  targets as authorization to read or write arbitrary host files. Temporary
  storage and permissions remain subject to the host application's policy.

These are security requirements for supported releases. They do not assert
that no vulnerability can exist. We investigate a violation even when a
parser dependency is responsible.

## Default OOXML package limits

The native OOXML admission helpers and Python ZIP inspection helpers use these
defaults:

| Bound | Default | Environment override |
| --- | --- | --- |
| ZIP entries | 200,000 | `WOLFXL_MAX_ZIP_ENTRIES` |
| Uncompressed bytes per entry | 512 MiB (536,870,912 bytes) | `WOLFXL_MAX_ZIP_ENTRY_BYTES` |
| Total uncompressed entry bytes | 4 GiB (4,294,967,296 bytes) | `WOLFXL_MAX_ZIP_TOTAL_BYTES` |
| Uncompressed/compressed bytes per entry | 1,000:1 | `WOLFXL_MAX_ZIP_COMPRESSION_RATIO` |

Every normal XLSX load also applies a per-load `wolfxl.Limits()` admission
pass before these ZIP inspection helpers. Its `max_part_bytes` and
`max_package_bytes` defaults are 512 MiB and 4 GiB, respectively. Both
checks must accept the input: raising an environment variable alone does
not raise the corresponding per-load budget, and raising a `Limits` budget
alone does not raise an environment-controlled ZIP ceiling. To tighten a
workload, pass stricter `limits=` values and keep the ZIP checks at their
defaults or stricter values. `Limits` additionally bounds cells, shared-string
bytes, and materialized text; see the README's load-limit guidance.

Run with the defaults or stricter positive limits for untrusted input.
Increasing a limit expands the admitted workload and its resource risk. Do
not accept limit overrides from workbook content or an untrusted request.
Choose smaller limits for a service with less memory than these defaults.

These ZIP defaults and the `wolfxl.Limits` workbook budgets apply to
XLSX-family input. The `.xlsb` reader does not enforce the ZIP admission
checks listed above, and `limits=` is unsupported for `.xlsb` and `.xls`.
Those formats have no equivalent bounded-resource admission guarantee here.
Resource-exhaustion reports involving them remain in scope.

## Non-guarantees

- **No process sandbox.** The library runs inside the caller's process with its
  permissions. We do not provide tenant isolation, an OS memory quota, a CPU
  deadline, a disk quota, or a concurrency budget.[^resources]
- **No malware sanitization.** Preserved macros, formulas, hyperlinks, external
  links, and embedded files may be dangerous when a user opens the output in
  Excel or another application. Preserving bytes does not certify them as safe.
- **No downstream execution control.** Excel may recalculate formulas, resolve
  links, or offer to run macros after opening an output. The host application
  must decide what to retain and how to distribute it.
- **No guarantee of confidentiality after export.** Workbook contents can
  include hidden sheets, comments, metadata, cached values, and embedded data.
  Editing a visible cell does not redact the rest of the package.
- **No universal formula or fidelity guarantee.** Community supports a documented
  subset of Excel behavior. Correctness and preservation boundaries are
  distinct from the security contract; report data corruption as a defect.
- **No protection from a compromised host or modified installation.** Arbitrary
  application callbacks, optional third-party integrations, dependency tampering,
  caller-selected URLs, and permissive filesystem access require separate review.

## Deployment responsibilities

For an untrusted-upload service, run workbook processing as an unprivileged
worker with OS-enforced memory, CPU/time, disk, and concurrency limits. Deny
network egress, restrict filesystem access to the job's input/output area,
limit uploads before parsing, and cancel jobs that exceed their budget. Keep
security patches current and close workbooks after use. Validate and redact
outputs before sharing them or opening them in a desktop application.

If the service relies on XLSX-family admission limits, classify the input
content with `wolfxl.classify_file_format` and reject other formats before
calling `load_workbook()`. Classify and load the same immutable input.
Do this work inside the resource-limited worker: classification and filename
validation do not provide a sandbox. Treat `.xlsb` and `.xls` processing as
unbounded unless the application supplies effective external containment.

[^resources]: Package metadata limits bound admitted input, not all derived
    objects or work. Shared strings, dense cell expansion, styles, formula
    evaluation, and concurrent workbooks can use substantially more memory or
    CPU than the compressed input. Default limits alone are not a hard RSS or
    wall-clock guarantee. Apply OS limits even to an admitted package, and
    report exploitable exhaustion through the security policy.
