# Security policy

## Supported versions

Security fixes are applied to the latest released `0.1.x` version.

## Scope of the first release

AiOffice `0.1.0` generates DOCX packages but does not import Office Open XML packages.
JSON and Markdown inputs are still untrusted data: callers should impose file-size and
workspace path limits appropriate to their environment.

The CLI never overwrites an input document when applying a patch. A committed patch
requires an explicit output path.

Macro-enabled Office formats (`.docm`, `.xlsm`, and `.pptm`), remote assets, external
renderers, and arbitrary network access are not supported.

## Reporting a vulnerability

Please report vulnerabilities privately to the project maintainer before opening a
public issue. Include the affected version, a minimal reproduction, and the impact.
