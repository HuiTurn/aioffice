# Security policy

## Supported versions

Security fixes are applied to the latest released `0.1.x` version.

## Scope of the first release

AiOffice `0.2.0.dev0` imports DOCX packages through a bounded security scanner. It
rejects path traversal, duplicate or encrypted ZIP entries, suspicious compression
ratios, oversized packages and XML parts, DTD/entity expansion, and macro payloads.
External relationships may be indexed but are never fetched by the core engine.

The CLI never overwrites an input document when applying a patch. A committed patch
requires an explicit output path.

Macro-enabled Office formats (`.docm`, `.xlsm`, and `.pptm`), remote assets, external
renderers, and arbitrary network access are not supported.

## Reporting a vulnerability

Please report vulnerabilities privately to the project maintainer before opening a
public issue. Include the affected version, a minimal reproduction, and the impact.
