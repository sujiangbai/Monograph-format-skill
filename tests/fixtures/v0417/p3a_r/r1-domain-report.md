# V0.4.1.7 R1 Semantic-Domain Contract Report

Status: test-only candidate; not an R1 acceptance and not an R2 authorization.

- Base commit: `b4a6d19ecfa71e2844782812d4011cb7ce5ff3be`
- Raw contract SHA-256: `9893b2470f9318187a27d3061e6d752d793cb07c63582b0a46651300be64315f`
- Canonical contract SHA-256: `bde113cfbc4dfb7880c74d3bdfa9e088171b37e32cd15ecd04d93620f6133643`
- Decision rows: 171
- Atomic obligations: 580
- Entities: 580
- Unresolved: 0

## Authority identities

| Role | Path | Git blob OID |
| --- | --- | --- |
| semantic_authority | `docs/plans/v0.4.0-profile-system-and-long-document-reliability.md` | `1d5b43979987744b57a7b6b686f2f3cbbffc3f0b` |
| lifecycle_proof | `docs/plans/v0.4.1-profile-foundation-and-grouped-execution.md` | `4cb9119ccc91ec39876590a26ca0d7e8745c6d3b` |
| lifecycle_proof | `docs/plans/v0.4.1.1-p2-contract-correction.md` | `5ccfa31964672c4409caa7f0c4a7b50e6fd060a4` |
| lifecycle_proof | `docs/plans/v0.4.1.2-p3-capability-and-asset-freeze.md` | `e6716765534384f2e91922907b160166971cb7f7` |
| lifecycle_proof | `docs/plans/v0.4.1.3-p3a-c2-performance-contract-correction.md` | `56d389d1d2c5ccfa5457fd12f11e00f1260d28cd` |
| lifecycle_proof | `docs/plans/v0.4.1.4-horizontal-foundation-and-staged-vertical-validation.md` | `28b3c8658203050fec97201b565fbc5ecae1b0aa` |
| lifecycle_proof | `docs/plans/v0.4.1.5-append-only-production-registry-succession.md` | `3649f216dcd8af9deb0009697b0cd4defdd37e2b` |
| lifecycle_proof | `docs/plans/v0.4.1.6-registry-semantic-closure-and-matrix-succession.md` | `7e3c5503fb73cb9e57b6f2ef3e65fadc7d56f4fa` |
| ownership_route_disposition_only | `docs/plans/v0.4.1.8-atomic-ownership-authority-correction.md` | `abe29ab6b12126fa85f839471a35a4d1bb506eb7` |

## Classification counts

- `capability`: 204
- `constraint`: 128
- `later_rule`: 21
- `non_registry_contract`: 150
- `property`: 77

## Domain-scope counts

- `project_closed`: 114
- `project_open_constrained`: 31
- `source_preserved`: 60

## Semantic heading-level closure

- Typed semantic-level fields: 5
- Allowed levels: `1`, `2`, `3`, `4`
- Value-style authorities: `V040-E-002`, `V040-E-003`, `V040-E-004`
- Selection source: frozen semantic artifact or user/publisher-approved artifact
- `execution-p5b`: consumer/gate only; never a value source

## N007 border-value closure

- Immediate value targets: `V040-N-003`, `V040-N-004`, `V040-N-006` properties only
- `V040-N-005`: exact deferred non-value branch, blocked to `V0.4.2`
- Unknown table semantic: reject; ambiguous classification: blocked QA

## Boundaries

V0.4.0 is the sole semantic/value authority. V0.4.1.8 contributes only ownership, route, and atomic disposition precedence. The report does not accept R1 and does not authorize R2, C2, P3b, a public CLI, Ready, release, push, PR, or merge.
