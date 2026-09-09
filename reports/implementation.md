# CBD matcher v4 and effective GROUP verification

## Scope
The workspace contains reference/1.jpg, 2.jpg and 3.jpg, but no source CBD XLS/XLSX files. Representative-name diagnostics and synthetic financial fixtures are explicitly distinguished from all-STYLE production data. Actual all-STYLE before/after counts and source-unrounded total reconciliation cannot be asserted without those workbooks or saved state.

## Before
Recorded before changing matcher code in matcher-before.json. HANG TAG scored 0 because PANT/JERSEY remained identity tokens (core Jaccard below 0.8); season numbers were already auxiliary. BUCKLE PCS/SET and WEBBING YD/M were hard rejected before identity scoring. SET and NEW were already auxiliary, but terminal P remained core. GROUP incurred 5 points plus a hard rejection below 94. The Excel exporter replaced effective detail sums with original CBD subtotal values when they differed.

## After
Central normalization preserves Unicode NFKC, dimensions and codes; label garment context and WEBBING terminal P are auxiliary. Numeric tokens are no longer indiscriminately removed. Name, row size/width, codes, construction, position and type conflicts are evaluated independently of UNIT/GROUP. UNIT costs 2 points and GROUP costs 1; threshold 90 and margin 8 remain. Global Hungarian assignment and two-sided ambiguity checks precede quantitative N:1 evaluation. Locked/manual relationships are protected; mixed-group automatic N:1 is prevented.

Representative results: HANG TAG 0 → 96 (Context Match); MAIN BUCKLE 0 → 94 (Auto Matched + Unit Changed); NYLON WEBBING 0 → 94 (Auto Matched + Unit Changed). Three reference-only and three comparison-only rows become three independent 1:1 links; no automatic N:1. These are fixture counts, not whole-production STYLE counts.

## GROUP and export
Relationships store referenceOriginalGroup, comparisonOriginalGroup, effectiveGroup, groupAssignmentSource, manualGroupOverride. finalGroup remains a synchronized compatibility field. Uploaded row GROUP values remain untouched. Source groups are reconstructed during migration, including history snapshots. UI and Excel use shared group resolution/presentation. N displays directional changes; O includes original/effective groups and unit changes. Numeric differences are blank in Excel for unlike units; original Extended Cost is retained. Original CBD subtotals are reference-only and excluded from subtotal formulas. Row height accounts for wrapped text.

Financial regression uses the supplied rounded example: reference OUTSHELL 8.4099 → 8.2368, TRIMS 3.7540 → 3.9271. Comparison OUTSHELL 7.9605, TRIMS 3.6606. Total material cost is conserved at 13.6483 and 12.8774 (floating-point tolerance 0.0001). This demonstrates the calculation rule, not verification against missing original unrounded values.

## Audit
Download Matching Audit examines all selected STYLEs, top three unresolved candidates, blocking conditions, margins, group and unit changes, N:1 relations, row coverage, group totals and source-summary consistency. scripts/audit-comparison.mts accepts saved state JSON. Reports expose original-summary discrepancies instead of replacing detail totals with source subtotals.

## Verification artifacts
- matcher-before.json / matcher-after.json: original and normalized names, attributes, units, candidate scores and margins.
- fixture-current.xlsx / fixture-all.xlsx: generated using the application workbook builder.
- verification-state.json: reproducible synthetic fixture, clearly named FIXTURE.
- local/live-current.xlsx and local/live-all.xlsx: browser downloads when corresponding verification completes.
- local/live-verification.json: browser results, source URL and data scope.

GitHub Pages workflows were not modified or invoked. User changes under reference/ are not included in this change.
