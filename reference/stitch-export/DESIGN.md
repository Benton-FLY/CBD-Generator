---
name: Cost Breakdown Precision System
colors:
  surface: '#f8f9ff'
  surface-dim: '#cbdbf5'
  surface-bright: '#f8f9ff'
  surface-container-lowest: '#ffffff'
  surface-container-low: '#eff4ff'
  surface-container: '#e5eeff'
  surface-container-high: '#dce9ff'
  surface-container-highest: '#d3e4fe'
  on-surface: '#0b1c30'
  on-surface-variant: '#45464d'
  inverse-surface: '#213145'
  inverse-on-surface: '#eaf1ff'
  outline: '#76777d'
  outline-variant: '#c6c6cd'
  surface-tint: '#565e74'
  primary: '#000000'
  on-primary: '#ffffff'
  primary-container: '#131b2e'
  on-primary-container: '#7c839b'
  inverse-primary: '#bec6e0'
  secondary: '#0051d5'
  on-secondary: '#ffffff'
  secondary-container: '#316bf3'
  on-secondary-container: '#fefcff'
  tertiary: '#000000'
  on-tertiary: '#ffffff'
  tertiary-container: '#002113'
  on-tertiary-container: '#009668'
  error: '#ba1a1a'
  on-error: '#ffffff'
  error-container: '#ffdad6'
  on-error-container: '#93000a'
  primary-fixed: '#dae2fd'
  primary-fixed-dim: '#bec6e0'
  on-primary-fixed: '#131b2e'
  on-primary-fixed-variant: '#3f465c'
  secondary-fixed: '#dbe1ff'
  secondary-fixed-dim: '#b4c5ff'
  on-secondary-fixed: '#00174b'
  on-secondary-fixed-variant: '#003ea8'
  tertiary-fixed: '#6ffbbe'
  tertiary-fixed-dim: '#4edea3'
  on-tertiary-fixed: '#002113'
  on-tertiary-fixed-variant: '#005236'
  background: '#f8f9ff'
  on-background: '#0b1c30'
  surface-variant: '#d3e4fe'
typography:
  headline-xl:
    fontFamily: Inter
    fontSize: 24px
    fontWeight: '700'
    lineHeight: 32px
    letterSpacing: -0.02em
  headline-lg:
    fontFamily: Inter
    fontSize: 18px
    fontWeight: '600'
    lineHeight: 24px
    letterSpacing: -0.015em
  headline-md:
    fontFamily: Inter
    fontSize: 15px
    fontWeight: '600'
    lineHeight: 20px
    letterSpacing: -0.01em
  body-md:
    fontFamily: Inter
    fontSize: 13px
    fontWeight: '400'
    lineHeight: 18px
    letterSpacing: 0em
  body-sm:
    fontFamily: Inter
    fontSize: 12px
    fontWeight: '400'
    lineHeight: 16px
    letterSpacing: 0em
  table-header:
    fontFamily: Inter
    fontSize: 11px
    fontWeight: '700'
    lineHeight: 14px
    letterSpacing: 0.06em
  metric-tabular:
    fontFamily: JetBrains Mono
    fontSize: 12px
    fontWeight: '500'
    lineHeight: 16px
    letterSpacing: -0.02em
  metric-delta:
    fontFamily: JetBrains Mono
    fontSize: 11px
    fontWeight: '600'
    lineHeight: 14px
    letterSpacing: 0em
  caption:
    fontFamily: Inter
    fontSize: 11px
    fontWeight: '500'
    lineHeight: 14px
    letterSpacing: 0.01em
rounded:
  sm: 0.125rem
  DEFAULT: 0.25rem
  md: 0.375rem
  lg: 0.5rem
  xl: 0.75rem
  full: 9999px
spacing:
  sidebar-width: 240px
  header-height: 48px
  table-cell-py: 6px
  table-cell-px: 10px
  gap-xs: 4px
  gap-sm: 8px
  gap-md: 12px
  gap-lg: 16px
  gap-xl: 24px
  panel-padding: 16px
---

## Brand & Style

This design system serves high-stakes enterprise costing, multi-level Bill of Materials (BOM) reconciliation, and supply chain margin engineering. The users are enterprise cost engineers, sourcing executives, industrial procurement managers, and financial analysts whose workflows demand relentless data density, error-free auditing, and deterministic clarity.

The visual style blends **Corporate Modern** with **Technical Precision Data Systems**. It strips away decorative whimsy in favor of dense tabular ergonomics, razor-sharp visual hierarchy, and instant cognitive triage of financial deltas. The interface exudes institutional competence: deep slate navy framing provides rock-solid grounding, high-contrast cool whites reduce eye fatigue across multi-hour comparison audits, and deliberate semantic color coding brings immediate visibility to cost creep, margin savings, and unmapped BOM dependencies.

## Colors

The palette is engineered for tabular financial computation, where color is never ornamental—it serves solely as a signal for hierarchy, state, and cost variance.

### Core Canvas & Frame
- **Brand & Structural Slate**: Deep Navy (`#0F172A`) anchors the active global navigation, top hierarchy headers, and master table summary blocks. Mid Slate (`#1E293B`) defines sub-navigation, primary active pills, and sticky column header bars.
- **Interactive Action Blue**: Primary interactive triggers use `#2563EB`, transitioning to `#1D4ED8` on hover and `#1E40AF` on press. Subtle selection wash uses `#EFF6FF` with a border stroke of `#BFDBFE`.
- **Canvas & Containers**: Background viewport is rendered in cool neutral `#F8FAFC`, white data cards in `#FFFFFF`, and structural cell boundaries in `#E2E8F0` and `#CBD5E1`.

### Semantic Variance System
- **Cost Creep / Increase (Negative Variance)**: Muted Rose-Red (`#EF4444`). Badges and table cell callouts apply background `#FEF2F2`, border `#FECACA`, and high-contrast text `#991B1B`.
- **Cost Reduction / Savings (Positive Variance)**: Balanced Emerald (`#10B981`). Fills use `#ECFDF5`, borders `#A7F3D0`, and metrics text `#065F46`.
- **Warning / Review / Unmapped**: Industrial Amber (`#F59E0B`). Background `#FFFBEB`, border `#FDE68A`, and text `#92400E`.
- **Matched / Parity / Reference**: Slate Indigo (`#3B82F6` / `#64748B`). Background `#F1F5F9`, border `#E2E8F0`, and text `#334155`.

## Typography

Typography prioritizes ultra-dense vertical efficiency and exact tabular number scanning.

- **Primary Interface Font**: `Inter` handles all page titles, form labels, control triggers, and item descriptions. Font sizing scales tightly from 11px to 24px, ensuring that complex hierarchical BOM structures fit above the fold without sacrificing legibility.
- **Monospaced Data Font**: `JetBrains Mono` with forced tabular numbers (`font-variant-numeric: tabular-nums`) is mandatory for all currency values, part IDs, units of measure, material consumption ratios, and percentage differences. Decimal places align perfectly across columnar data sets.
- **Micro-Headers**: Table headers and section metadata utilize 11px uppercase styling with `0.06em` tracking (`letter-spacing`) to deliver unambiguous visual partitioning across 10+ column wide financial grids.

## Layout & Spacing

The layout architecture relies on a **dense application shell** tailored for 1440px+ enterprise workstations while maintaining graceful horizontal flow for wide analytical grids:

- **Global Shell**: A fixed vertical navigation bar (240px) at the left, housing brand signifiers, seasonal workspace selectors, and workflow stepper steps.
- **Main Canvas**: A scrollable viewport anchored by a thin 48px utility banner (displaying active seasonal run metadata, file count, and sync state).
- **Tabular Rhythm**: The data grid is compressed to `py-1.5` (6px) and `px-2.5` (10px) per cell. This compact scale ensures up to 28-35 rows are viewable simultaneously without vertical pagination friction.
- **Card & Workspace Panels**: Structural cards maintain an interior padding of 16px (`1rem`), with 12px separation between parent cost groups and nested variant lines.

## Shapes

The design system employs a **Soft Architectural (`1`)** shape language:

- **Data Tables & Containers**: Corner radius `4px` (`rounded-sm` / `0.25rem`) to maximize pixel utility and cleanly hold rigid grid matrices.
- **Form Selects & Inputs**: `4px` to match table cells and maintain sharp alignment within dense inline editing blocks.
- **Workflow Navigation Pills & Status Badges**: `4px` to `6px` radius. Pill buttons remain squared enough to preserve high-density text boundaries while avoiding bubbly consumer aesthetics.
- **Dropzones & Drag Target Boxes**: `6px` dashed borders (`#94A3B8`) with `4px` internal radius.

## Components

### Financial Comparison Data Grids
- **Header Cells**: Solid `#0F172A` or `#1E293B` background with `#FFFFFF` text. Uppercase 11px font with right-aligned indicators for quantitative fields (Units, Price, Usage, Delta).
- **Row States**: Default `#FFFFFF`, subtle alternate zebra `#F8FAFC`, hover row `#F1F5F9`. Selected/Matched rows highlight in `#EFF6FF`.
- **Inline Group Category Rows**: Full-width spanning divider with subtle `#E2E8F0` border and `#334155` medium-weight font to aggregate functional parts (e.g., OUTSHELL, TRIMS, PACKAGING).
- **Tabular Cells**: All currency and usage values must be right-aligned with fixed decimal representation (2 for USD `$0.00`, 4 for precision consumption `0.0000`).

### Variance Badges & Delta Chips
- **Format**: Structured compact badges (`px-2 py-0.5`, font-size 11px, font-mono).
- **Positive Creep**: `+0.9500` rendered with soft red background (`#FEF2F2`), border (`#FECACA`), text (`#991B1B`).
- **Cost Reduction**: `-0.0300` rendered with emerald background (`#ECFDF5`), border (`#A7F3D0`), text (`#065F46`).
- **Parity / Zero**: Subtle hyphen `—` in `#94A3B8` without colored pill.

### Workflow Stepper
- **Left Navigation Vertical Stepper**: Sequential numbered items (`1. Upload`, `2. Match Styles`, etc.).
- **Active Step**: High-contrast blue highlight bar or full solid container (`bg-blue-50`, `text-blue-700`, `font-semibold`).
- **Completed Step**: Slate text `#64748B` with a subtle check indicator.

### Inline Match Selector
- Compact native dropdown select (`h-7`, `text-xs`, `border-slate-300`, `bg-white`) embedded directly inside comparison cells to allow real-time manual overrides and BOM re-linking.

### Dropzone & File Onboarding Cards
- Crisp white cards with dual comparison panes (Reference Season vs. Target Season).
- Dashed borders with centered micro-upload icon and clear file taxonomy tags (`FLY RACING 27 MX PANT...xlsx`).