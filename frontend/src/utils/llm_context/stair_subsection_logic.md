# Stair Subsection Logic Reference

This document describes the stair logic pattern used in the **Demo stair on grade** subsection of the Demolition scope. Use this as a reference when adding similar stair logic to other subsections.

---

## Raw Items

The following raw data items are supported:

- `Demo stairs on grade @ Stair A1`
- `Demo landings on grade @ Stair A1`
- `Demo stairs on grade`
- `Demo landings on grade`

---

## Grouping Logic

1. **Items with `@`**: If an item contains `@`, the text after `@` becomes the group heading (e.g., "Stair A1").
2. **Items without `@`**: Items without `@` form a single group with no heading.
3. **Pairing**: For each group:
   - `Demo stairs on grade` pairs with `Demo landings on grade` when they share the same group key (same text after `@`, or both without `@`).
   - Stairs may exist without landings (landings are optional).
   - Landings may exist without stairs.

**Group structure**: Each group is `{ heading: string | null, stairs: item | null, landings: item | null }`.

---

## Row Structure & Calculation Logic

### Column Reference

| Col | Name    | Index |
|-----|---------|-------|
| A   | Estimate| 0     |
| B   | Particulars | 1 |
| C   | Takeoff | 2 |
| D   | Unit    | 3     |
| E   | QTY     | 4     |
| F   | Length  | 5     |
| G   | Width   | 6     |
| H   | Height  | 7     |
| I   | FT      | 8     |
| J   | SQ FT   | 9     |
| K   | LBS     | 10    |
| L   | CY      | 11    |
| M   | QTY     | 12    |

### Row 1: Landing (only if landing item exists in group)

| Col | Value | Notes |
|-----|-------|-------|
| B   | Landing item name from raw data | |
| C   | Takeoff from raw data | |
| D   | "Treads" if unit is EA, else FT or SQ FT | |
| E, F, G | Empty | |
| H   | 0.67 | |
| I   | Empty | |
| J   | = C | **Red color** |
| K   | Empty | |
| L   | = J*H/27 | **Red color** |
| M   | Empty | |

**After landing row**: Add one empty row.

### Row 2: Stairs

| Col | Value | Notes |
|-----|-------|-------|
| B   | Demo stairs item name from raw data | |
| C   | Takeoff from raw data | |
| D   | "Treads" if unit is EA, else FT or SQ FT | |
| E   | Empty | |
| F   | 11/12 | |
| G   | 4.5 | |
| H   | 7/12 | |
| I   | Empty | |
| J   | = C*G*F | |
| K   | Empty | |
| L   | = J*H/27 | |
| M   | = C | |

### Row 3: Stair slab (always added, not from raw data)

| Col | Value | Notes |
|-----|-------|-------|
| B   | "Stair slab" | Fixed text |
| C   | = 1.3*C(second row) | References stairs row |
| D   | FT | |
| E, F | Empty | |
| G   | = G(second row) | References stairs row |
| H   | 0.67 | |
| I   | = C | |
| J   | = I*H | |
| K   | Empty | |
| L   | = J*G/27 | |
| M   | Empty | |

---

## Sum Logic

- **Columns summed**: I, J, L, M
- **Exclude**: The landing row (first row) is excluded from the sum
- **Include**: Stairs row and Stair slab row only
- **Per group**: Each group has its own sum row immediately after its stairs and stair slab rows
- **Sum row formatting**: Column I — red, not bold; Columns J, L, M — red and bold

---

## Formatting

### Group Heading (when `@` is present)

- **Text**: Heading + colon (e.g., "Stair A1:")
- **Style**: Bold, italic, underlined, black color (#000000)

### Spacing Between Groups

- Add **two empty rows** before each group (except the first group)

---

## Formula Item Types

When implementing, use these formula item types:

| Item Type | Purpose |
|-----------|---------|
| `demo_stair_on_grade_heading` | Group heading row (bold, italic, underline, black) |
| `demo_stair_on_grade_landing` | Landing row: J=C, L=J*H/27 (J and L in red) |
| `demo_stair_on_grade_stairs` | Stairs row: F=11/12, H=7/12, J=C*G*F, L=J*H/27, M=C |
| `demo_stair_on_grade_stair_slab` | Stair slab row: C=1.3*C(stairsRefRow), G=G(stairsRefRow), I=C, J=I*H, L=J*G/27 |
| `demo_stair_on_grade_sum` | Sum row: SUM(I), SUM(J), SUM(L), SUM(M) over stairs+slab rows only (I red/not bold; J, L, M red/bold) |

For other subsections, create analogous item types (e.g., `subsection_stair_heading`, `subsection_stair_landing`, etc.).

---

## Files to Update When Adding Stair Logic

1. **Template** (`capstoneTemplate.js`): Add subsection to structure
2. **Processor** (e.g., `demolitionProcessor.js`): Add subsection detection, grouping logic
3. **Parser** (e.g., `dimensionParser.js`): Add parse case if needed
4. **generateCalculationSheet.js**: Add subsection rendering block
5. **ProposalDetail.jsx, Spreadsheet.jsx, ProposalSheet.jsx**: Add formula handlers for each item type
6. **buildProposalSheet.js**: Add DM reference pattern if proposal output is needed
