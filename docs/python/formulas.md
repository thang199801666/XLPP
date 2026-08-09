# Formulas and calculation

## Write formulas

```python
sheet["C2"].set_formula("=A2+B2")
sheet["D2"].set_formula("=SUM(B2:B100)")
sheet["E2"].set_formula("='Lookup Data'!A1")
```

Assigning `"=A1+B1"` to `.value` writes text. Use `set_formula()` for formulas.

## New Excel functions

```python
function = xlpp.xlfn("FILTER")
sheet["F2"].set_formula(f"={function}(A2:A100,B2:B100>0)")
```

`xlfn()` adds the compatibility prefix expected by OOXML for newer functions.

## Shared, array, and dynamic-array formulas

```python
sheet["D2"].set_shared_formula("=A2*2", 7, "D2:D100")
sheet["E2"].set_array_formula("=SUM(A2:B100)", "E2:E2")
sheet["F2"].set_dynamic_array_formula("=_xlfn.SORT(A2:A100)", "F2:F100")
```

Formula metadata is available through `cell.formula_metadata` for imported and
advanced formulas.

## Calculate formulas

```python
options = xlpp.CalculationOptions()
options.recursive_dependencies = True
options.update_cached_values = True
options.evaluate_volatile_functions = True
options.spill_dynamic_arrays = True
options.max_depth = 256

report = book.calculate_formulas(options)

print(report.formula_cells_visited)
print(report.formula_cells_evaluated)
print(report.cached_values_updated)
print(report.unsupported_formulas)
print(report.evaluation_errors)
print(report.warnings)
```

`report.success` is the convenient overall indicator. Inspect counters and
warnings when formulas can include unsupported functions or external data.

!!! warning "Not the complete Excel calculation engine"

    XL++ covers a broad catalog and advanced reference behavior, but does not
    reproduce every Excel function, coercion rule, locale behavior, financial
    edge case, cube function, or external-data calculation path. Unsupported
    formulas are reported. Excel may remain the authoritative calculator for
    workbooks that depend on complete Excel semantics.

## Dirty and targeted calculation

```python
options = xlpp.CalculationOptions()
changed = xlpp.CalculationCell()
changed.sheet = "Sales"
changed.cell = "B2"
options.changed_cells = [changed]
options.recursive_dependencies = True

report = book.calculate_formulas(options)
print(report.dirty_formula_cells_selected)
```

Use changed cells to limit recalculation to dependency-reachable formulas.

## Iterative calculation

```python
options = xlpp.CalculationOptions()
options.iterative_calculation = True
options.max_iterations = 100
options.max_change = 0.001

report = book.calculate_formulas(options)
print(report.iterative_iterations)
print(report.iterative_convergence_failures)
```

Enable this only for workbooks intentionally designed around circular
calculation.

## External references

```python
def resolve_external(workbook_name, sheet_name, reference):
    # Return the value expected by your application-specific source.
    return None

options = xlpp.CalculationOptions()
options.external_reference_resolver = resolve_external
report = book.calculate_formulas(options)
```

External resolution is application-provided. XL++ does not open linked
workbooks or connect to external systems automatically.

## Dependency graph

```python
graph = book.dependency_graph()

for edge in graph.edges:
    print(
        edge.dependent_sheet,
        edge.dependent_cell,
        edge.kind,
        edge.precedent_sheet,
        edge.precedent_reference,
    )

precedents = graph.precedents_of("Sales", "D2")
dependents = graph.dependents_of("Sales", "B2")
```

The report distinguishes cell/range, defined-name, table, external, and
volatile-reference edges:

```python
print(graph.report.formula_cells)
print(graph.report.edges)
print(graph.report.unresolved_symbols)
```

## Translate references

```python
result = xlpp.rename_worksheet_references(
    "=SUM(Old!A1:A10)",
    "Old",
    "New",
)
print(result.value)
print(result.references_changed)
```

Other helpers:

```python
edit = xlpp.StructuralEdit(
    "Sales",
    xlpp.StructuralEditKind.INSERT_ROWS,
    2,
    1,
)
xlpp.translate_formula_references(formula, "Sales", edit)
xlpp.translate_range_references(reference, "Sales", edit)
xlpp.invalidate_worksheet_references(formula, "RemovedSheet")
```

Workbook structural edits use the same translation infrastructure to update
formulas, names, tables, charts, pivots, and other reference-bearing features.

## Save-time calculation

```python
options = xlpp.SaveOptions()
options.calculate_formulas_before_save = True
book.save("calculated.xlsx", options)
```

For production workflows that need to inspect calculation diagnostics, call
`calculate_formulas()` explicitly before save rather than relying only on the
save option.

## Formula tips

- Store formulas in invariant OOXML/Excel syntax, not localized UI names.
- Quote worksheet names that contain spaces or punctuation.
- Use absolute references for charts and long-lived named formulas.
- Inspect unsupported/evaluation counters after native calculation.
- Set workbook calculation properties when Excel should recalculate on open.
- Test dynamic arrays and iterative formulas with the target Excel version.
