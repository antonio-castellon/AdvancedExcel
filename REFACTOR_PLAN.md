# AdvancedExcel Refactor Plan

**Date:** 2026-06-09  
**Branch:** `refactor/split-modules` (created for this work)  
**Repo location:** `C:\DEV\TE\AdvancedExcel`  
**Goal:** Split the monolithic `src/AdvancedExcel.py` (1966 lines) into logical modules **without breaking any existing caller code or documented usage patterns** (including sspipe `p(data.method)` chaining).

## Current State Summary (Analysis Findings)

- Single file `src/AdvancedExcel.py` contains the entire implementation.
- `class AdvancedExcel(object):` only defines **16 real methods** (including `__init__`).
- **51 public "methods" + 7 private helpers** are defined at module top level as `def foo(self, df, ...):`. They are **not** inside the class.
- Many of these top-level functions call `self.other_method(...)` (e.g. `transpose` → `self.setRowAsHeader`, various row removers → `self.removeRowIfUniqueValue`).
- No attachment / monkey-patch code exists that binds the top-level functions onto the class (only an `OpenpyxlReader._convert_cell` patch at the very end).
- Consequence: only the first ~15 methods are true instance methods today. Most row/column/cleaning/reshape operations documented in the README and examples are fragile or non-functional when accessed as `data.method` or `p(data.method)`.
- Package layout quirks:
  - `setup.py`: `name="advance_excel"` (typo, missing 'd'), `package_dir={'': 'src'}`, `find_packages(where='src')`.
  - `src/__init__.py`: `from . import AdvancedExcel`
  - README/examples: `from AdvancedExcel import AdvancedExcel`
- Other issues: sys.path hacks, heavy/bloated requirements.txt (many deps appear to belong to other projects), no tests, no type hints, Python-level cell loops in hot paths (e.g. `getAllRowsFromKey`), inconsistent copy vs. in-place behavior in some methods.

The split will **fix the above structural problem** as a side effect while keeping (and making reliable) the public API.

## Public API That Must Be Preserved (Non-Negotiable)

- Class name: `AdvancedExcel`
- All existing method names and signatures exactly (e.g. `removeEmptyRows(self, df)`, `getDataFrameBlocksByKeyName(self, df, key_name)`, `mergeColumns(self, df, init_col, end_col, separator, name_of_column)`, `p(data.setRowAsHeader, 0)`, etc.).
- Class attributes / constants: `ROW_INDEX`, `COL_INDEX`, `DATA`, `name`, `sheets`.
- Constructor: `AdvancedExcel(file, filetype=None)`
- Sheet loading + multi-engine behavior (openpyxl / xlrd / pyxlsb) + hyperlink support via the monkey patch.
- Return semantics (most methods return a new/copied DataFrame).
- Import surface for existing callers (`from AdvancedExcel import AdvancedExcel` and any current `import AdvancedExcel` patterns after `pip install -e .`).
- The sspipe-friendly usage shown in README (chaining via `p(instance.method, args)`).

After the refactor, **all 50+ "pretend" methods must resolve as real bound methods** on `AdvancedExcel` instances.

## Proposed Module Structure & Encapsulation

Use **mixin classes** for a clean, maintainable split. Each file defines a mixin that contributes methods. The final class inherits from all of them.

Recommended layout (under `src/`):

```
src/
├── __init__.py                 # Keep or improve re-exports for compat
├── AdvancedExcel.py            # Thin shell (or rename to core.py)
│                               #   class AdvancedExcel( ...Mixins... ): ...
├── io.py                       # Loading, engines, basic sheet access
├── lookup.py                   # Key/value "find next" navigation inside sheets
├── blocks.py                   # Multi-table / entity / block extraction (key requirement)
├── rows.py                     # All row-oriented operations
├── columns.py                  # All column-oriented operations
├── cleaning.py                 # Value/string/NaN/rename/dtype transforms
├── reshape.py                  # Transpose, melt, wide/long, df merging
└── utils.py                    # Shared private helpers + small utilities
```

### Detailed Method → Module Mapping

**Core / Class Definition (AdvancedExcel.py or core.py)**
- `class AdvancedExcel(...)` (inherits all the *Mixin classes below)
- Constants: `ROW_INDEX`, `COL_INDEX`, `DATA`
- `__init__` (can delegate loading to IO mixin or keep minimal coordinator logic here)
- `getSheet(self, sheetname)`
- `getSheetNames(self)`
- Class docstring

**io.py (IOMixin) — "Manage file/sheet loading and top-level entities"**
- Full `__init__` loading logic (engine selection by extension, `load_workbook` / `xlrd` / `pyxlsb`, loop over sheets with `pd.read_excel`, population of `self.sheets`)
- `getSheet`, `getSheetNames` (if not left in core)
- `_path_leaf`
- `_get_sheetnames`
- The `OpenpyxlReader._convert_cell = ...` monkey patch (or a dedicated `_patches.py`)
- Future: context manager, lazy loading, etc.

**lookup.py (LookupMixin) — Navigation / key-based value lookup**
- `getNextValue(self, _sheet, keyname, limitoffset=0)`
- `getNextValueFromKeyList(self, _sheet, keynames=[], limitoffset=0)`
- `getAllRowsFromKey(self, _sheet, keyname)`
- `getValueFromPosition(self, _sheet, col, row, as_string=True)`
- `getAllPairValues(self, _sheet)`
- `_getNextValue(self, list_like, initial_position, limitOffset=0)`

**blocks.py (BlockTableMixin / EntitiesMixin) — "Manage entities, tables, blocks" (core of the advanced use case)**
- `getAllTables(self, _sheet, mincol=3)`
- `getDataFrameBlocksByKeyName(self, df, key_name)`
- `getDataFrameBlocksByKeyColumn(self, df, key_column)`
- `_getTable(self, _sheet, init_table, end_table)`
- `_getAllColumnsFromKey(self, df, key_column)`
- `_headerColumnsAreEmpty(self, columns)`

**rows.py (RowMixin) — Manage rows**
- `removeEmptyRows(self, df)`
- `removeDuplicates(self, df)` (row semantics)
- `removeRowsByRegEx(self, df, columnName, regex)`
- `removeRowsNotInRegEx(self, df, columnName, regex)`
- `removeRowsWhereValueInColumnIsEmpty(self, df, columnName)`
- `removeRowsWhereValueIsEqualInColumn(self, df, columnName, rowValue)`
- `mergeRows(self, df, init_row, end_row, separator='')`
- `removeRowIfUniqueValue(self, df, init_row, end_row)`
- `removeRowsIfUniqueValueOnTop(self, df, offset)`
- `removeRowsIfUniqueValueOnBotton(self, df, offset)`
- `removeAllRowsFromAColumnNameWhenHasValue(...)`
- `removeAllRowsFromAColumnWhenHasValue(...)`
- `removeAllRowsBeforeColumnValue(self, df, init_row, index_column, value)`
- `removeRowByIndex(self, df, index_row)`
- `fillForwardRowValues(self, df, columnNames=[])`
- `setHeaderAsARow(self, df)`
- `setRowAsHeader(self, df, index_row, uppercase=False, char_space_to=" ")`
- `setRowAsHeaderByValue(self, df, column_to_search, value_to_match, ...)`
- `addHeaderAsARow(self, df)`
- `removeEmptyRows` (already listed)
- Parts of `transpose` that are row-oriented

**columns.py (ColumnMixin) — Manage columns**
- `removeEmptyColumns(self, df)`
- `removeAllColumnsWhereAnyRowHasAValue(self, df, value)`
- `removeAllColumnsWhereALLRowHasAValue(self, df, value)`
- `removeAllColumnsWhereALLRowAreEmpty(self, df)`
- `removeAllColumnsBeforeColumnValueHeader(self, df, header_value, skip_first=False)`
- `removeAllColumnsAfterColumnValueHeader(self, df, header_value, skip_first=False)`
- `removeColumnsByNameNaN(self, df)`
- `removeColumnByName(self, df, key_name, if_column_exists=None)`
- `removeColumnsByNames(self, df, column_names)`
- `removeAllColumnsExceptListOfColumnNames(self, df, column_names)`
- `mergeColumns(self, df, init_col, end_col, separator, name_of_column)`
- `mergeColumnsByName(self, df, column_destine, column_left, column_right, remove_origin=True, separator='_')`
- `splitOnColumns(self, df, column_to_split, split_character="_", list_of_columns=None)`
- `splitIn2ColumnsRight(self, df, column_to_split, columns_to_generate=None, split_character="_")`
- `splitIn2ColumnsLeft(self, df, column_to_split, columns_to_generate=None, split_character="_")`
- `moveToFirst(self, df, column_to_move_ahead)`
- `setColumn2Numeric(self, df, columnName)`
- `setListOfColumns2Numeric(self, df, columnNames=None)`

**cleaning.py (CleaningMixin) — Value / string / NaN / header transforms**
- `replaceNaNInColumnByNaNStrings(self, df, columnName)`
- `replaceNaNStrings(self, df)`
- `stripAll(self, df)`
- `removeAllCRLF(self, df, byValue='')`
- `replaceByDictionary(self, df, dictionary, column_to_apply)`
- `replaceAll(self, df, to_replace, replace_by, columns_to_clean=None, columns_not_to_clean=None)`
- `replaceAllByList(self, df, to_replace, columns_to_clean=None, columns_not_to_clean=None)`
- `replaceSpacesBySeparator(self, df, separator=';', columns_to_clean=None, columns_not_to_clean=None)`
- `transformAsNumeric(self, df, columnsToApply=None, columnsToAvoid=None, errors='coerce')`
- `fillColumnsWithLeftValue(self, df, index_row)`
- `renameHeaders(self, df, dictionary)`
- `renameHeadersByRegExp(self, df, regexp, value='')`
- `renameHeadersByRegExpList(self, df, listOfRegExp, value='')`
- `caseHeaders(self, df, uppercase=True)`
- `caseColumnValues(self, df, columnName=None, title=False, uppercase=True)`

**reshape.py (ReshapeMixin)**
- `transpose(self, df)` (note: internally calls `setRowAsHeader` — will resolve via mixin)
- `melt(self, df, column_to_melt, column_name_id=None, column_name_values="VALUE")`
- `melt_extended(self, df, column_to_generate_columns, column_name_id)`
- `mergeListDataFrames(self, df_list, column_id_name="GROUP_ID")`

**utils.py (UtilsMixin or just imported helpers)**
- `_path_leaf`, `_get_sheetnames` (can move to io)
- Any other small pure helpers that emerge during the split
- Keep the 7 private helpers close to the logic that uses them (some can stay as `_` methods on the owning mixin; others can be module-private functions called as `self._foo` or `Module._foo`).

### Wiring Strategy (for 100% Compatibility)

```python
# In AdvancedExcel.py (or core.py)
from .io import IOMixin
from .lookup import LookupMixin
from .blocks import BlockTableMixin
from .rows import RowMixin
from .columns import ColumnMixin
from .cleaning import CleaningMixin
from .reshape import ReshapeMixin
from .utils import UtilsMixin   # if needed

class AdvancedExcel(
    IOMixin,
    LookupMixin,
    BlockTableMixin,
    RowMixin,
    ColumnMixin,
    CleaningMixin,
    ReshapeMixin,
    UtilsMixin,
    object,
):
    """Main facade. All behavior comes from the mixins above."""
    ROW_INDEX = "row_index"
    COL_INDEX = "col_index"
    DATA = "data"

    # __init__ and getSheet* can live here or be pulled from IOMixin
```

- All `self.xxx()` calls inside methods will resolve at runtime on the final `AdvancedExcel` instance.
- No change to any caller code is required.
- The module-level `def foo(self, ...)` definitions will be turned into proper `def foo(self, ...)` inside their mixin class.

Alternative (if mixins feel heavy): after defining the class, do explicit attachment:

```python
from .rows import removeEmptyRows, mergeRows, ...
AdvancedExcel.removeEmptyRows = removeEmptyRows
# ... for all 51
```
Mixins are preferred for readability, IDE support, and future extension.

## Packaging & Import Surface

- Keep `from AdvancedExcel import AdvancedExcel` working after `pip install -e .` (or `pip install -e "C:\DEV\TE\AdvancedExcel"`).
- Decide on one of:
  A. Normalize everything to a proper `advanced_excel` package name (update setup.py, __init__.py, add deprecation shim for the old import name).
  B. Keep the current weird surface exactly as-is for zero breakage (recommended for this first refactor).
- Update `src/__init__.py` as needed (e.g. `from .AdvancedExcel import AdvancedExcel` or `from .core import AdvancedExcel`).
- Consider adding `pyproject.toml` later (out of scope for the minimal split).
- `requirements.txt` should be trimmed for the library itself (move dev/test/db/graph deps to extras or a separate file).

## Other Improvements (to Consider During or After the Split)

**During the split (cheap because we're touching every method):**
- Remove `sys.path.append` hacks.
- Make all methods real (the main structural fix).
- Add consistent `df = df.copy()` or clear "returns a copy" language in docstrings where it matters.
- Add a `tests/` directory + `test_smoke.py` (or `tests/test_core.py`) that exercises at least one method from each new module + a full sspipe-style pipeline + the multi-block "studies_excel" pattern using pure pandas DataFrames (no external .xlsx files required for unit tests).
- Basic type hints on public methods (`pd.DataFrame` in/out).
- `__all__` in each module.

**High-value follow-ups (can be separate PRs):**
- Vectorize hot loops (`getAllRowsFromKey` and similar) — big perf win.
- Better error messages and column-existence guards.
- Trim `requirements.txt` (core lib should not pull sqlalchemy, matplotlib, graphviz, boto3, psycopg2, ntlm, etc. unless they are actually used by this package).
- Add pytest + a few real Excel fixtures for integration-style tests.
- Richer "Table/Entity" objects instead of always returning raw DataFrames (aligns with "manage entities tables").
- GitHub Actions CI.
- Proper logging instead of silent behavior.
- Modern packaging (`pyproject.toml`, hatch or setuptools, version in one place).
- Optional: support for keeping formulas vs data_only.

## Implementation Order (Suggested)

1. Create branch (done).
2. Write this `REFACTOR_PLAN.md` (done).
3. Create the new mixin files with the grouped methods (properly indented as methods, not top-level `def xxx(self)`).
4. Update `AdvancedExcel.py` (or `core.py`) to define the class that inherits the mixins + minimal core logic.
5. Update `src/__init__.py` (and setup.py only if we decide to normalize names).
6. Delete/move the old monolithic content (or empty the old file and re-export for transition).
7. Run smoke verification:
   - `python -c "from AdvancedExcel import AdvancedExcel; ..."` (or the installed name)
   - Instantiate with a small in-memory scenario or a test .xlsx if present.
   - Exercise methods from every group.
   - Exercise a multi-step `p(data.xxx)` chain.
   - Exercise `getDataFrameBlocksByKeyName` + row/col cleaning pipeline (the "studies" pattern from README).
   - Confirm no AttributeError on the previously top-level methods.
   - `git status` + diff review.
8. Commit with clear message. Push if desired.
9. (Optional) Add basic tests + CI in follow-up commits.

## Verification Checklist (Must Pass Before "Done")

- [ ] `from AdvancedExcel import AdvancedExcel` (or equivalent) succeeds.
- [ ] `ae = AdvancedExcel(some_file); df = ae.getSheet("...")` works.
- [ ] `ae.removeEmptyRows(df)`, `ae.setRowAsHeader(df, 0)`, `ae.mergeColumns(...)`, `ae.replaceAll(...)`, `ae.transpose(df)`, `ae.melt(...)` etc. all resolve and run without AttributeError or signature breakage.
- [ ] sspipe style: `from sspipe import p; ... | p(ae.removeDuplicates) | p(ae.caseHeaders) | ...` works.
- [ ] Multi-table / block extraction still functions (`getAllTables`, `getDataFrameBlocksByKeyName`).
- [ ] The openpyxl hyperlink monkey patch is still applied.
- [ ] `ae.sheets`, `ae.name`, and the three constants are still present and behave the same.
- [ ] No regression in the 16 originally "real" methods.
- [ ] Package still installs (`pip install -e .` from the dir) and the import name(s) work as before.

## Notes / Risks

- Some methods have side-effecty index resets + column drops. We should not change observable behavior.
- Cross-mixin calls (e.g. reshape calling a row method) are safe once everything is mixed into one class.
- The original code has very few unit tests (none visible in the repo). Smoke + manual review of key examples is the main safety net.
- If any caller is doing `from AdvancedExcel import removeDuplicates` (module-level), we may want to keep a compatibility shim for one release.

---

**Status:** Analysis complete. Plan documented. Awaiting explicit approval to begin code moves on the `refactor/split-modules` branch.

When approved, the implementer will:
- Create the new `*.py` files under `src/`.
- Wire the mixins.
- Update package glue.
- Run the verification checklist.
- Leave the repo in a clean, working state.

This plan was generated after full inspection of the source (AST + full method inventory + tail of file + package files + README examples).
