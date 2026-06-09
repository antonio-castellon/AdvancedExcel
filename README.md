## AdvancedExcel: Simplifying Excel Data Processing in Python

**v2.0+ note:** The public API now uses PEP 8 `snake_case` method names (e.g. `get_sheet`, `remove_empty_rows`, `get_dataframe_blocks_by_key_name`). This is a breaking change from v1.x.

The `AdvancedExcel` class provides a powerful and convenient way to interact with Excel files in Python. It streamlines common Excel data processing tasks, from reading and extracting data to cleaning, transforming, and manipulating it using pandas DataFrames. This class aims to reduce the boilerplate code often associated with Excel handling, making your data analysis workflows more efficient and readable.

**Key Strengths:**

* **Effortless Data Extraction:**  Easily read data from Excel sheets, including handling complexities like merged cells, multiple tables within a sheet, and various data types.  Target specific tables using row and column indices or header values.
* **Comprehensive DataFrame Manipulation:**  A rich set of methods for cleaning, transforming, and manipulating pandas DataFrames.  Perform operations like removing rows/columns based on criteria, filling missing values using various strategies (forward fill, left fill), merging and splitting columns, changing data types, and more.
* **Flexible Header and Index Management:**  Gain fine-grained control over headers and indices.  Set rows as headers, add headers as rows, reset indices, and drop unnecessary index levels.
* **Streamlined String and Value Processing:**  Simplify string manipulation with methods for stripping whitespace, replacing values (including using regular expressions), and changing the case of strings (uppercase, lowercase, title case).
* **Convenient Utility Functions:**  Includes helper functions for common tasks such as transposing DataFrames, merging multiple DataFrames, extracting filenames from paths, and getting sheet names without fully opening the Excel file.
* **Seamless Integration with `sspipe`:** Designed for smooth integration with the `sspipe` library (https://github.com/sspipe/sspipe). Chain method calls in a readable and functional style, creating elegant and concise data pipelines.

**Why use `AdvancedExcel`?**

Working with Excel files in Python often involves repetitive and verbose code. `AdvancedExcel` encapsulates these common tasks into reusable methods, promoting code reusability, readability, and maintainability.  It handles many of the edge cases and complexities of Excel data, allowing you to focus on the core logic of your data analysis.  The integration with `sspipe` further enhances code clarity by enabling a functional programming paradigm.

# Getting Started

Create a Virtual Environment.

`py -m venv venv`

Activate the virtual environment

`venv\Scripts\activate.bat` (in Linux `source ./venv/bin/activate` )

Update PIP

`py -m pip install -U pip`


## Installation

```bash
pip install advanced-excel
```

For the functional piping examples using `sspipe`:

```bash
pip install "advanced-excel[sspipe]"
```

For development:

```bash
pip install -e ".[dev,sspipe]"
```

## Dependencies

Core dependencies (pandas, openpyxl, etc.) are declared in `pyproject.toml`. Optional extras are available for `sspipe` support and development.

## Usage

### Basic usage 

```python
from advanced_excel import AdvancedExcel

data = AdvancedExcel("simple_excel.xlsx")
sheet = data.get_sheet("Data")
result = data.get_next_value(sheet, "Study Number")
print(result)
```

> Note: The recommended import is `from advanced_excel import AdvancedExcel` (this is the stable API as of v2.0).

### Example using sspipe for complex chaining (requires `pip install "advanced-excel[sspipe]"`):

```python
from sspipe import p
import pandas as pd
from advanced_excel import AdvancedExcel

excel_processor = AdvancedExcel("complex_excel.xlsx")

raw = excel_processor.get_sheet("Raw data")
sheet1 = excel_processor.get_sheet("Sheet1")
sheet2 = excel_processor.get_sheet("Sheet2")

final_df_chained = (
    raw
    | p(excel_processor.remove_empty_rows)
    | p(excel_processor.set_column_to_numeric, columnName="Quantity")
    | p(pd.merge,
        sheet1
        | p(excel_processor.remove_empty_rows)
        | p(excel_processor.rename_headers, {"OldName": "CategoryName"}),
        on="ProductID", how="left")
    | p(pd.merge,
        sheet2
        | p(excel_processor.remove_empty_rows)
        | p(excel_processor.strip_all)
        | p(excel_processor.remove_duplicates),
        left_index=True, right_index=True)
)
```
### Another example of usage on a complex Excel (multiple tables per sheet):

```python
from advanced_excel import AdvancedExcel
from sspipe import p   # optional, for nice chaining

data = AdvancedExcel("studies_excel.xlsx")
results = data.get_sheet("Raw data") | p(data.get_dataframe_blocks_by_key_name, "Batch Number")

all_tables = []
for dblock in results:
    study_num = data.get_next_value(dblock, "Study No")
    plan_num = data.get_next_value(dblock, "Plan No")

    table = (
        dblock
        | p(data.get_dataframe_blocks_by_key_name, "Batch Number")
        | (lambda blocks: blocks[0] if blocks else pd.DataFrame())
        | p(data.replace_nan_strings)
        | p(data.remove_empty_rows)
        | p(data.merge_rows, 0, 2)
        | p(data.set_row_as_header, 0)
        | p(data.remove_rows_if_unique_value_on_bottom, 2)
        | p(data.merge_columns, 1, 2, '_', 'SAMPLE_REPLICATE')
        | p(data.remove_column_by_name, "Sample #_Replicat #")
    )
    all_tables.append(table)
```

```python
from advanced_excel import AdvancedExcel
from sspipe import p
import pandas as pd

data = AdvancedExcel("studies_excel.xlsx")

columns_to_rename = {"PLATEFORM": "PLATFORM"}
columns_not_to_clean = ["PARAMETER"]

results = (
    data.get_sheet("data_raw")
    | p(data.case_headers)
    | p(data.rename_headers, columns_to_rename)
    | p(data.strip_all)
    | p(data.remove_column_by_name, r"^T\d+$")
    | p(data.remove_column_by_name, r"^T\d+\.\d+$")
    | p(data.remove_all_crlf, ";")
    | p(data.replace_nan_strings)
    | p(data.replace_spaces_by_separator, ";", [], columns_not_to_clean)
    | p(data.replace_all, ",;", ";", [], columns_not_to_clean)
    | p(data.replace_all, "tbd", "TBD", [], [])
    # | p(data.replace_by_dictionary, unit_map, "UNIT")
    | p(data.case_column_values, "PARAMETER", title=True)
    # | p(data.replace_by_dictionary, variable_map, "PARAMETER")
)
```
