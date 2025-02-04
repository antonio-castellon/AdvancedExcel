## AdvancedExcel: Simplifying Excel Data Processing in Python

The `AdvancedExcel` class provides a powerful and convenient way to interact with Excel files in Python.  It streamlines common Excel data processing tasks, from reading and extracting data to cleaning, transforming, and manipulating it using pandas DataFrames.  This class aims to reduce the boilerplate code often associated with Excel handling, making your data analysis workflows more efficient and readable.

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


## Dependencies

``py -m pip install --upgrade setuptools wheel``

``py -m pip install -r ./requirements.txt``

## Usage

### Basic usage 

```python 
    data = AdvancedExcel(file) ``
    result = data.getSheet(sheetName) | p(data.getNextValue, "Study Number")
```

### Example using sspipe for complex chaining:
```python
    
    from sspipe import Pipe, p
    excel_processor = AdvancedExcel("complex_excel.xlsx")
    final_df_chained = (
                        excel_processor.getSheet(sheetName)
                        | p(excel_processor.removeEmptyRows)
                        | p(excel_processor.setColumn2Numeric, columnName="Quantity")
                        | p(pd.merge, excel_processor.getSheet("Sheet1") 
                                        | p(excel_processor.removeEmptyRows) 
                                        | p(excel_processor.renameHeaders, {"OldName": "CategoryName"}), on="ProductID", how="left")
                        | p(pd.merge, excel_processor.getSheet("Sheet2") 
                                        | p(excel_processor.removeEmptyRows) 
                                        | p(excel_processor.stripAll) 
                                        | p(excel_processor.removeDuplicates), left_index=True, right_index=True) # Example merge using indexes
                        | p(lambda df: pd.concat([df, excel_processor.read_excel("complex_excel.xlsx", "Sheet3") | p(excel_processor.transpose)], axis=1))
                        )
```
### Another example of usage on a complex Excel:

```python

      data = AdvancedExcel(file)
      results = data.getSheet(sheetName) | p(data.getDataFrameBlocksByKeyName, "Batch Number")
      allTables = []
      for dblock in results: 
            study_num = dblock | p(data.getNextValue,"Study No") 
            plan_num = dblock | p(data.getNextValue,"Plan No")
            table = dblock | p(data.getDataFrameBlocksByKeyName,"Batch Number") \
                           | px[0] \
                           | p(data.replaceNaNStrings) \
                           | p(data.removeEmptyRows) \
                           | p(data.mergeRows, 0, 2) \
                           | p(data.setRowAsHeader, 0) \
                           | p(data.removeRowsIfUniqueValueOnBotton, 2) \
                           | p(data.mergeColumns, 1, 2, '_', 'SAMPLE_REPLICATE') \
                           | p(data.removeColumnByName, "Sample #_Replicat #")
                          # | p(data.removeColumnByName, '^Com.*')
            allTables.append(table)
```
