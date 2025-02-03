# Introduction 

Repository that contains a global Python class that allows 
to manipulate Excel files inside Pipeline processes.

Please, consider to use an isolated installation of python in your local computer. 
i.e : `C:\SDK\Python\3.7_64` 

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
    
    final_df_chained = (
    excel_processor.read_excel("complex_excel.xlsx", "Sheet1", init_row=2, end_row=10)
    | p(excel_processor.removeEmptyRows)
    | p(excel_processor.setColumn2Numeric, columnName="Quantity")
    | p(pd.merge, excel_processor.read_excel("complex_excel.xlsx", "Sheet1", init_row=15, end_row=20) | p(excel_processor.removeEmptyRows) | p(excel_processor.renameHeaders, {"OldName": "CategoryName"}), on="ProductID", how="left")
    | p(pd.merge, excel_processor.read_excel("complex_excel.xlsx", "Sheet2") | p(excel_processor.removeEmptyRows) | p(excel_processor.stripAll) | p(excel_processor.removeDuplicates), left_index=True, right_index=True) # Example merge using indexes
    | p(lambda df: pd.concat([df, excel_processor.read_excel("complex_excel.xlsx", "Sheet3") | p(excel_processor.transpose)], axis=1))
    | p(excel_processor.write_excel, "chained_processed_data.xlsx", "FinalData") # Write the data
    )
```
### Another example of usage on a complex Excel:

```python

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
