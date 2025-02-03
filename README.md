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

``  data = AdvancedExcel(file) ``  

``  result = data.getSheet(sheetName) | p(data.getNextValue, "Study Number")
``

