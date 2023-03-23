# Excel Reader Script
This script is designed to read an Excel file and add data from a text file to specific cells in the spreadsheet. The script was written using Python and the [openpyxl](https://openpyxl.readthedocs.io/en/stable/) library for manipulating Excel files.

### Installation
1. Install Python on your system if it is not already installed. You can download Python from the official website: https://www.python.org/downloads/

2. Install the openpyxl library by running the following command in your terminal:

```sh
pip install -r requirements.txt
```

### Usage
1. Clone this repository or download the script excel_reader.py to your computer.

2. Open a terminal window and navigate to the directory where the script is located.

3. Run the script by typing the following command:

```sh
python excel_reader.py
```

4. The GUI window will appear. Click the "Выбери excel" button to select the Excel file you want to manipulate.

5. Click the "Выберите txt файл" button to select the text file containing the data you want to add to the Excel file.

6. Enter the cell name where you want to add the data in the "Введите номер ячейки:" field.

7. Click the "Выполнить" button to execute the script and add the data to the specified cell in the Excel file.
