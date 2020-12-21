# FILL WEB FORM USING VBA AND PYTHON

This example fills a web form using data in Excel. 
We use VBA to call the exe file. This file is created with python and then converted to an exe file.
The Python file uses the win32com library to read the data from the Excel file and, Selenium to fill the web form.
After the task is complete, one cell is changed and control returns to Excel VBA macro.


![FILL WEB FORM](https://github.com/josemaria500/VBA/blob/main/FILL_WEB_FORM/vba_python.gif)

## Files explanation
| File | Description |
| ------ | ------ |
| demo.xlsm | Excel file with the data to fill the web form and macro VBA.   |
| Módulo1.bas | Th Main procedure, call exe file. |
| Hoja3.cls | Sheet code that detects changes in cells. |
| demo.py | Python file used to create the exe file. |
| demo.exe | File that reads Excel and fills the web form. |
| chromedriver.exe | Download the appropriate version at: [chromedriver](https://chromedriver.chromium.org/downloads) |