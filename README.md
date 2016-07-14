# xlsx2csv
python program to convert a `.xlsx` workbook into a folder of `.csv` files, one for each worksheet.  
This came in handy quite a bit during automated data collection: `.xlsx` files are more convenient for humans than machines.

## Usage

    xlsx2csv --file='my_workbook.xlsx'

This command will create a new folder named `my_workbook`.  Then, for each worksheet in the workbook, a `csv` file
will be generated, with the title of that worksheet. 
