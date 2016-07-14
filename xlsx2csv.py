from openpyxl import *  # reading xlsx
import csv  # writing the csv
import argparse  # parsing args


class xlsx2csv:
    # multiple constructors in one.
    # yeah, this is 100% an abuse of dynamic languages.
    def __init__(self, filename_or_workbook, filename=None):
        if type(filename_or_workbook) is Workbook:
            self.workbook = filename_or_workbook
            self.dir_name = filename[:-5]
        else:
            self.dir_name = filename_or_workbook[:-5]  # trim off the `.xlsx` extension
            self.workbook = load_workbook(filename=filename_or_workbook, read_only=True)

    def execute(self):
        if not os.path.exists(self.dir_name):
            os.makedirs(self.dir_name)
        for sheet in self.workbook.worksheets:
            with open(self.dir_name + '/' + sheet.title + '.csv', 'w') as f:
                print('converting sheet with name: ' + sheet.title)
                c = csv.writer(f)
                for r in sheet.rows:
                    row = []
                    for cell in r:
                        row.append(cell.value)
                    c.writerow(row)


def parse_args():
    parser = argparse.ArgumentParser(description="Converts a given .xlsx file into a folder of .csv files")
    parser.add_argument('--file', type=str,
                        help="file to convert", required=True)
    return parser.parse_args()

# Execution begins here if PressureSensor.py is called through the interpreter.
if __name__ == '__main__':
    fn = parse_args().file
    instance = xlsx2csv(filename_or_workbook=fn)
    instance.execute()
