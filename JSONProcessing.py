import pprint
import json
import csv
import xlsxwriter


class FileManagement:
    def __init__(self):
        self.json_file_name = "input.json"
        self.excel_filename = "outputXls.xlsx"
        self.csv_filename = "outputCsv.csv"
        self.array_x = []
        self.array_y = []
        self.array_sum = []

    def read_text_file(self):
        try:
            with open(self.json_file_name) as data_file:
                data = json.load(data_file)
                for each_axis in data:
                    x = str(each_axis["color"])
                    y = str(each_axis["value"])
                    self.array_x.append(x)
                    self.array_y.append(y)
                    self.array_sum.append(x + y)
            return 1
        except FileNotFoundError:
            print("An error occurred while reading the file")
            return -1

    def parse_csv_file(self):
        source_file = open(self.json_file_name)
        data = json.load(source_file)
        output_csv = open(self.csv_filename, "w")

        output_writer = csv.writer(output_csv)

        print(data[0].keys())
        row_array = []
        for key in data[0].keys():
            row_array.append(key)
        output_writer.writerow(row_array)
        pprint.pprint(row_array)

        for element in data:
            row_array = []
            for attribute in element:
                row_array.append(element[attribute])
            output_writer.writerow(row_array)

    def parse_xlsx_file(self):
        workbook = xlsxwriter.Workbook(self.excel_filename)
        worksheet = workbook.add_worksheet()
        for index, value in enumerate(self.array_x):            
            worksheet.write(index, 0, self.array_x[index])
            worksheet.write(index, 1, self.array_y[index])
            worksheet.write(index, 2, self.array_sum[index])
        workbook.close()


if __name__ == '__main__':
    file_management = FileManagement()
    success = file_management.read_text_file()
    if success:
        print("Would you like to export file as..." + "\n" +
              "1. CSV" + "\n" +
              "2. Excel")
        print("Choose an option:", end='')
        option = int(input())
        if option == 1:
            file_management.parse_csv_file()
        elif option == 2:
            file_management.parse_xlsx_file()
    else:
        print("Unexpected error occurred, exiting")
