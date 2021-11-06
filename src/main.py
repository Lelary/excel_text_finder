
import pathlib
import sys
import text_finder
from os import listdir
from os.path import isfile, join

def find_from_file(file_path, sheet_name, targets, find_from_header):
    finder = text_finder.Finder(file_path, sheet_name, targets, find_from_header)
    print(finder.is_valid())

    finder.run()

    if find_from_header:
        print(finder.found_headers)
    else:
        [print(x.coordinate, ":", x.value) for x in finder.found_values]


def find_from_directory(file_path, sheet_name, targets, find_from_header):
    excel_suffixes = ['.xls', '.xlsx']
    files = [f for f in listdir(dir_path) if isfile(join(dir_path, f))]
    files = [f for f in files if pathlib.Path(f).suffix in excel_suffixes]

    for file_path in files:
        finder = text_finder.Finder(join(dir_path, file_path), sheet_name, targets, find_from_header)
        if not finder.is_valid():
            continue

        finder.run()

        print('file : ', file_path)
        if find_from_header:
            print(finder.found_headers)
        else:
            [print(x.coordinate, ":", x.value) for x in finder.found_values]



if __name__ == '__main__':
    if len(sys.argv) >= 4:
        dir_path = sys.argv[1]
        sheet_name = sys.argv[2]
        targets = [sys.argv[3]]

        #find_from_file(dir_path, sheet_name, targets, True)
        #find_from_file(dir_path, sheet_name, targets, False)
        find_from_directory(dir_path, sheet_name, targets, True)
        #find_from_directory(dir_path, sheet_name, targets, False)