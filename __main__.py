#!/usr/bin/python
import fnmatch
import zipfile
import pandas as pd
import argparse
import os
import xlsxwriter

_data = []


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('file', type=input_ok, help='should be a .zip file')
    options = parser.parse_args()
    process(options.file)
    print("Wait for processing.. may take 1-2 minutes max")
    data_saver()
    print("Done.")
    current_path = os.getcwd() + r'\DDR.xlsx'
    print(f"File is at {current_path}")


def input_ok(file):
    if not os.path.exists(file):
        raise argparse.ArgumentTypeError(f"Filename {file} doesn't exists in this directory.")
    if file[-4:] != ".zip" and not zipfile.is_zipfile(file):
        raise argparse.ArgumentTypeError(f"{file} is not a .zip file.")
    return file


def process(file):
    with zipfile.ZipFile(file) as zfile:
        for filename in zfile.namelist():
            if fnmatch.fnmatch(filename, r'messages\*\messages.csv'):
                data = zfile.open(filename)
                df = pd.read_csv(data)
                _data.append(df)
            else:
                pass


def data_saver():
    writer = pd.ExcelWriter('DDR.xlsx', engine='xlsxwriter')
    i = 0
    while i < len(_data):
        df = pd.DataFrame({'ID': _data[i].ID, 'Timestamp': _data[i].Timestamp, 'Contents': _data[i].Contents})
        df['ID'] = df['ID'].apply(lambda x: '{:.0f}'.format(x))
        df['Timestamp'] = df['Timestamp'].str[:-13]
        df.sort_values(by=['Timestamp'])
        df.to_excel(writer, sheet_name='Messages')

        i += 1

    writer.save()


if __name__ == "__main__":
    main()
