import os
import argparse
import pyexcel as pe


def merge(files_dir, dest_file_name, start_row=0):
    all_rows = list()
    files = os.listdir(files_dir)
    for index, file_name in enumerate(files):
        print("Collecting Data from {0}".format(file_name))

        file_path = os.path.join(files_dir, file_name)
        rows = pe.get_array(file_name=file_path, start_row=start_row)
        all_rows += rows

    print("Merging all the Data into {0}".format(dest_file_name))
    pe.save_as(array=all_rows, dest_file_name=dest_file_name)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Merge Multiple Excel Files into One File')

    parser.add_argument('files_dir', metavar='d', type=str, help='Files Directory to be merged')
    parser.add_argument('dest_file_name', metavar='f', type=str, help='Destination File Name')
    parser.add_argument('start_row', metavar='s', type=int, help='Starting Row')

    args = parser.parse_args()

    merge(args.files_dir, args.dest_file_name, args.start_row)