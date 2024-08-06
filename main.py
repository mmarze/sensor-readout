import argparse
from pathlib import Path
import sys
import os
import re
import pandas as pd
from datetime import datetime
import functions


def main():
    # ------ Create a parser ------
    parser = argparse.ArgumentParser(description="Check the status of methane sensors")
    parser.add_argument("--folder_path", type=str, help="Path to the folder with '.csv' files to import")
    parser.add_argument("--file_paths", type=str, nargs='+', help="List of'.csv' files paths to import")
    parser.add_argument("--print_results", type=bool,
                        help="Sensor status printed on the console if True; default = True")
    parser.add_argument("--save_results", type=bool,
                        help="Sensor status saved to '.csv' file if True; default = False")
    parser.add_argument("--date_from", type=str, help="Start date, format: DD-MM-YYYY")
    parser.add_argument("--date_to", type=str, help="End date, format: DD-MM-YYYY")
    args = parser.parse_args()

    # close app if no paths to files/directory provided
    if args.folder_path is None and args.file_paths is None:
        print('No data to process. The application is closed.')
        sys.exit()

    # convert path strings to paths
    if args.folder_path:
        args.folder_path = Path(args.folder_path)
    if args.file_paths:
        for i, elem in enumerate(args.file_paths):
            args.file_paths[i] = Path(elem)

    # convert date strings to datetime
    if args.date_from:
        args.date_from = datetime.strptime(args.date_from, '%d-%m-%Y')
    if args.date_to:
        args.date_to = datetime.strptime(args.date_to, "%d-%m-%Y")

    # default values for print_results and save_results
    if args.print_results is None:
        args.print_results = True
    if args.save_results is None:
        args.save_results = False

    # ------ read all files to pandas.DataFrames ------
    # if path to the folder
    if args.folder_path:
        # read all csv files from the directory
        my_files_list = [f for f in os.listdir(args.folder_path) if re.match(r'\d\d\.\d\d\.\d\d\d\d\.csv', f)]
        # delete out-of-analysis files
        my_files_list = functions.select_files(my_files_list, args.date_from, args.date_to)
        # filepaths to all filtered files
        my_file_paths = []
        for elem in my_files_list:
            my_file_paths.append(os.path.join(args.folder_path, elem))
        # concatenate files
        my_sensor_data = functions.create_dataset(my_file_paths, my_files_list)
    # else - list of paths to files
    else:
        # create a list with names of all files
        my_files_list = []
        for elem in args.file_paths:
            head, tail = os.path.split(elem)
            my_files_list.append(tail)
        # select out-of-analysis files
        my_files_list = functions.select_files(my_files_list, args.date_from, args.date_to)
        # delete out-of-analysis files
        for i, elem in enumerate(args.file_paths):
            head, tail = os.path.split(elem)
            if tail not in my_files_list:
                args.file_paths[i] = ''
        args.file_paths = [elem for elem in args.file_paths if elem != '']
        # concatenate files
        my_sensor_data = functions.create_dataset(args.file_paths, my_files_list)

    # get sensor status
    my_sensor_data = functions.get_sensor_status(my_sensor_data)
    # get formatted pandas.DataFrame with sensor status
    my_result_data = functions.get_result_data(my_sensor_data)

    # print results_data DataFrame in console
    if args.print_results:
        print(my_result_data)

    # save results_data DataFrame to xlsx file:
    if args.save_results:
        # prepare filename
        my_filename = functions.generate_filename(my_files_list)
        # Create a Pandas Excel writer using XlsxWriter as the engine
        with pd.ExcelWriter(my_filename, engine='xlsxwriter') as writer:
            # Convert the dataframe to an XlsxWriter Excel object.
            my_result_data.to_excel(writer, sheet_name='Arkusz1', startrow=1, startcol=1, index=True)
        functions.format_excel_file(my_filename,  my_result_data.shape[0], my_result_data.shape[1])


if __name__ == '__main__':
    main()
