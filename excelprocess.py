from __future__ import division
from __future__ import print_function
import argparse
import os
import csv
from time import sleep
from tqdm import tqdm
from openpyxl import load_workbook

def folder_process(folder_address):
    '''Find out all the xlsx files in the folder with 
    going into all the sub-directories, and process'''

    file_list = []
    list_dirs = os.walk(folder_address)
    for root, dirs, files in list_dirs:
        for f in files:
            try:
                if f[-5:] == ".xlsx" or f[-5:] == ".XLSX":
                    file_list.append(os.path.join(root, f))
            except Exception as e:
                raise e

    pbar = tqdm(total=len(file_list),dynamic_ncols=True,unit='file')

    for temp_file in file_list:
        pbar.set_description(desc='Processing: 'temp_file)
        file_process(temp_file)
        sleep(0.01)
        pbar.update(1)
        
    pbar.close()



def file_process(file_name):
    '''Read the target file and do the processing'''

    try:
        work_book = load_workbook(filename=file_name)
        pass

    except IOError as e:
        print("Error: No such file, please check out the file name")

    else:
        # Read by columns and write by rows 
        sheet_names = work_book.get_sheet_names()
        for temp_sheet_name in sheet_names:
            with open(file_name[:-5]+'_'+temp_sheet_name+'.csv','wb') as tempfile:
                tempwriter = csv.writer(tempfile)
                temp_sheet = work_book[temp_sheet_name]
                for col in temp_sheet.iter_cols(values_only=True):
                	tempwriter.writerow(col)

        print('File '+file_name+' Processed Successfully')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='excelprocess')
    parser.add_argument('--file', type=str, default='test.xlsx', help='determine which excel file to be processed, for example: extract.py --file example.xlsx')
    parser.add_argument('--folder', type=str, help='determine the folder where all the xlsx files will to be processed, for example: extract.py --folder ./address')
    # parser.add_argument('--functions', type=str, nargs='+', default=['transform','sheet_split'], help='determine what to do on the original files, \nfunctions include: transform, sheet_split')

    arg = parser.parse_args()

    if arg.folder:
        folder_process(arg.folder)

    else:
        file_process(arg.file)