# Excelprocess
## About this script
To deal with some issues on excel files, specified as xlsx, xlsm files etc.. The script is coded in Python 2.7 and tested in Python 2.7.15rc1. In Python 3, as the 'xrange' function is replaced by 'range', and variable type subdivided into bytes-like object and string and something else, some error will raise. As for Python3 version is not availble now, you have to fix these bugs if you have to run it in Python3 
For more information and support, please contact Maxwell Huang <410116635@qq.com>

## Prerequisites
* tqdm
* openpyxl

You may have to install the packages mentioned above before running this script by running

```shell
pip install tqdm
pip install openpyxl
```
## Usage
### Testing
With no files missing, you can test the script by the command

```shell
python excelprocess.py
```

You should get 3 csv files in the same directory, named as 'test_Sheet1.csv', 'test_Sheet2.csv' and 'test_Sheet3.csv', after the message 'File test.xlsx Processed Successfully' is broadcasted in the console.

### File mode
You can tell the script to deal with the file you delare by command

```shell
python excelprocess.py --file yourfile.xlsx
```

This command means to process the yourfile.xlsx.

### Folder mode
You can tell the script the folder in which all the xlsx files (including all files in the subdirectories) will be processed with command

```shell
python excelprocess.py --folder ./
```

This command means to process all the xlsx files under the directory ./, a.k.a the directory where the script is.

### Function Selection
You can determine which function to be used, which will give different process on the xlsx files.
So far, function list is:
* 1 transform and split(default)
* 2 average per six columns

```shell
python excelprocess.py --folder ./ --function 2
```

This command means to process all the xlsx files within function 2 rather than the function 1 as default under the directory ./, a.k.a the directory where the script is.

### Help
You can also look up for help with command

```shell
python excelprocess.py -h 
```
or
```shell
python excelprocess.py --help
```
