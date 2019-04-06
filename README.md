# Excelprocess
## About this script
To deal with some issues on excel files, specified as xlsx, xlsm files etc.. The script is coded in Python 2.7 and tested in Python 2.7.15rc1, but it should also be effective under Python 3. 
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

### Help
You can also look up for help with command

```shell
python excelprocess.py -h 
```
or
```shell
python excelprocess.py --help
```
