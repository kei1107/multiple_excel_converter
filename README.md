# multiple_excel_converter

Script to output Excel files in a specified folder as a single pdf file.

## Requirements

- Windows + Python3
- Office

## Setup

```shell
pip install -r requirements.txt
```

## Usage

```shell
usage: multiple_excel_converter.py [-h] [--verbose] [--disable-recursive-search] [--fit-page] [--set-header] [--set-footer] [--disable-workspace-deletion] directory

convert multiple excel to the one pdf)

positional arguments:
  directory             excel directory

options:
  -h, --help            show this help message and exit
  --verbose             verbose output
  --disable-recursive-search
                        Disable recursive excel file search.
  --fit-page            fit page mode
  --set-header          set header(file name)
  --set-footer          set footer(file name)
  --disable-workspace-deletion
                        Disable workspace (tmp_*****) deletion.
```
