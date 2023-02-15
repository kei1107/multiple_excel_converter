#!/usr/bin/python3
import argparse
import datetime
from natsort import natsorted
from hashids import Hashids
import os
import pprint
import PyPDF2
import shutil
import win32com.client

EXCEL_BOOK_EXT = [
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".xls",
    ".xlw"
]

PDF_TYPE = 0


def get_excel_file_list(dir: str, disable_recursive_search: bool = False) -> list[tuple[str, str]]:
  output_list = []
  for f in natsorted(os.listdir(dir)):
    if os.path.isdir(os.path.join(dir, f)):
      if not disable_recursive_search:
        sub_dir_list = get_excel_file_list(os.path.join(dir, f))
        output_list.extend(sub_dir_list)
    else:
      # if os.path.isdir(f)
      name, ext = os.path.splitext(f)
      if ext in EXCEL_BOOK_EXT and not name.startswith('~$'):
        output_list.append((f, os.path.join(dir, f)))
  return output_list


def main():
  parser = argparse.ArgumentParser(
      description='convert multiple excel to the one pdf)')
  parser.add_argument('directory', type=str, help='excel directory')
  parser.add_argument('--verbose', action='store_true', help='verbose output')
  parser.add_argument('--disable-recursive-search', action='store_true',
                      help='Disable recursive excel file search.')
  parser.add_argument('--fit-page', action='store_true',
                      help='fit page mode')
  parser.add_argument('--set-header', action='store_true',
                      help='set header(file name)')
  parser.add_argument('--set-footer', action='store_true',
                      help='set footer(file name)')
  parser.add_argument('--disable-workspace-deletion', action='store_true', 
                      help='Disable workspace (tmp_*****) deletion.')
  args = parser.parse_args()

  directory_abs = os.path.abspath(args.directory)
  verbose = args.verbose
  disable_recursive_search = args.disable_recursive_search
  fit_page = args.fit_page
  set_header = args.set_header
  set_footer = args.set_footer
  disable_workspace_deletion = args.disable_workspace_deletion

  if not os.path.isdir(directory_abs):
    print(f"${directory_abs} is not directory.")
    parser.print_help()
    exit(1)

  excel_file_list = get_excel_file_list(
      directory_abs, disable_recursive_search=disable_recursive_search)
  if verbose:
    pprint.pprint(excel_file_list)

  dt_now = datetime.datetime.now()
  hashids = Hashids(min_length=8)
  working_dir = os.path.join(
      directory_abs,
      f"tmp_{hashids.encode(dt_now.hour,dt_now.minute,dt_now.second)}")
  excel_app = win32com.client.Dispatch("Excel.Application")
  excel_app.Visible = False
  excel_app.DisplayAlerts = False

  merger = PyPDF2.PdfMerger()

  if verbose:
    print(f"Create workspace : {working_dir}")
  os.makedirs(working_dir, exist_ok=True)

  try:
    counter = 0
    for excel_name, abs_excel_file in excel_file_list:
      if verbose:
        print(f"Working... in {abs_excel_file}")
      tmp_pdf_name = os.path.join(working_dir, f"{counter}.pdf")

      wb = excel_app.Workbooks.Open(abs_excel_file)
      if verbose:
        print(f"Create {tmp_pdf_name} for {abs_excel_file}")

      if any([fit_page, set_footer, set_header]):
        sheets_count = wb.Sheets.Count
        for sheet_idx in range(1, sheets_count+1):
          ws = wb.Worksheets(sheet_idx)
          ws.Activate()
          if fit_page:
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = 1
            ws.PageSetup.CenterHorizontally = True
          if set_header:
            ws.PageSetup.CenterHeader = excel_name
          if set_footer:
            ws.PageSetup.CenterFooter = excel_name
      wb.ExportAsFixedFormat(PDF_TYPE, tmp_pdf_name)
      try:
        # for safety close
        wb.SaveAs(os.path.join(working_dir, excel_name))
        wb.Close(False)
      except Exception:
        print(f"[WARNING] Unsafe Close : {abs_excel_file}")

      merger.append(tmp_pdf_name)
      counter += 1

    output_pdf_name = os.path.join(directory_abs, "output.pdf")
    merger.write(output_pdf_name)
    print(f"Created {output_pdf_name}")
  finally:
    excel_app.quit()
    merger.close()
    if not disable_workspace_deletion:
      if os.path.exists(working_dir):
        if verbose:
          print(f"Remove workspace : {working_dir}")
        shutil.rmtree(working_dir)

if __name__ == '__main__':
  main()
