#!/usr/bin/env python3
'''
2 functions:
1. make_guests_per_caller_lists(in_filename) -> Caller_lists
2. make_caller_pdfs(caller_mapping_dict, guest_dict, date_str)

The input file is an Excel file with 3 sheets: 'guest-to-caller', 'callers', 'guests'
See class definition below for the NamedTuple Caller_lists returned by make_guests_per_caller_lists(in_filename)

The output is a PDF file for each caller with a list of guests for the next Friday.

'''

import sys, os
try:
   from openpyxl import load_workbook
   import csv
   from fpdf import FPDF
except Exception as e:
   print(e)
   sys.exit(1)

import argparse
from pathlib import Path
from flask import current_app

from typing import NamedTuple  #not to be confused with namedtuple in collections

class guest_lists(NamedTuple):
    success: bool = False
    message: str = ''
    guest_list: list = []


def make_guest_list(in_filename):
   # Pickup_lists()
   guest_list = []

   with open(in_filename, newline='') as csvfile:
      reader = csv.DictReader(csvfile)
      # print(f"{reader.fieldnames=}")
      row_count = 0
      for row in reader:
         guest_list.append((row['First'], row['Last'], row['Route or Pickup Time']))
         # print(row['First'], row['Last'], row['Route or Pickup Time'])
         if row_count == 2:
            break
         row_count += 1
   return guest_list


def make_label_pdfs(guest_list, out_pdf_path, type='pickup'):
   # PDF writing examples:
   #  https://medium.com/@mahijain9211/creating-a-python-class-for-generating-pdf-tables-from-a-pandas-dataframe-using-fpdf2-c0eb4b88355c
   #  https://py-pdf.github.io/fpdf2/Tutorial.html
   font_size = 24
   line_height = 24
   try:
      pdf = FPDF(orientation="L", unit="pt", format=(144,288)) # default units are mm; heigth, width are in points - 72 points = 1 inch
      pdf.set_margins(0, 5, 0) #left, top, right in points
      pdf.set_font("Helvetica", "B", size=font_size) # Arial not available in fpdf2
      for row in guest_list:
         pdf.add_page()
         if type == 'delivery':
            pdf.cell(0, font_size, f"{row[2]}", align="C")
         pdf.ln(line_height)
         pdf.cell(0, font_size, f"{row[0]}", align="C")
         pdf.ln(line_height)
         pdf.cell(0, font_size, f"{row[1]}", align="C")
      pdf.output(out_pdf_path)

   except Exception as e:
      # try:
      #    current_app.logger.warning(f"PDF for {guest} failed: {e}")
      # except:
      print(f"PDF for {last}_{first} failed: {e}")

if __name__ == "__main__":
   argParser = argparse.ArgumentParser()
   argParser.add_argument("friday_pickups_filename", type=str, help="input filename with path")
   argParser.add_argument("saturday_pickups_filename", type=str, help="input filename with path")
   argParser.add_argument("delivery_filename", type=str, help="input filename with path")

   args = argParser.parse_args()

   guest_filename_list = []
   if args.friday_pickups_filename is None:
      sys.exit("Missing friday_pickups_filename.")
   elif Path(args.friday_pickups_filename).is_file():
      guest_filename_list.append(args.friday_pickups_filename)
   else:
      sys.exit("friday_pickups_filename is not a file.")

   if args.saturday_pickups_filename is None:
      sys.exit("Missing saturday_pickups_filename.")
   elif Path(args.saturday_pickups_filename).is_file():
      guest_filename_list.append(args.saturday_pickups_filename)
   else:
      sys.exit("saturday_pickups_filename is not a file.")

   if args.delivery_filename is None:
      sys.exit("Missing delivery_filename.")
   elif Path(args.delivery_filename).is_file():
      guest_filename_list.append(args.delivery_filename)
   else:
      sys.exit("delivery_filename is not a file.")


   guest_lists = [None] * len(guest_filename_list)
   for i in range(len(guest_filename_list)):
      guest_lists[i] = make_guest_list(guest_filename_list[i])
      if len(guest_lists[i]) == 0:
         print(f"Failure: Pickup_lists {i} had no guests.")
         sys.exit(1)
      # print(f"{guest_lists[i]=}")
      if i == 2:
         list_type = 'delivery'
      else:
         list_type = 'pickup'
      make_label_pdfs(guest_lists[i], f"/tmp/guest_list_{i}.pdf", type=list_type)
