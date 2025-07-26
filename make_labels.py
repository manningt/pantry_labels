#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.12"
# dependencies = [
#   "fpdf2",
#   "typing_extensions",
# ]
# ///

'''
Note: before loading uv, the first line was: #!/usr/bin/env python3

This script generates PDF labels for guests based on their item counts and delivery or pickup times.

The inputs are 3 CSV files:
1. Visits_with_Tallied_Inventory_Distribution.csv - contains the item counts for each guest.
2. Guest list CSV files - contain the first name, last name, route or pickup time, and item count for each guest.
   There is one one csv file for delivery and one for AM/PM pickup.
   For the pickup lists, AM pickup times are Saturday morning, and PM pickup times are Friday afternoon.

3 functions:
1. make_full_guest_dict(in_filename) - reads the Visits_with_Tallied_Inventory_Distribution.csv file
   and returns a dictionary with the guest names as keys and their item counts as values.
2. make_guest_list(in_filename, guest_dict, start_time=12, end_time=15) - reads a guest list CSV file and
    returns a list of tuples with the guest's first name, last name, route or pickup time, and item count.
3. make_label_pdfs(guest_list, type, out_pdf_path) - generates a PDF file with labels for each guest in the guest list.
   The number of pages in the label PDF file is determined by the item count for each guest.

The output is are the following files:
1. guest_list_0_Delivery.pdf - contains labels for all guests with delivery times.
2. guest_list_0_Pickup_Saturday.pdf - contains labels for all guests with AM pickup times (Saturday morning).
3. guest_list_0_Pickup_Friday_before_3.pdf - contains labels for all guests with PM pickup times (Friday afternoon).
4. guest_list_0_Pickup_Friday_after_3.pdf - contains labels for all guests with PM pickup times (Friday evening).
The output files are saved in the current directory.
The script also generates a report file make_tags_report.txt with the status of the label generation.

'''

import sys, os
import glob
try:
   import csv
   from fpdf import FPDF
except Exception as e:
   print(e)
   sys.exit(1)

import argparse

# from flask import current_app
# from typing import NamedTuple  #not to be confused with namedtuple in collections

DELIVERY_TYPE = 'Delivery'  # used for delivery guest lists
AM_PM_TYPE = 'AM_PM'  # used for AM/PM guest lists

# full_guest_dict has item count
def make_full_guest_dict(in_filename):
   guest_dict = {}

   with open(in_filename, newline='') as csvfile:
      reader = csv.DictReader(csvfile)
      if 'Client' not in reader.fieldnames:
         sys.exit(f"Failure: {in_filename} does not have a Client column.")
      # print(f"{reader.fieldnames=}")
      row_count = 0
      for row in reader:
         # full_name = row['Client'].strip()
         row_count += 1
         try:
            name_key = f"{row['Client'].split(',')[1]}_{row['Client'].split(',')[0]}".strip()
         except:
            print(f"Warning: row {row_count} has no client name in the second column.")
            continue
            
         try:
            guest_dict[name_key] = int(float(row['Total Quantity']))
         except:
            guest_dict[name_key] = 1
            print(f"Warning: {name_key} has no item count; defaulting to {guest_dict[name_key]}.")
         # if row_count == 3:
         #    break
   return guest_dict


def get_guest_list_type(in_filename):
   # Determine if the guest list is for delivery or AM/PM pickup
   with open(in_filename, newline='') as csvfile:
      reader = csv.DictReader(csvfile)
      if 'Route or Pickup Time' not in reader.fieldnames:
         sys.exit(f"Failure: {in_filename} does not have a Route or Pickup Time column.")
      row = next(reader)
      if ':' in row['Route or Pickup Time'] and (' AM' in row['Route or Pickup Time'] or ' PM' in row['Route or Pickup Time']):
         return AM_PM_TYPE
      else:
         return DELIVERY_TYPE
      

def make_guest_list(in_filename, guest_dict, start_time=12, end_time=15):
   # time format: 12:50 PM or 07:00 AM
   guest_list = []

   # print(f"make_guest_list {in_filename} {AM_PM_String=}")

   with open(in_filename, newline='') as csvfile:
      reader = csv.DictReader(csvfile)
      # print(f"{reader.fieldnames=}")
      row_count = 0
      for row in reader:

         if row_count == 0:
            if (':' in row['Route or Pickup Time'] and (' AM' in row['Route or Pickup Time'] or ' PM' in row['Route or Pickup Time'])):
               type = AM_PM_TYPE
            else:
               type = DELIVERY_TYPE

         name_key = row['First'] + '_' + row['Last'].replace('*', '')
         if name_key in guest_dict:
            item_count = guest_dict[name_key]
         else:
            item_count = 1
            print(f"Warning: {name_key} not found in guest_dict; defaulting item_count to {item_count}.")

         if type == DELIVERY_TYPE:
            guest_list.append((row['First'], row['Last'], row['Route or Pickup Time'], item_count))
         else:
            time_string = row['Route or Pickup Time'].split(' ')[0]
            hour, minute = map(int, time_string.split(':'))
            am_pm = row['Route or Pickup Time'].split(' ')[1]
            if am_pm == 'PM' and hour < 12:
               hour += 12
            if hour >= start_time and hour < end_time:
               guest_list.append((row['First'].title(), row['Last'].title(), row['Route or Pickup Time'], item_count))

         # print(f"  {row_count} = {guest_list[-1]}")
         # if row_count == 2:
         #    break
         row_count += 1
   return guest_list


def item_count_to_label_count(item_count):
   limits = [0, 9, 17, 25, 32, 40, 49,57,67,73,82,90,100,107,112,139,200]
   if item_count > 200:
      return 16
   for i in range(len(limits)-1):
      if item_count > limits[i] and item_count <= limits[i+1]:
         return i + 1


def make_label_pdfs(guest_list, type, pdf_filename, output_directory="."):

   # PDF writing examples:
   #  https://medium.com/@mahijain9211/creating-a-python-class-for-generating-pdf-tables-from-a-pandas-dataframe-using-fpdf2-c0eb4b88355c
   #  https://py-pdf.github.io/fpdf2/Tutorial.html
   route_font_size = 28 # allows longer names
   name_font_size = 36
   label_count_font_size = 12
   label_height = 144 #points
   label_width = 288
   number_of_labels = 0
   cell_width = 0
   cell_height = 0

   if len(guest_list) == 0:
      status_string = f"Failure: no guests in the guest_list to generate {out_pdf_path}."
   else:
      try:
         pdf = FPDF(orientation="L", unit="pt", format=(label_height,label_width))
      except Exception as e:
         status_string = f"Failure: could not create PDF for {out_pdf_path} exception: {e}"
         return status_string
      
      try:
         pdf.set_margins(0, 18, 0) #left, top, right in points
         pdf.set_auto_page_break(auto=False)
         pdf.set_font("Helvetica", "B") # Arial not available in fpdf2
         for row in guest_list:
            label_count = int(item_count_to_label_count(row[3]))
            # print(f"  {row[0]} {row[1]} {row[2]} has {row[3]} items, which is {label_count} labels.")
            for i in range(label_count):
               pdf.add_page()
               # if row[2] is a time, then don't print it; only print if it's a route
               if type == DELIVERY_TYPE:
                  pdf.set_font_size(route_font_size)
                  pdf.cell(cell_width, cell_height, f'{row[2].replace(" - ", ": ")}', align="L")
                  pdf.line(0, 36, label_width, 36) # line from left to right
               pdf.ln(route_font_size+10)
               pdf.set_font_size(name_font_size)
               pdf.cell(cell_width, cell_height, f"{row[0].title()}", align="C")
               pdf.ln(name_font_size+4)
               pdf.cell(cell_width, cell_height, f"{row[1][0:15].title()}", align="C")
               pdf.ln(name_font_size+4)
               pdf.set_font_size(label_count_font_size)
               pdf.cell(cell_width, cell_height, f"{i+1} of {label_count}", align="R")
               number_of_labels += 1
      except Exception as e:
         status_string = f"Failure: while adding cells for {out_pdf_path} exception: {e}"
         return status_string
      
      try:
         out_pdf_path = os.path.join(output_directory, pdf_filename)
         pdf.output(out_pdf_path)
         status_string = f"{pdf_filename} has {len(guest_list)} guests and {number_of_labels} labels."
      except Exception as e:
         #    current_app.logger.warning(f"PDF for {guest} failed: {e}")
         status_string = f"failed to generate {pdf_filename} exception: {e}"

   return status_string


def test_label_pdfs(out_pdf_path):
   route_font_size = 28 # allows longer names
   name_font_size = 36
   item_count_font_size = 12
   try:
      pdf = FPDF(orientation="L", unit="pt", format=(144,288)) # default units are mm; heigth, width are in points - 72 points = 1 inch
      pdf.set_margins(0, 10, 0) #left, top, right in points
      pdf.set_auto_page_break(auto=False)
      pdf.set_font("Helvetica", "B") # Arial not available in fpdf2
      pdf.add_page()
      pdf.set_font_size(route_font_size)
      pdf.cell(0, None, f"Route67890123456789012", align="L", border=1)
      pdf.ln(route_font_size+8)
      pdf.set_font_size(name_font_size)
      pdf.cell(0, None, f"First", align="C", border=1)
      pdf.ln(name_font_size+4)
      pdf.cell(0, None, f"Last", align="C", border=1)
      pdf.ln(name_font_size+4)
      pdf.set_font_size(item_count_font_size)
      pdf.cell(0, None, f"1 of 8", align="R", border=1)
      pdf.output(out_pdf_path)

   except Exception as e:
      print(f"PDF for test_label failed: {e}")


def write_report_file(guest_list, report_filename, output_directory="."):
   text_report_filename = report_filename.replace('.pdf', '.txt')
   text_report_path = os.path.join(output_directory, report_filename.replace('.pdf', '.txt'))
   try:
      with open(text_report_path, "w") as f:
         f.write(f"\n{text_report_filename}\n")
         # f.write(f"  First       Last     Time     Bags\n")
         for guest in guest_list:
            f.write(f"  {guest[0]:<12} {guest[1]:<12}     Time={guest[2]}     bags={guest[3]}\n")
   except Exception as e:
      print(f"Failed to write report file {report_filename}: {e}")  


def process_files(file_list, output_directory="."):

   if not os.path.isdir(output_directory):
      sys.exit(f"Failure: '{output_directory}' does not exists or is not a directory.")

   STRING_IN_ITEM_COUNT_FILENAME = "Tallied"
   STRING_IN_DELIVERY_FILENAME = "Delivery"
   STRING_IN_PICKUP_FILENAME = "Pickup"

   item_count_index = None
   delivery_index = None
   pickup_index = None

   full_guest_dict = {}
   report_strings = []

   for i in range(len(file_list)):
      if not os.path.isfile(file_list[i]):
         sys.exit(f"{file_list[i]} is not a file.")
      if STRING_IN_ITEM_COUNT_FILENAME in file_list[i]:
         item_count_index = i
      elif STRING_IN_DELIVERY_FILENAME in file_list[i]:
         delivery_index = i
      elif STRING_IN_PICKUP_FILENAME in file_list[i]:
         pickup_index = i
      else:
         print(f"Warning: '{file_list[i]}' does not match any expected file name.")

   if delivery_index is None and pickup_index is None:
      sys.exit("Failure: Neither Delivery or Pickup csv files were found.")

   if item_count_index is None:
      sys.exit("Failure: No 'Visits_with_Tallied_Inventory_Distribution' file found.")
   else:
      full_guest_dict = make_full_guest_dict(file_list[item_count_index])
      if 0:
         print(f"{full_guest_dict=}")
         sys.exit(0)

   if len(full_guest_dict) == 0:
      sys.exit("Failure: 'Visits_with_Tallied_Inventory_Distribution.csv' had no guests.")

   if delivery_index is not None:
      delivery_filename = file_list[delivery_index]
      delivery_guest_list = make_guest_list(delivery_filename, full_guest_dict)
      if len(delivery_guest_list) == 0:
            status_string = f"Warning: no guests in {delivery_filename}."
      else:
         delivery_pdf_filename = f'tags-for-{STRING_IN_DELIVERY_FILENAME}.pdf'
         status_string = make_label_pdfs(delivery_guest_list, DELIVERY_TYPE, delivery_pdf_filename, output_directory)
         write_report_file(delivery_guest_list, delivery_pdf_filename, output_directory)
      print(status_string)
      report_strings.append(status_string)

   if pickup_index is not None:
      timeslots_dict = {
         'Saturday': (7, 12),
         'Friday-before-3': (12, 15),
         'Friday-after-3': (15, 23)}
      for timeslot, (start, end) in timeslots_dict.items():
         pickup_pdf_filename = f'tags-for-{STRING_IN_PICKUP_FILENAME}-{timeslot}.pdf'
         pickup_guest_list = make_guest_list(file_list[pickup_index], full_guest_dict, start_time=start, end_time=end)
         if len(pickup_guest_list) == 0:
            status_string = f"Warning: no guests in {pickup_pdf_filename}."
         else:
            pickup_guest_list.sort(key=lambda x: (x[1], x[0]))
            status_string = make_label_pdfs(pickup_guest_list, type, pickup_pdf_filename, output_directory)
            write_report_file(pickup_guest_list, pickup_pdf_filename, output_directory)
         print(status_string)
         report_strings.append(status_string)

   text_report_path = os.path.join(output_directory, "make_tags_report.txt")
   with open(text_report_path, "w") as report_file:
      for line in report_strings:
         report_file.write(line + "\n")


if __name__ == "__main__":

   if 0:
      test_array = [1,11,17,25,32,40,49,57,67,73,82,90,100,107,112,139,200, 201]
      for item_count in test_array:
         print(f"{item_count} is {item_count_to_label_count(item_count)} labels")   
      sys.exit(0)

   if 0:
      test_label_pdfs("/tmp/test_label.pdf")
      sys.exit(0)

   argParser = argparse.ArgumentParser()
   argParser.add_argument("file_path", type=str, help="input filename with path", nargs='*')

   '''
   There can be multiple files:
      One must be the Visits_with_Tallied_Inventory_Distribution file, which has the item counts for each guest.
         The Inventory Distribution file is parsed into the full_guest_dict, which has the LAST_FIRST as the key and the item count as the value.

      The rest are the guest lists, which have the first name, last name, route or pickup time, and item count.
      guest_filename_list is a list of filenames that are guest lists.

      guests_lists[] is a list of lists, where each list is a guest list, typically one for delivery and one for pickup.
   '''

   args = argParser.parse_args()
   file_list= args.file_path

   if len(file_list) == 0:
      print("Using current directory for CSV files.")
      file_list = glob.glob('*.csv')
   # print(f"{file_list=}")

   process_files(file_list)
   # process_files(file_list, output_directory="/tmp")
