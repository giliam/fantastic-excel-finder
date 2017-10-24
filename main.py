import argparse
import os

import pandas as pd
from colorama import init as colorama_init
from colorama import Fore, Back, Style

colorama_init()

parser = argparse.ArgumentParser(description='Find a string in Excel files located in input folder.')
parser.add_argument('match', metavar='s', type=str, 
                    help='the string to find in files')
args = parser.parse_args()

for dirpath, _, filenames in os.walk('input/'):
    for filename in filenames:
        xlsx = pd.ExcelFile(os.path.join(dirpath, filename))
        for sheet_name in xlsx.sheet_names:
            sheet = pd.read_excel(xlsx, sheet_name, header=None)
            if args.match in str(sheet):
                for i, row in enumerate(sheet.index):
                    for j, col in enumerate(sheet.columns):
                        if args.match in str(sheet.get_value(row, col)):
                            print("Found in", Fore.CYAN + sheet.get_value(row, col) + Style.RESET_ALL, "\n", Fore.YELLOW + "%s%s" % (chr(col+65), row+1) + Fore.WHITE, "\t\tsheet", Fore.GREEN + sheet_name + Fore.WHITE, "\t\tfile", Fore.RED + filename + Fore.WHITE)