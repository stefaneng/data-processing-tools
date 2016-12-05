#!/usr/bin/env python3

from xlrd import open_workbook
import csv 
import sys
from itertools import chain

# flatmap from http://www.markhneedham.com/blog/2015/03/23/python-equivalent-to-flatmap-for-flattening-an-array-of-arrays/
# Modified for python 3

def workbook_to_dict(filename):
    wb = open_workbook(filename)

    # Store all the results as a list of dicts
    all_results = []
    for s in wb.sheets():
        for i in range(1, s.nrows):
            col_names = map(lambda x: str(x.value).lower(), s.row(0))
            data = {name: s.cell(i,col).value for name, col in zip(col_names, range(s.ncols))}
            data['sheet'] = s.name
            all_results.append(data)
    return all_results

def flatmap(f, items):
    return chain.from_iterable(map(f, items))

if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(1)

    workbook_data = list(flatmap(workbook_to_dict, sys.argv[1:]))

    # Get the full list of headers
    csv_headers = set.union(*map(lambda x: set(x), workbook_data))

    # Write data to stdout
    dict_writer = csv.DictWriter(sys.stdout, sorted(csv_headers))
    dict_writer.writeheader()
    dict_writer.writerows(workbook_data)

