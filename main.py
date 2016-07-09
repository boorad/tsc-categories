#!/usr/bin/env python

from openpyxl import load_workbook
from openpyxl.styles import Font
import os, re, sys, datetime

PAST_TITLES=False
CURR_SUBTOTAL=""
EOF=False
CURR_ROW=0


def datetype(val):
    return type(val) is datetime.datetime or type(val) is datetime.date


def print_row(row, wk, yr):
    global CURR_ROW

    if CURR_ROW > 14:
        return

    CURR_ROW = CURR_ROW + 1
    for cell in row:
        value = cell.value
        print yr, wk, value, type(value), repr(value)
    print


def process_row(row, wk, yr):
    global PAST_TITLES, CURR_SUBTOTAL, EOF, CURR_ROW

#    if CURR_ROW > 23:
#        return

    a = row[0].value
    if row[0].font.bold:
        subtotal = True
        CURR_SUBTOTAL=a
    else:
        subtotal = False

    if a[:5] == "TOTAL": # we are at the end of the file
        EOF=True
        return

    if PAST_TITLES and not subtotal:
        CURR_ROW = CURR_ROW + 1
        print "{},{},{},{}".format(yr, wk, CURR_SUBTOTAL, ",".join([str(cell.value) for cell in row]))

    if a == "Group":  # check after the row printing, because we don't need the titles row
        PAST_TITLES=True


def process_file(fn, wk, yr):
    global PAST_TITLES, EOF, CURR_ROW
    PAST_TITLES=False
    EOF=False
    CURR_ROW = 0

    wb = load_workbook(filename=fn)
    ws = wb['Report']

    # process rows
    for row in ws.rows:
        if EOF:
            return

        process_row(row, wk, yr)
#        print_row(row, wk, yr)


def parse_filename(fn):
    p = re.compile(r'[Cc]ategory[Ww]eek(?P<wk>[0-9]+)(?P<yr>[0-9]{4}).xlsx')
    m = p.search(fn)
    return m.group("wk"), m.group("yr")


def main(d="data/"):
    for subdir, dirs, files in os.walk(d):
        for file in files:
            filepath = subdir + os.sep + file
            if filepath.endswith(".xlsx"):
                wk, yr = parse_filename(file)
                process_file(filepath, wk, yr)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print "usage: {} directory".format(sys.argv[0])
    else:
        main(sys.argv[1])
