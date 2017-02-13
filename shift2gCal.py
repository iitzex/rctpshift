# -*- coding: utf-8 -*-
"""
Created on Tue Jun  9 10:00:32 2015

@author: QN

"""

from datetime import timedelta
from xlrd import *
from gCal import *


def open_sheet(filename, sheetname):
    return open_workbook(filename).sheet_by_name(sheetname)


def parse_excel(initial, sheet, year, month):
    shifts = []
    shift_count = 0
    hour_count = 0

    for row in range(3, sheet.nrows):
        for col in range(3, 22):  # D - V
            if 9 < col < 12 or 16 < col < 19:
                continue

            if sheet.cell(row, col).value == initial:
                # print col
                position = sheet.cell(1, col).value
                daystr = int(sheet.cell(row, 0).value)
                position, start, end, hours = update_shift_inteval(position, year, month, daystr)
                hour_count += hours
                if col < 18:  # bypass shift could
                    shift_count += 1

                if col < 11:
                    others = sheet.row_slice(row, 2, 10)
                else:
                    others = sheet.row_slice(row, 11, 17)
                    # print others
                    others += sheet.row_slice(row, 9, 10)
                    # print others

                des = u''
                for other in others:
                    des += other.value + ", "
                # print des
                des = des[:-2]
                des += u"\n班數:%d" % shift_count
                des += u", 時數:%d" % hour_count

                position = initial + "." + position
                info = (position, start, end, des)
                shifts.append(info)

    return shifts


def update_shift_inteval(pos, year, month, day):
    start_d = date(year, month, day)

    if pos[0] == 'N' or pos == 'TT7':
        end_d = date(year, month, day) + timedelta(1)
    else:
        end_d = date(year, month, day)

    if pos == "TT1" or pos == "TT3" or pos == 'TT2':
        start_t = time(7, 30)
        end_t = time(18, 30)
        hours = 11
    elif pos == "TT4":
        start_t = time(7, 30)
        end_t = time(17, 30)
        hours = 10
    elif pos == "TT5":
        start_t = time(8, 00)
        end_t = time(18, 00)
        hours = 10
    elif pos == "TT6":
        start_t = time(8, 00)
        end_t = time(17, 00)
        hours = 9
    elif pos == "TT7":
        start_t = time(13, 00)
        end_t = time(7, 00)
        hours = 11
    elif pos == "NTT1":
        start_t = time(19, 00)
        end_t = time(8, 00)
        hours = 13
    elif pos == "NTT2":
        start_t = time(18, 00)
        end_t = time(6, 00)
        hours = 12
    elif pos == "NTT3":
        start_t = time(18, 00)
        end_t = time(8, 00)
        hours = 14
    elif pos == "NTT4":
        start_t = time(18, 30)
        end_t = time(8, 30)
        hours = 14
    elif pos == "NTT5":
        start_t = time(18, 00)
        end_t = time(8, 00)
        hours = 14
    else:
        start_t = time(0, 00)
        end_t = time(0, 00)
        hours = 0

    start_time = datetime.combine(start_d, start_t).isoformat()
    end_time = datetime.combine(end_d, end_t).isoformat()

    return pos, start_time, end_time, hours


def get_config(config):
    year = int(config.cell(0, 2).value)
    month = int(config.cell(1, 2).value)

    return year, month


def parse_affair(initial, sheet, year, month):
    affairs = []

    for col in range(4, 35):  # E - AC
        if initial != sheet.cell(1, col).value:
            continue

        for row in range(2, sheet.nrows - 1):
            affair = sheet.cell(row, col).value
            day = int(sheet.cell(row, 0).value)

            if affair == '':
                continue

            (start, end, des) = update_affair_inteval(affair, sheet, year, month, day)
            summary = initial + '.' + affair
            info = (summary, start, end, des)

            affairs.append(info)
        return affairs


def update_affair_inteval(affair, sheet, year, month, day):
    for row in range(1, 25):
        start_d = date(year, month, day)

        start_t = time(9, 00)
        end_t = time(17, 00)

        if affair[0] == u'上':
            end_t = time(12, 00)

        elif affair[0] == u'下':
            start_t = time(13, 00)

        start_time = datetime.combine(start_d, start_t).isoformat()
        end_time = datetime.combine(start_d, end_t).isoformat()

        return start_time, end_time, affair


def main():
    filename = u'塔臺10506班表.xlsx'

    config_sheet = open_sheet(filename, u'設定')
    (year, month) = get_config(config_sheet)
    shift_sheet = open_sheet(filename, u'班表')
    affair_sheet = open_sheet(filename, u'人力一覽')

    members = ['QN']
    for initial in members:
        shifts = parse_excel(initial, shift_sheet, year, month)

        # print shifts
        for item in shifts:
            print(item)
            # insert_event(service, calId, position, start, end, descripition)

        affairs = parse_affair(initial, affair_sheet, year, month)
        # print affairs
        for (pos, start, end, affair) in affairs:
            print(pos, start, end, affair)
            # insert_event(service, calId, affair, start, end)


if __name__ == '__main__':
    main()
