# -*- coding: utf-8 -*-
"""
Created on Tue Jun  9 10:00:32 2015

@author: QN

"""
from qndropbox import download, check_month
from shift2gCal import *


def main():
    files = check_month()
    for filename in files:
        try:
            download(filename)
            parse_update(filename)
        except:
            pass


def parse_update(filename):
    # members = ['QN']
    members = ['QN', 'NN', 'HJ', 'BT', 'XP', 'MN', 'EH']
    members += ['BJ', 'FW', 'VB', 'ZD', 'TS', 'GQ']
    members += ['EZ', 'BZ', 'AS', 'EC', 'ZD', 'AW']
    members = sorted(members)

    config_sheet = open_sheet(filename, u'設定')
    year, month = get_config(config_sheet)
    shift_sheet = open_sheet(filename, u'班表')
    affair_sheet = open_sheet(filename, u'人力一覽')

    service = get_service()

    #    #get shifts
    for initial in members:
        calId = list_calendarlist(service, initial)
        del_month_event(service, calId, year, month)
        shifts = parse_excel(initial, shift_sheet, year, month)
        # print shifts
        for (position, start, end, description) in shifts:
            insert_event(service, calId, position, start, end, description)

        affairs = parse_affair(initial, affair_sheet, year, month)
        #        print affairs
        if affairs is not None:
            for (pos, start, end, affair) in affairs:
                insert_event(service, calId, pos, start, end, affair)

    # store log
    now = datetime.now().isoformat()
    logId = list_calendarlist(service, 'LOG')
    insert_event(service, logId, now, now, now, filename)


if __name__ == '__main__':
    main()
