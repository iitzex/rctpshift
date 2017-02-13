# -*- coding: utf-8 -*-
import dropbox
from datetime import date


def check_month():
    this_month = date.today()
    if this_month.month + 1 > 12:
        next_month = date(this_month.year + 1, 1, 1)
    else:
        next_month = date(this_month.year, this_month.month + 1, 1)

    last = [u'塔臺%3d%02d班表.xlsx' % (this_month.year - 1911, this_month.month)]

    this = []
    #3 days left to next month
    left_day = abs(next_month - this_month).days
    if left_day <= 5:
        this += [u'塔臺%3d%02d班表.xlsx' % (next_month.year - 1911, next_month.month)]

    this += last
    return this


def download(filename):
#   get token from dropbox
    token = ''

    dpx = dropbox.Dropbox(token)

    filepath = u'/RCTP TWR/班表/修正班表/'
#    filename = u'塔臺104%02d班表.xlsx' % date.today().month
    try:
        metadata = dpx.files_download_to_file(filename, filepath+filename)
    except:
        import sys
        print("Unexpected error:", sys.exc_info()[0])
        raise

    print(metadata)

if __name__ == '__main__':
    monthes = check_month()
    print(monthes)
    for month in monthes:
        print(month)
        try:
            download(month)
        except:
            pass
