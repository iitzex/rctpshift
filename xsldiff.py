# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 18:09:32 2015

@author: QN
"""

import codecs
from xlrd import *
from shift2gCal import *

def xsldiff(orig, later):
    events = []
    modified = []

    line = u'日期, 席位: 原本->後來\n'
    for row in range(3, orig.nrows):
        for col in range(3, 18): #D - J
            if (col > 9) and (col < 12):
                continue

            origCell = orig.cell(row, col).value
            laterCell = later.cell(row, col).value
            if origCell != laterCell:
                position = orig.cell(1, col).value
                date = int(orig.cell(row, 0).value)
                line += "%2d, %4s: %s -> %s\n" % (date, position, origCell, laterCell)
                events.append((date, position, origCell, laterCell))
                modified.append(origCell)
                modified.append(laterCell)

    members = sorted(set(modified))
    line = toView(members, events) + "\n" #+ line
    print line
    return line

def toView(members, events):
    line = u''
    for initial in members:
        line += "%s\n" % initial
        personal = []
        for event in events:
            if initial == event[2] or initial == event[3]:
                personal.append(event)
#        print personal

        switch = 0
        for i in range(len(personal)):
            first = personal[i]
            if switch :
                switch = 0
                continue
            if(i+1 < len(personal)):
                second = personal[i+1]
                if first[0] == second[0]:
                    switch = 1
                    if initial == first[2]:
                        line += "\t%2d:%s->%s\n" % (first[0], first[1], second[1])
                    else:
                        line += "\t%2d:%s->%s\n" % (first[0], second[1], first[1])
                elif initial == first[2] :
                    line += "\t%2d:-%s\n" % (first[0], first[1])
                elif initial == first[3] :
                    line += "\t%2d:+%s\n" % (first[0], first[1])
            else: #only for the last one
                if initial == first[2] :
                    line += "\t%2d:-%s\n" % (first[0], first[1])
                elif initial == first[3] :
                    line += "\t%2d:+%s\n" % (first[0], first[1])
    return line
def main():
    origFile = u'K_10407_v1.2_debug.xls'
    laterFile = u'K_10407_v1.3_debug.xls'
    outputFile = r'v1.2.v1.3.diff.txt'

    origSheet = open_sheet(origFile, u'班表')
    laterSheet = open_sheet(laterFile, u'班表')
    #output = codecs.open(outputFile, encoding='utf-8', mode='w')
    output = open(outputFile, 'w')

    line =  xsldiff(origSheet, laterSheet)

    output.write(line)
    output.close()

if __name__ == '__main__':
    main()
