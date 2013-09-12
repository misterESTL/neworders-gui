# -*- coding: utf-8 -*-
# Created on May 31, 2013
# @author: Tom Eaton
#
# This is the business logic for incoming purchase orders from Amazon.
# Orders are converted from XMLSS format into Everest Advanced Edition
# 5.0 compatible CSV header and data files.

import xml.etree.cElementTree as ET
import csv
import os
import datetime

# Only accept xls files as input.
filetypes = [('Excel File', ('.xls'))]

# Amazon specific title & Description
title = 'Amazon New Order Conversion'
description =  "This program will convert Amazon's Excel (XMLSS) purchase " \
               "orders into the format required by Everest for import." \
               "\n\nChoose the sales order file in the top field and the " \
               "directory for the output files in the second field." \
               "\n\n::Press Convert when ready::"

class ConvertOrder():
    '''This class reads orders and converts them to Everest Advanced Edition \
    format'''

    def __init__(self, paths):
        self.inFile = paths[0]
        self.dFile = paths[1] + '\Amazon New Order Data ' + \
          datetime.datetime.now().strftime('%m-%d-%Y %I%M%p') + '.csv'

        self.statusText = 'Processing...'

        order = self.readXMLSS(self.inFile)
        filteredOrder = self.filterData(order)
        self.writeCSV(self.dFile, filteredOrder)

    def removeNonASCII(self, x):
        return "".join(i for i in x if ord(i)<128)

    def writeCSV(self, filename, writetable):
        with open(filename, 'wb') as f:
            writer = csv.writer(f)
            writer.writerows(writetable)    

    def readXMLSS(self, inFile):
        tree = ET.parse(self.inFile)
        root = tree.getroot()
        line_items = root[4][0]

        tagPrefix = '{urn:schemas-microsoft-com:office:spreadsheet}'
        rowTag = tagPrefix + 'Row'
        cellTag = tagPrefix + 'Cell'
        dataTag = tagPrefix + 'Data'

        order_data = []
        for row in line_items.findall(rowTag):
            new_row = []
            for cell in row.iter(cellTag):
                for cell_data in cell.iter(dataTag):
                    new_row.append(cell_data.text)

        if new_row != []:
            order_data.append(new_row)

        return order_data

    def filterData(self, orderData):
        keepCols  = [0, 1, 2, 4, 7, 8, 13, 22] # Keep from original file
        logCols   = [0, 6, 7, 8, 1, 2, 3, 5, 4] # New order for log file
        dataCols  = [0, 8, 2, 2, 5, 4] # New order for data file
        
        filteredData = []
        for rowCount, row in enumerate(orderData):
            newRow = []
            for colCount, cell in enumerate(orderData[rowCount]):
                if rowCount in keepCols:
                    if colCount == 2:
                        newRow.append(cell)
                    elif colCount == 4:
                        newRow.append(self.removeNonASCII(cell))
                    elif colCount == 7:
                        newRow.append(float(cell.translate(None, '$, ')))
                    elif colCount == 8:
                        newRow.append(int(cell.translate(None, ', ')))
                    elif colCount == 13:
                        newRow.append(datetime.datetime.strptime(cell[:10], \
                          '%Y-%m-%d').strftime('%m/%d/%Y'))
                    else:
                        newRow.append(cell)
            if newRow != []:
                filteredData.append(newRow)

        return filteredData



'''
Old filter logic:
        order_data = []
        for row_count, row in enumerate(line_items.findall(rowTag)):
            new_row = []
            for col_count, cell in enumerate(row.iter(cellTag)):
                for cell_data in cell.iter(dataTag):
                    if row_count >= 4 and col_count in keep_cols:
                        if col_count == 2:
                            new_row.append(cell_data.text)
                        elif col_count == 4:
                            new_row.append(removeNonASCII(cell_data.text))
                        elif col_count == 7:
                            new_row.append(float(cell_data.text.translate(None, '$, '))) # Remove special characters from price.
                        elif col_count == 8:
                            new_row.append(int(cell_data.text.translate(None, ', '))) # Remove special characters from quantity.
                        elif col_count == 13:
                            new_row.append(datetime.datetime.strptime(cell_data.text[:10], '%Y-%m-%d').strftime('%m/%d/%Y')) # Convert date
                        else:
                            new_row.append(cell_data.text)
            if new_row != []:
                order_data.append(new_row)
        for row in order_data:
            print row
'''