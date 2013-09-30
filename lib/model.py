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


filetypes = [('Excel File', ('.xls'))]

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

        self.processOrder(self.inFile, self.dFile)

    def processOrder(self, iFile, oFile):
        self.statusText = 'Processing...'

        newOrder = self.readXMLSS(iFile)
        newOrder = self.removeNonData(newOrder)
        newOrder = self.filterData(newOrder)
        newOrder = self.removeCols(newOrder)
        self.writeCSV(oFile, newOrder)


    def removeNonASCII(self, x):
        return "".join(i for i in x if ord(i)<128)

    def readXMLSS(self, inFile):
        tree = ET.parse(self.inFile)
        root = tree.getroot()
        lineItems = root[4][0]

        tagPrefix = '{urn:schemas-microsoft-com:office:spreadsheet}'
        rowTag = tagPrefix + 'Row'
        cellTag = tagPrefix + 'Cell'
        dataTag = tagPrefix + 'Data'

        orderData = []
        for row in lineItems.findall(rowTag):
            newRow = []
            for cell in row.iter(cellTag):
                for cellData in cell.iter(dataTag):
                    newRow.append(cellData.text)

            if newRow != []:
                orderData.append(newRow)

        return orderData

    def removeNonData(self, order):
        for x in range(3):
            del order[0]

        return order

    def filterData(self, orderData):
        head = orderData.pop(0)
        filteredData = []

        for rowCount, row in enumerate(orderData):
            newRow = []
            for colCount, cell in enumerate(orderData[rowCount]):
                if colCount == 4:
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

        self.addHeader(head, filteredData)
        return filteredData

    def removeCols(self, order):
        keepCols  = [0, 1, 2, 4, 7, 8, 13, 22] # Keep from original file
        logCols   = [0, 6, 7, 8, 1, 2, 3, 5, 4] # New order for log file
        dataCols  = [0, 8, 2, 2, 5, 4] # New order for data file

        return order

    def addHeader(self, head, data):
        data.insert(0, head)
        return data   

    def filterCols(self, order):
        pass

    def reorderCols(self, order):
        pass

    def writeCSV(self, filename, table):
        with open(filename, 'wb') as f:
            writer = csv.writer(f)
            writer.writerows(table)    