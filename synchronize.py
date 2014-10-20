#!/usr/bin/env python
#
# Copyright (c) 2014 Steven MARTINS
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

__author__ = "Steven MARTINS"

from pydrive.auth import GoogleAuth

import gsp
import gsp.exceptions
import re

from local_config import config

class Synchronize(object):
    def __init__(self):
        self.g = gsp.login(config["username"], config["password"])
        self.sheet = None
        self.worksheet = None
        self.openSheetByKey(config["worksheet_key"])
        self.openWorkSheetByIndex(config["sheet_pos"])
    
    def openSheetByKey(self, key):
        self.sheet = self.g.open_by_key(key)

    def openSheetByName(self, name):
        self.sheet = self.g.open(name)
    
    def openWorkSheetByIndex(self, index):
        self.worksheet = self.sheet.get_worksheet(index)

    def openWorkSheetByName(self, name):
        self.worksheet = self.sheet.worksheet(name)

    def getRow(self, row):
        return self.worksheet.row_values(row)

    def getCol(self, col):
        return self.worksheet.col_values(col)

    def getCell(self, row, col):
        return self.worksheet.cell(row, col)

    def getACell(self, label):
        return self.worksheet.acell(label)

    def getRowAsDict(self, row, row_hdr=1):
        d = {}
        row = self.getRow(row)
        hdr = self.getTop(row_hdr)
        for i in range(min(len(hdr), len(row))):
            d[hdr[i]] = row[i]
        return d

    def getTop(self, row=1):
        return self.getRow(row)

    def getRows(self):
        return self.worksheet.get_all_values()

    def getAllRows(self):
        return self.worksheet.get_all_records()

    def rowCount(self):
        return self.worksheet.row_count

    def colCount(self):
        return self.worksheet.col_count

    def update(self, cells):
        return self.worksheet.update_cells(cells)

    def addRows(self, nb):
        return self.worksheet.add_rows(nb)

    def addCols(self, nb):
        return self.worksheet.add_cols(nb)

    def appendRow(self, values):
        if isinstance(values, dict):
            v = []
            hdr = self.getTop()
            for h in hdr:
                v.append(values[h] if h in values else '')
            values = v
        first_col = self.getCol(1)
        if len(first_col) < self.rowCount():
            return self.updateRowByIndex(len(first_col) + 1, values)
        return self.worksheet.append_row(values)

    def find(self, q):
        return self.worksheet.find(q)

    def findRow(self, query):
        try:
            cell = self.find(query)
            if cell and cell.row:
                return self.getRow(cell.row)
        except gsp.exceptions.CellNotFound:
            pass
        return None

    def getAddrFromInt(self, row, col):
        return self.worksheet.get_addr_int(row, col)

    def getIntFromAddr(self, label):
        return self.worksheet.get_int_addr(label)

    def getRange(self, r):
        return self.worksheet.range(r)

    def updateRowByIndex(self, index, values, autocommit=True):
        if isinstance(values, dict):
            v = []
            hdr = self.getTop()
            for h in hdr:
                v.append(values[h] if h in values else '')
            values = v
        i = 1
        ran = self.getRange("%s:%s" % (self.getAddrFromInt(index, 1), 
                                       self.getAddrFromInt(index, len(values))))
        for cell in ran:
            print("row(%s), col(%s), val(%s)" % (cell.row, cell.col, values[cell.col - 1]))
            cell.value = values[cell.col - 1]
            i += 1
        if autocommit:
            self.worksheet.update_cells(ran)
        return ran

    def updateRowByField(self, name, values, autocommit=True):
        try:
            cell = self.find(name)
            if cell and cell.row:
                diff = [x for x in self.getRow(cell.row) if x not in values]
                if len(diff) > 0:
                    return self.updateRowByIndex(cell.row, values, autocommit)
        except gsp.exceptions.CellNotFound:
            return self.appendRow(values)
        return None

    def wipeCells(self):
        nb_cols = self.colCount()
        nb_rows = self.rowCount()
        split_by = 40
        i = 0
        print("nb_cols(%s), nb_rows(%s)" % (nb_cols, nb_rows))

        while i < round(nb_rows / split_by) + 1:
            _from = i * split_by
            _end = _from + split_by if _from + split_by <= nb_rows else nb_rows
            cells = self.getRange("%s:%s" % (self.getAddrFromInt(_from + 1, 1), 
                                             self.getAddrFromInt(_end, nb_cols)))
            for cell in cells:
                cell.value = ""
            self.worksheet.update_cells(cells)
            i += 1
        

    def fastLoad(self, rows, name_idx=0):
        self.wipeCells()
        # update col Size and row Size
        nb_cols = len(max(rows, key=lambda row: len(row)))
        nb_rows = len(rows)
        if nb_rows < self.rowCount():
            self.addRows(nb_rows - self.rowCount())
        if nb_cols < self.colCount():
            self.addCols(nb_cols - self.colCount())

        split_by = 40
        i = 0
        print("nb_cols(%s), nb_rows(%s)" % (nb_cols, nb_rows))

        while i < round(nb_rows / split_by) + 1:
            _from = i * split_by
            _end = _from + split_by if _from + split_by <= nb_rows else nb_rows
            print("i(%s), _from(%s), _end(%s)" % (i, _from, _end))
            cells = self.getRange("%s:%s" % (self.getAddrFromInt(_from + 1, 1), 
                                             self.getAddrFromInt(_end, nb_cols)))
            print(len(cells))
            for cell in cells:
                row = rows[cell.row - 1]
                cell.value = row[cell.col - 1]
            self.worksheet.update_cells(cells)
            i += 1

    def load(self, rows, name_idx=0):
        update = []
        if len(rows) < self.rowCount():
            self.addRows(rows - self.rowCount() + 1)
        if isinstance(rows, dict):
            for k, v in rows:
                self.updateRowByField(k, v)
        else:
            i = 0
            for row in rows:
                res = self.updateRowByField(row[name_idx], row, False)
                if res:
                    update += res
                if i % 5 == 0 and len(update) > 0:
                    self.worksheet.update_cells(update)
                    update = []
                i += 1
        if len(update) > 0:
            self.worksheet.update_cells(update)


    def export(self, as_dict=True, key_name="A1"):
        if not as_dict:
            return self.getAllRows()
        key = self.getACell(key_name).value
        ret = {}
        for row in self.getAllRows():
            ret[row[key]] = row
        return ret

if __name__ == "__main__":
    o = Synchronize()
    #print(str(o.getRow(1)))
    #print(str(o.getRowAsDict(2)))
    #print(str(o.getAllRows()))
    o.updateRowByField("Cjour01", {
        "slug": "Cjour01",
        "type": "directory",
        "path": "Piscine",
        "triche": "TRUE",
        "rule": "01",
        "correction": "TRUE"
    })
