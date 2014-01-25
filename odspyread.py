#!/usr/bin/python

#system
import sys
import os
from optparse import OptionParser # deprecated

#local
import odf.opendocument
from odf.table import *
from odf.text import P

#globals

def log(arg, file = 1):
  if file == 1:
    sys.stdout.write(arg + "\n")
  elif file == 2:
    sys.stderr.write(arg + "\n")

def cell2text(cell):
  text = ""
  textNodes = cell.getElementsByType(P)
  for tn in textNodes:
    for tn2 in tn.childNodes:
#      if (tn.nodeType == 3): # avoid non-textNode types
      text = text + unicode(tn2.data)
  return text

def table2array(table):
  data = []
  rows = table.getElementsByType(TableRow)
  for row in rows:
    lRepeated = int(row.getAttribute("numberrowsrepeated") or 1)
    cells = row.getElementsByType(TableCell)
    values = []
    for cell in cells:
      lRepeated2 = int(cell.getAttribute("numbercolumnsrepeated") or 1)
      text = cell2text(cell)
      values.append(text)
      if lRepeated2 > 1:
        for _ in range(lRepeated2 - 1):
          values.append(text)
    for _ in range(lRepeated):
      data.append(values)
  return data

def getTable(sheet, lRowStart, lColumnStart):
  table = None
  # look for pivot
  lFields = 0
  bHeader = False
  rows = sheet.getElementsByType(TableRow)
  lRow = 1
  lRowEmptyCount = 0
  for row in rows:
    lColumn = 1
    if lRowEmptyCount >= options.lSeparationRowCount:
      # end of table
      break
    if lRow >= lRowStart:
      cells = row.getElementsByType(TableCell)
      if len(cells) == 1 and cell2text(cells[0]) == "":
        # empty row (cell repeated)
        lRepeated = int(row.getAttribute("numberrowsrepeated") or 1)
        lRow += lRepeated
        if bHeader:
          lRowEmptyCount += lRepeated
        continue
      tr = TableRow()
      if not bHeader:
        bHeader = True
        table = Table()
        bData = False
        for cell in cells:
          lRepeated = int(cell.getAttribute("numbercolumnsrepeated") or 1)
          if lColumn >= lColumnStart:
            text = cell2text(cell)
            if text != "":
              if not bData:
                bData = True
                lColumnFirst = lColumn;
              table.addElement(TableColumn(numbercolumnsrepeated = lRepeated))
              tr.addElement(cell)
              lFields = lFields + lRepeated
            else:
              if bData:
                # end of table header
                break
#              else:
#                # skip
          lColumn += lRepeated
        table.addElement(tr)
      else:
        bEmpty = True
        for cell in cells:
          lRepeated = int(cell.getAttribute("numbercolumnsrepeated") or 1)
          if ((lColumn >= lColumnFirst or lColumn + lRepeated - 1 >= lColumnFirst) and
             lColumn <= lColumnFirst + lFields - 1):
            # trim overlap
            if lColumn < lColumnFirst and lColumn + lRepeated - 1 >= lColumnFirst:
              lRepeated = lColumn + lRepeated - lColumnFirst
              lColumn = lColumnFirst
              cell.setAttribute("numbercolumnsrepeated", str(lRepeated))
            if lColumn + lRepeated - 1 > lColumnFirst + lFields - 1:
              lRepeated = lColumnFirst + lFields - lColumn
              cell.setAttribute("numbercolumnsrepeated", str(lRepeated))
            if bEmpty:
              if cell2text(cell) != "":
                bEmpty = False
            tr.addElement(cell)
          lColumn += lRepeated
        if bEmpty:
          lRepeated = int(row.getAttribute("numberrowsrepeated") or 1)
          lRowEmptyCount += lRepeated
        else:
          if lRowEmptyCount > 0:
            tr2 = TableRow(numberrowsrepeated = lRowEmptyCount)
            tr2.addElement(TableCell(numbercolumnsrepeated = lFields))
            table.addElement(tr2)
            lRowEmptyCount = 0
          table.addElement(tr)

    lRepeated = int(row.getAttribute("numberrowsrepeated") or 1)
    lRow += lRepeated
  return table

#args
optionParser = OptionParser()
def optionsListCallback(option, opt, value, parser):
  setattr(optionParser.values, option.dest, tuple(value.split(',')))
def optionsPathExpansionCallback(option, opt, value, parser):
  setattr(optionParser.values, option.dest, os.path.expandvars(os.path.expanduser(value)).split(','))
optionParser.add_option("-d", "--document", metavar = "DOC", default = "",
                        type = "string", dest = "sDoc",
                        help = "the spreadsheet path")
optionParser.add_option("-e", "--sheet", metavar = "SHEET", default = "", 
                        type = "string", dest = "sSheet",
                        help = "[optional] sheet name of interest [default: first sheet]")
optionParser.add_option("-i", "--idx", metavar = "IDX", default = "",
                        type = "string", dest = "sKeyName",
                        help = "[optional] name of the index/key field for searching in")
optionParser.add_option("-s", "--search", metavar = "SEARCH", default = ("*"),
                        type = "string", dest = "sKeyValues",
                        action = "callback", callback = optionsListCallback,
                        help = "[optional] comma-delimited list of value(s) to search for in the index/key field [default: *]")
optionParser.add_option("-m", "--allow-duplicates", metavar = "DUPLICATES",
                        dest = "bDuplicates", default = False,
                        action = "store_true",
                        help = "[optional] continue searching for duplicates after match [default: false]")
optionParser.add_option('-f', '--fields', metavar = "FIELDS", default = ("*"),
                        type = "string", dest = "sFields",
                        action = "callback", callback = optionsListCallback,
                        help = "[optional] comma delimited list of field(s) to extract data from [default: *]")
optionParser.add_option("-r", "--header-row", metavar = "HEADERROWSTART",
                        type = "int", dest = "lRowStart", default = 1,
                        help = "[optional] row number for specifying table location where multiple tables exist in a sheet [default: 1]")
optionParser.add_option("-c", "--header-column", metavar = "HEADERCOLUMNSTART",
                        type = "int", dest = "lColumnStart", default = 1,
                        help = "[optional] column number for specifying table location where multiple tables exist in a sheet [default: 1]")
optionParser.add_option("-l", "--delimiter", metavar = "DELIMITER",
                        type = "string", dest = "sDelimiter", default = u' | ',
                        help = "[optional] change the data output delimiter [default: ' | ']")
optionParser.add_option("-x", "--separation-row-count", metavar = "SEPARATIONROWCOUNT",
                        type = "int", dest = "lSeparationRowCount", default = 1,
                        help = "[optional] set the minimum number of concurrent empty rows for determining table extents [default: 1]")
optionParser.add_option("--header-to-stderr", metavar = "HEADERTOSTDERR",
                        dest = "bHeaderToStdErr", default = False,
                        action = "store_true",
                        help = "[optional] output first row to stderr [default: false]")
optionParser.add_option("-v", "--verbosity", metavar = "VERBOSITY",
                        type = "int", dest = "verbosity", default = 1,
                        action = "store",
                        help = "log level [default: 1]")

(options, args) = optionParser.parse_args() # options is a 'dict', args a 'list'


#process
try:

  if options.sDoc == "":
    raise Exception("[error] no document specified")

  if options.verbosity > 1:
    log("verbose mode: level " + str(options.verbosity))
    log("parsed args:")
    l = 0
    for name, value in options.__dict__.items():
      log( "idx: " + str(l) + ": '" + name + ": '" + str(value) + "'")
      l += 1

    log("unparsed args:")
    l = 0
    for value in args:
      log("idx: " + str(l) + ": '" + str(value) + "'")
      l += 1

  # load doc
  doc = odf.opendocument.load(options.sDoc)

  # get sheet
  sheet = None
  if options.sSheet:
    for ss in doc.spreadsheet.getElementsByType(Table):
      if ss.getAttribute("name") == options.sSheet:
        sheet = ss
        break
  else:
    sheet = doc.spreadsheet.getElementsByType(Table)[0]
  if not sheet:
    raise Exception("[error] cannot get sheet named: '" + sSheet + "'")

  # get table
  table = getTable(sheet, options.lRowStart, options.lColumnStart)
  # flatten table
  data = table2array(table)

  results = []

  # process table
  lKey = -1
  lFields = []
  aFields = []
  bHeader = False
  for value in options.sKeyValues:
    options.verbosity > 1 and log("searching for value: '" + value + "'")
    lRow = 0
    bReadRow = True
    while bReadRow:
      if lRow == 0 and not bHeader:
        if options.sKeyName != "":
          # find key, and setup dictionary for output fields
          for lCol in range(len(data[0])):
            if data[lRow][lCol] == options.sKeyName:
              lKey = lCol
              break
          if lKey == -1:
            raise Exception("key name: '" + options.sKeyName + "' not found in header")
        if options.sFields[0] == "*":
          lFields = range(len(data[0]))
          for lCol in lFields:
            aFields.append(data[lRow][lCol])
        else:
          aFields = options.sFields
          bSet = False
          for sField in aFields:
            for lCol in range(len(data[0])):
              if data[lRow][lCol] == sField:
                lFields.append(lCol)
                bSet = True
          if not bSet:
            lFields.append(-1)
        # add headers to results array
        results.append(aFields)
      else:
        # test key against key value
        if data[lRow][lKey] == value or value == "*":
          # success
          cols = []
          for l in lFields:
            cols.append(data[lRow][l])
          results.append(cols)
          if not options.bDuplicates and value != "*":
            bReadRow = False

      lRow += 1
      if lRow > len(data) - 1:
        bReadRow = False

  if len(results) > 1:
    # print results
    bHeader = False
    for row in results:
      text = ""
      for col in row:
        text = text + col.decode('utf-8') + options.sDelimiter.decode('utf-8') 
      text = text[:len(text) - len(options.sDelimiter)]
      if not bHeader:
        bHeader = True
        if options.bHeaderToStdErr:
          options.verbosity and log(text, 2)
        else:
          options.verbosity and log(text)
      else:
        options.verbosity and log(text)

#except Exception: pass
except Exception as e:
  print e

