
# odspyread

## description

python script to extract data from an ods format spreadsheet. handles multiple tables per sheet. optionally search for a value in a specified index/field and dump a selection/all fields' data for the records found

## dependencies
-odfpy

## help

Usage: odspyread.py [options]

Options:
  -h, --help            show this help message and exit
  -d DOC, --document=DOC
                        the spreadsheet path
  -e SHEET, --sheet=SHEET
                        [optional] sheet name of interest [default: first
                        sheet]
  -n NAME, --name=NAME  [optional] locate table position in sheet by name
  -r HEADERROWSTART, --header-row=HEADERROWSTART
                        [optional] locate table position in sheet by row
                        number [default: 1]
  -c HEADERCOLUMNSTART, --header-column=HEADERCOLUMNSTART
                        [optional] locate table position in sheet by column
                        number [default: 1]
  -i IDX, --idx=IDX     [optional] name of the index/key field for searching
                        in
  -s SEARCH, --search=SEARCH
                        [optional] comma-delimited list of value(s) to search
                        for under the index/key field [default: '*']
  -x REGEXPSEARCH, --regexp-search=REGEXPSEARCH
                        [optional] comma-delimited list of regular
                        expression(s) to search for under the index/key field
                        [default: '.*']
  -m, --allow-duplicates
                        [optional] continue searching for duplicates after
                        match [default: false]
  -f FIELDS, --fields=FIELDS
                        [optional] comma delimited list of field(s) to extract
                        data from [default: '*']
  --row-filter=ROWFILTER
                        [optional] comma delimited list of terms to be
                        filtered out results
  --comment-filter=COMMENTFILTER
                        [optional] comma delimited list of prefixes to ignore
                        when determining table position [default: '#']
  --delimiter=DELIMITER
                        [optional] change the data output delimiter [default:
                        ' | ']
  --header-to-stderr    [optional] output first row to stderr [default: false]
  --max-empty-rows=MAXEMPTYROWS
                        [optional] set the maximum number of concurrent empty
                        rows for determining table extents [default: 1]
  -v VERBOSITY, --verbosity=VERBOSITY
                        log level [default: 1]

## examples

the following examples are based on the content of 'test.odf':

*search for 'search 2' record under field 'idx'
> ./odspyread.py -d test.ods -i "idx" -s "search 2"

header 1 | header 2 | idx | header 4 | header 5
 | data1 2 2 | search 2 |  |

*dump a table which exists at some offset in the sheet
>./odspyread.py -d test.ods -r 10 -c 5

header 1 | idx | header 2 | header 4
 | search 1 |  |
 | search 2 | data3 2 2 |
 | search 3 |  | data3 3 4

*search for 'search2' in sheet 'test' in the first table found beyond row 9, returning data for fields 'idx' and 'header 3' only, delimited by ','
>./odspyread.py -d test.ods -r 9 -i 'idx' -s 'search 2' -f 'idx,header 3' --delimiter ','

idx,header 3
search 2,data2 2 3

*without comment filters the table is identifies as a 4x1, the comment viewed a header
>./odspyread.py -d test.ods -r 25 --max-empty-rows=2 --comment-filter=""
# table 5
header 1

data5 2 1

*filtering the '# table 5' comment we recover our intended 7x4 table. note here that the 2 rows below 'data 5 2 1' are now not seen as empty, there is data in columns 2 and 3 respectively, hence the table extends beyond that identified in the previous example
>./odspyread.py -d test.ods -r 25 --max-empty-rows=2 --comment-filter="#"
header 1 | header 2 | header 3 | header 4
 |  | data5 1 3 | data5 1 4
data5 2 1 |  |  |
 | data5 3 2 | data5 3 3 | data5 3 4
 |  | data5 4 3 | data5 4 4
 |  |  |
data5 6 1 | data5 6 2 | data5 6 3 | data5 6 4

*no row filter set (default)
>./odspyread.py -d test.ods -r 25 -c 8 --row-filter=''
header 1 | header 2 | header 3
data5 2 1 |  | data5 1 3
!! comment row |  |
data5 2 2 | data5 3 2 | data5 3 3

*rows are filtered if any entry in them are prefixed with a filter term
>./odspyread.py -d test.ods -r 25 -c 8 --row-filter='!'
header 1 | header 2 | header 3
data5 2 1 |  | data5 1 3
data5 2 2 | data5 3 2 | data5 3 3

*protect search terms containing commas, and always escape output which contains the delimiter string. without the escape in the search term here, terms 'atom1' and 'atom2' would be searched for independently
>./odspyread.py -d test.ods -r 36 -i groups -s "atom1\,atom2" -m --delimiter=","
groups,times
atom1\,atom2,3\,2\,3

*search for table by a title which must lie in a cells directly above any part of the table header. the '-n | --name' option can be supplemented with '-r | --header-row' and '-c | --header-column' options
>./odspyread.py -d test.ods -n "table 4" -c 2 -i idx -s "search2" --verbosity 2
[debug] found title: 'table 4' at: 'r19c3'
idx | header 2
search2 | data4b 2 2
>./odspyread.py -d test.ods -n "table 4c" -i idx -s "search2" --verbosity 2
[debug] found title: 'table 4c' at: 'r19c7'
idx | header 2 | header 3
search2 | data4c 2 2 | data4c 2 3

*regexp search for all 'atom1' related results
>./odspyread.py -d test.ods -n "table 7" -i groups -m -x "atom1.*atom3"
groups | times
atom1,atom3 | 5,4,4
atom1,atom2,atom3 | 1,1,1

## todo
-README markdown
-remove duplicate result rows when given multiple search values
-handle non-existent tables and fields
-create class
-deprecate optparse
