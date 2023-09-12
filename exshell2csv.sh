#!/bin/sh

####################### exshell2csv version 0.2 ########################
# exshell2csv: Small script to convert Excel to CSV, written in shell script only. No additional packages are required.
# Dependencies: Bourne Shell, sed, awk, and unzip.
# Usage: exshell2csv -h
# Reporitory: https://github.com/minamotorin/exshell2csv
# License: GNU General Public License Version 3 (https://www.gnu.org/licenses/gpl-3.0.html).
########################################################################

if [ "$1" = '-h' ]
then
  cat<<EOF
exshell2csv Version 0.2

Usage: `basename $0` [-h] [XLSX] [SHEET ID]

  -h               : Show this help
  [XLSX]           : Show list of SHEET IDS of XLSX file
  [XLSX] [SHEET ID]: Convert XLSX file’s SHEET ID to CSV (output to STDOUT)
EOF
exit
fi

if ! [ -e "$1" ]
then
  echo 'Error: file '"$1"' doesn’t exist' >&2
  exit 1
fi

if [ "$2" = '' ]
then
  python3 -c "import zipfile; print(zipfile.ZipFile('$1', 'r').extract('xl/workbook.xml').decode()) |
  sed '
    $!N
    H
    $!D
    x
    s/\n//g
    s/.*<sheets>//
    s/<\/sheets>.*//
    s/<sheet/\
&/g
  '                                                                    |
  sed 's/.* name="\(.*\)" sheetId="\([^"]*\)".*/\2: \1/; /^$/d'
  exit $?
fi

(
  python3 -c "import zipfile; print(zipfile.ZipFile('$1', 'r').extract('xl/sharedStrings.xml').decode()) |
  awk '{gsub("\\r", ""); print}'                                       |
  sed '
    1{
      s/^<?xml [^>]*>//
      /^$/d
      :loop
      /^[^<]/ s/.//
      t loop
    }
  '                                                                    |
  sed '
    1{/^<?xml/d;}
    s/<si>/\
&/g
    s/<si [^>]*>/\
<si>/g
  '                                                                    |
  sed '
    1d
    :loop
    /<\/si>/!{
      N
      bloop
    }

    :si
    /^<si>/! s/.//
    t si
    s/^<si>//

    s/\\/&&/g
    s/\n/\\n/g
    s/<\/si>.*//
  '                                                                    |
  sed '
    :topen
    /^<t>/!{ /^<t /!{
      s/^<[^>]*>//
      t topen
    }; }
    s/^<t>//
    s/^<t [^>]*>//

    H
    x
    s/<\/t>.*//
    s/\n//
    x

    :tclose
    /^<\/t>/!{
      s/.//
      t tclose
    }
    s/^<\/t>//

    /<\/t>/ b topen
    s/.*//
    x
    '                                                                  | tr -cd "[:print:]\n" |
    sed 's/^/l /'

  echo

  python3 -c "import zipfile; print(zipfile.ZipFile('$1', 'r').extract('xl/worksheets/sheet"$2".xml').decode()) |
  python3 parseSheet.py                                                |
  sed -E 's/^([A-Z]{1,3})([0-9]{1,7})/\1 \2/'
)                                                                      |
cat
