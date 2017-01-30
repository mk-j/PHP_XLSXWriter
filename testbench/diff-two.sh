#!/bin/bash
if [ ! -f "$1" ] || [ ! -f "$2" ]; then
    echo "Example Usage: $0 f1.xlsx f2.xlsx";
    exit
fi
./extract.sh $1
./extract.sh $2
echo "Now, run this command:"
echo "  meld openoffice/ test/";
#export DISPLAY=:0 && meld openoffice/ test/;

