#!/bin/bash
FILENAME=`basename --suffix=.xlsx $1`
DIRNAME=$FILENAME"_dir"
if [ "" == "$1" ]; then
    echo "Example Usage: $0 spreadsheet.xlsx";
    exit
fi
if [ ! -f "$1" ]; then
    echo "Example Usage: $0 spreadsheet.xlsx";
    exit
fi
mkdir -p $DIRNAME;
cp $1 $DIRNAME;
cd $DIRNAME;
unzip -o $1;

for FILE in *.xml; 
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;
for FILE in */*.xml
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;
for FILE in */*/*.xml
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;
xmllint --format "xl/_rels/workbook.xml.rels" > temp.xml;
mv temp.xml xl/_rels/workbook.xml.rels;
cd ..;
exit

