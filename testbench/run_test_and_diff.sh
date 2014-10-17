#!/bin/bash

mkdir -p test/;
rm -rf test/*;
php test.php;
cp test.xlsx test/;
cd test;
unzip test.xlsx;
rm -f test.xlsx;
cd ..;

mkdir -p openoffice/;
rm -rf openoffice/*; 
cp openoffice.xlsx openoffice/; 
cd openoffice/; 
unzip openoffice.xlsx; 
rm -f openoffice.xlsx;
cd ..;
#exit;

for FILE in test/*.xml
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;

for FILE in test/*/*.xml
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;

for FILE in test/*/*/*.xml
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;

xmllint --format "test/xl/_rels/workbook.xml.rels" > temp.xml;
mv temp.xml test/xl/_rels/workbook.xml.rels;

for FILE in openoffice/*.xml
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;

for FILE in openoffice/*/*.xml
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;

for FILE in openoffice/*/*/*.xml
do
    xmllint --format "$FILE" > temp.xml;
    mv temp.xml $FILE;
done;

xmllint --format "openoffice/xl/_rels/workbook.xml.rels" > temp.xml;
mv temp.xml openoffice/xl/_rels/workbook.xml.rels;

echo "Now, run this command:"
echo "  meld openoffice/ test/";
#export DISPLAY=:0 && meld openoffice/ test/;

