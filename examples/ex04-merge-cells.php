<?php
set_include_path( get_include_path().PATH_SEPARATOR."..");
include_once("xlsxwriter.class.php");

$header = array("string","string","string","string","string");
$row1 = array("Merge Cells Example");
$row2 = array(100, 200, 300, 400, 500);
$row3 = array(110, 210, 310, 410, 510);

$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', $header, $suppress_header_row = true);
$writer->writeSheetRow('Sheet1', $row1);
$writer->writeSheetRow('Sheet1', $row2);
$writer->writeSheetRow('Sheet1', $row3);
$writer->markMergedCell('Sheet1', $start_row=0, $start_col=0, $end_row=0, $end_col=4);
$writer->writeToFile('xlsx-merge-cells.xlsx');


