<?php
include_once("xlsxwriter.class.php");

$header = array("string","string","string","string","string");
$row1 = array("Merge Cells Example");
$row2 = array(100, 200, 300, 400, 500);
$row3 = array(110, 210, 310, 410, 510);

$sheet_name = 'Sheet1';
$writer = new XLSXWriter();
$writer->writeSheetHeader($sheet_name, $header, $suppress_header_row = true);
$writer->writeSheetRow($sheet_name, $row1);
$writer->writeSheetRow($sheet_name, $row2);
$writer->writeSheetRow($sheet_name, $row3);
$writer->markMergedCell($sheet_name, $start_row=0, $start_col=0, $end_row=0, $end_col=4);
$writer->writeToFile('example.xlsx');


