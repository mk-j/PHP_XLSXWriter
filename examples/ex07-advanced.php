<?php
require_once(__DIR__ . '/../vendor/autoload.php');

use pti\XLSXWriter\XLSXWriter;

$writer = new XLSXWriter();
$writer->setAuthor('Some Author');
$writer->setTempDir(sys_get_temp_dir());//set custom tempdir

//----
$sheet1 = 'merged_cells';
$header = array("string","string","string","string","string");
$rows = array(
    array("Merge Cells Example"),
    array(100, 200, 300, 400, 500),
    array(110, 210, 310, 410, 510),
);
$writer->writeSheetHeader($sheet1, $header, $suppress_header_row = true);
foreach($rows as $row)
	$writer->writeSheetRow($sheet1, $row);
$writer->markMergedCell($sheet1, $start_row=0, $start_col=0, $end_row=0, $end_col=4);

//----
$sheet2 = 'utf8_examples';
$rows = array(
    array('Spreadsheet','_'),
	array("Hoja de cálculo", "Hoja de c\xc3\xa1lculo"),
    array("Електронна таблица", "\xd0\x95\xd0\xbb\xd0\xb5\xd0\xba\xd1\x82\xd1\x80\xd0\xbe\xd0\xbd\xd0\xbd\xd0\xb0 \xd1\x82\xd0\xb0\xd0\xb1\xd0\xbb\xd0\xb8\xd1\x86\xd0\xb0"),//utf8 encoded
    array("電子試算表", "\xe9\x9b\xbb\xe5\xad\x90\xe8\xa9\xa6\xe7\xae\x97\xe8\xa1\xa8"),//utf8 encoded
);
$writer->writeSheet($rows, $sheet2);

//----
$sheet3 = 'font_example';
$format = array('font'=>'Arial','font-size'=>10,'font-style'=>'bold,italic', 'fill'=>'#eee','color'=>'#f00','fill'=>'#ffc', 'border'=>'top,bottom', 'halign'=>'center');

$rows = array(
    array(101,102,103,104,105,106,107,108,109,110),
    array(201,202,203,204,205,206,207,208,209,210),
);
foreach($rows as $row)
	$writer->writeSheetRow($sheet3, $row, $format);
$writer->writeToFile('xlsx-advanced.xlsx');


