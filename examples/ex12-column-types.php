<?php
set_include_path(get_include_path() . PATH_SEPARATOR . '..');
include_once("xlsxwriter.class.php");

$data = [
    ['Roman', '', '', '42'],
    ['Jackie', '', '', '43'],
    ['Steve', '', '', '17'],
];

$writer = new XLSXWriter();
$writer->setSheetColumnTypes('Sheet1', ['string', '@', '@', '@']);
$writer->writeSheetRow('Sheet1', ['Name', 'Not in use', 'Not in use', 'Value']);

foreach($data as $row)
{
    $writer->writeSheetRow('Sheet1', $row);
}

$writer->writeToFile('column-types.xlsx');

