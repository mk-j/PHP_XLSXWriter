<?php
include_once("xlsxwriter.class.php");

$rows = array(
    array('2003','=B2', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'),
);
$writer = new XLSXWriter();
$writer->setAuthor('Some Author');
foreach($rows as $row)
	$writer->writeSheetRow('Sheet1', $row);
$writer->writeToFile('example.xlsx');
//$writer->writeToStdOut();
//echo $writer->writeToString();
exit(0);


