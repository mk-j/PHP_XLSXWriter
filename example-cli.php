<?php
include_once("xlsxwriter.class.php");
ini_set('display_errors', 0);
ini_set('log_errors', 1);
error_reporting(E_ALL & ~E_NOTICE);

$filename = "example.xlsx";

$rows = array(
    array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    array('2003','=B1', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'),
);

$keywords = array('some','interesting','keywords');

$writer = new XLSXWriter();
$writer->setTitle('Some Title');
$writer->setSubject('Some Subject');
$writer->setAuthor('Some Author');
$writer->setCompany('Some Company');
$writer->setKeywords($keywords);
$writer->setDescription('Some interesting description');
foreach($rows as $row)
	$writer->writeSheetRow('Sheet1', $row);

$writer->writeToFile('example.xlsx');
