<?php
include_once("../xlsxwriter.class.php");

$header = array(
	'year'=>'string',
	'month'=>'string',
	'amount'=>'money',
);
$data = array(
	array('2003','1','-50.5'),
	array('2003','5', '23.5'),
);


$writer = new XLSXWriter();
$writer->setAuthor('Doc Author');
$writer->writeSheet($data,'Sheet1',$header);
$writer->writeToFile('test.xlsx');





