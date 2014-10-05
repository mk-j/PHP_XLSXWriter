<?php
//php test_xlsxwriter.php >out.xlsx
//Generates a spreadsheet with multiple sheets, 10K rows, 10 columns
include_once(__DIR__.'/../xlsxwriter.class.php');

$headers = array('id'=>'string', 'name'=>'string', 'description'=>'string', 'n1'=>'string', 'n2'=>'string', 'n3'=>'string', 'n4'=>'string', 'n5'=>'string', 'n6'=>'string', 'n7'=>'string');
$sheet_names = array('january','february','march','april','may','june');
$start = microtime(true);

$writer = new XLSXWriter();
foreach($sheet_names as $sheet_name)
{
	$writer->writeSheetHeader($sheet_name, $headers);
	for($i=0; $i<10000; $i++)
	{
		$writer->writeSheetRow($sheet_name, random_row());
	}
}
$writer->writeToStdOut();

file_put_contents("php://stderr", '#'.floor(memory_get_peak_usage()/1024/1024)."MB"."\n");
file_put_contents("php://stderr", '#'.sprintf("%1.2f", microtime(true) - $start) ."s"."\n");


function random_row() {
	return $row = array(rand()%10000,
				chr(rand(97,122)).chr(rand(97,122)).chr(rand(97,122)).chr(rand(97,122)).chr(rand(97,122)).chr(rand(97,122)).chr(rand(97,122)),
				md5(uniqid()),
				rand()%10000,
				rand()%10000,
				rand()%10000,
				rand()%10000,
				rand()%10000,
				rand()%10000,
				rand()%10000,
		);
}



