<?php
//UGLY code, but it generates a spreadsheet with multiple sheets, 10K rows, 10 columns
function rows()
{
	$rows = array();
	for($i=0; $i<10000; $i++)
	{
		$rows[] = array(rand()%10000,
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
	return $rows;
}
$headers = array('id'=>'string', 'name'=>'string', 'description'=>'string', 'n1'=>'string', 'n2'=>'string', 'n3'=>'string', 'n4'=>'string', 'n5'=>'string', 'n6'=>'string', 'n7'=>'string');
$sheets_array = array(
	'january'=>array(
		'headers'=>$headers,
		'data'=>rows(),
	),
	'february'=>array(
		'headers'=>$headers,
		'data'=>rows(),
	),
	'march'=>array(
		'headers'=>$headers,
		'data'=>rows(),
	),
	'april'=>array(
		'headers'=>$headers,
		'data'=>rows(),
	),
	'may'=>array(
		'headers'=>$headers,
		'data'=>rows(),
	),
	'june'=>array(
		'headers'=>$headers,
		'data'=>rows(),
	),
);
file_put_contents("php://stderr", '#'.floor(memory_get_peak_usage()/1024/1024)."MB"."\n");
$start = microtime(true);

include_once('PHP_XLSXWriter/xlsxwriter.class.php');
$data = array(
    array('year','month','amount'),
    array('2003','1','220'),
    array('2003','2','153.5'),
);

$writer = new XLSXWriter();
foreach($sheets_array as $sheet)
{
	$writer->writeSheet($sheet['data'],'',$sheet['headers']);
}
$writer->writeToStdOut();

file_put_contents("php://stderr", '#'.floor(memory_get_peak_usage()/1024/1024)."MB"."\n");
file_put_contents("php://stderr", '#'.sprintf("%1.2f", microtime(true) - $start) ."\n");








