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

set_include_path( __DIR__."/..");
include_once('phpexcel/PHPExcel.php');
include_once('phpexcel/PHPExcel/Writer/Excel2007.php');

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Doc Author");

for($i=1,$ix=count($sheets_array); $i<$ix; $i++)
{
	$objPHPExcel->createSheet($i);
}

$sheet_number=0;
foreach($sheets_array as $sheet_name=>$sheet_info)
{
	$theheaders = $sheet_info['headers'];
	$thedata    = $sheet_info['data'];
	$header_lookup = array_values($theheaders);

	$objPHPExcel->setActiveSheetIndex($sheet_number++);
	$objPHPExcel->getActiveSheet()->setTitle($sheet_name);
	
	$column_number=0;
	$row_number=0;
	foreach($theheaders as $field_name=>$field_type)
	{
		$cell = xlsCell($row_number, $column_number++);
		$objPHPExcel->getActiveSheet()->SetCellValue($cell, $field_name);
	}
	foreach($thedata as $row)
	{
		$row_number++;
		foreach($row as $column_number=>$cell_value)
		{
			//TODO handle numeric and date formatting
			//if ($header_lookup[$col]=='date')  { $format= $format_date; }
			//if ($header_lookup[$col]=='money') { $format= $format_num; }
			//if ($header_lookup[$col]=='dollar') { $format= $format_dollar; }
			//$objPHPExcel->getActiveSheet()->getStyle('E1')->getNumberFormat()->setFormatCode("##0;(-##0)");
			//$objPHPExcel->getActiveSheet()->getStyle('F1')->getNumberFormat()->setFormatCode('#,##0.00;[Red](#,##0.00)');

			$cell = xlsCell($row_number, $column_number);
			$objPHPExcel->getActiveSheet()->SetCellValue($cell, $cell_value);
		}
	}
}
$objPHPExcel->setActiveSheetIndex(0);

$tmp_filename = tempnam("/tmp", "phpexcel_").".xlsx";
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save($tmp_filename);
//readfile($tmp_filename);

function xlsCell($row_number, $column_number)
{
	$n = $column_number;
	for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
	{
		$r = chr($n%26 + 0x41) . $r;
	}
	return $r . ($row_number+1);
}
file_put_contents("php://stderr", '#'.floor(memory_get_peak_usage()/1024/1024)."MB"."\n");
file_put_contents("php://stderr", '#'.sprintf("%1.2f", microtime(true) - $start) ."\n");








