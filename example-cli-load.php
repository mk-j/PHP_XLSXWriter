<?php
include_once("xlsxwriter.class.php");

$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', array('c1'=>'string','c2'=>'string','c3'=>'string','c4'=>'string') );//optional
for($i=0; $i<50000; $i++)
{
    $writer->writeSheetRow('Sheet1', array(rand()%10000,rand()%10000,rand()%10000,rand()%10000) );
}
$writer->writeToFile('output.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
