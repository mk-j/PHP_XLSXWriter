<?php
require_once(__DIR__ . '/../vendor/autoload.php');

use pti\XLSXWriter\XLSXWriter;

$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', array('c1'=>'integer','c2'=>'integer','c3'=>'integer','c4'=>'integer') );//optional
for($i=0; $i<250000; $i++)
{
    $writer->writeSheetRow('Sheet1', array(rand()%10000,rand()%10000,rand()%10000,rand()%10000) );
}
$writer->writeToFile('xlsx-numbers-250k.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
