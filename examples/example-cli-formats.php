<?php
include_once("xlsxwriter.class.php");


$header = array(
  'c1'=>'dollar',
  'c2'=>'euro',
  'c3'=>'#,##0.00', //custom
  'c4'=>'#,##0.00 [$€-407]', //german euro
  'c5'=>'[$￥-411]#,##0;[RED]-[$￥-411]#,##0', //japanese yen
);
$row = array(100,200,300,400,500);
$writer = new XLSXWriter();
$writer->writeSheet(array($row),'Sheet1', $header);
$writer->writeToFile('example.xlsx');


