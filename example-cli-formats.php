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

//cell styling
//create a "fill" first: pattern, forground color, background color
$fill = $this->addFill('solid', 'FF9900', 'FFFFFF');
//create a style based on our fill
$style = $this->addCellStyle(array('fillId' => $fill));

//cell in column B will have the custom style
$row_style = array(0, $style, 0, 0, 0);
$this->writeSheetRow('sheet1', $row, $row_style);

$writer->writeToFile('example.xlsx');




