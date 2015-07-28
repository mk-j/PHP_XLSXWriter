<?php
include_once("xlsxwriter.class.php");

$header = array();
$data = array(
    array("Merge Cells Example"),
    array(100, 200, 300, 400, 500)
);
$merge_cells = array(
    array(
        array(0, 0),
        array(0, 4)
    )
);
$writer = new XLSXWriter();
$writer->writeSheet($data, 'Sheet1', $header, $merge_cells);
$writer->writeToFile('example.xlsx');


