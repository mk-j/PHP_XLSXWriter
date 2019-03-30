<?php
/**
 * Created by PhpStorm.
 * User: elminsondeoleobaez
 * Date: 10/3/18
 * Time: 5:52 PM
 */

//@todo json file? with data for composer and classes
//@todo form data,
//@todo Developer_name, Project_name, phpunit? (checkbox),
//@todo Create folders (src, tests)
//@todo Create MainClass => src (add autoload.php)
//@todo Create composer.json file
//@todo Create Test Cases class
//@todo Create Readme.md file
//@todo Zip content

namespace Mkj\XLSXWriter;

require __DIR__ . '/vendor/autoload.php';

$writer = new XLSXWriter();
$writer->setAuthor('Some Author');


$filename = "example.xlsx";
header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header('Content-Transfer-Encoding: binary');
header('Cache-Control: must-revalidate');
header('Pragma: public');

$rows = array(
    array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    array('2003','=B1', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'),
);


foreach($rows as $row)
    $writer->writeSheetRow('Sheet1', $row);
//$writer->writeToStdOut();
$writer->writeToFile('example.xlsx');
