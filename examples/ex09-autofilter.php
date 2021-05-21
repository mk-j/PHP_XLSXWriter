<?php
set_include_path( get_include_path().PATH_SEPARATOR."..");
include_once("xlsxwriter.class.php");

$chars = 'abcdefgh';

$writer = new XLSXWriter();

/**
 * auto_filter may also be set to numbers greater than 1 to offset the
 * start row for the auto filter data range. This will have the effect
 * of changing the start row for the data to be filtered and there-by
 * the row the dropdown filters show on to the row number specified
 * in this option.
 */
$writer->writeSheetHeader(
    'Sheet1',
    ['col-string'=>'string', 'col-numbers'=>'integer', 'col-timestamps'=>'datetime'],
    ['auto_filter'=>1, 'widths'=>[15, 15, 30]]
);
for($i=0; $i<1000; $i++)
{
    $writer->writeSheetRow('Sheet1', array(
        str_shuffle($chars),
        rand()%10000,
        date('Y-m-d H:i:s',time()-(rand()%31536000))
    ));
}
$writer->writeToFile('xlsx-autofilter.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
