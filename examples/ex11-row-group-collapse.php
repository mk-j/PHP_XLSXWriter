<?php
set_include_path( get_include_path().PATH_SEPARATOR."..");
include_once("xlsxwriter.class.php");

$chars = 'abcdefgh';

$writer = new XLSXWriter();

$column_header = array('C1' => 'string');
foreach (range(2, 11) as $col) {
    $column_header[] = array("C$col" => 'integer');
}
$writer->writeSheetHeader('Sheet1', $column_header);

$row_count = 0;
foreach (range("a", "j") as $letter) {
    $row_data = array($letter);
    foreach (range(1, 10) as $col_count) {
        $number = rand()%10000;
        $row_data[] = $number;
        @$total[$col_count] += $number;
    }

    /**
     * Set outlineLevel=1 for collapsed rows and every 5th row
     * set outlineLevel=0 and display totals
     */
    $row_options = array('collapse' => true, 'hidden' => true, 'outlineLevel' => 1);
    $writer->writeSheetRow('Sheet1', $row_data, $row_options);
    if (++$row_count % 5 === 0) {
        $row_options = array('collapse' => true, 'hidden' => false, 'outlineLevel' => 0);
        $writer->writeSheetRow('Sheet1', array_merge(array("Total after \"$letter\""), $total), $row_options);
    }

}

$writer->writeToFile('xlsx-group-row-collapse.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
