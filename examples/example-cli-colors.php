<?php
include_once("xlsxwriter.class.php");

file_put_contents("php://stderr", "Writing to example-cli-colors.xlsx\n");

$writer = new XLSXWriter();
$colors = array('ff','cc','99','66','33','00');
foreach($colors as $b) {
	foreach($colors as $g) {
		$rowdata = array();
		$rowstyle = array();
		foreach($colors as $r) {
			$rowdata[] = "#$r$g$b";
			$rowstyle[] = array('fill'=>"#$r$g$b");
		}
		$writer->writeSheetRow('Sheet1', $rowdata, $rowstyle );
	}
}
$writer->writeToFile('example-cli-colors.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";


