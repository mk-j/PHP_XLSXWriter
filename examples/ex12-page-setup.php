<?php
set_include_path( get_include_path().PATH_SEPARATOR."..");
include_once("xlsxwriter.class.php");

$header = array(
  'c1-text'=>'string',//text
  'c2-text'=>'@',//text
  'c3-integer'=>'integer',
  'c4-integer'=>'0',
  'c5-price'=>'price',
  'c6-price'=>'#,##0.00',//custom
  'c7-date'=>'date',
  'c8-date'=>'YYYY-MM-DD',
);
$rows = array(
  array('x101',102,103,104,105,106,'2018-01-07','2018-01-08'),
  array('x201',202,203,204,205,206,'2018-02-07','2018-02-08'),
  array('x301',302,303,304,305,306,'2018-03-07','2018-03-08'),
  array('x401',402,403,404,405,406,'2018-04-07','2018-04-08'),
  array('x501',502,503,504,505,506,'2018-05-07','2018-05-08'),
  array('x601',602,603,604,605,606,'2018-06-07','2018-06-08'),
  array('x701',702,703,704,705,706,'2018-07-07','2018-07-08'),
);
$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', $header, [
	'freeze_rows' => 1,
	'freeze_columns' => 2,
	// Use the "page_setup' to set up layout and print options of a page. Refer to the "page setup" dialog of Excel.
	'page_setup' => [
		'orientation' => 'landscape', // choose 'landscape' or 'portrait'
		'scale' => 80,                // percent
		//'fit_to_width' => 0,        // When "Fit to page", specify the number of pages
		//'fit_to_height' => 0,       // When "Fit to page", specify the number of pages
		'paper_size' => 9,            // 9=xlPaperA4 : specify XlPaperSize value. see https://msdn.microsoft.com/vba/excel-vba/articles/xlpapersize-enumeration-excel
		//'horizontal_dpi' => 600,
		//'vertical_dpi' => 600,
		//'first_page_number' => 1,
		//'use_first_page_number' => false,

		// Specify margin in inches (not in centimeters)
		'margin_left' => 0.1,
		'margin_right' => 0.2,
		'margin_top' => 0.3,
		'margin_bottom' => 0.4,
		'margin_header' => 0.5,
		'margin_footer' => 0.6,
		'horizontal_centered' => true,
		'vertical_centered' => true,

		// Header\Footer can be customized.
		// for details on how to write, refer to 'Remarks' of https://msdn.microsoft.com/library/documentformat.openxml.spreadsheet.evenheader.aspx
		'header' => 'Page Title',      // ex. fixed-text
		'footer' => '&amp;P / &amp;N', // ex. page number

		'print_area' => 'A1:F5',
		'print_titles' => '$1:$1',
		// Note : When setting multiple ranges, specify by array. ex. 'print_titles' => ['$1:$1', '$A:$A'],

		'page_order' => 'downThenOver', // choose 'overThenDown' or 'downThenOver'
		//'grid_lines' => false,
		//'black_and_white' => false,
		//'draft' => false,
		//'headings' => false,
		//'cell_comments' => 'none',    // 'asDisplayed', 'atEnd', 'none' can be selected
		//'errors' => 'displayed',      // 'blank', 'dash', 'displayed', 'NA' can be selected

		//'use_printer_defaults' => true,
		//'copies' => 1,
	]
]);
foreach ($rows as $row) {
	$writer->writeSheetRow('Sheet1', $row);
}
$writer->writeSheetHeader('Sheet2', $header, [
	'page_setup' => [
		'orientation' => 'portrait', // choose 'landscape' or 'portrait'
		'fit_to_width' => 2,        // When "Fit to page", specify the number of pages
		'fit_to_height' => 3,       // When "Fit to page", specify the number of pages
		'first_page_number' => 4,
		'use_first_page_number' => true,

		'header' => '', // no header
		'footer' => '', // no footer

		'print_area' => ['A1:B2', 'C3:D4'],   // multiple print area specification
		'print_titles' => ['$1:$1', '$A:$A'], // row print title & column print title specification
		'page_order' => 'overThenDown', // choose 'overThenDown' or 'downThenOver'

		'grid_lines' => true,
		'black_and_white' => true,
		//'draft' => true,
		'headings' => true,
	]
]);
foreach ($rows as $row) {
	$writer->writeSheetRow('Sheet2', $row);
}
$writer->writeToFile('xlsx-print-setup.xlsx');
