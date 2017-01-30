PHP_XLSXWriter
==============

This library is designed to be lightweight, and have relatively low memory usage.

It is designed to output an Excel spreadsheet in with (Office 2007+) xlsx format, with just basic features supported:
* supports PHP 5.2.1+
* takes UTF-8 encoded input
* multiple worksheets
* supports currency/date/numeric cell formatting, simple formulas

Give this library a try, if you find yourself [running out of memory writing spreadsheets with PHPExcel](http://www.zedwood.com/article/php-excel-writer-performance-comparison).

Simple PHP CLI example:
```php
$data = array(
    array('year','month','amount'),
    array('2003','1','220'),
    array('2003','2','153.5'),
);

$writer = new XLSXWriter();
$writer->writeSheet($data);
$writer->writeToFile('output.xlsx');
```

Simple/Advanced Cell Formats:
```php
//simple formats: date, datetime, integer, dollar, euro, string
$header = array(
  'created'=>'date',
  'product_id'=>'integer',
  'quantity'=>'#,##0',
  'amount'=>'price',
  'description'=>'string',
  'tax'=>'[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
);
$data = array(
    array('2015-01-01',873,1,'44.00','misc','=D2*0.05'),
    array('2015-01-12',324,2,'88.00','none','=D3*0.05'),
);

$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', $header );
foreach($data as $row)
	$writer->writeSheet('Sheet1', $row, );
$writer->writeToFile('example.xlsx');
```

50000 rows: (1.4s, 0MB memory usage)
```php
include_once("xlsxwriter.class.php");
$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', array('c1'=>'integer','c2'=>'integer','c3'=>'integer','c4'=>'integer') );
for($i=0; $i<50000; $i++)
{
    $writer->writeSheetRow('Sheet1', array($i, $i+1, $i+2, $i+3, $i+4) );
}
$writer->writeToFile('output.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
```
Memory usage will not increase due to rowcount

| rows   | time | memory |
| ------ | ---- | ------ |
|  50000 | 1.4s | 0MB    |
| 100000 | 2.7s | 0MB    |
| 150000 | 4.1s | 0MB    |
| 200000 | 5.7s | 0MB    |
| 250000 | 7.0s | 0MB    |

Simple cell formats map to more advanced cell formats
| simple formats | format code |
| string   | @ |
| integer  | 0 |
| date     | YYYY-MM-DD |
| datetime | YYYY-MM-DD HH:MM:SS |
| price    | #,##0.00 |
| dollar   | [$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00 |
| euro     | '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407] |


Basic cell styles have been available since version 0.30

| style      | allowed values |
| ---------- | ----  |
| font       | '', 'Arial', 'Times New Roman', 'Courier New', 'Comic Sans MS' |
| font-style | '', 'bold','italic','underline','strikethrough' or a combination ie: 'bold,italic' |
| font-size  | 8,9,10,11,12 |
| border     | '', 'left','right','top','bottom' or a combination ie: 'top,left' |
| color      | #RRGGBB, ie: #ff99cc or #f9c |
| fill       | #RRGGBB, ie: #eeffee or #efe |
| halign     | 'general','left','right','justify','center' |
| valign     | 'bottom','center','distributed' |


