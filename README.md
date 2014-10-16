PHP_XLSXWriter
==============

This library is designed to be lightweight, and have relatively low memory usage.

It is designed to output an Excel spreadsheet in with (Office 2007+) xlsx format, with just basic features supported:
* supports PHP 5.2.1+
* takes UTF-8 encoded input
* multiple worksheets
* supports a few simple cell formats:
  * simple $0.00 currency format 
  * simple ``Y-M-D/Y-M-D H:m:s`` format

Give this library a try, if you find yourself [running out of memory writing spreadsheets with PHPExcel](http://www.zedwood.com/article/php_xlsxwriter-performance-comparison).

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

Multiple Sheets:
```php
$data1 = array(  
     array('5','3'),
     array('1','6'),
);
$data2 = array(  
     array('2','7','9'),
     array('4','8','0'),
);

$writer = new XLSXWriter();
$writer->setAuthor('Doc Author');
$writer->writeSheet($data1);
$writer->writeSheet($data2);
echo $writer->writeToString();
```

Cell Formatting:
```php
$header = array(
  'create_date'=>'date',
  'quantity'=>'string',
  'product_id'=>'string',
  'amount'=>'money',
  'description'=>'string',
);
$data = array(
    array('2013-01-01',1,27,'44.00','twig'),
    array('2013-01-05',1,'=C1','-44.00','refund'),
);

$writer = new XLSXWriter();
$writer->writeSheet($data,'Sheet1', $header);
$writer->writeToFile('example.xlsx');
```

Load test with 50000 rows: (runs fast, with low memory usage)
```php
include_once("xlsxwriter.class.php");
$header = array('c1'=>'string','c2'=>'string','c3'=>'string','c4'=>'string');
$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', $header );//optional
for($i=0; $i<50000; $i++)
{
    $writer->writeSheetRow('Sheet1', array(rand()%10000,rand()%10000,rand()%10000,rand()%10000) );
}
$writer->writeToFile('output.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
```
