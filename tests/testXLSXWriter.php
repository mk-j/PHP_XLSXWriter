<?php
/**
 * Created by PhpStorm.
 * User: elminsondeoleobaez
 * Date: 10/3/18
 * Time: 1:52 PM
 */

namespace Mkj\XLSXWriter;

require __DIR__ . '/../vendor/autoload.php';

use PHPUnit\Framework\TestCase;

class testPHPXLSXWriter extends TestCase
{

    protected function setUp()
    {
        parent::setUp();
    }

    /**
     *
     */
    function testFirstTestCase()
    {
        $writer = new XLSXWriter();
        $writer->setAuthor('Some Author');


        $filename = "example.xlsx";

        $rows = array(
            array('2003', '1', '-50.5', '2010-01-01 23:00:00', '2012-12-31 23:00:00'),
            array('2003', '=B1', '23.5', '2010-01-01 00:00:00', '2012-12-31 00:00:00'),
        );


        foreach ($rows as $row)
            $writer->writeSheetRow('Sheet1', $row);
        $writer->writeToFile($filename);
        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }

    function testSimpleXLSX()
    {
        $filename = "xlsx-simple.xlsx";
        $writer = new XLSXWriter();
        $header = array(
            'c1-text' => 'string',//text
            'c2-text' => '@',//text
            'c3-integer' => 'integer',
            'c4-integer' => '0',
            'c5-price' => 'price',
            'c6-price' => '#,##0.00',//custom
            'c7-date' => 'date',
            'c8-date' => 'YYYY-MM-DD',
        );
        $rows = array(
            array('x101', 102, 103, 104, 105, 106, '2018-01-07', '2018-01-08'),
            array('x201', 202, 203, 204, 205, 206, '2018-02-07', '2018-02-08'),
            array('x301', 302, 303, 304, 305, 306, '2018-03-07', '2018-03-08'),
            array('x401', 402, 403, 404, 405, 406, '2018-04-07', '2018-04-08'),
            array('x501', 502, 503, 504, 505, 506, '2018-05-07', '2018-05-08'),
            array('x601', 602, 603, 604, 605, 606, '2018-06-07', '2018-06-08'),
            array('x701', 702, 703, 704, 705, 706, '2018-07-07', '2018-07-08'),
        );

        $writer->writeSheetHeader('Sheet1', $header);
        foreach ($rows as $row)
            $writer->writeSheetRow('Sheet1', $row);

        $writer->writeToFile($filename);
        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }

    function testMultipleXLSX()
    {

        $filename = "xlsx-sheets.xlsx";
        $writer = new XLSXWriter();


        $header = array(
            'year' => 'string',
            'month' => 'string',
            'amount' => 'price',
            'first_event' => 'datetime',
            'second_event' => 'date',
        );
        $data1 = array(
            array('2003', '1', '-50.5', '2010-01-01 23:00:00', '2012-12-31 23:00:00'),
            array('2003', '=B2', '23.5', '2010-01-01 00:00:00', '2012-12-31 00:00:00'),
            array('2003', "'=B2", '23.5', '2010-01-01 00:00:00', '2012-12-31 00:00:00'),
        );
        $data2 = array(
            array('2003', '01', '343.12', '4000000000'),
            array('2003', '02', '345.12', '2000000000'),
        );

        $writer->writeSheetHeader('Sheet1', $header);
        foreach ($data1 as $row)
            $writer->writeSheetRow('Sheet1', $row);
        foreach ($data2 as $row)
            $writer->writeSheetRow('Sheet2', $row);

        $writer->writeToFile($filename);
        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);

        unlink($filename);
    }


    function testFormats()
    {

        $filename = "xlsx-formats.xlsx";
        $writer = new XLSXWriter();

        $sheet1header = array(
            'c1-string' => 'string',
            'c2-integer' => 'integer',
            'c3-custom-integer' => '0',
            'c4-custom-1decimal' => '0.0',
            'c5-custom-2decimal' => '0.00',
            'c6-custom-percent' => '0%',
            'c7-custom-percent1' => '0.0%',
            'c8-custom-percent2' => '0.00%',
            'c9-custom-text' => '@',//text
        );
        $sheet2header = array(
            'col1-date' => 'date',
            'col2-datetime' => 'datetime',
            'custom-date1' => 'YYYY-MM-DD',
            'custom-date2' => 'MM/DD/YYYY',
            'custom-date3' => 'DD-MMM-YYYY HH:MM AM/PM',
            'custom-date4' => 'MM/DD/YYYY HH:MM:SS',
            'custom-date5' => 'YYYY-MM-DD HH:MM:SS',
            'custom-date6' => 'YY MMMM',
            'custom-date7' => 'QQ YYYY',
        );
        $sheet3header = array(
            'col1-dollar' => 'dollar',
            'col2-euro' => 'euro',
            'custom-amount1' => '0',
            'custom-amount2' => '0.0',//1 decimal place
            'custom-amount3' => '0.00',//2 decimal places
            'custom-currency1' => '#,##0.00',//currency 2 decimal places, no currency/dollar sign
            'custom-currency2' => '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',//w/dollar sign
            'custom-currency3' => '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]',//w/euro sign
            'custom-currency4' => '[$￥-411]#,##0;[RED]-[$￥-411]#,##0', //japanese yen
            'custom-scientific' => '0.00E+000',//-1.23E+003 scientific notation
        );
        $pi = 3.14159;
        $date = '2018-12-31 23:59:59';
        $amount = '5120.5';

        $writer->setAuthor('Some Author');
        $writer->writeSheetHeader('BasicFormats', $sheet1header);
        $writer->writeSheetRow('BasicFormats', array($pi, $pi, $pi, $pi, $pi, $pi, $pi, $pi, $pi));
        $writer->writeSheetHeader('Dates', $sheet2header);
        $writer->writeSheetRow('Dates', array($date, $date, $date, $date, $date, $date, $date, $date, $date));
        $writer->writeSheetHeader('Currencies', $sheet3header);
        $writer->writeSheetRow('Currencies', array($amount, $amount, $amount, $amount, $amount, $amount, $amount, $amount, $amount));
        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }

    function testStyles()
    {

        $filename = "xlsx-styles.xlsx";
        $writer = new XLSXWriter();
        $styles1 = array('font' => 'Arial', 'font-size' => 10, 'font-style' => 'bold', 'fill' => '#eee', 'halign' => 'center', 'border' => 'left,right,top,bottom');
        $styles2 = array(['font-size' => 6], ['font-size' => 8], ['font-size' => 10], ['font-size' => 16]);
        $styles3 = array(['font' => 'Arial'], ['font' => 'Courier New'], ['font' => 'Times New Roman'], ['font' => 'Comic Sans MS']);
        $styles4 = array(['font-style' => 'bold'], ['font-style' => 'italic'], ['font-style' => 'underline'], ['font-style' => 'strikethrough']);
        $styles5 = array(['color' => '#f00'], ['color' => '#0f0'], ['color' => '#00f'], ['color' => '#666']);
        $styles6 = array(['fill' => '#ffc'], ['fill' => '#fcf'], ['fill' => '#ccf'], ['fill' => '#cff']);
        $styles7 = array('border' => 'left,right,top,bottom');
        $styles8 = array(['halign' => 'left'], ['halign' => 'right'], ['halign' => 'center'], ['halign' => 'none']);
        $styles9 = array(array(), ['border' => 'left,top,bottom'], ['border' => 'top,bottom'], ['border' => 'top,bottom,right']);


        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles1);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles2);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles3);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles4);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles5);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles6);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles7);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles8);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles9);
        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";
        $this->assertFileExists($filename);
        unlink($filename);
    }

    function testColors()
    {

        $filename = "xlsx-colors.xlsx";
        $writer = new XLSXWriter();
        $colors = array('ff', 'cc', '99', '66', '33', '00');
        foreach ($colors as $b) {
            foreach ($colors as $g) {
                $rowdata = array();
                $rowstyle = array();
                foreach ($colors as $r) {
                    $rowdata[] = "#$r$g$b";
                    $rowstyle[] = array('fill' => "#$r$g$b");
                }
                $writer->writeSheetRow('Sheet1', $rowdata, $rowstyle);
            }
        }
        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }

    function testNumber250K()
    {

        $filename = "xlsx-numbers-250k.xlsx";
        $writer = new XLSXWriter();
        $writer->writeSheetHeader('Sheet1', array('c1' => 'integer', 'c2' => 'integer', 'c3' => 'integer', 'c4' => 'integer'));//optional
        for ($i = 0; $i < 250000; $i++) {
            $writer->writeSheetRow('Sheet1', array(rand() % 10000, rand() % 10000, rand() % 10000, rand() % 10000));
        }

        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }

    function testString250K()
    {
        $chars = "abcdefghijklmnopqrstuvwxyz0123456789 ";
        $s = '';
        for ($j = 0; $j < 16192; $j++)
            $s .= $chars[rand() % 36];

        $filename = "xlsx-strings-250k.xlsx";
        $writer = new XLSXWriter();
        $writer->writeSheetHeader('Sheet1', array('c1' => 'string', 'c2' => 'string', 'c3' => 'string', 'c4' => 'string'));//optional
        for ($i = 0; $i < 250000; $i++) {
            $s1 = substr($s, rand() % 4000, rand() % 5 + 5);
            $s2 = substr($s, rand() % 8000, rand() % 5 + 5);
            $s3 = substr($s, rand() % 12000, rand() % 5 + 5);
            $s4 = substr($s, rand() % 16000, rand() % 5 + 5);
            $writer->writeSheetRow('Sheet1', array($s1, $s2, $s3, $s4));
        }

        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }

    function testWidths()
    {

        $filename = "xlsx-widths.xlsx";
        $writer = new XLSXWriter();
        $writer->writeSheetHeader('Sheet1', $rowdata = array(300, 234, 456, 789), $col_options = ['widths' => [10, 20, 30, 40]]);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $row_options = ['height' => 20]);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $row_options = ['height' => 30]);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $row_options = ['height' => 40]);

        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }


    function testAdvanced()
    {

        $filename = "xlsx-advanced.xlsx";
        $writer = new XLSXWriter();
        $keywords = array('some', 'interesting', 'keywords');

        $writer->setTitle('Some Title');
        $writer->setSubject('Some Subject');
        $writer->setAuthor('Some Author');
        $writer->setCompany('Some Company');
        $writer->setKeywords($keywords);
        $writer->setDescription('Some interesting description');
        $writer->setTempDir(sys_get_temp_dir());//set custom tempdir

        $sheet1 = 'merged_cells';
        $header = array("string", "string", "string", "string", "string");
        $rows = array(
            array("Merge Cells Example"),
            array(100, 200, 300, 400, 500),
            array(110, 210, 310, 410, 510),
        );
        $writer->writeSheetHeader($sheet1, $header, $col_options = ['suppress_row' => true]);
        foreach ($rows as $row)
            $writer->writeSheetRow($sheet1, $row);
        $writer->markMergedCell($sheet1, $start_row = 0, $start_col = 0, $end_row = 0, $end_col = 4);

        $sheet2 = 'utf8';
        $rows = array(
            array('Spreadsheet', '_'),
            array("Hoja de cálculo", "Hoja de c\xc3\xa1lculo"),
            array("Електронна таблица", "\xd0\x95\xd0\xbb\xd0\xb5\xd0\xba\xd1\x82\xd1\x80\xd0\xbe\xd0\xbd\xd0\xbd\xd0\xb0 \xd1\x82\xd0\xb0\xd0\xb1\xd0\xbb\xd0\xb8\xd1\x86\xd0\xb0"),//utf8 encoded
            array("電子試算表", "\xe9\x9b\xbb\xe5\xad\x90\xe8\xa9\xa6\xe7\xae\x97\xe8\xa1\xa8"),//utf8 encoded
        );
        $writer->writeSheet($rows, $sheet2);

        $sheet3 = 'fonts';
        $format = array('font' => 'Arial', 'font-size' => 10, 'font-style' => 'bold,italic', 'fill' => '#eee', 'color' => '#f00', 'fill' => '#ffc', 'border' => 'top,bottom', 'halign' => 'center');
        $writer->writeSheetRow($sheet3, $row = array(101, 102, 103, 104, 105, 106, 107, 108, 109, 110), $format);
        $writer->writeSheetRow($sheet3, $row = array(201, 202, 203, 204, 205, 206, 207, 208, 209, 210), $format);


        $sheet4 = 'row_options';
        $writer->writeSheetHeader($sheet4, ["col1" => "string", "col2" => "string"], $col_options = array('widths' => [10, 10]));
        $writer->writeSheetRow($sheet4, array(101, 'this text will wrap'), $row_options = array('height' => 30, 'wrap_text' => true));
        $writer->writeSheetRow($sheet4, array(201, 'this text is hidden'), $row_options = array('height' => 30, 'hidden' => true));
        $writer->writeSheetRow($sheet4, array(301, 'this text will not wrap'), $row_options = array('height' => 30, 'collapsed' => true));

        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }


    function testAutofilter()
    {

        $filename = "xlsx-autofilter.xlsx";
        $chars = 'abcdefgh';

        $writer = new XLSXWriter();
        $writer->writeSheetHeader('Sheet1', array('col-string' => 'string', 'col-numbers' => 'integer', 'col-timestamps' => 'datetime'), ['auto_filter' => true, 'widths' => [15, 15, 30]]);
        for ($i = 0; $i < 1000; $i++) {
            $writer->writeSheetRow('Sheet1', array(
                str_shuffle($chars),
                rand() % 10000,
                date('Y-m-d H:i:s', time() - (rand() % 31536000))
            ));
        }
        $writer->writeToFile('xlsx-autofilter.xlsx');

        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }

 function testFreezeRowsColumns()
    {

        $filename = "xlsx-freeze-rows-columns.xlsx";

        $chars = 'abcdefgh';

        $writer = new XLSXWriter();
        $writer->writeSheetHeader('Sheet1', array('c1'=>'string','c2'=>'integer','c3'=>'integer','c4'=>'integer','c5'=>'integer'), ['freeze_rows'=>1, 'freeze_columns'=>1] );
        for($i=0; $i<250; $i++)
        {
            $writer->writeSheetRow('Sheet1', array(
                str_shuffle($chars),
                rand()%10000,
                rand()%10000,
                rand()%10000,
                rand()%10000
            ));
        }

        $writer->writeToFile($filename);

        echo 'Total Memory used: #' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertFileExists($filename);
        unlink($filename);
    }

}
