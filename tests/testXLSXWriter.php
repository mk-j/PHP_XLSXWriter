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
            array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
            array('2003','=B1', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'),
        );


        foreach($rows as $row)
            $writer->writeSheetRow('Sheet1', $row);
        $writer->writeToFile($filename);
        $this->assertFileExists($filename);
        unlink($filename);
    }

}
