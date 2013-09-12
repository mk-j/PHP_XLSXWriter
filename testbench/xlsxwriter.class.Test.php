<?php
include_once 'PHPUnit/Framework/TestCase.php'; // located in /usr/share/php
include_once 'xlsxwriter.class.php';

//Just a simple test, by no means comprehensive

class XLSXWriterTest extends PHPUnit_Framework_TestCase
{

    /**
     * @covers XLSXWriter::writeToFile
     */
	public function testWriteToFile()
	{
		$filename = tempnam("/tmp", "xlsx_writer");
		
		$header = array('0'=>'string','1'=>'string','2'=>'string','3'=>'string');
		$sheet = array(
				array('55','66','77','88'),
				array('10','11','12','13'),
		);

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheet($sheet,'mysheet',$header);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$this->assertTrue($r);

		$r = $zip->numFiles>0 ? true : false;
		$this->assertTrue($r);

		$out_sheet = array();
		for($z=0; $z<$zip->numFiles; $z++)
		{
			$inside_zip_filename = $zip->getNameIndex($z);
			if (preg_match("/sheet(\d+).xml/", basename($inside_zip_filename)))
			{
				$out_sheet = $this->stripCellsFromSheetXML($zip->getFromName($inside_zip_filename));
				array_shift($out_sheet);
				$out_sheet = array_values($out_sheet);
			}
		}
		$zip->close();
		@unlink($filename);
		
		$r1 = self::array_diff_assoc_recursive($out_sheet, $sheet);
		$r2 = self::array_diff_assoc_recursive($sheet, $out_sheet);
		$this->assertEmpty($r1);
		$this->assertEmpty($r2);
	}

	private function stripCellsFromSheetXML($sheet_xml)
	{
		$output=array();
		$xml = new SimpleXMLElement($sheet_xml);
		$i=0;
		foreach($xml->sheetData->row as $row)
		{
			$j=0;
			foreach($row->c as $c)
			{
				$output[$i][$j] = (string)$c->v;
				$j++;
			}
			$i++;
		}
		return $output;
	}

    public static function array_diff_assoc_recursive($array1, $array2)
    {
        foreach($array1 as $key => $value)
        {
            if(is_array($value))
            {
                if(!isset($array2[$key]) || !is_array($array2[$key]))
                {
                    $difference[$key] = $value;
                }
                else
                {
                    $new_diff = self::array_diff_assoc_recursive($value, $array2[$key]);
                    if(!empty($new_diff))
                    {
                        $difference[$key] = $new_diff;
                    }
                }
            }
            else if(!isset($array2[$key]) || $array2[$key] != $value)
            {
                $difference[$key] = $value;
            }
        }
        return !isset($difference) ? array() : $difference;
    }
}
