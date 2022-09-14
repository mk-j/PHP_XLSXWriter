<?php

namespace Test;

include_once __DIR__.'/../vendor/autoload.php';

use PHPUnit\Framework\TestCase;
use SimpleXMLElement;
use XLSXWriter;
use XLSXWriter_BuffererWriter;
use ZipArchive;

class _XLSXWriter_ extends XLSXWriter
{
	public function writeCell(XLSXWriter_BuffererWriter &$file, $row_number, $column_number, $value, $cell_format) {
		return call_user_func_array('parent::writeCell', [&$file, $row_number, $column_number, $value, $cell_format]);
	}
}
//Just a simple test, by no means comprehensive

class XLSXWriterTest extends TestCase
{
	/**
	 * @covers XLSXWriter::writeCell
	 */
	public function testWriteCell() {
		$filename = tempnam("/tmp", "xlsx_writer");
		$file_writer = new XLSXWriter_BuffererWriter($filename);

		$xlsx_writer = new _XLSXWriter_();
		$xlsx_writer->writeCell($file_writer, 0, 0, '0123', 'string');
		$file_writer->close();
		$cell_xml = file_get_contents($filename);
		$this->assertNotEquals('<c r="A1" s="0" t="n"><v>123</v></c>', $cell_xml);
		$this->assertEquals('<c r="A1" s="0" t="s"><v>0</v></c>', $cell_xml);//0123 should be the 0th index of the shared string array
		@unlink($filename);
	}

	/**
	 * @covers XLSXWriter::writeToFile
	 */
	public function testWriteToFile() {
		$filename = tempnam("/tmp", "xlsx_writer");

		$header = ['0'=>'string','1'=>'string','2'=>'string','3'=>'string'];
		$sheet = [
			['55','66','77','88'],
			['10','11','12','13'],
		];

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheet($sheet,'mysheet',$header);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$this->assertTrue($r);

		$this->assertNotEmpty(($zip->numFiles));

		$out_sheet = [];
		for($z=0; $z < $zip->numFiles; $z++) {
			$inside_zip_filename = $zip->getNameIndex($z);

			if (preg_match("/sheet(\d+).xml/", basename($inside_zip_filename))) {
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

	public function testMarkMergedCells() {
		$filename = tempnam("/tmp", "xlsx_writer");

		$header = ['0'=>'string','1'=>'string','2'=>'string','3'=>'string'];
		$sheet = [
			['55','66','77','88'],
			['10','11','12','13'],
		];

		$expected_merged_range = "B2:C3";

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheetHeader('mysheet', $header);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[0]);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[1]);
		$xlsx_writer->markMergedCell('mysheet', 1, 1, 2, 2);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$this->assertTrue($r);

		$this->assertNotEmpty(($zip->numFiles));

		for($z=0; $z < $zip->numFiles; $z++) {
			$inside_zip_filename = $zip->getNameIndex($z);
			$sheet_xml = $zip->getFromName($inside_zip_filename);
			if (preg_match("/sheet(\d+).xml/", basename($inside_zip_filename))) {
				$xml = new SimpleXMLElement($sheet_xml);
				$merged_cell_range = $xml->mergeCells->mergeCell["ref"][0];

				$this->assertEquals($expected_merged_range, $merged_cell_range);
			}
		}

		$zip->close();
		@unlink($filename);
	}

	/**
	 * @dataProvider getFreezeCellsScenarios
	 */
	public function testFreezeCells($freeze_cols, $freeze_rows, $expected_active_cells, $expected_pane) {
		$filename = tempnam("/tmp", "xlsx_writer");

		$header = ['0'=>'string','1'=>'string','2'=>'string','3'=>'string'];
		$sheet = [
			['55','66','77','88'],
			['10','11','12','13'],
		];

		$col_options = ['freeze_columns' => $freeze_cols, 'freeze_rows' => $freeze_rows];

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheetHeader('mysheet', $header, $format = 'xlsx', $delimiter = ';', $subheader = NULL, $col_options);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[0]);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[1]);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$this->assertTrue($r);

		$this->assertNotEmpty(($zip->numFiles));

		for($z=0; $z < $zip->numFiles; $z++) {
			$inside_zip_filename = $zip->getNameIndex($z);
			$sheet_xml = $zip->getFromName($inside_zip_filename);
			if (preg_match("/sheet(\d+).xml/", basename($inside_zip_filename))) {
				$xml = new SimpleXMLElement($sheet_xml);
				$sheet_view = $xml->sheetViews->sheetView;

				if (!empty($expected_pane)) {
					$pane = $sheet_view->pane;
					foreach ($expected_pane as $expected_key => $expected_value) {
						$attribute = (string) $pane[0][$expected_key];
						$this->assertEquals($expected_value, $attribute);
					}
				}

				$selections = $sheet_view->selection;
				for ($i = 0; $i < count($expected_active_cells); $i++) {
					$this->assertNotEmpty($selections[$i]);
					$this->assertEquals($expected_active_cells[$i]['cell'], $selections[$i]['activeCell']);
					$this->assertEquals($expected_active_cells[$i]['cell'], $selections[$i]['sqref']);
					$this->assertEquals($expected_active_cells[$i]['pane'], $selections[$i]['pane']);
				}
			}
		}

		$zip->close();
		@unlink($filename);
	}

	public static function getFreezeCellsScenarios() {
		return [
			"Not frozen" => [
				$freeze_cols = false,
				$freeze_rows = false,
				$expected_active_cells = [["cell" => "A1", "pane" => "topLeft"]],
				$expected_pane = [],
			],
			"Frozen Col B and Row 2" => [
				$freeze_cols = 1,
				$freeze_rows = 1,
				$expected_active_cells = [["cell" => "A2", "pane" => "topRight"], ["cell" => "B1", "pane" => "bottomLeft"], ["cell" => "B2", "pane" => "bottomRight"]],
				$expected_pane = ["ySplit" => $freeze_rows, "xSplit" => $freeze_cols, "topLeftCell" => "B2", "activePane" => "bottomRight"],
			],
			"Frozen Col B" => [
				$freeze_cols = 1,
				$freeze_rows = false,
				$expected_active_cells = [["cell" => "B1", "pane" => "topRight"]],
				$expected_pane = ["xSplit" => $freeze_cols, "topLeftCell" => "B1", "activePane" => "topRight"],
			],
			"Frozen Row 2" => [
				$freeze_cols = false,
				$freeze_rows = 1,
				$expected_active_cells = [["cell" => "A2", "pane" => "bottomLeft"]],
				$expected_pane = ["ySplit" => $freeze_rows, "topLeftCell" => "A2", "activePane" => "bottomLeft"],
			],
			"Frozen Col A and Row 1" => [
				$freeze_cols = 0,
				$freeze_rows = 0,
				$expected_active_cells = [["cell" => "A1", "pane" => "topRight"], ["cell" => "A1", "pane" => "bottomLeft"], ["cell" => "A1", "pane" => "bottomRight"]],
				$expected_pane = ["ySplit" => $freeze_rows, "xSplit" => $freeze_cols, "topLeftCell" => "A1", "activePane" => "bottomRight"],
			],
			"Frozen Col A" => [
				$freeze_cols = 0,
				$freeze_rows = false,
				$expected_active_cells = [["cell" => "A1", "pane" => "topRight"]],
				$expected_pane = ["xSplit" => $freeze_cols, "topLeftCell" => "A1", "activePane" => "topRight"],
			],
			"Frozen Row 1" => [
				$freeze_cols = false,
				$freeze_rows = 0,
				$expected_active_cells = [["cell" => "A1", "pane" => "bottomLeft"]],
				$expected_pane = ["ySplit" => $freeze_rows, "topLeftCell" => "A1", "activePane" => "bottomLeft"],
			],
		];
	}

	private function stripCellsFromSheetXML($sheet_xml) {
		$output = [];

		$xml = new SimpleXMLElement($sheet_xml);

		for ($i = 0; $i < count($xml->sheetData->row); $i++) {
			$row = $xml->sheetData->row[$i];
			for ($j = 0; $j < count($row->c); $j ++) {
				$output[$i][$j] = (string)$row->c[$j]->v;
			}
		}

		return $output;
	}

	public static function array_diff_assoc_recursive($array1, $array2) {
		$difference = [];
		foreach($array1 as $key => $value) {
			if(is_array($value)) {
				if(!isset($array2[$key]) || !is_array($array2[$key])) {
					$difference[$key] = $value;
				} else {
					$new_diff = self::array_diff_assoc_recursive($value, $array2[$key]);
					if(!empty($new_diff)) {
						$difference[$key] = $new_diff;
					}
				}
			} else if(!isset($array2[$key]) || $array2[$key] != $value) {
				$difference[$key] = $value;
			}
		}

		return empty($difference) ? [] : $difference;
	}
}
