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
	public function writeCell(XLSXWriter_BuffererWriter &$file, $row_number, $column_number, $value, $cell_format, $cell_style_idx = null) {
		return call_user_func_array('parent::writeCell', [&$file, $row_number, $column_number, $value, $cell_format, $cell_style_idx = null]);
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
		$xml = $this->extractSheetXml($zip);

		$this->assertTrue($r);
		$this->assertNotEmpty(($zip->numFiles));
		$this->assertNotEmpty($xml);

		$merged_cell_range = $xml->mergeCells->mergeCell["ref"][0];

		$this->assertEquals($expected_merged_range, $merged_cell_range);

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
		$xml = $this->extractSheetXml($zip);

		$this->assertTrue($r);
		$this->assertNotEmpty(($zip->numFiles));
		$this->assertNotEmpty($xml);

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

	public function testColumnsWidths() {
		$filename = tempnam("/tmp", "xlsx_writer");

		$header = ['0'=>'string','1'=>'string','2'=>'string','3'=>'string'];
		$sheet = [
			['55','66','77','88'],
			['10','11','12','13'],
		];

		$widths = [10, 20, 30, 40];

		$col_options = ['widths' => $widths];

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheetHeader('mysheet', $header, $format = 'xlsx', $delimiter = ';', $subheader = NULL, $col_options);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[0]);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[1]);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$xml = $this->extractSheetXml($zip);

		$this->assertTrue($r);
		$this->assertNotEmpty(($zip->numFiles));
		$this->assertNotEmpty($xml);

		$cols = $xml->cols->col;
		foreach ($widths as $col_index => $col_width) {
			$col = $cols[$col_index];
			$this->assertFalse(filter_var($col["collapsed"], FILTER_VALIDATE_BOOLEAN));
			$this->assertFalse(filter_var($col["hidden"], FILTER_VALIDATE_BOOLEAN));
			$this->assertTrue(filter_var($col["customWidth"], FILTER_VALIDATE_BOOLEAN));
			$this->assertEquals($col_index + 1, (string) $col["max"]);
			$this->assertEquals($col_index + 1, (string) $col["min"]);
			$this->assertEquals("0", (string) $col["style"]);
			$this->assertEquals($col_width, (string) $col["width"]);
		}
		$last_col_index = count($widths);
		$last_col = $cols[$last_col_index];
		$this->assertFalse(filter_var($last_col["collapsed"], FILTER_VALIDATE_BOOLEAN));
		$this->assertFalse(filter_var($last_col["hidden"], FILTER_VALIDATE_BOOLEAN));
		$this->assertFalse(filter_var($last_col["customWidth"], FILTER_VALIDATE_BOOLEAN));
		$this->assertEquals("1024", (string) $last_col["max"]);
		$this->assertEquals($last_col_index + 1, (string) $last_col["min"]);
		$this->assertEquals("0", (string) $last_col["style"]);
		$this->assertEquals("11.5", (string) $last_col["width"]);

		$zip->close();
		@unlink($filename);
	}

	public function testRowHeight() {
		$filename = tempnam("/tmp", "xlsx_writer");

		$sheet = [
			['55','66','77','88'],
			['10','11','12','13'],
		];

		$custom_height = 20.5;

		$row_options = ['height' => $custom_height];

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheetRow('mysheet', $sheet[0], $format = 'xlsx', $delimiter = ';', $row_options);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[1]);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$xml = $this->extractSheetXml($zip);

		$this->assertTrue($r);
		$this->assertNotEmpty(($zip->numFiles));
		$this->assertNotEmpty($xml);

		$rows = $xml->sheetData->row;
		$this->assertRowProperties($custom_height, $expected_custom_height = true, $expected_hidden = false, $expected_collapsed = false, $rows[0]);
		$this->assertRowProperties($expected_height = 12.1, $expected_custom_height = false, $expected_hidden = false, $expected_collapsed = false, $rows[1]);

		$zip->close();
		@unlink($filename);
	}

	public function testRowHidden() {
		$filename = tempnam("/tmp", "xlsx_writer");

		$sheet = [
			['55','66','77','88'],
			['10','11','12','13'],
		];

		$row_options = ['hidden' => true];

		$expected_height = 12.1;
		$expected_custom_height = false;
		$expected_collapsed = false;

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheetRow('mysheet', $sheet[0], $format = 'xlsx', $delimiter = ';', $row_options);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[1]);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$xml = $this->extractSheetXml($zip);

		$this->assertTrue($r);
		$this->assertNotEmpty(($zip->numFiles));
		$this->assertNotEmpty($xml);

		$rows = $xml->sheetData->row;
		$this->assertRowProperties($expected_height, $expected_custom_height, $expected_hidden = true, $expected_collapsed, $rows[0]);
		$this->assertRowProperties($expected_height, $expected_custom_height, $expected_hidden = false, $expected_collapsed, $rows[1]);

		$zip->close();
		@unlink($filename);
	}

	public function testRowCollapsed() {
		$filename = tempnam("/tmp", "xlsx_writer");

		$sheet = [
			['55','66','77','88'],
			['10','11','12','13'],
		];

		$row_options = ['collapsed' => true];

		$expected_height = 12.1;
		$expected_custom_height = false;
		$expected_hidden = false;

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheetRow('mysheet', $sheet[0], $format = 'xlsx', $delimiter = ';', $row_options);
		$xlsx_writer->writeSheetRow('mysheet', $sheet[1]);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$xml = $this->extractSheetXml($zip);

		$this->assertTrue($r);
		$this->assertNotEmpty(($zip->numFiles));
		$this->assertNotEmpty($xml);

		$rows = $xml->sheetData->row;
		$this->assertRowProperties($expected_height, $expected_custom_height, $expected_hidden, $expected_collapsed = true, $rows[0]);
		$this->assertRowProperties($expected_height, $expected_custom_height, $expected_hidden, $expected_collapsed = false, $rows[1]);

		$zip->close();
		@unlink($filename);
	}

	public function testAddBorder() {
		$filename = tempnam("/tmp", "xlsx_writer");

		$header = ["0"=>"string", "1"=>"string", "2"=>"string", "3"=>"string"];
		$sheet = [
			["55", "66", "77", "88"],
			["10", "11", "12", "13"],
		];

		$expected_borders = ["right", "left", "top", "bottom"];
		$expected_border_style = "thick";
		$expected_border_color_base = "ff99cc";
		$expected_border_color = "FFFF99CC";

		$row_options = [
			"border" => implode(",", $expected_borders),
			"border-style" => $expected_border_style,
			"border-color" => "#$expected_border_color_base" ,
		];

		$xlsx_writer = new XLSXWriter();
		$xlsx_writer->writeSheetHeader("mysheet", $header);
		$xlsx_writer->writeSheetRow("mysheet", $sheet[0], $format = "xlsx", $delimiter = ";", $row_options);
		$xlsx_writer->writeSheetRow("mysheet", $sheet[1]);
		$xlsx_writer->writeToFile($filename);

		$zip = new ZipArchive();
		$r = $zip->open($filename);
		$xml = $this->extractSheetXml($zip);
		$styles = $this->extractStyleXml($zip);

		$this->assertTrue($r);
		$this->assertNotEmpty(($zip->numFiles));
		$this->assertNotEmpty($xml);
		$this->assertNotEmpty($styles);

		$border_styles = $styles->borders;
		$this->assertBorderStyle($expected_border_style, $expected_border_color, $border = $border_styles->border[1]);
		$this->assertFillStyle($expected_pattern = "solid", $expected_bg_color = "FF003300", $styles->fills->fill[2]);
		$this->assertFontStyle($expected_font_name = "Arial", $expected_is_bold = "true", $styles->fonts->font[4]);

		$cell_styles = $styles->cellXfs->xf;
		$this->assertCellStyle($expected_apply_border_string = "false", $expected_border_id = 0, $expected_fill_id = 2, $expected_font_id = 4, $cell_styles[6]);
		$this->assertCellStyle($expected_apply_border_string = "true", $expected_border_id = 1, $expected_fill_id = 0, $expected_font_id = 0, $cell_styles[7]);

		$rows = $xml->sheetData->row;
		$this->assertRowHasStyleIndex($rows[0], $expected_header_style = 6);
		$this->assertRowHasStyleIndex($rows[1], $expected_style = 7);

		$zip->close();
		@unlink($filename);
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

	private function extractSheetXml($zip) {
		for($z=0; $z < $zip->numFiles; $z++) {
			$inside_zip_filename = $zip->getNameIndex($z);
			$sheet_xml = $zip->getFromName($inside_zip_filename);
			if (preg_match("/sheet(\d+).xml/", basename($inside_zip_filename))) {
				return new SimpleXMLElement($sheet_xml);
			}
		}

		return null;
	}

	private function extractStyleXml($zip) {
		for($z=0; $z < $zip->numFiles; $z++) {
			$inside_zip_filename = $zip->getNameIndex($z);
			$xml = $zip->getFromName($inside_zip_filename);
			if (preg_match("/styles.xml/", basename($inside_zip_filename))) {
				return new SimpleXMLElement($xml);
			}
		}

		return null;
	}

	private function assertRowProperties($expected_height, $expected_custom_height, $expected_hidden, $expected_collapsed, $row) {
		$this->assertEquals($expected_height, (string)$row['ht']);
		$this->assertEquals($expected_custom_height, filter_var($row['customHeight'], FILTER_VALIDATE_BOOLEAN));
		$this->assertEquals($expected_hidden, filter_var($row['hidden'], FILTER_VALIDATE_BOOLEAN));
		$this->assertEquals($expected_collapsed, filter_var($row['collapsed'], FILTER_VALIDATE_BOOLEAN));
	}

	private function assertCellStyle($expected_apply_border_string, $expected_border_id, $expected_fill_id, $expected_font_id, $cell_style) {
		$this->assertEquals($expected_apply_border_string, $cell_style["applyBorder"]);
		$this->assertEquals($expected_border_id, (int)$cell_style["borderId"]);
		$this->assertEquals($expected_fill_id, (int)$cell_style["fillId"]);
		$this->assertEquals($expected_font_id, (int)$cell_style["fontId"]);
	}

	private function assertBorderStyle($expected_border_style, $expected_border_color, $border) {
		$this->assertEquals($expected_border_style, $border->left["style"]);
		$this->assertEquals($expected_border_style, $border->right["style"]);
		$this->assertEquals($expected_border_style, $border->top["style"]);
		$this->assertEquals($expected_border_style, $border->bottom["style"]);

		$this->assertEquals($expected_border_color, $border->left->color["rgb"]);
		$this->assertEquals($expected_border_color, $border->right->color["rgb"]);
		$this->assertEquals($expected_border_color, $border->top->color["rgb"]);
		$this->assertEquals($expected_border_color, $border->bottom->color["rgb"]);
	}

	private function assertFillStyle($expected_pattern, $expected_bg_color, $fill) {
		$this->assertEquals($expected_pattern, $fill->patternFill["patternType"]);
		if (!empty($expected_bg_color)) {
			$this->assertEquals($expected_bg_color, $fill->patternFill->bgColor["rgb"]);
		}
	}

	private function assertFontStyle($expected_font_name, $expected_is_bold, $font) {
 		$this->assertEquals($expected_font_name, $font->name["val"]);
 		if (!empty($expected_is_bold)) {
		    $this->assertEquals($expected_is_bold, $font->b["val"]);
	    }
	}

	private function assertRowHasStyleIndex($row, $expected_style) {
		foreach ($row->c as $cell) {
			$this->assertEquals($expected_style, (int)$cell["s"]);
		}
	}
}
