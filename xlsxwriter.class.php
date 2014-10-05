<?php
/*
 * @license MIT License
 * */

if (!class_exists('ZipArchive')) { throw new Exception('ZipArchive not found'); }

class XLSXWriter
{
	//------------------------------------------------------------------
	protected $author ='Doc Author';
	protected $sheets_meta = array();
	protected $shared_strings = array();//unique set
	protected $shared_string_count = 0;//count of non-unique references to the unique set
	protected $temp_files = array();

	protected $file;
	protected $row_num;
	protected $header_offset;
	protected $cell_formats_arr;

	public function __construct()
	{
		if(!ini_get('date.timezone'))
		{
			//using date functions can kick out warning if this isn't set
			date_default_timezone_set('UTC');
		}
	}

	public function setAuthor($author='') { $this->author=$author; }

	public function __destruct()
	{
		if (!empty($this->temp_files)) {
			foreach($this->temp_files as $temp_file) {
				@unlink($temp_file);
			}
		}
	}

	protected function tempFilename()
	{
		$filename = tempnam("/tmp", "xlsx_writer_");
		$this->temp_files[] = $filename;
		return $filename;
	}

	public function writeToStdOut()
	{
		$temp_file = $this->tempFilename();
		self::writeToFile($temp_file);
		readfile($temp_file);
	}

	public function writeToString()
	{
		$temp_file = $this->tempFilename();
		self::writeToFile($temp_file);
		$string = file_get_contents($temp_file);
		return $string;
	}

	public function writeToFile($filename)
	{
		@unlink($filename);//if the zip already exists, overwrite it
		$zip = new ZipArchive();
		if (empty($this->sheets_meta))                  { self::log("Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined."); return; }
		if (!$zip->open($filename, ZipArchive::CREATE)) { self::log("Error in ".__CLASS__."::".__FUNCTION__.", unable to create zip."); return; }
		
		$zip->addEmptyDir("docProps/");
		$zip->addFromString("docProps/app.xml" , self::buildAppXML() );
		$zip->addFromString("docProps/core.xml", self::buildCoreXML());

		$zip->addEmptyDir("_rels/");
		$zip->addFromString("_rels/.rels", self::buildRelationshipsXML());

		$zip->addEmptyDir("xl/worksheets/");
		foreach($this->sheets_meta as $sheet_meta) {
			$zip->addFile($sheet_meta['filename'], "xl/worksheets/".$sheet_meta['xmlname'] );
		}
		if (!empty($this->shared_strings)) {
			$zip->addFile($this->writeSharedStringsXML(), "xl/sharedStrings.xml" );  //$zip->addFromString("xl/sharedStrings.xml",     self::buildSharedStringsXML() );
		}
		$zip->addFromString("xl/workbook.xml"         , self::buildWorkbookXML() );
		$zip->addFile($this->writeStylesXML(), "xl/styles.xml" );  //$zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
		$zip->addFromString("[Content_Types].xml"     , self::buildContentTypesXML() );

		$zip->addEmptyDir("xl/_rels/");
		$zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelsXML() );
		$zip->close();
	}

	public function writeSheet(array $data, $sheet_name='', array $header_types=array() )
	{
		$data = empty($data) ? array(array('')) : $data;

		$row_count = count($data);
		$column_count = count($data[self::array_first_key($data)]);

		$this->writeSheetHead($row_count, $column_count, $sheet_name, $header_types);

		foreach($data as $i=>$row)
		{
			$this->writeSheetRow($row);
		}

		$this->writeSheetFooter();
	}
	
	public function writeSheetHead($row_count, $column_count, $sheet_name='', array $header_types=array() )
	{
		$sheet_filename = $this->tempFilename();
		$sheet_default = 'Sheet' . (count($this->sheets_meta) + 1);
		$sheet_name = !empty($sheet_name) ? $sheet_name : $sheet_default;
		$this->sheets_meta[] = array('filename' => $sheet_filename, 'sheetname' => $sheet_name, 'xmlname' => strtolower($sheet_default) . ".xml");

		$this->header_offset = empty($header_types) ? 0 : 1;
		$row_count = $row_count + $this->header_offset;
		$max_cell = self::xlsCell($row_count - 1, $column_count - 1);

		$tabselected = count($this->sheets_meta) == 1 ? 'true' : 'false';//only first sheet is selected
		$this->cell_formats_arr = empty($header_types) ? array_fill(0, $column_count, 'string') : array_values($header_types);
		$header_row = empty($header_types) ? array() : array_keys($header_types);

		$this->file = new XLSXWriter_BuffererWriter($sheet_filename);
		$this->file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
		$this->file->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
		$this->file->write('<sheetPr filterMode="false">');
		$this->file->write('<pageSetUpPr fitToPage="false"/>');
		$this->file->write('</sheetPr>');
		$this->file->write('<dimension ref="A1:' . $max_cell . '"/>');
		$this->file->write('<sheetViews>');
		$this->file->write('<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="' . $tabselected . '" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
		$this->file->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
		$this->file->write('</sheetView>');
		$this->file->write('</sheetViews>');
		$this->file->write('<cols>');
		$this->file->write('<col collapsed="false" hidden="false" max="1025" min="1" style="0" width="11.5"/>');
		$this->file->write('</cols>');
		$this->file->write('<sheetData>');
		if (!empty($header_row)) {
			$this->file->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . (1) . '">');
			foreach ($header_row as $k => $v) {
				$this->writeCell($this->file, 0, $k, $v, $cell_format = 'string');
			}
			$this->file->write('</row>');
		}
		$this->row_num = 0;
	}

	public function writeSheetRow($row)
	{
		$this->file->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($this->row_num + $this->header_offset + 1) . '">');
		foreach ($row as $k => $v) {
			$this->writeCell($this->file, $this->row_num + $this->header_offset, $k, $v, $this->cell_formats_arr[$k]);
		}
		$this->file->write('</row>');
		$this->row_num++;
	}

	public function writeSheetFooter()
	{
		$this->file->write(    '</sheetData>');
		$this->file->write(    '<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
		$this->file->write(    '<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
		$this->file->write(    '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
		$this->file->write(    '<headerFooter differentFirst="false" differentOddEven="false">');
		$this->file->write(        '<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
		$this->file->write(        '<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
		$this->file->write(    '</headerFooter>');
		$this->file->write('</worksheet>');
		$this->file->close();
	}

	protected function writeCell(XLSXWriter_BuffererWriter &$file, $row_number, $column_number, $value, $cell_format)
	{
		static $styles = array('money'=>1,'dollar'=>1,'datetime'=>2,'date'=>3,'string'=>0);
		$cell = self::xlsCell($row_number, $column_number);
		$s = isset($styles[$cell_format]) ? $styles[$cell_format] : '0';
		
		if (!is_scalar($value) || $value=='') { //objects, array, empty
			$file->write('<c r="'.$cell.'" s="'.$s.'"/>');
		} elseif ($cell_format=='date') {
			$file->write('<c r="'.$cell.'" s="'.$s.'" t="n"><v>'.intval(self::convert_date_time($value)).'</v></c>');
		} elseif ($cell_format=='datetime') {
			$file->write('<c r="'.$cell.'" s="'.$s.'" t="n"><v>'.self::convert_date_time($value).'</v></c>');
		} elseif (!is_string($value)) {
			$file->write('<c r="'.$cell.'" s="'.$s.'" t="n"><v>'.($value*1).'</v></c>');//int,float, etc
		} elseif ($value{0}!='0' && is_numeric($value)){ //excel wants to trim leading zeros
			$file->write('<c r="'.$cell.'" s="'.$s.'" t="n"><v>'.($value*1).'</v></c>');//int,float, etc
		} elseif ($value{0}=='='){
			$file->write('<c r="'.$cell.'" s="'.$s.'" t="s"><f>'.self::xmlspecialchars($value).'</f></c>');
		} elseif ($value!==''){
			$file->write('<c r="'.$cell.'" s="'.$s.'" t="s"><v>'.self::xmlspecialchars($this->setSharedString($value)).'</v></c>');
		}
	}

	protected function writeStylesXML()
	{
		$temporary_filename = $this->tempFilename();
		$file = new XLSXWriter_BuffererWriter($temporary_filename);
		$file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		$file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
		$file->write('<numFmts count="4">');
		$file->write(		'<numFmt formatCode="GENERAL" numFmtId="164"/>');
		$file->write(		'<numFmt formatCode="[$$-1009]#,##0.00;[RED]\-[$$-1009]#,##0.00" numFmtId="165"/>');
		$file->write(		'<numFmt formatCode="YYYY-MM-DD\ HH:MM:SS" numFmtId="166"/>');
		$file->write(		'<numFmt formatCode="YYYY-MM-DD" numFmtId="167"/>');
		$file->write('</numFmts>');
		$file->write('<fonts count="4">');
		$file->write(		'<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write('</fonts>');
		$file->write('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
		$file->write('<borders count="1"><border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border></borders>');
		$file->write(	'<cellStyleXfs count="20">');
		$file->write(		'<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
		$file->write(		'<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
		$file->write(		'<protection hidden="false" locked="true"/>');
		$file->write(		'</xf>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
		$file->write(	'</cellStyleXfs>');
		$file->write(	'<cellXfs count="4">');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="165" xfId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="166" xfId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="167" xfId="0"/>');
		$file->write(	'</cellXfs>');
		$file->write(	'<cellStyles count="6">');
		$file->write(		'<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
		$file->write(		'<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
		$file->write(		'<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
		$file->write(		'<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
		$file->write(		'<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
		$file->write(		'<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
		$file->write(	'</cellStyles>');
		$file->write('</styleSheet>');
		$file->close();
		return $temporary_filename;
	}

	protected function setSharedString($v)
	{
		if (isset($this->shared_strings[$v]))
		{
			$string_value = $this->shared_strings[$v];
		}
		else
		{
			$string_value = count($this->shared_strings);
			$this->shared_strings[$v] = $string_value;
		}
		$this->shared_string_count++;//non-unique count
		return $string_value;
	}

	protected function writeSharedStringsXML()
	{
		$temporary_filename = $this->tempFilename();
		$file = new XLSXWriter_BuffererWriter($temporary_filename, $fd_flags='w', $check_utf8=true);
		$file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		$file->write('<sst count="'.($this->shared_string_count).'" uniqueCount="'.count($this->shared_strings).'" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
		foreach($this->shared_strings as $s=>$c)
		{
			$file->write('<si><t>'.self::xmlspecialchars($s).'</t></si>');
		}
		$file->write('</sst>');
		$file->close();
		
		return $temporary_filename;
	}

	protected function buildAppXML()
	{
		$app_xml="";
		$app_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$app_xml.='<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime></Properties>';
		return $app_xml;
	}

	protected function buildCoreXML()
	{
		$core_xml="";
		$core_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$core_xml.='<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
		$core_xml.='<dcterms:created xsi:type="dcterms:W3CDTF">'.date("Y-m-d\TH:i:s.00\Z").'</dcterms:created>';//$date_time = '2013-07-25T15:54:37.00Z';
		$core_xml.='<dc:creator>'.self::xmlspecialchars($this->author).'</dc:creator>';
		$core_xml.='<cp:revision>0</cp:revision>';
		$core_xml.='</cp:coreProperties>';
		return $core_xml;
	}

	protected function buildRelationshipsXML()
	{
		$rels_xml="";
		$rels_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$rels_xml.='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$rels_xml.='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
		$rels_xml.='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
		$rels_xml.='<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
		$rels_xml.="\n";
		$rels_xml.='</Relationships>';
		return $rels_xml;
	}

	protected function buildWorkbookXML()
	{
		$workbook_xml="";
		$workbook_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$workbook_xml.='<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		$workbook_xml.='<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
		$workbook_xml.='<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
		$workbook_xml.='<sheets>';
		foreach($this->sheets_meta as $i=>$sheet_meta) {
			$workbook_xml.='<sheet name="'.self::xmlspecialchars($sheet_meta['sheetname']).'" sheetId="'.($i+1).'" state="visible" r:id="rId'.($i+2).'"/>';
		}
		$workbook_xml.='</sheets>';
		$workbook_xml.='<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
		return $workbook_xml;
	}

	protected function buildWorkbookRelsXML()
	{
		$wkbkrels_xml="";
		$wkbkrels_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$wkbkrels_xml.='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$wkbkrels_xml.='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
		foreach($this->sheets_meta as $i=>$sheet_meta) {
			$wkbkrels_xml.='<Relationship Id="rId'.($i+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/'.($sheet_meta['xmlname']).'"/>';
		}
		if (!empty($this->shared_strings)) {
			$wkbkrels_xml.='<Relationship Id="rId'.(count($this->sheets_meta)+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
		}
		$wkbkrels_xml.="\n";
		$wkbkrels_xml.='</Relationships>';
		return $wkbkrels_xml;
	}

	protected function buildContentTypesXML()
	{
		$content_types_xml="";
		$content_types_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$content_types_xml.='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
		$content_types_xml.='<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		$content_types_xml.='<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		foreach($this->sheets_meta as $i=>$sheet_meta) {
			$content_types_xml.='<Override PartName="/xl/worksheets/'.($sheet_meta['xmlname']).'" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
		}
		if (!empty($this->shared_strings)) {
			$content_types_xml.='<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
		}
		$content_types_xml.='<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
		$content_types_xml.='<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
		$content_types_xml.='<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		$content_types_xml.='<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		$content_types_xml.="\n";
		$content_types_xml.='</Types>';
		return $content_types_xml;
	}

	//------------------------------------------------------------------
	/*
	 * @param $row_number int, zero based
	 * @param $column_number int, zero based
	 * @return Cell label/coordinates, ex: A1, C3, AA42
	 * */
	public static function xlsCell($row_number, $column_number)
	{
		$n = $column_number;
		for($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
			$r = chr($n%26 + 0x41) . $r;
		}
		return $r . ($row_number+1);
	}
	//------------------------------------------------------------------
	public static function log($string)
	{
		file_put_contents("php://stderr", date("Y-m-d H:i:s:").rtrim(is_array($string) ? json_encode($string) : $string)."\n");
	}
	//------------------------------------------------------------------
	public static function sanitize_filename($filename) //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
	{
		$nonprinting = array_map('chr', range(0,31));
		$invalid_chars = array('<', '>', '?', '"', ':', '|', '\\', '/', '*', '&');
		$all_invalids = array_merge($nonprinting,$invalid_chars);
		return str_replace($all_invalids, "", $filename);
	}
	//------------------------------------------------------------------
	public static function xmlspecialchars($val)
	{
		return str_replace("'", "&#39;", htmlspecialchars($val));
	}
	//------------------------------------------------------------------
	public static function array_first_key(array $arr)
	{
		reset($arr);
		$first_key = key($arr);
		return $first_key;
	}
	//------------------------------------------------------------------
	public static function convert_date_time($date_input) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
	{
		$days    = 0;    # Number of days since epoch
		$seconds = 0;    # Time expressed as fraction of 24h hours in seconds
		$year=$month=$day=0;
		$hour=$min  =$sec=0;

		$date_time = $date_input;
		if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches))
		{
			list($junk,$year,$month,$day) = $matches;
		}
		if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", $date_time, $matches))
		{
			list($junk,$hour,$min,$sec) = $matches;
			$seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
		}

		//using 1900 as epoch, not 1904, ignoring 1904 special case
		
		# Special cases for Excel.
		if ("$year-$month-$day"=='1899-12-31')  return $seconds      ;    # Excel 1900 epoch
		if ("$year-$month-$day"=='1900-01-00')  return $seconds      ;    # Excel 1900 epoch
		if ("$year-$month-$day"=='1900-02-29')  return 60 + $seconds ;    # Excel false leapday

		# We calculate the date by calculating the number of days since the epoch
		# and adjust for the number of leap days. We calculate the number of leap
		# days by normalising the year in relation to the epoch. Thus the year 2000
		# becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
		$epoch  = 1900;
		$offset = 0;
		$norm   = 300;
		$range  = $year - $epoch;

		# Set month days and check for leap year.
		$leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
		$mdays = array( 31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );

		# Some boundary checks
		if($year < $epoch || $year > 9999) return 0;
		if($month < 1     || $month > 12)  return 0;
		if($day < 1       || $day > $mdays[ $month - 1 ]) return 0;

		# Accumulate the number of days since the epoch.
		$days = $day;    # Add days for current month
		$days += array_sum( array_slice($mdays, 0, $month-1 ) );    # Add days for past months
		$days += $range * 365;                      # Add days for past years
		$days += intval( ( $range ) / 4 );             # Add leapdays
		$days -= intval( ( $range + $offset ) / 100 ); # Subtract 100 year leapdays
		$days += intval( ( $range + $offset + $norm ) / 400 );  # Add 400 year leapdays
		$days -= $leap;                                      # Already counted above

		# Adjust for Excel erroneously treating 1900 as a leap year.
		if ($days > 59) { $days++;}

		return $days + $seconds;
	}
	//------------------------------------------------------------------
}

class XLSXWriter_BuffererWriter
{
	protected $fd=null;
	protected $buffer='';
	protected $check_utf8=false;

	public function __construct($filename, $fd_fopen_flags='w', $check_utf8=false)
	{
		$this->check_utf8 = $check_utf8;
		$this->fd = fopen($filename, $fd_fopen_flags);
		if ($this->fd===false) {
			XLSXWriter::log("Unable to open $filename for writing.");
		}
	}

	public function write($string)
	{
		$this->buffer.=$string;
		if (isset($this->buffer[8191])) {
			$this->purge();
		}
	}

	protected function purge()
	{
		if ($this->fd) {
			if ($this->check_utf8 && !self::isValidUTF8($this->buffer)) {
				XLSXWriter::log("Error, invalid UTF8 encoding detected.");
				$this->check_utf8 = false;
			}
			fwrite($this->fd, $this->buffer);
			$this->buffer='';
		}
	}

	public function close()
	{
		$this->purge();
		if ($this->fd) {
			fclose($this->fd);
			$this->fd=null;
		}
	}

	public function __destruct() 
	{
		$this->close();
	}

	protected static function isValidUTF8($string)
	{
		if (function_exists('mb_check_encoding'))
		{
			return mb_check_encoding($string, 'UTF-8') ? true : false;
		}
		return preg_match("//u", $string) ? true : false;
	}
}



// vim: set filetype=php expandtab tabstop=4 shiftwidth=4 autoindent smartindent:
