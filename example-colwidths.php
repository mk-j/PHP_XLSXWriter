<?php
		include_once('xlsxwriter.class.php');
		ini_set('display_errors', 0);
		ini_set('log_errors', 1);
		error_reporting(E_ALL & ~E_NOTICE);
		$filename = "example-colwidths.xlsx";
		header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Transfer-Encoding: binary');
		header('Cache-Control: must-revalidate');
		header('Pragma: public');
		
		//setup heading row style
		$hstyle = array( 'font'=>'Arial','font-size'=>10,'font-style'=>'bold', 'halign'=>'center', 'border'=>'bottom');
		$hdr = array('Level'=>'string', 'Function'=>'string','ID'=>'string','Message'=>'string');
		
		//setup data array
		$db[]=array("1","test1","203","test function working");
		$db[]=array("2","test2","204","test function failing");
		
		//write header then sheet data and output file
		$writer = new XLSXWriter();
		
		//write sheet with standard widths
		$writer->writeSheetRow('Ex1',array_keys($hdr),$hstyle);
		$writer->writeSheet($db,'Ex1',$hdr,true);
		
		//set column width for all columns in a sheet and write sheet
		$writer->setColWidth("40");
		$writer->writeSheetRow('Ex2',array_keys($hdr),$hstyle);
		$writer->writeSheet($db,'Ex2',$hdr,true);
		
		//set array of column widths and write sheet
		$writer->setColWidths('Ex3',array(7,12,15,40));
		$writer->writeSheetRow('Ex3',array_keys($hdr),$hstyle);
		$writer->writeSheet($db,'Ex3',$hdr,true);
		
		// write file
		$writer->writeToStdOut();
		unset($writer);
exit(0);
?>