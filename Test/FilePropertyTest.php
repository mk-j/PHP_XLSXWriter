<?php

namespace Test;


use PHPUnit\Framework\TestCase;

class FilePropertyTest extends TestCase {

	public function testAuthor() {
		$expected_author = "John Doe";
		$writer = new \XLSXWriter();
		$writer->setAuthor($expected_author);

		$xlsx_properties = $writer->getFileProperties();

		$this->assertEquals($expected_author, $xlsx_properties["author"]);
	}

	public function testTitle() {
		$expected_title = "My Spreadsheet";
		$writer = new \XLSXWriter();
		$writer->setTitle($expected_title);

		$xlsx_properties = $writer->getFileProperties();

		$this->assertEquals($expected_title, $xlsx_properties["title"]);
	}

	public function testSubject() {
		$expected_subject = "My Spreadsheet is Wonderful";
		$writer = new \XLSXWriter();
		$writer->setSubject($expected_subject);

		$xlsx_properties = $writer->getFileProperties();

		$this->assertEquals($expected_subject, $xlsx_properties["subject"]);
	}

	public function testCompany() {
		$expected_company = "EBANX";
		$writer = new \XLSXWriter();
		$writer->setCompany($expected_company);

		$xlsx_properties = $writer->getFileProperties();

		$this->assertEquals($expected_company, $xlsx_properties["company"]);
	}

	public function testKeywords() {
		$expected_keywords = ["spreadsheet", "php", "EBANX"];
		$writer = new \XLSXWriter();
		$writer->setKeywords($expected_keywords);

		$xlsx_properties = $writer->getFileProperties();

		$this->assertEquals($expected_keywords, $xlsx_properties["keywords"]);
	}

}
