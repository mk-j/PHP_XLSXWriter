<?php

namespace Test;


use PHPUnit\Framework\TestCase;

class XlsxWriterTest extends TestCase {

	public function testFileProperties() {
		$expected_author = "John Doe";
		$writer = new \XLSXWriter();
		$writer->setAuthor($expected_author);

		$xlsx_properties = $writer->getFileProperties();

		$this->assertEquals($expected_author, $xlsx_properties["author"]);
	}

}
