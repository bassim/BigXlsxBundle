<?php

namespace Bassim\BigXlsxBundle\Tests\Entity;


use Bassim\BigXlsxBundle\Entity\SharedStringXml;
use Bassim\BigXlsxBundle\Entity\SheetXml;

class SheetXmlTest extends \PHPUnit_Framework_TestCase
{

	public function testWrite()
	{
		$data = array(
			array("1","2","3")
		);

		$sharedStringXml = new SharedStringXml();
		$sheetXml = new SheetXml($sharedStringXml);
		$sheetXml->addRow($data[0]);


		$file = $sheetXml->getFile();
		$xml=simplexml_load_file($file);

		$this->assertEquals((string)$xml->sheetData->row[0]->c[0]->v, 0);


	}

}