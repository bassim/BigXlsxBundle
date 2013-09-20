<?php
namespace Bassim\BigXlsxBundle\Tests\Entity;

use Bassim\BigXlsxBundle\Entity\SharedStringXml;

class SharedStringXmlTest extends \PHPUnit_Framework_TestCase
{

	public function testWrite()
	{
		$sharedStringXml = new SharedStringXml();

		$pos = $sharedStringXml->addString("1");
		$this->assertEquals(0, $pos);

		$pos = $sharedStringXml->addString("1");
		$this->assertEquals(1 , $pos);

		$file = $sharedStringXml->getFile();
		$xml=simplexml_load_file($file);
		$this->assertEquals($xml->si[0]->t, "1");
	}
}