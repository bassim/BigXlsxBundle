<?php
namespace Bassim\BigXlsxBundle\Tests\Services;

use Bassim\BigXlsxBundle\Services\BenchmarkService;
use Bassim\BigXlsxBundle\Services\BigXlsxService;
use Symfony\Bundle\FrameworkBundle\Test\WebTestCase;

class BigXlsxServiceTest extends \PHPUnit_Framework_TestCase
{

	private $_rowCount = 100;

	public function testService()
	{

		/** @var $service BigXlsxService */
		$service = new BigXlsxService();//get('bassim_big_xlsx.service');

		$data = array();
		for ($i=0;$i<$this->_rowCount;$i++) {
			$data[] = array("1_a_".$i, "1_b_".$i, "1_c_".$i);
		}

		$service->addSheet(0, "test Sheet_0", array("id","name"), $data);
		$data = array();
		for ($i=0;$i<$this->_rowCount;$i++) {
			$data[] = array("2_a_".$i, "2_b_".$i);
		}

		$service->addSheet(1, "test Sheet_1", array("id2","name2"), $data);
		$file = $service->get();

		//die(var_dump($file));




		//todo check whether file is valid xlsx


	}

	public function testBenchmarkService()
	{
		/** @var $service BigXlsxService */
		$service = new BenchmarkService();//->get('bassim_big_xlsx.benchmark.service');
		$service->create();

		$data = array();
		for ($i=0;$i<$this->_rowCount;$i++) {
			$data[] = array("1_a_".$i, "1_b_".$i, "1_c_".$i);
		}

		$service->addSheet(0, "test Sheet_0", array("id","name"), $data);
		$data = array();
		for ($i=0;$i<$this->_rowCount;$i++) {
			$data[] = array("2_a_".$i, "2_b_".$i);
		}

		$service->addSheet(1, "test Sheet_1", array("id2","name2"), $data);
		$file = $service->get();




	}



}