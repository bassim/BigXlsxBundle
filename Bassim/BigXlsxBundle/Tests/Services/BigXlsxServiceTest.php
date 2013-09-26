<?php
namespace Bassim\BigXlsxBundle\Tests\Services;

use Bassim\BigXlsxBundle\Services\BenchmarkService;
use Bassim\BigXlsxBundle\Services\BigXlsxService;

class BigXlsxServiceTest extends \PHPUnit_Framework_TestCase
{

	private $_rowCount = 100;

	public function testService()
	{

		/** @var $service BigXlsxService */
		$service = new BigXlsxService();//get('bassim_big_xlsx.service');

		$data[] = array("id","name");
		for ($i=0;$i<$this->_rowCount;$i++) {
			$data[] = array("1_a_".$i, "1_b_".$i, "1_c_".$i);
		}

		$service->addSheet(0, "test Sheet_0", $data);
		$data[] =  array("id2","name2");
		for ($i=0;$i<$this->_rowCount;$i++) {
			$data[] = array("2_a_".$i, "2_b_".$i);
		}

		$service->addSheet(1, "test Sheet_1", $data);
		$file = $service->getFile();

		$reader = new \PHPExcel_Reader_Excel2007();
		$reader->load($file);
		$this->assertEquals(2, count($reader->listWorksheetNames($file)));


	}

	public function testServiceAddCustomSheet()
	{
		/** @var $service BigXlsxService */
		$service = new BigXlsxService();//get('bassim_big_xlsx.service');

		$data[] = array("id","name");
		for ($i=0;$i<1;$i++) {
			$data[] = array("1_a_".$i, "1_b_".$i, "1_c_".$i);
		}

		$service->addSheet(0, "test Sheet_0", $data);
		$data[] =  array("id2","name2");
		for ($i=0;$i<1;$i++) {
			$data[] = array("2_a_".$i, "2_b_".$i);
		}

		$service->addSheet(1, "test Sheet_1", $data);

		$objPHPExcel = $service->getPHPExcel();

		//add third custom sheet
		$objPHPExcel->createSheet(2);
		$objPHPExcel->setActiveSheetIndex(2);
		$objPHPExcel->getActiveSheet()->setTitle("test");

		$file = $service->getFile();

		$objPHPExcel2 = new \PHPExcel_Reader_Excel2007();
		$objPHPExcel2->load($file);

		$this->assertEquals(3, count($objPHPExcel2->listWorksheetNames($file)));

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

		$reader = new \PHPExcel_Reader_Excel2007();
		$reader->load($file);
		$this->assertEquals(2, count($reader->listWorksheetNames($file)));



	}



}