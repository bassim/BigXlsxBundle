<?php

namespace Bassim\BigXlsxBundle\Services;


use Bassim\BigXlsxBundle\Entity\SharedStringXml;
use Bassim\BigXlsxBundle\Entity\SheetXml;
use Symfony\Bridge\Monolog\Logger;

class BenchmarkService
{

	/** @var $objPHPExcel \PHPExcel */
	private $_objPHPExcel;
	private $_columnName;
	private $_sheets = array();

	public function __construct()
	{
	}


	public function create()
	{
		$this->_columnName=null;

		$this->_objPHPExcel = new \PHPExcel();

	}

	public function addSheet($sheetNumber, $name, $columns, $data)
	{
		if ($sheetNumber>0) {
			$this->_objPHPExcel->createSheet($sheetNumber);
		}

		$this->_objPHPExcel->setActiveSheetIndex($sheetNumber);
		$this->_objPHPExcel->getActiveSheet()->setTitle($name);

		$this->_sheets[$sheetNumber] = array_merge(array($columns), $data);


		//set headers
		$columnNumber = 1;
		foreach ($columns as $rowIndex=>$row) {
			if (is_null($this->_columnName)) {
				$this->_columnName = 'a';
			}
			$cellName = strtoupper($this->_columnName++."1");
			$this->_objPHPExcel->getActiveSheet()->setCellValue($cellName, $row);
		}

		$this->_columnName = null;


		//set data
		$columnNumber = 2;
		foreach ($data as $rowIndex=>$row) {
			foreach ($row as $columnIndex=>$cell) {
				if (is_null($this->_columnName)) {
					$this->_columnName = 'a';
				}
				$cellName = strtoupper($this->_columnName++."".($rowIndex+2));
				$this->_objPHPExcel->getActiveSheet()->setCellValue($cellName, $cell);
			}
			$this->_columnName = null;
		}


	}

	public function get()
	{
		// Save Excel 2007 file
		$objWriter = new \PHPExcel_Writer_Excel2007($this->_objPHPExcel);
		$date = new \DateTime();
		$date = $date->format("Y-m-d_h-i");
		$filename = "/tmp/report_export_".$date.".xlsx";


		$objWriter->save($filename);

		return $filename;


	}

}