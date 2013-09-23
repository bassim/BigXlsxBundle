<?php

namespace Bassim\BigXlsxBundle\Services;


use Bassim\BigXlsxBundle\Entity\SharedStringXml;
use Bassim\BigXlsxBundle\Entity\SheetXml;


class BigXlsxService
{

	/** @var $objPHPExcel \PHPExcel */
	public $_objPHPExcel;
	public $_columnName;

	private $_sheets = array();

	public function __construct()
	{
		$this->_columnName=null;
		$this->_objPHPExcel = new \PHPExcel();
	}


	public function addSheet($sheetNumber, $name, $data)
	{
		if ($sheetNumber>0) {
			$this->_objPHPExcel->createSheet($sheetNumber);
		}

		$this->_objPHPExcel->setActiveSheetIndex($sheetNumber);
		$this->_objPHPExcel->getActiveSheet()->setTitle($name);

		$this->_sheets[$sheetNumber] = $data;
	}

	public function get()
	{
		// Save Excel 2007 file
		$objWriter = new \PHPExcel_Writer_Excel2007($this->_objPHPExcel);
		$date = new \DateTime();
		$date = $date->format("Y-m-d_h-i");
		$filename = "/tmp/report_export_".$date.".xlsx";
		$objWriter->save($filename);

		$zipArchive = new \ZipArchive();
		$zipArchive->open($filename);

		$sharedStringXml = new SharedStringXml();
		foreach ($this->_sheets as $key=>$sheet) {
			$sheetXml = new SheetXml($sharedStringXml);
			foreach ($sheet as $rowIndex => $row) {
				$sheetXml->addRow($row);
			}
			$zipArchive->addFile($sheetXml->getFile(), 'xl/worksheets/sheet'.($key+1).'.xml');
		}


		$zipArchive->addFile($sharedStringXml->getFile(), 'xl/sharedStrings.xml');
		$zipArchive->close();
		return $filename;


	}

	private function _convert($size)
	{
		$unit=array('b','kb','mb','gb','tb','pb');
		return @round($size/pow(1024,($i=floor(log($size,1024)))),2).' '.$unit[$i];
	}



}