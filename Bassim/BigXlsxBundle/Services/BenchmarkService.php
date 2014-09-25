<?php
namespace Bassim\BigXlsxBundle\Services;

class BenchmarkService
{

    /** @var $objPHPExcel \PHPExcel */
    private $objPHPExcel;
    private $columnName;
    private $sheets = array();

    public function __construct()
    {
    }


    public function create()
    {
        $this->columnName = null;

        $this->objPHPExcel = new \PHPExcel();

    }

    public function addSheet($sheetNumber, $name, $columns, $data)
    {
        if ($sheetNumber > 0) {
            $this->objPHPExcel->createSheet($sheetNumber);
        }

        $this->objPHPExcel->setActiveSheetIndex($sheetNumber);
        $this->objPHPExcel->getActiveSheet()->setTitle($name);

        $this->sheets[$sheetNumber] = array_merge(array($columns), $data);

        //set headers
        foreach ($columns as $row) {
            if (is_null($this->columnName)) {
                $this->columnName = 'a';
            }
            $cellName = strtoupper($this->columnName++ . "1");
            $this->objPHPExcel->getActiveSheet()->setCellValue($cellName, $row);
        }

        $this->columnName = null;

        //set data
        foreach ($data as $rowIndex => $row) {
            foreach ($row as $cell) {
                if (is_null($this->columnName)) {
                    $this->columnName = 'a';
                }
                $cellName = strtoupper($this->columnName++ . "" . ($rowIndex + 2));
                $this->objPHPExcel->getActiveSheet()->setCellValue($cellName, $cell);
            }
            $this->columnName = null;
        }
    }

    public function get()
    {
        // Save Excel 2007 file
        $objWriter = new \PHPExcel_Writer_Excel2007($this->objPHPExcel);
        $date = new \DateTime();
        $date = $date->format("Y-m-d_h-i");
        $filename = "/tmp/report_export_" . $date . ".xlsx";


        $objWriter->save($filename);

        return $filename;
    }
}
