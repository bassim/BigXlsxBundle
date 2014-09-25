<?php

namespace Bassim\BigXlsxBundle\Services;

/**
 * BigXlsxBundle
 *
 * Copyright (c) 2013 Bas Simons
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   BigXlsxBundle
 * @package    Services
 * @copyright  Copyright (c) 2013 Bas Simons (https://github.com/bassim)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    0.0.1, 2013-10-26
 */

use Bassim\BigXlsxBundle\Entity\SharedStringXml;
use Bassim\BigXlsxBundle\Entity\SheetXml;

class BigXlsxService
{

    /** @var $objPHPExcel \PHPExcel */
    private $objPHPExcel;

    /**
     * @var array
     */
    private $sheets = array();

    /**
     * Constructor.
     *
     * Initiate _objPHPExcel
     */
    public function __construct()
    {
        $this->objPHPExcel = new \PHPExcel();
    }

    /**
     * Add a Sheet
     *
     * @param int $sheetNumber The SheetNumber
     * @param string $sheetName The SheetName
     * @param array $sheetData The SheetData
     *
     * @return void
     */
    public function addSheet($sheetNumber, $sheetName, $sheetData)
    {
        if ($sheetNumber > 0) {
            $this->objPHPExcel->createSheet($sheetNumber);
        }

        $this->objPHPExcel->setActiveSheetIndex($sheetNumber);
        $this->objPHPExcel->getActiveSheet()->setTitle($sheetName);

        $this->sheets[$sheetNumber] = $sheetData;
    }

    /**
     * Returns the path to the xlsx file
     *
     * @param null|string $filePath The optional Custom FilePath
     *
     * @return string
     */
    public function getFile($filePath = null)
    {
        // Save Excel 2007 file
        $objWriter = new \PHPExcel_Writer_Excel2007($this->objPHPExcel);

        if (is_null($filePath)) {
            $filePath = realpath(sys_get_temp_dir()) . "/" . uniqid("xlsx") . ".xlsx";
        }

        $objWriter->save($filePath);

        $zipArchive = new \ZipArchive();
        $zipArchive->open($filePath);

        $sharedStringXml = new SharedStringXml();
        foreach ($this->sheets as $key => $sheet) {
            $sheetXml = new SheetXml($sharedStringXml);
            foreach ($sheet as $row) {
                $sheetXml->addRow($row);
            }
            $zipArchive->addFile($sheetXml->getFile(), 'xl/worksheets/sheet' . ($key + 1) . '.xml');
        }

        $zipArchive->addFile($sharedStringXml->getFile(), 'xl/sharedStrings.xml');
        $zipArchive->close();
        return $filePath;
    }

    /**
     * Returns the PHPExcel instance
     *
     * @return \PHPExcel
     */
    public function getPHPExcel()
    {
        return $this->objPHPExcel;
    }
}
