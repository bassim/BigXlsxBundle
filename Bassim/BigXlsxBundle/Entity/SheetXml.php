<?php

namespace Bassim\BigXlsxBundle\Entity;


class SheetXml
{
	/**
	 * @var SharedStringXml
	 */
	private $_sharedStringXml;
	private $_rowIndex = 1;
	private $_sheetFile;
	private $_fp;
	private $_string;
	private $_lineCount;

	public function __construct(SharedStringXml $sharedStringXml)
	{
		$this->_sheetFile = "/tmp/".uniqid();
		$this->_fp = fopen($this->_sheetFile, 'w');
		$this->_sharedStringXml = $sharedStringXml;
	}


	public function addRow($row)
	{
		$rowString = '<row r="'.$this->_rowIndex.'" spans="1:'.count($row).'">';

		$letter = "A";
		foreach ($row as $value) {
			$position = $this->_sharedStringXml->addString($value);
			$rowString .= '<c r="'.$letter++.''.$this->_rowIndex.'" t="s"><v>'.$position.'</v></c>';
		}

		$rowString .= '</row>';
		$this->_write($rowString);
		$this->_rowIndex++;
	}

	public function getFile()
	{
		$this->_flush();
		$this->_prependHeader();
		$this->_appendFooter();
		return $this->_sheetFile;
	}

	private function _prependHeader()
	{
		$string =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr><dimension ref="A1:B101"/><sheetViews><sheetView tabSelected="0" workbookViewId="0" showGridLines="true" showRowColHeaders="1"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="14.4" outlineLevelRow="0" outlineLevelCol="0"/>
    <sheetData>';

		$context = stream_context_create();
		$fp = fopen($this->_sheetFile, 'r', 1, $context);
		$tmpname = md5($string);
		file_put_contents($tmpname, $string);
		file_put_contents($tmpname, $fp, FILE_APPEND);
		fclose($fp);
		unlink($this->_sheetFile);
		rename($tmpname, $this->_sheetFile);

	}

	private function _appendFooter()
	{

		$context = stream_context_create();
		$fp = fopen($this->_sheetFile, 'a', 1, $context);
		fwrite($fp, '</sheetData>
<sheetProtection sheet="false" objects="false" scenarios="false" formatCells="false" formatColumns="false" formatRows="false" insertColumns="false" insertRows="false" insertHyperlinks="false" deleteColumns="false" deleteRows="false" selectLockedCells="false" sort="false" autoFilter="false" pivotTables="false" selectUnlockedCells="false"/><printOptions gridLines="false" gridLinesSet="true"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="1" orientation="default" scale="100" fitToHeight="1" fitToWidth="1"/><headerFooter differentOddEven="false" differentFirst="false" scaleWithDoc="true" alignWithMargins="true"><oddHeader></oddHeader><oddFooter></oddFooter><evenHeader></evenHeader><evenFooter></evenFooter><firstHeader></firstHeader><firstFooter></firstFooter></headerFooter></worksheet>');
		fclose($fp);
	}

	private function _write($string)
	{
		$this->_lineCount++;
		$this->_string .=$string;

//		if ($this->_lineCount>4000) {
//
//			fwrite($this->_fp,$this->_string);
//			$this->_string = null;
//			$this->_lineCount = 0;
//
//		}

	}

	private function _flush()
	{
		//fwrite($this->_fp,implode('', $this->_string));
		fwrite($this->_fp,$this->_string);
		fclose($this->_fp);
	}

}