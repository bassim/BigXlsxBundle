<?php
namespace Bassim\BigXlsxBundle\Entity;

class SheetXml
{
    /**
     * @var SharedStringXml
     */
    private $sharedStringXml;
    private $rowIndex = 1;
    private $sheetFile;
    private $filePointer;
    private $string;
    private $lineCount;

    public function __construct(SharedStringXml $sharedStringXml)
    {
        $this->sheetFile = realpath(sys_get_temp_dir()) . "/" . uniqid("sheetXml");
        $this->filePointer = fopen($this->sheetFile, 'w');
        $this->sharedStringXml = $sharedStringXml;
    }


    public function addRow($row)
    {
        $rowString = '<row r="' . $this->rowIndex . '" spans="1:' . count($row) . '">';

        $letter = "A";
        foreach ($row as $value) {
            $position = $this->sharedStringXml->addString($value);
            $rowString .= '<c r="' . $letter++ . '' . $this->rowIndex . '" t="s"><v>' . $position . '</v></c>';
        }

        $rowString .= '</row>';
        $this->write($rowString);
        $this->rowIndex++;
    }

    public function getFile()
    {
        $this->flush();
        $this->prependHeader();
        $this->appendFooter();
        return $this->sheetFile;
    }

    private function prependHeader()
    {
        $string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr><dimension ref="A1:B101"/><sheetViews><sheetView tabSelected="0" workbookViewId="0" showGridLines="true" showRowColHeaders="1"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="14.4" outlineLevelRow="0" outlineLevelCol="0"/>
    <sheetData>';

        $context = stream_context_create();
        $filePointer = fopen($this->sheetFile, 'r', 1, $context);
        $tmpName = md5($string);
        file_put_contents(realpath(sys_get_temp_dir()) . "/" . $tmpName, $string);
        file_put_contents(realpath(sys_get_temp_dir()) . "/" . $tmpName, $filePointer, FILE_APPEND);
        fclose($filePointer);
        unlink($this->sheetFile);
        rename(realpath(sys_get_temp_dir()) . "/" . $tmpName, $this->sheetFile);

    }

    private function appendFooter()
    {

        $context = stream_context_create();
        $filePointer = fopen($this->sheetFile, 'a', 1, $context);
        fwrite($filePointer, '</sheetData>
<sheetProtection sheet="false" objects="false" scenarios="false" formatCells="false" formatColumns="false" formatRows="false" insertColumns="false" insertRows="false" insertHyperlinks="false" deleteColumns="false" deleteRows="false" selectLockedCells="false" sort="false" autoFilter="false" pivotTables="false" selectUnlockedCells="false"/><printOptions gridLines="false" gridLinesSet="true"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="1" orientation="default" scale="100" fitToHeight="1" fitToWidth="1"/><headerFooter differentOddEven="false" differentFirst="false" scaleWithDoc="true" alignWithMargins="true"><oddHeader></oddHeader><oddFooter></oddFooter><evenHeader></evenHeader><evenFooter></evenFooter><firstHeader></firstHeader><firstFooter></firstFooter></headerFooter></worksheet>');
        fclose($filePointer);
    }

    private function write($string)
    {
        $this->lineCount++;
        $this->string .= $string;
    }

    private function flush()
    {
        fwrite($this->filePointer, $this->string);
        fclose($this->filePointer);
    }
}
