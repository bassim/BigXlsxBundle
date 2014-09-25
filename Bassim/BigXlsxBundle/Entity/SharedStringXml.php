<?php
namespace Bassim\BigXlsxBundle\Entity;

class SharedStringXml
{
    private $string;
    private $lineCount;
    private $position = -1;
    private $strings = array();
    private $filePointer;
    private $sharedStringsFile;

    public function __construct()
    {
        $this->sharedStringsFile = realpath(sys_get_temp_dir()) . "/" . uniqid("sharedStringXml");
        $this->filePointer = fopen($this->sharedStringsFile, 'w');
    }


    public function addString($string)
    {
        $pos = false; //array_search($string, $this->_strings);
        if ($pos === false) {
            $this->position++;
            $this->strings[] = $string;

            $this->write("<si><t>" . $string . "</t></si>");
            return $this->position;
        }
        return $pos;
    }

    public function getFile()
    {
        $this->flush();
        $this->prependHeader();
        $this->appendFooter();
        return $this->sharedStringsFile;
    }

    private function prependHeader()
    {
        $string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="' . count($this->strings) . '">
';

        $context = stream_context_create();
        $filePointer = fopen($this->sharedStringsFile, 'r', 1, $context);
        $tmpName = md5($string);
        file_put_contents(realpath(sys_get_temp_dir()) . "/" . $tmpName, $string);
        file_put_contents(realpath(sys_get_temp_dir()) . "/" . $tmpName, $filePointer, FILE_APPEND);
        fclose($filePointer);
        unlink($this->sharedStringsFile);
        rename(realpath(sys_get_temp_dir()) . "/" . $tmpName, $this->sharedStringsFile);
    }

    private function appendFooter()
    {
        $context = stream_context_create();
        $filePointer = fopen($this->sharedStringsFile, 'a', 1, $context);
        fwrite($filePointer, '</sst>');
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
