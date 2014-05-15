<?php

namespace Bassim\BigXlsxBundle\Entity;


class SharedStringXml
{
	private $_string;
	private $_lineCount;
	private $_position = -1;
	private $_strings = array();
	private $_fp;
	private $_sharedStringsFile;

	public function __construct()
	{
		$this->_sharedStringsFile = realpath(sys_get_temp_dir())."/".uniqid("sharedStringXml");
		$this->_fp = fopen($this->_sharedStringsFile, 'w');
	}


	public function addString($string)
	{
		$pos = false;//array_search($string, $this->_strings);
		if ($pos === false) {
			$this->_position++;
			$this->_strings[] = $string;

			$this->_write("<si><t>".$string."</t></si>");
			return $this->_position;
		}
		return $pos;
	}

	public function getFile()
	{
		$this->_flush();
		$this->_prependHeader();
		$this->_appendFooter();
		return $this->_sharedStringsFile;
	}

	private function _prependHeader()
	{
		$string =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="'.count($this->_strings).'">
';

		$context = stream_context_create();
		$fp = fopen($this->_sharedStringsFile, 'r', 1, $context);
		$tmpname = md5($string);
		file_put_contents(realpath(sys_get_temp_dir())."/".$tmpname, $string);
		file_put_contents(realpath(sys_get_temp_dir())."/".$tmpname, $fp, FILE_APPEND);
		fclose($fp);
		unlink($this->_sharedStringsFile);
		rename(realpath(sys_get_temp_dir())."/".$tmpname, $this->_sharedStringsFile);

	}

	private function _appendFooter()
	{
		$context = stream_context_create();
		$fp = fopen($this->_sharedStringsFile, 'a', 1, $context);
		fwrite($fp, '</sst>');
		fclose($fp);
	}

	private function _write($string)
	{
		$this->_lineCount++;
		$this->_string .=$string;

	}

	private function _flush()
	{
		fwrite($this->_fp,$this->_string);
		fclose($this->_fp);
	}


}