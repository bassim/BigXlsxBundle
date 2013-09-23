BigXlsxBundle
=============

Symfony2 Bundle for generating large multi-sheeted xlsx files with low memory usage.

This Bundle basically acts as a replacement for creating csv files with large datasets.
This, because the CSV format can be troublesome when it comes to difference in default cell separators en line-endings on specific operating systems.
Also, obviously, CSV cannot handle multiple sheets.

This bundle uses the 'codeplex/phpexcel' bundle but tries to get rid of the enormous execution time and memory consumption when one wants to store large datasets in a xlsx.

Installation
------------

add this to your composer.json:

		"bassim/big-xlsx-bundle": "dev-master"

add this to your AppKernel.php

		new Bassim\BigXlsxBundle\BassimBigXlsxBundle(),

Usage
-----

		/** @var $service BigXlsxService */
		$service = $container->get('bassim_big_xlsx.service');

		$data[] = array("id","name");
		for ($i=0;$i<1;$i++) {
			$data[] = array($i, "name_".$i);
		}

		$service->addSheet(0, "test Sheet_0", $data);
		$file = $service->get();

  
