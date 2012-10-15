<?php
require('php-excel-xml.php');

$e = new ExcelXML();
$e->setHeaders(array(
	'test' => new ExcelHeader('Test'),
	'foo' => new ExcelHeader('Foo Col', 'Number'),
	'bar' => new ExcelHeader('Column Three'),
));
$e->addRow(array(
	'test' => new ExcelCell('This is my test'),
	'foo' => new ExcelCell('12'), // Note we're missing a third column here
));
$e->addRow(array(
	'test' => new ExcelCell('Row 2', 'http://example.com'),
	'bar' => new ExcelCell('Hi!'), // Note we skipped a column here
));
$e->addRow(array(
	'test' => new ExcelCell('Row 3'),
	'foo' => new ExcelCell(14, null, 'Not really sure on this one...'),
	'bar' => new ExcelCell('Howdy!'),
));

$filename = $e->out('ooxml');
rename($filename, './test.zip');
copy('test.zip', 'test.xlsx');