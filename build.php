<?php

unlink('test.xlsx');
echo shell_exec('cd example && zip -r ../test.xlsx *');

$z = new ZipArchive();
$z->open('test.xlsx');

//$z->deleteName('xl/styles.xml'); // Makes everything plain, but still works
$z->close();

echo shell_exec('unzip -l test.xlsx');
