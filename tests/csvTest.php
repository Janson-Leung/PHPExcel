<?php
/**
 * Csv test
 *
 * @author Janson
 * @create 2017-11-25
 */
require __DIR__ . '/../autoload.php';

$start = microtime(true);
$memory = memory_get_usage();

$reader = Asan\PHPExcel\Excel::load('files/02.csv', function(Asan\PHPExcel\Reader\Csv $reader) {
    $reader->setRowLimit(5);
    $reader->setColumnLimit(10);

    $reader->ignoreEmptyRow(true);

    //$reader->setInputEncoding('UTF-8');
    $reader->setDelimiter("\t");
});

foreach ($reader as $row) {
    var_dump($row);
}

$reader->seek(2);

$count = $reader->count();
//$reader->seek(1);
$current = $reader->current();

$time = microtime(true) - $start;
$use = memory_get_usage() - $memory;
var_dump($current, $count, $time, $use/1024/1024);
