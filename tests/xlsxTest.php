<?php
/**
 * xlsx Test
 *
 * @author Janson
 * @create 2017-11-28
 */
require __DIR__ . '/../autoload.php';

$start = microtime(true);
$memory = memory_get_usage();

$reader = Asan\PHPExcel\Excel::load('files/01.xlsx', function(Asan\PHPExcel\Reader\Xlsx $reader) {
    $reader->setRowLimit(10);
    $reader->setColumnLimit(10);

    $reader->ignoreEmptyRow(true);

    //$reader->setSheetIndex(0);
});

foreach ($reader as $row) {
    var_dump($row);
}

//$reader->seek(50);

$count = $reader->count();
$reader->seek(2);
$current = $reader->current();

$sheets = $reader->sheets();

$time = microtime(true) - $start;
$use = memory_get_usage() - $memory;

var_dump($current, $count, $sheets, $time, $use/1024/1024);