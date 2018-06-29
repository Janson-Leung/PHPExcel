<?php
/**
 * Xls Test
 *
 * @author Janson
 * @create 2017-11-28
 */
require __DIR__ . '/../autoload.php';

$start = microtime(true);
$memory = memory_get_usage();

$reader = Asan\PHPExcel\Excel::load('files/01.xls', function(Asan\PHPExcel\Reader\Xls $reader) {
    //$reader->setRowLimit(5);
    $reader->setColumnLimit(10);

    //$reader->setSheetIndex(1);
});

foreach ($reader as $row) {
    var_dump($row);
}

$reader->seek(50);

//$reader->seek(5);
$count = $reader->count();
$current = $reader->current();

$sheets = $reader->sheets();

$time = microtime(true) - $start;
$use = memory_get_usage() - $memory;

var_dump($current, $count, $sheets, $time, $use/1024/1024);
