<?php

function classLoader($class) {
    $path = str_replace(
        ['\\', 'Asan' . DIRECTORY_SEPARATOR . 'PHPExcel' . DIRECTORY_SEPARATOR], [DIRECTORY_SEPARATOR, ''], $class
    );

    $file = __DIR__ . DIRECTORY_SEPARATOR . 'src' . DIRECTORY_SEPARATOR . $path . '.php';

    if (file_exists($file)) {
        require_once $file;
    }
}

spl_autoload_register('classLoader');
