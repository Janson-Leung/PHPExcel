<?php
/**
 * PHP Excel
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace Asan\PHPExcel;

use Asan\PHPExcel\Exception\ReaderException;

class Excel {
    /**
     * Load a file
     *
     * @param string $file
     * @param callback|null $callback
     * @param string|null $encoding
     * @param string $ext
     * @param string $logPath
     *
     * @throws ReaderException
     * @return \Asan\PHPExcel\Reader\BaseReader
     */
    public static function load($file, $callback = null, $encoding = null, $ext = '', $logPath = '') {
        set_error_handler(function($errorNo, $errorMsg, $errorFile, $errorLine) use ($logPath) {
            if ($logPath) {
                if (!file_exists($logPath)) {
                    mkdir($logPath, 0755, true);
                }

                $content = sprintf(
                    "%s\t%s.%s\t%s\t%s", date("Y-m-d H:i:s"), self::class, 'ERROR',
                    "[$errorNo]$errorMsg in $errorFile:$errorLine", PHP_EOL
                );

                file_put_contents("$logPath/excel-" . date('Y-m-d'). '.log', $content, FILE_APPEND);
            }
        }, E_ALL ^ E_ERROR);

        $ext = $ext ?: strtolower(pathinfo($file, PATHINFO_EXTENSION));

        $format = self::getFormatByExtension($ext);

        if (empty($format)) {
            throw new ReaderException("Could not identify file format for file [$file] with extension [$ext]");
        }

        $class = __NAMESPACE__ . '\\Reader\\' . $format;
        $reader = new $class;

        if ($callback) {
            if ($callback instanceof \Closure) {
                // Do the callback
                call_user_func($callback, $reader);
            } elseif (is_string($callback)) {
                // Set the encoding
                $encoding = $callback;
            }
        }

        if ($encoding && method_exists($reader, 'setInputEncoding')) {
            $reader->setInputEncoding($encoding);
        }

        return $reader->load($file);
    }

    /**
     * Identify file format
     *
     * @param string $ext
     * @return string
     */
    protected static function getFormatByExtension($ext) {
        $formart = '';

        switch ($ext) {
            /*
            |--------------------------------------------------------------------------
            | Excel 2007
            |--------------------------------------------------------------------------
            */
            case 'xlsx':
            case 'xlsm':
            case 'xltx':
            case 'xltm':
                $formart = 'Xlsx';
                break;

            /*
            |--------------------------------------------------------------------------
            | Excel5
            |--------------------------------------------------------------------------
            */
            case 'xls':
            case 'xlt':
                $formart = 'Xls';
                break;

            /*
            |--------------------------------------------------------------------------
            | CSV
            |--------------------------------------------------------------------------
            */
            case 'csv':
            case 'txt':
                $formart = 'Csv';
                break;
        }

        return $formart;
    }
}
