<?php
/**
 * Csv Reader
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace Asan\PHPExcel\Reader;

use Asan\PHPExcel\Exception\ReaderException;

class Csv extends BaseReader {
    /**
     * File handle
     *
     * @var resource
     */
    protected $fileHandle;

    /**
     * File read start
     *
     * @var int
     */
    protected $start = 0;

    /**
     * Input encoding
     *
     * @var string
     */
    protected $inputEncoding;

    /**
     * Delimiter
     *
     * @var string
     */
    protected $delimiter;

    /**
     * Enclosure
     *
     * @var string
     */
    protected $enclosure = '"';

    /**
     * Ignore empty row
     *
     * @var bool
     */
    protected $ignoreEmpty = false;

    /**
     * Loads Excel from file
     *
     * @param string $file
     *
     * @throws ReaderException
     * @return $this
     */
    public function load($file) {
        $lineEnding = ini_get('auto_detect_line_endings');
        ini_set('auto_detect_line_endings', true);

        // Open file
        $this->openFile($file);

        $this->autoDetection();

        $this->generator = $this->makeGenerator();

        ini_set('auto_detect_line_endings', $lineEnding);

        return $this;
    }

    /**
     * Count elements of the selected sheet
     *
     * @return int
     */
    public function count() {
        if ($this->count === null) {
            $position = ftell($this->fileHandle);
            $this->count = iterator_count($this->makeGenerator(true));
            fseek($this->fileHandle, $position);
        }

        return $this->count;
    }

    /**
     * Make the generator
     *
     * @param bool $calculate
     * @return \Generator
     */
    protected function makeGenerator($calculate = false) {
        fseek($this->fileHandle, $this->start);

        $finish = 0;
        while (($row = fgetcsv($this->fileHandle, 0, $this->delimiter, $this->enclosure)) !== false) {
            if ($this->ignoreEmpty && (empty($row) || trim(implode('', $row)) === '')) {
                continue;
            }

            if ($calculate) {
                yield;
                continue;
            }

            if ($this->rowLimit > 0 && ++$finish > $this->rowLimit) {
                break;
            }

            if ($this->columnLimit > 0) {
                $row = array_slice($row, 0, $this->columnLimit);
            }

            foreach ($row as &$value) {
                if ($value != '') {
                    if (is_numeric($value)) {
                        $value = (float)$value;
                    }

                    // Convert encoding if necessary
                    if ($this->inputEncoding !== 'UTF-8') {
                        $value = mb_convert_encoding($value, 'UTF-8', $this->inputEncoding);
                    }
                }
            }

            unset($value);

            yield $row;
        }
    }

    /**
     * Detect the file delimiter and encoding
     */
    protected function autoDetection() {
        if (($this->delimiter !== null && $this->inputEncoding !== null)
            || ($line = fgets($this->fileHandle)) === false) {

            return;
        }

        if ($this->delimiter === null) {
            $this->delimiter = ',';

            if ((strlen(trim($line, "\r\n")) == 5) && (stripos($line, 'sep=') === 0)) {
                $this->delimiter = substr($line, 4, 1);
            }
        }

        if ($this->inputEncoding === null) {
            $this->inputEncoding = 'UTF-8';

            if (($bom = substr($line, 0, 4)) == "\xFF\xFE\x00\x00" || $bom == "\x00\x00\xFE\xFF") {
                $this->start = 4;
                $this->inputEncoding = 'UTF-32';
            } elseif (($bom = substr($line, 0, 2)) == "\xFF\xFE" || $bom == "\xFE\xFF") {
                $this->start = 2;
                $this->inputEncoding = 'UTF-16';
            } elseif (($bom = substr($line, 0, 3)) == "\xEF\xBB\xBF") {
                $this->start = 3;
            }

            if (!$this->start) {
                $encoding = mb_detect_encoding($line, 'ASCII, UTF-8, GB2312, GBK');

                if ($encoding) {
                    if ($encoding == 'EUC-CN') {
                        $encoding = 'GB2312';
                    } elseif ($encoding == 'CP936') {
                        $encoding = 'GBK';
                    }

                    $this->inputEncoding = $encoding;
                }
            }
        }

        fseek($this->fileHandle, $this->start);
    }

    /**
     * Ignore empty row
     *
     * @param bool $ignoreEmpty
     *
     * @return $this
     */
    public function ignoreEmptyRow($ignoreEmpty = false) {
        $this->ignoreEmpty = $ignoreEmpty;

        return $this;
    }

    /**
     * Set input encoding
     *
     * @param string $encoding
     * @return $this
     */
    public function setInputEncoding($encoding = 'UTF-8') {
        $this->inputEncoding = $encoding;

        return $this;
    }

    /**
     * Get input encoding
     *
     * @return string
     */
    public function getInputEncoding() {
        return $this->inputEncoding;
    }

    /**
     * Set delimiter
     *
     * @param string $delimiter  Delimiter, defaults to ,
     * @return $this
     */
    public function setDelimiter($delimiter = ',') {
        $this->delimiter = $delimiter;

        return $this;
    }

    /**
     * Get delimiter
     *
     * @return string
     */
    public function getDelimiter() {
        return $this->delimiter;
    }

    /**
     * Set enclosure
     *
     * @param string $enclosure  Enclosure, defaults to "
     * @return $this
     */
    public function setEnclosure($enclosure = '"') {
        if ($enclosure == '') {
            $enclosure = '"';
        }

        $this->enclosure = $enclosure;

        return $this;
    }

    /**
     * Get enclosure
     *
     * @return string
     */
    public function getEnclosure() {
        return $this->enclosure;
    }

    /**
     * Can the current Reader read the file?
     *
     * @param string $file
     *
     * @return bool
     */
    public function canRead($file) {
        try {
            $this->openFile($file);
        } catch (\Exception $e) {
            return false;
        }

        fclose($this->fileHandle);

        return true;
    }

    /**
     * Open file for reading
     *
     * @param string $file
     *
     * @throws ReaderException
     */
    protected function openFile($file) {
        // Check if file exists
        if (!file_exists($file) || !is_readable($file)) {
            throw new ReaderException("Could not open file [$file] for reading! File does not exist.");
        }

        // Open file
        $this->fileHandle = fopen($file, 'r');
        if ($this->fileHandle === false) {
            throw new ReaderException("Could not open file [$file] for reading.");
        }
    }

    /**
     * Close file and release generator
     */
    public function __destruct() {
        if ($this->fileHandle) {
            fclose($this->fileHandle);
        }

        $this->generator = null;
    }
}
