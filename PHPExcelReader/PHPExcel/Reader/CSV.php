<?php

class PHPExcel_Reader_CSV implements Iterator, Countable {
    private $_fileHandle = null;

    private $filePath = '';

    private $_inputEncoding = 'UTF-8';

    private $_delimiter = ',';

    private $_enclosure = '"';

    private $_filter = 0;

    /**
     * @param string $filePath
     * @param int $filter filter empty row
     *
     * @throws Exception
     */
    public function __construct($filePath, $filter = 0) {
        if (! file_exists($filePath)) {
            throw new Exception("Could not open " . $filePath . " for reading! File does not exist.");
        }

        $this->filePath = $filePath;
        $this->_filter = $filter;
        ini_set('auto_detect_line_endings', true);

        $this->_fileHandle = fopen($filePath, 'r');
        $this->_detectEncoding();
    }

    /**
     * Move filepointer past any BOM marker
     */
    private function _detectEncoding() {
        $step = $BOMLength = 0;
        while ($step < 3) {
            $BOM = bin2hex(fread($this->_fileHandle, 2 + $step++));

            rewind($this->_fileHandle);

            if ($BOM == 'fffe' || $BOM == 'feff') {
                $BOMLength = 2;
                $this->_delimiter = "\t";
                $this->_inputEncoding = 'UTF-16';
                break;
            } else {
                if ($BOM == 'efbbbf') {
                    $BOMLength = 3;
                    break;
                } else {
                    if ($BOM == '0000feff' || $BOM == 'fffe0000') {
                        $BOMLength = 4;
                        $this->_delimiter = "\t";
                        $this->_inputEncoding = 'UTF-32';
                        break;
                    }
                }
            }
        }

        if (! $BOMLength) {
            $encoding = mb_detect_encoding(fgets($this->_fileHandle, 1024), 'ASCII, UTF-8, GB2312, GBK');
            rewind($this->_fileHandle);
            if ($encoding) {
                if ($encoding == 'EUC-CN') {
                    $this->_inputEncoding = 'GB2312';
                } else {
                    if ($encoding == 'CP936') {
                        $this->_inputEncoding = 'GBK';
                    } else {
                        $this->_inputEncoding = $encoding;
                    }
                }
            }
        }

        if ($this->_inputEncoding != 'UTF-8') {
            stream_filter_register("convert_iconv.*", "convert_iconv_filter");
            stream_filter_append($this->_fileHandle, 'convert_iconv.' . $this->_inputEncoding . '/UTF-8');
        }
    }

    /**
     * Returns information about sheets in the file.
     * @return array
     */
    public function Sheets() {
        return array(0 => basename($this->filePath));
    }

    /**
     * Changes sheet to another.
     *
     * @param int $index
     * @return bool
     */
    public function ChangeSheet($index) {
        if ($index == 0) {
            $this->rewind();

            return true;
        }

        return false;
    }

    /**
     * Rewind the Iterator to the first element.
     */
    public function rewind() {
        rewind($this->_fileHandle);
        $this->currentRow = null;
        $this->index = 0;
    }

    /**
     * Return the current element.
     * @return mixed
     */
    public function current() {
        if ($this->index == 0 && ! isset($this->currentRow)) {
            $this->rewind();
            $this->next();
            $this->index = 0;
        }

        return $this->currentRow;
    }

    /**
     * Move forward to next element.
     */
    public function next() {
        $this->currentRow = array();

        $this->index++;
        while (($row = fgetcsv($this->_fileHandle, 0, $this->_delimiter, $this->_enclosure)) !== false) {
            if (! $this->_filter || array_filter($row, array($this, 'filter'))) {
                $this->currentRow = $row;
                break;
            }
        }

        return $this->currentRow;
    }

    /**
     * Return the identifying key of the current element.
     * @return mixed
     */
    public function key() {
        return $this->index;
    }

    /**
     * Check if there is a current element after calls to rewind() or next().
     * @return bool
     */
    public function valid() {
        if ($this->currentRow || ! feof($this->_fileHandle)) {
            return true;
        } else {
            fclose($this->_fileHandle);

            return false;
        }
    }

    /**
     * return the count of the contained items
     * @return int
     */
    public function count() {
        if (! isset($this->rowCount)) {
            $total = 0;
            rewind($this->_fileHandle);
            while (($row = fgetcsv($this->_fileHandle, 0, $this->_delimiter, $this->_enclosure)) !== false) {
                if (! $this->_filter || array_filter($row, array($this, 'filter'))) {
                    $total++;
                }
            }

            $this->rowCount = $total;
        }

        return $this->rowCount;
    }

    /**
     * filter empty string
     *
     * @param mixed $value
     *
     * @return boolean
     */
    private function filter($value) {
        return trim($value) !== '';
    }
}

class convert_iconv_filter extends php_user_filter {
    private $modes;

    function filter($in, $out, &$consumed, $closing) {
        while ($bucket = stream_bucket_make_writeable($in)) {
            $bucket->data = mb_convert_encoding($bucket->data, $this->modes[1], $this->modes[0]);
            $consumed += $bucket->datalen;
            stream_bucket_append($out, $bucket);
        }

        return PSFS_PASS_ON;
    }

    function onCreate() {
        $format = explode('/', substr($this->filtername, 14));
        if (count($format) == 2) {
            $this->modes = $format;

            return true;
        } else {
            return false;
        }
    }
}
