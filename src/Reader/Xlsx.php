<?php
/**
 * Xlsx Reader
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace Asan\PHPExcel\Reader;

use Asan\PHPExcel\Parser\Excel2007;

class Xlsx extends BaseReader {
    /**
     * Xls parser
     *
     * @var Excel2007
     */
    protected $parser;

    /**
     * File row、column count
     *
     * @var array|int
     */
    protected $count;

    public function __construct() {
        $this->parser = new Excel2007();
    }

    /**
     * Loads Excel from file
     *
     * @param string $file
     *
     * @return $this
     */
    public function load($file) {
        $this->parser->loadZip($file);

        $this->generator = $this->makeGenerator();

        return $this;
    }

    /**
     * Count elements of an object
     *
     * @param bool $all
     * @return int|array
     */
    public function count($all = false) {
        if ($this->count === null) {
            $row = $column = 0;
            if ($sheet = $this->sheets($this->parser->getSheetIndex())) {
                $row = $sheet['totalRows'] ?? 0;
                $column = $sheet['totalColumns'] ?? 0;
            }

            $this->count = [
                $this->rowLimit > 0 ? min($row, $this->rowLimit) : $row,
                $this->columnLimit > 0 ? min($column, $this->columnLimit) : $column
            ];
        }

        return $all ? $this->count : $this->count[0];
    }

    /**
     * Get the work sheets info
     *
     * @param int $index
     * @return array
     */
    public function sheets($index = null) {
        $sheets = $this->parser->parseWorksheetInfo();

        if ($index !== null) {
            return $sheets[$index] ?? [];
        }

        return $sheets;
    }

    /**
     * Make the generator
     *
     * @return \Generator
     */
    protected function makeGenerator() {
        list($rowLimit, $columnLimit) = $this->count(true);

        $line = $finish = 0;
        while ($finish < $rowLimit && ($row = $this->parser->getRow($line++, $columnLimit)) !== false) {
            if ($this->parser->isIgnoreEmptyRow() && trim(implode('', $row)) === '') {
                continue;
            }

            $finish++;
            yield $row;
        }
    }

    /**
     * Ignore empty row
     *
     * @param bool $ignoreEmpty
     *
     * @return $this
     */
    public function ignoreEmptyRow($ignoreEmpty = false) {
        $this->parser->ignoreEmptyRow($ignoreEmpty);

        return $this;
    }

    /**
     * Set sheet index
     *
     * @param int $index
     * @return $this
     */
    public function setSheetIndex($index = 0) {
        if ($index != $this->parser->getSheetIndex()) {
            $this->parser->setSheetIndex($index);

            $this->count = null;
            $this->rewind();
        }

        return $this;
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
            $parser = new Excel2007();

            // open file
            $parser->openFile($file);
        } catch (\Exception $e) {
            return false;
        }

        return true;
    }

    /**
     * Release parser and generator
     */
    public function __destruct() {
        $this->parser = null;
        $this->generator = null;
    }
}
