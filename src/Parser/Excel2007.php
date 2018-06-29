<?php
/**
 * Excel2017 Parser
 *
 * @author Janson
 * @create 2017-12-02
 */
namespace Asan\PHPExcel\Parser;

use Asan\PHPExcel\Exception\ParserException;
use Asan\PHPExcel\Exception\ReaderException;

class Excel2007 {
    const CELL_TYPE_SHARED_STR = 's';

    /**
     * Temporary directory
     *
     * @var string
     */
    protected $tmpDir;

    /**
     * ZipArchive reader
     *
     * @var \ZipArchive
     */
    protected $zip;

    /**
     * Worksheet reader
     *
     * @var \XMLReader
     */
    protected $worksheetXML;

    /**
     * SharedStrings reader
     *
     * @var \XMLReader
     */
    protected $sharedStringsXML;

    /**
     * SharedStrings position
     *
     * @var array
     */
    private $sharedStringsPosition = -1;

    /**
     * The current sheet of the file
     *
     * @var int
     */
    private $sheetIndex = 0;

    /**
     * Ignore empty row
     *
     * @var bool
     */
    private $ignoreEmpty = false;

    /**
     * Style xfs
     *
     * @var array
     */
    private $styleXfs;

    /**
     * Number formats
     *
     * @var array
     */
    private $formats;

    /**
     * Parsed number formats
     *
     * @var array
     */
    private $parsedFormats;

    /**
     * Worksheets
     *
     * @var array
     */
    private $sheets;

    /**
     * Default options for libxml loader
     *
     * @var int
     */
    private static $libXmlLoaderOptions;

    /**
     * Base date
     *
     * @var \DateTime
     */
    private static $baseDate;
    private static $decimalSeparator = '.';
    private static $thousandSeparator = ',';
    private static $currencyCode = '';
    private static $runtimeInfo = ['GMPSupported' => false];

    /**
     * Use ZipArchive reader to extract the relevant data streams from the ZipArchive file
     *
     * @throws ParserException|ReaderException
     * @param string $file
     */
    public function loadZip($file) {
        $this->openFile($file);

        // Setting base date
        if (!self::$baseDate) {
            self::$baseDate = new \DateTime;
            self::$baseDate->setTimezone(new \DateTimeZone('UTC'));
            self::$baseDate->setDate(1900, 1, 0);
            self::$baseDate->setTime(0, 0, 0);
        }

        if (function_exists('gmp_gcd')) {
            self::$runtimeInfo['GMPSupported'] = true;
        }
    }

    /**
     * Ignore empty row
     *
     * @param bool $ignoreEmpty
     *
     * @return $this
     */
    public function ignoreEmptyRow($ignoreEmpty) {
        $this->ignoreEmpty = $ignoreEmpty;

        return $this;
    }

    /**
     * Whether is ignore empty row
     *
     * @return bool
     */
    public function isIgnoreEmptyRow() {
        return $this->ignoreEmpty;
    }

    /**
     * Set sheet index
     *
     * @param int $index
     *
     * @return $this
     */
    public function setSheetIndex($index) {
        if ($index != $this->sheetIndex) {
            $this->sheetIndex = $index;

            $this->getWorksheetXML();
        }

        return $this;
    }

    /**
     * Get sheet index
     *
     * @return int
     */
    public function getSheetIndex() {
        return $this->sheetIndex;
    }

    /**
     * Return worksheet info (Name, Last Column Letter, Last Column Index, Total Rows, Total Columns)
     *
     * @throws ReaderException
     * @return array
     */
    public function parseWorksheetInfo() {
        if ($this->sheets === null) {
            $workbookXML = simplexml_load_string(
                $this->securityScan($this->zip->getFromName('xl/workbook.xml')), 'SimpleXMLElement', self::getLibXmlLoaderOptions()
            );

            $this->sheets = [];
            if (isset($workbookXML->sheets) && $workbookXML->sheets) {
                $xml = new \XMLReader();

                $index = 0;
                foreach ($workbookXML->sheets->sheet as $sheet) {
                    $info = [
                        'name' => (string)$sheet['name'], 'lastColumnLetter' => '', 'lastColumnIndex' => 0,
                        'totalRows' => 0, 'totalColumns' => 0
                    ];

                    $this->zip->extractTo($this->tmpDir, $file = 'xl/worksheets/sheet' . (++$index) . '.xml');
                    $xml->open($this->tmpDir . '/' . $file, null, self::getLibXmlLoaderOptions());

                    $xml->setParserProperty(\XMLReader::DEFAULTATTRS, true);

                    $nonEmpty = false;
                    $columnLetter = '';
                    while ($xml->read()) {
                        if ($xml->name == 'row') {
                            if (!$this->ignoreEmpty && $xml->nodeType == \XMLReader::ELEMENT) {
                                $info['totalRows'] = (int)$xml->getAttribute('r');
                            } elseif ($xml->nodeType == \XMLReader::END_ELEMENT) {
                                if ($this->ignoreEmpty && $nonEmpty) {
                                    $info['totalRows']++;
                                    $nonEmpty = false;
                                }

                                if ($columnLetter > $info['lastColumnLetter']) {
                                    $info['lastColumnLetter'] = $columnLetter;
                                }
                            }
                        } elseif ($xml->name == 'c' && $xml->nodeType == \XMLReader::ELEMENT) {
                            $columnLetter = preg_replace('{[^[:alpha:]]}S', '', $xml->getAttribute('r'));
                        } elseif ($this->ignoreEmpty && !$nonEmpty && $xml->name == 'v'
                            && $xml->nodeType == \XMLReader::ELEMENT && trim($xml->readString()) !== '') {

                            $nonEmpty = true;
                        }
                    }

                    if ($info['lastColumnLetter']) {
                        $info['totalColumns'] = Format::columnIndexFromString($info['lastColumnLetter']);
                        $info['lastColumnIndex'] = $info['totalColumns'] - 1;
                    }

                    $this->sheets[] = $info;
                }

                $xml->close();
            }
        }

        return $this->sheets;
    }

    /**
     * Get shared string
     *
     * @param int $position
     * @return string
     */
    protected function getSharedString($position) {
        $value = '';

        $file = 'xl/sharedStrings.xml';
        if ($this->sharedStringsXML === null) {
            $this->sharedStringsXML = new \XMLReader();

            $this->zip->extractTo($this->tmpDir, $file);
        }

        if ($this->sharedStringsPosition < 0 || $position < $this->sharedStringsPosition) {
            $this->sharedStringsXML->open($this->tmpDir . '/' . $file, null, self::getLibXmlLoaderOptions());

            $this->sharedStringsPosition = -1;
        }

        while ($this->sharedStringsXML->read()) {
            $name = $this->sharedStringsXML->name;
            $nodeType = $this->sharedStringsXML->nodeType;

            if ($name == 'si') {
                if ($nodeType == \XMLReader::ELEMENT) {
                    $this->sharedStringsPosition++;
                } elseif ($position == $this->sharedStringsPosition && $nodeType == \XMLReader::END_ELEMENT) {
                    break;
                }
            } elseif ($name == 't' && $position == $this->sharedStringsPosition && $nodeType == \XMLReader::ELEMENT) {
                $value .= trim($this->sharedStringsXML->readString());
            }
        }

        return $value;
    }

    /**
     * Parse styles info
     *
     * @throws ReaderException
     */
    protected function parseStyles() {
        if ($this->styleXfs === null) {
            $stylesXML = simplexml_load_string(
                $this->securityScan($this->zip->getFromName('xl/styles.xml')), 'SimpleXMLElement', self::getLibXmlLoaderOptions()
            );

            $this->styleXfs = $this->formats = [];
            if ($stylesXML) {
                if (isset($stylesXML->cellXfs->xf) && $stylesXML->cellXfs->xf) {
                    foreach ($stylesXML->cellXfs->xf as $xf) {
                        $numFmtId = isset($xf['numFmtId']) ? (int)$xf['numFmtId'] : 0;
                        if (isset($xf['applyNumberFormat']) || $numFmtId == 0) {
                            // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                            $this->styleXfs[] = $numFmtId;
                        } else {
                            // 0 for "General" format
                            $this->styleXfs[] = Format::FORMAT_GENERAL;
                        }
                    }
                }

                if (isset($stylesXML->numFmts->numFmt) && $stylesXML->numFmts->numFmt) {
                    foreach ($stylesXML->numFmts->numFmt as $numFmt) {
                        if (isset($numFmt['numFmtId'], $numFmt['formatCode'])) {
                            $this->formats[(int)$numFmt['numFmtId']] = (string)$numFmt['formatCode'];
                        }
                    }
                }
            }
        }
    }

    /**
     * Get worksheet XMLReader
     */
    protected function getWorksheetXML() {
        if ($this->worksheetXML === null) {
            $this->worksheetXML = new \XMLReader();
        }

        $this->worksheetXML->open(
            $this->tmpDir . '/xl/worksheets/sheet' . ($this->getSheetIndex() + 1) . '.xml',
            null, self::getLibXmlLoaderOptions()
        );
    }

    /**
     * Get row data
     *
     * @param int $rowIndex
     * @param int $columnLimit
     *
     * @throws ReaderException
     * @return array|bool
     */
    public function getRow($rowIndex, $columnLimit = 0) {
        $this->parseStyles();
        $rowIndex === 0 && $this->getWorksheetXML();

        $sharedString = false;
        $index = $styleId = 0;
        $row = $columnLimit ? array_fill(0, $columnLimit, '') : [];

        while ($canRead = $this->worksheetXML->read()) {
            $name = $this->worksheetXML->name;
            $type = $this->worksheetXML->nodeType;

            // End of row
            if ($name == 'row') {
                if (!$this->ignoreEmpty && $type == \XMLReader::ELEMENT
                    && $rowIndex+1 != (int)$this->worksheetXML->getAttribute('r')) {

                    $this->worksheetXML->moveToElement();
                    break;
                }

                if ($type == \XMLReader::END_ELEMENT) {
                    break;
                }
            }

            if ($columnLimit > 0 && $index >= $columnLimit) {
                continue;
            }

            switch ($name) {
                // Cell
                case 'c':
                    if ($type == \XMLReader::END_ELEMENT) {
                        continue;
                    }

                    $styleId = (int)$this->worksheetXML->getAttribute('s');
                    $letter = preg_replace('{[^[:alpha:]]}S', '', $this->worksheetXML->getAttribute('r'));
                    $index = Format::columnIndexFromString($letter) - 1;

                    // Determine cell type
                    $sharedString = false;
                    if ($this->worksheetXML->getAttribute('t') == self::CELL_TYPE_SHARED_STR) {
                        $sharedString = true;
                    }

                    break;

                // Cell value
                case 'v':
                case 'is':
                    if ($type == \XMLReader::END_ELEMENT) {
                        continue;
                    }

                    $value = $this->worksheetXML->readString();
                    if ($sharedString) {
                        $value = $this->getSharedString($value);
                    }

                    // Format value if necessary
                    if ($value !== '' && $styleId && isset($this->styleXfs[$styleId])) {
                        $value = $this->formatValue($value, $styleId);
                    } elseif ($value && is_numeric($value)) {
                        $value = (float)$value;
                    }

                    $row[$index] = $value;
                    break;
            }
        }

        if ($canRead === false) {
            return false;
        }

        return $row;
    }

    /**
     * Close ZipArchive、XMLReader and remove temp dir
     */
    public function __destruct() {
        if ($this->zip && $this->tmpDir) {
            $this->zip->close();
        }

        if ($this->worksheetXML) {
            $this->worksheetXML->close();
        }

        if ($this->sharedStringsXML) {
            $this->sharedStringsXML->close();
        }

        $this->removeDir($this->tmpDir);

        $this->zip = null;
        $this->worksheetXML = null;
        $this->sharedStringsXML = null;
        $this->tmpDir = null;
    }

    /**
     * Remove dir
     *
     * @param string $dir
     */
    protected function removeDir($dir) {
        if($dir && is_dir($dir)) {
            $handle = opendir($dir);

            while($item = readdir($handle)) {
                if ($item != '.' && $item != '..') {
                    is_file($item = $dir . '/' . $item) ? unlink($item) : $this->removeDir($item);
                }
            }

            closedir($handle);
            rmdir($dir);
        }
    }

    /**
     * Formats the value according to the index
     *
     * @param string $value
     * @param int $index Format index
     *
     * @throws \Exception
     * @return string Formatted cell value
     */
    private function formatValue($value, $index) {
        if (!is_numeric($value)) {
            return $value;
        }

        if (isset($this->styleXfs[$index]) && $this->styleXfs[$index] !== false) {
            $index = $this->styleXfs[$index];
        } else {
            return $value;
        }

        // A special case for the "General" format
        if ($index == 0) {
            return is_numeric($value) ? (float)$value : $value;
        }

        $format = $this->parsedFormats[$index] ?? [];

        if (empty($format)) {
            $format = [
                'code' => false, 'type' => false, 'scale' => 1, 'thousands' => false, 'currency' => false
            ];

            if (isset(Format::$buildInFormats[$index])) {
                $format['code'] = Format::$buildInFormats[$index];
            } elseif (isset($this->formats[$index])) {
                $format['code'] = str_replace('"', '', $this->formats[$index]);
            }

            // Format code found, now parsing the format
            if ($format['code']) {
                $sections = explode(';', $format['code']);
                $format['code'] = $sections[0];

                switch (count($sections)) {
                    case 2:
                        if ($value < 0) {
                            $format['code'] = $sections[1];
                        }

                        $value = abs($value);
                        break;

                    case 3:
                    case 4:
                        if ($value < 0) {
                            $format['code'] = $sections[1];
                        } elseif ($value == 0) {
                            $format['code'] = $sections[2];
                        }

                        $value = abs($value);
                        break;
                }
            }

            // Stripping colors
            $format['code'] = trim(preg_replace('/^\\[[a-zA-Z]+\\]/', '', $format['code']));

            // Percentages
            if (substr($format['code'], -1) == '%') {
                $format['type'] = 'Percentage';
            } elseif (preg_match('/(\[\$[A-Z]*-[0-9A-F]*\])*[hmsdy]/i', $format['code'])) {
                $format['type'] = 'DateTime';
                $format['code'] = trim(preg_replace('/^(\[\$[A-Z]*-[0-9A-F]*\])/i', '', $format['code']));
                $format['code'] = strtolower($format['code']);
                $format['code'] = strtr($format['code'], Format::$dateFormatReplacements);

                if (strpos($format['code'], 'A') === false) {
                    $format['code'] = strtr($format['code'], Format::$dateFormatReplacements24);
                } else {
                    $format['code'] = strtr($format['code'], Format::$dateFormatReplacements12);
                }
            } elseif ($format['code'] == '[$EUR ]#,##0.00_-') {
                $format['type'] = 'Euro';
            } else {
                // Removing skipped characters
                $format['code'] = preg_replace('/_./', '', $format['code']);

                // Removing unnecessary escaping
                $format['code'] = preg_replace("/\\\\/", '', $format['code']);

                // Removing string quotes
                $format['code'] = str_replace(['"', '*'], '', $format['code']);

                // Removing thousands separator
                if (strpos($format['code'], '0,0') !== false || strpos($format['code'], '#,#') !== false) {
                    $format['thousands'] = true;
                }

                $format['code'] = str_replace(['0,0', '#,#'], ['00', '##'], $format['code']);

                // Scaling (Commas indicate the power)
                $scale = 1;
                $matches = [];

                if (preg_match('/(0|#)(,+)/', $format['code'], $matches)) {
                    $scale = pow(1000, strlen($matches[2]));

                    // Removing the commas
                    $format['code'] = preg_replace(['/0,+/', '/#,+/'], ['0', '#'], $format['code']);
                }

                $format['scale'] = $scale;
                if (preg_match('/#?.*\?\/\?/', $format['code'])) {
                    $format['type'] = 'Fraction';
                } else {
                    $format['code'] = str_replace('#', '', $format['code']);
                    $matches = [];

                    if (preg_match('/(0+)(\.?)(0*)/', preg_replace('/\[[^\]]+\]/', '', $format['code']), $matches)) {
                        list(, $integer, $decimalPoint, $decimal) = $matches;

                        $format['minWidth'] = strlen($integer) + strlen($decimalPoint) + strlen($decimal);
                        $format['decimals'] = $decimal;
                        $format['precision'] = strlen($format['decimals']);
                        $format['pattern'] = '%0' . $format['minWidth'] . '.' . $format['precision'] . 'f';
                    }
                }

                $matches = [];
                if (preg_match('/\[\$(.*)\]/u', $format['code'], $matches)) {
                    $currencyCode = explode('-', $matches[1]);
                    if ($currencyCode) {
                        $currencyCode = $currencyCode[0];
                    }

                    if (!$currencyCode) {
                        $currencyCode = self::$currencyCode;
                    }

                    $format['currency'] = $currencyCode;
                }

                $format['code'] = trim($format['code']);
            }

            $this->parsedFormats[$index] = $format;
        }

        // Applying format to value
        if ($format) {
            if ($format['code'] == '@') {
                return (string)$value;
            } elseif ($format['type'] == 'Percentage') { // Percentages
                if ($format['code'] === '0%') {
                    $value = round(100*$value, 0) . '%';
                } else {
                    $value = sprintf('%.2f%%', round(100*$value, 2));
                }
            } elseif ($format['type'] == 'DateTime') { // Dates and times
                $days = (int)$value;

                // Correcting for Feb 29, 1900
                if ($days > 60) {
                    $days--;
                }

                // At this point time is a fraction of a day
                $time = ($value - (int)$value);

                // Here time is converted to seconds
                // Some loss of precision will occur
                $seconds = $time ? (int)($time*86400) : 0;

                $value = clone self::$baseDate;
                $value->add(new \DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));

                $value = $value->format($format['code']);
            } elseif ($format['type'] == 'Euro') {
                $value = 'EUR ' . sprintf('%1.2f', $value);
            } else {
                // Fractional numbers
                if ($format['type'] == 'Fraction' && ($value != (int)$value)) {
                    $integer = floor(abs($value));
                    $decimal = fmod(abs($value), 1);

                    // Removing the integer part and decimal point
                    $decimal *= pow(10, strlen($decimal) - 2);
                    $decimalDivisor = pow(10, strlen($decimal));

                    if (self::$runtimeInfo['GMPSupported']) {
                        $GCD = gmp_strval(gmp_gcd($decimal, $decimalDivisor));
                    } else {
                        $GCD = self::GCD($decimal, $decimalDivisor);
                    }

                    $adjDecimal = $decimal/$GCD;
                    $adjDecimalDivisor = $decimalDivisor/$GCD;

                    if (strpos($format['code'], '0') !== false || strpos($format['code'], '#') !== false
                        || substr($format['code'], 0, 3) == '? ?') {

                        // The integer part is shown separately apart from the fraction
                        $value = ($value < 0 ? '-' : '') . $integer ? $integer . ' '
                            : '' . $adjDecimal . '/' . $adjDecimalDivisor;
                    } else {
                        // The fraction includes the integer part
                        $adjDecimal += $integer * $adjDecimalDivisor;
                        $value = ($value < 0 ? '-' : '') . $adjDecimal . '/' . $adjDecimalDivisor;
                    }
                } else {
                    // Scaling
                    $value = $value/$format['scale'];
                    if (!empty($format['minWidth']) && $format['decimals']) {
                        if ($format['thousands']) {
                            $value = number_format(
                                $value, $format['precision'], self::$decimalSeparator, self::$thousandSeparator
                            );

                            $value = preg_replace('/(0+)(\.?)(0*)/', $value, $format['code']);
                        } else {
                            if (preg_match('/[0#]E[+-]0/i', $format['code'])) {
                                // Scientific format
                                $value = sprintf('%5.2E', $value);
                            } else {
                                $value = sprintf($format['pattern'], $value);
                                $value = preg_replace('/(0+)(\.?)(0*)/', $value, $format['code']);
                            }
                        }
                    }
                }

                // currency/Accounting
                if ($format['currency']) {
                    $value = preg_replace('', $format['currency'], $value);
                }
            }
        }

        return $value;
    }

    /**
     * Greatest common divisor calculation in case GMP extension is not enabled
     *
     * @param int $number1
     * @param int $number2
     *
     * @return int
     */
    private static function GCD($number1, $number2) {
        $number1 = abs($number1);
        $number2 = abs($number2);

        if ($number1 + $number2 == 0) {
            return 0;
        }

        $number = 1;
        while ($number1 > 0) {
            $number = $number1;
            $number1 = $number2 % $number1;
            $number2 = $number;
        }

        return $number;
    }

    /**
     * Open file for reading
     *
     * @param string $file
     *
     * @throws ParserException|ReaderException
     */
    public function openFile($file) {
        // Check if file exists
        if (!file_exists($file) || !is_readable($file)) {
            throw new ReaderException("Could not open file [$file] for reading! File does not exist.");
        }

        $this->zip = new \ZipArchive();

        $xl = false;
        if ($this->zip->open($file) === true) {
            $this->tmpDir = sys_get_temp_dir() . '/' . uniqid();

            // check if it is an OOXML archive
            $rels = simplexml_load_string(
                $this->securityScan($this->zip->getFromName('_rels/.rels')),
                'SimpleXMLElement', self::getLibXmlLoaderOptions()
            );

            if ($rels !== false) {
                foreach ($rels->Relationship as $rel) {
                    switch ($rel["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                            if ($rel["Target"] == 'xl/workbook.xml') {
                                $xl = true;
                            }

                            break;
                    }
                }
            }
        }

        if ($xl === false) {
            throw new ParserException("The file [$file] is not recognised as a zip archive");
        }
    }

    /**
     * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
     *
     * @param  string $xml
     *
     * @throws ReaderException
     * @return string
     */
    protected function securityScan($xml) {
        $pattern = sprintf('/\\0?%s\\0?/', implode('\\0?', str_split('<!DOCTYPE')));

        if (preg_match($pattern, $xml)) {
            throw new ReaderException(
                'Detected use of ENTITY in XML, spreadsheet file load() aborted to prevent XXE/XEE attacks'
            );
        }

        return $xml;
    }

    /**
     * Set default options for libxml loader
     *
     * @param int $options Default options for libxml loader
     */
    public static function setLibXmlLoaderOptions($options = null) {
        if (is_null($options) && defined(LIBXML_DTDLOAD)) {
            $options = LIBXML_DTDLOAD | LIBXML_DTDATTR;
        }

        if (version_compare(PHP_VERSION, '5.2.11') >= 0) {
            @libxml_disable_entity_loader($options == (LIBXML_DTDLOAD | LIBXML_DTDATTR));
        }

        self::$libXmlLoaderOptions = $options;
    }

    /**
     * Get default options for libxml loader.
     * Defaults to LIBXML_DTDLOAD | LIBXML_DTDATTR when not set explicitly.
     *
     * @return int Default options for libxml loader
     */
    public static function getLibXmlLoaderOptions() {
        if (is_null(self::$libXmlLoaderOptions) && defined(LIBXML_DTDLOAD)) {
            self::setLibXmlLoaderOptions(LIBXML_DTDLOAD | LIBXML_DTDATTR);
        }

        if (version_compare(PHP_VERSION, '5.2.11') >= 0) {
            @libxml_disable_entity_loader(self::$libXmlLoaderOptions == (LIBXML_DTDLOAD | LIBXML_DTDATTR));
        }

        return self::$libXmlLoaderOptions;
    }
}
