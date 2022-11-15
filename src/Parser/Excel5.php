<?php
/**
 * Excel5 Parser
 *
 * @author Janson
 * @create 2017-11-27
 */
namespace Asan\PHPExcel\Parser;

use Asan\PHPExcel\Exception\ParserException;
use Asan\PHPExcel\Parser\Excel5\OLERead;
use Asan\PHPExcel\Parser\Excel5\RC4;

class Excel5 {
    // ParseXL definitions
    const XLS_BIFF8 = 0x0600;
    const XLS_BIFF7 = 0x0500;
    const XLS_WORKBOOKGLOBALS = 0x0005;
    const XLS_WORKSHEET = 0x0010;

    // record identifiers
    const XLS_TYPE_FORMULA = 0x0006;
    const XLS_TYPE_EOF = 0x000a;
    const XLS_TYPE_DATEMODE = 0x0022;
    const XLS_TYPE_FILEPASS = 0x002f;
    const XLS_TYPE_CONTINUE = 0x003c;
    const XLS_TYPE_CODEPAGE = 0x0042;
    const XLS_TYPE_OBJ = 0x005d;
    const XLS_TYPE_SHEET = 0x0085;
    const XLS_TYPE_MULRK = 0x00bd;
    const XLS_TYPE_MULBLANK = 0x00be;
    const XLS_TYPE_XF = 0x00e0;
    const XLS_TYPE_SST = 0x00fc;
    const XLS_TYPE_LABELSST = 0x00fd;
    const XLS_TYPE_BLANK = 0x0201;
    const XLS_TYPE_NUMBER = 0x0203;
    const XLS_TYPE_LABEL = 0x0204;
    const XLS_TYPE_BOOLERR = 0x0205;
    const XLS_TYPE_STRING = 0x0207;
    const XLS_TYPE_ROW = 0x0208;
    const XLS_TYPE_INDEX = 0x020b;
    const XLS_TYPE_ARRAY = 0x0221;
    const XLS_TYPE_RK = 0x027e;
    const XLS_TYPE_FORMAT = 0x041e;
    const XLS_TYPE_BOF = 0x0809;

    // Encryption type
    const MS_BIFF_CRYPTO_NONE = 0;
    const MS_BIFF_CRYPTO_XOR = 1;
    const MS_BIFF_CRYPTO_RC4 = 2;

    // Size of stream blocks when using RC4 encryption
    const REKEY_BLOCK = 0x400;

    // Sheet state
    const SHEETSTATE_VISIBLE = 'visible';
    const SHEETSTATE_HIDDEN = 'hidden';
    const SHEETSTATE_VERYHIDDEN = 'veryHidden';

    private static $errorCode = [
        0x00 => '#NULL!',
        0x07 => '#DIV/0!',
        0x0F => '#VALUE!',
        0x17 => '#REF!',
        0x1D => '#NAME?',
        0x24 => '#NUM!',
        0x2A => '#N/A'
    ];

    /**
     * Base calendar year to use for calculations
     *
     * @var int
     */
    private static $excelBaseDate = Format::CALENDAR_WINDOWS_1900;

    /**
     * Decimal separator
     *
     * @var string
     */
    private static $decimalSeparator;

    /**
     * Thousands separator
     *
     * @var string
     */
    private static $thousandsSeparator;

    /**
     * Currency code
     *
     * @var string
     */
    private static $currencyCode;

    /**
     * Workbook stream data
     *
     * @var string
     */
    private $data;

    /**
     * Size in bytes of $this->data
     *
     * @var int
     */
    private $dataSize;

    /**
     * Current position in stream
     *
     * @var integer
     */
    private $pos;

    /**
     * Worksheets
     *
     * @var array
     */
    private $sheets;

    /**
     * BIFF version
     *
     * @var int
     */
    private $version;

    /**
     * Codepage set in the Excel file being read. Only important for BIFF5 (Excel 5.0 - Excel 95)
     * For BIFF8 (Excel 97 - Excel 2003) this will always have the value 'UTF-16LE'
     *
     * @var string
     */
    private $codePage;

    /**
     * Row data
     *
     * @var array
     */
    private $row;

    /**
     * Shared formats
     *
     * @var array
     */
    private $formats;

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
     * The current row index of the sheet
     *
     * @var int
     */
    private $rowIndex = 0;

    /**
     * Max column number
     *
     * @var int
     */
    private $columnLimit = 0;

    /**
     * Whether to the end of the row
     *
     * @var bool
     */
    private $eor = false;

    /**
     * Extended format record
     *
     * @var array
     */
    private $xfRecords = [];

    /**
     * Shared strings. Only applies to BIFF8.
     *
     * @var array
     */
    private $sst = [];

    /**
     * The type of encryption in use
     *
     * @var int
     */
    private $encryption = 0;

    /**
     * The position in the stream after which contents are encrypted
     *
     * @var int
     */
    private $encryptionStartPos = false;

    /**
     * The current RC4 decryption object
     *
     * @var RC4
     */
    private $rc4Key = null;

    /**
     * The position in the stream that the RC4 decryption object was left at
     *
     * @var int
     */
    private $rc4Pos = 0;

    /**
     * The current MD5 context state
     *
     * @var string
     */
    private $md5Ctxt = null;

    /**
     * Use OLE reader to extract the relevant data streams from the OLE file
     *
     * @param string $file
     */
    public function loadOLE($file) {
        $oleRead = new OLERead();
        $oleRead->read($file);
        $this->data = $oleRead->getStream($oleRead->workbook);
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
        $this->sheetIndex = $index;

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
     * @throws ParserException
     * @return array
     */
    public function parseWorksheetInfo() {
        if ($this->sheets === null) {
            // total byte size of Excel data (workbook global substream + sheet substreams)
            $this->dataSize = strlen($this->data);
            $this->pos = 0;
            $this->codePage = 'CP1252';
            $this->sheets = [];

            // Parse Workbook Global Substream
            while ($this->pos < $this->dataSize) {
                $code = Format::getUInt2d($this->data, $this->pos);

                switch ($code) {
                    case self::XLS_TYPE_BOF:
                        $this->readBof();
                        break;

                    case self::XLS_TYPE_FILEPASS:
                        $this->readFilepass();
                        break;

                    case self::XLS_TYPE_CODEPAGE:
                        $this->readCodepage();
                        break;

                    case self::XLS_TYPE_DATEMODE:
                        $this->readDateMode();
                        break;

                    case self::XLS_TYPE_FORMAT:
                        $this->readFormat();
                        break;

                    case self::XLS_TYPE_XF:
                        $this->readXf();
                        break;

                    case self::XLS_TYPE_SST:
                        $this->readSst();
                        break;

                    case self::XLS_TYPE_SHEET:
                        $this->readSheet();
                        break;

                    case self::XLS_TYPE_EOF:
                        $this->readDefault();
                        break 2;

                    default:
                        $this->readDefault();
                        break;
                }
            }

            // Parse the individual sheets
            foreach ($this->sheets as $key => $sheet) {
                if ($sheet['sheetType'] != 0x00) {
                    // 0x00: Worksheet
                    // 0x02: Chart
                    // 0x06: Visual Basic module
                    continue;
                }

                $sheet['lastColumnLetter'] = '';
                $sheet['lastColumnIndex'] = null;
                $sheet['totalRows'] = 0;
                $sheet['totalColumns'] = 0;

                $lastRowIndex = 0;
                $this->pos = $sheet['offset'];
                while ($this->pos <= $this->dataSize - 4) {
                    $code = Format::getUInt2d($this->data, $this->pos);

                    switch ($code) {
                        case self::XLS_TYPE_RK:
                        case self::XLS_TYPE_LABELSST:
                        case self::XLS_TYPE_NUMBER:
                        case self::XLS_TYPE_FORMULA:
                        case self::XLS_TYPE_BOOLERR:
                        case self::XLS_TYPE_LABEL:
                            $length = Format::getUInt2d($this->data, $this->pos + 2);
                            $recordData = substr($this->data, $this->pos + 4, $length);

                            // move stream pointer to next record
                            $this->pos += 4 + $length;

                            $rowIndex = Format::getUInt2d($recordData, 0) + 1;
                            $columnIndex = Format::getUInt2d($recordData, 2);

                            if ($this->ignoreEmpty) {
                                if ($lastRowIndex < $rowIndex) {
                                    $sheet['totalRows']++;
                                }

                                $lastRowIndex = $rowIndex;
                            } else {
                                $sheet['totalRows'] = max($sheet['totalRows'], $rowIndex);
                            }

                            $sheet['lastColumnIndex'] = max($columnIndex, $sheet['lastColumnIndex']);
                            break;

                        case self::XLS_TYPE_BOF:
                            $this->readBof();
                            break;

                        case self::XLS_TYPE_EOF:
                            $this->readDefault();
                            break 2;

                        default:
                            $this->readDefault();
                            break;
                    }
                }

                if ($sheet['lastColumnIndex'] !== null) {
                    $sheet['lastColumnLetter'] = Format::stringFromColumnIndex($sheet['lastColumnIndex']);
                } else {
                    $sheet['lastColumnIndex'] = 0;
                }

                if ($sheet['lastColumnLetter']) {
                    $sheet['totalColumns'] = $sheet['lastColumnIndex'] + 1;
                }

                $this->sheets[$key] = $sheet;
            }

            $this->pos = 0;
        }

        return $this->sheets;
    }

    /**
     * Get row data
     *
     * @param int $rowIndex
     * @param int $columnLimit
     *
     * @throws ParserException
     * @return array|bool
     */
    public function getRow($rowIndex, $columnLimit = 0) {
        $this->parseWorksheetInfo();

        // Rewind or change sheet
        if ($rowIndex === 0 || $this->pos < $this->sheets[$this->sheetIndex]['offset']) {
            $this->pos = $this->sheets[$this->sheetIndex]['offset'];
        }

        $endPos = $this->dataSize - 4;
        if (isset($this->sheets[$this->sheetIndex + 1]['offset'])) {
            $endPos = $this->sheets[$this->sheetIndex + 1]['offset'] - 4;
        }

        if ($this->pos >= $endPos) {
            return false;
        }

        $this->rowIndex = $rowIndex;
        $this->columnLimit = $columnLimit;
        $this->eor = false;
        $this->row = $columnLimit ? array_fill(0, $columnLimit, '') : [];

        while ($this->pos <= $endPos) {
            // Remember last position
            $lastPos = $this->pos;
            $code = Format::getUInt2d($this->data, $this->pos);

            switch ($code) {
                case self::XLS_TYPE_BOF:
                    $this->readBof();
                    break;

                case self::XLS_TYPE_RK:
                    $this->readRk();
                    break;

                case self::XLS_TYPE_LABELSST:
                    $this->readLabelSst();
                    break;

                case self::XLS_TYPE_MULRK:
                    $this->readMulRk();
                    break;

                case self::XLS_TYPE_NUMBER:
                    $this->readNumber();
                    break;

                case self::XLS_TYPE_FORMULA:
                    $this->readFormula();
                    break;

                case self::XLS_TYPE_BOOLERR:
                    $this->readBoolErr();
                    break;

                case self::XLS_TYPE_MULBLANK:
                case self::XLS_TYPE_BLANK:
                    $this->readBlank();
                    break;

                case self::XLS_TYPE_LABEL:
                    $this->readLabel();
                    break;

                case self::XLS_TYPE_EOF:
                    $this->readDefault();
                    break 2;

                default:
                    $this->readDefault();
                    break;
            }

            //End of row
            if ($this->eor) {
                //Recover current position
                $this->pos = $lastPos;
                break;
            }
        }

        return $this->row;
    }

    /**
     * Add cell data
     *
     * @param int $row
     * @param int $column
     * @param mixed $value
     * @param int $xfIndex
     * @return bool
     */
    private function addCell($row, $column, $value, $xfIndex) {
        if ($this->rowIndex != $row) {
            $this->eor = true;

            return false;
        }

        if (!$this->columnLimit || $column < $this->columnLimit) {
            $xfRecord = $this->xfRecords[$xfIndex];
            $this->row[$column] = self::toFormattedString($value, $xfRecord['format']);
        }

        return true;
    }

    /**
     * Read BOF
     *
     * @throws ParserException
     */
    private function readBof() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 2; size: 2; type of the following data
        $substreamType = Format::getUInt2d($recordData, 2);

        switch ($substreamType) {
            case self::XLS_WORKBOOKGLOBALS:
                $version = Format::getUInt2d($recordData, 0);
                if (($version != self::XLS_BIFF8) && ($version != self::XLS_BIFF7)) {
                    throw new ParserException('Cannot read this Excel file. Version is too old.', 1);
                }

                $this->version = $version;
                break;

            case self::XLS_WORKSHEET:
                // do not use this version information for anything
                // it is unreliable (OpenOffice doc, 5.8), use only version information from the global stream
                break;

            default:
                // substream, e.g. chart
                // just skip the entire substream
                do {
                    $code = Format::getUInt2d($this->data, $this->pos);
                    $this->readDefault();
                } while ($code != self::XLS_TYPE_EOF && $this->pos < $this->dataSize);

                break;
        }
    }

    /**
     * SHEET
     *
     * This record is located in the Workbook Globals Substream and represents a sheet inside the workbook.
     * One SHEET record is written for each sheet. It stores the sheet name and a stream offset to the BOF
     * record of the respective Sheet Substream within the Workbook Stream.
     */
    private function readSheet() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // offset: 0; size: 4; absolute stream position of the BOF record of the sheet
        // NOTE: not encrypted
        $offset = Format::getInt4d($this->data, $this->pos + 4);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 4; size: 1; sheet state
        switch (ord($recordData[4])) {
            case 0x00:
                $sheetState = self::SHEETSTATE_VISIBLE;
                break;

            case 0x01:
                $sheetState = self::SHEETSTATE_HIDDEN;
                break;

            case 0x02:
                $sheetState = self::SHEETSTATE_VERYHIDDEN;
                break;

            default:
                $sheetState = self::SHEETSTATE_VISIBLE;
                break;
        }

        // offset: 5; size: 1; sheet type
        $sheetType = ord($recordData[5]);

        // offset: 6; size: var; sheet name
        $name = '';
        if ($this->version == self::XLS_BIFF8) {
            $string = self::readUnicodeStringShort(substr($recordData, 6));
            $name = $string['value'];
        } elseif ($this->version == self::XLS_BIFF7) {
            $string = $this->readByteStringShort(substr($recordData, 6));
            $name = $string['value'];
        }

        // ignore hidden sheet
        if ($sheetState == self::SHEETSTATE_VISIBLE) {
            $this->sheets[] = [
                'name' => $name, 'offset' => $offset, 'sheetState' => $sheetState, 'sheetType' => $sheetType
            ];
        }
    }

    /**
     * Reads a general type of BIFF record.
     * Does nothing except for moving stream pointer forward to next record.
     */
    private function readDefault() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        //$recordData = $this->readRecordData($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;
    }

    /**
     * FILEPASS
     *
     * This record is part of the File Protection Block. It contains information about the read/write password of
     * the file. All record contents following this record will be encrypted.
     * The decryption functions and objects used from here on in are based on the source of Spreadsheet-ParseExcel:
     * http://search.cpan.org/~jmcnamara/Spreadsheet-ParseExcel/
     *
     * @throws ParserException
     */
    private function readFilepass() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);

        if ($length != 54) {
            throw new ParserException('Unexpected file pass record length', 2);
        }

        $recordData = $this->readRecordData($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        if (!$this->verifyPassword('VelvetSweatshop', substr($recordData, 6,  16), substr($recordData, 22, 16),
            substr($recordData, 38, 16), $this->md5Ctxt)) {

            throw new ParserException('Decryption password incorrect', 3);
        }

        $this->encryption = self::MS_BIFF_CRYPTO_RC4;

        // Decryption required from the record after next onwards
        $this->encryptionStartPos = $this->pos + Format::getUInt2d($this->data, $this->pos + 2);
    }

    /**
     * Read record data from stream, decrypting as required
     *
     * @param string $data Data stream to read from
     * @param int $pos Position to start reading from
     * @param int $len Record data length
     *
     * @throws ParserException
     * @return string Record data
     */
    private function readRecordData($data, $pos, $len) {
        $data = substr($data, $pos, $len);

        // File not encrypted, or record before encryption start point
        if ($this->encryption == self::MS_BIFF_CRYPTO_NONE || $pos < $this->encryptionStartPos) {
            return $data;
        }

        $recordData = '';
        if ($this->encryption == self::MS_BIFF_CRYPTO_RC4) {
            $oldBlock = floor($this->rc4Pos / self::REKEY_BLOCK);
            $block = floor($pos / self::REKEY_BLOCK);
            $endBlock = floor(($pos + $len) / self::REKEY_BLOCK);

            // Spin an RC4 decryptor to the right spot. If we have a decryptor sitting
            // at a point earlier in the current block, re-use it as we can save some time.
            if ($block != $oldBlock || $pos < $this->rc4Pos || !$this->rc4Key) {
                $this->rc4Key = $this->makeKey($block, $this->md5Ctxt);
                $step = $pos % self::REKEY_BLOCK;
            } else {
                $step = $pos - $this->rc4Pos;
            }

            $this->rc4Key->RC4(str_repeat("\0", $step));

            // Decrypt record data (re-keying at the end of every block)
            while ($block != $endBlock) {
                $step = self::REKEY_BLOCK - ($pos % self::REKEY_BLOCK);
                $recordData .= $this->rc4Key->RC4(substr($data, 0, $step));

                $data = substr($data, $step);
                $pos += $step;
                $len -= $step;
                $block++;

                $this->rc4Key = $this->makeKey($block, $this->md5Ctxt);
            }

            $recordData .= $this->rc4Key->RC4(substr($data, 0, $len));

            // Keep track of the position of this decryptor.
            // We'll try and re-use it later if we can to speed things up
            $this->rc4Pos = $pos + $len;

        } elseif ($this->encryption == self::MS_BIFF_CRYPTO_XOR) {
            throw new ParserException('XOr encryption not supported', 4);
        }

        return $recordData;
    }

    /**
     * Make an RC4 decryptor for the given block
     *
     * @param int $block Block for which to create decrypto
     * @param string $valContext MD5 context state
     *
     * @return RC4
     */
    private function makeKey($block, $valContext) {
        $pw = str_repeat("\0", 64);

        for ($i = 0; $i < 5; $i++) {
            $pw[$i] = $valContext[$i];
        }

        $pw[5] = chr($block & 0xff);
        $pw[6] = chr(($block >> 8) & 0xff);
        $pw[7] = chr(($block >> 16) & 0xff);
        $pw[8] = chr(($block >> 24) & 0xff);

        $pw[9] = "\x80";
        $pw[56] = "\x48";

        return new RC4(md5($pw));
    }

    /**
     * Verify RC4 file password
     *
     * @var string $password        Password to check
     * @var string $docid           Document id
     * @var string $salt_data       Salt data
     * @var string $hashedsalt_data Hashed salt data
     * @var string &$valContext     Set to the MD5 context of the value
     *
     * @return bool Success
     */
    private function verifyPassword($password, $docid, $salt_data, $hashedsalt_data, &$valContext) {
        $pw = str_repeat("\0", 64);

        for ($i = 0; $i < strlen($password); $i++) {
            $o = ord(substr($password, $i, 1));
            $pw[2 * $i] = chr($o & 0xff);
            $pw[2 * $i + 1] = chr(($o >> 8) & 0xff);
        }

        $pw[2 * $i] = chr(0x80);
        $pw[56] = chr(($i << 4) & 0xff);

        $mdContext1 = md5($pw);

        $offset = 0;
        $keyOffset = 0;
        $toCopy = 5;

        while ($offset != 16) {
            if ((64 - $offset) < 5) {
                $toCopy = 64 - $offset;
            }

            for ($i = 0; $i <= $toCopy; $i++) {
                $pw[$offset + $i] = $mdContext1[$keyOffset + $i];
            }

            $offset += $toCopy;

            if ($offset == 64) {
                $keyOffset = $toCopy;
                $toCopy = 5 - $toCopy;
                $offset = 0;
                continue;
            }

            $keyOffset = 0;
            $toCopy = 5;
            for ($i = 0; $i < 16; $i++) {
                $pw[$offset + $i] = $docid[$i];
            }
            $offset += 16;
        }

        $pw[16] = "\x80";
        for ($i = 0; $i < 47; $i++) {
            $pw[17 + $i] = "\0";
        }
        $pw[56] = "\x80";
        $pw[57] = "\x0a";

        $valContext = md5($pw);

        $key = $this->makeKey(0, $valContext);

        $salt = $key->RC4($salt_data);
        $hashedsalt = $key->RC4($hashedsalt_data);

        $salt .= "\x80" . str_repeat("\0", 47);
        $salt[56] = "\x80";

        $mdContext2 = md5($salt);

        return $mdContext2 == $hashedsalt;
    }

    /**
     * CODEPAGE
     *
     * This record stores the text encoding used to write byte strings, stored as MS Windows code page identifier.
     *
     * @throws ParserException
     */
    private function readCodepage() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; code page identifier
        $codePage = Format::getUInt2d($recordData, 0);
        $this->codePage = self::NumberToName($codePage);
    }

    /**
     * DATEMODE
     * This record specifies the base date for displaying date values. All dates are stored as count of days
     * past this base date. In BIFF2-BIFF4 this record is part of the Calculation Settings Block. In BIFF5-BIFF8
     * it is stored in the Workbook Globals Substream.
     */
    private function readDateMode() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; 0 = base 1900, 1 = base 1904
        self::$excelBaseDate = Format::CALENDAR_WINDOWS_1900;
        if (ord($recordData[0]) == 1) {
            self::$excelBaseDate = Format::CALENDAR_MAC_1904;
        }
    }

    /**
     * FORMAT
     *
     * This record contains information about a number format. All FORMAT records occur together in a sequential list.
     * In BIFF2-BIFF4 other records referencing a FORMAT record contain a zero-based index into this list. From BIFF5
     * on the FORMAT record contains the index itself that will be used by other records.
     */
    private function readFormat() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        $indexCode = Format::getUInt2d($recordData, 0);
        if ($this->version == self::XLS_BIFF8) {
            $string = self::readUnicodeStringLong(substr($recordData, 2));
        } else {
            // BIFF7
            $string = $this->readByteStringShort(substr($recordData, 2));
        }

        $formatString = $string['value'];
        $this->formats[$indexCode] = $formatString;
    }

    /**
     * XF - Extended Format
     *
     * This record contains formatting information for cells, rows, columns or styles.
     * According to http://support.microsoft.com/kb/147732 there are always at least 15 cell style XF and 1 cell XF.
     * Inspection of Excel files generated by MS Office Excel shows that XF records 0-14 are cell style XF and XF
     * record 15 is a cell XF. We only read the first cell style XF and skip the remaining cell style XF records
     * We read all cell XF records.
     */
    private function readXf() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 2; size: 2; Index to FORMAT record
        $numberFormatIndex = Format::getUInt2d($recordData, 2);
        if (isset($this->formats[$numberFormatIndex])) {
            // then we have user-defined format code
            $numberFormat = $this->formats[$numberFormatIndex];
        } elseif (isset(Format::$buildInFormats[$numberFormatIndex])) {
            // then we have built-in format code
            $numberFormat = Format::$buildInFormats[$numberFormatIndex];
        } else {
            // we set the general format code
            $numberFormat = Format::FORMAT_GENERAL;
        }

        $this->xfRecords[] = ['index' => $numberFormatIndex, 'format' => $numberFormat];
    }

    /**
     * SST - Shared String Table
     *
     * This record contains a list of all strings used anywhere in the workbook. Each string occurs only once.
     * The workbook uses indexes into the list to reference the strings.
     **/
    private function readSst() {
        // offset within (spliced) record data
        $pos = 0;

        // get spliced record data
        $splicedRecordData = $this->getSplicedRecordData();
        $recordData = $splicedRecordData['recordData'];
        $spliceOffsets = $splicedRecordData['spliceOffsets'];

        // offset: 0; size: 4; total number of strings in the workbook
        $pos += 4;

        // offset: 4; size: 4; number of following strings ($nm)
        $nm = Format::getInt4d($recordData, 4);

        $pos += 4;

        // loop through the Unicode strings (16-bit length)
        for ($i = 0; $i < $nm; ++$i) {
            if (!isset($recordData[$pos + 2])) {
                break;
            }

            // number of characters in the Unicode string
            $numChars = Format::getUInt2d($recordData, $pos);
            $pos += 2;

            // option flags
            $optionFlags = ord($recordData[$pos]);
            ++$pos;

            // bit: 0; mask: 0x01; 0 = compressed; 1 = uncompressed
            $isCompressed = (($optionFlags & 0x01) == 0) ;

            // bit: 2; mask: 0x02; 0 = ordinary; 1 = Asian phonetic
            $hasAsian = (($optionFlags & 0x04) != 0);

            // bit: 3; mask: 0x03; 0 = ordinary; 1 = Rich-Text
            $formattingRuns = 0;
            $hasRichText = (($optionFlags & 0x08) != 0);
            if ($hasRichText && isset($recordData[$pos])) {
                // number of Rich-Text formatting runs
                $formattingRuns = Format::getUInt2d($recordData, $pos);
                $pos += 2;
            }

            $extendedRunLength = 0;
            if ($hasAsian && isset($recordData[$pos])) {
                // size of Asian phonetic setting
                $extendedRunLength = Format::getInt4d($recordData, $pos);
                $pos += 4;
            }

            // expected byte length of character array if not split
            $len = ($isCompressed) ? $numChars : $numChars * 2;

            // look up limit position
            $limitPos = 0;
            foreach ($spliceOffsets as $spliceOffset) {
                // it can happen that the string is empty, therefore we need
                // <= and not just <
                if ($pos <= $spliceOffset) {
                    $limitPos = $spliceOffset;
                    break;
                }
            }

            if ($pos + $len <= $limitPos) {
                // character array is not split between records
                $retStr = substr($recordData, $pos, $len);
                $pos += $len;
            } else {
                // character array is split between records
                // first part of character array
                $retStr = substr($recordData, $pos, $limitPos - $pos);
                $bytesRead = $limitPos - $pos;

                // remaining characters in Unicode string
                $charsLeft = $numChars - (($isCompressed) ? $bytesRead : ($bytesRead / 2));
                $pos = $limitPos;

                // keep reading the characters
                while ($charsLeft > 0) {
                    // look up next limit position, in case the string span more than one continue record
                    foreach ($spliceOffsets as $spliceOffset) {
                        if ($pos < $spliceOffset) {
                            $limitPos = $spliceOffset;
                            break;
                        }
                    }

                    if (!isset($recordData[$pos])) {
                        break;
                    }

                    // repeated option flags
                    // OpenOffice.org documentation 5.21
                    $option = ord($recordData[$pos]);
                    ++$pos;

                    if ($isCompressed && ($option == 0)) {
                        // 1st fragment compressed
                        // this fragment compressed
                        $len = min($charsLeft, $limitPos - $pos);
                        $retStr .= substr($recordData, $pos, $len);
                        $charsLeft -= $len;
                        $isCompressed = true;
                    } elseif (!$isCompressed && ($option != 0)) {
                        // 1st fragment uncompressed
                        // this fragment uncompressed
                        $len = min($charsLeft * 2, $limitPos - $pos);
                        $retStr .= substr($recordData, $pos, $len);
                        $charsLeft -= $len / 2;
                        $isCompressed = false;
                    } elseif (!$isCompressed && ($option == 0)) {
                        // 1st fragment uncompressed
                        // this fragment compressed
                        $len = min($charsLeft, $limitPos - $pos);
                        for ($j = 0; $j < $len; ++$j) {
                            if (!isset($recordData[$pos + $j])) {
                                break;
                            }

                            $retStr .= $recordData[$pos + $j] . chr(0);
                        }

                        $charsLeft -= $len;
                        $isCompressed = false;
                    } else {
                        // 1st fragment compressed
                        // this fragment uncompressed
                        $newStr = '';
                        $jMax = strlen($retStr);
                        for ($j = 0; $j < $jMax; ++$j) {
                            $newStr .= $retStr[$j] . chr(0);
                        }

                        $retStr = $newStr;
                        $len = min($charsLeft * 2, $limitPos - $pos);
                        $retStr .= substr($recordData, $pos, $len);
                        $charsLeft -= $len / 2;
                        $isCompressed = false;
                    }

                    $pos += $len;
                }
            }

            // convert to UTF-8
            $retStr = self::encodeUTF16($retStr, $isCompressed);

            // read additional Rich-Text information, if any
            // $fmtRuns = [];
            if ($hasRichText) {
                // list of formatting runs
                /*for ($j = 0; $j < $formattingRuns; ++$j) {
                    // first formatted character; zero-based
                    $charPos = Format::getUInt2d($recordData, $pos + $j * 4);

                    // index to font record
                    $fontIndex = Format::getUInt2d($recordData, $pos + 2 + $j * 4);
                    $fmtRuns[] = ['charPos' => $charPos, 'fontIndex' => $fontIndex];
                }*/

                $pos += 4 * $formattingRuns;
            }

            // read additional Asian phonetics information, if any
            if ($hasAsian) {
                // For Asian phonetic settings, we skip the extended string data
                $pos += $extendedRunLength;
            }

            // store the shared sting
            $this->sst[] = ['value' => $retStr];
        }
    }

    /**
     * Read RK record
     *
     * This record represents a cell that contains an RK value (encoded integer or floating-point value). If a
     * floating-point value cannot be encoded to an RK value, a NUMBER record will be written. This record replaces
     * the record INTEGER written in BIFF2.
     */
    private function readRk() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; index to row
        $row = Format::getUInt2d($recordData, 0);

        // offset: 2; size: 2; index to column
        $column = Format::getUInt2d($recordData, 2);

        // offset: 4; size: 2; index to XF record
        $xfIndex = Format::getUInt2d($recordData, 4);

        // offset: 6; size: 4; RK value
        $rkNum = Format::getInt4d($recordData, 6);
        $numValue = self::getIEEE754($rkNum);

        // add cell
        $this->addCell($row, $column, $numValue, $xfIndex);
    }

    /**
     * Read LABELSST record
     *
     * This record represents a cell that contains a string. It replaces the LABEL record and RSTRING record used in
     * BIFF2-BIFF5.
     */
    private function readLabelSst() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        $this->pos += 4 + $length;
        $xfIndex = Format::getUInt2d($recordData, 4);
        $row = Format::getUInt2d($recordData, 0);
        $column = Format::getUInt2d($recordData, 2);

        // offset: 6; size: 4; index to SST record
        $index = Format::getInt4d($recordData, 6);
        $this->addCell($row, $column, $this->sst[$index]['value'], $xfIndex);
    }

    /**
     * Read MULRK record
     *
     * This record represents a cell range containing RK value cells. All cells are located in the same row.
     */
    private function readMulRk() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; index to row
        $row = Format::getUInt2d($recordData, 0);

        // offset: 2; size: 2; index to first column
        $colFirst = Format::getUInt2d($recordData, 2);

        // offset: var; size: 2; index to last column
        $colLast = Format::getUInt2d($recordData, $length - 2);
        $columns = $colLast - $colFirst + 1;

        // offset within record data
        $offset = 4;
        for ($i = 0; $i < $columns; ++$i) {
            // offset: var; size: 2; index to XF record
            $xfIndex = Format::getUInt2d($recordData, $offset);

            // offset: var; size: 4; RK value
            $numValue = self::getIEEE754(Format::getInt4d($recordData, $offset + 2));

            $this->addCell($row, $colFirst + $i, $numValue, $xfIndex);

            $offset += 6;
        }
    }

    /**
     * Read NUMBER record
     *
     * This record represents a cell that contains a floating-point value.
     */
    private function readNumber() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; index to row
        $row = Format::getUInt2d($recordData, 0);

        // offset: 2; size 2; index to column
        $column = Format::getUInt2d($recordData, 2);

        // offset 4; size: 2; index to XF record
        $xfIndex = Format::getUInt2d($recordData, 4);
        $numValue = self::extractNumber(substr($recordData, 6, 8));

        $this->addCell($row, $column, $numValue, $xfIndex);
    }

    /**
     * Read FORMULA record + perhaps a following STRING record if formula result is a string
     * This record contains the token array and the result of a formula cell.
     */
    private function readFormula() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; row index
        $row = Format::getUInt2d($recordData, 0);

        // offset: 2; size: 2; col index
        $column = Format::getUInt2d($recordData, 2);

        // offset 4; size: 2; index to XF record
        $xfIndex = Format::getUInt2d($recordData, 4);

        // offset: 6; size: 8; result of the formula
        if ((ord($recordData[6]) == 0) && (ord($recordData[12]) == 255) && (ord($recordData[13]) == 255)) {
            // read STRING record
            $value = $this->readString();
        } elseif ((ord($recordData[6]) == 1) && (ord($recordData[12]) == 255) && (ord($recordData[13]) == 255)) {
            // Boolean formula. Result is in +2; 0=false, 1=true
            $value = (bool) ord($recordData[8]);
        } elseif ((ord($recordData[6]) == 2) && (ord($recordData[12]) == 255) && (ord($recordData[13]) == 255)) {
            // Error formula. Error code is in +2
            $value = self::mapErrorCode(ord($recordData[8]));
        } elseif ((ord($recordData[6]) == 3) && (ord($recordData[12]) == 255) && (ord($recordData[13]) == 255)) {
            // Formula result is a null string
            $value = '';
        } else {
            // forumla result is a number, first 14 bytes like _NUMBER record
            $value = self::extractNumber(substr($recordData, 6, 8));
        }

        $this->addCell($row, $column, $value, $xfIndex);
    }

    /**
     * Read a STRING record from current stream position and advance the stream pointer to next record.
     * This record is used for storing result from FORMULA record when it is a string, and it occurs
     * directly after the FORMULA record
     *
     * @return string The string contents as UTF-8
     */
    private function readString() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;
        if ($this->version == self::XLS_BIFF8) {
            $string = self::readUnicodeStringLong($recordData);
            $value = $string['value'];
        } else {
            $string = $this->readByteStringLong($recordData);
            $value = $string['value'];
        }

        return $value;
    }

    /**
     * Read BOOLERR record
     *
     * This record represents a Boolean value or error value cell.
     */
    private function readBoolErr() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; row index
        $row = Format::getUInt2d($recordData, 0);

        // offset: 2; size: 2; column index
        $column = Format::getUInt2d($recordData, 2);

        // offset: 4; size: 2; index to XF record
        $xfIndex = Format::getUInt2d($recordData, 4);

        // offset: 6; size: 1; the boolean value or error value
        $boolError = ord($recordData[6]);

        // offset: 7; size: 1; 0=boolean; 1=error
        $isError = ord($recordData[7]);

        switch ($isError) {
            case 0: // boolean
                $value = (bool)$boolError;

                // add cell value
                $this->addCell($row, $column, $value, $xfIndex);
                break;
            case 1: // error type
                $value = self::mapErrorCode($boolError);

                // add cell value
                $this->addCell($row, $column, $value, $xfIndex);
                break;
        }
    }

    /**
     * Read BLANK record
     */
    private function readBlank() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; row index
        $row = Format::getUInt2d($recordData, 0);

        // offset: 2; size: 2; col index
        $column = Format::getUInt2d($recordData, 2);

        // offset: 4; size: 2; XF index
        $xfIndex = Format::getUInt2d($recordData, 4);

        $this->addCell($row, $column, '', $xfIndex);
    }

    /**
     * Read LABEL record
     *
     * This record represents a cell that contains a string. In BIFF8 it is usually replaced by the LABELSST record.
     * Excel still uses this record, if it copies unformatted text cells to the clipboard.
     */
    private function readLabel() {
        $length = Format::getUInt2d($this->data, $this->pos + 2);
        $recordData = substr($this->data, $this->pos + 4, $length);

        // move stream pointer to next record
        $this->pos += 4 + $length;

        // offset: 0; size: 2; index to row
        $row = Format::getUInt2d($recordData, 0);

        // offset: 2; size: 2; index to column
        $column = Format::getUInt2d($recordData, 2);

        // offset: 4; size: 2; XF index
        $xfIndex = Format::getUInt2d($recordData, 4);

        // add cell value
        if ($this->version == self::XLS_BIFF8) {
            $string = self::readUnicodeStringLong(substr($recordData, 6));
            $value = $string['value'];
        } else {
            $string = $this->readByteStringLong(substr($recordData, 6));
            $value = $string['value'];
        }

        $this->addCell($row, $column, $value, $xfIndex);
    }

    /**
     * Map error code, e.g. '#N/A'
     *
     * @param int $code
     * @return string
     */
    private static function mapErrorCode($code) {
        if (isset(self::$errorCode[$code])) {
            return self::$errorCode[$code];
        }

        return false;
    }

    /**
     * Convert a value in a pre-defined format to a PHP string
     *
     * @param mixed $value    Value to format
     * @param string $format  Format code
     * @return string
     */
    private static function toFormattedString($value = '0', $format = Format::FORMAT_GENERAL) {
        // For now we do not treat strings although section 4 of a format code affects strings
        if (!is_numeric($value)) {
            return $value;
        }

        // For 'General' format code, we just pass the value although this is not entirely the way Excel does it,
        // it seems to round numbers to a total of 10 digits.
        if (($format === Format::FORMAT_GENERAL) || ($format === Format::FORMAT_TEXT)) {
            return $value;
        }

        // Convert any other escaped characters to quoted strings, e.g. (\T to "T")
        $format = preg_replace('/(\\\(.))(?=(?:[^"]|"[^"]*")*$)/u', '"${2}"', $format);

        // Get the sections, there can be up to four sections, separated with a semi-colon (but only if not a quoted literal)
        $sections = preg_split('/(;)(?=(?:[^"]|"[^"]*")*$)/u', $format);

        // Extract the relevant section depending on whether number is positive, negative, or zero?
        // Text not supported yet.
        // Here is how the sections apply to various values in Excel:
        //   1 section:   [POSITIVE/NEGATIVE/ZERO/TEXT]
        //   2 sections:  [POSITIVE/ZERO/TEXT] [NEGATIVE]
        //   3 sections:  [POSITIVE/TEXT] [NEGATIVE] [ZERO]
        //   4 sections:  [POSITIVE] [NEGATIVE] [ZERO] [TEXT]
        switch (count($sections)) {
            case 1:
                $format = $sections[0];
                break;

            case 2:
                $format = ($value >= 0) ? $sections[0] : $sections[1];
                $value = abs($value); // Use the absolute value
                break;

            case 3:
                $format = ($value > 0) ? $sections[0] : ( ($value < 0) ? $sections[1] : $sections[2]);
                $value = abs($value); // Use the absolute value
                break;

            case 4:
                $format = ($value > 0) ? $sections[0] : ( ($value < 0) ? $sections[1] : $sections[2]);
                $value = abs($value); // Use the absolute value
                break;

            default:
                // something is wrong, just use first section
                $format = $sections[0];
                break;
        }

        // In Excel formats, "_" is used to add spacing,
        //    The following character indicates the size of the spacing, which we can't do in HTML, so we just use a standard space
        $format = preg_replace('/_./', ' ', $format);

        // Save format with color information for later use below
        //$formatColor = $format;

        // Strip color information
        $colorRegex = '/^\\[[a-zA-Z]+\\]/';
        $format = preg_replace($colorRegex, '', $format);

        // Let's begin inspecting the format and converting the value to a formatted string
        //  Check for date/time characters (not inside quotes)
        if (preg_match('/(\[\$[A-Z]*-[0-9A-F]*\])*[hmsdy](?=(?:[^"]|"[^"]*")*$)/miu', $format, $matches)) {
            // datetime format
            self::formatAsDate($value, $format);
        } elseif (preg_match('/%$/', $format)) {
            // % number format
            self::formatAsPercentage($value, $format);
        } else {
            if ($format === Format::FORMAT_CURRENCY_EUR_SIMPLE) {
                $value = 'EUR ' . sprintf('%1.2f', $value);
            } else {
                // Some non-number strings are quoted, so we'll get rid of the quotes, likewise any positional * symbols
                $format = str_replace(['"', '*'], '', $format);

                // Find out if we need thousands separator
                // This is indicated by a comma enclosed by a digit placeholder:
                // #,# or 0,0
                $useThousands = preg_match('/(#,#|0,0)/', $format);
                if ($useThousands) {
                    $format = preg_replace('/0,0/', '00', $format);
                    $format = preg_replace('/#,#/', '##', $format);
                }

                // Scale thousands, millions,...
                // This is indicated by a number of commas after a digit placeholder:
                // #, or 0.0,,
                $scale = 1; // same as no scale
                $matches = [];
                if (preg_match('/(#|0)(,+)/', $format, $matches)) {
                    $scale = pow(1000, strlen($matches[2]));

                    // strip the commas
                    $format = preg_replace('/0,+/', '0', $format);
                    $format = preg_replace('/#,+/', '#', $format);
                }

                if (preg_match('/#?.*\?\/\?/', $format, $m)) {
                    //echo 'Format mask is fractional '.$format.' <br />';
                    if ($value != (int)$value) {
                        self::formatAsFraction($value, $format);
                    }
                } else {
                    // Handle the number itself
                    // scale number
                    $value = $value / $scale;

                    // Strip #
                    $format = preg_replace('/\\#/', '0', $format);
                    $n = "/\[[^\]]+\]/";
                    $m = preg_replace($n, '', $format);
                    $numberRegex = "/(0+)(\.?)(0*)/";
                    if (preg_match($numberRegex, $m, $matches)) {
                        $left = $matches[1];
                        $dec = $matches[2];
                        $right = $matches[3];

                        // minimun width of formatted number (including dot)
                        $minWidth = strlen($left) + strlen($dec) + strlen($right);
                        if ($useThousands) {
                            $value = number_format(
                                $value,
                                strlen($right),
                                self::getDecimalSeparator(),
                                self::getThousandsSeparator()
                            );

                            $value = preg_replace($numberRegex, $value, $format);
                        } else {
                            if (preg_match('/[0#]E[+-]0/i', $format)) {
                                //Scientific format
                                $value = sprintf('%5.2E', $value);
                            } elseif (preg_match('/0([^\d\.]+)0/', $format)) {
                                $value = self::complexNumberFormatMask($value, $format);
                            } else {
                                $sprintfPattern = "%0$minWidth." . strlen($right) . "f";
                                $value = sprintf($sprintfPattern, $value);
                                $value = preg_replace($numberRegex, $value, $format);
                            }
                        }
                    }
                }

                if (preg_match('/\[\$(.*)\]/u', $format, $m)) {
                    //  Currency or Accounting
                    //$currencyFormat = $m[0];
                    $currencyCode = $m[1];
                    list($currencyCode) = explode('-', $currencyCode);

                    if ($currencyCode == '') {
                        $currencyCode = self::getCurrencyCode();
                    }

                    $value = preg_replace('/\[\$([^\]]*)\]/u', $currencyCode, $value);
                }
            }
        }

        return $value;
    }

    /**
     * Reads a record from current position in data stream and continues reading data as long as CONTINUE records
     * are found. Splices the record data pieces and returns the combined string as if record data is in one piece.
     * Moves to next current position in data stream to start of next record different from a CONtINUE record
     *
     * @return array
     */
    private function getSplicedRecordData() {
        $i = 0;
        $data = '';
        $spliceOffsets = [0];

        do {
            ++$i;
            // offset: 0; size: 2; identifier
            //$identifier = Cell::getInt2d($this->data, $this->pos);

            // offset: 2; size: 2; length
            $length = Format::getUInt2d($this->data, $this->pos + 2);
            $data .= substr($this->data, $this->pos + 4, $length);
            $spliceOffsets[$i] = $spliceOffsets[$i - 1] + $length;

            $this->pos += 4 + $length;
            $nextIdentifier = Format::getUInt2d($this->data, $this->pos);
        } while ($nextIdentifier == self::XLS_TYPE_CONTINUE);

        return ['recordData' => $data, 'spliceOffsets' => $spliceOffsets];
    }

    /**
     * Get the decimal separator. If it has not yet been set explicitly, try to obtain number formatting
     * information from locale.
     *
     * @return string
     */
    private static function getDecimalSeparator() {
        if (!isset(self::$decimalSeparator)) {
            $localeconv = localeconv();

            self::$decimalSeparator = ($localeconv['decimal_point'] != '') ? $localeconv['decimal_point']
                : $localeconv['mon_decimal_point'];

            if (self::$decimalSeparator == '') {
                // Default to .
                self::$decimalSeparator = '.';
            }
        }

        return self::$decimalSeparator;
    }

    /**
     * Get the thousands separator. If it has not yet been set explicitly, try to obtain number formatting
     * information from locale.
     *
     * @return string
     */
    private static function getThousandsSeparator() {
        if (!isset(self::$thousandsSeparator)) {
            $localeconv = localeconv();

            self::$thousandsSeparator = ($localeconv['thousands_sep'] != '') ? $localeconv['thousands_sep']
                : $localeconv['mon_thousands_sep'];

            if (self::$thousandsSeparator == '') {
                // Default to .
                self::$thousandsSeparator = ',';
            }
        }

        return self::$thousandsSeparator;
    }

    /**
     * Get the currency code. If it has not yet been set explicitly, try to obtain the symbol information from locale.
     *
     * @return string
     */
    private static function getCurrencyCode() {
        if (!isset(self::$currencyCode)) {
            $localeconv = localeconv();

            self::$currencyCode = ($localeconv['currency_symbol'] != '') ? $localeconv['currency_symbol']
                : $localeconv['int_curr_symbol'];

            if (self::$currencyCode == '') {
                // Default to $
                self::$currencyCode = '$';
            }
        }

        return self::$currencyCode;
    }

    private static function complexNumberFormatMask($number, $mask) {
        $sign = ($number < 0.0);
        $number = abs($number);

        if (strpos($mask, '.') !== false) {
            $numbers = explode('.', $number . '.0');
            $masks = explode('.', $mask . '.0');
            $result1 = self::complexNumberFormatMask($numbers[0], $masks[0]);
            $result2 = strrev(self::complexNumberFormatMask(strrev($numbers[1]), strrev($masks[1])));

            return (($sign) ? '-' : '') . $result1 . '.' . $result2;
        }

        $r = preg_match_all('/0+/', $mask, $result, PREG_OFFSET_CAPTURE);
        if ($r > 1) {
            $result = array_reverse($result[0]);

            $offset = 0;
            foreach ($result as $block) {
                $divisor = 1 . $block[0];
                $size = strlen($block[0]);
                $offset = $block[1];
                $blockValue = sprintf('%0' . $size . 'd', fmod($number, $divisor));

                $number = floor($number / $divisor);
                $mask = substr_replace($mask, $blockValue, $offset, $size);
            }

            if ($number > 0) {
                $mask = substr_replace($mask, $number, $offset, 0);
            }

            $result = $mask;
        } else {
            $result = $number;
        }

        return (($sign) ? '-' : '') . $result;
    }

    /**
     * Convert Microsoft Code Page Identifier to Code Page Name which iconv and mbstring understands
     *
     * @param int $codePage Microsoft Code Page Indentifier
     *
     * @throws ParserException
     * @return string Code Page Name
     */
    private static function NumberToName($codePage = 1252) {
        switch ($codePage) {
            case 367:
                return 'ASCII'; //ASCII

            case 437:
                return 'CP437'; //OEM US

            case 720:
                throw new ParserException('Code page 720 not supported.', 5); //OEM Arabic

            case 737:
                return 'CP737'; //OEM Greek

            case 775:
                return 'CP775'; //OEM Baltic

            case 850:
                return 'CP850'; //OEM Latin I

            case 852:
                return 'CP852'; //OEM Latin II (Central European)

            case 855:
                return 'CP855'; //OEM Cyrillic

            case 857:
                return 'CP857'; //OEM Turkish

            case 858:
                return 'CP858'; //OEM Multilingual Latin I with Euro

            case 860:
                return 'CP860'; //OEM Portugese

            case 861:
                return 'CP861'; //OEM Icelandic

            case 862:
                return 'CP862'; //OEM Hebrew

            case 863:
                return 'CP863'; //OEM Canadian (French)

            case 864:
                return 'CP864'; //OEM Arabic

            case 865:
                return 'CP865'; //OEM Nordic

            case 866:
                return 'CP866'; //OEM Cyrillic (Russian)

            case 869:
                return 'CP869'; //OEM Greek (Modern)

            case 874:
                return 'CP874'; //ANSI Thai

            case 932:
                return 'CP932'; //ANSI Japanese Shift-JIS

            case 936:
                return 'CP936'; //ANSI Chinese Simplified GBK

            case 949:
                return 'CP949'; //ANSI Korean (Wansung)

            case 950:
                return 'CP950'; //ANSI Chinese Traditional BIG5

            case 1200:
                return 'UTF-16LE'; //UTF-16 (BIFF8)

            case 1250:
                return 'CP1250'; //ANSI Latin II (Central European)

            case 1251:
                return 'CP1251'; //ANSI Cyrillic

            case 0: //CodePage is not always correctly set when the xls file was saved by Apple's Numbers program
            case 1252:
                return 'CP1252'; //ANSI Latin I (BIFF4-BIFF7)

            case 1253:
                return 'CP1253'; //ANSI Greek

            case 1254:
                return 'CP1254'; //ANSI Turkish

            case 1255:
                return 'CP1255'; //ANSI Hebrew

            case 1256:
                return 'CP1256'; //ANSI Arabic

            case 1257:
                return 'CP1257'; //ANSI Baltic

            case 1258:
                return 'CP1258'; //ANSI Vietnamese

            case 1361:
                return 'CP1361'; //ANSI Korean (Johab)

            case 10000:
                return 'MAC'; //Apple Roman

            case 10001:
                return 'CP932'; //Macintosh Japanese

            case 10002:
                return 'CP950'; //Macintosh Chinese Traditional

            case 10003:
                return 'CP1361'; //Macintosh Korean

            case 10004:
                return 'MACARABIC'; //	Apple Arabic

            case 10005:
                return 'MACHEBREW'; //Apple Hebrew

            case 10006:
                return 'MACGREEK'; //Macintosh Greek

            case 10007:
                return 'MACCYRILLIC'; //Macintosh Cyrillic

            case 10008:
                return 'CP936'; //Macintosh - Simplified Chinese (GB 2312)

            case 10010:
                return 'MACROMANIA'; //Macintosh Romania

            case 10017:
                return 'MACUKRAINE'; //Macintosh Ukraine

            case 10021:
                return 'MACTHAI'; //Macintosh Thai

            case 10029:
                return 'MACCENTRALEUROPE'; //Macintosh Central Europe

            case 10079:
                return 'MACICELAND'; //Macintosh Icelandic

            case 10081:
                return 'MACTURKISH'; //Macintosh Turkish

            case 10082:
                return 'MACCROATIAN'; //Macintosh Croatian

            case 21010:
                return 'UTF-16LE'; //UTF-16 (BIFF8) This isn't correct, but some Excel writer libraries erroneously
                                   // use Codepage 21010 for UTF-16LE

            case 32768:
                return 'MAC'; //Apple Roman

            case 32769:
                throw new ParserException('Code page 32769 not supported.', 6); //ANSI Latin I (BIFF2-BIFF3)

            case 65000:
                return 'UTF-7'; //Unicode (UTF-7)

            case 65001:
                return 'UTF-8'; //Unicode (UTF-8)
        }

        throw new ParserException("Unknown codepage: $codePage", 7);
    }

    /**
     * Read byte string (8-bit string length). OpenOffice documentation: 2.5.2
     *
     * @param string $subData
     *
     * @return array
     */
    private function readByteStringShort($subData) {
        // offset: 0; size: 1; length of the string (character count)
        $ln = ord($subData[0]);

        // offset: 1: size: var; character array (8-bit characters)
        $value = $this->decodeCodepage(substr($subData, 1, $ln));

        // size in bytes of data structure
        return ['value' => $value, 'size' => 1 + $ln];
    }

    /**
     * Read byte string (16-bit string length). OpenOffice documentation: 2.5.2
     *
     * @param string $subData
     * @return array
     */
    private function readByteStringLong($subData) {
        // offset: 0; size: 2; length of the string (character count)
        $ln = Format::getUInt2d($subData, 0);

        // offset: 2: size: var; character array (8-bit characters)
        $value = $this->decodeCodepage(substr($subData, 2));

        // size in bytes of data structure
        return ['value' => $value, 'size' => 2 + $ln];
    }

    private static function formatAsDate(&$value, &$format) {
        // strip off first part containing e.g. [$-F800] or [$USD-409]
        // general syntax: [$<Currency string>-<language info>]
        // language info is in hexadecimal
        $format = preg_replace('/^(\[\$[A-Z]*-[0-9A-F]*\])/i', '', $format);

        // OpenOffice.org uses upper-case number formats, e.g. 'YYYY', convert to lower-case;
        // but we don't want to change any quoted strings
        $format = preg_replace_callback('/(?:^|")([^"]*)(?:$|")/', ['self', 'setLowercaseCallback'], $format);

        // Only process the non-quoted blocks for date format characters
        $blocks = explode('"', $format);

        foreach($blocks as $key => &$block) {
            if ($key % 2 == 0) {
                $block = strtr($block, Format::$dateFormatReplacements);
                if (strpos($block, 'A') === false) {
                    // 24-hour time format
                    $block = strtr($block, Format::$dateFormatReplacements24);
                } else {
                    // 12-hour time format
                    $block = strtr($block, Format::$dateFormatReplacements12);
                }
            }
        }

        $format = implode('"', $blocks);

        // escape any quoted characters so that DateTime format() will render them correctly
        $format = preg_replace_callback('/"(.*)"/U', ['self', 'escapeQuotesCallback'], $format);
        $dateObj = self::ExcelToPHPObject($value);

        $value = $dateObj->format($format);
    }

    private static function setLowercaseCallback($matches) {
        return mb_strtolower($matches[0]);
    }

    private static function escapeQuotesCallback($matches) {
        return '\\' . implode('\\', str_split($matches[1]));
    }

    /**
     * Convert a date from Excel to a PHP Date/Time object
     *
     * @param int $dateValue Excel date/time value
     *
     * @return \DateTime PHP date/time object
     */
    private static function ExcelToPHPObject($dateValue = 0) {
        $dateTime = self::ExcelToPHP($dateValue);

        $days = floor($dateTime / 86400);
        $time = round((($dateTime / 86400) - $days) * 86400);
        $hours = round($time / 3600);
        $minutes = round($time / 60) - ($hours * 60);
        $seconds = round($time) - ($hours * 3600) - ($minutes * 60);

        $dateObj = new \DateTime("1-Jan-1970+$days days");
        $dateObj->setTime($hours, $minutes, $seconds);

        return $dateObj;
    }

    /**
     * Convert a date from Excel to PHP
     *
     * @param int $dateValue Excel date/time value
     *
     * @return int PHP serialized date/time
     */
    private static function ExcelToPHP($dateValue = 0) {
        if (self::$excelBaseDate == Format::CALENDAR_WINDOWS_1900) {
            $excelBaseDate = 25569;

            //Adjust for the spurious 29-Feb-1900 (Day 60)
            if ($dateValue < 60) {
                --$excelBaseDate;
            }
        } else {
            $excelBaseDate = 24107;
        }

        // Perform conversion
        if ($dateValue >= 1) {
            $utcDays = $dateValue - $excelBaseDate;
            $returnValue = round($utcDays * 86400);

            if (($returnValue <= PHP_INT_MAX) && ($returnValue >= -PHP_INT_MAX)) {
                $returnValue = (integer) $returnValue;
            }
        } else {
            $hours = round($dateValue * 24);
            $mins = round($dateValue * 1440) - round($hours * 60);
            $secs = round($dateValue * 86400) - round($hours * 3600) - round($mins * 60);

            $returnValue = (integer) gmmktime($hours, $mins, $secs);
        }

        return $returnValue;
    }

    private static function formatAsPercentage(&$value, &$format) {
        if ($format === Format::FORMAT_PERCENTAGE) {
            $value = round((100 * $value), 0) . '%';
        } else {
            if (preg_match('/\.[#0]+/i', $format, $m)) {
                $s = substr($m[0], 0, 1) . (strlen($m[0]) - 1);
                $format = str_replace($m[0], $s, $format);
            }

            if (preg_match('/^[#0]+/', $format, $m)) {
                $format = str_replace($m[0], strlen($m[0]), $format);
            }

            $format = '%' . str_replace('%', 'f%%', $format);
            $value = sprintf($format, 100 * $value);
        }
    }

    private static function formatAsFraction(&$value, &$format) {
        $sign = ($value < 0) ? '-' : '';
        $integerPart = floor(abs($value));
        $decimalPart = trim(fmod(abs($value), 1), '0.');
        $decimalLength = strlen($decimalPart);
        $decimalDivisor = pow(10, $decimalLength);

        $GCD = self::GCD([$decimalPart, $decimalDivisor]);
        $adjustedDecimalPart = $decimalPart/$GCD;
        $adjustedDecimalDivisor = $decimalDivisor/$GCD;

        if ((strpos($format, '0') !== false) || (strpos($format, '#') !== false) || (substr($format, 0, 3) == '? ?')) {
            if ($integerPart == 0) {
                $integerPart = '';
            }

            $value = "$sign$integerPart $adjustedDecimalPart/$adjustedDecimalDivisor";
        } else {
            $adjustedDecimalPart += $integerPart * $adjustedDecimalDivisor;
            $value = "$sign$adjustedDecimalPart/$adjustedDecimalDivisor";
        }
    }

    /**
     * GCD
     *
     * Returns the greatest common divisor of a series of numbers. The greatest common divisor is the largest
     * integer that divides both number1 and number2 without a remainder.
     * Excel Function:
     *     GCD(number1[,number2[, ...]])
     *
     * @param array $params
     *
     * @return integer Greatest Common Divisor
     */
    private static function GCD($params) {
        $returnValue = 1;
        $allValuesFactors = [];

        // Loop through arguments
        $flattenArr = self::flattenArray($params);
        foreach ($flattenArr as $value) {
            if (!is_numeric($value)) {
                return '#VALUE!';
            } elseif ($value == 0) {
                continue;
            } elseif ($value < 0) {
                return '#NULL!';
            }

            $factors = self::factors($value);
            $countedFactors = array_count_values($factors);
            $allValuesFactors[] = $countedFactors;
        }

        $allValuesCount = count($allValuesFactors);
        if ($allValuesCount == 0) {
            return 0;
        }

        $mergedArray = $allValuesFactors[0];
        for ($i=1; $i < $allValuesCount; ++$i) {
            $mergedArray = array_intersect_key($mergedArray, $allValuesFactors[$i]);
        }

        $mergedArrayValues = count($mergedArray);

        if ($mergedArrayValues == 0) {
            return $returnValue;
        } elseif ($mergedArrayValues > 1) {
            foreach ($mergedArray as $mergedKey => $mergedValue) {
                foreach ($allValuesFactors as $highestPowerTest) {
                    foreach ($highestPowerTest as $testKey => $testValue) {
                        if (($testKey == $mergedKey) && ($testValue < $mergedValue)) {
                            $mergedArray[$mergedKey] = $testValue;
                            $mergedValue = $testValue;
                        }
                    }
                }
            }

            $returnValue = 1;
            foreach ($mergedArray as $key => $value) {
                $returnValue *= pow($key, $value);
            }

            return $returnValue;
        } else {
            $keys = array_keys($mergedArray);
            $key = $keys[0];
            $value = $mergedArray[$key];

            foreach ($allValuesFactors as $testValue) {
                foreach ($testValue as $mergedKey => $mergedValue) {
                    if (($mergedKey == $key) && ($mergedValue < $value)) {
                        $value = $mergedValue;
                    }
                }
            }

            return pow($key, $value);
        }
    }

    /**
     * Convert a multi-dimensional array to a simple 1-dimensional array
     *
     * @param array $array Array to be flattened
     *
     * @return array Flattened array
     */
    private static function flattenArray($array) {
        if (!is_array($array)) {
            return (array) $array;
        }

        $arrayValues = [];
        foreach ($array as $value) {
            if (is_array($value)) {
                foreach ($value as $val) {
                    if (is_array($val)) {
                        foreach ($val as $v) {
                            $arrayValues[] = $v;
                        }
                    } else {
                        $arrayValues[] = $val;
                    }
                }
            } else {
                $arrayValues[] = $value;
            }
        }

        return $arrayValues;
    }

    /**
     * Return an array of the factors of the input value
     *
     * @param int $value
     *
     * @return array
     */
    private static function factors($value) {
        $startVal = floor(sqrt($value));
        $factorArray = [];

        for ($i = $startVal; $i > 1; --$i) {
            if (($value % $i) == 0) {
                $factorArray = array_merge($factorArray, self::factors($value / $i));
                $factorArray = array_merge($factorArray, self::factors($i));

                if ($i <= sqrt($value)) {
                    break;
                }
            }
        }

        if (!empty($factorArray)) {
            rsort($factorArray);

            return $factorArray;
        }

        return [(int) $value];
    }

    /**
     * Read Unicode string with no string length field, but with known character count this function is under
     * construction, needs to support rich text, and Asian phonetic settings
     *
     * @param string $subData
     * @param int $characterCount
     *
     * @return array
     */
    private static function readUnicodeString($subData, $characterCount) {
        // offset: 0: size: 1; option flags
        // bit: 0; mask: 0x01; character compression (0 = compressed 8-bit, 1 = uncompressed 16-bit)
        $isCompressed = !((0x01 & ord($subData[0])) >> 0);

        // offset: 1: size: var; character array
        // this offset assumes richtext and Asian phonetic settings are off which is generally wrong
        // needs to be fixed
        $value = self::encodeUTF16(
            substr($subData, 1, $isCompressed ? $characterCount : 2 * $characterCount), $isCompressed
        );

        // the size in bytes including the option flags
        return ['value' => $value, 'size' => $isCompressed ? 1 + $characterCount : 1 + 2 * $characterCount];
    }

    /**
     * Extracts an Excel Unicode short string (8-bit string length), this function will automatically find out
     * where the Unicode string ends.
     *
     * @param string $subData
     *
     * @return array
     */
    private static function readUnicodeStringShort($subData) {
        // offset: 0: size: 1; length of the string (character count)
        $characterCount = ord($subData[0]);
        $string = self::readUnicodeString(substr($subData, 1), $characterCount);

        // add 1 for the string length
        $string['size'] += 1;

        return $string;
    }

    /**
     * Extracts an Excel Unicode long string (16-bit string length), this function is under construction,
     * needs to support rich text, and Asian phonetic settings
     *
     * @param string $subData
     *
     * @return array
     */
    private static function readUnicodeStringLong($subData) {
        // offset: 0: size: 2; length of the string (character count)
        $characterCount = Format::getUInt2d($subData, 0);
        $string = self::readUnicodeString(substr($subData, 2), $characterCount);

        // add 2 for the string length
        $string['size'] += 2;

        return $string;
    }

    private static function getIEEE754($rkNum) {
        if (($rkNum & 0x02) != 0) {
            $value = $rkNum >> 2;
        } else {
            // changes by mmp, info on IEEE754 encoding from
            // research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
            // The RK format calls for using only the most significant 30 bits of the 64 bit floating point value.
            // The other 34 bits are assumed to be 0 so we use the upper 30 bits of $rknum as follows...
            $sign = ($rkNum & 0x80000000) >> 31;
            $exp = ($rkNum & 0x7ff00000) >> 20;

            $mantissa = (0x100000 | ($rkNum & 0x000ffffc));
            $value = $mantissa / pow(2, (20- ($exp - 1023)));

            if ($sign) {
                $value = -1 * $value;
            }
            //end of changes by mmp
        }

        if (($rkNum & 0x01) != 0) {
            $value /= 100;
        }

        return $value;
    }

    /**
     * Get UTF-8 string from (compressed or uncompressed) UTF-16 string
     *
     * @param string $string
     * @param bool $compressed
     *
     * @return string
     */
    private static function encodeUTF16($string, $compressed = false) {
        if ($compressed) {
            $string = self::uncompressByteString($string);
        }

        return mb_convert_encoding($string, 'UTF-8', 'UTF-16LE');
    }

    /**
     * Convert string to UTF-8. Only used for BIFF5.
     *
     * @param string $string
     *
     * @return string
     */
    private function decodeCodepage($string) {
        return mb_convert_encoding($string, 'UTF-8', $this->codePage);
    }

    /**
     * Convert UTF-16 string in compressed notation to uncompressed form. Only used for BIFF8.
     *
     * @param string $string
     *
     * @return string
     */
    private static function uncompressByteString($string) {
        $uncompressedString = '';
        $strLen = strlen($string);

        for ($i = 0; $i < $strLen; ++$i) {
            $uncompressedString .= $string[$i] . "\0";
        }

        return $uncompressedString;
    }

    /**
     * Reads first 8 bytes of a string and return IEEE 754 float
     *
     * @param string $data Binary string that is at least 8 bytes long
     *
     * @return float
     */
    private static function extractNumber($data) {
        $rkNumHigh = Format::getInt4d($data, 4);
        $rkNumLow = Format::getInt4d($data, 0);

        $sign = ($rkNumHigh & 0x80000000) >> 31;
        $exp = (($rkNumHigh & 0x7ff00000) >> 20) - 1023;
        $mantissa = (0x100000 | ($rkNumHigh & 0x000fffff));

        $mantissaLow1 = ($rkNumLow & 0x80000000) >> 31;
        $mantissaLow2 = ($rkNumLow & 0x7fffffff);
        $value = $mantissa / pow(2, (20 - $exp));

        if ($mantissaLow1 != 0) {
            $value += 1 / pow(2, (21 - $exp));
        }

        $value += $mantissaLow2 / pow(2, (52 - $exp));

        if ($sign) {
            $value *= -1;
        }

        return $value;
    }
}
