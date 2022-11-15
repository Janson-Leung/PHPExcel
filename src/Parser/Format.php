<?php
/**
 * Format helper
 *
 * @author Janson
 * @create 2017-11-27
 */
namespace Asan\PHPExcel\Parser;

use Asan\PHPExcel\Exception\ParserException;

class Format {
    //Base date of 1st Jan 1900 = 1.0
    const CALENDAR_WINDOWS_1900 = 1900;

    //Base date of 2nd Jan 1904 = 1.0
    const CALENDAR_MAC_1904 = 1904;

    // Pre-defined formats
    const FORMAT_GENERAL = 'General';
    const FORMAT_TEXT = '@';

    const FORMAT_PERCENTAGE = '0%';
    const FORMAT_PERCENTAGE_00 = '0.00%';
    const FORMAT_CURRENCY_EUR_SIMPLE = '[$EUR ]#,##0.00_-';

    public static $buildInFormats = [
        0 => self::FORMAT_GENERAL,
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',
        5 => '"$"#,##0_),("$"#,##0)',
        6 => '"$"#,##0_),[Red]("$"#,##0)',
        7 => '"$"#,##0.00_),("$"#,##0.00)',
        8 => '"$"#,##0.00_),[Red]("$"#,##0.00)',
        9 => '0%',
        10 => '0.00%',
        //11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'yyyy/m/d',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'yyyy/m/d h:mm',

        // 补充
        28 => 'm月d日',
        31 => 'yyyy年m月d日',
        32 => 'h时i分',
        33 => 'h时i分ss秒',
        34 => 'AM/PM h时i分',
        35 => 'AM/PM h时i分ss秒',
        55 => 'AM/PM h时i分',
        56 => 'AM/PM h时i分ss秒',
        58 => 'm月d日',

        37 => '#,##0_),(#,##0)',
        38 => '#,##0_),[Red](#,##0)',
        39 => '#,##0.00_),(#,##0.00)',
        40 => '#,##0.00_),[Red](#,##0.00)',
        41 => '_("$"* #,##0_),_("$"* (#,##0),_("$"* "-"_),_(@_)',
        42 => '_(* #,##0_),_(* (#,##0),_(* "-"_),_(@_)',
        43 => '_(* #,##0.00_),_(* (#,##0.00),_(* "-"??_),_(@_)',
        44 => '_("$"* #,##0.00_),_("$"* \(#,##0.00\),_("$"* "-"??_),_(@_)',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mm:ss.0',
        48 => '##0.0E+0',
        49 => '@',

        // CHT
        27 => 'yyyy年m月',
        30 => 'm/d/yy',
        36 => '[$-404]e/m/d',
        50 => '[$-404]e/m/d',
        57 => 'yyyy年m月',

        // THA
        59 => 't0',
        60 => 't0.00',
        61 => 't#,##0',
        62 => 't#,##0.00',
        67 => 't0%',
        68 => 't0.00%',
        69 => 't# ?/?',
        70 => 't# ??/??'
    ];

    /**
     * Search/replace values to convert Excel date/time format masks to PHP format masks
     *
     * @var array
     */
    public static $dateFormatReplacements = [
        // first remove escapes related to non-format characters
        '\\' => '',

        // 12-hour suffix
        'am/pm' => 'A',

        // 2-digit year
        'e' => 'Y',
        'yyyy' => 'Y',
        'yy' => 'y',

        // first letter of month - no php equivalent
        'mmmmm' => 'M',

        // full month name
        'mmmm' => 'F',

        // short month name
        'mmm' => 'M',

        // mm is minutes if time, but can also be month w/leading zero
        // so we try to identify times be the inclusion of a : separator in the mask
        // It isn't perfect, but the best way I know how
        ':mm' => ':i',
        'mm:' => 'i:',

        // month leading zero
        'mm' => 'm',
        'm' => 'n',

        // full day of week name
        'dddd' => 'l',

        // short day of week name
        'ddd' => 'D',

        // days leading zero
        'dd' => 'd',
        'd' => 'j',

        // seconds
        'ss' => 's',

        // fractional seconds - no php equivalent
        '.s' => ''
    ];

    /**
     * Search/replace values to convert Excel date/time format masks hours to PHP format masks (24 hr clock)
     *
     * @var array
     */
    public static $dateFormatReplacements24 = [
        'hh' => 'H',
        'h'  => 'G'
    ];

    /**
     * Search/replace values to convert Excel date/time format masks hours to PHP format masks (12 hr clock)
     *
     * @var array
     */
    public static $dateFormatReplacements12 = [
        'hh' => 'h',
        'h'  => 'g'
    ];

    /**
     * Column index from string
     *
     * @param string $label
     *
     * @throws \Exception
     * @return int
     */
    public static function columnIndexFromString($label = 'A') {
        // Using a lookup cache adds a slight memory overhead, but boosts speed
        // caching using a static within the method is faster than a class static,
        // though it's additional memory overhead
        static $indexCache = [];

        if (isset($indexCache[$label])) {
            return $indexCache[$label];
        }

        // It's surprising how costly the strtoupper() and ord() calls actually are, so we use a lookup array rather
        // than use ord() and make it case insensitive to get rid of the strtoupper() as well. Because it's a static,
        // there's no significant memory overhead either
        static $columnLookup = [
            'A' => 1, 'B' => 2, 'C' => 3, 'D' => 4, 'E' => 5, 'F' => 6, 'G' => 7, 'H' => 8, 'I' => 9, 'J' => 10,
            'K' => 11, 'L' => 12, 'M' => 13, 'N' => 14, 'O' => 15, 'P' => 16, 'Q' => 17, 'R' => 18, 'S' => 19,
            'T' => 20, 'U' => 21, 'V' => 22, 'W' => 23, 'X' => 24, 'Y' => 25, 'Z' => 26, 'a' => 1, 'b' => 2, 'c' => 3,
            'd' => 4, 'e' => 5, 'f' => 6, 'g' => 7, 'h' => 8, 'i' => 9, 'j' => 10, 'k' => 11, 'l' => 12, 'm' => 13,
            'n' => 14, 'o' => 15, 'p' => 16, 'q' => 17, 'r' => 18, 's' => 19, 't' => 20, 'u' => 21, 'v' => 22,
            'w' => 23, 'x' => 24, 'y' => 25, 'z' => 26
        ];

        // We also use the language construct isset() rather than the more costly strlen() function to match the length
        // of $pString for improved performance
        if (!isset($indexCache[$label])) {
            if (!isset($label[0]) || isset($label[3])) {
                throw new ParserException('Column string can not be empty or longer than 3 characters');
            }

            if (!isset($label[1])) {
                $indexCache[$label] = $columnLookup[$label];
            } elseif (!isset($label[2])) {
                $indexCache[$label] = $columnLookup[$label[0]] * 26 + $columnLookup[$label[1]];
            } else {
                $indexCache[$label] = $columnLookup[$label[0]] * 676 + $columnLookup[$label[1]] * 26
                    + $columnLookup[$label[2]];
            }
        }

        return $indexCache[$label];
    }

    /**
     * String from columnindex
     *
     * @param int $column
     * @return string
     */
    public static function stringFromColumnIndex($column = 0) {
        // Using a lookup cache adds a slight memory overhead, but boosts speed
        // caching using a static within the method is faster than a class static,
        // though it's additional memory overhead
        static $stringCache = [];

        if (!isset($stringCache[$column])) {
            // Determine column string
            if ($column < 26) {
                $stringCache[$column] = chr(65 + $column);
            } elseif ($column < 702) {
                $stringCache[$column] = chr(64 + ($column / 26)) . chr(65 + $column % 26);
            } else {
                $stringCache[$column] = chr(64 + (($column - 26) / 676)) . chr(65 + ((($column - 26) % 676) / 26))
                    . chr(65 + $column % 26);
            }
        }

        return $stringCache[$column];
    }

    /**
     * Read 16-bit unsigned integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getUInt2d($data, $pos) {
        return ord($data[$pos]) | (ord($data[$pos + 1]) << 8);
    }

    /**
     * Read 32-bit signed integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt4d($data, $pos) {
        // FIX: represent numbers correctly on 64-bit system
        // http://sourceforge.net/tracker/index.php?func=detail&aid=1487372&group_id=99160&atid=623334
        // Hacked by Andreas Rehm 2006 to ensure correct result of the <<24 block on 32 and 64bit systems
        $ord24 = ord($data[$pos + 3]);

        if ($ord24 >= 128) {
            // negative number
            $ord24 = -abs((256 - $ord24) << 24);
        } else {
            $ord24 = ($ord24 & 127) << 24;
        }

        return ord($data[$pos]) | (ord($data[$pos + 1]) << 8) | (ord($data[$pos + 2]) << 16) | $ord24;
    }
}
