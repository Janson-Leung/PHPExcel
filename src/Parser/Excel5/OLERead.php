<?php
/**
 * OLE File Read
 *
 * @author Janson
 * @create 2017-11-27
 */
namespace Asan\PHPExcel\Parser\Excel5;

use Asan\PHPExcel\Exception\ParserException;
use Asan\PHPExcel\Exception\ReaderException;
use Asan\PHPExcel\Parser\Format;

defined('IDENTIFIER_OLE') ||
    define('IDENTIFIER_OLE', pack('CCCCCCCC', 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1));

class OLERead {
    // OLE identifier
    const IDENTIFIER_OLE = IDENTIFIER_OLE;

    // Size of a sector = 512 bytes
    const BIG_BLOCK_SIZE = 0x200;

    // Size of a short sector = 64 bytes
    const SMALL_BLOCK_SIZE = 0x40;

    // Size of a directory entry always = 128 bytes
    const PROPERTY_STORAGE_BLOCK_SIZE = 0x80;

    // Minimum size of a standard stream = 4096 bytes, streams smaller than this are stored as short streams
    const SMALL_BLOCK_THRESHOLD = 0x1000;

    // header offsets
    const NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2c;
    const ROOT_START_BLOCK_POS = 0x30;
    const SMALL_BLOCK_DEPOT_BLOCK_POS = 0x3c;
    const EXTENSION_BLOCK_POS = 0x44;
    const NUM_EXTENSION_BLOCK_POS = 0x48;
    const BIG_BLOCK_DEPOT_BLOCKS_POS = 0x4c;

    // property storage offsets (directory offsets)
    const SIZE_OF_NAME_POS = 0x40;
    const TYPE_POS = 0x42;
    const START_BLOCK_POS = 0x74;
    const SIZE_POS = 0x78;
    public $workbook = null;
    public $summaryInformation = null;
    public $documentSummaryInformation = null;

    protected $data = '';
    protected $bigBlockChain = '';
    protected $smallBlockChain = '';
    protected $entry = '';
    protected $props = [];
    protected $rootEntry = 0;
    protected $sbdStartBlock = 0;
    protected $extensionBlock = 0;
    protected $rootStartBlock = 0;
    protected $numExtensionBlocks = 0;
    protected $numBigBlockDepotBlocks = 0;

    /**
     * Read the file
     *
     * @throws ParserException|ReaderException
     * @param string $file
     */
    public function read($file) {
        $this->openFile($file);

        // Total number of sectors used for the SAT
        $this->numBigBlockDepotBlocks = Format::getInt4d($this->data, self::NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);

        // SecID of the first sector of the directory stream
        $this->rootStartBlock = Format::getInt4d($this->data, self::ROOT_START_BLOCK_POS);

        // SecID of the first sector of the SSAT (or -2 if not extant)
        $this->sbdStartBlock = Format::getInt4d($this->data, self::SMALL_BLOCK_DEPOT_BLOCK_POS);

        // SecID of the first sector of the MSAT (or -2 if no additional sectors are used)
        $this->extensionBlock = Format::getInt4d($this->data, self::EXTENSION_BLOCK_POS);

        // Total number of sectors used by MSAT
        $this->numExtensionBlocks = Format::getInt4d($this->data, self::NUM_EXTENSION_BLOCK_POS);

        $bigBlockDepotBlocks = [];
        $pos = self::BIG_BLOCK_DEPOT_BLOCKS_POS;
        $bbdBlocks = $this->numBigBlockDepotBlocks;
        if ($this->numExtensionBlocks != 0) {
            $bbdBlocks = (self::BIG_BLOCK_SIZE - self::BIG_BLOCK_DEPOT_BLOCKS_POS) / 4;
        }

        for ($i = 0; $i < $bbdBlocks; ++$i) {
            $bigBlockDepotBlocks[$i] = Format::getInt4d($this->data, $pos);
            $pos += 4;
        }

        for ($j = 0; $j < $this->numExtensionBlocks; ++$j) {
            $pos = ($this->extensionBlock + 1) * self::BIG_BLOCK_SIZE;
            $blocksToRead = min($this->numBigBlockDepotBlocks - $bbdBlocks, self::BIG_BLOCK_SIZE / 4 - 1);

            for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; ++$i) {
                $bigBlockDepotBlocks[$i] = Format::getInt4d($this->data, $pos);
                $pos += 4;
            }

            $bbdBlocks += $blocksToRead;
            if ($bbdBlocks < $this->numBigBlockDepotBlocks) {
                $this->extensionBlock = Format::getInt4d($this->data, $pos);
            }
        }

        $this->bigBlockChain = '';
        $bbs = self::BIG_BLOCK_SIZE / 4;
        for ($i = 0; $i < $this->numBigBlockDepotBlocks; ++$i) {
            $pos = ($bigBlockDepotBlocks[$i] + 1) * self::BIG_BLOCK_SIZE;
            $this->bigBlockChain .= substr($this->data, $pos, 4 * $bbs);
        }

        $sbdBlock = $this->sbdStartBlock;
        $this->smallBlockChain = '';
        while ($sbdBlock != -2) {
            $pos = ($sbdBlock + 1) * self::BIG_BLOCK_SIZE;
            $this->smallBlockChain .= substr($this->data, $pos, 4 * $bbs);

            $sbdBlock = Format::getInt4d($this->bigBlockChain, $sbdBlock * 4);
        }

        // read the directory stream
        $block = $this->rootStartBlock;
        $this->entry = $this->readData($block);
        $this->readPropertySets();
    }

    /**
     * Open file for reading
     *
     * @param string $file
     *
     * @throws ReaderException|ParserException
     */
    public function openFile($file) {
        // Check if file exists
        if (!file_exists($file) || !is_readable($file)) {
            throw new ReaderException("Could not open file [$file] for reading! File does not exist.");
        }

        // Get the file data
        $this->data = file_get_contents($file);

        // Check OLE identifier
        if (empty($this->data) || substr($this->data, 0, 8) != self::IDENTIFIER_OLE) {
            throw new ParserException("The file [$file] is not recognised as an OLE file");
        }
    }

    /**
     * Extract binary stream data.
     *
     * @param int $stream
     *
     * @return string|null
     */
    public function getStream($stream) {
        if ($stream === null) {
            return null;
        }

        $streamData = '';
        if ($this->props[$stream]['size'] < self::SMALL_BLOCK_THRESHOLD) {
            $rootData = $this->readData($this->props[$this->rootEntry]['startBlock']);
            $block = $this->props[$stream]['startBlock'];

            while ($block != -2) {
                $pos = $block * self::SMALL_BLOCK_SIZE;
                $streamData .= substr($rootData, $pos, self::SMALL_BLOCK_SIZE);
                $block = Format::getInt4d($this->smallBlockChain, $block * 4);
            }

            return $streamData;
        }

        $numBlocks = $this->props[$stream]['size'] / self::BIG_BLOCK_SIZE;
        if ($this->props[$stream]['size'] % self::BIG_BLOCK_SIZE != 0) {
            ++$numBlocks;
        }

        if ($numBlocks == 0) {
            return '';
        }

        $block = $this->props[$stream]['startBlock'];
        while ($block != -2) {
            $pos = ($block + 1) * self::BIG_BLOCK_SIZE;
            $streamData .= substr($this->data, $pos, self::BIG_BLOCK_SIZE);
            $block = Format::getInt4d($this->bigBlockChain, $block * 4);
        }

        return $streamData;
    }

    /**
     * Read a standard stream (by joining sectors using information from SAT).
     *
     * @param int $bl Sector ID where the stream starts
     *
     * @return string
     */
    protected function readData($bl) {
        $block = $bl;
        $data = '';

        while ($block != -2) {
            $pos = ($block + 1) * self::BIG_BLOCK_SIZE;
            $data .= substr($this->data, $pos, self::BIG_BLOCK_SIZE);
            $block = Format::getInt4d($this->bigBlockChain, $block * 4);
        }

        return $data;
    }

    /**
     * Read entries in the directory stream.
     */
    protected function readPropertySets() {
        $offset = 0;

        // loop through entires, each entry is 128 bytes
        $entryLen = strlen($this->entry);
        while ($offset < $entryLen) {
            // entry data (128 bytes)
            $d = substr($this->entry, $offset, self::PROPERTY_STORAGE_BLOCK_SIZE);

            // size in bytes of name
            $nameSize = ord($d[self::SIZE_OF_NAME_POS]) | (ord($d[self::SIZE_OF_NAME_POS + 1]) << 8);

            // type of entry
            $type = ord($d[self::TYPE_POS]);

            // sectorID of first sector or short sector, if this entry refers to a stream (the case with workbook)
            // sectorID of first sector of the short-stream container stream, if this entry is root entry
            $startBlock = Format::getInt4d($d, self::START_BLOCK_POS);
            $size = Format::getInt4d($d, self::SIZE_POS);
            $name = str_replace("\x00", '', substr($d, 0, $nameSize));
            $this->props[] = [
                'name' => $name,
                'type' => $type,
                'startBlock' => $startBlock,
                'size' => $size,
            ];

            // tmp helper to simplify checks
            $upName = strtoupper($name);

            // Workbook directory entry (BIFF5 uses Book, BIFF8 uses Workbook)
            if (($upName === 'WORKBOOK') || ($upName === 'BOOK')) {
                $this->workbook = count($this->props) - 1;
            } elseif ($upName === 'ROOT ENTRY' || $upName === 'R') {
                // Root entry
                $this->rootEntry = count($this->props) - 1;
            }

            // Summary information
            if ($name == chr(5) . 'SummaryInformation') {
                $this->summaryInformation = count($this->props) - 1;
            }

            // Additional Document Summary information
            if ($name == chr(5) . 'DocumentSummaryInformation') {
                $this->documentSummaryInformation = count($this->props) - 1;
            }

            $offset += self::PROPERTY_STORAGE_BLOCK_SIZE;
        }
    }
}
