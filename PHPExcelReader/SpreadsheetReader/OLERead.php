<?php

defined('IDENTIFIER_OLE') ||
	define('IDENTIFIER_OLE', pack('CCCCCCCC', 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1));

class PHPExcel_Shared_OLERead {
	private $data = '';
	
	const IDENTIFIER_OLE 					= IDENTIFIER_OLE;	// OLE identifier
	const BIG_BLOCK_SIZE					= 0x200;
	const SMALL_BLOCK_SIZE					= 0x40;			// Size of a short sector = 64 bytes
	const PROPERTY_STORAGE_BLOCK_SIZE		= 0x80;			// Size of a directory entry always = 128 bytes
	const SMALL_BLOCK_THRESHOLD				= 0x1000;		// Minimum size of a standard stream = 4096 bytes, streams smaller than this are stored as short streams
	
	// header offsets
	const NUM_BIG_BLOCK_DEPOT_BLOCKS_POS	= 0x2c;
	const ROOT_START_BLOCK_POS				= 0x30;
	const SMALL_BLOCK_DEPOT_BLOCK_POS		= 0x3c;
	const EXTENSION_BLOCK_POS				= 0x44;
	const NUM_EXTENSION_BLOCK_POS			= 0x48;
	const BIG_BLOCK_DEPOT_BLOCKS_POS		= 0x4c;
	
	// property storage offsets (directory offsets)
	const SIZE_OF_NAME_POS					= 0x40;
	const TYPE_POS							= 0x42;
	const START_BLOCK_POS					= 0x74;
	const SIZE_POS							= 0x78;
	
	public $error = false;
	public $workbook = null;
	public $summaryInformation = null;
	public $documentSummaryInformation = null;
	
	public function read($file){
		if( ! is_readable($file)) {
			throw new Exception('SpreadsheetReader_XLS: File not readable (' . $file . ')');
		}
		
		$this->data = file_get_contents($file);
		if( ! $this->data || substr($this->data, 0, 8) != self::IDENTIFIER_OLE){
			$this->error = true;
			return false;
		}
		
		$this->numBigBlockDepotBlocks = self::_GetInt4d($this->data, self::NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);
		$this->rootStartBlock = self::_GetInt4d($this->data, self::ROOT_START_BLOCK_POS);
		$this->sbdStartBlock = self::_GetInt4d($this->data, self::SMALL_BLOCK_DEPOT_BLOCK_POS);
		$this->extensionBlock = self::_GetInt4d($this->data, self::EXTENSION_BLOCK_POS);
		$this->numExtensionBlocks = self::_GetInt4d($this->data, self::NUM_EXTENSION_BLOCK_POS);
		
		$bigBlockDepotBlocks = array();
		$pos = self::BIG_BLOCK_DEPOT_BLOCKS_POS;
		$bbdBlocks = $this->numExtensionBlocks == 0 ? $this->numBigBlockDepotBlocks : (self::BIG_BLOCK_SIZE - self::BIG_BLOCK_DEPOT_BLOCKS_POS) / 4;
		for ($i = 0; $i < $bbdBlocks; ++$i) {
			$bigBlockDepotBlocks[$i] = self::_GetInt4d($this->data, $pos);
			$pos += 4;
		}
		
		for ($j = 0; $j < $this->numExtensionBlocks; ++$j) {
			$pos = ($this->extensionBlock + 1) * self::BIG_BLOCK_SIZE;
			$blocksToRead = min($this->numBigBlockDepotBlocks - $bbdBlocks, self::BIG_BLOCK_SIZE / 4 - 1);
		
			for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; ++$i) {
				$bigBlockDepotBlocks[$i] = self::_GetInt4d($this->data, $pos);
				$pos += 4;
			}
		
			$bbdBlocks += $blocksToRead;
			if ($bbdBlocks < $this->numBigBlockDepotBlocks) {
				$this->extensionBlock = self::_GetInt4d($this->data, $pos);
			}
		}
		
		$pos = 0;
		$this->bigBlockChain = '';
		$bbs = self::BIG_BLOCK_SIZE / 4;
		for ($i = 0; $i < $this->numBigBlockDepotBlocks; ++$i) {
			$pos = ($bigBlockDepotBlocks[$i] + 1) * self::BIG_BLOCK_SIZE;
		
			$this->bigBlockChain .= substr($this->data, $pos, 4*$bbs);
			$pos += 4*$bbs;
		}
		
		$pos = 0;
		$sbdBlock = $this->sbdStartBlock;
		$this->smallBlockChain = '';
		while ($sbdBlock != -2) {
			$pos = ($sbdBlock + 1) * self::BIG_BLOCK_SIZE;
			$this->smallBlockChain .= substr($this->data, $pos, 4*$bbs);
			$pos += 4*$bbs;
			$sbdBlock = self::_GetInt4d($this->bigBlockChain, 4*$sbdBlock);
		}
		
		$block = $this->rootStartBlock;				// read the directory stream
		$this->entry = $this->_readData($block);
		
		$this->_readPropertySets();
	}
	
	/**
	 * Extract binary stream data
	 *
	 * @return string
	 */
	public function getStream($stream) {
		if ($stream === NULL) {
			return null;
		}
	
		$streamData = '';
		if ($this->props[$stream]['size'] < self::SMALL_BLOCK_THRESHOLD) {
			$rootdata = $this->_readData($this->props[$this->rootentry]['startBlock']);
			$block = $this->props[$stream]['startBlock'];
	
			while ($block != -2) {
				$pos = $block * self::SMALL_BLOCK_SIZE;
				$streamData .= substr($rootdata, $pos, self::SMALL_BLOCK_SIZE);
				$block = self::_GetInt4d($this->smallBlockChain, $block*4);
			}
		} 
		else {
			$numBlocks = $this->props[$stream]['size'] / self::BIG_BLOCK_SIZE;
			if ($this->props[$stream]['size'] % self::BIG_BLOCK_SIZE != 0) {
				++$numBlocks;
			}
	
			if($numBlocks){
				$block = $this->props[$stream]['startBlock'];
		
				while ($block != -2) {
					$pos = ($block + 1) * self::BIG_BLOCK_SIZE;
					$streamData .= substr($this->data, $pos, self::BIG_BLOCK_SIZE);
					$block = self::_GetInt4d($this->bigBlockChain, $block*4);
				}
			}
		}
		
		return $streamData;
	}
	
	/**
	 * Read a standard stream (by joining sectors using information from SAT)
	 *
	 * @param int $bl Sector ID where the stream starts
	 * @return string Data for standard stream
	 */
	private function _readData($block) {
		$data = '';
	
		while ($block != -2) {
			$pos = ($block + 1) * self::BIG_BLOCK_SIZE;
			$data .= substr($this->data, $pos, self::BIG_BLOCK_SIZE);
			$block = self::_GetInt4d($this->bigBlockChain, 4*$block);
		}
		return $data;
	}
	
	/**
	 * Read entries in the directory stream.
	 */
	private function _readPropertySets() {
		$offset = 0;
	
		$entryLen = strlen($this->entry);		// loop through entires, each entry is 128 bytes
		while ($offset < $entryLen) {
			$data = substr($this->entry, $offset, self::PROPERTY_STORAGE_BLOCK_SIZE);							// entry data (128 bytes)
			$nameSize = ord($data[self::SIZE_OF_NAME_POS]) | (ord($data[self::SIZE_OF_NAME_POS + 1]) << 8);		// size in bytes of name
			$name = str_replace("\x00", "", substr($data, 0, $nameSize));
			$this->props[] = array (
				'name' 		 =>	 $name,
				'type' 		 =>	 ord($data[self::TYPE_POS]),			// type of entry
				'size' 		 =>	 self::_GetInt4d($data, self::SIZE_POS),
				'startBlock' =>	 self::_GetInt4d($data, self::START_BLOCK_POS)
			);
	
			$upName = strtoupper($name);								// tmp helper to simplify checks
			if (($upName === 'WORKBOOK') || ($upName === 'BOOK')) {		// Workbook directory entry (BIFF5 uses Book, BIFF8 uses Workbook)
				$this->workbook = count($this->props) - 1;
			}
			else if ( $upName === 'ROOT ENTRY' || $upName === 'R') {
				$this->rootentry = count($this->props) - 1;				// Root entry
			}
			
			if ($name == chr(5) . 'SummaryInformation') {
				$this->summaryInformation = count($this->props) - 1;			// Summary information
			}
	
			if ($name == chr(5) . 'DocumentSummaryInformation') {
				$this->documentSummaryInformation = count($this->props) - 1;	// Additional Document Summary information
			}
	
			$offset += self::PROPERTY_STORAGE_BLOCK_SIZE;
		}
	
	}
	
	/**
	 * Read 4 bytes of data at specified position
	 * FIX: represent numbers correctly on 64-bit system. Hacked by Andreas Rehm 2006 to ensure correct result of the <<24 block on 32 and 64bit systems
	 * http://sourceforge.net/tracker/index.php?func=detail&aid=1487372&group_id=99160&atid=623334
	 * 
	 * @param string $data
	 * @param int $pos
	 * @return int
	 */
	private static function _GetInt4d($data, $pos){
		$_or_24 = ord($data[$pos + 3]);
		if ($_or_24 >= 128) {
			$_ord_24 = -abs((256 - $_or_24) << 24);		// negative number
		} else {
			$_ord_24 = ($_or_24 & 127) << 24;
		}
		return ord($data[$pos]) | (ord($data[$pos + 1]) << 8) | (ord($data[$pos + 2]) << 16) | $_ord_24;
	}
}
