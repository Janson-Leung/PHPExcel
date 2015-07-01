<?php

class SpreadsheetReader_CSV implements Iterator, Countable {
	private $options = array(
		'Delimiter' => ';',
		'Enclosure' => '"'
	);
	
	private $encoding = 'UTF-8';
	private $filePath = '';
	private $handle = false;
	private $index = 0;
	private $currentRow = null;
	private $rowCount = null;
	
	public function __construct($filePath, $options = null, $encoding = '') {
		if ( ! is_readable($filePath)) {
			throw new Exception('SpreadsheetReader_CSV: File not readable (' . $filePath . ')');
		}
		
		$this->filePath = $filePath;
		@ini_set('auto_detect_line_endings', true);
		$this->options = array_merge($this->options, $options);
		$encoding && $this->encoding = $encoding;
		$this->handle = fopen($filePath, 'r');
		
		// Checking the file for byte-order mark to determine encoding
		$BOM16 = bin2hex(fread($this->handle, 2));
		if ($BOM16 == 'fffe') {
			$this->Encoding = 'UTF-16LE';
			$this->BOMLength = 2;
		}
		elseif ($BOM16 == 'feff') {
			$this->Encoding = 'UTF-16BE';
			$this->BOMLength = 2;
		}
		
		if ( ! $this->BOMLength) {
			fseek($this->handle, 0);
			$BOM32 = bin2hex(fread($this->handle, 4));
			if ($BOM32 == '0000feff') {
				$this->Encoding = 'UTF-32';
				$this->BOMLength = 4;
			}
			elseif ($BOM32 == 'fffe0000') {
				$this->Encoding = 'UTF-32';
				$this->BOMLength = 4;
			}
		}
		
		fseek($this->handle, 0);
		$BOM8 = bin2hex(fread($this->handle, 3));
		if ($BOM8 == 'efbbbf') {
			$this->Encoding = 'UTF-8';
			$this->BOMLength = 3;
		}
		
		// Seeking the place right after BOM as the start of the real content
		if ($this->BOMLength) {
			fseek($this->handle, $this->BOMLength);
		}
		
		// Checking for the delimiter if it should be determined automatically
		if ( ! $this->options['Delimiter']) {
			$Semicolon = ';';		// fgetcsv needs single-byte separators
			$Tab = "\t";
			$Comma = ',';
		
			// Reading the first row and checking if a specific separator character
			// has more columns than others (it means that most likely that is the delimiter).
			$SemicolonCount = count(fgetcsv($this->handle, null, $Semicolon));
			fseek($this->handle, $this->BOMLength);
			$TabCount = count(fgetcsv($this->handle, null, $Tab));
			fseek($this->handle, $this->BOMLength);
			$CommaCount = count(fgetcsv($this->handle, null, $Comma));
			fseek($this->handle, $this->BOMLength);
		
			$Delimiter = $Semicolon;
			if ($TabCount > $SemicolonCount || $CommaCount > $SemicolonCount) {
				$Delimiter = $CommaCount > $TabCount ? $Comma : $Tab;
			}
		
			$this->options['Delimiter'] = $Delimiter;
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
	 * @param bool
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
		fseek($this->handle, 0);
		$this->currentRow = null;
		$this->index = 0;
	}
	
	/**
	 * Return the current element.
	 * @return mixed
	 */
	public function current() {
		if ($this->index == 0 && is_null($this->currentRow)) {
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

		if ($this->Encoding == 'UTF-16LE' || $this->Encoding == 'UTF-16BE')	{
			while ( ! feof($this->handle)) {
				$Char = ord(fgetc($this->handle));
				if ( ! $Char || $Char == 10 || $Char == 13)	{
					continue;												// While bytes are insignificant whitespace, do nothing
				}
				else {
					if ($this->Encoding == 'UTF-16LE') {
						fseek($this->handle, ftell($this->handle) - 1);		// When significant bytes are found, step back to the last place before them
					}
					else {
						fseek($this->handle, ftell($this->handle) - 2);
					}
					break;
				}
			}
		}
		
		$this->index++;
		$this->currentRow = fgetcsv($this->handle, null, $this->options['Delimiter'], $this->options['Enclosure']);
		if ($this->currentRow) {
			if ($this->encoding != 'ASCII' && $this->encoding != 'UTF-8') {
				foreach($this->currentRow as $key => $value) {
					$this->currentRow[$key] =  trim(trim(
						mb_convert_encoding($value, 'UTF-8', $this->encoding),
						$this->options['Enclosure']
					));
				}
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
		return ($this->currentRow || ! feof($this->handle));
	}
	
	/**
	 * return the count of the contained items
	 * @return int
	 */
	public function count() {
		if (is_null($this->rowCount)) {
			$total = 0;
			
			fseek($this->handle, 0);
			while ($row = fgetcsv($this->handle, null, $this->options['Delimiter'], $this->options['Enclosure'])) {
				$total++;
			}
			
			$this->rowCount = $total;
		}
		
		return $this->rowCount;
	}
}
