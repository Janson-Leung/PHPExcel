<?php
/**
 * PHPExcelReader class
 *
 * @version 1.0.0
 * @author Janson Leung
 */
class PHPExcelReader implements SeekableIterator, Countable {
	const TYPE_XLSX = 'XLSX';
	const TYPE_XLS = 'XLS';
	const TYPE_CSV = 'CSV';
	const TYPE_ODS = 'ODS';

	private $options = array(
		'Delimiter' => '',
		'Enclosure' => '"'
	);

	private $index = 0;
	private $handle = array();
	private $type = false;

	/**
	 * @param string Path to file
	 * @param string Original filename (in case of an uploaded file), used to determine file type, optional
	 * @param string MIME type from an upload, used to determine file type, optional
	 */
	public function __construct($filePath, $originalFileName = false, $mimeType = false) {
		if ( ! is_readable($filePath)) {
			throw new Exception('SpreadsheetReader: File (' . $filePath . ') not readable');
		}

		$defaultTimeZone = @date_default_timezone_get();
		if ($defaultTimeZone)	{
			date_default_timezone_set($defaultTimeZone);
		}

		// Checking the other parameters for correctness
		// This should be a check for string but we're lenient
		if ( ! empty($originalFileName) && ! is_scalar($originalFileName)) {
			throw new Exception('SpreadsheetReader: Original file (2nd parameter) path is not a string or a scalar value.');
		}
		if ( ! empty($mimeType) && ! is_scalar($mimeType)) {
			throw new Exception('SpreadsheetReader: Mime type (3nd parameter) path is not a string or a scalar value.');
		}

		// 1. Determine type
		if ( ! $originalFileName) {
			$originalFileName = $filePath;
		}

		$Extension = strtolower(pathinfo($originalFileName, PATHINFO_EXTENSION));
		if($mimeType) {
			switch ($mimeType) {
				case 'text/csv':
				case 'text/comma-separated-values':
				case 'text/plain':
					$this->type = self::TYPE_CSV;
					break;
				case 'application/vnd.ms-excel':
				case 'application/msexcel':
				case 'application/x-msexcel':
				case 'application/x-ms-excel':
				case 'application/vnd.ms-excel':
				case 'application/x-excel':
				case 'application/x-dos_ms_excel':
				case 'application/xls':
				case 'application/xlt':
				case 'application/x-xls':
					// Excel does weird stuff
					if (in_array($Extension, array('csv', 'tsv', 'txt'))) {
						$this->type = self::TYPE_CSV;
					}
					else {
						$this->type = self::TYPE_XLS;
					}
					break;
				case 'application/vnd.oasis.opendocument.spreadsheet':
				case 'application/vnd.oasis.opendocument.spreadsheet-template':
					$this->type = self::TYPE_ODS;
					break;
				case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
				case 'application/vnd.openxmlformats-officedocument.spreadsheetml.template':
				case 'application/xlsx':
				case 'application/xltx':
					$this->type = self::TYPE_XLSX;
					break;
				case 'application/xml':
					// Excel 2004 xml format uses this
					break;
			}
		}

		if ( ! $this->type)	{
			switch ($Extension) {
				case 'xlsx':
				case 'xltx': // XLSX template
				case 'xlsm': // Macro-enabled XLSX
				case 'xltm': // Macro-enabled XLSX template
					$this->type = self::TYPE_XLSX;
					break;
				case 'xls':
				case 'xlt':
					$this->type = self::TYPE_XLS;
					break;
				case 'ods':
				case 'odt':
					$this->type = self::TYPE_ODS;
					break;
				default:
					$this->type = self::TYPE_CSV;
					break;
			}
		}

		// Pre-checking XLS files, in case they are renamed CSV or XLSX files
		if ($this->type == self::TYPE_XLS) {
			self::Load(self::TYPE_XLS);
			$this->handle = new SpreadsheetReader_XLS($filePath);
			if ($this->handle->error) {
				$this->handle->__destruct();

				if (is_resource($Ziphandle = zip_open($filePath))) {
					$this->type = self::TYPE_XLSX;
					zip_close($Ziphandle);
				}
				else {
					$this->type = self::TYPE_CSV;
				}
			}
		}

		// 2. Create handle
		switch ($this->type) {
			case self::TYPE_XLSX:
				self::Load(self::TYPE_XLSX);
				$this->handle = new SpreadsheetReader_XLSX($filePath);
				break;
			case self::TYPE_CSV:
				self::Load(self::TYPE_CSV);
				$this->handle = new SpreadsheetReader_CSV($filePath, $this->options, 'GBK');
				break;
			case self::TYPE_XLS:
				// Everything already happens above
				break;
			case self::TYPE_ODS:
				self::Load(self::TYPE_ODS);
				$this->handle = new SpreadsheetReader_ODS($filePath, $this->options);
				break;
		}
	}
	
	/**
	 * get the type of file
	 * @return string
	 */
	public function getType() {
		return $this->type;
	}
	
	/**
	 * Gets information about separate sheets in the given file
	 *
	 * @return array Associative array where key is sheet index and value is sheet name
	 */
	public function Sheets() {
		return $this->handle->Sheets();
	}

	/**
	 * Changes the current sheet to another from the file.
	 *	Note that changing the sheet will rewind the file to the beginning, even if
	 *	the current sheet index is provided.
	 *
	 * @param int Sheet index
	 *
	 * @return bool True if sheet could be changed to the specified one,
	 *	false if not (for example, if incorrect index was provided.
	 */
	public function ChangeSheet($index) {
		return $this->handle->ChangeSheet($index);
	}

	/**
	 * Autoloads the required class for the particular spreadsheet type
	 *
	 * @param TYPE_* Spreadsheet type, one of TYPE_* constants of this class
	 */
	private static function Load($type) {
		if ( ! in_array($type, array(self::TYPE_XLSX, self::TYPE_XLS, self::TYPE_CSV, self::TYPE_ODS))) {
			throw new Exception('SpreadsheetReader: Invalid type (' . $type . ')');
		}

		// 2nd parameter is to prevent autoloading for the class.
		// If autoload works, the require line is unnecessary, if it doesn't, it ends badly.
		if ( ! class_exists('SpreadsheetReader_' . $type, false)) {
			require(dirname(__FILE__) . DIRECTORY_SEPARATOR . 'SpreadsheetReader' . DIRECTORY_SEPARATOR . 'SpreadsheetReader_' . $type . '.php');
		}
	}

	// !Iterator interface methods

	/** 
	 * Rewind the Iterator to the first element.
	 * Similar to the reset() function for arrays in PHP
	 */ 
	public function rewind() {
		$this->index = 0;
		if ($this->handle) {
			$this->handle->rewind();
		}
	}

	/** 
	 * Return the current element.
	 * Similar to the current() function for arrays in PHP
	 *
	 * @return mixed current element from the collection
	 */
	public function current() {
		if ($this->handle) {
			return $this->handle->current();
		}
		return null;
	}

	/** 
	 * Move forward to next element. 
	 * Similar to the next() function for arrays in PHP 
	 */ 
	public function next() {
		if ($this->handle) {
			$this->index++;
			return $this->handle->next();
		}
		return null;
	}

	/** 
	 * Return the identifying key of the current element.
	 * Similar to the key() function for arrays in PHP
	 *
	 * @return mixed either an integer or a string
	 */ 
	public function key() {
		if ($this->handle) {
			return $this->handle->key();
		}
		return null;
	}

	/** 
	 * Check if there is a current element after calls to rewind() or next().
	 * Used to check if we've iterated to the end of the collection
	 *
	 * @return boolean FALSE if there's nothing more to iterate over
	 */ 
	public function valid()	{
		if ($this->handle) {
			return $this->handle->valid();
		}
		return false;
	}

	/**
	 * total of file number
	 * return int
	 */
	public function count()	{
		if ($this->handle) {
			return $this->handle->count();
		}
		return 0;
	}

	/**
	 * Method for SeekableIterator interface. Takes a posiiton and traverses the file to that position
	 * The value can be retrieved with a `current()` call afterwards.
	 *
	 * @param int position in file
	 */
	public function seek($position)	{
		if ( ! $this->handle) {
			throw new OutOfBoundsException('SpreadsheetReader: No file opened');
		}

		$Currentindex = $this->handle->key();
		if ($Currentindex != $position) {
			if ($position < $Currentindex || is_null($Currentindex) || $position == 0) {
				$this->rewind();
			}

			while ($this->handle->valid() && ($position > $this->handle->key())) {
				$this->handle->next();
			}

			if ( ! $this->handle->valid()) {
				throw new OutOfBoundsException('SpreadsheetError: position ' . $position . ' not found');
			}
		}

		return null;
	}
}
