<?php

class PHPExcel_Reader_XLS implements Iterator, Countable {
	private $handle = false;
	private $index = 0;
	private $rowCount = null;
	private $currentSheet = 0;
	private $currentRow = null;

	public  $error = false;

	public function __construct($filePath) {
		if ( ! file_exists($filePath)) {
			throw new Exception("Could not open " . $filePath . " for reading! File does not exist.");
		}

		try {
			$this->handle = new PHPExcel_Reader_Excel5($filePath);

			return true;
		} catch (Exception $e) {
			$this->error = true;
			return false;
		}
	}

	public function __destruct() {
		unset($this->handle);
	}

	/**
	 * Retrieves an array with information about sheets in the current file
	 *
	 * @return array List of sheets (key is sheet index, value is name)
	 */
	public function Sheets() {
		$this->sheetInfo = $this->handle->getWorksheetInfo();
		$this->rowCount = $this->sheetInfo['totalRows'];

		return $this->sheetInfo;
	}

	/**
	 * Changes the current sheet in the file to another
	 * @param $index int
	 * @return bool
	 */
	public function ChangeSheet($index)	{
		return $this->handle->ChangeSheet($index);
	}

	/**
	 * Rewind the Iterator to the first element.
	 */
	public function rewind() {
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
			$this->Index--;
		}

		return $this->currentRow;
	}

	/**
	 * Move forward to next element.
	 */
	public function next() {
		$this->currentRow = array();
		if( ! $this->sheetInfo) {
			$this->Sheets();
		}

		$this->index++;
		$cell = $this->handle->getCell();
		for($i = 0; $i < $this->sheetInfo['totalColumns']; $i++) {
			$this->currentRow[$i] = isset($cell[$i]) ? $cell[$i] : '';
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
	 * @return boolean
	 */
	public function valid() {
		if ($this->error) {
			return false;
		}

		return ($this->index <= $this->count());
	}

	/**
	 * return the count of the contained items
	 */
	public function count() {
		if ($this->error) {
			return 0;
		}

		if( ! isset($this->rowCount)){
			$this->Sheets();
		}

		return $this->rowCount;
	}
}
