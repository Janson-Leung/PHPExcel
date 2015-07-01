**PHPExcelReader** is a lightweight Excel file reading class, support the CSV, XLS, XLSX file. It can read 
line by line as needed.

### Requirements:
*  PHP 5.3.0 or newer
*  PHP must have Zip file support (see http://php.net/manual/en/zip.installation.php)

### Usage:

All data is read from the file sequentially, with each row being returned as a numeric array.
This is about the easiest way to read a file:

	<?php
		require 'PHPExcelReader/PHPExcelReader.php';

		try{
			$Reader = new PHPExcelReader('test.xls');
			$total = $Reader->count();			// get the total rows of records
			//$current = $Reader->current();	// get the current row data
		
			/*
			$Reader->seek(4);					// skip to the 4th row 
			$row = $Reader->current();			// get the 4th row data
			*/
			
			/*
			foreach($Reader as $key => $row){
				$data[] = $row;					// loop obtain row data
			}
			*/
			
			var_dump($total);
		} catch (Exception $e) {
			die($e->getMessage());
		}
