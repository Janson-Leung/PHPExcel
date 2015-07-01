<?php
require 'PHPExcelReader/PHPExcelReader.php';

try{
	$Reader = new PHPExcelReader('test.xls');
	$total = $Reader->count();		// get the total rows of records
	//$current = $Reader->current();	// get the current row data
	
	/*
	$Reader->seek(4);			// skip to the 4th row 
	$row = $Reader->current();		// get the 4th row data
	*/
	
	/*
	foreach($Reader as $key => $row){
		$data[] = $row;			// loop obtain row data
	}
	*/
	
	var_dump($total);
} catch (Exception $e) {
	die($e->getMessage());
}
