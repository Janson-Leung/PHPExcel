<?php
/**
 * autoload library
 */

PHPExcel_Autoloader::Register();

class PHPExcel_Autoloader {
	/**
	 * Register the Autoloader with SPL
	 */
	public static function Register() {
		if (function_exists('__autoload')) {
			spl_autoload_register('__autoload');
		}

		return spl_autoload_register(array('PHPExcel_Autoloader', 'Load'));
	}
	
	
	/**
	 * Autoload a class identified by name
	 * @param	string	$pClassName	Name of the object to load
	 */
	public static function Load($pClassName){
		if ((class_exists($pClassName, FALSE)) || (strpos($pClassName, 'PHPExcel') !== 0)) {
			return FALSE;
		}
	
		$pClassFilePath = PHPEXCEL_ROOT . str_replace('_', DIRECTORY_SEPARATOR, $pClassName) . '.php';

		if ((file_exists($pClassFilePath) === FALSE) || (is_readable($pClassFilePath) === FALSE)) {
			return FALSE;
		}
	
		require $pClassFilePath;
	}
}