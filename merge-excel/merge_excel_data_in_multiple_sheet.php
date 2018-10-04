<?php
$filePath =  $argv[1];
$saving_name = "E:\\merge-data.xls";//change the name and location of file.
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/Writer/Excel2007.php';
if ($dh = opendir($filePath)){
	$dataArray = array();
	$counter = 1;
	$file_name = array();
	$new_xl = new PHPExcel();
	while (($file = readdir($dh)) !== false) {
		$path_parts = pathinfo($file);		
		if ($path_parts['extension'] === 'xlsx') {			
			$singlefile = $path_parts["filename"]; //file name without extention
			echo "Start merging the content from file " . $singlefile .".xlsx" . PHP_EOL;
			$workbook_file = $filePath.'\\'. $file;
			$reader = PHPExcel_IOFactory::createReader('Excel2007');
			$xl = $reader->load($workbook_file);
			$firstSheet = $xl->getSheet(0);
			$workbook_name = $singlefile; 
			$firstSheet->setTitle($workbook_name);
			$new_xl->addExternalSheet($firstSheet);
			echo "End merging the content from file " . $singlefile . PHP_EOL;
		}
	}
	
	$new_xl->removeSheetByIndex(0);
	$writer = PHPExcel_IOFactory::createWriter($new_xl, 'Excel2007');
	// name of file, which needs to be attached during email sending
	$writer->save($saving_name);
        
	echo "**********************************************************************". PHP_EOL;
	echo "All the contents of excel file are successfully merged in sheet." . PHP_EOL;
	echo "**********************************************************************". PHP_EOL;
}   
?>