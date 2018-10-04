# Project Title

Combine multiple excel files in multiple excel sheet of a excel using PHP

## Getting Started

This code is help you to combine `n` number of excel files data in a multiple sheet of a excel with in a second. You just need to download the source code and follow the below instructions.


### Prerequisites

This code is using php excel library and offcouse PHP. So you have to include PHPEXCEL library in you file.

```
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/Writer/Excel2007.php';
require_once 'Classes/PHPExcel/IOFactory.php';

```

### Installing

In this sample code you can see a folder name is "files". There are a number of excel files in it and a file "merge_excel_data_in_multiple_sheet.php". 

This is the main file which will be use to merge all these files data into the single excel sheet.

```
$filePath =  $argv[1];
$saving_name = "F:\commit-excel-merge-multiple\merge-data.xls";//change the name and location of file.
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/Writer/Excel2007.php';
require_once 'Classes/PHPExcel/IOFactory.php';

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

```

## Running the tests

After downloading all the files on your computer you need to run the above file using command prompt on windows system like below :

```

F:\commit-excel-merge-multiple\merge-excel> php merge_excel_data_in_multiple_sheet.php F:\commit-excel-merge-multiple\merge-excel\files

```

### Explain :

1- Go to the directory where put your downloaded folder. In the above example i put it into my "F:\" directory. 

2- So first i go into this folder and then i run the "merge_excel_data_in_single_sheet.php" file by using above command.

3- Now after finish the process you will see that all the excel file from "files" folder are combine into multiple sheet in "merge-data.xlsx" file. "E:\\mergeFileData.xlsx" file.


## Built With

* [PHPEXCEL](https://github.com/PHPOffice/PHPExcel) - The PHP EXCEL library is used.
* [PHP WAMP SERVER](http://www.wampserver.com/en/) - PHP WAMP server is used.

## Contributing

We welcome the new commit of changes in this code. If any body want to contribute in it. (http://phpsollutions.blogspot.com) Please submit a pull requests to us.

## Authors

* **Salman Ahmad** - *Initial work* - [PHPSOLLUTIONS.BLOGSPOT.COM](https://phpsollutions.blogspot.com/p/blog-page.html)

## License

This project is developed using the free open source. So any body are free to download and use this code. 
