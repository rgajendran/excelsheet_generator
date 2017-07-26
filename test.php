<?php

require_once 'lib/PHPExcel/Classes/PHPExcel.php';
require_once 'lib/PHPExcel/Classes/PHPExcel/IOFactory.php';


$inputFileName = 'sheet.xlsx';

//  Read your Excel workbook
try {
    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($inputFileName);
} catch(Exception $e) {
    die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}

//  Get worksheet dimensions
$sheet = $objPHPExcel->getSheet(0); 
$highestRow = $sheet->getHighestRow(); 
$highestColumn = $sheet->getHighestColumn();

echo $highestRow."<br>";
echo $highestColumn."</br></br>";

for ($row = 5; $row <= 5; $row++){ 
    //  Read a row of data into an array
    $rowData = $sheet->rangeToArray('A'.$row .':' . $highestColumn. $row,
                                    NULL,
                                    TRUE,
                                    FALSE);
									
	$t1_booth = $rowData[0][0];	if(empty($t1_booth)){$t1_booth = 0;}
	$t1_one = $rowData[0][1];	if(empty($t1_one)){$t1_one = 0;}
	$t1_two = $rowData[0][2];	if(empty($t1_two)){$t1_two = 0;}
	$t1_three = $rowData[0][3];	if(empty($t1_three)){$t1_three = 0;}
	$t1_others = $rowData[0][4];if(empty($t1_others)){$t1_others = 0;}	
	$t1_total = $rowData[0][5];	if(empty($t1_total)){$t1_total = 0;}		
	
	$t2_booth = $rowData[0][6];	if(empty($t2_booth)){$t2_booth = 0;}
	$t2_one = $rowData[0][7];	if(empty($t2_one)){$t2_one = 0;}
	$t2_two = $rowData[0][8];	if(empty($t2_two)){$t2_two = 0;}
	$t2_three = $rowData[0][9];	if(empty($t2_three)){$t2_three = 0;}
	$t2_others = $rowData[0][10];if(empty($t2_others)){$t2_others = 0;}		
	$t2_total = $rowData[0][11];if(empty($t2_total)){$t2_total = 0;}		
	
	$t3_booth = $rowData[0][12]; if(empty($t3_booth)){$t3_booth = 0;}	
	$t3_one = $rowData[0][13];	 if(empty($t3_one)){$t3_one = 0;}
	$t3_two = $rowData[0][14];	 if(empty($t3_two)){$t3_two = 0;}
	$t3_three = $rowData[0][15]; if(empty($t3_three)){$t3_three = 0;}	
	$t3_others = $rowData[0][16];if(empty($t3_others)){$t3_others = 0;}		
	$t3_total = $rowData[0][17];if(empty($t3_total)){$t3_total = 0;}				
	
	$t4_booth = $rowData[0][18];if(empty($t4_booth)){$t4_booth = 0;}	
	$t4_one = $rowData[0][19];	if(empty($t4_one)){$t4_one = 0;}
	$t4_two = $rowData[0][20];	if(empty($t4_two)){$t4_two = 0;}
	$t4_three = $rowData[0][21];if(empty($t4_three)){$t4_three = 0;}	
	$t4_others = $rowData[0][22];if(empty($t4_others)){$t4_others = 0;}		
	$t4_total = $rowData[0][23];	if(empty($t4_total)){$t4_total = 0;}		
	
	
	//###########################################################################################################################
	

	//###########################################################################################################################
	
												
}
?>