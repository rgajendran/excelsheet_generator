<?php

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);
require_once 'lib/PHPExcel/Classes/PHPExcel.php';
require_once 'lib/PHPExcel/Classes/PHPExcel/IOFactory.php';
require 'class/ManualStyle.php';
require 'class/Constants.php';

$inputFileName = 'nippani.xlsx';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$styles = new ManualStyle();

try {
    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($inputFileName);
} catch(Exception $e) {
    die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}

$sheet = $objPHPExcel->getSheet(0); 
$highestRow = $sheet->getHighestRow(); 
$highestColumn = $sheet->getHighestColumn();

//***********************************************************************************************************************************

for ($row = 198; $row <= 263; $row++){ 
    //  Read a row of data into an array
    $rowData = $sheet->rangeToArray('A'.$row .':' . $highestColumn. $row,
                                    NULL,
                                    TRUE,
                                    FALSE);
	$n = $row - 198;																	
	$t1_booth = $rowData[0][0];	 if(empty($t1_booth)){$t1_booth = 0;}
	$t1_one = $rowData[0][1];	 if(empty($t1_one)){$t1_one = 0;}
	$t1_two = $rowData[0][2];	 if(empty($t1_two)){$t1_two = 0;}
	$t1_three = $rowData[0][3];	 if(empty($t1_three)){$t1_three = 0;}
	$t1_others = $rowData[0][4]; if(empty($t1_others)){$t1_others = 0;}	
	$t1_total = $rowData[0][5];	 if(empty($t1_total)){$t1_total = 0;}		
	
	$t2_booth = $rowData[0][6];	 if(empty($t2_booth)){$t2_booth = 0;}
	$t2_one = $rowData[0][7];	 if(empty($t2_one)){$t2_one = 0;}
	$t2_two = $rowData[0][8];	 if(empty($t2_two)){$t2_two = 0;}
	$t2_three = $rowData[0][9];	 if(empty($t2_three)){$t2_three = 0;}
	$t2_others = $rowData[0][10];if(empty($t2_others)){$t2_others = 0;}		
	$t2_total = $rowData[0][11]; if(empty($t2_total)){$t2_total = 0;}		
	
	$t3_booth = $rowData[0][12]; if(empty($t3_booth)){$t3_booth = 0;}	
	$t3_one = $rowData[0][13];	 if(empty($t3_one)){$t3_one = 0;}
	$t3_two = $rowData[0][14];	 if(empty($t3_two)){$t3_two = 0;}
	$t3_three = $rowData[0][15]; if(empty($t3_three)){$t3_three = 0;}	
	$t3_others = $rowData[0][16];if(empty($t3_others)){$t3_others = 0;}		
	$t3_total = $rowData[0][17]; if(empty($t3_total)){$t3_total = 0;}				
	
	$t4_booth = $rowData[0][18]; if(empty($t4_booth)){$t4_booth = 0;}	
	$t4_one = $rowData[0][19];	 if(empty($t4_one)){$t4_one = 0;}
	$t4_two = $rowData[0][20];	 if(empty($t4_two)){$t4_two = 0;}
	$t4_three = $rowData[0][21]; if(empty($t4_three)){$t4_three = 0;}	
	$t4_others = $rowData[0][22];if(empty($t4_others)){$t4_others = 0;}		
	$t4_total = $rowData[0][23]; if(empty($t4_total)){$t4_total = 0;}	
	
	if(!empty($t1_booth)){
		$booth = $t1_booth;
	}else if(!empty($t2_booth)){
		$booth = $t2_booth;
	}else if(!empty($t3_booth)){
		$booth = $t3_booth;
	}else if(!empty($t4_booth)){
		$booth = $t4_booth;
	}

// Create a first sheet, representing sales data
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex($n+1);	
$objPHPExcel->getActiveSheet()->mergeCells('C1:E1');
$objPHPExcel->getActiveSheet()->setCellValue('C1', Constants::HEADING);
$objPHPExcel->getActiveSheet()->getStyle("C1")->applyFromArray($styles->sheading());

//Line 2
$objPHPExcel->getActiveSheet()->mergeCells('C2:E2');
$objPHPExcel->getActiveSheet()->setCellValue('C2', Constants::SUBHEADING); //Sub Heading
$objPHPExcel->getActiveSheet()->getStyle("C2")->applyFromArray($styles->ssubheading());


//Line 3 heading 
$objPHPExcel->getActiveSheet()->mergeCells('A4:C4');
$objPHPExcel->getActiveSheet()->setCellValue('A4', Constants::ONETITLE); //Title ONE Table
$objPHPExcel->getActiveSheet()->getStyle("A4")->applyFromArray($styles->title());

$objPHPExcel->getActiveSheet()->mergeCells('E4:G4');
$objPHPExcel->getActiveSheet()->setCellValue('E4', Constants::TWOTITLE); //Title TWO Table
$objPHPExcel->getActiveSheet()->getStyle("E4")->applyFromArray($styles->title());

$objPHPExcel->getActiveSheet()->mergeCells('A12:C12');
$objPHPExcel->getActiveSheet()->setCellValue('A12', Constants::THREETITLE); //Title THREE Table
$objPHPExcel->getActiveSheet()->getStyle("A12")->applyFromArray($styles->title());

$objPHPExcel->getActiveSheet()->mergeCells('E12:G12');
$objPHPExcel->getActiveSheet()->setCellValue('E12', Constants::FOURTITLE); //Title FOUR Table
$objPHPExcel->getActiveSheet()->getStyle("E12")->applyFromArray($styles->title());

//Set Width
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
//Line 4 SubHeading 
//##################################TABLE ONE##########################################################################################
$objPHPExcel->getActiveSheet()->setCellValue('A5', Constants::BOOTH); //Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('B5', Constants::PARTY); //Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('C5', Constants::VOTES); //Sub Heading
$objPHPExcel->getActiveSheet()->getStyle("A5:C5")->applyFromArray($styles->subTitlewithB());
$objPHPExcel->getActiveSheet()->getStyle("A10:C10")->applyFromArray($styles->border());
//1
$objPHPExcel->getActiveSheet()->setCellValue('A6',$t1_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B6',Constants::PARTY_ONE);	//INC
$objPHPExcel->getActiveSheet()->setCellValue('C6',$t1_one);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A6")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A6:C6")->applyFromArray($styles->border());
//2
$objPHPExcel->getActiveSheet()->setCellValue('A7',$t1_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B7',Constants::PARTY_TWO);	//BJP
$objPHPExcel->getActiveSheet()->setCellValue('C7',$t1_two);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A7")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A7:C7")->applyFromArray($styles->border());
//3
$objPHPExcel->getActiveSheet()->setCellValue('A8',$t1_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B8',Constants::PARTY_THREE);	//JD(S)
$objPHPExcel->getActiveSheet()->setCellValue('C8',$t1_three);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A8")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A8:C8")->applyFromArray($styles->border());
//4
$objPHPExcel->getActiveSheet()->setCellValue('A9',$t1_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B9',Constants::PARTY_FOUR);	//Others
$objPHPExcel->getActiveSheet()->setCellValue('C9',$t1_others);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A9")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A9:C9")->applyFromArray($styles->border());

//5
$objPHPExcel->getActiveSheet()->mergeCells('A10:B10');
$objPHPExcel->getActiveSheet()->setCellValue('A10', Constants::PARTY_TITLE_TOTAL);
$objPHPExcel->getActiveSheet()->getStyle("A10")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->setCellValue('C10',$t1_total);											//Votes
//##################################TABLE TWO##########################################################################################
$objPHPExcel->getActiveSheet()->setCellValue('E5', Constants::BOOTH); 		//Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('F5', Constants::PARTY); 		//Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('G5', Constants::VOTES); 		//Sub Heading
$objPHPExcel->getActiveSheet()->getStyle("E5:G5")->applyFromArray($styles->subTitlewithB());
$objPHPExcel->getActiveSheet()->getStyle("E10:G10")->applyFromArray($styles->border());
//1
$objPHPExcel->getActiveSheet()->setCellValue('E6',$t2_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F6',Constants::PARTY_ONE);	//INC
$objPHPExcel->getActiveSheet()->setCellValue('G6',$t2_one);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E6")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E6:G6")->applyFromArray($styles->border());
//2
$objPHPExcel->getActiveSheet()->setCellValue('E7',$t2_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F7',Constants::PARTY_TWO);	//BJP
$objPHPExcel->getActiveSheet()->setCellValue('G7',$t2_two);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E7")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E7:G7")->applyFromArray($styles->border());
//3
$objPHPExcel->getActiveSheet()->setCellValue('E8',$t2_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F8',Constants::PARTY_THREE);	//JD(S)
$objPHPExcel->getActiveSheet()->setCellValue('G8',$t2_three);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E8")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E8:G8")->applyFromArray($styles->border());
//4
$objPHPExcel->getActiveSheet()->setCellValue('E9',$t2_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F9',Constants::PARTY_FOUR);	//Others
$objPHPExcel->getActiveSheet()->setCellValue('G9',$t2_others);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E9")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E9:G9")->applyFromArray($styles->border());

//5
$objPHPExcel->getActiveSheet()->mergeCells('E10:F10');
$objPHPExcel->getActiveSheet()->setCellValue('E10', Constants::PARTY_TITLE_TOTAL);
$objPHPExcel->getActiveSheet()->getStyle("E10")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->setCellValue('G10',$t2_total);											//Votes

//##################################TABLE THREE##########################################################################################
$objPHPExcel->getActiveSheet()->setCellValue('A13', Constants::BOOTH); //Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('B13', Constants::PARTY); //Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('C13', Constants::VOTES); //Sub Heading
$objPHPExcel->getActiveSheet()->getStyle("A13:C13")->applyFromArray($styles->subTitlewithB());
$objPHPExcel->getActiveSheet()->getStyle("A18:C18")->applyFromArray($styles->border());
//1
$objPHPExcel->getActiveSheet()->setCellValue('A14',$t3_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B14',Constants::PARTY_ONE);	//INC
$objPHPExcel->getActiveSheet()->setCellValue('C14',$t3_one);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A14")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A14:C14")->applyFromArray($styles->border());
//2
$objPHPExcel->getActiveSheet()->setCellValue('A15',$t3_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B15',Constants::PARTY_TWO);	//BJP
$objPHPExcel->getActiveSheet()->setCellValue('C15',$t3_two);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A15")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A15:C15")->applyFromArray($styles->border());
//3
$objPHPExcel->getActiveSheet()->setCellValue('A16',$t3_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B16',Constants::PARTY_THREE);	//JD(S)
$objPHPExcel->getActiveSheet()->setCellValue('C16',$t3_three);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A16")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A16:C16")->applyFromArray($styles->border());
//4
$objPHPExcel->getActiveSheet()->setCellValue('A17',$t3_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B17',Constants::PARTY_FOUR);	//Others
$objPHPExcel->getActiveSheet()->setCellValue('C17',$t3_others);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A17")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A17:C17")->applyFromArray($styles->border());

//5
$objPHPExcel->getActiveSheet()->mergeCells('A18:B18');
$objPHPExcel->getActiveSheet()->setCellValue('A18', Constants::PARTY_TITLE_TOTAL);
$objPHPExcel->getActiveSheet()->getStyle("A18")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->setCellValue('C18',$t3_total);

//##################################TABLE FOUR##########################################################################################
$objPHPExcel->getActiveSheet()->setCellValue('E13', Constants::BOOTH); 		//Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('F13', Constants::PARTY); 		//Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('G13', Constants::VOTES); 		//Sub Heading
$objPHPExcel->getActiveSheet()->getStyle("E13:G13")->applyFromArray($styles->subTitlewithB());
$objPHPExcel->getActiveSheet()->getStyle("E18:G18")->applyFromArray($styles->border());

//1
$objPHPExcel->getActiveSheet()->setCellValue('E14',$t4_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F14',Constants::PARTY_ONE);	//INC
$objPHPExcel->getActiveSheet()->setCellValue('G14',$t4_one);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E14")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E14:G14")->applyFromArray($styles->border());
//2
$objPHPExcel->getActiveSheet()->setCellValue('E15',$t4_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F15',Constants::PARTY_TWO);	//BJP
$objPHPExcel->getActiveSheet()->setCellValue('G15',$t4_two);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E15")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E15:G15")->applyFromArray($styles->border());
//3
$objPHPExcel->getActiveSheet()->setCellValue('E16',$t4_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F16',Constants::PARTY_THREE);	//JD(S)
$objPHPExcel->getActiveSheet()->setCellValue('G16',$t4_three);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E16")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E16:G16")->applyFromArray($styles->border());
//4
$objPHPExcel->getActiveSheet()->setCellValue('E17',$t4_booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F17',Constants::PARTY_FOUR);	//Others
$objPHPExcel->getActiveSheet()->setCellValue('G17',$t4_others);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E17")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E17:G17")->applyFromArray($styles->border());

//5
$objPHPExcel->getActiveSheet()->mergeCells('E18:F18');
$objPHPExcel->getActiveSheet()->setCellValue('E18', Constants::PARTY_TITLE_TOTAL);
$objPHPExcel->getActiveSheet()->getStyle("E18")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->setCellValue('G18',$t4_total);											//Votes

//Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('Booth-'.$booth);

//***********************************************************************************************************************************

}
//Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="name_of_file.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');

?>