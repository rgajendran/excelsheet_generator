<?php

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);
require 'lib/PHPExcel/Classes/PHPExcel.php';
require 'lib/PHPExcel/Classes/PHPExcel/IOFactory.php';
require 'class/ManualStyle.php';
require 'class/Constants.php';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$styles = new ManualStyle();
$booth = 0;

// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
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
$objPHPExcel->getActiveSheet()->setCellValue('A6',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B6',Constants::PARTY_ONE);	//INC
$objPHPExcel->getActiveSheet()->setCellValue('C6',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A6")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A6:C6")->applyFromArray($styles->border());
//2
$objPHPExcel->getActiveSheet()->setCellValue('A7',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B7',Constants::PARTY_TWO);	//BJP
$objPHPExcel->getActiveSheet()->setCellValue('C7',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A7")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A7:C7")->applyFromArray($styles->border());
//3
$objPHPExcel->getActiveSheet()->setCellValue('A8',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B8',Constants::PARTY_THREE);	//JD(S)
$objPHPExcel->getActiveSheet()->setCellValue('C8',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A8")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A8:C8")->applyFromArray($styles->border());
//4
$objPHPExcel->getActiveSheet()->setCellValue('A9',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B9',Constants::PARTY_FOUR);	//Others
$objPHPExcel->getActiveSheet()->setCellValue('C9',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A9")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A9:C9")->applyFromArray($styles->border());

//5
$objPHPExcel->getActiveSheet()->mergeCells('A10:B10');
$objPHPExcel->getActiveSheet()->setCellValue('A10', Constants::PARTY_TITLE_TOTAL);
$objPHPExcel->getActiveSheet()->getStyle("A10")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->setCellValue('C10',100);											//Votes
//##################################TABLE TWO##########################################################################################
$objPHPExcel->getActiveSheet()->setCellValue('E5', Constants::BOOTH); 		//Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('F5', Constants::PARTY); 		//Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('G5', Constants::VOTES); 		//Sub Heading
$objPHPExcel->getActiveSheet()->getStyle("E5:G5")->applyFromArray($styles->subTitlewithB());
$objPHPExcel->getActiveSheet()->getStyle("E10:G10")->applyFromArray($styles->border());
//1
$objPHPExcel->getActiveSheet()->setCellValue('E6',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F6',Constants::PARTY_ONE);	//INC
$objPHPExcel->getActiveSheet()->setCellValue('G6',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E6")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E6:G6")->applyFromArray($styles->border());
//2
$objPHPExcel->getActiveSheet()->setCellValue('E7',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F7',Constants::PARTY_TWO);	//BJP
$objPHPExcel->getActiveSheet()->setCellValue('G7',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E7")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E7:G7")->applyFromArray($styles->border());
//3
$objPHPExcel->getActiveSheet()->setCellValue('E8',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F8',Constants::PARTY_THREE);	//JD(S)
$objPHPExcel->getActiveSheet()->setCellValue('G8',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E8")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E8:G8")->applyFromArray($styles->border());
//4
$objPHPExcel->getActiveSheet()->setCellValue('E9',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F9',Constants::PARTY_FOUR);	//Others
$objPHPExcel->getActiveSheet()->setCellValue('G9',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E9")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E9:G9")->applyFromArray($styles->border());

//5
$objPHPExcel->getActiveSheet()->mergeCells('E10:F10');
$objPHPExcel->getActiveSheet()->setCellValue('E10', Constants::PARTY_TITLE_TOTAL);
$objPHPExcel->getActiveSheet()->getStyle("E10")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->setCellValue('G10',100);											//Votes

//##################################TABLE THREE##########################################################################################
$objPHPExcel->getActiveSheet()->setCellValue('A13', Constants::BOOTH); //Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('B13', Constants::PARTY); //Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('C13', Constants::VOTES); //Sub Heading
$objPHPExcel->getActiveSheet()->getStyle("A13:C13")->applyFromArray($styles->subTitlewithB());
$objPHPExcel->getActiveSheet()->getStyle("A18:C18")->applyFromArray($styles->border());
//1
$objPHPExcel->getActiveSheet()->setCellValue('A14',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B14',Constants::PARTY_ONE);	//INC
$objPHPExcel->getActiveSheet()->setCellValue('C14',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A14")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A14:C14")->applyFromArray($styles->border());
//2
$objPHPExcel->getActiveSheet()->setCellValue('A15',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B15',Constants::PARTY_TWO);	//BJP
$objPHPExcel->getActiveSheet()->setCellValue('C15',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A15")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A15:C15")->applyFromArray($styles->border());
//3
$objPHPExcel->getActiveSheet()->setCellValue('A16',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B16',Constants::PARTY_THREE);	//JD(S)
$objPHPExcel->getActiveSheet()->setCellValue('C16',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A16")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A16:C16")->applyFromArray($styles->border());
//4
$objPHPExcel->getActiveSheet()->setCellValue('A17',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('B17',Constants::PARTY_FOUR);	//Others
$objPHPExcel->getActiveSheet()->setCellValue('C17',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("A17")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("A17:C17")->applyFromArray($styles->border());

//5
$objPHPExcel->getActiveSheet()->mergeCells('A18:B18');
$objPHPExcel->getActiveSheet()->setCellValue('A18', Constants::PARTY_TITLE_TOTAL);
$objPHPExcel->getActiveSheet()->getStyle("A18")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->setCellValue('C18',100);

//##################################TABLE FOUR##########################################################################################
$objPHPExcel->getActiveSheet()->setCellValue('E13', Constants::BOOTH); 		//Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('F13', Constants::PARTY); 		//Sub Heading
$objPHPExcel->getActiveSheet()->setCellValue('G13', Constants::VOTES); 		//Sub Heading
$objPHPExcel->getActiveSheet()->getStyle("E13:G13")->applyFromArray($styles->subTitlewithB());
$objPHPExcel->getActiveSheet()->getStyle("E18:G18")->applyFromArray($styles->border());

//1
$objPHPExcel->getActiveSheet()->setCellValue('E14',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F14',Constants::PARTY_ONE);	//INC
$objPHPExcel->getActiveSheet()->setCellValue('G14',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E14")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E14:G14")->applyFromArray($styles->border());
//2
$objPHPExcel->getActiveSheet()->setCellValue('E15',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F15',Constants::PARTY_TWO);	//BJP
$objPHPExcel->getActiveSheet()->setCellValue('G15',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E15")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E15:G15")->applyFromArray($styles->border());
//3
$objPHPExcel->getActiveSheet()->setCellValue('E16',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F16',Constants::PARTY_THREE);	//JD(S)
$objPHPExcel->getActiveSheet()->setCellValue('G16',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E16")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E16:G16")->applyFromArray($styles->border());
//4
$objPHPExcel->getActiveSheet()->setCellValue('E17',$booth);					//boothname
$objPHPExcel->getActiveSheet()->setCellValue('F17',Constants::PARTY_FOUR);	//Others
$objPHPExcel->getActiveSheet()->setCellValue('G17',1);						//Votes
$objPHPExcel->getActiveSheet()->getStyle("E17")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->getStyle("E17:G17")->applyFromArray($styles->border());

//5
$objPHPExcel->getActiveSheet()->mergeCells('E18:F18');
$objPHPExcel->getActiveSheet()->setCellValue('E18', Constants::PARTY_TITLE_TOTAL);
$objPHPExcel->getActiveSheet()->getStyle("E18")->applyFromArray($styles->centerhorizontal());
$objPHPExcel->getActiveSheet()->setCellValue('G18',100);											//Votes

//Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('Booth - '.$booth);
$objPHPExcel->createSheet();

//Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="name_of_file.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');

?>