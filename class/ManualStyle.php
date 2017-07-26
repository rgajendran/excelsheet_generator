<?php

class ManualStyle{

	public function sheading(){
			$sheading = array(
		    'font'  => array(
		        'bold'  => true,
		      	'color' => array('rgb' => '000000'),
		        'size'  => '18',
		        'name'  => 'Calibri'
		    ),
			'alignment' => array(
		            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		        )
			);	
			
			return $sheading;
	}
	//styles
	
	public function ssubheading(){
		$ssubheading = array(
		    'font'  => array(
		        'bold'  => true,
		      	'color' => array('rgb' => '000000'),
		        'size'  => '13',
		        'name'  => 'Calibri'
		    ),
			'alignment' => array(
		            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		        )
		);
		return $ssubheading;
	}

	public function title(){
		$title = array(
		    'font'  => array(
		        'bold'  => true,
		      	'color' => array('rgb' => '000000'),
		        'name'  => 'Calibri'
		    ),
			'alignment' => array(
		            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		    ),
		    'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => 'FFB266')
		    )
		);
		return $title;
	}

	public function titlewithB(){
		$titlewithB = array(
		    'font'  => array(
		        'bold'  => true,
		      	'color' => array('rgb' => '000000'),
		        'size'  => '13',
		        'name'  => 'Calibri'
		    ),
			'alignment' => array(
		            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		    ),
		    'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => 'FFB266')
		    ),
		    'borders' => array(
			    'allborders' => array(
			      	'style' => PHPExcel_Style_Border::BORDER_THIN
			    )
		  	 )
		);
		
		return $titlewithB;
	}

	public function subTitlewithB(){
		$subTitlewithB = array(
		    'font'  => array(
		      	'color' => array('rgb' => '000000'),
		        'name'  => 'Calibri'
		    ),
			'alignment' => array(
		            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		    ),
		    'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => 'E0E0E0')
		    ),
		    'borders' => array(
			    'allborders' => array(
			      	'style' => PHPExcel_Style_Border::BORDER_THIN
			    )
		  	 )
		);		
		return $subTitlewithB;
	}

	public function border(){
		$border = array(
		    'borders' => array(
			    'allborders' => array(
			      	'style' => PHPExcel_Style_Border::BORDER_THIN
			    )
		  	)
		);	
		return $border;	
	}
	
	public function centerhorizontal(){
			$sheading = array(
			'alignment' => array(
		            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		        )
			);	
			
			return $sheading;
	}

}


?>