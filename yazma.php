
<?php

	require_once ('/Excel/yazma/PHPExcel.php');

	$objPHPExcel = new PHPExcel();		//create obj
	//load 
/*	$objPHPExcel = PHPExcel_IOFactory::load("yaz1.xlsx");
	$objPHPExcel->setActiveSheetIndex(0);
	$row = $objPHPExcel->getActiveSheet()->getHighestRow()+1;
*/


	//sayfa1
		$objPHPExcel->createSheet(NULL, 0);
		$objPHPExcel->setActiveSheetIndex(0);
		$objPHPExcel->getActiveSheet()->setTitle('Sayfa1');		// sayfa

		$objPHPExcel->getActiveSheet()->getPageMargins()->setTop(0.5/1);
		$objPHPExcel->getActiveSheet()->getPageMargins()->setBottom(0.5/2.54);
		$objPHPExcel->getActiveSheet()->getPageMargins()->setLeft(0.5/2.54);
		$objPHPExcel->getActiveSheet()->getPageMargins()->setRight(0.5/2.54);

		$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(4.30);
		$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(11);
		$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(16);
		$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(13);
		$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(6);
		$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(14);
		$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(14);
		$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(7);
		$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(18);
		$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(8);
		$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10);
		$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(7);
		$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(5);
		
		$objPHPExcel->getActiveSheet()->getRowDimension('4')->setRowHeight(25);
		$objPHPExcel->getActiveSheet()->getStyle('H3:H4')->getAlignment()->setWrapText(true); 
		$objPHPExcel->getActiveSheet()->getStyle('K3:K4')->getAlignment()->setWrapText(true); 
		$objPHPExcel->getActiveSheet()->getStyle('L3:L4')->getAlignment()->setWrapText(true); 
		$objPHPExcel->getActiveSheet()->getStyle('M3:M4')->getAlignment()->setWrapText(true); 

		$objPHPExcel->getActiveSheet()->getStyle("A1:R50")->getFont()->setSize(10);		// font size
		$objPHPExcel->getActiveSheet()->getStyle('A1:M29')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	
		$objPHPExcel->getActiveSheet()->mergeCells('A1:M1');	
		$objPHPExcel->getActiveSheet()->getStyle('A1:M1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);		
		$objPHPExcel->getActiveSheet()->setCellValue('A1', 'SUNİ TOHUMLAMA UYGULAMA CETVELİ');
		
		$objPHPExcel->getActiveSheet()->mergeCells('A2:B2');	
		$objPHPExcel->getActiveSheet()->getStyle('A2:M2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('A2', 'İLÇESİ :ÇUMRA');
		
		$objPHPExcel->getActiveSheet()->mergeCells('C2:I2');	
		$objPHPExcel->getActiveSheet()->getStyle('C2:I2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		
		$objPHPExcel->getActiveSheet()->mergeCells('J2:M2');	
		$objPHPExcel->getActiveSheet()->getStyle('J2:M2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('J2', 'GEÇERLİ OLDUĞU AY: 6');
		
		$objPHPExcel->getActiveSheet()->mergeCells('A3:A4');	
		$objPHPExcel->getActiveSheet()->getStyle('A3:A4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('A3', 'S.No');
		
		$objPHPExcel->getActiveSheet()->mergeCells('B3:B4');	
		$objPHPExcel->getActiveSheet()->getStyle('B3:B4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('B3', 'İŞLETME NO');
		
		$objPHPExcel->getActiveSheet()->mergeCells('C3:E3');	
		$objPHPExcel->getActiveSheet()->getStyle('C3:E3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('C3', 'İŞLETME SAHİBİNİN');
		
		$objPHPExcel->getActiveSheet()->setCellValue('C4', 'ADI SOYADI');
		$objPHPExcel->getActiveSheet()->setCellValue('D4', 'KÖY');
		$objPHPExcel->getActiveSheet()->setCellValue('E4', 'B.ÜYE');
		
		$objPHPExcel->getActiveSheet()->mergeCells('F3:H3');	
		$objPHPExcel->getActiveSheet()->getStyle('F3:H3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('F3', 'SUNİ TOHUMLAMA YAPILAN HAYVANIN');
		
		$objPHPExcel->getActiveSheet()->setCellValue('F4', 'KULAK NO');
		$objPHPExcel->getActiveSheet()->setCellValue('G4', 'IRKI');
		$objPHPExcel->getActiveSheet()->setCellValue('H4', 'DOĞUM TARİHİ');
		
		$objPHPExcel->getActiveSheet()->mergeCells('I3:J3');	
		$objPHPExcel->getActiveSheet()->getStyle('I3:J3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('I3', 'SUNİ TOHUMLAMA BOĞASININ');
		
		$objPHPExcel->getActiveSheet()->setCellValue('I4', 'ADI ve KULAK NO');
		$objPHPExcel->getActiveSheet()->setCellValue('J4', 'IRKI');
		
		$objPHPExcel->getActiveSheet()->mergeCells('K3:K4');	
		$objPHPExcel->getActiveSheet()->getStyle('K3:K4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('K3', 'TOHUMLAMA TARİHİ');
		
		$objPHPExcel->getActiveSheet()->mergeCells('L3:L4');	
		$objPHPExcel->getActiveSheet()->getStyle('L3:L4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('L3', 'BELGE NO');
		
		$objPHPExcel->getActiveSheet()->mergeCells('M3:M4');	
		$objPHPExcel->getActiveSheet()->getStyle('M3:M4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('M3', '1.2.3 TOH');

		$objPHPExcel->getActiveSheet()->mergeCells('B31:C31');	
		$objPHPExcel->getActiveSheet()->getStyle('B31:C31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('B31', 'TOHUMLAMA YAPAN');
		
		$objPHPExcel->getActiveSheet()->setCellValue('B33', 'ADI SOYADI:');
		$objPHPExcel->getActiveSheet()->setCellValue('B34', 'KOD NO:');
	
		$objPHPExcel->getActiveSheet()->mergeCells('G31:H31');	
		$objPHPExcel->getActiveSheet()->getStyle('G31:H31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('G31', 'KONTROL EDEN');
		
		$objPHPExcel->getActiveSheet()->mergeCells('J31:K31');	
		$objPHPExcel->getActiveSheet()->getStyle('J31:K31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->setCellValue('J31', 'TASDİK OLUNUR');
		
		
	//Kaydet
	$Kaydet = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$Kaydet->save("yaz1.xlsx");
	echo "Dosya Olusturuldu";

	
	
	
	
	
	
	
	
	

/*
	$dosya = "yaz1.xls"; 
	$yaz = @fopen($dosya,'w+'); 
	 
	fwrite($yaz,"Ad\t Soyad\t Bolum\t Email\t Telefon\t \n");

	$sutun0 = mb_convert_encoding("ığüjşiöçIĞÜJŞİÖÇ", "iso-8859-9", "UTF-8");
	$sutun1 = "bb";
	$sutun2 = "cc";
	$sutun3 = "dd";
	$sutun4 = "ee"; 
	 
	fwrite($yaz,"$sutun0\t $sutun1\t $sutun2\t $sutun3\t $sutun4\t\n"); 
	 
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="yaz1.xls"'); 
	header ('Content-Transfer-Encoding: binary');
	header('Cache-Control: max-age=0');
	$dosyayolu = "yaz1.xls"; // <-- düzenleyin
	readfile($dosyayolu);
 */
?>

