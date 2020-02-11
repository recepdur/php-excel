<?php

	require_once ('/Excel/yazma/PHPExcel.php'); 
	$file = "veri.xlsx";
	
	$inputFileType = PHPExcel_IOFactory::identify($file);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);      
    $objReader->setReadDataOnly(true);       
    $objPHPExcel = $objReader->load($file); 			// excel dosyasini load etme
	
    $toplamSayfaSayisi=$objPHPExcel->getSheetCount(); 	// sayfa sayisi    
    $tumSayfaAdlari=$objPHPExcel->getSheetNames(); 	
    $objWorksheet = $objPHPExcel->setActiveSheetIndex(0); 
	
    $satirSayisi = $objWorksheet->getHighestRow(); 
    $sutunSayisi = $objWorksheet->getHighestColumn();  
    $sutunSayisiIndex = PHPExcel_Cell::columnIndexFromString($sutunSayisi);
	
	echo '<table border="1" width="100%" align="left">';
	for ($row = 1; $row <= $satirSayisi; $row++) 
	{
		echo '<tr>';	
		echo '<td>'. $row .')</td>';
		for ($col = 0; $col < $sutunSayisiIndex; $col++) 
		{	
			$value = $objWorksheet->getCellByColumnAndRow($col, $row)->getValue();
			echo '<td>'. $value .'</td>';        		
		}	
		echo '</tr>';
	}
	echo '</table>';

/*
	ini_set("max_execution_time", 0);
	require_once("db.php");	
	
	header("Content-Type: text/html; charset=ISO-8859-9"); 
	require_once ('Excel/okuma/reader.php');
	
	$data = new Spreadsheet_Excel_Reader();
	$data->setUTFEncoder('iconv'); 			// iconv metoduyla dil cevrimini sagliyoruz
	$data->setOutputEncoding('ISO-8859-9'); //turkce dil kodlamasi
	$data->read('veri.xls'); 				// demo.xls dosyasi okunuyor
	
	$satir=$data->sheets[0]['numRows']; 	//satir sayisi
	$sutun=$data->sheets[0]['numCols'];		//sutun sayisi
	echo "<b>Satir ". $satir;
	echo "<br>Sutun ". $sutun ."</b>";

	function excelDateToDate($readDate)
	{
		$phpexcepDate = $readDate-25569; //to offset to Unix epoch
		return strtotime("+$phpexcepDate days", mktime(0,0,0,1,1,1970));
	}
			
		$sayac_isletme=0;
		$sayac_hayvan=0;
		$sayac_asi=0;

	for ($i = 3; $i <= $satir; $i++) 
	{
		if(($i%50) == 0) 
			sleep(4);
			
		$kupe_no = mb_convert_encoding($data->sheets[0]['cells'][$i][1],"utf-8","iso-8859-9");
		$il = mb_convert_encoding($data->sheets[0]['cells'][$i][2],"utf-8","iso-8859-9");
		$ilce = mb_convert_encoding($data->sheets[0]['cells'][$i][3],"utf-8","iso-8859-9");
		$mahalle = mb_convert_encoding($data->sheets[0]['cells'][$i][4],"utf-8","iso-8859-9");
		$isletme_no = mb_convert_encoding($data->sheets[0]['cells'][$i][5],"utf-8","iso-8859-9");
		$sutun6 = mb_convert_encoding($data->sheets[0]['cells'][$i][6],"utf-8","iso-8859-9");
		$dogum_tarih1 = mb_convert_encoding($data->sheets[0]['cells'][$i][7],"utf-8","iso-8859-9");			
		$dogum_tarih = date('Y-m-d', excelDateToDate($dogum_tarih1));
		$irk = mb_convert_encoding($data->sheets[0]['cells'][$i][8],"utf-8","iso-8859-9");
		$tur = mb_convert_encoding($data->sheets[0]['cells'][$i][9],"utf-8","iso-8859-9");
		$durum = mb_convert_encoding($data->sheets[0]['cells'][$i][10],"utf-8","iso-8859-9");
		$sutun11 = mb_convert_encoding($data->sheets[0]['cells'][$i][11],"utf-8","iso-8859-9");
		$sahibi = mb_convert_encoding($data->sheets[0]['cells'][$i][12],"utf-8","iso-8859-9");
		$birlik = mb_convert_encoding($data->sheets[0]['cells'][$i][13],"utf-8","iso-8859-9");
		$asi_tarih1 = mb_convert_encoding($data->sheets[0]['cells'][$i][14],"utf-8","iso-8859-9");
		$asi_tarih = date('Y-m-d', excelDateToDate($asi_tarih1));
		$boga = mb_convert_encoding($data->sheets[0]['cells'][$i][15],"utf-8","iso-8859-9");
		$boga_irk = mb_convert_encoding($data->sheets[0]['cells'][$i][16],"utf-8","iso-8859-9");
		$asi_sayisi = mb_convert_encoding($data->sheets[0]['cells'][$i][17],"utf-8","iso-8859-9");
		$belge_no = mb_convert_encoding($data->sheets[0]['cells'][$i][18],"utf-8","iso-8859-9");

		
		$sorgu = $DB->query("SELECT * FROM vtr_isletme WHERE isletme_no='". $isletme_no ."' "); 				// isletme no
		if($row=$DB->farray($sorgu))
		{					
			$sorgu2 = $DB->query("SELECT * FROM vtr_isletme_hayvan WHERE kupe_no='". $kupe_no ."' ");		// kupe no
			if($row2=$DB->farray($sorgu2))
			{		
				$sonuc=$DB->query("INSERT INTO vtr_isletme_hayvan_asi (vtr_isletme_hayvan_id,asi_tarih,boga, boga_irk,asi_sayisi,belge_no) VALUES('".$row2["ID"]."','".$asi_tarih."','".$boga."','".$boga_irk."','".$asi_sayisi."','".$belge_no."')");
				$sayac_asi++;
			}else
			{
				$sonuc=$DB->query("INSERT INTO vtr_isletme_hayvan (vtr_isletme_id,kupe_no,irk,dogum_tarih,sutun6,sutun11,tur,durumu,birlik,il,ilce,mahalle) VALUES('".$row["ID"]."','".$kupe_no."','".$irk."','".$dogum_tarih."','".$sutun6."','".$sutun11."','".$tur."','".$durum."','".$birlik."','".$il."','".$ilce."','".$mahalle."')");
				$isletme_hayvan_id = $DB->insert();
				$sonuc=$DB->query("INSERT INTO vtr_isletme_hayvan_asi (vtr_isletme_hayvan_id,asi_tarih,boga, boga_irk,asi_sayisi,belge_no) VALUES('".$isletme_hayvan_id."','".$asi_tarih."','".$boga."','".$boga_irk."','".$asi_sayisi."','".$belge_no."')");
				$sayac_hayvan++;
				$sayac_asi++;
			}			
		}else
		{
			$sonuc=$DB->query("INSERT INTO vtr_isletme (vtr_id, isletme_no, sahibi) VALUES('"."1"."','".$isletme_no."','".$sahibi."')");
			$isletme_id = $DB->insert();
			$sonuc=$DB->query("INSERT INTO vtr_isletme_hayvan (vtr_isletme_id,kupe_no,irk,dogum_tarih,sutun6,sutun11,tur,durumu,birlik,il,ilce,mahalle) VALUES('".$isletme_id."','".$kupe_no."','".$irk."','".$dogum_tarih."','".$sutun6."','".$sutun11."','".$tur."','".$durum."','".$birlik."','".$il."','".$ilce."','".$mahalle."')");
			$isletme_hayvan_id = $DB->insert();
			$sonuc=$DB->query("INSERT INTO vtr_isletme_hayvan_asi (vtr_isletme_hayvan_id,asi_tarih,boga, boga_irk,asi_sayisi,belge_no) VALUES('".$isletme_hayvan_id."','".$asi_tarih."','".$boga."','".$boga_irk."','".$asi_sayisi."','".$belge_no."')");
			$sayac_isletme++;
			$sayac_hayvan++;
			$sayac_asi++;		
		}
	}
	
	echo "<br><br><b>isletme sayisi: ". $sayac_isletme;
	echo "<br>hayvan sayisi: ". $sayac_hayvan;
	echo "<br>asi sayisi : ". $sayac_asi;	
	echo ' <br><br> Kayit ekleme bitti </b>';
*/
	
/*
	echo '<table border="1" width="100%" align="left">';
	for ($i = 1; $i <= $satir; $i++) 
	{
		echo '<tr>';	
		for ($j = 1; $j <= $sutun; $j++) 
		{	
			echo '<td>'.$data->sheets[0]['cells'][$i][$j].'</td>';        		
		}	
		echo '</tr>';
	}
	echo '</table>';
*/
?>

