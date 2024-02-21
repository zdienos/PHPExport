<?php

/**
 * PHPExport for CodeIgniter
 * 
 * dependencies : PHPSpreadsheet
 * 
 * @author zdienos
 * @link https://github.com/zdienos/phpexport 
 * 
 * 
 * 
 * Cara Penggunaan 
 * 1. Load dulu Librarynya :	$this->load->library('PHPExport');
 * 2. Tampung data yang akan dieksport ke array
 * 3. $exportExcel= new PHPExport; 			
 *		$exportExcel
 *			->dataSet($data_set)				: mandatory
 *			->rataTengah('4,5')				: optional (untuk rata tengah field, isikan nomor kolom)
 *			->rataKanan('13')				: optional (untuk rata kanan field, isikan nomor kolom)
 *			->warnaHeader('555555','FFFFFF')		: optional (untuk warna header dan warna font, RBG value)
 *			->excel2003('Laporan-SPK_'.date('YmdHis'));	: mandatory (excel2003/excel2007, isikan nama filenya)
 */


use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;


class PHPExport
{

	public $dataSet;
	public $judulSheet;
	public $warnaHeader;
	public $warnaFontHeader;
	public $rataTengah;
	public $rataKanan;
	public $fieldAccounting;
	public $fieldText;


	private function _nomor_ke_kolom($num)
	{
		$column_name = array();
		for ($i = 0; $i <= $num; $i++) {
			$numeric = $i % 26;
			$letter = chr(65 + $numeric);
			$num2 = intval($i / 26);
			if ($num2 > 0) {
				$v_column = chr(64 + $num2) . $letter;
			} else {
				$v_column = $letter;
			}
			$column_name[] = $v_column;
		}
		return $column_name;
	}

	public function dataSet($dataset)
	{
		$this->dataSet = $dataset;
		return $this;
	}

	public function judulSheet($judulSheet = 'sheet')
	{
		$this->judulSheet = $judulSheet;
		return $this;
	}

	public function warnaHeader($bgColor, $fontColor)
	{
		$this->warnaHeader = $bgColor;
		$this->warnaFontHeader = $fontColor;
		return $this;
	}

	public function rataTengah($arrTengah)
	{
		$this->rataTengah = explode(',', $arrTengah);
		return $this;
	}

	public function rataKanan($arrKanan)
	{
		$this->rataKanan = explode(',', $arrKanan);
		return $this;
	}

	public function fieldAccounting($arrAccounting)
	{
		$this->fieldAccounting = explode(',', $arrAccounting);
		return $this;
	}

	/**
	 * setCellValueExplicit agar tidak dianggap rumus
	 * Untuk field yang diawali tanda sama dengan (=) agar tidak error di phpExcel
	 *
	 * @param string $arrText
	 * @return $this
	 */
	public function fieldText($arrText)
	{
		$this->fieldText = explode(',', $arrText);
		return $this;
	}

	public function generate()
	{
		$objSpreadsheet = new Spreadsheet();

		$data_set  	= $this->dataSet;
		
		if (isset($this->judulSheet)) {
			$sheet_title = $this->judulSheet;
		} else {
			$sheet_title = 'sheet';
		}

		if (!empty($data_set)) {
			//cek ada datanya tidak?
			
			if (count($data_set) > 0) {
				$column_name = array();
				$column_title = array();

				//$objPHPExcel = new PHPExcel();
				$objSpreadsheet
					->getProperties()
					->setCreator("KumalaGroup IT Development")
					->setTitle("KumalaConnect Export");

				$objSheet = $objSpreadsheet->setActiveSheetIndex(0); //inisiasi set object

				$objget = $objSpreadsheet->getActiveSheet();  //inisiasi get object
				$objget->setTitle($sheet_title); //sheet title
				$objget->getDefaultRowDimension()->setRowHeight(-1);

				// TABEL HEADER					
				foreach ($data_set[0] as $key => $value) {
					$column_title[] = strtoupper(str_replace('_', ' ', $key));
				}

				//$column_name = array("A" .. "N");				
				//membuat column name by number
				$column_name = $this->_nomor_ke_kolom(count($column_title) - 1);

				// styling column
				// Beri Warna Title Column
				$v_column_name 	= $column_name[count($column_title) - 1];

				$rata_tengah = array(
					'alignment' => array(
						'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
					)
				);
				$tengah_tengah = array(
					'alignment' => array(
						'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
						'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
					)
				);
				$kiri_tengah = array(
					'alignment' => array(
						'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
						'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
					)
				);
				$kanan_tengah = array(
					'alignment' => array(
						'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
						'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
					)
				);

				//isi header column dengan array
				for ($a = 0; $a < count($column_title); $a++) {
					$objSheet->setCellValue("$column_name[$a]" . "1", "$column_title[$a]");
					//ndak perlu lagi setwidth tiap column, cukup setAutoSize saja
					$objSpreadsheet->getActiveSheet()->getColumnDimension("$column_name[$a]")->setAutoSize(true);
					$objSpreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(30);
				}

				// beri warna header dan rata tengah	
				$objSpreadsheet->setActiveSheetIndex(0)->getStyle("A1:$v_column_name" . "1")->applyFromArray($tengah_tengah);
				$objSpreadsheet->setActiveSheetIndex(0)->getStyle("A1:$v_column_name" . "1")->getFont()->setBold(true)->setSize(12);
				if ((null !== $this->warnaHeader) && (null !== $this->warnaFontHeader)) {
					$warna_header 	=  array(
						'fill' => array(
							'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
							'startColor' => array('rgb' => $this->warnaHeader) //'2,218,240'
						),
						'font' => array(
							'color' => array('rgb' => $this->warnaFontHeader) //'FFFFFF
						)
					);
					$objSpreadsheet->setActiveSheetIndex(0)->getStyle("A1:$v_column_name" . "1")->applyFromArray($warna_header);
					// $objSpreadsheet->setActiveSheetIndex(0)->getStyle("A1:$v_column_name" . "1")->applyFromArray($warna_header);
				}

				$baris = 2;
				foreach ($data_set as $data_item) {
					$col = 0;
					foreach ($data_item as $value) {
						$objSheet->setCellValue("$column_name[$col]" . $baris, $value);
						//rata Kiri
						if (null !== $this->rataTengah) {
							foreach ($this->rataTengah as $c) {
								if ($c == $col) {
									$objSpreadsheet->setActiveSheetIndex(0)->getStyle("$column_name[$c]" . $baris)->applyFromArray($tengah_tengah);
								}
							}
						}
						//rata Kanan			
						if (null !== $this->rataKanan) {
							foreach ($this->rataKanan as $c) {
								if ($c == $col) {
									$objSpreadsheet->setActiveSheetIndex(0)->getStyle("$column_name[$c]" . $baris)->applyFromArray($kanan_tengah);
								}
							}
						}
						//accounting Format		
						if (null !== $this->fieldAccounting) {
							foreach ($this->fieldAccounting as $c) {
								if ($c == $col) {
									$objSpreadsheet->setActiveSheetIndex(0)->getStyle("$column_name[$c]" . $baris)->getNumberFormat()->setFormatCode('_("Rp"* #,##0.00_);_("Rp"* \(#,##0.00\);_("Rp"* "-"??_);_(@_)');
								}
							}
						}
						//text Format		
						if (null !== $this->fieldText) {
							foreach ($this->fieldText as $c) {
								if ($c == $col) {
									$objSpreadsheet->getActiveSheet()->setCellValueExplicit("$column_name[$c]" . $baris, $value, PHPExcel_Cell_DataType::TYPE_STRING);
								}
							}
						}
						$col++;
					}
					$baris++;
				}
				return $objSpreadsheet;
			}
		}
	}

	private function writeToFile($filename, $writerType = 'Xls', $mimes = 'application/vnd.ms-excel')
	{
		// $spreaddd = new Spreadsheet();
		// $writer = new Xlsx($spreaddd);
		
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="'. $filename); 
		header('Cache-Control: max-age=0');

		$writer = IOFactory::createWriter($this->generate(), $writerType);
		$writer->save('php://output');
	}

	public function excel2003($namafile = 'noname')
	{
		$this->writeToFile($namafile . '.xls');
	}

	public function excel2007($namafile = 'noname')
	{
		$this->writeToFile($namafile . '.xlsx', 'Xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	}

	public function csv($namafile = 'noname', $delimiter = ',')
	{
		$spreaddd = new Spreadsheet();
		$writer = new \PhpOffice\PhpSpreadsheet\Writer\Csv($this->generate());

		header("Content-Type: text/csv");
		// header("Content-Type: application/csv");
		header("Content-Disposition: attachment;filename=\"$namafile\".csv");
		header("Cache-Control: max-age=0");

		// $writer = PHPExcel_IOFactory::createWriter($this->generate(), 'Csv');
		$writer->setDelimiter($delimiter);
		$writer->setEnclosure('"');
		$writer->setLineEnding("\r\n");
		$writer->setSheetIndex(0);
		$writer->save('php://output');
	}

}
