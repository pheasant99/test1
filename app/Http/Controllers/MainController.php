<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use \App\Lib\excel2pdf;
use \PhpOffice\PhpSpreadsheet\IOFactory;

class MainController extends Controller
{
	public function index3()
	{
		$fname		= '..\storage\app\public/001_納品書_タテ型.xlsx';
//@		$fn			= '..\storage\app\public/temp0000.xlsx';
		$filename	= 'template.xlsx';

		$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
		$book	= $reader->load($fname);

		$writer = IOFactory::createWriter($book, 'Xlsx');

//@		$writer->save($fn);
//@
//@		$fn			= $fname;
//@		
//@		$file_size	= filesize($fn);

		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="'.$filename.'"');
//@		header("Content-Length: {$file_size}");
		header('Cache-Control: max-age=0');
		header('Cache-Control: max-age=1');
	//	header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
		header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
		header('Cache-Control: cache, must-revalidate');
		header('Pragma: public');

	
//@		readfile($fn);
		$writer->save('php://output');
exit();
		
//		$writer = new Xlsx($book);
//		$writer->save('php://output');
 

//		$extension	= pathinfo($fname)['extension'];
//		$file_size	= filesize($fname);


/*
		header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		header("Content-Length: {$file_size}");
		header("Content-Disposition: attachment; filename='{$filename}.{$extension}'");
	
		readfile($fname);
*/
	}
	public function index()
	{
		$ex	=new excel2pdf();
		$fname	= '..\storage\app\public/001_納品書_タテ型.xlsx';
	//	$fname	= '..\storage\app\public/test.xls';
	//	$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
	//	$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xls();
		$book	= null;
	//	$book	= $reader->load($fname);
		
		$ex->setExcelFilename($fname);
		
		$cw	= array();
/*		$cw[1]	= 118 * 1.1;
		$cw[2]	= 100 * 1.1;
		$cw[3]	=  42 * 1.1;
		$cw[4]	=  27 * 1.1;
		$cw[5]	= 132 * 1.1;
*/
/*		$cw[1]	= 118 * 1.1;
		$cw[2]	= 206 * 1.1;
		$cw[3]	=  73 * 1.1;
		$cw[4]	=  58 * 1.1;
		$cw[5]	=  18 * 1.1;
		$cw[6]	=  43 * 1.1;
		$cw[7]	=  28 * 1.1;
		$cw[8]	= 175 * 1.1;
*/
		for($i=1;$i<20;$i++) {
//			$cw[$i]	= 175;				//px
//			$cw[$i]	= 19.5;//25.4;		//mm
			$cw[$i]	= 11.0;				//mm
		}
//		$ex->setColumnWidthsPx($cw);
		$ex->setColumnWidthsmm($cw);
		
		$sheet	=null;
		if($book != null) {
			// シートが1枚の場合
			$sheet = $book->getSheet(0);
			//セルに値をセットしてみる（自動計算されるか）
		//	$sheet->setCellValue('C18', 1500);
		}
		
		$sheet	= $ex->getSheet();
		if($sheet != null) {
			$ex->setSheet($sheet);
			
/*	* /
		//	$garr	= $ex->sheet->getDrawingCollection();
			$garr	= $sheet->getDrawingCollection();
			$c	= count($garr);
			
			echo("count=${c}<br>");
			
			foreach($garr as $obj){
				$nm	= $obj->getName();
				$cn	= $obj->getCoordinates();
				$x	= $obj->getOffsetX();
				$y	= $obj->getOffsetY();
				$str	= "<br> position ${x} ${y} name:${nm} coordinates:${cn}";
				echo($str);
			}
			exit("<br>end<br");
/**/
			$ex->writePDF();
		/*	*/
		//	var_dump($ex->clmWidthPt);		//カラム幅の初期値
		//	echo "<br><br>";
		//	$ex->debugCell(0,0);		//カラム幅、行高さの出力
		/*	*/
			for($r=1;$r<27;$r++) {
				for($c=1;$c<18;$c++) {
					$ex->debugCell($c,$r);
				}
			}
		/*	*/
		}
	}
	
	//-----------------------------------
	public function index2()
	{
		// tinkerによるデバッグ
	//	eval(\Psy\sh());
echo "<br>BORDER_NONE           ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE;
echo "<br>BORDER_DASHDOT        ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOT;
echo "<br>BORDER_DASHDOTDOT     ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOTDOT;
echo "<br>BORDER_DASHED         ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHED;
echo "<br>BORDER_DOTTED         ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOTTED;
echo "<br>BORDER_DOUBLE         ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE;
echo "<br>BORDER_HAIR           ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR;
echo "<br>BORDER_MEDIUM         ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM;
echo "<br>BORDER_MEDIUMDASHDOT     ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOT;
echo "<br>BORDER_MEDIUMDASHDOTDOT  ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOTDOT;
echo "<br>BORDER_MEDIUMDASHED      ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED;
echo "<br>BORDER_SLANTDASHDOT      ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_SLANTDASHDOT;
echo "<br>BORDER_THICK             ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK;
echo "<br>BORDER_THIN              ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
echo "<br><br><br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_GENERAL 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER_CONTINUOUS 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_JUSTIFY 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_FILL 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_DISTRIBUTED 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_JUSTIFY 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_DISTRIBUTED 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::READORDER_CONTEXT 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::READORDER_LTR 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::READORDER_RTL 	."<br>";
echo "<br><br><br>";


echo "--------------------<br>";

		
		
		$ex	=new excel2pdf();
		
		$ret	= $ex->area2index('A1:d5');
		var_dump($ret);
		
		echo "<br><br><br><br>";
		
//		try {
		//	$reader = Excel::load('/home/matsu/Laravel/test1/tests/test.xlsx');
			$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
		//	$book	= $reader->load('C:\Users\matsumoto\Documents\業務\ララベル\test1\storage\app\public\test.xlsx');
			$book	= $reader->load('C:/Users/matsumoto/Documents/業務\ララベル/test1/storage/app/public/test.xlsx');
			
			if ($book != null) {
				// シートが1枚の場合
				$sheet = $book->getSheet(0);
				
//							$h		= $sheet->getRowDimension(3)->getRowHeight();
//							$w		= $sheet->getColumnDimension('A')->getWidth();
				
				$rdim	= $sheet->getRowDimensions();
				$c	= count($rdim);
				echo "行 : ${c}<br><br>";
				for($r=1;$r<$c;$r++) {
					if(isset($rdim[$r])) {
						$v	= $rdim[$r]->getRowHeight();
						echo "(${r}):${v} ";
					}
				}
				echo "<br><br>";
				
				
/*	*/			for($r=1;$r<36;$r++) {
					for($c=1;$c<10;$c++) {
						$cell	= $sheet->getCellByColumnAndRow($c,$r,false);
						if($cell!=null) {
						//	$val	= $cell->getValue();
							$val	= $cell->getFormattedValue();
							
							$mg		= $cell->isMergeRangeValueCell();	//代表セル
							$marge	= $cell->getMergeRange();
							$style	= $cell->getStyle();
							$brders	= $style->getBorders();
							$bdr	= $brders->getBottom();
							$bline	= $bdr->getBorderStyle();
							
						//	echo $val."(${bline})(${mg} ${marge})\t";
							echo $val."(${mg} ${marge})\t";
						}
						else {
							echo "(cell:null)\t";
						}
					}
					echo "<br>";
				}
/*	*/			
				
				
				$area	= $sheet->getPageSetup()->getPrintArea();
				$scale	= $sheet->getPageSetup()->getScale();
				echo '印刷範囲 '.$area.'<br>';
				echo 'スケール '.$scale.'<br>';
				// tinkerによるデバッグ
				eval(\Psy\sh());
				exit();
			}
			else {
				throw new \Exception('error.');
			}
//		}
//		catch (\Exception $e) {
//			Log::error($e->getMessage());
//		}
		// 
		// ビュー
		return view('welcome');
	}
}
