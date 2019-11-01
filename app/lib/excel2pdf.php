<?php
/*-----------------------------------------------------------------------------*
	
 *-----------------------------------------------------------------------------*/
namespace App\Lib;

//=======================================
/**
*/
class	ExcelCell
{
	public	$posx;
	public	$posy;
	public	$width;
	public	$height;
	
	public	$strVal;
	
	public	$Font;				//フォント名
	public	$FontSize;			//フォントサイズ(mm)
	public	$bold;				//Bool	
	public	$italic;			//Bool	
	public	$superscript;		//Bool	上付き
	public	$subscript;			//Bool	添え字
	public	$underline;			//Str	下線
 	public	$strikethrough;		//Bool	取消線
	public	$color;				//Str？？
	
	public	$HAlignment;		//str	
	public	$VAlignment;		//str	
	public	$wrapText;			//Bool	折り返し
	public	$shrinkToFit;		//Bool	縮小
	public	$indent;			//Int インデント
	//罫線　[0]str:borderStyle  [1]str:color
	public	$bdrtop;			//
	public	$bdrbottom;
	public	$bdrleft;
	public	$bdrright;
}
//=======================================
/**
*/
class	excel2pdf
{
	public	$errMessage		= '';
	public	static	$POINT	= 0.3528;			//1ポイントの長さ(mm)
	//PDF
	private	static $fontTbl	= array();			//使用するフォント達
	private	$pdfFileName	= 'sheet.pdf';		//出力PDFファイル名
	private	$tcpdf			= null;
	
	//Excel
	private	$excelFileName	= '';				//エクセルファイル名
	private	$book			= null;				//ブックオブジェクト
	private	$sheet			= null;				//シートオブジェクト
	
	//PDF
		//単位（mm）
		//用紙
		//方向
	//Excel
	private	$area	= '';			//印刷範囲
	private	$spclm	= 0;
	private	$sprow	= 0;
	private	$epclm	= 0;
	private	$eprow	= 0;
	private	$csize	= array();		//カラムサイズ
	private	$rsize	= array();		//行サイズ
	private	$cpos	= array();		//カラム位置（X）
	private	$rpos	= array();		//行位置（Y）
	private	$recio	= 1.0;			//印刷倍率
	private	$margentop;				//ページ余白
	private	$margenleft;			//ページ余白
	
	private	$cells	= array();		//印刷範囲のセル情報
	public	$clmWidthPt= array();	//カラム幅の初期値（エクセルファイルから取得できない時用）

	//-------------------------------------------------
	/**
		PDF出力で使用するフォント群を登録する
		@param	$ftbl	フォントファイルの配列（順番が番号となる）
		@return	TRUE/FLSE
	*/
	public static function setPdfFonts($ftbl)
	{
		//@@@@@@@@@@@@@@
	}
	//-------------------------------------------------
	/**
		エクセルのフォントに対応するPDF用フォントを指定する。
		@param	$font	エクセルのフォント名
		@param	$id		対応するPDFフォントの番号（0～）
		@return	なし
	*/
	public static function setUseFontNo($font,$id)
	{
		//@@@@@@@@@@@@@@
	}
	//-------------------------------------------------
	/**
		出力するPDFファイル名を設定する
		@param	$fn	ファイル名
	*/
	public function setPdfFilename($fn)
	{
		$this->pdfFileName	= $fn;
	}
	//-------------------------------------------------
	/**
		出力するPDFファイル名を取得する
		@return	ファイル名
	*/
	public function getPdfFilename()
	{
		return	$this->pdfFileName;
	}
	
	//-------------------------------------------------
	/**
		PDFファイル出力する
		@return	なし
	*/
	public function writePDF()
	{
		$this->tcpdf = new \TCPDF("P", "mm", "A4", true, "UTF-8" );
		// 出力するPDFの初期設定
	//	$this->tcpdf = new TcpdfFpdi('L', 'mm', 'A4');
		$this->tcpdf->setPrintHeader( false );    
		$this->tcpdf->setPrintFooter( false );
		$this->tcpdf->AddPage();
		$this->tcpdf->SetFont('kozminproregular','',14);
		
		for($r=$this->sprow;$r<=$this->eprow; $r++) {
			for($c=$this->spclm; $c<=$this->epclm; $c++) {
				//セルの取得
				if(!empty($this->cells[$r][$c])) {
					$cell	= $this->cells[$r][$c];
					if(!empty($cell)) {
						$this->writeCell($cell);
					}
				}
			}
		}
		
		$this->tcpdf->Output($this->pdfFileName, "I");
		return;
	}
	//-------------------------------------------------
	/**
		PDFにセルを出力する
		@param	$cell	セル情報
		@return	なし
	*/
	public function writeCell($cell)
	{
		$border		= '';			//枠線（後で作り変える？線種対応）
		$align		= 'L';			//横方向の位置
		$fill		= false;
		$stretch	= 0;			//テキストの伸縮モード
		$valign		= 'T';			//縦方向の位置
		if($cell->bdrtop[0]!='none')	$border .= 'T';
		if($cell->bdrbottom[0]!='none')	$border .= 'B';
		if($cell->bdrleft[0]!='none')	$border .= 'L';
		if($cell->bdrright[0]!='none')	$border .= 'R';
		
		if($cell->HAlignment=='right')	$align = 'R';
		if($cell->HAlignment=='center')	$align = 'C';

		if($cell->VAlignment=='center')	$align = 'M';
		if($cell->VAlignment=='bottom')	$align = 'B';
		
		if((!empty($border))||(!empty($cell->strVal))) {
			//ボックスの左上
			$this->tcpdf->SetXY( $cell->posx, $cell->posy, true);
			$this->tcpdf->SetFont('kozminproregular','',$cell->FontSize);
			//ボックス（セル）の描画
			$this->tcpdf->Cell( $cell->width, $cell->height, $cell->strVal,		//サイズ、文字列
								$border, 0, $align, $fill, '', $stretch, true, 'T', $valign );
		}
	}
	//-------------------------------------------------
	/**
		セルの確認（デバッグ用）
		@param	$c,$r	セル位置
		@return	なし
	*/
	public function debugCell($c,$r)
	{
		if($c==0) {
			echo "<br>カラムサイズ:<br>";
			var_dump($this->csize);
		}
		if($r==0) {
			echo "<br>行高さ:<br>";
			var_dump($this->rsize);
		}
		if(($r>0)&&($c>0)) {
			$cell	= $this->cells[$r][$c];
			if(!empty($cell)) {
				echo "<br>行：${r} カラム:${c}";
				echo "<br>座標　X:{$cell->posx}  Y:{$cell->posy}";
				echo "<br>サイズX:{$cell->width}  Y:{$cell->height}";
				echo "<br>データ　:{$cell->strVal}";
				echo "<br>フォント:{$cell->Font}　　size:{$cell->FontSize}";
				echo "<br>";
			}
			else {
				echo "<br>行：${r} カラム:${c} セル無し<br>";
			}
		}
	}
	//-------------------------------------------------
	/**
		カラム幅をピクセル値の配列で設定する
		@param	$wa	カラム幅(px)の配列（1～）
	*/
	public function setColumnWidthsPx($wa)
	{
		$this->clmWidthPt[0]	= \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToPoints(72);	//デフォルトの幅
		foreach($wa as $i => $v){
			$v	= \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToPoints($v);
			$this->clmWidthPt[$i]	= $v;
		}
	}
	//-------------------------------------------------
	/**
		エクセルファイル名を設定する
		@param	$fn	ファイル名
	*/
	public function setExcelFilename($fn)
	{
		//ファイルがあるか？
		if(file_exists($fn)) {
			$this->excelFileName	= $fn;
			//エクセルファイルを読込む
			$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
			$book	= $reader->load($fn);
			$this->setBook($book);
		}
		else {
			$errMessage	= 'File not Exist!! filename='.$fn;
		}
	}
	//-------------------------------------------------
	/**
		エクセルファイル名を取得する
		@return	ファイル名
	*/
	public function getExcelFilename()
	{
		return	$this->excelFileName;
	}
	//-------------------------------------------------
	/**
		エクセルブックを設定する
		@param	$book	ファイル名
	*/
	public function setBook($book)
	{
		$this->book	= $book;
		$sheet	= $book->getActiveSheet();
		//シートを設定
		$this->setSheet($sheet);
	}
	//-------------------------------------------------
	/**
		エクセルブックを取得する
		@return	エクセルブック
	*/
	public function getBook()
	{
		return	$this->book;
	}
	//-------------------------------------------------
	/**
		エクセルシートを設定する
		@param	$sheet	エクセルシート
	*/
	public function setSheet($sheet)
	{
		//シートを保持
		$this->sheet	= $sheet;
		//印刷エリアの取得
		$this->area		= $sheet->getPageSetup()->getPrintArea();
		//セル群を読込む
		$this->loadCells();
	}
	//-------------------------------------------------
	/**
		エクセルシートを取得する
		@return	エクセルシート
	*/
	public function getSheet()
	{
		return	$this->sheet;
	}
	//-------------------------------------------------
	/**
		エクセルシートからセルの情報を読取る
	*/
	public function loadCells()
	{
		//印刷範囲を内部に設定
		$r	= $this->area2index($this->area);
		$s	= $r['sp'];
		$e	= $r['ep'];
		$this->sprow	= $s[0];	//開始行
		$this->spclm	= $s[1];	//終了行
		$this->eprow	= $e[0];	//開始カラム
		$this->epclm	= $e[1];	//終了カラム
		//印刷倍率
		$this->scale	= $this->sheet->getPageSetup()->getScale();
		$this->recio	= $this->scale / 100.0;
		//マージン（単位は？）
		$this->margentop	= $this->sheet->getPageMargins()->getTop();
		$this->margenleft	= $this->sheet->getPageMargins()->getLeft();
		
		//カラムサイズ
		$this->sheet->calculateColumnWidths();			/* ???  */
		$def	= $this->sheet->getDefaultColumnDimension();
		$defw	= $def->getWidth();
		if($defw<=0) $defw = $this->clmWidthPt[0];
		$w		= 0.0;
		$w		= $this->margenleft;		//余白
		
	//	for($i=$this->spclm; $i<=$this->epclm; $i++) {
	//		if( isset($dims[$i]) ) {
	//			$dims[$i]->setAutoSize(false);
	//		}
	//	}
	//	$this->sheet->calculateColumnWidths();			/* ???  */
		
		$dims	= $this->sheet->getColumnDimensions();
		for($i=$this->spclm; $i<=$this->epclm; $i++) {
			$vv	= -1;
			if( isset($dims[$i]) ) {
				$vv	= $dims[$i]->getWidth();
			}
			if( $vv<=0 ) {
				if( isset($this->clmWidthPt[$i])) {
					$vv	= $this->clmWidthPt[$i];
				}
				else {
					$vv	= $defw;
				}
			}
			$v	= $vv * excel2pdf::$POINT;		//mm へ変換
			$this->csize[$i]	= $v;
			$this->cpos[$i]		= $w;
			$w	+= $v;
		}
		//行サイズ
		$def	= $this->sheet->getDefaultRowDimension();
		$defw	= $def->getRowHeight();
		if($defw<=0) $defw = 18.75;
		$dims	= $this->sheet->getRowDimensions();
		$w		= 0.0;
		$w		= $this->margentop;
		for($i=$this->sprow; $i<=$this->eprow; $i++) {
			if( isset($dims[$i]) ) {
				$v	= $dims[$i]->getRowHeight();
			}
			else {
				$v	= $defw;
			}
			$v	= $v * excel2pdf::$POINT;		//mm へ変換
			$this->rsize[$i]	= $v;
			$this->rpos[$i]		= $w;
			$w	+= $v;
		}
		//印刷範囲のセルの情報を取得、nullは結合セルなど
	//	$this->cells	= array();				//印刷範囲のセル情報
		for($r=$this->sprow;$r<=$this->eprow; $r++) {
			for($c=$this->spclm; $c<=$this->epclm; $c++) {
				//セルの取得
				$cell	= $this->sheet->getCellByColumnAndRow($c,$r,false);
				if($cell!=null) {
					$mg		= $cell->isMergeRangeValueCell();	//代表セル
					$marge	= $cell->getMergeRange();			//結合範囲　　結合してなければ空
					if(($mg==1) || empty($marge)) {				//生きているセル
						$ec	= $this->getExcelCell($r,$c,$cell,$marge);
						
						$this->cells[$r][$c]	= $ec;
					}
					else {
						$this->cells[$r][$c]	= null;
					}
				}
				else {
					$this->cells[$r][$c]	= null;
				}
			}
		}
	}
	
	//-------------------------------------------------
	/**
		セルの情報を取得する
		@param	$r		行番号
		@param	$c		カラム番号
		@param	$cell	PhpSpreadsheetのセルオブジェクト
		@param	$㎎		結合セルの範囲
		@return	ExcelCellクラスのインスタンス
	*/
	public function getExcelCell($r,$c,$cell,$mg)
	{
		$ec	= new ExcelCell();
		
		$val	= $cell->getFormattedValue();		//表示文字列
		$ec->strVal	= $val;
		//左上座標
		$ec->posx	= $this->cpos[$c];
		$ec->posy	= $this->rpos[$r];
		if(empty($mg)) {
			$ec->width	= $this->csize[$c];
			$ec->height	= $this->rsize[$r];
		}
		else {
			$ar	= excel2pdf::area2index($mg);		//結合範囲の合計
			$w	= 0.0;
			for($i=$ar['sp'][1]; $i<=$ar['ep'][1] ;$i++ ) {
				$w	+= $this->csize[$i];
			}
			$ec->width	= $w;
			
			$w	= 0.0;
			for($i=$ar['sp'][0]; $i<=$ar['ep'][0] ;$i++ ) {
				$w	+= $this->rsize[$i];
			}
			$ec->height	= $w;
		}
		
		$style	= $cell->getStyle();
		//フォント情報
		$f	= $style->getFont();
		$ec->Font			= $f->getName();
		$ec->FontSize		= $f->getSize() * excel2pdf::$POINT;
		$ec->bold			= $f->getBold();
		$ec->italic			= $f->getItalic();
		$ec->superscript	= $f->getSuperscript();
		$ec->subscript		= $f->getSubscript();
		$ec->underline		= $f->getUnderline();
		$ec->strikethrough	= $f->getStrikethrough();
		$ec->color			= $f->getColor()->getRGB();
		
		//アライメント等
		$a	= $style->getAlignment();
		$ec->HAlignment		= $a->getHorizontal();
		$ec->VAlignment		= $a->getVertical();
		$ec->wrapText		= $a->getWrapText();
		$ec->shrinkToFit	= $a->getShrinkToFit();
		$ec->indent			= $a->getIndent();
		
		//罫線
		$b	= $style->getBorders();
		$k	= $b->getTop();					//上
		$ec->bdrtop[0]		= $k->getBorderStyle();
		$ec->bdrtop[1]		= $k->getColor();
		$k	= $b->getBottom();			//下
		$ec->bdrbottom[0]	= $k->getBorderStyle();
		$ec->bdrbottom[1]	= $k->getColor();
		$k	= $b->getLeft();			//左
		$ec->bdrleft[0]		= $k->getBorderStyle();
		$ec->bdrleft[1]		= $k->getColor();
		$k	= $b->getRight();			//右
		$ec->bdrright[0]	= $k->getBorderStyle();
		$ec->bdrright[1]	= $k->getColor();
		
		return	$ec;
	}
	
	//-------------------------------------------------
	/**
		エリアを表す文字列からカラム・ロウの
		インデックス値(1～)へ変換する。
	
		@param	$area	セル範囲を表す文字列 ex)'B1:G12'
		@return	セル範囲を表す連想配列 $ret= ['sp'=>(行、列)、'ep'=>(行、列)]
	*/
	public static function area2index($area)
	{
		$p	= explode(':',$area);
		$sp	= excel2pdf::name2index($p[0]);
		$ep	= excel2pdf::name2index($p[1]);
		$ret	= array('sp'=>$sp,'ep'=>$ep);
		return $ret;
	}

	//-------------------------------------------------
	/**
		セル位置を表す文字列から行・列番号へ変換する
	
		@param	$cn	セル位置の文字列 ex)'C4'
		@return	セル位置のインデックスを表す配列 $ret=（0:行、1:列）
	*/
	public static function name2index($cn)
	{
		$cidx = $ridx = 0;
		for($i=0;$i< strlen($cn); $i++) {
			if(ctype_digit($cn[$i])==TRUE) {
				$clm	= substr( $cn, 0, $i );
				$row	= substr( $cn, $i );
				$cidx	= \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($clm);
				$ridx	= (int)$row;
				break;
			}
		}
		return array($ridx,$cidx);
	}
}
//----------------------------------- eof -----------------------------------
