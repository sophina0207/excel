<?php
/**
* @user : wangying
* @date : 2017年6月30日
* @desc : excel类
**/
class excel{
	protected $sheet;//当前工作sheet
	protected $phpexcel;
	public function __construct(){
		require_once 'PHPExcel.php';
	}
	public function exportFile($filename,$header,$data){
		$phpexcel = new PHPExcel();
		//设置当前的工作sheet
		$sheet = $phpexcel->getActiveSheet();
		$this->phpexcel = $phpexcel;
		$this->sheet = $sheet;
		$sheet->setTitle($filename);
		//生成header,colum表示列，从第0列开始，行从1开始
		$column = 0;
		$line = 1;
		foreach ($header as $columData)
		{
			//根据列，行的数字开始创建header。
			$return=$sheet->setCellValueByColumnAndRow($column,$line,$columData,true);
			//设置字体大小
			$this->setFontSize($column, $line, 16);
			//字体加粗
			$this->setFontBold( $column, $line);
			//设置自动宽度
			$this->setAutoWidth($column);
			//设置横向对齐方式
			$this->setAlign($column, $line,2);
			//设置边框
			$this->setBorderStyle($column, $line);
			$this->setBorderColor($column, $line, 'FF993300');
			$column++;
		}
		foreach ($data as $key =>$item)
		{
			$column=0;
			$line++;
			foreach ($item as $columData)
			{
				//添加每行记录
				$sheet->setCellValueByColumnAndRow($column,$line,$columData);
				$this->setAlign($column, $line,2);
				//设置边框
				$this->setBorderStyle($column, $line);
				$this->setBorderColor($column, $line, 'FF993300');
				$column++;
			}
		}
		//设置单元格边框
		$filename .= '_'.date('YmdHis',time());
		//导出文件
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
		header('Cache-Control: max-age=0');
		PHPExcel_IOFactory::createWriter($phpexcel,'excel5')->save('php://output');
	}
	/**
	 * 设置字体大小
	 * @param unknown $column
	 * @param unknown $line
	 * @param unknown $size
	 * @param unknown $sheet
	 */
	public function setFontSize($column,$line,$size,$sheet=''){
		if(empty($sheet)){
			$sheet = $this->sheet;
		}
		$sheet->getStyleByColumnAndRow($column,$line)->getFont()->setSize($size);
	}
	/**
	 * 设置字体加粗
	 * @param unknown $column
	 * @param string $boolean
	 * @param unknown $sheet
	 */
	public function setFontBold($column,$line,$boolean=true,$sheet=''){
		if(empty($sheet)){
			$sheet = $this->sheet;
		}
		$sheet->getStyleByColumnAndRow($column,$line)->getFont()->setBold($boolean);
	}
	/**
	 * 设置列的自动宽度
	 * @param unknown $column
	 * @param string $boolean
	 * @param unknown $sheet
	 */
	public function setAutoWidth($column,$boolean=true,$sheet=''){
		if(empty($sheet)){
			$sheet = $this->sheet;
		}
		$sheet->getColumnDimensionByColumn($column)->setAutoSize($boolean);
	}
	/**
	 * 设置列的宽度
	 * @param unknown $column
	 * @param unknown $value
	 * @param string $sheet
	 */
	public function setWidthValue($column,$value,$sheet=''){
		if(empty($sheet)){
			$sheet = $this->sheet;
		}
		$sheet->getColumnDimensionByColumn($column)->setWidth($value);		
	}
	/**
	 * 设置横向对齐方式
	 * @param unknown $colum
	 * @param unknown $line
	 * @param number $type 1-左对齐，2-居中对齐，3-右对齐，其余方式暂无，如需要可扩充
	 * @param unknown $sheet
	 */
	public function setAlign($column,$line,$type=1,$sheet=''){
		if(empty($sheet)){
			$sheet = $this->sheet;
		}
		switch ($type) {
			case 2://居中对齐
				$type = PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
				break;
			case 3://右对齐
				$type = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
			default://左对齐
				$type = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
			break;
		}
		$sheet->getStyleByColumnAndRow($column,$line)->getAlignment()->setHorizontal($type);
	}
	/**
	 * 设置单元格垂直方向对齐方式
	 * @param unknown $column
	 * @param unknown $line
	 * @param string $type
	 * @param string $sheet
	 */
	public function setVertical($column,$line,$type='',$sheet=''){
		if(empty($sheet)){
			$sheet = $this->sheet;
		}
		$type = strtolower($type);
		switch ($type) {
			case 'top':
				$type = PHPExcel_Style_Alignment::VERTICAL_TOP;break;
			case 'bottom':
				$type = PHPExcel_Style_Alignment::VERTICAL_BOTTOM;break;
			default:
				$type = PHPExcel_Style_Alignment::VERTICAL_CENTER;break;
		}
		$sheet->getStyleByColumnAndRow($column,$line)->getAlignment()->setVertical($type);
		
	}
	/**
	 * 设置边框的样式
	 * @param unknown $column
	 * @param unknown $line
	 * @param unknown $type
	 * @param unknown $border
	 * @param string $sheet
	 */
	public function setBorderStyle($column,$line,$type='',$border='',$sheet=''){
		if(empty($sheet)){
			$sheet = $this->sheet;
		}
		$border = $this->setBorderDirection($border);
		switch ($type) {
			case 'double':
				$type = PHPExcel_Style_Border::BORDER_DOUBLE;break;
			case 'medium':
				$type = PHPExcel_Style_Border::BORDER_MEDIUM;break;
			default:
				$type = PHPExcel_Style_Border::BORDER_THIN;break;
			break;
		}
		$sheet->getStyleByColumnAndRow($column,$line)->getBorders()->$border()->setBorderStyle($type);
		
	}
	/**
	 * 设置单元格的边框颜色
	 * @param unknown $column
	 * @param unknown $line
	 * @param unknown $color
	 * @param string $border
	 * @param string $sheet
	 */
	public function setBorderColor($column,$line,$color,$border='',$sheet=''){
		if(empty($sheet)){
			$sheet = $this->sheet;
		}
		$border = $this->setBorderDirection($border);
		$sheet->getStyleByColumnAndRow($column,$line)->getBorders()->$border()->getColor()->setARGB($color);
	}
	/**
	 * 设置边框的方向
	 * @param unknown $border top\rigth\bottom\left 默认四边
	 * @return string
	 */
	protected function setBorderDirection($border=''){
		$border = strtolower($border);
		switch ($border) {
			case 'top'://上边框
				$border = 'getTop';break;
			case 'right':
				$border = 'getRight';break;
			case 'getBottom':
				$border = 'getBottom';break;
			case 'left':
				$border = 'getLeft';break;
			default:
				$border = 'getAllBorders';break;
		}
		return $border;
	}
}