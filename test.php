<?php
// 导入PHPExcel类库  
require_once("PHPExcel.php");  
  
// 通常PHPExcel对象有两种实例化的方式  
$arr=array(
		array("排名"=>1,'name'=>'曾川玲1','vote_count'=>22),
		array("排名"=>2,'name'=>'曾川玲2','vote_count'=>33),
		array("排名"=>3,'name'=>'曾川玲3','vote_count'=>53),
		array("排名"=>4,'name'=>'曾川玲4','vote_count'=>123),
		array("排名"=>5,'name'=>'曾川玲5','vote_count'=>553),
		array("排名"=>6,'name'=>'曾川玲6','vote_count'=>44),
		array("排名"=>7,'name'=>'曾川玲7','vote_count'=>55),
		array("排名"=>8,'name'=>'曾川玲8','vote_count'=>32),
		array("排名"=>9,'name'=>'曾川玲9','vote_count'=>65),
		array("排名"=>10,'name'=>'曾川玲0','vote_count'=>753),
		array("排名"=>11,'name'=>'曾川玲11','vote_count'=>5323),
		array("排名"=>12,'name'=>'曾川玲12','vote_count'=>663),
		array("排名"=>13,'name'=>'曾川玲13','vote_count'=>53),
		array("排名"=>14,'name'=>'曾川玲14','vote_count'=>53),
		array("排名"=>15,'name'=>'曾川玲15','vote_count'=>543),
		array("排名"=>16,'name'=>'曾川玲16','vote_count'=>53),
		array("排名"=>17,'name'=>'曾川玲17','vote_count'=>736),
		array("排名"=>18,'name'=>'曾川玲18','vote_count'=>33),
		array("排名"=>19,'name'=>'曾川玲19','vote_count'=>66),
		array("排名"=>20,'name'=>'曾川玲20','vote_count'=>345),
		array("排名"=>21,'name'=>'曾川玲21','vote_count'=>345),
		array("排名"=>22,'name'=>'曾川玲22','vote_count'=>345),
		array("排名"=>23,'name'=>'曾川玲23','vote_count'=>345),
		array("排名"=>24,'name'=>'曾川玲24','vote_count'=>345),
		array("排名"=>25,'name'=>'曾川玲25','vote_count'=>345),
		array("排名"=>26,'name'=>'曾川玲26','vote_count'=>345),
		array("排名"=>27,'name'=>'曾川玲27','vote_count'=>345),
		array("排名"=>28,'name'=>'曾川玲28','vote_count'=>345),
		array("排名"=>29,'name'=>'曾川玲29','vote_count'=>345),
		array("排名"=>30,'name'=>'曾川玲30','vote_count'=>345),
		array("排名"=>31,'name'=>'曾川玲31','vote_count'=>345),
		array("排名"=>32,'name'=>'曾川玲32','vote_count'=>345),
);
$header=array('序号','名称','投票数');
$filename='test';
$phpexcel=new PHPExcel();
$sheet=$phpexcel->getActiveSheet();
$sheet->setTitle($filename);
$colum=0;
foreach ($header as $key =>$k)
{
	$return=$sheet->setCellValueByColumnAndRow($colum,1,$k,true);
	$colum++;
}
foreach ($arr as $key =>$item)
{
	$colum=0;
	foreach ($item as $k =>$v )
	{
		//添加每行记录
		$sheet->setCellValueByColumnAndRow($colum,$key+2,$v);
		$colum++;
	}
}
//设置单元格边框
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
header('Cache-Control: max-age=0');
PHPExcel_IOFactory::createWriter($phpexcel,'excel5')->save('php://output');


















// new exportExcel('export', $array);
class exportExcel{
	protected $sheet;
	public function __construct($file_name,$arr)
	{
		$phpexcel = new PHPExcel();
		$sheet=$phpexcel->getActiveSheet();
		$this->sheet=$sheet;
		$sheet->setTitle($file_name);
		if(empty($arr))
		{
			return ;
		}
		foreach ($arr as $key =>$item)
		{
			if($key == 0)
			{
				//设置head的行高
				$sheet->getRowDimension($key+1)->setRowHeight(30);
			}
			$colum=0;
			foreach ($item as $k =>$v )
			{
				if($key == 0)
				{
					//添加head记录
					$sheet->setCellValueByColumnAndRow($colum,$key+1,$k);
					$this->setStyle($colum, $key+1,true);
				}
				//添加每行记录
				$sheet->setCellValueByColumnAndRow($colum,$key+2,$v);
				$this->setStyle($colum, $key+2);
				$colum++;
			}
		}
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="'.$file_name.'.xls"');
		header('Cache-Control: max-age=0');
		PHPExcel_IOFactory::createWriter($phpexcel, 'Excel5')->save('php://output');
		
	}
 
	protected function setStyle($colum,$row,$head=false)
	{
		$sheet=$this->sheet;
		$style=$sheet->getStyleByColumnAndRow($colum,$row);
		if($head)
		{
			//设置字体
			$style->getFont()->setBold(true)
							 ->setSize(14);
		}else
		{
			//设置字体
			$style->getFont()->setSize(12);
		}
		//设置边框
		$objBorder=$style->getBorders();
		$objBorder->getAllBorders()->getColor()->setRGB('000000');
		$objBorder->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		//设置对齐方式--上下居中
		$style->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	}
}