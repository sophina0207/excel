<?php
// 导入PHPExcel类库  
require_once("excel.php");  
  
// 通常PHPExcel对象有两种实例化的方式  
$data=array(
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
$excel = new excel();
$excel->exportFile($filename, $header, $data);
