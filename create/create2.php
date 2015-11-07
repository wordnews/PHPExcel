<?php

/**
 * 生成excel文件
 */

define('PHPEXCEL_PATH', dirname(__DIR__));

include PHPEXCEL_PATH . '/PHPExcel/PHPExcel.php';

$PHPExcel = new PHPExcel();

$Sheet = $PHPExcel->getActiveSheet(); // 得到sheet
$Sheet->setTitle('demo'); // 设置sheet名称

// 全部填入（ 数据大了会占用内存，推荐用逐行填充数据 ）
$data = array(
    array('姓名', '分数'),
    array('张三', '99'),
    array('李四', '81'),
);
$Sheet->fromArray($data);

$Write = PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel2007');
$Write->save(__DIR__ . '/create2.xlsx');


