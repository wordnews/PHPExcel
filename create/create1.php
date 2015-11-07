<?php

/**
 * 生成excel文件
 */

define('PHPEXCEL_PATH', dirname(__DIR__));

include PHPEXCEL_PATH . '/PHPExcel/PHPExcel.php';

$PHPExcel = new PHPExcel();

$Sheet = $PHPExcel->getActiveSheet(); // 得到sheet
$Sheet->setTitle('demo'); // 设置sheet名称

// 逐行填入数据
$Sheet->setCellValue('A1', '姓名')->setCellValue('B1', '分数');
$Sheet->setCellValue('A2', '张三')->setCellValue('B2', '149');
$Sheet->setCellValue('A3', '李四')->setCellValue('B3', '130');

$Write = PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel2007');
$Write->save(__DIR__ . '/create1.xlsx');


