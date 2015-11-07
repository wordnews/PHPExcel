<?php
header('Content-Type:text/html; charset=utf-8');
echo '<pre>';
/**
 * 读取Excel的所有内容，文件大了耗内存
 */

define('PHPEXCEL_PATH', dirname(__DIR__));

// 导入读取Excel的类文件
include PHPEXCEL_PATH . '/PHPExcel/PHPExcel/IOFactory.php';

$PHPExcel = PHPExcel_IOFactory::load('../create/create3.xls'); // 加载Excel文件

$sheetCount = $PHPExcel->getSheetCount(); // 获取sheet个数

for ($i=0; $i<$sheetCount; $i++) {
    $data = $PHPExcel->getSheet($i)->toArray(); // 读取每个sheet里的数据
    var_dump($data);
}


