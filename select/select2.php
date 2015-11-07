<?php
header('Content-Type:text/html; charset=utf-8');
echo '<pre>';
/**
 * 读取Excel的所有内容，逐行逐列读取
 */

define('PHPEXCEL_PATH', dirname(__DIR__));

// 导入读取Excel的类文件
include PHPEXCEL_PATH . '/PHPExcel/PHPExcel/IOFactory.php';

$PHPExcel = PHPExcel_IOFactory::load('../create/create3.xls'); // 加载Excel文件

foreach ($PHPExcel->getWorksheetIterator() as $sheet) { // 循环获取sheet
    foreach ($sheet->getRowIterator() as $row) { // 逐行处理
        // $row->getRowIndex() 获取当前行
        if ($row->getRowIndex() < 2) { // 不要标题
            continue;
        }
        foreach ($row->getCellIterator() as $cell) { // 逐列读取
            echo $cell->getValue(); // 获取单元格数据
        }
        echo '<br/>';
    }
    echo '<br/>';
}



