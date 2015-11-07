<?php
header('Content-Type:text/html; charset=utf-8');
/**
 * 读取Excel的内容 (部分读取)
 */

define('PHPEXCEL_PATH', dirname(__DIR__));

// 导入读取Excel的类文件
include PHPEXCEL_PATH . '/PHPExcel/PHPExcel/IOFactory.php';

$filename = '../create/create3.xls';

$fileType = PHPExcel_IOFactory::identify($filename); // 获取文件类型

$Reader = PHPExcel_IOFactory::createReader($fileType); // 读取操作对象

$sheetName = ['1年级']; // 指定需要读取的sheet

$Reader->setLoadSheetsOnly($sheetName); // 加载指定的sheet
$PHPExcel = $Reader->load($filename); // 加载文件


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



