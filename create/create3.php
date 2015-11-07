<?php

/**
 * 生成excel文件
 */

// 按照年级分组创建excel ( 多个sheel )

/**
 * 待演示的数据
 */
$data = array(
    //        姓名    分数   年级   几班
    array(
        array('张三', '80', '一', '三'),
        array('李四', '80', '一', '二'),
        array('IOS', '80', '一', '二'),
        array('张龙', '80', '一', '二'),
        array('赵虎', '80', '一', '一'),
    ),
    array(
        array('元芳', '80', '二', '三'),
        array('诺基亚', '80', '二', '三'),
        array('小红', '80', '二', '三'),
        array('红红', '80', '二', '一'),
    ),
    array(
        array('王麻子', '80', '三', '一'),
        array('逆风', '80', '三', '一'),
        array('iphone', '80', '三', '二'),
        array('plus', '80', '三', '三'),
    ),
);


define('PHPEXCEL_PATH', dirname(__DIR__));

include PHPEXCEL_PATH . '/PHPExcel/PHPExcel.php';

$PHPExcel = new PHPExcel();

for ($i = 1; $i <= 3; $i++) {
    if ($i > 1) { // 默认有一个，所以这里少建一个
        $PHPExcel->createSheet(); // 创建新的内置表
    }
    $PHPExcel->setActiveSheetIndex($i - 1); // 把新创建的sheet设定为当前活动sheet
    $Sheet = $PHPExcel->getActiveSheet(); // 获取当前sheet
    $Sheet->setTitle($i . '年级'); // 给sheet设置标题
    $Sheet->setCellValue('A1', '姓名')->setCellValue('B1', '分数')->setCellValue('C1', '年级')->setCellValue('D1', '班级'); // 标题

    $j = 2;
    foreach ($data[$i - 1] as $val) {
        $Sheet->setCellValue('A' . $j, $val[0])->setCellValue('B' . $j, $val[1])->setCellValue('C' . $j, $val[2])->setCellValue('D' . $j, $val[3]); // 循环填充内容
        $j++;
    }

}

$Write = PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel5');
$Write->save(__DIR__ . '/create3.xls');




