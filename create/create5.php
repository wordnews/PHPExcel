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
        array('plus', '80', '三', '三'),
        array('plus', '80', '三', '三'),
    ),
);


define('PHPEXCEL_PATH', dirname(__DIR__));

include PHPEXCEL_PATH . '/PHPExcel/PHPExcel.php';

$PHPExcel = new PHPExcel();

$Sheet = $PHPExcel->getActiveSheet();

// 设置excel文件默认水平垂直居中
$Sheet->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
// 设置文字样式
$Color = new PHPExcel_Style_Color();

$Color->setRGB('FF0000'); // 设置颜色
$Sheet->getStyle('A3:Z3')->getFont()->setName('微软雅黑')->setSize(20)->setBold(true)->setColor($Color);

$Color->setRGB('A2CD5A');
$Sheet->getStyle('A4:Z4')->getFont()->setName('微软雅黑')->setSize(17)->setBold(true)->setColor($Color);

$Color->setRGB('1E90FF');
$Sheet->getStyle('A5:Z14')->getFont()->setName('微软雅黑')->setSize(14)->setColor($Color);

// 设置遇见 \n 就换行
$Sheet->getStyle('A1:Z100')->getAlignment()->setWrapText(true);

$i = 0;
foreach ($data as $value) {
    $j = $i * 4;
    $a = 5;
    // 设置每个单元格的头
    $Sheet->setCellValue(handleAbc($j) . '4', '姓名')->setCellValue(handleAbc($j + 1) . '4', '分数')->setCellValue(handleAbc($j + 2) . '4', '年级')->setCellValue(handleAbc($j + 3) . '4', '班级');
    // 循环给每个单元格填充数据
    foreach ($value as $k => $val) {
        $Sheet->setCellValue(handleAbc($j) . ($a), $val[0])->setCellValue(handleAbc($j + 1) . ($a), $val[1] . '500382199212212719')->setCellValue(handleAbc($j + 2) . ($a),$val[2])->setCellValue(handleAbc($j + 3) . ($a), $val[3]);
        $a++;
    }
    // 年级头
    $Sheet->setCellValue(handleAbc($j) . 3, ($i+1) . "年级\n换行");
    // 合并年级头的单元格
    $Sheet->mergeCells(handleAbc($j) . '3:' . handleAbc($j + 3) . '3');
    // 设置年级头背景颜色
    $Sheet->getStyle(handleAbc($j) . '3:' . handleAbc($j + 3) . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00BFFF');
    // 设置年级头边框
    $Sheet->getStyle(handleAbc($j) . '3:' . handleAbc($j + 3) . '3')->applyFromArray(border('FFB90F'));

    $i++;
}

$Write = PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel5');
$Write->save(__DIR__ . '/create5.xls');


function handleAbc($key)
{
    $array = range('A', 'Z');
    return $array[$key];
}

function border($color = 'FFB90F')
{
    return [
        'borders' => [
            'outline' => [
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => ['rgb' => $color]
            ]
        ]
    ];
}
