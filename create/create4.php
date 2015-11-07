<?php

/**
 * 生成excel文件
 */

// 输出到浏览器

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

// 设置输出到浏览器
headers('Excel5', 'create4');

$Write = PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel5');
$Write->save('php://output');




function headers($type = 'Excel2007', $filename = 'create4')
{
    // mac下 Excel2007 损坏了，我操
    if ($type == 'Excel2007') {
        // xlsx文件头
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $suffix = '.xlsx';
    } elseif ($type == 'Excel5') {
        // xls头
        header('Content-Type: application/vnd.ms-excel');
        $suffix = '.xls';
    } else {
        exit( '未知的excel类型' );
    }

    header('Content-Disposition: attachment;filename="' . $filename  . $suffix); // excel文件名
    header('Cache-Control: max-age=0'); // 禁止浏览器缓存
}
