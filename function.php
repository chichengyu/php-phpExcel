<?php 

/**
 * 导出Excel
 * @param  [Array]   $cellFormat    [表格头] 一维数组 
 * @param  [Array]   $expTableData  [数据集] 二维数组
 * @param  [string]  $dir           [保存路径]
 * @param  [string]  $fileName      [文件名]
 * @return [Array]   [description]
 */
function exportExcel($cellFormat, $expTableData, $dir, $fileName){
    // 格式化cell标题
    foreach($cellFormat as $k => $v){
        $expCellName[] = array($k, $v);
    }

    $fileDir = $_SERVER['DOCUMENT_ROOT'].$dir;
    $filePathName = $fileName.date('_YmdHis') . rand(1000,9999);
    $fileName = iconv('utf-8', 'gb2312', $filePathName);//文件名称
    if (!is_dir($fileDir)) {
        mkdir($fileDir, 0755, true);
    }

    $cellNum = count($expCellName);
    $dataNum = count($expTableData);
    Vendor('PHPExcel');
    //require_once "./ThinkPHP/Library/Vendor/PHPExcel.php";
    $objPHPExcel = new \PHPExcel();
    $cellName = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ');
    $objPHPExcel->getActiveSheet(0)->mergeCells('A1:'.$cellName[$cellNum-1].'1');//合并单元格
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A1', '导出时间:'.date('Y-m-d H:i:s'));  
    for($i=0;$i<$cellNum;$i++){
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cellName[$i].'2', $expCellName[$i][1]); 
        $objPHPExcel->getActiveSheet(0)->getColumnDimension($cellName[$i])->setAutoSize(true);
    }   
    for($i=0;$i<$dataNum;$i++){
        for($j=0;$j<$cellNum;$j++){
            $objPHPExcel->getActiveSheet(0)->setCellValueExplicit($cellName[$j].($i+3), $expTableData[$i][$expCellName[$j][0]]);
        }             
    }  
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');  
    header('Content-Disposition: attachment;filename='.$fileName.'.xls');  
    header('Cache-Control: max-age=0');  
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');//生成excel文件
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    
    $objWriter->save($fileDir.$fileName.'.xlsx'); // 文件保存

    $result = array(  
        'errcode' => 0,  
        'errmsg' => '导出成功',
        'file_path' => 'http://'.$_SERVER['HTTP_HOST'].$dir.$filePathName.'.xlsx',
    );
    return $result;
}