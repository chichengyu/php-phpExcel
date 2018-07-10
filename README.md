# php-phpExcel

使用phpExcel包导出Excel文件，phpExcel官网下载地址：https://github.com/PHPOffice/PHPExcel

下载之后,将Classes文件夹全部复制

function.php文件为封装的方法

条用方式如下：

$xlsCell = array(
	'name' 	=> '姓名',
	'sex'	=> '性别',
	'age'	=> '年龄'
);

$excelDate = array();

$excelDate[] = array(
	'name' 	=> '小二',
	'sex'	=> '男',
	'age'	=> '18'
);

$path = $_SERVER['DOCUMENT_ROOT'].'/Public/';

$res = exportExcel($xlsCell,$excelDate,$path,'二维码');