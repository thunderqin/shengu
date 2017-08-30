<?php
namespace Home\Controller;
use Think\Controller;
Vendor('PHPExcel.PHPExcel');
Vendor('PHPExcel.PHPExcel.Writer.Excel2007');

class IndexController extends Controller {
    public function index(){
    	$data = [];
    	array_push($data, $this->readExcel(0));
    	array_push($data, $this->readExcel(1));
    	array_push($data, $this->readExcel(2));
    	echo json_encode($data);
    }

    public function readExcel($sheet){
    	$PHPExcel = new \PHPExcel();
		/**默认用excel2007读取excel，若格式不对，则用之前的版本进行读取*/
		$PHPReader = new \PHPExcel_Reader_Excel2007();
    	if(!$sheet){
    		$sheet = 0;
    	}

    	/**对excel里的日期进行格式转化*/
		

		$filePath = './Public/Uploads/test.xlsx';

		if(!$PHPReader->canRead($filePath)){
		    $PHPReader = new PHPExcel_Reader_Excel5();
		    if(!$PHPReader->canRead($filePath)){
		        echo 'no Excel';
		        return ;
		    }
		}

		$PHPExcel = $PHPReader->load($filePath);
		/**读取excel文件中的第一个工作表*/
		$currentSheet = $PHPExcel->getSheet($sheet);
		/**取得最大的列号*/
		$allColumn = $currentSheet->getHighestColumn();
		/**取得一共有多少行*/
		$allRow = $currentSheet->getHighestRow();
		/**从第二行开始输出，因为excel表中第一行为列名*/
		$data = [];
		for($currentRow = 2;$currentRow <= $allRow;$currentRow++){
		/**从第A列开始输出*/
			$item = [];
			for($currentColumn= 'A';$currentColumn<= $allColumn; $currentColumn++){
		    	$val = $currentSheet->getCellByColumnAndRow(ord($currentColumn) - 65,$currentRow)->getValue();/**ord()将字符转为十进制数*/
		    	if(empty($val)){
		    		break;
		    	}
		    	array_push($item, $val);
		    	/*if($currentColumn == 'A') {
		        	//echo date("Y-m-d H:i:s", PHPExcel_Shared_Date::ExcelToPHP($val));

				}*/
			}
			if(empty($item[0])){
		    	break;
		    }
			array_push($data, $item);
		}
		return $data;
    }
}