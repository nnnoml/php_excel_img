<?php
require 'vendor/autoload.php';
require 'drawClass.php';
class loadExcel{
	public $data = [];

	function load(){

		    $reader = \PHPExcel_IOFactory::createReader('Excel2007'); // 读取 excel 文档


		    $PHPExcel = $reader->load("readexcel.xlsx"); // 文档名称
		    $objWorksheet = $PHPExcel->getActiveSheet();

		    $highestRow = $objWorksheet->getHighestRow(); // 取得总行数
		    $highestColumn = $objWorksheet->getHighestColumn(); // 取得总列数

		    $data = [];
		    for ($row = 2; $row <= $highestRow; $row++) {
		        for ($column = 'A'; $column <= $highestColumn; $column++) {
		            $val = $objWorksheet->getCellByColumnAndRow(ord($column) - 65,$row)->getValue();/**ord()将字符转为十进制数*/

		            switch($column){
		            	case 'A':
		            		$data[$row-1]['date'] = gmdate("Y-m-d H:i:s", PHPExcel_Shared_Date::ExcelToPHP($val));
		            		break;
		            	case 'B':
		            		$data[$row-1]['count'] = $val;
		            		break;
		            	case 'C':
		            		$data[$row-1]['s_price'] = $val;
		            		break;
		            	case 'D':
		            		$data[$row-1]['t_price'] = $val;
		            		break;
		            	case 'E':
		            		$data[$row-1]['name'] = $val;
		            		break;
		            	case 'F':
		            		$data[$row-1]['classname'] = $val;
		            		break;
		            }

		        }
		    }
		    $this->data = array_values($data);
	}

	function fenxi(){

		while(!empty($this->data)){
			$this->itemArray();
		};
	}

	function itemArray(){
		$item_array=[];
		foreach($this->data as $key=>$vo){
			if($key == 0){
				$now_item = $vo['name'];
				$item_array[] = $this->data[$key];
				unset($this->data[0]);
			}

			if($key>0 && $vo['name'] == $now_item){
				$item_array[] = $this->data[$key];
				unset($this->data[$key]);
			}
	    }
	    $this->data = array_values($this->data);
	    $this->createExcel($item_array);
	}

	function createExcel($item_array){
		$filename = 'excel/'.iconv("utf-8", "gb2312", $item_array[0]['name']).'.xls'; 

		if (file_exists($filename)) {

			return false;
		}

		$min_time = $item_array[0]['date'];
		$max_time = $item_array[0]['date'];

		$total_count = 0;
		$total_price = 0;
		$canvas_num = array_fill(6,18, 0);
		
		foreach($item_array as $key=>$vo){
			if($vo['date'] > $max_time) $max_time = $vo['date'];
			if($vo['date'] < $min_time) $min_time = $vo['date'];
			$total_count += $vo['count'];
			$total_price += $vo['t_price'];

			$now_h = date('H',strtotime($vo['date']));

			if($now_h >= 6 && $now_h <= 23)
				$canvas_num[(int)$now_h] += $vo['count'];
		}



		//一共卖了多少天
		$item_days = ceil((strtotime($max_time)-strtotime($min_time))/86400);
		$canvas_num2 = array_fill(0,$item_days,0);

		foreach($item_array as $key=>$vo){
			$canvas_num2_key=ceil((strtotime($vo['date'])-strtotime($min_time))/86400);
			@$canvas_num2[(int)$canvas_num2_key] += $vo['count'];
		}

		$img_src1 = '';
		$img_src2 = '';

		$img_src1 = $this->create_img(array_values($canvas_num),$item_array[0]['name'],'25');
		$img_src2 = $this->create_img(array_values($canvas_num2),$item_array[0]['name'],'10',strtotime($max_time),strtotime($min_time));
		// if($img_src1!=''){
		// 	$this->create_excel_do($item_array,$img_src1,$total_count,$total_price,$min_time,$max_time);
		// }
	}

	function create_img($data,$title,$size,$max_time=0,$min_time=0){
		
		if($size==25)
            $title .= '日销量';
        else
            $title .= '销售时间趋势';

		$title_file = iconv("utf-8", "gbk", $title);
		$title_file = str_replace(['/','\\',':','*','"','<','>','|','?'],'_',$title_file);

		 if($max_time==0 && $min_time==0)
		 	$xdata = array('6h','7h','8h','9h','10h','11h','12h','13h','14h','15h','16h','17h','18h','19h','20h','21h','22h','23h');
		 else{

		 	$xdata=array();
		 	$a = date('m-d',$min_time);
		 	array_push($xdata,$a);
		 	$middle_time = $min_time;
		 	while($middle_time < $max_time){
		 		$middle_time += 24*60*60;
				array_push($xdata,date('m-d',$middle_time));
		 	}
		 }

		 $ydata = $data;
		 $color = array();
		 $seriesName = '';
		 $title = $title;
		 $Img = new Chart($title,$xdata,$ydata,$seriesName,$size,$title_file);
		 return $Img->paintLineChart();
		}


	function create_excel_do($data,$imgsrc,$total_count,$total_price,$min_time,$max_time){

		

		$objPHPExcel = new PHPExcel();


		$objPHPExcel->setActiveSheetIndex(0)
		            ->setCellValue('A1', '销售时间')
		            ->setCellValue('B1', '销售数量')
		            ->setCellValue('C1', '销售价格')
		            ->setCellValue('D1', '销售总价')
		            ->setCellValue('E1', '名称')
		            ->setCellValue('F1', '类名')
		            ->setCellValue('L3', '销售总数')
		            ->setCellValue('M3', '销售总价')
		            ->setCellValue('N3', '最小时间')
		            ->setCellValue('O3', '最大时间')
		            ->setCellValue('P3', '售卖天数');


		foreach($data as $key=>$vo){
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.($key+2),$vo['date']);
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.($key+2),$vo['count']);
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.($key+2),$vo['s_price']);
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.($key+2),$vo['t_price']);
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.($key+2),$vo['name']);
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.($key+2),$vo['classname']);
		}


		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('L4',$total_count);
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('M4',$total_price);
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('N4',$min_time);
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('O4',$max_time);
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('P4',ceil((strtotime($max_time) - strtotime($min_time))/86400).'天');


		$objDrawing = new PHPExcel_Worksheet_Drawing();
		$objDrawing->setName('img');
		$objDrawing->setDescription('img');
		$objDrawing->setPath($imgsrc);
		$objDrawing->setCoordinates('I6');
		$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());


		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
		$new_excel_name = iconv("utf-8", "gb2312", $data[0]['name']);
		$new_excel_name = str_replace(['/','\\',':','*','"','<','>','|','?'],'_',$new_excel_name);
		$objWriter->save('excel/'.$new_excel_name.'.xls');

		}	
}



$a  = new loadExcel();
$a->load();
$a->fenxi();