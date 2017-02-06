<!DOCTYPE html>
<html>
<head>
	<title>归档情况表</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body>
<a href="/bingan.html" style="font-size: 20px">返回</a>
<table width="500" border=" 1" style="margin: 10px; font-size: 15px">
<tr>
	<th>序号</th>
	<th>出院专业</th>
	<th>住院号</th>
	<th>姓名</th>
	<th>出院时间</th>
</tr>
<?php
	$department=$_POST["department"];
	require_once 'Classes/PHPExcel.php';
	$inputfile=iconv('utf-8', 'gbk//IGNORE','F:\工作\出院病人勾选\新政17-1.xls');
	$objreader=PHPExcel_IOFactory::createReaderForFile($inputfile);
	$excelobj=$objreader->load($inputfile);
	$excelobj->setActiveSheetIndex(0);
	$worksheet=$excelobj->getActiveSheet();
	$lastRow=$worksheet->getHighestRow();
	switch ($department) {
		case '新政内一病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if($depart=='新政心血管专业' and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}
			break;
		case '新政内二病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if($depart=='新政呼吸专业' and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}
			break;
		case '新政内三病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新政消化专业' or $depart=='新政内分泌专业' or $depart=='新政风湿免疫专业') and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}
			break;
		case '新政内四病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新神内专业' or $depart=='新肾内专业') and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}			
			break;
		case '新政外一病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新政普外专业' or $depart=='新政脑外专业' or $depart=='新政肝胆外科专业') and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}
			break;
		case '新政外二病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新政骨科专业' or $depart=='新政骨创伤专业' or $depart=='新政骨关节专业' or $depart=='新政脊柱专业') and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}		
			break;
		case '新政外三病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新政胸外专业' or $depart=='新政泌尿专业' or $depart=='新政肛肠专业') and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}			
			break;
		case '新政急诊病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新政急诊外科' or $depart=='新政急诊内科')  and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}			
			break;
		case '新政妇产病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新政妇科专业' or $depart=='新政产科专业' or $depart=='新政计划生育专业')  and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}		
			break;
		case '新政儿科病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新政儿科专业' or $depart=='新政新生儿专业') and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}			
			break;
		case '新政康复病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if($depart=='新政康复专业' and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}		
			break;
		case '新政肿瘤病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if($depart=='新政肿瘤专业'  and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}		
			break;
		case '新政疼痛病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新政疼痛专业' or $depart=='新政疼痛科')  and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}		
			break;
		case '新政重症监护室':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if($depart=='新政_ICU'  and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}			
			break;
		case '新政五官病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='新眼科专业' or $depart=='新耳鼻咽喉专业' or $depart=='新口腔专业' or $depart=='新皮肤专业') and $dctime==""){
					echo "<tr><td>";
					echo $num;
					$num++;
					echo "</td><td>";
					echo $worksheet->getCell('B'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('D'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('E'.$row)->getValue();
					echo "</td><td>";
					echo $worksheet->getCell('F'.$row)->getValue();
					echo "</td></tr>";
				}
			}			
			break;
		default:
			echo 'sorry, this page 404!!!!!';
			break;
	}
?>
</table>
</body>
</html>
