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
		case '金城内一病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='金城心血管专业' or $depart=='金城消化专业' or $depart=='金城内分泌专业' or $depart=='金城儿科专业') and $dctime==""){
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
		case '金城内二病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='金城呼吸专业' or $depart=='金城神经专业' or $depart=='金城肾病专业' or $depart=='金城老年病专业') and $dctime==""){
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
		case '金城感染病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='金城传染专业' or $depart=='金城感染病区') and $dctime==""){
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
				if(($depart=='金城普外专业' or $depart=='金城脑外专业' or $depart=='金城肝胆外科') and $dctime==""){
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
		case '金城外二病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='金城骨科专业' or $depart=='金城骨创伤专业' or $depart=='金城骨关节专业' or $depart=='金城脊柱专业') and $dctime==""){
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
		case '金城外三病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='金城胸外专业' or $depart=='金城泌尿专业' or $depart=='金城肛肠外科') and $dctime==""){
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
		case '金城急诊病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='金城急诊外科' or $depart=='金城急诊内科')  and $dctime==""){
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
		case '金城妇产病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='金城妇科专业' or $depart=='金城产科专业' or $depart=='金城计划生育专业')  and $dctime==""){
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
		case '金城康复病区':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if(($depart=='金城康复专业' or $depart=='金城理疗专业') and $dctime==""){
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
		case '金城重症监护室':
			$num=1;
			for ($row=3;$row<=$lastRow;$row++){
				$depart=$worksheet->getCell('B'.$row)->getValue();
				$dctime=$worksheet->getCell('H'.$row)->getValue();
				if($depart=='金城重症监护室ICU'  and $dctime==""){
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
				if(($depart=='金眼科专业' or $depart=='金耳鼻喉专业' or $depart=='金城口腔专业' or $depart=='金皮肤专业') and $dctime==""){
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
