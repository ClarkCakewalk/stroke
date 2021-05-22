<?php
set_time_limit(0);
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function num2alpha($n)  //數字轉英文(0=>A、1=>B、26=>AA...以此類推)
{
    for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
        $r = chr($n%26 + 0x41) . $r; 
    return $r; 
}

function StrokeProcess ($v) {
	if (preg_match_all('/^([\x7f-\xff]+)$/', $v)) {
		require 'config.php';
		for($i=0; $i<mb_strlen($v); $i++) {
			$word=mb_substr($v, $i, 1, "utf-8");
			$tmp=substr(json_encode($word),3,-1);
			$qrystroke="select `stroke`, `oldstroke`, `KangXi` from `CNS` where `code` like '{$tmp}'";
			$result=$dbh->query($qrystroke);
			$words[$i]=$result->fetch();
			$words[$i]["word"]=$word;
		}
	}
	else $words='error';
	return $words;
}    
?>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>姓名筆畫查詢</title>
</head>

<body>
<script type="text/javascript" language="javascript">
function checkfile(sender) {
    var validExts = new Array(".xlsx", ".xls");
    var fileExt = sender.value;
    fileExt = fileExt.substring(fileExt.lastIndexOf('.'));
    if (validExts.indexOf(fileExt) < 0) {
      alert("檔案格式錯誤，請選擇" +
               validExts.toString() + "格式的檔案。");
      return false;
    }
    else return true;
}
</script>
<h4>中文字筆畫批次查詢</h4>
<form id="upload" name="upload" method="post" action="" enctype="multipart/form-data">
  <p>
    <label>請選擇上傳檔案
      <input type="file" name="file" id="file" onchange="checkfile(this);" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"/>
    </label>
  </p>
  <p>
    <label>
      <input type="submit" name="button" id="button" value="開始轉換"/>
    </label>
  </p>
</form>
<p>上傳檔案說明：</p>
<p>1. 上傳檔案限制為excel格式檔案。</p>
<p>2. 需計算筆畫文字（姓名）整輸入於B欄中，第一列為標題列，系統將從第二列開始進行轉換。</p>
<?php
if (!empty($_FILES)) {
$filetype=strrchr($_FILES["file"]["name"], ".");
$filename=substr($_FILES["file"]["name"],0, strripos($_FILES["file"]["name"], $filetype));
if ($filetype==".xls" or $filetype==".xlsx") {
	move_uploaded_file($_FILES["file"]["tmp_name"],"./filetmp/".$_FILES["file"]["name"]);
	$loadFile=PhpOffice\PhpSpreadsheet\IOFactory::load('./filetmp/'.$_FILES["file"]["name"]);
	$sheetData=$loadFile->getSheet(0)->toArray();
	for ($i=1; $i<count($sheetData); $i++) {
		$names=StrokeProcess($sheetData[$i][1]);
		$countwords=count($names);
		if (empty($maxwords) or $countwords>$maxwords) {$maxwords=$countwords;}
		for ($j=0; $j<$countwords; $j++) {
			$sheetData[$i][$j*3+2]=$names[$j]["stroke"];
			if (!empty($names[$j]["oldstroke"])) $sheetData[$i][$j*3+3]=$names[$j]["oldstroke"];
				else $sheetData[$i][$j*3+3]=$names[$j]["stroke"];
			if (!empty($names[$j]["KangXi"])) $sheetData[$i][$j*3+4]=$names[$j]["KangXi"];
				else $sheetData[$i][$j*3+4]=$names[$j]["stroke"];
		}
	}
	for ($i=0; $i<$maxwords; $i++) {
		$t=$i+1;
		$sheetData[0][$i*3+2]=$t."_new";
		$sheetData[0][$i*3+3]=$t."_old";
		$sheetData[0][$i*3+4]=$t."_KangXi";
	}
//	var_dump($sheetData);
	$newFilename=$filename.'_stroke.xlsx';
	$objwrite=new \PhpOffice\PhpSpreadsheet\Spreadsheet();
	$objwrite->getActiveSheet()->fromArray($sheetData, NULL, 'A1');
	$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($objwrite);
	$writer->save('filetmp/'.$newFilename);
	if (file_exists('filetmp/'.$newFilename)) {
		$zip = new ZipArchive();
		$zipfile=$filename.'.zip';
		$createZip=$zip->open(__DIR__.'/filetmp/'.$zipfile, ZipArchive::CREATE);
		if($createZip===true) {
			$zip->addFile(__DIR__.'/filetmp/'.$newFilename, $newFilename);
		}
		$zip->close();
		header("Content-type:application/zip");
		header("Content-Disposition:filename=".$zipfile);
		ob_clean();
		flush();
		readfile(__DIR__.'/filetmp/'.$zipfile);
		unlink("filetmp/".$_FILES["file"]["name"]);
		unlink("filetmp/".$newFilename);
		unlink("filetmp/".$zipfile);
	}
}
else { 
	echo "上傳檔案格式錯誤！！";
	unlink($_FILES["file"]["tmp_name"]);
}
}
?>
<hr />
<h4>中文筆畫單筆查詢</h4>
<form id="form1" name="form1" method="post" action="">
  <label>請輸入文字：
    <input type="text" name="name" id="name" />
  </label>
  <label>
    <input type="submit" name="button2" id="button2" value="查詢" />
  </label>
</form>
<?php
if (!empty($_POST["name"])) {
	$stroke=StrokeProcess($_POST["name"]);
?>
<p>您輸入的文字為：<?php echo $_POST["name"]; if($stroke=='error') echo "，請輸入中文字"; ?></p>
<?php if ($stroke!='error') { ?>
<table width="200" border="1" cellpadding="1">
  <tr>
    <th scope="col">字</th>
<?php 	for($i=0; $i<count($stroke); $i++) {?>
    <th scope="col"><?php echo $stroke[$i]["word"];?></th>
<?php } ?>
  </tr>
  <tr>
    <th scope="row">筆畫數</th>
<?php 	for($i=0; $i<count($stroke); $i++) {?>
    <td><?php echo $stroke[$i]["stroke"];?></td>
<?php } ?>
  </tr>
  <tr>
    <th scope="row">古字筆畫數</th>
<?php 	for($i=0; $i<count($stroke); $i++) {?>
    <td><?php if (empty($stroke[$i]["oldstroke"])) { echo $stroke[$i]["stroke"];} else { echo $stroke[$i]["oldstroke"];}?></td>
<?php } ?>
  </tr>
  <tr>
    <th scope="row">康熙筆畫數</th>
<?php 	for($i=0; $i<count($stroke); $i++) {?>
    <td><?php if (empty($stroke[$i]["KangXi"])) { echo $stroke[$i]["stroke"];} else { echo $stroke[$i]["KangXi"];}?></td>
<?php } ?>
  </tr> 
</table>
<p>&nbsp;</p>
<?php		
}}
?>

</body>
</html>
