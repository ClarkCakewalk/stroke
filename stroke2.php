<?php
$filetype=strrchr($_GET["id"], ".");
$filename=substr($_GET["id"],0, strripos($_GET["id"], $filetype))."_stroke.xlsx";
header("Content-type:application/vnd.ms-excel");
header("Content-Disposition:filename=".$filename);
ob_clean();
flush();
readfile("filetmp/".$filename);
unlink("filetmp/".$_GET["id"]);
unlink("filetmp/".$filename);
?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>無標題文件</title>
</head>

<body>
</body>
</html>
