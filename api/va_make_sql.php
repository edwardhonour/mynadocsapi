<?php

ini_set('display_errors',1);
ini_set('display_startup_errors',1);
ini_set('memory_limit',-1);
ini_set('max_execution_time', 3000);
ini_set('upload_max_filesize', '8M');
ini_set('memory_limit', -1);
ini_set('post_max_size', '8M');
ini_set('KeepAliveTimeout', '300');
header('Access-Control-Allow-Headers: Access-Control-Allow-Origin, Content-Type, Authorization');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET,PUT,POST,DELETE,PATCH,OPTIONS');
//header('Content-type: application/json');

require('class.XRDB.php');
require('simpleXLSX.php');

$X=new DB();

$sql="select id, ADDRESS, STATE, SUB_ORGANIZATION from va_facility_load where BUILDING_NBR = '' order by 2, 3";
$list=$X->sql($sql);

foreach($list as $l) {

	$b=substr($l['STATE'],0,2);
        if ($l['SUB_ORGANIZATION']=='National Cemetary Administration') $b.="VC";
        if ($l['SUB_ORGANIZATION']=='Veterans Health Administration') $b.="VH";
        if ($l['SUB_ORGANIZATION']=='Veterans Benefits Administration') $b.="VB";
        if ($l['SUB_ORGANIZATION']=='Department of Veterans Affairs') $b.="VA";
	if ($l['SUB_ORGANIZATION']=='Office of Acquisiton and Materials Management') $b.="VM";
        if ($l['SUB_ORGANIZATION']=='Office of Information Systems and Telecommunications') $b.="VS";
        if ($l['SUB_ORGANIZATION']=='Administrative Service') $b.="VV";
        if ($l['SUB_ORGANIZATION']=='Inspector General') $b.="VG";
        if ($l['SUB_ORGANIZATION']=='Office of Acquisition and Material Management') $b.="VM";

	$sql="select count(*) as c from va_facility_load where SUBSTR(BUILDING_NBR,1,4) = '" . $b . "'";
        $c=$X->sql($sql);
	$m=intval($c[0]['c']);
	$m++;
	if ($m<10) {
           $j="00" . $m;
	}
	if ($m>=10&&$m<100) {
           $j="0" . $m;
	}
	if ($m>=100) {
           $j=$m;
	}
	$b.=$j;
	$sql="update va_facility_load set BUILDING_NBR = '" . $b . "' where id = " . $l['id'];
	echo $sql . "<br>";
	$X->execute($sql);

}

?>
