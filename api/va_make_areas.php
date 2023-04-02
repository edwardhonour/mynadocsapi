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
header('Content-type: application/json');

require('class.XRDB.php');
require('simpleXLSX.php');

$X=new DB();

$sql="select distinct VISIN, ADDRESS, CITY, STATE, ZIP from va_facility_load order by 1, 2, 3, 4, 5";
$list=$X->sql($sql);

foreach($list as $l) {
	$sql="select count(*) as c from va_area where AREA_CODE = '" . $l['STATE'] . '-' . $l['ZIP'] . "'";
        $c=$X->sql($sql);
        $post=array();
        $post['table_name']="va_area";
	if ($c[0]['c']==0) {
             $post['AREA_CODE']=$l['STATE'] . '-' . $l['ZIP'];
	} else {
             $post['AREA_CODE']=$l['STATE'] . '-' . $l['ZIP'] . "-" . $c[0]['c'];
	}
        $post['AREA_NAME']=$l['VISIN'] . '-' . str_replace("'","",$l['ADDRESS']) . "-" . $l['CITY'];
	$post['DISTRICT_CODE']=$l['VISIN'];
	$post['REGION_ID']=$l['STATE'];
    	print_r($post);
	$X->post($post);
	$sql="update va_facility_load set p_beat = '" . str_replace("'","",$post['AREA_CODE']) . "' where ";
	$sql.=" VISIN = '" . $l['VISIN'] . "' and ADDRESS = '" . str_replace("'","",$l['ADDRESS']) . "' and CITY = '" . str_replace("'","",$l['CITY']) . "' AND ";
	$sql.=" STATE = '" . $l['STATE'] . "' and ZIP = '" . $l['ZIP'] . "'";
	echo $sql;
	$X->execute($sql);

}

?>
