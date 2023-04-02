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
$sql="select distinct company_name, company_id from nua_monthly_member_census where company_id in (select id from nua_company where ";
$sql.=" org_id = 17) and company_id in (select company_id from nua_company_invoice where billed_month_id = '2022-04') order by 1";
$list=$X->sql($sql);

foreach($list as $l) {
	$sql="select count(*) as c from nua_monthly_member_census where plan_type = '*MEDICAL*' and  month_id = '2022-04' and company_id = " . $l['company_id'];
	$b=$X->sql($sql);
	echo $l['company_name'] . "," . $b[0]['c'] . "<br>";
}

?>
