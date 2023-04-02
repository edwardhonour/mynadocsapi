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

$sql="select * from apa_claims";
$list=$X->sql($sql);

echo "COMPANY_ID,MEMVER_ID,PRIMARY_MEMBER_ID,DATE_OF_SERVICE,CHARGED_AMOUNT,ALLOWED_AMOUNT,PAID_AMOUNT,OUT_OF_POCKET,PAID_DATE,COMPANY_NAME,PLAN_CODE,COVERAGE_LEVEL,EFF_DT,DATE_OF_BIRTH,GENDER,GROUP_SIZE<br>";
foreach($list as $l) {
	echo $l['company_id'] . ",";
        echo '"' . $l['claim_number'] . '",';
        echo '"' . $l['member_id'] . '",';
        echo '"' . $l['primary_member_id'] . '",';
        echo '"' . $l['date_of_service'] . '",';
        echo '"' . $l['charged_amount'] . '",';
        echo '"' . $l['allowed_amount'] . '",';
        echo '"' . $l['paid_amount'] . '",';
        echo '"' . number_format(floatval($l['allowed_amount'])-floatval($l['paid_amount']),2) . '",';
        echo '"' . $l['paid_date'] . '",';
        echo '"' . $l['company_name'] . '",';
        echo '"' . $l['plan_code'] . '",';
        echo '"' . $l['coverage_level'] . '",';
        echo '"' . $l['eff_dt'] . '",';
        echo '"' . $l['date_of_birth'] . '",';
        echo '"' . $l['gender'] . '",';
        echo '"' . $l['group_size'] . '"<br>';
}

?>
