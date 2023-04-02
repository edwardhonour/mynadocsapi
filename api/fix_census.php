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

$sql="select distinct member_id, member_name from apa_claims";
$list=$X->sql($sql);

foreach($list as $l) {
	$sql="select * from nua_employee where apa_member_id = '" . $l['member_id'] . "'";
	$c=$X->sql($sql);
	if (sizeof($c)>0) {
              $sql="update apa_claims set company_id = " . $c[0]['company_id'] . " where member_id = '" . $l['member_id'] . "'";
	      $X->execute($sql);
              $sql="update apa_claims set plan_code = '" . $c[0]['medical_plan'] . "' where member_id = '" . $l['member_id'] . "'";
	      $X->execute($sql);
              $sql="update apa_claims set coverage_level = '" . $c[0]['medical_plan_level'] . "' where member_id = '" . $l['member_id'] . "'";
	      $X->execute($sql);
              $sql="update apa_claims set date_of_birth = '" . $c[0]['date_of_birth'] . "' where member_id = '" . $l['member_id'] . "'";
	      $X->execute($sql);
              $sql="update apa_claims set gender = '" . $c[0]['gender'] . "' where member_id = '" . $l['member_id'] . "'";
	      $X->execute($sql);
              $sql="update apa_claims set plan_code = '" . $c[0]['apa_medical_plan'] . "' where member_id = '" . $l['member_id'] . "'";
	      $X->execute($sql);
              $sql="update apa_claims set eff_dt = '" . $c[0]['apa_medical_eff_dt'] . "' where member_id = '" . $l['member_id'] . "'";
	      $X->execute($sql);
	}
}

?>
