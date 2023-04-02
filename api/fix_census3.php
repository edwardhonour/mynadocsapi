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

$sql="select distinct primary_member_id, member_name from apa_claims";
$list=$X->sql($sql);

foreach($list as $l) {
	$sql="select * from nua_census where replace(employee_name,' ','') = '" . str_replace("'","",$l['member_name']) . "'";
	$c=$X->sql($sql);
	if (sizeof($c)>0) {
		print_r($c);
   //           $sql="update apa_claims set plan_code = '" . $c[0]['plan'] . "' where membsr_id = '" . $l['member_id'] . "'";
//	      $X->execute($sql);
              $sql="update apa_claims set coverage_level = '" . $c[0]['coverage_level'] . "' where primary_member_id = '" . $l['primary_member_id'] . "'";
	      $X->execute($sql);
  //            $sql="update apa_claims set date_of_birth = '" . $c[0]['dob'] . "' where member_id = '" . $l['member_id'] . "'";
//	      $X->execute($sql);
  //            $sql="update apa_claims set gender = '" . $c[0]['gender'] . "' where member_id = '" . $l['member_id'] . "'";
//	      $X->execute($sql);
  //            $sql="update apa_claims set eff_dt = '" . $c[0]['eff_dt'] . "' where member_id = '" . $l['member_id'] . "'";
//	      $X->execute($sql);
  //            $sql="update apa_claims set dependent = '" . $c[0]['dependent'] . "' where member_id = '" . $l['member_id'] . "'";
//	      $X->execute($sql);
  //            $sql="update apa_claims set primary_member_id = '" . $c[0]['apa_employee_id'] . "' where member_id = '" . $l['member_id'] . "'";
//	      $X->execute($sql);
	}
}

?>
