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

$sql="select distinct company_name, company_id from apa_claims where company_id not in (0,-1)";
$list=$X->sql($sql);

foreach($list as $l) {
	print_r($l);
              $sql="update apa_claims set company_id = " . $l['company_id'] . " where company_name = '" . str_replace("'","''",$l['company_name']) . "' and company_id in (0,-1)";
	      $X->execute($sql);
}

?>
