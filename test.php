<?php
ini_set('display_errors',1);
ini_set('display_startup_errors',1);
header('Access-Control-Allow-Headers: Access-Control-Allow-Origin, Content-Type, Authorization');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET,PUT,POST,DELETE,PATCH,OPTIONS');
header('Content-type: application/json');
require('/var/www/classes/class.XRDB.php');

$X=new XRDB();

$sql="SELECT COUNT(*) AS C FROM BEP20_TOKEN";
$h=$X->sql($sql);
print_r($h);
?>
