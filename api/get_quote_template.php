<?php
ini_set('memory_limit', -1);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');
require_once('class.enrollment.php');

ini_set('display_errors',1);
ini_set('display_startup_errors',1);

$month_id="2022-06";

$X=new XRDB();
$E=new ENROLLMENT();

if (isset($argv[1])) { $_GET['id']=$argv[1]; $_GET['display']="F"; }
if (isset($argv[2])) { $_GET['display']=$argv[2]; }

if (isset($_GET['id'])) {
     $id = $_GET['id'];
     $company_id = $_GET['id'];
} else {
    die();
}
if (isset($_GET['display'])) {
     $display = $_GET['display'];
} else {
     $display = "B";
}

if (isset($_GET['uid'])) {
     $uid = $_GET['uid'];
     $sql="select * from nua_user where id = " . $uid;
     $u=$X->sql($sql);
     $role=$u[0]['role'];
} else {
     $uid = "0";
     $role = "broker";
}
$role="sadmin";
if ($company_id==0) {
    $company=array();
} else {
    $sql="select * from nua_company where id = " . $company_id;
    $t=$X->sql($sql);
    $company=$t[0];
}

$E->getQuoteTemplate($company_id, $month_id,"B");

?>


