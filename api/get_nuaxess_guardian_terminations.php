<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');
require_once('class.enrollment.php');

ini_set('display_errors',1);
ini_set('display_startup_errors',1);

$month_id="2022-05";

$X=new XRDB();
$E=new ENROLLMENT();

if (isset($argv[1])) { $_GET['id']=$argv[1]; $_GET['display']="F"; }
if (isset($argv[2])) { $_GET['display']=$argv[2]; }

if (isset($_GET['id'])) {
     $month_id = $_GET['id'];
} else {
    die();
}
$display = "B";

$sql="select * from nua_monthly_member_terminations where month_id >= '" . $month_id . "' and plan_type in ('*DENTAL*','*LIFE*','*ADD*') and company_id in (select id from nua_company where org_id = 17)";
$data=$X->sql($sql);
$E->getGuardianTerminations($data,"B");

?>


