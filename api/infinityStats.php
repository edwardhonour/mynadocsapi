<?php
ini_set('memory_limit', -1);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');
require_once('class.enrollment2.php');

ini_set('display_errors',1);
ini_set('display_startup_errors',1);


$X=new XRDB();
$E=new ENROLLMENT();
if (isset($_GET['billingCycle'])) {
	$billingCycle = $_GET['billingCycle'];
  }
else
{
   echo("Billing cycle not provided");
   die();
}

$agency_id = $_GET['agency_id'];
$E->getAgencyCommissionReport($billingCycle,$agency_id);

?>


