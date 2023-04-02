<?php
ini_set('memory_limit', -1);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');

ini_set('display_errors',1);
ini_set('display_startup_errors',1);


$X=new XRDB();
//$month_id = $_GET['month_id'];
//$agency_id = $_GET['agency_id'];
//$payment_date = $_GET['payment_date'];
//$payment_id = $_GET['payment_id'];

function payCommissions($billingCycle, $agency_id,$payment_date,$payment_id=0)
{
	$X=new XRDB();


	$sql="select id, org_name from nua_agency where id = " . $agency_id . " order by org_name ";
	$eft=$X->sql($sql);

        $row=1;

        foreach($eft as $eftRow) {
		
	    $commission_total=0;
            $row++;
	
            $sql="select nua_agency_company.company_id as id, nua_company.company_name as company_name, plan_type, ";
            $sql.=" commission_rate, commission_type from ";
    	    $sql.="nua_agency_company, nua_company where nua_agency_company.agency_id = " . $eftRow['id']; 
	    $sql.=" and nua_agency_company.company_id = nua_company.id order by nua_company.company_name";

	    echo $sql;
	    $c=$X->sql($sql);

	    print_r($c);

		foreach($c as $comp) {
			$row++;
	
		        if ($comp['commission_type']=="FLAT") {
		           	$sql="select count(*) as c from nua_monthly_member_census where month_id = '" . $billingCycle . "' and company_id = " . $comp['id'];
			        $sql.=" and plan_type = '" . $comp['plan_type'] . "' and dependent_code = ''";
	         		$b=$X->sql($sql);
			        $basis=floatval($b[0]['c']);
		} else {
		   	        $sql="select coverage_price from nua_monthly_member_census where month_id = '" . $billingCycle . "' and company_id = " . $comp['id'];
			        $sql.=" and plan_type = '" . $comp['plan_type'] . "' and dependent_code = ''";
			        $b=$X->sql($sql);
			        $total=0;
			        foreach($b as $c) {
                                    $total+=floatval($c['coverage_price']);
			        }
			        $basis=$total;
		        }
		        $sql="select id, paid from nua_company_invoice where company_id = " . $comp['id'] . " and month_id = '" . $billingCycle . "'";
		        $gg=$X->sql($sql);
			print_r($gg);
		        if (sizeof($gg)==0) {
			     $p="No";
		         } else {
		           $p=$gg[0]['paid'];
			   if ($p=='N') $p="No";
			   if ($p=='Par') $p="No";
			   if ($p=='Y') $p="Yes";
	          	}
		        $commission=0;
		        if ($p=='No') {
			    $commission=0;
			    $commission_paid="No";
		        } else {
		            if (sizeof($gg)>0) {
		                $sql="select count(*) as c from nua_agency_paid where agency_id = " . $agency_id;
		                $sql.=" and invoice_id = " . $gg[0]['id'] . " and plan_type = '" . $comp['plan_type'] . "'";
		                $h=$X->sql($sql);
				print_r($h);
		                if ($h[0]['c']>0) { 
			            $commission_paid="Yes";
		                } else {
			            $commission_paid="No";
				    $ppp=array();
				    $ppp['table_name']="nua_agency_paid";
				    $ppp['action']="insert";
				    $ppp['agency_id']=$agency_id;
				    $ppp['invoice_id']=$gg[0]['id'];
				    $ppp['plan_type']=$comp['plan_type'];
				    $ppp['payment_id']=$payment_id;
				    $ppp['payment_date']=$payment_date;
				    print_r($ppp);
				    $X->post($ppp);
		                }
		                } 
		}
		}
	}
}

$month_id="2022-10";
$agency_id=48;
$payment_date="2022-10-31";
$payment_id=0;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=49;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=50;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=51;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=52;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=53;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=54;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=55;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=56;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=57;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=58;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=59;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=60;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);
$agency_id=61;
payCommissions($month_id, $agency_id, $payment_date, $payment_id);

?>


