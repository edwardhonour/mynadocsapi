<?php

ini_set('display_errors',1);
ini_set('display_startup_errors',1);
ini_set('memory_limit',-1);
ini_set('max_execution_time', 3000);
ini_set('upload_max_filesize', '80M');
ini_set('memory_limit', -1);
ini_set('post_max_size', '8M');
ini_set('KeepAliveTimeout', '300');
header('Access-Control-Allow-Headers: Access-Control-Allow-Origin, Content-Type, Authorization');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET,PUT,POST,DELETE,PATCH,OPTIONS');
header('Content-type: application/json');
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require('/var/www/classes/class.XRDB.php');
require('class.XRDB.php');
require('simpleXLSX.php');
require('class.enrollment.php');

$Y=new XRDB();
$X=new DB();

function process_claim($elt,$company_name) {
           $X=new DB();
	   $p=array();
	   $p['table_name']="apa_claims";
	   $p['action']="insert";			
	   $p['claim_number']=$elt[0];
	   $p['member_id']=$elt[1];
	   $p['member_name']=str_replace(' ','',$elt[2]);
	   $p['provider_name']=$elt[3];
	   $p['date_of_service']=$elt[4];
	   $p['charged_amount']=$elt[5];
	   $p['allowed_amount']=$elt[6];
	   $p['paid_amount']=$elt[7];
	   $p['status']=$elt[8];
	   $p['paid_date']=$elt[9];
	   $p['company_name']=$company_name;
           $sql="select * from apa_claims where claim_number = '" . $p['claim_number'] . "'";
           $gg=$X->sql($sql);
           if (sizeof($gg)>0) {
               $p['id']=$gg[0]['id'];
	   }
	   print_r($p);
           $X->post($p);
}

print_r($_FILES);
$d=$_FILES["name"]["name"];
$target_file="/var/www/docs/" . $d;
if (move_uploaded_file($_FILES["name"]["tmp_name"], $target_file)) { 

if ( $xlsx = SimpleXLSX::parse($target_file) ) {
	  
	$company_name ="";

	$i=1;
	try {
	foreach ($xlsx->rows(0) as $elt) {
		print_r($elt);
		if ($elt[1]!="") { 
			if (substr($elt[1],0,1)=='0') { 
			    process_claim($elt,$company_name);
			} else {
			    $company_name = $elt[1];
			}
		}	
	    $i++;
	} 
        } catch (Exception $e) {

        }

}

}

?>
