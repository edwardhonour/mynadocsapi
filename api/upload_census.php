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

function process_claim($elt,$member_id,$company_name) {
           $X=new DB();
	   $p=array();
	   $p['table_name']="apa_census";
	   $p['action']="insert";			
	   $p['apa_employee_id']=$member_id;
	   $p['member_type']=$elt[1];
	   $p['coverage_level']=$elt[2];
	   $p['employee_name']=$elt[3];
	   $p['dob']=$elt[4];
	   $p['gender']=$elt[5];
	   $p['status']=$elt[6];
	   $p['eff_dt']=$elt[7];
	   $p['term_dt']=$elt[8];
	   $p['plan']=$elt[9];
	   $p['company_name']=$company_name;
           $sql="select * from apa_census where apa_employee_id = '" . $p['apa_employee_id'] . "' and employee_name = '" . $p['employee_name'] . "'";
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
	$member_id ="";

	$i=1;
	try {
	foreach ($xlsx->rows(0) as $elt) {
		print_r($elt);
		if ($elt[0]!="ID") { 
			if (substr($elt[0],0,1)=='0') { 
			    $member_id=$elt[0];
			} else {
			    if ($elt[0]!='') $company_name = $elt[0];
			}
			if ($elt[1]!=''&&$elt[1]!='Type') process_claim($elt,$member_id,$company_name);
		}	
	    $i++;
	} 
        } catch (Exception $e) {

        }

}

}

?>
