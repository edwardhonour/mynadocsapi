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
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require('/var/www/classes/class.XRDB.php');
require('class.XRDB.php');
require('simpleXLSX.php');
require('class.enrollment.php');

$Y=new XRDB();
$X=new DB();

function process_census($elt) {
           $X=new DB();
	   $p=array();
	   $p['table_name']="FPS_CMOD_PROJECT";
	   $p['action']="insert";			
	   $p['PROJECT_SOURCE']=$elt[0];
	   $p['REGION_ID']=$elt[1];
	   $p['BUILDING_NBR']=$elt[2];
	   $p['FACILITY_NAME']=$elt[3];
	   $p['CITY']=$elt[4];
	   $p['STATE']=$elt[5];
	   $p['ADDRESS']=$elt[6];
	   $p['FSL']=$elt[7];
	   $p['STATUS']=$elt[8];
	   $p['ACTION']=$elt[9];
	   $p['SCOPE']=$elt[10];
	   $p['FSC_APPROVAL_STATUS']=$elt[11];
	   $p['FSC_APPROVAL_DATE']=$elt[12];
	   $p['FUNDING_STATUS']=$elt[13];
	   $p['FUNDING_DATE']=$elt[14];
	   $p['FUNDING_VEHICLE']=$elt[15];
	   $p['ITAR_SUBMITTED']=$elt[16];
	   $p['ITAR_SUBMITTED_DATE']=$elt[17];
	   $p['ITAR_COMPLETED']=$elt[18];
	   $p['ITAR_COMPLETED_DATE']=$elt[19];
	   $p['PROJECT_SOLICITATION']=$elt[20];
	   $p['PROJECT_SOLICITATION_DATE']=$elt[21];
	   $p['CONTRACT_AWARD_STATUS']=$elt[23];
	   $p['CONTRACT_AWARD_DATE']=$elt[24];
	   $p['NTP']=$elt[25];
	   $p['NTP_DATE']=$elt[26];
	   $p['COMPLETION_STATUS']=$elt[27];
	   $p['COMPLETION_DATE']=$elt[28];
	   $p['LAST_UPDATE_DATE']=$elt[29];
	   $p['LAST_UPDATE_DAYS']=$elt[30];
	   $p['PCT_COMPLETE']=$elt[31];
	   $p['CURRENT_STATUS_NOTES']=$elt[32];
	   $p['PROJECT_IDENTIFIED']=$elt[34];
	   $p['FSA_APPROVE_DATE']=$elt[35];
	   $p['USMS_VSS_PROJECT']=$elt[36];
	   $p['ACTIVE_STATUS']=$elt[37];
	   $p['EST_COMPLETION_DATE']=$elt[38];
	   $p['STATUS_NOTES']=$elt[39];
	   print_r($p);
           $project_id=$X->post($p);
	   $notes=explode("\n",$elt[33]);
	   foreach($notes as $n) {
               $pp=array();
	       $pp['table_name']="FPS_CMOD_ARCHIVED_NOTES";
	       $pp['action']="insert";
	       $pp['PROJECT_ID']=$project_id;
	       $pp['NOTE_DATE']=substr($n,0,8);
	       $pp['NOTE']=$n;
	       print_r($pp);
	       $X->post($pp);
           }
}


$d=$_FILES["name"]["name"];
$target_file="/var/www/docs/" . $d;
if (move_uploaded_file($_FILES["name"]["tmp_name"], $target_file)) { 

if ( $xlsx = SimpleXLSX::parse($target_file) ) {
	  
	$i=1;
	try {
	foreach ($xlsx->rows(0) as $elt) {
		if ($i>1) { 
			if ($elt[0]!="") process_census($elt);
		}	
	    $i++;
	} 
        } catch (Exception $e) {

        }

}

}

?>
