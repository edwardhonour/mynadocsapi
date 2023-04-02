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
//header('Content-type: application/json');
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
	   $p['table_name']="va_facility_load";
	   $p['action']="insert";			
	   $p['FRPP_ASSET_ID']=$elt[0];
	   $p['ISC_C_ID']=$elt[1];
	   $p['ORGANIZATION']=$elt[2];
	   $p['SUB_ORGANIZATION']=$elt[3];
	   $p['PROPERTY_TYPE']=$elt[4];
	   $p['PROPERTY_USE']=$elt[5];
	   $p['LEASED_OWNED']=$elt[6];
	   $p['VISIN']=$elt[7];
	   $p['ADDRESS']=$elt[8];
	   $p['STATE']=$elt[9];
	   $p['CITY']=$elt[10];
	   $p['COUNTY']=$elt[11];
	   $p['ZIP']=$elt[12];
	   $p['LATITUDE']=$elt[13];
	   $p['LONGITUDE']=$elt[14];
	   $p['LEASE_DATE']=$elt[15];
	   $p['REPLACEMENT_COST']=$elt[16];
	   $p['SQUARE_FEET']=$elt[17];
	   $p['ACRES']=$elt[18];
	   $p['POPULATION']=$elt[19];
	   $p['CHILD_CARE']=$elt[20];
	   $p['LAST_FSA_DATE']=$elt[21];
           $sql="select * from va_facility_load where FRPP_ASSET_ID = '" . $p['FRPP_ASSET_ID'] . "'";
           $gg=$X->sql($sql);
           if (sizeof($gg)>0) {
               $p['id']=$gg[0]['id'];
	   }
	   print_r($p);
           $X->post($p);
}


$d=$_FILES["name"]["name"];
$target_file="/var/www/docs/" . $d;
print_r($_FILES);
if (move_uploaded_file($_FILES["name"]["tmp_name"], $target_file)) { 
$row = 1;
if (($handle = fopen($target_file, "r")) !== FALSE) {
    while (($data = fgetcsv($handle, 1000, ",")) !== FALSE) {
        $num = count($data);
        echo "<p> $num fields in line $row: <br /></p>\n";
        $row++;
	if ($data[0]!="") process_census($data);
    }
    fclose($handle);
}

//if ( $xlsx = SimpleXLSX::parse($target_file) ) {
//	  
//	$i=1;
//	try {
//	foreach ($xlsx->rows(0) as $elt) {
//		if ($i>10) { 
//			if ($elt[0]!="") process_census($elt);
//		}	
//	    $i++;
//	} 
  //      } catch (Exception $e) {
//
  //      }
//
//}

}

?>
