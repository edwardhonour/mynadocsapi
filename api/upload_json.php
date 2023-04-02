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
//require 'vendor/autoload.php';

//use PhpOffice\PhpSpreadsheet\Spreadsheet;
//use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require('/var/www/classes/class.XRDB.php');
require('class.XRDB.php');
//require('simpleXLSX.php');
//require('class.enrollment.php');

$Y=new XRDB();
$X=new DB();

function process_record($elt,$table_name) {
           $X=new DB();
	   if (sizeof($elt)>0) {
	   foreach($elt as $e) {
	      $p=array();
	      $p['table_name']=$table_name;
	      $p['action']="insert";			
              foreach($e as $name=>$value) {
                 $p[$name]=$value;
              }
	   print_r($p);
	      if (sizeof($p) > 0) {
		      $X->post($p);
	      }
	   }
	   }
}


$d=$_FILES["name"]["name"];
$target_file="/var/www/docs/" . $d;
print_r($_FILES);
print_r($_POST);
if (move_uploaded_file($_FILES["name"]["tmp_name"], $target_file)) { 
	$raw=file_get_contents($target_file);
        $data=json_decode($raw,true);
        process_record($data,$_POST['table_name']);
}

