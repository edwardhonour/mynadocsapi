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
require('/var/www/classes/class.XRDB.php');

$X=new XRDB();
$post=array();
$post['table_name']="nua_doc";
$post['action']="insert";
//echo json_encode($_FILES);
$post['file_type']=$_FILES["file"]["type"];
$post['file_name']=$_FILES["file"]["name"];
if (isset($_POST["company_id"])) { $post['company_id']=$_POST["company_id"]; }
if (isset($_POST["employee_id"])) { $post['employee_id']=$_POST["employee_id"]; }
if (isset($_POST["broker_id"])) { $post['broker_id']=$_POST["broker_id"]; }
if (isset($_POST["dsc"])) { $post['dsc']=$_POST["dsc"]; }
if (isset($_POST["doc_title"])) { $post['doc_title']=$_POST["doc_title"]; }
$post['user_id']=$_POST["user_id"];
$id=$X->doc($post);
$key="nuaxess".$id;
$m=md5($key);
$d=substr($m,0,7);
$f=substr($m,7,7);
$dir="/var/www/docs/" . $d;
mkdir($dir);
$target_file=$dir . "/" . $f;

if (move_uploaded_file($_FILES["file"]["tmp_name"], $target_file)) { 
	 $output=array();
	 $output['error_code']=0;
	 $output['error_message']="";
	 echo json_encode($output);
} else {
         $output=array();
	 $output['error_code']=1;
	 $output['error_message']="An Error has Occured";
	 echo json_encode($output);
}

?>
