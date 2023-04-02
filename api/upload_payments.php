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
require('/var/www/classes/class.XRDB.php');
require('class.XRDB.php');
require('simpleXLSX.php');

$Y=new XRDB();
$X=new DB();
//---------------------------------------------------------------------------------------------------------------------
//TEMPLATE FUNCTIONS
//---------------------------------------------------------------------------------------------------------------------
//
//
$message="";
$post=array();
$post['table_name']="nua_doc";
$post['action']="insert";
$post['file_type']=$_FILES["file"]["type"];
$post['file_name']=$_FILES["file"]["name"];
$key=$post['file_name'] . $post['file_type'];
$m=md5($key);
$d=substr($m,0,7);
$f=substr($m,7,7);
$dir="/var/www/docs/" . $d;
try {
	mkdir($dir);
}
catch(Exception $e) {
  echo 'Message: ' .$e->getMessage();
}
$target_file=$dir . "/" . $f;
if (move_uploaded_file($_FILES["file"]["tmp_name"], $target_file)) { 


if ( $xlsx = SimpleXLSX::parse($target_file) ) {
	  
    $i = 1;

    foreach ($xlsx->rows(0) as $elt) {
		$p=array();
		if ($i>1) {
			if ($elt[3]!='') {
			$sql="select * from nua_payment where reference_number = '" . $elt[2] . "'";
			$t=$X->sql($sql);
			if (sizeof($t)==0) {
			$post=array();
			$post['table_name']="nua_payment";
			$post['action']="insert";
			$post['company_id']=$elt[0];
			$post['reference_number']=$elt[2];
			$post['deposit_type']=strtoupper($elt[5]);
			if (strtoupper($elt[5])=='CREDIT')  {
				$post['bank']="CREDIT";
			} else {
			   $post['bank']="5/3";
			}
			$post['amount_received']=$elt[4];
			$post['amount_applied']=$elt[4];
			$post['payment_date']=substr($elt[3],5,2) . "/" . substr($elt[3],8,2) . "/" . substr($elt[3],0,4);
			$post['amount_remaining']=0;
			$sql="select * from nua_company_invoice where invoice_number='" . $elt[2] . "'";
			$b=$X->sql($sql);
                        if (sizeof($b)>0) {
                            $invoice=$b[0];
			    $post['amount_remaining']=floatval($elt[4])-$invoice['grand_total_float'];
			    if (floatval($post['amount_remaining'])<0) $post['amount_remaining']=0;
			    $post['applied_to_1']=$invoice['id'];
			    $post['amount_applied_1']=$elt[4];
			    $payment_id=$X->post($post);

                            $sql="update nua_company_invoice set paid = 'Yes',  final='Y', payment_amount = '" . $elt[4];
                            $sql.="', amount_left = 0  where id = " . $post['applied_to_1'];
                            $X->execute($sql);

                            $p=array();
                            $p['table_name']="nua_payment_detail";
                            $p['action']="insert";
                            $p['seq']=1;
                            $p['payment_id']=$payment_id;
                            $p['invoice_id']=$post['applied_to_1'];
                            $p['amount_applied']=$elt[4];
                            $X->post($p);
	     	            $message.="Payment: " . $elt[2] . " sucessful., ";
			} else {
	     	            $message.="	Invoice: " . $elt[2] . " not found., ";
                            $post['amount_remaining']=$elt[4];
			    $payment_id=$X->post($post);
			}
			} else {
	     	           $message.="Payment " . $elt[2] . " exists., ";
			}
			}
		  } // i > 10
          $i++;
        } // for each
  } else {
    echo SimpleXLSX::parseError();
  } // valid parse
}  // move uploaded file
$output=array();
$output['error_code']=0;
$output['message']=$message;

echo json_encode($output);
?>
