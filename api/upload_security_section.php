<?php

ini_set('memory_limit',-1);
ini_set('max_execution_time', 3000);
ini_set('display_errors',1);
ini_set('display_startup_errors',1);
ini_set('upload_max_filesize', '8M');
ini_set('memory_limit', -1);
ini_set('post_max_size', '8M');
ini_set('KeepAliveTimeout', '300');
header('Access-Control-Allow-Headers: Access-Control-Allow-Origin, Content-Type, Authorization');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET,PUT,POST,DELETE,PATCH,OPTIONS');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require('/var/www/classes/class.XRDB.php');
require('class.XRDB.php');
require('simpleXLSX.php');

$Y=new XRDB();
$X=new DB();


function process_enrollment($elt,$company_id) {

}

if (isset($_POST["section_id"])) { $post['section_id']=$_POST["section_id"]; }

$post=array();
$post['table_name']="psp_doc";
$post['action']="insert";
$post['file_type']=$_FILES["file"]["type"];
$post['file_name']=$_FILES["file"]["name"];

//$user_id=$_POST['user_id'];
$user_id=1;
$section_id=$_POST['section_id'];

$sql="select * from psp_survey_section where id = " . $section_id;
$b=$X->sql($sql);
$section=$b[0];
$template_id=$section['template_id'];

$sql="select * from psp_template where id = " . $template_id;
$b=$X->sql($sql);
$template=$b[0];


$id=$Y->doc($post);
$key="psp".$id;
$m=md5($key);
$d=substr($m,0,7);
$f=substr($m,7,7);
$dir="/var/www/docs/" . $d;
mkdir($dir);
$target_file=$dir . "/" . $f;
$target_file="/var/www/docs/test.xlsx";

if (move_uploaded_file($_FILES["file"]["tmp_name"], $target_file)) { 

if ( $xlsx = SimpleXLSX::parse($target_file) ) {
    $i = 1;
    $sql="insert into psp_survey_option_backup (id, create_timestamp, section_id, parent_id, option_text, option_type, ";
    $sql.=" option_group, option_order, third_person_text, help_text, validation_text, height, width, session_id) ";
    $sql.=" select id, create_timestamp, section_id, parent_id, option_text, option_type, option_group, option_order, ";
    $sql.=" third_person_text, help_text, validation_text, height, width, " . $id . " from ";
    $sql.=" psp_survey_option where section_id = " . $section_id;
    $X->execute($sql);

    $sql="delete from psp_survey_option where section_id = " . $section_id;
    $X->execute($sql);


    $option_id=intval($section_id)*1000;
    $order_id=0;
    $status="";

    $parents=array();
    
    $p=array();
    $p['tmp']=0;
    $p['act']=0;
    array_push($parents,$p);
    $current_option_group_radio=0;
    $current_option_group_checkbox=0;
    $current_group_type="";

       try {
	foreach ($xlsx->rows(0) as $elt) {
		if ($i>1) {
                      $option_id += 10;                  		       
		      $order_id += 10;
		      //
		      // Values in elt[1] are the fake parent id's
		      //
		      // $elt[0] = parent_id;
		      // $elt[1] = tmp_option_id;
		      // $elt[2] = option type
		      // $elt[3] = option text
		      //
		      $tmp_parent_id="";
		      $tmp_parent_id=$elt[0];
		      $tmp_option_id=$elt[1];
		      $option_type=strtoupper($elt[2]);
		      $option_text=$elt[3];
		      $third=$elt[4];
		      $help=$elt[5];

		      if ($tmp_option_id!="") { 
		          $p=array();
		          $p['tmp']=$tmp_option_id;
		          $p['act']=$option_id;	  
                          array_push($parents,$p);
		      };

                      $post=array();
		      $post['table_name']="psp_survey_option";
		      $post['action']="insert";
		      $post['user_id']=$user_id;
		      $post['section_id']=$section_id;
		      $post['template_id']=$template_id;
		      $post['option_id']=$option_id;
                      $post['option_order']=$option_id;
		      if ($tmp_parent_id==0) {
			      $post['parent_id']=0;
			      $current_parent=0;
		      } else {
                            foreach($parents as $pp) {
                               if ($pp['tmp']==$tmp_parent_id) $post['parent_id']=$pp['act'];
                            }			     
		      }
                      $found='N';
		      if ($option_type=="RADIO GROUP") {
                          $current_option_group_radio=$option_id;
                          $post['option_group']=0;
		      }
		      if ($option_type=="CHECKBOX GROUP") {
                          $current_option_group_checkbox=$option_id;
                          $post['option_group']=0;
		      }
		      if ($option_type=="SUBTITLE") {
                          $post['option_group']=0;
		      }
		      if ($option_type=="TITLE") {
                          $post['option_group']=0;
		      }
		      if ($option_type=="TEXT") {
                          $post['option_group']=0;
		      }
                      if ($option_type=="RADIO BUTTON") {
                         $post['option_group']=$current_option_group_radio;
                      }
                      if ($option_type=="CHECKBOX") {
                         $post['option_group']=$current_option_group_checkbox;
                      }
                      $post['option_type']=$elt[2];
                      $post['option_text']=$elt[3];
                      $post['third_person_text']=$elt[4];
                      $post['help_text']=$elt[5];
                      $X->post($post);
		} // I>2 
	    $i++;
	}  // FOREACH
	}  // TRY
          catch (Exception $e) {

	} //CATCH
  }

$output=array();
$output['error_code']="0";
echo json_encode($output);
}
