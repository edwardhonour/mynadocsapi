<?php

ini_set('memory_limit',-1);
ini_set('max_execution_time', 3000);
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
require('class.enrollment.php');

$Y=new XRDB();
$X=new DB();
$E=new ENROLLMENT();

//---------------------------------------------------------------------------------------------------------------------
//TEMPLATE FUNCTIONS
//---------------------------------------------------------------------------------------------------------------------
//
function makeMonthList($low, $high) {
	$monthlist=array();
	array_push($monthlist,"2021-01");
	array_push($monthlist,"2021-02");
	array_push($monthlist,"2021-03");
	array_push($monthlist,"2021-04");
	array_push($monthlist,"2021-05");
	array_push($monthlist,"2021-06");
	array_push($monthlist,"2021-07");
	array_push($monthlist,"2021-08");
	array_push($monthlist,"2021-09");
	array_push($monthlist,"2021-10");
	array_push($monthlist,"2021-11");
	array_push($monthlist,"2021-12");
	array_push($monthlist,"2022-01");
	array_push($monthlist,"2022-02");
	array_push($monthlist,"2022-03");
	array_push($monthlist,"2022-04");
	array_push($monthlist,"2022-07");
	array_push($monthlist,"2022-08");
	array_push($monthlist,"2022-09");
	array_push($monthlist,"2022-10");
	array_push($monthlist,"2022-11");
	array_push($monthlist,"2022-12");	
	array_push($monthlist,"2023-01");
	array_push($monthlist,"2023-02");
	array_push($monthlist,"2023-03");
	array_push($monthlist,"2023-04");
	array_push($monthlist,"2023-05");
	array_push($monthlist,"2023-06");
	array_push($monthlist,"2023-07");
	array_push($monthlist,"2023-08");
	array_push($monthlist,"2023-09");
	array_push($monthlist,"2023-10");
	array_push($monthlist,"2023-11");
	array_push($monthlist,"2023-12");	

	$start='N';
	$output=array();
	foreach($monthlist as $m) {
//		echo $m;
//		echo "|";
//		echo $low;
//		echo ",";
		//		
		if ($m==$low) { $start='Y'; }
		if ($start=='Y') {
			array_push($output,$m);
			if ($m==$high) { $start='N'; }
		}
	}
	return $output;
}

function makeDeleteList($low, $high) {
	$monthlist=array();
	array_push($monthlist,"2021-01");
	array_push($monthlist,"2021-02");
	array_push($monthlist,"2021-03");
	array_push($monthlist,"2021-04");
	array_push($monthlist,"2021-05");
	array_push($monthlist,"2021-06");
	array_push($monthlist,"2021-07");
	array_push($monthlist,"2021-08");
	array_push($monthlist,"2021-09");
	array_push($monthlist,"2021-10");
	array_push($monthlist,"2021-11");
	array_push($monthlist,"2021-12");
	array_push($monthlist,"2022-01");
	array_push($monthlist,"2022-02");
	array_push($monthlist,"2022-03");
	array_push($monthlist,"2022-04");
	array_push($monthlist,"2022-07");
	array_push($monthlist,"2022-08");
	array_push($monthlist,"2022-09");
	array_push($monthlist,"2022-10");
	array_push($monthlist,"2022-11");
	array_push($monthlist,"2022-12");	
	array_push($monthlist,"2023-01");
	array_push($monthlist,"2023-02");
	array_push($monthlist,"2023-03");
	array_push($monthlist,"2023-04");
	array_push($monthlist,"2023-05");
	array_push($monthlist,"2023-06");
	array_push($monthlist,"2023-07");
	array_push($monthlist,"2023-08");
	array_push($monthlist,"2023-09");
	array_push($monthlist,"2023-10");
	array_push($monthlist,"2023-11");
	array_push($monthlist,"2023-12");	

	$start='N';
	$output=array();
	foreach($monthlist as $m) {
		if ($m==$low) { $start='Y'; }
		if ($start=='N') {
			array_push($output,$m);
		}
		if ($m==$high) { $start='N'; }
	}
	return $output;
}

function convert_month($m) {
	$month=substr($m,0,2);
	$year=substr($m,6,4);
	$month_id=$year . "-" . $month;
	return $month_id;
}

function process_enrollment($elt,$company_id) {
/*
	print_r($elt);

       $X=new DB();
       $p=array();
       $p['table_name']="nua_monthly_member_census";
       $p['action']="insert";			
       $p['month_id']=$elt[0];
       // $p['employee_code']=$elt[1];
       // $p['dependent_code']=substr($elt[2],0,10);
	  $p['ssn']=$elt[3];
	  $p['company_id']=$company_id;
	  $p['company_name']=$company_name;
	  $p['last_name']=strtoupper($elt[4]);	
		   $p['first_name']=strtoupper($elt[5]);
		   $p['middle_initial']=strtoupper($elt[5]);	
		   $p['middle_name']=strtoupper($elt[5]);			   
		   $p['relationship']=$elt[7];
		   $p['dob']=substr($elt[7],0,10);		
		   $p['gender']=$elt[9];
		   $p['marital_status']=substr($elt[10],0,1);
		   $p['address']=$elt[12];	
		   $p['address2']=$elt[13];
		   $p['city']=$elt[14];			   
		   $p['state']=$elt[15];
		   $p['zip']=$elt[16];
		   $p['email']=$elt[17];
		   $p['phone']=$elt[18];
		   $p['hire_date']=substr($elt[19],0,10);
		   
		   if (strtoupper($p['relationship'])=="EMPLOYEE") {
			   
			   //-- Find out if Employee is already in company.
			   //-- 
			   
			   $sql="select * from nua_employee where company_id = " . $company_id . " and last_name = '" . str_replace("'","''",$p['last_name']) . "' AND ";
			   $sql.=" first_name = '" . str_replace("'","''",$p['first_name']) . "' AND ";
                           $sql.=" middle_name = '" . str_replace("'","''",$p['middle_name']) . "'";
			   
                           $e=$X->sql($sql);
			   $employee_id=0;
			   $employee_code='n0';
			   if (sizeof($e)>0) {
				$employee_id = $e[0]['id'];
				$employee_code = $e[0]['employee_code'];				   
			   } else {
				   //--
				   //-- Check if employee already exists in census.
				   //--
				   $sql="select * from nua_monthly_member_census where last_name = '" . str_replace("'","''",$p['last_name']) . "' AND ";
    			           $sql.=" first_name = '" . str_replace("'","''",$p['first_name']) . "' and dependent_code = '' order by month_id desc";
				   $e=$X->sql($sql);
                                   if (sizeof($e)>0) {
					   $employee_id = $e[0]['employee_id'];
					   $employee_code = $e[0]['employee_code'];
					   $census_company_id = $e[0]['company_id'];
					   if (($census_company_id != $company_id)&&($company_id != 0)) {
                                                   $sql="select company_name from nua_company where id = " . $company_id;
						   $c=$X->sql($sql);
						   $sql="update nua_census set alt_company_id = " . $company_id . " where last_name = '" . str_replace("'","''",$p['last_name']) . "' AND ";
                         			   $sql.=" first_name = '" . str_replace("'","''",$p['first_name']) . "'";
						   $X->execute($sql);
				                   $sql="update nua_census set alt_company_id = " . $c[0]['company_name'] . " where last_name = '" . str_replace("'","''",$p['last_name']) . "' AND ";
            			                   $sql.=" first_name = '" . str_replace("'","''",$p['first_name']) . "'";
						   $X->execute($sql);						   
						   
						   $sql="update nua_monthly_member_census set company_id = " . $company_id . " where employee_id = " . $employee_id;
						   $X->execute($sql);

						   $sql="update nua_monthly_member_additions set company_id = " . $company_id . " where employee_id = " . $employee_id;
						   $X->execute($sql);

						   $sql="update nua_monthly_member_terminations set company_id = " . $company_id . " where employee_id = " . $employee_id;
						   $X->execute($sql);
					   }
				   }
			   }
			   $emp=array();
			   $emp['table_name']="nua_employee";
			   $emp['action']="insert";
                           $emp['company_id']=$company_id;
			   $emp['last_name']=strtoupper($elt[4]);	
			   $emp['first_name']=strtoupper($elt[5]);
			   $emp['middle_name']=strtoupper($elt[6]);			   
			   $emp['social_security_number']=$elt[3];
			   
			   $emp['date_of_birth']=substr($elt[7],0,10);		
   		           $emp['gender']=$elt[9];
		           $emp['marital_status']=substr($elt[10],0,1);
			   $emp['address']=$elt[11];	
		       $emp['suite']=$elt[12];
		       $emp['city']=$elt[13];			   
		       $emp['state']=$elt[14];
		       $emp['zip']=$elt[15];
		       $emp['email']=$elt[16];
		       $emp['phone']=$elt[17];
		       $emp['date_hired']=substr($elt[18],0,10);
		       $emp['work_status']=$elt[19];		   
		       $emp['eff_dt']=substr($elt[20],0,10);	
		       $emp['medical_eff_dt']=substr($elt[20],0,10);		 
		       $emp['medical_plan']=$elt[21];		   
		       $emp['medical_plan_level']=$elt[22];
		       $emp['dental_plan']=$elt[23];
		       $emp['dental_plan_level']=$elt[24];
		       $emp['vision_plan']=$elt[25];
		       $emp['vision_plan_level']=$elt[26];
		       $sql="select * from nua_company_plan where apa_code = '" . $elt[22] . "' and company_id = " . $company_id;
		       $r=$X->sql($sql);
		       if (sizeof($r)>0) {
			      $emp['medical_plan']=$r[0]['plan_code'];
			      $emp['medical_plan_level']=$elt[23];
                              $ee_price=$r[0]['ee_price'];
                              $ees_price=$r[0]['ees_price'];
                              $eec_price=$r[0]['eec_price'];
                              $fam_price=$r[0]['fam_price'];
		       } else {
                              $emp['medical_plan']="";
			      $emp['medical_plan_level']="";
		       }
		       $sql="select * from nua_company_plan where apa_code = 'GUARDHIGH' and company_id = " . $company_id;
		       $r=$X->sql($sql);
		       if (sizeof($r)>0) {
                              $dee_price=$r[0]['ee_price'];
                              $dees_price=$r[0]['ees_price'];
                              $deec_price=$r[0]['eec_price'];
                              $dfam_price=$r[0]['fam_price'];
		       } else {
                              $dee_price="0.00";
                              $dees_price="0.00";
                              $deec_price="0.00";
                              $dfam_price="0.00";
		       }
		       $sql="select * from nua_company_plan where apa_code = 'VSP' and company_id = " . $company_id;
		       $r=$X->sql($sql);
		       if (sizeof($r)>0) {
                              $vee_price=$r[0]['ee_price'];
                              $vees_price=$r[0]['ees_price'];
                              $veec_price=$r[0]['eec_price'];
                              $vfam_price=$r[0]['fam_price'];
		       } else {
                              $vee_price="0.00";
                              $vees_price="0.00";
                              $veec_price="0.00";
                              $vfam_price="0.00";
		       }


		   if ($employee_id!=0) {
			   $emp['id']=$employee_id; 
		            $X->post($emp);
		   } else {
			   $employee_id = $X->post($emp);   
			   $sql="update nua_employee set employee_code = 'n" . $employee_id . "' where id = " . $employee_id;
			   $X->execute($sql);
		   }
		   $month_id=substr($elt[21],0,7);

		   $date=date_create();
		   $current_month=date_format($date,'Y-m');
		   if ($current_month=="2021-10") $current_month_id="2021-11";
		   if ($current_month=="2021-11") $current_month_id="2021-12";
		   if ($current_month=="2021-12") $current_month_id="2022-01";
		   if ($current_month=="2022-01") $current_month_id="2022-02";
		   if ($current_month=="2022-02") $current_month_id="2022-03";
		   if ($current_month=="2022-03") $current_month_id="2022-04";
		   if ($current_month=="2022-04") $current_month_id="2022-05";
		   if ($current_month=="2022-05") $current_month_id="2022-06";
		   if ($current_month=="2022-06") $current_month_id="2022-07";
		   if ($current_month=="2022-07") $current_month_id="2022-08";
		   if ($current_month=="2022-08") $current_month_id="2022-09";

		   $months=makeMonthList($month_id,$current_month_id);
		   $deletes=makeDeleteList($month_id,$current_month_id);
		   //--
		   //
		   //-- Insert the employee into each months census from the effective date to the 
		   //-- current billing month.
		   //--
		   foreach ($deletes as $m) {

                                   $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                   $sql.=" and month_id = '" . $m . "' ";
                                   $X->execute($sql);

	           }
		   foreach ($months as $m) {
			    $p['month_id']=$m;
			    $p['eff_dt']=substr($elt[21],0,10);		 
                            $p['employee_id']=$employee_id;
                            $p['employee_code']='n'.$employee_id;
                            $p['company_name']=$company_name;
                            $p['company_id']=$company_id;
                            if ($emp['medical_plan']!="") {
				   $p['client_plan']=$emp['medical_plan'];
                                   $p['coverage_level']=$emp['medical_plan_level'];
				   if ($p['coverage_level']=="EE"||$p['coverage_level']=="SI") { $p['coverage_price']=$ee_price; }
				   if ($p['coverage_level']=="ES"||$p['coverage_level']=="EES") { $p['coverage_price']=$ees_price; }
				   if ($p['coverage_level']=="EC"||$p['coverage_level']=="EEC") { $p['coverage_price']=$eec_price; }
				   if ($p['coverage_level']=="FAM"||$p['coverage_level']=="FA") { $p['coverage_price']=$fam_price; }
				   $p['apa_plan']=$elt[22];
				   $p['plan_type']="*MEDICAL*";
                                   $sql="SELECT id from nua_monthly_member_census where employee_id = " . $employee_id;
                                   $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                   $sql.=" and client_plan = '" . $p['client_plan'] . "' and dependent_code = ''";
                                   $dd=$X->sql($sql);
                                   if (sizeof($dd)>0) {
                                        $p['id']=$dd[0]['id'];
                                   } 
				   $X->post($p);
				}
                                if ($emp['dental_plan']!="") {
				   $p['client_plan']=$emp['dental_plan'];
                                   $p['coverage_level']=$emp['dental_plan_level'];
				   if ($p['coverage_level']=="EE"||$p['coverage_level']=="SI") { $p['coverage_price']=$dee_price; }
				   if ($p['coverage_level']=="ES"||$p['coverage_level']=="EES") { $p['coverage_price']=$dees_price; }
				   if ($p['coverage_level']=="EC"||$p['coverage_level']=="EEC") { $p['coverage_price']=$deec_price; }
				   if ($p['coverage_level']=="FAM"||$p['coverage_level']=="FA") { $p['coverage_price']=$dfam_price; }
				   $p['apa_plan']="GUARDHIGH";
				   $p['plan_type']="*DENTAL*";
                                   $sql="SELECT id from nua_monthly_member_census where employee_id = " . $employee_id;
                                   $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                   $sql.=" and client_plan = '" . $p['client_plan'] . "' and dependent_code = ''";
                                   $dd=$X->sql($sql);
                                   if (sizeof($dd)>0) {
                                        $p['id']=$dd[0]['id'];
                                   } 
				   $X->post($p);
				}				
                                if ($emp['vision_plan']!="") {
				   $p['client_plan']=$emp['vision_plan'];
                                   $p['coverage_level']=$emp['vision_plan_level'];
				   $p['apa_plan']="VSP";
				   if ($p['coverage_level']=="EE"||$p['coverage_level']=="SI") { $p['coverage_price']=$vee_price; }
				   if ($p['coverage_level']=="ES"||$p['coverage_level']=="EES") { $p['coverage_price']=$vees_price; }
				   if ($p['coverage_level']=="EC"||$p['coverage_level']=="EEC") { $p['coverage_price']=$veec_price; }
				   if ($p['coverage_level']=="FAM"||$p['coverage_level']=="FA") { $p['coverage_price']=$vfam_price; }
				   $p['plan_type']="*VISION*";
                                   $sql="SELECT id from nua_monthly_member_census where employee_id = " . $employee_id;
                                   $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                   $sql.=" and client_plan = '" . $p['client_plan'] . "' and dependent_code = ''";
                                   $dd=$X->sql($sql);
                                   if (sizeof($dd)>0) {
                                        $p['id']=$dd[0]['id'];
                                   } 
				   $X->post($p);
				}
		   }
	   
		   //-- Update nua_monthly_member_census
		   
		   } else {
			   //--
			   //-- Process Dependent.
			   //--
			   
			    //--
				//-- Get the Employee Record
				//--
				
				$sql="select id, employee_code from nua_employee where social_security_number = '" . $elt[3] . "'";
		     	        $emps=$X->sql($sql);
				$employee_id=$emps[0]['id'];
				$employee_code=$emps[0]['employee_code'];

				$dep=array();
				$dep['table_name']="nua_employee_dependent";
			        $dep['action']="insert";
				//--
				//-- Find Existing Dependent
				//--
   			       $sql="select * from nua_employee_dependent where employee_id = " . $employee_id . " and last_name = '" . str_replace("'","''",$p['last_name']) . "' AND ";
			       $sql.=" first_name = '" . str_replace("'","''",$p['first_name']) . "' AND ";
                               $sql.=" middle_name = '" . str_replace("'","''",$p['middle_name']) . "'";
                               $e=$X->sql($sql);
			   
			   $dep['employee_id']=$employee_id;
			   $dep['employee_code']=$employee_code;
			   $dep['company_id']=$company_id;
			   $dep['first_name']=$p['first_name'];
			   $dep['last_name']=$p['last_name'];
			   $dep['middle_name']=$p['middle_name'];
		           $dep['gender']=$elt[10];
		           $dep['marital_status']=substr($elt[11],0,1);
			   $dep['address']=$elt[12];	
		           $dep['suite']=$elt[13];
		           $dep['city']=$elt[14];			   
		           $dep['state']=$elt[15];
		           $dep['zip']=$elt[16];
		           $dep['email']=$elt[17];
		           $dep['phone']=$elt[18];
			   $dep['date_of_birth']=$elt[9];		
			   $dep['social_security_number']=$elt[4];
			   $dep['relation_to_insured']=$elt[8];
			   $dep['relationship']=$elt[8];
                           if (strtoupper($elt[8])=='WIFE') $dep['relation_type']="02";
                           if (strtoupper($elt[8])=='HUSBAND') $dep['relation_type']="02";
                           if (strtoupper($elt[8])=='SON') $dep['relation_type']="03";
                           if (strtoupper($elt[8])=='DAUGHTER') $dep['relation_type']="03";
			   if (sizeof($e)>0) {
				   $dep['id']=$e[0]['id'];   
				   $dep['dependent_id']=$e[0]['dependent_id'];   
			   } else {
                                   $dep['dependent_id']="";
                           }
                           if ($dep['dependent_id']=="") {
                                $sql="select count(*) as c from nua_employee_dependent where employee_id = " . $employee_id;
                                $ee=$X->sql($sql);
                                $did=$ee[0]['c']+1; 
                                $dep['dependent_id']="n".$employee_id.".".$did;
                          }
                          $X->post($dep);
                          $sql="select * from nua_monthly_member_census where employee_id = " . $employee_id . " and dependent_code =''";
                          $cc=$X->sql($sql);
                          foreach($cc as $c) {
                              $depc=$c;
			      $depc['table_name']="nua_monthly_member_census";
                              $depc['action']="insert";
                              unset($depc['create_timestamp']);  
                              unset($depc['id']);  
                              $depc['dependent_code']=$dep['dependent_id'];
			      $depc['first_name']=$dep['first_name'];
			      $depc['last_name']=$dep['last_name']=$p['last_name'];
         		      $depc['middle_initial']=$dep['middle_name'];
                              $depc['gender']=$dep['gender'];
                              $depc['marital_status']=$dep['marital_status'];
		              $depc['email']=$dep['email'];
		              $depc['phone']=$dep['phone']=$elt[18];
			      $depc['dob']=$dep['date_of_birth'];		
			      $depc['ssn']=$dep['social_security_number']=$elt[4];
			      $depc['relation_type']=$dep['relation_type'];
			      $depc['relationship']=$dep['relationship'];
			      $depc['relation_to_insured']=$dep['relation_to_insured'];
                              $sql="select * from nua_monthly_member_census where employee_id = " . $employee_id;
                              $sql.=" and dependent_code = '" . $depc['dependent_code'] . "' ";
                              $sql.=" and month_id = '" . $depc['month_id'] . "'";
                              $ff=$X->sql($sql);
                              if (sizeof($ff)>0) {
                                  $depc['id']=$ff[0]['id'];
			      }
                              $X->post($depc);
                                   $sql="SELECT id from nua_monthly_member_census where employee_id = " . $employee_id;
                                   $sql.=" and month_id = '" . $depc['month_id'] . "' ";
                                   $sql.=" and client_plan = '" . $depc['client_plan'] . "' and dependent_code = '" . $depc['dependent_code'] . "'";
                                   $dd=$X->sql($sql);
                                   if (sizeof($dd)>0) {
                                        $depc['id']=$dd[0]['id'];
                                   } 
                                   $X->post($depc);
                          }
		   }	   
 */
}
function process_census($elt,$company_id, $org_id, $company_name) {
           $X=new DB();
	   $p=array();
	   $p['table_name']="nua_preenrollment_census";
	   $p['action']="insert";			
	   $p['company_id']=$company_id;
	   $p['social_security_number']=$elt[0];
	   $p['dependent_social_security_number']=$elt[1];
	   $p['last_name']=strtoupper($elt[2]);	
	   $p['first_name']=strtoupper($elt[3]);
	   $p['middle_name']=strtoupper($elt[4]);			   
	   $p['relationship']=$elt[5];
	   $p['date_of_birth']=substr($elt[6],0,10);		
	   $p['gender']=$elt[7];
	   $p['marital_status']=substr($elt[8],0,1);
           $sql="select * from nua_preenrollment_census where social_security_number = '" . $p['social_security_number'] . "' and dependent_social_security_number = '";
           $sql.=$p['dependent_social_security_number'] . "'";
           $gg=$X->sql($sql);
           if (sizeof($gg)>0) {
               $p['id']=$gg[0]['id'];
	   }
           if (isset($elt[9])) {
             if (strtoupper($elt[9])=="DELETE") {
                $sql="delete from nua_preenrollment_census where social_security_number = '" . $p['social_security_number'] . "' and dependent_social_security_number = '";
                $sql.=$p['dependent_social_security_number'] . "'";
                $X->execute($sql);
	     } else {
                $X->post($p);
             }
	  } else {
                $X->post($p);
          }
}


function process_quoted_plans($elt,$company_id, $org_id, $company_name) {

           $X=new DB();
	   $p=array();
	   $p['table_name']="nua_quoted_plan";
	   $p['action']="insert";			
	   $p['company_id']=$company_id;
	   $p['plan_code']=strtoupper($elt[0]);
	   $p['APA_CODE']=strtoupper($elt[1]);
	   $p['ee_price']=$elt[2];	
	   $p['ees_price']=$elt[3];
	   $p['eec_price']=$elt[4];			   
	   $p['fam_price']=$elt[5];
	   $p['accepted']=$elt[6];
           if ($p['accepted']=="") $p['accepted']="N";
           $sql="select * from nua_quoted_plan where company_id = " . $company_id . " and plan_code = '" . $p['plan_code'] . "'";
           $gg=$X->sql($sql);
           if (sizeof($gg)>0) {
               $p['id']=$gg[0]['id'];
	   }
           if (isset($elt[7])) {
             if (strtoupper($elt[7])=="DELETE") {
	        $p['accepted']="NO";
                $sql="delete from nua_quoted_plan where company_id = " . $company_id . " and plan_code = '" . $p['plan_code'] . "'";
                $X->execute($sql);
	     } else {
                $X->post($p);
             }
	  } else {
                $X->post($p);
          }
          if (strtoupper($p['accepted'])=="Y"||strtoupper($p['accepted']=="YES")) {
             $p['table_name']="nua_company_plan";
             $p['id']="";
             $sql="select * from nua_company_plan where company_id = " . $company_id . " and plan_code = '" . $p['plan_code'] . "'";
             if (sizeof($gg)>0) { $p['id']=$gg[0]['id']; }
             $X->post($p);
          }
}

function process_company_plans($elt,$company_id, $org_id, $company_name) {

           $X=new DB();
	   $p=array();
	   $p['table_name']="nua_company_plan";
	   $p['action']="insert";			
	   $p['company_id']=$company_id;
	   $p['plan_code']=strtoupper($elt[0]);
	   $p['APA_CODE']=strtoupper($elt[1]);
	   $p['ee_price']=$elt[2];	
	   $p['ees_price']=$elt[3];
	   $p['eec_price']=$elt[4];			   
	   $p['fam_price']=$elt[5];
	   $p['plan_type']=$elt[6];
           $sql="select * from nua_company_plan where company_id = " . $company_id . " and plan_code = '" . $p['plan_code'] . "'";
           $gg=$X->sql($sql);
           if (sizeof($gg)>0) {
               $p['id']=$gg[0]['id'];
	   }
           if (isset($elt[7])) {
             if (strtoupper($elt[7])=="DELETE") {
                $sql="delete from nua_company_plan where company_id = " . $company_id . " and plan_code = '" . $p['plan_code'] . "'";
                $X->execute($sql);
	     } else {
                $X->post($p);
             }
	  } else {
                $X->post($p);
          }

}



//---------------------------------------------------------------------------------------------------------------------
//END TEMPLATE FUNCTIONS 
//---------------------------------------------------------------------------------------------------------------------
//

$post=array();
$post['table_name']="nua_doc";
$post['action']="insert";
//echo json_encode($_FILES);
$post['file_type']=$_FILES["file"]["type"];
$post['file_name']=$_FILES["file"]["name"];
if (isset($_POST["company_id"])) { $post['company_id']=$_POST["company_id"]; }
if (isset($_POST["employee_id"])) { $post['employee_id']=$_POST["employee_id"]; }
if (isset($_POST["dsc"])) { $post['dsc']=$_POST["dsc"]; }
if (isset($_POST["doc_title"])) { $post['doc_title']=$_POST["doc_title"]; }
$user_id=$_POST['user_id'];

$post['user_id']=$_POST["user_id"];
$id=$Y->doc($post);
$key="nuaxess".$id;
$m=md5($key);
$d=substr($m,0,7);
$f=substr($m,7,7);
$dir="/var/www/docs/" . $d;
mkdir($dir);
$target_file=$dir . "/" . $f;
if (move_uploaded_file($_FILES["file"]["tmp_name"], $target_file)) { 

//------------------------------------------------------------------------------------------------------------------
// Template Code
// -----------------------------------------------------------------------------------------------------------------

$sql="select * from nua_user where id = " . $user_id;
$u=$X->sql($sql);
$user=$u[0]; 

$Company_A=array();  
$Company_B=array();
$Company_C=array();
$Company_D=array();
$Company_E=array();
$Company_F=array();

if ( $xlsx = SimpleXLSX::parse($target_file) ) {
	  
    $i = 1;

	//--
	//-- Company/Prospect
	//--
	
	//--
	//-- We are reading company data from sheet 0.  The company name must match MyNuaxess exactly for 
	//-- existing companies.
	//
	
        foreach ($xlsx->rows(0) as $elt) {
		array_push($Company_A,$elt[0]);
		array_push($Company_B,$elt[1]);
		array_push($Company_C,$elt[2]);
		array_push($Company_D,$elt[3]);
		array_push($Company_E,$elt[4]);
		//array_push($Company_F,$elt[5]);     
                $i++;
        }
	$new_company=$Company_B[1];

        $post=array();
	$post['table_name']="nua_company";
	$post['action']="insert";

        //-- Company Data
	
	$raw_id=$Company_B[2];
        $company_id=0;
	$post['company_name']=strtoupper($Company_B[3]);
	$company_name=$post['company_name'];
	$post['broker_name']=$Company_B[4];
	$post['broker_email']=strtolower($Company_B[5]);
	$post['company_type']=$Company_B[6];
	$post['tax_id']=$Company_B[7];
	$post['contact_name']=$Company_B[8];
	$post['contact_phone']=$Company_B[9];
	$post['contact_email']=$Company_B[10];
	$post['employee_count']=$Company_B[11];
	$post['medical']=$Company_B[13];
	$post['dental']=$Company_B[14];
	$post['vision']=$Company_B[15];
	$post['current_provider']=$Company_B[17];
	
	//-- Company Address
        $post['address']=strtoupper($Company_E[2]);
	$post['suite']=strtoupper($Company_E[3]);
	$post['city']=strtoupper($Company_E[4]);
	$post['state']=strtoupper($Company_E[5]);
        $post['zip']=$Company_E[6];
	$post['website']=strtolower($Company_E[7]);

        //-- Billing
        $post['invoice_company_name']=strtoupper($Company_E[10]);
        $post['invoice_mailing_address']=strtoupper($Company_E[11]);
	$post['invoice_suite']=strtoupper($Company_E[12]);
	$post['invoice_city']=strtoupper($Company_E[13]);
	$post['invoice_state']=strtoupper($Company_E[14]);
        $post['invoice_zip']=$Company_E[15];
	$post['billing_contact_name']=strtolower($Company_E[17]);
	$post['billing_contact_phone']=strtolower($Company_E[18]);
	$post['billing_contact_email']=strtolower($Company_E[19]);	
	$post['billing_contact_email2']=strtolower($Company_E[20]);	
	$post['billing_contact_email3']=strtolower($Company_E[21]);	
	$post['billing_contact_email4']=strtolower($Company_E[22]);	
	$post['billing_contact_email5']=strtolower($Company_E[23]);		

        //--
        //-- Get the company id if it exists.
	//--
	
        if ($raw_id!="(NuAxess Use Only)"&&$raw_id!="") {
                $company_id=$raw_id;
         	$sql="select * from nua_company where id = " . $raw_id;
	        $b=$X->sql($sql);
	        if (sizeof($b)>0) {
   		    $post['id']=$b[0]['id'];
		    $company_id=$b[0]['id'];
		    $company_name=$b[0]['company_name'];
		} else {
                    $company_id = 0;
                }
	} else {
           $company_id = 0;
        }

        //--
	//-- Find the broker and the broker's organization by email address.
	//--
	
	$sql="select * from nua_broker where email = '" . strtolower($post['broker_email']) . "'";
	$b=$X->sql($sql);
	if (sizeof($b)>0) {
    	    $post['org_id']=$b[0]['org_id'];	
	    $post['created_by']=$b[0]['id'];	
            $post['broker_id']=$b[0]['id'];	
            $post['broker_name']=$b[0]['first_name'] . ' ' . $b[0]['last_name'];	
	    $post['user_id']=$b[0]['id'];				
	} else {
	    $post['org_id']=0;	
	    $post['created_by']=0;	
            $post['broker_id']=0;
            $post['user_id']=0;			
	 }
     
         $company_id = $X->post($post);
         $sql="update nua_doc set company_id = " . $company_id . " where id = " . $id;
         $X->execute($sql);
 
//-----------------------------------------------------------------------------------------------------------------
// PRE ENROLLMENT CENSUS
// ---------------------------------------------------------------------------------------------------------------
	$sql="select * from nua_company where id = " . $company_id;
	$b=$X->sql($sql);
	if (sizeof($b)>0) {
		$company_id=$b[0]['id'];
		$org_id=$b[0]['org_id'];
		$company_name=$b[0]['company_name'];
	} else {
    	        $company_id = 0;	
	        $org_id = 0;	
	        $company_name='ERROR';
        }

	$i=1;
	try {
	foreach ($xlsx->rows(1) as $elt) {
            if ($i>2) process_census($elt,$company_id,$org_id,$company_name); 
	    $i++;
	} 
        } catch (Exception $e) {

        }


//-----------------------------------------------------------------------------------------------------------------
// QUOTED PLANS
// ---------------------------------------------------------------------------------------------------------------

	$sql="select * from nua_company where id = " . $company_id;
	$b=$X->sql($sql);
	if (sizeof($b)>0) {
		$company_id=$b[0]['id'];
		$org_id=$b[0]['org_id'];
		$company_name=$b[0]['company_name'];
	} else {
    	        $company_id = 0;	
	        $org_id = 0;	
	        $company_name='ERROR';
        }

	$i=1;
	foreach ($xlsx->rows(2) as $elt) {
            if ($i>2) process_quoted_plans($elt,$company_id,$org_id,$company_name); 
	    $i++;
	} 


//-----------------------------------------------------------------------------------------------------------------
// COMPANY PLANS
// ---------------------------------------------------------------------------------------------------------------

	$sql="select * from nua_company where id = " . $company_id;
	$b=$X->sql($sql);
	if (sizeof($b)>0) {
		$company_id=$b[0]['id'];
		$org_id=$b[0]['org_id'];
		$company_name=$b[0]['company_name'];
	} else {
    	        $company_id = 0;	
	        $org_id = 0;	
	        $company_name='ERROR';
        }

	$i=1;
       try {
	foreach ($xlsx->rows(3) as $elt) {
            if ($i>2) process_company_plans($elt,$company_id,$org_id,$company_name); 
	    $i++;
	} 
        } catch (Exception $e) {

        }


//-----------------------------------------------------------------------------------------------------------------
// ENROLLMENT
// ---------------------------------------------------------------------------------------------------------------

	$sql="select * from nua_company where id = " . $company_id;
	$b=$X->sql($sql);
	if (sizeof($b)>0) {
		$company_id=$b[0]['id'];
		$org_id=$b[0]['org_id'];
		$company_name=$b[0]['company_name'];
	} else {
    	        $company_id = 0;	
	        $org_id = 0;	
	        $company_name='ERROR';
        }

	$i=1;
       try {
	$kill="N";
	foreach ($xlsx->rows(4) as $elt) {
            if ($i==1) {
	       if ($elt[0]=="DELETE") {
		   $kill='Y';
                   $sql="select * from nua_monthly_member_census where company_id = " . $company_id;
                   $u=$X->sql($sql);
                   foreach($u as $v) {
                       $post=array();
                       $post['table_name']="backup_monthly_member_census";
                       $post['action']="insert";
                       $post['action']="insert";
                       foreach($v as $name => $value) {
                           if ($name=="create_timestamp") {
                           $date=date_create();
                              $post['create_timestamp']=date_format($date,'Y-m-d H:i:s');
                           } else {
                              $post[$name]=$value;
                          }
                       }
                       $X->postZ[$name]=$value;
                   }
                   $sql="delete from nua_monthly_member_census where company_id = " . $company_id;
                   $X->execute($sql);


                   $sql="select * from nua_monthly_member_additions where company_id = " . $company_id;
                   $u=$X->sql($sql);
                   foreach($u as $v) {
                       $post=array();
                       $post['table_name']="backup_monthly_member_additions";
                       $post['action']="insert";
                       $post['action']="insert";
                       foreach($v as $name => $value) {
                           if ($name=="create_timestamp") {
                           $date=date_create();
                              $post['create_timestamp']=date_format($date,'Y-m-d H:i:s');
                           } else {
                              $post[$name]=$value;
                          }
                       }
                       $X->postZ[$name]=$value;
                   }
                   $sql="delete from nua_monthly_member_additions where company_id = " . $company_id;
                   $X->execute($sql);


                   $sql="select * from nua_monthly_member_terminations where company_id = " . $company_id;
                   $u=$X->sql($sql);
                   foreach($u as $v) {
                       $post=array();
                       $post['table_name']="backup_monthly_member_terminations";
                       $post['action']="insert";
                       $post['action']="insert";
                       foreach($v as $name => $value) {
                           if ($name=="create_timestamp") {
                           $date=date_create();
                              $post['create_timestamp']=date_format($date,'Y-m-d H:i:s');
                           } else {
                              $post[$name]=$value;
                          }
                       }
                       $X->postZ[$name]=$value;
                   }
                   $sql="delete from nua_monthly_member_terminations where company_id = " . $company_id;
                   $X->execute($sql);

                   $sql="select * from nua_employee where company_id = " . $company_id;
                   $u=$X->sql($sql);
                   foreach($u as $v) {
                       $post=array();
                       $post['table_name']="backup_employee";
                       $post['action']="insert";
                       foreach($v as $name => $value) {
                           if ($name=="create_timestamp") {
                              $date=date_create();
                              $post['create_timestamp']=date_format($date,'Y-m-d H:i:s');
			   } else {
                              $post[$name]=$value;
                           }
                       }
                       $X->postZ[$name]=$value;
                   }

                   $sql="delete from nua_employee where company_id = " . $company_id;
		   $X->execute($sql);


                   $sql="select * from nua_employee_dependent where company_id = " . $company_id;
                   $u=$X->sql($sql);
                   foreach($u as $v) {
                       $post=array();
                       $post['table_name']="backup_employee_dependent";
		       $post['action']="insert";
                       foreach($v as $name => $value) {
                           if ($name=="create_timestamp") {
                              $date=date_create();
                              $post['create_timestamp']=date_format($date,'Y-m-d H:i:s');
			   } else {
                              $post[$name]=$value;
                           }
                       }
                       $X->postZ[$name]=$value;
                   }

                   $sql="delete from nua_employee_dependent where company_id = " . $company_id;
                   $X->execute($sql);
	       }
            }
	    if ($i>1) {
               $data=array();
               $data['min_month_id']=$elt[0];
               $data['max_month_id']=$elt[26];
	       if ($kill=="N") { 
                  $data['employee_code']=$elt[27];
                  $data['dependent_code']=$elt[28];
	       } else {
                  $data['employee_code']="";
                  $data['dependent_code']="";
	       }
               $data['social_security_number']=$elt[1];
               $data['dependent_social_security_number']=$elt[2];
	       $data['last_name']=$elt[3];
               $data['first_name']=$elt[4];
               $data['middle_initial']=$elt[5];
               $data['relationship']=$elt[6]; 
               $data['date_of_birth']=$elt[7];
               $data['gender']=$elt[8];
               $data['marital_status']=$elt[9];
               $data['address']=$elt[10];
	       $data['suite']=$elt[11];
               $data['city']=$elt[12];
               $data['state']=$elt[13];
               $data['zip']=$elt[14];
               $data['email']=$elt[15];
               $data['contact_phone']=$elt[16];
               $data['date_of_hire']=$elt[17];
               $data['work_status']=$elt[18];
               $data['eff_dt']=$elt[19];
               //$data['term_dt']=$elt[23];
               $data['medical_plan']=$elt[20];
               $data['medical_coverage_level']=$elt[21];
               //$data['medical_price']=$elt[26];
               $data['dental_plan']=$elt[22];
               $data['dental_coverage_level']=$elt[23];
               //$data['dental_price']=$elt[29];
               $data['vision_plan']=$elt[24];
               $data['vision_coverage_level']=$elt[25];
              // $data['vision_price']=$elt[32];
               $data['cid']="";
               $data['admin']=""; 

               if (strtoupper($data['admin'])=="DELETE") {
                  if ($data['employee_code']!=""&&$data['dependent_code']=="") {
                    if ($data['cid']=="") {
                        $sql="delete from nua_monthly_member_census where company_id = " . $company_id . " "; 
		        $sql.=" and employee_code = '" . $data['employee_code'] . "' and month_id = '" . $data['max_month_id'] . "'";
			$X->execute($sql);
                    } else {
                       $t=explode(",",$data['cid']);
                       foreach($t as $u) {
                           if ($u!="") {
                               $sql="delete from nua_monthly_member_census where id = " . $u;
			       $X->execute($sql);
                           }
                       }
                    }
                  }
                  if ($data['employee_code']!=""&&$data['dependent_code']!="") {
                    $sql="delete from nua_employee_dependent where id = " . $data['cid']; 
                    $X->execute($sql);
                  }
	       }

	       if (strtoupper($data['admin'])=="DELETE-ALL"||strtoupper($data['admin'])=="DELETEALL") {
                  if ($data['employee_code']!=""&&$data['dependent_code']=="") {
                    $sql="delete from nua_monthly_member_census where company_id = " . $company_id . " "; 
		    $sql.=" and employee_code = '" . $data['employee_code'] . "'";
                    $X->execute($sql);
                  }
                  if ($data['employee_code']!=""&&$data['dependent_code']!="") {
                    $sql="delete from nua_employee_dependent where company_id = " . $company_id . " "; 
		    $sql.=" and employee_code = '" . $data['employee_code'] . "' and dependent_id = '" . $data['dependent_code'] . "'";
                    $X->execute($sql);
                  }
	       }

               if (strtoupper($data['admin'])=="DELETE-EMPLOYEE") {
                  if ($data['employee_code']!=""&&$data['dependent_code']=="") {
                    if ($data['cid']=="") {
                        $sql="delete from nua_monthly_member_census where company_id = " . $company_id . " "; 
		        $sql.=" and employee_code = '" . $data['employee_code'] . "'";
                        $X->execute($sql);
                    } else {
                       $t=explode(",",$data['cid']);
                       foreach($t as $u) {
                           if ($u!="") {
                               $sql="delete from nua_monthly_member_census where id = " . $u; 
			       $X->execute($sql);
                           }
                       }
                    }

                    $sql="delete from nua_employee where company_id = " . $company_id . " "; 
		    $sql.=" and employee_code = '" . $data['employee_code'] . "'";
		    $X->execute($sql);

                    $sql="delete from nua_employee_dependent where company_id = " . $company_id . " "; 
		    $sql.=" and employee_code = '" . $data['employee_code'] . "'";
		    $X->execute($sql);

                  }
                  if ($data['employee_code']!=""&&$data['dependent_code']!="") {
                    $sql="delete from nua_employee_dependent where company_id = " . $company_id . " "; 
		    $sql.=" and employee_code = '" . $data['employee_code'] . "' and dependent_id = '" . $data['dependent_code'] . "'";
                    $X->execute($sql);
                  }
	       }

       if (strtoupper($data['admin'])!="DELETE"&&strtoupper($data['admin'])!="DELETEALL"&&strtoupper($data['admin'])!="DELETE-ALL"&&strtoupper($data['admin'])!="DELETE-EMPLOYEE") {
	    
	  
                   $E->process_enrollment($data,$company_id); 
               }
	    }
	    $i++;
	} 
        } catch (Exception $e) {

        }
}


//-----------------------------------------------------------------------------------------------------------------
// End Template Code
// ---------------------------------------------------------------------------------------------------------------
         $filename="";
         $filename=$E->getQuoteTemplate($company_id,"2022-05","F");
	 $output=array();
	 $output['error_code']=0;
	 $output['company_id']=$company_id;
	 $output['company_name']=$company_name;
	 $output['filename']=$filename;
	 $output['error_message']="";
	 echo json_encode($output);

} else {
         $output=array();
	 $output['error_code']=1;
         $output['error_message']="An Error has Occured";
	 echo json_encode($output);
}

?>
