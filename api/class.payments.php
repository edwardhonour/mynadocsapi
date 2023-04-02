<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_once('/var/www/classes/class.XRDB.php');
require_once('class.XRDB.php');


class ENROLLMENT {
        protected $X;
        protected $demo;
        protected $current_month_id;
        protected $next_month_id;
        protected $last_month_id;

    function __construct() {
        $this->X=new DB();
        $this->demo='N';
        $date=date_create();
        $this->current_month_id=date_format($date,'Y-m');
        if ($this->current_month_id=="2022-05") { $this->last_month_id="2022-04"; $this->next_month_id="2022-06"; }
        if ($this->current_month_id=="2022-06") { $this->last_month_id="2022-05"; $this->next_month_id="2022-07"; }
        if ($this->current_month_id=="2022-07") { $this->last_month_id="2022-06"; $this->next_month_id="2022-08"; }
        if ($this->current_month_id=="2022-08") { $this->last_month_id="2022-07"; $this->next_month_id="2022-09"; }
        if ($this->current_month_id=="2022-09") { $this->last_month_id="2022-08"; $this->next_month_id="2022-10"; }
        if ($this->current_month_id=="2022-10") { $this->last_month_id="2022-09"; $this->next_month_id="2022-11"; }

    }

function getMonthId() {
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
      if ($current_month=="2022-09") $current_month_id="2022-10";
      if ($current_month=="2022-10") $current_month_id="2022-11";
      if ($current_month=="2022-11") $current_month_id="2022-12";
      if ($current_month=="2022-12") $current_month_id="2023-01";
      return $current_month_id;
}


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
        array_push($monthlist,"2022-05");
        array_push($monthlist,"2022-06");
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
        array_push($monthlist,"2022-05");
        array_push($monthlist,"2022-06");
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

function process_payments($data) {

        $date=date_create();
        $m_id=date_format($date,'Y-m');

        if ($m_id=="2022-05") $month_id="2022-06";
        if ($m_id=="2022-06") $month_id="2022-07";
        if ($m_id=="2022-07") $month_id="2022-08";
        if ($m_id=="2022-08") $month_id="2022-09";
        if ($m_id=="2022-09") $month_id="2022-10";
        if ($m_id=="2022-10") $month_id="2022-11";
        if ($m_id=="2022-11") $month_id="2022-12";
        if ($m_id=="2022-12") $month_id="2023-01";

        $X=new XRDB();
        $post=array();

	$post['table_name']="nua_payment";


        if (strtoupper($data['company_relationship'])=="EMPLOYEE") {
	    // Get Employee ID.
	    if ($data['employee_code']!="") {
                $data['id']=str_replace('n','',$data['employee_code']);
            } else {
                $data['id']="";
	    }


            if ($data['id']=="0"||$data['id']=="") {
$sql="select * from nua_employee where social_security_number = '" . $data['social_security_number'] . "' and company_id = " . $company_id;
               $y=$this->X->sql($sql);
               if (sizeof($y)>0) { $post['id']=$y[0]['id']; }
            } else {
                  $post['id']=$data['id'];
            }

            //-- Update Employee Record
            $post['table_name']="nua_employee";
            $post['action']="insert";

            $post['company_id']=$data['company_id'];
            $post['first_name']=strtoupper($data['first_name']);
            $post['last_name']=strtoupper($data['last_name']);
            $post['middle_name']=strtoupper($data['middle_initial']);
            $post['email']=strtolower($data['email']);
            $post['date_hired']=$data['date_hired'];
            $post['marital_status']=$data['marital_status'];
            $post['gender']=$data['gender'];
            $post['date_of_birth']=$data['date_of_birth'];
            $post['address']=$data['address'];
            $post['state']=$data['state'];
            $post['city']=$data['city'];
            $post['suite']=$data['suite'];
            $post['zip']=$data['zip'];
            $post['phone']=$data['contact_phone'];
            $post['employee_name']=strtoupper($data['last_name']) . ", " . strtoupper($data['first_name']);
            $post['work_status']=$data['work_status'];
            $post['social_security_number']=$data['social_security_number'];
            //$post['dependent_social_security_number
            //

            $eff_dt="2022-01-01";
            if ($data['min_month_id']=="2022-02") $eff_dt="02/01/2022";
            if ($data['min_month_id']=="2022-03") $eff_dt="03/01/2022";
            if ($data['min_month_id']=="2022-04") $eff_dt="04/01/2022";
            if ($data['min_month_id']=="2022-05") $eff_dt="05/01/2022";
            if ($data['min_month_id']=="2022-06") $eff_dt="06/01/2022";
            if ($data['min_month_id']=="2022-07") $eff_dt="07/01/2022";
            if ($data['min_month_id']=="2022-08") $eff_dt="08/01/2022";
            if ($data['min_month_id']=="2022-09") $eff_dt="09/01/2022";
            if ($data['min_month_id']=="2022-10") $eff_dt="10/01/2022";
            if ($data['min_month_id']=="2022-11") $eff_dt="11/01/2022";
            if ($data['min_month_id']=="2022-12") $eff_dt="12/01/2022";

            $ter_dt="";
            if ($data['max_month_id']=="2022-02") $ter_dt="02/28/2022";
            if ($data['max_month_id']=="2022-03") $ter_dt="03/31/2022";
            if ($data['max_month_id']=="2022-04") $ter_dt="04/30/2022";
            if ($data['max_month_id']=="2022-05") $ter_dt="05/31/2022";
            if ($data['max_month_id']=="2022-06") $ter_dt="06/30/2022";
            if ($data['max_month_id']=="2022-07") $ter_dt="07/31/2022";
            if ($data['max_month_id']=="2022-08") $ter_dt="08/31/2022";
            if ($data['max_month_id']=="2022-09") $ter_dt="09/30/2022";
            if ($data['max_month_id']=="2022-10") $ter_dt="10/31/2022";
            if ($data['max_month_id']=="2022-11") $ter_dt="11/30/2022";
            if ($data['max_month_id']=="2022-12") $ter_dt="12/31/2022";

            $post['effective_date']=$eff_dt;
            $post['medical_plan']=strtoupper($data['medical_plan']);
            $post['medical_plan_level']=strtoupper($data['medical_coverage_level']);
            $post['dental_plan']=strtoupper($data['dental_plan']);
            $post['dental_plan_level']=strtoupper($data['dental_coverage_level']);
            $post['vision_plan']=strtoupper($data['vision_plan']);
            $post['vision_plan_level']=strtoupper($data['vision_coverage_level']);

            $employee_id=$this->X->post($post);
            $post['id']=$employee_id;
            $post['employee_code']='n' . $employee_id;
            $this->X->post($post);
            if ($data['max_month_id']=="") {
                    $data['max_month_id']=$month_id;
                    $term="N";
            } else {
                    $term="Y";
            }
            $monthlist=$this->makeMonthList($data['min_month_id'],$data['max_month_id']);
            $deletes=$this->makeDeleteList($data['min_month_id'],$data['max_month_id']);

            foreach ($deletes as $m) {
                   $sql="delete from nua_monthly_member_census where employee_id = '" . $employee_id . "'";
                   $sql.=" and month_id = '" . $m . "' ";
                   $X->execute($sql);
            }

            foreach($monthlist as $mo_id) {

                   $p=array();
                   $p['table_name']="nua_monthly_member_census";
                   $p['action']="insert";
                   $p['month_id']=$mo_id;
                   $p['employee_code']='n' . $employee_id;
                   $p['dependent_code']="";
                   $p['ssn']=$data['social_security_number'];
                   $p['company_id']=$company_id;
                   $sql="select * from nua_company where id = " . $company_id;
                   $c=$X->sql($sql);

                   $p['company_name']=$c[0]['company_name'];
                   $company_name=$c[0]['company_name'];
                   $p['last_name']=strtoupper($data['last_name']);
                   $p['first_name']=strtoupper($data['first_name']);
                   $p['middle_initial']=strtoupper($data['middle_initial']);
                   $p['middle_name']=strtoupper($data['middle_initial']);

                   $data['relationship']="EMPLOYEE";
                   $p['relationship']=$data['relationship'];
                   $p['eff_dt']=$eff_dt;
                   $p['dob']=substr($data['date_of_birth'],0,10);
                   $p['ssn']=$data['social_security_number'];
                   $p['gender']=strtoupper($data['gender']);
                   $p['marital_status']=strtoupper(substr($data['marital_status'],0,1));
                   $p['address']=strtoupper($data['address']);
                   $p['address2']=strtoupper($data['suite']);
                   $p['city']=strtoupper($data['city']);
                   $p['state']=strtoupper($data['state']);
                   $p['zip']=strtoupper($data['zip']);
                   $p['email']=strtolower($data['email']);
                   $p['phone']=$data['contact_phone'];
                   $p['date_hired']=substr($data['date_hired'],0,10);

if ($data['medical_plan']!='') {
                               $sql="select * from nua_company_plan where plan_code = '" . strtoupper($data['medical_plan']) . "' and company_id = " . $company_id;
                               $r=$X->sql($sql);
                               if (sizeof($r)>0) {
                                  $ee_price=$r[0]['ee_price'];
                                  $ees_price=$r[0]['ees_price'];
                                  $eec_price=$r[0]['eec_price'];
                                  $fam_price=$r[0]['fam_price'];
                                  $med_apa_code=strtoupper($r[0]['APA_CODE']);
                                  $ep=array();
                                  $ep['table_name']="nua_employee_plan";
                                  $ep['action']="insert";
                                  $ep['employee_id']=$employee_id;
                                  $ep['plan_id']=strtoupper($data['medical_plan']);
                                  if ($data['medical_coverage_level']=="EE") $ep['peo_premium']=$ee_price;
                                  if ($data['medical_coverage_level']=="ES") $ep['peo_premium']=$ees_price;
                                  if ($data['medical_coverage_level']=="EC") $ep['peo_premium']=$eec_price;
                                  if ($data['medical_coverage_level']=="FAM") $ep['peo_premium']=$fam_price;
                                  $ep['effective_date']=$eff_dt;
                                  $ep['plan_type']=$data['medical_coverage_level'];
                                  $ep['plan_code_type']="*MEDICAL*";
                                  $sql="select * from nua_employee_plan where employee_id = " . $employee_id . " and ";
                                  $sql.=" plan_code_type = '*MEDICAL*'";
                                  $u=$this->X->sql($sql);
                                  if (sizeof($u)>0) $ep['id']=$u[0]['id'];
                                  $this->X->post($ep);
                               }  else {
                                  $emp['medical_plan']="";
                                  $emp['medical_plan_level']="";
                                  $med_apa_code="";
                               }
                   }
                   if ($data['dental_plan']!='') {
$sql="select * from nua_company_plan where plan_code = '" . strtoupper($data['dental_plan']) . "' and company_id = " . $company_id;
                               $r=$X->sql($sql);
                               if (sizeof($r)>0) {
                                  $dee_price=$r[0]['ee_price'];
                                  $dees_price=$r[0]['ees_price'];
                                  $deec_price=$r[0]['eec_price'];
                                  $dfam_price=$r[0]['fam_price'];
                                  $dental_apa_code="GUARDHIGH";
                                  $ep=array();
                                  $ep['table_name']="nua_employee_plan";
                                  $ep['action']="insert";
                                  $ep['employee_id']=$employee_id;
                                  $ep['plan_id']=strtoupper($data['medical_plan']);
                                  if ($data['medical_coverage_level']=="EE") $ep['peo_premium']=$ee_price;
                                  if ($data['medical_coverage_level']=="ES") $ep['peo_premium']=$ees_price;
                                  if ($data['medical_coverage_level']=="EC") $ep['peo_premium']=$eec_price;
                                  if ($data['medical_coverage_level']=="FAM") $ep['peo_premium']=$fam_price;
                                  $ep['effective_date']=$eff_dt;
                                  $ep['plan_type']=$data['medical_coverage_level'];
                                  $ep['plan_code_type']="*DENTAL*";
                                  $sql="select * from nua_employee_plan where employee_id = " . $employee_id . " and ";
                                  $sql.=" plan_code_type = '*DENTAL*'";
                                  $u=$this->X->sql($sql);
                                  if (sizeof($u)>0) $ep['id']=$u[0]['id'];
                                  $this->X->post($ep);
                               } else {
                                  $dee_price="0.00";
                                  $dees_price="0.00";
                                  $deec_price="0.00";
                                  $dfam_price="0.00";
                                  $dental_apa_code="GUARDHIGH";
                               }
                   }
                  if ($data['vision_plan']!='') {
                               $sql="select * from nua_company_plan where plan_code = '" . strtoupper($data['vision_plan']) . "' and company_id = " . $company_id;
                               $r=$X->sql($sql);
                               if (sizeof($r)>0) {
                                  $vee_price=$r[0]['ee_price'];
                                  $vees_price=$r[0]['ees_price'];
                                  $veec_price=$r[0]['eec_price'];
                                  $vfam_price=$r[0]['fam_price'];
                                  $vision_apa_code="VSP";
                                  $ep=array();
                                  $ep['table_name']="nua_employee_plan";
                                  $ep['action']="insert";
                                  $ep['employee_id']=$employee_id;
                                  $ep['plan_id']=strtoupper($data['medical_plan']);
                                  if ($data['medical_coverage_level']=="EE") $ep['peo_premium']=$ee_price;
                                  if ($data['medical_coverage_level']=="ES") $ep['peo_premium']=$ees_price;
                                  if ($data['medical_coverage_level']=="EC") $ep['peo_premium']=$eec_price;
                                  if ($data['medical_coverage_level']=="FAM") $ep['peo_premium']=$fam_price;
                                  $ep['effective_date']=$eff_dt;
                                  $ep['plan_type']=$data['medical_coverage_level'];
                                  $ep['plan_code_type']="*VISION*";
                                  $sql="select * from nua_employee_plan where employee_id = " . $employee_id . " and ";
                                  $sql.=" plan_code_type = '*VISION*'";
                                  $u=$this->X->sql($sql);
                                  if (sizeof($u)>0) $ep['id']=$u[0]['id'];
                                  $this->X->post($ep);
                               } else {
                                  $vee_price="0.00";
                                  $vees_price="0.00";
                                  $veec_price="0.00";
                                  $vfam_price="0.00";
                                  $vision_apa_code="VSP";
                               }
		  }
                           $p['month_id']=$mo_id;
                           $p['eff_dt']=$eff_dt;
                           $p['employee_id']=$employee_id;
                           $p['employee_code']='n'.$employee_id;
                           $p['company_name']=$company_name;
                           $p['company_id']=$company_id;
                           if ($data['medical_plan']!="") {
                                 $p['client_plan']=strtoupper($data['medical_plan']);
                                 $p['coverage_level']=$data['medical_coverage_level'];
                                 if ($p['coverage_level']=="EE"||$p['coverage_level']=="SI") { $p['coverage_price']=$ee_price; }
                                 if ($p['coverage_level']=="ES"||$p['coverage_level']=="EES") { $p['coverage_price']=$ees_price; }
                                 if ($p['coverage_level']=="EC"||$p['coverage_level']=="EEC") { $p['coverage_price']=$eec_price; }
                                 if ($p['coverage_level']=="FAM"||$p['coverage_level']=="FA") { $p['coverage_price']=$fam_price; }
                                 $p['apa_plan']=$med_apa_code;
                                 $p['plan_type']="*MEDICAL*";
                                 $sql="SELECT id from nua_monthly_member_census where employee_id = " . $employee_id;
                                 $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                 $sql.=" and plan_type = '*MEDICAL*' and dependent_code = ''";
                                 $dd=$X->sql($sql);
                                 if (sizeof($dd)==1) {
                                        $p['id']=$dd[0]['id'];
                                 } else {
                                        $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                        $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                        $sql.=" and plan_type = '*MEDICAL*' and dependent_code = ''";
                                        $X->execute($sql);
                                 }
                                 $this->X->post($p);
                                 if (isset($p['id'])) unset($p['id']);
                                 if ($mo_id==$data['min_month_id']) {
                                     $add=$p;
                                     $add['table_name']="nua_monthly_member_additions";
                                     $sql="SELECT id from nua_monthly_member_additions where employee_id = " . $employee_id;
                                     $sql.=" and month_id = '" . $add['month_id'] . "' ";
                                     $sql.=" and plan_type = '*MEDICAL*' and dependent_code = ''";
                                     $dd=$X->sql($sql);
                                     if (sizeof($dd)==1) {
                                        $add['id']=$dd[0]['id'];
                                     } else {
                                        $sql="delete from nua_monthly_member_additions where employee_id = " . $employee_id;
                                        $sql.=" and plan_type = '*MEDICAL*' and dependent_code = ''";
                                        $X->execute($sql);
                                     }
                                     $this->X->post($add);
                                   }
                                if ($term=='Y'&&$mo_id==$data['max_month_id']) {
                                     if (isset($p['id'])) unset($p['id']);
                                     $add=$p;
                                     $add['table_name']="nua_monthly_member_terminations";
                                     $add['term_dt']=$ter_dt;
                                     $sql="SELECT id from nua_monthly_member_terminations where employee_id = " . $employee_id;
                                     $sql.=" and month_id = '" . $add['month_id'] . "' ";
                                     $sql.=" and plan_type = '*MEDICAL*' and dependent_code = ''";
                                     $dd=$X->sql($sql);
                                     if (sizeof($dd)==1) {
                                        $add['id']=$dd[0]['id'];
                                     } else {
                                        $sql="delete from nua_monthly_member_terminations where employee_id = " . $employee_id;
                                        $sql.=" and plan_type = '*MEDICAL*' and dependent_code = ''";
                                        $X->execute($sql);
                                     }
                                     $this->X->post($add);
				     $sql="update nua_monthly_member_census set term_dt = '" . $ter_dt . "' where ";
				     $sql.=" employee_id = " . $employee_id . " and plan_type = '*MEDICAL*'";
				     $this->X->execute($sql);
                                   }
                            } else {
                                 $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                 $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                 $sql.=" and plan_type = '*MEDICAL*' and dependent_code = ''";
                                 $X->execute($sql);
                            }
                           if ($data['dental_plan']!="") {
                                if (isset($p['id'])) unset($p['id']);
                                $p['client_plan']=strtoupper($data['dental_plan']);
                                $p['coverage_level']=$data['dental_coverage_level'];
                                if ($p['coverage_level']=="EE"||$p['coverage_level']=="SI") { $p['coverage_price']=$dee_price; }
                                if ($p['coverage_level']=="ES"||$p['coverage_level']=="EES") { $p['coverage_price']=$dees_price; }
                                if ($p['coverage_level']=="EC"||$p['coverage_level']=="EEC") { $p['coverage_price']=$deec_price; }
                                if ($p['coverage_level']=="FAM"||$p['coverage_level']=="FA") { $p['coverage_price']=$dfam_price; }
                                $p['apa_plan']="GUARDHIGH";
                                $p['plan_type']="*DENTAL*";
                                $sql="SELECT id from nua_monthly_member_census where employee_id = " . $employee_id;
                                $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                $dd=$X->sql($sql);
                                if (sizeof($dd)==1) {
                                     $p['id']=$dd[0]['id'];
                                } else {
                                     $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                     $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                     $X->execute($sql);
                                 }
				$ttt=$this->X->post($p);
                                 if ($mo_id==$data['min_month_id']) {
                                     $add=$p;
				     if (isset($add['id'])) unset($add['id']);
                                     $add['table_name']="nua_monthly_member_additions";
                                     $sql="SELECT id from nua_monthly_member_additions where employee_id = " . $employee_id;
                                     $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                     $dd=$X->sql($sql);
                                     if (sizeof($dd)==1) {
                                        $add['id']=$dd[0]['id'];
                                     } else {
                                        $sql="delete from nua_monthly_member_additions where employee_id = " . $employee_id;
                                        $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                        $X->execute($sql);
                                     }
                                     $this->X->post($add);
                                   }
                                 if ($term=='Y'&&$mo_id==$data['max_month_id']) {
                                     $add=$p;
				     if (isset($add['id'])) unset($add['id']);
                                     $add['table_name']="nua_monthly_member_terminations";
                                     $add['term_dt']=$ter_dt;
                                     $sql="SELECT id from nua_monthly_member_terminations where employee_id = " . $employee_id;
                                     $sql.=" and month_id = '" . $add['month_id'] . "' ";
                                     $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                     $dd=$X->sql($sql);
                                     if (sizeof($dd)==1) {
                                        $add['id']=$dd[0]['id'];
                                     } else {
                                        $sql="delete from nua_monthly_member_terminations where employee_id = " . $employee_id;
                                        $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                        $X->execute($sql);
                                     }
                                     $this->X->post($add);
				     $sql="update nua_monthly_member_census set term_dt = '" . $ter_dt . "' where ";
				     $sql.=" employee_id = " . $employee_id . " and plan_type = '*DENTAL*'";
				     $this->X->execute($sql);
                                   }
                            } else {
                                $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                 $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                 $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                 $X->execute($sql);
                            }
                            if ($data['vision_plan']!="") {
                                if (isset($p['id'])) unset($p['id']);
                                 $p['client_plan']=strtoupper($data['vision_plan']);
                                 $p['coverage_level']=$data['vision_coverage_level'];
                                 $p['apa_plan']="VSP";
                                 if ($p['coverage_level']=="EE"||$p['coverage_level']=="SI") { $p['coverage_price']=$vee_price; }
                                 if ($p['coverage_level']=="ES"||$p['coverage_level']=="EES") { $p['coverage_price']=$vees_price; }
                                 if ($p['coverage_level']=="EC"||$p['coverage_level']=="EEC") { $p['coverage_price']=$veec_price; }
                                 if ($p['coverage_level']=="FAM"||$p['coverage_level']=="FA") { $p['coverage_price']=$vfam_price; }
                                 $p['plan_type']="*VISION*";
                                 $sql="SELECT id from nua_monthly_member_census where employee_id = " . $employee_id;
                                 $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                 $sql.=" and plan_type = '*VISION*' and dependent_code = ''";
                                 $dd=$X->sql($sql);
                                 if (sizeof($dd)==1) {
                                       $p['id']=$dd[0]['id'];
                                 } else {
                                      $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                      $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                      $sql.=" and plan_type = '*VISION*' and dependent_code = ''";
                                      $X->execute($sql);
                                 }
                                 $this->X->post($p);
                                 if (isset($p['id'])) unset($p['id']);
                                 if ($mo_id==$data['min_month_id']) {
                                     $add=$p;
                                     if (isset($add['id'])) unset($add['id']);
                                     $add['table_name']="nua_monthly_member_additions";
                                     $sql="SELECT id from nua_monthly_member_additions where employee_id = " . $employee_id;
                                     $sql.=" and plan_type = '*VISION*' and dependent_code = ''";
                                     $dd=$X->sql($sql);
                                     if (sizeof($dd)==1) {
                                        $add['id']=$dd[0]['id'];
                                     } else {
                                        $sql="delete from nua_monthly_member_additions where employee_id = " . $employee_id;
                                        $sql.=" and plan_type = '*VISION*' and dependent_code = ''";
                                        $X->execute($sql);
                                     }
                                     $this->X->post($add);
                                   }
                                if ($term=='Y'&&$mo_id==$data['max_month_id']) {
                                     $add=$p;
                                     if (isset($add['id'])) unset($add['id']);
                                     $add['table_name']="nua_monthly_member_terminations";
                                     $add['term_dt']=$ter_dt;
                                     $sql="SELECT id from nua_monthly_member_terminations where employee_id = " . $employee_id;
                                     $sql.=" and month_id = '" . $add['month_id'] . "' ";
                                     $sql.=" and plan_type = '*VISION*' and dependent_code = ''";
                                     $dd=$X->sql($sql);
                                     if (sizeof($dd)==1) {
                                        $add['id']=$dd[0]['id'];
                                     } else {
                                        $sql="delete from nua_monthly_member_terminations where employee_id = " . $employee_id;
                                        $sql.=" and plan_type = '*VISION*' and dependent_code = ''";
                                        $X->execute($sql);
                                     }
                                     $this->X->post($add);
				     $sql="update nua_monthly_member_census set term_dt = '" . $ter_dt . "' where ";
				     $sql.=" employee_id = " . $employee_id . " and plan_type = '*VISION*'";
				     $this->X->execute($sql);
                                   }
                            } else {
                                 $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                 $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                 $sql.=" and plan_type = '*VISION*' and dependent_code = ''";
                                 $X->execute($sql);
                            }
	    }  // Months 
	} // Employee                   
	 if (strtoupper($data['relationship'])!="EMPLOYEE") {
                        if ($data['employee_code']!="") {
                             $sql="select id, employee_code from nua_employee where company_id = " . $company_id . " and employee_code = '" . $data['employee_code'] . "'";
                        } else {
                             $sql="select id, employee_code from nua_employee where company_id = " . $company_id . " and social_security_number = '" . $data['social_security_number'] . "'";
                        }
                        $emps=$X->sql($sql);
                        if (sizeof($emps)>0) {
                                $employee_id=$emps[0]['id'];
                                $employee_code=$emps[0]['employee_code'];

                                $dep=array();
                                $dep['table_name']="nua_employee_dependent";
                                $dep['action']="insert";
                                //--
                                //-- Find Existing Dependent
                                //--
                                $sql="select * from nua_employee_dependent where employee_id = " . $employee_id;
                                $sql.=" and social_security_number = '" . $data['dependent_social_security_number'] . "'";
                                $e=$X->sql($sql);

                                $dep['employee_id']=$employee_id;
                                $dep['employee_code']="n" . $employee_id;
                                $dep['company_id']=$company_id;
                                $dep['first_name']=strtoupper($data['first_name']);
                                $dep['last_name']=strtoupper($data['last_name']);
                                $dep['middle_name']=strtoupper($data['middle_initial']);
                                $dep['gender']=strtoupper($data['gender']);
                                $dep['date_of_birth']=substr($data['date_of_birth'],0,10);
                                $dep['social_security_number']=$data['dependent_social_security_number'];
                                $dep['relationship']=strtoupper($data['relationship']);
                                if (strtoupper($dep['relationship'])=='WIFE') $dep['relation_type']="02";
                                if (strtoupper($dep['relationship'])=='HUSBAND') $dep['relation_type']="03";
                                if (strtoupper($dep['relationship'])=='SON') $dep['relation_type']="04";
                                if (strtoupper($dep['relationship'])=='DAUGHTER') $dep['relation_type']="05";
                                if (strtoupper($dep['relationship'])=='OTHER') $dep['relation_type']="06";
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
                               $this->X->post($dep);
}
                   }  // DEPENDENT
}  // FUNCTION



} 

