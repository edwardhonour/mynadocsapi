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

require('class.XRDB.php');
require('simpleXLSX.php');

$X=new DB();

$sql="select * from va_facility_load";
$list=$X->sql($sql);
$fid=27682976;
foreach($list as $l) {
	$fid++;
	$own="300029";
	$pt="";
        if ($l['LEASED_OWNED']=="Owned") $own="200004";
        if ($l['LEASED_OWNED']=="Leased") $own="200029";
	if ($l['PROPERTY_TYPE']=="Land") $pt="1";
	if ($l['PROPERTY_TYPE']=="Building") $pt="2";
	if ($l['PROPERTY_TYPE']=="Structure") $pt="3";
        if ($pt=="1") { $building="N"; } else { $building="Y"; }
        $b="";
        if ($l['SUB_ORGANIZATION']=='National Cemetary Administration') $b.="VC";
        if ($l['SUB_ORGANIZATION']=='Veterans Health Administration') $b.="VH";
        if ($l['SUB_ORGANIZATION']=='Veterans Benefits Administration') $b.="VB";
        if ($l['SUB_ORGANIZATION']=='Department of Veterans Affairs') $b.="VA";
        if ($l['SUB_ORGANIZATION']=='Inspector General') $b.="VI";
        if ($l['SUB_ORGANIZATION']=='Administrative Service') $b.="VS";
        if ($l['SUB_ORGANIZATION']=='Office of Acquisiton and Materials Management') $b.="VM";
        if ($l['SUB_ORGANIZATION']=='Office of Information Systems and Telecommunications') $b.="VS";
	$pp="";
	if ($l['PROPERTY_USE']=="Parks and Historic Sites") $pp=1;
	if ($l['PROPERTY_USE']=="Service") $pp=2;
	if ($l['PROPERTY_USE']=="Industrial") $pp=3;
	if ($l['PROPERTY_USE']=="All Other Land") $pp=4;
	if ($l['PROPERTY_USE']=="Utility Systems") $pp=5;
	if ($l['PROPERTY_USE']=="Warehouses") $pp=6;
	if ($l['PROPERTY_USE']=="Comfort Station/Restrooms") $pp=7;
	if ($l['PROPERTY_USE']=="Public Facing Facility") $pp=8;
	if ($l['PROPERTY_USE']=="Outpatient Healthcare") $pp=9;
	if ($l['PROPERTY_USE']=="Housing") $pp=10;
	if ($l['PROPERTY_USE']=="Institutional") $pp=11;
	if ($l['PROPERTY_USE']=="Family Housing") $pp=12;
	if ($l['PROPERTY_USE']=="Office Building Locations") $pp=13;
	if ($l['PROPERTY_USE']=="Hospital") $pp=14;
	if ($l['PROPERTY_USE']=="Roads and Bridges") $pp=15;
	if ($l['PROPERTY_USE']=="Vacant") $pp=16;
	if ($l['PROPERTY_USE']=="Other Institutional Uses") $pp=17;
	if ($l['PROPERTY_USE']=="Dormitories/Barracks") $pp=18;
	if ($l['PROPERTY_USE']=="Laboratories") $pp=19;
	if ($l['PROPERTY_USE']=="Parking Structures") $pp=20;
	if ($l['PROPERTY_USE']=="Data Center") $pp=21;
	if ($l['PROPERTY_USE']=="School") $pp=22;
	if ($l['PROPERTY_USE']=="Communications Systems") $pp=23;
	if ($l['PROPERTY_USE']=="Flood Control and Navigation") $pp=24;
	if ($l['PROPERTY_USE']=="Monuments and Memorials") $pp=25;
	if ($l['PROPERTY_USE']=="Museum") $pp=26;
	if ($l['PROPERTY_USE']=="Reclamation and Irrigation") $pp=27;
	if ($l['PROPERTY_USE']=="Power Development and Distribution") $pp=28;
	if ($l['PROPERTY_USE']=="Facility Security") $pp=29;
	if ($l['PROPERTY_USE']=="Child Care Center") $pp=30;
	if ($l['PROPERTY_USE']=="Post Offices") $pp=31;
	if ($l['PROPERTY_USE']=="Storage") $pp=32;
	if ($l['PROPERTY_USE']=="Renewable Energy System") $pp=33;

	echo "<br>";
	echo "DECLARE<br>";
	echo "FID NUMBER(11);<br>";
	echo "C NUMBER(11);<br>";
	echo "BEGIN<br>";
	echo "SELECT COUNT(*) INTO C FROM TBL_DIM_FACILITY WHERE BUILDING_NBR = '" . $l['BUILDING_NBR'] . "';<br>";
	echo "IF C = 0 THEN<br>";
	echo "SELECT MEANINGLESS_KEY_SEQ.NEXTVAL INTO FID FROM DUAL;<br>";
	echo "INSERT INTO TBL_DIM_FACILITY (SEGMENT_ID, REGION_ID, FACILITY_ID,<br>";
        echo "BUILDING_NBR, FACILITY_NAME, ADDRESS, CITY_NAME, STATE_ABBR, ZIPCODE, <br>";
	echo "OWNERSHIP_ID, PROPERTY_TYPE_ID, FSL, LATITUDE, LONGITUDE, COUNTY, <br>";
	echo "ACTIVE_FLAG, P_BEAT, POPULATION, RENTABLE_SQFT, CAMPUS, AGENCY_BUILDING_NBR, <br>";
	echo "NO_SEC_FEE, FPS_RESPONSIBLE, BUILDING, DISTRICT_ID,FACILITY_TYPE_ID) VALUES (<br>";
	echo "'VA'," . $l['REGION_ID'] . ",FID,'" . $l['BUILDING_NBR'] . "',<br>";
	echo "'" . str_replace("'","",strtoupper($l['ADDRESS'])) . "',<br>";
	echo "'" . str_replace("'","",$l['ADDRESS']) . "','" . str_replace("'","",$l['CITY']) . "',<br>";
	echo "'" . $l['STATE'] . "','" . $l['ZIP'] . "',<br>";
	echo $own . ", " . $pp . ", 3, '" . $l['LATITUDE'] . "',<br>";
        echo "'" . $l['LONGITUDE'] . "', '" . str_replace("'","",$l['COUNTY']) . "', 'Y','" . $l['p_beat'] . "', '" . $l['POPULATION'] . "'<br>";
        echo ", '" . $l['SQUARE_FEET'] . "','N','" . $l['FRPP_ASSET_ID'] . "','N','Y','" . $building . "','" . $l['VISIN'] . "'," . $pp . ");<br>";
	if ($l['SUB_ORGANIZATION']=='National Cemetary Administration') { 
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo "VALUES (FID, FID,'NATIONAL CEMETARY ADMINISTRATION','3610','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Inspector General') { 
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo "VALUES (FID, FID,'INSPECTOR GENERAL','3601','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Administrative Services') { 
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo "VALUES (FID, FID,'ADMINISTRATIVE SERVICES','3614','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Administrative Service') { 
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo "VALUES (FID, FID,'ADMINISTRATIVE SERVICES','3614','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Office of Acquisiton and Materials Management') {
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo "VALUES (FID, FID,'ACQUISITION AND MATERIALS MANAGEMENT','3605','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Office of Acquisiton and Materials Management') {
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo "VALUES (FID, FID,'ACQUISITION AND MATERIALS MANAGEMENT','3605','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Office of Information Systems and Telecommunications') {
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo "VALUES (FID, FID,'INFORMATION SYSTEMS AND TELECOMMUNICATIONS','3607','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Veterans Health Administration') { 
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo " VALUES (FID, FID,'VETERANS HEALTH ADMINISTRATION','3609','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Veterans Benefits Administration') { 
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo " VALUES (FID, FID,'VETERANS BENEFITS ADMINISTRATION','3600','Y','Y');<br>"; 
	}
	if ($l['SUB_ORGANIZATION']=='Department of Veterans Affairs') { 
		echo "INSERT INTO FPS_FACILITY_TENANTS (TENANT_ID, FACILITY_ID, AGENCY_NAME, AGENCY_CODE, ACTIVE_FLAG, FED_TENANT) ";
		echo " VALUES (FID, FID,'DEPARTMENT OF VETERANS AFFAIRS','3611','Y','Y');<br>"; 
	}
	echo "COMMIT;<br>";
	echo "UPDATE TBL_DIM_FACILITY SET SEGMENT_ID = 'VA' WHERE FACILITY_ID = FID;<br>";
	echo "END IF;<br>";
	echo "END;<br>";
	echo "/<br>";
}

?>
