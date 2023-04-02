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
	echo "UPDATE TBL_DIM_FACILITY SET BUILDING_NBR = '" . $l['BUILDING_NBR'] . "', STATE_ABBR = '" . $l['STATE'] . "' ";
	echo "WHERE AGENCY_BUILDING_NBR = '" . $l['FRPP_ASSET_ID'] . "';<br>";
	echo "COMMIT;<br>";
}

?>
