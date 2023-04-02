<?php
	ini_set('display_errors',1);
	ini_set('display_startup_errors',1);
	ini_set('max_execution_time',600);
	ini_set('set_time_limit',600);
	header('Access-Control-Allow-Headers: Access-Control-Allow-Origin, Content-Type, Authorization');
        header('Access-Control-Allow-Origin: *');
//        header('Access-Control-Allow-Methods: GET,PUT,POST,DELETE,PATCH,OPTIONS');
	error_reporting(-1);
	ini_set('memory_limit', '-1');

        $a=file_get_contents("/var/www/docs/first.pdf");
    header('Content-Type: application/pdf');
    header('Content-Disposition: inline; filename="'. $filename . '"');
    echo $a;
 
?>

