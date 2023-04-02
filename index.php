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
        require('/var/www/classes/class.XRDB.php');

	$X=new XRDB();

	$id=$_GET['id'];
        $id_arr=str_split($id);
	$len_str="";
	$flag=0;
	for($i=2;$i<sizeof($id_arr);$i++) {
		if ($flag==0) {
                    if ($id_arr[$i]=='a') $flag=1;
		    if ($flag==0) {
                       $len_str.=$id_arr[$i];
	            }
		}
	}
	$len=intval($len_str);
	$real_id='';
        $hashloc=4;
	while(strlen($real_id)<$len) {
           $real_id.=$id_arr[$hashloc];
	   $next=fmod(intval($id_arr[$hashloc]),5)+1;
	   $hashloc+=intval($next); 
	}
	$key=md5('nuaxess'.$real_id);
	$dir=substr($key,0,7);
	$f=substr($key,7,7);
        $file="/var/www/docs/" . $dir . "/" . $f;
        $a=file_get_contents("/var/www/docs/" . $dir . "/" . $f);
	$sql="select * from nua_doc where id = " . $real_id;
	$f=$X->sql($sql);
	if (sizeof($f)==0) {
            echo "Document not found";
	} else {
	     $file_name=$f[0]['file_name'];
             $ar_ext = explode('.', $file_name);
             $ext = strtolower(end($ar_ext));
             $extensions = array(
                'bmp' => 'image/bmp',
                'csv' => 'text/csv',
                'doc' => 'application/msword',
                'docx' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'exe' => 'application/octet-stream',
                'gif' => 'image/gif',
                'htm' => 'text/html',
                'html' => 'text/html',
                'ico' => 'image/vnd.microsoft.icon',
                'jpeg' => 'image/jpg',
                'jpe' => 'image/jpg',
                'jpg' => 'image/jpg',
                'pdf' => 'application/pdf',
                'png' => 'image/png',
                'ppt' => 'application/vnd.ms-powerpoint',
                'psd' => 'image/psd',
                'swf' => 'application/x-shockwave-flash',
                'tif' => 'image/tiff',
                'tiff' => 'image/tiff',
                'xhtml' => 'application/xhtml+xml',
                'xml' => 'application/xml',
                'xls' => 'application/vnd.ms-excel',
                'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'zip' => 'application/zip'
              );
               $ctype = isset($extensions[$ext]) ? $extensions[$ext] : 'application/force-download';

               if(ini_get('zlib.output_compression')) ini_set('zlib.output_compression', 'Off');


    header('Pragma: public'); 
    header('Expires: 0');
 //   header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
    header('Cache-Control: private',false); 
    header('Content-Type: '. $ctype);
//    header('Content-Disposition: attachment; filename='. $file_name .';' );
    header('Content-Transfer-Encoding: binary');
    header('Content-Length: '. filesize($file));
  //  readfile($file);
           echo $a;
     }
?>

