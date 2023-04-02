<?php
ini_set('memory_limit', -1);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');

ini_set('display_errors',1);
ini_set('display_startup_errors',1);


$X=new XRDB();
if (isset($_GET['billingCycle'])) {
	$billingCycle = $_GET['billingCycle'];
} 
else
{
   echo("Billing cycle not provided");
   die();
}

function infinitiStats($billingCycle)
{
	$X=new XRDB();


$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Plan Utilization');

$styleArrayEnrollment = [
        'font' => [
        'bold' => true,
        ],
        'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        ],
        'borders' => [
                'top' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'bottom' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'left' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'right' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
],
'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
],
];


$styleArrayLine = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
        'height' => '10',
    ],
    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    ],
];


	$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('A1', "Category");

	$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(50);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('B1', "Client Plan");

	$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('C1', "EE Count");
	
	$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('D1', "ES Count");

	$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('E1', "EC Count");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('F1', "FAM Count");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('G1', "Total Count");
		
	$plan_type=array();
        array_push($plan_type,'ESSENTIAL');
        array_push($plan_type,'PREMIER');
        array_push($plan_type,'PRIME');
        array_push($plan_type,'COPPER');
        array_push($plan_type,'BRONZE');
        array_push($plan_type,'SILVER');
        array_push($plan_type,'GOLD');
        array_push($plan_type,'PLATINUM');

        $row=1;
		$total_total_total=0;
		$total_total_ee=0;
		$total_total_es=0;
		$total_total_ec=0;
		$total_total_fam=0;

	foreach($plan_type as $t) {

		$total_total=0;
		$total_ee=0;
		$total_es=0;
		$total_ec=0;
		$total_fam=0;

	$sql="select distinct client_plan from ";
	$sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	$sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	$sql.=" and client_plan like '%" . $t . "%' order by client_plan";
	$eft=$X->sql($sql);


        foreach($eft as $e) {
		
		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and client_plan = '" . $e['client_plan']  . "'";
	        $data=$X->sql($sql);

		$total=$data[0]['c'];
		$total_total+=$total;
		$total_total_total+=$total;

		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and coverage_level in ('EE','SI') and client_plan = '" . $e['client_plan']  . "'";

	        $data=$X->sql($sql);

		$ee=$data[0]['c'];
		$total_ee+=$ee;
		$total_total_ee+=$ee;

		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and coverage_level in ('ES','ES2') and client_plan = '" . $e['client_plan']  . "'";
	        $data=$X->sql($sql);

		$es=$data[0]['c'];
		$total_es+=$es;
		$total_total_es+=$es;

		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and coverage_level in ('EC','EC2') and client_plan = '" . $e['client_plan']  . "'";
	        $data=$X->sql($sql);

		$ec=$data[0]['c'];
		$total_ec+=$ec;
		$total_total_ec+=$ec;


		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and coverage_level in ('FAM','FA') and client_plan = '" . $e['client_plan']  . "'";
	        $data=$X->sql($sql);

		$fam=$data[0]['c'];
		$total_fam+=$fam;
		$total_total_fam+=$fam;

            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $t);
	
		
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $e['client_plan']);
	
	
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $ee);
	
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $es);
	
	
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $ec);

		
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $fam);
		
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total);

	}

            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $t);
	
		
	
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_ee);
	
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_es);
	
	
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_ec);

		
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_fam);
		
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_total);
            $row++;


	}
            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, 'GRAND TOTAL');
	
		
	
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_total_ee);
	
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_total_es);
	
	
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_total_ec);

		
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_total_fam);
		
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $total_total_total);
            $row++;
$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Enrollment_Stats_Report.xlsx";

    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();

}

infinitiStats($billingCycle);

?>


