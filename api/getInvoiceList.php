<?php
ini_set('memory_limit', -1);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');

ini_set('display_errors',1);
ini_set('display_startup_errors',1);

function getOpenInvoiceReportExcel($billingCycle)
{
	$X=new XRDB();


$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Open Invoice Report');

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
	$spreadsheet->getActiveSheet()->setCellValue('A1', "Company ID");

	$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(50);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('B1', "Company Name");

	$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(50);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('C1', "Status");

	$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('D1', "Medical Count");
	
	$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('E1', "Dental Count");

	$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('F1', "Vision Count");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('G1', "ADD/Life Count");

	$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('H1', "Invoicing");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('I1', "EFT");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('J1', "PEO");
		
	$sql="select * from nua_company where status = 'Enrolled' order by company_name";
	$eft=$X->sql($sql);

        $row=1;

        foreach($eft as $eftRow) {
		
            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['id']);
		
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['company_name']);
	
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['status']);
	
		$sql="select count(*) as c from nua_monthly_member_census where month_id = '" . $month_id . "' ";
		$sql.="and company_id = " . $eftRow['id'] . " and plan_type = '*MEDICAL*'";
	        $e=$X->sql($sql);
		$med=$e[0]['c'];

		$sql="select count(*) as c from nua_monthly_member_census where month_id = '" . $month_id . "' ";
		$sql.="and company_id = " . $eftRow['id'] . " and plan_type = '*DENTAL*'";
	        $e=$X->sql($sql);
		$dental=$e[0]['c'];

		$sql="select count(*) as c from nua_monthly_member_census where month_id = '" . $month_id . "' ";
		$sql.="and company_id = " . $eftRow['id'] . " and plan_type = '*VISION*'";
	        $e=$X->sql($sql);
		$vision=$e[0]['c'];

		$sql="select count(*) as c from nua_monthly_member_census where month_id = '" . $month_id . "' ";
		$sql.="and company_id = " . $eftRow['id'] . " and plan_type in ('*ADD*','*LIFE*')";
	        $e=$X->sql($sql);
		$add=$e[0]['c'];

		$sql="select org_name from nua_org where id = " . $eftRow['org_id'];
	        $e=$X->sql($sql);
		if (sizeof($e)>0) {
	              $orgd=$e[0]['org_name'];
                } else {
	              $orgd="";
		}
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $med);
	
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $dental);
		
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $vision);
		
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $add);

            $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['invoicing']);

            $cell="I".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['eft_flag']);

            $cell="J".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $orgd);

	}

        $spreadsheet->setActiveSheetIndex(0);
        $writer = new Xlsx($spreadsheet);
        $filename="Active_Client_Report.xlsx";

        header('Content-disposition: attachment; filename=' . $filename);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
        header('Cache-Control: must-revalidate');
        sleep(1);
        $writer->save("php://output");
        die();

}

$X=new XRDB();
if (isset($_GET['billingCycle'])) {
	$billingCycle = $_GET['billingCycle'];
  }
else
{
   echo("Billing cycle not provided");
   die();
}

getOpenInvoiceReportExcel($billingCycle);
?>


