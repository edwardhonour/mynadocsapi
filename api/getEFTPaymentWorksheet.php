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

function getEFTPaymentWorkbookExcel($billingCycle)
{
	$X=new XRDB();


$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Payment Workbook Report');

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

	$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('C1', "Invoice Number");
	
	$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('D1', "Payment Date");

	$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('E1', "Invoice Total");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('F1', "Payment Amount");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('F1', "Deposit Type (ACH,EFT,WIRE,Check)");
		
	$sql="select company_id, company_name, invoice_number, grand_total, payment_amount, billing_date, due_date, amount_left, paid, month_id from nua_company_invoice ";
	$sql.=" where paid in ('No','N','Par') and amount_left <> 0 and company_id in (select id from nua_company where org_id not in (17,19,20,21,22,23,24,25) and flag_eft in ('Y','Yes','YES','y'))  and month_id = '".$billingCycle."' order by company_name";
	$eft=$X->sql($sql);

        $row=1;

        foreach($eft as $eftRow) {
		
		if ($eftRow['month_id']=='2022-01') $dd="2021-12-31";
		if ($eftRow['month_id']=='2022-02') $dd="2022-01-31";
		if ($eftRow['month_id']=='2022-03') $dd="2022-02-28";
		if ($eftRow['month_id']=='2022-04') $dd="2022-03-31";
		if ($eftRow['month_id']=='2022-05') $dd="2022-04-30";
		if ($eftRow['month_id']=='2022-06') $dd="2022-05-31";
		if ($eftRow['month_id']=='2022-07') $dd="07/01/22";
		if ($eftRow['month_id']=='2022-08') $dd="08/01/22";
		if ($eftRow['month_id']=='2022-09') $dd="09/01/22";
		if ($eftRow['month_id']=='2022-10') $dd="10/01/22";
		if ($eftRow['month_id']=='2022-11') $dd="11/01/22";
		if ($eftRow['month_id']=='2022-12') $dd="12/01/22";

            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['company_id']);
	
		
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['company_name']);
	
	
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['invoice_number']);
	
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $dd);
	
	
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['grand_total']);

		
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, "EFT");
		
	}


$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="EFT_Payment_Workbook.xlsx";

    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();

}

getEFTPaymentWorkbookExcel($billingCycle);
