<?php
ini_set('memory_limit', -1);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');
require_once('class.enrollment2.php');

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

function getOpenPaymentReport($billingCycle)
{
	$X=new XRDB();


$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Open Payments Report');

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
	$spreadsheet->getActiveSheet()->setCellValue('C1', "Payment Date");
	
	$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('D1', "Amount Received");

	$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('E1', "Amount Applied");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('F1', "Amount Remaining");
		
	$sql="select nua_payment.company_id as company_id, company_name, payment_date, amount_received, amount_applied, amount_remaining from nua_payment, nua_company ";
	$sql.=" where amount_remaining not in ('0','0.00','0.0') and nua_company.id = nua_payment.company_id order by company_name";
	$eft=$X->sql($sql);

        $row=1;

        foreach($eft as $eftRow) {
		
            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['company_id']);
	
		
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['company_name']);
	
	
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['payment_date']);
	
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['amount_received']);
	
	
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['amount_applied']);

		
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['amount_remaining']);
		
	}


$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Open_Payment_Report.xlsx";

    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();

}

getOpenPaymentReport($billingCycle);

?>


