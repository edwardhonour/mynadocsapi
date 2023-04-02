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
      return $current_month_id;
}

function getQuoteTemplate($company_id, $month_id,$display='F') {

$styleArrayEnrollment = [
    'font' => [
        'bold' => true,
    ],
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

$styleArrayBold = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
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

$styleArrayHead = [
    'font' => [
        'bold' => true,
        'size' => '24',
    ],
];

$styleArray = [
    'font' => [
        'bold' => true,
        'size' => '12',
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


$X=new XRDB();
$role="sadmin";
if ($company_id==0) {
    $company=array();
} else {
    $sql="select * from nua_company where id = " . $company_id;
    $t=$X->sql($sql);
    $company=$t[0];
}


}
?php>

//
// 0
//
//

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Company');

if ($company_id!='0') {
//
// 1
//
//
   $worksheet1 = $spreadsheet->createSheet();
   $worksheet1->setTitle('Preenrollment-Census');
}

$quoted_plan="N";
$company_plan="N";

if ($company_id!='0') {

    //--
    //-- Only Admins and Quuoting Plan Team gets page unless its filled out.
    //--

    if ($role=="sadmin"||$role=="quoting") {
        $quoted_plan="Y";
    } else {
        $sql="select * from nua_quoted_plan where company_id = " . $company_id;
        $z=$X->sql($sql);
        $sql="select * from nua_company_plan where company_id = " . $company_id;
        $y=$X->sql($sql);
        if (sizeof($z)>0||sizeof($y>0))  $quoted_plan="Y";
    }
    if ($quoted_plan=="Y") {
        //
        // 2
        //
        //
        $worksheet2 = $spreadsheet->createSheet();
        $worksheet2->setTitle('Quoted-Plans');
    }

    $sql="select * from nua_company_plan where company_id = " . $company_id;
    $z=$X->sql($sql);
    if (sizeof($z)==0) {
        $company_plan="N";
    } else {
        $company_plan="Y";
        $worksheet3 = $spreadsheet->createSheet();
        $worksheet3->setTitle('Accepted-Plans');
    }
    if ($company_plan=="Y") {
       $worksheet4 = $spreadsheet->createSheet();
       $worksheet4->setTitle('Enrollment');
    }
}


//================================================================================
//  COMPANY TEMPLATE
//================================================================================

      $sheet = $spreadsheet->getActiveSheet();
      //-- Column Widths
      $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32);
      $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(42);
      $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(8);
      $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(42);
      $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(42);
      //-- Company Name Title
      //
      $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
      if ($company_id=='0') {
              $spreadsheet->getActiveSheet()->setCellValue('A1', "NEW PROSPECT COMPANY DATA");
      }  else {
              $spreadsheet->getActiveSheet()->setCellValue('A1', $company['company_name']);
      }

      //-- Company ID
      $sheet->setCellValue('A3', "ID");
      $spreadsheet->getActiveSheet()->getStyle('B3')->applyFromArray($styleArrayBold);
      $spreadsheet->getActiveSheet()->getStyle('B3')->getFill()->getStartColor()->setARGB('DDDDDDDD');
      if ($company_id!='0') {
            $spreadsheet->getActiveSheet()->setCellValue('B3', $company['id']);
      } else {
             $spreadsheet->getActiveSheet()->setCellValue('B3', "(NuAxess Use Only)");
      }
                //-- Company Name
                 $sheet->setCellValue('A4', "Company Name");
                 $spreadsheet->getActiveSheet()->getStyle('B4')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B4')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('B4', $company['company_name']);

                 //-- Broker Name
                 $sheet->setCellValue('A5', "Broker Name");
                 $spreadsheet->getActiveSheet()->getStyle('B5')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B5')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('B5', $company['broker_name']);

                 //-- Broker Email
                 $sheet->setCellValue('A6', "Broker Email");
                 $spreadsheet->getActiveSheet()->getStyle('B6')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B6')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('B6', $company['broker_email']);

                 //-- Company Type
                 $sheet->setCellValue('A7', "Company Type");
                 $spreadsheet->getActiveSheet()->getStyle('B7')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B7')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('B7', $company['company_type']);

               //-- Tax ID
                 $sheet->setCellValue('A8', "Tax ID");
                 $spreadsheet->getActiveSheet()->getStyle('B8')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B8')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B8', $company['tax_id']);
                //-- Contact Name
                 $sheet->setCellValue('A9', "Contact Name");
                 $spreadsheet->getActiveSheet()->getStyle('B9')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B9')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B9', $company['contact_name']);

                 //-- Contact Phone
                 $sheet->setCellValue('A10', "Contact Phone");
                 $spreadsheet->getActiveSheet()->getStyle('B10')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B10')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B10', $company['contact_phone']);

                 //-- Contact Email
                 $sheet->setCellValue('A11', "Contact Email");
                 $spreadsheet->getActiveSheet()->getStyle('B11')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B11')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B11', $company['contact_email']);

                 //-- Contact Email
                 $sheet->setCellValue('A12', "Employee Count (est)");
                 $spreadsheet->getActiveSheet()->getStyle('B12')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B12')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B12', $company['employee_count']);

                 //-- Medical
                 $sheet->setCellValue('A14', "Medical (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B14')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B14')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B14', $company['medical']);
                 //
                 //-- Dental
                 $sheet->setCellValue('A15', "Dental (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B15')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B15')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B15', $company['dental']);
                 //
                 //-- Vision
                 $sheet->setCellValue('A16', "Vision (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B16')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B16')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B16', $company['vision']);
                 //
                 //-- Provider
                 $sheet->setCellValue('A18', "Previous Insurance Provider");
                 $spreadsheet->getActiveSheet()->getStyle('B18')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B18')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B18', $company['current_provider']);
                 //
                 //-- Contact Email
                 $sheet->setCellValue('D3', "Company Address");
                 $spreadsheet->getActiveSheet()->getStyle('E3')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E3')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E3', $company['address']);

                 //-- Suite
                 $sheet->setCellValue('D4', "Line 2");
                 $spreadsheet->getActiveSheet()->getStyle('E4')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E4')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E4', $company['suite']);

                 //-- City
                 $sheet->setCellValue('D5', "City");
                 $spreadsheet->getActiveSheet()->getStyle('E5')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E5')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E5', $company['city']);

                 //-- State
                 $sheet->setCellValue('D6', "State");
                 $spreadsheet->getActiveSheet()->getStyle('E6')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E6')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E6', $company['state']);

                 //-- Zip
                 $sheet->setCellValue('D7', "Zip");
                 $spreadsheet->getActiveSheet()->getStyle('E7')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E7')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E7', $company['zip']);

                 //-- Website
                 $sheet->setCellValue('D8', "Website");
                 $spreadsheet->getActiveSheet()->getStyle('E8')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E8')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E8', $company['website']);

                 //-- Billing Address
                 $sheet->setCellValue('D10', "Billing Address");

                 //-- Company Name
                 $sheet->setCellValue('D11', "Company Name");
                 $spreadsheet->getActiveSheet()->getStyle('E11')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E11')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E11', $company['billing_company_name']);

                //-- Mailing Address
                 $sheet->setCellValue('D12', "Address");
                 $spreadsheet->getActiveSheet()->getStyle('E12')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E12')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E12', $company['invoice_mailing_address']);

                 //-- Invoice Suite
                 $sheet->setCellValue('D13', "Suite");
                 $spreadsheet->getActiveSheet()->getStyle('E13')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E13')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E13', $company['invoice_suite']);

                 //-- Invoice City
                 $sheet->setCellValue('D14', "City");
                 $spreadsheet->getActiveSheet()->getStyle('E14')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E14')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E14', $company['invoice_city']);

                 //-- Invoice State
                 $sheet->setCellValue('D15', "State");
                 $spreadsheet->getActiveSheet()->getStyle('E15')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E15')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E15', $company['invoice_state']);


                 //-- Invoice Zip
                 //
                 $sheet->setCellValue('D16', "Zip");
                 $spreadsheet->getActiveSheet()->getStyle('E16')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E16')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E16', $company['invoice_zip']);

                 //-- Invoice Contact Name
                 $sheet->setCellValue('D18', "Billing Contact Name");
                 $spreadsheet->getActiveSheet()->getStyle('E18')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E18')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('E18', $company['billing_contact_name']);

                 //-- Invoice Contact Phone
                 $sheet->setCellValue('D19', "Billing Contact Phone");
                 $spreadsheet->getActiveSheet()->getStyle('E19')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E19')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E19', $company['billing_contact_phone']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D20', "Billing Contact Email");
                 $spreadsheet->getActiveSheet()->getStyle('E20')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E20')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E20', $company['billing_contact_email']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D21', "Billing Contact Email 2");
                 $spreadsheet->getActiveSheet()->getStyle('E21')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E21')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E21', $company['billing_contact_email2']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D22', "Billing Contact Email 3");
                 $spreadsheet->getActiveSheet()->getStyle('E22')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E22')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E22', $company['billing_contact_email3']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D23', "Billing Contact Email 4");
                 $spreadsheet->getActiveSheet()->getStyle('E23')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E23')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E23', $company['billing_contact_email4']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D24', "Billing Contact Email 5");
                 $spreadsheet->getActiveSheet()->getStyle('E24')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E24')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E24', $company['billing_contact_email5']);


//================================================================================
//  PREENROLLMENT CENSUS
//================================================================================

       if ($company_id!=0) {
        $spreadsheet->setActiveSheetIndex(1);

        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "PRE-ENROLLMENT CENSUS");

        //-- Company ID
        $spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('A2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('A2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A2', "Member ID (Social Security Number)");
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B2', "Dependent ID (Social Security Number)");
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('C2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C2', "Last Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D2', "First Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E2', "Middle Initial");

        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F2', "Relationship (Employee, Husband, Wife Son, Daughter)");

        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G2', "Date of Birth");

        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('H2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H2', "Gender (M/F)");

        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('I2')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I2', "Marital Status - S/M");

        $sql="select * from nua_preenrollment_census where company_id = " . $company_id . " order by last_name, first_name";
        $x=$X->sql($sql);
        $row=2;
        foreach($x as $q) {
      $row++;
      $cell="A".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['social_security_number']);
      $cell="B".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['dependent_social_security_number']);
      $cell="C".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['last_name']);
      $cell="D".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['first_name']);
      $cell="E".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['middle_name']);
      $cell="F".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['relationship']);
      $cell="G".$row;
     $spreadsheet->getActiveSheet()->setCellValue($cell, $q['date_of_birth']);
      $cell="H".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['gender']);
      $cell="I".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['marital_status']);

        }
}

if ($company_id!=0) {
if ($quoted_plan=="Y") {
                $spreadsheet->setActiveSheetIndex(2);
                 //-- Column Widths
                 $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32);
                 $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);
                 $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);
                 //-- Company Name Title
                 //
                 $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
                 $spreadsheet->getActiveSheet()->setCellValue('A1', "Quoted Plans");
                 $spreadsheet->getActiveSheet()->getStyle('A2')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('A2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('A2', "Plan Name");
                 $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('B2', "APA");
                 $spreadsheet->getActiveSheet()->getStyle('C2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('C2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('C2', "EE Price");
                 $spreadsheet->getActiveSheet()->getStyle('D2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('D2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('D2', "EES Price");
                 $spreadsheet->getActiveSheet()->getStyle('E2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('E2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('E2', "EEC Price");
                 $spreadsheet->getActiveSheet()->getStyle('F2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('F2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('F2', "FAM Price");
                 $spreadsheet->getActiveSheet()->getStyle('G2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('G2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('G2', "Accepted");
$sql="select * from nua_quoted_plan where company_id = " . $company_id . " order by plan_type, plan_code";
$p=$X->sql($sql);
$row=2;
foreach ($p as $q) {
      $row++;
      $cell="A".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['plan_code']);
      $cell="B".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['APA_CODE']);
      $cell="C".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['ee_price']);
      $cell="D".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['ees_price']);
     $cell="E".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['eec_price']);
      $cell="F".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['fam_price']);
      $cell="G".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['accepted']);
}
} // QUOTED PLAN
} // COPMANY !=0

if ($company_id!=0) {
if ($company_plan=="Y") {

                $spreadsheet->setActiveSheetIndex(3);
                 //-- Column Widths
                 $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32);
                 $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);
                 $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);
                 //-- Company Name Title
                 //
                 $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
                 $spreadsheet->getActiveSheet()->setCellValue('A1', "Active Plans");
                 $spreadsheet->getActiveSheet()->getStyle('A2')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('A2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('A2', "Plan Name");
                 $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('B2', "APA");
                 $spreadsheet->getActiveSheet()->getStyle('C2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('C2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('C2', "EE Price");
                 $spreadsheet->getActiveSheet()->getStyle('D2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('D2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('D2', "EES Price");
                 $spreadsheet->getActiveSheet()->getStyle('E2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('E2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('E2', "EEC Price");
                 $spreadsheet->getActiveSheet()->getStyle('F2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('F2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('F2', "FAM Price");
                 $spreadsheet->getActiveSheet()->getStyle('G2')->applyFromArray($styleArrayBold);
                $spreadsheet->getActiveSheet()->getStyle('G2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('G2', "Plan Type");

$sql="select * from nua_company_plan where company_id = " . $company_id . " order by plan_type, plan_code";
$p=$X->sql($sql);
$row=2;
foreach ($p as $q) {
      $row++;
      $cell="A".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['plan_code']);
      $cell="B".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['APA_CODE']);
      $cell="C".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['ee_price']);
      $cell="D".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['ees_price']);
     $cell="E".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['eec_price']);
      $cell="F".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['fam_price']);
      $cell="G".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['plan_type']);
}
}
}

if ($company_id!=0) {
if ($company_plan=="Y") {

        $spreadsheet->setActiveSheetIndex(4);
        $spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(100);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Month");

        $spreadsheet->getActiveSheet()->getColumnDimension('AA')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('AA1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('AA1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
        $spreadsheet->getActiveSheet()->getStyle('AA1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AA1', "End");

        $spreadsheet->getActiveSheet()->getColumnDimension('AB')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('AB1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('AB1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
        $spreadsheet->getActiveSheet()->getStyle('AB1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AB1', "Employee Code (Leave Blank)");

        $spreadsheet->getActiveSheet()->getColumnDimension('AC')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('AC1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('AC1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
        $spreadsheet->getActiveSheet()->getStyle('AC1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AC1', "Dependent Code (Leave Blank)");

        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B1', "Social Security Number");

        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C1', "Dependent SSN");

        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D1', "Last Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E1', "First Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F1', "Middle Initial");

        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G1', "Relationship (Employee, Husband, Wife Son, Daughter)");

        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H1', "Date of Birth");
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I1', "Gender (M/F)");

        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('J1', "Marital Status");

        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('K1', "Address");

        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('L1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('L1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('L1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('L1', "Suite / Apt");

        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('M1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('M1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('M1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('M1', "City");

        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('N1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('N1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('N1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('N1', "State");

        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('O1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('O1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('O1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('O1', "Zipcode");

        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('P1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('P1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('P1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('P1', "Email (Provide a personal email)");

        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Q1', "Contact Phone (Required if no email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('R1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('R1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('R1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('R1', "Date of Hire");

        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('S1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('S1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('S1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('S1', "Work Status (A (Full Time) / PT (Part Time)");

        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('T1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('T1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('T1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('T1', "Effective Date (If blank defaults to first day of month in column A)");
// MEDICAL

        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('U1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('U1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('U1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('U1', "Medical Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('V1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('V1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('V1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('V1', "Coverage Election (EE, ES, EC, FAM)");
// DENTAL

        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('W1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('W1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('W1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('W1', "Dental Plan (Blank if none)");

        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('X1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('X1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('X1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('X1', "Dental Coverage (EE,ES,EC,FAM)");
// VISION
//
        $spreadsheet->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Y1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Y1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Y1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Y1', "Vision Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('Z')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Z1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Z1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Z1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Z1', "Vision Coverage");
//--
//-- This gets Everyone
//--
	$month=$this->getMonthId();

        $sql="select employee_id, employee_code, min(month_id) as min_month_id, max(month_id) as max_month_id, last_name, first_name from ";
        $sql.=" nua_monthly_member_census where company_id = " . $company_id;
        $sql.=" and dependent_code = '' group by employee_id, employee_code, last_name, first_name order by last_name, first_name";

        $d=$X->sql($sql);
        $row=1;

        $cid="";
        foreach($d as $e) {

            $sql="select id, employee_id from nua_monthly_member_census where employee_id = " . $e['employee_id'] . " and ";
            $sql.=" employee_code = '" . $e['employee_code'] . "' and dependent_code = '' and last_name = '" . str_replace("'","''",$e['last_name']) . "' and ";
            $sql.=" first_name = '" . str_replace("'","''",$e['first_name']) . "'";
            $x=$X->sql($sql);
            $cid="";
	    foreach($x as $y) {
                   $cid=$y['id'] . ",";
            }

            $sql="select * from nua_employee where id = " . $e['employee_id'];
            $x=$X->sql($sql);
	    if (sizeof($x)>0) {
		    $employee=$x[0];
	    } else {
		    $employee=array();
            }
            $sql="select * from nua_monthly_member_census where company_id = " . $company_id . " and employee_id = " . $e['employee_id'] . " and ";
            $sql.=" dependent_code = '' and month_id = '" . $e['max_month_id'] . "'";
            $hh=$X->sql($sql);
            $census=$hh[0];

            $sql="select * from nua_monthly_member_terminations where company_id = " . $company_id . " and employee_id = " . $e['employee_id'] . " and ";
            $sql.=" dependent_code = '' order by month_id";
            $ii=$X->sql($sql);
            $term_dt="";
            foreach($ii as $jj) $term_dt=$jj['term_dt'];

            $row++;
            $cell="A".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['min_month_id']);
            $cell="AA".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if ($e['max_month_id']!=$month) {
                    $spreadsheet->getActiveSheet()->setCellValue($cell, $e['max_month_id']);
		}

            $cell="AB".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $census['employee_code']);
            $cell="AC".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $census['dependent_code']);
            $cell="B".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['social_security_number'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['ssn']);
		}
            $cell="C".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
            $cell="D".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['last_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['last_name']);
		}
            $cell="E".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['first_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['first_name']);
		}
            $cell="F".$row;
                    $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['middle_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['middle_initial']);
		}
            $cell="G".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="H".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['date_of_birth'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['dob']);
		}
            $cell="I".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['gender'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['gender']);
		}
            $cell="J".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
               if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['marital_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
		} else {
			
		}
            $cell="K".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['address'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['address']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['address']);
		}
            $cell="L".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['suite'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['suite']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['address2']);
		}
            $cell="M".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['city'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['city']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['city']);
		}
            $cell="N".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['state'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['state']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['state']);
		}
            $cell="O".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['zip'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['zip']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['zip']);
		}
            $cell="P".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['email'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['email']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['email']);
		}
            $cell="Q".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['phone'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['phone']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['phone']);
		}
            $cell="R".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['date_hired'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_hired']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['hire_date']);
		}
            $cell="S".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		if (isset($employee['work_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['work_status']);
		} else {
			//$spreadsheet->getActiveSheet()->setCellValue($cell, $census['work_status```']);
		}
            $cell="T".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $census['eff_dt']);
		/*
                         $cell="X".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $term_dt);
		 */
            $medical_plan="";
            $medical_coverage="";
            $dental_plan="";
            $dental_coverage="";
            $vision_plan="";
            $vision_coverage="";
            $medical_price="";
            $dental_price="";
            $vision_price="";
            foreach($hh as $ii) {
                if ($ii['plan_type']=="*MEDICAL*" ) {
                     $medical_plan=$ii['client_plan'];
                     $medical_coverage=$ii['coverage_level'];
                     $medical_price=$ii['coverage_price'];
                }
                if ($ii['plan_type']=="*DENTAL*" ) {
                     $dental_plan=$ii['client_plan'];
                     $dental_coverage=$ii['coverage_level'];
                     $dental_price=$ii['coverage_price'];
                }
                if ($ii['plan_type']=="*VISION*" ) {
                     $vision_plan=$ii['client_plan'];
                     $vision_coverage=$ii['coverage_level'];
                     $vision_price=$ii['coverage_price'];
                }
            }
            $cell="U".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $medical_plan);
            $cell="V".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $medical_coverage);
	/*    
		$cell="AA".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $medical_price);
	 */      
		$cell="W".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $dental_plan);
            $cell="X".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
   
		$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_coverage);
/*/    
		$cell="AD".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $dental_price);
*/     
		$cell="Y".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $vision_plan);
            $cell="Z".$row;
                    $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $vision_coverage);
/*     
		$cell="AG".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $vision_price);
            $cell="AH".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                $spreadsheet->getActiveSheet()->setCellValue($cell, $cid);
            $cell="AI".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
                if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
                $spreadsheet->getActiveSheet()->setCellValue($cell, $census['error_msg']);
 */
            $sql="select * from nua_employee_dependent where employee_id = " . $e['employee_id'] . " order by dependent_id";
            $z=$X->sql($sql);
            foreach($z as $ff) {
                $row++;
                $cell="AB".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $e['employee_code']);
                $cell="AC".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		     if ($ff['dependent_id']=="") $ff['dependnt_id']=$e['employee_code'] . "m";
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['dependent_id']);
                $cell="B".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['social_security_number'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['ssn']);
		}
                $cell="C".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['social_security_number']);
                $cell="D".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['last_name']);
                $cell="E".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['first_name']);
                $cell="F".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['middle_name']);
                $cell="G".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['relationship']);
                $cell="H".$row;
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['date_of_birth']);
                $cell="I".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                        $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['gender']);
            }

        }
}
}

$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
if ($company_id==0) {
   $filename="NuAxess_Quoting_Template.xlsx";
} else {
   $filename=$company_id . "_" . str_replace(' ','_',str_replace('#',' ',$company['company_name'])) . ".xlsx";
}
if ($display=="F") {
$file="sheets/" . $filename;
$writer->save($file);
return $file;
} else {
    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();
}
}

function getCommissionReport($billingCycle)
{
	$X=new XRDB();


$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Agency Commission Report');

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
	$spreadsheet->getActiveSheet()->setCellValue('A1', "Agency ID");

	$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(40);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('B1', "Agency Name");

	$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('C1', "Company ID");
	
	$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(40);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('D1', "Company Name");

	$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('E1', "Plan Type");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('F1', "Rate");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('G1', "Commssion Type");
	 	
	$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('H1', "Basis");
	 	
	$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('I1', "Invoice Paid");
	 	
	$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('J1', "Commission");
	 	
	$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('K1', "Commission Paid");
	 	
	$sql="select id, org_name from nua_agency order by org_name ";
	$eft=$X->sql($sql);

        $row=1;

	
        foreach($eft as $eftRow) {
		
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['id']);
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['org_name']);
		$commission_total=0;
            $row++;
	$sql="select nua_agency_company.company_id as id, nua_company.company_name as company_name, plan_type, commission_rate, commission_type from ";
		$sql.="nua_agency_company, nua_company where nua_agency_company.agency_id = " . $eftRow['id']; 
		$sql.=" and nua_agency_company.company_id = nua_company.id order by nua_company.company_name";

		$c=$X->sql($sql);

		foreach($c as $comp) {
	
	

		if ($comp['commission_type']=="FLAT") {
			$sql="select count(*) as c from nua_monthly_member_census where month_id = '" . $billingCycle . "' and company_id = " . $comp['id'];
			$sql.=" and plan_type = '" . $comp['plan_type'] . "' and dependent_code = ''";
			$b=$X->sql($sql);
			$basis=floatval($b[0]['c']);
		} else {
			$sql="select coverage_price from nua_monthly_member_census where month_id = '" . $billingCycle . "' and company_id = " . $comp['id'];
			$sql.=" and plan_type = '" . $comp['plan_type'] . "' and dependent_code = ''";
			$b=$X->sql($sql);
			$total=0;
			foreach($b as $c) {
                             $total+=floatval($c['coverage_price']);
			}
			$basis=$total;
		}
		$sql="select id, paid from nua_company_invoice where company_id = " . $comp['id'] . " and month_id = '" . $billingCycle . "'";
		$gg=$X->sql($sql);
		if (sizeof($gg)==0) {
			$p="No";
		} else {
		        $p=$gg[0]['paid'];
			if ($p=='N') $p="No";
			if ($p=='Y') $p="Yes";
		}
		$commission=0;
		if ($p=='No') {
			$commission=0;
			$commission_paid="No";
		} else {
		     $commission=$basis*$comp['commission_rate'];
		     if (sizeof($gg)>0) {
		         $sql="select count(*) as c from nua_agency_paid where agency_id = " . $eftRow['id'];
		         $sql.=" and invoice_id = " . $gg[0]['id'] . " and plan_type = '" . $comp['plan_type'] . "'";
		         $h=$X->sql($sql);
		         if ($h[0]['c']>0) { 
			     $commission_paid="Yes";
		         } else {
			     $commission_paid="No";
		         }
		     } else {
			     $commission_paid="No";
	             }
		}
		if ($commission_paid=="No") {
			$commission_total+=$commission;
			$row++;
                         $cell="C".$row;
                         $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	                 $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['id']);
                $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['company_name']);
                $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['plan_type']);
                $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['commission_rate']);
		$comm_rate=floatval($comp['commission_rate']);
                $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['commission_type']);
                $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $basis);
                $cell="I".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $p);
                $cell="J".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, "$" . number_format($commission,2));
                $cell="K".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $commission_paid);
		}

		}
		$row++;
                        $cell="J".$row;
                        $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	                $spreadsheet->getActiveSheet()->setCellValue($cell, "$" . number_format($commission_total,2));
		$row++;

	}


$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Commission_Report.xlsx";

    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();

}

function getAgencyCommissionReport($billingCycle,$agency_id)
{
	$X=new XRDB();


$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Agency Commission Report');

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
	$spreadsheet->getActiveSheet()->setCellValue('A1', "Agency ID");

	$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(40);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('B1', "Agency Name");

	$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('C1', "Company ID");
	
	$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(40);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('D1', "Company Name");

	$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('E1', "Plan Type");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('F1', "Rate");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('G1', "Commssion Type");
	 	
	$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('H1', "Basis");
	 	
	$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('I1', "Invoice Paid");
	 	
	$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('J1', "Commission");
	 	
	$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('K1', "Commission Paid");
	 	
	$sql="select id, org_name from nua_agency where id = " . $agency_id . " order by org_name ";
	$eft=$X->sql($sql);

        $row=1;

        foreach($eft as $eftRow) {
		
		$commission_total=0;
            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['id']);
	
		
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['org_name']);
	
	$sql="select nua_agency_company.company_id as id, nua_company.company_name as company_name, plan_type, commission_rate, commission_type from ";
		$sql.="nua_agency_company, nua_company where nua_agency_company.agency_id = " . $eftRow['id']; 
		$sql.=" and nua_agency_company.company_id = nua_company.id order by nua_company.company_name";

		$c=$X->sql($sql);

		foreach($c as $comp) {
	
	

		if ($comp['commission_type']=="FLAT") {
			$sql="select count(*) as c from nua_monthly_member_census where month_id = '" . $billingCycle . "' and company_id = " . $comp['id'];
			$sql.=" and plan_type = '" . $comp['plan_type'] . "' and dependent_code = ''";
			$b=$X->sql($sql);
			$basis=floatval($b[0]['c']);
		} else {
			$sql="select coverage_price from nua_monthly_member_census where month_id = '" . $billingCycle . "' and company_id = " . $comp['id'];
			$sql.=" and plan_type = '" . $comp['plan_type'] . "' and dependent_code = ''";
			$b=$X->sql($sql);
			$total=0;
			foreach($b as $c) {
                             $total+=floatval($c['coverage_price']);
			}
			$basis=$total;
		}
		$sql="select id, paid from nua_company_invoice where company_id = " . $comp['id'] . " and month_id = '" . $billingCycle . "'";
		$gg=$X->sql($sql);
		if (sizeof($gg)==0) {
			$p="No";
		} else {
		        $p=$gg[0]['paid'];
			if ($p=='N') $p="No";
			if ($p=='Y') $p="Yes";
		}
		$commission=0;
		if ($p=='No') {
			$commission=0;
			$commission_paid="No";
		} else {
		     $commission=$basis*$comp['commission_rate'];
		     if (sizeof($gg)>0) {
		         $sql="select count(*) as c from nua_agency_paid where agency_id = " . $eftRow['id'];
		         $sql.=" and invoice_id = " . $gg[0]['id'] . " and plan_type = '" . $comp['plan_type'] . "'";
		         $h=$X->sql($sql);
		         if ($h[0]['c']>0) { 
			     $commission_paid="Yes";
		         } else {
			     $commission_paid="No";
		         }
		     } else {
			     $commission_paid="No";
	             }
		}
		if ($commission_paid=="No") {
			$commission_total+=$commission;
			$row++;
                         $cell="C".$row;
                         $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	                 $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['id']);
                $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['company_name']);
                $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['plan_type']);
                $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['commission_rate']);
		$comm_rate=floatval($comp['commission_rate']);
                $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['commission_type']);
                $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $basis);
                $cell="I".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $p);
                $cell="J".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, "$" . number_format($commission,2));
                $cell="K".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $commission_paid);
		}

		}
		$row++;
                        $cell="J".$row;
                        $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	                $spreadsheet->getActiveSheet()->setCellValue($cell, "$" . number_format($commission_total,2));
		$row++;

	}


$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Commission_Report.xlsx";

    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();

}


function payCommissions($billingCycle, $agency_id,$payment_date,$payment_id=0)
{
	$X=new DB();


	$sql="select id, org_name from nua_agency where id = " . $agency_id . " order by org_name ";
	$eft=$X->sql($sql);

        $row=1;

        foreach($eft as $eftRow) {
		
	    $commission_total=0;
            $row++;
	
            $sql="select nua_agency_company.company_id as id, nua_company.company_name as company_name, plan_type, ";
            $sql.=" commission_rate, commission_type from ";
    	    $sql.="nua_agency_company, nua_company where nua_agency_company.agency_id = " . $eftRow['id']; 
	    $sql.=" and nua_agency_company.company_id = nua_company.id order by nua_company.company_name";

	    echo $sql;
	    $c=$X->sql($sql);

	    print_r($c);

		foreach($c as $comp) {
			$row++;
	
		        if ($comp['commission_type']=="FLAT") {
		           	$sql="select count(*) as c from nua_monthly_member_census where month_id = '" . $billingCycle . "' and company_id = " . $comp['id'];
			        $sql.=" and plan_type = '" . $comp['plan_type'] . "' and dependent_code = ''";
	         		$b=$X->sql($sql);
			        $basis=floatval($b[0]['c']);
		} else {
		   	        $sql="select coverage_price from nua_monthly_member_census where month_id = '" . $billingCycle . "' and company_id = " . $comp['id'];
			        $sql.=" and plan_type = '" . $comp['plan_type'] . "' and dependent_code = ''";
			        $b=$X->sql($sql);
			        $total=0;
			        foreach($b as $c) {
                                    $total+=floatval($c['coverage_price']);
			        }
			        $basis=$total;
		        }
		        $sql="select id, paid from nua_company_invoice where company_id = " . $comp['id'] . " and month_id = '" . $billingCycle . "'";
		        $gg=$X->sql($sql);
			print_r($gg);
		        if (sizeof($gg)==0) {
			     $p="No";
		         } else {
		           $p=$gg[0]['paid'];
			   if ($p=='N') $p="No";
			   if ($p=='Par') $p="No";
			   if ($p=='Y') $p="Yes";
	          	}
		        $commission=0;
		        if ($p=='No') {
			    $commission=0;
			    $commission_paid="No";
		        } else {
		            if (sizeof($gg)>0) {
		                $sql="select count(*) as c from nua_agency_paid where agency_id = " . $agency_id;
		                $sql.=" and invoice_id = " . $gg[0]['id'] . " and plan_type = '" . $comp['plan_type'] . "'";
		                $h=$X->sql($sql);
				print_r($h);
		                if ($h[0]['c']>0) { 
			            $commission_paid="Yes";
		                } else {
			            $commission_paid="No";
				    $ppp=array();
				    $ppp['table_name']="nua_agency_paid";
				    $ppp['action']="insert";
				    $ppp['agency_id']=$agency_id;
				    $ppp['invoice_id']=$gg[0]['id'];
				    $ppp['plan_type']=$comp['plan_type'];
				    $ppp['payment_id']=$payment_id;
				    $ppp['payment_date']=$payment_date;
				    print_r($ppp);
				    $X->post($ppp);
		                }
		                } 
		}
		}
	}
}

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
	$spreadsheet->getActiveSheet()->setCellValue('D1', "Due Date");

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
		
	$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('G1', "Amount Due");
		
	$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('H1', "Days Overdue");
	 	
	$sql="select company_id, company_name, invoice_number, grand_total, payment_amount, billing_date, ";
	$sql.=" due_date, amount_left, paid, month_id from nua_company_invoice ";
	$sql.=" where paid in ('No','N','Par') and amount_left <> 0 and company_id in ";
	$sql.=" (select id from nua_company where org_id not in (17,19,20,21,22,23,24,25))  and month_id <= '".$billingCycle."' and month_id >= '2022-07' order by company_name, month_id";
	$eft=$X->sql($sql);

        $row=1;

        foreach($eft as $eftRow) {
		
		if ($eftRow['month_id']=='2022-01') $dd="2021-12-31";
		if ($eftRow['month_id']=='2022-02') $dd="2022-01-31";
		if ($eftRow['month_id']=='2022-03') $dd="2022-02-28";
		if ($eftRow['month_id']=='2022-04') $dd="2022-03-31";
		if ($eftRow['month_id']=='2022-05') $dd="2022-04-30";
		if ($eftRow['month_id']=='2022-06') $dd="2022-05-31";
		if ($eftRow['month_id']=='2022-07') $dd="2022-06-30";
		if ($eftRow['month_id']=='2022-08') $dd="2022-07-31";
		if ($eftRow['month_id']=='2022-09') $dd="2022-08-31";
		if ($eftRow['month_id']=='2022-10') $dd="2022-09-30";
		if ($eftRow['month_id']=='2022-11') $dd="2022-10-31";
		if ($eftRow['month_id']=='2022-12') $dd="2022-11-31";

		$sql="SELECT DATEDIFF(CURDATE(), '" . $dd . "') AS days FROM dual";
		$days=$X->sql($sql);
		$d=$days[0]['days'];
	 
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
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['due_date']);
	
	
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['grand_total']);

		
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['payment_amount']);
		
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['amount_left']);

            $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $d);
	}


$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Open_Invoice_Report.xlsx";

    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();

}

				    print_r($ppp);
				    $X->post($ppp);
		                }
		                } 
		}
		}
	}
}

	/*
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
	foreach($plan_type as $t) {

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

		$total=$data['c'];

		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and coverage_level in ('EE','SI') and client_plan = '" . $e['client_plan']  . "'";

	        $data=$X->sql($sql);

		$ee=$data['c'];

		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and coverage_level in ('ES','ES2') and client_plan = '" . $e['client_plan']  . "'";
	        $data=$X->sql($sql);

		$es=$data['c'];

		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and coverage_level in ('EC','EC2') and client_plan = '" . $e['client_plan']  . "'";
	        $data=$X->sql($sql);

		$ec=$data['c'];

		$sql="select count(*) as c from ";
	        $sql.=" nua_monthly_member_census where month_id = '" . $billingCycle . "' ";
	        $sql.=" and company_id in (select id from nua_company where org_id in (17)) and plan_type = '*MEDICAL*' ";
	        $sql.=" and coverage_level in ('FAM','FA') and client_plan = '" . $e['client_plan']  . "'";
	        $data=$X->sql($sql);

		$fam=$data['c'];

            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $plan_type);
	
		
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $t['client_plan']);
	
	
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
	}
$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Enrollment_Stats_Invoice_Report.xlsx";

    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();

}

	 */
function getPaymentWorkbookExcel($billingCycle)
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
	$sql.=" where paid in ('No','N','Par') and amount_left <> 0 and company_id in (select id from nua_company where org_id not in (17,19,20,21,22,23,24,25) and flag_eft not in ('Y','Yes','YES','y'))  and month_id = '".$billingCycle."' order by company_name";
	$eft=$X->sql($sql);

        $row=1;

        foreach($eft as $eftRow) {
		
		if ($eftRow['month_id']=='2022-01') $dd="2021-12-31";
		if ($eftRow['month_id']=='2022-02') $dd="2022-01-31";
		if ($eftRow['month_id']=='2022-03') $dd="2022-02-28";
		if ($eftRow['month_id']=='2022-04') $dd="2022-03-31";
		if ($eftRow['month_id']=='2022-05') $dd="2022-04-30";
		if ($eftRow['month_id']=='2022-06') $dd="2022-05-31";
		if ($eftRow['month_id']=='2022-07') $dd="2022-06-30";
		if ($eftRow['month_id']=='2022-08') $dd="2022-07-31";
		if ($eftRow['month_id']=='2022-09') $dd="2022-08-31";
		if ($eftRow['month_id']=='2022-10') $dd="2022-09-30";
		if ($eftRow['month_id']=='2022-11') $dd="2022-10-31";
		if ($eftRow['month_id']=='2022-12') $dd="2022-11-31";

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
	
//            $cell="D".$row;
//                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
//	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['due_date']);
	
	
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['grand_total']);

		
//            $cell="F".$row;
//                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
//	        $spreadsheet->getActiveSheet()->setCellValue($cell, $eftRow['payment_amount']);
		
	}


$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Payment_Workbook.xlsx";

    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
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

function getAPAAdditions($data,$display='F') {

	$X=new XRDB();

$styleArrayEnrollment = [
    'font' => [
        'bold' => true,
    ],
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

$styleArrayBold = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
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

$styleArrayHead = [
    'font' => [
        'bold' => true,
        'size' => '24',
    ],
];

$styleArray = [
    'font' => [
        'bold' => true,
        'size' => '12',
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

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('APA Additions');

$quoted_plan="N";
$company_plan="N";


      $spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(100);

        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Company Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B1', "Social Security Number");

        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C1', "Dependent SSN");

        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D1', "Last Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E1', "First Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F1', "Middle Initial");

        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G1', "Relationship (Employee, Husband, Wife Son, Daughter)");

        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H1', "Date of Birth");
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I1', "Gender (M/F)");

        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('J1', "Marital Status");

        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('K1', "Address");

        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('L1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('L1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('L1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('L1', "Suite / Apt");

        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('M1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('M1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('M1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('M1', "City");

        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('N1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('N1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('N1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('N1', "State");

        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('O1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('O1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('O1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('O1', "Zipcode");

        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('P1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('P1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('P1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('P1', "Email (Provide a personal email)");

        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Q1', "Contact Phone (Required if no email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('R1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('R1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('R1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('R1', "Date of Hire");

        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('S1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('S1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('S1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('S1', "Work Status (A (Full Time) / PT (Part Time)");

        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('T1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('T1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('T1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('T1', "Effective Date");

        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('U1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('U1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('U1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('U1', "Medical Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('V1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('V1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('V1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('V1', "Coverage Election (EE, ES, EC, FAM)");

        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('W1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('W1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('W1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('W1', "Client Plan");

        $row=1;

        $cid="";
        foreach($data as $e) {

	    $sql="select * from nua_employee where id = " . $e['employee_id'];
	    $ee=$X->sql($sql);
	    $employee=$ee[0];
	    $sql="select * from nua_company where id = " . $e['company_id'];
	    $ec=$X->sql($sql);
	    $comp=$ec[0];
            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['company_name']);
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['social_security_number'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['ssn']);
		}
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['last_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['last_name']);
		}
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['first_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['first_name']);
		}
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['middle_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['middle_initial']);
		}
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['date_of_birth'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['dob']);
		}
            $cell="I".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['gender'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['gender']);
		}
            $cell="J".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['marital_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
		} else {
			
		}
            $cell="K".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['address'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['address']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['address']);
		}
            $cell="L".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['suite'])) {
		     $spreadsheet->getActiveSheet()->setCellValue($cell, $employee['suite']);
		} else {
		     $spreadsheet->getActiveSheet()->setCellValue($cell, $census['address2']);
		}
            $cell="M".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['city'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['city']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['city']);
		}
            $cell="N".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['state'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['state']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['state']);
		}
            $cell="O".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['zip'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['zip']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['zip']);
		}
            $cell="P".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['email'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['email']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['email']);
		}
            $cell="Q".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['phone'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['phone']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['phone']);
		}
            $cell="R".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['hire_date'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['hire_date']);
		} else {
		}
            $cell="S".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['work_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['work_status']);
		} else {
			//$spreadsheet->getActiveSheet()->setCellValue($cell, $census['work_status```']);
		}
            $cell="T".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['eff_dt']);
		$sql="select APA_CODE from nua_company_plan where plan_code = '" . $e['client_plan'] . "' and company_id = " . $employee['company_id'];
		$z=$X->sql($sql);
		if (sizeof($z) > 0) { 
                     $cell="U".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	             $spreadsheet->getActiveSheet()->setCellValue($cell, $z[0]['APA_CODE']);
		} 
                $cell="V".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['coverage_level']);
                $cell="W".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $e['client_plan']);

		if ($e['coverage_level']!='EE') {
                 $sql="select * from nua_employee_dependent where employee_id = " . $e['employee_id'] . " order by dependent_id";
                 $z=$X->sql($sql);
                 foreach($z as $ff) {
                    $row++;
                    $cell="B".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
                $cell="C".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['social_security_number']);
                $cell="D".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['last_name']);
                $cell="E".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['first_name']);
                $cell="F".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['middle_name']);
                $cell="G".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['relationship']);
                $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['date_of_birth']);
                $cell="I".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['gender']);
	    }
		}
	}
$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="APA_Enrollment_File.xlsx";
if ($display=="F") {
$file="sheets/" . $filename;
$writer->save($file);
return $file;
} else {
    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();
}
 
}

function getGuardianAdditions($data,$display='F') {

	$X=new XRDB();

$styleArrayEnrollment = [
    'font' => [
        'bold' => true,
    ],
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

$styleArrayBold = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
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

$styleArrayHead = [
    'font' => [
        'bold' => true,
        'size' => '24',
    ],
];

$styleArray = [
    'font' => [
        'bold' => true,
        'size' => '12',
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

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Additions');

$quoted_plan="N";
$company_plan="N";


      $spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(100);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Company Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B1', "Social Security Number");

        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C1', "Dependent SSN");

        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D1', "Last Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E1', "First Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F1', "Middle Initial");

        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G1', "Relationship (Employee, Husband, Wife Son, Daughter)");

        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H1', "Date of Birth");
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I1', "Gender (M/F)");

        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('J1', "Marital Status");

        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('K1', "Address");

        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('L1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('L1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('L1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('L1', "Suite / Apt");

        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('M1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('M1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('M1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('M1', "City");

        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('N1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('N1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('N1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('N1', "State");

        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('O1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('O1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('O1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('O1', "Zipcode");

        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('P1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('P1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('P1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('P1', "Email (Provide a personal email)");

        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Q1', "Contact Phone (Required if no email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('R1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('R1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('R1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('R1', "Date of Hire");

        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('S1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('S1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('S1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('S1', "Work Status (A (Full Time) / PT (Part Time)");

        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('T1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('T1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('T1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('T1', "Effective Date");

        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('U1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('U1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('U1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('U1', "Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('V1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('V1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('V1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('V1', "Coverage Election (EE, ES, EC, FAM)");

        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('W1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('W1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('W1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('W1', "Coverage");

        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('X1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('X1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('X1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('X1', "Spouse Coverage");

        $spreadsheet->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Y1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Y1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Y1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Y1', "Dependent Coverage");

        $row=1;

        $cid="";
        foreach($data as $e) {

	    $sql="select * from nua_employee where id = " . $e['employee_id'];
	    $ee=$X->sql($sql);
	    $employee=$ee[0];
	    $sql="select * from nua_company where id = " . $e['company_id'];
	    $ec=$X->sql($sql);
	    $comp=$ec[0];
            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['company_name']);
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['social_security_number'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['ssn']);
		}
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['last_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['last_name']);
		}
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['first_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['first_name']);
		}
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['middle_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['middle_initial']);
		}
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['date_of_birth'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['dob']);
		}
            $cell="I".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['gender'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['gender']);
		}
            $cell="J".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['marital_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
		} else {
			
		}
            $cell="K".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['address'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['address']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['address']);
		}
            $cell="L".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['suite'])) {
		     $spreadsheet->getActiveSheet()->setCellValue($cell, $employee['suite']);
		} else {
		     $spreadsheet->getActiveSheet()->setCellValue($cell, $census['address2']);
		}
            $cell="M".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['city'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['city']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['city']);
		}
            $cell="N".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['state'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['state']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['state']);
		}
            $cell="O".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['zip'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['zip']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['zip']);
		}
            $cell="P".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['email'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['email']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['email']);
		}
            $cell="Q".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['phone'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['phone']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['phone']);
		}
            $cell="R".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['hire_date'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['hire_date']);
		} else {
		}
            $cell="S".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['work_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['work_status']);
		} else {
			//$spreadsheet->getActiveSheet()->setCellValue($cell, $census['work_status```']);
		}
            $cell="T".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['eff_dt']);
            $cell="U".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['client_plan']);
                $cell="V".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['coverage_level']);

                if ($e['plan_type']=="*LIFE*"||$e['plan_type']=="*ADD*"||$e['plan_type']=="*VISION*") {

		    $sql="select * from inf_client_employee_plan where employeeId = '" . $e['employee_code'] . "' and planId = '" . $e['client_plan'] . "'";
		    $t=$X->sql($sql);
		    if (sizeof($t)>0) {
                       $cell="W".$row;
                       $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	               $spreadsheet->getActiveSheet()->setCellValue($cell, $t[0]['coverage']);
                       $cell="X".$row;
                       $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	               $spreadsheet->getActiveSheet()->setCellValue($cell, $t[0]['spouseCoverage']);
                       $cell="Y".$row;
                       $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	               $spreadsheet->getActiveSheet()->setCellValue($cell, $t[0]['dependentCoverage']);
		    }
		}


		if ($e['coverage_level']!='EE') {
                 $sql="select * from nua_employee_dependent where employee_id = " . $e['employee_id'] . " order by dependent_id";
                 $z=$X->sql($sql);
                 foreach($z as $ff) {
                    $row++;
                    $cell="B".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
                $cell="C".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['social_security_number']);
                $cell="D".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['last_name']);
                $cell="E".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['first_name']);
                $cell="F".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['middle_name']);
                $cell="G".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['relationship']);
                $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['date_of_birth']);
                $cell="I".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['gender']);
	    }
		}
	}
$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Guardian_Enrollment_File.xlsx";
if ($display=="F") {
$file="sheets/" . $filename;
$writer->save($file);
return $file;
} else {
    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();
}
 
}


function getGuardianTerminations($data,$display='F') {

	$X=new XRDB();

$styleArrayEnrollment = [
    'font' => [
        'bold' => true,
    ],
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

$styleArrayBold = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
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

$styleArrayHead = [
    'font' => [
        'bold' => true,
        'size' => '24',
    ],
];

$styleArray = [
    'font' => [
        'bold' => true,
        'size' => '12',
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

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('APA Additions');

$quoted_plan="N";
$company_plan="N";


      $spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(100);

        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Company Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B1', "Social Security Number");

        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C1', "Dependent SSN");

        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D1', "Last Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E1', "First Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F1', "Middle Initial");

        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G1', "Relationship (Employee, Husband, Wife Son, Daughter)");

        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H1', "Date of Birth");
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I1', "Gender (M/F)");

        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('J1', "Marital Status");

        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('K1', "Address");

        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('L1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('L1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('L1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('L1', "Suite / Apt");

        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('M1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('M1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('M1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('M1', "City");

        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('N1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('N1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('N1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('N1', "State");

        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('O1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('O1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('O1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('O1', "Zipcode");

        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('P1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('P1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('P1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('P1', "Email (Provide a personal email)");

        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Q1', "Contact Phone (Required if no email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('R1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('R1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('R1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('R1', "Date of Hire");

        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('S1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('S1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('S1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('S1', "Work Status (A (Full Time) / PT (Part Time)");

        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('T1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('T1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('T1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('T1', "Effective Date");

        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('U1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('U1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('U1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('U1', "Termination Date");

        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('V1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('V1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('V1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('V1', "Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('W1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('W1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('W1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('W1', "Coverage Election (EE, ES, EC, FAM)");

        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('X1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('X1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('X1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('X1', "Coverage");

        $spreadsheet->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Y1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Y1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Y1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Y1', "Spouse Coverage");

        $spreadsheet->getActiveSheet()->getColumnDimension('Z')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Z1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Z1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Z1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Z1', "Dependent Coverage");

        $row=1;

        $cid="";
        foreach($data as $e) {

	    $sql="select * from nua_employee where id = " . $e['employee_id'];
	    $ee=$X->sql($sql);
	    $employee=$ee[0];
	    $sql="select * from nua_company where id = " . $e['company_id'];
	    $ec=$X->sql($sql);
	    $comp=$ec[0];
            $row++;
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['company_name']);
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['social_security_number'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['ssn']);
		}
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['last_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['last_name']);
		}
            $cell="E".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['first_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['first_name']);
		}
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['middle_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['middle_initial']);
		}
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['date_of_birth'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['dob']);
		}
            $cell="I".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['gender'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['gender']);
		}
            $cell="J".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['marital_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
		} else {
			
		}
            $cell="K".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['address'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['address']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['address']);
		}
            $cell="L".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['suite'])) {
		     $spreadsheet->getActiveSheet()->setCellValue($cell, $employee['suite']);
		} else {
		     $spreadsheet->getActiveSheet()->setCellValue($cell, $census['address2']);
		}
            $cell="M".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['city'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['city']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['city']);
		}
            $cell="N".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['state'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['state']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['state']);
		}
            $cell="O".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['zip'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['zip']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['zip']);
		}
            $cell="P".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['email'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['email']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['email']);
		}
            $cell="Q".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['phone'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['phone']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['phone']);
		}
            $cell="R".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['hire_date'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['hire_date']);
		} else {
		}
            $cell="S".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['work_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['work_status']);
		} else {
			//$spreadsheet->getActiveSheet()->setCellValue($cell, $census['work_status```']);
		}
            $cell="T".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['eff_dt']);
            $cell="U".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['term_dt']);
            $cell="V".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['client_plan']);
                $cell="W".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['coverage_level']);

		if ($e['plan_type']=='*VISION*'||$e['plan_type']=="*LIFE*"||$e['plan_type']=="*ADD*") {

		    $sql="select * from inf_client_employee_plan where employeeId = '" . $e['employee_code'] . "' and planId = '" . $e['client_plan'] . "'";
		    $t=$X->sql($sql);
		    if (sizeof($t)>0) {
                       $cell="X".$row;
                       $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	               $spreadsheet->getActiveSheet()->setCellValue($cell, $t[0]['coverage']);
                       $cell="Y".$row;
                       $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	               $spreadsheet->getActiveSheet()->setCellValue($cell, $t[0]['spouseCoverage']);
                       $cell="Z".$row;
                       $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	               $spreadsheet->getActiveSheet()->setCellValue($cell, $t[0]['dependentCoverage']);
		    }
		}


		if ($e['coverage_level']!='EE') {
                 $sql="select * from nua_employee_dependent where employee_id = " . $e['employee_id'] . " order by dependent_id";
                 $z=$X->sql($sql);
                 foreach($z as $ff) {
                    $row++;
                    $cell="B".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
                $cell="C".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['social_security_number']);
                $cell="D".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['last_name']);
                $cell="E".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['first_name']);
                $cell="F".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['middle_name']);
                $cell="G".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['relationship']);
                $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['date_of_birth']);
                $cell="I".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['gender']);
	    }
		}
	}
$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="Guardian_Termination_File.xlsx";
if ($display=="F") {
$file="sheets/" . $filename;
$writer->save($file);
return $file;
} else {
    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();
}
 
}


function getAPATerminations($data,$display='F') {

	$X=new XRDB();

$styleArrayEnrollment = [
    'font' => [
        'bold' => true,
    ],
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

$styleArrayBold = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
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

$styleArrayHead = [
    'font' => [
        'bold' => true,
        'size' => '24',
    ],
];

$styleArray = [
    'font' => [
        'bold' => true,
        'size' => '12',
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

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('APA Additions');

$quoted_plan="N";
$company_plan="N";


      $spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(100);

        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Company Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B1', "Social Security Number");

        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C1', "Dependent SSN");

        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D1', "Last Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E1', "First Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F1', "Middle Initial");

        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G1', "Relationship (Employee, Husband, Wife Son, Daughter)");

        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H1', "Date of Birth");
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I1', "Gender (M/F)");

        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('J1', "Marital Status");

        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('K1', "Address");

        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('L1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('L1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('L1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('L1', "Suite / Apt");

        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('M1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('M1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('M1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('M1', "City");

        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('N1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('N1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('N1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('N1', "State");

        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('O1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('O1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('O1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('O1', "Zipcode");

        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('P1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('P1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('P1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('P1', "Email (Provide a personal email)");

        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('Q1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Q1', "Contact Phone (Required if no email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('R1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('R1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('R1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('R1', "Date of Hire");

        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('S1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('S1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('S1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('S1', "Work Status (A (Full Time) / PT (Part Time)");

        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('T1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('T1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('T1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('T1', "Effective Date");

        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('U1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('U1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('U1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('U1', "Termination Date");

        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('V1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('V1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('V1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('V1', "Medical Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('W1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('W1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('W1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('W1', "Coverage Election (EE, ES, EC, FAM)");

        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('X1')->applyFromArray($styleArrayEnrollment);
        $spreadsheet->getActiveSheet()->getStyle('X1')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->getStyle('X1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('X1', "Client Plan");

        $row=1;

        $cid="";
        foreach($data as $e) {

	    $sql="select * from nua_employee where id = " . $e['employee_id'];
	    $ee=$X->sql($sql);
	    $employee=$ee[0];
	    $sql="select * from nua_company where id = " . $e['company_id'];
	    $ec=$X->sql($sql);
	    $comp=$ec[0];
            $cell="A".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $comp['company_name']);
            $row++;
            $cell="B".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['social_security_number'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['ssn']);
		}
            $cell="C".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
            $cell="D".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['last_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['last_name']);
		}
            $cell="F".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['first_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['first_name']);
		}
            $cell="G".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['middle_name'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['middle_initial']);
		}
            $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="I".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['date_of_birth'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['dob']);
		}
            $cell="J".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['gender'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['gender']);
		}
            $cell="K".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['marital_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
		} else {
			
		}
            $cell="L".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['address'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['address']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['address']);
		}
            $cell="M".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['suite'])) {
		     $spreadsheet->getActiveSheet()->setCellValue($cell, $employee['suite']);
		} else {
		     $spreadsheet->getActiveSheet()->setCellValue($cell, $census['address2']);
		}
            $cell="N".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['city'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['city']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['city']);
		}
            $cell="O".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['state'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['state']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['state']);
		}
            $cell="P".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['zip'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['zip']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['zip']);
		}
            $cell="Q".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['email'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['email']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['email']);
		}
            $cell="R".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['phone'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['phone']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['phone']);
		}
            $cell="S".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['hire_date'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['hire_date']);
		} else {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['hire_date']);
		}
            $cell="T".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if (isset($employee['work_status'])) {
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['work_status']);
		} else {
			//$spreadsheet->getActiveSheet()->setCellValue($cell, $census['work_status```']);
		}
            $cell="U".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['eff_dt']);
		$sql="select APA_CODE from nua_company_plan where plan_code = '" . $e['client_plan'] . "' and company_id = " . $employee['company_id'];
		$z=$X->sql($sql);
		if (sizeof($z) > 0) { 
                     $cell="U".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	             $spreadsheet->getActiveSheet()->setCellValue($cell, $z[0]['APA_CODE']);
		} 
                $cell="V".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                $spreadsheet->getActiveSheet()->setCellValue($cell, $e['coverage_level']);
                $cell="W".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
	        $spreadsheet->getActiveSheet()->setCellValue($cell, $e['client_plan']);

		if ($e['coverage_level']!='EE') {
                 $sql="select * from nua_employee_dependent where employee_id = " . $e['employee_id'] . " order by dependent_id";
                 $z=$X->sql($sql);
                 foreach($z as $ff) {
                    $row++;
                    $cell="B".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
                $cell="C".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['social_security_number']);
                $cell="D".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['last_name']);
                $cell="E".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['first_name']);
                $cell="F".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['middle_name']);
                $cell="G".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['relationship']);
                $cell="H".$row;
                $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['date_of_birth']);
                $cell="I".$row;
                   $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                   $spreadsheet->getActiveSheet()->setCellValue($cell, $ff['gender']);
	    }
	}
}
$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$filename="APA_Enrollment_File.xlsx";
if ($display=="F") {
$file="sheets/" . $filename;
$writer->save($file);
return $file;
} else {
    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();
}
 
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
function process_enrollment($data,$company_id) {
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
        //
        //EMPLOYEE
        //
        if (strtoupper($data['relationship'])=="EMPLOYEE") {
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



function process_enrollment_old($data,$company_id) {
	$outp=array();
	$outp['c']=1;
        $month=$this->getMonthId();

	print_r($data);
           $X=new DB();
	   if (strtoupper($data['relationship'])=="EMPLOYEE"||
		   //--
		   // Processing Employee
		   // 

			($data['employee_code']!=""&&$data['dependent_code']=="")||
			($data['social_security_number']!=""&&$data['dependent_social_security_number']=="")) {
                        $error='';
			if ($data['min_month_id']=="") { 
			        $data['min_month_id']=$month;
			        $data['max_month_id']="";
			}
			$data['relationship']="Employee";
			

			$monthlist=$this->makeMonthList($data['min_month_id'],$data['max_month_id']);
			$deletes=$this->makeDeleteList($data['min_month_id'],$data['max_month_id']);

                        foreach ($deletes as $m) {
                           $sql="delete from nua_monthly_member_census where employee_code = '" . $data['employee_code'] . "'";
                           $sql.=" and month_id = '" . $m . "' ";
                           $X->execute($sql);
                        }

                 	$outp['c']=3;
			foreach($monthlist as $month_id) {

				$p=array();
				$p['table_name']="nua_monthly_member_census";
				$p['action']="insert";
				$p['month_id']=$month_id;
				$p['employee_code']=$data['employee_code'];
				$p['dependent_code']=$data['dependent_code'];
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

                                if ($data['relationship']=="") {
                                    if ($data['dependent_social_security_number']=="") {
                                        $data['relationship']="EMPLOYEE";
				    } else {
                                        $data['relationship']="UNKNOWN";
				        $error.=", Dependent relationship not entered"; 
                                    }
                                }
                                $p['relationship']=$data['relationship'];
                                $p['dob']=substr($data['date_of_birth'],0,10);
				$p['ssn']=$data['social_security_number'];
                                if (strtoupper($data['gender'])!="M"&&strtoupper($data['gender'])!="F") {
				    $error.=", Gender not M or F"; 
                                }
                                $p['gender']=strtoupper($data['gender']);
                                if (substr(strtoupper($data['marital_status']),0,1)!="M"&&substr(strtoupper($data['marital_status']),0,1)!="F") {
				  //  $error.=", Marital Status not M or S"; 
                                }
                                $p['marital_status']=strtoupper(substr($data['marital_status'],0,1));
                                $p['address']=strtoupper($data['address']);
                                $p['address2']=strtoupper($data['suite']);
                                $p['city']=strtoupper($data['city']);
                                $p['state']=strtoupper($data['state']);
                                $p['zip']=strtoupper($data['zip']);
                                $p['email']=strtolower($data['email']);
                                $p['phone']=$data['contact_phone'];
                                $p['date_hired']=substr($data['date_of_hire'],0,10);

	$outp['c']=4;

                           //-- Find out if Employee is already in company.
                           //--
if ($p['employee_code']=="") {
         $sql="select * from nua_employee where company_id = " . $company_id . " and last_name = '" . str_replace("'","''",$p['last_name']) . "' AND ";
         $sql.=" first_name = '" . str_replace("'","''",$p['first_name']) . "' AND ";
         $sql.=" middle_name = '" . str_replace("'","''",$p['middle_name']) . "'";

                           $e=$X->sql($sql);
                           $employee_id=0;
                           $employee_code='n0';
                           if (sizeof($e)>0) {
                                $employee_id = $e[0]['id'];
                                $p['employee_id']=$employee_id;
                                $employee_code = $e[0]['employee_code'];
                                $p['employee_code']=$employee_code;
                           }
} else {

         $sql="select * from nua_employee where employee_code = '" . $p['employee_code'] . "'";
                           $e=$X->sql($sql);

                           $employee_id=0;
                           $employee_code='n0';
                           if (sizeof($e)>0) {
                                $employee_id = $e[0]['id'];
                                $p['employee_id']=$employee_id;
                                $employee_code = $p['employee_code'];
			   } else {
				   $sql="insert into nua_employee (id, employee_code, create_timestamp) values (";
				   $sql.=str_replace('n','',$p['employee_code']) . ",'" . $p['employee_code'] . "','2022-01-01')";
				   $X->execute($sql);
                                   $employee_id=str_replace('n','',$p['employee_code']);
				   $employee_code=$p['employee_code'];
				   $p['employee_id']=$employee_id;
		           }
}
	$outp['c']=5;
                           $emp=array();
                           $emp['table_name']="nua_employee";
                           $emp['action']="insert";
                           $emp['company_id']=$company_id;
                           $emp['last_name']=strtoupper($data['last_name']);
                           $emp['first_name']=strtoupper($data['first_name']);
                           $emp['middle_name']=strtoupper($data['middle_initial']);
                           $emp['social_security_number']=$data['social_security_number'];

                           $emp['date_of_birth']=substr($data['date_of_birth'],0,10);
                           $emp['gender']=$data['gender'];
                           $emp['marital_status']=strtoupper(substr($data['marital_status'],0,1));
                           $emp['address']=strtoupper($data['address']);
                           $emp['suite']=strtoupper($data['suite']);
                           $emp['city']=strtoupper($data['city']);
                           $emp['state']=strtoupper($data['state']);
                           $emp['zip']=$data['zip'];
                           $emp['email']=strtolower($data['email']);
                           $emp['phone']=$data['contact_phone'];
                           $emp['date_hired']=substr($data['date_of_hire'],0,10);
                           $emp['work_status']=$data['work_status'];
                           $emp['eff_dt']=substr($data['eff_dt'],0,10);
                           $emp['medical_eff_dt']=substr($data['eff_dt'],0,10);
                           $emp['medical_plan']=strtoupper($data['medical_plan']);
                           $emp['medical_plan_level']=strtoupper($data['medical_coverage_level']);
                           $emp['dental_plan']=strtoupper($data['dental_plan']);
                           $emp['dental_plan_level']=strtoupper($data['dental_coverage_level']);
                           $emp['vision_plan']=strtoupper($data['vision_plan']);
                           $emp['vision_plan_level']=strtoupper($data['vision_coverage_level']);
                           if ($data['medical_plan']==""&&$data['dental_plan']==""&&$data['vision_plan']=="") {
                               $error.=", No plan selected";
                           }
                           if ($data['medical_plan']!='') {
                               $sql="select * from nua_company_plan where plan_code = '" . strtoupper($data['medical_plan']) . "' and company_id = " . $company_id;
                               $r=$X->sql($sql);
                               if (sizeof($r)>0) {
                                  $ee_price=$r[0]['ee_price'];
                                  $ees_price=$r[0]['ees_price'];
                                  $eec_price=$r[0]['eec_price'];
                                  $fam_price=$r[0]['fam_price'];
                                  $med_apa_code=strtoupper($r[0]['APA_CODE']);
                               }  else {
                                  $emp['medical_plan']="";
                                  $emp['medical_plan_level']="";
                                  $med_apa_code="";
                                  $error.=", INVALID MED Plan";
                               }
                           }
	$outp['c']=6;
                           if ($data['dental_plan']!='') {
                               $sql="select * from nua_company_plan where plan_code = '" . strtoupper($data['dental_plan']) . "' and company_id = " . $company_id;
                               $r=$X->sql($sql);
                               if (sizeof($r)>0) {
                                  $dee_price=$r[0]['ee_price'];
                                  $dees_price=$r[0]['ees_price'];
                                  $deec_price=$r[0]['eec_price'];
                                  $dfam_price=$r[0]['fam_price'];
                                  $dental_apa_code="GUARDHIGH";
                               } else {
                                  $dee_price="0.00";
                                  $dees_price="0.00";
                                  $deec_price="0.00";
                                  $dfam_price="0.00";
                                  $dental_apa_code="GUARDHIGH";
                                  $error.=", INVALID DENTAL Plan";
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
                               } else {
                                  $vee_price="0.00";
                                  $vees_price="0.00";
                                  $veec_price="0.00";
                                  $vfam_price="0.00";
                                  $vision_apa_code="VSP";
                                  $error.=", INVALID VISION Plan";
                               }
                           }
	$outp['c']=7;

                           if ($employee_id!=0) {
                                $emp['id']=$employee_id;
                                $emp['employee_code']='n'.$employee_id;
                                $this->X->post($emp);
                           } else {
                                $employee_id = $this->X->post($emp);
                                $sql="update nua_employee set employee_code = 'n" . $employee_id . "' where id = " . $employee_id;
                                $X->execute($sql);
                            }

                            //
                            //-- Insert the employee into each months census from the effective date to the
                            //-- current billing month.
                            //--
                            $p['month_id']=$month_id;
                            $p['eff_dt']=substr($data['eff_dt'],0,10);
                            $p['employee_id']=$employee_id;
                            $p['employee_code']='n'.$employee_id;
                            $p['company_name']=$company_name;
			    $p['company_id']=$company_id;
                            $p['error_msg']=$error;
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
                        //         $sql.=" and client_plan = '" . strtoupper($p['client_plan']) . "' and dependent_code = ''";
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
				   $p['id']="";
	$outp['c']=8;
			    } else {
                                 $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                 $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                 $sql.=" and plan_type = '*MEDICAL*' and dependent_code = ''";
                                 $X->execute($sql);
			    }

                                if ($data['dental_plan']!="") {
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
                                 //  $sql.=" and client_plan = '" . strtoupper($p['client_plan']) . "' and dependent_code = ''";
                                   $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                   $dd=$X->sql($sql);
                                   if (sizeof($dd)==1) {
                                        $p['id']=$dd[0]['id'];
				   } else {
                                        $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                        $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                        $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                        $X->execute($sql);
			           }
                                   $this->X->post($p);
				   if (isset($p['id'])) unset($p['id']);
				   $p['id']="";
	$outp['c']=9;
			    } else {
                                 $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                 $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                 $sql.=" and plan_type = '*DENTAL*' and dependent_code = ''";
                                 $X->execute($sql);

			    }
                                if ($data['vision_plan']!="") {
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
                                   //$sql.=" and client_plan = '" . strtoupper($p['client_plan']) . "' and dependent_code = ''";
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
				   $p['id']="";
			    } else {
                                 $sql="delete from nua_monthly_member_census where employee_id = " . $employee_id;
                                 $sql.=" and month_id = '" . $p['month_id'] . "' ";
                                 $sql.=" and plan_type = '*VISION*' and dependent_code = ''";
                                 $X->execute($sql);

			    }
	$outp['c']=10;
                   }  // Months 
			} // Employee 
                		if (strtoupper($data['relationship'])!="EMPLOYEE"||
				($data['employee_code']!=""&&$data['dependent_code']!="")||
				($data['social_security_number']!=""&&$data['dependent_social_security_number']!="")) {
				if ($data['employee_code']!="") {
                                     $sql="select id, employee_code from nua_employee where employee_code = '" . $data['employee_code'] . "'";
                                } else {
                                     $sql="select id, employee_code from nua_employee where social_security_number = '" . $data['social_security_number'] . "'";							
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
                                $sql.=" and last_name = '" . strtoupper(str_replace("'","''",$data['last_name'])) . "' AND ";
                                $sql.=" first_name = '" . strtoupper(str_replace("'","''",$data['first_name'])) . "' AND ";
                                $sql.=" middle_name = '" . strtoupper(str_replace("'","''",$data['middle_initial'])) . "'";
                                $e=$X->sql($sql);

                                $dep['employee_id']=$employee_id;
                                $dep['employee_code']=$employee_code;
                                $dep['company_id']=$company_id;
                                $dep['first_name']=strtoupper($data['first_name']);
                                $dep['last_name']=strtoupper($data['last_name']);
                                $dep['middle_name']=strtoupper($data['middle_initial']);
                                $dep['gender']=strtoupper($data['gender']);
                                $dep['date_of_birth']=substr($data['date_of_birth'],0,10);
                                $dep['social_security_number']=$data['dependent_social_security_number'];
                                $dep['relationship']=strtoupper($data['relationship']);
                                if (strtoupper($dep['relationship'])=='WIFE') $dep['relation_type']="02";
                                if (strtoupper($dep['relationship'])=='HUSBAND') $dep['relation_type']="02";
                                if (strtoupper($dep['relationship'])=='SON') $dep['relation_type']="03";
                                if (strtoupper($dep['relationship'])=='DAUGHTER') $dep['relation_type']="03";
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
} // CLASS

