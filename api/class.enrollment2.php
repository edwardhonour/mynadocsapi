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
}
