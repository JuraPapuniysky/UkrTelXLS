<?php
error_reporting(E_ALL);
set_time_limit(0);

date_default_timezone_set('Europe/Kiev');
require_once './Formatter.php';



$inputXLSFileName = './files/test.xls';
$outCSVFileName = './files/report.csv';


$formatter = new Formatter($inputXLSFileName, $outCSVFileName);
$formatter->save();



