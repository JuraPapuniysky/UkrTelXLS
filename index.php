<?php
error_reporting(E_ALL);
set_time_limit(0);

date_default_timezone_set('Europe/Kiev');
require_once './Formatter.php';



$inputXLSFileName = './files/report.xls';
$outCSVFileName = './files/report.csv';
$subdivisionFile = './config/subdivision.php';
$subdivisions = '';
if(file_exists($subdivisionFile) && is_readable($subdivisionFile)){
    $subdivisions = require_once $subdivisionFile;
}

$formatter = new Formatter($inputXLSFileName, $outCSVFileName, $subdivisions);
$formatter->save();



