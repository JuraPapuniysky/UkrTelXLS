<?php
/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . './Classes/');

/** PHPExcel_IOFactory */
include 'PHPExcel/IOFactory.php';

class Formatter{

    private $objPHPExcel;
    private $csvFp;
    private $sheetData;

    private $outCSVFileName;


    public function __construct($inputXLSFileName, $outCSVFileName)
    {
        $this->outCSVFileName = $outCSVFileName;
        echo 'Loading file ',pathinfo($inputXLSFileName,PATHINFO_BASENAME),' using IOFactory to identify the format';
        $this->objPHPExcel = PHPExcel_IOFactory::load($inputXLSFileName);
        $this->csvFp = fopen($outCSVFileName, 'w');
        $this->sheetData = $this->objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
    }

    public function save(){
        foreach ($this->sheetData as $item){

            if(is_numeric($item['B']) && (strlen($item['B']) == 1)){
                fputcsv($this->csvFp, [
                    '',
                    $this->formatDate($item['F']),
                    $this->formatTime($item['F']),
                    $item['C'],
                    str_replace(' ', '',$item['E']),
                    $item['D'],
                    $item['Q'],
                    $item['S'],
                    '',
                ], '|');
            }
        }

        fclose($this->csvFp);
        $f = file_get_contents($this->outCSVFileName);
        $f = iconv("UTF-8", "WINDOWS-1251", $f);
        file_put_contents($this->outCSVFileName, $f);
    }

    private function formatDate($xlsDateTime){
        return str_replace('/', '.', substr($xlsDateTime, 0, 5)).'.'.date('Y');
    }

    private function formatTime($xlsDateTime){
        return str_replace('.', ':', substr($xlsDateTime, 6));
    }

}