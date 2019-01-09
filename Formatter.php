<?php
/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . './Classes/');

/** PHPExcel_IOFactory */
include 'PHPExcel/IOFactory.php';

class Formatter{

    private $objPHPExcel;
    private $csvFp;
    private $sheetData;
    private $subdivisions;

    private $outCSVFileName;


    public function __construct($inputXLSFileName, $outCSVFileName, $subdivisions)
    {
        $this->outCSVFileName = $outCSVFileName;
        $this->subdivisions = $subdivisions;
        echo 'Loading file ',pathinfo($inputXLSFileName,PATHINFO_BASENAME),' using IOFactory to identify the format';
        $this->objPHPExcel = PHPExcel_IOFactory::load($inputXLSFileName);
        $this->csvFp = fopen($outCSVFileName, 'w');
        $this->sheetData = $this->objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
    }

    public function save(){
        $currentDivision = '';
        foreach ($this->sheetData as $item){
            if (is_numeric($item['B']) && (strlen($item['B']) != 1)){
                $currentDivision = $this->subdivisions[$item['B']];
            }
            if(is_numeric($item['B']) && (strlen($item['B']) == 1)){
                fputcsv($this->csvFp, [
                    '',
                    $this->formatDate($item['F']),
                    $this->formatTime($item['F']),
                    "($currentDivision)".$item['C'],
                    str_replace(' ', '',$item['E']),
                    $item['D'],
                    $item['Q'],
                    (float)$item['S'],
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
    	$currentYear = date('Y');
    	$currenntMonth = date('m');
    	if ($currenntMonth == '01'){
    		$currentYear = date('Y') - 1;
    	}
        return str_replace('/', '.', substr($xlsDateTime, 0, 5)) . '.' . $currentYear;
    }

    private function formatTime($xlsDateTime){
        return str_replace('.', ':', substr($xlsDateTime, 6));
    }

}
