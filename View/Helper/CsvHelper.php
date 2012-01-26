<?php
include(APP . DS . 'Plugin' . DS . 'PhpExcel' . DS . 'Vendor' . DS . 'PhpExcel' . DS . 'IOFactory.php');
include(APP . DS . 'Plugin' . DS . 'PhpExcel' . DS . 'Vendor' . DS . 'PhpExcel' . DS . 'PHPExcel.php');

class CsvHelper extends AppHelper { 
     
	var $objPHPExcel;
    var $writer; 
    var $sheet; 
    var $data; 
    var $blacklist = array(); 
     
    public function csvHelper() { 
		$this->objPHPExcel = new PHPExcel();
        $this->writer = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'CSV');
        $this->sheet = $this->objPHPExcel->getActiveSheet(); 
        $this->sheet->getDefaultStyle()->getFont()->setName('Verdana'); 
    } 
                  
    function generate(&$data, $title = 'Report') { 
         $this->data =& $data; 
         $this->_title($title); 
         $this->_headers(); 
         $this->_rows(); 
         $this->_output($title); 
         return true; 
    } 
     
    function _title($title) { 
        $this->sheet->setCellValue('A2', $title); 
        $this->sheet->getStyle('A2')->getFont()->setSize(14); 
        $this->sheet->getRowDimension('2')->setRowHeight(23); 
    } 

    function _headers() { 
        $i=0; 
        foreach ($this->data[0] as $field => $value) { 
            if (!in_array($field,$this->blacklist)) { 
                $columnName = Inflector::humanize($field); 
                $this->sheet->setCellValueByColumnAndRow($i++, 4, $columnName); 
            } 
        } 
		/**
        $this->sheet->getStyle('A4')->getFont()->setBold(true); 
        $this->sheet->getStyle('A4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID); 
        $this->sheet->getStyle('A4')->getFill()->getStartColor()->setRGB('969696'); 
        $this->sheet->duplicateStyle( $this->sheet->getStyle('A4'), 'B4:'.$this->sheet->getHighestColumn().'4'); 
        for ($j=1; $j<$i; $j++) { 
            $this->sheet->getColumnDimension(PHPExcel_Cell::stringFromColumnIndex($j))->setAutoSize(true); 
        } 
        ***/
    } 
         
    function _rows() { 
        $i=5; 
        foreach ($this->data as $row) { 
            $j=0; 
            foreach ($row as $field => $value) { 
                if(!in_array($field,$this->blacklist)) { 
                    $this->sheet->setCellValueByColumnAndRow($j++,$i, $value); 
                } 
            } 
            $i++; 
        } 
    } 
             
    function _output($title) { 
        header("Content-type: text/csv");  
        header('Content-Disposition: attachment;filename="'.$title.'.csv"'); 
        header('Cache-Control: max-age=0'); 
		$this->writer = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'CSV');
        //$this->writer->setTempDir(TMP); 
        $this->writer->save('php://output'); 
    } 
}

?>