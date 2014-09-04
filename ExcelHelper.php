<?php

//import (and instantiate) the required 3rd party autoload class for all required classes
App::import('Vendor', 'PHPExcel', false, null, 'php_excel' . DS . 'PHPExcel.php');
PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );

class ExcelHelper extends AppHelper {

    // Storage of default settings
    protected $settings = array(
        'author'         => '',
        'lastModifiedBy' => '',
        'title'          => '',
        'subject'        => '',
        'description'    => '',
        'keywords'       => array(),
        'category'       => ''
    );
    
    private $defaultFilename = 'Untitled';    //.xslx will be appended automatically
    private $converter       = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');

    public function create($options = array()){
        $this->settings = array_merge($this->settings, $options);
        $sheet = new PHPExcel();
        $sheet->getProperties()->setCreator($this->settings['author'])
                               ->setLastModifiedBy($this->settings['lastModifiedBy'])
                               ->setTitle($this->settings['title'])
                               ->setSubject($this->settings['subject'])
                               ->setDescription($this->settings['description'])
                               ->setKeywords(implode(' ', $this->settings['keywords']))
                               ->setCategory($this->settings['category']);
        return $sheet;    //type PHPExcel object
    }

    public function addModel($obj, $data, $sheetname = '', $autowidth = false){
        if (empty($data)) return $obj;
        
        $temp = array();
        $columns = array();
        $temp = array_keys($data[0]);
        foreach ($temp as $col) {
            $columns[] = Inflector::humanize($col);    // Ex: asset_category_id => 'Asset Category Id'
        }
        unset($temp);

        $obj->createSheet();
        $obj->setActiveSheetIndex($obj->getSheetCount()-1);
        foreach ($columns as $key => $column) {
            // I don't expect more than 26 fields in a table, if there are, '1' needs to increment here and below
            $obj->getActiveSheet()->setCellValue($this->converter[$key].'1', $column);
        }
        $obj->getActiveSheet()->fromArray($data, NULL, 'A2');
        $obj->getActiveSheet()->getStyle('A1:'.$this->converter[count($data[0])-1].'1')->getFont()->setBold(true);
        $obj->getActiveSheet()->setAutoFilter($obj->getActiveSheet()->calculateWorksheetDimension());
        if ($sheetname) {
            $obj->getActiveSheet()->setTitle($sheetname);    
        }
        
        // Autosizing of columns must be done after all data has been entered UNLESS you only want to initially see the column name
        if ($autowidth) {
            foreach ($columns as $key => $column) {
                $obj->getActiveSheet()->getColumnDimension($this->converter[$key])->setAutoSize(true);
            }
        }
        
        return $obj;
    }
    
    public function render($obj, $filename = '') {
        // If the filename is empty...
        if (!$filename) {
            $filename = $this->defaultFilename;
        }
        
        // Remove the first sheet created during instantiation, after adding our data sheets
        $obj->removeSheetByIndex(0);

        // Make sure the first sheet in the workbook is the active sheet
        $obj->setActiveSheetIndex(0);
        
        $objWriter = PHPExcel_IOFactory::createWriter($obj, 'Excel2007');
        // Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'.xlsx"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $objWriter = PHPExcel_IOFactory::createWriter($obj, 'Excel2007');
        $objWriter->save('php://output');
    }
}
