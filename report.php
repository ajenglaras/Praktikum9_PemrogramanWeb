<?php 
//mengoneksikan database dan excel
require 'vendor/autoload.php'; 
use PhpOffice\PhpSpreadsheet\Spreadsheet; 
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 
//untuk memberikan nama kolom dan isi data hello world pada excel
$spreadsheet = new Spreadsheet(); 
$sheet = $spreadsheet->getActiveSheet(); 
$sheet->setCellValue('A1', 'Hello World !'); 
//menyimpan file excel
$writer = new Xlsx($spreadsheet); 
$writer->save('hello world.xlsx');
?>