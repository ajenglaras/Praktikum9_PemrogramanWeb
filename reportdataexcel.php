<?php
//mengoneksikan database dan excel
include('koneksi.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//untuk memberikan nama tiap kolom pada excel
$spreadsheet = new Spreadsheet(); 
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No'); 
$sheet->setCellValue('B1', 'Nama');
$sheet->setCellValue('C1', 'Kelas');
$sheet->setCellValue('D1', 'Alamat');
//perintah query untuk mengambil data dari database kemudian mencetak data di file excel
$query = mysqli_query($koneksi,"select * from tb_siswa");
$i = 2; 
$no = 1;
while($row = mysqli_fetch_array($query))
{
$sheet->setCellValue('A'.$i, $no++);
$sheet->setCellValue('B'.$i, $row['nama']);
$sheet->setCellvalue('C'.$i, $row['kelas']); 
$sheet->setCellValue('D'.$i, $row['alamat']);
$i++;
}
//menambah style excel
$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$i = $i - 1; 
$sheet->getStyle('A1:D'.$i)->applyFromArray($styleArray);
//menyimpan file excel
$writer = new Xlsx($spreadsheet); 
$writer->save('Report Data Siswa.xlsx'); 
?>