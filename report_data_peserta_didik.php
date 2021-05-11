<?php
//mengoneksikan database dan excel
include('koneksi2.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//untuk memberikan nama tiap kolom pada excel
$spreadsheet = new Spreadsheet(); 
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No'); 
$sheet->setCellValue('B1', 'Jenis pendaftaran');
$sheet->setCellValue('C1', 'Tanggal mmasuk sekolah');
$sheet->setCellValue('D1', 'NIS');
$sheet->setCellValue('E1', 'Nomor peserta ujian');
$sheet->setCellValue('F1', 'Pernah paud');
$sheet->setCellValue('G1', 'Pernah tk');
$sheet->setCellValue('H1', 'No seri SKHUN');
$sheet->setCellValue('I1', 'No seri ijazah');
$sheet->setCellValue('J1', 'Hobi');
$sheet->setCellValue('K1', 'Cita-cita');
$sheet->setCellValue('L1', 'Nama Lengkap');
$sheet->setCellValue('M1', 'Jenis Kelamin');
$sheet->setCellValue('N1', 'NISN');
$sheet->setCellValue('O1', 'NIK');
$sheet->setCellValue('P1', 'Tempat lahir');
$sheet->setCellValue('Q1', 'Tanggal lahir');
$sheet->setCellValue('R1', 'Agama');
$sheet->setCellValue('S1', 'Berkebutuhan Khusus');
$sheet->setCellValue('T1', 'Alamat_jalan');
$sheet->setCellValue('U1', 'RT');
$sheet->setCellValue('V1', 'RW');
$sheet->setCellValue('W1', 'Dusun');
$sheet->setCellValue('X1', 'Kelurahan atau desa');
$sheet->setCellValue('Y1', 'Kecamatan');
$sheet->setCellValue('Z1', 'Kode pos');
$sheet->setCellValue('AA1', 'Tempat tinggal');
$sheet->setCellValue('AB1', 'Moda transportasi');
$sheet->setCellValue('AC1', 'Nomor HP');
$sheet->setCellValue('AD1', 'Nomor telepon');
$sheet->setCellValue('AE1', 'Email pribadi');
$sheet->setCellValue('AF1', 'Penerima KPS KKS PKH KIP');
$sheet->setCellValue('AG1', 'No KPS KKS PKH KIP');
$sheet->setCellValue('AH1', 'Kewarganegaraan');
//perintah query untuk mengambil data dari database kemudian mencetak data di file excel
$query = mysqli_query($koneksi,"select * from pendaftaran");
$i = 2; 
$no = 1;
while($row = mysqli_fetch_array($query))
{
$sheet->setCellValue('A'.$i, $no++);
$sheet->setCellValue('B'.$i, $row['jenis_pendaftaran']);
$sheet->setCellvalue('C'.$i, $row['tanggal_masuk_sekolah']); 
$sheet->setCellValue('D'.$i, $row['nis']);
$sheet->setCellValue('E'.$i, $row['nomor_peserta_ujian']);
$sheet->setCellValue('F'.$i, $row['pernah_paud']);
$sheet->setCellValue('G'.$i, $row['pernah_tk']);
$sheet->setCellValue('H'.$i, $row['no_seri_skhun']);
$sheet->setCellValue('I'.$i, $row['no_seri_ijazah']);
$sheet->setCellValue('J'.$i, $row['hobi']);
$sheet->setCellValue('K'.$i, $row['cita_cita']);
$sheet->setCellValue('L'.$i, $row['nama_lengkap']);
$sheet->setCellValue('M'.$i, $row['jenis_kelamin']);
$sheet->setCellValue('N'.$i, $row['nisn']);
$sheet->setCellValue('O'.$i, $row['nik']);
$sheet->setCellValue('P'.$i, $row['tempat_lahir']);
$sheet->setCellValue('Q'.$i, $row['tanggal_lahir']);
$sheet->setCellValue('R'.$i, $row['agama']);
$sheet->setCellValue('S'.$i, $row['berkebutuhan_khusus']);
$sheet->setCellValue('T'.$i, $row['alamat_jalan']);
$sheet->setCellValue('U'.$i, $row['rt']);
$sheet->setCellValue('V'.$i, $row['rw']);
$sheet->setCellValue('W'.$i, $row['dusun']);
$sheet->setCellValue('X'.$i, $row['kelurahan_desa']);
$sheet->setCellValue('Y'.$i, $row['kecamatan']);
$sheet->setCellValue('Z'.$i, $row['kode_pos']);
$sheet->setCellValue('AA'.$i, $row['tempat_tinggal']);
$sheet->setCellValue('AB'.$i, $row['moda_transportasi']);
$sheet->setCellValue('AC'.$i, $row['nomor_hp']);
$sheet->setCellValue('AD'.$i, $row['nomor_telepon']);
$sheet->setCellValue('AE'.$i, $row['email_pribadi']);
$sheet->setCellValue('AF'.$i, $row['penerima_kps_kks_pkh_kip']);
$sheet->setCellValue('AG'.$i, $row['no_kps_kks_pkh_kip']);
$sheet->setCellValue('AH'.$i, $row['kewarganegaraan']);
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
$sheet->getStyle('A1:AH'.$i)->applyFromArray($styleArray);
//menyimpan file excel
$writer = new Xlsx($spreadsheet); 
$writer->save('Report Data Peserta Didik.xlsx'); 
?>