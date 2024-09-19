<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class ExcelController extends Controller
{
    public function uploadForm()
    {
        return view('upload');
    }

    public function processUpload(Request $request)
{
    $request->validate([
        'file' => 'required|mimes:xlsx'
    ]);

    $file = $request->file('file');
    $spreadsheet = IOFactory::load($file->getRealPath());
    $sheet = $spreadsheet->getActiveSheet();

    // Mengubah kolom sesuai dengan logika yang diberikan
    foreach ($sheet->getRowIterator(2) as $row) { // mulai dari baris kedua
        $cellC = strtoupper($sheet->getCell('C'.$row->getRowIndex())->getValue()); // Manufaktur
        $cellE = strtoupper($sheet->getCell('E'.$row->getRowIndex())->getValue()); // Old Material Number

        // Ambil 3 huruf pertama dari kolom C (Manufaktur)
        $prefixC = substr($cellC, 0, 3);

        // Bersihkan simbol dari kolom E dan ubah menjadi kapital
        $cleanedE = preg_replace('/[^A-Za-z0-9]/', '', $cellE);

        // Gabungkan LG + 3 huruf dari kolom C + kolom E (yang sudah dibersihkan dan diubah kapital)
        $materialNumber = 'LG' . $prefixC . $cleanedE;

        // Jika panjang kombinasi kurang dari 18 karakter, tambahkan '0' setelah 3 huruf dari kolom C
        if (strlen($materialNumber) < 18) {
            $materialNumber = substr($materialNumber, 0, 5) . str_pad(substr($materialNumber, 5), 13, '0', STR_PAD_LEFT);
        }

        // Ubah material number menjadi kapital dan set nilai di kolom D
        $sheet->setCellValue('D'.$row->getRowIndex(), strtoupper($materialNumber));

        // Set nilai di kolom E (Length of Material Number)
        $sheet->setCellValue('E'.$row->getRowIndex(), strlen($materialNumber));

        // Kolom F, G, H, I, J tetap dari data awal, tapi diubah jadi kapital
        $sheet->setCellValue('F'.$row->getRowIndex(), strtoupper($sheet->getCell('F'.$row->getRowIndex())->getValue()));
        $sheet->setCellValue('G'.$row->getRowIndex(), strtoupper($sheet->getCell('G'.$row->getRowIndex())->getValue()));
        $sheet->setCellValue('H'.$row->getRowIndex(), strtoupper($sheet->getCell('H'.$row->getRowIndex())->getValue()));
        $sheet->setCellValue('I'.$row->getRowIndex(), strtoupper($sheet->getCell('I'.$row->getRowIndex())->getValue()));
        $sheet->setCellValue('J'.$row->getRowIndex(), strtoupper($sheet->getCell('J'.$row->getRowIndex())->getValue()));

        // Kolom K (Additional Data - Length of Material Number or Fixed Value)
        $sheet->setCellValue('K'.$row->getRowIndex(), '0200');
    }

    // Simpan file sebagai Excel baru
    $writer = new Xlsx($spreadsheet);
    $newFileName = 'template-' . time() . '.xlsx';
    $writer->save(storage_path('app/' . $newFileName));

    return response()->download(storage_path('app/' . $newFileName));
}




}

