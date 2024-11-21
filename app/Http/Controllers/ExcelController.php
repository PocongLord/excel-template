<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class ExcelController extends Controller
{

    private function generateMaterialNumber($cellC, $cellE)
{
    // Pengecekan kata khusus pada kolom manufaktur
    if (stripos($cellC, 'Atlas Copco') !== false) {
        $prefixC = 'ACP'; // Gunakan "ACP" sebagai prefix
    } elseif (stripos($cellC, 'Multi Flow') !== false || stripos($cellC, 'Multiflow') !== false) {
        $prefixC = 'MLF'; // Gunakan "MLF" sebagai prefix
    } else {
        // Ambil 3 huruf pertama dari kolom C (Manufaktur)
        $prefixC = substr(strtoupper($cellC), 0, 3);
    }

    // Bersihkan simbol dari kolom E dan ubah menjadi kapital
    $cleanedE = preg_replace('/[^A-Za-z0-9]/', '', strtoupper($cellE));

    // Gabungkan LG2 + prefixC + kolom E
    $materialNumber = 'LG2' . $prefixC . $cleanedE;

    // Jika panjang kombinasi lebih dari 18 karakter
    if (strlen($materialNumber) > 18) {
        // Kurangi karakter dari awal kolom E hingga panjangnya menjadi 18
        $excessLength = strlen($materialNumber) - 18;
        $materialNumber = 'LG2' . $prefixC . substr($cleanedE, $excessLength);
    }
    // Jika panjang kombinasi kurang dari 18 karakter, tambahkan '0' setelah LG2 dan prefixC
    elseif (strlen($materialNumber) < 18) {
        // Hitung jumlah karakter yang perlu ditambahkan agar panjangnya menjadi 18
        $remainingLength = 18 - strlen($materialNumber);

        // Tambahkan '0' di antara "LG2" + prefixC dan cleanedE
        $materialNumber = 'LG2' . $prefixC . str_pad($cleanedE, strlen($cleanedE) + $remainingLength, '0', STR_PAD_LEFT);
    }

    // Pastikan panjang tetap 18 karakter dan kembalikan dalam huruf kapital
    return strtoupper(substr($materialNumber, 0, 18));
}



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

    foreach ($sheet->getRowIterator(2) as $row) {
        $rowIndex = $row->getRowIndex();
        
        // Ambil nilai dari kolom yang ada di file asli dan ubah menjadi kapital
        $cellC = strtoupper($sheet->getCell('C'.$rowIndex)->getValue()); // Manufaktur
        $cellE = strtoupper($sheet->getCell('E'.$rowIndex)->getValue()); // Old Material Number
        $cellF = strtoupper($sheet->getCell('F'.$rowIndex)->getValue()); // Material Description
        $cellG = strtoupper($sheet->getCell('G'.$rowIndex)->getValue()); // Material Group
        $cellH = strtoupper($sheet->getCell('H'.$rowIndex)->getValue()); // External Material Group
        $cellI = strtoupper($sheet->getCell('I'.$rowIndex)->getValue()); // Material Type
        $cellJ = strtoupper($sheet->getCell('J'.$rowIndex)->getValue()); // UOM

        // Validasi jika kolom C (Manufaktur) atau kolom E (Old Material Number) kosong
        if (empty($cellC) || empty($cellE)) {
            // Jika kosong, lewati baris ini dan lanjutkan ke baris berikutnya
            continue;
        }

        // Panggil fungsi generateMaterialNumber
        $materialNumber = $this->generateMaterialNumber($cellC, $cellE);

        // Set nilai di kolom D (Material Number) menjadi kapital
        $sheet->setCellValue('D'.$rowIndex, strtoupper($materialNumber));

        // Set nilai di kolom E (Length of Material Number)
        $sheet->setCellValue('E'.$rowIndex, strlen($materialNumber));

        // Set nilai di kolom F (Old Material Number)
        $sheet->setCellValue('F'.$rowIndex, $cellE);

        // Set nilai di kolom G (Material Description)
        $sheet->setCellValue('G'.$rowIndex, $cellF);

        // Set nilai di kolom H (Material Group)
        $sheet->setCellValue('H'.$rowIndex, $cellG);

        // Set nilai di kolom I (External Material Group)
        $sheet->setCellValue('I'.$rowIndex, $cellH);

        // Set nilai di kolom J (Material Type)
        $sheet->setCellValue('J'.$rowIndex, $cellI);

        // Set nilai di kolom K (UOM)
        $sheet->setCellValue('K'.$rowIndex, $cellJ);
    }

    // Simpan file sebagai Excel baru
    $writer = new Xlsx($spreadsheet);
    $newFileName = 'template-' . time() . '.xlsx';
    $writer->save(storage_path('app/' . $newFileName));

    return response()->download(storage_path('app/' . $newFileName));
}










}

