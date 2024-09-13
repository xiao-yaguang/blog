<?php

namespace App\Components;

use PhpOffice\PhpSpreadsheet\Reader\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class Excel
{

    /**
     * @param $file
     * @param string $sheetName
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
    public static function readExcel($file, $sheetName = '') {
        $reader = IOFactory::createReader("Xlsx");
        $spreadsheet = $reader->load($file);
        if($sheetName){
            $worksheet = $spreadsheet->getSheetByName($sheetName);
        }else{
            $worksheet = $spreadsheet->getActiveSheet();
        }
        $highestRow = $worksheet->getHighestRow(); // 总行数
        $highestColumn = $worksheet->getHighestColumn(); // 总列数
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);
        $data = [];
        for ($i = 1; $i <= $highestRow; $i++) {
            for ($j = 1; $j <= $highestColumnIndex; $j++) {
                //严格模式下 格式复杂的excel 每一列处理方式可能得额外处理
                $value = $worksheet->getCellByColumnAndRow($j, $i)->getCalculatedValue();

                if (is_object($value)) {
                    $value = $value->__tostring();
                }
                $data[$i - 1][$j - 1] = $value;
            }
        }
        return array_values($data);
    }

    public static function exportExcel($header = [], $data = [], $filename = "") {
        $spreadsheet = new Spreadsheet;
        $sheet = $spreadsheet->getActiveSheet();
        foreach ($header as $key => $value) {
            $sheet->setCellValueByColumnAndRow($key + 1, 1, $value);
        }
        $i = 2;
        foreach ($data as $rows) {
            $j = 1;
            foreach ($rows as $value) {
                $sheet->setCellValueByColumnAndRow($j, $i, $value);
                $j++;
            }
            $i++;
        }
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        self::header($filename);
        $writer->save('php://output');
    }

    /**
     * 纵向导出
     * @param array $header
     * @param array $data
     * @param string $filename
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function exportExcelVertical($header = [], $data = [], $filename = "") {
        $spreadsheet = new Spreadsheet;
        $sheet = $spreadsheet->getActiveSheet();
        $spreadsheet->getDefaultStyle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER); //设置所有的列水平居中
        $sheet->getColumnDimension('A')->setWidth(24); //宽度
        foreach ($header as $key => $value) {
            $sheet->setCellValueByColumnAndRow(1, $key + 1, $value)
                ->getStyleByColumnAndRow(1, $key + 1)
                ->getFont()->setBold(true); //->setSize(20);
//            $sheet->getStyleByColumnAndRow(1, $key + 1)
//                ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF808080');
        }

        $i = 2;
        foreach ($data as $rows) {
            $j = 1;
            $sheet->getColumnDimensionByColumn($i)->setWidth(30);
            foreach ($rows as $value) {
                $sheet->setCellValueByColumnAndRow($i, $j, $value);
                $sheet->getStyleByColumnAndRow($i, $j)->getAlignment()->setWrapText(true);
                $j++;
            }
            $i++;
        }
        //设置边框
//        $sheet->getStyle('A1:C34')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        self::header($filename);
        $writer->save('php://output');
    }

    public static function header($filename = ''){
        header('Access-Control-Allow-Origin: *');
        $agent = strtolower($_SERVER['HTTP_USER_AGENT'] ?? '');
        if (strripos($agent, 'msie') || strripos($agent, 'triden')) {
            header('Access-Control-Allow-Headers: Origin, X-Requested-With, content-type, Accept, authorization');
        } else {
            header('Access-Control-Allow-Headers: *');
        }
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename=' . $filename . date('Y-m-d') . '.xlsx');
        header('Cache-Control: max-age=0');
        header('Content-type: charset=UTF-8');
    }

    public static function exportExcelToFile($header = [], $data = [], $filename = ""){
        $spreadsheet = new Spreadsheet;
        $sheet = $spreadsheet->getActiveSheet();
        foreach ($header as $key => $value) {
            $sheet->setCellValueByColumnAndRow($key + 1, 1, $value);
        }
        $i = 2;
        foreach ($data as $rows) {
            $j = 1;
            foreach ($rows as $value) {
                $sheet->setCellValueByColumnAndRow($j, $i, $value);
                $j++;
            }
            $i++;
        }
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($filename);
    }

    /**
     * @param $list [[header => [], data => [], name => '']]
     * @param string $fileName
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public static function exportMultiExcel($list, $fileName = ''){
        $spreadsheet = new Spreadsheet();
        foreach ($list as $k => $v){
            $sheet = $k == 0 ? $spreadsheet->getActiveSheet() : $spreadsheet->createSheet();
            $name = $v['name'] ?? $k;
            if(mb_strlen($name) > $sheet::SHEET_TITLE_MAXIMUM_LENGTH){
                $name = mb_substr($name, 0, $sheet::SHEET_TITLE_MAXIMUM_LENGTH - 4). '...';
            }

            $checkTitle = $sheet->getInvalidCharacters();
            $name = str_replace($checkTitle, ' ', $name);
            $sheet->setTitle($name);
            //设置单元格内容

            $row = 1; //从第二行开始
            if($v['header']){
                foreach ($v['header'] as $key => $value) {
                    $sheet->setCellValueByColumnAndRow($key + 1, $row, $value);
                }
                $row++;
            }
            foreach ($v['data'] ?? [] as $item) {
                $column = 1;
                foreach ($item as $value) {
                    $sheet->setCellValueByColumnAndRow($column, $row, $value);
                    $column++;
                }
                $row++;
            }
        }
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        self::header($fileName);
        $writer->save('php://output');
    }

    public static function exportCSV($header, $list, $fileName = 'excel')
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        //设置单元格内容
        foreach ($header as $key => $value) {
            $sheet->setCellValueByColumnAndRow($key + 1, 1, $value);
        }
        $row = 2; //从第二行开始
        foreach ($list as $item) {
            $column = 1;
            foreach ($item as $value) {
                $sheet->setCellValueByColumnAndRow($column, $row, $value);
                $column++;
            }
            $row++;
        }
        $writer = IOFactory::createWriter($spreadsheet, 'Csv');

        self::header($fileName);
        //要解决PHP生成CSV文件的乱码问题，只需要在文件的开始输出BOM头，告诉windows CSV文件的编码方式，从而让Excel打开CSV时采用正确的编码。
        echo "\xEF\xBB\xBF";// UTF-8 BOM
        $writer->save('php://output');
        exit();
    }
}
