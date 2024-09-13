<?php

namespace App\Services\Post;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use App\Models\Post\Post;
use App\Components\Excel;

class PostServices
{
    //导出单sheet Excel文件
    public function postExport()
    {
        $postData = Post::query()->get()->toArray();
        $newExcel = new Spreadsheet(); //创建一个新的Excel文档
        $sheet = $newExcel->getActiveSheet(); //获取当前的操作对象
        //$sheet = $newExcel->createSheet(); //创建一个新的工作区间sheet(第二个sheet)
        $header = ["ID", "名称", "标题", "内容"];
        $row = 1; //初始化行
        $j = 1; //初始化列
        foreach ($header as $key => $value) {
            $sheet->setCellValueByColumnAndRow($j, $row, $value);
            $j++;
        }
        $row = 2; //初始化行
        //遍历数据
        foreach ($postData as $key => $value) {
            $j = 1; //初始化列
            foreach ($value as $k => $v) {
                $sheet->setCellValueByColumnAndRow($j, $row, $v);
                $j++;
            }
            $row++;
        }
        // 重命名工作表
        $sheet->setTitle('My Excel Sheet');
        // 创建Excel写入器
        $writer = IOFactory::createWriter($newExcel, 'Xlsx');
        Excel::header('export_file');
        $writer->save('php://output');
    }

    //导出多sheet Excel文件
    public function postExportManySheet()
    {
        $postData = Post::query()->get()->toArray();
        $spreadsheet = new Spreadsheet(); //创建一个新的Excel文档
        $spreadsheet->setActiveSheetIndex(0);
        $sheet = $spreadsheet->getActiveSheet(); //获取当前的操作对象
        // 重命名工作表
        $sheet->setTitle('Sheet1');
        $header = ["ID", "名称", "标题", "内容"];
        $row = 1; //初始化行
        $j = 1; //初始化列
        foreach ($header as $key => $value) {
            $sheet->setCellValueByColumnAndRow($j, $row, $value);
            $j++;
        }
        $row = 2; //初始化行
        //遍历数据
        foreach ($postData as $key => $value) {
            $j = 1; //初始化列
            foreach ($value as $k => $v) {
                $sheet->setCellValueByColumnAndRow($j, $row, $v);
                $j++;
            }
            $row++;
        }


        // 建立第二个sheet工作表
        $spreadsheet->createSheet(); //创建一个新的工作区间sheet(第二个sheet)
        $sheet2 = $spreadsheet->setActiveSheetIndex(1);
        $sheet2->setTitle('Sheet2');
        $header = ["ID", "名称", "标题", "内容"];
        $row = 1; //初始化行
        $j = 1; //初始化列
        foreach ($header as $key => $value) {
            $sheet2->setCellValueByColumnAndRow($j, $row, $value);
            $j++;
        }
        $row = 2; //初始化行
        //遍历数据
        foreach ($postData as $key => $value) {
            $j = 1; //初始化列
            foreach ($value as $k => $v) {
                $sheet2->setCellValueByColumnAndRow($j, $row, $v);
                $j++;
            }
            $row++;
        }

        // 创建Excel写入器
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        Excel::header('export_file');
        $writer->save('php://output');
    }


    //导出单sheet Excel文件 (带格式)
    public function postExportTable()
    {
        $styleArray = [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ],
        ];

//        //设置格式
//        $styleArray1 = [
//            'borders' => [
//                'allBorders' => [
//                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN //细边框
//                ]
//            ]
//        ];


        $postData = Post::query()->get()->toArray();
        $newExcel = new Spreadsheet(); //创建一个新的Excel文档
        $sheet = $newExcel->getActiveSheet(); //获取当前的操作对象
        //$sheet = $newExcel->createSheet(); //创建一个新的工作区间sheet(第二个sheet)
        $header = ["ID", "名称", "标题", "内容"];
        $row = 1; //初始化行
        $j = 1; //初始化列
        $sheet->getColumnDimension('A')->setWidth(30); //设置某一列的宽度
        foreach ($header as $key => $value) {
            $sheet->setCellValueByColumnAndRow($j, $row, $value);
            $j++;
        }

        //设置灰色填充并居中
        $sheet->getStyle("A2:A" . (count($postData) + 1))->getFont()->getColor()->setARGB('FF0000'); //设置字体颜色（查找html颜色编码，且不用带#号）
        $sheet->getStyle("A2:A" . (count($postData) + 1))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('F0F0F0'); //设置填充颜色
        $sheet->getStyle("A2:A" . (count($postData) + 1))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER); //设置居中


        $row = 2; //初始化行
        //遍历数据
        foreach ($postData as $key => $value) {
            $j = 1; //初始化列
            foreach ($value as $k => $v) {
                //除去第一列设置其它列大小和样式(具体可根据需求自行调整)
                if ($j > 1) {
                    $line = $sheet->getColumnDimensionByColumn($j)->getColumnIndex(); //获取列字母
                    $sheet->getStyle($line . $row . ':' . $line . $row)->getFont()->getColor()->setARGB('00FF00'); //设置字体颜色（查找html颜色编码，且不用带#号） //黄色字体
                    $sheet->getStyle($line . $row . ':' . $line . $row)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFFF00');  //绿色填充
                    $sheet->getStyle($line . $row . ':' . $line . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT); //设置居左
                    $sheet->setCellValueByColumnAndRow($j, $row, $v);
                } else {
                    $sheet->setCellValueByColumnAndRow($j, $row, $v);
                }
                $j++;
            }

            $row++;
        }
        // 重命名工作表
        $sheet->setTitle('My Excel Sheet');
        // 创建Excel写入器
        $writer = IOFactory::createWriter($newExcel, 'Xlsx');
        Excel::header('export_file');
        $writer->save('php://output');
    }

    public function createOrder()
    {
        dd(1);
        $wheres = array();
        $wheres['goods_id'] = 1;
        $number = M('store')->where($wheres)->getField('number');
        $redis = new Redis();
        for ($i = 0; $i < $number; $i++) {
            $redis->lpush('goods_number', 1);
        }
        echo $redis->llen('goods_number');
    }

    //生成唯一订单号
    public function build_order_no()
    {
        return date('ymd') . substr(implode(NULL, array_map('ord', str_split(substr(uniqid(), 7, 13), 1))), 0, 8);
    }

    //记录日志
    public function insertLog($event, $type = 0)
    {
        $data['event'] = $event;
        $data['type'] = $type;
        $res = M('log')->add($data);
    }


    //模拟下单操作
    //下单前判断redis队列库存量
    public function order()
    {
        $sku_id = 11;  //传入已知的sku_id;

        $wheres = array();
        $wheres['sku_id'] = $sku_id;
        $good_info = M('store')->where($wheres)->find();

        $user_id = rand(1, 200);
        $goods_id = $good_info['goods_id'];
        $price = $good_info['price'];
        $number = 1;//抢购时每次买一件商品


        $redis = new Redis();
        $count = $redis->rpop('goods_number');  //下单时做rpop 从goods_number中取出1
        if ($count == 0) {
            $this->insertLog('error:no goods_number redis');
            return;
        }


        if (($good_info['number'] - $number) <= 0) {
            $this->insertLog('商品售罄');  //如果库存为0写入日志 并停止下单操作
            return;
        }


        //生成订单
        $order_sn = $this->build_order_no();

        $data = array();
        $data['order_sn'] = $order_sn;
        $data['user_id'] = $user_id;
        $data['goods_id'] = $goods_id;
        $data['sku_id'] = $sku_id;
        $data['number'] = $number;
        $data['price'] = $price;

        $order_rs = M('order')->add($data);

        //库存减少
        $wheres['sku_id'] = $sku_id;
        $store_rs = M('store')->where($wheres)->setDec('number', $number);
        if ($store_rs) {
            $this->insertLog('库存减少成功');
        } else {
            $this->insertLog('库存减少失败');
        }
    }


}
