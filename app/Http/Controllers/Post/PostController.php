<?php

namespace App\Http\Controllers\Post;

use App\Components\Excel;
use App\Components\RabbitMQ;
use App\Http\Controllers\Controller;
use App\Models\Kol\Platform;
use App\Models\Post\Post;
use App\Services\Post\PostServices;

use Illuminate\Support\Facades\Validator;
use Illuminate\Validation\Rule;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;


class PostController extends Controller
{

    /**
     * Create a new controller instance.
     *
     * @return void
     */
    public function __construct()
    {
        //
    }


    //导出单sheet Excel文件
    public function postExport()
    {
        for ($i = 0; $i <= 1000000; $i++){
            Post::insert([
                'name' => '张三'.$i,
                'title' => '标题'. $i,
                'content' => '内容'. $i,
            ]);
            echo $i;
            echo PHP_EOL;
        }
        return 1;


//        $services = new PostServices();
//        $services->postExport();
    }

    //导出多sheet Excel文件
    public function postExportMany()
    {
        $services = new PostServices();
        $services->postExportManySheet();
    }

    //导出多sheet Excel文件
    public function postExportTable()
    {
        $services = new PostServices();
        $services->postExportTable();
    }


    public function pushMq()
    {
        $params = \request(['name']);
        $validator = Validator::make($params, [
            'name' => 'string',
        ]);
        if ($validator->fails()) {
            return $this->apiError($validator->errors()->first());
        }
        RabbitMQ::push('xiaohongshu', json_encode($params), 'xiaohongshu_exchange');
        //RabbitMQ::push(self::MESSAGR_TYPE[$param['type']]['queue_name'], json_encode($pushData), self::MESSAGR_TYPE[$param['type']]['exchenge']);
        //RabbitMQ::push(self::MESSAGR_TYPE[$param['type']]['queue_name'], json_encode($pushData), self::MESSAGR_TYPE[$param['type']]['exchenge']);
        return 1;
    }

    public function popMq()
    {
        try {
            $callback = function ($messageObj) {
                $data = json_decode($messageObj->body, true);
                if (is_array($data) && !empty($data)) {
                    dd($data);
                }
                $messageObj->getChannel()->basic_ack($messageObj->getDeliveryTag());
            };

            RabbitMQ::consumer('xiaohongshu', $callback);
        } catch (\Exception $e) {

        }
    }





    //模拟订单
    public function createOrder()
    {
        $services = new PostServices();
        $services->createOrder();

    }



}
