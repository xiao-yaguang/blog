<?php


namespace App\Components;


use App\Models\MqErrorRecord;
use Closure;
use PhpAmqpLib\Connection\AMQPStreamConnection;
use PhpAmqpLib\Message\AMQPMessage;

class RabbitMQ
{

    //防止使用new直接创建对象
    private function __construct(){}

    //防止使用clone克隆对象
    private function __clone(){}

    private static function getConnect(){

//        return new AMQPStreamConnection('rabbitmq-serverless-cn-5ce3uypol04.cn-shenzhen.amqp-16.net.mq.amqp.aliyuncs.com', '5672', 'MjpyYWJiaXRtcS1zZXJ2ZXJsZXNzLWNuLTVjZTN1eXBvbDA0OkxUQUk1dERDVmU2NmZSNzd4aUZWWjdzRg==', 'NjlBMDQ4RkY4N0MwMkQyM0Y5OTNGREI0MDc4NUM4OTlBQ0I0RkZEMToxNzI0MDY2MDAwNjcy', 'rabbitmq-serverless-cn-5ce3uypol04',
//            false, 'AMQPLAIN', null, 'en_US', 10, 9);

        return new AMQPStreamConnection('rabbitmq-serverless-cn-5ce3uypol04.cn-shenzhen.amqp-16.net.mq.amqp.aliyuncs.com', '5672', 'MjpyYWJiaXRtcS1zZXJ2ZXJsZXNzLWNuLTVjZTN1eXBvbDA0OkxUQUk1dERDVmU2NmZSNzd4aUZWWjdzRg==', 'NjlBMDQ4RkY4N0MwMkQyM0Y5OTNGREI0MDc4NUM4OTlBQ0I0RkZEMToxNzI0MDY2MDAwNjcy', 'Host_elc',
            false, 'AMQPLAIN', null, 'en_US', 10, 9);


    }

    public static function push(string $queue, string $messageBody, $exchange = 'router'){
        $connection = self::getConnect();
        $channel = $connection->channel();
        $channel->queue_declare($queue, false, true, false, false);
        $channel->exchange_declare($exchange, 'direct', false, true, false);
        $channel->queue_bind($queue, $exchange); // 队列和交换器绑定
        $message = new AMQPMessage($messageBody, array('content_type' => 'text/plain', 'delivery_mode' => AMQPMessage::DELIVERY_MODE_PERSISTENT));
        $a = $channel->basic_publish($message, $exchange); // 推送消息
        $channel->close();
        $connection->close();
    }

    /**
     * @param string $queue
     * @param Closure $callback
     * @return bool
     */
    public static function pop($queue, $callback){
        $connection = self::getConnect();
        $channel = $connection->channel();
        $message = $channel->basic_get($queue); //取出消息
        if(!$message){
            //没有消息
            return false;
        }
        $res = $callback($message->body);
        if($res){
            //回传告诉 rabbitMQ 这个消息处理成功 清除此消息
            $channel->basic_ack($message->getDeliveryTag());
        }
        $channel->close();
        try {
            $connection->close();
        }catch (\Exception $e){
            return false;
        }
        return true;
    }

    /**
     * 消费多条消息
     * @param string $queue
     * @param Closure $callback
     * @param int $num
     * @return bool
     */
    public static function popMulti($queue, $callback, $num = 10){
        $connection = self::getConnect();
        $channel = $connection->channel();
        for($count = 0; $count < $num; $count++){
            $message = $channel->basic_get($queue); //取出消息
            if(!$message){
                //已经没有消息 返回成功
                return true;
            }
            $res = $callback($message->body);
            //如果失败了 需要在回调函数中记录日志 然后继续处理其他的消息，保持队列不中断
            if($res){
                //回传告诉 rabbitMQ 这个消息处理成功 清除此消息
                $channel->basic_ack($message->getDeliveryTag());
            }
            //不回传的话 本次处理消息会继续往下，但是下一次会重试处理没回传的消息
        }
        $channel->close();
        try {
            $connection->close();
        }catch (\Exception $e){
            return false;
        }
        return true;
    }

    public static function consumer(string $queueName, callable $callback, $num = 0){
        $connection = self::getConnect();
        $channel = $connection->channel();
        if($num) {
            $channel->basic_qos(null, $num, null);
        }
        $channel->basic_consume($queueName, '', false, false, false, false,
            $callback);

        while (count($channel->callbacks)) {
            try {
                $channel->wait();
            } catch (\ErrorException $e) {
                $insert = [
                    'queue_name'    => $queueName,
                    'code'          => $e->getCode(),
                    'message'       => $e->getMessage(),
                    'file'          => $e->getFile(),
                    'line'          => $e->getLine(),
                    'created_at'    => date('Y-m-d H:i:s'),
                ];
                MqErrorRecord::query()->insert($insert);
            }
        }
    }
}
