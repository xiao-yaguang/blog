<?php

/** @var \Laravel\Lumen\Routing\Router $router */

/*
|--------------------------------------------------------------------------
| Application Routes
|--------------------------------------------------------------------------
|
| Here is where you can register all of the routes for an application.
| It is a breeze. Simply tell Lumen the URIs it should respond to
| and give it the Closure to call when that URI is requested.
|
*/

$router->get('/', function () use ($router) {
    return $router->app->version();
});

$router->group(['prefix' => 'post', 'namespace' => 'Post'], function () use ($router) {
    $router->post('post_export', 'PostController@postExport');
    $router->post('post_export_many', 'PostController@postExportMany');
    $router->post('post_export_table', 'PostController@postExportTable');
    $router->post('push_mq', 'PostController@pushMq');
    $router->post('pop_mq', 'PostController@popMq');
    $router->post('create_order', 'PostController@createOrder');
});
