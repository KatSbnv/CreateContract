<?php
require_once(dirname(__FILE__) . '/wp-config.php');
$wp->init();
$wp->parse_request();
$wp->query_posts();
$wp->register_globals();
$wp->send_headers();

$source = array(
    'post_title' => $_POST["fio"],              // - заголовок материала.
    'post_name' => 'zagolovok-posta',                // - "слаг", синоним пути.
    'post_content' => $_POST["data"],    // - содержимое/контент.
    'post_status' => 'publish',                      // - статус материала: опубликованный.
    'post_author' => 1,                              // - автор материала: пользователь с id=1 (администратор).
    'post_type' => 'post',                           // - тип контента: запись.
);

# Вставка записи в базу данных:
$post_id = wp_insert_post($source);
echo $post_id;