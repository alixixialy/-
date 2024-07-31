<?php
// 数据库连接信息
$servername = "localhost";
$username = "111"; // 您的数据库用户名
$password = "111111"; // 您的数据库密码
$database = "WWW_newstudent_com"; // 您的数据库名

// 创建数据库连接
$GLOBALS['conn'] = new mysqli($servername, $username, $password, $database);

// 检查连接是否成功
if ($GLOBALS['conn']->connect_error) {
    die("连接失败: " . $GLOBALS['conn']->connect_error);
}