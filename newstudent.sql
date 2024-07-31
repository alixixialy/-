-- phpMyAdmin SQL Dump
-- version 4.8.5
-- https://www.phpmyadmin.net/
--
-- 主机： localhost
-- 生成日期： 2024-07-29 23:34:28
-- 服务器版本： 5.7.26
-- PHP 版本： 7.4.3

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- 数据库： `www_newstudent_com`
--

-- --------------------------------------------------------

--
-- 表的结构 `newstudent`
--

CREATE TABLE `newstudent` (
  `name` varchar(255) NOT NULL COMMENT '名字',
  `sex` varchar(255) NOT NULL COMMENT '性别',
  `major` varchar(255) NOT NULL COMMENT '专业',
  `college` varchar(255) NOT NULL COMMENT '学院',
  `QQ` varchar(255) NOT NULL COMMENT 'QQ号',
  `number` varchar(255) NOT NULL COMMENT '电话号码',
  `notes` varchar(255) NOT NULL COMMENT '备注',
  `first_choice` varchar(255) NOT NULL,
  `$second_choice` varchar(255) NOT NULL,
  `$third_choice` varchar(255) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

--
-- 转储表的索引
--

--
-- 表的索引 `newstudent`
--
ALTER TABLE `newstudent`
  ADD PRIMARY KEY (`number`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
