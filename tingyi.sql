-- phpMyAdmin SQL Dump
-- version 5.2.0
-- https://www.phpmyadmin.net/
--
-- 主機： 127.0.0.1:3306
-- 產生時間： 2023-05-02 02:29:22
-- 伺服器版本： 5.7.40
-- PHP 版本： 7.4.33

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- 資料庫： `tingyi`
--
CREATE DATABASE IF NOT EXISTS `tingyi` DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci;
USE `tingyi`;

-- --------------------------------------------------------

--
-- 資料表結構 `power_item`
--

DROP TABLE IF EXISTS `power_item`;
CREATE TABLE IF NOT EXISTS `power_item` (
  `pi_sn` int(11) NOT NULL AUTO_INCREMENT,
  `pi_id` char(2) DEFAULT NULL,
  `pi_name` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`pi_sn`)
) ENGINE=MyISAM AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;

--
-- 傾印資料表的資料 `power_item`
--

INSERT INTO `power_item` (`pi_sn`, `pi_id`, `pi_name`) VALUES
(1, '0A', '員工資料管理'),
(2, '0B', '會員客戶管理'),
(3, '0C', '產品管理'),
(4, '0D', '報表列印管理');

-- --------------------------------------------------------

--
-- 資料表結構 `roles`
--

DROP TABLE IF EXISTS `roles`;
CREATE TABLE IF NOT EXISTS `roles` (
  `r_sn` int(11) NOT NULL AUTO_INCREMENT,
  `r_name` varchar(30) DEFAULT NULL,
  `r_power_item` text,
  PRIMARY KEY (`r_sn`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- 資料表結構 `staff`
--

DROP TABLE IF EXISTS `staff`;
CREATE TABLE IF NOT EXISTS `staff` (
  `s_sn` int(11) NOT NULL AUTO_INCREMENT,
  `s_name` varchar(30) DEFAULT NULL,
  `s_id` varchar(30) DEFAULT NULL,
  `s_pw` varchar(30) DEFAULT NULL,
  `s_tel` varchar(50) DEFAULT NULL,
  `s_mobile` varchar(50) DEFAULT NULL,
  `s_memo` varchar(300) DEFAULT NULL,
  PRIMARY KEY (`s_sn`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- 資料表結構 `sys_para`
--

DROP TABLE IF EXISTS `sys_para`;
CREATE TABLE IF NOT EXISTS `sys_para` (
  `sp_sn` int(11) NOT NULL AUTO_INCREMENT,
  `sp_name` varchar(30) DEFAULT NULL,
  `sp_type` varchar(10) DEFAULT NULL,
  `sp_option` varchar(300) DEFAULT NULL,
  PRIMARY KEY (`sp_sn`)
) ENGINE=MyISAM AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;

--
-- 傾印資料表的資料 `sys_para`
--

INSERT INTO `sys_para` (`sp_sn`, `sp_name`, `sp_type`, `sp_option`) VALUES
(1, '全酒', '單選', '清湯,1期,3期'),
(2, '半酒', '單選', '清湯,1期,3期');
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
