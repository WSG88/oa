# Host: localhost  (Version: 5.7.26)
# Date: 2020-09-21 23:41:38
# Generator: MySQL-Front 5.3  (Build 4.234)

/*!40101 SET NAMES utf8 */;

#
# Structure for table "oatime"
#

DROP TABLE IF EXISTS `oatime`;
CREATE TABLE `oatime` (
  `Id` varchar(100) NOT NULL DEFAULT '',
  `name` varchar(255) DEFAULT NULL COMMENT '姓名',
  `day` varchar(255) DEFAULT NULL COMMENT '日期',
  `d1` varchar(255) DEFAULT NULL COMMENT '第1次考勤',
  `d2` varchar(255) DEFAULT NULL COMMENT '第2次考勤',
  `d3` varchar(255) DEFAULT NULL COMMENT '第3次考勤',
  `d4` varchar(255) DEFAULT NULL COMMENT '第4次考勤',
  `d5` varchar(255) DEFAULT NULL COMMENT '第5次考勤',
  `d6` varchar(255) DEFAULT NULL COMMENT '第6次考勤',
  `m` decimal(15,12) DEFAULT NULL COMMENT '上午',
  `a` decimal(15,12) DEFAULT NULL COMMENT '下午',
  `n` decimal(15,12) DEFAULT NULL COMMENT '晚上',
  `days` decimal(15,12) DEFAULT NULL COMMENT '出勤天数',
  `times` decimal(15,12) DEFAULT NULL COMMENT '加班时长',
  `room` int(11) DEFAULT NULL COMMENT '车间',
  PRIMARY KEY (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;
