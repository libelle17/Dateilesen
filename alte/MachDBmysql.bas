'Bauanleitung für eine Datenbank wie `//linux/mysql` vom 15.11.09 14:26:27
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 22, 42) As new CString, ArtZ&(3, 22)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "columns_priv"
 Str(0, 0, 1) = "`Host`"
 Str(0, 0, 2) = "`Db`"
 Str(0, 0, 3) = "`User`"
 Str(0, 0, 4) = "`Table_name`"
 Str(0, 0, 5) = "`Column_name`"
 Str(0, 0, 6) = "`Timestamp`"
 Str(0, 0, 7) = "`Column_priv`"
 Str(0, 0, 8) = "`Host`"
 ArtZ(0, 0) = 7
 ArtZ(1, 0) = 1
 Str(1, 0, 0) = "CREATE TABLE `columns_priv` ("
 Str(1, 0, 1) = " `Host` char(60) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 0, 2) = " `Db` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 0, 3) = " `User` char(16) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 0, 4) = " `Table_name` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 0, 5) = " `Column_name` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 0, 6) = " `Timestamp` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
 Str(1, 0, 7) = " `Column_priv` set('Select','Insert','Update','References') CHARACTER SET utf8 NOT NULL DEFAULT ''"
 Str(1, 0, 8) = "  PRIMARY KEY (`Host`,`Db`,`User`,`Table_name`,`Column_name`)"
 Str(1, 0, 9) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Column privileges'"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "db"
 Str(0, 1, 1) = "`Host`"
 Str(0, 1, 2) = "`Db`"
 Str(0, 1, 3) = "`User`"
 Str(0, 1, 4) = "`Select_priv`"
 Str(0, 1, 5) = "`Insert_priv`"
 Str(0, 1, 6) = "`Update_priv`"
 Str(0, 1, 7) = "`Delete_priv`"
 Str(0, 1, 8) = "`Create_priv`"
 Str(0, 1, 9) = "`Drop_priv`"
 Str(0, 1, 10) = "`Grant_priv`"
 Str(0, 1, 11) = "`References_priv`"
 Str(0, 1, 12) = "`Index_priv`"
 Str(0, 1, 13) = "`Alter_priv`"
 Str(0, 1, 14) = "`Create_tmp_table_priv`"
 Str(0, 1, 15) = "`Lock_tables_priv`"
 Str(0, 1, 16) = "`Create_view_priv`"
 Str(0, 1, 17) = "`Show_view_priv`"
 Str(0, 1, 18) = "`Create_routine_priv`"
 Str(0, 1, 19) = "`Alter_routine_priv`"
 Str(0, 1, 20) = "`Execute_priv`"
 Str(0, 1, 21) = "`Event_priv`"
 Str(0, 1, 22) = "`Trigger_priv`"
 Str(0, 1, 23) = "`Host`"
 Str(0, 1, 24) = "`User`"
 ArtZ(0, 1) = 22
 ArtZ(1, 1) = 2
 Str(1, 1, 0) = "CREATE TABLE `db` ("
 Str(1, 1, 1) = " `Host` char(60) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 1, 2) = " `Db` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 1, 3) = " `User` char(16) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 1, 4) = " `Select_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 5) = " `Insert_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 6) = " `Update_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 7) = " `Delete_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 8) = " `Create_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 9) = " `Drop_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 10) = " `Grant_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 11) = " `References_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 12) = " `Index_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 13) = " `Alter_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 14) = " `Create_tmp_table_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 15) = " `Lock_tables_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 16) = " `Create_view_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 17) = " `Show_view_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 18) = " `Create_routine_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 19) = " `Alter_routine_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 20) = " `Execute_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 21) = " `Event_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 22) = " `Trigger_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 1, 23) = "  PRIMARY KEY (`Host`,`Db`,`User`)"
 Str(1, 1, 24) = "  KEY `User` (`User`)"
 Str(1, 1, 25) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Database privileges'"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "event"
 Str(0, 2, 1) = "`db`"
 Str(0, 2, 2) = "`name`"
 Str(0, 2, 3) = "`body`"
 Str(0, 2, 4) = "`definer`"
 Str(0, 2, 5) = "`execute_at`"
 Str(0, 2, 6) = "`interval_value`"
 Str(0, 2, 7) = "`interval_field`"
 Str(0, 2, 8) = "`created`"
 Str(0, 2, 9) = "`modified`"
 Str(0, 2, 10) = "`last_executed`"
 Str(0, 2, 11) = "`starts`"
 Str(0, 2, 12) = "`ends`"
 Str(0, 2, 13) = "`status`"
 Str(0, 2, 14) = "`on_completion`"
 Str(0, 2, 15) = "`sql_mode`"
 Str(0, 2, 16) = "`comment`"
 Str(0, 2, 17) = "`originator`"
 Str(0, 2, 18) = "`time_zone`"
 Str(0, 2, 19) = "`character_set_client`"
 Str(0, 2, 20) = "`collation_connection`"
 Str(0, 2, 21) = "`db_collation`"
 Str(0, 2, 22) = "`body_utf8`"
 Str(0, 2, 23) = "`db`"
 ArtZ(0, 2) = 22
 ArtZ(1, 2) = 1
 Str(1, 2, 0) = "CREATE TABLE `event` ("
 Str(1, 2, 1) = " `db` char(64) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 2, 2) = " `name` char(64) NOT NULL DEFAULT ''"
 Str(1, 2, 3) = " `body` longblob NOT NULL"
 Str(1, 2, 4) = " `definer` char(77) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 2, 5) = " `execute_at` datetime DEFAULT NULL"
 Str(1, 2, 6) = " `interval_value` int(11) DEFAULT NULL"
 Str(1, 2, 7) = " `interval_field` enum('YEAR','QUARTER','MONTH','DAY','HOUR','MINUTE','WEEK','SECOND','MICROSECOND','YEAR_MONTH','DAY_HOUR','DAY_MINUTE','DAY_SECOND','HOUR_MINUTE','HOUR_SECOND','MINUTE_SECOND','DAY_MICROSECOND','HOUR_MICROSECOND','MINUTE_MICROSECOND','SECOND_MICROSECOND') DEFAULT NULL"
 Str(1, 2, 8) = " `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
 Str(1, 2, 9) = " `modified` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 2, 10) = " `last_executed` datetime DEFAULT NULL"
 Str(1, 2, 11) = " `starts` datetime DEFAULT NULL"
 Str(1, 2, 12) = " `ends` datetime DEFAULT NULL"
 Str(1, 2, 13) = " `status` enum('ENABLED','DISABLED','SLAVESIDE_DISABLED') NOT NULL DEFAULT 'ENABLED'"
 Str(1, 2, 14) = " `on_completion` enum('DROP','PRESERVE') NOT NULL DEFAULT 'DROP'"
 Str(1, 2, 15) = " `sql_mode` set('REAL_AS_FLOAT','PIPES_AS_CONCAT','ANSI_QUOTES','IGNORE_SPACE','NOT_USED','ONLY_FULL_GROUP_BY','NO_UNSIGNED_SUBTRACTION','NO_DIR_IN_CREATE','POSTGRESQL','ORACLE','MSSQL','DB2','MAXDB','NO_KEY_OPTIONS','NO_TABLE_OPTIONS','NO_FIELD_OPTIONS','MYSQL323','MYSQL40','ANSI','NO_AUTO_VALUE_ON_ZERO','NO_BACKSLASH_ESCAPES','STRICT_TRANS_TABLES','STRICT_ALL_TABLES','NO_ZERO_IN_DATE','NO_ZERO_DATE','INVALID_DATES','ERROR_FOR_DIVISION_BY_ZERO','TRADITIONAL','NO_AUTO_CREATE_USER','HIGH_NOT_PRECEDENCE','NO_ENGINE_SUBSTITUTION','PAD_CHAR_TO_FULL_LENGTH') NOT NULL DEFAULT ''"
 Str(1, 2, 16) = " `comment` char(64) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 2, 17) = " `originator` int(10) unsigned NOT NULL"
 Str(1, 2, 18) = " `time_zone` char(64) CHARACTER SET latin1 NOT NULL DEFAULT 'SYSTEM'"
 Str(1, 2, 19) = " `character_set_client` char(32) CHARACTER SET utf8 COLLATE utf8_bin DEFAULT NULL"
 Str(1, 2, 20) = " `collation_connection` char(32) CHARACTER SET utf8 COLLATE utf8_bin DEFAULT NULL"
 Str(1, 2, 21) = " `db_collation` char(32) CHARACTER SET utf8 COLLATE utf8_bin DEFAULT NULL"
 Str(1, 2, 22) = " `body_utf8` longblob"
 Str(1, 2, 23) = "  PRIMARY KEY (`db`,`name`)"
 Str(1, 2, 24) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='Events'"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "func"
 Str(0, 3, 1) = "`name`"
 Str(0, 3, 2) = "`ret`"
 Str(0, 3, 3) = "`dl`"
 Str(0, 3, 4) = "`type`"
 Str(0, 3, 5) = "`name`"
 ArtZ(0, 3) = 4
 ArtZ(1, 3) = 1
 Str(1, 3, 0) = "CREATE TABLE `func` ("
 Str(1, 3, 1) = " `name` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 3, 2) = " `ret` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 3, 3) = " `dl` char(128) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 3, 4) = " `type` enum('function','aggregate') CHARACTER SET utf8 NOT NULL"
 Str(1, 3, 5) = "  PRIMARY KEY (`name`)"
 Str(1, 3, 6) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='User defined functions'"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "general_log"
 Str(0, 4, 1) = "`event_time`"
 Str(0, 4, 2) = "`user_host`"
 Str(0, 4, 3) = "`thread_id`"
 Str(0, 4, 4) = "`server_id`"
 Str(0, 4, 5) = "`command_type`"
 Str(0, 4, 6) = "`argument`"
 ArtZ(0, 4) = 6
 Str(1, 4, 0) = "CREATE TABLE `general_log` ("
 Str(1, 4, 1) = " `event_time` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
 Str(1, 4, 2) = " `user_host` mediumtext NOT NULL"
 Str(1, 4, 3) = " `thread_id` int(11) NOT NULL"
 Str(1, 4, 4) = " `server_id` int(10) unsigned NOT NULL"
 Str(1, 4, 5) = " `command_type` varchar(64) NOT NULL"
 Str(1, 4, 6) = " `argument` mediumtext NOT NULL"
 Str(1, 4, 7) = " ENGINE=CSV DEFAULT CHARSET=utf8 COMMENT='General log'"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "help_category"
 Str(0, 5, 1) = "`help_category_id`"
 Str(0, 5, 2) = "`name`"
 Str(0, 5, 3) = "`parent_category_id`"
 Str(0, 5, 4) = "`url`"
 Str(0, 5, 5) = "`help_category_id`"
 Str(0, 5, 6) = "`name`"
 ArtZ(0, 5) = 4
 ArtZ(1, 5) = 2
 Str(1, 5, 0) = "CREATE TABLE `help_category` ("
 Str(1, 5, 1) = " `help_category_id` smallint(5) unsigned NOT NULL"
 Str(1, 5, 2) = " `name` char(64) NOT NULL"
 Str(1, 5, 3) = " `parent_category_id` smallint(5) unsigned DEFAULT NULL"
 Str(1, 5, 4) = " `url` char(128) NOT NULL"
 Str(1, 5, 5) = "  PRIMARY KEY (`help_category_id`)"
 Str(1, 5, 6) = "  UNIQUE KEY `name` (`name`)"
 Str(1, 5, 7) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='help categories'"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "help_keyword"
 Str(0, 6, 1) = "`help_keyword_id`"
 Str(0, 6, 2) = "`name`"
 Str(0, 6, 3) = "`help_keyword_id`"
 Str(0, 6, 4) = "`name`"
 ArtZ(0, 6) = 2
 ArtZ(1, 6) = 2
 Str(1, 6, 0) = "CREATE TABLE `help_keyword` ("
 Str(1, 6, 1) = " `help_keyword_id` int(10) unsigned NOT NULL"
 Str(1, 6, 2) = " `name` char(64) NOT NULL"
 Str(1, 6, 3) = "  PRIMARY KEY (`help_keyword_id`)"
 Str(1, 6, 4) = "  UNIQUE KEY `name` (`name`)"
 Str(1, 6, 5) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='help keywords'"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "help_relation"
 Str(0, 7, 1) = "`help_topic_id`"
 Str(0, 7, 2) = "`help_keyword_id`"
 Str(0, 7, 3) = "`help_keyword_id`"
 ArtZ(0, 7) = 2
 ArtZ(1, 7) = 1
 Str(1, 7, 0) = "CREATE TABLE `help_relation` ("
 Str(1, 7, 1) = " `help_topic_id` int(10) unsigned NOT NULL"
 Str(1, 7, 2) = " `help_keyword_id` int(10) unsigned NOT NULL"
 Str(1, 7, 3) = "  PRIMARY KEY (`help_keyword_id`,`help_topic_id`)"
 Str(1, 7, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='keyword-topic relation'"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "help_topic"
 Str(0, 8, 1) = "`help_topic_id`"
 Str(0, 8, 2) = "`name`"
 Str(0, 8, 3) = "`help_category_id`"
 Str(0, 8, 4) = "`description`"
 Str(0, 8, 5) = "`example`"
 Str(0, 8, 6) = "`url`"
 Str(0, 8, 7) = "`help_topic_id`"
 Str(0, 8, 8) = "`name`"
 ArtZ(0, 8) = 6
 ArtZ(1, 8) = 2
 Str(1, 8, 0) = "CREATE TABLE `help_topic` ("
 Str(1, 8, 1) = " `help_topic_id` int(10) unsigned NOT NULL"
 Str(1, 8, 2) = " `name` char(64) NOT NULL"
 Str(1, 8, 3) = " `help_category_id` smallint(5) unsigned NOT NULL"
 Str(1, 8, 4) = " `description` text NOT NULL"
 Str(1, 8, 5) = " `example` text NOT NULL"
 Str(1, 8, 6) = " `url` char(128) NOT NULL"
 Str(1, 8, 7) = "  PRIMARY KEY (`help_topic_id`)"
 Str(1, 8, 8) = "  UNIQUE KEY `name` (`name`)"
 Str(1, 8, 9) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='help topics'"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "host"
 Str(0, 9, 1) = "`Host`"
 Str(0, 9, 2) = "`Db`"
 Str(0, 9, 3) = "`Select_priv`"
 Str(0, 9, 4) = "`Insert_priv`"
 Str(0, 9, 5) = "`Update_priv`"
 Str(0, 9, 6) = "`Delete_priv`"
 Str(0, 9, 7) = "`Create_priv`"
 Str(0, 9, 8) = "`Drop_priv`"
 Str(0, 9, 9) = "`Grant_priv`"
 Str(0, 9, 10) = "`References_priv`"
 Str(0, 9, 11) = "`Index_priv`"
 Str(0, 9, 12) = "`Alter_priv`"
 Str(0, 9, 13) = "`Create_tmp_table_priv`"
 Str(0, 9, 14) = "`Lock_tables_priv`"
 Str(0, 9, 15) = "`Create_view_priv`"
 Str(0, 9, 16) = "`Show_view_priv`"
 Str(0, 9, 17) = "`Create_routine_priv`"
 Str(0, 9, 18) = "`Alter_routine_priv`"
 Str(0, 9, 19) = "`Execute_priv`"
 Str(0, 9, 20) = "`Trigger_priv`"
 Str(0, 9, 21) = "`Host`"
 ArtZ(0, 9) = 20
 ArtZ(1, 9) = 1
 Str(1, 9, 0) = "CREATE TABLE `host` ("
 Str(1, 9, 1) = " `Host` char(60) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 9, 2) = " `Db` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 9, 3) = " `Select_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 4) = " `Insert_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 5) = " `Update_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 6) = " `Delete_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 7) = " `Create_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 8) = " `Drop_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 9) = " `Grant_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 10) = " `References_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 11) = " `Index_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 12) = " `Alter_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 13) = " `Create_tmp_table_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 14) = " `Lock_tables_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 15) = " `Create_view_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 16) = " `Show_view_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 17) = " `Create_routine_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 18) = " `Alter_routine_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 19) = " `Execute_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 20) = " `Trigger_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 9, 21) = "  PRIMARY KEY (`Host`,`Db`)"
 Str(1, 9, 22) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Host privileges;  Merged with database privileges'"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "ndb_binlog_index"
 Str(0, 10, 1) = "`Position`"
 Str(0, 10, 2) = "`File`"
 Str(0, 10, 3) = "`epoch`"
 Str(0, 10, 4) = "`inserts`"
 Str(0, 10, 5) = "`updates`"
 Str(0, 10, 6) = "`deletes`"
 Str(0, 10, 7) = "`schemaops`"
 Str(0, 10, 8) = "`epoch`"
 ArtZ(0, 10) = 7
 ArtZ(1, 10) = 1
 Str(1, 10, 0) = "CREATE TABLE `ndb_binlog_index` ("
 Str(1, 10, 1) = " `Position` bigint(20) unsigned NOT NULL"
 Str(1, 10, 2) = " `File` varchar(255) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 10, 3) = " `epoch` bigint(20) unsigned NOT NULL"
 Str(1, 10, 4) = " `inserts` bigint(20) unsigned NOT NULL"
 Str(1, 10, 5) = " `updates` bigint(20) unsigned NOT NULL"
 Str(1, 10, 6) = " `deletes` bigint(20) unsigned NOT NULL"
 Str(1, 10, 7) = " `schemaops` bigint(20) unsigned NOT NULL"
 Str(1, 10, 8) = "  PRIMARY KEY (`epoch`)"
 Str(1, 10, 9) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr10

Sub FüllStr11()
 Str(0, 11, 0) = "plugin"
 Str(0, 11, 1) = "`name`"
 Str(0, 11, 2) = "`dl`"
 Str(0, 11, 3) = "`name`"
 ArtZ(0, 11) = 2
 ArtZ(1, 11) = 1
 Str(1, 11, 0) = "CREATE TABLE `plugin` ("
 Str(1, 11, 1) = " `name` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 11, 2) = " `dl` char(128) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 11, 3) = "  PRIMARY KEY (`name`)"
 Str(1, 11, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='MySQL plugins'"
End Sub ' FüllStr11

Sub FüllStr12()
 Str(0, 12, 0) = "proc"
 Str(0, 12, 1) = "`db`"
 Str(0, 12, 2) = "`name`"
 Str(0, 12, 3) = "`type`"
 Str(0, 12, 4) = "`specific_name`"
 Str(0, 12, 5) = "`language`"
 Str(0, 12, 6) = "`sql_data_access`"
 Str(0, 12, 7) = "`is_deterministic`"
 Str(0, 12, 8) = "`security_type`"
 Str(0, 12, 9) = "`param_list`"
 Str(0, 12, 10) = "`returns`"
 Str(0, 12, 11) = "`body`"
 Str(0, 12, 12) = "`definer`"
 Str(0, 12, 13) = "`created`"
 Str(0, 12, 14) = "`modified`"
 Str(0, 12, 15) = "`sql_mode`"
 Str(0, 12, 16) = "`comment`"
 Str(0, 12, 17) = "`character_set_client`"
 Str(0, 12, 18) = "`collation_connection`"
 Str(0, 12, 19) = "`db_collation`"
 Str(0, 12, 20) = "`body_utf8`"
 Str(0, 12, 21) = "`db`"
 ArtZ(0, 12) = 20
 ArtZ(1, 12) = 1
 Str(1, 12, 0) = "CREATE TABLE `proc` ("
 Str(1, 12, 1) = " `db` char(64) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 12, 2) = " `name` char(64) NOT NULL DEFAULT ''"
 Str(1, 12, 3) = " `type` enum('FUNCTION','PROCEDURE') NOT NULL"
 Str(1, 12, 4) = " `specific_name` char(64) NOT NULL DEFAULT ''"
 Str(1, 12, 5) = " `language` enum('SQL') NOT NULL DEFAULT 'SQL'"
 Str(1, 12, 6) = " `sql_data_access` enum('CONTAINS_SQL','NO_SQL','READS_SQL_DATA','MODIFIES_SQL_DATA') NOT NULL DEFAULT 'CONTAINS_SQL'"
 Str(1, 12, 7) = " `is_deterministic` enum('YES','NO') NOT NULL DEFAULT 'NO'"
 Str(1, 12, 8) = " `security_type` enum('INVOKER','DEFINER') NOT NULL DEFAULT 'DEFINER'"
 Str(1, 12, 9) = " `param_list` blob NOT NULL"
 Str(1, 12, 10) = " `returns` longblob NOT NULL"
 Str(1, 12, 11) = " `body` longblob NOT NULL"
 Str(1, 12, 12) = " `definer` char(77) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 12, 13) = " `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
 Str(1, 12, 14) = " `modified` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 12, 15) = " `sql_mode` set('REAL_AS_FLOAT','PIPES_AS_CONCAT','ANSI_QUOTES','IGNORE_SPACE','NOT_USED','ONLY_FULL_GROUP_BY','NO_UNSIGNED_SUBTRACTION','NO_DIR_IN_CREATE','POSTGRESQL','ORACLE','MSSQL','DB2','MAXDB','NO_KEY_OPTIONS','NO_TABLE_OPTIONS','NO_FIELD_OPTIONS','MYSQL323','MYSQL40','ANSI','NO_AUTO_VALUE_ON_ZERO','NO_BACKSLASH_ESCAPES','STRICT_TRANS_TABLES','STRICT_ALL_TABLES','NO_ZERO_IN_DATE','NO_ZERO_DATE','INVALID_DATES','ERROR_FOR_DIVISION_BY_ZERO','TRADITIONAL','NO_AUTO_CREATE_USER','HIGH_NOT_PRECEDENCE','NO_ENGINE_SUBSTITUTION','PAD_CHAR_TO_FULL_LENGTH') NOT NULL DEFAULT ''"
 Str(1, 12, 16) = " `comment` char(64) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 12, 17) = " `character_set_client` char(32) CHARACTER SET utf8 COLLATE utf8_bin DEFAULT NULL"
 Str(1, 12, 18) = " `collation_connection` char(32) CHARACTER SET utf8 COLLATE utf8_bin DEFAULT NULL"
 Str(1, 12, 19) = " `db_collation` char(32) CHARACTER SET utf8 COLLATE utf8_bin DEFAULT NULL"
 Str(1, 12, 20) = " `body_utf8` longblob"
 Str(1, 12, 21) = "  PRIMARY KEY (`db`,`name`,`type`)"
 Str(1, 12, 22) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='Stored Procedures'"
End Sub ' FüllStr12

Sub FüllStr13()
 Str(0, 13, 0) = "procs_priv"
 Str(0, 13, 1) = "`Host`"
 Str(0, 13, 2) = "`Db`"
 Str(0, 13, 3) = "`User`"
 Str(0, 13, 4) = "`Routine_name`"
 Str(0, 13, 5) = "`Routine_type`"
 Str(0, 13, 6) = "`Grantor`"
 Str(0, 13, 7) = "`Proc_priv`"
 Str(0, 13, 8) = "`Timestamp`"
 Str(0, 13, 9) = "`Host`"
 Str(0, 13, 10) = "`Grantor`"
 ArtZ(0, 13) = 8
 ArtZ(1, 13) = 2
 Str(1, 13, 0) = "CREATE TABLE `procs_priv` ("
 Str(1, 13, 1) = " `Host` char(60) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 13, 2) = " `Db` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 13, 3) = " `User` char(16) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 13, 4) = " `Routine_name` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 13, 5) = " `Routine_type` enum('FUNCTION','PROCEDURE') COLLATE utf8_bin NOT NULL"
 Str(1, 13, 6) = " `Grantor` char(77) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 13, 7) = " `Proc_priv` set('Execute','Alter Routine','Grant') CHARACTER SET utf8 NOT NULL DEFAULT ''"
 Str(1, 13, 8) = " `Timestamp` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
 Str(1, 13, 9) = "  PRIMARY KEY (`Host`,`Db`,`User`,`Routine_name`,`Routine_type`)"
 Str(1, 13, 10) = "  KEY `Grantor` (`Grantor`)"
 Str(1, 13, 11) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Procedure privileges'"
End Sub ' FüllStr13

Sub FüllStr14()
 Str(0, 14, 0) = "servers"
 Str(0, 14, 1) = "`Server_name`"
 Str(0, 14, 2) = "`Host`"
 Str(0, 14, 3) = "`Db`"
 Str(0, 14, 4) = "`Username`"
 Str(0, 14, 5) = "`Password`"
 Str(0, 14, 6) = "`Port`"
 Str(0, 14, 7) = "`Socket`"
 Str(0, 14, 8) = "`Wrapper`"
 Str(0, 14, 9) = "`Owner`"
 Str(0, 14, 10) = "`Server_name`"
 ArtZ(0, 14) = 9
 ArtZ(1, 14) = 1
 Str(1, 14, 0) = "CREATE TABLE `servers` ("
 Str(1, 14, 1) = " `Server_name` char(64) NOT NULL DEFAULT ''"
 Str(1, 14, 2) = " `Host` char(64) NOT NULL DEFAULT ''"
 Str(1, 14, 3) = " `Db` char(64) NOT NULL DEFAULT ''"
 Str(1, 14, 4) = " `Username` char(64) NOT NULL DEFAULT ''"
 Str(1, 14, 5) = " `Password` char(64) NOT NULL DEFAULT ''"
 Str(1, 14, 6) = " `Port` int(4) NOT NULL DEFAULT '0'"
 Str(1, 14, 7) = " `Socket` char(64) NOT NULL DEFAULT ''"
 Str(1, 14, 8) = " `Wrapper` char(64) NOT NULL DEFAULT ''"
 Str(1, 14, 9) = " `Owner` char(64) NOT NULL DEFAULT ''"
 Str(1, 14, 10) = "  PRIMARY KEY (`Server_name`)"
 Str(1, 14, 11) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='MySQL Foreign Servers table'"
End Sub ' FüllStr14

Sub FüllStr15()
 Str(0, 15, 0) = "slow_log"
 Str(0, 15, 1) = "`start_time`"
 Str(0, 15, 2) = "`user_host`"
 Str(0, 15, 3) = "`query_time`"
 Str(0, 15, 4) = "`lock_time`"
 Str(0, 15, 5) = "`rows_sent`"
 Str(0, 15, 6) = "`rows_examined`"
 Str(0, 15, 7) = "`db`"
 Str(0, 15, 8) = "`last_insert_id`"
 Str(0, 15, 9) = "`insert_id`"
 Str(0, 15, 10) = "`server_id`"
 Str(0, 15, 11) = "`sql_text`"
 ArtZ(0, 15) = 11
 Str(1, 15, 0) = "CREATE TABLE `slow_log` ("
 Str(1, 15, 1) = " `start_time` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
 Str(1, 15, 2) = " `user_host` mediumtext NOT NULL"
 Str(1, 15, 3) = " `query_time` time NOT NULL"
 Str(1, 15, 4) = " `lock_time` time NOT NULL"
 Str(1, 15, 5) = " `rows_sent` int(11) NOT NULL"
 Str(1, 15, 6) = " `rows_examined` int(11) NOT NULL"
 Str(1, 15, 7) = " `db` varchar(512) NOT NULL"
 Str(1, 15, 8) = " `last_insert_id` int(11) NOT NULL"
 Str(1, 15, 9) = " `insert_id` int(11) NOT NULL"
 Str(1, 15, 10) = " `server_id` int(10) unsigned NOT NULL"
 Str(1, 15, 11) = " `sql_text` mediumtext NOT NULL"
 Str(1, 15, 12) = " ENGINE=CSV DEFAULT CHARSET=utf8 COMMENT='Slow log'"
End Sub ' FüllStr15

Sub FüllStr16()
 Str(0, 16, 0) = "tables_priv"
 Str(0, 16, 1) = "`Host`"
 Str(0, 16, 2) = "`Db`"
 Str(0, 16, 3) = "`User`"
 Str(0, 16, 4) = "`Table_name`"
 Str(0, 16, 5) = "`Grantor`"
 Str(0, 16, 6) = "`Timestamp`"
 Str(0, 16, 7) = "`Table_priv`"
 Str(0, 16, 8) = "`Column_priv`"
 Str(0, 16, 9) = "`Host`"
 Str(0, 16, 10) = "`Grantor`"
 ArtZ(0, 16) = 8
 ArtZ(1, 16) = 2
 Str(1, 16, 0) = "CREATE TABLE `tables_priv` ("
 Str(1, 16, 1) = " `Host` char(60) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 16, 2) = " `Db` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 16, 3) = " `User` char(16) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 16, 4) = " `Table_name` char(64) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 16, 5) = " `Grantor` char(77) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 16, 6) = " `Timestamp` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
 Str(1, 16, 7) = " `Table_priv` set('Select','Insert','Update','Delete','Create','Drop','Grant','References','Index','Alter','Create View','Show view','Trigger') CHARACTER SET utf8 NOT NULL DEFAULT ''"
 Str(1, 16, 8) = " `Column_priv` set('Select','Insert','Update','References') CHARACTER SET utf8 NOT NULL DEFAULT ''"
 Str(1, 16, 9) = "  PRIMARY KEY (`Host`,`Db`,`User`,`Table_name`)"
 Str(1, 16, 10) = "  KEY `Grantor` (`Grantor`)"
 Str(1, 16, 11) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Table privileges'"
End Sub ' FüllStr16

Sub FüllStr17()
 Str(0, 17, 0) = "time_zone"
 Str(0, 17, 1) = "`Time_zone_id`"
 Str(0, 17, 2) = "`Use_leap_seconds`"
 Str(0, 17, 3) = "`Time_zone_id`"
 ArtZ(0, 17) = 2
 ArtZ(1, 17) = 1
 Str(1, 17, 0) = "CREATE TABLE `time_zone` ("
 Str(1, 17, 1) = " `Time_zone_id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 17, 2) = " `Use_leap_seconds` enum('Y','N') NOT NULL DEFAULT 'N'"
 Str(1, 17, 3) = "  PRIMARY KEY (`Time_zone_id`)"
 Str(1, 17, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='Time zones'"
End Sub ' FüllStr17

Sub FüllStr18()
 Str(0, 18, 0) = "time_zone_leap_second"
 Str(0, 18, 1) = "`Transition_time`"
 Str(0, 18, 2) = "`Correction`"
 Str(0, 18, 3) = "`Transition_time`"
 ArtZ(0, 18) = 2
 ArtZ(1, 18) = 1
 Str(1, 18, 0) = "CREATE TABLE `time_zone_leap_second` ("
 Str(1, 18, 1) = " `Transition_time` bigint(20) NOT NULL"
 Str(1, 18, 2) = " `Correction` int(11) NOT NULL"
 Str(1, 18, 3) = "  PRIMARY KEY (`Transition_time`)"
 Str(1, 18, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='Leap seconds information for time zones'"
End Sub ' FüllStr18

Sub FüllStr19()
 Str(0, 19, 0) = "time_zone_name"
 Str(0, 19, 1) = "`Name`"
 Str(0, 19, 2) = "`Time_zone_id`"
 Str(0, 19, 3) = "`Name`"
 ArtZ(0, 19) = 2
 ArtZ(1, 19) = 1
 Str(1, 19, 0) = "CREATE TABLE `time_zone_name` ("
 Str(1, 19, 1) = " `Name` char(64) NOT NULL"
 Str(1, 19, 2) = " `Time_zone_id` int(10) unsigned NOT NULL"
 Str(1, 19, 3) = "  PRIMARY KEY (`Name`)"
 Str(1, 19, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='Time zone names'"
End Sub ' FüllStr19

Sub FüllStr20()
 Str(0, 20, 0) = "time_zone_transition"
 Str(0, 20, 1) = "`Time_zone_id`"
 Str(0, 20, 2) = "`Transition_time`"
 Str(0, 20, 3) = "`Transition_type_id`"
 Str(0, 20, 4) = "`Time_zone_id`"
 ArtZ(0, 20) = 3
 ArtZ(1, 20) = 1
 Str(1, 20, 0) = "CREATE TABLE `time_zone_transition` ("
 Str(1, 20, 1) = " `Time_zone_id` int(10) unsigned NOT NULL"
 Str(1, 20, 2) = " `Transition_time` bigint(20) NOT NULL"
 Str(1, 20, 3) = " `Transition_type_id` int(10) unsigned NOT NULL"
 Str(1, 20, 4) = "  PRIMARY KEY (`Time_zone_id`,`Transition_time`)"
 Str(1, 20, 5) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='Time zone transitions'"
End Sub ' FüllStr20

Sub FüllStr21()
 Str(0, 21, 0) = "time_zone_transition_type"
 Str(0, 21, 1) = "`Time_zone_id`"
 Str(0, 21, 2) = "`Transition_type_id`"
 Str(0, 21, 3) = "`Offset`"
 Str(0, 21, 4) = "`Is_DST`"
 Str(0, 21, 5) = "`Abbreviation`"
 Str(0, 21, 6) = "`Time_zone_id`"
 ArtZ(0, 21) = 5
 ArtZ(1, 21) = 1
 Str(1, 21, 0) = "CREATE TABLE `time_zone_transition_type` ("
 Str(1, 21, 1) = " `Time_zone_id` int(10) unsigned NOT NULL"
 Str(1, 21, 2) = " `Transition_type_id` int(10) unsigned NOT NULL"
 Str(1, 21, 3) = " `Offset` int(11) NOT NULL DEFAULT '0'"
 Str(1, 21, 4) = " `Is_DST` tinyint(3) unsigned NOT NULL DEFAULT '0'"
 Str(1, 21, 5) = " `Abbreviation` char(8) NOT NULL DEFAULT ''"
 Str(1, 21, 6) = "  PRIMARY KEY (`Time_zone_id`,`Transition_type_id`)"
 Str(1, 21, 7) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='Time zone transition types'"
End Sub ' FüllStr21

Sub FüllStr22()
 Str(0, 22, 0) = "user"
 Str(0, 22, 1) = "`Host`"
 Str(0, 22, 2) = "`User`"
 Str(0, 22, 3) = "`Password`"
 Str(0, 22, 4) = "`Select_priv`"
 Str(0, 22, 5) = "`Insert_priv`"
 Str(0, 22, 6) = "`Update_priv`"
 Str(0, 22, 7) = "`Delete_priv`"
 Str(0, 22, 8) = "`Create_priv`"
 Str(0, 22, 9) = "`Drop_priv`"
 Str(0, 22, 10) = "`Reload_priv`"
 Str(0, 22, 11) = "`Shutdown_priv`"
 Str(0, 22, 12) = "`Process_priv`"
 Str(0, 22, 13) = "`File_priv`"
 Str(0, 22, 14) = "`Grant_priv`"
 Str(0, 22, 15) = "`References_priv`"
 Str(0, 22, 16) = "`Index_priv`"
 Str(0, 22, 17) = "`Alter_priv`"
 Str(0, 22, 18) = "`Show_db_priv`"
 Str(0, 22, 19) = "`Super_priv`"
 Str(0, 22, 20) = "`Create_tmp_table_priv`"
 Str(0, 22, 21) = "`Lock_tables_priv`"
 Str(0, 22, 22) = "`Execute_priv`"
 Str(0, 22, 23) = "`Repl_slave_priv`"
 Str(0, 22, 24) = "`Repl_client_priv`"
 Str(0, 22, 25) = "`Create_view_priv`"
 Str(0, 22, 26) = "`Show_view_priv`"
 Str(0, 22, 27) = "`Create_routine_priv`"
 Str(0, 22, 28) = "`Alter_routine_priv`"
 Str(0, 22, 29) = "`Create_user_priv`"
 Str(0, 22, 30) = "`Event_priv`"
 Str(0, 22, 31) = "`Trigger_priv`"
 Str(0, 22, 32) = "`ssl_type`"
 Str(0, 22, 33) = "`ssl_cipher`"
 Str(0, 22, 34) = "`x509_issuer`"
 Str(0, 22, 35) = "`x509_subject`"
 Str(0, 22, 36) = "`max_questions`"
 Str(0, 22, 37) = "`max_updates`"
 Str(0, 22, 38) = "`max_connections`"
 Str(0, 22, 39) = "`max_user_connections`"
 Str(0, 22, 40) = "`Host`"
 ArtZ(0, 22) = 39
 ArtZ(1, 22) = 1
 Str(1, 22, 0) = "CREATE TABLE `user` ("
 Str(1, 22, 1) = " `Host` char(60) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 22, 2) = " `User` char(16) COLLATE utf8_bin NOT NULL DEFAULT ''"
 Str(1, 22, 3) = " `Password` char(41) CHARACTER SET latin1 COLLATE latin1_bin NOT NULL DEFAULT ''"
 Str(1, 22, 4) = " `Select_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 5) = " `Insert_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 6) = " `Update_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 7) = " `Delete_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 8) = " `Create_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 9) = " `Drop_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 10) = " `Reload_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 11) = " `Shutdown_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 12) = " `Process_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 13) = " `File_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 14) = " `Grant_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 15) = " `References_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 16) = " `Index_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 17) = " `Alter_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 18) = " `Show_db_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 19) = " `Super_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 20) = " `Create_tmp_table_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 21) = " `Lock_tables_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 22) = " `Execute_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 23) = " `Repl_slave_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 24) = " `Repl_client_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 25) = " `Create_view_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 26) = " `Show_view_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 27) = " `Create_routine_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 28) = " `Alter_routine_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 29) = " `Create_user_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 30) = " `Event_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 31) = " `Trigger_priv` enum('N','Y') CHARACTER SET utf8 NOT NULL DEFAULT 'N'"
 Str(1, 22, 32) = " `ssl_type` enum('','ANY','X509','SPECIFIED') CHARACTER SET utf8 NOT NULL DEFAULT ''"
 Str(1, 22, 33) = " `ssl_cipher` blob NOT NULL"
 Str(1, 22, 34) = " `x509_issuer` blob NOT NULL"
 Str(1, 22, 35) = " `x509_subject` blob NOT NULL"
 Str(1, 22, 36) = " `max_questions` int(11) unsigned NOT NULL DEFAULT '0'"
 Str(1, 22, 37) = " `max_updates` int(11) unsigned NOT NULL DEFAULT '0'"
 Str(1, 22, 38) = " `max_connections` int(11) unsigned NOT NULL DEFAULT '0'"
 Str(1, 22, 39) = " `max_user_connections` int(11) unsigned NOT NULL DEFAULT '0'"
 Str(1, 22, 40) = "  PRIMARY KEY (`Host`,`User`)"
 Str(1, 22, 41) = " ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Users and global privileges'"
End Sub ' FüllStr22

Function doEx&(sql$, obtolerant%) ' SQL-Befehl ausführen, Fehler anzeigen
 Dim rAF&, FMeld$
 If obtolerant Then On Error Resume Next Else On Error GoTo fehler
 call cnz.Execute(sql, rAf)
 lErrNr = Err.Number
 FMeld = "Err.Nr " & lErrNr & ", rAf: " & rAF & " bei " & sql
 On Error GoTo fehler
 Debug.Print FMeld
 If obProt Then Print #302, FMeld
 DoEvents
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case Err.Number
 Case -2147467259 
  If InStrB(Err.Description, "nicht erzeugen") Then ' 'Kann Tabelle 'testDB1.faxe' nicht erzeugen (Fehler: 150)
   doEx = 150
   Exit Function
  ElseIf InStrB(Err.Description, "is not BASE TABLE") <> 0 Then
   doEx = 151
   Exit Function
  ElseIf InStrB(Err.Description, "MySQL server has gone away") Then
   cnz.Close
   cnz.Open
   Call doEx("use `" & hDBn & "`", 0)
   Resume
  End If
End Select
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doEx/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function ' doEx

Function SplitN&(ByVal q$, Sep$, erg$()) ' da Split() Speicher fraß
 Dim p1&, p2&, Slen&, obExit%, runde&
 On Error GoTo fehler
 If Not IsNull(q) Then
  Slen = Len(Sep)
  For runde = 1 To 2
   p2 = 0
   Do
    p1 = p2
    p2 = InStr(p1 + Slen, q, Sep)
    If p2 = 0 Then p2 = Len(q) + 1: obExit = True
    If p2 <> 0 Then
     If runde = 2 Then
      erg(SplitN) = Mid$(q, p1 + Slen, p2 - p1 - Slen)
     End If
     SplitN = SplitN + 1
    End If
    If obExit Then Exit Do
   Loop
   If runde = 1 Then
    ReDim erg(SplitN - 1)
    SplitN = 0
    obExit = 0
   End If
  Next runde
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SplitN/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' aufSplit

Public Function doMach_mysql(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
 Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$
 Dim i&, p1&, p2&, p3&, CLen&, CLen1&, obLT%
 Dim Index$()
 On Error Resume Next
 hDBn = DBn
 Open App.Path & "\MachDB.bas_prot.txt" For Output As #302
 obProt = (Err.Number = 0)
 On Error GoTo fehler
 If LenB(server) = 0 Then Server = GetServer(DbCn)
 cnzCStr = "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=" & Server & ";uid=mysql;pwd=97a5o6;"
 set cnz = Nothing
 cnz.open cnzCStr
 call doex("create database if not exists `" & DBN & "` character set latin1 collate latin1_german2_ci;",0)
 call doex("grant all privileges on `" & DBN & "`.* to 'praxis'@'%' identified by 'sonne' with grant option",0)
 call doex("grant all privileges on `" & DBN & "`.* to 'praxis'@'localhost' identified by 'sonne' with grant option",0)
 call doex("use `" & DBN & "`",0)
 call doex("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ",0)
 FüllStr0
 FüllStr1
 FüllStr2
 FüllStr3
 FüllStr4
 FüllStr5
 FüllStr6
 FüllStr7
 FüllStr8
 FüllStr9
 FüllStr10
 FüllStr11
 FüllStr12
 FüllStr13
 FüllStr14
 FüllStr15
 FüllStr16
 FüllStr17
 FüllStr18
 FüllStr19
 FüllStr20
 FüllStr21
 FüllStr22
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 22
  If Instr(Str(1, i, 0),"CREATE TABLE")<>0 then
   Tbl = Str(0, i, 0)
   ZZ = ArtZ(0, i) + ArtZ(1, i)
   sql = "CREATE TABLE IF NOT EXISTS `" & Tbl & "` (" & vbLf
   For j = 1 To ZZ
    sql.Append Str(1, i, j)
    If j < ZZ Then sql.Append "," & vbLf
   Next j
   ZZ = ZZ + ArtZ(2, i) + 1
   sql.Append vbLf & ")"
   sql.Append Str(1, i, ZZ)
   FNr = doEx(sql.Value, 0)
   Do
    set rsc = nothing
    rsc.Open "show create table `" & tbl & "`", cnz, adOpenStatic, adLockReadOnly
    sct = rsc.Fields(1)
    If InStrB(sct, "CREATE ALGORITHM") = 1 Then
     FNr = doEx("drop view `" & Tbl & "`", 0)
     FNr = doEx(sql.Value, 0)
    Else
     Exit Do
    End If
   Loop
   If InStrB(AIoZ(sct), AIoZ(Str(1, i, ZZ))) = 0 Then
    Call doEx("alter table `" & tbl & "`" & Str(1, i, ZZ), 0)
   End If
   TMt.Clear
   SplitN sct, vbLf, Spli
   For j = 1 To ArtZ(0, i) ' Tabellenfelder
    Dim k&, enthalten%, genau%, Posi$
    enthalten = 0
    genau = 0
    k = 0
    Set rsc = Nothing
    rsc.Open "show columns from `" & Tbl & "` where field = '" & Mid$(Str(0, i, j), 2, Len(Str(0, i, j)) - 2) & "'", cnz, adOpenStatic, adLockReadOnly
    enthalten = Not rsc.BOF
    If enthalten Then
     genau = (InStrB(sct, Str(1, i, j)) <> 0)
     If Not genau Then
      CLen = -1 ' Column-Length nicht kürzen
      obLT = (InStrB(sct, Str(0, i, j) & " longtext") <> 0)
      If Not obLT Then
       p1 = InStr(sct, "(")
       p2 = InStr(p1, sct, Str(0, i, j)) 'zCat.Tables(Tbl).Columns(k).Name & "`")
       If p2 = 0 Then p2 = InStr(p1, LCase$(sct), LCase(Str(0, i, j)))
       p1 = InStr(p2, sct, "(")
       p3 = InStr(p2, sct, ",")
       If p3 = 0 Then p3 = InStr(p2, sct, vbLf & ")")
       If p1 <> 0 And p1 < p3 Then
        p2 = InStr(p1, sct, ")")
        CLen = Mid(sct, p1 + 1, p2 - p1 - 1)
       End If
      End If
     End If
    End If
    If Not enthalten Or Not genau Then
     If j = 1 Then
      posi = " FIRST,"
     Else
      posi = " AFTER " & Str(0, i, j - 1) & ","
     End If
     If Not enthalten Then
      TMt.AppVar (Array(" add ", Str(1, i, j), posi))
     ElseIf Not genau Then
      If CLen <> -1 Or obLT Then
       p1 = InStr(Str(1, i, j), "(")
       If p1 <> 0 Then
        p2 = InStr(p1, Str(1, i, j), ")")
        If p2 <> 0 Then
         CLen1 = Mid(Str(1, i, j), p1 + 1, p2 - p1 - 1)
         If obLT Then
          Str(1, i, j).Replace "varchar(" & CLen1 & ")", "longtext"
         ElseIf CLen1 < CLen Then
          Str(1, i, j).Replace "(" & CLen1 & ")", "(" & CLen & ")"
         End If
         genau = (InStrB(sct, Str(1, i, j)) <> 0)
        End If
       End If
      End If
      If Not genau Then
       TMt.AppVar (Array(" modify ", Str(1, i, j), posi))
      End If
     End If
    End If
   Next j
   For j = ArtZ(0, i) + 1 To ArtZ(0, i) + ArtZ(1, i) ' Indices
    If InStrB(sct, Str(1, i, j)) = 0 Then
     If InStrB(Str(1, i, j).Value, "PRIMARY") <> 0 Then
      If InStrB(sct, "PRIMARY KEY (") <> 0 Then
       TMt.Append (" DROP PRIMARY KEY,")
      End If
     Else
      If InStrB(sct, "KEY " & Str(0, i, j).Value) <> 0 Then
       TMt.AppVar Array(" DROP KEY ", Str(0, i, j), ",")
      End If
     End If
     TMt.AppVar Array(" add ", Str(1, i, j), ",")
    End If
   Next j
   If TMt.Length <> 0 Then
    TMt.Cut (TMt.Length - 1)
    Call doEx("Alter Table `" & tbl & "` " & TMt.Value, -1)
   End If
  End If ' InStr(Str(1, i, 0), "CREATE TABLE") <> 0 Then
 Next i
 For i = 0 To 22
  If Instr(Str(1, i, 0),"CREATE TABLE")<>0 then
   Tbl = Str(0, i, 0)
   ZZ = ArtZ(0, i) + ArtZ(1, i)
   set rsc = nothing
   rsc.Open "show create table `" & tbl & "`", cnz, adOpenStatic, adLockReadOnly
   sct = rsc.Fields(1)
   ZZ = ZZ + ArtZ(2, i) + 1
   For j = ArtZ(0, i) + ArtZ(1, i) + 1 To ZZ - 1 'Constraints
    If InStrB(sct, Str(1, i, j)) = 0 Then
     If InStrB(sct, "CONSTRAINT " & Str(0, i, j)) <> 0 Then
      Call doEx("ALTER TABLE `" & Tbl & "` DROP FOREIGN KEY " & Str(0, i, j), 0)
     End If
     Call doEx("ALTER TABLE `" & Tbl & "` ADD" & Str(1, i, j), 0)
    End If
   Next j
  End If ' InStr(Str(1, i, 0), "CREATE TABLE") <> 0 Then
 Next i
 Dim runde%
 For runde = 0 to 4
  For i = 0 To 22
   If InStrB(Str(1, i, 0), "DEFINER VIEW") <> 0 Then
    Dim obCr%
    obCr = 0
    Set rsc = Nothing
    rsc.Open "show tables from `" & DBn & "` where `tables_in_" & DBn & "` = """ & Str(0, i, 0) & """", cnz, adOpenStatic, adLockReadOnly
    If rsc.BOF Then
     obCr = True
    Else
     Set rsc = Nothing
     rsc.Open "show create table `" & Str(0, i, 0) & "`", cnz, adOpenStatic, adLockReadOnly
     If rsc.Fields(1) <> Str(1, i, 0) Then
      Call doEx("DROP TABLE IF EXISTS `" & Str(0, i, 0) & "`", 0)
      Call doEx("DROP VIEW IF EXISTS `" & Str(0, i, 0) & "`", 0)
      obCr = True
     End If
    End If
    If obCr Then
     Call doEx(Str(1, i, 0).Value, True)
    End If
   End If
  Next i
 Next runde
 call doex("set FOREIGN_KEY_CHECKS = 1",0)
 If obProt then Close #302
 If not obstumm then
  MsgBox "Fertig mit doMach_mysql(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_mysql/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_mysql

Function GetServer$(DBCn As ADODB.Connection)
Dim spos&, sp2&
spos = InStr(LCase$(DBCn), "server=")
If spos <> 0 Then
 sp2 = InStr(spos, DBCn, ";")
 If sp2 = 0 Then sp2 = Len(DBCn)
 GetServer = Mid$(DBCn, spos + 7, sp2 - spos - 7)
End If
End Function ' GetServer

Function AIoZ(Ursp) As CString ' Ursp kann $ oder CString sein
 Const Such$ = "AUTO_INCREMENT="
 Set AIoZ = New CString
 AIoZ = Ursp
 Dim p0&, p1&
 p0 = AIoZ.Instr(Such)
 If p0 <> 0 Then
  p1 = AIoZ.Instr(" ", p0)
  AIoZ.Cut (p0 - 2)
  AIoZ.Append Mid(Ursp, p1)
 End If
End Function ' AIoZ(Ursp$) As CString
