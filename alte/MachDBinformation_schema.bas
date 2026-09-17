'Bauanleitung für eine Datenbank wie `//linux/information_schema` vom 15.11.09 14:25:39
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 27, 40) As new CString, ArtZ&(3, 27)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "CHARACTER_SETS"
 Str(0, 0, 1) = "`CHARACTER_SET_NAME`"
 Str(0, 0, 2) = "`DEFAULT_COLLATE_NAME`"
 Str(0, 0, 3) = "`DESCRIPTION`"
 Str(0, 0, 4) = "`MAXLEN`"
 ArtZ(0, 0) = 4
 Str(1, 0, 0) = "CREATE TEMPORARY TABLE `CHARACTER_SETS` ("
 Str(1, 0, 1) = " `CHARACTER_SET_NAME` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 0, 2) = " `DEFAULT_COLLATE_NAME` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 0, 3) = " `DESCRIPTION` varchar(60) NOT NULL DEFAULT ''"
 Str(1, 0, 4) = " `MAXLEN` bigint(3) NOT NULL DEFAULT '0'"
 Str(1, 0, 5) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "COLLATIONS"
 Str(0, 1, 1) = "`COLLATION_NAME`"
 Str(0, 1, 2) = "`CHARACTER_SET_NAME`"
 Str(0, 1, 3) = "`ID`"
 Str(0, 1, 4) = "`IS_DEFAULT`"
 Str(0, 1, 5) = "`IS_COMPILED`"
 Str(0, 1, 6) = "`SORTLEN`"
 ArtZ(0, 1) = 6
 Str(1, 1, 0) = "CREATE TEMPORARY TABLE `COLLATIONS` ("
 Str(1, 1, 1) = " `COLLATION_NAME` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 1, 2) = " `CHARACTER_SET_NAME` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 1, 3) = " `ID` bigint(11) NOT NULL DEFAULT '0'"
 Str(1, 1, 4) = " `IS_DEFAULT` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 1, 5) = " `IS_COMPILED` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 1, 6) = " `SORTLEN` bigint(3) NOT NULL DEFAULT '0'"
 Str(1, 1, 7) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "COLLATION_CHARACTER_SET_APPLICABILITY"
 Str(0, 2, 1) = "`COLLATION_NAME`"
 Str(0, 2, 2) = "`CHARACTER_SET_NAME`"
 ArtZ(0, 2) = 2
 Str(1, 2, 0) = "CREATE TEMPORARY TABLE `COLLATION_CHARACTER_SET_APPLICABILITY` ("
 Str(1, 2, 1) = " `COLLATION_NAME` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 2, 2) = " `CHARACTER_SET_NAME` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 2, 3) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "COLUMNS"
 Str(0, 3, 1) = "`TABLE_CATALOG`"
 Str(0, 3, 2) = "`TABLE_SCHEMA`"
 Str(0, 3, 3) = "`TABLE_NAME`"
 Str(0, 3, 4) = "`COLUMN_NAME`"
 Str(0, 3, 5) = "`ORDINAL_POSITION`"
 Str(0, 3, 6) = "`COLUMN_DEFAULT`"
 Str(0, 3, 7) = "`IS_NULLABLE`"
 Str(0, 3, 8) = "`DATA_TYPE`"
 Str(0, 3, 9) = "`CHARACTER_MAXIMUM_LENGTH`"
 Str(0, 3, 10) = "`CHARACTER_OCTET_LENGTH`"
 Str(0, 3, 11) = "`NUMERIC_PRECISION`"
 Str(0, 3, 12) = "`NUMERIC_SCALE`"
 Str(0, 3, 13) = "`CHARACTER_SET_NAME`"
 Str(0, 3, 14) = "`COLLATION_NAME`"
 Str(0, 3, 15) = "`COLUMN_TYPE`"
 Str(0, 3, 16) = "`COLUMN_KEY`"
 Str(0, 3, 17) = "`EXTRA`"
 Str(0, 3, 18) = "`PRIVILEGES`"
 Str(0, 3, 19) = "`COLUMN_COMMENT`"
 ArtZ(0, 3) = 19
 Str(1, 3, 0) = "CREATE TEMPORARY TABLE `COLUMNS` ("
 Str(1, 3, 1) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 3, 2) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 3, 3) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 3, 4) = " `COLUMN_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 3, 5) = " `ORDINAL_POSITION` bigint(21) unsigned NOT NULL DEFAULT '0'"
 Str(1, 3, 6) = " `COLUMN_DEFAULT` longtext"
 Str(1, 3, 7) = " `IS_NULLABLE` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 3, 8) = " `DATA_TYPE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 3, 9) = " `CHARACTER_MAXIMUM_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 3, 10) = " `CHARACTER_OCTET_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 3, 11) = " `NUMERIC_PRECISION` bigint(21) unsigned DEFAULT NULL"
 Str(1, 3, 12) = " `NUMERIC_SCALE` bigint(21) unsigned DEFAULT NULL"
 Str(1, 3, 13) = " `CHARACTER_SET_NAME` varchar(32) DEFAULT NULL"
 Str(1, 3, 14) = " `COLLATION_NAME` varchar(32) DEFAULT NULL"
 Str(1, 3, 15) = " `COLUMN_TYPE` longtext NOT NULL"
 Str(1, 3, 16) = " `COLUMN_KEY` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 3, 17) = " `EXTRA` varchar(27) NOT NULL DEFAULT ''"
 Str(1, 3, 18) = " `PRIVILEGES` varchar(80) NOT NULL DEFAULT ''"
 Str(1, 3, 19) = " `COLUMN_COMMENT` varchar(255) NOT NULL DEFAULT ''"
 Str(1, 3, 20) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "COLUMN_PRIVILEGES"
 Str(0, 4, 1) = "`GRANTEE`"
 Str(0, 4, 2) = "`TABLE_CATALOG`"
 Str(0, 4, 3) = "`TABLE_SCHEMA`"
 Str(0, 4, 4) = "`TABLE_NAME`"
 Str(0, 4, 5) = "`COLUMN_NAME`"
 Str(0, 4, 6) = "`PRIVILEGE_TYPE`"
 Str(0, 4, 7) = "`IS_GRANTABLE`"
 ArtZ(0, 4) = 7
 Str(1, 4, 0) = "CREATE TEMPORARY TABLE `COLUMN_PRIVILEGES` ("
 Str(1, 4, 1) = " `GRANTEE` varchar(81) NOT NULL DEFAULT ''"
 Str(1, 4, 2) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 4, 3) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 4, 4) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 4, 5) = " `COLUMN_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 4, 6) = " `PRIVILEGE_TYPE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 4, 7) = " `IS_GRANTABLE` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 4, 8) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "ENGINES"
 Str(0, 5, 1) = "`ENGINE`"
 Str(0, 5, 2) = "`SUPPORT`"
 Str(0, 5, 3) = "`COMMENT`"
 Str(0, 5, 4) = "`TRANSACTIONS`"
 Str(0, 5, 5) = "`XA`"
 Str(0, 5, 6) = "`SAVEPOINTS`"
 ArtZ(0, 5) = 6
 Str(1, 5, 0) = "CREATE TEMPORARY TABLE `ENGINES` ("
 Str(1, 5, 1) = " `ENGINE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 5, 2) = " `SUPPORT` varchar(8) NOT NULL DEFAULT ''"
 Str(1, 5, 3) = " `COMMENT` varchar(80) NOT NULL DEFAULT ''"
 Str(1, 5, 4) = " `TRANSACTIONS` varchar(3) DEFAULT NULL"
 Str(1, 5, 5) = " `XA` varchar(3) DEFAULT NULL"
 Str(1, 5, 6) = " `SAVEPOINTS` varchar(3) DEFAULT NULL"
 Str(1, 5, 7) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "EVENTS"
 Str(0, 6, 1) = "`EVENT_CATALOG`"
 Str(0, 6, 2) = "`EVENT_SCHEMA`"
 Str(0, 6, 3) = "`EVENT_NAME`"
 Str(0, 6, 4) = "`DEFINER`"
 Str(0, 6, 5) = "`TIME_ZONE`"
 Str(0, 6, 6) = "`EVENT_BODY`"
 Str(0, 6, 7) = "`EVENT_DEFINITION`"
 Str(0, 6, 8) = "`EVENT_TYPE`"
 Str(0, 6, 9) = "`EXECUTE_AT`"
 Str(0, 6, 10) = "`INTERVAL_VALUE`"
 Str(0, 6, 11) = "`INTERVAL_FIELD`"
 Str(0, 6, 12) = "`SQL_MODE`"
 Str(0, 6, 13) = "`STARTS`"
 Str(0, 6, 14) = "`ENDS`"
 Str(0, 6, 15) = "`STATUS`"
 Str(0, 6, 16) = "`ON_COMPLETION`"
 Str(0, 6, 17) = "`CREATED`"
 Str(0, 6, 18) = "`LAST_ALTERED`"
 Str(0, 6, 19) = "`LAST_EXECUTED`"
 Str(0, 6, 20) = "`EVENT_COMMENT`"
 Str(0, 6, 21) = "`ORIGINATOR`"
 Str(0, 6, 22) = "`CHARACTER_SET_CLIENT`"
 Str(0, 6, 23) = "`COLLATION_CONNECTION`"
 Str(0, 6, 24) = "`DATABASE_COLLATION`"
 ArtZ(0, 6) = 24
 Str(1, 6, 0) = "CREATE TEMPORARY TABLE `EVENTS` ("
 Str(1, 6, 1) = " `EVENT_CATALOG` varchar(64) DEFAULT NULL"
 Str(1, 6, 2) = " `EVENT_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 6, 3) = " `EVENT_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 6, 4) = " `DEFINER` varchar(77) NOT NULL DEFAULT ''"
 Str(1, 6, 5) = " `TIME_ZONE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 6, 6) = " `EVENT_BODY` varchar(8) NOT NULL DEFAULT ''"
 Str(1, 6, 7) = " `EVENT_DEFINITION` longtext NOT NULL"
 Str(1, 6, 8) = " `EVENT_TYPE` varchar(9) NOT NULL DEFAULT ''"
 Str(1, 6, 9) = " `EXECUTE_AT` datetime DEFAULT NULL"
 Str(1, 6, 10) = " `INTERVAL_VALUE` varchar(256) DEFAULT NULL"
 Str(1, 6, 11) = " `INTERVAL_FIELD` varchar(18) DEFAULT NULL"
 Str(1, 6, 12) = " `SQL_MODE` varchar(8192) NOT NULL DEFAULT ''"
 Str(1, 6, 13) = " `STARTS` datetime DEFAULT NULL"
 Str(1, 6, 14) = " `ENDS` datetime DEFAULT NULL"
 Str(1, 6, 15) = " `STATUS` varchar(18) NOT NULL DEFAULT ''"
 Str(1, 6, 16) = " `ON_COMPLETION` varchar(12) NOT NULL DEFAULT ''"
 Str(1, 6, 17) = " `CREATED` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 6, 18) = " `LAST_ALTERED` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 6, 19) = " `LAST_EXECUTED` datetime DEFAULT NULL"
 Str(1, 6, 20) = " `EVENT_COMMENT` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 6, 21) = " `ORIGINATOR` bigint(10) NOT NULL DEFAULT '0'"
 Str(1, 6, 22) = " `CHARACTER_SET_CLIENT` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 6, 23) = " `COLLATION_CONNECTION` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 6, 24) = " `DATABASE_COLLATION` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 6, 25) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "FILES"
 Str(0, 7, 1) = "`FILE_ID`"
 Str(0, 7, 2) = "`FILE_NAME`"
 Str(0, 7, 3) = "`FILE_TYPE`"
 Str(0, 7, 4) = "`TABLESPACE_NAME`"
 Str(0, 7, 5) = "`TABLE_CATALOG`"
 Str(0, 7, 6) = "`TABLE_SCHEMA`"
 Str(0, 7, 7) = "`TABLE_NAME`"
 Str(0, 7, 8) = "`LOGFILE_GROUP_NAME`"
 Str(0, 7, 9) = "`LOGFILE_GROUP_NUMBER`"
 Str(0, 7, 10) = "`ENGINE`"
 Str(0, 7, 11) = "`FULLTEXT_KEYS`"
 Str(0, 7, 12) = "`DELETED_ROWS`"
 Str(0, 7, 13) = "`UPDATE_COUNT`"
 Str(0, 7, 14) = "`FREE_EXTENTS`"
 Str(0, 7, 15) = "`TOTAL_EXTENTS`"
 Str(0, 7, 16) = "`EXTENT_SIZE`"
 Str(0, 7, 17) = "`INITIAL_SIZE`"
 Str(0, 7, 18) = "`MAXIMUM_SIZE`"
 Str(0, 7, 19) = "`AUTOEXTEND_SIZE`"
 Str(0, 7, 20) = "`CREATION_TIME`"
 Str(0, 7, 21) = "`LAST_UPDATE_TIME`"
 Str(0, 7, 22) = "`LAST_ACCESS_TIME`"
 Str(0, 7, 23) = "`RECOVER_TIME`"
 Str(0, 7, 24) = "`TRANSACTION_COUNTER`"
 Str(0, 7, 25) = "`VERSION`"
 Str(0, 7, 26) = "`ROW_FORMAT`"
 Str(0, 7, 27) = "`TABLE_ROWS`"
 Str(0, 7, 28) = "`AVG_ROW_LENGTH`"
 Str(0, 7, 29) = "`DATA_LENGTH`"
 Str(0, 7, 30) = "`MAX_DATA_LENGTH`"
 Str(0, 7, 31) = "`INDEX_LENGTH`"
 Str(0, 7, 32) = "`DATA_FREE`"
 Str(0, 7, 33) = "`CREATE_TIME`"
 Str(0, 7, 34) = "`UPDATE_TIME`"
 Str(0, 7, 35) = "`CHECK_TIME`"
 Str(0, 7, 36) = "`CHECKSUM`"
 Str(0, 7, 37) = "`STATUS`"
 Str(0, 7, 38) = "`EXTRA`"
 ArtZ(0, 7) = 38
 Str(1, 7, 0) = "CREATE TEMPORARY TABLE `FILES` ("
 Str(1, 7, 1) = " `FILE_ID` bigint(4) NOT NULL DEFAULT '0'"
 Str(1, 7, 2) = " `FILE_NAME` varchar(64) DEFAULT NULL"
 Str(1, 7, 3) = " `FILE_TYPE` varchar(20) NOT NULL DEFAULT ''"
 Str(1, 7, 4) = " `TABLESPACE_NAME` varchar(64) DEFAULT NULL"
 Str(1, 7, 5) = " `TABLE_CATALOG` varchar(64) DEFAULT NULL"
 Str(1, 7, 6) = " `TABLE_SCHEMA` varchar(64) DEFAULT NULL"
 Str(1, 7, 7) = " `TABLE_NAME` varchar(64) DEFAULT NULL"
 Str(1, 7, 8) = " `LOGFILE_GROUP_NAME` varchar(64) DEFAULT NULL"
 Str(1, 7, 9) = " `LOGFILE_GROUP_NUMBER` bigint(4) DEFAULT NULL"
 Str(1, 7, 10) = " `ENGINE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 7, 11) = " `FULLTEXT_KEYS` varchar(64) DEFAULT NULL"
 Str(1, 7, 12) = " `DELETED_ROWS` bigint(4) DEFAULT NULL"
 Str(1, 7, 13) = " `UPDATE_COUNT` bigint(4) DEFAULT NULL"
 Str(1, 7, 14) = " `FREE_EXTENTS` bigint(4) DEFAULT NULL"
 Str(1, 7, 15) = " `TOTAL_EXTENTS` bigint(4) DEFAULT NULL"
 Str(1, 7, 16) = " `EXTENT_SIZE` bigint(4) NOT NULL DEFAULT '0'"
 Str(1, 7, 17) = " `INITIAL_SIZE` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 18) = " `MAXIMUM_SIZE` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 19) = " `AUTOEXTEND_SIZE` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 20) = " `CREATION_TIME` datetime DEFAULT NULL"
 Str(1, 7, 21) = " `LAST_UPDATE_TIME` datetime DEFAULT NULL"
 Str(1, 7, 22) = " `LAST_ACCESS_TIME` datetime DEFAULT NULL"
 Str(1, 7, 23) = " `RECOVER_TIME` bigint(4) DEFAULT NULL"
 Str(1, 7, 24) = " `TRANSACTION_COUNTER` bigint(4) DEFAULT NULL"
 Str(1, 7, 25) = " `VERSION` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 26) = " `ROW_FORMAT` varchar(10) DEFAULT NULL"
 Str(1, 7, 27) = " `TABLE_ROWS` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 28) = " `AVG_ROW_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 29) = " `DATA_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 30) = " `MAX_DATA_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 31) = " `INDEX_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 32) = " `DATA_FREE` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 33) = " `CREATE_TIME` datetime DEFAULT NULL"
 Str(1, 7, 34) = " `UPDATE_TIME` datetime DEFAULT NULL"
 Str(1, 7, 35) = " `CHECK_TIME` datetime DEFAULT NULL"
 Str(1, 7, 36) = " `CHECKSUM` bigint(21) unsigned DEFAULT NULL"
 Str(1, 7, 37) = " `STATUS` varchar(20) NOT NULL DEFAULT ''"
 Str(1, 7, 38) = " `EXTRA` varchar(255) DEFAULT NULL"
 Str(1, 7, 39) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "GLOBAL_STATUS"
 Str(0, 8, 1) = "`VARIABLE_NAME`"
 Str(0, 8, 2) = "`VARIABLE_VALUE`"
 ArtZ(0, 8) = 2
 Str(1, 8, 0) = "CREATE TEMPORARY TABLE `GLOBAL_STATUS` ("
 Str(1, 8, 1) = " `VARIABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 8, 2) = " `VARIABLE_VALUE` varchar(1024) DEFAULT NULL"
 Str(1, 8, 3) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "GLOBAL_VARIABLES"
 Str(0, 9, 1) = "`VARIABLE_NAME`"
 Str(0, 9, 2) = "`VARIABLE_VALUE`"
 ArtZ(0, 9) = 2
 Str(1, 9, 0) = "CREATE TEMPORARY TABLE `GLOBAL_VARIABLES` ("
 Str(1, 9, 1) = " `VARIABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 9, 2) = " `VARIABLE_VALUE` varchar(1024) DEFAULT NULL"
 Str(1, 9, 3) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "KEY_COLUMN_USAGE"
 Str(0, 10, 1) = "`CONSTRAINT_CATALOG`"
 Str(0, 10, 2) = "`CONSTRAINT_SCHEMA`"
 Str(0, 10, 3) = "`CONSTRAINT_NAME`"
 Str(0, 10, 4) = "`TABLE_CATALOG`"
 Str(0, 10, 5) = "`TABLE_SCHEMA`"
 Str(0, 10, 6) = "`TABLE_NAME`"
 Str(0, 10, 7) = "`COLUMN_NAME`"
 Str(0, 10, 8) = "`ORDINAL_POSITION`"
 Str(0, 10, 9) = "`POSITION_IN_UNIQUE_CONSTRAINT`"
 Str(0, 10, 10) = "`REFERENCED_TABLE_SCHEMA`"
 Str(0, 10, 11) = "`REFERENCED_TABLE_NAME`"
 Str(0, 10, 12) = "`REFERENCED_COLUMN_NAME`"
 ArtZ(0, 10) = 12
 Str(1, 10, 0) = "CREATE TEMPORARY TABLE `KEY_COLUMN_USAGE` ("
 Str(1, 10, 1) = " `CONSTRAINT_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 10, 2) = " `CONSTRAINT_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 10, 3) = " `CONSTRAINT_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 10, 4) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 10, 5) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 10, 6) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 10, 7) = " `COLUMN_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 10, 8) = " `ORDINAL_POSITION` bigint(10) NOT NULL DEFAULT '0'"
 Str(1, 10, 9) = " `POSITION_IN_UNIQUE_CONSTRAINT` bigint(10) DEFAULT NULL"
 Str(1, 10, 10) = " `REFERENCED_TABLE_SCHEMA` varchar(64) DEFAULT NULL"
 Str(1, 10, 11) = " `REFERENCED_TABLE_NAME` varchar(64) DEFAULT NULL"
 Str(1, 10, 12) = " `REFERENCED_COLUMN_NAME` varchar(64) DEFAULT NULL"
 Str(1, 10, 13) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr10

Sub FüllStr11()
 Str(0, 11, 0) = "PARTITIONS"
 Str(0, 11, 1) = "`TABLE_CATALOG`"
 Str(0, 11, 2) = "`TABLE_SCHEMA`"
 Str(0, 11, 3) = "`TABLE_NAME`"
 Str(0, 11, 4) = "`PARTITION_NAME`"
 Str(0, 11, 5) = "`SUBPARTITION_NAME`"
 Str(0, 11, 6) = "`PARTITION_ORDINAL_POSITION`"
 Str(0, 11, 7) = "`SUBPARTITION_ORDINAL_POSITION`"
 Str(0, 11, 8) = "`PARTITION_METHOD`"
 Str(0, 11, 9) = "`SUBPARTITION_METHOD`"
 Str(0, 11, 10) = "`PARTITION_EXPRESSION`"
 Str(0, 11, 11) = "`SUBPARTITION_EXPRESSION`"
 Str(0, 11, 12) = "`PARTITION_DESCRIPTION`"
 Str(0, 11, 13) = "`TABLE_ROWS`"
 Str(0, 11, 14) = "`AVG_ROW_LENGTH`"
 Str(0, 11, 15) = "`DATA_LENGTH`"
 Str(0, 11, 16) = "`MAX_DATA_LENGTH`"
 Str(0, 11, 17) = "`INDEX_LENGTH`"
 Str(0, 11, 18) = "`DATA_FREE`"
 Str(0, 11, 19) = "`CREATE_TIME`"
 Str(0, 11, 20) = "`UPDATE_TIME`"
 Str(0, 11, 21) = "`CHECK_TIME`"
 Str(0, 11, 22) = "`CHECKSUM`"
 Str(0, 11, 23) = "`PARTITION_COMMENT`"
 Str(0, 11, 24) = "`NODEGROUP`"
 Str(0, 11, 25) = "`TABLESPACE_NAME`"
 ArtZ(0, 11) = 25
 Str(1, 11, 0) = "CREATE TEMPORARY TABLE `PARTITIONS` ("
 Str(1, 11, 1) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 11, 2) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 11, 3) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 11, 4) = " `PARTITION_NAME` varchar(64) DEFAULT NULL"
 Str(1, 11, 5) = " `SUBPARTITION_NAME` varchar(64) DEFAULT NULL"
 Str(1, 11, 6) = " `PARTITION_ORDINAL_POSITION` bigint(21) unsigned DEFAULT NULL"
 Str(1, 11, 7) = " `SUBPARTITION_ORDINAL_POSITION` bigint(21) unsigned DEFAULT NULL"
 Str(1, 11, 8) = " `PARTITION_METHOD` varchar(12) DEFAULT NULL"
 Str(1, 11, 9) = " `SUBPARTITION_METHOD` varchar(12) DEFAULT NULL"
 Str(1, 11, 10) = " `PARTITION_EXPRESSION` longtext"
 Str(1, 11, 11) = " `SUBPARTITION_EXPRESSION` longtext"
 Str(1, 11, 12) = " `PARTITION_DESCRIPTION` longtext"
 Str(1, 11, 13) = " `TABLE_ROWS` bigint(21) unsigned NOT NULL DEFAULT '0'"
 Str(1, 11, 14) = " `AVG_ROW_LENGTH` bigint(21) unsigned NOT NULL DEFAULT '0'"
 Str(1, 11, 15) = " `DATA_LENGTH` bigint(21) unsigned NOT NULL DEFAULT '0'"
 Str(1, 11, 16) = " `MAX_DATA_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 11, 17) = " `INDEX_LENGTH` bigint(21) unsigned NOT NULL DEFAULT '0'"
 Str(1, 11, 18) = " `DATA_FREE` bigint(21) unsigned NOT NULL DEFAULT '0'"
 Str(1, 11, 19) = " `CREATE_TIME` datetime DEFAULT NULL"
 Str(1, 11, 20) = " `UPDATE_TIME` datetime DEFAULT NULL"
 Str(1, 11, 21) = " `CHECK_TIME` datetime DEFAULT NULL"
 Str(1, 11, 22) = " `CHECKSUM` bigint(21) unsigned DEFAULT NULL"
 Str(1, 11, 23) = " `PARTITION_COMMENT` varchar(80) NOT NULL DEFAULT ''"
 Str(1, 11, 24) = " `NODEGROUP` varchar(12) NOT NULL DEFAULT ''"
 Str(1, 11, 25) = " `TABLESPACE_NAME` varchar(64) DEFAULT NULL"
 Str(1, 11, 26) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr11

Sub FüllStr12()
 Str(0, 12, 0) = "PLUGINS"
 Str(0, 12, 1) = "`PLUGIN_NAME`"
 Str(0, 12, 2) = "`PLUGIN_VERSION`"
 Str(0, 12, 3) = "`PLUGIN_STATUS`"
 Str(0, 12, 4) = "`PLUGIN_TYPE`"
 Str(0, 12, 5) = "`PLUGIN_TYPE_VERSION`"
 Str(0, 12, 6) = "`PLUGIN_LIBRARY`"
 Str(0, 12, 7) = "`PLUGIN_LIBRARY_VERSION`"
 Str(0, 12, 8) = "`PLUGIN_AUTHOR`"
 Str(0, 12, 9) = "`PLUGIN_DESCRIPTION`"
 Str(0, 12, 10) = "`PLUGIN_LICENSE`"
 ArtZ(0, 12) = 10
 Str(1, 12, 0) = "CREATE TEMPORARY TABLE `PLUGINS` ("
 Str(1, 12, 1) = " `PLUGIN_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 12, 2) = " `PLUGIN_VERSION` varchar(20) NOT NULL DEFAULT ''"
 Str(1, 12, 3) = " `PLUGIN_STATUS` varchar(10) NOT NULL DEFAULT ''"
 Str(1, 12, 4) = " `PLUGIN_TYPE` varchar(80) NOT NULL DEFAULT ''"
 Str(1, 12, 5) = " `PLUGIN_TYPE_VERSION` varchar(20) NOT NULL DEFAULT ''"
 Str(1, 12, 6) = " `PLUGIN_LIBRARY` varchar(64) DEFAULT NULL"
 Str(1, 12, 7) = " `PLUGIN_LIBRARY_VERSION` varchar(20) DEFAULT NULL"
 Str(1, 12, 8) = " `PLUGIN_AUTHOR` varchar(64) DEFAULT NULL"
 Str(1, 12, 9) = " `PLUGIN_DESCRIPTION` longtext"
 Str(1, 12, 10) = " `PLUGIN_LICENSE` varchar(80) DEFAULT NULL"
 Str(1, 12, 11) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr12

Sub FüllStr13()
 Str(0, 13, 0) = "PROCESSLIST"
 Str(0, 13, 1) = "`ID`"
 Str(0, 13, 2) = "`USER`"
 Str(0, 13, 3) = "`HOST`"
 Str(0, 13, 4) = "`DB`"
 Str(0, 13, 5) = "`COMMAND`"
 Str(0, 13, 6) = "`TIME`"
 Str(0, 13, 7) = "`STATE`"
 Str(0, 13, 8) = "`INFO`"
 ArtZ(0, 13) = 8
 Str(1, 13, 0) = "CREATE TEMPORARY TABLE `PROCESSLIST` ("
 Str(1, 13, 1) = " `ID` bigint(4) NOT NULL DEFAULT '0'"
 Str(1, 13, 2) = " `USER` varchar(16) NOT NULL DEFAULT ''"
 Str(1, 13, 3) = " `HOST` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 13, 4) = " `DB` varchar(64) DEFAULT NULL"
 Str(1, 13, 5) = " `COMMAND` varchar(16) NOT NULL DEFAULT ''"
 Str(1, 13, 6) = " `TIME` int(7) NOT NULL DEFAULT '0'"
 Str(1, 13, 7) = " `STATE` varchar(64) DEFAULT NULL"
 Str(1, 13, 8) = " `INFO` longtext"
 Str(1, 13, 9) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr13

Sub FüllStr14()
 Str(0, 14, 0) = "PROFILING"
 Str(0, 14, 1) = "`QUERY_ID`"
 Str(0, 14, 2) = "`SEQ`"
 Str(0, 14, 3) = "`STATE`"
 Str(0, 14, 4) = "`DURATION`"
 Str(0, 14, 5) = "`CPU_USER`"
 Str(0, 14, 6) = "`CPU_SYSTEM`"
 Str(0, 14, 7) = "`CONTEXT_VOLUNTARY`"
 Str(0, 14, 8) = "`CONTEXT_INVOLUNTARY`"
 Str(0, 14, 9) = "`BLOCK_OPS_IN`"
 Str(0, 14, 10) = "`BLOCK_OPS_OUT`"
 Str(0, 14, 11) = "`MESSAGES_SENT`"
 Str(0, 14, 12) = "`MESSAGES_RECEIVED`"
 Str(0, 14, 13) = "`PAGE_FAULTS_MAJOR`"
 Str(0, 14, 14) = "`PAGE_FAULTS_MINOR`"
 Str(0, 14, 15) = "`SWAPS`"
 Str(0, 14, 16) = "`SOURCE_FUNCTION`"
 Str(0, 14, 17) = "`SOURCE_FILE`"
 Str(0, 14, 18) = "`SOURCE_LINE`"
 ArtZ(0, 14) = 18
 Str(1, 14, 0) = "CREATE TEMPORARY TABLE `PROFILING` ("
 Str(1, 14, 1) = " `QUERY_ID` int(20) NOT NULL DEFAULT '0'"
 Str(1, 14, 2) = " `SEQ` int(20) NOT NULL DEFAULT '0'"
 Str(1, 14, 3) = " `STATE` varchar(30) NOT NULL DEFAULT ''"
 Str(1, 14, 4) = " `DURATION` decimal(9,6) NOT NULL DEFAULT '0.000000'"
 Str(1, 14, 5) = " `CPU_USER` decimal(9,6) DEFAULT NULL"
 Str(1, 14, 6) = " `CPU_SYSTEM` decimal(9,6) DEFAULT NULL"
 Str(1, 14, 7) = " `CONTEXT_VOLUNTARY` int(20) DEFAULT NULL"
 Str(1, 14, 8) = " `CONTEXT_INVOLUNTARY` int(20) DEFAULT NULL"
 Str(1, 14, 9) = " `BLOCK_OPS_IN` int(20) DEFAULT NULL"
 Str(1, 14, 10) = " `BLOCK_OPS_OUT` int(20) DEFAULT NULL"
 Str(1, 14, 11) = " `MESSAGES_SENT` int(20) DEFAULT NULL"
 Str(1, 14, 12) = " `MESSAGES_RECEIVED` int(20) DEFAULT NULL"
 Str(1, 14, 13) = " `PAGE_FAULTS_MAJOR` int(20) DEFAULT NULL"
 Str(1, 14, 14) = " `PAGE_FAULTS_MINOR` int(20) DEFAULT NULL"
 Str(1, 14, 15) = " `SWAPS` int(20) DEFAULT NULL"
 Str(1, 14, 16) = " `SOURCE_FUNCTION` varchar(30) DEFAULT NULL"
 Str(1, 14, 17) = " `SOURCE_FILE` varchar(20) DEFAULT NULL"
 Str(1, 14, 18) = " `SOURCE_LINE` int(20) DEFAULT NULL"
 Str(1, 14, 19) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr14

Sub FüllStr15()
 Str(0, 15, 0) = "REFERENTIAL_CONSTRAINTS"
 Str(0, 15, 1) = "`CONSTRAINT_CATALOG`"
 Str(0, 15, 2) = "`CONSTRAINT_SCHEMA`"
 Str(0, 15, 3) = "`CONSTRAINT_NAME`"
 Str(0, 15, 4) = "`UNIQUE_CONSTRAINT_CATALOG`"
 Str(0, 15, 5) = "`UNIQUE_CONSTRAINT_SCHEMA`"
 Str(0, 15, 6) = "`UNIQUE_CONSTRAINT_NAME`"
 Str(0, 15, 7) = "`MATCH_OPTION`"
 Str(0, 15, 8) = "`UPDATE_RULE`"
 Str(0, 15, 9) = "`DELETE_RULE`"
 Str(0, 15, 10) = "`TABLE_NAME`"
 Str(0, 15, 11) = "`REFERENCED_TABLE_NAME`"
 ArtZ(0, 15) = 11
 Str(1, 15, 0) = "CREATE TEMPORARY TABLE `REFERENTIAL_CONSTRAINTS` ("
 Str(1, 15, 1) = " `CONSTRAINT_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 15, 2) = " `CONSTRAINT_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 15, 3) = " `CONSTRAINT_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 15, 4) = " `UNIQUE_CONSTRAINT_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 15, 5) = " `UNIQUE_CONSTRAINT_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 15, 6) = " `UNIQUE_CONSTRAINT_NAME` varchar(64) DEFAULT NULL"
 Str(1, 15, 7) = " `MATCH_OPTION` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 15, 8) = " `UPDATE_RULE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 15, 9) = " `DELETE_RULE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 15, 10) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 15, 11) = " `REFERENCED_TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 15, 12) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr15

Sub FüllStr16()
 Str(0, 16, 0) = "ROUTINES"
 Str(0, 16, 1) = "`SPECIFIC_NAME`"
 Str(0, 16, 2) = "`ROUTINE_CATALOG`"
 Str(0, 16, 3) = "`ROUTINE_SCHEMA`"
 Str(0, 16, 4) = "`ROUTINE_NAME`"
 Str(0, 16, 5) = "`ROUTINE_TYPE`"
 Str(0, 16, 6) = "`DTD_IDENTIFIER`"
 Str(0, 16, 7) = "`ROUTINE_BODY`"
 Str(0, 16, 8) = "`ROUTINE_DEFINITION`"
 Str(0, 16, 9) = "`EXTERNAL_NAME`"
 Str(0, 16, 10) = "`EXTERNAL_LANGUAGE`"
 Str(0, 16, 11) = "`PARAMETER_STYLE`"
 Str(0, 16, 12) = "`IS_DETERMINISTIC`"
 Str(0, 16, 13) = "`SQL_DATA_ACCESS`"
 Str(0, 16, 14) = "`SQL_PATH`"
 Str(0, 16, 15) = "`SECURITY_TYPE`"
 Str(0, 16, 16) = "`CREATED`"
 Str(0, 16, 17) = "`LAST_ALTERED`"
 Str(0, 16, 18) = "`SQL_MODE`"
 Str(0, 16, 19) = "`ROUTINE_COMMENT`"
 Str(0, 16, 20) = "`DEFINER`"
 Str(0, 16, 21) = "`CHARACTER_SET_CLIENT`"
 Str(0, 16, 22) = "`COLLATION_CONNECTION`"
 Str(0, 16, 23) = "`DATABASE_COLLATION`"
 ArtZ(0, 16) = 23
 Str(1, 16, 0) = "CREATE TEMPORARY TABLE `ROUTINES` ("
 Str(1, 16, 1) = " `SPECIFIC_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 16, 2) = " `ROUTINE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 16, 3) = " `ROUTINE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 16, 4) = " `ROUTINE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 16, 5) = " `ROUTINE_TYPE` varchar(9) NOT NULL DEFAULT ''"
 Str(1, 16, 6) = " `DTD_IDENTIFIER` varchar(64) DEFAULT NULL"
 Str(1, 16, 7) = " `ROUTINE_BODY` varchar(8) NOT NULL DEFAULT ''"
 Str(1, 16, 8) = " `ROUTINE_DEFINITION` longtext"
 Str(1, 16, 9) = " `EXTERNAL_NAME` varchar(64) DEFAULT NULL"
 Str(1, 16, 10) = " `EXTERNAL_LANGUAGE` varchar(64) DEFAULT NULL"
 Str(1, 16, 11) = " `PARAMETER_STYLE` varchar(8) NOT NULL DEFAULT ''"
 Str(1, 16, 12) = " `IS_DETERMINISTIC` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 16, 13) = " `SQL_DATA_ACCESS` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 16, 14) = " `SQL_PATH` varchar(64) DEFAULT NULL"
 Str(1, 16, 15) = " `SECURITY_TYPE` varchar(7) NOT NULL DEFAULT ''"
 Str(1, 16, 16) = " `CREATED` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 16, 17) = " `LAST_ALTERED` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 16, 18) = " `SQL_MODE` varchar(8192) NOT NULL DEFAULT ''"
 Str(1, 16, 19) = " `ROUTINE_COMMENT` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 16, 20) = " `DEFINER` varchar(77) NOT NULL DEFAULT ''"
 Str(1, 16, 21) = " `CHARACTER_SET_CLIENT` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 16, 22) = " `COLLATION_CONNECTION` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 16, 23) = " `DATABASE_COLLATION` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 16, 24) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr16

Sub FüllStr17()
 Str(0, 17, 0) = "SCHEMATA"
 Str(0, 17, 1) = "`CATALOG_NAME`"
 Str(0, 17, 2) = "`SCHEMA_NAME`"
 Str(0, 17, 3) = "`DEFAULT_CHARACTER_SET_NAME`"
 Str(0, 17, 4) = "`DEFAULT_COLLATION_NAME`"
 Str(0, 17, 5) = "`SQL_PATH`"
 ArtZ(0, 17) = 5
 Str(1, 17, 0) = "CREATE TEMPORARY TABLE `SCHEMATA` ("
 Str(1, 17, 1) = " `CATALOG_NAME` varchar(512) DEFAULT NULL"
 Str(1, 17, 2) = " `SCHEMA_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 17, 3) = " `DEFAULT_CHARACTER_SET_NAME` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 17, 4) = " `DEFAULT_COLLATION_NAME` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 17, 5) = " `SQL_PATH` varchar(512) DEFAULT NULL"
 Str(1, 17, 6) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr17

Sub FüllStr18()
 Str(0, 18, 0) = "SCHEMA_PRIVILEGES"
 Str(0, 18, 1) = "`GRANTEE`"
 Str(0, 18, 2) = "`TABLE_CATALOG`"
 Str(0, 18, 3) = "`TABLE_SCHEMA`"
 Str(0, 18, 4) = "`PRIVILEGE_TYPE`"
 Str(0, 18, 5) = "`IS_GRANTABLE`"
 ArtZ(0, 18) = 5
 Str(1, 18, 0) = "CREATE TEMPORARY TABLE `SCHEMA_PRIVILEGES` ("
 Str(1, 18, 1) = " `GRANTEE` varchar(81) NOT NULL DEFAULT ''"
 Str(1, 18, 2) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 18, 3) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 18, 4) = " `PRIVILEGE_TYPE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 18, 5) = " `IS_GRANTABLE` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 18, 6) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr18

Sub FüllStr19()
 Str(0, 19, 0) = "SESSION_STATUS"
 Str(0, 19, 1) = "`VARIABLE_NAME`"
 Str(0, 19, 2) = "`VARIABLE_VALUE`"
 ArtZ(0, 19) = 2
 Str(1, 19, 0) = "CREATE TEMPORARY TABLE `SESSION_STATUS` ("
 Str(1, 19, 1) = " `VARIABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 19, 2) = " `VARIABLE_VALUE` varchar(1024) DEFAULT NULL"
 Str(1, 19, 3) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr19

Sub FüllStr20()
 Str(0, 20, 0) = "SESSION_VARIABLES"
 Str(0, 20, 1) = "`VARIABLE_NAME`"
 Str(0, 20, 2) = "`VARIABLE_VALUE`"
 ArtZ(0, 20) = 2
 Str(1, 20, 0) = "CREATE TEMPORARY TABLE `SESSION_VARIABLES` ("
 Str(1, 20, 1) = " `VARIABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 20, 2) = " `VARIABLE_VALUE` varchar(1024) DEFAULT NULL"
 Str(1, 20, 3) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr20

Sub FüllStr21()
 Str(0, 21, 0) = "STATISTICS"
 Str(0, 21, 1) = "`TABLE_CATALOG`"
 Str(0, 21, 2) = "`TABLE_SCHEMA`"
 Str(0, 21, 3) = "`TABLE_NAME`"
 Str(0, 21, 4) = "`NON_UNIQUE`"
 Str(0, 21, 5) = "`INDEX_SCHEMA`"
 Str(0, 21, 6) = "`INDEX_NAME`"
 Str(0, 21, 7) = "`SEQ_IN_INDEX`"
 Str(0, 21, 8) = "`COLUMN_NAME`"
 Str(0, 21, 9) = "`COLLATION`"
 Str(0, 21, 10) = "`CARDINALITY`"
 Str(0, 21, 11) = "`SUB_PART`"
 Str(0, 21, 12) = "`PACKED`"
 Str(0, 21, 13) = "`NULLABLE`"
 Str(0, 21, 14) = "`INDEX_TYPE`"
 Str(0, 21, 15) = "`COMMENT`"
 ArtZ(0, 21) = 15
 Str(1, 21, 0) = "CREATE TEMPORARY TABLE `STATISTICS` ("
 Str(1, 21, 1) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 21, 2) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 21, 3) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 21, 4) = " `NON_UNIQUE` bigint(1) NOT NULL DEFAULT '0'"
 Str(1, 21, 5) = " `INDEX_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 21, 6) = " `INDEX_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 21, 7) = " `SEQ_IN_INDEX` bigint(2) NOT NULL DEFAULT '0'"
 Str(1, 21, 8) = " `COLUMN_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 21, 9) = " `COLLATION` varchar(1) DEFAULT NULL"
 Str(1, 21, 10) = " `CARDINALITY` bigint(21) DEFAULT NULL"
 Str(1, 21, 11) = " `SUB_PART` bigint(3) DEFAULT NULL"
 Str(1, 21, 12) = " `PACKED` varchar(10) DEFAULT NULL"
 Str(1, 21, 13) = " `NULLABLE` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 21, 14) = " `INDEX_TYPE` varchar(16) NOT NULL DEFAULT ''"
 Str(1, 21, 15) = " `COMMENT` varchar(16) DEFAULT NULL"
 Str(1, 21, 16) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr21

Sub FüllStr22()
 Str(0, 22, 0) = "TABLES"
 Str(0, 22, 1) = "`TABLE_CATALOG`"
 Str(0, 22, 2) = "`TABLE_SCHEMA`"
 Str(0, 22, 3) = "`TABLE_NAME`"
 Str(0, 22, 4) = "`TABLE_TYPE`"
 Str(0, 22, 5) = "`ENGINE`"
 Str(0, 22, 6) = "`VERSION`"
 Str(0, 22, 7) = "`ROW_FORMAT`"
 Str(0, 22, 8) = "`TABLE_ROWS`"
 Str(0, 22, 9) = "`AVG_ROW_LENGTH`"
 Str(0, 22, 10) = "`DATA_LENGTH`"
 Str(0, 22, 11) = "`MAX_DATA_LENGTH`"
 Str(0, 22, 12) = "`INDEX_LENGTH`"
 Str(0, 22, 13) = "`DATA_FREE`"
 Str(0, 22, 14) = "`AUTO_INCREMENT`"
 Str(0, 22, 15) = "`CREATE_TIME`"
 Str(0, 22, 16) = "`UPDATE_TIME`"
 Str(0, 22, 17) = "`CHECK_TIME`"
 Str(0, 22, 18) = "`TABLE_COLLATION`"
 Str(0, 22, 19) = "`CHECKSUM`"
 Str(0, 22, 20) = "`CREATE_OPTIONS`"
 Str(0, 22, 21) = "`TABLE_COMMENT`"
 ArtZ(0, 22) = 21
 Str(1, 22, 0) = "CREATE TEMPORARY TABLE `TABLES` ("
 Str(1, 22, 1) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 22, 2) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 22, 3) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 22, 4) = " `TABLE_TYPE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 22, 5) = " `ENGINE` varchar(64) DEFAULT NULL"
 Str(1, 22, 6) = " `VERSION` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 7) = " `ROW_FORMAT` varchar(10) DEFAULT NULL"
 Str(1, 22, 8) = " `TABLE_ROWS` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 9) = " `AVG_ROW_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 10) = " `DATA_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 11) = " `MAX_DATA_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 12) = " `INDEX_LENGTH` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 13) = " `DATA_FREE` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 14) = " `AUTO_INCREMENT` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 15) = " `CREATE_TIME` datetime DEFAULT NULL"
 Str(1, 22, 16) = " `UPDATE_TIME` datetime DEFAULT NULL"
 Str(1, 22, 17) = " `CHECK_TIME` datetime DEFAULT NULL"
 Str(1, 22, 18) = " `TABLE_COLLATION` varchar(32) DEFAULT NULL"
 Str(1, 22, 19) = " `CHECKSUM` bigint(21) unsigned DEFAULT NULL"
 Str(1, 22, 20) = " `CREATE_OPTIONS` varchar(255) DEFAULT NULL"
 Str(1, 22, 21) = " `TABLE_COMMENT` varchar(80) NOT NULL DEFAULT ''"
 Str(1, 22, 22) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr22

Sub FüllStr23()
 Str(0, 23, 0) = "TABLE_CONSTRAINTS"
 Str(0, 23, 1) = "`CONSTRAINT_CATALOG`"
 Str(0, 23, 2) = "`CONSTRAINT_SCHEMA`"
 Str(0, 23, 3) = "`CONSTRAINT_NAME`"
 Str(0, 23, 4) = "`TABLE_SCHEMA`"
 Str(0, 23, 5) = "`TABLE_NAME`"
 Str(0, 23, 6) = "`CONSTRAINT_TYPE`"
 ArtZ(0, 23) = 6
 Str(1, 23, 0) = "CREATE TEMPORARY TABLE `TABLE_CONSTRAINTS` ("
 Str(1, 23, 1) = " `CONSTRAINT_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 23, 2) = " `CONSTRAINT_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 23, 3) = " `CONSTRAINT_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 23, 4) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 23, 5) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 23, 6) = " `CONSTRAINT_TYPE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 23, 7) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr23

Sub FüllStr24()
 Str(0, 24, 0) = "TABLE_PRIVILEGES"
 Str(0, 24, 1) = "`GRANTEE`"
 Str(0, 24, 2) = "`TABLE_CATALOG`"
 Str(0, 24, 3) = "`TABLE_SCHEMA`"
 Str(0, 24, 4) = "`TABLE_NAME`"
 Str(0, 24, 5) = "`PRIVILEGE_TYPE`"
 Str(0, 24, 6) = "`IS_GRANTABLE`"
 ArtZ(0, 24) = 6
 Str(1, 24, 0) = "CREATE TEMPORARY TABLE `TABLE_PRIVILEGES` ("
 Str(1, 24, 1) = " `GRANTEE` varchar(81) NOT NULL DEFAULT ''"
 Str(1, 24, 2) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 24, 3) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 24, 4) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 24, 5) = " `PRIVILEGE_TYPE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 24, 6) = " `IS_GRANTABLE` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 24, 7) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr24

Sub FüllStr25()
 Str(0, 25, 0) = "TRIGGERS"
 Str(0, 25, 1) = "`TRIGGER_CATALOG`"
 Str(0, 25, 2) = "`TRIGGER_SCHEMA`"
 Str(0, 25, 3) = "`TRIGGER_NAME`"
 Str(0, 25, 4) = "`EVENT_MANIPULATION`"
 Str(0, 25, 5) = "`EVENT_OBJECT_CATALOG`"
 Str(0, 25, 6) = "`EVENT_OBJECT_SCHEMA`"
 Str(0, 25, 7) = "`EVENT_OBJECT_TABLE`"
 Str(0, 25, 8) = "`ACTION_ORDER`"
 Str(0, 25, 9) = "`ACTION_CONDITION`"
 Str(0, 25, 10) = "`ACTION_STATEMENT`"
 Str(0, 25, 11) = "`ACTION_ORIENTATION`"
 Str(0, 25, 12) = "`ACTION_TIMING`"
 Str(0, 25, 13) = "`ACTION_REFERENCE_OLD_TABLE`"
 Str(0, 25, 14) = "`ACTION_REFERENCE_NEW_TABLE`"
 Str(0, 25, 15) = "`ACTION_REFERENCE_OLD_ROW`"
 Str(0, 25, 16) = "`ACTION_REFERENCE_NEW_ROW`"
 Str(0, 25, 17) = "`CREATED`"
 Str(0, 25, 18) = "`SQL_MODE`"
 Str(0, 25, 19) = "`DEFINER`"
 Str(0, 25, 20) = "`CHARACTER_SET_CLIENT`"
 Str(0, 25, 21) = "`COLLATION_CONNECTION`"
 Str(0, 25, 22) = "`DATABASE_COLLATION`"
 ArtZ(0, 25) = 22
 Str(1, 25, 0) = "CREATE TEMPORARY TABLE `TRIGGERS` ("
 Str(1, 25, 1) = " `TRIGGER_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 25, 2) = " `TRIGGER_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 25, 3) = " `TRIGGER_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 25, 4) = " `EVENT_MANIPULATION` varchar(6) NOT NULL DEFAULT ''"
 Str(1, 25, 5) = " `EVENT_OBJECT_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 25, 6) = " `EVENT_OBJECT_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 25, 7) = " `EVENT_OBJECT_TABLE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 25, 8) = " `ACTION_ORDER` bigint(4) NOT NULL DEFAULT '0'"
 Str(1, 25, 9) = " `ACTION_CONDITION` longtext"
 Str(1, 25, 10) = " `ACTION_STATEMENT` longtext NOT NULL"
 Str(1, 25, 11) = " `ACTION_ORIENTATION` varchar(9) NOT NULL DEFAULT ''"
 Str(1, 25, 12) = " `ACTION_TIMING` varchar(6) NOT NULL DEFAULT ''"
 Str(1, 25, 13) = " `ACTION_REFERENCE_OLD_TABLE` varchar(64) DEFAULT NULL"
 Str(1, 25, 14) = " `ACTION_REFERENCE_NEW_TABLE` varchar(64) DEFAULT NULL"
 Str(1, 25, 15) = " `ACTION_REFERENCE_OLD_ROW` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 25, 16) = " `ACTION_REFERENCE_NEW_ROW` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 25, 17) = " `CREATED` datetime DEFAULT NULL"
 Str(1, 25, 18) = " `SQL_MODE` varchar(8192) NOT NULL DEFAULT ''"
 Str(1, 25, 19) = " `DEFINER` varchar(77) NOT NULL DEFAULT ''"
 Str(1, 25, 20) = " `CHARACTER_SET_CLIENT` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 25, 21) = " `COLLATION_CONNECTION` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 25, 22) = " `DATABASE_COLLATION` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 25, 23) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr25

Sub FüllStr26()
 Str(0, 26, 0) = "USER_PRIVILEGES"
 Str(0, 26, 1) = "`GRANTEE`"
 Str(0, 26, 2) = "`TABLE_CATALOG`"
 Str(0, 26, 3) = "`PRIVILEGE_TYPE`"
 Str(0, 26, 4) = "`IS_GRANTABLE`"
 ArtZ(0, 26) = 4
 Str(1, 26, 0) = "CREATE TEMPORARY TABLE `USER_PRIVILEGES` ("
 Str(1, 26, 1) = " `GRANTEE` varchar(81) NOT NULL DEFAULT ''"
 Str(1, 26, 2) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 26, 3) = " `PRIVILEGE_TYPE` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 26, 4) = " `IS_GRANTABLE` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 26, 5) = " ENGINE=MEMORY DEFAULT CHARSET=utf8"
End Sub ' FüllStr26

Sub FüllStr27()
 Str(0, 27, 0) = "VIEWS"
 Str(0, 27, 1) = "`TABLE_CATALOG`"
 Str(0, 27, 2) = "`TABLE_SCHEMA`"
 Str(0, 27, 3) = "`TABLE_NAME`"
 Str(0, 27, 4) = "`VIEW_DEFINITION`"
 Str(0, 27, 5) = "`CHECK_OPTION`"
 Str(0, 27, 6) = "`IS_UPDATABLE`"
 Str(0, 27, 7) = "`DEFINER`"
 Str(0, 27, 8) = "`SECURITY_TYPE`"
 Str(0, 27, 9) = "`CHARACTER_SET_CLIENT`"
 Str(0, 27, 10) = "`COLLATION_CONNECTION`"
 ArtZ(0, 27) = 10
 Str(1, 27, 0) = "CREATE TEMPORARY TABLE `VIEWS` ("
 Str(1, 27, 1) = " `TABLE_CATALOG` varchar(512) DEFAULT NULL"
 Str(1, 27, 2) = " `TABLE_SCHEMA` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 27, 3) = " `TABLE_NAME` varchar(64) NOT NULL DEFAULT ''"
 Str(1, 27, 4) = " `VIEW_DEFINITION` longtext NOT NULL"
 Str(1, 27, 5) = " `CHECK_OPTION` varchar(8) NOT NULL DEFAULT ''"
 Str(1, 27, 6) = " `IS_UPDATABLE` varchar(3) NOT NULL DEFAULT ''"
 Str(1, 27, 7) = " `DEFINER` varchar(77) NOT NULL DEFAULT ''"
 Str(1, 27, 8) = " `SECURITY_TYPE` varchar(7) NOT NULL DEFAULT ''"
 Str(1, 27, 9) = " `CHARACTER_SET_CLIENT` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 27, 10) = " `COLLATION_CONNECTION` varchar(32) NOT NULL DEFAULT ''"
 Str(1, 27, 11) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr27

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

Public Function doMach_information_schema(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
 FüllStr23
 FüllStr24
 FüllStr25
 FüllStr26
 FüllStr27
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 27
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
 For i = 0 To 27
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
  For i = 0 To 27
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
  MsgBox "Fertig mit doMach_information_schema(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_information_schema/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_information_schema

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
