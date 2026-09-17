'Bauanleitung für eine Datenbank wie `//linux/faxeinp` vom 15.11.09 14:26:14
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 9, 31) As new CString, ArtZ&(3, 9)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "dateien"
 Str(0, 0, 1) = "`id`"
 Str(0, 0, 2) = "`pfad`"
 Str(0, 0, 3) = "`geändert`"
 Str(0, 0, 4) = "`geprüft`"
 Str(0, 0, 5) = "`id`"
 Str(0, 0, 6) = "`pfad`"
 Str(0, 0, 7) = "`id`"
 ArtZ(0, 0) = 4
 ArtZ(1, 0) = 2
 ArtZ(2, 0) = 1
 Str(1, 0, 0) = "CREATE TABLE `dateien` ("
 Str(1, 0, 1) = " `id` int(2) unsigned NOT NULL COMMENT 'Bezug zu faxe'"
 Str(1, 0, 2) = " `pfad` varchar(500) COLLATE latin1_german2_ci NOT NULL COMMENT 'Dateiname samt Pfad'"
 Str(1, 0, 3) = " `geändert` datetime NOT NULL COMMENT 'Zeile eingetragen'"
 Str(1, 0, 4) = " `geprüft` tinyint(1) unsigned NOT NULL COMMENT 'Zeile geprüft'"
 Str(1, 0, 5) = "  KEY `id` (`id`)"
 Str(1, 0, 6) = "  KEY `pfad` (`pfad`)"
 Str(1, 0, 7) = "  CONSTRAINT `id` FOREIGN KEY (`id`) REFERENCES `faxe` (`ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 0, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "faxe"
 Str(0, 1, 1) = "`Name`"
 Str(0, 1, 2) = "`erstellt`"
 Str(0, 1, 3) = "`geändert`"
 Str(0, 1, 4) = "`Größe`"
 Str(0, 1, 5) = "`NameInP`"
 Str(0, 1, 6) = "`Telefonnummer`"
 Str(0, 1, 7) = "`Absender`"
 Str(0, 1, 8) = "`kopiert`"
 Str(0, 1, 9) = "`altNameInP`"
 Str(0, 1, 10) = "`ID`"
 Str(0, 1, 11) = "`TMStart`"
 Str(0, 1, 12) = "`TMEnd`"
 Str(0, 1, 13) = "`fFNr`"
 Str(0, 1, 14) = "`ID`"
 Str(0, 1, 15) = "`kopiertgeändertGröße`"
 Str(0, 1, 16) = "`kopiertName`"
 Str(0, 1, 17) = "`S2`"
 Str(0, 1, 18) = "`Telefonnummer`"
 Str(0, 1, 19) = "`geändertGröße`"
 Str(0, 1, 20) = "`Name`"
 Str(0, 1, 21) = "`NameinP`"
 ArtZ(0, 1) = 13
 ArtZ(1, 1) = 8
 Str(1, 1, 0) = "CREATE TABLE `faxe` ("
 Str(1, 1, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci NOT NULL COMMENT 'Name der Faxdatei'"
 Str(1, 1, 2) = " `erstellt` datetime DEFAULT NULL COMMENT 'Datum, an dem die Datei zuletzt erstellt wurde'"
 Str(1, 1, 3) = " `geändert` datetime DEFAULT NULL COMMENT 'Datum, an dem die Datei zuletzt geändert wurde'"
 Str(1, 1, 4) = " `Größe` int(10) DEFAULT NULL COMMENT 'Größe in Bytes'"
 Str(1, 1, 5) = " `NameInP` varchar(1000) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Neuer Name im Patientenordner'"
 Str(1, 1, 6) = " `Telefonnummer` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Telefonnummer des Senders'"
 Str(1, 1, 7) = " `Absender` varchar(1000) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Absender, falls bekannt'"
 Str(1, 1, 8) = " `kopiert` tinyint(1) DEFAULT NULL COMMENT 'ob Fax schon kopiert wurde'"
 Str(1, 1, 9) = " `altNameInP` varchar(1000) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Alter Name im Patientenordner'"
 Str(1, 1, 10) = " `ID` int(2) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 1, 11) = " `TMStart` datetime DEFAULT NULL COMMENT 'TransmissionStart'"
 Str(1, 1, 12) = " `TMEnd` datetime DEFAULT NULL COMMENT 'TransmissionEnd'"
 Str(1, 1, 13) = " `fFNr` tinyint(1) unsigned NOT NULL COMMENT 'faxFolder Nr. (1= IncomingQueue, 2 = IncomingArchive'"
 Str(1, 1, 14) = "  PRIMARY KEY (`ID`)"
 Str(1, 1, 15) = "  KEY `kopiertgeändertGröße` (`kopiert`,`geändert`,`Größe`)"
 Str(1, 1, 16) = "  KEY `kopiertName` (`kopiert`,`Name`)"
 Str(1, 1, 17) = "  KEY `S2` (`erstellt`,`Größe`)"
 Str(1, 1, 18) = "  KEY `Telefonnummer` (`Telefonnummer`)"
 Str(1, 1, 19) = "  KEY `geändertGröße` (`geändert`,`Größe`)"
 Str(1, 1, 20) = "  KEY `Name` (`Name`) USING BTREE"
 Str(1, 1, 21) = "  KEY `NameinP` (`NameInP`(767))"
 Str(1, 1, 22) = " ENGINE=InnoDB AUTO_INCREMENT=11710 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "inca"
 Str(0, 2, 1) = "`titel`"
 Str(0, 2, 2) = "`tsid`"
 Str(0, 2, 3) = "`transe`"
 Str(0, 2, 4) = "`transs`"
 Str(0, 2, 5) = "`id`"
 Str(0, 2, 6) = "`fsize`"
 Str(0, 2, 7) = "`pages`"
 Str(0, 2, 8) = "`devname`"
 Str(0, 2, 9) = "`retries`"
 Str(0, 2, 10) = "`csid`"
 Str(0, 2, 11) = "`routi`"
 Str(0, 2, 12) = "`callerid`"
 Str(0, 2, 13) = "`Id`"
 Str(0, 2, 14) = "`transe`"
 ArtZ(0, 2) = 12
 ArtZ(1, 2) = 2
 Str(1, 2, 0) = "CREATE TABLE `inca` ("
 Str(1, 2, 1) = " `titel` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 2) = " `tsid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 3) = " `transe` datetime DEFAULT NULL"
 Str(1, 2, 4) = " `transs` datetime DEFAULT NULL"
 Str(1, 2, 5) = " `id` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 6) = " `fsize` int(10) DEFAULT NULL"
 Str(1, 2, 7) = " `pages` int(10) DEFAULT NULL"
 Str(1, 2, 8) = " `devname` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 9) = " `retries` int(10) DEFAULT NULL"
 Str(1, 2, 10) = " `csid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 11) = " `routi` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 12) = " `callerid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 13) = "  KEY `Id` (`id`)"
 Str(1, 2, 14) = "  KEY `transe` (`transe`)"
 Str(1, 2, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "incq"
 Str(0, 3, 1) = "`titel`"
 Str(0, 3, 2) = "`tsid`"
 Str(0, 3, 3) = "`transe`"
 Str(0, 3, 4) = "`transs`"
 Str(0, 3, 5) = "`id`"
 Str(0, 3, 6) = "`fsize`"
 Str(0, 3, 7) = "`curp`"
 Str(0, 3, 8) = "`devid`"
 Str(0, 3, 9) = "`status`"
 Str(0, 3, 10) = "`exts`"
 Str(0, 3, 11) = "`extsc`"
 Str(0, 3, 12) = "`jobt`"
 Str(0, 3, 13) = "`retries`"
 Str(0, 3, 14) = "`routi`"
 Str(0, 3, 15) = "`csid`"
 Str(0, 3, 16) = "`avop`"
 Str(0, 3, 17) = "`callerid`"
 Str(0, 3, 18) = "`Id`"
 Str(0, 3, 19) = "`transs`"
 Str(0, 3, 20) = "`transe`"
 ArtZ(0, 3) = 17
 ArtZ(1, 3) = 3
 Str(1, 3, 0) = "CREATE TABLE `incq` ("
 Str(1, 3, 1) = " `titel` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 2) = " `tsid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 3) = " `transe` datetime DEFAULT NULL"
 Str(1, 3, 4) = " `transs` datetime DEFAULT NULL"
 Str(1, 3, 5) = " `id` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 6) = " `fsize` int(10) DEFAULT NULL"
 Str(1, 3, 7) = " `curp` int(10) DEFAULT NULL"
 Str(1, 3, 8) = " `devid` int(10) DEFAULT NULL"
 Str(1, 3, 9) = " `status` int(10) DEFAULT NULL"
 Str(1, 3, 10) = " `exts` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 11) = " `extsc` int(10) DEFAULT NULL"
 Str(1, 3, 12) = " `jobt` int(10) DEFAULT NULL"
 Str(1, 3, 13) = " `retries` int(10) DEFAULT NULL"
 Str(1, 3, 14) = " `routi` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 15) = " `csid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 16) = " `avop` int(10) DEFAULT NULL"
 Str(1, 3, 17) = " `callerid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 18) = "  KEY `Id` (`id`)"
 Str(1, 3, 19) = "  KEY `transs` (`transs`)"
 Str(1, 3, 20) = "  KEY `transe` (`transe`)"
 Str(1, 3, 21) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "outa"
 Str(0, 4, 1) = "`titel`"
 Str(0, 4, 2) = "`tsid`"
 Str(0, 4, 3) = "`submt`"
 Str(0, 4, 4) = "`submid`"
 Str(0, 4, 5) = "`oscht`"
 Str(0, 4, 6) = "`subject`"
 Str(0, 4, 7) = "`docname`"
 Str(0, 4, 8) = "`id`"
 Str(0, 4, 9) = "`fsize`"
 Str(0, 4, 10) = "`pages`"
 Str(0, 4, 11) = "`devname`"
 Str(0, 4, 12) = "`retries`"
 Str(0, 4, 13) = "`prio`"
 Str(0, 4, 14) = "`rcfax`"
 Str(0, 4, 15) = "`rcname`"
 Str(0, 4, 16) = "`csid`"
 Str(0, 4, 17) = "`sender`"
 Str(0, 4, 18) = "`transs`"
 Str(0, 4, 19) = "`transe`"
 Str(0, 4, 20) = "`Id`"
 Str(0, 4, 21) = "`submt`"
 Str(0, 4, 22) = "`oscht`"
 ArtZ(0, 4) = 19
 ArtZ(1, 4) = 3
 Str(1, 4, 0) = "CREATE TABLE `outa` ("
 Str(1, 4, 1) = " `titel` varchar(388) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 2) = " `tsid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 3) = " `submt` datetime DEFAULT NULL"
 Str(1, 4, 4) = " `submid` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 5) = " `oscht` datetime DEFAULT NULL"
 Str(1, 4, 6) = " `subject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 7) = " `docname` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 8) = " `id` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 9) = " `fsize` int(10) DEFAULT NULL"
 Str(1, 4, 10) = " `pages` int(10) DEFAULT NULL"
 Str(1, 4, 11) = " `devname` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 12) = " `retries` int(10) DEFAULT NULL"
 Str(1, 4, 13) = " `prio` int(10) DEFAULT NULL"
 Str(1, 4, 14) = " `rcfax` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 15) = " `rcname` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 16) = " `csid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 17) = " `sender` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 18) = " `transs` datetime DEFAULT NULL"
 Str(1, 4, 19) = " `transe` datetime DEFAULT NULL"
 Str(1, 4, 20) = "  KEY `Id` (`id`)"
 Str(1, 4, 21) = "  KEY `submt` (`submt`)"
 Str(1, 4, 22) = "  KEY `oscht` (`oscht`)"
 Str(1, 4, 23) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "outq"
 Str(0, 5, 1) = "`titel`"
 Str(0, 5, 2) = "`tsid`"
 Str(0, 5, 3) = "`submt`"
 Str(0, 5, 4) = "`submid`"
 Str(0, 5, 5) = "`scht`"
 Str(0, 5, 6) = "`oscht`"
 Str(0, 5, 7) = "`subject`"
 Str(0, 5, 8) = "`docname`"
 Str(0, 5, 9) = "`id`"
 Str(0, 5, 10) = "`fsize`"
 Str(0, 5, 11) = "`pages`"
 Str(0, 5, 12) = "`curp`"
 Str(0, 5, 13) = "`devid`"
 Str(0, 5, 14) = "`status`"
 Str(0, 5, 15) = "`exts`"
 Str(0, 5, 16) = "`extsc`"
 Str(0, 5, 17) = "`retries`"
 Str(0, 5, 18) = "`prio`"
 Str(0, 5, 19) = "`rct`"
 Str(0, 5, 20) = "`rcfax`"
 Str(0, 5, 21) = "`rcname`"
 Str(0, 5, 22) = "`csid`"
 Str(0, 5, 23) = "`avop`"
 Str(0, 5, 24) = "`gbr`"
 Str(0, 5, 25) = "`sender`"
 Str(0, 5, 26) = "`obalt`"
 Str(0, 5, 27) = "`Id`"
 Str(0, 5, 28) = "`submt`"
 Str(0, 5, 29) = "`scht`"
 ArtZ(0, 5) = 26
 ArtZ(1, 5) = 3
 Str(1, 5, 0) = "CREATE TABLE `outq` ("
 Str(1, 5, 1) = " `titel` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 2) = " `tsid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 3) = " `submt` datetime DEFAULT NULL"
 Str(1, 5, 4) = " `submid` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 5) = " `scht` datetime DEFAULT NULL"
 Str(1, 5, 6) = " `oscht` datetime DEFAULT NULL"
 Str(1, 5, 7) = " `subject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 8) = " `docname` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 9) = " `id` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 10) = " `fsize` int(10) DEFAULT NULL"
 Str(1, 5, 11) = " `pages` int(10) DEFAULT NULL"
 Str(1, 5, 12) = " `curp` int(10) DEFAULT NULL"
 Str(1, 5, 13) = " `devid` int(10) DEFAULT NULL"
 Str(1, 5, 14) = " `status` int(10) DEFAULT NULL"
 Str(1, 5, 15) = " `exts` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 16) = " `extsc` int(10) DEFAULT NULL"
 Str(1, 5, 17) = " `retries` int(10) DEFAULT NULL"
 Str(1, 5, 18) = " `prio` int(10) DEFAULT NULL"
 Str(1, 5, 19) = " `rct` int(10) DEFAULT NULL"
 Str(1, 5, 20) = " `rcfax` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 21) = " `rcname` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 22) = " `csid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 23) = " `avop` int(10) DEFAULT NULL"
 Str(1, 5, 24) = " `gbr` tinyint(1) DEFAULT NULL"
 Str(1, 5, 25) = " `sender` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 26) = " `obalt` tinyint(1) DEFAULT NULL"
 Str(1, 5, 27) = "  KEY `Id` (`id`)"
 Str(1, 5, 28) = "  KEY `submt` (`submt`)"
 Str(1, 5, 29) = "  KEY `scht` (`scht`)"
 Str(1, 5, 30) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "versorgungsamt oberfranken"
 Str(1, 6, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `versorgungsamt oberfranken` AS select `i`.`titel` AS `titel`,`i`.`tsid` AS `tsid`,`i`.`transe` AS `transe`,`i`.`transs` AS `transs`,`i`.`id` AS `id`,`i`.`fsize` AS `fsize`,`i`.`pages` AS `pages`,`i`.`devname` AS `devname`,`i`.`retries` AS `retries`,`i`.`csid` AS `csid`,`i`.`routi` AS `routi`,`i`.`callerid` AS `callerid` from `inca` `i` where ((`i`.`tsid` like '%803599%') and (`i`.`pages` > 1)) order by `i`.`transe` desc"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "verz"
 Str(1, 7, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `verz` AS select `vzo`.`zid` AS `id`,`vzo`.`fid` AS `fid`,`verze`.`verz` AS `verz`,`vzo`.`geändert` AS `geändert` from (`vzo` left join `verze` on((`vzo`.`vid` = `verze`.`id`)))"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "verze"
 Str(0, 8, 1) = "`id`"
 Str(0, 8, 2) = "`verz`"
 Str(0, 8, 3) = "`id`"
 Str(0, 8, 4) = "`verz`"
 ArtZ(0, 8) = 2
 ArtZ(1, 8) = 2
 Str(1, 8, 0) = "CREATE TABLE `verze` ("
 Str(1, 8, 1) = " `id` int(2) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf Verzeichnis'"
 Str(1, 8, 2) = " `verz` varchar(255) CHARACTER SET latin1 COLLATE latin1_german1_ci NOT NULL COMMENT 'Verzeichnis (eindeutig)'"
 Str(1, 8, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 8, 4) = "  UNIQUE KEY `verz` (`verz`)"
 Str(1, 8, 5) = " ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "vzo"
 Str(0, 9, 1) = "`zid`"
 Str(0, 9, 2) = "`fid`"
 Str(0, 9, 3) = "`vid`"
 Str(0, 9, 4) = "`geändert`"
 Str(0, 9, 5) = "`zid`"
 Str(0, 9, 6) = "`vzo`"
 Str(0, 9, 7) = "`geändert`"
 Str(0, 9, 8) = "`fid`"
 Str(0, 9, 9) = "`FKid`"
 Str(0, 9, 10) = "`vzo`"
 ArtZ(0, 9) = 4
 ArtZ(1, 9) = 4
 ArtZ(2, 9) = 2
 Str(1, 9, 0) = "CREATE TABLE `vzo` ("
 Str(1, 9, 1) = " `zid` int(2) unsigned NOT NULL AUTO_INCREMENT COMMENT 'eindeutige ID dieser Tabelle'"
 Str(1, 9, 2) = " `fid` int(2) unsigned NOT NULL COMMENT 'Bezug auf Faxe'"
 Str(1, 9, 3) = " `vid` int(2) unsigned NOT NULL COMMENT 'Bezug auf verze'"
 Str(1, 9, 4) = " `geändert` datetime NOT NULL COMMENT 'Datum des Eintrags'"
 Str(1, 9, 5) = "  PRIMARY KEY (`zid`)"
 Str(1, 9, 6) = "  KEY `vzo` (`vid`)"
 Str(1, 9, 7) = "  KEY `geändert` (`geändert`)"
 Str(1, 9, 8) = "  KEY `fid` (`fid`)"
 Str(1, 9, 9) = "  CONSTRAINT `FKid` FOREIGN KEY (`fid`) REFERENCES `faxe` (`ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 9, 10) = "  CONSTRAINT `vzo` FOREIGN KEY (`vid`) REFERENCES `verze` (`id`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 9, 11) = " ENGINE=InnoDB AUTO_INCREMENT=17845 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr9

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

Public Function doMach_faxeinp(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 9
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
 For i = 0 To 9
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
  For i = 0 To 9
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
  MsgBox "Fertig mit doMach_faxeinp(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_faxeinp/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_faxeinp

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
