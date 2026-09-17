Attribute VB_Name = "Module1"
'Bauanleitung für eine Datenbank wie `//Linux1/quelle` vom 7.12.14 19:26:45
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.Connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 173, 212) As New CString, ArtZ&(3, 173)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "AlbCreMinMax"
 Str(1, 0, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `AlbCreMinMax` AS SELECT `n`.`Pat_ID` AS `pat_id`,max(if(((cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)) > cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2))) or isnull(`l1`.`Wert`)),cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)),cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2)))) AS `max`,min(if(((cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)) < cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2))) or isnull(`l1`.`Wert`)),cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)),cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2)))) AS `min` FROM ((`namen` `n` LEFT JOIN `labor2a` `l2` on(((`n`.`Pat_ID` = `l2`.`Pat_ID`) and (`l2`.`abk_ur` in ('albcre','a" & _
  "lbq','album','albup')) and (`l2`.`Langtext` like '%alb%') and (`l2`.`Einheit` like 'mg/g%') and (`l2`.`zeitpunkt` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date))))) LEFT JOIN `labor1a` `l1` on(((`n`.`Pat_ID` = `l1`.`Pat_ID`) and (`l1`.`abk_ur` in ('albcre','albq','album','albup')) and (`l1`.`Langtext` like '%alb%') and (`l1`.`Einheit` like 'mg/g%') and (`l1`.`ZeitPunkt` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date))))) WHERE (`n`.`Pat_ID` is not null) group by `n`.`Pat_ID` ORDER BY `n`.`Pat_ID`"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "CSII bei Typ 2"
 Str(1, 1, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `CSII bei Typ 2` AS SELECT `a`.`Pat_id` AS `pat_id`,`a`.`Nachname` AS `nachname`,`a`.`Vorname` AS `vorname`,`d`.`ICD` AS `icd` FROM (`anamnesebogen` `a` LEFT JOIN `diagnosen` `d` on(((`a`.`Pat_id` = `d`.`Pat_id`) and (`d`.`ICD` regexp '^E1[1234].|^O24.4')))) WHERE ((`a`.`Ther1` = 'CSII') and (`d`.`ICD` is not null))"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "DMP-HbA1c-Statistik"
 Str(1, 2, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `DMP-HbA1c-Statistik` AS SELECT sum(if((`h`.`letzter` >= 8.5),1,0)) AS `HbA1c>8,5`,sum(if((`h`.`letzter` is not null),1,0)) AS `HbA1c vorh`,count(0) AS `Zahl akt.Fälle`,round(((sum(if((`h`.`letzter` >= 8.5),1,0)) / sum(if((`h`.`letzter` is not null),1,0))) * 100),2) AS `Anteil_an_Vorh`,round(((sum(if((`h`.`letzter` >= 8.5),1,0)) / count(0)) * 100),2) AS `Anteil_an_Fallzahl` FROM ((`aktfv` `f` LEFT JOIN `namen` `n` on((`f`.`pat_id` = `n`.`Pat_ID`))) LEFT JOIN `lHbA1c` `h` on((`f`.`pat_id` = `h`.`pat_id`))) WHERE (`n`.`dmpklass` = 3)"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "DMPInkonsistenzen"
 Str(1, 3, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `DMPInkonsistenzen` AS SELECT `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr` FROM (`_lfaelle` `l` LEFT JOIN `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`)))) group by `f`.`Pat_ID`"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "EBM2010"
 Str(0, 4, 1) = "`MYID`"
 Str(0, 4, 2) = "`eingid`"
 Str(0, 4, 3) = "`Ziffer`"
 Str(0, 4, 4) = "`Leistungstext`"
 Str(0, 4, 5) = "`PWerte`"
 Str(0, 4, 6) = "`Euro`"
 Str(0, 4, 7) = "`MYID`"
 Str(0, 4, 8) = "`fk_L__`"
 Str(0, 4, 9) = "`Ziffer`"
 Str(0, 4, 10) = "`fk_L__`"
 ArtZ(0, 4) = 6
 ArtZ(1, 4) = 3
 ArtZ(2, 4) = 1
 Str(1, 4, 0) = "CREATE TABLE `EBM2010` ("
 Str(1, 4, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 4, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 4, 3) = " `Ziffer` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 4) = " `Leistungstext` varchar(236) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 5) = " `PWerte` varchar(57) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 6) = " `Euro` decimal(10,2) NOT NULL"
 Str(1, 4, 7) = "  PRIMARY KEY (`MYID`)"
 Str(1, 4, 8) = "  KEY `fk_L__` (`eingid`)"
 Str(1, 4, 9) = "  KEY `Ziffer` (`Ziffer`)"
 Str(1, 4, 10) = "  CONSTRAINT `fk_L__` FOREIGN KEY (`eingid`) REFERENCES `eingelesen1` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 4, 11) = " ENGINE=InnoDB AUTO_INCREMENT=2499 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: listenausgabe_ebm-ziffern.csv'"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "Frauenärzte"
 Str(0, 5, 1) = "`fg`"
 Str(0, 5, 2) = "`name`"
 Str(0, 5, 3) = "`vorname`"
 Str(0, 5, 4) = "`titelt`"
 Str(0, 5, 5) = "`fachgruppe`"
 Str(0, 5, 6) = "`strasse`"
 Str(0, 5, 7) = "`plz`"
 Str(0, 5, 8) = "`ort`"
 Str(0, 5, 9) = "`telefon`"
 Str(0, 5, 10) = "`fax`"
 Str(0, 5, 11) = "`kvnr`"
 Str(0, 5, 12) = "`aktdat`"
 Str(0, 5, 13) = "`id`"
 Str(0, 5, 14) = "`Ã¼berschrift`"
 Str(0, 5, 15) = "`dbnr`"
 Str(0, 5, 16) = "`bstelle`"
 Str(0, 5, 17) = "`anrede`"
 Str(0, 5, 18) = "`tel1`"
 Str(0, 5, 19) = "`tel2`"
 Str(0, 5, 20) = "`tel3`"
 Str(0, 5, 21) = "`tel4`"
 Str(0, 5, 22) = "`fax1`"
 Str(0, 5, 23) = "`fax2`"
 Str(0, 5, 24) = "`fax3`"
 Str(0, 5, 25) = "`email`"
 Str(0, 5, 26) = "`zulg`"
 Str(0, 5, 27) = "`arzttyp`"
 Str(0, 5, 28) = "`gemmit`"
 Str(0, 5, 29) = "`beme`"
 Str(0, 5, 30) = "`dmpt2`"
 Str(0, 5, 31) = "`dmpt1`"
 Str(0, 5, 32) = "`geschlecht`"
 Str(0, 5, 33) = "`titel`"
 ArtZ(0, 5) = 33
 Str(1, 5, 0) = "CREATE TABLE `Frauenärzte` ("
 Str(1, 5, 1) = " `fg` tinyint(4) NOT NULL"
 Str(1, 5, 2) = " `name` tinyint(4) NOT NULL"
 Str(1, 5, 3) = " `vorname` tinyint(4) NOT NULL"
 Str(1, 5, 4) = " `titelt` tinyint(4) NOT NULL"
 Str(1, 5, 5) = " `fachgruppe` tinyint(4) NOT NULL"
 Str(1, 5, 6) = " `strasse` tinyint(4) NOT NULL"
 Str(1, 5, 7) = " `plz` tinyint(4) NOT NULL"
 Str(1, 5, 8) = " `ort` tinyint(4) NOT NULL"
 Str(1, 5, 9) = " `telefon` tinyint(4) NOT NULL"
 Str(1, 5, 10) = " `fax` tinyint(4) NOT NULL"
 Str(1, 5, 11) = " `kvnr` tinyint(4) NOT NULL"
 Str(1, 5, 12) = " `aktdat` tinyint(4) NOT NULL"
 Str(1, 5, 13) = " `id` tinyint(4) NOT NULL"
 Str(1, 5, 14) = " `Ã¼berschrift` tinyint(4) NOT NULL"
 Str(1, 5, 15) = " `dbnr` tinyint(4) NOT NULL"
 Str(1, 5, 16) = " `bstelle` tinyint(4) NOT NULL"
 Str(1, 5, 17) = " `anrede` tinyint(4) NOT NULL"
 Str(1, 5, 18) = " `tel1` tinyint(4) NOT NULL"
 Str(1, 5, 19) = " `tel2` tinyint(4) NOT NULL"
 Str(1, 5, 20) = " `tel3` tinyint(4) NOT NULL"
 Str(1, 5, 21) = " `tel4` tinyint(4) NOT NULL"
 Str(1, 5, 22) = " `fax1` tinyint(4) NOT NULL"
 Str(1, 5, 23) = " `fax2` tinyint(4) NOT NULL"
 Str(1, 5, 24) = " `fax3` tinyint(4) NOT NULL"
 Str(1, 5, 25) = " `email` tinyint(4) NOT NULL"
 Str(1, 5, 26) = " `zulg` tinyint(4) NOT NULL"
 Str(1, 5, 27) = " `arzttyp` tinyint(4) NOT NULL"
 Str(1, 5, 28) = " `gemmit` tinyint(4) NOT NULL"
 Str(1, 5, 29) = " `beme` tinyint(4) NOT NULL"
 Str(1, 5, 30) = " `dmpt2` tinyint(4) NOT NULL"
 Str(1, 5, 31) = " `dmpt1` tinyint(4) NOT NULL"
 Str(1, 5, 32) = " `geschlecht` tinyint(4) NOT NULL"
 Str(1, 5, 33) = " `titel` tinyint(4) NOT NULL"
 Str(1, 5, 34) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "GFRMinMax"
 Str(1, 6, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `GFRMinMax` AS SELECT `n`.`Pat_ID` AS `pat_id`,floor((((pow(max(if(((cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)) > cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2))) or isnull(`l1`.`Wert`)),cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)),cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2)))),-(1.154)) * 186) * pow(((to_days(if(((cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)) > cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2))) or isnull(`l1`.`Wert`)),`l2`.`zeitpunkt`,`l1`.`ZeitPunkt`)) - to_days(`n`.`GebDat`)) * 0.00273792574745),-(0.203))) * if((`n`.`Geschlecht` = 'w'),0.742,1))) AS `gfr_max`,floor((((pow(min(if(((cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9" & _
  ",2)) < cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2))) or isnull(`l1`.`Wert`)),cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)),cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2)))),-(1.154)) * 186) * pow(((to_days(if(((cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)) < cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2))) or isnull(`l1`.`Wert`)),`l2`.`zeitpunkt`,`l1`.`ZeitPunkt`)) - to_days(`n`.`GebDat`)) * 0.00273792574745),-(0.203))) * if((`n`.`Geschlecht` = 'w'),0.742,1))) AS `gfr_min` FROM ((`namen` `n` LEFT JOIN `labor2a` `l2` on(((`n`.`Pat_ID` = `l2`.`Pat_ID`) and (`l2`.`abk_ur` in ('kre02','creat','krea','krea02')) and (not((`l2`.`Langtext` like '%urin%'))) and (`l2`.`Einheit` = 'mg/dl') and (`l2`.`zeitpunkt` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - " & _
  "interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date))))) LEFT JOIN `labor1a` `l1` on(((`n`.`Pat_ID` = `l1`.`Pat_ID`) and (`l1`.`abk_ur` in ('kre02','creat','krea','krea02')) and (not((`l1`.`Langtext` like '%urin%'))) and (`l1`.`Einheit` = 'mg/dl') and (`l1`.`ZeitPunkt` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date))))) WHERE (`n`.`Pat_ID` is not null) group by `n`.`Pat_ID` ORDER BY `n`.`Pat_ID`"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "GNRStat"
 Str(0, 7, 1) = "`id`"
 Str(0, 7, 2) = "`datei`"
 Str(0, 7, 3) = "`dateidat`"
 Str(0, 7, 4) = "`qinv`"
 Str(0, 7, 5) = "`id`"
 Str(0, 7, 6) = "`qinv`"
 ArtZ(0, 7) = 4
 ArtZ(1, 7) = 2
 Str(1, 7, 0) = "CREATE TABLE `GNRStat` ("
 Str(1, 7, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 7, 2) = " `datei` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 3) = " `dateidat` datetime DEFAULT NULL"
 Str(1, 7, 4) = " `qinv` int(5) DEFAULT NULL COMMENT 'Quartal nach Jahr'"
 Str(1, 7, 5) = "  PRIMARY KEY (`id`)"
 Str(1, 7, 6) = "  KEY `qinv` (`qinv`)"
 Str(1, 7, 7) = " ENGINE=InnoDB AUTO_INCREMENT=46 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "GNRZahl"
 Str(0, 8, 1) = "`id`"
 Str(0, 8, 2) = "`statid`"
 Str(0, 8, 3) = "`gnr`"
 Str(0, 8, 4) = "`leigru`"
 Str(0, 8, 5) = "`punkte`"
 Str(0, 8, 6) = "`euro`"
 Str(0, 8, 7) = "`m`"
 Str(0, 8, 8) = "`f`"
 Str(0, 8, 9) = "`r`"
 Str(0, 8, 10) = "`zahl`"
 Str(0, 8, 11) = "`wert`"
 Str(0, 8, 12) = "`uwert`"
 Str(0, 8, 13) = "`min`"
 Str(0, 8, 14) = "`id`"
 ArtZ(0, 8) = 13
 ArtZ(1, 8) = 1
 Str(1, 8, 0) = "CREATE TABLE `GNRZahl` ("
 Str(1, 8, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 8, 2) = " `statid` int(10) DEFAULT NULL"
 Str(1, 8, 3) = " `gnr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 4) = " `leigru` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 5) = " `punkte` int(5) DEFAULT NULL"
 Str(1, 8, 6) = " `euro` decimal(5,2) DEFAULT NULL"
 Str(1, 8, 7) = " `m` int(5) DEFAULT NULL"
 Str(1, 8, 8) = " `f` int(5) DEFAULT NULL"
 Str(1, 8, 9) = " `r` int(5) DEFAULT NULL"
 Str(1, 8, 10) = " `zahl` int(10) DEFAULT NULL"
 Str(1, 8, 11) = " `wert` decimal(9,2) DEFAULT NULL"
 Str(1, 8, 12) = " `uwert` decimal(9,2) DEFAULT NULL"
 Str(1, 8, 13) = " `min` int(10) DEFAULT NULL"
 Str(1, 8, 14) = "  PRIMARY KEY (`id`)"
 Str(1, 8, 15) = " ENGINE=InnoDB AUTO_INCREMENT=3915 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "HbA1cMinMax"
 Str(1, 9, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `HbA1cMinMax` AS SELECT `h`.`pat_id` AS `pat_id`,if(((`h`.`max1` > `h`.`max2`) or isnull(`h`.`max2`)),`h`.`max1`,`h`.`max2`) AS `max`,if(((`h`.`min1` < `h`.`min2`) or isnull(`h`.`min2`)),`h`.`min1`,`h`.`min2`) AS `min` FROM `_HbA1cMinMax` `h`"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "KreaMinMax"
 Str(1, 10, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `KreaMinMax` AS SELECT `n`.`Pat_ID` AS `pat_id`,max(if(((cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)) > cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2))) or isnull(`l1`.`Wert`)),cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)),cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2)))) AS `max`,min(if(((cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)) < cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2))) or isnull(`l1`.`Wert`)),cast(replace(concat('0',trim(`l2`.`Wert`)),',','.') as decimal(9,2)),cast(replace(concat('0',trim(`l1`.`Wert`)),',','.') as decimal(9,2)))) AS `min` FROM ((`namen` `n` LEFT JOIN `labor2a` `l2` on(((`n`.`Pat_ID` = `l2`.`Pat_ID`) and (`l2`.`abk_ur` in ('kre02','cre" & _
  "at','krea','krea02')) and (not((`l2`.`Langtext` like '%urin%'))) and (`l2`.`Einheit` = 'mg/dl') and (`l2`.`zeitpunkt` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date))))) LEFT JOIN `labor1a` `l1` on(((`n`.`Pat_ID` = `l1`.`Pat_ID`) and (`l1`.`abk_ur` in ('kre02','creat','krea','krea02')) and (not((`l1`.`Langtext` like '%urin%'))) and (`l1`.`Einheit` = 'mg/dl') and (`l1`.`ZeitPunkt` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date))))) WHERE (`n`.`Pat_ID` is not null) group by `n`.`Pat_ID` ORDER BY `n`.`Pat_ID`"
End Sub ' FüllStr10

Sub FüllStr11()
 Str(0, 11, 0) = "LaborDokumente eP"
 Str(1, 11, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `LaborDokumente eP` AS SELECT `dl`.`Pat_ID` AS `Pat_id`,`na`.`Nachname` AS `Nachname`,`na`.`Vorname` AS `Vorname`,`dl`.`ZeitPunkt` AS `ZeitPunkt`,`dl`.`DokName` AS `DokName`,`dl`.`absPos` AS `AbsPos`,`dl`.`AktZeit` AS `AktZeit`,`dl`.`DokPfad` AS `DokPfad`,`dl`.`DokGroe` AS `dokgroe`,`da`.`abgehakt` AS `abgehakt` FROM ((`dokumente` `dl` LEFT JOIN `dokumente abgehakt` `da` on((`dl`.`DokPfad` = `da`.`DokPfad`))) LEFT JOIN `namen` `na` on((`na`.`Pat_ID` = `dl`.`Pat_ID`))) WHERE ((`dl`.`DokName` like _latin1'*labor*') and (`dl`.`Pat_ID` = 2146)) ORDER BY `na`.`Nachname`,`na`.`Vorname`,`dl`.`ZeitPunkt` desc"
End Sub ' FüllStr11

Sub FüllStr12()
 Str(0, 12, 0) = "Schulungen 4.Quartal 2011"
 Str(1, 12, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `Schulungen 4.Quartal 2011` AS SELECT `e`.`Pat_ID` AS `pat_id`,concat(`n`.`Nachname`,' ',`n`.`Vorname`) AS `name`,`e`.`ZeitPunkt` AS `zeitpunkt`,`e`.`Inhalt` AS `inhalt`,`e`.`QS` AS `qs` FROM (`eintraege` `e` LEFT JOIN `namen` `n` on((`e`.`Pat_ID` = `n`.`Pat_ID`))) WHERE ((`e`.`Art` = 'schul') and (`e`.`ZeitPunkt` between '2011-10-01' and '2012-01-01')) ORDER BY `e`.`ZeitPunkt`"
End Sub ' FüllStr12

Sub FüllStr13()
 Str(0, 13, 0) = "Schulungen akt. Quartal"
 Str(1, 13, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `Schulungen akt. Quartal` AS SELECT `e`.`Pat_ID` AS `pat_id`,concat(`n`.`Nachname`,' ',`n`.`Vorname`) AS `name`,`e`.`ZeitPunkt` AS `zeitpunkt`,`e`.`Inhalt` AS `inhalt`,`e`.`QS` AS `qs` FROM (`eintraege` `e` LEFT JOIN `namen` `n` on((`e`.`Pat_ID` = `n`.`Pat_ID`))) WHERE ((`e`.`Art` = 'schul') and (`e`.`ZeitPunkt` between concat(year((now() - interval 29 day)),'-',(((quarter((now() - interval 29 day)) - 1) * 3) + 1),'-01') and concat((year((now() - interval 29 day)) + (quarter((now() - interval 29 day)) DIV 4)),'-',((((quarter((now() - interval 29 day)) - 1) * 3) + 4) % 12),'-01'))) ORDER BY `e`.`ZeitPunkt`"
End Sub ' FüllStr13

Sub FüllStr14()
 Str(0, 14, 0) = "Schulungsstatistik"
 Str(1, 14, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `Schulungsstatistik` AS SELECT year(`e`.`ZeitPunkt`) AS `Jahr`,(((month(`e`.`ZeitPunkt`) - 1) DIV 3) + 1) AS `Quartal`,count(0) AS `Zahl` FROM `eintraege` `e` WHERE (`e`.`Art` = 'schul') group by year(`e`.`ZeitPunkt`),(((month(`e`.`ZeitPunkt`) - 1) DIV 3) + 1)"
End Sub ' FüllStr14

Sub FüllStr15()
 Str(0, 15, 0) = "Trop-Tests"
 Str(0, 15, 1) = "`FID`"
 Str(0, 15, 2) = "`Pat_ID`"
 Str(0, 15, 3) = "`ZeitPunkt`"
 Str(0, 15, 4) = "`Art`"
 Str(0, 15, 5) = "`Inhalt`"
 Str(0, 15, 6) = "`absPos`"
 Str(0, 15, 7) = "`AktZeit`"
 Str(0, 15, 8) = "`QS`"
 Str(0, 15, 9) = "`QT`"
 Str(0, 15, 10) = "`StByte`"
 ArtZ(0, 15) = 10
 Str(1, 15, 0) = "CREATE TABLE `Trop-Tests` ("
 Str(1, 15, 1) = " `FID` int(10) DEFAULT NULL"
 Str(1, 15, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 15, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 15, 4) = " `Art` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 15, 5) = " `Inhalt` longtext COLLATE latin1_german2_ci"
 Str(1, 15, 6) = " `absPos` int(10) DEFAULT NULL"
 Str(1, 15, 7) = " `AktZeit` datetime DEFAULT NULL"
 Str(1, 15, 8) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 15, 9) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 15, 10) = " `StByte` int(10) DEFAULT NULL"
 Str(1, 15, 11) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr15

Sub FüllStr16()
 Str(0, 16, 0) = "_HbA1cMinMax"
 Str(1, 16, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_HbA1cMinMax` AS SELECT `n`.`Pat_ID` AS `pat_id`,(SELECT max(cast(replace(concat('0',trim(`ln`.`Wert`)),',','.') as decimal(9,2))) FROM `laborneu` `ln` WHERE ((`ln`.`Pat_ID` = `n`.`Pat_ID`) and (`ln`.`Abkü` regexp '^hba[1c]') and (cast(replace(concat('0',trim(`ln`.`Wert`)),',','.') as decimal(9,2)) < 21) and (`ln`.`Wert` regexp '^[0-9 .,]*$') and (`ln`.`ZeitPunkt` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date)))) AS `max1`,(SELECT max(cast(replace(concat('0',trim(`w`.`Wert`)),',','.') as decimal(9,2))) FROM (`laborxwert` `w` LEFT JOIN `laborxus` `u` on((`w`.`RefNr` = `u`.`RefNr`))) WHERE ((`u`.`Pat_id` = `n`.`Pat_ID`) and (`w`.`Abkü` regexp '^hba[1c]') and (cast(replace(concat('0',trim(`w`.`Wert`)),',','.') a" & _
  "s decimal(9,2)) < 21) and (`w`.`Wert` regexp '^[0-9 .,]*$') and (`u`.`Eingang` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date)))) AS `max2`,(SELECT min(cast(replace(concat('0',trim(`ln`.`Wert`)),',','.') as decimal(9,2))) FROM `laborneu` `ln` WHERE ((`ln`.`Pat_ID` = `n`.`Pat_ID`) and (`ln`.`Abkü` regexp '^hba[1c]') and (cast(replace(concat('0',trim(`ln`.`Wert`)),',','.') as decimal(9,2)) < 21) and (`ln`.`Wert` regexp '^[0-9 .,]*$') and (`ln`.`ZeitPunkt` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date)))) AS `min1`,(SELECT min(cast(replace(concat('0',trim(`w`.`Wert`)),',','.') as decimal(9,2))) FROM (`laborxwert` `w` LEFT JOIN `laborxus` `u` on((`w`.`RefNr` = `u`.`RefNr`))) WHERE ((`u`.`Pat_id` = `n`.`Pat_ID`) and (`" & _
  "w`.`Abkü` regexp '^hba[1c]') and (cast(replace(concat('0',trim(`w`.`Wert`)),',','.') as decimal(9,2)) < 21) and (`w`.`Wert` regexp '^[0-9 .,]*$') and (`u`.`Eingang` > cast((SELECT concat(year((now() - interval 200 day)),'-',((((month((now() - interval 200 day)) - 1) DIV 3) * 3) + 1),'-01')) as date)))) AS `min2` FROM `namen` `n`"
End Sub ' FüllStr16

Sub FüllStr17()
 Str(0, 17, 0) = "__fuerlmp"
 Str(1, 17, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `__fuerlmp` AS SELECT `mp`.`Pat_ID` AS `pat_id`,`mp`.`ZeitPunkt` AS `zeitpunkt`,`mp`.`MPNr` AS `mpnr` FROM `medplan` `mp` group by `mp`.`Pat_ID`,`mp`.`ZeitPunkt`,`mp`.`MPNr` ORDER BY `mp`.`Pat_ID`,`mp`.`ZeitPunkt` desc,`mp`.`MPNr` desc"
End Sub ' FüllStr17

Sub FüllStr18()
 Str(0, 18, 0) = "__kontakttage"
 Str(1, 18, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `__kontakttage` AS SELECT `e`.`Pat_ID` AS `pat_id`,`e`.`ZeitPunkt` AS `zeitpunkt` FROM `eintraege` `e` WHERE ((`e`.`ZeitPunkt` between concat(year((now() - interval 29 day)),'-',(((quarter((now() - interval 29 day)) - 1) * 3) + 1),'-01') and concat((year((now() - interval 29 day)) + (quarter((now() - interval 29 day)) DIV 4)),'-',((((quarter((now() - interval 29 day)) - 1) * 3) + 4) % 12),'-01')) and (`e`.`Art` in (_latin1'notiz',_latin1'telef',_latin1'ni',_latin1'gstel',_latin1'gs',_latin1'rz',_latin1'ep',_latin1'bga',_latin1'tk',_latin1'APK',_latin1'wr',_latin1'jl',_latin1'ga',_latin1'ih',_latin1'cr',_latin1'tst',_latin1'ke',_latin1'hz',_latin1'mh',_latin1'ag',_latin1'ph',_latin1'pq',_latin1'er',_latin1'ds',_latin1'st',_latin1'eb',_latin1'us',_latin1'sn',_latin1'vb',_latin1'mm',_latin1'rc',_lati" & _
  "n1'ik',_latin1'ks',_latin1'sb',_latin1'cb',_latin1'th',_latin1'sp',_latin1'fa',_latin1'bz',_latin1'rp',_latin1'uzu',_latin1'hypo',_latin1'colo',_latin1'coloauf',_latin1'aug',_latin1'beweg', _latin1'bewegung', _latin1'bew', _latin1'bewg', _latin1'bewe',_latin1'pros',_latin1'impf',_latin1'gyn',_latin1'caro',_latin1'beruf',_latin1'ap',_latin1'mu',_latin1'rauch',_latin1'alko',_latin1'fams',_latin1'schula',_latin1'ass',_latin1'kra',_latin1'proc',_latin1'au',_latin1'GPD',_latin1'ba',_latin1'ARCHIE2',_latin1'gewicht',_latin1'gewi',_latin1'rrvgl',_latin1'bzvgl',_latin1'fuß',_latin1'taille',_latin1'urin',_latin1'bzm',_latin1'bztp',_latin1'bks',_latin1'anal',_latin1'andm',_latin1'andm2',_latin1'usal',_latin1'usdm',_latin1'doppler',_latin1'duplex',_latin1'sono',_latin1'sd',_latin1'UKG',_latin1'Größe',_latin1'HbA1c',_latin1'hyper',_latin1'keto',_latin1'wv',_latin1'ulc',_latin1'kv',_latin1'debr',_latin1'EKG',_latin1'LZRR',_latin1'Lufu',_latin1'lactoset',_latin1'trop',_latin1't" & _
  "emp',_latin1'oGTT',_latin1'gpt',_latin1'bmi',_latin1'hüfte',_latin1'puls',_latin1'GDT',_latin1'bef'))) group by `e`.`Pat_ID`,cast(`e`.`ZeitPunkt` as date)"
End Sub ' FüllStr18

Sub FüllStr19()
 Str(0, 19, 0) = "__lfaelle"
 Str(1, 19, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `__lfaelle` AS SELECT `f`.`BhFB` AS `mbhfb`,`f`.`BhFE1` AS `bhfe1`,`f`.`Pat_ID` AS `pid` FROM `faelle` `f` ORDER BY `f`.`Pat_ID`,`f`.`BhFB`"
End Sub ' FüllStr19

Sub FüllStr20()
 Str(0, 20, 0) = "_f1"
 Str(1, 20, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_f1` AS SELECT `faelle`.`BhFB` AS `bhfb`,`faelle`.`Pat_ID` AS `pat_id` FROM `faelle` ORDER BY `faelle`.`Pat_ID`,`faelle`.`BhFB` desc,`faelle`.`SchGr`"
End Sub ' FüllStr20

Sub FüllStr21()
 Str(0, 21, 0) = "_faellenachschgr"
 Str(1, 21, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_faellenachschgr` AS SELECT `faelle`.`FID` AS `FID`,`faelle`.`Pat_ID` AS `Pat_ID`,`faelle`.`Quartal` AS `Quartal`,`faelle`.`Nachname` AS `Nachname`,`faelle`.`Vorname` AS `Vorname`,`faelle`.`lfdnr` AS `lfdnr`,`faelle`.`TMFNr` AS `TMFNr`,`faelle`.`VKNr` AS `VKNr`,`faelle`.`BhFB` AS `BhFB`,`faelle`.`BhFE1` AS `BhFE1`,`faelle`.`BhFE2` AS `BhFE2`,`faelle`.`f4202` AS `f4202`,`faelle`.`ausgst` AS `ausgst`,`faelle`.`KtrAbrB` AS `KtrAbrB`,`faelle`.`AbrAr` AS `AbrAr`,`faelle`.`lVorl` AS `lVorl`,`faelle`.`IK` AS `IK`,`faelle`.`KVKs` AS `KVKs`,`faelle`.`KVKserg` AS `KVKserg`,`faelle`.`Kasse` AS `Kasse`,`faelle`.`GebOr` AS `GebOr`,`faelle`.`AbrGb` AS `AbrGb`,`faelle`.`PersKreis` AS `PersKreis`,`faelle`.`SKtZusatz` AS `SKtZusatz`,`faelle`.`letzteRegel` AS `letzteRegel`,`faelle`.`ÜwText` AS `ÜwText`,`faelle`.`f" & _
  "4210` AS `f4210`,`faelle`.`AkfHAH` AS `AkfHAH`,`faelle`.`AkfAB0` AS `AkfAB0`,`faelle`.`AkfAK` AS `AkfAK`,`faelle`.`statNuller` AS `statNuller`,`faelle`.`ÜbwV` AS `ÜbwV`,`faelle`.`ÜbWVLANR` AS `ÜbWVLANR`,`faelle`.`ÜbWVBSNR` AS `ÜbWVBSNR`,`faelle`.`ÜbWVKVNR` AS `ÜbWVKVNR`,`faelle`.`AndÜw` AS `AndÜw`,`faelle`.`Übwr` AS `Übwr`,`faelle`.`ÜbwLANR` AS `ÜbwLANR`,`faelle`.`ÜWZiel` AS `ÜWZiel`,`faelle`.`ÜWNNr` AS `ÜWNNr`,`faelle`.`ÜWNaN` AS `ÜWNaN`,`faelle`.`ÜWTit` AS `ÜWTit`,`faelle`.`ÜWVor` AS `ÜWVor`,`faelle`.`ÜWVsw` AS `ÜWVsw`,`faelle`.`üwvid` AS `üwvid`,`faelle`.`Auftrag` AS `Auftrag`,`faelle`.`Verdacht` AS `Verdacht`,`faelle`.`Befund` AS `Befund`,`faelle`.`statKlasse` AS `statKlasse`,`faelle`.`f4237` AS `f4237`,`faelle`.`statBehTage` AS `statBehTage`,`faelle`.`SchGr` AS `SchGr`,`faelle`.`Weiterbeh` AS `Weiterbeh`,`faelle`.`PGeb` AS `PGeb`,`faelle`.`PGebErg` AS `PGebErg`,`faelle`.`Mahnfrist` " & _
  "AS `Mahnfrist`,`faelle`.`GOÄKatNr` AS `GOÄKatNr`,`faelle`.`GOÄKatName` AS `GOÄKatName`,`faelle`.`abrArzt` AS `abrArzt`,`faelle`.`privVers` AS `privVers`,`faelle`.`AdNam` AS `AdNam`,`faelle`.`AdStr` AS `AdStr`,`faelle`.`AdPlz` AS `AdPlz`,`faelle`.`AdOrt` AS `AdOrt`,`faelle`.`BhFE` AS `BhFE`,`faelle`.`s8000` AS `s8000`,`faelle`.`s8100` AS `s8100`,`faelle`.`AktZeit` AS `AktZeit`,`faelle`.`Fanf` AS `Fanf`,`faelle`.`altQuart` AS `altQuart`,`faelle`.`QAnf` AS `QAnf`,`faelle`.`QEnd` AS `QEnd`,`faelle`.`QS` AS `QS`,`faelle`.`QT` AS `QT`,`faelle`.`StByte` AS `StByte`,`faelle`.`absPos` AS `absPos`,`faelle`.`LANRid` AS `LANRid`,`faelle`.`f4108` AS `f4108`,`faelle`.`BGFallNr` AS `BGFallNr`,`faelle`.`lGewicht` AS `lGewicht` FROM `faelle` ORDER BY `faelle`.`SchGr`"
End Sub ' FüllStr21

Sub FüllStr22()
 Str(0, 22, 0) = "_fuerlmp"
 Str(1, 22, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_fuerlmp` AS SELECT `__fuerlmp`.`pat_id` AS `pat_id`,`__fuerlmp`.`mpnr` AS `mpnr` FROM `__fuerlmp` group by `__fuerlmp`.`pat_id`"
End Sub ' FüllStr22

Sub FüllStr23()
 Str(0, 23, 0) = "_kontaktzahl"
 Str(1, 23, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_kontaktzahl` AS SELECT count(0) AS `ct`,`__kontakttage`.`pat_id` AS `pat_id` FROM `__kontakttage` group by `__kontakttage`.`pat_id`"
End Sub ' FüllStr23

Sub FüllStr24()
 Str(0, 24, 0) = "_lfaelle"
 Str(1, 24, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_lfaelle` AS SELECT `__lfaelle`.`mbhfb` AS `mbhfb`,`__lfaelle`.`bhfe1` AS `bhfe1`,`__lfaelle`.`pid` AS `pid` FROM `__lfaelle` group by `__lfaelle`.`pid` desc"
End Sub ' FüllStr24

Sub FüllStr25()
 Str(0, 25, 0) = "_maxGluc"
 Str(1, 25, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_maxGluc` AS SELECT `n`.`Pat_ID` AS `pat_id`,(SELECT max(cast(replace(concat('0',trim(`ln`.`Wert`)),',','.') as decimal(9,2))) FROM `laborneu` `ln` WHERE ((`ln`.`Pat_ID` = `n`.`Pat_ID`) and (`ln`.`Abkü` regexp '^glu') and (`ln`.`Einheit` = 'mg/dl'))) AS `w1`,(SELECT max(cast(replace(concat('0',trim(`w`.`Wert`)),',','.') as decimal(9,2))) FROM (`laborxwert` `w` LEFT JOIN `laborxus` `u` on((`w`.`RefNr` = `u`.`RefNr`))) WHERE ((`u`.`Pat_id` = `n`.`Pat_ID`) and (`w`.`Abkü` regexp '^glu') and (`w`.`Einheit` = 'mg/dl'))) AS `w2` FROM `namen` `n`"
End Sub ' FüllStr25

Sub FüllStr26()
 Str(0, 26, 0) = "_maxHbA1c"
 Str(1, 26, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_maxHbA1c` AS SELECT `n`.`Pat_ID` AS `pat_id`,(SELECT max(cast(replace(concat('0',trim(`ln`.`Wert`)),',','.') as decimal(9,2))) FROM `laborneu` `ln` WHERE ((`ln`.`Pat_ID` = `n`.`Pat_ID`) and (`ln`.`Abkü` regexp '^hba[1c]') and (cast(replace(concat('0',trim(`ln`.`Wert`)),',','.') as decimal(9,2)) < 21) and (`ln`.`Wert` regexp '^[0-9 .,]*$'))) AS `w1`,(SELECT max(cast(replace(concat('0',trim(`w`.`Wert`)),',','.') as decimal(9,2))) FROM (`laborxwert` `w` LEFT JOIN `laborxus` `u` on((`w`.`RefNr` = `u`.`RefNr`))) WHERE ((`u`.`Pat_id` = `n`.`Pat_ID`) and (`w`.`Abkü` regexp '^hba[1c]') and (cast(replace(concat('0',trim(`w`.`Wert`)),',','.') as decimal(9,2)) < 21) and (`w`.`Wert` regexp '^[0-9 .,]*$'))) AS `w2` FROM `namen` `n`"
End Sub ' FüllStr26

Sub FüllStr27()
 Str(0, 27, 0) = "_qsumme"
 Str(1, 27, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_qsumme` AS SELECT (count(0) * if((`e`.`Euro` = 18.75 or e.euro = 19.05 Or euro = 14.25) and left(e.ziffer,1)='9',75,`e`.`Euro`)) AS `wert`,count(0) AS `zahl`,`e`.`Euro` AS `euro`,`f`.`Quartal` AS `quartal` FROM ((`faelle` `f` LEFT JOIN `leistungen` `l` on((`f`.`FID` = `l`.`FID`))) LEFT JOIN `EBM2010` `e` on(((`l`.`Leistung` = `e`.`Ziffer`) and (`e`.`Leistungstext` <> '')))) WHERE ((`f`.`SchGr` <> 90) and (`l`.`Leistung` <> '00000') and (`e`.`Euro` is not null)) group by `l`.`Leistung`,`f`.`Quartal`"
End Sub ' FüllStr27

Sub FüllStr28()
 Str(0, 28, 0) = "aktf"
 Str(1, 28, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktf` AS SELECT `faelle`.`Pat_ID` AS `pat_id`,`faelle`.`FID` AS `fid`,`faelle`.`SchGr` AS `schgr`,`faelle`.`VKNr` AS `vknr` FROM `faelle` WHERE ((`faelle`.`SchGr` <> '90') and (`faelle`.`GOÄKatNr` not in ('40','41')) and (`faelle`.`Nachname` <> (_latin1'Bereitschaftsdienst' collate latin1_german2_ci)) and (`faelle`.`Quartal` = (SELECT concat((((month((now() - interval 29 day)) - 1) DIV 3) + 1),(year((now() - interval 29 day)) collate latin1_german2_ci)) AS `lq`))) ORDER BY `faelle`.`Pat_ID`,`faelle`.`FID` desc,`faelle`.`SchGr`"
End Sub ' FüllStr28

Sub FüllStr29()
 Str(0, 29, 0) = "aktfaelle"
 Str(1, 29, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktfaelle` AS SELECT `f`.`Pat_ID` AS `pid`,`n`.`Notiz` AS `notiz`,`stru`.`Leistung` AS `stru`,`chron`.`Leistung` AS `chron`,`kt`.`ct` AS `kt`,`ebm`.`Leistung` AS `verspau`,`d`.`ICD` AS `icd`,`f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`letzteRegel` AS `letzteRegel`,`f`" & _
  ".`ÜwText` AS `ÜwText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`ÜbwV` AS `ÜbwV`,`f`.`ÜbWVLANR` AS `ÜbWVLANR`,`f`.`ÜbWVBSNR` AS `ÜbWVBSNR`,`f`.`ÜbWVKVNR` AS `ÜbWVKVNR`,`f`.`AndÜw` AS `AndÜw`,`f`.`Übwr` AS `Übwr`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`Auftrag` AS `Auftrag`,`f`.`Verdacht` AS `Verdacht`,`f`.`Befund` AS `Befund`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers" & _
  "` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8100` AS `s8100`,`f`.`AktZeit` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos`,`f`.`LANRid` AS `LANRid`,`f`.`f4108` AS `f4108`,`f`.`BGFallNr` AS `BGFallNr`,`f`.`lGewicht` AS `lGewicht`,`k`.`ID` AS `id`,`k`.`VK` AS `vk`,`k`.`Name` AS `kname`,`k`.`Kateg` AS `kateg`,`k`.`AnzahlIK` AS `anzahlik`,`k`.`AnzahlKTUG` AS `anzahlktug`,`k`.`GültigVon` AS `gültigvon`,`k`.`GültigBis` AS `gültigbis`,`k`.`GO` AS `go`,`k`.`Kurzname` AS `kurzname` FROM (((((((`faelle` `f` LEFT JOIN `kassenliste` `k` on(((`f`.`VKNr` = `k`.`VK`) and (`f`.`IK` = `k`.`IK`)))) LEFT JOIN `diagnosen` `d` on(((`f`.`Pat_ID` = `d`.`Pat_id" & _
  "`) and (`d`.`ICD` regexp '^E1[0-4].|^O24.') and (`d`.`DiagSicherheit` <> 'A') and ((`d`.`obDauer` <> 0) or (`d`.`FID` = `f`.`FID`))))) LEFT JOIN `leistungen` `ebm` on(((`f`.`FID` = `ebm`.`FID`) and ((`ebm`.`Leistung` like (_latin1'031%' collate latin1_german2_ci)) or (`ebm`.`Leistung` like (_latin1'01210' collate latin1_german2_ci)))))) LEFT JOIN `leistungen` `chron` on(((`f`.`FID` = `chron`.`FID`) and (`chron`.`Leistung` = (_latin1'03212' collate latin1_german2_ci))))) LEFT JOIN `leistungen` `stru` on(((`f`.`FID` = `stru`.`FID`) and (`stru`.`Leistung` like (_latin1'973%' collate latin1_german2_ci))))) LEFT JOIN `_kontaktzahl` `kt` on((`kt`.`pat_id` = `f`.`Pat_ID`))) LEFT JOIN `namen` `n` on((`n`.`Pat_ID` = `f`.`Pat_ID`))) WHERE ((`f`.`SchGr` <> '90') and (`f`.`GOÄKatNr` not in ('40','41')) and (`n`.`Nachname` <> (_latin1'Bereitschaftsdienst' collate latin1_german2_ci)) and (`f`.`Quartal` " & _
  "= (SELECT concat((((month((now() - interval 29 day)) - 1) DIV 3) + 1),(year((now() - interval 29 day)) collate latin1_german2_ci)) AS `lq`))) group by `f`.`FID` ORDER BY `f`.`Pat_ID`,`f`.`SchGr`,`d`.`ICD`"
End Sub ' FüllStr29

Sub FüllStr30()
 Str(0, 30, 0) = "aktfaellev"
 Str(1, 30, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktfaellev` AS SELECT `aktfaelle`.`pid` AS `pid`,`aktfaelle`.`notiz` AS `notiz`,`aktfaelle`.`stru` AS `stru`,`aktfaelle`.`chron` AS `chron`,`aktfaelle`.`kt` AS `kt`,`aktfaelle`.`verspau` AS `verspau`,`aktfaelle`.`icd` AS `icd`,`aktfaelle`.`FID` AS `FID`,`aktfaelle`.`Pat_ID` AS `Pat_ID`,`aktfaelle`.`Quartal` AS `Quartal`,`aktfaelle`.`Nachname` AS `Nachname`,`aktfaelle`.`Vorname` AS `Vorname`,`aktfaelle`.`lfdnr` AS `lfdnr`,`aktfaelle`.`TMFNr` AS `TMFNr`,`aktfaelle`.`VKNr` AS `VKNr`,`aktfaelle`.`BhFB` AS `BhFB`,`aktfaelle`.`BhFE1` AS `BhFE1`,`aktfaelle`.`BhFE2` AS `BhFE2`,`aktfaelle`.`f4202` AS `f4202`,`aktfaelle`.`ausgst` AS `ausgst`,`aktfaelle`.`KtrAbrB` AS `KtrAbrB`,`aktfaelle`.`AbrAr` AS `AbrAr`,`aktfaelle`.`lVorl` AS `lVorl`,`aktfaelle`.`IK` AS `IK`,`aktfaelle`.`KVKs` AS `KVKs`,`aktfaelle`.`KVK" & _
  "serg` AS `KVKserg`,`aktfaelle`.`Kasse` AS `Kasse`,`aktfaelle`.`GebOr` AS `GebOr`,`aktfaelle`.`AbrGb` AS `AbrGb`,`aktfaelle`.`PersKreis` AS `PersKreis`,`aktfaelle`.`SKtZusatz` AS `SKtZusatz`,`aktfaelle`.`letzteRegel` AS `letzteRegel`,`aktfaelle`.`ÜwText` AS `ÜwText`,`aktfaelle`.`f4210` AS `f4210`,`aktfaelle`.`AkfHAH` AS `AkfHAH`,`aktfaelle`.`AkfAB0` AS `AkfAB0`,`aktfaelle`.`AkfAK` AS `AkfAK`,`aktfaelle`.`statNuller` AS `statNuller`,`aktfaelle`.`ÜbwV` AS `ÜbwV`,`aktfaelle`.`ÜbWVLANR` AS `ÜbWVLANR`,`aktfaelle`.`ÜbWVBSNR` AS `ÜbWVBSNR`,`aktfaelle`.`ÜbWVKVNR` AS `ÜbWVKVNR`,`aktfaelle`.`AndÜw` AS `AndÜw`,`aktfaelle`.`Übwr` AS `Übwr`,`aktfaelle`.`ÜbwLANR` AS `ÜbwLANR`,`aktfaelle`.`ÜWZiel` AS `ÜWZiel`,`aktfaelle`.`ÜWNNr` AS `ÜWNNr`,`aktfaelle`.`ÜWNaN` AS `ÜWNaN`,`aktfaelle`.`ÜWTit` AS `ÜWTit`,`aktfaelle`.`ÜWVor` AS `ÜWVor`,`aktfaelle`.`ÜWVsw` AS `ÜWVsw`,`aktfaelle`.`üwvid` AS `üwvid`,`aktfaelle`" & _
  ".`Auftrag` AS `Auftrag`,`aktfaelle`.`Verdacht` AS `Verdacht`,`aktfaelle`.`Befund` AS `Befund`,`aktfaelle`.`statKlasse` AS `statKlasse`,`aktfaelle`.`f4237` AS `f4237`,`aktfaelle`.`statBehTage` AS `statBehTage`,`aktfaelle`.`SchGr` AS `SchGr`,`aktfaelle`.`Weiterbeh` AS `Weiterbeh`,`aktfaelle`.`PGeb` AS `PGeb`,`aktfaelle`.`PGebErg` AS `PGebErg`,`aktfaelle`.`Mahnfrist` AS `Mahnfrist`,`aktfaelle`.`GOÄKatNr` AS `GOÄKatNr`,`aktfaelle`.`GOÄKatName` AS `GOÄKatName`,`aktfaelle`.`abrArzt` AS `abrArzt`,`aktfaelle`.`privVers` AS `privVers`,`aktfaelle`.`AdNam` AS `AdNam`,`aktfaelle`.`AdStr` AS `AdStr`,`aktfaelle`.`AdPlz` AS `AdPlz`,`aktfaelle`.`AdOrt` AS `AdOrt`,`aktfaelle`.`BhFE` AS `BhFE`,`aktfaelle`.`s8000` AS `s8000`,`aktfaelle`.`s8100` AS `s8100`,`aktfaelle`.`AktZeit` AS `AktZeit`,`aktfaelle`.`Fanf` AS `Fanf`,`aktfaelle`.`altQuart` AS `altQuart`,`aktfaelle`.`QAnf` AS `QAnf`,`aktfaelle`.`QEnd` AS `" & _
  "QEnd`,`aktfaelle`.`QS` AS `QS`,`aktfaelle`.`QT` AS `QT`,`aktfaelle`.`StByte` AS `StByte`,`aktfaelle`.`absPos` AS `absPos`,`aktfaelle`.`LANRid` AS `LANRid`,`aktfaelle`.`f4108` AS `f4108`,`aktfaelle`.`BGFallNr` AS `BGFallNr`,`aktfaelle`.`lGewicht` AS `lGewicht`,`aktfaelle`.`id` AS `id`,`aktfaelle`.`vk` AS `vk`,`aktfaelle`.`kname` AS `kname`,`aktfaelle`.`kateg` AS `kateg`,`aktfaelle`.`anzahlik` AS `anzahlik`,`aktfaelle`.`anzahlktug` AS `anzahlktug`,`aktfaelle`.`gültigvon` AS `gültigvon`,`aktfaelle`.`gültigbis` AS `gültigbis`,`aktfaelle`.`go` AS `go`,`aktfaelle`.`kurzname` AS `kurzname` FROM `aktfaelle` group by `aktfaelle`.`Pat_ID`"
End Sub ' FüllStr30

Sub FüllStr31()
 Str(0, 31, 0) = "aktfv"
 Str(1, 31, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktfv` AS SELECT `faelle`.`Pat_ID` AS `pat_id`,`faelle`.`FID` AS `fid`,`faelle`.`SchGr` AS `schgr`,`faelle`.`GOÄKatNr` AS `goäkatnr`,`faelle`.`VKNr` AS `vknr` FROM `faelle` WHERE ((`faelle`.`SchGr` <> '90') and (`faelle`.`GOÄKatNr` not in ('40','41')) and (`faelle`.`Nachname` <> (_latin1'Bereitschaftsdienst' collate latin1_german2_ci)) and (`faelle`.`Quartal` = (SELECT concat((((month((now() - interval 29 day)) - 1) DIV 3) + 1),(year((now() - interval 29 day)) collate latin1_german2_ci)) AS `lq`))) group by `faelle`.`Pat_ID` ORDER BY `faelle`.`Pat_ID`"
End Sub ' FüllStr31

Sub FüllStr32()
 Str(0, 32, 0) = "althae"
 Str(0, 32, 1) = "`neuAI`"
 Str(0, 32, 2) = "`DBNr`"
 Str(0, 32, 3) = "`diff`"
 Str(0, 32, 4) = "`BStelle`"
 Str(0, 32, 5) = "`Anrede`"
 Str(0, 32, 6) = "`HAName`"
 Str(0, 32, 7) = "`ort`"
 Str(0, 32, 8) = "`KVNR`"
 Str(0, 32, 9) = "`KVNu`"
 Str(0, 32, 10) = "`LANR`"
 Str(0, 32, 11) = "`Tel1`"
 Str(0, 32, 12) = "`Tel2`"
 Str(0, 32, 13) = "`Tel3`"
 Str(0, 32, 14) = "`Tel4`"
 Str(0, 32, 15) = "`Fax1`"
 Str(0, 32, 16) = "`Fax1k`"
 Str(0, 32, 17) = "`Fax2`"
 Str(0, 32, 18) = "`Fax2k`"
 Str(0, 32, 19) = "`Fax3`"
 Str(0, 32, 20) = "`Fax3k`"
 Str(0, 32, 21) = "`Email`"
 Str(0, 32, 22) = "`ZulG`"
 Str(0, 32, 23) = "`Arzttyp`"
 Str(0, 32, 24) = "`GemMit`"
 Str(0, 32, 25) = "`beme`"
 Str(0, 32, 26) = "`DMPT2`"
 Str(0, 32, 27) = "`DMPT1`"
 Str(0, 32, 28) = "`Geschlecht`"
 Str(0, 32, 29) = "`Titel`"
 Str(0, 32, 30) = "`Vorname`"
 Str(0, 32, 31) = "`Nachname`"
 Str(0, 32, 32) = "`Straße`"
 Str(0, 32, 33) = "`PLZ`"
 Str(0, 32, 34) = "`gelöscht`"
 Str(0, 32, 35) = "`seit`"
 Str(0, 32, 36) = "`bis`"
 Str(0, 32, 37) = "`AktZeit`"
 Str(0, 32, 38) = "`neuAI`"
 Str(0, 32, 39) = "`Fax1k`"
 Str(0, 32, 40) = "`KVNr`"
 Str(0, 32, 41) = "`Name`"
 Str(0, 32, 42) = "`kvnu`"
 Str(0, 32, 43) = "`dbnr`"
 ArtZ(0, 32) = 37
 ArtZ(1, 32) = 6
 Str(1, 32, 0) = "CREATE TABLE `althae` ("
 Str(1, 32, 1) = " `neuAI` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 32, 2) = " `DBNr` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 32, 3) = " `diff` int(10) DEFAULT NULL"
 Str(1, 32, 4) = " `BStelle` longtext COLLATE latin1_german2_ci"
 Str(1, 32, 5) = " `Anrede` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 6) = " `HAName` longtext COLLATE latin1_german2_ci"
 Str(1, 32, 7) = " `ort` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 8) = " `KVNR` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 9) = " `KVNu` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 10) = " `LANR` int(9) unsigned NOT NULL"
 Str(1, 32, 11) = " `Tel1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 12) = " `Tel2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 13) = " `Tel3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 14) = " `Tel4` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 15) = " `Fax1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 16) = " `Fax1k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 17) = " `Fax2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 18) = " `Fax2k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 19) = " `Fax3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 20) = " `Fax3k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 21) = " `Email` longtext COLLATE latin1_german2_ci"
 Str(1, 32, 22) = " `ZulG` longtext COLLATE latin1_german2_ci"
 Str(1, 32, 23) = " `Arzttyp` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 24) = " `GemMit` longtext COLLATE latin1_german2_ci"
 Str(1, 32, 25) = " `beme` longtext COLLATE latin1_german2_ci"
 Str(1, 32, 26) = " `DMPT2` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 32, 27) = " `DMPT1` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 32, 28) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 29) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 30) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 31) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 32) = " `Straße` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 33) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 34) = " `gelöscht` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 32, 35) = " `seit` date NOT NULL"
 Str(1, 32, 36) = " `bis` date NOT NULL"
 Str(1, 32, 37) = " `AktZeit` datetime DEFAULT NULL"
 Str(1, 32, 38) = "  PRIMARY KEY (`neuAI`)"
 Str(1, 32, 39) = "  KEY `Fax1k` (`Fax1k`)"
 Str(1, 32, 40) = "  KEY `KVNr` (`KVNR`,`Nachname`,`Vorname`)"
 Str(1, 32, 41) = "  KEY `Name` (`Nachname`,`Vorname`,`ort`)"
 Str(1, 32, 42) = "  KEY `kvnu` (`KVNu`)"
 Str(1, 32, 43) = "  KEY `dbnr` (`DBNr`,`KVNR`) USING BTREE"
 Str(1, 32, 44) = " ENGINE=InnoDB AUTO_INCREMENT=2224 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr32

Sub FüllStr33()
 Str(0, 33, 0) = "anamnesebogen"
 Str(0, 33, 1) = "`Prim`"
 Str(0, 33, 2) = "`Pat_id`"
 Str(0, 33, 3) = "`Nachname`"
 Str(0, 33, 4) = "`Vorname`"
 Str(0, 33, 5) = "`NVorsatz`"
 Str(0, 33, 6) = "`Titel`"
 Str(0, 33, 7) = "`Anrede`"
 Str(0, 33, 8) = "`GebDat`"
 Str(0, 33, 9) = "`Tkz`"
 Str(0, 33, 10) = "`Versicherungsart`"
 Str(0, 33, 11) = "`Diabetestyp`"
 Str(0, 33, 12) = "`Diabetes seit`"
 Str(0, 33, 13) = "`Tabletten seit`"
 Str(0, 33, 14) = "`Insulin seit`"
 Str(0, 33, 15) = "`Grund für Vorstellung`"
 Str(0, 33, 16) = "`Familienanamnese`"
 Str(0, 33, 17) = "`Größe`"
 Str(0, 33, 18) = "`Gewicht`"
 Str(0, 33, 19) = "`bmi`"
 Str(0, 33, 20) = "`Tendenz`"
 Str(0, 33, 21) = "`DiabetesMedikament 1`"
 Str(0, 33, 22) = "`DiabetesMedikament 1 Menge`"
 Str(0, 33, 23) = "`DiabetesMedikament 2`"
 Str(0, 33, 24) = "`DiabetesMedikament 2 Menge`"
 Str(0, 33, 25) = "`DiabetesMedikament 3`"
 Str(0, 33, 26) = "`DiabetesMedikament 3 Menge`"
 Str(0, 33, 27) = "`DiabetesMedikament 4`"
 Str(0, 33, 28) = "`DiabetesMedikament 4 Menge`"
 Str(0, 33, 29) = "`Insulinpumpe`"
 Str(0, 33, 30) = "`Insulinpumpe seit`"
 Str(0, 33, 31) = "`Insulinpumpe Marke`"
 Str(0, 33, 32) = "`Broteinheiten gesamt`"
 Str(0, 33, 33) = "`Broteinheiten früh`"
 Str(0, 33, 34) = "`Broteinheiten ZM früh`"
 Str(0, 33, 35) = "`Broteinheiten mittags`"
 Str(0, 33, 36) = "`Broteinheiten nachmittags`"
 Str(0, 33, 37) = "`Broteinheiten abends`"
 Str(0, 33, 38) = "`Broteinheiten nachts`"
 Str(0, 33, 39) = "`Essenszeit früh`"
 Str(0, 33, 40) = "`Essenszeit vormittags`"
 Str(0, 33, 41) = "`Essenszeit mittags`"
 Str(0, 33, 42) = "`Essenszeit nachmittags`"
 Str(0, 33, 43) = "`Essenszeit abends`"
 Str(0, 33, 44) = "`Essenszeit spät`"
 Str(0, 33, 45) = "`Spritz-Eß-Abstand früh`"
 Str(0, 33, 46) = "`Spritz-Eß-Abstand mittags`"
 Str(0, 33, 47) = "`Spritz-Eß-Abstand abends`"
 Str(0, 33, 48) = "`Spritzstelle früh`"
 Str(0, 33, 49) = "`Spritzstelle mittags`"
 Str(0, 33, 50) = "`Spritzstelle abends`"
 Str(0, 33, 51) = "`Spritzstelle nachts`"
 Str(0, 33, 52) = "`Jahr letzte Diabetesschulung`"
 Str(0, 33, 53) = "`Ort Schulung`"
 Str(0, 33, 54) = "`letztes HbA1c`"
 Str(0, 33, 55) = "`gemessen am`"
 Str(0, 33, 56) = "`vorherige Werte`"
 Str(0, 33, 57) = "`BZMessungen selbst`"
 Str(0, 33, 58) = "`Gerät`"
 Str(0, 33, 59) = "`BZMessungen pW`"
 Str(0, 33, 60) = "`BZMessungen pW ndE`"
 Str(0, 33, 61) = "`BZMessungen p W nachts`"
 Str(0, 33, 62) = "`Aufschreiben`"
 Str(0, 33, 63) = "`BZWerte v d Essen`"
 Str(0, 33, 64) = "`BZWerte n d Essen`"
 Str(0, 33, 65) = "`UZ Tageszeit`"
 Str(0, 33, 66) = "`Unterzucker pM`"
 Str(0, 33, 67) = "`UZ rechtzeitig`"
 Str(0, 33, 68) = "`Fremde Hilfe pa`"
 Str(0, 33, 69) = "`Bewußtlos pa`"
 Str(0, 33, 70) = "`Keto pa`"
 Str(0, 33, 71) = "`BZgr300 pM`"
 Str(0, 33, 72) = "`Bluthochdruck`"
 Str(0, 33, 73) = "`BHD seit`"
 Str(0, 33, 74) = "`BHD beh mit`"
 Str(0, 33, 75) = "`Blutdruckwerte`"
 Str(0, 33, 76) = "`BDselbst`"
 Str(0, 33, 77) = "`Schwanger`"
 Str(0, 33, 78) = "`Schwanger seit`"
 Str(0, 33, 79) = "`Augensp zuletzt`"
 Str(0, 33, 80) = "`Augensp Befund`"
 Str(0, 33, 81) = "`Netzhaut gelasert`"
 Str(0, 33, 82) = "`Sehminderung unbehebbar`"
 Str(0, 33, 83) = "`Diabet Nierenschaden`"
 Str(0, 33, 84) = "`Albumin zuletzt`"
 Str(0, 33, 85) = "`erhöht?`"
 Str(0, 33, 86) = "`Dialyse`"
 Str(0, 33, 87) = "`Dialyse seit`"
 Str(0, 33, 88) = "`andere Nierenerkrankung`"
 Str(0, 33, 89) = "`Herzkrankheit`"
 Str(0, 33, 90) = "`Angina pectoris`"
 Str(0, 33, 91) = "`Herzinfarkt`"
 Str(0, 33, 92) = "`Herzinfarkt wann`"
 Str(0, 33, 93) = "`PTCA oder Stent`"
 Str(0, 33, 94) = "`Bypass kardial`"
 Str(0, 33, 95) = "`Bypass wann`"
 Str(0, 33, 96) = "`Herzschwäche`"
 Str(0, 33, 97) = "`Herzkrankheit Beschreibung`"
 Str(0, 33, 98) = "`Hirndurchblutungsstörung`"
 Str(0, 33, 99) = "`Schlaganfall`"
 Str(0, 33, 100) = "`Beindurchblutungsstörung`"
 Str(0, 33, 101) = "`Schaufensterkrankheit`"
 Str(0, 33, 102) = "`Bypaß peripher`"
 Str(0, 33, 103) = "`Geschwür`"
 Str(0, 33, 104) = "`Amputation`"
 Str(0, 33, 105) = "`pAVK Beschreibung`"
 Str(0, 33, 106) = "`Ameisenlaufen`"
 Str(0, 33, 107) = "`Ameisen Ausmaß`"
 Str(0, 33, 108) = "`Druckstellen`"
 Str(0, 33, 109) = "`Verformungen`"
 Str(0, 33, 110) = "`Verformungen Beschreibung`"
 Str(0, 33, 111) = "`Fußpflege`"
 Str(0, 33, 112) = "`Podologie`"
 Str(0, 33, 113) = "`Einlagen`"
 Str(0, 33, 114) = "`Neue Fußkomplikationen`"
 Str(0, 33, 115) = "`Entleerungsstörungen Magen`"
 Str(0, 33, 116) = "`Entleerungsstörungen Harnblase`"
 Str(0, 33, 117) = "`Schwindel Aufstehen`"
 Str(0, 33, 118) = "`Folgeerkrankungen Haut`"
 Str(0, 33, 119) = "`Bewegungseinschränkungen`"
 Str(0, 33, 120) = "`Sexualstörung`"
 Str(0, 33, 121) = "`Sexualstörung seit`"
 Str(0, 33, 122) = "`Weitere Anamnese`"
 Str(0, 33, 123) = "`Tabak`"
 Str(0, 33, 124) = "`tabakex`"
 Str(0, 33, 125) = "`tabakbis`"
 Str(0, 33, 126) = "`tabakakt`"
 Str(0, 33, 127) = "`tabakmenge`"
 Str(0, 33, 128) = "`Alkohol`"
 Str(0, 33, 129) = "`Mitarbeiter`"
 Str(0, 33, 130) = "`Weitere Medikation`"
 Str(0, 33, 131) = "`Liphypertrophien Abdomen`"
 Str(0, 33, 132) = "`Liphypertrophien Beine`"
 Str(0, 33, 133) = "`Liphypertrophien Arme`"
 Str(0, 33, 134) = "`Beinbefund`"
 Str(0, 33, 135) = "`Hyperkeratosen`"
 Str(0, 33, 136) = "`Ulcera`"
 Str(0, 33, 137) = "`Kraft Zehenheber`"
 Str(0, 33, 138) = "`Kraft Zehenbeuger`"
 Str(0, 33, 139) = "`Kraft Knie`"
 Str(0, 33, 140) = "`ASR`"
 Str(0, 33, 141) = "`PSR`"
 Str(0, 33, 142) = "`Oberflächensensibilität`"
 Str(0, 33, 143) = "`Monofilamenttest`"
 Str(0, 33, 144) = "`Kalt-Warm`"
 Str(0, 33, 145) = "`Vibration IK`"
 Str(0, 33, 146) = "`Vibration Großzehe`"
 Str(0, 33, 147) = "`Puls Leiste`"
 Str(0, 33, 148) = "`Puls Kniekehle`"
 Str(0, 33, 149) = "`Puls Atp`"
 Str(0, 33, 150) = "`Puls Adp`"
 Str(0, 33, 151) = "`RR`"
 Str(0, 33, 152) = "`RRTurboMed`"
 Str(0, 33, 153) = "`Herz`"
 Str(0, 33, 154) = "`Lunge`"
 Str(0, 33, 155) = "`Bauch`"
 Str(0, 33, 156) = "`WS`"
 Str(0, 33, 157) = "`NL`"
 Str(0, 33, 158) = "`SD`"
 Str(0, 33, 159) = "`Carotiden`"
 Str(0, 33, 160) = "`NNH`"
 Str(0, 33, 161) = "`Zähne`"
 Str(0, 33, 162) = "`Mundhöhle`"
 Str(0, 33, 163) = "`LK`"
 Str(0, 33, 164) = "`BeinödVen`"
 Str(0, 33, 165) = "`Neuro sonst`"
 Str(0, 33, 166) = "`Weitere Befunde`"
 Str(0, 33, 167) = "`Schulung`"
 Str(0, 33, 168) = "`DMP`"
 Str(0, 33, 169) = "`DMSchulz`"
 Str(0, 33, 170) = "`DMSchL`"
 Str(0, 33, 171) = "`RRSchulz`"
 Str(0, 33, 172) = "`DMPhier`"
 Str(0, 33, 173) = "`HANr`"
 Str(0, 33, 174) = "`HANr2`"
 Str(0, 33, 175) = "`letzte Änderung`"
 Str(0, 33, 176) = "`Diagnosen`"
 Str(0, 33, 177) = "`Vorgestellt`"
 Str(0, 33, 178) = "`Versicherung`"
 Str(0, 33, 179) = "`AktZeit`"
 Str(0, 33, 180) = "`Ther1`"
 Str(0, 33, 181) = "`TherAkt`"
 Str(0, 33, 182) = "`obAn1eing`"
 Str(0, 33, 183) = "`obAn2eing`"
 Str(0, 33, 184) = "`obAnAeing`"
 Str(0, 33, 185) = "`obCheck`"
 Str(0, 33, 186) = "`obBZausgew`"
 Str(0, 33, 187) = "`obOSaufgek`"
 Str(0, 33, 188) = "`obPodAufgek`"
 Str(0, 33, 189) = "`obMBlAusgeh`"
 Str(0, 33, 190) = "`obSchulaufgek`"
 Str(0, 33, 191) = "`obDMPaufgekl`"
 Str(0, 33, 192) = "`obMedNetz`"
 Str(0, 33, 193) = "`Hausarzt`"
 Str(0, 33, 194) = "`ob`"
 Str(0, 33, 195) = "`QS`"
 Str(0, 33, 196) = "`QT`"
 Str(0, 33, 197) = "`Prim`"
 Str(0, 33, 198) = "`PrimaryKey`"
 Str(0, 33, 199) = "`Auswahl`"
 Str(0, 33, 200) = "`DMPhier`"
 Str(0, 33, 201) = "`GebDat`"
 Str(0, 33, 202) = "`jlD`"
 Str(0, 33, 203) = "`lÄnd`"
 Str(0, 33, 204) = "`Pat_ID`"
 Str(0, 33, 205) = "`Ther1`"
 Str(0, 33, 206) = "`Vorgestellt`"
 Str(0, 33, 207) = "`HausärzteAnamnesebogen_AccRel`"
 Str(0, 33, 208) = "`KassenlisteAnamnesebogen_AccRel`"
 Str(0, 33, 209) = "`Haus??rzteAnamnesebogen_AccRel`"
 Str(0, 33, 210) = "`KassenlisteAnamnesebogen_AccRel`"
 ArtZ(0, 33) = 196
 ArtZ(1, 33) = 12
 ArtZ(2, 33) = 2
 Str(1, 33, 0) = "CREATE TABLE `anamnesebogen` ("
 Str(1, 33, 1) = " `Prim` int(10) unsigned NOT NULL COMMENT 'Primärschlüssel'"
 Str(1, 33, 2) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 33, 3) = " `Nachname` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '-'"
 Str(1, 33, 4) = " `Vorname` varchar(22) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 5) = " `NVorsatz` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 6) = " `Titel` varchar(18) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 7) = " `Anrede` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 8) = " `GebDat` datetime DEFAULT NULL COMMENT ', geb.'"
 Str(1, 33, 9) = " `Tkz` tinyint(1) unsigned DEFAULT NULL COMMENT 'Tod-Kennzeichen'"
 Str(1, 33, 10) = " `Versicherungsart` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 11) = " `Diabetestyp` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetes Typ'"
 Str(1, 33, 12) = " `Diabetes seit` varchar(152) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 33, 13) = " `Tabletten seit` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Tabletten seit'"
 Str(1, 33, 14) = " `Insulin seit` varchar(149) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Insulin seit'"
 Str(1, 33, 15) = " `Grund für Vorstellung` varchar(829) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 16) = " `Familienanamnese` varchar(291) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 17) = " `Größe` double DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 18) = " `Gewicht` double DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 19) = " `bmi` decimal(5,1) DEFAULT '0.0'"
 Str(1, 33, 20) = " `Tendenz` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Tendenz'"
 Str(1, 33, 21) = " `DiabetesMedikament 1` varchar(68) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesmedikation:'"
 Str(1, 33, 22) = " `DiabetesMedikament 1 Menge` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 33, 23) = " `DiabetesMedikament 2` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 33, 24) = " `DiabetesMedikament 2 Menge` varchar(126) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 33, 25) = " `DiabetesMedikament 3` varchar(44) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 33, 26) = " `DiabetesMedikament 3 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 33, 27) = " `DiabetesMedikament 4` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 33, 28) = " `DiabetesMedikament 4 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 33, 29) = " `Insulinpumpe` tinyint(1) unsigned DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 30) = " `Insulinpumpe seit` varchar(250) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 33, 31) = " `Insulinpumpe Marke` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Marke:'"
 Str(1, 33, 32) = " `Broteinheiten gesamt` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Broteinheiten:gesamt'"
 Str(1, 33, 33) = " `Broteinheiten früh` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, früh'"
 Str(1, 33, 34) = " `Broteinheiten ZM früh` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zwischenmahlzeit vormittags'"
 Str(1, 33, 35) = " `Broteinheiten mittags` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 33, 36) = " `Broteinheiten nachmittags` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 33, 37) = " `Broteinheiten abends` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 33, 38) = " `Broteinheiten nachts` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 33, 39) = " `Essenszeit früh` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Essenszeiten:früh'"
 Str(1, 33, 40) = " `Essenszeit vormittags` varchar(34) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vormittags'"
 Str(1, 33, 41) = " `Essenszeit mittags` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 33, 42) = " `Essenszeit nachmittags` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 33, 43) = " `Essenszeit abends` varchar(18) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 33, 44) = " `Essenszeit spät` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, spät'"
 Str(1, 33, 45) = " `Spritz-Eß-Abstand früh` varchar(114) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritz-Eß-Abstand:früh'"
 Str(1, 33, 46) = " `Spritz-Eß-Abstand mittags` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 33, 47) = " `Spritz-Eß-Abstand abends` varchar(62) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 33, 48) = " `Spritzstelle früh` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritzstellen:früh'"
 Str(1, 33, 49) = " `Spritzstelle mittags` varchar(22) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 33, 50) = " `Spritzstelle abends` varchar(139) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 33, 51) = " `Spritzstelle nachts` varchar(141) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 33, 52) = " `Jahr letzte Diabetesschulung` varchar(129) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesschulung:'"
 Str(1, 33, 53) = " `Ort Schulung` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<in'"
 Str(1, 33, 54) = " `letztes HbA1c` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letztes HbA1c:'"
 Str(1, 33, 55) = " `gemessen am` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, gemessen'"
 Str(1, 33, 56) = " `vorherige Werte` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vorher:'"
 Str(1, 33, 57) = " `BZMessungen selbst` varchar(68) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckermessung:Selbstmessung?'"
 Str(1, 33, 58) = " `Gerät` varchar(103) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<:'"
 Str(1, 33, 59) = " `BZMessungen pW` varchar(149) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl d.Messungen pro Woche:'"
 Str(1, 33, 60) = " `BZMessungen pW ndE` varchar(62) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, davon nach dem Essen:'"
 Str(1, 33, 61) = " `BZMessungen p W nachts` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts:'"
 Str(1, 33, 62) = " `Aufschreiben` varchar(93) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Dokumentation:'"
 Str(1, 33, 63) = " `BZWerte v d Essen` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckerwerte vor dem Essen:'"
 Str(1, 33, 64) = " `BZWerte n d Essen` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nach dem Essen:'"
 Str(1, 33, 65) = " `UZ Tageszeit` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Unterzucker:Bevorzugte Tages-/Uhrzeit'"
 Str(1, 33, 66) = " `Unterzucker pM` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl der schweren (<50 mg/dl) pro Monat:'"
 Str(1, 33, 67) = " `UZ rechtzeitig` varchar(101) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, rechtzeitig bemerkt:'"
 Str(1, 33, 68) = " `Fremde Hilfe pa` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, fremde Hilfe deshalb nötig:'"
 Str(1, 33, 69) = " `Bewußtlos pa` varchar(81) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, bewußtlos deshalb:'"
 Str(1, 33, 70) = " `Keto pa` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Zahl der Ketoazidosen pro Jahr:'"
 Str(1, 33, 71) = " `BZgr300 pM` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Zahl der Blutzucker > 300 mg/dl pro Monat:'"
 Str(1, 33, 72) = " `Bluthochdruck` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Bluthochdruck:'"
 Str(1, 33, 73) = " `BHD seit` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit:'"
 Str(1, 33, 74) = " `BHD beh mit` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, behandelt mit:'"
 Str(1, 33, 75) = " `Blutdruckwerte` varchar(186) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckwerte:'"
 Str(1, 33, 76) = " `BDselbst` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckselbstmessung:'"
 Str(1, 33, 77) = " `Schwanger` varchar(43) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Aktuelle Schwangerschaft:'"
 Str(1, 33, 78) = " `Schwanger seit` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, seit:'"
 Str(1, 33, 79) = " `Augensp zuletzt` varchar(129) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Augenspiegelung:'"
 Str(1, 33, 80) = " `Augensp Befund` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 33, 81) = " `Netzhaut gelasert` varchar(142) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Netzhaut schon gelasert:'"
 Str(1, 33, 82) = " `Sehminderung unbehebbar` varchar(157) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', mit Brille nicht behebbare Sehminderung:'"
 Str(1, 33, 83) = " `Diabet Nierenschaden` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetischer Nierenschaden:'"
 Str(1, 33, 84) = " `Albumin zuletzt` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', letztes Albumin:'"
 Str(1, 33, 85) = " `erhöht?` varchar(52) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 33, 86) = " `Dialyse` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 87) = " `Dialyse seit` varchar(36) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 33, 88) = " `andere Nierenerkrankung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', andere Nierenerkrankung:'"
 Str(1, 33, 89) = " `Herzkrankheit` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Herzkrankheit:'"
 Str(1, 33, 90) = " `Angina pectoris` varchar(134) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 91) = " `Herzinfarkt` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 92) = " `Herzinfarkt wann` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 33, 93) = " `PTCA oder Stent` varchar(94) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 94) = " `Bypass kardial` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 95) = " `Bypass wann` varchar(388) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 33, 96) = " `Herzschwäche` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 97) = " `Herzkrankheit Beschreibung` varchar(213) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung:'"
 Str(1, 33, 98) = " `Hirndurchblutungsstörung` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 99) = " `Schlaganfall` varchar(291) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 100) = " `Beindurchblutungsstörung` varchar(136) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 101) = " `Schaufensterkrankheit` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 102) = " `Bypaß peripher` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 103) = " `Geschwür` varchar(174) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 104) = " `Amputation` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 105) = " `pAVK Beschreibung` varchar(97) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung der Beinbeschwerden:'"
 Str(1, 33, 106) = " `Ameisenlaufen` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 107) = " `Ameisen Ausmaß` varchar(123) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Ausmaß:'"
 Str(1, 33, 108) = " `Druckstellen` varchar(173) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 109) = " `Verformungen` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 110) = " `Verformungen Beschreibung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Beschreibung:'"
 Str(1, 33, 111) = " `Fußpflege` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 112) = " `Podologie` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 113) = " `Einlagen` varchar(91) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', diabetesgerechte orthopädische Einlagen/Schuhe:'"
 Str(1, 33, 114) = " `Neue Fußkomplikationen` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Neue Fußkomplikationen in den letzten 12 Monaten:'"
 Str(1, 33, 115) = " `Entleerungsstörungen Magen` varchar(130) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 116) = " `Entleerungsstörungen Harnblase` varchar(157) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 117) = " `Schwindel Aufstehen` varchar(121) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 118) = " `Folgeerkrankungen Haut` varchar(197) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 119) = " `Bewegungseinschränkungen` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 120) = " `Sexualstörung` varchar(76) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 121) = " `Sexualstörung seit` varchar(149) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 33, 122) = " `Weitere Anamnese` varchar(989) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 123) = " `Tabak` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Tabak:'"
 Str(1, 33, 124) = " `tabakex` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, früher:'"
 Str(1, 33, 125) = " `tabakbis` varchar(119) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, bis:'"
 Str(1, 33, 126) = " `tabakakt` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, aktuell:'"
 Str(1, 33, 127) = " `tabakmenge` varchar(118) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Menge:'"
 Str(1, 33, 128) = " `Alkohol` varchar(3302) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Alkohol pro Woche:'"
 Str(1, 33, 129) = " `Mitarbeiter` varchar(17) COLLATE latin1_german2_ci DEFAULT '' COMMENT '<, Mitarbeiter:'"
 Str(1, 33, 130) = " `Weitere Medikation` varchar(298) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 131) = " `Liphypertrophien Abdomen` varchar(320) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Liphypertrophien:Abdomen'"
 Str(1, 33, 132) = " `Liphypertrophien Beine` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Beine:'"
 Str(1, 33, 133) = " `Liphypertrophien Arme` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Arme:'"
 Str(1, 33, 134) = " `Beinbefund` varchar(341) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 135) = " `Hyperkeratosen` varchar(207) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 136) = " `Ulcera` varchar(103) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 137) = " `Kraft Zehenheber` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Kraft:Zehenheber'"
 Str(1, 33, 138) = " `Kraft Zehenbeuger` varchar(136) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zehenbeuger:'"
 Str(1, 33, 139) = " `Kraft Knie` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Knie:'"
 Str(1, 33, 140) = " `ASR` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 141) = " `PSR` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 142) = " `Oberflächensensibilität` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 143) = " `Monofilamenttest` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 144) = " `Kalt-Warm` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Kalt-Warm-Diskrimination:'"
 Str(1, 33, 145) = " `Vibration IK` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Vibrationsempfinden Innenknöchel:'"
 Str(1, 33, 146) = " `Vibration Großzehe` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Großzehe:'"
 Str(1, 33, 147) = " `Puls Leiste` varchar(33) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Pulse:Leiste'"
 Str(1, 33, 148) = " `Puls Kniekehle` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Kniekehle:'"
 Str(1, 33, 149) = " `Puls Atp` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Innenknöchel:'"
 Str(1, 33, 150) = " `Puls Adp` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Fußrücken:'"
 Str(1, 33, 151) = " `RR` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruck:'"
 Str(1, 33, 152) = " `RRTurboMed` varchar(1362) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 153) = " `Herz` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 33, 154) = " `Lunge` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 155) = " `Bauch` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Abdomen:'"
 Str(1, 33, 156) = " `WS` varchar(56) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Wirbelsäule:'"
 Str(1, 33, 157) = " `NL` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nierenlager:'"
 Str(1, 33, 158) = " `SD` varchar(88) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Schilddrüse:'"
 Str(1, 33, 159) = " `Carotiden` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Halsschlagadern:'"
 Str(1, 33, 160) = " `NNH` varchar(56) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nasennebenhöhlen:'"
 Str(1, 33, 161) = " `Zähne` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 162) = " `Mundhöhle` varchar(64) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 33, 163) = " `LK` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Lymphknoten:'"
 Str(1, 33, 164) = " `BeinödVen` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beinödeme/ Venenkrankheiten:'"
 Str(1, 33, 165) = " `Neuro sonst` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Sonstige neurologische Befunde:'"
 Str(1, 33, 166) = " `Weitere Befunde` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', weitere Befunde:'"
 Str(1, 33, 167) = " `Schulung` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Schulungsbedarf'"
 Str(1, 33, 168) = " `DMP` varchar(85) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Pat. bei HA im DMP'"
 Str(1, 33, 169) = " `DMSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der DMP-Schulungen hier'"
 Str(1, 33, 170) = " `DMSchL` smallint(6) DEFAULT NULL COMMENT 'Zahl der abgerechneten DMP-Schulungen hier'"
 Str(1, 33, 171) = " `RRSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der Hypertonie-Schulungen hier'"
 Str(1, 33, 172) = " `DMPhier` datetime DEFAULT NULL COMMENT 'ob Pat hier im DMP'"
 Str(1, 33, 173) = " `HANr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 33, 174) = " `HANr2` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 33, 175) = " `letzte Änderung` datetime DEFAULT NULL COMMENT 'Datum der letzten Änderung'"
 Str(1, 33, 176) = " `Diagnosen` varchar(1071) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 177) = " `Vorgestellt` datetime DEFAULT NULL COMMENT 'Erstvorstellung'"
 Str(1, 33, 178) = " `Versicherung` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 179) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 33, 180) = " `Ther1` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, Komb, ICT, CSII'"
 Str(1, 33, 181) = " `TherAkt` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, Komb, ICT, CSII'"
 Str(1, 33, 182) = " `obAn1eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 1 eingegeben wurde'"
 Str(1, 33, 183) = " `obAn2eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 2 eingegeben wurde'"
 Str(1, 33, 184) = " `obAnAeing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt allgemein eingegeben wurde'"
 Str(1, 33, 185) = " `obCheck` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Checkliste vorliegt'"
 Str(1, 33, 186) = " `obBZausgew` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Blutzuckergerät ausgewechselt'"
 Str(1, 33, 187) = " `obOSaufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über orthopäd Schuhmacher aufgeklärt'"
 Str(1, 33, 188) = " `obPodAufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 33, 189) = " `obMBlAusgeh` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 33, 190) = " `obSchulaufgek` varchar(14) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 33, 191) = " `obDMPaufgekl` varchar(17) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 33, 192) = " `obMedNetz` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob von Med. Netz geschickt'"
 Str(1, 33, 193) = " `Hausarzt` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Hausarzt laut Anamnesebogen'"
 Str(1, 33, 194) = " `ob` tinyint(1) unsigned DEFAULT NULL COMMENT 'für verschiedene Aktionen'"
 Str(1, 33, 195) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 33, 196) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 33, 197) = "  PRIMARY KEY (`Prim`)"
 Str(1, 33, 198) = "  UNIQUE KEY `PrimaryKey` (`Prim`)"
 Str(1, 33, 199) = "  KEY `Auswahl` (`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 33, 200) = "  KEY `DMPhier` (`DMPhier`)"
 Str(1, 33, 201) = "  KEY `GebDat` (`GebDat`)"
 Str(1, 33, 202) = "  KEY `jlD` (`Jahr letzte Diabetesschulung`,`GebDat`)"
 Str(1, 33, 203) = "  KEY `lÄnd` (`letzte Änderung`)"
 Str(1, 33, 204) = "  KEY `Pat_ID` (`Pat_id`)"
 Str(1, 33, 205) = "  KEY `Ther1` (`Ther1`)"
 Str(1, 33, 206) = "  KEY `Vorgestellt` (`Vorgestellt`,`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 33, 207) = "  KEY `HausärzteAnamnesebogen_AccRel` (`HANr`)"
 Str(1, 33, 208) = "  KEY `KassenlisteAnamnesebogen_AccRel` (`Versicherung`)"
 Str(1, 33, 209) = "  CONSTRAINT `Haus??rzteAnamnesebogen_AccRel` FOREIGN KEY (`HANr`) REFERENCES `hausaerztealt` (`KVNr`)"
 Str(1, 33, 210) = "  CONSTRAINT `KassenlisteAnamnesebogen_AccRel` FOREIGN KEY (`Versicherung`) REFERENCES `kassenliste` (`VK`)"
 Str(1, 33, 211) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr33

Sub FüllStr34()
 Str(0, 34, 0) = "anbogalt"
 Str(0, 34, 1) = "`Prim`"
 Str(0, 34, 2) = "`Pat_id`"
 Str(0, 34, 3) = "`Nachname`"
 Str(0, 34, 4) = "`Vorname`"
 Str(0, 34, 5) = "`NVorsatz`"
 Str(0, 34, 6) = "`Titel`"
 Str(0, 34, 7) = "`Anrede`"
 Str(0, 34, 8) = "`GebDat`"
 Str(0, 34, 9) = "`Tkz`"
 Str(0, 34, 10) = "`Versicherungsart`"
 Str(0, 34, 11) = "`Diabetestyp`"
 Str(0, 34, 12) = "`Diabetes seit`"
 Str(0, 34, 13) = "`Tabletten seit`"
 Str(0, 34, 14) = "`Insulin seit`"
 Str(0, 34, 15) = "`Grund für Vorstellung`"
 Str(0, 34, 16) = "`Familienanamnese`"
 Str(0, 34, 17) = "`Größe`"
 Str(0, 34, 18) = "`Gewicht`"
 Str(0, 34, 19) = "`Tendenz`"
 Str(0, 34, 20) = "`DiabetesMedikament 1`"
 Str(0, 34, 21) = "`DiabetesMedikament 1 Menge`"
 Str(0, 34, 22) = "`DiabetesMedikament 2`"
 Str(0, 34, 23) = "`DiabetesMedikament 2 Menge`"
 Str(0, 34, 24) = "`DiabetesMedikament 3`"
 Str(0, 34, 25) = "`DiabetesMedikament 3 Menge`"
 Str(0, 34, 26) = "`DiabetesMedikament 4`"
 Str(0, 34, 27) = "`DiabetesMedikament 4 Menge`"
 Str(0, 34, 28) = "`Insulinpumpe`"
 Str(0, 34, 29) = "`Insulinpumpe seit`"
 Str(0, 34, 30) = "`Insulinpumpe Marke`"
 Str(0, 34, 31) = "`Broteinheiten gesamt`"
 Str(0, 34, 32) = "`Broteinheiten früh`"
 Str(0, 34, 33) = "`Broteinheiten ZM früh`"
 Str(0, 34, 34) = "`Broteinheiten mittags`"
 Str(0, 34, 35) = "`Broteinheiten nachmittags`"
 Str(0, 34, 36) = "`Broteinheiten abends`"
 Str(0, 34, 37) = "`Broteinheiten nachts`"
 Str(0, 34, 38) = "`Essenszeit früh`"
 Str(0, 34, 39) = "`Essenszeit vormittags`"
 Str(0, 34, 40) = "`Essenszeit mittags`"
 Str(0, 34, 41) = "`Essenszeit nachmittags`"
 Str(0, 34, 42) = "`Essenszeit abends`"
 Str(0, 34, 43) = "`Essenszeit spät`"
 Str(0, 34, 44) = "`Spritz-Eß-Abstand früh`"
 Str(0, 34, 45) = "`Spritz-Eß-Abstand mittags`"
 Str(0, 34, 46) = "`Spritz-Eß-Abstand abends`"
 Str(0, 34, 47) = "`Spritzstelle früh`"
 Str(0, 34, 48) = "`Spritzstelle mittags`"
 Str(0, 34, 49) = "`Spritzstelle abends`"
 Str(0, 34, 50) = "`Spritzstelle nachts`"
 Str(0, 34, 51) = "`Jahr letzte Diabetesschulung`"
 Str(0, 34, 52) = "`Ort Schulung`"
 Str(0, 34, 53) = "`letztes HbA1c`"
 Str(0, 34, 54) = "`gemessen am`"
 Str(0, 34, 55) = "`vorherige Werte`"
 Str(0, 34, 56) = "`BZMessungen selbst`"
 Str(0, 34, 57) = "`Gerät`"
 Str(0, 34, 58) = "`BZMessungen pW`"
 Str(0, 34, 59) = "`BZMessungen pW ndE`"
 Str(0, 34, 60) = "`BZMessungen p W nachts`"
 Str(0, 34, 61) = "`Aufschreiben`"
 Str(0, 34, 62) = "`BZWerte v d Essen`"
 Str(0, 34, 63) = "`BZWerte n d Essen`"
 Str(0, 34, 64) = "`UZ Tageszeit`"
 Str(0, 34, 65) = "`Unterzucker pM`"
 Str(0, 34, 66) = "`UZ rechtzeitig`"
 Str(0, 34, 67) = "`Fremde Hilfe pa`"
 Str(0, 34, 68) = "`Bewußtlos pa`"
 Str(0, 34, 69) = "`Keto pa`"
 Str(0, 34, 70) = "`BZgr300 pM`"
 Str(0, 34, 71) = "`Bluthochdruck`"
 Str(0, 34, 72) = "`BHD seit`"
 Str(0, 34, 73) = "`BHD beh mit`"
 Str(0, 34, 74) = "`Blutdruckwerte`"
 Str(0, 34, 75) = "`BDselbst`"
 Str(0, 34, 76) = "`Schwanger`"
 Str(0, 34, 77) = "`Schwanger seit`"
 Str(0, 34, 78) = "`Augensp zuletzt`"
 Str(0, 34, 79) = "`Augensp Befund`"
 Str(0, 34, 80) = "`Netzhaut gelasert`"
 Str(0, 34, 81) = "`Sehminderung unbehebbar`"
 Str(0, 34, 82) = "`Diabet Nierenschaden`"
 Str(0, 34, 83) = "`Albumin zuletzt`"
 Str(0, 34, 84) = "`erhöht?`"
 Str(0, 34, 85) = "`Dialyse`"
 Str(0, 34, 86) = "`Dialyse seit`"
 Str(0, 34, 87) = "`andere Nierenerkrankung`"
 Str(0, 34, 88) = "`Herzkrankheit`"
 Str(0, 34, 89) = "`Angina pectoris`"
 Str(0, 34, 90) = "`Herzinfarkt`"
 Str(0, 34, 91) = "`Herzinfarkt wann`"
 Str(0, 34, 92) = "`PTCA oder Stent`"
 Str(0, 34, 93) = "`Bypass kardial`"
 Str(0, 34, 94) = "`Bypass wann`"
 Str(0, 34, 95) = "`Herzschwäche`"
 Str(0, 34, 96) = "`Herzkrankheit Beschreibung`"
 Str(0, 34, 97) = "`Hirndurchblutungsstörung`"
 Str(0, 34, 98) = "`Schlaganfall`"
 Str(0, 34, 99) = "`Beindurchblutungsstörung`"
 Str(0, 34, 100) = "`Schaufensterkrankheit`"
 Str(0, 34, 101) = "`Bypaß peripher`"
 Str(0, 34, 102) = "`Geschwür`"
 Str(0, 34, 103) = "`Amputation`"
 Str(0, 34, 104) = "`pAVK Beschreibung`"
 Str(0, 34, 105) = "`Ameisenlaufen`"
 Str(0, 34, 106) = "`Ameisen Ausmaß`"
 Str(0, 34, 107) = "`Druckstellen`"
 Str(0, 34, 108) = "`Verformungen`"
 Str(0, 34, 109) = "`Verformungen Beschreibung`"
 Str(0, 34, 110) = "`Fußpflege`"
 Str(0, 34, 111) = "`Podologie`"
 Str(0, 34, 112) = "`Einlagen`"
 Str(0, 34, 113) = "`Neue Fußkomplikationen`"
 Str(0, 34, 114) = "`Entleerungsstörungen Magen`"
 Str(0, 34, 115) = "`Entleerungsstörungen Harnblase`"
 Str(0, 34, 116) = "`Schwindel Aufstehen`"
 Str(0, 34, 117) = "`Folgeerkrankungen Haut`"
 Str(0, 34, 118) = "`Bewegungseinschränkungen`"
 Str(0, 34, 119) = "`Sexualstörung`"
 Str(0, 34, 120) = "`Sexualstörung seit`"
 Str(0, 34, 121) = "`Weitere Anamnese`"
 Str(0, 34, 122) = "`Alkohol`"
 Str(0, 34, 123) = "`Tabak`"
 Str(0, 34, 124) = "`tabakex`"
 Str(0, 34, 125) = "`tabakbis`"
 Str(0, 34, 126) = "`tabakakt`"
 Str(0, 34, 127) = "`tabakmenge`"
 Str(0, 34, 128) = "`Weitere Medikation`"
 Str(0, 34, 129) = "`Liphypertrophien Abdomen`"
 Str(0, 34, 130) = "`Liphypertrophien Beine`"
 Str(0, 34, 131) = "`Liphypertrophien Arme`"
 Str(0, 34, 132) = "`Beinbefund`"
 Str(0, 34, 133) = "`Hyperkeratosen`"
 Str(0, 34, 134) = "`Ulcera`"
 Str(0, 34, 135) = "`Kraft Zehenheber`"
 Str(0, 34, 136) = "`Kraft Zehenbeuger`"
 Str(0, 34, 137) = "`Kraft Knie`"
 Str(0, 34, 138) = "`ASR`"
 Str(0, 34, 139) = "`PSR`"
 Str(0, 34, 140) = "`Oberflächensensibilität`"
 Str(0, 34, 141) = "`Monofilamenttest`"
 Str(0, 34, 142) = "`Kalt-Warm`"
 Str(0, 34, 143) = "`Vibration IK`"
 Str(0, 34, 144) = "`Vibration Großzehe`"
 Str(0, 34, 145) = "`Puls Leiste`"
 Str(0, 34, 146) = "`Puls Kniekehle`"
 Str(0, 34, 147) = "`Puls Atp`"
 Str(0, 34, 148) = "`Puls Adp`"
 Str(0, 34, 149) = "`RR`"
 Str(0, 34, 150) = "`RRTurboMed`"
 Str(0, 34, 151) = "`Herz`"
 Str(0, 34, 152) = "`Lunge`"
 Str(0, 34, 153) = "`Bauch`"
 Str(0, 34, 154) = "`WS`"
 Str(0, 34, 155) = "`NL`"
 Str(0, 34, 156) = "`SD`"
 Str(0, 34, 157) = "`Carotiden`"
 Str(0, 34, 158) = "`NNH`"
 Str(0, 34, 159) = "`Zähne`"
 Str(0, 34, 160) = "`Mundhöhle`"
 Str(0, 34, 161) = "`LK`"
 Str(0, 34, 162) = "`BeinödVen`"
 Str(0, 34, 163) = "`Neuro sonst`"
 Str(0, 34, 164) = "`Weitere Befunde`"
 Str(0, 34, 165) = "`Schulung`"
 Str(0, 34, 166) = "`DMP`"
 Str(0, 34, 167) = "`DMSchulz`"
 Str(0, 34, 168) = "`DMSchL`"
 Str(0, 34, 169) = "`RRSchulz`"
 Str(0, 34, 170) = "`DMPhier`"
 Str(0, 34, 171) = "`HANr`"
 Str(0, 34, 172) = "`HANr2`"
 Str(0, 34, 173) = "`letzte Änderung`"
 Str(0, 34, 174) = "`Diagnosen`"
 Str(0, 34, 175) = "`Vorgestellt`"
 Str(0, 34, 176) = "`Versicherung`"
 Str(0, 34, 177) = "`AktZeit`"
 Str(0, 34, 178) = "`Ther1`"
 Str(0, 34, 179) = "`TherAkt`"
 Str(0, 34, 180) = "`obAn1eing`"
 Str(0, 34, 181) = "`obAn2eing`"
 Str(0, 34, 182) = "`obAnAeing`"
 Str(0, 34, 183) = "`obCheck`"
 Str(0, 34, 184) = "`obBZausgew`"
 Str(0, 34, 185) = "`obOSaufgek`"
 Str(0, 34, 186) = "`obPodAufgek`"
 Str(0, 34, 187) = "`obMBlAusgeh`"
 Str(0, 34, 188) = "`obSchulaufgek`"
 Str(0, 34, 189) = "`obDMPaufgekl`"
 Str(0, 34, 190) = "`obMedNetz`"
 Str(0, 34, 191) = "`Hausarzt`"
 Str(0, 34, 192) = "`ob`"
 Str(0, 34, 193) = "`QS`"
 Str(0, 34, 194) = "`QT`"
 Str(0, 34, 195) = "`Prim`"
 Str(0, 34, 196) = "`PrimaryKey`"
 Str(0, 34, 197) = "`Auswahl`"
 Str(0, 34, 198) = "`DMPhier`"
 Str(0, 34, 199) = "`GebDat`"
 Str(0, 34, 200) = "`jlD`"
 Str(0, 34, 201) = "`lÄnd`"
 Str(0, 34, 202) = "`Pat_ID`"
 Str(0, 34, 203) = "`Ther1`"
 Str(0, 34, 204) = "`Vorgestellt`"
 Str(0, 34, 205) = "`HausärzteAnamnesebogen_AccRel`"
 Str(0, 34, 206) = "`KassenlisteAnamnesebogen_AccRel`"
 ArtZ(0, 34) = 194
 ArtZ(1, 34) = 12
 Str(1, 34, 0) = "CREATE TABLE `anbogalt` ("
 Str(1, 34, 1) = " `Prim` int(2) unsigned NOT NULL COMMENT 'Primärschlüssel'"
 Str(1, 34, 2) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 34, 3) = " `Nachname` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '-'"
 Str(1, 34, 4) = " `Vorname` varchar(19) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 5) = " `NVorsatz` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 6) = " `Titel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 7) = " `Anrede` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 8) = " `GebDat` datetime DEFAULT NULL COMMENT ', geb.'"
 Str(1, 34, 9) = " `Tkz` tinyint(1) unsigned DEFAULT NULL COMMENT 'Tod-Kennzeichen'"
 Str(1, 34, 10) = " `Versicherungsart` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 11) = " `Diabetestyp` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetes Typ'"
 Str(1, 34, 12) = " `Diabetes seit` varchar(152) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 34, 13) = " `Tabletten seit` varchar(66) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Tabletten seit'"
 Str(1, 34, 14) = " `Insulin seit` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Insulin seit'"
 Str(1, 34, 15) = " `Grund für Vorstellung` varchar(721) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 16) = " `Familienanamnese` varchar(291) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 17) = " `Größe` double DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 18) = " `Gewicht` double DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 19) = " `Tendenz` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Tendenz'"
 Str(1, 34, 20) = " `DiabetesMedikament 1` varchar(48) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesmedikation:'"
 Str(1, 34, 21) = " `DiabetesMedikament 1 Menge` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 34, 22) = " `DiabetesMedikament 2` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 34, 23) = " `DiabetesMedikament 2 Menge` varchar(126) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 34, 24) = " `DiabetesMedikament 3` varchar(37) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 34, 25) = " `DiabetesMedikament 3 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 34, 26) = " `DiabetesMedikament 4` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 34, 27) = " `DiabetesMedikament 4 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 34, 28) = " `Insulinpumpe` tinyint(1) unsigned DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 29) = " `Insulinpumpe seit` varchar(250) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 34, 30) = " `Insulinpumpe Marke` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Marke:'"
 Str(1, 34, 31) = " `Broteinheiten gesamt` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Broteinheiten:gesamt'"
 Str(1, 34, 32) = " `Broteinheiten früh` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, früh'"
 Str(1, 34, 33) = " `Broteinheiten ZM früh` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zwischenmahlzeit vormittags'"
 Str(1, 34, 34) = " `Broteinheiten mittags` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 34, 35) = " `Broteinheiten nachmittags` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 34, 36) = " `Broteinheiten abends` varchar(13) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 34, 37) = " `Broteinheiten nachts` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 34, 38) = " `Essenszeit früh` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Essenszeiten:früh'"
 Str(1, 34, 39) = " `Essenszeit vormittags` varchar(34) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vormittags'"
 Str(1, 34, 40) = " `Essenszeit mittags` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 34, 41) = " `Essenszeit nachmittags` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 34, 42) = " `Essenszeit abends` varchar(18) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 34, 43) = " `Essenszeit spät` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, spät'"
 Str(1, 34, 44) = " `Spritz-Eß-Abstand früh` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritz-Eß-Abstand:früh'"
 Str(1, 34, 45) = " `Spritz-Eß-Abstand mittags` varchar(27) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 34, 46) = " `Spritz-Eß-Abstand abends` varchar(62) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 34, 47) = " `Spritzstelle früh` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritzstellen:früh'"
 Str(1, 34, 48) = " `Spritzstelle mittags` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 34, 49) = " `Spritzstelle abends` varchar(139) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 34, 50) = " `Spritzstelle nachts` varchar(141) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 34, 51) = " `Jahr letzte Diabetesschulung` varchar(129) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesschulung:'"
 Str(1, 34, 52) = " `Ort Schulung` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<in'"
 Str(1, 34, 53) = " `letztes HbA1c` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letztes HbA1c:'"
 Str(1, 34, 54) = " `gemessen am` datetime DEFAULT NULL COMMENT '<, gemessen'"
 Str(1, 34, 55) = " `vorherige Werte` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vorher:'"
 Str(1, 34, 56) = " `BZMessungen selbst` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckermessung:Selbstmessung?'"
 Str(1, 34, 57) = " `Gerät` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<:'"
 Str(1, 34, 58) = " `BZMessungen pW` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl d.Messungen pro Woche:'"
 Str(1, 34, 59) = " `BZMessungen pW ndE` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, davon nach dem Essen:'"
 Str(1, 34, 60) = " `BZMessungen p W nachts` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts:'"
 Str(1, 34, 61) = " `Aufschreiben` varchar(93) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Dokumentation:'"
 Str(1, 34, 62) = " `BZWerte v d Essen` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckerwerte vor dem Essen:'"
 Str(1, 34, 63) = " `BZWerte n d Essen` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nach dem Essen:'"
 Str(1, 34, 64) = " `UZ Tageszeit` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Unterzucker:Bevorzugte Tages-/Uhrzeit'"
 Str(1, 34, 65) = " `Unterzucker pM` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl der schweren (<50 mg/dl) pro Monat:'"
 Str(1, 34, 66) = " `UZ rechtzeitig` varchar(101) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, rechtzeitig bemerkt:'"
 Str(1, 34, 67) = " `Fremde Hilfe pa` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, fremde Hilfe deshalb nötig:'"
 Str(1, 34, 68) = " `Bewußtlos pa` varchar(81) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, bewußtlos deshalb:'"
 Str(1, 34, 69) = " `Keto pa` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Zahl der Ketoazidosen pro Jahr:'"
 Str(1, 34, 70) = " `BZgr300 pM` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Zahl der Blutzucker > 300 mg/dl pro Monat:'"
 Str(1, 34, 71) = " `Bluthochdruck` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Bluthochdruck:'"
 Str(1, 34, 72) = " `BHD seit` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit:'"
 Str(1, 34, 73) = " `BHD beh mit` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, behandelt mit:'"
 Str(1, 34, 74) = " `Blutdruckwerte` varchar(186) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckwerte:'"
 Str(1, 34, 75) = " `BDselbst` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckselbstmessung:'"
 Str(1, 34, 76) = " `Schwanger` varchar(43) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Aktuelle Schwangerschaft:'"
 Str(1, 34, 77) = " `Schwanger seit` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, seit:'"
 Str(1, 34, 78) = " `Augensp zuletzt` varchar(107) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Augenspiegelung:'"
 Str(1, 34, 79) = " `Augensp Befund` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 34, 80) = " `Netzhaut gelasert` varchar(142) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Netzhaut schon gelasert:'"
 Str(1, 34, 81) = " `Sehminderung unbehebbar` varchar(151) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', mit Brille nicht behebbare Sehminderung:'"
 Str(1, 34, 82) = " `Diabet Nierenschaden` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetischer Nierenschaden:'"
 Str(1, 34, 83) = " `Albumin zuletzt` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', letztes Albumin:'"
 Str(1, 34, 84) = " `erhöht?` varchar(52) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 34, 85) = " `Dialyse` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 86) = " `Dialyse seit` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 34, 87) = " `andere Nierenerkrankung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', andere Nierenerkrankung:'"
 Str(1, 34, 88) = " `Herzkrankheit` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Herzkrankheit:'"
 Str(1, 34, 89) = " `Angina pectoris` varchar(134) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 90) = " `Herzinfarkt` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 91) = " `Herzinfarkt wann` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 34, 92) = " `PTCA oder Stent` varchar(94) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 93) = " `Bypass kardial` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 94) = " `Bypass wann` varchar(388) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 34, 95) = " `Herzschwäche` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 96) = " `Herzkrankheit Beschreibung` varchar(213) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung:'"
 Str(1, 34, 97) = " `Hirndurchblutungsstörung` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 98) = " `Schlaganfall` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 99) = " `Beindurchblutungsstörung` varchar(136) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 100) = " `Schaufensterkrankheit` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 101) = " `Bypaß peripher` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 102) = " `Geschwür` varchar(174) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 103) = " `Amputation` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 104) = " `pAVK Beschreibung` varchar(97) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung der Beinbeschwerden:'"
 Str(1, 34, 105) = " `Ameisenlaufen` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 106) = " `Ameisen Ausmaß` varchar(123) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Ausmaß:'"
 Str(1, 34, 107) = " `Druckstellen` varchar(173) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 108) = " `Verformungen` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 109) = " `Verformungen Beschreibung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Beschreibung:'"
 Str(1, 34, 110) = " `Fußpflege` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 111) = " `Podologie` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 112) = " `Einlagen` varchar(91) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', diabetesgerechte orthopädische Einlagen/Schuhe:'"
 Str(1, 34, 113) = " `Neue Fußkomplikationen` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Neue Fußkomplikationen in den letzten 12 Monaten:'"
 Str(1, 34, 114) = " `Entleerungsstörungen Magen` varchar(130) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 115) = " `Entleerungsstörungen Harnblase` varchar(157) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 116) = " `Schwindel Aufstehen` varchar(121) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 117) = " `Folgeerkrankungen Haut` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 118) = " `Bewegungseinschränkungen` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 119) = " `Sexualstörung` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 120) = " `Sexualstörung seit` varchar(149) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 34, 121) = " `Weitere Anamnese` varchar(989) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 122) = " `Alkohol` varchar(148) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 123) = " `Tabak` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 124) = " `tabakex` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 125) = " `tabakbis` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 126) = " `tabakakt` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 127) = " `tabakmenge` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 128) = " `Weitere Medikation` varchar(298) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 129) = " `Liphypertrophien Abdomen` varchar(176) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Liphypertrophien:Abdomen'"
 Str(1, 34, 130) = " `Liphypertrophien Beine` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Beine:'"
 Str(1, 34, 131) = " `Liphypertrophien Arme` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Arme:'"
 Str(1, 34, 132) = " `Beinbefund` varchar(272) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 133) = " `Hyperkeratosen` varchar(207) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 134) = " `Ulcera` varchar(103) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 135) = " `Kraft Zehenheber` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Kraft:Zehenheber'"
 Str(1, 34, 136) = " `Kraft Zehenbeuger` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zehenbeuger:'"
 Str(1, 34, 137) = " `Kraft Knie` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Knie:'"
 Str(1, 34, 138) = " `ASR` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 139) = " `PSR` varchar(67) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 140) = " `Oberflächensensibilität` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 141) = " `Monofilamenttest` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 142) = " `Kalt-Warm` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Kalt-Warm-Diskrimination:'"
 Str(1, 34, 143) = " `Vibration IK` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Vibrationsempfinden Innenknöchel:'"
 Str(1, 34, 144) = " `Vibration Großzehe` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Großzehe:'"
 Str(1, 34, 145) = " `Puls Leiste` varchar(33) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Pulse:Leiste'"
 Str(1, 34, 146) = " `Puls Kniekehle` varchar(26) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Kniekehle:'"
 Str(1, 34, 147) = " `Puls Atp` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Innenknöchel:'"
 Str(1, 34, 148) = " `Puls Adp` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Fußrücken:'"
 Str(1, 34, 149) = " `RR` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruck:'"
 Str(1, 34, 150) = " `RRTurboMed` varchar(1362) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 151) = " `Herz` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 34, 152) = " `Lunge` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 153) = " `Bauch` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Abdomen:'"
 Str(1, 34, 154) = " `WS` varchar(56) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Wirbelsäule:'"
 Str(1, 34, 155) = " `NL` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nierenlager:'"
 Str(1, 34, 156) = " `SD` varchar(88) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Schilddrüse:'"
 Str(1, 34, 157) = " `Carotiden` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Halsschlagadern:'"
 Str(1, 34, 158) = " `NNH` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nasennebenhöhlen:'"
 Str(1, 34, 159) = " `Zähne` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 160) = " `Mundhöhle` varchar(64) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 34, 161) = " `LK` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Lymphknoten:'"
 Str(1, 34, 162) = " `BeinödVen` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beinödeme/ Venenkrankheiten:'"
 Str(1, 34, 163) = " `Neuro sonst` varchar(74) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Sonstige neurologische Befunde:'"
 Str(1, 34, 164) = " `Weitere Befunde` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', weitere Befunde:'"
 Str(1, 34, 165) = " `Schulung` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Schulungsbedarf'"
 Str(1, 34, 166) = " `DMP` varchar(85) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Pat. bei HA im DMP'"
 Str(1, 34, 167) = " `DMSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der DMP-Schulungen hier'"
 Str(1, 34, 168) = " `DMSchL` smallint(6) DEFAULT NULL COMMENT 'Zahl der abgerechneten DMP-Schulungen hier'"
 Str(1, 34, 169) = " `RRSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der Hypertonie-Schulungen hier'"
 Str(1, 34, 170) = " `DMPhier` datetime DEFAULT NULL COMMENT 'ob Pat hier im DMP'"
 Str(1, 34, 171) = " `HANr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 34, 172) = " `HANr2` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 34, 173) = " `letzte Änderung` datetime DEFAULT NULL COMMENT 'Datum der letzten Änderung'"
 Str(1, 34, 174) = " `Diagnosen` varchar(1071) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 175) = " `Vorgestellt` datetime DEFAULT NULL COMMENT 'Erstvorstellung'"
 Str(1, 34, 176) = " `Versicherung` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 177) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 34, 178) = " `Ther1` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, ICT, CSII'"
 Str(1, 34, 179) = " `TherAkt` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, ICT, CSII'"
 Str(1, 34, 180) = " `obAn1eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 1 eingegeben wurde'"
 Str(1, 34, 181) = " `obAn2eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 2 eingegeben wurde'"
 Str(1, 34, 182) = " `obAnAeing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt allgemein eingegeben wurde'"
 Str(1, 34, 183) = " `obCheck` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Checkliste vorliegt'"
 Str(1, 34, 184) = " `obBZausgew` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Blutzuckergerät ausgewechselt'"
 Str(1, 34, 185) = " `obOSaufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über orthopäd Schuhmacher aufgeklärt'"
 Str(1, 34, 186) = " `obPodAufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 34, 187) = " `obMBlAusgeh` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 34, 188) = " `obSchulaufgek` varchar(14) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 34, 189) = " `obDMPaufgekl` varchar(17) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 34, 190) = " `obMedNetz` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob von Med. Netz geschickt'"
 Str(1, 34, 191) = " `Hausarzt` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Hausarzt laut Anamnesebogen'"
 Str(1, 34, 192) = " `ob` tinyint(1) unsigned DEFAULT NULL COMMENT 'für verschiedene Aktionen'"
 Str(1, 34, 193) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 34, 194) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 34, 195) = "  PRIMARY KEY (`Prim`)"
 Str(1, 34, 196) = "  UNIQUE KEY `PrimaryKey` (`Prim`)"
 Str(1, 34, 197) = "  KEY `Auswahl` (`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 34, 198) = "  KEY `DMPhier` (`DMPhier`)"
 Str(1, 34, 199) = "  KEY `GebDat` (`GebDat`)"
 Str(1, 34, 200) = "  KEY `jlD` (`Jahr letzte Diabetesschulung`,`GebDat`)"
 Str(1, 34, 201) = "  KEY `lÄnd` (`letzte Änderung`)"
 Str(1, 34, 202) = "  KEY `Pat_ID` (`Pat_id`)"
 Str(1, 34, 203) = "  KEY `Ther1` (`Ther1`)"
 Str(1, 34, 204) = "  KEY `Vorgestellt` (`Vorgestellt`,`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 34, 205) = "  KEY `HausärzteAnamnesebogen_AccRel` (`HANr`)"
 Str(1, 34, 206) = "  KEY `KassenlisteAnamnesebogen_AccRel` (`Versicherung`)"
 Str(1, 34, 207) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr34

Sub FüllStr35()
 Str(0, 35, 0) = "au"
 Str(0, 35, 1) = "`FID`"
 Str(0, 35, 2) = "`Pat_ID`"
 Str(0, 35, 3) = "`ZeitPunkt`"
 Str(0, 35, 4) = "`Beginn`"
 Str(0, 35, 5) = "`Ende`"
 Str(0, 35, 6) = "`ICDs`"
 Str(0, 35, 7) = "`absPos`"
 Str(0, 35, 8) = "`AktZeit`"
 Str(0, 35, 9) = "`StByte`"
 Str(0, 35, 10) = "`Auswahl`"
 Str(0, 35, 11) = "`FälleAU`"
 Str(0, 35, 12) = "`FID`"
 Str(0, 35, 13) = "`NamenAU`"
 Str(0, 35, 14) = "`F??lleAU_AccRel`"
 ArtZ(0, 35) = 9
 ArtZ(1, 35) = 4
 ArtZ(2, 35) = 1
 Str(1, 35, 0) = "CREATE TABLE `au` ("
 Str(1, 35, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 35, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 35, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 35, 4) = " `Beginn` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6285 1. Hälfte'"
 Str(1, 35, 5) = " `Ende` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6285 2. Hälfte'"
 Str(1, 35, 6) = " `ICDs` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6286'"
 Str(1, 35, 7) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 35, 8) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 35, 9) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 35, 10) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Beginn`,`Ende`)"
 Str(1, 35, 11) = "  KEY `FälleAU` (`FID`)"
 Str(1, 35, 12) = "  KEY `FID` (`FID`)"
 Str(1, 35, 13) = "  KEY `NamenAU` (`Pat_ID`)"
 Str(1, 35, 14) = "  CONSTRAINT `F??lleAU_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 35, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr35

Sub FüllStr36()
 Str(0, 36, 0) = "augenbefunde"
 Str(0, 36, 1) = "`ID`"
 Str(0, 36, 2) = "`Pat_ID`"
 Str(0, 36, 3) = "`Datum`"
 Str(0, 36, 4) = "`Retinopathie`"
 Str(0, 36, 5) = "`Klassifikation`"
 Str(0, 36, 6) = "`Maculopathie`"
 Str(0, 36, 7) = "`Sonstiges`"
 Str(0, 36, 8) = "`Visus re`"
 Str(0, 36, 9) = "`Visus li`"
 Str(0, 36, 10) = "`KontrolleinMonaten`"
 Str(0, 36, 11) = "`Augenarzt`"
 Str(0, 36, 12) = "`DokName`"
 Str(0, 36, 13) = "`DokPfad`"
 Str(0, 36, 14) = "`verglichen`"
 Str(0, 36, 15) = "`ID`"
 Str(0, 36, 16) = "`PrimaryKey`"
 Str(0, 36, 17) = "`Auswahl`"
 Str(0, 36, 18) = "`NamenAugenbefunde`"
 Str(0, 36, 19) = "`Text`"
 Str(0, 36, 20) = "`NamenAugenbefunde_AccRel`"
 ArtZ(0, 36) = 14
 ArtZ(1, 36) = 5
 ArtZ(2, 36) = 1
 Str(1, 36, 0) = "CREATE TABLE `augenbefunde` ("
 Str(1, 36, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 36, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 36, 3) = " `Datum` datetime DEFAULT NULL"
 Str(1, 36, 4) = " `Retinopathie` longtext COLLATE latin1_german2_ci"
 Str(1, 36, 5) = " `Klassifikation` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 36, 6) = " `Maculopathie` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 36, 7) = " `Sonstiges` longtext COLLATE latin1_german2_ci"
 Str(1, 36, 8) = " `Visus re` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 36, 9) = " `Visus li` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 36, 10) = " `KontrolleinMonaten` float DEFAULT NULL"
 Str(1, 36, 11) = " `Augenarzt` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 36, 12) = " `DokName` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 36, 13) = " `DokPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 36, 14) = " `verglichen` datetime DEFAULT NULL"
 Str(1, 36, 15) = "  PRIMARY KEY (`ID`)"
 Str(1, 36, 16) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 36, 17) = "  KEY `Auswahl` (`Pat_ID`,`DokPfad`(255))"
 Str(1, 36, 18) = "  KEY `NamenAugenbefunde` (`Pat_ID`)"
 Str(1, 36, 19) = "  KEY `Text` (`Pat_ID`,`Retinopathie`(255))"
 Str(1, 36, 20) = "  CONSTRAINT `NamenAugenbefunde_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 36, 21) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr36

Sub FüllStr37()
 Str(0, 37, 0) = "briefe"
 Str(0, 37, 1) = "`FID`"
 Str(0, 37, 2) = "`Pat_ID`"
 Str(0, 37, 3) = "`ZeitPunkt`"
 Str(0, 37, 4) = "`Pfad`"
 Str(0, 37, 5) = "`Art`"
 Str(0, 37, 6) = "`Name`"
 Str(0, 37, 7) = "`Quelldatum`"
 Str(0, 37, 8) = "`Typ`"
 Str(0, 37, 9) = "`AktZeit`"
 Str(0, 37, 10) = "`DokGroe`"
 Str(0, 37, 11) = "`DokAenD`"
 Str(0, 37, 12) = "`QS`"
 Str(0, 37, 13) = "`QT`"
 Str(0, 37, 14) = "`absPos`"
 Str(0, 37, 15) = "`StByte`"
 Str(0, 37, 16) = "`FälleBriefe`"
 Str(0, 37, 17) = "`NamenBriefe`"
 Str(0, 37, 18) = "`Quelldatum`"
 Str(0, 37, 19) = "`Auswahl`"
 Str(0, 37, 20) = "`Name`"
 Str(0, 37, 21) = "`F??lleBriefe_AccRel`"
 ArtZ(0, 37) = 15
 ArtZ(1, 37) = 5
 ArtZ(2, 37) = 1
 Str(1, 37, 0) = "CREATE TABLE `briefe` ("
 Str(1, 37, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 37, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 37, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 37, 4) = " `Pfad` varchar(128) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 37, 5) = " `Art` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 37, 6) = " `Name` varchar(1435) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 37, 7) = " `Quelldatum` datetime DEFAULT NULL COMMENT 'Datum, auf das sich das Dokument bezieht'"
 Str(1, 37, 8) = " `Typ` varchar(78) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 37, 9) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 37, 10) = " `DokGroe` int(10) DEFAULT NULL COMMENT 'Größe der Datei'"
 Str(1, 37, 11) = " `DokAenD` datetime DEFAULT NULL COMMENT 'Dokument-letzte Änderung'"
 Str(1, 37, 12) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 37, 13) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 37, 14) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 37, 15) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 37, 16) = "  KEY `FälleBriefe` (`FID`)"
 Str(1, 37, 17) = "  KEY `NamenBriefe` (`Pat_ID`)"
 Str(1, 37, 18) = "  KEY `Quelldatum` (`Quelldatum`)"
 Str(1, 37, 19) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Name`(200)) USING BTREE"
 Str(1, 37, 20) = "  KEY `Name` (`Name`(767))"
 Str(1, 37, 21) = "  CONSTRAINT `F??lleBriefe_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 37, 22) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr37

Sub FüllStr38()
 Str(0, 38, 0) = "custo"
 Str(1, 38, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `custo` AS SELECT concat(`n`.`Nachname`,', ',`n`.`Vorname`,', ',cast(`n`.`GebDat` as date)) AS `Patient`,`n`.`Pat_ID` AS `pat_id`,`e`.`ZeitPunkt` AS `zeitpunkt`,`e`.`Art` AS `art` FROM (`eintraege` `e` LEFT JOIN `namen` `n` on((`e`.`Pat_ID` = `n`.`Pat_ID`))) WHERE (`e`.`Art` in ('EKG','GDT','Lufu','LZRR')) ORDER BY `n`.`Nachname`,`n`.`Vorname`,`n`.`GebDat`,`e`.`ZeitPunkt`"
End Sub ' FüllStr38

Sub FüllStr39()
 Str(0, 39, 0) = "dateien"
 Str(0, 39, 1) = "`id`"
 Str(0, 39, 2) = "`pfad`"
 Str(0, 39, 3) = "`geändert`"
 Str(0, 39, 4) = "`geprüft`"
 Str(0, 39, 5) = "`id`"
 Str(0, 39, 6) = "`pfad`"
 Str(0, 39, 7) = "`id`"
 ArtZ(0, 39) = 4
 ArtZ(1, 39) = 2
 ArtZ(2, 39) = 1
 Str(1, 39, 0) = "CREATE TABLE `dateien` ("
 Str(1, 39, 1) = " `id` int(2) unsigned NOT NULL COMMENT 'Bezug zu faxe'"
 Str(1, 39, 2) = " `pfad` varchar(500) COLLATE latin1_german2_ci NOT NULL COMMENT 'Dateiname samt Pfad'"
 Str(1, 39, 3) = " `geändert` datetime NOT NULL COMMENT 'Zeile eingetragen'"
 Str(1, 39, 4) = " `geprüft` tinyint(1) unsigned NOT NULL COMMENT 'Zeile geprüft'"
 Str(1, 39, 5) = "  KEY `id` (`id`)"
 Str(1, 39, 6) = "  KEY `pfad` (`pfad`)"
 Str(1, 39, 7) = "  CONSTRAINT `id` FOREIGN KEY (`id`) REFERENCES `faxe` (`ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 39, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr39

Sub FüllStr40()
 Str(0, 40, 0) = "desktop"
 Str(0, 40, 1) = "`id`"
 Str(0, 40, 2) = "`IDS`"
 Str(0, 40, 3) = "`Pat_ID`"
 Str(0, 40, 4) = "`erstZP`"
 Str(0, 40, 5) = "`exoL`"
 Str(0, 40, 6) = "`hideT`"
 Str(0, 40, 7) = "`iconPath`"
 Str(0, 40, 8) = "`noteBkColor`"
 Str(0, 40, 9) = "`noteFgColor`"
 Str(0, 40, 10) = "`positionBottom`"
 Str(0, 40, 11) = "`positionLeft`"
 Str(0, 40, 12) = "`positionRight`"
 Str(0, 40, 13) = "`positionTop`"
 Str(0, 40, 14) = "`showAsNote`"
 Str(0, 40, 15) = "`syncInfoList`"
 Str(0, 40, 16) = "`titel`"
 Str(0, 40, 17) = "`toolTipText`"
 Str(0, 40, 18) = "`verankert`"
 Str(0, 40, 19) = "`absPos`"
 Str(0, 40, 20) = "`AktZeit`"
 Str(0, 40, 21) = "`StByte`"
 Str(0, 40, 22) = "`id`"
 Str(0, 40, 23) = "`titel`"
 Str(0, 40, 24) = "`pat_id`"
 ArtZ(0, 40) = 21
 ArtZ(1, 40) = 3
 Str(1, 40, 0) = "CREATE TABLE `desktop` ("
 Str(1, 40, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'"
 Str(1, 40, 2) = " `IDS` varchar(10) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'id='"
 Str(1, 40, 3) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 40, 4) = " `erstZP` datetime DEFAULT NULL COMMENT 'erstellungsZeitpunkt'"
 Str(1, 40, 5) = " `exoL` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'executeonLoad'"
 Str(1, 40, 6) = " `hideT` tinyint(1) DEFAULT NULL COMMENT 'hideTitel'"
 Str(1, 40, 7) = " `iconPath` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'iconPath'"
 Str(1, 40, 8) = " `noteBkColor` int(10) DEFAULT NULL COMMENT 'noteBkColor'"
 Str(1, 40, 9) = " `noteFgColor` int(10) DEFAULT NULL COMMENT 'noteFgColor'"
 Str(1, 40, 10) = " `positionBottom` int(10) DEFAULT NULL COMMENT 'positionBottom'"
 Str(1, 40, 11) = " `positionLeft` int(10) DEFAULT NULL COMMENT 'positionLeft'"
 Str(1, 40, 12) = " `positionRight` int(10) DEFAULT NULL COMMENT 'positionRight'"
 Str(1, 40, 13) = " `positionTop` int(10) DEFAULT NULL COMMENT 'positionTop'"
 Str(1, 40, 14) = " `showAsNote` tinyint(1) DEFAULT NULL COMMENT 'showAsNote'"
 Str(1, 40, 15) = " `syncInfoList` varchar(12) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'syncInfoList'"
 Str(1, 40, 16) = " `titel` varchar(416) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'titel'"
 Str(1, 40, 17) = " `toolTipText` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'toolTipText'"
 Str(1, 40, 18) = " `verankert` tinyint(1) DEFAULT NULL COMMENT 'verankert'"
 Str(1, 40, 19) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 40, 20) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 40, 21) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 40, 22) = "  PRIMARY KEY (`id`)"
 Str(1, 40, 23) = "  KEY `titel` (`titel`)"
 Str(1, 40, 24) = "  KEY `pat_id` (`Pat_ID`,`titel`) USING BTREE"
 Str(1, 40, 25) = " ENGINE=InnoDB AUTO_INCREMENT=453405 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr40

Sub FüllStr41()
 Str(0, 41, 0) = "diagnosen"
 Str(0, 41, 1) = "`ID1`"
 Str(0, 41, 2) = "`FID`"
 Str(0, 41, 3) = "`Pat_id`"
 Str(0, 41, 4) = "`DiagDatum`"
 Str(0, 41, 5) = "`DiagSicherheit`"
 Str(0, 41, 6) = "`DiagText`"
 Str(0, 41, 7) = "`DiagSeite`"
 Str(0, 41, 8) = "`DiagAttr`"
 Str(0, 41, 9) = "`ICD`"
 Str(0, 41, 10) = "`obDauer`"
 Str(0, 41, 11) = "`intBemerk`"
 Str(0, 41, 12) = "`absPos`"
 Str(0, 41, 13) = "`AktZeit`"
 Str(0, 41, 14) = "`StByte`"
 Str(0, 41, 15) = "`AusnBegr`"
 Str(0, 41, 16) = "`f6010`"
 Str(0, 41, 17) = "`f6011`"
 Str(0, 41, 18) = "`ID1`"
 Str(0, 41, 19) = "`Auswahl`"
 Str(0, 41, 20) = "`DiagSuch`"
 Str(0, 41, 21) = "`DiagText`"
 Str(0, 41, 22) = "`FälleDiagnosen`"
 Str(0, 41, 23) = "`ICD`"
 Str(0, 41, 24) = "`F??lleDiagnosen_AccRel`"
 ArtZ(0, 41) = 17
 ArtZ(1, 41) = 6
 ArtZ(2, 41) = 1
 Str(1, 41, 0) = "CREATE TABLE `diagnosen` ("
 Str(1, 41, 1) = " `ID1` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 41, 2) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 41, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT 'Bezug auf Anamneseblattt'"
 Str(1, 41, 4) = " `DiagDatum` datetime DEFAULT NULL"
 Str(1, 41, 5) = " `DiagSicherheit` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6003'"
 Str(1, 41, 6) = " `DiagText` longtext COLLATE latin1_german2_ci"
 Str(1, 41, 7) = " `DiagSeite` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6004'"
 Str(1, 41, 8) = " `DiagAttr` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6006 Diagnosenattribut (optionale Erläuterung)'"
 Str(1, 41, 9) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 41, 10) = " `obDauer` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Dauerdiagnose'"
 Str(1, 41, 11) = " `intBemerk` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6009 interne Bemerkung'"
 Str(1, 41, 12) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 41, 13) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 41, 14) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 41, 15) = " `AusnBegr` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6008 Ausnahmebegründung'"
 Str(1, 41, 16) = " `f6010` tinyint(1) DEFAULT NULL COMMENT '6010 Diagnose gelöscht (Karteikarteneintrag: bdd)'"
 Str(1, 41, 17) = " `f6011` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6011 8.12.10: bisher nur ""TM#?""'"
 Str(1, 41, 18) = "  PRIMARY KEY (`ID1`)"
 Str(1, 41, 19) = "  KEY `Auswahl` (`Pat_id`,`DiagDatum`,`DiagSicherheit`,`DiagSeite`,`DiagAttr`,`DiagText`(255),`ICD`,`obDauer`)"
 Str(1, 41, 20) = "  KEY `DiagSuch` (`Pat_id`,`ICD`,`DiagSicherheit`,`DiagSeite`)"
 Str(1, 41, 21) = "  KEY `DiagText` (`Pat_id`,`DiagText`(255))"
 Str(1, 41, 22) = "  KEY `FälleDiagnosen` (`FID`)"
 Str(1, 41, 23) = "  KEY `ICD` (`ICD`)"
 Str(1, 41, 24) = "  CONSTRAINT `F??lleDiagnosen_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 41, 25) = " ENGINE=InnoDB AUTO_INCREMENT=3272138 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr41

Sub FüllStr42()
 Str(0, 42, 0) = "diagnosen exportiert"
 Str(0, 42, 1) = "`ID`"
 Str(0, 42, 2) = "`Datum`"
 Str(0, 42, 3) = "`Pat_id`"
 Str(0, 42, 4) = "`ICD`"
 Str(0, 42, 5) = "`Diagnose`"
 Str(0, 42, 6) = "`übertragen`"
 Str(0, 42, 7) = "`ID`"
 Str(0, 42, 8) = "`PrimaryKey`"
 Str(0, 42, 9) = "`ID`"
 ArtZ(0, 42) = 6
 ArtZ(1, 42) = 3
 Str(1, 42, 0) = "CREATE TABLE `diagnosen exportiert` ("
 Str(1, 42, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Reihenfolge'"
 Str(1, 42, 2) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum'"
 Str(1, 42, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!Pat_id'"
 Str(1, 42, 4) = " `ICD` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD-Nummer der Diagnose'"
 Str(1, 42, 5) = " `Diagnose` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Text der Diagnose'"
 Str(1, 42, 6) = " `übertragen` datetime DEFAULT NULL COMMENT '"""", übertragen'"
 Str(1, 42, 7) = "  PRIMARY KEY (`ID`)"
 Str(1, 42, 8) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 42, 9) = "  KEY `ID` (`Pat_id`)"
 Str(1, 42, 10) = " ENGINE=InnoDB AUTO_INCREMENT=22941 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr42

Sub FüllStr43()
 Str(0, 43, 0) = "diagnosenexport"
 Str(0, 43, 1) = "`ID`"
 Str(0, 43, 2) = "`Name`"
 Str(0, 43, 3) = "`Pat_id`"
 Str(0, 43, 4) = "`ICD`"
 Str(0, 43, 5) = "`Diagnose`"
 Str(0, 43, 6) = "`Status`"
 Str(0, 43, 7) = "`Protokoll`"
 Str(0, 43, 8) = "`nurQuart`"
 Str(0, 43, 9) = "`Zeitpunkt`"
 Str(0, 43, 10) = "`ID`"
 Str(0, 43, 11) = "`ID`"
 Str(0, 43, 12) = "`pat_ID`"
 Str(0, 43, 13) = "`Suche`"
 ArtZ(0, 43) = 9
 ArtZ(1, 43) = 4
 Str(1, 43, 0) = "CREATE TABLE `diagnosenexport` ("
 Str(1, 43, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 43, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 43, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 43, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 43, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 43, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 43, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 43, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 43, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 43, 10) = "  PRIMARY KEY (`ID`)"
 Str(1, 43, 11) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 43, 12) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 43, 13) = "  KEY `Suche` (`Pat_id`,`ICD`)"
 Str(1, 43, 14) = " ENGINE=InnoDB AUTO_INCREMENT=1788 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr43

Sub FüllStr44()
 Str(0, 44, 0) = "diagnoseng1"
 Str(0, 44, 1) = "`lfdnr`"
 Str(0, 44, 2) = "`gruppe`"
 Str(0, 44, 3) = "`rf`"
 Str(0, 44, 4) = "`lfdnr`"
 Str(0, 44, 5) = "`gruppe`"
 Str(0, 44, 6) = "`rf`"
 ArtZ(0, 44) = 3
 ArtZ(1, 44) = 3
 Str(1, 44, 0) = "CREATE TABLE `diagnoseng1` ("
 Str(1, 44, 1) = " `lfdnr` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 44, 2) = " `gruppe` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Gruppenname'"
 Str(1, 44, 3) = " `rf` int(10) unsigned NOT NULL DEFAULT '0' COMMENT 'Reihenfolge'"
 Str(1, 44, 4) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 44, 5) = "  KEY `gruppe` (`gruppe`) USING BTREE"
 Str(1, 44, 6) = "  KEY `rf` (`rf`)"
 Str(1, 44, 7) = " ENGINE=InnoDB AUTO_INCREMENT=45 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Diagnosengruppierung 1'"
End Sub ' FüllStr44

Sub FüllStr45()
 Str(0, 45, 0) = "diagreihe"
 Str(0, 45, 1) = "`lfdnr`"
 Str(0, 45, 2) = "`ICD`"
 Str(0, 45, 3) = "`dg1`"
 Str(0, 45, 4) = "`dg2`"
 Str(0, 45, 5) = "`rf`"
 Str(0, 45, 6) = "`lfdnr`"
 Str(0, 45, 7) = "`ICD`"
 Str(0, 45, 8) = "`rf`"
 Str(0, 45, 9) = "`dg1`"
 Str(0, 45, 10) = "`dg1`"
 ArtZ(0, 45) = 5
 ArtZ(1, 45) = 4
 ArtZ(2, 45) = 1
 Str(1, 45, 0) = "CREATE TABLE `diagreihe` ("
 Str(1, 45, 1) = " `lfdnr` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 45, 2) = " `ICD` varchar(10) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 45, 3) = " `dg1` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Diagnosenguppierung 1'"
 Str(1, 45, 4) = " `dg2` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Diagnosenguppierung 2'"
 Str(1, 45, 5) = " `rf` int(10) unsigned NOT NULL DEFAULT '0' COMMENT 'Reihenfolge für Arztbrief'"
 Str(1, 45, 6) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 45, 7) = "  UNIQUE KEY `ICD` (`ICD`) USING BTREE"
 Str(1, 45, 8) = "  KEY `rf` (`rf`)"
 Str(1, 45, 9) = "  KEY `dg1` (`dg1`)"
 Str(1, 45, 10) = "  CONSTRAINT `dg1` FOREIGN KEY (`dg1`) REFERENCES `diagnoseng1` (`gruppe`) ON UPDATE CASCADE"
 Str(1, 45, 11) = " ENGINE=InnoDB AUTO_INCREMENT=1238 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Reihenfolge der Diagnosen für Arztbrief'"
End Sub ' FüllStr45

Sub FüllStr46()
 Str(0, 46, 0) = "dmp-uschr"
 Str(0, 46, 1) = "`Pat_id`"
 Str(0, 46, 2) = "`U1`"
 Str(0, 46, 3) = "`U2`"
 Str(0, 46, 4) = "`U3`"
 Str(0, 46, 5) = "`Arztwechsel`"
 Str(0, 46, 6) = "`Pat_id`"
 ArtZ(0, 46) = 5
 ArtZ(1, 46) = 1
 Str(1, 46, 0) = "CREATE TABLE `dmp-uschr` ("
 Str(1, 46, 1) = " `Pat_id` int(10) DEFAULT NULL COMMENT 'Bezug auf Namen'"
 Str(1, 46, 2) = " `U1` datetime DEFAULT NULL COMMENT 'vorliegendes Blatt mit DMP-Unterschrift'"
 Str(1, 46, 3) = " `U2` datetime DEFAULT NULL COMMENT 'vorliegendes 2. Blatt mit DMP-Unterschrift'"
 Str(1, 46, 4) = " `U3` datetime DEFAULT NULL COMMENT 'vorliegendes 3. Blatt mit DMP-Unterschrift'"
 Str(1, 46, 5) = " `Arztwechsel` bit(1) DEFAULT NULL COMMENT 'danach Arztwechsel'"
 Str(1, 46, 6) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 46, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr46

Sub FüllStr47()
 Str(0, 47, 0) = "dmpreihe"
 Str(0, 47, 1) = "`Abk`"
 Str(0, 47, 2) = "`Art`"
 Str(0, 47, 3) = "`KarteiDatum`"
 Str(0, 47, 4) = "`exportiert`"
 Str(0, 47, 5) = "`DokuDatum`"
 Str(0, 47, 6) = "`obvoll`"
 Str(0, 47, 7) = "`ok`"
 Str(0, 47, 8) = "`ausgedruckt`"
 Str(0, 47, 9) = "`NachName`"
 Str(0, 47, 10) = "`VorName`"
 Str(0, 47, 11) = "`GebDat`"
 Str(0, 47, 12) = "`Pat_id`"
 Str(0, 47, 13) = "`StByte`"
 Str(0, 47, 14) = "`AktZeit`"
 Str(0, 47, 15) = "`lanrid`"
 Str(0, 47, 16) = "`Pat_ID`"
 Str(0, 47, 17) = "`dmpreihenamen`"
 ArtZ(0, 47) = 15
 ArtZ(1, 47) = 1
 ArtZ(2, 47) = 1
 Str(1, 47, 0) = "CREATE TABLE `dmpreihe` ("
 Str(1, 47, 1) = " `Abk` varchar(30) CHARACTER SET latin1 COLLATE latin1_german1_ci DEFAULT NULL COMMENT 'Abkürzung der DMP-Art'"
 Str(1, 47, 2) = " `Art` varchar(2) COLLATE latin1_german2_ci NOT NULL COMMENT 'ED = Erstdoku, FD = Folgedoku'"
 Str(1, 47, 3) = " `KarteiDatum` date DEFAULT NULL COMMENT 'Datum des Karteikarteneintrags der Dokumentation'"
 Str(1, 47, 4) = " `exportiert` datetime DEFAULT NULL COMMENT 'Datum des Exports'"
 Str(1, 47, 5) = " `DokuDatum` datetime DEFAULT NULL COMMENT 'Datum der Dokumentation'"
 Str(1, 47, 6) = " `obvoll` bit(1) DEFAULT NULL COMMENT 'ob vollständig'"
 Str(1, 47, 7) = " `ok` bit(1) NOT NULL COMMENT 'ob ""ok""'"
 Str(1, 47, 8) = " `ausgedruckt` bit(1) NOT NULL COMMENT 'ob ""ausgedruckt""'"
 Str(1, 47, 9) = " `NachName` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 47, 10) = " `VorName` varchar(22) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 47, 11) = " `GebDat` date DEFAULT NULL"
 Str(1, 47, 12) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 47, 13) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 47, 14) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungzeit'"
 Str(1, 47, 15) = " `lanrid` int(3) unsigned NOT NULL COMMENT 'Bezug auf lanrpraxis.id'"
 Str(1, 47, 16) = "  KEY `Pat_ID` (`Pat_id`)"
 Str(1, 47, 17) = "  CONSTRAINT `dmpreihenamen` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 47, 18) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr47

Sub FüllStr48()
 Str(0, 48, 0) = "dokabkop"
 Str(0, 48, 1) = "`DokPfad`"
 Str(0, 48, 2) = "`AktZeit`"
 Str(0, 48, 3) = "`abgehakt`"
 Str(0, 48, 4) = "`DokPfad`"
 ArtZ(0, 48) = 3
 ArtZ(1, 48) = 1
 Str(1, 48, 0) = "CREATE TABLE `dokabkop` ("
 Str(1, 48, 1) = " `DokPfad` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 48, 2) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 48, 3) = " `abgehakt` bit(1) DEFAULT NULL"
 Str(1, 48, 4) = "  KEY `DokPfad` (`DokPfad`)"
 Str(1, 48, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr48

Sub FüllStr49()
 Str(0, 49, 0) = "dokumente"
 Str(0, 49, 1) = "`FID`"
 Str(0, 49, 2) = "`Pat_ID`"
 Str(0, 49, 3) = "`ZeitPunkt`"
 Str(0, 49, 4) = "`DokPfad`"
 Str(0, 49, 5) = "`DokArt`"
 Str(0, 49, 6) = "`DokName`"
 Str(0, 49, 7) = "`Quelldatum`"
 Str(0, 49, 8) = "`absPos`"
 Str(0, 49, 9) = "`AktZeit`"
 Str(0, 49, 10) = "`DokGroe`"
 Str(0, 49, 11) = "`DokAenD`"
 Str(0, 49, 12) = "`QS`"
 Str(0, 49, 13) = "`QT`"
 Str(0, 49, 14) = "`StByte`"
 Str(0, 49, 15) = "`DokName`"
 Str(0, 49, 16) = "`DokPfad`"
 Str(0, 49, 17) = "`PIDokPfad`"
 Str(0, 49, 18) = "`Quelldatum`"
 Str(0, 49, 19) = "`ZeitPunkt`"
 Str(0, 49, 20) = "`Auswahl`"
 Str(0, 49, 21) = "`PIDDokName`"
 Str(0, 49, 22) = "`FID`"
 Str(0, 49, 23) = "`F??lleDokumente_AccRel`"
 Str(0, 49, 24) = "`NamenDokumente_AccRel`"
 ArtZ(0, 49) = 14
 ArtZ(1, 49) = 8
 ArtZ(2, 49) = 2
 Str(1, 49, 0) = "CREATE TABLE `dokumente` ("
 Str(1, 49, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 49, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 49, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 49, 4) = " `DokPfad` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 49, 5) = " `DokArt` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 49, 6) = " `DokName` varchar(1578) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 49, 7) = " `Quelldatum` datetime DEFAULT NULL COMMENT 'Datum, auf das sich das Dokument bezieht'"
 Str(1, 49, 8) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 49, 9) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 49, 10) = " `DokGroe` int(10) DEFAULT NULL COMMENT 'Dokument-Größe'"
 Str(1, 49, 11) = " `DokAenD` datetime DEFAULT NULL COMMENT 'Dokument-letzte Änderung'"
 Str(1, 49, 12) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 49, 13) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 49, 14) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 49, 15) = "  KEY `DokName` (`DokName`(767))"
 Str(1, 49, 16) = "  KEY `DokPfad` (`DokPfad`)"
 Str(1, 49, 17) = "  KEY `PIDokPfad` (`Pat_ID`,`DokPfad`)"
 Str(1, 49, 18) = "  KEY `Quelldatum` (`Quelldatum`)"
 Str(1, 49, 19) = "  KEY `ZeitPunkt` (`ZeitPunkt`)"
 Str(1, 49, 20) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`DokArt`,`DokName`(80)) USING BTREE"
 Str(1, 49, 21) = "  KEY `PIDDokName` (`Pat_ID`,`DokName`(767))"
 Str(1, 49, 22) = "  KEY `FID` (`FID`,`DokName`(767)) USING BTREE"
 Str(1, 49, 23) = "  CONSTRAINT `F??lleDokumente_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 49, 24) = "  CONSTRAINT `NamenDokumente_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 49, 25) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr49

Sub FüllStr50()
 Str(0, 50, 0) = "dokumente abgehakt"
 Str(0, 50, 1) = "`DokPfad`"
 Str(0, 50, 2) = "`AktZeit`"
 Str(0, 50, 3) = "`abgehakt`"
 Str(0, 50, 4) = "`ungueltig`"
 Str(0, 50, 5) = "`DokPfad`"
 ArtZ(0, 50) = 4
 ArtZ(1, 50) = 1
 Str(1, 50, 0) = "CREATE TABLE `dokumente abgehakt` ("
 Str(1, 50, 1) = " `DokPfad` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 50, 2) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 50, 3) = " `abgehakt` bit(1) DEFAULT NULL"
 Str(1, 50, 4) = " `ungueltig` bit(1) DEFAULT NULL"
 Str(1, 50, 5) = "  KEY `DokPfad` (`DokPfad`)"
 Str(1, 50, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr50

Sub FüllStr51()
 Str(0, 51, 0) = "ebm2000plus"
 Str(0, 51, 1) = "`Leistung`"
 Str(0, 51, 2) = "`Titel`"
 Str(0, 51, 3) = "`Punktwert`"
 Str(0, 51, 4) = "`Euro`"
 Str(0, 51, 5) = "`Bericht`"
 Str(0, 51, 6) = "`Text`"
 Str(0, 51, 7) = "`Betr`"
 Str(0, 51, 8) = "`Schul`"
 Str(0, 51, 9) = "`Typ1`"
 Str(0, 51, 10) = "`Typ2`"
 Str(0, 51, 11) = "`Gest`"
 Str(0, 51, 12) = "`DFS`"
 Str(0, 51, 13) = "`DMP`"
 Str(0, 51, 14) = "`AOK`"
 Str(0, 51, 15) = "`BKK`"
 Str(0, 51, 16) = "`BKN`"
 Str(0, 51, 17) = "`EK`"
 Str(0, 51, 18) = "`IKK`"
 Str(0, 51, 19) = "`LKK`"
 Str(0, 51, 20) = "`Üw`"
 Str(0, 51, 21) = "`Insulin`"
 Str(0, 51, 22) = "`ICT`"
 Str(0, 51, 23) = "`CSII`"
 Str(0, 51, 24) = "`Erst`"
 Str(0, 51, 25) = "`Folge`"
 Str(0, 51, 26) = "`fid`"
 Str(0, 51, 27) = "`Leistung`"
 Str(0, 51, 28) = "`Titel`"
 ArtZ(0, 51) = 26
 ArtZ(1, 51) = 2
 Str(1, 51, 0) = "CREATE TABLE `ebm2000plus` ("
 Str(1, 51, 1) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Leistungsziffer'"
 Str(1, 51, 2) = " `Titel` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kurztext'"
 Str(1, 51, 3) = " `Punktwert` decimal(10,1) DEFAULT NULL COMMENT 'Punktwert'"
 Str(1, 51, 4) = " `Euro` decimal(15,4) DEFAULT NULL COMMENT '€'"
 Str(1, 51, 5) = " `Bericht` bit(1) DEFAULT NULL COMMENT 'Berichtspflicht'"
 Str(1, 51, 6) = " `Text` longtext COLLATE latin1_german2_ci COMMENT 'restlicher Leistungstext'"
 Str(1, 51, 7) = " `Betr` bit(1) DEFAULT NULL COMMENT 'Betreuung'"
 Str(1, 51, 8) = " `Schul` bit(1) DEFAULT NULL COMMENT 'Schulung'"
 Str(1, 51, 9) = " `Typ1` bit(1) DEFAULT NULL"
 Str(1, 51, 10) = " `Typ2` bit(1) DEFAULT NULL"
 Str(1, 51, 11) = " `Gest` bit(1) DEFAULT NULL"
 Str(1, 51, 12) = " `DFS` bit(1) DEFAULT NULL"
 Str(1, 51, 13) = " `DMP` bit(1) DEFAULT NULL"
 Str(1, 51, 14) = " `AOK` bit(1) DEFAULT NULL"
 Str(1, 51, 15) = " `BKK` bit(1) DEFAULT NULL"
 Str(1, 51, 16) = " `BKN` bit(1) DEFAULT NULL"
 Str(1, 51, 17) = " `EK` bit(1) DEFAULT NULL"
 Str(1, 51, 18) = " `IKK` bit(1) DEFAULT NULL"
 Str(1, 51, 19) = " `LKK` bit(1) DEFAULT NULL"
 Str(1, 51, 20) = " `Üw` bit(1) DEFAULT NULL COMMENT 'Überweisung durch HA nötig'"
 Str(1, 51, 21) = " `Insulin` bit(1) DEFAULT NULL"
 Str(1, 51, 22) = " `ICT` bit(1) DEFAULT NULL"
 Str(1, 51, 23) = " `CSII` bit(1) DEFAULT NULL"
 Str(1, 51, 24) = " `Erst` bit(1) DEFAULT NULL"
 Str(1, 51, 25) = " `Folge` bit(1) DEFAULT NULL"
 Str(1, 51, 26) = " `fid` int(11) DEFAULT NULL"
 Str(1, 51, 27) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 51, 28) = "  KEY `Titel` (`Titel`)"
 Str(1, 51, 29) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr51

Sub FüllStr52()
 Str(0, 52, 0) = "ebmsdat"
 Str(0, 52, 1) = "`id`"
 Str(0, 52, 2) = "`dateiname`"
 Str(0, 52, 3) = "`laend`"
 Str(0, 52, 4) = "`erstellt`"
 Str(0, 52, 5) = "`von`"
 Str(0, 52, 6) = "`bis`"
 Str(0, 52, 7) = "`id`"
 Str(0, 52, 8) = "`dateiname`"
 Str(0, 52, 9) = "`erstellt`"
 Str(0, 52, 10) = "`von`"
 Str(0, 52, 11) = "`bis`"
 ArtZ(0, 52) = 6
 ArtZ(1, 52) = 5
 Str(1, 52, 0) = "CREATE TABLE `ebmsdat` ("
 Str(1, 52, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Eindeutige Datensatznummer'"
 Str(1, 52, 2) = " `dateiname` varchar(260) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Dateiname ohne Pfad'"
 Str(1, 52, 3) = " `laend` datetime DEFAULT NULL COMMENT 'Zeitpunkt der letzten Änderung'"
 Str(1, 52, 4) = " `erstellt` datetime DEFAULT NULL COMMENT 'Erstellungdatum in Turbomed'"
 Str(1, 52, 5) = " `von` datetime DEFAULT NULL COMMENT 'erster Tag der Auswertung'"
 Str(1, 52, 6) = " `bis` datetime DEFAULT NULL COMMENT 'letzter Tag der Auswertung'"
 Str(1, 52, 7) = "  PRIMARY KEY (`id`)"
 Str(1, 52, 8) = "  KEY `dateiname` (`dateiname`)"
 Str(1, 52, 9) = "  KEY `erstellt` (`erstellt`)"
 Str(1, 52, 10) = "  KEY `von` (`von`)"
 Str(1, 52, 11) = "  KEY `bis` (`bis`)"
 Str(1, 52, 12) = " ENGINE=InnoDB AUTO_INCREMENT=38 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr52

Sub FüllStr53()
 Str(0, 53, 0) = "ebmstat"
 Str(0, 53, 1) = "`id`"
 Str(0, 53, 2) = "`datid`"
 Str(0, 53, 3) = "`ziffer`"
 Str(0, 53, 4) = "`punkte`"
 Str(0, 53, 5) = "`euro`"
 Str(0, 53, 6) = "`anzahl`"
 Str(0, 53, 7) = "`id`"
 Str(0, 53, 8) = "`datid`"
 Str(0, 53, 9) = "`ziffer`"
 Str(0, 53, 10) = "`FK_ebmstat_1`"
 ArtZ(0, 53) = 6
 ArtZ(1, 53) = 3
 ArtZ(2, 53) = 1
 Str(1, 53, 0) = "CREATE TABLE `ebmstat` ("
 Str(1, 53, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Eindeutige Datensatznummer'"
 Str(1, 53, 2) = " `datid` int(10) DEFAULT NULL COMMENT 'Bezug auf `ebmsdat`'"
 Str(1, 53, 3) = " `ziffer` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'EBM-Ziffer'"
 Str(1, 53, 4) = " `punkte` int(10) DEFAULT NULL COMMENT 'angegebener Punktwert'"
 Str(1, 53, 5) = " `euro` decimal(10,2) DEFAULT NULL COMMENT 'umgerechneter Eurowert, einschl. Pauschalenumrechnung von 18,75 auf 75 Euro'"
 Str(1, 53, 6) = " `anzahl` int(10) DEFAULT NULL COMMENT 'Zahl der erbrachten Leistungen'"
 Str(1, 53, 7) = "  PRIMARY KEY (`id`)"
 Str(1, 53, 8) = "  KEY `datid` (`datid`)"
 Str(1, 53, 9) = "  KEY `ziffer` (`ziffer`)"
 Str(1, 53, 10) = "  CONSTRAINT `FK_ebmstat_1` FOREIGN KEY (`datid`) REFERENCES `ebmsdat` (`id`)"
 Str(1, 53, 11) = " ENGINE=InnoDB AUTO_INCREMENT=1710 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr53

Sub FüllStr54()
 Str(0, 54, 0) = "eingelesen"
 Str(0, 54, 1) = "`eingid`"
 Str(0, 54, 2) = "`verzeich`"
 Str(0, 54, 3) = "`datei`"
 Str(0, 54, 4) = "`fdt`"
 Str(0, 54, 5) = "`datum`"
 Str(0, 54, 6) = "`tabelle`"
 Str(0, 54, 7) = "`trennz`"
 Str(0, 54, 8) = "`feldzahl`"
 Str(0, 54, 9) = "`eingid`"
 Str(0, 54, 10) = "`vd`"
 Str(0, 54, 11) = "`fdt`"
 Str(0, 54, 12) = "`datum`"
 Str(0, 54, 13) = "`tabelle`"
 ArtZ(0, 54) = 8
 ArtZ(1, 54) = 5
 Str(1, 54, 0) = "CREATE TABLE `eingelesen` ("
 Str(1, 54, 1) = " `eingid` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 54, 2) = " `verzeich` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 54, 3) = " `datei` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 54, 4) = " `fdt` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 54, 5) = " `datum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 54, 6) = " `tabelle` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 54, 7) = " `trennz` varchar(1) COLLATE latin1_german2_ci NOT NULL COMMENT 'Trennzeichen, z.Zt. entweder "";"" oder "",""'"
 Str(1, 54, 8) = " `feldzahl` int(10) unsigned NOT NULL DEFAULT '0'"
 Str(1, 54, 9) = "  PRIMARY KEY (`eingid`)"
 Str(1, 54, 10) = "  UNIQUE KEY `vd` (`verzeich`,`datei`)"
 Str(1, 54, 11) = "  KEY `fdt` (`fdt`)"
 Str(1, 54, 12) = "  KEY `datum` (`datum`)"
 Str(1, 54, 13) = "  KEY `tabelle` (`tabelle`)"
 Str(1, 54, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='eingelesene GNR-Statistik'"
End Sub ' FüllStr54

Sub FüllStr55()
 Str(0, 55, 0) = "eingelesen1"
 Str(0, 55, 1) = "`eingid`"
 Str(0, 55, 2) = "`verzeich`"
 Str(0, 55, 3) = "`datei`"
 Str(0, 55, 4) = "`fdt`"
 Str(0, 55, 5) = "`datum`"
 Str(0, 55, 6) = "`tabelle`"
 Str(0, 55, 7) = "`trennz`"
 Str(0, 55, 8) = "`feldzahl`"
 Str(0, 55, 9) = "`eingid`"
 Str(0, 55, 10) = "`vd`"
 Str(0, 55, 11) = "`fdt`"
 Str(0, 55, 12) = "`datum`"
 Str(0, 55, 13) = "`tabelle`"
 ArtZ(0, 55) = 8
 ArtZ(1, 55) = 5
 Str(1, 55, 0) = "CREATE TABLE `eingelesen1` ("
 Str(1, 55, 1) = " `eingid` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 55, 2) = " `verzeich` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 55, 3) = " `datei` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 55, 4) = " `fdt` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 55, 5) = " `datum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 55, 6) = " `tabelle` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 55, 7) = " `trennz` varchar(1) COLLATE latin1_german2_ci NOT NULL COMMENT 'Trennzeichen, z.Zt. entweder """";"""" oder """",""""'"
 Str(1, 55, 8) = " `feldzahl` int(10) unsigned NOT NULL DEFAULT '0'"
 Str(1, 55, 9) = "  PRIMARY KEY (`eingid`)"
 Str(1, 55, 10) = "  UNIQUE KEY `vd` (`verzeich`,`datei`)"
 Str(1, 55, 11) = "  KEY `fdt` (`fdt`)"
 Str(1, 55, 12) = "  KEY `datum` (`datum`)"
 Str(1, 55, 13) = "  KEY `tabelle` (`tabelle`)"
 Str(1, 55, 14) = " ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr55

Sub FüllStr56()
 Str(0, 56, 0) = "einstellungen"
 Str(0, 56, 1) = "`ID`"
 Str(0, 56, 2) = "`Formular`"
 Str(0, 56, 3) = "`Abfrage für Formular`"
 Str(0, 56, 4) = "`ID für Formular`"
 Str(0, 56, 5) = "`DatensatzNr`"
 Str(0, 56, 6) = "`ID`"
 Str(0, 56, 7) = "`PrimaryKey`"
 Str(0, 56, 8) = "`Formular`"
 Str(0, 56, 9) = "`ID für Formular`"
 ArtZ(0, 56) = 5
 ArtZ(1, 56) = 4
 Str(1, 56, 0) = "CREATE TABLE `einstellungen` ("
 Str(1, 56, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 56, 2) = " `Formular` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name des Formulars'"
 Str(1, 56, 3) = " `Abfrage für Formular` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name der Abfrage, die zuletzt für das Formular ""Anamnesebogen"" verwendet wurde'"
 Str(1, 56, 4) = " `ID für Formular` int(10) DEFAULT NULL COMMENT 'Pat_ID in dieser Abfrage'"
 Str(1, 56, 5) = " `DatensatzNr` int(10) DEFAULT NULL COMMENT 'Datensatz-Nr. in dieser Abfrage'"
 Str(1, 56, 6) = "  PRIMARY KEY (`ID`)"
 Str(1, 56, 7) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 56, 8) = "  KEY `Formular` (`Formular`)"
 Str(1, 56, 9) = "  KEY `ID für Formular` (`ID für Formular`)"
 Str(1, 56, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr56

Sub FüllStr57()
 Str(0, 57, 0) = "eintraege"
 Str(0, 57, 1) = "`FID`"
 Str(0, 57, 2) = "`Pat_ID`"
 Str(0, 57, 3) = "`ZeitPunkt`"
 Str(0, 57, 4) = "`Art`"
 Str(0, 57, 5) = "`Inhalt`"
 Str(0, 57, 6) = "`absPos`"
 Str(0, 57, 7) = "`AktZeit`"
 Str(0, 57, 8) = "`QS`"
 Str(0, 57, 9) = "`QT`"
 Str(0, 57, 10) = "`StByte`"
 Str(0, 57, 11) = "`id`"
 Str(0, 57, 12) = "`inhNum`"
 Str(0, 57, 13) = "`id`"
 Str(0, 57, 14) = "`id_UNIQUE`"
 Str(0, 57, 15) = "`Auswahl`"
 Str(0, 57, 16) = "`Art`"
 Str(0, 57, 17) = "`FälleEinträge`"
 Str(0, 57, 18) = "`Zeitpunkt`"
 Str(0, 57, 19) = "`F??lleEintr??ge_AccRel`"
 ArtZ(0, 57) = 12
 ArtZ(1, 57) = 6
 ArtZ(2, 57) = 1
 Str(1, 57, 0) = "CREATE TABLE `eintraege` ("
 Str(1, 57, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 57, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 57, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 57, 4) = " `Art` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6330'"
 Str(1, 57, 5) = " `Inhalt` varchar(5000) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8480'"
 Str(1, 57, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 57, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 57, 8) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 57, 9) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 57, 10) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnr. der Datenübertragung'"
 Str(1, 57, 11) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 57, 12) = " `inhNum` double NOT NULL COMMENT 'Inhalt numerisch'"
 Str(1, 57, 13) = "  PRIMARY KEY (`id`)"
 Str(1, 57, 14) = "  UNIQUE KEY `id_UNIQUE` (`id`)"
 Str(1, 57, 15) = "  KEY `Auswahl` (`Pat_ID`,`Art`,`ZeitPunkt`)"
 Str(1, 57, 16) = "  KEY `Art` (`Art`)"
 Str(1, 57, 17) = "  KEY `FälleEinträge` (`FID`,`Inhalt`(767)) USING BTREE"
 Str(1, 57, 18) = "  KEY `Zeitpunkt` (`ZeitPunkt`,`Art`) USING BTREE"
 Str(1, 57, 19) = "  CONSTRAINT `F??lleEintr??ge_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 57, 20) = " ENGINE=InnoDB AUTO_INCREMENT=45801968 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=FIXED"
End Sub ' FüllStr57

Sub FüllStr58()
 Str(0, 58, 0) = "eintragszahlen"
 Str(0, 58, 1) = "`Beginn`"
 Str(0, 58, 2) = "`StByte`"
 Str(0, 58, 3) = "`Zp1`"
 Str(0, 58, 4) = "`Zp2`"
 Str(0, 58, 5) = "`Zp3`"
 Str(0, 58, 6) = "`Zp4`"
 Str(0, 58, 7) = "`Zp5`"
 Str(0, 58, 8) = "`Zp6`"
 Str(0, 58, 9) = "`Zp7`"
 Str(0, 58, 10) = "`Zp8`"
 Str(0, 58, 11) = "`Fallzahl`"
 Str(0, 58, 12) = "`Sekunden`"
 Str(0, 58, 13) = "`Datei`"
 Str(0, 58, 14) = "`DateiAend`"
 Str(0, 58, 15) = "`SpeicherZt`"
 Str(0, 58, 16) = "`TabellenEntleeren`"
 Str(0, 58, 17) = "`ZurücksetzenLAktDat`"
 Str(0, 58, 18) = "`Pat_IDVon`"
 Str(0, 58, 19) = "`Pat_IDbis`"
 Str(0, 58, 20) = "`VorladenFFI`"
 Str(0, 58, 21) = "`ÜberTabelle`"
 Str(0, 58, 22) = "`SammelInsert`"
 Str(0, 58, 23) = "`bereinigeFormInhFeld`"
 Str(0, 58, 24) = "`LaborDirektEinlesen`"
 Str(0, 58, 25) = "`LaborDirektNeu`"
 Str(0, 58, 26) = "`LaborQuerVerb`"
 Str(0, 58, 27) = "`LaborQuerNeu`"
 Str(0, 58, 28) = "`AlterTab`"
 Str(0, 58, 29) = "`obmitEmails`"
 Str(0, 58, 30) = "`LaborPfadBeispiel`"
 Str(0, 58, 31) = "`obVglMitLetzterEinlesung`"
 Str(0, 58, 32) = "`NurInTabelle`"
 Str(0, 58, 33) = "`Beginn`"
 Str(0, 58, 34) = "`Beginn`"
 Str(0, 58, 35) = "`stbyte`"
 ArtZ(0, 58) = 32
 ArtZ(1, 58) = 3
 Str(1, 58, 0) = "CREATE TABLE `eintragszahlen` ("
 Str(1, 58, 1) = " `Beginn` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 58, 2) = " `StByte` int(10) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 58, 3) = " `Zp1` datetime DEFAULT NULL"
 Str(1, 58, 4) = " `Zp2` datetime DEFAULT NULL"
 Str(1, 58, 5) = " `Zp3` datetime DEFAULT NULL"
 Str(1, 58, 6) = " `Zp4` datetime DEFAULT NULL"
 Str(1, 58, 7) = " `Zp5` datetime DEFAULT NULL"
 Str(1, 58, 8) = " `Zp6` datetime DEFAULT NULL"
 Str(1, 58, 9) = " `Zp7` datetime DEFAULT NULL"
 Str(1, 58, 10) = " `Zp8` datetime DEFAULT NULL"
 Str(1, 58, 11) = " `Fallzahl` int(10) DEFAULT NULL"
 Str(1, 58, 12) = " `Sekunden` int(10) DEFAULT NULL"
 Str(1, 58, 13) = " `Datei` varchar(120) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 58, 14) = " `DateiAend` datetime DEFAULT NULL"
 Str(1, 58, 15) = " `SpeicherZt` datetime DEFAULT NULL"
 Str(1, 58, 16) = " `TabellenEntleeren` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 17) = " `ZurücksetzenLAktDat` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 18) = " `Pat_IDVon` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 58, 19) = " `Pat_IDbis` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 58, 20) = " `VorladenFFI` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 21) = " `ÜberTabelle` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 22) = " `SammelInsert` tinyint(1) unsigned NOT NULL DEFAULT '0'"
 Str(1, 58, 23) = " `bereinigeFormInhFeld` tinyint(1) unsigned NOT NULL COMMENT 'ob FormInhFeld bereinigt wird'"
 Str(1, 58, 24) = " `LaborDirektEinlesen` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 25) = " `LaborDirektNeu` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 26) = " `LaborQuerVerb` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 27) = " `LaborQuerNeu` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 28) = " `AlterTab` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 29) = " `obmitEmails` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 30) = " `LaborPfadBeispiel` longtext COLLATE latin1_german2_ci"
 Str(1, 58, 31) = " `obVglMitLetzterEinlesung` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 58, 32) = " `NurInTabelle` bit(1) DEFAULT NULL"
 Str(1, 58, 33) = "  PRIMARY KEY (`Beginn`)"
 Str(1, 58, 34) = "  UNIQUE KEY `Beginn` (`Beginn`)"
 Str(1, 58, 35) = "  UNIQUE KEY `stbyte` (`StByte`) USING BTREE"
 Str(1, 58, 36) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr58

Sub FüllStr59()
 Str(0, 59, 0) = "eintragzulanr"
 Str(0, 59, 1) = "`id`"
 Str(0, 59, 2) = "`lanrid`"
 Str(0, 59, 3) = "`art`"
 Str(0, 59, 4) = "`id`"
 Str(0, 59, 5) = "`FK_eintragzulanr_lanr`"
 Str(0, 59, 6) = "`FK_eintragzulanr_lanr`"
 ArtZ(0, 59) = 3
 ArtZ(1, 59) = 2
 ArtZ(2, 59) = 1
 Str(1, 59, 0) = "CREATE TABLE `eintragzulanr` ("
 Str(1, 59, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 59, 2) = " `lanrid` int(10) unsigned NOT NULL DEFAULT '0'"
 Str(1, 59, 3) = " `art` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 59, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 59, 5) = "  KEY `FK_eintragzulanr_lanr` (`lanrid`)"
 Str(1, 59, 6) = "  CONSTRAINT `FK_eintragzulanr_lanr` FOREIGN KEY (`lanrid`) REFERENCES `lanrpraxis` (`id`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 59, 7) = " ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Einträge, die auf bestimmten Arzt deuten'"
End Sub ' FüllStr59

Sub FüllStr60()
 Str(0, 60, 0) = "eintrhist"
 Str(1, 60, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `eintrhist` AS SELECT `eintrhist1`.`ID` AS `ID`,`eintrhist1`.`Pat_ID` AS `Pat_ID`,`eintrhist1`.`ZeitPunkt` AS `ZeitPunkt`,`eintrhist1`.`Art` AS `Art`,`eintrhist1`.`Inhalt` AS `Inhalt`,`eintrhist1`.`QS` AS `QS`,`eintrhist1`.`QT` AS `QT`,`eintrhist2`.`FID` AS `FID`,`eintrhist2`.`absPos` AS `absPos`,`eintrhist2`.`AktZeit` AS `AktZeit`,`eintrhist2`.`StByte` AS `StByte` FROM (`eintrhist1` join `eintrhist2` on((`eintrhist1`.`ID` = `eintrhist2`.`ID`)))"
End Sub ' FüllStr60

Sub FüllStr61()
 Str(0, 61, 0) = "eintrhist1"
 Str(0, 61, 1) = "`Pat_ID`"
 Str(0, 61, 2) = "`ZeitPunkt`"
 Str(0, 61, 3) = "`Art`"
 Str(0, 61, 4) = "`Inhalt`"
 Str(0, 61, 5) = "`QS`"
 Str(0, 61, 6) = "`QT`"
 Str(0, 61, 7) = "`ID`"
 Str(0, 61, 8) = "`ID`"
 Str(0, 61, 9) = "`Auswahl`"
 Str(0, 61, 10) = "`NamenEinträge`"
 Str(0, 61, 11) = "`Art`"
 ArtZ(0, 61) = 7
 ArtZ(1, 61) = 4
 Str(1, 61, 0) = "CREATE TABLE `eintrhist1` ("
 Str(1, 61, 1) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 61, 2) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 61, 3) = " `Art` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6330'"
 Str(1, 61, 4) = " `Inhalt` longtext COLLATE latin1_german2_ci COMMENT '8480'"
 Str(1, 61, 5) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 61, 6) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 61, 7) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 61, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 61, 9) = "  KEY `Auswahl` (`Pat_ID`,`Art`,`ZeitPunkt`)"
 Str(1, 61, 10) = "  KEY `NamenEinträge` (`Pat_ID`)"
 Str(1, 61, 11) = "  KEY `Art` (`Art`)"
 Str(1, 61, 12) = " ENGINE=InnoDB AUTO_INCREMENT=16129 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr61

Sub FüllStr62()
 Str(0, 62, 0) = "eintrhist2"
 Str(0, 62, 1) = "`FID`"
 Str(0, 62, 2) = "`absPos`"
 Str(0, 62, 3) = "`AktZeit`"
 Str(0, 62, 4) = "`StByte`"
 Str(0, 62, 5) = "`ID`"
 Str(0, 62, 6) = "`ideig`"
 Str(0, 62, 7) = "`ideig`"
 Str(0, 62, 8) = "`FälleEinträge`"
 Str(0, 62, 9) = "`FK_id`"
 ArtZ(0, 62) = 6
 ArtZ(1, 62) = 3
 Str(1, 62, 0) = "CREATE TABLE `eintrhist2` ("
 Str(1, 62, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 62, 2) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 62, 3) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 62, 4) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 62, 5) = " `ID` int(10) unsigned NOT NULL DEFAULT '0' COMMENT 'Bezug auf eintrhist1'"
 Str(1, 62, 6) = " `ideig` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'eigene ID'"
 Str(1, 62, 7) = "  UNIQUE KEY `ideig` (`ideig`)"
 Str(1, 62, 8) = "  KEY `FälleEinträge` (`FID`)"
 Str(1, 62, 9) = "  KEY `FK_id` (`ID`)"
 Str(1, 62, 10) = " ENGINE=InnoDB AUTO_INCREMENT=737397 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr62

Sub FüllStr63()
 Str(0, 63, 0) = "faelle"
 Str(0, 63, 1) = "`FID`"
 Str(0, 63, 2) = "`Pat_ID`"
 Str(0, 63, 3) = "`Quartal`"
 Str(0, 63, 4) = "`Nachname`"
 Str(0, 63, 5) = "`Vorname`"
 Str(0, 63, 6) = "`lfdnr`"
 Str(0, 63, 7) = "`TMFNr`"
 Str(0, 63, 8) = "`VKNr`"
 Str(0, 63, 9) = "`BhFB`"
 Str(0, 63, 10) = "`BhFE1`"
 Str(0, 63, 11) = "`BhFE2`"
 Str(0, 63, 12) = "`f4202`"
 Str(0, 63, 13) = "`ausgst`"
 Str(0, 63, 14) = "`KtrAbrB`"
 Str(0, 63, 15) = "`AbrAr`"
 Str(0, 63, 16) = "`lVorl`"
 Str(0, 63, 17) = "`IK`"
 Str(0, 63, 18) = "`KVKs`"
 Str(0, 63, 19) = "`KVKserg`"
 Str(0, 63, 20) = "`Kasse`"
 Str(0, 63, 21) = "`GebOr`"
 Str(0, 63, 22) = "`AbrGb`"
 Str(0, 63, 23) = "`PersKreis`"
 Str(0, 63, 24) = "`SKtZusatz`"
 Str(0, 63, 25) = "`letzteRegel`"
 Str(0, 63, 26) = "`ÜwText`"
 Str(0, 63, 27) = "`f4210`"
 Str(0, 63, 28) = "`AkfHAH`"
 Str(0, 63, 29) = "`AkfAB0`"
 Str(0, 63, 30) = "`AkfAK`"
 Str(0, 63, 31) = "`statNuller`"
 Str(0, 63, 32) = "`ÜbwV`"
 Str(0, 63, 33) = "`ÜbWVLANR`"
 Str(0, 63, 34) = "`ÜbWVBSNR`"
 Str(0, 63, 35) = "`ÜbWVKVNR`"
 Str(0, 63, 36) = "`AndÜw`"
 Str(0, 63, 37) = "`Übwr`"
 Str(0, 63, 38) = "`ÜbwLANR`"
 Str(0, 63, 39) = "`ÜWZiel`"
 Str(0, 63, 40) = "`ÜWNNr`"
 Str(0, 63, 41) = "`ÜWNaN`"
 Str(0, 63, 42) = "`ÜWTit`"
 Str(0, 63, 43) = "`ÜWVor`"
 Str(0, 63, 44) = "`ÜWVsw`"
 Str(0, 63, 45) = "`üwvid`"
 Str(0, 63, 46) = "`Auftrag`"
 Str(0, 63, 47) = "`Verdacht`"
 Str(0, 63, 48) = "`Befund`"
 Str(0, 63, 49) = "`statKlasse`"
 Str(0, 63, 50) = "`f4237`"
 Str(0, 63, 51) = "`statBehTage`"
 Str(0, 63, 52) = "`SchGr`"
 Str(0, 63, 53) = "`Weiterbeh`"
 Str(0, 63, 54) = "`PGeb`"
 Str(0, 63, 55) = "`PGebErg`"
 Str(0, 63, 56) = "`Mahnfrist`"
 Str(0, 63, 57) = "`GOÄKatNr`"
 Str(0, 63, 58) = "`GOÄKatName`"
 Str(0, 63, 59) = "`abrArzt`"
 Str(0, 63, 60) = "`privVers`"
 Str(0, 63, 61) = "`AdNam`"
 Str(0, 63, 62) = "`AdStr`"
 Str(0, 63, 63) = "`AdPlz`"
 Str(0, 63, 64) = "`AdOrt`"
 Str(0, 63, 65) = "`BhFE`"
 Str(0, 63, 66) = "`s8000`"
 Str(0, 63, 67) = "`s8100`"
 Str(0, 63, 68) = "`AktZeit`"
 Str(0, 63, 69) = "`Fanf`"
 Str(0, 63, 70) = "`altQuart`"
 Str(0, 63, 71) = "`QAnf`"
 Str(0, 63, 72) = "`QEnd`"
 Str(0, 63, 73) = "`QS`"
 Str(0, 63, 74) = "`QT`"
 Str(0, 63, 75) = "`StByte`"
 Str(0, 63, 76) = "`absPos`"
 Str(0, 63, 77) = "`LANRid`"
 Str(0, 63, 78) = "`f4108`"
 Str(0, 63, 79) = "`BGFallNr`"
 Str(0, 63, 80) = "`lGewicht`"
 Str(0, 63, 81) = "`FID`"
 Str(0, 63, 82) = "`PrimaryKey`"
 Str(0, 63, 83) = "`AktF`"
 Str(0, 63, 84) = "`Auswahl`"
 Str(0, 63, 85) = "`BhFB`"
 Str(0, 63, 86) = "`FanF`"
 Str(0, 63, 87) = "`NamenFälle`"
 Str(0, 63, 88) = "`pQ`"
 Str(0, 63, 89) = "`Quartal`"
 Str(0, 63, 90) = "`SchGr`"
 Str(0, 63, 91) = "`vknr`"
 Str(0, 63, 92) = "`KassenlisteF??lle_AccRel`"
 Str(0, 63, 93) = "`NamenF??lle_AccRel`"
 ArtZ(0, 63) = 80
 ArtZ(1, 63) = 11
 ArtZ(2, 63) = 2
 Str(1, 63, 0) = "CREATE TABLE `faelle` ("
 Str(1, 63, 1) = " `FID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 63, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 63, 3) = " `Quartal` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4101'"
 Str(1, 63, 4) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 63, 5) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 63, 6) = " `lfdnr` int(10) DEFAULT NULL COMMENT 'laufende Fallnummer'"
 Str(1, 63, 7) = " `TMFNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4144 Fallnummer in Turbomed'"
 Str(1, 63, 8) = " `VKNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4104'"
 Str(1, 63, 9) = " `BhFB` datetime DEFAULT NULL COMMENT '4150'"
 Str(1, 63, 10) = " `BhFE1` datetime DEFAULT NULL COMMENT '4151'"
 Str(1, 63, 11) = " `BhFE2` datetime DEFAULT NULL COMMENT '4152'"
 Str(1, 63, 12) = " `f4202` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4202'"
 Str(1, 63, 13) = " `ausgst` datetime DEFAULT NULL COMMENT '4102 (''ausgestellt am'')'"
 Str(1, 63, 14) = " `KtrAbrB` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))'"
 Str(1, 63, 15) = " `AbrAr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4107, Abrechnungsart (1 = Primärkassen)'"
 Str(1, 63, 16) = " `lVorl` datetime DEFAULT NULL COMMENT '4109, letzte Vorlage'"
 Str(1, 63, 17) = " `IK` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4111 Krankenkassennummer (IK)'"
 Str(1, 63, 18) = " `KVKs` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4112 Versichertenstatus VK'"
 Str(1, 63, 19) = " `KVKserg` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4113 Ost/West-Status VK'"
 Str(1, 63, 20) = " `Kasse` varchar(70) COLLATE latin1_german2_ci NOT NULL COMMENT '6299 Kasse (aus Formularen)'"
 Str(1, 63, 21) = " `GebOr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4121, Gebührenordnung (1 = BMÄ, 2)'"
 Str(1, 63, 22) = " `AbrGb` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4122, Abrechnungsgebiet (07 = Diabetes)'"
 Str(1, 63, 23) = " `PersKreis` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4123 Personenkreis/Untersuchungskategorie'"
 Str(1, 63, 24) = " `SKtZusatz` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4124 SKT-Zusatzangaben'"
 Str(1, 63, 25) = " `letzteRegel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4206, letzter Tag der Regel'"
 Str(1, 63, 26) = " `ÜwText` longtext COLLATE latin1_german2_ci COMMENT '4209: Auftrags- / erläuternder Text zur Überweisung'"
 Str(1, 63, 27) = " `f4210` tinyint(1) unsigned DEFAULT NULL COMMENT '4210, Ankreuzfeld LSR'"
 Str(1, 63, 28) = " `AkfHAH` tinyint(1) unsigned DEFAULT NULL COMMENT '4211 Ankreuzfeld HAH'"
 Str(1, 63, 29) = " `AkfAB0` tinyint(1) unsigned DEFAULT NULL COMMENT '4212 Ankreuzfeld AB0.RH'"
 Str(1, 63, 30) = " `AkfAK` tinyint(1) unsigned DEFAULT NULL COMMENT '4213 Ankreuzfeld AK'"
 Str(1, 63, 31) = " `statNuller` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4216, nu bei Musterfrau 16 Nuller'"
 Str(1, 63, 32) = " `ÜbwV` varchar(22) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4218, überwiesen von Arztnummer'"
 Str(1, 63, 33) = " `ÜbWVLANR` varchar(9) COLLATE latin1_german2_ci NOT NULL COMMENT '4218(1) überwiesen von LANR'"
 Str(1, 63, 34) = " `ÜbWVBSNR` varchar(11) COLLATE latin1_german2_ci NOT NULL COMMENT '4218(2) überwiesen von BSNR'"
 Str(1, 63, 35) = " `ÜbWVKVNR` varchar(9) COLLATE latin1_german2_ci NOT NULL COMMENT '4218(3) überwiesen von KVNR'"
 Str(1, 63, 36) = " `AndÜw` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4219, anderer Überweiser'"
 Str(1, 63, 37) = " `Übwr` varchar(9) COLLATE latin1_german2_ci NOT NULL COMMENT 'resultierender Überweiser (BSNR): 4218 oder 4219, je nachdem, was befüllt'"
 Str(1, 63, 38) = " `ÜbwLANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4242 LANR des Überweisers'"
 Str(1, 63, 39) = " `ÜWZiel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4220 Überweisung an'"
 Str(1, 63, 40) = " `ÜWNNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(4): KV-Nummer des Überweisers'"
 Str(1, 63, 41) = " `ÜWNaN` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(3): Nachname des Überweisers'"
 Str(1, 63, 42) = " `ÜWTit` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(3): Titel des Überweisers'"
 Str(1, 63, 43) = " `ÜWVor` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(2): Vorname des Überweisers'"
 Str(1, 63, 44) = " `ÜWVsw` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(2b): Vorsatzwort des Überweisers'"
 Str(1, 63, 45) = " `üwvid` int(10) unsigned NOT NULL COMMENT '4247 Bezug auf ueberwvon'"
 Str(1, 63, 46) = " `Auftrag` varchar(195) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4205 Auftrag bei Überweisung'"
 Str(1, 63, 47) = " `Verdacht` varchar(133) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4207 Verdacht bei Überweisung\n'"
 Str(1, 63, 48) = " `Befund` varchar(210) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4208 Befund bei Überweisung'"
 Str(1, 63, 49) = " `statKlasse` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4236 Klasse bei Behandlung'"
 Str(1, 63, 50) = " `f4237` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4237 Krankenhausname'"
 Str(1, 63, 51) = " `statBehTage` int(10) DEFAULT NULL COMMENT '4238 Krankenhausaufenthalt'"
 Str(1, 63, 52) = " `SchGr` decimal(2,0) DEFAULT NULL COMMENT '4239, Schein(unter)gruppe'"
 Str(1, 63, 53) = " `Weiterbeh` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4243, Weiterbehandelnder'"
 Str(1, 63, 54) = " `PGeb` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4401, Praxisgebühr'"
 Str(1, 63, 55) = " `PGebErg` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4402, ?'"
 Str(1, 63, 56) = " `Mahnfrist` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4403, Mahnfrist bis'"
 Str(1, 63, 57) = " `GOÄKatNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4580 (1): Katalog-Nummer'"
 Str(1, 63, 58) = " `GOÄKatName` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4580 (2): Privat-Abrechnungskatalog'"
 Str(1, 63, 59) = " `abrArzt` varchar(30) COLLATE latin1_german2_ci NOT NULL COMMENT '4585 abrechnender Arzt'"
 Str(1, 63, 60) = " `privVers` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4586 private Versicherung'"
 Str(1, 63, 61) = " `AdNam` varchar(28) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT '4602(1) Name Rechnungsanschrift'"
 Str(1, 63, 62) = " `AdStr` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(2) Straße Rechnungsanschrift'"
 Str(1, 63, 63) = " `AdPlz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(3) PLZ Rechnungsanschrift'"
 Str(1, 63, 64) = " `AdOrt` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(4) Ort Rechnungsanschrift'"
 Str(1, 63, 65) = " `BhFE` datetime DEFAULT NULL COMMENT '4604, Behandlungsfall: Ende, bei Privatpatienten'"
 Str(1, 63, 66) = " `s8000` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000, Satzidentifikation'"
 Str(1, 63, 67) = " `s8100` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge'"
 Str(1, 63, 68) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 63, 69) = " `Fanf` datetime DEFAULT NULL COMMENT 'Fallanfang'"
 Str(1, 63, 70) = " `altQuart` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 63, 71) = " `QAnf` datetime DEFAULT NULL COMMENT 'Quartalsanfang'"
 Str(1, 63, 72) = " `QEnd` datetime DEFAULT NULL COMMENT 'Quartalsende'"
 Str(1, 63, 73) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 63, 74) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 63, 75) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 63, 76) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 63, 77) = " `LANRid` int(3) unsigned NOT NULL COMMENT 'Bezug auf lanrpraxis.id'"
 Str(1, 63, 78) = " `f4108` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4108'"
 Str(1, 63, 79) = " `BGFallNr` varchar(12) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3603 BG-Fall-Nummer  '"
 Str(1, 63, 80) = " `lGewicht` decimal(5,1) NOT NULL COMMENT 'letztes Gewicht in kg'"
 Str(1, 63, 81) = "  PRIMARY KEY (`FID`)"
 Str(1, 63, 82) = "  UNIQUE KEY `PrimaryKey` (`FID`)"
 Str(1, 63, 83) = "  KEY `AktF` (`Pat_ID`,`BhFB`)"
 Str(1, 63, 84) = "  KEY `Auswahl` (`Pat_ID`,`Quartal`,`BhFB`,`BhFE1`)"
 Str(1, 63, 85) = "  KEY `BhFB` (`BhFB`)"
 Str(1, 63, 86) = "  KEY `FanF` (`Fanf`)"
 Str(1, 63, 87) = "  KEY `NamenFälle` (`Pat_ID`)"
 Str(1, 63, 88) = "  KEY `pQ` (`Pat_ID`,`Quartal`)"
 Str(1, 63, 89) = "  KEY `Quartal` (`Quartal`)"
 Str(1, 63, 90) = "  KEY `SchGr` (`SchGr`,`Nachname`,`Vorname`)"
 Str(1, 63, 91) = "  KEY `vknr` (`VKNr`)"
 Str(1, 63, 92) = "  CONSTRAINT `KassenlisteF??lle_AccRel` FOREIGN KEY (`VKNr`) REFERENCES `kassenliste` (`VK`)"
 Str(1, 63, 93) = "  CONSTRAINT `NamenF??lle_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 63, 94) = " ENGINE=InnoDB AUTO_INCREMENT=3182741 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr63

Sub FüllStr64()
 Str(0, 64, 0) = "faelleverschieden"
 Str(1, 64, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `faelleverschieden` AS SELECT `f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`letzteRegel` AS `letzteRegel`,`f`.`ÜwText` AS `ÜwText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`ÜbwV` " & _
  "AS `ÜbwV`,`f`.`ÜbWVLANR` AS `ÜbWVLANR`,`f`.`ÜbWVBSNR` AS `ÜbWVBSNR`,`f`.`ÜbWVKVNR` AS `ÜbWVKVNR`,`f`.`AndÜw` AS `AndÜw`,`f`.`Übwr` AS `Übwr`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`Auftrag` AS `Auftrag`,`f`.`Verdacht` AS `Verdacht`,`f`.`Befund` AS `Befund`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8100" & _
  "` AS `s8100`,`f`.`AktZeit` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos`,`f`.`LANRid` AS `LANRid`,`f`.`f4108` AS `f4108`,`f`.`BGFallNr` AS `BGFallNr`,`f`.`lGewicht` AS `lGewicht` FROM `_faellenachschgr` `f` group by `f`.`Pat_ID`,`f`.`Quartal`"
End Sub ' FüllStr64

Sub FüllStr65()
 Str(0, 65, 0) = "faelleverschiedenneu"
 Str(1, 65, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `faelleverschiedenneu` AS SELECT `f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`letzteRegel` AS `letzteRegel`,`f`.`ÜwText` AS `ÜwText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`Übw" & _
  "V` AS `ÜbwV`,`f`.`ÜbWVLANR` AS `ÜbWVLANR`,`f`.`ÜbWVBSNR` AS `ÜbWVBSNR`,`f`.`ÜbWVKVNR` AS `ÜbWVKVNR`,`f`.`AndÜw` AS `AndÜw`,`f`.`Übwr` AS `Übwr`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`Auftrag` AS `Auftrag`,`f`.`Verdacht` AS `Verdacht`,`f`.`Befund` AS `Befund`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8" & _
  "100` AS `s8100`,`f`.`AktZeit` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos`,`f`.`LANRid` AS `LANRid`,`f`.`f4108` AS `f4108`,`f`.`BGFallNr` AS `BGFallNr`,`f`.`lGewicht` AS `lGewicht`,(SELECT min(`f1`.`Fanf`) FROM `faelle` `f1` WHERE ((`f1`.`Pat_ID` = `f`.`Pat_ID`) and (`f1`.`Fanf` < `f`.`Fanf`))) AS `erst` FROM `_faellenachschgr` `f` group by `f`.`Pat_ID`,`f`.`Quartal`"
End Sub ' FüllStr65

Sub FüllStr66()
 Str(0, 66, 0) = "fallzahlstand"
 Str(0, 66, 1) = "`id`"
 Str(0, 66, 2) = "`Quartal`"
 Str(0, 66, 3) = "`Tage`"
 Str(0, 66, 4) = "`ArbT`"
 Str(0, 66, 5) = "`ArbTSchade`"
 Str(0, 66, 6) = "`ArbTKothny`"
 Str(0, 66, 7) = "`Kassenpat`"
 Str(0, 66, 8) = "`KassenpatRel`"
 Str(0, 66, 9) = "`KassenpatSchade`"
 Str(0, 66, 10) = "`KassenpatSchadeRel`"
 Str(0, 66, 11) = "`KassenpatKothny`"
 Str(0, 66, 12) = "`KassenpatKothnyRel`"
 Str(0, 66, 13) = "`DmRel`"
 Str(0, 66, 14) = "`GDM`"
 Str(0, 66, 15) = "`Neue`"
 Str(0, 66, 16) = "`NeueDmRel`"
 Str(0, 66, 17) = "`NeueSchade`"
 Str(0, 66, 18) = "`NeueKothny`"
 Str(0, 66, 19) = "`Doppler`"
 Str(0, 66, 20) = "`Duplex`"
 Str(0, 66, 21) = "`Sonos`"
 Str(0, 66, 22) = "`Schul`"
 Str(0, 66, 23) = "`BriefeSchade`"
 Str(0, 66, 24) = "`BriefeKothny`"
 Str(0, 66, 25) = "`DmBriefeSchade`"
 Str(0, 66, 26) = "`DmBriefeKothny`"
 Str(0, 66, 27) = "`aktzeit`"
 Str(0, 66, 28) = "`id`"
 Str(0, 66, 29) = "`Tage_Quartal`"
 ArtZ(0, 66) = 27
 ArtZ(1, 66) = 2
 Str(1, 66, 0) = "CREATE TABLE `fallzahlstand` ("
 Str(1, 66, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'fortlaufende Nummer'"
 Str(1, 66, 2) = " `Quartal` varchar(5) COLLATE latin1_german2_ci NOT NULL COMMENT 'Quartal [1-4]Jahr'"
 Str(1, 66, 3) = " `Tage` int(3) unsigned NOT NULL COMMENT 'Tage nach jeweiligem Quartalsbeginn'"
 Str(1, 66, 4) = " `ArbT` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Arbeitstage'"
 Str(1, 66, 5) = " `ArbTSchade` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Arbeitstage Schade'"
 Str(1, 66, 6) = " `ArbTKothny` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Arbeitstage Kothny'"
 Str(1, 66, 7) = " `Kassenpat` int(4) unsigned NOT NULL COMMENT 'Zahl der Kassenpatienten bis (Tage) nach Quartalsbeginn'"
 Str(1, 66, 8) = " `KassenpatRel` decimal(6,1) DEFAULT NULL COMMENT 'Steigerung bzw. Senkung der Zahl der Kassenpat.im Vgl.z.Vorquartal'"
 Str(1, 66, 9) = " `KassenpatSchade` int(4) unsigned NOT NULL COMMENT 'Zahl der Kassenpat.v.G.Schade'"
 Str(1, 66, 10) = " `KassenpatSchadeRel` decimal(4,1) NOT NULL COMMENT 'Anteil der Pat.v.G.Schade an allen Kassenpat.'"
 Str(1, 66, 11) = " `KassenpatKothny` int(4) unsigned NOT NULL COMMENT 'Zahl der Kassenpat.v.Dr.Kothny'"
 Str(1, 66, 12) = " `KassenpatKothnyRel` decimal(4,1) NOT NULL COMMENT 'Anteil der Pat.v.Dr.Kothny an allen Kassenpat.'"
 Str(1, 66, 13) = " `DmRel` decimal(4,1) DEFAULT NULL COMMENT 'Anteil der Diabetiker an allen Patienten'"
 Str(1, 66, 14) = " `GDM` int(3) unsigned NOT NULL COMMENT 'Zahl der Gestationsdiabetikerinnen'"
 Str(1, 66, 15) = " `Neue` int(3) unsigned NOT NULL COMMENT 'Zahl der neuen Patienten'"
 Str(1, 66, 16) = " `NeueDmRel` decimal(4,1) NOT NULL COMMENT 'Anteil der Diabetiker unter den Neuen'"
 Str(1, 66, 17) = " `NeueSchade` int(3) unsigned NOT NULL COMMENT 'Zahl der neuen Pat.bei G.Schade'"
 Str(1, 66, 18) = " `NeueKothny` int(3) unsigned NOT NULL COMMENT 'Zahl der neuen Pat.bei Dr.Kothny'"
 Str(1, 66, 19) = " `Doppler` int(3) unsigned NOT NULL COMMENT 'Zahl der Doppleruntersuchungen'"
 Str(1, 66, 20) = " `Duplex` int(3) unsigned NOT NULL COMMENT 'Zahl der Duplexuntersuchungen'"
 Str(1, 66, 21) = " `Sonos` int(3) unsigned NOT NULL COMMENT 'Zahl der Sonos'"
 Str(1, 66, 22) = " `Schul` int(4) unsigned NOT NULL COMMENT 'Zahl der Schulungen'"
 Str(1, 66, 23) = " `BriefeSchade` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Zahl der geschriebenen Arztbriefe Schade'"
 Str(1, 66, 24) = " `BriefeKothny` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Zahl der geschriebenen Arztbriefe Kothny'"
 Str(1, 66, 25) = " `DmBriefeSchade` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Zahl der Briefe bei Diabetikern Schade'"
 Str(1, 66, 26) = " `DmBriefeKothny` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Zahl der Briefe bei Diabetikern Kothny'"
 Str(1, 66, 27) = " `aktzeit` datetime NOT NULL COMMENT 'Eintragszeitpunkt des Datensatzes'"
 Str(1, 66, 28) = "  PRIMARY KEY (`id`)"
 Str(1, 66, 29) = "  KEY `Tage_Quartal` (`Tage`,`Quartal`)"
 Str(1, 66, 30) = " ENGINE=InnoDB AUTO_INCREMENT=9909 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=FIXED COMMENT='Datenspeicher für Fallzahlstand-Statistik'"
End Sub ' FüllStr66

Sub FüllStr67()
 Str(0, 67, 0) = "fallzahlstand 1"
 Str(1, 67, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 1` AS SELECT count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` FROM `faelleverschiedenneu` `f` WHERE ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 0 day)) - to_days(concat(year((now() - interval 0 day)),'-',((((month((now() - interval 0 day)) - 1) DIV 3) * 3) + 1),'-01')))) and ((`f`.`Pat_ID` < 3044) or (`f`.`Pat_ID` > 50000))) group by `f`.`Quartal` ORDER BY substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Sub ' FüllStr67

Sub FüllStr68()
 Str(0, 68, 0) = "fallzahlstand 2"
 Str(1, 68, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 2` AS SELECT count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` FROM `faelleverschiedenneu` `f` WHERE ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 0 day)) - to_days(concat(year((now() - interval 0 day)),'-',((((month((now() - interval 0 day)) - 1) DIV 3) * 3) + 1),'-01')))) and (`f`.`Pat_ID` > 3044)) group by `f`.`Quartal` ORDER BY substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Sub ' FüllStr68

Sub FüllStr69()
 Str(0, 69, 0) = "fallzahlstand 3"
 Str(1, 69, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 3` AS SELECT count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` FROM `faelleverschiedenneu` `f` WHERE ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 0 day)) - to_days(concat(year((now() - interval 0 day)),'-',((((month((now() - interval 0 day)) - 1) DIV 3) * 3) + 1),'-01')))) and 1 and (substr(`f`.`Quartal`,2,4) > '2008') and (`f`.`Quartal` <> '12009')) group by `f`.`Quartal` ORDER BY substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Sub ' FüllStr69

Sub FüllStr70()
 Str(0, 70, 0) = "faxe"
 Str(0, 70, 1) = "`Name`"
 Str(0, 70, 2) = "`erstellt`"
 Str(0, 70, 3) = "`geändert`"
 Str(0, 70, 4) = "`Größe`"
 Str(0, 70, 5) = "`NameInP`"
 Str(0, 70, 6) = "`Telefonnummer`"
 Str(0, 70, 7) = "`Absender`"
 Str(0, 70, 8) = "`kopiert`"
 Str(0, 70, 9) = "`altNameInP`"
 Str(0, 70, 10) = "`ID`"
 Str(0, 70, 11) = "`TMStart`"
 Str(0, 70, 12) = "`TMEnd`"
 Str(0, 70, 13) = "`fFNr`"
 Str(0, 70, 14) = "`ID`"
 Str(0, 70, 15) = "`kopiertgeändertGröße`"
 Str(0, 70, 16) = "`kopiertName`"
 Str(0, 70, 17) = "`S2`"
 Str(0, 70, 18) = "`Telefonnummer`"
 Str(0, 70, 19) = "`geändertGröße`"
 Str(0, 70, 20) = "`Name`"
 Str(0, 70, 21) = "`NameinP`"
 ArtZ(0, 70) = 13
 ArtZ(1, 70) = 8
 Str(1, 70, 0) = "CREATE TABLE `faxe` ("
 Str(1, 70, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci NOT NULL COMMENT 'Name der Faxdatei'"
 Str(1, 70, 2) = " `erstellt` datetime DEFAULT NULL COMMENT 'Datum, an dem die Datei zuletzt erstellt wurde'"
 Str(1, 70, 3) = " `geändert` datetime DEFAULT NULL COMMENT 'Datum, an dem die Datei zuletzt geändert wurde'"
 Str(1, 70, 4) = " `Größe` int(10) DEFAULT NULL COMMENT 'Größe in Bytes'"
 Str(1, 70, 5) = " `NameInP` varchar(1000) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Neuer Name im Patientenordner'"
 Str(1, 70, 6) = " `Telefonnummer` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Telefonnummer des Senders'"
 Str(1, 70, 7) = " `Absender` varchar(1000) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Absender, falls bekannt'"
 Str(1, 70, 8) = " `kopiert` tinyint(1) DEFAULT NULL COMMENT 'ob Fax schon kopiert wurde'"
 Str(1, 70, 9) = " `altNameInP` varchar(1000) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Alter Name im Patientenordner'"
 Str(1, 70, 10) = " `ID` int(2) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 70, 11) = " `TMStart` datetime DEFAULT NULL COMMENT 'TransmissionStart'"
 Str(1, 70, 12) = " `TMEnd` datetime DEFAULT NULL COMMENT 'TransmissionEnd'"
 Str(1, 70, 13) = " `fFNr` tinyint(1) unsigned NOT NULL COMMENT 'faxFolder Nr. (1= IncomingQueue, 2 = IncomingArchive'"
 Str(1, 70, 14) = "  PRIMARY KEY (`ID`)"
 Str(1, 70, 15) = "  KEY `kopiertgeändertGröße` (`kopiert`,`geändert`,`Größe`)"
 Str(1, 70, 16) = "  KEY `kopiertName` (`kopiert`,`Name`)"
 Str(1, 70, 17) = "  KEY `S2` (`erstellt`,`Größe`)"
 Str(1, 70, 18) = "  KEY `Telefonnummer` (`Telefonnummer`)"
 Str(1, 70, 19) = "  KEY `geändertGröße` (`geändert`,`Größe`)"
 Str(1, 70, 20) = "  KEY `Name` (`Name`) USING BTREE"
 Str(1, 70, 21) = "  KEY `NameinP` (`NameInP`(767))"
 Str(1, 70, 22) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr70

Sub FüllStr71()
 Str(0, 71, 0) = "forminhaltfeld"
 Str(0, 71, 1) = "`FeldVW`"
 Str(0, 71, 2) = "`Feld`"
 Str(0, 71, 3) = "`StByte`"
 Str(0, 71, 4) = "`FeldVW`"
 Str(0, 71, 5) = "`Feld`"
 ArtZ(0, 71) = 3
 ArtZ(1, 71) = 2
 Str(1, 71, 0) = "CREATE TABLE `forminhaltfeld` ("
 Str(1, 71, 1) = " `FeldVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 71, 2) = " `Feld` longtext COLLATE latin1_german2_ci"
 Str(1, 71, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 71, 4) = "  PRIMARY KEY (`FeldVW`)"
 Str(1, 71, 5) = "  KEY `Feld` (`Feld`(255))"
 Str(1, 71, 6) = " ENGINE=InnoDB AUTO_INCREMENT=13481 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr71

Sub FüllStr72()
 Str(0, 72, 0) = "forminhaltfeldinh"
 Str(0, 72, 1) = "`FeldInhVW`"
 Str(0, 72, 2) = "`FeldInh`"
 Str(0, 72, 3) = "`StByte`"
 Str(0, 72, 4) = "`FeldInhVW`"
 Str(0, 72, 5) = "`FeldInhVW`"
 Str(0, 72, 6) = "`FeldInh`"
 ArtZ(0, 72) = 3
 ArtZ(1, 72) = 3
 Str(1, 72, 0) = "CREATE TABLE `forminhaltfeldinh` ("
 Str(1, 72, 1) = " `FeldInhVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 72, 2) = " `FeldInh` longtext COLLATE latin1_german2_ci"
 Str(1, 72, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 72, 4) = "  PRIMARY KEY (`FeldInhVW`)"
 Str(1, 72, 5) = "  UNIQUE KEY `FeldInhVW` (`FeldInhVW`)"
 Str(1, 72, 6) = "  KEY `FeldInh` (`FeldInh`(255))"
 Str(1, 72, 7) = " ENGINE=InnoDB AUTO_INCREMENT=238927 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr72

Sub FüllStr73()
 Str(0, 73, 0) = "forminhaltform_abk"
 Str(0, 73, 1) = "`Form_AbkVW`"
 Str(0, 73, 2) = "`Form_Abk`"
 Str(0, 73, 3) = "`Form_AbkVW`"
 Str(0, 73, 4) = "`Form_AbkVW`"
 Str(0, 73, 5) = "`Form_Abk`"
 ArtZ(0, 73) = 2
 ArtZ(1, 73) = 3
 Str(1, 73, 0) = "CREATE TABLE `forminhaltform_abk` ("
 Str(1, 73, 1) = " `Form_AbkVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 73, 2) = " `Form_Abk` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 3) = "  PRIMARY KEY (`Form_AbkVW`)"
 Str(1, 73, 4) = "  UNIQUE KEY `Form_AbkVW` (`Form_AbkVW`)"
 Str(1, 73, 5) = "  UNIQUE KEY `Form_Abk` (`Form_Abk`)"
 Str(1, 73, 6) = " ENGINE=InnoDB AUTO_INCREMENT=69 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr73

Sub FüllStr74()
 Str(0, 74, 0) = "forminhfeld"
 Str(0, 74, 1) = "`FoID`"
 Str(0, 74, 2) = "`Nr`"
 Str(0, 74, 3) = "`FeldNr`"
 Str(0, 74, 4) = "`FeldVW`"
 Str(0, 74, 5) = "`FeldInhVW`"
 Str(0, 74, 6) = "`FoID`"
 Str(0, 74, 7) = "`FormInhaltFeldFormInhFeld`"
 Str(0, 74, 8) = "`FormInhaltFeldInhFormInhFeld`"
 Str(0, 74, 9) = "`FormInhFeldFormInhaltFeld`"
 Str(0, 74, 10) = "`FormInhFeldFormInhaltFeldInhalt`"
 Str(0, 74, 11) = "`FormInhFeldFormInhKopf`"
 ArtZ(0, 74) = 5
 ArtZ(1, 74) = 3
 ArtZ(2, 74) = 3
 Str(1, 74, 0) = "CREATE TABLE `forminhfeld` ("
 Str(1, 74, 1) = " `FoID` int(10) DEFAULT NULL"
 Str(1, 74, 2) = " `Nr` smallint(6) DEFAULT NULL"
 Str(1, 74, 3) = " `FeldNr` smallint(6) DEFAULT NULL"
 Str(1, 74, 4) = " `FeldVW` int(10) DEFAULT NULL"
 Str(1, 74, 5) = " `FeldInhVW` int(10) DEFAULT NULL"
 Str(1, 74, 6) = "  KEY `FoID` (`FoID`,`FeldVW`,`FeldNr`)"
 Str(1, 74, 7) = "  KEY `FormInhaltFeldFormInhFeld` (`FeldVW`)"
 Str(1, 74, 8) = "  KEY `FormInhaltFeldInhFormInhFeld` (`FeldInhVW`)"
 Str(1, 74, 9) = "  CONSTRAINT `FormInhFeldFormInhaltFeld` FOREIGN KEY (`FeldVW`) REFERENCES `forminhaltfeld` (`FeldVW`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 74, 10) = "  CONSTRAINT `FormInhFeldFormInhaltFeldInhalt` FOREIGN KEY (`FeldInhVW`) REFERENCES `forminhaltfeldinh` (`FeldInhVW`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 74, 11) = "  CONSTRAINT `FormInhFeldFormInhKopf` FOREIGN KEY (`FoID`) REFERENCES `forminhkopf` (`FoID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 74, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr74

Sub FüllStr75()
 Str(0, 75, 0) = "forminhkopf"
 Str(0, 75, 1) = "`FoID`"
 Str(0, 75, 2) = "`FID`"
 Str(0, 75, 3) = "`Pat_ID`"
 Str(0, 75, 4) = "`Form_ID`"
 Str(0, 75, 5) = "`ZeitPunkt`"
 Str(0, 75, 6) = "`AbsPos`"
 Str(0, 75, 7) = "`AktZeit`"
 Str(0, 75, 8) = "`StByte`"
 Str(0, 75, 9) = "`Satzart`"
 Str(0, 75, 10) = "`Satzlänge`"
 Str(0, 75, 11) = "`LANRid`"
 Str(0, 75, 12) = "`FoID`"
 Str(0, 75, 13) = "`PrimaryKey`"
 Str(0, 75, 14) = "`Auswahl`"
 Str(0, 75, 15) = "`FälleFormInhKopf`"
 Str(0, 75, 16) = "`FID`"
 Str(0, 75, 17) = "`FormulareFormInhKopf`"
 Str(0, 75, 18) = "`NamenFormInhKopf`"
 Str(0, 75, 19) = "`F??lleFormInhKopf_AccRel`"
 Str(0, 75, 20) = "`FormulareFormInhKopf_AccRel`"
 ArtZ(0, 75) = 11
 ArtZ(1, 75) = 7
 ArtZ(2, 75) = 2
 Str(1, 75, 0) = "CREATE TABLE `forminhkopf` ("
 Str(1, 75, 1) = " `FoID` int(10) NOT NULL DEFAULT '0'"
 Str(1, 75, 2) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 75, 3) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 75, 4) = " `Form_ID` int(10) DEFAULT NULL"
 Str(1, 75, 5) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 75, 6) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 75, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 75, 8) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 75, 9) = " `Satzart` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000'"
 Str(1, 75, 10) = " `Satzlänge` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100'"
 Str(1, 75, 11) = " `LANRid` int(3) unsigned NOT NULL COMMENT 'Bezug auf lanrpraxis.id'"
 Str(1, 75, 12) = "  PRIMARY KEY (`FoID`)"
 Str(1, 75, 13) = "  UNIQUE KEY `PrimaryKey` (`FoID`)"
 Str(1, 75, 14) = "  KEY `Auswahl` (`Pat_ID`,`Form_ID`,`ZeitPunkt`)"
 Str(1, 75, 15) = "  KEY `FälleFormInhKopf` (`FID`)"
 Str(1, 75, 16) = "  KEY `FID` (`FID`)"
 Str(1, 75, 17) = "  KEY `FormulareFormInhKopf` (`Form_ID`)"
 Str(1, 75, 18) = "  KEY `NamenFormInhKopf` (`Pat_ID`)"
 Str(1, 75, 19) = "  CONSTRAINT `F??lleFormInhKopf_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 75, 20) = "  CONSTRAINT `FormulareFormInhKopf_AccRel` FOREIGN KEY (`Form_ID`) REFERENCES `formulare` (`FormID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 75, 21) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr75

Sub FüllStr76()
 Str(0, 76, 0) = "formular"
 Str(1, 76, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `formular` AS SELECT `forminhkopf`.`FoID` AS `foid`,`forminhkopf`.`Pat_ID` AS `Pat_ID`,`forminhkopf`.`FID` AS `FID`,`forminhkopf`.`Form_ID` AS `Form_ID`,`forminhkopf`.`ZeitPunkt` AS `ZeitPunkt`,`forminhfeld`.`Nr` AS `Nr`,`forminhfeld`.`FeldNr` AS `FeldNr`,`forminhaltfeld`.`Feld` AS `Feld`,`forminhaltfeldinh`.`FeldInh` AS `FeldInh`,`formulare`.`Form_Abk` AS `form_abk`,`formulare`.`FormVorl` AS `FormVorl` FROM ((((`forminhfeld` LEFT JOIN `forminhkopf` on((`forminhfeld`.`FoID` = `forminhkopf`.`FoID`))) LEFT JOIN `formulare` on((`formulare`.`FormID` = `forminhkopf`.`Form_ID`))) LEFT JOIN `forminhaltfeld` on((`forminhfeld`.`FeldVW` = `forminhaltfeld`.`FeldVW`))) LEFT JOIN `forminhaltfeldinh` on((`forminhfeld`.`FeldInhVW` = `forminhaltfeldinh`.`FeldInhVW`))) ORDER BY `forminhkopf`.`FoID`"
End Sub ' FüllStr76

Sub FüllStr77()
 Str(0, 77, 0) = "formulare"
 Str(0, 77, 1) = "`FormID`"
 Str(0, 77, 2) = "`Form_Abk`"
 Str(0, 77, 3) = "`FormBez`"
 Str(0, 77, 4) = "`FormVorl`"
 Str(0, 77, 5) = "`AktZeit`"
 Str(0, 77, 6) = "`absPos`"
 Str(0, 77, 7) = "`StByte`"
 Str(0, 77, 8) = "`FormID`"
 Str(0, 77, 9) = "`FormID`"
 Str(0, 77, 10) = "`Auswahl`"
 Str(0, 77, 11) = "`FormBez`"
 Str(0, 77, 12) = "`FormInhaltForm_AbkFormulare`"
 Str(0, 77, 13) = "`FormInhaltForm_AbkFormulare_AccRel`"
 ArtZ(0, 77) = 7
 ArtZ(1, 77) = 5
 ArtZ(2, 77) = 1
 Str(1, 77, 0) = "CREATE TABLE `formulare` ("
 Str(1, 77, 1) = " `FormID` int(10) NOT NULL DEFAULT '0'"
 Str(1, 77, 2) = " `Form_Abk` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 77, 3) = " `FormBez` longtext COLLATE latin1_german2_ci"
 Str(1, 77, 4) = " `FormVorl` varchar(114) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 77, 5) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Aktualisierung'"
 Str(1, 77, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 77, 7) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 77, 8) = "  PRIMARY KEY (`FormID`)"
 Str(1, 77, 9) = "  UNIQUE KEY `FormID` (`FormID`)"
 Str(1, 77, 10) = "  KEY `Auswahl` (`Form_Abk`,`FormBez`(255),`FormVorl`)"
 Str(1, 77, 11) = "  KEY `FormBez` (`FormBez`(255))"
 Str(1, 77, 12) = "  KEY `FormInhaltForm_AbkFormulare` (`Form_Abk`)"
 Str(1, 77, 13) = "  CONSTRAINT `FormInhaltForm_AbkFormulare_AccRel` FOREIGN KEY (`Form_Abk`) REFERENCES `forminhaltform_abk` (`Form_Abk`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 77, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr77

Sub FüllStr78()
 Str(0, 78, 0) = "fuerdiagexp"
 Str(0, 78, 1) = "`ID`"
 Str(0, 78, 2) = "`Name`"
 Str(0, 78, 3) = "`Pat_id`"
 Str(0, 78, 4) = "`ICD`"
 Str(0, 78, 5) = "`Diagnose`"
 Str(0, 78, 6) = "`Status`"
 Str(0, 78, 7) = "`Protokoll`"
 Str(0, 78, 8) = "`nurQuart`"
 Str(0, 78, 9) = "`Zeitpunkt`"
 Str(0, 78, 10) = "`ID`"
 Str(0, 78, 11) = "`ID`"
 Str(0, 78, 12) = "`pat_ID`"
 Str(0, 78, 13) = "`Suche`"
 ArtZ(0, 78) = 9
 ArtZ(1, 78) = 4
 Str(1, 78, 0) = "CREATE TABLE `fuerdiagexp` ("
 Str(1, 78, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 78, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 78, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 78, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 78, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 78, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 78, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 78, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 78, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 78, 10) = "  PRIMARY KEY (`ID`)"
 Str(1, 78, 11) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 78, 12) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 78, 13) = "  KEY `Suche` (`Pat_id`,`ICD`,`Diagnose`)"
 Str(1, 78, 14) = " ENGINE=InnoDB AUTO_INCREMENT=1709 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr78

Sub FüllStr79()
 Str(0, 79, 0) = "fuerdiagexparchiv"
 Str(0, 79, 1) = "`ID`"
 Str(0, 79, 2) = "`Name`"
 Str(0, 79, 3) = "`Pat_id`"
 Str(0, 79, 4) = "`ICD`"
 Str(0, 79, 5) = "`Diagnose`"
 Str(0, 79, 6) = "`Status`"
 Str(0, 79, 7) = "`Protokoll`"
 Str(0, 79, 8) = "`nurQuart`"
 Str(0, 79, 9) = "`Zeitpunkt`"
 Str(0, 79, 10) = "`archiviert`"
 Str(0, 79, 11) = "`ID`"
 Str(0, 79, 12) = "`ID`"
 Str(0, 79, 13) = "`pat_ID`"
 Str(0, 79, 14) = "`Suche`"
 ArtZ(0, 79) = 10
 ArtZ(1, 79) = 4
 Str(1, 79, 0) = "CREATE TABLE `fuerdiagexparchiv` ("
 Str(1, 79, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 79, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 79, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 79, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 79, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 79, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 79, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 79, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 79, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 79, 10) = " `archiviert` datetime DEFAULT NULL COMMENT 'Datum der Übertragung in Tabelle Leistungsexport'"
 Str(1, 79, 11) = "  PRIMARY KEY (`ID`)"
 Str(1, 79, 12) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 79, 13) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 79, 14) = "  KEY `Suche` (`Pat_id`,`ICD`,`Diagnose`)"
 Str(1, 79, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr79

Sub FüllStr80()
 Str(0, 80, 0) = "fuerleistexp"
 Str(0, 80, 1) = "`ID`"
 Str(0, 80, 2) = "`Name`"
 Str(0, 80, 3) = "`Datum`"
 Str(0, 80, 4) = "`Pat_id`"
 Str(0, 80, 5) = "`SchGr`"
 Str(0, 80, 6) = "`Status`"
 Str(0, 80, 7) = "`Protokoll`"
 Str(0, 80, 8) = "`ID`"
 Str(0, 80, 9) = "`PrimaryKey`"
 Str(0, 80, 10) = "`ID`"
 Str(0, 80, 11) = "`NamenfürLeistExp`"
 Str(0, 80, 12) = "`Namenf??rLeistExp_AccRel`"
 ArtZ(0, 80) = 7
 ArtZ(1, 80) = 4
 ArtZ(2, 80) = 1
 Str(1, 80, 0) = "CREATE TABLE `fuerleistexp` ("
 Str(1, 80, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 80, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 80, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen'"
 Str(1, 80, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 80, 5) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 80, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 80, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 80, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 80, 9) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 80, 10) = "  KEY `ID` (`Pat_id`)"
 Str(1, 80, 11) = "  KEY `NamenfürLeistExp` (`Pat_id`)"
 Str(1, 80, 12) = "  CONSTRAINT `Namenf??rLeistExp_AccRel` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 80, 13) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr80

Sub FüllStr81()
 Str(0, 81, 0) = "fuerleistexparchiv"
 Str(0, 81, 1) = "`ID`"
 Str(0, 81, 2) = "`Name`"
 Str(0, 81, 3) = "`Datum`"
 Str(0, 81, 4) = "`Pat_id`"
 Str(0, 81, 5) = "`SchGr`"
 Str(0, 81, 6) = "`archiviert`"
 Str(0, 81, 7) = "`Protokoll`"
 Str(0, 81, 8) = "`ID`"
 Str(0, 81, 9) = "`PrimaryKey`"
 Str(0, 81, 10) = "`ID`"
 ArtZ(0, 81) = 7
 ArtZ(1, 81) = 3
 Str(1, 81, 0) = "CREATE TABLE `fuerleistexparchiv` ("
 Str(1, 81, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 81, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 81, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen'"
 Str(1, 81, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 81, 5) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 81, 6) = " `archiviert` datetime DEFAULT NULL COMMENT 'Datum der Übertragung in Tabelle Leistungsexport'"
 Str(1, 81, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 81, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 81, 9) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 81, 10) = "  KEY `ID` (`Pat_id`)"
 Str(1, 81, 11) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr81

Sub FüllStr82()
 Str(0, 82, 0) = "genehmigungen"
 Str(0, 82, 1) = "`MYID`"
 Str(0, 82, 2) = "`eingid`"
 Str(0, 82, 3) = "`Leistung`"
 Str(0, 82, 4) = "`Erklärung`"
 Str(0, 82, 5) = "`Kassen`"
 Str(0, 82, 6) = "`Kothny`"
 Str(0, 82, 7) = "`Schade`"
 Str(0, 82, 8) = "`DMP1`"
 Str(0, 82, 9) = "`DMP2`"
 Str(0, 82, 10) = "`DTyp`"
 Str(0, 82, 11) = "`Kht`"
 Str(0, 82, 12) = "`Weibl`"
 Str(0, 82, 13) = "`von`"
 Str(0, 82, 14) = "`bis`"
 Str(0, 82, 15) = "`Therarten`"
 Str(0, 82, 16) = "`mitÜw`"
 Str(0, 82, 17) = "`ICD`"
 Str(0, 82, 18) = "`LAusschluß`"
 Str(0, 82, 19) = "`obSchulung`"
 Str(0, 82, 20) = "`Wert`"
 Str(0, 82, 21) = "`MYID`"
 Str(0, 82, 22) = "`Leistung`"
 Str(0, 82, 23) = "`fk_LEKKSNNNNNNNN`"
 Str(0, 82, 24) = "`obschulung`"
 ArtZ(0, 82) = 20
 ArtZ(1, 82) = 4
 Str(1, 82, 0) = "CREATE TABLE `genehmigungen` ("
 Str(1, 82, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 82, 2) = " `eingid` int(10) unsigned NOT NULL DEFAULT '0'"
 Str(1, 82, 3) = " `Leistung` varchar(6) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 82, 4) = " `Erklärung` varchar(198) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 82, 5) = " `Kassen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 82, 6) = " `Kothny` date NOT NULL DEFAULT '0000-00-00'"
 Str(1, 82, 7) = " `Schade` date NOT NULL DEFAULT '0000-00-00'"
 Str(1, 82, 8) = " `DMP1` tinyint(1) unsigned DEFAULT NULL COMMENT 'Null=egal, 0 = nicht im DMP gefordert'"
 Str(1, 82, 9) = " `DMP2` tinyint(1) unsigned DEFAULT NULL COMMENT 'Null=egal, 0 = nicht im DMP gefordert'"
 Str(1, 82, 10) = " `DTyp` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diabetestypen (1,2,g), ggf. durch Komma getrennt'"
 Str(1, 82, 11) = " `Kht` int(10) unsigned NOT NULL DEFAULT '1' COMMENT 'Krankeit: 0=Sachkosten 1=D.m., 2=Hypertonie'"
 Str(1, 82, 12) = " `Weibl` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT '1=Weiblichkeit gefordert (""(geplante) Schwangerschaft"")'"
 Str(1, 82, 13) = " `von` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Alter von'"
 Str(1, 82, 14) = " `bis` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Alter bis'"
 Str(1, 82, 15) = " `Therarten` varchar(30) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'mögliche Therapiearten, durch ,  getrennt: Diät,OAD,CT,Komb,ICT, CSII'"
 Str(1, 82, 16) = " `mitÜw` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT '1=Überweisung nötig, 2=Überweisung nur wenn nicht bei uns im DMP'"
 Str(1, 82, 17) = " `ICD` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'ICD-Muster für regexp'"
 Str(1, 82, 18) = " `LAusschluß` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Leistungsausschluß, Muster für regexp'"
 Str(1, 82, 19) = " `obSchulung` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'ob Leistung eine Schulungsleistung ist'"
 Str(1, 82, 20) = " `Wert` decimal(5,2) DEFAULT NULL COMMENT 'Wert der Leistung in Euro'"
 Str(1, 82, 21) = "  PRIMARY KEY (`MYID`)"
 Str(1, 82, 22) = "  UNIQUE KEY `Leistung` (`Leistung`,`DTyp`) USING BTREE"
 Str(1, 82, 23) = "  KEY `fk_LEKKSNNNNNNNN` (`eingid`)"
 Str(1, 82, 24) = "  KEY `obschulung` (`obSchulung`)"
 Str(1, 82, 25) = " ENGINE=InnoDB AUTO_INCREMENT=128 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Tabelle aus Datei: genehmigungen.csv'"
End Sub ' FüllStr82

Sub FüllStr83()
 Str(0, 83, 0) = "genehmigungen0"
 Str(0, 83, 1) = "`MYID`"
 Str(0, 83, 2) = "`eingid`"
 Str(0, 83, 3) = "`Leistung`"
 Str(0, 83, 4) = "`Erklärung`"
 Str(0, 83, 5) = "`Kassen`"
 Str(0, 83, 6) = "`Kothny`"
 Str(0, 83, 7) = "`Schade`"
 Str(0, 83, 8) = "`DMP1`"
 Str(0, 83, 9) = "`DMP2`"
 Str(0, 83, 10) = "`DTyp`"
 Str(0, 83, 11) = "`Weibl`"
 Str(0, 83, 12) = "`von`"
 Str(0, 83, 13) = "`bis`"
 Str(0, 83, 14) = "`Therarten`"
 Str(0, 83, 15) = "`mitÜw`"
 Str(0, 83, 16) = "`ICD`"
 Str(0, 83, 17) = "`LAusschluß`"
 Str(0, 83, 18) = "`obSchulung`"
 Str(0, 83, 19) = "`Wert`"
 Str(0, 83, 20) = "`MYID`"
 Str(0, 83, 21) = "`Definition`"
 Str(0, 83, 22) = "`fk_LEKKSNNNNNNNN`"
 Str(0, 83, 23) = "`Leistung`"
 Str(0, 83, 24) = "`obschulung`"
 ArtZ(0, 83) = 19
 ArtZ(1, 83) = 5
 Str(1, 83, 0) = "CREATE TABLE `genehmigungen0` ("
 Str(1, 83, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 83, 2) = " `eingid` int(10) unsigned NOT NULL DEFAULT '0'"
 Str(1, 83, 3) = " `Leistung` varchar(6) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 83, 4) = " `Erklärung` varchar(198) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 83, 5) = " `Kassen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 83, 6) = " `Kothny` date NOT NULL DEFAULT '0000-00-00'"
 Str(1, 83, 7) = " `Schade` date NOT NULL DEFAULT '0000-00-00'"
 Str(1, 83, 8) = " `DMP1` tinyint(1) unsigned DEFAULT NULL COMMENT 'Null=egal, 0 = nicht im DMP gefordert'"
 Str(1, 83, 9) = " `DMP2` tinyint(1) unsigned DEFAULT NULL COMMENT 'Null=egal, 0 = nicht im DMP gefordert'"
 Str(1, 83, 10) = " `DTyp` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diabetestypen (1,2,g), ggf. durch Komma getrennt'"
 Str(1, 83, 11) = " `Weibl` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT '1=Weiblichkeit gefordert (""(geplante) Schwangerschaft"")'"
 Str(1, 83, 12) = " `von` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Alter von'"
 Str(1, 83, 13) = " `bis` int(3) unsigned NOT NULL DEFAULT '0' COMMENT 'Alter bis'"
 Str(1, 83, 14) = " `Therarten` varchar(30) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'mögliche Therapiearten, durch ,  getrennt: Diät,OAD,CT,Komb,ICT, CSII'"
 Str(1, 83, 15) = " `mitÜw` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT '1=Überweisung nötig'"
 Str(1, 83, 16) = " `ICD` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'ICD-Muster für regexp'"
 Str(1, 83, 17) = " `LAusschluß` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Leistungsausschluß, Muster für regexp'"
 Str(1, 83, 18) = " `obSchulung` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'ob Leistung eine Schulungsleistung ist'"
 Str(1, 83, 19) = " `Wert` decimal(5,2) NOT NULL COMMENT 'Wert der Leistung in Euro'"
 Str(1, 83, 20) = "  PRIMARY KEY (`MYID`)"
 Str(1, 83, 21) = "  UNIQUE KEY `Definition` (`Leistung`,`DTyp`)"
 Str(1, 83, 22) = "  KEY `fk_LEKKSNNNNNNNN` (`eingid`)"
 Str(1, 83, 23) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 83, 24) = "  KEY `obschulung` (`obSchulung`)"
 Str(1, 83, 25) = " ENGINE=InnoDB AUTO_INCREMENT=123 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: genehmigungen.csv'"
End Sub ' FüllStr83

Sub FüllStr84()
 Str(0, 84, 0) = "genehmigungen_alt"
 Str(0, 84, 1) = "`MYID`"
 Str(0, 84, 2) = "`eingid`"
 Str(0, 84, 3) = "`Leistung`"
 Str(0, 84, 4) = "`Erklärung`"
 Str(0, 84, 5) = "`Kassen`"
 Str(0, 84, 6) = "`Kothny`"
 Str(0, 84, 7) = "`Schade`"
 Str(0, 84, 8) = "`DMP1`"
 Str(0, 84, 9) = "`DMP2`"
 Str(0, 84, 10) = "`DTyp`"
 Str(0, 84, 11) = "`Weibl`"
 Str(0, 84, 12) = "`von`"
 Str(0, 84, 13) = "`bis`"
 Str(0, 84, 14) = "`Therarten`"
 Str(0, 84, 15) = "`mitÜw`"
 Str(0, 84, 16) = "`MYID`"
 Str(0, 84, 17) = "`fk_LEKKSNNNNNNNN`"
 Str(0, 84, 18) = "`Leistung`"
 Str(0, 84, 19) = "`fk_LEKKSNNNNNNNN`"
 ArtZ(0, 84) = 15
 ArtZ(1, 84) = 3
 ArtZ(2, 84) = 1
 Str(1, 84, 0) = "CREATE TABLE `genehmigungen_alt` ("
 Str(1, 84, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 84, 2) = " `eingid` int(10) unsigned NOT NULL DEFAULT '0'"
 Str(1, 84, 3) = " `Leistung` varchar(6) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 84, 4) = " `Erklärung` varchar(198) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 84, 5) = " `Kassen` varchar(29) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 84, 6) = " `Kothny` date NOT NULL DEFAULT '0000-00-00'"
 Str(1, 84, 7) = " `Schade` date NOT NULL DEFAULT '0000-00-00'"
 Str(1, 84, 8) = " `DMP1` tinyint(1) unsigned DEFAULT NULL COMMENT 'Null=egal, 0 = nicht im DMP gefordert'"
 Str(1, 84, 9) = " `DMP2` tinyint(1) unsigned DEFAULT NULL COMMENT 'Null=egal, 0 = nicht im DMP gefordert'"
 Str(1, 84, 10) = " `DTyp` varchar(3) COLLATE latin1_german2_ci NOT NULL COMMENT 'Diabetestypen (1,2,g), ggf. durch Komma getrennt'"
 Str(1, 84, 11) = " `Weibl` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT '1=Weiblichkeit gefordert (""(geplante) Schwangerschaft"")'"
 Str(1, 84, 12) = " `von` int(3) unsigned NOT NULL COMMENT 'Alter von'"
 Str(1, 84, 13) = " `bis` int(3) unsigned NOT NULL COMMENT 'Alter bis'"
 Str(1, 84, 14) = " `Therarten` varchar(30) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'mögliche Therapiearten, durch ,  getrennt: Diät,OAD,CT,Komb,ICT, CSII'"
 Str(1, 84, 15) = " `mitÜw` tinyint(1) unsigned NOT NULL COMMENT '1=Überweisung nötig'"
 Str(1, 84, 16) = "  PRIMARY KEY (`MYID`)"
 Str(1, 84, 17) = "  KEY `fk_LEKKSNNNNNNNN` (`eingid`)"
 Str(1, 84, 18) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 84, 19) = "  CONSTRAINT `fk_LEKKSNNNNNNNN` FOREIGN KEY (`eingid`) REFERENCES `eingelesen1` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 84, 20) = " ENGINE=InnoDB AUTO_INCREMENT=95 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: genehmigungen.csv'"
End Sub ' FüllStr84

Sub FüllStr85()
 Str(0, 85, 0) = "gewicht"
 Str(0, 85, 1) = "`id`"
 Str(0, 85, 2) = "`FID`"
 Str(0, 85, 3) = "`Pat_ID`"
 Str(0, 85, 4) = "`ZeitPunkt`"
 Str(0, 85, 5) = "`Gewicht`"
 Str(0, 85, 6) = "`absPos`"
 Str(0, 85, 7) = "`AktZeit`"
 Str(0, 85, 8) = "`QS`"
 Str(0, 85, 9) = "`QT`"
 Str(0, 85, 10) = "`StByte`"
 Str(0, 85, 11) = "`inhNum`"
 Str(0, 85, 12) = "`id`"
 ArtZ(0, 85) = 11
 ArtZ(1, 85) = 1
 Str(1, 85, 0) = "CREATE TABLE `gewicht` ("
 Str(1, 85, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 85, 2) = " `FID` int(10) DEFAULT NULL"
 Str(1, 85, 3) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 85, 4) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 85, 5) = " `Gewicht` decimal(5,1) DEFAULT NULL"
 Str(1, 85, 6) = " `absPos` int(10) DEFAULT NULL"
 Str(1, 85, 7) = " `AktZeit` datetime DEFAULT NULL"
 Str(1, 85, 8) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 85, 9) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 85, 10) = " `StByte` int(10) DEFAULT NULL"
 Str(1, 85, 11) = " `inhNum` double DEFAULT NULL"
 Str(1, 85, 12) = "  PRIMARY KEY (`id`)"
 Str(1, 85, 13) = " ENGINE=InnoDB AUTO_INCREMENT=43157 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr85

Sub FüllStr86()
 Str(0, 86, 0) = "ha2"
 Str(0, 86, 1) = "`ID`"
 Str(0, 86, 2) = "`Überschrift`"
 Str(0, 86, 3) = "`Name`"
 Str(0, 86, 4) = "`Vorname`"
 Str(0, 86, 5) = "`Nachname`"
 Str(0, 86, 6) = "`Anschrift`"
 Str(0, 86, 7) = "`KVNr`"
 Str(0, 86, 8) = "`Telefon`"
 Str(0, 86, 9) = "`Telefax`"
 Str(0, 86, 10) = "`E_Mail`"
 Str(0, 86, 11) = "`Zulassungsgebiet`"
 Str(0, 86, 12) = "`Arzttyp`"
 Str(0, 86, 13) = "`Gemeinschaftspraxis mit`"
 Str(0, 86, 14) = "`Schwerpunkt`"
 Str(0, 86, 15) = "`Zusatzbezeichnung`"
 Str(0, 86, 16) = "`Bemerkung`"
 Str(0, 86, 17) = "`Beme`"
 Str(0, 86, 18) = "`Sprechstunden`"
 Str(0, 86, 19) = "`von _ bis`"
 Str(0, 86, 20) = "`Internetadressen`"
 Str(0, 86, 21) = "`Behandlung in Fremdsprachen`"
 Str(0, 86, 22) = "`Rollstuhlgerechte Praxis`"
 Str(0, 86, 23) = "`Verkehrsmittel`"
 Str(0, 86, 24) = "`Linie`"
 Str(0, 86, 25) = "`Haltestelle Parkplätze`"
 Str(0, 86, 26) = "`Wegbeschreibung`"
 Str(0, 86, 27) = "`Entfernung zur Praxis`"
 Str(0, 86, 28) = "`Zahl`"
 Str(0, 86, 29) = "`nichtmehr`"
 Str(0, 86, 30) = "`Titel`"
 Str(0, 86, 31) = "`Geschlecht`"
 Str(0, 86, 32) = "`Straße`"
 Str(0, 86, 33) = "`PLZ`"
 Str(0, 86, 34) = "`Ort`"
 Str(0, 86, 35) = "`DMPT2`"
 Str(0, 86, 36) = "`DMPT1`"
 Str(0, 86, 37) = "`gelöscht`"
 Str(0, 86, 38) = "`ID`"
 Str(0, 86, 39) = "`PrimaryKey`"
 Str(0, 86, 40) = "`Auswahl`"
 Str(0, 86, 41) = "`KVNr`"
 ArtZ(0, 86) = 37
 ArtZ(1, 86) = 4
 Str(1, 86, 0) = "CREATE TABLE `ha2` ("
 Str(1, 86, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 86, 2) = " `Überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '""L"" = Liebe(r), ""H"" = Hallo'"
 Str(1, 86, 3) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 4) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 5) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 6) = " `Anschrift` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 7) = " `KVNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 8) = " `Telefon` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 9) = " `Telefax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 10) = " `E_Mail` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 11) = " `Zulassungsgebiet` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 12) = " `Arzttyp` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 13) = " `Gemeinschaftspraxis mit` longtext COLLATE latin1_german2_ci"
 Str(1, 86, 14) = " `Schwerpunkt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 15) = " `Zusatzbezeichnung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 16) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 86, 17) = " `Beme` longtext COLLATE latin1_german2_ci"
 Str(1, 86, 18) = " `Sprechstunden` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 19) = " `von _ bis` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 20) = " `Internetadressen` longtext COLLATE latin1_german2_ci"
 Str(1, 86, 21) = " `Behandlung in Fremdsprachen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 22) = " `Rollstuhlgerechte Praxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 23) = " `Verkehrsmittel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 24) = " `Linie` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 25) = " `Haltestelle Parkplätze` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 26) = " `Wegbeschreibung` longtext COLLATE latin1_german2_ci"
 Str(1, 86, 27) = " `Entfernung zur Praxis` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 28) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 86, 29) = " `nichtmehr` bit(1) DEFAULT NULL COMMENT 'Arzt nicht mehr im Verzeichnis'"
 Str(1, 86, 30) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 31) = " `Geschlecht` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 32) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 33) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 34) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 86, 35) = " `DMPT2` smallint(6) DEFAULT NULL"
 Str(1, 86, 36) = " `DMPT1` bit(1) DEFAULT NULL"
 Str(1, 86, 37) = " `gelöscht` bit(1) DEFAULT NULL"
 Str(1, 86, 38) = "  PRIMARY KEY (`ID`)"
 Str(1, 86, 39) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 86, 40) = "  KEY `Auswahl` (`KVNr`,`Name`)"
 Str(1, 86, 41) = "  KEY `KVNr` (`KVNr`,`Nachname`,`Vorname`)"
 Str(1, 86, 42) = " ENGINE=InnoDB AUTO_INCREMENT=28715 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr86

Sub FüllStr87()
 Str(0, 87, 0) = "hareal"
 Str(0, 87, 1) = "`Anrede`"
 Str(0, 87, 2) = "`Adressat`"
 Str(0, 87, 3) = "`Straße`"
 Str(0, 87, 4) = "`PLZOrt`"
 Str(0, 87, 5) = "`Fax`"
 Str(0, 87, 6) = "`Überschrift`"
 Str(0, 87, 7) = "`dmp2`"
 Str(0, 87, 8) = "`dmp1`"
 Str(0, 87, 9) = "`Niederlassungsgebiet`"
 Str(0, 87, 10) = "`Vorname`"
 Str(0, 87, 11) = "`Funktion`"
 Str(0, 87, 12) = "`InnereAllg`"
 Str(0, 87, 13) = "`kvnr`"
 Str(0, 87, 14) = "`Tel`"
 Str(0, 87, 15) = "`Nachname`"
 Str(0, 87, 16) = "`kvnr`"
 ArtZ(0, 87) = 15
 ArtZ(1, 87) = 1
 Str(1, 87, 0) = "CREATE TABLE `hareal` ("
 Str(1, 87, 1) = " `Anrede` tinyint(1) DEFAULT NULL COMMENT '0=Frau,1=Herrn'"
 Str(1, 87, 2) = " `Adressat` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Titel+Vorn+Nachn'"
 Str(1, 87, 3) = " `Straße` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 87, 4) = " `PLZOrt` varchar(34) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 87, 5) = " `Fax` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 87, 6) = " `Überschrift` varchar(56) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 87, 7) = " `dmp2` tinyint(1) DEFAULT NULL COMMENT '0=nein,1=ja'"
 Str(1, 87, 8) = " `dmp1` tinyint(1) DEFAULT NULL COMMENT '0=nein,1=ja'"
 Str(1, 87, 9) = " `Niederlassungsgebiet` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Med.Fachrichtung'"
 Str(1, 87, 10) = " `Vorname` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 87, 11) = " `Funktion` varchar(0) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'nur noch Schaltfeld, Inhalt in `namen`'"
 Str(1, 87, 12) = " `InnereAllg` tinyint(1) DEFAULT NULL COMMENT '1=Innere oder Allgemeinmedizin'"
 Str(1, 87, 13) = " `kvnr` int(10) unsigned NOT NULL"
 Str(1, 87, 14) = " `Tel` varchar(13) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 87, 15) = " `Nachname` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 87, 16) = "  PRIMARY KEY (`kvnr`)"
 Str(1, 87, 17) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr87

Sub FüllStr88()
 Str(0, 88, 0) = "harealalt"
 Str(0, 88, 1) = "`Anrede`"
 Str(0, 88, 2) = "`Adressat`"
 Str(0, 88, 3) = "`Straße`"
 Str(0, 88, 4) = "`PLZOrt`"
 Str(0, 88, 5) = "`Fax`"
 Str(0, 88, 6) = "`Überschrift`"
 Str(0, 88, 7) = "`dmp2`"
 Str(0, 88, 8) = "`dmp1`"
 Str(0, 88, 9) = "`Niederlassungsgebiet`"
 Str(0, 88, 10) = "`Vorname`"
 Str(0, 88, 11) = "`Funktion`"
 Str(0, 88, 12) = "`InnereAllg`"
 Str(0, 88, 13) = "`kvnr`"
 Str(0, 88, 14) = "`Tel`"
 Str(0, 88, 15) = "`Nachname`"
 Str(0, 88, 16) = "`kvnr`"
 ArtZ(0, 88) = 15
 ArtZ(1, 88) = 1
 Str(1, 88, 0) = "CREATE TABLE `harealalt` ("
 Str(1, 88, 1) = " `Anrede` bit(1) DEFAULT NULL COMMENT '0=Frau,1=Herrn'"
 Str(1, 88, 2) = " `Adressat` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Titel+Vorn+Nachn'"
 Str(1, 88, 3) = " `Straße` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 88, 4) = " `PLZOrt` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 88, 5) = " `Fax` varchar(17) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 88, 6) = " `Überschrift` varchar(56) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 88, 7) = " `dmp2` bit(1) DEFAULT NULL COMMENT '0=nein,1=ja'"
 Str(1, 88, 8) = " `dmp1` bit(1) DEFAULT NULL COMMENT '0=nein,1=ja'"
 Str(1, 88, 9) = " `Niederlassungsgebiet` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Med.Fachrichtung'"
 Str(1, 88, 10) = " `Vorname` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 88, 11) = " `Funktion` varchar(0) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'nur noch Schaltfeld, Inhalt in `namen`'"
 Str(1, 88, 12) = " `InnereAllg` bit(1) DEFAULT NULL COMMENT '1=Innere oder Allgemeinmedizin'"
 Str(1, 88, 13) = " `kvnr` int(10) unsigned NOT NULL"
 Str(1, 88, 14) = " `Tel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 88, 15) = " `Nachname` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 88, 16) = "  PRIMARY KEY (`kvnr`)"
 Str(1, 88, 17) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr88

Sub FüllStr89()
 Str(0, 89, 0) = "hausaerzte"
 Str(0, 89, 1) = "`ID`"
 Str(0, 89, 2) = "`Überschrift`"
 Str(0, 89, 3) = "`Name`"
 Str(0, 89, 4) = "`Vorname`"
 Str(0, 89, 5) = "`Nachname`"
 Str(0, 89, 6) = "`Anschrift`"
 Str(0, 89, 7) = "`KVNr`"
 Str(0, 89, 8) = "`Telefon`"
 Str(0, 89, 9) = "`Telefax`"
 Str(0, 89, 10) = "`E_Mail`"
 Str(0, 89, 11) = "`Zulassungsgebiet`"
 Str(0, 89, 12) = "`Arzttyp`"
 Str(0, 89, 13) = "`Gemeinschaftspraxis mit`"
 Str(0, 89, 14) = "`Schwerpunkt`"
 Str(0, 89, 15) = "`Zusatzbezeichnung`"
 Str(0, 89, 16) = "`Bemerkung`"
 Str(0, 89, 17) = "`Beme`"
 Str(0, 89, 18) = "`Sprechstunden`"
 Str(0, 89, 19) = "`von _ bis`"
 Str(0, 89, 20) = "`Internetadressen`"
 Str(0, 89, 21) = "`Behandlung in Fremdsprachen`"
 Str(0, 89, 22) = "`Rollstuhlgerechte Praxis`"
 Str(0, 89, 23) = "`Verkehrsmittel`"
 Str(0, 89, 24) = "`Linie`"
 Str(0, 89, 25) = "`Haltestelle Parkplätze`"
 Str(0, 89, 26) = "`Wegbeschreibung`"
 Str(0, 89, 27) = "`Entfernung zur Praxis`"
 Str(0, 89, 28) = "`Zahl`"
 Str(0, 89, 29) = "`nichtmehr`"
 Str(0, 89, 30) = "`Titel`"
 Str(0, 89, 31) = "`Geschlecht`"
 Str(0, 89, 32) = "`Straße`"
 Str(0, 89, 33) = "`PLZ`"
 Str(0, 89, 34) = "`Ort`"
 Str(0, 89, 35) = "`DMPT2`"
 Str(0, 89, 36) = "`DMPT1`"
 Str(0, 89, 37) = "`gelöscht`"
 Str(0, 89, 38) = "`ID`"
 Str(0, 89, 39) = "`PrimaryKey`"
 Str(0, 89, 40) = "`Auswahl`"
 Str(0, 89, 41) = "`KVNr`"
 ArtZ(0, 89) = 37
 ArtZ(1, 89) = 4
 Str(1, 89, 0) = "CREATE TABLE `hausaerzte` ("
 Str(1, 89, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 89, 2) = " `Überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '""L"" = Liebe(r), ""H"" = Hallo'"
 Str(1, 89, 3) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 4) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 5) = " `Nachname` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 6) = " `Anschrift` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 7) = " `KVNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 8) = " `Telefon` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 9) = " `Telefax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 10) = " `E_Mail` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 11) = " `Zulassungsgebiet` varchar(592) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 12) = " `Arzttyp` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 13) = " `Gemeinschaftspraxis mit` longtext COLLATE latin1_german2_ci"
 Str(1, 89, 14) = " `Schwerpunkt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 15) = " `Zusatzbezeichnung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 16) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 89, 17) = " `Beme` longtext COLLATE latin1_german2_ci"
 Str(1, 89, 18) = " `Sprechstunden` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 19) = " `von _ bis` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 20) = " `Internetadressen` longtext COLLATE latin1_german2_ci"
 Str(1, 89, 21) = " `Behandlung in Fremdsprachen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 22) = " `Rollstuhlgerechte Praxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 23) = " `Verkehrsmittel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 24) = " `Linie` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 25) = " `Haltestelle Parkplätze` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 26) = " `Wegbeschreibung` longtext COLLATE latin1_german2_ci"
 Str(1, 89, 27) = " `Entfernung zur Praxis` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 28) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 89, 29) = " `nichtmehr` tinyint(1) unsigned DEFAULT NULL COMMENT 'Arzt nicht mehr im Verzeichnis'"
 Str(1, 89, 30) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 31) = " `Geschlecht` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 32) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 33) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 34) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 35) = " `DMPT2` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 89, 36) = " `DMPT1` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 89, 37) = " `gelöscht` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 89, 38) = "  PRIMARY KEY (`ID`)"
 Str(1, 89, 39) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 89, 40) = "  KEY `Auswahl` (`KVNr`,`Name`)"
 Str(1, 89, 41) = "  KEY `KVNr` (`KVNr`,`Nachname`,`Vorname`)"
 Str(1, 89, 42) = " ENGINE=InnoDB AUTO_INCREMENT=1591 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr89

Sub FüllStr90()
 Str(0, 90, 0) = "hausaerztealt"
 Str(0, 90, 1) = "`ID`"
 Str(0, 90, 2) = "`Überschrift`"
 Str(0, 90, 3) = "`Name`"
 Str(0, 90, 4) = "`Vorname`"
 Str(0, 90, 5) = "`Nachname`"
 Str(0, 90, 6) = "`Anschrift`"
 Str(0, 90, 7) = "`KVNr`"
 Str(0, 90, 8) = "`Telefon`"
 Str(0, 90, 9) = "`Telefax`"
 Str(0, 90, 10) = "`E_Mail`"
 Str(0, 90, 11) = "`Zulassungsgebiet`"
 Str(0, 90, 12) = "`Arzttyp`"
 Str(0, 90, 13) = "`Gemeinschaftspraxis mit`"
 Str(0, 90, 14) = "`Schwerpunkt`"
 Str(0, 90, 15) = "`Zusatzbezeichnung`"
 Str(0, 90, 16) = "`Bemerkung`"
 Str(0, 90, 17) = "`Beme`"
 Str(0, 90, 18) = "`Sprechstunden`"
 Str(0, 90, 19) = "`von _ bis`"
 Str(0, 90, 20) = "`Internetadressen`"
 Str(0, 90, 21) = "`Behandlung in Fremdsprachen`"
 Str(0, 90, 22) = "`Rollstuhlgerechte Praxis`"
 Str(0, 90, 23) = "`Verkehrsmittel`"
 Str(0, 90, 24) = "`Linie`"
 Str(0, 90, 25) = "`Haltestelle Parkplätze`"
 Str(0, 90, 26) = "`Wegbeschreibung`"
 Str(0, 90, 27) = "`Entfernung zur Praxis`"
 Str(0, 90, 28) = "`Zahl`"
 Str(0, 90, 29) = "`nichtmehr`"
 Str(0, 90, 30) = "`Titel`"
 Str(0, 90, 31) = "`Geschlecht`"
 Str(0, 90, 32) = "`Straße`"
 Str(0, 90, 33) = "`PLZ`"
 Str(0, 90, 34) = "`Ort`"
 Str(0, 90, 35) = "`DMPT2`"
 Str(0, 90, 36) = "`DMPT1`"
 Str(0, 90, 37) = "`gelöscht`"
 Str(0, 90, 38) = "`ID`"
 Str(0, 90, 39) = "`PrimaryKey`"
 Str(0, 90, 40) = "`Auswahl`"
 Str(0, 90, 41) = "`KVNr`"
 ArtZ(0, 90) = 37
 ArtZ(1, 90) = 4
 Str(1, 90, 0) = "CREATE TABLE `hausaerztealt` ("
 Str(1, 90, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 90, 2) = " `Überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '""L"" = Liebe(r), ""H"" = Hallo'"
 Str(1, 90, 3) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 4) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 5) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 6) = " `Anschrift` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 7) = " `KVNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 8) = " `Telefon` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 9) = " `Telefax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 10) = " `E_Mail` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 11) = " `Zulassungsgebiet` varchar(217) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 12) = " `Arzttyp` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 13) = " `Gemeinschaftspraxis mit` longtext COLLATE latin1_german2_ci"
 Str(1, 90, 14) = " `Schwerpunkt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 15) = " `Zusatzbezeichnung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 16) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 90, 17) = " `Beme` longtext COLLATE latin1_german2_ci"
 Str(1, 90, 18) = " `Sprechstunden` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 19) = " `von _ bis` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 20) = " `Internetadressen` longtext COLLATE latin1_german2_ci"
 Str(1, 90, 21) = " `Behandlung in Fremdsprachen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 22) = " `Rollstuhlgerechte Praxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 23) = " `Verkehrsmittel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 24) = " `Linie` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 25) = " `Haltestelle Parkplätze` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 26) = " `Wegbeschreibung` longtext COLLATE latin1_german2_ci"
 Str(1, 90, 27) = " `Entfernung zur Praxis` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 28) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 90, 29) = " `nichtmehr` bit(1) DEFAULT NULL COMMENT 'Arzt nicht mehr im Verzeichnis'"
 Str(1, 90, 30) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 31) = " `Geschlecht` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 32) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 33) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 34) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 35) = " `DMPT2` smallint(6) DEFAULT NULL"
 Str(1, 90, 36) = " `DMPT1` bit(1) DEFAULT NULL"
 Str(1, 90, 37) = " `gelöscht` bit(1) DEFAULT NULL"
 Str(1, 90, 38) = "  PRIMARY KEY (`ID`)"
 Str(1, 90, 39) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 90, 40) = "  KEY `Auswahl` (`KVNr`,`Name`)"
 Str(1, 90, 41) = "  KEY `KVNr` (`KVNr`,`Nachname`,`Vorname`)"
 Str(1, 90, 42) = " ENGINE=InnoDB AUTO_INCREMENT=108232 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr90

Sub FüllStr91()
 Str(0, 91, 0) = "hausaerztealt2"
 Str(0, 91, 1) = "`ID`"
 Str(0, 91, 2) = "`Überschrift`"
 Str(0, 91, 3) = "`Name`"
 Str(0, 91, 4) = "`Vorname`"
 Str(0, 91, 5) = "`Nachname`"
 Str(0, 91, 6) = "`Anschrift`"
 Str(0, 91, 7) = "`KVNr`"
 Str(0, 91, 8) = "`Telefon`"
 Str(0, 91, 9) = "`Telefax`"
 Str(0, 91, 10) = "`E_Mail`"
 Str(0, 91, 11) = "`Zulassungsgebiet`"
 Str(0, 91, 12) = "`Arzttyp`"
 Str(0, 91, 13) = "`Gemeinschaftspraxis mit`"
 Str(0, 91, 14) = "`Schwerpunkt`"
 Str(0, 91, 15) = "`Zusatzbezeichnung`"
 Str(0, 91, 16) = "`Bemerkung`"
 Str(0, 91, 17) = "`Beme`"
 Str(0, 91, 18) = "`Sprechstunden`"
 Str(0, 91, 19) = "`von _ bis`"
 Str(0, 91, 20) = "`Internetadressen`"
 Str(0, 91, 21) = "`Behandlung in Fremdsprachen`"
 Str(0, 91, 22) = "`Rollstuhlgerechte Praxis`"
 Str(0, 91, 23) = "`Verkehrsmittel`"
 Str(0, 91, 24) = "`Linie`"
 Str(0, 91, 25) = "`Haltestelle Parkplätze`"
 Str(0, 91, 26) = "`Wegbeschreibung`"
 Str(0, 91, 27) = "`Entfernung zur Praxis`"
 Str(0, 91, 28) = "`Zahl`"
 Str(0, 91, 29) = "`nichtmehr`"
 Str(0, 91, 30) = "`Titel`"
 Str(0, 91, 31) = "`Geschlecht`"
 Str(0, 91, 32) = "`Straße`"
 Str(0, 91, 33) = "`PLZ`"
 Str(0, 91, 34) = "`Ort`"
 Str(0, 91, 35) = "`DMPT2`"
 Str(0, 91, 36) = "`DMPT1`"
 Str(0, 91, 37) = "`gelöscht`"
 Str(0, 91, 38) = "`ID`"
 Str(0, 91, 39) = "`PrimaryKey`"
 Str(0, 91, 40) = "`eindeut`"
 Str(0, 91, 41) = "`Auswahl`"
 Str(0, 91, 42) = "`KVNr`"
 ArtZ(0, 91) = 37
 ArtZ(1, 91) = 5
 Str(1, 91, 0) = "CREATE TABLE `hausaerztealt2` ("
 Str(1, 91, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 91, 2) = " `Überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '""L"" = Liebe(r), ""H"" = Hallo'"
 Str(1, 91, 3) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 4) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 5) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 6) = " `Anschrift` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 7) = " `KVNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 8) = " `Telefon` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 9) = " `Telefax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 10) = " `E_Mail` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 11) = " `Zulassungsgebiet` varchar(217) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 12) = " `Arzttyp` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 13) = " `Gemeinschaftspraxis mit` longtext COLLATE latin1_german2_ci"
 Str(1, 91, 14) = " `Schwerpunkt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 15) = " `Zusatzbezeichnung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 16) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 91, 17) = " `Beme` longtext COLLATE latin1_german2_ci"
 Str(1, 91, 18) = " `Sprechstunden` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 19) = " `von _ bis` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 20) = " `Internetadressen` longtext COLLATE latin1_german2_ci"
 Str(1, 91, 21) = " `Behandlung in Fremdsprachen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 22) = " `Rollstuhlgerechte Praxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 23) = " `Verkehrsmittel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 24) = " `Linie` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 25) = " `Haltestelle Parkplätze` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 26) = " `Wegbeschreibung` longtext COLLATE latin1_german2_ci"
 Str(1, 91, 27) = " `Entfernung zur Praxis` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 28) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 91, 29) = " `nichtmehr` bit(1) DEFAULT NULL COMMENT 'Arzt nicht mehr im Verzeichnis'"
 Str(1, 91, 30) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 31) = " `Geschlecht` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 32) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 33) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 34) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 91, 35) = " `DMPT2` smallint(6) DEFAULT NULL"
 Str(1, 91, 36) = " `DMPT1` bit(1) DEFAULT NULL"
 Str(1, 91, 37) = " `gelöscht` bit(1) DEFAULT NULL"
 Str(1, 91, 38) = "  PRIMARY KEY (`ID`)"
 Str(1, 91, 39) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 91, 40) = "  UNIQUE KEY `eindeut` (`Name`,`KVNr`,`DMPT2`,`DMPT1`,`gelöscht`,`Anschrift`,`Telefon`,`E_Mail`,`Beme`(30)) USING BTREE"
 Str(1, 91, 41) = "  KEY `Auswahl` (`KVNr`,`Name`)"
 Str(1, 91, 42) = "  KEY `KVNr` (`KVNr`,`Nachname`,`Vorname`)"
 Str(1, 91, 43) = " ENGINE=InnoDB AUTO_INCREMENT=108232 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr91

Sub FüllStr92()
 Str(0, 92, 0) = "inca"
 Str(0, 92, 1) = "`titel`"
 Str(0, 92, 2) = "`tsid`"
 Str(0, 92, 3) = "`transe`"
 Str(0, 92, 4) = "`transs`"
 Str(0, 92, 5) = "`id`"
 Str(0, 92, 6) = "`fsize`"
 Str(0, 92, 7) = "`pages`"
 Str(0, 92, 8) = "`devname`"
 Str(0, 92, 9) = "`retries`"
 Str(0, 92, 10) = "`csid`"
 Str(0, 92, 11) = "`routi`"
 Str(0, 92, 12) = "`callerid`"
 Str(0, 92, 13) = "`Id`"
 Str(0, 92, 14) = "`transe`"
 ArtZ(0, 92) = 12
 ArtZ(1, 92) = 2
 Str(1, 92, 0) = "CREATE TABLE `inca` ("
 Str(1, 92, 1) = " `titel` varchar(825) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 92, 2) = " `tsid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 92, 3) = " `transe` datetime DEFAULT NULL"
 Str(1, 92, 4) = " `transs` datetime DEFAULT NULL"
 Str(1, 92, 5) = " `id` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 92, 6) = " `fsize` int(10) DEFAULT NULL"
 Str(1, 92, 7) = " `pages` int(10) DEFAULT NULL"
 Str(1, 92, 8) = " `devname` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 92, 9) = " `retries` int(10) DEFAULT NULL"
 Str(1, 92, 10) = " `csid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 92, 11) = " `routi` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 92, 12) = " `callerid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 92, 13) = "  KEY `Id` (`id`)"
 Str(1, 92, 14) = "  KEY `transe` (`transe`)"
 Str(1, 92, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr92

Sub FüllStr93()
 Str(0, 93, 0) = "incq"
 Str(0, 93, 1) = "`titel`"
 Str(0, 93, 2) = "`tsid`"
 Str(0, 93, 3) = "`transe`"
 Str(0, 93, 4) = "`transs`"
 Str(0, 93, 5) = "`id`"
 Str(0, 93, 6) = "`fsize`"
 Str(0, 93, 7) = "`curp`"
 Str(0, 93, 8) = "`devid`"
 Str(0, 93, 9) = "`status`"
 Str(0, 93, 10) = "`exts`"
 Str(0, 93, 11) = "`extsc`"
 Str(0, 93, 12) = "`jobt`"
 Str(0, 93, 13) = "`retries`"
 Str(0, 93, 14) = "`routi`"
 Str(0, 93, 15) = "`csid`"
 Str(0, 93, 16) = "`avop`"
 Str(0, 93, 17) = "`callerid`"
 Str(0, 93, 18) = "`Id`"
 Str(0, 93, 19) = "`transs`"
 Str(0, 93, 20) = "`transe`"
 ArtZ(0, 93) = 17
 ArtZ(1, 93) = 3
 Str(1, 93, 0) = "CREATE TABLE `incq` ("
 Str(1, 93, 1) = " `titel` varchar(825) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 2) = " `tsid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 3) = " `transe` datetime DEFAULT NULL"
 Str(1, 93, 4) = " `transs` datetime DEFAULT NULL"
 Str(1, 93, 5) = " `id` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 6) = " `fsize` int(10) DEFAULT NULL"
 Str(1, 93, 7) = " `curp` int(10) DEFAULT NULL"
 Str(1, 93, 8) = " `devid` int(10) DEFAULT NULL"
 Str(1, 93, 9) = " `status` int(10) DEFAULT NULL"
 Str(1, 93, 10) = " `exts` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 11) = " `extsc` int(10) DEFAULT NULL"
 Str(1, 93, 12) = " `jobt` int(10) DEFAULT NULL"
 Str(1, 93, 13) = " `retries` int(10) DEFAULT NULL"
 Str(1, 93, 14) = " `routi` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 15) = " `csid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 16) = " `avop` int(10) DEFAULT NULL"
 Str(1, 93, 17) = " `callerid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 18) = "  KEY `Id` (`id`)"
 Str(1, 93, 19) = "  KEY `transs` (`transs`)"
 Str(1, 93, 20) = "  KEY `transe` (`transe`)"
 Str(1, 93, 21) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr93

Sub FüllStr94()
 Str(0, 94, 0) = "inl"
 Str(0, 94, 1) = "`id`"
 Str(0, 94, 2) = "`StByte`"
 Str(0, 94, 3) = "`breite`"
 Str(0, 94, 4) = "`kennung`"
 Str(0, 94, 5) = "`inhalt`"
 Str(0, 94, 6) = "`id`"
 Str(0, 94, 7) = "`stbyte`"
 Str(0, 94, 8) = "`kennung`"
 ArtZ(0, 94) = 5
 ArtZ(1, 94) = 3
 Str(1, 94, 0) = "CREATE TABLE `inl` ("
 Str(1, 94, 1) = " `id` int(15) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 94, 2) = " `StByte` int(10) unsigned NOT NULL"
 Str(1, 94, 3) = " `breite` char(3) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 94, 4) = " `kennung` char(4) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 94, 5) = " `inhalt` varchar(10000) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 94, 6) = "  PRIMARY KEY (`id`)"
 Str(1, 94, 7) = "  KEY `stbyte` (`StByte`) USING BTREE"
 Str(1, 94, 8) = "  KEY `kennung` (`kennung`,`inhalt`(6)) USING BTREE"
 Str(1, 94, 9) = " ENGINE=MyISAM AUTO_INCREMENT=51933093 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr94

Sub FüllStr95()
 Str(0, 95, 0) = "inlalt"
 Str(0, 95, 1) = "`id`"
 Str(0, 95, 2) = "`stbyte`"
 Str(0, 95, 3) = "`breite`"
 Str(0, 95, 4) = "`kennung`"
 Str(0, 95, 5) = "`inhalt`"
 Str(0, 95, 6) = "`id`"
 Str(0, 95, 7) = "`kennunginhalt`"
 ArtZ(0, 95) = 5
 ArtZ(1, 95) = 2
 Str(1, 95, 0) = "CREATE TABLE `inlalt` ("
 Str(1, 95, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 95, 2) = " `stbyte` int(10) DEFAULT NULL"
 Str(1, 95, 3) = " `breite` char(3) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 95, 4) = " `kennung` char(4) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 95, 5) = " `inhalt` varchar(10000) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 95, 6) = "  PRIMARY KEY (`id`)"
 Str(1, 95, 7) = "  KEY `kennunginhalt` (`kennung`,`inhalt`(6))"
 Str(1, 95, 8) = " ENGINE=MyISAM AUTO_INCREMENT=1364 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr95

Sub FüllStr96()
 Str(0, 96, 0) = "kassenliste"
 Str(0, 96, 1) = "`ID`"
 Str(0, 96, 2) = "`VK`"
 Str(0, 96, 3) = "`IK`"
 Str(0, 96, 4) = "`Name`"
 Str(0, 96, 5) = "`Kateg`"
 Str(0, 96, 6) = "`AnzahlIK`"
 Str(0, 96, 7) = "`AnzahlKTUG`"
 Str(0, 96, 8) = "`GültigVon`"
 Str(0, 96, 9) = "`GültigBis`"
 Str(0, 96, 10) = "`GO`"
 Str(0, 96, 11) = "`Kurzname`"
 Str(0, 96, 12) = "`rName`"
 Str(0, 96, 13) = "`Lantus2`"
 Str(0, 96, 14) = "`Levemir2`"
 Str(0, 96, 15) = "`Humalog`"
 Str(0, 96, 16) = "`Liprolog`"
 Str(0, 96, 17) = "`Novorapid`"
 Str(0, 96, 18) = "`Apidra`"
 Str(0, 96, 19) = "`ID`"
 Str(0, 96, 20) = "`PrimaryKey`"
 Str(0, 96, 21) = "`VK`"
 Str(0, 96, 22) = "`IK`"
 Str(0, 96, 23) = "`VKIK`"
 ArtZ(0, 96) = 18
 ArtZ(1, 96) = 5
 Str(1, 96, 0) = "CREATE TABLE `kassenliste` ("
 Str(1, 96, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 96, 2) = " `VK` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 96, 3) = " `IK` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 96, 4) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 96, 5) = " `Kateg` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kategorie'"
 Str(1, 96, 6) = " `AnzahlIK` int(4) unsigned DEFAULT NULL"
 Str(1, 96, 7) = " `AnzahlKTUG` int(4) unsigned DEFAULT NULL"
 Str(1, 96, 8) = " `GültigVon` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 96, 9) = " `GültigBis` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 96, 10) = " `GO` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 96, 11) = " `Kurzname` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 96, 12) = " `rName` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kurzname, falls nicht verfügbar: Name'"
 Str(1, 96, 13) = " `Lantus2` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Mehrwertvertrag für Lantus für Typ-2-D.m.'"
 Str(1, 96, 14) = " `Levemir2` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Mehrwertvertrag für Levemir für Typ-2-D.m.'"
 Str(1, 96, 15) = " `Humalog` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Rabattvertrag für Humalog'"
 Str(1, 96, 16) = " `Liprolog` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Rabattvertrag für Liprolog'"
 Str(1, 96, 17) = " `Novorapid` tinyint(7) unsigned NOT NULL DEFAULT '0' COMMENT 'Rabattvertrag für Novorapid'"
 Str(1, 96, 18) = " `Apidra` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Rabattvertrag für Apidra'"
 Str(1, 96, 19) = "  PRIMARY KEY (`ID`)"
 Str(1, 96, 20) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 96, 21) = "  KEY `VK` (`VK`)"
 Str(1, 96, 22) = "  KEY `IK` (`IK`)"
 Str(1, 96, 23) = "  KEY `VKIK` (`VK`,`IK`)"
 Str(1, 96, 24) = " ENGINE=InnoDB AUTO_INCREMENT=9424 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr96

Sub FüllStr97()
 Str(0, 97, 0) = "kheinweis"
 Str(0, 97, 1) = "`FID`"
 Str(0, 97, 2) = "`Pat_ID`"
 Str(0, 97, 3) = "`ZeitPunkt`"
 Str(0, 97, 4) = "`Ziel`"
 Str(0, 97, 5) = "`Diagnose`"
 Str(0, 97, 6) = "`absPos`"
 Str(0, 97, 7) = "`AktZeit`"
 Str(0, 97, 8) = "`StByte`"
 Str(0, 97, 9) = "`Auswahl`"
 Str(0, 97, 10) = "`FälleKHEinweis`"
 Str(0, 97, 11) = "`FID`"
 Str(0, 97, 12) = "`NamenKHEinweis`"
 Str(0, 97, 13) = "`F??lleKHEinweis_AccRel`"
 ArtZ(0, 97) = 8
 ArtZ(1, 97) = 4
 ArtZ(2, 97) = 1
 Str(1, 97, 0) = "CREATE TABLE `kheinweis` ("
 Str(1, 97, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 97, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 97, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 97, 4) = " `Ziel` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6291'"
 Str(1, 97, 5) = " `Diagnose` longtext COLLATE latin1_german2_ci COMMENT '6230'"
 Str(1, 97, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 97, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 97, 8) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 97, 9) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Ziel`)"
 Str(1, 97, 10) = "  KEY `FälleKHEinweis` (`FID`)"
 Str(1, 97, 11) = "  KEY `FID` (`FID`)"
 Str(1, 97, 12) = "  KEY `NamenKHEinweis` (`Pat_ID`)"
 Str(1, 97, 13) = "  CONSTRAINT `F??lleKHEinweis_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 97, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr97

Sub FüllStr98()
 Str(0, 98, 0) = "kontoausw"
 Str(1, 98, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `kontoausw` AS SELECT `e`.`eingid` AS `eingid`,`e`.`tabelle` AS `tabelle`,cast(`e`.`fdt` as date) AS `datum`,`e`.`datei` AS `datei`,if((left(`e`.`datei`,1) between '0' and '9'),left(`e`.`datei`,(locate('_',`e`.`datei`) - 1)),if((locate('_',`e`.`datei`,(locate('_',`e`.`datei`) + 1)) <> 0),substr(`e`.`datei`,(locate('_',`e`.`datei`) + 1),((locate('_',`e`.`datei`,(locate('_',`e`.`datei`) + 1)) - locate('_',`e`.`datei`)) - 1)),if((locate(' ',`e`.`datei`,(locate(' ',`e`.`datei`) + 1)) <> 0),substr(`e`.`datei`,(locate(' ',`e`.`datei`) + 1),((locate(' ',`e`.`datei`,(locate(' ',`e`.`datei`) + 1)) - locate(' ',`e`.`datei`)) - 1)),'?'))) AS `Konto` FROM `eingelesen` `e`"
End Sub ' FüllStr98

Sub FüllStr99()
 Str(0, 99, 0) = "kvnrue"
 Str(0, 99, 1) = "`lfdnr`"
 Str(0, 99, 2) = "`Pat_ID`"
 Str(0, 99, 3) = "`KVNr`"
 Str(0, 99, 4) = "`absPos`"
 Str(0, 99, 5) = "`AktZeit`"
 Str(0, 99, 6) = "`StByte`"
 Str(0, 99, 7) = "`lfdnr`"
 Str(0, 99, 8) = "`PrimaryKey`"
 Str(0, 99, 9) = "`zuord`"
 Str(0, 99, 10) = "`kvnruenamen`"
 ArtZ(0, 99) = 6
 ArtZ(1, 99) = 3
 ArtZ(2, 99) = 1
 Str(1, 99, 0) = "CREATE TABLE `kvnrue` ("
 Str(1, 99, 1) = " `lfdnr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 99, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 99, 3) = " `KVNr` varchar(9) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 99, 4) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 99, 5) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Zeit der Aktualisuerung aus der BDT-Datei'"
 Str(1, 99, 6) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 99, 7) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 99, 8) = "  UNIQUE KEY `PrimaryKey` (`lfdnr`)"
 Str(1, 99, 9) = "  UNIQUE KEY `zuord` (`Pat_ID`,`KVNr`)"
 Str(1, 99, 10) = "  CONSTRAINT `kvnruenamen` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 99, 11) = " ENGINE=InnoDB AUTO_INCREMENT=346347 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr99

Sub FüllStr100()
 Str(0, 100, 0) = "lGFR"
 Str(1, 100, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `lGFR` AS SELECT `n`.`Pat_ID` AS `pat_id`,if(((`l2`.`zeitpunkt` > `l1`.`ZeitPunkt`) or isnull(`l1`.`ZeitPunkt`)),`l2`.`zeitpunkt`,`l1`.`ZeitPunkt`) AS `lzp`,cast(replace(concat('0',trim(if(((`l2`.`zeitpunkt` > `l1`.`ZeitPunkt`) or isnull(`l1`.`ZeitPunkt`)),`l2`.`Wert`,`l1`.`Wert`))),',','.') as decimal(9,2)) AS `letzter` FROM ((`namen` `n` LEFT JOIN `labor2a` `l2` on(((`n`.`Pat_ID` = `l2`.`Pat_ID`) and (`l2`.`abk_ur` like '%gfr%') and (`l2`.`zeitpunkt` = (SELECT max(`labor2a`.`zeitpunkt`) FROM `labor2a` WHERE ((`labor2a`.`Pat_ID` = `n`.`Pat_ID`) and (`labor2a`.`abk_ur` like '%gfr%'))))))) LEFT JOIN `labor1a` `l1` on(((`n`.`Pat_ID` = `l1`.`Pat_ID`) and (`l1`.`abk_ur` like '%gfr%') and (`l1`.`ZeitPunkt` = (SELECT max(`labor1a`.`ZeitPunkt`) FROM `labor1a` WHERE ((`labor1a`.`Pat_ID` = `n`.`Pat_ID`) a" & _
  "nd (`labor1a`.`abk_ur` like '%gfr%'))))))) group by `n`.`Pat_ID`"
End Sub ' FüllStr100

Sub FüllStr101()
 Str(0, 101, 0) = "lHbA1c"
 Str(1, 101, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `lHbA1c` AS SELECT `n`.`Pat_ID` AS `pat_id`,if(((`l2`.`zeitpunkt` > `l1`.`ZeitPunkt`) or isnull(`l1`.`ZeitPunkt`)),`l2`.`zeitpunkt`,`l1`.`ZeitPunkt`) AS `lzp`,cast(replace(concat('0',trim(if(((`l2`.`zeitpunkt` > `l1`.`ZeitPunkt`) or isnull(`l1`.`ZeitPunkt`)),`l2`.`Wert`,`l1`.`Wert`))),',','.') as decimal(9,2)) AS `letzter` FROM ((`namen` `n` LEFT JOIN `labor2a` `l2` on(((`n`.`Pat_ID` = `l2`.`Pat_ID`) and (`l2`.`abk_ur` regexp '^hba[1c]') and (`l2`.`zeitpunkt` = (SELECT max(`labor2a`.`zeitpunkt`) FROM `labor2a` WHERE ((`labor2a`.`Pat_ID` = `n`.`Pat_ID`) and (`labor2a`.`abk_ur` regexp '^hba[1c]'))))))) LEFT JOIN `labor1a` `l1` on(((`n`.`Pat_ID` = `l1`.`Pat_ID`) and (`l1`.`abk_ur` regexp '^hba[1c]') and (`l1`.`ZeitPunkt` = (SELECT max(`labor1a`.`ZeitPunkt`) FROM `labor1a` WHERE ((`labor1a`.`Pat_ID` " & _
  "= `n`.`Pat_ID`) and (`labor1a`.`abk_ur` regexp '^hba[1c]'))))))) group by `n`.`Pat_ID`"
End Sub ' FüllStr101

Sub FüllStr102()
 Str(0, 102, 0) = "labor1"
 Str(1, 102, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `labor1` AS SELECT `n`.`Pat_ID` AS `Pat_ID`,cast(`n`.`ZeitPunkt` as date) AS `ZeitPunkt`,`n`.`FertigStGrad` AS `FertigStGrad`,if(isnull(`n2`.`Abkü`),`n`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`n`.`Abkü` AS `abk_ur`,`l`.`Langtext` AS `Langtext`,`n`.`Wert` AS `Wert`,if(isnull(`n2`.`Abkü`),`n`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`n`.`Einheit` AS `Einheit_ur`,`k`.`Kommentar` AS `Kommentar`,if(isnull(`n2`.`Abkü`),concat(if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`uNw`) or (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`oNw`) or (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`)),`nb2`.`NB`) AS `NB`,concat(if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`uNw`) or (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`oNw`) or (`p`.`oNw` = '')),`p`.`oN" & _
  "m`,`p`.`oNw`)) AS `NB_ur`,if(isnull(`n2`.`Abkü`),if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`uNw`) or (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),`nb2`.`uNg`) AS `uNg`,if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`uNw`) or (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`) AS `uNg_ur`,if(isnull(`n2`.`Abkü`),if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`oNw`) or (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`),`nb2`.`oNg`) AS `oNg`,if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`oNw`) or (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`) AS `oNg_ur`,'TM' AS `Labor` FROM (((((((`laborneu` `n` LEFT JOIN `namen` `na` on((`n`.`Pat_ID` = `na`.`Pat_ID`))) LEFT JOIN `laborlangtext` `l` on((`l`.`LangtextVW` = `n`.`LangtextVW`))) LEFT JOIN `laborkommentar` `k` on((`k`.`KommentarVW` = `n`.`KommentarVW`))) LEFT JOIN `laborparameter` `p` on(((`n`.`Abkü` = `p`.`Abkü`) and (`n`.`Einheit` = `p`.`Einheit`)))) LEFT JOIN `laborxpgl` `gl` on((`p`.`i" & _
  "d` = `gl`.`idpara`))) LEFT JOIN `laborxpneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) LEFT JOIN `laborxpnb` `nb2` on(((`n2`.`id` = `nb2`.`pid`) and (`nb2`.`Geschlecht` in (3,9,0,if((`na`.`Geschlecht` = 'm'),if(((`na`.`GebDat` + interval 18 year) > cast(`nb2`.`Eingang` as date)),1,4),if(((`na`.`GebDat` + interval 18 year) > cast(`nb2`.`Eingang` as date)),2,5))))))) WHERE ((`n`.`Wert` <> '') and (`n`.`Wert` is not null))"
End Sub ' FüllStr102

Sub FüllStr103()
 Str(0, 103, 0) = "labor1a"
 Str(1, 103, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `labor1a` AS SELECT `n`.`Pat_ID` AS `Pat_ID`,cast(`n`.`ZeitPunkt` as date) AS `ZeitPunkt`,`n`.`FertigStGrad` AS `FertigStGrad`,if(isnull(`n2`.`Abkü`),`n`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`n`.`Abkü` AS `abk_ur`,`l`.`Langtext` AS `Langtext`,`n`.`Wert` AS `Wert`,if(isnull(`n2`.`Abkü`),`n`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`n`.`Einheit` AS `Einheit_ur`,`k`.`Kommentar` AS `Kommentar`,if(isnull(`n2`.`Abkü`),concat(if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`uNw`) or (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`oNw`) or (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`)),`nb2`.`NB`) AS `NB`,concat(if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`uNw`) or (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`oNw`) or (`p`.`oNw` = '')),`p`.`o" & _
  "Nm`,`p`.`oNw`)) AS `NB_ur`,if(isnull(`n2`.`Abkü`),if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`uNw`) or (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),`nb2`.`uNg`) AS `uNg`,if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`uNw`) or (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`) AS `uNg_ur`,if(isnull(`n2`.`Abkü`),if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`oNw`) or (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`),`nb2`.`oNg`) AS `oNg`,if(((`na`.`Geschlecht` = 'm') or isnull(`p`.`oNw`) or (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`) AS `oNg_ur`,'TM' AS `Labor` FROM ((((((((`laborneu` `n` LEFT JOIN `namen` `na` on((`n`.`Pat_ID` = `na`.`Pat_ID`))) LEFT JOIN `laborlangtext` `l` on((`l`.`LangtextVW` = `n`.`LangtextVW`))) LEFT JOIN `laborkommentar` `k` on((`k`.`KommentarVW` = `n`.`KommentarVW`))) LEFT JOIN `laborparameter` `p` on(((`n`.`Abkü` = `p`.`Abkü`) and (`n`.`Einheit` = `p`.`Einheit`)))) LEFT JOIN `laborxpgl` `gl` on((`p`." & _
  "`id` = `gl`.`idpara`))) LEFT JOIN `laborxpneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) LEFT JOIN `laborxpnb` `nb2` on(((`n2`.`id` = `nb2`.`pid`) and (`nb2`.`Geschlecht` in (3,9,0,if((`na`.`Geschlecht` = 'm'),if(((`na`.`GebDat` + interval 18 year) > cast(`nb2`.`Eingang` as date)),1,4),if(((`na`.`GebDat` + interval 18 year) > cast(`nb2`.`Eingang` as date)),2,5))))))) LEFT JOIN `labor2a` `l2` on(((`l2`.`Pat_ID` = `n`.`Pat_ID`) and (`l2`.`abk_ur` = `n`.`Abkü`) and (`l2`.`FertigStGrad` = `n`.`FertigStGrad`) and ((`l2`.`Wert` = `n`.`Wert`) or (`l2`.`Kommentar` = `n`.`Wert`)) and (cast(`n`.`ZeitPunkt` as date) > (cast(`l2`.`zeitpunkt` as date) - interval 3 day)) and (cast(`n`.`ZeitPunkt` as date) < (cast(`l2`.`zeitpunkt` as date) + interval 6 day))))) WHERE ((`n`.`Wert` <> '') and (`n`.`Wert` is not null) and (isnull(`nb2`.`Eingang`) or (`nb2`.`Eingang` = (SELECT max(`laborxpnb`.`Eingang`) FROM " & _
  "`laborxpnb` WHERE (`laborxpnb`.`pid` = `n2`.`id`)))) and isnull(`l2`.`Pat_ID`))"
End Sub ' FüllStr103

Sub FüllStr104()
 Str(0, 104, 0) = "labor2a"
 Str(1, 104, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `labor2a` AS SELECT `u`.`Pat_id` AS `Pat_ID`,cast(`u`.`Eingang` as date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,if(isnull(`n2`.`Abkü`),`w`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`w`.`Abkü` AS `abk_ur`,`w`.`Langname` AS `Langtext`,`w`.`Wert` AS `Wert`,if(isnull(`n2`.`Abkü`),`w`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`w`.`Einheit` AS `Einheit_ur`,concat(if((`w`.`Erklärung` regexp '^:[ /*:]*$'),'',if((`w`.`Erklärung` regexp '^:[ /*]*:'),concat(substr(`w`.`Erklärung`,(locate(':',`w`.`Erklärung`,2) + 1)),';'),if((`w`.`Erklärung` = '.'),'',if((`w`.`Erklärung` = ''),'',concat(`w`.`Erklärung`,';'))))),`w`.`Kommentar`) AS `Kommentar`,`nb`.`NB` AS `NB`,`nb`.`NB` AS `NB_ur`,`nb`.`uNg` AS `uNg`,`nb`.`uNg` AS `uNG_ur`,if(((if(isnull(`n2`.`Abkü`),`w`.`Abkü`,`n2`.`Abkü`) = 'LDL') and (if(isnull(`n2`.`Abkü`),`w`.`" & _
  "Einheit`,`n2`.`Einheit`) = 'mg/dl')),'100',`nb`.`oNg`) AS `oNg`,`nb`.`oNg` AS `oNg_ur`,`s`.`Labor` AS `Labor` FROM (((((((`laborxus` `u` LEFT JOIN `namen` `na` on((`u`.`Pat_id` = `na`.`Pat_ID`))) LEFT JOIN `laborxwert` `w` on((`u`.`RefNr` = `w`.`RefNr`))) LEFT JOIN `laborxpnb` `nb` on((`w`.`nbid` = `nb`.`id`))) LEFT JOIN `laborxpneu` `neu` on((`nb`.`pid` = `neu`.`id`))) LEFT JOIN `laborxpgl` `gl` on((`neu`.`id` = `gl`.`idxpneu`))) LEFT JOIN `laborxpneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) LEFT JOIN `laborxsaetze` `s` on((`u`.`SatzID` = `s`.`SatzID`))) ORDER BY `u`.`Pat_id`,`u`.`Eingang`,`u`.`BefArt`"
End Sub ' FüllStr104

Sub FüllStr105()
 Str(0, 105, 0) = "labor2aNachweis"
 Str(1, 105, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `labor2aNachweis` AS SELECT `eg`.`Pfad` AS `pfad`,`u`.`Auftragsschlüssel` AS `auftragsschlüssel`,`u`.`Pat_id` AS `Pat_ID`,cast(`u`.`Eingang` as date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,if(isnull(`n2`.`Abkü`),`w`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`w`.`Abkü` AS `abk_ur`,`w`.`Langname` AS `Langtext`,`w`.`Wert` AS `Wert`,if(isnull(`n2`.`Abkü`),`w`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`w`.`Einheit` AS `Einheit_ur`,concat(if((`w`.`Erklärung` regexp '^:[ /*:]*$'),'',if((`w`.`Erklärung` regexp '^:[ /*]*:'),concat(substr(`w`.`Erklärung`,(locate(':',`w`.`Erklärung`,2) + 1)),';'),if((`w`.`Erklärung` = '.'),'',if((`w`.`Erklärung` = ''),'',concat(`w`.`Erklärung`,';'))))),`w`.`Kommentar`) AS `Kommentar`,if(isnull(`n2`.`Abkü`),`nb`.`NB`,`nb2`.`NB`) AS `NB`,`nb`.`NB` AS `NB_ur`,if(isnull(`n2`.`Abkü`),`n" & _
  "b`.`uNg`,`nb2`.`uNg`) AS `uNg`,`nb`.`uNg` AS `uNg_ur`,if(isnull(`n2`.`Abkü`),`nb`.`oNg`,`nb2`.`oNg`) AS `oNg`,`nb`.`oNg` AS `oNg_ur`,`s`.`Labor` AS `Labor` FROM (((((((((`laborxus` `u` LEFT JOIN `namen` `na` on((`u`.`Pat_id` = `na`.`Pat_ID`))) LEFT JOIN `laborxwert` `w` on((`u`.`RefNr` = `w`.`RefNr`))) LEFT JOIN `laborxpnb` `nb` on((`w`.`nbid` = `nb`.`id`))) LEFT JOIN `laborxpneu` `neu` on((`nb`.`pid` = `neu`.`id`))) LEFT JOIN `laborxpgl` `gl` on((`neu`.`id` = `gl`.`idxpneu`))) LEFT JOIN `laborxpneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) LEFT JOIN `laborxpnb` `nb2` on(((`n2`.`id` = `nb2`.`pid`) and (`nb2`.`Geschlecht` in (3,9,0,if((`na`.`Geschlecht` = 'm'),if(((`na`.`GebDat` + interval 18 year) > cast(`u`.`Eingang` as date)),1,4),if(((`na`.`GebDat` + interval 18 year) > cast(`u`.`Eingang` as date)),2,5))))))) LEFT JOIN `laborxsaetze` `s` on((`u`.`SatzID` = `s`.`SatzID`))) LEFT JOIN `lab" & _
  "orxeingel` `eg` on((`s`.`DatID` = `eg`.`DatID`))) WHERE ((isnull(`nb2`.`Eingang`) or (`nb2`.`Eingang` = (SELECT max(`laborxpnb`.`Eingang`) FROM `laborxpnb` WHERE (`laborxpnb`.`pid` = `n2`.`id`)))) and (`u`.`Pat_id` = 52287))"
End Sub ' FüllStr105

Sub FüllStr106()
 Str(0, 106, 0) = "labor_xls41674342743"
 Str(0, 106, 1) = "`id`"
 Str(0, 106, 2) = "`patient`"
 Str(0, 106, 3) = "`fehlerart`"
 Str(0, 106, 4) = "`id`"
 ArtZ(0, 106) = 3
 ArtZ(1, 106) = 1
 Str(1, 106, 0) = "CREATE TABLE `labor_xls41674342743` ("
 Str(1, 106, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 106, 2) = " `patient` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 106, 3) = " `fehlerart` varchar(3000) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 106, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 106, 5) = " ENGINE=InnoDB AUTO_INCREMENT=15 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr106

Sub FüllStr107()
 Str(0, 107, 0) = "labor_xls41674343877"
 Str(0, 107, 1) = "`id`"
 Str(0, 107, 2) = "`patient`"
 Str(0, 107, 3) = "`fehlerart`"
 Str(0, 107, 4) = "`id`"
 ArtZ(0, 107) = 3
 ArtZ(1, 107) = 1
 Str(1, 107, 0) = "CREATE TABLE `labor_xls41674343877` ("
 Str(1, 107, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 107, 2) = " `patient` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 3) = " `fehlerart` varchar(3000) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 107, 5) = " ENGINE=InnoDB AUTO_INCREMENT=15 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr107

Sub FüllStr108()
 Str(0, 108, 0) = "labor_xls41739330451"
 Str(0, 108, 1) = "`id`"
 Str(0, 108, 2) = "`patient`"
 Str(0, 108, 3) = "`fehlerart`"
 Str(0, 108, 4) = "`id`"
 ArtZ(0, 108) = 3
 ArtZ(1, 108) = 1
 Str(1, 108, 0) = "CREATE TABLE `labor_xls41739330451` ("
 Str(1, 108, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 108, 2) = " `patient` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 3) = " `fehlerart` varchar(3000) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 108, 5) = " ENGINE=InnoDB AUTO_INCREMENT=244 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr108

Sub FüllStr109()
 Str(0, 109, 0) = "labor_xls41752318703"
 Str(0, 109, 1) = "`id`"
 Str(0, 109, 2) = "`patient`"
 Str(0, 109, 3) = "`fehlerart`"
 Str(0, 109, 4) = "`id`"
 ArtZ(0, 109) = 3
 ArtZ(1, 109) = 1
 Str(1, 109, 0) = "CREATE TABLE `labor_xls41752318703` ("
 Str(1, 109, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 109, 2) = " `patient` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 109, 3) = " `fehlerart` varchar(3000) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 109, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 109, 5) = " ENGINE=InnoDB AUTO_INCREMENT=693 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr109

Sub FüllStr110()
 Str(0, 110, 0) = "labor_xls41880798287"
 Str(0, 110, 1) = "`id`"
 Str(0, 110, 2) = "`patient`"
 Str(0, 110, 3) = "`fehlerart`"
 Str(0, 110, 4) = "`id`"
 ArtZ(0, 110) = 3
 ArtZ(1, 110) = 1
 Str(1, 110, 0) = "CREATE TABLE `labor_xls41880798287` ("
 Str(1, 110, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 110, 2) = " `patient` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 110, 3) = " `fehlerart` varchar(3000) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 110, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 110, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr110

Sub FüllStr111()
 Str(0, 111, 0) = "labor_xls41880798831"
 Str(0, 111, 1) = "`id`"
 Str(0, 111, 2) = "`patient`"
 Str(0, 111, 3) = "`fehlerart`"
 Str(0, 111, 4) = "`id`"
 ArtZ(0, 111) = 3
 ArtZ(1, 111) = 1
 Str(1, 111, 0) = "CREATE TABLE `labor_xls41880798831` ("
 Str(1, 111, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 111, 2) = " `patient` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 111, 3) = " `fehlerart` varchar(3000) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 111, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 111, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr111

Sub FüllStr112()
 Str(0, 112, 0) = "labor_xls41880799814"
 Str(0, 112, 1) = "`id`"
 Str(0, 112, 2) = "`patient`"
 Str(0, 112, 3) = "`fehlerart`"
 Str(0, 112, 4) = "`id`"
 ArtZ(0, 112) = 3
 ArtZ(1, 112) = 1
 Str(1, 112, 0) = "CREATE TABLE `labor_xls41880799814` ("
 Str(1, 112, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 112, 2) = " `patient` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 112, 3) = " `fehlerart` varchar(3000) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 112, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 112, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr112

Sub FüllStr113()
 Str(0, 113, 0) = "laborgruppen"
 Str(0, 113, 1) = "`Laborgruppe`"
 Str(0, 113, 2) = "`Erklärung`"
 Str(0, 113, 3) = "`Laborgruppe`"
 Str(0, 113, 4) = "`Gruppe`"
 ArtZ(0, 113) = 2
 ArtZ(1, 113) = 2
 Str(1, 113, 0) = "CREATE TABLE `laborgruppen` ("
 Str(1, 113, 1) = " `Laborgruppe` int(10) NOT NULL DEFAULT '0'"
 Str(1, 113, 2) = " `Erklärung` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 113, 3) = "  PRIMARY KEY (`Laborgruppe`)"
 Str(1, 113, 4) = "  UNIQUE KEY `Gruppe` (`Laborgruppe`)"
 Str(1, 113, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr113

Sub FüllStr114()
 Str(0, 114, 0) = "laborkommentar"
 Str(0, 114, 1) = "`KommentarVW`"
 Str(0, 114, 2) = "`Kommentar`"
 Str(0, 114, 3) = "`KommentarVW`"
 Str(0, 114, 4) = "`KommentarVW`"
 Str(0, 114, 5) = "`Kommentar`"
 ArtZ(0, 114) = 2
 ArtZ(1, 114) = 3
 Str(1, 114, 0) = "CREATE TABLE `laborkommentar` ("
 Str(1, 114, 1) = " `KommentarVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 114, 2) = " `Kommentar` longtext COLLATE latin1_german2_ci"
 Str(1, 114, 3) = "  PRIMARY KEY (`KommentarVW`)"
 Str(1, 114, 4) = "  UNIQUE KEY `KommentarVW` (`KommentarVW`)"
 Str(1, 114, 5) = "  KEY `Kommentar` (`Kommentar`(255))"
 Str(1, 114, 6) = " ENGINE=InnoDB AUTO_INCREMENT=51887 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr114

Sub FüllStr115()
 Str(0, 115, 0) = "laborlangtext"
 Str(0, 115, 1) = "`LangtextVW`"
 Str(0, 115, 2) = "`Langtext`"
 Str(0, 115, 3) = "`LangtextVW`"
 Str(0, 115, 4) = "`LangtextVW`"
 Str(0, 115, 5) = "`Langtext`"
 ArtZ(0, 115) = 2
 ArtZ(1, 115) = 3
 Str(1, 115, 0) = "CREATE TABLE `laborlangtext` ("
 Str(1, 115, 1) = " `LangtextVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 115, 2) = " `Langtext` longtext COLLATE latin1_german2_ci"
 Str(1, 115, 3) = "  PRIMARY KEY (`LangtextVW`)"
 Str(1, 115, 4) = "  UNIQUE KEY `LangtextVW` (`LangtextVW`)"
 Str(1, 115, 5) = "  KEY `Langtext` (`Langtext`(255))"
 Str(1, 115, 6) = " ENGINE=InnoDB AUTO_INCREMENT=1532 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr115

Sub FüllStr116()
 Str(0, 116, 0) = "laborneu"
 Str(0, 116, 1) = "`FID`"
 Str(0, 116, 2) = "`Pat_ID`"
 Str(0, 116, 3) = "`ZeitPunkt`"
 Str(0, 116, 4) = "`FertigStGrad`"
 Str(0, 116, 5) = "`Abkü`"
 Str(0, 116, 6) = "`LangtextVW`"
 Str(0, 116, 7) = "`Wert`"
 Str(0, 116, 8) = "`Einheit`"
 Str(0, 116, 9) = "`KommentarVW`"
 Str(0, 116, 10) = "`AbsPos`"
 Str(0, 116, 11) = "`AktZeit`"
 Str(0, 116, 12) = "`Refnr`"
 Str(0, 116, 13) = "`StByte`"
 Str(0, 116, 14) = "`AbküWert`"
 Str(0, 116, 15) = "`Auswahl`"
 Str(0, 116, 16) = "`FälleLaborNeu`"
 Str(0, 116, 17) = "`LaborKommentarLaborNeu`"
 Str(0, 116, 18) = "`LaborLangtextLaborNeu`"
 Str(0, 116, 19) = "`LaborParameterLaborNeu`"
 Str(0, 116, 20) = "`NamenLaborNeu`"
 Str(0, 116, 21) = "`Prüf`"
 Str(0, 116, 22) = "`F??lleLaborNeu_AccRel`"
 Str(0, 116, 23) = "`LaborKommentarLaborNeu_AccRel`"
 Str(0, 116, 24) = "`LaborLangtextLaborNeu_AccRel`"
 Str(0, 116, 25) = "`LaborParameterLaborNeu_AccRel`"
 ArtZ(0, 116) = 13
 ArtZ(1, 116) = 8
 ArtZ(2, 116) = 4
 Str(1, 116, 0) = "CREATE TABLE `laborneu` ("
 Str(1, 116, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 116, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 116, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 116, 4) = " `FertigStGrad` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8401'"
 Str(1, 116, 5) = " `Abkü` varchar(64) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410'"
 Str(1, 116, 6) = " `LangtextVW` int(10) DEFAULT NULL COMMENT '8411'"
 Str(1, 116, 7) = " `Wert` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8420'"
 Str(1, 116, 8) = " `Einheit` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421'"
 Str(1, 116, 9) = " `KommentarVW` int(10) DEFAULT NULL COMMENT '8480'"
 Str(1, 116, 10) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 116, 11) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 116, 12) = " `Refnr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborXUS'"
 Str(1, 116, 13) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 116, 14) = "  KEY `AbküWert` (`Abkü`,`Wert`)"
 Str(1, 116, 15) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`FertigStGrad`,`Abkü`)"
 Str(1, 116, 16) = "  KEY `FälleLaborNeu` (`FID`)"
 Str(1, 116, 17) = "  KEY `LaborKommentarLaborNeu` (`KommentarVW`)"
 Str(1, 116, 18) = "  KEY `LaborLangtextLaborNeu` (`LangtextVW`)"
 Str(1, 116, 19) = "  KEY `LaborParameterLaborNeu` (`Abkü`,`Einheit`)"
 Str(1, 116, 20) = "  KEY `NamenLaborNeu` (`Pat_ID`)"
 Str(1, 116, 21) = "  KEY `Prüf` (`Pat_ID`,`Abkü`,`Wert`)"
 Str(1, 116, 22) = "  CONSTRAINT `F??lleLaborNeu_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 116, 23) = "  CONSTRAINT `LaborKommentarLaborNeu_AccRel` FOREIGN KEY (`KommentarVW`) REFERENCES `laborkommentar` (`KommentarVW`)"
 Str(1, 116, 24) = "  CONSTRAINT `LaborLangtextLaborNeu_AccRel` FOREIGN KEY (`LangtextVW`) REFERENCES `laborlangtext` (`LangtextVW`) ON UPDATE CASCADE"
 Str(1, 116, 25) = "  CONSTRAINT `LaborParameterLaborNeu_AccRel` FOREIGN KEY (`Abkü`, `Einheit`) REFERENCES `laborparameter` (`Abkü`, `Einheit`)"
 Str(1, 116, 26) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr116

Sub FüllStr117()
 Str(0, 117, 0) = "laborparameter"
 Str(0, 117, 1) = "`Abkü`"
 Str(0, 117, 2) = "`AbküN`"
 Str(0, 117, 3) = "`Labor`"
 Str(0, 117, 4) = "`Langtext`"
 Str(0, 117, 5) = "`Einheit`"
 Str(0, 117, 6) = "`Gruppe`"
 Str(0, 117, 7) = "`Reihe`"
 Str(0, 117, 8) = "`uNm`"
 Str(0, 117, 9) = "`oNm`"
 Str(0, 117, 10) = "`uNw`"
 Str(0, 117, 11) = "`oNw`"
 Str(0, 117, 12) = "`NB`"
 Str(0, 117, 13) = "`AktZeit`"
 Str(0, 117, 14) = "`StByte`"
 Str(0, 117, 15) = "`id`"
 Str(0, 117, 16) = "`id`"
 Str(0, 117, 17) = "`Abkü`"
 Str(0, 117, 18) = "`Fehlende`"
 Str(0, 117, 19) = "`LaborParameterAbkü`"
 Str(0, 117, 20) = "`Reihe`"
 Str(0, 117, 21) = "`LaborgruppenLaborParameter_AccRel`"
 ArtZ(0, 117) = 15
 ArtZ(1, 117) = 5
 ArtZ(2, 117) = 1
 Str(1, 117, 0) = "CREATE TABLE `laborparameter` ("
 Str(1, 117, 1) = " `Abkü` varchar(70) COLLATE latin1_german2_ci NOT NULL COMMENT '8410(1)'"
 Str(1, 117, 2) = " `AbküN` varchar(70) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Abkürzung mit gleicher Bedeutung, gleicher Einheit und gleichem Normbereich'"
 Str(1, 117, 3) = " `Labor` varchar(40) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT '8410(2)'"
 Str(1, 117, 4) = " `Langtext` varchar(60) COLLATE latin1_german2_ci NOT NULL COMMENT '8411'"
 Str(1, 117, 5) = " `Einheit` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421'"
 Str(1, 117, 6) = " `Gruppe` int(10) DEFAULT NULL COMMENT 'Ordnungsgruppe'"
 Str(1, 117, 7) = " `Reihe` int(10) DEFAULT NULL COMMENT 'Reihenfolge innerhalb der Gruppe'"
 Str(1, 117, 8) = " `uNm` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'unterer Normwert männlich'"
 Str(1, 117, 9) = " `oNm` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'oberer Normwert männlich'"
 Str(1, 117, 10) = " `uNw` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'unterer Normwert weiblich'"
 Str(1, 117, 11) = " `oNw` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'oberer Normwert weiblich'"
 Str(1, 117, 12) = " `NB` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Normbereich aus laborxwert'"
 Str(1, 117, 13) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 117, 14) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 117, 15) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 117, 16) = "  PRIMARY KEY (`id`)"
 Str(1, 117, 17) = "  UNIQUE KEY `Abkü` (`Abkü`,`Einheit`)"
 Str(1, 117, 18) = "  KEY `Fehlende` (`Gruppe`,`AktZeit`)"
 Str(1, 117, 19) = "  KEY `LaborParameterAbkü` (`Abkü`)"
 Str(1, 117, 20) = "  KEY `Reihe` (`Gruppe`,`Abkü`)"
 Str(1, 117, 21) = "  CONSTRAINT `LaborgruppenLaborParameter_AccRel` FOREIGN KEY (`Gruppe`) REFERENCES `laborgruppen` (`Laborgruppe`)"
 Str(1, 117, 22) = " ENGINE=InnoDB AUTO_INCREMENT=4499 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr117

Sub FüllStr118()
 Str(0, 118, 0) = "laborparameteralt"
 Str(0, 118, 1) = "`Abkü`"
 Str(0, 118, 2) = "`AbküN`"
 Str(0, 118, 3) = "`Labor`"
 Str(0, 118, 4) = "`Langtext`"
 Str(0, 118, 5) = "`Einheit`"
 Str(0, 118, 6) = "`Gruppe`"
 Str(0, 118, 7) = "`Reihe`"
 Str(0, 118, 8) = "`uNm`"
 Str(0, 118, 9) = "`oNm`"
 Str(0, 118, 10) = "`uNw`"
 Str(0, 118, 11) = "`oNw`"
 Str(0, 118, 12) = "`NB`"
 Str(0, 118, 13) = "`AktZeit`"
 Str(0, 118, 14) = "`StByte`"
 Str(0, 118, 15) = "`id`"
 Str(0, 118, 16) = "`id`"
 Str(0, 118, 17) = "`Abkü`"
 Str(0, 118, 18) = "`Fehlende`"
 Str(0, 118, 19) = "`LaborParameterAbkü`"
 Str(0, 118, 20) = "`Reihe`"
 ArtZ(0, 118) = 15
 ArtZ(1, 118) = 5
 Str(1, 118, 0) = "CREATE TABLE `laborparameteralt` ("
 Str(1, 118, 1) = " `Abkü` varchar(70) COLLATE latin1_german2_ci NOT NULL COMMENT '8410(1)'"
 Str(1, 118, 2) = " `AbküN` varchar(70) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Abkürzung mit gleicher Bedeutung, gleicher Einheit und gleichem Normbereich'"
 Str(1, 118, 3) = " `Labor` varchar(40) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT '8410(2)'"
 Str(1, 118, 4) = " `Langtext` varchar(60) COLLATE latin1_german2_ci NOT NULL COMMENT '8411'"
 Str(1, 118, 5) = " `Einheit` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421'"
 Str(1, 118, 6) = " `Gruppe` int(10) DEFAULT NULL COMMENT 'Ordnungsgruppe'"
 Str(1, 118, 7) = " `Reihe` int(10) DEFAULT NULL COMMENT 'Reihenfolge innerhalb der Gruppe'"
 Str(1, 118, 8) = " `uNm` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'unterer Normwert männlich'"
 Str(1, 118, 9) = " `oNm` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'oberer Normwert männlich'"
 Str(1, 118, 10) = " `uNw` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'unterer Normwert weiblich'"
 Str(1, 118, 11) = " `oNw` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'oberer Normwert weiblich'"
 Str(1, 118, 12) = " `NB` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Normbereich aus laborxwert'"
 Str(1, 118, 13) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 118, 14) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 118, 15) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 118, 16) = "  PRIMARY KEY (`id`)"
 Str(1, 118, 17) = "  UNIQUE KEY `Abkü` (`Abkü`,`Einheit`)"
 Str(1, 118, 18) = "  KEY `Fehlende` (`Gruppe`,`AktZeit`)"
 Str(1, 118, 19) = "  KEY `LaborParameterAbkü` (`Abkü`)"
 Str(1, 118, 20) = "  KEY `Reihe` (`Gruppe`,`Abkü`)"
 Str(1, 118, 21) = " ENGINE=InnoDB AUTO_INCREMENT=4237 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr118

Sub FüllStr119()
 Str(0, 119, 0) = "laborxbakt"
 Str(0, 119, 1) = "`RefNr`"
 Str(0, 119, 2) = "`Verf`"
 Str(0, 119, 3) = "`KuQu`"
 Str(0, 119, 4) = "`Quelle`"
 Str(0, 119, 5) = "`QSpez`"
 Str(0, 119, 6) = "`AbnDat`"
 Str(0, 119, 7) = "`Kommentar`"
 Str(0, 119, 8) = "`Erklärung`"
 Str(0, 119, 9) = "`Keimzahl`"
 Str(0, 119, 10) = "`LaborXUSLaborXBakt`"
 Str(0, 119, 11) = "`LaborXUSLaborXBakt_AccRel`"
 ArtZ(0, 119) = 9
 ArtZ(1, 119) = 1
 ArtZ(2, 119) = 1
 Str(1, 119, 0) = "CREATE TABLE `laborxbakt` ("
 Str(1, 119, 1) = " `RefNr` int(10) DEFAULT NULL"
 Str(1, 119, 2) = " `Verf` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 119, 3) = " `KuQu` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8428 Probenmaterial-Ident (Turbomed)'"
 Str(1, 119, 4) = " `Quelle` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8430 Probenmaterial-Bezeichnung (Turbomed)'"
 Str(1, 119, 5) = " `QSpez` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8431 Probenmaterial-Spezifikation (Turbomed)'"
 Str(1, 119, 6) = " `AbnDat` datetime DEFAULT NULL COMMENT '8432 Abnahmedatum (Turbomed)'"
 Str(1, 119, 7) = " `Kommentar` longtext COLLATE latin1_german2_ci COMMENT '8480 Ergebnistest (Turbomed)'"
 Str(1, 119, 8) = " `Erklärung` longtext COLLATE latin1_german2_ci"
 Str(1, 119, 9) = " `Keimzahl` varchar(36) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 119, 10) = "  KEY `LaborXUSLaborXBakt` (`RefNr`)"
 Str(1, 119, 11) = "  CONSTRAINT `LaborXUSLaborXBakt_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 119, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr119

Sub FüllStr120()
 Str(0, 120, 0) = "laborxeingel"
 Str(0, 120, 1) = "`DatID`"
 Str(0, 120, 2) = "`Pfad`"
 Str(0, 120, 3) = "`Name`"
 Str(0, 120, 4) = "`Zp`"
 Str(0, 120, 5) = "`fertig`"
 Str(0, 120, 6) = "`DatID`"
 Str(0, 120, 7) = "`NamePfad`"
 ArtZ(0, 120) = 5
 ArtZ(1, 120) = 2
 Str(1, 120, 0) = "CREATE TABLE `laborxeingel` ("
 Str(1, 120, 1) = " `DatID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf LaborEingelesen'"
 Str(1, 120, 2) = " `Pfad` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Pfadname'"
 Str(1, 120, 3) = " `Name` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name der eingelesenen Labordatei ohne Endung'"
 Str(1, 120, 4) = " `Zp` datetime DEFAULT NULL COMMENT 'Einlesezeitpunkt'"
 Str(1, 120, 5) = " `fertig` bit(1) DEFAULT NULL COMMENT 'ob Einlesen fertig'"
 Str(1, 120, 6) = "  PRIMARY KEY (`DatID`)"
 Str(1, 120, 7) = "  KEY `NamePfad` (`Name`,`Pfad`)"
 Str(1, 120, 8) = " ENGINE=InnoDB AUTO_INCREMENT=4131 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr120

Sub FüllStr121()
 Str(0, 121, 0) = "laborxleist"
 Str(0, 121, 1) = "`RefNr`"
 Str(0, 121, 2) = "`Abkü`"
 Str(0, 121, 3) = "`Verf`"
 Str(0, 121, 4) = "`EBM`"
 Str(0, 121, 5) = "`goä`"
 Str(0, 121, 6) = "`Anzahl`"
 Str(0, 121, 7) = "`abrd`"
 Str(0, 121, 8) = "`LaborXUSLaborXLeist`"
 Str(0, 121, 9) = "`LaborXUSLaborXLeist_AccRel`"
 ArtZ(0, 121) = 7
 ArtZ(1, 121) = 1
 ArtZ(2, 121) = 1
 Str(1, 121, 0) = "CREATE TABLE `laborxleist` ("
 Str(1, 121, 1) = " `RefNr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborUS'"
 Str(1, 121, 2) = " `Abkü` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410 Test-Ident (Turbomed)'"
 Str(1, 121, 3) = " `Verf` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8434'"
 Str(1, 121, 4) = " `EBM` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5001 GNR (Turbomed)'"
 Str(1, 121, 5) = " `goä` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8406'"
 Str(1, 121, 6) = " `Anzahl` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5005'"
 Str(1, 121, 7) = " `abrd` varchar(1) COLLATE latin1_german2_ci NOT NULL COMMENT '8614 Abrechnung durch: 1 = Labor, 2 = Einweiser'"
 Str(1, 121, 8) = "  KEY `LaborXUSLaborXLeist` (`RefNr`) USING BTREE"
 Str(1, 121, 9) = "  CONSTRAINT `LaborXUSLaborXLeist_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 121, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr121

Sub FüllStr122()
 Str(0, 122, 0) = "laborxpgl"
 Str(0, 122, 1) = "`id`"
 Str(0, 122, 2) = "`idxpneu`"
 Str(0, 122, 3) = "`idpara`"
 Str(0, 122, 4) = "`idxpbez`"
 Str(0, 122, 5) = "`ergänzt`"
 Str(0, 122, 6) = "`id`"
 Str(0, 122, 7) = "`idxpneu`"
 Str(0, 122, 8) = "`idpara`"
 Str(0, 122, 9) = "`idxpbez`"
 Str(0, 122, 10) = "`FK_laborxpgl_1`"
 Str(0, 122, 11) = "`FK_laborxpgl_2`"
 Str(0, 122, 12) = "`FK_laborxpgl_3`"
 ArtZ(0, 122) = 5
 ArtZ(1, 122) = 4
 ArtZ(2, 122) = 3
 Str(1, 122, 0) = "CREATE TABLE `laborxpgl` ("
 Str(1, 122, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 122, 2) = " `idxpneu` int(10) unsigned DEFAULT NULL COMMENT 'laborxpneu'"
 Str(1, 122, 3) = " `idpara` int(10) unsigned DEFAULT NULL COMMENT 'laborparameter'"
 Str(1, 122, 4) = " `idxpbez` int(10) unsigned NOT NULL COMMENT '.. ist identisch mit laborxpneu'"
 Str(1, 122, 5) = " `ergänzt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Ergänzung'"
 Str(1, 122, 6) = "  PRIMARY KEY (`id`)"
 Str(1, 122, 7) = "  KEY `idxpneu` (`idxpneu`)"
 Str(1, 122, 8) = "  KEY `idpara` (`idpara`)"
 Str(1, 122, 9) = "  KEY `idxpbez` (`idxpbez`)"
 Str(1, 122, 10) = "  CONSTRAINT `FK_laborxpgl_1` FOREIGN KEY (`idxpneu`) REFERENCES `laborxpneu` (`id`)"
 Str(1, 122, 11) = "  CONSTRAINT `FK_laborxpgl_2` FOREIGN KEY (`idpara`) REFERENCES `laborparameter` (`id`)"
 Str(1, 122, 12) = "  CONSTRAINT `FK_laborxpgl_3` FOREIGN KEY (`idxpbez`) REFERENCES `laborxpneu` (`id`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 122, 13) = " ENGINE=InnoDB AUTO_INCREMENT=1636 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Parametergleichheiten'"
End Sub ' FüllStr122

Sub FüllStr123()
 Str(0, 123, 0) = "laborxplab"
 Str(0, 123, 1) = "`id`"
 Str(0, 123, 2) = "`Labor`"
 Str(0, 123, 3) = "`id`"
 Str(0, 123, 4) = "`Labor`"
 ArtZ(0, 123) = 2
 ArtZ(1, 123) = 2
 Str(1, 123, 0) = "CREATE TABLE `laborxplab` ("
 Str(1, 123, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'eindeutige Kennung'"
 Str(1, 123, 2) = " `Labor` varchar(36) COLLATE latin1_german2_ci NOT NULL COMMENT '8300 (maximale Länge: 36)'"
 Str(1, 123, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 123, 4) = "  UNIQUE KEY `Labor` (`Labor`) USING BTREE"
 Str(1, 123, 5) = " ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr123

Sub FüllStr124()
 Str(0, 124, 0) = "laborxpnb"
 Str(0, 124, 1) = "`id`"
 Str(0, 124, 2) = "`pid`"
 Str(0, 124, 3) = "`Geschlecht`"
 Str(0, 124, 4) = "`Eingang`"
 Str(0, 124, 5) = "`uNg`"
 Str(0, 124, 6) = "`oNg`"
 Str(0, 124, 7) = "`NB`"
 Str(0, 124, 8) = "`zahl`"
 Str(0, 124, 9) = "`uid`"
 Str(0, 124, 10) = "`id`"
 Str(0, 124, 11) = "`pid`"
 Str(0, 124, 12) = "`laborxpneu_xpnb`"
 ArtZ(0, 124) = 9
 ArtZ(1, 124) = 2
 ArtZ(2, 124) = 1
 Str(1, 124, 0) = "CREATE TABLE `laborxpnb` ("
 Str(1, 124, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'eindeutige Kennung'"
 Str(1, 124, 2) = " `pid` int(10) unsigned NOT NULL COMMENT 'Bezug auf laborxpneu'"
 Str(1, 124, 3) = " `Geschlecht` int(1) unsigned NOT NULL COMMENT '1=Mann, 2=Frau, 3=unbek, 4=Knabe, 5=Mädchen, 0=Name fehlt, 9=beide'"
 Str(1, 124, 4) = " `Eingang` datetime DEFAULT NULL COMMENT 'Eingangsdatum im Labor'"
 Str(1, 124, 5) = " `uNg` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'untere Normgrenze'"
 Str(1, 124, 6) = " `oNg` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'obere  Normgrenze'"
 Str(1, 124, 7) = " `NB` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Normbereich'"
 Str(1, 124, 8) = " `zahl` int(10) unsigned NOT NULL COMMENT 'Häufigkeit eines Laborparameters'"
 Str(1, 124, 9) = " `uid` int(10) unsigned NOT NULL COMMENT 'laborxus.id des ersten Eintrags'"
 Str(1, 124, 10) = "  PRIMARY KEY (`id`)"
 Str(1, 124, 11) = "  KEY `pid` (`pid`)"
 Str(1, 124, 12) = "  CONSTRAINT `laborxpneu_xpnb` FOREIGN KEY (`pid`) REFERENCES `laborxpneu` (`id`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 124, 13) = " ENGINE=InnoDB AUTO_INCREMENT=2506 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Normbereiche'"
End Sub ' FüllStr124

Sub FüllStr125()
 Str(0, 125, 0) = "laborxpneu"
 Str(0, 125, 1) = "`id`"
 Str(0, 125, 2) = "`Abkü`"
 Str(0, 125, 3) = "`Langtext`"
 Str(0, 125, 4) = "`Einheit`"
 Str(0, 125, 5) = "`lid`"
 Str(0, 125, 6) = "`Gruppe`"
 Str(0, 125, 7) = "`id`"
 Str(0, 125, 8) = "`Abkü`"
 Str(0, 125, 9) = "`Labore`"
 Str(0, 125, 10) = "`Gruppe`"
 Str(0, 125, 11) = "`Gruppe`"
 Str(0, 125, 12) = "`Labore`"
 ArtZ(0, 125) = 6
 ArtZ(1, 125) = 4
 ArtZ(2, 125) = 2
 Str(1, 125, 0) = "CREATE TABLE `laborxpneu` ("
 Str(1, 125, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'eindeutige Kennung'"
 Str(1, 125, 2) = " `Abkü` varchar(8) COLLATE latin1_german2_ci NOT NULL COMMENT '8410 (maximale Länge: 8)'"
 Str(1, 125, 3) = " `Langtext` varchar(40) COLLATE latin1_german2_ci NOT NULL COMMENT '8411 (maximale Länge: 40)'"
 Str(1, 125, 4) = " `Einheit` varchar(12) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421 (maximale Länge: 12)'"
 Str(1, 125, 5) = " `lid` int(10) unsigned NOT NULL COMMENT 'Bezug auf laborxplab.id'"
 Str(1, 125, 6) = " `Gruppe` int(10) DEFAULT NULL COMMENT 'Bezug auf laborgruppen.laborgruppe'"
 Str(1, 125, 7) = "  PRIMARY KEY (`id`)"
 Str(1, 125, 8) = "  UNIQUE KEY `Abkü` (`Abkü`,`Einheit`,`lid`) USING BTREE"
 Str(1, 125, 9) = "  KEY `Labore` (`lid`)"
 Str(1, 125, 10) = "  KEY `Gruppe` (`Gruppe`)"
 Str(1, 125, 11) = "  CONSTRAINT `Gruppe` FOREIGN KEY (`Gruppe`) REFERENCES `laborgruppen` (`Laborgruppe`)"
 Str(1, 125, 12) = "  CONSTRAINT `Labore` FOREIGN KEY (`lid`) REFERENCES `laborxplab` (`id`)"
 Str(1, 125, 13) = " ENGINE=InnoDB AUTO_INCREMENT=1828 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Langtexte, Einheiten und Gruppen zu den Verfahren'"
End Sub ' FüllStr125

Sub FüllStr126()
 Str(0, 126, 0) = "laborxsaetze"
 Str(0, 126, 1) = "`SatzID`"
 Str(0, 126, 2) = "`DatID`"
 Str(0, 126, 3) = "`Satzart`"
 Str(0, 126, 4) = "`Satzlänge`"
 Str(0, 126, 5) = "`SatzlängeSchluss`"
 Str(0, 126, 6) = "`VersionSatzb`"
 Str(0, 126, 7) = "`Arztnr`"
 Str(0, 126, 8) = "`Arztname`"
 Str(0, 126, 9) = "`StraßePraxis`"
 Str(0, 126, 10) = "`Arzt`"
 Str(0, 126, 11) = "`LANR`"
 Str(0, 126, 12) = "`PLZPraxis`"
 Str(0, 126, 13) = "`OrtPraxis`"
 Str(0, 126, 14) = "`Labor`"
 Str(0, 126, 15) = "`StraßeLabor`"
 Str(0, 126, 16) = "`PLZLabor`"
 Str(0, 126, 17) = "`OrtLabor`"
 Str(0, 126, 18) = "`KBVPrüfnr`"
 Str(0, 126, 19) = "`Zeichensatz`"
 Str(0, 126, 20) = "`Kundenarztnr`"
 Str(0, 126, 21) = "`Erstellungsdatum`"
 Str(0, 126, 22) = "`Gesamtlänge`"
 Str(0, 126, 23) = "`SatzID`"
 Str(0, 126, 24) = "`DatID`"
 Str(0, 126, 25) = "`Name`"
 Str(0, 126, 26) = "`laborsaetzedateien`"
 ArtZ(0, 126) = 22
 ArtZ(1, 126) = 3
 ArtZ(2, 126) = 1
 Str(1, 126, 0) = "CREATE TABLE `laborxsaetze` ("
 Str(1, 126, 1) = " `SatzID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'zum Bezug für LaborUS'"
 Str(1, 126, 2) = " `DatID` int(10) DEFAULT NULL COMMENT 'Bezug zu LaborEingelesen'"
 Str(1, 126, 3) = " `Satzart` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000 Satzart (Turbomed)'"
 Str(1, 126, 4) = " `Satzlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed)'"
 Str(1, 126, 5) = " `SatzlängeSchluss` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000'"
 Str(1, 126, 6) = " `VersionSatzb` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9212 Version der Satzbeschreibung (Turbomed)'"
 Str(1, 126, 7) = " `Arztnr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '201 Arztnummer (Turbomed)'"
 Str(1, 126, 8) = " `Arztname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '203 Arztname (Turbomed)'"
 Str(1, 126, 9) = " `StraßePraxis` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '205 Straße der Praxis (Turbomed)'"
 Str(1, 126, 10) = " `Arzt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ' 211 Ausführender Arzt'"
 Str(1, 126, 11) = " `LANR` varchar(11) COLLATE latin1_german2_ci NOT NULL COMMENT ' 212 LANR'"
 Str(1, 126, 12) = " `PLZPraxis` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '215 PLZ der Praxis (Turbomed)'"
 Str(1, 126, 13) = " `OrtPraxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '216 Ort der Praxis (Turbomed)'"
 Str(1, 126, 14) = " `Labor` varchar(36) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8320 Labor'"
 Str(1, 126, 15) = " `StraßeLabor` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8321 Straße der Laboradresse (Turbomed)'"
 Str(1, 126, 16) = " `PLZLabor` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8322 PLZ der Laboradresse (Turbomed)'"
 Str(1, 126, 17) = " `OrtLabor` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8323 Ort der Laboradresse (Turbomed)'"
 Str(1, 126, 18) = " `KBVPrüfnr` varchar(16) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '101 KBV-Prüfnummer (Turbomed)'"
 Str(1, 126, 19) = " `Zeichensatz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9106 verwendeter Zeichensatz (Turbomed)'"
 Str(1, 126, 20) = " `Kundenarztnr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8312 Kundenarztnummer (Turbomed)'"
 Str(1, 126, 21) = " `Erstellungsdatum` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9103 Erstellungsdatum (Turbomed)'"
 Str(1, 126, 22) = " `Gesamtlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9202 Gesamtlänge des Datenpaketes (Turbomed)'"
 Str(1, 126, 23) = "  PRIMARY KEY (`SatzID`)"
 Str(1, 126, 24) = "  KEY `DatID` (`DatID`)"
 Str(1, 126, 25) = "  KEY `Name` (`PLZLabor`,`OrtLabor`)"
 Str(1, 126, 26) = "  CONSTRAINT `laborsaetzedateien` FOREIGN KEY (`DatID`) REFERENCES `laborxeingel` (`DatID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 126, 27) = " ENGINE=InnoDB AUTO_INCREMENT=9187 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr126

Sub FüllStr127()
 Str(0, 127, 0) = "laborxus"
 Str(0, 127, 1) = "`RefNr`"
 Str(0, 127, 2) = "`DatID`"
 Str(0, 127, 3) = "`SatzID`"
 Str(0, 127, 4) = "`Satzart`"
 Str(0, 127, 5) = "`Satzlänge`"
 Str(0, 127, 6) = "`Auftragsnummer`"
 Str(0, 127, 7) = "`Auftragsschlüssel`"
 Str(0, 127, 8) = "`Eingang`"
 Str(0, 127, 9) = "`Berichtsdatum`"
 Str(0, 127, 10) = "`Pat_id`"
 Str(0, 127, 11) = "`Nachname`"
 Str(0, 127, 12) = "`Vorname`"
 Str(0, 127, 13) = "`GebDat`"
 Str(0, 127, 14) = "`Titel`"
 Str(0, 127, 15) = "`NVorsatz`"
 Str(0, 127, 16) = "`BefArt`"
 Str(0, 127, 17) = "`Abrechnungstyp`"
 Str(0, 127, 18) = "`GebüOrd`"
 Str(0, 127, 19) = "`Auftraggeber`"
 Str(0, 127, 20) = "`Patienteninformation`"
 Str(0, 127, 21) = "`Geschlecht`"
 Str(0, 127, 22) = "`AuftrHinw`"
 Str(0, 127, 23) = "`Pat_idUrsp`"
 Str(0, 127, 24) = "`Pat_idErwVNG`"
 Str(0, 127, 25) = "`Pat_idErwVN`"
 Str(0, 127, 26) = "`Pat_idErwG`"
 Str(0, 127, 27) = "`Pat_idErwGB`"
 Str(0, 127, 28) = "`Pat_idErwGL`"
 Str(0, 127, 29) = "`Pat_idLaborNeu`"
 Str(0, 127, 30) = "`ZeitpunktLaborneu`"
 Str(0, 127, 31) = "`ZdüP`"
 Str(0, 127, 32) = "`ZdiP`"
 Str(0, 127, 33) = "`LWerte`"
 Str(0, 127, 34) = "`verglichen`"
 Str(0, 127, 35) = "`AfN`"
 Str(0, 127, 36) = "`RefNr`"
 Str(0, 127, 37) = "`LaborXEingelLaborXUS`"
 Str(0, 127, 38) = "`LaborXSätzeLaborXUS`"
 Str(0, 127, 39) = "`Name`"
 Str(0, 127, 40) = "`Pat_id`"
 Str(0, 127, 41) = "`LaborXEingelLaborXUS_AccRel`"
 Str(0, 127, 42) = "`LaborXS??tzeLaborXUS_AccRel`"
 ArtZ(0, 127) = 35
 ArtZ(1, 127) = 5
 ArtZ(2, 127) = 2
 Str(1, 127, 0) = "CREATE TABLE `laborxus` ("
 Str(1, 127, 1) = " `RefNr` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf LaborWert'"
 Str(1, 127, 2) = " `DatID` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborEingelesen'"
 Str(1, 127, 3) = " `SatzID` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborXSätze'"
 Str(1, 127, 4) = " `Satzart` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000 Satzart (Turbomed)'"
 Str(1, 127, 5) = " `Satzlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed)'"
 Str(1, 127, 6) = " `Auftragsnummer` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8310 Anforderungsident (Turbomed)'"
 Str(1, 127, 7) = " `Auftragsschlüssel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8311 Anforderungsnr d Labors (Turbomed)'"
 Str(1, 127, 8) = " `Eingang` datetime DEFAULT NULL COMMENT '8301 Eingangsdatum in Datumsform'"
 Str(1, 127, 9) = " `Berichtsdatum` varchar(13) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8302 Berichtsdatum'"
 Str(1, 127, 10) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 127, 11) = " `Nachname` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 127, 12) = " `Vorname` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 127, 13) = " `GebDat` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3103'"
 Str(1, 127, 14) = " `Titel` varchar(12) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3104'"
 Str(1, 127, 15) = " `NVorsatz` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3100'"
 Str(1, 127, 16) = " `BefArt` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8401 Befundart (Turbomed) / Fertigstellungsgrad (""E""=Endbefund, ""T"" = Teilbefund)'"
 Str(1, 127, 17) = " `Abrechnungstyp` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)'"
 Str(1, 127, 18) = " `GebüOrd` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8403 Gebührenordnung (Turbomed)'"
 Str(1, 127, 19) = " `Auftraggeber` varchar(10) COLLATE latin1_german2_ci NOT NULL COMMENT '8615 Auftraggeber (LANR)'"
 Str(1, 127, 20) = " `Patienteninformation` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8405 Patienteninformation (Turbomed)'"
 Str(1, 127, 21) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8407 Geschlecht (Turbomed)'"
 Str(1, 127, 22) = " `AuftrHinw` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8490 Auftragsbezogene Hinweise (Turbomed)'"
 Str(1, 127, 23) = " `Pat_idUrsp` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor'"
 Str(1, 127, 24) = " `Pat_idErwVNG` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag'"
 Str(1, 127, 25) = " `Pat_idErwVN` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Vornamen und Nachnamen'"
 Str(1, 127, 26) = " `Pat_idErwG` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag'"
 Str(1, 127, 27) = " `Pat_idErwGB` varchar(34) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung'"
 Str(1, 127, 28) = " `Pat_idErwGL` varchar(14) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor'"
 Str(1, 127, 29) = " `Pat_idLaborNeu` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Pat_ids von in Laborneu zuordnbaren Patienten'"
 Str(1, 127, 30) = " `ZeitpunktLaborneu` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde'"
 Str(1, 127, 31) = " `ZdüP` smallint(6) DEFAULT NULL COMMENT 'Zahl der verglichenen Parameter'"
 Str(1, 127, 32) = " `ZdiP` int(10) DEFAULT NULL COMMENT 'Zahl der infragekommenden Patienten'"
 Str(1, 127, 33) = " `LWerte` longtext COLLATE latin1_german2_ci COMMENT 'Laborwerte, die zur Zuordnung geführt haben'"
 Str(1, 127, 34) = " `verglichen` datetime DEFAULT NULL COMMENT 'Datum, zu dem Datensatz zuletzt verglichen wurde'"
 Str(1, 127, 35) = " `AfN` smallint(6) DEFAULT NULL COMMENT 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu'"
 Str(1, 127, 36) = "  PRIMARY KEY (`RefNr`)"
 Str(1, 127, 37) = "  KEY `LaborXEingelLaborXUS` (`DatID`)"
 Str(1, 127, 38) = "  KEY `LaborXSätzeLaborXUS` (`SatzID`)"
 Str(1, 127, 39) = "  KEY `Name` (`Nachname`,`Vorname`)"
 Str(1, 127, 40) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 127, 41) = "  CONSTRAINT `LaborXEingelLaborXUS_AccRel` FOREIGN KEY (`DatID`) REFERENCES `laborxeingel` (`DatID`) ON UPDATE CASCADE"
 Str(1, 127, 42) = "  CONSTRAINT `LaborXS??tzeLaborXUS_AccRel` FOREIGN KEY (`SatzID`) REFERENCES `laborxsaetze` (`SatzID`) ON UPDATE CASCADE"
 Str(1, 127, 43) = " ENGINE=InnoDB AUTO_INCREMENT=137759 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr127

Sub FüllStr128()
 Str(0, 128, 0) = "laborxwert"
 Str(0, 128, 1) = "`RefNr`"
 Str(0, 128, 2) = "`Abkü`"
 Str(0, 128, 3) = "`Langname`"
 Str(0, 128, 4) = "`Quelle`"
 Str(0, 128, 5) = "`QSpez`"
 Str(0, 128, 6) = "`AbnDat`"
 Str(0, 128, 7) = "`Wert`"
 Str(0, 128, 8) = "`Einheit`"
 Str(0, 128, 9) = "`Grenzwerti`"
 Str(0, 128, 10) = "`Kommentar`"
 Str(0, 128, 11) = "`Teststatus`"
 Str(0, 128, 12) = "`Erklärung`"
 Str(0, 128, 13) = "`AuftrHinw`"
 Str(0, 128, 14) = "`nbid`"
 Str(0, 128, 15) = "`LaborXUSLaborXWert`"
 Str(0, 128, 16) = "`LaborXWertAbkü`"
 Str(0, 128, 17) = "`nbid`"
 ArtZ(0, 128) = 14
 ArtZ(1, 128) = 3
 Str(1, 128, 0) = "CREATE TABLE `laborxwert` ("
 Str(1, 128, 1) = " `RefNr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborUS'"
 Str(1, 128, 2) = " `Abkü` varchar(16) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410 Test-Ident  (Turbomed)'"
 Str(1, 128, 3) = " `Langname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8411 Testbezeichnung (Turbomed)'"
 Str(1, 128, 4) = " `Quelle` varchar(43) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8430 Probenmaterial-Bezeichnung (Turbomed)'"
 Str(1, 128, 5) = " `QSpez` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8431 Probenmaterial-Spezifikation (Turbomed)'"
 Str(1, 128, 6) = " `AbnDat` datetime DEFAULT NULL COMMENT '8432 Abnahmedatum (Turbomed)'"
 Str(1, 128, 7) = " `Wert` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8420 Ergebniswert (Turbomed)'"
 Str(1, 128, 8) = " `Einheit` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421 Einheit (Turbomed)'"
 Str(1, 128, 9) = " `Grenzwerti` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8422 Grenzwertindikator (Turbomed)'"
 Str(1, 128, 10) = " `Kommentar` varchar(1385) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8480 Ergebnistext (Turbomed)'"
 Str(1, 128, 11) = " `Teststatus` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8418 Teststatus (Turbomed)'"
 Str(1, 128, 12) = " `Erklärung` varchar(1588) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8470 Testbezogene Hinweise (Turbomed)'"
 Str(1, 128, 13) = " `AuftrHinw` varchar(180) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8490 Auftragsbezogene Hinweise (Turbomed)'"
 Str(1, 128, 14) = " `nbid` int(10) unsigned NOT NULL COMMENT 'Bezug zu laborxplab.id'"
 Str(1, 128, 15) = "  KEY `LaborXUSLaborXWert` (`RefNr`)"
 Str(1, 128, 16) = "  KEY `LaborXWertAbkü` (`Abkü`,`Einheit`)"
 Str(1, 128, 17) = "  KEY `nbid` (`nbid`)"
 Str(1, 128, 18) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr128

Sub FüllStr129()
 Str(0, 129, 0) = "lanrpraxis"
 Str(0, 129, 1) = "`id`"
 Str(0, 129, 2) = "`LANR`"
 Str(0, 129, 3) = "`id`"
 Str(0, 129, 4) = "`LANR`"
 ArtZ(0, 129) = 2
 ArtZ(1, 129) = 2
 Str(1, 129, 0) = "CREATE TABLE `lanrpraxis` ("
 Str(1, 129, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'"
 Str(1, 129, 2) = " `LANR` int(9) unsigned NOT NULL COMMENT 'LANR'"
 Str(1, 129, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 129, 4) = "  KEY `LANR` (`LANR`)"
 Str(1, 129, 5) = " ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='LANR der in der Praxis tätigen Ärzte, für  Fälle, Formulare'"
End Sub ' FüllStr129

Sub FüllStr130()
 Str(0, 130, 0) = "lbanforderungen"
 Str(0, 130, 1) = "`FID`"
 Str(0, 130, 2) = "`Pat_ID`"
 Str(0, 130, 3) = "`ZeitPunkt`"
 Str(0, 130, 4) = "`AnfText`"
 Str(0, 130, 5) = "`absPos`"
 Str(0, 130, 6) = "`AktZeit`"
 Str(0, 130, 7) = "`StByte`"
 Str(0, 130, 8) = "`Auswahl`"
 Str(0, 130, 9) = "`FälleLbAnforderungen`"
 Str(0, 130, 10) = "`FID`"
 Str(0, 130, 11) = "`NamenLbAnforderungen`"
 Str(0, 130, 12) = "`F??lleLbAnforderungen_AccRel`"
 ArtZ(0, 130) = 7
 ArtZ(1, 130) = 4
 ArtZ(2, 130) = 1
 Str(1, 130, 0) = "CREATE TABLE `lbanforderungen` ("
 Str(1, 130, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 130, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 130, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 130, 4) = " `AnfText` longtext COLLATE latin1_german2_ci COMMENT '6280'"
 Str(1, 130, 5) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 130, 6) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 130, 7) = " `StByte` int(11) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 130, 8) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`AnfText`(255))"
 Str(1, 130, 9) = "  KEY `FälleLbAnforderungen` (`FID`)"
 Str(1, 130, 10) = "  KEY `FID` (`FID`)"
 Str(1, 130, 11) = "  KEY `NamenLbAnforderungen` (`Pat_ID`)"
 Str(1, 130, 12) = "  CONSTRAINT `F??lleLbAnforderungen_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 130, 13) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr130

Sub FüllStr131()
 Str(0, 131, 0) = "leistungen"
 Str(0, 131, 1) = "`id`"
 Str(0, 131, 2) = "`FID`"
 Str(0, 131, 3) = "`Pat_ID`"
 Str(0, 131, 4) = "`ZeitPunkt`"
 Str(0, 131, 5) = "`Leistung`"
 Str(0, 131, 6) = "`f5002`"
 Str(0, 131, 7) = "`f5005`"
 Str(0, 131, 8) = "`f5006`"
 Str(0, 131, 9) = "`f5009`"
 Str(0, 131, 10) = "`Med`"
 Str(0, 131, 11) = "`f5015`"
 Str(0, 131, 12) = "`f5016`"
 Str(0, 131, 13) = "`f5021`"
 Str(0, 131, 14) = "`f5026`"
 Str(0, 131, 15) = "`Faktor`"
 Str(0, 131, 16) = "`f5098`"
 Str(0, 131, 17) = "`LANR`"
 Str(0, 131, 18) = "`letzVorg`"
 Str(0, 131, 19) = "`Ausn`"
 Str(0, 131, 20) = "`Beme`"
 Str(0, 131, 21) = "`absPos`"
 Str(0, 131, 22) = "`AktZeit`"
 Str(0, 131, 23) = "`QS`"
 Str(0, 131, 24) = "`QT`"
 Str(0, 131, 25) = "`StByte`"
 Str(0, 131, 26) = "`LANRid`"
 Str(0, 131, 27) = "`Sachkbez`"
 Str(0, 131, 28) = "`Sachkct`"
 Str(0, 131, 29) = "`Zone`"
 Str(0, 131, 30) = "`id`"
 Str(0, 131, 31) = "`Auswahl`"
 Str(0, 131, 32) = "`FälleLeistungen`"
 Str(0, 131, 33) = "`Leistung`"
 Str(0, 131, 34) = "`NamenLeistungen`"
 Str(0, 131, 35) = "`Auswahl2`"
 Str(0, 131, 36) = "`F??lleLeistungen_AccRel`"
 ArtZ(0, 131) = 29
 ArtZ(1, 131) = 6
 ArtZ(2, 131) = 1
 Str(1, 131, 0) = "CREATE TABLE `leistungen` ("
 Str(1, 131, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'eindeutige ID, hinzugefügt 26.3.11'"
 Str(1, 131, 2) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 131, 3) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 131, 4) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '5000 + 6201'"
 Str(1, 131, 5) = " `Leistung` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5001 Leistungsziffer'"
 Str(1, 131, 6) = " `f5002` varchar(32) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5002 Art der Untersuchung'"
 Str(1, 131, 7) = " `f5005` varchar(2) COLLATE latin1_german2_ci NOT NULL DEFAULT '1' COMMENT '5005 Anzahl'"
 Str(1, 131, 8) = " `f5006` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5006 um Uhrzeit'"
 Str(1, 131, 9) = " `f5009` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5009 freier Begründungstext'"
 Str(1, 131, 10) = " `Med` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5010 Medikament'"
 Str(1, 131, 11) = " `f5015` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5015 Organ'"
 Str(1, 131, 12) = " `f5016` varchar(28) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5016 Name des Arztes (Briefempfänger)'"
 Str(1, 131, 13) = " `f5021` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5021 Datum letzte Krebsvorsorge'"
 Str(1, 131, 14) = " `f5026` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5026 Entlassungsdatum'"
 Str(1, 131, 15) = " `Faktor` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5062 Multiplikator für GOÄ-Rechnung'"
 Str(1, 131, 16) = " `f5098` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5098 0000000000'"
 Str(1, 131, 17) = " `LANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5099 LANR'"
 Str(1, 131, 18) = " `letzVorg` datetime DEFAULT NULL COMMENT '5101 letzter Vorgang'"
 Str(1, 131, 19) = " `Ausn` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3677 Ausnahme/Begründung für abweichendes Geschlecht'"
 Str(1, 131, 20) = " `Beme` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '         Bemerkung'"
 Str(1, 131, 21) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 131, 22) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 131, 23) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 131, 24) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 131, 25) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 131, 26) = " `LANRid` int(3) unsigned NOT NULL COMMENT 'Bezug auf lanrpraxis.id'"
 Str(1, 131, 27) = " `Sachkbez` varchar(34) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5011 Sachkostenbezeichnung'"
 Str(1, 131, 28) = " `Sachkct` int(8) unsigned DEFAULT NULL COMMENT '5012 Sach-/Materialkosten in ct'"
 Str(1, 131, 29) = " `Zone` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5018 Zone bei Besuchen'"
 Str(1, 131, 30) = "  PRIMARY KEY (`id`)"
 Str(1, 131, 31) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Leistung`)"
 Str(1, 131, 32) = "  KEY `FälleLeistungen` (`FID`)"
 Str(1, 131, 33) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 131, 34) = "  KEY `NamenLeistungen` (`Pat_ID`)"
 Str(1, 131, 35) = "  KEY `Auswahl2` (`Pat_ID`,`Leistung`,`ZeitPunkt`)"
 Str(1, 131, 36) = "  CONSTRAINT `F??lleLeistungen_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 131, 37) = " ENGINE=InnoDB AUTO_INCREMENT=26765365 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr131

Sub FüllStr132()
 Str(0, 132, 0) = "leistungen exportiert"
 Str(0, 132, 1) = "`ID`"
 Str(0, 132, 2) = "`Datum`"
 Str(0, 132, 3) = "`Pat_id`"
 Str(0, 132, 4) = "`SchGr`"
 Str(0, 132, 5) = "`Leistung`"
 Str(0, 132, 6) = "`übertragen`"
 Str(0, 132, 7) = "`ID`"
 Str(0, 132, 8) = "`PrimaryKey`"
 Str(0, 132, 9) = "`ID`"
 ArtZ(0, 132) = 6
 ArtZ(1, 132) = 3
 Str(1, 132, 0) = "CREATE TABLE `leistungen exportiert` ("
 Str(1, 132, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Reihenfolge'"
 Str(1, 132, 2) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum'"
 Str(1, 132, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!Pat_id'"
 Str(1, 132, 4) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 132, 5) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ziffer aus EBM / GOÄ'"
 Str(1, 132, 6) = " `übertragen` datetime DEFAULT NULL COMMENT '"""", übertragen'"
 Str(1, 132, 7) = "  PRIMARY KEY (`ID`)"
 Str(1, 132, 8) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 132, 9) = "  KEY `ID` (`Pat_id`)"
 Str(1, 132, 10) = " ENGINE=InnoDB AUTO_INCREMENT=3970 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr132

Sub FüllStr133()
 Str(0, 133, 0) = "leistungsexport"
 Str(0, 133, 1) = "`Prim`"
 Str(0, 133, 2) = "`PatID`"
 Str(0, 133, 3) = "`Datum`"
 Str(0, 133, 4) = "`UZeit`"
 Str(0, 133, 5) = "`Leistung`"
 Str(0, 133, 6) = "`SchGr`"
 Str(0, 133, 7) = "`Status`"
 Str(0, 133, 8) = "`Prim`"
 Str(0, 133, 9) = "`PatID`"
 ArtZ(0, 133) = 7
 ArtZ(1, 133) = 2
 Str(1, 133, 0) = "CREATE TABLE `leistungsexport` ("
 Str(1, 133, 1) = " `Prim` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Reihenfolge'"
 Str(1, 133, 2) = " `PatID` int(10) DEFAULT NULL COMMENT '-> Namen!Pat_id'"
 Str(1, 133, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum'"
 Str(1, 133, 4) = " `UZeit` datetime DEFAULT NULL COMMENT 'Leistungszeit, falls gewünscht'"
 Str(1, 133, 5) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ziffer aus EBM / GOÄ'"
 Str(1, 133, 6) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 133, 7) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '"""", übertragen'"
 Str(1, 133, 8) = "  PRIMARY KEY (`Prim`)"
 Str(1, 133, 9) = "  KEY `PatID` (`PatID`)"
 Str(1, 133, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr133

Sub FüllStr134()
 Str(0, 134, 0) = "letze faelle"
 Str(1, 134, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `letze faelle` AS SELECT `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik` FROM (`_lfaelle` `l` LEFT JOIN `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`))))"
End Sub ' FüllStr134

Sub FüllStr135()
 Str(0, 135, 0) = "letztefaelleverschieden"
 Str(1, 135, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `letztefaelleverschieden` AS SELECT `i`.`bhfb` AS `bhfb`,`i`.`pat_id` AS `pat_id` FROM `_f1` `i` group by `i`.`pat_id`"
End Sub ' FüllStr135

Sub FüllStr136()
 Str(0, 136, 0) = "lfaelle"
 Str(1, 136, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `lfaelle` AS SELECT `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr`,`f`.`ÜbWVBSNR` AS `übwvbsnr`,`f`.`ÜbWVKVNR` AS `übwvkvnr`,`f`.`ÜbWVLANR` AS `übwvlanr`,`f`.`Übwr` AS `übwr` FROM (`_lfaelle` `l` LEFT JOIN `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`))))"
End Sub ' FüllStr136

Sub FüllStr137()
 Str(0, 137, 0) = "lfaellev"
 Str(1, 137, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `lfaellev` AS SELECT `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr` FROM (`_lfaelle` `l` LEFT JOIN `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`)))) group by `f`.`Pat_ID`"
End Sub ' FüllStr137

Sub FüllStr138()
 Str(0, 138, 0) = "listenausgabeuew"
 Str(0, 138, 1) = "`name`"
 Str(0, 138, 2) = "`vorname`"
 Str(0, 138, 3) = "`titelt`"
 Str(0, 138, 4) = "`fachgruppe`"
 Str(0, 138, 5) = "`strasse`"
 Str(0, 138, 6) = "`plz`"
 Str(0, 138, 7) = "`ort`"
 Str(0, 138, 8) = "`telefon`"
 Str(0, 138, 9) = "`fax`"
 Str(0, 138, 10) = "`kvnr`"
 Str(0, 138, 11) = "`aktdat`"
 Str(0, 138, 12) = "`id`"
 Str(0, 138, 13) = "`überschrift`"
 Str(0, 138, 14) = "`dbnr`"
 Str(0, 138, 15) = "`bstelle`"
 Str(0, 138, 16) = "`anrede`"
 Str(0, 138, 17) = "`tel1`"
 Str(0, 138, 18) = "`tel2`"
 Str(0, 138, 19) = "`tel3`"
 Str(0, 138, 20) = "`tel4`"
 Str(0, 138, 21) = "`fax1`"
 Str(0, 138, 22) = "`fax2`"
 Str(0, 138, 23) = "`fax3`"
 Str(0, 138, 24) = "`email`"
 Str(0, 138, 25) = "`zulg`"
 Str(0, 138, 26) = "`arzttyp`"
 Str(0, 138, 27) = "`gemmit`"
 Str(0, 138, 28) = "`beme`"
 Str(0, 138, 29) = "`dmpt2`"
 Str(0, 138, 30) = "`dmpt1`"
 Str(0, 138, 31) = "`geschlecht`"
 Str(0, 138, 32) = "`titel`"
 Str(0, 138, 33) = "`id`"
 Str(0, 138, 34) = "`ident`"
 ArtZ(0, 138) = 32
 ArtZ(1, 138) = 2
 Str(1, 138, 0) = "CREATE TABLE `listenausgabeuew` ("
 Str(1, 138, 1) = " `name` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 2) = " `vorname` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 3) = " `titelt` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 4) = " `fachgruppe` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 5) = " `strasse` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 6) = " `plz` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 7) = " `ort` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 8) = " `telefon` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 138, 9) = " `fax` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 138, 10) = " `kvnr` varchar(9) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 138, 11) = " `aktdat` date DEFAULT NULL"
 Str(1, 138, 12) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 138, 13) = " `überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 14) = " `dbnr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 15) = " `bstelle` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 16) = " `anrede` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 17) = " `tel1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 18) = " `tel2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 19) = " `tel3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 20) = " `tel4` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 21) = " `fax1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 22) = " `fax2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 23) = " `fax3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 24) = " `email` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 25) = " `zulg` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 26) = " `arzttyp` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 27) = " `gemmit` longtext COLLATE latin1_german2_ci"
 Str(1, 138, 28) = " `beme` longtext COLLATE latin1_german2_ci"
 Str(1, 138, 29) = " `dmpt2` tinyint(1) DEFAULT NULL"
 Str(1, 138, 30) = " `dmpt1` tinyint(1) DEFAULT NULL"
 Str(1, 138, 31) = " `geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 32) = " `titel` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 138, 33) = "  PRIMARY KEY (`id`)"
 Str(1, 138, 34) = "  UNIQUE KEY `ident` (`name`,`vorname`,`plz`,`kvnr`)"
 Str(1, 138, 35) = " ENGINE=InnoDB AUTO_INCREMENT=1727 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr138

Sub FüllStr139()
 Str(0, 139, 0) = "listenausgabeuew_alt"
 Str(0, 139, 1) = "`name`"
 Str(0, 139, 2) = "`vorname`"
 Str(0, 139, 3) = "`titelt`"
 Str(0, 139, 4) = "`fachgruppe`"
 Str(0, 139, 5) = "`strasse`"
 Str(0, 139, 6) = "`plz`"
 Str(0, 139, 7) = "`ort`"
 Str(0, 139, 8) = "`telefon`"
 Str(0, 139, 9) = "`fax`"
 Str(0, 139, 10) = "`kvnr`"
 Str(0, 139, 11) = "`aktdat`"
 Str(0, 139, 12) = "`id`"
 Str(0, 139, 13) = "`überschrift`"
 Str(0, 139, 14) = "`dbnr`"
 Str(0, 139, 15) = "`bstelle`"
 Str(0, 139, 16) = "`anrede`"
 Str(0, 139, 17) = "`tel1`"
 Str(0, 139, 18) = "`tel2`"
 Str(0, 139, 19) = "`tel3`"
 Str(0, 139, 20) = "`tel4`"
 Str(0, 139, 21) = "`fax1`"
 Str(0, 139, 22) = "`fax2`"
 Str(0, 139, 23) = "`fax3`"
 Str(0, 139, 24) = "`email`"
 Str(0, 139, 25) = "`zulg`"
 Str(0, 139, 26) = "`arzttyp`"
 Str(0, 139, 27) = "`gemmit`"
 Str(0, 139, 28) = "`beme`"
 Str(0, 139, 29) = "`dmpt2`"
 Str(0, 139, 30) = "`dmpt1`"
 Str(0, 139, 31) = "`geschlecht`"
 Str(0, 139, 32) = "`titel`"
 Str(0, 139, 33) = "`id`"
 Str(0, 139, 34) = "`ident`"
 ArtZ(0, 139) = 32
 ArtZ(1, 139) = 2
 Str(1, 139, 0) = "CREATE TABLE `listenausgabeuew_alt` ("
 Str(1, 139, 1) = " `name` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 2) = " `vorname` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 3) = " `titelt` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 4) = " `fachgruppe` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 5) = " `strasse` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 6) = " `plz` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 7) = " `ort` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 8) = " `telefon` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 139, 9) = " `fax` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 139, 10) = " `kvnr` varchar(9) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 139, 11) = " `aktdat` date DEFAULT NULL"
 Str(1, 139, 12) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 139, 13) = " `überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 14) = " `dbnr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 15) = " `bstelle` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 16) = " `anrede` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 17) = " `tel1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 18) = " `tel2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 19) = " `tel3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 20) = " `tel4` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 21) = " `fax1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 22) = " `fax2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 23) = " `fax3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 24) = " `email` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 25) = " `zulg` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 26) = " `arzttyp` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 27) = " `gemmit` longtext COLLATE latin1_german2_ci"
 Str(1, 139, 28) = " `beme` longtext COLLATE latin1_german2_ci"
 Str(1, 139, 29) = " `dmpt2` tinyint(1) DEFAULT NULL"
 Str(1, 139, 30) = " `dmpt1` tinyint(1) DEFAULT NULL"
 Str(1, 139, 31) = " `geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 32) = " `titel` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 139, 33) = "  PRIMARY KEY (`id`)"
 Str(1, 139, 34) = "  UNIQUE KEY `ident` (`name`,`vorname`,`plz`,`kvnr`)"
 Str(1, 139, 35) = " ENGINE=InnoDB AUTO_INCREMENT=1705 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr139

Sub FüllStr140()
 Str(0, 140, 0) = "lmp"
 Str(1, 140, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `lmp` AS SELECT `mp`.`FID` AS `FID`,`mp`.`Pat_ID` AS `Pat_ID`,`mp`.`MPNr` AS `MPNr`,`mp`.`ZeitPunkt` AS `ZeitPunkt`,`mp`.`Datum` AS `Datum`,`mp`.`Medikament` AS `Medikament`,`mp`.`MedAnfang` AS `MedAnfang`,`mp`.`FeldNr` AS `FeldNr`,`mp`.`mo` AS `mo`,`mp`.`mi` AS `mi`,`mp`.`nm` AS `nm`,`mp`.`ab` AS `ab`,`mp`.`zn` AS `zn`,`mp`.`bBed` AS `bBed`,`mp`.`Bemerkung` AS `Bemerkung`,`mp`.`AbsPos` AS `AbsPos`,`mp`.`AktZeit` AS `AktZeit`,`mp`.`StByte` AS `StByte` FROM (`_fuerlmp` `i` join `medplan` `mp` on(((`i`.`pat_id` = `mp`.`Pat_ID`) and (`i`.`mpnr` = `mp`.`MPNr`))))"
End Sub ' FüllStr140

Sub FüllStr141()
 Str(0, 141, 0) = "maxGluc"
 Str(1, 141, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `maxGluc` AS SELECT `_maxGluc`.`pat_id` AS `pat_id`,if(((`_maxGluc`.`w1` > `_maxGluc`.`w2`) or isnull(`_maxGluc`.`w2`)),`_maxGluc`.`w1`,`_maxGluc`.`w2`) AS `Wert` FROM `_maxGluc`"
End Sub ' FüllStr141

Sub FüllStr142()
 Str(0, 142, 0) = "maxHbA1c"
 Str(1, 142, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `maxHbA1c` AS SELECT `_maxHbA1c`.`pat_id` AS `pat_id`,if(((`_maxHbA1c`.`w1` > `_maxHbA1c`.`w2`) or isnull(`_maxHbA1c`.`w2`)),`_maxHbA1c`.`w1`,`_maxHbA1c`.`w2`) AS `Wert` FROM `_maxHbA1c`"
End Sub ' FüllStr142

Sub FüllStr143()
 Str(0, 143, 0) = "medarten"
 Str(0, 143, 1) = "`Medikament`"
 Str(0, 143, 2) = "`Langname`"
 Str(0, 143, 3) = "`Pat_ID`"
 Str(0, 143, 4) = "`Anzahl`"
 Str(0, 143, 5) = "`Glib`"
 Str(0, 143, 6) = "`Metf`"
 Str(0, 143, 7) = "`GlucI`"
 Str(0, 143, 8) = "`SHGlin`"
 Str(0, 143, 9) = "`Glit`"
 Str(0, 143, 10) = "`DPP4`"
 Str(0, 143, 11) = "`GLP1`"
 Str(0, 143, 12) = "`SGLT2`"
 Str(0, 143, 13) = "`SonstAD`"
 Str(0, 143, 14) = "`Ins`"
 Str(0, 143, 15) = "`Anal`"
 Str(0, 143, 16) = "`InsArt`"
 Str(0, 143, 17) = "`HMG`"
 Str(0, 143, 18) = "`Hypt`"
 Str(0, 143, 19) = "`Thro`"
 Str(0, 143, 20) = "`Antib`"
 Str(0, 143, 21) = "`and`"
 Str(0, 143, 22) = "`hinzugefügt`"
 Str(0, 143, 23) = "`Tstr`"
 Str(0, 143, 24) = "`Puzu`"
 Str(0, 143, 25) = "`VMat`"
 Str(0, 143, 26) = "`PenN`"
 Str(0, 143, 27) = "`Neurp`"
 Str(0, 143, 28) = "`AutNP`"
 Str(0, 143, 29) = "`Fetts`"
 Str(0, 143, 30) = "`Hsre`"
 Str(0, 143, 31) = "`AntiMyk`"
 Str(0, 143, 32) = "`Glauk`"
 Str(0, 143, 33) = "`COLD`"
 Str(0, 143, 34) = "`Pros`"
 Str(0, 143, 35) = "`Urä`"
 Str(0, 143, 36) = "`HyThy`"
 Str(0, 143, 37) = "`Ostp`"
 Str(0, 143, 38) = "`KHK`"
 Str(0, 143, 39) = "`HerzI`"
 Str(0, 143, 40) = "`Stru`"
 Str(0, 143, 41) = "`AVK`"
 Str(0, 143, 42) = "`PanI`"
 Str(0, 143, 43) = "`Vari`"
 Str(0, 143, 44) = "`Östr`"
 Str(0, 143, 45) = "`AntiDep`"
 Str(0, 143, 46) = "`AntiDem`"
 Str(0, 143, 47) = "`AntiEp`"
 Str(0, 143, 48) = "`Park`"
 Str(0, 143, 49) = "`AntiPern`"
 Str(0, 143, 50) = "`Appet`"
 Str(0, 143, 51) = "`Anäm`"
 Str(0, 143, 52) = "`Antiherp`"
 Str(0, 143, 53) = "`NSAR`"
 Str(0, 143, 54) = "`Antikoag`"
 Str(0, 143, 55) = "`Betabl`"
 Str(0, 143, 56) = "`ACEH`"
 Str(0, 143, 57) = "`AT1`"
 Str(0, 143, 58) = "`CalcA`"
 Str(0, 143, 59) = "`Diur`"
 Str(0, 143, 60) = "`falsch`"
 Str(0, 143, 61) = "`ID`"
 Str(0, 143, 62) = "`ID`"
 Str(0, 143, 63) = "`Medikament`"
 Str(0, 143, 64) = "`pat_id`"
 ArtZ(0, 143) = 61
 ArtZ(1, 143) = 3
 Str(1, 143, 0) = "CREATE TABLE `medarten` ("
 Str(1, 143, 1) = " `Medikament` varchar(38) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT ''"
 Str(1, 143, 2) = " `Langname` varchar(150) COLLATE latin1_german2_ci DEFAULT '' COMMENT 'Beispiel-Langname'"
 Str(1, 143, 3) = " `Pat_ID` int(10) DEFAULT '0' COMMENT 'Beispiel-PatID'"
 Str(1, 143, 4) = " `Anzahl` int(10) DEFAULT '0' COMMENT 'Anzahl der Vorkommen'"
 Str(1, 143, 5) = " `Glib` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Glibenclamid'"
 Str(1, 143, 6) = " `Metf` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Metformin'"
 Str(1, 143, 7) = " `GlucI` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Glucosidase-Inhibitoren'"
 Str(1, 143, 8) = " `SHGlin` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'andere Sulfonylharnstoffe oder Glinide'"
 Str(1, 143, 9) = " `Glit` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Glitanzone'"
 Str(1, 143, 10) = " `DPP4` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'DPP4-Hemmer'"
 Str(1, 143, 11) = " `GLP1` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'GLP1-Analogon'"
 Str(1, 143, 12) = " `SGLT2` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'SGLT2-Hemmer'"
 Str(1, 143, 13) = " `SonstAD` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Sonstige antidiabetische Medikation'"
 Str(1, 143, 14) = " `Ins` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Insulin'"
 Str(1, 143, 15) = " `Anal` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Insulin-Analoga'"
 Str(1, 143, 16) = " `InsArt` varchar(1) COLLATE latin1_german2_ci DEFAULT '0' COMMENT '1= schnell, 2 = langsam, 3 = Misch'"
 Str(1, 143, 17) = " `HMG` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'HMG-CoA-Reduktase-Inhibitoren'"
 Str(1, 143, 18) = " `Hypt` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Hypertonie-Mittel'"
 Str(1, 143, 19) = " `Thro` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Thrombozyten-Hemmer'"
 Str(1, 143, 20) = " `Antib` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Antibiotika'"
 Str(1, 143, 21) = " `and` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'andere'"
 Str(1, 143, 22) = " `hinzugefügt` datetime DEFAULT '0000-00-00 00:00:00'"
 Str(1, 143, 23) = " `Tstr` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Teststreifen'"
 Str(1, 143, 24) = " `Puzu` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Pumpenzubehör'"
 Str(1, 143, 25) = " `VMat` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Verbandsmaterial'"
 Str(1, 143, 26) = " `PenN` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Pennadeln'"
 Str(1, 143, 27) = " `Neurp` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Neuropathie-Behandlungsmittel'"
 Str(1, 143, 28) = " `AutNP` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Autonome Neuropathie'"
 Str(1, 143, 29) = " `Fetts` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Fibrate, Ezetrol, Niaspan u.a.'"
 Str(1, 143, 30) = " `Hsre` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Hyperuriämie-Mittel'"
 Str(1, 143, 31) = " `AntiMyk` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Antimykotika'"
 Str(1, 143, 32) = " `Glauk` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Glaukom'"
 Str(1, 143, 33) = " `COLD` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'COLD und Asthma'"
 Str(1, 143, 34) = " `Pros` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Prostatahypertrophie'"
 Str(1, 143, 35) = " `Urä` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Urämie-spezifische'"
 Str(1, 143, 36) = " `HyThy` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'thyreostatische Mittel'"
 Str(1, 143, 37) = " `Ostp` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Osteoporosemittel'"
 Str(1, 143, 38) = " `KHK` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'KHK-spezifisch'"
 Str(1, 143, 39) = " `HerzI` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Herzinsuffizienz-spezifische'"
 Str(1, 143, 40) = " `Stru` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Struma- und Hypothyreosemittel'"
 Str(1, 143, 41) = " `AVK` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'AVK-Mittel'"
 Str(1, 143, 42) = " `PanI` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Pankreasinsuffizienz'"
 Str(1, 143, 43) = " `Vari` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Varikosemittel'"
 Str(1, 143, 44) = " `Östr` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Östrogene, Gestagene usw.'"
 Str(1, 143, 45) = " `AntiDep` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Antidepressiva'"
 Str(1, 143, 46) = " `AntiDem` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Antidementika'"
 Str(1, 143, 47) = " `AntiEp` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Antiepileptika'"
 Str(1, 143, 48) = " `Park` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Parkinson-Medikament'"
 Str(1, 143, 49) = " `AntiPern` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Antiperniziosa'"
 Str(1, 143, 50) = " `Appet` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Appetitzügler'"
 Str(1, 143, 51) = " `Anäm` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Anämiebehandlungsmittel'"
 Str(1, 143, 52) = " `Antiherp` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Anti-Herpes-Mittel'"
 Str(1, 143, 53) = " `NSAR` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'NSAR'"
 Str(1, 143, 54) = " `Antikoag` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Antikoagulatien'"
 Str(1, 143, 55) = " `Betabl` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Betablocker'"
 Str(1, 143, 56) = " `ACEH` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'ACE-Hemmer'"
 Str(1, 143, 57) = " `AT1` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'AT-1-Blocker'"
 Str(1, 143, 58) = " `CalcA` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Calcium-Antagonist'"
 Str(1, 143, 59) = " `Diur` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'Diuretikum'"
 Str(1, 143, 60) = " `falsch` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'falsch geschrieben oder nicht erkennbar'"
 Str(1, 143, 61) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'"
 Str(1, 143, 62) = "  PRIMARY KEY (`ID`)"
 Str(1, 143, 63) = "  UNIQUE KEY `Medikament` (`Medikament`)"
 Str(1, 143, 64) = "  KEY `pat_id` (`Pat_ID`)"
 Str(1, 143, 65) = " ENGINE=InnoDB AUTO_INCREMENT=7231 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr143

Sub FüllStr144()
 Str(0, 144, 0) = "medplan"
 Str(0, 144, 1) = "`FID`"
 Str(0, 144, 2) = "`Pat_ID`"
 Str(0, 144, 3) = "`MPNr`"
 Str(0, 144, 4) = "`ZeitPunkt`"
 Str(0, 144, 5) = "`Datum`"
 Str(0, 144, 6) = "`Medikament`"
 Str(0, 144, 7) = "`MedAnfang`"
 Str(0, 144, 8) = "`FeldNr`"
 Str(0, 144, 9) = "`mo`"
 Str(0, 144, 10) = "`mi`"
 Str(0, 144, 11) = "`nm`"
 Str(0, 144, 12) = "`ab`"
 Str(0, 144, 13) = "`zn`"
 Str(0, 144, 14) = "`bBed`"
 Str(0, 144, 15) = "`Bemerkung`"
 Str(0, 144, 16) = "`AbsPos`"
 Str(0, 144, 17) = "`AktZeit`"
 Str(0, 144, 18) = "`StByte`"
 Str(0, 144, 19) = "`FälleMedPlan`"
 Str(0, 144, 20) = "`MedPlanMedikament`"
 Str(0, 144, 21) = "`NamenMedPlan`"
 Str(0, 144, 22) = "`MedArtenMedPlan_AccRel`"
 Str(0, 144, 23) = "`MPNrPat_ID`"
 Str(0, 144, 24) = "`Auswahl`"
 Str(0, 144, 25) = "`F??lleMedPlan_AccRel`"
 Str(0, 144, 26) = "`MedArtenMedPlan_AccRel`"
 ArtZ(0, 144) = 18
 ArtZ(1, 144) = 6
 ArtZ(2, 144) = 2
 Str(1, 144, 0) = "CREATE TABLE `medplan` ("
 Str(1, 144, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 144, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 144, 3) = " `MPNr` int(10) DEFAULT NULL COMMENT 'Ordnungsziffer für Medikamentenplan'"
 Str(1, 144, 4) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt, der Speicherung im Turbomed'"
 Str(1, 144, 5) = " `Datum` datetime DEFAULT NULL COMMENT 'Zeitpunkt aus dem Kopf des Medikamentenplans'"
 Str(1, 144, 6) = " `Medikament` varchar(57) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 144, 7) = " `MedAnfang` varchar(36) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL"
 Str(1, 144, 8) = " `FeldNr` smallint(6) DEFAULT NULL"
 Str(1, 144, 9) = " `mo` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 144, 10) = " `mi` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 144, 11) = " `nm` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 144, 12) = " `ab` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 144, 13) = " `zn` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 144, 14) = " `bBed` bit(1) DEFAULT NULL"
 Str(1, 144, 15) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 144, 16) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 144, 17) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 144, 18) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 144, 19) = "  KEY `FälleMedPlan` (`FID`)"
 Str(1, 144, 20) = "  KEY `MedPlanMedikament` (`Medikament`)"
 Str(1, 144, 21) = "  KEY `NamenMedPlan` (`Pat_ID`)"
 Str(1, 144, 22) = "  KEY `MedArtenMedPlan_AccRel` (`MedAnfang`)"
 Str(1, 144, 23) = "  KEY `MPNrPat_ID` (`Pat_ID`,`MPNr`)"
 Str(1, 144, 24) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`MPNr`) USING BTREE"
 Str(1, 144, 25) = "  CONSTRAINT `F??lleMedPlan_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 144, 26) = "  CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`)"
 Str(1, 144, 27) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr144

Sub FüllStr145()
 Str(0, 145, 0) = "namen"
 Str(0, 145, 1) = "`Pat_ID`"
 Str(0, 145, 2) = "`lfdnr`"
 Str(0, 145, 3) = "`NVorsatz`"
 Str(0, 145, 4) = "`Nachname`"
 Str(0, 145, 5) = "`Vorname`"
 Str(0, 145, 6) = "`GebDat`"
 Str(0, 145, 7) = "`Straße`"
 Str(0, 145, 8) = "`KVKStatus`"
 Str(0, 145, 9) = "`Geschlecht`"
 Str(0, 145, 10) = "`Plz`"
 Str(0, 145, 11) = "`Ort`"
 Str(0, 145, 12) = "`Postfach`"
 Str(0, 145, 13) = "`Weggeldzone`"
 Str(0, 145, 14) = "`WeggzZahl`"
 Str(0, 145, 15) = "`AufnDat`"
 Str(0, 145, 16) = "`LANR`"
 Str(0, 145, 17) = "`BStNr`"
 Str(0, 145, 18) = "`Titel`"
 Str(0, 145, 19) = "`Versichertennummer`"
 Str(0, 145, 20) = "`PrivatTel`"
 Str(0, 145, 21) = "`KVNr`"
 Str(0, 145, 22) = "`KVNr2`"
 Str(0, 145, 23) = "`KVNr3`"
 Str(0, 145, 24) = "`KVNr4`"
 Str(0, 145, 25) = "`PrivatTel_2`"
 Str(0, 145, 26) = "`PrivatFax`"
 Str(0, 145, 27) = "`DienstTel`"
 Str(0, 145, 28) = "`PrivatMobil`"
 Str(0, 145, 29) = "`Email`"
 Str(0, 145, 30) = "`Arbeitgeber`"
 Str(0, 145, 31) = "`AnAllgda`"
 Str(0, 145, 32) = "`An1da`"
 Str(0, 145, 33) = "`An2da`"
 Str(0, 145, 34) = "`Checkda`"
 Str(0, 145, 35) = "`DMTypaD`"
 Str(0, 145, 36) = "`AktZeit`"
 Str(0, 145, 37) = "`absPos`"
 Str(0, 145, 38) = "`StByte`"
 Str(0, 145, 39) = "`Cave`"
 Str(0, 145, 40) = "`Notiz`"
 Str(0, 145, 41) = "`f3800`"
 Str(0, 145, 42) = "`dmpklass`"
 Str(0, 145, 43) = "`dmpbeg`"
 Str(0, 145, 44) = "`dmpkhkklass`"
 Str(0, 145, 45) = "`dmpkhkbeg`"
 Str(0, 145, 46) = "`dmpcopdklass`"
 Str(0, 145, 47) = "`dmpcopdbeg`"
 Str(0, 145, 48) = "`getHA0`"
 Str(0, 145, 49) = "`fnHA0`"
 Str(0, 145, 50) = "`getHA1`"
 Str(0, 145, 51) = "`fnHA1`"
 Str(0, 145, 52) = "`getHA2`"
 Str(0, 145, 53) = "`fnHA2`"
 Str(0, 145, 54) = "`zubenach`"
 Str(0, 145, 55) = "`Verwandt`"
 Str(0, 145, 56) = "`Sprache`"
 Str(0, 145, 57) = "`lAktTM`"
 Str(0, 145, 58) = "`PAT_ID`"
 Str(0, 145, 59) = "`Auswahl`"
 Str(0, 145, 60) = "`HausärzteNamen_AccRel`"
 Str(0, 145, 61) = "`weggszahl`"
 Str(0, 145, 62) = "`weggeldzone`"
 Str(0, 145, 63) = "`Haus??rzteNamen_AccRel`"
 ArtZ(0, 145) = 57
 ArtZ(1, 145) = 5
 ArtZ(2, 145) = 1
 Str(1, 145, 0) = "CREATE TABLE `namen` ("
 Str(1, 145, 1) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 145, 2) = " `lfdnr` int(10) DEFAULT NULL COMMENT 'laufende Patientennummer'"
 Str(1, 145, 3) = " `NVorsatz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3100'"
 Str(1, 145, 4) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 145, 5) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 145, 6) = " `GebDat` datetime DEFAULT NULL COMMENT '3103'"
 Str(1, 145, 7) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3107'"
 Str(1, 145, 8) = " `KVKStatus` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3108'"
 Str(1, 145, 9) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3110'"
 Str(1, 145, 10) = " `Plz` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3112'"
 Str(1, 145, 11) = " `Ort` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3113'"
 Str(1, 145, 12) = " `Postfach` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3216'"
 Str(1, 145, 13) = " `Weggeldzone` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3631 (1) Weggeldzone mit Z'"
 Str(1, 145, 14) = " `WeggzZahl` decimal(1,0) DEFAULT NULL COMMENT '3631 (2) Weggeldzone, Zahl in Feld 2'"
 Str(1, 145, 15) = " `AufnDat` datetime DEFAULT NULL COMMENT '3610'"
 Str(1, 145, 16) = " `LANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP'"
 Str(1, 145, 17) = " `BStNr` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3536 Betriebsstättennummer'"
 Str(1, 145, 18) = " `Titel` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3104'"
 Str(1, 145, 19) = " `Versichertennummer` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3105'"
 Str(1, 145, 20) = " `PrivatTel` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 145, 21) = " `KVNr` varchar(9) COLLATE latin1_german2_ci NOT NULL DEFAULT '""""' COMMENT '3630 Hausarzt'"
 Str(1, 145, 22) = " `KVNr2` varchar(9) COLLATE latin1_german2_ci NOT NULL DEFAULT '""""' COMMENT '3630 Hausarzt (2.Eintrag)'"
 Str(1, 145, 23) = " `KVNr3` varchar(9) COLLATE latin1_german2_ci NOT NULL DEFAULT '""""' COMMENT '3630 Hausarzt (3.Eintrag)'"
 Str(1, 145, 24) = " `KVNr4` varchar(9) COLLATE latin1_german2_ci NOT NULL COMMENT '3630 Hausarzt (4.Eintrag)'"
 Str(1, 145, 25) = " `PrivatTel_2` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 145, 26) = " `PrivatFax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 145, 27) = " `DienstTel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 145, 28) = " `PrivatMobil` varchar(52) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 145, 29) = " `Email` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Email'"
 Str(1, 145, 30) = " `Arbeitgeber` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3625'"
 Str(1, 145, 31) = " `AnAllgda` bit(1) DEFAULT NULL COMMENT 'Anamnese allgemein da'"
 Str(1, 145, 32) = " `An1da` bit(1) DEFAULT NULL COMMENT 'Anamnese S.1 da'"
 Str(1, 145, 33) = " `An2da` bit(1) DEFAULT NULL COMMENT 'Anamnese S.2 da'"
 Str(1, 145, 34) = " `Checkda` bit(1) DEFAULT NULL COMMENT 'Checkliste da'"
 Str(1, 145, 35) = " `DMTypaD` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'aus Diagnosen'"
 Str(1, 145, 36) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 145, 37) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 145, 38) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 145, 39) = " `Cave` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3654'"
 Str(1, 145, 40) = " `Notiz` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>, DMP ausgeschrieben <datum>'"
 Str(1, 145, 41) = " `f3800` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3800'"
 Str(1, 145, 42) = " `dmpklass` int(1) unsigned DEFAULT '0' COMMENT '1 = DMP nein, 2 = DMP HA, 3 = DMP hier, 4 = DMP ausgeschrieben'"
 Str(1, 145, 43) = " `dmpbeg` date DEFAULT '0000-00-00' COMMENT 'Datum der aktuellen DMP-Klassifikation'"
 Str(1, 145, 44) = " `dmpkhkklass` int(1) unsigned DEFAULT '0' COMMENT '1 = DMP nein, 2 = DMP HA, 3 = DMP hier'"
 Str(1, 145, 45) = " `dmpkhkbeg` date DEFAULT '0000-00-00' COMMENT 'Datum der aktuellen DMP-Klassifikation'"
 Str(1, 145, 46) = " `dmpcopdklass` int(1) unsigned DEFAULT '0' COMMENT '1 = DMP nein, 2 = DMP HA, 3 = DMP hier'"
 Str(1, 145, 47) = " `dmpcopdbeg` date DEFAULT '0000-00-00' COMMENT 'Datum der aktuellen DMP-Klassifikation'"
 Str(1, 145, 48) = " `getHA0` int(10) unsigned DEFAULT NULL COMMENT 'KVNr aus getHausarzt -> Üw(12,0)'"
 Str(1, 145, 49) = " `fnHA0` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Funktion aus getHausarzt -> Üw(10,0)'"
 Str(1, 145, 50) = " `getHA1` int(10) unsigned DEFAULT NULL COMMENT 'KVNr aus getHausarzt -> Üw(12,1)'"
 Str(1, 145, 51) = " `fnHA1` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Funktion aus getHausarzt -> Üw(10,1)'"
 Str(1, 145, 52) = " `getHA2` int(10) unsigned DEFAULT NULL COMMENT 'KVNr aus getHausarzt -> Üw(12,2)'"
 Str(1, 145, 53) = " `fnHA2` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Funktion aus getHausarzt -> Üw(10,2)'"
 Str(1, 145, 54) = " `zubenach` varchar(105) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3633'"
 Str(1, 145, 55) = " `Verwandt` varchar(488) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3632'"
 Str(1, 145, 56) = " `Sprache` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3628'"
 Str(1, 145, 57) = " `lAktTM` datetime NOT NULL COMMENT 'letzte Aktualisierung in Turbomed'"
 Str(1, 145, 58) = "  UNIQUE KEY `PAT_ID` (`Pat_ID`)"
 Str(1, 145, 59) = "  KEY `Auswahl` (`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 145, 60) = "  KEY `HausärzteNamen_AccRel` (`KVNr`)"
 Str(1, 145, 61) = "  KEY `weggszahl` (`WeggzZahl`)"
 Str(1, 145, 62) = "  KEY `weggeldzone` (`Weggeldzone`)"
 Str(1, 145, 63) = "  CONSTRAINT `Haus??rzteNamen_AccRel` FOREIGN KEY (`KVNr`) REFERENCES `hausaerztealt` (`KVNr`)"
 Str(1, 145, 64) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr145

Sub FüllStr146()
 Str(0, 146, 0) = "ni_abr"
 Str(0, 146, 1) = "`id`"
 Str(0, 146, 2) = "`pat_id`"
 Str(0, 146, 3) = "`DmICD`"
 Str(0, 146, 4) = "`maxHbA1c`"
 Str(0, 146, 5) = "`maxGluc`"
 Str(0, 146, 6) = "`eGFR`"
 Str(0, 146, 7) = "`npICD`"
 Str(0, 146, 8) = "`niICD`"
 Str(0, 146, 9) = "`pZ`"
 Str(0, 146, 10) = "`nZ`"
 Str(0, 146, 11) = "`gesZ`"
 Str(0, 146, 12) = "`minDat`"
 Str(0, 146, 13) = "`maxAlb`"
 Str(0, 146, 14) = "`id`"
 ArtZ(0, 146) = 13
 ArtZ(1, 146) = 1
 Str(1, 146, 0) = "CREATE TABLE `ni_abr` ("
 Str(1, 146, 1) = " `id` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 146, 2) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 146, 3) = " `DmICD` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 146, 4) = " `maxHbA1c` float DEFAULT NULL"
 Str(1, 146, 5) = " `maxGluc` float DEFAULT NULL"
 Str(1, 146, 6) = " `eGFR` float DEFAULT NULL"
 Str(1, 146, 7) = " `npICD` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 146, 8) = " `niICD` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 146, 9) = " `pZ` int(3) DEFAULT NULL"
 Str(1, 146, 10) = " `nZ` int(3) DEFAULT NULL"
 Str(1, 146, 11) = " `gesZ` int(3) DEFAULT NULL"
 Str(1, 146, 12) = " `minDat` date DEFAULT NULL"
 Str(1, 146, 13) = " `maxAlb` decimal(8,2) DEFAULT NULL"
 Str(1, 146, 14) = "  PRIMARY KEY (`id`)"
 Str(1, 146, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr146

Sub FüllStr147()
 Str(0, 147, 0) = "outa"
 Str(0, 147, 1) = "`titel`"
 Str(0, 147, 2) = "`tsid`"
 Str(0, 147, 3) = "`submt`"
 Str(0, 147, 4) = "`submid`"
 Str(0, 147, 5) = "`oscht`"
 Str(0, 147, 6) = "`subject`"
 Str(0, 147, 7) = "`docname`"
 Str(0, 147, 8) = "`id`"
 Str(0, 147, 9) = "`fsize`"
 Str(0, 147, 10) = "`pages`"
 Str(0, 147, 11) = "`devname`"
 Str(0, 147, 12) = "`retries`"
 Str(0, 147, 13) = "`prio`"
 Str(0, 147, 14) = "`rcfax`"
 Str(0, 147, 15) = "`rcname`"
 Str(0, 147, 16) = "`csid`"
 Str(0, 147, 17) = "`sender`"
 Str(0, 147, 18) = "`transs`"
 Str(0, 147, 19) = "`transe`"
 Str(0, 147, 20) = "`Pat_id`"
 Str(0, 147, 21) = "`Id`"
 Str(0, 147, 22) = "`submt`"
 Str(0, 147, 23) = "`oscht`"
 Str(0, 147, 24) = "`docname`"
 Str(0, 147, 25) = "`Pid`"
 ArtZ(0, 147) = 20
 ArtZ(1, 147) = 5
 Str(1, 147, 0) = "CREATE TABLE `outa` ("
 Str(1, 147, 1) = " `titel` varchar(596) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 2) = " `tsid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 3) = " `submt` datetime DEFAULT NULL"
 Str(1, 147, 4) = " `submid` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 5) = " `oscht` datetime DEFAULT NULL"
 Str(1, 147, 6) = " `subject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 7) = " `docname` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 8) = " `id` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 9) = " `fsize` int(10) DEFAULT NULL"
 Str(1, 147, 10) = " `pages` int(10) DEFAULT NULL"
 Str(1, 147, 11) = " `devname` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 12) = " `retries` int(10) DEFAULT NULL"
 Str(1, 147, 13) = " `prio` int(10) DEFAULT NULL"
 Str(1, 147, 14) = " `rcfax` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 15) = " `rcname` varchar(596) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 16) = " `csid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 17) = " `sender` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 147, 18) = " `transs` datetime DEFAULT NULL"
 Str(1, 147, 19) = " `transe` datetime DEFAULT NULL"
 Str(1, 147, 20) = " `Pid` int(10) unsigned DEFAULT NULL"
 Str(1, 147, 21) = "  KEY `Id` (`id`)"
 Str(1, 147, 22) = "  KEY `submt` (`submt`)"
 Str(1, 147, 23) = "  KEY `oscht` (`oscht`)"
 Str(1, 147, 24) = "  KEY `docname` (`docname`)"
 Str(1, 147, 25) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 147, 26) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr147

Sub FüllStr148()
 Str(0, 148, 0) = "outq"
 Str(0, 148, 1) = "`titel`"
 Str(0, 148, 2) = "`tsid`"
 Str(0, 148, 3) = "`submt`"
 Str(0, 148, 4) = "`submid`"
 Str(0, 148, 5) = "`scht`"
 Str(0, 148, 6) = "`oscht`"
 Str(0, 148, 7) = "`subject`"
 Str(0, 148, 8) = "`docname`"
 Str(0, 148, 9) = "`id`"
 Str(0, 148, 10) = "`fsize`"
 Str(0, 148, 11) = "`pages`"
 Str(0, 148, 12) = "`curp`"
 Str(0, 148, 13) = "`devid`"
 Str(0, 148, 14) = "`status`"
 Str(0, 148, 15) = "`exts`"
 Str(0, 148, 16) = "`extsc`"
 Str(0, 148, 17) = "`retries`"
 Str(0, 148, 18) = "`prio`"
 Str(0, 148, 19) = "`rct`"
 Str(0, 148, 20) = "`rcfax`"
 Str(0, 148, 21) = "`rcname`"
 Str(0, 148, 22) = "`csid`"
 Str(0, 148, 23) = "`avop`"
 Str(0, 148, 24) = "`gbr`"
 Str(0, 148, 25) = "`sender`"
 Str(0, 148, 26) = "`obalt`"
 Str(0, 148, 27) = "`Pat_id`"
 Str(0, 148, 28) = "`Id`"
 Str(0, 148, 29) = "`submt`"
 Str(0, 148, 30) = "`scht`"
 Str(0, 148, 31) = "`docname`"
 Str(0, 148, 32) = "`pat_id`"
 ArtZ(0, 148) = 27
 ArtZ(1, 148) = 5
 Str(1, 148, 0) = "CREATE TABLE `outq` ("
 Str(1, 148, 1) = " `titel` varchar(568) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 2) = " `tsid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 3) = " `submt` datetime DEFAULT NULL"
 Str(1, 148, 4) = " `submid` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 5) = " `scht` datetime DEFAULT NULL"
 Str(1, 148, 6) = " `oscht` datetime DEFAULT NULL"
 Str(1, 148, 7) = " `subject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 8) = " `docname` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 9) = " `id` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 10) = " `fsize` int(10) DEFAULT NULL"
 Str(1, 148, 11) = " `pages` int(10) DEFAULT NULL"
 Str(1, 148, 12) = " `curp` int(10) DEFAULT NULL"
 Str(1, 148, 13) = " `devid` int(10) DEFAULT NULL"
 Str(1, 148, 14) = " `status` int(10) DEFAULT NULL"
 Str(1, 148, 15) = " `exts` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 16) = " `extsc` int(10) DEFAULT NULL"
 Str(1, 148, 17) = " `retries` int(10) DEFAULT NULL"
 Str(1, 148, 18) = " `prio` int(10) DEFAULT NULL"
 Str(1, 148, 19) = " `rct` int(10) DEFAULT NULL"
 Str(1, 148, 20) = " `rcfax` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 21) = " `rcname` varchar(568) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 22) = " `csid` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 23) = " `avop` int(10) DEFAULT NULL"
 Str(1, 148, 24) = " `gbr` tinyint(1) DEFAULT NULL"
 Str(1, 148, 25) = " `sender` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 148, 26) = " `obalt` tinyint(1) DEFAULT NULL"
 Str(1, 148, 27) = " `Pat_id` int(10) unsigned DEFAULT NULL"
 Str(1, 148, 28) = "  KEY `Id` (`id`)"
 Str(1, 148, 29) = "  KEY `submt` (`submt`)"
 Str(1, 148, 30) = "  KEY `scht` (`scht`)"
 Str(1, 148, 31) = "  KEY `docname` (`docname`)"
 Str(1, 148, 32) = "  KEY `pat_id` (`Pat_id`)"
 Str(1, 148, 33) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=FIXED"
End Sub ' FüllStr148

Sub FüllStr149()
 Str(0, 149, 0) = "pauschalen"
 Str(0, 149, 1) = "`Leistung`"
 Str(0, 149, 2) = "`Betreuung`"
 ArtZ(0, 149) = 2
 Str(1, 149, 0) = "CREATE TABLE `pauschalen` ("
 Str(1, 149, 1) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 149, 2) = " `Betreuung` bit(1) DEFAULT NULL"
 Str(1, 149, 3) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr149

Sub FüllStr150()
 Str(0, 150, 0) = "pumpenträger"
 Str(0, 150, 1) = "`fall`"
 Str(0, 150, 2) = "`anrede`"
 Str(0, 150, 3) = "`nachname`"
 Str(0, 150, 4) = "`vorname`"
 Str(0, 150, 5) = "`gebdat`"
 Str(0, 150, 6) = "`privattel`"
 Str(0, 150, 7) = "`privattel_2`"
 Str(0, 150, 8) = "`privatfax`"
 Str(0, 150, 9) = "`diensttel`"
 Str(0, 150, 10) = "`straße`"
 Str(0, 150, 11) = "`plz`"
 Str(0, 150, 12) = "`ort`"
 Str(0, 150, 13) = "`mail1`"
 Str(0, 150, 14) = "`mail2`"
 Str(0, 150, 15) = "`email`"
 ArtZ(0, 150) = 15
 Str(1, 150, 0) = "CREATE TABLE `pumpenträger` ("
 Str(1, 150, 1) = " `fall` varchar(2) DEFAULT NULL"
 Str(1, 150, 2) = " `anrede` varchar(4) DEFAULT NULL"
 Str(1, 150, 3) = " `nachname` varchar(21) DEFAULT NULL"
 Str(1, 150, 4) = " `vorname` varchar(19) DEFAULT NULL"
 Str(1, 150, 5) = " `gebdat` datetime DEFAULT NULL"
 Str(1, 150, 6) = " `privattel` varchar(100) DEFAULT NULL"
 Str(1, 150, 7) = " `privattel_2` varchar(50) DEFAULT NULL"
 Str(1, 150, 8) = " `privatfax` varchar(50) DEFAULT NULL"
 Str(1, 150, 9) = " `diensttel` varchar(50) DEFAULT NULL"
 Str(1, 150, 10) = " `straße` varchar(50) DEFAULT NULL"
 Str(1, 150, 11) = " `plz` varchar(20) DEFAULT NULL"
 Str(1, 150, 12) = " `ort` varchar(70) DEFAULT NULL"
 Str(1, 150, 13) = " `mail1` varchar(100) DEFAULT NULL"
 Str(1, 150, 14) = " `mail2` varchar(100) DEFAULT NULL"
 Str(1, 150, 15) = " `email` varchar(100) DEFAULT NULL"
 Str(1, 150, 16) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr150

Sub FüllStr151()
 Str(0, 151, 0) = "qsumme"
 Str(1, 151, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `qsumme` AS SELECT sum(`_qsumme`.`wert`) AS `summe`,`_qsumme`.`quartal` AS `quartal` FROM `_qsumme` group by `_qsumme`.`quartal` ORDER BY substr(`_qsumme`.`quartal`,2) desc,`_qsumme`.`quartal` desc"
End Sub ' FüllStr151

Sub FüllStr152()
 Str(0, 152, 0) = "relationen"
 Str(0, 152, 1) = "`ccolumn`"
 Str(0, 152, 2) = "`grbit`"
 Str(0, 152, 3) = "`icolumn`"
 Str(0, 152, 4) = "`szColumn`"
 Str(0, 152, 5) = "`szObject`"
 Str(0, 152, 6) = "`szReferencedColumn`"
 Str(0, 152, 7) = "`szReferencedObject`"
 Str(0, 152, 8) = "`szRelationship`"
 Str(0, 152, 9) = "`szObject`"
 Str(0, 152, 10) = "`szReferencedObject`"
 Str(0, 152, 11) = "`szRelationship`"
 ArtZ(0, 152) = 8
 ArtZ(1, 152) = 3
 Str(1, 152, 0) = "CREATE TABLE `relationen` ("
 Str(1, 152, 1) = " `ccolumn` int(10) DEFAULT NULL"
 Str(1, 152, 2) = " `grbit` int(10) DEFAULT NULL"
 Str(1, 152, 3) = " `icolumn` int(10) DEFAULT NULL"
 Str(1, 152, 4) = " `szColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 152, 5) = " `szObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 152, 6) = " `szReferencedColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 152, 7) = " `szReferencedObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 152, 8) = " `szRelationship` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 152, 9) = "  KEY `szObject` (`szObject`)"
 Str(1, 152, 10) = "  KEY `szReferencedObject` (`szReferencedObject`)"
 Str(1, 152, 11) = "  KEY `szRelationship` (`szRelationship`)"
 Str(1, 152, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr152

Sub FüllStr153()
 Str(0, 153, 0) = "rezepteintraege"
 Str(0, 153, 1) = "`FID`"
 Str(0, 153, 2) = "`Pat_ID`"
 Str(0, 153, 3) = "`ZeitPunkt`"
 Str(0, 153, 4) = "`Rezept`"
 Str(0, 153, 5) = "`Rezeptklasse`"
 Str(0, 153, 6) = "`Rezklkurz`"
 Str(0, 153, 7) = "`Rezkllang`"
 Str(0, 153, 8) = "`kbez`"
 Str(0, 153, 9) = "`Medikament`"
 Str(0, 153, 10) = "`PZN`"
 Str(0, 153, 11) = "`absPos`"
 Str(0, 153, 12) = "`AktZeit`"
 Str(0, 153, 13) = "`QS`"
 Str(0, 153, 14) = "`QT`"
 Str(0, 153, 15) = "`StByte`"
 Str(0, 153, 16) = "`LANRid`"
 Str(0, 153, 17) = "`id`"
 Str(0, 153, 18) = "`id`"
 Str(0, 153, 19) = "`Auswahl`"
 Str(0, 153, 20) = "`FälleRezeptEinträge`"
 Str(0, 153, 21) = "`FID`"
 Str(0, 153, 22) = "`NamenRezeptEinträge`"
 Str(0, 153, 23) = "`F??lleRezeptEintr??ge_AccRel`"
 ArtZ(0, 153) = 17
 ArtZ(1, 153) = 5
 ArtZ(2, 153) = 1
 Str(1, 153, 0) = "CREATE TABLE `rezepteintraege` ("
 Str(1, 153, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 153, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 153, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 153, 4) = " `Rezept` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6210, 3652(1), 6218(1)'"
 Str(1, 153, 5) = " `Rezeptklasse` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)'"
 Str(1, 153, 6) = " `Rezklkurz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'letztes Split-Feld, z.B. ''rp'' oder ''lar'''"
 Str(1, 153, 7) = " `Rezkllang` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erstes Split-Feld, z.B. ''Sprechstundenbedarf'''"
 Str(1, 153, 8) = " `kbez` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'vorletztes Split-Feld, z.B. ''DTronAnthra'''"
 Str(1, 153, 9) = " `Medikament` varchar(236) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3652(2), 6218(4)'"
 Str(1, 153, 10) = " `PZN` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6210(2), 6218(3)'"
 Str(1, 153, 11) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 153, 12) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 153, 13) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 153, 14) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 153, 15) = " `StByte` int(11) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 153, 16) = " `LANRid` int(3) unsigned NOT NULL COMMENT 'Bezug auf lanrpraxis.id'"
 Str(1, 153, 17) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 153, 18) = "  PRIMARY KEY (`id`)"
 Str(1, 153, 19) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Rezept`,`Medikament`)"
 Str(1, 153, 20) = "  KEY `FälleRezeptEinträge` (`FID`)"
 Str(1, 153, 21) = "  KEY `FID` (`FID`)"
 Str(1, 153, 22) = "  KEY `NamenRezeptEinträge` (`Pat_ID`)"
 Str(1, 153, 23) = "  CONSTRAINT `F??lleRezeptEintr??ge_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 153, 24) = " ENGINE=InnoDB AUTO_INCREMENT=13741767 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr153

Sub FüllStr154()
 Str(0, 154, 0) = "rr"
 Str(0, 154, 1) = "`FID`"
 Str(0, 154, 2) = "`Pat_ID`"
 Str(0, 154, 3) = "`ZeitPunkt`"
 Str(0, 154, 4) = "`RR`"
 Str(0, 154, 5) = "`Puls`"
 Str(0, 154, 6) = "`Quelle`"
 Str(0, 154, 7) = "`Bemerkung`"
 Str(0, 154, 8) = "`absPos`"
 Str(0, 154, 9) = "`AktZeit`"
 Str(0, 154, 10) = "`StByte`"
 Str(0, 154, 11) = "`Auswahl`"
 Str(0, 154, 12) = "`FälleRR`"
 Str(0, 154, 13) = "`FID`"
 Str(0, 154, 14) = "`NamenRR`"
 Str(0, 154, 15) = "`F??lleRR_AccRel`"
 ArtZ(0, 154) = 10
 ArtZ(1, 154) = 4
 ArtZ(2, 154) = 1
 Str(1, 154, 0) = "CREATE TABLE `rr` ("
 Str(1, 154, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 154, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 154, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 154, 4) = " `RR` longtext COLLATE latin1_german2_ci COMMENT '6230'"
 Str(1, 154, 5) = " `Puls` int(3) unsigned DEFAULT NULL COMMENT 'Puls'"
 Str(1, 154, 6) = " `Quelle` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Informationsquelle'"
 Str(1, 154, 7) = " `Bemerkung` varchar(294) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Bemerkung'"
 Str(1, 154, 8) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 154, 9) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 154, 10) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 154, 11) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`RR`(255))"
 Str(1, 154, 12) = "  KEY `FälleRR` (`FID`)"
 Str(1, 154, 13) = "  KEY `FID` (`FID`)"
 Str(1, 154, 14) = "  KEY `NamenRR` (`Pat_ID`)"
 Str(1, 154, 15) = "  CONSTRAINT `F??lleRR_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 154, 16) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr154

Sub FüllStr155()
 Str(0, 155, 0) = "rrparse"
 Str(0, 155, 1) = "`Pat_id`"
 Str(0, 155, 2) = "`Zeitpunkt`"
 Str(0, 155, 3) = "`RRSyst`"
 Str(0, 155, 4) = "`RRDiast`"
 Str(0, 155, 5) = "`Quelle`"
 Str(0, 155, 6) = "`ID`"
 Str(0, 155, 7) = "`rrparsenamen`"
 ArtZ(0, 155) = 5
 ArtZ(1, 155) = 1
 ArtZ(2, 155) = 1
 Str(1, 155, 0) = "CREATE TABLE `rrparse` ("
 Str(1, 155, 1) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 155, 2) = " `Zeitpunkt` datetime DEFAULT NULL"
 Str(1, 155, 3) = " `RRSyst` smallint(6) DEFAULT NULL"
 Str(1, 155, 4) = " `RRDiast` smallint(6) DEFAULT NULL"
 Str(1, 155, 5) = " `Quelle` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 155, 6) = "  KEY `ID` (`Pat_id`,`Zeitpunkt`,`RRSyst`,`RRDiast`)"
 Str(1, 155, 7) = "  CONSTRAINT `rrparsenamen` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 155, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr155

Sub FüllStr156()
 Str(0, 156, 0) = "tabfuell"
 Str(0, 156, 1) = "`id`"
 Str(0, 156, 2) = "`zp`"
 Str(0, 156, 3) = "`tabname`"
 Str(0, 156, 4) = "`fehler`"
 Str(0, 156, 5) = "`tabtype`"
 Str(0, 156, 6) = "`ds`"
 Str(0, 156, 7) = "`id`"
 Str(0, 156, 8) = "`tabname`"
 Str(0, 156, 9) = "`zp`"
 ArtZ(0, 156) = 6
 ArtZ(1, 156) = 3
 Str(1, 156, 0) = "CREATE TABLE `tabfuell` ("
 Str(1, 156, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 156, 2) = " `zp` datetime DEFAULT NULL"
 Str(1, 156, 3) = " `tabname` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 156, 4) = " `fehler` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 156, 5) = " `tabtype` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 156, 6) = " `ds` int(10) DEFAULT NULL"
 Str(1, 156, 7) = "  PRIMARY KEY (`id`)"
 Str(1, 156, 8) = "  KEY `tabname` (`tabname`)"
 Str(1, 156, 9) = "  KEY `zp` (`zp`)"
 Str(1, 156, 10) = " ENGINE=MyISAM AUTO_INCREMENT=455 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr156

Sub FüllStr157()
 Str(0, 157, 0) = "terminakt"
 Str(0, 157, 1) = "`id`"
 Str(0, 157, 2) = "`abgerufen`"
 Str(0, 157, 3) = "`aktzeit`"
 Str(0, 157, 4) = "`id`"
 Str(0, 157, 5) = "`aktzeit`"
 ArtZ(0, 157) = 3
 ArtZ(1, 157) = 2
 Str(1, 157, 0) = "CREATE TABLE `terminakt` ("
 Str(1, 157, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 157, 2) = " `abgerufen` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Datei?nderung TMFTools'"
 Str(1, 157, 3) = " `aktzeit` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Aktualisierung'"
 Str(1, 157, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 157, 5) = "  KEY `aktzeit` (`aktzeit`,`abgerufen`)"
 Str(1, 157, 6) = " ENGINE=InnoDB AUTO_INCREMENT=129828 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='?nderungszeitpunkt von TMFTools'"
End Sub ' FüllStr157

Sub FüllStr158()
 Str(0, 158, 0) = "termine"
 Str(0, 158, 1) = "`id`"
 Str(0, 158, 2) = "`taid`"
 Str(0, 158, 3) = "`pid`"
 Str(0, 158, 4) = "`zp`"
 Str(0, 158, 5) = "`raum`"
 Str(0, 158, 6) = "`zusatz`"
 Str(0, 158, 7) = "`aktzeit`"
 Str(0, 158, 8) = "`id`"
 Str(0, 158, 9) = "`zp`"
 Str(0, 158, 10) = "`pid`"
 Str(0, 158, 11) = "`taid`"
 Str(0, 158, 12) = "`raum`"
 ArtZ(0, 158) = 7
 ArtZ(1, 158) = 5
 Str(1, 158, 0) = "CREATE TABLE `termine` ("
 Str(1, 158, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT COMMENT 'eindeutige ID, Prim?rschl?ssel'"
 Str(1, 158, 2) = " `taid` int(10) DEFAULT NULL COMMENT 'Bezug auf `terminakt`'"
 Str(1, 158, 3) = " `pid` int(10) DEFAULT NULL COMMENT 'Pat_Id aus `namen`'"
 Str(1, 158, 4) = " `zp` datetime DEFAULT NULL COMMENT 'Terminzeitpunkt'"
 Str(1, 158, 5) = " `raum` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Terminart'"
 Str(1, 158, 6) = " `zusatz` varchar(400) CHARACTER SET utf8 COLLATE utf8_unicode_ci DEFAULT NULL COMMENT 'Terminbeschreibung und eintragender Mitarbeiter'"
 Str(1, 158, 7) = " `aktzeit` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Aktualisierung'"
 Str(1, 158, 8) = "  PRIMARY KEY (`id`)"
 Str(1, 158, 9) = "  KEY `zp` (`zp`)"
 Str(1, 158, 10) = "  KEY `pid` (`pid`)"
 Str(1, 158, 11) = "  KEY `taid` (`taid`)"
 Str(1, 158, 12) = "  KEY `raum` (`raum`(30))"
 Str(1, 158, 13) = " ENGINE=InnoDB AUTO_INCREMENT=3343509 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='?nderungszeitpunkt von TMFTools'"
End Sub ' FüllStr158

Sub FüllStr159()
 Str(0, 159, 0) = "test0"
 Str(1, 159, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `test0` AS SELECT `u`.`Pat_id` AS `Pat_ID`,cast(`u`.`Eingang` as date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,if(isnull(`n2`.`Abkü`),`w`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`w`.`Abkü` AS `abk_ur`,`w`.`Langname` AS `Langtext`,`w`.`Wert` AS `Wert`,if(isnull(`n2`.`Abkü`),`w`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`w`.`Einheit` AS `Einheit_ur`,concat(if((`w`.`Erklärung` regexp '^:[ /*:]*$'),'',if((`w`.`Erklärung` regexp '^:[ /*]*:'),concat(substr(`w`.`Erklärung`,(locate(':',`w`.`Erklärung`,2) + 1)),';'),if((`w`.`Erklärung` = '.'),'',if((`w`.`Erklärung` = ''),'',concat(`w`.`Erklärung`,';'))))),`w`.`Kommentar`) AS `Kommentar`,if(isnull(`n2`.`Abkü`),`nb`.`NB`,`nb2`.`NB`) AS `NB`,`nb`.`NB` AS `NB_ur`,if(isnull(`n2`.`Abkü`),`nb`.`uNg`,`nb2`.`uNg`) AS `uNg`,`nb`.`uNg` AS `uNg_ur`,if(isnull(`n2`.`Abkü`),`n" & _
  "b`.`oNg`,`nb2`.`oNg`) AS `oNg`,`nb`.`oNg` AS `oNg_ur`,`s`.`Labor` AS `Labor` FROM ((((((((`laborxus` `u` LEFT JOIN `namen` `na` on((`u`.`Pat_id` = `na`.`Pat_ID`))) LEFT JOIN `laborxwert` `w` on((`u`.`RefNr` = `w`.`RefNr`))) LEFT JOIN `laborxpnb` `nb` on((`w`.`nbid` = `nb`.`id`))) LEFT JOIN `laborxpneu` `neu` on((`nb`.`pid` = `neu`.`id`))) LEFT JOIN `laborxpgl` `gl` on((`neu`.`id` = `gl`.`idxpneu`))) LEFT JOIN `laborxpneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) LEFT JOIN `laborxpnb` `nb2` on(((`n2`.`id` = `nb2`.`pid`) and (`nb2`.`Geschlecht` in (3,9,0,if((`na`.`Geschlecht` = 'm'),if(((`na`.`GebDat` + interval 18 year) > cast(`u`.`Eingang` as date)),1,4),if(((`na`.`GebDat` + interval 18 year) > cast(`u`.`Eingang` as date)),2,5))))))) LEFT JOIN `laborxsaetze` `s` on((`u`.`SatzID` = `s`.`SatzID`))) WHERE (`u`.`Pat_id` = 14)"
End Sub ' FüllStr159

Sub FüllStr160()
 Str(0, 160, 0) = "test2"
 Str(1, 160, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `test2` AS SELECT `nb2`.`Eingang` AS `eingang`,`u`.`Pat_id` AS `Pat_ID`,cast(`u`.`Eingang` as date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,if(isnull(`n2`.`Abkü`),`w`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`w`.`Abkü` AS `abk_ur`,`w`.`Langname` AS `Langtext`,`w`.`Wert` AS `Wert`,if(isnull(`n2`.`Abkü`),`w`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`w`.`Einheit` AS `Einheit_ur`,concat(if((`w`.`Erklärung` regexp '^:[ /*:]*$'),'',if((`w`.`Erklärung` regexp '^:[ /*]*:'),concat(substr(`w`.`Erklärung`,(locate(':',`w`.`Erklärung`,2) + 1)),';'),if((`w`.`Erklärung` = '.'),'',if((`w`.`Erklärung` = ''),'',concat(`w`.`Erklärung`,';'))))),`w`.`Kommentar`) AS `Kommentar`,if(isnull(`n2`.`Abkü`),`nb`.`NB`,`nb2`.`NB`) AS `NB`,`nb`.`NB` AS `NB_ur`,if(isnull(`n2`.`Abkü`),`nb`.`uNg`,`nb2`.`uNg`) AS `uNg`,`nb`.`uNg` AS `uNg_" & _
  "ur`,if(isnull(`n2`.`Abkü`),`nb`.`oNg`,`nb2`.`oNg`) AS `oNg`,`nb`.`oNg` AS `oNg_ur`,`s`.`Labor` AS `Labor` FROM ((((((((`laborxus` `u` LEFT JOIN `namen` `na` on((`u`.`Pat_id` = `na`.`Pat_ID`))) LEFT JOIN `laborxwert` `w` on((`u`.`RefNr` = `w`.`RefNr`))) LEFT JOIN `laborxpnb` `nb` on((`w`.`nbid` = `nb`.`id`))) LEFT JOIN `laborxpneu` `neu` on((`nb`.`pid` = `neu`.`id`))) LEFT JOIN `laborxpgl` `gl` on((`neu`.`id` = `gl`.`idxpneu`))) LEFT JOIN `laborxpneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) LEFT JOIN `laborxpnb` `nb2` on(((`n2`.`id` = `nb2`.`pid`) and (`nb2`.`Geschlecht` in (3,9,0,if((`na`.`Geschlecht` = 'm'),if(((`na`.`GebDat` + interval 18 year) > cast(`u`.`Eingang` as date)),1,4),if(((`na`.`GebDat` + interval 18 year) > cast(`u`.`Eingang` as date)),2,5))))))) LEFT JOIN `laborxsaetze` `s` on((`u`.`SatzID` = `s`.`SatzID`))) WHERE ((`u`.`Pat_id` = 14) and (`n2`.`Abkü` = 'alkp02') and (`nb2" & _
  "`.`Eingang` = (SELECT max(`laborxpnb`.`Eingang`) FROM `laborxpnb` WHERE (`laborxpnb`.`pid` = `n2`.`id`)))) group by `w`.`RefNr`,`w`.`Abkü` ORDER BY `nb2`.`Eingang` desc"
End Sub ' FüllStr160

Sub FüllStr161()
 Str(0, 161, 0) = "testid"
 Str(0, 161, 1) = "`id`"
 Str(0, 161, 2) = "`name`"
 Str(0, 161, 3) = "`id`"
 ArtZ(0, 161) = 2
 ArtZ(1, 161) = 1
 Str(1, 161, 0) = "CREATE TABLE `testid` ("
 Str(1, 161, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 161, 2) = " `name` char(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 161, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 161, 4) = " ENGINE=MyISAM AUTO_INCREMENT=8011 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr161

Sub FüllStr162()
 Str(0, 162, 0) = "therarten"
 Str(0, 162, 1) = "`id`"
 Str(0, 162, 2) = "`pat_id`"
 Str(0, 162, 3) = "`zp`"
 Str(0, 162, 4) = "`mpnr`"
 Str(0, 162, 5) = "`therart`"
 Str(0, 162, 6) = "`absPos`"
 Str(0, 162, 7) = "`AktZeit`"
 Str(0, 162, 8) = "`StByte`"
 Str(0, 162, 9) = "`id`"
 Str(0, 162, 10) = "`pat_id`"
 ArtZ(0, 162) = 8
 ArtZ(1, 162) = 2
 Str(1, 162, 0) = "CREATE TABLE `therarten` ("
 Str(1, 162, 1) = " `id` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 162, 2) = " `pat_id` int(11) DEFAULT NULL"
 Str(1, 162, 3) = " `zp` date DEFAULT NULL"
 Str(1, 162, 4) = " `mpnr` int(11) DEFAULT NULL"
 Str(1, 162, 5) = " `therart` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 162, 6) = " `absPos` int(10) DEFAULT NULL"
 Str(1, 162, 7) = " `AktZeit` datetime DEFAULT NULL"
 Str(1, 162, 8) = " `StByte` int(10) DEFAULT NULL"
 Str(1, 162, 9) = "  PRIMARY KEY (`id`)"
 Str(1, 162, 10) = "  KEY `pat_id` (`pat_id`)"
 Str(1, 162, 11) = " ENGINE=MyISAM AUTO_INCREMENT=398914 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr162

Sub FüllStr163()
 Str(0, 163, 0) = "tmpfif"
 Str(0, 163, 1) = "`FeldVW`"
 Str(0, 163, 2) = "`Feld`"
 Str(0, 163, 3) = "`StByte`"
 Str(0, 163, 4) = "`FeldVW`"
 Str(0, 163, 5) = "`Feld`"
 ArtZ(0, 163) = 3
 ArtZ(1, 163) = 2
 Str(1, 163, 0) = "CREATE TABLE `tmpfif` ("
 Str(1, 163, 1) = " `FeldVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 163, 2) = " `Feld` longtext COLLATE latin1_german2_ci"
 Str(1, 163, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 163, 4) = "  PRIMARY KEY (`FeldVW`)"
 Str(1, 163, 5) = "  KEY `Feld` (`Feld`(255))"
 Str(1, 163, 6) = " ENGINE=InnoDB AUTO_INCREMENT=13279 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr163

Sub FüllStr164()
 Str(0, 164, 0) = "tmprelationen"
 Str(0, 164, 1) = "`ccolumn`"
 Str(0, 164, 2) = "`grbit`"
 Str(0, 164, 3) = "`icolumn`"
 Str(0, 164, 4) = "`szColumn`"
 Str(0, 164, 5) = "`szObject`"
 Str(0, 164, 6) = "`szReferencedColumn`"
 Str(0, 164, 7) = "`szReferencedObject`"
 Str(0, 164, 8) = "`szRelationship`"
 ArtZ(0, 164) = 8
 Str(1, 164, 0) = "CREATE TABLE `tmprelationen` ("
 Str(1, 164, 1) = " `ccolumn` int(10) DEFAULT NULL"
 Str(1, 164, 2) = " `grbit` int(10) DEFAULT NULL"
 Str(1, 164, 3) = " `icolumn` int(10) DEFAULT NULL"
 Str(1, 164, 4) = " `szColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 164, 5) = " `szObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 164, 6) = " `szReferencedColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 164, 7) = " `szReferencedObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 164, 8) = " `szRelationship` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 164, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr164

Sub FüllStr165()
 Str(0, 165, 0) = "ueberwvon"
 Str(0, 165, 1) = "`ID`"
 Str(0, 165, 2) = "`KVNr`"
 Str(0, 165, 3) = "`Titel`"
 Str(0, 165, 4) = "`Vorname`"
 Str(0, 165, 5) = "`Zusatz`"
 Str(0, 165, 6) = "`Nachname`"
 Str(0, 165, 7) = "`ID`"
 Str(0, 165, 8) = "`KVNr`"
 Str(0, 165, 9) = "`Name`"
 ArtZ(0, 165) = 6
 ArtZ(1, 165) = 3
 Str(1, 165, 0) = "CREATE TABLE `ueberwvon` ("
 Str(1, 165, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 165, 2) = " `KVNr` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 165, 3) = " `Titel` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 165, 4) = " `Vorname` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 165, 5) = " `Zusatz` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 165, 6) = " `Nachname` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 165, 7) = "  PRIMARY KEY (`ID`) USING BTREE"
 Str(1, 165, 8) = "  KEY `KVNr` (`KVNr`)"
 Str(1, 165, 9) = "  KEY `Name` (`Nachname`,`Vorname`)"
 Str(1, 165, 10) = " ENGINE=InnoDB AUTO_INCREMENT=6424 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='4247 Überwiesen von'"
End Sub ' FüllStr165

Sub FüllStr166()
 Str(0, 166, 0) = "unbekannte kennungen"
 Str(0, 166, 1) = "`Kennung`"
 Str(0, 166, 2) = "`absPos`"
 Str(0, 166, 3) = "`StByte`"
 Str(0, 166, 4) = "`Pat_id`"
 Str(0, 166, 5) = "`Inhalt`"
 Str(0, 166, 6) = "`unbekannte kennungennamen`"
 Str(0, 166, 7) = "`Kennung`"
 Str(0, 166, 8) = "`unbekannte?kennungennamen`"
 ArtZ(0, 166) = 5
 ArtZ(1, 166) = 2
 ArtZ(2, 166) = 1
 Str(1, 166, 0) = "CREATE TABLE `unbekannte kennungen` ("
 Str(1, 166, 1) = " `Kennung` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 166, 2) = " `absPos` int(10) DEFAULT NULL"
 Str(1, 166, 3) = " `StByte` int(10) DEFAULT NULL"
 Str(1, 166, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT 'zugehöriger Patient für spätere Ermittlungen'"
 Str(1, 166, 5) = " `Inhalt` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Inhalt Zeile zum Wiederauffinden'"
 Str(1, 166, 6) = "  KEY `unbekannte kennungennamen` (`Pat_id`)"
 Str(1, 166, 7) = "  KEY `Kennung` (`Kennung`) USING BTREE"
 Str(1, 166, 8) = "  CONSTRAINT `unbekannte?kennungennamen` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 166, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr166

Sub FüllStr167()
 Str(0, 167, 0) = "versorgungsamt oberfranken"
 Str(1, 167, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `versorgungsamt oberfranken` AS SELECT `i`.`titel` AS `titel`,`i`.`tsid` AS `tsid`,`i`.`transe` AS `transe`,`i`.`transs` AS `transs`,`i`.`id` AS `id`,`i`.`fsize` AS `fsize`,`i`.`pages` AS `pages`,`i`.`devname` AS `devname`,`i`.`retries` AS `retries`,`i`.`csid` AS `csid`,`i`.`routi` AS `routi`,`i`.`callerid` AS `callerid` FROM `inca` `i` WHERE ((`i`.`tsid` like '%803599%') and (`i`.`pages` > 1)) ORDER BY `i`.`transe` desc"
End Sub ' FüllStr167

Sub FüllStr168()
 Str(0, 168, 0) = "verz"
 Str(1, 168, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `verz` AS SELECT `vzo`.`zid` AS `id`,`vzo`.`fid` AS `fid`,`verze`.`verz` AS `verz`,`vzo`.`geändert` AS `geändert` FROM (`vzo` LEFT JOIN `verze` on((`vzo`.`vid` = `verze`.`id`)))"
End Sub ' FüllStr168

Sub FüllStr169()
 Str(0, 169, 0) = "verze"
 Str(0, 169, 1) = "`id`"
 Str(0, 169, 2) = "`verz`"
 Str(0, 169, 3) = "`id`"
 Str(0, 169, 4) = "`verz`"
 ArtZ(0, 169) = 2
 ArtZ(1, 169) = 2
 Str(1, 169, 0) = "CREATE TABLE `verze` ("
 Str(1, 169, 1) = " `id` int(2) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf Verzeichnis'"
 Str(1, 169, 2) = " `verz` varchar(255) CHARACTER SET latin1 COLLATE latin1_german1_ci NOT NULL COMMENT 'Verzeichnis (eindeutig)'"
 Str(1, 169, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 169, 4) = "  UNIQUE KEY `verz` (`verz`)"
 Str(1, 169, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr169

Sub FüllStr170()
 Str(0, 170, 0) = "vzo"
 Str(0, 170, 1) = "`zid`"
 Str(0, 170, 2) = "`fid`"
 Str(0, 170, 3) = "`vid`"
 Str(0, 170, 4) = "`geändert`"
 Str(0, 170, 5) = "`zid`"
 Str(0, 170, 6) = "`vzo`"
 Str(0, 170, 7) = "`geändert`"
 Str(0, 170, 8) = "`fid`"
 Str(0, 170, 9) = "`FKid`"
 Str(0, 170, 10) = "`vzo`"
 ArtZ(0, 170) = 4
 ArtZ(1, 170) = 4
 ArtZ(2, 170) = 2
 Str(1, 170, 0) = "CREATE TABLE `vzo` ("
 Str(1, 170, 1) = " `zid` int(2) unsigned NOT NULL AUTO_INCREMENT COMMENT 'eindeutige ID dieser Tabelle'"
 Str(1, 170, 2) = " `fid` int(2) unsigned NOT NULL COMMENT 'Bezug auf Faxe'"
 Str(1, 170, 3) = " `vid` int(2) unsigned NOT NULL COMMENT 'Bezug auf verze'"
 Str(1, 170, 4) = " `geändert` datetime NOT NULL COMMENT 'Datum des Eintrags'"
 Str(1, 170, 5) = "  PRIMARY KEY (`zid`)"
 Str(1, 170, 6) = "  KEY `vzo` (`vid`)"
 Str(1, 170, 7) = "  KEY `geändert` (`geändert`)"
 Str(1, 170, 8) = "  KEY `fid` (`fid`)"
 Str(1, 170, 9) = "  CONSTRAINT `FKid` FOREIGN KEY (`fid`) REFERENCES `faxe` (`ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 170, 10) = "  CONSTRAINT `vzo` FOREIGN KEY (`vid`) REFERENCES `verze` (`id`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 170, 11) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr170

Sub FüllStr171()
 Str(0, 171, 0) = "werte_scheingruppen"
 Str(0, 171, 1) = "`schgr`"
 Str(0, 171, 2) = "`Erklärung`"
 Str(0, 171, 3) = "`schgr`"
 ArtZ(0, 171) = 2
 ArtZ(1, 171) = 1
 Str(1, 171, 0) = "CREATE TABLE `werte_scheingruppen` ("
 Str(1, 171, 1) = " `schgr` decimal(2,0) NOT NULL DEFAULT '0' COMMENT '4239'"
 Str(1, 171, 2) = " `Erklärung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 171, 3) = "  PRIMARY KEY (`schgr`) USING BTREE"
 Str(1, 171, 4) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Scheingruppen in Turbomed'"
End Sub ' FüllStr171

Sub FüllStr172()
 Str(0, 172, 0) = "werte_weggeldzonen"
 Str(0, 172, 1) = "`Weggeldzone`"
 Str(0, 172, 2) = "`Zonennr`"
 Str(0, 172, 3) = "`Bereich`"
 Str(0, 172, 4) = "`Weggeldzone`"
 Str(0, 172, 5) = "`Zonennr`"
 ArtZ(0, 172) = 3
 ArtZ(1, 172) = 2
 Str(1, 172, 0) = "CREATE TABLE `werte_weggeldzonen` ("
 Str(1, 172, 1) = " `Weggeldzone` varchar(2) COLLATE latin1_german2_ci NOT NULL COMMENT '3631 (1) Weggeldzone'"
 Str(1, 172, 2) = " `Zonennr` decimal(1,0) DEFAULT NULL COMMENT '3631 (2) Weggeldzonenziffer'"
 Str(1, 172, 3) = " `Bereich` varchar(30) COLLATE latin1_german2_ci NOT NULL COMMENT 'Kilometer-Bereich'"
 Str(1, 172, 4) = "  PRIMARY KEY (`Weggeldzone`) USING BTREE"
 Str(1, 172, 5) = "  KEY `Zonennr` (`Zonennr`)"
 Str(1, 172, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Weggeldzonen'"
End Sub ' FüllStr172

Sub FüllStr173()
 Str(0, 173, 0) = "zz"
 Str(0, 173, 1) = "`id`"
 Str(0, 173, 2) = "`v1`"
 Str(0, 173, 3) = "`v2`"
 Str(0, 173, 4) = "`i1`"
 Str(0, 173, 5) = "`i2`"
 Str(0, 173, 6) = "`id`"
 ArtZ(0, 173) = 5
 ArtZ(1, 173) = 1
 Str(1, 173, 0) = "CREATE TABLE `zz` ("
 Str(1, 173, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 173, 2) = " `v1` varchar(2) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 173, 3) = " `v2` varchar(1) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 173, 4) = " `i1` int(20) unsigned DEFAULT NULL"
 Str(1, 173, 5) = " `i2` decimal(15,0) NOT NULL"
 Str(1, 173, 6) = "  PRIMARY KEY (`id`)"
 Str(1, 173, 7) = " ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr173

Function doEx&(sql$, obtolerant%) ' SQL-Befehl ausführen, Fehler anzeigen
 Dim rAF&, FMeld$
 If obtolerant Then On Error Resume Next Else On Error GoTo fehler
 Call cnz.Execute(sql, rAF)
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
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
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
   Call doEx("USE `" & hDBn & "`", 0)
   Resume
  End If
End Select
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), "", CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doEx/" & AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function ' doEx

Function SplitN&(ByRef q$, Sep$, erg$()) ' da Split() Speicher fraß
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
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SplitN/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SplitN aus aufSplit

Public Function doMach_quelle(DBn$, Optional Server$, Optional obStumm% = True) ' Datenbankname
 Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$
 Dim i&, p1&, p2&, p3&, CLen&, CLen1&, obLT%
 Dim Index$()
 On Error Resume Next
 hDBn = DBn
 Open App.path & "\MachDB.bas_prot.txt" For Output As #302
 obProt = (Err.Number = 0)
 On Error GoTo fehler
 If LenB(Server) = 0 Then Server = GetServr(DBCn)
 cnzCStr = "PROVIDER=MSDASQL;driver={MySQL ODBC 5.3 Unicode Driver};server=linux1;uid=mysql;pwd=97a5o6;"
 Set cnz = Nothing
 cnz.Open cnzCStr
 Call doEx("create database if not exists `" & DBn & "` character set latin1 collate latin1_german2_ci;", 0)
 Call doEx("grant all privileges ON `" & DBn & "`.* to 'praxis'@'%' identified by 'sonne' with grant option", 0)
 Call doEx("grant all privileges ON `" & DBn & "`.* to 'praxis'@'localhost' identified by 'sonne' with grant option", 0)
 Call doEx("USE `" & DBn & "`", 0)
 Call doEx("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", 0)
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
 FüllStr28
 FüllStr29
 FüllStr30
 FüllStr31
 FüllStr32
 FüllStr33
 FüllStr34
 FüllStr35
 FüllStr36
 FüllStr37
 FüllStr38
 FüllStr39
 FüllStr40
 FüllStr41
 FüllStr42
 FüllStr43
 FüllStr44
 FüllStr45
 FüllStr46
 FüllStr47
 FüllStr48
 FüllStr49
 FüllStr50
 FüllStr51
 FüllStr52
 FüllStr53
 FüllStr54
 FüllStr55
 FüllStr56
 FüllStr57
 FüllStr58
 FüllStr59
 FüllStr60
 FüllStr61
 FüllStr62
 FüllStr63
 FüllStr64
 FüllStr65
 FüllStr66
 FüllStr67
 FüllStr68
 FüllStr69
 FüllStr70
 FüllStr71
 FüllStr72
 FüllStr73
 FüllStr74
 FüllStr75
 FüllStr76
 FüllStr77
 FüllStr78
 FüllStr79
 FüllStr80
 FüllStr81
 FüllStr82
 FüllStr83
 FüllStr84
 FüllStr85
 FüllStr86
 FüllStr87
 FüllStr88
 FüllStr89
 FüllStr90
 FüllStr91
 FüllStr92
 FüllStr93
 FüllStr94
 FüllStr95
 FüllStr96
 FüllStr97
 FüllStr98
 FüllStr99
 FüllStr100
 FüllStr101
 FüllStr102
 FüllStr103
 FüllStr104
 FüllStr105
 FüllStr106
 FüllStr107
 FüllStr108
 FüllStr109
 FüllStr110
 FüllStr111
 FüllStr112
 FüllStr113
 FüllStr114
 FüllStr115
 FüllStr116
 FüllStr117
 FüllStr118
 FüllStr119
 FüllStr120
 FüllStr121
 FüllStr122
 FüllStr123
 FüllStr124
 FüllStr125
 FüllStr126
 FüllStr127
 FüllStr128
 FüllStr129
 FüllStr130
 FüllStr131
 FüllStr132
 FüllStr133
 FüllStr134
 FüllStr135
 FüllStr136
 FüllStr137
 FüllStr138
 FüllStr139
 FüllStr140
 FüllStr141
 FüllStr142
 FüllStr143
 FüllStr144
 FüllStr145
 FüllStr146
 FüllStr147
 FüllStr148
 FüllStr149
 FüllStr150
 FüllStr151
 FüllStr152
 FüllStr153
 FüllStr154
 FüllStr155
 FüllStr156
 FüllStr157
 FüllStr158
 FüllStr159
 FüllStr160
 FüllStr161
 FüllStr162
 FüllStr163
 FüllStr164
 FüllStr165
 FüllStr166
 FüllStr167
 FüllStr168
 FüllStr169
 FüllStr170
 FüllStr171
 FüllStr172
 FüllStr173
 Call doEx("SET FOREIGN_KEY_CHECKS = 0", 0)

 Dim j&, ZZ&, Tbl$, sql As New CString
 For i = 0 To 173
  If InStrB(Str(1, i, 0), "CREATE TABLE") <> 0 Then
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
    Set rsc = Nothing
    rsc.Open "show CREATE TABLE `" & Tbl & "`", cnz, adOpenStatic, adLockReadOnly
    sct = rsc.Fields(1)
    If InStrB(sct, "CREATE ALGORITHM") = 1 Then
     FNr = doEx("drop view `" & Tbl & "`", 0)
     FNr = doEx(sql.Value, 0)
    Else
     Exit Do
    End If
   Loop
   If InStrB(AIoZ(sct), AIoZ(Str(1, i, ZZ))) = 0 Then
    Call doEx("ALTER TABLE `" & Tbl & "`" & Str(1, i, ZZ), 0)
   End If
   TMt.Clear
   SplitN sct, vbLf, Spli
   For j = 1 To ArtZ(0, i) ' Tabellenfelder
    Dim k&, enthalten%, genau%, Posi$
    enthalten = 0
    genau = 0
    k = 0
    Set rsc = Nothing
    rsc.Open "show columns FROM `" & Tbl & "` WHERE field = '" & Mid$(Str(0, i, j), 2, Len(Str(0, i, j)) - 2) & "'", cnz, adOpenStatic, adLockReadOnly
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
      Posi = " FIRST,"
     Else
      Posi = " AFTER " & Str(0, i, j - 1) & ","
     End If
     If Not enthalten Then
      TMt.AppVar (Array(" add ", Str(1, i, j), Posi))
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
       TMt.AppVar (Array(" modify ", Str(1, i, j), Posi))
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
   If TMt.length <> 0 Then
    TMt.Cut (TMt.length - 1)
    Call doEx("ALTER TABLE `" & Tbl & "` " & TMt.Value, -1)
   End If
  End If ' InStr(Str(1, i, 0), "CREATE TABLE") <> 0 Then
 Next i
 For i = 0 To 173
  If InStrB(Str(1, i, 0), "CREATE TABLE") <> 0 Then
   Tbl = Str(0, i, 0)
   ZZ = ArtZ(0, i) + ArtZ(1, i)
   Set rsc = Nothing
   rsc.Open "show CREATE TABLE `" & Tbl & "`", cnz, adOpenStatic, adLockReadOnly
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
 For runde = 0 To 4
  For i = 0 To 173
   If InStrB(Str(1, i, 0), "DEFINER VIEW") <> 0 Then
    Dim obCr%
    obCr = 0
    Set rsc = Nothing
    rsc.Open "show tables FROM `" & DBn & "` WHERE `tables_in_" & DBn & "` = """ & Str(0, i, 0) & """", cnz, adOpenStatic, adLockReadOnly
    If rsc.BOF Then
     obCr = True
    Else
     Set rsc = Nothing
     rsc.Open "show CREATE TABLE `" & Str(0, i, 0) & "`", cnz, adOpenStatic, adLockReadOnly
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
 Call doEx("set FOREIGN_KEY_CHECKS = 1", 0)
 If obProt Then Close #302
 If Not obStumm Then
  MsgBox "Fertig mit doMach_quelle(" & DBn & "," & Server & ")!"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), "", CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_quelle/" & AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_quelle

Function GetServr$(DBCn As ADODB.Connection)
Dim spos&, sp2&
spos = InStr(LCase$(DBCn), "server=")
If spos <> 0 Then
 sp2 = InStr(spos, DBCn, ";")
 If sp2 = 0 Then sp2 = Len(DBCn)
 GetServr = Mid$(DBCn, spos + 7, sp2 - spos - 7)
End If
End Function ' GetServr

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
