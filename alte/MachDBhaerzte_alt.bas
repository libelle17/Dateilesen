Attribute VB_Name = "Module2"
'Bauanleitung für eine Datenbank wie `//linux1/haerzte` vom 10.7.11 21:43:04
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.Connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 35, 21) As New CString, ArtZ&(3, 35)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "ÄrzteundPsychotherapeuten"
 Str(1, 0, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `ÄrzteundPsychotherapeuten` AS SELECT CONCAT(if((`titel`.`Titel` <> ''),CONCAT(`titel`.`Titel`,' '),''),`arzt`.`Vorname`,' ',`arzt`.`Nachname`) AS `Arzt`,`arzt`.`LANR` AS `Lanr`,CAST(GROUP_CONCAT(DISTINCT CONCAT(`bsart`.`BSArt`,' ',`bs`.`BSNR`,': ',if((`bs`.`Name` <> ''),CONCAT(`bs`.`Name`,': '),''),CONCAT(`bs`.`Straße`,' ',`bs`.`Hausnr`,', ',`bs`.`PLZ`,' ',`ort`.`Ort`)),if((`tel`.`Tel` <> ''),CONCAT(', Tel: ',`tel`.`Tel`),''),if((`fax`.`Fax` <> ''),CONCAT(', Fax: ',`fax`.`Fax`),''),if((`mail`.`Mail` <> ''),CONCAT(', Email: ',`mail`.`Mail`),''),repeat(' ',20) SEPARATOR ',  ') AS char charset utf8) AS `Betriebsstätten`,(SELECT GROUP_CONCAT(DISTINCT `fachrichtung`.`Fachrichtung` SEPARATOR ',') FROM (`arzt_has_fachrichtung` `ahg` LEFT JOIN `fachrichtung` on((`ahg`.`fachrichtung_id` = `fachrichtung`.`i" & _
  "dFachrichtung`))) WHERE (`ahg`.`arzt_id` = `arzt`.`idarzt`)) AS `Fachrichtung`,(SELECT GROUP_CONCAT(DISTINCT `zusatzbezeichnung`.`Zusatzbezeichnung` SEPARATOR ',') FROM (`arzt_has_zusatzbezeichnung` `ahg` LEFT JOIN `zusatzbezeichnung` on((`ahg`.`Zusatzbezeichnung_id` = `zusatzbezeichnung`.`idZusatzbezeichnung`))) WHERE (`ahg`.`arzt_id` = `arzt`.`idarzt`)) AS `Zusatzbezeichnung`,(SELECT GROUP_CONCAT(DISTINCT `weiterbildung`.`Weiterbildung` SEPARATOR ',') FROM (`arzt_has_weiterbildung` `ahg` LEFT JOIN `weiterbildung` on((`ahg`.`weiterbildung_id` = `weiterbildung`.`idWeiterbildung`))) WHERE (`ahg`.`arzt_id` = `arzt`.`idarzt`)) AS `Weiterbildung`,(SELECT GROUP_CONCAT(DISTINCT `genehmigung`.`genehmigung` SEPARATOR ',') FROM (`arzt_has_genehmigung` `ahg` LEFT JOIN `genehmigung` on((`ahg`.`genehmigung_id` = `genehmigung`.`idgenehmigung`))) WHERE (`ahg`.`arzt_id` = `arzt`.`idarzt`)) AS `Genehmig" & _
  "ung`,(SELECT GROUP_CONCAT(DISTINCT `vertragsangebot`.`vertragsangebot` SEPARATOR ',') FROM (`arzt_has_vertragsangebot` `ahg` LEFT JOIN `vertragsangebot` on((`ahg`.`vertragsangebot_id` = `vertragsangebot`.`idvertragsangebot`))) WHERE (`ahg`.`arzt_id` = `arzt`.`idarzt`)) AS `Vertragsangebot`,(SELECT GROUP_CONCAT(DISTINCT `leistungsumfang`.`leistungsumfang` SEPARATOR ',') FROM (`arzt_has_leistungsumfang` `ahg` LEFT JOIN `leistungsumfang` on((`ahg`.`leistungsumfang_id` = `leistungsumfang`.`idLeistungsumfang`))) WHERE (`ahg`.`arzt_id` = `arzt`.`idarzt`)) AS `Leistungsumfang`,(SELECT GROUP_CONCAT(DISTINCT `fremdsprache`.`fremdsprache` SEPARATOR ',') FROM (`arzt_has_fremdsprache` `ahg` LEFT JOIN `fremdsprache` on((`ahg`.`fremdsprache_id` = `fremdsprache`.`idfremdsprache`))) WHERE (`ahg`.`arzt_id` = `arzt`.`idarzt`)) AS `Fremdsprache`,`sprechzeiten`.`Sprechzeiten` AS `Sprechzeiten`,`arzt`.`aktze" & _
  "it` AS `AktZeit` FROM (((((((((`arzt` LEFT JOIN `arzt_has_bs` `ahb` on(((`arzt`.`idarzt` = `ahb`.`arzt_id`) AND (CAST(`ahb`.`aktzeit` AS date) = CAST((SELECT MAX(`arzt_has_bs`.`aktzeit`) FROM `arzt_has_bs` WHERE (`arzt_has_bs`.`arzt_id` = `arzt`.`idarzt`)) AS date))))) LEFT JOIN `bs` on((`bs`.`idbs` = `ahb`.`bs_id`))) LEFT JOIN `tel` on((`bs`.`idbs` = `tel`.`bs_id`))) LEFT JOIN `fax` on((`bs`.`idbs` = `fax`.`bs_id`))) LEFT JOIN `mail` on((`bs`.`idbs` = `mail`.`bs_id`))) LEFT JOIN `titel` on((`titel`.`idtitel` = `arzt`.`titel_id`))) LEFT JOIN `ort` on((`bs`.`Ort_id` = `ort`.`idOrt`))) LEFT JOIN `bsart` on((`bs`.`bsart_id` = `bsart`.`idbsart`))) LEFT JOIN `sprechzeiten` on((`bs`.`sprechzeiten_id` = `sprechzeiten`.`idsprechzeiten`))) WHERE (`ahb`.`bs_id` is not null) GROUP BY `arzt`.`idarzt`"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "Übersicht"
 Str(1, 1, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `Übersicht` AS SELECT COUNT(0) AS `Zahl`,CAST(`arzt_has_bs`.`aktzeit` AS date) AS `Datum` FROM `arzt_has_bs` GROUP BY CAST(`arzt_has_bs`.`aktzeit` AS date) ORDER BY CAST(`arzt_has_bs`.`aktzeit` AS date) DESC"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "ArztmitGenehmigungen"
 Str(1, 2, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `ArztmitGenehmigungen` AS SELECT `arzt`.`Nachname` AS `nachname`,`arzt`.`Vorname` AS `vorname`,`arzt`.`LANR` AS `lanr`,`ahg`.`genehmigung_id` AS `genehmigung_id`,`g`.`genehmigung` AS `genehmigung` FROM ((`arzt` LEFT JOIN `arzt_has_genehmigung` `ahg` on((`arzt`.`idarzt` = `ahg`.`arzt_id`))) LEFT JOIN `genehmigung` `g` on((`ahg`.`genehmigung_id` = `g`.`idgenehmigung`)))"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "RegelBetriebsstätten"
 Str(1, 3, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `RegelBetriebsstätten` AS SELECT CAST(GROUP_CONCAT(DISTINCT CONCAT(if((`titel`.`Titel` <> ''),CONCAT(`titel`.`Titel`,' '),''),`arzt`.`Vorname`,' ',`arzt`.`Nachname`,if((`arzt`.`LANR` is not null),CONCAT(' (LANR ',`arzt`.`LANR`,')'),'')) SEPARATOR ', ') AS char charset utf8) AS `Ärzte`,`bs`.`Name` AS `Name`,`bs`.`Straße` AS `Straße`,`bs`.`Hausnr` AS `Hausnr`,`bs`.`PLZ` AS `Plz`,`ort`.`Ort` AS `Ort`,`tel`.`Tel` AS `Tel`,`fax`.`Fax` AS `Fax`,`mail`.`Mail` AS `Mail`,`bs`.`BSNR` AS `Bsnr`,`bsart`.`BSArt` AS `BSArt`,`sprechzeiten`.`Sprechzeiten` AS `Sprechzeiten`,`bs`.`Rollst` AS `Rollst`,`bs`.`aktzeit` AS `Aktzeit`,(SELECT GROUP_CONCAT(DISTINCT `genehmigung`.`genehmigung` SEPARATOR ',') FROM ((`arzt_has_bs` `ahb` LEFT JOIN `arzt_has_genehmigung` `ahg` on((`ahg`.`arzt_id` = `ahb`.`arzt_id`))) LEFT JOIN `" & _
  "genehmigung` on((`ahg`.`genehmigung_id` = `genehmigung`.`idgenehmigung`))) WHERE (`ahb`.`bs_id` = `bs`.`idbs`)) AS `Genehmigung`,(SELECT GROUP_CONCAT(DISTINCT `fachrichtung`.`Fachrichtung` SEPARATOR ',') FROM ((`arzt_has_bs` `ahb` LEFT JOIN `arzt_has_fachrichtung` `ahg` on((`ahg`.`arzt_id` = `ahb`.`arzt_id`))) LEFT JOIN `fachrichtung` on((`ahg`.`fachrichtung_id` = `fachrichtung`.`idFachrichtung`))) WHERE (`ahb`.`bs_id` = `bs`.`idbs`)) AS `Fachrichtung`,(SELECT GROUP_CONCAT(DISTINCT `leistungsumfang`.`leistungsumfang` SEPARATOR ',') FROM ((`arzt_has_bs` `ahb` LEFT JOIN `arzt_has_leistungsumfang` `ahg` on((`ahg`.`arzt_id` = `ahb`.`arzt_id`))) LEFT JOIN `leistungsumfang` on((`ahg`.`leistungsumfang_id` = `leistungsumfang`.`idLeistungsumfang`))) WHERE (`ahb`.`bs_id` = `bs`.`idbs`)) AS `Leistungsumfang`,(SELECT GROUP_CONCAT(DISTINCT `vertragsangebot`.`vertragsangebot` SEPARATOR ',') FROM ((`ar" & _
  "zt_has_bs` `ahb` LEFT JOIN `arzt_has_vertragsangebot` `ahg` on((`ahg`.`arzt_id` = `ahb`.`arzt_id`))) LEFT JOIN `vertragsangebot` on((`ahg`.`vertragsangebot_id` = `vertragsangebot`.`idvertragsangebot`))) WHERE (`ahb`.`bs_id` = `bs`.`idbs`)) AS `Vertragsangebot`,(SELECT GROUP_CONCAT(DISTINCT `weiterbildung`.`Weiterbildung` SEPARATOR ',') FROM ((`arzt_has_bs` `ahb` LEFT JOIN `arzt_has_weiterbildung` `ahg` on((`ahg`.`arzt_id` = `ahb`.`arzt_id`))) LEFT JOIN `weiterbildung` on((`ahg`.`weiterbildung_id` = `weiterbildung`.`idWeiterbildung`))) WHERE (`ahb`.`bs_id` = `bs`.`idbs`)) AS `Weiterbildung`,(SELECT GROUP_CONCAT(DISTINCT `zusatzbezeichnung`.`Zusatzbezeichnung` SEPARATOR ',') FROM ((`arzt_has_bs` `ahb` LEFT JOIN `arzt_has_zusatzbezeichnung` `ahg` on((`ahg`.`arzt_id` = `ahb`.`arzt_id`))) LEFT JOIN `zusatzbezeichnung` on((`ahg`.`Zusatzbezeichnung_id` = `zusatzbezeichnung`.`idZusatzbezeichnung" & _
  "`))) WHERE (`ahb`.`bs_id` = `bs`.`idbs`)) AS `Zusatzbezeichnung`,(SELECT GROUP_CONCAT(DISTINCT `fremdsprache`.`fremdsprache` SEPARATOR ',') FROM ((`arzt_has_bs` `ahb` LEFT JOIN `arzt_has_fremdsprache` `ahg` on((`ahg`.`arzt_id` = `ahb`.`arzt_id`))) LEFT JOIN `fremdsprache` on((`ahg`.`fremdsprache_id` = `fremdsprache`.`idfremdsprache`))) WHERE (`ahb`.`bs_id` = `bs`.`idbs`)) AS `Fremdsprache` FROM (((((((((`bs` LEFT JOIN `arzt_has_bs` `ahb` on(((`bs`.`idbs` = `ahb`.`bs_id`) AND (CAST(`ahb`.`aktzeit` AS date) = CAST((SELECT MAX(`arzt_has_bs`.`aktzeit`) FROM `arzt_has_bs` WHERE (`arzt_has_bs`.`bs_id` = `bs`.`idbs`)) AS date))))) LEFT JOIN `arzt` on((`arzt`.`idarzt` = `ahb`.`arzt_id`))) LEFT JOIN `titel` on((`titel`.`idtitel` = `arzt`.`titel_id`))) LEFT JOIN `ort` on((`bs`.`Ort_id` = `ort`.`idOrt`))) LEFT JOIN `bsart` on((`bs`.`bsart_id` = `bsart`.`idbsart`))) LEFT JOIN `tel` on((`bs`.`idbs` =" & _
  " `tel`.`bs_id`))) LEFT JOIN `fax` on((`bs`.`idbs` = `fax`.`bs_id`))) LEFT JOIN `mail` on((`bs`.`idbs` = `mail`.`bs_id`))) LEFT JOIN `sprechzeiten` on((`bs`.`sprechzeiten_id` = `sprechzeiten`.`idsprechzeiten`))) WHERE (`ahb`.`bs_id` is not null) GROUP BY `bs`.`idbs`"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "arzt"
 Str(0, 4, 1) = "`idarzt`"
 Str(0, 4, 2) = "`Nachname`"
 Str(0, 4, 3) = "`Vorname`"
 Str(0, 4, 4) = "`titel_id`"
 Str(0, 4, 5) = "`Namenszusatz`"
 Str(0, 4, 6) = "`LANR`"
 Str(0, 4, 7) = "`nlart_id`"
 Str(0, 4, 8) = "`obweibl`"
 Str(0, 4, 9) = "`seit`"
 Str(0, 4, 10) = "`aktzeit`"
 Str(0, 4, 11) = "`idarzt`"
 Str(0, 4, 12) = "`lanr`"
 Str(0, 4, 13) = "`fk_arzt_titel1`"
 Str(0, 4, 14) = "`fk_arzt_nlart1`"
 Str(0, 4, 15) = "`namen`"
 Str(0, 4, 16) = "`fk_arzt_nlart1`"
 Str(0, 4, 17) = "`fk_arzt_titel1`"
 ArtZ(0, 4) = 10
 ArtZ(1, 4) = 5
 ArtZ(2, 4) = 2
 Str(1, 4, 0) = "CREATE TABLE `arzt` ("
 Str(1, 4, 1) = " `idarzt` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 4, 2) = " `Nachname` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 3) = " `Vorname` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 4) = " `titel_id` int(11) NOT NULL"
 Str(1, 4, 5) = " `Namenszusatz` varchar(14) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 6) = " `LANR` int(11) DEFAULT NULL"
 Str(1, 4, 7) = " `nlart_id` int(11) NOT NULL"
 Str(1, 4, 8) = " `obweibl` tinyint(1) unsigned NOT NULL DEFAULT '0'"
 Str(1, 4, 9) = " `seit` datetime DEFAULT NULL"
 Str(1, 4, 10) = " `aktzeit` datetime DEFAULT NULL"
 Str(1, 4, 11) = "  PRIMARY KEY (`idarzt`)"
 Str(1, 4, 12) = "  UNIQUE KEY `lanr` (`LANR`)"
 Str(1, 4, 13) = "  KEY `fk_arzt_titel1` (`titel_id`)"
 Str(1, 4, 14) = "  KEY `fk_arzt_nlart1` (`nlart_id`)"
 Str(1, 4, 15) = "  KEY `namen` (`Nachname`,`Vorname`)"
 Str(1, 4, 16) = "  CONSTRAINT `fk_arzt_nlart1` FOREIGN KEY (`nlart_id`) REFERENCES `nlart` (`idnlart`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 4, 17) = "  CONSTRAINT `fk_arzt_titel1` FOREIGN KEY (`titel_id`) REFERENCES `titel` (`idtitel`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 4, 18) = " ENGINE=InnoDB AUTO_INCREMENT=25653 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "arzt_has_bs"
 Str(0, 5, 1) = "`bs_id`"
 Str(0, 5, 2) = "`arzt_id`"
 Str(0, 5, 3) = "`obneben`"
 Str(0, 5, 4) = "`obang`"
 Str(0, 5, 5) = "`seit`"
 Str(0, 5, 6) = "`aktzeit`"
 Str(0, 5, 7) = "`bs_id`"
 Str(0, 5, 8) = "`fk_Betriebsstätte_has_hausarzt_Betriebsstätte1`"
 Str(0, 5, 9) = "`fk_Betriebsstätte_has_hausarzt_hausarzt1`"
 Str(0, 5, 10) = "`fk_Betriebsstätte_has_hausarzt_Betriebsstätte1`"
 Str(0, 5, 11) = "`fk_Betriebsstätte_has_hausarzt_hausarzt1`"
 ArtZ(0, 5) = 6
 ArtZ(1, 5) = 3
 ArtZ(2, 5) = 2
 Str(1, 5, 0) = "CREATE TABLE `arzt_has_bs` ("
 Str(1, 5, 1) = " `bs_id` int(11) NOT NULL"
 Str(1, 5, 2) = " `arzt_id` int(11) NOT NULL"
 Str(1, 5, 3) = " `obneben` tinyint(1) unsigned NOT NULL DEFAULT '0'"
 Str(1, 5, 4) = " `obang` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT 'ob Arzt hier angestellt'"
 Str(1, 5, 5) = " `seit` datetime DEFAULT NULL"
 Str(1, 5, 6) = " `aktzeit` datetime DEFAULT NULL"
 Str(1, 5, 7) = "  PRIMARY KEY (`bs_id`,`arzt_id`)"
 Str(1, 5, 8) = "  KEY `fk_Betriebsstätte_has_hausarzt_Betriebsstätte1` (`bs_id`)"
 Str(1, 5, 9) = "  KEY `fk_Betriebsstätte_has_hausarzt_hausarzt1` (`arzt_id`)"
 Str(1, 5, 10) = "  CONSTRAINT `fk_Betriebsstätte_has_hausarzt_Betriebsstätte1` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 5, 11) = "  CONSTRAINT `fk_Betriebsstätte_has_hausarzt_hausarzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 5, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "arzt_has_fachrichtung"
 Str(0, 6, 1) = "`arzt_id`"
 Str(0, 6, 2) = "`fachrichtung_id`"
 Str(0, 6, 3) = "`arzt_id`"
 Str(0, 6, 4) = "`fk_arzt_has_fachrichtung_arzt1`"
 Str(0, 6, 5) = "`fk_arzt_has_fachrichtung_fachrichtung1`"
 Str(0, 6, 6) = "`fk_arzt_has_fachrichtung_arzt1`"
 Str(0, 6, 7) = "`fk_arzt_has_fachrichtung_fachrichtung1`"
 ArtZ(0, 6) = 2
 ArtZ(1, 6) = 3
 ArtZ(2, 6) = 2
 Str(1, 6, 0) = "CREATE TABLE `arzt_has_fachrichtung` ("
 Str(1, 6, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 6, 2) = " `fachrichtung_id` int(11) NOT NULL"
 Str(1, 6, 3) = "  PRIMARY KEY (`arzt_id`,`fachrichtung_id`)"
 Str(1, 6, 4) = "  KEY `fk_arzt_has_fachrichtung_arzt1` (`arzt_id`)"
 Str(1, 6, 5) = "  KEY `fk_arzt_has_fachrichtung_fachrichtung1` (`fachrichtung_id`)"
 Str(1, 6, 6) = "  CONSTRAINT `fk_arzt_has_fachrichtung_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 6, 7) = "  CONSTRAINT `fk_arzt_has_fachrichtung_fachrichtung1` FOREIGN KEY (`fachrichtung_id`) REFERENCES `fachrichtung` (`idFachrichtung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 6, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "arzt_has_fremdsprache"
 Str(0, 7, 1) = "`arzt_id`"
 Str(0, 7, 2) = "`fremdsprache_id`"
 Str(0, 7, 3) = "`arzt_id`"
 Str(0, 7, 4) = "`fk_arzt_has_fremdsprache_arzt1`"
 Str(0, 7, 5) = "`fk_arzt_has_fremdsprache_fremdsprache1`"
 Str(0, 7, 6) = "`fk_arzt_has_fremdsprache_arzt1`"
 Str(0, 7, 7) = "`fk_arzt_has_fremdsprache_fremdsprache1`"
 ArtZ(0, 7) = 2
 ArtZ(1, 7) = 3
 ArtZ(2, 7) = 2
 Str(1, 7, 0) = "CREATE TABLE `arzt_has_fremdsprache` ("
 Str(1, 7, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 7, 2) = " `fremdsprache_id` int(11) NOT NULL"
 Str(1, 7, 3) = "  PRIMARY KEY (`arzt_id`,`fremdsprache_id`)"
 Str(1, 7, 4) = "  KEY `fk_arzt_has_fremdsprache_arzt1` (`arzt_id`)"
 Str(1, 7, 5) = "  KEY `fk_arzt_has_fremdsprache_fremdsprache1` (`fremdsprache_id`)"
 Str(1, 7, 6) = "  CONSTRAINT `fk_arzt_has_fremdsprache_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 7, 7) = "  CONSTRAINT `fk_arzt_has_fremdsprache_fremdsprache1` FOREIGN KEY (`fremdsprache_id`) REFERENCES `fremdsprache` (`idfremdsprache`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 7, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "arzt_has_genehmigung"
 Str(0, 8, 1) = "`arzt_id`"
 Str(0, 8, 2) = "`genehmigung_id`"
 Str(0, 8, 3) = "`arzt_id`"
 Str(0, 8, 4) = "`fk_hausarzt_has_genehmigung_hausarzt1`"
 Str(0, 8, 5) = "`fk_hausarzt_has_genehmigung_genehmigung1`"
 Str(0, 8, 6) = "`fk_hausarzt_has_genehmigung_genehmigung1`"
 Str(0, 8, 7) = "`fk_hausarzt_has_genehmigung_hausarzt1`"
 ArtZ(0, 8) = 2
 ArtZ(1, 8) = 3
 ArtZ(2, 8) = 2
 Str(1, 8, 0) = "CREATE TABLE `arzt_has_genehmigung` ("
 Str(1, 8, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 8, 2) = " `genehmigung_id` int(11) NOT NULL"
 Str(1, 8, 3) = "  PRIMARY KEY (`arzt_id`,`genehmigung_id`)"
 Str(1, 8, 4) = "  KEY `fk_hausarzt_has_genehmigung_hausarzt1` (`arzt_id`)"
 Str(1, 8, 5) = "  KEY `fk_hausarzt_has_genehmigung_genehmigung1` (`genehmigung_id`)"
 Str(1, 8, 6) = "  CONSTRAINT `fk_hausarzt_has_genehmigung_genehmigung1` FOREIGN KEY (`genehmigung_id`) REFERENCES `genehmigung` (`idgenehmigung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 8, 7) = "  CONSTRAINT `fk_hausarzt_has_genehmigung_hausarzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 8, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "arzt_has_leistungsumfang"
 Str(0, 9, 1) = "`arzt_id`"
 Str(0, 9, 2) = "`leistungsumfang_id`"
 Str(0, 9, 3) = "`arzt_id`"
 Str(0, 9, 4) = "`fk_arzt_has_leistungsumfang_arzt1`"
 Str(0, 9, 5) = "`fk_arzt_has_leistungsumfang_leistungsumfang1`"
 Str(0, 9, 6) = "`fk_arzt_has_leistungsumfang_arzt1`"
 Str(0, 9, 7) = "`fk_arzt_has_leistungsumfang_leistungsumfang1`"
 ArtZ(0, 9) = 2
 ArtZ(1, 9) = 3
 ArtZ(2, 9) = 2
 Str(1, 9, 0) = "CREATE TABLE `arzt_has_leistungsumfang` ("
 Str(1, 9, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 9, 2) = " `leistungsumfang_id` int(11) NOT NULL"
 Str(1, 9, 3) = "  PRIMARY KEY (`arzt_id`,`leistungsumfang_id`)"
 Str(1, 9, 4) = "  KEY `fk_arzt_has_leistungsumfang_arzt1` (`arzt_id`)"
 Str(1, 9, 5) = "  KEY `fk_arzt_has_leistungsumfang_leistungsumfang1` (`leistungsumfang_id`)"
 Str(1, 9, 6) = "  CONSTRAINT `fk_arzt_has_leistungsumfang_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 9, 7) = "  CONSTRAINT `fk_arzt_has_leistungsumfang_leistungsumfang1` FOREIGN KEY (`leistungsumfang_id`) REFERENCES `leistungsumfang` (`idLeistungsumfang`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 9, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "arzt_has_vertragsangebot"
 Str(0, 10, 1) = "`arzt_id`"
 Str(0, 10, 2) = "`vertragsangebot_id`"
 Str(0, 10, 3) = "`arzt_id`"
 Str(0, 10, 4) = "`fk_arzt_has_vertragsangebot_arzt1`"
 Str(0, 10, 5) = "`fk_arzt_has_vertragsangebot_vertragsangebot1`"
 Str(0, 10, 6) = "`fk_arzt_has_vertragsangebot_arzt1`"
 Str(0, 10, 7) = "`fk_arzt_has_vertragsangebot_vertragsangebot1`"
 ArtZ(0, 10) = 2
 ArtZ(1, 10) = 3
 ArtZ(2, 10) = 2
 Str(1, 10, 0) = "CREATE TABLE `arzt_has_vertragsangebot` ("
 Str(1, 10, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 10, 2) = " `vertragsangebot_id` int(11) NOT NULL"
 Str(1, 10, 3) = "  PRIMARY KEY (`arzt_id`,`vertragsangebot_id`)"
 Str(1, 10, 4) = "  KEY `fk_arzt_has_vertragsangebot_arzt1` (`arzt_id`)"
 Str(1, 10, 5) = "  KEY `fk_arzt_has_vertragsangebot_vertragsangebot1` (`vertragsangebot_id`)"
 Str(1, 10, 6) = "  CONSTRAINT `fk_arzt_has_vertragsangebot_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 10, 7) = "  CONSTRAINT `fk_arzt_has_vertragsangebot_vertragsangebot1` FOREIGN KEY (`vertragsangebot_id`) REFERENCES `vertragsangebot` (`idvertragsangebot`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 10, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr10

Sub FüllStr11()
 Str(0, 11, 0) = "arzt_has_weiterbildung"
 Str(0, 11, 1) = "`arzt_id`"
 Str(0, 11, 2) = "`weiterbildung_id`"
 Str(0, 11, 3) = "`arzt_id`"
 Str(0, 11, 4) = "`fk_arzt_has_weiterbildung_arzt1`"
 Str(0, 11, 5) = "`fk_arzt_has_weiterbildung_weiterbildung1`"
 Str(0, 11, 6) = "`fk_arzt_has_weiterbildung_arzt1`"
 Str(0, 11, 7) = "`fk_arzt_has_weiterbildung_weiterbildung1`"
 ArtZ(0, 11) = 2
 ArtZ(1, 11) = 3
 ArtZ(2, 11) = 2
 Str(1, 11, 0) = "CREATE TABLE `arzt_has_weiterbildung` ("
 Str(1, 11, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 11, 2) = " `weiterbildung_id` int(11) NOT NULL"
 Str(1, 11, 3) = "  PRIMARY KEY (`arzt_id`,`weiterbildung_id`)"
 Str(1, 11, 4) = "  KEY `fk_arzt_has_weiterbildung_arzt1` (`arzt_id`)"
 Str(1, 11, 5) = "  KEY `fk_arzt_has_weiterbildung_weiterbildung1` (`weiterbildung_id`)"
 Str(1, 11, 6) = "  CONSTRAINT `fk_arzt_has_weiterbildung_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 11, 7) = "  CONSTRAINT `fk_arzt_has_weiterbildung_weiterbildung1` FOREIGN KEY (`weiterbildung_id`) REFERENCES `weiterbildung` (`idWeiterbildung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 11, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr11

Sub FüllStr12()
 Str(0, 12, 0) = "arzt_has_zusatzbezeichnung"
 Str(0, 12, 1) = "`arzt_id`"
 Str(0, 12, 2) = "`Zusatzbezeichnung_id`"
 Str(0, 12, 3) = "`arzt_id`"
 Str(0, 12, 4) = "`fk_hausarzt_has_Zusatzbezeichnung_hausarzt1`"
 Str(0, 12, 5) = "`fk_hausarzt_has_Zusatzbezeichnung_Zusatzbezeichnung1`"
 Str(0, 12, 6) = "`fk_hausarzt_has_Zusatzbezeichnung_hausarzt1`"
 Str(0, 12, 7) = "`fk_hausarzt_has_Zusatzbezeichnung_Zusatzbezeichnung1`"
 ArtZ(0, 12) = 2
 ArtZ(1, 12) = 3
 ArtZ(2, 12) = 2
 Str(1, 12, 0) = "CREATE TABLE `arzt_has_zusatzbezeichnung` ("
 Str(1, 12, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 12, 2) = " `Zusatzbezeichnung_id` int(11) NOT NULL"
 Str(1, 12, 3) = "  PRIMARY KEY (`arzt_id`,`Zusatzbezeichnung_id`)"
 Str(1, 12, 4) = "  KEY `fk_hausarzt_has_Zusatzbezeichnung_hausarzt1` (`arzt_id`)"
 Str(1, 12, 5) = "  KEY `fk_hausarzt_has_Zusatzbezeichnung_Zusatzbezeichnung1` (`Zusatzbezeichnung_id`)"
 Str(1, 12, 6) = "  CONSTRAINT `fk_hausarzt_has_Zusatzbezeichnung_hausarzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 12, 7) = "  CONSTRAINT `fk_hausarzt_has_Zusatzbezeichnung_Zusatzbezeichnung1` FOREIGN KEY (`Zusatzbezeichnung_id`) REFERENCES `zusatzbezeichnung` (`idZusatzbezeichnung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 12, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr12

Sub FüllStr13()
 Str(0, 13, 0) = "bs"
 Str(0, 13, 1) = "`idbs`"
 Str(0, 13, 2) = "`Straße`"
 Str(0, 13, 3) = "`Hausnr`"
 Str(0, 13, 4) = "`PLZ`"
 Str(0, 13, 5) = "`Ort_id`"
 Str(0, 13, 6) = "`BSNR`"
 Str(0, 13, 7) = "`bsart_id`"
 Str(0, 13, 8) = "`sprechzeiten_id`"
 Str(0, 13, 9) = "`Rollst`"
 Str(0, 13, 10) = "`seit`"
 Str(0, 13, 11) = "`aktzeit`"
 Str(0, 13, 12) = "`Name`"
 Str(0, 13, 13) = "`idbs`"
 Str(0, 13, 14) = "`fk_Betriebsstätte_Ort1`"
 Str(0, 13, 15) = "`fk_betriebsstätte_bsart1`"
 Str(0, 13, 16) = "`fk_bs_sprechzeiten1`"
 Str(0, 13, 17) = "`fk_betriebsstätte_bsart1`"
 Str(0, 13, 18) = "`fk_Betriebsstätte_Ort1`"
 Str(0, 13, 19) = "`fk_bs_sprechzeiten1`"
 ArtZ(0, 13) = 12
 ArtZ(1, 13) = 4
 ArtZ(2, 13) = 3
 Str(1, 13, 0) = "CREATE TABLE `bs` ("
 Str(1, 13, 1) = " `idbs` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 13, 2) = " `Straße` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 13, 3) = " `Hausnr` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 13, 4) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 13, 5) = " `Ort_id` int(11) NOT NULL"
 Str(1, 13, 6) = " `BSNR` int(11) DEFAULT NULL"
 Str(1, 13, 7) = " `bsart_id` int(11) NOT NULL"
 Str(1, 13, 8) = " `sprechzeiten_id` int(11) NOT NULL"
 Str(1, 13, 9) = " `Rollst` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT '""Rollstuhlgerechte Praxis""'"
 Str(1, 13, 10) = " `seit` datetime DEFAULT NULL"
 Str(1, 13, 11) = " `aktzeit` datetime DEFAULT NULL"
 Str(1, 13, 12) = " `Name` varchar(557) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 13, 13) = "  PRIMARY KEY (`idbs`)"
 Str(1, 13, 14) = "  KEY `fk_Betriebsstätte_Ort1` (`Ort_id`)"
 Str(1, 13, 15) = "  KEY `fk_betriebsstätte_bsart1` (`bsart_id`)"
 Str(1, 13, 16) = "  KEY `fk_bs_sprechzeiten1` (`sprechzeiten_id`)"
 Str(1, 13, 17) = "  CONSTRAINT `fk_betriebsstätte_bsart1` FOREIGN KEY (`bsart_id`) REFERENCES `bsart` (`idbsart`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 13, 18) = "  CONSTRAINT `fk_Betriebsstätte_Ort1` FOREIGN KEY (`Ort_id`) REFERENCES `ort` (`idOrt`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 13, 19) = "  CONSTRAINT `fk_bs_sprechzeiten1` FOREIGN KEY (`sprechzeiten_id`) REFERENCES `sprechzeiten` (`idsprechzeiten`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 13, 20) = " ENGINE=InnoDB AUTO_INCREMENT=20464 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Betriebsstätte'"
End Sub ' FüllStr13

Sub FüllStr14()
 Str(0, 14, 0) = "bs_has_genehmigung"
 Str(0, 14, 1) = "`bs_id`"
 Str(0, 14, 2) = "`genehmigung_id`"
 Str(0, 14, 3) = "`bs_id`"
 Str(0, 14, 4) = "`fk_hausbs_has_genehmigung_hausbs1`"
 Str(0, 14, 5) = "`fk_hausbs_has_genehmigung_genehmigung1`"
 Str(0, 14, 6) = "`fk_hausbs_has_genehmigung_genehmigung1`"
 Str(0, 14, 7) = "`fk_hausbs_has_genehmigung_hausbs1`"
 ArtZ(0, 14) = 2
 ArtZ(1, 14) = 3
 ArtZ(2, 14) = 2
 Str(1, 14, 0) = "CREATE TABLE `bs_has_genehmigung` ("
 Str(1, 14, 1) = " `bs_id` int(11) NOT NULL"
 Str(1, 14, 2) = " `genehmigung_id` int(11) NOT NULL"
 Str(1, 14, 3) = "  PRIMARY KEY (`bs_id`,`genehmigung_id`)"
 Str(1, 14, 4) = "  KEY `fk_hausbs_has_genehmigung_hausbs1` (`bs_id`)"
 Str(1, 14, 5) = "  KEY `fk_hausbs_has_genehmigung_genehmigung1` (`genehmigung_id`)"
 Str(1, 14, 6) = "  CONSTRAINT `fk_hausbs_has_genehmigung_genehmigung1` FOREIGN KEY (`genehmigung_id`) REFERENCES `genehmigung` (`idgenehmigung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 14, 7) = "  CONSTRAINT `fk_hausbs_has_genehmigung_hausbs1` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 14, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr14

Sub FüllStr15()
 Str(0, 15, 0) = "bs_has_leistungsumfang"
 Str(0, 15, 1) = "`bs_id`"
 Str(0, 15, 2) = "`leistungsumfang_id`"
 Str(0, 15, 3) = "`bs_id`"
 Str(0, 15, 4) = "`fk_bs_has_leistungsumfang_bs1`"
 Str(0, 15, 5) = "`fk_bs_has_leistungsumfang_leistungsumfang1`"
 Str(0, 15, 6) = "`fk_bs_has_leistungsumfang_bs1`"
 Str(0, 15, 7) = "`fk_bs_has_leistungsumfang_leistungsumfang1`"
 ArtZ(0, 15) = 2
 ArtZ(1, 15) = 3
 ArtZ(2, 15) = 2
 Str(1, 15, 0) = "CREATE TABLE `bs_has_leistungsumfang` ("
 Str(1, 15, 1) = " `bs_id` int(11) NOT NULL"
 Str(1, 15, 2) = " `leistungsumfang_id` int(11) NOT NULL"
 Str(1, 15, 3) = "  PRIMARY KEY (`bs_id`,`leistungsumfang_id`)"
 Str(1, 15, 4) = "  KEY `fk_bs_has_leistungsumfang_bs1` (`bs_id`)"
 Str(1, 15, 5) = "  KEY `fk_bs_has_leistungsumfang_leistungsumfang1` (`leistungsumfang_id`)"
 Str(1, 15, 6) = "  CONSTRAINT `fk_bs_has_leistungsumfang_bs1` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 15, 7) = "  CONSTRAINT `fk_bs_has_leistungsumfang_leistungsumfang1` FOREIGN KEY (`leistungsumfang_id`) REFERENCES `leistungsumfang` (`idLeistungsumfang`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 15, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr15

Sub FüllStr16()
 Str(0, 16, 0) = "bs_has_vertragsangebot"
 Str(0, 16, 1) = "`bs_id`"
 Str(0, 16, 2) = "`vertragsangebot_id`"
 Str(0, 16, 3) = "`bs_id`"
 Str(0, 16, 4) = "`fk_bs_has_vertragsangebot_bs1`"
 Str(0, 16, 5) = "`fk_bs_has_vertragsangebot_vertragsangebot1`"
 Str(0, 16, 6) = "`fk_bs_has_vertragsangebot_bs1`"
 Str(0, 16, 7) = "`fk_bs_has_vertragsangebot_vertragsangebot1`"
 ArtZ(0, 16) = 2
 ArtZ(1, 16) = 3
 ArtZ(2, 16) = 2
 Str(1, 16, 0) = "CREATE TABLE `bs_has_vertragsangebot` ("
 Str(1, 16, 1) = " `bs_id` int(11) NOT NULL"
 Str(1, 16, 2) = " `vertragsangebot_id` int(11) NOT NULL"
 Str(1, 16, 3) = "  PRIMARY KEY (`bs_id`,`vertragsangebot_id`)"
 Str(1, 16, 4) = "  KEY `fk_bs_has_vertragsangebot_bs1` (`bs_id`)"
 Str(1, 16, 5) = "  KEY `fk_bs_has_vertragsangebot_vertragsangebot1` (`vertragsangebot_id`)"
 Str(1, 16, 6) = "  CONSTRAINT `fk_bs_has_vertragsangebot_bs1` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 16, 7) = "  CONSTRAINT `fk_bs_has_vertragsangebot_vertragsangebot1` FOREIGN KEY (`vertragsangebot_id`) REFERENCES `vertragsangebot` (`idvertragsangebot`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 16, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr16

Sub FüllStr17()
 Str(0, 17, 0) = "bsart"
 Str(0, 17, 1) = "`idbsart`"
 Str(0, 17, 2) = "`BSArt`"
 Str(0, 17, 3) = "`idbsart`"
 Str(0, 17, 4) = "`bsart`"
 ArtZ(0, 17) = 2
 ArtZ(1, 17) = 2
 Str(1, 17, 0) = "CREATE TABLE `bsart` ("
 Str(1, 17, 1) = " `idbsart` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 17, 2) = " `BSArt` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 3) = "  PRIMARY KEY (`idbsart`)"
 Str(1, 17, 4) = "  UNIQUE KEY `bsart` (`BSArt`)"
 Str(1, 17, 5) = " ENGINE=InnoDB AUTO_INCREMENT=15 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr17

Sub FüllStr18()
 Str(0, 18, 0) = "ermächtigteEinrichtungen"
 Str(1, 18, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `ermächtigteEinrichtungen` AS SELECT `bs`.`Name` AS `Name`,`bs`.`Straße` AS `Straße`,`bs`.`Hausnr` AS `Hausnr`,`bs`.`PLZ` AS `Plz`,`ort`.`Ort` AS `Ort`,`bs`.`BSNR` AS `Bsnr`,`bsart`.`BSArt` AS `BSArt`,`sprechzeiten`.`Sprechzeiten` AS `Sprechzeiten`,`bs`.`Rollst` AS `Rollst`,(SELECT GROUP_CONCAT(`genehmigung`.`genehmigung` SEPARATOR ',') FROM (`bs_has_genehmigung` `bhg` LEFT JOIN `genehmigung` on((`bhg`.`genehmigung_id` = `genehmigung`.`idgenehmigung`))) WHERE (`bs`.`idbs` = `bhg`.`bs_id`)) AS `Genehmigung`,(SELECT GROUP_CONCAT(`leistungsumfang`.`leistungsumfang` SEPARATOR ',') FROM (`bs_has_leistungsumfang` `bhg` LEFT JOIN `leistungsumfang` on((`bhg`.`leistungsumfang_id` = `leistungsumfang`.`idLeistungsumfang`))) WHERE (`bs`.`idbs` = `bhg`.`bs_id`)) AS `Leistungsumfang`,(SELECT GROUP_CONCAT(`vertr" & _
  "agsangebot`.`vertragsangebot` SEPARATOR ',') FROM (`bs_has_vertragsangebot` `bhg` LEFT JOIN `vertragsangebot` on((`bhg`.`vertragsangebot_id` = `vertragsangebot`.`idvertragsangebot`))) WHERE (`bs`.`idbs` = `bhg`.`bs_id`)) AS `Vertragsangebot`,`bs`.`aktzeit` AS `Aktzeit` FROM ((((`bs` LEFT JOIN `arzt_has_bs` `ahb` on((`bs`.`idbs` = `ahb`.`bs_id`))) LEFT JOIN `ort` on((`bs`.`Ort_id` = `ort`.`idOrt`))) LEFT JOIN `bsart` on((`bs`.`bsart_id` = `bsart`.`idbsart`))) LEFT JOIN `sprechzeiten` on((`bs`.`sprechzeiten_id` = `sprechzeiten`.`idsprechzeiten`))) WHERE ISNULL(`ahb`.`bs_id`)"
End Sub ' FüllStr18

Sub FüllStr19()
 Str(0, 19, 0) = "fachrichtung"
 Str(0, 19, 1) = "`idFachrichtung`"
 Str(0, 19, 2) = "`Fachrichtung`"
 Str(0, 19, 3) = "`idFachrichtung`"
 Str(0, 19, 4) = "`fachrichtung`"
 ArtZ(0, 19) = 2
 ArtZ(1, 19) = 2
 Str(1, 19, 0) = "CREATE TABLE `fachrichtung` ("
 Str(1, 19, 1) = " `idFachrichtung` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 19, 2) = " `Fachrichtung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 3) = "  PRIMARY KEY (`idFachrichtung`)"
 Str(1, 19, 4) = "  UNIQUE KEY `fachrichtung` (`Fachrichtung`)"
 Str(1, 19, 5) = " ENGINE=InnoDB AUTO_INCREMENT=195 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr19

Sub FüllStr20()
 Str(0, 20, 0) = "fax"
 Str(0, 20, 1) = "`idfax`"
 Str(0, 20, 2) = "`Fax`"
 Str(0, 20, 3) = "`bs_id`"
 Str(0, 20, 4) = "`FaxZahl`"
 Str(0, 20, 5) = "`idfax`"
 Str(0, 20, 6) = "`fax`"
 Str(0, 20, 7) = "`fk_fax_bs`"
 Str(0, 20, 8) = "`faxzahl`"
 Str(0, 20, 9) = "`fk_fax_bs`"
 ArtZ(0, 20) = 4
 ArtZ(1, 20) = 4
 ArtZ(2, 20) = 1
 Str(1, 20, 0) = "CREATE TABLE `fax` ("
 Str(1, 20, 1) = " `idfax` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 20, 2) = " `Fax` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 3) = " `bs_id` int(11) NOT NULL"
 Str(1, 20, 4) = " `FaxZahl` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 5) = "  PRIMARY KEY (`idfax`)"
 Str(1, 20, 6) = "  KEY `fax` (`Fax`)"
 Str(1, 20, 7) = "  KEY `fk_fax_bs` (`bs_id`)"
 Str(1, 20, 8) = "  KEY `faxzahl` (`FaxZahl`)"
 Str(1, 20, 9) = "  CONSTRAINT `fk_fax_bs` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 20, 10) = " ENGINE=InnoDB AUTO_INCREMENT=347697 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr20

Sub FüllStr21()
 Str(0, 21, 0) = "fremdsprache"
 Str(0, 21, 1) = "`idfremdsprache`"
 Str(0, 21, 2) = "`fremdsprache`"
 Str(0, 21, 3) = "`idfremdsprache`"
 Str(0, 21, 4) = "`fremdsprache`"
 ArtZ(0, 21) = 2
 ArtZ(1, 21) = 2
 Str(1, 21, 0) = "CREATE TABLE `fremdsprache` ("
 Str(1, 21, 1) = " `idfremdsprache` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 21, 2) = " `fremdsprache` varchar(414) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 3) = "  PRIMARY KEY (`idfremdsprache`)"
 Str(1, 21, 4) = "  UNIQUE KEY `fremdsprache` (`fremdsprache`)"
 Str(1, 21, 5) = " ENGINE=InnoDB AUTO_INCREMENT=77 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr21

Sub FüllStr22()
 Str(0, 22, 0) = "genehmigung"
 Str(0, 22, 1) = "`idgenehmigung`"
 Str(0, 22, 2) = "`genehmigung`"
 Str(0, 22, 3) = "`idgenehmigung`"
 Str(0, 22, 4) = "`genehmigung`"
 ArtZ(0, 22) = 2
 ArtZ(1, 22) = 2
 Str(1, 22, 0) = "CREATE TABLE `genehmigung` ("
 Str(1, 22, 1) = " `idgenehmigung` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 22, 2) = " `genehmigung` varchar(335) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 22, 3) = "  PRIMARY KEY (`idgenehmigung`)"
 Str(1, 22, 4) = "  UNIQUE KEY `genehmigung` (`genehmigung`)"
 Str(1, 22, 5) = " ENGINE=InnoDB AUTO_INCREMENT=274 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr22

Sub FüllStr23()
 Str(0, 23, 0) = "hae"
 Str(1, 23, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `hae` AS SELECT GROUP_CONCAT(DISTINCT `a`.`Nachname` SEPARATOR ',') AS `haname`,`ort`.`Ort` AS `ort`,CAST(left(`bs`.`BSNR`,7) AS char charset utf8) AS `kvnu`,CONCAT(left(`bs`.`BSNR`,2),'/',substr(`bs`.`BSNR`,3,5)) AS `kvnr`,replace(`tel1`.`Tel`,'-','') AS `tel1`,replace(`tel2`.`Tel`,'-','') AS `tel2`,replace(`fax1`.`Fax`,'-','') AS `fax1`,replace(`fax2`.`Fax`,'-','') AS `fax2`,`mail1`.`Mail` AS `email`,if(`a`.`obweibl`,'Frau','Herr') AS `anrede`,`t`.`Titel` AS `titel`,`a`.`Vorname` AS `vorname`,`a`.`Nachname` AS `nachname`,MAX(if(ISNULL((SELECT `g1`.`genehmigung` AS `genehmigung` FROM (`arzt_has_genehmigung` `ahg1` LEFT JOIN `genehmigung` `g1` on((`g1`.`idgenehmigung` = `ahg1`.`genehmigung_id`))) WHERE ((`ahg1`.`arzt_id` = `a`.`idarzt`) AND (`g1`.`genehmigung` = 'DMP-DM1_Koordinierender Arzt_Hausa" & _
  "rzt')))),'0','1')) AS `dmpt1`,MAX(if(ISNULL((SELECT `g2`.`genehmigung` AS `genehmigung` FROM (`arzt_has_genehmigung` `ahg2` LEFT JOIN `genehmigung` `g2` on((`g2`.`idgenehmigung` = `ahg2`.`genehmigung_id`))) WHERE ((`ahg2`.`arzt_id` = `a`.`idarzt`) AND (`g2`.`genehmigung` = 'DMP-DM2_Koordinierender Arzt')))),'0','1')) AS `dmpt2`,`a`.`LANR` AS `lanr`,CONCAT(`bs`.`Straße`,' ',`bs`.`Hausnr`) AS `straße`,`bs`.`PLZ` AS `plz`,GROUP_CONCAT(DISTINCT `fr`.`Fachrichtung` SEPARATOR ',') AS `zulg`,`nlart`.`Niederlassungsart` AS `Arzttyp`,`bs`.`aktzeit` AS `aktzeit` FROM ((((((((((((`bs` LEFT JOIN `ort` on((`bs`.`Ort_id` = `ort`.`idOrt`))) LEFT JOIN `tel` `tel1` on((`tel1`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `tel` `tel2` on(((`tel2`.`bs_id` = `bs`.`idbs`) AND (`tel2`.`Tel` <> `tel1`.`Tel`)))) LEFT JOIN `fax` `fax1` on((`fax1`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `fax` `fax2` on(((`fax2`.`bs_id` = `bs`.`i" & _
  "dbs`) AND (`fax2`.`Fax` <> `fax1`.`Fax`)))) LEFT JOIN `mail` `mail1` on((`mail1`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `arzt_has_bs` `ahb` on((`ahb`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `arzt` `a` on((`a`.`idarzt` = `ahb`.`arzt_id`))) LEFT JOIN `titel` `t` on((`t`.`idtitel` = `a`.`titel_id`))) LEFT JOIN `arzt_has_fachrichtung` `ahf` on((`ahf`.`arzt_id` = `a`.`idarzt`))) LEFT JOIN `fachrichtung` `fr` on((`fr`.`idFachrichtung` = `ahf`.`fachrichtung_id`))) LEFT JOIN `nlart` on((`nlart`.`idnlart` = `a`.`nlart_id`))) GROUP BY CAST(left(`bs`.`BSNR`,7) AS char charset utf8)"
End Sub ' FüllStr23

Sub FüllStr24()
 Str(0, 24, 0) = "haea"
 Str(1, 24, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `haea` AS SELECT GROUP_CONCAT(DISTINCT `a`.`Nachname` SEPARATOR ',') AS `haname`,`ort`.`Ort` AS `ort`,CAST(left(`bs`.`BSNR`,7) AS char charset utf8) AS `kvnu`,CONCAT(left(`bs`.`BSNR`,2),'/',substr(`bs`.`BSNR`,3,5)) AS `kvnr`,replace(`tel1`.`Tel`,'-','') AS `tel1`,replace(`tel2`.`Tel`,'-','') AS `tel2`,replace(`fax1`.`Fax`,'-','') AS `fax1`,replace(`fax2`.`Fax`,'-','') AS `fax2`,`mail1`.`Mail` AS `email`,if(`a`.`obweibl`,'Frau','Herr') AS `anrede`,`t`.`Titel` AS `titel`,`a`.`Vorname` AS `vorname`,`a`.`Nachname` AS `nachname`,MAX(if(ISNULL((SELECT `g1`.`genehmigung` AS `genehmigung` FROM (`arzt_has_genehmigung` `ahg1` LEFT JOIN `genehmigung` `g1` on((`g1`.`idgenehmigung` = `ahg1`.`genehmigung_id`))) WHERE ((`ahg1`.`arzt_id` = `a`.`idarzt`) AND (`g1`.`genehmigung` = 'DMP-DM1_Koordinierender Arzt_Haus" & _
  "arzt')))),'0','1')) AS `dmpt1`,MAX(if(ISNULL((SELECT `g2`.`genehmigung` AS `genehmigung` FROM (`arzt_has_genehmigung` `ahg2` LEFT JOIN `genehmigung` `g2` on((`g2`.`idgenehmigung` = `ahg2`.`genehmigung_id`))) WHERE ((`ahg2`.`arzt_id` = `a`.`idarzt`) AND (`g2`.`genehmigung` = 'DMP-DM2_Koordinierender Arzt')))),'0','1')) AS `dmpt2`,`a`.`LANR` AS `lanr`,CONCAT(`bs`.`Straße`,' ',`bs`.`Hausnr`) AS `straße`,`bs`.`PLZ` AS `plz`,GROUP_CONCAT(DISTINCT `fr`.`Fachrichtung` SEPARATOR ',') AS `zulg`,`nlart`.`Niederlassungsart` AS `Arzttyp`,`bs`.`aktzeit` AS `aktzeit` FROM ((((((((((((`bs` LEFT JOIN `ort` on((`bs`.`Ort_id` = `ort`.`idOrt`))) LEFT JOIN `tel` `tel1` on((`tel1`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `tel` `tel2` on(((`tel2`.`bs_id` = `bs`.`idbs`) AND (`tel2`.`Tel` <> `tel1`.`Tel`)))) LEFT JOIN `fax` `fax1` on((`fax1`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `fax` `fax2` on(((`fax2`.`bs_id` = `bs`.`" & _
  "idbs`) AND (`fax2`.`Fax` <> `fax1`.`Fax`)))) LEFT JOIN `mail` `mail1` on((`mail1`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `arzt_has_bs` `ahb` on((`ahb`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `arzt` `a` on((`a`.`idarzt` = `ahb`.`arzt_id`))) LEFT JOIN `titel` `t` on((`t`.`idtitel` = `a`.`titel_id`))) LEFT JOIN `arzt_has_fachrichtung` `ahf` on((`ahf`.`arzt_id` = `a`.`idarzt`))) LEFT JOIN `fachrichtung` `fr` on((`fr`.`idFachrichtung` = `ahf`.`fachrichtung_id`))) LEFT JOIN `nlart` on((`nlart`.`idnlart` = `a`.`nlart_id`))) GROUP BY CAST(left(`a`.`LANR`,7) AS char charset utf8),`ort`.`Ort`"
End Sub ' FüllStr24

Sub FüllStr25()
 Str(0, 25, 0) = "haekurz"
 Str(1, 25, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `haekurz` AS SELECT `ort`.`Ort` AS `ort`,`bs`.`BSNR` AS `KVNu`,replace(`tel1`.`Tel`,'-','') AS `tel1`,replace(`fax1`.`Fax`,'-','') AS `fax1`,if(`a`.`obweibl`,'Frau','Herr') AS `anrede`,`t`.`Titel` AS `titel`,`a`.`Vorname` AS `vorname`,`a`.`Nachname` AS `nachname`,CONCAT(`bs`.`Straße`,' ',`bs`.`Hausnr`) AS `straße`,`bs`.`PLZ` AS `plz`,GROUP_CONCAT(DISTINCT `fr`.`Fachrichtung` SEPARATOR ',') AS `zulg`,`bs`.`aktzeit` AS `aktzeit` FROM ((((((((`bs` LEFT JOIN `ort` on((`bs`.`Ort_id` = `ort`.`idOrt`))) LEFT JOIN `tel` `tel1` on((`tel1`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `fax` `fax1` on((`fax1`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `arzt_has_bs` `ahb` on((`ahb`.`bs_id` = `bs`.`idbs`))) LEFT JOIN `arzt` `a` on((`a`.`idarzt` = `ahb`.`arzt_id`))) LEFT JOIN `titel` `t` on((`t`.`idtitel` = `a`.`titel_id`))) left" & _
  " join `arzt_has_fachrichtung` `ahf` on((`ahf`.`arzt_id` = `a`.`idarzt`))) LEFT JOIN `fachrichtung` `fr` on((`fr`.`idFachrichtung` = `ahf`.`fachrichtung_id`))) GROUP BY CAST(left(`bs`.`BSNR`,7) AS char charset utf8)"
End Sub ' FüllStr25

Sub FüllStr26()
 Str(0, 26, 0) = "leistungsumfang"
 Str(0, 26, 1) = "`idLeistungsumfang`"
 Str(0, 26, 2) = "`leistungsumfang`"
 Str(0, 26, 3) = "`idLeistungsumfang`"
 Str(0, 26, 4) = "`Leistungsumfang`"
 ArtZ(0, 26) = 2
 ArtZ(1, 26) = 2
 Str(1, 26, 0) = "CREATE TABLE `leistungsumfang` ("
 Str(1, 26, 1) = " `idLeistungsumfang` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 26, 2) = " `leistungsumfang` varchar(1311) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 26, 3) = "  PRIMARY KEY (`idLeistungsumfang`)"
 Str(1, 26, 4) = "  KEY `Leistungsumfang` (`leistungsumfang`(300)) USING BTREE"
 Str(1, 26, 5) = " ENGINE=InnoDB AUTO_INCREMENT=24286 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr26

Sub FüllStr27()
 Str(0, 27, 0) = "mail"
 Str(0, 27, 1) = "`idMail`"
 Str(0, 27, 2) = "`Mail`"
 Str(0, 27, 3) = "`bs_id`"
 Str(0, 27, 4) = "`idMail`"
 Str(0, 27, 5) = "`mail`"
 Str(0, 27, 6) = "`fk_mail_bs`"
 Str(0, 27, 7) = "`fk_mail_bs`"
 ArtZ(0, 27) = 3
 ArtZ(1, 27) = 3
 ArtZ(2, 27) = 1
 Str(1, 27, 0) = "CREATE TABLE `mail` ("
 Str(1, 27, 1) = " `idMail` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 27, 2) = " `Mail` varchar(72) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 27, 3) = " `bs_id` int(11) NOT NULL"
 Str(1, 27, 4) = "  PRIMARY KEY (`idMail`) USING BTREE"
 Str(1, 27, 5) = "  KEY `mail` (`Mail`)"
 Str(1, 27, 6) = "  KEY `fk_mail_bs` (`bs_id`)"
 Str(1, 27, 7) = "  CONSTRAINT `fk_mail_bs` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 27, 8) = " ENGINE=InnoDB AUTO_INCREMENT=177116 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr27

Sub FüllStr28()
 Str(0, 28, 0) = "nlart"
 Str(0, 28, 1) = "`idnlart`"
 Str(0, 28, 2) = "`Niederlassungsart`"
 Str(0, 28, 3) = "`idnlart`"
 Str(0, 28, 4) = "`nlart`"
 ArtZ(0, 28) = 2
 ArtZ(1, 28) = 2
 Str(1, 28, 0) = "CREATE TABLE `nlart` ("
 Str(1, 28, 1) = " `idnlart` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 28, 2) = " `Niederlassungsart` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 28, 3) = "  PRIMARY KEY (`idnlart`)"
 Str(1, 28, 4) = "  UNIQUE KEY `nlart` (`Niederlassungsart`)"
 Str(1, 28, 5) = " ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Niederlassungsart'"
End Sub ' FüllStr28

Sub FüllStr29()
 Str(0, 29, 0) = "ort"
 Str(0, 29, 1) = "`idOrt`"
 Str(0, 29, 2) = "`Ort`"
 Str(0, 29, 3) = "`idOrt`"
 Str(0, 29, 4) = "`ort`"
 ArtZ(0, 29) = 2
 ArtZ(1, 29) = 2
 Str(1, 29, 0) = "CREATE TABLE `ort` ("
 Str(1, 29, 1) = " `idOrt` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 29, 2) = " `Ort` varchar(46) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 29, 3) = "  PRIMARY KEY (`idOrt`)"
 Str(1, 29, 4) = "  UNIQUE KEY `ort` (`Ort`)"
 Str(1, 29, 5) = " ENGINE=InnoDB AUTO_INCREMENT=1869 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr29

Sub FüllStr30()
 Str(0, 30, 0) = "sprechzeiten"
 Str(0, 30, 1) = "`idsprechzeiten`"
 Str(0, 30, 2) = "`Sprechzeiten`"
 Str(0, 30, 3) = "`idsprechzeiten`"
 ArtZ(0, 30) = 2
 ArtZ(1, 30) = 1
 Str(1, 30, 0) = "CREATE TABLE `sprechzeiten` ("
 Str(1, 30, 1) = " `idsprechzeiten` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 30, 2) = " `Sprechzeiten` varchar(3860) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 30, 3) = "  PRIMARY KEY (`idsprechzeiten`)"
 Str(1, 30, 4) = " ENGINE=InnoDB AUTO_INCREMENT=13389 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr30

Sub FüllStr31()
 Str(0, 31, 0) = "tel"
 Str(0, 31, 1) = "`idTel`"
 Str(0, 31, 2) = "`Tel`"
 Str(0, 31, 3) = "`bs_id`"
 Str(0, 31, 4) = "`idTel`"
 Str(0, 31, 5) = "`tel`"
 Str(0, 31, 6) = "`fk_tel_bs`"
 Str(0, 31, 7) = "`fk_tel_bs`"
 ArtZ(0, 31) = 3
 ArtZ(1, 31) = 3
 ArtZ(2, 31) = 1
 Str(1, 31, 0) = "CREATE TABLE `tel` ("
 Str(1, 31, 1) = " `idTel` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 31, 2) = " `Tel` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 31, 3) = " `bs_id` int(11) NOT NULL"
 Str(1, 31, 4) = "  PRIMARY KEY (`idTel`)"
 Str(1, 31, 5) = "  KEY `tel` (`Tel`)"
 Str(1, 31, 6) = "  KEY `fk_tel_bs` (`bs_id`)"
 Str(1, 31, 7) = "  CONSTRAINT `fk_tel_bs` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 31, 8) = " ENGINE=InnoDB AUTO_INCREMENT=442522 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr31

Sub FüllStr32()
 Str(0, 32, 0) = "titel"
 Str(0, 32, 1) = "`idtitel`"
 Str(0, 32, 2) = "`Titel`"
 Str(0, 32, 3) = "`idtitel`"
 Str(0, 32, 4) = "`titel`"
 ArtZ(0, 32) = 2
 ArtZ(1, 32) = 2
 Str(1, 32, 0) = "CREATE TABLE `titel` ("
 Str(1, 32, 1) = " `idtitel` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 32, 2) = " `Titel` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 3) = "  PRIMARY KEY (`idtitel`)"
 Str(1, 32, 4) = "  UNIQUE KEY `titel` (`Titel`)"
 Str(1, 32, 5) = " ENGINE=InnoDB AUTO_INCREMENT=684 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr32

Sub FüllStr33()
 Str(0, 33, 0) = "vertragsangebot"
 Str(0, 33, 1) = "`idvertragsangebot`"
 Str(0, 33, 2) = "`vertragsangebot`"
 Str(0, 33, 3) = "`idvertragsangebot`"
 Str(0, 33, 4) = "`vertragsangebot`"
 ArtZ(0, 33) = 2
 ArtZ(1, 33) = 2
 Str(1, 33, 0) = "CREATE TABLE `vertragsangebot` ("
 Str(1, 33, 1) = " `idvertragsangebot` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 33, 2) = " `vertragsangebot` varchar(1287) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 3) = "  PRIMARY KEY (`idvertragsangebot`)"
 Str(1, 33, 4) = "  KEY `vertragsangebot` (`vertragsangebot`(30))"
 Str(1, 33, 5) = " ENGINE=InnoDB AUTO_INCREMENT=106 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr33

Sub FüllStr34()
 Str(0, 34, 0) = "weiterbildung"
 Str(0, 34, 1) = "`idWeiterbildung`"
 Str(0, 34, 2) = "`Weiterbildung`"
 Str(0, 34, 3) = "`idWeiterbildung`"
 Str(0, 34, 4) = "`Weiterbildung`"
 ArtZ(0, 34) = 2
 ArtZ(1, 34) = 2
 Str(1, 34, 0) = "CREATE TABLE `weiterbildung` ("
 Str(1, 34, 1) = " `idWeiterbildung` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 34, 2) = " `Weiterbildung` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 3) = "  PRIMARY KEY (`idWeiterbildung`)"
 Str(1, 34, 4) = "  UNIQUE KEY `Weiterbildung` (`Weiterbildung`)"
 Str(1, 34, 5) = " ENGINE=InnoDB AUTO_INCREMENT=30 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr34

Sub FüllStr35()
 Str(0, 35, 0) = "zusatzbezeichnung"
 Str(0, 35, 1) = "`idZusatzbezeichnung`"
 Str(0, 35, 2) = "`Zusatzbezeichnung`"
 Str(0, 35, 3) = "`idZusatzbezeichnung`"
 Str(0, 35, 4) = "`zusatzbezeichnung`"
 ArtZ(0, 35) = 2
 ArtZ(1, 35) = 2
 Str(1, 35, 0) = "CREATE TABLE `zusatzbezeichnung` ("
 Str(1, 35, 1) = " `idZusatzbezeichnung` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 35, 2) = " `Zusatzbezeichnung` varchar(141) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 35, 3) = "  PRIMARY KEY (`idZusatzbezeichnung`)"
 Str(1, 35, 4) = "  UNIQUE KEY `zusatzbezeichnung` (`Zusatzbezeichnung`)"
 Str(1, 35, 5) = " ENGINE=InnoDB AUTO_INCREMENT=59 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr35

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
 Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
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
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'SplitN ' aufSplit

Public Function doMach_haerzte(DBn$, DBCn As ADODB.Connection, Optional Server$, Optional obStumm% = True) ' Datenbankname
 Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$
 Dim i&, p1&, p2&, p3&, CLen&, CLen1&, obLT%
 Dim Index$()
 On Error Resume Next
 hDBn = DBn
 Open App.path & "\MachDB.bas_prot.txt" For Output As #302
 obProt = (Err.Number = 0)
 On Error GoTo fehler
 If LenB(Server) = 0 Then Server = GetServr(DBCn)
 cnzCStr = "PROVIDER=MSDASQL;driver={" & ODBCStr & "};server=linux1;uid=mysql;pwd=97a5o6;"
 Set cnz = Nothing
 cnz.Open cnzCStr
 Call doEx("CREATE DATABASE IF NOT EXISTS `" & DBn & "` CHARACTER SET latin1 COLLATE latin1_german2_ci;", 0)
 Call doEx("GRANT ALL ON `" & DBn & "`.* to 'praxis'@'%' IDENTIFIED BY 'sonne' WITH GRANT OPTION", 0)
 Call doEx("GRANT ALL ON `" & DBn & "`.* to 'praxis'@'localhost' IDENTIFIED BY 'sonne' WITH GRANT OPTION", 0)
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
 Call doEx("SET FOREIGN_KEY_CHECKS = 0", 0)

 Dim j&, ZZ&, Tbl$, sql As New CString
 For i = 0 To 35
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
     FNr = doEx("DROP VIEW `" & Tbl & "`", 0)
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
  End If ' InStrB(Str(1, i, 0), "CREATE TABLE") <> 0 Then
 Next i
 For i = 0 To 35
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
  For i = 0 To 35
   If InStrB(Str(1, i, 0), "DEFINER VIEW") <> 0 Then
    Dim obCr%
    obCr = 0
    Set rsc = Nothing
    rsc.Open "SHOW TABLES FROM `" & DBn & "` WHERE `tables_in_" & DBn & "` = """ & Str(0, i, 0) & """", cnz, adOpenStatic, adLockReadOnly
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
 Call doEx("SET FOREIGN_KEY_CHECKS = 1", 0)
 If obProt Then Close #302
 If Not obStumm Then
  MsgBox "Fertig mit doMach_haerzte(" & DBn & ",DBCn," & Server & ")!"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), "", CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_haerzte/" & AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_haerzte

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
End Function ' AIoZ(Ursp$) AS CString
