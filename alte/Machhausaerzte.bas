Attribute VB_Name = "Module2"
'Bauanleitung für eine Datenbank wie `//LINUX/hausaerzte` vom 9.11.09 01:07:08
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.Connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 22, 18) As New CString, ArtZ&(3, 22)
Dim hDBn$ ' hiesiger Datenbankname


Function FüllStr0()
 Str(0, 0, 0) = "arzt"
 Str(0, 0, 1) = "`idarzt`"
 Str(0, 0, 2) = "`Nachname`"
 Str(0, 0, 3) = "`Vorname`"
 Str(0, 0, 4) = "`titel_id`"
 Str(0, 0, 5) = "`Namenszusatz`"
 Str(0, 0, 6) = "`LANR`"
 Str(0, 0, 7) = "`nlart_id`"
 Str(0, 0, 8) = "`obweibl`"
 Str(0, 0, 9) = "`idarzt`"
 Str(0, 0, 10) = "`lanr`"
 Str(0, 0, 11) = "`fk_arzt_titel1`"
 Str(0, 0, 12) = "`fk_arzt_nlart1`"
 Str(0, 0, 13) = "`fk_arzt_titel1`"
 Str(0, 0, 14) = "`fk_arzt_nlart1`"
 ArtZ(0, 0) = 8
 ArtZ(1, 0) = 4
 ArtZ(2, 0) = 2
 Str(1, 0, 0) = "CREATE TABLE `arzt` ("
 Str(1, 0, 1) = " `idarzt` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 0, 2) = " `Nachname` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 3) = " `Vorname` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 4) = " `titel_id` int(11) NOT NULL"
 Str(1, 0, 5) = " `Namenszusatz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 6) = " `LANR` int(11) DEFAULT NULL"
 Str(1, 0, 7) = " `nlart_id` int(11) NOT NULL"
 Str(1, 0, 8) = " `obweibl` tinyint(1) unsigned NOT NULL DEFAULT '0'"
 Str(1, 0, 9) = "  PRIMARY KEY (`idarzt`)"
 Str(1, 0, 10) = "  UNIQUE KEY `lanr` (`LANR`)"
 Str(1, 0, 11) = "  KEY `fk_arzt_titel1` (`titel_id`)"
 Str(1, 0, 12) = "  KEY `fk_arzt_nlart1` (`nlart_id`)"
 Str(1, 0, 13) = "  CONSTRAINT `fk_arzt_titel1` FOREIGN KEY (`titel_id`) REFERENCES `titel` (`idtitel`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 0, 14) = "  CONSTRAINT `fk_arzt_nlart1` FOREIGN KEY (`nlart_id`) REFERENCES `nlart` (`idnlart`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 0, 15) = " ENGINE=InnoDB AUTO_INCREMENT=24031 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr0

Function FüllStr1()
 Str(0, 1, 0) = "arzt_has_bs"
 Str(0, 1, 1) = "`bs_id`"
 Str(0, 1, 2) = "`arzt_id`"
 Str(0, 1, 3) = "`obneben`"
 Str(0, 1, 4) = "`bs_id`"
 Str(0, 1, 5) = "`fk_Betriebsstätte_has_hausarzt_Betriebsstätte1`"
 Str(0, 1, 6) = "`fk_Betriebsstätte_has_hausarzt_hausarzt1`"
 Str(0, 1, 7) = "`fk_Betriebsstätte_has_hausarzt_Betriebsstätte1`"
 Str(0, 1, 8) = "`fk_Betriebsstätte_has_hausarzt_hausarzt1`"
 ArtZ(0, 1) = 3
 ArtZ(1, 1) = 3
 ArtZ(2, 1) = 2
 Str(1, 1, 0) = "CREATE TABLE `arzt_has_bs` ("
 Str(1, 1, 1) = " `bs_id` int(11) NOT NULL"
 Str(1, 1, 2) = " `arzt_id` int(11) NOT NULL"
 Str(1, 1, 3) = " `obneben` tinyint(1) unsigned NOT NULL DEFAULT '0'"
 Str(1, 1, 4) = "  PRIMARY KEY (`bs_id`,`arzt_id`)"
 Str(1, 1, 5) = "  KEY `fk_Betriebsstätte_has_hausarzt_Betriebsstätte1` (`bs_id`)"
 Str(1, 1, 6) = "  KEY `fk_Betriebsstätte_has_hausarzt_hausarzt1` (`arzt_id`)"
 Str(1, 1, 7) = "  CONSTRAINT `fk_Betriebsstätte_has_hausarzt_Betriebsstätte1` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 1, 8) = "  CONSTRAINT `fk_Betriebsstätte_has_hausarzt_hausarzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 1, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr1

Function FüllStr2()
 Str(0, 2, 0) = "arzt_has_fachrichtung"
 Str(0, 2, 1) = "`arzt_id`"
 Str(0, 2, 2) = "`fachrichtung_id`"
 Str(0, 2, 3) = "`arzt_id`"
 Str(0, 2, 4) = "`fk_arzt_has_fachrichtung_arzt1`"
 Str(0, 2, 5) = "`fk_arzt_has_fachrichtung_fachrichtung1`"
 Str(0, 2, 6) = "`fk_arzt_has_fachrichtung_arzt1`"
 Str(0, 2, 7) = "`fk_arzt_has_fachrichtung_fachrichtung1`"
 ArtZ(0, 2) = 2
 ArtZ(1, 2) = 3
 ArtZ(2, 2) = 2
 Str(1, 2, 0) = "CREATE TABLE `arzt_has_fachrichtung` ("
 Str(1, 2, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 2, 2) = " `fachrichtung_id` int(11) NOT NULL"
 Str(1, 2, 3) = "  PRIMARY KEY (`arzt_id`,`fachrichtung_id`)"
 Str(1, 2, 4) = "  KEY `fk_arzt_has_fachrichtung_arzt1` (`arzt_id`)"
 Str(1, 2, 5) = "  KEY `fk_arzt_has_fachrichtung_fachrichtung1` (`fachrichtung_id`)"
 Str(1, 2, 6) = "  CONSTRAINT `fk_arzt_has_fachrichtung_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 2, 7) = "  CONSTRAINT `fk_arzt_has_fachrichtung_fachrichtung1` FOREIGN KEY (`fachrichtung_id`) REFERENCES `fachrichtung` (`idFachrichtung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 2, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr2

Function FüllStr3()
 Str(0, 3, 0) = "arzt_has_fremdsprache"
 Str(0, 3, 1) = "`arzt_id`"
 Str(0, 3, 2) = "`fremdsprache_id`"
 Str(0, 3, 3) = "`arzt_id`"
 Str(0, 3, 4) = "`fk_arzt_has_fremdsprache_arzt1`"
 Str(0, 3, 5) = "`fk_arzt_has_fremdsprache_fremdsprache1`"
 Str(0, 3, 6) = "`fk_arzt_has_fremdsprache_arzt1`"
 Str(0, 3, 7) = "`fk_arzt_has_fremdsprache_fremdsprache1`"
 ArtZ(0, 3) = 2
 ArtZ(1, 3) = 3
 ArtZ(2, 3) = 2
 Str(1, 3, 0) = "CREATE TABLE `arzt_has_fremdsprache` ("
 Str(1, 3, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 3, 2) = " `fremdsprache_id` int(11) NOT NULL"
 Str(1, 3, 3) = "  PRIMARY KEY (`arzt_id`,`fremdsprache_id`)"
 Str(1, 3, 4) = "  KEY `fk_arzt_has_fremdsprache_arzt1` (`arzt_id`)"
 Str(1, 3, 5) = "  KEY `fk_arzt_has_fremdsprache_fremdsprache1` (`fremdsprache_id`)"
 Str(1, 3, 6) = "  CONSTRAINT `fk_arzt_has_fremdsprache_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 3, 7) = "  CONSTRAINT `fk_arzt_has_fremdsprache_fremdsprache1` FOREIGN KEY (`fremdsprache_id`) REFERENCES `fremdsprache` (`idfremdsprache`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 3, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr3

Function FüllStr4()
 Str(0, 4, 0) = "arzt_has_genehmigung"
 Str(0, 4, 1) = "`arzt_id`"
 Str(0, 4, 2) = "`genehmigung_id`"
 Str(0, 4, 3) = "`arzt_id`"
 Str(0, 4, 4) = "`fk_hausarzt_has_genehmigung_hausarzt1`"
 Str(0, 4, 5) = "`fk_hausarzt_has_genehmigung_genehmigung1`"
 Str(0, 4, 6) = "`fk_hausarzt_has_genehmigung_hausarzt1`"
 Str(0, 4, 7) = "`fk_hausarzt_has_genehmigung_genehmigung1`"
 ArtZ(0, 4) = 2
 ArtZ(1, 4) = 3
 ArtZ(2, 4) = 2
 Str(1, 4, 0) = "CREATE TABLE `arzt_has_genehmigung` ("
 Str(1, 4, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 4, 2) = " `genehmigung_id` int(11) NOT NULL"
 Str(1, 4, 3) = "  PRIMARY KEY (`arzt_id`,`genehmigung_id`)"
 Str(1, 4, 4) = "  KEY `fk_hausarzt_has_genehmigung_hausarzt1` (`arzt_id`)"
 Str(1, 4, 5) = "  KEY `fk_hausarzt_has_genehmigung_genehmigung1` (`genehmigung_id`)"
 Str(1, 4, 6) = "  CONSTRAINT `fk_hausarzt_has_genehmigung_hausarzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 4, 7) = "  CONSTRAINT `fk_hausarzt_has_genehmigung_genehmigung1` FOREIGN KEY (`genehmigung_id`) REFERENCES `genehmigung` (`idgenehmigung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 4, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr4

Function FüllStr5()
 Str(0, 5, 0) = "arzt_has_vertragsangebot"
 Str(0, 5, 1) = "`arzt_id`"
 Str(0, 5, 2) = "`vertragsangebot_id`"
 Str(0, 5, 3) = "`arzt_id`"
 Str(0, 5, 4) = "`fk_arzt_has_vertragsangebot_arzt1`"
 Str(0, 5, 5) = "`fk_arzt_has_vertragsangebot_vertragsangebot1`"
 Str(0, 5, 6) = "`fk_arzt_has_vertragsangebot_arzt1`"
 Str(0, 5, 7) = "`fk_arzt_has_vertragsangebot_vertragsangebot1`"
 ArtZ(0, 5) = 2
 ArtZ(1, 5) = 3
 ArtZ(2, 5) = 2
 Str(1, 5, 0) = "CREATE TABLE `arzt_has_vertragsangebot` ("
 Str(1, 5, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 5, 2) = " `vertragsangebot_id` int(11) NOT NULL"
 Str(1, 5, 3) = "  PRIMARY KEY (`arzt_id`,`vertragsangebot_id`)"
 Str(1, 5, 4) = "  KEY `fk_arzt_has_vertragsangebot_arzt1` (`arzt_id`)"
 Str(1, 5, 5) = "  KEY `fk_arzt_has_vertragsangebot_vertragsangebot1` (`vertragsangebot_id`)"
 Str(1, 5, 6) = "  CONSTRAINT `fk_arzt_has_vertragsangebot_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 5, 7) = "  CONSTRAINT `fk_arzt_has_vertragsangebot_vertragsangebot1` FOREIGN KEY (`vertragsangebot_id`) REFERENCES `vertragsangebot` (`idvertragsangebot`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 5, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr5

Function FüllStr6()
 Str(0, 6, 0) = "arzt_has_weiterbildung"
 Str(0, 6, 1) = "`arzt_id`"
 Str(0, 6, 2) = "`weiterbildung_id`"
 Str(0, 6, 3) = "`arzt_id`"
 Str(0, 6, 4) = "`fk_arzt_has_weiterbildung_arzt1`"
 Str(0, 6, 5) = "`fk_arzt_has_weiterbildung_weiterbildung1`"
 Str(0, 6, 6) = "`fk_arzt_has_weiterbildung_arzt1`"
 Str(0, 6, 7) = "`fk_arzt_has_weiterbildung_weiterbildung1`"
 ArtZ(0, 6) = 2
 ArtZ(1, 6) = 3
 ArtZ(2, 6) = 2
 Str(1, 6, 0) = "CREATE TABLE `arzt_has_weiterbildung` ("
 Str(1, 6, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 6, 2) = " `weiterbildung_id` int(11) NOT NULL"
 Str(1, 6, 3) = "  PRIMARY KEY (`arzt_id`,`weiterbildung_id`)"
 Str(1, 6, 4) = "  KEY `fk_arzt_has_weiterbildung_arzt1` (`arzt_id`)"
 Str(1, 6, 5) = "  KEY `fk_arzt_has_weiterbildung_weiterbildung1` (`weiterbildung_id`)"
 Str(1, 6, 6) = "  CONSTRAINT `fk_arzt_has_weiterbildung_arzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 6, 7) = "  CONSTRAINT `fk_arzt_has_weiterbildung_weiterbildung1` FOREIGN KEY (`weiterbildung_id`) REFERENCES `weiterbildung` (`idWeiterbildung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 6, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr6

Function FüllStr7()
 Str(0, 7, 0) = "arzt_has_zusatzbezeichnung"
 Str(0, 7, 1) = "`arzt_id`"
 Str(0, 7, 2) = "`Zusatzbezeichnung_id`"
 Str(0, 7, 3) = "`arzt_id`"
 Str(0, 7, 4) = "`fk_hausarzt_has_Zusatzbezeichnung_hausarzt1`"
 Str(0, 7, 5) = "`fk_hausarzt_has_Zusatzbezeichnung_Zusatzbezeichnung1`"
 Str(0, 7, 6) = "`fk_hausarzt_has_Zusatzbezeichnung_hausarzt1`"
 Str(0, 7, 7) = "`fk_hausarzt_has_Zusatzbezeichnung_Zusatzbezeichnung1`"
 ArtZ(0, 7) = 2
 ArtZ(1, 7) = 3
 ArtZ(2, 7) = 2
 Str(1, 7, 0) = "CREATE TABLE `arzt_has_zusatzbezeichnung` ("
 Str(1, 7, 1) = " `arzt_id` int(11) NOT NULL"
 Str(1, 7, 2) = " `Zusatzbezeichnung_id` int(11) NOT NULL"
 Str(1, 7, 3) = "  PRIMARY KEY (`arzt_id`,`Zusatzbezeichnung_id`)"
 Str(1, 7, 4) = "  KEY `fk_hausarzt_has_Zusatzbezeichnung_hausarzt1` (`arzt_id`)"
 Str(1, 7, 5) = "  KEY `fk_hausarzt_has_Zusatzbezeichnung_Zusatzbezeichnung1` (`Zusatzbezeichnung_id`)"
 Str(1, 7, 6) = "  CONSTRAINT `fk_hausarzt_has_Zusatzbezeichnung_hausarzt1` FOREIGN KEY (`arzt_id`) REFERENCES `arzt` (`idarzt`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 7, 7) = "  CONSTRAINT `fk_hausarzt_has_Zusatzbezeichnung_Zusatzbezeichnung1` FOREIGN KEY (`Zusatzbezeichnung_id`) REFERENCES `zusatzbezeichnung` (`idZusatzbezeichnung`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 7, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr7

Function FüllStr8()
 Str(0, 8, 0) = "bs"
 Str(0, 8, 1) = "`idbs`"
 Str(0, 8, 2) = "`Straße`"
 Str(0, 8, 3) = "`Hausnr`"
 Str(0, 8, 4) = "`PLZ`"
 Str(0, 8, 5) = "`Ort_id`"
 Str(0, 8, 6) = "`BSNR`"
 Str(0, 8, 7) = "`bsart_id`"
 Str(0, 8, 8) = "`sprechzeiten_id`"
 Str(0, 8, 9) = "`Rollst`"
 Str(0, 8, 10) = "`idbs`"
 Str(0, 8, 11) = "`fk_Betriebsstätte_Ort1`"
 Str(0, 8, 12) = "`fk_betriebsstätte_bsart1`"
 Str(0, 8, 13) = "`fk_bs_sprechzeiten1`"
 Str(0, 8, 14) = "`fk_Betriebsstätte_Ort1`"
 Str(0, 8, 15) = "`fk_betriebsstätte_bsart1`"
 Str(0, 8, 16) = "`fk_bs_sprechzeiten1`"
 ArtZ(0, 8) = 9
 ArtZ(1, 8) = 4
 ArtZ(2, 8) = 3
 Str(1, 8, 0) = "CREATE TABLE `bs` ("
 Str(1, 8, 1) = " `idbs` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 8, 2) = " `Straße` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 3) = " `Hausnr` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 4) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 5) = " `Ort_id` int(11) NOT NULL"
 Str(1, 8, 6) = " `BSNR` int(11) DEFAULT NULL"
 Str(1, 8, 7) = " `bsart_id` int(11) NOT NULL"
 Str(1, 8, 8) = " `sprechzeiten_id` int(11) NOT NULL"
 Str(1, 8, 9) = " `Rollst` tinyint(1) unsigned NOT NULL DEFAULT '0' COMMENT '""Rollstuhlgerechte Praxis""'"
 Str(1, 8, 10) = "  PRIMARY KEY (`idbs`)"
 Str(1, 8, 11) = "  KEY `fk_Betriebsstätte_Ort1` (`Ort_id`)"
 Str(1, 8, 12) = "  KEY `fk_betriebsstätte_bsart1` (`bsart_id`)"
 Str(1, 8, 13) = "  KEY `fk_bs_sprechzeiten1` (`sprechzeiten_id`)"
 Str(1, 8, 14) = "  CONSTRAINT `fk_Betriebsstätte_Ort1` FOREIGN KEY (`Ort_id`) REFERENCES `ort` (`idOrt`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 8, 15) = "  CONSTRAINT `fk_betriebsstätte_bsart1` FOREIGN KEY (`bsart_id`) REFERENCES `bsart` (`idbsart`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 8, 16) = "  CONSTRAINT `fk_bs_sprechzeiten1` FOREIGN KEY (`sprechzeiten_id`) REFERENCES `sprechzeiten` (`idsprechzeiten`) ON DELETE NO ACTION ON UPDATE NO ACTION"
 Str(1, 8, 17) = " ENGINE=InnoDB AUTO_INCREMENT=18858 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Betriebsstätte'"
End Function ' FüllStr8

Function FüllStr9()
 Str(0, 9, 0) = "bsart"
 Str(0, 9, 1) = "`idbsart`"
 Str(0, 9, 2) = "`BSArt`"
 Str(0, 9, 3) = "`idbsart`"
 Str(0, 9, 4) = "`bsart`"
 ArtZ(0, 9) = 2
 ArtZ(1, 9) = 2
 Str(1, 9, 0) = "CREATE TABLE `bsart` ("
 Str(1, 9, 1) = " `idbsart` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 9, 2) = " `BSArt` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Betriebsstättenart'"
 Str(1, 9, 3) = "  PRIMARY KEY (`idbsart`)"
 Str(1, 9, 4) = "  UNIQUE KEY `bsart` (`BSArt`)"
 Str(1, 9, 5) = " ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr9

Function FüllStr10()
 Str(0, 10, 0) = "fachrichtung"
 Str(0, 10, 1) = "`idFachrichtung`"
 Str(0, 10, 2) = "`Fachrichtung`"
 Str(0, 10, 3) = "`idFachrichtung`"
 Str(0, 10, 4) = "`fachrichtung`"
 ArtZ(0, 10) = 2
 ArtZ(1, 10) = 2
 Str(1, 10, 0) = "CREATE TABLE `fachrichtung` ("
 Str(1, 10, 1) = " `idFachrichtung` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 10, 2) = " `Fachrichtung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 3) = "  PRIMARY KEY (`idFachrichtung`)"
 Str(1, 10, 4) = "  UNIQUE KEY `fachrichtung` (`Fachrichtung`)"
 Str(1, 10, 5) = " ENGINE=InnoDB AUTO_INCREMENT=124 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr10

Function FüllStr11()
 Str(0, 11, 0) = "fax"
 Str(0, 11, 1) = "`idfax`"
 Str(0, 11, 2) = "`Fax`"
 Str(0, 11, 3) = "`bs_id`"
 Str(0, 11, 4) = "`FaxZahl`"
 Str(0, 11, 5) = "`idfax`"
 Str(0, 11, 6) = "`fax`"
 Str(0, 11, 7) = "`fk_fax_bs`"
 Str(0, 11, 8) = "`faxzahl`"
 Str(0, 11, 9) = "`fk_fax_bs`"
 ArtZ(0, 11) = 4
 ArtZ(1, 11) = 4
 ArtZ(2, 11) = 1
 Str(1, 11, 0) = "CREATE TABLE `fax` ("
 Str(1, 11, 1) = " `idfax` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 11, 2) = " `Fax` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 11, 3) = " `bs_id` int(11) NOT NULL"
 Str(1, 11, 4) = " `FaxZahl` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 11, 5) = "  PRIMARY KEY (`idfax`)"
 Str(1, 11, 6) = "  KEY `fax` (`Fax`)"
 Str(1, 11, 7) = "  KEY `fk_fax_bs` (`bs_id`)"
 Str(1, 11, 8) = "  KEY `faxzahl` (`FaxZahl`)"
 Str(1, 11, 9) = "  CONSTRAINT `fk_fax_bs` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 11, 10) = " ENGINE=InnoDB AUTO_INCREMENT=43677 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr11

Function FüllStr12()
 Str(0, 12, 0) = "fremdsprache"
 Str(0, 12, 1) = "`idfremdsprache`"
 Str(0, 12, 2) = "`fremdsprache`"
 Str(0, 12, 3) = "`idfremdsprache`"
 Str(0, 12, 4) = "`fremdsprache`"
 ArtZ(0, 12) = 2
 ArtZ(1, 12) = 2
 Str(1, 12, 0) = "CREATE TABLE `fremdsprache` ("
 Str(1, 12, 1) = " `idfremdsprache` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 12, 2) = " `fremdsprache` varchar(67) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 12, 3) = "  PRIMARY KEY (`idfremdsprache`)"
 Str(1, 12, 4) = "  UNIQUE KEY `fremdsprache` (`fremdsprache`)"
 Str(1, 12, 5) = " ENGINE=InnoDB AUTO_INCREMENT=60 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr12

Function FüllStr13()
 Str(0, 13, 0) = "genehmigung"
 Str(0, 13, 1) = "`idgenehmigung`"
 Str(0, 13, 2) = "`genehmigung`"
 Str(0, 13, 3) = "`idgenehmigung`"
 Str(0, 13, 4) = "`genehmigung`"
 ArtZ(0, 13) = 2
 ArtZ(1, 13) = 2
 Str(1, 13, 0) = "CREATE TABLE `genehmigung` ("
 Str(1, 13, 1) = " `idgenehmigung` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 13, 2) = " `genehmigung` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 13, 3) = "  PRIMARY KEY (`idgenehmigung`)"
 Str(1, 13, 4) = "  UNIQUE KEY `genehmigung` (`genehmigung`)"
 Str(1, 13, 5) = " ENGINE=InnoDB AUTO_INCREMENT=135 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr13

Function FüllStr14()
 Str(0, 14, 0) = "mail"
 Str(0, 14, 1) = "`idfax`"
 Str(0, 14, 2) = "`Mail`"
 Str(0, 14, 3) = "`bs_id`"
 Str(0, 14, 4) = "`idfax`"
 Str(0, 14, 5) = "`mail`"
 Str(0, 14, 6) = "`fk_mail_bs`"
 Str(0, 14, 7) = "`fk_mail_bs`"
 ArtZ(0, 14) = 3
 ArtZ(1, 14) = 3
 ArtZ(2, 14) = 1
 Str(1, 14, 0) = "CREATE TABLE `mail` ("
 Str(1, 14, 1) = " `idfax` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 14, 2) = " `Mail` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 3) = " `bs_id` int(11) NOT NULL"
 Str(1, 14, 4) = "  PRIMARY KEY (`idfax`)"
 Str(1, 14, 5) = "  KEY `mail` (`Mail`)"
 Str(1, 14, 6) = "  KEY `fk_mail_bs` (`bs_id`)"
 Str(1, 14, 7) = "  CONSTRAINT `fk_mail_bs` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 14, 8) = " ENGINE=InnoDB AUTO_INCREMENT=23257 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr14

Function FüllStr15()
 Str(0, 15, 0) = "nlart"
 Str(0, 15, 1) = "`idnlart`"
 Str(0, 15, 2) = "`Niederlassungsart`"
 Str(0, 15, 3) = "`idnlart`"
 Str(0, 15, 4) = "`nlart`"
 ArtZ(0, 15) = 2
 ArtZ(1, 15) = 2
 Str(1, 15, 0) = "CREATE TABLE `nlart` ("
 Str(1, 15, 1) = " `idnlart` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 15, 2) = " `Niederlassungsart` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 15, 3) = "  PRIMARY KEY (`idnlart`)"
 Str(1, 15, 4) = "  UNIQUE KEY `nlart` (`Niederlassungsart`)"
 Str(1, 15, 5) = " ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Niederlassungsart'"
End Function ' FüllStr15

Function FüllStr16()
 Str(0, 16, 0) = "ort"
 Str(0, 16, 1) = "`idOrt`"
 Str(0, 16, 2) = "`Ort`"
 Str(0, 16, 3) = "`idOrt`"
 Str(0, 16, 4) = "`ort`"
 ArtZ(0, 16) = 2
 ArtZ(1, 16) = 2
 Str(1, 16, 0) = "CREATE TABLE `ort` ("
 Str(1, 16, 1) = " `idOrt` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 16, 2) = " `Ort` varchar(46) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 16, 3) = "  PRIMARY KEY (`idOrt`)"
 Str(1, 16, 4) = "  UNIQUE KEY `ort` (`Ort`)"
 Str(1, 16, 5) = " ENGINE=InnoDB AUTO_INCREMENT=1834 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr16

Function FüllStr17()
 Str(0, 17, 0) = "sprechzeiten"
 Str(0, 17, 1) = "`idsprechzeiten`"
 Str(0, 17, 2) = "`Sprechzeiten`"
 Str(0, 17, 3) = "`idsprechzeiten`"
 ArtZ(0, 17) = 2
 ArtZ(1, 17) = 1
 Str(1, 17, 0) = "CREATE TABLE `sprechzeiten` ("
 Str(1, 17, 1) = " `idsprechzeiten` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 17, 2) = " `Sprechzeiten` varchar(1658) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 3) = "  PRIMARY KEY (`idsprechzeiten`)"
 Str(1, 17, 4) = " ENGINE=InnoDB AUTO_INCREMENT=11743 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr17

Function FüllStr18()
 Str(0, 18, 0) = "tel"
 Str(0, 18, 1) = "`idTel`"
 Str(0, 18, 2) = "`Tel`"
 Str(0, 18, 3) = "`bs_id`"
 Str(0, 18, 4) = "`idTel`"
 Str(0, 18, 5) = "`tel`"
 Str(0, 18, 6) = "`fk_tel_bs`"
 Str(0, 18, 7) = "`fk_tel_bs`"
 ArtZ(0, 18) = 3
 ArtZ(1, 18) = 3
 ArtZ(2, 18) = 1
 Str(1, 18, 0) = "CREATE TABLE `tel` ("
 Str(1, 18, 1) = " `idTel` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 18, 2) = " `Tel` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 3) = " `bs_id` int(11) NOT NULL"
 Str(1, 18, 4) = "  PRIMARY KEY (`idTel`)"
 Str(1, 18, 5) = "  KEY `tel` (`Tel`)"
 Str(1, 18, 6) = "  KEY `fk_tel_bs` (`bs_id`)"
 Str(1, 18, 7) = "  CONSTRAINT `fk_tel_bs` FOREIGN KEY (`bs_id`) REFERENCES `bs` (`idbs`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 18, 8) = " ENGINE=InnoDB AUTO_INCREMENT=51979 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr18

Function FüllStr19()
 Str(0, 19, 0) = "titel"
 Str(0, 19, 1) = "`idtitel`"
 Str(0, 19, 2) = "`Titel`"
 Str(0, 19, 3) = "`idtitel`"
 Str(0, 19, 4) = "`titel`"
 ArtZ(0, 19) = 2
 ArtZ(1, 19) = 2
 Str(1, 19, 0) = "CREATE TABLE `titel` ("
 Str(1, 19, 1) = " `idtitel` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 19, 2) = " `Titel` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 3) = "  PRIMARY KEY (`idtitel`)"
 Str(1, 19, 4) = "  UNIQUE KEY `titel` (`Titel`)"
 Str(1, 19, 5) = " ENGINE=InnoDB AUTO_INCREMENT=704 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr19

Function FüllStr20()
 Str(0, 20, 0) = "vertragsangebot"
 Str(0, 20, 1) = "`idvertragsangebot`"
 Str(0, 20, 2) = "`vertragsangebot`"
 Str(0, 20, 3) = "`idvertragsangebot`"
 Str(0, 20, 4) = "`vertragsangebot`"
 ArtZ(0, 20) = 2
 ArtZ(1, 20) = 2
 Str(1, 20, 0) = "CREATE TABLE `vertragsangebot` ("
 Str(1, 20, 1) = " `idvertragsangebot` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 20, 2) = " `vertragsangebot` varchar(71) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 3) = "  PRIMARY KEY (`idvertragsangebot`)"
 Str(1, 20, 4) = "  UNIQUE KEY `vertragsangebot` (`vertragsangebot`)"
 Str(1, 20, 5) = " ENGINE=InnoDB AUTO_INCREMENT=42 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr20

Function FüllStr21()
 Str(0, 21, 0) = "weiterbildung"
 Str(0, 21, 1) = "`idWeiterbildung`"
 Str(0, 21, 2) = "`Weiterbildung`"
 Str(0, 21, 3) = "`idWeiterbildung`"
 Str(0, 21, 4) = "`Weiterbildung`"
 ArtZ(0, 21) = 2
 ArtZ(1, 21) = 2
 Str(1, 21, 0) = "CREATE TABLE `weiterbildung` ("
 Str(1, 21, 1) = " `idWeiterbildung` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 21, 2) = " `Weiterbildung` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 3) = "  PRIMARY KEY (`idWeiterbildung`)"
 Str(1, 21, 4) = "  UNIQUE KEY `Weiterbildung` (`Weiterbildung`)"
 Str(1, 21, 5) = " ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr21

Function FüllStr22()
 Str(0, 22, 0) = "zusatzbezeichnung"
 Str(0, 22, 1) = "`idZusatzbezeichnung`"
 Str(0, 22, 2) = "`Zusatzbezeichnung`"
 Str(0, 22, 3) = "`idZusatzbezeichnung`"
 Str(0, 22, 4) = "`zusatzbezeichnung`"
 ArtZ(0, 22) = 2
 ArtZ(1, 22) = 2
 Str(1, 22, 0) = "CREATE TABLE `zusatzbezeichnung` ("
 Str(1, 22, 1) = " `idZusatzbezeichnung` int(11) NOT NULL AUTO_INCREMENT"
 Str(1, 22, 2) = " `Zusatzbezeichnung` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 22, 3) = "  PRIMARY KEY (`idZusatzbezeichnung`)"
 Str(1, 22, 4) = "  UNIQUE KEY `zusatzbezeichnung` (`Zusatzbezeichnung`)"
 Str(1, 22, 5) = " ENGINE=InnoDB AUTO_INCREMENT=55 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr22

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
 AnwPfad = CurrentDb.Name
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SplitN/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' aufSplit

Public Function doMach_hausaerzte(DBn$, Optional Server$, Optional obStumm% = True) ' Datenbankname
 Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$
 Dim i&, p1&, p2&, p3&, CLen&, CLen1&, obLT%
 Dim Index$()
 On Error Resume Next
 hDBn = DBn
 Open App.path & "\MachDB.bas_prot.txt" For Output As #302
 obProt = (Err.Number = 0)
 On Error GoTo fehler
 If LenB(Server) = 0 Then Server = GetServer(DBCn)
 cnzCStr = "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=" & Server & ";uid=mysql;pwd=97a5o6;"
 Set cnz = Nothing
 cnz.Open cnzCStr
 Call doEx("create database if not exists `" & DBn & "` character set latin1 collate latin1_german2_ci;", 0)
 Call doEx("grant all privileges on `" & DBn & "`.* to 'praxis'@'%' identified by 'sonne' with grant option", 0)
 Call doEx("grant all privileges on `" & DBn & "`.* to 'praxis'@'localhost' identified by 'sonne' with grant option", 0)
 Call doEx("use `" & DBn & "`", 0)
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
 Call doEx("SET FOREIGN_KEY_CHECKS = 0", 0)

 Dim j&, ZZ&, Tbl$, sql As New CString
 For i = 0 To 22
  If InStr(Str(1, i, 0), "CREATE TABLE") <> 0 Then
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
    rsc.Open "show create table `" & Tbl & "`", cnz, adOpenStatic, adLockReadOnly
    sct = rsc.Fields(1)
    If InStrB(sct, "CREATE ALGORITHM") = 1 Then
     FNr = doEx("drop view `" & Tbl & "`", 0)
     FNr = doEx(sql.Value, 0)
    Else
     Exit Do
    End If
   Loop
   If InStrB(AIoZ(sct), AIoZ(Str(1, i, ZZ))) = 0 Then
    Call doEx("alter table `" & Tbl & "`" & Str(1, i, ZZ), 0)
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
      Posi = " FIRST,"
     Else
      Posi = " AFTER " & Str(1, i, j - 1) & ","
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
   If TMt.Length <> 0 Then
    TMt.Cut (TMt.Length - 1)
    Call doEx("Alter Table `" & Tbl & "` " & TMt.Value, -1)
   End If
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
 Call doEx("set FOREIGN_KEY_CHECKS = 1", 0)
 If obProt Then Close #302
 If Not obStumm Then
  MsgBox "Fertig mit doMach_hausaerzte(" & DBn & "," & Server & ")!"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_hausaerzte/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_hausaerzte

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
  AIoZ.Cut (p0 + Len(Such) - 1)
  AIoZ.Append Mid(Ursp, p1)
 End If
End Function ' AIoZ(Ursp$) As CString
