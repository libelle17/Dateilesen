'Bauanleitung für eine Datenbank wie `//linux/office` vom 15.11.09 14:26:29
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 29, 69) As new CString, ArtZ&(3, 29)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "SPPausHae"
 Str(1, 0, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `SPPausHae` AS select distinct concat(`h2`.`Nachname`,_latin1' ',`h2`.`Vorname`) AS `Name`,`h2`.`Email` AS `Email` from (`kvaerzte`.`hae` `h` left join `kvaerzte`.`hae` `h2` on((`h`.`KVNR` = `h2`.`KVNR`))) where (((`h`.`beme` like _latin1'%diabetolog. qualifizierter Arzt%') or (`h`.`beme` like _latin1'%Diabetologie%')) and (not(`h2`.`gelöscht`)) and (`h2`.`Email` <> _latin1'')) order by `h2`.`Nachname`,`h2`.`Vorname`"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "adresse"
 Str(0, 1, 1) = "`Adresse_Nr`"
 Str(0, 1, 2) = "`Name`"
 Str(0, 1, 3) = "`Vorname`"
 Str(0, 1, 4) = "`Titel`"
 Str(0, 1, 5) = "`Anschrift1`"
 Str(0, 1, 6) = "`Anschrift2`"
 Str(0, 1, 7) = "`Postfach`"
 Str(0, 1, 8) = "`Ort`"
 Str(0, 1, 9) = "`Land`"
 Str(0, 1, 10) = "`Telefon_privat`"
 Str(0, 1, 11) = "`Telefon_geschaeftlich`"
 Str(0, 1, 12) = "`Fax`"
 Str(0, 1, 13) = "`Faxk`"
 Str(0, 1, 14) = "`eMail_privat`"
 Str(0, 1, 15) = "`Mobilfunk`"
 Str(0, 1, 16) = "`Reserviert`"
 Str(0, 1, 17) = "`WebSite`"
 Str(0, 1, 18) = "`Briefkopf_Nr`"
 Str(0, 1, 19) = "`Brief_Anrede_Nr`"
 Str(0, 1, 20) = "`Geburtstag`"
 Str(0, 1, 21) = "`Abteilung`"
 Str(0, 1, 22) = "`AdrGru_Nr`"
 Str(0, 1, 23) = "`Ansprechpartner`"
 Str(0, 1, 24) = "`eMail_geschaeftlich`"
 Str(0, 1, 25) = "`Zusatzfeld1`"
 Str(0, 1, 26) = "`Zusatzfeld2`"
 Str(0, 1, 27) = "`Zusatzfeld3`"
 Str(0, 1, 28) = "`Zusatzfeld4`"
 Str(0, 1, 29) = "`Zusatzfeld5`"
 Str(0, 1, 30) = "`Zusatzfeld6`"
 Str(0, 1, 31) = "`Zusatzfeld7`"
 Str(0, 1, 32) = "`Zusatzfeld8`"
 Str(0, 1, 33) = "`Zusatzfeld9`"
 Str(0, 1, 34) = "`Zusatzfeld10`"
 Str(0, 1, 35) = "`Standardsatz`"
 Str(0, 1, 36) = "`Bemerkung`"
 Str(0, 1, 37) = "`UserID`"
 Str(0, 1, 38) = "`ExchangeStatus`"
 Str(0, 1, 39) = "`Abbildung`"
 Str(0, 1, 40) = "`ZIP`"
 Str(0, 1, 41) = "`Branche`"
 Str(0, 1, 42) = "`unsichtbar`"
 Str(0, 1, 43) = "`Geb_unsichtbar`"
 Str(0, 1, 44) = "`geaendert`"
 Str(0, 1, 45) = "`obFest`"
 Str(0, 1, 46) = "`EtikAnr`"
 Str(0, 1, 47) = "`ehemalig`"
 Str(0, 1, 48) = "`Handy`"
 Str(0, 1, 49) = "`Ausdruck`"
 Str(0, 1, 50) = "`Geschlecht`"
 Str(0, 1, 51) = "`KNr`"
 Str(0, 1, 52) = "`Adresse_Nr`"
 Str(0, 1, 53) = "`Identität`"
 Str(0, 1, 54) = "`AdresseBriefAnrede`"
 Str(0, 1, 55) = "`AdresseBriefAnrede_MySQLRel`"
 Str(0, 1, 56) = "`AdresseBriefkopf`"
 Str(0, 1, 57) = "`AdresseBriefkopf_MySQLRel`"
 Str(0, 1, 58) = "`AdrGruAdresse`"
 Str(0, 1, 59) = "`AdrGruAdresse_MySQLRel`"
 Str(0, 1, 60) = "`Faxk`"
 Str(0, 1, 61) = "`geaendert`"
 Str(0, 1, 62) = "`UserID`"
 Str(0, 1, 63) = "`Index_12`"
 Str(0, 1, 64) = "`Fax`"
 Str(0, 1, 65) = "`AdresseBriefAnrede_MySQLRel_AccRel`"
 Str(0, 1, 66) = "`AdresseBriefkopf_MySQLRel_AccRel`"
 Str(0, 1, 67) = "`AdrGruAdresse_MySQLRel_AccRel`"
 ArtZ(0, 1) = 51
 ArtZ(1, 1) = 13
 ArtZ(2, 1) = 3
 Str(1, 1, 0) = "CREATE TABLE `adresse` ("
 Str(1, 1, 1) = " `Adresse_Nr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 1, 2) = " `Name` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 3) = " `Vorname` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 4) = " `Titel` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 5) = " `Anschrift1` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 6) = " `Anschrift2` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 7) = " `Postfach` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 8) = " `Ort` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 9) = " `Land` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 10) = " `Telefon_privat` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 11) = " `Telefon_geschaeftlich` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 12) = " `Fax` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 13) = " `Faxk` varchar(50) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 1, 14) = " `eMail_privat` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 15) = " `Mobilfunk` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 16) = " `Reserviert` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 17) = " `WebSite` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 18) = " `Briefkopf_Nr` int(10) DEFAULT NULL"
 Str(1, 1, 19) = " `Brief_Anrede_Nr` int(10) DEFAULT NULL"
 Str(1, 1, 20) = " `Geburtstag` datetime DEFAULT NULL"
 Str(1, 1, 21) = " `Abteilung` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 22) = " `AdrGru_Nr` int(10) DEFAULT NULL"
 Str(1, 1, 23) = " `Ansprechpartner` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 24) = " `eMail_geschaeftlich` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 25) = " `Zusatzfeld1` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 26) = " `Zusatzfeld2` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 27) = " `Zusatzfeld3` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 28) = " `Zusatzfeld4` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 29) = " `Zusatzfeld5` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 30) = " `Zusatzfeld6` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 31) = " `Zusatzfeld7` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 32) = " `Zusatzfeld8` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 33) = " `Zusatzfeld9` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 34) = " `Zusatzfeld10` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 35) = " `Standardsatz` tinyint(1) DEFAULT NULL"
 Str(1, 1, 36) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 1, 37) = " `UserID` int(10) DEFAULT NULL"
 Str(1, 1, 38) = " `ExchangeStatus` int(10) DEFAULT NULL"
 Str(1, 1, 39) = " `Abbildung` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 40) = " `ZIP` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 41) = " `Branche` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 42) = " `unsichtbar` int(10) DEFAULT NULL"
 Str(1, 1, 43) = " `Geb_unsichtbar` int(10) DEFAULT NULL"
 Str(1, 1, 44) = " `geaendert` datetime DEFAULT NULL"
 Str(1, 1, 45) = " `obFest` int(10) DEFAULT NULL"
 Str(1, 1, 46) = " `EtikAnr` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 47) = " `ehemalig` tinyint(1) DEFAULT NULL COMMENT 'SoleLuna: Ehemaliges Mitglied; sonst: ungültig geworden'"
 Str(1, 1, 48) = " `Handy` tinyint(1) DEFAULT NULL COMMENT 'Übertragung ans Handy'"
 Str(1, 1, 49) = " `Ausdruck` tinyint(1) DEFAULT NULL COMMENT 'Ausdruck auf Adressenliste'"
 Str(1, 1, 50) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'w,m'"
 Str(1, 1, 51) = " `KNr` int(10) unsigned NOT NULL"
 Str(1, 1, 52) = "  UNIQUE KEY `Adresse_Nr` (`Adresse_Nr`)"
 Str(1, 1, 53) = "  UNIQUE KEY `Identität` (`Name`,`Vorname`,`Geburtstag`,`Abteilung`,`Adresse_Nr`)"
 Str(1, 1, 54) = "  KEY `AdresseBriefAnrede` (`Brief_Anrede_Nr`)"
 Str(1, 1, 55) = "  KEY `AdresseBriefAnrede_MySQLRel` (`Brief_Anrede_Nr`)"
 Str(1, 1, 56) = "  KEY `AdresseBriefkopf` (`Briefkopf_Nr`)"
 Str(1, 1, 57) = "  KEY `AdresseBriefkopf_MySQLRel` (`Briefkopf_Nr`)"
 Str(1, 1, 58) = "  KEY `AdrGruAdresse` (`AdrGru_Nr`)"
 Str(1, 1, 59) = "  KEY `AdrGruAdresse_MySQLRel` (`AdrGru_Nr`)"
 Str(1, 1, 60) = "  KEY `Faxk` (`Faxk`)"
 Str(1, 1, 61) = "  KEY `geaendert` (`geaendert`)"
 Str(1, 1, 62) = "  KEY `UserID` (`UserID`)"
 Str(1, 1, 63) = "  KEY `Index_12` (`KNr`,`AdrGru_Nr`)"
 Str(1, 1, 64) = "  KEY `Fax` (`Fax`)"
 Str(1, 1, 65) = "  CONSTRAINT `AdresseBriefAnrede_MySQLRel_AccRel` FOREIGN KEY (`Brief_Anrede_Nr`) REFERENCES `brief_anrede` (`Brief_Anrede_Nr`) ON UPDATE CASCADE"
 Str(1, 1, 66) = "  CONSTRAINT `AdresseBriefkopf_MySQLRel_AccRel` FOREIGN KEY (`Briefkopf_Nr`) REFERENCES `briefkopf` (`Briefkopf_Nr`) ON UPDATE CASCADE"
 Str(1, 1, 67) = "  CONSTRAINT `AdrGruAdresse_MySQLRel_AccRel` FOREIGN KEY (`AdrGru_Nr`) REFERENCES `adrgru` (`AdrGru_Nr`) ON UPDATE CASCADE"
 Str(1, 1, 68) = " ENGINE=InnoDB AUTO_INCREMENT=9403 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=FIXED"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "adrgru"
 Str(0, 2, 1) = "`AdrGru_Nr`"
 Str(0, 2, 2) = "`AdrGru_Name`"
 Str(0, 2, 3) = "`Reihenfolge`"
 Str(0, 2, 4) = "`sichtbar`"
 Str(0, 2, 5) = "`geaendert`"
 Str(0, 2, 6) = "`AdrGru_Nr`"
 Str(0, 2, 7) = "`AdrGru_Nr`"
 Str(0, 2, 8) = "`AdrGru_Name`"
 Str(0, 2, 9) = "`Index_4g`"
 ArtZ(0, 2) = 5
 ArtZ(1, 2) = 4
 Str(1, 2, 0) = "CREATE TABLE `adrgru` ("
 Str(1, 2, 1) = " `AdrGru_Nr` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Adressgruppen-Nr'"
 Str(1, 2, 2) = " `AdrGru_Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Adressgruppen-Name'"
 Str(1, 2, 3) = " `Reihenfolge` int(10) DEFAULT NULL COMMENT 'Reihenfolge 1 = oben'"
 Str(1, 2, 4) = " `sichtbar` tinyint(1) DEFAULT NULL COMMENT 'Adressgruppe sichtbar?'"
 Str(1, 2, 5) = " `geaendert` datetime NOT NULL"
 Str(1, 2, 6) = "  PRIMARY KEY (`AdrGru_Nr`)"
 Str(1, 2, 7) = "  UNIQUE KEY `AdrGru_Nr` (`AdrGru_Nr`)"
 Str(1, 2, 8) = "  UNIQUE KEY `AdrGru_Name` (`AdrGru_Name`)"
 Str(1, 2, 9) = "  KEY `Index_4g` (`geaendert`)"
 Str(1, 2, 10) = " ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=FIXED"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "adrgruzuord"
 Str(0, 3, 1) = "`KNr`"
 Str(0, 3, 2) = "`AdrGru_Nr`"
 Str(0, 3, 3) = "`alt`"
 Str(0, 3, 4) = "`ab`"
 Str(0, 3, 5) = "`bis`"
 Str(0, 3, 6) = "`geändert`"
 Str(0, 3, 7) = "`KNr`"
 Str(0, 3, 8) = "`KNr`"
 Str(0, 3, 9) = "`AdrGru_Nr`"
 Str(0, 3, 10) = "`AdrGruAdrGruZuord`"
 Str(0, 3, 11) = "`AdrGruAdrGruZuord_MySQLRel`"
 Str(0, 3, 12) = "`KontakteAdrGruZuord`"
 Str(0, 3, 13) = "`KontakteAdrGruZuord_MySQLRel`"
 Str(0, 3, 14) = "`AdrGruAdrGruZuord_MySQLRel_AccRel`"
 Str(0, 3, 15) = "`KontakteAdrGruZuord_MySQLRel_AccRel`"
 ArtZ(0, 3) = 6
 ArtZ(1, 3) = 7
 ArtZ(2, 3) = 2
 Str(1, 3, 0) = "CREATE TABLE `adrgruzuord` ("
 Str(1, 3, 1) = " `KNr` int(10) NOT NULL DEFAULT '0' COMMENT 'Kontakt-Nr'"
 Str(1, 3, 2) = " `AdrGru_Nr` int(10) NOT NULL DEFAULT '0' COMMENT 'Adressgruppen-Nr'"
 Str(1, 3, 3) = " `alt` tinyint(1) DEFAULT NULL COMMENT 'nicht mehr gültig?'"
 Str(1, 3, 4) = " `ab` datetime DEFAULT NULL COMMENT 'gültig ab'"
 Str(1, 3, 5) = " `bis` datetime DEFAULT NULL COMMENT 'gültig bis'"
 Str(1, 3, 6) = " `geändert` datetime DEFAULT NULL"
 Str(1, 3, 7) = "  PRIMARY KEY (`KNr`,`AdrGru_Nr`)"
 Str(1, 3, 8) = "  UNIQUE KEY `KNr` (`KNr`,`AdrGru_Nr`)"
 Str(1, 3, 9) = "  KEY `AdrGru_Nr` (`AdrGru_Nr`)"
 Str(1, 3, 10) = "  KEY `AdrGruAdrGruZuord` (`AdrGru_Nr`)"
 Str(1, 3, 11) = "  KEY `AdrGruAdrGruZuord_MySQLRel` (`AdrGru_Nr`)"
 Str(1, 3, 12) = "  KEY `KontakteAdrGruZuord` (`KNr`)"
 Str(1, 3, 13) = "  KEY `KontakteAdrGruZuord_MySQLRel` (`KNr`)"
 Str(1, 3, 14) = "  CONSTRAINT `AdrGruAdrGruZuord_MySQLRel_AccRel` FOREIGN KEY (`AdrGru_Nr`) REFERENCES `adrgru` (`AdrGru_Nr`) ON UPDATE CASCADE"
 Str(1, 3, 15) = "  CONSTRAINT `KontakteAdrGruZuord_MySQLRel_AccRel` FOREIGN KEY (`KNr`) REFERENCES `kontakte` (`KNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 3, 16) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=FIXED"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "anschr"
 Str(0, 4, 1) = "`AnsNr`"
 Str(0, 4, 2) = "`KNr`"
 Str(0, 4, 3) = "`Straße`"
 Str(0, 4, 4) = "`Hausnr`"
 Str(0, 4, 5) = "`PLZ`"
 Str(0, 4, 6) = "`Ort`"
 Str(0, 4, 7) = "`Land`"
 Str(0, 4, 8) = "`Postfach`"
 Str(0, 4, 9) = "`alt`"
 Str(0, 4, 10) = "`ab`"
 Str(0, 4, 11) = "`bis`"
 Str(0, 4, 12) = "`geändert`"
 Str(0, 4, 13) = "`AnsNr`"
 Str(0, 4, 14) = "`AnsNr`"
 Str(0, 4, 15) = "`Ident`"
 Str(0, 4, 16) = "`Knr`"
 Str(0, 4, 17) = "`KontakteAnschrNeu`"
 Str(0, 4, 18) = "`KontakteAnschrNeu_AccRel`"
 ArtZ(0, 4) = 12
 ArtZ(1, 4) = 5
 ArtZ(2, 4) = 1
 Str(1, 4, 0) = "CREATE TABLE `anschr` ("
 Str(1, 4, 1) = " `AnsNr` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Anschrift Nummer'"
 Str(1, 4, 2) = " `KNr` int(10) DEFAULT NULL COMMENT 'Kontakt-Nr'"
 Str(1, 4, 3) = " `Straße` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Straße'"
 Str(1, 4, 4) = " `Hausnr` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Hausnummer'"
 Str(1, 4, 5) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Postleitzahl'"
 Str(1, 4, 6) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ort'"
 Str(1, 4, 7) = " `Land` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Land'"
 Str(1, 4, 8) = " `Postfach` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Postfach'"
 Str(1, 4, 9) = " `alt` tinyint(1) DEFAULT NULL COMMENT 'ob alt'"
 Str(1, 4, 10) = " `ab` datetime DEFAULT NULL COMMENT 'gültig ab'"
 Str(1, 4, 11) = " `bis` datetime DEFAULT NULL COMMENT 'gültig bis'"
 Str(1, 4, 12) = " `geändert` datetime DEFAULT NULL COMMENT 'zuletzt geändert'"
 Str(1, 4, 13) = "  PRIMARY KEY (`AnsNr`)"
 Str(1, 4, 14) = "  UNIQUE KEY `AnsNr` (`AnsNr`)"
 Str(1, 4, 15) = "  KEY `Ident` (`Straße`,`Hausnr`,`PLZ`,`Ort`,`Land`,`Postfach`)"
 Str(1, 4, 16) = "  KEY `Knr` (`KNr`)"
 Str(1, 4, 17) = "  KEY `KontakteAnschrNeu` (`KNr`)"
 Str(1, 4, 18) = "  CONSTRAINT `KontakteAnschrNeu_AccRel` FOREIGN KEY (`KNr`) REFERENCES `kontakte` (`KNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 4, 19) = " ENGINE=InnoDB AUTO_INCREMENT=4806 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Anschrift, über Kontaktnummer einem Kontakt zugeordnet'"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "benutzer"
 Str(0, 5, 1) = "`UserID`"
 Str(0, 5, 2) = "`Name`"
 Str(0, 5, 3) = "`Active`"
 Str(0, 5, 4) = "`Status`"
 Str(0, 5, 5) = "`Beschreibung`"
 Str(0, 5, 6) = "`UserID`"
 Str(0, 5, 7) = "`PrimaryINDEX`"
 Str(0, 5, 8) = "`UserID`"
 ArtZ(0, 5) = 5
 ArtZ(1, 5) = 3
 Str(1, 5, 0) = "CREATE TABLE `benutzer` ("
 Str(1, 5, 1) = " `UserID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 5, 2) = " `Name` varchar(32) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 3) = " `Active` tinyint(1) DEFAULT NULL"
 Str(1, 5, 4) = " `Status` int(10) DEFAULT NULL"
 Str(1, 5, 5) = " `Beschreibung` varchar(32) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 6) = "  PRIMARY KEY (`UserID`)"
 Str(1, 5, 7) = "  UNIQUE KEY `PrimaryINDEX` (`UserID`)"
 Str(1, 5, 8) = "  KEY `UserID` (`UserID`)"
 Str(1, 5, 9) = " ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "blocktermin"
 Str(0, 6, 1) = "`Blocktermin_Nr`"
 Str(0, 6, 2) = "`Blocktermin_Text`"
 Str(0, 6, 3) = "`Start`"
 Str(0, 6, 4) = "`Ende`"
 Str(0, 6, 5) = "`Blocktermin_Nr`"
 Str(0, 6, 6) = "`MyConst`"
 ArtZ(0, 6) = 4
 ArtZ(1, 6) = 2
 Str(1, 6, 0) = "CREATE TABLE `blocktermin` ("
 Str(1, 6, 1) = " `Blocktermin_Nr` int(10) NOT NULL DEFAULT '0'"
 Str(1, 6, 2) = " `Blocktermin_Text` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 3) = " `Start` datetime DEFAULT NULL"
 Str(1, 6, 4) = " `Ende` datetime DEFAULT NULL"
 Str(1, 6, 5) = "  PRIMARY KEY (`Blocktermin_Nr`)"
 Str(1, 6, 6) = "  UNIQUE KEY `MyConst` (`Blocktermin_Nr`)"
 Str(1, 6, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "brief_anrede"
 Str(0, 7, 1) = "`Brief_Anrede_Nr`"
 Str(0, 7, 2) = "`Brief_Anrede`"
 Str(0, 7, 3) = "`Geschlecht`"
 Str(0, 7, 4) = "`KorrespNr`"
 Str(0, 7, 5) = "`Vertrautheit`"
 Str(0, 7, 6) = "`BAw`"
 Str(0, 7, 7) = "`BAm`"
 Str(0, 7, 8) = "`BAn`"
 Str(0, 7, 9) = "`Brief_Anrede_Nr`"
 Str(0, 7, 10) = "`Brief_Anrede_Nr`"
 Str(0, 7, 11) = "`Brief_Anrede`"
 Str(0, 7, 12) = "`System`"
 ArtZ(0, 7) = 8
 ArtZ(1, 7) = 4
 Str(1, 7, 0) = "CREATE TABLE `brief_anrede` ("
 Str(1, 7, 1) = " `Brief_Anrede_Nr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 7, 2) = " `Brief_Anrede` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 3) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 4) = " `KorrespNr` int(10) DEFAULT NULL"
 Str(1, 7, 5) = " `Vertrautheit` int(10) DEFAULT NULL COMMENT '1,2,3'"
 Str(1, 7, 6) = " `BAw` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Brief-Anrede weiblich'"
 Str(1, 7, 7) = " `BAm` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Brief-Anrede männlich'"
 Str(1, 7, 8) = " `BAn` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Brief-Anrede neutral'"
 Str(1, 7, 9) = "  PRIMARY KEY (`Brief_Anrede_Nr`)"
 Str(1, 7, 10) = "  UNIQUE KEY `Brief_Anrede_Nr` (`Brief_Anrede_Nr`)"
 Str(1, 7, 11) = "  KEY `Brief_Anrede` (`Brief_Anrede`)"
 Str(1, 7, 12) = "  KEY `System` (`Geschlecht`,`Vertrautheit`)"
 Str(1, 7, 13) = " ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "briefkopf"
 Str(0, 8, 1) = "`Briefkopf_Nr`"
 Str(0, 8, 2) = "`Briefkopf`"
 Str(0, 8, 3) = "`Geschlecht`"
 Str(0, 8, 4) = "`KorrespNr`"
 Str(0, 8, 5) = "`Form`"
 Str(0, 8, 6) = "`BKw`"
 Str(0, 8, 7) = "`BKm`"
 Str(0, 8, 8) = "`BKn`"
 Str(0, 8, 9) = "`Briefkopf_Nr`"
 Str(0, 8, 10) = "`Briefkopf_Nr`"
 Str(0, 8, 11) = "`Briefkopf`"
 Str(0, 8, 12) = "`System`"
 ArtZ(0, 8) = 8
 ArtZ(1, 8) = 4
 Str(1, 8, 0) = "CREATE TABLE `briefkopf` ("
 Str(1, 8, 1) = " `Briefkopf_Nr` int(10) NOT NULL DEFAULT '0'"
 Str(1, 8, 2) = " `Briefkopf` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 3) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'w,m'"
 Str(1, 8, 4) = " `KorrespNr` int(10) DEFAULT NULL COMMENT 'für jew. anderes Geschlecht'"
 Str(1, 8, 5) = " `Form` int(10) DEFAULT NULL COMMENT '0,1,2'"
 Str(1, 8, 6) = " `BKw` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Briefkopf weiblich'"
 Str(1, 8, 7) = " `BKm` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Briefkopf männlich'"
 Str(1, 8, 8) = " `BKn` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Briefkopf neutral'"
 Str(1, 8, 9) = "  PRIMARY KEY (`Briefkopf_Nr`)"
 Str(1, 8, 10) = "  UNIQUE KEY `Briefkopf_Nr` (`Briefkopf_Nr`)"
 Str(1, 8, 11) = "  KEY `Briefkopf` (`Briefkopf`)"
 Str(1, 8, 12) = "  KEY `System` (`Geschlecht`,`Form`)"
 Str(1, 8, 13) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "briefkopfa"
 Str(0, 9, 1) = "`Briefkopf_Nr`"
 Str(0, 9, 2) = "`Briefkopf`"
 Str(0, 9, 3) = "`Briefkopf_Nr`"
 Str(0, 9, 4) = "`PrimaryINDEX`"
 ArtZ(0, 9) = 2
 ArtZ(1, 9) = 2
 Str(1, 9, 0) = "CREATE TABLE `briefkopfa` ("
 Str(1, 9, 1) = " `Briefkopf_Nr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 9, 2) = " `Briefkopf` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 3) = "  PRIMARY KEY (`Briefkopf_Nr`)"
 Str(1, 9, 4) = "  UNIQUE KEY `PrimaryINDEX` (`Briefkopf_Nr`)"
 Str(1, 9, 5) = " ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "commandbars"
 Str(0, 10, 1) = "`ID`"
 Str(0, 10, 2) = "`BuiltIn`"
 Str(0, 10, 3) = "`Context`"
 Str(0, 10, 4) = "`Controls`"
 Str(0, 10, 5) = "`Creator`"
 Str(0, 10, 6) = "`Enabled`"
 Str(0, 10, 7) = "`Height`"
 Str(0, 10, 8) = "`Index`"
 Str(0, 10, 9) = "`Left`"
 Str(0, 10, 10) = "`Name`"
 Str(0, 10, 11) = "`NameLocal`"
 Str(0, 10, 12) = "`Parent`"
 Str(0, 10, 13) = "`Position`"
 Str(0, 10, 14) = "`Protection`"
 Str(0, 10, 15) = "`RowIndex`"
 Str(0, 10, 16) = "`Top`"
 Str(0, 10, 17) = "`Type`"
 Str(0, 10, 18) = "`Visible`"
 Str(0, 10, 19) = "`Width`"
 ArtZ(0, 10) = 19
 Str(1, 10, 0) = "CREATE TABLE `commandbars` ("
 Str(1, 10, 1) = " `ID` decimal(15,4) DEFAULT NULL"
 Str(1, 10, 2) = " `BuiltIn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 3) = " `Context` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 4) = " `Controls` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 5) = " `Creator` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 6) = " `Enabled` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 7) = " `Height` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 8) = " `Index` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 9) = " `Left` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 10) = " `Name` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 11) = " `NameLocal` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 12) = " `Parent` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 13) = " `Position` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 14) = " `Protection` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 15) = " `RowIndex` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 16) = " `Top` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 17) = " `Type` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 18) = " `Visible` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 19) = " `Width` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 10, 20) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr10

Sub FüllStr11()
 Str(0, 11, 0) = "dokument"
 Str(0, 11, 1) = "`Dokument_Nr`"
 Str(0, 11, 2) = "`Adresse_Nr`"
 Str(0, 11, 3) = "`Dokument_Info`"
 Str(0, 11, 4) = "`Dokument_Name`"
 Str(0, 11, 5) = "`Dokument_Vom`"
 Str(0, 11, 6) = "`Dokument_Nr`"
 Str(0, 11, 7) = "`MyConst`"
 ArtZ(0, 11) = 5
 ArtZ(1, 11) = 2
 Str(1, 11, 0) = "CREATE TABLE `dokument` ("
 Str(1, 11, 1) = " `Dokument_Nr` int(10) NOT NULL DEFAULT '0'"
 Str(1, 11, 2) = " `Adresse_Nr` int(10) DEFAULT NULL"
 Str(1, 11, 3) = " `Dokument_Info` varchar(64) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 11, 4) = " `Dokument_Name` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 11, 5) = " `Dokument_Vom` datetime DEFAULT NULL"
 Str(1, 11, 6) = "  PRIMARY KEY (`Dokument_Nr`)"
 Str(1, 11, 7) = "  UNIQUE KEY `MyConst` (`Dokument_Nr`)"
 Str(1, 11, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr11

Sub FüllStr12()
 Str(0, 12, 0) = "emails"
 Str(1, 12, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `emails` AS select `ez`.`KNr` AS `knr`,`ez`.`alt` AS `ezalt`,`ez`.`ab` AS `ab`,`ez`.`bis` AS `bis`,`ez`.`geändert` AS `ezgeändert`,`mails`.`ENr` AS `ENr`,`mails`.`Adresse_nr` AS `Adresse_nr`,`mails`.`Name` AS `Name`,`mails`.`Vorname` AS `Vorname`,`mails`.`Geburtstag` AS `Geburtstag`,`mails`.`IdKz` AS `IdKz`,`mails`.`Email` AS `Email`,`mails`.`Nr` AS `Nr`,`mails`.`alt` AS `alt`,`mails`.`geändert` AS `geändert` from (`emailzuordnung` `ez` left join `mails` on((`ez`.`ENr` = `mails`.`ENr`)))"
End Sub ' FüllStr12

Sub FüllStr13()
 Str(0, 13, 0) = "emailzuordnung"
 Str(0, 13, 1) = "`ENr`"
 Str(0, 13, 2) = "`KNr`"
 Str(0, 13, 3) = "`alt`"
 Str(0, 13, 4) = "`ab`"
 Str(0, 13, 5) = "`bis`"
 Str(0, 13, 6) = "`geändert`"
 Str(0, 13, 7) = "`ENr`"
 Str(0, 13, 8) = "`ENr`"
 Str(0, 13, 9) = "`EmailsEMailZuordnung`"
 Str(0, 13, 10) = "`EmailsEMailZuordnung_MySQLRel`"
 Str(0, 13, 11) = "`KontakteEMailZuordnung`"
 Str(0, 13, 12) = "`KontakteEMailZuordnung_MySQLRel`"
 Str(0, 13, 13) = "`EmailsEMailZuordnung_MySQLRel_AccRel`"
 Str(0, 13, 14) = "`KontakteEMailZuordnung_MySQLRel_AccRel`"
 ArtZ(0, 13) = 6
 ArtZ(1, 13) = 6
 ArtZ(2, 13) = 2
 Str(1, 13, 0) = "CREATE TABLE `emailzuordnung` ("
 Str(1, 13, 1) = " `ENr` int(10) NOT NULL DEFAULT '0' COMMENT 'Email-Nummer'"
 Str(1, 13, 2) = " `KNr` int(10) NOT NULL DEFAULT '0' COMMENT 'Kontakt-Nr'"
 Str(1, 13, 3) = " `alt` tinyint(1) DEFAULT NULL COMMENT 'nicht mehr gültig?'"
 Str(1, 13, 4) = " `ab` datetime DEFAULT NULL COMMENT 'gültig ab'"
 Str(1, 13, 5) = " `bis` datetime DEFAULT NULL COMMENT 'gültig bis'"
 Str(1, 13, 6) = " `geändert` datetime DEFAULT NULL COMMENT 'letzte Änderung'"
 Str(1, 13, 7) = "  PRIMARY KEY (`ENr`,`KNr`)"
 Str(1, 13, 8) = "  UNIQUE KEY `ENr` (`ENr`,`KNr`)"
 Str(1, 13, 9) = "  KEY `EmailsEMailZuordnung` (`ENr`)"
 Str(1, 13, 10) = "  KEY `EmailsEMailZuordnung_MySQLRel` (`ENr`)"
 Str(1, 13, 11) = "  KEY `KontakteEMailZuordnung` (`KNr`)"
 Str(1, 13, 12) = "  KEY `KontakteEMailZuordnung_MySQLRel` (`KNr`)"
 Str(1, 13, 13) = "  CONSTRAINT `EmailsEMailZuordnung_MySQLRel_AccRel` FOREIGN KEY (`ENr`) REFERENCES `mails` (`ENr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 13, 14) = "  CONSTRAINT `KontakteEMailZuordnung_MySQLRel_AccRel` FOREIGN KEY (`KNr`) REFERENCES `kontakte` (`KNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 13, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr13

Sub FüllStr14()
 Str(0, 14, 0) = "inst"
 Str(0, 14, 1) = "`Adresse_Nr`"
 Str(0, 14, 2) = "`Anschrift2`"
 Str(0, 14, 3) = "`Telefon 1`"
 Str(0, 14, 4) = "`Telefon 2`"
 Str(0, 14, 5) = "`Fax`"
 Str(0, 14, 6) = "`Mobilfunk`"
 Str(0, 14, 7) = "`WebSite`"
 Str(0, 14, 8) = "`Briefkopf_Nr`"
 Str(0, 14, 9) = "`Brief_Anrede_Nr`"
 Str(0, 14, 10) = "`Abteilung`"
 Str(0, 14, 11) = "`AdrGru_Nr`"
 Str(0, 14, 12) = "`Firma`"
 Str(0, 14, 13) = "`Firma2`"
 Str(0, 14, 14) = "`Öffnungszeiten`"
 Str(0, 14, 15) = "`FeldName1`"
 Str(0, 14, 16) = "`FeldInhalt1`"
 Str(0, 14, 17) = "`FeldName2`"
 Str(0, 14, 18) = "`FeldInhalt2`"
 Str(0, 14, 19) = "`Bemerkung`"
 Str(0, 14, 20) = "`Branche`"
 Str(0, 14, 21) = "`unsichtbar`"
 Str(0, 14, 22) = "`geaendert`"
 Str(0, 14, 23) = "`EtikAnr`"
 Str(0, 14, 24) = "`ehemalig`"
 Str(0, 14, 25) = "`Handy`"
 Str(0, 14, 26) = "`Ausdruck`"
 Str(0, 14, 27) = "`KNr`"
 Str(0, 14, 28) = "`Adresse_Nr`"
 Str(0, 14, 29) = "`KNr`"
 Str(0, 14, 30) = "`KontakteInst`"
 Str(0, 14, 31) = "`geaendert`"
 Str(0, 14, 32) = "`KontakteInst_MySQLRel_AccRel`"
 ArtZ(0, 14) = 27
 ArtZ(1, 14) = 4
 ArtZ(2, 14) = 1
 Str(1, 14, 0) = "CREATE TABLE `inst` ("
 Str(1, 14, 1) = " `Adresse_Nr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 14, 2) = " `Anschrift2` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 3) = " `Telefon 1` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 4) = " `Telefon 2` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 5) = " `Fax` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 6) = " `Mobilfunk` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 7) = " `WebSite` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 8) = " `Briefkopf_Nr` int(10) DEFAULT NULL"
 Str(1, 14, 9) = " `Brief_Anrede_Nr` int(10) DEFAULT NULL"
 Str(1, 14, 10) = " `Abteilung` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 11) = " `AdrGru_Nr` int(10) DEFAULT NULL"
 Str(1, 14, 12) = " `Firma` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 13) = " `Firma2` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 14) = " `Öffnungszeiten` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 15) = " `FeldName1` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 16) = " `FeldInhalt1` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 17) = " `FeldName2` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 18) = " `FeldInhalt2` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 19) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 14, 20) = " `Branche` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 21) = " `unsichtbar` tinyint(1) DEFAULT NULL"
 Str(1, 14, 22) = " `geaendert` datetime DEFAULT NULL"
 Str(1, 14, 23) = " `EtikAnr` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 24) = " `ehemalig` tinyint(1) DEFAULT NULL COMMENT 'SoleLuna: Ehemaliges Mitglied; sonst: ungültig geworden'"
 Str(1, 14, 25) = " `Handy` tinyint(1) DEFAULT NULL COMMENT 'Übertragung ans Handy'"
 Str(1, 14, 26) = " `Ausdruck` tinyint(1) DEFAULT NULL COMMENT 'Ausdruck auf Adressenliste'"
 Str(1, 14, 27) = " `KNr` int(10) DEFAULT NULL COMMENT 'Kontakt-Nummer'"
 Str(1, 14, 28) = "  UNIQUE KEY `Adresse_Nr` (`Adresse_Nr`)"
 Str(1, 14, 29) = "  UNIQUE KEY `KNr` (`KNr`)"
 Str(1, 14, 30) = "  UNIQUE KEY `KontakteInst` (`KNr`)"
 Str(1, 14, 31) = "  KEY `geaendert` (`geaendert`)"
 Str(1, 14, 32) = "  CONSTRAINT `KontakteInst_MySQLRel_AccRel` FOREIGN KEY (`KNr`) REFERENCES `kontakte` (`KNr`)"
 Str(1, 14, 33) = " ENGINE=InnoDB AUTO_INCREMENT=9403 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr14

Sub FüllStr15()
 Str(0, 15, 0) = "kalender"
 Str(0, 15, 1) = "`Datum`"
 Str(0, 15, 2) = "`Feiertag`"
 Str(0, 15, 3) = "`obfeier`"
 Str(0, 15, 4) = "`Geburtstage`"
 Str(0, 15, 5) = "`Termin`"
 Str(0, 15, 6) = "`HG`"
 Str(0, 15, 7) = "`VG`"
 Str(0, 15, 8) = "`Samba`"
 Str(0, 15, 9) = "`SonAuf`"
 Str(0, 15, 10) = "`SonUnt`"
 Str(0, 15, 11) = "`Mond`"
 Str(0, 15, 12) = "`Zusat`"
 Str(0, 15, 13) = "`geaendert`"
 Str(0, 15, 14) = "`Datum`"
 Str(0, 15, 15) = "`geaendert`"
 ArtZ(0, 15) = 13
 ArtZ(1, 15) = 2
 Str(1, 15, 0) = "CREATE TABLE `kalender` ("
 Str(1, 15, 1) = " `Datum` datetime DEFAULT NULL COMMENT 'Datum'"
 Str(1, 15, 2) = " `Feiertag` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name des Feiertags'"
 Str(1, 15, 3) = " `obfeier` int(10) DEFAULT NULL COMMENT '0 = kein Feiertag, 1 = Feiertag, 2 = halber Feiertag'"
 Str(1, 15, 4) = " `Geburtstage` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Geburtstage'"
 Str(1, 15, 5) = " `Termin` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'sonstige Termine'"
 Str(1, 15, 6) = " `HG` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Hintergrunddienst = ""H "", KV-Dienst = ""KV""'"
 Str(1, 15, 7) = " `VG` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Vordergrunddienst = Dienstart'"
 Str(1, 15, 8) = " `Samba` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Sambaauftritt'"
 Str(1, 15, 9) = " `SonAuf` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Sonnenaufgang'"
 Str(1, 15, 10) = " `SonUnt` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Sonnenuntergang'"
 Str(1, 15, 11) = " `Mond` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'N = Neumond, V = Vollmond'"
 Str(1, 15, 12) = " `Zusat` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Zusatzinfo, die nicht auf dem Ausdruck erscheint'"
 Str(1, 15, 13) = " `geaendert` datetime DEFAULT NULL COMMENT 'Datum und Uhrzeit der letzten Änderung'"
 Str(1, 15, 14) = "  UNIQUE KEY `Datum` (`Datum`)"
 Str(1, 15, 15) = "  KEY `geaendert` (`geaendert`)"
 Str(1, 15, 16) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr15

Sub FüllStr16()
 Str(0, 16, 0) = "kalender ab heute"
 Str(1, 16, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `kalender ab heute` AS select sql_no_cache `kalender`.`Datum` AS `Datum`,`kalender`.`Feiertag` AS `Feiertag`,`kalender`.`obfeier` AS `obfeier`,`kalender`.`Geburtstage` AS `Geburtstage`,`kalender`.`Termin` AS `Termin`,`kalender`.`HG` AS `HG`,`kalender`.`VG` AS `VG`,`kalender`.`Samba` AS `Samba`,`kalender`.`SonAuf` AS `SonAuf`,`kalender`.`SonUnt` AS `SonUnt`,`kalender`.`Mond` AS `Mond`,`kalender`.`Zusat` AS `Zusat`,`kalender`.`geaendert` AS `geaendert` from `kalender` where (`kalender`.`Datum` >= (now() - 7)) order by `kalender`.`Datum`"
End Sub ' FüllStr16

Sub FüllStr17()
 Str(0, 17, 0) = "kontakte"
 Str(0, 17, 1) = "`KNr`"
 Str(0, 17, 2) = "`Bezug`"
 Str(0, 17, 3) = "`alt`"
 Str(0, 17, 4) = "`geändert`"
 Str(0, 17, 5) = "`KNr`"
 ArtZ(0, 17) = 4
 ArtZ(1, 17) = 1
 Str(1, 17, 0) = "CREATE TABLE `kontakte` ("
 Str(1, 17, 1) = " `KNr` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Kontakt-Nummer'"
 Str(1, 17, 2) = " `Bezug` int(10) DEFAULT NULL COMMENT '1= Person (Tabelle Pers), 2 = Institution (Tabelle Inst)'"
 Str(1, 17, 3) = " `alt` tinyint(1) DEFAULT NULL COMMENT 'nicht mehr gültig?'"
 Str(1, 17, 4) = " `geändert` datetime DEFAULT NULL COMMENT 'letzte Änderung'"
 Str(1, 17, 5) = "  PRIMARY KEY (`KNr`)"
 Str(1, 17, 6) = " ENGINE=InnoDB AUTO_INCREMENT=4205 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr17

Sub FüllStr18()
 Str(0, 18, 0) = "konten"
 Str(0, 18, 1) = "`KNr`"
 Str(0, 18, 2) = "`Kto`"
 Str(0, 18, 3) = "`BLZ`"
 Str(0, 18, 4) = "`Bank`"
 Str(0, 18, 5) = "`Inhaber`"
 Str(0, 18, 6) = "`alt`"
 Str(0, 18, 7) = "`ab`"
 Str(0, 18, 8) = "`bis`"
 Str(0, 18, 9) = "`geändert`"
 Str(0, 18, 10) = "`KNr`"
 Str(0, 18, 11) = "`KontoVerb`"
 Str(0, 18, 12) = "`BVerb`"
 Str(0, 18, 13) = "`KontakteKonten`"
 Str(0, 18, 14) = "`KontenKNr`"
 Str(0, 18, 15) = "`Konto-Nummer`"
 Str(0, 18, 16) = "`KontakteKonten_AccRel`"
 ArtZ(0, 18) = 9
 ArtZ(1, 18) = 6
 ArtZ(2, 18) = 1
 Str(1, 18, 0) = "CREATE TABLE `konten` ("
 Str(1, 18, 1) = " `KNr` int(10) NOT NULL DEFAULT '0' COMMENT 'Kontakt-Nr'"
 Str(1, 18, 2) = " `Kto` varchar(50) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Kontonummer'"
 Str(1, 18, 3) = " `BLZ` varchar(50) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Bankleitzahl'"
 Str(1, 18, 4) = " `Bank` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Bank'"
 Str(1, 18, 5) = " `Inhaber` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Inhaber'"
 Str(1, 18, 6) = " `alt` tinyint(1) DEFAULT NULL COMMENT 'Konto alt?'"
 Str(1, 18, 7) = " `ab` datetime DEFAULT NULL COMMENT 'gültig ab'"
 Str(1, 18, 8) = " `bis` datetime DEFAULT NULL COMMENT 'gültig bis'"
 Str(1, 18, 9) = " `geändert` datetime DEFAULT NULL COMMENT 'letzte Änderung'"
 Str(1, 18, 10) = "  PRIMARY KEY (`KNr`,`Kto`,`BLZ`)"
 Str(1, 18, 11) = "  UNIQUE KEY `KontoVerb` (`KNr`,`Kto`,`BLZ`)"
 Str(1, 18, 12) = "  KEY `BVerb` (`Kto`,`BLZ`)"
 Str(1, 18, 13) = "  KEY `KontakteKonten` (`KNr`)"
 Str(1, 18, 14) = "  KEY `KontenKNr` (`KNr`)"
 Str(1, 18, 15) = "  KEY `Konto-Nummer` (`Kto`)"
 Str(1, 18, 16) = "  CONSTRAINT `KontakteKonten_AccRel` FOREIGN KEY (`KNr`) REFERENCES `kontakte` (`KNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 18, 17) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr18

Sub FüllStr19()
 Str(0, 19, 0) = "mails"
 Str(0, 19, 1) = "`ENr`"
 Str(0, 19, 2) = "`Adresse_nr`"
 Str(0, 19, 3) = "`Name`"
 Str(0, 19, 4) = "`Vorname`"
 Str(0, 19, 5) = "`Geburtstag`"
 Str(0, 19, 6) = "`IdKz`"
 Str(0, 19, 7) = "`Email`"
 Str(0, 19, 8) = "`Nr`"
 Str(0, 19, 9) = "`alt`"
 Str(0, 19, 10) = "`geändert`"
 Str(0, 19, 11) = "`ENr`"
 Str(0, 19, 12) = "`ENr`"
 Str(0, 19, 13) = "`EindInh`"
 Str(0, 19, 14) = "`Adresse_nr`"
 Str(0, 19, 15) = "`AdresseEmails`"
 Str(0, 19, 16) = "`AdresseEmails_MySQLRel`"
 Str(0, 19, 17) = "`Email`"
 Str(0, 19, 18) = "`EmailAkt`"
 Str(0, 19, 19) = "`IdKz`"
 Str(0, 19, 20) = "`Zuordnung`"
 Str(0, 19, 21) = "`AdresseEmails_MySQLRel_AccRel`"
 ArtZ(0, 19) = 10
 ArtZ(1, 19) = 10
 ArtZ(2, 19) = 1
 Str(1, 19, 0) = "CREATE TABLE `mails` ("
 Str(1, 19, 1) = " `ENr` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel der Email-Adresse'"
 Str(1, 19, 2) = " `Adresse_nr` int(10) DEFAULT NULL"
 Str(1, 19, 3) = " `Name` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 4) = " `Vorname` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 5) = " `Geburtstag` datetime DEFAULT NULL"
 Str(1, 19, 6) = " `IdKz` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 7) = " `Email` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Email-Adresse'"
 Str(1, 19, 8) = " `Nr` int(10) DEFAULT NULL COMMENT 'Reihenfolgenummer der Email für Adressen'"
 Str(1, 19, 9) = " `alt` tinyint(1) DEFAULT NULL COMMENT 'nicht mehr akutell?'"
 Str(1, 19, 10) = " `geändert` datetime DEFAULT NULL COMMENT 'letzte Änderung'"
 Str(1, 19, 11) = "  PRIMARY KEY (`ENr`)"
 Str(1, 19, 12) = "  UNIQUE KEY `ENr` (`ENr`)"
 Str(1, 19, 13) = "  UNIQUE KEY `EindInh` (`Adresse_nr`,`Email`)"
 Str(1, 19, 14) = "  KEY `Adresse_nr` (`Adresse_nr`,`ENr`)"
 Str(1, 19, 15) = "  KEY `AdresseEmails` (`Adresse_nr`)"
 Str(1, 19, 16) = "  KEY `AdresseEmails_MySQLRel` (`Adresse_nr`)"
 Str(1, 19, 17) = "  KEY `Email` (`Email`)"
 Str(1, 19, 18) = "  KEY `EmailAkt` (`alt`,`Email`)"
 Str(1, 19, 19) = "  KEY `IdKz` (`IdKz`)"
 Str(1, 19, 20) = "  KEY `Zuordnung` (`Name`,`Vorname`,`Geburtstag`,`IdKz`)"
 Str(1, 19, 21) = "  CONSTRAINT `AdresseEmails_MySQLRel_AccRel` FOREIGN KEY (`Adresse_nr`) REFERENCES `adresse` (`Adresse_Nr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 19, 22) = " ENGINE=InnoDB AUTO_INCREMENT=1362 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr19

Sub FüllStr20()
 Str(0, 20, 0) = "notiz"
 Str(0, 20, 1) = "`Notiz_Nr`"
 Str(0, 20, 2) = "`Text`"
 Str(0, 20, 3) = "`Termin_Nr`"
 Str(0, 20, 4) = "`Erstellt_am`"
 Str(0, 20, 5) = "`Geaendert_am`"
 Str(0, 20, 6) = "`Thema`"
 Str(0, 20, 7) = "`Notiz_Nr`"
 Str(0, 20, 8) = "`PrimaryINDEX`"
 ArtZ(0, 20) = 6
 ArtZ(1, 20) = 2
 Str(1, 20, 0) = "CREATE TABLE `notiz` ("
 Str(1, 20, 1) = " `Notiz_Nr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 20, 2) = " `Text` longtext COLLATE latin1_german2_ci"
 Str(1, 20, 3) = " `Termin_Nr` int(10) DEFAULT NULL"
 Str(1, 20, 4) = " `Erstellt_am` datetime DEFAULT NULL"
 Str(1, 20, 5) = " `Geaendert_am` datetime DEFAULT NULL"
 Str(1, 20, 6) = " `Thema` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 7) = "  PRIMARY KEY (`Notiz_Nr`)"
 Str(1, 20, 8) = "  UNIQUE KEY `PrimaryINDEX` (`Notiz_Nr`)"
 Str(1, 20, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr20

Sub FüllStr21()
 Str(0, 21, 0) = "pers"
 Str(0, 21, 1) = "`Adresse_Nr`"
 Str(0, 21, 2) = "`Name`"
 Str(0, 21, 3) = "`Vorname`"
 Str(0, 21, 4) = "`Titel`"
 Str(0, 21, 5) = "`Geburtsname`"
 Str(0, 21, 6) = "`IdKz`"
 Str(0, 21, 7) = "`Gschl`"
 Str(0, 21, 8) = "`Briefkopf_Nr`"
 Str(0, 21, 9) = "`Brief_Anrede_Nr`"
 Str(0, 21, 10) = "`Geburtstag`"
 Str(0, 21, 11) = "`Bemerkung`"
 Str(0, 21, 12) = "`unsichtbar`"
 Str(0, 21, 13) = "`Geb_unsichtbar`"
 Str(0, 21, 14) = "`geaendert`"
 Str(0, 21, 15) = "`obFest`"
 Str(0, 21, 16) = "`EtikAnr`"
 Str(0, 21, 17) = "`ehemalig`"
 Str(0, 21, 18) = "`Handy`"
 Str(0, 21, 19) = "`Ausdruck`"
 Str(0, 21, 20) = "`KNr`"
 Str(0, 21, 21) = "`Adresse_Nr`"
 Str(0, 21, 22) = "`Identität`"
 Str(0, 21, 23) = "`KontaktePers`"
 Str(0, 21, 24) = "`Brief_AnredePers`"
 Str(0, 21, 25) = "`Brief_AnredePers_MySQLRel`"
 Str(0, 21, 26) = "`BriefkopfPers`"
 Str(0, 21, 27) = "`BriefkopfPers_MySQLRel`"
 Str(0, 21, 28) = "`geaendert`"
 Str(0, 21, 29) = "`idkz`"
 Str(0, 21, 30) = "`KNr`"
 Str(0, 21, 31) = "`PersBriefkopf_Nr`"
 Str(0, 21, 32) = "`BriefkopfPers_MySQLRel_AccRel`"
 Str(0, 21, 33) = "`Brief_AnredePers_MySQLRel_AccRel`"
 Str(0, 21, 34) = "`KontaktePers_MySQLRel_AccRel`"
 ArtZ(0, 21) = 20
 ArtZ(1, 21) = 11
 ArtZ(2, 21) = 3
 Str(1, 21, 0) = "CREATE TABLE `pers` ("
 Str(1, 21, 1) = " `Adresse_Nr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 21, 2) = " `Name` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 3) = " `Vorname` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 4) = " `Titel` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 5) = " `Geburtsname` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 6) = " `IdKz` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Identitätskennzeichen, falls Name+Vorname+Gebd nicht eindeutig'"
 Str(1, 21, 7) = " `Gschl` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 8) = " `Briefkopf_Nr` int(10) DEFAULT NULL"
 Str(1, 21, 9) = " `Brief_Anrede_Nr` int(10) DEFAULT NULL"
 Str(1, 21, 10) = " `Geburtstag` datetime DEFAULT NULL"
 Str(1, 21, 11) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 21, 12) = " `unsichtbar` tinyint(1) DEFAULT NULL"
 Str(1, 21, 13) = " `Geb_unsichtbar` tinyint(1) DEFAULT NULL"
 Str(1, 21, 14) = " `geaendert` datetime DEFAULT NULL"
 Str(1, 21, 15) = " `obFest` tinyint(1) DEFAULT NULL"
 Str(1, 21, 16) = " `EtikAnr` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 17) = " `ehemalig` tinyint(1) DEFAULT NULL COMMENT 'SoleLuna: Ehemaliges Mitglied; sonst: ungültig geworden'"
 Str(1, 21, 18) = " `Handy` tinyint(1) DEFAULT NULL COMMENT 'Übertragung ans Handy'"
 Str(1, 21, 19) = " `Ausdruck` tinyint(1) DEFAULT NULL COMMENT 'Ausdruck auf Adressenliste'"
 Str(1, 21, 20) = " `KNr` int(10) DEFAULT NULL COMMENT 'Kontakt-Nummer'"
 Str(1, 21, 21) = "  UNIQUE KEY `Adresse_Nr` (`Adresse_Nr`)"
 Str(1, 21, 22) = "  UNIQUE KEY `Identität` (`Name`,`Vorname`,`Geburtstag`,`IdKz`)"
 Str(1, 21, 23) = "  UNIQUE KEY `KontaktePers` (`KNr`)"
 Str(1, 21, 24) = "  KEY `Brief_AnredePers` (`Brief_Anrede_Nr`)"
 Str(1, 21, 25) = "  KEY `Brief_AnredePers_MySQLRel` (`Brief_Anrede_Nr`)"
 Str(1, 21, 26) = "  KEY `BriefkopfPers` (`Briefkopf_Nr`)"
 Str(1, 21, 27) = "  KEY `BriefkopfPers_MySQLRel` (`Briefkopf_Nr`)"
 Str(1, 21, 28) = "  KEY `geaendert` (`geaendert`)"
 Str(1, 21, 29) = "  KEY `idkz` (`IdKz`)"
 Str(1, 21, 30) = "  KEY `KNr` (`KNr`)"
 Str(1, 21, 31) = "  KEY `PersBriefkopf_Nr` (`Briefkopf_Nr`)"
 Str(1, 21, 32) = "  CONSTRAINT `BriefkopfPers_MySQLRel_AccRel` FOREIGN KEY (`Briefkopf_Nr`) REFERENCES `briefkopf` (`Briefkopf_Nr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 21, 33) = "  CONSTRAINT `Brief_AnredePers_MySQLRel_AccRel` FOREIGN KEY (`Brief_Anrede_Nr`) REFERENCES `brief_anrede` (`Brief_Anrede_Nr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 21, 34) = "  CONSTRAINT `KontaktePers_MySQLRel_AccRel` FOREIGN KEY (`KNr`) REFERENCES `kontakte` (`KNr`)"
 Str(1, 21, 35) = " ENGINE=InnoDB AUTO_INCREMENT=9403 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr21

Sub FüllStr22()
 Str(0, 22, 0) = "persinstzuord"
 Str(0, 22, 1) = "`PIZ_NR`"
 Str(0, 22, 2) = "`PKnr`"
 Str(0, 22, 3) = "`IKnr`"
 Str(0, 22, 4) = "`alt`"
 Str(0, 22, 5) = "`ab`"
 Str(0, 22, 6) = "`bis`"
 Str(0, 22, 7) = "`geändert`"
 Str(0, 22, 8) = "`PIZ_NR`"
 Str(0, 22, 9) = "`PrimaryINDEX`"
 Str(0, 22, 10) = "`IKnr`"
 Str(0, 22, 11) = "`InstPersInstZuord`"
 Str(0, 22, 12) = "`InstPersInstZuord_MySQLRel`"
 Str(0, 22, 13) = "`PersPersInstZuord`"
 Str(0, 22, 14) = "`PersPersInstZuord_MySQLRel`"
 Str(0, 22, 15) = "`PKnr`"
 Str(0, 22, 16) = "`InstPersInstZuord_MySQLRel_AccRel`"
 Str(0, 22, 17) = "`PersPersInstZuord_MySQLRel_AccRel`"
 ArtZ(0, 22) = 7
 ArtZ(1, 22) = 8
 ArtZ(2, 22) = 2
 Str(1, 22, 0) = "CREATE TABLE `persinstzuord` ("
 Str(1, 22, 1) = " `PIZ_NR` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Person-Institutions-Zuordnungs-Nummer'"
 Str(1, 22, 2) = " `PKnr` int(10) DEFAULT NULL COMMENT 'Adreß-Kontakt-Nummer'"
 Str(1, 22, 3) = " `IKnr` int(10) DEFAULT NULL COMMENT 'Institutions-Kontakt-Nummer'"
 Str(1, 22, 4) = " `alt` tinyint(1) DEFAULT NULL COMMENT 'ob Zuordnung ungültig'"
 Str(1, 22, 5) = " `ab` datetime DEFAULT NULL COMMENT 'Zuordnung gültig ab'"
 Str(1, 22, 6) = " `bis` datetime DEFAULT NULL COMMENT 'Zuordnung gültig bis'"
 Str(1, 22, 7) = " `geändert` datetime DEFAULT NULL COMMENT 'letzte Änderung'"
 Str(1, 22, 8) = "  PRIMARY KEY (`PIZ_NR`)"
 Str(1, 22, 9) = "  UNIQUE KEY `PrimaryINDEX` (`PIZ_NR`)"
 Str(1, 22, 10) = "  KEY `IKnr` (`IKnr`,`PKnr`)"
 Str(1, 22, 11) = "  KEY `InstPersInstZuord` (`IKnr`)"
 Str(1, 22, 12) = "  KEY `InstPersInstZuord_MySQLRel` (`IKnr`)"
 Str(1, 22, 13) = "  KEY `PersPersInstZuord` (`PKnr`)"
 Str(1, 22, 14) = "  KEY `PersPersInstZuord_MySQLRel` (`PKnr`)"
 Str(1, 22, 15) = "  KEY `PKnr` (`PKnr`,`IKnr`)"
 Str(1, 22, 16) = "  CONSTRAINT `InstPersInstZuord_MySQLRel_AccRel` FOREIGN KEY (`IKnr`) REFERENCES `inst` (`KNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 22, 17) = "  CONSTRAINT `PersPersInstZuord_MySQLRel_AccRel` FOREIGN KEY (`PKnr`) REFERENCES `pers` (`KNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 22, 18) = " ENGINE=InnoDB AUTO_INCREMENT=366 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr22

Sub FüllStr23()
 Str(0, 23, 0) = "pfade"
 Str(0, 23, 1) = "`ID`"
 Str(0, 23, 2) = "`Tabelle`"
 Str(0, 23, 3) = "`Computer`"
 Str(0, 23, 4) = "`obAktiv`"
 Str(0, 23, 5) = "`Pfad`"
 Str(0, 23, 6) = "`Zweck`"
 Str(0, 23, 7) = "`MitFunktion`"
 Str(0, 23, 8) = "`geaendert1`"
 Str(0, 23, 9) = "`geaendert2`"
 Str(0, 23, 10) = "`word8`"
 Str(0, 23, 11) = "`ID`"
 Str(0, 23, 12) = "`PrimaryINDEX`"
 Str(0, 23, 13) = "`GetPfad`"
 Str(0, 23, 14) = "`Such`"
 ArtZ(0, 23) = 10
 ArtZ(1, 23) = 4
 Str(1, 23, 0) = "CREATE TABLE `pfade` ("
 Str(1, 23, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 23, 2) = " `Tabelle` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Dateiname ohne Endung'"
 Str(1, 23, 3) = " `Computer` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Computername aus HKEY_LOCAL_MACHINESystemCurrentControlSetComputerNameComputerName, dort ComputerName'"
 Str(1, 23, 4) = " `obAktiv` tinyint(1) DEFAULT NULL COMMENT 'aktuelle Tabellen/Computer-Kombination aktiv?'"
 Str(1, 23, 5) = " `Pfad` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Pfad auf diesem Computer mit ohne ohne Slash'"
 Str(1, 23, 6) = " `Zweck` int(10) DEFAULT NULL COMMENT '0 = inaktiv, 1= Daten sollen auf diesem Computer aus der genannten Datei geholt werden, 2 = Ja = Pfad zum Schreiben der Ausdruckdaten Adressenliste (Tabelle: Adresse) und Kalenderdaten (Tabelle: Kalender)'"
 Str(1, 23, 7) = " `MitFunktion` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Funktion, mit der die Daten aus dieser Datei geholt werden sollen'"
 Str(1, 23, 8) = " `geaendert1` datetime DEFAULT NULL COMMENT 'Änderungsdatum der ersten auswärtigen Tabelle, die aus der unter Pfad+Tabelle zusammengesetzten Datenbank stammt'"
 Str(1, 23, 9) = " `geaendert2` datetime DEFAULT NULL COMMENT 'Änderungsdatum derzweiten auswärtigen Tabelle, die aus der unter Pfad+Tabelle zusammengesetzten Datenbank stammt'"
 Str(1, 23, 10) = " `word8` tinyint(1) DEFAULT NULL COMMENT 'ja = Word 97 wird verwendet, nein = Word 95 wird verwendet'"
 Str(1, 23, 11) = "  PRIMARY KEY (`ID`)"
 Str(1, 23, 12) = "  UNIQUE KEY `PrimaryINDEX` (`ID`)"
 Str(1, 23, 13) = "  KEY `GetPfad` (`Tabelle`,`Computer`)"
 Str(1, 23, 14) = "  KEY `Such` (`Zweck`,`Computer`)"
 Str(1, 23, 15) = " ENGINE=InnoDB AUTO_INCREMENT=45 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr23

Sub FüllStr24()
 Str(0, 24, 0) = "system"
 Str(0, 24, 1) = "`Version`"
 Str(0, 24, 2) = "`LastAccess`"
 Str(0, 24, 3) = "`Password`"
 ArtZ(0, 24) = 3
 Str(1, 24, 0) = "CREATE TABLE `system` ("
 Str(1, 24, 1) = " `Version` int(10) DEFAULT NULL"
 Str(1, 24, 2) = " `LastAccess` datetime DEFAULT NULL"
 Str(1, 24, 3) = " `Password` varchar(32) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 24, 4) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr24

Sub FüllStr25()
 Str(0, 25, 0) = "telefon"
 Str(0, 25, 1) = "`KNr`"
 Str(0, 25, 2) = "`Num`"
 Str(0, 25, 3) = "`Art`"
 Str(0, 25, 4) = "`Bemerkung`"
 Str(0, 25, 5) = "`beru`"
 Str(0, 25, 6) = "`alt`"
 Str(0, 25, 7) = "`ab`"
 Str(0, 25, 8) = "`bis`"
 Str(0, 25, 9) = "`geändert`"
 Str(0, 25, 10) = "`KNr`"
 Str(0, 25, 11) = "`KNrNum`"
 Str(0, 25, 12) = "`Num`"
 Str(0, 25, 13) = "`telefonKontakte_MySQLRel`"
 Str(0, 25, 14) = "`telefonKontakte_MySQLRel_AccRel`"
 ArtZ(0, 25) = 9
 ArtZ(1, 25) = 4
 ArtZ(2, 25) = 1
 Str(1, 25, 0) = "CREATE TABLE `telefon` ("
 Str(1, 25, 1) = " `KNr` int(10) DEFAULT NULL COMMENT 'Kontakt-Nr, Bezug auf Pers und Inst'"
 Str(1, 25, 2) = " `Num` varchar(150) CHARACTER SET utf8 DEFAULT NULL COMMENT 'Telefon-, Fax- oder Handynummer'"
 Str(1, 25, 3) = " `Art` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 't=Telefon, f=Fax, b=beides, m=Mobil'"
 Str(1, 25, 4) = " `Bemerkung` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kommentar'"
 Str(1, 25, 5) = " `beru` tinyint(1) DEFAULT NULL COMMENT 'wahr = beruflich'"
 Str(1, 25, 6) = " `alt` tinyint(1) DEFAULT NULL COMMENT 'nicht mehr gültig?'"
 Str(1, 25, 7) = " `ab` datetime DEFAULT NULL COMMENT 'gültig ab'"
 Str(1, 25, 8) = " `bis` datetime DEFAULT NULL COMMENT 'gültig bis'"
 Str(1, 25, 9) = " `geändert` datetime DEFAULT NULL COMMENT 'letzte Änderung'"
 Str(1, 25, 10) = "  KEY `KNr` (`KNr`)"
 Str(1, 25, 11) = "  KEY `KNrNum` (`KNr`,`Num`)"
 Str(1, 25, 12) = "  KEY `Num` (`Num`)"
 Str(1, 25, 13) = "  KEY `telefonKontakte_MySQLRel` (`KNr`)"
 Str(1, 25, 14) = "  CONSTRAINT `telefonKontakte_MySQLRel_AccRel` FOREIGN KEY (`KNr`) REFERENCES `kontakte` (`KNr`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 25, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr25

Sub FüllStr26()
 Str(0, 26, 0) = "telefonliste"
 Str(1, 26, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `telefonliste` AS select `pers`.`Name` AS `name`,`pers`.`Vorname` AS `vorname`,`pers`.`Titel` AS `titel`,`telefon`.`Num` AS `num`,`telefon`.`Art` AS `art`,`telefon`.`beru` AS `beru`,`telefon`.`alt` AS `alt`,`telefon`.`ab` AS `ab`,`telefon`.`bis` AS `bis`,`telefon`.`geändert` AS `geändert` from (`pers` left join `telefon` on((`pers`.`KNr` = `telefon`.`KNr`))) order by `pers`.`Name`,`pers`.`Vorname`"
End Sub ' FüllStr26

Sub FüllStr27()
 Str(0, 27, 0) = "termin"
 Str(0, 27, 1) = "`Termin_Nr`"
 Str(0, 27, 2) = "`Termin_Text`"
 Str(0, 27, 3) = "`Adresse_Nr`"
 Str(0, 27, 4) = "`Termin_Typ`"
 Str(0, 27, 5) = "`Wiedervorlage_am`"
 Str(0, 27, 6) = "`TerminBeginn`"
 Str(0, 27, 7) = "`Status`"
 Str(0, 27, 8) = "`TerminEnde`"
 Str(0, 27, 9) = "`AutoTermin`"
 Str(0, 27, 10) = "`todo`"
 Str(0, 27, 11) = "`Prioritaet`"
 Str(0, 27, 12) = "`AutoErinnerung`"
 Str(0, 27, 13) = "`UserID`"
 Str(0, 27, 14) = "`ExchangeStatus`"
 Str(0, 27, 15) = "`TerminAlarm`"
 Str(0, 27, 16) = "`AlarmWert`"
 Str(0, 27, 17) = "`AlarmEinheit`"
 Str(0, 27, 18) = "`AlarmWiederholungstyp`"
 Str(0, 27, 19) = "`Blocktermin_Nr`"
 Str(0, 27, 20) = "`Termin_Nr`"
 Str(0, 27, 21) = "`PrimaryINDEX`"
 Str(0, 27, 22) = "`TerminAdresse_Nr`"
 Str(0, 27, 23) = "`UserID`"
 Str(0, 27, 24) = "`Zeitpunkt`"
 ArtZ(0, 27) = 19
 ArtZ(1, 27) = 5
 Str(1, 27, 0) = "CREATE TABLE `termin` ("
 Str(1, 27, 1) = " `Termin_Nr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 27, 2) = " `Termin_Text` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 27, 3) = " `Adresse_Nr` int(10) DEFAULT NULL"
 Str(1, 27, 4) = " `Termin_Typ` int(10) DEFAULT NULL"
 Str(1, 27, 5) = " `Wiedervorlage_am` datetime DEFAULT NULL"
 Str(1, 27, 6) = " `TerminBeginn` datetime DEFAULT NULL"
 Str(1, 27, 7) = " `Status` int(10) DEFAULT NULL"
 Str(1, 27, 8) = " `TerminEnde` datetime DEFAULT NULL"
 Str(1, 27, 9) = " `AutoTermin` tinyint(1) DEFAULT NULL"
 Str(1, 27, 10) = " `todo` tinyint(1) DEFAULT NULL"
 Str(1, 27, 11) = " `Prioritaet` int(10) DEFAULT NULL"
 Str(1, 27, 12) = " `AutoErinnerung` tinyint(1) DEFAULT NULL"
 Str(1, 27, 13) = " `UserID` int(10) DEFAULT NULL"
 Str(1, 27, 14) = " `ExchangeStatus` int(10) DEFAULT NULL"
 Str(1, 27, 15) = " `TerminAlarm` datetime DEFAULT NULL"
 Str(1, 27, 16) = " `AlarmWert` int(10) DEFAULT NULL"
 Str(1, 27, 17) = " `AlarmEinheit` int(10) DEFAULT NULL"
 Str(1, 27, 18) = " `AlarmWiederholungstyp` int(10) DEFAULT NULL"
 Str(1, 27, 19) = " `Blocktermin_Nr` int(10) DEFAULT NULL"
 Str(1, 27, 20) = "  PRIMARY KEY (`Termin_Nr`)"
 Str(1, 27, 21) = "  UNIQUE KEY `PrimaryINDEX` (`Termin_Nr`)"
 Str(1, 27, 22) = "  KEY `TerminAdresse_Nr` (`Adresse_Nr`)"
 Str(1, 27, 23) = "  KEY `UserID` (`UserID`)"
 Str(1, 27, 24) = "  KEY `Zeitpunkt` (`TerminBeginn`)"
 Str(1, 27, 25) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr27

Sub FüllStr28()
 Str(0, 28, 0) = "test"
 Str(0, 28, 1) = "`id`"
 Str(0, 28, 2) = "`t`"
 Str(0, 28, 3) = "`id`"
 ArtZ(0, 28) = 2
 ArtZ(1, 28) = 1
 Str(1, 28, 0) = "CREATE TABLE `test` ("
 Str(1, 28, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 28, 2) = " `t` bit(1) DEFAULT NULL"
 Str(1, 28, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 28, 4) = " ENGINE=MyISAM AUTO_INCREMENT=4 DEFAULT CHARSET=latin1"
End Sub ' FüllStr28

Sub FüllStr29()
 Str(0, 29, 0) = "zusatzfeld"
 Str(0, 29, 1) = "`Zusatzfeld_Nr`"
 Str(0, 29, 2) = "`Zusatzfeld_Name`"
 Str(0, 29, 3) = "`Zusatzfeld_Nr`"
 Str(0, 29, 4) = "`Zusatzfeld_Nr`"
 Str(0, 29, 5) = "`Zusatzfeld_Name`"
 ArtZ(0, 29) = 2
 ArtZ(1, 29) = 3
 Str(1, 29, 0) = "CREATE TABLE `zusatzfeld` ("
 Str(1, 29, 1) = " `Zusatzfeld_Nr` int(10) NOT NULL DEFAULT '0'"
 Str(1, 29, 2) = " `Zusatzfeld_Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 29, 3) = "  PRIMARY KEY (`Zusatzfeld_Nr`)"
 Str(1, 29, 4) = "  UNIQUE KEY `Zusatzfeld_Nr` (`Zusatzfeld_Nr`)"
 Str(1, 29, 5) = "  KEY `Zusatzfeld_Name` (`Zusatzfeld_Name`)"
 Str(1, 29, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr29

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

Public Function doMach_office(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
 FüllStr28
 FüllStr29
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 29
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
 For i = 0 To 29
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
  For i = 0 To 29
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
  MsgBox "Fertig mit doMach_office(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_office/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_office

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
