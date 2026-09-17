'Bauanleitung für eine Datenbank wie `//linux/kvaerzte` vom 15.11.09 14:26:22
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 10, 44) As new CString, ArtZ&(3, 10)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "GynäkologenImEinzugsbereich"
 Str(1, 0, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `GynäkologenImEinzugsbereich` AS select `hae`.`Vorname` AS `vorname`,`hae`.`Nachname` AS `nachname`,`hae`.`Straße` AS `straße`,`hae`.`PLZ` AS `plz`,`hae`.`ort` AS `ort` from `hae` where ((((convert(`hae`.`ZulG` using utf8) like '%Gebu%') and ((convert(`hae`.`BStelle` using utf8) like '%dachau%') or (convert(`hae`.`BStelle` using utf8) like '%Fürstenf%') or (convert(`hae`.`BStelle` using utf8) like '%freising%') or (`hae`.`PLZ` in (80992,80993,80995,80997,80999,81247,81249,81243,81245,81249,82110,85716,85386,85375,85777,85402,85375,85354,85302,85305)))) or (convert(`hae`.`HAName` using utf8) like '%Ringmaier%')) and (convert(`hae`.`Nachname` using utf8) <> 'Peschers') and (convert(`hae`.`Nachname` using utf8) <> 'Fehlhaber') and (convert(`hae`.`ort` using utf8) <> 'Moosburg')) group by `hae`.`PLZ`,`ha" & _ 
  "e`.`Straße`"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "SPP"
 Str(0, 1, 1) = "`idSPP`"
 Str(0, 1, 2) = "`Name`"
 Str(0, 1, 3) = "`Email`"
 Str(0, 1, 4) = "`ungültig`"
 Str(0, 1, 5) = "`idSPP`"
 Str(0, 1, 6) = "`Email`"
 ArtZ(0, 1) = 4
 ArtZ(1, 1) = 2
 Str(1, 1, 0) = "CREATE TABLE `SPP` ("
 Str(1, 1, 1) = " `idSPP` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 1, 2) = " `Name` varchar(45) NOT NULL"
 Str(1, 1, 3) = " `Email` varchar(150) NOT NULL"
 Str(1, 1, 4) = " `ungültig` bit(1) NOT NULL"
 Str(1, 1, 5) = "  PRIMARY KEY (`idSPP`)"
 Str(1, 1, 6) = "  UNIQUE KEY `Email` (`Email`)"
 Str(1, 1, 7) = " ENGINE=InnoDB AUTO_INCREMENT=461 DEFAULT CHARSET=latin1"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "bezirke"
 Str(0, 2, 1) = "`id`"
 Str(0, 2, 2) = "`Bezirk`"
 Str(0, 2, 3) = "`Nr`"
 Str(0, 2, 4) = "`MaxNr`"
 Str(0, 2, 5) = "`id`"
 Str(0, 2, 6) = "`id`"
 ArtZ(0, 2) = 4
 ArtZ(1, 2) = 2
 Str(1, 2, 0) = "CREATE TABLE `bezirke` ("
 Str(1, 2, 1) = " `id` int(10) NOT NULL DEFAULT '0'"
 Str(1, 2, 2) = " `Bezirk` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 3) = " `Nr` int(10) DEFAULT NULL"
 Str(1, 2, 4) = " `MaxNr` int(10) DEFAULT NULL"
 Str(1, 2, 5) = "  PRIMARY KEY (`id`)"
 Str(1, 2, 6) = "  UNIQUE KEY `id` (`id`)"
 Str(1, 2, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "chefaerzte"
 Str(0, 3, 1) = "`titel`"
 Str(0, 3, 2) = "`vorname`"
 Str(0, 3, 3) = "`nachname`"
 Str(0, 3, 4) = "`tel`"
 Str(0, 3, 5) = "`klinik`"
 Str(0, 3, 6) = "`abteilung`"
 Str(0, 3, 7) = "`email`"
 Str(0, 3, 8) = "`strasse`"
 Str(0, 3, 9) = "`ort`"
 Str(0, 3, 10) = "`Nr`"
 Str(0, 3, 11) = "`Nr`"
 ArtZ(0, 3) = 10
 ArtZ(1, 3) = 1
 Str(1, 3, 0) = "CREATE TABLE `chefaerzte` ("
 Str(1, 3, 1) = " `titel` varchar(15) DEFAULT NULL"
 Str(1, 3, 2) = " `vorname` varchar(15) DEFAULT NULL"
 Str(1, 3, 3) = " `nachname` varchar(15) DEFAULT NULL"
 Str(1, 3, 4) = " `tel` varchar(20) DEFAULT NULL"
 Str(1, 3, 5) = " `klinik` varchar(20) DEFAULT NULL"
 Str(1, 3, 6) = " `abteilung` varchar(50) DEFAULT NULL"
 Str(1, 3, 7) = " `email` varchar(50) DEFAULT NULL"
 Str(1, 3, 8) = " `strasse` varchar(20) DEFAULT NULL"
 Str(1, 3, 9) = " `ort` varchar(30) DEFAULT NULL"
 Str(1, 3, 10) = " `Nr` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 3, 11) = "  PRIMARY KEY (`Nr`)"
 Str(1, 3, 12) = " ENGINE=MyISAM AUTO_INCREMENT=13 DEFAULT CHARSET=latin1"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "efn"
 Str(0, 4, 1) = "`Nachname`"
 Str(0, 4, 2) = "`Vorname`"
 Str(0, 4, 3) = "`efn`"
 Str(0, 4, 4) = "`referent`"
 Str(0, 4, 5) = "`lernziel`"
 Str(0, 4, 6) = "`id`"
 Str(0, 4, 7) = "`id`"
 Str(0, 4, 8) = "`name`"
 ArtZ(0, 4) = 6
 ArtZ(1, 4) = 2
 Str(1, 4, 0) = "CREATE TABLE `efn` ("
 Str(1, 4, 1) = " `Nachname` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 2) = " `Vorname` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 3) = " `efn` decimal(15,0) DEFAULT NULL"
 Str(1, 4, 4) = " `referent` bit(1) NOT NULL DEFAULT b'0'"
 Str(1, 4, 5) = " `lernziel` tinyint(3) unsigned NOT NULL DEFAULT '0'"
 Str(1, 4, 6) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 4, 7) = "  PRIMARY KEY (`id`)"
 Str(1, 4, 8) = "  KEY `name` (`Nachname`,`Vorname`)"
 Str(1, 4, 9) = " ENGINE=MyISAM AUTO_INCREMENT=57 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "hae"
 Str(0, 5, 1) = "`DBNr`"
 Str(0, 5, 2) = "`diff`"
 Str(0, 5, 3) = "`BStelle`"
 Str(0, 5, 4) = "`Anrede`"
 Str(0, 5, 5) = "`HAName`"
 Str(0, 5, 6) = "`ort`"
 Str(0, 5, 7) = "`KVNR`"
 Str(0, 5, 8) = "`KVNu`"
 Str(0, 5, 9) = "`LANR`"
 Str(0, 5, 10) = "`Tel1`"
 Str(0, 5, 11) = "`Tel2`"
 Str(0, 5, 12) = "`Tel3`"
 Str(0, 5, 13) = "`Tel4`"
 Str(0, 5, 14) = "`Fax1`"
 Str(0, 5, 15) = "`Fax1k`"
 Str(0, 5, 16) = "`Fax2`"
 Str(0, 5, 17) = "`Fax2k`"
 Str(0, 5, 18) = "`Fax3`"
 Str(0, 5, 19) = "`Fax3k`"
 Str(0, 5, 20) = "`Email`"
 Str(0, 5, 21) = "`ZulG`"
 Str(0, 5, 22) = "`Arzttyp`"
 Str(0, 5, 23) = "`GemMit`"
 Str(0, 5, 24) = "`beme`"
 Str(0, 5, 25) = "`DMPT2`"
 Str(0, 5, 26) = "`DMPT1`"
 Str(0, 5, 27) = "`Geschlecht`"
 Str(0, 5, 28) = "`Titel`"
 Str(0, 5, 29) = "`Vorname`"
 Str(0, 5, 30) = "`Nachname`"
 Str(0, 5, 31) = "`Straße`"
 Str(0, 5, 32) = "`PLZ`"
 Str(0, 5, 33) = "`gelöscht`"
 Str(0, 5, 34) = "`seit`"
 Str(0, 5, 35) = "`AktZeit`"
 Str(0, 5, 36) = "`DBNr`"
 Str(0, 5, 37) = "`dbnr`"
 Str(0, 5, 38) = "`Fax1k`"
 Str(0, 5, 39) = "`KVNr`"
 Str(0, 5, 40) = "`Name`"
 Str(0, 5, 41) = "`kvnu`"
 ArtZ(0, 5) = 35
 ArtZ(1, 5) = 6
 Str(1, 5, 0) = "CREATE TABLE `hae` ("
 Str(1, 5, 1) = " `DBNr` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 5, 2) = " `diff` int(10) DEFAULT NULL"
 Str(1, 5, 3) = " `BStelle` longtext COLLATE latin1_german2_ci"
 Str(1, 5, 4) = " `Anrede` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 5) = " `HAName` longtext COLLATE latin1_german2_ci"
 Str(1, 5, 6) = " `ort` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 7) = " `KVNR` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 8) = " `KVNu` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 9) = " `LANR` int(9) unsigned NOT NULL"
 Str(1, 5, 10) = " `Tel1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 11) = " `Tel2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 12) = " `Tel3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 13) = " `Tel4` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 14) = " `Fax1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 15) = " `Fax1k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 16) = " `Fax2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 17) = " `Fax2k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 18) = " `Fax3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 19) = " `Fax3k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 20) = " `Email` longtext COLLATE latin1_german2_ci"
 Str(1, 5, 21) = " `ZulG` longtext COLLATE latin1_german2_ci"
 Str(1, 5, 22) = " `Arzttyp` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 23) = " `GemMit` longtext COLLATE latin1_german2_ci"
 Str(1, 5, 24) = " `beme` longtext COLLATE latin1_german2_ci"
 Str(1, 5, 25) = " `DMPT2` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 5, 26) = " `DMPT1` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 5, 27) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 28) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 29) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 30) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 31) = " `Straße` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 32) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 33) = " `gelöscht` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 5, 34) = " `seit` date NOT NULL"
 Str(1, 5, 35) = " `AktZeit` datetime DEFAULT NULL"
 Str(1, 5, 36) = "  PRIMARY KEY (`DBNr`)"
 Str(1, 5, 37) = "  UNIQUE KEY `dbnr` (`DBNr`)"
 Str(1, 5, 38) = "  KEY `Fax1k` (`Fax1k`)"
 Str(1, 5, 39) = "  KEY `KVNr` (`KVNR`,`Nachname`,`Vorname`)"
 Str(1, 5, 40) = "  KEY `Name` (`Nachname`,`Vorname`,`ort`)"
 Str(1, 5, 41) = "  KEY `kvnu` (`KVNu`)"
 Str(1, 5, 42) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "hae2"
 Str(0, 6, 1) = "`DBNr`"
 Str(0, 6, 2) = "`diff`"
 Str(0, 6, 3) = "`BStelle`"
 Str(0, 6, 4) = "`Anrede`"
 Str(0, 6, 5) = "`HAName`"
 Str(0, 6, 6) = "`ort`"
 Str(0, 6, 7) = "`KVNR`"
 Str(0, 6, 8) = "`KVNu`"
 Str(0, 6, 9) = "`LANR`"
 Str(0, 6, 10) = "`Tel1`"
 Str(0, 6, 11) = "`Tel2`"
 Str(0, 6, 12) = "`Tel3`"
 Str(0, 6, 13) = "`Tel4`"
 Str(0, 6, 14) = "`Fax1`"
 Str(0, 6, 15) = "`Fax1k`"
 Str(0, 6, 16) = "`Fax2`"
 Str(0, 6, 17) = "`Fax2k`"
 Str(0, 6, 18) = "`Fax3`"
 Str(0, 6, 19) = "`Fax3k`"
 Str(0, 6, 20) = "`Email`"
 Str(0, 6, 21) = "`ZulG`"
 Str(0, 6, 22) = "`Arzttyp`"
 Str(0, 6, 23) = "`GemMit`"
 Str(0, 6, 24) = "`beme`"
 Str(0, 6, 25) = "`DMPT2`"
 Str(0, 6, 26) = "`DMPT1`"
 Str(0, 6, 27) = "`Geschlecht`"
 Str(0, 6, 28) = "`Titel`"
 Str(0, 6, 29) = "`Vorname`"
 Str(0, 6, 30) = "`Nachname`"
 Str(0, 6, 31) = "`Straße`"
 Str(0, 6, 32) = "`PLZ`"
 Str(0, 6, 33) = "`gelöscht`"
 Str(0, 6, 34) = "`AktZeit`"
 Str(0, 6, 35) = "`DBNr`"
 Str(0, 6, 36) = "`dbnr`"
 Str(0, 6, 37) = "`Fax1k`"
 Str(0, 6, 38) = "`KVNr`"
 Str(0, 6, 39) = "`Name`"
 Str(0, 6, 40) = "`kvnu`"
 ArtZ(0, 6) = 34
 ArtZ(1, 6) = 6
 Str(1, 6, 0) = "CREATE TABLE `hae2` ("
 Str(1, 6, 1) = " `DBNr` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 6, 2) = " `diff` int(10) DEFAULT NULL"
 Str(1, 6, 3) = " `BStelle` longtext COLLATE latin1_german2_ci"
 Str(1, 6, 4) = " `Anrede` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 5) = " `HAName` longtext COLLATE latin1_german2_ci"
 Str(1, 6, 6) = " `ort` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 7) = " `KVNR` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 8) = " `KVNu` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 9) = " `LANR` int(9) unsigned NOT NULL"
 Str(1, 6, 10) = " `Tel1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 11) = " `Tel2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 12) = " `Tel3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 13) = " `Tel4` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 14) = " `Fax1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 15) = " `Fax1k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 16) = " `Fax2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 17) = " `Fax2k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 18) = " `Fax3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 19) = " `Fax3k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 20) = " `Email` longtext COLLATE latin1_german2_ci"
 Str(1, 6, 21) = " `ZulG` longtext COLLATE latin1_german2_ci"
 Str(1, 6, 22) = " `Arzttyp` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 23) = " `GemMit` longtext COLLATE latin1_german2_ci"
 Str(1, 6, 24) = " `beme` longtext COLLATE latin1_german2_ci"
 Str(1, 6, 25) = " `DMPT2` bit(1) DEFAULT NULL"
 Str(1, 6, 26) = " `DMPT1` bit(1) DEFAULT NULL"
 Str(1, 6, 27) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 28) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 29) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 30) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 31) = " `Straße` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 32) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 33) = " `gelöscht` bit(1) DEFAULT NULL"
 Str(1, 6, 34) = " `AktZeit` datetime DEFAULT NULL"
 Str(1, 6, 35) = "  PRIMARY KEY (`DBNr`)"
 Str(1, 6, 36) = "  UNIQUE KEY `dbnr` (`DBNr`)"
 Str(1, 6, 37) = "  KEY `Fax1k` (`Fax1k`)"
 Str(1, 6, 38) = "  KEY `KVNr` (`KVNR`,`Nachname`,`Vorname`)"
 Str(1, 6, 39) = "  KEY `Name` (`Nachname`,`Vorname`,`ort`)"
 Str(1, 6, 40) = "  KEY `kvnu` (`KVNu`)"
 Str(1, 6, 41) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "haeaktliste"
 Str(1, 7, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `haeaktliste` AS select cast(`h`.`AktZeit` as date) AS `date(aktzeit)`,count(0) AS `zahl`,`h`.`gelöscht` AS `gelöscht`,`h`.`DBNr` AS `DBNr` from `hae` `h` group by cast(`h`.`AktZeit` as date),`h`.`gelöscht` order by cast(`h`.`AktZeit` as date) desc,`h`.`gelöscht`,`h`.`DBNr`"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "haealt"
 Str(0, 8, 1) = "`DBNr`"
 Str(0, 8, 2) = "`diff`"
 Str(0, 8, 3) = "`BStelle`"
 Str(0, 8, 4) = "`Anrede`"
 Str(0, 8, 5) = "`HAName`"
 Str(0, 8, 6) = "`ort`"
 Str(0, 8, 7) = "`KVNR`"
 Str(0, 8, 8) = "`KVNu`"
 Str(0, 8, 9) = "`LANR`"
 Str(0, 8, 10) = "`Tel1`"
 Str(0, 8, 11) = "`Tel2`"
 Str(0, 8, 12) = "`Tel3`"
 Str(0, 8, 13) = "`Tel4`"
 Str(0, 8, 14) = "`Fax1`"
 Str(0, 8, 15) = "`Fax1k`"
 Str(0, 8, 16) = "`Fax2`"
 Str(0, 8, 17) = "`Fax2k`"
 Str(0, 8, 18) = "`Fax3`"
 Str(0, 8, 19) = "`Fax3k`"
 Str(0, 8, 20) = "`Email`"
 Str(0, 8, 21) = "`ZulG`"
 Str(0, 8, 22) = "`Arzttyp`"
 Str(0, 8, 23) = "`GemMit`"
 Str(0, 8, 24) = "`beme`"
 Str(0, 8, 25) = "`DMPT2`"
 Str(0, 8, 26) = "`DMPT1`"
 Str(0, 8, 27) = "`Geschlecht`"
 Str(0, 8, 28) = "`Titel`"
 Str(0, 8, 29) = "`Vorname`"
 Str(0, 8, 30) = "`Nachname`"
 Str(0, 8, 31) = "`Straße`"
 Str(0, 8, 32) = "`PLZ`"
 Str(0, 8, 33) = "`gelöscht`"
 Str(0, 8, 34) = "`seit`"
 Str(0, 8, 35) = "`bis`"
 Str(0, 8, 36) = "`AktZeit`"
 Str(0, 8, 37) = "`DBNr`"
 Str(0, 8, 38) = "`dbnr`"
 Str(0, 8, 39) = "`Fax1k`"
 Str(0, 8, 40) = "`KVNr`"
 Str(0, 8, 41) = "`Name`"
 Str(0, 8, 42) = "`kvnu`"
 ArtZ(0, 8) = 36
 ArtZ(1, 8) = 6
 Str(1, 8, 0) = "CREATE TABLE `haealt` ("
 Str(1, 8, 1) = " `DBNr` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 8, 2) = " `diff` int(10) DEFAULT NULL"
 Str(1, 8, 3) = " `BStelle` longtext COLLATE latin1_german2_ci"
 Str(1, 8, 4) = " `Anrede` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 5) = " `HAName` longtext COLLATE latin1_german2_ci"
 Str(1, 8, 6) = " `ort` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 7) = " `KVNR` varchar(15) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 8, 8) = " `KVNu` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 9) = " `LANR` int(9) unsigned NOT NULL"
 Str(1, 8, 10) = " `Tel1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 11) = " `Tel2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 12) = " `Tel3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 13) = " `Tel4` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 14) = " `Fax1` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 15) = " `Fax1k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 16) = " `Fax2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 17) = " `Fax2k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 18) = " `Fax3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 19) = " `Fax3k` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 20) = " `Email` longtext COLLATE latin1_german2_ci"
 Str(1, 8, 21) = " `ZulG` longtext COLLATE latin1_german2_ci"
 Str(1, 8, 22) = " `Arzttyp` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 23) = " `GemMit` varchar(426) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 24) = " `beme` longtext COLLATE latin1_german2_ci"
 Str(1, 8, 25) = " `DMPT2` bit(1) DEFAULT NULL"
 Str(1, 8, 26) = " `DMPT1` bit(1) DEFAULT NULL"
 Str(1, 8, 27) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 28) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 29) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 30) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 31) = " `Straße` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 32) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 33) = " `gelöscht` bit(1) DEFAULT NULL"
 Str(1, 8, 34) = " `seit` date NOT NULL"
 Str(1, 8, 35) = " `bis` date NOT NULL"
 Str(1, 8, 36) = " `AktZeit` datetime DEFAULT NULL"
 Str(1, 8, 37) = "  PRIMARY KEY (`DBNr`,`KVNR`)"
 Str(1, 8, 38) = "  UNIQUE KEY `dbnr` (`DBNr`,`KVNR`)"
 Str(1, 8, 39) = "  KEY `Fax1k` (`Fax1k`)"
 Str(1, 8, 40) = "  KEY `KVNr` (`KVNR`,`Nachname`,`Vorname`)"
 Str(1, 8, 41) = "  KEY `Name` (`Nachname`,`Vorname`,`ort`)"
 Str(1, 8, 42) = "  KEY `kvnu` (`KVNu`)"
 Str(1, 8, 43) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "haxls"
 Str(0, 9, 1) = "`Name`"
 Str(0, 9, 2) = "`Vorname`"
 Str(0, 9, 3) = "`KVNummer`"
 Str(0, 9, 4) = "`Tel`"
 Str(0, 9, 5) = "`Tel2`"
 Str(0, 9, 6) = "`Fax`"
 Str(0, 9, 7) = "`Straße`"
 Str(0, 9, 8) = "`Plz`"
 Str(0, 9, 9) = "`Ort`"
 Str(0, 9, 10) = "`Fachgruppe`"
 Str(0, 9, 11) = "`Titel`"
 Str(0, 9, 12) = "`name`"
 Str(0, 9, 13) = "`vorname`"
 ArtZ(0, 9) = 11
 ArtZ(1, 9) = 2
 Str(1, 9, 0) = "CREATE TABLE `haxls` ("
 Str(1, 9, 1) = " `Name` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 2) = " `Vorname` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 3) = " `KVNummer` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 4) = " `Tel` varchar(15) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 9, 5) = " `Tel2` varchar(20) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 9, 6) = " `Fax` varchar(20) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 9, 7) = " `Straße` varchar(30) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 9, 8) = " `Plz` varchar(5) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 9, 9) = " `Ort` varchar(25) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 9, 10) = " `Fachgruppe` varchar(255) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 9, 11) = " `Titel` varchar(40) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 9, 12) = "  KEY `name` (`Name`)"
 Str(1, 9, 13) = "  KEY `vorname` (`Vorname`)"
 Str(1, 9, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "stand"
 Str(0, 10, 1) = "`Stand`"
 Str(0, 10, 2) = "`Datum`"
 Str(0, 10, 3) = "`Stand`"
 Str(0, 10, 4) = "`Stand`"
 ArtZ(0, 10) = 2
 ArtZ(1, 10) = 2
 Str(1, 10, 0) = "CREATE TABLE `stand` ("
 Str(1, 10, 1) = " `Stand` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 10, 2) = " `Datum` datetime DEFAULT NULL"
 Str(1, 10, 3) = "  PRIMARY KEY (`Stand`)"
 Str(1, 10, 4) = "  UNIQUE KEY `Stand` (`Stand`)"
 Str(1, 10, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr10

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

Public Function doMach_kvaerzte(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 10
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
 For i = 0 To 10
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
  For i = 0 To 10
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
  MsgBox "Fertig mit doMach_kvaerzte(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_kvaerzte/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_kvaerzte

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
