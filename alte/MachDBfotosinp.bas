'Bauanleitung für eine Datenbank wie `//linux/fotosinp` vom 15.11.09 14:26:16
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 7, 41) As new CString, ArtZ&(3, 7)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "deff"
 Str(0, 0, 1) = "`ID`"
 Str(0, 0, 2) = "`NNa`"
 Str(0, 0, 3) = "`VNa`"
 Str(0, 0, 4) = "`Verz`"
 Str(0, 0, 5) = "`Datei`"
 Str(0, 0, 6) = "`USDat`"
 Str(0, 0, 7) = "`UArt`"
 Str(0, 0, 8) = "`Fold`"
 Str(0, 0, 9) = "`obEx`"
 Str(0, 0, 10) = "`FileDate`"
 Str(0, 0, 11) = "`FileLen`"
 Str(0, 0, 12) = "`ID`"
 Str(0, 0, 13) = "`Ident`"
 ArtZ(0, 0) = 11
 ArtZ(1, 0) = 2
 Str(1, 0, 0) = "CREATE TABLE `deff` ("
 Str(1, 0, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 0, 2) = " `NNa` varchar(30) NOT NULL"
 Str(1, 0, 3) = " `VNa` varchar(30) NOT NULL"
 Str(1, 0, 4) = " `Verz` varchar(255) NOT NULL COMMENT 'Unterverzeichnis ""IM""'"
 Str(1, 0, 5) = " `Datei` varchar(150) NOT NULL"
 Str(1, 0, 6) = " `USDat` datetime NOT NULL"
 Str(1, 0, 7) = " `UArt` varchar(50) NOT NULL"
 Str(1, 0, 8) = " `Fold` varchar(255) NOT NULL COMMENT 'Grundverzeichnis'"
 Str(1, 0, 9) = " `obEx` bit(1) NOT NULL COMMENT 'ob Datei existiert'"
 Str(1, 0, 10) = " `FileDate` datetime NOT NULL"
 Str(1, 0, 11) = " `FileLen` int(10) unsigned NOT NULL"
 Str(1, 0, 12) = "  PRIMARY KEY (`ID`)"
 Str(1, 0, 13) = "  KEY `Ident` (`NNa`,`VNa`,`USDat`,`UArt`)"
 Str(1, 0, 14) = " ENGINE=InnoDB AUTO_INCREMENT=497587 DEFAULT CHARSET=latin1"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "jpg"
 Str(0, 1, 1) = "`Name`"
 Str(0, 1, 2) = "`erstellt`"
 Str(0, 1, 3) = "`geändert`"
 Str(0, 1, 4) = "`Größe`"
 Str(0, 1, 5) = "`aktgröße`"
 Str(0, 1, 6) = "`Pfad`"
 Str(0, 1, 7) = "`pexist`"
 Str(0, 1, 8) = "`tPfad`"
 Str(0, 1, 9) = "`existiert`"
 Str(0, 1, 10) = "`größestimmt`"
 Str(0, 1, 11) = "`gelöscht`"
 Str(0, 1, 12) = "`bearbeitet`"
 Str(0, 1, 13) = "`verwendet`"
 Str(0, 1, 14) = "`NeuerName`"
 Str(0, 1, 15) = "`NNgelöscht`"
 Str(0, 1, 16) = "`WavPfad`"
 Str(0, 1, 17) = "`wavexist`"
 Str(0, 1, 18) = "`WavErstellt`"
 Str(0, 1, 19) = "`WavGröße`"
 Str(0, 1, 20) = "`WavGelöscht`"
 Str(0, 1, 21) = "`Pict`"
 Str(0, 1, 22) = "`Körperteil`"
 Str(0, 1, 23) = "`Beschreibung`"
 Str(0, 1, 24) = "`PatDatum`"
 Str(0, 1, 25) = "`WA`"
 Str(0, 1, 26) = "`Pat_id`"
 Str(0, 1, 27) = "`PatName`"
 Str(0, 1, 28) = "`Helligkeit`"
 Str(0, 1, 29) = "`Kontrast`"
 Str(0, 1, 30) = "`Gamma`"
 Str(0, 1, 31) = "`Pfad`"
 Str(0, 1, 32) = "`erstellt`"
 Str(0, 1, 33) = "`ErstGrö`"
 Str(0, 1, 34) = "`Name`"
 Str(0, 1, 35) = "`neuerName`"
 Str(0, 1, 36) = "`Pat_id`"
 Str(0, 1, 37) = "`pict`"
 Str(0, 1, 38) = "`Pict-Nummer`"
 Str(0, 1, 39) = "`WavErstGrö`"
 ArtZ(0, 1) = 30
 ArtZ(1, 1) = 9
 Str(1, 1, 0) = "CREATE TABLE `jpg` ("
 Str(1, 1, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 2) = " `erstellt` datetime DEFAULT NULL"
 Str(1, 1, 3) = " `geändert` datetime DEFAULT NULL"
 Str(1, 1, 4) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 1, 5) = " `aktgröße` int(10) unsigned NOT NULL"
 Str(1, 1, 6) = " `Pfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 7) = " `pexist` tinyint(1) unsigned NOT NULL"
 Str(1, 1, 8) = " `tPfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 9) = " `existiert` tinyint(1) unsigned NOT NULL"
 Str(1, 1, 10) = " `größestimmt` tinyint(1) unsigned NOT NULL"
 Str(1, 1, 11) = " `gelöscht` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 1, 12) = " `bearbeitet` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 1, 13) = " `verwendet` tinyint(1) NOT NULL DEFAULT '0' COMMENT 'wird für Turbomed verwendet (Patientenfoto)'"
 Str(1, 1, 14) = " `NeuerName` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 15) = " `NNgelöscht` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 1, 16) = " `WavPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 1, 17) = " `wavexist` tinyint(1) unsigned NOT NULL"
 Str(1, 1, 18) = " `WavErstellt` datetime DEFAULT NULL"
 Str(1, 1, 19) = " `WavGröße` int(10) DEFAULT NULL"
 Str(1, 1, 20) = " `WavGelöscht` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 1, 21) = " `Pict` smallint(5) DEFAULT NULL"
 Str(1, 1, 22) = " `Körperteil` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 23) = " `Beschreibung` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 24) = " `PatDatum` datetime DEFAULT NULL"
 Str(1, 1, 25) = " `WA` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 26) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 1, 27) = " `PatName` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 28) = " `Helligkeit` int(10) DEFAULT NULL"
 Str(1, 1, 29) = " `Kontrast` int(10) DEFAULT NULL"
 Str(1, 1, 30) = " `Gamma` int(10) DEFAULT NULL"
 Str(1, 1, 31) = "  UNIQUE KEY `Pfad` (`Pfad`)"
 Str(1, 1, 32) = "  KEY `erstellt` (`erstellt`)"
 Str(1, 1, 33) = "  KEY `ErstGrö` (`erstellt`,`Größe`)"
 Str(1, 1, 34) = "  KEY `Name` (`Name`)"
 Str(1, 1, 35) = "  KEY `neuerName` (`NeuerName`)"
 Str(1, 1, 36) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 1, 37) = "  KEY `pict` (`Pict`)"
 Str(1, 1, 38) = "  KEY `Pict-Nummer` (`Pict`)"
 Str(1, 1, 39) = "  KEY `WavErstGrö` (`WavErstellt`,`WavGröße`)"
 Str(1, 1, 40) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "jpg mit allen gelöschten"
 Str(0, 2, 1) = "`Name`"
 Str(0, 2, 2) = "`erstellt`"
 Str(0, 2, 3) = "`geändert`"
 Str(0, 2, 4) = "`Größe`"
 Str(0, 2, 5) = "`Pfad`"
 Str(0, 2, 6) = "`gelöscht`"
 Str(0, 2, 7) = "`bearbeitet`"
 Str(0, 2, 8) = "`verwendet`"
 Str(0, 2, 9) = "`NeuerName`"
 Str(0, 2, 10) = "`NNgelöscht`"
 Str(0, 2, 11) = "`WavPfad`"
 Str(0, 2, 12) = "`WavErstellt`"
 Str(0, 2, 13) = "`WavGröße`"
 Str(0, 2, 14) = "`WavGelöscht`"
 Str(0, 2, 15) = "`Pict`"
 Str(0, 2, 16) = "`Körperteil`"
 Str(0, 2, 17) = "`Beschreibung`"
 Str(0, 2, 18) = "`PatDatum`"
 Str(0, 2, 19) = "`WA`"
 Str(0, 2, 20) = "`Pat_id`"
 Str(0, 2, 21) = "`Pfad`"
 Str(0, 2, 22) = "`erstellt`"
 Str(0, 2, 23) = "`ErstGrö`"
 Str(0, 2, 24) = "`Name`"
 Str(0, 2, 25) = "`neuerName`"
 Str(0, 2, 26) = "`Pat_id`"
 Str(0, 2, 27) = "`pict`"
 Str(0, 2, 28) = "`Pict-Nummer`"
 Str(0, 2, 29) = "`WavErstGrö`"
 ArtZ(0, 2) = 20
 ArtZ(1, 2) = 9
 Str(1, 2, 0) = "CREATE TABLE `jpg mit allen gelöschten` ("
 Str(1, 2, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 2) = " `erstellt` datetime DEFAULT NULL"
 Str(1, 2, 3) = " `geändert` datetime DEFAULT NULL"
 Str(1, 2, 4) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 2, 5) = " `Pfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 6) = " `gelöscht` tinyint(1) NOT NULL"
 Str(1, 2, 7) = " `bearbeitet` tinyint(1) NOT NULL"
 Str(1, 2, 8) = " `verwendet` tinyint(1) NOT NULL"
 Str(1, 2, 9) = " `NeuerName` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 10) = " `NNgelöscht` tinyint(1) NOT NULL"
 Str(1, 2, 11) = " `WavPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 2, 12) = " `WavErstellt` datetime DEFAULT NULL"
 Str(1, 2, 13) = " `WavGröße` int(10) DEFAULT NULL"
 Str(1, 2, 14) = " `WavGelöscht` tinyint(1) NOT NULL"
 Str(1, 2, 15) = " `Pict` smallint(5) DEFAULT NULL"
 Str(1, 2, 16) = " `Körperteil` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 17) = " `Beschreibung` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 18) = " `PatDatum` datetime DEFAULT NULL"
 Str(1, 2, 19) = " `WA` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 20) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 2, 21) = "  UNIQUE KEY `Pfad` (`Pfad`)"
 Str(1, 2, 22) = "  KEY `erstellt` (`erstellt`)"
 Str(1, 2, 23) = "  KEY `ErstGrö` (`erstellt`,`Größe`)"
 Str(1, 2, 24) = "  KEY `Name` (`Name`)"
 Str(1, 2, 25) = "  KEY `neuerName` (`NeuerName`)"
 Str(1, 2, 26) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 2, 27) = "  KEY `pict` (`Pict`)"
 Str(1, 2, 28) = "  KEY `Pict-Nummer` (`Pict`)"
 Str(1, 2, 29) = "  KEY `WavErstGrö` (`WavErstellt`,`WavGröße`)"
 Str(1, 2, 30) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "jpg nach beschreibungsergänzung"
 Str(0, 3, 1) = "`Name`"
 Str(0, 3, 2) = "`erstellt`"
 Str(0, 3, 3) = "`geändert`"
 Str(0, 3, 4) = "`Größe`"
 Str(0, 3, 5) = "`Pfad`"
 Str(0, 3, 6) = "`gelöscht`"
 Str(0, 3, 7) = "`bearbeitet`"
 Str(0, 3, 8) = "`verwendet`"
 Str(0, 3, 9) = "`NeuerName`"
 Str(0, 3, 10) = "`NNgelöscht`"
 Str(0, 3, 11) = "`WavPfad`"
 Str(0, 3, 12) = "`WavErstellt`"
 Str(0, 3, 13) = "`WavGröße`"
 Str(0, 3, 14) = "`WavGelöscht`"
 Str(0, 3, 15) = "`Pict`"
 Str(0, 3, 16) = "`Körperteil`"
 Str(0, 3, 17) = "`Beschreibung`"
 Str(0, 3, 18) = "`PatDatum`"
 Str(0, 3, 19) = "`WA`"
 Str(0, 3, 20) = "`Pat_id`"
 Str(0, 3, 21) = "`PatName`"
 Str(0, 3, 22) = "`Pfad`"
 Str(0, 3, 23) = "`erstellt`"
 Str(0, 3, 24) = "`ErstGrö`"
 Str(0, 3, 25) = "`Name`"
 Str(0, 3, 26) = "`neuerName`"
 Str(0, 3, 27) = "`Pat_id`"
 Str(0, 3, 28) = "`pict`"
 Str(0, 3, 29) = "`Pict-Nummer`"
 Str(0, 3, 30) = "`WavErstGrö`"
 ArtZ(0, 3) = 21
 ArtZ(1, 3) = 9
 Str(1, 3, 0) = "CREATE TABLE `jpg nach beschreibungsergänzung` ("
 Str(1, 3, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 2) = " `erstellt` datetime DEFAULT NULL"
 Str(1, 3, 3) = " `geändert` datetime DEFAULT NULL"
 Str(1, 3, 4) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 3, 5) = " `Pfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 6) = " `gelöscht` tinyint(1) NOT NULL"
 Str(1, 3, 7) = " `bearbeitet` tinyint(1) NOT NULL"
 Str(1, 3, 8) = " `verwendet` tinyint(1) NOT NULL"
 Str(1, 3, 9) = " `NeuerName` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 10) = " `NNgelöscht` tinyint(1) NOT NULL"
 Str(1, 3, 11) = " `WavPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 3, 12) = " `WavErstellt` datetime DEFAULT NULL"
 Str(1, 3, 13) = " `WavGröße` int(10) DEFAULT NULL"
 Str(1, 3, 14) = " `WavGelöscht` tinyint(1) NOT NULL"
 Str(1, 3, 15) = " `Pict` smallint(5) DEFAULT NULL"
 Str(1, 3, 16) = " `Körperteil` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 17) = " `Beschreibung` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 18) = " `PatDatum` datetime DEFAULT NULL"
 Str(1, 3, 19) = " `WA` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 20) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 3, 21) = " `PatName` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 22) = "  UNIQUE KEY `Pfad` (`Pfad`)"
 Str(1, 3, 23) = "  KEY `erstellt` (`erstellt`)"
 Str(1, 3, 24) = "  KEY `ErstGrö` (`erstellt`,`Größe`)"
 Str(1, 3, 25) = "  KEY `Name` (`Name`)"
 Str(1, 3, 26) = "  KEY `neuerName` (`NeuerName`)"
 Str(1, 3, 27) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 3, 28) = "  KEY `pict` (`Pict`)"
 Str(1, 3, 29) = "  KEY `Pict-Nummer` (`Pict`)"
 Str(1, 3, 30) = "  KEY `WavErstGrö` (`WavErstellt`,`WavGröße`)"
 Str(1, 3, 31) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "kopie von jpg 03102006"
 Str(0, 4, 1) = "`Name`"
 Str(0, 4, 2) = "`erstellt`"
 Str(0, 4, 3) = "`geändert`"
 Str(0, 4, 4) = "`Größe`"
 Str(0, 4, 5) = "`Pfad`"
 Str(0, 4, 6) = "`gelöscht`"
 Str(0, 4, 7) = "`bearbeitet`"
 Str(0, 4, 8) = "`verwendet`"
 Str(0, 4, 9) = "`NeuerName`"
 Str(0, 4, 10) = "`NNgelöscht`"
 Str(0, 4, 11) = "`WavPfad`"
 Str(0, 4, 12) = "`WavErstellt`"
 Str(0, 4, 13) = "`WavGröße`"
 Str(0, 4, 14) = "`WavGelöscht`"
 Str(0, 4, 15) = "`Pict`"
 Str(0, 4, 16) = "`Körperteil`"
 Str(0, 4, 17) = "`Beschreibung`"
 Str(0, 4, 18) = "`PatDatum`"
 Str(0, 4, 19) = "`WA`"
 Str(0, 4, 20) = "`Pat_id`"
 Str(0, 4, 21) = "`PatName`"
 Str(0, 4, 22) = "`Pfad`"
 Str(0, 4, 23) = "`erstellt`"
 Str(0, 4, 24) = "`ErstGrö`"
 Str(0, 4, 25) = "`Name`"
 Str(0, 4, 26) = "`neuerName`"
 Str(0, 4, 27) = "`Pat_id`"
 Str(0, 4, 28) = "`pict`"
 Str(0, 4, 29) = "`Pict-Nummer`"
 Str(0, 4, 30) = "`WavErstGrö`"
 ArtZ(0, 4) = 21
 ArtZ(1, 4) = 9
 Str(1, 4, 0) = "CREATE TABLE `kopie von jpg 03102006` ("
 Str(1, 4, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 2) = " `erstellt` datetime DEFAULT NULL"
 Str(1, 4, 3) = " `geändert` datetime DEFAULT NULL"
 Str(1, 4, 4) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 4, 5) = " `Pfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 6) = " `gelöscht` tinyint(1) NOT NULL"
 Str(1, 4, 7) = " `bearbeitet` tinyint(1) NOT NULL"
 Str(1, 4, 8) = " `verwendet` tinyint(1) NOT NULL"
 Str(1, 4, 9) = " `NeuerName` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 10) = " `NNgelöscht` tinyint(1) NOT NULL"
 Str(1, 4, 11) = " `WavPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 4, 12) = " `WavErstellt` datetime DEFAULT NULL"
 Str(1, 4, 13) = " `WavGröße` int(10) DEFAULT NULL"
 Str(1, 4, 14) = " `WavGelöscht` tinyint(1) NOT NULL"
 Str(1, 4, 15) = " `Pict` smallint(5) DEFAULT NULL"
 Str(1, 4, 16) = " `Körperteil` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 17) = " `Beschreibung` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 18) = " `PatDatum` datetime DEFAULT NULL"
 Str(1, 4, 19) = " `WA` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 20) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 4, 21) = " `PatName` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 22) = "  UNIQUE KEY `Pfad` (`Pfad`)"
 Str(1, 4, 23) = "  KEY `erstellt` (`erstellt`)"
 Str(1, 4, 24) = "  KEY `ErstGrö` (`erstellt`,`Größe`)"
 Str(1, 4, 25) = "  KEY `Name` (`Name`)"
 Str(1, 4, 26) = "  KEY `neuerName` (`NeuerName`)"
 Str(1, 4, 27) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 4, 28) = "  KEY `pict` (`Pict`)"
 Str(1, 4, 29) = "  KEY `Pict-Nummer` (`Pict`)"
 Str(1, 4, 30) = "  KEY `WavErstGrö` (`WavErstellt`,`WavGröße`)"
 Str(1, 4, 31) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "kopie von jpg vor erstanwendung"
 Str(0, 5, 1) = "`Name`"
 Str(0, 5, 2) = "`erstellt`"
 Str(0, 5, 3) = "`geändert`"
 Str(0, 5, 4) = "`Größe`"
 Str(0, 5, 5) = "`Pfad`"
 Str(0, 5, 6) = "`gelöscht`"
 Str(0, 5, 7) = "`bearbeitet`"
 Str(0, 5, 8) = "`verwendet`"
 Str(0, 5, 9) = "`NeuerName`"
 Str(0, 5, 10) = "`NNgelöscht`"
 Str(0, 5, 11) = "`WavPfad`"
 Str(0, 5, 12) = "`WavErstellt`"
 Str(0, 5, 13) = "`WavGröße`"
 Str(0, 5, 14) = "`WavGelöscht`"
 Str(0, 5, 15) = "`Pict`"
 Str(0, 5, 16) = "`Körperteil`"
 Str(0, 5, 17) = "`Beschreibung`"
 Str(0, 5, 18) = "`PatDatum`"
 Str(0, 5, 19) = "`WA`"
 Str(0, 5, 20) = "`Pat_id`"
 Str(0, 5, 21) = "`PatName`"
 Str(0, 5, 22) = "`Pfad`"
 Str(0, 5, 23) = "`erstellt`"
 Str(0, 5, 24) = "`ErstGrö`"
 Str(0, 5, 25) = "`Name`"
 Str(0, 5, 26) = "`neuerName`"
 Str(0, 5, 27) = "`Pat_id`"
 Str(0, 5, 28) = "`pict`"
 Str(0, 5, 29) = "`Pict-Nummer`"
 Str(0, 5, 30) = "`WavErstGrö`"
 ArtZ(0, 5) = 21
 ArtZ(1, 5) = 9
 Str(1, 5, 0) = "CREATE TABLE `kopie von jpg vor erstanwendung` ("
 Str(1, 5, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 2) = " `erstellt` datetime DEFAULT NULL"
 Str(1, 5, 3) = " `geändert` datetime DEFAULT NULL"
 Str(1, 5, 4) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 5, 5) = " `Pfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 6) = " `gelöscht` tinyint(1) NOT NULL"
 Str(1, 5, 7) = " `bearbeitet` tinyint(1) NOT NULL"
 Str(1, 5, 8) = " `verwendet` tinyint(1) NOT NULL"
 Str(1, 5, 9) = " `NeuerName` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 10) = " `NNgelöscht` tinyint(1) NOT NULL"
 Str(1, 5, 11) = " `WavPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 5, 12) = " `WavErstellt` datetime DEFAULT NULL"
 Str(1, 5, 13) = " `WavGröße` int(10) DEFAULT NULL"
 Str(1, 5, 14) = " `WavGelöscht` tinyint(1) NOT NULL"
 Str(1, 5, 15) = " `Pict` smallint(5) DEFAULT NULL"
 Str(1, 5, 16) = " `Körperteil` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 17) = " `Beschreibung` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 18) = " `PatDatum` datetime DEFAULT NULL"
 Str(1, 5, 19) = " `WA` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 20) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 5, 21) = " `PatName` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 22) = "  UNIQUE KEY `Pfad` (`Pfad`)"
 Str(1, 5, 23) = "  KEY `erstellt` (`erstellt`)"
 Str(1, 5, 24) = "  KEY `ErstGrö` (`erstellt`,`Größe`)"
 Str(1, 5, 25) = "  KEY `Name` (`Name`)"
 Str(1, 5, 26) = "  KEY `neuerName` (`NeuerName`)"
 Str(1, 5, 27) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 5, 28) = "  KEY `pict` (`Pict`)"
 Str(1, 5, 29) = "  KEY `Pict-Nummer` (`Pict`)"
 Str(1, 5, 30) = "  KEY `WavErstGrö` (`WavErstellt`,`WavGröße`)"
 Str(1, 5, 31) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "orig"
 Str(0, 6, 1) = "`Name`"
 Str(0, 6, 2) = "`erstellt`"
 Str(0, 6, 3) = "`geändert`"
 Str(0, 6, 4) = "`Größe`"
 Str(0, 6, 5) = "`gelöscht`"
 Str(0, 6, 6) = "`NeuerName`"
 Str(0, 6, 7) = "`Pict`"
 Str(0, 6, 8) = "`erstellt`"
 Str(0, 6, 9) = "`ErstGrö`"
 Str(0, 6, 10) = "`Name`"
 Str(0, 6, 11) = "`Pict-Nummer`"
 ArtZ(0, 6) = 7
 ArtZ(1, 6) = 4
 Str(1, 6, 0) = "CREATE TABLE `orig` ("
 Str(1, 6, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 2) = " `erstellt` datetime DEFAULT NULL"
 Str(1, 6, 3) = " `geändert` datetime DEFAULT NULL"
 Str(1, 6, 4) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 6, 5) = " `gelöscht` tinyint(1) NOT NULL"
 Str(1, 6, 6) = " `NeuerName` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 7) = " `Pict` smallint(5) DEFAULT NULL"
 Str(1, 6, 8) = "  KEY `erstellt` (`erstellt`)"
 Str(1, 6, 9) = "  KEY `ErstGrö` (`erstellt`,`Größe`)"
 Str(1, 6, 10) = "  KEY `Name` (`Name`)"
 Str(1, 6, 11) = "  KEY `Pict-Nummer` (`Pict`)"
 Str(1, 6, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "wav"
 Str(0, 7, 1) = "`Name`"
 Str(0, 7, 2) = "`erstellt`"
 Str(0, 7, 3) = "`geändert`"
 Str(0, 7, 4) = "`Größe`"
 Str(0, 7, 5) = "`Pfad`"
 Str(0, 7, 6) = "`Pfad`"
 Str(0, 7, 7) = "`erstellt`"
 Str(0, 7, 8) = "`Name`"
 ArtZ(0, 7) = 5
 ArtZ(1, 7) = 3
 Str(1, 7, 0) = "CREATE TABLE `wav` ("
 Str(1, 7, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 2) = " `erstellt` datetime DEFAULT NULL"
 Str(1, 7, 3) = " `geändert` datetime DEFAULT NULL"
 Str(1, 7, 4) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 7, 5) = " `Pfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 6) = "  UNIQUE KEY `Pfad` (`Pfad`)"
 Str(1, 7, 7) = "  KEY `erstellt` (`erstellt`)"
 Str(1, 7, 8) = "  KEY `Name` (`Name`)"
 Str(1, 7, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr7

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

Public Function doMach_fotosinp(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 7
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
 For i = 0 To 7
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
  For i = 0 To 7
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
  MsgBox "Fertig mit doMach_fotosinp(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_fotosinp/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_fotosinp

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
