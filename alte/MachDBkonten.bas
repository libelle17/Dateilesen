'Bauanleitung für eine Datenbank wie `//linux/konten` vom 15.11.09 14:26:20
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 7, 49) As new CString, ArtZ&(3, 7)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "2"
 Str(0, 0, 1) = "`ID`"
 Str(0, 0, 2) = "`eingid`"
 Str(0, 0, 3) = "`23#09#2005 00:07:17 72888 U:\Konto\apobank 22#9#05#txt`"
 Str(0, 0, 4) = "`ID`"
 Str(0, 0, 5) = "`fk_2`"
 Str(0, 0, 6) = "`fk_2`"
 ArtZ(0, 0) = 3
 ArtZ(1, 0) = 2
 ArtZ(2, 0) = 1
 Str(1, 0, 0) = "CREATE TABLE `2` ("
 Str(1, 0, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 0, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 0, 3) = " `23#09#2005 00:07:17 72888 U:\Konto\apobank 22#9#05#txt` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 4) = "  PRIMARY KEY (`ID`)"
 Str(1, 0, 5) = "  KEY `fk_2` (`eingid`)"
 Str(1, 0, 6) = "  CONSTRAINT `fk_2` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 0, 7) = " ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Date: SMDateiliste.txt'"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "BBBBBBBKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW"
 Str(0, 1, 1) = "`ID`"
 Str(0, 1, 2) = "`eingid`"
 Str(0, 1, 3) = "`Betrag`"
 Str(0, 1, 4) = "`Buchungstext`"
 Str(0, 1, 5) = "`Betrag - Währung`"
 Str(0, 1, 6) = "`Buchungstag`"
 Str(0, 1, 7) = "`Begünstigter/Absender - Bankleitzahl`"
 Str(0, 1, 8) = "`Begünstigter/Absender - Kontonummer`"
 Str(0, 1, 9) = "`Begünstigter/Absender - Name`"
 Str(0, 1, 10) = "`Kategorie`"
 Str(0, 1, 11) = "`Kostenstelle`"
 Str(0, 1, 12) = "`Laufende Nummer`"
 Str(0, 1, 13) = "`Marker`"
 Str(0, 1, 14) = "`Orginalbetrag`"
 Str(0, 1, 15) = "`Originalbetrag - Währung`"
 Str(0, 1, 16) = "`Primanota`"
 Str(0, 1, 17) = "`Saldo`"
 Str(0, 1, 18) = "`Saldo - Währung`"
 Str(0, 1, 19) = "`Storno`"
 Str(0, 1, 20) = "`Storno - Originalbetrag`"
 Str(0, 1, 21) = "`Splittbuchung - Auftraggeber / Name`"
 Str(0, 1, 22) = "`Splittbuchung - Kategorie`"
 Str(0, 1, 23) = "`Splittbuchung - Kostenstelle`"
 Str(0, 1, 24) = "`Splittbuchung - Orginal Betrag`"
 Str(0, 1, 25) = "`Splittbuchung - Unterkategorie`"
 Str(0, 1, 26) = "`Splittbuchung - Verwendungszweckzeile 1`"
 Str(0, 1, 27) = "`Textschlüssel`"
 Str(0, 1, 28) = "`Unterkategorie`"
 Str(0, 1, 29) = "`Verwendungszweckzeile 1`"
 Str(0, 1, 30) = "`Verwendungszweckzeile 2`"
 Str(0, 1, 31) = "`Verwendungszweckzeile 3`"
 Str(0, 1, 32) = "`Verwendungszweckzeile 4`"
 Str(0, 1, 33) = "`Verwendungszweckzeile 5`"
 Str(0, 1, 34) = "`Verwendungszweckzeile 6`"
 Str(0, 1, 35) = "`Verwendungszweckzeile 7`"
 Str(0, 1, 36) = "`Verwendungszweckzeile 8`"
 Str(0, 1, 37) = "`Verwendungszweckzeile 9`"
 Str(0, 1, 38) = "`Verwendungszweckzeile 10`"
 Str(0, 1, 39) = "`Verwendungszweckzeile 11`"
 Str(0, 1, 40) = "`Verwendungszweckzeile 12`"
 Str(0, 1, 41) = "`Verwendungszweckzeile 13`"
 Str(0, 1, 42) = "`Verwendungszweckzeile 14`"
 Str(0, 1, 43) = "`Wertstellungstag`"
 Str(0, 1, 44) = "`ID`"
 Str(0, 1, 45) = "`fk_BBBBBBBKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW`"
 Str(0, 1, 46) = "`fk_BBBBBBBKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW`"
 ArtZ(0, 1) = 43
 ArtZ(1, 1) = 2
 ArtZ(2, 1) = 1
 Str(1, 1, 0) = "CREATE TABLE `BBBBBBBKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW` ("
 Str(1, 1, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 1, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 1, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 1, 7) = " `Begünstigter/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 8) = " `Begünstigter/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 9) = " `Begünstigter/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 10) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 11) = " `Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 12) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 13) = " `Marker` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 14) = " `Orginalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 15) = " `Originalbetrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 16) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 17) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 18) = " `Saldo - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 19) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 20) = " `Storno - Originalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 21) = " `Splittbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 22) = " `Splittbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 23) = " `Splittbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 24) = " `Splittbuchung - Orginal Betrag` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 25) = " `Splittbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 26) = " `Splittbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 27) = " `Textschlüssel` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 28) = " `Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 29) = " `Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 30) = " `Verwendungszweckzeile 2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 31) = " `Verwendungszweckzeile 3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 32) = " `Verwendungszweckzeile 4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 33) = " `Verwendungszweckzeile 5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 34) = " `Verwendungszweckzeile 6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 35) = " `Verwendungszweckzeile 7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 36) = " `Verwendungszweckzeile 8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 37) = " `Verwendungszweckzeile 9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 38) = " `Verwendungszweckzeile 10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 39) = " `Verwendungszweckzeile 11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 40) = " `Verwendungszweckzeile 12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 41) = " `Verwendungszweckzeile 13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 42) = " `Verwendungszweckzeile 14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 43) = " `Wertstellungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 1, 44) = "  PRIMARY KEY (`ID`)"
 Str(1, 1, 45) = "  KEY `fk_BBBBBBBKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW` (`eingid`)"
 Str(1, 1, 46) = "  CONSTRAINT `fk_BBBBBBBKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 1, 47) = " ENGINE=InnoDB AUTO_INCREMENT=24631 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Date: 297626808_20080814.csv'"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "BBBBEEEKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW"
 Str(0, 2, 1) = "`ID`"
 Str(0, 2, 2) = "`eingid`"
 Str(0, 2, 3) = "`Betrag`"
 Str(0, 2, 4) = "`Buchungstext`"
 Str(0, 2, 5) = "`Betrag - Währung`"
 Str(0, 2, 6) = "`Buchungstag`"
 Str(0, 2, 7) = "`Empfänger/Absender - Bankleitzahl`"
 Str(0, 2, 8) = "`Empfänger/Absender - Kontonummer`"
 Str(0, 2, 9) = "`Empfänger/Absender - Name`"
 Str(0, 2, 10) = "`Kategorie`"
 Str(0, 2, 11) = "`Kostenstelle`"
 Str(0, 2, 12) = "`Laufende Nummer`"
 Str(0, 2, 13) = "`Marker`"
 Str(0, 2, 14) = "`Orginalbetrag`"
 Str(0, 2, 15) = "`Originalbetrag - Währung`"
 Str(0, 2, 16) = "`Primanota`"
 Str(0, 2, 17) = "`Saldo`"
 Str(0, 2, 18) = "`Saldo - Währung`"
 Str(0, 2, 19) = "`Storno`"
 Str(0, 2, 20) = "`Storno - Originalbetrag`"
 Str(0, 2, 21) = "`Splitbuchung - Auftraggeber / Name`"
 Str(0, 2, 22) = "`Splitbuchung - Kategorie`"
 Str(0, 2, 23) = "`Splitbuchung - Kostenstelle`"
 Str(0, 2, 24) = "`Splitbuchung - Orginal Betrag`"
 Str(0, 2, 25) = "`Splitbuchung - Unterkategorie`"
 Str(0, 2, 26) = "`Splitbuchung - Verwendungszweckzeile 1`"
 Str(0, 2, 27) = "`Textschlüssel`"
 Str(0, 2, 28) = "`Unterkategorie`"
 Str(0, 2, 29) = "`Verwendungszweckzeile 1`"
 Str(0, 2, 30) = "`Verwendungszweckzeile 2`"
 Str(0, 2, 31) = "`Verwendungszweckzeile 3`"
 Str(0, 2, 32) = "`Verwendungszweckzeile 4`"
 Str(0, 2, 33) = "`Verwendungszweckzeile 5`"
 Str(0, 2, 34) = "`Verwendungszweckzeile 6`"
 Str(0, 2, 35) = "`Verwendungszweckzeile 7`"
 Str(0, 2, 36) = "`Verwendungszweckzeile 8`"
 Str(0, 2, 37) = "`Verwendungszweckzeile 9`"
 Str(0, 2, 38) = "`Verwendungszweckzeile 10`"
 Str(0, 2, 39) = "`Verwendungszweckzeile 11`"
 Str(0, 2, 40) = "`Verwendungszweckzeile 12`"
 Str(0, 2, 41) = "`Verwendungszweckzeile 13`"
 Str(0, 2, 42) = "`Verwendungszweckzeile 14`"
 Str(0, 2, 43) = "`Wertstellungstag`"
 Str(0, 2, 44) = "`ID`"
 Str(0, 2, 45) = "`fk_BBBBEEEKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW`"
 Str(0, 2, 46) = "`fk_BBBBEEEKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW`"
 ArtZ(0, 2) = 43
 ArtZ(1, 2) = 2
 ArtZ(2, 2) = 1
 Str(1, 2, 0) = "CREATE TABLE `BBBBEEEKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW` ("
 Str(1, 2, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 2, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 2, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 2, 7) = " `Empfänger/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 8) = " `Empfänger/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 9) = " `Empfänger/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 10) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 11) = " `Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 12) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 13) = " `Marker` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 14) = " `Orginalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 15) = " `Originalbetrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 16) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 17) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 18) = " `Saldo - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 19) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 20) = " `Storno - Originalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 21) = " `Splitbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 22) = " `Splitbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 23) = " `Splitbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 24) = " `Splitbuchung - Orginal Betrag` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 25) = " `Splitbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 26) = " `Splitbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 27) = " `Textschlüssel` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 28) = " `Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 29) = " `Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 30) = " `Verwendungszweckzeile 2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 31) = " `Verwendungszweckzeile 3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 32) = " `Verwendungszweckzeile 4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 33) = " `Verwendungszweckzeile 5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 34) = " `Verwendungszweckzeile 6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 35) = " `Verwendungszweckzeile 7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 36) = " `Verwendungszweckzeile 8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 37) = " `Verwendungszweckzeile 9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 38) = " `Verwendungszweckzeile 10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 39) = " `Verwendungszweckzeile 11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 40) = " `Verwendungszweckzeile 12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 41) = " `Verwendungszweckzeile 13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 42) = " `Verwendungszweckzeile 14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 43) = " `Wertstellungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 2, 44) = "  PRIMARY KEY (`ID`)"
 Str(1, 2, 45) = "  KEY `fk_BBBBEEEKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW` (`eingid`)"
 Str(1, 2, 46) = "  CONSTRAINT `fk_BBBBEEEKKLMOOPSSSSSSSSSSTUVVVVVVVVVVVVVVW` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 2, 47) = " ENGINE=InnoDB AUTO_INCREMENT=278 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Date: pkon2006.csv'"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "D"
 Str(0, 3, 1) = "`ID`"
 Str(0, 3, 2) = "`eingid`"
 Str(0, 3, 3) = "`Datum, Zeit, Zeitzone, Name, Art, Status, WÃ¤hrung, Brutto, GebÃ`"
 Str(0, 3, 4) = "`ID`"
 Str(0, 3, 5) = "`fk_D`"
 Str(0, 3, 6) = "`fk_D`"
 ArtZ(0, 3) = 3
 ArtZ(1, 3) = 2
 ArtZ(2, 3) = 1
 Str(1, 3, 0) = "CREATE TABLE `D` ("
 Str(1, 3, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 3, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 3, 3) = " `Datum, Zeit, Zeitzone, Name, Art, Status, WÃ¤hrung, Brutto, GebÃ` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 3, 4) = "  PRIMARY KEY (`ID`)"
 Str(1, 3, 5) = "  KEY `fk_D` (`eingid`)"
 Str(1, 3, 6) = "  CONSTRAINT `fk_D` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 3, 7) = " ENGINE=InnoDB AUTO_INCREMENT=30 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Date: paypal20080413.csv'"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "LWWWAAAKKPZSSSSWWKKOODDGGrrSSAAADDDDDDDDDD"
 Str(0, 4, 1) = "`ID`"
 Str(0, 4, 2) = "`eingid`"
 Str(0, 4, 3) = "`Laufende Nummer`"
 Str(0, 4, 4) = "`Wertpapiername`"
 Str(0, 4, 5) = "`Wertpapierlangname`"
 Str(0, 4, 6) = "`Wertpapierkennung`"
 Str(0, 4, 7) = "`Aktueller Kurs`"
 Str(0, 4, 8) = "`Aktueller Kurs - Währung`"
 Str(0, 4, 9) = "`Aktueller Kurs - Börsenplatz`"
 Str(0, 4, 10) = "`Kaufkurs`"
 Str(0, 4, 11) = "`Kaufkurs - Währung`"
 Str(0, 4, 12) = "`Preisherkunft`"
 Str(0, 4, 13) = "`Zusatz der Preisherkunft`"
 Str(0, 4, 14) = "`Stück / Nominal`"
 Str(0, 4, 15) = "`Stück / Nominal - Währung`"
 Str(0, 4, 16) = "`Stückzinsen`"
 Str(0, 4, 17) = "`Stückzinsen - Währung`"
 Str(0, 4, 18) = "`Wert`"
 Str(0, 4, 19) = "`Wert - Währung`"
 Str(0, 4, 20) = "`Kaufpreis`"
 Str(0, 4, 21) = "`Kaufpreis - Währung`"
 Str(0, 4, 22) = "`Orderkosten`"
 Str(0, 4, 23) = "`Orderkosten - Währung`"
 Str(0, 4, 24) = "`Durchschittspreis`"
 Str(0, 4, 25) = "`Durchschittspreis - Währung`"
 Str(0, 4, 26) = "`Gewinn`"
 Str(0, 4, 27) = "`Gewinn - Währung`"
 Str(0, 4, 28) = "`relativer Gewinn`"
 Str(0, 4, 29) = "`relativer Gewinn - Währung`"
 Str(0, 4, 30) = "`Stückzinsen in Depotwährung`"
 Str(0, 4, 31) = "`Stückzinsen in Depotwährung - Währung`"
 Str(0, 4, 32) = "`Anzahl aufgelaufener Tage`"
 Str(0, 4, 33) = "`Aktueller Wert des gesamten Depots`"
 Str(0, 4, 34) = "`Aktueller Wert des gesamten Depots - Währung`"
 Str(0, 4, 35) = "`Depotunterposten - Stück / Nominal`"
 Str(0, 4, 36) = "`Depotunterposten - Stück / Nominal - Währung`"
 Str(0, 4, 37) = "`Depotunterposten - Verfügbar`"
 Str(0, 4, 38) = "`Depotunterposten - Lagerland`"
 Str(0, 4, 39) = "`Depotunterposten - Depotschlüssel`"
 Str(0, 4, 40) = "`Depotunterposten - Verwahrart`"
 Str(0, 4, 41) = "`Depotunterposten - Vertriebswegkennzeichen`"
 Str(0, 4, 42) = "`Depotunterposten - Lagerstelle`"
 Str(0, 4, 43) = "`Depotunterposten - Sperre bis`"
 Str(0, 4, 44) = "`Depotunterposten - Sperrvermerk / Zusatzvermerk des Kreditinstit`"
 Str(0, 4, 45) = "`ID`"
 Str(0, 4, 46) = "`fk_LWWWAAAKKPZSSSSWWKKOODDGGrrSSAAADDDDDDDDDD`"
 Str(0, 4, 47) = "`fk_LWWWAAAKKPZSSSSWWKKOODDGGrrSSAAADDDDDDDDDD`"
 ArtZ(0, 4) = 44
 ArtZ(1, 4) = 2
 ArtZ(2, 4) = 1
 Str(1, 4, 0) = "CREATE TABLE `LWWWAAAKKPZSSSSWWKKOODDGGrrSSAAADDDDDDDDDD` ("
 Str(1, 4, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 4, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 4, 3) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 4) = " `Wertpapiername` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 5) = " `Wertpapierlangname` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 6) = " `Wertpapierkennung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 7) = " `Aktueller Kurs` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 8) = " `Aktueller Kurs - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 9) = " `Aktueller Kurs - Börsenplatz` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 10) = " `Kaufkurs` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 11) = " `Kaufkurs - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 12) = " `Preisherkunft` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 13) = " `Zusatz der Preisherkunft` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 14) = " `Stück / Nominal` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 15) = " `Stück / Nominal - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 16) = " `Stückzinsen` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 17) = " `Stückzinsen - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 18) = " `Wert` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 19) = " `Wert - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 20) = " `Kaufpreis` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 21) = " `Kaufpreis - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 22) = " `Orderkosten` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 23) = " `Orderkosten - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 24) = " `Durchschittspreis` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 25) = " `Durchschittspreis - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 26) = " `Gewinn` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 27) = " `Gewinn - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 28) = " `relativer Gewinn` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 29) = " `relativer Gewinn - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 30) = " `Stückzinsen in Depotwährung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 31) = " `Stückzinsen in Depotwährung - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 32) = " `Anzahl aufgelaufener Tage` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 33) = " `Aktueller Wert des gesamten Depots` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 4, 34) = " `Aktueller Wert des gesamten Depots - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 35) = " `Depotunterposten - Stück / Nominal` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 36) = " `Depotunterposten - Stück / Nominal - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 37) = " `Depotunterposten - Verfügbar` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 38) = " `Depotunterposten - Lagerland` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 39) = " `Depotunterposten - Depotschlüssel` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 40) = " `Depotunterposten - Verwahrart` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 41) = " `Depotunterposten - Vertriebswegkennzeichen` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 42) = " `Depotunterposten - Lagerstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 43) = " `Depotunterposten - Sperre bis` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 44) = " `Depotunterposten - Sperrvermerk / Zusatzvermerk des Kreditinstit` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 45) = "  PRIMARY KEY (`ID`)"
 Str(1, 4, 46) = "  KEY `fk_LWWWAAAKKPZSSSSWWKKOODDGGrrSSAAADDDDDDDDDD` (`eingid`)"
 Str(1, 4, 47) = "  CONSTRAINT `fk_LWWWAAAKKPZSSSSWWKKOODDGGrrSSAAADDDDDDDDDD` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 4, 48) = " ENGINE=InnoDB AUTO_INCREMENT=219 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Date: 8655587700_20080919.csv'"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "O1FFF"
 Str(0, 5, 1) = "`ID`"
 Str(0, 5, 2) = "`eingid`"
 Str(0, 5, 3) = "`Obligo`"
 Str(0, 5, 4) = "`15#09#2008 02:41`"
 Str(0, 5, 5) = "`F3`"
 Str(0, 5, 6) = "`F4`"
 Str(0, 5, 7) = "`F5`"
 Str(0, 5, 8) = "`ID`"
 Str(0, 5, 9) = "`fk_O1FFF`"
 Str(0, 5, 10) = "`fk_O1FFF`"
 ArtZ(0, 5) = 7
 ArtZ(1, 5) = 2
 ArtZ(2, 5) = 1
 Str(1, 5, 0) = "CREATE TABLE `O1FFF` ("
 Str(1, 5, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 5, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 5, 3) = " `Obligo` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 4) = " `15#09#2008 02:41` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 5) = " `F3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 6) = " `F4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 7) = " `F5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 5, 9) = "  KEY `fk_O1FFF` (`eingid`)"
 Str(1, 5, 10) = "  CONSTRAINT `fk_O1FFF` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 5, 11) = " ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Date: obligo8655587700_20080915.csv'"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "SSAAASOBBOOTVVVVVVVVVVVVVVBWPKUK"
 Str(0, 6, 1) = "`ID`"
 Str(0, 6, 2) = "`eingid`"
 Str(0, 6, 3) = "`Saldo`"
 Str(0, 6, 4) = "`SdoWaehr`"
 Str(0, 6, 5) = "`AgBlz`"
 Str(0, 6, 6) = "`AgKto`"
 Str(0, 6, 7) = "`AgName1`"
 Str(0, 6, 8) = "`Storno`"
 Str(0, 6, 9) = "`OrigBtg`"
 Str(0, 6, 10) = "`Betrag`"
 Str(0, 6, 11) = "`BtgWaehr`"
 Str(0, 6, 12) = "`OCMTBetr`"
 Str(0, 6, 13) = "`OCMTWaehr`"
 Str(0, 6, 14) = "`Textschl`"
 Str(0, 6, 15) = "`VWZ1`"
 Str(0, 6, 16) = "`VWZ2`"
 Str(0, 6, 17) = "`VWZ3`"
 Str(0, 6, 18) = "`VWZ4`"
 Str(0, 6, 19) = "`VWZ5`"
 Str(0, 6, 20) = "`VWZ6`"
 Str(0, 6, 21) = "`VWZ7`"
 Str(0, 6, 22) = "`VWZ8`"
 Str(0, 6, 23) = "`VWZ9`"
 Str(0, 6, 24) = "`VWZ10`"
 Str(0, 6, 25) = "`VWZ11`"
 Str(0, 6, 26) = "`VWZ12`"
 Str(0, 6, 27) = "`VWZ13`"
 Str(0, 6, 28) = "`VWZ14`"
 Str(0, 6, 29) = "`BuchDatum`"
 Str(0, 6, 30) = "`WertDatum`"
 Str(0, 6, 31) = "`Primanota`"
 Str(0, 6, 32) = "`Kategorie`"
 Str(0, 6, 33) = "`Unterkat`"
 Str(0, 6, 34) = "`Kostenst`"
 Str(0, 6, 35) = "`ID`"
 Str(0, 6, 36) = "`fk_SSAAASOBBOOTVVVVVVVVVVVVVVBWPKUK`"
 Str(0, 6, 37) = "`fk_SSAAASOBBOOTVVVVVVVVVVVVVVBWPKUK`"
 ArtZ(0, 6) = 34
 ArtZ(1, 6) = 2
 ArtZ(2, 6) = 1
 Str(1, 6, 0) = "CREATE TABLE `SSAAASOBBOOTVVVVVVVVVVVVVVBWPKUK` ("
 Str(1, 6, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 6, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 6, 3) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 4) = " `SdoWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 5) = " `AgBlz` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 6) = " `AgKto` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 7) = " `AgName1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 8) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 9) = " `OrigBtg` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 10) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 11) = " `BtgWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 12) = " `OCMTBetr` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 13) = " `OCMTWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 14) = " `Textschl` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 15) = " `VWZ1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 16) = " `VWZ2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 17) = " `VWZ3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 18) = " `VWZ4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 19) = " `VWZ5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 20) = " `VWZ6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 21) = " `VWZ7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 22) = " `VWZ8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 23) = " `VWZ9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 24) = " `VWZ10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 25) = " `VWZ11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 26) = " `VWZ12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 27) = " `VWZ13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 28) = " `VWZ14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 29) = " `BuchDatum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 6, 30) = " `WertDatum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 6, 31) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 32) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 33) = " `Unterkat` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 34) = " `Kostenst` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 35) = "  PRIMARY KEY (`ID`)"
 Str(1, 6, 36) = "  KEY `fk_SSAAASOBBOOTVVVVVVVVVVVVVVBWPKUK` (`eingid`)"
 Str(1, 6, 37) = "  CONSTRAINT `fk_SSAAASOBBOOTVVVVVVVVVVVVVVBWPKUK` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 6, 38) = " ENGINE=InnoDB AUTO_INCREMENT=278 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Date: pb2005.txt'"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "eingelesen"
 Str(0, 7, 1) = "`eingid`"
 Str(0, 7, 2) = "`verzeich`"
 Str(0, 7, 3) = "`datei`"
 Str(0, 7, 4) = "`datum`"
 Str(0, 7, 5) = "`eingid`"
 Str(0, 7, 6) = "`vd`"
 ArtZ(0, 7) = 4
 ArtZ(1, 7) = 2
 Str(1, 7, 0) = "CREATE TABLE `eingelesen` ("
 Str(1, 7, 1) = " `eingid` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 7, 2) = " `verzeich` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 7, 3) = " `datei` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 7, 4) = " `datum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 7, 5) = "  PRIMARY KEY (`eingid`)"
 Str(1, 7, 6) = "  UNIQUE KEY `vd` (`verzeich`,`datei`)"
 Str(1, 7, 7) = " ENGINE=InnoDB AUTO_INCREMENT=27 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
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

Public Function doMach_konten(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
  MsgBox "Fertig mit doMach_konten(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_konten/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_konten

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
