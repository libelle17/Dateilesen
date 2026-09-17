'Bauanleitung für eine Datenbank wie `//linux/depot` vom 15.11.09 14:25:58
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 10, 80) As new CString, ArtZ&(3, 10)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "T0006097316"
 Str(0, 0, 1) = "`MYID`"
 Str(0, 0, 2) = "`eingid`"
 Str(0, 0, 3) = "`Betrag`"
 Str(0, 0, 4) = "`Buchungstext`"
 Str(0, 0, 5) = "`Betrag - Währung`"
 Str(0, 0, 6) = "`Buchungstag`"
 Str(0, 0, 7) = "`Empfänger/Absender - Bankleitzahl`"
 Str(0, 0, 8) = "`Empfänger/Absender - Kontonummer`"
 Str(0, 0, 9) = "`Empfänger/Absender - Name`"
 Str(0, 0, 10) = "`Kategorie`"
 Str(0, 0, 11) = "`Kostenstelle`"
 Str(0, 0, 12) = "`Laufende Nummer`"
 Str(0, 0, 13) = "`Marker`"
 Str(0, 0, 14) = "`Orginalbetrag`"
 Str(0, 0, 15) = "`Originalbetrag - Währung`"
 Str(0, 0, 16) = "`Primanota`"
 Str(0, 0, 17) = "`Saldo`"
 Str(0, 0, 18) = "`Saldo - Währung`"
 Str(0, 0, 19) = "`Storno`"
 Str(0, 0, 20) = "`Storno - Originalbetrag`"
 Str(0, 0, 21) = "`Splittbuchung - Auftraggeber / Name`"
 Str(0, 0, 22) = "`Splittbuchung - Kategorie`"
 Str(0, 0, 23) = "`Splittbuchung - Kostenstelle`"
 Str(0, 0, 24) = "`Splittbuchung - Orginal Betrag`"
 Str(0, 0, 25) = "`Splittbuchung - Unterkategorie`"
 Str(0, 0, 26) = "`Splittbuchung - Verwendungszweckzeile 1`"
 Str(0, 0, 27) = "`Textschlüssel`"
 Str(0, 0, 28) = "`Unterkategorie`"
 Str(0, 0, 29) = "`Verwendungszweckzeile 1`"
 Str(0, 0, 30) = "`Verwendungszweckzeile 2`"
 Str(0, 0, 31) = "`Verwendungszweckzeile 3`"
 Str(0, 0, 32) = "`Verwendungszweckzeile 4`"
 Str(0, 0, 33) = "`Verwendungszweckzeile 5`"
 Str(0, 0, 34) = "`Verwendungszweckzeile 6`"
 Str(0, 0, 35) = "`Verwendungszweckzeile 7`"
 Str(0, 0, 36) = "`Verwendungszweckzeile 8`"
 Str(0, 0, 37) = "`Verwendungszweckzeile 9`"
 Str(0, 0, 38) = "`Verwendungszweckzeile 10`"
 Str(0, 0, 39) = "`Verwendungszweckzeile 11`"
 Str(0, 0, 40) = "`Verwendungszweckzeile 12`"
 Str(0, 0, 41) = "`Verwendungszweckzeile 13`"
 Str(0, 0, 42) = "`Verwendungszweckzeile 14`"
 Str(0, 0, 43) = "`Wertstellungstag`"
 Str(0, 0, 44) = "`ID`"
 Str(0, 0, 45) = "`Begünstigter/Absender - Bankleitzahl`"
 Str(0, 0, 46) = "`Begünstigter/Absender - Kontonummer`"
 Str(0, 0, 47) = "`Begünstigter/Absender - Name`"
 Str(0, 0, 48) = "`SdoWaehr`"
 Str(0, 0, 49) = "`AgBlz`"
 Str(0, 0, 50) = "`AgKto`"
 Str(0, 0, 51) = "`AgName1`"
 Str(0, 0, 52) = "`OrigBtg`"
 Str(0, 0, 53) = "`BtgWaehr`"
 Str(0, 0, 54) = "`OCMTBetr`"
 Str(0, 0, 55) = "`OCMTWaehr`"
 Str(0, 0, 56) = "`Textschl`"
 Str(0, 0, 57) = "`VWZ1`"
 Str(0, 0, 58) = "`VWZ2`"
 Str(0, 0, 59) = "`VWZ3`"
 Str(0, 0, 60) = "`VWZ4`"
 Str(0, 0, 61) = "`VWZ5`"
 Str(0, 0, 62) = "`VWZ6`"
 Str(0, 0, 63) = "`VWZ7`"
 Str(0, 0, 64) = "`VWZ8`"
 Str(0, 0, 65) = "`VWZ9`"
 Str(0, 0, 66) = "`VWZ10`"
 Str(0, 0, 67) = "`VWZ11`"
 Str(0, 0, 68) = "`VWZ12`"
 Str(0, 0, 69) = "`VWZ13`"
 Str(0, 0, 70) = "`VWZ14`"
 Str(0, 0, 71) = "`BuchDatum`"
 Str(0, 0, 72) = "`WertDatum`"
 Str(0, 0, 73) = "`Unterkat`"
 Str(0, 0, 74) = "`Kostenst`"
 Str(0, 0, 75) = "`BuchText`"
 Str(0, 0, 76) = "`MYID`"
 Str(0, 0, 77) = "`fk_T0006097316`"
 Str(0, 0, 78) = "`fk_T0006097316`"
 ArtZ(0, 0) = 75
 ArtZ(1, 0) = 2
 ArtZ(2, 0) = 1
 Str(1, 0, 0) = "CREATE TABLE `T0006097316` ("
 Str(1, 0, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 0, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 0, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 0, 7) = " `Empfänger/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 8) = " `Empfänger/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 9) = " `Empfänger/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 10) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 11) = " `Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 12) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 13) = " `Marker` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 14) = " `Orginalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 15) = " `Originalbetrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 16) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 17) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 18) = " `Saldo - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 19) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 20) = " `Storno - Originalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 21) = " `Splittbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 22) = " `Splittbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 23) = " `Splittbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 24) = " `Splittbuchung - Orginal Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 25) = " `Splittbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 26) = " `Splittbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 27) = " `Textschlüssel` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 28) = " `Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 29) = " `Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 30) = " `Verwendungszweckzeile 2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 31) = " `Verwendungszweckzeile 3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 32) = " `Verwendungszweckzeile 4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 33) = " `Verwendungszweckzeile 5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 34) = " `Verwendungszweckzeile 6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 35) = " `Verwendungszweckzeile 7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 36) = " `Verwendungszweckzeile 8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 37) = " `Verwendungszweckzeile 9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 38) = " `Verwendungszweckzeile 10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 39) = " `Verwendungszweckzeile 11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 40) = " `Verwendungszweckzeile 12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 41) = " `Verwendungszweckzeile 13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 42) = " `Verwendungszweckzeile 14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 43) = " `Wertstellungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 0, 44) = " `ID` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 45) = " `Begünstigter/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 46) = " `Begünstigter/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 47) = " `Begünstigter/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 48) = " `SdoWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 49) = " `AgBlz` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 50) = " `AgKto` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 51) = " `AgName1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 52) = " `OrigBtg` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 53) = " `BtgWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 54) = " `OCMTBetr` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 55) = " `OCMTWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 56) = " `Textschl` decimal(10,2) DEFAULT NULL"
 Str(1, 0, 57) = " `VWZ1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 58) = " `VWZ2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 59) = " `VWZ3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 60) = " `VWZ4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 61) = " `VWZ5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 62) = " `VWZ6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 63) = " `VWZ7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 64) = " `VWZ8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 65) = " `VWZ9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 66) = " `VWZ10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 67) = " `VWZ11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 68) = " `VWZ12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 69) = " `VWZ13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 70) = " `VWZ14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 71) = " `BuchDatum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 0, 72) = " `WertDatum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 0, 73) = " `Unterkat` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 74) = " `Kostenst` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 75) = " `BuchText` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 76) = "  PRIMARY KEY (`MYID`)"
 Str(1, 0, 77) = "  KEY `fk_T0006097316` (`eingid`)"
 Str(1, 0, 78) = "  CONSTRAINT `fk_T0006097316` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 0, 79) = " ENGINE=InnoDB AUTO_INCREMENT=11640 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 0006097316_20071014.csv'"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "T0086097316"
 Str(0, 1, 1) = "`MYID`"
 Str(0, 1, 2) = "`eingid`"
 Str(0, 1, 3) = "`Laufende Nummer`"
 Str(0, 1, 4) = "`Wertpapiername`"
 Str(0, 1, 5) = "`Wertpapierlangname`"
 Str(0, 1, 6) = "`Wertpapierkennung`"
 Str(0, 1, 7) = "`Aktueller Kurs`"
 Str(0, 1, 8) = "`Aktueller Kurs - Währung`"
 Str(0, 1, 9) = "`Aktueller Kurs - Börsenplatz`"
 Str(0, 1, 10) = "`Kaufkurs`"
 Str(0, 1, 11) = "`Kaufkurs - Währung`"
 Str(0, 1, 12) = "`Preisherkunft`"
 Str(0, 1, 13) = "`Zusatz der Preisherkunft`"
 Str(0, 1, 14) = "`Stück / Nominal`"
 Str(0, 1, 15) = "`Stück / Nominal - Währung`"
 Str(0, 1, 16) = "`Stückzinsen`"
 Str(0, 1, 17) = "`Stückzinsen - Währung`"
 Str(0, 1, 18) = "`Wert`"
 Str(0, 1, 19) = "`Wert - Währung`"
 Str(0, 1, 20) = "`Kaufpreis`"
 Str(0, 1, 21) = "`Kaufpreis - Währung`"
 Str(0, 1, 22) = "`Orderkosten`"
 Str(0, 1, 23) = "`Orderkosten - Währung`"
 Str(0, 1, 24) = "`Durchschittspreis`"
 Str(0, 1, 25) = "`Durchschittspreis - Währung`"
 Str(0, 1, 26) = "`Gewinn`"
 Str(0, 1, 27) = "`Gewinn - Währung`"
 Str(0, 1, 28) = "`relativer Gewinn`"
 Str(0, 1, 29) = "`relativer Gewinn - Währung`"
 Str(0, 1, 30) = "`Stückzinsen in Depotwährung`"
 Str(0, 1, 31) = "`Stückzinsen in Depotwährung - Währung`"
 Str(0, 1, 32) = "`Anzahl aufgelaufener Tage`"
 Str(0, 1, 33) = "`Aktueller Wert des gesamten Depots`"
 Str(0, 1, 34) = "`Aktueller Wert des gesamten Depots - Währung`"
 Str(0, 1, 35) = "`Depotunterposten - Stück / Nominal`"
 Str(0, 1, 36) = "`Depotunterposten - Stück / Nominal - Währung`"
 Str(0, 1, 37) = "`Depotunterposten - Verfügbar`"
 Str(0, 1, 38) = "`Depotunterposten - Lagerland`"
 Str(0, 1, 39) = "`Depotunterposten - Depotschlüssel`"
 Str(0, 1, 40) = "`Depotunterposten - Verwahrart`"
 Str(0, 1, 41) = "`Depotunterposten - Vertriebswegkennzeichen`"
 Str(0, 1, 42) = "`Depotunterposten - Lagerstelle`"
 Str(0, 1, 43) = "`Depotunterposten - Sperre bis`"
 Str(0, 1, 44) = "`Depotunterposten - Sperrvermerk / Zusatzvermerk des Kreditinstit`"
 Str(0, 1, 45) = "`MYID`"
 Str(0, 1, 46) = "`fk_T0086097316`"
 Str(0, 1, 47) = "`fk_T0086097316`"
 ArtZ(0, 1) = 44
 ArtZ(1, 1) = 2
 ArtZ(2, 1) = 1
 Str(1, 1, 0) = "CREATE TABLE `T0086097316` ("
 Str(1, 1, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 1, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 1, 3) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 4) = " `Wertpapiername` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 5) = " `Wertpapierlangname` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 6) = " `Wertpapierkennung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 7) = " `Aktueller Kurs` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 8) = " `Aktueller Kurs - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 9) = " `Aktueller Kurs - Börsenplatz` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 10) = " `Kaufkurs` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 11) = " `Kaufkurs - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 12) = " `Preisherkunft` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 13) = " `Zusatz der Preisherkunft` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 14) = " `Stück / Nominal` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 15) = " `Stück / Nominal - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 16) = " `Stückzinsen` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 17) = " `Stückzinsen - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 18) = " `Wert` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 19) = " `Wert - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 20) = " `Kaufpreis` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 21) = " `Kaufpreis - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 22) = " `Orderkosten` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 23) = " `Orderkosten - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 24) = " `Durchschittspreis` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 25) = " `Durchschittspreis - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 26) = " `Gewinn` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 27) = " `Gewinn - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 28) = " `relativer Gewinn` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 29) = " `relativer Gewinn - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 30) = " `Stückzinsen in Depotwährung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 31) = " `Stückzinsen in Depotwährung - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 32) = " `Anzahl aufgelaufener Tage` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 33) = " `Aktueller Wert des gesamten Depots` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 1, 34) = " `Aktueller Wert des gesamten Depots - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 35) = " `Depotunterposten - Stück / Nominal` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 36) = " `Depotunterposten - Stück / Nominal - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 37) = " `Depotunterposten - Verfügbar` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 38) = " `Depotunterposten - Lagerland` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 39) = " `Depotunterposten - Depotschlüssel` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 40) = " `Depotunterposten - Verwahrart` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 41) = " `Depotunterposten - Vertriebswegkennzeichen` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 42) = " `Depotunterposten - Lagerstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 43) = " `Depotunterposten - Sperre bis` decimal(10,2) DEFAULT NULL"
 Str(1, 1, 44) = " `Depotunterposten - Sperrvermerk / Zusatzvermerk des Kreditinstit` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 1, 45) = "  PRIMARY KEY (`MYID`)"
 Str(1, 1, 46) = "  KEY `fk_T0086097316` (`eingid`)"
 Str(1, 1, 47) = "  CONSTRAINT `fk_T0086097316` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 1, 48) = " ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 0086097316_20081107.csv'"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "T0106097316"
 Str(0, 2, 1) = "`MYID`"
 Str(0, 2, 2) = "`eingid`"
 Str(0, 2, 3) = "`Betrag`"
 Str(0, 2, 4) = "`Buchungstext`"
 Str(0, 2, 5) = "`Betrag - Währung`"
 Str(0, 2, 6) = "`Buchungstag`"
 Str(0, 2, 7) = "`Begünstigter/Absender - Bankleitzahl`"
 Str(0, 2, 8) = "`Begünstigter/Absender - Kontonummer`"
 Str(0, 2, 9) = "`Begünstigter/Absender - Name`"
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
 Str(0, 2, 21) = "`Splittbuchung - Auftraggeber / Name`"
 Str(0, 2, 22) = "`Splittbuchung - Kategorie`"
 Str(0, 2, 23) = "`Splittbuchung - Kostenstelle`"
 Str(0, 2, 24) = "`Splittbuchung - Orginal Betrag`"
 Str(0, 2, 25) = "`Splittbuchung - Unterkategorie`"
 Str(0, 2, 26) = "`Splittbuchung - Verwendungszweckzeile 1`"
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
 Str(0, 2, 44) = "`Empfänger/Absender - Bankleitzahl`"
 Str(0, 2, 45) = "`Empfänger/Absender - Kontonummer`"
 Str(0, 2, 46) = "`Empfänger/Absender - Name`"
 Str(0, 2, 47) = "`ID`"
 Str(0, 2, 48) = "`MYID`"
 Str(0, 2, 49) = "`fk_T0106097316`"
 Str(0, 2, 50) = "`fk_T0106097316`"
 ArtZ(0, 2) = 47
 ArtZ(1, 2) = 2
 ArtZ(2, 2) = 1
 Str(1, 2, 0) = "CREATE TABLE `T0106097316` ("
 Str(1, 2, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 2, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 2, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 2, 7) = " `Begünstigter/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 8) = " `Begünstigter/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 9) = " `Begünstigter/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
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
 Str(1, 2, 21) = " `Splittbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 22) = " `Splittbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 23) = " `Splittbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 24) = " `Splittbuchung - Orginal Betrag` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 25) = " `Splittbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 26) = " `Splittbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
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
 Str(1, 2, 44) = " `Empfänger/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 45) = " `Empfänger/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 46) = " `Empfänger/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 47) = " `ID` decimal(10,2) DEFAULT NULL"
 Str(1, 2, 48) = "  PRIMARY KEY (`MYID`)"
 Str(1, 2, 49) = "  KEY `fk_T0106097316` (`eingid`)"
 Str(1, 2, 50) = "  CONSTRAINT `fk_T0106097316` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 2, 51) = " ENGINE=InnoDB AUTO_INCREMENT=32 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 0106097316_20081107.csv'"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "T0655587766"
 Str(0, 3, 1) = "`MYID`"
 Str(0, 3, 2) = "`eingid`"
 Str(0, 3, 3) = "`Betrag`"
 Str(0, 3, 4) = "`Buchungstext`"
 Str(0, 3, 5) = "`Betrag - Währung`"
 Str(0, 3, 6) = "`Buchungstag`"
 Str(0, 3, 7) = "`Empfänger/Absender - Bankleitzahl`"
 Str(0, 3, 8) = "`Empfänger/Absender - Kontonummer`"
 Str(0, 3, 9) = "`Empfänger/Absender - Name`"
 Str(0, 3, 10) = "`Kategorie`"
 Str(0, 3, 11) = "`Kostenstelle`"
 Str(0, 3, 12) = "`Laufende Nummer`"
 Str(0, 3, 13) = "`Marker`"
 Str(0, 3, 14) = "`Orginalbetrag`"
 Str(0, 3, 15) = "`Originalbetrag - Währung`"
 Str(0, 3, 16) = "`Primanota`"
 Str(0, 3, 17) = "`Saldo`"
 Str(0, 3, 18) = "`Saldo - Währung`"
 Str(0, 3, 19) = "`Storno`"
 Str(0, 3, 20) = "`Storno - Originalbetrag`"
 Str(0, 3, 21) = "`Splittbuchung - Auftraggeber / Name`"
 Str(0, 3, 22) = "`Splittbuchung - Kategorie`"
 Str(0, 3, 23) = "`Splittbuchung - Kostenstelle`"
 Str(0, 3, 24) = "`Splittbuchung - Orginal Betrag`"
 Str(0, 3, 25) = "`Splittbuchung - Unterkategorie`"
 Str(0, 3, 26) = "`Splittbuchung - Verwendungszweckzeile 1`"
 Str(0, 3, 27) = "`Textschlüssel`"
 Str(0, 3, 28) = "`Unterkategorie`"
 Str(0, 3, 29) = "`Verwendungszweckzeile 1`"
 Str(0, 3, 30) = "`Verwendungszweckzeile 2`"
 Str(0, 3, 31) = "`Verwendungszweckzeile 3`"
 Str(0, 3, 32) = "`Verwendungszweckzeile 4`"
 Str(0, 3, 33) = "`Verwendungszweckzeile 5`"
 Str(0, 3, 34) = "`Verwendungszweckzeile 6`"
 Str(0, 3, 35) = "`Verwendungszweckzeile 7`"
 Str(0, 3, 36) = "`Verwendungszweckzeile 8`"
 Str(0, 3, 37) = "`Verwendungszweckzeile 9`"
 Str(0, 3, 38) = "`Verwendungszweckzeile 10`"
 Str(0, 3, 39) = "`Verwendungszweckzeile 11`"
 Str(0, 3, 40) = "`Verwendungszweckzeile 12`"
 Str(0, 3, 41) = "`Verwendungszweckzeile 13`"
 Str(0, 3, 42) = "`Verwendungszweckzeile 14`"
 Str(0, 3, 43) = "`Wertstellungstag`"
 Str(0, 3, 44) = "`ID`"
 Str(0, 3, 45) = "`Begünstigter/Absender - Bankleitzahl`"
 Str(0, 3, 46) = "`Begünstigter/Absender - Kontonummer`"
 Str(0, 3, 47) = "`Begünstigter/Absender - Name`"
 Str(0, 3, 48) = "`MYID`"
 Str(0, 3, 49) = "`fk_T0655587766`"
 Str(0, 3, 50) = "`fk_T0655587766`"
 ArtZ(0, 3) = 47
 ArtZ(1, 3) = 2
 ArtZ(2, 3) = 1
 Str(1, 3, 0) = "CREATE TABLE `T0655587766` ("
 Str(1, 3, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 3, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 3, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 3, 7) = " `Empfänger/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 8) = " `Empfänger/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 9) = " `Empfänger/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 10) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 11) = " `Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 12) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 13) = " `Marker` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 14) = " `Orginalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 15) = " `Originalbetrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 16) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 17) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 18) = " `Saldo - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 19) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 20) = " `Storno - Originalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 21) = " `Splittbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 22) = " `Splittbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 23) = " `Splittbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 24) = " `Splittbuchung - Orginal Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 25) = " `Splittbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 26) = " `Splittbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 27) = " `Textschlüssel` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 28) = " `Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 29) = " `Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 30) = " `Verwendungszweckzeile 2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 31) = " `Verwendungszweckzeile 3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 32) = " `Verwendungszweckzeile 4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 33) = " `Verwendungszweckzeile 5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 34) = " `Verwendungszweckzeile 6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 35) = " `Verwendungszweckzeile 7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 36) = " `Verwendungszweckzeile 8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 37) = " `Verwendungszweckzeile 9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 38) = " `Verwendungszweckzeile 10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 39) = " `Verwendungszweckzeile 11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 40) = " `Verwendungszweckzeile 12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 41) = " `Verwendungszweckzeile 13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 42) = " `Verwendungszweckzeile 14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 43) = " `Wertstellungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 3, 44) = " `ID` decimal(10,2) DEFAULT NULL"
 Str(1, 3, 45) = " `Begünstigter/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 46) = " `Begünstigter/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 47) = " `Begünstigter/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 3, 48) = "  PRIMARY KEY (`MYID`)"
 Str(1, 3, 49) = "  KEY `fk_T0655587766` (`eingid`)"
 Str(1, 3, 50) = "  CONSTRAINT `fk_T0655587766` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 3, 51) = " ENGINE=InnoDB AUTO_INCREMENT=157 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 0655587766_20071014.csv'"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "T1402064"
 Str(0, 4, 1) = "`MYID`"
 Str(0, 4, 2) = "`eingid`"
 Str(0, 4, 3) = "`Betrag`"
 Str(0, 4, 4) = "`Buchungstext`"
 Str(0, 4, 5) = "`Betrag - Währung`"
 Str(0, 4, 6) = "`Buchungstag`"
 Str(0, 4, 7) = "`Begünstigter/Absender - Bankleitzahl`"
 Str(0, 4, 8) = "`Begünstigter/Absender - Kontonummer`"
 Str(0, 4, 9) = "`Begünstigter/Absender - Name`"
 Str(0, 4, 10) = "`Kategorie`"
 Str(0, 4, 11) = "`Kostenstelle`"
 Str(0, 4, 12) = "`Laufende Nummer`"
 Str(0, 4, 13) = "`Marker`"
 Str(0, 4, 14) = "`Orginalbetrag`"
 Str(0, 4, 15) = "`Originalbetrag - Währung`"
 Str(0, 4, 16) = "`Primanota`"
 Str(0, 4, 17) = "`Saldo`"
 Str(0, 4, 18) = "`Saldo - Währung`"
 Str(0, 4, 19) = "`Storno`"
 Str(0, 4, 20) = "`Storno - Originalbetrag`"
 Str(0, 4, 21) = "`Splittbuchung - Auftraggeber / Name`"
 Str(0, 4, 22) = "`Splittbuchung - Kategorie`"
 Str(0, 4, 23) = "`Splittbuchung - Kostenstelle`"
 Str(0, 4, 24) = "`Splittbuchung - Orginal Betrag`"
 Str(0, 4, 25) = "`Splittbuchung - Unterkategorie`"
 Str(0, 4, 26) = "`Splittbuchung - Verwendungszweckzeile 1`"
 Str(0, 4, 27) = "`Textschlüssel`"
 Str(0, 4, 28) = "`Unterkategorie`"
 Str(0, 4, 29) = "`Verwendungszweckzeile 1`"
 Str(0, 4, 30) = "`Verwendungszweckzeile 2`"
 Str(0, 4, 31) = "`Verwendungszweckzeile 3`"
 Str(0, 4, 32) = "`Verwendungszweckzeile 4`"
 Str(0, 4, 33) = "`Verwendungszweckzeile 5`"
 Str(0, 4, 34) = "`Verwendungszweckzeile 6`"
 Str(0, 4, 35) = "`Verwendungszweckzeile 7`"
 Str(0, 4, 36) = "`Verwendungszweckzeile 8`"
 Str(0, 4, 37) = "`Verwendungszweckzeile 9`"
 Str(0, 4, 38) = "`Verwendungszweckzeile 10`"
 Str(0, 4, 39) = "`Verwendungszweckzeile 11`"
 Str(0, 4, 40) = "`Verwendungszweckzeile 12`"
 Str(0, 4, 41) = "`Verwendungszweckzeile 13`"
 Str(0, 4, 42) = "`Verwendungszweckzeile 14`"
 Str(0, 4, 43) = "`Wertstellungstag`"
 Str(0, 4, 44) = "`Empfänger/Absender - Bankleitzahl`"
 Str(0, 4, 45) = "`Empfänger/Absender - Kontonummer`"
 Str(0, 4, 46) = "`Empfänger/Absender - Name`"
 Str(0, 4, 47) = "`ID`"
 Str(0, 4, 48) = "`MYID`"
 Str(0, 4, 49) = "`fk_T1402064`"
 Str(0, 4, 50) = "`fk_T1402064`"
 ArtZ(0, 4) = 47
 ArtZ(1, 4) = 2
 ArtZ(2, 4) = 1
 Str(1, 4, 0) = "CREATE TABLE `T1402064` ("
 Str(1, 4, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 4, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 4, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 4, 7) = " `Begünstigter/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 8) = " `Begünstigter/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 9) = " `Begünstigter/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 10) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 11) = " `Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 12) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 13) = " `Marker` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 14) = " `Orginalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 15) = " `Originalbetrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 16) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 17) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 18) = " `Saldo - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 19) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 20) = " `Storno - Originalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 21) = " `Splittbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 22) = " `Splittbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 23) = " `Splittbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 24) = " `Splittbuchung - Orginal Betrag` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 25) = " `Splittbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 26) = " `Splittbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 27) = " `Textschlüssel` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 28) = " `Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 29) = " `Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 30) = " `Verwendungszweckzeile 2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 31) = " `Verwendungszweckzeile 3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 32) = " `Verwendungszweckzeile 4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 33) = " `Verwendungszweckzeile 5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 34) = " `Verwendungszweckzeile 6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 35) = " `Verwendungszweckzeile 7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 36) = " `Verwendungszweckzeile 8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 37) = " `Verwendungszweckzeile 9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 38) = " `Verwendungszweckzeile 10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 39) = " `Verwendungszweckzeile 11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 40) = " `Verwendungszweckzeile 12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 41) = " `Verwendungszweckzeile 13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 42) = " `Verwendungszweckzeile 14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 43) = " `Wertstellungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 4, 44) = " `Empfänger/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 45) = " `Empfänger/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 46) = " `Empfänger/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 47) = " `ID` decimal(10,2) DEFAULT NULL"
 Str(1, 4, 48) = "  PRIMARY KEY (`MYID`)"
 Str(1, 4, 49) = "  KEY `fk_T1402064` (`eingid`)"
 Str(1, 4, 50) = "  CONSTRAINT `fk_T1402064` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 4, 51) = " ENGINE=InnoDB AUTO_INCREMENT=261 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 1402064_20081107.csv'"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "T1506135621"
 Str(0, 5, 1) = "`MYID`"
 Str(0, 5, 2) = "`eingid`"
 Str(0, 5, 3) = "`Betrag`"
 Str(0, 5, 4) = "`Buchungstext`"
 Str(0, 5, 5) = "`Betrag - Währung`"
 Str(0, 5, 6) = "`Buchungstag`"
 Str(0, 5, 7) = "`Begünstigter/Absender - Bankleitzahl`"
 Str(0, 5, 8) = "`Begünstigter/Absender - Kontonummer`"
 Str(0, 5, 9) = "`Begünstigter/Absender - Name`"
 Str(0, 5, 10) = "`Kategorie`"
 Str(0, 5, 11) = "`Kostenstelle`"
 Str(0, 5, 12) = "`Laufende Nummer`"
 Str(0, 5, 13) = "`Marker`"
 Str(0, 5, 14) = "`Orginalbetrag`"
 Str(0, 5, 15) = "`Originalbetrag - Währung`"
 Str(0, 5, 16) = "`Primanota`"
 Str(0, 5, 17) = "`Saldo`"
 Str(0, 5, 18) = "`Saldo - Währung`"
 Str(0, 5, 19) = "`Storno`"
 Str(0, 5, 20) = "`Storno - Originalbetrag`"
 Str(0, 5, 21) = "`Splittbuchung - Auftraggeber / Name`"
 Str(0, 5, 22) = "`Splittbuchung - Kategorie`"
 Str(0, 5, 23) = "`Splittbuchung - Kostenstelle`"
 Str(0, 5, 24) = "`Splittbuchung - Orginal Betrag`"
 Str(0, 5, 25) = "`Splittbuchung - Unterkategorie`"
 Str(0, 5, 26) = "`Splittbuchung - Verwendungszweckzeile 1`"
 Str(0, 5, 27) = "`Textschlüssel`"
 Str(0, 5, 28) = "`Unterkategorie`"
 Str(0, 5, 29) = "`Verwendungszweckzeile 1`"
 Str(0, 5, 30) = "`Verwendungszweckzeile 2`"
 Str(0, 5, 31) = "`Verwendungszweckzeile 3`"
 Str(0, 5, 32) = "`Verwendungszweckzeile 4`"
 Str(0, 5, 33) = "`Verwendungszweckzeile 5`"
 Str(0, 5, 34) = "`Verwendungszweckzeile 6`"
 Str(0, 5, 35) = "`Verwendungszweckzeile 7`"
 Str(0, 5, 36) = "`Verwendungszweckzeile 8`"
 Str(0, 5, 37) = "`Verwendungszweckzeile 9`"
 Str(0, 5, 38) = "`Verwendungszweckzeile 10`"
 Str(0, 5, 39) = "`Verwendungszweckzeile 11`"
 Str(0, 5, 40) = "`Verwendungszweckzeile 12`"
 Str(0, 5, 41) = "`Verwendungszweckzeile 13`"
 Str(0, 5, 42) = "`Verwendungszweckzeile 14`"
 Str(0, 5, 43) = "`Wertstellungstag`"
 Str(0, 5, 44) = "`MYID`"
 Str(0, 5, 45) = "`fk_T1506135621`"
 Str(0, 5, 46) = "`fk_T1506135621`"
 ArtZ(0, 5) = 43
 ArtZ(1, 5) = 2
 ArtZ(2, 5) = 1
 Str(1, 5, 0) = "CREATE TABLE `T1506135621` ("
 Str(1, 5, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 5, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 5, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 5, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 5, 7) = " `Begünstigter/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 8) = " `Begünstigter/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 9) = " `Begünstigter/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 10) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 11) = " `Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 12) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 5, 13) = " `Marker` decimal(10,2) DEFAULT NULL"
 Str(1, 5, 14) = " `Orginalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 5, 15) = " `Originalbetrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 16) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 17) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 5, 18) = " `Saldo - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 19) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 5, 20) = " `Storno - Originalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 5, 21) = " `Splittbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 22) = " `Splittbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 23) = " `Splittbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 24) = " `Splittbuchung - Orginal Betrag` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 25) = " `Splittbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 26) = " `Splittbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 27) = " `Textschlüssel` decimal(10,2) DEFAULT NULL"
 Str(1, 5, 28) = " `Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 29) = " `Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 30) = " `Verwendungszweckzeile 2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 31) = " `Verwendungszweckzeile 3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 32) = " `Verwendungszweckzeile 4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 33) = " `Verwendungszweckzeile 5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 34) = " `Verwendungszweckzeile 6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 35) = " `Verwendungszweckzeile 7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 36) = " `Verwendungszweckzeile 8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 37) = " `Verwendungszweckzeile 9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 38) = " `Verwendungszweckzeile 10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 39) = " `Verwendungszweckzeile 11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 40) = " `Verwendungszweckzeile 12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 41) = " `Verwendungszweckzeile 13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 42) = " `Verwendungszweckzeile 14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 43) = " `Wertstellungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 5, 44) = "  PRIMARY KEY (`MYID`)"
 Str(1, 5, 45) = "  KEY `fk_T1506135621` (`eingid`)"
 Str(1, 5, 46) = "  CONSTRAINT `fk_T1506135621` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 5, 47) = " ENGINE=InnoDB AUTO_INCREMENT=24 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 1506135621_20081107.csv'"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "T297626808"
 Str(0, 6, 1) = "`MYID`"
 Str(0, 6, 2) = "`eingid`"
 Str(0, 6, 3) = "`Betrag`"
 Str(0, 6, 4) = "`Buchungstext`"
 Str(0, 6, 5) = "`Betrag - Währung`"
 Str(0, 6, 6) = "`Buchungstag`"
 Str(0, 6, 7) = "`Empfänger/Absender - Bankleitzahl`"
 Str(0, 6, 8) = "`Empfänger/Absender - Kontonummer`"
 Str(0, 6, 9) = "`Empfänger/Absender - Name`"
 Str(0, 6, 10) = "`Kategorie`"
 Str(0, 6, 11) = "`Kostenstelle`"
 Str(0, 6, 12) = "`Laufende Nummer`"
 Str(0, 6, 13) = "`Marker`"
 Str(0, 6, 14) = "`Orginalbetrag`"
 Str(0, 6, 15) = "`Originalbetrag - Währung`"
 Str(0, 6, 16) = "`Primanota`"
 Str(0, 6, 17) = "`Saldo`"
 Str(0, 6, 18) = "`Saldo - Währung`"
 Str(0, 6, 19) = "`Storno`"
 Str(0, 6, 20) = "`Storno - Originalbetrag`"
 Str(0, 6, 21) = "`Splittbuchung - Auftraggeber / Name`"
 Str(0, 6, 22) = "`Splittbuchung - Kategorie`"
 Str(0, 6, 23) = "`Splittbuchung - Kostenstelle`"
 Str(0, 6, 24) = "`Splittbuchung - Orginal Betrag`"
 Str(0, 6, 25) = "`Splittbuchung - Unterkategorie`"
 Str(0, 6, 26) = "`Splittbuchung - Verwendungszweckzeile 1`"
 Str(0, 6, 27) = "`Textschlüssel`"
 Str(0, 6, 28) = "`Unterkategorie`"
 Str(0, 6, 29) = "`Verwendungszweckzeile 1`"
 Str(0, 6, 30) = "`Verwendungszweckzeile 2`"
 Str(0, 6, 31) = "`Verwendungszweckzeile 3`"
 Str(0, 6, 32) = "`Verwendungszweckzeile 4`"
 Str(0, 6, 33) = "`Verwendungszweckzeile 5`"
 Str(0, 6, 34) = "`Verwendungszweckzeile 6`"
 Str(0, 6, 35) = "`Verwendungszweckzeile 7`"
 Str(0, 6, 36) = "`Verwendungszweckzeile 8`"
 Str(0, 6, 37) = "`Verwendungszweckzeile 9`"
 Str(0, 6, 38) = "`Verwendungszweckzeile 10`"
 Str(0, 6, 39) = "`Verwendungszweckzeile 11`"
 Str(0, 6, 40) = "`Verwendungszweckzeile 12`"
 Str(0, 6, 41) = "`Verwendungszweckzeile 13`"
 Str(0, 6, 42) = "`Verwendungszweckzeile 14`"
 Str(0, 6, 43) = "`Wertstellungstag`"
 Str(0, 6, 44) = "`ID`"
 Str(0, 6, 45) = "`Begünstigter/Absender - Bankleitzahl`"
 Str(0, 6, 46) = "`Begünstigter/Absender - Kontonummer`"
 Str(0, 6, 47) = "`Begünstigter/Absender - Name`"
 Str(0, 6, 48) = "`SdoWaehr`"
 Str(0, 6, 49) = "`AgBlz`"
 Str(0, 6, 50) = "`AgKto`"
 Str(0, 6, 51) = "`AgName1`"
 Str(0, 6, 52) = "`OrigBtg`"
 Str(0, 6, 53) = "`BtgWaehr`"
 Str(0, 6, 54) = "`OCMTBetr`"
 Str(0, 6, 55) = "`OCMTWaehr`"
 Str(0, 6, 56) = "`Textschl`"
 Str(0, 6, 57) = "`VWZ1`"
 Str(0, 6, 58) = "`VWZ2`"
 Str(0, 6, 59) = "`VWZ3`"
 Str(0, 6, 60) = "`VWZ4`"
 Str(0, 6, 61) = "`VWZ5`"
 Str(0, 6, 62) = "`VWZ6`"
 Str(0, 6, 63) = "`VWZ7`"
 Str(0, 6, 64) = "`VWZ8`"
 Str(0, 6, 65) = "`VWZ9`"
 Str(0, 6, 66) = "`VWZ10`"
 Str(0, 6, 67) = "`VWZ11`"
 Str(0, 6, 68) = "`VWZ12`"
 Str(0, 6, 69) = "`VWZ13`"
 Str(0, 6, 70) = "`VWZ14`"
 Str(0, 6, 71) = "`BuchDatum`"
 Str(0, 6, 72) = "`WertDatum`"
 Str(0, 6, 73) = "`Unterkat`"
 Str(0, 6, 74) = "`Kostenst`"
 Str(0, 6, 75) = "`BuchText`"
 Str(0, 6, 76) = "`MYID`"
 Str(0, 6, 77) = "`fk_T297626808`"
 Str(0, 6, 78) = "`fk_T297626808`"
 ArtZ(0, 6) = 75
 ArtZ(1, 6) = 2
 ArtZ(2, 6) = 1
 Str(1, 6, 0) = "CREATE TABLE `T297626808` ("
 Str(1, 6, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 6, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 6, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 6, 7) = " `Empfänger/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 8) = " `Empfänger/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 9) = " `Empfänger/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 10) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 11) = " `Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 12) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 13) = " `Marker` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 14) = " `Orginalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 15) = " `Originalbetrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 16) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 17) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 18) = " `Saldo - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 19) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 20) = " `Storno - Originalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 21) = " `Splittbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 22) = " `Splittbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 23) = " `Splittbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 24) = " `Splittbuchung - Orginal Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 25) = " `Splittbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 26) = " `Splittbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 27) = " `Textschlüssel` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 28) = " `Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 29) = " `Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 30) = " `Verwendungszweckzeile 2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 31) = " `Verwendungszweckzeile 3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 32) = " `Verwendungszweckzeile 4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 33) = " `Verwendungszweckzeile 5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 34) = " `Verwendungszweckzeile 6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 35) = " `Verwendungszweckzeile 7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 36) = " `Verwendungszweckzeile 8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 37) = " `Verwendungszweckzeile 9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 38) = " `Verwendungszweckzeile 10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 39) = " `Verwendungszweckzeile 11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 40) = " `Verwendungszweckzeile 12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 41) = " `Verwendungszweckzeile 13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 42) = " `Verwendungszweckzeile 14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 43) = " `Wertstellungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 6, 44) = " `ID` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 45) = " `Begünstigter/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 46) = " `Begünstigter/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 47) = " `Begünstigter/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 48) = " `SdoWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 49) = " `AgBlz` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 50) = " `AgKto` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 51) = " `AgName1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 52) = " `OrigBtg` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 53) = " `BtgWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 54) = " `OCMTBetr` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 55) = " `OCMTWaehr` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 56) = " `Textschl` decimal(10,2) DEFAULT NULL"
 Str(1, 6, 57) = " `VWZ1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 58) = " `VWZ2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 59) = " `VWZ3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 60) = " `VWZ4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 61) = " `VWZ5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 62) = " `VWZ6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 63) = " `VWZ7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 64) = " `VWZ8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 65) = " `VWZ9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 66) = " `VWZ10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 67) = " `VWZ11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 68) = " `VWZ12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 69) = " `VWZ13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 70) = " `VWZ14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 71) = " `BuchDatum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 6, 72) = " `WertDatum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 6, 73) = " `Unterkat` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 74) = " `Kostenst` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 75) = " `BuchText` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 6, 76) = "  PRIMARY KEY (`MYID`)"
 Str(1, 6, 77) = "  KEY `fk_T297626808` (`eingid`)"
 Str(1, 6, 78) = "  CONSTRAINT `fk_T297626808` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 6, 79) = " ENGINE=InnoDB AUTO_INCREMENT=10024 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 297626808_20080220 für steuererklärungen '"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "T6097316"
 Str(0, 7, 1) = "`MYID`"
 Str(0, 7, 2) = "`eingid`"
 Str(0, 7, 3) = "`Betrag`"
 Str(0, 7, 4) = "`Buchungstext`"
 Str(0, 7, 5) = "`Betrag - Währung`"
 Str(0, 7, 6) = "`Buchungstag`"
 Str(0, 7, 7) = "`Begünstigter/Absender - Bankleitzahl`"
 Str(0, 7, 8) = "`Begünstigter/Absender - Kontonummer`"
 Str(0, 7, 9) = "`Begünstigter/Absender - Name`"
 Str(0, 7, 10) = "`Kategorie`"
 Str(0, 7, 11) = "`Kostenstelle`"
 Str(0, 7, 12) = "`Laufende Nummer`"
 Str(0, 7, 13) = "`Marker`"
 Str(0, 7, 14) = "`Orginalbetrag`"
 Str(0, 7, 15) = "`Originalbetrag - Währung`"
 Str(0, 7, 16) = "`Primanota`"
 Str(0, 7, 17) = "`Saldo`"
 Str(0, 7, 18) = "`Saldo - Währung`"
 Str(0, 7, 19) = "`Storno`"
 Str(0, 7, 20) = "`Storno - Originalbetrag`"
 Str(0, 7, 21) = "`Splittbuchung - Auftraggeber / Name`"
 Str(0, 7, 22) = "`Splittbuchung - Kategorie`"
 Str(0, 7, 23) = "`Splittbuchung - Kostenstelle`"
 Str(0, 7, 24) = "`Splittbuchung - Orginal Betrag`"
 Str(0, 7, 25) = "`Splittbuchung - Unterkategorie`"
 Str(0, 7, 26) = "`Splittbuchung - Verwendungszweckzeile 1`"
 Str(0, 7, 27) = "`Textschlüssel`"
 Str(0, 7, 28) = "`Unterkategorie`"
 Str(0, 7, 29) = "`Verwendungszweckzeile 1`"
 Str(0, 7, 30) = "`Verwendungszweckzeile 2`"
 Str(0, 7, 31) = "`Verwendungszweckzeile 3`"
 Str(0, 7, 32) = "`Verwendungszweckzeile 4`"
 Str(0, 7, 33) = "`Verwendungszweckzeile 5`"
 Str(0, 7, 34) = "`Verwendungszweckzeile 6`"
 Str(0, 7, 35) = "`Verwendungszweckzeile 7`"
 Str(0, 7, 36) = "`Verwendungszweckzeile 8`"
 Str(0, 7, 37) = "`Verwendungszweckzeile 9`"
 Str(0, 7, 38) = "`Verwendungszweckzeile 10`"
 Str(0, 7, 39) = "`Verwendungszweckzeile 11`"
 Str(0, 7, 40) = "`Verwendungszweckzeile 12`"
 Str(0, 7, 41) = "`Verwendungszweckzeile 13`"
 Str(0, 7, 42) = "`Verwendungszweckzeile 14`"
 Str(0, 7, 43) = "`Wertstellungstag`"
 Str(0, 7, 44) = "`MYID`"
 Str(0, 7, 45) = "`fk_T6097316`"
 Str(0, 7, 46) = "`fk_T6097316`"
 ArtZ(0, 7) = 43
 ArtZ(1, 7) = 2
 ArtZ(2, 7) = 1
 Str(1, 7, 0) = "CREATE TABLE `T6097316` ("
 Str(1, 7, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 7, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 7, 3) = " `Betrag` decimal(10,2) DEFAULT NULL"
 Str(1, 7, 4) = " `Buchungstext` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 5) = " `Betrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 6) = " `Buchungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 7, 7) = " `Begünstigter/Absender - Bankleitzahl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 8) = " `Begünstigter/Absender - Kontonummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 9) = " `Begünstigter/Absender - Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 10) = " `Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 11) = " `Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 12) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 7, 13) = " `Marker` decimal(10,2) DEFAULT NULL"
 Str(1, 7, 14) = " `Orginalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 7, 15) = " `Originalbetrag - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 16) = " `Primanota` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 17) = " `Saldo` decimal(10,2) DEFAULT NULL"
 Str(1, 7, 18) = " `Saldo - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 19) = " `Storno` decimal(10,2) DEFAULT NULL"
 Str(1, 7, 20) = " `Storno - Originalbetrag` decimal(10,2) DEFAULT NULL"
 Str(1, 7, 21) = " `Splittbuchung - Auftraggeber / Name` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 22) = " `Splittbuchung - Kategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 23) = " `Splittbuchung - Kostenstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 24) = " `Splittbuchung - Orginal Betrag` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 25) = " `Splittbuchung - Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 26) = " `Splittbuchung - Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 27) = " `Textschlüssel` decimal(10,2) DEFAULT NULL"
 Str(1, 7, 28) = " `Unterkategorie` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 29) = " `Verwendungszweckzeile 1` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 30) = " `Verwendungszweckzeile 2` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 31) = " `Verwendungszweckzeile 3` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 32) = " `Verwendungszweckzeile 4` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 33) = " `Verwendungszweckzeile 5` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 34) = " `Verwendungszweckzeile 6` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 35) = " `Verwendungszweckzeile 7` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 36) = " `Verwendungszweckzeile 8` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 37) = " `Verwendungszweckzeile 9` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 38) = " `Verwendungszweckzeile 10` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 39) = " `Verwendungszweckzeile 11` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 40) = " `Verwendungszweckzeile 12` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 41) = " `Verwendungszweckzeile 13` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 42) = " `Verwendungszweckzeile 14` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 7, 43) = " `Wertstellungstag` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 7, 44) = "  PRIMARY KEY (`MYID`)"
 Str(1, 7, 45) = "  KEY `fk_T6097316` (`eingid`)"
 Str(1, 7, 46) = "  CONSTRAINT `fk_T6097316` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 7, 47) = " ENGINE=InnoDB AUTO_INCREMENT=1195 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 6097316_20070615.csv'"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "T6944009"
 Str(0, 8, 1) = "`MYID`"
 Str(0, 8, 2) = "`eingid`"
 Str(0, 8, 3) = "`Laufende Nummer`"
 Str(0, 8, 4) = "`Wertpapiername`"
 Str(0, 8, 5) = "`Wertpapierlangname`"
 Str(0, 8, 6) = "`Wertpapierkennung`"
 Str(0, 8, 7) = "`Aktueller Kurs`"
 Str(0, 8, 8) = "`Aktueller Kurs - Währung`"
 Str(0, 8, 9) = "`Aktueller Kurs - Börsenplatz`"
 Str(0, 8, 10) = "`Kaufkurs`"
 Str(0, 8, 11) = "`Kaufkurs - Währung`"
 Str(0, 8, 12) = "`Preisherkunft`"
 Str(0, 8, 13) = "`Zusatz der Preisherkunft`"
 Str(0, 8, 14) = "`Stück / Nominal`"
 Str(0, 8, 15) = "`Stück / Nominal - Währung`"
 Str(0, 8, 16) = "`Stückzinsen`"
 Str(0, 8, 17) = "`Stückzinsen - Währung`"
 Str(0, 8, 18) = "`Wert`"
 Str(0, 8, 19) = "`Wert - Währung`"
 Str(0, 8, 20) = "`Kaufpreis`"
 Str(0, 8, 21) = "`Kaufpreis - Währung`"
 Str(0, 8, 22) = "`Orderkosten`"
 Str(0, 8, 23) = "`Orderkosten - Währung`"
 Str(0, 8, 24) = "`Durchschittspreis`"
 Str(0, 8, 25) = "`Durchschittspreis - Währung`"
 Str(0, 8, 26) = "`Gewinn`"
 Str(0, 8, 27) = "`Gewinn - Währung`"
 Str(0, 8, 28) = "`relativer Gewinn`"
 Str(0, 8, 29) = "`relativer Gewinn - Währung`"
 Str(0, 8, 30) = "`Stückzinsen in Depotwährung`"
 Str(0, 8, 31) = "`Stückzinsen in Depotwährung - Währung`"
 Str(0, 8, 32) = "`Anzahl aufgelaufener Tage`"
 Str(0, 8, 33) = "`Aktueller Wert des gesamten Depots`"
 Str(0, 8, 34) = "`Aktueller Wert des gesamten Depots - Währung`"
 Str(0, 8, 35) = "`Depotunterposten - Stück / Nominal`"
 Str(0, 8, 36) = "`Depotunterposten - Stück / Nominal - Währung`"
 Str(0, 8, 37) = "`Depotunterposten - Verfügbar`"
 Str(0, 8, 38) = "`Depotunterposten - Lagerland`"
 Str(0, 8, 39) = "`Depotunterposten - Depotschlüssel`"
 Str(0, 8, 40) = "`Depotunterposten - Verwahrart`"
 Str(0, 8, 41) = "`Depotunterposten - Vertriebswegkennzeichen`"
 Str(0, 8, 42) = "`Depotunterposten - Lagerstelle`"
 Str(0, 8, 43) = "`Depotunterposten - Sperre bis`"
 Str(0, 8, 44) = "`Depotunterposten - Sperrvermerk / Zusatzvermerk des Kreditinstit`"
 Str(0, 8, 45) = "`ID`"
 Str(0, 8, 46) = "`MYID`"
 Str(0, 8, 47) = "`fk_T6944009`"
 Str(0, 8, 48) = "`fk_T6944009`"
 ArtZ(0, 8) = 45
 ArtZ(1, 8) = 2
 ArtZ(2, 8) = 1
 Str(1, 8, 0) = "CREATE TABLE `T6944009` ("
 Str(1, 8, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 8, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 8, 3) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 8, 4) = " `Wertpapiername` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 5) = " `Wertpapierlangname` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 6) = " `Wertpapierkennung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 7) = " `Aktueller Kurs` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 8) = " `Aktueller Kurs - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 9) = " `Aktueller Kurs - Börsenplatz` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 10) = " `Kaufkurs` decimal(10,2) DEFAULT NULL"
 Str(1, 8, 11) = " `Kaufkurs - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 12) = " `Preisherkunft` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 13) = " `Zusatz der Preisherkunft` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 14) = " `Stück / Nominal` decimal(10,2) DEFAULT NULL"
 Str(1, 8, 15) = " `Stück / Nominal - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 16) = " `Stückzinsen` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 17) = " `Stückzinsen - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 18) = " `Wert` decimal(10,2) DEFAULT NULL"
 Str(1, 8, 19) = " `Wert - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 20) = " `Kaufpreis` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 21) = " `Kaufpreis - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 22) = " `Orderkosten` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 23) = " `Orderkosten - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 24) = " `Durchschittspreis` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 25) = " `Durchschittspreis - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 26) = " `Gewinn` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 27) = " `Gewinn - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 28) = " `relativer Gewinn` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 29) = " `relativer Gewinn - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 30) = " `Stückzinsen in Depotwährung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 31) = " `Stückzinsen in Depotwährung - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 32) = " `Anzahl aufgelaufener Tage` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 33) = " `Aktueller Wert des gesamten Depots` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 8, 34) = " `Aktueller Wert des gesamten Depots - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 35) = " `Depotunterposten - Stück / Nominal` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 36) = " `Depotunterposten - Stück / Nominal - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 37) = " `Depotunterposten - Verfügbar` decimal(10,2) DEFAULT NULL"
 Str(1, 8, 38) = " `Depotunterposten - Lagerland` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 39) = " `Depotunterposten - Depotschlüssel` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 40) = " `Depotunterposten - Verwahrart` decimal(10,2) DEFAULT NULL"
 Str(1, 8, 41) = " `Depotunterposten - Vertriebswegkennzeichen` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 42) = " `Depotunterposten - Lagerstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 43) = " `Depotunterposten - Sperre bis` decimal(10,2) DEFAULT NULL"
 Str(1, 8, 44) = " `Depotunterposten - Sperrvermerk / Zusatzvermerk des Kreditinstit` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 8, 45) = " `ID` decimal(10,2) DEFAULT NULL"
 Str(1, 8, 46) = "  PRIMARY KEY (`MYID`)"
 Str(1, 8, 47) = "  KEY `fk_T6944009` (`eingid`)"
 Str(1, 8, 48) = "  CONSTRAINT `fk_T6944009` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 8, 49) = " ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 6944009_20071014.csv'"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "T8655587700"
 Str(0, 9, 1) = "`MYID`"
 Str(0, 9, 2) = "`eingid`"
 Str(0, 9, 3) = "`Laufende Nummer`"
 Str(0, 9, 4) = "`Wertpapiername`"
 Str(0, 9, 5) = "`Wertpapierlangname`"
 Str(0, 9, 6) = "`Wertpapierkennung`"
 Str(0, 9, 7) = "`Aktueller Kurs`"
 Str(0, 9, 8) = "`Aktueller Kurs - Währung`"
 Str(0, 9, 9) = "`Aktueller Kurs - Börsenplatz`"
 Str(0, 9, 10) = "`Kaufkurs`"
 Str(0, 9, 11) = "`Kaufkurs - Währung`"
 Str(0, 9, 12) = "`Preisherkunft`"
 Str(0, 9, 13) = "`Zusatz der Preisherkunft`"
 Str(0, 9, 14) = "`Stück / Nominal`"
 Str(0, 9, 15) = "`Stück / Nominal - Währung`"
 Str(0, 9, 16) = "`Stückzinsen`"
 Str(0, 9, 17) = "`Stückzinsen - Währung`"
 Str(0, 9, 18) = "`Wert`"
 Str(0, 9, 19) = "`Wert - Währung`"
 Str(0, 9, 20) = "`Kaufpreis`"
 Str(0, 9, 21) = "`Kaufpreis - Währung`"
 Str(0, 9, 22) = "`Orderkosten`"
 Str(0, 9, 23) = "`Orderkosten - Währung`"
 Str(0, 9, 24) = "`Durchschittspreis`"
 Str(0, 9, 25) = "`Durchschittspreis - Währung`"
 Str(0, 9, 26) = "`Gewinn`"
 Str(0, 9, 27) = "`Gewinn - Währung`"
 Str(0, 9, 28) = "`relativer Gewinn`"
 Str(0, 9, 29) = "`relativer Gewinn - Währung`"
 Str(0, 9, 30) = "`Stückzinsen in Depotwährung`"
 Str(0, 9, 31) = "`Stückzinsen in Depotwährung - Währung`"
 Str(0, 9, 32) = "`Anzahl aufgelaufener Tage`"
 Str(0, 9, 33) = "`Aktueller Wert des gesamten Depots`"
 Str(0, 9, 34) = "`Aktueller Wert des gesamten Depots - Währung`"
 Str(0, 9, 35) = "`Depotunterposten - Stück / Nominal`"
 Str(0, 9, 36) = "`Depotunterposten - Stück / Nominal - Währung`"
 Str(0, 9, 37) = "`Depotunterposten - Verfügbar`"
 Str(0, 9, 38) = "`Depotunterposten - Lagerland`"
 Str(0, 9, 39) = "`Depotunterposten - Depotschlüssel`"
 Str(0, 9, 40) = "`Depotunterposten - Verwahrart`"
 Str(0, 9, 41) = "`Depotunterposten - Vertriebswegkennzeichen`"
 Str(0, 9, 42) = "`Depotunterposten - Lagerstelle`"
 Str(0, 9, 43) = "`Depotunterposten - Sperre bis`"
 Str(0, 9, 44) = "`Depotunterposten - Sperrvermerk / Zusatzvermerk des Kreditinstit`"
 Str(0, 9, 45) = "`ID`"
 Str(0, 9, 46) = "`Wertpapier`"
 Str(0, 9, 47) = "`Kennnummer`"
 Str(0, 9, 48) = "`Boersenpl`"
 Str(0, 9, 49) = "`St_Nom`"
 Str(0, 9, 50) = "`St_NomW`"
 Str(0, 9, 51) = "`WertW`"
 Str(0, 9, 52) = "`Orderkost`"
 Str(0, 9, 53) = "`OrderkostW`"
 Str(0, 9, 54) = "`KaufpreisW`"
 Str(0, 9, 55) = "`Abw_abs`"
 Str(0, 9, 56) = "`Abw_absW`"
 Str(0, 9, 57) = "`Abw_rel`"
 Str(0, 9, 58) = "`Abw_relW`"
 Str(0, 9, 59) = "`Kurs`"
 Str(0, 9, 60) = "`KursW`"
 Str(0, 9, 61) = "`Kursdatum`"
 Str(0, 9, 62) = "`MYID`"
 Str(0, 9, 63) = "`fk_T8655587700`"
 Str(0, 9, 64) = "`fk_T8655587700`"
 ArtZ(0, 9) = 61
 ArtZ(1, 9) = 2
 ArtZ(2, 9) = 1
 Str(1, 9, 0) = "CREATE TABLE `T8655587700` ("
 Str(1, 9, 1) = " `MYID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 9, 2) = " `eingid` int(10) unsigned NOT NULL"
 Str(1, 9, 3) = " `Laufende Nummer` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 4) = " `Wertpapiername` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 5) = " `Wertpapierlangname` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 6) = " `Wertpapierkennung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 7) = " `Aktueller Kurs` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 8) = " `Aktueller Kurs - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 9) = " `Aktueller Kurs - Börsenplatz` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 10) = " `Kaufkurs` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 11) = " `Kaufkurs - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 12) = " `Preisherkunft` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 13) = " `Zusatz der Preisherkunft` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 14) = " `Stück / Nominal` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 15) = " `Stück / Nominal - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 16) = " `Stückzinsen` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 17) = " `Stückzinsen - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 18) = " `Wert` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 19) = " `Wert - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 20) = " `Kaufpreis` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 21) = " `Kaufpreis - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 22) = " `Orderkosten` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 23) = " `Orderkosten - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 24) = " `Durchschittspreis` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 25) = " `Durchschittspreis - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 26) = " `Gewinn` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 27) = " `Gewinn - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 28) = " `relativer Gewinn` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 29) = " `relativer Gewinn - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 30) = " `Stückzinsen in Depotwährung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 31) = " `Stückzinsen in Depotwährung - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 32) = " `Anzahl aufgelaufener Tage` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 33) = " `Aktueller Wert des gesamten Depots` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 9, 34) = " `Aktueller Wert des gesamten Depots - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 35) = " `Depotunterposten - Stück / Nominal` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 36) = " `Depotunterposten - Stück / Nominal - Währung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 37) = " `Depotunterposten - Verfügbar` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 38) = " `Depotunterposten - Lagerland` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 39) = " `Depotunterposten - Depotschlüssel` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 40) = " `Depotunterposten - Verwahrart` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 41) = " `Depotunterposten - Vertriebswegkennzeichen` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 42) = " `Depotunterposten - Lagerstelle` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 43) = " `Depotunterposten - Sperre bis` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 44) = " `Depotunterposten - Sperrvermerk / Zusatzvermerk des Kreditinstit` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 45) = " `ID` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 46) = " `Wertpapier` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 47) = " `Kennnummer` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 48) = " `Boersenpl` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 49) = " `St_Nom` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 50) = " `St_NomW` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 51) = " `WertW` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 52) = " `Orderkost` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 53) = " `OrderkostW` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 54) = " `KaufpreisW` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 55) = " `Abw_abs` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 56) = " `Abw_absW` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 57) = " `Abw_rel` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 58) = " `Abw_relW` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 59) = " `Kurs` decimal(10,2) DEFAULT NULL"
 Str(1, 9, 60) = " `KursW` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 9, 61) = " `Kursdatum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 9, 62) = "  PRIMARY KEY (`MYID`)"
 Str(1, 9, 63) = "  KEY `fk_T8655587700` (`eingid`)"
 Str(1, 9, 64) = "  CONSTRAINT `fk_T8655587700` FOREIGN KEY (`eingid`) REFERENCES `eingelesen` (`eingid`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 9, 65) = " ENGINE=InnoDB AUTO_INCREMENT=2119 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Tabelle aus Datei: 8655587700_20080919.csv'"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "eingelesen"
 Str(0, 10, 1) = "`eingid`"
 Str(0, 10, 2) = "`verzeich`"
 Str(0, 10, 3) = "`datei`"
 Str(0, 10, 4) = "`fdt`"
 Str(0, 10, 5) = "`datum`"
 Str(0, 10, 6) = "`eingid`"
 Str(0, 10, 7) = "`vd`"
 Str(0, 10, 8) = "`fdt`"
 Str(0, 10, 9) = "`datum`"
 ArtZ(0, 10) = 5
 ArtZ(1, 10) = 4
 Str(1, 10, 0) = "CREATE TABLE `eingelesen` ("
 Str(1, 10, 1) = " `eingid` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 10, 2) = " `verzeich` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 10, 3) = " `datei` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 10, 4) = " `fdt` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 10, 5) = " `datum` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 10, 6) = "  PRIMARY KEY (`eingid`)"
 Str(1, 10, 7) = "  UNIQUE KEY `vd` (`verzeich`,`datei`)"
 Str(1, 10, 8) = "  KEY `fdt` (`fdt`)"
 Str(1, 10, 9) = "  KEY `datum` (`datum`)"
 Str(1, 10, 10) = " ENGINE=InnoDB AUTO_INCREMENT=155 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
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

Public Function doMach_depot(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
  MsgBox "Fertig mit doMach_depot(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_depot/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_depot

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
