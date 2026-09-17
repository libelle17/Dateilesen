'Bauanleitung für eine Datenbank wie `//linux/dp` vom 15.11.09 14:26:12
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 7, 29) As new CString, ArtZ&(3, 7)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "arten"
 Str(0, 0, 1) = "`ArtNr`"
 Str(0, 0, 2) = "`erkl`"
 Str(0, 0, 3) = "`Std`"
 Str(0, 0, 4) = "`Farbe`"
 Str(0, 0, 5) = "`zusatz`"
 Str(0, 0, 6) = "`ArtNr`"
 ArtZ(0, 0) = 5
 ArtZ(1, 0) = 1
 Str(1, 0, 0) = "CREATE TABLE `arten` ("
 Str(1, 0, 1) = " `ArtNr` varchar(30) COLLATE latin1_german2_ci NOT NULL COMMENT 'Artnr'"
 Str(1, 0, 2) = " `erkl` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Erklärung'"
 Str(1, 0, 3) = " `Std` double(3,1) unsigned NOT NULL COMMENT 'Stunden'"
 Str(1, 0, 4) = " `Farbe` int(4) unsigned NOT NULL"
 Str(1, 0, 5) = " `zusatz` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 0, 6) = "  PRIMARY KEY (`ArtNr`)"
 Str(1, 0, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Dienstarten; InnoDB free: 261120 kB'"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "bilanzen"
 Str(0, 1, 1) = "`PersNr`"
 Str(0, 1, 2) = "`Jahr`"
 Str(0, 1, 3) = "`Urlaub`"
 Str(0, 1, 4) = "`Überstunden`"
 Str(0, 1, 5) = "`Fortbildung`"
 Str(0, 1, 6) = "`PersNr`"
 Str(0, 1, 7) = "`PersNr2`"
 ArtZ(0, 1) = 5
 ArtZ(1, 1) = 1
 ArtZ(2, 1) = 1
 Str(1, 1, 0) = "CREATE TABLE `bilanzen` ("
 Str(1, 1, 1) = " `PersNr` int(5) unsigned NOT NULL DEFAULT '0'"
 Str(1, 1, 2) = " `Jahr` int(4) unsigned NOT NULL"
 Str(1, 1, 3) = " `Urlaub` double(4,1) NOT NULL COMMENT 'Tage'"
 Str(1, 1, 4) = " `Überstunden` double(4,1) NOT NULL COMMENT 'Stunden'"
 Str(1, 1, 5) = " `Fortbildung` double(4,1) NOT NULL COMMENT 'Tage'"
 Str(1, 1, 6) = "  PRIMARY KEY (`PersNr`,`Jahr`) USING BTREE"
 Str(1, 1, 7) = "  CONSTRAINT `PersNr2` FOREIGN KEY (`PersNr`) REFERENCES `mitarbeiter` (`PersNr`)"
 Str(1, 1, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Urlaubs- und Überstundenbilanzen'"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "dienstplan"
 Str(0, 2, 1) = "`ID`"
 Str(0, 2, 2) = "`tag`"
 Str(0, 2, 3) = "`PersNr`"
 Str(0, 2, 4) = "`ArtNr`"
 Str(0, 2, 5) = "`Std`"
 Str(0, 2, 6) = "`ID`"
 Str(0, 2, 7) = "`Persnr`"
 Str(0, 2, 8) = "`Artnr`"
 Str(0, 2, 9) = "`ArtNr`"
 Str(0, 2, 10) = "`Persnr`"
 ArtZ(0, 2) = 5
 ArtZ(1, 2) = 3
 ArtZ(2, 2) = 2
 Str(1, 2, 0) = "CREATE TABLE `dienstplan` ("
 Str(1, 2, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 2, 2) = " `tag` date NOT NULL"
 Str(1, 2, 3) = " `PersNr` int(4) unsigned NOT NULL"
 Str(1, 2, 4) = " `ArtNr` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 5) = " `Std` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 2, 6) = "  PRIMARY KEY (`ID`)"
 Str(1, 2, 7) = "  KEY `Persnr` (`PersNr`)"
 Str(1, 2, 8) = "  KEY `Artnr` (`ArtNr`)"
 Str(1, 2, 9) = "  CONSTRAINT `ArtNr` FOREIGN KEY (`ArtNr`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 2, 10) = "  CONSTRAINT `Persnr` FOREIGN KEY (`PersNr`) REFERENCES `mitarbeiter` (`PersNr`)"
 Str(1, 2, 11) = " ENGINE=InnoDB AUTO_INCREMENT=3026 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Dienstplan'"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "einstellungen"
 Str(0, 3, 1) = "`Einstellung`"
 Str(0, 3, 2) = "`Wert`"
 Str(0, 3, 3) = "`Einstellung`"
 ArtZ(0, 3) = 2
 ArtZ(1, 3) = 1
 Str(1, 3, 0) = "CREATE TABLE `einstellungen` ("
 Str(1, 3, 1) = " `Einstellung` varchar(70) NOT NULL DEFAULT ''"
 Str(1, 3, 2) = " `Wert` varchar(70) DEFAULT NULL"
 Str(1, 3, 3) = "  PRIMARY KEY (`Einstellung`)"
 Str(1, 3, 4) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Einstellungen'"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "mitarbeiter"
 Str(0, 4, 1) = "`PersNr`"
 Str(0, 4, 2) = "`Kuerzel`"
 Str(0, 4, 3) = "`Nachname`"
 Str(0, 4, 4) = "`Vorname`"
 Str(0, 4, 5) = "`Aus`"
 Str(0, 4, 6) = "`KalDB`"
 Str(0, 4, 7) = "`KalTab`"
 Str(0, 4, 8) = "`KalDatSp`"
 Str(0, 4, 9) = "`KalSp1`"
 Str(0, 4, 10) = "`KalSp2`"
 Str(0, 4, 11) = "`KalSp3`"
 Str(0, 4, 12) = "`KalSp4`"
 Str(0, 4, 13) = "`PersNr`"
 Str(0, 4, 14) = "`Aus`"
 ArtZ(0, 4) = 12
 ArtZ(1, 4) = 2
 Str(1, 4, 0) = "CREATE TABLE `mitarbeiter` ("
 Str(1, 4, 1) = " `PersNr` int(5) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Personal-Nummer'"
 Str(1, 4, 2) = " `Kuerzel` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kürzel'"
 Str(1, 4, 3) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 4) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 5) = " `Aus` date DEFAULT NULL COMMENT 'Austritt'"
 Str(1, 4, 6) = " `KalDB` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 7) = " `KalTab` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 8) = " `KalDatSp` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 9) = " `KalSp1` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 10) = " `KalSp2` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 11) = " `KalSp3` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 12) = " `KalSp4` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 13) = "  PRIMARY KEY (`PersNr`)"
 Str(1, 4, 14) = "  KEY `Aus` (`Aus`)"
 Str(1, 4, 15) = " ENGINE=InnoDB AUTO_INCREMENT=68 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "protok"
 Str(0, 5, 1) = "`ID`"
 Str(0, 5, 2) = "`tag`"
 Str(0, 5, 3) = "`PersNr`"
 Str(0, 5, 4) = "`ArtNrV`"
 Str(0, 5, 5) = "`ArtNr`"
 Str(0, 5, 6) = "`AendDat`"
 Str(0, 5, 7) = "`AendPC`"
 Str(0, 5, 8) = "`AendUser`"
 Str(0, 5, 9) = "`user`"
 Str(0, 5, 10) = "`ID`"
 Str(0, 5, 11) = "`Persnr`"
 Str(0, 5, 12) = "`Artnr`"
 Str(0, 5, 13) = "`user`"
 Str(0, 5, 14) = "`ProtArtNr`"
 Str(0, 5, 15) = "`ProtPersnr`"
 ArtZ(0, 5) = 9
 ArtZ(1, 5) = 4
 ArtZ(2, 5) = 2
 Str(1, 5, 0) = "CREATE TABLE `protok` ("
 Str(1, 5, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 5, 2) = " `tag` date NOT NULL"
 Str(1, 5, 3) = " `PersNr` int(5) unsigned NOT NULL"
 Str(1, 5, 4) = " `ArtNrV` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 5) = " `ArtNr` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 6) = " `AendDat` datetime DEFAULT NULL"
 Str(1, 5, 7) = " `AendPC` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 8) = " `AendUser` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 5, 9) = " `user` varchar(45) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 5, 10) = "  PRIMARY KEY (`ID`)"
 Str(1, 5, 11) = "  KEY `Persnr` (`PersNr`)"
 Str(1, 5, 12) = "  KEY `Artnr` (`ArtNr`)"
 Str(1, 5, 13) = "  KEY `user` (`user`)"
 Str(1, 5, 14) = "  CONSTRAINT `ProtArtNr` FOREIGN KEY (`ArtNr`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 5, 15) = "  CONSTRAINT `ProtPersnr` FOREIGN KEY (`PersNr`) REFERENCES `mitarbeiter` (`PersNr`)"
 Str(1, 5, 16) = " ENGINE=InnoDB AUTO_INCREMENT=4439 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Änderungsprotokoll für Dienstplan'"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "user"
 Str(0, 6, 1) = "`user`"
 Str(0, 6, 2) = "`Passwort`"
 Str(0, 6, 3) = "`hinzugefügt`"
 Str(0, 6, 4) = "`geändert`"
 Str(0, 6, 5) = "`ID`"
 Str(0, 6, 6) = "`ID`"
 Str(0, 6, 7) = "`user`"
 ArtZ(0, 6) = 5
 ArtZ(1, 6) = 2
 Str(1, 6, 0) = "CREATE TABLE `user` ("
 Str(1, 6, 1) = " `user` varchar(45) CHARACTER SET latin1 NOT NULL"
 Str(1, 6, 2) = " `Passwort` blob NOT NULL"
 Str(1, 6, 3) = " `hinzugefügt` datetime NOT NULL"
 Str(1, 6, 4) = " `geändert` datetime NOT NULL"
 Str(1, 6, 5) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 6, 6) = "  PRIMARY KEY (`ID`)"
 Str(1, 6, 7) = "  UNIQUE KEY `user` (`user`)"
 Str(1, 6, 8) = " ENGINE=InnoDB AUTO_INCREMENT=25 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Benutzer für Datenänderungen'"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "wochenplan"
 Str(0, 7, 1) = "`PersNr`"
 Str(0, 7, 2) = "`ab`"
 Str(0, 7, 3) = "`Mo`"
 Str(0, 7, 4) = "`Di`"
 Str(0, 7, 5) = "`Mi`"
 Str(0, 7, 6) = "`Do`"
 Str(0, 7, 7) = "`Fr`"
 Str(0, 7, 8) = "`Sa`"
 Str(0, 7, 9) = "`So`"
 Str(0, 7, 10) = "`WAZ`"
 Str(0, 7, 11) = "`Urlaub`"
 Str(0, 7, 12) = "`PersNr`"
 Str(0, 7, 13) = "`Mo`"
 Str(0, 7, 14) = "`Di`"
 Str(0, 7, 15) = "`Mi`"
 Str(0, 7, 16) = "`Do`"
 Str(0, 7, 17) = "`Fr`"
 Str(0, 7, 18) = "`Sa`"
 Str(0, 7, 19) = "`So`"
 Str(0, 7, 20) = "`DiArt`"
 Str(0, 7, 21) = "`DoArt`"
 Str(0, 7, 22) = "`FrArt`"
 Str(0, 7, 23) = "`MiArt`"
 Str(0, 7, 24) = "`MoArt`"
 Str(0, 7, 25) = "`Persnr1`"
 Str(0, 7, 26) = "`SaArt`"
 Str(0, 7, 27) = "`SoArt`"
 ArtZ(0, 7) = 11
 ArtZ(1, 7) = 8
 ArtZ(2, 7) = 8
 Str(1, 7, 0) = "CREATE TABLE `wochenplan` ("
 Str(1, 7, 1) = " `PersNr` int(5) unsigned NOT NULL DEFAULT '0' COMMENT 'Personal-Nr.'"
 Str(1, 7, 2) = " `ab` date NOT NULL DEFAULT '2007-01-01' COMMENT 'Gültigkeitsbeginn'"
 Str(1, 7, 3) = " `Mo` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Montag'"
 Str(1, 7, 4) = " `Di` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Dienstag'"
 Str(1, 7, 5) = " `Mi` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Mittwoch'"
 Str(1, 7, 6) = " `Do` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Donnerstag'"
 Str(1, 7, 7) = " `Fr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitag'"
 Str(1, 7, 8) = " `Sa` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Samstag'"
 Str(1, 7, 9) = " `So` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Sonntag'"
 Str(1, 7, 10) = " `WAZ` double(3,1) NOT NULL DEFAULT '38.5'"
 Str(1, 7, 11) = " `Urlaub` int(2) NOT NULL DEFAULT '26' COMMENT 'Urlaubstage pro Jahr'"
 Str(1, 7, 12) = "  PRIMARY KEY (`PersNr`,`ab`)"
 Str(1, 7, 13) = "  KEY `Mo` (`Mo`)"
 Str(1, 7, 14) = "  KEY `Di` (`Di`)"
 Str(1, 7, 15) = "  KEY `Mi` (`Mi`)"
 Str(1, 7, 16) = "  KEY `Do` (`Do`)"
 Str(1, 7, 17) = "  KEY `Fr` (`Fr`)"
 Str(1, 7, 18) = "  KEY `Sa` (`Sa`)"
 Str(1, 7, 19) = "  KEY `So` (`So`)"
 Str(1, 7, 20) = "  CONSTRAINT `DiArt` FOREIGN KEY (`Di`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 7, 21) = "  CONSTRAINT `DoArt` FOREIGN KEY (`Do`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 7, 22) = "  CONSTRAINT `FrArt` FOREIGN KEY (`Fr`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 7, 23) = "  CONSTRAINT `MiArt` FOREIGN KEY (`Mi`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 7, 24) = "  CONSTRAINT `MoArt` FOREIGN KEY (`Mo`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 7, 25) = "  CONSTRAINT `Persnr1` FOREIGN KEY (`PersNr`) REFERENCES `mitarbeiter` (`PersNr`)"
 Str(1, 7, 26) = "  CONSTRAINT `SaArt` FOREIGN KEY (`Sa`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 7, 27) = "  CONSTRAINT `SoArt` FOREIGN KEY (`So`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 7, 28) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Wochenplan'"
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

Public Function doMach_dp(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
  MsgBox "Fertig mit doMach_dp(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_dp/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_dp

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
