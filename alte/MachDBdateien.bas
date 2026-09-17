'Bauanleitung für eine Datenbank wie `//linux/dateien` vom 15.11.09 14:25:53
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 4, 14) As new CString, ArtZ&(3, 4)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "Fehlerordner"
 Str(1, 0, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `Fehlerordner` AS select `f`.`ID` AS `ID`,`f`.`ErrNr` AS `ErrNr`,`f`.`ErrDes` AS `ErrDes`,`f`.`ErrLastDLLErr` AS `ErrLastDLLErr`,`f`.`ErrSource` AS `ErrSource`,`f`.`eingetragen` AS `eingetragen`,`f`.`Ordner` AS `Ordner`,`o`.`Name` AS `oname`,`o1`.`Name` AS `o1name`,`o2`.`Name` AS `o2name`,`o3`.`Name` AS `o3name`,`o4`.`Name` AS `o4name` from (((((`fehler` `f` left join `ordner` `o` on((`f`.`Ordner` = `o`.`ID`))) left join `ordner` `o1` on((`o`.`Parent` = `o1`.`ID`))) left join `ordner` `o2` on((`o1`.`Parent` = `o2`.`ID`))) left join `ordner` `o3` on((`o2`.`Parent` = `o3`.`ID`))) left join `ordner` `o4` on((`o3`.`Parent` = `o4`.`ID`)))"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "dateien"
 Str(0, 1, 1) = "`ID`"
 Str(0, 1, 2) = "`Name`"
 Str(0, 1, 3) = "`Größe`"
 Str(0, 1, 4) = "`geändert`"
 Str(0, 1, 5) = "`eingetragen`"
 Str(0, 1, 6) = "`Parent`"
 Str(0, 1, 7) = "`ID`"
 Str(0, 1, 8) = "`Name`"
 Str(0, 1, 9) = "`Größe`"
 Str(0, 1, 10) = "`geändert`"
 Str(0, 1, 11) = "`Parent`"
 Str(0, 1, 12) = "`dateienordner`"
 ArtZ(0, 1) = 6
 ArtZ(1, 1) = 5
 ArtZ(2, 1) = 1
 Str(1, 1, 0) = "CREATE TABLE `dateien` ("
 Str(1, 1, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 1, 2) = " `Name` varchar(256) COLLATE latin1_german2_ci NOT NULL COMMENT 'Dateiname'"
 Str(1, 1, 3) = " `Größe` int(15) DEFAULT NULL COMMENT 'Größe'"
 Str(1, 1, 4) = " `geändert` datetime DEFAULT NULL COMMENT 'Letzte Änderung'"
 Str(1, 1, 5) = " `eingetragen` datetime DEFAULT NULL COMMENT 'Eintrag in diese Tabelle'"
 Str(1, 1, 6) = " `Parent` int(10) unsigned NOT NULL COMMENT 'ID des Ordners'"
 Str(1, 1, 7) = "  PRIMARY KEY (`ID`)"
 Str(1, 1, 8) = "  KEY `Name` (`Name`)"
 Str(1, 1, 9) = "  KEY `Größe` (`Größe`)"
 Str(1, 1, 10) = "  KEY `geändert` (`geändert`)"
 Str(1, 1, 11) = "  KEY `Parent` (`Parent`)"
 Str(1, 1, 12) = "  CONSTRAINT `dateienordner` FOREIGN KEY (`Parent`) REFERENCES `ordner` (`ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 1, 13) = " ENGINE=InnoDB AUTO_INCREMENT=1284072 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Dateien'"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "fehler"
 Str(0, 2, 1) = "`ID`"
 Str(0, 2, 2) = "`ErrNr`"
 Str(0, 2, 3) = "`ErrDes`"
 Str(0, 2, 4) = "`ErrLastDLLErr`"
 Str(0, 2, 5) = "`ErrSource`"
 Str(0, 2, 6) = "`eingetragen`"
 Str(0, 2, 7) = "`Ordner`"
 Str(0, 2, 8) = "`ID`"
 Str(0, 2, 9) = "`ErrNr`"
 Str(0, 2, 10) = "`Ordner`"
 Str(0, 2, 11) = "`fehlerordner`"
 ArtZ(0, 2) = 7
 ArtZ(1, 2) = 3
 ArtZ(2, 2) = 1
 Str(1, 2, 0) = "CREATE TABLE `fehler` ("
 Str(1, 2, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 2, 2) = " `ErrNr` int(10) NOT NULL COMMENT 'Fehlernummer'"
 Str(1, 2, 3) = " `ErrDes` varchar(256) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Beschreibung'"
 Str(1, 2, 4) = " `ErrLastDLLErr` int(10) DEFAULT NULL COMMENT 'Letzter DLL-Fehler'"
 Str(1, 2, 5) = " `ErrSource` varchar(256) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Fehler-Entstehung'"
 Str(1, 2, 6) = " `eingetragen` datetime DEFAULT NULL COMMENT 'Eintrag in diese Tabelle'"
 Str(1, 2, 7) = " `Ordner` int(10) unsigned NOT NULL COMMENT 'ID des Ordners'"
 Str(1, 2, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 2, 9) = "  KEY `ErrNr` (`ErrNr`)"
 Str(1, 2, 10) = "  KEY `Ordner` (`Ordner`)"
 Str(1, 2, 11) = "  CONSTRAINT `fehlerordner` FOREIGN KEY (`Ordner`) REFERENCES `ordner` (`ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 2, 12) = " ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Fehler'"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "ordner"
 Str(0, 3, 1) = "`ID`"
 Str(0, 3, 2) = "`Name`"
 Str(0, 3, 3) = "`Parent`"
 Str(0, 3, 4) = "`Suche`"
 Str(0, 3, 5) = "`Anfang`"
 Str(0, 3, 6) = "`Ende`"
 Str(0, 3, 7) = "`ID`"
 Str(0, 3, 8) = "`Name`"
 Str(0, 3, 9) = "`ordnersuchen`"
 Str(0, 3, 10) = "`ordnersuchen`"
 ArtZ(0, 3) = 6
 ArtZ(1, 3) = 3
 ArtZ(2, 3) = 1
 Str(1, 3, 0) = "CREATE TABLE `ordner` ("
 Str(1, 3, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 3, 2) = " `Name` varchar(256) COLLATE latin1_german2_ci NOT NULL COMMENT 'auch Laufwerk'"
 Str(1, 3, 3) = " `Parent` int(10) unsigned NOT NULL COMMENT 'ID des übergeordneten Ordners'"
 Str(1, 3, 4) = " `Suche` int(10) unsigned NOT NULL COMMENT 'ID der Suche'"
 Str(1, 3, 5) = " `Anfang` datetime DEFAULT NULL COMMENT 'Ordner angefangen eingelesen'"
 Str(1, 3, 6) = " `Ende` datetime DEFAULT NULL COMMENT 'Ordner fertig eingelesen'"
 Str(1, 3, 7) = "  PRIMARY KEY (`ID`)"
 Str(1, 3, 8) = "  KEY `Name` (`Name`)"
 Str(1, 3, 9) = "  KEY `ordnersuchen` (`Suche`)"
 Str(1, 3, 10) = "  CONSTRAINT `ordnersuchen` FOREIGN KEY (`Suche`) REFERENCES `suchen` (`ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 3, 11) = " ENGINE=InnoDB AUTO_INCREMENT=37903 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Verzeichnisse'"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "suchen"
 Str(0, 4, 1) = "`ID`"
 Str(0, 4, 2) = "`Name`"
 Str(0, 4, 3) = "`Anfang`"
 Str(0, 4, 4) = "`Ende`"
 Str(0, 4, 5) = "`ID`"
 Str(0, 4, 6) = "`Name`"
 Str(0, 4, 7) = "`Anfang`"
 ArtZ(0, 4) = 4
 ArtZ(1, 4) = 3
 Str(1, 4, 0) = "CREATE TABLE `suchen` ("
 Str(1, 4, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 4, 2) = " `Name` varchar(256) COLLATE latin1_german2_ci NOT NULL COMMENT 'auch Laufwerk'"
 Str(1, 4, 3) = " `Anfang` datetime DEFAULT NULL COMMENT 'Ordner angefangen eingelesen'"
 Str(1, 4, 4) = " `Ende` datetime DEFAULT NULL COMMENT 'Ordner fertig eingelesen'"
 Str(1, 4, 5) = "  PRIMARY KEY (`ID`)"
 Str(1, 4, 6) = "  KEY `Name` (`Name`)"
 Str(1, 4, 7) = "  KEY `Anfang` (`Anfang`)"
 Str(1, 4, 8) = " ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Suchvorgänge'"
End Sub ' FüllStr4

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

Public Function doMach_dateien(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 4
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
 For i = 0 To 4
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
  For i = 0 To 4
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
  MsgBox "Fertig mit doMach_dateien(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_dateien/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_dateien

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
