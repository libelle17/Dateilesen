Option Explicit
Const DBn$ = "quelle7" ' Datenbankname
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz as new ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen

function doEx%(sql$, obtolerant%)
 Dim rAF&, FMeld$
 if obtolerant then on error resume next else on error goto fehler
 call cnz.execute(sql,rAf)
 lErrNr = Err.Number
 FMeld = "Err.Nr " & lErrNr & ", rAf: " & rAF & " bei " & sql
 On error goto fehler
 Debug.print FMeld
 If obProt then Print #302, FMeld
 doEvents
 exit function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case Err.Number
 Case -2147467259 'Kann Tabelle 'testDB1.faxe' nicht erzeugen (Fehler: 150)
  doEx = 150
  Exit Function
End Select
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doEx/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): ende
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

Function doGenMachDB()
 Dim zCat as new ADOX.CATALOG, rsc As New ADODB.Recordset,sct$,Spli$(),tStr$,TMt as New CString,TabEig$, i&, p1&, p2&, p3&, CLen&, CLen1&
 Dim Index$()
 On Error Resume Next
 Open "u:\programmierung\dateilesen\MachD1B.bas_prot.txt" for Output as #302
 obProt = (Err.Number = 0)
 On Error goto fehler
 cnzCStr = "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=Linux;uid=mysql;pwd=97a5o6;"
 cnz.open cnzCStr
 call doEx("create database if not exists `" & DBN & "` character set latin1 collate latin1_german2_ci;",0)
 call doEx("grant all privileges on `" & DBN & "`.* to 'praxis'@'%' with grant option",0)
 call doEx("grant all privileges on `" & DBN & "`.* to 'praxis'@'localhost' with grant option",0)
 call doEx("use `" & DBN & "`",0)
 call doEx("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ",0)
 call doEx("SET FOREIGN_KEY_CHECKS = 0",0)
