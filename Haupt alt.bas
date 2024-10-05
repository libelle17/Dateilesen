Attribute VB_Name = "Haupt"
Option Explicit
Public Const opti& = 2 + 4 + 8 + 32 ' 131118, 32 ' + 2048 + 16384
Public Const CSStr$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;uid=praxis;pwd=sonne;OPTION=" & opti ' database=quelle1;
Public DBCn As New ADODB.Connection
Public ErrNumber&, ErrDescription$
Public rs As New ADODB.Recordset
Public rsAdo As ADODB.Recordset
Public FSO As New FileSystemObject
Public ConStr$ ' Connection String
Public Const CStrAcc$ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public Const CStrMy$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;user=praxis;pwd=sonne;database="
Public imAufBau1 As Boolean, imAufBau2 As Boolean
Public Const RegStelle$ = "Software\GSProducts\DateiLese"
Public FPos& ' Fehlerposition
Public obMySQL% ' ob MySQL oder Access
Public myDB$ ' quelle, quelle1, quelle2
Public kla$, klz$ ' Klammer auf, Klammer zu f¸r Tabellen- und Feldnamen in Access [], in MySQL ``
'Public dlg As Dialog
Const TbZ1% = 19
Const TbZ2% = TbZ1 + 7
Public Lese1
Public AllePat% ' alle Patienten sind in der Datei
Public obVorb% ' ob Formularvorbereitung
Public obStart% ' ob Startvorgang (da keinen Datei÷ffnen-Dialog zeigen) 'ConnectionString verbinden usw.)
Dim Tbn$(TbZ2), Tbk$(TbZ2)
Public ZielDbS$
Public BrichAb% ' bricht gleich ab
Public ProgL‰uft%

Function machAbK¸$()
 Dim i%, j%, K%
 Tbn(0) = "Namen"
 Tbn(1) = "Faelle"
 Tbn(2) = "AU"
 Tbn(3) = "Briefe"
 Tbn(4) = "Diagnosen"
 Tbn(5) = "Dokumente"
 Tbn(6) = "Eintraege"
 Tbn(7) = "FormInhaltForm_Abk"
 Tbn(8) = "Formulare"
 Tbn(9) = "FormInhKopf" 'altNeu"
 Tbn(10) = "FormInhFeld"
 Tbn(11) = "KHEinweis"
 Tbn(12) = "LBAnforderungen"
 Tbn(13) = "LaborNeu"
 Tbn(14) = "Leistungen"
 Tbn(15) = "MedPlan"
 Tbn(16) = "RezeptEintraege"
 Tbn(17) = "RR"
 Tbn(18) = "KVNrUe"
 Tbn(19) = "unbekannte kennungen"
 Tbn(20) = "laborxsaetze"
 Tbn(21) = "laborxeingel"
 Tbn(22) = "laborxus"
 Tbn(23) = "laborxbakt"
 Tbn(24) = "laborxwert"
 Tbn(25) = "laborxleist"
 Tbn(26) = "anamnesebogen"
 
' Tbn(17) = "FormInhaltForm_abk"
' Tbn(18) = "FormInhaltFeldinh"
' Tbn(19) = "FormInhaltFeld"
 For i = 0 To TbZ2
  Select Case Tbn(i)
   Case "FormInhaltForm_Abk": Tbk(i) = "rFi"
'   Case "Formulare": Tbk(i) = "rFo"
   Case "laborxsaetze": Tbk(i) = "rLs"
'   Case "laborxbakt":   Tbk(i) = "rLb"
   Case "laborxleist":  Tbk(i) = "rLL"
   Case "laborxwert":   Tbk(i) = "rLw"
   Case "laborxus":     Tbk(i) = "rLu"
   Case "laborxeingel": Tbk(i) = "rLg"
   Case "anamnesebogen": Tbk(i) = "rAna"
   Case Else
    For K = 2 To Len(Tbn(i))
     Tbk(i) = "r" + UCase(Left(Tbn(i), 1)) + LCase(Mid(Tbn(i), K, 1))
     For j = 0 To i - 1
      If Tbk(j) = Tbk(i) Then GoTo hammaschon
     Next j
     Exit For
hammaschon:
    Next K
  End Select
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Zeichensatz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' machAbK¸$()
Function doMacheTypen(Tabelle$)
 Dim Tabl$, Art$, rsAdSc As New ADODB.Recordset
 Dim Lenge&, obZusatzFeld%
 On Error GoTo fehler
 Tabl = LCase(Tabelle)
 Print #257, ""
 Print #257, "Public type " + Replace(Tabl, " ", "_")
 Set rsAdSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, Tabl, Empty))
 Do While Not rsAdSc.EOF
'  Debug.Print rsAdSc.Fields(2), colw(rsadsc!column_name), rsAdSc!column_propid, rsAdSc!column_flags, rsAdSc!data_type, rsAdSc!type_guid, rsAdSc!character_maximum_length, rsAdSc!numeric_scale, rsAdSc!numeric_precision, rsAdSc!Description
  Dim colN$
  colN = colW(rsAdSc!column_name)
  Art = " " + colN + " as "
  Select Case rsAdSc!data_type
   Case 3, 17, 19, 20, 21, 205, 201: Art = Art + "long"
   Case 2, 11, 16, 17, 18, 128: Art = Art + "integer" ' z.T.: "boolean"
   Case 8, 129, 130, 200, 202: Art = Art + "string"
   Case 7, 133, 134, 135: Art = Art + "date"
   Case 4, 5: Art = Art + "single"
   Case 6, 14, 131, 139: Art = Art + "double"
   Case Else: Art = Art + "variant " & " '" & colW(rsAdSc!column_name) & " " & rsAdSc!data_type
  End Select
  Art = Art & " '" & rsAdSc!Description
  Print #257, Art
  If colN Like "Feld*VW" Or colN = "LangtextVW" Or colN = "KommentarVW" Then
   Print #257, " " + Left(colN, Len(colN) - 2) + " as string"
  End If
  rsAdSc.MoveNext
 Loop
#If False Then
' If obMySQL Then
  Set rsAdo = DBCn.Execute("show columns from `" & Tabl & "`")
' End If
 Do While Not rsAdo.EOF
  Art = " " + rsAdo.Fields(0) + " as "
  If rsAdo.Fields(1) Like "varchar*" Then
   Lenge = Len(rsAdo.Fields(1))
   Art = Art + "string" '  * " + Left(Mid(rsADO.Fields(1), 9), lenge - 9)
  ElseIf rsAdo.Fields(1) Like "char*" Then
   Lenge = Len(rsAdo.Fields(1))
   Art = Art + "string" ' * " + Left(Mid(rsADO.Fields(1), 6), lenge - 6)
  ElseIf rsAdo.Fields(1) Like "int(10)*" Or rsAdo.Fields(1) Like "int(4)*" Or rsAdo.Fields(1) Like "integer" Or rsAdo.Fields(1) Like "int(11)*" Then
   Art = Art + "long"
  ElseIf rsAdo.Fields(1) Like "date*" Then
    Art = Art + "date"
  ElseIf rsAdo.Fields(1) = "longtext" Then
    Art = Art + "string"
  ElseIf rsAdo.Fields(1) = "tinyint(1)" Then
    Art = Art + "boolean"
  ElseIf rsAdo.Fields(1) = "smallint(6)" Or rsAdo.Fields(1) = "tinyint(4)" Then
   Art = Art + "integer"
  ElseIf rsAdo.Fields(1) = "bit(8)" Or rsAdo.Fields(1) = "int(1) unsigned" Then
   Art = Art + "byte"
  Else
    Err.Raise 999, , "Fehler in doMacheTypen: rsAdo.Fields(1): " & rsAdo.Fields(1) & "(Datentyp nicht vorgesehen)"
  End If
  Print #257, Art
  If rsAdo.Fields(0) Like "Feld*VW" Then
   Print #257, " " + Left(rsAdo.Fields(0), Len(rsAdo.Fields(0)) - 2) + " as string"
  End If
  rsAdo.Move 1
 Loop
#End If
 Print #257, "end type"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMacheTypen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doMacheTypen
Function obAutoIncr%(Tabl$, rs As ADODB.Recordset)
 Dim test%, Spalt$
 On Error GoTo fehler
#If False Then
 Dim rsDir As New ADODB.Recordset, fi%, pi%
 rsDir.CursorLocation = adUseServer
 rsDir.Open Tabl, DBCn, adOpenStatic, adLockReadOnly
' For fi = 0 To rsDir.Fields.Count - 1
'  For pi = 0 To rsDir.Fields(fi).Properties.Count - 1
'   Debug.Print rsDir.Fields(fi).Properties(pi).Name, rsDir.Fields(fi).Properties(pi).Value
   If rsDir.Fields(Sp).Properties("ISAUTOINCREMENT") Then
    obAutoIncr = True
   End If
'  Next pi
' Next fi
#End If
 If obMySQL Then
  If rs!data_type = 3 And rs!column_flags = 120 And rs!column_hasdefault = True And rs!column_default = 0 Then
   obAutoIncr = True
  End If
 Else
  Static catx As New ADOX.Catalog
  If catx.ActiveConnection Is Nothing Then
   catx.ActiveConnection = ConStr
  End If
  On Error Resume Next
  Spalt = rs!column_name
  test = catx.Tables(Tabl).Columns(Spalt).Properties.Count
  If Err.Number <> 0 Then
   Exit Function
  End If
  On Error GoTo fehler
  If catx.Tables(Tabl).Columns(Spalt).Properties.Count > 0 Then
   If catx.Tables(Tabl).Columns(Spalt).Properties("AutoIncrement") Then
    obAutoIncr = True
   End If
  End If
 End If ' obMySQL
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obAutoIncr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' obAutoIncr
Function colW$(col$)
 colW = Replace(Replace(Replace(col, " ", "_"), "?", ""), "-", "_")
End Function ' colW$(col$)
Function doMachSQL0$(TName$, NobAI%) ' NobAI = es existiert eine Zeile ohne Autoincrement
 On Error GoTo fehler
 Dim Tabl$, Art$, VStr$, Lenge&, rsAdSc As New ADODB.Recordset, SpZ%
 Dim trennz$
 Dim obHierKAI%, colN$ ' ob hier kein AutoIncrement
 NobAI = 0
 Tabl = LCase(TName)
 doMachSQL0 = " ("
' If dbcn.State = 0 Then Call dbcn.Open(CSStr)
' Call dbcn.Execute("use quelle1;")
 Set rsAdSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, Tabl, Empty))
 If Not rsAdSc.BOF Then
  Do While Not rsAdSc.EOF
   Do
    obHierKAI = Not obAutoIncr(Tabl, rsAdSc)
    If obHierKAI Then colN = colW(rsAdSc!column_name)
    rsAdSc.Move 1
   Loop Until obHierKAI Or rsAdSc.EOF
   If rsAdSc.EOF Then trennz = ")" Else trennz = ","
   If obHierKAI Then
    doMachSQL0 = doMachSQL0 & colN & trennz
    NobAI = True
   End If
   If Not rsAdSc.EOF Then If SpZ Mod 10 = 0 Then doMachSQL0 = doMachSQL0 + """ & _" + vbCrLf + "     """
   SpZ = SpZ + 1
  Loop
 End If
#If False Then
 Set rsAdo = DBCn.Execute("show columns from `" & Tabl & "`")
 Do While Not rsAdo.EOF
  doMachSQL0 = doMachSQL0 + rsAdo.Fields(0)
  doMachSQL0 = doMachSQL0 & ","
  rsAdo.Move 1
 Loop
#End If
' doMachSQL0 = Left(doMachSQL0, Len(doMachSQL0) - 1) + ")"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMachSQL0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doMachSQL1(ti%, sql$)

Function doMachSQL1$(ti%, ptxt$)
 Dim Tabl$, Art$, VStr$
 Dim Lenge&, umbrZahl%
 Dim trennz$, tza$, tze1$, tze2$
 Dim obHierKAI%, colN$, rDT$ ' ob hier kein AutoIncrement / Column-Name / Recordset-Datatype
 Dim rsAdSc As New ADODB.Recordset
 umbrZahl = 1
 Tabl = LCase(Tbn(ti))
 
 Set rsAdSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, Tabl, Empty))
 If Not rsAdSc.BOF Then
  Do While Not rsAdSc.EOF
   Do
    obHierKAI = Not obAutoIncr(Tabl, rsAdSc)
    If obHierKAI Then
     colN = colW(rsAdSc!column_name)
     rDT = rsAdSc!data_type
    End If
    rsAdSc.Move 1
   Loop Until obHierKAI Or rsAdSc.EOF
   If rsAdSc.EOF Then trennz = ")""" Else trennz = ","
   If obHierKAI Then
    VStr = Tbk(ti) & "(i)." & colN
    Select Case rDT
'    Case 2: ' "integer"
     Case 3, 17, 19, 20, 21: tza = """ & ": tze1 = " & ": tze2 = """" ' "long"
     Case 4, 5: tza = """ & ": tze1 = " & ": tze2 = """" 'single/ double
     Case 2, 11, 16, 17, 18, 128: tza = """ & cstr(cint(": tze1 = ")) & ": tze2 = """" ' VStr = "cStr(cint(" & VStr & "))" & trennz ' int, z.T.: "boolean"
     Case 8, 129, 130: tza = "'"" & ": tze1 = " & ": tze2 = """'" 'VStr = """'"" + " & VStr & " + ""'" & trennz & """"
     Case 7, 133, 134, 135: tza = """ & datForm(": tze1 = ") & ": tze2 = """" 'VStr = " datForm(" & VStr & ")" & trennz & """"
     Case Else: Err.Raise 999, , "Fehler in doMachSQL1: rDT: " & rDT & " (Datentyp nicht vorgesehen)"
    End Select
    ptxt = ptxt & tza & VStr & tze1
    
    If Len(ptxt) / umbrZahl > 150 Then
     umbrZahl = umbrZahl + 1
     If umbrZahl Mod 15 = 1 Then
'      Print #257, ptxt & """" ' & " _"
      ptxt = ptxt & """" & vbCrLf & "  sql = sql & "
     Else
'      Print #257, ptxt & " _"
      ptxt = ptxt & " _" & vbCrLf & "   "
     End If
    End If
    ptxt = ptxt & tze2 & trennz
   End If
  Loop
 End If ' not rsAdSc.BOF
 
#If False Then
 Set rsAdo = DBCn.Execute("show columns from `" & Tabl & "`")
 Do While Not rsAdo.EOF
  sql = sql & """ & "
  VStr = Tbk(ti) & "(i)." & rsAdo.Fields(0)
  If rsAdo.Fields(1) Like "*char*" Or rsAdo.Fields(1) Like "*text*" Then
'   VStr = """'"" + " + "replace(replace(trim(" + VStr + "),""'"",""\'""),""\"",""\\"")" + " + ""'"""
'   VStr = """'"" + " + "trim(" + VStr + ")" + " + ""'"""
   VStr = """'"" + " + VStr + " + ""'"""
  ElseIf rsAdo.Fields(1) = "tinyint(1)" Then
   VStr = "cStr(cint(" & VStr & "))"
  ElseIf rsAdo.Fields(1) Like "*int*" Then
'   VStr = "cStr(clng(" + VStr + "))"
  ElseIf rsAdo.Fields(1) Like "*date*" Or rsAdo.Fields(1) Like "*time*" Then
'   VStr = """'""" + " + format(" + VStr + ", ""yyyy-mm-dd hh:mm:ss"") + " + """'"""
   VStr = " datForm(" & VStr & ")" & ""
  Else
   Stop
  End If
  sql = sql + VStr
  If Len(sql) > 150 Then
   Print #257, sql & " _"
   sql = " "
  End If
  sql = sql & " & "","
  rsAdo.Move 1
 Loop
#End If
End Function ' doMachSQL1(ti%, sql$)
#If False Then
Function doMachSQL0(ti%, sql$)
 Dim Tabl$, Art$, VStr$, pos%
 Dim Lenge&
 Tabl = LCase(Tbn(ti))
 Set rsAdo = DBCn.Execute("show columns from `" & Tabl & "`")
 Do While Not rsAdo.EOF
  If rsAdo.Fields(0) <> "StByte" Then
   sql = sql + rsAdo.Fields(0) & " = "" & "
   VStr = Tbn(ti) & "(i)." + rsAdo.Fields(0)
   If rsAdo.Fields(1) Like "*char*" Or rsAdo.Fields(1) Like "*text*" Then
    VStr = """'"" & " & "trim(" & VStr & ")" & " & ""'"""
   ElseIf rsAdo.Fields(1) Like "*int*" Then
    'VStr = "cStr(clng(" & VStr & "))"
   ElseIf rsAdo.Fields(1) Like "*date*" Or rsAdo.Fields(1) Like "*time*" Then
    VStr = """'""" & " & format(" & VStr & ", ""yyyy-mm-dd hh:mm:ss"") & " & """'"""
   Else
    Stop
   End If
   sql = sql + VStr
   If Len(sql) > 150 Then
    Print #257, sql & " & _"
    sql = "  "" and "
   Else
    sql = sql & " & "" and "
   End If
  End If
  rsAdo.Move 1
 Loop
End Function ' doMachSQL0(ti%, sql$)
#End If
'Function tabK¸$(Tl$)
' tabK¸ = "r" + UCase(Left(Tl, 1)) + LCase(Mid(Tl, 2, 1))
'End Function

Function MacheTypen(frm As Lese)
 Const SammelIns% = 0
 Dim rsAdSc As New ADODB.Recordset
 Dim i%, sql$, ptext$
 Dim typFile$
 On Error GoTo fehler
 typFile = App.Path + "\" + "typen.bas"
 Open typFile For Output As #257
 Call machAbK¸
' dbcn.Open CSStr
' Call dbcn.Execute("use " & myDB)
 Print #257, "Dim sql$"
' Print #257, ""
 For i = 0 To TbZ2
  Call doMacheTypen(Tbn(i))
 Next i
 Print #257, ""
 For i = 0 To TbZ2
  Print #257, "Public " + Tbk(i) + "() as " + LCase(Replace(Tbn(i), " ", "_"))
 Next i
 Print #257, ""
 Print #257, "Public Function Tinit()"
 Print #257, " static wdh%"
 Print #257, " ReDim rAna(0)"
' Print #257, " ReDim rFm(0)"
 For i = 0 To TbZ1
  Select Case Tbk(i)
   Case "rFo", "rFi"
   Print #257, " if wdh = 0 then reDim " + Tbk(i) + "(0)"
   Case Else
   Print #257, " ReDim " + Tbk(i) + "(0)"
  End Select
 Next i
 Print #257, " wdh = -1"
 Print #257, "End Function ' Tinit"
 Print #257, ""
 Print #257, "Public Function LabInit()"
 Print #257, " static wdh%"
' Print #257, " ReDim rFm(0)"
 For i = TbZ1 + 1 To TbZ2 - 1
  Select Case Tbk(i)
   Case "rFo", "rFi"
   Print #257, " if wdh = 0 then reDim " + Tbk(i) + "(0)"
   Case Else
   Print #257, " ReDim " + Tbk(i) + "(0)"
  End Select
 Next i
 Print #257, " wdh = -1"
 Print #257, "End Function ' Tinit"
 Print #257, ""
 Print #257, "Public Function doLˆsch(frm as lese, Tbl$)"
 Print #257, " Set rs = dbcn.Execute(""select count(*) as ct from "" + kla + Tbl + klz)"
 Print #257, " frm.Ausgabe = ""Lˆsche: "" & kla & Tbl & klz & "" ("" & rs!ct & "" Datens‰tze)"" & vbCrLf & frm.Ausgabe"
 Print #257, " If obMySQL Then"
 Print #257, "  sql = ""truncate table "" & kla & Tbl & klz"
 Print #257, " Else"
 Print #257, "  sql = ""delete from "" & kla & Tbl & klz"
 Print #257, " End If"
 Print #257, " Call dbcn.Execute(sql) ' ,,adasyncexecute"
 Print #257, " DoEvents"
 Print #257, "End Function ' doLˆsch"
 Print #257, ""
 Print #257, "Public Function AllesLˆsch(frm as lese)"
 Print #257, " Dim ct&, rs as new ADODB.recordset"
 Print #257, " on error goto fehler"
 Print #257, " if obmysql and obForeign then Call dbcn.Execute(""set FOREIGN_KEY_CHECKS = 0"")"
 For i = TbZ1 To 0 Step -1
  Print #257, " call doLˆsch(frm, """ & LCase(Tbn(i)) & """)"
 Next i
 Print #257, " if obmysql and obForeign then Call dbcn.Execute(""set FOREIGN_KEY_CHECKS = 1"")"
 Print #257, " exit function"
 Print #257, "fehler:"
 Print #257, " Dim AnwPfad$"
 Print #257, " #If VBA6 Then"
 Print #257, " AnwPfad = CurrentDb.Name"
 Print #257, " #Else"
 Print #257, " AnwPfad = App.Path"
 Print #257, " #End If"
 Print #257, " Select Case MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(IsNull(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""Aufgefangener Fehler in allesLˆsch/"" + AnwPfad)"
 Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): End"
 Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
 Print #257, " end Select"
 Print #257, "End Function ' AllesLˆsch"
 Print #257, ""
 Print #257, "Public Function LabLˆsch(frm as lese)"
 Print #257, " Dim ct&, rs as new ADODB.recordset"
 Print #257, " on error goto fehler"
 Print #257, " if obmysql and obForeign then Call dbcn.Execute(""set FOREIGN_KEY_CHECKS = 0"")"
 For i = TbZ2 - 1 To TbZ1 + 1 Step -1
  Print #257, " call doLˆsch(frm, """ & LCase(Tbn(i)) & """)"
 Next i
 Print #257, " if obmysql and obForeign then Call dbcn.Execute(""set FOREIGN_KEY_CHECKS = 1"")"
 Print #257, " exit function"
 Print #257, "fehler:"
 Print #257, " Dim AnwPfad$"
 Print #257, " #If VBA6 Then"
 Print #257, " AnwPfad = CurrentDb.Name"
 Print #257, " #Else"
 Print #257, " AnwPfad = App.Path"
 Print #257, " #End If"
 Print #257, " Select Case MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(IsNull(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""Aufgefangener Fehler in LabLˆsch/"" + AnwPfad)"
 Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): End"
 Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
 Print #257, " end Select"
 Print #257, "End Function ' LabLˆsch"
 Print #257, ""
#If False Then
 Print #257, "Public Function rAnaHin(rs As ADODB.Recordset, rAna As " & LCase(Tbn(TbZ2)) & ")" ' anamnesebogen; auf die Ersetzung von "rAna" durch tbk(tbz2) wurde aus mehreren Gr¸nden verzichtet
 Set rsAdSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, LCase(Tbn(TbZ2)), Empty))
 Do While Not rsAdSc.EOF
  Print #257, " rAna." & colW(rsAdSc!column_name) & " = rs.fields(""" & rsAdSc!column_name & """)"
  rsAdSc.Move 1
 Loop
 Print #257, "End Function ' rAnaHin"
#End If
 For i = 0 To TbZ2 - 1
  Print #257, ""
  ptext = "Public Function " + LCase(Replace(Tbn(i), " ", "_")) + "Speichern"
  Select Case Tbn(i)
   Case "laborxus" ' , "laborxwert", "laborxbakt", "laborxleist"
    ptext = ptext + "(j&)"
  End Select
  Print #257, ptext
  Print #257, " Dim i%, sql0$, rAf&" ', dbcn As New adodb.Connection"
  Print #257, " on error goto fehler"
'  Print #257, " dbcn.Open CSStr"
'  Print #257, " Call dbcn.Execute(""use quelle1"")"
  If i <= TbZ1 Then
   Print #257, " if not Allepat then"
'  Print #257, "  on error resume next"
   Select Case LCase(Tbn(i))
    Case "formulare", "forminhaltform_abk", "forminhfeld", "unbekannte kennungen"
    Case Else
     Print #257, "  sql = ""delete from "" & kla & """ + LCase(Tbn(i)) + """ & klz & "" where Pat_ID = """ + " + CStr(rNa(0).Pat_ID)"
     Print #257, "  Call dbcn.Execute(sql)"
   End Select
   Print #257, " End if ' not allePat"
  End If ' i <= tbz1
'  Print #257, " If obMySQL Then On Error GoTo fehler Else On Error Resume Next"
  Dim InsertBefehl$, InsBefFld$, NobAI% ' True = es existiert ein Feld ohne AutoIncrement
  InsertBefehl = "  sql0 = ""insert ignore into "" & kla & """ + LCase(Tbn(i)) + """ & klz & """
  InsBefFld = doMachSQL0(Tbn(i), NobAI)
  Print #257, " if obMysql then"
  Print #257, IIf(SammelIns, Replace(InsertBefehl, "sql0", "sql"), InsertBefehl) + IIf(LCase(Tbn(i)) = "forminhfeld" Or NobAI, InsBefFld, "") + " values"
  Print #257, " else"
  Print #257, Replace(InsertBefehl, " ignore", "") + InsBefFld + " values"
  Print #257, " end if"
  Dim iAnf$, iEnd$
  Select Case LCase(Tbn(i))
   Case "namen", "laborxsaetze": iAnf = "0"
   Case "formulare": iAnf = "rFo1 + 1"
   Case "forminhaltform_abk": iAnf = "rFi1 + 1"
   Case "unbekannte kennungen": iAnf = "rUn1 + 1"
   Case "laborxus": iAnf = "j"
'   Case "laborxwert", "laborxbakt", "laborxleist": iAnf = "j"
   Case Else: iAnf = "1"
  End Select
  Select Case LCase(Tbn(i))
   Case "laborxus": iEnd = "j"
'   case "laborxwert", "laborxbakt", "laborxleist": iend = "j"
   Case Else: iEnd = "ubound(" + Tbk(i) + ")"
  End Select
  Print #257, " For i = " & iAnf & " to " & iEnd
  Print #257, "'  " + Tbk(i) + "(i).AktZeit = now()"
  If i <= TbZ1 Then
   Select Case LCase(Tbn(i))
    Case "forminhfeld", "forminhaltform_abk"
    Case Else
     Print #257, "  " + Tbk(i) + "(i).StByte = CStr(AktByte)"
   End Select
  End If
  'sql = "  sql = ""select pat_id from " + LCase(Tbn(i)) + " where "
  'Call doMachSQL0(Tbn(i), sql)
  'sql = Left(sql, Len(sql) - 5) + ""
  'Print #257, sql
  'Print #257, " If rs.State = 1 Then rs.Close"
  'Print #257, " rs.Open sql, dbcn, adOpenStatic, adLockReadOnly"
  'Print #257, " if rs.bof then"
  If SammelIns Then
   sql = "  sql = iif(obMySQL,sql,sql0) & ""("
  Else
   sql = "  sql = sql0 & ""("
  End If
  Call doMachSQL1(i, sql)
  Print #257, sql
  If SammelIns Then Print #257, "  if obMySQl then"
  If SammelIns Then Print #257, "   if i < ubound(" & Tbk(i) & ") then sql = sql & "","""
  If SammelIns Then Print #257, "  else"
  Print #257, "  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)"
  Print #257, "  If obMySQL and obmitAlterTab Then"
  Print #257, "   Set rs = dbcn.Execute(""show warnings"")"
  Print #257, "   If not rs.BOF() then"
  Print #257, "    If rs!code = 1265 Then"
  Print #257, "     Err.Raise -2147217833"
  Print #257, "    End If"
  Print #257, "   End If"
  Print #257, "  End If"
  If SammelIns Then Print #257, "  end if"
  Print #257, " Next i"
  Select Case LCase(Tbn(i))
   Case "formulare": iAnf = "rFo1"
   Case "forminhaltform_abk": iAnf = "rFi1"
   Case "unbekannte kennungen": iAnf = "rUn1"
   Case Else: iAnf = ""
  End Select
  If iAnf <> "" Then
   Print #257, " " + iAnf + " = ubound(" + Tbk(i) + ")"
  End If
  If SammelIns Then Print #257, " if obMySQl then if ubound(" & Tbk(i) & ")>0 then call dbcn.Execute(sql)', , adAsyncExecute)"
  Print #257, " exit function"
  Print #257, "fehler:"
  
  Print #257, "If Err.Number = -2147217833 Then"
  Print #257, " Dim rsc As ADODB.Recordset"
  If i <= TbZ1 Then
   Print #257, " dbcn.CommitTrans"
  End If
  Print #257, " Set rsc = New ADODB.Recordset"
  Print #257, " Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, """ & LCase(Tbn(i)) & """, Empty))"
  Print #257, " Do While Not rsc.EOF"
  Print #257, "  Select Case rsc!column_name"
  Set rsAdSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, LCase(Tbn(i)), Empty))
  Do While Not rsAdSc.EOF
   Select Case rsAdSc!data_type
    Case 8, 129, 130
     Print #257, "   Case """ & colW(rsAdSc!column_name) & """:  If SpMod(" & Tbk(i) & "(i)." & colW(rsAdSc!column_name) & ", rsc!character_maximum_length, """ & LCase(Tbn(i)) & """, """ & colW(rsAdSc!column_name) & """) Then Exit Do"
'     Print #257, "    If Len(" & Tbk(i) & "(i)." & colW(rsAdSc!column_name) & ") > rsc!character_maximum_length Then"
'     Print #257, "     Call dbcn.Execute(""Alter Table "" & kla & """ + LCase(Tbn(i)) + """ & klz & "" """ + " & iif(obMySQL,""MODIFY"",""ALTER"") & " + """ Column " & colW(rsAdSc!column_name) & " """ + " & iif(obMySQL,""VARCHAR"",""TEXT"") & " + """(""" & " & Len(" & Tbk(i) & "(i)." & colW(rsAdSc!column_name) & ") & "")"")"
'     Print #257, "     exit do"
'     Print #257, "    end if"
   End Select
   rsAdSc.Move 1
  Loop
  Print #257, "  End Select"
  Print #257, "  rsc.Move 1"
  Print #257, " Loop"
  If i <= TbZ1 Then
   Print #257, " dbcn.BeginTrans"
  End If
  Print #257, " if obmysql then Resume next else resume"
  Print #257, "End If"
  
  Print #257, " Dim AnwPfad$"
  Print #257, "#If VBA6 Then"
  Print #257, " AnwPfad = CurrentDb.Name"
  Print #257, "#Else"
  Print #257, " AnwPfad = App.Path"
  Print #257, "#End If"
  Print #257, " Select Case MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(IsNull(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""Aufgefangener Fehler in " + LCase(Replace(Tbn(i), " ", "_")) + "Speichern" + "/"" + AnwPfad)"
  Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): End"
  Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
  Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
  Print #257, " End Select"
  Print #257, "End Function ' " + LCase(Replace(Tbn(i), " ", "_")) + "Speichern"
 Next i
 Print #257, ""
 Print #257, "Public Function doSpeichern"
 Print #257, " on error goto fehler"
 For i = 0 To TbZ1
  Print #257, " call " + LCase(Replace(Tbn(i), " ", "_")) + "Speichern"
  Select Case LCase(Tbn(i))
  Case "faelle", "forminhaltform_abk"
  Print #257, "   if not obMySql then"
  Print #257, "    Call dbcn.CommitTrans"
  Print #257, "    Call dbcn.BeginTrans"
  Print #257, "   end if"
  End Select
 Next i
 Print #257, " exit function"
 Print #257, "fehler:"
 Print #257, " Dim AnwPfad$"
 Print #257, "#If VBA6 Then"
 Print #257, " AnwPfad = CurrentDb.Name"
 Print #257, "#Else"
 Print #257, " AnwPfad = App.Path"
 Print #257, "#End If"
 Print #257, " Select Case MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(IsNull(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""Aufgefangener Fehler in doSpeichern/"" + AnwPfad)"
 Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): End"
 Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
 Print #257, " End Select"

 Print #257, "End Function ' doSpeichern"
 Print #257, ""
 Print #257, "Function SpMod%(SpVal$, BishL&, TName$, SpName$)"
 Print #257, " If Len(SpVal) > BishL Then"
 Print #257, "  Call DBCn.Execute(""Alter Table "" & kla & TName & klz & "" "" & IIf(obMySQL, ""MODIFY"", ""ALTER"") & "" Column "" & SpName & "" "" & IIf(obMySQL, ""VARCHAR"", ""TEXT"") & ""("" & Len(SpVal) & "")"")"
 Print #257, "  SpMod = True"
 Print #257, " End If"
 Print #257, "End Function ' SpMod"
 
 Close #257
 Lese.Ausgabe = "Fertig mit MachTypen" + vbCrLf + frm.Ausgabe
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MacheTypen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' MacheTypen

Public Function HolReg(dlg As Dialog)
 Dim reg
 obStart = True
 dlg.obAcc = GetReg(1, RegStelle, "obAcc")
 reg = GetReg(1, RegStelle, "obMyQuelle")
 If reg = "" Then dlg.obMyQuelle = False Else dlg.obMyQuelle = reg
 reg = GetReg(1, RegStelle, "obMyQuelle1")
 If reg = "" Then dlg.obMyQuelle1 = False Else dlg.obMyQuelle1 = reg
 reg = GetReg(1, RegStelle, "obMyQuelle2")
 If reg = "" Then dlg.obMyQuelle2 = False Else dlg.obMyQuelle2 = reg
 If dlg.obAcc = 0 And dlg.obMyQuelle1 = 0 And dlg.obMyQuelle2 = 0 Then dlg.obMyQuelle = 1
 dlg.MdB = GetReg(1, RegStelle, "MdB")
 dlg.EmDatei = GetReg(1, RegStelle, "Email-Datei")
 obStart = False
' dlg.hfrm.Ziel = dlg.MdB
End Function ' HolReg(dlg As Dialog)
Public Function PutReg(dlg As Dialog)
 Call fStSpei(HCU, RegStelle, "obAcc", dlg.obAcc)
 Call fStSpei(HCU, RegStelle, "obMyQuelle", dlg.obMyQuelle)
 Call fStSpei(HCU, RegStelle, "obMyQuelle1", dlg.obMyQuelle1)
 Call fStSpei(HCU, RegStelle, "obMyQuelle2", dlg.obMyQuelle2)
 Call fStSpei(HCU, RegStelle, "MdB", dlg.MdB)
 Call fStSpei(HCU, RegStelle, "Email-Datei", dlg.EmDatei)
 dlg.hfrm.obAcc = dlg.obAcc
 dlg.hfrm.obMy = Not dlg.obAcc
 dlg.hfrm.Ziel = dlg.MdB
End Function ' PubReg
Public Function ConstrFestleg(ByVal Art%, Optional dlg As Dialog) ' dlg ist f¸r art= 0 und 1 nˆtig
 Const opti& = 2 + 4 + 8   ' 131118, 32 ' 1 + 2048 + 16384 + 131072
 On Error GoTo fehler
'ConStr$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;uid=praxis;pwd=sonne;option=" & opti
 Select Case Art
  Case 0
   If dlg.obAcc Then
    Art = 1
   Else
    If dlg.obMyQuelle1 Then
     Art = 3
    ElseIf dlg.obMyQuelle2 Then
     Art = 4
    ElseIf dlg.obMyQuelle Then
     Art = 2
    End If
   End If
 End Select
 If Art = 1 Then
 ' ConStr = CStrAcc + """" + dlg.MdB + """"
  If Not obStart Then
   If dlg.MdB = "" Then Call dlg.MdBFestleg
  End If
  ConStr = CStrAcc + dlg.MdB
  obMySQL = 0
 Else
  If Art = 3 Then
   myDB = "quelle1"
  ElseIf Art = 4 Then
   myDB = "quelle2"
  Else 'If dlg.obMyQuelle Then
   myDB = "quelle"
  End If
  ConStr$ = CStrMy & myDB & ";option=" & opti
  obMySQL = -1
 End If
 If Not obStart Then
' der ¸bern‰chste Befehl steht im regul‰ren Programmablauf nur hier
  If Not DBCn Is Nothing Then Set DBCn = Nothing
  DBCn.Open ConStr
'  Call MacheTypen
  If obMySQL Then
   Call DBCn.Execute("use " & myDB)
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ConstrFestleg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'ConstrFestleg(dlg As dialog)
Public Function Vergleiche(frm As Lese)
 Const SpBr% = 34
 Dim fld$(), i%, j%, n%, K%, cnA As New ADODB.Recordset, maxf%, cn As New ADODB.Connection
 For i = 0 To 2
  Call ConstrFestleg(i + 1, frm.dlg)
'  If i > 0 Then
   If Not cn Is Nothing Then Set cn = Nothing 'If cn.State = 1 Then cn.Close
   cn.Open ConStr
'  End If
  For j = 0 To TbZ2
   If i > 0 Then
'    On Error Resume Next
'    cn.Execute ("alter table " & LCase(Tbn(j)) & " change column abspos AbsPos integer unsigned")
'    On Error GoTo 0
   End If
   If Not cnA Is Nothing Then If cnA.State = 1 Then cnA.Close
   On Error Resume Next
   cnA.Open LCase(Tbn(j)), cn, adOpenStatic, adLockReadOnly
   If Err.Number <> 0 Then GoTo weiter
   On Error GoTo fehler
   If cnA.Fields.Count > maxf Then: maxf = cnA.Fields.Count: ReDim Preserve fld(2, TbZ2, 3, maxf)
   For K = 0 To cnA.Fields.Count - 1
    fld(i, j, 0, K) = cnA.Fields(K).Name
    fld(i, j, 1, K) = cnA.Fields(K).Type
    fld(i, j, 2, K) = cnA.Fields(K).Attributes
    fld(i, j, 3, K) = cnA.Fields(K).DefinedSize
   Next K
weiter:
  Next j
 Next i
 frm.Ausgabe = Left(frm.dlg.MdB + Space(SpBr), SpBr) + " " + Left("quelle" + Space(SpBr), SpBr) + " " + Left("quelle1" + Space(SpBr), SpBr) + " " + Left("quelle2" + Space(SpBr), SpBr) + vbCrLf + vbCrLf
  For j = 0 To TbZ2
   frm.Ausgabe = frm.Ausgabe & LCase(Tbn(j)) + vbCrLf
   For K = 0 To maxf
    frm.Ausgabe = frm.Ausgabe & _
    Left(Left(fld(0, j, 0, K) + Space(16), 15) & " " & Right(Space(3) & fld(0, j, 1, K), 3) & " " & fld(0, j, 2, K) & " " & Right(Space(3) & fld(0, j, 3, K), 3) & " " & Space(SpBr), SpBr) & _
    Left(Left(fld(1, j, 0, K) + Space(16), 15) & " " & Right(Space(3) & fld(1, j, 1, K), 3) & " " & fld(1, j, 2, K) & " " & Right(Space(3) & fld(1, j, 3, K), 3) & " " & Space(SpBr), SpBr) & _
    Left(Left(fld(2, j, 0, K) + Space(16), 15) & " " & Right(Space(3) & fld(2, j, 1, K), 3) & " " & fld(2, j, 2, K) & " " & Right(Space(3) & fld(2, j, 3, K), 3) & " " & Space(SpBr), SpBr) & vbCrLf
    If fld(0, j, 0, K) = "" And fld(1, j, 0, K) = "" And fld(2, j, 0, K) = "" Then GoTo weiter2
   Next K
weiter2:
  Next j
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Vergleiche/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Vergleiche

Function Zeichensatz$(Roh$)
Dim i%, ZS$, bq$, bz$
 On Error GoTo fehler
ZS = ""
For i = 1 To Len(Roh)
 bq = Mid(Roh, i, 1)
 Select Case bq
  Case "Ä"
   bz = "_"
  Case "Ø", "Ó"
   bz = "ﬂ"
  Case "ı"
   bz = "‰"
  Case "˜"
   bz = "ˆ"
  Case "≥"
   bz = "¸"
'  Case "-" ' f¸hrt bei den Labor-Langbezeichnungen zu schlechten Ergebnissen
'   bz = "ƒ"
  Case "Õ"
   bz = "÷"
  Case "_"
   bz = "‹"
  Case "¡"
   bz = "µ"
  Case "ﬂ"
   bz = "·"
  Case "”"
   bz = "‡"
  Case "⁄"
   bz = "È"
  Case "ﬁ"
   bz = "Ë"
  Case "›"
   bz = "Ì"
  Case "˝"
   bz = "Ï"
  Case "æ"
   bz = "Û"
  Case "="
   bz = "Ú"
  Case "∑"
   bz = "˙"
  Case "®"
   bz = "˘"
  Case "‘"
   bz = "‚"
  Case "€"
   bz = "Í"
  Case "Ø"
   bz = "Ó"
  Case "∂"
   bz = "Ù"
  Case "π"
   bz = "˚"
  Case "¡"
   bz = "µ"
  Case "¶"
   bz = "≤"
  Case "¶"
   bz = "≥"
  Case "∞"
   bz = "ß"
  Case "«"
   bz = "Ä"
  Case Else
   bz = bq
 End Select
 ZS = ZS + bz
Next i
Zeichensatz = ZS
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Zeichensatz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'Zeichensatz
Function ZSU1$(Roh$)
Dim i%, ZS$, bq$, bz$
 On Error GoTo fehler
ZS = ""
For i = 1 To Len(Roh)
 bq = Mid(Roh, i, 1)
 Select Case bq
  Case "·"
   bz = "ﬂ"
  Case "Ñ"
   bz = "‰"
  Case "î"
   bz = "ˆ"
  Case "Å"
   bz = "¸"
  Case "é"
   bz = "ƒ"
  Case "ô"
   bz = "÷"
  Case "ö"
   bz = "‹"
  Case "†"
   bz = "·"
  Case "Ö"
   bz = "‡"
  Case "Ç"
   bz = "È"
  Case "ä"
   bz = "Ë"
  Case "°"
   bz = "Ì"
  Case "ç"
   bz = "Ï"
  Case "¢"
   bz = "Û"
  Case "ï"
   bz = "Ú"
  Case "£"
   bz = "˙"
  Case "ó"
   bz = "˘"
  Case "É"
   bz = "‚"
  Case "à"
   bz = "Í"
  Case "å"
   bz = "Ó"
  Case "ì"
   bz = "Ù"
  Case "ñ"
   bz = "˚"
  Case "Ê"
   bz = "µ"
  Case "˝"
   bz = "≤"
  Case "¸"
   bz = "≥"
  Case "ı"
   bz = "ß"
'  Case "_" ' Ä wird zwar als _ dargestellt, aber _ auch
'   bz = "Ä"
  Case Else
   bz = bq
 End Select
 ZS = ZS + bz
Next i
ZSU1 = ZS
'If ZSU1 <> roh Then
' Debug.Print "ZSU: " + ZSU1 + " -> " + roh
'End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZSU1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZSU
#If False Then
' Stellt in der Datenbank programmst‰nde.Lauf den Wert Abbruch auf 'obabbruch', falls er nicht schon darauf steht und falls nicht "Nicht‰ndern" eingegeben wurde
' Ansonsten liefert die Funktion "true"
Public Function ProgrammLauf(obAbbruch%, Optional Cpt$, Optional Nichtƒndern%) As Boolean
 Dim i%, plConn As New ADODB.Connection, rsA As ADODB.Recordset, lcConStr$
 Dim cat As New ADOX.Catalog
 Dim Pfad$
 If DBCn.State = 1 Then
 On Error Resume Next
 For i = 1 To 6
' a) Verbindung aufbauen
  If obMySQL Then
   plConn.Open DBCn.ConnectionString
   plConn.Execute ("use programmstaende;")
  Else
   Pfad = Mid(DBCn, InStr(DBCn, "Source=") + 7)
   Pfad = Left(Pfad, InStr(Pfad, ";"))
   Do While Right(Pfad, 1) <> "\"
    Pfad = Left(Pfad, Len(Pfad) - 1)
   Loop
   lcConStr = ("Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & Pfad + "programmstaende" + ".mdb;" & _
               "Jet OLEDB:Engine Type=5")
   plConn.Open lcConStr
  End If
' b) falls nˆtig, Datenbank erstellen
  If Err.Number <> 0 Then
   If obMySQL Then
      Call plConn.Execute("create database programmstaende;")
   Else
      cat.Create lcConStr
#If False Then
      Dim WS, dbe
      Set dbe = New DAO.DBEngine
      Set WS = dbe.Workspaces(0)
      Call WS.CreateDatabase(Pfad + "programmstaende" + ".mdb", dbLangGeneral)
#End If
   End If
  Else ' err.number <> 0 nach connection
   Set rsA = New ADODB.Recordset
   Call rsA.Open("select Computer,Programm,Abbruch from Lauf where programm = ""DateiLese"";", plConn, adOpenDynamic, adLockOptimistic)
   If Err.Number = 0 And Not rsA.BOF Then
    Exit For
   Else ' Err.Number = 0 And Not rsA.BOF Then
    Err.Clear
    Call plConn.Execute("insert into Lauf (Programm) values (""DateiLese"");")
    If Err.Number = 0 Then
     Exit For
    Else '  Err.Number = 0 Then
     Err.Clear
     Call plConn.Execute("create table Lauf (Programm varchar(30), abbruch bit, Computer varchar(50))")
     If Err.Number <> 0 Then
      MsgBox "Unbehebbarer Fehler bei Programmlauf"
     End If
    End If ' Err.Number = 0 Then
   End If ' Err.Number = 0 And Not rsA.BOF Then
  End If ' err.number <> 0 nach connection
 Next i
 On Error GoTo fehler
 Cpt = IIf(IsNull(rsA!Computer), "", rsA!Computer)
 If rsA!abbruch = obAbbruch Then
  ProgrammLauf = True
 ElseIf Not Nichtƒndern Then
   If obAbbruch Then BrichAb = True
   rsA!Computer = CptName
   rsA!abbruch = obAbbruch
   rsA.Update
 End If
 End If ' dbcn.state
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ProgrammLauf/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ProgrammLauf(obAbbruch%, Optional Nichtƒndern%) As Boolean
#End If

Function adoxtest(dlg As Dialog)
 Dim i%, j&, K&, l&
 Dim cat As New ADOX.Catalog
 Dim Tb As New ADOX.Table
 For i = 2 To 2
  Call ConstrFestleg(i, dlg)
  cat.ActiveConnection = ConStr
  For K = 1 To cat.Tables.Count - 1
   For j = 1 To cat.Tables(K).Columns.Count - 1
    Debug.Print cat.Tables(K).Columns(j).Name, cat.Tables(K).Columns(j).Attributes, cat.Tables(K).Columns(j).Properties.Count
    For l = 0 To cat.Tables(K).Columns(j).Properties.Count - 1
     Debug.Print cat.Tables(K).Columns(j).Properties(l).Name, cat.Tables(K).Columns(j).Properties(l).Value
    Next l
'   Debug.Print cat.Tables(1).Columns(j).SortOrder
'    Debug.Print cat.Tables(k).Columns(j).Properties.Count
'    Debug.Print cat.Tables(k).Columns(j).Attributes
   Next j
  Next K
 Next
End Function ' adoxtest
Function getDokPfad$(Optional Abschnitt$)
 On Error GoTo fehler
 Dim iDt As TMIniDatei
 If Abschnitt <> "" Or IsNull(PcDokPfad) Or PcDokPfad = "" Then
  Set iDt = New TMIniDatei
  If IsNull(PcDokPfad) Or PcDokPfad = "" Then
   PcDokPfad = iDt.GetProp("Verzeichnisse/TurboMed/Dokumente", "Pfad")
   getDokPfad = PcDokPfad
  End If
  If Abschnitt <> "" Then
   getDokPfad = iDt.GetProp("Verzeichnisse/TurboMed/" + Abschnitt, "Pfad")
  End If
 Else
  getDokPfad = PcDokPfad
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getDokPfad/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getTMExeV

Public Function key(KeyCode%, Shift%, frm As Form)
 On Error GoTo fehler
 If KeyCode = 27 Then
    Call HolReg(frm)
    frm.Hide
    If frm.Name = "Lese" Then End
 End If
' If KeyCode = 33 Then Call doR¸ckw‰rtsCmd(frm)
' If KeyCode = 34 Then Call doVorw‰rtsCmd(frm) <- stellt den aktuellen Feldinhalt falsch ein!
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in key/" + App.Path)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' key

