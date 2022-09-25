VERSION 5.00
Begin VB.Form PatAuswahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Patientenauswahl"
   ClientHeight    =   11310
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11310
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox FaellepHA 
      Height          =   5520
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   12135
   End
   Begin VB.ComboBox HAAusw 
      Height          =   315
      Left            =   480
      TabIndex        =   11
      Top             =   5280
      Width           =   11055
   End
   Begin VB.TextBox Hausarzt 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   4950
      Width           =   11775
   End
   Begin VB.ComboBox Pat_id 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   6015
   End
   Begin VB.ListBox Faelle 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   12135
   End
   Begin VB.TextBox pat_idDaten 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   10815
   End
   Begin VB.ComboBox PatName 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   11040
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   11040
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Anzahl 
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Zahl 
      Height          =   255
      Left            =   11640
      TabIndex        =   13
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label HA_Bez 
      Caption         =   "&HA:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4980
      Width           =   375
   End
   Begin VB.Label entspricht 
      Caption         =   "entspricht"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "&Eingabe Pat_ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Auswahl mit F4:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PatAuswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public hlese As Lese
Public PatID&
Private haGewählt%
#Const EinmalDB = True

Private Sub CancelButton_Click()
 Call Key(27, 0, Me)
End Sub ' CancelButton_Click

Private Sub CancelButton_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
End Sub ' CancelButton_KeyDown

Private Sub Faelle_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
End Sub ' Faelle_KeyDown

Public Sub FaellepHA_Click()
If Me.FaellepHA.ListCount > 0 Then
 Dim Inx&
 Inx = Me.FaellepHA.ListIndex
 If Inx < 0 Then Inx = 0
 If Inx > Me.FaellepHA.ListCount Then Inx = Me.FaellepHA.ListCount
 Me.Pat_id = Left(Me.FaellepHA.List(Inx), InStr(Me.FaellepHA.List(Inx), " ") - 1)
End If
End Sub ' FaellepHA_Click

Private Sub FaellepHA_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
End Sub ' FaellepHA_KeyDown

Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
End Sub ' Form_KeyDown

Private Sub HAAusw_GotFocus()
 If Not haGewählt Then Call AuswHA(Me)
End Sub

Private Sub HAAusw_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
End Sub ' HAAusw_KeyDown

Private Sub Hausarzt_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
End Sub ' Hausarzt_KeyDown

Private Sub OKButton_Click()
 Me.Visible = False
 Me.hlese.los
End Sub ' OKButton_Click

Private Sub OKButton_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
End Sub ' OKButton_KeyDown
Private Sub Pat_ID_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
End Sub ' Pat_ID_KeyDown
Private Sub Form_Load()
 On Error GoTo fehler
  If Me.Visible Then Screen.MousePointer = vbHourglass
  DoEvents
  Me.Caption = "Patientenauswahl " & IIf(lies.obMySQL, "MySQL: " & hlese.MyDB, Me.hlese.dlg.MdB)
  Call AuswName(Me)
'  Call AuswHA(Me)
#If EinmalDB Then
#Else
  Call AuswPat_id(Me)
#End If
  Me.Faelle.Clear
  If Not lies.obMySQL Then
   Call KVAccSuch
  End If
'  Me.hlese.MousePointer = vbDefault
 Screen.MousePointer = vbDefault
 DoEvents
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in form_load/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Form_Load
' kommt vor in doPatNameChange, SchreibDatensatz und FDC_indvorWechsel
Function getPat_id&(PatName$)
 Dim Spl$(), sql$
 Dim rNaA As New ADODB.Recordset
 If IsNull(PatName) Or PatName = "" Then
  getPat_id = -1
 Else
  Spl = Split(PatName, "|")
  If UBound(Spl) < 1 Then
   getPat_id = -1
  Else
   If Not IsNumeric(Spl(1)) Then
    getPat_id = -1
   Else
    getPat_id = CLng(Spl(1))
   End If
  End If
 End If
 If getPat_id = -1 And Me.PatName <> "" Then
  If IsNumeric(Me.PatName) Then
   sql = "select nachname, vorname, gebdat, pat_id from namen where pat_id like '" & Me.PatName & "%'"
  Else
   sql = "select nachname, vorname, gebdat, pat_id from namen where concat(nachname¡ ', '¡ vorname) like '" & Me.PatName & "%'"
  End If
  sql = SqlU(sql, hlese.obMySQL) ' 11.7.09
'  If hlese.obMySQL Then
'   SQL = replace$(SQL, "¡", ",")
'  Else
'   SQL = replace$(replace$(SQL, "concat", ""), "¡", " & ")
'  End If
  rNaA.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rNaA.EOF Then
   getPat_id = rNaA!Pat_id
   PatID = getPat_id
  End If
 End If
End Function ' getPat_id
Function AuswName(frm As PatAuswahl)
 Dim zwi$, i&, t$
 syscmd 4, "Wähle Patientien aus"
 DoEvents
 #If EinmalDB Then
 Dim sEl As SortierPat_IDText, sLi As New SortierListe
 #End If
 On Error GoTo fehler
 Dim rNaA As New ADODB.Recordset
' rNaA.Open "select nachname, vorname, gebdat, pat_id from namen order by nachname, vorname, gebdat;", dbcn, adOpenStatic, adLockReadOnly
 #If EinmalDB Then
 sql = "select pat_id, cast(concat(cast(nachname as char)¡ "", ""¡ vorname¡ "",*""¡ date_format(gebdat,'%e.%c.%Y')¡ "" | ""¡ pat_id) as char) as f0, concat(pat_id¡"" | ""¡ cast(nachname as char)¡ "", ""¡ vorname¡ "",*""¡ date_format(gebdat,'%e.%c.%Y')) as f1 from namen order by nachname, vorname, gebdat;"
 #Else
 sql = "select cast(concat(cast(nachname as char)¡ "", ""¡ vorname¡ "",*""¡ date_format(gebdat,'%e.%c.%Y')¡ "" | ""¡ pat_id) as char) as f0 from namen order by nachname, vorname, gebdat;"
 #End If
 If Forms(0).obMySQL Then
  sql = Replace$(sql, "¡", ",")
 Else
  sql = Replace$(Replace$(Replace$(Replace$(sql, "¡", " & "), "concat", ""), "date_format", "format"), "'%e.%c.%Y'", """short date""")
  sql = Replace$(Replace$(sql, "cast(", ""), " as char)", "")
 End If
 rNaA.Open sql, DBCn, adOpenStatic, adLockReadOnly
 zwi = frm.PatName
 frm.PatName.Clear
 frm.PatName = zwi
 Do While Not rNaA.EOF
  t = rNaA!F0 ' rNaA!NachName + ", " & rNaA!VorName + ",*" + format$(rNaA!GebDat, "D.M.YY") & " | " + CStr(rNaA!Pat_id)
  frm!PatName.AddItem t
  i = i + 1
  If i = 10 Then
   syscmd 4, "Pat: " & rNaA!F0 'rNaA!NachName & " " & rNaA!VorName
   DoEvents
   i = 0
  End If
  #If EinmalDB Then
  Set sEl = New SortierPat_IDText
  sEl.Pat_id = -rNaA!Pat_id
  sEl.Text = rNaA!f1 'CStr(rNaA!Pat_id) & " | " & rNaA!NachName & ", " & rNaA!VorName + ",*" + format$(rNaA!GebDat, "D.M.YY")
  Call sLi.sCAdd(sEl)
  #End If
  rNaA.Move 1
 Loop
 #If EinmalDB Then
 frm.Pat_id.Clear
 For i = 1 To sLi.Count
  frm.Pat_id.AddItem sLi.Item(i).Text
 Next i
 #End If
 syscmd 5
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AuswName/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AuswName
Function AuswPat_id(frm As PatAuswahl)
 Dim zwi$, i&
 On Error GoTo fehler
 'Dim cnADO As New ADODB.Connection
 'cnADO.Open ConStr
 Dim rNaA As New ADODB.Recordset
 sql = "select concat(pat_id¡"" | ""¡ nachname¡ "", ""¡ vorname¡ "",*""¡ date_format(gebdat,'%e.%c.%Y')) as f1 from namen order by pat_id;"
 If Forms(0).obMySQL Then
  sql = Replace$(sql, "¡", ",")
 Else
  sql = Replace$(Replace$(Replace$(Replace$(sql, "¡", " & "), "concat", ""), "date_format", "format"), "'%e.%c.%Y'", """short date""")
 End If
 rNaA.Open sql, DBCn, adOpenStatic, adLockReadOnly
 zwi = frm.Pat_id
 frm.Pat_id.Clear
 frm.Pat_id = zwi
 Dim t$
 Do While Not rNaA.EOF
  t = rNaA!f1 ' CStr(rNaA!Pat_id) & " | " & rNaA!NachName & ", " & rNaA!VorName + ",*" + format$(rNaA!GebDat, "D.M.YY")
  frm.Pat_id.AddItem t
  i = i + 1
  If i = 10 Then
   syscmd 4, "Pat: " & rNaA!f1 'rNaA!NachName & " " & rNaA!VorName
   DoEvents
   i = 0
  End If
  rNaA.Move 1
 Loop
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Auswpat_id/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AuswPat_id
Function repl(q$, r$, z$)
 repl = Replace$(q, r, z)
End Function
Function AuswHA(frm As PatAuswahl)
 Dim zwi$, sql$, i%, hier%, ItemStr$
 Dim kvPfad$, p1&, p2&
 On Error GoTo fehler
 syscmd 4, "Hausarztauswahl ..."
 DoEvents
 'Dim cnADO As New ADODB.Connection
 'cnADO.Open ConStr
 Dim rHS As New ADODB.Recordset, rKVNr As New ADODB.Recordset
 Call rKVNr.Open("select distinct kvnr from namen union select distinct übw from faelle", DBCn, adOpenStatic, adLockReadOnly)
 If Not Forms(0).obMySQL Then
   Call acon(HaT)
   If Not Forms(0).obMySQL Then
    p1 = InStr(HAECn, "Data Source=")
    If p1 > 0 Then
     p1 = p1 + 12
     p2 = InStr(p1, HAECn, ";")
     kvPfad = Mid$(HAECn, p1, p2 - p1)
    End If
   End If
 End If
 Do While Not rKVNr.EOF
  If lies.obMySQL Or kvPfad <> "" Then
'   sql = "select dbnr, bstelle, anrede, """" as überschrift, haname, ort, kvnu as kvnr, tel1, tel2, tel3, tel4, fax1, fax1k,fax2,fax2k, fax3,fax3k, email, zulg, arzttyp, gemmit, beme, -dmpt2 as j_dmpt2, -dmpt1 as j_dmpt1, geschlecht, titel, vorname, nachname, straße, plz from kvaerzte.hae where kvnr = '" & left(rKVNr!KVNr, 2) & "/" & mid$(rKVNr!KVNr, 3, 5) & "' " & _
         "union select dbnr, bstelle, anrede, """" as überschrift, haname, ort, kvnu as kvnr, tel1, tel2, tel3, tel4, fax1, fax1k,fax2,fax2k, fax3,fax3k, email, zulg, arzttyp, gemmit, beme, -dmpt2 as j_dmpt2, -dmpt1 as j_dmpt1, geschlecht, titel, vorname, nachname, straße, plz from kvaerzte.haealt where kvnr = '" & left(rKVNr!KVNr, 2) & "/" & mid$(rKVNr!KVNr, 3, 5) & "' " & _
         "union select id as dbnr, """" as bstelle, """" as anrede, überschrift, name as haname, ort, kvnr, telefon as tel1, """" as tel2, """" as tel3, """" as tel4, telefax as fax1, replace$(replace$(replace$(telefax,""/"",""""),""-"",""""),"" "","""") as fax1k,"""" as fax2,"""" as fax2k, """" as fax3, """" as fax3k, e_mail as email, zulassungsgebiet as zulg, arzttyp, `gemeinschaftspraxis mit` as gemmit, beme, dmpt2, dmpt2, geschlecht, titel, vorname, nachname, straße, plz from hausaerzte where kvnr = '" & rKVNr!KVNr & "' " & _
         "union select """" as dbnr, """" as bstelle, """" as anrede, """" as überschrift, name as haname, ort, kvnr, telefon as tel1, """" as tel2, """" as tel3, """" as tel4, """" as fax1, fax as fax1k,"""" as fax2,"""" as fax2k, """" as fax3, """" as fax3k, """" as email, fachgruppe as zulg, """" as arzttyp, """" as gemmit, """" as beme, -2 as dmpt2, -2 as dmpt1, 0 as geschlecht, titel, vorname, name as nachname, strasse, plz from listenausgabeuew where kvnr = '" & rKVNr!KVNr & "'"
    sql = "select dbnr, bstelle, anrede, """" as überschrift, haname, ort, kvnu as kvnr, tel1, tel2, tel3, tel4, fax1, fax1k,fax2,fax2k, fax3,fax3k, email, zulg, arzttyp, gemmit, beme, dmpt2, dmpt1, geschlecht, titel, vorname, nachname, straße, plz from kvaerzte.hae where kvnr = '" & Left(rKVNr!KVNr, 2) & "/" & Mid$(rKVNr!KVNr, 3, 5) & "' " & _
         "union select dbnr, bstelle, anrede, """" as überschrift, haname, ort, kvnu as kvnr, tel1, tel2, tel3, tel4, fax1, fax1k,fax2,fax2k, fax3,fax3k, email, zulg, arzttyp, gemmit, beme, dmpt2, dmpt1, geschlecht, titel, vorname, nachname, straße, plz from kvaerzte.haealt where kvnr = '" & Left(rKVNr!KVNr, 2) & "/" & Mid$(rKVNr!KVNr, 3, 5) & "' " & _
         "union select id as dbnr, """" as bstelle, """" as anrede, überschrift, name as haname, ort, kvnr, telefon as tel1, """" as tel2, """" as tel3, """" as tel4, telefax as fax1, replace$(replace$(replace$(telefax,""/"",""""),""-"",""""),"" "","""") as fax1k,"""" as fax2,"""" as fax2k, """" as fax3, """" as fax3k, e_mail as email, zulassungsgebiet as zulg, arzttyp, `gemeinschaftspraxis mit` as gemmit, beme, dmpt2, dmpt2, geschlecht, titel, vorname, nachname, straße, plz from hausaerzte where kvnr = '" & rKVNr!KVNr & "' " & _
         "union select """" as dbnr, """" as bstelle, """" as anrede, """" as überschrift, name as haname, ort, kvnr, telefon as tel1, """" as tel2, """" as tel3, """" as tel4, """" as fax1, fax as fax1k,"""" as fax2,"""" as fax2k, """" as fax3, """" as fax3k, """" as email, fachgruppe as zulg, """" as arzttyp, """" as gemmit, """" as beme, -2 as dmpt2, -2 as dmpt1, 0 as geschlecht, titel, vorname, name as nachname, strasse, plz from listenausgabeuew where kvnr = '" & rKVNr!KVNr & "'"
   If Not Forms(0).obMySQL Then
    sql = Replace$(sql, "kvaerzte.", kla & kvPfad & klz & ".")
    sql = Replace$(sql, "replace$(replace$(replace$(telefax,""/"",""""),""-"",""""),"" "","""")", "telefax")
   End If
   Set rHS = Nothing
   rHS.Open sql, DBCn, adOpenStatic, adLockReadOnly
   If Not rHS.BOF Then
    ItemStr = rHS!Nachname & ", " & rHS!Titel & ", " & rHS!Vorname & ", KVNr:" & rHS!KVNr & ", T." & rHS!tel1 & " F." & rHS!fax1 & ", " & rHS!Straße & ", " & rHS!Plz & " " & rHS!Ort
    For i = 0 To frm.HAAusw.ListCount
     If frm.HAAusw.List(i) > ItemStr Then
      Call frm.HAAusw.AddItem(ItemStr, i)
      GoTo hamma
     End If
    Next i
    Call frm.HAAusw.AddItem(ItemStr)
hamma:
    syscmd 4, "Hausarzt: " & rHS!haname
   End If
  End If
  rKVNr.Move 1
  DoEvents
 Loop
 haGewählt = True
 syscmd 5
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AuswHA/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AuswPat_id

Private Sub HAAusw_Click()
 Call HAAusw_Change
End Sub
Private Sub HAAusw_Change()
 Dim rNaA As New ADODB.Recordset
 Dim sql$, KVNr$
 KVNr = Mid$(Me.HAAusw, InStr(Me.HAAusw, "KVNr") + 5, 7)
 sql = "select distinct * from(" & _
       "select namen.pat_id, quartal, bhfb, namen.nachname, namen.vorname, namen.gebdat  from namen left join (select pat_id, min(quartal) as quartal, min(bhfb) as bhfb from (select pat_id, quartal, bhfb from faelle order by bhfb desc) as innen group by pat_id) as minnen on minnen.pat_id = namen.pat_id where kvnr = '" & KVNr & "' union " & _
       "select  namen.pat_id, quartal, bhfb, namen.nachname, namen.vorname, namen.gebdat  from (select pat_id, min(quartal) as quartal, min(bhfb) as bhfb, min(übw) as übw from (select pat_id, quartal, bhfb, übw from faelle order by bhfb desc) as innen group by pat_id) as minnen left join namen on minnen.pat_id = namen.pat_id where übw = '" & KVNr & "') as in1 order by pat_id desc"

  If DBCn.ConnectionString = "" Then
'   Call Me.hlese.ConstrFestleg(0, Me.hlese)
   Call acon(quelleT)
  End If
Me.FaellepHA.Clear
rNaA.Open sql, DBCn, adOpenStatic, adLockReadOnly
Me.Zahl = 0
Do While Not rNaA.EOF
 Me.Zahl = Me.Zahl + 1
 Me.FaellepHA.AddItem rNaA!Pat_id & " " & rNaA!Quartal & " " & rNaA!Nachname & " " & rNaA!Vorname & " " & rNaA!GebDat
 rNaA.Move 1
Loop
Call Me.FaellepHA_Click
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in HAAusw_Change/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' haausw_change

Private Sub Pat_ID_Change()
 Dim rNaA As New ADODB.Recordset
 Dim rFaA As New ADODB.Recordset
 Dim HAStr$, obHA%
 Static rKVA As New ADODB.Recordset
' Call KVÄVorb
  Call acon(HaT)
 Me.MousePointer = vbHourglass
 Me.Faelle.Clear
 Me.Hausarzt = ""
 If Me.Pat_id <> "" Then
  On Error Resume Next
  Me.PatID = getPid(Me.Pat_id)
  On Error GoTo fehler
  If Me.PatID = 0 Then Exit Sub
  If DBCn.ConnectionString = "" Then
'   Call Me.hlese.ConstrFestleg(0, Me.hlese)
   Call acon(quelleT)
  End If
  rNaA.Open "select nachname, vorname, gebdat, pat_id from namen where pat_id = " & Me.PatID & " order by nachname, vorname, gebdat", DBCn, adOpenStatic, adLockReadOnly
  If rNaA.EOF Then
   Me.pat_idDaten = ""
  Else
   Me.pat_idDaten = rNaA!Nachname & ", " & rNaA!Vorname + ",*" + Format$(rNaA!GebDat, "D.M.YY") + " | " + CStr(rNaA!Pat_id)
  End If
' Call KVÄVorb
  Call acon(HaT)
  rFaA.Open "select * from faelle left join (select distinct vk, kateg, name from kassenliste) as kl on faelle.vknr = kl.vk where pat_id = " & Me.PatID & " order by bhfb desc", DBCn, adOpenDynamic, adLockReadOnly
  Do While Not rFaA.EOF
   HAStr = ""
   If Not rFaA.BOF Then
    If Not rFaA!Übw = "" Then
     Set rKVA = Nothing
'     rKVA.Open "select -dmpt1 as j_dmpt1, -dmpt2 as j_dmpt2, hae.* from hae where kvnr = '" & left(rFaA!Übw, 2) & "/" & right$(rFaA!Übw, 5) & "'", HAECn, adOpenDynamic, adLockReadOnly
     rKVA.Open "select * from hae where kvnr = '" & Left(rFaA!Übw, 2) & "/" & Right$(rFaA!Übw, 5) & "'", HAECn, adOpenDynamic, adLockReadOnly
     If Not rKVA.BOF And Not obHA And Not rFaA!Übw = "6419153" Then
      HAStr = rKVA!haname & ", " & rKVA!Ort & IIf(Not IsNull(rKVA!Email) And rKVA!Email <> "", " Email: " & rKVA!Email, "") & _
              "   Tel: " & Replace$(IIf(IsNull(rKVA!tel1), "", rKVA!tel1), " ", "") & "  Fax: " & IIf(IsNull(rKVA!fax1k), "", rKVA!fax1k)
      Dim obDMP2%, obDMP1%
      Do While Not rKVA.EOF
'       If rKVA!j_dmpt2 <> 0 Then obDMP2 = True
       If rKVA!dmpt2 <> 0 Then obDMP2 = True
'       If rKVA!j_dmpt1 <> 0 Then obDMP1 = True
       If rKVA!dmpt1 <> 0 Then obDMP1 = True
       rKVA.Move 1
      Loop
      HAStr = HAStr & "  DMP 2: " & IIf(obDMP2, "+", "-") & "  1: " & IIf(obDMP1, "+", "-")
      Me.Hausarzt = HAStr
      obHA = -1
     End If
    End If
   End If
   Me.Faelle.AddItem (rFaA!BhFB & " - " & rFaA!BhFE1 & " " & rFaA!SchGr & "  " & Left(IIf(IsNull(rFaA!Name), "", rFaA!Name), 20) & " " & IIf(IsNull(rFaA!Kateg), "", rFaA!Kateg) & "  " & rFaA!ÜWNNr & " " & rFaA!ÜWNaN & " " & rFaA!ÜWVor)
   rFaA.Move 1
  Loop
 End If
 Me.MousePointer = vbDefault
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Pat_ID_Change()/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Pat_ID_Change()
Public Function getPid(Text$)
 Dim pos&
 pos = InStr(Text, " ")
 If pos > 0 Then
  getPid = Left(Me.Pat_id, pos - 1)
 Else
  getPid = Me.Pat_id
 End If
End Function

Private Sub Pat_id_Click()
 Call Pat_ID_Change
End Sub

Private Sub PatName_Click()
 Me.Pat_id = getPat_id(Me.PatName)
End Sub


Private Sub PatName_Change()
 Dim rs As New ADODB.Recordset
 Me.Pat_id = getPat_id(Me.PatName)
 rs.Open "select count(pat_id) as ct from namen where concat(nachname, ', ',vorname) like '" & Me.PatName & "%'", DBCn, adOpenStatic, adLockReadOnly
 Me.Anzahl = rs!ct
End Sub ' PatName_Change()

Private Sub PatName_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me)
 Me.Pat_id = getPat_id(Me.PatName)
End Sub

