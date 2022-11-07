VERSION 5.00
Begin VB.Form PatAuswahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Patientenauswahl"
   ClientHeight    =   14370
   ClientLeft      =   3840
   ClientTop       =   435
   ClientWidth     =   12285
   LinkTopic       =   "PatAuswahl"
   MinButton       =   0   'False
   ScaleHeight     =   14370
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Therapiearten 
      Caption         =   "&Therartenanz."
      Height          =   315
      Left            =   7440
      TabIndex        =   31
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Zeilenzahl 
      Height          =   285
      Left            =   7920
      TabIndex        =   30
      Text            =   "8"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton VorDat 
      Caption         =   "Vor&Dat.ermit"
      Height          =   255
      Left            =   7440
      TabIndex        =   28
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Abr 
      Caption         =   "A&br"
      Height          =   375
      Left            =   6840
      TabIndex        =   26
      Top             =   600
      Width           =   495
   End
   Begin VB.ListBox Vorbriefe 
      Height          =   2400
      Left            =   360
      TabIndex        =   24
      Top             =   4680
      Width           =   11895
   End
   Begin VB.CommandButton Patientenlaufzettel 
      Caption         =   "&PLZ"
      Height          =   375
      Left            =   6290
      TabIndex        =   23
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton inTm 
      Caption         =   "T&urb"
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Angeforderte 
      Caption         =   "Angeforderte Briefe &gs"
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   21
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Angeforderte 
      Caption         =   "Angeforderte Briefe &tk"
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   20
      Top             =   0
      Width           =   2295
   End
   Begin VB.ListBox Diagnosen 
      Height          =   2790
      Left            =   7560
      TabIndex        =   17
      Top             =   1800
      Width           =   4695
   End
   Begin VB.ComboBox Vorlage 
      Height          =   315
      Left            =   9600
      TabIndex        =   16
      Top             =   600
      Width           =   2535
   End
   Begin VB.ListBox FaellepHA 
      Height          =   6495
      Left            =   120
      TabIndex        =   13
      Top             =   7800
      Width           =   12135
   End
   Begin VB.ComboBox HAAusw 
      Height          =   315
      Left            =   360
      TabIndex        =   12
      Top             =   7440
      Width           =   11055
   End
   Begin VB.TextBox Hausarzt 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   7125
      Width           =   11775
   End
   Begin VB.ComboBox Pat_id 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.ListBox Faelle 
      Height          =   2985
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox pat_idDaten 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   10815
   End
   Begin VB.ComboBox PatName 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   11330
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   10920
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label ZeilenzahlL 
      Caption         =   "&Zeilen"
      Height          =   255
      Left            =   7200
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label keinBericht 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11040
      TabIndex        =   27
      Top             =   600
      Width           =   975
   End
   Begin VB.Label nötig 
      BackColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Arzt 
      BackColor       =   &H8000000B&
      Height          =   255
      Left            =   8040
      TabIndex        =   19
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label ArztLbl 
      Caption         =   "Arzt:"
      Height          =   255
      Left            =   7560
      TabIndex        =   18
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label VorlageLab 
      Caption         =   "&Vorlage:"
      Height          =   255
      Left            =   9000
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Anzahl 
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Zahl 
      Height          =   255
      Left            =   11640
      TabIndex        =   14
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label HA_Bez 
      Caption         =   "&HA:"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label entspricht 
      Caption         =   "entspricht"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
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
Const HADBName$ = "haerzte"
Private haGewählt%, VorlageGewählt%, PatGewählt%, PatZuWählen%, Pat_IDGewählt%
Dim RegPos$, cR As New Registry
#Const EinmalDB = True
Dim Zulp As Date ' zuletzt - letzter Zeitpunkt
Dim lbeh% ' letzter Behandler: 1 = Kothny, 0 = unentschieden, 1 = Schade
Public VorBriefID& ' letzte Pat_id, zu der ein Vorbrief rausgesucht wurde
Public obRueck%
Public geladen%
Public briefneu%

Private Sub Angeforderte_Click(Index As Integer)
 Dim rs As New ADODB.Recordset
 Dim arztbed$
' SELECT CASE Index
'  Case 0: arztbed = "(e.art IN ('tk','ARCHIE2','APK') OR e.inhalt LIKE '%(tk)%')"
'  Case 1: arztbed = "(e.art IN ('gs','doppler','duplex') OR e.inhalt LIKE '%(gs)%')"
' END SELECT
 DoCmd.Maximize
 Me.Faelle.Clear
 Me.Diagnosen.Clear
 Me.FaellepHA.Clear
 Me.Vorbriefe.Clear
' myFrag rs, "SELECT f.pat_id, LEFT(CONCAT(n.nachname,',',n.vorname,IF(n.titel='','',','),n.titel,IF(n.nvorsatz='','',' '),n.nvorsatz),24) AS name, DATE_FORMAT(f.bhfb,'%d.%m.%y') BhFB, IF(e.art IN ('gs','doppler','duplex') OR e.inhalt LIKE '%(gs)%','gs','tk') arzt, f.auftrag, f.verdacht, f.befund FROM faelle f LEFT JOIN namen n ON f.pat_id = n.pat_id LEFT JOIN briefe b ON f.pat_id = b.pat_id AND name LIKE '%brief%' AND b.zeitpunkt > f.bhfb LEFT JOIN eintraege e ON e.pat_id = f.pat_id AND " & arztbed & " WHERE (auftrag LIKE '%beten%' OR auftrag LIKE '%bitte%' OR verdacht LIKE '%beten%' OR verdacht LIKE '%bitte%' OR befund LIKE '%beten%' OR befund LIKE '%bitte%') AND ISNULL(b.name) GROUP BY f.pat_id ORDER BY arzt, f.bhfb, f.pat_id, e.zeitpunkt DESC"
' myFrag rs, "SELECT f.pat_id, LEFT(CONCAT(n.nachname,',',n.vorname,IF(n.titel='','',','),n.titel,IF(n.nvorsatz='','',' '),n.nvorsatz),24) AS name, DATE_FORMAT(f.bhfb,'%d.%m.%y') BhFB, IF(e.art IN ('gs','doppler','duplex') OR e.inhalt LIKE '%(gs)%','gs','tk') arzt, f.auftrag, f.verdacht, f.befund FROM faelle f LEFT JOIN namen n ON f.pat_id = n.pat_id LEFT JOIN briefe b ON f.pat_id = b.pat_id AND name LIKE '%brief%' AND b.zeitpunkt > f.bhfb LEFT JOIN eintraege e ON e.pat_id = f.pat_id AND " & arztbed & " LEFT JOIN `faxeinp`.`outa` o ON o.pid = f.pat_id AND o.transe > f.bhfb AND o.docname LIKE '%brief%' WHERE (auftrag LIKE '%beten%' OR auftrag LIKE '%bitte%' OR verdacht LIKE '%beten%' OR verdacht LIKE '%bitte%' OR befund LIKE '%beten%' OR befund LIKE '%bitte%') AND ISNULL(b.name) AND ISNULL(o.transe) GROUP BY f.pat_id ORDER BY f.pat_id, e.zeitpunkt DESC"
myFrag rs, "SELECT * FROM (SELECT f.pat_id, LEFT(CONCAT(n.nachname,',',n.vorname,IF(n.titel='','',','),n.titel,IF(n.nvorsatz='','',' '),n.nvorsatz),24) AS name, DATE_FORMAT(f.bhfb,'%d.%m.%y') BhFB, (SELECT COUNT(0) FROM eintraege e WHERE pat_id = f.pat_id AND (e.art IN ('gs','doppler','duplex') OR e.inhalt LIKE '%(gs)%')) gsz, (SELECT MAX(zeitpunkt) FROM eintraege e WHERE pat_id = f.pat_id AND (e.art IN ('gs','doppler','duplex') OR e.inhalt LIKE '%(gs)%')) gsl, (SELECT COUNT(0) FROM eintraege e WHERE pat_id = f.pat_id AND (e.art IN ('tk') OR e.inhalt LIKE '%(tk)%')) tkz, (SELECT MAX(zeitpunkt) FROM eintraege e WHERE pat_id = f.pat_id AND (e.art IN ('tk') OR e.inhalt LIKE '%(tk)%')) tkl," & _
        "f.auftrag, f.verdacht, f.befund FROM faelle f LEFT JOIN namen n ON f.pat_id = n.pat_id LEFT JOIN briefe b ON f.pat_id = b.pat_id AND name LIKE '%brief%' AND b.zeitpunkt > f.bhfb LEFT JOIN `faxeinp`.`outa` o ON o.pid = f.pat_id AND o.transe > f.bhfb AND o.docname LIKE '%brief% ' WHERE (auftrag LIKE '%beten%' OR auftrag LIKE '%bitte%' OR verdacht LIKE '%beten%' OR verdacht LIKE '%bitte%' OR befund LIKE '%beten%' OR befund LIKE '%bitte%') AND ISNULL(b.name) AND ISNULL(o.transe) GROUP BY f.pat_id) i WHERE " & IIf(Index = 0, "tkz", "gsz") & "<>0 ORDER BY pat_id"
 If Not rs.BOF Then
  Do While Not rs.EOF
   Me.Vorbriefe.AddItem rs!Pat_id & Space$(1.7 * (5 - Len(rs!Pat_id))) & " " & Left$(rs!name, 20) & Space$(IIf(Len(rs!name) >= 20, 1, 1.5 * (20 - Len(rs!name)))) & "   (" & IIf(rs!gsz <> 0, IIf(rs!gsl > rs!tkl Or (rs!gsz > (3 * rs!Tkz)), " GS: ", " gs: ") & rs!gsz & " K.,zul.: " & Format(rs!gsl, "Dd.mm.yy") & IIf(rs!Tkz <> 0, ", ", ""), Space$(40)) & IIf(rs!Tkz <> 0, IIf(rs!tkl > rs!gsl Or (rs!Tkz > (3 * rs!gsz)), "  TK: ", "  tk: ") & rs!Tkz & " K.,zul.: " & Format(rs!tkl, "Dd.mm.yy"), Space$(40)) & ")     " & rs!BhFB & "          " & rs!Auftrag & " " & rs!Verdacht & " " & rs!Befund
   rs.MoveNext
  Loop
 End If ' Not rs.BOF
End Sub ' Angeforderte_click

Private Sub CancelButton_Click()
 Call Key(27, 0, Me)
End Sub ' CancelButton_Click

Private Sub CancelButton_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' CancelButton_KeyDown

Private Sub Faelle_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
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

Private Sub FaellepHA_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' FaellepHA_KeyDown

Private Sub Form_Activate()
 Me.Pat_id.SetFocus
 If Me.obRueck Then Me.Patientenlaufzettel.FontItalic = True Else Me.Patientenlaufzettel.FontItalic = False
 On Error Resume Next
' IF Me.hlese.Visible THEN Me.hlese.Hide
End Sub ' Form_Activate()

Private Sub Form_Deactivate()
 On Error Resume Next
 If Not Me.hlese.Visible Then Me.hlese.Show
End Sub ' Form_Deactivate()

Private Sub Form_GotFocus()
 Me.Pat_id.SetFocus
End Sub ' Form_GotFocus()

Private Sub Form_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' Form_KeyDown

Private Sub Form_Unload(Cancel%)
 geladen = False
End Sub ' Form_Unload

Private Sub HAAusw_GotFocus()
 If Not haGewählt Then Call AuswHA(Me)
End Sub ' HAAusw_GotFocus()

Private Sub inTM_Click()
 If IsNumeric(Me.Pat_id) Then inTMAnz Me.Pat_id
End Sub 'inTM_Click()

Private Sub Patientenlaufzettel_Click()
 Dim zzn%
 zzn = 8
 If IsNumeric(Me.Zeilenzahl) Then zzn = CInt(zzn)
 If IsNumeric(Me.Pat_id) Then
  Call dodoPLZ(Me.Pat_id, plzVz, Now, Now - Int(Now), True, "", zzn, obRueck)
 End If
End Sub ' Patientenlaufzettel_Click()

Private Sub Abr_Click()
 If IsNumeric(Me.Pat_id) Then Call tubriefStandalone(Me.PatID, False, , Me.Vorlage)
End Sub ' Abr_Click()

Private Sub PatName_DblClick()
' Stop
End Sub ' PatName_DblClick

Private Sub PatName_DragDrop(source As Control, x As Single, Y As Single)
' Stop
End Sub ' PatName_DragDrop

Private Sub PatName_DropDown()
 Call patname_gotfocus
End Sub ' PatName_DropDown

Private Sub pat_id_dropdown()
 If Not Pat_IDGewählt Then
  Call Me.AuswPat_id(Me)
  Pat_IDGewählt = True
 End If
End Sub ' pat_id_dropdown()

Private Sub pat_id_gotfocus()
 PatZuWählen = True
End Sub ' pat_id_gotfocus()

Private Sub patname_gotfocus()
 If PatZuWählen Then
  If Not PatGewählt Then
   Call AuswName(Me)
   PatGewählt = True
  End If
 End If
 Me.PatName.SelStart = Len(Me.PatName)
 Me.PatName.SelLength = 0
End Sub ' patname_gotfocus()

Private Sub PatName_KeyPress(KeyAscii As Integer)
' Stop
End Sub

' Knopf Thererartenanz
Private Sub Therapiearten_Click()
 Dim rs As New ADODB.Recordset, spmaxü
 spmaxü = Array(10, 5, 200)
 If Me.Pat_id <> 0 Then
  myFrag rs, "SELECT * FROM therarten WHERE pat_id=" & CStr(Me.Pat_id) & " ORDER BY zp DESC, mpnr DESC"
  TabAusgeb rs, Me, , , , , spmaxü, , "Therapiearten von Pat. " & CStr(Me.Pat_id) '& " (" & gesname(PID) & ")"
 End If
End Sub ' Therapiearten_Click()

Private Sub VorDat_Click() 'Vordaten der Briefe raussuchen
 do_Pat_ID_Change (True)
End Sub

Private Sub Vorlage_GotFocus()
  Select Case hlese.Aktion
   Case Briefschreiben
    If Not VorlageGewählt Then Call AuswVorlage(Me)
   Case Else
    Me.VorlageLab.Visible = False
    Me.Vorlage.Visible = True
  End Select
End Sub ' Vorlage_GotFocus

Private Sub HAAusw_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
 If KeyCode = 13 Then
  Dim rNaA As New ADODB.Recordset
  Dim sql$, KVNr$
  sql = "SELECT kvnr FROM hareal WHERE adressat LIKE '%" & Me.HAAusw & "%'"
  myFrag rNaA, sql
  If Not rNaA.BOF Then
   KVNr = rNaA!KVNr
   If KVNr <> vNS Then
    Set rNaA = Nothing
    sql = "SELECT DISTINCT * from(" & _
       "SELECT `namen`.pat_id, quartal, bhfb, `namen`.nachname, `namen`.vorname, `namen`.gebdat  FROM `namen` LEFT JOIN (SELECT pat_id, MIN(quartal) AS quartal, MIN(bhfb) AS bhfb FROM (SELECT pat_id, quartal, bhfb FROM `faelle` ORDER BY bhfb DESC) AS innen GROUP BY pat_id) AS minnen ON minnen.pat_id = `namen`.pat_id WHERE kvnr = '" & KVNr & "' UNION " & _
       "SELECT `namen`.pat_id, quartal, bhfb, `namen`.nachname, `namen`.vorname, `namen`.gebdat  FROM (SELECT pat_id, MIN(quartal) AS quartal, MIN(bhfb) AS bhfb, MIN(übwr) AS übwr FROM (SELECT pat_id, quartal, bhfb, übwr FROM `faelle` ORDER BY bhfb DESC) AS innen GROUP BY pat_id) AS minnen LEFT JOIN `namen` ON minnen.pat_id = `namen`.pat_id WHERE übwr = '" & KVNr & "') AS in1 ORDER BY pat_id DESC"

    If LenB(DBCnS) = 0 Then 'DBCn.ConnectionString = "" THEN
'   Call Me.hlese.ConstrFestleg(0, Me.hlese)
     Call acon(quelleT)
    End If
    Me.FaellepHA.Clear
    myFrag rNaA, sql
    Me.Zahl = 0
    Do While Not rNaA.EOF
     Me.Zahl = Me.Zahl + 1
     Me.FaellepHA.AddItem rNaA!Pat_id & " " & rNaA!Quartal & " " & rNaA!Nachname & " " & rNaA!Vorname & " " & rNaA!GebDat
     rNaA.Move 1
    Loop
    Call Me.FaellepHA_Click
    Exit Sub
   End If
  End If
 End If
End Sub ' HAAusw_KeyDown

Private Sub Hausarzt_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Hausarzt_KeyDown

Private Sub Vorlage_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Vorlage_KeyDown(KeyCode As Integer, Shift As Integer)
Private Sub OKButton_Click()
 Me.Visible = False
 cR.WriteKey Me.Vorlage, "Vorlage", RegPos, HKEY_CURRENT_USER, REG_SZ
 Me.hlese.los
End Sub ' OKButton_Click

Private Sub OKButton_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' OKButton_KeyDown

Private Sub Pat_ID_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Pat_ID_KeyDown
        
Private Sub Form_Load()
 geladen = True
 On Error GoTo fehler
  Me.Zeilenzahl.Visible = False
  Me.ZeilenzahlL.Visible = False
  If Me.Visible Then Screen.MousePointer = vbHourglass
  DoEvents
  If hlese Is Nothing Then Set hlese = Lese
  Me.Caption = "Patientenauswahl " & IIf(lies.obMySQL, "MySQL: " & Lese.MyDB, Lese.dlg.MdB)
'  Call AuswName(Me)
'  Call AuswHA(Me)
#If EinmalDB Then
#Else
  Call AuswPat_id(Me)
#End If
  Select Case hlese.Aktion
   Case Briefschreiben
    Me.Angeforderte(1).Visible = True
    Me.Angeforderte(0).Visible = True
    Me.VorlageLab.Visible = True
    Me.Vorlage.Visible = True
    RegPos = RegWurzel & App.EXEName & "\PatAuswahl"
    Me.Vorlage = cR.ReadKey("Vorlage", RegPos, HKEY_CURRENT_USER)
   Case Else
    Me.Angeforderte(1).Visible = False
    Me.Angeforderte(0).Visible = False
    Me.VorlageLab.Visible = False
    Me.Vorlage.Visible = False
  End Select
  
  Me.Faelle.Clear
  Me.Diagnosen.Clear
  Me.Arzt = vNS
  If Not lies.obMySQL Then
   Call KVAccSuch
  End If
'  Me.hlese.MousePointer = vbDefault
 Screen.MousePointer = vbDefault
 DoEvents
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in form_load/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Form_Load

' kommt vor IN doPatNameChange, SchreibDatensatz und FDC_indvorWechsel
Function getPat_id&(PatName$)
 Dim Spl$(), sql$
 Dim rNaA As New ADODB.Recordset
 On Error GoTo fehler
 If IsNull(PatName) Or PatName = vNS Then
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
 If getPat_id = -1 And Me.PatName <> vNS Then
  If IsNumeric(Me.PatName) Then
   sql = "SELECT nachname, vorname, gebdat, pat_id FROM `namen` WHERE pat_id LIKE '" & UmwfSQL(Me.PatName) & "%'"
  Else
   sql = "SELECT nachname, vorname, gebdat, pat_id FROM `namen` WHERE CONCAT(nachname¡ ', '¡ vorname) LIKE '" & UmwfSQL(Me.PatName) & "%'"
  End If
  sql = SqlU(sql, hlese.obMySQL) ' 11.7.09
'  IF hlese.obMySQL THEN
'   SQL = replace$(SQL, "¡", ",")
'  Else
'   SQL = replace$(replace$(SQL, "concat", ""), "¡", " & ")
'  END IF
  myFrag rNaA, sql
  If Not rNaA.EOF Then
   getPat_id = rNaA!Pat_id
   PatID = getPat_id
  End If
 End If
 Exit Function
fehler:
ErrDescription = Err.Description
ErrNumber = Err.Number
ErrLastDllError = Err.LastDllError
ErrSource = Err.source
' Der ODBC-Treiber unterstützt die angeforderten Eigenschaften nicht.
' -2147217887
If InStrB(ErrDescription, "Der ODBC-Treiber unterstützt") <> 0 Then
 Call DBCnOpen
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in getPat_id/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getPat_id

Function AuswName(frm As PatAuswahl)
 Dim zwi$, i&, t$
 syscmd 4, "Wähle Patientien aus"
 DoEvents
altobTrans:
 #If EinmalDB Then
 Dim sEl As SortierPat_IDText
' Dim sLi As New SortierListe
 #End If
 On Error GoTo fehler
 Dim rNaA As New ADODB.Recordset
' myFrag rNaA, "SELECT nachname, vorname, gebdat, pat_id FROM `namen` ORDER BY nachname, vorname, gebdat;"
 #If EinmalDB Then
 sql = "SELECT pat_id, CAST(CONCAT(CAST(nachname AS CHAR)¡ "", ""¡ vorname¡ "",*""¡ DATE_FORMAT(gebdat,'%e.%c.%Y')¡ "" | ""¡ pat_id) AS char) f0, CONCAT(pat_id¡"" | ""¡ CAST(nachname AS char)¡ "", ""¡ vorname¡ "",*""¡ DATE_FORMAT(gebdat,'%e.%c.%Y')) AS f1 FROM `namen` ORDER BY nachname, vorname, gebdat;"
 #Else
 sql = "SELECT CAST(CONCAT(CAST(nachname AS char)¡ "", ""¡ vorname¡ "",*""¡ DATE_FORMAT(gebdat,'%e.%c.%Y')¡ "" | ""¡ pat_id) AS char) f0 FROM `namen` ORDER BY nachname, vorname, gebdat;"
 #End If
 If Forms(0).obMySQL Then
  sql = REPLACE$(sql, "¡", ",")
 Else
  sql = REPLACE$(REPLACE$(REPLACE$(REPLACE$(sql, "¡", " & "), "CONCAT", ""), "DATE_FORMAT", "FORMAT"), "'%e.%c.%Y'", """SHORT DATE""")
  sql = REPLACE$(REPLACE$(sql, "CAST(", ""), " AS CHAR)", "")
 End If
 myFrag rNaA, sql
 zwi = frm.PatName
 frm.PatName.Clear
 frm.PatName = zwi
 frm.PatName.SelStart = Len(Me.PatName.Text)
 Do While Not rNaA.EOF
  If Not IsNull(rNaA!F0) Then
  t = rNaA!F0 ' rNaA!NachName + ", " & rNaA!VorName + ",*" + format$(rNaA!GebDat, "D.M.YY") & " | " + CStr(rNaA!Pat_id)
  frm!PatName.AddItem t
  i = i + 1
  If i = 30 Then
   syscmd 4, "Pat: " & rNaA!F0 'rNaA!NachName & " " & rNaA!VorName
   DoEvents
   i = 0
  End If
  #If EinmalDB Then
  Set sEl = New SortierPat_IDText
  sEl.Pat_id = -rNaA!Pat_id
  sEl.Text = rNaA!f1 'CStr(rNaA!Pat_id) & " | " & rNaA!NachName & ", " & rNaA!VorName + ",*" + format$(rNaA!GebDat, "D.M.YY")
'  Call sLi.sCAdd(sEl)
  #End If
  End If
  rNaA.Move 1
 Loop

' #If EinmalDB THEN
' frm.Pat_id.Clear
' For i = 1 To sLi.Count
'  frm.Pat_id.AddItem sLi.Item(i).Text
' Next i
' #END IF
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147467259 Or Err.Number = -2147217887 Then ' Server has gone away / ungültiger Eigenschaftswert
 Call DBCnOpen
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AuswName/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
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
' sql = "SELECT CONCAT(pat_id¡"" : ""¡ nachname¡ "", ""¡ vorname¡ "",*""¡ DATE_FORMAT(gebdat,'%e.%c.%Y')) AS f1 FROM `namen` ORDER BY pat_id;"
' IF Forms(0).obMySQL THEN
'  sql = replace$(sql, "¡", ",")
' Else
'  sql = replace$(replace$(replace$(replace$(sql, "¡", " & "), "concat", ""), "DATE_FORMAT", "format"), "'%e.%c.%Y'", """short date""")
' END IF
 sql = "SELECT CONCAT(CAST(pat_id AS char),' | ', nachname, ', ', vorname, ',*', DATE_FORMAT(gebdat,'%e.%c.%Y')) AS f1 FROM `namen` ORDER BY pat_id DESC"
 myFrag rNaA, sql
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Auswpat_id/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AuswPat_id

Function repl(q$, r$, z$)
 repl = REPLACE$(q, r, z)
End Function ' repl(q$, r$, z$)

' zuvor auch IN form_load, jetzt nur noch IN HAAusw_GotFocus
Function AuswHA(frm As PatAuswahl)
 Dim zwi$, sql$, i%, hier%, ItemStr$
 Dim kvPfad$, p1&, p2&
 On Error GoTo fehler
 syscmd 4, "Hausarztauswahl ..."
 DoEvents
 'Dim cnADO As New ADODB.Connection
 'cnADO.Open ConStr
 Dim rHS As New ADODB.Recordset, rKVNr As New ADODB.Recordset
 myFrag rKVNr, "SELECT DISTINCT kvnr FROM `namen` UNION SELECT DISTINCT übwr FROM `faelle`" ' DISTINCT könnte hier entfallen, ist durch UNION impliziert
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
'   sql = "SELECT dbnr, bstelle, anrede, """" AS überschrift, haname, ort, kvnu AS kvnr, tel1, tel2, tel3, tel4, fax1, fax1k,fax2,fax2k, fax3,fax3k, email, zulg, arzttyp, gemmit, beme, -dmpt2 AS j_dmpt2, -dmpt1 AS j_dmpt1, geschlecht, titel, vorname, nachname, straße, plz FROM `kvaerzte`.hae WHERE kvnr = '" & LEFT(rKVNr!KVNr, 2) & "/" & mid$(rKVNr!KVNr, 3, 5) & "' " & _
         "UNION SELECT dbnr, bstelle, anrede, """" AS überschrift, haname, ort, kvnu AS kvnr, tel1, tel2, tel3, tel4, fax1, fax1k,fax2,fax2k, fax3,fax3k, email, zulg, arzttyp, gemmit, beme, -dmpt2 AS j_dmpt2, -dmpt1 AS j_dmpt1, geschlecht, titel, vorname, nachname, straße, plz FROM `kvaerzte`.`haealt` WHERE kvnr = '" & LEFT(rKVNr!KVNr, 2) & "/" & mid$(rKVNr!KVNr, 3, 5) & "' " & _
         "UNION SELECT id AS dbnr, """" AS bstelle, """" AS anrede, überschrift, name AS haname, ort, kvnr, telefon AS tel1, """" AS tel2, """" AS tel3, """" AS tel4, telefax AS fax1, replace$(replace$(replace$(telefax,""/"",""""),""-"",""""),"" "","""") AS fax1k,"""" AS fax2,"""" AS fax2k, """" AS fax3, """" AS fax3k, e_mail AS email, zulassungsgebiet AS zulg, arzttyp, `gemeinschaftspraxis mit` AS gemmit, beme, dmpt2, dmpt2, geschlecht, titel, vorname, nachname, straße, plz FROM `hausaerzte` WHERE kvnr = '" & rKVNr!KVNr & "' " & _
         "UNION SELECT """" AS dbnr, """" AS bstelle, """" AS anrede, """" AS überschrift, name AS haname, ort, kvnr, telefon AS tel1, """" AS tel2, """" AS tel3, """" AS tel4, """" AS fax1, fax AS fax1k,"""" AS fax2,"""" AS fax2k, """" AS fax3, """" AS fax3k, """" AS email, fachgruppe AS zulg, """" AS arzttyp, """" AS gemmit, """" AS beme, -2 AS dmpt2, -2 AS dmpt1, 0 AS geschlecht, titel, vorname, name AS nachname, strasse, plz FROM `liuez` WHERE kvnr = '" & rKVNr!KVNr & "'"
'    sql = "SELECT dbnr, bstelle, anrede, """" AS überschrift, haname, ort, kvnu AS kvnr, tel1, tel2, tel3, tel4, fax1, fax1k,fax2,fax2k, fax3,fax3k, email, zulg, arzttyp, gemmit, beme, dmpt2, dmpt1, geschlecht, titel, vorname, nachname, straße, plz FROM `kvaerzte`.`hae` WHERE kvnr = '" & LEFT(rKVNr!KVNr, 2) & "/" & Mid$(rKVNr!KVNr, 3, 5) & "' " & _
         "UNION SELECT dbnr, bstelle, anrede, """" AS überschrift, haname, ort, kvnu AS kvnr, tel1, tel2, tel3, tel4, fax1, fax1k,fax2,fax2k, fax3,fax3k, email, zulg, arzttyp, gemmit, beme, dmpt2, dmpt1, geschlecht, titel, vorname, nachname, straße, plz FROM `kvaerzte`.`hae`alt WHERE kvnr = '" & LEFT(rKVNr!KVNr, 2) & "/" & Mid$(rKVNr!KVNr, 3, 5) & "' " & _
         "UNION SELECT id AS dbnr, """" AS bstelle, """" AS anrede, überschrift, name AS haname, ort, kvnr, telefon AS tel1, """" AS tel2, """" AS tel3, """" AS tel4, telefax AS fax1, REPLACE(REPLACE(REPLACE(telefax,""/"",""""),""-"",""""),"" "","""") AS fax1k,"""" AS fax2,"""" AS fax2k, """" AS fax3, """" AS fax3k, e_mail AS email, zulassungsgebiet AS zulg, arzttyp, `gemeinschaftspraxis mit` AS gemmit, beme, dmpt2, dmpt2, geschlecht, titel, vorname, nachname, straße, plz FROM `hausaerzte` WHERE kvnr = '" & rKVNr!KVNr & "' " & _
         "UNION SELECT """" AS dbnr, """" AS bstelle, """" AS anrede, """" AS überschrift, name AS haname, ort, kvnr, telefon AS tel1, """" AS tel2, """" AS tel3, """" AS tel4, """" AS fax1, fax AS fax1k,"""" AS fax2,"""" AS fax2k, """" AS fax3, """" AS fax3k, """" AS email, fachgruppe AS zulg, """" AS arzttyp, """" AS gemmit, """" AS beme, -2 AS dmpt2, -2 AS dmpt1, 0 AS geschlecht, titel, vorname, name AS nachname, strasse, plz FROM `liuez` WHERE kvnr = '" & rKVNr!KVNr & "'"
     sql = "SELECT * FROM (" & _
          "SELECT GROUP_CONCAT(DISTINCT nachname) haname, ort, CAST(LEFT(bsnr,7) AS char) kvnr, REPLACE(tel,'-','') tel1, REPLACE(fax,'-','') fax1, titel, vorname, nachname, CONCAT(straße,' ',hausnr) straße, plz FROM " & HADBName & ".bs LEFT JOIN " & HADBName & ".ort ON ort_id = idort LEFT JOIN " & HADBName & ".tel ON tel.bs_id = idbs LEFT JOIN " & HADBName & ".fax ON fax.bs_id = idbs LEFT JOIN " & HADBName & ".arzt_has_bs ahb ON ahb.bs_id = idbs LEFT JOIN " & HADBName & ".arzt a ON a.idarzt = ahb.arzt_id LEFT JOIN " & HADBName & ".titel t ON t.idtitel = a.titel_id WHERE bsnr = '" & IIf(IsNull(rKVNr!KVNr), "000000", rKVNr!KVNr) & "00' " & _
          "UNION SELECT haname, ort, CAST(kvnu AS char) kvnr, tel1, fax1, titel, vorname, nachname, straße, plz FROM `kvaerzte`.`hae` WHERE kvnr = '" & Left(IIf(IsNull(rKVNr!KVNr), "000000", rKVNr!KVNr), 2) & "/" & Mid$(IIf(IsNull(rKVNr!KVNr), "000000", rKVNr!KVNr), 3, 5) & "' " & _
          "UNION SELECT haname, ort, CAST(kvnu AS char) kvnr, tel1, fax1, titel, vorname, nachname, straße, plz FROM `kvaerzte`.`haealt` WHERE kvnr = '" & Left(IIf(IsNull(rKVNr!KVNr), "000000", rKVNr!KVNr), 2) & "/" & Mid$(IIf(IsNull(rKVNr!KVNr), "000000", rKVNr!KVNr), 3, 5) & "' " & _
          "UNION SELECT name haname, ort, CAST(kvnr AS char) kvnr, telefon AS tel1, telefax AS fax1, titel, vorname, nachname, straße, plz FROM `hausaerzte` WHERE kvnr = '" & IIf(IsNull(rKVNr!KVNr), "000000", rKVNr!KVNr) & "' " & _
          "UNION SELECT name haname, ort, CAST(kvnr AS char) kvnr, telefon AS tel1, '' AS fax1, titel, vorname, name AS nachname, strasse, plz FROM `aktlue` WHERE kvnro = '" & IIf(IsNull(rKVNr!KVNr), "000000", rKVNr!KVNr) & "' " & _
          ") i WHERE NOT ISNULL(nachname) ORDER BY haname DESC;"
' geht nicht wegen der Reihenfolge der Felder
'    sql = "SELECT * FROM (" & _
'          kvneuroh & "WHERE bsnr = '" & rKVNr!KVNr & "00' " & _
'          "UNION SELECT haname, ort, CAST(kvnu AS char) kvnr, tel1, fax1, titel, vorname, nachname, straße, plz FROM `kvaerzte`.`hae` WHERE kvnr = '" & LEFT(rKVNr!KVNr, 2) & "/" & Mid$(rKVNr!KVNr, 3, 5) & "' " & _
'          "UNION SELECT haname, ort, CAST(kvnu AS char) kvnr, tel1, fax1, titel, vorname, nachname, straße, plz FROM `kvaerzte`.`haealt` WHERE kvnr = '" & LEFT(rKVNr!KVNr, 2) & "/" & Mid$(rKVNr!KVNr, 3, 5) & "' " & _
'          "UNION SELECT name haname, ort, CAST(kvnr AS char) kvnr, telefon AS tel1, telefax AS fax1, titel, vorname, nachname, straße, plz FROM `hausaerzte` WHERE kvnr = '" & rKVNr!KVNr & "' " & _
'          "UNION SELECT name haname, ort, CAST(kvnr AS char) kvnr, telefon AS tel1, '' AS fax1, titel, vorname, name AS nachname, strasse, plz FROM `liuez` WHERE kvnr = '" & rKVNr!KVNr & "' " & _
'          ") i WHERE NOT ISNULL(nachname) ORDER BY haname DESC;"
   If Not Forms(0).obMySQL Then
    sql = REPLACE$(sql, "`kvaerzte`.", "`" & kvPfad & "`.")
    sql = REPLACE$(sql, "replace$(replace$(replace$(telefax,""/"",""""),""-"",""""),"" "","""")", "telefax")
   End If
   Set rHS = Nothing
   myFrag rHS, sql
   If Not rHS.BOF Then
' Nachname, Titel, Vorname, KVNr, Tel1, Fax1, Straße, PLZ, Ort
    ItemStr = rHS!Nachname & ", " & rHS!Titel & ", " & rHS!Vorname & ", KVNr:" & rHS!KVNr & ", T." & rHS!tel1 & " F." & rHS!fax1 & ", " & rHS!Straße & ", " & rHS!plz & " " & rHS!ort
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AuswHA/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AuswHA

Function AuswVorlage(frm As PatAuswahl)
 Dim zwi$, sql$, i%, hier%, ItemStr$, Verz$
 Dim kvPfad$, p1&, p2&
 On Error GoTo fehler
 syscmd 4, "Vorlagen-Auswahl ..."
 DoEvents
 Verz = getDokPfad("Vorlagen")
 Me.Vorlage.Clear
 ItemStr = Dir(Verz & "\*brief*.dot")
 Do While ItemStr <> vNS
  For i = 0 To frm.Vorlage.ListCount
   If frm.Vorlage.List(i) > ItemStr Then
    Call frm.Vorlage.AddItem(ItemStr, i)
    GoTo hamma
   End If
  Next i
  Call frm.Vorlage.AddItem(ItemStr)
hamma:
  syscmd 4, "Vorlage: " & ItemStr
  ItemStr = Dir
  DoEvents
 Loop
' Me.Vorlage = "Accessbrief.dot"
 VorlageGewählt = True
 syscmd 5
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AuswVorlage/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AuswVorlage

Private Sub HAAusw_Click()
 Call HAAusw_Change
End Sub ' HAAusw_Click()

Private Sub HAAusw_Change()
 Dim rNaA As New ADODB.Recordset
 Dim sql$, KVNr$
 KVNr = Mid$(Me.HAAusw, InStr(Me.HAAusw, "KVNr") + 5, 7)
 If KVNr <> vNS Then
  sql = "SELECT DISTINCT * from(" & _
       "SELECT `namen`.pat_id, quartal, bhfb, `namen`.nachname, `namen`.vorname, `namen`.gebdat  FROM `namen` LEFT JOIN (SELECT pat_id, MIN(quartal) AS quartal, MIN(bhfb) AS bhfb FROM (SELECT pat_id, quartal, bhfb FROM `faelle` ORDER BY bhfb DESC) AS innen GROUP BY pat_id) AS minnen ON minnen.pat_id = `namen`.pat_id WHERE kvnr = '" & KVNr & "' UNION " & _
       "SELECT `namen`.pat_id, quartal, bhfb, `namen`.nachname, `namen`.vorname, `namen`.gebdat  FROM (SELECT pat_id, MIN(quartal) AS quartal, MIN(bhfb) AS bhfb, MIN(übwr) AS übwr FROM (SELECT pat_id, quartal, bhfb, übwr FROM `faelle` ORDER BY bhfb DESC) AS innen GROUP BY pat_id) AS minnen LEFT JOIN `namen` ON minnen.pat_id = `namen`.pat_id WHERE übwr = '" & KVNr & "') AS in1 ORDER BY pat_id DESC"

  If LenB(DBCnS) = 0 Then 'DBCn.ConnectionString = "" THEN
'   Call Me.hlese.ConstrFestleg(0, Me.hlese)
   Call acon(quelleT)
  End If
 Me.FaellepHA.Clear
 myFrag rNaA, sql
 Me.Zahl = 0
 Do While Not rNaA.EOF
  Me.Zahl = Me.Zahl + 1
  Me.FaellepHA.AddItem rNaA!Pat_id & " " & rNaA!Quartal & " " & rNaA!Nachname & " " & rNaA!Vorname & " " & rNaA!GebDat
  rNaA.Move 1
 Loop
 Call Me.FaellepHA_Click
  Exit Sub
End If ' kvnr <> vns
Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in HAAusw_Change/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' haausw_change

Private Sub Pat_ID_Change()
 Call do_Pat_ID_Change
End Sub ' Pat_ID_Change()

Private Sub do_Pat_ID_Change(Optional mitVorDat%)
 Dim rNaA As New ADODB.Recordset
 Dim rFaA As New ADODB.Recordset
 Dim rDi As New ADODB.Recordset
 Dim rEin As New ADODB.Recordset
 Dim HAStr$, obHA%
 Static rKVA As New ADODB.Recordset
 Static innen%
 Dim pos&
 If innen Then Exit Sub
' Call KVÄVorb
 Me.MousePointer = vbHourglass
 pos = InStr(Me.Pat_id, " |")
 If pos <> 0 Then innen = True: Me.Pat_id = Left(Me.Pat_id, pos - 1): innen = False
 If Me.Pat_id <> vNS And Me.Pat_id <> "-1" Then
  On Error Resume Next
  If Not IsNumeric(Me.Pat_id) Then
   innen = True
   Me.PatName = Me.Pat_id
   Me.Pat_id = vNS
   Me.PatName.SetFocus
   innen = False
   Exit Sub
  End If
  Me.PatID = getPid(Me.Pat_id)
  On Error GoTo fehler
  If Me.PatID = 0 Then Exit Sub
  If Me.Pat_id <> Me.VorBriefID Or InStrB(Me.PatName, "|") <> 0 Then ' aufwändiges Raussuchen des Vorbriefdatums nur+immer bei Auswahl aus der Patientenliste mit der Maus
'   Call acon(HaT)
   If LenB(DBCn) = 0 Or DBCn = "" Then Call acon(quelleT)
   Me.Faelle.Clear
   Me.Diagnosen.Clear
   Me.Vorbriefe.Clear
   If LenB(Me.HAAusw) = 0 Then Me.Vorbriefe.Clear
   Me.Hausarzt = vNS
   Me.Arzt = vNS
   If LenB(DBCnS) = 0 Then ' DBCn.ConnectionString = "" THEN
'   Call Me.hlese.ConstrFestleg(0, Me.hlese)
    Call acon(quelleT)
   End If
   myFrag rNaA, "SELECT nachname, vorname, gebdat, pat_id FROM `namen` WHERE pat_id = " & Me.PatID & " ORDER BY nachname, vorname, gebdat"
   If rNaA.EOF Then
    Me.pat_idDaten = vNS
   Else
    Me.pat_idDaten = rNaA!Nachname & ", " & rNaA!Vorname + ",*" + Format$(rNaA!GebDat, "D.M.YY") + " | " + CStr(rNaA!Pat_id)
' Call KVÄVorb
'    Call acon(HaT)
'    rFaA.Open "SELECT Übwr, BhfB, bhFE1, SchGr, Name, Kateg, ÜwnNr, ÜwNaN, ÜwVor FROM `faelle` LEFT JOIN (SELECT vknr, ik, kateg, name FROM `kassenliste` GROUP by ik,vknr) AS kl ON `faelle`.vknr = kl.vknr AND faelle.ik=kl.ik WHERE pat_id = " & Me.PatID & " ORDER BY bhfb DESC", DBCn, adOpenDynamic, adLockReadOnly
    myFrag rFaA, "SELECT Übwr, BhfB, bhFE1, SchGr, Name, Kateg, ÜwnNr, ÜwNaN, ÜwVor FROM `faelle` LEFT JOIN (SELECT vknr, ik, kateg, name FROM `kassenliste` GROUP by ik,vknr) AS kl ON `faelle`.vknr = kl.vknr AND faelle.ik=kl.ik WHERE pat_id = " & Me.PatID & " ORDER BY bhfb DESC"
    If Not rFaA.BOF Then
     Do While Not rFaA.EOF
      HAStr = vNS
      If Not rFaA!Übwr = vNS Then
       Set rKVA = Nothing
 '    rKVA.Open "SELECT -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2, `hae`.* FROM `hae` WHERE kvnr = '" & LEFT(rFaA!Übwr, 2) & "/" & right$(rFaA!Übw, 5) & "'", HAECn, adOpenDynamic, adLockReadOnly
       Dim infos$()
       Call getHausarzt(Me.Pat_id, infos)
       Me.Hausarzt = infos(10, 0) & ":  " & infos(1, 0) & " " & infos(3, 0) & " " & infos(4, 0) & "  " & IIf(infos(6, 0) = "X", "DMP2", vNS) & " " & IIf(infos(7, 0) = "X", "DMP1", vNS)
 '      rKVA.Open "SELECT HAName, Ort, Email, Tel1, Fax1k, DMPT1, DMPT2 FROM `hae` WHERE kvnr = '" & LEFT(rFaA!Übw, 2) & "/" & Right$(rFaA!Übw, 5) & "'", HAECn, adOpenDynamic, adLockReadOnly
 '      IF Not rKVA.BOF AND NOT obHA AND NOT rFaA!Übw = "6419153" THEN
'        HAStr = rKVA!haname & ", " & rKVA!Ort & IIf(NOT ISNULL(rKVA!Email) AND LenB(rKVA!Email) <> 0, " Email: " & rKVA!Email, "") & _
 '               "   Tel: " & replace$(IIf(ISNULL(rKVA!tel1), vNS, rKVA!tel1), " ", "") & "  Fax: " & IIf(ISNULL(rKVA!fax1k), vNS, rKVA!fax1k)
 '       Dim obDMP2%, obDMP1%
 '       Do While Not rKVA.EOF
 ' '      IF rKVA!j_dmpt2 <> 0 THEN obDMP2 = True
 '        IF rKVA!dmpt2 <> 0 THEN obDMP2 = True
 ' '      IF rKVA!j_dmpt1 <> 0 THEN obDMP1 = True
 '        IF rKVA!dmpt1 <> 0 THEN obDMP1 = True
 '        rKVA.Move 1
 '       Loop
 '       HAStr = HAStr & "  DMP 2: " & IIf(obDMP2, "+", "-") & "  1: " & IIf(obDMP1, "+", "-")
 '       Me.Hausarzt = HAStr
        obHA = -1
 '      END IF
      End If
      Me.Faelle.AddItem (rFaA!BhFB & " - " & rFaA!BhFE1 & " " & rFaA!SchGr & "  " & Left(IIf(IsNull(rFaA!name), vNS, rFaA!name), 20) & " " & IIf(IsNull(rFaA!Kateg), vNS, rFaA!Kateg) & "  " & rFaA!ÜWNNr & " " & rFaA!ÜWNaN & " " & rFaA!ÜWVor)
      rFaA.Move 1
     Loop
    End If ' Not rFaA.BOF THEN
    myFrag rDi, "SELECT diagsicherheit, icd, diagtext, DATE_FORMAT(diagdatum,'%d.%m.%y') diagdatum, obdauer " & vbCrLf & _
                "FROM diagview d " & vbCrLf & _
                "WHERE pat_id = " & Me.PatID & " GROUP BY diagsicherheit, icd, diagtext, obdauer ORDER BY icd" ' AND COALESCE(d.f6010,0)=0
    If Not rDi.BOF Then
     Do While Not rDi.EOF
      Me.Diagnosen.AddItem rDi!DiagSicherheit & " " & rDi!DiagDatum & " " & IIf(rDi!obDauer = 0, "f", "d") & " " & Left$(rDi!ICD & Space$(7), 7) & rDi!DiagText
      rDi.MoveNext
     Loop
    End If
    Zulp = 0
    lbeh = 0
    Me.Arzt.BackColor = &H80FF&
    Me.nötig.BackColor = &HFF& ' rot
    myFrag rEin, "SELECT COUNT(0) tk, DATE_FORMAT(MAX(zeitpunkt),'%e.%c.%y') zp FROM `eintraege` WHERE (art IN ('tk','ARCHIE2','APK') OR inhalt LIKE '%(tk)%') AND pat_id = " & Me.Pat_id
    If Not rEin.BOF Then
     If rEin!tk <> 0 Then
      Zulp = MAX(Zulp, rEin!Zp)
      Me.Arzt = "Kothny (" & rEin!tk & ", zul.: " & rEin!Zp & ") "
      lbeh = rEin!tk
      Me.Arzt.BackColor = &HFF&
     End If
    End If
    Set rEin = Nothing
    myFrag rEin, "SELECT COUNT(0) gs, DATE_FORMAT(MAX(zeitpunkt),'%e.%c.%y') zp FROM `eintraege` WHERE ((art IN ('gs','doppler','duplex') AND NOT inhalt LIKE '%(tk)%') OR inhalt LIKE '%(gs)%') AND pat_id = " & Me.Pat_id
    If Not rEin.BOF Then
     If rEin!gs <> 0 Then
       If rEin!Zp > Zulp And rEin!gs > lbeh Then
        Me.Arzt.BackColor = &HFFFF&
       ElseIf rEin!Zp > Zulp Or rEin!gs > lbeh Then
        Me.Arzt.BackColor = &H80FF&
       Else
        Me.Arzt.BackColor = &HFF&
       End If
      If LenB(Me.Arzt) <> 0 Then Me.Arzt = Me.Arzt & ", "
      Zulp = MAX(Zulp, rEin!Zp)
      Me.Arzt = Me.Arzt & " Schade (" & rEin!gs & ", zul.: " & rEin!Zp & ")"
     End If
    End If
   
'   IF LenB(Me.HAAusw) = 0 THEN
     Dim rs As New ADODB.Recordset
     myFrag rs, "SELECT transe, docname , RCFax, pages, fsize, Retries FROM `faxeinp`.`outa` o WHERE docname RLIKE 'PID " & Pat_id & "[^0123456789]' ORDER BY transe DESC"
     If Not rs.BOF Then
      Do While Not rs.EOF
       If rs!transe >= Zulp Then Me.nötig.BackColor = &HFF0000 ' blau
       Me.Vorbriefe.AddItem "gefaxt: " & rs!transe & "  |  " & rs!docName & "  |  " & rs!rcfax & "  |  " & rs!Pages & "  |  " & rs!fsize & "  |  " & rs!Retries
       rs.MoveNext
      Loop
     End If ' Not rs.BOF THEN
     Dim vd As Date
     Dim zeitp1 As Date, name$
'     IF hlese.Aktion = BriefSchreiben AND InStrB(Me.PatName, "|") <> 0 THEN ' aufwändiges Raussuchen des Vorbriefdatums nur+immer bei Auswahl aus der Patientenliste mit der Maus
      If mitVorDat Then
      vd = GetVorDat(Pat_id, -1, -1, 0, zeitp1, name)
      If vd <> 0 Then
       If vd >= Zulp Then Me.nötig.BackColor = &HFF0000 ' blau
       Me.Vorbriefe.AddItem "Kart'k'eintr v.: " & zeitp1 & ": " & name & ", geschrieben: " & vd
      End If
     Else
      vd = GetVorDat(Pat_id, -1, -1, -1, zeitp1, name)
     End If ' hlese.Aktion = BriefSchreiben THEN
  '   END IF ' LenB(Me.HAAusw) = 0 THEN
   End If ' rNaA.EOF
   Dim sql$, rdesk As New ADODB.Recordset
   sql = "SELECT (0) FROM desktop WHERE pat_id = " & Me.Pat_id & " AND titel LIKE '%kein%Bericht%'"
   Set rdesk = Nothing
   myFrag rdesk, sql
   If Not rdesk.BOF Then
    Me.keinBericht = "kein Bericht!"
   Else
    Me.keinBericht = ""
   End If
   Me.VorBriefID = Me.Pat_id
  End If ' me.vorbriefid <> me.pat_id
 End If ' Me.Pat_id <> vNS THEN
 Me.MousePointer = vbDefault
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Pat_ID_Change()/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' do_Pat_ID_Change

Public Function getPid(Text$)
 Dim pos&
 pos = InStr(Text, " ")
 If pos > 0 Then
  getPid = Left(Me.Pat_id, pos - 1)
 Else
  getPid = Me.Pat_id
 End If
End Function ' getPid(Text$)

Private Sub Pat_id_Click()
 Call Pat_ID_Change
End Sub ' Pat_id_Click()

Private Sub PatName_Click()
' erst wenn Leerzeichen oder Komma enthalten, Infos raussuchen
 If InStrB(Me.PatName, ",") <> 0 Or InStrB(Me.PatName, " ") <> 0 Then
  Me.Pat_id = getPat_id(Me.PatName)
 End If
 If InStrB(Me.PatName, "|") <> 0 Then ' aufwändiges Raussuchen des Vorbriefdatums nur+immer bei Auswahl aus der Patientenliste mit der Maus
  Call Pat_ID_Change
 End If
End Sub ' PatName_Click()

Private Sub PatName_Change()
 Dim rs As New ADODB.Recordset, sn$(), s1$
' Me.Pat_id = getPat_id(Me.PatName)
 If InStrB(Me.PatName, ",") <> 0 Or InStrB(Me.PatName, " ") <> 0 Then
  SplitNeu Me.PatName, ",", sn
  If UBound(sn) = 0 Then
   SplitNeu Me.PatName, " ", sn
  End If
  If UBound(sn) > 0 Then s1 = sn(1) Else s1 = vNS
  myFrag rs, "SELECT COUNT(0) ct, MAX(pat_id) pid FROM `namen` WHERE nachname LIKE '" & UmwfSQL(sn(0)) & "%' AND vorname LIKE '" & UmwfSQL(s1) & "%'"
  If Not rs.EOF Then
   Me.Anzahl = rs!ct
   If Me.Anzahl > 0 Then
    Me.Pat_id = rs!pid
   Else
    Me.Pat_id = -1
   End If
  End If ' Not rs.EOF Then
 End If ' Not rs.EOF Then
End Sub ' PatName_Change()

Private Sub PatName_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 18 Then ' Alt-Taste
  Call Key(KeyCode, Shift, Me)
  Call PatName_Click
 End If
End Sub ' PatName_KeyDown(KeyCode As Integer, Shift As Integer)

