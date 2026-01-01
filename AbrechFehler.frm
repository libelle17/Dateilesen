VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form AbrechFehler 
   Caption         =   "Abrechnungsfehler"
   ClientHeight    =   11055
   ClientLeft      =   1245
   ClientTop       =   1050
   ClientWidth     =   15120
   LinkTopic       =   "AbrechFehler"
   ScaleHeight     =   11055
   ScaleWidth      =   15120
   Begin VB.CommandButton Automatisch 
      Caption         =   "Au&tomatisch"
      Height          =   375
      Left            =   14400
      TabIndex        =   13
      Top             =   0
      Width           =   1335
   End
   Begin VB.CheckBox Private 
      Caption         =   "&Private"
      Height          =   255
      Left            =   8640
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch ohne Speichern"
      Height          =   375
      Left            =   12480
      TabIndex        =   9
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton alleMarkieren 
      Caption         =   "alle &Demarkieren"
      Height          =   375
      Index           =   1
      Left            =   11040
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton alleMarkieren 
      Caption         =   "&alle Markieren"
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.CheckBox mitSQL 
      Caption         =   "&Liste"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CheckBox aktfDirekt 
      Caption         =   "Quaral &vorausgerechnet"
      Height          =   195
      Left            =   6360
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.CheckBox Debug 
      Caption         =   "A&usdruck"
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.ListBox Quartal 
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Start"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG 
      Height          =   11295
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   19923
      _Version        =   393216
      WordWrap        =   -1  'True
      MergeCells      =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label SQLin 
      Caption         =   "SQL in:"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Version 
      Height          =   255
      Left            =   12360
      TabIndex        =   10
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Quartallab 
      Caption         =   "&Quartal:"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "AbrechFehler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const obmitalterform = 0
'Const RegWurzel$ = "Software\GSProducts\"
Const J2a$ = "%", J1a$ = "_"
Const J2d$ = "*", J1d$ = "?"
Const lf$ = _
 "SELECT * FROM `faelle` INNER JOIN (SELECT MIN(pat_id) AS pid, MAX(bhfb) AS bb FROM `faelle` GROUP BY pat_id) AS sel " + _
 "ON (`faelle`.BhFB=sel.bb) AND (`faelle`.Pat_ID=sel.pid) "
Const tAusgS$ = "\Abrechnungsfehler"
Dim AktQ$    ' aktuelles Quartal
Dim tAusgSg$ ' Ausgabedatei mit Verzeichnis
Dim AbrVerz$ ' Verzeichnis davon
' Dim J2$, J1$
Dim CoSt$ ' ConnectionString
Dim FNr&  ' FehlerNummer
Dim AwN$() ' Auswahl-Namen
Dim sql$() ' SQL-Befehle
Dim obmo%() ' ob der aktuelle SQL-Befehl auf Medical Office aufgerufen werden soll
Dim mins%() ' Minimale Ausgabebreite
Dim maxs%() ' Maximale Ausgabebreite
Dim aktr&, altr&, alttop&, altC&
Dim noenter%
Dim fgespei% ' Farbe gespeichert
Dim altFarbe&
Dim AWz%, AWlf%
Dim nichtspeichern%
Dim sqlgezeigt%
#If obmitalterform Then
 Public AbrF As New AbrFForm
#End If

Public Function LiesDatei$(ByVal dname$)
Dim sLine$
Open dname For Input As #1
While Not EOF(1)
  Line Input #1, sLine
  If Left$(sLine, 2) <> "--" Then LiesDatei = LiesDatei & " " & sLine & vbCrLf
Wend
Close #1
End Function ' LiesDatei$(ByVal dname$)

'Public daw As New Datenauswahl
Private Sub aktfDirekt_Click()
 Call ZeigSQL
End Sub ' aktfDirekt_Click()

Private Sub Automatisch_Click()
 Call tuStart_click(1)
End Sub ' Automatisch_Click()

Private Sub Debug_KeyDown(KeyCode%, Shift%)
 Call KeyDown(KeyCode, Shift)
End Sub ' Debug_KeyDown(KeyCode%, Shift%)
Private Sub KeyDown(KeyCode%, Shift)
 If KeyCode = 27 Then End
End Sub ' keydown(KeyCode&, Shift)

Private Sub aktfDirekt_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me, Me.ActiveControl.name)
End Sub ' aktfDirekt_KeyDown(KeyCode%, Shift%)

Private Sub alleMarkieren_KeyDown(Index%, KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me, Me.ActiveControl.name)
End Sub ' alleMarkieren_KeyDown(Index%, KeyCode%, Shift%)

Private Sub altesFormular_Click()
#If obmitalterform Then
 AbrF.Show
#End If
End Sub ' altesFormular_Click()

Private Sub Einstellungen_Click()
' Stop
End Sub

Private Sub Form_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me, Me.ActiveControl.name)
End Sub ' Form_KeyDown(KeyCode%, Shift%)

'Private Sub Hintergrund_KeyDown(KeyCode As Integer, Shift As Integer)
' Call key(KeyCode, Shift, Me, Me.ActiveControl.Name)
'End Sub

Private Sub MFG_KeyDown(KeyCode%, Shift%)
 If KeyCode = 67 Then ' Ctrl+C
  Clipboard.SetText Me.MFG.Text
 Else
  Call Key(KeyCode, Shift, Me, Me.ActiveControl.name)
 End If
End Sub ' MFG_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Abbruch_Click()
 nichtspeichern = True
 Unload Me
' End
End Sub ' Private Sub Abbruch_Click()

Private Sub alleMarkieren_Click(Index As Integer)
 With MFG
 noenter = True
 .col = 1
 For AWlf = 1 To AWz
  .Row = AWlf
  If Index = 0 Then
   .TextMatrix(AWlf, 1) = "X"
   .CellBackColor = vbYellow
  Else
   .TextMatrix(AWlf, 1) = vNS
   .CellBackColor = vbWhite
  End If
 Next AWlf
 noenter = 0
 End With
End Sub ' Private Sub alleMarkieren_Click(Index As Integer)

Private Function cmd$(sql$, obAcc%)
 If obAcc Then
  cmd = sql
'  Cmd = replace$(Cmd, "`", "")
  cmd = REPLACE$(cmd, "date_add", "")
  cmd = REPLACE$(cmd, "DATE_SUB", "")
  cmd = REPLACE$(cmd, ",interval", "-")
  cmd = REPLACE$(cmd, ", interval", "+")
  cmd = REPLACE$(cmd, "INTERVAL ", " ")
  cmd = REPLACE$(cmd, "day)", ")")
  cmd = REPLACE$(cmd, "month)", "*30)")
  cmd = REPLACE$(cmd, "CAST(", "cdate(int(")
  cmd = REPLACE$(cmd, "STR_TO_DATE(", "cdate(")
  cmd = REPLACE$(cmd, ",'%d.%m.%Y'", "")
  cmd = REPLACE$(cmd, "as date", ")")
  cmd = REPLACE$(cmd, "_utf8", " ")
  cmd = REPLACE$(cmd, "concat", "")
  cmd = REPLACE$(cmd, "substr(", "MID(")
  cmd = REPLACE$(cmd, "¡", " & ")
  cmd = REPLACE$(cmd, "intaccdatemy", "int")
  cmd = REPLACE$(cmd, "intacc", "int")
  cmd = REPLACE$(cmd, "SELECTmy", "((")
  cmd = REPLACE$(cmd, "divmy", "/")
  cmd = REPLACE$(cmd, "to_days", "int")
'  Cmd = replace$(replace$(Cmd, "CAST(", ""), " As Date)", "")
  cmd = REPLACE$(cmd, "SUBDATE(", "(")
  cmd = REPLACE$(cmd, "adddate(", "(")
 '  Call cat.Views.Append(vn(i), Cmd)
  cmd = REPLACE$(cmd, "STR_TO_DATE(", "")
  cmd = REPLACE$(cmd, ",'#%m/%d/%Y#')", ")")
  cmd = REPLACE$(cmd, ",""#%m/%d/%Y#"")", ")")
'  Call Shell(App.Path & "\viewserst.exe """ & Forms(0).dbv.Datei & """ """ & QName & """ """ & Cmd & """")
 Else
  cmd = sql
  cmd = REPLACE$(cmd, "SELECTmy", "SELECT")
  cmd = REPLACE$(cmd, "divmy", "div")
  cmd = REPLACE$(cmd, "¡", ",")
  cmd = REPLACE$(cmd, "intaccdatemy", "date")
  cmd = REPLACE$(cmd, "intacc", "")
  cmd = REPLACE$(cmd, "cdate(", "(")
  cmd = REPLACE$(cmd, "first(", "(")
'  Call myEFrag("DROP VIEW " & ifexists & " `" & QName & "`;")
'  Call myEFrag("CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`" & Forms(0).dbv.uid & "`@`%` SQL SECURITY DEFINER VIEW `" & QName & "` AS " & Cmd)
 End If
End Function ' Cmd

Private Sub Form_Load()
 Me.WindowState = vbMaximized
 Me.Version = App.Major & "." & App.Minor & "." & App.Revision & ", Datenbank: " & DefDB(Lese.dbv.wCn)
 Call QuartalFüll
 AWz = 200
 On Error GoTo fehler
 ReDim sql(AWz - 1)
 ReDim obmo(AWz - 1)
 ReDim AwN(AWz - 1)
 ReDim mins(AWz - 1)
 ReDim maxs(AWz - 1)
' J2 = J2a
' J1 = J1a
#If obmitalterform Then
' If Me.AbrF.DAO Then J2 = J2d: J1 = J1d
 If Me.AbrF.ADOJet Then
  CoSt = CoJ
 Else
  Dim DBank$
  If Me.AbrF.AdoMySQL Then
   DBank = "quelle"
  ElseIf Me.AbrF.AdoMySQL1 Then
   DBank = "quelle1"
  ElseIf Me.AbrF.AdoMySQL2 Then
   DBank = "quelle2"
  End If
#Else
 Const DBank$ = vNS
 If True Then
#End If
  CoSt = Lese.dbv.cnVorb(DBank, "anamnesebogen", "Patientendaten")
 End If
 
 Dim erge
 erge = fWertLesen(HCU, RegWurzel & App.EXEName, "mitSQL")
 noenter = True
 If VarType(erge) = 8209 Then
  mitSQL = -CBool(erge(0))
 End If
 erge = fWertLesen(HCU, RegWurzel & App.EXEName, "aktfDirekt")
 noenter = True
 If VarType(erge) = 8209 Then
  aktfDirekt = -CBool(erge(0))
 End If
 
 noenter = 0
' tAusgSg = GetSpecialFolder(ssfPERSONAL) & tAusgS
 AbrVerz = pVerz & "Abrechnung"
 VerzPrüf AbrVerz
 tAusgSg = AbrVerz & tAusgS
' Call ZeigSQL
 
 Call SizeColumns(MFG, Me)
 Me.MFG.AllowUserResizing = 3 ' flexResizeBoth, in C:\Windows\SysWow64\MSHFLXGD.oca
' Me.KeyPreview = True

 Me.Visible = True '
  Call Form_Resize
 Me.MFG.SetFocus
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in form_load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Private Sub Form_Load()



' in ZeigSQL
Public Function ZeigSprivat(FristS$)
' 24.10.09: ändern in WHERE bhfe = '1899-12-30' AND schgr = 90
' ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Abrechnung von Beratungen (1,3)"
 sql(AWlf) = "SELECT n.pat_id,gesnameg(n.pat_id) Name,e.zeitpunkt zp, e.art art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid  " & vbCrLf & _
 "WHERE art IN (" & artspezG & ") " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt)) AND leistung IN ('1','3','30A','31A')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 '"WHERE art IN (" & artspezG & ") " & vbCrLf & _

' ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Zuschlag für Leistungen außerhalb der Sprechstunde -> A"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE art IN (" & artspezG & ") " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt)) AND leistung = 'A') " & vbCrLf & _
 "AND (day(e.zeitpunkt) IN (4,6) AND hour(e.zeitpunkt) BETWEEN 15 AND 19) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 ' ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Nachtzuschlag (20:00-22:00, 6:00-8:00 außerhalb Sprechstunde) -> (B)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE art IN (" & artspezG & ") " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung = 'B') " & vbCrLf & _
 "AND hour(e.zeitpunkt) IN (20,21,6) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 ' ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Tiefenachtzuschlag (22:00-6:00) -> (C)"
 sql(AWlf) = "SELECT n.pat_id AS Pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zp, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.pat_id = f.pat_id AND e.zeitpunkt BETWEEN qanf() AND qend() " & vbCrLf & _
 "WHERE art IN (" & artspezG & ") " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung = 'C') " & vbCrLf & _
 "AND NOT hour(e.zeitpunkt) BETWEEN 6 AND 22 " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Wochenendzuschlag (D)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, LEFT(DATE_FORMAT(e.zeitpunkt,'%a, %d.%m.%y %H:%i'),20) AS Datum, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE art IN (" & artspezG & ") " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung = 'D') " & vbCrLf & _
 "AND dayofweek(e.zeitpunkt) IN (7,1) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende ausführliche Beratung im Behandlungsfall (3)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, MIN(e.zeitpunkt) AS minzp, MAX(e.zeitpunkt) AS maxzp, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE art IN (" & artspezG & ") " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt)=DATE(e.zeitpunkt) AND leistung = '3') " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%') GROUP BY pat_id ORDER BY pat_id, e.zeitpunkt"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Untersuchung (5,6,7,8)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((art IN (" & artSpezEintr & ") " & vbCrLf & _
 "AND (inhalt LIKE '%ausk%' OR inhalt LIKE '%ödem%')) " & vbCrLf & _
 "OR (art IN (" & artSpezUS & "))) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt)=DATE(e.zeitpunkt) AND leistung IN ('5','6','7','8')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 AwN(AWlf) = "Evtl. fehlende Chronikerziffer (15)" ' AND COALESCE(d.Dggel,0)=0
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, MIN(e.zeitpunkt) AS minzp, MAX(e.zeitpunkt) AS maxzp, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "LEFT JOIN `diagview` d ON e.pat_id = d.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)))" & vbCrLf & _
 "WHERE ((art IN (" & artSpezEintr & ") " & vbCrLf & _
 ") " & vbCrLf & _
 "OR (art IN (" & artSpezUS & "))) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` l WHERE pat_id = e.pat_id AND YEAR(zeitpunkt) = YEAR(e.zeitpunkt) AND leistung = '15') " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%') GROUP BY pat_id ORDER BY pat_id, e.zeitpunkt"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Schulungen (20,33)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE art IN ('schul') " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt)=DATE(e.zeitpunkt) AND leistung IN ('20','33')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Untersuchungen in der Schwangerschaft (24)"
 '  AND COALESCE(d.Dggel,0)=0
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "LEFT JOIN `diagview` d ON e.pat_id = d.pat_id AND (d.icd LIKE 'O24%' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))" & vbCrLf & _
 "WHERE art IN (" & artSpezBerat & ") " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('24')) " & vbCrLf & _
 "AND NOT ISNULL(diagdatum) AND ADDDATE(diagdatum, INTERVAL 180 DAY) > e.zeitpunkt " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Arztbrief (75)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, b.zeitpunkt AS zeitpunkt " & vbCrLf & _
 "FROM `tmbrie` b LEFT JOIN `namen` n ON b.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON b.fid = f.fid " & vbCrLf & _
 "WHERE (b.name LIKE '%Arztbr%' OR b.name LIKE '%Nachr%') " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = b.pat_id AND zeitpunkt > b.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = b.pat_id AND DATE(zeitpunkt) = DATE(b.zeitpunkt) AND leistung IN ('75')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND ADDDATE(b.zeitpunkt, INTERVAL 365 DAY) > now()"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender KompressionsverbAND (204)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((inhalt LIKE '%Kompressionsv%' OR inhalt RLIKE '.*[^sHc]kv.*') OR (art LIKE '%kv%'))" & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('204')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Blutentnahme (250) nach Text"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((inhalt LIKE '%blutentn%' OR inhalt LIKE '%blutabn%') OR (art LIKE 'bz%'))" & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('250')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Blutentnahme (250) nach Laborleistung"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt " & vbCrLf & _
 "FROM `leistungen` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE (e.leistung='3585H1') " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('250')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Spritze (252, 253)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((inhalt LIKE '%i.v%' OR inhalt LIKE '%s.c%' OR inhalt LIKE '%i.m%' OR inhalt LIKE '%quadd%'))" & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('252','253')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Infusion (271,272,277,278)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((inhalt LIKE '%Alpha%' OR inhalt LIKE '%lipon%' OR inhalt LIKE '%thioc%' OR inhalt LIKE '%prostav%') OR false)" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('271','272','277','278')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Aderlaß (285)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((inhalt LIKE '%Aderla%' ) OR false)" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('285')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Duplexzuschlag nach Texteinträgen (401)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE art RLIKE 'duplex|dup|sono'" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('401')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = n.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Duplexzuschlag nach Bildern (401)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, b.quelldatum quelldatum, name " & vbCrLf & _
 "FROM tmbrie b LEFT JOIN `namen` n ON b.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON b.pat_id = f.pat_id " & vbCrLf & _
 "WHERE ((b.name LIKE '%sonobild %'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = b.pat_id AND zeitpunkt > b.quelldatum)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = b.pat_id AND DATE(zeitpunkt) = DATE(b.quelldatum) AND leistung IN ('401')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = n.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%') GROUP BY pat_id, quelldatum"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Frequenzanalysezuschlag nach Texteinträgen (404)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE art RLIKE '^dup$|duplex'" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('404')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Frequenzanalysezuschlag nach Bildern (404)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, b.quelldatum quelldatum, name " & vbCrLf & _
 "FROM tmbrie b LEFT JOIN `namen` n ON b.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON b.pat_id = f.pat_id " & vbCrLf & _
 "WHERE ((name LIKE '%sonobild %'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = b.pat_id AND zeitpunkt > b.quelldatum)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = b.pat_id AND DATE(zeitpunkt) = DATE(b.quelldatum) AND leistung IN ('404')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = n.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%') GROUP BY pat_id, quelldatum"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Sonoziffer Schilddrüse nach Texteinträgen (417)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((art LIKE '%sono%') AND (inhalt LIKE '%schild%' OR inhalt LIKE '%SD:%'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND ((DATE(zeitpunkt) > DATE(e.zeitpunkt)-7) AND (DATE(zeiptunkt) < DATE(e.zeitpunkt)+7)) AND leistung IN ('417','410')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Sonoziffer Schilddrüse nach Bildern (417)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, b.quelldatum AS quelldatum, name " & vbCrLf & _
 "FROM tmbrie b LEFT JOIN `namen` n ON b.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON b.pat_id = f.pat_id " & vbCrLf & _
 "WHERE ((name LIKE '%wchtl.sch%'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = b.pat_id AND zeitpunkt > b.quelldatum)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = b.pat_id AND DATE(zeitpunkt) = DATE(b.quelldatum) AND leistung IN ('417','410')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = n.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')  GROUP BY pat_id, quelldatum"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Sonoziffer Organ nach Texteinträgen (410)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((art LIKE '%sono%') AND NOT (inhalt LIKE '%schild%' OR inhalt LIKE '%SD:%'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('410')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = n.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Sonoziffer nach Bildern (410)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, b.quelldatum AS quelldatum, name " & vbCrLf & _
 "FROM tmbrie b LEFT JOIN `namen` n ON b.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON b.pat_id = f.pat_id " & vbCrLf & _
 "WHERE (name LIKE '%SonoBild%' AND name LIKE '% Abd%')" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = b.pat_id AND zeitpunkt > b.quelldatum)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = b.pat_id AND DATE(zeitpunkt) = DATE(b.quelldatum) AND leistung IN ('410')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = n.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Sonoziffer Organ nach Texteinträgen (420)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((art LIKE '%sono%') AND NOT (inhalt LIKE '%schild%' OR inhalt LIKE '%SD:%'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('420')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlende Lufu (605, 605a)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((art LIKE '%lufu%'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND (NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('605')) " & vbCrLf & _
 "OR NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('605A'))) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlendes EKG (651)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((art LIKE '%EKG%'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('651')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
'ktag fehlerhaft
 AwN(AWlf) = "Evtl. fehlender Neurostatus (800)"
 sql(AWlf) = "SELECT n.pat_id AS pat_id, gesname(n.pat_id) Name, e.zeitpunkt AS zeitpunkt, e.art AS art, inhalt " & vbCrLf & _
 "FROM `eintraege` e LEFT JOIN `namen` n ON e.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` f ON e.fid = f.fid " & vbCrLf & _
 "WHERE ((art LIKE '%usd%'))" & vbCrLf & _
 "AND (false OR NOT EXISTS (SELECT pat_id FROM `eintraege` bez WHERE art = 'rech' AND pat_id = e.pat_id AND zeitpunkt > e.zeitpunkt)) " & vbCrLf & _
 "AND NOT EXISTS (SELECT pat_id FROM `leistungen` WHERE pat_id = e.pat_id AND DATE(zeitpunkt) = DATE(e.zeitpunkt) AND leistung IN ('800')) " & vbCrLf & _
 "AND n.nachname <> 'ZuTun' AND n.pat_id <> 2 AND n.pat_id <> 2 AND schgr = 90 AND zeitpunkt > (SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = e.pat_id AND art = 'pvs' AND NOT inhalt LIKE '%storniert%')"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
 AwN(AWlf) = "Evtl. fehlender CGM (A659)"
 sql(AWlf) = _
 "SELECT i.pat_id, gesname(i.pat_id) PName, QD, tx FROM ( " & vbCrLf & _
 " SELECT f0.pat_id " & vbCrLf & _
 " , DATE(qdm) QD " & vbCrLf & _
 " , LEFT(b.name,40) tx " & vbCrLf & _
 " FROM (SELECT pat_id,bhfe1,bhfb FROM faelle WHERE schgr=90 AND bhfe1=18991230 AND bhfb>SUBDATE(NOW(),500) GROUP BY pat_id) f0 " & vbCrLf & _
 " LEFT JOIN tmbrie b ON b.pat_id=f0.pat_id AND b.name RLIKE 'Omnipod ?5|Simplera|Guardian|Enlite|Dexc?om|Eversen|CGM|Libre|clarity|glooko|care ?link' AND NOT b.name RLIKE 'cgm bmp' AND b.qdm>" & qtAnf(FristS) & " " & vbCrLf & _
 " UNION ALL " & vbCrLf & _
 " SELECT f1.pat_id " & vbCrLf & _
 " , DATE(fr.zeitpunkt) QD " & vbCrLf & _
 " , fr.feldinh tx " & vbCrLf & _
 " FROM (SELECT pat_id,bhfe1,bhfb FROM faelle WHERE schgr=90 AND bhfe1=18991230 AND bhfb>SUBDATE(NOW(),500) GROUP BY pat_id) f1 " & vbCrLf & _
 " LEFT JOIN formular fr ON fr.pat_id=f1.pat_id AND fr.form_abk IN ('prp','plar') AND fr.feld RLIKE 'medikament|verordnungszeile|txtMedKey' AND fr.feldinh RLIKE 'Omnipod ?5|Simplera|Guardian|Enlite|Dexc?om|Eversen|CGM|Libre|clarity|glooko|care ?link' AND fr.zeitpunkt>" & qtAnf(FristS) & " " & vbCrLf & _
 " UNION ALL " & vbCrLf & _
 " SELECT f2.pat_id " & vbCrLf & _
 " , DATE(e.zeitpunkt) QD " & vbCrLf & _
 " , e.inhalt tx " & vbCrLf & _
 " FROM (SELECT pat_id,bhfe1,bhfb FROM faelle WHERE schgr=90 AND bhfe1=18991230 AND bhfb>SUBDATE(NOW(),500) GROUP BY pat_id) f2 " & vbCrLf & _
 " LEFT JOIN eintraege e ON e.pat_id=f2.pat_id AND e.inhalt RLIKE 'Omnipod ?5|Simplera|Guardian|Enlite|Dexc?om|Eversen|CGM|Libre|clarity|glooko|care ?link' AND NOT e.inhalt LIKE '%vorgestellt%will%nicht%' AND e.zeitpunkt>" & qtAnf(FristS) & " " & vbCrLf & _
 ") i " & vbCrLf & _
 " LEFT JOIN leistungen l ON  l.pat_id=i.pat_id AND DATE(l.zeitpunkt)=i.qd AND leistung='A659' " & vbCrLf & _
 "WHERE NOT ISNULL(qd) AND ISNULL(leistung) ORDER BY pat_id,qd;"
 mins(AWlf) = 5
 maxs(AWlf) = 100
 AWlf = AWlf + 1
End Function

Public Function ZeigSQL() ' Abrechnungsfehler
'Const nname$ = "LEFT(CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')),IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')),n.nachname,', ',n.vorname),25) Name, "
Dim FristS$
AktQ = Me.Quartal
FristS = Str(Int(Now() - fctQEnd(AktQ) + 1))
Static aktf$, qanfS$, qendS$
On Error GoTo fehler
AWlf = 0
If Me.Private <> 0 Then
 ZeigSprivat (FristS)
 
Else ' me.private = 0 => Kassenpatienten
' faq$ = "SELECT * FROM `faelle` WHERE quartal = '" + aktQ + "'"
' Dim aktf$, qanfS$, qendS$
' BetrPausch = "'97350A','97360A','97350B','97360B','97370B','97371B', '97310','97320','97321','97312','97322','97313','97323','97333'"
 Const BetrPausch$ = "'97310','97312','97320','97321','97322','97333'"
 If aktfDirekt <> 0 Then
  aktf = "(SELECT pat_id,nachname,vorname,quartal,fid,schgr,goäkatnr,ik,vknr,kid FROM `faelle` WHERE schgr <> '90' AND NOT goäkatnr IN ('40','41') AND nachname <> 'Bereitschaftsdienst' AND quartal = '" & AktQ & "') AS f "
 Else
  aktf = "(SELECT pat_id,nachname,vorname, quartal,fid,schgr,goäkatnr,ik,vknr,kid FROM `faelle` WHERE schgr <> '90' AND NOT goäkatnr IN ('40','41') AND nachname <> 'Bereitschaftsdienst' AND quartal = (SELECTmy CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & FristS & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & FristS & " DAY)))) " & vbCrLf & _
         "ORDER BY pat_id, fid DESC, schgr) AS f "
  aktf = cmd(aktf, InStrB(Lese.dbv.CnStr, "MySQL") = 0) ' .wCn.ConnectionString ' 28.12.08
 End If
 
 If aktfDirekt <> 0 Or InStrB(Lese.dbv.CnStr, "MySQL") = 0 Then ' .wCn.ConnectionString ' 28.12.08
  qanfS = DatFor_k(fctQAnf(AktQ))
  qendS = DatFor_k(fctQEnd(AktQ) + 1)
 Else
  qanfS = "TIMESTAMP((CONCAT(YEAR((NOW() - INTERVAL " & FristS & " DAY))¡ '-'¡ (intacc(((month((NOW() - INTERVAL " & FristS & " DAY)) - 1) divmy 3) * 3) + 1), '-01'))) "
  qendS = "TIMESTAMP((CONCAT(YEAR((NOW() - INTERVAL " & FristS & " DAY))¡'-'¡(intacc(((month((NOW() - INTERVAL " & FristS & " DAY)) - 1) divmy 3) * 3) + 1),'-01') + INTERVAL 3 MONTH)) "
  qanfS = cmd(qanfS, InStrB(Lese.dbv.CnStr, "MySQL") = 0) ' .wCn.ConnectionString ' 28.12.08
  qendS = cmd(qendS, InStrB(Lese.dbv.CnStr, "MySQL") = 0) ' .wCn.ConnectionString ' 28.12.08
 End If ' aktfDirekt <> 0 OR  else
  
  ' für Ärzte:
  
' sql(AWlf) = _
' "SELECT anf1.pat_id PatID, gesnameg(anf1.pat_id) Name, diabetestyp DTyp, icd, quartal " + _
'"FROM (SELECT an.*,lf.pat_id AS lfpid,lf.fid,quartal " & vbCrLf & _
'"FROM `anamnesebogen` AS an INNER JOIN (" + lf + ") AS lf " + _
' "ON an.`pat_id`=lf.pat_id) AS anf1 " + _
' "LEFT JOIN (SELECT d.icd, d.diagicherheit, d.Dggel, F0.pat_id AS f0pid, F0.fid AS F0fid FROM (`diagnosen` d LEFT JOIN `faelle` F0 " + _
' "ON d.pat_id = F0.pat_id AND (d.fid = F0.fid OR ISNULL(d.fid))) " + _


' neuView
AwN(AWlf) = "Diagnosen"
sql(AWlf) = "ü"
 AWlf = AWlf + 1

' 1
 AwN(AWlf) = "Widersprüchlicher Diabetestyp Anamnese/ ICD (quartalsübergreifende Abfrage!) (vorher 0)"
 ' AND COALESCE(d.Dggel,0)=0
 ' SELECT pat_id, gesname(pat_id), diabetestyp, a.* FROM anamnesebogen a WHERE pat_id IN
 ' <dann in notepad die Spalte markieren mit Alt+Shift+Maus> usw.
 sql(AWlf) = _
"SELECT Pat_id PID, PName, wer, LEFT(Diabetestyp,10) DTyp, IF(ISNULL(anam),'(null)',anam) `DTyp aus andm`, MaxHbA1c, CONCAT(g0,' ',g2) OGTT, icd " & vbCrLf & _
"FROM ( " & vbCrLf & _
" SELECT f.pat_id, gesname(f.pat_id) Pname, az.wer, a.diabetestyp," & vbCrLf & _
" TRIM(MID(e.inhalt,INSTR(e.inhalt,'Diabetes Typ ')+12,20)) Anam," & vbCrLf & _
" IF(xH.max1>xH.max2,xH.max1, xH.max2) maxHbA1c, " & vbCrLf & _
" IF(xG.max1>xG.max2,xG.max1, xG.max2) maxGluc, g0(IF(LENGTH(og.inhalt)>1000,LEFT(og.inhalt,1000),og.inhalt)) g0, g2(IF(LENGTH(og.inhalt)>1000,LEFT(og.inhalt,1000),og.inhalt)) g2, " & vbCrLf & _
" (SELECT MIN(icd) FROM diagview d WHERE d.pat_id=f.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)))) icd " & vbCrLf & _
"FROM aktfvz f " & vbCrLf & _
"LEFT JOIN anamnesebogen a ON a.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN `_maxHbA1c` xH ON xH.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN `_maxGluc` xG ON xG.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id = f.pat_id AND e.art RLIKE '^andm' AND e.ZeitPunkt=(SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=f.pat_id AND art RLIKE '^andm')" & vbCrLf & _
"LEFT JOIN eintraege og ON f.pat_id = og.pat_id AND og.art = 'ogtt' " & vbCrLf & _
"LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
") i WHERE " & vbCrLf & _
"NOT (diabetestyp IN ('1','1?','s') AND icd LIKE 'E10%') AND " & vbCrLf & _
"NOT (diabetestyp IN ('2','2?','s','mody') AND icd LIKE 'E11%') AND " & vbCrLf & _
"NOT (diabetestyp IN ('-','','u','p') AND (ISNULL(icd) OR icd='O24.4')) AND " & vbCrLf & _
"NOT (diabetestyp IN ('-','p','','?','g') AND (ISNULL(icd) OR icd='R73.0')) AND " & vbCrLf & _
"NOT (diabetestyp = 's' AND icd LIKE 'E13%') AND " & vbCrLf & _
"NOT (diabetestyp='g' AND icd ='O24.4');"
 mins(AWlf) = 5
 maxs(AWlf) = 18
 AWlf = AWlf + 1

 ' die 1000 s. Views g0, g2 und _vorNr
 ' vorher: lfaellev statt aktfvs

' 2
 AwN(AWlf) = "Möglicherweise fehlende Diabetes/Gestationsdiabetes-Diagnosen (vorher 1)"
 sql(AWlf) = _
"SELECT f.pat_id, gesname(f.pat_id) PName, wer, xH.Wert HbA1cMax, xG.Wert GlucMax, dd.icd `DM-ICD`, gd.icd `GDM-ICD` " & vbCrLf & _
"FROM aktfvz f " & vbCrLf & _
"LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
"LEFT JOIN `maxHbA1c` xH ON xH.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN `maxGluc` xG ON xG.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN diagview dd " & vbCrLf & _
" ON dd.pat_id = f.pat_id AND dd.gICD RLIKE '^E1[0-4]\.' AND dd.obdauer<>0 " & vbCrLf & _
"LEFT JOIN diagview gd " & vbCrLf & _
" ON gd.pat_id = f.pat_id AND (gd.icd='O24.4' AND gd.Dggel=0 AND gd.diagsicherheit IN ('G',' ') AND gd.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))" & " " & vbCrLf & _
"WHERE " & vbCrLf & _
" (xH.Wert>6.4 AND ISNULL(dd.icd)) OR (xG.Wert>199 AND ISNULL(dd.icd) AND ISNULL(gd.icd)) " & vbCrLf & _
"OR (NOT ISNULL(dd.icd) AND NOT ISNULL(gd.icd))"

 'sql(AWlf) = _
 "SELECT * FROM (" & vbCrLf & _
 "SELECT a.pat_id, " & qtAnf(FristS) & ", GesName(a.pat_id), xH.wert maxHbA1c, xG.wert maxGluc, REPLACE(REPLACE(n.notiz,char(13),''),char(10),'') Notiz, IF(f.letzteregel<>'',DATE_FORMAT(ADDDATE(f.letzteRegel,274),'%e.%c.%y'),'') TdE, " & vbCrLf & _
 "(SELECT GROUP_CONCAT(DISTINCT g.leistung) FROM leistungen l LEFT JOIN genehmigungen g ON l.leistung = g.leistung AND g.leistung NOT IN (33060,33061,33076) AND l.zeitpunkt > SUBDATE(NOW(),INTERVAL 120 DAY) WHERE l.pat_id = a.pat_id) leistung," & vbCrLf & _
 "GROUP_CONCAT(CONCAT(info.diagdatum,' ',info.diagsicherheit,' ',info.diagtext)) diagtext, MAX(info.diagdatum) diagdat FROM aktfvs a " & vbCrLf & _
 "LEFT JOIN faelle f ON a.fid = f.fid " & vbCrLf & _
 "LEFT JOIN namen n ON n.pat_id = a.pat_id " & vbCrLf & _
 "LEFT JOIN diagnosen d ON a.pat_id = d.pat_id AND ((d.icd RLIKE '^E1[0-4]' AND d.obdauer<>0 AND d.diagsicherheit NOT IN ('A','Z') AND COALESCE(d.Dggel,0)=0) OR (d.icd LIKE 'O24%' AND d.diagsicherheit NOT IN ('A','Z') AND COALESCE(d.Dggel,0)=0 AND d.fid = a.fid)) " & vbCrLf & _
 "LEFT JOIN diagnosen abkl ON a.pat_id = abkl.pat_id AND abkl.icd = 'Z13.1' AND abkl.diagsicherheit <> 'A' AND COALESCE(abkl.Dggel,0)=0 " & vbCrLf & _
 "LEFT JOIN diagnosen info ON info.pat_id = a.pat_id AND info.diagsicherheit NOT IN ('A','Z') AND info.diagsicherheit=0 " & vbCrLf & _
 "                    AND info.diagdatum > SUBDATE(f.qanf,INTERVAL 90 DAY) AND info.diagtext LIKE '%gesta%' " & vbCrLf & _
 "LEFT JOIN `maxHbA1c` xH ON xH.pat_id = a.pat_id " & vbCrLf & _
 "LEFT JOIN `maxGluc` xG ON xG.pat_id = a.pat_id " & vbCrLf & _
 "WHERE ISNULL(d.ICD) AND ISNULL(abkl.ICD) " & vbCrLf & _
 "GROUP BY a.pat_id) i " & vbCrLf & _
 "WHERE (maxHbA1c > 6.4 AND NOT ISNULL(maxHbA1c))" & vbCrLf & _
 "   OR (maxGluc > 199 AND NOT ISNULL(maxGluc)) " & vbCrLf & _
 "   OR notiz<>'' OR (NOT ISNULL(leistung)) OR (diagdat>" & qtAnf(FristS) & ") " & vbCrLf & _
 "ORDER BY maxHbA1c DESC" '  OR INSTR(diagtext,'schwang')<>0
 mins(AWlf) = 5
 maxs(AWlf) = 18
 AWlf = AWlf + 1
 
' 3
 AwN(AWlf) = "Fehlende Gestationsdiabetes-ICD nach Makroeintrag 'vkgd' (vorher 4)"
 sql(AWlf) = "SELECT f.pat_id, gesname(f.pat_id) PName, wer " & vbCrLf & _
             "FROM `aktfvz` f " & vbCrLf & _
             "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
             "LEFT JOIN `eintraege` e ON f.pat_id = e.pat_id AND e.zeitpunkt BETWEEN qanf() AND qend() " & vbCrLf & _
             "LEFT JOIN diagview d ON ((d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)) OR (f.pat_id = d.pat_id AND (d.icd RLIKE '^E1[0-4]') AND d.diagsicherheit NOT IN ('A','Z','V'))) " & vbCrLf & _
             "LEFT JOIN namen n ON f.pat_id = n.pat_id WHERE e.art IN ('vkgd','vkgd2') AND ISNULL(icd) GROUP BY f.pat_id" '  AND COALESCE(d.Dggel,0)=0
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1

' 4
AwN(AWlf) = "evtl. fehlende O24.4G nach OGTT (voher 102)"
' AND COALESCE(d.Dggel,0)=0
sql(AWlf) = vbCrLf & _
"SELECT * FROM (" & vbCrLf & _
" SELECT e.inhalt, g0(e.inhalt) g0,g1(e.inhalt) g1,g2(e.inhalt) g2, s.voret, f.pat_id,gesname(f.pat_id) PName, wer" & vbCrLf & _
", (SELECT MAX(icd) FROM diagview d WHERE d.pat_id=f.pat_id AND (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)) " & vbCrLf & _
"  ) gdd, dd.icd" & vbCrLf & _
" FROM aktfvz f" & vbCrLf & _
" LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
" LEFT JOIN sws s ON s.pat_id=f.pat_id AND s.voret>qanf()" & vbCrLf & _
" LEFT JOIN diagview dd ON dd.pat_id=f.pat_id AND dd.gicd RLIKE '^E1[0-4]\.' AND dd.obdauer<>0 " & vbCrLf & _
" LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.art = 'ogtt' AND e.zeitpunkt BETWEEN SUBDATE(" & qtAnf(FristS) & ",INTERVAL 180 DAY) AND " & qtEnd(FristS) & " AND e.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=f.pat_id AND art='ogtt') " & vbCrLf & _
" AND f.schgr<>90 AND NOT ISNULL(s.voret)" & vbCrLf & _
" ) i " & vbCrLf & _
" WHERE (i.g0 >= 92 OR i.G1 >= 180 OR i.G2 >= 153)" & vbCrLf & _
" AND ISNULL(gdd) AND ISNULL(icd) " & vbCrLf & _
"ORDER BY pat_id"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
'" LEFT JOIN faelle et ON f.pat_id=et.pat_id AND et.voret BETWEEN " & KhtsflZuk() & " " & vbCrLf & _

 ' 5
 AwN(AWlf) = "Diabetestypen E13 und E14 (vorher 2)"
 sql(AWlf) = "SELECT d.pat_id,GesNameg(d.pat_id) PName,wer,d.diagdatum Datum,d.diagsicherheit Si, d.DiagText,d.ICD " & vbCrLf & _
             "FROM `diagview` d" & vbCrLf & _
             "LEFT JOIN anaktk az ON az.pid=d.pat_id" & vbCrLf & _
             "LEFT JOIN faelle f ON f.pat_id = d.pat_id AND d.diagdatum BETWEEN qanf() AND qend() " & vbCrLf & _
             "LEFT JOIN `diagview` r ON d.pat_id = r.pat_id AND r.gicd RLIKE '^E1[01]'" & vbCrLf & _
             "WHERE d.gzICDok RLIKE '^E1[234]' AND ISNULL(r.gicd) AND d.diagsicherheit NOT IN ('A','Z') AND schgr<> 90" ' AND COALESCE(d.Dggel,0)=0
 mins(AWlf) = 5
 maxs(AWlf) = 18
 AWlf = AWlf + 1
 
' 6
 AwN(AWlf) = "DMP-Diabetes-Patienten ohne Diabetes-ICD (vorher 5)"
 sql(AWlf) = "SELECT f.pat_id, gesname(f.pat_id) PName, wer " & vbCrLf & _
             "FROM aktfvz f " & vbCrLf & _
             "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
             "LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
             "LEFT JOIN `diagview` d ON f.pat_id = d.pat_id AND d.gICD RLIKE '^E1[0-4]\.' AND d.obdauer<>0 WHERE ((dmpklass=2 OR (dmpklass=3 AND dmpbeg<=qend()))) AND ISNULL(d.gicd)"
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1


' 7
AwN(AWlf) = "Möglicherweise doppelte Diabetesdiagnosen (vorher 8)"
'Dggel=1 => 'ddg' steht in der Art-Spalte
'bestimmte Simultandiagnosen nötig für Nephropathiepauschale
'sql(AWlf) = "SELECT a.Pat_ID, d.GesName, COUNT(0) Zahl, GROUP_CONCAT(CONCAT(icd,diagsicherheit,diagattr)) ICDs FROM `aktfvz` a LEFT JOIN `diagnosen` d ON d.pat_id = a.pat_id AND diagsicherheit IN (' ','G','V') AND COALESCE(Dggel,0)=0 AND icd RLIKE '^E1[0-4]|^O24' AND (obdauer<>0 OR d.fid = a.fid OR d.fid=0 OR ISNULL(d.fid)) GROUP BY a.pat_id HAVING zahl<>1"
' AND COALESCE(d.Dggel,0)=0
' sql(AWlf) = "SELECT pid, gesnameg(pid) PName, wer, zahl, icd, e.Art " & vbCrLf & _
            "FROM (" & vbCrLf & _
            "SELECT f.Pat_ID PID, wer, COUNT(0) Zahl, GROUP_CONCAT(CONCAT(icd,diagsicherheit,diagattr)) ICD, " & vbCrLf & _
            "(SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = f.pat_id AND (art IN ('gs','tk','ah')OR(art='tb'AND ersteller IN ('gs','tk','ah')) OR inhalt LIKE '%(gs)%' OR inhalt LIKE '%(tk)%' OR inhalt LIKE '%(ah)%')) lEintr " & vbCrLf & _
            "FROM `aktfvz` f " & vbCrLf & _
            "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
            "LEFT JOIN `diageview` d ON d.pat_id = f.pat_id AND ((d.gicd REGEXP '^E1[0-4]\.|^R73' AND d.obdauer<>0 AND d.diagsicherheit='G' AND d.diagdatum > qbegs(f.quartal)) OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)))" & vbCrLf & _
            "GROUP BY f.pat_id HAVING zahl<>1) i " & vbCrLf & _
            "LEFT JOIN `eintraege` e ON i.pid = e.pat_id AND e.zeitpunkt = i.leintr AND (art IN ('tk','gs','ah')OR(art='tb'AND ersteller IN ('gs','tk','ah')) OR inhalt LIKE '%(tk)%' OR inhalt LIKE '%(gs)%' OR inhalt LIKE '%(ah)%') " & vbCrLf & _
            "WHERE NOT (icd RLIKE 'E1[01].7[^5]' AND icd RLIKE 'E1[01].75') AND NOT (icd RLIKE 'E1[01].7[^4]' AND icd RLIKE 'E1[01].74') AND NOT (icd RLIKE 'E1[01].2' AND icd RLIKE 'E1[01].[^2]') " & vbCrLf & _
            "GROUP BY pid,leintr" & vbCrLf & _
            "HAVING NOT ICD RLIKE 'O24.4.*R73.0' AND NOT ICD RLIKE 'R73.0.*O24.4'"
#If langsam Then
sql(AWlf) = "" & _
"WITH sel AS (SELECT fbehgrundnr fb,FEintragsart ea,FIcdcode ICD FROM ltag WHERE fpatnr IN (SELECT fpatnr FROM patfall f WHERE 18900101+INTERVAL FVon DAY>=STR_TO_DATE(CONCAT(YEAR(NOW()-INTERVAL " & Verspätung & " DAY),LPAD((MONTH(NOW()-INTERVAL " & Verspätung & " DAY)-1)DIV 3*3+1,2,'0'),'01'),'%Y%m%d')))" & vbCrLf & _
"SELECT * FROM (" & vbCrLf & _
" SELECT FPatnr" & vbCrLf & _
" ,LEFT(CONCAT(IF(FTitel='','',CONCAT(FTitel,' ')),IF(FNamensvorsatz='','',CONCAT(FNamensvorsatz,' ')),FNachname,', ',FVorname),25) PatName" & vbCrLf & _
" ,GROUP_CONCAT(icd ORDER BY FDatum SEPARATOR' | ') ICDs" & vbCrLf & _
" ,GROUP_CONCAT(18900101 + INTERVAL FDatum DAY ORDER BY FDatum SEPARATOR' | ') Diagdaten" & vbCrLf & _
" ,GROUP_CONCAT(CASE FStatus WHEN 1 THEN 'ak' WHEN 2 THEN 'an' WHEN 3 THEN 'hi' WHEN 4 THEN 'ab' WHEN 5 THEN 'da' ELSE ' ' END ORDER BY FDatum SEPARATOR' | ')STATUS" & vbCrLf & _
" ,GROUP_CONCAT(COALESCE(IF(RIGHT(FErlaeuterung,1)=0,LEFT(FErlaeuterung,LENGTH(FErlaeuterung)-1),FErlaeuterung),'')ORDER BY FDatum SEPARATOR ' | ') Zusätze" & vbCrLf & _
" ,GROUP_CONCAT(nu.FInitialen ORDER BY FDatum SEPARATOR' | ') Nutzer" & vbCrLf & _
" FROM sel lt" & vbCrLf & _
" LEFT JOIN behgrund b ON b.FSurogat=lt.fb" & vbCrLf & _
" LEFT JOIN nutzerneu nu ON nu.FSurogat= b.FNutzernr" & vbCrLf & _
" LEFT JOIN patstamm p ON p.FSurogat=b.FPatnr" & vbCrLf & _
" WHERE FStatus<>3 AND lt.ICD=b.FIcdcode" & vbCrLf & _
" AND lt.icd RLIKE '^E1[0-4]\.'" & vbCrLf & _
" AND 18900101 + INTERVAL FDatum DAY + INTERVAL FZeit SECOND >=STR_TO_DATE(CONCAT(YEAR(NOW()-INTERVAL " & Verspätung & " DAY),LPAD((MONTH(NOW()-INTERVAL " & Verspätung & " DAY)-1)DIV 3*3+1,2,'0'),'01'),'%Y%m%d')" & vbCrLf & _
" AND CASE (FKlasse MOD 15)MOD 10 WHEN 1 THEN 'V' WHEN 2 THEN 'G' WHEN 3 THEN 'Z' WHEN 4 THEN 'A' ELSE ' ' END IN ('G',' ')" & vbCrLf
sql(AWlf) = sql(AWlf) & _
" GROUP BY FPatnr" & vbCrLf & _
")i WHERE icds RLIKE '\|'" & vbCrLf & _
";" & vbCrLf & _
""
#Else
sql(AWlf) = _
"SELECT f.FPatnr" & _
",LEFT(CONCAT(IF(FTitel='','',CONCAT(FTitel,' ')),IF(FNamensvorsatz='','',CONCAT(FNamensvorsatz,' ')),FNachname,', ',FVorname),25) PatName" & vbCrLf & _
",GROUP_CONCAT(lt.FICDCode ORDER BY b.FDatum SEPARATOR' | ') ICDs" & vbCrLf & _
",GROUP_CONCAT(18900101 + INTERVAL b.FDatum DAY ORDER BY b.FDatum SEPARATOR' | ') Diagdaten" & vbCrLf & _
",GROUP_CONCAT(CASE b.FStatus WHEN 1 THEN 'ak' WHEN 2 THEN 'an' WHEN 3 THEN 'hi' WHEN 4 THEN 'ab' WHEN 5 THEN 'da' ELSE ' ' END ORDER BY b.FDatum SEPARATOR' | ')STATUS" & vbCrLf & _
",GROUP_CONCAT(COALESCE(IF(RIGHT(FErlaeuterung,1)=0,LEFT(FErlaeuterung,LENGTH(FErlaeuterung)-1),FErlaeuterung),'')ORDER BY b.FDatum SEPARATOR ' | ') Zusätze" & vbCrLf & _
",GROUP_CONCAT(nu.FInitialen ORDER BY b.FDatum SEPARATOR' | ') Nutzer" & vbCrLf & _
"FROM patfall f" & vbCrLf & _
"LEFT JOIN ltag lt USING (fpatnr)" & vbCrLf & _
"LEFT JOIN behgrund b ON b.FSurogat=lt.FBehgrundnr" & vbCrLf & _
"LEFT JOIN nutzerneu nu ON nu.FSurogat=b.FNutzernr" & vbCrLf & _
" LEFT JOIN patstamm p ON p.FSurogat=b.FPatnr" & vbCrLf & _
"WHERE f.FSurogat=(SELECT MAX(fsurogat) FROM patfall WHERE fpatnr=f.fpatnr AND 18900101+INTERVAL Fbis DAY>=STR_TO_DATE(CONCAT(YEAR(NOW()-INTERVAL " & Verspätung & " DAY),LPAD((MONTH(NOW()-INTERVAL " & Verspätung & " DAY)-1)DIV 3*3+1,2,'0'),'01'),'%Y%m%d'))" & vbCrLf & _
"AND b.FStatus<>3 AND b.FStatus<>4 AND lt.FICDCode=b.FIcdcode" & vbCrLf & _
"AND lt.FICDCode RLIKE '^E1[0-4]\.'" & vbCrLf & _
"AND 18900101+INTERVAL b.FDatum DAY + INTERVAL b.FZeit SECOND >=STR_TO_DATE(CONCAT(YEAR(NOW()-INTERVAL " & Verspätung & " DAY),LPAD((MONTH(NOW()-INTERVAL " & Verspätung & " DAY)-1)DIV 3*3+1,2,'0'),'01'),'%Y%m%d')" & vbCrLf & _
"AND CASE (FKlasse MOD 15)MOD 10 WHEN 1 THEN 'V' WHEN 2 THEN 'G' WHEN 3 THEN 'Z' WHEN 4 THEN 'A' ELSE ' ' END IN ('G',' ')" & vbCrLf & _
"GROUP BY FPatnr" & vbCrLf & _
"HAVING icds RLIKE '\|'" & vbCrLf & _
""
#End If

 mins(AWlf) = 10
 maxs(AWlf) = 60
 obmo(AWlf) = True
 AWlf = AWlf + 1
 
 ' 8
 ' diagsicherheit und Dggel in aktfaelle.icd schon eingebaut
 AwN(AWlf) = "Unverwertbare Diabetes-Klassifikation (E12, E13, E14, O24.4) bei 'DMP hier ...' oder 'DMP HA ...' -> Korrektur z.B. 'D.m.(sekundär)' + ICD E10.91 (vorher 3)"
 sql(AWlf) = "SELECT f.pat_id, gesname(f.pat_id) PName, wer, dmpklass, f.icd icd " & vbCrLf & _
    "FROM `aktfaellev` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id  " & vbCrLf & _
    "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
    "WHERE NOT (f.icd RLIKE '^E1[01]\.|^R73|^O24.4')"
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1

 ' 9
 AwN(AWlf) = "Fehlende Fußdiagnose trotz Leistung 97314, 97324 oder 02311 (vorher 6)"
 sql(AWlf) = _
 "SELECT eid, Pat_id, Patient, wer, ICD, Quelldatum, FotoStadium, FotoSeite, Eintr_Seite, Eintr_Zeitpunkt, Art, Eintr_Inhalt, Leistn, name FROM (" & vbCrLf & _
 " SELECT e.id eid, f.pat_id, gesname(f.pat_id) Patient, wer, MAX(di.icd) ICD, Quelldatum" & vbCrLf & _
 " , CONCAT('L89.',CAST(MAX(IF(MID(name,INSTR(name,'WA ')+3,1)>2,2,MID(name,INSTR(name,'WA ')+3,1)))+1 AS CHAR),IF(INSTR(name,'Ferse')<>0,'7','8')) Fotostadium" & vbCrLf & _
 " , REPLACE(REPLACE(GROUP_CONCAT(DISTINCT IF(INSTR(name,' li.')<>0,MID(name,INSTR(name,' li.')+1,2),IF(INSTR(name,' re.')<>0,MID(name,INSTR(name,' re.')+1,2),''))),',',''),'lire','bds') FotoSeite" & vbCrLf & _
 " , IF(INSTR(e.inhalt,' bds.')<>0,' bds',IF(INSTR(e.inhalt,' li ')<>0,' li', IF(INSTR(e.inhalt,' re ')<>0,' re',''))) Eintr_Seite" & vbCrLf & _
 " , e.zeitpunkt Eintr_Zeitpunkt, e.Art, e.inhalt Eintr_Inhalt" & vbCrLf & _
 ", Leistn" & vbCrLf & _
 ", b.name" & vbCrLf & _
 " FROM `aktf` f " & vbCrLf & _
 " JOIN `faelle` fa USING (fid)" & vbCrLf & _
 " LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
 " LEFT JOIN `namen` n ON f.pat_id = n.pat_id" & vbCrLf & _
 " LEFT JOIN (SELECT pat_id,GROUP_CONCAT(DISTINCT CONCAT(DATE_FORMAT(ZeitPunkt,'%d.%m.%y'),': ',leistung) SEPARATOR ', ') Leistn FROM leistungen WHERE leistung IN ('97314','97324','02311') AND zeitpunkt BETWEEN qanf() AND qend() GROUP BY pat_id) l ON l.pat_id = f.pat_id" & vbCrLf & _
 " LEFT JOIN tmbrie b ON f.pat_id = b.pat_id AND b.quelldatum BETWEEN fa.qanf AND fa.qend AND b.name RLIKE 'WA [0-5]'" & vbCrLf & _
 " LEFT JOIN `eintraege` e ON f.pat_id = e.pat_id AND DATE(e.zeitpunkt) BETWEEN fa.qanf AND fa.qend AND (e.art RLIKE 'debr|ulcus' OR (e.inhalt RLIKE 'ebrid|resekt' AND NOT inhalt RLIKE 'Leber.*rese[ck]t|Resektionshöhle|Resektion der Schild|Pan[ck]reas.*resekt|Resektion Leberzyste|Resektionsbereich|Re[ck]tumrese[ck]t|Rese[ck]tion Lunge|(darm|magen|milz|acg-|schilddrüsen|struma|sd-|nieren|teil|links|total|gebärmutter|mu[ck]osa|gallenblasen|elektro|wurzel|truma|igma|pan[ck]reaskopf)resekt|Hypoph.*resekt'))" & vbCrLf & _
 " LEFT JOIN `diagview` di ON f.pat_id = di.pat_id AND di.ICD RLIKE '^L89\.[1-5]' AND (obdauer<>0 OR DATE(di.diagdatum) BETWEEN qbeg(b.quelldatum) AND qende(b.quelldatum))" & vbCrLf & _
 " WHERE not ISNULL(Leistn)" & vbCrLf & _
 " GROUP BY f.pat_id,e.art, e.inhalt, b.name" & vbCrLf & _
 " HAVING ISNULL(ICD) OR (NOT ISNULL(fotostadium) AND (ICD COLLATE utf8mb4_german2_ci < Fotostadium ))" & vbCrLf & _
 " ) i" & vbCrLf & _
 "GROUP BY pat_id"
' "" & vbCrLf & _
' gestrichen vor Fotostadium: "COLLATE utf8mb4_general_ci" (dieses führt hier zu einem SQL-Fehler, wenn auch nicht in HeidiSQL

 ' AND obdauer = 0
 ' -- 'WA |WA -'
 ' ausrangiert 7.1.23:
 ' "-- , CONCAT('L89.',CAST(MAX(IF(REPLACE(MID(name,INSTR(name,'WA ')+3,1),'-','0')>2,2,REPLACE(MID(name,INSTR(name,'WA ')+3,1),'-','0')))+1 AS char),IF(INSTR(name,'Ferse')<>0,'7','8')) Fotostadium" & vbCrLf & _

 mins(AWlf) = 7
 maxs(AWlf) = 120
 AWlf = AWlf + 1

 ' 10
AwN(AWlf) = "Podologierezepte u.ä. ohne Neuropathiediagnose G63.2 u.ä. (vorher 24)"
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id) PName, wer" & vbCrLf & _
",GROUP_CONCAT(DISTINCT CASE WHEN ra.Medikament RLIKE 'podo' THEN 'Pod' WHEN ra.Medikament RLIKE 'Orthon' THEN 'Nagel' WHEN ra.Medikament RLIKE 'orthop|therapies' THEN 'Schuh' WHEN ra.Medikament RLIKE 'bettu' THEN 'Bett' ELSE '' END) Rp_akt" & vbCrLf & _
",date_format(ra.ZeitPunkt,'%d.%m.%y') `verordnet am`,IF(ISNULL(d.ICD),'G62.9','G63.2') `Fehlende ICD`" & vbCrLf & _
"FROM aktfvz f" & vbCrLf & _
"LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
"LEFT JOIN rezepteintraege ra ON ra.pat_id=f.pat_id AND ra.rezkllang IN ('Langrezepteintrag','Heilmittelverordnung') AND ra.medikament RLIKE 'podo|bettu|therapies|ortho[np]' AND ra.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"LEFT JOIN diagview d ON d.pat_id=f.pat_id AND d.obdauer<>0 AND d.gICD REGEXP '^E1[0-4]\.'" & vbCrLf & _
"LEFT JOIN diagview nd ON nd.pat_id=f.pat_id AND nd.obdauer<>0 AND nd.gICD = 'G63.2'" & vbCrLf & _
"LEFT JOIN diagview na ON na.pat_id=f.pat_id AND na.obdauer<>0 AND na.gICD RLIKE '^G6[0-3]'" & vbCrLf & _
"WHERE (NOT ISNULL(d.icd) AND ISNULL(nd.icd)) OR (ISNULL(d.icd) AND ISNULL(na.ICD))" & vbCrLf & _
"GROUP BY f.pat_id" & vbCrLf & _
"HAVING Rp_akt<>''" & vbCrLf & _
""
' AwN(AWlf) = "Fehlende 97314 (Fußsyndr.b.Typ1), 97324 (b.Typ2) für Fußsyndrom nach Diagnosen"
' sql(AWlf) = _
 "SELECT * FROM " & vbCrLf & _
 "(SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, l.leistung AS Leistung, kateg, d.icd, diab.icd AS DTyp FROM " & aktf & " " & vbCrLf & _
 "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('97314','97324')) AS l ON f.fid = l.fid " & vbCrLf & _
 "LEFT JOIN `diagnosen` diab ON (f.pat_id = diab.pat_id AND diab.obdauer <> 0 AND diab.diagsicherheit<> 'A' AND COALESCE(diab.Dggel,0)=0) AND diab.icd REGEXP '^E1[0-4]\.' " & vbCrLf & _
 "LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & " AND d.diagsicherheit NOT IN ('A','Z') AND COALESCE(d.Dggel,0)=0 AND d.icd REGEXP '^L89[1234]\.' " & vbCrLf & _
 "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr AND f.ik=kl.ik " & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
 "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
 "WHERE (ISNULL(leistung)) " & vbCrLf & _
 "AND NOT ISNULL(dtyp) " & vbCrLf & _
 "AND NOT ISNULL(icd) "
' sql(AWlf) = sql(AWlf) & " WHERE NOT ISNULL(icd)"

 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1
 
  
 ' 11
 AwN(AWlf) = "Untertriebene Niereninsuffizienz => N18.3 oder N18.4 (vorher 85)"
 sql(AWlf) = "SELECT f.pat_id, gesname(f.pat_id) PName, wer, _lGFR(f.pat_id) eGFR, d.ICD " & vbCrLf & _
 "FROM aktfvz f " & vbCrLf & _
 "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
 "LEFT JOIN diagview d ON f.pat_id = d.pat_id AND d.gicd IN ('N18.1','N18.2')" & vbCrLf & _
 "WHERE COALESCE(_lgfr(f.pat_id),0) BETWEEN 1 AND 59 AND NOT ISNULL(d.icd) "
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 12
 AwN(AWlf) = "Fehlende Ulcus-Diagnose nach Fuß-Makro (vorher 96)"
 '  AND COALESCE(ni.Dggel,0)=0
 sql(AWlf) = "" & vbCrLf & _
 "SELECT u.pat_id, gesnameg(u.pat_id) PName, wer, u.zeitpunkt Makrozt, u.ulcera, di.icd, u.mitarbeiter " & vbCrLf & _
 "FROM aktfvz f " & vbCrLf & _
 "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
 "LEFT JOIN fuss u ON f.pat_id = u.pat_id AND u.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
 "LEFT JOIN diagview di ON f.pat_id = di.pat_id AND di.diagdatum BETWEEN qanf() AND qend() AND (di.icd RLIKE '^L89\.[1-5]'" & _
 "OR ((SELECT MAX(icd) FROM diagview dd  WHERE dd.pat_id = f.pat_id AND dd.gicd REGEXP '^E1[0-4]') IS NULL AND di.gicd RLIKE '^L97') " & _
 ") AND obdauer = 0 AND di.diagdatum BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
 "WHERE ulcera <>'nein' " & vbCrLf & _
 " AND zeitpunkt > 20170717080000 " & vbCrLf & _
 " AND ISNULL(icd) " & vbCrLf & _
 "ORDER BY zeitpunkt;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 13
Dim ulcusicd$ ' nur in 'Fehlende Ulcus-Diagnose'
ulcusicd = "'^L89\.[123]|T14.9|L02.9'"
 AwN(AWlf) = "Fehlende Ulcus-Diagnose (" & ulcusicd & ") nach Ulcus-Makro (voher 97)"
 ' AND COALESCE(dd.Dggel,0
 sql(AWlf) = "" & vbCrLf & _
 "SELECT e.pat_id, gesnameg(e.pat_id) PName, wer, e.zeitpunkt Makrozt, e.Inhalt " & vbCrLf & _
 "FROM aktfvz f " & vbCrLf & _
 "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
 "LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
 "LEFT JOIN diagview di ON f.pat_id = di.pat_id AND di.diagdatum BETWEEN qanf() AND qend() AND (di.gicd RLIKE " & ulcusicd & " " & _
 "OR ((SELECT MAX(icd) FROM `diagview` dd  WHERE dd.pat_id = f.pat_id AND dd.gicd REGEXP '^E1[0-4]' ) IS NULL AND di.gicd RLIKE '^L97') " & _
 ") AND obdauer = 0 AND di.diagdatum BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
 "WHERE art='ulcus' AND inhalt NOT LIKE '%Lokalisation: Bauch%' AND ISNULL(di.icd) " & vbCrLf & _
 "ORDER BY e.pat_id DESC, zeitpunkt;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 14
AwN(AWlf) = "Fehler bei ICD oder Schwangerschaftseintrag bei Schwangeren mit D.m. oder GDM (vorher 101)"
' AND COALESCE(d.Dggel,0)=0
sql(AWlf) = vbCrLf & _
"SELECT i.pat_id PID, PName,wer, voret Entbindung,Eintr Eintrag,dicd `Diabetes-ICD`,sicd `Sws-ICD(s)`, sdicd `Sws-Dauer-ICD` " & vbCrLf & _
", concat(IF((dicd='' AND sicd not RLIKE '^O24.4')  OR (dicd RLIKE '^E10' AND sicd NOT LIKE 'O24.0%') OR (dicd RLIKE '^E11' AND sicd NOT LIKE 'O24.1%'),IF(dicd='','Sws-ICD fehlt','Diabetes-ICD passt nicht zu Sws-ICD,'),'')," & vbCrLf & _
"         IF(sicd='','Sws-ICD fehlt,',''),IF(voret='','Entbindung fehlt,',''),IF(sicd LIKE '%,%','Sws-Diagnose zu oft eingetragen,',''),IF(sdicd<>'',concat(sdicd,' falsch als Dauerdiagnose'),'')) Fehler " & vbCrLf & _
"FROM (" & vbCrLf & _
"SELECT pat_id, gesname(pat_id) PName, wer " & vbCrLf & _
", COALESCE((SELECT MAX(voret) FROM sws WHERE voret>" & qtAnf(FristS) & " AND pat_id=v.pat_id),'') voret " & vbCrLf & _
", COALESCE((SELECT GROUP_CONCAT(art) FROM eintraege e WHERE v.pat_id=e.pat_id AND e.art RLIKE '^vkgd|^aufgd' AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & "),'') Eintr " & vbCrLf & _
", COALESCE((SELECT GROUP_CONCAT(icd) FROM diagview d WHERE v.pat_id=d.pat_id AND d.gicd RLIKE '^E1[0-4]' AND obdauer<>0 ),'') dicd " & vbCrLf & _
", COALESCE((SELECT GROUP_CONCAT(icd) FROM diagview d WHERE v.pat_id=d.pat_id AND (d.icd LIKE 'O24.%' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(v.quartal) AND qends(v.quartal))),'') sicd " & vbCrLf & _
", COALESCE((SELECT GROUP_CONCAT(icd) FROM diagview d WHERE v.pat_id=d.pat_id AND (d.icd LIKE 'O24.%' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(v.quartal) AND qends(v.quartal)) AND obdauer<>0),'') sdicd " & vbCrLf & _
"FROM aktfvz v" & vbCrLf & _
"LEFT JOIN anaktk az ON az.pid=v.pat_id" & vbCrLf & _
") i " & vbCrLf & _
"WHERE (voret<>'' OR sicd<>'' OR eintr<>'') AND NOT (dicd='' AND sicd='') " & vbCrLf & _
"AND ((dicd='' AND sicd NOT RLIKE '^O24.4') OR (dicd RLIKE '^E10' AND sicd NOT LIKE 'O24.0%') OR (dicd RLIKE '^E11' AND sicd NOT LIKE 'O24.1%') " & vbCrLf & _
"      OR sicd='' OR voret='' OR sdicd<>'') " & vbCrLf & _
"OR sicd LIKE '%,%'"

'sql(AWlf) = vbCrLf & _
"SELECT f.pat_id,gesname(f.pat_id) Name, e.art,d.icd,d.obdauer=0 obqd,dm.icd DM, et.voret, ISNULL(d.icd) `O-Diagn.fehlt`, MID(d.icd,5,1)!=MID(dm.icd,3,1) FalscheICD, d.obdauer<>0 FaleschDauer, ISNULL(et.voret) KeinVorETermin FROM aktfvz f " & vbCrLf & _
"LEFT JOIN eintraege e ON f.pat_id=e.pat_id AND e.art IN ('vkgd','vkgd2','aufgd') AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"LEFT JOIN diagnosen d ON d.pat_id=f.pat_id AND d.icd LIKE 'O24%' AND d.diagsicherheit IN ('g',' ') AND d.Dggel<>1 AND d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"LEFT JOIN diagnosen dm ON dm.pat_id=f.pat_id AND dm.icd RLIKE '^E1[0-4]' AND dm.obdauer<>0 AND dm.Dggel<>1 " & vbCrLf & _
"LEFT JOIN faelle et ON f.pat_id=et.pat_id AND et.voret BETWEEN " & KhtsflZuk() & " " & vbCrLf & _
"WHERE (NOT ISNULL(art) OR NOT ISNULL(d.icd) OR (NOT ISNULL(voret) AND NOT ISNULL(dm.icd))) AND ( " & vbCrLf & _
"(ISNULL(d.icd) OR (NOT ISNULL(dm.icd) AND (ISNULL(d.icd)) AND MID(d.icd,5,1)!=MID(dm.icd,3,1))) OR " & vbCrLf & _
"(d.obdauer<>0) OR " & vbCrLf & _
"(ISNULL(voret)) " & vbCrLf & _
") GROUP BY f.pat_id"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 15
 AwN(AWlf) = "Fehlende Diabetes-Entgleisung im ICD-Code (vorher 129)"
 ' AND COALESCE(Dggel,0)=0
sql(AWlf) = _
"SELECT f.Pat_id,gesname(f.pat_id) PName,wer,h.lzp HbA1cZp, h.letzter lHbA1c,d.icd " & vbCrLf & _
"FROM aktfvz f " & vbCrLf & _
"LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
"LEFT JOIN lHbA1c h USING (pat_id) " & vbCrLf & _
"LEFT JOIN diagview d ON d.pat_id=f.pat_id AND d.gicd RLIKE '^E1[0-4]' AND d.obdauer<>0 " & vbCrLf & _
"WHERE (h.letzter>7.4 OR false) AND (d.icd RLIKE '^E1[0-4]\.[0-6]0' OR false);"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
 ' 16
 AwN(AWlf) = "nicht mehr aktuelle Nieren-ICDs und/oder herabzustufende EBM-Ziffern (vorher 130)"
 ' AND COALESCE(Dggel,0)=0
 sql(AWlf) = _
 "SELECT pat_id, gesname(pat_id) PName, wer, max1, eGFR" & vbCrLf & _
 ",IF(icd1='','',CONCAT(icd1,' -> ',IF(max1>20,'V.a.','Z.n.'))) ICD_1 " & vbCrLf & _
 ",IF(icd2='','',CONCAT(icd2,' -> ',IF(egfr>59,'Z.n.','V.a.'))) ICD_2 " & vbCrLf & _
 ",IF(fallei='','',CONCAT(fallei,' -> ',97600)) Leistung " & vbCrLf & _
 "FROM " & vbCrLf & _
 "(SELECT f.pat_id, wer, MAX1,_lGFR(f.pat_id) egfr,obnephrop(f.pat_id) obne " & vbCrLf & _
 ",COALESCE((SELECT GROUP_CONCAT(icd) FROM diagview WHERE pat_id=f.pat_id AND gicd RLIKE '^N1[89]' AND obdauer<>0 ),'') ICD1 " & vbCrLf & _
 ",COALESCE((SELECT GROUP_CONCAT(icd) FROM diagview WHERE pat_id=f.pat_id AND gicd ='N08.3' AND obdauer<>0),'') ICD2 " & vbCrLf & _
 ",COALESCE((SELECT MAX(leistung) FROM leistungen WHERE pat_id = f.pat_id AND zeitpunkt>" & qtAnf(FristS) & " AND leistung IN ('97601','97602')),'') fallei " & vbCrLf & _
 "FROM aktfvz f LEFT JOIN _AlbCreMinMax USING (pat_id)" & vbCrLf & _
 "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
 ") i " & vbCrLf & _
 "WHERE (ICD1<>'' OR ICD2<>'') " & vbCrLf & _
 " AND obne=0 AND egfr<1000 " & vbCrLf & _
 " AND NOT EXISTS (SELECT leistung FROM leistungen WHERE pat_id = i.pat_id AND zeitpunkt<" & qtAnf(FristS) & " AND leistung IN ('97601','97602')) " & vbCrLf & _
 " AND NOT EXISTS (SELECT icd FROM diagnosen WHERE pat_id = i.pat_id AND icd = 'Z94.0' AND diagsicherheit IN ('G','Z',' ')) " & vbCrLf & _
 "" ' kein schon dokumentiertes pathologisches DAK-Modul, keine Nierentransplantation
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

' 17
 AwN(AWlf) = "Pat. mit 'DMP HA' ohne Diabetes-ICD oder ohne nachweisbaren Überweiser (vorher 132)"
 ' AND COALESCE(Dggel,0)=0
 ' n.getha0
 
sql(AWlf) = _
"SELECT * FROM (SELECT f.pat_id, gesnameg(f.pat_id)" & vbCrLf & _
", COALESCE((SELECT ICD FROM diagview" & vbCrLf & _
"  WHERE Pat_ID=f.Pat_ID AND gICD REGEXP '^E1[01]' AND (obDauer<>0 OR pat_id = f.pat_id AND diagdatum BETWEEN qanf() AND qend()) " & vbCrLf & _
"  ORDER BY ICD " & vbCrLf & _
"  LIMIT 1),'fehlt') ICD " & vbCrLf & _
",f.übwvbsnr kvnr, CONCAT(h.name,', ',h.vorname,', ',h.ort) Üw " & vbCrLf & _
",Eintr,dt Desktop,termin,Azu " & vbCrLf & _
"FROM (" & vbCrLf & _
"  SELECT f.*, CASE WHEN gszul>tkzul AND gsz>tkz THEN 'gs' WHEN gszul<tkzul AND gsz<tkz THEN 'tk' ELSE '' END Eintr " & vbCrLf & _
"  FROM ( " & vbCrLf & _
"   SELECT f.pat_id,f.fid,f.quartal,f.schgr,f.goäkatnr,f.nachname,f.übwvlanr,f.übwvbsnr " & vbCrLf & _
"   , COALESCE((SELECT CASE WHEN lanrid=1 THEN 'gs' WHEN lanrid=2 THEN 'tk' WHEN lanrid=5 THEN 'ah' ELSE '?' END FROM faelle WHERE pat_id=f.pat_id AND lanrid IN (1,2,5) ORDER BY bhfb DESC LIMIT 1),'') Azu " & vbCrLf & _
"   , (SELECT COUNT(0) FROM eintraege WHERE (((art IN ('gs','doppler','dop','duplex','dup') OR (art='tb' AND ersteller='gs')) AND NOT inhalt LIKE '%(tk)%') OR inhalt LIKE '%(gs)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)) gsz " & vbCrLf & _
"   , COALESCE((SELECT DATE(MAX(zeitpunkt)) FROM eintraege WHERE (((art IN ('gs','doppler','dop','duplex','dup') OR (art='tb' AND ersteller='gs')) AND NOT inhalt LIKE '%(tk)%') OR inhalt LIKE '%(gs)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)),0) gszul " & vbCrLf & _
"   , (SELECT COUNT(0) FROM eintraege WHERE ((art IN ('tk','ARCHIE2','APK') OR (art='tb' AND ersteller='tk')) OR inhalt LIKE '%(tk)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)) tkz " & vbCrLf & _
"   , COALESCE((SELECT DATE(MAX(zeitpunkt)) FROM eintraege WHERE ((art IN ('tk','ARCHIE2','APK') OR (art='tb' AND ersteller='tk')) OR inhalt LIKE '%(tk)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)),0) tkzul " & vbCrLf & _
"   , COALESCE((SELECT CONCAT(MAX(IF(raum='Schade','gs','tk')),' ', DATE_FORMAT(MAX(zp),'%e.%c.%y')) FROM termine t WHERE pid = f.pat_id AND zp =(SELECT MIN(zp) FROM termine WHERE zp>=CURRENT_TIMESTAMP() AND pid=t.pid AND raum IN ('Kothny','Schade') AND NOT zusatz RLIKE 'doppler|duplex|dup')),'') Termin " & vbCrLf & _
"   ,TRIM(CONCAT(IF(obtk,'tk ',''),IF(obgs,'gs ',''),IF(obdw,'wd|ah ',''))) dt " & vbCrLf & _
"   FROM aktfvz af JOIN faelle f USING (fid)" & vbCrLf & _
"   LEFT JOIN anakt ON anakt.pat_id=f.pat_id " & vbCrLf & _
"  ) f " & vbCrLf & _
" ) f " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
"LEFT JOIN liuez h ON h.kvnri=f.übwvbsnr AND f.übwvbsnr<>0 " & vbCrLf & _
"LEFT JOIN `desktop` dt ON n.pat_id = dt.pat_id AND dt.titel LIKE '%kein%bericht%' " & vbCrLf & _
"WHERE f.quartal IN ( " & vbCrLf & _
"(SELECT CONCAT((MONTH(CURRENT_TIMESTAMP() - INTERVAL " & Verspätung & " DAY) - 1) DIV 3 + 1, YEAR(CURRENT_TIMESTAMP() - INTERVAL " & Verspätung & " DAY))), " & vbCrLf & _
"(SELECT CONCAT((MONTH(CURRENT_TIMESTAMP() - INTERVAL " & (Verspätung + 90) & " DAY) - 1) DIV 3 + 1, YEAR(CURRENT_TIMESTAMP() - INTERVAL " & (Verspätung + 90) & " DAY)))) " & vbCrLf & _
"AND f.schgr<>'90' AND f.`GOÄKatNr` NOT IN ('40','41') AND f.nachname<>'Bereitschaftsdienst'" & vbCrLf & _
"AND ISNULL(dt.titel) AND n.dmpklass=2) i WHERE (kvnr=0 OR NOT ICD REGEXP '^E1[01]') " & vbCrLf & _
"GROUP BY pat_id " & vbCrLf & _
"ORDER BY kvnr DESC " & vbCrLf & _
";"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

' 18
AwN(AWlf) = "Fehlende Diagnosen bei Impflingen (vorher 142)"
sql(AWlf) = "" & _
"  SELECT pat_id, gesname(pat_id) PName, rr, DATE(rr.zeitpunkt) Datum " & vbCrLf & _
"  , case when RRsyst<140 AND RRdiast<95 THEN 'Ausschluß arter.Hypertonie (I10.90A)' ELSE 'V.a. art. Hypertonie (I10.90A)' END Diagnose " & vbCrLf & _
"   FROM aktfvz f " & vbCrLf & _
"  LEFT JOIN rr USING (pat_id) " & vbCrLf & _
"  WHERE NOT ISNULL(rr.rr) " & vbCrLf & _
"  AND EXISTS(SELECT * FROM diagnosen WHERE pat_id=f.pat_id AND icd='Z26.9') " & vbCrLf & _
"  AND NOT EXISTS(SELECT * FROM diagnosen WHERE pat_id=f.pat_id AND icd<>'Z26.9') " & vbCrLf & _
"  GROUP BY f.pat_id ORDER BY diagnose,pat_id; "
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 19
 AwN(AWlf) = "Kernschatten ohne CLL oder Ausschluss CLL (vorher 156)"
sql(AWlf) = "" & _
"SELECT IF(lwer=wer AND NOT ISNULL(lwer),'',CONCAT(CHR(13),CHR(10))) ``" & vbCrLf & _
" , i.pat_id,IF(wer='-','?',wer) Arzt,Name,Patalter,Zp,`%`,lbeh" & vbCrLf & _
"  FROM (SELECT " & vbCrLf & _
"  l.pat_id, az.Wer, wia(l.pat_id) labArzt, gesname(l.pat_id) NAME, Patalter(l.pat_id) patalter" & vbCrLf & _
"  , GROUP_CONCAT(l.zeitpunkt SEPARATOR '|') zp, MAX(zeitpunkt) mzp" & vbCrLf & _
"  , GROUP_CONCAT(l.wert SEPARATOR '|') `%` ,date(lbeh) lbeh" & vbCrLf & _
"  , CASE when wer='-' then 0 WHEN wer='tk' THEN 1 WHEN wer='ah' THEN 2 ELSE 3 END rh" & vbCrLf & _
"  , LAG(wer) OVER(ORDER BY rh,mzp DESC) lwer" & vbCrLf & _
"  FROM labor2a l " & vbCrLf & _
"  LEFT JOIN diagnosen d ON d.Pat_id=l.pat_id AND d.icd LIKE 'C91%' " & vbCrLf & _
"  LEFT JOIN anamnesebogen a ON a.pat_id=l.pat_id" & vbCrLf & _
"  LEFT JOIN anaktk az ON az.pid=a.pat_id" & vbCrLf & _
"  LEFT JOIN namenlb n ON n.pat_id=a.pat_id" & vbCrLf & _
"  WHERE abkü='ksm' AND ISNULL(d.icd) AND a.tkz<>1 " & vbCrLf & _
"  GROUP BY pat_id " & vbCrLf & _
" ) i ORDER BY rh, mzp DESC;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 20
 AwN(AWlf) = "Kodierrichtlinien (1): Diabetes-ICD"
 sql(AWlf) = "" & _
"SELECT i.pat_id,i.PName,wer,ddia" & vbCrLf & _
", CASE WHEN dsnr IN (2,3,4,5,6,7) THEN IF((drei<>dsnr AND drei<>7) OR (dsnr=7 AND ((obfu=5 AND vier NOT IN (4,5)) ) OR obfu=5 AND drei<>7),CONCAT(zwei,if(obfu=5,7,dsnr),IF(obfu=5 OR dsnr=7,obfu,1)),'')" & vbCrLf & _
"       WHEN dsnr=9 THEN IF(obfu=3,'',IF(drei<>7 OR vier NOT IN (4,5),CONCAT(zwei,7,5),''))" & vbCrLf & _
"  END soll" & vbCrLf & _
"  , grund,dsnr,obfu" & vbCrLf & _
"FROM (" & vbCrLf & _
"SELECT f.pat_id, gesname(f.pat_id) PName, ddia, zwei, drei,vier" & vbCrLf & _
"        , CASE WHEN obnp+oblu+obap+obfl+obau+obne>1 THEN '7'" & vbCrLf & _
"                                    WHEN obnp OR oblu THEN '4'" & vbCrLf & _
"                                    WHEN obap THEN '5'" & vbCrLf & _
"                                    WHEN obfl THEN '6'" & vbCrLf & _
"                                    WHEN obau THEN '3'" & vbCrLf & _
"                                    WHEN obne THEN '2'" & vbCrLf & _
"                                    Else '9' END dsnr" & vbCrLf & _
", obfu, grund" & vbCrLf
sql(AWlf) = sql(AWlf) & _
"FROM (SELECT f.*, divd.icd ddia, LEFT(divd.icd,4) zwei, MID(divd.icd,5,1) drei, MID(divd.icd,6,1) vier" & vbCrLf & _
"            , max(IF(ISNULL(npvd.icd),0,1)) obnp" & vbCrLf & _
"            , max(IF(ISNULL(luvd.icd),0,1)) oblu" & vbCrLf & _
"            , max(IF(ISNULL(apvd.icd),0,1)) obap" & vbCrLf & _
"            , max(IF(ISNULL(flvd.icd),0,1)) obfl" & vbCrLf & _
"            , max(IF(ISNULL(auvd.icd),0,1)) obau" & vbCrLf & _
"            , max(IF(ISNULL(nevd.icd),0,1)) obne" & vbCrLf & _
"            , max(IF(ISNULL(fuvd.icd),3,5)) obfu" & vbCrLf & _
"            ,CONCAT(COALESCE(GROUP_CONCAT(distinct npvd.ICD),''),',',COALESCE(GROUP_CONCAT(distinct luvd.ICD),''),',',COALESCE(GROUP_CONCAT(distinct apvd.ICD),''),','" & vbCrLf & _
"                   ,COALESCE(GROUP_CONCAT(distinct flvd.ICD),''),',',COALESCE(GROUP_CONCAT(distinct auvd.ICD),''),',',COALESCE(GROUP_CONCAT(distinct nevd.ICD),''),','" & vbCrLf & _
"                     ,COALESCE(GROUP_CONCAT(distinct fuvd.ICD),'')) Grund" & vbCrLf & _
"FROM aktfvz f" & vbCrLf
sql(AWlf) = sql(AWlf) & _
"LEFT JOIN (SELECT pat_id, icd FROM (" & vbCrLf & _
"  SELECT pat_id, icd,RANK() OVER (PARTITION BY pat_id ORDER BY IF(MID(icd,5,1)='9','-1',MID(icd,5,1)) DESC,MID(icd,6,1) DESC) rang FROM diagview WHERE gicd RLIKE '^E1[0-4]' AND obdauer<>0" & vbCrLf & _
"          ) d WHERE rang=1) divd ON divd.pat_id=f.pat_id" & vbCrLf & _
"LEFT JOIN diagview npvd ON npvd.pat_id=f.pat_id AND npvd.gicd IN ('G59.0','G63.2','G73.0','G99.0') AND npvd.obdauer<>0" & vbCrLf & _
"LEFT JOIN diagview luvd ON luvd.pat_id=f.pat_id AND luvd.gicd IN ('N31.1','N31.2') AND luvd.obdauer<>0" & vbCrLf & _
"LEFT JOIN diagview apvd ON apvd.pat_id=f.pat_id AND apvd.gicd IN ('I79.2') AND apvd.obdauer<>0" & vbCrLf & _
"LEFT JOIN diagview flvd ON flvd.pat_id=f.pat_id AND flvd.gicd IN ('K77.8','M14.2','M14.6') AND flvd.obdauer<>0" & vbCrLf & _
"LEFT JOIN diagview auvd ON auvd.pat_id=f.pat_id AND auvd.gicd IN ('H28.0','H36.0') AND auvd.obdauer<>0" & vbCrLf & _
"LEFT JOIN diagview nevd ON nevd.pat_id=f.pat_id AND (nevd.gicd RLIKE '^N18\.[1-5]' OR nevd.gicd='N08.3') AND nevd.obdauer<>0" & vbCrLf & _
"LEFT JOIN diagview fuvd ON fuvd.pat_id=f.pat_id AND fuvd.gicd RLIKE '^L89' AND fuvd.obdauer<>0" & vbCrLf & _
"GROUP BY f.pat_id, divd.icd" & vbCrLf & _
") f WHERE NOT ISNULL(ddia)" & vbCrLf & _
") i " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"LEFT JOIN anaktk az ON az.pid=i.pat_id" & vbCrLf & _
"HAVING soll<>'';"
 mins(AWlf) = 10
 maxs(AWlf) = 600
 AWlf = AWlf + 1
 
' "-- LEFT JOIN diagview divd ON divd.pat_id=f.pat_id AND divd.gicd RLIKE '^E1[0-4]\.' AND divd.obdauer<>0" & vbCrLf & _
"-- ORDER BY f.pat_id, MID(divd.icd,5,1) DESC" & vbCrLf & _

' 21
 AwN(AWlf) = "Kodierrichtlinien (2): Myokardinfarkt"
 sql(AWlf) = "" & _
 "SELECT f.Pat_ID, gesname(f.pat_id) PName, az.wer, date_format(Diagdatum,'%e.%c.%y') DiagDatum,DiagSicherheit DS, Diagtext, ICD" & vbCrLf & _
 "FROM aktfvz f" & vbCrLf & _
 "LEFT JOIN anaktk az ON az.pid=f.pat_id" & vbCrLf & _
 "LEFT JOIN diagview d ON d.pat_id=f.pat_id" & vbCrLf & _
 "WHERE gicd LIKE 'i21%' OR (gicd LIKE 'i25.2' AND diagsicherheit='z')"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 22
 AwN(AWlf) = "Nicht mit R32 kodierte, im ADL eingetragene Harninkontinenz (vorher 151)"
 ' AND COALESCE(Dggel,0)=0
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id) PName " & vbCrLf & _
",(SELECT concat(feldinh,'   ',zeitpunkt) FROM formular WHERE pat_id=f.pat_id AND LCASE(formvorl)='\$\\turbomed\\formulare\\patientenmenue\\adl.tmf' AND feldnr IN (20,21,22) AND feld='txtdyn' ORDER BY zeitpunkt DESC, foid DESC LIMIT 1) Harnkontinenzpunkte " & vbCrLf & _
",(SELECT icd FROM diagview d WHERE pat_id=f.pat_id AND gicd LIKE 'R32%' LIMIT 1) R32 " & vbCrLf & _
"FROM aktfvz f " & vbCrLf & _
"HAVING LEFT(Harnkontinenzpunkte,2)<>10 AND ISNULL(R32);"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
'",(SELECT icd FROM diagnosen d WHERE pat_id=f.pat_id AND icd LIKE 'N31%' AND diagsicherheit IN ('G',' ') AND COALESCE(Dggel,0)=0 LIMIT 1) N31 " & vbCrLf & _

' 23
 AwN(AWlf) = "Möglicherweise nachzutragende Pflegestufendiagnose Z74.9 (vorher 62)"
' AND COALESCE(dd.Dggel,0)=0
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patAlter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'And ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'And ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'And ersteller='ah'))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",GROUP_CONCAT(pfl.inhalt SEPARATOR ' | ') Pflegestufentext " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0123]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gicd LIKE 'Z74%' " & vbCrLf & _
            "LEFT JOIN eintraege pfl ON pfl.pat_id=v.pat_id AND  pfl.inhalt RLIKE 'pflege(grad|stufe)' AND pfl.inhalt NOT RLIKE 'Pflegegrad 3-5    \\[ \\]|Pflegestufe 2-3 \\[ \\]|(beantrag|will|soll|keine|Gatt|Ehe(mann|frau)|pflege(grad|stufe) -)'" & vbCrLf & _
            "WHERE NOT ISNULL(pfl.inhalt) AND ISNULL(pfld.icd)" & vbCrLf & _
            "GROUP BY pat_id;"
'"LEFT JOIN eintraege pfl ON v.pat_id = pfl.pat_id AND pfl.inhalt AND pfl.inhalt NOT LIKE '%pflegestufe nein%' AND pfl.inhalt NOT LIKE '%Pflegestufe -%' AND pfl.inhalt NOT LIKE '%auf -%'"
 mins(AWlf) = 10
 maxs(AWlf) = 1500
 AWlf = AWlf + 1
 
' 24
' kann man mit nur 'schwindel' und .. AND schw.art NOT IN ('andm','andm2') auch zum Finden falsch kategorisierter Diabetesanamnesen verwenden
 AwN(AWlf) = "Möglicherweise nachzutragender Altersschwindel R42 (vorher 63)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patAlter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'And ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'And ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'And ersteller='ah'))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",DATE(schw.zeitpunkt) Zp, schw.inhalt Schw " & vbCrLf & _
            ",GROUP_CONCAT(DISTINCT mp.medikament) Schwindelmedikament " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0123]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview schwd ON v.pat_id = schwd.pat_id AND schwd.gicd LIKE 'R42%' " & vbCrLf & _
            "LEFT JOIN eintraege schw ON v.pat_id = schw.pat_id AND schw.inhalt RLIKE /*'schwindel'*/ '(?<!nicht |kein |keinen |ohne )schwindel(?! beim Aufstehen[?] nein| beim Aufstehen[?]  S)' /*AND schw.art NOT IN ('andm','andm2')*/  " & vbCrLf & _
            "LEFT JOIN medplan mp ON v.pat_id = mp.pat_id AND (medikament LIKE '%dimen%' OR (medikament LIKE '%vert%' AND medikament NOT LIKE '%verteil%' AND medikament NOT LIKE '%vertr%' AND medikament NOT LIKE '%Hevert%') " & vbCrLf & _
                  "OR medikament LIKE '%vasomotal%' OR medikament LIKE '%betahistin%' OR medikament LIKE '%quamen%' OR medikament LIKE '%fluna%' OR medikament LIKE '%natil%' OR medikament LIKE '%sibelium%') " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND ISNULL(schwd.icd) AND ((NOT ISNULL(schw.inhalt) OR NOT ISNULL(mp.medikament))) " & vbCrLf & _
            "AND schw.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 3 MONTH) " & vbCrLf & _
            "GROUP BY pat_id, schw.id;"
 mins(AWlf) = 10
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
' 25
' Achtung: jedes Regexp_replace kommt nochmal bei Where vor
 AwN(AWlf) = "Möglicherweise nachzutragende Stuhlinkontinenz R15 oder Harninkontinenz R32 (vorher 64)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'And ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'And ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'And ersteller='ah'))) ah " & vbCrLf & _
            ",dd.gicd,pfld.gicd" & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",REGEXP_REPLACE(bar.inhalt,'^.*Harninkontinenz:.* [(]([0-9]+)[)].*$','\\1') Hrnpkte,hdg.gicd HDg" & vbCrLf & _
            ",REGEXP_REPLACE(bar.inhalt,'^.*([(]?AP[)]?-Versorgung|stuhlinkontinent)[^(]* [(]([0-9]+)[)].*$','\\2') Stpkte,sdg.gicd SDg" & vbCrLf & _
            ",DATE(pfl.zeitpunkt) Zpt, pfl.inhalt Inkontinenztext " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN `faelle` f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN eintraege bar ON v.pat_id = bar.pat_id AND bar.art='bar' AND bar.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art='bar')" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0123]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview sdg ON v.pat_id = sdg.pat_id AND sdg.gicd = 'R15' " & vbCrLf & _
            "LEFT JOIN diagview hdg ON v.pat_id = hdg.pat_id AND hdg.gicd = 'R32' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gicd IN ('R15','R32') " & vbCrLf & _
            "LEFT JOIN eintraege pfl ON v.pat_id = pfl.pat_id AND pfl.inhalt RLIKE 'inkont[^r]' AND NOT pfl.inhalt RLIKE 'keine *inkont[^r]' AND NOT (pfl.inhalt RLIKE 'inkontinent *(5|0)' AND NOT pfl.inhalt RLIKE 'inkontinent *(5|0) *(5|0)')" & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND NOT ISNULL(pfl.inhalt)" & vbCrLf & _
            "AND((pfl.inhalt NOT LIKE'Barthel-Index%'AND pfld.gicd IS NULL)OR(pfl.inhalt LIKE'Barthel-Index%'AND((sdg.gicd IS NULL AND REGEXP_REPLACE(bar.inhalt,'^.*([(]?AP[)]?-Versorgung|stuhlinkontinent)[^(]* [(]([0-9]+)[)].*$','\\2')<10)OR(hdg.gicd IS NULL AND REGEXP_REPLACE(bar.inhalt,'^.*Harninkontinenz:.* [(]([0-9]+)[)].*$','\\1')<10))))" & vbCrLf & _
            "AND pfl.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) " & vbCrLf & _
            "GROUP BY pat_id, pfl.id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 26
 AwN(AWlf) = "Möglicherweise nachzutragender Demenz F01-F03 oder Verwirrtheit R41.0 (vorher 65)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='tk'OR(art='tb'AND ersteller='tk')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk')))) tk " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='gs'OR(art='tb'AND ersteller='gs')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs')))) gs " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='ah'OR(art='tb'AND ersteller='ah')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah')))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",DATE(schw.zeitpunkt) Zpt, schw.inhalt DemenzInhalt " & vbCrLf & _
            ",GROUP_CONCAT(DISTINCT mp.medikament) Medikament " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagnosen dd ON v.pat_id = dd.pat_id AND dd.icd RLIKE '^F0[0-3]|^G20' AND dd.diagsicherheit NOT IN ('A','Z','V') " & vbCrLf & _
            "LEFT JOIN diagview schwd ON v.pat_id = schwd.pat_id AND schwd.gicd RLIKE '^F0[0123]|R41\.[038]' " & vbCrLf & _
            "LEFT JOIN eintraege schw ON v.pat_id = schw.pat_id AND (schw.inhalt RLIKE '[^öÖ]demen[^t]' OR schw.inhalt LIKE '%vergessl%' OR schw.inhalt LIKE '%vergeßl%' OR (schw.inhalt LIKE '%dächtn%' AND NOT schw.inhalt LIKE '%aus dem Gedächtnis%' AND NOT schw.inhalt LIKE '%Gedächtnisprotokoll%' AND NOT schw.inhalt LIKE '%nach Gedächtnis %' AND NOT schw.inhalt LIKE '%chmerzgedächtnis%') OR schw.inhalt LIKE '%verwirr%') " & vbCrLf & _
            "LEFT JOIN medplan mp ON v.pat_id = mp.pat_id AND medikament RLIKE 'exiba|ergobel|exelon|galanta|galnora|mema[cn]|nicer[ig]|nimvast|prometax|reminyl|rivastig|sermion|yasnal|Donepe|Donez|Aricept|Axura' " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND ISNULL(schwd.icd) AND (NOT ISNULL(schw.inhalt) OR NOT ISNULL(mp.medikament)) " & vbCrLf & _
            "AND schw.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 3 MONTH) " & vbCrLf & _
            "GROUP BY pat_id, schw.id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 27
 AwN(AWlf) = "Möglicherweise nachzutragendes chronisches Schmerzsyndrom R52.2 oder F45.51 (psych) (vorher 66)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='tk'OR(art='tb'AND ersteller='tk')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk')))) tk " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='gs'OR(art='tb'AND ersteller='gs')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs')))) gs " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='ah'OR(art='tb'AND ersteller='ah')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah')))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",pfl.art, DATE(pfl.zeitpunkt) Zpt, pfl.inhalt Schmerztext " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0123]|^G20'" & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gicd IN ('R52.2','F45.51') " & vbCrLf & _
            "LEFT JOIN eintraege pfl ON v.pat_id = pfl.pat_id AND pfl.art <> 'ufrag' AND pfl.inhalt RLIKE '(?<!keine |kein )schmerz(?!frei|los)' AND NOT pfl.inhalt RLIKE 'bekommen Sie regelm|wie oft am Tag m' AND pfl.art NOT IN ('andm','andm2') AND pfl.art<>'htxt' AND pfl.zeitpunkt> SUBDATE(" & qtAnf(FristS) & ",INTERVAL 6 MONTH) " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND NOT ISNULL(pfl.inhalt) AND ISNULL(pfld.icd) " & vbCrLf & _
            "AND pfl.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 3 MONTH) " & vbCrLf & _
            "GROUP BY pat_id, pfl.id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 28
 AwN(AWlf) = "Möglicherweise nachzutragende Fallneigung R29.6 (vorher 67)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patAlter(v.pat_id) PAlter " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='tk'OR(art='tb'AND ersteller='tk')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk')))) tk " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='gs'OR(art='tb'AND ersteller='gs')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs')))) gs " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='ah'OR(art='tb'AND ersteller='ah')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah')))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",pfl.art, DATE(pfl.zeitpunkt) Zpt, pfl.inhalt Sturztext " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|^G20'" & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gicd IN ('R29.6') " & vbCrLf & _
            "LEFT JOIN eintraege pfl ON v.pat_id = pfl.pat_id AND (pfl.inhalt RLIKE '[^h]fallneig' OR (pfl.inhalt LIKE '%sturz%' AND NOT pfl.inhalt LIKE '%hörsturz%' AND NOT pfl.inhalt LIKE '%radlsturz%' AND NOT pfl.inhalt LIKE '%radsturz%')) AND pfl.art NOT IN ('andm','andm2') " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND NOT ISNULL(pfl.inhalt) AND ISNULL(pfld.icd) " & vbCrLf & _
            "AND pfl.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 3 MONTH) " & vbCrLf & _
            "GROUP BY pat_id, pfl.id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 29
 AwN(AWlf) = "Möglicherweise nachzutragende Bettlägrigkeit R26.3 (vorher 68)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'And ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'And ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'And ersteller='ah'))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",DATE(pfl.zeitpunkt) Zpt, pfl.inhalt Bettlägrigkeitstext " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gicd IN ('R26.3') " & vbCrLf & _
            "LEFT JOIN eintraege pfl ON v.pat_id = pfl.pat_id AND pfl.inhalt LIKE '%bettläg%' AND NOT (pfl.inhalt RLIKE 'Bettlägerig *0' AND NOT pfl.inhalt RLIKE 'Bettlägerig *0 *0')" & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND NOT ISNULL(pfl.inhalt) AND ISNULL(pfld.icd)" & vbCrLf & _
            "AND pfl.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 3 MONTH) " & vbCrLf & _
            "GROUP BY pat_id, pfl.id;"
 mins(AWlf) = 10
 maxs(AWlf) = 600
 AWlf = AWlf + 1

 ' 30
 AwN(AWlf) = "Möglicherweise nachzutragende verminderte körperliche Aktivität R68.8 (vorher 69)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'And ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'And ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'And ersteller='ah'))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",DATE(pfl.zeitpunkt) Zpt, pfl.inhalt Aktivitätstext " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gicd IN ('R68.8') " & vbCrLf & _
            "LEFT JOIN eintraege pfl ON v.pat_id = pfl.pat_id AND pfl.art = 'beweg' AND (pfl.inhalt LIKE '%eingeschr%' OR pfl.inhalt LIKE '%nicht%' OR pfl.inhalt LIKE '%wenig%') AND NOT pfl.inhalt LIKE '%wetter%' " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND NOT ISNULL(pfl.inhalt) AND ISNULL(pfld.icd)" & vbCrLf & _
            "AND pfl.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 3 MONTH) " & vbCrLf & _
            "GROUP BY pat_id, pfl.id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

 ' 31
 AwN(AWlf) = "Möglicherweise nachzutragende multifaktorielle Mobilitätsstörung R26.8 (vorher 70)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='tk'OR(art='tb'AND ersteller='tk')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk')))) tk " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='gs'OR(art='tb'AND ersteller='gs')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs')))) gs " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='ah'OR(art='tb'AND ersteller='ah')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah')))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",DATE(pfl.zeitpunkt) Zpt, pfl.inhalt Mobilitätstext " & vbCrLf & _
            ",DATE(tug.zeitpunkt)," & qtAnf(FristS) & ",DATE(" & qtEnd(FristS) & ") " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gicd IN ('R26.8') " & vbCrLf & _
            "LEFT JOIN eintraege pfl ON v.pat_id = pfl.pat_id AND pfl.art IN('TUG','247') AND (MID(pfl.inhalt,1,INSTR(pfl.inhalt,' ')) > 10) " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND NOT ISNULL(pfl.inhalt) AND ISNULL(pfld.icd)" & vbCrLf & _
            "AND pfl.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) " & vbCrLf & _
            "GROUP BY pat_id ORDER BY pfl.inhalt, v.pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 32
 AwN(AWlf) = "Möglicherweise nachzutragende Dysphagie R13.9 (vorher 71)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='tk'OR(art='tb'AND ersteller='tk')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk')))) tk " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='gs'OR(art='tb'AND ersteller='gs')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs')))) gs " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='ah'OR(art='tb'AND ersteller='ah')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah')))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",DATE(pfl.zeitpunkt) Zpt, pfl.inhalt Dysphagietext " & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gicd IN ('R13.9') " & vbCrLf & _
            "LEFT JOIN eintraege pfl ON v.pat_id = pfl.pat_id AND ((pfl.inhalt LIKE '%schluckst%' AND NOT pfl.inhalt LIKE '%keine schluckst%') OR pfl.inhalt LIKE '%dysphag%') " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND NOT ISNULL(pfl.inhalt) AND ISNULL(pfld.icd)" & vbCrLf & _
            "AND pfl.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) " & vbCrLf & _
            "GROUP BY pat_id, pfl.id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 33
' LPAD(MID(tug.inhalt,1,INSTR(tug.inhalt,' ')),4,' '), 's/ '
' LPAD(MID(adl.inhalt,17),3,' '),' P') Tests " & vbCrLf & _

 AwN(AWlf) = "Patienten mit Barthel- und TUG-Test, die als gesund eingestuft sind (keine icd in ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R41.0','R41.3','R41.8','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.icd RLIKE '^F0[012]') (vorher 72)"
' sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'And ersteller='tk'))) tk" & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'And ersteller='gs'))) gs" & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'And ersteller='ah'))) ah" & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*enötigte Zeit\: *([0-9]+) [sS].*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests" & vbCrLf & _
            "FROM aktfvz v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND (pfld.gicd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.gicd RLIKE '^F0[012]')" & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND ISNULL(pfld.icd) " & vbCrLf & _
            "AND NOT ISNULL(tug.inhalt) AND NOT ISNULL(adl.inhalt) " & vbCrLf & _
            "GROUP BY pat_id;"
            
sql(AWlf) = _
"WITH bar AS" & vbCrLf & _
"(SELECT" & vbCrLf & _
"  CASE WHEN art='bar' THEN SUM(CAST(REGEXP_REPLACE(SUBSTRING_INDEX(inhalt, ')', n),'^.*[(]([0-9]*)$','\\1') AS UNSIGNED)) WHEN art='adl' THEN LPAD(REGEXP_REPLACE(inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')ELSE'Bar'END bis," & vbCrLf & _
"  sub.*" & vbCrLf & _
" FROM (" & vbCrLf & _
"  SELECT e.*, n" & vbCrLf & _
"   FROM (" & vbCrLf & _
"    SELECT ROW_NUMBER() OVER(PARTITION BY pat_id ORDER BY zeitpunkt desc) rn, e.* FROM eintraege e" & vbCrLf & _
"    WHERE e.art IN ('ADL','284','bar')" & vbCrLf & _
"    AND pat_id IN (SELECT pat_id FROM aktfvz)" & vbCrLf & _
"     ORDER BY pat_id, zeitpunkt DESC" & vbCrLf & _
"    )e" & vbCrLf & _
"   JOIN (" & vbCrLf & _
"    WITH RECURSIVE cte AS (SELECT 1 N UNION ALL SELECT N + 1 FROM cte WHERE N < 16)" & vbCrLf & _
"    SELECT n FROM cte" & vbCrLf & _
"   ) nrn" & vbCrLf & _
"   WHERE inhalt IS NOT NULL" & vbCrLf & _
"   AND e.rn=1" & vbCrLf & _
"   GROUP BY id,n" & vbCrLf & _
"  ) sub" & vbCrLf & _
"  GROUP BY id" & vbCrLf & _
"  )," & vbCrLf & _
" tug AS (SELECT * FROM eintraege WHERE art IN ('TUG','247'))" & vbCrLf

sql(AWlf) = sql(AWlf) & _
" SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter" & vbCrLf & _
" ,(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'And ersteller='tk'))) tk" & vbCrLf & _
" ,(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'And ersteller='gs'))) gs" & vbCrLf & _
" ,(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'And ersteller='ah'))) ah" & vbCrLf & _
" ,CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*enötigte Zeit\: *([0-9]+) [sS].*','\\1'),3,' '),'s/ ', adl.bis) Tests" & vbCrLf & _
", DATE_FORMAT(GREATEST(tug.zeitpunkt,adl.zeitpunkt),'%d.%m.%Y') zuletzt" & vbCrLf & _
"FROM aktfvz v" & vbCrLf & _
" JOIN `faelle` f USING (fid)" & vbCrLf & _
" LEFT JOIN namen n ON v.pat_id = n.pat_id" & vbCrLf & _
" LEFT JOIN tug ON v.pat_id = tug.pat_id AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM tug WHERE pat_id = v.pat_id)" & vbCrLf & _
" LEFT JOIN bar adl ON v.pat_id = adl.pat_id AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM bar WHERE pat_id = v.pat_id)" & vbCrLf & _
" LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|^G20'" & vbCrLf & _
" LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND (pfld.gicd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.gicd RLIKE '^F0[012]')" & vbCrLf & _
" WHERE (DATEDIFF(20250701, n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) AND ISNULL(pfld.icd)" & vbCrLf & _
" AND NOT ISNULL(tug.inhalt) AND NOT ISNULL(adl.inhalt)" & vbCrLf & _
" GROUP BY pat_id"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' für Arzthelferinnen:
' 34
 AwN(AWlf) = "Haus- und Praxisärzte"
sql(AWlf) = "ü"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 35
AwN(AWlf) = "Liste der möglicherweise fehlenden Hausärzte (vorher 9)"
' AND COALESCE(d.Dggel,0)=0
sql(AWlf) = "SELECT n.Pat_id, gesnameg(n.pat_id) Name,f.Schgr,f.VKNr,ICD" & vbCrLf & _
            ",CASE WHEN dmpklass = 1 THEN 'nein' WHEN dmpklass = 2 THEN 'HA' WHEN dmpklass = 3 THEN 'hier' WHEN dmpklass = 4 THEN 'ausgeschrieben' ELSE '?' END `DMP` " & vbCrLf & _
            ",(SELECT COUNT(DISTINCT icd) FROM diagnosen WHERE pat_id=n.pat_id) diagnzahl " & vbCrLf & _
            ",(SELECT COUNT(0) FROM eintraege WHERE pat_id=n.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND art IN ('vac','c19i','cia')) impfeintr " & vbCrLf & _
            ",(SELECT COUNT(0) FROM eintraege WHERE pat_id=n.pat_id AND art NOT IN ('cia','c19i','vac','pa')) sonsteintr " & vbCrLf & _
            ",fa.ÜBWVLANR,fa.ÜBWVBSNR,p.FBezeichnung" & vbCrLf & _
            "FROM `namen` n  " & vbCrLf & _
            "LEFT JOIN `aktfvz` f ON f.pat_id = n.pat_id  " & vbCrLf & _
            "JOIN faelle fa USING (fid)" & vbCrLf & _
            "LEFT JOIN epraxis p ON p.FBetriebsnr=fa.ÜBWVBSNR" & vbCrLf & _
            "LEFT JOIN `diagview` d ON d.pat_id = n.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))) " & vbCrLf & _
            "WHERE n.kvnr IN ('') AND NOT ISNULL(f.pat_id) " & vbCrLf & _
            " AND f.SchGr NOT IN (41,44) AND f.VKNr <> 71800 " & vbCrLf & _
             "AND (" & vbCrLf & _
              "NOT EXISTS (SELECT 0 FROM diagview WHERE pat_id=n.pat_id AND gicd = 'Z26.9')" & vbCrLf & _
               "OR (" & vbCrLf & _
                "EXISTS (SELECT 0 FROM diagview WHERE pat_id=n.pat_id AND NOT ((gicd IN ('Z26.9','I10.90')) OR (gicd IN ('U07.1') )) )" & vbCrLf & _
                "AND" & vbCrLf & _
                "EXISTS (SELECT 0 FROM leistungen WHERE pat_id=n.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND leistung NOT RLIKE '^883|^0300|^00000|^99215')" & vbCrLf & _
               ")" & vbCrLf & _
             ")" & vbCrLf & _
            "GROUP BY n.pat_id " & vbCrLf & _
            "HAVING impfeintr=0 OR sonsteintr>2" & vbCrLf & _
            "ORDER BY n.kvnr, n.pat_id DESC"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
 ' 36
 ' diagsicherheit in aktfaellev schon eingebaut
 AwN(AWlf) = "'DMP HA'-Einträge in Notiz bei fehlendem Nachweis der DMP-Teilnahme seitens des Hausarztes im Internet für akt. Diabetestyp (vorher 20)"
 sql(AWlf) = _
 "SELECT n.pat_id, gesname(n.pat_id) Name, ICD, n.getha0 ÜWNNr, CONCAT(IF(h.anrede,'Herr','Frau'), ' ', h.adressat) Hausarzt " & vbCrLf & _
 "FROM `aktfaellev` f " & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `hareal` h ON n.getha0 = h.kvnr " & vbCrLf & _
 "WHERE n.dmpklass = 2 AND f.icd RLIKE '^E1[0-4]\.' " & vbCrLf & _
 "AND NOT ((dmp1<>0 AND icd RLIKE '^E1[0234]') OR (dmp2<>0 AND icd RLIKE '^E1[1234]')) ORDER BY Hausarzt, Name"
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1

 ' 37
 AwN(AWlf) = "Auffällige Arztzuordnung (vorher 51)"
#If True Then
' sql(AWlf) = _
" SELECT pat_id, gesname(pat_id) PName, Azu, Eintr,dt `Dt'farbe`,Termin " & vbCrLf & _
" , CASE WHEN (eintr RLIKE dt AND NOT Azu RLIKE dt) OR (eintr='' AND NOT Azu RLIKE dt) THEN 'Azu falsch?' " & vbCrLf & _
"  WHEN eintr<>'' AND (Azu<>'' AND Azu<>eintr) THEN 'Azu+Desktop falsch?' " & vbCrLf & _
"  WHEN termin<>'' AND ((eintr<>'' AND LEFT(termin,2)<>eintr AND NOT termin LIKE 'ah%') OR (dt<>'' AND termin<>'' AND LEFT(termin,2) NOT RLIKE dt AND NOT dt RLIKE 'ah|wd' AND NOT LEFT(termin,2) RLIKE 'ah|wd')) AND NOT (termin RLIKE 'Impf|Moderna|Biontech|Carotis|Halssch' OR (termin LIKE 'ah%' AND aktq()=_latin1'22022')) THEN ' Termin falsch?' " & vbCrLf & _
"  ELSE 'bitte prüfen' END `mögl.Fehler` " & vbCrLf & _
" ,dt, gsz gszahl,gszul gszuletzt,tkz tkzahl,tkzul tkzuletzt " & vbCrLf & _
" FROM ( " & vbCrLf & _
"  SELECT f.*, CASE WHEN gszul>tkzul AND gsz>tkz THEN 'gs' WHEN gszul<tkzul AND gsz<tkz THEN 'tk' ELSE '' END Eintr " & vbCrLf & _
"  FROM ( " & vbCrLf & _
"   SELECT f.pat_id " & vbCrLf & _
"   , COALESCE((SELECT CASE WHEN lanrid=1 THEN 'gs' WHEN lanrid=2 THEN 'tk' END FROM faelle WHERE pat_id=f.pat_id AND lanrid IN (1,2) AND bhfb <= qend() ORDER BY bhfb DESC LIMIT 1),'') Azu " & vbCrLf & _
"   , (SELECT COUNT(0) FROM eintraege WHERE ((art IN ('gs','doppler','dop','duplex') AND NOT inhalt LIKE '%(tk)%') OR inhalt LIKE '%(gs)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)) gsz " & vbCrLf & _
"   , COALESCE((SELECT DATE(MAX(zeitpunkt)) FROM eintraege WHERE ((art IN ('gs','doppler','dop','duplex') AND NOT inhalt LIKE '%(tk)%') OR inhalt LIKE '%(gs)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)),0) gszul " & vbCrLf & _
"   , (SELECT COUNT(0) FROM eintraege WHERE (art IN ('tk','ARCHIE2','APK') OR inhalt LIKE '%(tk)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)) tkz " & vbCrLf & _
"   , COALESCE((SELECT DATE(MAX(zeitpunkt)) FROM eintraege WHERE (art IN ('tk','ARCHIE2','APK') OR inhalt LIKE '%(tk)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)),0) tkzul " & vbCrLf & _
"   , COALESCE((SELECT CONCAT(CASE WHEN raum='Schade' THEN 'gs' WHEN raum='Kothny' THEN 'tk' WHEN raum='Wagner' THEN 'wd' WHEN raum LIKE 'Hamm%' THEN 'ah' ELSE '' END,' ', DATE_FORMAT(zp,'%e.%c.%y'),' ',REPLACE(REPLACE(LEFT(zusatz,IF(LENGTH(zusatz)-1<25,LENGTH(zusatz)-1,25)),chr(10),''),chr(13),'')) FROM termine t WHERE pid = f.pat_id AND raum<>'Labor' AND zp =(SELECT MIN(zp) FROM termine WHERE zp>=CURRENT_TIMESTAMP() AND pid=t.pid AND raum RLIKE 'Kothny|Schade|Wagner|^Hamm' AND NOT zusatz RLIKE 'doppler|duplex|carotis') AND aktzeit=(SELECT MAX(aktzeit) FROM termine WHERE pid=t.pid AND raum<>'Labor' AND zp=t.zp) LIMIT 1),'') Termin " & vbCrLf & _
"   ,REPLACE(TRIM(CONCAT(IF(obtk,'tk ',''),IF(obgs,'gs ',''),IF(obdw,'wd',IF(obah,'ah','')))),' ','|') dt " & vbCrLf & _
"   FROM aktfv f " & vbCrLf & _
"   LEFT JOIN anaktk ON anaktk.pat_id=f.pat_id " & vbCrLf & _
"  ) f " & vbCrLf & _
" ) i " & vbCrLf & _
" WHERE (eintr<>'' AND dt RLIKE 'tk|gs' AND INSTR(dt,eintr)=0) OR (Azu<>'' AND dt RLIKE 'tk|gs' AND INSTR(dt,Azu)=0) OR (((eintr<>'' AND Termin<>'' AND eintr<>LEFT(Termin,2)) OR (dt<>'' AND Termin<>'' AND LEFT(termin,2)<>dt AND NOT termin LIKE 'ah%' AND NOT dt RLIKE 'ah|wd' AND NOT LEFT(termin,2) RLIKE 'ah|wd')) AND NOT (termin RLIKE 'Impf|Moderna|Biontech|Carotis|Halssch' OR (termin LIKE 'ah%' AND aktq()=_latin1'22022')))" & vbCrLf & _
" ORDER BY `mögl.Fehler`,Termin;"
sql(AWlf) = _
" SELECT RANK() OVER(ORDER BY `mögl.Fehler`,Termin,pat_id) rang, pat_id, gesname(pat_id) PName, Azu, Eintr,dt `Dt'farbe`,Termin " & vbCrLf & _
"  , CASE WHEN (eintr RLIKE dt AND NOT Azu RLIKE dt) OR (eintr='' AND NOT Azu RLIKE dt) THEN 'Azu falsch?' " & vbCrLf & _
"   WHEN eintr<>'' AND (Azu<>'' AND Azu<>eintr) THEN 'Azu+Desktop falsch?' " & vbCrLf & _
"   WHEN termin<>'' AND (e_t OR d_t) AND NOT ungt THEN ' Termin falsch?' " & vbCrLf & _
"   ELSE 'bitte prüfen' END `mögl.Fehler` " & vbCrLf & _
"  ,dt, gsz gszahl,gszul gszuletzt,tkz tkzahl,tkzul tkzuletzt " & vbCrLf & _
"  FROM ( " & vbCrLf & _
"   SELECT i.*" & vbCrLf & _
"   , eintr<>'' AND dt RLIKE 'tk|gs' AND INSTR(dt,eintr)=0 e_d" & vbCrLf & _
"   , Azu<>'' AND dt RLIKE 'tk|gs' AND INSTR(dt,Azu)=0 a_d" & vbCrLf & _
"   , eintr<>'' AND Termin<>'' AND eintr<>LEFT(Termin,2) AND NOT termin LIKE 'ah%' e_t" & vbCrLf & _
"   , dt<>'' AND Termin<>'' AND LEFT(termin,2) NOT RLIKE dt AND NOT termin LIKE 'ah%' AND NOT dt RLIKE 'ah|wd' AND NOT LEFT(termin,2) RLIKE 'ah|wd' d_t" & vbCrLf & _
"   , termin RLIKE 'Impf|Moderna|Biontech|Carotis|Halssch|Belastungs|SD Sono|SD-Sono|Arterien|Venen' OR (termin LIKE 'ah%' AND aktq()=_latin1'22022') ungt" & vbCrLf & _
"   FROM (" & vbCrLf & _
"    SELECT f.*, CASE WHEN gszul>tkzul AND gsz>tkz THEN 'gs' WHEN gszul<tkzul AND gsz<tkz THEN 'tk' ELSE '' END Eintr " & vbCrLf & _
"    FROM ( " & vbCrLf & _
"     SELECT f.pat_id " & vbCrLf & _
"     , COALESCE((SELECT CASE WHEN lanrid=1 THEN 'gs' WHEN lanrid=2 THEN 'tk' WHEN lanrid=5 THEN 'ah' ELSE '?' END FROM `faelle` WHERE pat_id=f.pat_id AND lanrid IN (1,2,5) AND bhfb <= qend() ORDER BY bhfb DESC LIMIT 1),'') Azu " & vbCrLf & _
"     , (SELECT COUNT(0) FROM eintraege WHERE (((art IN ('gs','doppler','dop','duplex','dup') OR (art='tb' AND ersteller='gs')) AND NOT inhalt LIKE '%(tk)%') OR inhalt LIKE '%(gs)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)) gsz " & vbCrLf & _
"     , COALESCE((SELECT DATE(MAX(zeitpunkt)) FROM eintraege WHERE (((art IN ('gs','doppler','dop','duplex','dup') OR (art='tb' AND ersteller='gs')) AND NOT inhalt LIKE '%(tk)%') OR inhalt LIKE '%(gs)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)),0) gszul " & vbCrLf & _
"     , (SELECT COUNT(0) FROM eintraege WHERE ((art IN ('tk','ARCHIE2','APK') OR (art='tb' AND ersteller='tk')) OR inhalt LIKE '%(tk)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)) tkz " & vbCrLf & _
"     , COALESCE((SELECT DATE(MAX(zeitpunkt)) FROM eintraege WHERE ((art IN ('tk','ARCHIE2','APK') OR (art='tb' AND ersteller='tk')) OR inhalt LIKE '%(tk)%') AND pat_id = f.pat_id AND zeitpunkt>SUBDATE(CURRENT_TIMESTAMP(),365)),0) tkzul " & vbCrLf & _
"     , COALESCE((SELECT CONCAT(CASE WHEN raum='Schade' THEN 'gs' WHEN raum='Kothny' THEN 'tk' WHEN raum='Wagner' THEN 'wd' WHEN raum LIKE 'Hamm%' THEN 'ah' ELSE '' END,' ', DATE_FORMAT(zp,'%e.%c.%y'),' ',REPLACE(REPLACE(LEFT(zusatz,IF(LENGTH(zusatz)-1<25,LENGTH(zusatz)-1,25)),chr(10),''),chr(13),'')) FROM termine t WHERE pid = f.pat_id AND raum<>'Labor' AND zp =(SELECT MIN(zp) FROM termine WHERE zp>=CURRENT_TIMESTAMP() AND pid=t.pid AND raum RLIKE 'Kothny|Schade|Wagner|^Hamm' AND NOT zusatz RLIKE 'doppler|duplex|dup|carotis|belastung|SD Sono|SD-Sono|Arterien|Venen') AND aktzeit=(SELECT MAX(aktzeit) FROM termine WHERE pid=t.pid AND raum<>'Labor' AND zp=t.zp) LIMIT 1),'') Termin " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"     ,REPLACE(TRIM(CONCAT(IF(obtk,'tk ',''),IF(obgs,'gs ',''),IF(obwd,'wd',IF(obah,'ah','')))),' ','|') dt " & vbCrLf & _
"     FROM aktfvz f " & vbCrLf & _
"     LEFT JOIN anaktk ON anaktk.pid=f.pat_id " & vbCrLf & _
"    ) f " & vbCrLf & _
"   ) i" & vbCrLf & _
"  ) i " & vbCrLf & _
"  WHERE e_d OR a_d OR ((e_t OR (d_t)) AND NOT ungt)" & vbCrLf & _
"  ORDER BY `mögl.Fehler`,Termin" & vbCrLf & _
";"

#Else
' sql(AWlf) = "SELECT Pat_id,Name, CASE WHEN lanrid=1 THEN 'gs' WHEN lanrid=2 THEN 'tk' WHEN lanrid=5 THEN 'ah' ELSE '?' END LANRid,IF(Eintrag=1,'gs','tk') Eintrag " & vbCrLf & _
             ", (SELECT IF(COUNT(0)=0,'-',CONCAT(CAST(LPAD(COUNT(0),03,' ') AS char), ',zul.',DATE_FORMAT(MAX(zeitpunkt),'%e.%c.%y'))) FROM `eintraege` WHERE (((art IN ('gs','doppler','dop','duplex') OR (art='tb' AND ersteller='gs')) AND NOT inhalt LIKE '%(tk)%') OR inhalt LIKE '%(gs)%') AND pat_id = i.pat_id) Schade " & vbCrLf & _
             ", (SELECT IF(COUNT(0)=0,'-',CONCAT(CAST(LPAD(COUNT(0),03,' ') AS char), ',zul.',DATE_FORMAT(MAX(zeitpunkt),'%e.%c.%y'))) FROM `eintraege` WHERE ((art IN ('tk','ARCHIE2','APK') OR (art='tb' AND ersteller='tk')) OR inhalt LIKE '%(tk)%') AND pat_id = i.pat_id) Kothny " & vbCrLf & _
             ", (SELECT CONCAT(MAX(raum),' ', DATE_FORMAT(MAX(zp),'%e.%c.%y')) FROM termine  t WHERE pid = i.pat_id AND zp =(SELECT MIN(zp) FROM termine t0 WHERE t0.zp>" & qtAnf(FristS) & " AND t0.pid=t.pid AND raum IN ('Kothny','Schade') )) Termin " & vbCrLf & _
             ", Schgr " & vbCrLf & _
             ", dt Desktop " & vbCrLf & _
             "FROM ( " & vbCrLf & _
             "SELECT pat_id,Name, DATE_FORMAT(qanf,'%e.%c.%y') QAnf,DATE_FORMAT(qend,'%e.%c.%y') qend, IF(Schgr=24,'Ü',IF(schgr=0 AND NOT goäkatnr IN ('40','41'),'O',CAST(schgr AS char))) Schgr,LANRid,GROUP_CONCAT(DISTINCT wer) Eintrag, dt FROM ( " & vbCrLf & _
             "SELECT af.fid, f.qanf, f.qend, f.schgr, f.pat_id, f.goäkatnr, gesname(f.pat_id) Name, f.lanrid, IF(art IN (('gs','doppler','dop','duplex') OR (art='tb' AND ersteller='gs')) OR inhalt LIKE '%(gs)%','1',IF(art IN ('tk','ARCHIE2','APK') OR inhalt LIKE '%(tk)%','2','0')) wer " & vbCrLf & _
             ",TRIM(CONCAT(IF(obtk,'tk ',''),IF(obgs,'gs ',''),IF(obdw,'wd|ah ',''))) dt" & vbCrLf & _
             "FROM `aktfvz` af LEFT JOIN `faelle` f ON af.fid = f.fid " & vbCrLf & _
             "LEFT JOIN `eintraege` e ON e.pat_id = f.pat_id AND (((art IN ('gs','doppler','dop','duplex','tk','ARCHIE2','APK') OR (art='tb' AND ersteller IN ('gs','tk','ah'))) OR inhalt LIKE '%(tk)%' OR inhalt LIKE '%(gs)%') AND e.zeitpunkt >= f.qanf AND e.zeitpunkt <= f.qend " & vbCrLf & _
             "LEFT JOIN anakt ON anakt.pat_id=af.pat_id" & vbCrLf & _
             ") i " & vbCrLf & _
             "GROUP BY fid,pat_id) i " & vbCrLf & _
             "WHERE ((lanrid='2' AND eintrag='1') OR (lanrid = '1' AND eintrag='2') OR dt='' OR (lanrid=1 AND INSTR(dt,'gs')=0) OR (lanrid=2 AND INSTR(dt,'tk')=0)) " & vbCrLf & _
             "ORDER BY lanrid DESC, pat_id, qanf, schgr DESC"
#End If
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1


' 38
 AwN(AWlf) = "Texteinträge, Notizfelder, Krankenkassen"
sql(AWlf) = "ü"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
  
' 39
' noch zu tun!!!!!!!
 AwN(AWlf) = "Unleserliche Diabetesdauer (vorher 35)"
 sql(AWlf) = _
 "SELECT Pat_id, Name, i.dseit " & vbCrLf & _
 "FROM ( " & vbCrLf & _
 "SELECT f.pat_id, gesname(f.pat_id) Name, a.`diabetes seit` dseit, vorgestellt, " & vbCrLf & _
 "d.ICD " & vbCrLf & _
 "FROM aktfvz f LEFT JOIN anamnesebogen a ON f.pat_id = a.pat_id " & vbCrLf & _
 "LEFT JOIN `diagview` d " & vbCrLf & _
 " ON d.pat_id = f.pat_id AND d.obdauer <> 0 AND d.gicd REGEXP '^E1[0-4]\.' ) i " & vbCrLf & _
 "WHERE NOT CAST(dseit AS DECIMAL) BETWEEN YEAR(NOW())-90 AND YEAR(NOW()) AND zudatum(i.dseit,i.vorgestellt) =i.vorgestellt AND NOT dseit RLIKE '^[0123456789.]{1,2}\.[0123456789.]{1,2}.[0123456789.]{2,4}$' AND NOT ISNULL(icd) "
 mins(AWlf) = 7
 maxs(AWlf) = 20
 AWlf = AWlf + 1

' 40
' ktag fehlerhaft
AwN(AWlf) = "Sono, Doppler oder Duplex ohne Befund (vorher 7)"
' sql(AWlf) = _
"SELECT innen.*, ezp, LEFT(art,7) AS art, inhalt from" & vbCrLf & _
 "(SELECT f.`pid` AS Pat_ID, name, " & vbCrLf & _
         "MID(name,locate(' ',name,locate('SonoBild ',name)+9),9) AS udatroh, " & vbCrLf & _
         "STR_TO_DATE(MID(name,locate(' ',name,locate('SonoBild ',name)+9),9),'%d.%m.%y') AS udat " & vbCrLf & _
          "FROM " & vbCrLf & _
     "(SELECT pat_id AS pid, nachname, vorname, fid,schgr,ik,vknr FROM `faelle` WHERE schgr <> '90' AND NOT goäkatnr IN ('40','41') AND quartal = '" & AktQ & "') AS f " & vbCrLf & _
       "LEFT JOIN tmbrie b ON pid = b.pat_id AND name LIKE '%sonobild %' " & vbCrLf & _
     ") AS innen " & vbCrLf & _
    "LEFT JOIN " & vbCrLf & _
      "(SELECT pat_id pid, art, inhalt, zeitpunkt ezp FROM `eintraege` " & vbCrLf & _
          "WHERE (art IN ('doppler','dop','duplex','dup','sono')OR(art='tb'AND ersteller='gs')) AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") AS eint " & vbCrLf & _
          "ON innen.pat_id = eint.pid AND DATE(ezp) = innen.udat " & vbCrLf & _
    "WHERE udat BETWEEN " & lQAnfuEnd(FristS) & " AND ISNULL(ezp)" ' , 1)
 sql(AWlf) = _
 "SELECT i.*" & vbCrLf & _
 " FROM (" & vbCrLf & _
 "  SELECT /*COUNT(0)zahl,*/" & vbCrLf & _
 "   b.pat_id,gesnameg(b.Pat_id)PName,CASE WHEN name RLIKE'_Schade_'THEN'Sch'WHEN name RLIKE '_Dr. Kothny_'THEN'Kot'WHEN name RLIKE'_Hammerschmidt_'THEN'Ham'ELSE''END Arzt" & vbCrLf & _
 "   ,udat,uag,name" & vbCrLf & _
 "   ,CASE WHEN uag IN('abd_general','Vasc_superficial')THEN'^(dup|sono)' WHEN uag LIKE'vasc%'THEN'^d[uo]p'ELSE'^sono'END Muster" & vbCrLf & _
 "  FROM " & vbCrLf & _
 "  (SELECT" & vbCrLf & _
 "   STR_TO_DATE(REGEXP_REPLACE(NAME,'^.*_([0-9]{8})_?([0-9]{6})([_a]| [0-9]{1,2}| - Kopie( \\([0-9]{1,3}\\))?)?\.png.*$','\\1 \\2'),'%Y%m%d')udat" & vbCrLf & _
 "   ,REGEXP_REPLACE(NAME,'^.*_(?:(?:(?:Schade_li)?L12-[35]|[CS]5-2)|(?:[FMO]|[MW]?__(?:C5_2|L11_3))?|_)_([a-zA-Z ]*(?:_[a-z]+)?)_.*$','\\1')uag" & vbCrLf & _
 "   ,b.* FROM" & vbCrLf & _
 "   tmbrie b" & vbCrLf & _
 "  )b" & vbCrLf & _
 "  WHERE b.NAME LIKE '%.png'" & vbCrLf & _
 "  AND b.NAME RLIKE '_[0-9]{8}_?[0-9]{6}([_a]| [0-9]{1,2}| - Kopie( \\([0-9]{1,3}\\))?)?\.png'" & vbCrLf & _
 "  GROUP BY b.pat_id,b.udat,b.uag)i" & vbCrLf & _
 "  WHERE (SELECT inhalt FROM eintraege" & vbCrLf & _
 "   WHERE pat_id=i.pat_id " & vbCrLf & _
 "   AND inhalt RLIKE '^(Abdomen|(Hals(schlagadern|venen|arterien|weichteile|sono)|Wei?chteile|Pleura|(Sono )?Schid?l?l?dd?d?r?r?s?ür?s?s?e?|(Bein|Arm)(venen|arterien)( (re|li(nks)?|bds.))?)|Nierenarterien?|Restharn|SD|Darmarterien|Abdomen|Belastung(suntersuchung)?).*:'" & vbCrLf & _
 "   AND art RLIKE muster" & vbCrLf & _
 "   AND DATE(zeitpunkt)=i.udat LIMIT 1" & vbCrLf & _
 "  ) IS NULL" & vbCrLf
 sql(AWlf) = sql(AWlf) & _
 "ORDER BY udat DESC"
 mins(AWlf) = 10
 maxs(AWlf) = 360
 AWlf = AWlf + 1

 ' 41
 AwN(AWlf) = "Unverwertbare DMP-Einteilung im Notiz-Feld (vorher 10)"
 sql(AWlf) = "SELECT n.pat_id, gesname(n.pat_id) Name, n.dmpklass, n.dmpbeg,n.dmpkhkklass, n.DMPKHKBeg, n.DMPCopdKlass, n.DmpCOPDBeg, n.DMPABKlass, n.DMPABBeg, REPLACE(n.notiz,CONCAT(char(13),char(10)),'') Notiz " & vbCrLf & _
             ",(SELECT GROUP_CONCAT(CONCAT(MID(abk,INSTR(abk,'DMP')+3),'_',art,'(',IF(ok,'1','0'),'/',IF(ausgedruckt,'1','0'),'/',IF(exportiert=18991230,'-',DATE_FORMAT(exportiert,'%d.%m.%y')),')') ORDER BY exportiert DESC) FROM dmpreihe d WHERE karteidatum > NOW()-INTERVAL 2 YEAR AND pat_id=n.pat_id) DMPzuletzt" & vbCrLf & _
             "FROM aktfvz f " & vbCrLf & _
             "LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
             "WHERE notiz LIKE '%dmp%' AND NOT notiz = CONCAT('DMP NEIN',char(13),char(10)) AND NOT notiz = CONCAT('DMP AUSGESCHRIEBEN',char(13),char(10)) AND NZNr=0 " & vbCrLf & _
             "AND NOT (notiz LIKE 'DMP KHK%' AND n.dmpkhkklass) " & vbCrLf & _
             "AND NOT (notiz LIKE 'DMP COPD%' AND n.dmpcopdklass) " & vbCrLf & _
             "AND NOT (notiz LIKE 'DMP AB%' AND n.dmpabklass) " & vbCrLf & _
             "AND NOT notiz RLIKE 'DMP[ -](?:halbjährlich|(?:beim )?HA abgeklärt|Regress)'" & vbCrLf & _
             "GROUP BY n.pat_id " & vbCrLf & _
             "ORDER BY pat_id DESC"
 mins(AWlf) = 7
 maxs(AWlf) = 120
 AWlf = AWlf + 1
  
' 42
AwN(AWlf) = "Fehlende oder zu alte DMP-(Teilnahme)bestätigung in der Karteikarte (vorher 93)"
sql(AWlf) = "" & vbCrLf & _
"SELECT v.pat_id, gesnameg(v.pat_id) Name, b.zeitpunkt, b.name " & vbCrLf & _
"FROM aktfvz v " & vbCrLf & _
"LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
"LEFT JOIN tmbrie b ON v.pat_id = b.pat_id " & vbCrLf & _
" AND ((name LIKE '%dmp-%' AND NOT name LIKE '%dmp-daten%') OR name LIKE '%dmp%teilnahme%' OR name LIKE '%dmp%bestätigung%' OR name LIKE '%teilnahme%dmp%' OR name LIKE '%bestätigung%dmp%') " & vbCrLf & _
" AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM tmbrie WHERE pat_id = b.pat_id " & vbCrLf & _
" AND ((name LIKE '%dmp-%' AND NOT name LIKE '%dmp-daten%') OR name LIKE '%dmp%teilnahme%' OR name LIKE '%dmp%bestätigung%' OR name LIKE '%dmp%bestätigung%' OR name LIKE '%teilnahme%dmp%' OR name LIKE '%bestätigung%dmp%')) " & vbCrLf & _
"WHERE dmpklass = 2 AND (ISNULL(b.Zeitpunkt) OR DATEDIFF(" & qtAnf(FristS) & ", b.Zeitpunkt) > 0) " & vbCrLf & _
"GROUP BY v.pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

 ' 43
 ' vorher liste_43() laufen lassen
 AwN(AWlf) = "Fehlende oder mehrdeutige Krankenkassenkategorie (INSERT INTO kassenliste (vknr, ik, NAME, kateg, anzahlik,anzahlktug, gültigvon, gültigbis,  go,kurzname,rname, eingef, geaen) SELECT vknr, ik, NAME, kateg, anzahlik,anzahlktug, gültigvon, gültigbis,  go,kurzname,rname, eingef, geaen FROM kassenliste WHERE id=...; usw.)"
 sql(AWlf) = _
" SELECT Pat_id,PatName,Ik,VKNr,KategZahl,KategN FROM ( " & vbCrLf & _
" SELECT f.Pat_id, gesname(f.pat_id) Patname, Ik,Vknr, Kateg " & vbCrLf & _
"  , (SELECT COUNT(DISTINCT kateg) FROM kassenliste kal WHERE kal.id=f.kid) KategZahl " & vbCrLf & _
"  , (SELECT GROUP_CONCAT(CONCAT(kateg,IF(ISNULL(pid),'',CONCAT('(',pid,')')))) " & vbCrLf & _
" FROM kassenliste kal WHERE kal.id=f.kid) Kategn " & vbCrLf & _
" FROM aktfv f " & vbCrLf & _
") i " & vbCrLf & _
"WHERE i.KategZahl<>1 OR i.kateg='' OR ISNULL(i.kateg);"
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1
  
' 44
'Briefe umfassen alle Dokumente und dann noch die mit pdf und doc usw.
'(SELECT pfad, zeitpunkt, name FROM `tmbrie` WHERE pat_id=" & Pat_ID & " UNION " & vbCrLf & _
'    " SELECT datname, laend, datname FROM `dokprotlist` WHERE pat_id=" & Pat_ID & ") i" & vbCrLf &

AwN(AWlf) = "Nicht nachweisbar eingescannte und importierte Überweisungsscheine (vorher 47)"
 sql(AWlf) = "SELECT  a.pat_id,gesnameg(a.pat_id) Name, ausgst, lvorl" & vbCrLf & _
 "FROM `aktfvs` a" & vbCrLf & _
 "JOIN faelle f USING (fid)" & vbCrLf & _
 "LEFT JOIN (SELECT pat_id, pfad, zeitpunkt, name FROM `tmbrie`) b " & vbCrLf & _
 "  ON b.pat_ID = a.pat_id AND (b.name LIKE '%Üw%" & Left$(AktQ, 1) & "%" & Right$(AktQ, 2) & "%' OR name LIKE '%Uew%" & Left$(AktQ, 1) & "%" & Right$(AktQ, 2) & "%')" & vbCrLf & _
 "LEFT JOIN `namen` n ON a.pat_id = n.pat_id " & vbCrLf & _
 "WHERE ISNULL(name) AND a.schgr = 24 " & vbCrLf & _
 "    AND NOT EXISTS (SELECT pat_id FROM `aktf` WHERE pat_id = a.pat_id AND schgr = 0 AND NOT f.goäkatnr IN ('40','41')) " & vbCrLf & _
 "GROUP BY a.pat_id"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
#If vor31032018 Then
AwN(AWlf) = "Fehlende 99300 für NI oder RI"
'  AND COALESCE(ni.Dggel,0)=0
sql(AWlf) = "" & vbCrLf & _
"SELECT f.pat_id, gesname(f.pat_id),dm.icd Dm,ni.icd NI,ne.icd Nephr, copd.icd COPD, ri.icd RI, k.name Kasse " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN kassenliste k ON k.id=f.kid" & vbCrLf & _
"LEFT JOIN diagview dm ON f.pat_id = dm.pat_id AND dm.gicd RLIKE '^E1[013]\.2[01]|^E1[013]\.7[23]' " & vbCrLf & _
"LEFT JOIN diagview ni ON f.pat_id = ni.pat_id AND ni.gicd RLIKE '^N18\.[12]' " & vbCrLf & _
"LEFT JOIN diagview ne ON f.pat_id = ne.pat_id AND ne.gicd = 'N08.3'" & vbCrLf & _
"LEFT JOIN diagview copd ON f.pat_id = copd.pat_id AND copd.gicd RLIKE '^J44\.[018][01]'" & vbCrLf & _
"LEFT JOIN diagview ri ON f.pat_id = ri.pat_id AND ri.gicd RLIKE '^J96\.1[019]')" & vbCrLf & _
"LEFT JOIN leistungen l ON l.fid = f.fid AND l.leistung = '99300' " & vbCrLf & _
"WHERE ((NOT ISNULL(dm.icd) AND NOT ISNULL(ni.icd)) " & vbCrLf & _
"   OR (NOT ISNULL(copd.icd) AND NOT ISNULL(ri.icd))) " & vbCrLf & _
"   AND ISNULL(l.leistung) AND k.kateg IN ('AOK','EK') " & vbCrLf & _
"GROUP BY f.pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
#End If

' 45
 AwN(AWlf) = "Atypisch klassifizierte Doppler/Duplex-Untersuchungen (Halsschlagadern, Halsvenen, Thoraxvenen, Darmarterien, Beinarterien, Beinvenen, Nierenarterien, Temporalarterien, Abdomen, Armvenen, Armarterien, Belastung) (vorher 87, vorher 113)"
 sql(AWlf) = vbCrLf & _
 "SELECT DISTINCT pat_id, gesname(pat_id) PName, COUNT(0) zahl, MIN(zeitpunkt) Zt1,MAX(zeitpunkt) Zt2, GROUP_CONCAT(LEFT(inhalt,20) SEPARATOR ' | ') Inhalt " & vbCrLf & _
 "FROM eintraege WHERE art IN ('doppler','dop','duplex','dup') " & vbCrLf & _
 "AND inhalt NOT LIKE 'Halsschlagadern%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Halsart%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Halsvenen%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Thoraxvenen%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Darmarterien%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Beinarterien%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Beinvenen%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Nierenarterie%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Temporalarterien%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Abdomen%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Armvenen%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Armarterien%' " & vbCrLf & _
 "AND inhalt NOT LIKE 'Belastung%' " & vbCrLf & _
 "and NOT (inhalt='alle Fußarteriensignale biphasisch' AND zeitpunkt='20070619120300' AND pat_id=1912) " & vbCrLf & _
 "GROUP BY pat_id, zeitpunkt ORDER BY pat_id DESC, zeitpunkt"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
' Abdomen: Bauchaorta:

' 46
AwN(AWlf) = "falscher Eintrag 'vacc' (vorher 92)" ' : [Mögliche DAK/KKH/HEK/TK-Ziffern Niereninsuffizienz (97600,97601,97603)]"
sql(AWlf) = "select pat_id, gesname(pat_id) PName, Zeitpunkt, art, Inhalt FROM eintraege WHERE art='vacc'"
'sql(AWlf) = "" & vbCrLf & _
"SELECT * FROM ( " & vbCrLf & _
"SELECT f.pat_id, gesnameg(f.pat_id), l.leistung Lstg, l.zeitpunkt Lzp, alb.wert `Alb/Cre jüngste links`, " & vbCrLf & _
"lHbA(f.pat_id) HbA1c, lCre(f.pat_id) Crea, rr.rr `RR jüngster links` " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN (SELECT pat_id FROM eintraege WHERE Art = 'dak' AND Inhalt LIKE 'dak: hier%' GROUP BY pat_id) e ON e.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN (SELECT pat_id,leistung,zeitpunkt FROM leistungen li WHERE leistung IN ('97600','97601','97602') AND zeitpunkt >= SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) " & vbCrLf & _
"AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM leistungen WHERE pat_id = li.pat_id AND leistung IN ('97600','97601','97602')) " & vbCrLf & _
"GROUP BY pat_id) l ON l.pat_id = e.pat_id " & vbCrLf & _
"LEFT JOIN ( " & vbCrLf & _
"SELECT pat_id, zp, LEFT(GROUP_CONCAT(wert ORDER BY zp DESC SEPARATOR ', '),22) wert FROM ( " & vbCrLf & _
"SELECT pat_id, DATE(zeitpunkt) zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) Wert " & vbCrLf & _
"FROM `labor1a` " & vbCrLf & _
"WHERE " & vbCrLf & _
"((abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g %') OR (abkü IN ('ALBU','ALBUMU') AND (einheit = 'mg/l' OR einheit = ''))) " & vbCrLf & _
"UNION SELECT pat_id, DATE(Zeitpunkt) zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',Kommentar),wert) Wert FROM `labor2a` u " & vbCrLf & _
"WHERE " & vbCrLf & _
"((abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g %') OR (abkü IN ('ALBU','ALBUMU') AND (einheit = 'mg/l' OR einheit = ''))) " & vbCrLf & _
") alb GROUP BY pat_id DESC) alb ON f.pat_id = alb.pat_id " & vbCrLf & _
"LEFT JOIN (SELECT pat_id, LEFT(GROUP_CONCAT(REPLACE(REPLACE(REPLACE(rr.rr,' mmHg',''),'mmHg',''),'mmhg','') ORDER BY zeitpunkt DESC SEPARATOR ', '),50) rr FROM rr WHERE rr.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) GROUP BY pat_id) rr ON rr.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN (SELECT pat_id, icd FROM diagnosen WHERE diagsicherheit IN ('g',' ') AND (icd LIKE 'E1%.2' OR icd LIKE 'N18%' OR icd LIKE 'N19%' OR icd LIKE 'I12.0%' OR icd LIKE 'I13.1%' OR icd LIKE 'I13.2%' OR icd RLIKE '^Z49\.[012]' OR icd LIKE 'Z99.2%') AND DiagDatum<" & qtAnf(FristS) & " AND COALESCE(Dggel,0)=0) d ON d.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN leistungen lz ON f.pat_id = lz.pat_id AND lz.leistung IN ('97600','97601','97602') AND lz.zeitpunkt>= SUBDATE(" & qtAnf(FristS) & ",INTERVAL 3 MONTH) " & vbCrLf & _
"WHERE NOT ISNULL(e.pat_id) AND ISNULL(d.icd) AND ISNULL(lz.pat_id) AND ISNULL(l.leistung) ORDER BY f.pat_id) i;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 47
AwN(AWlf) = "Tippfehler beim Blutdruck ('7' statt '/', vorher 94)"
sql(AWlf) = "" & vbCrLf & _
"SELECT pat_id, gesnameg(pat_id) Name, zeitpunkt, rr " & vbCrLf & _
"FROM rr " & vbCrLf & _
"WHERE rr REGEXP '[^0-9][0-9]{2,3}7[0-9]{2,3}[^0-9]'" & vbCrLf & _
"ORDER BY zeitpunkt DESC;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 48
AwN(AWlf) = "Mißgestaltetes Fuß-Makro (vorher 95)"
sql(AWlf) = "" & vbCrLf & _
"SELECT u.pat_id, gesnameg(u.pat_id) Name, u.Zeitpunkt, ulcera, Zn_Amput, Fußdeform, Mitarbeiter " & vbCrLf & _
"FROM aktfvz f LEFT JOIN fuss u " & vbCrLf & _
" ON f.pat_id = u.pat_id AND u.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & _
"WHERE (NOT ulcera IN ('nein','obfl','tief') OR (zn_amput=fußdeform AND zn_amput<>'nein')) " & vbCrLf & _
" AND zeitpunkt > 20170717080000 " & vbCrLf & _
"ORDER BY zeitpunkt;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
 ' 49
 AwN(AWlf) = "als schnell wirksame Insuline Gekennzeichnete mit Namen von langsam wirkenden (vorher 110)"
 sql(AWlf) = vbCrLf & _
 "SELECT insart,medikament,langname,pat_id FROM medarten mp WHERE insart=1 AND ( " & vbCrLf & _
"((langname REGEXP 'basal[^r]' OR langname LIKE '%basal') AND NOT langname LIKE '%rapid%') OR " & vbCrLf & _
"langname LIKE '%abasa%' OR " & vbCrLf & _
"langname LIKE '%suliq%' OR " & vbCrLf & _
"langname LIKE '%levi%' OR " & vbCrLf & _
"langname LIKE '%leve%' OR " & vbCrLf & _
"langname LIKE '%phan%' OR " & vbCrLf & _
"langname LIKE '%prot%' OR " & vbCrLf & _
"langname LIKE '%tres%' OR " & vbCrLf & _
"langname LIKE '%semil%' OR " & vbCrLf & _
"langname LIKE '%touj%' OR " & vbCrLf & _
"langname LIKE '%tard%' OR " & vbCrLf & _
"langname LIKE '%nph%' OR " & vbCrLf & _
"langname LIKE '%nhp%' OR " & vbCrLf & _
"langname LIKE '%semgl%' OR " & vbCrLf & _
"langname LIKE '% sem %' OR " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"(langname LIKE '%nacht%' AND NOT langname LIKE '%mg/dl%') OR " & vbCrLf & _
"(Langname RLIKE 'lan[tg][^e]' AND NOT Langname LIKE '%Apidra + Lantus%') OR " & vbCrLf & _
"langname LIKE '%lanuts%' OR langname LIKE '%lanus%' OR langname LIKE '%lanctus%' OR " & vbCrLf & _
"langname LIKE '%xult%' OR " & vbCrLf & _
"langname LIKE '%verz%' OR " & vbCrLf & _
"langname LIKE '%glarg%' OR " & vbCrLf & _
"langname LIKE '%schwenk%' OR " & vbCrLf & _
"langname LIKE '%nadel%' OR " & vbCrLf & _
"langname LIKE '%tuj%' OR " & vbCrLf & _
"langname LIKE '%abasa%' " & vbCrLf & _
");"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

 ' 50
 AwN(AWlf) = "Träger von Vornamen mit verschiedenen Geschlechtern (vorher 109)"
 sql(AWlf) = vbCrLf & _
 "SELECT vorname, geschlecht, pat_id, gesname(pat_id) Gesname FROM namen n " & vbCrLf & _
 "WHERE (SELECT COUNT(DISTINCT geschlecht) FROM namen ref WHERE ref.vorname=n.vorname)<>1 " & vbCrLf & _
 " AND n.vorname NOT IN ('Aliye','Andrea','Areti','Ayhan','Eike','Gabriele','Michele','Milan','Nikola','Nuran','Rong','Wei','Yüksel') " & vbCrLf & _
 "ORDER BY vorname, geschlecht;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
 ' 51
 ' mit PCRE, ab Mariadb 10.0.5
 AwN(AWlf) = "Als Verzögerungsinsulin klassifiziertes Insulin, das nicht als solches bekannt ist (vorher 111)"
 sql(AWlf) = vbCrLf & _
 "SELECT insart,medikament,langname,pat_id FROM medarten mp WHERE insart=2 AND NOT ( " & vbCrLf & _
"(medikament = 'Insulin 10' OR (langname REGEXP 'basal[^r]' OR langname LIKE '%basal') AND NOT langname LIKE '%rapid%') OR " & vbCrLf & _
"langname REGEXP 'absag|abasa|awi[qk]l[iy]|suliq|lant|lev|phan|prot|tres|semgl|semil|touj|touej|tard|nph|nhp| sem |nacht((?!(mg/dl)).)*$|lan[tg][^e]|lanuts|lanus|lanctus|xult|verz|glarg|schwenk|nadel|tuj|abasa'" & vbCrLf & _
");"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
' "langname LIKE '%lant%' OR " & vbCrLf & _
"langname LIKE '%lev%' OR " & vbCrLf & _
"langname LIKE '%phan%' OR " & vbCrLf & _
"langname LIKE '%prot%' OR " & vbCrLf & _
"langname LIKE '%tres%' OR " & vbCrLf & _
"langname LIKE '%semil%' OR " & vbCrLf & _
"langname LIKE '%touj%' OR " & "langname LIKE '%touej%' OR " & vbCrLf & _
"langname LIKE '%tard%' OR " & vbCrLf & _
"langname LIKE '%nph%' OR " & vbCrLf & _
"langname LIKE '%nhp%' OR " & vbCrLf & _
"langname LIKE '% sem %' OR " & vbCrLf & _
"(langname LIKE '%nacht%' AND NOT langname LIKE '%mg/dl%') OR " & vbCrLf & _
"Langname RLIKE 'lan[tg][^e]' OR " & vbCrLf & _
"langname LIKE '%lanuts%' OR langname LIKE '%lanus%' OR langname LIKE '%lanctus%' OR " & vbCrLf & _
"langname LIKE '%xult%' OR " & vbCrLf & _
"langname LIKE '%verz%' OR " & vbCrLf & _
"langname LIKE '%glarg%' OR " & vbCrLf & _
"langname LIKE '%schwenk%' OR " & vbCrLf & _
"langname LIKE '%nadel%' OR " & vbCrLf & _
"langname LIKE '%tuj%' OR " & vbCrLf & _
"langname LIKE '%abasa%' " & vbCrLf & _

' 52
 AwN(AWlf) = "Auffällige Arteinträge (vorher 112)"
 sql(AWlf) = AuffArtSql
' "SELECT art, gesname(pat_id) Name, COUNT(0) Zahl, GROUP_CONCAT(DISTINCT Pat_id) FROM eintraege " & vbCrLf & _
' "WHERE NOT art IN (" & artspezG & "," & artSpezSonst & ") " & vbCrLf & _
' "GROUP BY art ORDER BY COUNT(0) DESC; "
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

' 53
 AwN(AWlf) = "Möglicherweise ausstehende Blutabnahmedokumentationen ('ba' oder 'bla') (vorher 121)"
 sql(AWlf) = vbCrLf & _
"SELECT f.pat_id, gesname(f.pat_id) Name, DATE(kva.zeitpunkt) kva, kva.inhalt " & vbCrLf & _
",COALESCE((SELECT DATE_FORMAT(MAX(zeitpunkt),'%d.%m.%y') FROM eintraege WHERE pat_id=f.pat_id AND art='ba' AND DATE(zeitpunkt) BETWEEN DATE(kva.zeitpunkt - INTERVAL 8 DAY) AND DATE(kva.zeitpunkt)),'-') `Vor-ba`" & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN eintraege kva ON f.pat_id=kva.pat_id AND kva.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND kva.art='kva' AND obbla(kva.inhalt) " & vbCrLf & _
"LEFT JOIN eintraege ba ON f.pat_id=ba.pat_id AND DATE(ba.zeitpunkt) IN (DATE(kva.zeitpunkt), SUBDATE(DATE(kva.zeitpunkt),1)) AND ba.art IN ('ba','bla') " & vbCrLf & _
"WHERE NOT ISNULL(kva.Art) AND ISNULL(ba.Art) " & vbCrLf & _
"GROUP BY f.pat_id,DATE(kva.zeitpunkt);"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
 ' 54
 AwN(AWlf) = "falsch aufgebaute Datensätze aus Kassenliste (vorher 122)"
 sql(AWlf) = vbCrLf & _
 "SELECT k1.vknr,k1.ik,k2.ik,k1.name,k1.kurzname,k1.kateg,k2.name,k2.kurzname,k2.kateg " & vbCrLf & _
 "FROM kassenliste k1 " & vbCrLf & _
 "LEFT JOIN kassenliste k2 USING (ik,vknr) " & vbCrLf & _
 "WHERE k1.vknr<>k2.vknr;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 55
AwN(AWlf) = "Pat. mit 'DMP ausgeschrieben', die in den letzten 5a da waren, unter 90a alt sind und noch keinen Hinweis auf DMP-Wiedereinschreibung auf dem dem Desktop haben (vorher 145)"
sql(AWlf) = "" & _
"SELECT (SELECT MAX(bhfb) FROM faelle WHERE pat_id=n.pat_id) lfall, n.pat_id, gesname(n.pat_id) pname, patalter(n.pat_id) PAlter " & vbCrLf & _
"FROM namen n LEFT JOIN anamnesebogen USING (pat_id) " & vbCrLf & _
"LEFT JOIN desktop dt ON dt.pat_id=n.pat_id AND dt.titel LIKE '%DMP%' " & vbCrLf & _
"WHERE dmpklass=4 AND tkz=0 AND notiz LIKE '%ausge%' AND ISNULL(dt.titel) " & vbCrLf & _
"HAVING lfall > qanf() - INTERVAL 5 YEAR AND palter<90 "
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 56
AwN(AWlf) = "Pat. mit 'DMP ausgeschrieben', die trotzdem " & AktQ & " einen neuen Fall haben (vorher 146)"
sql(AWlf) = "" & _
"SELECT i.pat_id, gesname(i.pat_id) NAME, i.notiz, i.LetzteDoku, GROUP_CONCAT(f.Quartal ORDER BY bhfb DESC) FallDanach FROM (" & vbCrLf & _
"SELECT n.pat_id,notiz,MAX(STR_TO_DATE(feldinh,'%d.%m.%y')) letztedoku, qende(MAX(STR_TO_DATE(feldinh,'%d.%m.%y'))) qen " & vbCrLf & _
"FROM namen n " & vbCrLf & _
"LEFT JOIN formular fo ON n.pat_id=fo.pat_id " & vbCrLf & _
"WHERE notiz RLIKE 'dmp ausge' AND formvorl RLIKE 'DMP' AND feld='Datum' " & vbCrLf & _
"GROUP BY n.pat_id " & vbCrLf & _
") i LEFT JOIN faelle f ON f.pat_id=i.pat_id AND f.bhfb>qen " & vbCrLf & _
"WHERE NOT ISNULL(f.pat_id) " & vbCrLf & _
"GROUP BY i.pat_id " & vbCrLf & _
"HAVING falldanach RLIKE aktq();"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 57
AwN(AWlf) = "Möglicherweise fehlende Libre- bzw. CGM-Icons auf dem Patientendesktop (vorher 147)"
'sql(AWlf) = "" & _
"SELECT n.pat_id PID, gesname(n.pat_id) PName " & vbCrLf & _
",COALESCE((SELECT LEFT(titel,6) FROM desktop WHERE pat_id=n.pat_id AND iconPath LIKE '%CGM.ico%' LIMIT 1),'-') Icon " & vbCrLf & _
", COALESCE((SELECT CONCAT(DATE(zeitpunkt),' ',LEFT(name,40)) FROM tmbrie WHERE pat_id=n.pat_id AND name RLIKE 'libre|dexcom|cgm|clarity|medtronic|care link|eversen' AND NOT name RLIKE 'cgm bmp' AND zeitpunkt> qanf() - INTERVAL 3 month ORDER BY qdm DESC LIMIT 1),'') Einles " & vbCrLf & _
", COALESCE((SELECT CONCAT(DATE(zeitpunkt),' ',LEFT(feldinh,40)) FROM formular WHERE pat_id=n.pat_id AND form_abk='lar' AND feld='medikament' AND feldinh RLIKE 'Guardian|CGM StartSet|Enlite|Dexcom|Dexom|Eversen|CGM|Libre' AND zeitpunkt>qanf()-INTERVAL 3 MONTH ORDER BY zeitpunkt desc LIMIT 1),'') Rezept " & vbCrLf & _
"FROM namen n " & vbCrLf & _
"HAVING icon='-' AND (einles<>'' OR Rezept<>'') " & vbCrLf & _
";"
sql(AWlf) = "" & _
"SELECT n.pat_id PID, gesname(n.pat_id) PName" & vbCrLf & _
", COALESCE((SELECT CONCAT(DATE(zeitpunkt),' ',LEFT(name,80)) FROM tmbrie WHERE pat_id=n.pat_id AND name RLIKE 'Omnipod ?5|Simplera|Guardian|Enlite|Dexc?om|Eversen|CGM|Libre|clarity|glooko|care ?link' AND NOT name RLIKE 'Privat|Leih|Ablehnung|Antrag|Gutachten|Widerspruch|Schreiben MD|barmer|cgm bmp' AND zeitpunkt> qanf() - INTERVAL 3 month ORDER BY qdm DESC LIMIT 1),'') Einles " & vbCrLf & _
", COALESCE((SELECT CONCAT(DATE(zeitpunkt),' ',LEFT(feldinh,80)) FROM formular WHERE pat_id=n.pat_id AND form_abk='lar' AND feld='medikament' AND feldinh RLIKE 'Omnipod ?5|Simplera|Guardian|Enlite|Dexc?om|Eversen|CGM|Libre|clarity|glooko|care ?link' AND zeitpunkt>qanf()-INTERVAL 3 MONTH ORDER BY zeitpunkt DESC LIMIT 1),'') Rezept " & vbCrLf & _
"FROM namen n" & vbCrLf & _
"WHERE cgm=0 AND COALESCE(sdatum>19000101,0)=0 AND n.pat_id<>0" & vbCrLf & _
"HAVING (einles<>'' OR Rezept<>'')" & vbCrLf & _
"ORDER BY n.pat_id"

 mins(AWlf) = 10
 maxs(AWlf) = 100
 AWlf = AWlf + 1

' 58
AwN(AWlf) = "Auffälliges Geschlecht bei schwangerschaftsbez. Leistungen '92277','92281','97313','97323','97277' (vorher 148)"
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id) PName, n.Geschlecht, l.zeitpunkt, l.leistung FROM aktfv f " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.Leistung IN ('92277','92281','97313','97323','97277') AND l.ZeitPunkt BETWEEN qanf() AND qend() " & vbCrLf & _
"LEFT JOIN namen n ON n.pat_id = f.pat_id " & vbCrLf & _
"WHERE NOT ISNULL(l.leistung) AND n.geschlecht <>'w';"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 59
' s. 58 und 59
 AwN(AWlf) = "Falsche Arteinträge (nicht: 'bzvgl' oder 'bz') bei Blutzuckermessungen mit dem Biosen (vorher 155)"
sql(AWlf) = "" & _
"SELECT pat_id,gesname(pat_id) PName, zeitpunkt,art,inhalt" & vbCrLf & _
", (SELECT group_concat(concat(art,': ',inhalt)) FROM eintraege WHERE pat_id=e.pat_id AND DATE(zeitpunkt)=DATE(e.zeitpunkt) AND art IN ('bzvgl','bz','ogtt')) Zweiteintrag" & vbCrLf & _
"FROM BiosenMessung e " & vbCrLf & _
"WHERE art NOT IN ('bzvgl','bz','ogtt','gluv')" & vbCrLf & _
"ORDER BY pat_id, zeitpunkt DESC" & vbCrLf & _
";"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 60, neuView
AwN(AWlf) = "Fehler bei der Fallzuordnung"
sql(AWlf) = "ü"
 AWlf = AWlf + 1

' 61
 AwN(AWlf) = "Fehlende Arzt-Markierungen"
sql(AWlf) = _
"SELECT pat_id,PName" & vbCrLf & _
",LPAD(REGEXP_REPLACE(COALESCE(CONVERT(SUM(obtk),CHAR),''),'^0$',''),7,' ') EintrTKz, COALESCE(IF(SUM(obtk)=0,'',tkzp),'') tkzp" & vbCrLf & _
",LPAD(REGEXP_REPLACE(COALESCE(CONVERT(SUM(obgs),CHAR),''),'^0$',''),7,' ') EintrGSz ,COALESCE(IF(SUM(obgs)=0,'',gszp),'') gszp" & vbCrLf & _
",LPAD(REGEXP_REPLACE(COALESCE(CONVERT(SUM(obah),CHAR),''),'^0$',''),7,' ') EintrAHz ,COALESCE(IF(SUM(obah)=0,'',ahzp),'') ahzp" & vbCrLf & _
"/*,COALESCE((SELECT 'x' FROM desktop WHERE pat_id=n.pat_id AND iconpath LIKE '%4eckblau%' AND showasnote=0 LIMIT 1),'') icon_k" & vbCrLf & _
",COALESCE((SELECT 'x' FROM desktop WHERE pat_id =n.pat_id AND iconpath LIKE '%4eckgelb%' AND showasnote=0 LIMIT 1),'') icon_s" & vbCrLf & _
",COALESCE((SELECT 'x' FROM desktop WHERE pat_id =n.pat_id AND iconpath LIKE '%4eckHellgruen%' AND showasnote=0 LIMIT 1),'') icon_h" & vbCrLf & _
",COALESCE((SELECT 1 FROM desktop WHERE pat_id =n.pat_id AND iconpath LIKE '%4eckmagneta%' AND showasnote=0 LIMIT 1),'') icon_w*/" & vbCrLf & _
",COALESCE((SELECT CONCAT(CASE WHEN raum='Schade' THEN 'gs' WHEN raum='Kothny' THEN 'tk' WHEN raum='Wagner' THEN 'wd' WHEN raum LIKE 'Hamm%' THEN 'ah' ELSE '' END,' ', DATE_FORMAT(zp,'%e.%c.%y'),' ',REPLACE(REPLACE(LEFT(zusatz,IF(LENGTH(zusatz)-1<25,LENGTH(zusatz)-1,25)),chr(10),''),chr(13),'')) FROM termine t WHERE pid = n.pat_id AND raum<>'Labor' AND zp =(SELECT MIN(zp) FROM termine WHERE zp>=CURRENT_TIMESTAMP() AND pid=t.pid AND raum RLIKE 'Kothny|Schade|Wagner|^Hamm' AND NOT zusatz RLIKE 'doppler|duplex|dup|carotis|belastung|SD Sono|SD-Sono|Arterien|Venen') AND aktzeit=(SELECT MAX(aktzeit) FROM termine WHERE pid=t.pid AND raum<>'Labor' AND zp=t.zp) LIMIT 1),'') Termin" & vbCrLf & _
"FROM (" & vbCrLf & _
"SELECT n.pat_id, gesname(n.pat_id) PName,ersteller,obk,obs,obh" & vbCrLf & _
",FIRST_VALUE(e.zeitpunkt) OVER(PARTITION BY e.pat_id ORDER BY obtk DESC,zeitpunkt DESC) tkzp" & vbCrLf & _
",FIRST_VALUE(e.zeitpunkt) OVER(PARTITION BY e.pat_id ORDER BY obgs DESC,zeitpunkt DESC) gszp" & vbCrLf & _
",FIRST_VALUE(e.zeitpunkt) OVER(PARTITION BY e.pat_id ORDER BY obah DESC,zeitpunkt DESC) ahzp" & vbCrLf & _
",RANK() OVER(PARTITION BY pat_id ORDER BY zeitpunkt DESC) rang, e.art, e.zeitpunkt, e.inhalt,obtk,obgs,obah" & vbCrLf & _
"FROM namen n " & vbCrLf & _
"LEFT JOIN (SELECT e.*" & vbCrLf & _
",art ='tk' OR (art='tb' AND ersteller ='tk') OR inhalt LIKE '%(tk)%' obtk" & vbCrLf & _
",art ='gs' OR (art='tb' AND ersteller ='gs') OR inhalt LIKE '%(gs)%' obgs" & vbCrLf & _
",art ='ah' OR (art='tb' AND ersteller ='ah') OR inhalt LIKE '%(ah)%' obah" & vbCrLf & _
"FROM eintraege e" & vbCrLf
sql(AWlf) = sql(AWlf) & _
") e ON e.pat_id=n.Pat_ID AND (art IN ('tk','gs','wd','ah') OR (art='tb' AND ersteller IN ('tk','gs','wd','ah')) OR inhalt RLIKE '\\((gs|tk|wd|ah)\\)')" & vbCrLf & _
"WHERE EXISTS(SELECT 0 FROM faelle WHERE pat_id=n.pat_id AND bhfb>20240101 LIMIT 1)" & vbCrLf & _
"AND obk=0 AND obs=0 AND obh=0" & vbCrLf & _
") n" & vbCrLf & _
"WHERE rang<16" & vbCrLf & _
"GROUP BY pat_id"
' "-- and n.pat_id=47" & vbCrLf & _


'AwN(AWlf) = "Möglicherweise zu Unrecht dem Bayerischen Landesinstitut für Gesundheit zugeordnete Leistungen (vorher 139)"
' sql(AWlf) = "" & _
"SELECT f.Pat_ID, gesname(f.pat_id) PName, f.Schgr, f.BhFB, l.Zeitpunkt,l.leistung,COALESCE(l.LZahl,0) LZahl,l.LfBegr Begründung,l.LOrgan Organ,l.Faktor " & vbCrLf & _
"FROM faelle f " & vbCrLf & _
"LEFT JOIN leistungen l ON l.fid=f.fid AND DATE(l.zeitpunkt) BETWEEN f.bhfb AND f.bhfe1 " & vbCrLf & _
"WHERE f.VKNr = 71800 " & vbCrLf & _
"HAVING NOT ISNULL(leistung) AND ((f.schgr = '90') = (leistung RLIKE '^88122|^8832[02345]|^8833[12578][ABRVWXGHK]|^88333[ABVWGH]|^88334[YI]{0,1}|^88336[ABVWGH]|^8835[01235]|^8831[12]|^88371|^98060')) " & vbCrLf & _
"ORDER BY f.pat_id, zeitpunkt " & vbCrLf & _
";"
' AND f.bhfb BETWEEN qanf() AND qend()
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 62
 AwN(AWlf) = "Leistungsdatum außerhalb Quartal (vorher 37)"
 sql(AWlf) = _
 "SELECT gf.pat_id, gesname(gf.pat_id) Name, l.zeitpunkt, l.leistung,gf.schgr,gf.goäkatnr, gf.quartal,gf.qanf,gf.qend " & vbCrLf & _
 "FROM faelle gf " & vbCrLf & _
 "LEFT JOIN leistungen l USING(fid) " & vbCrLf & _
 "WHERE ((l.zeitpunkt BETWEEN " & qtAnf(FristS) & " AND " & qtEnd(FristS) & " AND (gf.qanf> " & qtEnd(FristS) & " OR qend < " & qtAnf(FristS) & "))" & vbCrLf & _
 "OR NOT l.zeitpunkt BETWEEN " & qtAnf(FristS) & " AND " & qtEnd(FristS) & " AND (gf.qanf<=" & qtEnd(FristS) & " AND qend>" & qtAnf(FristS) & ")) " & vbCrLf & _
 "AND gf.schgr<>90 AND goäkatnr<>'40'"
 mins(AWlf) = 7
 maxs(AWlf) = 120
 AWlf = AWlf + 1
' "(zeitpunkt > DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 14 DAY)),'-',(((month(SUBDATE(NOW(),INTERVAL 14 DAY))-1) div 3) * 3)+1,'-1')) " & vbCrLf & _
 "AND zeitpunkt < ADDDATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 14 DAY)),'-',(((month(SUBDATE(NOW(),INTERVAL 14 DAY))-1) div 3) * 3)+1,'-1'),INTERVAL 3 MONTH)) "

' AwN(AWlf) = "Fehlende Ziffer 99995"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, LEFT(CONCAT(IF(titel='','',CONCAT(titel,' ')),IF(nvorsatz='','',CONCAT(nvorsatz,' ')),nachname,', ',vorname),25) Name, l.leistung AS Leistung, icd FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('99995')) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND obdauer <> 0 AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0 AND icd REGEXP '^E1[0-4]\.|^O24\.' " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (leistung <> '99995' OR ISNULL(leistung)) " & vbCrLf & _
' "AND NOT ISNULL(icd) "
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1

' 63
AwN(AWlf) = "Vermutlich durch fehlerhaften Leistungsimport generierte Doppelfälle (vorher 144)"
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id) PName, l.Leistung, l.Zeitpunkt " & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN leistungen l ON f.pat_id=l.pat_id AND l.zeitpunkt BETWEEN qanf()AND qend()" & vbCrLf & _
" AND NOT EXISTS (SELECT 0 FROM leistungen WHERE pat_id=f.pat_id AND zeitpunkt BETWEEN qanf()AND qend() AND leistung <> l.leistung)" & vbCrLf & _
" AND (SELECT COUNT(0) FROM faelle WHERE pat_id=f.pat_id AND schgr<>90 AND bhfb BETWEEN qanf()AND qend())>1 " & vbCrLf & _
"WHERE NOT ISNULL(l.leistung)" & vbCrLf & _
";"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1


' 64, neuView
AwN(AWlf) = "DMP"
sql(AWlf) = "ü"
 AWlf = AWlf + 1
  
  ' 65
AwN(AWlf) = "Unzufriedenstellende DMP-Klassifikation bei D.m. (vorher 17)"
sql(AWlf) = _
"SELECT Pat_ID,gesnameg(pat_id) Name,`Alter[a]`,Leistung,ICD, CASE WHEN dmpklass = 1 THEN 'nein' WHEN dmpklass = 2 THEN 'HA' WHEN dmpklass = 3 THEN 'hier' WHEN dmpklass = 4 THEN 'ausgeschrieben' ELSE '?' END `DMP`, Notiz,Kateg,SchGr,maxtha(pat_id)`max.Ther`,TherAkt,Ther1 FROM " & vbCrLf & _
"(SELECT * FROM " & vbCrLf & _
"(SELECT f.pat_id AS pat_id, gesnameg(f.pat_id) Name, DATEDIFF(" & qtAnf(FristS) & ", n.gebdat) DIV 365.24 `Alter[a]`, l.leistung AS Leistung, icd, dmpklass, REPLACE(REPLACE(notiz,char(13),''),char(10),'') Notiz, kateg, kl.name klname, schgr, Therakt, Ther1  FROM " & aktf & " " & vbCrLf & _
"LEFT JOIN (SELECT pat_id,leistung,zeitpunkt FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"LEFT JOIN `diagview` d ON (f.pat_id = d.pat_id AND d.obdauer <> 0) AND gicd RLIKE '^E1[0-4]\.' " & vbCrLf & _
"LEFT JOIN `kassenliste` kl ON kl.id=f.kid" & vbCrLf & _
"LEFT JOIN `namen` n ON f.pat_id = n.pat_id" & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id" & vbCrLf & _
"WHERE  n.nachname<>'Bereitschaftsdienst'" & vbCrLf & _
"ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen" & vbCrLf & _
"WHERE dmpklass NOT IN (2,3,4)" & vbCrLf & _
"AND NOT kateg IN ('PBe', 'SHV', '') " & vbCrLf & _
"AND klname NOT RLIKE 'BKK SCHEUFELEN'" & vbCrLf & _
"AND NOT ISNULL(icd)"
maxs(AWlf) = 26
AWlf = AWlf + 1
' 'LKK',
  
  
' 66, neuView
AwN(AWlf) = "Ziffern für Diabetes"
sql(AWlf) = "ü"
 AWlf = AWlf + 1
  
 ' 67
 ' 15.10.14: Komb und CT noch nicht bewertet, da unsicher und z.T. falsch ausgewertet
 ' 12.10.19: maxtha braucht lang, soll nur in Auswahl von namen berechnet werden
 AwN(AWlf) = "Fehlende oder falsche Betreuungspauschale 97320/97310 (ICT oder CSII), 97321 (and.Therapie), 97322/97312 (Kinder) im DMP (vorher 18)"
sql(AWlf) = _
"SELECT * FROM ( " & vbCrLf & _
"SELECT f.pat_id PID, gesnameg(f.pat_id) Name, DTyp, n.dmpklass DMP,f.Kateg " & vbCrLf & _
", f.mTA Maxth " & vbCrLf & _
", IF(dtyp='g','97333',IF(ADDDATE(n.gebdat,INTERVAL 18 YEAR)>" & qtAnf(FristS) & ",IF(dtyp='1','97312','97322'),IF(dtyp='1','97310',IF(f.rang IN (8,9,10),'97320','97321')))) `Soll` " & vbCrLf & _
", COALESCE(l.leistung,'') Ist, COALESCE(SUM(l.lzahl),0) Zahl " & vbCrLf & _
"FROM aktfvmta f " & vbCrLf & _
"LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND leistung IN ('97310','97312','97320','97321','97322','97333') AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"WHERE dtyp IN ('1','2','g') AND f.kateg NOT IN ('SHV') " & vbCrLf & _
"GROUP BY pid " & vbCrLf & _
") i WHERE (soll<>ist OR (soll='97333' AND zahl<>2)) AND NOT (ist='' AND (dmp=1 OR kateg=''))"

'sql(AWlf) = _
" SELECT Pat_id, gesnameg(pat_id), ICD, Maxth, Leistung, Soll FROM " & vbCrLf & _
"(SELECT maxtha(i.pat_id) Maxth, CASE WHEN !oberw THEN '97322' ELSE CASE WHEN rang(maxtha(i.pat_id)) IN (6,7,8) THEN '97320' ELSE '97321' END END `Soll`, " & vbCrLf & _
"        i.* FROM " & vbCrLf & _
"     (SELECT ADDDATE(n.gebdat,INTERVAL 18 YEAR)<" & qtAnf(FristS) & " oberw, l.Leistung, d.icd, f.* FROM aktfvs f " & vbCrLf & _
"       LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
"       LEFT JOIN kassenliste kl ON f.vknr = kl.vknr " & vbCrLf & _
"       LEFT JOIN diagnosen d USING (pat_id) " & vbCrLf & _
"       LEFT JOIN leistungen l ON l.fid=f.fid AND l.leistung IN (" & BetrPausch & ") " & vbCrLf & _
"       WHERE d.obdauer <> 0 AND d.diagsicherheit IN ('G',' ') AND COALESCE(d.Dggel,0)=0 AND d.icd LIKE 'E11%' " & vbCrLf & _
"         AND (dmpklass IN (2,3) OR kl.kateg IN ('LKK','PBe')) " & vbCrLf & _
"       GROUP BY l.id " & vbCrLf & _
"     ) i " & vbCrLf & _
") i " & vbCrLf & _
"WHERE ISNULL(Leistung) OR Leistung <> Soll " & vbCrLf & _
"GROUP BY pat_id;"

' sql(AWlf) = _
 "SELECT * FROM " & vbCrLf & _
 "(SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id pat_id, gesnameg(f.pat_id) Name, DATEDIFF(" & qtAnf(FristS) & ", n.gebdat) div 365.24 `Alter[a]`, maxtha(f.pat_id) `max.Therapie`, l.leistung AS Leistung, icd, dmpklass, REPLACE(REPLACE(notiz,char(13),''),char(10),'') Notiz, kateg, fa.schgr, IF(fa.letzteregel<>'',DATE_FORMAT(ADDDATE(STR_TO_DATE(fa.letzteregel,'%d.%m.%Y'),274),'%e.%c.%y'),'') TdE, n.gebdat " & vbCrLf & _
 "FROM aktfvs f " & vbCrLf & _
 "LEFT JOIN faelle fa USING (fid) " & vbCrLf & _
 "LEFT JOIN `namen` n ON n.pat_id=f.pat_id " & vbCrLf & _
 "LEFT JOIN `leistungen` l USING (fid) " & vbCrLf & _
 "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND d.obdauer <> 0 AND d.diagsicherheit IN ('G',' ') AND COALESCE(d.Dggel,0)=0) AND icd LIKE 'E11%' " & vbCrLf & _
 "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
 "WHERE leistung IN (" & BetrPausch & ") " & vbCrLf & _
 "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
 "WHERE (ISNULL(leistung) OR " & vbCrLf & _
 "  (ADDDATE(gebdat,INTERVAL 18 YEAR) >= " & qtAnf(FristS) & " AND leistung <> '97322') OR " & vbCrLf & _
 "  (ADDDATE(gebdat,INTERVAL 18 YEAR)  < " & qtAnf(FristS) & " AND leistung =  '97322') OR " & vbCrLf & _
 "  (ADDDATE(gebdat,INTERVAL 18 YEAR) < " & qtAnf(FristS) & " AND " & vbCrLf & _
 "  ((rang(`max.Therapie`)in(6,7,8) AND leistung <> '97320') OR (rang(`max.Therapie`)in(0,1,2,3,4,5) AND leistung<>'97321')) ) " & vbCrLf & _
 " ) " & vbCrLf & _
 "AND NOT ISNULL(icd) " & vbCrLf & _
 "AND (dmpklass IN (2,3) OR kateg IN ('LKK','PBe'))"
' "AND notiz LIKE '%DMP%' " & vbCrLf & _
' "AND NOT ISNULL(notiz) " & vbCrLf & _
 mins(AWlf) = 7
' "    (((NOT ((ther1 LIKE '%Diät%' OR ther1 LIKE 'OAD%' OR ther1 LIKE 'Komb%' OR ther1 LIKE 'CT%') AND (therakt LIKE '%Diät%' OR therakt LIKE 'OAD%' OR therakt LIKE 'Komb%' OR therakt LIKE 'CT%')) ) AND leistung <> '97320') OR " & vbCrLf & _
' "     ((    ((ther1 LIKE '%Diät%' OR ther1 LIKE 'OAD%' OR ther1 LIKE 'Komb%' OR ther1 LIKE 'CT%') AND (therakt LIKE '%Diät%' OR therakt LIKE 'OAD%' OR therakt LIKE 'Komb%' OR therakt LIKE 'CT%')) ) AND leistung <> '97321'))) " & vbCrLf & _
' ""
 maxs(AWlf) = 26
 AWlf = AWlf + 1
 
 ' 68
 AwN(AWlf) = "Fehlende oder falsche Betreuungspauschale 97310 (ICT oder CSII), 97312 (Kinder) für D.m.1 mit DMP (vorher 19)"
 ' AND COALESCE(d.Dggel,0)=0
sql(AWlf) = _
" SELECT Pat_id, gesnameg(pat_id),ICD,Leistung,Soll,kateg,kasse FROM " & vbCrLf & _
"(SELECT CASE WHEN !oberw THEN '97312' ELSE '97310' END `Soll`, " & vbCrLf & _
"        i.* FROM " & vbCrLf & _
"     (SELECT ADDDATE(n.gebdat,INTERVAL 18 YEAR)<" & qtAnf(FristS) & " oberw, l.leistung, d.icd, f.* " & vbCrLf & _
"       ,kl.Kateg,CONCAT(kl.name,' (',kl.kurzname,')')Kasse" & vbCrLf & _
"       FROM aktfvs f" & vbCrLf & _
"       LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
"       LEFT JOIN diagview d USING (pat_id)" & vbCrLf & _
"       LEFT JOIN kassenliste kl ON kl.id=f.kid" & vbCrLf & _
"       LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.leistung IN (" & BetrPausch & ") AND l.zeitpunkt BETWEEN qanf() and qend()" & vbCrLf & _
"       WHERE d.obdauer <> 0 AND d.gicd LIKE 'E10%' " & vbCrLf & _
"         AND (((dmpklass=2 OR (dmpklass=3 AND dmpbeg<=qend()))) OR kl.kateg IN ('PBe','')) " & vbCrLf & _
"       GROUP BY l.id " & vbCrLf & _
"     ) i " & vbCrLf & _
") i " & vbCrLf & _
"WHERE ISNULL(Leistung) OR Leistung <> Soll " & vbCrLf & _
"GROUP BY pat_id;"
' 'LKK',
' 'SHV'

' sql(AWlf) = _
"SELECT f.pat_id pat_id, gesnameg(f.pat_id) Name, PatAlter(f.pat_id) `Alter[a]`, l.leistung Leistung, obmednetz, icd, dmpklass, REPLACE(REPLACE(notiz,char(13),''),char(10),'') Notiz, kateg, schgr, maxtha(f.pat_id) `max.Therapie`, Therakt, Ther1 " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN leistungen l ON f.fid=l.fid AND l.leistung IN (" & BetrPausch & ") " & vbCrLf & _
"LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND d.obdauer <> 0 AND d.diagsicherheit IN ('G',' ') AND COALESCE(d.Dggel,0)=0) AND icd LIKE 'E10%' " & vbCrLf & _
"LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
"LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
" WHERE (ISNULL(leistung) OR " & vbCrLf & _
"   (ADDDATE(n.gebdat,INTERVAL 18 YEAR) >= " & qtAnf(FristS) & " AND leistung <> '97312') OR " & vbCrLf & _
"   (ADDDATE(n.gebdat,INTERVAL 18 YEAR)  < " & qtAnf(FristS) & " AND leistung =  '97312') " & vbCrLf & _
"   ) " & vbCrLf & _
" AND NOT ISNULL(icd) " & vbCrLf & _
" AND ((obMedNetz = 0) OR ISNULL(obmednetz)) " & vbCrLf & _
" AND (dmpklass IN (2,3) OR kateg IN ('LKK','PBe')) " & vbCrLf & _
" GROUP BY f.pat_id "
 maxs(AWlf) = 26
 AWlf = AWlf + 1
'
'n.gebdat, ADDDATE(n.gebdat,INTERVAL 18 YEAR) adddat,
'gebdat,
 
' 19.2.11
' AwN(AWlf) = "Fehlende 97350B für D.m.1 bei LKK"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, l.leistung AS Leistung, obMedNetz, icd, REPLACE(notiz,char(13),'') AS notiz, kateg, schgr  FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND obdauer <> 0 AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0) AND icd LIKE 'E10%' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (leistung <> '97350B' OR ISNULL(leistung)) " & vbCrLf & _
' "AND kateg IN ('LKK') " & vbCrLf & _
' "AND NOT ISNULL(icd) " & vbCrLf & _
' "AND obmednetz = 0 " & vbCrLf & _
' "AND true"
' mins(AWlf) = 7
' maxs(AWlf) = 30
' AWlf = AWlf + 1
 

' 19.2.11
' AwN(AWlf) = "Fehlende 97360B für D.m.2 bei LKK"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, l.leistung AS Leistung, obMedNetz, icd, REPLACE(notiz,char(13),'') AS notiz, kateg, schgr  FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND obdauer <> 0 AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0) AND icd LIKE 'E11%' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (leistung <> '97360B' OR ISNULL(leistung)) " & vbCrLf & _
' "AND kateg IN ('LKK') " & vbCrLf & _
' "AND NOT ISNULL(icd) " & vbCrLf & _
' "AND obmednetz = 0 " & vbCrLf & _
' "AND true"
' mins(AWlf) = 7
' maxs(AWlf) = 30
' AWlf = AWlf + 1

' 19.2.11
' AwN(AWlf) = "Fehlende 97370B für D.m.2 ohne DMP mit ICT bei BKK, BKN, EK"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, l.leistung AS Leistung, obMedNetz, icd, REPLACE(notiz,char(13),'') AS notiz, ther1, therakt, kateg, schgr  FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND obdauer <> 0 AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0) AND icd LIKE 'E11%' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (leistung <> '97370B' OR ISNULL(leistung)) " & vbCrLf & _
' "AND kateg IN ('BKK', 'BKN', 'EK') " & vbCrLf & _
' "AND NOT ((ther1 LIKE '%Diät%' OR ther1 LIKE 'OAD%') AND (therakt LIKE '%Diät%' OR therakt LIKE 'OAD%')) " & vbCrLf & _
' "AND NOT ISNULL(icd) " & vbCrLf & _
' "AND obmednetz = 0 " & vbCrLf & _
' "AND notiz REGEXP 'DMP *nein'"
'' "AND notiz LIKE '%DMP%' " & vbCrLf & _
' mins(AWlf) = 7
' maxs(AWlf) = 30
' AWlf = AWlf + 1

 ' 69
 ' diagsicherheit in aktfaellev schon eingebaut
' AwN(AWlf) = "'DMP HA'-Einträge in Notiz bei Patienten ohne Diabetesdiagnose, bereits IN 5) eingebaut"
' sql(AWlf) = "SELECT n.pat_id, LEFT(CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')),IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')),n.nachname,', ',n.vorname),25) Name  FROM `aktfaellev` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN `hareal` h ON n.getha0 = h.kvnr WHERE n.dmpklass = 2 AND ISNULL(f.icd)"
 AwN(AWlf) = "Motivationskandidaten ('92278','92282') (vorher 21)"
 sql(AWlf) = motsql()
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1

' 19.2.11
' AwN(AWlf) = "Fehlende 97371B für Gestationsdiabetes bei BKK, BKN, IKK, LKK"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, l.leistung AS Leistung, obMedNetz, icd, REPLACE(notiz,char(13),'') AS notiz, kateg, schgr  FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0) AND icd LIKE 'O24.4%' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (leistung <> '97371B' OR ISNULL(leistung)) " & vbCrLf & _
' "AND kateg IN ('BKK', 'BKN', 'IKK','LKK') " & vbCrLf & _
' "AND NOT ISNULL(icd) " & vbCrLf & _
' "AND obmednetz = 0 "
' mins(AWlf) = 7
' maxs(AWlf) = 30
' AWlf = AWlf + 1

 ' 70
' QUARTAL(lei.zeitpunkt)=aktq() hier (zumindest bei 0 Einträgen) auch nicht langsamer
 AwN(AWlf) = "Potentiell falsche Pauschalenzahl (" & BetrPausch & ") (vorher 22)"
' sql(AWlf) = _
"SELECT f.pat_id AS PatID, f.nachname AS Nachname, f.vorname AS Vorname, lei.ct Pauschalenzahl, IF(MinL='97333' AND MaxL='97333',2,1) Soll, MinL `Pauschale 1`, IF(MaxL=MinL,'',MaxL) `ggf. Pauschale 2` " & vbCrLf & _
"FROM (SELECT f.pat_id, nachname, vorname, fid,schgr,goäkatnr,ik,vknr FROM `aktfv` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY pat_id, fid DESC, schgr) AS f " & vbCrLf & _
"LEFT JOIN (SELECT pat_id, COALESCE(SUM(lzahl),0) ct, MIN(Leistung) MinL, MAX(Leistung) MaxL, Zeitpunkt FROM `leistungen` l WHERE leistung IN (" & BetrPausch & ") GROUP BY fid) lei " & vbCrLf & _
"ON lei.pat_id = f.pat_id AND lei.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"WHERE (Minl='97333' AND Maxl='97333' AND ct<>2) OR ((Minl<>'97333' OR MaxL<>'97333') AND lei.ct <> 1)"
 sql(AWlf) = _
 "SELECT f.pat_id, gesname(f.pat_id) PName, leistung, SUM(lzahl) LZahl, IF(l.Leistung='97333',2,1) Soll" & vbCrLf & _
 ", (SELECT COUNT(0) FROM faelle WHERE pat_id=f.pat_id AND bhfb BETWEEN qanf() AND qend() AND Schgr<>90) fallzahl" & vbCrLf & _
 "FROM aktfv f" & vbCrLf & _
 "LEFT JOIN leistungen l ON l.pat_id = f.pat_id AND l.leistung IN ('97310','97312','97320','97321','97322','97333')" & vbCrLf & _
 " AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
 "GROUP BY f.pat_id,leistung" & vbCrLf & _
 "HAVING lzahl<>soll*fallzahl" & vbCrLf & _
 ";" ' *fallzahl eingefügt am 23.5.24 aufgrund Patienten vom 4.Quartal 23: 2532, 52386, 52438, 68957 mit Kassenwechsel
 mins(AWlf) = 7
 maxs(AWlf) = 15
 AWlf = AWlf + 1
' QUARTAL(lei.zeitpunkt)=aktq() " &

 ' 71
 ' ktag fehlerhaft
 AwN(AWlf) = "Leistungen 97324 oder 97314 und 02311 am selben Tag (vorher 23)"
 ' AND COALESCE(diab.Dggel,0)=0
 sql(AWlf) = _
 "SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id AS pat_id,gesname(f.pat_id) Name,l.leistung AS Leistung,kateg,ldat Tag,diab.icd,kid AS DTyp " & vbCrLf & _
 "FROM " & aktf & " " & vbCrLf & _
 "LEFT JOIN (SELECT pat_id,leistung,DATE(zeitpunkt) ldat FROM `leistungen` WHERE leistung IN ('97314','97324')) AS l ON f.pat_id = l.pat_id and l.ldat BETWEEN qbegs(f.quartal) AND qends(f.quartal) " & vbCrLf & _
 "LEFT JOIN (SELECT pat_id,leistung,DATE(zeitpunkt) l1dat FROM `leistungen` WHERE leistung IN ('02311')) AS l1 ON f.pat_id = l1.pat_id and l1.l1dat BETWEEN qbegs(f.quartal) AND qends(f.quartal) " & vbCrLf & _
 "LEFT JOIN diagview diab ON f.pat_id = diab.pat_id AND (diab.gicd REGEXP '^E1[0-4]\.|^R73' OR (diab.icd='O24.4' AND diab.Dggel=0 AND diab.diagsicherheit IN ('G',' ') AND diab.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)))" & vbCrLf & _
 "LEFT JOIN `kassenliste` kl ON kl.id=f.kid" & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
 "WHERE (NOT ISNULL(l.leistung) AND NOT ISNULL(l1.leistung)) AND ldat=l1dat " & vbCrLf & _
 "ORDER BY leistung DESC) i GROUP BY pat_id"
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1

 
' 72
 AwN(AWlf) = "Fehlende 97314,97324 für Fußsyndrom nach Einträgen, Fotos und Diagnosen (vorher 25)"
 ' NOT ISNULL(eArt) e1, Fuß_akt<>'' e2, NOT ISNULL(uArt) e3, WFc RLIKE '^0D|^[12][BCD]|^[3-5]' e4, Wchron RLIKE '\\.[2-5]' e5,
 '  AND COALESCE(wc.Dggel,0)=0
 sql(AWlf) = _
 "SELECT i.Pat_id,PName,`E1..`,`G63.2`,Wchron,Wakt,WFc,WFa,Rp_je,Rp_akt,Fuß_akt,uArt FROM (" & vbCrLf & _
 "SELECT f.pat_id, gesname(f.pat_id) PName,IF(d.ICD RLIKE '....[47]','~',d.ICD) `E1..`, IF(ISNULL(n.ICD),'fehlt','~') `G63.2`" & vbCrLf & _
 " ,COALESCE(GROUP_CONCAT(DISTINCT wc.ICD),'-') Wchron, COALESCE(GROUP_CONCAT(DISTINCT wa.ICD),'-') Wakt" & vbCrLf & _
 " ,(SELECT COALESCE(MAX(MID(name,INSTR(name,'WA ')+3,2)),'-') FROM tmbrie WHERE pat_id=f.pat_id AND name RLIKE 'WA [0-5]') WFc" & vbCrLf & _
 " ,(SELECT COALESCE(MAX(MID(name,INSTR(name,'WA ')+3,2)),'-') FROM tmbrie WHERE pat_id=f.pat_id AND name RLIKE 'WA [0-5]' AND zeitpunkt BETWEEN qanf() AND qend()) WFa" & vbCrLf & _
 " ,GROUP_CONCAT(DISTINCT CASE WHEN r.Medikament RLIKE 'podo' THEN 'Pod' WHEN r.Medikament RLIKE 'Orthon' THEN 'Nagel' WHEN r.Medikament RLIKE 'orthop|therapies' THEN 'Schuh' WHEN r.Medikament RLIKE 'bettu' THEN 'Bett' ELSE '' END) Rp_je" & vbCrLf & _
 " ,GROUP_CONCAT(DISTINCT CASE WHEN ra.Medikament RLIKE 'podo' THEN 'Pod' WHEN ra.Medikament RLIKE 'Orthon' THEN 'Nagel' WHEN ra.Medikament RLIKE 'orthop|therapies' THEN 'Schuh' WHEN ra.Medikament RLIKE 'bettu' THEN 'Bett' ELSE '' END) Rp_akt" & vbCrLf & _
 " ,COALESCE(GROUP_CONCAT(DISTINCT CONCAT(fu.art,':',LEFT(fu.Inhalt,4)) SEPARATOR ', '),'') Fuß_akt, e.Art eArt, u.Art uArt, r.Medikament rRz, ra.Medikament raRz, d.ICD dICD, dmpKlass,dmpbeg,Kateg" & vbCrLf & _
 " ,COALESCE(GROUP_CONCAT(DISTINCT CONCAT(e.art,':',e.Inhalt)),'') Debr_akt" & vbCrLf & _
 " FROM aktfv f LEFT JOIN namen USING (pat_id)" & vbCrLf & _
 " LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.zeitpunkt BETWEEN qanf() AND qend() AND (e.inhalt LIKE '%ebrid%' OR e.art LIKE 'debr%' OR (e.inhalt LIKE '%resekt%' AND NOT e.inhalt RLIKE 'Leber.*rese[ck]t|Resektionshöhle|Gebärmutterrese|Segmentresekt|Teilresekt|Linksresekt|Totalresekt|Prostataresekt|SD-resekt|Mucosaresektion|Nachresektion|Resektion der Schild|Strumaresekt|Schilddrüsenresekt|Gallenblase[n]{0,1}resektion|Elektroresektion|wurzelresektion|Nierenresektion|Pan[ck]reas.*resektion|trumektomie|igmaresektion|Resektion Leberzyste|Resektionsbereich|Pan[ck]reaskopfresektion|Re[ck]tumrese[ck]t|darmresekt|olonresekt|olon-resekt|milzresekt|Rese[ck]tion Lunge'))" & vbCrLf & _
 " LEFT JOIN eintraege fu ON fu.pat_id=f.pat_id AND fu.zeitpunkt BETWEEN qanf() AND qend() AND fu.Art RLIKE 'fuss|fuß|usd|ulcus'" & vbCrLf & _
 " LEFT JOIN eintraege u ON u.pat_id=f.pat_id AND u.zeitpunkt BETWEEN qanf() AND qend() AND u.Art = 'ulcus'" & vbCrLf & _
 " LEFT JOIN diagview d ON d.pat_id=f.pat_id AND d.obdauer<>0 AND d.gICD REGEXP '^E1[0-4]\.'" & vbCrLf & _
 " LEFT JOIN diagview n ON n.pat_id=f.pat_id AND n.obdauer<>0 AND n.gICD = 'G63.2'" & vbCrLf & _
 " LEFT JOIN diagview wc ON wc.pat_id=f.pat_id AND wc.obdauer<>0 AND wc.ICD RLIKE 'L89' AND wc.diagsicherheit IN ('G',' ','Z')" & vbCrLf & _
 " LEFT JOIN diagview wa ON wa.pat_id=f.pat_id AND wa.obdauer=0 AND wa.gICD RLIKE 'L89' AND wa.diagdatum BETWEEN qanf() AND qend()" & vbCrLf & _
 " LEFT JOIN rezepteintraege r ON r.pat_id=f.pat_id AND r.rezkllang IN ('Langrezepteintrag','Heilmittelverordnung') AND r.medikament RLIKE 'podo|bettu|therapies|ortho[np]'" & vbCrLf & _
 " LEFT JOIN rezepteintraege ra ON ra.pat_id=f.pat_id AND ra.rezkllang IN ('Langrezepteintrag','Heilmittelverordnung') AND ra.medikament RLIKE 'podo|bettu|therapies|ortho[np]' AND ra.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
 " GROUP BY f.pat_id" & vbCrLf & _
 ") i" & vbCrLf & _
 "LEFT JOIN leistungen l ON l.pat_id = i.pat_id AND leistung IN ('97314','97324') AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
 "WHERE (NOT ISNULL(eArt) OR ((Fuß_akt<>'' OR NOT ISNULL(uArt)) AND (WFc RLIKE '^0D|^[12][BCD]|^[3-5]' OR Wchron RLIKE '\\.[2-5]')) OR (WFa RLIKE '^0D|^[12][BCD]|^[3-5]' OR Wakt RLIKE '\\.[2-5]')) AND NOT ISNULL(dICD) AND ISNULL(leistung) AND (((dmpklass=2 OR (dmpklass=3 AND dmpbeg<=qend()))) OR kateg IN ('PBe'))"
' 'LKK',
'"WHERE (NOT ISNULL(eArt) OR ((Fuß_akt<>'' OR NOT ISNULL(uArt)) AND (WFc RLIKE '^0D|^[12][BCD]|^[3-5]' OR Wchron RLIKE '\\.[2-5]')) OR (WFa RLIKE '^0D|^[12][BCD]|^[3-5]' OR Wakt RLIKE '\\.[2-5]')) OR NOT ISNULL(raRz)) AND NOT ISNULL(dICD) AND ISNULL(leistung) AND (dmpklass IN (2,3) OR kateg IN ('LKK','PBe'))"
'"WHERE (NOT ISNULL(eArt) OR ((Fuß_akt<>'' OR NOT ISNULL(uArt)) AND (WFc RLIKE '^0D|^[12][BCD]|^[3-5]' OR Wchron RLIKE '\\.[2-5]')) OR (WFa RLIKE '^0D|^[12][BCD]|^[3-5]' OR Wakt RLIKE '\\.[2-5]')) OR NOT ISNULL(rRz)) AND NOT ISNULL(dICD) AND ISNULL(leistung) AND (dmpklass IN (2,3) OR kateg IN ('LKK','PBe'))"

' "WHERE (NOT ISNULL(eArt) OR NOT ISNULL(uArt) OR NOT ISNULL(rRz) OR Wchron<>'-' OR WFc<>'-') AND NOT ISNULL(dICD) AND ISNULL(leistung) AND (dmpklass IN (2,3) OR kateg IN ('LKK','PBe'))"
'  "WHERE (NOT ISNULL(fu.art) OR NOT ISNULL(u.art) OR NOT ISNULL(e.art)) AND NOT ISNULL(ra.Medikament) AND NOT ISNULL(d.ICD) AND ISNULL(leistung) AND (dmpklass IN (2,3) OR kateg IN ('LKK','PBe'))" & vbCrLf & _ ' 133 Einträge
'  "WHERE (NOT ISNULL(fu.art) OR NOT ISNULL(u.art) OR NOT ISNULL(e.art)) AND NOT ISNULL(r.Medikament) AND NOT ISNULL(d.ICD) AND ISNULL(leistung) AND (dmpklass IN (2,3) OR kateg IN ('LKK','PBe'))" & vbCrLf & _ ' 319 Einträge
'  "WHERE (NOT ISNULL(e.art) OR NOT ISNULL(u.art) OR NOT ISNULL(r.Medikament)) AND NOT ISNULL(d.ICD) AND ISNULL(leistung) AND (dmpklass IN (2,3) OR kateg IN ('LKK','PBe'))" & vbCrLf & _ ' ca. 648 Einträge
 ' AND NOT ISNULL(d.icd)
 mins(AWlf) = 7
 maxs(AWlf) = 90
 AWlf = AWlf + 1
' "AND zeitpunkt > DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 14 DAY)),'-',(((month(SUBDATE(NOW(),INTERVAL 14 DAY))-1) div 3) * 3)+1,'-1')) " & vbCrLf & _
 "AND zeitpunkt < ADDDATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 14 DAY)),'-',(((month(SUBDATE(NOW(),INTERVAL 14 DAY))-1) div 3) * 3)+1,'-1'),INTERVAL 3 MONTH) " & vbCrLf & _

' 73
 AwN(AWlf) = "Fehlende 02312, 02311 bzw. 02300 für Fußsyndrom nach Einträgen (vorher 27)"
 sql(AWlf) = _
 "SELECT f.pat_id, dmtyp(f.pat_id) dt, gesname(f.pat_id) PName, DATE_FORMAT(e.zeitpunkt,'%e.%c.%y') Zp, GROUP_CONCAT(ndl.leistung) ndl, GROUP_CONCAT(dil.leistung) dil, e.art, e.inhalt " & vbCrLf & _
 "FROM aktfvs f " & vbCrLf & _
 "LEFT JOIN namen USING (pat_id) " & vbCrLf & _
 "LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND (inhalt LIKE '%ebrid%' OR art LIKE 'debr%' OR (inhalt LIKE '%resekt%' AND NOT inhalt RLIKE 'Leber.*rese[ck]t|Resektionshöhle|Resektion der Schild|Resektionsrand Pankreas|Pan[ck]rea(titi|)s.*resekt|Resektion Leberzyste|Resektionsbereich|Re[ck]tumrese[ck]t|Rese[ck]tion Lunge|(darm|magen|milz|acg-|ICC-|schilddrüsen|struma|sd-|nieren|teil|links|total|gebärmutter|mu[ck]osa|gallenblasen|elektro|wurzel|truma|igma|pan[ck]reaskopf|olon|olon-|olypen|operative )resekt|Hypoph.*resekt')) " & vbCrLf & _
 "LEFT JOIN leistungen ndl ON ndl.pat_id = f.pat_id AND ndl.leistung IN ('02300') AND DATE(ndl.zeitpunkt)=DATE(e.zeitpunkt)" & vbCrLf & _
 "LEFT JOIN leistungen dil ON dil.pat_id = f.pat_id AND dil.leistung IN ('02311') AND DATE(dil.zeitpunkt)=DATE(e.zeitpunkt)" & vbCrLf & _
 "LEFT JOIN leistungen dix ON dix.pat_id = f.pat_id AND dix.leistung IN ('02312') AND DATE(dix.zeitpunkt)=DATE(e.zeitpunkt)" & vbCrLf & _
 "WHERE NOT ISNULL(e.art) AND dix.leistung IS NULL AND ((dmtyp(f.pat_id) IN (1,2) AND ISNULL(dil.leistung)) OR (dmtyp(f.pat_id) NOT IN (1,2) AND ISNULL(ndl.leistung)))" & vbCrLf & _
 "GROUP BY f.pat_id, date(e.zeitpunkt) " & vbCrLf & _
 "ORDER BY f.pat_id;" & vbCrLf & _
 ""
' sql(AWlf) = _
 "SELECT pat_id, name, zeitpunkt, inhalt FROM " & vbCrLf & _
 "(SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, dok.zeitpunkt, dok.inhalt, l.leistung AS Leistung FROM " & aktf & " " & vbCrLf & _
 "LEFT JOIN (SELECT pat_id, zeitpunkt, art, inhalt FROM `eintraege` WHERE (art LIKE 'debr%' OR inhalt LIKE '%ebrid%' OR (inhalt LIKE '%resekt%' AND NOT inhalt RLIKE 'Teilresekt|Totalresekt|Prostataresekt|SD-resekt|Resektion der Schild|Strumaresekt|Schilddrüsenresekt|Gallenblasenresektion|Elektroresektion|wurzelresektion|Pankreas.*resektion|trumektomie|igmaresektion')) " & vbCrLf & _
 "AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & _
 ") AS dok ON f.pat_id = dok.pat_id " & vbCrLf & _
 "LEFT JOIN (SELECT fid,leistung,zeitpunkt FROM `leistungen` WHERE leistung IN ('02311','02312')) AS l ON l.fid = f.fid AND DATE(l.zeitpunkt) = DATE(dok.zeitpunkt)" & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
 "WHERE (ISNULL(leistung)) " & vbCrLf & _
 "AND NOT ISNULL(inhalt)"
 mins(AWlf) = 7
 maxs(AWlf) = 100
 AWlf = AWlf + 1
 
' 74
 AwN(AWlf) = "Fehlende 02311 usw. für Fußsyndrom nach Fotostadium (vorher 28)"
 'sql(AWlf) = _
 "SELECT * FROM " & vbCrLf & _
 "(SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, Fototag, fotostad, `02311/02312`, `02313`, `02300`, `97314/97324`, L3Tag FROM " & aktf & " " & vbCrLf & _
 "LEFT JOIN (SELECT * FROM (SELECT MAX(pat_id) AS PID, MAX(MID(name,INSTR(name,""WA ""),5)) AS fotoStad, Quelldatum Fototag FROM `tmbrie` GROUP BY pat_id, Fototag) AS innen WHERE " & vbCrLf & _
 "fotoStad = 'WA 0D' OR (fotostad LIKE 'WA %' AND NOT fotostad IN ('WA 1A','WA 2A') AND NOT fotostad LIKE 'WA 0%' AND NOT fotostad LIKE 'WA -%' ) " & vbCrLf & _
 ") AS dok ON f.pat_id = dok.pid " & vbCrLf & _
 "        AND Fototag BETWEEN " & lQAnfuEnd(FristS) & vbCrLf & _
 "LEFT JOIN (SELECT pat_id,DATE(Zeitpunkt) LTag, leistung `02311/02312` FROM `leistungen` WHERE leistung IN ('02311','02312')) AS l ON f.pat_id = l.pat_id AND ltag = fototag AND ltag BETWEEN " & lqanfuend(FristS) & " " & vbCrLf & _
 "LEFT JOIN (SELECT pat_id,DATE(zeitpunkt) L1Tag, leistung `02313` FROM `leistungen` WHERE leistung IN ('02313')) AS l1 ON f.pat_id = l1.pat_id AND l1tag = fototag AND l1tag BETWEEN " & lqanfuend(FristS) & " " & vbCrLf & _
 "LEFT JOIN (SELECT pat_id,DATE(zeitpunkt) L2Tag, leistung `02300` FROM `leistungen` WHERE leistung IN ('02300')) AS l2 ON f.pat_id = l2.pat_id AND l2tag = fototag AND l2tag BETWEEN " & lqanfuend(FristS) & " " & vbCrLf & _
 "LEFT JOIN (SELECT pat_id,DATE(zeitpunkt) L3Tag, leistung `97314/97324` FROM `leistungen` WHERE leistung IN ('97314','97324')) AS l3 ON f.pat_id = l3.pat_id AND l3tag BETWEEN " & lqanfuend(FristS) & " " & vbCrLf & _
 "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY `02311/02312` DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
 "WHERE (`02311/02312` <> '02311' OR ISNULL(`02311/02312`)) " & vbCrLf & _
 "AND NOT (`02311/02312` = '02312' AND `02313` = '02313' AND NOT ISNULL(`02311/02312`) AND NOT ISNULL(`02313`)) " & vbCrLf & _
 "AND NOT ISNULL(fotostad) " & vbCrLf & _
 "GROUP BY Pat_id,Fototag,L3Tag"
 
sql(AWlf) = _
" SELECT ia.Pat_id, ia.PName, ia.ICD,ia.FotoTag," & vbCrLf & _
" IF(INSTR(kvinh,'bds')<>0 AND L13z<>2 AND LFLeist<>'02311','zu wenige 02313', " & vbCrLf & _
"   IF(INSTR(wvinh,'bds')<>0 AND LFz<>2,'zu wenige 02311/02312', " & vbCrLf & _
"     IF(ISNULL(LFLeist),'02311/02312 dazu', " & vbCrLf & _
"        IF((LFLeist='02312' OR ISNULL(LFLeist)) AND ISNULL(L13Leist) AND L13z<>0,'02313 dazu', " & vbCrLf & _
"           IF(LFLeist='02311' AND NOT ISNULL(L00tLeist),'2300 stört an dem Tag', " & vbCrLf & _
"            IF(LFLeist='02312' AND NOT ISNULL(L00Leist),'2300 stört in dem Quartal', " & vbCrLf & _
"              IF(ISNULL(icd),'Diabetesdiagnose dazu', " & vbCrLf & _
"                IF(ISNULL(L97Leist),CONCAT(IF(icd LIKE 'E10%','97314','97324'),' dazu'),'falscher Fehler') " & vbCrLf & _
"              ) " & vbCrLf & _
"            ) " & vbCrLf & _
"           ) " & vbCrLf & _
"        ) " & vbCrLf & _
"     ) " & vbCrLf & _
"   ) " & vbCrLf & _
" ) Fehler " & vbCrLf & _
",ia.FotoStad, LFLeist `02311/02312`, LFz lfZahl,WVInh, L13Leist `02313`,L13z L13Zahl,KVInh, " & vbCrLf & _
"L00Leist `02300`, L00tLeist, L00tz,L97Leist `97314/97324`, ia.PName " & vbCrLf & _
" FROM ( " & vbCrLf & _
"  SELECT f.pat_id, d.icd, f.fid, wv.inhalt WVInh, kv.inhalt KVInh, gesname(f.pat_id) PName," & vbCrLf & _
"  lf.lei LFLeist,COALESCE(lf.z,0) LFz,L13.lei L13Leist,COALESCE(L13.z,0) L13z,L00.leistung L00Leist,L00t.leistung L00tLeist, COALESCE(L00t.z,0) L00tz, L97.leistung L97Leist," & vbCrLf & _
"  MAX(MID(b.name,b.Fpos+4,2)) FotoStad,  b.QD FotoTag,  b.name name " & vbCrLf & _
"  FROM aktfvs f " & vbCrLf & _
"  LEFT JOIN (SELECT pat_id,DATE(qdm) QD,INSTR(name,' WA ') Fpos,Name FROM tmbrie) b ON b.pat_id = f.pat_id AND b.QD BETWEEN qanf() AND qend() AND b.Fpos<>0 AND MID(b.name,b.Fpos+4,1)<>'-' AND MID(b.name,b.Fpos+4,2) NOT IN ('0A','0B','0C','1A','2A') " & vbCrLf
' AND COALESCE(d.Dggel,0)=0
sql(AWlf) = sql(AWlf) & _
"  LEFT JOIN diagview d ON d.pat_id=f.pat_id AND d.gicd RLIKE '^E1[0-4]' " & vbCrLf & _
"  LEFT JOIN eintraege wv ON wv.pat_id=f.pat_id AND wv.art='wv' AND DATE(wv.zeitpunkt)=QD " & vbCrLf & _
"  LEFT JOIN eintraege kv ON kv.pat_id=f.pat_id AND kv.art='kv' AND DATE(kv.zeitpunkt)=QD " & vbCrLf & _
"  LEFT JOIN (SELECT leistung lei,pat_id,fid,DATE(zeitpunkt) zp,SUM(lzahl) z FROM leistungen WHERE leistung IN ('02311','02312') GROUP BY pat_id,leistung,DATE(zeitpunkt)) lf ON lf.pat_id=f.pat_id AND lf.zp=b.qd " & vbCrLf & _
"  LEFT JOIN (SELECT leistung lei,pat_id,fid,DATE(zeitpunkt) zp,SUM(lzahl) z FROM leistungen WHERE leistung IN ('02313') GROUP BY pat_id,leistung,DATE(zeitpunkt)) L13 ON L13.pat_id=f.pat_id AND L13.zp=b.qd " & vbCrLf & _
"  LEFT JOIN leistungen L00 ON L00.leistung IN ('02300') AND L00.pat_id=f.pat_id AND L00.zeitpunkt BETWEEN qanf() AND qend() " & vbCrLf & _
"  LEFT JOIN (SELECT leistung,pat_id,DATE(zeitpunkt) zp,SUM(lzahl) z FROM leistungen WHERE leistung='02300' GROUP BY pat_id,leistung,DATE(zeitpunkt)) L00t ON L00t.pat_id=f.pat_id AND L00t.zp=b.qd " & vbCrLf & _
"  LEFT JOIN leistungen L97 ON L97.leistung IN ('97314','97324') AND L97.pat_id=f.pat_id AND L97.zeitpunkt BETWEEN qanf() AND qend() " & vbCrLf & _
"  GROUP BY f.pat_id,QD " & vbCrLf & _
") ia " & vbCrLf & _
"WHERE " & vbCrLf & _
" (fotoStad<>'' AND " & vbCrLf & _
"   (   (ISNULL(LFLeist) AND (ISNULL(L00tz) OR L00tz<3)) " & vbCrLf & _
"    OR ((LFLeist='02312' OR ISNULL(LFLeist)) AND ISNULL(L13Leist) AND L13z<>0) " & vbCrLf & _
"    OR (LFLeist='02311' AND COALESCE(L13Leist,0)<>0) " & vbCrLf & _
"    OR (LFLeist='02311' AND COALESCE(L00tLeist,0)<>0) " & vbCrLf & _
"    OR (LFLeist='02312' AND COALESCE(L00Leist,0)<>0) " & vbCrLf & _
"     OR (NOT ISNULL(icd) AND ISNULL(L97Leist)) " & vbCrLf & _
"   ) " & vbCrLf & _
" ) " & vbCrLf & _
" OR INSTR(wvinh,'bds')<>0 AND LFz<>2 " & vbCrLf & _
" OR INSTR(kvinh,'bds')<>0 AND L13z<>2 AND LFLeist<>'02311' " & vbCrLf & _
"GROUP BY ia.fid,FotoTag; "

 
 mins(AWlf) = 7
 maxs(AWlf) = 120
 AWlf = AWlf + 1

' 75
' ktag fehlerhaft
 AwN(AWlf) = "Fehlende 02310/ 02311/ 02312 /02300 nach Wundverband (wv) (vorher 29)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung IN ('02310','02311','02312','02300')),0) `wv(023..)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('wv')" & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `wv(023..)`=0;"

'sql(AWlf) = _
"SELECT f.`pat_id` Pat_ID, ezp Zeitpunkt,  LEFT(art,7) art, Inhalt FROM " & aktf & " " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.`pat_id`= a.`pat_id` " & vbCrLf & _
"LEFT JOIN (SELECT * FROM " & vbCrLf & _
"(SELECT pat_id pid,  fid, LEFT(art,7) art, inhalt, zeitpunkt ezp, DATE(zeitpunkt) zp FROM `eintraege` e " & vbCrLf & _
"WHERE art IN ('wv') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") eint " & vbCrLf & _
"LEFT JOIN (SELECT pat_id wvpat_id, fid wvfid, DATE(zeitpunkt) wvtag, leistung wv FROM `leistungen` WHERE leistung IN ('02311','02312','02300')) wv " & vbCrLf & _
"ON eint.pid = wv.wvpat_id AND eint.zp = wv.wvtag) " & vbCrLf & _
"as ges " & vbCrLf & _
"ON (ges.PID = f.`Pat_ID`) " & vbCrLf & _
"WHERE NOT ISNULL(ezp) AND ISNULL(wvtag) " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1

 ' 76
 ' ktag fehlerhaft
 AwN(AWlf) = "Fehlender Kompressionsverband: (02313) (vorher 30)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung IN ('02313')),0) `kv(02313)`" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung IN ('02310','02311','02350','07340','10330','18340','30501')),0) `debr(..)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('kv')" & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `kv(02313)`=0 AND `debr(..)`=0;"

'sql(AWlf) = _
"SELECT f.`pat_id` Pat_ID, ezp Zeitpunkt,  LEFT(art,7) art, Inhalt FROM " & aktf & " " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.`pat_id`= a.`pat_id` " & vbCrLf & _
"LEFT JOIN (SELECT * FROM " & vbCrLf & _
"((SELECT pat_id pid,  fid, LEFT(art,7) art, inhalt, zeitpunkt ezp, DATE(zeitpunkt) edt FROM `eintraege` e " & vbCrLf & _
"WHERE art IN ('kv') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") eint " & vbCrLf & _
"LEFT JOIN (SELECT pat_id kvpat_id, fid kvfid, DATE(zeitpunkt) kvtag, leistung kv FROM `leistungen` WHERE leistung IN ('02313')) kv " & vbCrLf & _
"ON eint.pid = kv.kvpat_id AND edt = kv.kvtag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id debrpat_id, fid debrfid, DATE(zeitpunkt) debrtag, leistung debr FROM `leistungen` WHERE leistung IN ('02310','02311','02350','07340','10330','18340','30501')) debr " & vbCrLf & _
"ON eint.pid = debr.debrpat_id AND debrtag BETWEEN " & lQAnfuEnd(FristS) & ") " & vbCrLf & _
"as ges " & vbCrLf & _
"ON (ges.PID = f.`Pat_ID`) " & vbCrLf & _
"WHERE NOT ISNULL(ezp) AND ISNULL(kvtag) AND ISNULL(debrtag)" & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1


 ' 78
 ' ktag fehlerhaft
 AwN(AWlf) = "Mehrfache 02300 am selben Tag ohne T01.1-T01.9-Diagnose (vorher 31)"
 sql(AWlf) = "SELECT f.Pat_id, gesname(f.Pat_id) PName, DATE(l.zeitpunkt) Datum, Leistung, COALESCE(SUM(lzahl)) Zahl, d.icd " & vbCrLf & _
             "FROM `aktfvs` f " & vbCrLf & _
             "LEFT JOIN `leistungen` l ON f.pat_id = l.pat_id AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND leistung = '02300' " & vbCrLf & _
             "LEFT JOIN `diagnosen` d ON d.pat_id = f.pat_id AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND d.icd LIKE 'T01.%' " & vbCrLf & _
             "WHERE NOT ISNULL(leistung) GROUP BY f.pat_id, datum " & vbCrLf & _
             "HAVING (zahl > 1 AND ISNULL(icd)) OR zahl>5"
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1

#If ebmalt Then
' ktag fehlerhaft
' AND COALESCE(dg.Dggel,0)=0
 AwN(AWlf) = "Fehlende Chronikerpauschale 03212 (vorher 33)"
 sql(AWlf) = _
 "SELECT fpat_id pat_id, ffid fid, name, Kontakte, " & vbCrLf & _
 "ICD, ICDg, Kdet, Schgr, MinSchgr, LAlter `Alter`, Grundl, Chron, `Zahl(Grundleist)`, `Zahl(Chron'pausch.)` FROM " & vbCrLf & _
 "(SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id Fpat_id, f.fid FFid, LEFT(CONCAT(n.Nachname, ' ', n.Vorname),25) Name, timestampdiff(YEAR, n.gebdat," & qtAnf(FristS) & ") lalter, f.Schgr, MIN(MSchGr.Schgr) MinSchGr, gl.leistung AS Grundl, l.leistung AS chron, COUNT(DISTINCT pat_gl.pat_id) `Zahl(Grundleist)`, COUNT(DISTINCT pat_l.pat_id) `Zahl(Chron'pausch.)`, d.icd ICD, dg.icd ICDg, " & vbCrLf & _
 "(SELECT COUNT(DISTINCT DATE(zp)) FROM (SELECT DATE(zeitpunkt) zp,pat_id,art FROM `eintraege` e UNION SELECT DATE(zeitpunkt) zp,pat_id,'notiz' art FROM rezepteintraege rz) erz WHERE erz.pat_id = fpat_id AND art IN (" & artspezG & ")  AND art NOT IN ('andm','andm2','anal') AND erz.zp BETWEEN " & lQAnfuEnd(FristS, 1) & " ) Kontakte, " & vbCrLf & _
 "(SELECT CONCAT(GROUP_CONCAT(DISTINCT DATE_FORMAT(zp,'%e.%c.%y') SEPARATOR '  '),' / ',GROUP_CONCAT(DISTINCT art)) FROM (SELECT DATE(zeitpunkt) zp,pat_id,art FROM `eintraege` e UNION SELECT DATE(zeitpunkt) zp,pat_id,'notiz' art FROM rezepteintraege rz) erz WHERE erz.pat_id = fpat_id AND art IN (" & artspezG & ") AND art NOT IN ('andm','andm2','anal') AND erz.zp BETWEEN " & lQAnfuEnd(FristS, 1) & " ) Kdet " & vbCrLf & _
 " FROM " & aktf & " " & vbCrLf & _
 "LEFT JOIN `leistungen` gl ON f.fid = gl.fid AND gl.leistung LIKE '030%' " & vbCrLf & _
 "LEFT JOIN `leistungen` l ON f.fid = l.fid AND l.leistung IN ('0322%') " & vbCrLf & _
 "LEFT JOIN `leistungen` pat_gl ON pat_gl.leistung LIKE '030%' AND pat_gl.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND f.pat_id = pat_gl.pat_id " & vbCrLf & _
 "LEFT JOIN `leistungen` pat_l ON pat_l.leistung LIKE '0322%' AND pat_l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND pat_l.pat_id = f.pat_id " & vbCrLf & _
 "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
 "LEFT JOIN `faelle` MSchGr ON f.pat_id = MSchGr.pat_id AND MSchGr.quartal = '" & AktQ & "' " & vbCrLf & _
 "LEFT JOIN `diagview` d ON d.pat_id = f.pat_id AND (d.obdauer <> 0 OR d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & ") AND d.icd REGEXP '^E1[0-4]\.' AND d.diagsicherheit <> 'A'" & vbCrLf & _
 "LEFT JOIN `diagview` dg ON dg.pat_id = f.pat_id AND (dg.icd='O24.4' AND dg.Dggel=0 AND dg.diagsicherheit IN ('G',' ') AND dg.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))" & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
 "GROUP BY f.fid ORDER BY gl.leistung DESC) AS innen WHERE kontakte > 1) AS innen " & vbCrLf & _
 "WHERE " & vbCrLf & _
 "(ISNULL(`Zahl(Chron'pausch.)`)) OR `Zahl(Chron'pausch.)` <> 1 " & vbCrLf & _
 "ORDER BY pat_id "

 mins(AWlf) = 7
 maxs(AWlf) = 20
 AWlf = AWlf + 1
#Else
'•
' 33
' verlagert in AbrFausg(
' myEFrag "SET @qn:=(SELECT quartal FROM aktfvs f LEFT JOIN faelle fa USING (fid) LIMIT 1)"
' myEFrag "SET @qe:=vorquart(@qn,1)"
' myEFrag "SET @qz:=vorquart(@qn,2)"
' myEFrag "SET @qd:=vorquart(@qn,3)"
' myEFrag "SET @qv:=vorquart(@qn,4)"

' 78
AwN(AWlf) = "Mögliche Fehler bei Chronikerpauschale 03220, 03221 (lauto) (vorher 33)"
#If Not False Then
'  AND d.diagsicherheit NOT IN ('A','V')
' AND COALESCE(gd.Dggel,0)=0
 sql(AWlf) = _
"SELECT i.Pat_id PID, gesname(i.pat_id) Name, CONCAT(F0,IF(F0='' OR F1='','',', '),F1) LEIFehler, IF(i.obChrPa,'ja','nein') chron, IF(i.obvorkoz,'ja','nein') 3VorKte, i.aktkoz aktKtzl, i.l0z 03220zl, i.l1z 03221zl, sonst Sonstige_Diagnosen, LEIDAT, LANRID,czp Kontakte,cart KontaktArt" & vbCrLf & _
"FROM ( " & vbCrLf & _
"SELECT i.Pat_id, i.obChrPa, i.obvorkoz, i.aktkoz, i.l0z, i.l1z, i.sonst " & vbCrLf & _
", CASE " & vbCrLf & _
"    WHEN obChrPa=1 AND obvorkoz=1 THEN " & vbCrLf & _
"     CASE " & vbCrLf & _
"      WHEN aktkoz=0 THEN " & vbCrLf & _
"       CASE WHEN l0z=0 THEN '' ELSE '03220 falsch' END " & vbCrLf & _
"      ELSE " & vbCrLf & _
"       CASE WHEN l0z=0 THEN '03220 dazu' WHEN l0z>1 THEN CONCAT('03220 zu häufig (',l0z,'x)') ELSE '' END " & vbCrLf & _
"     END " & vbCrLf & _
"    ELSE " & vbCrLf & _
"     CASE WHEN l0z=0 THEN '' ELSE '03220 falsch' END " & vbCrLf & _
"  END F0, " & vbCrLf & _
"  CASE " & vbCrLf & _
"    WHEN obChrPa=1 AND obvorkoz=1 THEN " & vbCrLf & _
"     CASE " & vbCrLf & _
"      WHEN aktkoz<2 THEN " & vbCrLf & _
"       CASE WHEN l1z=0 THEN '' ELSE '03221 falsch' END " & vbCrLf & _
"      ELSE " & vbCrLf & _
"       CASE WHEN l1z=0 THEN '03221 dazu' WHEN l1z>1 THEN CONCAT('03221 zu häufig (',l1z,'x)') ELSE '' END " & vbCrLf & _
"     END " & vbCrLf & _
"    ELSE "
sql(AWlf) = sql(AWlf) & vbCrLf & _
"     CASE WHEN l1z=0 THEN '' ELSE '03221 falsch' END " & vbCrLf & _
"  END F1 " & vbCrLf & _
"  , LANRID, LEIDAT,czp,cart " & vbCrLf & _
"FROM ( " & vbCrLf & _
" SELECT f.pat_id,czp,f.art cart " & vbCrLf & _
", NOT ISNULL(d.icd) AND ISNULL(pau.leistung) obChrPa " & vbCrLf & _
", IF((SELECT IF(COUNT(0)>0,1,0) FROM faelle WHERE pat_id=f.pat_id AND quartal=@qe) " & vbCrLf & _
" +(SELECT IF(COUNT(0)>0,1,0) FROM faelle WHERE pat_id=f.pat_id AND quartal=@qz) " & vbCrLf & _
" +(SELECT IF(COUNT(0)>0,1,0) FROM faelle WHERE pat_id=f.pat_id AND quartal=@qd) " & vbCrLf & _
" +(SELECT IF(COUNT(0)>0,1,0) FROM faelle WHERE pat_id=f.pat_id AND quartal=@qv) " & vbCrLf & _
" >=3,1,0) obvorkoz " & vbCrLf & _
", IF(ISNULL(f.koz),0,f.koz) aktkoz " & vbCrLf & _
", (SELECT COUNT(leistung) FROM leistungen WHERE pat_id=f.pat_id AND leistung IN ('03220','03220H') AND zeitpunkt BETWEEN qanf() AND qend()) l0z " & vbCrLf & _
", (SELECT COUNT(leistung) FROM leistungen WHERE pat_id=f.pat_id AND leistung IN ('03221','03221H') AND zeitpunkt BETWEEN qanf() AND qend()) l1z " & vbCrLf & _
", GROUP_CONCAT(DISTINCT CONCAT(sd.icd,sd.diagsicherheit,sd.Dggel,sd.obdauer,' ',sd.diagtext)) sonst " & vbCrLf & _
", f.LANRID, DATE(LETZT) LEIDAT " & vbCrLf & _
"FROM aktfkvs f " & vbCrLf & _
"LEFT JOIN diagview d ON d.pat_id=f.pat_id AND d.gicd RLIKE '^E78\.0|^E03.9|^I87.0|^F17.1|^E27\.1|^D35\.2|^L20|^E66\.9|^T78\.[134]|^K74|^I[1234567]\.|^I0[56789]\.|^E1[0-4]\.|^E[234589]\.|^E0[356]\.|^M|^N80|^K5[01]|^C|^J4|^E78\.0|^D50|^D68|^R52.2|^I89.0|^D6[34]|^F[^1]\.|^G' AND NOT d.gicd RLIKE '^M6[5-8]\.' AND (d.obdauer <> 0 OR d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & ") " & vbCrLf & _
"LEFT JOIN diagview gd ON gd.pat_id=f.pat_id AND (gd.icd='O24.4' AND gd.Dggel=0 AND gd.diagsicherheit IN ('G',' ') AND gd.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))" & vbCrLf & _
"LEFT JOIN diagview sd ON sd.pat_id=f.pat_id AND sd.diagsicherheit NOT IN ('A','Z') " & vbCrLf & _
"LEFT JOIN leistungen pau ON pau.pat_id=f.pat_id AND pau.leistung='97146' AND pau.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"GROUP BY f.pat_id) i " & vbCrLf & _
") i " & vbCrLf & _
"WHERE F0<>'' OR F1<>'' ;"
#Else

' ktag fehlerhaft
' CONCAT((MID(f.quartal,1,1)-1) MOD 4+1,MID(f.quartal,2)+ (MID(f.quartal,1,1)-8) div 4)
' AND COALESCE(d.Dggel,0)=0
' sql(AWlf) = _
'"SELECT PID,Name,vorkz,koz,ICDchr,GD,GL,CP,DMseit,Fehler,mHbA1c,mGluc,czp,konze,@qe,qbegs(@qe),qends(@qe) FROM ( " & vbCrLf & _
'"SELECT f.pat_id PID, gesname(f.pat_id) Name, COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id) vorkz, " & vbCrLf & _
'"IF(ISNULL(k.koz),0,k.koz) koz, d.icd ICDchr, dg.icd GD, pat_gl.lst gl, pat_zl.lst zl, GROUP_CONCAT(pat_l.lst)  CP, zuDatum(an.`diabetes seit`,an.vorgestellt) DMseit, " & vbCrLf & _
'"IF(((ISNULL(d.icd) AND (NOT ISNULL(dg.icd) OR ((ISNULL(mH.Wert) OR mH.Wert<6.5) AND (ISNULL(mG.Wert) OR mG.Wert<200)))) AND NOT ISNULL(pat_l.lst)),'ChrP aber evtl. kein D.m.'," & vbCrLf & _
'"IF(((ISNULL(d.icd) AND (ISNULL(dg.icd) AND ((NOT ISNULL(mH.Wert) AND mH.Wert>=6.5) OR (NOT ISNULL(mG.Wert) AND mG.Wert>=200)))) AND NOT ISNULL(pat_l.lst)),'ChrP aber ICD fehlt'," & vbCrLf & _
'"IF(((ISNULL(d.icd) AND (ISNULL(dg.icd) AND ((NOT ISNULL(mH.Wert) AND mH.Wert>=6.5) OR (NOT ISNULL(mG.Wert) AND mG.Wert>=200)))) AND ISNULL(pat_l.lst)),'ChrP und ICD könnten fehlen'," & vbCrLf & _
'"IF(NOT ISNULL(d.icd) AND ISNULL(pat_l.lst) AND COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id) > 2 AND k.koz > 1,'03220 und 03221 fehlen'," & vbCrLf & _
'"IF(NOT ISNULL(d.icd) AND ISNULL(pat_l.lst) AND COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id) > 2 AND k.koz > 0,'03220 dazu'," & vbCrLf & _
'"IF(NOT ISNULL(d.icd) AND ISNULL(pat_l.lst) AND COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id) <=2 AND k.koz > 1 AND zuDatum(an.`diabetes seit`,an.vorgestellt) BETWEEN " & Khtsfl & ",'03221H+03220H scheinen zu fehlen'," & vbCrLf & _
'"IF(NOT ISNULL(d.icd) AND ISNULL(pat_l.lst) AND COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id) <=2 AND k.koz > 0 AND zuDatum(an.`diabetes seit`,an.vorgestellt) BETWEEN " & Khtsfl & ",'03220H scheint zu fehlen'," & vbCrLf & _
'"IF(NOT ISNULL(d.icd) AND NOT ISNULL(pat_l.lst) AND NOT pat_l.lst LIKE '%H' AND COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id) <=2 AND k.koz > 0 AND zuDatum(an.`diabetes seit`,an.vorgestellt) BETWEEN " & Khtsfl & ", CONCAT(IF(k.koz>1,'03221H','03220H'),' statt ',pat_l.lst)," & vbCrLf & _
'"IF(NOT ISNULL(d.icd) AND NOT ISNULL(pat_l.lst) AND pat_l.lst LIKE '%H' AND COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id) > 2 AND k.koz > 0,CONCAT(IF(k.koz>1,'03221','03220'),' statt ',pat_l.lst)," & vbCrLf & _
'"IF(NOT ISNULL(d.icd) AND NOT ISNULL(pat_l.lst) AND pat_l.lst IN ('03221H','03221') AND k.koz<=1,CONCAT('03220',IF(COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id)<= 2,'H',''),' evtl.statt ',pat_l.lst)," & vbCrLf & _
'"IF(NOT ISNULL(d.icd) AND NOT ISNULL(pat_l.lst) AND pat_l.lst IN ('03220H','03220') AND k.koz> 1,CONCAT('03221',IF(COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id)<= 2,'H',''),' dazu')," & vbCrLf & _
'"IF(COUNT(pat_l.lst)>2,CONCAT(IF(ISNULL(d.icd) AND ((ISNULL(mH.Wert) OR mH.Wert<6.5) OR (ISNULL(mG.Wert) OR mG.Wert<200)),'-',CONCAT('nur ',IF(k.koz>1,'03221','03220'),IF(COUNT(fme.pat_id)+COUNT(fmz.pat_id)+COUNT(fmd.pat_id)+COUNT(fmv.pat_id)<= 2,'H',''))),' statt ',GROUP_CONCAT(pat_l.lst))," & vbCrLf & _
'"'')))))))))))) Fehler, kvnr, "
'sql(AWlf) = sql(AWlf) & _
'"mH.Wert mHbA1c, mG.Wert mGluc, k.czp,konze.koz konze,@qe,qbegs(@qe),qends(@qe) " & vbCrLf & _
'"FROM aktfvs a JOIN faelle f USING (fid)" & vbCrLf & _
'"LEFT JOIN namen n ON f.pat_id = n.pat_id" & vbCrLf & _
'"LEFT JOIN `anamnesebogen` an ON f.pat_id = an.pat_id " & vbCrLf & _
'"LEFT JOIN (SELECT pat_id, quartal FROM `faelle` f GROUP BY f.pat_id,quartal) fme ON f.pat_id = fme.pat_id AND fme.quartal = @qe " & vbCrLf & _
'"LEFT JOIN (SELECT pat_id, quartal FROM `faelle` f GROUP BY f.pat_id,quartal) fmz ON f.pat_id = fmz.pat_id AND fmz.quartal = @qz " & vbCrLf & _
'"LEFT JOIN (SELECT pat_id, quartal FROM `faelle` f GROUP BY f.pat_id,quartal) fmd ON f.pat_id = fmd.pat_id AND fmd.quartal = @qd " & vbCrLf & _
'"LEFT JOIN (SELECT pat_id, quartal FROM `faelle` f GROUP BY f.pat_id,quartal) fmv ON f.pat_id = fmv.pat_id AND fmv.quartal = @qv " & vbCrLf & _
'"LEFT JOIN (SELECT pat_id, COUNT(DISTINCT DATE(zeitpunkt)) koz FROM eintraege e WHERE art IN (" & artspezG & ") AND art NOT IN ('andm','andm2','anal') AND zeitpunkt BETWEEN qbegs(@qe) AND qends(@qe) GROUP BY pat_id) konze ON f.pat_id=konze.pat_id " & vbCrLf & _
'"LEFT JOIN (SELECT pat_id, COUNT(DISTINCT DATE(zp)) koz, GROUP_CONCAT(DATE_FORMAT(zp,'%e.%c.') SEPARATOR '•') czp FROM ( " & vbCrLf & _
'"(SELECT pat_id, DATE(zeitpunkt) zp FROM eintraege WHERE art IN (" & artspezG & ") AND art NOT IN ('andm','andm2','anal') " & vbCrLf & _
'"  AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") " & vbCrLf & _
'"UNION (SELECT pat_id, DATE(zeitpunkt) zp FROM rezepteintraege WHERE zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ")) k GROUP BY pat_id ORDER BY zp) k ON f.pat_id = k.pat_id " & vbCrLf & _
'"LEFT JOIN (SELECT GROUP_CONCAT(leistung) lst, fid FROM `leistungen` WHERE leistung LIKE '030%' GROUP BY fid) pat_gl ON f.fid = pat_gl.fid " & vbCrLf & _
'"LEFT JOIN (SELECT GROUP_CONCAT(leistung) lst, fid FROM `leistungen` WHERE leistung LIKE '0322%' GROUP BY fid) pat_l ON f.fid = pat_l.fid " & vbCrLf & _
'"LEFT JOIN (SELECT GROUP_CONCAT(leistung) lst, fid FROM `leistungen` WHERE leistung = '97146' GROUP BY fid) pat_zl ON f.fid = pat_zl.fid " & vbCrLf & _
'"LEFT JOIN `maxHbA1c` mH ON f.pat_id = mH.pat_id " & vbCrLf & _
'"LEFT JOIN `maxGluc` mG ON f.pat_id = mG.pat_id " & vbCrLf & _
'"LEFT JOIN (SELECT MIN(icd) icd,pat_id,obdauer,diagdatum,diagsicherheit FROM `diagnosen` d WHERE (d.icd RLIKE '^L20' OR d.icd RLIKE '^E66\.9' OR d.icd IN ('T78.4', 'T78.3', 'T78.1') OR d.icd LIKE 'K74%' OR d.icd REGEXP '^I[1-7]\.' OR d.icd REGEXP '^I0[5-9]\.' OR d.icd REGEXP '^E1[0-4]\.' OR d.icd RLIKE '^E[234589]\.' OR d.icd RLIKE '^E0[356]\.' OR (d.icd LIKE 'M%' AND NOT d.icd RLIKE '^M6[5678]\.') OR d.icd LIKE 'N80%' OR d.icd LIKE 'K51%' OR d.icd LIKE 'K50%' OR d.icd LIKE 'C%' OR d.icd LIKE 'J4%' OR d.icd = 'E78.0' OR d.icd LIKE 'D50%' OR d.icd LIKE 'D68%' OR d.icd LIKE 'R52.2%' OR d.icd LIKE 'I89.0' OR d.icd LIKE 'D63%' OR icd LIKE 'D64%' OR d.icd REGEXP '^F[^1]\.' OR d.icd LIKE 'G%') AND d.diagsicherheit NOT IN ('A','V') AND (d.obdauer <> 0 OR d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & ") GROUP BY pat_id) d ON d.pat_id = f.pat_id " & vbCrLf & _
'"LEFT JOIN (SELECT MIN(icd) icd,pat_id,obdauer,diagdatum,diagsicherheit FROM `diagview` d WHERE (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))) GROUP BY pat_id) dg ON dg.pat_id = f.pat_id " & vbCrLf & _
'"GROUP BY f.pat_id) i WHERE ( "
' sql(AWlf) = sql(AWlf) & _
'"((ISNULL(ICDchr) AND (NOT ISNULL(GD) OR ((ISNULL(mHbA1c) OR mHbA1c<6.5) AND (ISNULL(mGluc) OR mGluc<200)))) AND NOT ISNULL(cp)) OR " & vbCrLf & _
'"((ISNULL(ICDchr) AND (ISNULL(GD) AND ((NOT ISNULL(mHbA1c) AND mHbA1c>=6.5) OR (NOT ISNULL(mGluc) AND mGluc>=200)))) AND NOT ISNULL(cp))  OR " & vbCrLf & _
'"((ISNULL(ICDchr) AND (ISNULL(GD) AND ((NOT ISNULL(mHbA1c) AND mHbA1c>=6.5) OR (NOT ISNULL(mGluc) AND mGluc>=200)))) AND ISNULL(cp)) OR " & vbCrLf & _
'"(NOT ISNULL(ICDchr) AND ISNULL(CP) AND vorkz > 2 AND koz > 0) OR " & vbCrLf & _
'"(NOT ISNULL(ICDchr) AND ISNULL(CP) AND vorkz <=2 AND koz > 0 AND dmseit BETWEEN " & Khtsfl & " OR " & vbCrLf & _
'"(NOT ISNULL(ICDchr) AND NOT ISNULL(CP) AND NOT CP LIKE '%H' AND vorkz <=2 AND koz > 0 AND dmseit BETWEEN " & Khtsfl & " OR " & vbCrLf & _
'"(NOT ISNULL(ICDchr) AND NOT ISNULL(CP) AND cp LIKE '%H' AND vorkz > 2 AND koz > 0) OR " & vbCrLf & _
'"(NOT ISNULL(ICDchr) AND NOT ISNULL(CP) AND cp IN ('03221H','03221') AND koz <=1) OR " & vbCrLf & _
'"(NOT ISNULL(ICDchr) AND NOT ISNULL(CP) AND cp IN ('03220H','03220') AND koz > 1) OR " & vbCrLf & _
'"INSTR(cp,',') OR " & vbCrLf & _
'"false) AND ISNULL(zl) AND fehler<>''" & _
'" AND (NOT fehler LIKE '%H %' OR kvnr='' OR kvnr RLIKE '^" & KVNr & "' OR kvnr='933284903' OR kvnr='889690003')"
#End If
 mins(AWlf) = 8
 maxs(AWlf) = 20
 AWlf = AWlf + 1

' 79
 AwN(AWlf) = "03221% ohne 03220% (vorher 34)"
sql(AWlf) = "SELECT f.pat_id, gesname(f.pat_id), l1.leistung, l0.leistung " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN leistungen l1 ON f.pat_id=l1.pat_id AND l1.zeitpunkt BETWEEN qanf() AND qend() AND l1.leistung LIKE '%03221%' " & vbCrLf & _
"LEFT JOIN leistungen l0 ON f.pat_id=l0.pat_id AND l1.zeitpunkt BETWEEN qanf() AND qend() AND l0.leistung LIKE '%03220%' " & vbCrLf & _
"WHERE NOT ISNULL(l1.leistung) AND ISNULL(l0.leistung);"
 mins(AWlf) = 7
 maxs(AWlf) = 20
 AWlf = AWlf + 1

#End If

' 80
AwN(AWlf) = "Pauschalenfehler nach Tabelle (vorher 48)"
'17.10.12: Prinzip: für alle Fälle (aktfvs) werden die leistungen rausgesucht, die in der Tabelle `genehmigungen` behandelt werden
'(WHERE l.leistung IN (SELECT leistung FROM `genehmigungen`)).
'Dazu wird jeweils eine Zeile aus `genehmigungen` mit derselben Leistung und dem selben Diabetestyp ausgewählt
'(LEFT JOIN `genehmigungen` g ON l.leistung = g.leistung " & vbCrLf & _ " AND IF(d.icd RLIKE '^E10',_latin1'1', IF(d.icd RLIKE '^E11',_latin1'2',IF(d.icd = 'O24.4',_latin1'g',null))) = g.dtyp " & vbCrLf & _).
'Gibt es keine solche, so entsteht ein fDTyp-Fehler: ... OR ISNULL(myid),'fDTyp!','') fDTyp
'Gibt es eine, so werden alle anderen Bedingungen verglichen und ggf. gemeldet.

'sql(AWlf) = "SELECT a.Pat_id, gesname(a.pat_id) Name, l.id, l.Zeitpunkt, l.Leistung, an.Therakt, g.`therarten`, IF(therarten<>'' AND INSTR(therarten,therakt)=0 AND NOT (d.icd = 'O24.4' AND l.leistung=97271),'FTh!','') fTha, IF(d.icd RLIKE '^E10','1', IF(d.icd RLIKE '^E11','2',IF(d.icd = 'O24.4','g',null))) Dm, GROUP_CONCAT(DTyp) D_Typ, IF(dtyp<>'' AND INSTR(dtyp,IF(d.icd RLIKE '^E11',2,IF(d.icd = 'O24.4','g',null)))=0,'fDTyp!','') DmTyp, IF(k.kateg='EK','vdek',k.kateg) Kasse, Kassen, IF (kassen<>'' AND INSTR(kassen,IF(k.kateg='EK','vdek',k.kateg))=0,'fKas!','') fKas, l.lanr, kothny, schade, IF((l.lanr=889690003 AND schade = 0) OR (l.lanr=933284903 AND kothny=0),'fLANR!','') fLANR, ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24) `Alter`, CAST(CONCAT(IF(von=0,'',von),'-',IF(bis=0,'',bis)) AS char) AlStuf, IF(von<>0 AND (ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24)<von OR ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24)>bis),'fAlter!','') fAlter, " & vbCrLf & _
"IF(n.geschlecht='w',1,0) obweibl, IF(g.weibl=1 AND n.geschlecht<>'w','fWeibl','') fWeibl, IF(mitüw=1 AND a.schgr='00' AND NOT a.goäkatnr IN (40,41) ,'fÜw','') fÜw, IF(n.dmpklass IN (2,3),1,0) dmp, IF((dmp1=1 OR dmp2=1) AND NOT n.dmpklass IN (2,3) AND d.icd = 'O24.4','fDMP','') fDMP " & vbCrLf & _
"FROM `aktfvs` a LEFT JOIN `namen` n ON a.pat_id = n.pat_id LEFT JOIN `faelle` f ON a.fid = f.fid LEFT JOIN `kassenliste` k ON a.vknr = k.vknr AND f.ik = k.ik LEFT JOIN `leistungen` l ON a.fid = l.fid LEFT JOIN `diagnosen` d ON d.pat_id = a.pat_id AND diagsicherheit IN (' ','G','V') AND COALESCE(Dggel,0)=0 AND icd RLIKE '^E1[0-4]|^O24\.4' AND (obdauer<>0 OR d.fid = a.fid OR d.fid=0 OR ISNULL(d.fid)) LEFT JOIN `anamnesebogen` an ON a.pat_id = an.pat_id LEFT JOIN `genehmigungen` g ON l.leistung = g.leistung WHERE l.leistung IN (SELECT leistung FROM `genehmigungen`) HAVING fTha<>'' OR dmtyp<>'' OR fkas<>'' OR flanr<>'' OR falter<>'' OR fweibl<>'' OR füw<>'' OR fdmp<>''"
' INSTR(DTyp,P_DTyp)=0 OR (...) fDTyp: unnötig
' 9.10.20: bei LKK keine Überweisung verlangen, nach Richtigstellungsmitteilung 1/20 Leistung 97320 nicht moniert (Pat_id 64606)
'sql(AWlf) = "SELECT * FROM (SELECT Pat_id, Name, PAlter, geschlecht P_Gsl, Zeitpunkt, Leistung, Therakt P_Ther, GROUP_CONCAT(DISTINCT Therarten ORDER BY dtyp SEPARATOR '/') L_Ther, MIN(fTha) fTha, P_DTyp, GROUP_CONCAT(DISTINCT DTyp ORDER BY dtyp) L_DTyp," & vbCrLf & _
"IF(ISNULL(myid),'fDTyp!','') fDTyp, Kasse P_Kas, GROUP_CONCAT(DISTINCT Kassen) L_Kas, IF(INSTR(kassen,kasse)=0 AND kassen<>'','fKas!','') fkas, P_Arzt, MIN(Kothny) L_Kothny, MIN(Schade) L_Schade, MAX(fLANR) fLANR, P_Alter, GROUP_CONCAT(DISTINCT AlStuf ORDER BY dtyp) L_Alter, MIN(fAlter) fAlter, IF(INSTR(l_gsl,geschlecht)=0,'fWeibl!','') fweibl, MIN(fÜw) fÜw, P_DMP, GROUP_CONCAT(DISTINCT L_DMP SEPARATOR '/') L_DMP, MIN(fDMP) fDMP, KName " & vbCrLf & _
" FROM ( SELECT f.Pat_id, gesname(f.pat_id) Name,patAlter(f.pat_id) PAlter, " & vbCrLf & _
" l.id LID, l.Zeitpunkt, l.Leistung, maxtha(f.pat_id) Therakt, g.Therarten, " & vbCrLf & _
" IF(therarten<>'' AND INSTR(therarten,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(maxtha(f.pat_id),'?',''),'GLP1ICT','ICT'),'GLP1Ins','Komb'),'GLP1','OAD'),'Komb',IF(d.icd RLIKE '^E10','ICT','Komb')))=0 AND NOT ((d.icd = 'O24.4' AND d.diagsicherheit NOT IN ('A','Z') AND COALESCE(d.Dggel,0)=0) AND l.leistung=97271),'FTh!','') fTha, " & vbCrLf & _
" IF(d.icd RLIKE '^E10',_latin1'1', IF(d.icd RLIKE '^E11',_latin1'2',IF(d.icd = 'O24.4',_latin1'g',null))) P_DTyp, GROUP_CONCAT(DISTINCT g.DTyp) DTyp, g.leistung gleistung, myid, " & vbCrLf & _
" IF(k.kateg='EK','vdek',IF(k.kateg='',k.name,k.kateg)) Kasse, GROUP_CONCAT(DISTINCT g.Kassen) Kassen, IF(k.name='',k.kurzname,k.name) KName" & vbCrLf & _
" ,IF(l.lanr='889690003','Schade',IF(l.lanr='933284903','Kothny',l.lanr)) P_Arzt" & vbCrLf & _
" ,g.kothny Kothny, g.schade Schade, IF((l.lanr='889690003' AND schade = 0) OR (l.lanr='933284903' AND kothny=0),'fLANR!','') fLANR," & vbCrLf & _
" ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24) P_Alter, CAST(CONCAT(IF(von=0,'',von),'-',IF(bis=0,'',bis)) AS char) AlStuf," & vbCrLf & _
" MIN(IF(von<>0 AND (ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24)<von OR ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24)>bis),'fAlter!','')) fAlter," & vbCrLf & _
" n.geschlecht, GROUP_CONCAT(DISTINCT IF(g.weibl=1,'w','wm')) L_Gsl," & vbCrLf & _
" IF(g.mitüw>1 AND f.schgr='00' AND NOT (g.mitüw=2 AND (n.dmpklass=3 OR k.kateg='LKK')) AND NOT n.kvnr IN ('889690003','9333284903'),'fÜw','') fÜw, IF(n.dmpklass IN (2,3),'+','-') P_DMP," & vbCrLf & _
" CONCAT(IF(g.dmp1=1,'T1',''),IF(g.dmp1=1 AND g.dmp2=1,'/',''),IF(g.dmp2=1,'T2','')) L_DMP," & vbCrLf & _
" IF((g.dmp1=1 OR g.dmp2=1) AND NOT n.dmpklass IN (2,3) AND NOT ((k.kateg IN ('LKK','PBe','SHV','') AND l.leistung IN ('97310','973212','97313','97314','97320','97321','97322','97323','97324')) OR (k.kateg='PBe' AND l.leistung IN ('97268','92298D','92698P','92698C','92698S','92698B','92698A','92277','92278','92267A','97267B','92268A','97268B','92266A','97266B','92265A','97265B','92269A','97269B','92264A','97264B','92263A','97263B','92262A','97262B','92261A','97261B','92281','92282','92292S','97267B','92292E','97268B','92292D','97266B','92292C','97265B','92292B','97274','92292A','97268','92711G','92710G','92711S','92710S','92710A','92711A','92710B','97269','97280','97271','97274','97267B','97280S','97276','97277','97270','92278','92281','92282','97312','97310','97320','97321','97322','97313','97323','97333'))),'fDMP','') fDMP" & vbCrLf & _
" FROM `aktfvs` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
" LEFT JOIN `kassenliste` k ON f.vknr = k.vknr AND f.ik = k.ik " & vbCrLf & _
" LEFT JOIN `leistungen` l ON f.fid = l.fid " & vbCrLf & _
" LEFT JOIN `diagnosen` d ON d.pat_id = f.pat_id AND d.diagsicherheit IN (' ','G','V') AND COALESCE(d.Dggel,0)=0 AND icd RLIKE '^E1[0-4]|^O24\.4' AND (obdauer<>0 OR d.fid = f.fid OR d.fid=0 OR ISNULL(d.fid)) " & vbCrLf & _
" LEFT JOIN `anamnesebogen` an ON f.pat_id = an.pat_id LEFT JOIN `genehmigungen` g ON l.leistung = g.leistung " & vbCrLf & _
" AND (IF(d.icd RLIKE '^E10',_latin1'1', IF(d.icd RLIKE '^E11',_latin1'2',IF(d.icd = 'O24.4',_latin1'g',null))) = g.dtyp OR ISNULL(g.dtyp) OR g.dtyp=' ') " & vbCrLf & _
" WHERE l.leistung IN (SELECT leistung FROM `genehmigungen`) GROUP BY pat_id, lid, myid" & vbCrLf & _
" HAVING true) i GROUP BY pat_id, lid) i WHERE fTha<>'' OR fDTyp<>'' OR fkas<>'' OR flanr<>'' OR falter<>'' OR fweibl<>0 OR füw<>'' OR fdmp<>'' OR ISNULL(l_ther)"

' das gleiche, übersichtlicher:
' 1.10.23: AND COALESCE(Dggel,0)=0 wg. Pat. 67849 und
sql(AWlf) = _
"SELECT Pat_id, Name, PAlter, geschlecht P_Gsl, Zeitpunkt, Leistung, Therakt P_Ther " & vbCrLf & _
" ,GROUP_CONCAT(DISTINCT Therarten ORDER BY dtyp SEPARATOR '/') L_Ther, MIN(fTha) fTha, P_DTyp " & vbCrLf & _
" ,GROUP_CONCAT(DISTINCT DTyp ORDER BY dtyp) L_DTyp " & vbCrLf & _
" ,IF(ISNULL(myid),'fDTyp!','') fDTyp, Kasse P_Kas, GROUP_CONCAT(DISTINCT Kassen) L_Kas, IF(INSTR(kassen,kasse)=0 AND kassen<>'','fKas!','') fkas " & vbCrLf & _
" ,P_Arzt, MIN(Kothny) L_Kothny, MIN(Schade) L_Schade, MAX(fLANR) fLANR, P_Alter " & vbCrLf & _
" ,GROUP_CONCAT(DISTINCT AlStuf ORDER BY dtyp) L_Alter, MIN(fAlter) fAlter, IF(INSTR(l_gsl,geschlecht)=0,'fWeibl!','') fweibl " & vbCrLf & _
" ,mitüw,schgr,kateg,kvnr" & vbCrLf & _
" ,MIN(fÜw) fÜw, P_DMP, REGEXP_REPLACE(GROUP_CONCAT(DISTINCT L_DMP SEPARATOR '/'),'^/','')L_DMP,DMPwo,DmpBeg,MIN(fDMP) fDMP, KName " & vbCrLf & _
"FROM ( " & vbCrLf & _
"  SELECT f.Pat_id, gesname(f.pat_id) Name,patAlter(f.pat_id) PAlter, l.id LID, l.Zeitpunkt, l.Leistung, maxtha(f.pat_id) Therakt, g.Therarten " & vbCrLf & _
"  ,IF(therarten<>'' AND INSTR(therarten,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(maxtha(f.pat_id),_utf8mb4'?',_utf8mb4''),_utf8mb4'GLP1ICT',_utf8mb4'ICT'),_utf8mb4'GLP1Ins',_utf8mb4'Komb'),_utf8mb4'GLP1',_utf8mb4'OAD'),_utf8mb4'Komb',IF(d.icd RLIKE '^E10',_utf8mb4'ICT',_utf8mb4'Komb')))=0 AND NOT (d.ICD = 'O24.4' AND l.leistung=97271),'FTh!','') fTha " & vbCrLf & _
"  ,dmtypicd(d.icd) P_DTyp, GROUP_CONCAT(DISTINCT g.DTyp) DTyp, g.leistung gleistung, myid " & vbCrLf & _
"  ,IF(k.kateg='EK','vdek',IF(k.kateg='',k.name,k.kateg)) Kasse, GROUP_CONCAT(DISTINCT g.Kassen) Kassen, IF(k.name='',k.kurzname,k.name) KName " & vbCrLf & _
"  ,IF(l.lanr='889690003','Schade',IF(l.lanr='933284903','Kothny',l.lanr)) P_Arzt " & vbCrLf & _
"  ,g.kothny Kothny, g.schade Schade, IF((l.lanr='889690003' AND schade = 0) OR (l.lanr='933284903' AND kothny=0),'fLANR!','') fLANR " & vbCrLf & _
"  ,ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24) P_Alter, CAST(CONCAT(IF(von=0,'',von),'-',IF(bis=0,'',bis)) AS char) AlStuf " & vbCrLf & _
"  ,MIN(IF(von<>0 AND (ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24)<vON OR ROUND(DATEDIFF(l.zeitpunkt,n.gebdat)/365.24)>bis),'fAlter!','')) fAlter " & vbCrLf & _
"  ,n.geschlecht, GROUP_CONCAT(DISTINCT IF(g.weibl=1,'w','wm')) L_Gsl " & vbCrLf & _
"  ,IF(g.mitüw>1 AND f.schgr='00' AND NOT (g.mitüw=2 AND ((n.dmpklass=3 AND n.dmpbeg<=qend()) /*OR k.kateg='LKK'*/)) AND NOT n.kvnr IN ('889690003','9333284903'),'fÜw','') fÜw " & vbCrLf & _
"  ,IF(n.dmpklass IN (2,3),'+','-') P_DMP,CONCAT(IF(g.dmp1=1,'T1',''),IF(g.dmp1=1 AND g.dmp2=1,'/',''),IF(g.dmp2=1,'T2','')) L_DMP " & vbCrLf & _
"  ,CASE n.DmpKlass WHEN 1 THEN 'nein' WHEN 2 THEN 'HA' WHEN 3 THEN 'hier' WHEN 4 THEN 'ausg.' END DMPwo,dmpbeg" & vbCrLf & _
"  ,IF((g.dmp1=1 OR g.dmp2=1) AND NOT ((dmpklass=2 OR (n.dmpklass=3 AND n.dmpbeg<=qend()))) AND NOT ((k.kateg IN (/*'LKK',*/'PBe','SHV','') AND l.leistung IN ('97310','973212','97313','97314','97320','97321','97322','97323','97324')) OR (k.kateg='PBe' AND l.leistung IN ('97268','92298D','92698P','92698C','92698S','92698B','92698A','92277','92278','92267A','97267B','92268A','97268B','92266A','97266B','92265A','97265B','92269A','97269B','92264A','97264B','92263A','97263B','92262A','97262B','92261A','97261B','92281','92282','92292S','97267B','92292E','97268B','92292D','97266B','92292C','97265B','92292B','97274','92292A','97268','92711G','92710G','92711S','92710S','92710A','92711A','92710B','97269','97280','97271','97274','97267B','97280S','97276','97277','97270','92278','92281','92282','97312','97310','97320','97321','97322','97313','97323','97333'))),'fDMP','') fDMP " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"  ,mitüw,schgr,k.kateg,n.kvnr,dmpklass" & vbCrLf & _
"FROM `aktfvs` f" & vbCrLf & _
"LEFT JOIN `namen` n ON f.pat_id = n.pat_id" & vbCrLf & _
"LEFT JOIN `kassenliste` k ON k.id=f.kid" & vbCrLf & _
"LEFT JOIN `leistungen` l ON l.pat_id = f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() " & vbCrLf & _
"LEFT JOIN `diagview` d ON d.pat_id = f.pat_id AND d.diagsicherheit IN (' ','G') AND COALESCE(Dggel,0)=0 AND icd RLIKE '^E1[0-4]|^O24\.4' AND ((icd rlike '^E1[0-4]' AND obdauer<>0) OR (obdauer=0 AND icd='O24.4' AND diagdatum BETWEEN qanf() AND qend())) " & vbCrLf & _
"LEFT JOIN `anamnesebogen` an ON f.pat_id = an.pat_id " & vbCrLf & _
"LEFT JOIN `genehmigungen` g ON l.leistung = g.leistung AND (dmtypicd(d.icd) = g.dtyp OR ISNULL(g.dtyp) OR g.dtyp=' ') " & vbCrLf & _
"WHERE l.leistung IN (SELECT leistung FROM `genehmigungen`) GROUP BY pat_id, lid, myid) i " & vbCrLf & _
"GROUP BY pat_id, lid " & vbCrLf & _
"HAVING fTha<>'' OR fDTyp<>'' OR fkas<>'' OR flanr<>'' OR falter<>'' OR fweibl<>0 OR füw<>'' OR fdmp<>'' OR ISNULL(l_ther) " & vbCrLf & _
";"
' -- IF(d.icd RLIKE '^E10',_latin1'1', IF(d.icd RLIKE '^E11',_latin1'2',IF(d.icd = 'O24.4',_latin1'g',null))) " & vbCrLf & _
' 19.10.23: 'V' bei diagview entfernt

mins(AWlf) = 10
maxs(AWlf) = 60
AWlf = AWlf + 1
' ...  IF(g.mitüw>1 AND a.schgr='00' AND NOT f.goäkatnr IN ('40','41') AND NOT (g.mitüw=2 AND  => goäkatnr schon aktfv ausgeschlossen
  
' 81
 AwN(AWlf) = "D.m. bei LKK ohne 97155"
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id) PName" & vbCrLf & _
",dt.DmTyp DmTyp" & vbCrLf & _
",(SELECT COUNT(0) FROM leistungen WHERE pat_id=f.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND leistung=97155) Zahl_d_97155" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN namen n USING (pat_id)" & vbCrLf & _
"LEFT JOIN kassenliste k ON k.ID=f.kid" & vbCrLf & _
"LEFT JOIN dtypview dt ON dt.pat_id=f.pat_id" & vbCrLf & _
"WHERE f.kateg='LKK' AND k.kurzname NOT RLIKE 'Bundeswehr' AND dt.DmTyp<>'-' AND n.dmpklass NOT IN (2,3)" & vbCrLf & _
"HAVING Zahl_d_97155<>1"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 82
' ktag fehlerhaft
 AwN(AWlf) = "Falsche Anzahl Grundpauschalen (03001,03002,03003,03004,03005,03011,03012,03013,03014,03015 (lauto) (vorher 54)" '"03110, 03111, 03112, 03120, 03121, 03122)"
 sql(AWlf) = _
 "SELECT f.pat_id PID, gesnameg(f.pat_id) PName, f.Schgr, f.LANRID, COALESCE(SUM(lzahl),0) Zahl" & vbCrLf & _
 ", COALESCE(GROUP_CONCAT(DISTINCT DATE(l.zeitpunkt) SEPARATOR ', '),'') Zeitpunkte" & vbCrLf & _
 ", COALESCE(GROUP_CONCAT(leistung SEPARATOR ', '),'') Leistungen" & vbCrLf & _
 ", IF(ISNULL(SUM(lzahl)),DATE(GREATEST(bhfb,qanf())),DATE(MAX(l.zeitpunkt))) LEIDAT" & vbCrLf & _
 ", CASE WHEN ISNULL(SUM(lzahl)) THEN" & vbCrLf & _
 "   CONCAT(CASE WHEN patalter(f.pat_id)<4 THEN '03001' WHEN patalter(f.pat_id)<18 THEN '03002' WHEN patalter(f.pat_id)<54 THEN '03003' WHEN patalter(f.pat_id)<75 THEN '03004' ELSE '03005' END,' dazu')" & vbCrLf & _
 "   ELSE CONCAT(MIN(l.leistung),' weg')" & vbCrLf & _
 "  END LEIFEHLER" & vbCrLf & _
 ", COALESCE((SELECT 1 FROM rr WHERE pat_id=f.pat_id AND zeitpunkt < qend() LIMIT 1),0) obRR" & vbCrLf & _
 ", COALESCE((SELECT GROUP_CONCAT(DISTINCT art) FROM eintraege WHERE pat_id=f.pat_id AND zeitpunkt < qend() AND art NOT IN ('cia','vac','pa','pvs')),'') Einträge" & vbCrLf & _
 ", COALESCE((SELECT GROUP_CONCAT(DISTINCT art) FROM eintraege WHERE pat_id=f.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND art IN ('cia','vac')),'') Impfeinträge" & vbCrLf & _
 ", COALESCE((SELECT GROUP_CONCAT(DISTINCT form_abk) frm FROM forminhkopf k  LEFT JOIN formulare fo ON k.Form_ID=fo.formid WHERE pat_id=f.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND form_abk NOT IN ('covze','covge')),'') Formulare" & vbCrLf & _
 "FROM aktfvs f" & vbCrLf & _
 "LEFT JOIN `faelle` fl USING (fid)" & vbCrLf & _
 "LEFT JOIN namen n ON n.pat_id=f.pat_id" & vbCrLf & _
 "LEFT JOIN leistungen l ON l.pat_id = f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() AND leistung IN ('03001','03002','03003','03004','03005','03011','03012','03013','03014','03015')" & vbCrLf & _
 "WHERE f.SchGr <> 41 AND f.VKNr <> 71800 AND n.inaktiv<>4 AND fl.nachname NOT LIKE '%Bereitschaft%'" & vbCrLf & _
 "GROUP BY f.pat_id" & vbCrLf & _
 "HAVING (zahl<>0 AND (obRR=0 AND Einträge='' AND Impfeinträge<>'' AND Formulare='')) " & vbCrLf & _
 "    OR (zahl=0 AND (obRR<>0 OR Einträge<>'' OR Formulare<>''));" ' '03110', '03111', '03112', '03120', '03121', '03122'
             ' oder qs=aktq()

 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
#If ebmalt Then
' ktag fehlerhaft
' AwN(AWlf) = "Falsche Anzahl Chronikerpauschalen 03212:"
' sql(AWlf) = "SELECT f.pat_id, gesnameg(f.pat_id) Name, COALESCE(SUM(lzahl),0) Zahl, GROUP_CONCAT(DATE(zeitpunkt) SEPARATOR ', ') Zeitpunkte, GROUP_CONCAT(leistung SEPARATOR ', ') Leistungen " & vbCrLf & _
' "FROM aktfvs f LEFT JOIN namen n ON f.pat_id = n.pat_id LEFT JOIN leistungen l ON f.fid = l.fid AND leistung IN ('03212') AND schgr<>90 GROUP BY pat_id HAVING zahl>1"
' mins(AWlf) = 10
' maxs(AWlf) = 60
' AWlf = AWlf + 1
#End If

 ' 83
 AwN(AWlf) = "Mehr als 6 Insulineinweisungen 97276 (vorher 52)"
 sql(AWlf) = "SELECT f.pat_id, gesnameg(f.pat_id) Name, lz.zeitpunkt " & vbCrLf & _
              "FROM aktfvs f " & vbCrLf & _
              "LEFT JOIN namen n ON f.pat_id = n.pat_id LEFT JOIN leistungen lz ON f.pat_id = lz.pat_id " & vbCrLf & _
              "WHERE EXISTS (SELECT 0 FROM leistungen l WHERE pat_id=f.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND l.leistung = '97276') " & vbCrLf & _
                " AND lz.leistung=97276 AND lz.zeitpunkt BETWEEN " & Khtsfl & "" & vbCrLf & _
                " AND COALESCE((SELECT SUM(lzahl) zahl FROM leistungen " & vbCrLf & _
                "WHERE pat_id = f.pat_id AND leistung = '97276' AND zeitpunkt BETWEEN " & Khtsfl & "),0) > 6 " & vbCrLf & _
                "ORDER BY pat_id, lz.zeitpunkt DESC"
'                "WHERE pat_id = f.pat_id AND leistung = '97276' AND zeitpunkt > SUBDATE(" & qtAnf(FristS) & ", INTERVAL 1 YEAR)),0) > 6 " & vbCrLf & _

 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 84
 AwN(AWlf) = "Potential für 97276 (vorher 153)"
sql(AWlf) = "" & _
"SELECT Pat_id, gesname(pat_id) PName, therart, DATE_FORMAT(zp,'%e.%c.%y') ab" & vbCrLf & _
", InsSeit, InsSeitDatum, CGMseit, 97276_Zahl, 97276_Daten FROM (" & vbCrLf & _
" SELECT f.pat_id, zp, t.therart,a.`Insulin seit` InsSeit, zuDatum(a.`Insulin seit`,qanf()) InsSeitDatum, a.`CGM seit` CGMseit " & vbCrLf & _
" ,(SELECT MIN(zp) FROM therarten WHERE pat_id=f.pat_id AND therart IN ('ICT','GLP1ICT')) ThaICTBeg" & vbCrLf & _
" ,COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE leistung='97276' AND pat_id=f.pat_id AND zeitpunkt BETWEEN qanf()-INTERVAL 9 MONTH AND qend()),0) 97276_Zahl" & vbCrLf & _
" ,COALESCE((SELECT GROUP_CONCAT(DATE_FORMAT(zeitpunkt,'%e.%c.%y')SEPARATOR'|') FROM leistungen WHERE leistung='97276' AND pat_id=f.pat_id AND zeitpunkt BETWEEN qanf()-INTERVAL 9 MONTH AND qend()),'') 97276_Daten" & vbCrLf & _
" FROM aktfv f LEFT JOIN therarten t ON f.pat_id=t.pat_id" & vbCrLf & _
" LEFT JOIN anamnesebogen a ON a.pat_id=f.pat_id" & vbCrLf & _
") i" & vbCrLf & _
"WHERE zp BETWEEN qanf() AND qend()" & vbCrLf & _
"AND therart IN ('ICT','GLP1ICT')" & vbCrLf & _
"AND NOT insSeitDatum<zuDatum('-',qanf())" & vbCrLf & _
"AND ThaICTBeg>=zuDatum('-',qanf())" & vbCrLf & _
"" & vbCrLf & _
""
' "-- , VorTher" & vbCrLf & _
' "-- ,LAG(therart,1) over(PARTITION BY pat_id ORDER BY zp) Vorther" & vbCrLf & _
' "-- AND (Vorther NOT IN ('ICT','GLP1ICT') OR ISNULL(Vorther))" & vbCrLf & _

 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 85
 AwN(AWlf) = "Potential für 97270 (vorher 154)"
sql(AWlf) = "" & _
"SELECT Pat_id, gesname(pat_id) PName, therart, DATE_FORMAT(zp,'%e.%c.%y') ab, VorTher, 97270_Zahl, 97270_Daten FROM (" & vbCrLf & _
"SELECT f.pat_id, zp, t.therart " & vbCrLf & _
", LAG(therart,1) over(PARTITION BY pat_id ORDER BY zp) Vorther" & vbCrLf & _
",COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE leistung='97270' AND pat_id=f.pat_id AND zeitpunkt BETWEEN qanf()-INTERVAL 9 MONTH AND qend()),0) 97270_Zahl" & vbCrLf & _
",(SELECT GROUP_CONCAT(DATE_FORMAT(zeitpunkt,'%e.%c.%y')SEPARATOR'|') FROM leistungen WHERE leistung='97270' AND pat_id=f.pat_id AND zeitpunkt BETWEEN qanf()-INTERVAL 9 MONTH AND qend()) 97270_Daten" & vbCrLf & _
", COALESCE(a.insulinpumpe,0) inspu" & vbCrLf & _
"FROM aktfv f LEFT JOIN therarten t ON f.pat_id=t.pat_id" & vbCrLf & _
"LEFT JOIN anamnesebogen a ON a.pat_id=f.pat_id" & vbCrLf & _
") i" & vbCrLf & _
"WHERE zp BETWEEN qanf() AND qend()" & vbCrLf & _
"AND therart IN ('CSII')" & vbCrLf & _
"AND inspu<>1" & vbCrLf & _
"AND (Vorther NOT IN ('CSII') OR ISNULL(Vorther))" & vbCrLf & _
""
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 86
 AwN(AWlf) = "Möglicherweise zu rasche Motivationsfolge 92278 oder 92282 (vorher 53)"
 sql(AWlf) = "SELECT f.pat_id, gesnameg(f.pat_id) Name, l.leistung, l.zeitpunkt Zp, lz.zeitpunkt lZp " & vbCrLf & _
             "FROM aktfvs f " & vbCrLf & _
             "LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
             "LEFT JOIN leistungen l ON l.pat_id = f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() " & vbCrLf & _
             "LEFT JOIN leistungen lz ON f.pat_id = lz.pat_id " & vbCrLf & _
             "WHERE l.leistung IN ('92278','92282') AND f.quartal <> quartal(lz.zeitpunkt) AND" & vbCrLf & _
             " lz.leistung = l.leistung AND lz.zeitpunkt > SUBDATE(" & qtEnd(FristS) & ",INTERVAL 1 YEAR) "
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' Gesundheitsuntersuchung: 01732 Glucose Cholesterin, Urin, Blutdruck ab 36.LJ mit uns als HA
 
' 87
 ' ktag fehlerhaft
 ' Anmerkung: der selbe reguläre Ausdruck wird auch IN 59 und 155 verwendet
 ' Gegenkontrolle:
 ' SELECT pat_id,zeitpunkt,art,inhalt FROM eintraege WHERE inhalt RLIKE 'biosen'
 ' AND NOT (inhalt RLIKE '[0-9]{1,3}(,[0-9]?|)[ ]?([m]?mg(/dl|)[,]?|aktuell|)( nüchtern| [0-9]{1,2}(\\.[0-9]{2}|) Uhr[,]?|)[ ]?((\\(|)(mit |)|)Biosen(\\)|)|biosen (gemessen|verglichen)|BZ.*[0-9]{2,}.*mit Biosen|Kontrolle mit Biosen|Biosen.*EDTA|
 '  Biosen(Messung|messwert|sor|[ ]{0,2}m|_|s|se|meßwert| \\(plasma\\)| |)(([,:]?[ ]?|\\()[0-9]{1,2}([.:][0-9]{2}|)(\\)|[ ]?(Uhr[:]?|h nü)[ ]?|)|Messvergleich|: Kontrolle zum vorherigen Wert|)[>.;:]*[ ]*([ ]*BZ[:]?|)[ ]*(p[\.]?p |postprandial|nach dem Frühstück|n\.M |nach 1h |nach 2 be |nach [0-9]{1,2}[ ]?min(.|uten)( warten|)|hier |.*gegessen.*|[ ]*(\\(|)nü(\.[:]?|chtern(\\)|))|Leihgerät|re.|[0-9]{1,3} min.*|nach dem Abendessen |)[ ]*[)>.;:]*[ ]*[0-9]+[ ]?[.,/]?[ ]?[0-9o]{0,2}([ ,]{0,2}$|: |[ #/]*([=0-][ ]*[0-9]+(,0|)[ ]*|)([.,] | AB:|\\(|[;,] Ab|\\+ [0-9]|nochmal|verglichen|Kontrolle |spe|Abw.|\., Ab|,mg|m/|/dl|bz/mg|smg|m[, ]g|mm|mg|ng/dl|g/dl|mmg|mh|gm|m%|%|/min|mmHg|,g/|kg|mmmg| Abwei|[,]? umgerech|[0-9]{2}|-))' AND NOT inhalt RLIKE 'Biosen(:|)( [-~] |[ ]*)mg/dl|biosen:$')
 ' AND NOT inhalt RLIKE 'Biosen:$|Biosenglucose,|Antibiosen|Biosen: ~ [^ 0-9]|^([^0-9]+|)Biosen([^0-9]+|)$|Biosen:[ ]{1,2}(-|,|mg/dl)|Biosen[s]?[:]?[ ]+mg/dl|Biosen:[ ]?-(-|#|.|)[ ]+mg|OGTT mit Biosen|Schuhen, Biosen|gleich nochmal mit Biosen|als unser Biosen|Biosen schon abgeschaltet|Biosen(-Gerät|) (schon aus|am )|Biosen war|Biosen zu hoch|Biosen funktioniert|Biosen Geräte hat|Biosen aus( [^E]|$)|Biosen[ ]?(heute|derzeit|) nicht|biosen[^0-9]*(kaput|defekt)|morgen.*Biosen-Gerät|Biosen unauffällig|Biosen mißt.*falsch|Biosentic|Biosen tauschen'
 ' sowohl mit den beiden NOT als auch ohne sie sollen 0 Datensätze rauskommen

AwN(AWlf) = "Fehlende 32025 für Blutzuckermessungen (bz, bzvgl, ogtt) außerhalb der Schwangerschaft (lauto) (vorher 58)"
' sql(AWlf) = "SELECT * FROM (" & vbCrLf & _
      "SELECT f.pat_id, gesname(f.pat_id) Name," & vbCrLf & _
      "DATE(e.zeitpunkt) Tag, COUNT(art) `BZ-Zahl`, SUM(IF(ISNULL(l.LAnzl),0,IF(l.LAnzl='',1,l.LAnzl))) `Leistungs-Zahl`," & vbCrLf & _
      "STR_TO_DATE(fl.letzteRegel,'TM#%d%m%Y') `letzte Regel` " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN faelle fl ON f.fid = fl.fid " & vbCrLf & _
"LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
"LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND art IN ('bz','bzvgl') AND (fl.letzteRegel='' OR NOT e.zeitpunkt BETWEEN STR_TO_DATE(fl.letzteRegel,'TM#%d%m%Y') AND ADDDATE(STR_TO_DATE(fl.letzteRegel,'TM#%d%m%Y'),274)) " & vbCrLf & _
"LEFT JOIN leistungen l ON f.fid = l.fid AND leistung = '32025' AND DATE(l.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
"WHERE NOT ISNULL(art) GROUP BY f.pat_id, DATE(e.zeitpunkt)) i WHERE `BZ-Zahl`<>`Leistungs-Zahl`;"
 sql(AWlf) = "SELECT PID, Name, LEIDAT, LEIFEHLER, LANRID, i.Art, i.Inhalt FROM (" & vbCrLf & _
"SELECT f.pat_id PID, gesname(f.pat_id) Name, DATE(e.zeitpunkt) LEIDAT, COUNT(art) artz,COALESCE(lz,0) lzz, '32025 dazu' LEIFEHLER, f.LANRID, e.Art, e.Inhalt" & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"JOIN faelle fl USING (fid)" & vbCrLf & _
"LEFT JOIN BiosenMessung e ON e.pat_id = f.pat_id " & vbCrLf & _
"      AND e.zeitpunkt BETWEEN fl.bhfb AND fl.bhfe1 " & vbCrLf & _
"      AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & vbCrLf & _
"AND NOT EXISTS (SELECT 0 FROM sws WHERE pat_id=f.pat_id AND e.zeitpunkt BETWEEN voret - INTERVAL 280 DAY AND voret) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id, SUM(lzahl) lz, DATE(zeitpunkt) lzp FROM leistungen WHERE leistung = '32025' GROUP BY pat_id,DATE(zeitpunkt)) l ON f.pat_id = l.pat_id AND lzp = DATE(e.zeitpunkt) " & vbCrLf & _
"WHERE e.inhalt NOT LIKE '%Biosen: - %' AND NOT ISNULL(e.Zeitpunkt) " & vbCrLf & _
"GROUP BY e.id " & vbCrLf & _
") i WHERE artz>lzz " & vbCrLf & _
"ORDER BY PID "
' AND NOT inhalt LIKE '%Kontroll%lösung%' AND NOT inhalt LIKE '%test%lösung%'
' statt voret stand vorher drin: ADDDATE(STR_TO_DATE(fl.letzteRegel,'%d.%m.%Y'),274)
' AND inhalt NOT RLIKE 'Kontroll|lösung|Zielbereich'
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
 ' 88
 AwN(AWlf) = "Eventuell fehlende 97146 (Neumanifestation mit Ernährungsberatung, BMI + ggf. Augen-Üw) (vorher 61)"
 sql(AWlf) = "SELECT v.Pat_id, gesnameg(v.pat_id) Name, DsJ, `Diabetes seit` " & vbCrLf & _
            ",d.icd " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='tk'OR(art='tb'AND ersteller='tk')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk')))) tk " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='gs'OR(art='tb'AND ersteller='gs')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs')))) gs " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='ah'OR(art='tb'AND ersteller='ah')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah')))) ah " & vbCrLf & _
            ",IF(NOT d.icd RLIKE '^E1[01]' OR ISNULL(d.icd),'E10 oder E11','') ICD_zutun " & vbCrLf & _
            ",'ggf.97146 dazu' Leistung_zutun " & vbCrLf & _
            "FROM aktfvs v " & vbCrLf & _
            "LEFT JOIN anb_neuman a on v.pat_id=a.pat_id" & vbCrLf
' AND COALESCE(d.Dggel,0)=0
 sql(AWlf) = sql(AWlf) & _
            "LEFT JOIN diagnosen d ON v.pat_id = d.pat_id AND d.icd RLIKE '^E1[0-4]\.' AND d.diagsicherheit NOT IN ('A','Z','V') " & vbCrLf & _
            "LEFT JOIN _maxHbA1c xh ON v.pat_id = xh.pat_id " & vbCrLf & _
            "LEFT JOIN _maxGluc xg ON v.pat_id = xg.pat_id " & vbCrLf & _
            "LEFT JOIN leistungen chr ON chr.pat_id=v.pat_id AND chr.zeitpunkt BETWEEN qanf()AND qend() AND chr.leistung IN('03220','03221','03220H','03221H') " & vbCrLf & _
            "WHERE (NOT ISNULL(d.ICD) OR xh.max1 >= 6.5 OR xh.max2 >= 6.5 OR xg.max1 >= 200 OR xg.max2 >= 200) AND (DATEDIFF(NOW(), d.DiagDatum) < 180 OR ISNULL(d.DiagDatum)) " & vbCrLf & _
            "AND a.obneu" & vbCrLf & _
            "AND COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id = v.pat_id AND NOT zeitpunkt BETWEEN qanf() AND qend() AND leistung = '97146'),0) < 2 " & vbCrLf & _
            "AND COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id = v.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND leistung = '97146'),0) = 0 " & vbCrLf & _
            "AND NOT ISNULL(a.bmi) AND a.bmi<>0 " & vbCrLf & _
            "AND ISNULL(chr.leistung) " & vbCrLf & _
            "GROUP BY v.pat_id"
' RLIKE '^[0-9][0-9][0-9][0-9]' AND LEFT(IF(LEFT(`Diabetes seit`,4)='ca. ',MID(`Diabetes seit`,5),`Diabetes seit`),4)

 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 89
' ktag fehlerhaft
AwN(AWlf) = "97146 (Ersteinstellung D.m.) neben 03220, 03221, 04220 oder 04221 (vorher 80)"
sql(AWlf) = "SELECT f.pat_id, gesname(f.pat_id) Gesname, DATE(l.zeitpunkt) Tag1, l.leistung Leist1, DATE(l1.zeitpunkt) Tag2, l1.leistung Leist2 " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"       LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() AND l.leistung = '97146' " & vbCrLf & _
"       LEFT JOIN leistungen l1 ON l1.pat_id=f.pat_id AND l1.zeitpunkt BETWEEN qanf() AND qend() AND l1.leistung IN ('03220','03221','04220','04221') " & vbCrLf & _
"WHERE NOT ISNULL(l.leistung) AND NOT ISNULL(l1.leistung);"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 90
' ktag fehlerhaft
AwN(AWlf) = "97146 zu lange nach Erstmanifestation (vorher 81)"
sql(AWlf) = "SELECT f.pat_id, gesname(f.pat_id) Gesname, DATE(l.zeitpunkt) Tag1, l.leistung Leist1, DATE(l1.zeitpunkt) Tag2, l1.leistung Leist2 " & vbCrLf & _
" FROM aktfvs f " & vbCrLf & _
"       LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qanf()AND qend() AND l.leistung = '97146' " & vbCrLf & _
"       LEFT JOIN leistungen l1 ON f.pat_id = l1.pat_id AND l1.leistung = '97146' AND NOT l.zeitpunkt BETWEEN qanf()AND qend()" & vbCrLf & _
" WHERE NOT ISNULL(l.leistung) AND NOT ISNULL(l1.leistung) AND DATEDIFF(qbeg(l.zeitpunkt),qbeg(l1.zeitpunkt))>100;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 91
' AwN(AWlf) = "Fehlende CGM-Ziffern 03355 (-10/Krkhtsf.) nach Rezepten"
' sql(AWlf) = "" & vbCrLf & _
 "SELECT * FROM (" & vbCrLf & _
 "SELECT f.Pat_id, gesname(f.pat_id) Name " & vbCrLf & _
 ", (SELECT SUM(IF(l.LAnzl='',1,l.LAnzl)) FROM leistungen l WHERE pat_id = f.pat_id AND leistung ='03355' AND zeitpunkt >=SUBDATE(" & qtAnf(FristS) & ",274)) bisher " & vbCrLf & _
 ", (SELECT GROUP_CONCAT(DISTINCT DATE(e.zeitpunkt)) FROM eintraege e LEFT JOIN leistungen l ON DATE(e.zeitpunkt) = DATE(l.zeitpunkt) AND l.leistung='03355' WHERE e.pat_id=f.pat_id AND e.zeitpunkt BETWEEN " & lqanfuend(FristS) & " AND e.art IN (" & artSpezBerat & artSpezMA & ") AND ISNULL(l.leistung)) `z.B. fehlend` " & vbCrLf & _
 ", GROUP_CONCAT(DISTINCT DATE(o.zeitpunkt)) rpzp " & vbCrLf & _
 ", GROUP_CONCAT(DISTINCT feldinh) rpfeldinh " & vbCrLf & _
 "FROM aktfvs f LEFT JOIN formular o ON f.pat_id = o.pat_id " & vbCrLf & _
 "WHERE o.form_abk='lar'AND feld='medikament' " & vbCrLf & _
 "AND feldinh RLIKE 'Guardian|CGM StartSet|Enlite|Dexcom|Dexcom|Eversen|CGM' " & vbCrLf & _
 "AND o.zeitpunkt>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 12 MONTH) " & vbCrLf & _
 "GROUP BY f.pat_id) i WHERE bisher<10;"

' 91
' ktag fehlerhaft
' 6.7.25: hier aktuell die einzige Verwendung von kstreng
AwN(AWlf) = "Möglicherweise fehlende 03355 (lauto) (vorher 98)"
 sql(AWlf) = "" & _
 "SELECT PID, LEFT(gesname(PID),15)PName,L3`Q-3`,L2`Q-2`,L1`Q-1`,L0`akt.Q`,bisher`=bisher`,bisher+rng`+rng`" & vbCrLf & _
 ",CASE cgm WHEN 1 THEN'Libre Handy'WHEN 2 THEN'Libre Gerät'WHEN 3 THEN'Dexcom Hdy'WHEN 4 THEN'Dexcom G.'WHEN 5 THEN'Simplera'WHEN 6 THEN'Eversense'WHEN 0 THEN CASE insanw WHEN 3 THEN'Combo'WHEN 4 THEN'Kaleido'WHEN 5 THEN 'Medtronic'WHEN 6 THEN'Omnipod 5'WHEN 8 THEN'T:slim'WHEN 9 THEN'Ypso'ELSE'-'END ELSE'sonst.'END Icon" & vbCrLf & _
 ",LEIDAT, LEIFEHLER, was, Einles, Langrz, Eintr, LANRID" & vbCrLf & _
 "FROM (" & vbCrLf & _
 " SELECT PID, LEIDAT, was" & vbCrLf & _
 " , CONCAT('03355 ',IF(ISNULL(Einles) AND ISNULL(langrz),'fehlt evtl.','dazu')) LEIFEHLER" & vbCrLf & _
 " ,  L3,L2,L1,L0,L3+L2+L1+L0 bisher, Einles, Langrz, Eintr, LANRID,Nachname,cgm,insanw" & vbCrLf & _
 " , ROW_NUMBER()OVER(PARTITION BY i.PID ORDER BY i.zp)rng" & vbCrLf
 sql(AWlf) = sql(AWlf) & _
 " FROM (" & vbCrLf & _
 "  SELECT *" & vbCrLf & _
 "  FROM (" & vbCrLf & _
 "   SELECT f.pat_id PID, DATE(zp) LEIDAT, was" & vbCrLf & _
 "   ,COALESCE((SELECT SUM(lzahl)FROM leistungen WHERE pat_id=f.pat_id AND leistung=03355 AND zeitpunkt BETWEEN qanf()-INTERVAL 9 MONTH AND qend()-INTERVAL 9 MONTH),0)L3" & vbCrLf & _
 "   ,COALESCE((SELECT SUM(lzahl)FROM leistungen WHERE pat_id=f.pat_id AND leistung=03355 AND zeitpunkt BETWEEN qanf()-INTERVAL 6 MONTH AND qend()-INTERVAL 6 MONTH),0)L2" & vbCrLf & _
 "   ,COALESCE((SELECT SUM(lzahl)FROM leistungen WHERE pat_id=f.pat_id AND leistung=03355 AND zeitpunkt BETWEEN qanf()-INTERVAL 3 MONTH AND qend()-INTERVAL 3 MONTH),0)L1" & vbCrLf & _
 "   ,COALESCE((SELECT SUM(lzahl)FROM leistungen WHERE pat_id=f.pat_id AND leistung=03355 AND zeitpunkt BETWEEN qanf()-INTERVAL 0 MONTH AND qend()-INTERVAL 0 MONTH),0)L0" & vbCrLf & _
 "   ,(SELECT CONCAT(DATE_FORMAT(zeitpunkt,'%d.%m.%y'),' ',GROUP_CONCAT(DISTINCT LEFT(name,40)ORDER BY zeitpunkt DESC SEPARATOR'; ')) FROM tmbrie WHERE pat_id=f.pat_id AND name RLIKE 'Omnipod ?5|Simplera|Guardian|Enlite|Dexc?om|Eversen|CGM|Libre|clarity|glooko|care ?link' AND NOT name RLIKE 'cgm bmp' AND qdm>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 6 MONTH)) Einles" & vbCrLf & _
 "   ,(SELECT CONCAT(DATE_FORMAT(zeitpunkt,'%d.%m.%y'),' ',GROUP_CONCAT(DISTINCT LEFT(medikament,20)ORDER BY zeitpunkt DESC SEPARATOR'; ')) FROM rezepteintraege WHERE pat_id=f.pat_id AND medikament RLIKE 'Omnipod ?5|Simplera|Guardian|Enlite|Dexcom|Eversen|CGM|Libre')Langrz" & vbCrLf & _
 "   ,(SELECT CONCAT(DATE_FORMAT(zeitpunkt,'%d.%m.%y'),' ',GROUP_CONCAT(DISTINCT CONCAT(art,' ',LEFT(inhalt,40))ORDER BY zeitpunkt DESC SEPARATOR'; ')) FROM eintraege e WHERE pat_id=f.pat_id AND inhalt RLIKE 'Omnipod ?5|Simplera|Guardian|Enlite|Dexc?om|Eversen|CGM|Libre|clarity|glooko|care ?link' AND NOT e.inhalt LIKE '%vorgestellt%will%nicht%' AND zeitpunkt BETWEEN qanf()-INTERVAL 3 MONTH AND qend() ORDER BY zeitpunkt DESC) Eintr" & vbCrLf & _
 "   ,COALESCE((SELECT lzahl FROM leistungen WHERE pat_id=f.pat_id AND leistung='03355' AND DATE(zeitpunkt)=DATE(f.zp) GROUP BY pat_id, DATE(zeitpunkt)),0) bish " & vbCrLf & _
 "   ,LANRID,zp,cgm,insanw,Nachname" & vbCrLf & _
 "   FROM kstreng f" & vbCrLf & _
 "   LEFT JOIN namen n ON n.pat_id=f.pat_id" & vbCrLf & _
 "  ) i WHERE (einles IS NOT NULL OR langrz IS NOT NULL OR eintr IS NOT NULL OR cgm<>0) AND bish=0 AND NOT ISNULL(LEIDAT)" & vbCrLf & _
 " ) i" & vbCrLf & _
 ") i" & vbCrLf & _
 "WHERE rng+bisher<=7" & vbCrLf & _
 " AND EXISTS (SELECT 0 FROM diagview WHERE pat_id=PID AND (gicd REGEXP '^E1[0-4]' OR (icd='O24.4' AND Dggel=0 AND diagsicherheit IN ('G',' ') AND diagdatum BETWEEN qanf()AND qend())))" & vbCrLf & _
 "ORDER BY pid,leidat"
 ' mit HAVING geht's 5% langsamer als mit stärker geschachtelter Abfrage! 13.7.22
' ",COALESCE((SELECT DATE(MAX(dokaend)) FROM tmbrie WHERE pat_id=PID AND name LIKE '%.pdf' " & vbCrLf & _
 "  AND name LIKE CONCAT('%',Nachname,'%') AND name NOT RLIKE ' DS | DS+| ÜW |Arztb|labor|Ergebnis|Bericht|Mutterpass|Seminar|Notfall| BZ |OGT|Radiologie|plan|DMP|befund|ausstell|verord|Aufkl|schein|covid|fvk|quarant|ber.| rr|gravid|rezept|nachsorg|ausweis|amcl|iief|protokoll|liste|anamnese|ekg|impf|zbfs|versorg|schreiben|programm|erklär|abdomen|BZ-Werte|endosk|gutachten|klinik|bestät|unterschr|behandl|doku|rechtsanw|attest|füw|mdk|frage|modul|entbind|wunsch|änder|kh ffb|ausgef|maßnahm|studie|metrie|kardiol|rechnu| kd |dachau|zentrum|enddarm|befrei|neuro|neurpad|rechnung|erlaubn|formul|ausgef.|lmu |ögd|colo|entlass|reha|blutdr|schema|mrt|thorax|angio| au |sono|szinti|radiol rp |knochen|skopie|gpd|augen|zuzahl|arteri| ct |wund|zeugn|an fax|zusatz|symptom|verläng|brief|praxis|pflege|rezept|anford|zettel|kranken|ernährung|vertrag|eigen|kompress|kassenä|paß|orthopäd|röntg|teiln|ablehn|pathol|training|contour|einschreib'),'-') `ltz.Einles`" & vbCrLf & _
 ",COALESCE((SELECT DATE(MAX(dokaend)) FROM tmbrie WHERE pat_id=PID AND name LIKE '%libre%.pdf'),'-') `ltz.'Libre'`" & vbCrLf & _

' @ geht mit dem ADO nicht richtig
' sql(AWlf) = "" & _
 "SELECT PID, LEFT(gesname(PID),15) PName,bisher,bisher+rng " & vbCrLf & _
 ",COALESCE((SELECT LEFT(titel,6) FROM desktop WHERE pat_id=PID AND iconPath LIKE '%CGM.ico%' LIMIT 1),'-') Icon " & vbCrLf & _
 ",COALESCE((SELECT DATE(MAX(dokaend)) FROM tmbrie WHERE pat_id=PID AND name LIKE '%.pdf' " & vbCrLf & _
 "  AND name LIKE CONCAT('%',n.Nachname,'%') AND name NOT RLIKE ' DS | DS+| ÜW |Arztb|labor|Ergebnis|Bericht|Mutterpass|Seminar|Notfall| BZ |OGT|Radiologie|plan|DMP|befund|ausstell|verord|Aufkl|schein|covid|fvk|quarant|ber.| rr|gravid|rezept|nachsorg|ausweis|amcl|iief|protokoll|liste|anamnese|ekg|impf|zbfs|versorg|schreiben|programm|erklär|abdomen|BZ-Werte|endosk|gutachten|klinik|bestät|unterschr|behandl|doku|rechtsanw|attest|füw|mdk|frage|modul|entbind|wunsch|änder|kh ffb|ausgef|maßnahm|studie|metrie|kardiol|rechnu| kd |dachau|zentrum|enddarm|befrei|neuro|neurpad|rechnung|erlaubn|formul|ausgef.|lmu |ögd|colo|entlass|reha|blutdr|schema|mrt|thorax|angio| au |sono|szinti|radiol rp |knochen|skopie|gpd|augen|zuzahl|arteri| ct |wund|zeugn|an fax|zusatz|symptom|verläng|brief|praxis|pflege|rezept|anford|zettel|kranken|ernährung|vertrag|eigen|kompress|kassenä|paß|orthopäd|röntg|teiln|ablehn|pathol|training|contour|einschreib'),'-') `ltz.Einles`" & vbCrLf & _
 ",COALESCE((SELECT DATE(MAX(dokaend)) FROM tmbrie WHERE pat_id=PID AND name LIKE '%libre%.pdf'),'-') `ltz.'Libre'`" & vbCrLf & _
 ",LEIDAT, LEIFEHLER, was, Einles, Langrz, Eintr, bisher, rng, LANRID " & vbCrLf & _
 "FROM (" & vbCrLf & _
 " SELECT PID, LEIDAT, was" & vbCrLf & _
 " , CONCAT('03355 ',IF(ISNULL(Einles) AND ISNULL(langrz),'fehlt evtl.','dazu')) LEIFEHLER" & vbCrLf & _
 " , bisher, Einles, Langrz, Eintr, LANRID" & vbCrLf & _
 " , CASE WHEN @vorp=i.PID THEN @nr:=@nr+1 ELSE @nr:=1 END rng" & vbCrLf & _
 " , @vorp:=i.PID vorpid" & vbCrLf
' sql(AWlf) = sql(AWlf) & _
 " FROM (" & vbCrLf & _
 "  SELECT *" & vbCrLf & _
 "  FROM (" & vbCrLf & _
 "   SELECT pat_id PID, DATE(zp) LEIDAT, was" & vbCrLf & _
 "   , COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='03355' AND zeitpunkt BETWEEN " & Khtsfl & "),0) bisher" & vbCrLf & _
 "   , (SELECT CONCAT(DATE(zeitpunkt),' ',LEFT(name,40)) FROM tmbrie WHERE pat_id=f.pat_id AND name RLIKE 'libre|dexcom|cgm|clarity|medtronic|care link|eversen' AND NOT name RLIKE 'cgm bmp' AND qdm>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 6 MONTH) LIMIT 1) Einles" & vbCrLf & _
 "   , (SELECT CONCAT(DATE(fr.zeitpunkt),' ',LEFT(fr.feldinh,40)) FROM formular fr WHERE fr.pat_id=f.pat_id AND fr.form_abk='lar' AND fr.feld='medikament' AND fr.feldinh RLIKE 'Guardian|CGM StartSet|Enlite|Dexcom|Dexom|Eversen|CGM|Libre' AND fr.zeitpunkt>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 12 MONTH) LIMIT 1) Langrz" & vbCrLf & _
 "   , (SELECT CONCAT(art,' ',inhalt) FROM eintraege e WHERE e.pat_id=f.pat_id AND e.inhalt RLIKE 'cgm|eversen|guardian|enlite|dexcom|libre' AND NOT e.inhalt LIKE '%vorgestellt%will%nicht%' AND e.zeitpunkt BETWEEN " & Khtsfl & " LIMIT 1) Eintr" & vbCrLf & _
 "   , (SELECT lzahl FROM leistungen WHERE pat_id=f.pat_id AND leistung='03355' AND DATE(zeitpunkt)=DATE(f.zp) GROUP BY pat_id, DATE(zeitpunkt)) bish " & vbCrLf & _
 "   , LANRID" & vbCrLf & _
 "   FROM kstreng f" & vbCrLf & _
 "  ) i WHERE (NOT ISNULL(Einles) OR NOT ISNULL(Langrz) OR NOT ISNULL(eintr)) AND ISNULL(bish) AND NOT ISNULL(LEIDAT)" & vbCrLf & _
 " ) i" & vbCrLf & _
 ") i LEFT JOIN namen n ON n.pat_id=PID WHERE rng+bisher<=10 " & vbCrLf & _
 "ORDER BY pid,leidat"
 
' sql(AWlf) = vbCrLf & _
"SELECT * FROM ( " & vbCrLf & _
"SELECT f.pat_id PID, LEFT(gesname(f.pat_id),20) PatName, " & vbCrLf & _
"(SELECT SUM(IF(lb.LAnzl='',1,lb.LAnzl)) FROM leistungen lb WHERE lb.pat_id=f.pat_id AND lb.leistung='03355' AND lb.zeitpunkt>=SUBDATE(" & qtAnf(FristS) & ",274)) bisher, " & vbCrLf & _
"DATE_FORMAT(kon.zeitpunkt,'%d.%m.%y') KontZP, GROUP_CONCAT(DISTINCT kon.art) KArt, DATE_FORMAT(IF(b.quelldatum=18991230,b.zeitpunkt,b.quelldatum),'%d.%m.%y') EinlZP, LEFT(b.name,40) Einlesung, DATE_FORMAT(fr.zeitpunkt,'%d.%m.%y') LangrzZP,LEFT(fr.feldinh,40) LangRZ,DATE_FORMAT(e.zeitpunkt,'%d.%m.%y') EintrZP,e.inhalt Eintrag " & vbCrLf & _
"FROM aktfv f " & vbCrLf & _
"LEFT JOIN tmbrie b ON b.pat_id=f.pat_id AND name LIKE '%libre%' AND IF(b.quelldatum=18991230,b.zeitpunkt,b.quelldatum)>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 6 MONTH) " & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.inhalt RLIKE 'cgm|eversen|guardian|enlite|dexcom|libre' AND NOT e.inhalt LIKE '%vorgestellt%will%nicht%' AND e.zeitpunkt>SUBDATE(" & qtAnf(FristS) & ",274) " & vbCrLf & _
"LEFT JOIN formular fr ON fr.pat_id=f.pat_id AND fr.form_abk='lar' AND fr.feld='medikament' AND fr.feldinh RLIKE 'Guardian|CGM StartSet|Enlite|Dexcom|Dexcom|Eversen|CGM|Libre' AND fr.zeitpunkt>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 12 MONTH) " & vbCrLf & _
"LEFT JOIN eintraege kon ON kon.pat_id=f.pat_id AND kon.zeitpunkt BETWEEN " & lqanfuend(FristS) & " AND kon.art IN (" & artSpezBerat & ",'usdm2','sem','uzu','hypo','notiz','rauch','debr','anal','andm','usal','sono','doppler','dop','duplex','ulcus','wv','kv')" & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.leistung='03355' AND DATE(l.zeitpunkt)=DATE(kon.zeitpunkt) " & vbCrLf & _
"WHERE (NOT ISNULL(b.name) OR (NOT ISNULL(e.inhalt) AND e.zeitpunkt>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 12 MONTH)) " & vbCrLf & _
"OR NOT ISNULL(fr.feldinh) " & vbCrLf & _
") AND ISNULL(l.leistung) AND NOT ISNULL(kon.zeitpunkt) " & vbCrLf & _
"GROUP BY f.pat_id, DATE(kon.zeitpunkt)) i WHERE bisher<10 ORDER BY PID, KontZP"
', r.medikament Rz, r.zeitpunkt RzZp
'"LEFT JOIN rezepteintraege r ON r.pat_id=f.pat_id AND r.medikament LIKE '%libre%' AND r.zeitpunkt>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 12 MONTH) "
'OR NOT ISNULL(r.medikament)
 mins(AWlf) = 10
 maxs(AWlf) = 400
 AWlf = AWlf + 1

' 92
' ktag fehlerhaft
 AwN(AWlf) = "evtl. fehlende Eversense-cgma-Makro oder 03355(+Zusatz)-Leistung (lauto) (vorher 99)"
 ' AND COALESCE(dd.Dggel,0)=0
 sql(AWlf) = _
 "SELECT PID, gesname(pid) PName" & vbCrLf & _
 ",COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=i.pid AND l.zeitpunkt BETWEEN " & Khtsfl & " AND l.leistung='03355'),0) Geszahl " & vbCrLf & _
 ",DATE(zp) LEIDAT " & vbCrLf & _
 ",IF(ISNULL(l.leistung),'03355 dazu','') LEIFEHLER " & vbCrLf & _
 ",CONCAT(l.leistung,'(',LEFT(l.LfBegr,10),IF(ISNULL(l.LfBegr),'','...'),')') Leist " & vbCrLf & _
 ",DATE(ec.zeitpunkt) MakroZp,ec.inhalt Makro " & vbCrLf & _
 ",GROUP_CONCAT(inh SEPARATOR ';') inh,i.LANRID " & vbCrLf & _
 "FROM ( " & vbCrLf & _
 " SELECT i.*,f.lanrid FROM aktfvs f" & vbCrLf & _
 " LEFT JOIN (" & vbCrLf & _
 "   SELECT pat_id pid, zeitpunkt zp, CONCAT(art,': ',inhalt) inh FROM eintraege WHERE pat_id IN (SELECT pat_id FROM aktfvs) AND  inhalt LIKE '%eversen%' AND inhalt RLIKE 'plant|einges|setz|entf|gelegt|gsetz|wechs' AND zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
 "    UNION ALL" & vbCrLf & _
 "   SELECT pat_id PID, IF(quelldatum<20000101,zeitpunkt,quelldatum) zp, CONCAT('Dok (',DATE(zeitpunkt),'): ',name) inh FROM tmbrie WHERE pat_id IN (SELECT pat_id FROM aktfvs) AND name LIKE '%eversen%' HAVING zp BETWEEN qanf() AND qend()" & vbCrLf & _
 "    UNION ALL" & vbCrLf & _
 "   SELECT PID, zp, CONCAT('T: ',REPLACE(LEFT(zusatz,40),'\n','')) inh FROM termine WHERE pid IN (SELECT pat_id FROM aktfvs) AND zusatz RLIKE 'everse' AND zp BETWEEN qanf() AND qend()" & vbCrLf & _
 " ) i ON i.pid=f.pat_id " & vbCrLf & _
 ") i " & vbCrLf & _
 "LEFT JOIN leistungen l ON l.pat_id=i.pid AND l.leistung='03355' AND DATE(l.zeitpunkt)=DATE(i.zp) " & vbCrLf & _
 "LEFT JOIN eintraege ec ON ec.pat_id=i.pid AND ec.art RLIKE 'cgm' AND ec.inhalt RLIKE 'eversen' AND DATE(ec.zeitpunkt)<=i.zp " & vbCrLf & _
 "LEFT JOIN `diagnosen` dd ON i.pid = dd.pat_id AND dd.icd REGEXP '^E1[0-4]' AND dd.diagsicherheit NOT IN ('A','Z','V') " & vbCrLf & _
 "WHERE NOT ISNULL(dd.icd) AND NOT ISNULL(zp) AND (ISNULL(l.leistung) OR ISNULL(ec.pat_id)) GROUP BY pid,DATE(zp) HAVING Geszahl<7"
 
' halb so schnell:
'  " SELECT f.pat_id PID, e.zeitpunkt zp, CONCAT(e.art,': ',e.inhalt) inh,f.LANRID " & vbCrLf & _
 " FROM aktfvs f " & vbCrLf & _
 "  LEFT JOIN eintraege e ON f.pat_id=e.pat_id AND e.inhalt LIKE '%eversen%' AND e.inhalt RLIKE 'plant|einges|setz|entf|gelegt|gsetz|wechs' AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
 " UNION ALL " & vbCrLf & _
 " SELECT f.pat_id PID, qd zp, CONCAT('Dok (',DATE(b.zeitpunkt),'): ',b.name) inh,f.LANRID " & vbCrLf & _
 " FROM aktfvs f " & vbCrLf & _
 "  LEFT JOIN (SELECT IF(quelldatum<20000101,zeitpunkt,quelldatum) qd,name,zeitpunkt,pat_id FROM tmbrie) b ON f.pat_id=b.pat_id AND b.qd BETWEEN " & lQAnfuEnd(FristS) & " AND b.name LIKE '%eversen%' " & vbCrLf & _
 " UNION ALL " & vbCrLf & _
 " SELECT f.pat_id PID, t.zp zp, CONCAT('T: ',REPLACE(LEFT(t.zusatz,40),'\n','')) inh,f.LANRID " & vbCrLf & _
 " FROM aktfvs f " & vbCrLf & _
 "  LEFT JOIN termine t ON f.pat_id = t.pid AND t.zusatz RLIKE 'everse' AND t.zp BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _

 
' sql(AWlf) = vbCrLf & _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, COALESCE(DATE(l.zeitpunkt),DATE(t.zp),DATE(ec.zeitpunkt),DATE(b.zeitpunkt)) LEIDDAT" & vbCrLf & _
",IF(ISNULL(l.leistung),'03355 dazu','') LEIFEHLER" & vbCrLf & _
",CONCAT(l.leistung,'(',LEFT(l.LfBegr,10),IF(ISNULL(l.LfBegr),'','...'),')') Leist" & vbCrLf & _
",DATE(ec.zeitpunkt) MakroZp,ec.inhalt Makro" & vbCrLf & _
",CONCAT(DATE(t.zp),'(',REPLACE(LEFT(t.zusatz,20),'\n',''),IF(ISNULL(t.zusatz),'','...'),')') Termin" & vbCrLf & _
",DATE(e.zeitpunkt) EintrZP,CONCAT(e.art,': ',e.inhalt) Eintrag,DATE(b.zeitpunkt) DokZp,b.qd DokQellDt,b.name Dokument" & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN eintraege e ON f.pat_id=e.pat_id AND e.inhalt LIKE '%eversen%' AND e.inhalt RLIKE 'plant|einges|setz|entf|gelegt|gsetz|wechs|bestell' AND QUARTAL(e.zeitpunkt)=aktq() " & vbCrLf & _
"LEFT JOIN (SELECT DATE(fIF(quelldatum<20000101,zeitpunkt,quelldatum)) qd,name,zeitpunkt,pat_id FROM tmbrie) b ON f.pat_id=b.pat_id AND QUARTAL(b.qd)=aktq() AND b.name LIKE '%eversen%' " & vbCrLf & _
"LEFT JOIN leistungen l ON l.fid=f.fid AND l.leistung='03355' AND DATE(l.zeitpunkt) IN (b.qd,DATE(e.zeitpunkt)) " & vbCrLf & _
"LEFT JOIN eintraege ec ON ec.pat_id=f.pat_id AND ec.art RLIKE 'cgm' AND ec.inhalt RLIKE 'eversen' AND DATE(ec.zeitpunkt) IN (DATE(b.qd),DATE(e.zeitpunkt)) " & vbCrLf & _
"LEFT JOIN termine t ON f.pat_id = t.pid AND t.zusatz RLIKE 'evers' AND QUARTAL(t.zp)=aktq() " & vbCrLf & _
"WHERE (NOT e.art IS NULL OR NOT b.name IS NULL OR NOT ISNULL(t.Zp)) " & vbCrLf & _
"GROUP BY f.pat_id,DATE(e.zeitpunkt),b.qd,DATE(t.zp) ORDER BY f.pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 500
 AWlf = AWlf + 1
 ' AND l.LfBegr RLIKE 'eversen'

' AwN(AWlf) = "Eventuell fehlende CGM-Ziffern (03355) nach Einträgen"
' sql(AWlf) = "" & vbCrLf & _
 "SELECT e.Pat_id, gesname(e.pat_id) Pat_ID, DATE(e.zeitpunkt) Tag,e.Art, e.Inhalt FROM aktfvs f LEFT JOIN eintraege e ON f.pat_id = e.pat_id " & vbCrLf & _
 "LEFT JOIN leistungen l ON l.pat_id = e.pat_id AND DATE(l.zeitpunkt)=DATE(e.zeitpunkt) AND l.leistung='03355' " & vbCrLf & _
 "WHERE e.inhalt RLIKE 'cgm|eversen|guardian|enlite|dexcom' AND e.zeitpunkt > " & qtAnf(FristS) & " " & vbCrLf & _
 "AND ISNULL(l.leistung) " & vbCrLf & _
 "AND e.pat_id NOT IN " & vbCrLf & _
 " (SELECT pat_id FROM ( " & vbCrLf & _
 " SELECT f.Pat_id, gesname(f.pat_id) Name " & vbCrLf & _
 " , (SELECT COUNT(leistung) FROM leistungen WHERE pat_id = f.pat_id AND leistung ='03355' AND zeitpunkt >=SUBDATE(" & qtAnf(FristS) & ",274)) bisher " & vbCrLf & _
 " FROM aktfvs f LEFT JOIN formular o ON f.pat_id = o.pat_id " & vbCrLf & _
 " WHERE o.form_abk='lar'AND feld='medikament' " & vbCrLf & _
 " AND feldinh RLIKE 'Guardian|CGM StartSet|Enlite|Dexcom|Eversen' " & vbCrLf & _
 " AND o.zeitpunkt>SUBDATE(" & qtAnf(FristS) & ",INTERVAL 3 YEAR) " & vbCrLf & _
 " GROUP BY f.pat_id) i WHERE bisher<10) " & vbCrLf & _
 "GROUP BY e.pat_id ORDER BY e.pat_id;"
'  mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
 
' 93
 AwN(AWlf) = "Mehr als 7 Leistungen 03355 im Krankheitsfall (vorher 113)"
 sql(AWlf) = vbCrLf & _
"SELECT * FROM (SELECT " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.zeitpunkt BETWEEN " & Khtsfl & " AND l.leistung='03355'),0) Geszahl, " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='03355' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & "),0) Aktzahl, " & vbCrLf & _
"gesname(f.pat_id), f.pat_id " & vbCrLf & _
", (SELECT GROUP_CONCAT(zeitpunkt SEPARATOR ', ') FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='03355' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Zpkte " & vbCrLf & _
"FROM aktfvs f) i WHERE geszahl>7 AND aktzahl>0;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 94, neuView
AwN(AWlf) = "Schulungen"
sql(AWlf) = "ü"
 AWlf = AWlf + 1

' 95
' ktag fehlerhaft
AwN(AWlf) = "Schulungseintrag ohne erkennbare Schulungsziffer (" & myEFrag("SELECT GROUP_CONCAT(leistung) FROM genehmigungen WHERE obschulung<>0").Fields(0) & ") (vorher 49)"
sql(AWlf) = "SELECT e.Pat_id, gesname(e.pat_id) Name, DATE(e.zeitpunkt), e.inhalt FROM `eintraege` e " & vbCrLf & _
            "INNER JOIN `aktfvs` f ON e.pat_id = f.pat_id " & vbCrLf & _
            "LEFT JOIN `leistungen` l ON l.pat_id = e.pat_id AND DATE(l.zeitpunkt)= DATE(e.zeitpunkt) AND l.leistung IN (SELECT DISTINCT leistung FROM `genehmigungen` WHERE obschulung<>0) " & vbCrLf & _
            "WHERE e.art = 'schul' AND NOT e.inhalt RLIKE 'fehlt|entschuldigt' AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND ISNULL(leistung) ORDER BY pat_id, DATE(e.zeitpunkt)"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 83 ' in 33 enthalten
' "Doppelte 03221/ 03221H"
' "SELECT * FROM (" & vbCrLf & _
"       SELECT f.pat_id, gesname(f.pat_id) Name, SUM(l0.lzahl) Zahl," & vbCrLf & _
"       GROUP_CONCAT(l0.leistung) Leistung0" & vbCrLf & _
"       FROM aktfvs f " & vbCrLf & _
"       LEFT JOIN leistungen l0 ON f.pat_id = l0.pat_id AND l0.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"       AND l0.leistung IN ('03221','03221H')" & vbCrLf & _
"       GROUP BY f.fid" & vbCrLf & _
"       ) i WHERE Zahl>1"
' 96
AwN(AWlf) = "Mehr Leistungen seit 9 Monaten vor Quartalsanfang als in `genehmigungen` eingetragen (vorher 83)"
sql(AWlf) = "SELECT i.pat_id, gesname(i.pat_id) PName,COALESCE(SUM(zl.lzahl),0) lzahl,i.maxzdue MaxZ,i.Leistung" & vbCrLf & _
",GROUP_CONCAT(DATE_FORMAT(zl.zeitpunkt,'%e.%c.%y')SEPARATOR'|') lzdat, i.Erklärung" & vbCrLf & _
"FROM" & vbCrLf & _
"(SELECT f.pat_id,l.leistung,g.maxzdue,g.Erklärung" & vbCrLf & _
" FROM aktfv f" & vbCrLf & _
" LEFT JOIN leistungen l ON l.Pat_ID=f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
" LEFT JOIN genehmigungen g ON g.leistung=l.leistung" & vbCrLf & _
" WHERE NOT ISNULL(g.myid) AND NOT (minzdue=0 AND maxzdue=0)" & vbCrLf & _
" GROUP BY f.pat_id,l.leistung) i" & vbCrLf & _
"LEFT JOIN leistungen zl ON zl.pat_id=i.pat_id AND zl.Leistung=i.Leistung AND zl.zeitpunkt BETWEEN qanf()-INTERVAL 9 MONTH AND qend()" & vbCrLf & _
"GROUP BY i.pat_id, i.leistung" & vbCrLf & _
"HAVING lzahl>maxzdue" & vbCrLf & _
";"
mins(AWlf) = 10
maxs(AWlf) = 60
AWlf = AWlf + 1
  
' 97 ' in 33
 AwN(AWlf) = "Übersicht Einzelschulungen (vorher 117)"
 sql(AWlf) = _
"SELECT l.pat_id" & vbCrLf & _
",IF(RANK() OVER(PARTITION BY l.pat_id ORDER BY l.pat_id,l.zeitpunkt)=1,gesname(l.pat_id),'') PName" & vbCrLf & _
",l.zeitpunkt,l.leistung,IF(NOT ISNULL(sws.voret),1,0) obschwanger, n.obneu,d.dmtyp,IF(lanr LIKE '9%','tk','gs') Arzt" & vbCrLf & _
",IF(ISNULL(sws.voret) AND NOT (obneu AND d.dmtyp=1),CAST(COUNT(0) OVER(PARTITION BY lanr,IF(NOT ISNULL(sws.voret),1,0)) AS CHAR),'-') Arztzahl" & vbCrLf & _
"FROM leistungen l" & vbCrLf & _
"INNER JOIN (SELECT leistung FROM genehmigungen WHERE obschulung=2 GROUP BY leistung) g ON l.leistung=g.leistung" & vbCrLf & _
"LEFT JOIN sws on sws.pat_id=l.pat_id AND sws.voret>l.zeitpunkt" & vbCrLf & _
"LEFT JOIN anb_neuman n ON n.pat_id=l.pat_id" & vbCrLf & _
"LEFT JOIN (SELECT pat_id,IF(icd RLIKE '^E',MID(icd,3,1)+1,'g') dmtyp FROM diagview d WHERE (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qanf() AND qend())) GROUP BY pat_id) d ON d.pat_id=l.pat_id" & vbCrLf & _
"WHERE l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"ORDER BY l.pat_id, l.zeitpunkt;"
 
 ' "Mehr als 1 Leistung 03221/03221H im Behandlungsfall"
' sql(AWlf) = "" ' vbCrLf & _
"SELECT * FROM (SELECT " & vbCrLf & _
"(SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung IN ('03221','03221H') AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Aktzahl, " & vbCrLf & _
"gesname(f.pat_id), f.pat_id " & vbCrLf & _
", (SELECT GROUP_CONCAT(zeitpunkt SEPARATOR ', ') FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung IN ('03221','03221H') AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Zpkte " & vbCrLf & _
"FROM aktfvs f) i WHERE aktzahl>1;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 98
 AwN(AWlf) = "Prüfung auf Einzelschulungen bei T1D (ICT DMP 92392A, DV 97368E, Linda DMP 92392B, DV 97274E, Primas DMP 92392F, DV 97281E, Kinder DMP 92390A, DV 97278E, Jug DMP 92391A, DV 97279E, Hypert DMP 92392C, DV 97265E, HPSB DMP 92392D, DV 97266E, IMP DMP 92392E, DV 97268E, BGAT DV 97269E, Hypos DV 97280E) (vorher 150)"
'sql(AWlf) = "" & _
" SELECT f.pat_id " & vbCrLf & _
", IF (ROW_NUMBER() over (PARTITION BY f.pat_id ORDER BY e.zeitpunkt)=1,gesname(f.pat_id),'') PName " & vbCrLf & _
", IF (ROW_NUMBER() over (PARTITION BY f.pat_id ORDER BY e.zeitpunkt)=1,IF(s.voret=18991230,'',DATE_FORMAT(s.voret,'%e.%c.%y')),'') vorET " & vbCrLf & _
", IF (ROW_NUMBER() over (PARTITION BY f.pat_id ORDER BY e.zeitpunkt)=1,dt.ityp,'') dtyp " & vbCrLf & _
", DATE_FORMAT(e.zeitpunkt,'%d.%m.%y %H:%i') zp, e.art, e.inhalt " & vbCrLf & _
"FROM aktfv f " & vbCrLf & _
"LEFT JOIN faelle fa USING (fid) " & vbCrLf & _
"LEFT JOIN sws s ON s.pat_id=f.pat_id AND s.voret>qanf()" & vbCrLf & _
"LEFT JOIN dtypen dt ON dt.pat_id=f.pat_id " & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND DATE(l.zeitpunkt)=DATE(e.zeitpunkt) AND l.leistung IN ('92392A','97368E','92392B', '97274E','92392F','97281E', '92390A','97278E','92391A','97279E','92392C','97265E','92392D','97266E','92392E','97268E','97269E','97280E') " & vbCrLf & _
"WHERE (dt.ityp='1' OR (s.voret>qanf() AND dt.ityp IN ('1','2','g'))) " & vbCrLf & _
"AND e.ZeitPunkt BETWEEN qanf() AND qend() " & vbCrLf & _
"AND e.Art IN (" & artSpezBerat & ") " & vbCrLf & _
"AND ISNULL(l.leistung); "
sql(AWlf) = "SELECT f.pat_id " & vbCrLf & _
", IF (ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY e.zeitpunkt)=1,CONCAT(gesname(f.pat_id),' (',patalter(f.pat_id),')'),'') PName " & vbCrLf & _
", IF (ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY e.zeitpunkt)=1,CONCAT(dt.ityp,' ',COALESCE(IF(s.voret=18991230,'',DATE_FORMAT(s.voret,'(%e.%c.%y)')),'')),'') dtyp " & vbCrLf & _
", IF (ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY e.zeitpunkt)=1,REPLACE(maxtha(f.pat_id),_utf8mb4'GLP',_utf8mb4''),'') maxtha " & vbCrLf & _
", DATE_FORMAT(e.zeitpunkt,'%d.%m.%y %H:%i') zp, e.art, e.inhalt" & vbCrLf & _
", (SELECT COUNT(0) FROM leistungen l LEFT JOIN genehmigungen g USING (leistung) WHERE l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qbeg(e.zeitpunkt) - INTERVAL 1 YEAR AND e.zeitpunkt AND g.obschulung=1 AND l.leistung NOT IN ('92282','92278') AND g.Erklärung NOT RLIKE 'Sachkosten') gsz" & vbCrLf & _
", COALESCE((SELECT CONCAT(COUNT(0),':',GROUP_CONCAT(DISTINCT leistung)) FROM leistungen l LEFT JOIN genehmigungen g USING (leistung) WHERE l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qbeg(e.zeitpunkt) - INTERVAL 1 YEAR AND e.zeitpunkt AND g.obschulung=2 AND l.leistung NOT IN ('92282','92278') AND g.Erklärung NOT RLIKE 'Sachkosten'),'') `bish.ES`" & vbCrLf & _
", IF (ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY e.zeitpunkt)=1, " & vbCrLf & _
"   (SELECT COALESCE(GROUP_CONCAT(DISTINCT leistung SEPARATOR ' '),'') FROM genehmigungen WHERE obschulung=2 " & vbCrLf & _
"      AND patalter(f.pat_id) BETWEEN von AND bis " & vbCrLf & _
"      AND (kht<>2 OR ISNULL((SELECT MAX(0) FROM diagview WHERE pat_id=f.pat_id AND gICD='I10.90' AND obdauer<>0))) " & vbCrLf & _
"      AND INSTR(kassen,f.kateg)<>0 " & vbCrLf & _
"     AND (INSTR(erklärung,'DMP')=0 OR ((n.dmpklass=2 OR (n.dmpklass=3 AND n.dmpbeg<=qend())))) " & vbCrLf & _
"     AND dtyp=dt.ityp " & vbCrLf & _
"     AND (therarten='' OR INSTR(therarten,REPLACE(maxtha(f.pat_id),_utf8mb4'GLP',_utf8mb4''))<>0) " & vbCrLf & _
"     AND kht=2 " & vbCrLf & _
"     ),'') mögl_HT" & vbCrLf
sql(AWlf) = sql(AWlf) & _
", IF (ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY e.zeitpunkt)=1, " & vbCrLf & _
"   (SELECT COALESCE(GROUP_CONCAT(DISTINCT leistung SEPARATOR ' '),'') FROM genehmigungen WHERE obschulung=2 " & vbCrLf & _
"      AND patalter(f.pat_id) BETWEEN von AND bis " & vbCrLf & _
"      AND (kht<>2 OR ISNULL((SELECT MAX(0) FROM diagview WHERE pat_id=f.pat_id AND gICD='I10.90' AND obdauer<>0))) " & vbCrLf & _
"      AND INSTR(kassen,f.kateg)<>0 " & vbCrLf & _
"     AND (INSTR(erklärung,'DMP')=0 OR ((n.dmpklass=2 OR (n.dmpklass=3 AND n.dmpbeg<=qend())))) " & vbCrLf & _
"     AND dtyp=dt.ityp " & vbCrLf & _
"     AND (therarten='' OR INSTR(therarten,REPLACE(maxtha(f.pat_id),_utf8mb4'GLP',_utf8mb4''))<>0) " & vbCrLf & _
"     AND kht=1 " & vbCrLf & _
"     ),'') mögl_DM " & vbCrLf & _
"FROM aktfvmta f" & vbCrLf & _
"LEFT JOIN faelle fa USING (fid)" & vbCrLf & _
"LEFT JOIN namen n ON f.pat_id=n.pat_id" & vbCrLf & _
"LEFT JOIN sws s ON s.pat_id=f.pat_id AND s.voret>qanf()" & vbCrLf & _
"LEFT JOIN dtypen dt ON dt.pat_id=f.pat_id" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id" & vbCrLf & _
"LEFT JOIN leistungen l" & vbCrLf & _
" ON l.pat_id=f.pat_id AND DATE(l.zeitpunkt)=DATE(e.zeitpunkt) AND leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung<>0)" & vbCrLf & _
"WHERE (dt.ityp='1' OR (s.voret>qanf() AND dt.ityp IN ('1','2','g')))" & vbCrLf & _
" AND e.ZeitPunkt BETWEEN qanf() AND qend()" & vbCrLf & _
" AND e.Art IN (" & artSpezBerat & ")" & vbCrLf & _
" AND ISNULL(l.leistung)" & vbCrLf & _
"GROUP BY f.pat_id,e.id" & vbCrLf & _
"HAVING gsz<4" & vbCrLf
sql(AWlf) = sql(AWlf) & _
"ORDER BY f.pat_id,e.zeitpunkt"

 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

 ' 99
 AwN(AWlf) = "Potential für Einzelschulungen bei Schwangerschaft (vorher 82 und 131)"
 sql(AWlf) = "SELECT f.Pat_ID,gesname(f.pat_id) PName" & _
 ", COALESCE((SELECT dmtyp FROM dtypview WHERE pat_id=f.pat_id),'-') dtyp" & vbCrLf & _
 ", maxtha(f.pat_id) maxTha" & vbCrLf & _
 ", COALESCE((SELECT voret FROM sws WHERE pat_id=f.pat_id AND voret>qanf() ORDER BY voret DESC LIMIT 1),'') VorET" & vbCrLf & _
 ", IF((SELECT obneu FROM anb_neuman WHERE pat_id=f.pat_id)<>0,'ja','') obneu" & vbCrLf & _
 ", f.KoZ, f.czp KontaktZp,f.Art" & vbCrLf & _
 ", SUM(l.lzahl) ZdBSL" & vbCrLf & _
 ",COALESCE(GROUP_CONCAT(CONCAT(l.leistung,DATE_FORMAT(l.zeitpunkt,'•%e.%c.')) ORDER BY zeitpunkt SEPARATOR ' '),'') Bish_Schul_Lei" & vbCrLf & _
 ",COALESCE((SELECT GROUP_CONCAT(CONCAT(leistung,DATE_FORMAT(zeitpunkt,'•%e.%c.')) ORDER BY zeitpunkt SEPARATOR ' ')" & vbCrLf & _
 "   FROM leistungen WHERE pat_id=f.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND leistung IN ('97268','97274','97271')" & vbCrLf & _
 "      ),'') Lei_97268_97274_97271" & vbCrLf & _
 ",(SELECT GROUP_CONCAT(leistung) FROM leistungen WHERE pat_id=f.pat_id AND ZeitPunkt BETWEEN qanf() AND qend() " & vbCrLf & _
 "    AND leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung=2 AND erklärung NOT RLIKE 'sach|material')) EinzLeist_aktQ " & vbCrLf & _
 ",(SELECT GROUP_CONCAT(leistung) FROM leistungen WHERE pat_id=f.pat_id AND ZeitPunkt BETWEEN qanf()-INTERVAL 3 MONTH AND qend()-INTERVAL 3 MONTH " & vbCrLf & _
 "    AND leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung=2 AND erklärung NOT RLIKE 'sach|material')) EinzLeist_vorQ " & vbCrLf & _
 "FROM aktfpkvs f" & vbCrLf & _
 "LEFT JOIN leistungen l ON l.Pat_ID=f.pat_id " & vbCrLf & _
 "   AND l.leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung<>0 AND erklärung NOT RLIKE 'sach|material') AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
 "GROUP BY f.pat_id   " & vbCrLf & _
 "HAVING " & vbCrLf & _
 "(((VorET<>'' AND dtyp IN ('1','2')) OR (obneu AND dtyp IN ('1','2')) OR (dtyp='g' AND koz>2)) AND GROUP_CONCAT(DISTINCT DATE_FORMAT(l.zeitpunkt,'%e.%c.') ORDER BY l.zeitpunkt SEPARATOR '•')<>czp) OR" & vbCrLf & _
 " (voret<>'' AND dtyp='g' AND koz BETWEEN 1 AND 3)" & vbCrLf & _
 ";"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

#If vorherig Then
#If True Then
' aktfpkvs
AwN(AWlf) = "Potential für Einzel- anstatt Gruppenschulungen ('97268','97274','97271') bei Gestationsdiabetes mit 1-3 Kontakten (vorher 82)"
sql(AWlf) = "SELECT " & vbCrLf & _
"  IF(ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY l.zeitpunkt)=1,maxtha(f.pat_id),'') Tha" & vbCrLf & _
", IF(ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY l.zeitpunkt)=1,dt.ityp,'') DTyp" & vbCrLf & _
", f.pat_id " & vbCrLf & _
", IF(ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY l.zeitpunkt)=1,gesname(f.pat_id),'') PName " & vbCrLf & _
", COALESCE(l.ZeitPunkt,'') Zeitpunkt,COALESCE(l.Leistung,'') Leistung " & vbCrLf & _
",(SELECT GROUP_CONCAT(leistung) FROM leistungen WHERE pat_id=f.pat_id AND ZeitPunkt BETWEEN qanf() AND qend() AND leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung=2)) EinzLeist_aktQ " & vbCrLf & _
",(SELECT GROUP_CONCAT(leistung) FROM leistungen WHERE pat_id=f.pat_id AND ZeitPunkt BETWEEN qanf()-INTERVAL 3 MONTH AND qend()-INTERVAL 3 MONTH AND leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung=2)) EinzLeist_vorQ " & vbCrLf & _
",f.koz, f.czp, f.Art" & vbCrLf & _
" FROM aktfpkvs f " & vbCrLf & _
"LEFT JOIN dtypen dt ON dt.pat_id=f.pat_id " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.ZeitPunkt BETWEEN qanf() AND qend() AND l.leistung IN ('97268','97274','97271') " & vbCrLf & _
"LEFT JOIN sws s ON s.pat_id=f.pat_id AND s.voret>qanf() " & vbCrLf & _
"WHERE NOT ISNULL(s.voret) AND NOT (ityp='-' AND ISNULL(leistung)) AND dt.ityp='g' AND f.koz BETWEEN 1 AND 3" & vbCrLf & _
"ORDER BY f.pat_id,l.zeitpunkt;"
#Else
AwN(AWlf) = "Potential für Einzel- anstatt Gruppenschulungen ('97268','97274','97271') bei Schwangeren"
sql(AWlf) = "SELECT " & vbCrLf & _
"  IF(ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY l.zeitpunkt)=1,maxtha(f.pat_id),'') Tha" & vbCrLf & _
", IF(ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY l.zeitpunkt)=1,dt.ityp,'') DTyp" & vbCrLf & _
", IF(ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY l.zeitpunkt)=1,dt.ityp,'') voret" & vbCrLf & _
", f.pat_id " & vbCrLf & _
", IF(ROW_NUMBER() OVER (PARTITION BY f.pat_id ORDER BY l.zeitpunkt)=1,gesname(f.pat_id),'') PName " & vbCrLf & _
", COALESCE(l.ZeitPunkt,'') Zeitpunkt,COALESCE(l.Leistung,'') Leistung " & vbCrLf & _
",(SELECT GROUP_CONCAT(leistung) FROM leistungen WHERE pat_id=f.pat_id AND ZeitPunkt BETWEEN qanf() AND qend() AND leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung=2)) EinzLeist_aktQ " & vbCrLf & _
",(SELECT GROUP_CONCAT(leistung) FROM leistungen WHERE pat_id=f.pat_id AND ZeitPunkt BETWEEN qanf()-INTERVAL 3 MONTH AND qend()-INTERVAL 3 MONTH AND leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung=2)) EinzLeist_vorQ " & vbCrLf & _
",czp Kontakte" & vbCrLf & _
" FROM aktfkvs f " & vbCrLf & _
"LEFT JOIN dtypen dt ON dt.pat_id=f.pat_id " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.ZeitPunkt BETWEEN qanf() AND qend() AND l.leistung IN ('97268','97274','97271') " & vbCrLf & _
"LEFT JOIN sws s ON s.pat_id=f.pat_id AND s.voret>qanf() " & vbCrLf & _
"WHERE NOT ISNULL(s.voret) AND NOT (ityp='-' AND ISNULL(leistung)) " & vbCrLf & _
"ORDER BY f.pat_id,l.zeitpunkt;"
#End If ' true
' 82 ' in 33 enthalten
' AwN(AWlf) = "" ' "Doppelte 03220/ 03220H"
' sql(AWlf) = "" ' "SELECT * FROM (" & vbCrLf & _
"       SELECT f.pat_id, gesname(f.pat_id) Name, SUM(l0.lzahl) Zahl," & vbCrLf & _
"       GROUP_CONCAT(l0.leistung) Leistung0" & vbCrLf & _
"       FROM aktfvs f " & vbCrLf & _
"       LEFT JOIN leistungen l0 ON f.pat_id = l0.pat_id AND l0.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"       AND l0.leistung IN ('03220','03220H')" & vbCrLf & _
"       GROUP BY f.fid" & vbCrLf & _
"       ) i WHERE Zahl>1"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 131
' ktag fehlerhaft
' AwN(AWlf) = "Mögliche Telefonkontakte (lauto)"
'     sql(AWlf) = _
    " SELECT pat_id PID, gesname(pat_id) NAME, DATE(zp) LEIDAT,'01434 dazu' LEIFEHLER, rang,zahl,koz,zzahl, flz, Art, czp, Typen, Arten, LANRID FROM ( " & vbCrLf & _
    "  SELECT zp,pat_id,rang,zahl,trang,tzahl,koz,COUNT(*) OVER (PARTITION BY pat_id) zzahl, lz " & vbCrLf & _
    "  ,SUM(lz) OVER (PARTITION BY pat_id,CAST(zp As Date)) tlz " & vbCrLf & _
    "  ,SUM(lz) OVER (PARTITION BY pat_id) flz " & vbCrLf & _
    "  ,GROUP_CONCAT(Art ORDER BY Art) Art,czp,Typen,Arten,LANRID FROM ( " & vbCrLf & _
    "   SELECT e.zeitpunkt zp " & vbCrLf & _
    "   ,RANK() OVER (PARTITION BY f.pat_id ORDER BY e.zeitpunkt) rang " & vbCrLf & _
    "   ,COUNT(*) OVER (PARTITION BY f.pat_id) zahl " & vbCrLf & _
    "   ,RANK() OVER (PARTITION BY f.pat_id,CAST(e.zeitpunkt As Date) ORDER BY e.zeitpunkt) trang " & vbCrLf & _
    "   ,COUNT(*) OVER (PARTITION BY f.pat_id,CAST(e.zeitpunkt As Date)) tzahl " & vbCrLf & _
    "   ,f.pat_id, f.fid,e.Art,koz,czp,Typen,Arten,f.lanrid " & vbCrLf & _
    "   ,(SELECT COUNT(0) FROM leistungen l WHERE l.fid = f.fid AND leistung LIKE '01434' AND DATE(l.zeitpunkt)=DATE(zp)) lz " & vbCrLf & _
    "   FROM (SELECT pat_id, fid, schgr,ik,vknr,LANRID,koz,czp,typ Typen,art Arten FROM aktfkvs) f " & vbCrLf & _
    "   JOIN eintraege e ON e.pat_id = f.pat_id " & vbCrLf
'    sql(AWlf) = sql(AWlf) & _
    "    AND e.zeitpunkt " & DiesQ() & vbCrLf & _
    "    AND art IN (" & artSpezBerat & ",'ni','gstel','tel','fa','fam','familie','beruf','uzu','uz','hypo','bzm','bztp','ufrag','uabfrag') " & vbCrLf & _
    "   WHERE CAST(e.zeitpunkt As Date)>20201101 " & vbCrLf & _
    "   GROUP BY pat_id,DATE(e.zeitpunkt) " & vbCrLf & _
    "  ) i WHERE koz>2+zahl-rang GROUP BY pat_id,zp" & vbCrLf & _
    " ) i WHERE zzahl-(zahl-rang)<7-flz AND tlz=0" ' AND PAT_ID IN (51612)" ' AND pat_id=2047" ' AND pat_id = 3159"

'    AwN(AWlf) = "-"
'    sql(AWlf) = "-"

AwN(AWlf) = "Übersicht Schwangere mit T1Dm oder T2Dm oder, falls >=3 Kontakte, GDM, sowie Neumanifestation Typ 1-Diabetes, mit Hinblick auf Einzelschulungen (vorher 131)"
sql(AWlf) = _
"SELECT f.pat_id, gesname(f.pat_id) PName,COALESCE(sws.voret,'') VorET" & vbCrLf & _
",IF(n.obneu=0,' ',n.obneu) obneu, dmtyp,COUNT(DISTINCT DATE(l.zeitpunkt)) lz" & vbCrLf & _
",GROUP_CONCAT(DISTINCT DATE_FORMAT(l.zeitpunkt,'%e.%c.') ORDER BY l.zeitpunkt SEPARATOR '•') lzp" & vbCrLf & _
",f.koz ,f.czp kozp,f.Art" & vbCrLf & _
"FROM aktfpkvs f" & vbCrLf & _
"LEFT JOIN leistungen l ON l.Pat_ID=f.pat_id AND l.leistung IN (SELECT leistung FROM genehmigungen WHERE obschulung<>0) AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"LEFT JOIN sws on sws.pat_id=f.pat_id AND sws.voret>qanf()" & vbCrLf & _
"LEFT JOIN anb_neuman n ON n.pat_id=f.pat_id" & vbCrLf & _
"LEFT JOIN (SELECT pat_id,IF(icd RLIKE '^E',MID(icd,3,1)+1,'g') dmtyp FROM diagview" & vbCrLf & _
"           WHERE (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))) GROUP BY pat_id) d" & vbCrLf & _
"           ON d.pat_id=f.pat_id" & vbCrLf & _
"GROUP BY f.pat_id" & vbCrLf & _
"HAVING ((VorET<>'' AND dmtyp IN ('1','2')) OR (obneu AND dmtyp='1') OR (dmtyp='g' AND koz>2)) AND lzp<>czp" & vbCrLf & _
";"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
#End If ' vorherig
 
' 100
' ktag fehlerhaft
AwN(AWlf) = "GDM Schulungsübersicht aktuell/gesamt (vorher 107)"
'"(SELECT GROUP_CONCAT(DISTINCT insart) FROM therarten t WHERE t.pat_id=f.pat_id AND t.zp BETWEEN qbeg(SUBDATE(et.voret,274)) AND " & qtEnd(FristS) & ") Insarten, " & vbCrLf & _
""
' "(SELECT IF(MAX(insart)AND NOT ISNULL(MAX(insart)),'x','') FROM therarten t WHERE t.pat_id=f.pat_id AND t.zp BETWEEN qbeg(SUBDATE(et.voret,274)) AND " & qtEnd(FristS) & ") Ins, " & vbCrLf & _
""
' AND COALESCE(d.Dggel,0)=0
sql(AWlf) = "" & _
"SELECT f.pat_id, LEFT(gesname(f.pat_id),20) Nam, lr, voret, MAX(icd) icd," & vbCrLf & _
"COALESCE((SELECT therart FROM therarten WHERE pat_id=f.pat_id AND zp=(SELECT MAX(zp) FROM therarten WHERE pat_id=f.pat_id AND zp<qend())ORDER BY zp DESC LIMIT 1),'Diät') TherArt," & vbCrLf & _
"(SELECT COUNT(DISTINCT DATE(zeitpunkt)) FROM eintraege e WHERE e.pat_id=f.pat_id AND e.art IN ('ni','gstel','gs','rz','ep','bga','tk','APK','wd','ah','ta','tb','tt','tn','wr','jl','ga','ih','cr','tn','be','lf','kb','mn','lt','sk','th') AND zeitpunkt BETWEEN qanf() AND IF(qend()<sws.voret,qend(),sws.voret)) KoZ," & vbCrLf & _
"CONCAT(" & vbCrLf & _
"CAST(COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97276' AND zeitpunkt BETWEEN qanf() AND IF(qend()<sws.voret,qend(),sws.voret)),0) AS char CHARACTER SET utf8),'/'," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97276' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0)) `97276`," & vbCrLf & _
"CONCAT(" & vbCrLf & _
"CAST(COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97271' AND zeitpunkt BETWEEN qanf() AND IF(qend()<sws.voret,qend(),sws.voret)),0) AS char CHARACTER SET utf8),'/'," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97271' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0)) `97271`," & vbCrLf & _
"CONCAT(" & vbCrLf & _
"CAST(COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97271E' AND zeitpunkt BETWEEN qanf() AND IF(qend()<sws.voret,qend(),sws.voret)),0) AS char CHARACTER SET utf8),'/'," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97271E' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0)) `97271E`," & vbCrLf & _
"CONCAT(" & vbCrLf & _
"CAST(COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97268' AND zeitpunkt BETWEEN qanf() AND IF(qend()<sws.voret,qend(),sws.voret)),0) AS char CHARACTER SET utf8),'/'," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97268' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0)) `97268`," & vbCrLf & _
"CONCAT(" & vbCrLf & _
"CAST(COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97368E' AND zeitpunkt BETWEEN qanf() AND IF(qend()<sws.voret,qend(),sws.voret)),0) AS char CHARACTER SET utf8),'/'," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97368E' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0)) `97368E`," & vbCrLf & _
"CONCAT(" & vbCrLf
sql(AWlf) = sql(AWlf) & _
"CAST(COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97274' AND zeitpunkt BETWEEN qanf() AND IF(qend()<sws.voret,qend(),sws.voret)),0) AS char CHARACTER SET utf8),'/'," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97274' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0)) `97274`," & vbCrLf & _
"CONCAT(" & vbCrLf & _
"CAST(COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97274E' AND zeitpunkt BETWEEN qanf() AND IF(qend()<sws.voret,qend(),sws.voret)),0) AS char CHARACTER SET utf8),'/'," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97274E' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0)) `97274E`," & vbCrLf & _
"CONCAT(" & vbCrLf & _
"CAST(COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97267B' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0) AS char CHARACTER SET utf8),'/'," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='97267B' AND zeitpunkt BETWEEN qanf()-INTERVAL 2 YEAR AND IF(qend()<sws.voret,qend(),sws.voret)),0)) `97267B`" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"JOIN sws USING (pat_id)" & vbCrLf & _
"JOIN diagview dd ON dd.pat_id=f.pat_id AND (dd.icd RLIKE '^O24' AND dd.Dggel=0 AND dd.diagsicherheit IN ('G',' ') AND (dd.icd RLIKE '^E1[0-4]' OR dd.diagdatum BETWEEN qanf() AND qend()))" & vbCrLf & _
"  AND dd.id1=(SELECT MAX(id1) FROM diagview WHERE pat_id=f.pat_id AND (icd RLIKE '^O24' AND Dggel=0 AND diagsicherheit IN ('G',' ') AND (icd RLIKE '^E1[0-4]' OR diagdatum BETWEEN qanf() AND qend())))" & vbCrLf & _
"WHERE lr+INTERVAL 280 DAY>=qanf()" & vbCrLf & _
"AND dd.icd IS NOT NULL" & vbCrLf & _
"GROUP BY f.pat_id" & vbCrLf & _
"ORDER BY (SELECT IF(MAX(insart)AND NOT ISNULL(MAX(insart)),'x','') FROM therarten t WHERE t.pat_id=f.pat_id AND t.zp BETWEEN qbeg(sws.lR) AND qend()), therart, f.pat_id" & vbCrLf

' "(SELECT IF(instr(GROUP_CONCAT(DISTINCT insart),'1')<>0 OR instr(GROUP_CONCAT(DISTINCT insart),'4')<>0,'ICT',IF(instr(GROUP_CONCAT(DISTINCT insart),'2')<>0,'Basal','Diät'))  FROM therarten t WHERE t.pat_id=f.pat_id AND t.zp BETWEEN qbeg(STR_TO_DATE(et.letzteRegel,'%d.%m.%Y')) AND " & qtEnd(FristS) & ") TherArt, " & vbCrLf & _
''"LEFT JOIN faelle et ON f.pat_id=et.pat_id AND et.voret BETWEEN " & KhtsflZuk() & vbCrLf & _
mins(AWlf) = 10
maxs(AWlf) = 80
AWlf = AWlf + 1
 
' 101, neuView
AwN(AWlf) = "Gestationsdiabetes, Diabetes und Schwangerschaft"
sql(AWlf) = "ü"
 AWlf = AWlf + 1

' 102
' ktag fehlerhaft
 AwN(AWlf) = "Eventuell falsch abgerechnete OGTT in der Schwangerschaft (01777: 1/Schwt, 01812: 3/OGTT) (vorher 60)"
' sql(AWlf) = "SELECT `01777`, `01812`, (SELECT SUM(IF(LAnzl='',1,LAnzl)) FROM leistungen l WHERE l.pat_id= pat_id AND DATE(l.zeitpunkt)<messzeitpunkt AND DATE(l.zeitpunkt) > `letzte Regel` AND l.leistung='01777') Vor01777, Pat_ID, Name, `letzte Regel`, Messzeitpunkt, `OGTT-Dokumentation` FROM " & vbCrLf & _
'    "(SELECT SUM(IF(ogtt.LAnzl='',1,ogtt.LAnzl))  `01777`, SUM(IF(gluc.LAnzl='',1,gluc.LAnzl)) `01812`, e.pat_id Pat_ID, LEFT(CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')),IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')),n.nachname,', ',n.vorname),25) Name, STR_TO_DATE(f.letzteRegel,'TM#%d%m%Y') `letzte Regel`, DATE(e.zeitpunkt) Messzeitpunkt, e.inhalt `OGTT-Dokumentation`, e.fid fid " & vbCrLf & _
'    "FROM eintraege e " & vbCrLf & _
'    "LEFT JOIN faelle f ON e.fid = f.fid " & vbCrLf & _
'    "LEFT JOIN leistungen ogtt ON f.fid = ogtt.fid AND ogtt.leistung = '01777' AND DATE(ogtt.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
'    "LEFT JOIN leistungen gluc ON f.fid = gluc.fid AND gluc.leistung = '01812' AND DATE(gluc.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
'    "LEFT JOIN namen n ON e.pat_id = n.pat_id " & vbCrLf & _
'    "WHERE e.art = 'ogtt' AND e.zeitpunkt BETWEEN " & lqanfuend(FristS) & " AND letzteRegel<>'' AND ADDDATE(STR_TO_DATE(f.letzteRegel,'TM#%d%m%Y'),274) > e.zeitpunkt GROUP BY e.fid,e.zeitpunkt) i " & vbCrLf & _
'    "WHERE `01777`<>(IF((SELECT SUM(IF(LAnzl='',1,LAnzl)) FROM leistungen l WHERE l.pat_id= pat_id AND DATE(l.zeitpunkt)<messzeitpunkt AND DATE(l.zeitpunkt) > `letzte Regel` AND l.leistung='01777')=0,1,0)) OR `01812` <>3 "
sql(AWlf) = "SELECT Pat_ID, Name, Messzeitpunkt, `01812`,`Vor-01812`,Soll, `01777`, `Vor-01777`" & vbCrLf & _
", ob50, letzteRegel, CONCAT(DATE_FORMAT(voret - INTERVAL 280 DAY,'%d.%m.%y'),'-',DATE_FORMAT(voret,'%d.%m.%y')) `SWS-von-bis`" & vbCrLf & _
", einh `OGTT-Dokumentation` " & vbCrLf & _
"FROM (" & vbCrLf & _
"SELECT COALESCE(SUM(ogtt.lzahl),0) `01777`" & vbCrLf & _
",COALESCE((SELECT MAX(IF(inhalt RLIKE 'ja am|t *ja|am *[0-9]' OR inhalt RLIKE 'chgeführt\?.\{0,2\}ja',1,IF(inhalt RLIKE 'nein am|- am|-,' OR inhalt RLIKE 'chgeführt\?.\{0,2\}nein',0,'?'))) FROM eintraege WHERE pat_id = f.pat_id AND art RLIKE '^angd|^50g$' AND DATE(zeitpunkt)BETWEEN qbeg(e.zeitpunkt)AND DATE(e.zeitpunkt)),'u') ob50" & vbCrLf & _
",COALESCE(SUM(gluc.lzahl),0) `01812`" & vbCrLf & _
",COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id= e.pat_id AND DATE(zeitpunkt)<DATE(e.zeitpunkt) AND DATE(zeitpunkt)> et.letzteRegel AND leistung='01812'),0) `Vor-01812`" & vbCrLf & _
",et.letzteRegel" & vbCrLf & _
",gesname(f.pat_id) Name, DATE(e.zeitpunkt) Messzeitpunkt, e.inhalt einh, e.fid fid, e.pat_id pat_id" & vbCrLf & _
",COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id= e.pat_id AND DATE(zeitpunkt)<DATE(e.zeitpunkt) AND DATE(zeitpunkt)> et.letzteRegel AND leistung='01777'),0) `Vor-01777`" & vbCrLf & _
",et.voret, (SELECT COUNT(0) FROM eintraege WHERE pat_id=e.pat_id AND DATE(zeitpunkt)=DATE(e.zeitpunkt) AND art IN ('bzvgl','bz')) + CASE WHEN e.inhalt LIKE '%mg%mg%mg%' AND NOT e.inhalt LIKE '%mg%mg%~%mg%' THEN 3 WHEN e.inhalt LIKE '%mg%mg%' AND NOT e.inhalt LIKE '%mg%~%mg%' THEN 2 WHEN e.inhalt LIKE '%mg%' THEN 1 ELSE 0 END Soll" & vbCrLf & _
",f.kid" & vbCrLf & _
"FROM aktfv v JOIN faelle f USING (fid) LEFT JOIN namen n ON n.pat_id=v.pat_id" & vbCrLf & _
"LEFT JOIN eintraege e ON e.Pat_ID=v.pat_id AND e.art='ogtt' AND e.zeitpunkt BETWEEN qanf()AND qend()" & vbCrLf & _
"LEFT JOIN (SELECT IF(LR=18991230,IF(efLR=18991230,IF(erLR=18991230,IF(voret<19500101,voret+INTERVAL 100 YEAR,voret)-INTERVAL 280 day,erlr),efLR),IF(LR<19500101,LR+INTERVAL 100 YEAR,LR)) letzteRegel, voret,pat_id FROM sws)et ON et.Pat_ID=f.pat_id AND et.voret>qanf() AND et.voret-INTERVAL 280 DAY<e.zeitpunkt" & vbCrLf & _
"LEFT JOIN leistungen ogtt ON f.pat_id = ogtt.pat_id AND ogtt.leistung = '01777' AND DATE(ogtt.zeitpunkt) = DATE(e.zeitpunkt)" & vbCrLf & _
"LEFT JOIN leistungen gluc ON f.pat_id = gluc.pat_id AND gluc.leistung = '01812' AND DATE(gluc.zeitpunkt) = DATE(e.zeitpunkt)" & vbCrLf & _
"WHERE EXISTS (SELECT 0 FROM sws WHERE pat_id=f.pat_id AND e.zeitpunkt BETWEEN voret - INTERVAL 280 day AND voret)" & vbCrLf & _
"GROUP BY e.pat_id,e.zeitpunkt)i" & vbCrLf & _
"WHERE ((SELECT MAX(kateg) FROM kassenliste WHERE id=i.kid)<>'SHV' AND" & vbCrLf & _
"        `01777`<>(IF(`Vor-01777`=0 AND ob50 IN(1,'u'),1,0)))" & vbCrLf & _
"   OR `01812` < LEAST(Soll,8-`Vor-01812`)" & vbCrLf & _
";"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
'       "letzteRegel<>'' AND ADDDATE(STR_TO_DATE(f.letzteRegel,'TM#%d%m%Y'),274) > e.zeitpunkt " & vbCrLf & _
'             ",COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id= e.pat_id AND DATE(zeitpunkt)<DATE(e.zeitpunkt) AND DATE(zeitpunkt)> STR_TO_DATE(f.letzteRegel,'%d.%m.%Y') AND leistung='01777'),0) `Vor-01777` " & vbCrLf & _

 ' 103
AwN(AWlf) = "01777 ohne dokumentierten, auswärtigen 50g-Vortest (vorher 103)"
sql(AWlf) = vbCrLf & _
"SELECT l.pat_id, gesname(l.pat_id), l.zeitpunkt,LfBegr " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN leistungen l ON f.pat_id = l.pat_id AND leistung = 01777 AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"WHERE NOT ISNULL(leistung) AND LfBegr NOT RLIKE '50g(r.?)?( (OGTT|Test|oraler Glucosetoleranztest))? bei(m)? (Gyn|Frauenarzt)( bereits)? erfolgt'"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 104
AwN(AWlf) = "Mehrfache 01777 (vorher 88)"
sql(AWlf) = "" & vbCrLf & _
"SELECT * FROM ( " & vbCrLf & _
"SELECT f.pat_id,gesnameg(f.pat_id) Name, TRIM(GROUP_CONCAT(' ', DATE_FORMAT(l.zeitpunkt,'%d.%m.%y'))) Zeitpunkte, " & vbCrLf & _
"COALESCE(SUM(l.lzahl),0) Zahl " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id = f.pat_id " & vbCrLf & _
"WHERE leistung = '01777' " & vbCrLf & _
"AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"GROUP BY f.pat_id) i WHERE Zahl>1; "
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 105
 AwN(AWlf) = "Mehr als 8 Leistungen 01812 im Krankheitsfall (vorher 114)"
 sql(AWlf) = vbCrLf & _
"SELECT * FROM (SELECT " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.zeitpunkt BETWEEN " & Khtsfl & " AND l.leistung='01812'),0) Geszahl, " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='01812' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & "),0) Aktzahl, " & vbCrLf & _
"gesname(f.pat_id) Gesname, f.pat_id " & vbCrLf & _
", (SELECT GROUP_CONCAT(zeitpunkt SEPARATOR ', ') FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='01812' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Zpkte " & vbCrLf & _
"FROM aktfvs f) i WHERE geszahl>8 AND aktzahl>0;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
 ' 106
 ' ktag fehlerhaft
 AwN(AWlf) = "Falsche Zahl 01812 oder 32025 für Blutzuckermessungen (bz, bzvgl, ogtt) ín der Schwangerschaft (vorher 59)"
 sql(AWlf) = "SELECT * FROM (" & vbCrLf & _
 "SELECT i.pid, gesname(i.pid) Name, i.eTag Tag, i.BZZahl `BZ-Zahl`" & vbCrLf & _
 ",COALESCE(SUM(lzahl),0) `Leistungs-Zahl`" & vbCrLf & _
 ",COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id= pid AND DATE(zeitpunkt)<=DATE(eTag) AND DATE(zeitpunkt)>et.letzteRegel AND leistung='01812'),0) `Vor-01812`" & vbCrLf & _
 ", i.Art, i.Inhalt" & vbCrLf & _
 "FROM (SELECT f.pat_id pid, DATE(e.zeitpunkt) eTag, SUM(CASE WHEN art LIKE 'bz%' THEN 1 WHEN art= 'angd' THEN CASE WHEN inhalt RLIKE 'BZ0.*[0-9][ ]*mg.*[0-9][ ]*mg.*[0-9][ ]*mg.*Grenze' THEN 3 when inhalt RLIKE 'BZ0.*[0-9][ ]*mg.*[0-9][ ]*mg.*Grenze' THEN 2 WHEN inhalt LIKE 'BZ0.*[0-9][ ]*mg.*Grenze' THEN 1 ELSE 0 END ELSE CASE WHEN inhalt RLIKE '[0-9][ ]*mg.*[0-9][ ]*mg.*[0-9][ ]*mg' THEN 3 when inhalt RLIKE '[0-9][ ]*mg.*[0-9][ ]*mg' THEN 2 WHEN inhalt LIKE '[0-9][ ]*mg' THEN 1 ELSE 0 END END) BZZahl, e.Art, e.inhalt FROM " & vbCrLf & _
 "aktfvs f " & vbCrLf & _
 "LEFT JOIN `faelle` fl USING(fid) " & vbCrLf & _
 "LEFT JOIN BiosenMessung e ON f.pat_id = e.pat_id AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & vbCrLf & _
 "WHERE NOT ISNULL(art) AND EXISTS (SELECT 0 FROM sws WHERE pat_id=f.pat_id AND e.zeitpunkt BETWEEN voret - INTERVAL 280 DAY AND voret) " & vbCrLf & _
 "GROUP BY f.pat_id, DATE(e.zeitpunkt)) i " & vbCrLf & _
 "LEFT JOIN (SELECT IF(LR=18991230,IF(efLR=18991230,IF(erLR=18991230,IF(voret<19500101,voret+INTERVAL 100 YEAR,voret)-INTERVAL 280 day,erlr),efLR),IF(LR<19500101,LR+INTERVAL 100 YEAR,LR)) letzteRegel, voret,pat_id FROM sws) et ON et.Pat_ID=i.pid AND et.voret>qanf() AND et.voret - INTERVAL 280 DAY<eTag" & vbCrLf & _
 "LEFT JOIN leistungen l ON l.pat_id = i.pid AND leistung IN ('01812','32025') AND DATE(l.zeitpunkt) = i.eTag " & vbCrLf & _
 "GROUP BY i.pid, i.eTag) i WHERE `BZ-Zahl`>`Leistungs-Zahl` AND `Vor-01812`<8;"
' AwN(AWlf) = "Fehlende 01812 für Blutzuckermessungen (bz, bzvgl) außerhalb der Schwangerschaft:"
' sql(AWlf) = "SELECT * FROM (" & vbCrLf & _
      "SELECT f.pat_id, LEFT(CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')),IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')),n.nachname,', ',n.vorname),25) Name," & vbCrLf & _
      "DATE(e.zeitpunkt) Tag, COUNT(art) `BZ-Zahl`, SUM(IF(ISNULL(l.LAnzl),0,IF(l.LAnzl='',1,l.LAnzl))) `Leistungs-Zahl`," & vbCrLf & _
      "STR_TO_DATE(fl.letzteRegel,'TM#%d%m%Y') `letzte Regel` " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN faelle fl ON f.fid = fl.fid " & vbCrLf & _
"LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
"LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND art IN ('bz','bzvgl') AND (fl.letzteRegel<>'' AND e.zeitpunkt BETWEEN STR_TO_DATE(fl.letzteRegel,'TM#%d%m%Y') AND ADDDATE(STR_TO_DATE(fl.letzteRegel,'TM#%d%m%Y'),274)) " & vbCrLf & _
"LEFT JOIN leistungen l ON f.fid = l.fid AND leistung = '01812' AND DATE(l.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
"WHERE NOT ISNULL(art) GROUP BY f.pat_id, DATE(e.zeitpunkt)) i WHERE `BZ-Zahl`<>`Leistungs-Zahl`;"
'  "WHERE NOT ISNULL(art) AND EXISTS (SELECT 0 FROM `faelle` WHERE pat_id=f.pat_id AND e.zeitpunkt BETWEEN STR_TO_DATE(letzteRegel,'%d.%m.%Y') AND voret) " & vbCrLf & _
mins(AWlf) = 10
maxs(AWlf) = 60
AWlf = AWlf + 1
 
' 107
' ktag fehlerhaft
AwN(AWlf) = "01812, 01776 oder 01777 ohne Ausnahmekennziffer 32007 (lauto) (vorher 104)"
sql(AWlf) = vbCrLf & _
"SELECT DISTINCT f.pat_id PID, gesname(f.pat_id) Name " & vbCrLf & _
",DATE(l.zeitpunkt) LEIDAT, GROUP_CONCAT(l.leistung) Leistungen, '32007 dazu' LEIFEHLER, f.LANRID " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN leistungen l ON f.pat_id=l.pat_id AND l.leistung IN ('01777','01776','01812') AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"LEFT JOIN leistungen a ON f.pat_id=a.pat_id AND a.leistung = '32007' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"WHERE NOT ISNULL(l.leistung) AND ISNULL(a.leistung)" & vbCrLf & _
"GROUP BY f.fid,DATE(l.zeitpunkt)"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

' 108
 AwN(AWlf) = "Fehlende 97333 für Gestationsdiabetes (vorher 12)"
 ' AND COALESCE(d.Dggel,0)=0
 sql(AWlf) = _
 "SELECT f.pat_id, gesname(f.pat_id) Name,d.icd,f.kateg " & vbCrLf & _
 ", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE leistung='97333' AND pat_id=f.pat_id AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & "),0) `Zahl 97333` " & vbCrLf & _
 "FROM aktfv f " & vbCrLf & _
 "LEFT JOIN `diagview` d ON f.pat_id = d.pat_id AND (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)) " & vbCrLf & _
 "WHERE NOT ISNULL(d.ICD) AND f.kateg NOT IN ('SHV')" & vbCrLf & _
 "HAVING `Zahl 97333`<>2;"
' sql(AWlf) = _
 "SELECT * FROM " & vbCrLf & _
 "(SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, l.leistung AS Leistung, icd, REPLACE(REPLACE(notiz,char(13),''),char(10),'') Notiz, kateg, schgr  FROM " & aktf & " " & vbCrLf & _
 "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
 "LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND d.diagsicherheit IN (' ','G','V') AND COALESCE(d.Dggel,0)=0 AND icd LIKE 'O24.4%' AND diagdatum BETWEEN  " & lQAnfuEnd(FristS) & " " & _
 "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
 "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
 "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
 "WHERE (leistung <> '97333' OR ISNULL(leistung)) " & vbCrLf & _
 "AND kateg IN ('AOK','BKK','BKN','IKK','LKK','vdek') " & vbCrLf & _
 "AND NOT ISNULL(icd) "
' "AND obmednetz = 0 " & vbCrLf & _
' "AND notiz LIKE '%DMP%' " & vbCrLf & _
' "AND NOT notiz LIKE '%DMP%nein%'"
 mins(AWlf) = 7
 maxs(AWlf) = 30
 AWlf = AWlf + 1
 
 ' 109
 ' 97312=Betreuung T1DM i.d.Schwangerschaft, 97323=~T2DM~, 97333=~GDM~
 ' AND COALESCE(ddm.Dggel,0)=0
 AwN(AWlf) = "evtl. falsche oder fehlende Diagnosen oder Leistungen bei evtl. Schwangerschaft (vorher 16)"
 sql(AWlf) = _
 "SELECT * FROM ( SELECT i.Pat_id, i.Name, i.voret, i.`OGTT`,i.`ICD D.m.`,i.`ICD Schw`, " & vbCrLf & _
 "      IF(`ICD Schw`=swdg,'-',swdg) `evtl.fehl.Diag.`, l.leistung, COALESCE(SUM(lzahl),0) LZahl, " & vbCrLf & _
 "      IF(swdg='O24.4' AND kateg<>'SHV' AND (l.leistung<>97333 OR COALESCE(SUM(lzahl),0)<>2),'97333 2-mal',IF(swdg='O24.0' AND l.leistung<>97313,'97313',IF(swdg='O24.1' AND l.leistung<>97323,'97323','-'))) `evtl.fehl.Leistg`, " & vbCrLf & _
 "       i.Inhalt , briefname, i.DkZp, i.FID, geschlecht, quartal,kateg " & vbCrLf & _
 "    FROM ( " & vbCrLf & _
 " SELECT n.pat_id, n.geschlecht, et.voret, gesname(n.pat_id) Name, CONCAT(DATE_FORMAT(ogtt.zeitpunkt,'%e.%c.%y'),':',g0(ogtt.inhalt),'/',g1(ogtt.inhalt),'/',g2(ogtt.inhalt)) `OGTT`, " & vbCrLf & _
 "     ddm.icd `ICD D.m.`, dgd.icd `ICD Schw`, " & vbCrLf & _
 "     IF(ddm.icd LIKE 'E10%','O24.0', IF(ISNULL(ddm.icd),'O24.4','O24.1')) swdg, " & vbCrLf & _
 "     e.inhalt, dk.name Briefname, dk.qdm DkZp, f.fid, f.quartal,f.kateg " & vbCrLf & _
 " FROM aktfv f " & vbCrLf & _
 " LEFT JOIN sws et ON et.pat_id=f.pat_id AND et.voret>qanf()" & vbCrLf & _
 " LEFT JOIN diagview ddm ON f.Pat_id = ddm.pat_id AND ddm.gicd RLIKE '^E1[0-4]\.' " & vbCrLf & _
 " LEFT JOIN diagnosen dgd ON f.pat_id = dgd.pat_id AND(dgd .icd LIKE 'O24%' AND dgd.Dggel=0 AND dgd.diagsicherheit IN ('G',' ') AND dgd.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)) " & vbCrLf & _
 " LEFT JOIN namen n ON f.pat_id = n.pat_id "
 sql(AWlf) = sql(AWlf) & vbCrLf & _
 " LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.inhalt RLIKE 'Wohlempfinden|schwanger' AND NOT e.inhalt RLIKE 'war.*schwanger|schwangerschaft|schwanger werden|nicht schwanger|schwanger war|schwanger zu werden|Gattin schwanger' AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
 " LEFT JOIN eintraege ogtt ON ogtt.pat_id=f.pat_id AND ogtt.art='ogtt' AND ogtt.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=f.pat_id AND art='ogtt') AND ADDDATE(ogtt.zeitpunkt,180)>" & qtAnf(FristS) & " " & _
 " LEFT JOIN tmbrie dk ON f.pat_id = dk.pat_id AND dk.name LIKE '%mutterpa%' AND dk.qdm BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
 " WHERE (NOT ISNULL(dk.name) OR NOT ISNULL(e.inhalt) OR NOT ISNULL(dgd.icd)) " & vbCrLf & _
 "        AND NOT (g0(ogtt.inhalt)<92 AND g1(ogtt.inhalt)<180 AND g2(ogtt.inhalt)<153) " & vbCrLf & _
 "        AND NOT et.voret BETWEEN ADDDATE(" & qtAnf(FristS) & ",-90) AND ADDDATE(" & qtAnf(FristS) & ",-1) " & vbCrLf & _
 "        AND ogtt.zeitpunkt < et.voret" & vbCrLf & _
 " GROUP BY f.pat_id) i " & vbCrLf & _
 " LEFT JOIN `leistungen` l ON l.pat_id=i.pat_id AND quartal(l.zeitpunkt) = i.quartal AND l.leistung IN ('97313','97323','97333')" & vbCrLf & _
 " GROUP BY i.fid) i " & vbCrLf & _
 " WHERE (`evtl.fehl.Diag.`<>'-' OR `evtl.fehl.Leistg`<>'-') " & vbCrLf & _
 " AND NOT EXISTS (SELECT icd FROM diagnosen WHERE pat_id = i.pat_id AND (icd LIKE 'O24%' AND Dggel=0 AND diagsicherheit = 'A' AND diagdatum BETWEEN qbegs(i.quartal) AND qends(i.quartal)))"
 maxs(AWlf) = 26
 AWlf = AWlf + 1
 ' " LEFT JOIN faelle et ON et.pat_id=f.pat_id AND voret<>18991230 AND et.qanf=(SELECT MAX(qanf) FROM faelle WHERE pat_id=f.pat_id AND voret<>18991230) AND et.voret BETWEEN qbeg(SUBDATE(" & qtAnf(FristS) & ",274)) AND qende(ADDDATE(" & qtEnd(FristS) & ",274)) " & vbCrLf & _
' " LEFT JOIN sws et ON et.pat_id=f.pat_id AND et.voret>qanf()" & vbCrLf & _

' 110
' ktag fehlerhaft
 AwN(AWlf) = "Evl. fehlende Beratungsgespraeche in der oder bei geplanter Schwangerschaft (92281 bei T1DM, 92277 bei T2DM) (vorher 100)"
 ' AND COALESCE(dd.Dggel,0)=0
 sql(AWlf) = vbCrLf & _
 "SELECT f.pat_id AS Pat_ID,gesnameG(f.pat_id) Name, dd.ICD DiabICD, ds.icd SchwICD, DATE(ds.diagdatum) SchwICDZp, DATE(e.zeitpunkt) EintrZp, e.art EintrArt, e.inhalt EintrInhalt " & vbCrLf & _
      "FROM `aktf` f " & vbCrLf & _
      "LEFT JOIN `namen` n USING (pat_id) " & vbCrLf & _
      "LEFT JOIN `kassenliste` kl ON kl.id=f.kid" & vbCrLf & _
      "LEFT JOIN `diagnosen` dd ON f.pat_id = dd.pat_id AND dd.icd REGEXP '^E1[0-4]' AND dd.diagsicherheit NOT IN ('A','Z','V') " & vbCrLf & _
      "LEFT JOIN `diagnosen` ds ON f.pat_id = ds.pat_id AND ds.icd LIKE 'O24.%' AND ds.diagdatum BETWEEN " & lQAnfuEnd(FristS) & " AND ds.diagsicherheit NOT IN ('A','Z','V') AND ds.Dggel=0" & vbCrLf & _
      "LEFT JOIN `eintraege` e ON f.pat_id = e.pat_id AND (e.art IN ('andm','andm2'" & artSpezBerat & ")) " & vbCrLf & _
      "AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND e.inhalt RLIKE 'schwang|schwg' " & vbCrLf & _
      "LEFT JOIN leistungen l0 ON l0.pat_id = f.pat_id AND l0.leistung IN ('92281','92277') AND l0.zeitpunkt>SUBDATE(qende(e.zeitpunkt),365) " & vbCrLf & _
      "WHERE NOT ISNULL(dd.id1) AND NOT ISNULL(e.art) AND ISNULL(l0.leistung) AND NOT kl.kateg IN (/*'LKK',*/'PBe','SHV') AND ((NOT ISNULL(ds.id1) AND ds.diagdatum>" & qtAnf(FristS) & "-INTERVAL 6 MONTH)OR NOT e.inhalt IS NULL) " & vbCrLf & _
      "AND n.geschlecht<>'m' AND patalter(n.pat_id) BETWEEN 16 AND 50 GROUP BY f.pat_id, e.art, e.zeitpunkt ORDER BY MID(dd.icd,2,1), pat_id, e.zeitpunkt;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

 
' 111
AwN(AWlf) = "GDM mit falscher Zahl an 03003 (lauto) (vorher 105)"
' AND COALESCE(d.Dggel,0)=0
sql(AWlf) = vbCrLf & _
"SELECT i.*,IF(lz=0,'03003 dazu',CONCAT(lz-1,' 03003 weg')) LEIFEHLER FROM ( " & vbCrLf & _
"SELECT f.pat_id PID, gesname(f.pat_id) Name " & vbCrLf & _
",COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE f.pat_id=l.pat_id AND l.leistung='03003' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & "),0) lz " & vbCrLf & _
", f.LANRID " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN diagview d ON f.pat_id=d.pat_id AND (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))" & vbCrLf & _
"WHERE NOT ISNULL(d.icd)) i WHERE lz<>1"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 112
AwN(AWlf) = "GDM mit verdächtiger Zahl an 97277 (vorher 106)"
'  AND COALESCE(d.Dggel,0)=0
sql(AWlf) = vbCrLf & _
"SELECT f.pat_id, gesname(f.pat_id) Name, s.VorET," & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE l.pat_id=f.pat_id AND l.leistung='97277' AND l.zeitpunkt BETWEEN qbeg(s.voret - INTERVAL 280 DAY) AND qend()),0) `97277-Zahl` " & vbCrLf & _
",(SELECT GROUP_CONCAT(DISTINCT DATE_FORMAT(zeitpunkt,'%e.%c.%y') SEPARATOR ',') FROM leistungen l WHERE l.pat_id=f.pat_id AND l.leistung='97277' AND l.zeitpunkt BETWEEN qbeg(s.voret - INTERVAL 280 DAY) AND qend()) `Zp` " & vbCrLf & _
"FROM aktfvmta f " & vbCrLf & _
"LEFT JOIN sws s ON s.pat_id=f.pat_id AND s.voret>qanf()" & vbCrLf & _
"LEFT JOIN diagview d ON f.pat_id=d.pat_id AND (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)) " & vbCrLf & _
"LEFT JOIN diagview dm ON f.pat_id=dm.pat_id AND dm.gicd RLIKE '^E1[0-4]\.' " & vbCrLf & _
"WHERE NOT ISNULL(d.icd) AND f.kateg<> 'SHV'" & vbCrLf & _
"HAVING `97277-Zahl`<>1"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
'"LEFT JOIN faelle et USING (fid) " & vbCrLf & _
",d.icd,d.diagdatum" & vbCrLf & _
" -- OR (NOT ISNULL(s.voret) AND ISNULL(dm.icd)) " entfernt 14.7.22, da auch Schwangere mit gar keinem Diabetes möglich sind

'AwN(AWlf) = "GDM mit Hypertonie"
''  AND COALESCE(d.Dggel,0)=0
'sql(AWlf) = vbCrLf & _
'"SELECT gesname(f.pat_id) PName, f.pat_id, h.icd, mp.medikament, zdzhB, GROUP_CONCAT(rr) rr" & vbCrLf & _
'"FROM aktfv f" & vbCrLf & _
'"LEFT JOIN diagview d ON f.pat_id = d.pat_id AND (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))" & vbCrLf & _
'"LEFT JOIN diagview h ON f.pat_id = h.pat_id AND (h.icd RLIKE '^i10' AND h.diagsicherheit IN ('G',' '))" & vbCrLf & _
'"LEFT JOIN (SELECT mp.* from medplan mp inner JOIN medarten ma ON mp.medanfang=ma.medikament where tag between qanf() and qend() AND hypt<>0) mp ON mp.pat_id=f.pat_id" & vbCrLf & _
'"LEFT JOIN (select COUNT(0) OVER (PARTITION BY pat_id) zdzhB, pat_id, zeitpunkt, rr from rr WHERE (rr.rrsyst>=140 OR rr.rrdiast>=90)) rr  ON rr.pat_id=f.pat_id" & vbCrLf & _
'"Where Not IsNull(d.ICD)" & vbCrLf & _
'"AND (NOT ISNULL(h.icd) OR NOT isnull(medikament) OR zdzhB>1" & vbCrLf & _
'")" & vbCrLf & _
'"GROUP BY f.pat_id" & vbCrLf & _
'";"
' mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
'

' 113, neuView
AwN(AWlf) = "Weitere Leistungen"
sql(AWlf) = "ü"
 AWlf = AWlf + 1

' 114
' ktag fehlerhaft
 AwN(AWlf) = "Unvorhergesehene Inanspruchnahme 1 (01100) Mo-Fr 19-22, WE 7-19 (vorher 13)"
' sql(AWlf) = "SELECT e.Pat_ID AS pat_id,LEFT(DATE_FORMAT(e.ZeitPunkt,_utf8'%a, %d.%m.%y %H:%i'),20) AS Datum, LEFT(e.Art,5) AS art,e.Inhalt AS inhalt FROM `eintraege` e LEFT JOIN `faelle` f USING (fid) WHERE ((e.ZeitPunkt BETWEEN CONCAT(YEAR((NOW() - INTERVAL 14 DAY)),_latin1'-',((((month((NOW() - INTERVAL 14 DAY)) - 1) DIV 3) * 3) + 1),_latin1'-01') AND (CONCAT(YEAR((NOW() - INTERVAL 14 DAY)),_latin1'-',((((month((NOW() - INTERVAL 14 DAY)) - 1) DIV 3) * 3) + 1),_latin1'-01') + INTERVAL 3 MONTH)) AND (e.Art IN (_latin1'notiz',_latin1'ni',_latin1'telef',_latin1'gs',_latin1'ep',_latin1'wr',_latin1'bga',_latin1'gstel',_latin1'rz',_latin1'ga'))) AND (((time(zeitpunkt) BETWEEN '19.00' AND '22:00') AND weekday(zeitpunkt) BETWEEN 0 AND 5) OR ((weekday(zeitpunkt) BETWEEN 6 AND 7) AND (time(zeitpunkt) BETWEEN '07:00' AND '19:00'))) AND schgr <> '90' AND NOT goäkatnr IN (40,41) ORDER BY e.Pat_ID"
 sql(AWlf) = "SELECT e.Pat_id, gesname(e.pat_id)PName,LEFT(DATE_FORMAT(e.zeitpunkt,'%a, %d.%m.%y %H:%i'),20)Datum,CONCAT(RPAD(LEFT(e.Art,5),6),' ',RPAD(e.Ersteller,4),' ',e.`Änderer`)`Art/Erstel./Änd.`,e.Inhalt " & vbCrLf & _
             "FROM aktfv f LEFT JOIN eintraege e ON e.pat_id = f.pat_id AND e.art NOT IN ('vac','cia','c19i') AND e.inhalt NOT RLIKE 'impfung|armschmerzen|gefaxt|Rezepte erstellt|Urogenitalflora' " & vbCrLf & _
             "LEFT JOIN `leistungen` l ON l.pat_id = e.pat_id AND leistung IN ('01100','01101') AND DATE(l.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
             "WHERE e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND (e.Art IN (" & artspezG & ")) " & vbCrLf & _
             "AND ((WEEKDAY(e.zeitpunkt)=0 AND TIME(e.zeitpunkt) BETWEEN '19:30' AND '22:00') OR (WEEKDAY(e.zeitpunkt) IN (1,3) AND time(e.zeitpunkt) BETWEEN '20:00' AND '22:00') OR (weekday(e.zeitpunkt) IN (2,4) AND time(e.zeitpunkt) BETWEEN '19:00' AND '22:00') OR (weekday(e.zeitpunkt) IN (5,6) AND time(e.zeitpunkt) BETWEEN '07:00' AND '19:00')) AND schgr <> '90' AND NOT goäkatnr IN ('40','41') " & vbCrLf & _
             "AND e.inhalt NOT LIKE '%nachgef%' AND e.inhalt NOT LIKE '=>%' " & vbCrLf & _
             "AND NOT EXISTS (SELECT 0 FROM leistungen WHERE pat_id=e.pat_id AND leistung IN (01101,01102,01205,01207,01210,01212,01214,01216,01218,01410,01411,01412,01413,01415,01418,01949,01950,01951,01953,03373,04373,37306) AND zeitpunkt BETWEEN e.zeitpunkt AND e.zeitpunkt + INTERVAL 1 hour)" & vbCrLf & _
             "AND NOT (WEEKDAY(e.zeitpunkt)=5 AND EXISTS (SELECT 0 FROM leistungen WHERE pat_id=e.pat_id AND DATE(zeitpunkt)=DATE(e.zeitpunkt) AND leistung='01102'))" & vbCrLf & _
             "AND ISNULL(leistung) " & vbCrLf & _
             "ORDER BY e.zeitpunkt"
 mins(AWlf) = 7
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
 ' 115
' ktag fehlerhaft
 AwN(AWlf) = "Unvorhergesehene Inanspruchnahme 2 (01101) Mo-Fr 22-7, WE 19-7 (vorher 14)"
' sql(AWlf) = "SELECT e.Pat_ID AS pat_id,LEFT(DATE_FORMAT(e.ZeitPunkt,_utf8'%a, %d.%m.%y %H:%i'),20) AS Datum, LEFT(e.Art,5) AS art,e.Inhalt AS inhalt FROM `eintraege` e LEFT JOIN `faelle` f USING (fid) WHERE ((e.ZeitPunkt BETWEEN CONCAT(YEAR((NOW() - INTERVAL 14 DAY)),_latin1'-',((((month((NOW() - INTERVAL 14 DAY)) - 1) DIV 3) * 3) + 1),_latin1'-01') AND (CONCAT(YEAR((NOW() - INTERVAL 14 DAY)),_latin1'-',((((month((NOW() - INTERVAL 14 DAY)) - 1) DIV 3) * 3) + 1),_latin1'-01') + INTERVAL 3 MONTH)) AND (e.Art IN (_latin1'notiz',_latin1'ni',_latin1'telef',_latin1'gs',_latin1'ep',_latin1'wr',_latin1'bga',_latin1'gstel',_latin1'rz',_latin1'ga'))) AND (((time(e.zeitpunkt) <'07:00' OR time(e.zeitpunkt) >'22:00') AND weekday(e.zeitpunkt) BETWEEN 0 AND 5) OR ((weekday(e.zeitpunkt) BETWEEN 6 AND 7) AND NOT (time(e.zeitpunkt) BETWEEN '07:00' AND '19:00'))) AND schgr <> '90' AND NOT goäkatnr IN (40,41) ORDER BY e.Pat_ID;"
 sql(AWlf) = "SELECT e.pat_id AS Pat_id, LEFT(DATE_FORMAT(e.zeitpunkt,'%a, %d.%m.%y %H:%i'),20) AS Datum, LEFT(e.Art,5) AS Art, e.Inhalt AS Inhalt " & vbCrLf & _
             "FROM `eintraege` e LEFT JOIN `aktfv` f ON e.pat_id = f.pat_id AND e.art NOT IN ('vac','cia','c19i') AND e.inhalt NOT RLIKE 'impfung|armschmerzen|gefaxt|Rezepte erstellt|Urogenitalflora' " & vbCrLf & _
             "LEFT JOIN `leistungen` l ON l.pat_id = e.pat_id AND leistung = '01101' AND DATE(l.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
             "WHERE e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND (e.Art IN (" & artspezG & ")) " & vbCrLf & _
             "AND (((time(e.zeitpunkt) <'07:00' OR time(e.zeitpunkt) >'22:00') AND weekday(e.zeitpunkt) BETWEEN 0 AND 4) OR ((weekday(e.zeitpunkt) BETWEEN 5 AND 6) AND NOT (time(e.zeitpunkt) BETWEEN '07:00' AND '19:00'))) AND schgr <> '90' AND NOT goäkatnr IN ('40','41') " & vbCrLf & _
             " AND e.inhalt NOT LIKE '%nachgef%' AND e.inhalt NOT LIKE '=>%' " & vbCrLf & _
             "AND ISNULL(leistung) AND art<>'EKG' AND art<>'LZRR' AND art<>'Lufu'" & vbCrLf & _
             "AND NOT EXISTS (SELECT 0 FROM leistungen WHERE pat_id=e.pat_id AND leistung IN (01100,01101,01205,01207,01210,01212,01214,01216,01218,01410,01411,01412,01415,01418,01949,01950,01951,01953,03373,04373,04564,04565,04566,04572,04573,13610,13611,13612,13620,13621,13622,37306) AND zeitpunkt BETWEEN e.zeitpunkt AND e.zeitpunkt + INTERVAL 1 hour)" & vbCrLf & _
             "ORDER BY e.Pat_ID"
 mins(AWlf) = 7
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
 ' 116
 ' ktag fehlerhaft
 AwN(AWlf) = "Inanspruchnahme an Samstagen (01102) 7-14 Spr. (vorher 15)"
 sql(AWlf) = "SELECT e.Pat_ID AS pat_id,gesname(e.pat_id) Name, LEFT(DATE_FORMAT(e.ZeitPunkt,'%a, %d.%m.%y %H:%i'),20) AS Datum, LEFT(e.Art,5) AS art,e.Inhalt AS inhalt " & vbCrLf & _
             "FROM `eintraege` e LEFT JOIN aktfv USING (fid) " & vbCrLf & _
             "LEFT JOIN `leistungen` l ON l.pat_id = e.pat_id AND leistung = '01102' AND DATE(l.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
             "LEFT JOIN leistungen besuch ON besuch.pat_id = e.pat_id AND DATE(besuch.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
             "WHERE (e.Art IN (" & artspezG & ")) " & vbCrLf & _
             "AND (((time(e.zeitpunkt) BETWEEN '07:00' AND '14:00') AND weekday(e.zeitpunkt) = 5)) " & vbCrLf & _
             " AND e.inhalt NOT LIKE '%nachgef%' " & vbCrLf & _
             "AND ISNULL(l.leistung) " & vbCrLf & _
             "AND ISNULL(besuch.leistung) " & vbCrLf & _
             "AND schgr <> '90' AND NOT goäkatnr IN ('40','41') ORDER BY e.Pat_ID"
 mins(AWlf) = 7
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 19.2.11
' AwN(AWlf) = "Überweisungsleistungen auf Originalschein"
' sql(AWlf) = "SELECT f.pat_id,schgr,leistung FROM " & aktf & " LEFT JOIN `leistungen` USING (fid) WHERE leistung IN ('97360A', '97360B', '97370B', '97371B', '97320', '97321', '97322', '97313', '97323', '97333') AND schgr = '00' AND NOT goäkatnr IN (40,41)"
' mins(AWlf) = 7
' maxs(AWlf) = 16
' AWlf = AWlf + 1
 
' 19.2.11
' AwN(AWlf) = "Fehlende 97350A für D.m.1 mit DMP bei BKK, BKN, IKK, EK"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, l.leistung AS Leistung, icd, REPLACE(notiz,char(13),'') AS notiz, kateg, schgr  FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND obdauer <> 0 AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0) AND icd LIKE 'E10%' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (leistung <> '97350A' OR ISNULL(leistung)) " & vbCrLf & _
' "AND kateg IN ('BKK','BKN','IKK','EK') " & vbCrLf & _
' "AND NOT ISNULL(icd) " & vbCrLf & _
' "AND NOT notiz REGEXP 'DMP *nein'"
'' "AND notiz LIKE '%DMP%' " & vbCrLf & _
'' "AND NOT ISNULL(notiz) " & vbCrLf & _
' mins(AWlf) = 7
' maxs(AWlf) = 30
' AWlf = AWlf + 1
 
' AwN(AWlf) = "Fehlende oder falsche 97310 (Erwachsene), 97312 (Kinder), 97313 (Schwangere) für D.m.1 mit DMP bei AOK"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesnameg(f.pat_id) Name, l.leistung AS Leistung, icd, grav.fid AS schwanger, REPLACE(notiz,char(13),'') AS notiz, kateg, schgr  FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND obdauer <> 0 AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0) AND icd LIKE 'E10%' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "LEFT JOIN (SELECT fid FROM `diagnosen` d WHERE icd LIKE 'O24%' AND diagsicherheit IN (' ','G') AND COALESCE(Dggel,0)=0) AS grav ON f.fid = grav.fid " & vbCrLf & _
' "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (ISNULL(leistung) OR (ADDDATE(gebdat,INTERVAL 18 YEAR) >= " & qanfS & " AND leistung <> '97312' AND ISNULL(schwanger)) OR (ADDDATE(gebdat,INTERVAL 18 YEAR) < " & qanfS & " AND leistung <> '97310' AND ISNULL(schwanger))  OR (ADDDATE(gebdat,INTERVAL 18 YEAR) < " & qanfS & " AND leistung <> '97313' AND NOT ISNULL(schwanger))) " & vbCrLf & _
' "AND kateg IN ('AOK') " & vbCrLf & _
' "AND NOT ISNULL(icd) " & vbCrLf & _
' "AND NOT notiz REGEXP 'DMP *nein'"
'' "AND notiz LIKE '%DMP%' " & vbCrLf & _
'' "AND NOT ISNULL(notiz) " & vbCrLf & _
' mins(AWlf) = 7
' maxs(AWlf) = 30
' AWlf = AWlf + 1
 
' 19.2.11
' AwN(AWlf) = "Fehlende 97360A für D.m.2 mit DMP bei BKK, BKN, IKK, EK"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesnameg(f.pat_id) Name, l.leistung AS Leistung, obmednetz, icd, REPLACE(notiz,char(13),'') AS notiz, kateg, schgr  FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND obdauer <> 0 AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0) AND icd LIKE 'E11%' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (leistung <> '97360A' OR ISNULL(leistung)) " & vbCrLf & _
' "AND kateg IN ('BKK','BKN','IKK','EK') " & vbCrLf & _
' "AND NOT ISNULL(icd) " & vbCrLf & _
' "AND obMedNetz = 0 " & vbCrLf & _
' "AND NOT notiz REGEXP 'DMP *nein'"
' mins(AWlf) = 7
' maxs(AWlf) = 30
' AWlf = AWlf + 1
' "AND notiz LIKE '%DMP%' " & vbCrLf & _
' "AND NOT ISNULL(notiz) " & vbCrLf & _

' AwN(AWlf) = "Fehlende oder falsche 97320, 97321, 97322, 97323, für D.m.2 mit DMP bei AOK"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesnameg(f.pat_id) Name, l.leistung AS Leistung, obmednetz, icd, grav.fid AS schwanger, REPLACE(notiz,char(13),'') AS notiz, kateg, schgr, Therakt, Ther1  FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN (" & BetrPausch & ")) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON (f.pat_id = d.pat_id AND obdauer <> 0 AND diagsicherheit<> 'A' AND COALESCE(Dggel,0)=0) AND icd LIKE 'E11%' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "LEFT JOIN (SELECT fid FROM `diagnosen` d WHERE icd = 'Z34' AND diagsicherheit IN ('G',' ') AND COALESCE(Dggel,0)=0) AS grav ON f.fid = grav.fid " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (ISNULL(leistung) OR " & vbCrLf & _
' "  (ADDDATE(gebdat,INTERVAL 18 YEAR) >= " & qanfS & " AND leistung <> '97322' AND ISNULL(schwanger)) OR " & vbCrLf & _
' "  (ADDDATE(gebdat,INTERVAL 18 YEAR) < " & qanfS & " AND ISNULL(schwanger) AND " & vbCrLf & _
' "    (((NOT ((ther1 LIKE '%Diät%' OR ther1 LIKE 'OAD%') AND (therakt LIKE '%Diät%' OR therakt LIKE 'OAD%')) ) AND leistung <> '97320') OR " & vbCrLf & _
' "     ((    ((ther1 LIKE '%Diät%' OR ther1 LIKE 'OAD%') AND (therakt LIKE '%Diät%' OR therakt LIKE 'OAD%')) ) AND leistung <> '97321'))) OR " & vbCrLf & _
' "  (ADDDATE(gebdat,INTERVAL 18 YEAR) < " & qanfS & " AND NOT ISNULL(schwanger) AND leistung <> '97323')) " & vbCrLf & _
' "AND kateg IN ('AOK') " & vbCrLf & _
' "AND NOT ISNULL(icd) " & vbCrLf & _
' "AND obMedNetz = 0 " & vbCrLf & _
' "AND NOT notiz REGEXP 'DMP *nein'"
'' "AND notiz LIKE '%DMP%' " & vbCrLf & _
'' "AND NOT ISNULL(notiz) " & vbCrLf & _
' mins(AWlf) = 7
' maxs(AWlf) = 26
' AWlf = AWlf + 1
 
 ' 117
 AwN(AWlf) = "Fehlende 01740 für Beratung zur Früherkennung des kolorektalen Carcinoms nach Einträgen ab 50./55. Lebensjahr (vorher 32)"
#If ganzalt Then
 sql(AWlf) = _
 "SELECT * FROM " & vbCrLf & _
 "(SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id pat_id, gesname(f.pat_id) Name, patAlter(f.pat_id) PAlter, DATE(zeitpunkt) Tag, dok.Art, dok.Inhalt FROM " & aktf & " " & vbCrLf & _
 "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('01740')) AS l ON l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qanf()AND qend()" & vbCrLf & _
 "LEFT JOIN (SELECT pat_id, zeitpunkt, Art, Inhalt FROM `eintraege` WHERE art LIKE 'colo%' OR MATCH(inhalt) AGAINST ('Vorsorgecolo* Kontrollcolo* Vorbeugecolo* Pancolo* Gastrocolo* colo* Darmkrebs* Vorsorgedarmsp* Darmspie* Magendarmspie*' IN BOOLEAN MODE) " & vbCrLf & _
 ") AS dok ON f.pat_id = dok.pat_id " & vbCrLf & _
 "AND zeitpunkt > DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 14 DAY)),'-',(((month(SUBDATE(NOW(),INTERVAL 14 DAY))-1) div 3) * 3)+1,'-1')) " & vbCrLf & _
 "AND zeitpunkt < ADDDATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 14 DAY)),'-',(((month(SUBDATE(NOW(),INTERVAL 14 DAY))-1) div 3) * 3)+1,'-1'),INTERVAL 3 MONTH) " & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
 "WHERE (leistung <> '01740' OR ISNULL(leistung)) " & vbCrLf & _
 "AND palter>=50 " & vbCrLf & _
 "AND NOT ISNULL(inhalt)"
#ElseIf alt Then
sql(AWlf) = _
"SELECT f.pat_id, gesname(f.pat_id) Name, DATE_FORMAT(e.zeitpunkt,'%e.%c.%y') zp, e.art, e.inhalt " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
"LEFT JOIN eintraege e USING (pat_id) " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id = f.pat_id AND l.leistung = '01740' " & vbCrLf & _
"WHERE (art LIKE 'colo%' OR MATCH(inhalt) AGAINST ('Vorsorgecolo* Kontrollcolo* Vorbeugecolo* Pancolo* Gastrocolo* colo* Darmkrebs* Vorsorgedarmsp* Darmspie* Magendarmspie*' IN BOOLEAN MODE)) " & vbCrLf & _
"   AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"   AND ((e.inhalt LIKE '%geraten%' AND NOT e.inhalt LIKE '%abgeraten%') OR e.inhalt LIKE '%empfohlen%' OR e.inhalt LIKE '%erläut%' OR e.art='coloauf') " & vbCrLf & _
"   AND ISNULL(l.leistung) AND ADDDATE(n.gebdat,INTERVAL IF(n.geschlecht='w',55,50) YEAR) < e.zeitpunkt;"
#Else
sql(AWlf) = _
"SELECT f.pat_id, gesname(f.pat_id) Name, DATE_FORMAT(e.zeitpunkt,'%e.%c.%y') zp, e.art, e.inhalt" & vbCrLf & _
"FROM aktfvs f" & vbCrLf & _
"LEFT JOIN namen n USING (pat_id)" & vbCrLf & _
"LEFT JOIN eintraege e USING (pat_id)" & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id = f.pat_id AND l.leistung = '01740'" & vbCrLf & _
"WHERE (art LIKE 'colo%' OR inhalt RLIKE 'colo|darm(krebs|sp)')" & vbCrLf & _
"   AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"   AND ((e.inhalt LIKE '%geraten%' AND NOT e.inhalt LIKE '%abgeraten%') OR e.inhalt LIKE '%empfohlen%' OR e.inhalt LIKE '%erläut%' OR e.art='coloauf') " & vbCrLf & _
"   AND ISNULL(l.leistung) AND ADDDATE(n.gebdat,INTERVAL IF(n.geschlecht='w',55,50) YEAR) < e.zeitpunkt;"

#End If
 mins(AWlf) = 7
 maxs(AWlf) = 120
 AWlf = AWlf + 1
' extrem lückenhaft
' AwN(AWlf) = "Fehlende Versichertenpauschale"
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, LEFT(CONCAT(n.Nachname, ' ', n.Vorname),25) AS Name, obMedNetz, schgr, l.leistung AS Leistung FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('01210', '03110','03111','03112','03120','03121','03122')) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (ISNULL(leistung)) AND obmednetz = 0 AND pat_id NOT IN (2, 2420)" ' Musterwoman, Bereitschaftsdienst
' mins(AWlf) = 7
' maxs(AWlf) = 120
' AWlf = AWlf + 1
'
' sehr lückenhaft
' AwN(AWlf) = "Fehlende Chronikerpauschale 03212 alt mit union"
' sql(AWlf) = _
' "SELECT pat_id, name, ct, icd, Schgr, Leistung FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, LEFT(CONCAT(n.Nachname, ' ', n.Vorname),25) AS Name, Schgr, l.leistung AS leistung, ct, icd FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('03212')) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN (SELECT COUNT(0) AS ct, pat_id FROM (SELECT pat_id, DATE(zeitpunkt) AS zp FROM " & aktf & _
' " LEFT JOIN `eintraege` e USING (pat_id) WHERE (art IN (" & latinis(artspezG) & ")) " & vbCrLf & _
' "AND zeitpunkt BETWEEN " & qanfS & " AND " & qendS & " " & vbCrLf & _
' "GROUP BY pat_id UNION " & vbCrLf & _
' "SELECT pat_id, DATE(zeitpunkt) AS zp FROM " & aktf & " LEFT JOIN `rezepteintraege`  r USING (pat_id) WHERE true " & vbCrLf & _
' "AND zeitpunkt BETWEEN " & qanfS & " AND " & qendS & " " & vbCrLf & _
' "GROUP BY pat_id, zp) AS innen GROUP BY pat_id) AS dok ON f.pat_id = dok.pat_id " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON d.pat_id = f.pat_id AND (obdauer <> 0 OR (d.diagdatum BETWEEN " & qanfS & " AND " & qendS & ")) AND (((icd LIKE 'E1%' AND NOT icd LIKE 'E15%') OR icd LIKE 'I10%') AND diagsicherheit <> 'A' AND COALESCE(Dggel,0)=0) " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE (ISNULL(leistung)) AND ct > 1 AND NOT ISNULL(icd) "
' mins(AWlf) = 7
' maxs(AWlf) = 120
' AWlf = AWlf + 1
'
' vollständig, aber gemischt
' AwN(AWlf) = "Falsche Versicherungspauschale 03110,03111,03112 bzw. 03120,03121,03122 alt mit union"
' sql(AWlf) = _
' "SELECT fpat_id AS pat_id, ffid AS fid, name, ct, icd, obMedNetz, Schgr, MSchgr, Grundl, ChrON FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS Fpat_id, f.fid AS FFid, LEFT(CONCAT(n.Nachname, ' ', n.Vorname),25) AS Name, obMedNetz, Schgr, MSchGr, gl.leistung AS Grundl, l.leistung AS chron, ct, icd FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('03110','03111','03112','03120','03121','03122')) AS gl ON f.fid = gl.fid " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('03212')) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "LEFT JOIN (SELECT COUNT(0) AS ct, pat_id FROM (SELECT pat_id, DATE(zeitpunkt) AS zp FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN `eintraege` e USING (pat_id) WHERE (art IN (" & latinis(artspezG) & ")) " & vbCrLf & _
' "AND zeitpunkt BETWEEN " & qanfS & " AND " & qendS & " " & vbCrLf & _
' "GROUP BY pat_id UNION " & vbCrLf & _
' "SELECT pat_id, DATE(zeitpunkt) AS zp FROM " & aktf & " LEFT JOIN `rezepteintraege`  r USING (pat_id) WHERE true " & vbCrLf & _
' "AND zeitpunkt BETWEEN " & qanfS & " AND " & qendS & " " & vbCrLf & _
' "GROUP BY pat_id, zp) AS innen GROUP BY pat_id) AS dok ON f.pat_id = dok.pat_id " & vbCrLf & _
' "LEFT JOIN (SELECT pat_id, MIN(schgr) AS mschgr,quartal FROM `faelle` GROUP BY pat_id, quartal) AS mschgr ON f.pat_id = mschgr.pat_id AND mschgr.quartal = '" & aktQ & "' " & vbCrLf & _
' "LEFT JOIN (SELECT * FROM (SELECT * FROM `diagnosen` d WHERE (obdauer <> 0 OR (d.diagdatum BETWEEN " & qanfS & " AND " & qendS & ")) AND (((icd LIKE 'E1%' AND NOT icd LIKE 'E15%') OR icd LIKE 'I10%') AND diagsicherheit <> 'A' AND COALESCE(Dggel,0)=0) ORDER BY icd) AS d GROUP BY pat_id) AS d ON d.pat_id = f.pat_id " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY gl.leistung DESC) AS innen) AS innen " & vbCrLf & _
' "WHERE (schgr = '24' AND grundl IN ('03110','03111','03112')) OR (schgr = '00' AND NOT goäkatnr IN (40,41) AND grundl IN ('03120','03121','03122')) " & vbCrLf & _
' "ORDER BY pat_id "
'' "WHERE (ISNULL(leistung)) AND ct > 1 AND NOT ISNULL(icd) "
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
'
#If ebmalt Then
' ktag fehlerhaft
' AwN(AWlf) = "Dem falschen Fall zugeordnete Chronikerpauschale 03212"
' '  AND COALESCE(d.Dggel,0)=0
' sql(AWlf) = _
' "SELECT fpat_id AS pat_id, ffid AS fid, name, ct, ICD, Schgr, MSchgr, Grundl, ChrON FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS Fpat_id, f.fid AS FFid, LEFT(CONCAT(n.Nachname, ' ', n.Vorname),25) AS Name, Schgr, MSchGr, gl.leistung AS Grundl, l.leistung AS chron, ct, ICD FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('03110','03111','03112','03120','03121','03122')) AS gl ON f.fid = gl.fid " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('03212')) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN (SELECT COUNT(0) AS ct, pat_id FROM (SELECT pat_id, DATE(zeitpunkt) zp FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN `eintraege` e USING (pat_id) WHERE (art IN (" & artspezG & ")) " & vbCrLf & _
' "AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
' "GROUP BY pat_id UNION " & vbCrLf & _
' "SELECT pat_id, DATE(zeitpunkt) AS zp FROM " & aktf & " LEFT JOIN `rezepteintraege`  r USING (pat_id) WHERE true " & vbCrLf & _
' "AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
' "GROUP BY pat_id, zp) AS innen GROUP BY pat_id) AS dok ON f.pat_id = dok.pat_id " & vbCrLf & _
' "LEFT JOIN (SELECT pat_id, MIN(schgr) AS mschgr,quartal FROM `faelle` GROUP BY pat_id, quartal) AS mschgr ON f.pat_id = mschgr.pat_id AND mschgr.quartal = '" & AktQ & "' " & vbCrLf & _
' "LEFT JOIN (SELECT * FROM (SELECT * FROM `diagnosen` d WHERE (d.obdauer <> 0 OR (d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & ")) AND ((d.icd RLIKE '^E1[0-4]' OR d.icd LIKE 'I10%') AND d.diagsicherheit <> 'A') ORDER BY d.icd) AS d GROUP BY pat_id) AS d ON d.pat_id = f.pat_id " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY gl.leistung DESC) AS innen) AS innen " & vbCrLf & _
' "WHERE ((schgr = '24' AND mschgr = 0 AND NOT ISNULL(chron))) " & vbCrLf & _
' "ORDER BY pat_id "
'' "WHERE (ISNULL(leistung)) AND ct > 1 AND NOT ISNULL(icd) "
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
#Else
#End If

 ' 118
 AwN(AWlf) = "Fehlende Ziffer 32022 (Ausnahmeziffer Labor) (lauto) (vorher 38)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) Name, GROUP_CONCAT(vorh.leistung) Laborleistung, '32022 dazu' LEIFEHLER, DATE(min(vorh.zeitpunkt)) LEIDAT, f.LANRID " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN leistungen vorh ON vorh.pat_id=f.pat_id AND vorh.zeitpunkt BETWEEN qanf()AND qend() AND vorh.leistung BETWEEN '32030' AND '32135' " & vbCrLf & _
"LEFT JOIN leistungen ausn ON ausn.pat_id=f.pat_id AND ausn.zeitpunkt BETWEEN qanf()AND qend() AND ausn.leistung = '32022' " & vbCrLf & _
"LEFT JOIN `diagview` d ON d.pat_id = f.pat_id AND (d.gicd REGEXP '^E1[0-4]\.' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))) " & vbCrLf & _
"WHERE ISNULL(ausn.Leistung) AND NOT ISNULL(vorh.Leistung) AND NOT ISNULL(d.ICD) " & vbCrLf & _
"GROUP BY f.pat_id;"

' "SELECT * FROM " & vbCrLf & _
 "(SELECT * FROM " & vbCrLf & _
 "(SELECT f.pat_id PID, gesname(f.pat_id) Name, l.leistung AS leistung, zif.leistung AS ziffer, icd FROM " & aktf & " " & vbCrLf & _
 "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung >= '32030' AND leistung <= '32135' GROUP BY fid) AS l ON f.fid = l.fid " & vbCrLf & _
 "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('32022')) AS zif ON f.fid = zif.fid " & vbCrLf & _
 "LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND d.obdauer <> 0 AND d.diagsicherheit NOT IN ('A','Z') AND COALESCE(d.Dggel,0)=0 AND icd REGEXP '^E1[0-4]\.|^O24.4' " & vbCrLf & _
 "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
 "WHERE (ziffer <> '32022' OR ISNULL(ziffer)) " & vbCrLf & _
 "AND NOT ISNULL(leistung) " & vbCrLf & _
 "AND NOT ISNULL(icd) "
 mins(AWlf) = 7
 maxs(AWlf) = 20
 AWlf = AWlf + 1

' 119
 AwN(AWlf) = "Fehlende Ausschlußziffer Antikoagulation 32015: (lauto) (vorher 55)"
' sql(AWlf) = "SELECT pat_id PID, Name,DATE(" & qtAnf(FristS) & ") LEIDAT, '32015 dazu' LEIFEHLER, icd, medanfang, datum, LANRID FROM (" & vbCrLf & _
             " SELECT f.pat_id, gesnameg(f.pat_id) Name, SUM(lzahl) Zahl, d.icd, mp.medanfang, CAST(mp.datum As Date) Datum, f.LANRID " & vbCrLf & _
             " FROM aktfvs f " & vbCrLf & _
             " LEFT JOIN medplan mp ON f.pat_id = mp.pat_id " & vbCrLf & _
             " LEFT JOIN medarten m ON mp.medanfang = m.medikament " & vbCrLf & _
             "  AND ((mp.zeitpunkt >= (SELECT MAX(zeitpunkt) FROM medplan WHERE zeitpunkt < " & qtAnf(FristS) & " AND pat_id = f.pat_id))  OR (SELECT MAX(zeitpunkt) FROM medplan WHERE zeitpunkt < " & qtAnf(FristS) & " AND pat_id = f.pat_id) IS NULL AND mp.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
             "  AND ((mp.zeitpunkt <  (SELECT MIN(zeitpunkt) FROM medplan WHERE zeitpunkt > " & qtEnd(FristS) & " AND pat_id = f.pat_id))  OR (SELECT MIN(zeitpunkt) FROM medplan WHERE zeitpunkt > " & qtEnd(FristS) & " AND pat_id = f.pat_id) IS NULL AND mp.zeitpunkt<=" & qtEnd(FristS) & ") " & vbCrLf & _
             " LEFT JOIN leistungen l ON f.fid = l.fid AND l.leistung = '32015' " & vbCrLf & _
             " LEFT JOIN diagnosen d ON d.pat_id=f.pat_id AND d.icd='Z92.1' AND d.obdauer<>0 AND COALESCE(d.Dggel,0)=0 AND d.diagsicherheit IN (' ','G') " & vbCrLf & _
             "WHERE antikoag<>0 " & vbCrLf & _
             "GROUP BY pat_id) i WHERE zahl=0 OR ISNULL(zahl)"
' CAST( ... As Date) und DATE(...) sind im Test haargenausoschnell
' gleichwertig mit obigem 15.6.22:
' AND COALESCE(d.Dggel,0)=0
 sql(AWlf) = "SELECT f.pat_id PID, gesnameg(f.pat_id) PName, DATE(qanf()) LEIDAT, '32015 dazu' LEIFEHLER, d.icd, mp.medanfang, CAST(mp.datum As Date) Datum, f.LANRID " & vbCrLf & _
             "FROM aktfvs f " & vbCrLf & _
             "LEFT JOIN medplan mp ON f.pat_id = mp.pat_id " & vbCrLf & _
             "LEFT JOIN medarten m ON mp.medanfang = m.medikament " & vbCrLf & _
             " AND ((mp.zeitpunkt >= (SELECT MAX(zeitpunkt) FROM medplan WHERE zeitpunkt < " & qtAnf(FristS) & " AND pat_id = f.pat_id))  OR (SELECT MAX(zeitpunkt) FROM medplan WHERE zeitpunkt < " & qtAnf(FristS) & " AND pat_id = f.pat_id) IS NULL AND mp.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
             " AND ((mp.zeitpunkt <  (SELECT MIN(zeitpunkt) FROM medplan WHERE zeitpunkt > " & qtEnd(FristS) & " AND pat_id = f.pat_id))  OR (SELECT MIN(zeitpunkt) FROM medplan WHERE zeitpunkt > " & qtEnd(FristS) & " AND pat_id = f.pat_id) IS NULL AND mp.zeitpunkt<=" & qtEnd(FristS) & ") " & vbCrLf & _
             "LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() AND l.leistung = '32015' " & vbCrLf & _
             "LEFT JOIN diagview d ON d.pat_id=f.pat_id AND d.gicd='Z92.1' AND d.obdauer<>0 " & vbCrLf & _
             "WHERE antikoag<>0 " & vbCrLf & _
             "GROUP BY f.pat_id" & vbCrLf & _
             "HAVING COALESCE(SUM(lzahl),0)=0" & vbCrLf & _
             "ORDER BY f.pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 120
' ktag fehlerhaft
AwN(AWlf) = "Fehlende Leistung für cerebrale Doppler/Duplex (33060) (vorher 39)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33060'),0) `cwzl(33060)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('doppler','dop','duplex','dup') " & vbCrLf & _
"AND (e.inhalt RLIKE 'hals|arotis|ACI|A.car.|Carot' OR (inhalt RLIKE '%acc%' AND NOT inhalt RLIKE 'saph|ACC\.')) " & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `cwzl(33060)`=0" & vbCrLf & _
"ORDER BY f.pat_id"
'", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33076'),0) `komprzl(33076)`" & vbCrLf & _
'", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung IN ('33042','33042A')),0) `sonozl(33042/A)`" & vbCrLf & _

'sql(AWlf) = _
"SELECT f.`pat_id` AS Pat_ID, ezp AS Zeitpunkt, art, Inhalt FROM " & vbCrLf & _
 aktf & " " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.`pat_id`= a.`pat_id` " & vbCrLf & _
"LEFT JOIN (SELECT * FROM " & vbCrLf & _
"(((SELECT pat_id AS pid,  LEFT(art,7) AS art, inhalt, zeitpunkt AS ezp, DATE(zeitpunkt) zp FROM `eintraege` e " & vbCrLf & _
"WHERE art IN ('doppler','dop','duplex') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") AS eint " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS cwpat_id, fid AS cwfid, DATE(zeitpunkt) cwltag, leistung AS cw FROM `leistungen` WHERE leistung IN ('33060')) AS cw " & vbCrLf & _
"ON eint.pid = cw.cwpat_id AND eint.zp = cw.cwltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS komprpat_id, fid AS komprfid, DATE(zeitpunkt) komprltag, leistung AS kompr FROM `leistungen` WHERE leistung IN ('33076')) AS kompr " & vbCrLf & _
"ON eint.pid = kompr.komprpat_id AND eint.zp = kompr.komprltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS sonopat_id, fid AS sonofid, DATE(zeitpunkt) sonoltag, leistung AS sono FROM `leistungen` WHERE leistung IN ('33042','33042A')) AS sono " & vbCrLf & _
"ON eint.pid = sono.sonopat_id AND eint.zp = sono.sonoltag) " & vbCrLf & _
"as ges " & vbCrLf & _
"ON (ges.PID = f.`Pat_ID`) " & vbCrLf & _
"WHERE NOT ISNULL(ezp) AND ISNULL(cwltag) " & vbCrLf & _
"AND (inhalt RLIKE 'hals|arotis|ACI|A.car.|Carot' OR (inhalt LIKE '%ACC%' AND NOT inhalt LIKE '%SAPH%' AND NOT inhalt LIKE '%ACC\.%')) " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1

' 121
' ktag fehlerhaft
AwN(AWlf) = "Fehlende Leistung für Extremitäten-Doppler/Duplex (33061) (vorher 40)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33061'),0) `cwzl(33061)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('doppler','dop','duplex','dup')" & vbCrLf & _
"AND (e.inhalt RLIKE '^Bein|Belastung|A.d.|A.t.|biph|monoph|postst|Fußart|IDR|A. fem.|A.fem.|in Ruhe|in Ruhe|aorta' AND NOT e.inhalt RLIKE 'halsvenen|^halsschlagadern|halsart|Nierenarterien') " & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `cwzl(33061)`=0" & vbCrLf & _
"ORDER BY f.pat_id"
'sql(AWlf) = _
"SELECT f.pat_id, gesname(f.pat_id),e.zeitpunkt, e.art,e.inhalt " & vbCrLf & _
" FROM aktfvs f " & vbCrLf & _
" LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('doppler','dop','duplex') " & vbCrLf & _
"   AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"   AND e.inhalt RLIKE '^Bein|Belastung|A.d.|A.t.|biph|monoph|postst|Fußart|IDR|A. fem.|A.fem.|in Ruhe|in Ruhe|aorta' " & vbCrLf & _
"   AND NOT e.inhalt RLIKE '^halsschlagadern|Nierenarterien' " & vbCrLf & _
"LEFT JOIN leistungen cwl ON cwl.fid=f.fid AND cwl.leistung = '33061' AND DATE(cwl.zeitpunkt)=DATE(e.zeitpunkt) " & vbCrLf & _
"WHERE NOT ISNULL(e.pat_id) AND ISNULL(cwl.pat_id) " & vbCrLf & _
"GROUP BY f.pat_id,DATE(e.zeitpunkt); "

'sql(AWlf) = _
"SELECT f.`pat_id` AS Pat_ID, ezp AS Zeitpunkt, art, Inhalt FROM " & aktf & " " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.`pat_id`= a.`pat_id` " & vbCrLf & _
"LEFT JOIN (SELECT * FROM " & vbCrLf & _
"(((SELECT pat_id AS pid,  LEFT(art,7) AS art, inhalt, zeitpunkt AS ezp, DATE(zeitpunkt) AS tag FROM `eintraege` e " & vbCrLf & _
"WHERE art IN ('doppler','dop','duplex') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") AS eint " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS cwpat_id, fid AS cwfid, DATE(zeitpunkt) AS cwltag, leistung AS cw FROM `leistungen` WHERE leistung IN ('33061')) AS cw " & vbCrLf & _
"ON eint.pid = cw.cwpat_id AND eint.tag = cw.cwltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS komprpat_id, fid AS komprfid, DATE(zeitpunkt) AS komprltag, leistung AS kompr FROM `leistungen` WHERE leistung IN ('33076')) AS kompr " & vbCrLf & _
"ON eint.pid = kompr.komprpat_id AND eint.tag = kompr.komprltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS sonopat_id, fid AS sonofid, DATE(zeitpunkt) AS sonoltag, leistung AS sono FROM `leistungen` WHERE leistung IN ('33042','33042A')) AS sono " & vbCrLf & _
"ON eint.pid = sono.sonopat_id AND eint.tag = sono.sonoltag) " & vbCrLf & _
"as ges " & vbCrLf & _
"ON (ges.PID = f.`Pat_ID`) " & vbCrLf & _
"WHERE NOT ISNULL(ezp) AND ISNULL(cwltag) " & vbCrLf & _
"AND inhalt RLIKE '^Bein|Belastung|A.d.|A.t.|biph|monoph|postst|Fußart|IDR|A. fem.|A.fem.|in Ruhe|in Ruhe|aorta' " & vbCrLf & _
"AND NOT inhalt RLIKE '^halsschlagadern|Nierenarterien' " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1

' 122
' ktag fehlerhaft
AwN(AWlf) = "Doppelte Leistung für arterielle Doppler/Duplex (33061) (vorher 41)"
sql(AWlf) = _
"SELECT COALESCE(SUM(l.lzahl),0) zahl, f.pat_id, DATE(l.zeitpunkt) zp " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN `leistungen` l ON l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qanf()AND qend()" & vbCrLf & _
"     AND leistung = '33061'" & vbCrLf & _
"GROUP BY f.pat_id, DATE(l.zeitpunkt) " & vbCrLf & _
"HAVING zahl>1 " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1


' 123
' ktag fehlerhaft
' imponieren
AwN(AWlf) = "Fehlende Leistung für Nierenarterien-/Mesenterialgefäß-Doppler/Duplex: (33042) (vorher 42)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung IN ('33042','33042A')),0) `sonozl(33042/A)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('doppler','dop','duplex','dup')" & vbCrLf & _
"AND (e.inhalt RLIKE '([^o]|^)Nieren|Mesent') " & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `sonozl(33042/A)`=0" & vbCrLf & _
"ORDER BY f.pat_id"
'", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33061'),0) `cwzl(33061)`" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33076'),0) `komprzl(33076)`" & vbCrLf & _

'sql(AWlf) = _
"SELECT f.`pat_id` AS Pat_ID, ezp AS Zeitpunkt,  LEFT(art,7) AS art, Inhalt, ges.fid " & vbCrLf & _
"FROM " & aktf & " " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.`pat_id`= a.`pat_id` " & vbCrLf & _
"LEFT JOIN (SELECT * FROM " & vbCrLf & _
"(((SELECT pat_id AS pid,  LEFT(art,7) AS art, inhalt, zeitpunkt AS ezp, DATE(zeitpunkt) zp, fid FROM `eintraege` e " & vbCrLf & _
"WHERE art IN ('doppler','dop','duplex') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") AS eint " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS cwpat_id, fid AS cwfid, DATE(zeitpunkt) cwltag, leistung AS cw FROM `leistungen` WHERE leistung IN ('33061')) AS cw " & vbCrLf & _
"ON eint.pid = cw.cwpat_id AND eint.zp = cw.cwltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS komprpat_id, fid AS komprfid, DATE(zeitpunkt) komprltag, leistung AS kompr FROM `leistungen` WHERE leistung IN ('33076')) AS kompr " & vbCrLf & _
"ON eint.pid = kompr.komprpat_id AND eint.zp = kompr.komprltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS sonopat_id, fid AS sonofid, DATE(zeitpunkt) sonoltag, leistung AS sono FROM `leistungen` WHERE leistung IN ('33042','33042A')) AS sono " & vbCrLf & _
"ON eint.pid = sono.sonopat_id AND eint.zp = sono.sonoltag) " & vbCrLf & _
"as ges " & vbCrLf & _
"ON (ges.PID = f.`Pat_ID`) " & vbCrLf & _
"WHERE NOT ISNULL(ezp) AND ISNULL(sonoltag) " & vbCrLf & _
"AND (inhalt LIKE '%Nieren%' OR inhalt LIKE '%Mesent%') " & vbCrLf & _
"AND COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE fid=ges.fid AND leistung IN ('33042','33042A')),0)<=2 " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1

' 124
' ktag fehlerhaft
AwN(AWlf) = "Fehlende Leistung für Schilddrüsensono: (33012) (vorher 43)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33012'),0) `sdl(33012)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('sono')" & vbCrLf & _
"AND (e.inhalt LIKE '%Schild%') " & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `sdl(33012)`=0" & vbCrLf & _
"ORDER BY f.pat_id"
'sql(AWlf) = _
"SELECT f.`pat_id` AS Pat_ID, ezp AS Zeitpunkt,  LEFT(art,7) AS art, Inhalt FROM " & aktf & " " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.`pat_id`= a.`pat_id` " & vbCrLf & _
"LEFT JOIN (SELECT * FROM " & vbCrLf & _
"(SELECT pat_id AS pid, LEFT(art,7) AS art, inhalt, zeitpunkt AS ezp, DATE(zeitpunkt) zp FROM `eintraege` e " & vbCrLf & _
"WHERE art IN ('sono') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") AS eint " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS cwpat_id, fid AS cwfid, DATE(zeitpunkt) sdltag, leistung AS cw FROM `leistungen` WHERE leistung IN ('33012')) AS sd " & vbCrLf & _
"ON eint.pid = sd.cwpat_id AND eint.zp = sd.sdltag) " & vbCrLf & _
"as ges " & vbCrLf & _
"ON (ges.PID = f.`Pat_ID`) " & vbCrLf & _
"WHERE NOT ISNULL(ezp) AND ISNULL(sdltag) " & vbCrLf & _
"AND (inhalt LIKE '%Schild%') " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1

' 125
' ktag fehlerhaft
AwN(AWlf) = "Fehlende Leistung für Abdomensono: (33042, 33042A) (vorher 44)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung IN ('33042','33042A')),0) `sonozl(33042/A)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('sono')" & vbCrLf & _
"AND (e.inhalt RLIKE 'Abdom|[^k]leber|(?<!o|r|den )nier(?!eng)') " & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `sonozl(33042/A)`=0;"

'sql(AWlf) = _
"SELECT f.pat_ID, gesname(f.pat_id) Name, e.zeitpunkt, LEFT(art,7) Art, Inhalt FROM " & vbCrLf & _
"aktfvs f " & vbCrLf & _
"LEFT JOIN faelle gf ON gf.fid = f.fid " & vbCrLf & _
"LEFT JOIN anamnesebogen a ON f.pat_id = a.pat_id " & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id = f.pat_id AND DATE(e.zeitpunkt) BETWEEN gf.qanf AND gf.qend AND art = 'sono' AND inhalt RLIKE 'Abd|leber|nier' " & vbCrLf & _
"LEFT JOIN leistungen l ON l.fid = f.fid AND leistung IN ('33042','33042A') " & vbCrLf & _
"WHERE NOT ISNULL(e.pat_id) AND ISNULL(l.pat_id) " & vbCrLf & _
"AND COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE fid=f.fid AND leistung IN ('33042','33042A')),0)<=2 " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1

' 126
 AwN(AWlf) = "Mehr als 2 Leistungen 33042/33042A im Behandlungsfall (vorher 115)"
 sql(AWlf) = vbCrLf & _
"SELECT * FROM (SELECT " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung IN ('33042','33042A') AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & "),0) Aktzahl, " & vbCrLf & _
"gesname(f.pat_id), f.pat_id " & vbCrLf & _
", (SELECT GROUP_CONCAT(zeitpunkt SEPARATOR ', ') FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung IN ('33042','33042A') AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Zpkte " & vbCrLf & _
"FROM aktfvs f) i WHERE aktzahl>2;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
#If vor31032018 Then
' zum 31.3.18 gestrichen: 99300
' 127
' ktag fehlerhaft
 AwN(AWlf) = "Mögliche Kandidaten für Atemnotpauschale (J44.[018][01] plus J96.1[019] (vorher 86)"
 sql(AWlf) = "SELECT f.pat_id,gesname(f.pat_id) Name, patAlter(f.pat_id) PAlter " & vbCrLf & _
       ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='tk'OR(art='tb'AND ersteller='tk')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk')))) tk " & vbCrLf & _
       ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='gs'OR(art='tb'AND ersteller='gs')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs')))) gs " & vbCrLf & _
       ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = v.pat_id AND (art='ah'OR(art='tb'AND ersteller='ah')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah')))) ah " & vbCrLf & _
       ", GROUP_CONCAT(DISTINCT jd.icd) COPDicd, GROUP_CONCAT(DISTINCT dd.icd) RIicd, GROUP_CONCAT(DISTINCT LEFT(dd.diagtext,17)) `Resp.Ins.`,DATE(e.zeitpunkt) zpt, e.inhalt Eintrag, LEFT(k.name,20) Kasse  " & vbCrLf & _
       "FROM aktfvs v " & vbCrLf & _
       "JOIN faelle f USING (fid)" & vbCrLf & _
       "LEFT JOIN kassenliste k ON k.id=f.kid" & vbCrLf & _
       "LEFT JOIN diagview jd ON f.pat_id = jd.pat_id AND jd.gicd LIKE 'J44%' " & vbCrLf & _
       "LEFT JOIN diagview dd ON f.pat_id = dd.pat_id AND dd.diagsicherheit NOT IN ('A','Z') AND (dd.icd LIKE 'R06%' OR dd.icd LIKE 'J96%' OR dd.diagtext LIKE '%dyspnoe%' OR dd.diagtext LIKE '%luftnot%' OR dd.diagtext LIKE '%atemnot%') " & vbCrLf & _
       "LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND (e.zeitpunkt > SUBDATE(" & qtAnf(FristS) & ",INTERVAL 2 YEAR) AND (e.inhalt LIKE '%dyspnoe%' OR (e.inhalt LIKE '%luftnot%' AND e.art NOT IN ('andm','andm2')) OR (e.inhalt LIKE '%atemnot%' AND e.art NOT IN ('andm','andm2')))) " & vbCrLf & _
       "WHERE kateg IN ('AOK','EK') AND NOT (ISNULL(jd.ICD) AND ISNULL(dd.ICD) AND ISNULL(e.Inhalt)) AND NOT (jd.icd RLIKE 'J44.[018][01]' AND dd.icd RLIKE 'J96.1[019]') " & vbCrLf & _
       "GROUP BY v.pat_id, e.zeitpunkt, e.inhalt " & vbCrLf & _
       "ORDER BY f.pat_id "
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
#End If
 
' 127
AwN(AWlf) = "In 33042A umzuwandelnde 033042 (vorher 86)"
sql(AWlf) = vbCrLf & _
"SELECT f.Pat_id, gesname(f.pat_id) PName,l1.zeitpunkt Zeit_01748,l2.zeitpunkt Zeit_33042" & vbCrLf & _
"FROM aktfvs f" & vbCrLf & _
"JOIN leistungen l1 ON l1.pat_id=f.pat_id and l1.zeitpunkt BETWEEN qanf() AND qend() AND l1.Leistung='01748'" & vbCrLf & _
"JOIN leistungen l2 ON l2.pat_id=f.pat_id AND l2.zeitpunkt BETWEEN qanf() AND qend() AND l2.Leistung='33042' AND DATE(l2.zeitpunkt)=DATE(l1.zeitpunkt)" & vbCrLf & _
"WHERE COALESCE((SELECT COALESCE(SUM(lzahl),0) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='01748'),0)<=1"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

' 128
' ktag fehlerhaft
AwN(AWlf) = "Fehlende Leistung für Venen-Doppler/Duplex (33076 Kompressionssono / 33061 cw-Doppler) (vorher 45)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33061'),0) `cwzl(33061)`" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33076'),0) `komprzl(33076)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('duplex','dup') " & vbCrLf & _
"AND ((e.inhalt RLIKE 'vene|varik|v[.]s[.]|saph|insuff|reflux' AND NOT e.inhalt RLIKE 'mesent|venengl|halsven|venös') OR (inhalt LIKE '%thrombos%' AND NOT inhalt RLIKE 'aorta|halsven')) " & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `cwzl(33061)`=0 OR `komprzl(33076)`=0;"
'", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung IN ('33042','33042A')),0) `sonozl(33042/A)`" & vbCrLf & _

'"SELECT f.`pat_id` AS Pat_ID, ezp AS Zeitpunkt, art, Inhalt FROM " & aktf & " " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.`pat_id`= a.`pat_id` " & vbCrLf & _
"LEFT JOIN (SELECT * FROM " & vbCrLf & _
"(((SELECT pat_id AS pid,  LEFT(art,7) AS art, inhalt, zeitpunkt AS ezp, DATE(zeitpunkt) zp FROM `eintraege` e " & vbCrLf & _
"WHERE art IN ('duplex') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") AS eint " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS cwpat_id, fid AS cwfid, DATE(zeitpunkt) cwltag, leistung AS cw FROM `leistungen` WHERE leistung IN ('33061')) AS cw " & vbCrLf & _
"ON eint.pid = cw.cwpat_id AND eint.zp = cw.cwltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS komprpat_id, fid AS komprfid, DATE(zeitpunkt) komprltag, leistung AS kompr FROM `leistungen` WHERE leistung IN ('33076')) AS kompr " & vbCrLf & _
"ON eint.pid = kompr.komprpat_id AND eint.zp = kompr.komprltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS sonopat_id, fid AS sonofid, DATE(zeitpunkt) sonoltag, leistung AS sono FROM `leistungen` WHERE leistung IN ('33042','33042A')) AS sono " & vbCrLf & _
"ON eint.pid = sono.sonopat_id AND eint.zp = sono.sonoltag) " & vbCrLf & _
"as ges " & vbCrLf & "ON (ges.PID = f.`Pat_ID`) " & vbCrLf & _
"WHERE NOT ISNULL(ezp) AND (ISNULL(kompr) OR ISNULL(cw)) " & vbCrLf & _
"AND ((inhalt LIKE '%vene%' AND NOT inhalt LIKE '%mesent%' AND NOT inhalt LIKE '%halsven%') " & vbCrLf & _
"OR inhalt LIKE '%varik%' OR inhalt LIKE '%v.s.%' OR inhalt LIKE '%saph%' " & vbCrLf & _
"OR (inhalt LIKE '%thrombos%' AND NOT inhalt LIKE '%aorta%' AND NOT inhalt LIKE '%halsven%') OR inhalt LIKE '%reflux%') " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 50
 AWlf = AWlf + 1

' 129
' ktag fehlerhaft
AwN(AWlf) = "Potentiell fehlende Leistung für sonstigen Doppler/Duplex (33076) (vorher 46)"
sql(AWlf) = _
"SELECT f.pat_id PID, gesname(f.pat_id) PName, DATE(e.zeitpunkt) eZP" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33061'),0) `cwzl(33061)`" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung='33076'),0) `komprzl(33076)`" & vbCrLf & _
", COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND DATE(zeitpunkt)=DATE(e.ZeitPunkt) AND leistung IN ('33042','33042A')),0) `sonozl(33042/A)`" & vbCrLf & _
", Inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art IN ('doppler','dop','duplex','dup')" & vbCrLf & _
"AND (e.inhalt NOT RLIKE 'hals|arotis|aci|A.car.|ACC|Carot|A.d.|A.t.|biph|monoph|postst|Fußart|IDR|A.fem.|A. fem.|in Ruhe|Nieren|Vene|Varik|V.s.|Thrombos|reflux') " & vbCrLf & _
"AND e.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"WHERE NOT ISNULL(e.Pat_id)" & vbCrLf & _
"HAVING `cwzl(33061)`=0 AND `komprzl(33076)`=0 AND `sonozl(33042/A)`=0" & vbCrLf & _
"ORDER BY f.pat_id"

'sql(AWlf) = _
"SELECT f.`pat_id` AS Pat_ID, ezp AS Zeitpunkt, LEFT(art,7) AS art, Inhalt FROM " & aktf & " " & vbCrLf & _
"LEFT JOIN `anamnesebogen` a ON f.`pat_id`= a.`pat_id` " & vbCrLf & _
"LEFT JOIN (SELECT * FROM " & vbCrLf & _
"(((SELECT pat_id AS pid, art, inhalt, zeitpunkt AS ezp, DATE(zeitpunkt) zp FROM `eintraege` e " & vbCrLf & _
"WHERE art IN ('doppler','dop','duplex') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") AS eint " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS cwpat_id, fid AS cwfid, DATE(zeitpunkt) cwltag, leistung AS cw FROM `leistungen` WHERE leistung IN ('33061')) AS cw " & vbCrLf & _
"ON eint.pid = cw.cwpat_id AND eint.zp = cw.cwltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS komprpat_id, fid AS komprfid, DATE(zeitpunkt) komprltag, leistung AS kompr FROM `leistungen` WHERE leistung IN ('33076')) AS kompr " & vbCrLf & _
"ON eint.pid = kompr.komprpat_id AND eint.zp = kompr.komprltag) " & vbCrLf & _
"LEFT JOIN (SELECT pat_id AS sonopat_id, fid AS sonofid, DATE(zeitpunkt) sonoltag, leistung AS sono FROM `leistungen` WHERE leistung IN ('33042','33042A')) AS sono " & vbCrLf & _
"ON eint.pid = sono.sonopat_id AND eint.zp = sono.sonoltag) " & vbCrLf & _
"as ges " & vbCrLf & _
"ON (ges.PID = f.`Pat_ID`) " & vbCrLf & _
"WHERE NOT ISNULL(ezp) AND ISNULL(kompr) AND ISNULL(sono) AND ISNULL(cw) " & vbCrLf & _
"AND NOT (" & vbCrLf & _
"(inhalt LIKE '%hals%' OR inhalt LIKE '%arotis%'  OR inhalt LIKE '%ACI%' OR inhalt LIKE '%A.car.%'  OR inhalt LIKE '%ACC%' OR inhalt LIKE '%Carot%') " & vbCrLf & _
"OR (inhalt LIKE '%A.d.%' OR inhalt LIKE '%A.t.%'  OR inhalt LIKE '%biph%'  OR inhalt LIKE '%monoph%' OR inhalt LIKE '%postst%' OR inhalt LIKE '%Fußart%' OR inhalt LIKE '%IDR%' OR inhalt LIKE '%A. fem.%' OR inhalt LIKE '%A.fem.%' OR inhalt LIKE '%in Ruhe%' OR inhalt LIKE '%in Ruhe%') " & vbCrLf & _
"OR (inhalt LIKE '%Nieren%') " & vbCrLf & _
"OR (inhalt LIKE '%vene%' OR inhalt LIKE '%varik%' OR inhalt LIKE '%v.s.%' OR inhalt LIKE '%thrombos%' OR inhalt LIKE '%reflux%')) " & vbCrLf & _
"ORDER BY f.`pat_id`" ' AND ISNULL(ltag)
 mins(AWlf) = 7
 maxs(AWlf) = 120
 AWlf = AWlf + 1

' Gesundheitsuntersuchung: 01732 Glucose Cholesterin, Urin, Blutdruck ab 36.LJ mit uns als HA
' 130
AwN(AWlf) = "Gesundheitsuntersuchung 01732 (lauto) (vorher 50)"
sql(AWlf) = "SELECT n.Pat_id PID, gesname(n.pat_id) Name, patAlter(n.pat_id) PAlter, DATE(gluc.zeitpunkt) LEIDAT, '01732 dazu' LEIFEHLER, f.LANRID " & vbCrLf & _
" FROM `aktfvs` f " & vbCrLf & _
" LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
" LEFT JOIN eintraege gluc ON gluc.pat_id = f.pat_id AND gluc.art='kva' AND gluc.inhalt RLIKE '32094|32057' AND gluc.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
" LEFT JOIN eintraege chol ON chol.pat_id = f.pat_id AND chol.art='kva' AND chol.inhalt RLIKE '32060|32062' AND chol.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
" LEFT JOIN eintraege urin ON urin.pat_id = f.pat_id AND urin.art='urin' AND urin.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
" LEFT JOIN eintraege rr ON rr.pat_id = f.pat_id AND rr.zeitpunkt BETWEEN SUBDATE(" & qtAnf(FristS) & ",182) AND " & qtEnd(FristS) & " " & vbCrLf & _
" LEFT JOIN laborneu albu ON albu.pat_id=f.pat_id AND albu.zeitpunkt BETWEEN qanf()AND qend() AND albu.abkü LIKE '%albu%' " & vbCrLf & _
" LEFT JOIN `leistungen` l ON n.pat_id = l.pat_id AND l.leistung = '01732' AND DATEDIFF(" & qtAnf(FristS) & ",l.zeitpunkt)<365*3 " & vbCrLf & _
" WHERE n.kvnr IN ('" & KVNr & "','889690003','9333284903') AND DATEDIFF(" & qtAnf(FristS) & ",gebdat) > 365*35 AND ISNULL(leistung) " & vbCrLf & _
" AND NOT ISNULL(gluc.art) AND NOT ISNULL(chol.art) AND (NOT ISNULL(urin.art) OR NOT ISNULL(albu.fid)) " & vbCrLf & _
" GROUP BY n.pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 131
AwN(AWlf) = "Potentiell undokumentierte Op-Vorbereitung 31011, 31012, 31013 (vorher 56)"
' Kriterien eintraege ei koordinieren mit Liste 120 Urin
 sql(AWlf) = vbCrLf & _
"SELECT f.pat_id, gesname(f.pat_id),ei.zeitpunkt,ei.art,ei.inhalt, b.zeitpunkt, b.quelldatum, b.dokaend, b.name,ea.zeitpunkt,ea.inhalt" & vbCrLf & _
"FROM aktf f " & vbCrLf & _
"LEFT JOIN tmbrie b ON f.pat_id=b.pat_id AND name LIKE '%op%vorb%' " & vbCrLf & _
"AND (quelldatum BETWEEN " & lQAnfuEnd(FristS) & " OR dokaEND BETWEEN " & lQAnfuEnd(FristS) & ") " & vbCrLf & _
"LEFT JOIN eintraege ea ON f.pat_id=ea.pat_id AND ea.art='usal' AND ea.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"LEFT JOIN eintraege ei ON f.pat_id=ei.pat_id AND ei.inhalt LIKE '%op%vorb%' AND NOT ei.inhalt RLIKE 'op.*vorb(ei|eug)|Coloskopievorb|war.*op.*vorb|op.*vorb.*angefangen|bei .*op.*vorb' AND ei.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id and l.zeitpunkt BETWEEN qanf()AND qend() AND leistung IN ('31011','31012','31013') " & vbCrLf & _
"WHERE (NOT ISNULL(b.pfad) OR NOT ISNULL(ea.art) OR NOT ISNULL(ei.art)) AND ISNULL(l.leistung) AND NOT ISNULL(ei.inhalt);" & vbCrLf & _
""
mins(AWlf) = 10
maxs(AWlf) = 80
AWlf = AWlf + 1
 
 
' 132
 AwN(AWlf) = "Fehlende Leistung 03360 für Barthel- und TUG-Test (lauto) (vorher 73)"
 sql(AWlf) = "SELECT * FROM ( " & vbCrLf & _
 "SELECT v.pat_id PID, COALESCE(SUM(lz.lzahl),0)*COUNT(DISTINCT lz.id)/COUNT(lz.id) LZ, GROUP_CONCAT(DISTINCT lz.zeitpunkt) lzzp, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah'))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",GROUP_CONCAT(DISTINCT tug.zeitpunkt) tugzp " & vbCrLf & _
            ",GROUP_CONCAT(DISTINCT adl.zeitpunkt) adlzp " & vbCrLf & _
            ",(SELECT DATE(zeitpunkt) FROM leistungen WHERE pat_id=v.pat_id AND zeitpunkt>=qanf() ORDER BY zeitpunkt LIMIT 1) LEIDAT, '03360 dazu' LEIFEHLER,v.LANRID " & vbCrLf & _
            "FROM aktfvs v " & vbCrLf
 sql(AWlf) = sql(AWlf) & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id=tug.pat_id AND tug.art IN('TUG','247')AND tug.zeitpunkt=(SELECT MAX(zeitpunkt)FROM eintraege WHERE pat_id=v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id=adl.pat_id AND adl.art IN('ADL','284')AND adl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0123]|^G20' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND (pfld.gicd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.gicd RLIKE '^F0[012]') " & vbCrLf & _
            "LEFT JOIN leistungen l ON l.pat_id=v.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() AND l.leistung = '03360' " & vbCrLf & _
            "LEFT JOIN leistungen lz ON v.pat_id = lz.pat_id AND lz.leistung = '03360' AND lz.zeitpunkt BETWEEN " & Khtsfl & " " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) " & vbCrLf & _
            "AND NOT ISNULL(tug.inhalt) AND NOT ISNULL(adl.inhalt) " & vbCrLf & _
            "AND ISNULL(l.leistung) " & vbCrLf & _
            "AND NOT ISNULL(pfld.icd) " & vbCrLf & _
            "GROUP BY v.pat_id) i " & vbCrLf & _
            "WHERE lz<2"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
'             ",IF(tug.zeitpunkt>" & qtAnf(FristS) & ",DATE(tug.zeitpunkt),DATE(" & qtEnd(FristS) & ")) LEIDAT, '03360 dazu' LEIFEHLER,v.LANRID "
'            "LEFT JOIN leistungen lz ON v.pat_id = lz.pat_id AND lz.leistung = '03360' AND YEAR(lz.zeitpunkt)=YEAR(SUBDATE(NOW(),INTERVAL " & FristS & " DAY)) "

' 133
 AwN(AWlf) = "Leistung 03360 für Barthel- und TUG-Test bei Patienten ohne Pflegediagnose (vorher 74)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah'))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            "FROM aktfvs v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|G20' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND (pfld.gicd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.gicd RLIKE '^F0[0-3]') " & vbCrLf & _
            "LEFT JOIN leistungen l ON l.pat_id=v.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() AND l.leistung = '03360' " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) " & vbCrLf & _
            "AND NOT ISNULL(tug.inhalt) AND NOT ISNULL(adl.inhalt) " & vbCrLf & _
            "AND NOT ISNULL(l.leistung) " & vbCrLf & _
            "AND ISNULL(pfld.icd) " & vbCrLf & _
            "GROUP BY pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' AwN(AWlf) = "fehlende Medikamentenplanaktualisierung im aktuellen Quartal für Ziffer 03362"
' sql(AWlf) = "SELECT v.pat_id, gesnameg(v.pat_id) Name " & vbCrLf & _
'            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND art = 'tk') tk " & vbCrLf & _
'            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND art = 'gs') gs " & vbCrLf & _
'            ",CONCAT(LPAD(MID(tug.inhalt,1,INSTR(tug.inhalt,' ')),4,' '), 's/ ', LPAD(MID(adl.inhalt,17),3,' '),' P') Tests " & vbCrLf & _
'            ",k.konz Konz " & vbCrLf & _
'            "FROM aktfvs v " & vbCrLf & _
'            "LEFT JOIN faelle f ON v.fid = f.fid " & vbCrLf & _
'            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
'            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND DATE(tug.zeitpunkt) BETWEEN SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) AND " & qtEnd(FristS) & " " & vbCrLf & _
'            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND DATE(adl.zeitpunkt) BETWEEN SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) AND " & qtEnd(FristS) & " " & vbCrLf & _
'            "LEFT JOIN diagnosen dd ON v.pat_id = dd.pat_id AND (dd.icd RLIKE '^F0[0123]' OR dd.icd LIKE 'G20%') AND dd.diagsicherheit NOT IN ('A','V') AND COALESCE(dd.Dggel,0)=0 " & vbCrLf & _
'            "LEFT JOIN diagnosen pfld ON v.pat_id = pfld.pat_id AND (pfld.icd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R41.3','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.icd RLIKE '^F0[012]') AND pfld.diagsicherheit NOT IN ('A','V') AND COALESCE(pfld.Dggel,0)=0 " & vbCrLf & _
'            "LEFT JOIN leistungen l ON v.fid = l.fid  AND l.leistung = '03362' " & vbCrLf & _
'"LEFT JOIN (SELECT pat_id, COUNT(zp) konz, GROUP_CONCAT(DATE_FORMAT(zp,'%e.%c.') SEPARATOR '•') czp FROM ( " & vbCrLf & _
'"(SELECT pat_id, DATE(zeitpunkt) zp FROM eintraege WHERE art IN (" & artspezG & ")  AND art <> 'andm' AND art <> 'anal' " & vbCrLf & _
'"AND DATE(zeitpunkt) BETWEEN " & lqanfuend(FristS) & ") " & vbCrLf & _
'"UNION (SELECT pat_id, DATE(zeitpunkt) zp FROM rezepteintraege WHERE DATE(zeitpunkt) BETWEEN " & lqanfuend(FristS) & ")) k GROUP BY pat_id ORDER BY zp) k ON v.pat_id = k.pat_id " & vbCrLf & _
'            "LEFT JOIN medplan mp ON v.pat_id = mp.pat_id AND DATE(mp.zeitpunkt) BETWEEN " & lqanfuend(FristS) & " " & vbCrLf & _
'            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) " & vbCrLf & _
'            "AND NOT ISNULL(tug.inhalt) AND NOT ISNULL(adl.inhalt) " & vbCrLf & _
'            "AND NOT ISNULL(pfld.icd) " & vbCrLf & _
'            "AND ISNULL(mp.zeitpunkt) " & vbCrLf & _
'            "GROUP BY pat_id;"
' mins(AWlf) = 10
' maxs(AWlf) = 60
' AWlf = AWlf + 1
'
 ' 134
 AwN(AWlf) = "Fehlende 03362 (hausärztlich-geriatrischer Betreuungskomplex) (lauto) (vorher 75)"
 sql(AWlf) = "SELECT v.pat_id PID, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='ah'AND ersteller='ah'))) ah " & vbCrLf & _
            ",CONCAT(pfld.icd,'(',pfld.diagtext,')') `Pflegediagnose`" & vbCrLf & _
            ", '03362 dazu' LEIFEHLER, DATE(letzt)+INTERVAL IF((SELECT 1 FROM leistungen WHERE pat_id=v.pat_id AND leistung IN ('03000','03001','03002','03003','03004','03005','03030') AND DATE(zeitpunkt)=DATE(letzt)) IS NULL,0,1) DAY LEIDAT " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",v.koz KZahl, czp Kontakte, v.art KontaktArt, v.LANRID " & vbCrLf & _
            "FROM aktfkvs v " & vbCrLf & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagnosen dd ON v.pat_id = dd.pat_id AND (dd.icd RLIKE '^F0[0123]' OR dd.icd LIKE 'G20%') AND dd.diagsicherheit NOT IN ('A','V','Z') " & vbCrLf & _
            "LEFT JOIN diagnosen pfld ON v.pat_id = pfld.pat_id AND (pfld.icd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.icd RLIKE '^F0[012]') AND pfld.diagsicherheit NOT IN ('A','V','Z') " & vbCrLf & _
            "LEFT JOIN leistungen l ON v.pat_id = l.pat_id  AND l.leistung = '03362' AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf
sql(AWlf) = sql(AWlf) & _
            "LEFT JOIN medplan mp ON v.pat_id = mp.pat_id " & vbCrLf & _
            "AND ((mp.zeitpunkt >= (SELECT MAX(zeitpunkt) FROM medplan WHERE zeitpunkt < " & qtAnf(FristS) & " AND pat_id = f.pat_id))  OR (SELECT MAX(zeitpunkt) FROM medplan WHERE zeitpunkt < " & qtAnf(FristS) & " AND pat_id = f.pat_id) IS NULL AND mp.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
            "AND ((mp.zeitpunkt <  (SELECT MIN(zeitpunkt) FROM medplan WHERE zeitpunkt > " & qtEnd(FristS) & " AND pat_id = f.pat_id))  OR (SELECT MIN(zeitpunkt) FROM medplan WHERE zeitpunkt > " & qtEnd(FristS) & " AND pat_id = f.pat_id) IS NULL AND mp.zeitpunkt<=" & qtEnd(FristS) & ") " & vbCrLf & _
            "WHERE (DATEDIFF(" & qtAnf(FristS) & ", n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) " & vbCrLf & _
            "AND NOT ISNULL(tug.inhalt) AND NOT ISNULL(adl.inhalt) " & vbCrLf & _
            "AND ISNULL(l.leistung) " & vbCrLf & _
            "AND NOT ISNULL(pfld.icd) " & vbCrLf & _
            "AND koz > 1 " & vbCrLf & _
            "GROUP BY v.pat_id;"
' Änderung 13.7.14
'            "AND NOT ISNULL(mp.zeitpunkt) " & vbCrLf & _
'
'"LEFT JOIN (SELECT pat_id, COUNT(DISTINCT DATE(zp)) koz, GROUP_CONCAT(DATE_FORMAT(zp,'%e.%c.') ORDER BY zp SEPARATOR '•') czp FROM ( " & vbCrLf & _
"(SELECT pat_id, DATE(zeitpunkt) zp FROM eintraege WHERE art IN (" & artspezG & ")  AND art <> 'andm' AND art <> 'anal' " & vbCrLf & _
"AND zeitpunkt BETWEEN " & lqanfuend(FristS) & ") " & vbCrLf & _
"UNION (SELECT pat_id, DATE(zeitpunkt) zp FROM rezepteintraege WHERE zeitpunkt BETWEEN " & lqanfuend(FristS) & ")) k GROUP BY pat_id ORDER BY zp) k ON v.pat_id = k.pat_id " & vbCrLf & _

 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 135
 AwN(AWlf) = "möglicherweise unberechtigte 03362 (vorher 76)"
 sql(AWlf) = "SELECT v.pat_id, gesname(v.pat_id) Name, patalter(v.pat_id) PAlter " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk'))) tk " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs'))) gs " & vbCrLf & _
            ",(SELECT COUNT(art) FROM eintraege WHERE pat_id = v.pat_id AND (art = 'ah'OR(art='ah'AND ersteller='ah'))) ah " & vbCrLf & _
            ",CONCAT(LPAD(REGEXP_REPLACE(tug.inhalt,'.*benötigte Zeit\: *([0-9]+) Sek.*','\\1'),3,' '),'s/ ', LPAD(REGEXP_REPLACE(adl.inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')) Tests " & vbCrLf & _
            ",CONCAT(IF(v.koz+(IF(ISNULL(l2.leistung),0,1))<=1,'Kontaktzahl ','            '),IF(ISNULL(pfld.icd),'Diagnose ','         ') ) Fehler " & vbCrLf & _
            ", h.letzter ltzt_HbA1c,v.czp Kontakte,v.art KontaktArt " & vbCrLf & _
            "FROM aktfkvs v " & vbCrLf
sql(AWlf) = sql(AWlf) & _
            "JOIN faelle f USING (fid)" & vbCrLf & _
            "LEFT JOIN leistungen l2 ON l2.pat_id=v.pat_id AND l2.zeitpunkt BETWEEN qanf()AND qend() AND l2.leistung LIKE '03221%' " & vbCrLf & _
            "LEFT JOIN namen n ON v.pat_id = n.pat_id " & vbCrLf & _
            "LEFT JOIN eintraege tug ON v.pat_id = tug.pat_id AND tug.art IN('TUG','247') AND tug.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('TUG','247'))" & vbCrLf & _
            "LEFT JOIN eintraege adl ON v.pat_id = adl.pat_id AND adl.art IN('ADL','284') AND adl.zeitpunkt = (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id = v.pat_id AND art IN('ADL','284'))" & vbCrLf & _
            "LEFT JOIN diagview dd ON v.pat_id = dd.pat_id AND dd.gicd RLIKE '^F0[0-3]|G20\.[12]|G30' " & vbCrLf & _
            "LEFT JOIN diagview pfld ON v.pat_id = pfld.pat_id AND pfld.gICD RLIKE '^F0[0-3]|^F3|^F45\.41|^F69|^F79|^G20\.[12]|^G30\.9|^M62\.50|^R1[35]|^R26\.[38]|^R29\.6|^R32|^R4[12]|^R46\.4|^R52\.[12]|^R5[34]|^R63\.4|^R68\.8|^Z74\.[09]'" & vbCrLf & _
            "LEFT JOIN leistungen l ON l.pat_id=v.pat_id AND l.zeitpunkt BETWEEN qanf()AND qend() AND l.leistung = '03362' " & vbCrLf & _
            "LEFT JOIN medplan mp ON v.pat_id = mp.pat_id " & vbCrLf & _
            "AND ((mp.zeitpunkt >= (SELECT MAX(zeitpunkt) FROM medplan WHERE zeitpunkt < " & qtAnf(FristS) & " AND pat_id = f.pat_id))  OR (SELECT MAX(zeitpunkt) FROM medplan WHERE zeitpunkt < " & qtAnf(FristS) & " AND pat_id = f.pat_id) IS NULL AND mp.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
            "AND ((mp.zeitpunkt <  (SELECT MIN(zeitpunkt) FROM medplan WHERE zeitpunkt > " & qtEnd(FristS) & " AND pat_id = f.pat_id))  OR (SELECT MIN(zeitpunkt) FROM medplan WHERE zeitpunkt > " & qtEnd(FristS) & " AND pat_id = f.pat_id) IS NULL AND mp.zeitpunkt<=" & qtEnd(FristS) & ") " & vbCrLf & _
            "LEFT JOIN lHbA1c h ON h.Pat_id = v.pat_id " & vbCrLf & _
            "WHERE (DATEDIFF(SUBDATE(" & qtAnf(FristS) & ",INTERVAL " & FristS & " DAY), n.GebDat) > 70 * 365 OR NOT ISNULL(dd.ICD)) " & vbCrLf & _
            "AND NOT ISNULL(l.leistung) " & vbCrLf & _
            "AND ((ISNULL(pfld.icd) AND COALESCE(h.letzter,0)<=7.5) " & vbCrLf & _
            "OR v.koz+(IF(ISNULL(l2.leistung),0,1)) <= 1 ) GROUP BY pat_id;"
'"LEFT JOIN (SELECT pat_id, COUNT(DISTINCT DATE(zp)) koz, GROUP_CONCAT(DATE_FORMAT(zp,'%e.%c.') SEPARATOR '•') czp FROM ( " & vbCrLf & _
"(SELECT pat_id, DATE(zeitpunkt) zp FROM eintraege WHERE art IN (" & artspezG & ")  AND art <> 'andm' AND art <> 'anal' " & vbCrLf & _
"AND zeitpunkt BETWEEN " & lqanfuend(FristS) & ") " & vbCrLf & _
"UNION (SELECT pat_id, DATE(zeitpunkt) zp FROM rezepteintraege WHERE zeitpunkt BETWEEN " & lqanfuend(FristS) & ")) k GROUP BY pat_id ORDER BY zp) k ON v.pat_id = k.pat_id " & vbCrLf & _

' Änderungen 13.7.14
'            ",CONCAT(IF(k.koz<=1,'Kontaktzahl ','            '),IF(ISNULL(pfld.icd),'Diagnose ','         '),IF(ISNULL(mp.zeitpunkt),'Medplan ','        ') ) Fehler " & vbCrLf & _
'
'            "AND NOT ISNULL(tug.inhalt) AND NOT ISNULL(adl.inhalt) " & vbCrLf & _
'
'            "OR ISNULL(mp.zeitpunkt) " & vbCrLf & _
'
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 136
 AwN(AWlf) = "Doppelte 03362 (vorher 77)"
 sql(AWlf) = "SELECT * FROM (" & vbCrLf & _
"SELECT COALESCE(SUM(l.lzahl),0) zahl, f.pat_id, gesname(f.pat_id) Gesname, patalter(f.pat_id) PAlter, l.zeitpunkt, l.leistung " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id = f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() AND leistung = '03362' " & vbCrLf & _
"WHERE NOT ISNULL(l.Leistung) " & vbCrLf & _
"GROUP BY f.pat_id) i " & vbCrLf & _
"WHERE zahl<>1;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 137
' ktag fehlerhaft
 AwN(AWlf) = "03362 gleichzeitg mit 03000, 03001, 03002, 03003, 03004, 03005 (vorher 78)"
 sql(AWlf) = "SELECT f.pat_id, gesname(f.pat_id) Gesname, DATE(l.zeitpunkt), l.leistung, l1.leistung " & vbCrLf & _
            ",f.koz KZahl, czp Kontakte,f.art Kontaktart " & vbCrLf & _
            "FROM aktfkvs f " & vbCrLf & _
"       LEFT JOIN leistungen l ON f.pat_id = l.pat_id AND l.leistung = '03362' AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"       LEFT JOIN leistungen l1 ON f.pat_id = l1.pat_id AND l1.leistung IN ('03000','03001','03002','03003','03004','03005','03030') AND l1.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
" WHERE NOT ISNULL(l.leistung) AND NOT ISNULL(l1.leistung) AND DATE(l.zeitpunkt)=DATE(l1.zeitpunkt); "
 
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 138
AwN(AWlf) = "03360 öfter als 2 x pro Krankheitsfall (vorher 79)"
 sql(AWlf) = "SELECT * FROM ( " & vbCrLf & _
"SELECT COALESCE(SUM(l.lzahl),0) zahl, f.pat_id, gesname(f.pat_id) Gesname, patalter(f.pat_id) PAlter, l.zeitpunkt, l.leistung " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"       LEFT JOIN leistungen l ON f.pat_id = l.pat_id AND l.leistung = '03360' AND l.zeitpunkt BETWEEN " & Khtsfl & " " & vbCrLf & _
"       LEFT JOIN leistungen la ON f.pat_id = la.pat_id AND la.leistung = '03360' AND la.zeitpunkt BETWEEN " & lQAnfuEnd(Verspätung) & " " & vbCrLf & _
"WHERE NOT ISNULL(la.Leistung) " & vbCrLf & _
"GROUP BY f.pat_id) i " & vbCrLf & _
"WHERE zahl>2;"
'"       LEFT JOIN leistungen l ON f.pat_id = l.pat_id AND leistung = '03360' AND YEAR(l.zeitpunkt)=YEAR(SUBDATE(" & qtAnf(FristS) & ",INTERVAL " & FristS & " DAY)) "

mins(AWlf) = 10
maxs(AWlf) = 60
AWlf = AWlf + 1

' 139
' ktag fehlerhaft
AwN(AWlf) = "Fehlende 32150 für Toponintest (lauto) (vorher 84)"
sql(AWlf) = "" & vbCrLf & _
"SELECT f.pat_id PID, gesname(f.pat_id),DATE(e.zeitpunkt) LEIDAT, '32150 dazu' LEIFEHLER, e.Art, e.Inhalt, f.LANRID " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.zeitpunkt BETWEEN qanf() AND qend() AND (inhalt RLIKE 'trop-?test' OR art='trop')" & vbCrLf & _
"LEFT JOIN leistungen l ON f.pat_id = l.pat_id AND leistung = '32150' AND DATE(l.zeitpunkt)=DATE(e.zeitpunkt) " & vbCrLf & _
"WHERE NOT ISNULL(e.inhalt) AND ISNULL(l.leistung)"
mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
#If vor31032018 Then
' zum 31.3.18 gestrichen: 99300
sql = _
"SELECT f.pat_id, gesname(f.pat_id),dm.icd Dm,ni.icd NI,ne.icd Nephr, lGFR.letzter eGFR, copd.icd COPD, ri.icd RI " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN diagview dm ON f.pat_id = dm.pat_id AND dm.gicd RLIKE '^E1[013]\.2[01]|^E1[013]\.7[23]'" & vbCrLf & _
"LEFT JOIN diagview ni ON f.pat_id = ni.pat_id AND ni.gicd RLIKE '^N18\.[12]'" & vbCrLf & _
"LEFT JOIN diagview ne ON f.pat_id = ne.pat_id AND ne.gicd = 'N08.3' " & vbCrLf & _
"LEFT JOIN diagview copd ON f.pat_id = copd.pat_id AND copd.gicd RLIKE '^J44\.[018][01]'  " & vbCrLf & _
"LEFT JOIN diagview ri ON f.pat_id = ri.pat_id AND ri.gicd RLIKE '^J96\.1[019]' " & vbCrLf & _
"LEFT JOIN lGFR ON f.pat_id = lGFR.pat_id " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id = f.pat_id AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND l.leistung = '99300' " & vbCrLf & _
"WHERE (NOT ISNULL(dm.ICD) AND NOT ISNULL(ni.ICD) AND ISNULL(l.Leistung)) " & vbCrLf & _
"   OR (NOT ISNULL(copd.icd) AND NOT ISNULL(ri.icd)) " & vbCrLf & _
"GROUP BY f.pat_id; "
#End If

' 140
' ktag fehlerhaft
 AwN(AWlf) = "Fehlende Leistungsziffern 01747 und 01748 Screening auf Bauchaortenaneurysmen (vorher 108)"
 sql(AWlf) = vbCrLf & _
 "SELECT f.pat_id, gesname(f.pat_id) Name, DATE(e.zeitpunkt) Zp, n.gebdat, l1.leistung L1,l2.leistung L2 " & vbCrLf & _
 "FROM aktfvs f " & vbCrLf & _
 "LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
 "LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND e.art = 'sono'  AND e.inhalt LIKE '%aort%' " & vbCrLf & _
 "LEFT JOIN leistungen l1 ON l1.pat_id = f.pat_id AND l1.leistung = '01747' " & vbCrLf & _
 "LEFT JOIN leistungen l2 ON l2.pat_id = f.pat_id AND l2.leistung = '01748' " & vbCrLf & _
 "WHERE true " & vbCrLf & _
 "AND NOT ISNULL(e.art) " & vbCrLf & _
 "AND f.schgr<>90 " & vbCrLf & _
 "AND n.geschlecht='m' " & vbCrLf & _
 "AND (ISNULL(l1.leistung) OR ISNULL(l2.leistung)) " & vbCrLf & _
 "AND ADDDATE(n.gebdat,INTERVAL 65 YEAR) < e.zeitpunkt " & vbCrLf & _
 "AND COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='01748'),0) <1 "
' "AND patalter(f.pat_id) > 64 "
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 141
 AwN(AWlf) = "Mehr als 1 Leistung 01747 im Leben (vorher 118)"
 sql(AWlf) = vbCrLf & _
"SELECT * FROM (SELECT " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='01747'),0) Geszahl, " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND leistung='01747' AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & "),0) Aktzahl, " & vbCrLf & _
"gesname(f.pat_id), f.pat_id " & vbCrLf & _
", (SELECT GROUP_CONCAT(zeitpunkt SEPARATOR ', ') FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='01747' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Zpkte " & vbCrLf & _
"FROM aktfvs f) i WHERE geszahl>1 AND aktzahl>0;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 142
 AwN(AWlf) = "Mehr als 1 Leistung 01748 im Leben, darunter >0 aktuell (vorher 119)"
 sql(AWlf) = vbCrLf & _
"SELECT " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='01748'),0) Geszahl, " & vbCrLf & _
"COALESCE((SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='01748' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & "),0) Aktzahl, " & vbCrLf & _
"gesname(f.pat_id), f.pat_id " & vbCrLf & _
", (SELECT GROUP_CONCAT(zeitpunkt SEPARATOR ', ') FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung='01748' AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Zpkte " & vbCrLf & _
"FROM aktfvs f HAVING geszahl>1 AND aktzahl>0;"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 143
' ktag fehlerhaft
'     CASE "LANRID": lanrda = True
'     CASE "LEIFEHLER": leida = True
'     CASE "LEIDAT": ldda = True
'     CASE "PID": pda = True
 AwN(AWlf) = "Fehlende 32033 oder Op-Vorb. 31010-31013 für Urinuntersuchungen (urin): (lauto) (vorher 120)"
 ' Kriterien eintraege ei koordinieren mit Liste 56 Op-Vorbereitung
 'sql(AWlf) = "SELECT * FROM (SELECT f.pat_id, gesnameg(f.pat_id) Name, DATE(e.zeitpunit) Tag, COUNT(art) `Urin-Zahl`, SUM(IF(LAnzl='',1,LAnzl)) `Leistungs-Zahl` FROM aktfvs f LEFT JOIN namen n ON f.pat_id = n.pat_id LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND DATE(e.zeitpunkt) BETWEEN " & lqanfuend(FristS) & " AND art IN ('urin')  LEFT JOIN leistungen l ON f.fid = l.fid AND leistung IN ('32033','31010','31011','31012','31013') AND DATE(l.zeitpunkt)=DATE(e.zeitpunkt) WHERE NOT ISNULL(art) GROUP BY f.pat_id, DATE(e.zeitpunkt)) i WHERE `Urin-Zahl`<>`Leistungs-Zahl`"
 sql(AWlf) = vbCrLf & _
" SELECT e.pat_id PID, gesnameg(e.pat_id) Name, DATE(e.zeitpunkt) LEIDAT, e.inhalt " & vbCrLf & _
" , IF(ISNULL(ea.art) AND ISNULL(ei.art),'32033 dazu','31010-13 fehlt') LEIFEHLER " & vbCrLf & _
" , f.LANRID " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id = f.pat_id AND e.art='Urin' AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id = e.pat_id AND l.leistung IN ('32033','31010','31011','31012','31013') " & vbCrLf & _
      "AND DATE(e.zeitpunkt) = DATE(l.zeitpunkt) " & vbCrLf & _
"LEFT JOIN eintraege ea ON f.pat_id=ea.pat_id AND ea.art='usal' AND ea.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"LEFT JOIN eintraege ei ON f.pat_id=ei.pat_id AND ei.inhalt LIKE '%op%vorb%' AND NOT ei.inhalt RLIKE 'op.*vorbei|Coloskopievorb|war.*op.*vorb|op.*vorb.*angefangen|bei .*op.*vorb' AND ei.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
"WHERE e.art='urin' AND ISNULL(l.leistung) " & vbCrLf & _
"GROUP BY e.pat_id "

mins(AWlf) = 10
maxs(AWlf) = 60
AWlf = AWlf + 1

' 144
' ktag fehlerhaft
AwN(AWlf) = "Möglicherweise vergessene 03230 (vorher 135)"
sql(AWlf) = "" & _
"SELECT IF(zeitpunkt=minzp,pat_id,'') PID, gesname(pat_id), DATE_FORMAT(zeitpunkt,'%d.%m.%y %k:%i') ZeitP, DATE_FORMAT(minzp,'%k:%i') MinZP, DATE_FORMAT(maxzp,'%k:%i') MaxZP, Minuten, Art,Inhalt FROM (" & vbCrLf & _
"SELECT ROUND((to_seconds(maxzp) - to_seconds(minzp))/60) Minuten, i.* FROM (" & vbCrLf & _
"SELECT f.pat_id, e.zeitpunkt, " & vbCrLf & _
"(SELECT MIN(zeitpunkt) FROM eintraege ez WHERE ez.pat_id = e.pat_id AND " & vbCrLf & _
"  ez.art IN (" & artSpezBerat & ") AND ez.inhalt NOT LIKE '%nachgefordert%' AND " & vbCrLf & _
"  DATE(ez.zeitpunkt) = DATE(e.zeitpunkt)) minzp, " & vbCrLf & _
"(SELECT MAX(zeitpunkt) FROM eintraege ez WHERE ez.pat_id = e.pat_id AND " & vbCrLf & _
"  ez.art IN (" & artSpezBerat & ") AND ez.inhalt NOT LIKE '%nachgefordert%' AND " & vbCrLf & _
"  DATE(ez.zeitpunkt)=DATE(e.zeitpunkt)) maxzp, " & vbCrLf & _
"e.Art , e.Inhalt, l.Leistung " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id = f.pat_id AND e.qt = aktq() AND art IN (" & artSpezBerat & ") AND inhalt NOT LIKE '%nachgefordert%' " & vbCrLf & _
"LEFT JOIN leistungen l ON e.pat_id = l.pat_id AND DATE(l.zeitpunkt)=DATE(e.zeitpunkt) AND l.leistung = '03230' " & vbCrLf & _
") i " & vbCrLf & _
") j WHERE Minuten>9 AND ISNULL(leistung) " & vbCrLf & _
"ORDER BY pat_id, zeitpunkt"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 145
' ktag fehlerhaft
AwN(AWlf) = "PCO-Syndrom für 03230 (vorher 136)"
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id), DATE(ez.zeitpunkt) FROM aktfv f " & vbCrLf & _
"LEFT JOIN diagnosen d ON d.Pat_id=f.pat_id AND d.diagdatum > qanf()-100 AND d.icd = 'E28.2'" & vbCrLf & _
"LEFT JOIN eintraege ez ON ez.pat_id=f.pat_id AND ez.ZeitPunkt BETWEEN qanf() AND qend() AND " & vbCrLf & _
"ez.art IN (" & artSpezBerat & ") AND ez.inhalt NOT LIKE '%nachgefordert%' " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.Leistung = '03230' AND DATE(l.zeitpunkt) = DATE(ez.zeitpunkt) " & vbCrLf & _
"WHERE NOT ISNULL(DiagDatum) AND IsNull(Leistung) " & vbCrLf & _
"GROUP BY f.pat_id, DATE(ez.zeitpunkt);"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 146, neuView
AwN(AWlf) = "Fehlende Leistungen für 24-h-RR und Lufu"
sql(AWlf) = "" & _
"SELECT  f.pat_id, gesname(f.pat_id) PName, IF(e.Inhalt RLIKE '24 h','03324','03330') Soll_Leistung, e.zeitpunkt, e.art, e.inhalt" & vbCrLf & _
"FROM aktfv f" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.ZeitPunkt BETWEEN qanf() AND qend() AND e.art IN ('LZRR','EKG','LUFU') AND e.inhalt RLIKE '24 h|SOLL *Messwert'" & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND date(l.zeitpunkt)=DATE(e.zeitpunkt) AND ((e.Inhalt RLIKE '24 h' AND l.leistung='03324') OR (e.inhalt RLIKE 'soll *messwert' AND l.leistung='03330'))" & vbCrLf & _
"WHERE e.Pat_id IS NOT NULL And l.Pat_id IS NULL" & vbCrLf & _
"GROUP BY f.pat_id, DATE(e.zeitpunkt);"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 147, neuView
AwN(AWlf) = "Impfungen"
sql(AWlf) = "ü"
 AWlf = AWlf + 1

#If vor0415 Then
'#If ebmalt THEN
' AwN(AWlf) = "Originalschein mit fehlender oder falscher Grundpauschale 03110,03111,03112"
' sql(AWlf) = _
' "SELECT fpat_id pat_id, ffid fid, name, ICD, Schgr, MinSchgr, LAlter `Alter`, Grundl, Chron, `Zahl(Grundleist)`, `Zahl(Chron'pausch.)` FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS Fpat_id, f.fid AS FFid, LEFT(CONCAT(n.Nachname, ' ', n.Vorname),25) AS Name, timestampdiff(YEAR, n.gebdat," & qtAnf(FristS) & ") AS lalter, f.Schgr, MIN(MSchGr.Schgr) MinSchGr, gl.leistung AS Grundl, l.leistung AS chron, COUNT(DISTINCT pat_gl.pat_id) `Zahl(Grundleist)`, COUNT(DISTINCT pat_l.pat_id) `Zahl(Chron'pausch.)`, ICD FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN `leistungen` gl ON f.fid = gl.fid AND gl.leistung IN ('03110','03111','03112','03120','03121','03122') " & vbCrLf & _
' "LEFT JOIN `leistungen` l ON f.fid = l.fid AND l.leistung IN ('03212') " & vbCrLf & _
' "LEFT JOIN `leistungen` pat_gl ON pat_gl.leistung IN ('03110','03111','03112','03120','03121','03122') AND pat_gl.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND f.pat_id = pat_gl.pat_id " & vbCrLf & _
' "LEFT JOIN `leistungen` pat_l ON pat_l.leistung IN ('03212') AND pat_l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND pat_l.pat_id = f.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "LEFT JOIN `faelle` MSchGr ON f.pat_id = MSchGr.pat_id AND MSchGr.quartal = '" & AktQ & "' " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON d.pat_id = f.pat_id AND (d.obdauer <> 0 OR d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & ") AND d.icd REGEXP '^E1[0-4]\.' AND d.diagsicherheit <> 'A' AND COALESCE(d.Dggel,0)=0 " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "GROUP BY f.fid ORDER BY gl.leistung DESC) AS innen) AS innen " & vbCrLf & _
' "WHERE " & vbCrLf & _
' "schgr = 0  AND ((ISNULL(grundl) OR grundl NOT IN ('03110','03111','03112') OR " & vbCrLf & _
' "(lalter < 6 AND grundl <> '03110') OR ((lalter >= 6 AND lalter < 59) AND grundl <> '03111') OR (lalter >= 59 AND grundl <> '03112'))) " & vbCrLf & _
' "ORDER BY pat_id "
' ' AND NOT f.goäkatnr IN (40,41)
''  "WHERE (ct > 1 AND NOT NOT ISNULL(`Zahl(Grundleist)`) AND ISNULL(chron)) " & vbCrLf & _
'' "WHERE (ISNULL(leistung)) AND ct > 1 AND NOT ISNULL(icd) "
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
'#Else 'lalter
'
' AwN(AWlf) = "Originalschein oder Üw vom Facharzt mit fehlender oder falscher Grundpauschale 0300x"
' sql(AWlf) = "SELECT * FROM (" & vbCrLf & _
'             "SELECT Pat_ID, Name, riLei `erwartete richt.Ziff`, i.gesleist `vorh.Ziff`, i.gpzahl  `Zahl d.vorh.Ziff`, Schgr, IF(schgr<>0,CONCAT_WS(' | ',IF(ISNULL(übwlanr),'',übwlanr),IF(ISNULL(üw),'',üw),IF(ISNULL(NLArt),'',nlart),IF(ISNULL(lanr2),'',LANR2),IF(ISNULL(Üw2),'',üw2),IF(ISNULL(nlart2),'', nlart2) ),'') `Üw`, i.falter `Alter n.Fall`, i.fanf `Fallanf` " & vbCrLf & _
'             "FROM (SELECT Name, übwlanr, üw, NLArt, LANR2,Üw2,nlart2, i.Pat_ID, Schgr, i.leistung,IF(lalter<4,'03001',IF(lalter<18,'03002',IF(lalter<54,'03003',IF(lalter<75,'03004','03005')))) riLei, i.gesleist, i.gpzahl, i.lalter, i.falter, i.fanf " & vbCrLf & _
'             "FROM (SELECT gesname(f.pat_id) Name, vp.übwlanr, vp.üw, vp.NLArt,vp.LANR2,vp.Üw2,vp.nlart2, op.leistung, GROUP_CONCAT(gp.leistung) gesleist, SUM(gp.lzahl) gpzahl, f.pat_id, gf.schgr, timestampdiff(YEAR,n.gebdat,gp.zeitpunkt) lalter, timestampdiff(YEAR,n.gebdat,IF(gf.fanf>gf.bhfb,gf.fanf,gf.bhfb)) falter, DATE(IF(gf.fanf>gf.bhfb,gf.fanf,gf.bhfb)) fanf " & vbCrLf & _
'             "FROM aktfvs f " & vbCrLf & _
'             "LEFT JOIN faelle gf ON f.fid = gf.fid " & vbCrLf & _
'             "LEFT JOIN leistungen op ON f.fid = op.fid AND op.leistung LIKE '0300%' " & vbCrLf & _
'             "LEFT JOIN leistungen gp ON f.fid = gp.fid AND (gp.leistung RLIKE '^030[01]' OR gp.leistung RLIKE '^132[12][012]') " & vbCrLf & _
'             "LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
'             "LEFT JOIN ( " & vbCrLf & _
'             "SELECT f.pat_id, f.fid, gesname(f.pat_id) Pat, vpau, f.übwlanr, CONCAT_WS(',',arzt.nachname,arzt.vorname) ÜW, IF(nlart.niederlassungsart='Hausarzt','h',LEFT(nlart.niederlassungsart,1)) NLArt, arzt2.lanr LANR2, CONCAT_WS(',',arzt2.nachname,arzt2.vorname) ÜW2, IF(nlart2.niederlassungsart='Hausarzt','h',LEFT(nlart2.niederlassungsart,1)) nlart2 FROM (`aktfvs` JOIN `faelle` f ON `aktfvs`.fid = f.fid AND f.schgr BETWEEN 20 AND 29) LEFT JOIN (haerzte.arzt JOIN haerzte.nlart ON arzt.nlart_id = nlart.idnlart LEFT JOIN (haerzte.arzt_has_bs ahb LEFT JOIN haerzte.arzt_has_bs ahb2 ON ahb.bs_id = ahb2.bs_id AND ahb.arzt_id <> ahb2.arzt_id JOIN  (haerzte.arzt arzt2 JOIN haerzte.nlart nlart2 ON arzt2.nlart_id = nlart2.idnlart) ON ahb2.arzt_id = arzt2.idarzt   ) ON arzt.idarzt = ahb.arzt_id AND nlart2.Niederlassungsart = 'Facharzt' AND nlart.niederlassungsart = 'Hausarzt') ON übwlanr = arzt.lanr " & vbCrLf & _
'             "LEFT JOIN (SELECT `aktfvs`.pat_id, leistung vpau FROM `aktfvs` JOIN `leistungen` l ON `aktfvs`.fid = l.fid AND (leistung RLIKE '^030[12]' OR leistung = '01210')) vpau ON f.pat_id = vpau.pat_id WHERE nlart.niederlassungsart = 'Facharzt' OR NOT ISNULL(arzt2.lanr) GROUP BY f.fid" & vbCrLf & _
'             ") vp ON f.fid = vp.fid " & vbCrLf & _
'             "WHERE f.schgr = 00 OR vp.NLArt='F' OR vp.NLArt2='F' " & vbCrLf & _
'             "GROUP BY f.fid ) i) i WHERE i.Leistung <>riLei OR ISNULL(i.Leistung)) i WHERE `erwartete richt.Ziff`<>`vorh.Ziff` ORDER BY pat_id;"
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
'
'#If False THEN
' AwN(AWlf) = "alte Liste Originalschein mit fehlender oder falscher Grundpauschale 0300x"
' sql(AWlf) = _
' "SELECT fpat_id pat_id, ffid fid, name, ICD, Schgr, MinSchgr, LAlter `Alter`, Grundl, Chron, `Zahl(Grundleist)`, `Zahl(Chron'pausch.)` FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS Fpat_id, f.fid AS FFid, gesnamep(f.pat_id) Name, timestampdiff(YEAR, n.gebdat,gl.zeitpunkt) AS lalter, f.Schgr, MIN(MSchGr.Schgr) MinSchGr, gl.leistung AS Grundl, l.leistung AS chron, COUNT(DISTINCT pat_gl.pat_id) `Zahl(Grundleist)`, COUNT(DISTINCT pat_l.pat_id) `Zahl(Chron'pausch.)`, ICD FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN `leistungen` gl ON f.fid = gl.fid AND gl.leistung LIKE '030%' " & vbCrLf & _
' "LEFT JOIN `leistungen` l ON f.fid = l.fid AND l.leistung IN ('03220','03221') " & vbCrLf & _
' "LEFT JOIN `leistungen` pat_gl ON pat_gl.leistung LIKE '030%' AND pat_gl.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND f.pat_id = pat_gl.pat_id " & vbCrLf & _
' "LEFT JOIN `leistungen` pat_l ON pat_l.leistung IN ('03220','03221') AND pat_l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND pat_l.pat_id = f.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "LEFT JOIN `faelle` MSchGr ON f.pat_id = MSchGr.pat_id AND MSchGr.quartal = '" & AktQ & "' " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON d.pat_id = f.pat_id AND (obdauer <> 0 OR d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & ") AND d.icd REGEXP '^E1[0-4]\.' AND d.diagsicherheit <> 'A' AND COALESCE(d.Dggel,0)=0 " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "GROUP BY f.fid ORDER BY gl.leistung DESC) AS innen) AS innen " & vbCrLf & _
' "WHERE " & vbCrLf & _
' "schgr = 0  AND ((ISNULL(grundl) OR grundl NOT LIKE '0300%' OR " & vbCrLf & _
' "(lalter < 4 AND grundl <> '03001') OR ((lalter >= 4 AND lalter < 18) AND grundl <> '03002') OR ((lalter >= 18 AND lalter < 54) AND grundl <> '03003') OR ((lalter >= 54 AND lalter < 75) AND grundl <> '03004') OR (lalter >= 75 AND grundl <> '03005'))) " & vbCrLf & _
' "ORDER BY pat_id "
' ' AND NOT f.goäkatnr IN (40,41)
''  "WHERE (ct > 1 AND NOT NOT ISNULL(`Zahl(Grundleist)`) AND ISNULL(chron)) " & vbCrLf & _
'' "WHERE (ISNULL(leistung)) AND ct > 1 AND NOT ISNULL(icd) "
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
'#END IF
'#END IF
'
'#If ebmalt THEN
' AwN(AWlf) = "Überweisungsschein mit fehlender oder falscher Grundpauschale 03120,03121,03122"
' sql(AWlf) = _
' "SELECT fpat_id pat_id, ffid fid, name, ICD, Schgr, MinSchgr, LAlter `Alter`, Grundl, Chron, `Zahl(Grundleist)`, `Zahl(Chron'pausch.)` FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS Fpat_id, f.fid AS FFid, gesname(f.pat_id) AS Name, timestampdiff(YEAR, n.gebdat," & qtAnf(FristS) & ") AS lalter, f.Schgr, MIN(MSchGr.Schgr) MinSchGr, gl.leistung AS Grundl, l.leistung AS chron, COUNT(DISTINCT pat_gl.pat_id) `Zahl(Grundleist)`, COUNT(DISTINCT pat_l.pat_id) `Zahl(Chron'pausch.)`, ICD FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN `leistungen` gl ON f.fid = gl.fid AND gl.leistung IN ('03110','03111','03112','03120','03121','03122') " & vbCrLf & _
' "LEFT JOIN `leistungen` l ON f.fid = l.fid AND l.leistung IN ('03212') " & vbCrLf & _
' "LEFT JOIN `leistungen` pat_gl ON pat_gl.leistung IN ('03110','03111','03112','03120','03121','03122') AND pat_gl.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND f.pat_id = pat_gl.pat_id " & vbCrLf & _
' "LEFT JOIN `leistungen` pat_l ON pat_l.leistung IN ('03212') AND pat_l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND pat_l.pat_id = f.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "LEFT JOIN `faelle` MSchGr ON f.pat_id = MSchGr.pat_id AND MSchGr.quartal = '" & AktQ & "' " & vbCrLf & _
' "LEFT JOIN `eintraege` e ON e.pat_id = f.pat_id AND art IN (" & artspezG & ") AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
' "LEFT JOIN `rezepteintraege` rz ON rz.pat_id = f.pat_id AND art IN (" & artspezG & ") AND rz.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON d.pat_id = f.pat_id AND (d.obdauer <> 0 OR d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & ") AND d.icd REGEXP '^E1[0-4]\.' AND d.diagsicherheit <> 'A' AND COALESCE(d.Dggel,0)=0 " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "GROUP BY f.fid ORDER BY gl.leistung DESC) AS innen) AS innen " & vbCrLf & _
' "WHERE " & vbCrLf & _
' "schgr <> 0 AND (((minschgr = 0 AND (NOT ISNULL(grundl) OR NOT ISNULL(chron)))) OR " & vbCrLf & _
' "((minschgr = schgr    ) AND (grundl NOT IN ('03120','03121','03122') OR " & vbCrLf & _
' "(lalter < 6 AND grundl <> '03120') OR ((lalter >= 6 AND lalter < 59) AND grundl <> '03121') OR (lalter >= 59 AND grundl <> '03122')))) " & vbCrLf & _
' "ORDER BY pat_id "
''  "WHERE (ct > 1 AND NOT NOT ISNULL(`Zahl(Grundleist)`) AND ISNULL(chron)) " & vbCrLf & _
'' "WHERE (ISNULL(leistung)) AND ct > 1 AND NOT ISNULL(icd) "
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
'#Else 'lalter
'' SELECT i.* FROM (
''SELECT üp.leistung, GROUP_CONCAT(gp.leistung) gesleist, COUNT(gp.leistung) gpzahl, f.pat_id, timestampdiff(YEAR,n.gebdat,üp.zeitpunkt) lalter FROM aktfvs f LEFT JOIN leistungen üp ON f.fid = üp.fid AND üp.leistung LIKE '0301%' LEFT JOIN leistungen gp ON f.fid = gp.fid AND (gp.leistung RLIKE '^030' OR gp.leistung RLIKE '^132[12][012]') LEFT JOIN namen n ON f.pat_id = n.pat_id WHERE schgr = 24 GROUP BY f.fid
'') i;
'
'#If False THEN
' AwN(AWlf) = "Überweisungsschein mit fehlender oder falscher Grundpauschale 0301x (neu)"
' sql(AWlf) = "SELECT Pat_ID, riLei `erwartete richt.Ziff`, i.gesleist `vorh.Ziff`, i.gpzahl  `Zahl d.vorh.Ziff`, i.lalter `Alter n.Leist.`, i.falter `Alter n.Fall`, i.fanf `Fallanf` FROM " & vbCrLf & _
'             "(SELECT i.Pat_ID, i.leistung," & vbCrLf & _
'             "IF(lalter<4,'03011',IF(falter<18,'03012',IF(falter<54,'03013',IF(falter<75,'03014','03015')))) riLei," & vbCrLf & _
'             "i.gesleist, i.gpzahl, i.lalter, i.falter, i.fanf FROM " & vbCrLf & _
'             "(SELECT üp.leistung, GROUP_CONCAT(gp.leistung) gesleist, SUM(gp.lzahl) gpzahl, f.pat_id, timestampdiff(YEAR,n.gebdat,üp.zeitpunkt) lalter, timestampdiff(YEAR,n.gebdat,IF(gf.fanf>gf.bhfb,gf.fanf,gf.bhfb)) falter, DATE(IF(gf.fanf>gf.bhfb,gf.fanf,gf.bhfb)) fanf FROM aktfvs f LEFT JOIN faelle gf ON f.fid = gf.fid LEFT JOIN leistungen üp ON f.fid = üp.fid AND üp.leistung LIKE '0301%' LEFT JOIN leistungen gp ON f.fid = gp.fid AND (gp.leistung RLIKE '^030' OR gp.leistung RLIKE '^132[12][012]') LEFT JOIN namen n ON f.pat_id = n.pat_id WHERE f.schgr = 24 GROUP BY f.fid " & vbCrLf & _
'             ") i) i WHERE i.Leistung <>riLei OR ISNULL(i.Leistung) ORDER BY pat_id"
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
'#END IF
'
' AwN(AWlf) = "Überweisungsschein vom Hausarzt mit fehlender oder falscher Grundpauschale 0301x"
' sql(AWlf) = "SELECT * FROM (" & vbCrLf & _
'             "SELECT Pat_ID, Name, riLei `erwartete richt.Ziff`, i.gesleist `vorh.Ziff`, i.gpzahl  `Zahl d.vorh.Ziff`, Schgr, i.falter `Alter n.Fall`, i.fanf `Fallanf` " & vbCrLf & _
'             "FROM (SELECT Name, übwlanr, üw, NLArt, LANR2,Üw2,nlart2, i.Pat_ID, Schgr, i.leistung,IF(lalter<4,'03011',IF(falter<18,'03012',IF(falter<54,'03013',IF(falter<75,'03014','03015')))) riLei,i.gesleist, i.gpzahl, i.lalter, i.falter, i.fanf " & vbCrLf & _
'             "FROM (SELECT CONCAT(n.nachname,', ',n.vorname) Name, vp.übwlanr, vp.üw, vp.NLArt,vp.LANR2,vp.Üw2,vp.nlart2, üp.leistung, GROUP_CONCAT(gp.leistung) gesleist, SUM(gp.lzahl) gpzahl, f.pat_id, gf.schgr, timestampdiff(YEAR,n.gebdat,gp.zeitpunkt) lalter, timestampdiff(YEAR,n.gebdat,IF(gf.fanf>gf.bhfb,gf.fanf,gf.bhfb)) falter, DATE(IF(gf.fanf>gf.bhfb,gf.fanf,gf.bhfb)) fanf " & vbCrLf & _
'             "FROM aktfvs f " & vbCrLf & _
'             "LEFT JOIN faelle gf ON f.fid = gf.fid " & vbCrLf & _
'             "LEFT JOIN leistungen üp ON f.fid = üp.fid AND üp.leistung LIKE '0301%' " & vbCrLf & _
'             "LEFT JOIN leistungen gp ON f.fid = gp.fid AND (gp.leistung RLIKE '^030[01]' OR gp.leistung RLIKE '^132[12][012]') " & vbCrLf & _
'             "LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
'             "LEFT JOIN ( " & vbCrLf & _
'             "SELECT f.pat_id, f.fid, CONCAT_WS(',', f.nachname, f.vorname) Pat, vpau, f.übwlanr, CONCAT_WS(',',arzt.nachname,arzt.vorname) ÜW, IF(nlart.niederlassungsart='Hausarzt','h',LEFT(nlart.niederlassungsart,1)) NLArt, arzt2.lanr LANR2, CONCAT_WS(',',arzt2.nachname,arzt2.vorname) ÜW2, IF(nlart2.niederlassungsart='Hausarzt','h',LEFT(nlart2.niederlassungsart,1)) nlart2 FROM (`aktfvs` JOIN `faelle` f ON `aktfvs`.fid = f.fid AND f.schgr BETWEEN 20 AND 29) LEFT JOIN (haerzte.arzt JOIN haerzte.nlart ON arzt.nlart_id = nlart.idnlart LEFT JOIN (haerzte.arzt_has_bs ahb LEFT JOIN haerzte.arzt_has_bs ahb2 ON ahb.bs_id = ahb2.bs_id AND ahb.arzt_id <> ahb2.arzt_id JOIN  (haerzte.arzt arzt2 JOIN haerzte.nlart nlart2 ON arzt2.nlart_id = nlart2.idnlart) ON ahb2.arzt_id = arzt2.idarzt   ) ON arzt.idarzt = ahb.arzt_id AND nlart2.Niederlassungsart = 'Facharzt' AND nlart.niederlassungsart = 'Hausarzt') ON übwlanr = arzt.lanr " & vbCrLf & _
'             "LEFT JOIN (SELECT `aktfvs`.pat_id, leistung vpau FROM `aktfvs` JOIN `leistungen` l ON `aktfvs`.fid = l.fid AND (leistung RLIKE '^030[12]' OR leistung = '01210')) vpau ON f.pat_id = vpau.pat_id WHERE nlart.niederlassungsart = 'Facharzt' OR NOT ISNULL(arzt2.lanr) GROUP BY f.fid" & vbCrLf & _
'             ") vp ON f.fid = vp.fid " & vbCrLf & _
'             "WHERE f.schgr BETWEEN 20 AND 29 AND (ISNULL(vp.NLArt) OR vp.NLArt<>'F') AND (ISNULL(vp.NLArt2) OR vp.NLArt2<>'F') " & vbCrLf & _
'             "GROUP BY f.fid ) i) i WHERE i.leistung <>riLei OR ISNULL(i.Leistung)) i WHERE `erwartete richt.Ziff`<>`vorh.Ziff` ORDER BY pat_id;"
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
'
'
'#If False THEN
' AwN(AWlf) = "Überweisungsschein mit fehlender oder falscher Grundpauschale 0301x"
' sql(AWlf) = _
' "SELECT fpat_id pat_id, ffid fid, name, ICD, Schgr, MinSchgr, LAlter `Alter`, Grundl, Chron, `Zahl(Grundleist)`, `Zahl(Chron'pausch.)` FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS Fpat_id, f.fid AS FFid, LEFT(CONCAT(n.Nachname, ' ', n.Vorname),25) AS Name, timestampdiff(YEAR, n.gebdat,gl.zeitpunkt) AS lalter, f.Schgr, MIN(MSchGr.Schgr) MinSchGr, gl.leistung AS Grundl, l.leistung AS chron, COUNT(DISTINCT pat_gl.pat_id) `Zahl(Grundleist)`, COUNT(DISTINCT pat_l.pat_id) `Zahl(Chron'pausch.)`, ICD FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN `leistungen` gl ON f.fid = gl.fid AND gl.leistung LIKE '030%' " & vbCrLf & _
' "LEFT JOIN `leistungen` l ON f.fid = l.fid AND l.leistung IN ('03220','03221') " & vbCrLf & _
' "LEFT JOIN `leistungen` pat_gl ON pat_gl.leistung LIKE '030%' AND pat_gl.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND f.pat_id = pat_gl.pat_id " & vbCrLf & _
' "LEFT JOIN `leistungen` pat_l ON pat_l.leistung IN ('03220','03221') AND pat_l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " AND pat_l.pat_id = f.pat_id " & vbCrLf & _
' "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
' "LEFT JOIN `faelle` MSchGr ON f.pat_id = MSchGr.pat_id AND MSchGr.quartal = '" & AktQ & "' " & vbCrLf & _
' "LEFT JOIN `eintraege` e ON e.pat_id = f.pat_id AND art IN (" & artspezG & ") AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
' "LEFT JOIN `rezepteintraege` rz ON rz.pat_id = f.pat_id AND art IN (" & artspezG & ") AND rz.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
' "LEFT JOIN `diagnosen` d ON d.pat_id = f.pat_id AND (d.obdauer <> 0 OR d.diagdatum BETWEEN " & lQAnfuEnd(FristS) & ") AND d.icd REGEXP '^E1[0-4]\.' AND d.diagsicherheit <> 'A' AND COALESCE(d.Dggel,0)=0 " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id " & vbCrLf & _
' "GROUP BY f.fid ORDER BY gl.leistung DESC) AS innen) AS innen " & vbCrLf & _
' "WHERE " & vbCrLf & _
' "schgr <> 0 AND (((minschgr = 0 AND (NOT ISNULL(grundl) OR NOT ISNULL(chron)))) OR " & vbCrLf & _
' "((minschgr = schgr) AND (grundl NOT LIKE '0300%' OR " & vbCrLf & _
' "(lalter < 4 AND grundl <> '03011') OR ((lalter >= 4 AND lalter < 18) AND grundl <> '03012') OR ((lalter >= 18 AND lalter < 54) AND grundl <> '03013') OR ((lalter >= 54 AND lalter < 75) AND grundl <> '03014') OR (lalter >= 75 AND grundl <> '03015')))) " & vbCrLf & _
' "ORDER BY pat_id "
''  "WHERE (ct > 1 AND NOT NOT ISNULL(`Zahl(Grundleist)`) AND ISNULL(chron)) " & vbCrLf & _
'' "WHERE (ISNULL(leistung)) AND ct > 1 AND NOT ISNULL(icd) "
' mins(AWlf) = 7
' maxs(AWlf) = 20
' AWlf = AWlf + 1
'#END IF
'#END IF

#Else
' AwN(AWlf) = "Falsche Zahl Grundpauschalen 0301x"
#End If

' #Const mitcovid = 1
#If mitcovid Then
' 148
 AwN(AWlf) = "Falsche Zahl Genesenenzertifikatabrechnungen 88371 (lauto) (vorher 36)"
' sql(AWlf) = _
"  SELECT f.pat_id, gesname(f.pat_id), " & vbCrLf & _
"  (SELECT SUM(lzahl) FROM leistungen WHERE pat_id=f.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND leistung LIKE '030%') zahl, " & vbCrLf & _
"  (SELECT COALESCE(GROUP_CONCAT(leistung),'') FROM leistungen WHERE pat_id=f.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND leistung LIKE '030%') leistung " & vbCrLf & _
" FROM aktfvs f HAVING zahl <> 1;"
sql(AWlf) = _
" SELECT f.pat_id PID, gesname(f.pat_id) PName, f.LANRID, DATE(k.zeitpunkt) LEIDAT " & vbCrLf & _
" ,CASE WHEN COALESCE(SUM(lzahl),0)=0 THEN '88371 dazu' ELSE COALESCE(CONCAT(SUM(lzahl),0)-1,' mal 88371 weg') END LEIFEHLER " & vbCrLf & _
"  FROM aktfvmi f " & vbCrLf & _
" LEFT JOIN forminhkopf k ON f.pat_id= k.pat_id " & vbCrLf & _
" LEFT JOIN formulare fo ON k.form_id=fo.formid " & vbCrLf & _
" LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND DATE(l.zeitpunkt)=DATE(k.zeitpunkt) AND leistung = '88371' " & vbCrLf & _
" WHERE form_abk='covge' AND k.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
" GROUP BY f.pat_id" & vbCrLf & _
" HAVING COALESCE(SUM(lzahl),0)<>1"
'Dim rs As New ADODB.Recordset
'myFrag rs, sql(AWlf)
mins(AWlf) = 7
maxs(AWlf) = 20
AWlf = AWlf + 1
' 149
#Else
' 148
#End If
' ktag fehlerhaft
 AwN(AWlf) = "Fehlende Leistungen für Impfungen (vac) außer Hepatitis (vorher 57)"
' sql(AWlf) = "SELECT f.pat_id, gesnameg(f.pat_id) Name, e.Tag, COUNT(art) `Impfungs-Zahl`, SUM(lzahl) `Leistungs-Zahl`, inhalt, MAX(leistung) " & vbCrLf & _
             "FROM aktfvs f " & vbCrLf & _
             "LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
             "LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
             "     AND art IN ('vac') AND NOT impfart(inhalt) IN (9) " & vbCrLf & _
             "LEFT JOIN leistungen l ON f.pat_id=l.pat_id AND DATE(l.zeitpunkt)=e.tag AND leistung IN ('89109','89111','89112','89112Z','89120R','89120','89102A','89102B','89102R','89124A','89122A','89122B','89122R','89124B','89124R','89303','89303R','89302','89302R','89300','89300A','89300B','89400E','89201R','89112Z','89119R') AND DATE(l.zeitpunkt) = DATE(e.zeitpunkt) " & vbCrLf & _
             "WHERE NOT ISNULL(art) " & vbCrLf & _
             "GROUP BY f.pat_id, DATE(e.zeitpunkt) " & vbCrLf & _
             "HAVING `Impfungs-Zahl`<>`Leistungs-Zahl`"
sql(AWlf) = "" & _
"WITH eintr AS ( " & vbCrLf & _
" SELECT DATE(zeitpunkt) Tag, impfart(inhalt) iart, Art, Inhalt, Pat_id FROM eintraege " & vbCrLf & _
") " & vbCrLf & _
"SELECT f.Pat_id, gesnameg(f.pat_id) NAME, e.Tag, e.iart, l.leistung, e.Inhalt " & vbCrLf & _
"FROM aktfv f " & vbCrLf & _
"LEFT JOIN eintr e " & vbCrLf & _
" ON e.pat_id=f.pat_id AND e.Tag BETWEEN qanf() AND now() AND art = 'vac' " & vbCrLf & _
"LEFT JOIN leistungen l " & vbCrLf & _
" ON l.pat_id=f.pat_id AND DATE(l.zeitpunkt)=e.TaG " & vbCrLf
sql(AWlf) = sql(AWlf) & _
" AND ( " & vbCrLf & _
"    (iart IN (1,101) AND leistung IN ('89111','89112','89112Y','89112Z','89112B'))" & vbCrLf & _
" OR (iart IN (2,102) AND leistung IN ('89118A','89118B','89119','89119R','89120','89120R','89120V','89120X'))" & vbCrLf & _
" OR (iart IN (3,103) AND leistung IN ('89122R','89302','89302R','89303','89303R','89303Y','89400','89400E','89400K','89500A','89500B','89600A','89600B'))" & vbCrLf & _
" OR (iart IN (4,104) AND leistung IN ('89128A','89128B','89129A','89129B'))" & vbCrLf & _
" OR (iart IN (5,105) AND leistung IN ('89102A','89102B','89102R','89102V','89102W','89102X'))" & vbCrLf & _
" OR (iart IN (6,106) AND leistung IN ('89401A','89401B','89401V','89401W'))" & vbCrLf & _
" OR (iart IN (7,107) AND leistung IN ('89124A','89124B','89124R','89201A','89201B','89201R','89400E','89400K'))" & vbCrLf & _
" OR (iart IN (9,109) AND leistung IN ('89105A','89105B','89105R','89105V','89105W','89105X','89106A','89106B','89107A','89107B','89107R','89107V','89107W','89107X','89108A','89108B','89108R'))" & vbCrLf & _
" OR (iart IN (10,110) AND leistung IN ('89114','89115A','89115B','89115R','89115C','89115D','89115S','89115V','89115W','89115X'))" & vbCrLf & _
" OR (iart IN (12,112) AND leistung IN ('89110A'))" & vbCrLf & _
" OR (iart IN (13,113) AND leistung IN ('89133Y','89133V'))" & vbCrLf & _
" OR (iart IN (14,114) AND leistung IN ('89201A','89201B','89201R'))" & vbCrLf & _
" OR (iart IN (16,116) AND leistung RLIKE '^89132[VWX]')" & vbCrLf & _
" OR (iart IN (17,117) AND leistung IN ('89138','89137'))" & vbCrLf
sql(AWlf) = sql(AWlf) & _
" OR (iart IN (127) AND leistung RLIKE '^88337[ABRVWXGHK]') " & vbCrLf & _
" OR (iart IN (121) AND leistung RLIKE '^88331[ABRVWXGHK]|88345[RX]') " & vbCrLf & _
" OR (iart IN (128) AND leistung RLIKE '^88338[ABRVWXGHK]') " & vbCrLf & _
" OR (iart IN (122) AND leistung RLIKE '^88332[ABRVWXGHK]') " & vbCrLf & _
" OR (iart IN (124) AND leistung RLIKE '^88334[ABRVWXGHK]') " & vbCrLf & _
" OR (iart IN (125) AND leistung RLIKE '^88335[ABRVWXGHK]') " & vbCrLf & _
" OR (iart IN (126) AND leistung RLIKE '^88336[ABVWGH]') " & vbCrLf & _
") " & vbCrLf & _
"WHERE NOT ISNULL(inhalt) AND (schgr<>90 OR impfart(e.inhalt)=8) AND ISNULL(leistung) " & vbCrLf & _
"AND NOT inhalt RLIKE 'Twinrix|havrix|Engerix|strova|Bexsero|Nimenrix'" & vbCrLf & _
"AND NOT (inhalt RLIKE 'tetagam' AND NOT inhalt RLIKE 'tetanol|boostrix|grundimm')" & vbCrLf & _
"GROUP BY f.pat_id, iart, leistung" & vbCrLf & _
";"
' 14.7.24 tetagam ausgeschlossen
'    "   WHEN inh RLIKE 'Astra|Vaxzev' THEN SET erg='123';" & vbCrLf & _
    "   WHEN inh RLIKE 'Spikevax|Moderna' THEN SET erg='122';" & vbCrLf & _
    "   WHEN inh RLIKE 'biontech|comirnaty|corminaty|cominarty|comiarty|coirnaty|cominary|comirnary|commirnaty|comirnarty|cominaty|comitnaty' THEN IF inh RLIKE ': G' THEN SET erg='127'; ELSE SET erg='121'; END IF;" & vbCrLf & _
    "   WHEN inh RLIKE 'Novavax' THEN SET erg='125';" & vbCrLf & _
    "   WHEN inh RLIKE 'Valneva' THEN SET erg='126';" & vbCrLf & _


' 8833[12][ABRVWXGHK]|88333[ABVWGH]|88334[YI]{0,1}
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
' inhalt RLIKE 'havr|twin|shingrix|zostavax|enger|hbvax|vaqta|ambirix|prevenar|priorix|m-m-rva|mmrvax' "
'             "OR (iart IN (8,108) AND leistung NOT IN ('88331A','88331B','88331V','88331W','88331G','88331H','88332A','88332B','88332V','88332W','88332G','88332H','88333A','88333B','88333V','88333W','88333G','88333H')) " & vbCrLf & _

 AwN(AWlf) = "Unwohlgeformte Covid-Impfeinträge"
 sql(AWlf) = _
 "SELECT pat_id, gesname(pat_id) PName, Zeitpunkt, Inhalt" & vbCrLf & _
 "FROM eintraege" & vbCrLf & _
 "WHERE art='vac' and inhalt RLIKE 'covid' and zeitpunkt >= qbeg(qanf()- INTERVAL 1 DAY)" & vbCrLf & _
 "AND (inhalt RLIKE ' .Impfung' OR inhalt NOT RLIKE 'ch.-b.: [^ ]* ?, [1-9]')" & vbCrLf & _
 ""
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
#If mitcovid Then
' 151
' ktag fehlerhaft
 AwN(AWlf) = "Covid-Impfberatungen (Makro cib) ohne Impfung mit Zahl d.Leistungen 88322 <>1 (lauto) (vorher 133)"
 sql(AWlf) = _
 "SELECT i.*, IF(88322_Zahl=0,'88322 dazu',CONCAT(88322_Zahl-1,' 83322 weg')) LEIFEHLER FROM (" & vbCrLf & _
 "SELECT f.pat_id PID, gesname(f.pat_id) Name, e.Zeitpunkt LEIDAT, e.art, e.Inhalt EName " & vbCrLf & _
 ",(SELECT COUNT(0) FROM leistungen WHERE pat_id=f.pat_id AND leistung='88322' AND DATE(zeitpunkt)=DATE(e.zeitpunkt)) 88322_Zahl" & vbCrLf & _
 ",(SELECT COUNT(0) FROM leistungen WHERE pat_id=f.pat_id AND leistung RLIKE '^8833[12][ABRVWXGHK]|^88333[ABVWGH]|^88334[YI]{0,1}' AND zeitpunkt BETWEEN qend() - INTERVAL 1 YEAR AND qend()) ImpfZahl" & vbCrLf & _
 ", f.LANRID " & vbCrLf & _
 "FROM aktfvmi f" & vbCrLf & _
 "LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.ZeitPunkt BETWEEN qanf() AND qend() " & vbCrLf & _
 "WHERE art='cib' AND inhalt LIKE 'Covid-Impfberatung%'" & vbCrLf & _
 "HAVING 88322_Zahl<>1 AND ImpfZahl=0) i" & vbCrLf & _
 ";"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1

' 152
' ktag fehlerhaft
 AwN(AWlf) = "Covid-Testzeugnisse mit Zahl d.Leistungen 88320 <>1 (lauto) (vorher 134)"
 sql(AWlf) = _
 "SELECT i.*, IF(88320_Zahl=0,'88320 dazu',CONCAT(88320_Zahl-1,' 83320 weg')) LEIFEHLER FROM (" & vbCrLf & _
 "SELECT f.pat_id PID, gesname(f.pat_id) Name, b.Zeitpunkt LEIDAT, b.name DName " & vbCrLf & _
 ",(SELECT COUNT(0) FROM leistungen WHERE pat_id=f.pat_id AND leistung='88320' AND DATE(zeitpunkt)=DATE(b.zeitpunkt)) 88320_Zahl" & vbCrLf & _
 ", '88320 dazu' LEIFEHLER, f.LANRID" & vbCrLf & _
 "FROM aktfv f " & vbCrLf & _
 "LEFT JOIN tmbrie b ON f.pat_id=b.pat_id AND b.ZeitPunkt BETWEEN qanf() AND qend() " & vbCrLf & _
 "WHERE b.name RLIKE 'covid|corona|impf' AND b.name RLIKE 'Attest|Zeugnis' AND NOT b.name LIKE '%.pdf' " & vbCrLf & _
 "AND NOT EXISTS(SELECT 0 FROM eintraege WHERE pat_id=f.pat_id AND art='vac' AND zeitpunkt BETWEEN qend() - INTERVAL 1 YEAR AND qend() AND inhalt RLIKE 'Janssen|Johnson|Astra|Vaxzev|Spikevax|Moderna|biontech|comirnaty|corminaty|cominarty|comiarty|coirnaty|cominary|comirnary|commirnaty|comirnarty|cominaty|comitnaty|novavax')" & vbCrLf & _
 "GROUP BY f.pat_id, DATE(leidat) " & vbCrLf & _
 "HAVING 88320_Zahl<>1) i" & vbCrLf & _
 ";"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
' 153
AwN(AWlf) = "Zu ungenaue Covid-Impf-Abrechnung (lauto) (vorher 137)"
' AND COALESCE(Dggel,0)=0
sql(AWlf) = "" & _
"SELECT i.pat_id PID, gesname(i.pat_id) PName, DATE(i.zeitpunkt) LEIDAT " & vbCrLf & _
", COALESCE((SELECT MAX(lanrid) FROM faelle WHERE pat_id=i.pat_id AND quartal=quartal(i.zeitpunkt)),0) LANRID " & vbCrLf & _
" , CASE WHEN COALESCE(SUM(lzahl),0)=0 THEN CONCAT(solllei,' dazu') WHEN COALESCE(SUM(lzahl),0)>1 THEN CONCAT(SUM(lzahl) - 1,' ',leistung,' weniger') ELSE CONCAT(solllei,' statt ',COALESCE(GROUP_CONCAT(DISTINCT(leistung)),'')) END LEIFEHLER " & vbCrLf & _
", COALESCE(SUM(lzahl),0) lzahl " & vbCrLf & _
", solllei " & vbCrLf & _
", COALESCE(GROUP_CONCAT(DISTINCT(leistung)),'') istlei " & vbCrLf & _
", dnr,vorvac,znc,vacinh23,c19i,vorze Vorzert,ob23" & vbCrLf & _
", inhalt" & vbCrLf & _
"FROM ( " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"SELECT e.pat_id, e.ZeitPunkt, e.inhalt " & vbCrLf & _
", dnr,vorvac,znc,vacinh23,c19i,vorze,ob23" & vbCrLf & _
", CASE WHEN inhalt RLIKE 'Janssen|Johnson' THEN '88334' " & vbCrLf & _
"       WHEN inhalt RLIKE 'Astra|Vaxzev' THEN " & vbCrLf & _
"           CASE WHEN ob23>2 THEN '88333R' WHEN ob23=2 THEN '88333B' ELSE '88333A' END " & vbCrLf & _
"       WHEN inhalt RLIKE 'Spikevax|Moderna' THEN " & vbCrLf & _
"           CASE WHEN ob23>2 THEN '88332R' WHEN ob23=2 THEN '88332B' ELSE '88332A' END " & vbCrLf & _
"       WHEN inhalt RLIKE 'biontech|comirnaty|corminaty|cominarty|comiarty|coirnaty|cominary|comirnary|commirnaty|comirnarty|cominaty|comitnaty' THEN " & vbCrLf & _
"           IF(inhalt RLIKE ':[ ]*G'," & vbCrLf & _
"            CASE WHEN ob23>2 THEN '88337R' WHEN ob23=2 THEN '88337B' ELSE '88337A' END," & vbCrLf & _
"            CASE WHEN ob23>2 THEN '88331R' WHEN ob23=2 THEN '88331B' ELSE '88331A' END)" & vbCrLf & _
"       ELSE '' END solllei " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"FROM " & vbCrLf & _
" (SELECT e.*, CASE WHEN dnr>0 THEN dnr WHEN vacinh23>1 THEN vacinh23 WHEN vorvac OR znc OR c19i OR vorze THEN 2 ELSE 1 END ob23 " & vbCrLf & _
"  FROM" & vbCrLf & _
"  (SELECT e.*" & vbCrLf & _
"   , EXISTS(SELECT 0 FROM eintraege WHERE pat_id=e.pat_id AND art = 'vac' AND inhalt RLIKE 'Janssen|Johnson|Astra|Vaxzev|Spikevax|Moderna|biontech|comirnaty|corminaty|cominarty|comiarty|coirnaty|cominary|comirnary|commirnaty|comirnarty|cominaty|comitnaty|novavax' AND DATE(zeitpunkt)<DATE(e.zeitpunkt)) vorvac" & vbCrLf & _
"   , CASE WHEN inhalt RLIKE '2\\.' THEN 2 WHEN inhalt RLIKE '[3-9]\\.|[1-9][0-9]\\.' THEN 3 ELSE 1 END vacinh23 " & vbCrLf & _
"   , EXISTS(SELECT 0 FROM eintraege WHERE pat_id=e.pat_id AND art='c19i' AND zuDatum(MID(inhalt,INSTR(inhalt,' am ')),qanf()) BETWEEN 20201201 AND DATE(e.zeitpunkt) - INTERVAL 1 DAY) c19i" & vbCrLf & _
"   , EXISTS(SELECT feldinh FROM formular WHERE pat_id=e.pat_id AND form_abk IN ('covze','covge') AND feld IN ('impfdatum','testpositivdatum') AND STR_TO_DATE(feldinh,'%d.%m.%Y')<DATE(e.zeitpunkt)) vorze" & vbCrLf & _
"   , COALESCE((SELECT DATE_FORMAT(diagdatum,'%d.%m.%y') FROM diagnosen WHERE pat_id=e.pat_id AND diagdatum<DATE(e.zeitpunkt) AND icd='U07.1' AND diagsicherheit IN ('G','Z',' ') ORDER BY zeitpunkt LIMIT 1),'') znc " & vbCrLf & _
"   , COALESCE((SELECT MAX(LEFT(fd.feldinh,1)) FROM formular fd LEFT JOIN formular f ON fd.foid=f.foid AND fd.feld='dosis' AND f.feld='impfdatum' WHERE f.pat_id=e.pat_id AND f.form_abk='covze' AND STR_TO_DATE(f.feldinh,'%d.%m.%Y')=DATE(e.zeitpunkt)),0) dnr " & vbCrLf & _
"   FROM eintraege e" & vbCrLf & _
"   WHERE zeitpunkt>=qanf()" & vbCrLf & _
"  ) e" & vbCrLf & _
" ) e" & vbCrLf & _
"" & vbCrLf & _
"WHERE art = 'vac' AND zeitpunkt > 20210331 AND inhalt RLIKE 'Janssen|Johnson|Astra|Vaxzev|Spikevax|Moderna|biontech|comirnaty|corminaty|cominarty|comiarty|coirnaty|cominary|comirnary|commirnaty|comirnarty|cominaty|comitnaty|novavax' " & vbCrLf & _
"HAVING solllei <> '') i " & vbCrLf & _
"LEFT JOIN leistungen l ON l.pat_id=i.pat_id AND DATE(l.zeitpunkt)=date(i.zeitpunkt) AND leistung RLIKE '^8833[12578][ABRVWXGHK]|^88333[ABVWGH]|^88334[YI]{0,1}|^88336[ABVWGH]' " & vbCrLf & _
"GROUP BY i.pat_id,i.zeitpunkt " & vbCrLf & _
"HAVING solllei<>REPLACE(REPLACE(REPLACE(istlei,'X','R'),'W','B'),'V','A') OR lzahl<>1; "
' vacinh23 = in der Inhaltsspalte von "vac" steht "2." oder "3.", sonst 1
' vorvac = es existiert ein älterer "vac"-Eintrag mit einer Corona-Firma im Namen,
' c19i = ein c19i-Makro verweist auf einen früheren Zeitpunkt,
' vorze = ein Zertifikat verweist auf früheren Zeitpunkt,
' ob23 = vacinh23>1, dann das, sonst wenn mindestens eine dieser Bedingungen ist erfüllt => 2, sonst 1
' "   , (inhalt RLIKE '[2-9]\\.' AND NOT EXISTS(SELECT 0 FROM eintraege WHERE pat_id=e.pat_id AND art = 'vac' AND inhalt RLIKE 'Janssen|Johnson|Astra|Vaxzev|Spikevax|Moderna|biontech|comirnaty|corminaty|cominarty|comiarty|coirnaty|cominary|comirnary|commirnaty|comirnarty|cominaty|comitnaty|novavax' AND DATE(zeitpunkt)>DATE(e.zeitpunkt))) vacinh23" & vbCrLf &
' 31.12.21:
' 1) 3. und folgende Impfungen sollten jetzt erkannt werden,
' 2) ein zusätzliches Feld "dnr" bestimmt mit höchster Priorität die Reihenfolge der Impfung aus der entsprechenden Kennzeichnung im Impfzertifikat,
' 3) berufliche Impfungen werden, wenn eingegeben, berücksichtigt, aber nicht aktiv vorgeschlagen.

 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
 ' 154
 AwN(AWlf) = "Covidimpfzertifikatabrechnungsauffälligkeiten (lauto) (vorher 138)"
 sql(AWlf) = "" & _
"SELECT PID,gesname(pid) PName,ZP,COALESCE(impfdd,'') Impfdatum,COALESCE(slei,'') `Soll-Leistung`,COALESCE(srang,'') `Soll-Rang`,COALESCE(vaczl,'') vaczl,COALESCE(vorfzl,'') vorfzl,COALESCE(ilei,'') ilei,COALESCE(lzahl,'') lzahl,COALESCE(irang,'') irang,LANRID,LEIDAT " & vbCrLf & _
",CONCAT(LEIFEHLER,IF(isnull(slei),CONCAT(' (',COUNT(0),' mal)'),'')) LEIFEHLER " & vbCrLf & _
"FROM ( " & vbCrLf & _
"SELECT IF(isnull(spid),ipid,spid) PID, IF(isnull(szp),izp,szp) ZP,impfdd,slei,srang,ilei,lzahl,irang " & vbCrLf & _
",IF(isnull(sli),ili,sli) LANRID,vaczl,vorfzl,szp,DATE(IF(isnull(szp),izp,szp)) LEIDAT " & vbCrLf & _
",IF(isnull(ilei),CONCAT(slei,' dazu'),IF(isnull(slei),CONCAT(ilei,' weg'),'')) LEIFEHLER " & vbCrLf & _
"FROM ( " & vbCrLf & _
" WITH soll AS (  " & vbCrLf & _
"  WITH cfor AS ( " & vbCrLf & _
"     SELECT pat_id,zeitpunkt szp,form_abk,STR_TO_DATE(feldinh,'%d.%m.%Y') impfdd,foid" & vbCrLf & _
"     FROM formular" & vbCrLf & _
"     WHERE form_abk='covze' AND feld='impfdatum'" & vbCrLf & _
"     GROUP BY pat_id,DATE(zeitpunkt),feldinh" & vbCrLf & _
"  ) " & vbCrLf & _
"  , impfeintr AS (" & vbCrLf & _
"    SELECT pat_id,inhalt,art,zeitpunkt FROM eintraege WHERE art='vac' AND inhalt RLIKE 'Janssen|Johnson|Astra|Vaxzev|Spikevax|Moderna|biontech|comirnaty|corminaty|cominarty|comiarty|coirnaty|cominary|comirnary|commirnaty|comirnarty|cominaty|comitnaty|novavax' GROUP BY pat_id,DATE(zeitpunkt)" & vbCrLf & _
"  )" & vbCrLf & _
"  SELECT spid, szp, impfdd,LANRID sli,vaczl,vorfzl " & vbCrLf & _
"  , CASE WHEN EXISTS(SELECT pat_id FROM impfeintr WHERE pat_id=spid AND DATE(zeitpunkt)=DATE(impfdd)) THEN '88351' " & vbCrLf & _
"         ELSE '88352' END slei " & vbCrLf & _
"    , RANK() OVER(PARTITION BY spid,slei,DATE(szp) ORDER BY slei,szp,foid) srang " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"  FROM (SELECT f.LANRID, i.pat_id spid, i.szp,impfdd,foid " & vbCrLf & _
"        , COALESCE((SELECT COUNT(0) OVER() FROM impfeintr " & vbCrLf & _
"                    WHERE pat_id=i.pat_id AND DATE(zeitpunkt)<=DATE(i.szp) " & vbCrLf & _
"                    GROUP BY pat_id,DATE(zeitpunkt) LIMIT 1),0) vaczl " & vbCrLf & _
"        , (SELECT COUNT(0) FROM cfor WHERE pat_id=i.pat_id AND szp BETWEEN qanf() AND qend() AND (szp<i.szp OR (szp=i.szp AND impfdd<i.impfdd))" & vbCrLf & _
"           AND NOT EXISTS(SELECT pat_id FROM impfeintr WHERE pat_id=cfor.pat_id AND DATE(zeitpunkt)=DATE(cfor.impfdd))" & vbCrLf & _
"          ) vorfzl " & vbCrLf & _
"        FROM faelle f LEFT JOIN cfor i USING (pat_id) " & vbCrLf & _
"        WHERE NOT ISNULL(Form_Abk) " & vbCrLf & _
"        AND szp BETWEEN qanf() AND qend() " & vbCrLf & _
"        GROUP BY f.pat_id, impfdd" & vbCrLf & _
"  ) i " & vbCrLf & _
") " & vbCrLf
sql(AWlf) = sql(AWlf) & _
",    ist AS ( " & vbCrLf & _
" WITH RECURSIVE " & vbCrLf & _
"  dlei AS (SELECT pat_id ipid,zeitpunkt izp,leistung ilei,lzahl,id,lanrid ili FROM leistungen WHERE leistung IN ('88350','88351','88352','88353')), " & vbCrLf & _
"  numbers AS (SELECT 1 Dt UNION SELECT Dt + 1 FROM numbers WHERE Dt <= COALESCE((SELECT MAX(lzahl)FROM dlei),0)) " & vbCrLf & _
" SELECT ipid, izp, ilei, lzahl,ili,RANK() OVER(PARTITION BY ipid,ilei,DATE(izp) ORDER BY id,dt) irang " & vbCrLf & _
" FROM dlei " & vbCrLf & _
" LEFT JOIN numbers ON DT<=lzahl " & vbCrLf & _
" WHERE izp BETWEEN qanf() AND qend() " & vbCrLf & _
" ORDER BY ipid,ilei,izp,Dt " & vbCrLf & _
") " & vbCrLf & _
"SELECT * FROM soll LEFT JOIN ist ON soll.spid=ipid AND DATE(soll.szp)=DATE(izp) AND soll.slei=ist.ilei AND srang=irang" & vbCrLf & _
"UNION DISTINCT " & vbCrLf & _
"SELECT * FROM soll RIGHT JOIN ist ON soll.spid=ipid AND DATE(soll.szp)=DATE(izp) AND soll.slei=ist.ilei AND srang=irang" & vbCrLf & _
") i " & vbCrLf
sql(AWlf) = sql(AWlf) & _
"WHERE ISNULL(slei) OR isnull(ilei) " & vbCrLf & _
") i " & vbCrLf & _
"GROUP BY PID,DATE(zp),ilei,slei,srang " & vbCrLf & _
"ORDER BY PID,DATE(zp),IF(isnull(slei),ilei,slei),IF(isnull(srang),irang,srang) " & vbCrLf & _
";"
' WHEN vaczl>IF(vorfzl>1,1,vorfzl)
' "         WHEN vaczl<vorfzl THEN '88353' " & vbCrLf & _

 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 155
AwN(AWlf) = "Möglicherweise unerlaubt abgerechnete Impfberatungen (88322) (vorher 140)"
 sql(AWlf) = "" & _
 "SELECT l.pat_id,gesname(l.pat_id) PName, l.Zeitpunkt,l.Leistung, n.ZeitPunkt, n.Leistung " & vbCrLf & _
 "FROM leistungen l " & vbCrLf & _
 "LEFT JOIN leistungen n ON l.pat_id=.n.pat_id AND n.leistung RLIKE '^8833[12578][ABRVWXGHK]|^88333[ABVWGH]|^88334[YI]{0,1}|^88336[ABVWGH]' " & vbCrLf & _
 "AND n.ZeitPunkt BETWEEN qende(l.zeitpunkt) - INTERVAL 1 YEAR AND qende(l.zeitpunkt) + INTERVAL 3 MONTH " & vbCrLf & _
 "WHERE l.leistung='88322' AND NOT ISNULL (n.pat_id) AND l.zeitpunkt>=qanf()" & vbCrLf & _
 "GROUP BY l.id " & vbCrLf & _
 "ORDER BY l.pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

' 156
#Else
' 150
#End If
AwN(AWlf) = "Unbekannte Impfstoffe (vorher 141)"
sql(AWlf) = "" & _
"SELECT pat_id, gesname(pat_id) PName, Zeitpunkt, Art, Inhalt FROM eintraege WHERE art='vac' AND impfart(inhalt)=999 AND NOT inhalt rlike 'tetagam|Bexsero|Nimenrix';"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
#If mitcovid Then
' 157
AwN(AWlf) = "Abrechnung Genesenenzertifikate (lauto) (vorher 143)"
sql(AWlf) = "" & _
"  SELECT f.pat_id PID, gesname(f.pat_id) PName,  DATE(fo.zeitpunkt) LEIDAT, '88371 dazu' LEIFEHLER, f.lanrid LANRID " & vbCrLf & _
"  FROM aktfv f " & vbCrLf & _
"  LEFT JOIN formular fo on fo.Pat_ID = f.pat_id AND fo.form_abk ='covge' AND fo.feld='testpositivdatum' AND fo.zeitpunkt BETWEEN qanf() AND qend() " & vbCrLf & _
"  LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.leistung='88371' AND DATE(l.ZeitPunkt)=DATE(fo.zeitpunkt) " & vbCrLf & _
"  WHERE NOT ISNULL(fo.pat_ID) AND ISNULL(l.leistung) " & vbCrLf & _
"  GROUP BY f.pat_id, DATE(fo.zeitpunkt);"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
' 158
AwN(AWlf) = "Covid-Impfungen ohne Chargenangabe (vorher 149)"
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id) PName, l.Zeitpunkt, l.Leistung, l.Charge " & vbCrLf & _
"FROM aktfv f " & vbCrLf & _
"LEFT JOIN leistungen l ON f.pat_id=l.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend() " & vbCrLf & _
"WHERE leistung RLIKE '^8833' AND COALESCE(l.Charge,'')='' " & vbCrLf & _
"ORDER BY f.pat_id, l.zeitpunkt;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
 
' 159, neuView
#Else
' 151
#End If

#If mitcovid Then
' 159
#Else
' 151
#End If
AwN(AWlf) = "Covid-Impfabrechnung ohne Angabe der Ordnungsnummer der Impfung zur Leistung"
sql(AWlf) = "" & vbCrLf & _
"SELECT f.pat_id, gesname(f.pat_id) PName, Zeitpunkt, leistung, lfbegr " & vbCrLf & _
"FROM aktfv f LEFT JOIN leistungen l USING (pat_id) " & vbCrLf & _
"WHERE leistung LIKE '88%' AND zeitpunkt>qanf() AND lfbegr NOT RLIKE '^[0-9]+\.'" & vbCrLf & _
"ORDER BY f.pat_id DESC;" & vbCrLf & _
""
mins(AWlf) = 10
maxs(AWlf) = 60
AWlf = AWlf + 1



#If mitcovid Then
' 160
#Else
' 152
#End If
AwN(AWlf) = "DAK-Modul"
sql(AWlf) = "ü"
 AWlf = AWlf + 1
 

#If mitcovid Then
' 161
#Else
' 153
#End If
AwN(AWlf) = "Keinem Patienten zugeordnete DAK-Faxe /KKH-Faxe(bitte im MySQL-Query-Browser zuordnen über 'SELECT pid,docname,eind FROM faxeinp.outa WHERE eind=...') (vorher 89)"
sql(AWlf) = "" & vbCrLf & _
"SELECT eind, pid, docname FROM faxeinp.outa " & _
"WHERE (docname LIKE '%6971042276004%' OR docname LIKE '%51180684684%') AND pid=0 " & _
"ORDER BY docname; "
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

#If mitcovid Then
' 162
#Else
' 154
#End If
AwN(AWlf) = "Evtl. nicht angekommene oder fehlerhaft benannte DAK/KKH/HEK-Einverständnis-Faxe (nicht berücksichtigbar: Techniker Kk.) (vorher 90)"
' 07433967297004 = neue Nr. DAK
' 051180684684 = KKH
' 040656961201 = HEK (Hanseatische Krankenkasse)
' 0404606626279 = TK, aber nicht per fax erwünscht
sql(AWlf) = "" & vbCrLf & _
"SELECT f.pat_id, gesname(f.pat_id) Name " & vbCrLf & _
"FROM aktfvs f " & vbCrLf & _
"JOIN faelle fa USING (fid)" & vbCrLf & _
"LEFT JOIN namen n ON n.pat_id = f.pat_id " & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.zeitpunkt BETWEEN qanf()AND qend()" & vbCrLf & _
"LEFT JOIN faxeinp.outa odak ON odak.docname RLIKE CONCAT(n.nachname,'.*',n.vorname,'.*(6971042276004|67297004|51180684684|40656961201|404606626279)')" & vbCrLf & _
"LEFT JOIN eintraege br ON br.pat_id=f.pat_id AND br.inhalt RLIKE 'Zusatzprogramm.*(gemailt|geschickt)'" & vbCrLf & _
"WHERE e.art = 'dak' AND e.inhalt LIKE 'dak: hier%' AND ISNULL(odak.docname) AND ISNULL(br.pat_id) AND fa.kkasse_2 NOT RLIKE 'Techniker' AND fa.kkasse_2 <> 'TK';"
'"LEFT JOIN faxeinp.outa odak ON odak.docname LIKE CONCAT('%',n.nachname,' %',n.vorname,'%6971042276004%') " & vbCrLf & _
" OR odak.docname LIKE CONCAT('%',n.nachname,' %',n.vorname,'%7433967297004%') " & vbCrLf & _
" OR odak.docname LIKE CONCAT('%',n.nachname,' %',n.vorname,'%Zusatzprogramm%') " & vbCrLf & _
"LEFT JOIN faxeinp.outa okkh ON okkh.docname LIKE CONCAT('%',n.nachname,' %',n.vorname,'%51180684684%') " & vbCrLf & _
"LEFT JOIN faxeinp.outa ohek ON ohek.docname LIKE CONCAT('%',n.nachname,' %',n.vorname,'%40656961201%') " & vbCrLf & _
"LEFT JOIN faxeinp.outa otk ON otk.docname LIKE CONCAT('%',n.nachname,' %',n.vorname,'%404606626279%') " & vbCrLf & _
"WHERE e.art = 'dak' AND e.inhalt LIKE 'dak: hier%' AND ISNULL(odak.docname) AND ISNULL(okkh.docname) AND ISNULL(ohek.docname) AND fa.kkasse_2 NOT RLIKE 'Techniker';"
' AND ISNULL(otk.docname);"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

#If mitcovid Then
' 163
#Else
' 155
#End If
AwN(AWlf) = "Evtl. fehlende DAK(/KKH/HEK/TK)-Makros (vorher 91)"
sql(AWlf) = "" & vbCrLf & _
"SELECT docname, pid, n.pat_id, gesname(n.pat_id) Name, DATE_FORMAT(transe,'%d.%m.%Y') gefaxt " & vbCrLf & _
", wia(n.pat_id) iArzt " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = n.pat_id AND (art='tk'OR(art='tb'AND ersteller='tk')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = n.pat_id AND (art = 'tk'OR(art='tb'AND ersteller='tk')))) tk " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = n.pat_id AND (art='gs'OR(art='tb'AND ersteller='gs')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = n.pat_id AND (art = 'gs'OR(art='tb'AND ersteller='gs')))) gs " & vbCrLf & _
            ",IF((SELECT MAX(art) FROM eintraege WHERE pat_id = n.pat_id AND (art='ah'OR(art='tb'AND ersteller='ah')))IS NULL,'',(SELECT CONCAT(LPAD(CAST(COUNT(art) AS char),4,' '),DATE_FORMAT(MAX(zeitpunkt),'(-%d.%m.%y)')) FROM eintraege WHERE pat_id = n.pat_id AND (art = 'ah'OR(art='tb'AND ersteller='ah')))) ah " & vbCrLf & _
"FROM faxeinp.outa " & vbCrLf & _
"LEFT JOIN namen n ON n.pat_id=pid " & vbCrLf & _
"LEFT JOIN eintraege e ON pid = e.pat_id AND (art= 'dak') " & vbCrLf & _
"WHERE docname RLIKE '6971042276004|7433967297004|51180684684|40656961201|404606626279' " & vbCrLf & _
"AND NOT docname RLIKE 'Attest' AND ISNULL(art) AND pid<>0 " & vbCrLf & _
"AND EXISTS(SELECT 1 FROM `faelle` f WHERE pat_id=pid AND qbeg(fanf)=qbeg(submt) AND ((SELECT kateg FROM kassenliste WHERE id=f.kid)) IN ('','BKK','BKN','EK','IKK'))" & vbCrLf & _
"ORDER BY docname;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

'#If False THEN
'AwN(AWlf) = "Mögliche DAK/KKH-Ziffern Niereninsuffizienz (97600,97601,97603)"
'sql(AWlf) = "" & vbCrLf & _
'"SELECT * FROM ( " & vbCrLf & _
'"SELECT f.pat_id, gesnameg(f.pat_id), l.leistung Lstg, l.zeitpunkt Lzp, GROUP_CONCAT(DISTINCT l1.wert SEPARATOR ', ') AlbCre, IF(MAX(l1.zeitpunkt)>MAX(l2.zeitpunkt OR ISNULL(l2.zeitpunkt)),MAX(l1.zeitpunkt),MAX(l2.zeitpunkt)) AlbCreZp, " & vbCrLf & _
'"lHbA(f.pat_id) HbA1c, lCre(f.pat_id) Crea, LEFT(GROUP_CONCAT(REPLACE(REPLACE(REPLACE(rr.rr,' mmHg',''),'mmHg',''),'mmhg','') SEPARATOR ', '),50) rr " & vbCrLf & _
'"FROM aktfvs f " & vbCrLf & _
'"LEFT JOIN eintraege e ON f.pat_id = e.pat_id " & vbCrLf & _
'"LEFT JOIN leistungen l ON l.pat_id = e.pat_id AND leistung IN ('97600','97601','97602') AND l.zeitpunkt >= SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) " & vbCrLf & _
'"AND l.zeitpunkt = (SELECT MAX(zeitpunkt) FROM leistungen WHERE pat_id = e.pat_id AND leistung IN ('97600','97601','97602')) " & vbCrLf & _
'"LEFT JOIN labor1a l1 ON l1.pat_id = f.pat_id AND ((l1.abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND l1.einheit LIKE 'mg/g %') OR (l1.abkü IN ('ALBU','ALBUMU','ALBUU') AND l1.einheit IN ('mg/l',''))) " & vbCrLf & _
'"AND l1.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) " & vbCrLf & _
'"LEFT JOIN labor2a l2 ON l2.pat_id = f.pat_id AND ((l2.abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND l2.einheit LIKE 'mg/g %') OR (l2.abkü IN ('ALBU','ALBUMU','ALBUU') AND l2.einheit IN ('mg/l',''))) " & vbCrLf & _
'"AND l2.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) " & vbCrLf & _
'"AND NOT l2.wert=l1.wert " & vbCrLf & _
'"LEFT JOIN rr ON rr.pat_id = f.pat_id AND rr.zeitpunkt >SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) " & vbCrLf & _
'"WHERE e.Art = 'dak' AND e.Inhalt LIKE 'dak: hier%' " & vbCrLf & _
'"GROUP BY f.pat_id " & vbCrLf & _
'") i"
'#END IF

#If mitcovid Then
' 164 DAK-Module
#Else
' 156 DAK-Module
#End If
'UPDATE eintraege FORCE INDEX (inhfull) SET art='dakluts' WHERE inhalt LIKE 'wie oft am tag%';
'SELECT * from eintraege FORCE INDEX (inhfull) WHERE inhalt LIKE 'Bekommen Sie regelmäßig nach bestimmten Gehstrecken%';
'UPDATE eintraege FORCE INDEX (inhfull) SET art='dakap' WHERE inhalt LIKE 'Bekommen Sie regelmäßig nach bestimmten Gehstrecken%';
'SELECT * from eintraege FORCE INDEX (inhfull) WHERE inhalt LIKE 'Schweißtest: %';
'UPDATE eintraege FORCE INDEX (inhfull) SET art='daknp' WHERE inhalt LIKE 'Schweißtest: %';
 sql(AWlf) = LiesDatei(uVerz & "dakges.qbquery")
 sql(AWlf) = Mid$(sql(AWlf), InStr(1, sql(AWlf), "SELECT", vbTextCompare))
 AwN(AWlf) = "Fehler bei DAK-Modul: Synthese (vorher 123)"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
#If mitcovid Then
 ' 165 DAK Neuropathie zur Kontrolle
#Else
 ' 157 DAK Neuropathie zur Kontrolle
#End If
 sql(AWlf) = LiesDatei(uVerz & "daknp.qbquery")
 sql(AWlf) = Mid$(sql(AWlf), InStr(1, sql(AWlf), "SELECT", vbTextCompare))
 AwN(AWlf) = "DAK-Modul-Kontrolle Neuropathie (vorher 124)"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
#If mitcovid Then
 ' 166 DAK LUTS zur Kontrolle
#Else
 ' 158 DAK LUTS zur Kontrolle
#End If
 sql(AWlf) = LiesDatei(uVerz & "daklu.qbquery")
 sql(AWlf) = Mid$(sql(AWlf), InStr(1, sql(AWlf), "SELECT", vbTextCompare))
 AwN(AWlf) = "DAK-Modul-Kontrolle LUTS  (vorher 125)"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
#If mitcovid Then
 ' 167 DAK Angiopathie zur Kontrolle
#Else
 ' 159 DAK Angiopathie zur Kontrolle
#End If
 sql(AWlf) = LiesDatei(uVerz & "dakap.qbquery")
 sql(AWlf) = Mid$(sql(AWlf), InStr(sql(AWlf), "select"))
 AwN(AWlf) = "DAK-Modul-Kontrolle Angiopathie (vorher 126)"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
#If mitcovid Then
 ' 168 DAK Hepatopathie zur Kontrolle
#Else
 ' 160 DAK Hepatopathie zur Kontrolle
#End If
 sql(AWlf) = LiesDatei(uVerz & "dakfl.qbquery")
 sql(AWlf) = Mid$(sql(AWlf), InStr(sql(AWlf), "select"))
 AwN(AWlf) = "DAK-Modul-Kontrolle Hepatopathie (vorher 127)"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
#If mitcovid Then
 ' 169 DAK Nephropathie zur Kontrolle
#Else
 ' 161 DAK Nephropathie zur Kontrolle
#End If
 sql(AWlf) = LiesDatei(uVerz & "dakne.qbquery")
 sql(AWlf) = Mid$(sql(AWlf), InStr(LCase$(sql(AWlf)), "select"))
 AwN(AWlf) = "DAK-Modul-Kontrolle Nephropathie (vorher 128)"
 mins(AWlf) = 10
 maxs(AWlf) = 80
 AWlf = AWlf + 1
 
Dim ob_tkeinschr$, ob_tkversandt$
ob_tkeinschr = "'tk.*(Modul|Programm)'"
ob_tkversandt = "'tk.*(geschickt|gefaxt)'"
#If mitcovid Then
' 170
#Else
' 162
#End If
 AwN(AWlf) = "TK-Modul-Einschreibungen (" & ob_tkeinschr & ") ohne Versandeintrag (" & ob_tkversandt & ") (vorher 152)"
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id) PName, n.gebdat, LEFT(b.Name,80) BName, IF(b.quelldatum=18991230,'',DATE_FORMAT(b.quelldatum,'%d.%m.%y')) Quelldatum, DATE_FORMAT(b.Zeitpunkt,'%d.%m.%y') Zeitpunkt " & vbCrLf & _
"FROM aktfv f " & vbCrLf & _
"LEFT JOIN namen n USING(pat_id)" & vbCrLf & _
"LEFT JOIN tmbrie b ON b.pat_id=f.pat_id AND b.Name RLIKE " & ob_tkeinschr & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.inhalt RLIKE " & ob_tkversandt & vbCrLf & _
"WHERE NOT ISNULL(b.name) AND ISNULL(e.zeitpunkt);"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

#If mitcovid Then
' 171
#Else
' 163
#End If
 AwN(AWlf) = "ergänzende Listen"
sql(AWlf) = "ü"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
  
#If mitcovid Then
' 172
#Else
' 164
#End If
 AwN(AWlf) = "Leistung 01435 zu Fall mit Grundpauschale oder taggleicher sonstiger Leistung"
sql(AWlf) = "" & _
"SELECT f.pat_id, gesname(f.pat_id), l.zeitpunkt " & vbCrLf & _
",(SELECT zeitpunkt FROM leistungen WHERE pat_id =f.pat_id AND zeitpunkt BETWEEN qanf() AND qend() AND leistung IN (03001,03002,03003,03004,03005,03011,03012,03013,03014,03015)) ZeitpunktGrundPauschale" & vbCrLf & _
", COALESCE((SELECT group_concat(leistung) FROM leistungen WHERE pat_id =f.pat_id AND date(zeitpunkt)=DATE(l.zeitpunkt) AND leistung<>l.leistung),'') LeistGleicherTag" & vbCrLf & _
"FROM aktfv f LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.ZeitPunkt BETWEEN qanf() AND qend()" & vbCrLf & _
"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.inhalt RLIKE " & ob_tkversandt & vbCrLf & _
"WHERE l.leistung='01435'" & vbCrLf & _
"ORDER BY f.pat_id;"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

#If mitcovid Then
' 173
#Else
' 165
#End If
 AwN(AWlf) = "Überweisungen mit unbekannten Betriebsstättennummern"
sql(AWlf) = "" & _
"SELECT p.fpatnr, 18900101+INTERVAL fvon DAY von, 18900101+INTERVAL fbis DAY bis, REGEXP_REPLACE(CONVERT(p.FMemo USING latin1), '[[:cntrl:]]+', ' ') patfall_FMemo, p.fsurogat " & vbCrLf & _
"FROM patfall p" & vbCrLf & _
"LEFT JOIN epraxis b ON p.FMemo RLIKE b.FBetriebsnr" & vbCrLf & _
"WHERE p.FScheintyp=0 AND p.FTarif=2" & vbCrLf & _
"AND b.FBezeichnung IS NULL AND 18900101+INTERVAL fvon DAY >= STR_TO_DATE(CONCAT(YEAR(NOW()-INTERVAL " & Verspätung & " DAY),'-',((MONTH(NOW()-INTERVAL " & Verspätung & " DAY)-1) DIV 3)*3+1,'-1'),'%Y-%m-%e')" & vbCrLf & _
"ORDER BY fpatnr;"
 obmo(AWlf) = True
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1
 
#If mitcovid Then
' 174
#Else
' 166
#End If
 AwN(AWlf) = "Möglicherweise fehlende 99055 (Schulungsabschluss)"
sql(AWlf) = _
"SELECT Pat_id,PName,DTyp,Leistung,Zeitpunkt,Bereich,Geszahl,Schulungsdaten `Daten vorausgegangener Leistungen`,Termine `Termine nach heute` FROM (" & vbCrLf & _
" SELECT i.*,LAST_VALUE(geszahl)OVER(PARTITION BY pat_id)lv FROM (" & vbCrLf & _
"  SELECT l.Pat_ID,gesname(l.pat_id) PName" & vbCrLf & _
"  ,(SELECT COUNT(0)FROM leistungen WHERE pat_id=l.pat_id AND leistung=l.Leistung AND zeitpunkt BETWEEN l.zeitpunkt-INTERVAL 1 YEAR AND l.zeitpunkt-INTERVAL 1 SECOND)+1 geszahl" & vbCrLf & _
"  ,CONCAT(g.minZdUE,'-',g.maxZdUE)Bereich,g.minZdUE" & vbCrLf & _
"  ,d.ityp DTyp/*,COUNT(0)OVER(PARTITION BY l.id)lzahl,l.id*/,l.leistung,l.zeitpunkt" & vbCrLf & _
"  ,(SELECT GROUP_CONCAT(DATE_FORMAT(zeitpunkt,'%d.%m.%y')ORDER BY zeitpunkt SEPARATOR ', ')FROM leistungen WHERE pat_id=l.pat_id AND leistung=l.Leistung AND zeitpunkt BETWEEN l.zeitpunkt-INTERVAL 1 YEAR AND l.zeitpunkt-INTERVAL 1 SECOND)Schulungsdaten" & vbCrLf & _
"  ,sl.leistung `99055`" & vbCrLf & _
"  ,(SELECT GROUP_CONCAT(CONCAT(DATE_FORMAT(zp,'%d.%m.%y('),raum,')')ORDER BY zp SEPARATOR ', ') termin FROM termine WHERE pid=f.pat_id AND zp>=NOW())termine" & vbCrLf & _
"  FROM aktfv f" & vbCrLf & _
"  JOIN leistungen l ON l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qanf()AND qend()" & vbCrLf & _
"  LEFT JOIN dtypen d ON d.pat_id=f.pat_id" & vbCrLf & _
"  JOIN genehmigungen g ON g.leistung=l.leistung AND(d.ityp=g.dtyp OR d.ityp IS NULL OR g.dtyp IS NULL)AND g.minzdue>0" & vbCrLf & _
"  LEFT JOIN leistungen sl ON sl.Pat_ID=f.pat_id AND sl.Leistung='99055'AND sl.ZeitPunkt BETWEEN qanf()AND qend()" & vbCrLf & _
"  WHERE sl.Leistung IS NULL" & vbCrLf & _
"  GROUP BY l.id" & vbCrLf & _
"  HAVING geszahl>=minzdue" & vbCrLf & _
"  )i" & vbCrLf & _
")i WHERE lv=geszahl"
 obmo(AWlf) = False
 mins(AWlf) = 10
 maxs(AWlf) = 120
 AWlf = AWlf + 1

#If mitcovid Then
' 175
#Else
' 167
#End If

AwN(AWlf) = "Zu viele 01647 oder 01648"
sql(AWlf) = _
"SELECT SUM(IF(leistung=01647,lzahl,0))`01647-Zahl`,SUM(IF(leistung=01648,lzahl,0))`01648-Zahl`," & vbCrLf & _
" GROUP_CONCAT(zeitpunkt SEPARATOR ',  ')Zeitpunkte," & vbCrLf & _
" l.Pat_id, gesnameg(l.pat_id)PName" & vbCrLf & _
" FROM aktfv f" & vbCrLf & _
" JOIN  leistungen l" & vbCrLf & _
" ON l.Pat_ID=f.pat_id AND l.ZeitPunkt BETWEEN qanf() AND qend() AND l.Leistung IN(01647,01648)" & vbCrLf & _
" GROUP BY l.pat_id" & vbCrLf & _
" HAVING `01647-Zahl`+`01648-Zahl`>1"
 obmo(AWlf) = False
 mins(AWlf) = 10
 maxs(AWlf) = 120
 AWlf = AWlf + 1

#If mitcovid Then
' 176
#Else
' 168
#End If
 AwN(AWlf) = "weitere Diagnosen"
sql(AWlf) = "ü"
 mins(AWlf) = 10
 maxs(AWlf) = 60
 AWlf = AWlf + 1

#If mitcovid Then
' 177
#Else
' 169
#End If
 AwN(AWlf) = "Inkretinmimetika ohne schwere Niereninsuffizienz, ohne Metformin und ohne Metforminunverträglichkeit (T88.7 mit ""Metformin"" in Diagnosentext oder Zusatztext oder N18.4 oder N18.5 dazu)"
sql(AWlf) = _
"SELECT * FROM (" & vbCrLf & _
"SELECT CASE WHEN n.obh THEN 'H' WHEN n.obk THEN 'K' WHEN n.obs THEN 'S' END Arzt" & vbCrLf & _
",ityp `DM-Typ`,f.pat_id,gesname(f.pat_id)PName,rz.zeitpunkt,rz.medikament/*,glp.zeitpunkt,glp.medikament*/" & vbCrLf & _
",ROW_NUMBER()OVER(PARTITION BY pat_id ORDER BY rz.zeitpunkt DESC)rn" & vbCrLf & _
",COALESCE((SELECT GROUP_CONCAT(concat(zeitpunkt,': ',inhalt)SEPARATOR '; ') FROM eintraege WHERE pat_id=f.pat_id AND inhalt LIKE '%Metformin%'),'') `Einträge mit ""Metformin""`" & vbCrLf & _
"/* , n.obk,n.obs,n.obh, d.ICD,d.DiagText,d.DiagAttr,  f.*,glp.* ,gma.* */" & vbCrLf & _
"FROM aktfvz f" & vbCrLf & _
"JOIN namen n USING(pat_id)" & vbCrLf & _
"LEFT JOIN dtypen USING (pat_id)" & vbCrLf & _
"/*JOIN lmp glp on glp.pat_id=f.pat_id JOIN medarten gma ON gma.medikament=glp.medanfang AND gma.glp1<>0 AND glp.medikament NOT LIKE '%Wegovy%' AND CONCAT(bemerkung,' ',Grund) NOT RLIKE 'selbst|privat'*/" & vbCrLf & _
"JOIN rezepteintraege rz ON rz.Pat_ID=f.pat_id AND rz.medikament RLIKE 'Liraglutid|victoza|ozempic|truliciy|mounjaro'and rezklkurz='rp' AND zeitpunkt>qanf()-INTERVAL 6 MONTH AND obst=0" & vbCrLf & _
"LEFT JOIN " & vbCrLf & _
"(SELECT metf.pat_id from lmp metf JOIN medarten mma ON mma.medikament=metf.medanfang AND mma.metf<>0)i" & vbCrLf & _
"ON i.pat_id=f.pat_id" & vbCrLf & _
"LEFT JOIN diagview d ON d.pat_id=f.pat_id AND d.ICD='T88.7' AND d.DiagSicherheit IN (' ','G') AND d.obdauer<>0" & vbCrLf & _
"AND (d.diagtext LIKE '%Metformin%' OR d.DiagAttr LIKE '%Metformin%')" & vbCrLf & _
"LEFT JOIN diagview ni ON ni.pat_id=f.pat_id AND ni.ICD RLIKE 'N18.[45]' AND ni.DiagSicherheit IN (' ','G') AND ni.obdauer<>0" & vbCrLf & _
" WHERE i.pat_id IS NULL AND d.ICD IS NULL AND ni.ICD IS NULL" & vbCrLf & _
")i WHERE rn=1" & vbCrLf & _
"ORDER BY Arzt,`DM-Typ`,pat_id"
 mins(AWlf) = 10
 maxs(AWlf) = 6000
 AWlf = AWlf + 1

#If mitcovid Then
' 178
#Else
' 170
#End If
 AwN(AWlf) = "Bempedoinsäure ohne Statin und ohne Statinunverträglichkeit (T88.7 mit ""Statin"" in Diagnosentext oder Zusatztext)"
sql(AWlf) = _
"SELECT CASE WHEN n.obh THEN 'H' WHEN n.obk THEN 'K' WHEN n.obs THEN 'S' END Arzt" & vbCrLf & _
" ,f.pat_id,gesname(f.pat_id)PName,glp.zeitpunkt,glp.medikament" & vbCrLf & _
" ,COALESCE((SELECT GROUP_CONCAT(concat(zeitpunkt,': ',inhalt)SEPARATOR '; ') FROM eintraege WHERE pat_id=f.pat_id AND inhalt LIKE '%Statin%'),'') `Einträge mit ""Statin""`" & vbCrLf & _
"/* , n.obk,n.obs,n.obh, d.ICD,d.DiagText,d.DiagAttr,  f.*,glp.* ,gma.* */" & vbCrLf & _
"FROM aktfvz f" & vbCrLf & _
"JOIN namen n USING (pat_id)" & vbCrLf & _
" JOIN lmp glp on glp.pat_id=f.pat_id JOIN medarten gma ON gma.medikament=glp.medanfang AND gma.medikament RLIKE 'Nilemdo|Nustendi'" & vbCrLf & _
"LEFT JOIN " & vbCrLf & _
" (SELECT metf.pat_id,metf.medikament from lmp metf JOIN medarten mma ON mma.medikament=metf.medanfang AND mma.HMG<>0)i" & vbCrLf & _
"  ON i.pat_id=f.pat_id" & vbCrLf & _
" LEFT JOIN diagview d ON d.pat_id=f.pat_id AND d.ICD='T88.7' AND d.DiagSicherheit IN (' ','G') AND d.obdauer<>0" & vbCrLf & _
"  AND CONCAT(d.diagtext,' ',d.diagattr) LIKE '%statin%'" & vbCrLf & _
" WHERE i.pat_id IS NULL AND d.ICD IS NULL" & vbCrLf & _
"  ORDER BY Arzt,f.pat_id"
 mins(AWlf) = 10
 maxs(AWlf) = 6000
 AWlf = AWlf + 1

#If mitcovid Then
' 179
#Else
' 171
#End If
 AwN(AWlf) = "Sonobefunde mit falscher Art"
sql(AWlf) = _
"SELECT pat_id,gesname(pat_id)PName,Zeitpunkt,Art,Ersteller,Änderer,Inhalt FROM eintraege e WHERE inhalt RLIKE ':'" & vbCrLf & _
"AND inhalt RLIKE '^(Abdomen|(Hals(schlagadern|venen|arterien|weichteile|sono)|Wei?chteile|Pleura|(Sono )?Schid?l?l?dd?d?r?r?s?ür?s?s?e?|(Bein|Arm)(venen|arterien)( (re|li(nks)?|bds.))?)|Nierenarterien?|Restharn|SD|Darmarterien|Abdomen|Belastung(suntersuchung)?).*:'" & vbCrLf & _
"AND NOT inhalt RLIKE '(^Belastungsdyspnoe|szinti|ekg|-ct|^sd-(ex|unterfunktion)|sd -?(op|wurde)|sd code|untersuchung|lmu|mvz|nicht|ko in|sd-werte|sds-|Dr.|empfohlen).*:|(.*Antikörper abnehmen| vor langer Zeit| Verdächtiger Befund|Abdomen-Sono 7/25:)'" & vbCrLf & _
"AND NOT art IN ('Sono','dup','dop','doppler','duplex','ana','utxt','caro','anal','lar','plar')" & vbCrLf & _
"ORDER BY pat_id"
 mins(AWlf) = 10
 maxs(AWlf) = 6000
 AWlf = AWlf + 1

' neuView
 AwN(AWlf) = "- Ende -"
 sql(AWlf) = "ü"
 AWlf = AWlf + 1
 
 
 
' 19.2.11
' AwN(AWlf) = "97350A, 97360A, 97370B, 92282, 92278, 92281, 92277 für LKK-Patienten:"
' sql(AWlf) = _
' "SELECT f.pat_id, nachname, vorname, leistung, kateg FROM (" & aktf & " " & vbCrLf & _
'"LEFT JOIN (SELECT * FROM `leistungen` WHERE leistung IN ( '97350A','97360A','97370B', '92282', '92278', '92281', '92277')) AS leist " & vbCrLf & _
'"ON f.fid = leist.fid) " & vbCrLf & _
'"LEFT JOIN  `kassenliste` AS kl ON f.vknr = kl.vknr " & vbCrLf & _
'"WHERE (kateg = 'LKK' AND NOT ISNULL(leist.leistung)) GROUP BY f.pat_id, kateg"
' mins(AWlf) = 5
' maxs(AWlf) = 18
' AWlf = AWlf + 1

' 19.2.11
' AwN(AWlf) = "97350B, 97360B für Nicht-LKK-Patienten:"
' sql(AWlf) = _
'"SELECT f.pat_id, nachname, vorname, leistung, kateg FROM (" & aktf & " " & vbCrLf & _
'"LEFT JOIN (SELECT * FROM `leistungen` WHERE leistung IN ( '97350B','97360B')) AS leist " & vbCrLf & _
'"ON f.fid = leist.fid) " & vbCrLf & _
'"LEFT JOIN  `kassenliste` AS kl ON f.vknr = kl.vknr " & vbCrLf & _
'"WHERE (kateg <> 'LKK' AND kateg <> 'AOK' AND NOT ISNULL(leist.leistung))  GROUP BY f.pat_id, kateg"
' mins(AWlf) = 5
' maxs(AWlf) = 18
' AWlf = AWlf + 1
 
' 19.2.11
' AwN(AWlf) = "97310, 97320, 97321, 97312, 97322, 97313, 97323, 97333, 97314, 97324 für Nicht-AOK-Patienten:"
' sql(AWlf) = _
'"SELECT f.pat_id, nachname, vorname, leistung, kateg FROM (" & aktf & " " & vbCrLf & _
'"LEFT JOIN (SELECT * FROM `leistungen` WHERE leistung IN ('97310', '97320', '97321', '97312', '97322', '97313', '97323', '97333', '97314', '97324')) AS leist " & vbCrLf & _
'"ON f.fid = leist.fid) " & vbCrLf & _
'"LEFT JOIN  `kassenliste` AS kl ON f.vknr = kl.vknr " & vbCrLf & _
'"WHERE (kateg <> 'LKK' AND kateg <> 'AOK' AND NOT ISNULL(leist.leistung))  GROUP BY f.pat_id, kateg"
' mins(AWlf) = 5
' maxs(AWlf) = 18
' AWlf = AWlf + 1
 
' AwN(AWlf) = "97370B für nicht- BKK, BKN oder EK -Patienten:"
' sql(AWlf) = _
'"SELECT f.pat_id, nachname, vorname, leist.leistung, kateg FROM (" & aktf & " " & vbCrLf & _
'"LEFT JOIN (SELECT * FROM `leistungen` WHERE leistung IN ( ""97370B"")) AS leist " & vbCrLf & _
'"ON f.fid = leist.fid) " & vbCrLf & _
'"LEFT JOIN `kassenliste` AS kl ON f.vknr = kl.vknr " & vbCrLf & _
'"WHERE (kateg <> ""BKK"" AND kateg <> ""BKN"" AND kateg <> ""EK"" AND NOT ISNULL(leist.leistung))  GROUP BY f.pat_id, kateg"
' mins(AWlf) = 5
' maxs(AWlf) = 18
' AWlf = AWlf + 1

#If alt Then
' sql(AWlf) = _
' "SELECT * FROM " & vbCrLf & _
' "(SELECT * FROM " & vbCrLf & _
' "(SELECT f.pat_id AS pat_id, gesname(f.pat_id) Name, zeitpunkt, fotostad, l.leistung AS Leistung, kateg, dmpklass, diab.icd DTyp FROM " & aktf & " " & vbCrLf & _
' "LEFT JOIN (SELECT fid,leistung FROM `leistungen` WHERE leistung IN ('97314','97324')) AS l ON f.fid = l.fid " & vbCrLf & _
' "LEFT JOIN (SELECT * FROM (SELECT MAX(pat_id) AS PID, MAX(MID(name,INSTR(name,""WA ""),5)) fotoStad, Quelldatum zeitpunkt FROM tmbrie GROUP BY pat_id, zeitpunkt) innen WHERE " & vbCrLf & _
' "fotoStad = 'WA 0D' OR (fotostad LIKE 'WA %' AND NOT fotostad IN ('WA 1A','WA 2A') AND NOT fotostad LIKE 'WA 0%' AND NOT fotostad LIKE 'WA -%' ) " & vbCrLf & _
' ") AS dok ON f.pat_id = dok.pid " & vbCrLf & _
' "AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & _
' "LEFT JOIN `diagnosen` diab ON (f.pat_id = diab.pat_id AND diab.obdauer <> 0 AND diab.diagsicherheit<> 'A') AND diab.icd REGEXP '^E1[0-4]\.' " & vbCrLf & _
' "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr AND f.ik=kl.ik " & vbCrLf & _
' "LEFT JOIN `namen` n ON f.pat_id = n.pat_id ORDER BY leistung DESC) AS innen GROUP BY pat_id) AS innen " & vbCrLf & _
' "WHERE true " & vbCrLf & _
' "AND NOT ISNULL(fotostad) " & vbCrLf & _
' "AND NOT ISNULL(dtyp) " & vbCrLf & _
' "AND (ISNULL(leistung)) " & vbCrLf & _
' "AND kateg NOT IN ('SA','') " & vbCrLf & _
' "AND (((dmpklass=2 OR (dmpklass=3 AND dmpbeg<=qend()))) OR kateg IN ('LKK','PBe') " & vbCrLf & _
' "AND true "
'#ElseIf schonwiederalt Then
'' 26 "Fehlende 97314, 97324 für Fußsyndrom nach Fotostadium"
'' MAX(IF(REPLACE(MID(name,INSTR(name,'WA ')+3,1),'-','0')>2,2,REPLACE(MID(name,INSTR(name,'WA ')+3,1),'-','0')))+1
'' AND COALESCE(Dggel,0)=0
' sql(AWlf) = _
' "SELECT f.pat_id, gesname(f.pat_id) Name, d.fotostad,DATE(qdat) qdat, diab.icd, kateg FROM aktfv f " & vbCrLf & _
' "LEFT JOIN leistungen l ON f.fid=l.fid AND leistung IN ('97314','97324') " & vbCrLf & _
' "LEFT JOIN (SELECT MAX(fotostad) fotostad,MAX(pat_id) pat_id,DATE(MAX(IF(quelldatum<20000101,zeitpunkt,quelldatum))) qdat, MAX(zeitpunkt) zp FROM (SELECT MID(name,INSTR(name,'WA '),5) fotostad, b.* FROM tmbrie b) d  WHERE fotoStad = 'WA 0D' OR (fotostad LIKE 'WA %' AND NOT fotostad IN ('WA 1A','WA 2A') AND NOT fotostad LIKE 'WA 0%' AND NOT fotostad LIKE 'WA -%' ) GROUP BY b.pat_id) d ON b.pat_id=f.pat_id AND qdat BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
' "LEFT JOIN (SELECT pat_id, MAX(icd) icd FROM `diagnosen` WHERE obdauer <> 0 AND diagsicherheit<> 'A' AND icd REGEXP '^E1[0-4]\.' GROUP BY pat_id) diab ON f.pat_id = diab.pat_id " & vbCrLf & _
' "LEFT JOIN namen n ON n.pat_id=f.pat_id " & vbCrLf & _
' "WHERE true " & vbCrLf & _
' "AND ISNULL(leistung) " & vbCrLf & _
' "AND NOT ISNULL(fotostad) AND fotostad<>'' AND fotostad<>'wa -.' " & vbCrLf & _
' "AND NOT ISNULL(diab.icd) " & vbCrLf & _
' "AND NOT ISNULL(kateg) AND kateg NOT IN ('SA','') " & vbCrLf & _
' "AND (((dmpklass=2 OR (dmpklass=3 AND dmpbeg<=qend()))) OR kateg IN ('LKK','PBe')) " & vbCrLf & _
' "ORDER BY f.pat_id "
'#ElseIf auschonwiederalt Then
'' 26
'' AwN(AWlf) = "Fehlende 97314, 97324 für Fußsyndrom nach Fotostadium"
' AwN(AWlf) = "-"
'  sql(AWlf) = "-"
' ' Sozialamt
' mins(AWlf) = 7
' maxs(AWlf) = 120
' AWlf = AWlf + 1
'
'
'
'
'#If False Then
'' 116 ' in 33
' AwN(AWlf) = "Mehr als 1 Leistung 03220/03220H im Behandlungsfall"
' sql(AWlf) = "" ' vbCrLf & _
'"SELECT * FROM (SELECT " & vbCrLf & _
'"(SELECT SUM(lzahl) FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung IN ('03220','03220H') AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Aktzahl, " & vbCrLf & _
'"gesname(f.pat_id), f.pat_id " & vbCrLf & _
'", (SELECT GROUP_CONCAT(zeitpunkt SEPARATOR ', ') FROM leistungen l WHERE pat_id=f.pat_id AND l.leistung IN ('03220','03220H') AND l.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") Zpkte " & vbCrLf & _
'"FROM aktfvs f) i WHERE aktzahl>1;"
' mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
'
' ' 123
' AwN(AWlf) = "Fehler bei DAK-Modul: Neuropathie (lauto)"
' '  AND COALESCE(Dggel,0
'sql(AWlf) = _
'"SELECT * FROM ( " & vbCrLf & _
'" SELECT i.Pat_ID PID,i.DAKab,lztLeiTag,lztLei,aktLeiTag,aktLei,vorlZl,Eintr,Dicd,Nicd " & vbCrLf & _
'" ,CONCAT(CASE WHEN lztlei='' OR lztlei='97560' THEN " & vbCrLf & _
'"   CASE " & vbCrLf & _
'"     WHEN eintr='' AND nicd<>'' THEN " & vbCrLf & _
'"       IF(aktlei IN ('97560','97562') AND lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR),CONCAT(aktlei,' streichen'),'-') " & vbCrLf & _
'"     ELSE " & vbCrLf & _
'"       CASE WHEN Epath OR nicd<>'' THEN " & vbCrLf & _
'"         CASE " & vbCrLf & _
'"           WHEN aktlei='97561' THEN '-' " & vbCrLf & _
'"           WHEN aktlei='' THEN '97561 dazu' " & vbCrLf & _
'"           ELSE CONCAT('97561 statt ',aktlei) " & vbCrLf & _
'"           END " & vbCrLf & _
'"       ELSE " & vbCrLf & _
'"         CASE " & vbCrLf & _
'"           WHEN lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) THEN IF(aktlei='','-',CONCAT(aktlei,' streichen')) " & vbCrLf & _
'"         ELSE " & vbCrLf & _
'"           CASE " & vbCrLf & _
'"             WHEN aktlei='97560' THEN '-' " & vbCrLf & _
'"             WHEN aktlei='' THEN '97560 dazu' " & vbCrLf & _
'"             ELSE CONCAT('97560 statt ',aktlei) " & vbCrLf & _
'"           END " & vbCrLf & _
'"         END " & vbCrLf & _
'"       END " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'"   END " & vbCrLf & _
'" ELSE " & vbCrLf & _
'"    CASE WHEN vorlzl>=2 THEN CONCAT(aktlei,' streichen') " & vbCrLf & _
'"    ELSE " & vbCrLf & _
'"      CASE " & vbCrLf & _
'"       WHEN aktlei='' THEN '97562 dazu' " & vbCrLf & _
'"       WHEN aktlei<>'97562' THEN CONCAT('97562 statt ',aktlei) " & vbCrLf & _
'"       ELSE '-' " & vbCrLf & _
'"      END " & vbCrLf & _
'"    END " & vbCrLf & _
'" END ,IF(eintr<>'' AND matel='',', 97563 dazu','')) LeiFehler " & vbCrLf & _
'",CASE " & vbCrLf & _
'"   WHEN lztlei>'97560' OR aktlei>'97560' OR Epath OR nicd<>'' THEN " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[^47]',CONCAT(LEFT(dicd,4),IF(MID(dicd,5,1)='9',CONCAT('4',MID(dicd,6)),IF(obf,'73','75')),' statt ',dicd),'-') " & vbCrLf & _
'"   ELSE " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[4]',CONCAT(LEFT(dicd,4),'9',MID(dicd,6),' statt ',dicd),'-') " & vbCrLf & _
'" END DicdFehler " & vbCrLf & _
'", IF((lztlei>'97560' OR aktlei>'97560' OR Epath) AND nicd='','G63.2 dazu','-') NicdFehler " & vbCrLf & _
'", IF(lztlei>'97560' OR aktlei>'97560' OR Epath OR nicd<>'',1,0) obpath " & vbCrLf & _
'", DATE(IF(einh IS NULL," & qtEnd(FristS) & ",ezp)) LeiDat, LANRid " & vbCrLf & _
'"FROM ( " & vbCrLf & _
'"SELECT f.pat_id, n.dakab " & vbCrLf & _
'", COALESCE(DATE(lztl.zeitpunkt),'') lztleitag, COALESCE(LEFT(lztl.leistung,5),'') lztlei " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'", COALESCE(DATE(aktl.zeitpunkt),'') aktleitag, COALESCE(LEFT(aktl.leistung,5),'') aktlei " & vbCrLf & _
'", (SELECT COUNT(0) FROM leistungen WHERE pat_id=f.pat_id AND leistung IN ('97560','97561','97562') AND YEAR(zeitpunkt)=YEAR(" & qtAnf(FristS) & ") AND zeitpunkt<" & qtAnf(FristS) & ") vorlzl " & vbCrLf & _
'", COALESCE(LEFT(e.inhalt,20),'') Eintr,instr(e.inhalt,'stör')<>0 EPath, COALESCE(dd.icd,'') Dicd " & vbCrLf & _
'", COALESCE((SELECT icd FROM diagview WHERE pat_id=f.pat_id AND gicd IN ('G59.0','G63.2','G99.0'))=0 LIMIT 1),'') Nicd " & vbCrLf & _
'", COALESCE((SELECT leistung FROM leistungen WHERE pat_id=f.pat_id AND leistung='97563' AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " LIMIT 1),'') matel " & vbCrLf & _
'", IF((SELECT 1 FROM diagview WHERE pat_id=f.pat_id AND gicd RLIKE '^L89' LIMIT 1) IS NULL,1,0) obf " & vbCrLf & _
'", e.inhalt einh, e.zeitpunkt ezp, f.LANRid " & vbCrLf & _
'"FROM aktfvs f " & vbCrLf & _
'"LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
'"LEFT JOIN leistungen lztl ON lztl.pat_id=f.pat_id AND lztl.leistung IN ('97560','97561','97562') AND " & vbCrLf & _
'"   lztl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97560','97561','97562') AND lm.zeitpunkt<" & qtAnf(FristS) & ") " & vbCrLf & _
'"LEFT JOIN leistungen aktl ON aktl.pat_id=f.pat_id AND aktl.leistung IN ('97560','97561','97562') AND " & vbCrLf & _
'"   aktl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97560','97561','97562') AND lm.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
'"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art='daknp' AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
'"LEFT JOIN diagnosen dd ON dd.pat_id=f.pat_id AND dd.ricdok RLIKE '^E1[0-4]\.' AND dd.obdauer<>0 " & vbCrLf & _
'"WHERE n.dakab>20000101 AND (dakab<=DATE(lztl.zeitpunkt) OR ISNULL(lztl.zeitpunkt))) i " & vbCrLf & _
'") i " & vbCrLf & _
'"WHERE leifehler<>'-' OR dicdfehler<>'-' OR nicdfehler<>'-' " & vbCrLf & _
'"GROUP BY PID,leifehler,dicdfehler,nicdfehler "
' mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
'
'' 124
'AwN(AWlf) = "Fehler bei DAK-Modul: LUTS (lauto)"
'' AND COALESCE(Dggel,0)=0
'sql(AWlf) = vbCrLf & _
'"SELECT * FROM ( " & vbCrLf & _
'" SELECT i.Pat_ID PID,i.DAKab,lztLeiTag,lztLei,aktLeiTag,aktLei,vorlZl,Eintr,Dicd,Nicd " & vbCrLf & _
'" ,CASE WHEN lztlei='' OR lztlei='97570' THEN " & vbCrLf & _
'"   CASE " & vbCrLf & _
'"     WHEN eintr='-' AND nicd<>'' THEN " & vbCrLf & _
'"       IF(aktlei IN ('97570','97572') AND lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR),CONCAT(aktlei,' streichen'),'-') " & vbCrLf & _
'"     ELSE " & vbCrLf & _
'"       CASE WHEN Epath OR nicd<>'' THEN " & vbCrLf & _
'"         CASE " & vbCrLf & _
'"           WHEN aktlei='97571' THEN '-' " & vbCrLf & _
'"           WHEN aktlei='' THEN '97571 dazu' " & vbCrLf & _
'"           ELSE CONCAT('97571 statt ',aktlei) " & vbCrLf & _
'"           END " & vbCrLf & _
'"       ELSE " & vbCrLf & _
'"         CASE " & vbCrLf & _
'"           WHEN lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) OR eintr='-' THEN IF(aktlei='','-',CONCAT(aktlei,' streichen')) " & vbCrLf & _
'"         ELSE " & vbCrLf & _
'"           CASE " & vbCrLf & _
'"             WHEN aktlei='97570' OR Eintr='-' THEN '-' " & vbCrLf & _
'"             WHEN aktlei='' THEN '97570 dazu' " & vbCrLf & _
'"             ELSE CONCAT('97570 statt ',aktlei) " & vbCrLf & _
'"           END " & vbCrLf & _
'"         END " & vbCrLf & _
'"       END " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'"   END " & vbCrLf & _
'" ELSE " & vbCrLf & _
'"    CASE WHEN vorlzl>=2 THEN CONCAT(aktlei,' streichen') " & vbCrLf & _
'"    ELSE " & vbCrLf & _
'"      CASE " & vbCrLf & _
'"       WHEN aktlei='' THEN '97572 dazu' " & vbCrLf & _
'"       WHEN aktlei<>'97572' THEN CONCAT('97572 statt ',aktlei) " & vbCrLf & _
'"       ELSE '-' " & vbCrLf & _
'"      END " & vbCrLf & _
'"    END " & vbCrLf & _
'" END LeiFehler " & vbCrLf & _
'",CASE " & vbCrLf & _
'"   WHEN lztlei>'97570' OR aktlei>'97570' OR Epath OR NOT nicd='' THEN " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[^47]',CONCAT(LEFT(dicd,4),IF(MID(dicd,5,1)='9',CONCAT('4',MID(dicd,6)),IF(obf,'73','75')),' statt ',dicd),'-') " & vbCrLf & _
'"   ELSE " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[4]',CONCAT(LEFT(dicd,4),'9',MID(dicd,6),' statt ',dicd),'-') " & vbCrLf & _
'" END DicdFehler " & vbCrLf & _
'", IF((lztlei>'97570' OR aktlei>'97570' OR Epath) AND nicd='','N31.2 dazu','-') NicdFehler " & vbCrLf & _
'", IF(lztlei>'97570' OR aktlei>'97570' OR Epath OR NOT nicd='',1,0) obpath " & vbCrLf & _
'", DATE(IF(einh IS NULL," & qtEnd(FristS) & ",ezp)) LeiDat, LANRid " & vbCrLf & _
'"FROM ( " & vbCrLf & _
'"SELECT f.pat_id, n.dakab " & vbCrLf & _
'", COALESCE(DATE(lztl.zeitpunkt),'') lztleitag,COALESCE(LEFT(lztl.leistung,5),'') lztlei " & vbCrLf & _
'", COALESCE(DATE(aktl.zeitpunkt),'') aktleitag,COALESCE(LEFT(aktl.leistung,5),'') aktlei " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'", (SELECT COUNT(0) FROM leistungen WHERE pat_id=f.pat_id AND leistung IN ('97570','97571','97572') AND YEAR(zeitpunkt)=YEAR(" & qtAnf(FristS) & ") AND zeitpunkt<" & qtAnf(FristS) & ") vorlzl " & vbCrLf & _
'", IF(e.inhalt IS NULL,'-','ja') Eintr " & vbCrLf & _
'", COALESCE(MID(inhalt,instr(inhalt,'entleeren?')+11) RLIKE '^n|^Inko' OR MID(inhalt,instr(inhalt,'Wasserlassen)?')+15) RLIKE '^öfter|^mehr|^[2-9]' OR MID(inhalt,instr(inhalt,'abgeschwächt?')+14) RLIKE '^j' OR MID(inhalt,instr(inhalt,'zu entleeren?')+14) RLIKE '^j' OR MID(inhalt,instr(inhalt,'gang auf?')+10) RLIKE '^j' OR MID(inhalt,instr(inhalt,'beschreiben)')+12) RLIKE '^j',0) EPath, " & vbCrLf & _
'"  COALESCE(dd.icd,'') Dicd,COALESCE(nd.icd,'') Nicd " & vbCrLf & _
'", IF((SELECT 1 FROM diagview WHERE pat_id=f.pat_id AND gicd RLIKE '^L89' LIMIT 1) IS NULL,1,0) obf " & vbCrLf & _
'", e.inhalt einh, e.zeitpunkt ezp, f.LANRid " & vbCrLf & _
'"FROM aktfvs f " & vbCrLf & _
'"LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
'"LEFT JOIN leistungen lztl ON lztl.pat_id=f.pat_id AND lztl.leistung IN ('97570','97571','97572') AND " & vbCrLf & _
'"   lztl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97570','97571','97572') AND lm.zeitpunkt<" & qtAnf(FristS) & ") " & vbCrLf & _
'"LEFT JOIN leistungen aktl ON aktl.pat_id=f.pat_id AND aktl.leistung IN ('97570','97571','97572') AND " & vbCrLf & _
'"   aktl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97570','97571','97572') AND lm.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
'"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art='dakluts' AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
'"LEFT JOIN diagview dd ON dd.pat_id=f.pat_id AND dd.gicd RLIKE '^E1[0-4]\.' AND dd.obdauer<>0 " & vbCrLf & _
'"LEFT JOIN diagview nd ON nd.pat_id=f.pat_id AND nd.gicd IN ('N31.1','N31.2') AND nd.obdauer<>0 " & vbCrLf & _
'"WHERE n.dakab>20000101 AND (dakab<=DATE(lztl.zeitpunkt) OR ISNULL(lztl.zeitpunkt))) i " & vbCrLf & _
'") i " & vbCrLf & _
'"WHERE leifehler<>'-' OR dicdfehler<>'-' OR nicdfehler<>'-' " & vbCrLf & _
'"GROUP BY PID,leifehler,dicdfehler,nicdfehler " & vbCrLf & _
'";"
' mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
'
'' 125
' AwN(AWlf) = "Fehler bei DAK-Modul: Angiopathie (lauto)"
' ' AND COALESCE(Dggel,0)=0
'sql(AWlf) = _
'"SELECT * FROM ( " & vbCrLf & _
'" SELECT i.Pat_ID PID,i.DAKab,lztLeiTag,lztLei,aktLeiTag,aktLei,vorlZl,Eintr,Dicd,Nicd1,Nicd2 " & vbCrLf & _
'" ,CASE WHEN lztlei='' OR lztlei='97580' THEN " & vbCrLf & _
'"   CASE " & vbCrLf & _
'"     WHEN eintr='-' AND (nicd1<>'' OR nicd2<>'') THEN " & vbCrLf & _
'"       IF(aktlei IN ('97580','97582') AND lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR),CONCAT(aktlei,' streichen'),'-') " & vbCrLf & _
'"     ELSE " & vbCrLf & _
'"       CASE WHEN Epath OR nicd1<>'' OR nicd2<>'' OR sicd<>'' THEN " & vbCrLf & _
'"         CASE " & vbCrLf & _
'"           WHEN aktlei='97581' THEN '-' " & vbCrLf & _
'"           WHEN aktlei='' THEN '97581 dazu' " & vbCrLf & _
'"           ELSE CONCAT('97581 statt ',aktlei) " & vbCrLf & _
'"           END " & vbCrLf & _
'"       ELSE " & vbCrLf & _
'"         CASE " & vbCrLf & _
'"           WHEN lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) OR eintr='-' THEN IF(aktlei='','-',CONCAT(aktlei,' streichen')) " & vbCrLf & _
'"         ELSE " & vbCrLf & _
'"           CASE " & vbCrLf & _
'"             WHEN aktlei='97580' OR Eintr='-' THEN '-' " & vbCrLf & _
'"             WHEN aktlei='' THEN '97580 dazu' " & vbCrLf & _
'"             ELSE CONCAT('97580 statt ',aktlei) " & vbCrLf & _
'"           END " & vbCrLf & _
'"         END " & vbCrLf & _
'"       END " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'"   END " & vbCrLf & _
'" ELSE " & vbCrLf & _
'"    CASE WHEN vorlzl>=2 THEN CONCAT(aktlei,' streichen') " & vbCrLf & _
'"    ELSE " & vbCrLf & _
'"      CASE " & vbCrLf & _
'"       WHEN aktlei='' THEN '97582 dazu' " & vbCrLf & _
'"       WHEN aktlei<>'97582' THEN CONCAT('97582 statt ',aktlei) " & vbCrLf & _
'"       ELSE '-' " & vbCrLf & _
'"      END " & vbCrLf & _
'"    END " & vbCrLf & _
'" END LeiFehler " & vbCrLf & _
'",CASE " & vbCrLf & _
'"   WHEN lztlei>'97580' OR aktlei>'97580' OR Epath OR nicd1<>'' OR nicd2<>'' OR sicd<>'' THEN " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[^57]',CONCAT(LEFT(dicd,4),IF(MID(dicd,5,1)='9',CONCAT('5',MID(dicd,6)),IF(obf,'73','75')),' statt ',dicd),'-') " & vbCrLf & _
'"   ELSE " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[5]',CONCAT(LEFT(dicd,4),'9',MID(dicd,6),' statt ',dicd),'-') " & vbCrLf & _
'" END DicdFehler " & vbCrLf & _
'", IF((lztlei>'97580' OR aktlei>'97580' OR Epath OR nicd2<>'' OR NOT sicd='') AND nicd1='','I79.2 dazu','-') Nicd1Fehler " & vbCrLf & _
'", IF((lztlei>'97580' OR aktlei>'97580' OR Epath OR nicd1<>'' OR NOT sicd='') AND nicd2='','I70.2x dazu','-') Nicd2Fehler " & vbCrLf & _
'", IF(lztlei>'97580' OR aktlei>'97580' OR Epath OR nicd1<>'' OR nicd2<>'',1,0) OR sicd<>'' obpath " & vbCrLf & _
'", DATE(IF(einh IS NULL," & qtEnd(FristS) & ",ezp)) LeiDat, LANRid " & vbCrLf & _
'"FROM ( " & vbCrLf & _
'"SELECT f.pat_id, n.dakab " & vbCrLf & _
'", COALESCE(DATE(lztl.zeitpunkt),'') lztleitag,COALESCE(LEFT(lztl.leistung,5),'') lztlei " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'", COALESCE(DATE(aktl.zeitpunkt),'') aktleitag,COALESCE(LEFT(aktl.leistung,5),'') aktlei " & vbCrLf & _
'", (SELECT COUNT(0) FROM leistungen WHERE pat_id=f.pat_id AND leistung IN ('97580','97581','97582') AND YEAR(zeitpunkt)=YEAR(" & qtAnf(FristS) & ") AND zeitpunkt<" & qtAnf(FristS) & ") vorlzl " & vbCrLf & _
'", IF(e.inhalt IS NULL,'-','ja') Eintr " & vbCrLf & _
'", (TRIM(MID(inhalt,instr(inhalt,'zwingen?')+8)) LIKE 'j%' AND NOT TRIM(MID(inhalt,instr(inhalt,'genauso auf?')+12)) LIKE 'j%' " & vbCrLf & _
'"   AND TRIM(MID(inhalt,instr(inhalt,'genauso?')+8)) LIKE 'j%') OR " & vbCrLf & _
'"   (CONVERT(MID(inhalt,instr(inhalt,'Rötung re:')+10),integer)>5)+ " & vbCrLf & _
'"   (CONVERT(MID(inhalt,instr(inhalt,'li:')+3),integer)>5)+ " & vbCrLf & _
'"   (CONVERT(MID(inhalt,instr(inhalt,'Venenfüllung: re:')+17),integer)>10)+ " & vbCrLf & _
'"   (CONVERT(MID(inhalt,instr(inhalt,', li:')+6),integer)>10) >1 EPath, " & vbCrLf & _
'"  COALESCE(dd.icd,'') Dicd,COALESCE(nd1.icd,'') Nicd1,COALESCE(nd2.icd,'') Nicd2,COALESCE(sd.icd,'') sicd " & vbCrLf & _
'", IF((SELECT 1 FROM diagview WHERE pat_id=f.pat_id AND gicd RLIKE '^L89' LIMIT 1) IS NULL,1,0) obf " & vbCrLf & _
'", e.inhalt einh, e.zeitpunkt ezp, f.LANRid " & vbCrLf & _
'"FROM aktfvs f " & vbCrLf & _
'"LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
'"LEFT JOIN leistungen lztl ON lztl.pat_id=f.pat_id AND lztl.leistung IN ('97580','97581','97582') AND " & vbCrLf & _
'"   lztl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97580','97581','97582') AND lm.zeitpunkt<" & qtAnf(FristS) & ") " & vbCrLf & _
'"LEFT JOIN leistungen aktl ON aktl.pat_id=f.pat_id AND aktl.leistung IN ('97580','97581','97582') AND " & vbCrLf & _
'"   aktl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97580','97581','97582') AND lm.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
'"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art='dakap' AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
'"LEFT JOIN diagview dd ON dd.pat_id=f.pat_id AND dd.gicd RLIKE '^E1[0-4]\.' AND dd.obdauer<>0 " & vbCrLf & _
'"LEFT JOIN diagview nd1 ON nd1.pat_id=f.pat_id AND nd1.gicd IN ('I79.2') AND nd1.obdauer<>0 " & vbCrLf & _
'"LEFT JOIN diagview nd2 ON nd2.pat_id=f.pat_id AND nd2.gicd LIKE 'I70.2%' AND nd2.obdauer<>0 " & vbCrLf & _
'"LEFT JOIN diagview sd ON sd.pat_id=f.pat_id AND sd.gicd = 'I73.9' AND sd.obdauer<>0 " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'"WHERE n.dakab>20000101 AND (dakab<=DATE(lztl.zeitpunkt) OR ISNULL(lztl.zeitpunkt))) i " & vbCrLf & _
'") i " & vbCrLf & _
'"WHERE leifehler<>'-' OR dicdfehler<>'-' OR nicd1fehler<>'-' OR nicd2fehler<>'-' " & vbCrLf & _
'"GROUP BY PID,leifehler,dicdfehler,nicd1fehler,nicd2fehler " & vbCrLf & _
'";"
' mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
'
'' 126
' AwN(AWlf) = "Fehler bei DAK-Modul: Hepatopathie (lauto)"
' ' AND COALESCE(Dggel,0)=0
'sql(AWlf) = vbCrLf & _
'"SELECT * FROM ( " & vbCrLf & _
'" SELECT i.Pat_ID PID,i.DAKab,lztLeiTag,lztLei,aktLeiTag,aktLei,vorlZl,Dicd,Nicd1 " & vbCrLf & _
'" ,CASE WHEN lztlei='' OR lztlei='97590' THEN " & vbCrLf & _
'"   CASE " & vbCrLf & _
'"     WHEN eintr='' AND nicd1<>'' THEN " & vbCrLf & _
'"       IF(aktlei IN ('97590','97592') AND lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR),CONCAT(aktlei,' streichen'),'-') " & vbCrLf & _
'"     ELSE " & vbCrLf & _
'"       CASE WHEN Epath OR nicd1<>'' OR sicd<>'' THEN " & vbCrLf & _
'"         CASE " & vbCrLf & _
'"           WHEN aktlei='97591' THEN '-' " & vbCrLf & _
'"           WHEN aktlei='' THEN '97591 dazu' " & vbCrLf & _
'"           ELSE CONCAT('97591 statt ',aktlei) " & vbCrLf & _
'"           END " & vbCrLf & _
'"       ELSE " & vbCrLf & _
'"         CASE " & vbCrLf & _
'"           WHEN lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) OR eintr='' THEN IF(aktlei='','-',CONCAT(aktlei,' streichen')) " & vbCrLf & _
'"         ELSE " & vbCrLf & _
'"           CASE " & vbCrLf & _
'"             WHEN aktlei='97590' OR Eintr='' THEN '-' " & vbCrLf & _
'"             WHEN aktlei='' THEN '97590 dazu' " & vbCrLf & _
'"             ELSE CONCAT('97590 statt ',aktlei) " & vbCrLf & _
'"           END " & vbCrLf & _
'"         END " & vbCrLf & _
'"       END " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'"   END " & vbCrLf & _
'" ELSE " & vbCrLf & _
'"    CASE WHEN vorlzl>=2 THEN CONCAT(aktlei,' streichen') " & vbCrLf & _
'"    ELSE " & vbCrLf & _
'"      CASE " & vbCrLf & _
'"       WHEN aktlei='' THEN '97592 dazu' " & vbCrLf & _
'"       WHEN aktlei<>'97592' THEN CONCAT('97592 statt ',aktlei) " & vbCrLf & _
'"       ELSE '-' " & vbCrLf & _
'"      END " & vbCrLf & _
'"    END " & vbCrLf & _
'" END LeiFehler " & vbCrLf & _
'",CASE " & vbCrLf & _
'"   WHEN lztlei>'97590' OR aktlei>'97590' OR Epath OR nicd1<>'' OR sicd<>'' THEN " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[^67]',CONCAT(LEFT(dicd,4),IF(MID(dicd,5,1)='9',CONCAT('6',MID(dicd,6)),IF(obf,'73','75')),' statt ',dicd),'-') " & vbCrLf & _
'"   ELSE " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[6]',CONCAT(LEFT(dicd,4),'9',MID(dicd,6),' statt ',dicd),'-') " & vbCrLf & _
'" END DicdFehler " & vbCrLf & _
'", IF((lztlei>'97590' OR aktlei>'97590' OR Epath OR sicd<>'') AND nicd1='','K77.8 dazu','-') Nicd1Fehler " & vbCrLf & _
'", IF(lztlei>'97590' OR aktlei>'97590' OR Epath OR nicd1<>'',1,0) OR sicd<>'' obpath " & vbCrLf & _
'" ,EPath,Eintr " & vbCrLf & _
'", DATE(IF(einh IS NULL," & qtEnd(FristS) & ",ezp)) LeiDat, LANRid " & vbCrLf & _
'"FROM ( " & vbCrLf & _
'"SELECT f.pat_id, n.dakab " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'", COALESCE(DATE(lztl.zeitpunkt),'') lztleitag,COALESCE(LEFT(lztl.leistung,5),'') lztlei " & vbCrLf & _
'", COALESCE(DATE(aktl.zeitpunkt),'') aktleitag,COALESCE(LEFT(aktl.leistung,5),'') aktlei " & vbCrLf & _
'", (SELECT COUNT(0) FROM leistungen WHERE pat_id=f.pat_id AND leistung IN ('97590','97591','97592') AND YEAR(zeitpunkt)=YEAR(" & qtAnf(FristS) & ") AND zeitpunkt<" & qtAnf(FristS) & ") vorlzl " & vbCrLf & _
'", COALESCE(LEFT(e.inhalt,120),'') Eintr " & vbCrLf & _
'", COALESCE(e.inhalt RLIKE 'Leber(((?!((re|li) Niere|Nieren|Pan[ck]r)).)*?((?<!kaum ver|nicht ver)fett|(?<!leicht |kaum |nicht |keine |gering )verdicht)|.*fetteinl|zi)|Fettleber',0) EPath " & vbCrLf & _
'", COALESCE(dd.icd,'') Dicd,COALESCE(nd1.icd,'') Nicd1,COALESCE(sd.icd,'') sicd " & vbCrLf & _
'", IF((SELECT 1 FROM diagview WHERE pat_id=f.pat_id AND gicd RLIKE '^L89' LIMIT 1) IS NULL,1,0) obf " & vbCrLf & _
'", e.inhalt einh, e.zeitpunkt ezp, f.LANRid " & vbCrLf & _
'"FROM aktfvs f " & vbCrLf & _
'"LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
'"LEFT JOIN leistungen lztl ON lztl.pat_id=f.pat_id AND lztl.leistung IN ('97590','97591','97592') AND " & vbCrLf & _
'"   lztl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97590','97591','97592') AND lm.zeitpunkt<" & qtAnf(FristS) & ") " & vbCrLf & _
'"LEFT JOIN leistungen aktl ON aktl.pat_id=f.pat_id AND aktl.leistung IN ('97590','97591','97592') AND " & vbCrLf & _
'"   aktl.zeitpunkt=(SELECT MAX(zeitpunkt) FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97590','97591','97592') AND lm.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
'"LEFT JOIN eintraege e ON e.pat_id=f.pat_id AND e.art='sono' AND e.inhalt LIKE 'Abdomen:%' AND e.zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " " & vbCrLf & _
'"LEFT JOIN diagview dd ON dd.pat_id=f.pat_id AND dd.gicd RLIKE '^E1[0-4]\.' AND dd.obdauer<>0 " & vbCrLf & _
'"LEFT JOIN diagview nd1 ON nd1.pat_id=f.pat_id AND nd1.gicd IN ('K77.8') AND nd1.obdauer<>0 " & vbCrLf & _
'"LEFT JOIN diagview sd ON sd.pat_id=f.pat_id AND sd.gicd RLIKE '^K76\.[09]|^K71' AND sd.obdauer<>0 " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'"WHERE n.dakab>20000101 AND (dakab<=DATE(lztl.zeitpunkt) OR ISNULL(lztl.zeitpunkt))) i " & vbCrLf & _
'") i " & vbCrLf & _
'"WHERE leifehler<>'-' OR dicdfehler<>'-' OR nicd1fehler<>'-' " & vbCrLf & _
'"GROUP BY PID,leifehler,dicdfehler,nicd1fehler " & vbCrLf & _
'";"
' mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
'
'' 127
'' bei fehlender Albuminurie und bisheriger Klassifikation als nicht pathologisch => Toleranzbreite bis eGFR 45 ml/min
' AwN(AWlf) = "Fehler bei DAK-Modul: Nephropathie (lauto)"
'' AND COALESCE(Dggel,0)=0
'sql(AWlf) = _
'"SELECT * FROM ( " & vbCrLf & _
'"  SELECT i.Pat_ID PID,i.DAKab,lztLeiTag,lztLei,aktLeiTag,aktLei,vorlZl,_lGFR(i.pat_id) eGFR, EPath,Dicd,Nicd " & vbCrLf & _
'"  ,CASE WHEN lztlei='' OR lztlei='97600' THEN " & vbCrLf & _
'"        CASE WHEN Epath OR nicd<>'' THEN " & vbCrLf & _
'"          CASE " & vbCrLf & _
'"            WHEN aktlei='97601' THEN '-' " & vbCrLf & _
'"            WHEN aktlei='' THEN '97601 dazu' " & vbCrLf & _
'"            ELSE CONCAT('97601 statt ',aktlei) " & vbCrLf & _
'"            END " & vbCrLf & _
'"        ELSE " & vbCrLf & _
'"          CASE " & vbCrLf & _
'"            WHEN lztleitag>=SUBDATE(" & qtAnf(FristS) & ",INTERVAL 1 YEAR) THEN IF(aktlei='','-',CONCAT(aktlei,' streichen')) " & vbCrLf & _
'"          ELSE " & vbCrLf & _
'"            CASE " & vbCrLf & _
'"              WHEN aktlei='97600' THEN '-' " & vbCrLf & _
'"              WHEN aktlei='' THEN '97600 dazu' " & vbCrLf & _
'"              ELSE CONCAT('97600 statt ',aktlei) " & vbCrLf & _
'"            END " & vbCrLf & _
'"          END " & vbCrLf & _
'"    END " & vbCrLf & _
'"  ELSE " & vbCrLf & _
'"     CASE WHEN vorlzl>=2 THEN CONCAT(aktlei,' streichen') " & vbCrLf & _
'"     ELSE " & vbCrLf & _
'"       CASE "
'sql(AWlf) = sql(AWlf) & vbCrLf & _
'"        WHEN aktlei='' THEN '97602 dazu' " & vbCrLf & _
'"        WHEN aktlei<>'97602' THEN CONCAT('97602 statt ',aktlei) " & vbCrLf & _
'"        ELSE '-' " & vbCrLf & _
'"       END " & vbCrLf & _
'"     END " & vbCrLf & _
'"  END LeiFehler " & vbCrLf & _
'" ,CASE " & vbCrLf & _
'"    WHEN lztlei>'97600' OR aktlei>'97600' OR Epath OR NOT nicd='' THEN " & vbCrLf & _
'"     IF(dicd RLIKE '^E1.\.[^27]',CONCAT(LEFT(dicd,4),IF(MID(dicd,5,1)='9',CONCAT('2',MID(dicd,6)),IF(obf,'73','75')),' statt ',dicd),'-') " & vbCrLf & _
'"    ELSE " & vbCrLf & _
'"      IF(dicd RLIKE '^E1.\.[2]',CONCAT(LEFT(dicd,4),'9',MID(dicd,6),' statt ',dicd),'-') " & vbCrLf & _
'"  END DicdFehler " & vbCrLf & _
'" , IF((lztlei>'97600' OR aktlei>'97600' OR Epath) AND nicd='',CONCAT(sicd,IF(nicd='',' dazu',CONCAT('statt ',nicd))),IF(nicd<>'' AND nicd<>sicd,CONCAT(sicd,' statt ',nicd),'-')) NicdFehler " & vbCrLf & _
'" , IF(lztlei>'97600' OR aktlei>'97600' OR Epath OR NOT nicd='',1,0) obpath " & vbCrLf & _
'", DATE(" & qtEnd(FristS) & ") LeiDat, LANRid " & vbCrLf & _
'"FROM ( " & vbCrLf & _
'" SELECT f.pat_id, n.dakab " & vbCrLf & _
'" , COALESCE(DATE(lztl.zeitpunkt),'') lztleitag,COALESCE(LEFT(lztl.leistung,5),'') lztlei " & vbCrLf & _
'" , COALESCE(DATE(aktl.zeitpunkt),'') aktleitag,COALESCE(LEFT(aktl.leistung,5),'') aktlei " & vbCrLf
'sql(AWlf) = sql(AWlf) & _
'" , (SELECT COUNT(0) " & vbCrLf & _
'"FROM leistungen WHERE pat_id=f.pat_id AND leistung IN ('97600','97601','97602') AND YEAR(zeitpunkt)=YEAR(" & qtAnf(FristS) & ") AND zeitpunkt<" & qtAnf(FristS) & ") vorlzl " & vbCrLf & _
'", obnephrop(f.pat_id) EPath " & vbCrLf & _
'", COALESCE(dd.icd,'') Dicd,COALESCE(nd.icd,'') Nicd,CASE WHEN _lGFR(f.pat_id)>60 THEN 'N18.2' WHEN _lGFR(f.pat_id)>30 THEN 'N18.3' WHEN _lGFR(f.pat_id)>15 THEN 'N18.4' ELSE 'N18.5' END sICD " & vbCrLf & _
'", IF((SELECT 1 FROM diagnosen WHERE pat_id=f.pat_id AND icd RLIKE '^L89' LIMIT 1) IS NULL,1,0) obf " & vbCrLf & _
'", f.LANRid " & vbCrLf & _
'"FROM aktfvs f " & vbCrLf & _
'" LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
'" LEFT JOIN leistungen lztl ON lztl.pat_id=f.pat_id AND lztl.leistung IN ('97600','97601','97602') AND " & vbCrLf & _
'"    lztl.zeitpunkt=(SELECT MAX(zeitpunkt) " & vbCrLf & _
'"FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97600','97601','97602') AND lm.zeitpunkt<" & qtAnf(FristS) & ") " & vbCrLf & _
'" LEFT JOIN leistungen aktl ON aktl.pat_id=f.pat_id AND aktl.leistung IN ('97600','97601','97602') AND " & vbCrLf & _
'"    aktl.zeitpunkt=(SELECT MAX(zeitpunkt) " & vbCrLf & _
'"FROM leistungen lm WHERE lm.pat_id=f.pat_id AND lm.leistung IN ('97600','97601','97602') AND lm.zeitpunkt>=" & qtAnf(FristS) & ") " & vbCrLf & _
'" LEFT JOIN diagview dd ON dd.pat_id=f.pat_id AND dd.gicd RLIKE '^E1[0-4]\.' AND dd.obdauer<>0 " & vbCrLf & _
'" LEFT JOIN diagview nd ON nd.pat_id=f.pat_id AND nd.icd RLIKE '^N1[89]' AND NOT nd.gicd RLIKE '^N18\.[1]' AND nd.obdauer<>0 " & vbCrLf & _
'" WHERE n.dakab>20000101 AND (dakab<=DATE(lztl.zeitpunkt) OR ISNULL(lztl.zeitpunkt))) i " & vbCrLf & _
'" ) i " & vbCrLf & _
'" WHERE leifehler<>'-' OR dicdfehler<>'-' OR nicdfehler<>'-' " & vbCrLf & _
'" GROUP BY PID,leifehler,dicdfehler,nicdfehler " & vbCrLf & _
'";"
' mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
'
'' 128
'AwN(AWlf) = "abschließende Diabetesdiagnosenkorrektur bei DAK-Modul"
'sql(AWlf) = _
'"SELECT pid,gesname(pid) Name, Zp,icd Ist,Soll,ni Niere,au Augen,ne Nerven,an Gefäße,so Sonst,fu Fuß FROM (" & vbCrLf & _
'" SELECT pid,Zp,icd,CONCAT(LEFT(icd,4)," & vbCrLf & _
'"  case ni+au+ne+an+so+fu when 0 THEN '9' WHEN 1 THEN IF(ni,'2',IF(au,'3',IF(ne,'4',IF(an,'5',IF(so OR fu,'6','8'))))) ELSE '7' end," & vbCrLf & _
'"  IF((MID(icd,5,1)=7 AND ni+au+ne+an+so+fu>1)OR(MID(icd,5,1)<>7 AND ni+au+ne+an+so+fu<=1),IF(MID(icd,6,1)<=1,MID(icd,6,1),IF(fu,'5','3')),IF(ni+au+ne+an+so+fu<=1,'1',IF(fu,'5','3')))) Soll,ni,au,ne,an,so,fu FROM (" & vbCrLf & _
'"   SELECT f.pat_id pid,dm.diagdatum Zp,dm.icd,IF(ISNULL(ni.icd),0,1)NI,IF(ISNULL(au.icd),0,1)AU,IF(ISNULL(ne.icd),0,1)NE,IF(ISNULL(an.icd),0,1)AN,IF(ISNULL(so.icd),0,1)SO,IF(ISNULL(fu.icd),0,1)FU " & vbCrLf & _
'"    FROM aktfv f " & vbCrLf & _
'"    LEFT JOIN namen n ON n.pat_id=f.pat_id " & vbCrLf & _
'"    LEFT JOIN diagview dm ON dm.pat_id=f.pat_id AND dm.gicd RLIKE '^E1[0-4]\.'                                   AND dm.obdauer<>0 " & vbCrLf & _
'"    LEFT JOIN diagview ni ON ni.pat_id=f.pat_id AND ni.gicd='N08.3'                                              AND ni.obdauer<>0 " & vbCrLf & _
'"    LEFT JOIN diagview au ON au.pat_id=f.pat_id AND au.gicd IN ('H28.0','H36.0')                                 AND au.obdauer<>0 " & vbCrLf & _
'"    LEFT JOIN diagview ne ON ne.pat_id=f.pat_id AND ne.gicd IN ('G73.0','G99.0','G59.0','G63.2','N31.1','N31.2') AND ne.obdauer<>0 " & vbCrLf & _
'"    LEFT JOIN diagview an ON an.pat_id=f.pat_id AND an.gicd RLIKE '^I79.2|^I25|^I21|^I7[049]^I6[35689]'          AND an.obdauer<>0 " & vbCrLf & _
'"    LEFT JOIN diagview so ON so.pat_id=f.pat_id AND so.gicd RLIKE '^M14|K77.8|^K76\.[09]|^K71'                   AND so.obdauer<>0 " & vbCrLf & _
'"    LEFT JOIN diagview fu ON fu.pat_id=f.pat_id AND fu.gicd RLIKE '^L89'                                         " & vbCrLf & _
'"   WHERE NOT ISNULL(dm.icd) " & vbCrLf & _
'"      AND n.dakab>20000101 " & vbCrLf & _
'"   GROUP BY dm.id1) i " & vbCrLf & _
'"  ) i " & vbCrLf & _
'"WHERE icd<>Soll ORDER BY pid;"
' mins(AWlf) = 10
' maxs(AWlf) = 80
' AWlf = AWlf + 1
'
'#End If ' false
'
' ' 135
' AwN(AWlf) = "Unbekannte Impfungen "
#End If

End If ' Private / Kassenpatienten
' SELECT e.pat_id, e.zeitpunkt, e.inhalt FROM eintraege e
' WHERE art='sono' AND inhalt NOT RLIKE '^Schilddrüse|^Abdomen|^Bauchaorta|^SD|^Pleura|^Thorax|^Weichteile'
' ORDER BY (SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=e.pat_id AND art='sono' AND inhalt NOT RLIKE '^Schilddrüse|^Abdomen|^Bauchaorta|^SD|^Pleura|^Thorax|^Weichteile') DESC, zeitpunkt DESC;
 

' "SELECT dok.pid, dok.inhalt, dok.zeitpunkt FROM (SELECT pat_id AS pid, zeitpunkt, inhalt FROM `eintraege` WHERE inhalt LIKE ""%ebrid%"" OR inhalt LIKE ""%resekt%"") AS dok " & vbCrLf & _
 "LEFT JOIN (SELECT * FROM `leistungen` WHERE leistung = ""97272"" AND zeitpunkt >= " & QAnfS & " AND zeitpunkt < " & qendS & ") AS l1 ON dok.pid = l1.pat_id " & vbCrLf & _
 " WHERE dok.zeitpunkt>= " & QAnfS & " AND dok.zeitpunkt < " & qendS & " AND (ISNULL(l1.leistung)) ORDER BY dok.pid, dok.zeitpunkt"
 AWz = AWlf
 ReDim Preserve sql(AWz - 1)
 ReDim Preserve AwN(AWz - 1)
 ReDim Preserve mins(AWz - 1)
 ReDim Preserve maxs(AWz - 1)
 With MFG
  .cols = 6
  .Rows = AWz + 1
  .col = 1
  .FormatString = "|*|<Erklärung"
  Dim Zahl As Variant, Länge&
   Zahl = fWertLesen(HCU, RegWurzel & App.EXEName, "Wert", Länge)
  For AWlf = 1 To AWz
   If sql(AWlf - 1) = "ü" Then
    .Row = AWlf
    .col = 2
    .CellFontBold = True
    .col = 1
    .CellBackColor = vbBlack
   End If
   .TextMatrix(AWlf, 2) = IIf(sql(AWlf - 1) = "ü", "", (AWlf - 1) & ". ") & AwN(AWlf - 1)
   If mitSQL <> 0 Then
    .TextMatrix(AWlf, 3) = dowr(sql(AWlf - 1))
    .TextMatrix(AWlf, 4) = mins(AWlf - 1)
    .TextMatrix(AWlf, 5) = maxs(AWlf - 1)
   End If
   Dim zl%
   zl = 0
   On Error Resume Next
   zl = Zahl(AWlf - 1)
   On Error GoTo fehler
   If sql(AWlf - 1) = "ü" Then
   ElseIf zl <> 0 Then
    .TextMatrix(AWlf, 1) = "X"
    .Row = AWlf
    .CellBackColor = vbYellow
   End If
'   .Row = i
'   .Col = 2
'   .text = AwN(i - 1)
'   .Col = 3
'   .text = dowr(sql(i - 1))
  Next AWlf
  .Row = 1
  .col = 1
  altFarbe = vbWhite
 End With
 sqlgezeigt = True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZeigSQL/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function      ' ZeigSQL

' in mitSQL_Click, tuStart_Click, ZeigSQL
Public Function dowr$(ByVal txt$)
 Dim pos&, zwi$
 Do While txt <> vNS
  If dowr <> vNS Then dowr = dowr & vbLf
  pos = InStr(130, txt, " ")
  If pos = 0 Then
   dowr = dowr & txt
   txt = vNS
  Else
   dowr = dowr & Left$(txt, pos - 1)
   txt = Mid$(txt, pos + 1)
  End If
 Loop
End Function ' dowr

' Form_Resize
Public Sub ObenAusricht()
 On Error GoTo fehler
  With MFG
    If .Row - altr > (.Height / .CellHeight * 0.9) Then
     If .Row < 2 Then
      .TopRow = 1
     Else
      .TopRow = .Row - 1
     End If
    End If
  End With
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ObenAusricht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' ObenAusricht

' in Form_Load
Private Sub Form_Resize()
' Me.Hintergrund.Top = 0
' Me.Hintergrund.Height = Me.Height - 700
 MFG.Height = Me.Height - MFG.Top - 400
 MFG.Width = Me.Width - MFG.Left - 400
 Call Me.ObenAusricht
End Sub ' Private Sub Form_Resize()

' in Form_Unload, tuStart_click
Public Sub EinstSpeichern()
 Dim Zahl() As Byte
 ReDim Zahl(AWz)
 Call fBiSpei(HCU, RegWurzel & App.EXEName, "mitSQL", CByte(mitSQL))
 Call fBiSpei(HCU, RegWurzel & App.EXEName, "aktfDirekt", CByte(aktfDirekt))
 If Not nichtspeichern Then
  For AWlf = 1 To AWz
   Zahl(AWlf - 1) = (MFG.TextMatrix(AWlf, 1) = "X")
  Next AWlf
  Call fBiSpei(HCU, RegWurzel & App.EXEName, "Wert", Zahl)
 End If ' not nichtspeichern then
End Sub ' Public Sub EinstSpeichern()

' tuStart_click
Public Sub SQLvorZeigSQL()
  For AWlf = 1 To AWz
   If (MFG.TextMatrix(AWlf, 1) = "X" And sql(AWlf - 1) <> "-") Then
    If InStr(MFG.TextMatrix(AWlf, 2), "Nicht zugeordnete DAK-Faxe /KKH-Faxe") <> 0 Then
     myEFrag ("UPDATE faxeinp.outa o SET pid = (" & vbCrLf & _
      "SELECT GROUP_CONCAT(n.pat_id) FROM namen n " & vbCrLf & _
      "LEFT JOIN faelle f ON f.pat_id = n.pat_id " & vbCrLf & _
      "WHERE o.docname LIKE CONCAT('%',n.nachname,'%',n.vorname,'%') " & vbCrLf & _
      "AND f.bhfb=qbeg(o.submt) AND (f.kasse LIKE '%DAK%' OR f.kasse LIKE '%KKH%' OR f.kasse LIKE '%Kaufmännische K%')" & vbCrLf & _
      ") WHERE pid=0 AND (docname LIKE '%6971042276004%' OR docname LIKE '%51180684684%') AND NOT ISNULL(" & vbCrLf & _
      "(SELECT GROUP_CONCAT(n.pat_id) FROM namen n " & vbCrLf & _
      "LEFT JOIN faelle f ON f.pat_id = n.pat_id " & vbCrLf & _
      "WHERE o.docname LIKE CONCAT('%',n.nachname,'%',n.vorname,'%') " & vbCrLf & _
      "AND f.bhfb=qbeg(o.submt) AND (f.kasse LIKE '%DAK%' OR f.kasse LIKE '%KKH%' OR f.kasse LIKE '%Kaufmännische K%')))")
     myEFrag ("UPDATE faxeinp.outa SET pid = (SELECT MAX(n.pat_id) FROM namen n " & vbCrLf & _
     "LEFT JOIN faelle f ON f.pat_id = n.pat_id AND fanf = (SELECT MAX(fanf) FROM faelle WHERE pat_id = n.pat_id) " & vbCrLf & _
     "LEFT JOIN kassenliste k ON k.id=f.kid" & vbCrLf & _
     "WHERE REPLACE(REPLACE(docname,' ',''),'-','') LIKE CONCAT(REPLACE(REPLACE(CONCAT(n.nachname,n.vorname),' ',''),'-',''),'%') " & _
     "and (k.name LIKE 'DAK%' OR k.name LIKE 'KKH%' OR k.name LIKE '%Kaufmännische K%')" & _
     ") WHERE docname LIKE '%971042276004%' AND pid=0;")
    ElseIf MFG.TextMatrix(AWlf, 2) Like "* Widersprüchlicher Diabetestyp *" Then
' ältere Updatess
     myEFrag ("UPDATE anamnesebogen a" & vbCrLf & _
      "INNER JOIN eintraege e ON e.pat_id = a.pat_id AND e.art LIKE 'andm%'" & vbCrLf & _
      "SET diabetestyp=LEFT(TRIM(MID(e.inhalt,INSTR(e.inhalt,'Diabetes Typ ')+12,20)),LENGTH(diabetestyp))" & vbCrLf & _
      "WHERE NOT ISNULL(e.inhalt) AND NOT LEFT(TRIM(MID(e.inhalt,INSTR(e.inhalt,'Diabetes Typ ')+12,20)),LENGTH(diabetestyp))=diabetestyp AND diabetestyp='?';")
     myEFrag ("UPDATE anamnesebogen a" & vbCrLf & _
      "INNER JOIN dtypen d ON a.pat_id=d.pat_id" & vbCrLf & _
      "SET Diabetestyp = d.ttyp" & vbCrLf & _
      "WHERE diabetestyp='?' AND d.ttyp<>'-';")
' neuere Updates
     myEFrag (" UPDATE anamnesebogen a" & vbCrLf & _
      "LEFT JOIN (SELECT pat_id,min(gicd) gicd FROM diagview WHERE gicd RLIKE '^E1[01]' GROUP BY pat_id) d ON a.pat_id=d.pat_id" & vbCrLf & _
      "SET diabetestyp=CASE WHEN gicd LIKE 'E10%' THEN 1 WHEN gicd LIKE 'E11%' THEN 2 WHEN icd='O24.4' THEN 'g' WHEN gicd LIKE 'R73%' THEN 'p' ELSE diabetestyp END" & vbCrLf & _
      "WHERE diabetestyp = '?' AND NOT ISNULL(gicd);")
     myEFrag (" UPDATE anamnesebogen a" & vbCrLf & _
      "LEFT JOIN (SELECT pat_id,min(icd) gicd FROM diagview WHERE (icd='O24.4' AND Dggel=0 AND diagsicherheit IN ('G',' ')) GROUP BY pat_id) gd ON a.pat_id=gd.pat_id" & vbCrLf & _
      "LEFT JOIN (SELECT pat_id,min(gicd) gicd FROM diagview WHERE gicd ='^E1[01]' GROUP BY pat_id) d ON a.pat_id=d.pat_id" & vbCrLf & _
      "SET diabetestyp='g'" & vbCrLf & _
      "WHERE diabetestyp = '?' AND NOT ISNULL(gd.gicd) AND ISNULL(d.gicd)")
     myEFrag (" UPDATE anamnesebogen a" & vbCrLf & _
      "LEFT JOIN (SELECT pat_id,min(gicd) gicd FROM diagview where gicd ='R73.0' GROUP BY pat_id) pd ON a.pat_id=pd.pat_id" & vbCrLf & _
      "LEFT JOIN (SELECT pat_id,min(icd) gicd FROM diagview where (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ')) GROUP BY pat_id) gd ON a.pat_id=gd.pat_id" & vbCrLf & _
      "LEFT JOIN (SELECT pat_id,min(gicd) gicd FROM diagview where gicd ='^E1[01]' GROUP BY pat_id) d ON a.pat_id=d.pat_id" & vbCrLf & _
      "SET diabetestyp='p'" & vbCrLf & _
      "WHERE diabetestyp = '?' and not isnull(pd.gicd) AND ISNULL(gd.gicd) AND ISNULL(d.gicd)")
     myEFrag (" UPDATE anamnesebogen a" & vbCrLf & _
      "LEFT JOIN (select pat_id, min(art) art, min(inhalt) inhalt FROM eintraege WHERE art LIKE 'andm%' GROUP BY pat_id) e ON e.pat_id=a.pat_id" & vbCrLf & _
      "LEFT JOIN (SELECT pat_id,MIN(gicd) gicd FROM diagview d WHERE (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' '))) GROUP BY pat_id) gd ON a.pat_id=gd.pat_id" & vbCrLf & _
      "SET diabetestyp=LEFT(TRIM(MID(e.inhalt,INSTR(e.inhalt,'Diabetes Typ ')+12,20)),1)" & vbCrLf & _
      "WHERE LEFT(Diabetestyp,10)<> 's' and LEFT(Diabetestyp,10)<>case when MID(gicd,3,1)='0' then '1' when MID(gicd,3,1)='1' then '2' END" & vbCrLf & _
      "AND TRIM(MID(e.inhalt,INSTR(e.inhalt,'Diabetes Typ ')+12,20)) RLIKE '^[12] ' AND ((LEFT(TRIM(MID(e.inhalt,INSTR(e.inhalt,'Diabetes Typ ')+12,20)),1)='1' AND gicd LIKE 'E10%') OR (LEFT(TRIM(MID(e.inhalt,INSTR(e.inhalt,'Diabetes Typ ')+12,20)),1)='2' AND gicd LIKE 'E11%'))")
     myEFrag (" UPDATE anamnesebogen a" & vbCrLf & _
      "LEFT JOIN diagview d ON d.pat_id=a.pat_id (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ')))" & vbCrLf & _
      "SET diabetestyp=CASE WHEN gicd LIKE 'E10%' THEN 1 WHEN gicd LIKE 'E11%' THEN 2 WHEN icd='O24.4' THEN 'g' WHEN gicd LIKE 'R73%' THEN 'p' ELSE diabetestyp END" & vbCrLf & _
      "WHERE ISNULL((SELECT 0 FROM eintraege WHERE pat_id=a.pat_id AND art LIKE 'andm%' LIMIT 1))")
    End If ' InStr(MFG.TextMatrix(AWlf, 2), "Nicht zugeordnete DAK-Faxe /KKH-Faxe") <> 0 Then
   End If ' (MFG.TextMatrix(AWlf, 1) = "X") Then
  Next AWlf
End Sub ' Public Sub SQLvorZeigSQL()

Private Sub Form_Unload(Cancel As Integer)
 Call EinstSpeichern
' End
End Sub ' Private Sub Form_Unload(Cancel As Integer)

Public Sub MFG_Click()
 If noenter = 0 Then
  If True Or fgespei = 0 Then
   If Me.MFG.col = 1 And Me.MFG.CellBackColor <> vbBlack Then ' Zwischenüberschriften
    If Me.MFG.Text = "X" Then
     Me.MFG.Text = vNS
     Me.MFG.CellBackColor = altFarbe
    Else
     altFarbe = Me.MFG.CellBackColor
     Me.MFG.Text = "X"
     Me.MFG.CellBackColor = vbYellow
    End If
   End If
   fgespei = -1
  End If ' True Or fgespei = 0 Then
 End If ' noenter = 0 Then
End Sub ' Sub MFG_Click()

Private Sub mfg_entercell()
 Exit Sub
End Sub ' mfg_entercell()

Private Sub MFG_leavecell()
 If noenter = 0 Then
'  Me.MFG.CellBackColor = IIf(Me.MFG.Row = 0, vbActiveBorder, altFarbe)
'  IF MFG.CellBackColor = 0 THEN Stop
  fgespei = 0
 End If
End Sub
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Call key(KeyCode, Shift, Me, Me.ActiveControl.Name)
'End Sub

Private Sub mitSQL_Click() ' SQL in Liste
 If noenter <> 0 Then Exit Sub
 With MFG
 For AWlf = 1 To AWz
  If mitSQL <> 0 Then
   .TextMatrix(AWlf, 3) = dowr(sql(AWlf - 1))
   .TextMatrix(AWlf, 4) = mins(AWlf - 1)
   .TextMatrix(AWlf, 5) = maxs(AWlf - 1)
  Else
   .TextMatrix(AWlf, 3) = vNS
   .TextMatrix(AWlf, 4) = vNS
   .TextMatrix(AWlf, 5) = vNS
  End If
 Next AWlf
 End With
 Call SizeColumns(MFG, Me)
End Sub ' mitSQL_Click()

Private Sub mitSQL_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me, Me.ActiveControl.name)
End Sub ' mitSQL_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub mitSQLMenü_Click()
 Call mitSQL_Click
End Sub ' mitSQLMenü_Click()

Private Sub Private_Click()
 Dim i&
 If noenter Then Exit Sub
 If Me.Private <> 0 Then
  Me.Quartal.Enabled = 0
  Me.aktfDirekt.Enabled = 0
 Else
  Me.Quartal.Enabled = True
  Me.aktfDirekt.Enabled = True
 End If
 AWz = 200
 ReDim sql(AWz - 1)
 ReDim obmo(AWz - 1)
 ReDim AwN(AWz - 1)
 ReDim mins(AWz - 1)
 ReDim maxs(AWz - 1)
' For i = 0 To Me.MFG.Cols - 1
'  ON Error Resume Next
'  Me.MFG.TextMatrix(i, 1) = "y"
' Next i
 Call ZeigSQL
 Call SizeColumns(MFG, Me)
End Sub 'Private_Click()

Private Sub Quartal_Click()
 If noenter Then Exit Sub
 Call ZeigSQL
End Sub ' Quartal_click

Private Sub Quartal_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me, Me.ActiveControl.name)
End Sub ' Sub Quartal_KeyDown(KeyCode%, Shift%)

Private Sub Start_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me, Me.Start.name)
End Sub ' Start_KeyDown(KeyCode%, Shift%)

' Starte die Abrechnungsfehlerlisten
Private Sub Start_Click()
 tuStart_click (0)
End Sub ' Start_Click()

Private Sub tuStart_click(obauto%)
 Dim lfSQL$, i&, StartZeit As Date, überschrift As New CString
 Static rc As New ADODB.Connection
 Dim rLF As ADODB.Recordset, rIn As ADODB.Recordset
 StartZeit = Now()
 If rc.State <> 1 Then
   Set rc = Lese.dbv.wCn
 End If
 Set rLF = New ADODB.Recordset
 Call EinstSpeichern
 Call SQLvorZeigSQL
 Dim AbrFlrDt$, AbrAutDt$
 Screen.MousePointer = vbHourglass
 
 AbrFlrDt = tAusgSg + "_" & AktQ & "_" & Format(Now, "yyyy-mm-dd.hh.mm.ss") & ".txt"
 AbrAutDt = AbrVerz & "\Abrechnungsprotokoll_" & AktQ & "_" & Format(Now, "yyyy-mm-dd.hh.mm.ss") & ".txt"
 überschrift.AppVar Array("Abrechnungsfehler für Quartal ", AktQ, ", ODBC-Verbindung:", Lese.dbv.Constr, vbCrLf)
' Open AbrFlrDt For Output AS #359
' Print #359, "Abrechnungsfehler für Quartal " & aktQ

' IF Me.AbrF.DAO THEN
'  Print #359, "Abfragesprache DAO"
' ElseIf Me.AbrF.ADOJet THEN
'  Print #359, "Abfragesprache ADODB, Jet"
' ElseIf Me.AbrF.AdoMySQL THEN
'  Print #359, "Abfragesprache ADODB, MySQL: quelle"
' ElseIf Me.AbrF.AdoMySQL1 THEN
'  Print #359, "Abfragesprache ADODB, MySQL: quelle1"
' ElseIf Me.AbrF.AdoMySQL2 THEN
'  Print #359, "Abfragesprache ADODB, MySQL: quelle2"
' END IF
' Print #359, "ODBC-Verbindung:" & Lese.dbv.Constr

 With MFG
 Dim obappend%
 Dim angefangen%
 Dim BDT As BDTSchreib
 obappend = False
 If obauto Then Open AbrAutDt For Output As #317
 For AWlf = 1 To .Rows - 1
  If sql(AWlf - 1) = "ü" Then ' Zwischenüberschrift
  ElseIf .TextMatrix(AWlf, 1) = "X" And sql(AWlf - 1) <> "-" Then
    Do While Not AbrFausg(Str(AWlf - 1) & ". " & AwN(AWlf - 1), REPLACE$(dowr(sql(AWlf - 1)), vbLf, " "), obmo(AWlf - 1), AbrFlrDt, mins(AWlf - 1), maxs(AWlf - 1), überschrift, obappend, AWlf - 1, obauto, angefangen, BDT)
     Dim altAWlf%
     altAWlf = AWlf
     MsgBox "Stop in Start_Click" & vbCrLf & "AWlf: " & AWlf
     Stop
     Call ZeigSQL(obauto)
     AWlf = altAWlf
    Loop
    überschrift = vNS
    obappend = True
  End If
  DoEvents
 Next AWlf
 If obauto Then Close #317
 Set BDT = Nothing
 End With
 If rLF.State = 1 Then rLF.Close
 ' Close #359
 Dim erg ' Task-ID
' erg = Shell("dir """ & AbrFlrDt & """", vbMaximizedFocus)
 If FSO.FileExists(AbrFlrDt) Then erg = zeigan(AbrFlrDt)
 If obauto Then erg = zeigan(AbrAutDt)
 Screen.MousePointer = vbNormal
' Beep 440, 500
' MsgBox "Fertig mit Ermitteln der Abrechnungsfehler!" & vbCrLf & "Start: " & StartZeit & vbCrLf & "Ende: " & Now()
End Sub ' tu_Start_Click

Function QuartalFüll()
 Dim mon%, Jahr%, quart$
 On Error GoTo fehler
 noenter = True
 mon = 1
 Jahr = 2005
 Do
  Do
   quart = CStr(mon) + CStr(Jahr)
   Quartal.AddItem (quart)
   If quart = ZQuart(Now - Verspätung) Then
    Quartal.Selected(Quartal.ListCount - 1) = -1
   End If
   If quart = ZQuart(Now) Then Exit Function
   mon = mon + 1
   If mon > 4 Then Exit Do
  Loop
  mon = 1
  Jahr = Jahr + 1
 Loop
 noenter = False
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in QuartalFüll/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' QuartalFüll

'Public FUNCTION ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung
'Dim j AS String * 4, Q AS String * 1
'On Error GoTo fehler:
'j = YEAR(Datum)
'SELECT CASE Datum
' Case Is < CDate("1.4." + j): Q = "1"
' Case Is < CDate("1.7." + j): Q = "2"
' Case Is < CDate("1.10." + j): Q = "3"
' Case Else: Q = "4"
'End SELECT
'ZQuart = Q + j
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZQuart/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung






' Make the FlexGrid's columns big enough to hold all values.
Sub SizeColumns(ByVal flx As MSHFlexGrid, frm As Form)
Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer
Dim zwz%, pos&
On Error GoTo fehler
If Not sqlgezeigt Then Call ZeigSQL
FNr = 1
    max_row = flx.Rows - 1
FNr = 2
    For c = 0 To flx.cols - 1
        max_wid = 0
        For r = 0 To max_row
FNr = 3
            wid = Lese.dbv.TextWidth(REPLACE$(flx.TextMatrix(r, c), vbCrLf, "--"))
FNr = 4
            If max_wid < wid Then max_wid = wid
FNr = 5
        Next r
FNr = 6
        flx.ColWidth(c) = (max_wid + 400)
FNr = 7
    Next c
    For r = 0 To max_row
        zwz = 0
        pos = 0
        Do
FNr = 8
         pos = InStr(pos + 1, flx.TextMatrix(r, 3), vbLf)
FNr = 9
         If pos > 0 Then
          zwz = zwz + 1
         Else
          Exit Do
         End If
        Loop
FNr = 10
        flx.RowHeight(r) = zwz * (flx.RowHeight(r) * 0.85) + 250
FNr = 11
    Next r
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", Err.number = " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SizeColumns/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' SizeColumns

' aufgerufen in tuStart_Click
Public Function AbrFausg(name$, sql$, obmo%, Datei$, mins%, ByVal maxs%, überschrift As CString, obappend%, sqlnr%, obauto%, ByRef angefangen%, ByRef BDT As BDTSchreib) ' Abrechnungsfehler ausgeben
 Dim ÜberschrAkt As New CString
 Dim ErrNr&, ErrDes$
 ÜberschrAkt = überschrift
 On Error GoTo fehler
 Dim T1!
 Static rc As New ADODB.Connection
 Static rE As ADODB.Recordset
#If obmitalterform Then
 Dim rD As DAO.Recordset
 If Me.AbrF.DAO Then
  If Dtb Is Nothing Then Call dtbInit(Datei)
  Set rD = Dtb.OpenRecordset(sql)
  ausg = do_ausg(rD, mins, maxs)
 Else
#Else
 If True Then
#End If
'  If rE Is Nothing Then' auskommentiert 26.3.23
  Set rE = New ADODB.Recordset
  If obmo Then
   Call MOConInit(, "Abfrageausgeb: " & name)
   Set rc = MOCon
  Else '  If rc.State <> 1 Then
   Set rc = Lese.dbv.wCn
  End If
'  If rE.State = 1 Then rE.Close ' auskommentiert 26.3.23
'  IF (me.AbrF.AdoMySQL OR me.AbrF.AdoMySQL1 OR me.AbrF.AdoMySQL2) THEN ON Error Resume Next
'  IF (me.AbrF.AdoMySQL OR me.AbrF.AdoMySQL1 OR me.AbrF.AdoMySQL2) THEN sql = replace$(replace$(replace$(replace$(replace$(replace$(sql, "`", "`"), "`", "`"), " faelle ", " `faelle` "), "iif(", "if("), "DATE(", "date("), "int(", "date(")
'  IF (me.AbrF.AdoMySQL OR me.AbrF.AdoMySQL1 OR me.AbrF.AdoMySQL2) THEN 'concat
  If InStrB(CoSt, "MySQL") <> 0 Then
   On Error Resume Next
   sql = REPLACE$(REPLACE$(REPLACE$(REPLACE$(sql, " faelle ", " `faelle` "), "iif(", "if("), "datevalue(", "date("), "int(", "date(")
   sql = REPLACE$(sql, "||", ",")
  Else
   sql = REPLACE$(sql, "||", " & ")
   sql = REPLACE$(sql, "concat", "")
  End If
  T1 = Timer
  On Error GoTo fehler
  FNr = 999
  Dim rcsql$
  On Error Resume Next
  ÜberschrAkt.AppVar Array(name)
  If InStrB(name, "Fehler bei Chronikerpauschale") <> 0 Then 'sqlnr = 33 THEN
   myEFrag "SET @qn:=(SELECT f.quartal FROM aktfv f LEFT JOIN faelle fa USING (fid) LIMIT 1)", , rc
   myEFrag "SET @qe:=vorquart(@qn,1)", , rc
   myEFrag "SET @qz:=vorquart(@qn,2)", , rc
   myEFrag "SET @qd:=vorquart(@qn,3)", , rc
   myEFrag "SET @qv:=vorquart(@qn,4)", , rc
  End If
'  rc.Execute "SET GROUP_CONCAT_MAX_LEN = " & maxs
  Err.Clear
'  rE.Open sql, rc, adOpenStatic, adLockReadOnly ' 24.6.09: Hier ging der Kontakt zum Server verloren, evtl. Zufall
  On Error GoTo fehler
  Dim rAf&
  rc.Close
  rc.Open
  myFrag rE, sql, adOpenStatic, rc, adLockReadOnly, maxs, rAf, True, ErrNr, ErrDes
  ' Listen mit (lauto)
  If obauto Then
   Dim lanrda%, leida%, ldda%, pda%, i%, Arztnr&, Zahl&, Protdat$
   Print #317, "Trage Leistungen für Liste '" & name & "' ein: "
   For i = 0 To rE.Fields.COUNT ' - 1
   On Error Resume Next ' 11.1.23 wg. fields.count
    Select Case UCase$(rE.Fields(i).name)
     Case "LANRID": lanrda = True
     Case "LEIFEHLER": leida = True
     Case "LEIDAT": ldda = True
     Case "PID": pda = True
    End Select
   Next i
   On Error GoTo fehler
   If lanrda And leida And ldda And pda Then
    For Arztnr = 1 To 2
'     angefangen = 0
'     Protdat = ""
'     Zahl = 0
     Dim rEPid&
     Do While Not rE.EOF
      rEPid = rE!pid
      If rE!lanrid = Arztnr Then
       Dim pos%, Lei$
       pos = InStr(rE!LEIFEHLER, " dazu")
       If pos <> 0 Then
        Lei = Left$(rE!LEIFEHLER, pos - 1)
        If Not angefangen Then
         Set BDT = New BDTSchreib
         If Not BDT.Start(hVerz, "Leist", 0, Lei) Then ' Arztnr) THEN
          Exit Function
         End If
         Call BDT.ImportFolderHerricht
         Call BDT.BDTKopf
         Protdat = BDT.DMPImp
         angefangen = True
        End If
        Dim pruefdat As Date
        pruefdat = DateValue(rE!LEIDAT) + IIf(TimeValue(rE!LEIDAT) = 0, CDate("18:00"), TimeValue(rE!LEIDAT))
        pruefdat = myEFrag("SELECT naemin(" & rE!pid & ",'" & Format$(pruefdat, "YYYYmmddHHMM") & "00')").Fields(0)
        If LeistungsExport1(BDT, rEPid, Lei, pruefdat, Format$(pruefdat, "HH:MM:SS"), True, Arztnr) Then
         Print #317, " " & rE!pid & " " & Lei & " " & rE!LEIDAT & " " & Arztnr
         Zahl = Zahl + 1
        End If ' LeistungsExport1
       End If ' pos<>0 Then
      End If ' re!lanrid=arztnr
      rE.MoveNext
     Loop ' While Not rE.EOF
     Print #317, "   " & Zahl & " Leistungen in " & Protdat & " eingetragen."
     rE.MoveFirst
    Next Arztnr
    If angefangen Then
     BDT.Schreib
    End If ' angefangen Then
   End If  'lanrda And leida And ldda And pda Then
  Else ' obauto
   If ErrNr <> 0 Then
    rcsql = rc
    Set rc = Nothing
    If obmo Then Set rc = MOCon Else Set rc = Lese.dbv.wCn
    If sql = "" Then AbrFausg = True: Exit Function
    Set rE = Nothing
'    On Error GoTo fehler2
    myFrag rE, sql, adOpenStatic, rc, adLockReadOnly, , , True, ErrNumber, ErrDescr
    If ErrNumber <> 0 Then
     Open Datei For Append As #307
'     Print #307, "Error " & ErrNumber & ": " & ErrDescr
     Print #307, ÜberschrAkt.Value & vbCrLf; "SQL-Fehler:" & vbCrLf & "Error " & ErrNumber & ", " & ErrDescr & vbCrLf & "bei:" & vbCrLf & sql & vbCrLf
' ausgetauscht 11.12.22:
'   myFrag rE, "SHOW ERRORS", adOpenStatic, rc, adLockReadOnly
'   If Not rE.EOF Then
'    Print #307, rE.Fields(0) & " " & rE.Fields(1) & " " & rE.Fields(2)
'    Print #307, ÜberschrAkt.Value & vbCrLf; "SQL-Fehler:" & vbCrLf & rE.Fields(0) & " " & rE.Fields(1) & " " & rE.Fields(2) & vbCrLf & sql & vbCrLf
'   Else
'    Print #307, ErrDescr
'   End If
    Close #307
' Dim tmpsql$
' tmpsql = Environ("tmp") & "\sql.txt"
' Open tmpsql For Output AS #307
' Print #307, sql
' Close #307
' zeigan tmpsql,vbNormalFocus
    GoTo zurück
    End If
   End If ' ErrNr <> 0 Then
  
   On Error GoTo fehler
'  IF Not rE.EOF THEN
   If Me.Debug <> 0 Then
    ÜberschrAkt.AppVar Array(vbCrLf, "SQL: ", vbCrLf, REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(sql, "FROM", vbCrLf + "FROM"), "from", vbCrLf + "FROM"), " ON", vbCrLf + " ON"), "LEFT JOIN", vbCrLf + "LEFT JOIN"), "INNER JOIN", vbCrLf + "INNER JOIN"), "RIGHT JOIN", vbCrLf + "RIGHT JOIN"), "WHERE", vbCrLf + "WHERE"), "ORDER BY", vbCrLf + "ORDER BY"), "ORDER BY", vbCrLf + "ORDER BY"))
   End If
'  END IF
   FNr = 1000
   TabAusgeb rE, Me, False, , rE.EOF, , , , Datei, , rE.EOF, ÜberschrAkt.Value, "Abrechnungsfehler: " & name, obappend, True
  End If ' obauto
zurück:
  AbrFausg = True
'  Exit Function
'  ON Error GoTo fehler
  
'  IF Not rE.EOF THEN
'   Print #359, vbCrLf & vbCrLf & Name
'   IF Me.Debug THEN
'    Print #359, "SQL: " + vbCrLf + replace$(replace$(replace$(replace$(replace$(replace$(replace$(replace$(replace$(sql, "FROM", vbCrLf + "FROM"), "from", vbCrLf + "FROM"), " ON", vbCrLf + " ON"), "LEFT JOIN", vbCrLf + "LEFT JOIN"), "INNER JOIN", vbCrLf + "INNER JOIN"), "RIGHT JOIN", vbCrLf + "RIGHT JOIN"), "WHERE", vbCrLf + "WHERE"), "ORDER BY", vbCrLf + "ORDER BY"), "ORDER BY", vbCrLf + "ORDER BY")
'   END IF
'   FNr = 1000
'   AbrFausg = do_ausg(rE, mins, maxs, T1)
'  Else
'zurück:
'   AbrFausg = True
'  END IF
 End If
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrLastDllError = Err.LastDllError
 ErrDescr = Err.Description
 ErrSource = IIf(IsNull(Err.source), vNS, Err.source)
fehler2:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + ErrSource + vbCrLf + "Description: " + ErrDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in AbrF_ausg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AbrF_ausg

'Public FUNCTION do_ausg%(rE, mins%, maxs%, T1!)
' Const lMax$ = 1390
' Dim i%, j%, AuS$, defS&, DSNr%, obDAO%
' Dim FldS&(), FldLen&
' Dim T2!
' #If obmitalterform THEN
'  obDAO = Me.AbrF.DAO
' #END IF
' ON Error GoTo fehler
' ReDim FldS(rE.Fields.Count)
' Do While Not rE.EOF
'  For i = 0 To rE.Fields.Count - 1
'   IF ISNULL(rE.Fields(i)) THEN
'    FldLen = 4
'   Else
'    FldLen = Len(LTrim$(rE.Fields(i)))
'   END IF
'   IF FldLen > FldS(i) THEN
'    FldS(i) = FldLen
'   END IF
'  Next i
'  rE.Move 1
' Loop
' IF Not rE.BOF THEN rE.MoveFirst
' For i = 0 To rE.Fields.Count - 1
'  FldLen = Len(rE.Fields(i).Name)
'  IF FldLen > FldS(i) THEN FldS(i) = FldLen
'  IF FldS(i) < mins THEN FldS(i) = mins
'  IF FldS(i) > maxs THEN FldS(i) = maxs
' Next i
' AuS = "Nr.  "
' For i = 0 To rE.Fields.Count - 1
''  defS = DoDefs(rE.Fields(i), mins, maxs)
'  AuS = AuS & Left$(rE.Fields(i).Name & Space(FldS(i)), FldS(i)) & " " ' defs
' Next i
' AuS = Left$(AuS, lMax)
' Print #359, String(Len(AuS), "_")
' Print #359, AuS
' Print #359, String(Len(AuS), "_")
' AuS = vNS
' IF rE.BOF THEN do_ausg = True: Exit Function
' rE.MoveFirst
' DSNr = 0
' Do While Not rE.EOF
'  DSNr = DSNr + 1
'  AuS = Format(DSNr, "0000 ")
'  For i = 0 To rE.Fields.Count - 1
'   defS = DoDefs(rE.Fields(i), mins, maxs)
'   Dim AusgS$
'   IF ISNULL(rE.Fields(i).Value) THEN
'    AusgS = "Null"
'' dbdate = 8, adDate = 7
'' dbdouble = 7, addouble = 5
'   ElseIf rE.Fields(i).Type = 204 THEN  ' 11.7.10, sonst Fehler
'    AusgS = CStr(rE.Fields(i).Value)
'   ElseIf rE.Fields(i).Type = IIf(obDAO, 8, 7) OR rE.Fields(i).Type = 135 OR ((rE.Fields(i).Type = IIf(obDAO, 7, 5) AND rE.Fields(i).Value > 0 AND rE.Fields(i).Value < 44198) AND (InStr(rE.Fields(i).Name, "zp") > 0 OR InStr(rE.Fields(i).Name, "datum") > 0 OR InStr(rE.Fields(i).Name, "zeit") > 0)) THEN
'    IF LCase$(rE.Fields(i).Name) = "zeitpunkt" OR LCase$(rE.Fields(i).Name) = "zp" THEN
'     AusgS = Format(rE.Fields(i).Value, "dd.mm.yy hh:mm")
'    Else
'     AusgS = Format(rE.Fields(i).Value, "dd.mm.yy")
'    END IF
'   Else
'    AusgS = CStr(rE.Fields(i).Value)
'   END IF
'   AusgS = Left$(AusgS & Space$(FldS(i)), FldS(i)) ' defS
''   IF Len(AusgS) < defS THEN Stop
'   AuS = AuS + AusgS + " "
'  Next i
'  AuS = Left$(AuS, lMax)
'  Print #359, AuS
'  AuS = vNS
'  rE.Move 1
' Loop
' T2 = Timer
' Print #359, "Zeitverbrauch: " & Format(T1 / 60 / 60 / 24, "general date") & " - " & Format(T2 / 60 / 60 / 24, "general date") & " (" & ROUND(T2 - T1, 3) & ")"
' do_ausg = True
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_ausg/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume: do_ausg = False: Exit Function
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION 'do_ausg

Public Function DoDefs&(fld, mins%, maxs%)
 Dim nmMin%, fS&, obDAO%
 On Error GoTo fehler
#If obmitalterform Then
 obDAO = Forms(0).AbrF.DAO
#End If
   If obDAO Then fS = fld.size
   If fld.Type = IIf(obDAO, 8, 7) Or fld.Type = 135 Then ' dbDate
    DoDefs = 8
   ElseIf fld.Type = 12 Or (obDAO And fS = 0) Then ' dbMemo = 12
    DoDefs = maxs
   Else
    If obDAO Then
     DoDefs = fld.size
     If DoDefs = 0 Then
      On Error Resume Next
      DoDefs = fld.FieldSize
      On Error GoTo fehler
     End If
    Else
     DoDefs = fld.DefinedSize
'     IF fld.ActualSize <> 0 AND fld.ActualSize < DoDefs THEN DoDefs = fld.ActualSize
    End If
    If DoDefs > maxs Or DoDefs = -1 Then DoDefs = maxs
    nmMin = Len(fld.name)
    If nmMin > mins Then nmMin = mins
    If DoDefs < mins Then DoDefs = mins
   End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doDefs/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DoDefs

