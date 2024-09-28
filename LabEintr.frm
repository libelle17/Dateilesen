VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form LabEintr 
   Caption         =   "namen"
   ClientHeight    =   10005
   ClientLeft      =   1050
   ClientTop       =   540
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "LabEintr"
   ScaleHeight     =   10005
   ScaleWidth      =   15120
   Begin VB.CommandButton inTurbomedAnzeigen 
      Caption         =   "in &Turbomed anzeigen"
      Height          =   255
      Left            =   8520
      TabIndex        =   10
      Top             =   9360
      Width           =   2535
   End
   Begin VB.CheckBox nurHeuer 
      Caption         =   "nur &heuer + letztes Jahr"
      Height          =   255
      Left            =   11160
      TabIndex        =   9
      Top             =   9360
      Width           =   2175
   End
   Begin VB.TextBox Zeilen 
      Height          =   285
      Left            =   6840
      TabIndex        =   1
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton abhaken 
      Caption         =   "abha&ken"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   9360
      Width           =   855
   End
   Begin VB.CommandButton Photoimpact 
      Caption         =   "&Photoimpact"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   9360
      Width           =   1095
   End
   Begin VB.CommandButton zurück 
      Caption         =   "&zurück"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   9360
      Width           =   855
   End
   Begin VB.CommandButton Irfan 
      Caption         =   "&Irfan"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   9360
      Width           =   1095
   End
   Begin VB.CommandButton weiter 
      Caption         =   "&Weiter"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   9360
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Neu lesen"
      Height          =   300
      Left            =   13560
      TabIndex        =   2
      Top             =   9360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "S&chließen"
      Height          =   300
      Left            =   14760
      TabIndex        =   0
      Top             =   9360
      Width           =   1080
   End
   Begin MSAdodcLib.Adodc dpS 
      Align           =   2  'Unten ausrichten
      Height          =   330
      Left            =   0
      Top             =   9675
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   9255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   16325
      _Version        =   393216
      Cols            =   8
      AllowUserResizing=   3
      FormatString    =   "|pat_id| zeitpunkt| dokpfad| dokname| quelldatum| dokgroe| abgehakt"
      _NumberOfBands  =   1
      _Band(0).BandIndent=   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imgSort 
      Height          =   480
      Index           =   1
      Left            =   2730
      Top             =   2355
      Width           =   1200
   End
End
Attribute VB_Name = "LabEintr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MARGIN_SIZE = 60      ' in Twips
' Variablen für das Aktivieren der Spaltensortierung
Private m_iSortCol As Integer
Private m_iSortType As Integer

' Variablen für das Ziehen von Spalten
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer
Public hlese As Lese
Private DatPfad$
'Const AbgehSQL$ = "SELECT briefe.Pat_ID, briefe.Zeitpunkt, briefe.pfad, briefe.Name, DATE(briefe.DokAenD) AS Quelldt, briefe.DokGroe, CONCAT(hkv.nachname, ' ', hkv.ort) AS kv, CONCAT(hüw.nachname,' ',hüw.ort) üb, DATE(bhfb) AS Fallbg, IF(abgehakt <> 0,'X',' ') AS abgeh FROM ((`briefe` LEFT JOIN `br_abgehakt` ON `briefe`.pfad = `br_abgehakt`.pfad) LEFT JOIN `namen` ON `briefe`.pat_id = `namen`.pat_id) LEFT JOIN `lfaelle` ON `briefe`.pat_id = `lfaelle`.pat_id LEFT JOIN `hausaerzte` hkv ON `namen`.kvnr = hkv.kvnr LEFT JOIN `hausaerzte` hüw ON übwv = hüw.kvnr WHERE (((`briefe`.Name) LIKE '%Fremdlabor%')) AND (ISNULL(abgehakt) OR abgehakt = 0) GROUP BY pfad ORDER BY pat_id DESC, zeitpunkt DESC, übwv DESC"
'Const AbgehSQL$ = "SELECT `briefe`.Pat_ID, `briefe`.Zeitpunkt, `briefe`.pfad, `briefe`.Name, DATE(`briefe`.DokAenD) AS Quelldt, `briefe`.DokGroe, kvnr, übwv, DATE(bhfb) AS Fallbg, IF(abgehakt <> 0,'X',' ') AS abgeh FROM ((`briefe` LEFT JOIN `br_abgehakt` ON `briefe`.pfad = `br_abgehakt`.pfad) LEFT JOIN `namen` ON `briefe`.pat_id = `namen`.pat_id) LEFT JOIN `lfaelle` ON `briefe`.pat_id = `lfaelle`.pat_id WHERE (((`briefe`.Name) LIKE '%Fremdlabor%')) AND (ISNULL(abgehakt) OR abgehakt = 0) GROUP BY pfad ORDER BY pat_id DESC, zeitpunkt DESC, übwv DESC"
Const AbgehSQL$ = "SELECT d.Pat_ID, Zeitpunkt, d.pfad, Name, intaccdatemy(Quelldatum) Quelldt, DokGroe, kvnr, übwr, intaccdatemy(bhfb) Fallbg, IF(abgehakt <> 0,'X',' ') abgeh, ROW_NUMBER() OVER(ORDER BY intaccdatemy(Quelldatum) DESC, Name) Rang FROM ((`briefe` d LEFT JOIN `br_abgehakt` da ON d.pfad = da.dokpfad) LEFT JOIN `namen` n ON d.pat_id = n.pat_id) LEFT JOIN `lfaelle` f ON d.pat_id = f.pat_id WHERE ((Name LIKE '%Fremdlabor%')) AND ((ISNULL(abgehakt) OR abgehakt = 0) OR ungueltig =1 OR Quelldatum < '19000000') GROUP BY d.pfad ORDER BY d.pat_id DESC, zeitpunkt DESC, übwr DESC"

Private Sub abhaken_Click()
  Dim rAf&, ob%, obbof%
  Dim rAd As ADODB.Recordset
  On Error GoTo fehler
  Me.MSHFlexGrid1.col = 10
  Me.MSHFlexGrid1.Text = IIf(Me.MSHFlexGrid1.Text = " ", "X", " ")
  ob = IIf(Me.MSHFlexGrid1 = "X", 1, 0)
  Me.MSHFlexGrid1.col = 3
  myFrag rAd, "SELECT 0 FROM `br_abgehakt` WHERE dokpfad = '" & doUmwfSQL(Me.MSHFlexGrid1.Text, lies.obMySQL) & "'"
  obbof = rAd.BOF
  Set rAd = Nothing
  If obbof Then
   InsKorr DBCn, "INSERT INTO `br_abgehakt` (abgehakt,aktzeit,dokpfad) VALUES(" & ob & "," & DatFor_k(Now) & ",'" & doUmwfSQL(Me.MSHFlexGrid1.Text, lies.obMySQL) & "')", rAf
  Else ' obbof
   Call myEFrag("UPDATE `br_abgehakt` SET abgehakt = " & ob & ", aktzeit = " & DatFor_k(Now) & " WHERE dokpfad = '" & doUmwfSQL(Me.MSHFlexGrid1.Text, lies.obMySQL) & "'", rAf)
  End If ' obbof
 Call weiter_Click
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then
' Call hlese.ConstrFestleg(0, hlese)
 Call acon(quelleT)
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in abhaken/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' abhaken_Click

'Public Sub Pause(Millisekunden AS Long)
'Sleep Millisekunden
'End Sub

Private Sub inTurbomedAnzeigen_Click() ' in Turbomed anzeigen
 Const obStumm% = 0
 Dim hnd&, Pat_id& ' , dtyp%, alttop&, altr&
 Const Pat_IDSp& = 1
' Dim DT As DMPClass, j%
' Dim rTyp As New ADODB.Recordset
 On Error GoTo fehler
 With Me.MSHFlexGrid1
  .SetFocus
'  alttop = .TopRow
'  altr = .Row
'  altC = .col
  Pat_id = .TextMatrix(.Row, Pat_IDSp)
'  .col = altC
 End With
' Call DMPAusgeb0(dt, Pat_id, Not obstumm)
 hnd = FensterHandle("TurboMed")
 If hnd <> 0 Then
  Call Shell(IrfanString, vbMaximizedFocus)
  Pause (Pausenlänge)
  AppActivate "TurboMed", True
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ENTER}", True
  Pause (Pausenlänge)
  SendKeys "{F12}", True
  Pause (Pausenlänge)
'  SendKeys "{ESC}", True
'  Pause (Pausenlänge)
'  SendKeys "{ESC}", True
'  Pause (Pausenlänge)
  SendKeys "p", True
  Pause (Pausenlänge)
  SendKeys "" & "{bs}" & Pat_id & "", True
  Pause (Pausenlänge)
  SendKeys "{ENTER}", True
  Pause (Pausenlänge)
  SendKeys "{F3}", True
  Pause (Pausenlänge)
  AppActivate "TurboMed", True
'  Pause (Pausenlänge)
'  SendKeys "^{F7}", True
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in inTurbomedAnzeigen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' InTurbomed anzeigen


'Private Sub dpS_Error(ByVal ErrorNumber AS Long, Description AS String, ByVal Scode AS Long, ByVal Source AS String, ByVal HelpFile AS String, ByVal HelpContext AS Long, fCancelDisplay AS Boolean)
' Stop
'End Sub

'Private Sub dpS_FieldChangeComplete(ByVal cFields AS Long, Fields AS Variant, ByVal pError AS ADODB.Error, adStatus AS ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
' Stop
'End Sub

'Private Sub dpS_MoveComplete(ByVal adReason AS ADODB.EventReasonEnum, ByVal pError AS ADODB.Error, adStatus AS ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
' Stop
'End Sub

'Private Sub dpS_RecordChangeComplete(ByVal adReason AS ADODB.EventReasonEnum, ByVal cRecords AS Long, ByVal pError AS ADODB.Error, adStatus AS ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
' Stop
'End Sub

'Private Sub dpS_RecordsetChangeComplete(ByVal adReason AS ADODB.EventReasonEnum, ByVal pError AS ADODB.Error, adStatus AS ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
' Stop
'End Sub

'Private Sub dpS_WillChangeField(ByVal cFields AS Long, Fields AS Variant, adStatus AS ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
' Stop
'End Sub

'Private Sub dpS_WillChangeRecord(ByVal adReason AS ADODB.EventReasonEnum, ByVal cRecords AS Long, adStatus AS ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
' Stop
'End Sub

'Private Sub dpS_WillChangeRecordset(ByVal adReason AS ADODB.EventReasonEnum, adStatus AS ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
' Stop
'End Sub

'Private Sub dpS_WillMove(ByVal adReason AS ADODB.EventReasonEnum, adStatus AS ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
' Stop
'End Sub

Private Sub MSHFlexGrid1_EnterCell()
 Me.MSHFlexGrid1.CellBackColor = vbYellow
End Sub ' MSHFlexGrid1_EnterCell

Private Sub MSHFlexGrid1_LeaveCell()
 Me.MSHFlexGrid1.CellBackColor = IIf(Me.MSHFlexGrid1.Row = 0, vbActiveBorder, vbWhite)
End Sub ' MSHFlexGrid1_LeaveCell

Private Sub nurHeuer_Click()
 Call LadDaten
End Sub ' nurHeuer_Click

Private Sub weiter_Click()
 On Error GoTo fehler
 Me.MSHFlexGrid1.CellBackColor = IIf(Me.MSHFlexGrid1.Row = 0, vbActiveBorder, vbWhite)
 Me.MSHFlexGrid1.col = 3
 If Me.MSHFlexGrid1.Row < Me.MSHFlexGrid1.Rows - 1 Then
  If Me.Visible Then
   Me.MSHFlexGrid1.SetFocus
   On Error Resume Next
   Err.Clear
   SendKeys "{DOWN}"
   If Err.Number <> 0 Then
    Me.MSHFlexGrid1.Row = Me.MSHFlexGrid1.Row + 1
   End If
   On Error GoTo fehler
  Else
   Me.MSHFlexGrid1.Row = Me.MSHFlexGrid1.Row + 1
  End If
 End If
 If Me.Visible Then Call Me.MSHFlexGrid1.SetFocus
 Me.MSHFlexGrid1.CellBackColor = vbYellow
' Me.MSHFlexGrid1.Recordset.Move 1
' dpS.Move 1
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in weiter/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' weiter_Click


Private Sub zurück_Click()
 On Error GoTo fehler
 Me.MSHFlexGrid1.CellBackColor = IIf(Me.MSHFlexGrid1.Row = 0, vbActiveBorder, vbWhite)
 Me.MSHFlexGrid1.col = 3
 If Me.MSHFlexGrid1.Row > 1 Then
  If Me.Visible Then
   Me.MSHFlexGrid1.SetFocus
   On Error Resume Next
   Err.Clear
   SendKeys "{UP}"
   If Err.Number <> 0 Then
    Me.MSHFlexGrid1.Row = Me.MSHFlexGrid1.Row - 1
   End If
   On Error GoTo fehler
  Else
   Me.MSHFlexGrid1.Row = Me.MSHFlexGrid1.Row - 1
  End If
 End If
 If Me.Visible Then Call Me.MSHFlexGrid1.SetFocus
 Me.MSHFlexGrid1.CellBackColor = vbYellow
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in zurück/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub 'zurück_Click


Private Sub Form_Initialize()
'    rAd.Open sql, dbcn, adOpenDynamic, adLockOptimistic
'    SET Me.MSHFlexGrid1.DataSource = rAd

End Sub ' Form_Initialize

Private Sub LadDaten()
' PROVIDER=MSDASQL;dsn=MySQLpraxis;uid=...;pwd=...;database=quelle;
 Dim sql$
 On Error GoTo fehler
 If Not obStart Then
  hlese.Visible = False
 End If
 sql = SqlU(IIf(Me.nurHeuer, "SELECT * FROM (" & AbgehSQL & ") i WHERE YEAR(quelldt)>=year(SUBDATE(NOW(),INTERVAL 90 DAY)+1) OR quelldt < '19000000' OR ISNULL(quelldt) ORDER BY quelldt DESC, Name", AbgehSQL), Lese.obMySQL)
 If LenB(DBCnS) = 0 Then  ' DBCn.ConnectionString = "" THEN
  Call acon(quelleT)
 End If
 Dim rs As New ADODB.Recordset
 If Lese.obMySQL = 0 Then
  myFrag rs, SqlU(sql, Lese.obMySQL) ' um die Fehler im SQL-String besser zu erkennen
 End If
    dpS.ConnectionString = Lese.dbv.CnStr 'DBCn.ConnectionString  '& ";Persist Security Info=False"
    dpS.RecordSource = sql '"shape {" & sql & "} AS dokab"
    On Error Resume Next
    If Lese.obMySQL Then dpS.Refresh
    If Err.Number = -2147467259 And InStrB(dpS.ConnectionString, "Jet.") > 0 Then 'And Err.Description LIKE "*konnte nicht verwendet werden*" THEN
     Dim LoeDat$
     LoeDat = REPLACE$(Mid$(dpS.ConnectionString, InStr(dpS.ConnectionString, "Data Source") + 12), "mdb", "ldb")
     Call FSO.DeleteFile(LoeDat, True)
     On Error GoTo fehler
     dpS.Refresh
    End If
    On Error GoTo fehler
    DoEvents
    Screen.MousePointer = vbHourglass
    dpS.Visible = False
    Set Me.MSHFlexGrid1.DataSource = dpS
    Me.Zeilen = Me.MSHFlexGrid1.Rows
    Screen.MousePointer = vbNormal
    dpS.Visible = True
    Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LadDaten/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' LadDaten

Private Sub FlexAusricht()
 On Error GoTo fehler
    With MSHFlexGrid1
        .cols = 12
        .Redraw = False
        ' Spaltenbreite für Tabelle festlegen
        .ColWidth(0) = 300
        .ColWidth(1) = 600
        .ColWidth(2) = 1700
        .ColWidth(3) = 500
        .ColWidth(4) = 7000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 300
        .ColWidth(10) = 500

        ' Stil für Tabelle festlegen
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' Spaltenkopf Fett formatieren
        .Row = 0
        .col = 7
        .Text = "HA (Namen)"
        .col = 8
        .Text = "HA (Fälle)"
        .col = 0
        .RowSel = .FixedRows - 1
        On Error Resume Next
        .ColSel = .cols - 1
        On Error GoTo fehler
        .CellFontBold = True

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FlexAusricht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' FlexAusricht

Private Sub Form_Load() ' Labor eintragen
 On Error GoTo fehler
 Screen.MousePointer = vbHourglass
 Me.nurHeuer = 1
' Call LadDaten ' wird bereits beim Knopf "nurHeuer_Click" geladen
 Call FlexAusricht
 Call weiter_Click
 Me.WindowState = 2
 Screen.MousePointer = vbNormal
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Form_load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub 'Form_Load

Private Sub Form_Unload(Cancel As Integer)
 hlese.Visible = True
End Sub ' Form_Unload

Private Function IrfanString$()
 Dim DateiName$
 On Error GoTo fehler
 Me.MSHFlexGrid1.col = 3
 If LenB(PcDokPfad) = 0 Then getDokPfad
 With Me.MSHFlexGrid1
  DateiName = Stutz(REPLACE$(LCase$(.Text), "$\turbomed\dokumente", PcDokPfad))
  IrfanString = getIViewPfad & " " & DateiName & " /title='" & .TextMatrix(.Row, 4) & "'   (" & DateiName & ")" ' Environ("ProgramFiles") & "\irfanview\i_view32.exe
 End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in IrfanString/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' IrfanString

Private Sub Irfan_Click()
 On Error GoTo fehler
#If False Then
   Dim irf&
   irf = Shell(DatPfad, vbMaximizedFocus)
   Me.Hide
   Do While Not WartAufProzeß(irf)
   Loop
#Else
 Me.Hide
 Call Syncshell(IrfanString, 0, False, False)
#End If
   Me.Show
   Call SwitchTo(GFGW)
   Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Irfan_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Irfan_Click

Private Sub MSHFlexGrid1_DragDrop(source As Control, x As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' Code in den DragDrop-, MouseDown-, MouseMove- und MouseUp-Ereignissen der Tabelle aktiviert das Ziehen von Spalten
'-------------------------------------------------------------------------------------------
    If m_iDragCol = -1 Then Exit Sub    ' es wurde nicht gezogen
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub
    If MSHFlexGrid1.FixedCols = 1 And MSHFlexGrid1.MouseCol = 0 Then Exit Sub

    With MSHFlexGrid1
        .Redraw = False
        .ColPosition(m_iDragCol) = .MouseCol
        .Redraw = True
    End With
End Sub ' MSHFlexGrid1_DragDrop

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' Code in den DragDrop-, MouseDown-, MouseMove- und MouseUp-Ereignissen der Tabelle aktiviert das Ziehen von Spalten
'-------------------------------------------------------------------------------------------
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub
    If MSHFlexGrid1.MouseCol = 0 And MSHFlexGrid1.FixedCols = 1 Then Exit Sub
    xdn = x
    ydn = Y
    m_iDragCol = -1     ' Zieh-Attribut löschen
    m_bDragOK = True
End Sub ' MSHFlexGrid1_MouseDown

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' Code in den DragDrop-, MouseDown-, MouseMove- und MouseUp-Ereignissen der Tabelle aktiviert das Ziehen von Spalten
'-------------------------------------------------------------------------------------------
    ' testen, um zu sehen, ob zum dem Ziehen begonnen werden kann
    If Not m_bDragOK Then Exit Sub
    If Button <> 1 Then Exit Sub                        ' falsche Schaltfläche
    If m_iDragCol <> -1 Then Exit Sub                   ' es wird bereits gezogen
    If Abs(xdn - x) + Abs(ydn - Y) < 50 Then Exit Sub   ' noch nicht genug bewegt
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub         ' Spaltenkopf muß gezogen werden

    ' wenn Sie bis hierher gekommen sind, dann starten Sie den Ziehvorgang
    m_iDragCol = MSHFlexGrid1.MouseCol
    MSHFlexGrid1.Drag vbBeginDrag
End Sub ' MSHFlexGrid1_MouseMove

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' Code in den DragDrop-, MouseDown-, MouseMove- und MouseUp-Ereignissen der Tabelle aktiviert das Ziehen von Spalten
'-------------------------------------------------------------------------------------------
    m_bDragOK = False
End Sub ' MSHFlexGrid1_MouseUp

Private Sub MSHFlexGrid1_DblClick()
'-------------------------------------------------------------------------------------------
' Code in DblClick-Ereignis der Tabelle aktiviert Spaltensortierung
'-------------------------------------------------------------------------------------------
    Dim i%
    ' nur dann sortieren, wenn eine feste Zeile angeklickt wurde
    If MSHFlexGrid1.MouseRow >= MSHFlexGrid1.FixedRows Then Exit Sub
    i = m_iSortCol                  ' alte Spalte speichern
    m_iSortCol = MSHFlexGrid1.col   ' neue Spalte festlegen
    ' Sortiertyp inkrementieren
    If i <> m_iSortCol Then
        ' wenn eine neue Spalte geklickt wird, mit aufsteigender Sortierung beginnen
        m_iSortType = 1
    Else
        ' wenn dieselbe Spalte geklickt wird, zwischen aufsteigender und absteigender Sortierung umschalten
        m_iSortType = m_iSortType + 1
    If m_iSortType = 3 Then m_iSortType = 1
    End If
    DoColumnSort
End Sub ' MSHFlexGrid1_DblClick

Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' Führt Exchange-Sortierung von column m_iSortCol durch
'-------------------------------------------------------------------------------------------
    With MSHFlexGrid1
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .col = m_iSortCol
        .Sort = m_iSortType
        .Redraw = True
    End With
End Sub ' DoColumnSort

Private Sub Form_Resize()
    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single
    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight
        ' Schaltfläche 'Schließen' in untere rechte Ecke verschieben
        With .cmdClose
                sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE) - Me.dpS.Height - MARGIN_SIZE
                .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With
        With .cmdRefresh
                .Move Me.cmdClose.Left - .Width - MARGIN_SIZE, sngButtonTop
        End With
        With .Irfan
           .Move .Left, sngButtonTop
        End With
        With .Photoimpact
           .Move .Left, sngButtonTop
        End With
        With .abhaken
         .Move .Left, sngButtonTop
        End With
        With .zurück
         .Move .Left, sngButtonTop
        End With
        With .weiter
         .Move .Left, sngButtonTop
        End With
        With .dpS
            .Move .Left, sngButtonTop + MARGIN_SIZE + Me.cmdRefresh.Height
        End With
        .MSHFlexGrid1.Move MARGIN_SIZE, _
            MARGIN_SIZE, _
            sngScaleWidth - (2 * MARGIN_SIZE), _
            sngButtonTop - (2 * MARGIN_SIZE)
        .Zeilen.Left = .abhaken.Left + .abhaken.Width + 300
        .Zeilen.Top = .abhaken.Top
        .inTurbomedAnzeigen.Left = .Zeilen.Left + .Zeilen.Width + 300
        .inTurbomedAnzeigen.Top = .Zeilen.Top
        .nurHeuer.Left = .inTurbomedAnzeigen.Left + .inTurbomedAnzeigen.Width + 300
        .nurHeuer.Top = .inTurbomedAnzeigen.Top
    End With
    Exit Sub
Form_Resize_Error:
    ' Fehler bei negativen Werten vermeiden
    Resume Next
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in zurück/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Form_Resize

Private Sub cmdRefresh_Click()
  'Dies wird nur für Mehrbenutzeranwendungen benötigt
  On Error GoTo RefreshErr
  dpS.Refresh
  Me.Zeilen = Me.MSHFlexGrid1.Rows
  Me.MSHFlexGrid1.Row = 0
  Call weiter_Click
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub ' cmdRefresh_Click

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  dpS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub ' cmdUpdate_Click

Private Sub cmdClose_Click()
    Unload Me
End Sub ' cmdClose_Click


Private Sub Photoimpact_Click()
  Dim irf&
  On Error GoTo fehler
  If LenB(PcDokPfad) = 0 Then getDokPfad
  With Me.MSHFlexGrid1
   DatPfad = Environ("ProgramFiles") & "\Ulead Systems\Ulead PhotoImpact 6\iedit.exe " & "" & Stutz(REPLACE$(LCase$(.TextMatrix(.Row, 3)), "$\turbomed\dokumente", PcDokPfad)) + ""
  End With
  irf = Shell(DatPfad, vbMaximizedFocus)
'   Me.Hide
'   Do While WartAufProzeß(irf)
'   Loop
'   Me.Show
'   Call SwitchTo(GFGW)
   Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Photoimpact_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Photoimpact_Click

