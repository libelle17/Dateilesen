VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Auswahl von Datei und Datenbank"
   ClientHeight    =   7695
   ClientLeft      =   2010
   ClientTop       =   1380
   ClientWidth     =   9750
   Icon            =   "Datenbank.frx":0000
   LinkTopic       =   "Dialog"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox nuraktfaelle 
      Caption         =   "&nur aktuelle Fälle"
      Height          =   375
      Left            =   5640
      TabIndex        =   43
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox NurInTabelle 
      Caption         =   "&Nur dorthin einlesen"
      Height          =   195
      Left            =   2400
      TabIndex        =   18
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CheckBox BeziehungsfehlerSpeichern 
      Caption         =   "Be&ziehungsfehler speichern?"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CheckBox bereinigeFormInhFeld 
      Caption         =   "&FormInhFeld bereinigen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   20
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CheckBox ÜberTabelle 
      Caption         =   "&Über Tabelle `inl`einlesen"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Zeitgewinn 
      Height          =   285
      Left            =   1320
      TabIndex        =   40
      Top             =   420
      Width           =   4095
   End
   Begin VB.CheckBox SammelInsert 
      Caption         =   "&Sammelinsert(MySQL)"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton DatenbankAuswahl 
      Caption         =   "&Datenbankauswahl"
      Height          =   255
      Left            =   2400
      TabIndex        =   38
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox BDTAend 
      BackColor       =   &H80000000&
      ForeColor       =   &H8000000C&
      Height          =   285
      Left            =   5520
      TabIndex        =   36
      Top             =   420
      Width           =   2535
   End
   Begin VB.CommandButton BDTaussuchen 
      Caption         =   "a&ussuchen"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton EmailsAussuchen 
      Caption         =   "aussuc&hen"
      Height          =   285
      Left            =   7920
      TabIndex        =   29
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton AccAussuchen 
      Caption         =   "au&ssuchen"
      Height          =   405
      Left            =   7080
      TabIndex        =   5
      Top             =   990
      Width           =   975
   End
   Begin VB.CommandButton komprimieren 
      Caption         =   "&komprimieren"
      Height          =   255
      Left            =   7320
      TabIndex        =   35
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox LDateiAend 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000013&
      Height          =   285
      Left            =   7200
      TabIndex        =   34
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox LDatei 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   33
      Top             =   2880
      Width           =   7215
   End
   Begin VB.CheckBox obVglMitLetzterEinlesung 
      Caption         =   "&Vergleich mit letzter Einlesung"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox LAktDatum 
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox AlterTab 
      Caption         =   "Tabellenfelder nötigenfalls verl&ängern?"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox LaborPfadBeispiel 
      Height          =   285
      Left            =   2040
      TabIndex        =   26
      Top             =   6600
      Width           =   6495
   End
   Begin VB.TextBox BDTDatei 
      Height          =   375
      Left            =   1290
      TabIndex        =   1
      Top             =   0
      Width           =   7305
   End
   Begin VB.CheckBox obBDT 
      Caption         =   "&BDT-Datei:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.CheckBox LaborQuerNeu 
      Caption         =   "neu"
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      Top             =   6045
      Width           =   615
   End
   Begin VB.CheckBox LaborDirektNeu 
      Caption         =   "neu"
      Height          =   375
      Left            =   2640
      TabIndex        =   22
      Top             =   5640
      Width           =   735
   End
   Begin VB.CheckBox LaborQuerVerb 
      Caption         =   "Labor &querverbinden"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6075
      Width           =   2295
   End
   Begin VB.CheckBox LaborDirektEinlesen 
      Caption         =   "&Labordirektdateien einlesen"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CheckBox VorladenFFI 
      Caption         =   "Vo&rladen Formularfeldinhalte"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox Pat_IDBis 
      Height          =   285
      Left            =   2400
      TabIndex        =   14
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Pat_IDVon 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   3720
      Width           =   855
   End
   Begin VB.CheckBox ZurücksetzenLAktDat 
      Caption         =   "&Zurücksetzen des letzten Aktualiserungsdatums, ggf. nur für Pat., die Einträge enthalten nach dem:"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CheckBox TabellenEntleeren 
      Caption         =   "Tabellen vorher ent&leeren?"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Start 
      Caption         =   "==>       S&tart"
      Height          =   615
      Left            =   3720
      TabIndex        =   32
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CheckBox obmitEmails 
      Caption         =   "mit &Emails?"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox EmDatei 
      Height          =   285
      Left            =   1320
      TabIndex        =   28
      Top             =   7080
      Width           =   6495
   End
   Begin MSComDlg.CommonDialog CommonDialogDialog 
      Left            =   1680
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton obMySQL 
      Caption         =   "&MySQL"
      Height          =   285
      Left            =   90
      TabIndex        =   6
      Top             =   1440
      Width           =   1425
   End
   Begin VB.OptionButton obAcc 
      Caption         =   "&Access"
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   1020
      Width           =   1065
   End
   Begin VB.TextBox MdB 
      Height          =   405
      Left            =   1290
      TabIndex        =   4
      Top             =   990
      Width           =   5745
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   8400
      TabIndex        =   31
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   8400
      TabIndex        =   30
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label KlZu 
      Caption         =   ")"
      Height          =   255
      Left            =   8160
      TabIndex        =   41
      Top             =   420
      Width           =   135
   End
   Begin VB.Label ConStrLabel 
      Height          =   495
      Left            =   120
      TabIndex        =   39
      Top             =   1800
      Width           =   7815
   End
   Begin VB.Label Zeitgewinn1 
      Height          =   255
      Left            =   5640
      TabIndex        =   37
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Label LaborpfadBez 
      Caption         =   "Laborp&fad (Beispieldatei):"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Pat_IDBisBez 
      Caption         =   "-"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label PatIDVonBez 
      Caption         =   "&Pat_ID von:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   975
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public hlese As Lese
Public lBDTDatei$
Public noreact%
Public ConstrCn$
Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1

Private Sub AccAussuchen_Click()
 Call AccTest
End Sub ' AccAussuchen_Click

Private Sub AccAussuchen_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' AccAussuchen_KeyDown

Public Sub BDTAend_Change()
 If Me.Visible Then Screen.MousePointer = vbHourglass
 If Me.BDTAend <> vNS And Me.LDateiAend <> vNS Then
  Me.Zeitgewinn.Caption = Format$(CDate(Me.BDTAend) - CDate(Me.LDateiAend), "##,##0.###") & " Tage Zeitgewinn (" & Me.LDateiAend & " - "
 Else
  Me.Zeitgewinn.Caption = vNS
 End If
 If Me.Visible Then Screen.MousePointer = vbDefault
End Sub ' BDTAend_Change

Private Sub BDTaussuchen_Click()
  Call BDTDateiDialog
End Sub ' BDTaussuchen_Click

Private Sub BDTaussuchen_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' BDTaussuchen_KeyDown

Private Sub BDTDatei_Change()
 If Me.Visible Then Screen.MousePointer = vbHourglass
 If FSO.FileExists(Me.BDTDatei) Then Me.BDTAend = FSO.GetFile(Me.BDTDatei).DateLastModified
End Sub ' BDTDatei_Change()

Private Sub BDTDatei_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
 If Me.Visible Then Screen.MousePointer = vbDefault
End Sub ' BDTDatei_KeyDown

Private Sub DatenbankAuswahl_Click()
 Call lies.dbv.Auswahl("", "anamnesebogen", "Patientendaten")
' SetDBCn Nothing
 DBCnS = vNS
 Call Me.FrmLEinlesung
 Call Me.BDTAend_Change
End Sub ' DatenbankAuswahl_Click

Private Sub dbv_wCnAendern(CnStr As String)
 Static altCnStr$
 If CnStr <> altCnStr Then ' 13.4.08
  Me.ConstrCn = dbv.CnStr
  Me.ConStrLabel = dbv.Constr
  obStart = True
  Me.obMySQL = InStr(1, dbv.CnStr, "MYSQL", vbTextCompare) <> 0
' obStart = False ' 13.4.08, da bei Auswahl von Access im Lese-Fenster sonst dann im Dialog-Fenster Zeitgewinn nicht stimmt
  Me.obAcc = Not Me.obMySQL
  If Me.obAcc Then
   Me.MdB = Me.dbv.Datei
  End If
' SetDBCn Me.dbv.wCn ' 29.3.08
  altCnStr = CnStr
 End If
 obStart = False
End Sub ' dbv_wCnAendern(CnStr AS String)

Private Sub EmailsAussuchen_Click()
  Call EmDateiDialog
End Sub ' EmailsAussuchen_Click

Private Sub EmailsAussuchen_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' EmailsAussuchen_KeyDown

Private Sub Form_Activate()
 If Me.hlese.Visible Then Me.hlese.Hide
 If Me.BDTDatei = vNS Then
  Me.BDTDatei = getLDatei(hlese.dlg.lBDTDatei, "*.bdt")
 End If
 If FSO.FileExists(Me.BDTDatei) Then
  Dim gesb# ' ab 21.2.18 Grenze von long überschritten!
  gesb = FSO.GetFile(Me.BDTDatei).size
  If gesb > 500000000 Then
   Me.LaborDirektEinlesen = 1
   Me.LaborQuerVerb = 1
  Else
   Me.LaborDirektEinlesen = 0
   Me.LaborQuerVerb = 0
  End If
 End If
'   Me.BDTDatei = getLDatei(Me.BDTDatei, "*.bdt")
'    Me.BDTAend = FSO.GetFile(Me.BDTDatei).DateLastModified
End Sub ' Form_Activate

Private Sub Form_Deactivate()
 On Error Resume Next
 If Not Me.hlese.Visible Then Me.hlese.Show
End Sub ' Form_Deactivate

Private Sub Form_GotFocus()
'    IF Me.BDTDatei = "" THEN Me.BDTDatei = getLDatei(hlese.dlg.lBDTDatei, "*.bdt")
''    dlg.BDTDatei = getLDatei(dlg.BDTDatei, "*.bdt")
''     Me.BDTAend = FSO.GetFile(Me.BDTDatei).DateLastModified
End Sub ' Form_GotFocus

Private Sub Form_Initialize()
'
 
End Sub ' Form_Initialize

Private Sub komprimieren_Click()
 If Me.Visible Then Screen.MousePointer = vbHourglass
 Me.hlese.obAusgedehnt = True
 Call kompakt(Me.hlese)
 If Me.Visible Then Screen.MousePointer = vbDefault
End Sub ' komprimieren_Click

Private Sub LaborDirektEinlesen_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' LaborDirektEinlesen_KeyDown

Private Sub LaborDirektNeu_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub LaborPfadBeispiel_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)

End Sub ' LaborPfadBeispiel_KeyDown

Private Sub LaborpfadBez_Click()
 LaborPfadBeispiel = Lese.LaborPfadDialog(LaborPfadBeispiel)
End Sub ' LaborpfadBez_Click

Private Sub LaborQuerNeu_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' LaborQuerNeu_KeyDown

Private Sub LaborQuerVerb_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' LaborQuerVerb_KeyDown

Private Sub mitEmails_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' mitEmails_KeyDown

Private Sub LDateiAend_Change()
 Call Me.BDTAend_Change
End Sub ' LDateiAend_Change

Private Sub MdB_Change()
 dbv.Datei = Me.MdB
 hlese.Ziel = Me.MdB
End Sub ' MdB_Change

Private Sub MdB_Click()
 dbv.Datei = Me.MdB
 hlese.Ziel = Me.MdB
End Sub ' MdB_Click

Private Sub obMySQL_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' obMySQL_KeyDown

Private Sub obAcc_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' obAcc_KeyDown

Private Sub obVglMitLetzterEinlesung_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' obVglMitLetzterEinlesung_KeyDown

Private Sub obAcc_GotFocus()
' Me.obAcc = True
' Call AccTest
End Sub ' obAcc_GotFocus

Private Function AccTest()
 If (Not imAufbauLese And Not imAufbauDialog) Or Me.MdB = "" Then
  Call MdBFestleg
 End If
 Call dbv.cnVorb("", "anamnesebogen", "Patientendaten")
' IF Not obStart THEN Call PutReg(Me.hlese)
' Call hlese.ConstrFestleg(1, hlese)
' Call acon(QuelleT, accDtb)
' IF Not obStart THEN Call hlese.HolEinstvonDB
End Function ' AccTest

Private Sub obBDT_Click()
 If Not imAufbauDialog And Me.BDTDatei = "" Then
  Call BDTDateiDialog
 End If
End Sub ' obBDT_Click

Private Sub CancelButton_Click()
' ggf ursprüngliche Einstellungen wiederherstellen
 If Me.Visible Then Screen.MousePointer = vbHourglass
 Call HolReg(hlese)
' Me.Hide
 Me.hlese.Show
 If Me.Visible Then Screen.MousePointer = vbDefault
End Sub ' CancelButton_Click

Private Sub CancelButton_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' CancelButton_KeyDown

Private Sub EmDatei_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' EmDatei_KeyDown

Private Sub Form_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' Form_KeyDown

Private Sub Form_Load()
 Set Me.dbv = hlese.dbv
 If Not obStart Then
  hlese.Visible = False
  imAufbauDialog = True
' Call HolReg(hlese)
' Call hlese.ConstrFestleg(0, hlese)
  Call acon(quelleT)
  Call hlese.HolEinstvonDB
  imAufbauDialog = False
 Else
 End If
 With Me.CommonDialogDialog
  .Orientation = cdlLandscape
  .flags = 0
' .Flags = .Flags OR FileOpenConstants.cdlOFNExplorer
  .flags = .flags Or FileOpenConstants.cdlOFNHideReadOnly ' Schreibgeschützt-Checkbox entfernen
  .flags = .flags Or FileOpenConstants.cdlOFNLongNames ' '   Lange Dateinamen erlauben (nur sinnvoll bei Nicht-Win95-Design)
  .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
  .flags = .flags Or FileOpenConstants.cdlOFNPathMustExist
 End With
End Sub ' Form_Load()

Public Function FrmLEinlesung()
 Dim rs As New ADODB.Recordset ', catx As New ADOX.Catalog
 Dim zl%, ErrNr&, ErrDes$
 On Error GoTo fehler
' Exit Function
 If LenB(DBCn) = 0 Or DBCn.State = 0 Then
'   Call CnOpen(False, ConStr, Me.hlese)
   Call acon(quelleT)
   If lies.obMySQL Then Call myEFrag("USE " & hlese.MyDB)
 End If
 Me.obVglMitLetzterEinlesung = 1
' Me.obVglMitLetzterEinlesung.Enabled = False ' auskommentiert 13.5.23
' catx.ActiveConnection = ConStr
' zl = catx.Tables("eintragszahlen").Columns("dateiaend").Properties.Count
 On Error Resume Next
 For zl = 1 To 2
  If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
'  sql = "SELECT datei, dateiaend FROM `eintragszahlen` WHERE NOT ISNULL(datei) AND NOT ISNULL(zp3) AND (zp3 < " & DatFor_k("6.1.2007") & " OR fallzahl > 10) ORDER BY beginn DESC;"
  sql = "SELECT SUBSTRING_INDEX(datei, '\\', -1) datei, dateiaend FROM `eintragszahlen` WHERE NOT ISNULL(datei) AND NOT ISNULL(zp3) AND fallzahl > 10 AND NOT datei LIKE '%HB_%' ORDER BY beginn DESC LIMIT 30;" ' 4.10.20, 28.11.21, 22.10.22
  myFrag rs, sql, , , , , , True, ErrNr, ErrDes
  If ErrNr = 0 Then
   Exit For
  Else
   If zl = 0 Then
    Call myEFrag("ALTER TABLE `eintragszahlen` add COLUMN `dateiaend` `datetime`" & IIf(lies.obMySQL, " after `datei`", ""))
   Else
    MsgBox ("Fehler bei der Ausführung von: '" & sql & "'" & vbCrLf & "bei der Verbindung: '" & DBCn & "':" & vbCrLf & Err.Number & ": " & Err.Description)
    Exit Function
   End If
  End If
 Next zl
 Me.LDatei = vNS
 Me.LDateiAend = vNS
 On Error GoTo fehler
 Dim SuchVerz(1)
 SuchVerz(0) = "C:\TMExport\"
 SuchVerz(1) = üVerz
 If Not rs Is Nothing Then
  If Not rs.BOF Then
   Me.LDatei = rs!Datei
   Do While Not rs.EOF
    Dim ai%, Datei$
    For ai = 0 To UBound(SuchVerz)
     Datei = SuchVerz(ai) & rs!Datei
     If FSO.FileExists(Datei) Then
' Für Bearbeitungen während des Sommerzeitwechsels
    Dim diff#, diff2#
      diff = FSO.GetFile(Datei).DateLastModified - rs!DateiAend - 1 / 24
      diff2 = diff + 2 / 24
      If diff < 0 Then diff = -diff
      If diff2 < 0 Then diff2 = -diff2
      If FSO.GetFile(Datei).DateLastModified = rs!DateiAend Or (diff < 0.000000001) Or (diff2 < 0.000000001) Then
       Me.LDatei = Datei
       Me.LDateiAend = rs!DateiAend
       Me.obVglMitLetzterEinlesung.Enabled = True
       Exit Do
      End If
     End If
    Next
    rs.Move 1
   Loop
  End If
  If Me.obVglMitLetzterEinlesung.Enabled = False Then
   Me.obVglMitLetzterEinlesung = False
  End If
  rs.Close
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FrmLEinlesung/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FrmEinlesung

Private Sub EmdateiBez_Click()
End Sub ' EmdateiBez_Click()

Private Sub Form_Terminate()
'
End Sub ' Form_Terminate

Private Sub Form_Unload(Cancel%)
 hlese.Visible = True
 Call PutReg(Me.hlese)
 Call Me.hlese.PutEinstAufDB
End Sub ' form_unload

Private Sub MdB_KeyDown(KeyCode%, Shift%)
 hlese.Ziel = Me.MdB
 Call Key(KeyCode, Shift, Me)
End Sub ' MdB_KeyDown

Private Sub obAcc_DblClick()
 Call AccTest
End Sub ' obAcc_DblClick

Private Sub obBDT_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' obBDT_KeyDown

Private Sub obmitEmails_Click()
 If Not imAufbauDialog And Me.EmDatei = vNS Then
  Call EmDateiDialog
 End If
End Sub ' obmitEmails_Click

Private Sub doobMyQuelle_Click(nr%)
  Dim altBDTDatei$
  hlese.obMySQL = True
  Call Zinit(hlese.obMySQL)
  Call hlese.PutEinstAufDB
' Call hlese.ConstrFestleg(0, hlese)
  Call acon(quelleT)
  altBDTDatei = Me.BDTDatei
  Call hlese.HolEinstvonDB
  Me.BDTDatei = altBDTDatei
'  IF FSO.FileExists(Me.BDTDatei) THEN Me.BDTAend = FSO.GetFile(Me.BDTDatei).DateLastModified
'  Call FrmLEinlesung(Me)
  Set Lese.pataw = Nothing
  If Me.Visible Then Screen.MousePointer = vbDefault
End Sub ' doobMyQuelle_Click

Private Sub obAcc_Click()
  Dim altChangeStill%
  If Not obStart Then
'  Lies.obMySQL = False
   obStart = True
   altChangeStill = dbv.changeStill
   dbv.changeStill = True
   If Me.obAcc <> False Then Me.dbv.ODBC = "Microsoft Access Drivers (*.mdb)"
   dbv.changeStill = altChangeStill
   obStart = False
   If Me.Visible Then Screen.MousePointer = vbHourglass
   lies.obAcc = True
   lies.Ziel = Me.MdB
  End If ' 13.4.08
  SetDBCn dbv.wCn, dbv.CnStr ' 13.4.08
  Call Me.FrmLEinlesung
  Call Me.BDTAend_Change
'  END IF
  imAufbauDialog = True
'  Call AccTest
  imAufbauDialog = False
  Set Lese.pataw = Nothing
  If Not obStart Then If Me.Visible Then Screen.MousePointer = vbDefault
End Sub ' obAcc_Click()

Private Sub obMySQL_Click()
 If Me.obMySQL <> False And Me.dbv.ODBC Like "*Access*" Then Me.obMySQL = False
 If Not obStart Then
  If Me.Visible Then Screen.MousePointer = vbHourglass
  obStart = True
  hlese.MyDB = hlese.dbv.DaBa
  obStart = False
  hlese.obMySQL = True
  Call doobMyQuelle_Click(2)
  Call Me.FrmLEinlesung
  Call Me.BDTAend_Change
  If Me.Visible Then Screen.MousePointer = vbDefault
 End If
End Sub ' obMyQuelle_Click()

Private Sub obVglMitLetzterEinlesung_Click()
 Call AuswEnable
End Sub ' obVglMitLetzterEinlesung_Click

Private Sub Pat_IDBis_GotFocus()
 Me.Pat_IDBis.SelLength = Len(Me.Pat_IDBis)
End Sub ' Pat_IDBis_GotFocus

Private Sub Pat_IDVon_GotFocus()
 Me.Pat_IDVon.SelLength = Len(Me.Pat_IDVon)
End Sub ' Pat_IDVon_GotFocus

Private Sub Zeitgewinn_Click()
 Call EintragZusatz
 Call BDTAend_Change
End Sub ' Zeitgewinn_Click

Private Sub ZurücksetzenLAktDat_Click()
 Call AuswEnable
End Sub ' ZurücksetzenLAktDat_Click

Private Sub TabellenEntleeren_Click()
 Call AuswEnable
End Sub ' TabellenEntleeren_Click

Private Function AuswEnable()
 Me.TabellenEntleeren.Enabled = True
 Me.ZurücksetzenLAktDat.Enabled = True
' Me.Pat_IDVon.Enabled = False
' Me.Pat_IDBis.Enabled = False
 Me.LAktDatum.Enabled = False
 Me.nuraktfaelle.Enabled = True
 If Me.LDateiAend <> vNS Then
  Me.obVglMitLetzterEinlesung.Enabled = True
 End If
 If Me.obVglMitLetzterEinlesung Then
  Me.TabellenEntleeren.Enabled = False
  Me.ZurücksetzenLAktDat.Enabled = False
  Me.nuraktfaelle.Enabled = False
 ElseIf Me.ZurücksetzenLAktDat Then
  Me.obVglMitLetzterEinlesung.Enabled = False
  Me.TabellenEntleeren.Enabled = False
  Me.LAktDatum.Enabled = True
'  Me.Pat_IDVon.Enabled = True
'  Me.Pat_IDBis.Enabled = True
 ElseIf Me.TabellenEntleeren Then
  Me.obVglMitLetzterEinlesung.Enabled = False
  Me.ZurücksetzenLAktDat.Enabled = False
 Else
  Me.obVglMitLetzterEinlesung.Enabled = True
 End If
End Function ' AuswEnable

Private Sub OKButton_Click()
' Unload Me
 Call PutReg(Me.hlese)
 Call Me.hlese.PutEinstAufDB
 Call Me.dbv.RegSpeichern
' Me.Hide
 Me.hlese.Show
End Sub ' OKButton_Click

Private Sub OKButton_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' OKButton_KeyDown

Private Sub Pat_IDBis_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' Pat_IDBis_KeyDown

Private Sub Pat_IDVon_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' Pat_IDVon_KeyDown

Public Sub Start_Click()
 If Not obStart Then
 If Not lies.obMySQL Then Me.hlese.obAusgedehnt = True
 Call PutReg(Me.hlese) ' da in hide nicht dabei
' hlese.Visible = True
 hlese.Show
' Me.Hide
 Call Einlies
 If Not lies.obMySQL And Me.hlese.obAusgedehnt Then Call kompakt(Me.hlese)
 End If
End Sub ' Start_Click()

Private Sub Start_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' KeyDown

Private Sub TabellenEntleeren_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' TabellenEntleeren_KeyDown

Private Sub SammelInsert_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' SammelInsert_KeyDown

Private Sub bereinigeFormInhFeld_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' bereinigeFormInhFeld_KeyDown

Private Sub VorladenFFI_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' VorladenFFI_KeyDown

Private Sub ÜberTabelle_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' ÜberTabelle_KeyDown

Private Sub NurInTabelle_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' NurInTabelle_KeyDown

Private Sub ZurücksetzenLAktDat_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' ZurücksetzenLAktDat_KeyDown

Function EmDateiDialog()
 If Me.obmitEmails Then
  With Me.CommonDialogDialog
   .DialogTitle = "Email-Excel-Datei"
   .initDir = FSO.GetParentFolderName(Me.EmDatei)
   If .initDir = vNS Then .initDir = pügVerz   ' uverz & "anamnese"
   .Filename = "*.xls"
   .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
   .ShowOpen
   If FSO.FileExists(.Filename) Then
    Me.EmDatei = .Filename
   End If
  End With
 End If
End Function ' EmDateiDialog

Function BDTDateiDialog()
 If Me.obBDT Then
  With Me.CommonDialogDialog
   .Filename = "*.BDT"
   .DialogTitle = "Einzulesende BDT-Datei"
   If FSO.FileExists(Me.BDTDatei) Then
    .initDir = FSO.GetParentFolderName(FSO.GetFile(Me.BDTDatei).path)
   Else
    .initDir = üVerz ' uverz & "tmexport"
   End If
   .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
   .ShowOpen
   If FSO.FileExists(.Filename) Then
    Me.BDTDatei = .Filename
'    Me.BDTAend = FSO.GetFile(Me.BDTDatei).DateLastModified
   End If
  End With
 End If
End Function ' BDTDateiDialog()

Public Function MdBFestleg()
  Dim cat As New ADOX.Catalog
  With Me.CommonDialogDialog
   .DialogTitle = "MdB festlegen"
   .initDir = FSO.GetParentFolderName(Me.MdB)
   If .initDir = vNS Then .initDir = aVerz ' uverz & "Anamnese"
   .Filename = "*.mdb"
   .flags = .flags And Not FileOpenConstants.cdlOFNFileMustExist
   .ShowOpen
   ZielDbS = .Filename
    If LCase$(Right$(ZielDbS, 3) <> "mdb") Then ZielDbS = ZielDbS & ".mdb"
   If Not FSO.FileExists(ZielDbS) And InStrB(ZielDbS, "*") = 0 Then
'    SET dbe = New DAO.DBEngine
'    SET ws = dbe.Workspaces(0)
'    Call ws.CreateDatabase(.FileName, dbLangGeneral)
    cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ZielDbS
#If False Then ' Kommentar 21.9.08 G.Schade
    Call MachSammelTabA
#End If
   Else
    ZielDbS = .Filename
   End If
   If FSO.FileExists(ZielDbS) Then
    Me.MdB = ZielDbS
    hlese.Ziel = ZielDbS
   End If
  End With
End Function ' MdBFestleg

' in Start_Click
Public Sub Einlies()
 Dim i%
 obMitAlterTab = Me.AlterTab
 obVorber = Me.VorladenFFI
 If Me.obBDT Then
  Do
   For i = 1 To 2
    If FSO.FileExists(Me.BDTDatei) Then Exit Do
    Call BDTDateiDialog
   Next i
   Exit Sub
  Loop ' Me.obBDT Then
 End If
 If Me.LaborDirektEinlesen And (Me.LaborPfadBeispiel = vNS Or Me.LaborPfadBeispiel = "0") Then
  LaborPfadBeispiel = Lese.LaborPfadDialog(LaborPfadBeispiel)
 End If
 Call Lese.ProgStart
 Call doEinles(Me.TabellenEntleeren)
 Call hlese.ProgEnde
End Sub ' EinlesenVorb_Click()

' 30.3.21: nur in Einlies()
Private Sub doEinles(obevtlAlle%)
' Const obZS% = -1
 Dim Fil As File, erg%, sql$, obAktZeit%
 Dim oblies%
 On Error GoTo fehler
 
 Dim EinlAb&, EinlBis&
 Call acon(quelleT, , , , , , , True)
 ComTrans
 AllePat = 0
 If Me.obBDT Then
  Set Fil = FSO.GetFile(Me.BDTDatei)
  Me.hlese.QDatei = Fil.path
  Me.hlese.QDatum = Fil.DateLastModified
  If IsNumeric(Fil.size) Then Me.hlese.GesBytes = Format$(Fil.size, "###,###,###,###,###,###,##0")
  If obevtlAlle Then
   erg = MsgBox("Wollen Sie wirklich alle Patientendatensätze aus `" & DefDB(DBCn) & "` löschen?", vbYesNo)
   If erg = vbNo Then Exit Sub
   AllePat = True
  ElseIf Me.ZurücksetzenLAktDat Then
   sql = "UPDATE `namen` SET aktzeit = 0"
   If Not IsNull(Me.LAktDatum) Then
    If IsDate(Me.LAktDatum) Then
     sql = sql & " WHERE aktzeit > " & DatFor_k(Me.LAktDatum)
     obAktZeit = -1
    End If
   End If
'   update namen n LEFT JOIN aktfv f ON n.pat_id=f.pat_id SET n.aktzeit = 0 WHERE NOT ISNULL(f.pat_id);
   Dim obvon%, obbis%
   If IsNumeric(Me.Pat_IDVon) Then If Me.Pat_IDVon > 0 Then obvon = -1
   If IsNumeric(Me.Pat_IDBis) Then If Me.Pat_IDBis > 0 Then obbis = -1
   If obvon Or obbis Then
    sql = sql & IIf(obAktZeit, " AND", " WHERE") & " pat_id "
    If obvon Then sql = sql & ">= " & Me.Pat_IDVon: If obbis Then sql = sql & " AND pat_id "
    If obbis Then sql = sql & "<= " & Me.Pat_IDBis
   End If
   If Me.nuraktfaelle <> 0 Then
    sql = sql & IIf(obAktZeit Or obvon Or obbis, " AND", " WHERE") & " pat_id IN (SELECT pat_id FROM aktfv)"
   End If
   erg = MsgBox("Wollen Sie wirklich alle " & IIf(Me.nuraktfaelle <> 0, " aktuellen ", "") & " Patienten" & IIf(obvon, " ab " & Me.Pat_IDVon, vNS) & IIf(obbis, " bis " & Me.Pat_IDBis, vNS) & " zurücksetzen?", vbYesNo)
   If erg = vbNo Then Exit Sub
   Call myEFrag(sql)
  End If
 End If ' Me.obBDT Then
 Call EintragStart(Me.hlese)
 T1 = Now
 If IsNumeric(Me.Pat_IDVon) Then EinlAb = Me.Pat_IDVon Else EinlAb = 0
 If IsNumeric(Me.Pat_IDBis) Then EinlBis = Me.Pat_IDBis Else EinlBis = 0
 If (Me.obBDT = 0) Then Me.BDTDatei = ""
 Dim FilName$
 If Not Me.obBDT Then FilName = "" Else FilName = Fil.name
 obHausBesuch = (UCase$(Me.BDTDatei) Like "*\HB_*.BDT")
 Call GesLies(Me.hlese, Me.BDTDatei, FilName, EinlAb, EinlBis, Me.LaborDirektEinlesen, Me.LaborDirektNeu, Me.LaborQuerVerb, Me.LaborQuerNeu, Me.obmitEmails, Me.EmDatei, oblies%)
 DoKassenkategorienBestimmen
 If Command = "auto" Then
 Else
'  MsgBox CStr((Now - T1) * 60 * 60 * 24) + " s"
  Lese.Prozent = "100 %"
  syscmd 4, IIf(oblies, "Fertig mit Einlesen von ", "nicht eingelesen: ") & Me.BDTDatei & ", Zeitdauer gesamt: " & Format$((Now - T1) * 60 * 60 * 24, "###,###,###,###,###,###,##0") & "s, " & "Zeitdauer nach Haupteinlesen: " & Format$((Now - T1a) * 60 * 60 * 24, "###,###,###,###,###,###,##0") & " s"
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doEinles/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' doEinles

