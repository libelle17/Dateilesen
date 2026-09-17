VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm Lese 
   BackColor       =   &H8000000C&
   Caption         =   "Turbomeddaten nach MySQL übertragen"
   ClientHeight    =   5055
   ClientLeft      =   1290
   ClientTop       =   1335
   ClientWidth     =   15600
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   1  'Oben ausrichten
      Height          =   5055
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   15540
      TabIndex        =   0
      Top             =   0
      Width           =   15600
      Begin VB.TextBox QDatum 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   26
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox Ziel 
         Height          =   285
         Left            =   10560
         TabIndex        =   23
         Top             =   3960
         Width           =   4935
      End
      Begin VB.CommandButton ÜbertrageCd 
         Caption         =   "&Übertrage"
         Height          =   275
         Left            =   0
         TabIndex        =   18
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton nachCd 
         Caption         =   "&nach"
         Height          =   275
         Left            =   8280
         TabIndex        =   20
         Top             =   3960
         Width           =   615
      End
      Begin VB.OptionButton obAcc 
         Caption         =   "&Access"
         Height          =   275
         Left            =   9720
         TabIndex        =   22
         Top             =   3960
         Width           =   855
      End
      Begin VB.OptionButton obMy 
         Caption         =   "M&ySQL"
         Height          =   275
         Left            =   8880
         TabIndex        =   21
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox QDatei 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   3960
         Width           =   4335
      End
      Begin VB.PictureBox Picture2 
         Height          =   315
         Left            =   15210
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   24
         Top             =   4350
         Width           =   315
         Begin VB.Label SBez 
            BackColor       =   &H00E0E0E0&
            Caption         =   "S"
            Height          =   225
            Left            =   30
            TabIndex        =   25
            Top             =   30
            Width           =   195
         End
      End
      Begin VB.TextBox Ausgabe 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3885
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   17
         Top             =   30
         Width           =   15435
      End
      Begin VB.TextBox Ende 
         Enabled         =   0   'False
         Height          =   285
         Left            =   14160
         TabIndex        =   16
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox GesDauer 
         Enabled         =   0   'False
         Height          =   285
         Left            =   12600
         TabIndex        =   14
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Beginn 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10200
         TabIndex        =   12
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox Sekunden 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         TabIndex        =   10
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Prozent 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   8
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox GesBytes 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox Zeilen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox Bytes 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Top             =   4320
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label EndeBez 
         Caption         =   "Ende:"
         Height          =   275
         Left            =   13680
         TabIndex        =   15
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label GesamtDauerBez 
         Caption         =   "Ges'dauer vor.:"
         Height          =   275
         Left            =   11520
         TabIndex        =   13
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label sbisherBez 
         Caption         =   "s bisher; Beginn:"
         Height          =   275
         Left            =   9000
         TabIndex        =   11
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label ProzentBez 
         Caption         =   "%,"
         Enabled         =   0   'False
         Height          =   275
         Left            =   7560
         TabIndex        =   9
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label GleichBez 
         Caption         =   " ="
         Enabled         =   0   'False
         Height          =   275
         Left            =   6360
         TabIndex        =   7
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label GesBytesBez 
         Caption         =   " /"
         Enabled         =   0   'False
         Height          =   275
         Left            =   4320
         TabIndex        =   5
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label BytesBez 
         Caption         =   "&Bytes:"
         Height          =   275
         Left            =   2160
         TabIndex        =   3
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label ZeilenBez 
         Caption         =   "&Zeilen:"
         Height          =   275
         Left            =   120
         TabIndex        =   1
         Top             =   4320
         Width           =   615
      End
   End
   Begin VB.Menu Datei 
      Caption         =   "&Datei"
      Begin VB.Menu EinlesenEinzeln 
         Caption         =   "&Einlesen einzeln mit Alter Table"
      End
      Begin VB.Menu EinlesenSchnell 
         Caption         =   "Einlesen einzeln &ohne Alter Table"
      End
      Begin VB.Menu Einlesen 
         Caption         =   "E&inlesen gesammelt"
         Index           =   2
      End
      Begin VB.Menu EinlesenVorb 
         Caption         =   "Ei&nlesen gesammelt mit Vorladen der Formularfeldinhalte"
      End
      Begin VB.Menu EinlesenEinzelnVorb 
         Caption         =   "Ein&lesen einzeln mit Vorladen der Formularfeldinhalte"
      End
      Begin VB.Menu EinlesLDI 
         Caption         =   "Einles mit Load &Data Infile"
      End
      Begin VB.Menu ÜbertragenNachLDInF 
         Caption         =   "&Übertragen nach Load Data Infile ohne Neueinlesen"
      End
      Begin VB.Menu Zurücksetzen 
         Caption         =   "&Zurücksetzen des Programmlaufs"
         Visible         =   0   'False
      End
      Begin VB.Menu LaborAlle 
         Caption         =   "L&abor alle"
      End
      Begin VB.Menu LaborFehlende 
         Caption         =   "Labor &fehlende"
      End
      Begin VB.Menu Laborquerverbinden 
         Caption         =   "Labor &querverbinden"
      End
      Begin VB.Menu LaborQuerverbindenMitLöschen 
         Caption         =   "Labor q&uerverbinden mit Löschen"
      End
      Begin VB.Menu Abbrechen 
         Caption         =   "Abbre&chen (vor Speichern eines Patienten)"
      End
      Begin VB.Menu Beenden 
         Caption         =   "&Beenden"
         Index           =   3
      End
   End
   Begin VB.Menu Sekundärfunktionen 
      Caption         =   "Sekund&ärfunktionen"
      Begin VB.Menu NachzuholendeLaborimporte 
         Caption         =   "Nach&zuholende Laborimporte"
      End
   End
   Begin VB.Menu Sonderfunktionen 
      Caption         =   "&Sonderfunktionen"
      Begin VB.Menu MachTypen 
         Caption         =   "&MachTypen (Datei Typen.bas erstellen)"
         Index           =   1
      End
      Begin VB.Menu AnamnesebogenHolen 
         Caption         =   "Anamnesebogen von u:\Anamnese\Quelle.mdb &kopieren"
      End
      Begin VB.Menu Vergleichen 
         Caption         =   "&Vergleichen der Datenbanken"
      End
   End
   Begin VB.Menu Testfunktionen 
      Caption         =   "&Testfunktionen"
      Begin VB.Menu t7 
         Caption         =   "t&7"
      End
      Begin VB.Menu ADOXtet 
         Caption         =   "ADO&Xtet"
      End
      Begin VB.Menu SyncTest 
         Caption         =   "&SyncTest"
      End
   End
   Begin VB.Menu Optionen 
      Caption         =   "&Optionen"
      Begin VB.Menu Datenbank 
         Caption         =   "&Datenbank"
      End
   End
End
Attribute VB_Name = "Lese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dlg As Dialog


Private Sub Einlesen_Click(index As Integer)
 Call doEinlesen(True)
End Sub ' Einlesen_Click(Index As Integer)
Private Sub EinlesenSchnell_Click()
 obMitAlterTab = False
 Call doEinlesen(False)
End Sub
Private Sub EinlesenEinzeln_Click()
 obMitAlterTab = True
 Call doEinlesen(False)
End Sub ' EinlesenEinzeln_Click()
Private Sub EinlesenVorb_Click()
 obVorb = True
 Call doEinlesen(True)
End Sub ' EinlesenVorb_Click()
Private Sub EinlesenEinzelnVorb_Click()
 obVorb = True
 Call doEinlesen(False)
End Sub ' EinlesenEinzelnVorb_Click()
Private Sub doEinlesen(obevtlAlle%, Optional LDInF%)
 Dim Fil As File
 On Error GoTo fehler
 AllePat = 0
 Call ProgStart
 If LDInF <> 2 Then
  With Me.CommonDialog1
   .FileName = "*.BDT"
   If LDInF <> 0 Then .InitDir = "\\linux\daten" Else .InitDir = "u:\tmexport"
   .ShowOpen
   If .FileName <> "*.BDT" Then
   If FSO.FileExists(.FileName) Then
    Set Fil = FSO.GetFile(.FileName)
    Me.QDatei = Fil.Path
    Me.QDatum = Fil.DateLastModified
   End If
   End If
  End With
  If Fil.Size > 75000 And obevtlAlle Then AllePat = True
  Me.GesBytes = Format(Fil.Size, "###,###,###,###,###,###,##0")
 Else
 End If
 If LDInF = 2 Or Not Fil Is Nothing Then
    Call EintragStart(Me)
    t1 = Now
    Call GesLies(Me, Me.QDatei, True, t1, LDInF)
    MsgBox CStr((Now - t1) * 60 * 60 * 24) + " s"
 End If
 Call ProgEnde
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doEinlesen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' doEinlesen
Private Sub AnamnesebogenHolen_Click()
 Call ProgStart
 Call holAB(Me)
 Call ProgEnde
End Sub ' AnamnesebogenHolen_Click()

Private Sub EinlesLDI_Click()
' Call ProgStart
 obMitAlterTab = True
 Call doEinlesen(False, True)
' Call LaborDirektImport(Me, -1)
' Call LaborErgPatId(Me)
' Call ProgEnde
End Sub

Private Sub LaborAlle_Click()
 Call ProgStart
 Call LaborDirektImport(Me, -1)
 Call LaborErgPatId(Me)
 Call ProgEnde
End Sub ' LaborAlle_Click()

Private Sub LaborFehlende_Click()
 Call ProgStart
 Call LaborDirektImport(Me, 0)
 Call LaborErgPatId(Me)
 Call ProgEnde
End Sub ' LaborFehlende_Click()

Private Sub Laborquerverbinden_Click()
 Call ProgStart
 Call LaborErgPatId(Me)
 Call ProgEnde
End Sub ' Laborquerverbinden_Click()

Private Sub LaborQuerverbindenMitLöschen_Click()
 Call ProgStart
 Call LaborErgPatId(Me, -1)
 Call ProgEnde
End Sub ' LaborQuerverbindenMitLöschen_Click()

Private Sub MachTypen_Click(index As Integer)
 Call ProgStart
 Call MacheTypen(Me)
 Call ProgEnde
End Sub ' MachTypen_Click(index As Integer)

Private Sub NachzuholendeLaborimporte_Click()
 Dim rs As ADODB.Recordset, lfdnr&
 Const ErgebDatei$ = "u:\anamnese\Nachholen.txt"
 Open ErgebDatei For Output As #301
 Call ProgStart
 Set rs = DBCn.Execute("SELECT distinct pfad,lwerte,laborxus.pat_id,eingang,auftragsnummer, namen.nachname, namen.vorname, laborxus.refnr FROM (laborxus left join laborxeingel on laborxus.datid = laborxeingel.datid) left join namen on laborxus.pat_id = namen.pat_id where afn = 0 and zdip = 0 and not laborxus.pat_id = 0 and zdüp>0 order by refnr;")
 If Not rs.BOF Then
  lfdnr = 1
  Do While Not rs.EOF
   Print #301, Right("   " & lfdnr, 3) & ": Pat: " & Right("   " & rs!Pat_id, 4) & " Auftrag: " & Right("       " & rs!Auftragsnummer, 11) & " " & rs!Eingang & " " & IIf(Len(rs!Pfad) < 50, Right(Space(50) & rs!Pfad, 50), rs!Pfad) & " (Pat: " & rs!Nachname & ", " & rs!Vorname & " Werte: " & Trim(rs!LWerte) & ")"
'   Print #301, lfdnr & ": Pat: " & rs!Pat_id & " (" & rs!Nachname & ", " & rs!Vorname & "), für: " & rs!Eingang & ", Auftragsnummer: " & rs!Auftragsnummer & " (Werte: " & rs!LWerte & ")"
'   Print #301, "   dann importieren: " & rs!Pfad
   lfdnr = lfdnr + 1
   rs.Move 1
  Loop
 End If
 Close #301
 Call Shell("notepad " + ErgebDatei, vbNormalFocus)
 Call ProgEnde
End Sub ' NachzuholendeLaborimporte_Click()

Private Sub t7_Click()
 Call ProgStart
 Call test7
 Call ProgEnde
End Sub

Private Sub ÜbertragenNachLDInF_Click()
 obMitAlterTab = True
 Call doEinlesen(False, 2)
End Sub

Private Sub Vergleichen_Click()
 Call ProgStart
 Call Vergleiche(Me)
 Call ProgEnde
End Sub ' Vergleichen_Click()

Private Sub ADOXtet_Click()
 Call ProgStart
 Call adoxtest(dlg)
 Call ProgEnde
End Sub

Private Sub Abbrechen_Click()
'#If False Then
' MsgBox "Stopp Programm gleich"
' Call ProgrammLauf(-1) ' Programm stoppen
'#End If
 BrichAb = True
End Sub ' Abbrechen_Click


Private Sub Beenden_Click(index As Integer)
'#If False Then
' If ProgrammLauf(-1) Or dbcn.State = 0 Then ' Wenn Programm schon gestoppt war, dann abbrechen, sonst stoppen
'#End If
 If Not ProgLäuft Or DBCn.State = 0 Then
  End
 End If
End Sub ' Beenden_Click(Index As Integer)


Private Sub Beginn_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Bytes_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Ende_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub GesBytes_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub GesDauer_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub


Private Sub nachCd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub obAcc_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub obMy_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Prozent_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Quelle_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Sekunden_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub SyncTest_Click()
 Call formInhMach
End Sub ' SyncTest_Click

Private Sub ÜbertrageCd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Zeilen_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Ziel_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 27 Then End
 Call key(KeyCode, Shift, Me)
End Sub ' Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Ausgabe_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me)
End Sub

#If False Then
Private Sub Zurücksetzen_Click()
 Call ProgrammLauf(-1) ' falls es fälschlich auf 0 steht: 0 = Programm läuft, -1 = nicht
 Dim Ctl As Control
 For Each Ctl In Me.Controls
  If Ctl.Name Like "*inlesen*" Then Ctl.Enabled = True
 Next Ctl
 Me.Zurücksetzen.Enabled = False
End Sub
#End If

Private Sub Datenbank_Click()
 dlg.Show
End Sub ' Datenbank_Click()


Private Sub MDIForm_Load()
 imAufBau1 = True
 Set dlg = New Dialog
 Set dlg.hfrm = Me
 dlg.obAcc = dlg.obAcc
 Me.Zeilen = 0
 Me.Bytes = 0
 With Me.CommonDialog1
  .DialogTitle = "Einzulesende Datei"
  .InitDir = üVerz '"u:\tmexport"
  .FileName = "*.BDT"
  .Orientation = cdlLandscape
 .Flags = 0
' .Flags = .Flags Or FileOpenConstants.cdlOFNExplorer
 .Flags = .Flags Or FileOpenConstants.cdlOFNHideReadOnly ' Schreibgeschützt-Checkbox entfernen
 .Flags = .Flags Or FileOpenConstants.cdlOFNLongNames ' '   Lange Dateinamen erlauben (nur sinnvoll bei Nicht-Win95-Design)
 .Flags = .Flags Or FileOpenConstants.cdlOFNFileMustExist
 .Flags = .Flags Or FileOpenConstants.cdlOFNPathMustExist
 End With
 Call MDIForm_Resize
 imAufBau1 = False
 Dim i%
 For i = 0 To Me.Controls.Count - 1
  If Me.Controls(i).Name = "Abbrechen" Then
   Me.Controls(i).Enabled = False
   Exit For
  End If
 Next i
End Sub ' MDIForm_Load()
Function ProgStart()
 Dim i&
 BrichAb = 0
 t1 = Now
 Me.Controls!Abbrechen.Enabled = True
' Me.Controls!Beenden.Enabled = False
 For i = 1 To Me.Controls.Count
  If Me.Controls(i).Name = "Beenden" Then
   Me.Controls(i).Enabled = False
   Exit For
  End If
 Next i
 ProgLäuft = True
 Call ConstrFestleg(0, dlg)
End Function
Function ProgEnde()
 Dim i&
 Me.Controls!Abbrechen.Enabled = False
' Me.Controls!Beenden.Enabled = True
 For i = 1 To Me.Controls.Count
  If Me.Controls(i).Name = "Beenden" Then
   Me.Controls(i).Enabled = True
   Exit For
  End If
 Next i
 ProgLäuft = False
End Function ' ProgEnde
Private Sub MDIForm_Resize()
 Picture1.Height = Me.Height - 5745 + 5055
 Ausgabe.Height = Me.Height - 5745 + 4250
 Me.ÜbertrageCd.Top = Me.Ausgabe.Top + Me.Ausgabe.Height + 40
 Me.QDatei.Top = Me.ÜbertrageCd.Top
 Me.QDatum.Top = Me.ÜbertrageCd.Top
 Me.nachCd.Top = Me.ÜbertrageCd.Top
 Me.obMy.Top = Me.ÜbertrageCd.Top
 Me.obAcc.Top = Me.ÜbertrageCd.Top
 Me.Ziel.Top = Me.ÜbertrageCd.Top
 Me.ZeilenBez.Top = Me.Height - 5745 + 4645
 Me.Zeilen.Top = Me.ZeilenBez.Top
 Me.BytesBez.Top = Me.ZeilenBez.Top + 20
 Me.Bytes.Top = Me.ZeilenBez.Top
 Me.GesBytesBez.Top = Me.ZeilenBez.Top + 20
 Me.GesBytes.Top = Me.ZeilenBez.Top
 Me.GleichBez.Top = Me.ZeilenBez.Top + 20
 Me.Prozent.Top = Me.ZeilenBez.Top
 Me.ProzentBez.Top = Me.ZeilenBez.Top + 20
 Me.Sekunden.Top = Me.ZeilenBez.Top
 Me.sbisherBez.Top = Me.ZeilenBez.Top + 20
 Me.Beginn.Top = Me.ZeilenBez.Top
 Me.EndeBez.Top = Me.ZeilenBez.Top + 20
 Me.Ende.Top = Me.ZeilenBez.Top
 Me.SBez.Top = Me.Ende.Top
 Me.Picture2.Top = Me.Ende.Top
 Me.GesamtDauerBez.Top = Me.ZeilenBez.Top + 20
 Me.GesDauer.Top = Me.ZeilenBez.Top
End Sub ' MDIForm_Resize()

Private Sub MDIForm_Unload(Cancel As Integer)
 End
End Sub ' MDIForm_Unload(Cancel As Integer)

