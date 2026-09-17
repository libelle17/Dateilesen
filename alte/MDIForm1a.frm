VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm Lese 
   BackColor       =   &H8000000C&
   Caption         =   "Turbomeddaten nach MySQL übertragen"
   ClientHeight    =   5055
   ClientLeft      =   1290
   ClientTop       =   1335
   ClientWidth     =   15600
   Icon            =   "MDIForm1a.frx":0000
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
      Begin MSComDlg.CommonDialog CommonDialogLese 
         Left            =   14040
         Top             =   4680
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
      Begin VB.Menu Anamnesebogen 
         Caption         =   "&Anamnesebogen"
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
   Begin VB.Menu Anwendungsfunktionen 
      Caption         =   "An&wendungsfunktionen"
      Begin VB.Menu DMPSend 
         Caption         =   "&DMPSend"
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
      Begin VB.Menu AnamnesebogenPacken 
         Caption         =   "Anamnesebogen &packen (Stringfeldlängen optimieren)"
      End
      Begin VB.Menu MachTypen 
         Caption         =   "&MachTypen (Datei Typen.bas erstellen)"
         Index           =   1
      End
      Begin VB.Menu AnamnesebogenHolen 
         Caption         =   "Anamnesebogen von u:\Anamnese\Quelle.mdb &kopieren"
      End
      Begin VB.Menu DokumenteAbgehaktkopieren 
         Caption         =   "&Dokumente abgehakt von u:\Anamnese\Quelle.mdb kopieren"
      End
      Begin VB.Menu KassenlisteKopieren 
         Caption         =   "&Kassenliste von u:\Anamnese\Quelle.mdb kopieren"
      End
      Begin VB.Menu DokumentPfadeKorrigieren 
         Caption         =   "&Dokumentpfade korrigieren"
      End
      Begin VB.Menu Vergleichen 
         Caption         =   "&Vergleichen der Datenbanken"
      End
      Begin VB.Menu BooleanFelder 
         Caption         =   "&Boolean-Felder in MySQL-Datenbanken erstellen"
      End
   End
   Begin VB.Menu Testfunktionen 
      Caption         =   "&Testfunktionen"
      Begin VB.Menu Spinstr 
         Caption         =   "S&pinstr"
      End
      Begin VB.Menu t7 
         Caption         =   "t&7"
      End
      Begin VB.Menu Eintragszusatz 
         Caption         =   "Eintrags&zusatz"
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
         Caption         =   "&Einlesen"
      End
      Begin VB.Menu Sonstige 
         Caption         =   "&Sonstige"
      End
   End
End
Attribute VB_Name = "Lese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dlg As New Dialog
Public snst As New Sonstige
Public anb As New Anamnese
'Public EinlAb&

Private Sub Anamnesebogen_Click()
 anb.Show
End Sub

'Private Sub Einlesen_Click(index As Integer)
' obMitAlterTab = True
' Call ProgStart
' Call doEinlesen(True)
' Call ProgEnde
'End Sub ' Einlesen_Click(Index As Integer)

'Private Sub EinlesenAb_Click()
' Dim erg
' obMitAlterTab = True
' Call ProgStart
' erg = InputBox("Ab welcher Patientennummer soll eingelesen werden?", "Einlesen ab Patientennummer")
' If IsNumeric(erg) Then
'  EinlAb = erg
'  Call doEinlesen(False)
'  Call ProgEnde
' End If
'End Sub

'Private Sub EinlesenEingelesene_Click()
' obMitAlterTab = True
' Call ProgStart
' Call DBCn.Execute("update quelle.namen set aktzeit = 0")
' Call doEinlesen(False)
' Call ProgEnde
'End Sub ' EinlesenEingelesene_Click()

'Private Sub EinlesenohneLabor_Click()
' obMitAlterTab = True
' Call ProgStart
' Call doEinlesen(False, True)
' Call ProgEnde
'End Sub ' EinlesenohneLabor_Click()

'Private Sub EinlesenSchnell_Click()
' obMitAlterTab = False
' Call ProgStart
' Call doEinlesen(False)
' Call ProgEnde
'End Sub ' EinlesenSchnell_Click()
'Private Sub EinlesenEinzeln_Click()
' obMitAlterTab = True
' Call ProgStart
' Call doEinlesen(False)
' Call ProgEnde
'End Sub ' EinlesenEinzeln_Click()
'Private Sub EinlesenVorb_Click()
' obMitAlterTab = True
' obVorb = True
' Call ProgStart
' Call doEinlesen(True)
' Call ProgEnde
'End Sub ' EinlesenVorb_Click()
'Private Sub EinlesenEinzelnVorb_Click()
' obMitAlterTab = True
' obVorb = True
' Call ProgStart
' Call doEinlesen(False)
' Call ProgEnde
'End Sub ' EinlesenEinzelnVorb_Click()
Private Sub AnamnesebogenHolen_Click()
 Call ProgStart
 Call holAB(Me)
 Call ProgEnde
End Sub ' AnamnesebogenHolen_Click()

Private Sub DMPSend_Click()
 Call ProgStart
 Call alleDMPs(Me)
 Call ProgEnde
End Sub
Private Sub KassenlisteKopieren_Click()
 Call ProgStart
 Call holKassen(Me)
 Call ProgEnde
End Sub

Private Sub DokumenteAbgehaktkopieren_Click()
 Call ProgStart
 Call holDA(Me)
 Call ProgEnde
End Sub
Private Sub DokumentPfadeKorrigieren_Click()
 Call ProgStart
 Call DokPfadKorr(Me)
 Call ProgEnde
End Sub

Private Sub AnamnesebogenPacken_Click()
 Call ProgStart
 Call AnPack
 Call ProgEnde
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

Private Sub MDIForm_Terminate()
 End
End Sub

Private Sub NachzuholendeLaborimporte_Click()
 Dim rs As ADODB.Recordset, lfdnr&
 Open snst.DateiNachzuholen For Output As #301
 Call ProgStart
 Set rs = DBCn.Execute("SELECT distinct pfad,lwerte,laborxus.pat_id,eingang,auftragsnummer, namen.nachname, namen.vorname, laborxus.refnr FROM (laborxus left join laborxeingel on laborxus.datid = laborxeingel.datid) left join namen on laborxus.pat_id = namen.pat_id where afn = 0 and zdip = 0 and not laborxus.pat_id = 0 and zdüp>0 order by refnr;")
 If Not rs.BOF Then
  lfdnr = 1
  Do While Not rs.EOF
   Print #301, Right("   " & lfdnr, 3) & ": Pat: " & Right("   " & rs!Pat_ID, 4) & " Auftrag: " & Right("       " & rs!Auftragsnummer, 11) & " " & rs!Eingang & " " & IIf(Len(rs!Pfad) < 50, Right(Space(50) & rs!Pfad, 50), rs!Pfad) & " (Pat: " & rs!Nachname & ", " & rs!Vorname & " Werte: " & Trim(rs!LWerte) & ")"
'   Print #301, lfdnr & ": Pat: " & rs!Pat_id & " (" & rs!Nachname & ", " & rs!Vorname & "), für: " & rs!Eingang & ", Auftragsnummer: " & rs!Auftragsnummer & " (Werte: " & rs!LWerte & ")"
'   Print #301, "   dann importieren: " & rs!Pfad
   lfdnr = lfdnr + 1
   rs.Move 1
  Loop
 End If
 Close #301
 Call Shell("notepad " + snst.DateiNachzuholen, vbNormalFocus)
 Call ProgEnde
End Sub ' NachzuholendeLaborimporte_Click()

Private Sub Sonstige_Click()
 snst.Show
End Sub

Private Sub Spinstr_Click()
 Dim St1$, St2$
 St1 = "noch warst du n.och"
 St2 = "noch"
 Call doSp(St1, St2)
End Sub
Private Sub t7_Click()
 Call ProgStart
' Call test7
' Call löschBezügeausLaborux(65)
 Call LöschDateiEintrag(48)
 Call ProgEnde
End Sub
Private Sub Eintragszusatz_Click()
 Call ProgStart
 Call EintragZusatz
 Call ProgEnde
End Sub

Private Sub Vergleichen_Click()
 Call ProgStart
 Call Vergleiche(Me)
 Call ProgEnde
End Sub ' Vergleichen_Click()
Private Sub BooleanFelder_Click()
 Call ProgStart
 Call BooleanFld(Me)
 Call ProgEnde
End Sub

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
Public Function PutEinstAufDB()
 If 1 = 0 Then
 If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
 If Not DBCn Is Nothing Then
  If DBCn.State = 1 Then
   Set rs = New ADODB.Recordset
   rs.Open "select * from eintragszahlen where beginn = (select max(beginn) from eintragszahlen)", DBCn, adOpenDynamic, adLockOptimistic ' "select * from eintragszahlen order by beginn desc" soll bei älteren MySQL-Versionen nicht immer ganz funktionieren
   rs!TabellenLöschen = dlg.TabellenLöschen
   rs!ZurücksetzenLAktDat = dlg.ZurücksetzenLAktDat
   rs!obVglMitLetzterEinlesung = dlg.obVglMitLetzterEinlesung
   rs!Pat_IDBis = IIf(IsNumeric(dlg.Pat_IDBis), dlg.Pat_IDBis, 0)
   rs!Pat_IDBis = IIf(IsNumeric(dlg.Pat_IDBis), dlg.Pat_IDBis, 0)
   rs!AlterTab = dlg.AlterTab
   rs!VorladenFFI = dlg.VorladenFFI
   rs!LaborDirektEinlesen = dlg.LaborDirektEinlesen
   rs!LaborDirektNeu = dlg.LaborDirektNeu
   rs!LaborQuerVerb = dlg.LaborQuerVerb
   rs!LaborQuerNeu = dlg.LaborQuerNeu
   rs!LaborPfadBeispiel = dlg.LaborPfadBeispiel
'   rs!Datei = dlg.BDTDatei
   rs.Update
   End If
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
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PutEinstAufDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PutEinstAufDB
Public Function HolEinstvonDB()
 On Error GoTo fehler
 If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
 If Not DBCn Is Nothing Then
  If DBCn.State = 1 Then
   On Error Resume Next
   Set rs = DBCn.Execute("select * from eintragszahlen where beginn = (select max(beginn) from eintragszahlen)")
   If Err.Number <> 0 Then
    MsgBox "Fehler bei Tabelle 'Eintragszahlen' in Datenbank '" & DBCn.Properties("Data Source Name") & "'"
    Exit Function
   End If
   On Error GoTo fehler
   dlg.TabellenLöschen = Abs(HolEinstFeld(rs, "TabellenLöschen", 0, "BIT"))
   dlg.TabellenLöschen = 0
   dlg.ZurücksetzenLAktDat = Abs(HolEinstFeld(rs, "ZurücksetzenLAktDat", 0, "BIT"))
   dlg.obVglMitLetzterEinlesung = Abs(HolEinstFeld(rs, "obVglMitLetzterEinlesung", 0, "BIT"))
   dlg.ZurücksetzenLAktDat = 0
   dlg.Pat_IDVon = HolEinstFeld(rs, "Pat_IDvon", 0, "TEXT", 6)
   dlg.Pat_IDBis = HolEinstFeld(rs, "Pat_IDbis", 0, "TEXT", 6)
   dlg.AlterTab = Abs(HolEinstFeld(rs, "AlterTab", -1, "BIT"))
   dlg.VorladenFFI = Abs(HolEinstFeld(rs, "VorladenFFI", 0, "BIT"))
   dlg.LaborDirektEinlesen = Abs(HolEinstFeld(rs, "LaborDirektEinlesen", 0, "BIT"))
   dlg.LaborDirektNeu = Abs(HolEinstFeld(rs, "LaborDirektNeu", 0, "BIT"))
   dlg.LaborQuerVerb = Abs(HolEinstFeld(rs, "LaborQuerVerb", -1, "BIT"))
   dlg.LaborQuerNeu = Abs(HolEinstFeld(rs, "LaborQuerNeu", 0, "BIT"))
   dlg.BDTDatei = HolEinstFeld(rs, "Datei", üVerz, "TEXT", 120)
   dlg.BDTDatei = getLDatei(dlg.BDTDatei, "*.bdt")
   dlg.LaborPfadBeispiel = HolEinstFeld(rs, "LaborPfadBeispiel", LabTransPfad, "TEXT")
   dlg.obmitEmails = Abs(HolEinstFeld(rs, "obmitEmails", -1, "BIT"))
   dlg.obBDT = Abs(True)
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
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HolEinstVonDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Private Function HolEinstFeld(rs As ADODB.Recordset, FName$, Default, ByVal Typ$, Optional Lenge&)
 Dim FNr&
 On Error Resume Next
 HolEinstFeld = rs(FName)
 FNr = Err.Number
 If FNr <> 0 Then
'  On Error GoTo 0
  rs.Close
  If obMySQL Then
   Select Case Typ
    Case "TEXT": Typ = "VARCHAR": If Lenge = 0 Then Typ = "VARCHAR(300)"
    Case "BIT": If Lenge = 0 Then Typ = "BIT(1)" '  Typ = "TINYINT": If Lenge = 0 Then Typ = "TINYINT(1)"
    Case "DATE", "TIME": Typ = "DATETIME"
    ' identisch verwendbar sind evtl.: "LONGTEXT","DECIMAL","FLOAT","BINARY"
   End Select
  End If
  Call DBCn.Execute("Alter Table " & kla & "eintragszahlen" & klz & " ADD Column " & kla & FName & klz & " " & Typ & IIf(Lenge <> 0, "(" & Lenge & ")", "") & IIf(Null, " NULL", ""))
  Set rs = DBCn.Execute("select * from eintragszahlen where beginn = (select max(beginn) from eintragszahlen)")
  HolEinstFeld = Default
 End If
  Select Case VarType(HolEinstFeld)
   Case Is < 8
   Case Else
    If HolEinstFeld = "" Then HolEinstFeld = Default
  End Select
 On Error Resume Next
 If InStr(HolEinstFeld, ":\\") > 0 Or InStr(HolEinstFeld, "\\\\") > 0 Then HolEinstFeld = Replace(HolEinstFeld, "\\", "\")
 On Error GoTo 0
 If IsNull(HolEinstFeld) Then HolEinstFeld = 0
 On Error GoTo 0
End Function ' HolEinstFeld(rs As ADODB.Recordset, FName$, Default, ByVal Typ$, Optional Lenge&)

Private Sub MDIForm_Load()
 imAufbauLese = True
 Set dlg = New Dialog
 Set dlg.hlese = Me
 Set snst = New Sonstige
 Set snst.hlese = Me
 Me.Zeilen = 0
 Me.Bytes = 0
 With Me.CommonDialogLese
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
 imAufbauLese = False
 Call AbbrechDisable(Me)
End Sub ' MDIForm_Load()
Public Function AbbrechDisable(frm As Lese)
 Dim i%
 For i = 0 To frm.Controls.Count - 1
  If frm.Controls(i).Name = "Abbrechen" Then
   frm.Controls(i).Enabled = False
   Exit For
  End If
 Next i
End Function

Public Function ProgStart()
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
Public Function ProgEnde()
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

Private Sub MDIForm_Unload(Cancel As Integer) ' geht nur beim Anklicken des Kreuzes oben
 End
End Sub ' MDIForm_Unload(Cancel As Integer)

Public Function ZeigDateien()
 Me.obAcc = Me.dlg.obAcc
 Me.obMy = Not Me.dlg.obAcc
 If Me.dlg.obAcc Then
  Me.Ziel = Me.dlg.MdB
 Else
  If Me.dlg.obMyQuelle Then
   Me.Ziel = "quelle"
  ElseIf Me.dlg.obMyQuelle1 Then
   Me.Ziel = "quelle1"
  ElseIf Me.dlg.obMyQuelle2 Then
   Me.Ziel = "quelle2"
  End If
 End If
End Function ' zeigdateien


