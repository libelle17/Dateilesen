VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm Lese 
   BackColor       =   &H8000000C&
   Caption         =   "Patientendaten Diabetespraxis Dachau"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   990
   ClientWidth     =   19080
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":030A
   Begin VB.PictureBox Picture1 
      Align           =   1  'Oben ausrichten
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   19020
      TabIndex        =   0
      Top             =   0
      Width           =   19080
      Begin VB.TextBox DurchschnDauer 
         BackColor       =   &H80000018&
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   17760
         TabIndex        =   31
         Top             =   4320
         Width           =   1215
      End
      Begin VB.ComboBox MyDB 
         Height          =   315
         ItemData        =   "MDIForm1.frx":0614
         Left            =   8040
         List            =   "MDIForm1.frx":0616
         TabIndex        =   21
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox Fuß 
         BackColor       =   &H80000001&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   28
         Top             =   4680
         Width           =   19095
      End
      Begin VB.CommandButton DMPForts 
         Caption         =   "D&MP Fortsetzen"
         Height          =   375
         Left            =   17520
         TabIndex        =   27
         Top             =   3960
         Width           =   1455
      End
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
         BackColor       =   &H00FFFFFF&
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
      Begin VB.OptionButton obAcc 
         Caption         =   "&Access"
         Height          =   275
         Left            =   9720
         TabIndex        =   22
         Top             =   3960
         Width           =   855
      End
      Begin VB.OptionButton obMySQL 
         Caption         =   "M&ySQL"
         Height          =   275
         Left            =   7080
         TabIndex        =   20
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
         Width           =   19035
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
      Begin VB.Label DurschnDauerLab 
         Caption         =   "DurschnittsDauer[s]:"
         Height          =   255
         Left            =   16320
         TabIndex        =   32
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Version 
         Height          =   255
         Left            =   15600
         TabIndex        =   30
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label ConStri 
         Caption         =   "hier könnte der Connection-String stehen"
         Height          =   975
         Left            =   0
         TabIndex        =   29
         Top             =   4920
         Width           =   18975
         WordWrap        =   -1  'True
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
      Begin VB.Menu ÜbertragenenAnamnesebogen 
         Caption         =   "&Übertragenen Anamnesebogen laden"
      End
      Begin VB.Menu Anamnesebogen 
         Caption         =   "&Anamnesebogen experimentell"
         Enabled         =   0   'False
      End
      Begin VB.Menu LaborEintragen 
         Caption         =   "Labor &eintragen"
      End
      Begin VB.Menu Zurücksetzen 
         Caption         =   "&Zurücksetzen des Programmlaufs"
         Visible         =   0   'False
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
         Caption         =   "Alle &DMP-Dokumente an Hausärzte in p:\zufaxen erstellen"
      End
      Begin VB.Menu DMPhierListe 
         Caption         =   "D&MP hier Liste"
      End
      Begin VB.Menu DuplexKontrollieren 
         Caption         =   "&Duplex Kontrollieren"
      End
      Begin VB.Menu DMPLeiDok 
         Caption         =   "DMP-Brief-&Leistungs-Doku"
      End
      Begin VB.Menu BriefeLeiDok 
         Caption         =   "&Leistungen zu Arztbriefen bei Kassenpatienten in u:\tmimport\ dokumentieren"
      End
      Begin VB.Menu VorhandeneBriefe 
         Caption         =   "&Vorhandene Briefe korrigieren"
      End
      Begin VB.Menu BriefeBerichtspflicht 
         Caption         =   "&Briefe zu Patienten mit Berichtspflicht schreiben"
      End
      Begin VB.Menu UngeschriebeneBriefe 
         Caption         =   "&Briefe zu Pat., zu denen bisher kein Brief geschrieben wurde"
      End
      Begin VB.Menu doFollowUp 
         Caption         =   "&Briefe zu Patienten, deren erster Brief schon > 1a her ist"
      End
      Begin VB.Menu BriefeZuListeSchreiben 
         Caption         =   "Briefe zu &Pat_ID-Liste aus Datei schreiben"
      End
      Begin VB.Menu RestlicheBriefe 
         Caption         =   "&Restliche Briefe für aktuelles Quartal schreiben"
      End
      Begin VB.Menu BriefSchreiben 
         Caption         =   "&Brief schreiben"
         Shortcut        =   ^B
      End
      Begin VB.Menu BriefOhneMaske 
         Caption         =   "Brief &ohne Maske schreiben"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Sekundärfunktionen 
      Caption         =   "Sekund&ärfunktionen"
      Begin VB.Menu ListeDerFehlendenDokumente 
         Caption         =   "&Liste der fehlenden Dokumente"
      End
      Begin VB.Menu NachzuholendeLaborimporte 
         Caption         =   "Nach&zuholende Laborimporte"
      End
      Begin VB.Menu FehlendeHausaerzte 
         Caption         =   "Liste der fehlenden &Hausärzte"
      End
      Begin VB.Menu Statistiken 
         Caption         =   "&Statistiken"
         Begin VB.Menu Hilfsmittelklassifikationen 
            Caption         =   "&Hilfsmittelklassifikationen"
         End
         Begin VB.Menu Fallzahlstand 
            Caption         =   "&Fallzahlstand"
         End
      End
      Begin VB.Menu Patientenlaufzettel 
         Caption         =   "&Patientenlaufzettel"
      End
      Begin VB.Menu PLZeinzeln 
         Caption         =   "P&atientenlaufzettel einzeln"
      End
   End
   Begin VB.Menu Sonderfunktionen 
      Caption         =   "&Sonderfunktionen"
      Begin VB.Menu AnamnesebogenPacken 
         Caption         =   "Anamnesebogen pa&cken (Stringfeldlängen optimieren)"
      End
      Begin VB.Menu MachTypen 
         Caption         =   "&MachTypen (Datei Typen.bas erstellen)"
         Index           =   1
      End
      Begin VB.Menu AnamnesebogenHolen 
         Caption         =   "Anamnesebogen von u:\Anamnese\Quelle.mdb &kopieren"
         Visible         =   0   'False
      End
      Begin VB.Menu DokumenteAbgehaktkopieren 
         Caption         =   "&Dokumente abgehakt von u:\Anamnese\Quelle.mdb kopieren"
         Visible         =   0   'False
      End
      Begin VB.Menu KassenlisteKopieren 
         Caption         =   "&Kassenliste von u:\Anamnese\Quelle.mdb kopieren"
         Visible         =   0   'False
      End
      Begin VB.Menu holMedArten 
         Caption         =   "&Medarten von u:\Anamnese\Quelle.mdb kopieren"
         Visible         =   0   'False
      End
      Begin VB.Menu holLaborParameter 
         Caption         =   "&Laborparameter von u:\Anamnese\Quelle.mdb kopieren"
         Visible         =   0   'False
      End
      Begin VB.Menu TabKop 
         Caption         =   "&Tabelle kopieren"
      End
      Begin VB.Menu MedartenEditieren 
         Caption         =   "&Medarten editieren"
      End
      Begin VB.Menu WSt0Erg 
         Caption         =   "&Fußsyndrome Wagner-Stadium 0 ergänzen"
      End
      Begin VB.Menu holEBM2000plus 
         Caption         =   "&EBM2000plus-Ziffern holen"
      End
      Begin VB.Menu TherapieartenFestlegen 
         Caption         =   "&Therapiearten festlegen"
      End
      Begin VB.Menu HAUebertrag 
         Caption         =   "Haus&ärzte übertragen"
      End
      Begin VB.Menu KassenkategorienBestimmen 
         Caption         =   "&Kassenkategorien bestimmen"
      End
      Begin VB.Menu DokPfadKorrigieren 
         Caption         =   "&Dokumentpfade korrigieren"
      End
      Begin VB.Menu Vergleichen 
         Caption         =   "Ve&rgleich der Datenbankstrukturen"
      End
      Begin VB.Menu TabVergleich 
         Caption         =   "Vergleich der Datenbanktabellenf&üllungen"
      End
      Begin VB.Menu Excelliste 
         Caption         =   "E&xcelliste"
      End
      Begin VB.Menu DMPListe 
         Caption         =   "DM&P-Liste erstellen"
      End
      Begin VB.Menu CallDMPString 
         Caption         =   "DMP-St&ring"
      End
      Begin VB.Menu DiagnosenSortieren 
         Caption         =   "D&iagnosen sortieren"
      End
      Begin VB.Menu DiagnosenExportieren 
         Caption         =   "&Diagnosen exportieren"
      End
      Begin VB.Menu DiagnosenExportierenTest 
         Caption         =   "&Diagnosen exportieren (test)"
      End
      Begin VB.Menu FormulareÜbertragen 
         Caption         =   "&Formulare übertragen"
      End
      Begin VB.Menu BooleanFelder 
         Caption         =   "&Boolean-Felder in MySQL-Datenbanken erstellen"
      End
      Begin VB.Menu ViewsErstellen 
         Caption         =   "&ViewsErstellen"
      End
      Begin VB.Menu Leistungseingabe 
         Caption         =   "&Leistungseingabe"
         Begin VB.Menu Porto 
            Caption         =   "&Porto"
         End
      End
      Begin VB.Menu WiedereinbestellungenDMP 
         Caption         =   "&Wiedereinbestellungen DMP"
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
      Begin VB.Menu CallUSDM 
         Caption         =   "&Usdm"
      End
      Begin VB.Menu gefaxttrennen 
         Caption         =   "gefaxttrennen"
      End
      Begin VB.Menu doppelteFAxe 
         Caption         =   "&doppelteFAxe"
      End
      Begin VB.Menu alteHausärzte 
         Caption         =   "alteHausärzte"
      End
      Begin VB.Menu MedKlassT 
         Caption         =   "&MedKlass"
      End
      Begin VB.Menu test 
         Caption         =   "&test"
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
      Begin VB.Menu Datenbankverbindung 
         Caption         =   "&Datenbankverbindung Patientendaten"
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
Public pau As New PatAuswahl
Public ple As New LabEintr
Public pal As New PatListe
Public Tü As New TabÜbertr
'Public ple2 As New LabEin2
'Public ple3 As New LabEin3
#Const mitab = True ' auch in Formular
#If mitab Then
Public anBogÜ As New AnBog
Public labtest As New LaborTest
#End If
'Public DMPlst As New DMPListen
Public obAusgedehnt% ' ob Accessdatenbank noch komprimiert werden muss
Enum AktionTyp
 nix
 BriefSchreiben
 RestlicheBriefe
End Enum
Public Aktion As AktionTyp
Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1
'Public WithEvents qdb As QuelleDBC

'Public EinlAb&
Private Sub TestFos_Click()
 Dim rs As New ADODB.Recordset
 Call ProgStart
 Call rs.Open("select medikament, Urä, Ostp from medarten where medikament = 'FOSAVANCE'", DBCn, adOpenStatic, adLockReadOnly)
 Debug.Print rs!Medikament, rs!urä, rs!ostp
 Call ProgEnde
End Sub

Private Sub alteHausärzte_Click()
 Dim ahc As New ADODB.Connection, ahr As New ADODB.Recordset, acr As New ADODB.Recordset
 Dim runde%
' Set ahc = acon(haT, accDtb)
 Call acon(HaT, q1Dtb)
 
 Dim Fls As Files, Fl As File, lastDat#
 Dim FSO As New FileSystemObject
    Set Fls = FSO.GetFolder(AnamneseVerZeichnis1).Files
    For Each Fl In Fls
     If Fl.name Like "KV*rzte*.mdb" Then
      Set ahc = Nothing
      ahc.Open CStrAcc & Fl.path
      
 For runde = 1 To 2
 Set ahr = Nothing
 On Error Resume Next
 ahr.Open "select * from " & IIf(runde = 1, "hae", "haealt") & " where gelöscht", ahc, adOpenDynamic, adLockReadOnly
 If Err.Number <> 0 Then GoTo nrunde
 On Error GoTo 0
 Do While Not ahr.EOF
  Set acr = Nothing
  acr.Open "select * from hae where nachname = '" & ahr!Nachname & "' and kvnr = '" & ahr!KVNr & "'", HAECn, adOpenStatic, adLockReadOnly
  If acr.EOF Then
   Set acr = Nothing
   acr.Open "select * from haealt where nachname = '" & ahr!Nachname & "' and kvnr = '" & ahr!KVNr & "'", HAECn, adOpenStatic, adLockReadOnly
   If acr.EOF Then
    Stop
   End If
  End If
  ahr.Move 1
 Loop
nrunde:
 On Error GoTo 0
 Next runde
      
     End If
    Next Fl
 
End Sub

Private Sub Anamnesebogen_Click()
 anb.Show
End Sub

Private Sub BriefeBerichtspflicht_Click()
 Call doBriefeBerichtspflicht
End Sub


Private Sub BriefeZuListeSchreiben_Click()
 Dim bzl As New Lade
 bzl.FenArt = 1
 bzl.Show
End Sub

Private Sub BriefOhneMaske_Click()
 Dim erg$
 erg = InputBox("Bitte Pat_ID eingeben:")
 If IsNumeric(erg) Then
  Call tubriefStandalone(CLng(erg), 0)
 End If
End Sub

Private Sub Datenbankverbindung_Click()
' Call dbv.rücksetzBedTbl
' Call dbv.setzBedTbl("anamnesebogen")
' dbv.Show 1
' Set cn = Nothing
' cn.Open dbv.CnStr
 Call dbv.Auswahl("", "anamnesebogen", "Patientendaten")
 obStart = True
 If InStrB(UCase$(dbv.CnStr), "MYSQL") > 0 Then
  Me.obMySQL = True
  Me.obAcc = False
 Else
  Me.obMySQL = False
  Me.obAcc = True
 End If
 obStart = False
' Unload Me
' Me.Show
End Sub

Private Sub DiagnosenExportieren_Click()
 Call GesDiagExp
End Sub

Private Sub DiagnosenExportierenTest_Click()
 Call GesDiagExp(obTest:=True)
End Sub

Private Sub DiagnosenSortieren_Click()
 Dim pad As New PatListe
 pad.art = Diag
 Set pad.hlese = Me
 pad.Show
End Sub

Private Sub DMPhierListe_Click()
 Call ProgStart
 Set pal.hlese = Me
 Me.Hide
 pal.Show
 Call ProgEnde
End Sub ' DMPhierListe_Click()

Private Sub DokPfadKorrigieren_Click() ' Dokumentpfade korrigieren
 Call ProgStart
 Call dokpfadänder(Me)
 Call ProgEnde
End Sub

Private Sub doppelteFAxe_Click()
 Dim DatNam$
 DatNam = "p:\schongefaxte " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".txt"
 Open DatNam For Output As #323
 Call acon(quelleT)
 Call acon(FaxT)
 Call dodoppeltefaxe("p:\unkorrigiert")
 Print #323, "Fertig!"
 Debug.Print "Fertig!"
 Close #323
 Shell Environ("systemroot") & "\system32\notepad.exe " & DatNam, vbMaximizedFocus
End Sub ' doppeltefaxe
Private Sub dodoppeltefaxe(V$)
 Static FSO As New FileSystemObject
 Dim Fil As File, pid$, pos%, buch$
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
 Print #323, V
 Debug.Print V
 For Each Fil In FSO.GetFolder(V).Files
  If Fil.name Like "*PID *" Then
   pos = InStr(Fil.name, "PID ") + 4
   pid = ""
   Do
    buch = Mid$(Fil.name, pos, 1)
    If IsNumeric(buch) Then
     pid = pid & buch
    Else
     Exit Do
    End If
    pos = pos + 1
   Loop
   If pid <> "" Then
    pid = pid & ","
    Set rs = Nothing
    Call rs.Open("select * from outg where docn like '%PID " & pid & "%' And docn Like '%Arztbrief%'", FxCn, adOpenStatic, adLockReadOnly)
    If Not rs.EOF Then
     If FSO.FileExists("p:\" & rs!docn) Then
      Print #323, "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe
      Debug.Print "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe
      Kill Fil.path
     Else
      Set rs1 = Nothing
      Call rs1.Open("select * from briefe where name like '%Arztbrief%' and name like '%PID " & pid & "%'", DBCn, adOpenStatic, adLockReadOnly)
      If Not rs1.EOF Then
       Print #323, "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe, rs1!name
       Debug.Print "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe, rs1!name
       Kill Fil.path
      Else
       Print #323, rs!docn, rs!origst, rs!submt, rs!transe
       Debug.Print rs!docn, rs!origst, rs!submt, rs!transe
      End If
     End If
    End If
   End If
  End If
 Next Fil
 Dim fld As Folder
 For Each fld In FSO.GetFolder(V).SubFolders
  Call dodoppeltefaxe(fld.path)
 Next fld
End Sub ' dodoppeltefaxe

Private Sub DuplexKontrollieren_Click()
 Call ProgStart
 Call doDuplexkontrollieren
 Call ProgEnde
End Sub
#If False Then
Private Sub falschebriefelöschen_Click()
 Dim Fil As File, pid&, pos&, p2&, rs As New ADODB.Recordset, an As New ADODB.Recordset, Infos$(), rd As New ADODB.Recordset
 Call ProgStart
 For Each Fil In FSO.GetFolder("p:\unkorrigiert").Files
  pos = InStr(Fil.name, "PID ")
  If pos > 0 Then
   p2 = InStr(pos, Fil.name, " ")
   pid = Mid$(Fil.name, pos + 4, p2 - pos + 1)
   Set rs = Nothing
   Call rs.Open("select * from faelle where pat_id = " & pid & " order by bhfb desc;", DBCn, adOpenStatic, adLockOptimistic)
   
   Set an = Nothing
'   Call an.Open("select -tkz as j_tkz, a.* from anamnesebogen a where pat_id = " & pid, DBCn, adOpenStatic, adLockOptimistic)
   Call an.Open("select * from anamnesebogen a where pat_id = " & pid, DBCn, adOpenStatic, adLockOptimistic)
   If rs!SchGr = 41 Or rs!SchGr = 42 Or rs!SchGr = 43 Then
    FSO.DeleteFile (Fil.path)
'   ElseIf an!j_tkz <> 0 Then
   ElseIf an!Tkz <> 0 Then
'    Stop
    FSO.DeleteFile (Fil.path)
   Else
     Call getHausarzt(pid, Infos())
     If Infos(1, 0) = "" Then
      FSO.DeleteFile (Fil.path)
     ElseIf Infos(1, 0) Like "*Schade" Then
      Debug.Print Fil.name
      Stop
     End If
    Set rd = Nothing
    Call rd.Open("Select * from briefe where pat_id = " & pid & " and name like 'Brief an %'", DBCn, adOpenStatic, adLockOptimistic)
    If Not rd.BOF Then
     FSO.DeleteFile (Fil.path)
    End If
'    Set rd = Nothing
'    Call rd.Open("Select * from briefe where pat_id = " & pid & " and name not like '%labor%' and name not like '%Tagebuch%' and name not like '%Nachricht an%' and name not like '% BZ%' and name not like '%DMP-%' and name not like '%EKG%' and name not like '%anmeldung%' and name not like '%datenliste%' and name not like '%Herzzentrum%' and name not like '%Termin%' and name not like '%Attest%' and name not like '%Schreiben%' and name not like '%-Bakt%' and name not like '%Bescheinigung%' and name not like '%Dokumentation%' and name not like '%Vorschläge%' and name not like '%Pumpeneinstellung%' and name not like '%anforderung%' and name not like '%Medikamentenplan%' and name not like '%befund%' and name not like '%blutdruck%' and name not like '%analyse%' and name not like '%sonogramm%' and name not like '%übersicht%' and name not like '%auswertung%' and name not like '%GPD%' and name not like '%untersuchung%' and name not like '%standar_tag%'", DBCn, adOpenStatic, adLockOptimistic)
'    If Not rd.BOF Then
'     Debug.Print rd!Name
'     Stop
'    End If
   End If
  Else
   Stop
  End If
 Next Fil
 Call ProgEnde
End Sub
#End If

Private Sub Excelliste_Click()
 Dim lad As New Lade
 Call lad.Show
End Sub

Private Sub Fallzahlstand_Click()
 Dim tage$
 Call ProgStart
 Do While Not IsNumeric(tage)
  tage = InputBox("Vor wie viel Tagen?", "Rückfrage", 0)
 Loop
 Call dofallzahlstand(Me, CDbl(tage))
 Call ProgEnde
End Sub

Private Sub Hilfsmittelklassifikationen_Click()
 Call ProgStart
 Call doHilfsmittelklassifikationen(Me)
 Call ProgEnde
End Sub

Private Sub gefaxttrennen_Click()
 Call ProgStart
' Call dogefaxttrennen
 Call ProgEnde
End Sub

Private Sub KassenkategorienBestimmen_Click()
 Dim rs As ADODB.Recordset, rAF&
 Call ProgStart
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'AOK' where name like 'aok%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'BKK' where name like '%bkk%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'BKK' where name like '%betriebsk%' or name like '%sbk%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'IKK' where name like '%ikk%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'LKK' where name like '%lkk%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'BKN' where name like '%bkn%' or name like '%knapp%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'EK' where (name like '%BARMER%' or name like '%DAK%' or name like '%TKK%' or name like '%KKH%' or name like '%HEK%' or name like '%HMK%' or name like '%HKK%' or name like '%GEK%' or name like '%HZK%' or name like '%KEH%' or name like '%HAMBURG-MÜNCHENER%' or name like '%Handelskrankenkasse%' or name like '%Techniker%' or name like '%HANSEATISCHE%' or name like 'AUS %') and isnull(Kateg)", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'EK' where name like '%EK %' and isnull(Kateg)", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'AOK' where kurzname like 'aok%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'BKK' where kurzname like '%bkk%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'BKK' where kurzname like '%betriebsk%' or kurzname like '%sbk%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'IKK' where kurzname like '%ikk%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'LKK' where kurzname like '%lkk%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'BKN' where kurzname like '%bkn%' or kurzname like '%knapp%'", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'EK' where (kurzname like '%BARMER%' or kurzname like '%DAK%' or kurzname like '%TKK%' or kurzname like '%KKH%' or kurzname like '%HEK%' or kurzname like '%HMK%' or kurzname like '%HKK%' or kurzname like '%GEK%' or kurzname like '%HZK%' or kurzname like '%KEH%' or kurzname like '%HAMBURG-MÜNCHENER%' or kurzname like '%Handelskrankenkasse%' or kurzname like '%Techniker%' or kurzname like '%HANSEATISCHE%' or kurzname like 'AUS %') and isnull(Kateg)", rAF)
 Debug.Print rs.source, rAF
 Set rs = DBCn.Execute("update kassenliste set Kateg = 'EK' where kurzname like '%EK %' and isnull(Kateg)", rAF)
 Debug.Print rs.source, rAF
 Call ProgEnde
End Sub

Private Sub LaborEintragen_Click()
 Call ProgStart
' Set ple3.hlese = Me
' Set ple2.hlese = Me
 Set ple.hlese = Me
 ple.Show
 Call ProgEnde
End Sub

Private Sub MDIForm_Activate()
 On Error Resume Next
 Me.dlg.Hide
End Sub

Private Sub MedartenEditieren_Click()
 Dim mda As New Medarten
 mda.Show
End Sub

Private Sub MyDB_Change()
 If Not obStart Then
   Screen.MousePointer = vbHourglass
   If Not obStart Then Call PutReg(Me)
'   obStart = True
   Me.obMySQL = True
'   obStart = False
   Me.dbv.DaBa = Me.MyDB
'   If Not Me.Visible Then Stop
   Me.Visible = True
   Screen.MousePointer = 0
 End If
End Sub

Private Sub MyDB_Click()
 Call MyDB_Change
End Sub

Private Sub MyDB_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me, Me.MyDB.name)
End Sub
Private Sub obMySQL_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me, Me.MyDB.name)
End Sub

Private Sub MyDB_LostFocus()
 Call MyDB_Change
End Sub


Private Sub obAcc_Click()
 If Not obStart Then
  If Me.Visible Then Screen.MousePointer = vbHourglass
'  Me.dlg.obAcc = Me.obAcc
  Me.dbv.changeStill = True
  Me.dbv.Datei = Me.Ziel
  Me.dbv.changeStill = False
  Me.dbv.ODBC = "Microsoft Access Driver (*.mdb)"
'  Call Me.dbv.RegSpeichern
  Call acon(quelleT)
'  Call dbv.cnVorb("", "anamnesebogen", "Patientendaten")
  Call PutReg(Me)
  If Me.Visible Then Screen.MousePointer = vbNormal
 End If
End Sub

Private Sub obMySQL_Click()
 If Not obStart Then
  If Me.Visible Then Screen.MousePointer = vbHourglass
  Call dbv.cnVorb("", "anamnesebogen", "Patientendaten")
  Me.dbv.ODBC = "MySQL ODBC 3.51 Driver"
  Call acon(quelleT)
  Call PutReg(Me)
  If Me.Visible Then Screen.MousePointer = 0
 End If
End Sub

Private Sub Patientenlaufzettel_Click()
 Call doPatientenlaufzettel
End Sub

Private Sub PLZeinzeln_Click()
 Call doPLZeinzeln
End Sub ' PLZeinzeln_Click

Private Sub Porto_Click()
 Call ProgStart
 Call doPorto
 Call ProgEnde
End Sub

Private Sub RestlicheBriefe_Click() ' restliche Briefe
 Call ProgStart
 Aktion = RestlicheBriefe
 Set pau.hlese = Me
 pau.Show
' Aktion = nix
 Call ProgEnde
End Sub
Public Sub los()
 Select Case Aktion
  Case BriefSchreiben
   Call tubriefStandalone(Me.pau.PatID, False)
  Case RestlicheBriefe
   Call doRestlicheBriefe(Me, Me.pau.PatID)
 End Select
End Sub ' los
Private Sub BriefSchreiben_Click() ' Brief schreiben
 Call ProgStart
 Aktion = BriefSchreiben
 Set pau.hlese = Me
 Screen.MousePointer = vbHourglass
 pau.Show
 Screen.MousePointer = vbNormal
' Aktion = nix
 Call ProgEnde
End Sub ' BriefSchreiben_Click

Private Sub CallDMPString_Click()
 Dim erg
 Call ProgStart
 erg = InputBox("Bitte Patientenummer eingeben!")
 If IsNumeric(erg) Then
  Call doCallDMP(ByVal erg)
 End If
 Call ProgEnde
End Sub ' CallDMPString_Click()
Private Sub DMPForts_Click()
 Static Ausw As New ADODB.Recordset
 If DBCn Is Nothing Then Call ProgStart
 If Ausw.State = 0 Then
  Ausw.Open "SELECT distinct pat_id, nachname, vorname, gebdat FROM dmpreihe where datum > " & datform(Now() - 365 * 0.5) & " order by nachname, vorname;", DBCn, adOpenDynamic, adLockReadOnly
 End If
 Ausw.Find "pat_id = " & lDMPPat_id, , adSearchBackward, 1
 If Not Ausw.EOF Then
  Ausw.Move 1
  Call doCallDMP(Ausw!Pat_id)
 End If
End Sub
Sub doCallDMP(ByVal pid&)
 Dim dmpstD$, erg$, DT As dmptyp ' Dateiname
 Dim rAna As New ADODB.Recordset
 ReDim rNa(0)
 rNa(0).Pat_id = pid
 rAna.Open "select * from namen where pat_id = " & pid, DBCn, adOpenStatic, adLockReadOnly
 If Not rAna.BOF And Not IsNull(rAna!Nachname) And Not IsNull(rAna!Vorname) Then
  dmpstD = GesNam(rAna) & " (" & pid & ") "
 Else
  dmpstD = "p:\DmpString "
 End If
 rAna.Close
 erg = DMPString$(rNa(0).Pat_id, DT)
 If lies.obMySQL Then
  dmpstD = dmpstD & Me.MyDB
 Else
  dmpstD = dmpstD & Replace(Replace(dlg.MdB, "\", "_"), ":", ".")
 End If
 dmpstD = dmpstD & " " & Format$(Date, "dd.mm.yy") & ".txt"
 Open dmpstD For Output As #391
 Print #391, erg
 Close #391
 Call Shell(Environ("systemroot") & "\system32\notepad.exe " & dmpstD, vbMaximizedFocus)
End Sub

Private Sub CallUSDM_Click()
 Dim erg
 Call ProgStart
 ReDim rNa(0)
 erg = InputBox("Bitte Patientenummer eingeben!")
 If Not IsNumeric(erg) Then Exit Sub
 rNa(0).Pat_id = erg
 Call usDM(True)
 Call ProgEnde
End Sub ' CallUSDM_Click()

'Private Sub DMPListeAnzeigen_Click()
' Call ProgStart
' Call doDMPListeAnzeigen(Me)
' Call ProgEnde
'End Sub

Private Sub FehlendeHausaerzte_Click()
'select distinct namen.pat_id, namen.nachname, namen.vorname from quelle.namen left join quelle.faelle on faelle.pat_id = namen.pat_id where (namen.kvnr = "" or isnull(namen.kvnr)) and faelle.schgr = 24 order by pat_id desc;
 Call ProgStart
 Call doLdFH(Me)
 Call ProgEnde
End Sub ' FehlendeHausaerzte_Click()

Private Sub holLaborParameter_Click()
 Call ProgStart
 Call holAllg(Me, "laborparameter", "", 0)
 Call ProgEnde
End Sub
Private Sub holEBM2000plus_Click()
 Call ProgStart
 Call holAllg(Me, "ebm2000plus", "", 0)
 Call ergEBM(Me)
 Call ProgEnde
End Sub

Private Sub ListeDerFehlendenDokumente_Click()
 Call ProgStart
 Call doLdFD
 Call ProgEnde
End Sub ' ListeDerFehlendenDokumente_Click()

Private Sub TabKop_Click()
 Set Tü.F0 = Me
 Tü.Show
End Sub

Private Sub test_Click()
 Call bittest1
End Sub

'Private Sub LaborTestLaden_Click()
'Set labtest.dlg = Me.dlg
' labtest.Show
'End Sub
Private Sub TherapieartenFestlegen_Click()
 Dim nTher As TherapieArt, rAF&
 Call ProgStart
 Dim rF As New ADODB.Recordset
 If 1 = 0 Then
 Call rF.Open("select * from faelle order by fid ", DBCn, adOpenStatic, adLockReadOnly)
 Do While Not rF.EOF
  nTher = TherArt(rF!Pat_id, True, , QEnd(rF!Quartal), rF!Quartal)
  If IsNull(rF!TherArt) Or rF!TherArt <> nTher Then
   Call DBCn.Execute("update faelle set therart = " & nTher & " where pat_id = " & rF!Pat_id & " and quartal = '" & rF!Quartal & "'", rAF)
   syscmd 4, "Therapieart für " & rF!Pat_id & " " & rF!Nachname & " " & rF!Vorname & ", Quartal: " & rF!Quartal & " zu " & TherUmw(nTher) & " festgelegt."
  End If
  rF.MoveNext
 Loop
 End If
 
 Dim rsAna As New ADODB.Recordset
' rsAna.Open "select pat_id,diabetestyp,-insulinpumpe as j_insulinpumpe,ther1,therakt from anamnesebogen", DBCn, adOpenStatic, adLockOptimistic
 rsAna.Open "select pat_id,diabetestyp,insulinpumpe,ther1,therakt from anamnesebogen", DBCn, adOpenStatic, adLockOptimistic
 Do While Not rsAna.EOF
  Call TherapieArtEinzelnFestlegen(rsAna!Pat_id, rsAna)
  rsAna.Move 1
 Loop
 Call ProgEnde
End Sub ' TherapieartenFestlegen_Click()

#If mitab Then

Private Sub Medklass_Click()
 Call ProgStart
 Call doMedklassT
 Call ProgEnde
End Sub ' MedklassT_Click()

Private Sub ÜbertragenenAnamnesebogen_Click()
 anBogÜ.Show
End Sub
#End If

Private Sub MedklassT_Click()
 Call ProgStart
 Call doMedklassT
 Call ProgEnde
End Sub ' MedklassT_Click()


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
Private Sub HAUebertrag_Click()
 Call ProgStart
 Call holAllg(Me, "hausaerzte", "ID", -1)
 Call ProgEnde
End Sub
Private Sub holMedArten_Click()
 Call ProgStart
'Function holMA(frm As Lese)
' Rökan <> Roekan =>
'alter table medplan drop foreign key MedArtenMedPlan_AccRel;
'ALTER TABLE medplan MODIFY COLUMN `MedAnfang` VARCHAR(35) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL;
'delete from medarten;
'alter table medarten modify column Medikament varchar(50) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL;
'insert into medarten (medikament) SELECT distinct medanfang FROM medplan m;
'alter TABLE  medplan add CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`);
' If lies.obmysql Then
'  Call DBCn.Execute("set foreign_key_checks = 0")
'  Call DBCn.Execute("delete from medarten")
' End If
 Call holAllg(Me, "medarten", "Medikament", 0)
' If lies.obmysql Then
'  Call DBCn.Execute("set foreign_key_checks = 1")
' End If
 Call ProgEnde
End Sub

'Private Sub DMPListe_Click()
' Call ProgStart
' Call DMPlst.init(Me)
' DMPlst.Show
'End Sub

Private Sub DMPSend_Click()
 Call ProgStart
 Call alleDMPs(Me)
 Call ProgEnde
End Sub
Private Sub DMPLeiDok_Click()
 Call ProgStart
 Call alleDMPLeiDok(Me)
 Call ProgEnde
End Sub
Private Sub BriefeLeiDok_Click()
 Call ProgStart
 Call tuBriefeLeiDok(Me)
 Call ProgEnde
End Sub

Private Sub FormulareÜbertragen_Click()
 Call ProgStart
 Call ÜbertragFormulare(Me, "_Medarten", "Medarten", "u:\zugriff.mdb", "medarten")
 Exit Sub
 Call ÜbertragFormulare(Me, "Anamnesebogen", "AnBog", "u:\zugriff.mdb", "anamnesebogen")
 Call ÜbertragFormulare(Me, "Labordokumente eP", "Labordokumente", "u:\zugriff.mdb", "dokumente")
 Call ProgEnde
End Sub

Private Sub KassenlisteKopieren_Click()
 Call ProgStart
 Call holAllg(Me, "kassenliste", "ID", 0)
 Call ProgEnde
End Sub ' KassenlisteKopieren_Click()

Private Sub DokumenteAbgehaktkopieren_Click()
 Call ProgStart
 Call holDA(Me)
 Call ProgEnde
End Sub
'Private Sub DokumentPfadeKorrigieren_Click()
' Call ProgStart
' Call DokPfadKorr(Me)
' Call ProgEnde
'End Sub

Private Sub AnamnesebogenPacken_Click()
 Call ProgStart
 Call AnPack
 Call ProgEnde
End Sub

'Private Sub LaborAlle_Click()
' Call ProgStart
' Call LaborDirektImport(Me, -1)
' Call LaborErgPatId(Me)
' Call ProgEnde
'End Sub ' LaborAlle_Click()

'Private Sub LaborFehlende_Click()
' Call ProgStart
' Call LaborDirektImport(Me, 0)
' Call LaborErgPatId(Me)
' Call ProgEnde
'End Sub ' LaborFehlende_Click()

'Private Sub Laborquerverbinden_Click()
' Call ProgStart
' Call LaborErgPatId(Me)
' Call ProgEnde
'End Sub ' Laborquerverbinden_Click()

'Private Sub LaborQuerverbindenMitLöschen_Click()
' Call ProgStart
' Call LaborErgPatId(Me, -1)
' Call ProgEnde
'End Sub ' LaborQuerverbindenMitLöschen_Click()

Private Sub MachTypen_Click(Index As Integer)
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
   Print #301, Right$("   " & lfdnr, 3) & ": Pat: " & Right$("   " & rs!Pat_id, 4) & " Auftrag: " & Right$("       " & rs!Auftragsnummer, 11) & " " & rs!Eingang & " " & IIf(Len(rs!Pfad) < 50, Right$(Space$(50) & rs!Pfad, 50), rs!Pfad) & " (Pat: " & rs!Nachname & ", " & rs!Vorname & " Werte: " & Trim$(rs!LWerte) & ")"
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
 Call test7
' Call löschBezügeausLaborux(65)
' Call LöschDateiEintrag(419)
 Call ProgEnde
End Sub
Private Sub Eintragszusatz_Click()
 Call ProgStart
 Call EintragZusatz
 Call ProgEnde
End Sub

Private Sub TabVergleich_Click()
 Call ProgStart
 Call VergleichTab(Me)
 Call ProgEnde
End Sub

Private Sub UngeschriebeneBriefe_Click() ' Briefe zu Pat., zu denen bisher noch kein Brief geschrieben wurde
 Call ProgStart
 Call doUngeschriebeneBriefe(Me)
 Call ProgEnde
End Sub
Private Sub doFollowUp_Click() ' &Briefe zu Patienten, deren erster Brief schon > 1a her ist
 Call ProgStart
 Call dodoFollowUp(Me)
 Call ProgEnde
End Sub

Private Sub Vergleichen_Click() ' Vergleich der Datenbanken
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


Private Sub Beenden_Click(Index As Integer)
'#If False Then
' If ProgrammLauf(-1) Or DBCn.State = 0 Then ' Wenn Programm schon gestoppt war, dann abbrechen, sonst stoppen
'#End If
 If Not ProgLäuft Or DBCn.State = 0 Then
  Unload Me
  End
 End If
End Sub ' Beenden_Click(Index As Integer)


Private Sub Beginn_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub Bytes_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub Ende_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub GesBytes_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub GesDauer_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub


Private Sub nachCd_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub obAcc_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub obMy_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub Picture2_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub Prozent_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub


Private Sub Sekunden_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub SyncTest_Click()
 Call formInhMach
End Sub ' SyncTest_Click

Private Sub ÜbertrageCd_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub ViewsErstellen_Click()
 Call doViewsErstellen
End Sub

Private Sub VorhandeneBriefe_Click()
 Dim ergZ&
 Call ProgStart
 ergZ = doVorhandene
 Me.Ausgabe = "Fertig mit Auswechseln der Vorhandenen: " & ergZ & IIf(BrichAb, ",5", "") & " Briefe bearbeitet. " & vbCrLf & altAusgabe
 altAusgabe = Me.Ausgabe
 Call ProgEnde
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VorhandeneBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' VorhandeneBriefe_Click

Private Sub WiedereinbestellungenDMP_Click()
 Dim sql$, Zp$, obDruck%
 Dim r1 As New ADODB.Recordset
 Dim r2 As New ADODB.Recordset
 Const WDatei$ = "p:\Wiedereinbestellungen.txt"
 Open WDatei For Output As #339
 'r1.Open "select pat_id, notiz from namen na where notiz like '%hier%'", dbv.wCn, adOpenStatic, adLockReadOnly
 'select pat_id, notiz from namen na join (select pat_id from dmpreihe dr union select pat_id from namen na where notiz like '%hier%' order by pat_id) as innen using (pat_id)
 r1.Open "select pat_id, na.nachname, na.vorname, notiz from namen na join (select pat_id from dmpreihe dr union select pat_id from namen na where notiz like '%hier%' order by pat_id) as innen using (pat_id) join anamnesebogen an using (pat_id) where tkz=0", dbv.wCn, adOpenStatic, adLockReadOnly
 Do While Not r1.EOF
'  If obhierdmp(r1!notiz) Then
   Set r2 = Nothing
   r2.Open "select * from labor1 where pat_id = " & r1!Pat_id & " and abkü = 'hba1c' union select * from labor2 where pat_id = " & r1!Pat_id & " and abkü = 'hba1c' order by zeitpunkt desc limit 1", dbv.wCn, adOpenStatic, adLockReadOnly
   Zp = ""
   obDruck = True
   If Not r2.EOF Then
    Zp = r2!Zeitpunkt & " " & Right$(Space$(3) & r2!Wert, 3)
    If r2!Zeitpunkt >= CDate("1.4.08") Then obDruck = False
   End If
   If obDruck Then
    Debug.Print r1!Pat_id, r1!Nachname, r1!Vorname, IIf(obhierdmp(r1!Notiz), "X", " "), Replace(Replace(r1!Notiz, vbCr, ""), vbLf, ""), Zp
    Print #339, Right$(Space$(4) & r1!Pat_id, 4) & " " & Left(r1!Nachname & Space$(15), 15) & " " & Left(r1!Vorname & Space$(11), 11) & "   " & IIf(obhierdmp(r1!Notiz), "X", " ") & "   " & Left(Replace(Replace(r1!Notiz, vbCr, ""), vbLf, "") & Space$(55), 55) & " " & Zp
    Print #339, String$(110, "_")
   End If
'  End If
  r1.Move 1
 Loop
 Close #339
 Call Shell(Environ("systemroot") & "\system32\notepad.exe " & WDatei, vbNormalFocus)
End Sub

Private Sub WSt0Erg_Click()
 Call ProgStart
 Call doWSt0Erg
 Call ProgEnde
End Sub

Private Sub Zeilen_KeyDown(keyCode As Integer, Shift As Integer)
 Call key(keyCode, Shift, Me)
End Sub

Private Sub Ziel_Change()
 dlg.MdB = Me.Ziel
 If Me.obAcc Then
  Call dbv.cnVorb("quelle", "anamnesebogen", "Patientendaten")
 End If
End Sub

Private Sub Ziel_Click()
 dlg.MdB = Me.Ziel
End Sub

Private Sub Ziel_KeyDown(keyCode As Integer, Shift As Integer)
 dlg.MdB = Me.Ziel
 Call key(keyCode, Shift, Me)
End Sub

Private Sub Picture1_KeyDown(keyCode As Integer, Shift As Integer)
' If KeyCode = 27 Then End
 Call key(keyCode, Shift, Me)
End Sub ' Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Ausgabe_KeyDown(keyCode As Integer, Shift As Integer)
 Static obCtrl%, pa&, pe&, Zahl$
 Call key(keyCode, Shift, Me)
 If keyCode = 17 Then
  obCtrl = True
 Else
  If obCtrl Then
   If keyCode = 68 Then ' D
    pe = Me.Ausgabe.SelStart
    Do While Mid$(Ausgabe.Text, pe, 2) <> vbCrLf
     pe = pe + 1
     If pe = Len(Ausgabe.Text) - 1 Then
      pe = Len(Ausgabe.Text)
      Exit Do
     End If
    Loop
    pa = pe
    Do While Mid$(Me.Ausgabe.Text, pa, 1) <> " " And pa > 0
     pa = pa - 1
    Loop
    Zahl = Mid$(Me.Ausgabe.Text, pa + 1, pe - pa - 1)
    Clipboard.SetText Zahl
    Call doCallDMP(ByVal Zahl)
   End If
  End If
  obCtrl = False
 End If
' Debug.Print "Schlüssel: " & KeyCode & " " & Shift
End Sub ' Ausgabe_Keydown

#If False Then
Private Sub Zurücksetzen_Click()
 Call ProgrammLauf(-1) ' falls es fälschlich auf 0 steht: 0 = Programm läuft, -1 = nicht
 Dim ctl As Control
 For Each ctl In Me.Controls
  If ctl.name Like "*inlesen*" Then ctl.Enabled = True
 Next ctl
 Me.Zurücksetzen.Enabled = False
End Sub
#End If

Private Sub Datenbank_Click()
 Screen.MousePointer = vbHourglass
 dlg.BDTDatei = getLDatei(dlg.BDTDatei, "*.bdt")
 dlg.Show
 Screen.MousePointer = vbDefault
End Sub ' Datenbank_Click()
Public Function PutEinstAufDB()
 Dim rs As New ADODB.Recordset
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
   rs!SammelInsert = dlg.SammelInsert
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
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PutEinstAufDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PutEinstAufDB
Public Function HolEinstvonDB()
 Dim rs As New ADODB.Recordset
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
   dlg.SammelInsert = Abs(HolEinstFeld(rs, "SammelInsert", 0, "BIT"))
   dlg.LaborDirektEinlesen = Abs(HolEinstFeld(rs, "LaborDirektEinlesen", 0, "BIT"))
   dlg.LaborDirektNeu = Abs(HolEinstFeld(rs, "LaborDirektNeu", 0, "BIT"))
   dlg.LaborQuerVerb = Abs(HolEinstFeld(rs, "LaborQuerVerb", -1, "BIT"))
   dlg.LaborQuerNeu = Abs(HolEinstFeld(rs, "LaborQuerNeu", 0, "BIT"))
   dlg.lBDTDatei = HolEinstFeld(rs, "Datei", üVerz, "TEXT", 120)
   Call dlg.FrmLEinlesung
'   dlg.BDTDatei = getLDatei(dlg.BDTDatei, "*.bdt")
   
'   dlg.BDTAend = FSO.GetFile(dlg.BDTDatei).DateLastModified
'   dlg.BDTDatei = getLDatei(dlg.BDTDatei, "*.bdt")
   dlg.LaborPfadBeispiel = HolEinstFeld(rs, "LaborPfadBeispiel", LabTransPfad, "TEXT")
   dlg.obmitEmails = Abs(HolEinstFeld(rs, "obmitEmails", -1, "BIT"))
   dlg.obBDT = Abs(True)
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HolEinstVonDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function HolEinstFeld(rs As ADODB.Recordset, FName$, Default, ByVal Typ$, Optional lenge&)
 Dim FNr&
 On Error Resume Next
 HolEinstFeld = rs(FName)
 FNr = Err.Number
 If FNr <> 0 Then
'  On Error GoTo fehler
  rs.Close
  If lies.obMySQL Then
   Select Case Typ
    Case "TEXT": Typ = "VARCHAR": If lenge = 0 Then Typ = "VARCHAR(300)"
    Case "BIT": If lenge = 0 Then Typ = "BIT(1)" '  Typ = "TINYINT": If Lenge = 0 Then Typ = "TINYINT(1)"
    Case "DATE", "TIME": Typ = "DATETIME"
    ' identisch verwendbar sind evtl.: "LONGTEXT","DECIMAL","FLOAT","BINARY"
   End Select
  End If
  Call DBCn.Execute("Alter Table `eintragszahlen` ADD Column `" & FName & "` " & Typ & IIf(lenge <> 0, "(" & lenge & ")", "") & IIf(Null, " NULL", ""))
  Set rs = DBCn.Execute("select * from eintragszahlen where beginn = (select max(beginn) from eintragszahlen)")
  HolEinstFeld = Default
 End If
  Select Case VarType(HolEinstFeld)
   Case Is < 8
   Case Else
    If HolEinstFeld = "" Then HolEinstFeld = Default
  End Select
 On Error Resume Next
 If InStrB(HolEinstFeld, ":\\") > 0 Or InStrB(HolEinstFeld, "\\\\") > 0 Then HolEinstFeld = Replace(HolEinstFeld, "\\", "\")
 On Error GoTo fehler
 If IsNull(HolEinstFeld) Then HolEinstFeld = 0
 On Error GoTo fehler
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HolEinstFeld/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' HolEinstFeld(rs As ADODB.Recordset, FName$, Default, ByVal Typ$, Optional Lenge&)
Private Sub dbv_wCnAendern(CnStr$)
 On Error GoTo fehler
 ConStri = dbv.Constr
 obStart = True
 Me.obMySQL = InStrB(UCase$(ConStri), "MYSQL") <> 0
 Me.obAcc = Not Me.obMySQL
 
 If Me.obMySQL <> 0 Then
  If InStrB(dbv.Tabellen, "anamnesebogen") <> 0 Then
   If dbv.DaBa = "" And Not DBCn Is Nothing Then
    Me.MyDB = DBCn.Properties("Current Catalog")
   Else
    Me.MyDB = dbv.DaBa
   End If
  End If
 Else
  Me.Ziel = dbv.Datei
 End If
 
' 2.7.08: Kommentar der nä 3 Zeilen fentfernt, da sonst bei Änderungen der Mysql-Tabellen weder in
' Lese noch in DBVerb DBCn korrigiert wird
 If obQuelle Then ' Me.dbv.DaBa = InStrB(Me.dbv.Ü2, "Patienten") <> 0 Then  ' sonst bekommt DBCn Bedeutungen von Hausärzten usw
  Set DBCn = Me.dbv.wCn
 End If
 LVobMySQL = InStrB(UCase$(CnStr), "MYSQL") <> 0 '(Not (cDtb = accDtb))
 Me.obMySQL = LVobMySQL
 Me.obAcc = Not Me.obMySQL
 obStart = False
 Call Zinit(LVobMySQL)
 On Error Resume Next
 If Not DBCn Is Nothing Then
  Forms(0).ConStri = "geöffnet: " & Forms(0).dbv.Constr
 Else
  Forms(0).ConStri = "geschlossen: " & Forms(0).dbv.Constr
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in wcn_Aendern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' dbv_wCnAendern
Private Sub MDIForm_Load()
  Dim Diff#
  On Error GoTo fehler
  Call WD
  
 imAufbauLese = True
' Set qdb = New QuelleDBC
 Set dbv = New DBVerb
 Set lies = Me
' Set dlg = New Dialog
 Set dlg.hlese = Me
' Set snst = New Sonstige
 Set snst.hlese = Me
 Call HolReg(Me)
 Me.Zeilen = 0
 Me.Bytes = 0
 With Me.CommonDialogLese
  .DialogTitle = "Einzulesende Datei"
  .InitDir = üVerz '"u:\tmexport"
  .Filename = "*.BDT"
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
 Call Me.MyDB.AddItem("quelle")
 Call Me.MyDB.AddItem("quelle1")
 Call Me.MyDB.AddItem("quelle2")
 If Command = "auto" Then
  Me.dlg.obAcc = True
  Me.dlg.obVglMitLetzterEinlesung = 1
  Me.dlg.LaborDirektEinlesen = 1
  Me.dlg.LaborQuerVerb = 1
  Me.dlg.LaborDirektNeu = False
  Me.dlg.LaborQuerNeu = False
  Diff = CDate(Me.dlg.BDTAend) - CDate(Me.dlg.LDateiAend)
  If Diff > 0 Then
   Call Me.dlg.Start_Click
  End If
  Call snie
  
  Me.dlg.obMySQL = True
  Call dbv.cnVorb("quelle", "anamnesebogen", "Patientendaten")
  Me.dlg.obVglMitLetzterEinlesung = 1
  Me.dlg.LaborDirektEinlesen = 1
  Me.dlg.LaborQuerVerb = 1
  Me.dlg.LaborDirektNeu = False
  Me.dlg.LaborQuerNeu = False
  Diff = CDate(Me.dlg.BDTAend) - CDate(Me.dlg.LDateiAend)
  If Diff > 0 Then
   Call Me.dlg.Start_Click
  End If
  Call snie
  Call snie
  
'  Me.dlg.obMyQuelle1 = True
  Call dbv.cnVorb("quelle1", "anamnesebogen", "Patientendaten")
  Me.dlg.obVglMitLetzterEinlesung = 1
  Me.dlg.LaborDirektEinlesen = 1
  Me.dlg.LaborQuerVerb = 1
  Me.dlg.LaborDirektNeu = False
  Me.dlg.LaborQuerNeu = False
  Diff = CDate(Me.dlg.BDTAend) - CDate(Me.dlg.LDateiAend)
  If Diff > 0 Then
   Call Me.dlg.Start_Click
  End If
  Call snie
  Call snie
  Call snie
  
'  Me.dlg.obMyQuelle2 = True
  Call dbv.cnVorb("quelle2", "anamnesebogen", "Patientendaten")
  Me.dlg.obVglMitLetzterEinlesung = 1
  Me.dlg.LaborDirektEinlesen = 1
  Me.dlg.LaborQuerVerb = 1
  Me.dlg.LaborDirektNeu = False
  Me.dlg.LaborQuerNeu = False
  Diff = CDate(Me.dlg.BDTAend) - CDate(Me.dlg.LDateiAend)
  If Diff > 0 Then
   Call Me.dlg.Start_Click
  End If
  Call snie
  Call snie
  Call snie
  Call snie
  
  Unload Me
  End
  ' Me.dlg.Visible
 End If
 Me.Version = App.Major & " " & App.Minor & " " & App.Revision
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MDIForm_Load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' MDIForm_Load()
Public Function AbbrechDisable(frm As Lese)
 Dim i%
 For i = 0 To frm.Controls.Count - 1
  If frm.Controls(i).name = "Abbrechen" Then
   frm.Controls(i).Enabled = False
   Exit For
  End If
 Next i
End Function

Public Function ProgStart()
 BrichAb = 0
 T1 = Now
 Me.Controls!Abbrechen.Enabled = True
' Me.Controls!Beenden.Enabled = False
 Call BeendenBlend(False)
' Call Me.ConstrFestleg(0, Me)
 Call acon(quelleT)
End Function ' ProgStart

Public Function ProgEnde()
 Dim i&
 Me.Controls!Abbrechen.Enabled = False
' Me.Controls!Beenden.Enabled = True
 Call BeendenBlend(True)
End Function ' ProgEnde

Private Sub BeendenBlend(obTrue%)
 Dim i&
 For i = 1 To Me.Controls.Count
  If Me.Controls(i).name = "Beenden" Then
   Me.Controls(i).Enabled = obTrue
   Exit For
  End If
 Next i
 ProgLäuft = Not obTrue
End Sub ' BeendenBlend
Private Sub MDIForm_Resize()
 On Error Resume Next
 Picture1.Height = Me.Height - 5745 + 5055
 Ausgabe.Height = Me.Height - 5745 + 4250 - Me.Fuß.Height - Me.ConStri.Height
 Me.ÜbertrageCd.Top = Me.Ausgabe.Top + Me.Ausgabe.Height + 40
 Me.DMPForts.Top = Me.ÜbertrageCd.Top
 Me.QDatei.Top = Me.ÜbertrageCd.Top
 Me.QDatum.Top = Me.ÜbertrageCd.Top
 Me.obMySQL.Top = Me.ÜbertrageCd.Top
 Me.obAcc.Top = Me.ÜbertrageCd.Top
 Me.Ziel.Top = Me.ÜbertrageCd.Top
 Me.ZeilenBez.Top = Me.Height - 5745 + 4645 - Me.Fuß.Height - Me.ConStri.Height
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
 Me.DurschnDauerLab.Top = Me.ZeilenBez.Top
 Me.DurchschnDauer.Top = Me.ZeilenBez.Top + 20
 Me.Beginn.Top = Me.ZeilenBez.Top
 Me.EndeBez.Top = Me.ZeilenBez.Top + 20
 Me.Ende.Top = Me.ZeilenBez.Top
 Me.SBez.Top = Me.Ende.Top
 Me.Picture2.Top = Me.Ende.Top
 Me.MyDB.Top = Me.ÜbertrageCd.Top
 Me.GesamtDauerBez.Top = Me.ZeilenBez.Top + 20
 Me.GesDauer.Top = Me.ZeilenBez.Top
 Me.Fuß.Top = Picture1.Height - Me.Fuß.Height - Me.ConStri.Height
 Me.ConStri.Top = Me.Fuß.Top + Me.Fuß.Height
End Sub ' MDIForm_Resize()

Private Sub MDIForm_Unload(Cancel As Integer) ' geht nur beim Anklicken des Kreuzes oben
 Call PutEinstAufDB
 Call PutReg(Me)
 End
End Sub ' MDIForm_Unload(Cancel As Integer)
#If False Then
Public Function ConstrFestleg(ByVal art As ConDtb, Optional hlese As Lese)   ' dlg ist für art= 0 und 1 nötig
 On Error GoTo fehler
'ConStr$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;uid=praxis;pwd=sonne;option=" & opti
 Select Case art
  Case 0
   If hlese.obAcc Then
    art = accDtb
   Else
    If hlese.MyDB = "quelle" Then
     art = qDtb
    ElseIf hlese.MyDB = "quelle1" Then
     art = q1Dtb
    ElseIf hlese.MyDB = "quelle2" Then
     art = q2dtb
    End If
   End If
 End Select
 Dim MdB$
 If art = 1 Then
 ' ConStr = CStrAcc + """" + dlg.MdB + """"
  If Not obStart Then
   If hlese.Ziel = "" Or Not FSO.FileExists(hlese.Ziel) Then Call hlese.dlg.MdBFestleg
  End If
'  MdB = dlg.MdB
' Else
'  MdB = ""
 End If
 Call doConstrFestleg(art, obStart, hlese.Ziel, Me)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ConstrFestleg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ConstrFestleg
#End If
