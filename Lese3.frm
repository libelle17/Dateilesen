VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm Lese 
   BackColor       =   &H8000000C&
   Caption         =   "Patientendaten Diabetespraxis Dachau"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   990
   ClientWidth     =   19080
   Icon            =   "Lese3.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Lese3.frx":030A
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
         ItemData        =   "Lese3.frx":0614
         Left            =   8040
         List            =   "Lese3.frx":0616
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
         Caption         =   "Access"
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
      Begin VB.TextBox EndZp 
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
      Begin VB.Menu Optionen 
         Caption         =   "&Optionen"
      End
      Begin VB.Menu Datenbankverbindung 
         Caption         =   "&Datenbankverbindung Patientendaten"
      End
      Begin VB.Menu Zurücksetzen 
         Caption         =   "&Zurücksetzen des Programmlaufs"
         Visible         =   0   'False
      End
      Begin VB.Menu Beenden 
         Caption         =   "&Beenden"
         Index           =   3
      End
   End
   Begin VB.Menu FunktionenfürArzthelferinundArzt 
      Caption         =   "&Funktionen: für Arzthelferin und Arzt"
      Begin VB.Menu Datenbank 
         Caption         =   "&Einlesen"
      End
      Begin VB.Menu Abbrechen 
         Caption         =   "Abbre&chen (vor Speichern eines Patienten)"
      End
      Begin VB.Menu LaborEintragen 
         Caption         =   "Labor &eintragen"
      End
      Begin VB.Menu FalschAbgehakteUngueltig 
         Caption         =   "&Falsch abgehakte Dokumente ungültig stempeln"
      End
      Begin VB.Menu Abrechnungsfehler 
         Caption         =   "&Abrechnungsfehler"
      End
      Begin VB.Menu Patientenlaufzettel 
         Caption         =   "&Patientenlaufzettel"
         Shortcut        =   ^P
      End
      Begin VB.Menu PLZeinzeln 
         Caption         =   "P&atientenlaufzettel einzeln"
         Shortcut        =   ^E
      End
      Begin VB.Menu WiedereinbestellungenDMP 
         Caption         =   "&Wiedereinbestellungen DMP"
      End
      Begin VB.Menu HausärztemitalterKVNrergänzen 
         Caption         =   "&Hausärzte mit alter KV-Nr ergänzen"
      End
      Begin VB.Menu PatientenlistefürHausarztmodell 
         Caption         =   "&Patientenliste für Hausarztmodell"
      End
      Begin VB.Menu PatientenlistefürVollpauschale 
         Caption         =   "&Patientenliste für Vollpauschale"
      End
      Begin VB.Menu FalscheKarteikarteneinträge 
         Caption         =   "Falsche &Karteikarteneinträge"
      End
   End
   Begin VB.Menu FunktionenfürArzt 
      Caption         =   "... für &Arzt"
      Begin VB.Menu ÜbertragenenAnamnesebogen 
         Caption         =   "&Anamnesebogen (Diagnosen eingeben)"
      End
      Begin VB.Menu Anamnesebogen 
         Caption         =   "&Anamnesebogen experimentell"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu DMPSend 
         Caption         =   "Alle &DMP-Dokumente an Hausärzte in p:\zufaxen erstellen"
      End
      Begin VB.Menu DMPhierListe 
         Caption         =   "D&MP hier Liste"
      End
      Begin VB.Menu DMPKHKAsthma 
         Caption         =   "DM&P KHK Asthma"
      End
      Begin VB.Menu DuplexKontrollieren 
         Caption         =   "&Duplex Kontrollieren"
      End
      Begin VB.Menu DMPLeiDok 
         Caption         =   "DMP-Brief-&Leistungs-Doku"
      End
      Begin VB.Menu doPorto 
         Caption         =   "Po&rto abrechnen"
      End
      Begin VB.Menu BriefeLeiDok 
         Caption         =   "&Leistungen zu Arztbriefen bei Kassenpatienten in u:\tmimport\ dokumentieren"
      End
      Begin VB.Menu PathLabAnschau 
         Caption         =   "Pa&thologische Laborwerte anschauen"
      End
      Begin VB.Menu VorhandeneBriefe 
         Caption         =   "&Vorhandene Briefe korrigieren"
      End
      Begin VB.Menu doFollowUp 
         Caption         =   "B&riefe zu Patienten, deren erster Brief schon > 1a her ist"
      End
      Begin VB.Menu UngeschriebeneBriefe 
         Caption         =   "B&riefe zu Pat., zu denen bisher kein Brief geschrieben wurde"
      End
      Begin VB.Menu RestlicheBriefe 
         Caption         =   "&Restliche Briefe für aktuelles Quartal schreiben"
      End
      Begin VB.Menu BriefeBerichtspflicht 
         Caption         =   "B&riefe zu Patienten mit Berichtspflicht schreiben"
      End
      Begin VB.Menu BriefeZuListeSchreiben 
         Caption         =   "Briefe zu &Pat_ID-Liste aus Datei schreiben"
      End
      Begin VB.Menu BriefSchreiben 
         Caption         =   "&Brief schreiben"
         Shortcut        =   ^B
      End
      Begin VB.Menu BriefOhneMaske 
         Caption         =   "Brief &ohne Maske schreiben"
         Shortcut        =   ^O
      End
      Begin VB.Menu ListeDerFehlendenDokumente 
         Caption         =   "&Liste der fehlenden Dokumente"
      End
      Begin VB.Menu NachzuholendeLaborimporte 
         Caption         =   "Nach&zuholende Laborimporte"
      End
      Begin VB.Menu VerdächtigeÜberweiser 
         Caption         =   "Verd&ächtige Überweiser"
      End
      Begin VB.Menu FehlendeHausaerzte 
         Caption         =   "Liste der fehlenden &Hausärzte"
      End
      Begin VB.Menu DoppelteDiagnosen 
         Caption         =   "&Doppelte Diagnosen ermitteln"
      End
      Begin VB.Menu Motivationsgesprächskandidaten 
         Caption         =   "&Motivationsgesprächskandidaten"
      End
      Begin VB.Menu MedartenEditieren 
         Caption         =   "&Medarten editieren"
      End
   End
   Begin VB.Menu Statistik 
      Caption         =   "&Statistik"
      Begin VB.Menu HausärzteBKK 
         Caption         =   "&Hausärzte von Pat. in der BKK"
      End
      Begin VB.Menu Überweiserstatistik 
         Caption         =   "&Überweiserstatistik"
      End
      Begin VB.Menu Schulungsstatistik 
         Caption         =   "S&chulungsstatistik"
      End
      Begin VB.Menu Hilfsmittelklassifikationen 
         Caption         =   "&Hilfsmittelklassifikationen"
         Shortcut        =   ^H
      End
      Begin VB.Menu Fallzahlstand 
         Caption         =   "&Fallzahlstand"
         Shortcut        =   ^F
      End
      Begin VB.Menu Pumpenträgerliste 
         Caption         =   "&Pumpenträgerliste"
      End
      Begin VB.Menu Gestationsdiabetikerinnen 
         Caption         =   "&Gestationsdiabetikerinnen"
      End
   End
   Begin VB.Menu EDV 
      Caption         =   "&EDV"
      Begin VB.Menu MachDB 
         Caption         =   "&MachDB"
      End
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
      Begin VB.Menu korrQD 
         Caption         =   "korrQD"
      End
      Begin VB.Menu Apothekenrezepte 
         Caption         =   "Apothekenrezepte"
      End
      Begin VB.Menu EinlesungenAnzeigen 
         Caption         =   "&Einlesungen anzeigen"
      End
      Begin VB.Menu DokumenteAbgehaktPrüfen 
         Caption         =   "Dokumente abgeha&kt prüfen"
      End
      Begin VB.Menu DokumenteNeuAbhaken 
         Caption         =   "Dokumente neu a&bhaken"
      End
      Begin VB.Menu HausärzteEinlesen 
         Caption         =   "Haus&ärzte einlesen"
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
      Begin VB.Menu SortierungÄndern 
         Caption         =   "&Sortierung ändern"
      End
      Begin VB.Menu calldoGenMachDB 
         Caption         =   "&Generierte MachDB aufrufen"
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
'Public anb As New Anamnese
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
 Patientenlaufzetteleinzeln
 DMPZettel
End Enum
Public Aktion As AktionTyp
Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1
'Public WithEvents qdb As QuelleDBC
Enum UAbTyp
 Name = 0
 NLrt
' Fachrichtung
' Straße
' Ort
' Tel
' Fax
' Mail
' BSArt
' Sprechzeiten
' LANRBSNR
 Zusatzbezeichnung
 Weiterbildung
 Genehmigung
 Vertragsangebot
 Fremdsprache
 ÄrztederPraxis
End Enum

Private Type ArztTyp
 idarzt As Long
 Nachname As String
 Vorname As String
 titel_id As Long
' fachrichtung_id As Long
 LANR As Long
 nlart_id As Long
End Type

Private Type BSTyp
 idbs As Long
 Straße As New CString
 Hausnr As New CString
 Plz As String
 Ort_id As Long
 BSNR As String
 obNBS As Integer
 bsart_id As Long
 sprechzeiten_id As Long
 Rollst As Integer
End Type

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
     If Fl.Name Like "KV*rzte*.mdb" Then
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
    MsgBox "Stop in alteHausärzte_Click: " & vbCrLf & "acr.EOF"
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

Private Sub Apothekenrezepte_Click()
' select nachname, vorname, date(gebdat) as geb, fr.zeitpunkt, fa.feldinh from formular fr left join formular fa using (foid) left join namen on fr.pat_id = namen.pat_id where fr.feldinh like "%Gerald Schade;" and not isnull(fr.pat_id) and fr.formvorl like '%rezept%' and ((fr.formvorl like '%lang%' and fa.feld = 'medikament') or (fr.formvorl not like '%lang%' and fa.nr in (4,9,10,11))) and fr.zeitpunkt between '2008-02-01' and now() and not fa.feldinh like '%-  -%'
 Dim rs As New ADODB.Recordset
 Call ProgStart
 Open uVerz & "Apotheke.csv" For Output As #333
 rs.Open "select foid, nachname, vorname, date(gebdat) as geb, fr.zeitpunkt as Zeitp, fa.feldinh as text from formular fr left join formular fa using (foid) left join namen on fr.pat_id = namen.pat_id where fr.feldinh like '%Gerald Schade;' and not isnull(fr.pat_id) and fr.formvorl like '%rezept%' and ((fr.formvorl like '%lang%' and fa.feld = 'medikament') or (fr.formvorl not like '%lang%' and fa.nr in (4,8,9,10,11))) and fr.zeitpunkt between '2008-02-01' and now() and not fa.feldinh like '%-  -%'", DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  Print #333, rs!FoID & ";" & rs!Nachname & ";" & rs!Vorname & ";" & rs!Geb & ";" & rs!Zeitp & ";" & rs!Text
  rs.Move 1
 Loop
 Close #333
 MsgBox "Fertig mit Apotheke!"
End Sub

'Private Sub Anamnesebogen_Click()
' anb.Show
'End Sub

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

Private Sub calldoGenMachDB_Click()
 Dim DBn$
 DBn = InputBox("Datenbankname", "Eingabe der Datenbank", "quelle")
 If LenB(DBn) <> 0 Then
  Call doMach_quelle(DBn, DBVerb.Cpt)
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
 pad.Art = artDiag
 Set pad.hlese = Me
 pad.Show
End Sub

Private Sub DMPhierListe_Click()
 Call ProgStart
 Set pal.hlese = Me
 pal.Art = artPat
 Me.Hide
 pal.Show
 Call ProgEnde
End Sub ' DMPhierListe_Click()

Private Sub DMPKHKAsthma_Click()
 Dim rs As New ADODB.Recordset
 Call ProgStart
 rs.Open "select n.pat_id,n.nachname,n.vorname,f.form_abk, f.zeitpunkt from namen n left join formular f on n.pat_id = f.pat_id where (form_abk like 'dmpkhk%' or form_abk like 'edmpkhk%' or form_abk like 'edmpab%' or form_abk like 'dmpab%' or form_abk like 'dmpcopd%' or form_abk like 'edmpcopd%') group by n.pat_id,form_abk, date(zeitpunkt) order by n.pat_id, zeitpunkt", DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  Call Ausgeb(rs!Pat_id & " " & rs!Nachname & " " & rs!Vorname & " " & rs!Form_Abk & " " & rs!Zeitpunkt, True)
  rs.Move 1
 Loop
End Sub

Private Sub DokPfadKorrigieren_Click() ' Dokumentpfade korrigieren
 Call ProgStart
 Call dokpfadänder(Me)
 Call ProgEnde
End Sub
Private Sub DokumenteNeuAbhaken_Click()
 Dim rs As New ADODB.Recordset, rAF&
 Call ProgStart
 Call DBCn.Execute("delete from `dokumente abgehakt`", rAF)
 'Call DBCn.Execute("insert into `dokumente abgehakt`(aktzeit,abgehakt,dokpfad) select now() as aktzeit,1 as abgehakt, replace$(replace$(dokpfad,'\\','\\\\'),'\'','\\\'') from (select * from (select pat_id, zeitpunkt from labor1 group by pat_id, zeitpunkt order by pat_id, zeitpunkt) as i left join (select pat_id, date(quelldatum) as zeitpunkt, dokpfad, dokname from dokumente d where dokname like '%fremdlabor%') as d using (pat_id,zeitpunkt)) as i where not isnull(dokpfad)", rAF)
 Debug.Print "gelöscht:", rAF
 Call DBCn.Execute("insert into `dokumente abgehakt`(aktzeit,abgehakt,dokpfad) select now() as aktzeit,1 as abgehakt, dokpfad from (select * from (select pat_id, zeitpunkt from labor1 group by pat_id, zeitpunkt order by pat_id, zeitpunkt) as i left join (select pat_id, date(quelldatum) as zeitpunkt, dokpfad, dokname from dokumente d where dokname like '%fremdlabor%') as d using (pat_id,zeitpunkt)) as i where not isnull(dokpfad)", rAF)
 Debug.Print "eingefügt:", rAF
 Beep 1000, 1000
End Sub

Private Sub EinlesungenAnzeigen_Click()
 Dim rEinl As New ADODB.Recordset, sql$
 sql = "select st.*, e.datei as Datei, e.dateiaend as `letzte Änderung`, e.beginn as Einl_Beg, e.zp1, e.zp6, e.zp7, e.fallzahl as Fälle, e.sekunden as s from (select count(0) as Zahl, max(pat_id) as Pat_id, stbyte as StByte from namen group by stbyte) as st left join eintragszahlen e on st.stbyte = e.stbyte"
 Call ProgStart
 rEinl.Open sql, DBCn, adOpenStatic, adLockReadOnly
 TabAusgeb rEinl, True
End Sub

Public Sub FalschAbgehakteUngueltig_Click()
 Dim rAF&, zwg&, zug&, rs As New ADODB.Recordset, rl As New ADODB.Recordset
 Call ProgStart
 rs.Open "select --abgehakt ab, --ungueltig ug, pat_id, d.dokpfad, quelldatum qd from `dokumente abgehakt` da inner join dokumente d on da.dokpfad = d.dokpfad", DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  Set rl = Nothing
  rl.Open "select pat_id from laborneu where pat_id = " & rs!Pat_id & " and date(zeitpunkt) = " & DatForm(rs!qd), DBCn, adOpenStatic, adLockReadOnly
  If rl.EOF And rs!ug <> 1 Then
   DBCn.Execute "update `dokumente abgehakt` set ungueltig = 1 where dokpfad = '" & doUmwfSQL(rs!DokPfad, lies.obMySQL) & "'", rAF
   zug = zug + rAF
   If rAF = 0 Then
    MsgBox "Fehler beim Auffinden von " & doUmwfSQL(rs!DokPfad, lies.obMySQL) & " in `dokumente abgehakt` (Orginaldokpfad: " & rs!DokPfad & ")"
   End If
  ElseIf Not rl.EOF And rs!ug <> 0 Then
   DBCn.Execute "update `dokumente abgehakt` set ungueltig = 0 where dokpfad = '" & doUmwfSQL(rs!DokPfad, lies.obMySQL) & "'", rAF
   zwg = zwg + rAF
   If rAF = 0 Then
    MsgBox "Fehler beim Auffinden von " & doUmwfSQL(rs!DokPfad, lies.obMySQL) & " in `dokumente abgehakt` (Orginaldokpfad: " & rs!DokPfad & ")"
   End If
  End If
  rs.Move 1
 Loop
 Ausgeb zwg & " Dokumente wieder gültig gestempelt", True
 Ausgeb zug & " Dokumente ungültig gestempelt", True
End Sub ' FalschAbgehakteUngueltig_Click

Private Sub FalscheKarteikarteneinträge_Click()
 Dim ErgDat$
 ErgDat$ = pVerz & "FalscheKarteikarteneinträge.txt"
 Dim altArt$
 Dim rs As New ADODB.Recordset
 Call ProgStart
 Open ErgDat For Output As #327
 Print #327, "Falsche karteikarteneinträge vom " & Now() & ":"
 rs.Open "select pat_id,zeitpunkt,art,inhalt from eintraege where art not in (" & artSpezEintr & "," & artSpezUS & "," & artSpezSonst & ") order by art, pat_id, zeitpunkt", DBCn
 Do While Not rs.EOF
  If rs!Art <> altArt Then
   Print #327, String$(80, "_")
   altArt = rs!Art
  End If
  Print #327, Right$(Space$(4) & rs!Pat_id, 4) & "|" & Left$(rs!Zeitpunkt & Space$(19), 19) & "|" & Left$(rs!Art & Space$(10), 10) & "|" & Left$(rs!Inhalt, 50)
  rs.Move 1
 Loop
 Close #327
' Shell Environ("systemroot") & "\system32\notepad.exe " & ErgDat, vbMaximizedFocus
 Call GetWord
 With Wapp
   .Visible = True
   .WindowState = wdWindowStateMaximize
   .documents.Open ErgDat
   .activedocument.Range.Font.size = 9
   .Activate
 End With
End Sub

Function indIns&(cn As ADODB.Connection, Tb$, fld$, Wert$, idFld$)
 Dim rs As New ADODB.Recordset, rAF&
 Dim varlen&, pos&
 Dim STyp$
 On Error GoTo fehler
 rs.Open "select * from `" & Tb & "` where `" & fld & "` = '" & Wert & "'", cn, adOpenStatic, adLockReadOnly
 If rs.BOF Then
  Set rs = Nothing
  rs.Open "show full columns from `" & Tb & "` where field = '" & fld & "'", cn
  If Not rs.BOF Then
   STyp = rs.Fields("Type")
   If InStrB(STyp, "varchar(") <> 0 Then
    pos = InStr(STyp, "(")
    varlen = Mid(STyp, pos + 1, InStr(STyp, ")") - pos - 1)
    If varlen < Len(Wert) Then
     cn.Execute "Alter Table `" & Tb & "` modify `" & fld & "` varchar(" & Len(Wert) & ")"
    End If
   End If
  End If
  cn.Execute "insert into `" & Tb & "`(`" & fld & "`) values('" & Wert & "')", rAF
  Set rs = cn.Execute("select last_insert_id()")
  indIns = rs.Fields(0)
 Else
  indIns = rs.Fields(idFld)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in indIns/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' indIns
Public Function getHAPDF$(Optional obstumm%)
  With Me.CommonDialogLese
   .DialogTitle = "PDF-Datei mit Hausarztdaten von KVB aussuchen"
   .InitDir = vVerz
' "c:\v\KVB_Arztsuche_-20090919-221246.pdf"
   .DefaultExt = "pdf"
   .Filename = "KVB_Arztsuche*.pdf"
   Dim jFil$, jDat As Date, erg$
   erg = Dir(.InitDir & .Filename)
   Do While LenB(erg) <> 0
    If erg <> "." And erg <> ".." Then
     If FileDateTime(.InitDir & erg) > jDat Then
      jDat = FileDateTime(.InitDir & erg)
      jFil = erg
     End If
    End If
    erg = Dir
   Loop
   .Filename = .InitDir & jFil
'   .Flags = .Flags And Not FileOpenConstants.cdlOFNFileMustExist
   If Not obstumm Then .ShowOpen
   getHAPDF = .Filename
  End With
End Function

Private Sub HausärzteEinlesen_Click()
 Const DBName$ = "hausaerzte"
 Dim PDTextS As Acrobat.CAcroPDTextSelect
 Dim Result&, PDDoc As Acrobat.CAcroPDDoc
 Dim PDPage As Acrobat.CAcroPDPage
 Dim PDHili As Acrobat.CAcroHiliteList
 Dim B As Boolean, erg$
 Dim PDFStream$, Filename$
' Strings Falls Text von Anfang ausgelesen werden soll (bzw bis zum Ende)
 Dim FromFirst$, ToLast$
 Dim Str As New CString, NTL&
 Dim i%, j&, jj&, splitt$(), s2$(), Pid$, Datum As Date
 Dim Uhrzeit$
 Dim rAF&
 
 Filename = getHAPDF(True)
 FromFirst = Chr(169) & Chr(170) & Chr(172)
 ToLast = Chr(163) & Chr(165) & Chr(164)
 Set PDDoc = CreateObject("AcroExch.pdDoc")
 Result = PDDoc.Open(Filename)
 If Not Result Then
  MsgBox "Fehler beim Öffnen von: " & Filename, vbCritical, "Hausärzte einlesen"
  Exit Sub
 End If

#Const hadbneu = 0
#If hadbneu Then
 Call doMach_hausaerzte(DBName, DBVerb.Cpt)
#End If
' Filename = Environ("userprofile") & "\Desktop\" & "KVB_Arztsuche_-20090905-045301.pdf"
' Nehme die Erste Seite - Index 0
 Do
  Set PDPage = PDDoc.AcquirePage(j)
  ' Erzeuge ein Highlight Objekt und weise ihm 2000 Elemente bei (keine Grenzprobleme)
  Set PDHili = CreateObject("AcroExch.HiliteList")
  B = PDHili.Add(0, 32000)
  ' Erzeuge eine Textauswahl aus dem gesamten Text
  On Error Resume Next
  Set PDTextS = PDPage.CreatePageHilite(PDHili)
  If Err.Number <> 0 Then Exit Do
  On Error GoTo fehler
  ' Hole Anzahl der "Textblöcke"
  NTL = PDTextS.GetNumText
  ' Gebe den Text der Textauswahl zurück
  For i = 0 To NTL - 1
   Str.Append PDTextS.GetText(i)
  Next i
  DoEvents
  Str.Append vbCrLf
  j = j + 1
  lies.Ausgeb "Hausärzte einesen: Lese S.: " & j, 0
' die ersten 30 Seiten einlesen
'  If j > 334 Then Exit Do
 Loop
' in jeder Zeile von splitt steht eine Zeile aus der PDF-Datei
 SplitNeu Str.Value, vbCrLf, splitt, "(", ")"
 
 Dim HACn As New ADODB.Connection, rs As New ADODB.Recordset
 HACn.Open (DBCn.ConnectionString)
 HACn.Execute "use `" & DBName & "`"
 HACn.Execute "truncate `bs`", rAF
 HACn.Execute "truncate `bsart`", rAF
 HACn.Execute "truncate `arzt`", rAF
 HACn.Execute "truncate `fachrichtung`", rAF
 HACn.Execute "truncate `ort`", rAF
 HACn.Execute "truncate `bsart`", rAF
 HACn.Execute "truncate `fremdsprache`", rAF
 HACn.Execute "truncate `genehmigung`", rAF
 HACn.Execute "truncate `nlart`", rAF
 HACn.Execute "truncate `sprechzeiten`", rAF
 HACn.Execute "truncate `titel`", rAF
 HACn.Execute "truncate `vertragsangebot`", rAF
 HACn.Execute "truncate `zusatzbezeichnung`", rAF
 HACn.Execute "truncate `weiterbildung`", rAF
 
 Dim AbschBis&(), AbZ&
 AbZ = 0
 ReDim AbschBis(0)
 Dim aktab&, laktabFA% ' ob die letzte Zeile eine Facharztzeile war -> dann nicht zählen
 For aktab = LBound(splitt) To UBound(splitt)
  Select Case splitt(aktab)
   Case "Fachärztin für Psychosomatische Medizin und", "Facharzt für Psychosomatische Medizin und"
    splitt(aktab) = splitt(aktab) & " " & splitt(aktab + 1)
    splitt(aktab + 1) = ""
  End Select
'  If aktab = 46637 Then Stop
  Select Case splitt(aktab)
' eine der folgenden Ausdrücke steht hinter dem Namen, wodurch dieser als solches klassifizierbar ist
   Case "Facharzt", "Hausarzt", "Fach- und Hausarzt", "Psychologischer Psychotherapeut", "Psychologische Psychotherapeutin", "Psychotherapeutisch tätige Ärztin", "Psychotherapeutisch tätiger Arzt", "Kinder- und Jugendlichen-Psychotherapeutin", "Kinder- und Jugendlichen-Psychotherapeut", "Fachärztin für Psychiatrie und Psychotherapie", "Facharzt für Psychiatrie und Psychotherapie", "Facharzt für Neurologie", "Facharzt für Nervenheilkunde", "Fachärztin für Neurologie", "Fachärztin für Nervenheilkunde", "Fachärztin für Psychosomatische Medizin und Psychotherapie", "Facharzt für Psychosomatische Medizin und Psychotherapie"
    If laktabFA = 0 Then
neuerAbschnitt:
     If AbZ = UBound(AbschBis) Then
      ReDim Preserve AbschBis(UBound(AbschBis) + 500)
     End If
     AbZ = AbZ + 1
     If aktab = 1 Then
      AbschBis(AbZ) = aktab - 2
     ElseIf InStrB(splitt(aktab - 1), " ") = 0 Then ' wenn der Name sich über zwei Zeilen zieht, dann ist in der zweiten Zeile evtl. kein Leerzeichen mehr; Angerpointner
      AbschBis(AbZ) = aktab - 3
      splitt(aktab - 2) = splitt(aktab - 2) & " " & splitt(aktab - 1)
      splitt(aktab - 1) = ""
     Else
      AbschBis(AbZ) = aktab - 2
     End If
'     aktab = aktab + 1 ' sonst wird "Psychotherapie" nicht zusammengezogen, s. Becker-Jakubaß
    End If
    laktabFA = True
   Case "" ' Nach Zusammenzug von "Fachärztin für Psychosomatische Medizin und Psychotherapie"
'    Stop
   Case Else
    If laktabFA = True Then
     If Left$(splitt(aktab), 4) = "LANR" Then
      laktabFA = 0
     End If
' wenn laktabFA = 0, darf normal kein Ort mehr kommen, weil der schon vor LANR kommt.
' Kommt doch noch später einer, so handelt es sich um einen neuen Arzt ohne Niederlassungsart -> Gerald Beier
    ElseIf laktabFA = 0 Then
     If IsNumeric(Left(splitt(aktab), 1)) Then ' weil "like" so lang brauchen soll
      If splitt(aktab) Like "##### *" Then ' PLZ mit Ort
       GoTo neuerAbschnitt
      End If
     End If
    End If ' laktabFA = True / else
  End Select
 Next aktab
 If AbZ = UBound(AbschBis) Then
  ReDim Preserve AbschBis(UBound(AbschBis) + 500)
 End If
 AbZ = AbZ + 1
 AbschBis(AbZ) = UBound(splitt) - 1
 For aktab = 2 To AbZ
  Call proTeilnehmer(splitt, AbschBis, aktab, HACn)
 Next aktab
 lies.Ausgeb (AbZ - 1) & " Ärzte aus '" & Filename & "' eingelesen.", 1
 SplitNeu splitt(0), " ", s2
 For j = 0 To UBound(s2)
  If InStrB(s2(j), "-") > 0 Then
   s2 = Split(s2(j), "-")
   Datum = CDate(s2(0))
   Exit For
  End If
 Next j
 Result = PDDoc.Close
 Set PDPage = Nothing
 Set PDHili = Nothing
 Set PDTextS = Nothing
 Set PDDoc = Nothing
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HausärzteEinlesen_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub

Sub proTeilnehmer(ByRef splitt$(), ByRef AbschBis&(), aktab&, ByRef HACn As ADODB.Connection)
  Dim Uab As UAbTyp, Arzt As ArztTyp, BS As BSTyp, Ort As New CString, SZ As New CString, s2$()
  Dim tel$(), telz&, fax$(), faxz&, mail() As New CString, mailz&
  Dim OrtZeile&, LANRZeile&, szzz&  ' SprechzeitenZeilenZahl
  Dim fren&() ' Fachrichtungen des aktuellen Teilnehmers
  ReDim fren&(0)
  Dim j&, jj&, szz&, buch$, pos&, p2&
  Dim idIns&, idIns2, rAF&, rs As New ADODB.Recordset
  On Error GoTo fehler
' alle Zeilen des aktuellen Teilnehmers
  Uab = Name
  For j = AbschBis(aktab - 1) + 1 To AbschBis(aktab)
'   If Left$(splitt(j), 12) = "Gerald Beier" Then Stop
   Select Case Uab
    Case Name
'     If InStrB(splitt(j), "Massoudy") <> 0 Or InStrB(splitt(j), "Al-Iassin") <> 0 Or InStrB(splitt(j), "Guarch") <> 0 Then Stop
     SplitNeu splitt(j), " ", s2
     If UBound(s2) > 1 Then
      If s2(UBound(s2) - 1) = "Al" Then
       s2(UBound(s2) - 1) = s2(UBound(s2) - 1) & " " & s2(UBound(s2))
       ReDim Preserve s2(UBound(s2) - 1)
      End If
     End If
     Arzt.Nachname = s2(UBound(s2))
     
     If UBound(s2) = 1 Then
      Arzt.Vorname = s2(0)
     Else
      If s2(0) = "Jorj" Or s2(0) = "Roser" Or s2(0) = "Wolf" Or s2(0) = "Karl" Or s2(0) = "Mirjam" Or s2(0) = "Gerald" Or s2(1) = "von" Or s2(1) = "van" Or s2(1) = "van den" Or s2(1) = "v." Then
       s2(1) = s2(0) & " " & s2(1)
       s2(0) = ""
      End If
      Arzt.Vorname = s2(1)
      For jj = 2 To UBound(s2) - 1
       Arzt.Vorname = Arzt.Vorname & " " & s2(jj)
      Next jj
      Do
       pos = InStr(Arzt.Vorname, ". ")
       If pos = 0 Then Exit Do
       s2(0) = s2(0) & Left$(Arzt.Vorname, pos)
       Arzt.Vorname = Mid$(Arzt.Vorname, pos + 2)
      Loop
      pos = InStr(Arzt.Vorname, ")")
      If pos <> 0 Then
       s2(0) = s2(0) & Left$(Arzt.Vorname, pos)
       Arzt.Vorname = Trim$(Mid$(Arzt.Vorname, pos + 1))
      End If
      pos = InStr(Arzt.Vorname, "/")
      If pos <> 0 Then
       p2 = InStr(pos, Arzt.Vorname, " ")
       If p2 <> 0 Then
        s2(0) = s2(0) & Left$(Arzt.Vorname, p2 - 1)
        Arzt.Vorname = Mid$(Arzt.Vorname, p2 + 1)
       End If
      End If
      pos = InStr(Arzt.Vorname, "Beier Perlasberger")
      If pos <> 0 Then
       Arzt.Vorname = Trim(Left$(Arzt.Vorname, pos - 1))
       Arzt.Nachname = "Beier"
      End If
      Arzt.titel_id = indIns(HACn, "titel", "titel", s2(0), "idtitel")
     End If ' rs.bof
     lies.Ausgeb "Bearbeite: " & Arzt.Nachname & ", " & Arzt.Vorname, 0
     DoEvents
    Case NLrt
' die Zeile, in der PLZ und Ort stehen, enthält als erstes eine Ziffer
     For jj = j To AbschBis(aktab) ' j+1 gestrichen wegen Gerald Beier
      If IsNumeric(Left$(splitt(jj), 1)) Then OrtZeile = jj: Exit For
     Next jj
     If OrtZeile >= AbschBis(aktab) Then Stop
' in der Zeile nach dem Namen könnte, muß aber nicht, eine Niederlassungsart stehen
     Select Case splitt(j)
      Case "Facharzt", "Hausarzt", "Fach- und Hausarzt", "Psychologischer Psychotherapeut", "Psychologische Psychotherapeutin"
       Arzt.nlart_id = indIns(HACn, "nlart", "niederlassungsart", splitt(j), "idnlart")
       j = j + 1
     End Select
     Dim jjj&
' über die Zeile getrennte Fachrichtungen zusammenführen
     For jj = j To OrtZeile - 2
      jjj = 0
      While (InStrB(splitt(jj), "(") <> 0 And InStrB(splitt(jj), ")") = 0) Or Right$(splitt(jj), 2) = "/-"
       jjj = jjj + 1
       If Right$(splitt(jj), 1) <> " " Then splitt(jj) = splitt(jj) & " "
       splitt(jj) = splitt(jj) & splitt(jj + jjj)
       splitt(jj + jjj) = ""
      Wend
     Next jj
' Teilgebiete und Schwerpunkte gesondert aufnehmen
     Dim FrSep$(1), ergstr$(), ez&
     FrSep(0) = "Schwerpunkt"
     FrSep(1) = "Teilgebiet"
'     If Arzt.Nachname = "Abelein" And Arzt.Vorname = "Richard" Then Stop
     For jj = j To OrtZeile - 2
      If LenB(splitt(jj)) <> 0 Then
       ez = SplitNeuArr(splitt(jj), FrSep, ergstr, True)
       For jjj = 0 To ez - 1
        ReDim Preserve fren(UBound(fren) + 1)
        buch = Right$(ergstr(jjj), 1)
        If buch = "(" Or buch = ")" Or buch = "," Then
         ergstr(jjj) = Left$(ergstr(jjj), Len(ergstr(jjj)) - 1)
        End If
        fren(UBound(fren)) = indIns(HACn, "fachrichtung", "Fachrichtung", ergstr(jjj), "idFachrichtung")
       Next jjj
      End If
     Next jj
     
     jj = 1
     Do
      buch = Mid$(splitt(OrtZeile - 1), jj, 1)
      If IsNumeric(buch) Then
       BS.Hausnr = Mid$(splitt(OrtZeile - 1), jj)
       Exit Do
      End If
      jj = jj + 1
      If jj > Len(splitt(OrtZeile - 1)) Then Exit Do
     Loop
     BS.Straße = Trim$(Left$(splitt(OrtZeile - 1), jj - 1))
     pos = InStr(BS.Straße, "Perlasberger")
     If pos <> 0 Then
      BS.Straße = Mid(BS.Straße, pos)
     End If
     SplitNeu splitt(OrtZeile), " ", s2
     BS.Plz = s2(0)
     If UBound(s2) > 0 Then
      Ort = s2(1)
      For jj = 1 To UBound(s2) - 1
       Ort.Append " "
       Ort.Append s2(jj)
      Next jj
      BS.Ort_id = indIns(HACn, "ort", "Ort", Ort.Value, "idOrt")
     End If
'     Select Case splitt(j)
'      Case "Facharzt", "Hausarzt" ' danach immer Fachrichtung
'      Case "Psychologischer Psychotherapeut", "Psychologische Psychotherapeutin"
'       ' manchmal allein, manchmal vor dem nächsten
'      Case "Kinder- und Jugendlichen-Psychotherapeutin", "Kinder- und Jugendlichen-Psychotherapeut"
'       ' manchmal allein stehend, manchmal nach dem letzten, danach immer Straße
'      Case "Fachärztin für Psychiatrie und Psychotherapie", "Facharzt für Psychiatrie und Psychotherapie"
'       ' manchmal nach Facharzt, manchmal allein, manchmal vor dem nächsten
'      Case "Psychotherapeutisch tätige Ärztin", "Psychotherapeutisch tätiger Arzt"
'       ' manchmal allein, manchmal nach dem letzten
'       Uab = Uab + 1
'      Case Else
'     End Select
     j = OrtZeile + 1
     telz = -1
     ReDim tel(0)
     Do While Left$(splitt(j), 5) = "Tel.:"
      telz = telz + 1
      If telz > UBound(tel) Then ReDim Preserve tel(telz)
      tel(telz) = Mid$(splitt(j), 7)
      j = j + 1
     Loop
     faxz = -1
     ReDim fax(0)
     Do While Left$(splitt(j), 5) = "Fax.:"
      faxz = faxz + 1
      If faxz > UBound(fax) Then ReDim Preserve fax(faxz)
      fax(faxz) = Mid$(splitt(j), 7)
      j = j + 1
     Loop
     For jj = j + 1 To AbschBis(aktab)
      If splitt(jj) = "Sprechzeiten:" Then
       szz = jj
       Exit For
      End If
     Next jj
     If szz = 0 Then Stop
     mailz = -1
     ReDim mail(0)
     For jj = j To szz - 2
'      If splitt(jj) = "ag" Then Stop
      If False And (Left$(splitt(jj), 8) = "E-Mail.:") Then ' in einer Zeile stand: "E-Mail.:", dann nä Zeile
       mailz = mailz + 1
       If mailz > UBound(mail) Then ReDim Preserve mail(mailz)
       mail(mailz).Append Mid$(splitt(jj), 10)
      End If
'       mail(mailz).Append splitt(jj)
      If Left$(splitt(jj), 8) = "E-Mail.:" Then ' in einer Zeile stand: "E-Mail.:", dann nä Zeile
       mailz = mailz + 1
       If mailz > UBound(mail) Then ReDim Preserve mail(mailz)
       mail(mailz).Append Mid$(splitt(jj), 10)
      Else
       mail(mailz).Append splitt(jj)
      End If
     Next jj
     BS.bsart_id = indIns(HACn, "bsart", "BSArt", splitt(szz - 1), "idbsart")
     j = szz
     For jj = j + 1 To AbschBis(aktab)
      If splitt(jj) = "LANR: BSNR:" Then LANRZeile = jj: BS.obNBS = False: Exit For
      If splitt(jj) = "LANR: NBSNR:" Then LANRZeile = jj: BS.obNBS = 1: Exit For
     Next jj
     szzz = (LANRZeile - j - 1) * 0.5
     SZ.Clear
     For jj = j + 1 To j + szzz + 1
      If jj + szzz >= LANRZeile Then Exit For
      SZ.Append splitt(jj)
      If jj + szzz >= LANRZeile Or szzz = 0 Then Exit For
      SZ.Append " "
      SZ.Append splitt(jj + szzz)
      If jj < j + 1 + szzz Then SZ.Append ", "
     Next jj
'     If Left$(SZ.Value, 6) <> "Montag" And Left$(SZ.Value, 8) <> "Dienstag" And Left$(SZ.Value, 8) <> "Mittwoch" And Left$(SZ.Value, 10) <> "Donnerstag" And Left$(SZ.Value, 7) <> "Freitag" And SZ.Value <> "keine Angaben" Then Stop
     BS.sprechzeiten_id = indIns(HACn, "sprechzeiten", "Sprechzeiten", doUmwfSQL(SZ.Value, obMySQL), "idsprechzeiten")
     SplitNeu splitt(LANRZeile + 1), " ", s2
     If UBound(s2) <> 1 Then Stop
     Arzt.LANR = s2(0)
'     If BS.obNBS = 1 Then BS.NBSNR = s2(1): BS.BSNR = 0 Else BS.NBSNR = 0: BS.BSNR = s2(1)
     BS.BSNR = s2(1)
     Dim KennZeile&(Zusatzbezeichnung To ÄrztederPraxis)
     Dim ZB() As New CString, Wb() As New CString, Geneh() As New CString, VertA() As New CString, fS() As New CString, ÄP() As New CString
     Dim ZBz&, WBz&, Genehz&, VertAz&, FSz&, ÄPz&, überspringen&
     For jj = LANRZeile To AbschBis(aktab)
      Select Case splitt(jj)
       Case "Zusatzbezeichnungen:":        KennZeile(Zusatzbezeichnung) = jj: Uab = Zusatzbezeichnung
       Case "Weiterbildungen:":            KennZeile(Weiterbildung) = jj: Uab = Weiterbildung
       Case "Genehmigungen:":              KennZeile(Genehmigung) = jj: Uab = Genehmigung
       Case "Besondere Vertragsangebote:": KennZeile(Vertragsangebot) = jj: Uab = Vertragsangebot
       Case "Fremdsprachen:":              KennZeile(Fremdsprache) = jj: Uab = Fremdsprache
       Case "Ärzte der Praxis:":           KennZeile(ÄrztederPraxis) = jj: Uab = ÄrztederPraxis
       Case "Rollstuhlgerechte Praxis":    BS.Rollst = 1
       Case Else
        überspringen = 0
        Do While jj <= AbschBis(aktab) And (Right$(splitt(jj), 3) = "und" Or Right$(splitt(jj), 2) = "u." Or Right$(splitt(jj), 1) = ";" Or Left$(splitt(jj + 1 + überspringen), 4) = "bzw." Or Right$(splitt(jj), 2) = "f." Or Right$(splitt(jj), 3) = "mit" Or Left$(splitt(jj + 1 + überspringen), 1) = "(" Or Left$(splitt(jj + 1 + überspringen), 1) = "/" Or Right$(splitt(jj), 1) = "/" Or Right$(splitt(jj), 1) = "-" Or Right$(splitt(jj), 1) = "," Or Right$(splitt(jj), 5) = "sches" Or Left$(splitt(jj + 1 + überspringen), 3) = "SGB" Or Left$(splitt(jj + 1 + überspringen), 14) = "sonographisch)" Or Left$(splitt(jj + 1 + überspringen), 8) = "mammogr." Or Right$(splitt(jj), 5) = "spez." Or Left$(splitt(jj + 1 + überspringen), 3) = "IPM" Or Left$(splitt(jj + 1 + überspringen), 18) = "Kinder/Jugendliche" Or Left$(splitt(jj + 1 + überspringen), 12) = "Operationen)" Or Right$(splitt(jj), 6) = "in der" Or splitt(jj) = "DMP Brustkrebs teilnehmender Facharzt")
         überspringen = überspringen + 1
'         If Uab <> ÄrztederPraxis Then Stop
         splitt(jj) = splitt(jj) & " " & splitt(jj + überspringen)
        Loop
        Select Case Uab
         Case Zusatzbezeichnung
          ZBz = ZBz + 1
          If ZBz = 1 Then ReDim ZB(60) Else If ZBz > UBound(ZB) Then ReDim Preserve ZB(UBound(ZB) + 60)
          ZB(ZBz) = splitt(jj)
         Case Weiterbildung
          WBz = WBz + 1
          If WBz = 1 Then ReDim Wb(60) Else If WBz > UBound(Wb) Then ReDim Preserve Wb(UBound(Wb) + 60)
          Wb(WBz) = splitt(jj)
         Case Genehmigung
          Genehz = Genehz + 1
          If Genehz = 1 Then ReDim Geneh(60) Else If Genehz > UBound(Geneh) Then ReDim Preserve Geneh(UBound(Geneh) + 60)
          Geneh(Genehz) = splitt(jj)
         Case Vertragsangebot
          VertAz = VertAz + 1
          If VertAz = 1 Then ReDim VertA(60) Else If VertAz > UBound(VertA) Then ReDim Preserve VertA(UBound(VertA) + 60)
          VertA(VertAz) = splitt(jj)
         Case Fremdsprache
          FSz = FSz + 1
          If FSz = 1 Then ReDim fS(60) Else If FSz > UBound(fS) Then ReDim Preserve fS(UBound(fS) + 60)
'          If Not FS(FSz + 1) Is Nothing Then If FS(FSz + 1).Length <> 0 Then Stop
          fS(FSz) = splitt(jj)
'         For jjj = 1 To FSz - 1
'           If FS(jjj) = FS(FSz) Then Stop
'          Next jjj
         Case ÄrztederPraxis
          ÄPz = ÄPz + 1
          If ÄPz = 1 Then ReDim ÄP(60) Else If ÄPz > UBound(ÄP) Then ReDim Preserve ÄP(UBound(ÄP) + 60)
          ÄP(ÄPz) = splitt(jj)
        End Select
      End Select
      jj = jj + überspringen
     Next jj
     Set rs = Nothing
     rs.Open "select idarzt from arzt where lanr = " & Arzt.LANR, HACn, adOpenStatic, adLockReadOnly
     If rs.BOF Then
      If Arzt.titel_id = 0 Then Arzt.titel_id = indIns(HACn, "titel", "titel", "", "idtitel")
      If Arzt.nlart_id = 0 Then Arzt.nlart_id = indIns(HACn, "nlart", "niederlassungsart", "", "idnlart")
      HACn.Execute "insert into `arzt`(`Nachname`,`Vorname`,`titel_id`,`LANR`,`nlart_id`) values('" & doUmwfSQL(Arzt.Nachname, obMySQL) & "','" & doUmwfSQL(Arzt.Vorname, obMySQL) & "'," & Arzt.titel_id & "," & Arzt.LANR & "," & Arzt.nlart_id & ")"
      Set rs = HACn.Execute("select last_insert_id()")
      idIns2 = rs.Fields(0)
      For jj = 1 To ZBz
       idIns = indIns(HACn, "zusatzbezeichnung", "Zusatzbezeichnung", ZB(jj).Value, "idZusatzbezeichnung")
       HACn.Execute "insert into `arzt_has_zusatzbezeichnung`(`arzt_id`,`Zusatzbezeichnung_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      For jj = 1 To WBz
       idIns = indIns(HACn, "weiterbildung", "Weiterbildung", Wb(jj).Value, "idWeiterbildung")
       HACn.Execute "insert into `arzt_has_weiterbildung`(`arzt_id`,`Weiterbildung_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      For jj = 1 To Genehz
       idIns = indIns(HACn, "genehmigung", "genehmigung", Geneh(jj).Value, "idgenehmigung")
       HACn.Execute "insert into `arzt_has_genehmigung`(`arzt_id`,`genehmigung_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      For jj = 1 To VertAz
       idIns = indIns(HACn, "vertragsangebot", "vertragsangebot", VertA(jj).Value, "idvertragsangebot")
       HACn.Execute "insert into `arzt_has_vertragsangebot`(`arzt_id`,`vertragsangebot_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      For jj = 1 To FSz
       idIns = indIns(HACn, "fremdsprache", "fremdsprache", fS(jj).Value, "idfremdsprache")
       HACn.Execute "insert into `arzt_has_fremdsprache`(`arzt_id`,`fremdsprache_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      For jj = 1 To UBound(fren)
       ' Fachrichtungen können offenbar versehentlich auch doppelt genannt werden (-> Eckhard Rudolf, hat 2 x Unfallchirurgie und 2 x Viszeralchirurgie)
       Set rs = Nothing
       rs.Open "select arzt_id from `arzt_has_fachrichtung` where arzt_id = " & idIns2 & " and fachrichtung_id = " & fren(jj), HACn, adOpenStatic, adLockReadOnly
       If rs.BOF Then
        HACn.Execute "insert into `arzt_has_fachrichtung`(`arzt_id`,`fachrichtung_id`) values(" & idIns2 & "," & fren(jj) & ")"
       End If
      Next jj
     Else
      idIns2 = rs.Fields(0) ' für arzt_has_bs
     End If
   End Select
   Select Case Uab
    Case Name, NLrt
     Uab = Uab + 1
   End Select
  Next j
'  HACn.Execute "insert into `arzt`(nachname,vorname,titel_id,fachrichtung_id,lanr,nlart_id) values('" & Arzt.Nachname & "','" & Arzt.Vorname & "'," & Arzt.titel_id & "," & Arzt.fachrichtung_id & "," & Arzt.LANR & " ," & Arzt.nlart_id & ") on duplicate key update nachname = '" & Arzt.Nachname & "', vorname = '" & Arzt.Vorname & "', titel_id = " & Arzt.titel_id & ", fachrichtung_id = " & Arzt.fachrichtung_id & ", nlart_id = " & Arzt.nlart_id, rAF
  Set rs = Nothing
  rs.Open "select idbs from `bs` where bsnr = " & BS.BSNR, HACn, adOpenStatic, adLockReadOnly
  If rs.BOF Then
   HACn.Execute "insert into `bs`(straße,hausnr,plz,ort_id,bsnr,bsart_id,sprechzeiten_id,rollst) values('" & doUmwfSQL(BS.Straße.Value, obMySQL) & "','" & BS.Hausnr & "','" & BS.Plz & "'," & BS.Ort_id & ",'" & BS.BSNR & "'," & BS.bsart_id & "," & BS.sprechzeiten_id & "," & BS.Rollst & ")"
   Set rs = HACn.Execute("select last_insert_id()")
  End If
  idIns = rs.Fields(0)
  For jj = 0 To telz
   HACn.Execute "insert into `tel`(`Tel`,`bs_id`) values('" & tel(jj) & "'," & idIns & ")", rAF
  Next jj
  For jj = 0 To faxz
   HACn.Execute "insert into `fax`(`Fax`,`bs_id`) values('" & fax(jj) & "'," & idIns & ")", rAF
  Next jj
  For jj = 0 To mailz
   HACn.Execute "insert into `mail`(`mail`,`bs_id`) values('" & mail(jj) & "'," & idIns & ")", rAF
  Next jj
  
  Set rs = Nothing ' da ein Arzt versehentlich doppelt drin stehen kann (Georg Hochheuser)
  rs.Open "select obneben from `arzt_has_bs` where `bs_id` = " & idIns & " and `arzt_id` = " & idIns2, HACn, adOpenStatic, adLockReadOnly
  If rs.BOF Then
   HACn.Execute "insert into `arzt_has_bs`(`bs_id`,`arzt_id`,`obneben`) values(" & idIns & "," & idIns2 & "," & BS.obNBS & ")", rAF
  End If
' Plausibilitätskontrollen:
'select * from mail where not mail rlike '^[\-\_\.[:alnum:]]+@[\-\_\.[:alnum:]]+\.(org|de|com|ag|info|net|biz)+$';
'select * from tel where not tel rlike '^[-0123456789]+$';
'select * from fax where not fax rlike '^[-0123456789]+$';
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in proTeilnehmer/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' proTeilnehmer

Private Sub HausärztemitalterKVNrergänzen_Click()
 Dim fha As New frmalthae, rAF&, rs As New ADODB.Recordset
 ' erst mal die leeren KV-Nummern einfügen
 DBCn.Execute "insert into althae (kvnu,kvnr) select kvnu, kvnr from (select n.kvnr kvnu, left(n.kvnr,2),'/',right(n.kvnr,5) kvnr, HAName hHA, concat_ws(', ',l.name, l.vorname) lHA from aktfv f left join namen n on f.pat_id = n.pat_id left join listenausgabeuew l on n.kvnr = l.kvnr left join althae h on n.kvnr = h.kvnu group by n.pat_id) innen where (isnull(lha) or lha='') and (isnull(hha) or hha='') and kvnu <> '' and not exists (select kvnu from althae where kvnu = innen.kvnu)", rAF
' Set fha.datprimaryRS = n
' Call fha.vorbereit
 rs.Open "select group_concat(kvnu) nrn from (select n.kvnr kvnu, HAName hHA, concat_ws(', ',l.name, l.vorname) lHA from aktfv f left join namen n on f.pat_id = n.pat_id left join listenausgabeuew l on n.kvnr = l.kvnr left join althae h on n.kvnr = h.kvnu group by n.pat_id) innen where (isnull(lha) or lha='') and (isnull(hha) or hha='') and kvnu <> ''", DBCn, adOpenStatic, adLockReadOnly
 If LenB(rs!nrn) <> 0 Then
  If fha.Vorbereit(rs!nrn) Then
   fha.Show
  End If
 Else
  MsgBox "Keine Datensätze zu editieren!"
 End If
 Exit Sub
End Sub

Private Sub korrQD_Click() ' Quelldatum für alle korrigieren
 Dim rs As New ADODB.Recordset, nQD As Date, rAF&, rsum&
 Call ProgStart
 rs.Open "select * from dokumente", DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  nQD = doQuelldatum(rs!DokName)
  If nQD <> rs!Quelldatum Then
   DBCn.Execute "update dokumente set quelldatum = " & DatForm(nQD) & " where pat_id = " & rs!Pat_id & " and fid = " & rs!FID & " and dokpfad = '" & doUmwfSQL(rs!DokPfad, lies.obMySQL) & "'", rAF
   If rAF <> 1 Then
    rsum = rsum & rAF
   End If
  End If
  rs.Move 1
 Loop
 Ausgeb rsum & " Quelldaten korrigiert", True
 Debug.Print "Fertig"
 Beep 1000, 1000
End Sub

Private Sub DokumenteAbgehaktPrüfen_Click()
 Dim rs As New ADODB.Recordset, n&, n1&
 On Error GoTo fehler
 Dim ErgDat$, erg$
 ErgDat = uVerz & "FD.txt"
 Do
  erg = Dir(ErgDat)
  If LenB(erg) = 0 Then Exit Do
  ErgDat = Replace$(ErgDat, "FD", "FD1")
 Loop
 Open ErgDat For Output As #7
 Call ProgStart
 rs.Open "SELECT dokname, da.*, d.* FROM `dokumente abgehakt` da left join dokumente d using (dokpfad) order by d.pat_id, d.zeitpunkt", DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  Dim eDat As Date
  If Not IsNull(rs!DokName) Then
   eDat = GetDatumAusString(rs!DokName)
'   Debug.Print eDat & ", " & rs!DokName
   If eDat = 0 Then
    Debug.Print n, rs!Pat_id, rs!Zeitpunkt, rs!DokName
    Print #7, n, rs!Pat_id, rs!Zeitpunkt, rs!DokName
    n = n + 1
   Else
    n1 = n1 + 1
   End If
  End If
  rs.Move 1
 Loop
 Debug.Print "Richtige:"; n1
 Print #7, "Richtige:", n1
 Close #7
 Shell Environ("systemroot") & "\system32\notepad.exe " & ErgDat, vbMaximizedFocus
 Ende
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DokumenteAbgehaktPrüfen_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DokumenteAbgehaktPrüfen_Click

Private Sub PathLabAnschau_Click()
 Dim ePL As New PathLabForm
 Set ePL.eLese = Me
 ePL.Show
End Sub
Private Sub Gestationsdiabetikerinnen_Click()
' SELECT f.pat_id,f.fid, left(concat(a.nachname,' ',a.vorname),20) as name, date_format(a.gebdat,'%d.%m.%y') as geb, d.icd, d.diagsicherheit as dsi, diabetestyp FROM aktfv f left join diagnosen d on f.fid = d.fid and icd = 'O24.4' and diagsicherheit not in ('A','Z') left join anamnesebogen a on f.pat_id = a.pat_id where (not isnull(icd) or a.diabetestyp = 'g');
 Dim rs As New ADODB.Recordset, rez As New ADODB.Recordset, Pat_id&, AusgStr$, DatNam$, TA1$, STA1$(), i&
 Dim SpMin%(2)
 SpMin%(0) = 6
 Call ProgStart
 DatNam = pVerz & "Gestationsdiabetikerinnen " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".txt"
' Open DatNam For Output As #327
 rs.Open "SELECT f.pat_id,f.fid, left(concat(a.nachname,' ',a.vorname),20) as name, date_format(a.gebdat,'%d.%m.%y') as geb, d.icd, d.diagsicherheit as dsi, diabetestyp FROM aktfv f left join diagnosen d on f.fid = d.fid and icd = 'O24.4' and diagsicherheit not in ('A','Z') left join anamnesebogen a on f.pat_id = a.pat_id where (not isnull(icd) or a.diabetestyp = 'g')", DBCn, adOpenStatic, adLockReadOnly
 TabAusgeb rs, True, , , , , , DatNam
End Sub

Private Sub PatientenlistefürHausarztmodell_Click()
 Dim rs As New ADODB.Recordset
 rs.Open "select n.kvnr `KV-Nr.HA`, if(l.name <> '' or l.vorname <> '' and not isnull(l.name),concat_ws(',',l.name, left(l.vorname,1), concat(if(l.telefon<>'','T.',''),l.telefon)), if(h.nachname <> '' and not isnull(h.nachname),concat_ws(',',h.nachname,left(h.vorname,1),concat(if(h.tel1<>'','T.',''),h.tel1)),'?')) Hausarzt, n.pat_id Pat_ID, concat_ws(',*',concat_ws(',', n.nachname, n.vorname),date_format(gebdat,'%d.%m.%y')) Patient, mid(n.notiz,instr(n.notiz,'HM '),if(instr(n.notiz,'PG ')<>0,instr(n.notiz,'PG ')-2-instr(n.notiz,'HM '),length(n.notiz))) Hausarztmodell from aktfv f left join namen n on f.pat_id = n.pat_id left join listenausgabeuew l on n.kvnr = l.kvnr left join althae h on n.kvnr = h.kvnu group by pat_id order by n.kvnr, patient", DBCn, adOpenStatic, adLockReadOnly
 TabAusgeb rs, True, , , , , True, "p:\PatientenlisteFürHausarztmodell.txt"
End Sub
Private Sub PatientenlistefürVollpauschale_Click()
 Dim rs As New ADODB.Recordset
 rs.Open "select f.pat_id, concat_ws(',', f.nachname, f.vorname) Pat, vpau, f.übwlanr, concat_ws(',',arzt.nachname,arzt.vorname) ÜW, if(nlart.niederlassungsart='Hausarzt','h',left(nlart.niederlassungsart,1)) NLArt, arzt2.lanr LANR2, concat_ws(',',arzt2.nachname,arzt2.vorname) ÜW2, if(nlart2.niederlassungsart='Hausarzt','h',left(nlart2.niederlassungsart,1)) nlart2 from (aktfv join faelle f on aktfv.fid = f.fid and f.schgr between 20 and 29) left join (hausaerzte.arzt join hausaerzte.nlart on arzt.nlart_id = nlart.idnlart left join (hausaerzte.arzt_has_bs ahb left join hausaerzte.arzt_has_bs ahb2 on ahb.bs_id = ahb2.bs_id and ahb.arzt_id <> ahb2.arzt_id join  (hausaerzte.arzt arzt2 join hausaerzte.nlart nlart2 on arzt2.nlart_id = nlart2.idnlart) on ahb2.arzt_id = arzt2.idarzt   ) on arzt.idarzt = ahb.arzt_id and nlart2.Niederlassungsart = 'Facharzt' and nlart.niederlassungsart = 'Hausarzt') on übwlanr = arzt.lanr" & _
         " left join (select aktfv.pat_id, leistung vpau from aktfv join leistungen l on aktfv.fid = l.fid and (leistung like '031%' or leistung like '01210')) vpau on f.pat_id = vpau.pat_id where nlart.niederlassungsart = 'Facharzt' or isnull(nlart.niederlassungsart) or not isnull(arzt2.lanr) group by pat_id", DBCn, adOpenStatic, adLockReadOnly
 TabAusgeb rs, True, , , , , True, "p:\PatientenlisteFürVollpauschale.txt"
End Sub

Private Sub Pumpenträgerliste_Click()
 Dim rs As New ADODB.Recordset, rez As New ADODB.Recordset, Pat_id&, AusgStr$, DatNam$, TA1$, STA1$(), i&
 Dim SpMin%(2)
 SpMin%(0) = 6
 Call ProgStart
 DatNam = pVerz & "Pumpenträger " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".txt"
 Open DatNam For Output As #326
 ' für Acrobat Querdruck 70%
 rs.Open "select a.pat_id, left(concat(a.nachname,',',a.vorname,if(a.titel='','',','),a.titel,if(a.nvorsatz='','',' '),a.nvorsatz,' (',a.anrede,')'),24) as name, date_format(a.gebdat,'%d.%m.%y') as geb, left(concat(`diabetes seit`,' ',a.`insulin seit`),12) as 'D.m./Ins.', left(concat(if(a.insulinpumpe=1,'+','-'),' ',a.`insulinpumpe seit`,' ',a.`insulinpumpe marke`),24) as 'Pumpe b.Anamn./seit/Marke', left(a.ther1,4) as Ther1, left(a.dmp,9) as DMP, f.schgr as SG, date_format(f.bhfb,'%d.%m.%y') as bhfb, left(concat(privattel, '|',privatmobil,'|',diensttel,'|',email,'|',privatfax,'|',privattel_2),60) as kontakt from anamnesebogen a left join namen n on a.pat_id = n.pat_id left join lfaellev f on a.pat_id = f.pat_id where therakt = 'CSII' and tkz = 0", DBCn, adOpenStatic, adLockReadOnly
 TA1 = TabAusgeb(rs, True).Value
 Print #326, TA1
 STA1 = Split(TA1, vbCrLf)
 Print #326, ''
 Print #326, ''
 i = 1
 Do While Not rs.EOF
 ' pumpentr
  Set rez = Nothing
  Pat_id = rs!Pat_id
  sql = forminhalt & " where form_abk = 'lar' and feld in ('medikament','txtmedKey') and " & _
        "(feldinh like ('%Rapid D Link%') or feldinh like ('%Rap D Li%') or feldinh like ('%Rapid-D Li%') or feldinh like ('%Check Spirit%') or feldinh like ('%Chek Spirit%') or feldinh like ('%Pumpenträg%')or feldinh like ('%Kunststoffampu%')or feldinh like ('%Spritzampull%')or feldinh like ('%batteriefachdeckel%')or feldinh like ('%H-Tron%') or feldinh like ('%D-Tron%') or feldinh like ('%Paradigm%') or feldinh like ('%CSII%') or feldinh like ('%linpumpe%')) " & _
        "and zeitpunkt > " & DatForm(rs!BhFB - 640) & " and pat_id = " & Pat_id & " order by zeitpunkt desc limit 10"
  rez.Open "select Pat_ID, Zeitpunkt, Feldinh from (" & sql & ") i", DBCn, adOpenStatic, adLockReadOnly
  Print #326, STA1(i)
  Print #326, TabAusgeb(rEinl:=rez, obMitausgeb:=False, ohneÜbers:=True, SpMinÜ:=SpMin).Value
'  Do While Not rez.EOF
'   rez.Move 1
'  Loop
  rs.Move 1
  i = i + 1
 Loop
 Close #326
 Shell Environ("systemroot") & "\system32\notepad.exe " & DatNam, vbMaximizedFocus
End Sub

' select count(0),leistung from leistungen where leistung in ('97261B','92261A',
'92292C','92292D','92292E','97268','97269','97271','97274','97278','97279','92292A','92292B','92290A','92291A') and year(zeitpunkt)= '2008' group by leistung;

Private Sub Schulungsstatistik_Click()
 Dim col As New Collection, el, rs As New ADODB.Recordset, ausg$, DatNam$
 DatNam = pVerz & "Schulungsstatistik " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".csv"
 Open DatNam For Output As #325
 Print #325, "Leistung Titel                                               Anzahl in " & Year(Now() - 15)
 col.Add "97261B"
 col.Add "92261A"
 col.Add "97262B"
 col.Add "92262A"
 col.Add "97263B"
 col.Add "92263A"
 col.Add "97269B"
 col.Add "92269A"
 col.Add "97264B"
 col.Add "92264A"
 col.Add "97265B"
 col.Add "92292C"
 col.Add "92265A"
 col.Add "97266B"
 col.Add "92292D"
 col.Add "92266A"
 col.Add "97268B"
 col.Add "92292E"
 col.Add "92268A"
 col.Add "97268"
 col.Add "92292A"
 col.Add "97269"
 col.Add "97271"
 col.Add "97274"
 col.Add "92292B"
 col.Add "97278"
 col.Add "92290A"
 col.Add "97279"
 col.Add "92291A"
 
 Call ProgStart
 For Each el In col
  Set rs = Nothing
  rs.Open "select leistung,titel,(select count(0) from leistungen where leistung = e.leistung and year(zeitpunkt) = year(subdate(now(),interval 15 day))) ct from ebm2000plus e where leistung = '" & el & "'", DBCn, adOpenStatic, adLockReadOnly
  If rs.EOF Then rs.Close: rs.Open "select count(0) ct, titel from leistungen l left join ebm2000plus e on l.leistung = e.leistung where l.leistung = '" & el & "' and year(zeitpunkt) = year(subdate(now(),interval 15 day))", DBCn, adOpenStatic, adLockReadOnly
  If el = "97278" Then
   MsgBox "Stop in Schulungsstatistik_Click:" & vbCrLf & "el = '97278'"
   Stop
  End If
  ausg = Left$(el & Space(7), 7) & "(" & Left(rs!Titel & Space$(60), 60) & "): " & rs!ct
  Print #325, ausg
 Next el
 MsgBox "Fertig!"
 Close #325
 Shell Environ("systemroot") & "\system32\notepad.exe " & DatNam, vbMaximizedFocus
End Sub

Private Sub Motivationsgesprächskandidaten_Click()
 Dim rv As New ADODB.Recordset, rs As New ADODB.Recordset, DatNam$, i&, ausg$, TA1$, SpMax%(5), fristS$, sql$
 Call ProgStart
 rv.Open "show create view aktfv", DBCn, adOpenStatic, adLockReadOnly
 fristS = rv.Fields(1)
 Set rv = Nothing
 fristS = Mid(fristS, InStr(fristS, "interval ") + 9)
 fristS = Left$(fristS, InStr(fristS, " ") - 1)
 If Not IsNumeric(fristS) Then
  MsgBox "Ungeeignete Abfrage Aktfv, evtl. erst Views erstellen"
  Exit Sub
 End If
 SpMax(1) = 32
 SpMax(5) = 300
 DatNam = pVerz & "Motivationsgesprächskandidaten " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".csv"
' Print #325, "Lfdnr. Pat_id Name                          ICD"
' rs.Open "select f.pat_id,gesname, icd from aktfv f left join diagnosen d on f.pat_id = d.pat_id and icd regexp '^E1[01234]' and diagsicherheit <> 'A' left join leistungen l on f.pat_id = l.pat_id and (leistung like '972%' or leistung like '922%' or leistung = '92282' or leistung = '92278') and not leistung in ('97272','97276','97277') and year(zeitpunkt) = year(subdate(now(),interval 20 day)) and adddate(zeitpunkt,interval 365 day) > now() where not isnull(id1) and isnull(leistung) group by f.pat_id", DBCn, adOpenStatic, adLockReadOnly
 sql = "select f.pat_id as Pat_ID,gesname as Name, ICD, date(e.zeitpunkt) as Zeitpunkt, e.art as Art, e.inhalt as Inhalt from aktfv f left join diagnosen d on f.pat_id = d.pat_id and icd regexp '^E1[01234]' and diagsicherheit <> 'A' left join leistungen l on f.pat_id = l.pat_id and (leistung like '972%' or leistung like '922%' or leistung = '92282' or leistung = '92278') and not leistung in ('97272','97276','97277') and year(zeitpunkt) = year(subdate(now(),interval " & fristS & " day)) " & _
         "and adddate(zeitpunkt,interval 365 day) > now() left join eintraege e on f.pat_id = e.pat_id and e.art in ('andm','wr','tk','gs') and e.zeitpunkt between str_to_date(concat(year(subdate(now(),interval " & fristS & " day)),'/',((month(subdate(now(),interval " & fristS & " day))-1) div 3)*3+1,'/1'),'%Y/%m/%d') and subdate(adddate(str_to_date(concat(year(subdate(now(),interval " & fristS & " day)),'/',((month(subdate(now(),interval 20 day))-1) div 3)*3+1,'/1'),'%Y/%m/%d'),interval 3 month),interval 1 day) where not isnull(id1) and isnull(l.leistung) group by f.pat_id, art, zeitpunkt"
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 TA1 = TabAusgeb(rs, , , , , SpMax, True, DatNam)
' Open DatNam For Output As #325
' Print #325, TA1
' i = 0
' Do While Not rs.EOF
'  i = i + 1
'  ausg = Right("    " & i, 4) & " " & Right("     " & rs!Pat_id, 5) & " " & Left(rs!GesName & Space$(30), 30) & " " & rs!ICD
'  Print #325, ausg
'  rs.Move 1
' Loop
' Close #325
 MsgBox "Fertig!"
 Call ProgEnde
 Shell Environ("systemroot") & "\system32\notepad.exe " & DatNam, vbMaximizedFocus
End Sub ' Sub Motivationsgesprächskandidaten_Click()

Private Sub Überweiserstatistik_Click()
 Dim rs As New ADODB.Recordset, DatNam$, i&, ausg$
 DatNam = pVerz & "Überweiserstatistik " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".csv"
 Open DatNam For Output As #326
 Call ProgStart
 rs.Open "select kvnu,anrede, haname,plz,ort,tel1,tel2,fax1,fax2,zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,gelöscht,ct from (select count(0) as ct, left(übwv,7) as kvnu from faelle where bhfb > '" & DatForm(Now - 365) & "' and übwv <> '' group by übwv union select count(0) as ct, left(andüw,7) as kvnu from faelle where bhfb > '" & DatForm(Now - 365) & "' and andüw <> '' group by andüw) as i left join kvaerzte.hae using (kvnu) where not gelöscht and not isnull(kvnu) and kvnu <> '6419153' order by ct desc", DBCn, adOpenStatic, adLockReadOnly
 For i = 0 To rs.Fields.Count - 1
  ausg = ausg & """" & rs.Fields(i).Name & """;"
 Next
 ausg = Left(ausg, Len(ausg) - 1)
 Print #326, ausg
 Do While Not rs.EOF
  ausg = ""
  For i = 0 To rs.Fields.Count - 1
   ausg = ausg & """" & rs.Fields(i).Value & """;"
  Next
  ausg = Left(ausg, Len(ausg) - 1)
  Print #326, ausg
  rs.Move 1
 Loop
 MsgBox "Fertig!"
 Close #326
 Call ProgEnde
 Shell Environ("systemroot") & "\system32\notepad.exe " & DatNam, vbMaximizedFocus
End Sub
' select kvnu,anrede, haname,plz,ort,tel1,tel2,fax1,fax2,zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,gelöscht,ct from (select count(0) as ct, left(übwv,7) as kvnu from faelle where bhfb > '2007-09-30' and übwv <> '' group by übwv union select count(0) as ct, left(andüw,7) as kvnu from faelle where bhfb > '2007-09-30' and andüw <> '' group by andüw) as i left join kvaerzte.hae using (kvnu) where not gelöscht and not isnull(kvnu) and kvnu <> '6419153' order by ct desc;

Private Sub DoppelteDiagnosen_Click()
 Dim rs As New ADODB.Recordset, DatNam$, i&
 DatNam = pVerz & "DoppelteDiagnosen " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".txt"
 Open DatNam For Output As #327
 Print #327, "Nr. Zahl Pat_id Name                        ICD Sicherheit -Seite -text"
 Call ProgStart
 rs.Open "select * from (select count(0) as ct, pat_id, icd, gesname, diagsicherheit si, diagtext tx, diagseite se from diagnosen where obdauer <> 0 group by pat_id, icd, diagsicherheit, diagseite) as i where ct > 1 order by pat_id desc", DBCn, adOpenStatic, adLockReadOnly
 i = 0
 Do While Not rs.EOF
  i = i + 1
  Print #327, Right("    " & i, 4) & " " & Left$(rs!ct & "  ", 2) & " " & Right$("     " & rs!Pat_id, 5) & " " & Left$(rs!GesName & Space$(30), 30) & " " & Left$(rs!ICD & "      ", 6) & " " & rs!SI & " " & rs!SE & " " & Left$(rs!tx & Space$(50), 50)
  rs.Move 1
 Loop
 Print #327, "Fertig!"
 Debug.Print "Fertig!"
 Close #327
 Call ProgEnde
 Shell Environ("systemroot") & "\system32\notepad.exe " & DatNam, vbMaximizedFocus
End Sub

Private Sub doppelteFAxe_Click()
 Dim DatNam$
 DatNam = pVerz & "schongefaxte " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".txt"
 Open DatNam For Output As #323
 Call acon(quelleT)
 Call acon(FaxT)
 Call dodoppeltefaxe(pVerz & "unkorrigiert")
 Print #323, "Fertig!"
 Debug.Print "Fertig!"
 Close #323
 Shell Environ("systemroot") & "\system32\notepad.exe " & DatNam, vbMaximizedFocus
End Sub ' doppeltefaxe
Private Sub dodoppeltefaxe(V$)
 Static FSO As New FileSystemObject
 Dim Fil As File, Pid$, pos%, buch$
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
 Print #323, V
 Debug.Print V
 For Each Fil In FSO.GetFolder(V).Files
  If Fil.Name Like "*PID *" Then
   pos = InStr(Fil.Name, "PID ") + 4
   Pid = ""
   Do
    buch = Mid$(Fil.Name, pos, 1)
    If IsNumeric(buch) Then
     Pid = Pid & buch
    Else
     Exit Do
    End If
    pos = pos + 1
   Loop
   If Pid <> "" Then
    Pid = Pid & ","
    Set rs = Nothing
    Call rs.Open("select * from outg where docn like '%PID " & Pid & "%' And docn Like '%Arztbrief%'", FxCn, adOpenStatic, adLockReadOnly)
    If Not rs.EOF Then
     If FSO.FileExists(pVerz & rs!docn) Then
      Print #323, "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe
      Debug.Print "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe
      Kill Fil.path
     Else
      Set rs1 = Nothing
      Call rs1.Open("select * from briefe where name like '%Arztbrief%' and name like '%PID " & Pid & "%'", DBCn, adOpenStatic, adLockReadOnly)
      If Not rs1.EOF Then
       Print #323, "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe, rs1!Name
       Debug.Print "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe, rs1!Name
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
 Dim Fil As File, Pid&, pos&, p2&, rs As New ADODB.Recordset, an As New ADODB.Recordset, Infos$(), rD As New ADODB.Recordset
 Call ProgStart
 For Each Fil In FSO.GetFolder(pVerz & "unkorrigiert").Files
  pos = InStr(Fil.Name, "PID ")
  If pos > 0 Then
   p2 = InStr(pos, Fil.Name, " ")
   Pid = Mid$(Fil.Name, pos + 4, p2 - pos + 1)
   Set rs = Nothing
   Call rs.Open("select * from faelle where pat_id = " & Pid & " order by bhfb desc;", DBCn, adOpenStatic, adLockOptimistic)
   
   Set an = Nothing
'   Call an.Open("select -tkz as j_tkz, a.* from anamnesebogen a where pat_id = " & pid, DBCn, adOpenStatic, adLockOptimistic)
   Call an.Open("select * from anamnesebogen a where pat_id = " & Pid, DBCn, adOpenStatic, adLockOptimistic)
   If rs!SchGr = 41 Or rs!SchGr = 42 Or rs!SchGr = 43 Then
    FSO.DeleteFile (Fil.path)
'   ElseIf an!j_tkz <> 0 Then
   ElseIf an!Tkz <> 0 Then
'
    FSO.DeleteFile (Fil.path)
   Else
     Call getHausarzt(Pid, Infos())
     If Infos(1, 0) = "" Then
      FSO.DeleteFile (Fil.path)
     ElseIf Infos(1, 0) Like "*Schade" Then
      Debug.Print Fil.Name
      MsgBox "Stop in falschebriefelöschen_Click:" & "infos(1,0): " & Infos(1, 0)
      Stop
     End If
    Set rD = Nothing
    Call rD.Open("Select * from briefe where pat_id = " & Pid & " and name like 'Brief an %'", DBCn, adOpenStatic, adLockOptimistic)
    If Not rD.BOF Then
     FSO.DeleteFile (Fil.path)
    End If
'    Set rd = Nothing
'    Call rd.Open("Select * from briefe where pat_id = " & pid & " and name not like '%labor%' and name not like '%Tagebuch%' and name not like '%Nachricht an%' and name not like '% BZ%' and name not like '%DMP-%' and name not like '%EKG%' and name not like '%anmeldung%' and name not like '%datenliste%' and name not like '%Herzzentrum%' and name not like '%Termin%' and name not like '%Attest%' and name not like '%Schreiben%' and name not like '%-Bakt%' and name not like '%Bescheinigung%' and name not like '%Dokumentation%' and name not like '%Vorschläge%' and name not like '%Pumpeneinstellung%' and name not like '%anforderung%' and name not like '%Medikamentenplan%' and name not like '%befund%' and name not like '%blutdruck%' and name not like '%analyse%' and name not like '%sonogramm%' and name not like '%übersicht%' and name not like '%auswertung%' and name not like '%GPD%' and name not like '%untersuchung%' and name not like '%standar_tag%'", DBCn, adOpenStatic, adLockOptimistic)
'    If Not rd.BOF Then
'     Debug.Print rd!Name
'    End If
   End If
  Else
   MsgBox "Stop in falschebriefelöschen_Click:" & "pos <= 0"
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

Public Sub KassenkategorienBestimmen_Click()
 Dim namen$(20)
 Call ProgStart
 namen(0) = "aok": Call doKassKat("AOK", namen)
 namen(0) = "bkk": namen(1) = "betriebsk": namen(2) = "sbk": Call doKassKat("BKK", namen)
 namen(0) = "ikk": Call doKassKat("IKK", namen)
 namen(0) = "lkk": Call doKassKat("LKK", namen)
 namen(0) = "bkn": namen(1) = "knapp": namen(2) = "": Call doKassKat("BKN", namen)
 namen(0) = "BARMER"
 namen(1) = "DAK"
 namen(2) = "TKK"
 namen(3) = "KKH"
 namen(4) = "HEK"
 namen(5) = "HMK"
 namen(6) = "HKK"
 namen(7) = "GEK"
 namen(8) = "HZK"
 namen(9) = "KEH"
 namen(10) = "HAMBURG-MÜNCHENER"
 namen(11) = "Handelskrankenkasse"
 namen(12) = "Techniker"
 namen(13) = "HANSEATISCHE"
 namen(14) = "AUS "
 namen(15) = "EK"
 Call doKassKat("EK", namen, True)
 Call ProgEnde
End Sub
Sub doKassKat(Kateg$, namen$(), Optional isn%)
 Dim rs As ADODB.Recordset, rAF&, sql$, i%
 sql = "update kassenliste set Kateg = '" & Kateg & "' where ("
 i = 0
 Do
  If LenB(namen(i)) = 0 Then Exit Do
  sql = sql & "name like '%" & namen(i) & "%' or kurzname like '%" & namen(i) & "%'"
  namen(i) = ""
  i = i + 1
  If LenB(namen(i)) <> 0 Then
   sql = sql & "or "
  End If
 Loop
 sql = sql & ") and Kateg <> '" & Kateg & "'"
 If isn <> 0 Then
  sql = sql & " and isnull(Kateg)"
 End If
 Set rs = DBCn.Execute(sql, rAF)
 If rAF <> 0 Then
  Ausgeb "In die Kategorie '" & Kateg & "' wurden " & rAF & " Kassen eingeteilt.", True
 End If
End Sub

Private Sub LaborEintragen_Click()
 Call ProgStart
' Set ple3.hlese = Me
' Set ple2.hlese = Me
 Set ple.hlese = Me
 ple.Caption = "Laborwerte eintragen"
 ple.Show
 Call ProgEnde
End Sub
Private Sub Abrechnungsfehler_Click()
 Dim AbrF As New AbrechFehler
 AbrF.Show
 Set AbrF = Nothing
End Sub

Private Sub MachDB_Click()
 Dim MdB As New MachDatenbank
 MdB.Show
 Set MdB = Nothing
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
   Me.Visible = True
   Screen.MousePointer = 0
 End If
End Sub

Private Sub MyDB_Click()
 Call MyDB_Change
End Sub

Private Sub MyDB_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me, Me.MyDB.Name)
End Sub
Private Sub obMySQL_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me, Me.MyDB.Name)
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
  Call machODBCAcc
'  Call Me.dbv.RegSpeichern
  Call acon(quelleT)
'  Call dbv.cnVorb("", "anamnesebogen", "Patientendaten")
  Call PutReg(Me)
  If Me.Visible Then Screen.MousePointer = vbNormal
 End If
End Sub
Sub machODBCAcc()
  Me.dbv.ODBC = "Microsoft Access Driver (*.mdb)"
End Sub
Private Sub obMySQL_Click()
 If Not obStart Then
  If Me.Visible Then Screen.MousePointer = vbHourglass
  Call dbv.cnVorb("", "anamnesebogen", "Patientendaten")
  Call machODBCMy
  Call acon(quelleT)
  Call PutReg(Me)
  If Me.Visible Then Screen.MousePointer = 0
 End If
End Sub
Sub machODBCMy()
  Me.dbv.ODBC = "MySQL ODBC 3.51 Driver"
End Sub
Private Sub Patientenlaufzettel_Click()
 Call doPatientenlaufzettel
End Sub

Private Sub PLZeinzeln_Click()
 Call ProgStart
 Aktion = Patientenlaufzetteleinzeln
 Set pau.hlese = Me
 pau.Show
 Call ProgEnde
' Call doPLZeinzeln
End Sub ' PLZeinzeln_Click

Private Sub doPorto_Click()
 Call ProgStart
 Call dodoPorto
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
  Case Patientenlaufzetteleinzeln
'   Call doPLZeinzeln(Me.pau.PatID)
   Call dodoPLZ(Me.pau.PatID, Now, Now - Int(Now), True)
  Case DMPZettel
   Call einDMP(Me.pau.Pat_id)
 End Select
End Sub ' los
Private Sub BriefSchreiben_Click() ' Brief schreiben
 Call ProgStart
 Aktion = BriefSchreiben
 Set pau.hlese = Me
 pau.Show
' Aktion = nix
 Call ProgEnde
End Sub ' BriefSchreiben_Click

Private Sub CallDMPString_Click()
 Dim erg$
 Call ProgStart
 erg = InputBox("Bitte Patientenummer eingeben!")
 If IsNumeric(erg) Then
  Call doCallDMP(ByVal erg)
 End If
 Call ProgEnde
End Sub ' CallDMPString_Click()
Private Sub DMPForts_Click()
 Static Ausw As New ADODB.Recordset
 Call ProgStart
 If Ausw.State = 0 Then
  Ausw.Open "SELECT distinct pat_id, nachname, vorname, gebdat FROM dmpreihe where datum > " & DatForm(Now() - 365 * 0.5) & " order by nachname, vorname;", DBCn, adOpenDynamic, adLockReadOnly
 End If
 Ausw.Find "pat_id = " & lDMPPat_id, , adSearchBackward, 1
 If Not Ausw.EOF Then
  Ausw.Move 1
  Call doCallDMP(Ausw!Pat_id)
 End If
End Sub
Sub doCallDMP(ByVal Pid&)
 Dim dmpstD$, erg$, DT As dmptyp ' Dateiname
 Dim rAna As New ADODB.Recordset
 ReDim rNa(0)
 rNa(0).Pat_id = Pid
 rAna.Open "select * from namen where pat_id = " & Pid, DBCn, adOpenStatic, adLockReadOnly
 If Not rAna.BOF And Not IsNull(rAna!Nachname) And Not IsNull(rAna!Vorname) Then
  dmpstD = GesNam(rAna) & " (" & Pid & ") "
 Else
  dmpstD = pVerz & "DmpString "
 End If
 rAna.Close
 erg = DMPString$(rNa(0).Pat_id, DT)
 If lies.obMySQL Then
  dmpstD = dmpstD & Me.MyDB
 Else
  dmpstD = dmpstD & Replace$(Replace$(dlg.MdB, "\", "_"), ":", ".")
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
' call progstart
' Call doDMPListeAnzeigen(Me)
' Call ProgEnde
'End Sub

Private Sub FehlendeHausaerzte_Click()
'select distinct namen.pat_id, namen.nachname, namen.vorname from quelle.namen left join quelle.faelle on faelle.pat_id = namen.pat_id where (namen.kvnr = "" or isnull(namen.kvnr)) and faelle.schgr = 24 order by pat_id desc;
 Call ProgStart
 Call doLdFH(Me)
 Call ProgEnde
End Sub ' FehlendeHausaerzte_Click()
Private Sub HausärzteBKK_Click()
 Call ProgStart
 Call doHABKK(Me)
End Sub
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

Private Sub SortierungÄndern_Click()
 Call ProgStart
 Call doSortierungÄndern0
End Sub

Private Sub TabKop_Click() ' Tabelle kopieren
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
 Ausgeb "Fertig mit Festlegen der Therapiearten", True
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
' call progstart
' Call doEinlesen(True)
' Call ProgEnde
'End Sub ' Einlesen_Click(Index As Integer)

'Private Sub EinlesenAb_Click()
' Dim erg
' obMitAlterTab = True
' call progstart
' erg = InputBox("Ab welcher Patientennummer soll eingelesen werden?", "Einlesen ab Patientennummer")
' If IsNumeric(erg) Then
'  EinlAb = erg
'  Call doEinlesen(False)
'  Call ProgEnde
' End If
'End Sub

'Private Sub EinlesenEingelesene_Click()
' obMitAlterTab = True
' call progstart
' Call DBCn.Execute("update quelle.namen set aktzeit = 0")
' Call doEinlesen(False)
' Call ProgEnde
'End Sub ' EinlesenEingelesene_Click()

'Private Sub EinlesenohneLabor_Click()
' obMitAlterTab = True
' call progstart
' Call doEinlesen(False, True)
' Call ProgEnde
'End Sub ' EinlesenohneLabor_Click()

'Private Sub EinlesenSchnell_Click()
' obMitAlterTab = False
' call progstart
' Call doEinlesen(False)
' Call ProgEnde
'End Sub ' EinlesenSchnell_Click()
'Private Sub EinlesenEinzeln_Click()
' obMitAlterTab = True
' call progstart
' Call doEinlesen(False)
' Call ProgEnde
'End Sub ' EinlesenEinzeln_Click()
'Private Sub EinlesenVorb_Click()
' obMitAlterTab = True
' obVorb = True
' call progstart
' Call doEinlesen(True)
' Call ProgEnde
'End Sub ' EinlesenVorb_Click()
'Private Sub EinlesenEinzelnVorb_Click()
' obMitAlterTab = True
' obVorb = True
' call progstart
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
 Call ergänzeliste
' Call holAllg(Me, "hausaerzte", "ID", -1)
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
' call progstart
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
 Call ÜbertragFormulare(Me, "_Medarten", "Medarten", uVerz & "zugriff.mdb", "medarten")
 Exit Sub
 Call ÜbertragFormulare(Me, "Anamnesebogen", "AnBog", uVerz & "zugriff.mdb", "anamnesebogen")
 Call ÜbertragFormulare(Me, "Labordokumente eP", "Labordokumente", uVerz & "zugriff.mdb", "dokumente")
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
' call progstart
' Call DokPfadKorr(Me)
' Call ProgEnde
'End Sub

Private Sub AnamnesebogenPacken_Click()
 Call ProgStart
 Call AnPack
 Call ProgEnde
End Sub

'Private Sub LaborAlle_Click()
' call progstart
' Call LaborDirektImport(Me, -1)
' Call LaborErgPatId(Me)
' Call ProgEnde
'End Sub ' LaborAlle_Click()

'Private Sub LaborFehlende_Click()
' call progstart
' Call LaborDirektImport(Me, 0)
' Call LaborErgPatId(Me)
' Call ProgEnde
'End Sub ' LaborFehlende_Click()

'Private Sub Laborquerverbinden_Click()
' call progstart
' Call LaborErgPatId(Me)
' Call ProgEnde
'End Sub ' Laborquerverbinden_Click()

'Private Sub LaborQuerverbindenMitLöschen_Click()
' call progstart
' Call LaborErgPatId(Me, -1)
' Call ProgEnde
'End Sub ' LaborQuerverbindenMitLöschen_Click()

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

Private Sub VerdächtigeÜberweiser_Click()
 Call ProgStart
 Call doVerdächtigeÜberweiser
End Sub

Private Sub Vergleichen_Click() ' Vergleich der Datenbankstrukturen
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
' If ProgrammLauf(-1) Or DBCn.State = 0 Then ' Wenn Programm schon gestoppt war, dann abbrechen, sonst stoppen
'#End If
 If Not ProgLäuft Or DBCn.State = 0 Then
  Unload Me
  End
 End If
End Sub ' Beenden_Click(Index As Integer)


Private Sub Beginn_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Bytes_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Ende_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub GesBytes_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub GesDauer_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub


Private Sub nachCd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub obAcc_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub obMy_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Prozent_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub


Private Sub Sekunden_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub SyncTest_Click()
 Call formInhMach
End Sub ' SyncTest_Click

Private Sub ÜbertrageCd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub ViewsErstellen_Click()
 Call doViewsErstellen
End Sub

Private Sub VorhandeneBriefe_Click()
 Dim ergZ&
 Call ProgStart
 ergZ = doVorhandene
 Me.Ausgeb "Fertig mit Auswechseln der Vorhandenen: " & ergZ & IIf(BrichAb, ",5", "") & " Briefe bearbeitet. ", True
 Call ProgEnde
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VorhandeneBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' VorhandeneBriefe_Click
Private Sub WiedereinbestellungenDMP_Click()
 Dim rs As New ADODB.Recordset, sql$
 sql = "select * from (select f.quartal `Schein`, n.pat_id Pat_ID, concat(concat_ws(',*',concat_ws(',',n.nachname, n.vorname),date_format(n.gebdat,'%y')),', T: ',concat_ws(',',n.privattel,n.privattel_2,n.privatmobil)) Name, replace(replace(replace(notiz,char(13),' '),char(10),''),'DMP ','') `DMP-Notiz`, date_format(dokudatum,'%d.%m.%y') DMPDoku, if(dokudatum > subdate(concat(year(now()),'-',((month(now())-1) div 3)*3+1,'-1'),interval 3 month),'',if(dokudatum > subdate(concat(year(now()),'-',((month(now())-1) div 3)*3+1,'-1'),interval 6 month),'!','-')) fri, a1c.wert HbA1c, date_format(a1c.zeitpunkt,'%d.%m.%y') `HbA1c-Zpkt`, if(a1c.zeitpunkt > subdate(concat(year(now()),'-',((month(now())-1) div 3)*3+1,'-1'),interval 0 month),'',if(a1c.zeitpunkt > subdate(concat(year(now()),'-',((month(now())-1) div 3)*3+1,'-1'),interval 3 month),'!','!!')) alt from namen n left join aktfv af on n.pat_id = af.pat_id left join faelle f on n.pat_id = f.pat_id " & _
 "left join dmpreihe dr on dr.pat_id = n.pat_id left join (SELECT pat_id,zeitpunkt,wert FROM labor1 where abkü = 'hba1c' union select pat_id,zeitpunkt,wert from labor2 where abkü = 'hba1c') a1c on n.pat_id = a1c.pat_id left join anamnesebogen a on a.pat_id = n.pat_id where a.tkz = 0 and instr(notiz,'DMP hier')> 0 and isnull(af.vknr) and dokudatum > subdate(now(),interval 9 month) order by n.pat_id, mid(f.quartal,2) desc, f.quartal desc, dokudatum desc, a1c.zeitpunkt desc) i group by pat_id order by mid(`Schein`,2) desc, `Schein` desc, name;"
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 TabAusgeb rs, True, , , , Array(0, 0, 55, 25, 0, 3, 0, 0, 3), True, pVerz & "DMP-Wiedereinbestellungen"
End Sub
Private Sub WiedereinbestellungenDMP_Click_alt()
 Dim sql$, Zp$, obDruck%, AusgStr$, Grenze As Date
 Dim r1 As New ADODB.Recordset
 Dim r2 As New ADODB.Recordset
 Dim WDatei$
 On Error GoTo fehler
' Select Wert,Zeitpunkt from (select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"" as NB from (select n.Pat_ID AS Pat_ID,n.ZeitPunkt AS ZeitPunkt,n.FertigStGrad AS FertigStGrad,n.Abkü AS Abkü,l.Langtext AS Langtext,n.Wert AS Wert,n.Einheit AS Einheit,k.Kommentar AS Kommentar,n.AbsPos AS AbsPos,n.AktZeit AS AktZeit from (laborlangtext l inner join (laborkommentar k inner join laborneu n on ((k.KommentarVW = n.KommentarVW))) on ((l.LangtextVW = n.LangtextVW))) where pat_id = 105) as labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar, Normbereich as NB FROM laborxus LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr WHERE pat_id = 105 and not exists (select * from laborneu where pat_id = 105 and abkü = laborxwert.Abkü and wert = laborxwert.wert and zeitpunkt > laborxus.Eingang -3 and zeitpunkt < laborxus.Eingang+6)) as sql1 where abkü = "HBA1C" order by zeitpunkt desc
 Call ProgStart
 Grenze = QAnf(ZQuart(Now() - 90)) 'CDate("1.4.08")
 WDatei$ = pVerz & "Wiedereinbestellungen " & Format(Now, "d.m.yy hh.mm") & ".txt"
 Open WDatei For Output As #339
 Print #339, Right$(Space$(4) & "Pat_ID", 4) & " " & Left("Nachname" & Space$(15), 15) & " " & Left("Vorname" & Space$(9), 9) & " 'X'= DMP hier (Notiz-Eintrag / Versicherung)        letzter Fallbeginn  letztes HbA1c"
 Print #339, String$(110, "_")
 'r1.Open "select pat_id, notiz from namen na where notiz like '%hier%'", dbv.wCn, adOpenStatic, adLockReadOnly
 'select pat_id, notiz from namen na join (select pat_id from dmpreihe dr union select pat_id from namen na where notiz like '%hier%' order by pat_id) as innen using (pat_id)
 'r1.Open "select pat_id, na.nachname, na.vorname, lf.bhfb, notiz,rname from namen na join (select pat_id from dmpreihe dr " & _
         "where dokudatum > adddate(now(),-365) union select pat_id from _lfaelle f left join diagnosen d " & _
         "using (pat_id) where icd like 'E1%' and not icd like 'E16%' and diagsicherheit <> 'A'  and " & _
         "bhfb > adddate(now(),-180) order by pat_id) as innen using (pat_id) " & _
         "join lfaellev lf using (pat_id) " & _
         "join kassenliste k on k.ik = lf.ik " & _
         "join anamnesebogen an using (pat_id) where tkz=0 group by pat_id", dbv.wCn, adOpenStatic, adLockReadOnly '  ... union select pat_id from namen na where notiz like '%hier%'
 r1.Open "select pat_id, na.nachname, na.vorname, lf.bhfb, notiz,rname from namen na join (select pat_id from dmpreihe dr where dokudatum > adddate(now(),-365) union select pid pat_id from _lfaelle f left join diagnosen d on d.pat_id = f.pid where icd like 'E1%' and not icd like 'E16%' and diagsicherheit <> 'A'  and mbhfb > adddate(now(),-180) order by pat_id) as innen using (pat_id) join lfaellev lf using (pat_id) join kassenliste k on k.ik = lf.ik join anamnesebogen an using (pat_id) where tkz=0 group by pat_id", dbv.wCn, adOpenStatic, adLockReadOnly  '  ... union select pat_id from namen na where notiz like '%hier%'
 Do While Not r1.EOF
'  If obhierdmp(r1!notiz) Then
   Set r2 = Nothing
   r2.Open "select * from labor1 where pat_id = " & r1!Pat_id & " and abkü = 'hba1c' union select * from labor2 where pat_id = " & r1!Pat_id & " and abkü = 'hba1c' order by zeitpunkt desc limit 1", dbv.wCn, adOpenStatic, adLockReadOnly
   Zp = ""
   obDruck = True
   If Not r2.EOF Then
    Zp = r2!Zeitpunkt & " " & Right$(Space$(3) & r2!Wert, 3)
    If r2!Zeitpunkt >= Grenze Then obDruck = False
   End If
   If obDruck Then
    AusgStr = Right$(Space$(4) & r1!Pat_id, 4) & " " & Left(r1!Nachname & Space$(15), 15) & " " & Left(r1!Vorname & Space$(11), 11) & "   " & IIf(obhierdmp(r1!Notiz), "X", " ") & " (" & Left(IIf(IsNull(r1!Notiz) Or r1!Notiz = "", r1!rname, Replace$(Replace$(r1!Notiz, vbCr, ""), vbLf, "")) & ")" & Space$(42), 42) & " " & Left(r1!BhFB & Space$(10), 10) & " " & Zp
    Debug.Print AusgStr
    Me.Ausgeb AusgStr & vbCrLf & altAusgabe, True
    DoEvents
    Print #339, AusgStr
    Print #339, String$(110, "_")
   End If
'  End If
  r1.Move 1
 Loop
 Close #339
 Call Shell(Environ("systemroot") & "\system32\notepad.exe " & WDatei, vbNormalFocus)
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WiedereinbestellungenDMP_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' WiedereinbestellungenDMP_Click()

Private Sub WSt0Erg_Click()
 Call ProgStart
 Call doWSt0Erg
 Call ProgEnde
End Sub

Private Sub Zeilen_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
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

Private Sub Ziel_KeyDown(KeyCode As Integer, Shift As Integer)
 dlg.MdB = Me.Ziel
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 27 Then End
 Call Key(KeyCode, Shift, Me)
End Sub ' Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Ausgabe_KeyDown(KeyCode As Integer, Shift As Integer)
 Static obCtrl%, pa&, pe&, Zahl$
 Call Key(KeyCode, Shift, Me)
 If KeyCode = 17 Then
  obCtrl = True
 Else
  If obCtrl Then
   If KeyCode = 68 Then ' D
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
  If ctl.Name Like "*inlesen*" Then ctl.Enabled = True
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
   rs!TabellenEntleeren = dlg.TabellenEntleeren
   rs!ZurücksetzenLAktDat = dlg.ZurücksetzenLAktDat
   rs!obVglMitLetzterEinlesung = dlg.obVglMitLetzterEinlesung
   rs!Pat_IDBis = IIf(IsNumeric(dlg.Pat_IDBis), dlg.Pat_IDBis, 0)
   rs!Pat_IDBis = IIf(IsNumeric(dlg.Pat_IDBis), dlg.Pat_IDBis, 0)
   rs!AlterTab = dlg.AlterTab
   rs!VorladenFFI = dlg.VorladenFFI
   rs!ÜberTabelle = dlg.ÜberTabelle
   rs!SammelInsert = dlg.SammelInsert
   rs!bereinigeFormInhFeld = dlg.bereinigeFormInhFeld
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
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PutEinstAufDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
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
    MsgBox "Fehler bei Tabelle 'Eintragszahlen' in Datenbank '" & DefDB(DBCn) & "'" ' DBCn.Properties("Data Source Name")
    Exit Function
   End If
   On Error GoTo fehler
   dlg.TabellenEntleeren = Abs(HolEinstFeld(rs, "TabellenEntleeren", 0, "BIT"))
   dlg.TabellenEntleeren = 0
   dlg.ZurücksetzenLAktDat = Abs(HolEinstFeld(rs, "ZurücksetzenLAktDat", 0, "BIT"))
   dlg.obVglMitLetzterEinlesung = Abs(HolEinstFeld(rs, "obVglMitLetzterEinlesung", 0, "BIT"))
   dlg.ZurücksetzenLAktDat = 0
   dlg.Pat_IDVon = HolEinstFeld(rs, "Pat_IDvon", 0, "TEXT", 6)
   dlg.Pat_IDBis = HolEinstFeld(rs, "Pat_IDbis", 0, "TEXT", 6)
   dlg.AlterTab = Abs(HolEinstFeld(rs, "AlterTab", -1, "BIT"))
   dlg.VorladenFFI = Abs(HolEinstFeld(rs, "VorladenFFI", 0, "BIT"))
   dlg.ÜberTabelle = Abs(HolEinstFeld(rs, "ÜberTabelle", 0, "BIT"))
   dlg.SammelInsert = Abs(HolEinstFeld(rs, "SammelInsert", 0, "BIT"))
   dlg.bereinigeFormInhFeld = Abs(HolEinstFeld(rs, "bereinigeFormInhFeld", 0, "BIT"))
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HolEinstVonDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
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
 If InStrB(HolEinstFeld, ":\\") > 0 Or InStrB(HolEinstFeld, "\\\\") > 0 Then HolEinstFeld = Replace$(HolEinstFeld, "\\", "\")
 On Error GoTo fehler
 If IsNull(HolEinstFeld) Then HolEinstFeld = 0
 On Error GoTo fehler
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HolEinstFeld/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' HolEinstFeld(rs As ADODB.Recordset, FName$, Default, ByVal Typ$, Optional Lenge&)
Private Sub dbv_wCnAendern(CnStr$)
 On Error GoTo fehler
 ConStri = dbv.Constr
 obStart = True
 Me.obMySQL = InStrB(Me.dbv.wCn.ConnectionString, "MySQL") <> 0
 Me.obAcc = Not Me.obMySQL
 
 If Me.obMySQL <> 0 Then
  If InStrB(Me.dbv.Tabellen, "anamnesebogen") <> 0 Then
   If dbv.DaBa = "" And Not DBCn Is Nothing And DBCn.State <> 0 Then
    Me.MyDB = DBCn.Properties("Current Catalog")
   Else
    Me.MyDB = dbv.DaBa
   End If
  End If
 Else
  Me.Ziel = dbv.Datei
 End If
 
' 2.7.08: Kommentar der nä 3 Zeilen entfernt, da sonst bei Änderungen der Mysql-Tabellen weder in
' Lese noch in DBVerb DBCn korrigiert wird
 If Me.dbv.obQuelle Then ' Me.dbv.DaBa = InStrB(Me.dbv.Ü2, "Patienten") <> 0 Then  ' sonst bekommt DBCn Bedeutungen von Hausärzten usw
  DBCnS = Me.dbv.CnStr
  SetDBCn Me.dbv.wCn
 End If
 LVobMySQL = InStrB(Me.dbv.wCn.ConnectionString, "MySQL") <> 0 '(Not (cDtb = accDtb))
 Me.obMySQL = LVobMySQL
 Me.obAcc = Not Me.obMySQL
 obStart = False
 Call Zinit(LVobMySQL)
' Kommentar 15.5.09, da sonst nach Start lEinlesung automatisch geändert wird
' If DBCn.State <> 0 Then
'  Call Me.dlg.FrmLEinlesung
' End If
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in wcn_Aendern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' dbv_wCnAendern
Private Sub Konstanten()
 Dim Cpt$
 Cpt = LCase$(CptName)
 Select Case Cpt
  Case "pc08", "gerald03"
   uVerz = "c:\u\"
   pVerz = "c:\p\"
   vVerz = "c:\v\"
  Case Else
   uVerz = "u:\"
   pVerz = "p:\"
   vVerz = "v:\"
 End Select
 hVerz$ = uVerz & "TMImport\" ' HochladeVerzeichnis
 aVerz$ = uVerz & "Anamnese\"
 QmdB$ = aVerz & "Quelle.mdb" ' uverz & "Anamnese\Quelle.mdb
 eVerz$ = pVerz & "Patientenübergreifendes"
 üVerz$ = uVerz & "TMExport\"
 LabTransPfad$ = "\\anmeldr\BioWinBACKUP"
 StACCDB$ = AnamneseVerZeichnis1 & "quelle.mdb"
 StFxDB$ = uVerz & "FaxeinP.mdb"
 StFtDB$ = uVerz & "FotosinP.mdb"
 StOffDB$ = uVerz & "office.mdb"
 KVÄDatei1$ = AnamneseVerZeichnis1 + "KV-Ärzte neu.mdb"
 BriefZiel$ = pVerz
 AutoBriefZiel$ = BriefZiel & "unkorrigiert\"
 AutoBriefProtok$ = BriefZiel & "zufaxen.txt"
 plzVerz$ = pVerz & "plz\"
 FaxSendDatei$ = uVerz & "test1.txt"
 DMPVorlage$ = uVerz & "vorlagen\DMP-Vorlage gemein.dot"
End Sub ' Konstanten
Private Sub mdiForm_Load()
  Dim Diff#
  On Error GoTo fehler
  Call WD
  Call Konstanten
  
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
  .InitDir = üVerz ' uverz & "tmexport"
  .Filename = "*.BDT"
  .Orientation = cdlLandscape
 .Flags = 0
' .Flags = .Flags Or FileOpenConstants.cdlOFNExplorer
 .Flags = .Flags Or FileOpenConstants.cdlOFNHideReadOnly ' Schreibgeschützt-Checkbox entfernen
 .Flags = .Flags Or FileOpenConstants.cdlOFNLongNames ' '   Lange Dateinamen erlauben (nur sinnvoll bei Nicht-Win95-Design)
 .Flags = .Flags Or FileOpenConstants.cdlOFNFileMustExist
 .Flags = .Flags Or FileOpenConstants.cdlOFNPathMustExist
 End With
 Call mdiForm_Resize
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MDIForm_Load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
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
 If LenB(DBCn.ConnectionString) = 0 Then
  BrichAb = 0
  T1 = Now
  Me.Controls!Abbrechen.Enabled = True
' Me.Controls!Beenden.Enabled = False
  Call BeendenBlend(False)
' Call Me.ConstrFestleg(0, Me)
  Call acon(quelleT)
 ElseIf DBCn.State = 0 Then
  DBCnOpen  ' wenn o.g. Prozess abgebrochen wurde
 End If
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
  If Me.Controls(i).Name = "Beenden" Then
   Me.Controls(i).Enabled = obTrue
   Exit For
  End If
 Next i
 ProgLäuft = Not obTrue
End Sub ' BeendenBlend
Private Sub mdiForm_Resize()
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
 Me.EndZp.Top = Me.ZeilenBez.Top
 Me.SBez.Top = Me.EndZp.Top
 Me.Picture2.Top = Me.EndZp.Top
 Me.MyDB.Top = Me.ÜbertrageCd.Top
 Me.GesamtDauerBez.Top = Me.ZeilenBez.Top + 20
 Me.GesDauer.Top = Me.ZeilenBez.Top
 Me.Fuß.Top = Picture1.Height - Me.Fuß.Height - Me.ConStri.Height
 Me.ConStri.Top = Me.Fuß.Top + Me.Fuß.Height
End Sub ' MDIForm_Resize()

Private Sub mdiform_unload(Cancel As Integer) ' geht nur beim Anklicken des Kreuzes oben
 Call PutEinstAufDB
 Call PutReg(Me)
' Call Me.dbv.RegSpeichern ' Kommentar 25.9.09
 End
End Sub ' MDIForm_Unload(Cancel As Integer)
#If False Then
Public Function ConstrFestleg(ByVal Art As ConDtb, Optional hlese As Lese)   ' dlg ist für art= 0 und 1 nötig
 On Error GoTo fehler
'ConStr$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;uid=praxis;pwd=***REMOVED***;option=" & opti
 Select Case Art
  Case 0
   If hlese.obAcc Then
    Art = accDtb
   Else
    If hlese.MyDB = "quelle" Then
     Art = qDtb
    ElseIf hlese.MyDB = "quelle1" Then
     Art = q1Dtb
    ElseIf hlese.MyDB = "quelle2" Then
     Art = q2dtb
    End If
   End If
 End Select
 Dim MdB$
 If Art = 1 Then
 ' ConStr = CStrAcc + """" + dlg.MdB + """"
  If Not obStart Then
   If hlese.Ziel = "" Or Not FSO.FileExists(hlese.Ziel) Then Call hlese.dlg.MdBFestleg
  End If
'  MdB = dlg.MdB
' Else
'  MdB = ""
 End If
 Call doConstrFestleg(Art, obStart, hlese.Ziel, Me)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ConstrFestleg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ConstrFestleg
#End If

Public Function Ausgeb(Text$, obDauer%)
 Me.Ausgabe = Text & vbCrLf & altAusgabe
 If obDauer Then
  altAusgabe = Me.Ausgabe
 End If
 If InStrB(Text, "READ-COMMITTED") <> 0 Then
  MsgBox "Beinahe-Stop in Ausgeb:" & vbCrLf & "instrb(text, 'READ-COMMITTED') <> 0" & vbCrLf & "Text: " & Text
 End If
End Function
