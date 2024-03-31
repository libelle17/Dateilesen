VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm Lese 
   BackColor       =   &H8000000C&
   Caption         =   "Patientendaten Diabetespraxis Dachau"
   ClientHeight    =   6360
   ClientLeft      =   270
   ClientTop       =   1275
   ClientWidth     =   15240
   Icon            =   "Lese5.frx":0000
   LinkTopic       =   "Lese"
   Picture         =   "Lese5.frx":030A
   Begin VB.PictureBox Picture1 
      Align           =   1  'Oben ausrichten
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.TextBox DurchschnDauer 
         BackColor       =   &H80000018&
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   16920
         TabIndex        =   31
         Top             =   3960
         Width           =   1215
      End
      Begin VB.ComboBox MyDB 
         Height          =   315
         ItemData        =   "Lese5.frx":0614
         Left            =   9840
         List            =   "Lese5.frx":0616
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
         Left            =   16800
         TabIndex        =   27
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox QDatum 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   26
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox Ziel 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   12360
         TabIndex        =   23
         Top             =   3960
         Width           =   3135
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
         Left            =   11520
         TabIndex        =   22
         Top             =   3960
         Width           =   855
      End
      Begin VB.OptionButton obMySQL 
         Caption         =   "M&ySQL"
         Height          =   275
         Left            =   9000
         TabIndex        =   20
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox QDatei 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   3960
         Width           =   6255
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
         BackColor       =   &H00C0FFFF&
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
         Width           =   15495
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
         Caption         =   "Durschn.Dauer[s]:"
         Height          =   255
         Left            =   15600
         TabIndex        =   32
         Top             =   3960
         Width           =   1335
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
      Begin VB.Menu FalscheDiabetesdiagnosen 
         Caption         =   "&Falsche Diabetesdiagnosen"
      End
      Begin VB.Menu DMPHAKorr 
         Caption         =   "DMP-Teilnahme der Haus&ärzte korrigieren"
      End
      Begin VB.Menu Motivationsgesprächskandidaten 
         Caption         =   "&Motivationsgesprächskandidaten"
      End
      Begin VB.Menu Statistik_03355 
         Caption         =   "&Quartalsstatistik zu Leistung, z.B. 03355"
      End
      Begin VB.Menu Statistiken_zu_03230 
         Caption         =   "Statistiken zu &03230"
         Begin VB.Menu Statistik_zu_03230nachPatient 
            Caption         =   "03230-Zahl nach Patient ab 3"
         End
         Begin VB.Menu Statistik_zu_03230nachTag 
            Caption         =   "03230-Zahl nach Tag"
         End
         Begin VB.Menu Statistik_zu_03230nachTagundArzt 
            Caption         =   "03230-Zahl nach Tag und Arzt"
         End
         Begin VB.Menu Statistik_zu_03230nachArzt 
            Caption         =   "03230-Zahl nach Arzt"
         End
         Begin VB.Menu Statistik_zu_03230_einzeln 
            Caption         =   "03230-Zahl einzeln"
         End
      End
      Begin VB.Menu Abrechnungsfehler 
         Caption         =   "&Abrechnungsfehler"
      End
      Begin VB.Menu Ziffer30u31Ausschlüsse 
         Caption         =   "Ziffer &30 und 31-Ausschlüsse"
      End
      Begin VB.Menu Niereninsuffizienzpauschalendiabetiker 
         Caption         =   "&Niereninsuffizienzpauschalendiabetiker"
      End
      Begin VB.Menu Patientenlaufzettel_mit_Rueckmeldung 
         Caption         =   "&Patientenlaufzettel"
         Shortcut        =   ^P
      End
      Begin VB.Menu Patientenlaufzettel 
         Caption         =   "Patienten&laufzettel alt"
      End
      Begin VB.Menu Plzeinzeln_mr 
         Caption         =   "Pa&tientenlaufzettel einzeln"
         Shortcut        =   ^E
      End
      Begin VB.Menu PLZeinzeln 
         Caption         =   "Patientenla&ufzettel alt"
      End
      Begin VB.Menu Patientenlaufzettel_aus_zuplz 
         Caption         =   "Patientenlaufzettel aus &zuplz.txt"
      End
      Begin VB.Menu DMPBriefEinzeln 
         Caption         =   "&DMP-Brief einzeln"
      End
      Begin VB.Menu WiedereinbestellungenDMP 
         Caption         =   "&Wiedereinbestellungen DMP"
      End
      Begin VB.Menu Hausärzte_aus_Listenausgabe_Ueberweiser_einlesen 
         Caption         =   "&Hausärzte aus Listenausgabe_Ueberweiser einlesen"
      End
      Begin VB.Menu HausärztemitalterKVNrergänzen 
         Caption         =   "&Hausärzte mit alter KV-Nr ergänzen"
      End
      Begin VB.Menu PatientenlistefürHausarztmodell 
         Caption         =   "Patientenliste für Hausarztmodell"
      End
      Begin VB.Menu PatientenlistefürVollpauschale 
         Caption         =   "Patientenliste für &Vollpauschale"
      End
      Begin VB.Menu Dokumentnamenprüfung 
         Caption         =   "&Dokumentnamenprüfung"
      End
      Begin VB.Menu FehlendeHausaerzte 
         Caption         =   "Patienten mit uns als &Hausarzt (vorher fehlende Hausärzte ergänzen!)"
      End
      Begin VB.Menu FehlendeÜberweisungsscheine 
         Caption         =   "&Fehlende Überweisungsscheine"
      End
      Begin VB.Menu FalscheKarteikarteneinträge 
         Caption         =   "Falsche &Karteikarteneinträge"
      End
      Begin VB.Menu VerhunzteFotosEinfärben 
         Caption         =   "Verhunzte Fotos einf&ärben"
      End
      Begin VB.Menu Anwaltsunterlagen_für_Pat_zusammenstellen 
         Caption         =   "A&nwaltsunterlagen für Pat. zusammenstellen"
      End
      Begin VB.Menu SonderpatientenAnzeigen 
         Caption         =   "&Sonderpatienten anzeigen"
      End
      Begin VB.Menu Hausarzt_anzeigen 
         Caption         =   "Hausar&zt anzeigen"
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
      Begin VB.Menu DiabetesQuartalsdiagnosenInDauerdiagnosenUmwandeln 
         Caption         =   "&Diabetes-Quartalsdiagnosen in Dauerdiagnosen umwandeln (manuell)"
      End
      Begin VB.Menu DMPSend 
         Caption         =   "Alle &DMP-Dokumente an Hausärzte in p:\zufaxen erstellen"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu AlleDMPanHA 
         Caption         =   "Alle &DMP-Dokumente an Hausärzte faxen"
      End
      Begin VB.Menu DMP_Dokumente_an_HA_Nachweis 
         Caption         =   "DMP-Dokumente an HA-Nachweis"
      End
      Begin VB.Menu Kontrolllisten_für_DMP_HA 
         Caption         =   "K&ontrolllisten für DMP HA"
      End
      Begin VB.Menu UnverwertbareDMPEinträge 
         Caption         =   "Unverwertbare DMP-Einträge"
      End
      Begin VB.Menu DMPhierListe 
         Caption         =   "D&MP hier Liste"
      End
      Begin VB.Menu DMPhierListeoT 
         Caption         =   "D&MP hier Liste ohne Pat.mit Terminen"
      End
      Begin VB.Menu DMPKHKAsthma 
         Caption         =   "DM&P KHK Asthma"
      End
      Begin VB.Menu DMP_Übersicht 
         Caption         =   "DMP &Übersicht"
      End
      Begin VB.Menu DuplexKontrollieren 
         Caption         =   "&Duplex Kontrollieren"
      End
      Begin VB.Menu HausärztemitDMPPatienten 
         Caption         =   "Haus&ärzte mit DMP-Patienten"
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
         Shortcut        =   ^L
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
      Begin VB.Menu Briefschreiben 
         Caption         =   "&Brief schreiben"
         Shortcut        =   ^B
      End
      Begin VB.Menu Briefschreibenneu 
         Caption         =   "Br&ief schreiben neu"
      End
      Begin VB.Menu Briefnochmal 
         Caption         =   "Br&ief nochmal"
      End
      Begin VB.Menu BriefImport 
         Caption         =   "Brief zu letztem &Import schreiben"
         Shortcut        =   ^I
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
      Begin VB.Menu Banksachen 
         Caption         =   "Ba&nksachen"
         Begin VB.Menu LfdKosten 
            Caption         =   "&Lfd.Kosten"
         End
         Begin VB.Menu LfdKostenMitBetrag 
            Caption         =   "Lfd.Kosten mit &Betrag"
         End
         Begin VB.Menu KVÜberw 
            Caption         =   "&KV-Überw"
         End
         Begin VB.Menu LfdKostenEigenbetrieb 
            Caption         =   "Lfd.Kosten &Eigenbetrieb"
         End
         Begin VB.Menu LfdKostenEigenbetrmBetrag 
            Caption         =   "Lfd.Kosten &Eigenbetr.m.Betrag"
         End
         Begin VB.Menu LfdKostenPGiro 
            Caption         =   "Lfd.Kosten &PGiro"
         End
         Begin VB.Menu LfdKostenPGiromBetrag 
            Caption         =   "Lfd.Kosten &PGiro mit Betrag"
         End
      End
      Begin VB.Menu Faxnachweis 
         Caption         =   "Fa&xnachweis"
      End
      Begin VB.Menu Faxwarteschlange 
         Caption         =   "Fax&warteschlange"
      End
      Begin VB.Menu Faxe_gescheitert 
         Caption         =   "Fa&xe gescheitert"
      End
      Begin VB.Menu Pat_loeschen 
         Caption         =   "Pat. l&öschen"
      End
      Begin VB.Menu GefaxteBriefeAnzeigen 
         Caption         =   "Gefa&xte Briefe anzeigen"
      End
      Begin VB.Menu PLZfuerMedikament 
         Caption         =   "Patientenlaufzettel f&ür alle Pat.mit best. Medikament"
      End
      Begin VB.Menu VerdächtigeÜberweiser 
         Caption         =   "Verd&ächtige Überweiser"
      End
      Begin VB.Menu DoppelteDiagnosen 
         Caption         =   "&Doppelte Diagnosen ermitteln"
      End
      Begin VB.Menu KassenEditieren 
         Caption         =   "&KassenEditieren (Rabattverträge etc.)"
      End
      Begin VB.Menu MedartenEditieren 
         Caption         =   "&Medarten editieren"
      End
      Begin VB.Menu LaborparameterZusammenfassen 
         Caption         =   "Laborparameter z&usammenfassen"
      End
   End
   Begin VB.Menu Statistik 
      Caption         =   "&Statistik"
      Begin VB.Menu Covid_Impfliste 
         Caption         =   "&Covid-Impfliste"
      End
      Begin VB.Menu HausärzteBKK 
         Caption         =   "&Hausärzte von Pat. in der BKK"
      End
      Begin VB.Menu Überweiserstatistik 
         Caption         =   "&Überweiserstatistik"
      End
      Begin VB.Menu Überweiserstatistik2 
         Caption         =   "&Überweiserstatistik d.letzten 2a"
      End
      Begin VB.Menu Schulungsstatistik 
         Caption         =   "S&chulungsstatistik nach Schulungsart"
      End
      Begin VB.Menu GruppenschulungsstatisiknachZiffernzahlproQuartal 
         Caption         =   "&Gruppenschulungsstatisik nach Ziffernzahl pro Quartal"
      End
      Begin VB.Menu Schlungsziffer_analyse 
         Caption         =   "Sch&ulungszifferanalyse"
      End
      Begin VB.Menu Schulungsziffereinzelnachweis 
         Caption         =   "Sc&hulungsziffereinzelnachweis"
      End
      Begin VB.Menu Motivationsgesprächsstatistik 
         Caption         =   "&Motivationsgesprächsstatistik"
      End
      Begin VB.Menu PatientenMitAOKKriterien 
         Caption         =   "&Patienten mit AOK-Kriterien"
      End
      Begin VB.Menu GestationsdiabetikerinnenProQuartal 
         Caption         =   "&Gestationsdiabetikerinnen pro Quartal"
      End
      Begin VB.Menu Gestationsdiabetikerinnen 
         Caption         =   "&Gestationsdiabetikerinnen"
      End
      Begin VB.Menu PLZausListe 
         Caption         =   "&PLZ aus Liste (kein Insulin mehr)"
      End
      Begin VB.Menu PioglitazonRezepte 
         Caption         =   "&Pioglitazon-Rezepte ab 1.4.11"
      End
      Begin VB.Menu DiabetikerOhneSchulungLetztesJahr 
         Caption         =   "&Diabetiker ohne Schulung letztes Jahr"
      End
      Begin VB.Menu Therapieartenwechsel 
         Caption         =   "&Therapieartenwechsel"
      End
      Begin VB.Menu HbA1cStatistik 
         Caption         =   "HbA&1c-Statistik (dauert ...)"
      End
      Begin VB.Menu Hilfsmittelklassifikationen 
         Caption         =   "&Hilfsmittelklassifikationen"
         Shortcut        =   ^H
      End
      Begin VB.Menu Einlesungen 
         Caption         =   "&Einlesungen"
      End
      Begin VB.Menu EinlesungenAnzeigen 
         Caption         =   "&Einlesungen anzeigen"
      End
      Begin VB.Menu Fallzahlstand 
         Caption         =   "&Fallzahlstand"
         Shortcut        =   ^F
      End
      Begin VB.Menu Wohnortstatistik 
         Caption         =   "&Wohnortstatistik"
      End
      Begin VB.Menu GNR_Statistiken_einl 
         Caption         =   "&GNR-Statistiken einlesen"
      End
      Begin VB.Menu GNR_Statistik 
         Caption         =   "&GNR-Statistik"
      End
      Begin VB.Menu Quartalsvergleich 
         Caption         =   "&Quartalsvergleich"
      End
      Begin VB.Menu Pumpenträgerliste 
         Caption         =   "&Pumpenträgerliste"
      End
      Begin VB.Menu suchTel 
         Caption         =   "&suchTel"
      End
   End
   Begin VB.Menu EDV 
      Caption         =   "&EDV"
      Begin VB.Menu MachTypen 
         Caption         =   "&MachTypen (Datei Typen.bas erstellen)"
         Index           =   1
      End
      Begin VB.Menu MachDB 
         Caption         =   "&MachDB"
      End
      Begin VB.Menu AlleFallzahlstände 
         Caption         =   "Alle Fallzahlst&ände"
      End
      Begin VB.Menu Leistungen_zu_Patient_anzeigen 
         Caption         =   "&Leistungen zu Patient anzeigen"
      End
      Begin VB.Menu Therapiearten_anzeigen 
         Caption         =   "&Therapiearten anzeigen"
      End
      Begin VB.Menu DokumenteInDatenbank 
         Caption         =   "&Dokumente in Datenbank"
      End
      Begin VB.Menu AnamnesebogenPacken 
         Caption         =   "Anamnesebogen pa&cken (Stringfeldlängen optimieren)"
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
      Begin VB.Menu ZahlEintrag 
         Caption         =   "&Zahlen aus Einträgen ermitteln"
      End
      Begin VB.Menu WSt0Erg 
         Caption         =   "&Fußsyndrome Wagner-Stadium 0 ergänzen"
      End
      Begin VB.Menu TherapieartenFestlegen 
         Caption         =   "&Therapiearten festlegen"
      End
      Begin VB.Menu HAUebertrag 
         Caption         =   "Haus&ärzte übertragen"
         Enabled         =   0   'False
         Visible         =   0   'False
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
      Begin VB.Menu KassenLesen 
         Caption         =   "Kassen aus Turbomed-Ex&cel-Datei einlesen"
      End
      Begin VB.Menu Excelliste 
         Caption         =   "E&xcelliste"
      End
      Begin VB.Menu falschenLaboreintragZuPatlöschen 
         Caption         =   "&falschen Laboreintrag zu Pat. löschen"
      End
      Begin VB.Menu LabortestsZuordnen 
         Caption         =   "&Labortests zuordnen"
      End
      Begin VB.Menu Laborvergleich 
         Caption         =   "&Laborvergleich"
      End
      Begin VB.Menu LaborLöschenAb 
         Caption         =   "Labor (direkt -> ""X"") l&öschen ab"
      End
      Begin VB.Menu DMPListe 
         Caption         =   "DM&P-Liste erstellen"
         Visible         =   0   'False
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
      Begin VB.Menu Formulare_bereinigen 
         Caption         =   "&Formulare bereinigen"
      End
      Begin VB.Menu BooleanFelder 
         Caption         =   "&Boolean-Felder in MySQL-Datenbanken erstellen"
      End
      Begin VB.Menu ViewsErstellen 
         Caption         =   "&ViewsErstellen"
      End
      Begin VB.Menu FalscheDokumente 
         Caption         =   "&Falsche Dokumente"
      End
      Begin VB.Menu korrQD 
         Caption         =   "Quelldatum für alle Dokumente korrigieren"
      End
      Begin VB.Menu korrQB 
         Caption         =   "Quelldatum für alle Brief korrigieren"
      End
      Begin VB.Menu Apothekenrezepte 
         Caption         =   "Apothekenrezepte"
      End
      Begin VB.Menu DokumenteAbgehaktPrüfen 
         Caption         =   "Dokumente abgeha&kt prüfen"
      End
      Begin VB.Menu tabfuell 
         Caption         =   "&Tabellenfüllungen untersuchen (tabfuell)"
      End
      Begin VB.Menu DokumenteNeuAbhaken 
         Caption         =   "Dokumente neu a&bhaken"
      End
      Begin VB.Menu harealNeu 
         Caption         =   "&hareal neu aufbauen und namen.getha-Felder befüllen"
      End
      Begin VB.Menu HausärzteEinlesen 
         Caption         =   "Haus&ärzte aus kv-pdf-Datei einlesen"
      End
      Begin VB.Menu alleHausärzteEinlesen 
         Caption         =   "alle KV-PDF-Dateien der Haus&ärzte einlesen"
      End
      Begin VB.Menu Punktwerte 
         Caption         =   "&Punktwerte EBM2010"
      End
      Begin VB.Menu MedOffZpSetzen 
         Caption         =   "Med&OffZpSetzen"
      End
      Begin VB.Menu MedOffTabZahl 
         Caption         =   "MedOffTab&Zahl"
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
         Visible         =   0   'False
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
         Visible         =   0   'False
      End
      Begin VB.Menu MedKlassT 
         Caption         =   "&MedKlass"
      End
      Begin VB.Menu test 
         Caption         =   "&test"
      End
      Begin VB.Menu Gewichte 
         Caption         =   "&Gewichte"
      End
      Begin VB.Menu Gewichtsabnahmekandidaten 
         Caption         =   "&Gewichtsabnahmekandidaten"
      End
      Begin VB.Menu SortierungÄndern 
         Caption         =   "&Sortierung ändern"
      End
      Begin VB.Menu calldoGenMachDB 
         Caption         =   "&Generierte MachDB aufrufen"
      End
      Begin VB.Menu testlqanf 
         Caption         =   "&testlqanf"
      End
      Begin VB.Menu PatvonMO 
         Caption         =   "P&atvonMO"
      End
   End
   Begin VB.Menu Fenster 
      Caption         =   "Fe&nster"
      Begin VB.Menu Wechseln 
         Caption         =   "&Wechseln"
         Shortcut        =   ^W
      End
   End
End
Attribute VB_Name = "Lese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MoSer$ = "szn4"
Const MOCStr$ = "DRIVER={MySQL ODBC 8.0 Unicode Driver};server=" & MoSer & ";option=0;database=medoff;uid=medoff;pwd=medoff;port=2020;"
Public MOCon As New ADODB.Connection
Public rsco As New ADODB.Recordset
Public dlg As New Dialog
Public opt As New Optionen
Public snst As New Sonstige
'Public anb As New Anamnese
Public pataw As New PatAuswahl
Public ple As New LabEintr
Public pal As New PatListe
Public haanz As New Hausarztanzeigen
Public Tü As New TabÜbertr
'Public ple2 As New LabEin2
'Public ple3 As New LabEin3
Const HADBName$ = "haerzte"
#Const mitab = True ' auch in Formular
#If mitab Then
Public anBogÜ As New AnBog
#If zutesten Then
Public labtest As New LaborTest
#End If
#End If
'Public DMPlst As New DMPListen
Public obAusgedehnt% ' ob Accessdatenbank noch komprimiert werden muss
Public lzPID& ' letzte PATid des Importes
Enum AktionTyp
 nix
 Briefschreiben
 GefaxteAnzeigen
 RestlicheBriefe
 Patientenlaufzetteleinzeln
 DMPZettel
 Anwalt
End Enum
Public Aktion As AktionTyp
Public SpPat_id&, SpName$ ' Sonderpatient Pat_id
Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1
Public obRueck%
'Public WithEvents qdb AS QuelleDBC

' Datei -> Datenbankverbindung Patientendaten
Private Sub Datenbankverbindung_Click()
' Call dbv.rücksetzBedTbl
' Call dbv.setzBedTbl("anamnesebogen")
' dbv.Show 1
' SET cn = Nothing
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
End Sub ' Datenbankverbindung_Click

' Datei -> Beenden
Private Sub Beenden_Click(Index As Integer)
'#If False THEN
' IF ProgrammLauf(-1) OR DBCn.State = 0 THEN ' Wenn Programm schon gestoppt war, dann abbrechen, sonst stoppen
'#END IF
 If Not ProgLäuft Or DBCn.State = 0 Then
  Unload Me
  End
 End If
End Sub ' Beenden_Click

' Funktionen für Arzthelferin und Arzt -> Einlesen
Private Sub Datenbank_Click()
 Screen.MousePointer = vbHourglass
 dlg.BDTDatei = getLDatei(dlg.BDTDatei, "*.bdt")
 dlg.Show
 Screen.MousePointer = vbDefault
End Sub ' Datenbank_Click

' für Arzt -> DMP Übersicht
Private Sub DMP_Übersicht_Click()
 Dim rs As New ADODB.Recordset
 Dim quart$
 quart = InputBox("Quartal?", "Quartalseingabe", Left$(ZQuart(Now() - vgbVerspätung), 1) & Right$(ZQuart(Now() - vgbVerspätung), 2))
 sql = "SELECT NachName, VorName, GebDat, Pat_id, LanrID, Karteidatum, DATE(exportiert) EXP, DATE(dokudatum) Doku, Abk, Art " & vbCrLf & _
 "FROM dmpreihe e WHERE karteidatum BETWEEN " & Format(QAnf(quart), "YYYYmmdd") & " AND " & Format(QEnd(quart), "YYYYmmdd") & " AND exportiert<>18991230 ORDER BY lanrid, REPLACE(nachname,'€','C'), vorname, gebdat;"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , "DMP-Dok'en " & quart & ", nach LANRID, Nachname, Vorname, Geb'dat sortiert", , True, , , , , True
End Sub ' DMP_Übersicht_Click()

' EDV -> Formulare bereinigen
Private Sub Formulare_bereinigen_Click()
 Dim rAf&
 sql = "DELETE FROM forminhaltfeld WHERE NOT EXISTS (SELECT 1 FROM forminhfeld where feldvw = forminhaltfeld.feldvw LIMIT 1)"
 syscmd 4, sql
 myEFrag sql, rAf
 sql = "DELETE f FROM forminhfeld f LEFT JOIN forminhkopf k USING (foid) WHERE k.foid IS NULL;"
 syscmd 4, sql
 myEFrag sql, rAf
 sql = "DELETE fif FROM forminhaltfeldinh WHERE NOT EXISTS (SELECT 1 FROM forminhfeld WHERE feldinhvw = forminhaltfeldinh.feldinhvw LIMIT 1)"
 ' fif LEFT JOIN forminhfeld fi USING (feldinhvw) WHERE fi.feldinhvw IS NULL;"
 syscmd 4, sql
 myEFrag sql, rAf
 syscmd 4, "Fertig mit Bereinigen der Formulare"
End Sub ' Formulare_bereinigen_Click()

Private Sub MedOffZpSetzen_Click()
 Dim vorhin As Date, Tn$, tr&, jS$
 jS = DatFor_k(Now())
 On Error Resume Next
 MOCon.Open MOCStr
 On Error GoTo 0
' vorhin = DBCn.Execute("SELECT COALESCE((SELECT datum FROM moprot WHERE server='" & MoSer & "' LIMIT 1),0)").Fields(0)
 vorhin = myEFrag("SELECT COALESCE((SELECT datum FROM moprot WHERE server='" & MoSer & "' LIMIT 1),0)", , DBCn).Fields(0)
' DBCn.BeginTrans
' DBCn.Execute "TRUNCATE mozp;"
' DBCn.Execute "INSERT INTO mozp VALUES(now());"
' DBCn.CommitTrans
 myEFrag "TRUNCATE mozp;", , DBCn
 myEFrag "INSERT INTO mozp VALUES(now());", , DBCn
 Set rsco = Nothing
 rsco.Open "SELECT TABLE_NAME, TABLE_ROWS From information_schema.tables WHERE table_schema = 'medoff' AND table_name NOT LIKE '%_fsurogat_seq'", MOCon, adOpenStatic, adLockReadOnly
 Do While Not rsco.EOF
  Tn = rsco!table_name
'  tr = rsco!table_rows
  Dim obabfr%
  obabfr = 0
  If vorhin = 0 Then
   obabfr = -1
  Else
   tr = MOCon.Execute("SELECT COUNT(0) FROM `" & Tn & "`").Fields(0)
   If tr <> myEFrag("SELECT table_rows FROM moprot WHERE server='" & MoSer & "' AND table_name='" & Tn & "' AND datum=(SELECT MAX(datum) FROM moprot WHERE server='" & MoSer & "' AND table_name='" & Tn & "')", , DBCn).Fields(0) Then
    obabfr = -1
   End If
  End If
  If obabfr Then
   myEFrag "INSERT INTO moprot(server,datum,table_name,table_rows) VALUES('" & MoSer & "'," & jS & ",'" & Tn & "','" & tr & "')", , DBCn
  End If
  rsco.MoveNext
 Loop
 MsgBox "Datenbanken zum Zeitpunkt " & jS & " gemerkt."
 Set rsco = Nothing
End Sub ' MedOffZpSetzen_Click

' zeigt Tabellenänderung in medoff an, nachdem zum letzten Mal MedOffZpSetzen_Click aufgerufen wurde
Private Sub MedOffTabZahl_Click()
 Const datnam$ = "v:\moaend.txt"
 Dim rcol As New ADODB.Recordset, raen As New ADODB.Recordset
 Dim lzp As Date
 Dim Tn$, tr&, jS$, cols$, colN$, sql$, runde%, Prim$, colz& ', AnzS$
 Dim Wt() ' Werte
 On Error Resume Next
 MOCon.Open MOCStr
 On Error GoTo pfadfehler
 Open datnam For Output As #220
 On Error GoTo fehler
 jS = DatFor_k(Now())
 lzp = DBCn.Execute("SELECT COALESCE((SELECT letzt FROM mozp),0)").Fields(0)
 Debug.Print "letzter Zeitpunkt: ", lzp
 Set rsco = Nothing
 rsco.Open "SELECT TABLE_NAME, TABLE_ROWS From information_schema.tables WHERE table_schema = 'medoff' AND table_name NOT LIKE '%_fsurogat_seq'", MOCon, adOpenStatic, adLockReadOnly
 Do While Not rsco.EOF
  Tn = rsco!table_name
  syscmd 4, "Tabelle: " & Tn
'  tr = rsco!table_rows
  Debug.Print Tn
  Set rcol = Nothing
  cols = "": colN = "": colz = 0
  rcol.Open "SELECT column_name FROM information_schema.columns WHERE table_schema = 'medoff' and TABLE_NAME='" & Tn & "' ORDER BY ordinal_position", MOCon, adOpenStatic, adLockReadOnly
  Do While Not rcol.EOF
   colN = IIf(colN = "", rcol.Fields(0), colN & "," & rcol.Fields(0))
   colz = colz + 1
   If cols <> "" Then cols = cols & ",' ',"
   cols = cols & "CONCAT('" & rcol.Fields(0) & ":',CONVERT(COALESCE(LEFT(" & rcol.Fields(0) & ",20),'') USING 'utf8mb4'))"
   rcol.MoveNext
  Loop
  ReDim Wt(colz)
  Prim = MOCon.Execute("SELECT GROUP_CONCAT(DISTINCT COLUMN_NAME) sp FROM information_schema.key_column_usage i WHERE CONSTRAINT_NAME='PRIMARY' AND table_schema='medoff' AND table_name='" & Tn & "' AND column_name NOT IN ('row_start','row_end') GROUP BY table_catalog,table_schema,TABLE_NAME ORDER BY table_catalog,table_schema,table_name,ordinal_position").Fields(0)
'  sql = "SELECT CONCAT(" & cols & ",' ',row_start,' ',row_end) FROM `" & Tn & "` WHERE row_start>" & Format(lzp, "yyyymmddHHMMSS") ' FOR SYSTEM_TIME BETWEEN " & Format(lzp, "yyyymmddHHMMSS") & " AND NOW()"
  sql = "SELECT * FROM (SELECT CONCAT(" & cols & ",' ',row_start,' ',row_end) sp, row_start, LEAD(row_start,1) OVER (PARTITION BY " & Prim & " ORDER BY row_start) nrs, LAG(row_start,1) OVER (PARTITION BY " & Prim & " ORDER BY row_start) vrs, a.* FROM `" & Tn & "` FOR system_time ALL a) i WHERE row_start>" & Format(lzp, "yyyymmddHHMMSS") & " or nrs>" & Format(lzp, "yyyymmddHHMMSS") & " ORDER BY " & Prim & ",row_start;"
  Const Offs% = 4 ' Offset der ersten Spalte aus a.*
  runde = 0
  raen.Open sql, MOCon, adOpenStatic, adLockReadOnly
'  If Tn = "patfall" Then Stop
  Do While Not raen.EOF
   If runde = 0 Then
    tr = MOCon.Execute("SELECT COUNT(0) FROM `" & Tn & "`").Fields(0)
    Print #220, vbCrLf & Tn & " (" & DBCn.Execute("SELECT table_rows FROM moprot WHERE server='" & MoSer & "' AND table_name='" & Tn & "' AND datum=(SELECT MAX(datum) FROM moprot WHERE server='" & MoSer & "' AND table_name='" & Tn & "')").Fields(0) & " -> " & tr & "):"
    Print #220, colN & ":"
   End If
   Print #220, raen.Fields(0)
   rcol.MoveFirst
   colz = 0
   Do While Not rcol.EOF
    If runde = 0 Then
     If IsNull(raen.Fields(colz + Offs)) Then Wt(colz) = "NULL" Else Wt(colz) = raen.Fields(colz + Offs)
    ElseIf Not IsNull(raen!vrs) Then ' wenn Datensatz nicht neu eingefügt wurde
     If CStr(Wt(colz)) <> CStr(IIf(IsNull(raen.Fields(colz + Offs)), "NULL", raen.Fields(colz + Offs))) Then
      Print #220, rcol.Fields(0) & ": " & Wt(colz) & " -> " & raen.Fields(colz + Offs)
     End If
    End If
    colz = colz + 1
    rcol.MoveNext
   Loop
   runde = 1
   raen.MoveNext
  Loop
  Set raen = Nothing
' On Error Resume Next
' MOCon.Execute "ALTER TABLE `" & Tn & "` ADD SYSTEM VERSIONING"
' On Error GoTo fehler
'  If vorhin = 0 Or tr <> DBCn.Execute("SELECT table_rows FROM moprot WHERE server='" & moser & "' AND table_name='" & tn & "' AND datum=(SELECT MAX(datum) FROM moprot WHERE server='" & moser & "' AND table_name='" & tn & "')").Fields(0) Then
'   If vorhin <> 0 Then AnzS = IIf(AnzS = "", tn, AnzS & vbCrLf & tn)
'   DBCn.Execute "INSERT INTO moprot(server,datum,table_name,table_rows) VALUES('" & moser & "'," & jS & ",'" & tn & "','" & tr & "')"
'  End If
  rsco.MoveNext
 Loop
 Close #220
 Set rsco = Nothing
 zeigan datnam
' If AnzS <> "" Then
'  MsgBox "Folgende Tabellen befüllt: " & vbCrLf & AnzS
' Else
'  MsgBox "MedOffTabZahl_Click: keine Tabelle geändert"
' End If
 Exit Sub
pfadfehler:
 Open REPLACE$(datnam, "v:", "\\linux1\daten\down") For Output As #220
 Resume Next
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then
 DBCn.Close
 DBCn.Open
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MedOffTabZahl/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' MedOffTabZahl_Click


' Datei -> Optionen
Private Sub Optionen_Click()
 opt.Show
End Sub ' Optionen_Click()

Private Sub PatvonMO_Click()
 Call doPatvonMO(2112, 12112)
End Sub ' PatvonMO_Click

Private Sub Ziffer30u31Ausschlüsse_Click()
  Dim rs As New ADODB.Recordset, spmax%(3)
  spmax(0) = 6
  spmax(1) = 20
  spmax(2) = 10
  spmax(3) = 200
' sql$ = "SELECT n.Pat_id, gesname(n.Pat_id) `Name mit Telnr. " & tel & "`, PrivatTel, Privattel_2, Diensttel, PrivatMobil, PrivatFax FROM `namen` n WHERE privattel LIKE " & telm & " OR privattel_2 LIKE  " & telm & " OR privatfax LIKE  " & telm & " OR diensttel LIKE  " & telm & " OR privatmobil LIKE  " & telm
  sql = "SELECT d.pat_id, gesname(d.pat_id) PName,d.erstZP, d.titel Titel, CONCAT(PrivatMobil,' / ',Privattel,' / ',Privattel_2) Tel" & vbCrLf & _
 "FROM faelle f" & vbCrLf & _
 "LEFT JOIN namen n USING (pat_id)" & vbCrLf & _
 "LEFT JOIN desktop d USING (pat_id)" & vbCrLf & _
 "WHERE schgr=90 AND bhfe1> NOW() - INTERVAL 1 YEAR" & vbCrLf & _
 "AND d.titel RLIKE '[^1-9]15($|[^1-9])|[^1-9]30($|[^1-9])|[^1-9]31($|[^1-9])'" & vbCrLf & _
 " GROUP BY pat_id" & vbCrLf & _
 "ORDER BY pname" & vbCrLf
 myFrag rs, sql
 Call TabAusgeb(rs, Me, , , , , , True, "Ausschlüsse von Ziffern 15, 30 und 31")
End Sub ' Ziffer30u31Ausschlüsse_Click()

#If False Then
Private Sub Zurücksetzen_Click() ' nicht sichtbar: "Datei -> &Zurücksetzen des Programmlaufs"
 Call ProgrammLauf(-1) ' falls es fälschlich auf 0 steht: 0 = Programm läuft, -1 = nicht
 Dim ctl As Control
 For Each ctl In Me.Controls
  If ctl.name Like "*inlesen*" Then ctl.Enabled = True
 Next ctl
 Me.Zurücksetzen.Enabled = False
End Sub ' Zurücksetzen_Click
#End If

' Funktionen für Arzthelferin und Arzt -> Abbrechen (vor Speichern eines Patienten)
Private Sub Abbrechen_Click()
'#If False THEN
' MsgBox "Stopp Programm gleich"
' Call ProgrammLauf(-1) ' Programm stoppen
'#END IF
 BrichAb = True
End Sub ' Abbrechen_Click

' Funktionen für Arzthelferin und Arzt -> Labor eintragen
Private Sub LaborEintragen_Click()
 Call ProgStart
' SET ple3.hlese = Me
' SET ple2.hlese = Me
 Set ple.hlese = Me
 ple.Caption = "Laborwerte eintragen"
 ple.Show
 Call ProgEnde
End Sub ' LaborEintragen_Click

' Funktionen für Arzthelferin und Arzt -> Falsch abgehakte Dokumente ungültig stempeln
Public Sub FalschAbgehakteUngueltig_Click()
 Dim rAf&, zwg&, zug&, rs As New ADODB.Recordset, rl As New ADODB.Recordset
 Call ProgStart
 Me.Ausgeb "FalschAbgehakteUngültig ...", False
 myFrag rs, "SELECT --abgehakt ab, --ungueltig ug, pat_id, d.dokpfad, quelldatum qd FROM `dokumente abgehakt` da INNER JOIN `dokumente` d ON da.dokpfad = d.dokpfad"
 Do While Not rs.EOF
  Set rl = Nothing
  myFrag rl, "SELECT pat_id FROM `laborneu` WHERE pat_id = " & rs!Pat_id & " AND " & SelDatum("zeitpunkt", rs!qd)
  If rl.EOF And rs!ug <> 1 Then
   myEFrag "UPDATE `dokumente abgehakt` SET ungueltig = 1 WHERE dokpfad = '" & doUmwfSQL(rs!DokPfad, lies.obMySQL) & "'", rAf
   zug = zug + rAf
   If rAf = 0 Then
    MsgBox "Fehler beim Auffinden von " & doUmwfSQL(rs!DokPfad, lies.obMySQL) & " in `dokumente abgehakt` (Orginaldokpfad: " & rs!DokPfad & ")"
   End If
  ElseIf Not rl.EOF And rs!ug <> 0 Then
   myEFrag "UPDATE `dokumente abgehakt` SET ungueltig = 0 WHERE dokpfad = '" & doUmwfSQL(rs!DokPfad, lies.obMySQL) & "'", rAf
   zwg = zwg + rAf
   If rAf = 0 Then
    MsgBox "Fehler beim Auffinden von " & doUmwfSQL(rs!DokPfad, lies.obMySQL) & " in `dokumente abgehakt` (Orginaldokpfad: " & rs!DokPfad & ")"
   End If
  End If
  rs.Move 1
 Loop
 Ausgeb zwg & " Dokumente wieder gültig gestempelt", True
 Ausgeb zug & " Dokumente ungültig gestempelt", True
End Sub ' FalschAbgehakteUngueltig_Click

' Funktionen für Arzthelferin und Arzt -> Falsche Diabetesdiagnosen
Private Sub FalscheDiabetesdiagnosen_Click()
 Dim rs As New ADODB.Recordset
 ' diagsicherheit unf f6010 in aktfaelle.icd schon eingebaut
 myFrag rs, "SELECT f.pat_id, dmpklass, f.icd FROM `aktfaellev` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id WHERE NOT (icd REGEXP '^E1[01]\.|^R73|^O24.4')"
 TabAusgeb rs, Me, , , , , , , "Falsche Diabetesdiagnosen (E12, E13, E14; bitte nach Medikation, Anammnese, Antikörpern einordnen, ggf. z.B. 'Diabetes mellitus (sekundär) [E10.91]'"
End Sub ' FalscheDiabetesdiagnosen_Click

' Funktionen für Arzthelferin und Arzt -> DMP-Teilnahme der Hausärzte korrigieren
Private Sub DMPHAKorr_Click()
 Call ProgStart
 Set pal = Nothing
 pal.PLArt = artHA
 Set pal.hlese = Me
 Me.Hide
 pal.Show
 Call ProgEnde
End Sub ' DMPHAKorr_Click

' Funktionen für Arzthelferin und Arzt -> Motivationsgesprächskandidaten
Private Sub Motivationsgesprächskandidaten_Click()
' Dim rv As New ADODB.Recordset, rs As New ADODB.Recordset, i&, ausg$, TA1$, SpMax%(5), fristS$, sql$
 Dim rs As New ADODB.Recordset, spmax%(5), sql$

 Call ProgStart
' myFrag rv, "SHOW CREATE VIEW `aktfv`"
' fristS = rv.Fields(1)
' SET rv = Nothing
' fristS = Mid$(fristS, InStr(fristS, "INTERVAL ") + 9)
' fristS = Left$(fristS, InStr(fristS, " ") - 1)
' IF Not IsNumeric(fristS) THEN
'  MsgBox "Ungeeignete Abfrage `aktfv`, evtl. erst Views erstellen"
'  Exit Sub
' END IF

 spmax(1) = 32
 spmax(5) = 300
' Const Schulungsleistungen$ = "(leistung LIKE '972%' OR leistung LIKE '922%' OR Leistung='92282' OR Leistung='92278' OR Leistung='92281' OR Leistung='92277') AND NOT leistung IN ('97272','97276','97277')"
' Print #325, "Lfdnr. Pat_id Name                          ICD"
' myFrag rs, "SELECT f.pat_id,gesname, icd FROM `aktfv` f LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND icd REGEXP '^E1[0-4]' AND diagsicherheit <> 'A' LEFT JOIN leistungen l ON f.pat_id = l.pat_id AND (leistung LIKE '972%' OR leistung LIKE '922%' OR leistung = '92282' OR leistung = '92278') AND NOT leistung IN ('97272','97276','97277') AND YEAR(zeitpunkt) = YEAR(SUBDATE(NOW(),INTERVAL 20 DAY)) AND adddate(zeitpunkt,INTERVAL 365 DAY) > now() WHERE NOT ISNULL(id1) AND ISNULL(leistung) GROUP BY f.pat_id"
' sql = "SELECT f.pat_id AS Pat_ID,gesname AS Name, ICD, DATE(e.zeitpunkt) AS Zeitpunkt, e.art AS Art, e.inhalt AS Inhalt FROM `aktfv` f LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND icd REGEXP '^E1[0-4]|^O24' AND diagsicherheit <> 'A' LEFT JOIN leistungen l ON f.pat_id = l.pat_id AND (leistung LIKE '972%' OR leistung LIKE '922%' OR leistung = '92282' OR leistung = '92278') AND NOT leistung IN ('97272','97276','97277') AND YEAR(zeitpunkt) = YEAR(SUBDATE(NOW(),INTERVAL " & FristS & " DAY)) " & _
         "and adddate(zeitpunkt,INTERVAL 365 DAY) > now() LEFT JOIN `eintraege` e ON f.pat_id = e.pat_id AND e.art IN ('andm','wr','tk','gs') AND e.zeitpunkt BETWEEN STR_TO_DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & FristS & " DAY)),'/',((month(SUBDATE(NOW(),INTERVAL " & FristS & " DAY))-1) div 3)*3+1,'/1'),'%Y/%m/%d') AND SUBDATE(adddate(STR_TO_DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & FristS & " DAY)),'/',((month(SUBDATE(NOW(),INTERVAL 20 DAY))-1) div 3)*3+1,'/1'),'%Y/%m/%d'),INTERVAL 3 MONTH),INTERVAL 1 DAY) WHERE NOT ISNULL(id1) AND ISNULL(l.leistung) GROUP BY f.pat_id, art, zeitpunkt"
' ' folgendes nach "and " & Schulungsleistungen & " auskommentiert 10.4.12: " AND YEAR(zeitpunkt) = YEAR(SUBDATE(NOW(),INTERVAL " & FristS & " DAY)) " & \_
 sql = motsql()
 myFrag rs, sql
Call TabAusgeb(rs, Me, , , , , spmax, True, "Kandidaten f.neue Motivationsziffern (92278(T2DM), 92282(T1DM))")
' Open DatNam For Output AS #325
' Print #325, TA1
' i = 0
' Do While Not rs.EOF
'  i = i + 1
'  ausg = Right$("    " & i, 4) & " " & Right$("     " & rs!Pat_id, 5) & " " & LEFT(rs!GesName & Space$(30), 30) & " " & rs!ICD
'  Print #325, ausg
'  rs.Move 1
' Loop
' Close #325
' MsgBox "Fertig!"
' Call ProgEnde
' zeigan DatNam
End Sub ' Sub Motivationsgesprächskandidaten_Click

' Funktionen für Arzthelferin und Arzt -> Quartalsstatistik zu Leistung, z.B. 03355
Private Sub Statistik_03355_Click()
 Dim sql$, rs As New ADODB.Recordset
 Dim Leistung$
 Leistung = InputBox("Leistung: ", "Rückfrage", "03355")
 sql = "SELECT COUNT(0) `Zahl 03355`, Quartal FROM leistungen LEFT JOIN faelle f USING (fid) WHERE leistung='" & Leistung & "' GROUP BY f.quartal ORDER BY MID(quartal,2) DESC,quartal DESC"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , , , , , "Quartalsstatistik zu Leistung " & Leistung
End Sub ' Statistik_03355_Click

' Funktionen für Arzthelferin und Arzt -> Statistiken zu 03230 -> 03230-Zahl nach Patient ab 3
Private Sub Statistik_zu_03230nachPatient_Click()
 Const Zahl% = 3
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT f.pat_id, gesname(f.pat_id) PName, DATE(zeitpunkt) Tag, CASE WHEN l.lanrid=1 THEN 'gs' WHEN l.lanrid=2 THEN 'tk' END Arzt, SUM(lzahl) Zahl " & vbCrLf & _
         "FROM aktfv f LEFT JOIN leistungen l ON l.pat_id=f.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
         "WHERE leistung='03230' GROUP BY f.pat_id HAVING SUM(lzahl)>= " & Zahl & ";"
 TabAusgeb rs, Me, , , , , , , "03230 nach Patient ab " & Zahl & " Leistungen", 1
End Sub ' Statistik_zu_03230nachPatient_Click

' Funktionen für Arzthelferin und Arzt -> Statistiken zu 03230 -> 03230-Zahl nach Tag
Private Sub Statistik_zu_03230nachTag_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT f.pat_id, gesname(f.pat_id) PName, DATE(zeitpunkt) Tag, sum(lzahl) Zahl, ROUND(sum(lzahl)/6,1) Stunden" & vbCrLf & _
         "FROM aktfv f LEFT JOIN leistungen l ON f.pat_id=l.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
         "WHERE leistung='03230' GROUP BY DATE(zeitpunkt) ORDER BY DATE(zeitpunkt);"
 TabAusgeb rs, Me, , , , , , , "03230 nach Tag ", 1
End Sub ' Statistik_zu_03230nachTag_Click

' Funktionen für Arzthelferin und Arzt -> Statistiken zu 03230 -> 03230-Zahl nach Tag und Arzt
Private Sub Statistik_zu_03230nachTagundArzt_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT f.pat_id, gesname(f.pat_id) PName, DATE(zeitpunkt) Tag, CASE WHEN l.lanrid=1 THEN 'gs' WHEN l.lanrid=2 THEN 'tk' END Arzt, sum(lzahl) Zahl, ROUND(sum(lzahl)/6,1) Stunden" & vbCrLf & _
         "FROM aktfv f LEFT JOIN leistungen l ON f.pat_id=l.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
         "WHERE leistung='03230' GROUP BY DATE(zeitpunkt),l.lanrid ORDER BY DATE(zeitpunkt);"
 TabAusgeb rs, Me, , , , , , , "03230 nach Tag und Arzt", 1
End Sub ' Statistik_zu_03230nachTagundArzt_Click

' Funktionen für Arzthelferin und Arzt -> Statistiken zu 03230 -> 03230-Zahl nach Arzt
Private Sub Statistik_zu_03230nachArzt_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT CASE WHEN l.lanrid=1 THEN 'gs' WHEN l.lanrid=2 THEN 'tk' END Arzt, sum(lzahl) Zahl, ROUND(sum(lzahl)/6,1) Stunden " & vbCrLf & _
         "FROM aktfv f LEFT JOIN leistungen l ON f.pat_id=l.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
         "WHERE leistung='03230' GROUP BY l.lanrid WITH ROLLUP;"
 TabAusgeb rs, Me, , , , , , , "03230 nach Arzt", 1
End Sub ' Statistik_zu_03230nachArzt_Click

' Statistik_zu_03230_einzeln
Private Sub Statistik_zu_03230_einzeln_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT f.pat_id, gesname(f.pat_id) PName, zeitpunkt zp, CASE WHEN l.lanrid=1 THEN 'gs' WHEN l.lanrid=2 THEN 'tk' END Arzt, lzahl " & vbCrLf & _
         "FROM aktfv f LEFT JOIN leistungen l ON f.pat_id=l.pat_id AND l.zeitpunkt BETWEEN qanf() AND qend()" & vbCrLf & _
         "WHERE leistung='03230' ORDER BY f.pat_id, l.zeitpunkt;"
 TabAusgeb rs, Me, , , , , , , "03230 nach Arzt", 1
End Sub

' Funktionen für Arzthelferin und Arzt -> Abrechnungsfehler
Private Sub Abrechnungsfehler_Click()
 Dim AbrF As New AbrechFehler
 AbrF.Show
 Set AbrF = Nothing
End Sub ' Abrechnungsfehler_Click

' 24.5.14 hier zu arbeiten
' Funktionen für Arzthelferin und Arzt -> Niereninsuffizienzpauschalendiabetiker
Private Sub Niereninsuffizienzpauschalendiabetiker_Click()
Dim rs As New ADODB.Recordset, rsa As New ADODB.Recordset, sql$, sqla$, gesZ%, pz%, nz%, obNP%, maxAlb#, aktAlb#
myEFrag ("DROP TABLE IF EXISTS `ni_abr`")
myEFrag ("CREATE TABLE `quelle`.`ni_abr`(`id` int(11) NOT NULL AUTO_INCREMENT key,`pat_id` int(10),DmICD varchar(8),maxHbA1c FLOAT, maxGluc FLOAT, eGFR FLOAT, npICD varchar(8), niICD varchar(8), pZ int(3), nZ int(3), gesZ int(3), minDat date, maxAlb DECIMAL(8,2), kasse varchar(20))")
sql = "SELECT f.pat_id, d.icd DmICD, IF(xh.max1>xh.max2,xh.max1, xh.max2) maxHbA1c, IF(xg.max1>xg.max2,xg.max1, xg.max2) maxGluc, _lGFR(f.pat_id) eGFR, dn.icd npICD, di.icd niICD, LEFT(k.name,20) Kasse " & _
        "FROM aktfvs v " & _
        "LEFT JOIN kassenliste k ON v.vknr=k.vknr AND v.ik=k.ik" & _
        "LEFT JOIN faelle f ON v.fid = f.fid " & _
        "LEFT JOIN diagview d ON f.pat_id = d.pat_id AND d.gICD RLIKE '^E1[0-4]\.' " & _
        "LEFT JOIN diagview dn ON f.pat_id = dn.pat_id AND dn.gICD LIKE 'N08.3%' " & _
        "LEFT JOIN diagview di ON f.pat_id = di.pat_id AND di.gICD LIKE 'N18%' " & _
        "LEFT JOIN _maxHbA1c xh ON f.pat_id = xh.pat_id " & _
        "LEFT JOIN _maxGluc xg ON f.pat_id = xg.pat_id " & _
        "WHERE k.kateg IN ('AOK','EK') AND (NOT ISNULL(d.icd) OR xh.max1>=6.5 OR xh.max2>= 6.5 OR xg.max1>=200 OR xg.max2>=200) AND _lGFR(f.pat_id)>59 " & _
        "AND (ISNULL(dn.icd) OR ISNULL(di.icd) OR (_lGFR(f.pat_id)<90 AND di.icd<>'N18.2') OR (_lGFR(f.pat_id)>=90 AND di.icd<>'N18.1')) " & _
        "GROUP BY f.pat_id"
myFrag rs, sql
If Not rs.BOF Then
 Do While Not rs.EOF
    gesZ = 0: pz = 0: nz = 0: obNP = 0: maxAlb = 0
'  sqla = LabEPatS(AlbCre, rs!Pat_id)
    Set rsa = Nothing
'   myFrag rsa, sqla
    Dim lwZahl&, aktlwx&
    Dim lab() As labtyp
    Set rsa = hollabor(rs!Pat_id, "", 0, 0, 0, lwZahl)
    If Not rsa.BOF And lwZahl Then
     ReDim lab(lwZahl)
     aktlwx = 0
     Do While Not rsa.EOF
      lab(aktlwx).Abkü = rsa!Abkü
      lab(aktlwx).WertSg = rsa!Wert
      lab(aktlwx).Einheit = rsa!Einheit
      lab(aktlwx).Zp = rsa!Zeitpunkt
      aktlwx = aktlwx + 1
      rsa.MoveNext
     Loop
     If lwZahl Then
      Dim Zp As Date
      For aktlwx = 0 To lwZahl
       If obLabI(LA_AlbCre, lab(aktlwx)) Then
   
'   IF Not rsa.BOF THEN
'    Do While Not rsa.EOF
'     Debug.Print rs!Pat_id, rsa.Fields(0), rsa.Fields(1)
        gesZ = gesZ + 1
        Zp = lab(aktlwx).Zp
        aktAlb = MachNumerisch(lab(aktlwx).WertSg) 'rsa!wert)
        If aktAlb > maxAlb Then maxAlb = aktAlb
        If aktAlb >= 20 Then pz = pz + 1 Else nz = nz + 1
        If pz > 1 And Not (nz >= pz + pz) Then obNP = 1: Exit For ' Do
'     rsa.MoveNext
'    Loop
'   END IF
       End If ' IF obLabI(LA_AlbCre, lab(aktlwx)) THEN
      Next aktlwx
'  sqla = "SELECT DATE(zeitpunkt) zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) Wert FROM `laborneu` ln LEFT JOIN laborkommentar lk ON ln.kommentarvw = lk.kommentarvw WHERE ((abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g %') OR (abkü IN ('ALBU','ALBUMU') AND (einheit = 'mg/l' OR einheit = ''))) AND pat_id = 262 UNION SELECT DATE(u.eingang) zp, IF(ISNULL(w.wert),IF(ISNULL(w.kommentar),'',w.kommentar),w.wert) Wert FROM `laborxus` u LEFT JOIN laborxwert w ON u.refnr = w.refnr WHERE ((abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g %') OR (abkü IN ('ALBU','ALBUMU') AND (einheit = 'mg/l' OR einheit = ''))) AND pat_id = 262 GROUP BY zp ORDER BY zp DESC;"
      If obNP Then
       myEFrag ("INSERT INTO ni_abr(pat_id,DmICD,maxHbA1c,maxGluc,eGFR,npICD,niICD,pZ,nZ,gesZ,minDat,maxAlb,kasse) VALUES(" & rs!Pat_id & ",'" & rs!DmICD & "','" & REPLACE$(IIf(IsNull(rs!maxHbA1c), 0, rs!maxHbA1c), ",", ".") & "','" & REPLACE(IIf(IsNull(rs!maxGluc), "0", rs!maxGluc), ",", ".") & "','" & REPLACE(rs!eGFR, ",", ".") & "','" & IIf(IsNull(rs!npICD), "", rs!npICD) & "','" & IIf(IsNull(rs!niICD), "", rs!niICD) & "'," & pz & "," & nz & "," & gesZ & "," & Format(Zp, "YYYYMMDD") & "," & REPLACE$(maxAlb, ",", ".") & ",'" & rs!Kasse & "')")
      End If ' obNP
     End If ' lwzahl
    End If ' not rsa.BOF
    rs.MoveNext
 Loop ' Do While Not rs.EOF
End If ' IF Not rs.BOF THEN
Set rs = Nothing
sql = "SELECT pat_id, gesnameg(pat_id) Name, DmICD, eGFR, IF(dmicd NOT IN ('E10.20','E10.21','E11.20','E11.21','E13.20','E13.21','E10.72','E10.73','E11.72','E11.73','E13.72','E13.73'),CONCAT('E1',MID(dmicd,3,1),'.',IF(MID(dmicd,5,1) IN ('9','2'),'2','7'),IF(MID(dmicd,6,1) IN ('1','3','5'),IF(MID(dmicd,5,1) IN ('9','2'),'1','3'),IF(MID(dmicd,5,1) IN ('9','2'),'0','2')),IF(MID(dmicd,5,1)='7','  dazu',CONCAT('  statt ',dmicd))),'') fICD, IF(npicd='','N08.3 dazu','') fNP " & _
      ",IF(eGFR<90 AND niICD<>'N18.2',CONCAT('N18.2 ',IF(niICD='','dazu',CONCAT('statt ',niicd))),CONCAT('N18.1 ',IF(niICD='','dazu',CONCAT('statt ',niicd)))) fNI, kasse " & _
      "FROM ni_abr"
myFrag rs, sql
TabAusgeb rs, Me, , , , , , , "Korrekturbedarf für AOK-Vergütung Nephropathiekodierung", 1
End Sub ' Niereninsuffizienzpauschalendiabetiker_Click

' Funktionen für Arzthelferin und Arzt -> Patientenlaufzettel
Private Sub Patientenlaufzettel_mit_Rueckmeldung_Click()
 Call doPatientenlaufzettel(, obphp:=True)
End Sub ' Patientenlaufzettel_mit_Rueckmeldung_Click

' Funktionen für Arzthelferin und Arzt -> Patientenlaufzettel einzeln
Private Sub Plzeinzeln_mr_Click()
 Call ProgStart
 Aktion = Patientenlaufzetteleinzeln
 Set pataw.hlese = Me
 pataw.obRueck = True ' für PLZ-Knopf
 obRueck = True ' für OK-Knopf
 pataw.ZeilenzahlL.Visible = True
 pataw.Zeilenzahl.Visible = True
 pataw.Show
 Call ProgEnde
' Call doPLZeinzeln
End Sub ' Plzeinzeln_mr_Click

' Funktionen für Arzthelferin und Arzt -> Patientenlaufzettel alt (1. Vorkommen)
Private Sub Patientenlaufzettel_Click()
 Call doPatientenlaufzettel
End Sub ' Patientenlaufzettel_Click

' Funktionen für Arzthelferin und Arzt -> Patientenlaufzettel alt (2. Vorkommen)
Private Sub PLZeinzeln_Click()
 Call ProgStart
 Aktion = Patientenlaufzetteleinzeln
 Set pataw.hlese = Me
 pataw.obRueck = False
 obRueck = False
 pataw.ZeilenzahlL.Visible = True
 pataw.Zeilenzahl.Visible = True
 pataw.Show
 Call ProgEnde
' Call doPLZeinzeln
End Sub ' PLZeinzeln_Click

' Funktionen für Arzthelferin und Arzt -> Patientenlaufzettel aus zuplz.txt
Private Sub Patientenlaufzettel_aus_zuplz_Click()
Const dname$ = "Abrechnung\zuplz.txt"
Dim sLine$, pid$
Open pVerz & dname For Input As #1
While Not EOF(1)
  Line Input #1, sLine
  pid = LTrim$(sLine)
  If (InStr(pid, " ")) Then pid = Left$(pid, InStr(pid, " ") - 1)
  If (InStr(pid, ",")) Then pid = Left$(pid, InStr(pid, ",") - 1)
'  Debug.Print pid
  If IsNumeric(pid) Then
    dodoplz pid, plzVz, Now, Now - Int(Now), True, ""
  End If ' IsNumeric(pid) Then
Wend ' Not EOF(1)
Close #1
End Sub ' Patientenlaufzettel_aus_zuplz_Click

' Funktionen für Arzthelferin und Arzt -> DMP-Brief einzeln
Private Sub DMPBriefEinzeln_Click()
 Call ProgStart
 Aktion = DMPZettel
 Set pataw.hlese = Me
 pataw.ZeilenzahlL.Visible = True
 pataw.Zeilenzahl.Visible = True
 pataw.Show
 Call ProgEnde
' Call doPLZeinzeln
End Sub ' DMPBriefEinzeln_Click

' Funktionen für Arzthelferin und Arzt -> Wiedereinbestellungen DMP
Private Sub WiedereinbestellungenDMP_Click()
' wegen falsch eingetragener Fremdlabore gestrichen: AND einheit = '%'
 Dim rs As New ADODB.Recordset, sql$
 'sql = "SELECT * FROM (SELECT f.quartal `Schein`, n.pat_id Pat_ID, CONCAT(CONCAT_WS(',*',CONCAT_WS(',',n.nachname, n.vorname),DATE_FORMAT(n.gebdat,'%y')),', T: ',CONCAT_WS(',',n.privattel,n.privattel_2,n.privatmobil)) Name, REPLACE(REPLACE(REPLACE(notiz,char(13),' '),char(10),''),'DMP ','') `DMP-Notiz`, DATE_FORMAT(dokudatum,'%d.%m.%y') DMPDoku, IF(dokudatum > SUBDATE(CONCAT(YEAR(NOW()),'-',((month(NOW())-1) div 3)*3+1,'-1'),INTERVAL 3 MONTH),'',IF(dokudatum > SUBDATE(CONCAT(YEAR(NOW()),'-',((month(NOW())-1) div 3)*3+1,'-1'),INTERVAL 6 MONTH),'!','-')) fri, a1c.wert HbA1c, DATE_FORMAT(a1c.zp,'%d.%m.%y') `HbA1c-Zpkt`, IF(a1c.zp > SUBDATE(CONCAT(YEAR(NOW()),'-',((month(NOW())-1) div 3)*3+1,'-1'),INTERVAL 0 MONTH),'',IF(a1c.zp > SUBDATE(CONCAT(YEAR(NOW()),'-',((month(NOW())-1) div 3)*3+1,'-1'),INTERVAL 3 MONTH),'!','!!')) alt FROM `namen` n LEFT JOIN `aktfvs` af ON n.pat_id = af.pat_id LEFT JOIN `faelle` f ON n.pat_id = f.pat_id " & _
 "LEFT JOIN `dmpreihe` dr ON dr.pat_id = n.pat_id  AND (dr.Abk LIKE 'eDMPDM%' OR dr.Abk LIKE 'DMPDTYP%') LEFT JOIN (SELECT pat_id, zeitpunkt zp, wert FROM labor1a ln WHERE abkü RLIKE 'hba[c1]' AND ln.wert < 22 UNION SELECT 2a.pat_id, 2a.zeitpunkt zp, 2a.wert FROM labor2a 2a WHERE abkü RLIKE 'hba[c1]' AND 2a.wert < 22 ORDER BY pat_id,zp DESC) a1c ON n.pat_id = a1c.pat_id LEFT JOIN `anamnesebogen` a ON a.pat_id = n.pat_id WHERE a.tkz = 0 AND (instr(notiz,'DMP hier')> 0 OR dmpklass = 3) AND ISNULL(af.vknr) AND dokudatum > SUBDATE(NOW(),INTERVAL 9 MONTH) ORDER BY n.pat_id, MID(f.quartal,2) DESC, f.quartal DESC, dokudatum DESC, a1c.zp DESC) i GROUP BY pat_id ORDER BY MID(`Schein`,2) DESC, `Schein` DESC, name;"
 sql = _
"  SELECT" & vbCrLf & _
"   f.quartal `Schein`, f.Pat_ID" & vbCrLf & _
"   , CONCAT(" & vbCrLf & _
"     gesname(f.pat_id),', ',patalter(f.pat_id),'a, T: '" & vbCrLf & _
"      ,CONCAT_WS(',',n.privattel,n.privattel_2,n.privatmobil)) NAME" & vbCrLf & _
"    , DATE_FORMAT(dokudatum,'%d.%m.%y') DMPDoku" & vbCrLf & _
"    , IF(dokudatum > SUBDATE(CONCAT(YEAR(NOW()),'-',((MONTH(NOW())-1) DIV 3)*3+1,'-1'),INTERVAL 3 MONTH),'',IF(dokudatum > SUBDATE(CONCAT(YEAR(NOW()),'-',((MONTH(NOW())-1) DIV 3)*3+1,'-1'),INTERVAL 6 MONTH),'!','-')) fri" & vbCrLf & _
"    , h.letzter HbA1c" & vbCrLf & _
"    , DATE_FORMAT(h.lzp,'%d.%m.%y') `HbA1c-Zpkt`" & vbCrLf & _
"    , IF(h.lzp > SUBDATE(CONCAT(YEAR(NOW()),'-',((MONTH(NOW())-1) DIV 3)*3+1,'-1'),INTERVAL 0 MONTH),'',IF(h.lzp > SUBDATE(CONCAT(YEAR(NOW()),'-',((MONTH(NOW())-1) DIV 3)*3+1,'-1'),INTERVAL 3 MONTH),'!','!!')) alt" & vbCrLf & _
"   , REPLACE(REPLACE(COALESCE(GROUP_CONCAT(CONCAT(DATE_FORMAT(zp,' %d.%m.%y '), LEFT(raum,3),' ',LEFT(zusatz,15)) ORDER BY zp SEPARATOR ', ' LIMIT 10),''),CHR(10),''),CHR(13),'') Termin" & vbCrLf & _
"  FROM namen n" & vbCrLf & _
"  LEFT JOIN (SELECT f.*,ROW_NUMBER() OVER(PARTITION BY pat_id ORDER BY bhfb DESC) fzn FROM faelle f" & vbCrLf & _
"            )  f   ON f.Pat_ID=n.pat_id AND fzn=1" & vbCrLf & _
"  LEFT JOIN (SELECT dr.*,ROW_NUMBER() OVER(PARTITION BY pat_id ORDER BY dokudatum DESC) zn FROM dmpreihe dr" & vbCrLf & _
"              WHERE dr.dokudatum>NOW()-INTERVAL 9 MONTH AND dr.Abk RLIKE '^eDMPDM|^DMPDTYP'" & vbCrLf & _
"            ) dr ON dr.pat_id=n.pat_id AND dr.zn=1" & vbCrLf & _
"  LEFT JOIN lHbA1c h ON h.pat_id=n.pat_id" & vbCrLf & _
"  LEFT JOIN `anamnesebogen` a ON a.pat_id=n.pat_id" & vbCrLf & _
"  LEFT JOIN (SELECT ROW_NUMBER() OVER(PARTITION BY pid ORDER BY zp) rn, t.* FROM termine t" & vbCrLf & _
"              WHERE zp>=date(NOW()) group BY t.pid,DATE(t.zp),t.raum,t.zusatz) t ON t.pid=n.pat_id AND rn<10" & vbCrLf & _
"  WHERE a.Tkz = 0 AND n.dmpklass = 3 AND f.BhFB < QAnf() AND NOT ISNULL(dr.DokuDatum)" & vbCrLf & _
"  GROUP BY n.Pat_ID" & vbCrLf & _
"  ORDER BY MID(f.quartal,2) DESC, f.quartal DESC, n.nachNAME, n.vorname;"
' "     CONVERT(gesname(f.pat_id) USING LATIN1),', ',patalter(f.pat_id),'a, T: '" & vbCrLf & _

 myFrag rs, sql
 TabAusgeb rs, Me, True, , , , Array(0, 0, 55, 25, 0, 3, 0, 0, 100), True, "DMP-Wiedereinbestellungen"
End Sub ' WiedereinbestellungenDMP_Click

#If False Then
Private Sub WiedereinbestellungenDMP_Click()
 Dim sql$, Zp$, obDruck%, AusgStr$, grenze As Date
 Dim r1 As New ADODB.Recordset
 Dim r2 As New ADODB.Recordset
 Dim WDatei$
 On Error GoTo fehler
' SELECT Wert,Zeitpunkt FROM (SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"" AS NB FROM (SELECT n.Pat_ID AS Pat_ID,n.ZeitPunkt AS ZeitPunkt,n.FertigStGrad AS FertigStGrad,n.Abkü AS Abkü,l.Langtext AS Langtext,n.Wert AS Wert,n.Einheit AS Einheit,k.Kommentar AS Kommentar,n.AbsPos AS AbsPos,n.AktZeit AS AktZeit FROM (`laborlangtext` l INNER JOIN (laborkommentar k INNER JOIN `laborneu` n ON ((k.KommentarVW = n.KommentarVW))) ON ((l.LangtextVW = n.LangtextVW))) WHERE pat_id = 105) AS labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar, Normbereich AS NB FROM `laborxus` LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr WHERE pat_id = 105 AND NOT EXISTS (SELECT * FROM `laborneu` WHERE pat_id = 105 AND abkü = laborxwert.Abkü AND wert = laborxwert.wert AND zeitpunkt > laborxus.Eingang -3 AND zeitpunkt < laborxus.Eingang+6)) AS sql1 WHERE abkü = "HBA1C" ORDER BY zeitpunkt desc
 Call ProgStart
 grenze = QAnf(ZQuart(Now() - 90)) 'CDate("1.4.08")
 WDatei$ = pVerz & "Wiedereinbestellungen " & Format(Now, "d.m.yy hh.mm") & ".txt"
 Open WDatei For Output As #339
 Print #339, Right$(Space$(4) & "Pat_ID", 4) & " " & Left$("Nachname" & Space$(15), 15) & " " & Left$("Vorname" & Space$(9), 9) & " 'X'= DMP hier (Notiz-Eintrag / Versicherung)        letzter Fallbeginn  letztes HbA1c"
 Print #339, String$(110, "_")
 'r1.Open "SELECT pat_id, notiz FROM `namen` na WHERE notiz LIKE '%hier%'", dbv.wCn, adOpenStatic, adLockReadOnly
 'SELECT pat_id, notiz FROM `namen` na JOIN (SELECT pat_id FROM `dmpreihe` dr UNION SELECT pat_id FROM `namen` na WHERE notiz LIKE '%hier%' ORDER BY pat_id) AS innen USING (pat_id)
 'r1.Open "SELECT pat_id, na.nachname, na.vorname, lf.bhfb, notiz,rname FROM `namen` na JOIN (SELECT pat_id FROM `dmpreihe` dr " & _
         "WHERE dokudatum > adddate(NOW(),-365) UNION SELECT pat_id FROM _lfaelle f LEFT JOIN `diagnosen` d " & _
         "using (pat_id) WHERE icd LIKE 'E1%' AND NOT icd LIKE 'E16%' AND diagsicherheit <> 'A'  AND " & _
         "bhfb > adddate(NOW(),-180) ORDER BY pat_id) AS innen USING (pat_id) " & _
         "join lfaellev lf USING (pat_id) " & _
         "join `kassenliste` k ON k.ik = lf.ik " & _
         "join `anamnesebogen` an USING (pat_id) WHERE tkz=0 GROUP BY pat_id", dbv.wCn, adOpenStatic, adLockReadOnly '  ... UNION SELECT pat_id FROM `namen` na WHERE notiz LIKE '%hier%'
 myFrag r1, "SELECT pat_id, na.nachname, na.vorname, na.dmpklass, na.dmpbeg, lf.bhfb, notiz,rname FROM `namen` na JOIN (SELECT pat_id FROM `dmpreihe` dr WHERE dokudatum > adddate(NOW(),-365)  AND (dr.Abk LIKE 'eDMPDM%' OR dr.Abk LIKE 'DMPDTYP%') UNION SELECT pid pat_id " & vbCrLf & _
            "FROM _lfaelle f " & vbCrLf & _
            "LEFT JOIN diagview d ON d.pat_id = f.pid " & vbCrLf & _
            "WHERE gICD REGEXP '^E1[0-4]\.' AND mbhfb > adddate(NOW(),-180) ORDER BY pat_id) innen USING (pat_id) JOIN lfaellev lf USING (pat_id) JOIN `kassenliste` k ON k.ik = lf.ik JOIN `anamnesebogen` an USING (pat_id) WHERE tkz=0 GROUP BY pat_id", adOpenStatic, dbv.wCn, adLockReadOnly  '  ... UNION SELECT pat_id FROM `namen` na WHERE notiz LIKE '%hier%'
 Do While Not r1.EOF
'  IF obhierdmp(r1!notiz) THEN
   Set r2 = Nothing
   ' wegen falscher Fremdlabore gestrichen: einheit = '%'
'   r2.Open "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & r1!Pat_id & " AND abkü RLIKE '^hba[1c]' AND CAST(wert AS decimal) < 22 UNION SELECT * FROM `labor1a` WHERE pat_id = " & r1!Pat_id & " AND abkü RLIKE '^hba[1c]' AND CAST(wert AS decimal) < 22) i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb ORDER BY zeitpunkt DESC LIMIT 1", dbv.wCn, adOpenStatic, adLockReadOnly
   Set r2 = hollabor(r1!Pat_id, "HBA[1C]", 0, 22)
   Zp = vNS
   obDruck = True
   If Not r2.EOF Then
    Zp = r2!Zeitpunkt & " " & Right$(Space$(3) & r2!Wert, 3)
    If r2!Zeitpunkt >= grenze Then obDruck = False
   End If
   If obDruck Then
'    AusgStr = Right$(Space$(4) & r1!Pat_id, 4) & " " & LEFT(r1!Nachname & Space$(15), 15) & " " & LEFT(r1!Vorname & Space$(11), 11) & "   " & IIf(obhierdmp(r1!Notiz), "X", " ") & " (" & LEFT(IIf(ISNULL(r1!Notiz) OR LenB(r1!Notiz) = 0, r1!rname, replace$(replace$(r1!Notiz, vbCr, ""), vbLf, "")) & ")" & Space$(42), 42) & " " & LEFT(r1!BhFB & Space$(10), 10) & " " & Zp
    AusgStr = Right$(Space$(4) & r1!Pat_id, 4) & " " & Left$(r1!Nachname & Space$(15), 15) & " " & Left$(r1!Vorname & Space$(11), 11) & "   " & IIf(r1!dmpklass = 3, "X", " ") & " (" & Left$(IIf(IsNull(r1!Notiz) Or LenB(r1!Notiz) = 0, r1!rname, REPLACE$(REPLACE$(r1!Notiz, vbCr, ""), vbLf, "")) & ")" & Space$(42), 42) & " " & Left$(r1!BhFB & Space$(10), 10) & " " & Zp
'    Debug.Print AusgStr
    Me.Ausgeb AusgStr & vbCrLf & altAusgabe, True
    Print #339, AusgStr
    Print #339, String$(110, "_")
   End If
'  END IF
  r1.Move 1
 Loop
 Close #339
 zeigan WDatei, vbNormalFocus
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WiedereinbestellungenDMP_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' WiedereinbestellungenDMP_Click
#End If

' Funktionen für Arzthelferin und Arzt -> Hausärzte aus Listenausgabe_Ueberweiser einlesen
Private Sub Hausärzte_aus_Listenausgabe_Ueberweiser_einlesen_Click()
 Call doHAAkt(Me)
End Sub ' Hausärzte_aus_Listenausgabe_Ueberweiser_einlesen_Click

' Funktionen für Arzthelferin und Arzt -> Hausärzte mit alter KV-Nr ergänzen
Private Sub HausärztemitalterKVNrergänzen_Click()
 Dim fha As New frmalthae, rAf&, rs As New ADODB.Recordset
 ' erst mal die leeren KV-Nummern einfügen
 InsKorr DBCn, DBCnS, "INSERT INTO `althae` (kvnu,kvnr) SELECT kvnu, kvnr FROM (SELECT n.kvnr kvnu, LEFT(n.kvnr,2),'/',RIGHT(n.kvnr,5) kvnr, HAName hHA, CONCAT_WS(', ',l.name, l.vorname) lHA FROM `aktfvs` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN `aktlue` l ON n.kvnr = l.kvnro LEFT JOIN althae h ON n.kvnr = h.kvnu GROUP BY n.pat_id) innen WHERE (ISNULL(lha) OR lha='') AND (ISNULL(hha) OR hha='') AND kvnu <> '' AND NOT EXISTS (SELECT kvnu FROM althae WHERE kvnu = innen.kvnu)", rAf
' SET fha.datprimaryRS = n
' Call fha.vorbereit
 myFrag rs, "SELECT GROUP_CONCAT(kvnu) nrn FROM (SELECT n.kvnr kvnu, HAName hHA, CONCAT_WS(', ',l.name, l.vorname) lHA FROM `aktfvs` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN `aktlue` l ON n.kvnr = l.kvnro LEFT JOIN althae h ON n.kvnr = h.kvnu GROUP BY n.pat_id) innen WHERE (ISNULL(lha) OR lha='') AND (ISNULL(hha) OR hha='') AND kvnu <> ''"
 If LenB(rs!nrn) <> 0 Then
  If fha.Vorbereit(rs!nrn) Then
   fha.Show
   Exit Sub
  End If
 End If ' LenB(rs!nrn) <> 0 Then
 MsgBox "Keine Datensätze zu editieren!"
 Exit Sub
End Sub ' HausärztemitalterKVNrergänzen_Click

' Funktionen für Arzthelferin und Arzt -> Patientenliste für Hausarztmodell
Private Sub PatientenlistefürHausarztmodell_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT n.kvnr `KV-Nr.HA`, IF(l.name <> '' OR l.vorname <> '' AND NOT ISNULL(l.name),CONCAT_WS(',',l.name, LEFT(l.vorname,1), CONCAT(IF(l.telefon<>'','T.',''),l.telefon)), IF(h.nachname <> '' AND NOT ISNULL(h.nachname),CONCAT_WS(',',h.nachname,LEFT(h.vorname,1),CONCAT(IF(h.tel1<>'','T.',''),h.tel1)),'?')) Hausarzt, n.pat_id Pat_ID, CONCAT_WS(',*',CONCAT_WS(',', n.nachname, n.vorname),DATE_FORMAT(gebdat,'%d.%m.%y')) Patient, MID(n.notiz,instr(n.notiz,'HM '),IF(instr(n.notiz,'PG ')<>0,instr(n.notiz,'PG ')-2-instr(n.notiz,'HM '),LENGTH(n.notiz))) Hausarztmodell FROM `aktfvs` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN `aktlue` l ON n.kvnr = l.kvnro LEFT JOIN althae h ON n.kvnr = h.kvnu GROUP BY pat_id ORDER BY n.kvnr, patient"
 TabAusgeb rs, Me, True, , , , , True, "PatientenlisteFürHausarztmodell"
End Sub ' PatientenlistefürHausarztmodell_Click

' Funktionen für Arzthelferin und Arzt -> Patientenliste für Vollpauschale
Private Sub PatientenlistefürVollpauschale_Click()
 Dim rs As New ADODB.Recordset
#If ebmalt Then
 myFrag rs, "SELECT f.pat_id, CONCAT_WS(',', f.nachname, f.vorname) Pat, vpau, f.übwlanr, CONCAT_WS(',',arzt.nachname,arzt.vorname) ÜW, IF(nlart.niederlassungsart='Hausarzt','h',LEFT(nlart.niederlassungsart,1)) NLArt, arzt2.lanr LANR2, CONCAT_WS(',',arzt2.nachname,arzt2.vorname) ÜW2, IF(nlart2.niederlassungsart='Hausarzt','h',LEFT(nlart2.niederlassungsart,1)) nlart2 FROM (`aktfvs` JOIN `faelle` f ON `aktfvs`.fid = f.fid AND f.schgr BETWEEN 20 AND 29) LEFT JOIN (" & HADBName & ".arzt JOIN " & HADBName & ".nlart ON arzt.nlart_id = nlart.idnlart LEFT JOIN (" & HADBName & ".arzt_has_bs ahb LEFT JOIN " & HADBName & ".arzt_has_bs ahb2 ON ahb.bs_id = ahb2.bs_id AND ahb.arzt_id <> ahb2.arzt_id JOIN  (" & HADBName & ".arzt arzt2 JOIN " & HADBName & ".nlart nlart2 ON arzt2.nlart_id = nlart2.idnlart) ON ahb2.arzt_id = arzt2.idarzt   ) ON arzt.idarzt = ahb.arzt_id AND nlart2.Niederlassungsart = 'Facharzt' AND nlart.niederlassungsart = 'Hausarzt') ON übwlanr = arzt.lanr" & _
         " LEFT JOIN (SELECT `aktfvs`.pat_id, leistung vpau FROM `aktfvs` JOIN `leistungen` l ON `aktfvs`.fid = l.fid AND (leistung LIKE '031%' OR leistung LIKE '01210')) vpau ON f.pat_id = vpau.pat_id WHERE nlart.niederlassungsart = 'Facharzt' OR ISNULL(nlart.niederlassungsart) OR NOT ISNULL(arzt2.lanr) GROUP BY pat_id"
#Else
 myFrag rs, "SELECT f.pat_id, CONCAT_WS(',', f.nachname, f.vorname) Pat, vpau, f.übwlanr, CONCAT_WS(',',arzt.nachname,arzt.vorname) ÜW, IF(nlart.niederlassungsart='Hausarzt','h',LEFT(nlart.niederlassungsart,1)) NLArt, arzt2.lanr LANR2, CONCAT_WS(',',arzt2.nachname,arzt2.vorname) ÜW2, IF(nlart2.niederlassungsart='Hausarzt','h',LEFT(nlart2.niederlassungsart,1)) nlart2 FROM (`aktfvs` JOIN `faelle` f ON `aktfvs`.fid = f.fid AND f.schgr BETWEEN 20 AND 29) LEFT JOIN (" & HADBName & ".arzt JOIN " & HADBName & ".nlart ON arzt.nlart_id = nlart.idnlart LEFT JOIN (" & HADBName & ".arzt_has_bs ahb LEFT JOIN " & HADBName & ".arzt_has_bs ahb2 ON ahb.bs_id = ahb2.bs_id AND ahb.arzt_id <> ahb2.arzt_id JOIN  (" & HADBName & ".arzt arzt2 JOIN " & HADBName & ".nlart nlart2 ON arzt2.nlart_id = nlart2.idnlart) ON ahb2.arzt_id = arzt2.idarzt   ) ON arzt.idarzt = ahb.arzt_id AND nlart2.Niederlassungsart = 'Facharzt' AND nlart.niederlassungsart = 'Hausarzt') ON übwlanr = arzt.lanr" & _
         " LEFT JOIN (SELECT `aktfvs`.pat_id, leistung vpau FROM `aktfvs` JOIN `leistungen` l ON `aktfvs`.fid = l.fid AND (leistung RLIKE '^030[12]' OR leistung LIKE '01210')) vpau ON f.pat_id = vpau.pat_id WHERE nlart.niederlassungsart = 'Facharzt' OR ISNULL(nlart.niederlassungsart) OR NOT ISNULL(arzt2.lanr) GROUP BY pat_id"
#End If
 TabAusgeb rs, Me, True, , , , , True, "PatientenlisteFürVollpauschale"
End Sub ' PatientenlistefürVollpauschale_Click

' Funktionen für Arzthelferin und Arzt -> Dokumentnamenprüfung
Private Sub Dokumentnamenprüfung_Click()
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT b.pat_id, b.zeitpunkt, CONCAT(n.nachname, ', ',n.vorname) Name,  b.name Dokname " & _
       "FROM briefe b LEFT JOIN namen n ON b.pat_id = n.pat_id " & _
       "WHERE ((zuuml(b.name) NOT LIKE CONCAT('%',zuuml(n.nachname),'%') " & _
       "and zuuml(CONCAT('ZZZ',b.name)) NOT LIKE CONCAT('%',zuuml(n.nachname),'%')) " & _
       "or zuuml(b.name) NOT LIKE CONCAT('%',zuuml(n.vorname),'%')) " & _
       "and not b.name IN ('GDT Import Datei') AND NOT b.name LIKE 'Brief an %' AND b.name<>'CGM BMP gedruckt' ORDER BY b.pat_id DESC;"
 myFrag rs, sql
 TabAusgeb rs, Me, True, , , , , , , , , , "Dokumentnamenprüfung"
End Sub ' Dokumentnamenprüfung_Click

' Funktionen für Arzthelferin und Arzt -> Patienten mit uns als Hausarzt (vorher fehlende Hausärzte ergänzen!)
Private Sub FehlendeHausaerzte_Click()
'SELECT DISTINCT `namen`.pat_id, `namen`.nachname, `namen`.vorname FROM quelle.`namen` LEFT JOIN quelle.faelle ON faelle.pat_id = `namen`.pat_id WHERE (`namen`.kvnr = "" OR ISNULL(`namen`.kvnr)) AND faelle.schgr = 24 ORDER BY pat_id DESC;
 Call ProgStart
 Call doLdFH(Me)
 Call ProgEnde
End Sub ' FehlendeHausaerzte_Click

' Funktionen für Arzthelferin und Arzt -> fehlende Überweisungsscheine
Private Sub FehlendeÜberweisungsscheine_Click()
 Call ProgStart
 Call doFÜwS(Me)
 Call ProgEnde
End Sub ' FehlendeÜberweisungsscheine_Click

' 102
' Funktionen für Arzthelferin und Arzt -> falsche Karteikarteneinträge
Private Sub FalscheKarteikarteneinträge_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, AuffArtSql
 TabAusgeb rs, Me, True, , , , , , "Falsche_Karteikarteneinträge"
 Exit Sub ' FalscheKarteikarteneinträge_Click

#If False Then
 Dim ErgDat$
 ErgDat$ = pVerz & "FalscheKarteikarteneinträge.txt"
 Dim altArt$
 Dim rs As New ADODB.Recordset
 Call ProgStart
 Open ErgDat For Output As #327
 Print #327, "Falsche karteikarteneinträge vom " & Now() & ":"
 myFrag rs, "SELECT pat_id,zeitpunkt,art,inhalt FROM `eintraege` WHERE art NOT IN (" & artSpezEintr & "," & artSpezUS & "," & artSpezSonst & ") ORDER BY art, pat_id, zeitpunkt", adOpenStatic, DBCn
 Do While Not rs.EOF
  If rs!art <> altArt Then
   Print #327, String$(80, "_")
   altArt = rs!art
  End If
  Print #327, Right$(Space$(4) & rs!Pat_id, 4) & "|" & Left$(rs!Zeitpunkt & Space$(19), 19) & "|" & Left$(rs!art & Space$(10), 10) & "|" & Left$(rs!Inhalt, 50)
  rs.Move 1
 Loop
 Close #327
' zeigan ergdat
 Call GetWord
 With Wapp
   .Visible = True
   .WindowState = wdWindowStateMaximize
   .documents.Open ErgDat
   .activedocument.Range.Font.size = 9
   .Activate
 End With
#End If
End Sub ' FalscheKarteikarteneinträge_Click

' Funktionen für Arzthelferin und Arzt -> verhunzte Fotos einfärben
Private Sub VerhunzteFotosEinfärben_Click() ' zu knapp gespeicherte Fotos nochmal verbessern
Dim rDok As New ADODB.Recordset, IViewPfad$, KStr$, FPfad$, FNam$, DokPfad$, erg$, eingefärbt&
 Dim rsFiP As New ADODB.Recordset, Nam$, namspl$(), j&, sql$, nr&, altnam$, neunam$
 Dim fold As Folder, Fil As File
 Dim APfad$
 Dim DokName$
 ProgStart
 APfad = pügVerz & "VerhunzteFotosEinfärbenProt_" & Year(Now()) & "_" & Month(Now()) & "_" & Day(Now()) & "_" & Hour(Now()) & "_" & Minute(Now()) & ".txt"
 Open APfad For Output As #311
 Close #311
 zeigan APfad
 IViewPfad = getIViewPfad() ' Environ("ProgramFiles") & "\irfanview\i_view32.exe
 getDokPfad
 For Each Fil In FSO.GetFolder(pVerz & "eingelesen\" & Year(Now() - 7)).Files ' "eingelesen\2010").Files ' "eingelesen\" & YEAR(NOW() - 7)).Files
  If Fil.name Like "*Foto*.jpg" Then
   If Fil.size < 70000 Then
    Nam = Left$(Fil.name, Len(Fil.name) - 4)
    SplitNeu Nam, "_", namspl
    If UBound(namspl) > 0 Then
     If IsNumeric(namspl(UBound(namspl))) Then
     Nam = namspl(0)
     For j = 1 To UBound(namspl) - 1
      Nam = Nam & namspl(j)
     Next j
     End If
    End If
    If InStrB(Nam, "'") > 0 Then Nam = REPLACE$(Nam, "'", "''")
    sql = "SELECT * FROM `fotosinp`.`jpg` WHERE neuername LIKE '" & Nam & "%'"
    Set rsFiP = Nothing
    myFrag rsFiP, sql
    If rsFiP.BOF Then
     Open APfad For Append As #311
     Print #311, "Fehlt in fotosinp: " & Fil.size & " " & Fil.path
     Close #311
    Else ' IF Not IsNumeric(rsFiP!compression) OR rsFiP!compression < 20 THEN
     nr = nr + 1
     Me.Ausgeb nr & ": " & rsFiP!tpfad, 0
     If FSO.FileExists(rsFiP!tpfad) Then
      Set rDok = Nothing
nochmal:
      DokName = REPLACE$(REPLACE$(Nam, " ", "%"), "Wärmflasch,", "Wärmflasche,") & "%'"
      sql = "SELECT * FROM `dokumente` WHERE dokname LIKE '" & DokName
      myFrag rDok, sql
      If rDok.BOF Then
       Open APfad For Append As #311
       Print #311, "Fehlt in Tabelle `dokumente`: '" & DokName & "'"
       Close #311
       Set rDok = Nothing
'       GoTo nochmal
      Else
       FPfad = Fil.path
       FNam = Fil.name
'       KStr = "cmd /c ren """ & FPfad & """ """ & FPfad & "_kaputt"""
'       erg = Shell(KStr, vbNormalFocus)
       If FSO.FileExists(FPfad & "_kaputt") Then Kill FPfad & "_kaputt"
       Name FPfad As FPfad & "_kaputt"
       KStr = IViewPfad & " """ & rsFiP!tpfad & """ /convert=""" & FPfad & """ /jpgq=30 " & IIf(rsFiP!helligkeit <> 90, " /bright=" & rsFiP!helligkeit, vNS) & IIf(rsFiP!kontrast <> 90, " /contrast=" & rsFiP!kontrast, vNS)
       erg = Shell(KStr, vbNormalFocus)
       If FSO.FileExists(FPfad) Then
        If FSO.GetFile(FPfad).size >= FSO.GetFile(FPfad & "_kaputt").size * 0.5 Then ' 21.8.21 statt FileLen
         Kill FPfad & "_kaputt"
        Else
         Open APfad For Append As #311
         Print #311, "'" & FPfad & "' nicht mindestens doppelt so groß wie '" & FPfad & "_kaputt':" & vbCrLf & "(" & FSO.GetFile(FPfad).size & " vs. " & FSO.GetFile(FPfad).size & "_kaputt)" ' 21.8.21 statt Filelen
         Close #311
        End If
       Else
        Open APfad For Append As #311
        Print #311, "'" & FPfad & "' existiert nicht."
        Close #311
       End If
       Do While Not rDok.EOF
        DokPfad = REPLACE$(LCase$(rDok!DokPfad), "$\turbomed\dokumente", PcDokPfad)
        If FSO.FileExists(DokPfad) Then
'        KStr = "cmd /c ren """ & DokPfad & """ """ & DokPfad & "_kaputt"""
'        Shell KStr, vbNormalFocus
         Name DokPfad As DokPfad & "_kaputt"
'        KStr = "xcopy """ & FPfad & """ """ & DokPfad & """ /y /h /r /k /c"
'        erg = Shell(KStr, vbNormalFocus)
         FileCopy FPfad, DokPfad
         If FSO.FileExists(DokPfad) Then
          If FSO.GetFile(DokPfad).size >= FSO.GetFile(DokPfad & "_kaputt").size * 0.5 Then ' 21.8.21 statt FileLen
           Kill DokPfad & "_kaputt"
           eingefärbt = eingefärbt + 1
          Else
           Open APfad For Append As #311
           Print #311, "'" & DokPfad & "' nicht größer als '" & DokPfad & "_kaputt':" & vbCrLf & "(" & FSO.GetFile(DokPfad).size & " vs. " & FSO.GetFile(DokPfad).size & "_kaputt)" ' 21.8.21 statt FileLen
           Close #311
          End If
         Else
          Open APfad For Append As #311
          Print #311, "'" & DokPfad & "' existiert nach `Filecopy '" & FPfad & "' '" & DokPfad & "'` nicht."
          Close #311
         End If
        Else
         Open APfad For Append As #311
         Print #311, "'" & DokPfad & "' existiert von Haus aus nicht."
         Close #311
        End If
        rDok.MoveNext
       Loop
      End If
     Else
      Open APfad For Append As #311
      Print #311, "Die in `fotosinp` zu `neuername` '" & Nam & "' genannte Datei '" & rsFiP!tpfad & "' existiert nicht."
      Close #311
     End If
    End If
   End If
  End If
 Next Fil
 MsgBox "Fertig mit VerhunzteFotosEinfärben: " & eingefärbt & " Fotos eingefärbt"
End Sub ' VerhunzteFotosEinfärben_Click

' Funktionen für Arzthelferin und Arzt -> Anwaltsunterlagen für Pat. zusammenstellen
Private Sub Anwaltsunterlagen_für_Pat_zusammenstellen_Click()
 Call ProgStart
 Aktion = Anwalt
 Set pataw.hlese = Me
 pataw.obRueck = False
 obRueck = False
 pataw.ZeilenzahlL.Visible = True
 pataw.Zeilenzahl.Visible = True
 pataw.Show
 Call ProgEnde
' Call doPLZeinzeln
End Sub ' Anwaltsunterlagen_für_Pat_zusammenstellen_Click

' Funktionen für Arzthelferin und Arzt -> Sonderpatienten anzeigen
Private Sub SonderpatientenAnzeigen_Click()
 Dim rs As New ADODB.Recordset
 Static spa As New Sonderpatientauswahl
 Set spa.aufRufer = Me
 spa.Show 1
 myFrag rs, "SELECT Zeitpunkt,Art,Inhalt FROM `eintraege` WHERE pat_id = " & SpPat_id & " ORDER BY zeitpunkt DESC"
 TabAusgeb rs, Me, , , , , , , "Einträge für " & SpName
End Sub ' SonderpatientenAnzeigen_Click

' Funktionen für Arzthelferin und Arzt -> Hausarzt anzeigen
Private Sub Hausarzt_anzeigen_Click()
 Call ProgStart
 Set haanz = Nothing
 Set haanz.hlese = Me
 haanz.Show
End Sub ' Hausarzt_anzeigen_Click

' ... für &Arzt

#If mitab Then
' ...für Arzt -> Anamnesebogen (Diagnosen eingeben)
Private Sub ÜbertragenenAnamnesebogen_Click()
 anBogÜ.Show
End Sub ' ÜbertragenenAnamnesebogen_Click
#End If

' nicht sichtbar: ...für Arzt -> Anamnesebogen experimentell
'Private Sub Anamnesebogen_Click()
' anb.Show
'End Sub

' ...für Arzt -> Diabetes-Quartalsdiagnosen in Dauerdiagnosen umwandeln (manuell)
Private Sub DiabetesQuartalsdiagnosenInDauerdiagnosenUmwandeln_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT d.Pat_Id, GesNameG(d.pat_id) GesName, d.DiagDatum, d.ICD " & vbCrLf & _
            "FROM faelle f " & vbCrLf & _
            "LEFT JOIN diagview d ON f.fid = d.fid " & vbCrLf & _
            "LEFT JOIN diagview dd ON f.pat_id = dd.pat_id AND dd.gICD RLIKE '^E1[0-4]\.' AND dd.obdauer<>0 " & vbCrLf & _
            "WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND d.gICD RLIKE '^E1[0-4]\.' AND d.obdauer = 0 AND ISNULL(dd.pat_ID) " & vbCrLf & _
            "ORDER BY d.pat_id"
 TabAusgeb rs, Me, , , , , , , "Diabetes-Quartalsdiagnosen in Dauerdiagnosen umwandeln (manuell)", 1
End Sub ' DiabetesQuartalsdiagnosenInDauerdiagnosenUmwandeln_Click

' 5.7.10: jetzt nicht mehr nötig, in PatListe integriert
' nicht sichtbar: ...für Arzt -> Alle &DMP-Dokumente an Hausärzte in p:\zufaxen erstellen
'Private Sub DMPSend_Click()
' Call ProgStart
' Call alleDMPs(Me)
' Call ProgEnde
'End Sub ' DMPSend_Click

' ...für Arzt -> alle DMP-Dokumente an Hausärzte faxen
Private Sub AlleDMPanHA_Click() ' Alle DMP-Dokumente an Hausärzte faxen
#If False Then
 Dim pad As New PatListe
 pad.PLArt = artDMP
 Set pad.hlese = Me
 pad.Show
#Else
 Call ProgStart
 Set pal = Nothing
 pal.PLArt = artDMP
 Set pal.hlese = Me
 Me.Hide
 pal.Show
 Call ProgEnde
#End If
End Sub ' AlleDMPanHA_Click


' ...für Arzt -> DMP-Dokumente an HA-Nachweis
Private Sub DMP_Dokumente_an_HA_Nachweis_Click() ' s.DMPFüll
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT f0.pat_id, gesname(f0.pat_id) name, " & vbCrLf & _
       "f.icd, n.getha0 Üwnnr, CONCAT_WS(', ',h.name, h.vorname, h.titelt) Adressat " & vbCrLf & _
       ", NOT ISNULL(f.pat_id) AND ISNULL(dt.titel) AND n.dmpklass = 2 AND f.icd RLIKE '^E1[0-4]\.' AND h.kvnr<>'' obfax" & vbCrLf & _
       ", NOT ISNULL(dt.titel) `keinB.anHA`, NOT ISNULL(hk.kvnr) HAwillkeinDMP, n.dmpklass `DKlass(2=HA)`, f.ICD, h.KVNr " & vbCrLf & _
       "From aktfv f0 " & vbCrLf & _
       "LEFT JOIN aktfaellev f USING(fid)" & vbCrLf & _
       "LEFT JOIN namen n ON f0.pat_id = n.pat_id " & vbCrLf & _
       "LEFT JOIN liuez h ON h.kvnri=n.getha0 AND n.getha0<>0 " & vbCrLf & _
       "LEFT JOIN `hakeinedmpinfo` hk ON n.getha0=hk.kvnr " & vbCrLf & _
       "LEFT JOIN `desktop` dt ON n.pat_id = dt.pat_id AND dt.titel LIKE '%kein%bericht%' " & vbCrLf & _
       "GROUP BY f0.pat_id;"
'       "WHERE ISNULL(dt.titel) AND n.dmpklass = 2 AND f.icd RLIKE '^E1[0-4]\.' AND h.kvnr<>'' " & vbCrLf & _

 myFrag rs, sql
 TabAusgeb rs, Me, True, , , , , , , , , , "DMP_Dokumente_an_HA_Nachweis_Click (DMP-Klass 2=relevant)"
End Sub ' DMP_Dokumente_an_HA_Nachweis_Click

' ...für Arzt -> Kontrolllisten für DMP HA
Private Sub Kontrolllisten_für_DMP_HA_Click()
 Dim rdh As New ADODB.Recordset
 Dim altgetha&
 Const ZielVz$ = "zufaxen\DMP\"
 Shell ("cmd /c del /q " & pVerz & ZielVz & "*.*")
 Const qd$ = "Kontroll"
 Dim Ziel$
 sql = _
 "SELECT * FROM (" & vbCrLf & _
 "SELECT sum(vdoku) OVER(PARTITION BY getha0) vzahl, COUNT(0) OVER(PARTITION BY getha0) pzahl, i.* FROM (" & vbCrLf & _
 "SELECT f.pat_id, gesnameg(f.pat_id) gesnam " & vbCrLf & _
 ",COALESCE((SELECT CASE WHEN icd REGEXP '^E10' THEN 'Typ 1' WHEN icd REGEXP '^E11' THEN 'Typ 2' WHEN '^O24' THEN 'Gest' ELSE 'Sonst' END " & vbCrLf & _
 "FROM diagview " & vbCrLf & _
 "WHERE Pat_ID=f.Pat_ID AND (gicd REGEXP '^E1[0-4]\.|^R73' OR (icd='O24.4' AND f6010=0 AND diagsicherheit IN ('G',' ') AND diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))) " & vbCrLf & _
 "ORDER BY ICD " & vbCrLf & _
 "LIMIT 1),'fehlt') typ " & vbCrLf & _
 ", COALESCE((SELECT 1 FROM briefe WHERE pat_id=f.pat_id AND name RLIKE 'dmp[ -]doku' AND zeitpunkt > IF((SELECT COUNT(0) FROM eintraege WHERE pat_id=f.pat_id AND zeitpunkt>qanf())>0,qanf(),qbeg(SUBDATE(NOW(),111))) LIMIT 1),0) vdoku" & vbCrLf & _
 ",n.getha0 " & vbCrLf & _
 ",h.* " & vbCrLf & _
 "FROM faelle f " & vbCrLf & _
 "LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
 "LEFT JOIN liuez h ON h.kvnri=n.getha0 AND n.getha0<>0 " & vbCrLf & _
 "LEFT JOIN `desktop` dt ON n.pat_id = dt.pat_id AND dt.titel LIKE '%kein%bericht%' " & vbCrLf & _
 "WHERE f.quartal IN ( " & vbCrLf & _
 "(SELECT CONCAT((MONTH(CURRENT_TIMESTAMP() - INTERVAL 21 DAY) - 1) DIV 3 + 1, YEAR(CURRENT_TIMESTAMP() - INTERVAL 21 DAY))), " & vbCrLf & _
 "(SELECT CONCAT((MONTH(CURRENT_TIMESTAMP() - INTERVAL 101 DAY) - 1) DIV 3 + 1, YEAR(CURRENT_TIMESTAMP() - INTERVAL 101 DAY)))) " & vbCrLf & _
 "AND f.schgr<>'90' AND f.`GOÄKatNr` NOT IN ('40','41') AND f.nachname<>'Bereitschaftsdienst' " & vbCrLf & _
 "AND ISNULL(dt.titel) AND n.dmpklass=2 " & vbCrLf & _
 "AND getha0<>0) i) i " & vbCrLf & _
 "WHERE vzahl<pzahl " & vbCrLf & _
 "GROUP BY pat_id " & vbCrLf & _
 "ORDER BY getha0,typ,gesnam"
 ' AND exists (SELECT 1 FROM faxeinp.outa WHERE docname LIKE 'dmp-rückfrage%' AND erfolg=0 AND rcfax=i.fax)
 ' "WHERE vzahl>0 AND COALESCE((SELECT erfolg FROM faxeinp.outa a WHERE rcfax=i.fax AND docname LIKE 'DMP-Rückfrage%' AND NOT EXISTS (SELECT 1 FROM faxeinp.inca WHERE tsid=a.rcfax AND transe>a.transe) LIMIT 1),0)=1 " & vbCrLf & _
 ' WHERE vdoku AND COALESCE((SELECT erfolg FROM faxeinp.outa a WHERE rcfax=i.fax AND docname LIKE 'DMP-Rückfrage%' AND NOT EXISTS (SELECT 1 FROM faxeinp.inca WHERE tsid=a.rcfax AND transe>a.transe) LIMIT 1),0)=1 " & vbCrLf & _

 myFrag rdh, sql
 Do While True
  If Not rdh.EOF Then
   If rdh!getHA0 = altgetha Then GoTo w1
   If altgetha = 0 Then GoTo w2:
  End If
  Close #312
  FileJoin vVerz & qd & "3.rtf", Ziel
w2:
  If rdh.EOF Then Exit Do
  Ziel = pVerz & ZielVz & "DMP-Rückfrage an Dr.med." & rdh!name & ", " & rdh!Vorname & " an Fax " & rdh!fax & ".rtf"
  syscmd 4, Ziel
  FSO.CopyFile vVerz & qd & "1.rtf", Ziel, True
  Open Ziel For Append As #312
  Print #312, Format(Now(), "dd.mm.yy")
  Close #312
  FileJoin vVerz & qd & "2b.rtf", Ziel
  Open Ziel For Append As #312
  altgetha = rdh!getHA0
w1:
  Print #312, "\pard\plain \s2\ql \li0\ri0\keepn\widctlpar\intbl\aspalpha\aspnum\faauto\outlinelevel1\adjustright\rin0\lin0 \f1\fs24\lang1031\langfe1031\cgrid\langnp1031\langfenp1031 {" & umlweg(rdh!gesnam) & "\cell }\pard\plain \ql \li0\ri0\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin0\lin0 \fs20\lang1031\langfe1031\cgrid\langnp1031\langfenp1031 {\f1\fs24 " & rdh!Typ & "\cell " & IIf(rdh!vdoku, "\u8730\'76", "") & "\cell }\pard \ql \li0\ri0\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\f1\fs24 \trowd \trgaph70\trleft-70\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 \trbrdrh\brdrs\brdrw10 \trbrdrv\brdrs\brdrw10 \trftsWidth1\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddfr3 \clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb\clftsWidth3\clwWidth7630 \cellx7560\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl " & _
   "\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb\clftsWidth3\clwWidth900 \cellx8460\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb\clftsWidth3\clwWidth682 \cellx9142 \row }"
  rdh.MoveNext
 Loop
 syscmd 5
 MsgBox "Fertig mit DMP-Rückfragen in Verzeichnis " & pVerz & ZielVz
End Sub ' Kontrolllisten_für_DMP_HA_Click

' ...für Arzt -> Unverwertbare DMP-Einträge
Private Sub UnverwertbareDMPEinträge_Click() ' Unverwertbare DMP-Einträge
 Dim rs As New ADODB.Recordset, rez As New ADODB.Recordset, Pat_id&, AusgStr$, TA1$, STA1$(), i&
 Dim SpMin%(2)
 SpMin%(0) = 6
 Call ProgStart
' Open DatNam For Output AS #327
 myFrag rs, "SELECT pat_id, dmpklass, dmpbeg,dmpkhkklass, DMPKHKBeg, DMPCopdKlass, DmpCOPDBeg, DMPABKlass, DMPABBeg, REPLACE(notiz,CONCAT(char(13),char(10)),0) Notiz FROM `namen` WHERE notiz LIKE '%dmp%' AND NOT notiz = CONCAT('DMP NEIN',char(13),char(10)) ORDER BY pat_id DESC"
 TabAusgeb rs, Me, True, , , , , , "Unverwertbare DMP-Einträge"
End Sub ' UnverwertbareDMPEinträge_Click ' Unverwertbare DMP-Einträge

' ...für Arzt -> DMP hier Liste
Private Sub DMPhierListe_Click()
 Call ProgStart
 Set pal.hlese = Me
 Set pal = Nothing
 pal.PLArt = artPat
 Me.Hide
 pal.Show
 Call ProgEnde
End Sub ' DMPhierListe_Click

' ...für Arzt -> DMP hier Liste ohne Pat.mit Terminen
Private Sub DMPhierListeoT_Click()
 Call ProgStart
 Set pal.hlese = Me
 Set pal = Nothing
 pal.PLArt = artPat
 pal.ohneTermine = True
 Me.Hide
 pal.Show
 Call ProgEnde
End Sub ' DMPhierListe_Click

' ...für Arzt -> DMP KHK Asthma
Private Sub DMPKHKAsthma_Click()
 Dim rs As New ADODB.Recordset
 Dim sql$
 Call ProgStart
 sql = "SELECT n.pat_id,gesnameg(n.pat_id), f.form_abk, f.zeitpunkt " & vbCrLf & _
       "FROM `namen` n " & vbCrLf & _
       "LEFT JOIN `formular` f ON n.pat_id = f.pat_id " & vbCrLf & _
       "WHERE ((form_abk LIKE 'dmpkhk%' OR form_abk LIKE 'dmpab%' OR form_abk LIKE 'dmpcopd%') " & vbCrLf & _
       "   OR (form_abk LIKE 'edmp%' AND feld = 'Einschreibung' AND Feldnr IN (2,3,4))) " & vbCrLf & _
       "   AND f.zeitpunkt > SUBDATE(NOW(),INTERVAL 365 DAY) " & vbCrLf & _
       "GROUP BY n.pat_id,form_abk, DATE(zeitpunkt)"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , "DMP KHK Asthma " & Format(Now(), "dd.mm.YYYY")
End Sub ' DMPKHKAsthma_Click

' ...für Arzt -> Duplex Kontrollieren
Private Sub DuplexKontrollieren_Click()
 Call ProgStart
 Call doDuplexkontrollieren
 Call ProgEnde
End Sub ' DuplexKontrollieren_Click

' ...für Arzt -> Hausärzte mit DMP-Patienten
Private Sub HausärztemitDMPPatienten_Click()
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT CONCAT(LEFT(CONCAT(h.kvnr,'        '),8),'| '," & vbCrLf & _
       "LEFT(CONCAT(h.titelt,h.Name,' ',h.Vorname,'                           '),36),'|'," & vbCrLf & _
       "RIGHT(CONCAT('                      ',fax),21),' | '," & vbCrLf & _
       "LEFT(CONCAT(h.ort,', ',strasse,', ',fachgruppe,'                            '),46),'|') feld" & vbCrLf & _
       "FROM faelle f" & vbCrLf & _
 "LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
 "LEFT JOIN liuez h ON CAST(n.getha0 As Integer) = h.kvnri " & vbCrLf & _
 "WHERE bhfb > SUBDATE(NOW(),INTERVAL 1 YEAR) AND n.dmpklass=2 AND getha0<>0 AND instr(fachgruppe,'Diabet')=0 " & vbCrLf & _
 "GROUP BY h.fax;"
 myFrag rs, sql
 TabAusgeb rs, Me, , , True, , , , "HA-DMP.txt", , , , "Hausaerzte mit Patienten im DMP"
End Sub ' HausärztemitDMPPatienten_Click

' ...für Arzt -> DMP-Brief-Leistungs-Doku
Private Sub DMPLeiDok_Click()
 Call ProgStart
 Call alleDMPLeiDok(Me, 1)
 Call alleDMPLeiDok(Me, 2)
' Call ProgEnde
End Sub ' DMPLeiDok_Click

' ...für Arzt -> Porto abrechnen
Private Sub doPorto_Click()
 Call ProgStart
 Call dodoPorto
 Call ProgEnde
End Sub ' doPorto_Click

' ...für Arzt -> Leistungen zu Arztbriefen bei Kassenpatienten in u:\tmimport\ dokumentieren
Private Sub BriefeLeiDok_Click()
 Call ProgStart
 Call tuBriefeLeiDok(Me)
 Call ProgEnde
End Sub ' BriefeLeiDok_Click

' ...für Arzt -> Pathologische Laborwerte anschauen
Private Sub PathLabAnschau_Click()
 Dim ePL As New PathLabForm
 Set ePL.eLese = Me
 ePL.Show
End Sub ' PathLabAnschau_Click

' ...für Arzt -> Vorhandene Briefe korrigieren
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
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VorhandeneBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' VorhandeneBriefe_Click

' ...für Arzt -> Briefe zu Patienten, deren erster Brief schon > 1a her ist
Private Sub doFollowUp_Click() ' &Briefe zu Patienten, deren erster Brief schon > 1a her ist
 Call ProgStart
 Call dodoFollowUp(Me)
 Call ProgEnde
End Sub ' doFollowUp_Click

' ...für Arzt -> Briefe zu Pat., zu denen bisher noch kein Brief geschrieben wurde
Private Sub UngeschriebeneBriefe_Click()
 Call ProgStart
 Call doUngeschriebeneBriefe(Me)
 Call ProgEnde
End Sub ' UngeschriebeneBriefe_Click

' ...für Arzt -> Restliche Briefe für aktuelles Quartal schreiben
Private Sub RestlicheBriefe_Click()
 Call ProgStart
 Aktion = RestlicheBriefe
 Set pataw.hlese = Me
 pataw.Show
' Aktion = nix
 Call ProgEnde
End Sub ' RestlicheBriefe_Click

' ...für Arzt -> Briefe zu Patienten mit Berichtspflicht schreiben
Private Sub BriefeBerichtspflicht_Click()
 Call doBriefeBerichtspflicht
End Sub ' BriefeBerichtspflicht_Click

' ...für Arzt -> "Briefe zu Pat_ID-Liste aus Datei schreiben
Private Sub BriefeZuListeSchreiben_Click()
 Dim bzl As New Lade
 bzl.FenArt = 1
 bzl.Show
End Sub ' BriefeZuListeSchreiben_Click

' ...für Arzt -> Brief schreiben
Private Sub BriefSchreiben_Click() ' Brief schreiben
 Call ProgStart
 Aktion = Briefschreiben
 Set pataw.hlese = Me
 pataw.briefneu = False
 pataw.Show
' Aktion = nix
 Call ProgEnde
End Sub ' BriefSchreiben_Click

' ...für Arzt -> "Brief schreiben neu
Private Sub Briefschreibenneu_Click()
 Call ProgStart
 Aktion = Briefschreiben
 Set pataw.hlese = Me
 pataw.briefneu = True
 pataw.Show
' Aktion = nix
 Call ProgEnde
End Sub ' Briefschreibenneu_Click

Private Sub Briefnochmal_Click()
 Call ProgStart
 If Me.pataw.Pat_id = "" Then Me.pataw.Pat_id = 681
 Call tubriefStandalone(Me.pataw.Pat_id, 0, "", "", 0, True, True)
' Aktion = nix
 Call ProgEnde
End Sub ' Briefnochmal_Click

' ...für Arzt -> Brief zu letztem Import schreiben
Private Sub BriefImport_Click()
 Call tubriefStandalone(CStr(lzPID), 0, , Me.pataw.Vorlage)
End Sub ' Sub BriefImport_Click

' ...für Arzt -> Brief ohne Maske schreiben
Private Sub BriefOhneMaske_Click()
 Dim erg$
 erg = InputBox("Bitte Pat_ID eingeben:")
 If IsNumeric(erg) Then
  Call tubriefStandalone(CLng(erg), 0)
 End If
End Sub  ' BriefOhneMaske_Click

' ...für Arzt -> Liste der fehlenden Dokumente
Private Sub ListeDerFehlendenDokumente_Click()
 Call ProgStart
 Call doLdFD
 Call ProgEnde
End Sub ' ListeDerFehlendenDokumente_Click

' ...für Arzt -> Nach&zuholende Laborimporte
Private Sub NachzuholendeLaborimporte_Click()
 Dim rs As ADODB.Recordset, lfdnr&
 Open snst.DateiNachzuholen For Output As #301
 Call ProgStart
 Set rs = myEFrag("SELECT DISTINCT pfad,lwerte,laborxus.pat_id,eingang,auftragsnummer, `namen`.nachname, `namen`.vorname, laborxus.refnr FROM (`laborxus` LEFT JOIN laborxeingel ON laborxus.datid = laborxeingel.datid) LEFT JOIN `namen` ON laborxus.pat_id = `namen`.pat_id WHERE afn = 0 AND zdip = 0 AND NOT laborxus.pat_id = 0 AND zdüp>0 ORDER BY refnr;")
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
End Sub ' NachzuholendeLaborimporte_Click

' ...für Arzt -> Banksachen

' ...für Arzt -> Banksachen -> Lfd.Kosten
Private Sub LfdKosten_Click()
 Dim pwd$
 pwd = InputBox("Bitte M-Net-Kennwort eingeben", "Rückfrage")
 If pwd <> "17raga" Then Exit Sub
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT COUNT(0) zahl, LEFT(CAST(buchungstag AS char),7) Monat, ROUND(SUM(betrag)) `Kosten` FROM `konten`.`girokonten` g WHERE kontonummer = '7710127' AND betrag < 0 AND CONCAT(`verwendungszweckzeile 1`,`verwendungszweckzeile 2`) NOT LIKE '%Fehlüberweisung%' AND NOT (`begünstigter/Absender - Kontonummer` IN ('0297626808', '0006097316','0230113763','2000573363') OR `verwendungszweckzeile 1`like 'entnahme%') GROUP BY LEFT(buchungstag,7) ORDER BY buchungstag DESC"
 TabAusgeb rs, Me, , , , , , , "Laufende Kosten der Praxis"
End Sub ' LfdKosten_Click

' ...für Arzt -> Banksachen -> Lfd.Kosten mit &Betrag
Private Sub LfdKostenMitBetrag_Click()
 Dim pwd$
 pwd = InputBox("Bitte M-Net-Kennwort eingeben", "Rückfrage")
 If pwd <> "17raga" Then Exit Sub
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT COUNT(0) zahl, LEFT(CAST(buchungstag AS char),7) Monat, Buchungstag, ROUND(SUM(betrag)) `Kosten`, CONCAT(`verwendungszweckzeile 1`, `verwendungszweckzeile 2`, `verwendungszweckzeile 3`, `verwendungszweckzeile 4`, `verwendungszweckzeile 5`, `verwendungszweckzeile 6`) Verwendungszweck, Myid, EingID, Buchungstext, Kategorie, Unterkategorie, Textschlüssel, `Begünstigter/Absender - Name` Begünstigter, `Begünstigter/Absender - Bankleitzahl` BLZ, `Begünstigter/Absender - Kontonummer` Kto, Wertstellungstag FROM `konten`.`girokonten` g WHERE kontonummer = '7710127' AND betrag < 0 AND CONCAT(`verwendungszweckzeile 1`,`verwendungszweckzeile 2`) NOT LIKE '%Fehlüberweisung%' AND NOT (`begünstigter/Absender - Kontonummer` IN ('0297626808', '0006097316','0230113763','2000573363') OR `verwendungszweckzeile 1`like 'entnahme%') GROUP BY myid ORDER BY LEFT(buchungstag,7) DESC, kosten"
 TabAusgeb rs, Me, , , , , , , "Laufende Kosten der Praxis nach Monat und Betrag"
End Sub ' LfdKostenMitBetrag_Click

' ...für Arzt -> Banksachen -> KV-Überw
Private Sub KVÜberw_Click()
 Dim pwd$
 pwd = InputBox("Bitte M-Net-Kennwort eingeben", "Rückfrage")
 If pwd <> "17raga" Then Exit Sub
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT LEFT(CAST(buchungstag AS char),10) Buchungstag, ROUND(SUM(betrag)) `Einnahme`, CONCAT(`verwendungszweckzeile 1`, `verwendungszweckzeile 2`, `verwendungszweckzeile 3`, `verwendungszweckzeile 4`, `verwendungszweckzeile 5`, `verwendungszweckzeile 6`) Verwendungszweck, Myid, EingID, Buchungstext, Kategorie, Unterkategorie, Textschlüssel, `Begünstigter/Absender - Name` Begünstigter, `Begünstigter/Absender - Bankleitzahl` BLZ, `Begünstigter/Absender - Kontonummer` Kto, Wertstellungstag FROM `konten`.`girokonten` g WHERE kontonummer = '7710127' AND betrag > 0 AND CONCAT(`verwendungszweckzeile 1`,`verwendungszweckzeile 2`) NOT LIKE '%Fehlüberweisung%' AND NOT (`begünstigter/Absender - Kontonummer` IN ('0297626808', '0006097316','0230113763','2000573363')  OR `verwendungszweckzeile 1`like 'entnahme%') AND `begünstigter/absender - name` LIKE 'kvb oberbay%' GROUP BY myid ORDER BY buchungstag DESC"
 TabAusgeb rs, Me, , , , , , , "KV-Überweisungen"
End Sub ' KVÜberw_Click

' ...für Arzt -> Banksachen -> Lfd.Kosten &Eigenbetrieb
Private Sub LfdKostenEigenbetrieb_Click()
 Dim pwd$
 pwd = InputBox("Bitte M-Net-Kennwort eingeben", "Rückfrage")
 If pwd <> "17raga" Then Exit Sub
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT COUNT(0) zahl, LEFT(CAST(buchungstag AS char),7) Monat, ROUND(SUM(betrag)) `Kosten` FROM `konten`.`girokonten` g WHERE kontonummer = '6097316' AND betrag < 0 AND CONCAT(`verwendungszweckzeile 1`,`verwendungszweckzeile 2`) NOT LIKE '%Fehlüberweisung%' AND NOT (`begünstigter/Absender - Kontonummer` IN ('0297626808', '0006097316','0230113763','2000573363') OR `verwendungszweckzeile 1`like 'entnahme%'  OR `verwendungszweckzeile 1`like '%aufbank%') GROUP BY LEFT(buchungstag,7) ORDER BY buchungstag DESC"
 TabAusgeb rs, Me, , , , , , , "Laufende Kosten Eigenbetrieb"
End Sub ' LfdKostenEigenbetrieb_Click

' ...für Arzt -> Banksachen -> Lfd.Kosten &Eigenbetr.m.Betrag
Private Sub LfdKostenEigenbetrmBetrag_Click()
 Dim pwd$
 pwd = InputBox("Bitte M-Net-Kennwort eingeben", "Rückfrage")
 If pwd <> "17raga" Then Exit Sub
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT COUNT(0) zahl, LEFT(CAST(buchungstag AS char),7) Monat, LEFT(CAST(buchungstag AS char),10) Buchungstag, ROUND(SUM(betrag)) `Kosten`, CONCAT(`verwendungszweckzeile 1`, `verwendungszweckzeile 2`, `verwendungszweckzeile 3`, `verwendungszweckzeile 4`, `verwendungszweckzeile 5`, `verwendungszweckzeile 6`) Verwendungszweck, Myid, EingID, Buchungstext, Kategorie, Unterkategorie, Textschlüssel, `Begünstigter/Absender - Name` Begünstigter, `Begünstigter/Absender - Bankleitzahl` BLZ, `Begünstigter/Absender - Kontonummer` Kto, Wertstellungstag FROM `konten`.`girokonten` g WHERE kontonummer = '6097316' AND betrag < 0 AND CONCAT(`verwendungszweckzeile 1`,`verwendungszweckzeile 2`) NOT LIKE _utf8'%Fehlüberweisung%' AND NOT (`begünstigter/Absender - Kontonummer` IN ('0297626808', '0006097316','0230113763','2000573363') OR `verwendungszweckzeile 1`like 'entnahme%' OR `verwendungszweckzeile 1`like '%aufbank%') GROUP BY myid ORDER BY LEFT(buchungstag,7) DESC, kosten"
 TabAusgeb rs, Me, , , , , , , "Laufende Kosten Eigenbetrieb mit Betrag"
End Sub ' LfdKostenEigenbetrmBetrag_Click

' ...für Arzt -> Banksachen -> Lfd.Kosten PGiro
Private Sub LfdKostenPGiro_Click()
 Dim pwd$
 pwd = InputBox("Bitte M-Net-Kennwort eingeben", "Rückfrage")
 If pwd <> "17raga" Then Exit Sub
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT COUNT(0) zahl, LEFT(CAST(buchungstag AS char),7) Monat, LEFT(CAST(buchungstag AS char),10) Buchungstag, ROUND(SUM(betrag)) `Kosten`, CONCAT(`verwendungszweckzeile 1`, `verwendungszweckzeile 2`, `verwendungszweckzeile 3`, `verwendungszweckzeile 4`, `verwendungszweckzeile 5`, `verwendungszweckzeile 6`) Verwendungszweck, Myid, EingID, Buchungstext, Kategorie, Unterkategorie, Textschlüssel, `Begünstigter/Absender - Name` Begünstigter, `Begünstigter/Absender - Bankleitzahl` BLZ, `Begünstigter/Absender - Kontonummer` Kto, Wertstellungstag FROM `konten`.`girokonten` g WHERE kontonummer = '297626808' AND betrag < 0 AND CONCAT(`verwendungszweckzeile 1`,`verwendungszweckzeile 2`) NOT LIKE _utf8'%Fehlüberweisung%' AND NOT (`begünstigter/Absender - Kontonummer` IN ('0297626808', '0006097316','0230113763','2000573363') OR `verwendungszweckzeile 1`like 'entnahme%' OR `verwendungszweckzeile 1`like '%aufbank%') GROUP BY myid ORDER BY LEFT(buchungstag,7) DESC, kosten"
 TabAusgeb rs, Me, , , , , , , "Laufende Kosten PGiro"
End Sub ' LfdKostenPGiro_Click

' ...für Arzt -> Banksachen -> Lfd.Kosten PGiro mit Betrag
Private Sub LfdKostenPGiromBetrag_Click()
 Dim pwd$
 pwd = InputBox("Bitte M-Net-Kennwort eingeben", "Rückfrage")
 If pwd <> "17raga" Then Exit Sub
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT COUNT(0) zahl, LEFT(CAST(buchungstag AS char),7) Monat, Buchungstag, ROUND(SUM(betrag)) `Kosten`, CONCAT(`verwendungszweckzeile 1`, `verwendungszweckzeile 2`, `verwendungszweckzeile 3`, `verwendungszweckzeile 4`, `verwendungszweckzeile 5`, `verwendungszweckzeile 6`) Verwendungszweck, Myid, EingID, Buchungstext, Kategorie, Unterkategorie, Textschlüssel, `Begünstigter/Absender - Name` Begünstigter, `Begünstigter/Absender - Bankleitzahl` BLZ, `Begünstigter/Absender - Kontonummer` Kto, Wertstellungstag FROM `konten`.`girokonten` g WHERE kontonummer = '297626808' AND betrag < 0 AND CONCAT(`verwendungszweckzeile 1`,`verwendungszweckzeile 2`) NOT LIKE '%Fehlüberweisung%' AND NOT (`begünstigter/Absender - Kontonummer` IN ('0297626808', '0006097316','0230113763','2000573363') OR `verwendungszweckzeile 1`like 'entnahme%') GROUP BY myid ORDER BY LEFT(buchungstag,7) DESC, kosten"
 TabAusgeb rs, Me, , , , , , , "Laufende Kosten PGiro mit Betrag"
End Sub ' LfdKostenPGiromBetrag_Click


' ...für Arzt -> Faxnachweis
Private Sub Faxnachweis_Click()
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT docname, rcname, rcfax, transe, gesname(o.pid) Name, o.pid, submid FROM faxeinp.outa o LEFT JOIN namen n ON o.pid = n.pat_id WHERE erfolg<>'0' ORDER BY transe DESC LIMIT 2500"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , , , , , "Nachweis übermittelter Faxe"
End Sub ' Faxnachweis_Click

' ...für Arzt -> Faxwarteschlange
Private Sub Faxwarteschlange_Click()
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT cdateidatum,original, telnr FROM faxeinp.spool s ORDER BY cdateizeit DESC"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , , , , , "Faxwarteschlange auf " & LiName & " (Fritzcard)"
End Sub ' Faxwarteschlange_Click

' ...für Arzt -> Faxe gescheitert
Private Sub Faxe_gescheitert_Click()
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT docname, rcname, rcfax, transe, gesname(o.pid) Name, o.pid, submid FROM faxeinp.outa o LEFT JOIN namen n ON o.pid = n.pat_id WHERE erfolg='0' ORDER BY transe DESC LIMIT 2500"
' sql = "SELECT titel, submt, submid, docname,fsize,rcfax, rcname,transe, pid FROM faxeinp.outf o ORDER BY submt DESC"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , , , , , "Faxe gescheitert auf " & LiName & " (Fritzcard)"
End Sub ' Faxe_gescheitert_Click

' ...für Arzt -> Pat. löschen
Private Sub Pat_loeschen_Click()
 Dim Pat_id&, erg&
 Pat_id = InputBox("Welchen Patienten wollen Sie löschen?")
 Dim rsPat As New ADODB.Recordset
 myFrag rsPat, "SELECT gesname(" & Pat_id & ")"
 If Not rsPat.BOF() Then
  erg = MsgBox("Wollen Sie wirklich den Patienten `" & Pat_id & " (" & rsPat.Fields(0) & ")` löschen?", vbYesNo)
  If erg = vbYes Then
   Call LöschePat(Pat_id, True)
  End If
 End If ' Not rsPat.BOF() Then
End Sub ' Pat_loeschen_Click

' ...für Arzt -> Gefaxte Briefe anzeigen
Private Sub GefaxteBriefeAnzeigen_Click()
 Call ProgStart
 Aktion = GefaxteAnzeigen
 Set pataw.hlese = Me
 pataw.Show
End Sub ' GefaxteBriefeAnzeigen_Click

' ...für Arzt -> Patientenlaufzettel für alle Pat.mit best. Medikament
Private Sub PLZfuerMedikament_Click()
 Dim Med$ ', plzverzalt$
 Dim plzVerz$
 Dim rs As New ADODB.Recordset
 Med = InputBox("Bitte Medikamentenname eingeben", "Medikamentenname", "Forxiga")
 If Med <> "" Then
'  plzverzalt = plzVerz
  plzVerz = plzVz & Med & "\"
  MkDir plzVerz
  myFrag rs, "SELECT DISTINCT pat_id FROM medplan WHERE medikament LIKE '" & Med & "%' ORDER BY pat_id DESC"
  If Not rs.BOF Then
   Do While Not rs.EOF
    dodoplz rs!Pat_id, plzVz
    rs.MoveNext
   Loop
  End If
'  plzVerz = plzverzalt
 End If
 Ausgeb "Fertig mit PLZ für " & Med, 1
End Sub ' PLZfuerMedikament_Click

' ...für Arzt -> Verdächtige Überweiser
Private Sub VerdächtigeÜberweiser_Click()
 Call ProgStart
 Call doVerdächtigeÜberweiser
End Sub ' VerdächtigeÜberweiser_Click

' ...für Arzt -> Doppelte Diagnosen ermitteln
Private Sub DoppelteDiagnosen_Click()
 Dim rs As New ADODB.Recordset, datnam$, i&
 datnam = pVerz & "DoppelteDiagnosen " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".txt"
 Open datnam For Output As #327
 Print #327, "Nr. Zahl Pat_id Name                        ICD Sicherheit -Seite -text"
 Call ProgStart
 myFrag rs, "SELECT * FROM (SELECT COUNT(0) ct, pat_id, icd, gesnameg(pat_id) GesName, diagsicherheit si, diagtext tx, diagseite se FROM `diagnosen` WHERE obdauer <> 0 GROUP BY pat_id, icd, diagsicherheit, diagseite) AS i WHERE ct > 1 ORDER BY pat_id DESC"
 i = 0
 Do While Not rs.EOF
  i = i + 1
  Print #327, Right$("    " & i, 4) & " " & Left$(rs!ct & "  ", 2) & " " & Right$("     " & rs!Pat_id, 5) & " " & Left$(rs!gesname & Space$(30), 30) & " " & Left$(rs!ICD & "      ", 6) & " " & rs!SI & " " & rs!SE & " " & Left$(rs!tx & Space$(50), 50)
  rs.Move 1
 Loop
 Print #327, "Fertig!"
' Debug.Print "Fertig!"
 Close #327
 Call ProgEnde
 zeigan datnam
End Sub ' DoppelteDiagnosen_Click

' ...für Arzt -> KassenEditieren (Rabattverträge etc.)
Private Sub KassenEditieren_Click()
 Dim ked As New KassenEditieren, i&, rs As New ADODB.Recordset, DBName$
 ProgStart
 DBName = DefDB(DBCn)
 myFrag rs, "SELECT column_name cn, data_type dt, column_type ct, column_comment cc FROM information_schema.`COLUMNS` C WHERE table_schema = '" & DBName & "' AND table_name = 'kassenliste' AND column_type LIKE 'tinyint(1)%'", , DBCn
 Do While Not rs.EOF
  Debug.Print rs!Cn, rs!DT, rs!ct, rs!CC
  ked.Check1(ked.Check1.COUNT - 1).Caption = "&" & rs!Cn & " (" & rs!CC & ")"
  ked.Check1(ked.Check1.COUNT - 1).Tag = rs!Cn
  Load ked.Check1(ked.Check1.COUNT)
  ked.Check1(ked.Check1.COUNT - 1).Top = ked.Check1(ked.Check1.COUNT - 2).Top + 300
  ked.Check1(ked.Check1.COUNT - 1).Visible = True
  rs.Move 1
 Loop ' While Not rs.EOF
 Unload ked.Check1(ked.Check1.COUNT - 1)
 ked.Show
End Sub ' KassenEditieren_Click

' ...für Arzt -> Medarten editieren
Private Sub MedartenEditieren_Click()
 Dim mda As New Medarten
 mda.Show
End Sub ' MedartenEditieren_Click()

' ...für Arzt -> Laborparameter zusammenfassen
Private Sub LaborparameterZusammenfassen_Click()
 Call ProgStart
 Set pal = Nothing
 pal.PLArt = artLPar
 Set pal.hlese = Me
 Me.Hide
 pal.Show
 Call ProgEnde
End Sub ' LaborparameterZusammenfassen_Click

' Statistik

' Statistik -> Covid-Impfliste
Private Sub Covid_Impfliste_Click()
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT Inhalt, concat(nachNAME,', ',vorname) NAME" & vbCrLf & _
  ",LEFT(CONCAT(COALESCE((SELECT titel FROM desktop WHERE pat_id=i.pat_id AND iconPath RLIKE 'Raum5'),''),SPACE(15)),15) Notiz," & vbCrLf & _
  "CONCAT('->',Privattel) Tel" & vbCrLf & _
  ",IF(privatmobil='',concat('->',Privattel_2),concat('->',Privatmobil,IF(privattel_2='','',CONCAT(',',privattel_2)))) `Tel2/Mobil`" & vbCrLf & _
  "FROM (SELECT LEFT(inhalt,instr(inhalt,' ')) pat_id, inhalt" & vbCrLf & _
  "       FROM eintraege WHERE pat_id=(SELECT pat_id FROM namen WHERE Nachname='Covid')) i " & vbCrLf & _
  "LEFT JOIN namen USING (pat_id)" & vbCrLf & _
  "WHERE NOT ISNULL(lfdnr)" & vbCrLf & _
  "ORDER BY gebdat;"
 myFrag rs, sql
 TabAusgeb rs, Me, True, , , , , , "Covid19-Liste.csv", , , , "Covid-Impfliste, mit 'Raum5' für Notizen", , 0
End Sub ' Covid_Impfliste_Click

' Statistik -> Hausärzte von Pat. in der BKK
Private Sub HausärzteBKK_Click()
 Call ProgStart
 Call doHABKK(Me)
End Sub ' Sub HausärzteBKK_Click

' Statistik -> Überweiserstatistik
Private Sub Überweiserstatistik_Click()
 Dim rs As New ADODB.Recordset, datnam$, i&, ausg$, sql$
 datnam = pVerz & "Überweiserstatistik " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".csv"
 Open datnam For Output As #326
 Call ProgStart
' myFrag rs, "SELECT kvnu,anrede, haname,plz,ort,tel1,tel2,fax1,fax2,zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,gelöscht,ct FROM (SELECT COUNT(0) AS ct, LEFT(übwv,7) AS kvnu FROM `faelle` WHERE bhfb > " & DatFor_k(Now - 365) & " AND übwv <> '' GROUP BY übwv " & _
   "UNION SELECT COUNT(0) AS ct, LEFT(andüw,7) AS kvnu FROM `faelle` WHERE bhfb > " & DatFor_k(Now - 365) & " AND andüw <> '' GROUP BY andüw) AS i LEFT JOIN `kvaerzte`.`hae` USING (kvnu) WHERE not gelöscht AND NOT ISNULL(kvnu) AND kvnu <> '" & kvnr & "' ORDER BY ct DESC"
' myFrag rs, "SELECT kvnu,anrede, haname,plz,ort,tel1,tel2,fax1,fax2,zulg,arzttyp,dmpt2,dmpt1 FROM (SELECT COUNT(0) AS ct, LEFT(übwv,7) AS kvnu FROM `faelle` WHERE bhfb > " & DatFor_k(Now - 365) & " AND übwv <> '' GROUP BY übwv " & _
   "UNION SELECT COUNT(0) AS ct, LEFT(andüw,7) AS kvnu FROM `faelle` WHERE bhfb > " & DatFor_k(Now - 365) & " AND andüw <> '' GROUP BY andüw) AS i LEFT JOIN `kvaerzte`.`hae` USING (kvnu) WHERE not gelöscht AND NOT ISNULL(kvnu) AND kvnu <> '" & kvnr & "' ORDER BY ct DESC"
 sql = "SELECT ct,haname,ort,dmpt2,dmpt1,i.kvnu,lname,pat_id,bhfb FROM (SELECT COUNT(0) AS ct, übwr kvnu, GROUP_CONCAT(DISTINCT CAST(pat_id AS char)) pat_id, bhfb FROM quelle.faelle f WHERE bhfb > '2008-12-05 21:39:20' AND übwr <> '' AND übwr <> '" & BSNR & "' GROUP BY kvnu) i LEFT JOIN " & HADBName & ".`hae` hae ON i.kvnu = hae.kvnu LEFT JOIN (SELECT GROUP_CONCAT(DISTINCT name) lname,kvnr FROM `aktlue` l WHERE kvnro<>'' GROUP BY kvnr) l ON i.kvnu = l.kvnr ORDER BY ct DESC;"
 myFrag rs, sql
 For i = 0 To rs.Fields.COUNT - 1
  ausg = ausg & """" & rs.Fields(i).name & """;"
 Next
 ausg = Left$(ausg, Len(ausg) - 1)
 Print #326, ausg
 Do While Not rs.EOF
  ausg = vNS
  For i = 0 To rs.Fields.COUNT - 1
   ausg = ausg & """" & rs.Fields(i).Value & """;"
  Next
  ausg = Left$(ausg, Len(ausg) - 1)
  Print #326, ausg
  rs.Move 1
 Loop
 MsgBox "Fertig!"
 Close #326
 Call ProgEnde
 zeigan datnam
End Sub ' Überweiserstatistik_Click
' SELECT kvnu,anrede, haname,plz,ort,tel1,tel2,fax1,fax2,zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,gelöscht,ct FROM (SELECT COUNT(0) AS ct, LEFT(übwv,7) AS kvnu FROM `faelle` WHERE bhfb > '2007-09-30' AND übwv <> '' GROUP BY übwv UNION SELECT COUNT(0) AS ct, LEFT(andüw,7) AS kvnu FROM `faelle` WHERE bhfb > '2007-09-30' AND andüw <> '' GROUP BY andüw) AS i LEFT JOIN `kvaerzte`.`hae` USING (kvnu) WHERE not gelöscht AND NOT ISNULL(kvnu) AND kvnu <> '" & kvnr & "' ORDER BY ct DESC;

' Statistik -> &Überweiserstatistik d.letzten 2a
Private Sub Überweiserstatistik2_Click()
 Dim rs As New ADODB.Recordset
 sql = "SELECT COUNT(0) `Überw.Pat.`, übwvlanr, Titel, Vorname, Name,Fachgruppe,Straße,PLZ,Ort, GROUP_CONCAT(DISTINCT pid) PIDs FROM (" & _
          "SELECT übwvlanr, IF(ISNULL(l.titelt),f.üwtit,l.titelt) Titel," & _
          "if(ISNULL(l.vorname),f.üwvor,l.vorname) Vorname, IF(ISNULL(l.name),f.üwnan,l.name) Name," & _
          "if(ISNULL(l.name),'',l.fachgruppe) Fachgruppe, IF(ISNULL(l.name),'',l.strasse) Straße," & _
          "if(ISNULL(l.name),'',l.plz) PLZ, IF(ISNULL(l.name),'',l.ort) Ort," & _
          "GROUP_CONCAT(DISTINCT pat_id) pid, l.id lid " & _
          "FROM faelle f " & _
          "LEFT JOIN aktlue l ON f.üwnan = l.nameo AND f.üwvor = l.vno AND f.üwtit = l.titelto " & _
       "WHERE BhFB > SUBDATE(NOW(), 2 * 365) " & _
       "GROUP BY pat_id, übwvlanr) i " & _
       "WHERE übwvlanr<>'' " & _
       "GROUP BY lid ORDER BY COUNT(0) DESC;"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , "Überweiserstatistik 2a bis " & Format(Now(), "dd.mm.YYYY")
End Sub ' Überweiserstatistik2_Click

' SELECT COUNT(0),leistung FROM `leistungen` WHERE leistung IN ('97261B','92261A',
'92292C','92292D','92292E','97268','97269','97271','97274','97278','97279','92292A','92292B','92290A','92291A') AND YEAR(zeitpunkt)= '2008' GROUP BY leistung;
'1.7.12:
' schul:
' SELECT YEAR(zeitpunkt) Jahr, (month(zeitpunkt) -1) div 3 +1 Quartal, COUNT(0) Zahl FROM eintraege e WHERE art = 'schul' GROUP BY YEAR(zeitpunkt), (month(zeitpunkt) -1) div 3 +1;

' Statistik -> Schulungsstatistik nach Schulungsart
Private Sub Schulungsstatistik_Click()
 Dim col As New Collection, el, rs As New ADODB.Recordset, ausg$, TA1$, spmax%(5), rAf&
 myEFrag "INSERT INTO `ebm2000plus`(leistung,titel,euro) SELECT g.leistung, g.erklärung, g.wert FROM `genehmigungen` g LEFT JOIN `ebm2000plus` e ON g.leistung=e.leistung WHERE ISNULL(e.leistung)", rAf
 myFrag rs, "SELECT leistung FROM `genehmigungen` WHERE obschulung<>0"
 Do While Not rs.EOF
  ausg = rs!Leistung
  col.Add ausg
  rs.MoveNext
 Loop
 Set rs = Nothing
' col.Add "97261B"
' col.Add "92261A"
' col.Add "97262B"
' col.Add "92262A"
' col.Add "97263B"
' col.Add "92263A"
' col.Add "97269B"
' col.Add "92269A"
' col.Add "97264B"
' col.Add "92264A"
' col.Add "97265B"
' col.Add "92292C"
' col.Add "92265A"
' col.Add "97266B"
' col.Add "92292D"
' col.Add "92266A"
' col.Add "97268B"
' col.Add "92292E"
' col.Add "92268A"
' col.Add "97268"
' col.Add "92292A"
' col.Add "97269"
' col.Add "97271"
' col.Add "97274"
' col.Add "92292B"
' col.Add "97278"
' col.Add "92290A"
' col.Add "97279"
' col.Add "92291A"
 Call ProgStart
 
#If True Then
  Dim sql As New CString, lst As New CString
  For Each el In col:  lst.AppVar Array("'", el, "',"): Next el
  lst.Cut (lst.length - 1)
  sql.AppVar Array("SELECT e.Leistung,Titel,COUNT(pat_id) Zahl,CAST(GROUP_CONCAT(pat_id) AS char) Pat_IDs FROM `ebm2000plus` e LEFT JOIN `leistungen` l ON l.leistung = e.leistung AND YEAR(SUBDATE(NOW(),INTERVAL 15 DAY)) = YEAR(l.zeitpunkt) WHERE e.leistung IN (", lst.Value, ") GROUP BY e.leistung")
  myFrag rs, sql.Value
  spmax(3) = 100
  TA1 = TabAusgeb(rs, Me, , , , , spmax, , "Schulungsstatistik")
#Else ' True
'  Dim DatNam$
'  DatNam = pVerz & "Schulungsstatistik " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".csv"
'  Open DatNam For Output AS #325
'  Print #325, "Leistung Titel                                               Anzahl IN " & YEAR(NOW() - 15)
'  For Each el In col
'   SET rs = Nothing
'   myFrag rs, "SELECT leistung,titel,(SELECT COUNT(0) FROM `leistungen` WHERE leistung = e.leistung AND YEAR(zeitpunkt) = YEAR(SUBDATE(NOW(),INTERVAL 15 DAY))) ct FROM `ebm2000plus` e WHERE leistung = '" & el & "'"
'   IF rs.EOF THEN rs.Close: myFrag rs, "SELECT COUNT(0) ct, titel FROM `leistungen` l LEFT JOIN `ebm2000plus` e ON l.leistung = e.leistung WHERE l.leistung = '" & el & "' AND YEAR(zeitpunkt) = YEAR(SUBDATE(NOW(),INTERVAL 15 DAY))"
''   IF el = "97278" THEN
''    MsgBox "Stop in Schulungsstatistik_Click:" & vbCrLf & "el = '97278'"
''    Stop
''   END IF
'   ausg = Left$(el & Space(7), 7) & "(" & LEFT(rs!Titel & Space$(60), 60) & "): " & rs!ct
'   Print #325, ausg
'  Next el
'  MsgBox "Fertig!"
'  Close #325
' zeigan datnam
#End If ' True
End Sub ' Schulungsstatistik_Click

' Statistik -> Gruppenschulungsstatisik nach Ziffernzahl pro Quartal
Private Sub GruppenSchulungsstatisiknachZiffernzahlproQuartal_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT YEAR(zeitpunkt) Jahr, (month(zeitpunkt) -1) div 3 +1 Quartal, COUNT(0) Zahl FROM leistungen l LEFT JOIN `genehmigungen` g ON l.leistung = g.leistung WHERE  obschulung<>0  AND erklärung NOT LIKE '%Sach%' AND erklärung NOT LIKE '%material%' AND erklärung NOT LIKE '%schwang%' AND erklärung NOT LIKE '%buch%' AND erklärung NOT LIKE '%gestat%' AND g.leistung NOT IN ('92278','92282') GROUP BY YEAR(zeitpunkt), (month(zeitpunkt) -1) div 3 +1"
 TabAusgeb rs, Me, , , , , , , "Gruppenschulungsstatisik nach Ziffernzahl pro Quartal"
End Sub  ' GruppenSchulungsstatisiknachZiffernzahlproQuartal_Click

' Statistik -> Sch&ulungszifferanalyse
Private Sub Schlungsziffer_analyse_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT g.Leistung, YEAR(zeitpunkt) Jahr, (month(zeitpunkt) -1) div 3 +1 Quartal, COUNT(0) Zahl, g.Erklärung, g.Wert FROM leistungen l LEFT JOIN `genehmigungen` g ON l.leistung = g.leistung WHERE  obschulung<>0 GROUP BY g.Leistung, YEAR(zeitpunkt), (month(zeitpunkt) -1) div 3 +1"
 TabAusgeb rs, Me, , , , , , , "Schulungsziffer-Analyse"
End Sub ' Schlungsziffer-Analyse_Click

' Statistik -> Schulungsziffereinzelnachweis
Private Sub Schulungsziffereinzelnachweis_Click()
 Dim Ziffer$, rs As New ADODB.Recordset, spmaxü
 spmaxü = Array(10, 5, 200)
 Ziffer = InputBox("Für welche Ziffer (z.b. '97268', '97274'?")
' ktag fehlerhaft
 myFrag rs, "SELECT DATE(zeitpunkt) Datum, COUNT(0) Zahl, GROUP_CONCAT(pat_id) Pat_id FROM leistungen WHERE leistung = '" & Ziffer & "' GROUP BY tag ORDER BY tag"
 TabAusgeb rs, Me, , , , , spmaxü, , "Ziffereinzelnachweis für Ziffer " & Ziffer
End Sub ' Schulungsziffereinzelnachweis_Click

' Statistik -> Motivationsgesprächsstatistik
Private Sub Motivationsgesprächsstatistik_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT YEAR(zeitpunkt) Jahr, quarter(zeitpunkt) Quartal, COUNT(0) Zahl FROM leistungen l WHERE leistung IN ('92282','92278') GROUP BY YEAR(zeitpunkt),quarter(zeitpunkt) ORDER BY YEAR(zeitpunkt) DESC,quarter(zeitpunkt) DESC"
 TabAusgeb rs, Me, , , , , , , "Motivationsgesprächsstatistik"
End Sub ' Motivationsgesprächsstatistik_Click

' Statistik -> Patienten mit AOK-Kriterien
Private Sub PatientenMitAOKKriterien_Click()
' 2 x gestrichen: AND ln.einheit = '%'
 Dim rs As New ADODB.Recordset
' sql = "SELECT IF (HbA1c>=7.3 OR GFR <= 40,'X','') Krit, i.* from" & _
       "(SELECT `aktfvs`.pat_id, DATE(hba1c.zp) `HbA1c-Tag`, hba1c.wert HbA1c, DATE(krea.zp) `Krea-Tag`, krea.wert Krea, IF(REPLACE(krea.wert,',','.')=0 OR timestampdiff(year,n.gebdat,krea.zp)<=0,'-',round(186 * pow(REPLACE(krea.wert,',','.') * 1.0526,-1.154) * pow(timestampdiff(year,n.gebdat,krea.zp),-0.203)*if(geschlecht='w',0.742,1),0)) GFR  " & _
       "FROM `aktfvs` LEFT JOIN " & _
       "(SELECT pat_id, zeitpunkt zp, abkü, einheit, wert FROM `laborneu` WHERE abkü RLIKE '[ck]rea$|[ck]rea[^u]' AND einheit = 'mg/dl' UNION SELECT u.pat_id pat_id, u.eingang zp, w.abkü abkü, w.einheit einheit, w.wert wert FROM `laboryus` u LEFT JOIN laborywert w ON u.id = w.usid WHERE abkü RLIKE '[ck]rea$|[ck]rea[^u]' AND einheit = 'mg/dl' ORDER BY pat_id,zp DESC) krea ON `aktfvs`.pat_id = krea.pat_id LEFT JOIN `namen` n ON krea.pat_id = n.pat_id " & _
       "LEFT JOIN (SELECT pat_id, zeitpunkt zp, wert FROM `laborneu` ln WHERE abkü RLIKE 'hba[c1]' AND CAST(wert AS decimal) < 22 UNION SELECT u.pat_id, u.eingang zp, w.wert FROM `laboryus` u LEFT JOIN laborywert w ON u.id = w.usid WHERE abkü RLIKE 'hba[c1]' AND CAST(wert AS decimal) < 22 ORDER BY pat_id,zp DESC) hba1c ON `aktfvs`.pat_id = hba1c.pat_id WHERE datediff(NOW(), hba1c.zp) <100 OR datediff(NOW(),krea.zp)<100 GROUP BY `aktfvs`.pat_id) i"
 sql = "SELECT IF (h.letzter>=7.3 OR _lGFR(f.pat_id)<= 40,'X','') Krit, f.pat_id,h.letzter HbA1c,h.lzp HbA1cZp,k.letzter Krea, k.lzp KreaZp, _lGFR(f.pat_id) eGFR FROM " & vbCrLf & _
       "aktfvs f LEFT JOIN lHbA1c h USING (pat_id) " & vbCrLf & _
       "LEFT JOIN lKrea k USING (pat_id); "

 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , "PatMitAokKrit"
End Sub ' PatientenMitAOKKriterien_Click

' Statistik -> Gestationsdiabetikerinnen pro Quartal
Private Sub GestationsdiabetikerinnenProQuartal_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT COUNT(0) Zahl, quartal(diagdatum) Quartal" & vbCrLf & _
            "FROM diagview d " & vbCrLf & _
            "WHERE (d.icd='O24.4' AND d.f6010=0 AND d.diagsicherheit in ('G',' ')) " & vbCrLf & _
            "GROUP BY quartal(diagdatum)" & vbCrLf & _
            "ORDER BY MID(quartal(diagdatum),2) DESC, quartal(diagdatum) DESC;"
 TabAusgeb rs, Me, , , , , , , "Gestationsdiabetikerinnen pro Quartal", 0, True
End Sub ' GestationsdiabetikerinnenProQuartal_Click

' Statistik -> Gestationsdiabetikerinnen
Private Sub Gestationsdiabetikerinnen_Click()
' SELECT f.pat_id,f.fid, LEFT(CONCAT(a.nachname,' ',a.vorname),20) AS name, DATE_FORMAT(a.gebdat,'%d.%m.%y') AS geb, d.icd, d.diagsicherheit AS dsi, diabetestyp FROM `aktfvs` f LEFT JOIN `diagnosen` d ON f.fid = d.fid AND icd = 'O24.4' AND diagsicherheit NOT IN ('A','Z') LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id WHERE (NOT ISNULL(icd) OR a.diabetestyp = 'g');
 Dim rs As New ADODB.Recordset, rez As New ADODB.Recordset, Pat_id&, AusgStr$, TA1$, STA1$(), i&
 Dim SpMin%(2)
 SpMin%(0) = 6
 Call ProgStart
' Open DatNam For Output AS #327
 myFrag rs, "SELECT f.pat_id,f.fid, LEFT(CONCAT(a.nachname,' ',a.vorname),20) name, DATE_FORMAT(a.gebdat,'%d.%m.%y') AS geb, d.icd, d.diagsicherheit AS dsi, diabetestyp " & vbCrLf & _
            "FROM `aktfvs` f LEFT JOIN `diagview` d ON f.pat_id = d.pat_id AND (d.icd='O24.4' AND d.f6010=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))" & vbCrLf & _
            "LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id " & vbCrLf & _
            "WHERE (NOT ISNULL(icd) OR a.diabetestyp = 'g')"
 TabAusgeb rs, Me, True, , , , , , "Gestationsdiabetikerinnen"
End Sub ' Gestationsdiabetikerinnen_Click

' Statistik -> PLZ aus Liste (kein Insulin mehr)
Private Sub PLZausListe_Click()
' Dim plzverzalt$
 Dim sql$, Zahl&, plzVerz$
 Dim rs As New ADODB.Recordset
 On Error GoTo fehler
'  plzverzalt = plzVerz
  plzVerz = plzVz & "keinInsulinmehr\"
  On Error Resume Next
  MkDir plzVerz
  On Error GoTo fehler
'  sql = "SELECT t1.pat_id, gesnameg(t1.pat_id) Name, t1.therart, COUNT(gs.art) gsz, COUNT(tk.art) tkz, t1.zp t1zp,t2.therart,t2.zp t2zp FROM therarten t1 " & _
        "LEFT JOIN therarten t2 ON t1.pat_id = t2.pat_id AND t2.zp> t1.zp " & _
        "LEFT JOIN eintraege gs ON t1.pat_id = gs.pat_id AND gs.art = 'gs' " & _
        "LEFT JOIN eintraege tk ON t1.pat_id = tk.pat_id AND tk.art = 'tk' " & _
        "WHERE NOT ISNULL(t2.zp) AND t1.therart IN ('ict','ct','Komb') AND t2.therart IN ('diät','oad') " & _
        "AND NOT EXISTS (SELECT icd FROM diagnosen d WHERE d.pat_id = t1.pat_id AND d.icd LIKE 'E10%' AND d.diagsicherheit IN ('G',' ')) " & _
        "GROUP BY pat_id ORDER BY t2.zp DESC; "
' liefert auf SQL-Ebene z.T. falsche Ergebnisse für gsz und tkz (jeweils die gleiche hohe Zahl), liegt wohl am Group (müßte noch tk.id und gs.id enthalten
  sql = "SELECT t1.pat_id, gesnameg(t1.pat_id) Name, t1.therart, " & _
          "(SELECT COUNT(0) FROM eintraege WHERE pat_id = t1.pat_id AND art = 'gs') gsz, " & _
          "(SELECT COUNT(0) FROM eintraege WHERE pat_id = t1.pat_id AND art = 'tk') tkz, " & _
          "t1.zp t1zp,t2.therart,t2.zp t2zp FROM therarten t1 " & _
        "LEFT JOIN therarten t2 ON t1.pat_id = t2.pat_id AND t2.zp> t1.zp " & _
        "WHERE NOT ISNULL(t2.zp) AND t1.therart IN ('ict','ct','Komb') AND t2.therart IN ('diät','oad') " & _
        "AND NOT EXISTS (SELECT icd FROM diagview d WHERE d.pat_id = t1.pat_id AND d.gICD LIKE 'E10%' " & _
        "GROUP BY pat_id ORDER BY t2.zp DESC, t2.MPNr desc; "
  myFrag rs, sql
  If Not rs.BOF Then
   Do While Not rs.EOF
    Zahl = Zahl + 1
    dodoplz rs!Pat_id, plzVz, , , , rs!T1Zp & " - " & rs!T2Zp & " (gs " & rs!gsz & ", tk " & rs!Tkz & ")", 0
    rs.MoveNext
   Loop
  End If
'  plzVerz = plzverzalt
 Ausgeb "Fertig mit PLZ für " & Zahl & " Patienten mit keinem Insulin mehr", 1
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PLZausListe_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' PLZausListe_Click

' Statistik -> Pioglitazon-Rezepte ab 1.4.11
Private Sub PioglitazonRezepte_Click() ' Pioglitazon-Rezepte ab 1.4.11
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT r.pat_id, r.zeitpunkt, r.medikament, f.schgr FROM `rezepteintraege` r LEFT JOIN `faelle` f ON r.fid = f.fid WHERE (medikament LIKE '%actos%' OR medikament LIKE '%competact%') AND rezklkurz<>'prp' AND zeitpunkt > '2011-03-31' AND schgr <> 90"
 TabAusgeb rs, Me, , , , , , , "Pioglitazon-Rezepte ab 1.4.11"
End Sub ' PioglitazonRezepte_Click

' Statistik -> Diabetiker ohne Schulung letztes Jahr
Private Sub DiabetikerOhneSchulungLetztesJahr_Click()
 Dim rs As New ADODB.Recordset
 sql = "SELECT f.Pat_ID, gesname(n.pat_id) PName FROM `aktfvs` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN diagview d ON f.pat_id = d.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.f6010=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)))" & vbCrLf & _
       "LEFT JOIN `eintraege` e ON f.pat_id = e.pat_id AND art = 'schul' AND YEAR(zeitpunkt) = YEAR(SUBDATE(NOW(),INTERVAL 25 DAY)) " & vbCrLf & _
       "WHERE ISNULL(art) AND NOT ISNULL(icd) AND schgr <> 90 GROUP BY pat_id"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , "DiabetikerOhneSchulungLetztesJahr"
End Sub ' DiabetikerOhneSchulungLetztesJahr_Click

' Statistik -> Therapieartenwechsel
Private Sub Therapieartenwechsel_Click() ' s. therart_erm
 Dim rs As New ADODB.Recordset, sql$, altpat_id&, altTherArt$, rAf&, erg&, T1!, T2!
 Const FristS$ = "25"
 Call ProgStart
 erg = MsgBox("Mit Neuauswertung der Therapiearten?", vbYesNo + vbQuestion + vbDefaultButton2, "Rückfrage")
 If erg = vbYes Then
  T1 = Timer
  Ausgeb "Bitte warten", 0
'  myEFrag "CREATE TABLE IF NOT EXISTS `therarten`(id integer key auto_increment,pat_id integer, zp datetime, mpnr integer, therart varchar(7), index pat_id(pat_id))", rAF
  myEFrag "CREATE TABLE IF NOT EXISTS `therarten`(id INT(11) NOT NULL AUTO_INCREMENT,pat_id INT(11) NULL DEFAULT NULL,zp DATETIME NULL DEFAULT NULL,mpnr INT(11) NULL DEFAULT NULL,therart VARCHAR(7) NULL DEFAULT NULL COLLATE 'utf8mb4_german2_ci',insart INT(1) UNSIGNED NOT NULL DEFAULT '0' COMMENT '\'0=keines, 1=nur Mahlzeiten,2=nur Verzögerungs,3=nur Misch, 4=verschiedene',Grund VARCHAR(1000) NULL DEFAULT NULL COMMENT 'Grund/Gründe für Zuordnung' COLLATE 'utf8mb4_german2_ci',absPos INT(10) NULL DEFAULT NULL,AktZeit DATETIME NULL DEFAULT NULL,StByte INT(10) NULL DEFAULT NULL,PRIMARY KEY (id) USING BTREE,INDEX pat_id (pat_id) USING BTREE,INDEX zp (zp,mpnr) USING BTREE) COLLATE='utf8mb4_german2_ci' ENGINE = MyISAM"
  myEFrag "TRUNCATE `therarten`"
#If Not thaalt Then
' 22.10.22: führt bei Aufruf über Ado zumindest bis zur Mariadb-Version 10.9 immer wieder zum Server-Crash, s.ähnliche Bug-Hinweise früherer Versionen
#Const mitfensters = False
#If mitfensters Then
  rufauf "ssh", "root@" & LiName & " mysql --defaults-extra-file=~/.mysqlpwd quelle -e'CALL fuellThaP(0)'", 2, "c:\windows\system32\openssh\", -1, 0
#Else
 Call TheraErmitt(0)
#End If
'  myEFrag "CALL fuellThaP(0)"
#Else
  sql = "SELECT pat_id, zp, mpnr, IF(purez OR puzu,'CSII',IF(obict=1,'ICT',IF(insu=0,IF(oad=1,'OAD','Diät'),IF(oad=1,'Komb','CT')))) therart FROM (SELECT mpü.pat_id, mpü.zeitpunkt zp, mpü.mpnr, " & _
        "(SELECT COUNT(0)<>0 oad FROM `medplan` mpu LEFT JOIN `medarten` ma ON mpu.medanfang = ma.medikament WHERE mpu.pat_id = mpü.pat_id AND mpu.mpnr = mpü.mpnr AND (glib<>0 OR metf<>0 OR gluci<>0 OR shglin<>0 OR glit<>0 OR dpp4<>0 OR sglt2<>0 OR sonstad<>0)) oad, " & _
        "(SELECT COUNT(0)<>0 oad FROM `medplan` mpu LEFT JOIN `medarten` ma ON mpu.medanfang = ma.medikament WHERE mpu.pat_id = mpü.pat_id AND mpu.mpnr = mpü.mpnr AND (glp1<>0)) glp1, " & _
        "(SELECT (MAX((mpu.mo<>'')+(mpu.mi<>'')+(mpu.ab<>'')+(mpu.zn<>''))>0) obict FROM `medplan` mpu LEFT JOIN `medarten` ma ON mpu.medanfang = ma.medikament WHERE mpu.pat_id = mpü.pat_id AND mpu.mpnr = mpü.mpnr AND insart =1) obict, " & _
        "(SELECT COUNT(0)<>0 insu FROM `medplan` mpu LEFT JOIN `medarten` ma ON mpu.medanfang = ma.medikament WHERE mpu.pat_id = mpü.pat_id AND mpu.mpnr = mpü.mpnr AND (ins<>0 OR anal<>0)) insu, " & _
        "(SELECT COUNT(0)<>0 puzu FROM `medplan` mpu LEFT JOIN `medarten` ma ON mpu.medanfang = ma.medikament WHERE mpu.pat_id = mpü.pat_id AND mpu.mpnr = mpü.mpnr AND puzu<>0) purez, " & _
        "(SELECT COUNT(0)<> 0 FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid=`forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid=`forminhkopf`.form_id) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw  WHERE form_abk IN ('lar','plar') AND feld IN ('medikament','txtmedKey') AND zeitpunkt > SUBDATE(mpü.zeitpunkt, INTERVAL 0.5 YEAR) AND pat_id = mpü.pat_id) puzu " & _
        "(feldinh LIKE '%reservoir%' OR feldinh LIKE '%Rapid D Link%' OR feldinh LIKE '%Rap D Li%' OR feldinh LIKE '%Rapid-D Li%' OR feldinh LIKE '%TenderL%' OR feldinh LIKE '%FlexL%' OR feldinh LIKE '%Check Spirit%' OR feldinh LIKE '%Insight%' OR feldinh LIKE '%Chek Spirit%' OR feldinh LIKE '%Pumpenträg%' OR feldinh LIKE '%Kunststoffampu%' OR feldinh LIKE '%Spritzampull%' OR feldinh LIKE '%batteriefachdeckel%' OR feldinh LIKE '%H-Tron%' OR feldinh LIKE '%D-Tron%' OR feldinh LIKE '%Paradigm%' OR feldinh LIKE '%CSII%' OR feldinh LIKE '%linpumpe%' OR feldinh LIKE '%omnipod%' OR feldinh LIKE '%ypso pump%' OR " & _
        " feldinh LIKE '%MiniMed%' OR feldinh LIKE '%640G%' OR feldinh LIKE '%CareLink%' OR Feldinh LIKE '%Mio %' OR feldinh LIKE '%Quick%set%' OR feldinh LIKE '%Silhouette%' OR feldinh LIKE '%Sure-T%' OR feldinh LIKE '%Sure T%' OR feldinh LIKE '%Paradigm%' OR feldinh LIKE '% Veo%' OR feldinh LIKE '%Animas%' OR feldinh LIKE '%Car%idge%') AND NOT feldinh LIKE  AND NOT feldinh LIKE '%menveo%' AND zeitpunkt > SUBDATE(mpü.zeitpunkt, INTERVAL 0.5 YEAR) AND pat_id = mpü.pat_id) puzu " & _
        "FROM `medplan` mpü GROUP BY mpü.pat_id, mpü.mpnr, mpü.zeitpunkt) i"
  myFrag rs, sql
  Do While Not rs.EOF
   If rs!Pat_id <> altpat_id Or rs!therart <> altTherArt Then
    InsKorr DBCn, DBCnS, "INSERT INTO `therarten`(pat_id,zp,mpnr,therart) VALUES(" & rs!Pat_id & "," & DatFor_k(rs!Zp) & "," & rs!MPNr & ",'" & rs!therart & "')", rAf
    altpat_id = rs!Pat_id
    altTherArt = rs!therart
   End If
   rs.Move 1
  Loop
#End If
  T2 = Timer
  Ausgeb "Fertig mit Auffüllen der Tabelle Therapiearten, " & CStr(T2 - T1) & " Sekunden", True, True
 End If ' erg=vbYes
 Set rs = Nothing
' ' LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND icd REGEXP '^E1[0-4]' AND diagsicherheit <> 'A'
 sql = "SELECT a.Pat_ID, gesname(a.pat_id) Name, ICD, AufnDat, SchGr, t.zp `Beginn`,TherArt " & _
       "FROM `aktfvs` a " & vbCrLf & _
       "LEFT JOIN diagview d ON a.pat_id = d.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.f6010=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(a.quartal) AND qends(a.quartal)))" & _
       "LEFT JOIN `therarten` t ON a.pat_id = t.pat_id " & _
       "WHERE (therart ='ICT' OR therart = 'CSII') AND " & _
       "zp BETWEEN " & lQAnfuEnd(FristS) ', 1)
' "zp BETWEEN STR_TO_DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & fristS & " DAY)),'/',((month(SUBDATE(NOW(),INTERVAL " & fristS & " DAY))-1) div 3)*3+1,'/1'),'%Y/%m/%d') AND SUBDATE(adddate(STR_TO_DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & fristS & " DAY)),'/',((month(SUBDATE(NOW(),INTERVAL 20 DAY))-1) div 3)*3+1,'/1'),'%Y/%m/%d'),INTERVAL 3 MONTH),INTERVAL 1 DAY)"
 myFrag rs, sql
 Call TabAusgeb(rs, Me, , , , , , , "NeueICT")
End Sub ' Sub Therapieartenwechsel_Click

' Statistik -> HbA&1c-Statistik (dauert ...)
Private Sub HbA1cStatistik_Click() ' HbA1c-Stastisik
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT COUNT(0) Zahl, ROUND(avg(HbA1c),1) `mittl.HbA1c`, Quartal,ICD " & vbCrLf & _
 "FROM (" & vbCrLf & _
  "SELECT f.quartal, f.pat_id, MIN(d.icd) icd, l.letzter HbA1c FROM faelle f " & vbCrLf & _
  "LEFT JOIN diagview d ON f.pat_id = d.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.f6010=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal)))" & vbCrLf & _
  "LEFT JOIN lHbA1c l ON f.pat_id = l.pat_id " & vbCrLf & _
  "WHERE NOT ISNULL(icd) GROUP BY f.quartal, f.pat_id) i " & vbCrLf & _
 "GROUP BY quartal, icd " & vbCrLf & _
 "ORDER BY icd,MID(quartal,2),quartal"
 TabAusgeb rs, Me, , , , , , , "HbA1c-Stastik nach Diabetes-Typ und Quartal"
End Sub ' HbA1c-Statistik

' Statistik -> Hilfsmittelklassifikationen
Private Sub Hilfsmittelklassifikationen_Click()
 Call ProgStart
 Call doHilfsmittelklassifikationen(Me)
 Call ProgEnde
End Sub ' Hilfsmittelklassifikationen_Click

' Statistik -> Einlesungen
Private Sub Einlesungen_Click()
 Dim sql$, rs As New ADODB.Recordset
' sql = "SELECT COUNT(0) Zahl, n.Stbyte, MAX(pat_id) lPat, gesname(max(pat_id)) PName, COALESCE(e.Zp1,'') Zp1, COALESCE(e.zp4,'') Zp4, COALESCE(e.Fallzahl,'') Fallzahl, e.Datei, e.DateiAend FROM namen n LEFT JOIN `eintragszahlen` e ON n.stbyte = e.stbyte GROUP BY n.stbyte ORDER BY n.stbyte DESC"
' die beiden Folgenden sind anfangs gleichschnell, nach wiederholtem Aufruf ist das zweite schneller, ferner einfacher
' SELECT n.aktzeit, f.aktzeit impzeit, n.pat_id, gesname(n.pat_id) NAME, n.StByte from namen n LEFT JOIN faelle f ON n.pat_id = f.pat_id AND f.aktzeit=(SELECT MAX(aktzeit) FROM faelle WHERE pat_id=n.pat_id) GROUP BY n.pat_id ORDER BY stbyte DESC, impzeit DESC) i GROUP BY stbyte ORDER BY stbyte DESC;
' SELECT COUNT(0) zahl, i.* FROM (SELECT n.aktzeit, (select max(aktzeit) from faelle where pat_id=n.pat_id) impzeit, n.pat_id, gesname(n.pat_id) NAME, n.StByte from namen n GROUP BY n.pat_id ORDER BY stbyte DESC, impzeit DESC) i GROUP BY stbyte ORDER BY stbyte DESC;
 sql = _
 "SELECT COUNT(0) Zahl, i.Aktzeit, i.Impzeit Importzeit, Pat_id, PName" & vbCrLf & _
 ", COALESCE(e.Zp1,'') Zp1, COALESCE(e.zp4,'') Zp4, COALESCE(e.Fallzahl,'') Fallzahl, e.Datei, e.DateiAend" & vbCrLf & _
 "FROM (" & vbCrLf & _
 " SELECT COALESCE(aktzeit,'') Aktzeit, COALESCE((SELECT MAX(aktzeit) FROM faelle WHERE pat_id=n.pat_id),'') impzeit, IF(Pat_id=0,'',pat_id) pat_id, gesname(pat_id) PName, StByte" & vbCrLf & _
 " FROM namen n" & vbCrLf & _
 " GROUP BY pat_id ORDER BY stbyte DESC, impzeit DESC" & vbCrLf & _
 ") i" & vbCrLf & _
 "LEFT JOIN eintragszahlen e ON e.stbyte = i.stbyte" & vbCrLf & _
 "GROUP BY i.stbyte ORDER BY i.stbyte DESC;"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , , , , , "Einlesungen"
End Sub ' Einlesungen_Click

' Statistik -> Fallzahlstand
Private Sub Fallzahlstand_Click()
 Dim tage$
 Call ProgStart
 Do While Not IsNumeric(tage) And tage <> "-"
  tage = InputBox("Vor wie vielen Tagen ('-' = gesamtes Quartal)?", "Rückfrage", 0)
  If LenB(tage) = 0 Then Exit Sub
 Loop
 If tage <> "-" Then tage = CDbl(tage)
 Call dofallzahlstand(Me, tage)
 Call ProgEnde
End Sub ' Fallzahlstand_Click

' Statistik -> Wohnortstatistik
Private Sub Wohnortstatistik_Click()
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT * FROM (SELECT COUNT(0) Zahl, ROUND(COUNT(0)/(SELECT COUNT(0) FROM `aktfvs`)*100,2) Prozent, Ort, Plz FROM `aktfvs` f LEFT JOIN namen n USING (pat_id) GROUP BY plz) i ORDER BY prozent DESC"
 TabAusgeb rs, Me, , , , , , , "Wohnortstatistik"
End Sub ' Herkunftsstatistik_Click

' Statistik -> GNR-Statistiken einlesen
Private Sub GNR_Statistiken_einl_Click() ' GNR-Stastistiken einlesen, für GNR_Statistik und Quartalsvergleich
 Call doGNR_Statistiken_einl_Click
End Sub ' GNR_Statistiken_einl_Click

' Statistik -> GNR-Statistik
Private Sub GNR_Statistik_Click()
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT ROUND(SUM(uwert)) wert, COUNT(z.id) ZdLstZiff," & _
 "(SELECT COUNT(DISTINCT pat_id) FROM faelle WHERE quartal=concat(MID(qinv,5,1),LEFT(qinv,4)) AND schgr <> '90' AND NOT goäkatnr IN ('40','41') AND nachname <> 'Bereitschaftsdienst') FZahl," & _
 "s.Datei, s.DateiDat, s.Qinv FROM GNRStat s LEFT JOIN GNRZahl z ON s.id = z.statid GROUP BY s.id ORDER BY qinv DESC"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , , , , , "GNR-Statistik"
End Sub ' GNR_Statistik_Click

' Statistik -> Quartalsvergleich
Private Sub Quartalsvergleich_Click()
 Dim sql$, rs As New ADODB.Recordset
 Dim q1$, q2$, ID1$, id2$
 q1 = InputBox("Bitte das untersuchende Quartal eingeben (yyyym)", "Eingabe 1", QuartalStr(Now() - 21, True))
 q2 = InputBox("Bitte das Vergleichsquartal eingeben (yyyym)", "Eingabe 2", QuartalStr(Now() - 21 - 90, True))
 Set rs = Nothing
 myFrag rs, "SELECT id FROM GNRStat WHERE qinv = '" & q1 & "'"
 If rs.EOF Then
  MsgBox "kein Datensatz zu " & q1 & "gefunden!"
  Exit Sub
 Else
  ID1 = rs!id
 End If
 Set rs = Nothing
 myFrag rs, "SELECT id FROM GNRStat WHERE qinv = '" & q2 & "'"
 If rs.EOF Then
  MsgBox "kein Datensatz zu " & q2 & "gefunden!"
  Exit Sub
 Else
  id2 = rs!id
 End If
 sql = "SELECT i.*, l.leistungstext FROM (" & _
       "SELECT gnr, ROUND(wert1) _" & q1 & ",round(wert2) _" & q2 & ",round(wert1-wert2) Diff FROM (SELECT z1.gnr, z1.euro euro1, z1.uwert wert1, IF(ISNULL(z2.euro),0,z2.euro) euro2 , IF(ISNULL(z2.uwert),0,z2.uwert) wert2 FROM GNRZahl z1 LEFT JOIN GNRZahl z2 ON z1.gnr = z2.gnr AND z2.statid=" & id2 & " WHERE z1.statid=" & ID1 & ") i " & _
       "UNION DISTINCT " & _
       "SELECT gnr, ROUND(wert2) _" & q1 & ",round(wert1) _" & q2 & ",round(wert2-wert1) Diff FROM (SELECT z1.gnr, z1.euro euro1, z1.uwert wert1, IF(ISNULL(z2.euro),0,z2.euro) euro2 , IF(ISNULL(z2.uwert),0,z2.uwert) wert2 FROM GNRZahl z1 LEFT JOIN GNRZahl z2 ON z1.gnr = z2.gnr AND z2.statid=" & ID1 & " WHERE z1.statid=" & id2 & ") i " & _
       ") i LEFT JOIN EBM2010 l ON IF(LEFT(l.ziffer,1)='0',MID(l.ziffer,2),l.ziffer) = i.gnr AND leistungstext<>'' " & _
       "ORDER BY diff"
 Set rs = Nothing
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , "Quartalsvergleich " & q1 & " vs. " & q2, 1
End Sub ' Quartalsvergleich_Click

' Statistik -> Pumpenträgerliste
' künftig ähnlich:
' SELECT gesname(th.pat_id) Name, th.Pat_id, DATE_FORMAT(th.zp,'%d.%m.%Y') CSII_hier_seit, (SELECT DATE_FORMAT(MAX(bhfb),'%d.%m.%Y') FROM faelle f WHERE f.pat_id= th.pat_id) BhFB FROM therarten th LEFT JOIN anamnesebogen a ON th.pat_id = a.pat_id WHERE therart = 'CSII' AND zp = (SELECT MAX(zp) FROM therarten t WHERE t.pat_id = th.pat_id) AND tkz=0;
Private Sub Pumpenträgerliste_Click() ' s. therart_erm
 Dim rs As New ADODB.Recordset, rez As New ADODB.Recordset, Pat_id&, AusgStr$, datnam$, TA1$, STA1$(), i&
 Dim SpMin%(2)
 SpMin%(0) = 6
 Call ProgStart
 datnam = pVerz & "Pumpenträger " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".txt"
 Open datnam For Output As #326
 ' für Acrobat Querdruck 70%
 myFrag rs, "SELECT a.pat_id, LEFT(CONCAT(a.nachname,',',a.vorname,IF(a.titel='','',','),a.titel,IF(a.nvorsatz='','',' '),a.nvorsatz,' (',a.anrede,')'),24) AS name, DATE_FORMAT(a.gebdat,'%d.%m.%y') AS geb, LEFT(CONCAT(`diabetes seit`,' ',a.`insulin seit`),12) AS 'D.m./Ins.', LEFT(CONCAT(IF(a.insulinpumpe=1,'+','-'),' ',a.`insulinpumpe seit`,' ',a.`insulinpumpe marke`),24) AS 'Pumpe b.Anamn./seit/Marke', LEFT(a.ther1,4) AS Ther1, LEFT(a.dmp,9) AS DMP, f.schgr AS SG, DATE_FORMAT(f.bhfb,'%d.%m.%y') AS bhfb, LEFT(CONCAT(privattel, '|',privatmobil,'|',diensttel,'|',email,'|',privatfax,'|',privattel_2),60) AS kontakt FROM `anamnesebogen` a LEFT JOIN `namen` n ON a.pat_id = n.pat_id LEFT JOIN lfaellev f ON a.pat_id = f.pat_id WHERE therakt = 'CSII' AND tkz = 0"
 TA1 = TabAusgeb(rs, Me, True).Value
 Print #326, TA1
 STA1 = Split(TA1, vbCrLf)
 Print #326, ''
 Print #326, ''
 i = 1
 Do While Not rs.EOF
 ' pumpentr
  Set rez = Nothing
  Pat_id = rs!Pat_id
  ' ,'txtmedKey'
  sql = forminhalt & " WHERE form_abk IN ('lar','plar') AND feld IN ('medikament','txtmedKey') AND " & _
        "(feldinh LIKE '%reservoir% OR feldinh LIKE '%Rapid D Link%' OR feldinh LIKE '%Rap D Li%' OR feldinh LIKE '%Rapid-D Li%' OR feldinh LIKE '%TenderL%' OR feldinh LIKE '%FlexL%' OR feldinh LIKE '%Check Spirit%' OR feldinh LIKE '%Insight%' OR feldinh LIKE '%Chek Spirit%' OR feldinh LIKE '%Pumpenträg%' OR feldinh LIKE '%Kunststoffampu%' OR feldinh LIKE '%Spritzampull%' OR feldinh LIKE '%batteriefachdeckel%' OR feldinh LIKE '%H-Tron%' OR feldinh LIKE '%D-Tron%' OR feldinh LIKE '%Paradigm%' OR feldinh LIKE '%CSII%' OR feldinh LIKE '%linpumpe%' OR feldinh LIKE '%omnipod%' OR feldinh LIKE '%ypso pump%' OR feldinh LIKE '%MiniMed%' OR feldinh LIKE '%640G%' OR feldinh LIKE '%CareLink%' OR Feldinh LIKE '%Mio %' OR feldinh LIKE '%Quick%set%' OR feldinh LIKE '%Silhouette%' OR feldinh LIKE '%Sure-T%' OR feldinh LIKE '%Sure T%' OR feldinh LIKE '%Paradigm%' OR feldinh LIKE '% Veo%' OR feldinh LIKE '%Animas%' OR feldinh LIKE '%Car%idge%') AND NOT feldinh LIKE  AND NOT feldinh LIKE '%menveo%'" & _
        "AND zeitpunkt > " & DatFor_k(rs!BhFB - 640) & " AND pat_id = " & Pat_id & " ORDER BY zeitpunkt DESC LIMIT 10"
  myFrag rez, "SELECT Pat_ID, Zeitpunkt, Feldinh FROM (" & sql & ") i"
  Print #326, STA1(i)
  Print #326, TabAusgeb(rEinl:=rez, AusgebFrm:=Me, obMitausgeb:=False, ohneKopfZ:=True, SpMinÜ:=SpMin).Value
'  Do While Not rez.EOF
'   rez.Move 1
'  Loop
  rs.Move 1
  i = i + 1
 Loop
 Close #326
 zeigan datnam
End Sub ' Pumpenträgerliste_Click

' Statistik -> suchTel
Private Sub suchTel_Click()
 ProgStart
 Call doSuchTel(Me)
End Sub ' suchTel_Click

' EDV

' EDV -> MachTypen (Datei Typen.bas erstellen)
Private Sub MachTypen_Click(Index%)
 Call ProgStart
 Call MacheTypen(Me)
 Call ProgEnde
End Sub ' MachTypen_Click

' EDV -> MachDB
Private Sub MachDB_Click()
 Dim MdB As New MachDatenbank
 MdB.Show
 Set MdB = Nothing
End Sub ' MachDB_Click

' EDV -> MachDB
Private Sub AlleFallzahlstände_Click()
 Dim i%, erg&
 erg = MsgBox("Sollen alle Fallzahlstände berechnet werden?", vbYesNo Or vbDefaultButton2, "Rückfrage")
 If erg = vbYes Then
  For i = 93 To 0 Step -1
   dofallzahlstand Me, CStr(i)
  Next i
 End If ' erg = vbYes THEN
 dofallzahlstand Me, "-"
End Sub ' AlleFallzahlstände_Click

' EDV -> Leistungen zu Patient anzeigen
Private Sub Leistungen_zu_Patient_anzeigen_Click()
 Dim rs As New ADODB.Recordset, rsa As New ADODB.Recordset, spmaxü
 Dim pid$
 spmaxü = Array(10, 5, 200)
 pid = InputBox("Bitte PAT_ID eingeben")
 If pid <> 0 Then
  myFrag rs, "SELECT l.QS, l.AktZeit, l.Zeitpunkt,l.Leistung, IF (ISNULL(e2.titel), e.Leistungstext,e2.titel) Titel,f5002, f5005, f5006, f5009, Med, f5015, f5016, f5021, f5026, Faktor, f5098, LANR, letzVorg, Ausn, Beme, absPos, QT, StByte, LANRid, Sachkbez, Sachkct, Zone, l.FID, l.id FROM leistungen l LEFT JOIN ebm2000plus e2 USING (leistung) LEFT JOIN EBM2010 e ON l.leistung = e.ziffer WHERE pat_id=" & CStr(pid) & " ORDER BY zeitpunkt DESC"
  myFrag rsa, "SELECT * FROM namen WHERE pat_id=" & pid
  TabAusgeb rs, Me, , , , , spmaxü, , "Leistungen zu Pat. " & CStr(pid) & " (" & GesNamFn(rsa) & ")           "
 End If
End Sub ' Leistungen_zu_Patient_anzeigen_Click

' EDV -> Therapiearten anzeigen
Private Sub Therapiearten_anzeigen_Click()
 Dim rs As New ADODB.Recordset, spmaxü
 Dim pid$
 spmaxü = Array(10, 5, 200)
 pid = InputBox("Bitte PAT_ID eingeben")
 If pid <> 0 Then
  myFrag rs, "SELECT * FROM therarten WHERE pat_id=" & CStr(pid) & " ORDER BY zp DESC, mpnr DESC"
  TabAusgeb rs, Me, , , , , spmaxü, , "Therapiearten von Pat. " & CStr(pid) '& " (" & gesname(PID) & ")"
 End If
End Sub ' Therapiearten_anzeigen_Click

' EDV -> Dokumente in Datenbank
Private Sub DokumenteInDatenbank_Click()
 Const TMDok$ = "tmdok"
 Dim idt As TMIniDatei, DPfad$
 Call ProgStart
 Set idt = New TMIniDatei
 DPfad = idt.GetProp("Verzeichnisse/TurboMed/Dokumente", "Pfad")
 myEFrag "DROP DATABASE `" & TMDok & "`"
 myEFrag "CREATE DATABASE IF NOT EXISTS `" & TMDok & "` default CHARACTER SET 'utf8mb4' COLLATE 'utf8mb4_german2_ci'"
 myEFrag "USE `" & TMDok & "`"
 myEFrag "CREATE TABLE IF NOT EXISTS `tmdok`.`Dokumente` (  `id` int(11) NOT NULL AUTO_INCREMENT,  `Pfad` varchar(255),  `Datei` varchar(255),  `Größe` int(11) DEFAULT NULL,  `geändert` datetime DEFAULT NULL,  PRIMARY KEY (`id`),  KEY `Name` (`Pfad`,`Datei`),  KEY `Größe` (`Größe`),  KEY `geändert` (`geändert`)) ENGINE=InnoDB DEFAULT CHARSET='utf8mb4' COLLATE=utf8mb4_german2_ci ROW_FORMAT=DYNAMIC"
 Call dverz(DPfad)
End Sub ' DokumenteInDatenbank_Click

' EDV -> Anamnesebogen pa&cken (Stringfeldlängen optimieren)
Private Sub AnamnesebogenPacken_Click()
 Call ProgStart
 Call AnPack
 Call ProgEnde
End Sub ' AnamnesebogenPacken_Click

' EDV -> Anamnesebogen von u:\Anamnese\Quelle.mdb kopieren
' unsichtbar
#If False Then
Private Sub AnamnesebogenHolen_Click() ' nicht sichtbar: "Anamnesebogen von u:\Anamnese\Quelle.mdb &kopieren
 Call ProgStart
 Call holAB(Me)
 Call ProgEnde
End Sub ' AnamnesebogenHolen_Click
#End If

' EDV -> Dokumente abgehakt von u:\Anamnese\Quelle.mdb kopieren
' unsichtbar
#If False Then
Private Sub DokumenteAbgehaktkopieren_Click() ' nicht sichtbar: &Dokumente abgehakt von u:\Anamnese\Quelle.mdb kopieren"
 Call ProgStart
 Call holDA(Me)
 Call ProgEnde
End Sub ' DokumenteAbgehaktkopieren_Click
#End If

' EDV -> Kassenliste von u:\Anamnese\Quelle.mdb kopieren
' unsichtbar
#If False Then
Private Sub KassenlisteKopieren_Click() ' nicht sichtbar: &Kassenliste von u:\Anamnese\Quelle.mdb kopieren"
 Call ProgStart
 Call holAllg(Me, "kassenliste", "ID", 0)
 Call ProgEnde
End Sub ' KassenlisteKopieren_Click
#End If

' EDV -> Medarten von u:\Anamnese\Quelle.mdb kopieren
' unsichtbar
#If False Then
Private Sub holMedArten_Click() ' nicht sichtbar: &Medarten von u:\Anamnese\Quelle.mdb kopieren"
 Call ProgStart
'Function holMA(frm AS Lese)
' Rökan <> Roekan =>
'ALTER TABLE `medplan` drop FOREIGN KEY MedArtenMedPlan_AccRel;
'ALTER TABLE `medplan` MODIFY COLUMN `MedAnfang` VARCHAR(35) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL;
'DELETE FROM `medarten`;
'ALTER TABLE `medarten` modify column Medikament varchar(50) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL;
'INSERT INTO `medarten` (medikament) SELECT DISTINCT medanfang FROM `medplan` m;
'ALTER TABLE  `medplan` ADD CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`);
' IF lies.obmysql THEN
'  Call myEFrag("SET foreign_key_checks = 0")
'  Call myEFrag("DELETE FROM `medarten`")
' END IF
 Call holAllg(Me, "medarten", "Medikament", 0)
' IF lies.obmysql THEN
'  Call myEFrag("SET foreign_key_checks = 1")
' END IF
 Call ProgEnde
End Sub ' holMedArten_Click
#End If

' EDV -> Laborparameter von u:\Anamnese\Quelle.mdb kopieren
' unsichtbar
#If False Then
Private Sub holLaborParameter_Click() ' nicht sichtbar: "&Laborparameter von u:\Anamnese\Quelle.mdb kopieren
 Call ProgStart
 Call holAllg(Me, "laborparameter", vNS, 0)
 Call ProgEnde
End Sub ' holLaborParameter_Click
#End If

' EDV -> Tabelle kopieren
Private Sub TabKop_Click() ' Tabelle kopieren
 Set Tü.F0 = Me
 Tü.Show
End Sub ' TabKop_Click

' EDV -> Zahlen aus Einträgen ermitteln
Private Sub ZahlEintrag_Click()
 Call Haupt.ZahlEintrag
End Sub ' ZahlEintrag_Click

' EDV -> Fußsyndrome Wagner-Stadium 0 ergänzen
Private Sub WSt0Erg_Click()
 Call ProgStart
 Call doWSt0Erg
 Call ProgEnde
End Sub ' WSt0Erg_Click

#If uralt Then
'' EDV -> EBM2000plus-Ziffern holen
'Private Sub holEBM2000plus_Click()
' Call ProgStart
' Call holAllg(Me, "ebm2000plus", vNS, 0)
' Call ergEBM(Me)
' Call ProgEnde
'End Sub ' holEBM2000plus_Click
#End If

' EDV -> Therapiearten festlegen
Private Sub TherapieartenFestlegen_Click() ' Therapiearten festlegen
' Dim nTher AS TherapieArt, rAF&
 Call ProgStart
' Dim rf As New ADODB.Recordset
' IF 1 = 0 THEN
'  myFrag rF, "SELECT * FROM `faelle` ORDER BY fid "
'  Do While Not rF.EOF
'   nTher = therart(rF!Pat_id, True, , qend(rF!Quartal), rF!Quartal)
'   IF ISNULL(rF!therart) OR rF!therart <> nTher THEN
'    Call myEFrag("UPDATE `faelle` SET therart = " & nTher & " WHERE pat_id = " & rF!Pat_id & " AND quartal = '" & rF!Quartal & "'", rAF)
'    syscmd 4, "Therapieart für " & rF!Pat_id & " " & rF!Nachname & " " & rF!Vorname & ", Quartal: " & rF!Quartal & " zu " & TherUmw(nTher) & " festgelegt."
'   END IF
'   rF.MoveNext
'  Loop
' END IF
 
 Dim rsAna As New ADODB.Recordset
' rsAna.Open "SELECT pat_id,diabetestyp,-insulinpumpe AS j_insulinpumpe,ther1,therakt FROM `anamnesebogen`", DBCn, adOpenStatic, adLockOptimistic
' rsAna.Open "SELECT pat_id,diabetestyp,insulinpumpe,ther1,therakt FROM `anamnesebogen`", DBCn, adOpenStatic, adLockOptimistic
' rsAna.Open "SELECT an.pat_id, diabetestyp,insulinpumpe,ther1,therakt, (SELECT MAX(fanf) FROM `faelle` WHERE pat_id = an.pat_id) fanf FROM `anamnesebogen` an LEFT JOIN `faelle` f ON an.pat_id = f.pat_id GROUP BY pat_id ORDER BY (SELECT MAX(fanf) FROM `faelle` WHERE pat_id = an.pat_id) DESC", DBCn, adOpenStatic, adLockOptimistic
 myFrag rsAna, "SELECT n.pat_id, gesname(n.pat_id), MIN(f.fanf) FROM namen n LEFT JOIN faelle f ON n.pat_id = f.pat_id WHERE f.pat_id>000 AND f.pat_id <= 70000 GROUP BY n.pat_id ORDER BY fanf DESC"
 Do While Not rsAna.EOF
'  Call TherapieArtEinzelnFestlegen(rsAna!Pat_id, rsAna)
  Call rufThFestleg(rsAna!Pat_id)
  rsAna.Move 1
 Loop
 Ausgeb "Fertig mit Festlegen der Therapiearten", True
 Call ProgEnde
End Sub ' TherapieartenFestlegen_Click

' EDV -> Hausärzte übertragen
' unsichtbar
#If False Then
'Private Sub HAUebertrag_Click() ' nicht sichtbar: "Haus&ärzte übertragen
' Call ProgStart
' Call ergänzeliste
' Call holAllg(Me, "hausaerzte", "ID", -1)
' Call ProgEnde
'End Sub ' HAUebertrag_Click
#End If

' EDV -> Kassenkategorien bestimmen
Public Sub KassenkategorienBestimmen_Click()
 RufKassenKategorienBestimmen
Forms(0).Ausgeb "Fertig mit Bestimmen der Krankenkassenkategorien!", True
End Sub ' KassenkategorienBestimmen_Click

' EDV -> Dokumentpfade korrigieren
Private Sub DokPfadKorrigieren_Click() ' Dokumentpfade korrigieren
 Call ProgStart
 Call dokpfadänder(Me)
 Call ProgEnde
End Sub ' DokPfadKorrigieren_Click

' EDV -> Vergleich der Datenbankstrukturen
Private Sub Vergleichen_Click() ' Vergleich der Datenbankstrukturen
 Call ProgStart
 Call Vergleiche(Me)
 Call ProgEnde
End Sub ' Vergleichen_Click()

' EDV -> Vergleich der Datenbanktabellenf&üllungen
Private Sub TabVergleich_Click()
 Call ProgStart
 Call VergleichTab(Me)
 Call ProgEnde
End Sub ' TabVergleich_Click

' EDV -> Kassen aus Turbomed-Ex&cel-Datei einlesen
Private Sub KassenLesen_Click()
 Call LiesKassen
End Sub ' KassenLesen_Click

' EDV -> Excelliste
Private Sub Excelliste_Click()
 Dim lad As New Lade
 Call lad.Show
End Sub ' Excelliste_Click

' EDV -> falschen Laboreintrag zu Pat. löschen
Private Sub falschenLaboreintragZuPatlöschen_Click()
 Dim PIDStr$, pid&, TagStr$, Tag As Date, erg$, PName$, ltxt$, rAf&
 Dim rs As New Recordset
 PIDStr = InputBox("Bitte Pat_ID eingeben", "1.Rückfrage", 0)
 If Not IsNumeric(PIDStr) Then Exit Sub
 pid = CLng(PIDStr)
 If pid = 0 Then Exit Sub
 PIDStr = pid
 TagStr = InputBox("bitte zu löschendes Datum eingeben", "2.Rückfrage", "1.3.2002")
 If Not IsDate(TagStr) Then Exit Sub
 Tag = CDate(TagStr)
 myFrag rs, "SELECT gesnameg(pat_id) Name FROM namen WHERE pat_id = " & PIDStr
 If Not rs.BOF Then If Not rs.EOF Then PName = rs!name
 erg = MsgBox("Wollen Sie Labordaten von Pat. " & PIDStr & " (" & PName & ") vom " & Format(Tag, "dd.mm.yyyy") & " löschen?" & vbCrLf & "Bitte vorher prüfen, ob nicht durch Korrektur von 'pat_id' IN laboryus eine Zuordnung zu anderem Patienten nötig", vbYesNo)
 If erg <> vbYes Then Exit Sub
 Set rs = Nothing
 myFrag rs, "SELECT u.id,u.datid,u.Nachname,u.vorname,u.gebdat,u.eingang,e.pfad,u.lwerte," & _
 "(SELECT COUNT(0) FROM laborybakt WHERE usid = u.id) bzl, " & _
 "(SELECT COUNT(0) FROM laboryleist WHERE usid = u.id) lzl, " & _
 "(SELECT COUNT(0) FROM laboryus WHERE usid = u.id) uzl, " & _
 "(SELECT COUNT(0) FROM laborywert WHERE usid = u.id) bzw " & _
 "FROM laboryus u " & _
 "LEFT JOIN laborydat e ON u.datid = e.datid " & _
 "WHERE pat_id = " & PIDStr & " AND DATE(eingang) = " & Format(Tag, "yyyymmdd")
 If Not rs.BOF Then
  Open pVerz & "löschegleich " & Format(Now(), "yyyymmdd_MMHHSS") & ".txt" For Output As #59
  Do While Not rs.EOF
   ltxt = "USID: " & rs!id & ", DatID: " & rs!DatID & ", Name: " & rs!Nachname & ", Vorname: " & rs!Vorname & ", Geb: " & rs!GebDat & ", Eingang: " & rs!Eingang & vbCrLf & "aus Datei: " & rs!Pfad & vbCrLf & rs!LWerte & vbCrLf & "Baktzahl: " & rs!bzl & ", Leistzahl: " & rs!lzl & ", USZahl: " & rs!uzl & ", Wertzahl: " & rs!bzw
   erg = MsgBox("Löschen? " & ltxt, vbYesNo)
   If erg = vbYes Then
    Print #59, ltxt
    myEFrag "DELETE FROM laborybakt WHERE usid = " & rs!id, rAf
    ltxt = rAf & " Sätze aus laborybakt gelöscht"
    Print #59, ltxt
    myEFrag "DELETE FROM laboryleist WHERE usid = " & rs!id, rAf
    ltxt = rAf & " Sätze aus laboryleist gelöscht"
    Print #59, ltxt
    myEFrag "DELETE FROM laborywert WHERE usid = " & rs!id, rAf
    ltxt = rAf & " Sätze aus laborywert gelöscht"
    Print #59, ltxt
    myEFrag "DELETE FROM laboryus WHERE usid = " & rs!id, rAf
    ltxt = rAf & " Sätze aus laboryus gelöscht"
    Print #59, ltxt
   End If
   rs.MoveNext
  Loop
  Close #59
 End If
End Sub ' falschenLaboreintragZuPatlöschen_Click

' EDV -> Labortests zuordnen
Private Sub LabortestsZuordnen_Click()
 Dim lad As New Lade
 lad.FenArt = LaborArten
 Call lad.Show
End Sub ' LabortestsZuordnen_Click

' EDV -> Laborvergleich
Private Sub Laborvergleich_Click()
 Dim rv As New ADODB.Recordset, rs As New ADODB.Recordset, i&, ausg$, TA1$, spmax%(5), FristS$, sql$
 sql = "SELECT * FROM (SELECT COUNT(0) Zahl,Labor,trim(CONCAT(LEFT(CONCAT(abkü,'          '),10),LEFT(CONCAT(einheit,'            '),12),LEFT(nb,26))) `Verfahren/Einheit/   Normbereich`,Langtext, MAX(zeitpunkt) MaxEing, MIN(zeitpunkt) MinEing FROM labor2a GROUP BY abkü, einheit, nb, langtext, labor ORDER BY langtext, `Verfahren/Einheit/   Normbereich`) i;"
 myFrag rs, sql
 TA1 = TabAusgeb(rs, Me, , , , , , False, "Laborvergleich")
End Sub ' Laborvergleich_Click

' EDV -> Labor (direkt -> ""X"") l&öschen ab
Private Sub LaborLöschenAb_Click()
 Dim sql$, rs As New ADODB.Recordset, rAf&
 Dim DatumS$, Datum As Date, nr&
 Dim krit0$, krit1$, krit2$, krit3$, erg$
 Do
  DatumS = InputBox("ab welchem Datum löschen?")
  If IsDate(DatumS) Then Exit Do
 Loop
 Datum = CDate(DatumS)
 
 Ausgeb "", True
 
 Set rs = Nothing
 myFrag rs, "SELECT COUNT(0) FROM labor2a WHERE zeitpunkt >= " & Format(Datum, "yyyymmdd")
 nr = 0
 If Not rs.BOF Then nr = rs.Fields(0)
 Ausgeb "Aus laborywert würden gelöscht: " & nr & " Sätze", True
 
 Set rs = Nothing
 myFrag rs, "SELECT COUNT(0) FROM labor2bakt WHERE eingang >= " & Format(Datum, "yyyymmdd")
 nr = 0
 If Not rs.BOF Then nr = rs.Fields(0)
 Ausgeb "Aus laborybakt würden gelöscht: " & nr & " Sätze", True
 
 Set rs = Nothing
 myFrag rs, "SELECT COUNT(0) FROM laboryleist l LEFT JOIN laboryus u ON l.usid=u.id WHERE eingang >= " & Format(Datum, "yyyymmdd")
 nr = 0
 If Not rs.BOF Then nr = rs.Fields(0)
 Ausgeb "Aus laboryleist würden gelöscht: " & nr & " Sätze", True
 
 Set rs = Nothing
 myFrag rs, "SELECT COUNT(0) FROM `laboryus` WHERE eingang>= " & Format(Datum, "yyyymmdd")
 nr = 0
 If Not rs.BOF Then nr = rs.Fields(0)
 Ausgeb "Aus laboryus würden gelöscht: " & nr & " Sätze", True
 
 Set rs = Nothing
 myFrag rs, "SELECT COUNT(DISTINCT d.satzid) FROM laborysaetze d LEFT JOIN laboryus u ON u.satzid = d.satzid WHERE eingang >= " & Format(Datum, "yyyymmdd")
 nr = 0
 If Not rs.BOF Then nr = rs.Fields(0)
 Ausgeb "Aus laborysaetze würden gelöscht: " & nr & " Sätze", True
 
 Set rs = Nothing
 myFrag rs, "SELECT COUNT(DISTINCT d.datid) FROM laborydat d LEFT JOIN laboryus u ON u.datid = d.datid WHERE eingang >= " & Format(Datum, "yyyymmdd")
 nr = 0
 If Not rs.BOF Then nr = rs.Fields(0)
 Ausgeb "Aus laborydat würden gelöscht: " & nr & " Sätze", True
 
 erg = MsgBox("Wollen Sie wirklich alle LaboreINträge ab " & Datum & " löschen?", vbYesNo Or vbDefaultButton2, "Rückfrage")
 If erg = vbNo Then Exit Sub
 ' myEFrag "DELETE FROM `laborxwert` WHERE refnr IN " & krit0, rAF
 ' myEFrag "DELETE FROM `laborxbakt` WHERE refnr IN " & krit0, rAF
 ' myEFrag "DELETE FROM `laborxleist` WHERE refnr IN " & krit0, rAF
 myEFrag "DELETE d FROM laborydat d LEFT JOIN laboryus u ON u.datid = d.datid WHERE eingang >= " & Format(Datum, "yyyymmdd"), rAf
End Sub ' LaborLöschenAB

' EDV -> DMP-Liste erstellen
' unsichtbar
'Private Sub DMPListe_Click()
' call progstart
' Call DMPlst.init(Me)
' DMPlst.Show
'End Sub

' EDV -> DMP-String
Sub CallDMPString_Click()
 Dim erg$
 Call ProgStart
 erg = InputBox("Bitte Patientenummer eingeben!")
 If IsNumeric(erg) Then
  Call doCallDMP(ByVal erg)
 End If
 Call ProgEnde
End Sub ' CallDMPString_Click

' EDV -> Diagnosen sortieren
Private Sub DiagnosenSortieren_Click()
 Dim pad As New PatListe
 pad.PLArt = artDiag
 Set pad.hlese = Me
 pad.Show
End Sub ' DiagnosenSortieren_Click

' EDV -> Diagnosen exportieren
Private Sub DiagnosenExportieren_Click()
 Call GesDiagExp
End Sub ' DiagnosenExportieren_Click

' EDV -> Diagnosen exportieren (test)
Private Sub DiagnosenExportierenTest_Click()
 Call GesDiagExp(obTest:=True)
End Sub ' DiagnosenExportierenTest_Click

' EDV -> Formulare übertragen
Private Sub FormulareÜbertragen_Click()
 Call ProgStart
 Call ÜbertragFormulare(Me, "_Medarten", "Medarten", uVerz & "zugriff.mdb", "medarten")
 Exit Sub
 Call ÜbertragFormulare(Me, "Anamnesebogen", "AnBog", uVerz & "zugriff.mdb", "anamnesebogen")
 Call ÜbertragFormulare(Me, "Labordokumente eP", "Labordokumente", uVerz & "zugriff.mdb", "dokumente")
 Call ProgEnde
End Sub ' FormulareÜbertragen_Click

' EDV -> Boolean-Felder in MySQL-Datenbanken erstellen
Private Sub BooleanFelder_Click()
 Call ProgStart
 Call BooleanFld(Me)
 Call ProgEnde
End Sub ' BooleanFelder_Click

' EDV -> ViewsErstellen
Private Sub ViewsErstellen_Click() ' ViewsErstellen
 Call doViewsErstellen
End Sub ' ViewsErstellen_Click()

' EDV -> Falsche Dokumente
Private Sub FalscheDokumente_Click()
 Dim sql$, rs As New ADODB.Recordset, erg$(), dokn$, dokr$, i%, rs2 As New ADODB.Recordset, Pat_id&
 sql = "SELECT d.pat_id, dokname FROM dokumente d LEFT JOIN namen n ON d.pat_id = n.pat_id LIMIT 10000"
 myFrag rs, sql
 Do While Not rs.EOF
  dokn = rs!DokName
  dokr = REPLACE$(dokn, ",", " ")
  For i = 1 To 10
   dokr = REPLACE$(dokr, "  ", " ")
  Next
  SplitNeu dokr, " ", erg
  If UBound(erg) > 1 Then
   Pat_id = rs!Pat_id
   Set rs2 = Nothing
   myFrag rs2, "SELECT * FROM namen WHERE nachname = "" & erg(0) & "" AND vorname = "" & erg(1) & "" AND pat_id <> " & Pat_id
   If Not rs2.EOF() Then
    Debug.Print rs2!Pat_id, erg(0), erg(1), Pat_id, rs!DokName
   End If
  End If
  rs.MoveNext
 Loop
End Sub ' FalscheDokumente_Click


' Leistungseingabe (gelöscht)

' EDV -> Quelldatum für alle Dokumente korrigieren
Private Sub korrQD_Click() ' Quelldatum für alle Dokumente korrigieren
 Dim rs As New ADODB.Recordset, nQD As Date, rAf&, rsum&
 Call ProgStart
 myEFrag "UPDATE dokumente SET quelldatum=quelldat(dokname,DokAenD);", rsum
' myFrag rs, "SELECT * FROM `dokumente`"
' Do While Not rs.EOF
'  nQD = doQuelldatum(rs!DokName)
'  IF nQD <> rs!Quelldatum OR (nQD <> 0 AND ISNULL(rs!Quelldatum)) THEN
'   myEFrag "UPDATE `dokumente` SET quelldatum = " & DatFor_k(nQD) & " WHERE pat_id = " & rs!Pat_id & " AND fid = " & rs!FID & " AND dokpfad = '" & doUmwfSQL(rs!DokPfad, lies.obMySQL) & "'", rAF
'   IF rAF <> 1 THEN
'    rsum = rsum & rAF
'   END IF
'  END IF
'  rs.Move 1
' Loop
 Ausgeb rsum & " Quelldaten korrigiert", True, True
 Beep 1000, 1000
End Sub ' korrQD_Click

' EDV -> Quelldatum für alle Brief korrigieren
Private Sub korrQB_Click() ' Quelldatum für alle Briefe korrigieren
 Dim rs As New ADODB.Recordset, nQD As Date, rAf&, rsum&
 Call ProgStart
 myEFrag "UPDATE quelle.briefe SET quelldatum=quelldat(name,DokAenD);", rsum
' myFrag rs; "SELECT * FROM `briefe`"
' Do While Not rs.EOF
'  nQD = doFLQuelldatum(rs!name)
'  IF nQD <> rs!Quelldatum OR (nQD <> 0 AND ISNULL(rs!Quelldatum)) THEN
'   myEFrag "UPDATE `briefe` SET quelldatum = " & DatFor_k(nQD) & " WHERE pat_id = " & rs!Pat_id & " AND fid = " & rs!FID & " AND pfad = '" & doUmwfSQL(rs!Pfad, lies.obMySQL) & "'", rAF
'   IF rAF <> 1 THEN
''    rsum = rsum & rAF
'   END IF
'  END IF
'  rs.Move 1
' Loop
 Ausgeb rsum & " Quelldaten korrigiert", True, True
 Beep 1000, 1000
End Sub ' korrQB_Click

' EDV -> Apothekenrezepte
Private Sub Apothekenrezepte_Click()
' SELECT nachname, vorname, DATE(gebdat) AS geb, fr.zeitpunkt, fa.feldinh from`formular`fr LEFT JOIN `formular` fa USING (foid) LEFT JOIN `namen` ON fr.pat_id = `namen`.pat_id WHERE fr.feldinh LIKE "%Gerald Schade;" AND NOT ISNULL(fr.pat_id) AND fr.formvorl LIKE '%rezept%' AND ((fr.formvorl LIKE '%lang%' AND fa.feld = 'medikament') OR (fr.formvorl NOT LIKE '%lang%' AND fa.nr IN (4,9,10,11))) AND fr.zeitpunkt BETWEEN '2008-02-01' AND now() AND NOT fa.feldinh LIKE '%-  -%'
 Dim rs As New ADODB.Recordset
 Call ProgStart
 Open uVerz & "Apotheke.csv" For Output As #333
 myFrag rs, "SELECT foid, nachname, vorname, DATE(gebdat) AS geb, fr.zeitpunkt AS Zeitp, fa.feldinh AS text FROM `formular` fr LEFT JOIN `formular` fa USING (foid) LEFT JOIN `namen` ON fr.pat_id = `namen`.pat_id WHERE fr.feldinh LIKE '%Gerald Schade;' AND NOT ISNULL(fr.pat_id) AND fr.formvorl LIKE '%rezept%' AND ((fr.formvorl LIKE '%lang%' AND fa.feld = 'medikament') OR (fr.formvorl NOT LIKE '%lang%' AND fa.nr IN (4,8,9,10,11))) AND fr.zeitpunkt BETWEEN '2008-02-01' AND now() AND NOT fa.feldinh LIKE '%-  -%'"
 Do While Not rs.EOF
  Print #333, rs!Foid & ";" & rs!Nachname & ";" & rs!Vorname & ";" & rs!Geb & ";" & rs!Zeitp & ";" & rs!text
  rs.Move 1
 Loop
 Close #333
 MsgBox "Fertig mit Apotheke!"
End Sub ' Apothekenrezepte_Click

' EDV -> Einlesungen anzeigen
Private Sub EinlesungenAnzeigen_Click()
 Dim rEinl As New ADODB.Recordset, sql$
' sql = "SELECT st.*, e.datei As Datei, e.dateiaend AS `letzte Änderung`, e.beginn AS Einl_Beg, e.zp1, e.zp6, e.zp7, e.fallzahl AS Fälle, e.sekunden AS s FROM (SELECT COUNT(0) AS Zahl, MAX(pat_id) AS Pat_id, stbyte AS StByte FROM `namen` GROUP BY stbyte) AS st LEFT JOIN `eintragszahlen` e ON st.stbyte = e.stbyte"
 sql = "SELECT Beginn, Zp1, COALESCE(Zp4,'') Zp4, COALESCE(Fallzahl,'') Fallzahl, COALESCE(Sekunden,'') s, Datei, DateiAend, obvglmitletzterEinlesung obVgl FROM eintragszahlen ORDER BY beginn DESC LIMIT 100"
 Call ProgStart
 myFrag rEinl, sql
 TabAusgeb rEinl, Me, True
End Sub ' EinlesungenAnzeigen_Click

' EDV -> Dokumente abgehakt prüfen
Private Sub DokumenteAbgehaktPrüfen_Click()
 Dim rs As New ADODB.Recordset, n&, n1&
 On Error GoTo fehler
 Dim ErgDat$, erg$
 ErgDat = uVerz & "FD.txt"
 Do
  erg = Dir(ErgDat)
  If LenB(erg) = 0 Then Exit Do
  ErgDat = REPLACE$(ErgDat, "FD", "FD1")
 Loop
 Open ErgDat For Output As #7
 Call ProgStart
 myFrag rs, "SELECT dokname, da.*, d.* FROM `dokumente abgehakt` da LEFT JOIN `dokumente` d USING (dokpfad) ORDER BY d.pat_id, d.zeitpunkt"
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
 zeigan ErgDat
 ProgEnde
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DokumenteAbgehaktPrüfen_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DokumenteAbgehaktPrüfen_Click

' EDV -> Tabellenfüllungen untersuchen (tabfuell)
Private Sub tabfuell_Click()
 Call TabFuellSnSh
End Sub ' tabfuell_Click

' EDV -> Dokumente neu abhaken
Private Sub DokumenteNeuAbhaken_Click()
 Dim rs As New ADODB.Recordset, rAf&
 Call ProgStart
 Call myEFrag("DELETE FROM `dokumente abgehakt`", rAf)
 'Call myEFrag("INSERT INTO `dokumente abgehakt`(aktzeit,abgehakt,dokpfad) SELECT now() AS aktzeit,1 AS abgehakt, replace$(replace$(dokpfad,'\\','\\\\'),'\'','\\\'') FROM (SELECT * FROM (SELECT pat_id, zeitpunkt FROM labor1 GROUP BY pat_id, zeitpunkt ORDER BY pat_id, zeitpunkt) AS i LEFT JOIN (SELECT pat_id, DATE(quelldatum) AS zeitpunkt, dokpfad, dokname FROM `dokumente` d WHERE dokname LIKE '%fremdlabor%') AS d USING (pat_id,zeitpunkt)) AS i WHERE NOT ISNULL(dokpfad)", rAF)
 Debug.Print "gelöscht:", rAf
 InsKorr DBCn, DBCnS, "INSERT INTO `dokumente abgehakt`(aktzeit,abgehakt,dokpfad) " & vbCrLf & _
 "SELECT now() AS aktzeit,1 AS abgehakt, pfad FROM (" & vbCrLf & _
 "SELECT * FROM (" & vbCrLf & _
 "SELECT pat_id, zeitpunkt FROM labor1a GROUP BY pat_id, zeitpunkt ORDER BY pat_id, zeitpunkt) AS i " & vbCrLf & _
 "LEFT JOIN (SELECT pat_id, DATE(quelldatum) AS zeitpunkt, pfad, name FROM `briefe` d " & vbCrLf & _
 "WHERE name LIKE '%fremdlabor%') AS d " & vbCrLf & _
 "using (pat_id,zeitpunkt)) AS i WHERE NOT ISNULL(pfad)", rAf
 Debug.Print "eingefügt:", rAf
 Beep 1000, 1000
End Sub ' DokumenteNeuAbhaken_Click

' EDV -> hareal neu aufbauen und namen.getha-Felder befüllen
Private Sub harealNeu_Click() ' `hareal` neu aufbauen
 Dim rAf&, ohd%, od%, infos$()
 Dim i&
 ProgStart
 '  0: Frau/Herrn
 '  1: Titel+Vorn+Nachn,
 '  2: Straße,
 '  3: PLZ+Ort,
 '  4: Faxnr,
 '  5: S.g./Liebe,
 '  6: DMPTyp2,
 '  7: DMPTyp1,
 '  8: Niederlassungsgebiet 3. Feld für einmal austauschen,
 '  9: Vorname,
 ' 10: Funktion ("Üw 207, HA"),
 ' 11: InnereAllg
 ' 12: KV-Nummer
 ' 13: Tel'nr.
 ' 14: Nachname,
 On Error Resume Next
 myEFrag "DROP TABLE `hareal`"
 myEFrag "TRUNCATE table `hausaerzte`"
 On Error GoTo fehler
 myEFrag "CREATE TABLE IF NOT EXISTS `hareal`(Anrede tinyint(1) comment '0=Frau,1=Herrn',Adressat varchar(1) comment 'Titel+Vorn+Nachn', Straße varchar(1), PLZOrt varchar(1), Fax varchar(1), Überschrift varchar(1), dmp2 tinyint(1) comment '0=nein,1=ja', dmp1 tinyint(1) comment '0=nein,1=ja', Niederlassungsgebiet varchar(1) comment 'Med.Fachrichtung', Vorname varchar(1), Funktion varchar(0) comment 'nur noch Schaltfeld, Inhalt in `namen`', InnereAllg tinyint(1) comment '1=Innere oder Allgemeinmedizin', kvnr integer(10) UNSIGNED primary key, Tel varchar(1), Nachname varchar(1))"
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT pat_id FROM `namen`"
 Do While Not rs.EOF
  ReDim infos(0)
'  Call getHausarztAlt(rs!pat_id, Infos, True)
  Dim rFa() As Faelle
  Dim rKv1() As kvnrue
  getHausarzt1 infos, rFa, rKv1, , rs!Pat_id
  If LenB(infos(12, 0)) <> 0 Then
   For i = 0 To UBound(infos, 2)
'    IF Infos(4, i) = "08131-85028" THEN Stop
    If i < 3 Then
     myEFrag "UPDATE `namen` SET getha" & CStr(i) & " = " & IIf(infos(12, i) = vNS, 0, infos(12, i)) & ", fnHA" & CStr(i) & " = '" & IIf(infos(10, i) = vNS, 0, infos(10, i)) & "' WHERE pat_id = " & rs!Pat_id, rAf
    End If
'    IF False THEN
     If infos(12, i) <> vNS Then
      myEFrag "SELECT kvnr FROM `hareal` WHERE LEFT(kvnr,7) = " & Left$(infos(12, i), 7) & " AND vorname = '" & infos(9, i) & "' AND nachname = '" & infos(14, i) & "'", rAf
      If rAf = 0 Then
       InsKorr DBCn, DBCnS, "INSERT INTO `hareal`(Anrede,Adressat,Straße,PLZOrt,Fax,Überschrift,dmp2,dmp1,Niederlassungsgebiet,Vorname,InnereAllg,kvnr,Tel,Nachname) VALUES(" & IIf(infos(0, i) = "Herr", 1, 0) & ",'" & infos(1, i) & "','" & infos(2, i) & "','" & infos(3, i) & "','" & infos(4, i) & "','" & infos(5, i) & "'," & IIf(infos(6, i) = vNS, 0, 1) & "," & IIf(infos(7, i) = vNS, 0, 1) & ",'" & infos(8, i) & "','" & infos(9, i) & "'," & IIf(infos(11, i) = "", 0, 1) & "," & IIf(infos(12, i) = vNS, 0, infos(12, i)) & ",'" & infos(13, i) & "','" & infos(14, i) & "')", rAf
'       IF rAF = 0 THEN Stop
      Else
       myEFrag "UPDATE `hareal` SET Anrede=" & IIf(infos(0, i) = "Herr", 1, 0) & ",Adressat='" & infos(1, i) & "',Straße='" & infos(2, i) & "',PLZOrt='" & infos(3, i) & "',Fax='" & infos(4, i) & "',Überschrift='" & infos(5, i) & "',dmp2=" & IIf(infos(6, i) = vNS, 0, 1) & ",dmp1=" & IIf(infos(7, i) = vNS, 0, 1) & ",Niederlassungsgebiet='" & infos(8, i) & "',Vorname='" & infos(9, i) & "',InnereAllg=" & IIf(infos(11, i) = vNS, 0, 1) & ",Tel='" & infos(13, i) & "',Nachname='" & infos(14, i) & "' WHERE kvnr = " & infos(12, i) & " AND vorname = '" & infos(9, i) & "' AND nachname = '" & infos(14, i) & "'", rAf
'       IF rAF = 0 THEN Stop
      End If
'     END IF
    End If
    Lese.Ausgeb "Pat_id: " & rs!Pat_id & ": geändert: " & rAf, False, True
   Next i
  End If
  rs.Move 1
 Loop
 Lese.Ausgeb "Fertig mit Neuaufbau von `hareal`!", True, True
 Exit Sub
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in harealNeu/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' harealNeu

' EDV -> Hausärzte aus kv-pdf-Datei einlesen
Private Sub HausärzteEinlesen_Click()
 Call doHausärzteEinlesen
End Sub ' HausärzteEinlesen_Click

' EDV -> Hausärzte aus kv-pdf-Datei einlesen
Private Sub alleHausärzteEinlesen_Click()
 Call doalleKVDateien
End Sub ' alleHausärzteEinlesen

' EDV -> Punktwerte EBM2010
Private Sub Punktwerte_Click()
 Dim Str$, Zahl#, pos&, rAf&
ProgStart
Dim rs As New ADODB.Recordset
myFrag rs, "SELECT * FROM `EBM2010`"
Do While Not rs.EOF
 Str = rs!pwerte
 Do
  pos = InStr(Str, "|")
  If pos <> 0 Then
   Str = Mid$(Str, pos + 1)
  Else
   Exit Do
  End If
 Loop
 Str = REPLACE(Str, " Euro", vNS)
 If IsNumeric(Str) Then
  Zahl = CDbl(REPLACE(Str, ".", ""))
 ElseIf InStrB(Str, " Punkte") <> 0 Then
  Str = REPLACE$(Str, " Punkte", vNS)
  If IsNumeric(Str) Then
   Zahl = CDbl(REPLACE(Str, ".", "")) * 0.03505
  Else
   Debug.Print rs!pwerte, Str
   Stop
  End If
 ElseIf InStrB(Str, " unbewertet") <> 0 Or rs!pwerte = vNS Then
  Zahl = 0
 Else
  Debug.Print rs!pwerte, Str
  Stop
 End If
 myEFrag "UPDATE `EBM2010` SET euro = '" & REPLACE(Zahl, ",", ".") & "' WHERE myid = " & rs!myid, rAf
 If rAf <> 1 And Zahl <> rs!euro Then
  MsgBox "Fehler in Punktwerte_Click: rAF <> 1 AND Zahl <> rs!Euro"
  Stop
 End If
 rs.MoveNext
Loop
MsgBox "Fertig mit Punktwerten!"
End Sub ' Punktwerte_Click

' Testfunktionen

' Testfunktionen -> Spinstr
Private Sub Spinstr_Click()
 Dim St1$, St2$
 St1 = "noch warst du n.och"
 St2 = "noch"
 Call doSp(St1, St2)
End Sub ' Spinstr_Click

' Testfunktionen -> t7
Private Sub t7_Click()
 Call ProgStart
 Call test7
' Call löschBezügeausLaborux(65)
' Call LöschDateiEintrag(419)
 Call ProgEnde
End Sub ' t7_Click

' Testfunktionen -> ADOXtet
Private Sub ADOXtet_Click()
 Call ProgStart
 Call adoxtest(dlg)
 Call ProgEnde
End Sub ' ADOXtet_Click

' Testfunktionen -> SyncTest
' auskommentiert 10.10.22
'Private Sub SyncTest_Click()
' Call formInhMach
'End Sub ' SyncTest_Click

' Testfunktionen -> Usdm
Private Sub CallUSDM_Click()
 Dim erg
 Call ProgStart
 ReDim rNa(0)
 erg = InputBox("Bitte Patientenummer eingeben!")
 If Not IsNumeric(erg) Then Exit Sub
 rNa(0).Pat_id = erg
 Call usdmAlt(True)
 Call ProgEnde
End Sub ' CallUSDM_Click

' Testfunktionen -> gefaxttrennen
Private Sub gefaxttrennen_Click()
 Call ProgStart
' Call dogefaxttrennen
 Call ProgEnde
End Sub ' gefaxttrennen_Click

' Testfunktionen -> doppelteFAxe
Private Sub doppelteFaxe_Click()
 Dim datnam$
 datnam = pVerz & "schongefaxte " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".txt"
 Open datnam For Output As #323
 Call acon(quelleT)
 Call acon(FaxT)
 Call dodoppelteFaxe(pVerz & "unkorrigiert")
 Print #323, "Fertig mit doppelteFaxe_Click!"
' Debug.Print "Fertig mit doppelteFaxe_Click!"
 Close #323
 zeigan datnam
End Sub ' doppelteFaxe_Click

' Testfunktionen -> alteHausärzte
' auskommentiert 10.10.22
'Private Sub alteHausärzte_Click()
' Dim ahc As New Adodb.Connection, ahr As New Adodb.Recordset, acr As New Adodb.Recordset
' Dim runde%
'' SET ahc = acon(haT, accDtb)
'' Call acon(HaT, q1Dtb)
' If LenB(DBCn) = 0 Or DBCn = "" Then Call acon(quelleT)
'
' Dim Fls As Files, Fl As File, lastDat#
' Dim FSO As New FileSystemObject
' On Error GoTo fehler
' Set Fls = FSO.GetFolder(AnamneseVerZeichnis1).Files
' For Each Fl In Fls
'  If Fl.name Like "KV*rzte*.mdb" Then
'   Set ahc = Nothing
'   ahc.Open CStrAcc & Fl.path
'   For runde = 1 To 2
'    Set ahr = Nothing
'    On Error Resume Next
'    ahr.Open "SELECT * FROM `" & IIf(runde = 1, "`hae`", "`haealt`") & "` WHERE gelöscht", ahc, adOpenDynamic, adLockReadOnly
'    If Err.Number <> 0 Then GoTo nrunde
'    On Error GoTo fehler
'    Do While Not ahr.EOF
'     Set acr = Nothing
'     myFrag acr, "SELECT * FROM `kvaerzte`.`hae` WHERE nachname = '" & ahr!Nachname & "' AND kvnr = '" & ahr!KVNr & "'" 'haecn
'     If acr.EOF Then
'      Set acr = Nothing
'      myFrag acr, "SELECT * FROM `kvaerzte`.`haealt` WHERE nachname = '" & ahr!Nachname & "' AND kvnr = '" & ahr!KVNr & "'" 'haecn
'      If acr.EOF Then
'       MsgBox "Stop in alteHausärzte_Click: " & vbCrLf & "acr.EOF"
'       Stop
'      End If
'     End If
'     ahr.Move 1
'    Loop
'nrunde:
'    On Error GoTo fehler
'   Next runde
'  End If
' Next Fl
' Exit Sub
'fehler:
'Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.name
'#Else
' AnwPfad = App.path
'#End If
'Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in alteHausärzte/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Sub ' alteHausärzte

' Testfunktionen -> MedKlass
Private Sub MedklassT_Click() ' Testfunktionen
 Call ProgStart
 Call doMedklassT
 Call ProgEnde
End Sub ' MedklassT_Click

' Testfunktionen -> test
Private Sub test_Click()
 Call bittest1
End Sub ' test_Click

' Testfunktionen -> Gewichte
Private Sub Gewichte_Click()
 Dim sql$, rs As New ADODB.Recordset, rAf&
  On Error Resume Next
  myEFrag "DROP TABLE `gewicht`", rAf
  On Error GoTo 0
  myEFrag "CREATE TABLE `gewicht`(id integer(10) auto_increment key, FID int(10), Pat_ID int(10), ZeitPunkt datetime, Gewicht DECIMAL(5,1), absPos int(10), AktZeit datetime, QS varchar(5), QT varchar(5), StByte int(10), inhNum double)"
 sql = "SELECT * FROM eintraege WHERE lower(art) = 'gewicht'"
 myFrag rs, sql
 Do While Not rs.EOF
  sql = "INSERT INTO `gewicht`(FID,Pat_ID,ZeitPunkt,Gewicht,absPos,AktZeit,QS,QT,StByte,inhNum) VALUES(" & rs!FID & "," & rs!Pat_id & "," & Format(rs!Zeitpunkt, "yyyymmddHHMMSS") & "," & REPLACE(MachNumerisch(rs!Inhalt, 0), ",", ".") & "," & rs!absPos & "," & Format(rs!AktZeit, "yyyymmddHHMMSS") & ",'" & rs!QS & "','" & rs!QT & "'," & rs!StByte & "," & rs!inhNum & ")"
  myEFrag sql, rAf
  rs.Move 1
 Loop
 Lese.Ausgeb "Fertig mit Gewichte_Click!", True
End Sub ' Gewichte_Click

' Testfunktionen -> Gewichtsabnahmekandidaten
Private Sub Gewichtsabnahmekandidaten_Click()
 Dim sql$, rs As New ADODB.Recordset
 sql = "SELECT tict.pat_id, " & _
 "LEFT(CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')),IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')),n.nachname,', ',n.vorname),25) Name, " & _
 "tict.zp zpicd,toad.zp zpoad, gmax.gewicht maxgew, gmax.ZeitPunkt maxzp, gmin.gewicht mingew, gmin.ZeitPunkt minzp, mp.medikament, d.icd Typ2ICD, " & _
 "(SELECT COUNT(art) FROM eintraege WHERE pat_id = tict.pat_id AND art = 'tk') etkz, " & _
 "(SELECT COUNT(art) FROM eintraege WHERE pat_id = tict.pat_id AND art = 'gs') egsz " & _
 "FROM therarten tict LEFT JOIN therarten toad ON toad.pat_id = tict.pat_id AND toad.zp > tict.zp " & _
 "LEFT JOIN (SELECT g.pat_id, g.gewicht, g.zeitpunkt FROM gewicht g JOIN (SELECT pat_id, MAX(gewicht) gewicht FROM gewicht GROUP BY pat_id) g2 ON g.pat_id = g2.pat_id AND g.gewicht= g2.gewicht GROUP BY pat_id) gmax ON gmax.pat_id = tict.pat_id " & _
 "LEFT JOIN (SELECT g.pat_id, g.gewicht, g.zeitpunkt FROM gewicht g JOIN (SELECT pat_id, MIN(gewicht) gewicht FROM gewicht GROUP BY pat_id) g2 ON g.pat_id = g2.pat_id AND g.gewicht= g2.gewicht GROUP BY pat_id) gmin ON gmin.pat_id = tict.pat_id " & _
 "LEFT JOIN namen n ON n.pat_id = tict.pat_id " & _
  "LEFT JOIN medplan mp ON mp.pat_id = tict.pat_id " & _
  "and (medikament LIKE '%by%' OR medikament LIKE '%lyx%' OR medikament LIKE '%vict%')" & _
  "LEFT JOIN diagview d ON d.pat_id = tict.pat_id AND d.gICD RLIKE '^E1[123]\.' " & _
"WHERE tict.therart = 'ICT' AND NOT ISNULL(toad.zp) AND gmax.Zeitpunkt < gmin.ZeitPunkt " & _
 "  AND gmax.Zeitpunkt < toad.zp AND gmin.ZeitPunkt > tict.zp " & _
 "GROUP BY tict.pat_id, tict.zp, toad.zp " & _
 "ORDER BY tict.pat_id"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , , , , , "Gewichtsabnahmekandidaten", , 0
End Sub ' Gewichtsabnahmekandidaten_Click

' Testfunktionen -> Sortierung ändern
Private Sub SortierungÄndern_Click()
 Call ProgStart
 Call doSortierungÄndern0
End Sub ' SortierungÄndern_Click

' Testfunktionen -> Generierte MachDB aufrufen
Private Sub calldoGenMachDB_Click()
 Dim DBn$
 DBn = InputBox("Datenbankname", "Eingabe der Datenbank", "quelle")
 If LenB(DBn) <> 0 Then
#If Modul_MachDBquelle_bas_dabei Then
  Call doMach_quelle(DBn, DBCn, DBVerb.Cpt, False)
#End If
 End If
End Sub ' Sub calldoGenMachDB_Click

' Testfunktionen -> testlqanf
Private Sub testlqanf_Click()
 Dim sql$, rs As New ADODB.Recordset
 Lese.ProgStart
 Dim rv As New ADODB.Recordset
 Dim FristS$
 myFrag rv, "SHOW CREATE VIEW `aktf`"
 FristS = rv.Fields(1)
 Set rv = Nothing
 FristS = Mid$(FristS, InStr(FristS, "inverval ") + 9) ' interval muss klein sein
 FristS = Left$(FristS, InStr(FristS, " ") - 1)
 If Not IsNumeric(FristS) Then
  MsgBox "Ungeeignete Abfrage `aktf`, evtl. erst Views erstellen"
  Exit Sub
 End If
' ktag fehlerhaft
' sql = "SELECT zp FROM (SELECT DATE(zeitpunkt) zp FROM eintraege WHERE zeitpunkt BETWEEN " & lQAnfuEnd(FristS, 1) & ") i GROUP BY zp ORDER BY zp DESC"
 sql = "SELECT zp FROM (SELECT DATE(zeitpunkt) zp FROM eintraege WHERE zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & ") i GROUP BY zp ORDER BY zp DESC"
 myFrag rs, sql
 TabAusgeb rs, Me, , , , , , , , , , , "Tage mit Einträgen dieses Quartal (FristS = " & FristS
End Sub ' testlqanf_Click

' Fenster

' Testfunktionen -> Wechseln
Private Sub Wechseln_Click()
 On Error Resume Next
 If pataw.geladen Then pataw.SetFocus
' IF Err.Number <> 0 THEN IF ple.geladen THEN ple.SetFocus
End Sub ' Wechseln_Click


' hierweiter

#If False Then
'Public EinlAb&
Private Sub TestFos_Click()
 Dim rs As New ADODB.Recordset
 Call ProgStart
 myFrag rs, "SELECT medikament, Urä, Ostp FROM `medarten` WHERE medikament = 'FOSAVANCE'"
 Debug.Print rs!Medikament, rs!urä, rs!ostp
 Call ProgEnde
End Sub ' TestFos_Click
#End If

' in Kontrolllisten_für_DMP_HA_Click
Function FileJoin&(strFile1$, strFile2$)
    Dim b() As Byte
    Dim nFileNum%
    Dim lFileLen&
    nFileNum = FreeFile
    Open strFile1 For Binary Access Read As nFileNum
        b() = InputB(LOF(nFileNum), nFileNum)
    Close nFileNum
    nFileNum = FreeFile
    lFileLen = FSO.GetFile(strFile2).size ' 21.8.21 statt Filelen
    Open strFile2 For Binary Access Write As nFileNum
        Put nFileNum, (lFileLen + 1), b()
    Close nFileNum
    FileJoin = lFileLen
End Function ' FileJoin

' in GNR_Statistiken_einl_Click
Sub doGNR_Statistiken_einl_Click(Optional obneu = 0)
 Const XStra = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
 Const XStrb = ";Extended Properties=""Excel 8.0;HDR=no;IMEX=1"""
 Dim Verz$
 Verz$ = "t:\kv-abrechnungen"
 Const GStat$ = "GNRStat"
 Const GZahl$ = "GNRZahl"
 Dim fgnr%, fleigru%, fpunkte%, feuro%, fm%, ff%, fr%, FZahl%, fmin%, FNr% ' Feldnummern
 Dim rX As New ADOX.Catalog, sql$, ka%, ke%, runde%, angefangen%, obAnfang%, i&, rAf&, erg$, labxtb$, DateiDat As Date
 Dim doeintr%, statid&
 Dim XCon As New ADODB.Connection
 Dim rEx As New ADODB.Recordset, rs As New ADODB.Recordset, rTest As New ADODB.Recordset
 
 If obneu Then
  On Error GoTo fehler
  DBCnS = DBCn
  On Error Resume Next
  labxtb = "labor_xls" & Int(CDbl(Now()) * 1000000)
  myEFrag ("DROP TABLE `" & GStat & "`")
  On Error GoTo fehler
  myEFrag "CREATE TABLE `" & GStat & "` (id integer(10) auto_increment key, datei varchar(255), dateidat datetime, qinv int(5) comment 'Quartal nach Jahr', index qinv(qinv))", rAf
  On Error Resume Next
  myEFrag ("DROP TABLE `" & GZahl & "`")
  On Error GoTo fehler
  myEFrag "CREATE TABLE `" & GZahl & "` (id integer(10) auto_increment key, statid integer(10), gnr varchar(20), leigru varchar(10), punkte integer(5), euro DECIMAL(5,2), m integer(5), f integer(5), r integer(5), zahl integer(10), wert DECIMAL(9,2), uwert DECIMAL(9,2), min integer(10))", rAf
 End If
 
 erg = Dir(Verz & "\GNR-Statistik*")
 If erg = "" Then
  Verz = tVerz & "kv-abrechnungen"
  erg = Dir$(Verz & "\GNR-Statistik*")
 End If
 Do While erg <> ""
'  IF erg = "GNR-Statistik_12.07.2021.xls" THEN Stop
  DateiDat = FileDateTime(Verz & "\" & erg)
  Ausgeb erg & " " & DateiDat, True
  Set XCon = Nothing
  On Error GoTo nichtoeffnen
  XCon.Open XStra & Verz & "\" & erg & XStrb
  On Error GoTo fehler
  Set rX = Nothing
  rX.ActiveConnection = XCon
  rEx.Open "`" & rX.Tables(rX.Tables.COUNT - 1).name & "`", XCon ' Hier Excel, nicht obmysql = 0!
  doeintr = 0
'  If InStr(erg, "GNR-Statistik_13.01.2023") <> 0 Then Stop
  Do While Not rEx.EOF
   Dim F0$
   If Not IsNull(rEx.Fields(0)) Then
    F0 = rEx.Fields(0)
    If InStrB(F0, "Erstellt am") = 1 Then
     Dim pZeitr%, Dat0 As Date, Dat1 As Date, q0$, q1$
     pZeitr = InStr(F0, "Zeitraum")
     If pZeitr <> 0 Then
      Dat0 = CDate(Mid$(F0, pZeitr + 11, 10))
      Dat1 = CDate(Mid$(F0, pZeitr + 26, 10))
      q0 = QuartalStr(Dat0)
      q1 = QuartalStr(Dat1)
      If q0 <> q1 Or Dat0 <> QAnf(q0) Or Dat1 <> QEnd(q1) Then
'       Stop
      Else
       Set rTest = Nothing
       myFrag rTest, "SELECT id, dateidat FROM `" & GStat & "` WHERE qinv = '" & Mid$(q0, 2) & Left$(q0, 1) & "'"
       If Not rTest.EOF Then
        If rTest!DateiDat >= DateiDat Then
         Exit Do ' nur die jüngste Datei eintragen
        Else
         myEFrag ("DELETE FROM `" & GZahl & "` WHERE statid = " & rTest!id)
         myEFrag ("DELETE FROM `" & GStat & "` WHERE id = " & rTest!id)
        End If
       End If
       InsKorr DBCn, DBCnS, "INSERT INTO `" & GStat & "` (datei,dateidat,qinv) values ('" & UmwfSQL(Verz & "\" & erg) & "'," & DatFor_k(DateiDat) & ",'" & Mid$(q0, 2) & Left$(q0, 1) & "')", rAf
       Set rTest = Nothing
'       Set rTest = myEFrag("SELECT last_insert_id()")
       Set rTest = myEFrag("SELECT id FROM `" & GStat & "` WHERE DATEI='" & UmwfSQL(Verz & "\" & erg) & "'")
       statid = rTest.Fields(0)
       If statid = 0 Then MsgBox "Fehler in doGNR_Statistiken_einl_Click: last_insert_id()=0"
       doeintr = 1
      End If ' q0 <> q1 OR Dat0 <> QAnf(q0) OR Dat1 <> qend(q1) Then else
     End If ' pZeitr <> 0 Then
    ElseIf doeintr = 1 Then
     If F0 = "GNR" Then
      fleigru = 0: fpunkte = 0: feuro = 0: fm = 0: ff = 0: fr = 0: FZahl = 0: fmin = 0
      fgnr = 0
      For FNr = 0 To rEx.Fields.COUNT - 1
       Select Case rEx.Fields(FNr)
        Case "Lei.Grp.": fleigru = FNr
        Case "Punkte": fpunkte = FNr
        Case "€": feuro = FNr
        Case "M": fm = FNr
        Case "F": ff = FNr
        Case "R": fr = FNr
        Case "Anzahl": FZahl = FNr
        Case "min": fmin = FNr
       End Select ' Case rEx.Fields(FNr)
      Next FNr
      doeintr = 2
     End If ' F0 = "GNR" Then
    ElseIf doeintr = 2 Then
      Dim punkte&, m&, F&, r&, Zahl&, MIN&, ppunkte&, pm&, pf&, pr&, pzahl&, pmin&, euro#, Wert#, uwert#
      If F0 = "Summe" Then Exit Do
      ppunkte = InStr(rEx.Fields(fpunkte), ","): If ppunkte = 0 Then punkte = IIf(IsNumeric(rEx.Fields(fpunkte)), rEx.Fields(fpunkte), "0") Else punkte = Left$(rEx.Fields(fpunkte), ppunkte - 1)
      pm = InStr(rEx.Fields(fm), ","): If pm = 0 Then m = rEx.Fields(fm) Else m = Left$(rEx.Fields(fm), pm - 1)
      pf = InStr(rEx.Fields(ff), ","): If pf = 0 Then F = rEx.Fields(ff) Else F = Left$(rEx.Fields(ff), pf - 1)
      pr = InStr(rEx.Fields(fr), ","): If pr = 0 Then r = rEx.Fields(fr) Else r = Left$(rEx.Fields(fr), pr - 1)
      pzahl = InStr(rEx.Fields(FZahl), ","): If pzahl = 0 Then Zahl = rEx.Fields(FZahl) Else Zahl = Left$(rEx.Fields(FZahl), pzahl - 1)
      pmin = InStr(rEx.Fields(fmin), ","): If pmin = 0 Then MIN = rEx.Fields(fmin) Else MIN = Left$(rEx.Fields(fmin), pmin - 1)
      euro = "0" & REPLACE$(rEx.Fields(feuro), ".", ",")
      Wert = euro * Zahl
      uwert = IIf((euro = 18.75 Or euro = 19.05 Or euro = 14.25) And Left$(rEx.Fields(fgnr), 1) = "9", 75, euro) * Zahl
      InsKorr DBCn, DBCnS, "INSERT INTO `" & GZahl & "` (statid,gnr,leigru,punkte,euro,m,f,r,zahl,wert,uwert,min) values (" & statid & ",'" & rEx.Fields(fgnr) & "','" & rEx.Fields(fleigru) & "'," & punkte & "," & REPLACE(euro, ",", ".") & "," & m & "," & F & "," & r & "," & Zahl & "," & REPLACE(Wert, ",", ".") & "," & REPLACE(uwert, ",", ".") & "," & MIN & ")", rAf
    End If ' InStrB(F0, "Erstellt am") = 1 Then elseif elseif doeintr = 2
   End If ' NOT ISNULL(rEx.Fields(0)) Then
'   Debug.Print rEx.Fields(0) ' , rEx.Fields(1), rEx.Fields(2), rEx.Fields(3), rEx.Fields(4), rEx.Fields(5), rEx.Fields(6), rEx.Fields(7), rEx.Fields(8), rEx.Fields(9), rEx.Fields(10), rEx.Fields(11), rEx.Fields(12)
   rEx.Move 1
  Loop
naechstedatei:
  Set rEx = Nothing
'  Exit Sub
  erg = Dir
 Loop
 
 Ausgeb "Fertig!", True
 Exit Sub
nichtoeffnen:
 GoTo naechstedatei
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doGNR_Statistiken_einl_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' doGNR_Statistiken_einl_Click

' in DokumenteInDatenbank_Click
Private Sub dverz(DPfad$)
 Dim FSOPfad As Folder, Fil As File, SubF As Folder, rAf&
 Set FSOPfad = FSO.GetFolder(DPfad)
 For Each Fil In FSOPfad.Files
  InsKorr DBCn, DBCnS, "INSERT INTO `Dokumente`(Pfad,Datei,größe,geändert) VALUES('" & doUmwfSQL(DPfad, True) & "','" & doUmwfSQL(Fil.name, True) & "'," & Fil.size & "," & DatFor_k(Fil.DateLastModified) & ")", rAf
 Next Fil
 For Each SubF In FSOPfad.SubFolders
  Call dverz(SubF.path)
 Next SubF
End Sub ' dverz(DPfad$)

' nur in doppelteFaxe und dodoppelteFaxe selbst
Private Sub dodoppelteFaxe(V$)
 Static FSO As New FileSystemObject
 Dim Fil As File, pid$, pos%, buch$
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
 Print #323, V
 Debug.Print "dodoppelteFaxe(" & V & ")"
 For Each Fil In FSO.GetFolder(V).Files
  If Fil.name Like "*PID *" Then
   pos = InStr(Fil.name, "PID ") + 4
   pid = vNS
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
    myFrag rs, "SELECT * FROM `outg` WHERE docn LIKE '%PID " & pid & "%' AND docn LIKE '%Arztbrief%'", adOpenStatic, FxCn, adLockReadOnly
    If Not rs.EOF Then
     If FSO.FileExists(pVerz & rs!docn) Then
      Print #323, "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe
      Debug.Print "zulöschen: ", rs!docn, rs!origst, rs!submt, rs!transe
      Kill Fil.path
     Else
      Set rs1 = Nothing
      myFrag rs1, "SELECT * FROM `briefe` WHERE name LIKE '%Arztbrief%' AND name LIKE '%PID " & pid & "%'"
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
  Call dodoppelteFaxe(fld.path)
 Next fld
End Sub ' dodoppeltefaxe

#If False Then
Private Sub falschebriefelöschen_Click()
 Dim Fil As File, pid&, pos&, p2&, rs As New ADODB.Recordset, an As New ADODB.Recordset, infos$(), rD As New ADODB.Recordset
 Call ProgStart
 For Each Fil In FSO.GetFolder(pVerz & "unkorrigiert").Files
  pos = InStr(Fil.name, "PID ")
  If pos > 0 Then
   p2 = InStr(pos, Fil.name, " ")
   pid = Mid$(Fil.name, pos + 4, p2 - pos + 1)
   Set rs = Nothing
   myFrag rs, "SELECT * FROM `faelle` WHERE pat_id = " & pid & " ORDER BY bhfb DESC;", adOpenStatic, DBCn, adLockOptimistic
   
   Set an = Nothing
'   Call an.Open("SELECT -tkz AS j_tkz, a.* FROM `anamnesebogen` a WHERE pat_id = " & pid, DBCn, adOpenStatic, adLockOptimistic)
   myFrag an, "SELECT * FROM `anamnesebogen` a WHERE pat_id = " & pid, adOpenStatic, DBCn, adLockOptimistic
   If rs!SchGr = 41 Or rs!SchGr = 42 Or rs!SchGr = 43 Then
    FSO.DeleteFile (Fil.path)
'   ElseIf an!j_tkz <> 0 THEN
   ElseIf an!Tkz <> 0 Then
'
    FSO.DeleteFile (Fil.path)
   Else
'     Call getHausarztAlt(Pid, Infos())
     Dim rFa() As Faelle
     Dim rKv1() As kvnrue
     getHausarzt1 infos(), rFa, rKv1, , pid
     If LenB(infos(1, 0)) = 0 Then
      FSO.DeleteFile Fil.path
     ElseIf infos(1, 0) Like "*Schade" Then
      Debug.Print Fil.name
      MsgBox "Stop in falschebriefelöschen_Click:" & "infos(1,0): " & infos(1, 0)
      Stop
     End If
    Set rD = Nothing
    myFrag rD, "SELECT * FROM `briefe` WHERE pat_id = " & pid & " AND name LIKE 'Brief an %'", adOpenStatic, DBCn, adLockOptimistic
    If Not rD.BOF Then
     FSO.DeleteFile (Fil.path)
    End If
'    SET rd = Nothing
'    myFrag rd, "SELECT * FROM `briefe` WHERE pat_id = " & pid & " AND name NOT LIKE '%labor%' AND name NOT LIKE '%Tagebuch%' AND name NOT LIKE '%Nachricht an%' AND name NOT LIKE '% BZ%' AND name NOT LIKE '%DMP-%' AND name NOT LIKE '%EKG%' AND name NOT LIKE '%anmeldung%' AND name NOT LIKE '%datenliste%' AND name NOT LIKE '%Herzzentrum%' AND name NOT LIKE '%Termin%' AND name NOT LIKE '%Attest%' AND name NOT LIKE '%Schreiben%' AND name NOT LIKE '%-Bakt%' AND name NOT LIKE '%Bescheinigung%' AND name NOT LIKE '%Dokumentation%' AND name NOT LIKE '%Vorschläge%' AND name NOT LIKE '%Pumpeneinstellung%' AND name NOT LIKE '%anforderung%' AND name NOT LIKE '%Medikamentenplan%' AND name NOT LIKE '%befund%' AND name NOT LIKE '%blutdruck%' AND name NOT LIKE '%analyse%' AND name NOT LIKE '%sonogramm%' AND name NOT LIKE '%übersicht%' AND name NOT LIKE '%auswertung%' AND name NOT LIKE '%GPD%' AND name NOT LIKE '%untersuchung%' AND name NOT LIKE '%standar_tag%'", adOpenStatic, DBCn, adLockOptimistic
'    IF Not rd.BOF THEN
'     Debug.Print rd!Name
'    END IF
   End If
  Else
   MsgBox "Stop in falschebriefelöschen_Click:" & "pos <= 0"
   Stop
  End If
 Next Fil
 Call ProgEnde
End Sub ' falschebriefelöschen_Click
#End If ' false

Private Sub MDIForm_Activate()
 If Command = "plz" Then
  Call doPatientenlaufzettel(obohnerueckfrage:=True, obphp:=True)
  Call ProgEnde
 ElseIf Command = "labor" Then
  ProgStart
  Lese.obMySQL = True
  LVobMySQL = Lese.obMySQL
  BezFeh = pVerz & "BezFehler_" & DefDB(DBCn) & "_" & Format(Now(), "YYYYMMDD_hhmmss") & ".txt"
  obMitAlterTab = True
  Call LaborDirektImport(Lese, 0, True, 0, xVerz & "Labor\backup\", 0)
  Call ProgEnde
 End If
 On Error Resume Next
 Me.dlg.Hide
End Sub ' MDIForm_Activate()

' in MyDB_Click, MyDB_LostFocus
Private Sub MyDB_Change()
 If Not obStart Then
   Screen.MousePointer = vbHourglass
   If Not obStart Then Call PutReg(Me)
'   obStart = True
   Me.obMySQL = True
'   obStart = False
   Me.dbv.DaBa = Me.MyDB
'   On Error Resume Next
'   DBCn.CommitTrans: obTrans = 0
'   On Error GoTo 0
   ComTrans
   If DBCn.State <> 0 Then DBCn.Close ' 12.12.09 + nächste Zeile
   DBCnS = Me.dbv.CnStr
   DBCn.Open DBCnS ' Me.dbv.wCn.ConnectionString
   Me.Visible = True
   Screen.MousePointer = 0
 End If
End Sub ' MyDB_Change()

Private Sub MyDB_Click()
 Call MyDB_Change
End Sub ' MyDB_Click

Private Sub MyDB_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me, Me.MyDB.name)
End Sub ' MyDB_KeyDown

Private Sub obMySQL_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me, Me.MyDB.name)
End Sub ' obMySQL_KeyDown

Private Sub MyDB_LostFocus()
 Call MyDB_Change
End Sub ' MyDB_LostFocus

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
End Sub ' obAcc_Click

' in QuelleDB.aCStr und obAcc_Click
Sub machODBCAcc()
  Me.dbv.ODBC = "Microsoft Access Driver (*.mdb)"
End Sub ' machODBCAcc

' in: MDIForm_Load -> Haupt.holreg
Private Sub obMySQL_Click()
 If Not obStart Then
  If Me.Visible Then Screen.MousePointer = vbHourglass
  Call dbv.cnVorb("", "anamnesebogen", "Patientendaten")
  Call machODBCMy
  Call acon(quelleT)
  Call PutReg(Me)
  If Me.Visible Then Screen.MousePointer = 0
 End If
End Sub ' obMySQL_Click

' in QuelleDB.aCStr und obMySQL_Click
Sub machODBCMy()
  Me.dbv.ODBC = ODBCStr()
End Sub ' machODBCMy

' in los
Public Sub ZeigGefaxteAn(Pat_id&, Optional PatName$)
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT transe `Übertragungsende`, docname `Dokumentname`, RCFax, pages `Seiten`, fsize `Größe`,Retries FROM `faxeinp`.`outa` o WHERE docname LIKE '%PID " & Pat_id & "%' ORDER BY transe DESC"
 TabAusgeb rs, Me, True, , , , , , vNS
 Me.Ausgeb "Gefaxt wurden an: " & PatName, 1
End Sub 'ZeitGefaxteAn

' aufgerufen in patauswahl: Sub OKButton_Click()
Public Sub los()
 Dim zzn%
 Select Case Aktion
  Case GefaxteAnzeigen
   Call ZeigGefaxteAn(Me.pataw.PatID, Me.pataw.PatName)
  Case Briefschreiben
   Call tubriefStandalone(Me.pataw.PatID, False, , Me.pataw.Vorlage, , Me.pataw.briefneu, Me.pataw.nichtherricht)
  Case RestlicheBriefe
   Call doRestlicheBriefe(Me, Me.pataw.PatID)
  Case Patientenlaufzetteleinzeln
'   Call doPLZeinzeln(Me.pataw.PatID)
    zzn = 8
    If IsNumeric(Me.pataw.Zeilenzahl) Then zzn = CInt(Me.pataw.Zeilenzahl)
   Call dodoplz(Me.pataw.PatID, plzVz, Now, Now - Int(Now), True, "", zzn, obRueck)
  Case DMPZettel
   Call einDMP(Me.pataw.Pat_id)
  Case Anwalt
   Call doAnwalt(Me.pataw.Pat_id)
 End Select
End Sub ' los

Private Sub DMPForts_Click()
 Static Ausw As New ADODB.Recordset
 Call ProgStart
 If Ausw.State = 0 Then
'  Ausw.Open "SELECT DISTINCT pat_id, nachname, vorname, gebdat FROM `dmpreihe` WHERE  (Abk LIKE 'eDMPDM%' OR Abk LIKE 'DMPDTYP%') AND datum > " & DatFor_k(Now() - 365 * 0.5) & " ORDER BY nachname, vorname;", DBCn, adOpenDynamic, adLockReadOnly
  myFrag Ausw, "SELECT DISTINCT pat_id, nachname, vorname, gebdat FROM `dmpreihe` WHERE  (Abk LIKE 'eDMPDM%' OR Abk LIKE 'DMPDTYP%') AND datum > " & DatFor_k(Now() - 365 * 0.5) & " ORDER BY nachname, vorname;", adOpenStatic
 End If
 Ausw.Find "pat_id = " & lDMPPat_id, , adSearchBackward, 1
 If Not Ausw.EOF Then
  Ausw.Move 1
  Call doCallDMP(Ausw!Pat_id)
 End If
End Sub ' DMPForts_Click

' in Ausgabe_KeyDown, CallDMPString_Click, DMPForts_Click
Sub doCallDMP(ByVal pid&)
 Dim dmpstD$, erg$, DT As DMPClass ' Dateiname
 Dim rsNa As New ADODB.Recordset
 ReDim rNa(0)
 rNa(0).Pat_id = pid
 myFrag rsNa, "SELECT * FROM `namen` WHERE pat_id = " & pid
 If Not rsNa.BOF And Not IsNull(rsNa!Nachname) And Not IsNull(rsNa!Vorname) Then
  dmpstD = GesNamFn(rsNa) & " (" & pid & ") "
 Else
  dmpstD = pVerz & "DmpString "
 End If
 rsNa.Close
 erg = DMPString$(rNa(0).Pat_id, DT)
 If lies.obMySQL Then
  dmpstD = dmpstD & Me.MyDB
 Else
  dmpstD = dmpstD & REPLACE$(REPLACE$(dlg.MdB, "\", "_"), ":", ".")
 End If
 dmpstD = Environ("userprofile") & "\documents\" & dmpstD & " " & Format$(Date, "dd.mm.yy") & ".txt"
 Open dmpstD For Output As #391
 Print #391, erg
 Close #391
 zeigan dmpstD
End Sub ' doCallDMP

'Private Sub DMPListeAnzeigen_Click()
' call progstart
' Call doDMPListeAnzeigen(Me)
' Call ProgEnde
'End Sub

'Private Sub LaborTestLaden_Click()
'Set labtest.dlg = Me.dlg
' labtest.Show
'End Sub

'Private Sub Einlesen_Click(index As Integer)
' obMitAlterTab = True
' call progstart
' Call doEinles(True)
' Call ProgEnde
'End Sub ' Einlesen_Click(Index As Integer)

'Private Sub EinlesenAb_Click()
' Dim erg
' obMitAlterTab = True
' call progstart
' erg = InputBox("Ab welcher Patientennummer soll eingelesen werden?", "Einlesen ab Patientennummer")
' IF IsNumeric(erg) THEN
'  EinlAb = erg
'  Call doEinles(False)
'  Call ProgEnde
' END IF
'End Sub

'Private Sub EinlesenEingelesene_Click()
' obMitAlterTab = True
' call progstart
' Call myEFrag("UPDATE quelle.`namen` SET aktzeit = 0")
' Call doEinles(False)
' Call ProgEnde
'End Sub ' EinlesenEingelesene_Click()

'Private Sub EinlesenohneLabor_Click()
' obMitAlterTab = True
' call progstart
' Call doEinles(False, True)
' Call ProgEnde
'End Sub ' EinlesenohneLabor_Click()

'Private Sub EinlesenSchnell_Click()
' obMitAlterTab = False
' call progstart
' Call doEinles(False)
' Call ProgEnde
'End Sub ' EinlesenSchnell_Click()
'Private Sub EinlesenEinzeln_Click()
' obMitAlterTab = True
' call progstart
' Call doEinles(False)
' Call ProgEnde
'End Sub ' EinlesenEinzeln_Click()
'Private Sub EinlesenVorb_Click()
' obMitAlterTab = True
' obVorber = True
' call progstart
' Call doEinles(True)
' Call ProgEnde
'End Sub ' EinlesenVorb_Click()
'Private Sub EinlesenEinzelnVorb_Click()
' obMitAlterTab = True
' obVorber = True
' call progstart
' Call doEinles(False)
' Call ProgEnde
'End Sub ' EinlesenEinzelnVorb_Click()

'Function ergänzeliste()
' Dim rs As New ADODB.Recordset, rs2 As New ADODB.Recordset, rAF&
' myFrag rs, "SELECT s1.name, s2.nachname, s1.vorname AS s1vorname, s2.vorname, s1.kvnr, s2.nkvnr FROM `liuez` s1 LEFT JOIN (SELECT nkvnr, `hae`.* FROM (SELECT n.kvnr nkvnr FROM `namen` n LEFT JOIN `liuez` h ON n.kvnr = h.kvnr WHERE ISNULL(h.name)) AS innen LEFT JOIN `kvaerzte`.`hae` ON innen.nkvnr = `hae`.kvnu  WHERE NOT ISNULL(dbnr) GROUP BY nkvnr) AS s2 ON s1.name = s2.nachname WHERE NOT ISNULL(nachname) AND (ISNULL(s1.kvnr) OR s1.kvnr='')"
'
' myFrag rs, "SELECT s1.name,s2.nachname, s1.vorname AS s1vorname, s2.vorname, s1.kvnr, s2.nkvnr FROM `liuez` s1 LEFT JOIN (SELECT nkvnr, `hae`.* FROM (SELECT n.kvnr nkvnr FROM `namen` n LEFT JOIN `liuez` h ON n.kvnr = h.kvnr WHERE ISNULL(h.name)) AS innen LEFT JOIN `kvaerzte`.`hae` ON innen.nkvnr = `hae`.kvnu  WHERE NOT ISNULL(dbnr) GROUP BY nkvnr) AS s2 ON s1.name = s2.nachname WHERE NOT ISNULL(nachname) AND (ISNULL(s1.kvnr) OR s1.kvnr='')"
' Do While Not rs.EOF
'  Call myEFrag("UPDATE `liuez` SET kvnr = '" & rs!nkvnr & "' WHERE name = '" & rs!Name & "' AND vorname = '" & rs!s1vorname & "'", rAF)
'  Debug.Print rs!Name, rs!Nachname, rs!s1vorname, rs!nkvnr, rAF
'  rs.Move 1
' Loop
'End Function

'Private Sub DokumentPfadeKorrigieren_Click()
' call progstart
' Call DokPfadKorr(Me)
' Call ProgEnde
'End Sub

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

Private Sub MDIForm_Terminate()
 End
End Sub ' MDIForm_Terminate

Private Sub Sonstige_Click()
 snst.Show
End Sub ' Sonstige_Click

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

Private Sub ÜbertrageCd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

'Private Sub VermutlichFehlendeGestationsdiabetesVerschlüsselung_Click() ' vermutlich fehlende Gestationsdiabetes-Verschlüsselung
' Dim rs As New ADODB.Recordset, sql$
' sql = "SELECT f.pat_id, CONCAT(n.nachname,',',n.vorname) Name FROM `aktfv` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND diagsicherheit <> 'A' AND (icd REGEXP '^E1[0-4]' OR (icd LIKE 'O24%' AND obdauer <> 0 AND f.fid = d.fid)) LEFT JOIN `eintraege` e ON f.pat_id = e.pat_id AND inhalt LIKE '%gestat%' WHERE ISNULL(icd) AND NOT ISNULL(inhalt) GROUP BY pat_id"
' myFrag rs, sql
' TabAusgeb rs, Me, , , , , , , "FehlendeGest"
'End Sub

Private Sub Zeilen_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Zeilen_KeyDown

Private Sub Ziel_Change()
 dlg.MdB = Me.Ziel
 If Me.obAcc Then
  Call dbv.cnVorb("quelle", "anamnesebogen", "Patientendaten")
 End If
End Sub ' Ziel_Change

Private Sub Ziel_Click()
 dlg.MdB = Me.Ziel
End Sub ' Ziel_Click

Private Sub Ziel_KeyDown(KeyCode As Integer, Shift As Integer)
 dlg.MdB = Me.Ziel
 Call Key(KeyCode, Shift, Me)
End Sub ' Ziel_KeyDown

Private Sub Picture1_KeyDown(KeyCode%, Shift%)
' IF KeyCode = 27 THEN End
 Call Key(KeyCode, Shift, Me)
End Sub ' Picture1_KeyDown

Private Sub Ausgabe_KeyDown(KeyCode As Integer, Shift As Integer)
 Static obCtrl%, pa&, pe&, Zahl$
 Call Key(KeyCode, Shift, Me)
 If KeyCode = 17 Then
  obCtrl = True
 Else
  If obCtrl Then
   If KeyCode = 68 Then ' D
    pe = Me.Ausgabe.SelStart
    Do While Mid$(Ausgabe.text, pe, 2) <> vbCrLf
     pe = pe + 1
     If pe = Len(Ausgabe.text) - 1 Then
      pe = Len(Ausgabe.text)
      Exit Do
     End If
    Loop
    pa = pe
    Do While Mid$(Me.Ausgabe.text, pa, 1) <> " " And pa > 0
     pa = pa - 1
    Loop
    Zahl = Mid$(Me.Ausgabe.text, pa + 1, pe - pa - 1)
    Clipboard.SetText Zahl
    Call doCallDMP(ByVal Zahl)
   End If
  End If
  obCtrl = False
 End If
' Debug.Print "Schlüssel: " & KeyCode & " " & Shift
End Sub ' Ausgabe_Keydown

' in ' in mdiform_unload, Dialog.Form_uload, dialog.doobMyQuelle_Click, dialog.OKButton_Click
Public Function PutEinstAufDB()
 Dim rs As New ADODB.Recordset
#If False Then
  If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
  If Not DBCn Is Nothing Then
   If DBCn.State = 1 Then
    Set rs = New ADODB.Recordset
'    rs.Open "SELECT * FROM `eintragszahlen` WHERE beginn = (SELECT MAX(beginn) FROM `eintragszahlen`)", DBCn, adOpenDynamic, adLockOptimistic ' "SELECT * FROM `eintragszahlen` ORDER BY beginn DESC" soll bei älteren MySQL-Versionen nicht immer ganz funktionieren
    myFrag rs, "SELECT * FROM `eintragszahlen` WHERE beginn = (SELECT MAX(beginn) FROM `eintragszahlen`)"
    rs!TabellenEntleeren = dlg.TabellenEntleeren
    rs!ZurücksetzenLAktDat = dlg.ZurücksetzenLAktDat
    rs!obVglMitLetzterEinlesung = dlg.obVglMitLetzterEinlesung
    rs!Pat_IDBis = IIf(IsNumeric(dlg.Pat_IDBis), dlg.Pat_IDBis, 0)
    rs!Pat_IDBis = IIf(IsNumeric(dlg.Pat_IDBis), dlg.Pat_IDBis, 0)
    rs!AlterTab = dlg.AlterTab
    rs!VorladenFFI = dlg.VorladenFFI
    rs!ÜberTabelle = dlg.ÜberTabelle
    rs!NurInTabelle = dlg.NurInTabelle
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
 #End If ' false
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PutEinstAufDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PutEinstAufDB

' in haupt.hol_reg, dialog.form_load, dialog.doobMyQuelle_Click
Public Function HolEinstvonDB()
 Dim rs As New ADODB.Recordset
 Dim ErrNr&, ErrDes$
 On Error GoTo fehler
 If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
 If Not DBCn Is Nothing Then
  If DBCn.State = 1 Then
   On Error Resume Next
   Set rs = myEFrag("SELECT * FROM `eintragszahlen` WHERE beginn = (SELECT MAX(beginn) FROM `eintragszahlen`)", , , True, ErrNr, ErrDes)
   If ErrNr <> 0 Then
    MsgBox "Fehler bei Tabelle 'Eintragszahlen' in Datenbank '" & DefDB(DBCn) & "'" ' DBCn.Properties("Data Source Name")
    Exit Function
   End If
   On Error GoTo fehler
   dlg.TabellenEntleeren = Abs(HolEinstFeld(rs, "TabellenEntleeren", 0, "BIT"))
   dlg.TabellenEntleeren = 0
   dlg.ZurücksetzenLAktDat = Abs(HolEinstFeld(rs, "ZurücksetzenLAktDat", 0, "BIT"))
   dlg.obVglMitLetzterEinlesung = Abs(HolEinstFeld(rs, "obVglMitLetzterEinlesung", 0, "BIT"))
   dlg.ZurücksetzenLAktDat = 0
'   dlg.Pat_IDVon = HolEinstFeld(rs, "Pat_IDvon", 0, "TEXT", 6) ' auskommentiert 20.8.23
'   dlg.Pat_IDBis = HolEinstFeld(rs, "Pat_IDbis", 0, "TEXT", 6)
   dlg.AlterTab = Abs(HolEinstFeld(rs, "AlterTab", -1, "BIT"))
   dlg.VorladenFFI = Abs(HolEinstFeld(rs, "VorladenFFI", 0, "BIT"))
   dlg.ÜberTabelle = Abs(HolEinstFeld(rs, "ÜberTabelle", 0, "BIT"))
   dlg.NurInTabelle = Abs(HolEinstFeld(rs, "NurInTabelle", 0, "BIT"))
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
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HolEinstVonDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' HolEinstvonDB

' in HolEinstvonDB
Public Function HolEinstFeld(rs As ADODB.Recordset, FName$, Default, ByVal Typ$, Optional lenge&)
 Dim FNr&
 On Error Resume Next
 HolEinstFeld = rs(FName)
 FNr = Err.Number
 If FNr <> 0 Then
'  ON Error GoTo fehler
  rs.Close
  If lies.obMySQL Then
   Select Case Typ
    Case "TEXT": Typ = "VARCHAR": If lenge = 0 Then Typ = "VARCHAR(300)"
    Case "BIT": If lenge = 0 Then Typ = "BIT(1)" '  Typ = "TINYINT": IF Lenge = 0 THEN Typ = "TINYINT(1)"
    Case "DATE", "TIME": Typ = "DATETIME"
    ' identisch verwendbar sind evtl.: "LONGTEXT","DECIMAL","FLOAT","BINARY"
   End Select
  End If
  Call myEFrag("ALTER TABLE `eintragszahlen` ADD COLUMN `" & FName & "` " & Typ & IIf(lenge <> 0, "(" & lenge & ")", "") & IIf(Null, " NULL", ""))
  Set rs = myEFrag("SELECT * FROM `eintragszahlen` WHERE beginn = (SELECT MAX(beginn) FROM `eintragszahlen`)")
  HolEinstFeld = Default
 End If
  Select Case VarType(HolEinstFeld)
   Case Is < 8
   Case Else
    If HolEinstFeld = "" Then HolEinstFeld = Default
  End Select
 On Error Resume Next
 If InStrB(HolEinstFeld, ":\\") > 0 Or InStrB(HolEinstFeld, "\\\\") > 0 Then HolEinstFeld = REPLACE$(HolEinstFeld, "\\", "\")
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HolEinstFeld/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' HolEinstFeld

Private Sub dbv_wCnAendern(CnStr$)
 On Error GoTo fehler
 ConStri = dbv.Constr
 obStart = True
' IF Me.dbv.wCn.ConnectionString = "Provider=MSDASQL.1;" THEN
'  Me.obMySQL = (Me.dbv.wCn.Properties("DBMS Name") = "MySQL")
' Else
  Me.obMySQL = (InStrB(Me.dbv.CnStr, "MySQL") <> 0) '  Me.dbv.wCn.ConnectionString
' END IF
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
 If Me.dbv.obQuelle Then  ' Me.dbv.DaBa = InStrB(Me.dbv.Ü2, "Patienten") <> 0 THEN  ' sonst bekommt DBCn Bedeutungen von Hausärzten usw
  SetDBCn Me.dbv.wCn, Me.dbv.CnStr
 End If
' IF Me.dbv.wCn.ConnectionString = "Provider=MSDASQL.1;" THEN
'  LVobMySQL = (Me.dbv.wCn.Properties("DBMS Name") = "MySQL")
' Else
  LVobMySQL = (InStrB(Me.dbv.CnStr, "MySQL") <> 0)  ' Me.dbv.wCn.ConnectionString
' END IF

 Me.obMySQL = LVobMySQL
 Me.obAcc = Not Me.obMySQL
 obStart = False
 Call Zinit(LVobMySQL)
' Kommentar 15.5.09, da sonst nach Start lEinlesung automatisch geändert wird
' IF DBCn.State <> 0 THEN
'  Call Me.dlg.FrmLEinlesung
' END IF
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in wcn_Aendern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' dbv_wCnAendern

' mdiForm_Load
Private Sub Konstanten()
 Dim Cpt$
 Cpt = LCase$(CptName)
 Select Case Cpt
  Case "pc08", "gerald03"
   uVerz = "c:\u\"
   pVerz = "c:\p\"
   vVerz = "c:\v\"
  Case Else
   Call SetProgV
'   uVerz = "u:\"
'   pVerz = "p:\"
'   vVerz = "v:\"
 End Select
 hVerz$ = uVerz & "TMImport\" ' HochladeVerzeichnis
 aVerz$ = uVerz & "Anamnese\"
 QmdB$ = aVerz & "Quelle.mdb" ' uverz & "Anamnese\Quelle.mdb
 pügVerz$ = pVerz & "Patientenübergreifendes\"
 üVerz$ = uVerz & "TMExport\"
 LabTransPfad$ = "\\anmeldr\BioWinBACKUP"
 StACCDB$ = AnamneseVerZeichnis1 & "quelle.mdb"
 StFxDB$ = uVerz & "FaxeinP.mdb"
 StFtDB$ = uVerz & "FotosinP.mdb"
 StOffDB$ = uVerz & "office.mdb"
 KVÄDatei1$ = AnamneseVerZeichnis1 + "KV-Ärzte neu.mdb"
 BriefZiel$ = pVerz
 AutoBriefZiel$ = BriefZiel & "unkorrigiert\"
 AutoBriefProtok$ = pügVerz & "zufaxen.txt"
' plzVerz$ = plzVz
 FaxSendDatei$ = uVerz & "test1.txt"
 DMPVorlage$ = uVerz & "vorlagen\DMP-Vorlage gemein.dot"
End Sub ' Konstanten

Private Sub mdiForm_Load()
  userprof = Environ("userprofile")
  Dim lab$
  On Error Resume Next
  Me.Caption = "Patientendaten Diabetespraxis Dachau, Programm: " & App.path & "\" & App.EXEName & ".exe"
  Me.Caption = Me.Caption & ", erstellt: " & FSO.GetFile(App.path & "\" & App.EXEName & ".exe").DateLastModified
  Dim diff#
  On Error GoTo fehler
  Call WD
  Call Konstanten
  Dim a#, b#
  a = 0
  b = 50000
  ' passable Farben zwischen &H00C00000& und &H00FFFFFF&; 4194303 ist die Differenz
 ' Me.Ausgabe.BackColor = &HCC0000 + (((App.Revision + a) * b) Mod 4194303)
  Dim z0#, z1#, z2#, z3#, z4#
  z0 = 1 / App.Revision
  z1 = z0 * 10000000
  z1 = z1 - Int(z1)
  z2 = z1 * 100
  z2 = z2 - Int(z2)
  z3 = z2 * 100
  z3 = z3 - Int(z3)
'  Me.Ausgabe.BackColor = &HA0FFFF
  Me.Ausgabe.BackColor = RGB(160 + z1 * 95, 160 + z2 * 95, 160 + z3 * 95)
 imAufbauLese = True
' SET qdb = New QuelleDBC
 Set dbv = New DBVerb
 Set lies = Me
' SET dlg = New Dialog
 Set dlg.hlese = Me
' SET snst = New Sonstige
 Set snst.hlese = Me
 Call HolReg(Me)
 Call HolKRein(Me)
 Verspätung = vgbVerspätung
 Call holFrist
' Dim rs As New ADODB.Recordset
' myFrag rs, "SHOW create FUNCTION qanf"
' IF Not rs.EOF THEN
'  Dim Text$
'  Text = rs.Fields(2)
'  Dim p1%, p2%
'  Dim versp$
'  Const interval$ = "interval"
'  p1 = InStr(Text, interval)
'  IF p1 > 0 THEN
'   p2 = InStr(p1, Text, "day")
'   p1 = p1 + Len(interval)
'   Verspätung = Mid$(Text, p1 + 1, p2 - p1 - 2)
'  END IF
' END IF
 Me.Zeilen = 0
 Me.Bytes = 0
 With Me.CommonDialogLese
  .DialogTitle = "Einzulesende Datei"
  .initDir = üVerz ' uverz & "tmexport"
  .Filename = "*.BDT"
  .Orientation = cdlLandscape
 .flags = 0
' .Flags = .Flags OR FileOpenConstants.cdlOFNExplorer
 .flags = .flags Or FileOpenConstants.cdlOFNHideReadOnly ' Schreibgeschützt-Checkbox entfernen
 .flags = .flags Or FileOpenConstants.cdlOFNLongNames ' '   Lange Dateinamen erlauben (nur sinnvoll bei Nicht-Win95-Design)
 .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
 .flags = .flags Or FileOpenConstants.cdlOFNPathMustExist
 End With
 
 Dim Verbi As New MCn, rAf&
 Call Verbi.init(DBCnS)
 Dim maxpid&

'' Verbi.Cn.BeginTrans
' Verbi.führaus ("truncate test1")
' Verbi.führaus ("insert into test1(a,b) values(62,10)")
'' Verbi.Cn.CommitTrans
' Verbi.führaus ("insert into test1(a,b) values(63,10)")
' Verbi.Begin
' Verbi.führaus ("insert into test1(a,b) values(74,10)")
' Verbi.führaus ("delete FROM faelle WHERE pat_id=67789")
' Verbi.Commit
' Set Verbi = New MCn
' Call Verbi.init(DBCnS)
 
' DBCn.Execute ("insert into test1(a,b) values(35,10)")
'' DBCn.BeginTrans
' DBCn.Execute ("insert into test1(a,b) values(36,11)")
'' DBCn.CommitTrans
' DBCn.Execute ("start transaction")
' DBCn.Execute ("insert into test1(a,b) values(37,12)")
' DBCn.Execute ("start transaction")
' DBCn.Execute ("insert into test1(a,b) values(38,13)")
' DBCn.Execute ("commit")
' DBCn.Execute ("start transaction")
' DBCn.Execute ("insert into test1(a,b) values(39,14)")
' DBCn.Execute ("start transaction")
' DBCn.Execute ("insert into test1(a,b) values(51,15)")
' DBCn.Execute ("commit")
' DBCn.Close
 
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
  diff = CDate(Me.dlg.BDTAend) - CDate(Me.dlg.LDateiAend)
  If diff > 0 Then
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
  diff = CDate(Me.dlg.BDTAend) - CDate(Me.dlg.LDateiAend)
  If diff > 0 Then
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
  diff = CDate(Me.dlg.BDTAend) - CDate(Me.dlg.LDateiAend)
  If diff > 0 Then
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
  diff = CDate(Me.dlg.BDTAend) - CDate(Me.dlg.LDateiAend)
  If diff > 0 Then
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
 aktOSV = GetOSVersion()
 If aktOSV < win_vista Then Pausenlänge = 100 Else Pausenlänge = 200
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MDIForm_Load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' MDIForm_Load()

' in MDIForm_Load
Public Function AbbrechDisable(frm As Lese)
 Dim i%
 For i = 0 To frm.Controls.COUNT - 1
  If frm.Controls(i).name = "Abbrechen" Then
   frm.Controls(i).Enabled = False
   Exit For
  End If
 Next i
End Function ' AbbrechDisable

Public Function ProgStart()
 If LenB(DBCnS) = 0 Then ' DBCn.ConnectionString
  syscmd 4, "ProgStart"
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

' ProgStart
Private Sub BeendenBlend(obTrue%)
 Dim i&
 For i = 1 To Me.Controls.COUNT
  If Me.Controls(i).name = "Beenden" Then
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
End Sub ' MDIForm_Resize

Private Sub mdiform_unload(Cancel As Integer) ' geht nur beim Anklicken des Kreuzes oben
 Call PutEinstAufDB
 Call PutReg(Me)
' Call Me.dbv.RegSpeichern ' Kommentar 25.9.09
 End
End Sub ' MDIForm_Unload

#If False Then
Public Function ConstrFestleg(ByVal art As ConDtb, Optional hlese As Lese)   ' dlg ist für art= 0 und 1 nötig
 On Error GoTo fehler
'ConStr$ = "DRIVER={MySQL ODBC 3.51 Driver};server=" & LiName & ";uid=praxis;pwd=***REMOVED***;option=" & opti
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
     art = q2Dtb
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
'  MdB = vns
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ConstrFestleg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ConstrFestleg
#End If

Public Function Ausgeb(text$, obDauer%, Optional obdebug%)
 Me.Ausgabe = text & vbCrLf & altAusgabe
 If obDauer <> 0 Then
  altAusgabe = Me.Ausgabe
 End If
 If InStrB(text, "READ-COMMITTED") <> 0 Then
  MsgBox "Beinahe-Stop in Ausgeb:" & vbCrLf & "instrb(text, 'READ-COMMITTED') <> 0" & vbCrLf & "Text: " & text
 End If
 If obdebug <> 0 Then Debug.Print text
 DoEvents
End Function ' Ausgeb

' in ImportiereLabor.LaborDirektImport, Dialog.LaborpfadBez, Dialog.Einlesen
Public Function LaborPfadDialog(LPBeisp$) As String
' IF Me.obBDT THEN
  With Me.CommonDialogLese
   .Filename = "*.ld?;*.dat"
   .DialogTitle = "Einzulesende Labordateien"
   If FSO.FileExists(LPBeisp) Then
    .initDir = FSO.GetParentFolderName(FSO.GetFile(LPBeisp).path)
   Else
    .initDir = LabTransPfad ' uverz & "tmexport"
   End If
   .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
   .ShowOpen
   If FSO.FileExists(.Filename) Then
    LaborPfadDialog = FSO.GetParentFolderName(.Filename)
   End If
  End With
' END IF
End Function ' LaborPfadDialog

