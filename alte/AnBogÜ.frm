VERSION 5.00
Begin VB.Form AnBog 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "AnBog"
   ClientHeight    =   13395
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   17910
   KeyPreview      =   -1  'True
   LinkTopic       =   "AnBog"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13395
   ScaleWidth      =   17910
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Pat_ID"
      Height          =   238
      Index           =   1
      Left            =   113
      TabIndex        =   0
      Top             =   57
      Width           =   516
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Vorname"
      Height          =   238
      Index           =   2
      Left            =   2834
      TabIndex        =   2
      Top             =   56
      Width           =   2040
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "GebDat"
      Height          =   238
      Index           =   3
      Left            =   850
      TabIndex        =   3
      Top             =   340
      Width           =   1017
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H000080FF&
      DataField       =   "Diabetestyp"
      Height          =   285
      Index           =   4
      Left            =   114
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   7
      Top             =   879
      Width           =   1140
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H000080FF&
      DataField       =   "Diabetes seit"
      Height          =   285
      Index           =   5
      Left            =   1299
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   8
      Top             =   879
      Width           =   1125
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H000080FF&
      DataField       =   "Tabletten seit"
      Height          =   285
      Index           =   6
      Left            =   2484
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   9
      Top             =   879
      Width           =   1140
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H000080FF&
      DataField       =   "Insulin seit"
      Height          =   285
      Index           =   7
      Left            =   3669
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      Top             =   879
      Width           =   1023
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Grund für Vorstellung"
      Height          =   645
      Index           =   8
      Left            =   114
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   11
      Top             =   1404
      Width           =   7638
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Familienanamnese"
      Height          =   645
      Index           =   9
      Left            =   114
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   12
      Top             =   2289
      Width           =   7638
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H009D627B&
      DataField       =   "Größe"
      Height          =   238
      Index           =   10
      Left            =   794
      TabIndex        =   13
      Top             =   3005
      Width           =   840
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H009D627B&
      DataField       =   "Gewicht"
      Height          =   238
      Index           =   11
      Left            =   2438
      TabIndex        =   14
      Top             =   3005
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H009D627B&
      DataField       =   "Tendenz"
      Height          =   238
      Index           =   12
      Left            =   4081
      TabIndex        =   15
      Top             =   3004
      Width           =   450
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 1"
      Height          =   256
      Index           =   13
      Left            =   1019
      TabIndex        =   16
      Top             =   3344
      Width           =   3633
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 1 Menge"
      Height          =   256
      Index           =   14
      Left            =   5385
      TabIndex        =   17
      Top             =   3346
      Width           =   2490
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 2"
      Height          =   256
      Index           =   15
      Left            =   737
      TabIndex        =   18
      Top             =   3614
      Width           =   3915
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 2 Menge"
      Height          =   256
      Index           =   16
      Left            =   5385
      TabIndex        =   19
      Top             =   3616
      Width           =   2478
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 3"
      Height          =   256
      Index           =   17
      Left            =   737
      TabIndex        =   20
      Top             =   3883
      Width           =   3915
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 3 Menge"
      Height          =   256
      Index           =   18
      Left            =   5385
      TabIndex        =   21
      Top             =   3885
      Width           =   2490
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 4"
      Height          =   256
      Index           =   19
      Left            =   737
      TabIndex        =   22
      Top             =   4181
      Width           =   3018
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 4 Menge"
      Height          =   256
      Index           =   20
      Left            =   5385
      TabIndex        =   23
      Top             =   4183
      Width           =   2490
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Insulinpumpe"
      Height          =   165
      Index           =   1
      Left            =   1077
      TabIndex        =   24
      Top             =   4592
      Width           =   180
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Insulinpumpe seit"
      Height          =   285
      Index           =   21
      Left            =   1700
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   25
      Top             =   4535
      Width           =   1290
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Insulinpumpe Marke"
      Height          =   285
      Index           =   22
      Left            =   3685
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   26
      Top             =   4479
      Width           =   4188
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten gesamt"
      Height          =   244
      Index           =   23
      Left            =   907
      TabIndex        =   27
      Top             =   4876
      Width           =   585
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten früh"
      Height          =   244
      Index           =   24
      Left            =   1927
      TabIndex        =   28
      Top             =   4875
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten ZM früh"
      Height          =   244
      Index           =   25
      Left            =   3117
      TabIndex        =   29
      Top             =   4875
      Width           =   513
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten mittags"
      Height          =   244
      Index           =   26
      Left            =   3912
      TabIndex        =   30
      Top             =   4875
      Width           =   750
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten nachmittags"
      Height          =   244
      Index           =   27
      Left            =   4989
      TabIndex        =   31
      Top             =   4875
      Width           =   630
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten abends"
      Height          =   244
      Index           =   28
      Left            =   6009
      TabIndex        =   32
      Top             =   4875
      Width           =   858
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten nachts"
      Height          =   244
      Index           =   29
      Left            =   7257
      TabIndex        =   33
      Top             =   4762
      Width           =   615
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit früh"
      Height          =   244
      Index           =   30
      Left            =   1927
      TabIndex        =   34
      Top             =   5160
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit mittags"
      Height          =   244
      Index           =   31
      Left            =   3968
      TabIndex        =   36
      Top             =   5160
      Width           =   690
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit abends"
      Height          =   244
      Index           =   32
      Left            =   6009
      TabIndex        =   38
      Top             =   5160
      Width           =   870
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit spät"
      Height          =   244
      Index           =   33
      Left            =   7256
      TabIndex        =   39
      Top             =   5047
      Width           =   630
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Spritz-Eß-Abstand früh"
      Height          =   244
      Index           =   34
      Left            =   1927
      TabIndex        =   40
      Top             =   5443
      Width           =   783
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Spritz-Eß-Abstand mittags"
      Height          =   244
      Index           =   35
      Left            =   3968
      TabIndex        =   41
      Top             =   5443
      Width           =   720
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Spritz-Eß-Abstand abends"
      Height          =   244
      Index           =   36
      Left            =   6066
      TabIndex        =   42
      Top             =   5443
      Width           =   810
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Spritzstelle früh"
      Height          =   244
      Index           =   37
      Left            =   1927
      TabIndex        =   43
      Top             =   5727
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Spritzstelle mittags"
      Height          =   244
      Index           =   38
      Left            =   3968
      TabIndex        =   44
      Top             =   5727
      Width           =   723
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Spritzstelle abends"
      Height          =   244
      Index           =   39
      Left            =   6066
      TabIndex        =   45
      Top             =   5727
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Spritzstelle nachts"
      Height          =   244
      Index           =   40
      Left            =   7256
      TabIndex        =   46
      Top             =   5614
      Width           =   705
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Jahr letzte Diabetesschulung"
      Height          =   285
      Index           =   41
      Left            =   2211
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   47
      Top             =   6122
      Width           =   2115
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Ort Schulung"
      Height          =   285
      Index           =   42
      Left            =   5102
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   48
      Top             =   6122
      Width           =   2778
   End
   Begin VB.TextBox vTextB 
      DataField       =   "letztes HbA1c"
      Height          =   244
      Index           =   43
      Left            =   1077
      TabIndex        =   49
      Top             =   6405
      Width           =   1065
   End
   Begin VB.TextBox vTextB 
      DataField       =   "gemessen am"
      Height          =   244
      Index           =   44
      Left            =   2769
      TabIndex        =   50
      Top             =   6405
      Width           =   1545
   End
   Begin VB.TextBox vTextB 
      DataField       =   "vorherige Werte"
      Height          =   244
      Index           =   45
      Left            =   5555
      TabIndex        =   51
      Top             =   6405
      Width           =   2310
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZMessungen pW"
      Height          =   244
      Index           =   46
      Left            =   850
      TabIndex        =   55
      Top             =   7003
      Width           =   1365
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZMessungen pW ndE"
      Height          =   244
      Index           =   47
      Left            =   3174
      TabIndex        =   56
      Top             =   7003
      Width           =   858
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZMessungen p W nachts"
      Height          =   244
      Index           =   48
      Left            =   5272
      TabIndex        =   57
      Top             =   7003
      Width           =   915
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Aufschreiben"
      Height          =   244
      Index           =   49
      Left            =   6236
      TabIndex        =   54
      Top             =   6720
      Width           =   570
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZWerte v d Essen"
      Height          =   244
      Index           =   50
      Left            =   1190
      TabIndex        =   58
      Top             =   7371
      Width           =   1590
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZWerte n d Essen"
      Height          =   244
      Index           =   51
      Left            =   3514
      TabIndex        =   59
      Top             =   7370
      Width           =   1488
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Unterzucker pM"
      Height          =   244
      Index           =   52
      Left            =   680
      TabIndex        =   61
      Top             =   7653
      Width           =   510
   End
   Begin VB.TextBox vTextB 
      DataField       =   "UZ rechtzeitig"
      Height          =   244
      Index           =   53
      Left            =   1984
      TabIndex        =   62
      Top             =   7653
      Width           =   330
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Fremde Hilfe pa"
      Height          =   244
      Index           =   54
      Left            =   3004
      TabIndex        =   63
      Top             =   7653
      Width           =   570
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Bewußtlos pa"
      Height          =   244
      Index           =   55
      Left            =   4308
      TabIndex        =   64
      Top             =   7653
      Width           =   450
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Keto pa"
      Height          =   244
      Index           =   56
      Left            =   5442
      TabIndex        =   65
      Top             =   7653
      Width           =   453
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZgr300 pM"
      Height          =   244
      Index           =   57
      Left            =   6746
      TabIndex        =   66
      Top             =   7653
      Width           =   960
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Bluthochdruck"
      Height          =   285
      Index           =   58
      Left            =   56
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   69
      Top             =   8687
      Width           =   735
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "BHD seit"
      Height          =   285
      Index           =   59
      Left            =   850
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   70
      Top             =   8687
      Width           =   1125
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "BHD beh mit"
      Height          =   285
      Index           =   60
      Left            =   2035
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   71
      Top             =   8687
      Width           =   3000
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Blutdruckwerte"
      Height          =   285
      Index           =   61
      Left            =   6039
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   73
      Top             =   8687
      Width           =   1848
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BDselbst"
      Height          =   285
      Index           =   62
      Left            =   5102
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   72
      Top             =   8687
      Width           =   900
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Schwanger"
      Height          =   244
      Index           =   63
      Left            =   1077
      TabIndex        =   67
      Top             =   8050
      Width           =   630
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Schwanger seit"
      Height          =   244
      Index           =   64
      Left            =   3004
      TabIndex        =   68
      Top             =   8050
      Width           =   1185
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Augensp zuletzt"
      Height          =   285
      Index           =   65
      Left            =   56
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   74
      Top             =   9310
      Width           =   1140
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Augensp Befund"
      Height          =   285
      Index           =   66
      Left            =   1256
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   75
      Top             =   9310
      Width           =   3453
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Netzhaut gelasert"
      Height          =   285
      Index           =   67
      Left            =   4762
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   76
      Top             =   9310
      Width           =   1260
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Sehminderung unbehebbar"
      Height          =   285
      Index           =   68
      Left            =   6055
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   77
      Top             =   9310
      Width           =   1800
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Diabet Nierenschaden"
      Height          =   285
      Index           =   69
      Left            =   56
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   78
      Top             =   9933
      Width           =   900
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Albumin zuletzt"
      Height          =   285
      Index           =   70
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   79
      Top             =   9933
      Width           =   1125
   End
   Begin VB.TextBox vTextB 
      DataField       =   "erhöht?"
      Height          =   285
      Index           =   71
      Left            =   2205
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   80
      Top             =   9933
      Width           =   1185
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Dialyse"
      Height          =   285
      Index           =   2
      Left            =   3390
      TabIndex        =   81
      Top             =   9933
      Width           =   288
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Dialyse seit"
      Height          =   285
      Index           =   72
      Left            =   3664
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   82
      Top             =   9933
      Width           =   1155
   End
   Begin VB.TextBox vTextB 
      DataField       =   "andere Nierenerkrankung"
      Height          =   285
      Index           =   73
      Left            =   4822
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   83
      Top             =   9933
      Width           =   3030
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Herzkrankheit"
      Height          =   285
      Index           =   74
      Left            =   56
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   84
      Top             =   10500
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Angina pectoris"
      Height          =   285
      Index           =   75
      Left            =   907
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   85
      Top             =   10500
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Herzinfarkt"
      Height          =   285
      Index           =   76
      Left            =   1757
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   86
      Top             =   10500
      Width           =   513
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Herzinfarkt wann"
      Height          =   285
      Index           =   77
      Left            =   2324
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   87
      Top             =   10500
      Width           =   675
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PTCA oder Stent"
      Height          =   285
      Index           =   78
      Left            =   3061
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   88
      Top             =   10500
      Width           =   510
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Bypass kardial"
      Height          =   285
      Index           =   3
      Left            =   3120
      TabIndex        =   89
      Top             =   10800
      Width           =   1095
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Bypass wann"
      Height          =   285
      Index           =   79
      Left            =   3798
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   90
      Top             =   10500
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Herzschwäche"
      Height          =   285
      Index           =   80
      Left            =   4648
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   91
      Top             =   10487
      Width           =   963
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Herzkrankheit Beschreibung"
      Height          =   285
      Index           =   81
      Left            =   5585
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   92
      Top             =   10500
      Width           =   2265
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Hirndurchblutungsstörung"
      Height          =   285
      Index           =   82
      Left            =   7880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   93
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Schlaganfall"
      Height          =   285
      Index           =   83
      Left            =   8900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   94
      Top             =   240
      Width           =   2310
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Beindurchblutungsstörung"
      Height          =   285
      Index           =   84
      Left            =   11225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   95
      Top             =   240
      Width           =   858
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Schaufensterkrankheit"
      Height          =   285
      Index           =   85
      Left            =   12132
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   96
      Top             =   240
      Width           =   900
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Bypaß peripher"
      Height          =   285
      Index           =   4
      Left            =   13096
      TabIndex        =   97
      Top             =   240
      Width           =   270
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Geschwür"
      Height          =   285
      Index           =   86
      Left            =   13322
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   98
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Amputation"
      Height          =   285
      Index           =   87
      Left            =   14222
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   99
      Top             =   240
      Width           =   2433
   End
   Begin VB.TextBox vTextB 
      DataField       =   "pAVK Beschreibung"
      Height          =   645
      Index           =   88
      Left            =   7880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   100
      Top             =   750
      Width           =   7638
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Ameisenlaufen"
      Height          =   285
      Index           =   89
      Left            =   7880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   101
      Top             =   1635
      Width           =   465
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Ameisen Ausmaß"
      Height          =   285
      Index           =   90
      Left            =   8390
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   102
      Top             =   1657
      Width           =   2100
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Druckstellen"
      Height          =   285
      Index           =   91
      Left            =   10530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   103
      Top             =   1657
      Width           =   525
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Verformungen"
      Height          =   285
      Index           =   92
      Left            =   11111
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   104
      Top             =   1657
      Width           =   510
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Verformungen Beschreibung"
      Height          =   285
      Index           =   93
      Left            =   11678
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   105
      Top             =   1657
      Width           =   1413
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Neue Fußkomplikationen"
      Height          =   285
      Index           =   94
      Left            =   7937
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   109
      Top             =   2281
      Width           =   2265
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Entleerungsstörungen Magen"
      Height          =   285
      Index           =   95
      Left            =   10266
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   110
      Top             =   2281
      Width           =   1590
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Entleerungsstörungen Harnblase"
      Height          =   285
      Index           =   96
      Left            =   11853
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   111
      Top             =   2281
      Width           =   1155
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Schwindel Aufstehen"
      Height          =   285
      Index           =   97
      Left            =   13044
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   112
      Top             =   2281
      Width           =   1248
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Folgeerkrankungen Haut"
      Height          =   285
      Index           =   98
      Left            =   7937
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   113
      Top             =   2848
      Width           =   1830
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H009595FF&
      DataField       =   "Bewegungseinschränkungen"
      Height          =   285
      Index           =   99
      Left            =   9808
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   114
      Top             =   2848
      Width           =   2310
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Sexualstörung"
      Height          =   285
      Index           =   100
      Left            =   12188
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   115
      Top             =   2848
      Width           =   1080
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Sexualstörung seit"
      Height          =   285
      Index           =   101
      Left            =   13268
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   116
      Top             =   2848
      Width           =   1938
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Weitere Anamnese"
      Height          =   645
      Index           =   102
      Left            =   7937
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   117
      Top             =   3415
      Width           =   7638
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Alkohol"
      Height          =   285
      Index           =   103
      Left            =   11055
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   119
      Top             =   4082
      Width           =   4485
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Tabak"
      Height          =   285
      Index           =   104
      Left            =   8617
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   118
      Top             =   4082
      Width           =   1698
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Weitere Medikation"
      Height          =   645
      Index           =   105
      Left            =   7937
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   120
      Top             =   4605
      Width           =   7638
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Liphypertrophien Abdomen"
      Height          =   285
      Index           =   106
      Left            =   7937
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   121
      Top             =   5512
      Width           =   2310
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Liphypertrophien Beine"
      Height          =   285
      Index           =   107
      Left            =   10247
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   122
      Top             =   5512
      Width           =   2310
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Liphypertrophien Arme"
      Height          =   285
      Index           =   108
      Left            =   12557
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   123
      Top             =   5512
      Width           =   3018
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Hyperkeratosen"
      Height          =   285
      Index           =   109
      Left            =   7993
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   125
      Top             =   6419
      Width           =   3285
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Ulcera"
      Height          =   285
      Index           =   110
      Left            =   11338
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   126
      Top             =   6419
      Width           =   1485
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Kraft Zehenheber"
      Height          =   285
      Index           =   111
      Left            =   12869
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   127
      Top             =   6419
      Width           =   915
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Kraft Zehenbeuger"
      Height          =   285
      Index           =   112
      Left            =   13833
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   128
      Top             =   6419
      Width           =   843
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Kraft Knie"
      Height          =   285
      Index           =   113
      Left            =   14740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   129
      Top             =   6419
      Width           =   855
   End
   Begin VB.TextBox vTextB 
      DataField       =   "ASR"
      Height          =   285
      Index           =   114
      Left            =   7993
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   130
      Top             =   6986
      Width           =   1080
   End
   Begin VB.TextBox vTextB 
      DataField       =   "PSR"
      Height          =   285
      Index           =   115
      Left            =   9127
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   131
      Top             =   6986
      Width           =   990
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Oberflächensensibilität"
      Height          =   285
      Index           =   116
      Left            =   10148
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   132
      Top             =   6986
      Width           =   1638
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Monofilamenttest"
      Height          =   285
      Index           =   117
      Left            =   11848
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   133
      Top             =   6986
      Width           =   1560
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Kalt-Warm"
      Height          =   285
      Index           =   118
      Left            =   13408
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   134
      Top             =   6986
      Width           =   1560
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Vibration IK"
      Height          =   285
      Index           =   119
      Left            =   7993
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   135
      Top             =   7553
      Width           =   1560
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Vibration Großzehe"
      Height          =   285
      Index           =   120
      Left            =   9581
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   136
      Top             =   7553
      Width           =   1560
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Puls Leiste"
      Height          =   285
      Index           =   121
      Left            =   11225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   137
      Top             =   7553
      Width           =   1023
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Puls Kniekehle"
      Height          =   285
      Index           =   122
      Left            =   12302
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   138
      Top             =   7553
      Width           =   900
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Puls Atp"
      Height          =   285
      Index           =   123
      Left            =   13266
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   139
      Top             =   7553
      Width           =   1080
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Puls Adp"
      Height          =   285
      Index           =   124
      Left            =   14400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   140
      Top             =   7553
      Width           =   1125
   End
   Begin VB.TextBox vTextB 
      DataField       =   "RR"
      Height          =   285
      Index           =   125
      Left            =   8447
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   141
      Top             =   7937
      Width           =   7128
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Herz"
      Height          =   238
      Index           =   126
      Left            =   8503
      TabIndex        =   143
      Top             =   8505
      Width           =   2655
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Lunge"
      Height          =   238
      Index           =   127
      Left            =   13548
      TabIndex        =   145
      Top             =   8505
      Width           =   2040
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Bauch"
      Height          =   238
      Index           =   128
      Left            =   8616
      TabIndex        =   146
      Top             =   8755
      Width           =   1968
   End
   Begin VB.TextBox vTextB 
      DataField       =   "WS"
      Height          =   238
      Index           =   129
      Left            =   10998
      TabIndex        =   147
      Top             =   8755
      Width           =   1035
   End
   Begin VB.TextBox vTextB 
      DataField       =   "NL"
      Height          =   238
      Index           =   130
      Left            =   12302
      TabIndex        =   148
      Top             =   8755
      Width           =   1080
   End
   Begin VB.TextBox vTextB 
      DataField       =   "SD"
      Height          =   238
      Index           =   131
      Left            =   8220
      TabIndex        =   149
      Top             =   8993
      Width           =   1200
   End
   Begin VB.TextBox vTextB 
      DataField       =   "NNH"
      Height          =   238
      Index           =   132
      Left            =   9874
      TabIndex        =   150
      Top             =   8993
      Width           =   1290
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Zähne"
      Height          =   238
      Index           =   133
      Left            =   11792
      TabIndex        =   151
      Top             =   8993
      Width           =   1023
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Mundhöhle"
      Height          =   238
      Index           =   134
      Left            =   8957
      TabIndex        =   152
      Top             =   9243
      Width           =   1260
   End
   Begin VB.TextBox vTextB 
      DataField       =   "LK"
      Height          =   238
      Index           =   135
      Left            =   8390
      TabIndex        =   154
      Top             =   9581
      Width           =   1380
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Neuro sonst"
      Height          =   238
      Index           =   136
      Left            =   8731
      TabIndex        =   156
      Top             =   10035
      Width           =   2328
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Weitere Befunde"
      Height          =   540
      Index           =   137
      Left            =   11162
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   155
      Top             =   9751
      Width           =   4428
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Schulung"
      Height          =   256
      Index           =   138
      Left            =   13889
      TabIndex        =   158
      Top             =   10320
      Width           =   855
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DMP"
      Height          =   285
      Index           =   139
      Left            =   7993
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   159
      Top             =   10828
      Width           =   900
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Ausgabe"
      Height          =   291
      Index           =   1
      Left            =   15583
      TabIndex        =   160
      Top             =   566
      Width           =   1551
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "HANr"
      Height          =   240
      Index           =   140
      Left            =   10148
      TabIndex        =   161
      Top             =   10885
      Width           =   1176
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Beinbefund"
      Height          =   240
      Index           =   141
      Left            =   9015
      TabIndex        =   124
      Top             =   5839
      Width           =   6576
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H009D627B&
      DataField       =   "=IIf([Größe]=0,"",[Gewicht]/[Größe]/[Größe]*IIf([Größe]>3,10000,1))"
      Height          =   238
      Index           =   142
      Left            =   5272
      TabIndex        =   162
      Top             =   3005
      Width           =   636
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "letzte Änderung"
      Height          =   227
      Index           =   143
      Left            =   17064
      TabIndex        =   163
      Top             =   0
      Width           =   1349
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Fußpflege"
      Height          =   240
      Index           =   144
      Left            =   13152
      TabIndex        =   106
      Top             =   1701
      Width           =   696
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Podologie"
      Height          =   240
      Index           =   145
      Left            =   13889
      TabIndex        =   107
      Top             =   1701
      Width           =   696
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Einlagen"
      Height          =   240
      Index           =   146
      Left            =   14626
      TabIndex        =   108
      Top             =   1701
      Width           =   906
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   4591
      Index           =   147
      Left            =   15647
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   164
      Top             =   850
      Width           =   13215
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   227
      Index           =   148
      Left            =   10374
      TabIndex        =   165
      Top             =   10601
      Width           =   8148
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "[Vorgestellt]"
      Height          =   238
      Index           =   149
      Left            =   2834
      TabIndex        =   4
      Top             =   340
      Width           =   1011
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "RRTurboMed"
      Height          =   240
      Index           =   150
      Left            =   8447
      TabIndex        =   142
      Top             =   8220
      Width           =   7131
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Versicherung"
      Height          =   240
      Index           =   151
      Left            =   9070
      TabIndex        =   157
      Top             =   10318
      Width           =   3846
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Titel"
      Height          =   240
      Index           =   152
      Left            =   4932
      TabIndex        =   166
      Top             =   56
      Width           =   681
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Anrede"
      Height          =   240
      Index           =   153
      Left            =   5669
      TabIndex        =   167
      Top             =   56
      Width           =   801
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FF0080&
      DataField       =   "Versicherungsart"
      Height          =   240
      Index           =   154
      Left            =   6576
      TabIndex        =   168
      Top             =   56
      Width           =   336
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit vormittags"
      Height          =   244
      Index           =   155
      Left            =   3118
      TabIndex        =   35
      Top             =   5160
      Width           =   516
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit nachmittags"
      Height          =   244
      Index           =   156
      Left            =   4988
      TabIndex        =   37
      Top             =   5160
      Width           =   681
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZMessungen selbst"
      Height          =   244
      Index           =   157
      Left            =   1870
      TabIndex        =   52
      Top             =   6720
      Width           =   336
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Gerät"
      Height          =   244
      Index           =   158
      Left            =   2890
      TabIndex        =   53
      Top             =   6720
      Width           =   2091
   End
   Begin VB.TextBox vTextB 
      DataField       =   "UZ Tageszeit"
      Height          =   244
      Index           =   159
      Left            =   6236
      TabIndex        =   60
      Top             =   7258
      Width           =   1701
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Nachname"
      Height          =   238
      Index           =   160
      Left            =   680
      TabIndex        =   1
      Top             =   56
      Width           =   2100
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BeinödVen"
      Height          =   510
      Index           =   161
      Left            =   11399
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   153
      Top             =   9243
      Width           =   4191
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Carotiden"
      Height          =   238
      Index           =   162
      Left            =   11632
      TabIndex        =   144
      Top             =   8505
      Width           =   1230
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "DMSchulz"
      Height          =   256
      Index           =   163
      Left            =   15817
      TabIndex        =   169
      Top             =   10320
      Width           =   510
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "RRSchulz"
      Height          =   256
      Index           =   164
      Left            =   17351
      TabIndex        =   170
      Top             =   10318
      Width           =   510
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "DMPhier"
      Height          =   285
      Index           =   165
      Left            =   8957
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   171
      Top             =   10828
      Width           =   900
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Tkz"
      Height          =   225
      Index           =   5
      Left            =   7426
      TabIndex        =   172
      Top             =   56
      Width           =   270
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Ther1"
      Height          =   225
      Index           =   166
      Left            =   5782
      TabIndex        =   6
      Top             =   626
      Width           =   750
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00808080&
      DataField       =   "=[Nachname]+"
      Height          =   239
      Index           =   167
      Left            =   11338
      TabIndex        =   173
      Top             =   4365
      Width           =   4202
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   4599
      Index           =   168
      Left            =   15760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   174
      Top             =   5499
      Width           =   13094
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   227
      Index           =   169
      Left            =   12812
      TabIndex        =   175
      Top             =   10885
      Width           =   1486
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   341
      Index           =   170
      Left            =   7993
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   176
      Top             =   11168
      Width           =   6353
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "An1Aufruf"
      Height          =   276
      Index           =   2
      Left            =   56
      TabIndex        =   177
      Top             =   10885
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "An2Aufruf"
      Height          =   291
      Index           =   3
      Left            =   56
      TabIndex        =   178
      Top             =   11221
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "AnAAufruf"
      Height          =   321
      Index           =   4
      Left            =   3571
      TabIndex        =   179
      Top             =   10840
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "CheckAufruf"
      Height          =   291
      Index           =   5
      Left            =   3571
      TabIndex        =   180
      Top             =   11225
      Width           =   2826
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obAn1eing"
      Height          =   240
      Index           =   6
      Left            =   2891
      TabIndex        =   181
      Top             =   10885
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obAn2eing"
      Height          =   240
      Index           =   7
      Left            =   2891
      TabIndex        =   182
      Top             =   11245
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obAnAeing"
      Height          =   240
      Index           =   8
      Left            =   6462
      TabIndex        =   183
      Top             =   10885
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obCheck"
      Height          =   240
      Index           =   9
      Left            =   6462
      TabIndex        =   184
      Top             =   11225
      Width           =   260
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Labor"
      Height          =   336
      Index           =   6
      Left            =   6776
      TabIndex        =   185
      Top             =   10828
      Width           =   1071
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Brief"
      Height          =   290
      Index           =   7
      Left            =   6803
      TabIndex        =   186
      Top             =   11225
      Width           =   1011
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Dokumente"
      Height          =   336
      Index           =   8
      Left            =   56
      TabIndex        =   187
      Top             =   11565
      Width           =   696
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Briefe"
      Height          =   336
      Index           =   9
      Left            =   1190
      TabIndex        =   188
      Top             =   11565
      Width           =   2031
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   1
      Left            =   15649
      TabIndex        =   189
      Top             =   5511
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   2
      Left            =   15649
      TabIndex        =   190
      Top             =   5707
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   3
      Left            =   15649
      TabIndex        =   191
      Top             =   5903
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   4
      Left            =   15649
      TabIndex        =   192
      Top             =   6098
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   5
      Left            =   15649
      TabIndex        =   193
      Top             =   6294
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   6
      Left            =   15649
      TabIndex        =   194
      Top             =   6490
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   7
      Left            =   15649
      TabIndex        =   195
      Top             =   6686
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   8
      Left            =   15649
      TabIndex        =   196
      Top             =   6882
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   9
      Left            =   15649
      TabIndex        =   197
      Top             =   7078
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   10
      Left            =   15649
      TabIndex        =   198
      Top             =   7274
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   11
      Left            =   15649
      TabIndex        =   199
      Top             =   7470
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   12
      Left            =   15649
      TabIndex        =   200
      Top             =   7666
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   13
      Left            =   15649
      TabIndex        =   201
      Top             =   7862
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   14
      Left            =   15649
      TabIndex        =   202
      Top             =   8058
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   15
      Left            =   15649
      TabIndex        =   203
      Top             =   8254
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   16
      Left            =   15649
      TabIndex        =   204
      Top             =   8450
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   17
      Left            =   15649
      TabIndex        =   205
      Top             =   8646
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   18
      Left            =   15647
      TabIndex        =   206
      Top             =   8842
      Width           =   113
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "DReset"
      Height          =   216
      Index           =   10
      Left            =   17971
      TabIndex        =   207
      Top             =   10318
      Width           =   576
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obBZausgew"
      Height          =   240
      Index           =   10
      Left            =   14400
      TabIndex        =   208
      Top             =   10830
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obOSaufgek"
      Height          =   240
      Index           =   11
      Left            =   14400
      TabIndex        =   210
      Top             =   11287
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obPodAufgek"
      Height          =   240
      Index           =   12
      Left            =   14400
      TabIndex        =   211
      Top             =   11532
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obMBlAusgeh"
      Height          =   240
      Index           =   13
      Left            =   14400
      TabIndex        =   209
      Top             =   11080
      Width           =   260
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "obDMPaufgekl"
      Height          =   238
      Index           =   171
      Left            =   16157
      TabIndex        =   213
      Top             =   11533
      Width           =   405
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "obSchulaufgek"
      Height          =   238
      Index           =   172
      Left            =   16157
      TabIndex        =   212
      Top             =   11281
      Width           =   405
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   19
      Left            =   15647
      TabIndex        =   214
      Top             =   9038
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   20
      Left            =   15647
      TabIndex        =   215
      Top             =   9234
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   21
      Left            =   15647
      TabIndex        =   216
      Top             =   9430
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   22
      Left            =   15647
      TabIndex        =   217
      Top             =   9626
      Width           =   113
   End
   Begin VB.OptionButton vOptionB 
      Height          =   193
      Index           =   23
      Left            =   15647
      TabIndex        =   218
      Top             =   9822
      Width           =   113
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "TherAkt"
      Height          =   238
      Index           =   173
      Left            =   6973
      TabIndex        =   219
      Top             =   623
      Width           =   780
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Versicherungsart"
      Height          =   226
      Index           =   174
      Left            =   7313
      TabIndex        =   220
      Top             =   340
      Width           =   454
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obMedNetz"
      Height          =   240
      Index           =   14
      Left            =   7483
      TabIndex        =   221
      Top             =   907
      Width           =   260
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "BZKurven"
      Height          =   336
      Index           =   11
      Left            =   3344
      TabIndex        =   222
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "DMPAusgeb"
      Height          =   290
      Index           =   12
      Left            =   17177
      TabIndex        =   223
      Top             =   566
      Width           =   1296
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "AugenBefunde"
      Height          =   337
      Index           =   13
      Left            =   5499
      TabIndex        =   224
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "DokDown"
      Height          =   291
      Index           =   14
      Left            =   793
      TabIndex        =   225
      Top             =   11565
      Width           =   336
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   227
      Index           =   175
      Left            =   11735
      TabIndex        =   226
      Top             =   10885
      Width           =   1032
   End
   Begin VB.ComboBox vComboB 
      Height          =   284
      Index           =   1
      Left            =   8569
      TabIndex        =   410
      Top             =   11508
      Width           =   2599
   End
   Begin VB.ComboBox vComboB 
      Height          =   286
      Index           =   2
      Left            =   11791
      TabIndex        =   411
      Top             =   11508
      Width           =   2551
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HA1nicht"
      Height          =   240
      Index           =   15
      Left            =   8333
      TabIndex        =   229
      Top             =   11792
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HA2nicht"
      Height          =   240
      Index           =   16
      Left            =   11451
      TabIndex        =   230
      Top             =   11735
      Width           =   260
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "HA1_Bez"
      Height          =   285
      Index           =   15
      Left            =   7990
      TabIndex        =   231
      Top             =   11508
      Width           =   570
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "HA2_Bez"
      Height          =   270
      Index           =   16
      Left            =   11168
      TabIndex        =   232
      Top             =   11508
      Width           =   615
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   284
      Index           =   176
      Left            =   16440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   233
      Top             =   10885
      Width           =   1984
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   1140
      Index           =   177
      Left            =   4315
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   234
      Top             =   11962
      Width           =   4127
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   1185
      Index           =   178
      Left            =   9139
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   235
      Top             =   11905
      Width           =   4632
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   737
      Index           =   179
      Left            =   14301
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   236
      Top             =   11905
      Width           =   4350
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   1140
      Index           =   180
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   237
      Top             =   11962
      Width           =   3635
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   226
      Index           =   181
      Left            =   4025
      TabIndex        =   5
      Top             =   340
      Width           =   1190
   End
   Begin VB.ComboBox vComboB 
      Height          =   315
      Index           =   3
      Left            =   0
      TabIndex        =   420
      Top             =   13080
      Width           =   3814
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Unausgefülle"
      Height          =   291
      Index           =   17
      Left            =   4080
      TabIndex        =   422
      Top             =   13080
      Width           =   1251
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "alleDatensätze"
      Height          =   291
      Index           =   18
      Left            =   7680
      TabIndex        =   423
      Top             =   13080
      Width           =   1251
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "UnausgefNZ"
      Height          =   291
      Index           =   19
      Left            =   5520
      TabIndex        =   424
      Top             =   13080
      Width           =   1866
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "=currentdb.Name"
      Height          =   227
      Index           =   182
      Left            =   11225
      TabIndex        =   425
      Top             =   0
      Width           =   7257
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Datenbank-Aufruf"
      Height          =   291
      Index           =   20
      Left            =   11225
      TabIndex        =   426
      Top             =   226
      Width           =   7311
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "GebDat"
      Height          =   238
      Index           =   1
      Left            =   113
      TabIndex        =   227
      Top             =   340
      Width           =   687
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Diabetestyp&:"
      Height          =   240
      Index           =   2
      Left            =   114
      TabIndex        =   228
      Top             =   639
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "Diabetes seit"
      Height          =   240
      Index           =   3
      Left            =   1299
      TabIndex        =   238
      Top             =   639
      Width           =   1125
   End
   Begin VB.Label vLab 
      Caption         =   "Tabletten seit"
      Height          =   240
      Index           =   4
      Left            =   2484
      TabIndex        =   239
      Top             =   639
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "Insulin seit"
      Height          =   240
      Index           =   5
      Left            =   3669
      TabIndex        =   240
      Top             =   639
      Width           =   1023
   End
   Begin VB.Label vLab 
      Caption         =   "Grund für Vorstellung"
      Height          =   240
      Index           =   6
      Left            =   114
      TabIndex        =   241
      Top             =   1164
      Width           =   1578
   End
   Begin VB.Label vLab 
      Caption         =   "Familienanamnese"
      Height          =   240
      Index           =   7
      Left            =   114
      TabIndex        =   242
      Top             =   2049
      Width           =   7638
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Gr&öße"
      Height          =   238
      Index           =   8
      Left            =   113
      TabIndex        =   243
      Top             =   3005
      Width           =   570
   End
   Begin VB.Label vLab 
      Caption         =   "Gewicht"
      Height          =   238
      Index           =   9
      Left            =   1701
      TabIndex        =   244
      Top             =   3005
      Width           =   690
   End
   Begin VB.Label vLab 
      Caption         =   "Tendenz"
      Height          =   238
      Index           =   10
      Left            =   3288
      TabIndex        =   245
      Top             =   3005
      Width           =   750
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Diab&.Med. 1"
      Height          =   256
      Index           =   11
      Left            =   56
      TabIndex        =   246
      Top             =   3344
      Width           =   963
   End
   Begin VB.Label vLab 
      Caption         =   "Menge"
      Height          =   256
      Index           =   12
      Left            =   4762
      TabIndex        =   247
      Top             =   3346
      Width           =   615
   End
   Begin VB.Label vLab 
      Caption         =   "2"
      Height          =   256
      Index           =   13
      Left            =   56
      TabIndex        =   248
      Top             =   3614
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "Menge"
      Height          =   256
      Index           =   14
      Left            =   4705
      TabIndex        =   249
      Top             =   3616
      Width           =   663
   End
   Begin VB.Label vLab 
      Caption         =   "3"
      Height          =   256
      Index           =   15
      Left            =   56
      TabIndex        =   250
      Top             =   3883
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "Menge"
      Height          =   256
      Index           =   16
      Left            =   4762
      TabIndex        =   251
      Top             =   3885
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "4"
      Height          =   256
      Index           =   17
      Left            =   56
      TabIndex        =   252
      Top             =   4181
      Width           =   633
   End
   Begin VB.Label vLab 
      Caption         =   "Menge"
      Height          =   256
      Index           =   18
      Left            =   4648
      TabIndex        =   253
      Top             =   4183
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Insulinpumpe"
      Height          =   240
      Index           =   19
      Left            =   56
      TabIndex        =   254
      Top             =   4535
      Width           =   1005
   End
   Begin VB.Label vLab 
      Caption         =   "seit"
      Height          =   240
      Index           =   20
      Left            =   1303
      TabIndex        =   255
      Top             =   4535
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "Marke"
      Height          =   240
      Index           =   21
      Left            =   3061
      TabIndex        =   256
      Top             =   4479
      Width           =   573
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "BE &gesamt"
      Height          =   244
      Index           =   22
      Left            =   0
      TabIndex        =   257
      Top             =   4876
      Width           =   855
   End
   Begin VB.Label vLab 
      Caption         =   "früh"
      Height          =   244
      Index           =   23
      Left            =   1587
      TabIndex        =   258
      Top             =   4875
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "vormittags"
      Height          =   244
      Index           =   24
      Left            =   2777
      TabIndex        =   259
      Top             =   4875
      Width           =   348
   End
   Begin VB.Label vLab 
      Caption         =   "mittags"
      Height          =   244
      Index           =   25
      Left            =   3685
      TabIndex        =   260
      Top             =   4875
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "nachmittags"
      Height          =   244
      Index           =   26
      Left            =   4705
      TabIndex        =   261
      Top             =   4875
      Width           =   285
   End
   Begin VB.Label vLab 
      Caption         =   "abends"
      Height          =   244
      Index           =   27
      Left            =   5669
      TabIndex        =   262
      Top             =   4875
      Width           =   348
   End
   Begin VB.Label vLab 
      Caption         =   "spät"
      Height          =   244
      Index           =   28
      Left            =   6916
      TabIndex        =   263
      Top             =   4762
      Width           =   345
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Essenszeit  fr&üh:"
      Height          =   244
      Index           =   29
      Left            =   632
      TabIndex        =   264
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label vLab 
      Caption         =   "mittags"
      Height          =   244
      Index           =   30
      Left            =   3685
      TabIndex        =   265
      Top             =   5160
      Width           =   285
   End
   Begin VB.Label vLab 
      Caption         =   "abends"
      Height          =   244
      Index           =   31
      Left            =   5669
      TabIndex        =   266
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "spät"
      Height          =   244
      Index           =   32
      Left            =   6916
      TabIndex        =   267
      Top             =   5047
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "Spritz-Eß-Abstand früh:"
      Height          =   244
      Index           =   33
      Left            =   56
      TabIndex        =   268
      Top             =   5442
      Width           =   1818
   End
   Begin VB.Label vLab 
      Caption         =   "mittags"
      Height          =   244
      Index           =   34
      Left            =   3685
      TabIndex        =   269
      Top             =   5443
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "abends"
      Height          =   244
      Index           =   35
      Left            =   5669
      TabIndex        =   270
      Top             =   5443
      Width           =   330
   End
   Begin VB.Label vLab 
      Caption         =   "Spritzstelle früh (B, O, A)"
      Height          =   244
      Index           =   36
      Left            =   56
      TabIndex        =   271
      Top             =   5727
      Width           =   1785
   End
   Begin VB.Label vLab 
      Caption         =   "mittags"
      Height          =   244
      Index           =   37
      Left            =   3685
      TabIndex        =   272
      Top             =   5727
      Width           =   288
   End
   Begin VB.Label vLab 
      Caption         =   "abends"
      Height          =   244
      Index           =   38
      Left            =   5669
      TabIndex        =   273
      Top             =   5727
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "nachts"
      Height          =   244
      Index           =   39
      Left            =   6859
      TabIndex        =   274
      Top             =   5614
      Width           =   405
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Jahr letzte Diabetesschulung"
      Height          =   240
      Index           =   40
      Left            =   56
      TabIndex        =   275
      Top             =   6122
      Width           =   2115
   End
   Begin VB.Label vLab 
      Caption         =   "Ort Schulung"
      Height          =   240
      Index           =   41
      Left            =   4365
      TabIndex        =   276
      Top             =   6122
      Width           =   738
   End
   Begin VB.Label vLab 
      Caption         =   "letztes HbA1c"
      Height          =   244
      Index           =   42
      Left            =   56
      TabIndex        =   277
      Top             =   6405
      Width           =   960
   End
   Begin VB.Label vLab 
      Caption         =   "gem am"
      Height          =   244
      Index           =   43
      Left            =   2097
      TabIndex        =   278
      Top             =   6405
      Width           =   615
   End
   Begin VB.Label vLab 
      Caption         =   "vorherige Werte"
      Height          =   244
      Index           =   44
      Left            =   4308
      TabIndex        =   279
      Top             =   6405
      Width           =   1170
   End
   Begin VB.Label vLab 
      Caption         =   "Zahl pW"
      Height          =   244
      Index           =   45
      Left            =   113
      TabIndex        =   280
      Top             =   7003
      Width           =   675
   End
   Begin VB.Label vLab 
      Caption         =   "davon ndE"
      Height          =   244
      Index           =   46
      Left            =   2324
      TabIndex        =   281
      Top             =   7003
      Width           =   843
   End
   Begin VB.Label vLab 
      Caption         =   "davon nachts"
      Height          =   244
      Index           =   47
      Left            =   4195
      TabIndex        =   282
      Top             =   7003
      Width           =   1020
   End
   Begin VB.Label vLab 
      Caption         =   "Aufschreiben"
      Height          =   244
      Index           =   48
      Left            =   5159
      TabIndex        =   283
      Top             =   6720
      Width           =   1020
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "BZ &v d Essen"
      Height          =   244
      Index           =   49
      Left            =   113
      TabIndex        =   284
      Top             =   7371
      Width           =   1020
   End
   Begin VB.Label vLab 
      Caption         =   "danach"
      Height          =   244
      Index           =   50
      Left            =   2834
      TabIndex        =   285
      Top             =   7371
      Width           =   633
   End
   Begin VB.Label vLab 
      Caption         =   "UZ pM"
      Height          =   244
      Index           =   51
      Left            =   113
      TabIndex        =   286
      Top             =   7653
      Width           =   525
   End
   Begin VB.Label vLab 
      Caption         =   "rechtzeitig"
      Height          =   244
      Index           =   52
      Left            =   1246
      TabIndex        =   287
      Top             =   7653
      Width           =   675
   End
   Begin VB.Label vLab 
      Caption         =   "Fremde Hilfe pa"
      Height          =   244
      Index           =   53
      Left            =   2381
      TabIndex        =   288
      Top             =   7653
      Width           =   570
   End
   Begin VB.Label vLab 
      Caption         =   "Bewußtl pa"
      Height          =   244
      Index           =   54
      Left            =   3628
      TabIndex        =   289
      Top             =   7653
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "Keto pa"
      Height          =   244
      Index           =   55
      Left            =   4762
      TabIndex        =   290
      Top             =   7653
      Width           =   633
   End
   Begin VB.Label vLab 
      Caption         =   "gr300 pM"
      Height          =   244
      Index           =   56
      Left            =   5952
      TabIndex        =   291
      Top             =   7653
      Width           =   735
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Bl&uthochdruck"
      Height          =   240
      Index           =   57
      Left            =   56
      TabIndex        =   292
      Top             =   8447
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "BHD seit"
      Height          =   240
      Index           =   58
      Left            =   850
      TabIndex        =   293
      Top             =   8447
      Width           =   1125
   End
   Begin VB.Label vLab 
      Caption         =   "BHD beh mit"
      Height          =   240
      Index           =   59
      Left            =   2035
      TabIndex        =   294
      Top             =   8447
      Width           =   3000
   End
   Begin VB.Label vLab 
      Caption         =   "Blutdruckwerte"
      Height          =   240
      Index           =   60
      Left            =   6039
      TabIndex        =   295
      Top             =   8447
      Width           =   1848
   End
   Begin VB.Label vLab 
      Caption         =   "BDselbst"
      Height          =   240
      Index           =   61
      Left            =   5102
      TabIndex        =   296
      Top             =   8447
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Schwanger"
      Height          =   244
      Index           =   62
      Left            =   113
      TabIndex        =   297
      Top             =   8050
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Schwanger seit"
      Height          =   244
      Index           =   63
      Left            =   1757
      TabIndex        =   298
      Top             =   8050
      Width           =   1185
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Augens&p zuletzt"
      Height          =   240
      Index           =   64
      Left            =   60
      TabIndex        =   299
      Top             =   9074
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "Augensp Befund"
      Height          =   240
      Index           =   65
      Left            =   1256
      TabIndex        =   300
      Top             =   9070
      Width           =   3453
   End
   Begin VB.Label vLab 
      Caption         =   "Netzhaut gelasert"
      Height          =   240
      Index           =   66
      Left            =   4762
      TabIndex        =   301
      Top             =   9070
      Width           =   1260
   End
   Begin VB.Label vLab 
      Caption         =   "Sehminderung unbehebbar"
      Height          =   240
      Index           =   67
      Left            =   6055
      TabIndex        =   302
      Top             =   9070
      Width           =   1800
   End
   Begin VB.Label vLab 
      Caption         =   "Diabet Nierenschaden"
      Height          =   240
      Index           =   68
      Left            =   56
      TabIndex        =   303
      Top             =   9693
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Albumin zuletzt"
      Height          =   240
      Index           =   69
      Left            =   1020
      TabIndex        =   304
      Top             =   9693
      Width           =   1125
   End
   Begin VB.Label vLab 
      Caption         =   "erhöht?"
      Height          =   240
      Index           =   70
      Left            =   2205
      TabIndex        =   305
      Top             =   9693
      Width           =   1185
   End
   Begin VB.Label vLab 
      Caption         =   "Dialyse"
      Height          =   240
      Index           =   71
      Left            =   3390
      TabIndex        =   306
      Top             =   9693
      Width           =   288
   End
   Begin VB.Label vLab 
      Caption         =   "Dialyse seit"
      Height          =   240
      Index           =   72
      Left            =   3664
      TabIndex        =   307
      Top             =   9693
      Width           =   1155
   End
   Begin VB.Label vLab 
      Caption         =   "andere Nierenerkrankung"
      Height          =   240
      Index           =   73
      Left            =   4818
      TabIndex        =   308
      Top             =   9689
      Width           =   3030
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Herzkrankheit"
      Height          =   240
      Index           =   74
      Left            =   60
      TabIndex        =   309
      Top             =   10259
      Width           =   795
   End
   Begin VB.Label vLab 
      Caption         =   "Angina pectoris"
      Height          =   240
      Index           =   75
      Left            =   907
      TabIndex        =   310
      Top             =   10260
      Width           =   855
   End
   Begin VB.Label vLab 
      Caption         =   "Herzinfakt"
      Height          =   240
      Index           =   76
      Left            =   1757
      TabIndex        =   311
      Top             =   10260
      Width           =   573
   End
   Begin VB.Label vLab 
      Caption         =   "wann"
      Height          =   240
      Index           =   77
      Left            =   2324
      TabIndex        =   312
      Top             =   10260
      Width           =   675
   End
   Begin VB.Label vLab 
      Caption         =   "PTCA oder Stent"
      Height          =   240
      Index           =   78
      Left            =   3061
      TabIndex        =   313
      Top             =   10260
      Width           =   510
   End
   Begin VB.Label vLab 
      Caption         =   "Bypass kardial"
      Height          =   240
      Index           =   79
      Left            =   3571
      TabIndex        =   314
      Top             =   10260
      Width           =   480
   End
   Begin VB.Label vLab 
      Caption         =   "wann"
      Height          =   240
      Index           =   80
      Left            =   4081
      TabIndex        =   315
      Top             =   10261
      Width           =   525
   End
   Begin VB.Label vLab 
      Caption         =   "Herzschwäche"
      Height          =   240
      Index           =   81
      Left            =   4648
      TabIndex        =   316
      Top             =   10260
      Width           =   963
   End
   Begin VB.Label vLab 
      Caption         =   "Herzkrankheit Beschreibung"
      Height          =   240
      Index           =   82
      Left            =   5585
      TabIndex        =   317
      Top             =   10260
      Width           =   2265
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "H&irndurchblutungsstörung"
      Height          =   240
      Index           =   83
      Left            =   7880
      TabIndex        =   318
      Top             =   0
      Width           =   975
   End
   Begin VB.Label vLab 
      Caption         =   "Schlaganfall"
      Height          =   240
      Index           =   84
      Left            =   8900
      TabIndex        =   319
      Top             =   0
      Width           =   2310
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "B&eindurchblutungsstörung"
      Height          =   240
      Index           =   85
      Left            =   11225
      TabIndex        =   320
      Top             =   0
      Width           =   858
   End
   Begin VB.Label vLab 
      Caption         =   "Schaufensterkrankheit"
      Height          =   240
      Index           =   86
      Left            =   12132
      TabIndex        =   321
      Top             =   0
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Bypaß peripher"
      Height          =   240
      Index           =   87
      Left            =   13096
      TabIndex        =   322
      Top             =   0
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "Geschwür"
      Height          =   240
      Index           =   88
      Left            =   13322
      TabIndex        =   323
      Top             =   0
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Amputation"
      Height          =   240
      Index           =   89
      Left            =   14222
      TabIndex        =   324
      Top             =   0
      Width           =   2433
   End
   Begin VB.Label vLab 
      Caption         =   "pAVK Beschreibung"
      Height          =   240
      Index           =   90
      Left            =   7880
      TabIndex        =   325
      Top             =   510
      Width           =   7638
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ameisenlaufe&n"
      Height          =   240
      Index           =   91
      Left            =   7880
      TabIndex        =   326
      Top             =   1395
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "Ausmaß"
      Height          =   240
      Index           =   92
      Left            =   9080
      TabIndex        =   327
      Top             =   1417
      Width           =   675
   End
   Begin VB.Label vLab 
      Caption         =   "Druckstellen"
      Height          =   240
      Index           =   93
      Left            =   10091
      TabIndex        =   328
      Top             =   1417
      Width           =   960
   End
   Begin VB.Label vLab 
      Caption         =   "Verformungen"
      Height          =   240
      Index           =   94
      Left            =   11055
      TabIndex        =   329
      Top             =   1417
      Width           =   1020
   End
   Begin VB.Label vLab 
      Caption         =   "Beschreibung"
      Height          =   240
      Index           =   95
      Left            =   12132
      TabIndex        =   330
      Top             =   1417
      Width           =   828
   End
   Begin VB.Label vLab 
      Caption         =   "Neue Fußkomplikationen"
      Height          =   240
      Index           =   96
      Left            =   7937
      TabIndex        =   331
      Top             =   2041
      Width           =   2265
   End
   Begin VB.Label vLab 
      Caption         =   "Entleerungsst Magen"
      Height          =   240
      Index           =   97
      Left            =   10261
      TabIndex        =   332
      Top             =   2044
      Width           =   1575
   End
   Begin VB.Label vLab 
      Caption         =   "Harnblase"
      Height          =   240
      Index           =   98
      Left            =   11853
      TabIndex        =   333
      Top             =   2041
      Width           =   1155
   End
   Begin VB.Label vLab 
      Caption         =   "Schwindel Aufstehen"
      Height          =   240
      Index           =   99
      Left            =   13044
      TabIndex        =   334
      Top             =   2041
      Width           =   1248
   End
   Begin VB.Label vLab 
      Caption         =   "Folgeerkrankungen Haut"
      Height          =   240
      Index           =   100
      Left            =   7937
      TabIndex        =   335
      Top             =   2608
      Width           =   1830
   End
   Begin VB.Label vLab 
      Caption         =   "Bewegungseinschränkungen"
      Height          =   240
      Index           =   101
      Left            =   9808
      TabIndex        =   336
      Top             =   2608
      Width           =   2310
   End
   Begin VB.Label vLab 
      Caption         =   "Sexualstörung"
      Height          =   240
      Index           =   102
      Left            =   12188
      TabIndex        =   337
      Top             =   2608
      Width           =   1080
   End
   Begin VB.Label vLab 
      Caption         =   "Sexualstörung seit"
      Height          =   240
      Index           =   103
      Left            =   13268
      TabIndex        =   338
      Top             =   2608
      Width           =   1938
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Weitere Anamnese"
      Height          =   240
      Index           =   104
      Left            =   7937
      TabIndex        =   339
      Top             =   3175
      Width           =   7638
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Alk&ohol"
      Height          =   240
      Index           =   105
      Left            =   10318
      TabIndex        =   340
      Top             =   4082
      Width           =   690
   End
   Begin VB.Label vLab 
      Caption         =   "Tabak"
      Height          =   240
      Index           =   106
      Left            =   7937
      TabIndex        =   341
      Top             =   4082
      Width           =   633
   End
   Begin VB.Label vLab 
      Caption         =   "Weitere Medikation"
      Height          =   240
      Index           =   107
      Left            =   7937
      TabIndex        =   342
      Top             =   4365
      Width           =   1653
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Liph&ypertrophien Abdomen"
      Height          =   240
      Index           =   108
      Left            =   7937
      TabIndex        =   343
      Top             =   5272
      Width           =   2310
   End
   Begin VB.Label vLab 
      Caption         =   "Liphypertrophien Beine"
      Height          =   240
      Index           =   109
      Left            =   10247
      TabIndex        =   344
      Top             =   5272
      Width           =   2310
   End
   Begin VB.Label vLab 
      Caption         =   "Liphypertrophien Arme"
      Height          =   240
      Index           =   110
      Left            =   12557
      TabIndex        =   345
      Top             =   5272
      Width           =   3018
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Hyper&keratosen"
      Height          =   240
      Index           =   111
      Left            =   7993
      TabIndex        =   346
      Top             =   6179
      Width           =   3285
   End
   Begin VB.Label vLab 
      Caption         =   "Ulcera"
      Height          =   240
      Index           =   112
      Left            =   11338
      TabIndex        =   347
      Top             =   6179
      Width           =   1485
   End
   Begin VB.Label vLab 
      Caption         =   "Kraft Zehenheber"
      Height          =   240
      Index           =   113
      Left            =   12869
      TabIndex        =   348
      Top             =   6179
      Width           =   915
   End
   Begin VB.Label vLab 
      Caption         =   "Kraft Zehenbeuger"
      Height          =   240
      Index           =   114
      Left            =   13833
      TabIndex        =   349
      Top             =   6179
      Width           =   843
   End
   Begin VB.Label vLab 
      Caption         =   "Kraft Knie"
      Height          =   240
      Index           =   115
      Left            =   14740
      TabIndex        =   350
      Top             =   6179
      Width           =   855
   End
   Begin VB.Label vLab 
      Caption         =   "ASR"
      Height          =   240
      Index           =   116
      Left            =   7993
      TabIndex        =   351
      Top             =   6746
      Width           =   1080
   End
   Begin VB.Label vLab 
      Caption         =   "PSR"
      Height          =   240
      Index           =   117
      Left            =   9127
      TabIndex        =   352
      Top             =   6746
      Width           =   990
   End
   Begin VB.Label vLab 
      Caption         =   "Oberflächensensibilität"
      Height          =   240
      Index           =   118
      Left            =   10148
      TabIndex        =   353
      Top             =   6746
      Width           =   1638
   End
   Begin VB.Label vLab 
      Caption         =   "Monofilamenttest"
      Height          =   240
      Index           =   119
      Left            =   11848
      TabIndex        =   354
      Top             =   6746
      Width           =   1560
   End
   Begin VB.Label vLab 
      Caption         =   "Kalt-Warm"
      Height          =   240
      Index           =   120
      Left            =   13408
      TabIndex        =   355
      Top             =   6746
      Width           =   1560
   End
   Begin VB.Label vLab 
      Caption         =   "Vibration IK"
      Height          =   240
      Index           =   121
      Left            =   7993
      TabIndex        =   356
      Top             =   7313
      Width           =   1560
   End
   Begin VB.Label vLab 
      Caption         =   "Vibration Großzehe"
      Height          =   240
      Index           =   122
      Left            =   9581
      TabIndex        =   357
      Top             =   7313
      Width           =   1560
   End
   Begin VB.Label vLab 
      Caption         =   "Puls Leiste"
      Height          =   240
      Index           =   123
      Left            =   11225
      TabIndex        =   358
      Top             =   7313
      Width           =   1023
   End
   Begin VB.Label vLab 
      Caption         =   "Puls Kniekehle"
      Height          =   240
      Index           =   124
      Left            =   12302
      TabIndex        =   359
      Top             =   7313
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Puls Atp"
      Height          =   240
      Index           =   125
      Left            =   13266
      TabIndex        =   360
      Top             =   7313
      Width           =   1080
   End
   Begin VB.Label vLab 
      Caption         =   "Puls Adp"
      Height          =   240
      Index           =   126
      Left            =   14400
      TabIndex        =   361
      Top             =   7313
      Width           =   1140
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "&RR"
      Height          =   240
      Index           =   127
      Left            =   7993
      TabIndex        =   362
      Top             =   7937
      Width           =   393
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Her&z"
      Height          =   238
      Index           =   128
      Left            =   7993
      TabIndex        =   363
      Top             =   8505
      Width           =   450
   End
   Begin VB.Label vLab 
      Caption         =   "Lunge"
      Height          =   238
      Index           =   129
      Left            =   12925
      TabIndex        =   364
      Top             =   8505
      Width           =   555
   End
   Begin VB.Label vLab 
      Caption         =   "Bauch"
      Height          =   238
      Index           =   130
      Left            =   7993
      TabIndex        =   365
      Top             =   8755
      Width           =   573
   End
   Begin VB.Label vLab 
      Caption         =   "WS"
      Height          =   238
      Index           =   131
      Left            =   10658
      TabIndex        =   366
      Top             =   8755
      Width           =   285
   End
   Begin VB.Label vLab 
      Caption         =   "NL"
      Height          =   238
      Index           =   132
      Left            =   12076
      TabIndex        =   367
      Top             =   8755
      Width           =   240
   End
   Begin VB.Label vLab 
      Caption         =   "SD"
      Height          =   238
      Index           =   133
      Left            =   7993
      TabIndex        =   368
      Top             =   8993
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "NNH"
      Height          =   238
      Index           =   134
      Left            =   9467
      TabIndex        =   369
      Top             =   8993
      Width           =   390
   End
   Begin VB.Label vLab 
      Caption         =   "Zähne"
      Height          =   238
      Index           =   135
      Left            =   11225
      TabIndex        =   370
      Top             =   8993
      Width           =   573
   End
   Begin VB.Label vLab 
      Caption         =   "Mundhöhle"
      Height          =   238
      Index           =   136
      Left            =   7993
      TabIndex        =   371
      Top             =   9243
      Width           =   915
   End
   Begin VB.Label vLab 
      Caption         =   "LK"
      Height          =   238
      Index           =   137
      Left            =   7993
      TabIndex        =   372
      Top             =   9581
      Width           =   330
   End
   Begin VB.Label vLab 
      Caption         =   "Neuro sonst"
      Height          =   238
      Index           =   138
      Left            =   8050
      TabIndex        =   373
      Top             =   10035
      Width           =   618
   End
   Begin VB.Label vLab 
      Caption         =   "Weitere Befunde"
      Height          =   240
      Index           =   139
      Left            =   9807
      TabIndex        =   374
      Top             =   9751
      Width           =   1308
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "S&chulung"
      Height          =   256
      Index           =   140
      Left            =   13039
      TabIndex        =   375
      Top             =   10320
      Width           =   840
   End
   Begin VB.Label vLab 
      Caption         =   "DMP"
      Height          =   240
      Index           =   141
      Left            =   7993
      TabIndex        =   376
      Top             =   10601
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "HA"
      Height          =   240
      Index           =   142
      Left            =   9977
      TabIndex        =   377
      Top             =   10601
      Width           =   285
   End
   Begin VB.Label vLab 
      Caption         =   "Beinbefund"
      Height          =   240
      Index           =   143
      Left            =   7993
      TabIndex        =   378
      Top             =   5839
      Width           =   960
   End
   Begin VB.Label vLab 
      Caption         =   "BMI"
      Height          =   238
      Index           =   144
      Left            =   4762
      TabIndex        =   379
      Top             =   3005
      Width           =   450
   End
   Begin VB.Label vLab 
      Caption         =   "kg/m²"
      Height          =   238
      Index           =   145
      Left            =   5953
      TabIndex        =   380
      Top             =   3005
      Width           =   624
   End
   Begin VB.Label vLab 
      Caption         =   "letzte Änderung"
      Height          =   240
      Index           =   146
      Left            =   15817
      TabIndex        =   381
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label vLab 
      Caption         =   "Fußplfege"
      Height          =   240
      Index           =   147
      Left            =   13096
      TabIndex        =   382
      Top             =   1417
      Width           =   795
   End
   Begin VB.Label vLab 
      Caption         =   "Podologie"
      Height          =   240
      Index           =   148
      Left            =   13889
      TabIndex        =   383
      Top             =   1417
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Orthop.Einlg/Schuhe diab&'ger"
      Height          =   240
      Index           =   149
      Left            =   14683
      TabIndex        =   384
      Top             =   1417
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Diagnosen:"
      Height          =   240
      Index           =   150
      Left            =   14400
      TabIndex        =   385
      Top             =   1984
      Width           =   1125
   End
   Begin VB.Label vLab 
      Caption         =   "vorgestellt"
      Height          =   238
      Index           =   151
      Left            =   1927
      TabIndex        =   386
      Top             =   340
      Width           =   795
   End
   Begin VB.Label vLab 
      Caption         =   "RR TurboMed:"
      Height          =   240
      Index           =   152
      Left            =   7993
      TabIndex        =   387
      Top             =   8221
      Width           =   450
   End
   Begin VB.Label vLab 
      Caption         =   "Versicherung:"
      Height          =   240
      Index           =   153
      Left            =   8050
      TabIndex        =   388
      Top             =   10318
      Width           =   960
   End
   Begin VB.Label vLab 
      Caption         =   "vormittags"
      Height          =   244
      Index           =   154
      Left            =   2721
      TabIndex        =   389
      Top             =   5160
      Width           =   405
   End
   Begin VB.Label vLab 
      Caption         =   "nachmittags"
      Height          =   244
      Index           =   155
      Left            =   4705
      TabIndex        =   390
      Top             =   5160
      Width           =   270
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "BZ-&Messungen: selbst?"
      Height          =   244
      Index           =   156
      Left            =   56
      TabIndex        =   391
      Top             =   6720
      Width           =   1755
   End
   Begin VB.Label vLab 
      Caption         =   "Gerät:"
      Height          =   244
      Index           =   157
      Left            =   2323
      TabIndex        =   392
      Top             =   6720
      Width           =   525
   End
   Begin VB.Label vLab 
      Caption         =   "UZ Tageszeit"
      Height          =   244
      Index           =   158
      Left            =   5102
      TabIndex        =   393
      Top             =   7257
      Width           =   1065
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Beinödeme&, Venenkrkht:"
      Height          =   238
      Index           =   159
      Left            =   10261
      TabIndex        =   394
      Top             =   9243
      Width           =   1080
   End
   Begin VB.Label vLab 
      Caption         =   "Carotiden"
      Height          =   238
      Index           =   160
      Left            =   11225
      TabIndex        =   395
      Top             =   8505
      Width           =   390
   End
   Begin VB.Label vLab 
      Caption         =   "Dm-Schlgn:"
      Height          =   256
      Index           =   161
      Left            =   14850
      TabIndex        =   396
      Top             =   10320
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "RR-Schlgn:"
      Height          =   256
      Index           =   162
      Left            =   16384
      TabIndex        =   397
      Top             =   10318
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "DMP hier:"
      Height          =   240
      Index           =   163
      Left            =   8957
      TabIndex        =   398
      Top             =   10601
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Tkz"
      Height          =   225
      Index           =   164
      Left            =   6973
      TabIndex        =   399
      Top             =   56
      Width           =   390
   End
   Begin VB.Label vLab 
      Caption         =   "Therapieart"
      Height          =   224
      Index           =   165
      Left            =   4818
      TabIndex        =   400
      Top             =   626
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "BZ-Meßger ausgew"
      Height          =   210
      Index           =   166
      Left            =   14630
      TabIndex        =   401
      Top             =   10830
      Width           =   1470
   End
   Begin VB.Label vLab 
      Caption         =   "orthop. SM aufgekl."
      Height          =   225
      Index           =   167
      Left            =   14630
      TabIndex        =   402
      Top             =   11282
      Width           =   1470
   End
   Begin VB.Label vLab 
      Caption         =   "üb.Podologie aufgek."
      Height          =   240
      Index           =   168
      Left            =   14630
      TabIndex        =   403
      Top             =   11532
      Width           =   1530
   End
   Begin VB.Label vLab 
      Caption         =   "&8: Merkbl DFS ausgehändigt"
      Height          =   240
      Index           =   169
      Left            =   14626
      TabIndex        =   404
      Top             =   11080
      Width           =   1470
   End
   Begin VB.Label vLab 
      Caption         =   " DMP-Beteiligung"
      Height          =   210
      Index           =   170
      Left            =   16611
      TabIndex        =   405
      Top             =   11508
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "&9: Schulungsaufkl"
      Height          =   210
      Index           =   171
      Left            =   16611
      TabIndex        =   406
      Top             =   11281
      Width           =   1170
   End
   Begin VB.Label vLab 
      Caption         =   "-&>"
      Height          =   227
      Index           =   172
      Left            =   6689
      TabIndex        =   407
      Top             =   623
      Width           =   178
   End
   Begin VB.Label vLab 
      Caption         =   "SchGr"
      Height          =   227
      Index           =   173
      Left            =   6746
      TabIndex        =   408
      Top             =   340
      Width           =   518
   End
   Begin VB.Label vLab 
      Caption         =   "Med.Netz:"
      Height          =   240
      Index           =   174
      Left            =   6633
      TabIndex        =   409
      Top             =   907
      Width           =   795
   End
   Begin VB.Label vLab 
      Caption         =   "-"
      Height          =   180
      Index           =   175
      Left            =   8047
      TabIndex        =   412
      Top             =   11735
      Width           =   240
   End
   Begin VB.Label vLab 
      Caption         =   "-"
      Height          =   180
      Index           =   176
      Left            =   11168
      TabIndex        =   413
      Top             =   11735
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "&+"
      Height          =   285
      Index           =   177
      Left            =   16157
      TabIndex        =   414
      Top             =   10885
      Width           =   255
   End
   Begin VB.Label vLab 
      Caption         =   "Duplex"
      Height          =   283
      Index           =   178
      Left            =   6292
      TabIndex        =   415
      Top             =   1026
      Width           =   850
   End
   Begin VB.Label vLab 
      Caption         =   "Doppler"
      Height          =   240
      Index           =   179
      Left            =   3628
      TabIndex        =   416
      Top             =   11962
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "Duplex"
      Height          =   226
      Index           =   180
      Left            =   8447
      TabIndex        =   417
      Top             =   11905
      Width           =   628
   End
   Begin VB.Label vLab 
      Caption         =   "Sono"
      Height          =   227
      Index           =   181
      Left            =   13776
      TabIndex        =   418
      Top             =   11848
      Width           =   462
   End
   Begin VB.Label vLab 
      Caption         =   "-"
      Height          =   226
      Index           =   182
      Left            =   3855
      TabIndex        =   419
      Top             =   340
      Width           =   134
   End
   Begin VB.Label vLab 
      Caption         =   "Daten&quelle:"
      Height          =   240
      Index           =   183
      Left            =   113
      TabIndex        =   421
      Top             =   56
      Width           =   1005
   End
End
Attribute VB_Name = "AnBog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alleDatensätze_Click()
 Call DqC_A(Me)
End Sub

Private Sub An1Aufruf_Click()
 Call do_An1Aufruf_click(Me)
End Sub
Private Sub An2Aufruf_Click()
 Call do_An2Aufruf_click(Me)
End Sub
Private Sub AnAAufruf_Click()
 Call do_AnAAufruf_click(Me)
End Sub

Private Sub AugenBefunde_Click()
 Call do_AugenBefunde_Click(Me)
End Sub

Private Sub Brief_Click() ' Winword-Brief schreiben
 Call tu_brief(Me.vTextB(1), 0)
 Call WD
 Call Sound(WinDir + "\media\Windows XP-Batterie niedrig.wav")
End Sub ' Brief_Click() ' Winword-Brief schreiben

Private Sub Briefe_Click()
 Call do_Briefe_Click(Me)
End Sub

Private Sub BZKurven_Click()
 Call do_BZKurven_Click(Me)
End Sub

Private Sub CheckAufruf_Click()
 Call do_CheckAufruf_click(Me)
End Sub

Private Sub Ausgabe_Click()
 Call do_Ausgabe_Click(Me)
End Sub

Private Sub Datenbank_Aufruf_Click()
 Call do_Datenbank_Aufruf_Click(Me)
End Sub

Private Sub Datenquelle_Change()
 Call do_Datenquelle_Change(Me)
End Sub

Private Sub DMPAusgeb_Click()
 Call do_DMPAusgebStandAlone(Me.vTextB(1))
End Sub

Private Sub DokDown_Click()
 Call do_DokDown(Me)
End Sub

Private Sub Dokumente_Click()
 Call do_Dokumente_Click(Me)
End Sub

Private Sub DReset_Click()
 Call do_Diagnosen_Reset(Me)
End Sub

Private Sub Form_AfterInsert()
 Call do_Form_AfterUpdate(Me)
End Sub

Private Sub Form_AfterUpdate()
 Call do_Form_AfterUpdate(Me)
End Sub

Private Sub Form_BeforeUpdate(Cancel%)
 Call do_form_beforeUpdate(Me)
End Sub

Private Sub Form_Current()
 Call do_Form_Current(Me)
End Sub

Private Sub Form_Open(Cancel%)
 Call do_Form_Open(Cancel, Me)
End Sub
Private Sub Form_Load()
 'Me.Controls!Diagnosen.ControlSource = "=replace(replace(recordset!Diagnosen,chr(11),chr(13)+chr(10)),chr(9),"" "")"
End Sub
Sub Form_Close()
  Call do_Form_Close(Me)
End Sub

Private Sub HA_Auswahl_Change()
 ReDim Preserve HANrBf(2)
 If HA_Auswahl Like "##/#####*" Then
  If Replace(HA_Auswahl, "/", "") <> HANrBf(1) Then
   HANrBf(1) = Replace(HA_Auswahl, "/", "")
   Call Sound(WinDir + "\media\Windows XP-Batterie niedrig.wav")
  End If
 End If
End Sub

Private Sub HA_Auswahl2_Change()
 ReDim Preserve HANrBf(2)
 If HA_Auswahl2 Like "##/#####*" Then
  If Replace(HA_Auswahl2, "/", "") <> HANrBf(2) Then
   HANrBf(2) = Replace(HA_Auswahl2, "/", "")
   Call Sound(WinDir + "\media\Windows XP-Batterie niedrig.wav")
  End If
 End If
End Sub

Private Sub HA_Click()
 ' Hier Auswahl
End Sub
Private Sub HA1_Bez_Click()
 Call do_ha_click(HA_Auswahl)
End Sub

Private Sub HA2_Bez_Click()
 Call do_ha_click(HA_Auswahl2)
End Sub

Private Sub HANr_Click()
 Call do_ha_click(HANr)
End Sub

Private Sub HANr2_Click()
 Call do_ha_click(HANr2)
End Sub

Private Sub Hirndurchblutungsstörung_Change()
' Dim Cancel%
' Call do_Hirndurchblutungsstörung_Exit(Cancel, Me)
End Sub

Private Sub Hirndurchblutungsstörung_Exit(Cancel%)
' Call do_Hirndurchblutungsstörung_Exit(Cancel, Me)
End Sub

Private Sub Hirndurchblutungsstörung_LostFocus()
' Dim Cancel%
' Call do_Hirndurchblutungsstörung_Exit(Cancel, Me)
End Sub

Private Sub Labor_Click()
 Call do_Labor_Click(Me)
End Sub

Private Sub Labor_KeyPress(KeyCode%)
  Call Taste(KeyCode, 0, Me)
End Sub
Private Sub U1_Click()
 Call do_u_Click(Me, 1)
End Sub

Private Sub U2_Click()
 Call do_u_Click(Me, 2)
End Sub
Private Sub U3_Click()
 Call do_u_Click(Me, 3)
End Sub
Private Sub U4_Click()
 Call do_u_Click(Me, 4)
End Sub
Private Sub U5_Click()
 Call do_u_Click(Me, 5)
End Sub
Private Sub U6_Click()
 Call do_u_Click(Me, 6)
End Sub
Private Sub U7_Click()
 Call do_u_Click(Me, 7)
End Sub
Private Sub U8_Click()
 Call do_u_Click(Me, 8)
End Sub
Private Sub U9_Click()
 Call do_u_Click(Me, 9)
End Sub
Private Sub U10_Click()
 Call do_u_Click(Me, 10)
End Sub
Private Sub U11_Click()
 Call do_u_Click(Me, 11)
End Sub
Private Sub U12_Click()
 Call do_u_Click(Me, 12)
End Sub
Private Sub U13_Click()
 Call do_u_Click(Me, 13)
End Sub
Private Sub U14_Click()
 Call do_u_Click(Me, 14)
End Sub
Private Sub U15_Click()
 Call do_u_Click(Me, 15)
End Sub
Private Sub U16_Click()
 Call do_u_Click(Me, 16)
End Sub
Private Sub U17_Click()
 Call do_u_Click(Me, 17)
End Sub
Private Sub U18_Click()
 Call do_u_Click(Me, 18)
End Sub
Private Sub U19_Click()
 Call do_u_Click(Me, 19)
End Sub
Private Sub U20_Click()
 Call do_u_Click(Me, 20)
End Sub
Private Sub U21_Click()
 Call do_u_Click(Me, 21)
End Sub
Private Sub U22_Click()
 Call do_u_Click(Me, 22)
End Sub
Private Sub U23_Click()
 Call do_u_Click(Me, 23)
End Sub
Function do_u_Click(frm As Form, nr%)
 Debug.Print frm.Controls("U" + CStr(nr)).Value
 On Error GoTo 0
 On Error Resume Next
 Err.Clear
 rfDE.Seek "=", frm!Pat_id, MDIICD(nr), MDIDiag(nr)
 If Err.Number > 0 Then Exit Function
 If rfDE.NoMatch And frm.Controls("U" + CStr(nr)).Value = True Then
  rfDE.AddNew
  rfDE!Pat_id = frm!Pat_id
  rfDE!Name = GesNam(frm!Recordset) 'frm!Nachname + " " + frm!Vorname
  rfDE!ICD = MDIICD(nr)
  rfDE!Diagnose = MDIDiag(nr)
  rfDE.Update
 ElseIf Not rfDE.NoMatch And frm.Controls("U" + CStr(nr)).Value = False Then
  rfDE.Delete
 End If
End Function ' do_u_Click(frm As Form, Nr%)


Private Sub UnausgefNZ_Click()
 Call DqC_UNZ(Me)
End Sub

Private Sub Unausgefülle_Click()
 Call DqC_U(Me)
End Sub
