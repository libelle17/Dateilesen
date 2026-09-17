VERSION 5.00
Begin VB.Form AnBog 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "AnBog"
   ClientHeight    =   13530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   28860
   KeyPreview      =   -1  'True
   LinkTopic       =   "AnBog"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13530
   ScaleWidth      =   28860
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Pat_ID"
      Height          =   238
      Index           =   1
      Left            =   113
      TabIndex        =   1
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
      TabIndex        =   4
      Top             =   340
      Width           =   1017
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Diabetestyp"
      Height          =   285
      Index           =   4
      Left            =   114
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   6
      Top             =   879
      Width           =   1140
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
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
      BackColor       =   &H0000FF00&
      DataField       =   "Tabletten seit"
      Height          =   285
      Index           =   6
      Left            =   2484
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      Top             =   879
      Width           =   1140
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Insulin seit"
      Height          =   285
      Index           =   7
      Left            =   3669
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   12
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
      TabIndex        =   14
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
      TabIndex        =   16
      Top             =   2289
      Width           =   7638
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H8000000F&
      DataField       =   "Größe"
      Height          =   238
      Index           =   10
      Left            =   794
      TabIndex        =   18
      Top             =   3005
      Width           =   840
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H8000000F&
      DataField       =   "Gewicht"
      Height          =   238
      Index           =   11
      Left            =   2438
      TabIndex        =   20
      Top             =   3005
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H8000000F&
      DataField       =   "Tendenz"
      Height          =   238
      Index           =   12
      Left            =   4081
      TabIndex        =   22
      Top             =   3004
      Width           =   450
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 1"
      Height          =   256
      Index           =   13
      Left            =   1019
      TabIndex        =   24
      Top             =   3344
      Width           =   3633
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 1 Menge"
      Height          =   256
      Index           =   14
      Left            =   5385
      TabIndex        =   26
      Top             =   3346
      Width           =   2490
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 2"
      Height          =   256
      Index           =   15
      Left            =   737
      TabIndex        =   28
      Top             =   3614
      Width           =   3915
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 2 Menge"
      Height          =   256
      Index           =   16
      Left            =   5385
      TabIndex        =   30
      Top             =   3616
      Width           =   2478
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 3"
      Height          =   256
      Index           =   17
      Left            =   737
      TabIndex        =   32
      Top             =   3883
      Width           =   3915
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 3 Menge"
      Height          =   256
      Index           =   18
      Left            =   5385
      TabIndex        =   34
      Top             =   3885
      Width           =   2490
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 4"
      Height          =   256
      Index           =   19
      Left            =   737
      TabIndex        =   36
      Top             =   4181
      Width           =   3018
   End
   Begin VB.TextBox vTextB 
      DataField       =   "DiabetesMedikament 4 Menge"
      Height          =   256
      Index           =   20
      Left            =   5385
      TabIndex        =   38
      Top             =   4183
      Width           =   2490
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Insulinpumpe"
      DataField       =   "Insulinpumpe"
      Height          =   165
      Index           =   1
      Left            =   1077
      TabIndex        =   40
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
      TabIndex        =   42
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
      TabIndex        =   44
      Top             =   4479
      Width           =   4188
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten gesamt"
      Height          =   244
      Index           =   23
      Left            =   907
      TabIndex        =   46
      Top             =   4876
      Width           =   585
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten früh"
      Height          =   244
      Index           =   24
      Left            =   1927
      TabIndex        =   48
      Top             =   4875
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten ZM früh"
      Height          =   244
      Index           =   25
      Left            =   3117
      TabIndex        =   50
      Top             =   4875
      Width           =   513
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten mittags"
      Height          =   244
      Index           =   26
      Left            =   3912
      TabIndex        =   52
      Top             =   4875
      Width           =   750
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten nachmittags"
      Height          =   244
      Index           =   27
      Left            =   4989
      TabIndex        =   54
      Top             =   4875
      Width           =   630
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten abends"
      Height          =   244
      Index           =   28
      Left            =   6009
      TabIndex        =   56
      Top             =   4875
      Width           =   858
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Broteinheiten nachts"
      Height          =   244
      Index           =   29
      Left            =   7257
      TabIndex        =   58
      Top             =   4762
      Width           =   615
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit früh"
      Height          =   244
      Index           =   30
      Left            =   1927
      TabIndex        =   60
      Top             =   5160
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit mittags"
      Height          =   244
      Index           =   31
      Left            =   3968
      TabIndex        =   62
      Top             =   5160
      Width           =   690
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit abends"
      Height          =   244
      Index           =   32
      Left            =   6009
      TabIndex        =   64
      Top             =   5160
      Width           =   870
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit spät"
      Height          =   244
      Index           =   33
      Left            =   7256
      TabIndex        =   66
      Top             =   5047
      Width           =   630
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Spritz-Eß-Abstand früh"
      Height          =   244
      Index           =   34
      Left            =   1927
      TabIndex        =   68
      Top             =   5443
      Width           =   783
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Spritz-Eß-Abstand mittags"
      Height          =   244
      Index           =   35
      Left            =   3968
      TabIndex        =   70
      Top             =   5443
      Width           =   720
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Spritz-Eß-Abstand abends"
      Height          =   244
      Index           =   36
      Left            =   6066
      TabIndex        =   72
      Top             =   5443
      Width           =   810
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Spritzstelle früh"
      Height          =   244
      Index           =   37
      Left            =   1927
      TabIndex        =   74
      Top             =   5727
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Spritzstelle mittags"
      Height          =   244
      Index           =   38
      Left            =   3968
      TabIndex        =   76
      Top             =   5727
      Width           =   723
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Spritzstelle abends"
      Height          =   244
      Index           =   39
      Left            =   6066
      TabIndex        =   78
      Top             =   5727
      Width           =   795
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Spritzstelle nachts"
      Height          =   244
      Index           =   40
      Left            =   7256
      TabIndex        =   80
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
      TabIndex        =   82
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
      TabIndex        =   84
      Top             =   6122
      Width           =   2778
   End
   Begin VB.TextBox vTextB 
      DataField       =   "letztes HbA1c"
      Height          =   244
      Index           =   43
      Left            =   1077
      TabIndex        =   86
      Top             =   6405
      Width           =   1065
   End
   Begin VB.TextBox vTextB 
      DataField       =   "gemessen am"
      Height          =   244
      Index           =   44
      Left            =   2769
      TabIndex        =   88
      Top             =   6405
      Width           =   1545
   End
   Begin VB.TextBox vTextB 
      DataField       =   "vorherige Werte"
      Height          =   244
      Index           =   45
      Left            =   5555
      TabIndex        =   90
      Top             =   6405
      Width           =   2310
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZMessungen pW"
      Height          =   244
      Index           =   46
      Left            =   850
      TabIndex        =   92
      Top             =   7003
      Width           =   1365
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZMessungen pW ndE"
      Height          =   244
      Index           =   47
      Left            =   3174
      TabIndex        =   94
      Top             =   7003
      Width           =   858
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZMessungen p W nachts"
      Height          =   244
      Index           =   48
      Left            =   5272
      TabIndex        =   96
      Top             =   7003
      Width           =   915
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Aufschreiben"
      Height          =   244
      Index           =   49
      Left            =   6236
      TabIndex        =   98
      Top             =   6720
      Width           =   570
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZWerte v d Essen"
      Height          =   244
      Index           =   50
      Left            =   1190
      TabIndex        =   100
      Top             =   7371
      Width           =   1590
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZWerte n d Essen"
      Height          =   244
      Index           =   51
      Left            =   3514
      TabIndex        =   102
      Top             =   7370
      Width           =   1488
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Unterzucker pM"
      Height          =   244
      Index           =   52
      Left            =   680
      TabIndex        =   104
      Top             =   7653
      Width           =   510
   End
   Begin VB.TextBox vTextB 
      DataField       =   "UZ rechtzeitig"
      Height          =   244
      Index           =   53
      Left            =   1984
      TabIndex        =   106
      Top             =   7653
      Width           =   330
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Fremde Hilfe pa"
      Height          =   244
      Index           =   54
      Left            =   3004
      TabIndex        =   108
      Top             =   7653
      Width           =   570
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Bewußtlos pa"
      Height          =   244
      Index           =   55
      Left            =   4308
      TabIndex        =   110
      Top             =   7653
      Width           =   450
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Keto pa"
      Height          =   244
      Index           =   56
      Left            =   5442
      TabIndex        =   112
      Top             =   7653
      Width           =   453
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZgr300 pM"
      Height          =   244
      Index           =   57
      Left            =   6746
      TabIndex        =   114
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
      TabIndex        =   116
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
      TabIndex        =   118
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
      TabIndex        =   120
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
      TabIndex        =   122
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
      TabIndex        =   124
      Top             =   8687
      Width           =   900
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Schwanger"
      Height          =   244
      Index           =   63
      Left            =   1077
      TabIndex        =   126
      Top             =   8050
      Width           =   630
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Schwanger seit"
      Height          =   244
      Index           =   64
      Left            =   3004
      TabIndex        =   128
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
      TabIndex        =   130
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
      TabIndex        =   132
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
      TabIndex        =   134
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
      TabIndex        =   136
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
      TabIndex        =   138
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
      TabIndex        =   140
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
      TabIndex        =   142
      Top             =   9933
      Width           =   1185
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Dialyse"
      DataField       =   "Dialyse"
      Height          =   285
      Index           =   2
      Left            =   3390
      TabIndex        =   144
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
      TabIndex        =   146
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
      TabIndex        =   148
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
      TabIndex        =   150
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
      TabIndex        =   152
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
      TabIndex        =   154
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
      TabIndex        =   156
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
      TabIndex        =   158
      Top             =   10500
      Width           =   510
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Bypass kardial"
      DataField       =   "Bypass kardial"
      Height          =   285
      Index           =   3
      Left            =   3571
      TabIndex        =   160
      Top             =   10500
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
      TabIndex        =   162
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
      TabIndex        =   164
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
      TabIndex        =   166
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
      TabIndex        =   168
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
      TabIndex        =   170
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
      TabIndex        =   172
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
      TabIndex        =   174
      Top             =   240
      Width           =   900
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Bypaß peripher"
      DataField       =   "Bypaß peripher"
      Height          =   285
      Index           =   4
      Left            =   13096
      TabIndex        =   176
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
      TabIndex        =   178
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
      TabIndex        =   180
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
      TabIndex        =   182
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
      TabIndex        =   184
      Top             =   1635
      Width           =   915
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Ameisen Ausmaß"
      Height          =   285
      Index           =   90
      Left            =   8795
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   186
      Top             =   1657
      Width           =   1695
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Druckstellen"
      Height          =   285
      Index           =   91
      Left            =   10530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   188
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
      TabIndex        =   190
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
      TabIndex        =   192
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
      TabIndex        =   194
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
      TabIndex        =   196
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
      TabIndex        =   198
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
      TabIndex        =   200
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
      TabIndex        =   202
      Top             =   2848
      Width           =   1830
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Bewegungseinschränkungen"
      Height          =   285
      Index           =   99
      Left            =   9808
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   204
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
      TabIndex        =   206
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
      TabIndex        =   208
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
      TabIndex        =   210
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
      TabIndex        =   212
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
      TabIndex        =   214
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
      TabIndex        =   216
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
      TabIndex        =   218
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
      TabIndex        =   220
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
      TabIndex        =   222
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
      TabIndex        =   224
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
      TabIndex        =   226
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
      TabIndex        =   228
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
      TabIndex        =   230
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
      TabIndex        =   232
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
      TabIndex        =   234
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
      TabIndex        =   236
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
      TabIndex        =   238
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
      TabIndex        =   240
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
      TabIndex        =   242
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
      TabIndex        =   244
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
      TabIndex        =   246
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
      TabIndex        =   248
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
      TabIndex        =   250
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
      TabIndex        =   252
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
      TabIndex        =   254
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
      TabIndex        =   256
      Top             =   7937
      Width           =   7128
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Herz"
      Height          =   238
      Index           =   126
      Left            =   8503
      TabIndex        =   258
      Top             =   8505
      Width           =   2655
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Lunge"
      Height          =   238
      Index           =   127
      Left            =   13548
      TabIndex        =   260
      Top             =   8505
      Width           =   2040
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Bauch"
      Height          =   238
      Index           =   128
      Left            =   8616
      TabIndex        =   262
      Top             =   8755
      Width           =   1968
   End
   Begin VB.TextBox vTextB 
      DataField       =   "WS"
      Height          =   238
      Index           =   129
      Left            =   10998
      TabIndex        =   264
      Top             =   8755
      Width           =   1035
   End
   Begin VB.TextBox vTextB 
      DataField       =   "NL"
      Height          =   238
      Index           =   130
      Left            =   12302
      TabIndex        =   266
      Top             =   8755
      Width           =   1080
   End
   Begin VB.TextBox vTextB 
      DataField       =   "SD"
      Height          =   238
      Index           =   131
      Left            =   8220
      TabIndex        =   268
      Top             =   8993
      Width           =   1200
   End
   Begin VB.TextBox vTextB 
      DataField       =   "NNH"
      Height          =   238
      Index           =   132
      Left            =   9874
      TabIndex        =   270
      Top             =   8993
      Width           =   1290
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Zähne"
      Height          =   238
      Index           =   133
      Left            =   11792
      TabIndex        =   272
      Top             =   8993
      Width           =   1023
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Mundhöhle"
      Height          =   238
      Index           =   134
      Left            =   8957
      TabIndex        =   274
      Top             =   9243
      Width           =   1260
   End
   Begin VB.TextBox vTextB 
      DataField       =   "LK"
      Height          =   238
      Index           =   135
      Left            =   8390
      TabIndex        =   276
      Top             =   9581
      Width           =   1380
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Neuro sonst"
      Height          =   238
      Index           =   136
      Left            =   8731
      TabIndex        =   278
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
      TabIndex        =   280
      Top             =   9751
      Width           =   4428
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Schulung"
      Height          =   256
      Index           =   138
      Left            =   13889
      TabIndex        =   282
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
      TabIndex        =   284
      Top             =   10828
      Width           =   900
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Ausgabe"
      Height          =   291
      Index           =   1
      Left            =   15583
      TabIndex        =   285
      Top             =   566
      Width           =   1551
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "HANr"
      Height          =   240
      Index           =   140
      Left            =   10148
      TabIndex        =   287
      Top             =   10885
      Width           =   1176
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Beinbefund"
      Height          =   240
      Index           =   141
      Left            =   9015
      TabIndex        =   289
      Top             =   5839
      Width           =   6576
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H8000000F&
      Height          =   238
      Index           =   142
      Left            =   5272
      TabIndex        =   291
      Top             =   3005
      Width           =   636
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "letzte Änderung"
      Height          =   227
      Index           =   143
      Left            =   17064
      TabIndex        =   294
      Top             =   0
      Width           =   1349
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Fußpflege"
      Height          =   240
      Index           =   144
      Left            =   13152
      TabIndex        =   296
      Top             =   1701
      Width           =   696
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Podologie"
      Height          =   240
      Index           =   145
      Left            =   13889
      TabIndex        =   298
      Top             =   1701
      Width           =   696
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Einlagen"
      Height          =   240
      Index           =   146
      Left            =   14626
      TabIndex        =   300
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
      TabIndex        =   302
      Top             =   850
      Width           =   13215
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   227
      Index           =   148
      Left            =   10374
      TabIndex        =   303
      Top             =   10601
      Width           =   8148
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Vorgestellt"
      Height          =   238
      Index           =   149
      Left            =   2834
      TabIndex        =   305
      Top             =   340
      Width           =   1011
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "RRTurboMed"
      Height          =   240
      Index           =   150
      Left            =   8447
      TabIndex        =   307
      Top             =   8220
      Width           =   7131
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Versicherung"
      Height          =   240
      Index           =   151
      Left            =   9070
      TabIndex        =   309
      Top             =   10318
      Width           =   3846
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Titel"
      Height          =   240
      Index           =   152
      Left            =   4932
      TabIndex        =   310
      Top             =   56
      Width           =   681
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Anrede"
      Height          =   240
      Index           =   153
      Left            =   5669
      TabIndex        =   311
      Top             =   56
      Width           =   801
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FF0080&
      DataField       =   "Versicherungsart"
      Height          =   240
      Index           =   154
      Left            =   6576
      TabIndex        =   312
      Top             =   56
      Width           =   336
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit vormittags"
      Height          =   244
      Index           =   155
      Left            =   3118
      TabIndex        =   314
      Top             =   5160
      Width           =   516
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Essenszeit nachmittags"
      Height          =   244
      Index           =   156
      Left            =   4988
      TabIndex        =   316
      Top             =   5160
      Width           =   681
   End
   Begin VB.TextBox vTextB 
      DataField       =   "BZMessungen selbst"
      Height          =   244
      Index           =   157
      Left            =   1870
      TabIndex        =   318
      Top             =   6720
      Width           =   336
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Gerät"
      Height          =   244
      Index           =   158
      Left            =   2890
      TabIndex        =   320
      Top             =   6720
      Width           =   2091
   End
   Begin VB.TextBox vTextB 
      DataField       =   "UZ Tageszeit"
      Height          =   244
      Index           =   159
      Left            =   6236
      TabIndex        =   322
      Top             =   7258
      Width           =   1701
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0000FF00&
      DataField       =   "Nachname"
      Height          =   238
      Index           =   160
      Left            =   680
      TabIndex        =   323
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
      TabIndex        =   325
      Top             =   9243
      Width           =   4191
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Carotiden"
      Height          =   238
      Index           =   162
      Left            =   11632
      TabIndex        =   327
      Top             =   8505
      Width           =   1230
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "DMSchulz"
      Height          =   256
      Index           =   163
      Left            =   15817
      TabIndex        =   329
      Top             =   10320
      Width           =   510
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "RRSchulz"
      Height          =   256
      Index           =   164
      Left            =   17351
      TabIndex        =   331
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
      TabIndex        =   333
      Top             =   10828
      Width           =   900
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Tkz"
      DataField       =   "Tkz"
      Height          =   225
      Index           =   5
      Left            =   7426
      TabIndex        =   335
      Top             =   56
      Width           =   270
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Ther1"
      Height          =   225
      Index           =   166
      Left            =   5782
      TabIndex        =   337
      Top             =   626
      Width           =   750
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00808080&
      Height          =   239
      Index           =   167
      Left            =   11338
      TabIndex        =   338
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
      TabIndex        =   339
      Top             =   5499
      Width           =   13094
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   227
      Index           =   169
      Left            =   12812
      TabIndex        =   340
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
      TabIndex        =   341
      Top             =   11168
      Width           =   6353
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "An1Aufruf"
      Height          =   276
      Index           =   2
      Left            =   56
      TabIndex        =   342
      Top             =   10885
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "An2Aufruf"
      Height          =   291
      Index           =   3
      Left            =   56
      TabIndex        =   343
      Top             =   11221
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "AnAAufruf"
      Height          =   321
      Index           =   4
      Left            =   3571
      TabIndex        =   344
      Top             =   10840
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "CheckAufruf"
      Height          =   291
      Index           =   5
      Left            =   3571
      TabIndex        =   345
      Top             =   11225
      Width           =   2826
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obAn1eing"
      DataField       =   "obAn1eing"
      Height          =   240
      Index           =   6
      Left            =   2891
      TabIndex        =   346
      Top             =   10885
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obAn2eing"
      DataField       =   "obAn2eing"
      Height          =   240
      Index           =   7
      Left            =   2891
      TabIndex        =   347
      Top             =   11245
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obAnAeing"
      DataField       =   "obAnAeing"
      Height          =   240
      Index           =   8
      Left            =   6462
      TabIndex        =   348
      Top             =   10885
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obCheck"
      DataField       =   "obCheck"
      Height          =   240
      Index           =   9
      Left            =   6462
      TabIndex        =   349
      Top             =   11225
      Width           =   260
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Labor"
      Height          =   336
      Index           =   6
      Left            =   6776
      TabIndex        =   350
      Top             =   10828
      Width           =   1071
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Brief"
      Height          =   290
      Index           =   7
      Left            =   6803
      TabIndex        =   351
      Top             =   11225
      Width           =   1011
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Dokumente"
      Height          =   336
      Index           =   8
      Left            =   56
      TabIndex        =   352
      Top             =   11565
      Width           =   696
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Briefe"
      Height          =   336
      Index           =   9
      Left            =   1190
      TabIndex        =   353
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   1
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   354
      Top             =   5511
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   2
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   355
      Top             =   5707
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   3
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   356
      Top             =   5903
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   4
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   357
      Top             =   6098
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   5
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   358
      Top             =   6294
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   6
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   359
      Top             =   6490
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   7
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   360
      Top             =   6686
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   8
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   361
      Top             =   6882
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   9
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   362
      Top             =   7078
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   10
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   363
      Top             =   7274
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   11
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   364
      Top             =   7470
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   12
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   365
      Top             =   7666
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   13
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   366
      Top             =   7862
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   14
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   367
      Top             =   8058
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   15
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   368
      Top             =   8254
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   16
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   369
      Top             =   8450
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   17
      Left            =   15649
      Style           =   1  'Grafisch
      TabIndex        =   370
      Top             =   8646
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obCheck"
      Height          =   193
      Index           =   18
      Left            =   15647
      Style           =   1  'Grafisch
      TabIndex        =   371
      Top             =   8842
      Width           =   113
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "DReset"
      Height          =   276
      Index           =   10
      Left            =   17971
      TabIndex        =   372
      Top             =   10318
      Width           =   576
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obBZausgew"
      DataField       =   "obBZausgew"
      Height          =   240
      Index           =   10
      Left            =   14400
      TabIndex        =   374
      Top             =   10830
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obOSaufgek"
      DataField       =   "obOSaufgek"
      Height          =   240
      Index           =   11
      Left            =   14400
      TabIndex        =   376
      Top             =   11287
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obPodAufgek"
      DataField       =   "obPodAufgek"
      Height          =   240
      Index           =   12
      Left            =   14400
      TabIndex        =   378
      Top             =   11532
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obMBlAusgeh"
      DataField       =   "obMBlAusgeh"
      Height          =   240
      Index           =   13
      Left            =   14400
      TabIndex        =   380
      Top             =   11080
      Width           =   260
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "obDMPaufgekl"
      Height          =   238
      Index           =   171
      Left            =   16157
      TabIndex        =   382
      Top             =   11533
      Width           =   405
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "obSchulaufgek"
      Height          =   238
      Index           =   172
      Left            =   16157
      TabIndex        =   384
      Top             =   11281
      Width           =   405
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obSchulaufgek"
      Height          =   193
      Index           =   19
      Left            =   15647
      Style           =   1  'Grafisch
      TabIndex        =   385
      Top             =   9038
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obSchulaufgek"
      Height          =   193
      Index           =   20
      Left            =   15647
      Style           =   1  'Grafisch
      TabIndex        =   386
      Top             =   9234
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obSchulaufgek"
      Height          =   193
      Index           =   21
      Left            =   15647
      Style           =   1  'Grafisch
      TabIndex        =   387
      Top             =   9430
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obSchulaufgek"
      Height          =   193
      Index           =   22
      Left            =   15647
      Style           =   1  'Grafisch
      TabIndex        =   388
      Top             =   9626
      Width           =   113
   End
   Begin VB.CheckBox vOptionB 
      DataField       =   "obSchulaufgek"
      Height          =   193
      Index           =   23
      Left            =   15647
      Style           =   1  'Grafisch
      TabIndex        =   389
      Top             =   9822
      Width           =   113
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "TherAkt"
      Height          =   238
      Index           =   173
      Left            =   6973
      TabIndex        =   391
      Top             =   623
      Width           =   780
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Versicherungsart"
      Height          =   226
      Index           =   174
      Left            =   7313
      TabIndex        =   393
      Top             =   340
      Width           =   454
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "obMedNetz"
      DataField       =   "obMedNetz"
      Height          =   240
      Index           =   14
      Left            =   7483
      TabIndex        =   395
      Top             =   907
      Width           =   260
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "BZKurven"
      Height          =   336
      Index           =   11
      Left            =   3344
      TabIndex        =   396
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "DMPAusgeb"
      Height          =   290
      Index           =   12
      Left            =   17177
      TabIndex        =   397
      Top             =   566
      Width           =   1296
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "AugenBefunde"
      Height          =   337
      Index           =   13
      Left            =   5499
      TabIndex        =   398
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "DokDown"
      Height          =   291
      Index           =   14
      Left            =   793
      TabIndex        =   399
      Top             =   11565
      Width           =   336
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   227
      Index           =   175
      Left            =   11735
      TabIndex        =   400
      Top             =   10885
      Width           =   1032
   End
   Begin VB.ComboBox vComboB 
      Height          =   284
      Index           =   1
      Left            =   8569
      TabIndex        =   401
      Top             =   11508
      Width           =   2599
   End
   Begin VB.ComboBox vComboB 
      Height          =   286
      Index           =   2
      Left            =   11791
      TabIndex        =   402
      Top             =   11508
      Width           =   2551
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HA1nicht"
      Height          =   240
      Index           =   15
      Left            =   8333
      TabIndex        =   404
      Top             =   11792
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HA2nicht"
      Height          =   240
      Index           =   16
      Left            =   11451
      TabIndex        =   406
      Top             =   11735
      Width           =   260
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "HA1_Bez"
      Height          =   285
      Index           =   15
      Left            =   7990
      TabIndex        =   407
      Top             =   11508
      Width           =   570
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "HA2_Bez"
      Height          =   270
      Index           =   16
      Left            =   11168
      TabIndex        =   408
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
      TabIndex        =   410
      Top             =   10885
      Width           =   1984
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   680
      Index           =   177
      Left            =   4315
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   411
      Top             =   11962
      Width           =   4127
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   725
      Index           =   178
      Left            =   9139
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   413
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
      TabIndex        =   416
      Top             =   11905
      Width           =   4350
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   680
      Index           =   180
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   417
      Top             =   11962
      Width           =   3635
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   226
      Index           =   181
      Left            =   4025
      TabIndex        =   419
      Top             =   340
      Width           =   1190
   End
   Begin VB.ComboBox vComboB 
      Height          =   281
      Index           =   3
      Left            =   1231
      TabIndex        =   421
      Top             =   12642
      Width           =   3424
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Unausgefülle"
      Height          =   291
      Index           =   17
      Left            =   4705
      TabIndex        =   422
      Top             =   12642
      Width           =   1131
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "alleDatensätze"
      Height          =   291
      Index           =   18
      Left            =   7256
      TabIndex        =   423
      Top             =   12642
      Width           =   1251
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "UnausgefNZ"
      Height          =   291
      Index           =   19
      Left            =   5839
      TabIndex        =   424
      Top             =   12642
      Width           =   1416
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      Height          =   227
      Index           =   182
      Left            =   8503
      TabIndex        =   425
      Top             =   12642
      Width           =   2667
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Datenbank-Aufruf"
      Height          =   291
      Index           =   20
      Left            =   11225
      TabIndex        =   426
      Top             =   12642
      Width           =   7311
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   28860
      TabIndex        =   0
      Top             =   12930
      Width           =   28860
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Hinzufügen"
         Height          =   300
         Left            =   59
         TabIndex        =   427
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Bearbeiten"
         Height          =   300
         Left            =   1213
         TabIndex        =   428
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Löschen"
         Height          =   300
         Left            =   2367
         TabIndex        =   429
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "neu lesen"
         Height          =   300
         Left            =   3521
         TabIndex        =   430
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "S&chließen"
         Height          =   300
         Left            =   4675
         TabIndex        =   431
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Aktualisieren"
         Height          =   300
         Left            =   59
         TabIndex        =   432
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Abbrechen"
         Height          =   300
         Left            =   1213
         TabIndex        =   433
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   28860
      TabIndex        =   434
      Top             =   13230
      Width           =   28860
      Begin VB.CommandButton Suchen 
         Caption         =   "&Suchen"
         Height          =   255
         Left            =   4920
         TabIndex        =   440
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox suche 
         Height          =   285
         Left            =   5640
         TabIndex        =   441
         Top             =   0
         Width           =   2655
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "AnBog.frx":0000
         Style           =   1  'Grafisch
         TabIndex        =   435
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&ü"
         Height          =   300
         Left            =   345
         Picture         =   "AnBog.frx":0342
         Style           =   1  'Grafisch
         TabIndex        =   436
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&ä"
         Height          =   300
         Left            =   4080
         Picture         =   "AnBog.frx":0684
         Style           =   1  'Grafisch
         TabIndex        =   437
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   465
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "AnBog.frx":09C6
         Style           =   1  'Grafisch
         TabIndex        =   438
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   690
         TabIndex        =   439
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "GebDat"
      Height          =   238
      Index           =   1
      Left            =   113
      TabIndex        =   3
      Top             =   340
      Width           =   687
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Diabetestyp&:"
      Height          =   240
      Index           =   2
      Left            =   114
      TabIndex        =   5
      Top             =   639
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "Diabetes seit"
      Height          =   240
      Index           =   3
      Left            =   1299
      TabIndex        =   7
      Top             =   639
      Width           =   1125
   End
   Begin VB.Label vLab 
      Caption         =   "Tabletten seit"
      Height          =   240
      Index           =   4
      Left            =   2484
      TabIndex        =   9
      Top             =   639
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "Insulin seit"
      Height          =   240
      Index           =   5
      Left            =   3669
      TabIndex        =   11
      Top             =   639
      Width           =   1023
   End
   Begin VB.Label vLab 
      Caption         =   "Grund für Vorstellung"
      Height          =   240
      Index           =   6
      Left            =   114
      TabIndex        =   13
      Top             =   1164
      Width           =   1578
   End
   Begin VB.Label vLab 
      Caption         =   "Familienanamnese"
      Height          =   240
      Index           =   7
      Left            =   114
      TabIndex        =   15
      Top             =   2049
      Width           =   7638
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Gr&öße"
      Height          =   238
      Index           =   8
      Left            =   113
      TabIndex        =   17
      Top             =   3005
      Width           =   570
   End
   Begin VB.Label vLab 
      Caption         =   "Gewicht"
      Height          =   238
      Index           =   9
      Left            =   1701
      TabIndex        =   19
      Top             =   3005
      Width           =   690
   End
   Begin VB.Label vLab 
      Caption         =   "Tendenz"
      Height          =   238
      Index           =   10
      Left            =   3288
      TabIndex        =   21
      Top             =   3005
      Width           =   750
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Diab&.Med. 1"
      Height          =   256
      Index           =   11
      Left            =   56
      TabIndex        =   23
      Top             =   3344
      Width           =   963
   End
   Begin VB.Label vLab 
      Caption         =   "Menge"
      Height          =   256
      Index           =   12
      Left            =   4762
      TabIndex        =   25
      Top             =   3346
      Width           =   615
   End
   Begin VB.Label vLab 
      Caption         =   "2"
      Height          =   256
      Index           =   13
      Left            =   56
      TabIndex        =   27
      Top             =   3614
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "Menge"
      Height          =   256
      Index           =   14
      Left            =   4705
      TabIndex        =   29
      Top             =   3616
      Width           =   663
   End
   Begin VB.Label vLab 
      Caption         =   "3"
      Height          =   256
      Index           =   15
      Left            =   56
      TabIndex        =   31
      Top             =   3883
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "Menge"
      Height          =   256
      Index           =   16
      Left            =   4762
      TabIndex        =   33
      Top             =   3885
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "4"
      Height          =   256
      Index           =   17
      Left            =   56
      TabIndex        =   35
      Top             =   4181
      Width           =   633
   End
   Begin VB.Label vLab 
      Caption         =   "Menge"
      Height          =   256
      Index           =   18
      Left            =   4648
      TabIndex        =   37
      Top             =   4183
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Insulinpumpe"
      Height          =   240
      Index           =   19
      Left            =   56
      TabIndex        =   39
      Top             =   4535
      Width           =   1005
   End
   Begin VB.Label vLab 
      Caption         =   "seit"
      Height          =   240
      Index           =   20
      Left            =   1303
      TabIndex        =   41
      Top             =   4535
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "Marke"
      Height          =   240
      Index           =   21
      Left            =   3061
      TabIndex        =   43
      Top             =   4479
      Width           =   573
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "BE &gesamt"
      Height          =   244
      Index           =   22
      Left            =   0
      TabIndex        =   45
      Top             =   4876
      Width           =   855
   End
   Begin VB.Label vLab 
      Caption         =   "früh"
      Height          =   244
      Index           =   23
      Left            =   1587
      TabIndex        =   47
      Top             =   4875
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "vormittags"
      Height          =   244
      Index           =   24
      Left            =   2777
      TabIndex        =   49
      Top             =   4875
      Width           =   348
   End
   Begin VB.Label vLab 
      Caption         =   "mittags"
      Height          =   244
      Index           =   25
      Left            =   3685
      TabIndex        =   51
      Top             =   4875
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "nachmittags"
      Height          =   244
      Index           =   26
      Left            =   4705
      TabIndex        =   53
      Top             =   4875
      Width           =   285
   End
   Begin VB.Label vLab 
      Caption         =   "abends"
      Height          =   244
      Index           =   27
      Left            =   5669
      TabIndex        =   55
      Top             =   4875
      Width           =   348
   End
   Begin VB.Label vLab 
      Caption         =   "spät"
      Height          =   244
      Index           =   28
      Left            =   6916
      TabIndex        =   57
      Top             =   4762
      Width           =   345
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Essenszeit  fr&üh:"
      Height          =   244
      Index           =   29
      Left            =   632
      TabIndex        =   59
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label vLab 
      Caption         =   "mittags"
      Height          =   244
      Index           =   30
      Left            =   3685
      TabIndex        =   61
      Top             =   5160
      Width           =   285
   End
   Begin VB.Label vLab 
      Caption         =   "abends"
      Height          =   244
      Index           =   31
      Left            =   5669
      TabIndex        =   63
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "spät"
      Height          =   244
      Index           =   32
      Left            =   6916
      TabIndex        =   65
      Top             =   5047
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "Spritz-Eß-Abstand früh:"
      Height          =   244
      Index           =   33
      Left            =   56
      TabIndex        =   67
      Top             =   5442
      Width           =   1818
   End
   Begin VB.Label vLab 
      Caption         =   "mittags"
      Height          =   244
      Index           =   34
      Left            =   3685
      TabIndex        =   69
      Top             =   5443
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "abends"
      Height          =   244
      Index           =   35
      Left            =   5669
      TabIndex        =   71
      Top             =   5443
      Width           =   330
   End
   Begin VB.Label vLab 
      Caption         =   "Spritzstelle früh (B, O, A)"
      Height          =   244
      Index           =   36
      Left            =   56
      TabIndex        =   73
      Top             =   5727
      Width           =   1785
   End
   Begin VB.Label vLab 
      Caption         =   "mittags"
      Height          =   244
      Index           =   37
      Left            =   3685
      TabIndex        =   75
      Top             =   5727
      Width           =   288
   End
   Begin VB.Label vLab 
      Caption         =   "abends"
      Height          =   244
      Index           =   38
      Left            =   5669
      TabIndex        =   77
      Top             =   5727
      Width           =   345
   End
   Begin VB.Label vLab 
      Caption         =   "nachts"
      Height          =   244
      Index           =   39
      Left            =   6859
      TabIndex        =   79
      Top             =   5614
      Width           =   405
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Jahr letzte Diabetesschulung"
      Height          =   240
      Index           =   40
      Left            =   56
      TabIndex        =   81
      Top             =   6122
      Width           =   2115
   End
   Begin VB.Label vLab 
      Caption         =   "Ort Schulung"
      Height          =   240
      Index           =   41
      Left            =   4365
      TabIndex        =   83
      Top             =   6122
      Width           =   738
   End
   Begin VB.Label vLab 
      Caption         =   "letztes HbA1c"
      Height          =   244
      Index           =   42
      Left            =   56
      TabIndex        =   85
      Top             =   6405
      Width           =   960
   End
   Begin VB.Label vLab 
      Caption         =   "gem am"
      Height          =   244
      Index           =   43
      Left            =   2097
      TabIndex        =   87
      Top             =   6405
      Width           =   615
   End
   Begin VB.Label vLab 
      Caption         =   "vorherige Werte"
      Height          =   244
      Index           =   44
      Left            =   4308
      TabIndex        =   89
      Top             =   6405
      Width           =   1170
   End
   Begin VB.Label vLab 
      Caption         =   "Zahl pW"
      Height          =   244
      Index           =   45
      Left            =   113
      TabIndex        =   91
      Top             =   7003
      Width           =   675
   End
   Begin VB.Label vLab 
      Caption         =   "davon ndE"
      Height          =   244
      Index           =   46
      Left            =   2324
      TabIndex        =   93
      Top             =   7003
      Width           =   843
   End
   Begin VB.Label vLab 
      Caption         =   "davon nachts"
      Height          =   244
      Index           =   47
      Left            =   4195
      TabIndex        =   95
      Top             =   7003
      Width           =   1020
   End
   Begin VB.Label vLab 
      Caption         =   "Aufschreiben"
      Height          =   244
      Index           =   48
      Left            =   5159
      TabIndex        =   97
      Top             =   6720
      Width           =   1020
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "BZ &v d Essen"
      Height          =   244
      Index           =   49
      Left            =   113
      TabIndex        =   99
      Top             =   7371
      Width           =   1020
   End
   Begin VB.Label vLab 
      Caption         =   "danach"
      Height          =   244
      Index           =   50
      Left            =   2834
      TabIndex        =   101
      Top             =   7371
      Width           =   633
   End
   Begin VB.Label vLab 
      Caption         =   "UZ pM"
      Height          =   244
      Index           =   51
      Left            =   113
      TabIndex        =   103
      Top             =   7653
      Width           =   525
   End
   Begin VB.Label vLab 
      Caption         =   "rechtzeitig"
      Height          =   244
      Index           =   52
      Left            =   1246
      TabIndex        =   105
      Top             =   7653
      Width           =   675
   End
   Begin VB.Label vLab 
      Caption         =   "Fremde Hilfe pa"
      Height          =   244
      Index           =   53
      Left            =   2381
      TabIndex        =   107
      Top             =   7653
      Width           =   570
   End
   Begin VB.Label vLab 
      Caption         =   "Bewußtl pa"
      Height          =   244
      Index           =   54
      Left            =   3628
      TabIndex        =   109
      Top             =   7653
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "Keto pa"
      Height          =   244
      Index           =   55
      Left            =   4762
      TabIndex        =   111
      Top             =   7653
      Width           =   633
   End
   Begin VB.Label vLab 
      Caption         =   "gr300 pM"
      Height          =   244
      Index           =   56
      Left            =   5952
      TabIndex        =   113
      Top             =   7653
      Width           =   735
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Bl&uthochdruck"
      Height          =   240
      Index           =   57
      Left            =   56
      TabIndex        =   115
      Top             =   8447
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "BHD seit"
      Height          =   240
      Index           =   58
      Left            =   850
      TabIndex        =   117
      Top             =   8447
      Width           =   1125
   End
   Begin VB.Label vLab 
      Caption         =   "BHD beh mit"
      Height          =   240
      Index           =   59
      Left            =   2035
      TabIndex        =   119
      Top             =   8447
      Width           =   3000
   End
   Begin VB.Label vLab 
      Caption         =   "Blutdruckwerte"
      Height          =   240
      Index           =   60
      Left            =   6039
      TabIndex        =   121
      Top             =   8447
      Width           =   1848
   End
   Begin VB.Label vLab 
      Caption         =   "BDselbst"
      Height          =   240
      Index           =   61
      Left            =   5102
      TabIndex        =   123
      Top             =   8447
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Schwanger"
      Height          =   244
      Index           =   62
      Left            =   113
      TabIndex        =   125
      Top             =   8050
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Schwanger seit"
      Height          =   244
      Index           =   63
      Left            =   1757
      TabIndex        =   127
      Top             =   8050
      Width           =   1185
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Augens&p zuletzt"
      Height          =   240
      Index           =   64
      Left            =   60
      TabIndex        =   129
      Top             =   9074
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "Augensp Befund"
      Height          =   240
      Index           =   65
      Left            =   1256
      TabIndex        =   131
      Top             =   9070
      Width           =   3453
   End
   Begin VB.Label vLab 
      Caption         =   "Netzhaut gelasert"
      Height          =   240
      Index           =   66
      Left            =   4762
      TabIndex        =   133
      Top             =   9070
      Width           =   1260
   End
   Begin VB.Label vLab 
      Caption         =   "Sehminderung unbehebbar"
      Height          =   240
      Index           =   67
      Left            =   6055
      TabIndex        =   135
      Top             =   9070
      Width           =   1800
   End
   Begin VB.Label vLab 
      Caption         =   "Diabet Nierenschaden"
      Height          =   240
      Index           =   68
      Left            =   56
      TabIndex        =   137
      Top             =   9693
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Albumin zuletzt"
      Height          =   240
      Index           =   69
      Left            =   1020
      TabIndex        =   139
      Top             =   9693
      Width           =   1125
   End
   Begin VB.Label vLab 
      Caption         =   "erhöht?"
      Height          =   240
      Index           =   70
      Left            =   2205
      TabIndex        =   141
      Top             =   9693
      Width           =   1185
   End
   Begin VB.Label vLab 
      Caption         =   "Dialyse"
      Height          =   240
      Index           =   71
      Left            =   3390
      TabIndex        =   143
      Top             =   9693
      Width           =   288
   End
   Begin VB.Label vLab 
      Caption         =   "Dialyse seit"
      Height          =   240
      Index           =   72
      Left            =   3664
      TabIndex        =   145
      Top             =   9693
      Width           =   1155
   End
   Begin VB.Label vLab 
      Caption         =   "andere Nierenerkrankung"
      Height          =   240
      Index           =   73
      Left            =   4818
      TabIndex        =   147
      Top             =   9689
      Width           =   3030
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Herzkrankheit"
      Height          =   240
      Index           =   74
      Left            =   60
      TabIndex        =   149
      Top             =   10259
      Width           =   795
   End
   Begin VB.Label vLab 
      Caption         =   "Angina pectoris"
      Height          =   240
      Index           =   75
      Left            =   907
      TabIndex        =   151
      Top             =   10260
      Width           =   855
   End
   Begin VB.Label vLab 
      Caption         =   "Herzinfakt"
      Height          =   240
      Index           =   76
      Left            =   1757
      TabIndex        =   153
      Top             =   10260
      Width           =   573
   End
   Begin VB.Label vLab 
      Caption         =   "wann"
      Height          =   240
      Index           =   77
      Left            =   2324
      TabIndex        =   155
      Top             =   10260
      Width           =   675
   End
   Begin VB.Label vLab 
      Caption         =   "PTCA oder Stent"
      Height          =   240
      Index           =   78
      Left            =   3061
      TabIndex        =   157
      Top             =   10260
      Width           =   510
   End
   Begin VB.Label vLab 
      Caption         =   "Bypass kardial"
      Height          =   240
      Index           =   79
      Left            =   3571
      TabIndex        =   159
      Top             =   10260
      Width           =   480
   End
   Begin VB.Label vLab 
      Caption         =   "wann"
      Height          =   240
      Index           =   80
      Left            =   4081
      TabIndex        =   161
      Top             =   10261
      Width           =   525
   End
   Begin VB.Label vLab 
      Caption         =   "Herzschwäche"
      Height          =   240
      Index           =   81
      Left            =   4648
      TabIndex        =   163
      Top             =   10260
      Width           =   963
   End
   Begin VB.Label vLab 
      Caption         =   "Herzkrankheit Beschreibung"
      Height          =   240
      Index           =   82
      Left            =   5585
      TabIndex        =   165
      Top             =   10260
      Width           =   2265
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "H&irndurchblutungsstörung"
      Height          =   240
      Index           =   83
      Left            =   7880
      TabIndex        =   167
      Top             =   0
      Width           =   975
   End
   Begin VB.Label vLab 
      Caption         =   "Schlaganfall"
      Height          =   240
      Index           =   84
      Left            =   8900
      TabIndex        =   169
      Top             =   0
      Width           =   2310
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "B&eindurchblutungsstörung"
      Height          =   240
      Index           =   85
      Left            =   11225
      TabIndex        =   171
      Top             =   0
      Width           =   858
   End
   Begin VB.Label vLab 
      Caption         =   "Schaufensterkrankheit"
      Height          =   240
      Index           =   86
      Left            =   12132
      TabIndex        =   173
      Top             =   0
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Bypaß peripher"
      Height          =   240
      Index           =   87
      Left            =   13096
      TabIndex        =   175
      Top             =   0
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "Geschwür"
      Height          =   240
      Index           =   88
      Left            =   13322
      TabIndex        =   177
      Top             =   0
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Amputation"
      Height          =   240
      Index           =   89
      Left            =   14222
      TabIndex        =   179
      Top             =   0
      Width           =   2433
   End
   Begin VB.Label vLab 
      Caption         =   "pAVK Beschreibung"
      Height          =   240
      Index           =   90
      Left            =   7880
      TabIndex        =   181
      Top             =   510
      Width           =   7638
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ameisenlaufe&n"
      Height          =   240
      Index           =   91
      Left            =   7880
      TabIndex        =   183
      Top             =   1395
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "Ausmaß"
      Height          =   240
      Index           =   92
      Left            =   9080
      TabIndex        =   185
      Top             =   1417
      Width           =   675
   End
   Begin VB.Label vLab 
      Caption         =   "Druckstellen"
      Height          =   240
      Index           =   93
      Left            =   10091
      TabIndex        =   187
      Top             =   1417
      Width           =   960
   End
   Begin VB.Label vLab 
      Caption         =   "Verformungen"
      Height          =   240
      Index           =   94
      Left            =   11055
      TabIndex        =   189
      Top             =   1417
      Width           =   1020
   End
   Begin VB.Label vLab 
      Caption         =   "Beschreibung"
      Height          =   240
      Index           =   95
      Left            =   12132
      TabIndex        =   191
      Top             =   1417
      Width           =   828
   End
   Begin VB.Label vLab 
      Caption         =   "Neue Fußkomplikationen"
      Height          =   240
      Index           =   96
      Left            =   7937
      TabIndex        =   193
      Top             =   2041
      Width           =   2265
   End
   Begin VB.Label vLab 
      Caption         =   "Entleerungsst Magen"
      Height          =   240
      Index           =   97
      Left            =   10261
      TabIndex        =   195
      Top             =   2044
      Width           =   1575
   End
   Begin VB.Label vLab 
      Caption         =   "Harnblase"
      Height          =   240
      Index           =   98
      Left            =   11853
      TabIndex        =   197
      Top             =   2041
      Width           =   1155
   End
   Begin VB.Label vLab 
      Caption         =   "Schwindel Aufstehen"
      Height          =   240
      Index           =   99
      Left            =   13044
      TabIndex        =   199
      Top             =   2041
      Width           =   1248
   End
   Begin VB.Label vLab 
      Caption         =   "Folgeerkrankungen Haut"
      Height          =   240
      Index           =   100
      Left            =   7937
      TabIndex        =   201
      Top             =   2608
      Width           =   1830
   End
   Begin VB.Label vLab 
      Caption         =   "Bewegungseinschränkungen"
      Height          =   240
      Index           =   101
      Left            =   9808
      TabIndex        =   203
      Top             =   2608
      Width           =   2310
   End
   Begin VB.Label vLab 
      Caption         =   "Sexualstörung"
      Height          =   240
      Index           =   102
      Left            =   12188
      TabIndex        =   205
      Top             =   2608
      Width           =   1080
   End
   Begin VB.Label vLab 
      Caption         =   "Sexualstörung seit"
      Height          =   240
      Index           =   103
      Left            =   13268
      TabIndex        =   207
      Top             =   2608
      Width           =   1938
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Weitere Anamnese"
      Height          =   240
      Index           =   104
      Left            =   7937
      TabIndex        =   209
      Top             =   3175
      Width           =   7638
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Alk&ohol"
      Height          =   240
      Index           =   105
      Left            =   10318
      TabIndex        =   211
      Top             =   4082
      Width           =   690
   End
   Begin VB.Label vLab 
      Caption         =   "Tabak"
      Height          =   240
      Index           =   106
      Left            =   7937
      TabIndex        =   213
      Top             =   4082
      Width           =   633
   End
   Begin VB.Label vLab 
      Caption         =   "Weitere Medikation"
      Height          =   240
      Index           =   107
      Left            =   7937
      TabIndex        =   215
      Top             =   4365
      Width           =   1653
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Liph&ypertrophien Abdomen"
      Height          =   240
      Index           =   108
      Left            =   7937
      TabIndex        =   217
      Top             =   5272
      Width           =   2310
   End
   Begin VB.Label vLab 
      Caption         =   "Liphypertrophien Beine"
      Height          =   240
      Index           =   109
      Left            =   10247
      TabIndex        =   219
      Top             =   5272
      Width           =   2310
   End
   Begin VB.Label vLab 
      Caption         =   "Liphypertrophien Arme"
      Height          =   240
      Index           =   110
      Left            =   12557
      TabIndex        =   221
      Top             =   5272
      Width           =   3018
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Hyper&keratosen"
      Height          =   240
      Index           =   111
      Left            =   7993
      TabIndex        =   223
      Top             =   6179
      Width           =   3285
   End
   Begin VB.Label vLab 
      Caption         =   "Ulcera"
      Height          =   240
      Index           =   112
      Left            =   11338
      TabIndex        =   225
      Top             =   6179
      Width           =   1485
   End
   Begin VB.Label vLab 
      Caption         =   "Kraft Zehenheber"
      Height          =   240
      Index           =   113
      Left            =   12869
      TabIndex        =   227
      Top             =   6179
      Width           =   915
   End
   Begin VB.Label vLab 
      Caption         =   "Kraft Zehenbeuger"
      Height          =   240
      Index           =   114
      Left            =   13833
      TabIndex        =   229
      Top             =   6179
      Width           =   843
   End
   Begin VB.Label vLab 
      Caption         =   "Kraft Knie"
      Height          =   240
      Index           =   115
      Left            =   14740
      TabIndex        =   231
      Top             =   6179
      Width           =   855
   End
   Begin VB.Label vLab 
      Caption         =   "ASR"
      Height          =   240
      Index           =   116
      Left            =   7993
      TabIndex        =   233
      Top             =   6746
      Width           =   1080
   End
   Begin VB.Label vLab 
      Caption         =   "PSR"
      Height          =   240
      Index           =   117
      Left            =   9127
      TabIndex        =   235
      Top             =   6746
      Width           =   990
   End
   Begin VB.Label vLab 
      Caption         =   "Oberflächensensibilität"
      Height          =   240
      Index           =   118
      Left            =   10148
      TabIndex        =   237
      Top             =   6746
      Width           =   1638
   End
   Begin VB.Label vLab 
      Caption         =   "Monofilamenttest"
      Height          =   240
      Index           =   119
      Left            =   11848
      TabIndex        =   239
      Top             =   6746
      Width           =   1560
   End
   Begin VB.Label vLab 
      Caption         =   "Kalt-Warm"
      Height          =   240
      Index           =   120
      Left            =   13408
      TabIndex        =   241
      Top             =   6746
      Width           =   1560
   End
   Begin VB.Label vLab 
      Caption         =   "Vibration IK"
      Height          =   240
      Index           =   121
      Left            =   7993
      TabIndex        =   243
      Top             =   7313
      Width           =   1560
   End
   Begin VB.Label vLab 
      Caption         =   "Vibration Großzehe"
      Height          =   240
      Index           =   122
      Left            =   9581
      TabIndex        =   245
      Top             =   7313
      Width           =   1560
   End
   Begin VB.Label vLab 
      Caption         =   "Puls Leiste"
      Height          =   240
      Index           =   123
      Left            =   11225
      TabIndex        =   247
      Top             =   7313
      Width           =   1023
   End
   Begin VB.Label vLab 
      Caption         =   "Puls Kniekehle"
      Height          =   240
      Index           =   124
      Left            =   12302
      TabIndex        =   249
      Top             =   7313
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Puls Atp"
      Height          =   240
      Index           =   125
      Left            =   13266
      TabIndex        =   251
      Top             =   7313
      Width           =   1080
   End
   Begin VB.Label vLab 
      Caption         =   "Puls Adp"
      Height          =   240
      Index           =   126
      Left            =   14400
      TabIndex        =   253
      Top             =   7313
      Width           =   1140
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "&RR"
      Height          =   240
      Index           =   127
      Left            =   7993
      TabIndex        =   255
      Top             =   7937
      Width           =   393
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Her&z"
      Height          =   238
      Index           =   128
      Left            =   7993
      TabIndex        =   257
      Top             =   8505
      Width           =   450
   End
   Begin VB.Label vLab 
      Caption         =   "Lunge"
      Height          =   238
      Index           =   129
      Left            =   12925
      TabIndex        =   259
      Top             =   8505
      Width           =   555
   End
   Begin VB.Label vLab 
      Caption         =   "Bauch"
      Height          =   238
      Index           =   130
      Left            =   7993
      TabIndex        =   261
      Top             =   8755
      Width           =   573
   End
   Begin VB.Label vLab 
      Caption         =   "WS"
      Height          =   238
      Index           =   131
      Left            =   10658
      TabIndex        =   263
      Top             =   8755
      Width           =   285
   End
   Begin VB.Label vLab 
      Caption         =   "NL"
      Height          =   238
      Index           =   132
      Left            =   12076
      TabIndex        =   265
      Top             =   8755
      Width           =   240
   End
   Begin VB.Label vLab 
      Caption         =   "SD"
      Height          =   238
      Index           =   133
      Left            =   7993
      TabIndex        =   267
      Top             =   8993
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "NNH"
      Height          =   238
      Index           =   134
      Left            =   9467
      TabIndex        =   269
      Top             =   8993
      Width           =   390
   End
   Begin VB.Label vLab 
      Caption         =   "Zähne"
      Height          =   238
      Index           =   135
      Left            =   11225
      TabIndex        =   271
      Top             =   8993
      Width           =   573
   End
   Begin VB.Label vLab 
      Caption         =   "Mundhöhle"
      Height          =   238
      Index           =   136
      Left            =   7993
      TabIndex        =   273
      Top             =   9243
      Width           =   915
   End
   Begin VB.Label vLab 
      Caption         =   "LK"
      Height          =   238
      Index           =   137
      Left            =   7993
      TabIndex        =   275
      Top             =   9581
      Width           =   330
   End
   Begin VB.Label vLab 
      Caption         =   "Neuro sonst"
      Height          =   238
      Index           =   138
      Left            =   8050
      TabIndex        =   277
      Top             =   10035
      Width           =   618
   End
   Begin VB.Label vLab 
      Caption         =   "Weitere Befunde"
      Height          =   240
      Index           =   139
      Left            =   9807
      TabIndex        =   279
      Top             =   9751
      Width           =   1308
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "S&chulung"
      Height          =   256
      Index           =   140
      Left            =   13039
      TabIndex        =   281
      Top             =   10320
      Width           =   840
   End
   Begin VB.Label vLab 
      Caption         =   "DMP"
      Height          =   240
      Index           =   141
      Left            =   7993
      TabIndex        =   283
      Top             =   10601
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "HA"
      Height          =   240
      Index           =   142
      Left            =   9977
      TabIndex        =   286
      Top             =   10601
      Width           =   285
   End
   Begin VB.Label vLab 
      Caption         =   "Beinbefund"
      Height          =   240
      Index           =   143
      Left            =   7993
      TabIndex        =   288
      Top             =   5839
      Width           =   960
   End
   Begin VB.Label vLab 
      Caption         =   "BMI"
      Height          =   238
      Index           =   144
      Left            =   4762
      TabIndex        =   290
      Top             =   3005
      Width           =   450
   End
   Begin VB.Label vLab 
      Caption         =   "kg/m²"
      Height          =   238
      Index           =   145
      Left            =   5953
      TabIndex        =   292
      Top             =   3005
      Width           =   624
   End
   Begin VB.Label vLab 
      Caption         =   "letzte Änderung"
      Height          =   240
      Index           =   146
      Left            =   15817
      TabIndex        =   293
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label vLab 
      Caption         =   "Fußplfege"
      Height          =   240
      Index           =   147
      Left            =   13096
      TabIndex        =   295
      Top             =   1417
      Width           =   795
   End
   Begin VB.Label vLab 
      Caption         =   "Podologie"
      Height          =   240
      Index           =   148
      Left            =   13889
      TabIndex        =   297
      Top             =   1417
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Orthop.Einlg/Schuhe diab&'ger"
      Height          =   240
      Index           =   149
      Left            =   14683
      TabIndex        =   299
      Top             =   1417
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Diagnosen:"
      Height          =   240
      Index           =   150
      Left            =   14400
      TabIndex        =   301
      Top             =   1984
      Width           =   1125
   End
   Begin VB.Label vLab 
      Caption         =   "vorgestellt"
      Height          =   238
      Index           =   151
      Left            =   1927
      TabIndex        =   304
      Top             =   340
      Width           =   795
   End
   Begin VB.Label vLab 
      Caption         =   "RR TurboMed:"
      Height          =   240
      Index           =   152
      Left            =   7993
      TabIndex        =   306
      Top             =   8221
      Width           =   450
   End
   Begin VB.Label vLab 
      Caption         =   "Versicherung:"
      Height          =   240
      Index           =   153
      Left            =   8050
      TabIndex        =   308
      Top             =   10318
      Width           =   960
   End
   Begin VB.Label vLab 
      Caption         =   "vormittags"
      Height          =   244
      Index           =   154
      Left            =   2721
      TabIndex        =   313
      Top             =   5160
      Width           =   405
   End
   Begin VB.Label vLab 
      Caption         =   "nachmittags"
      Height          =   244
      Index           =   155
      Left            =   4705
      TabIndex        =   315
      Top             =   5160
      Width           =   270
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "BZ-&Messungen: selbst?"
      Height          =   244
      Index           =   156
      Left            =   56
      TabIndex        =   317
      Top             =   6720
      Width           =   1755
   End
   Begin VB.Label vLab 
      Caption         =   "Gerät:"
      Height          =   244
      Index           =   157
      Left            =   2323
      TabIndex        =   319
      Top             =   6720
      Width           =   525
   End
   Begin VB.Label vLab 
      Caption         =   "UZ Tageszeit"
      Height          =   244
      Index           =   158
      Left            =   5102
      TabIndex        =   321
      Top             =   7257
      Width           =   1065
   End
   Begin VB.Label vLab 
      BackColor       =   &H0000FFFF&
      Caption         =   "Beinödeme&, Venenkrkht:"
      Height          =   463
      Index           =   159
      Left            =   10261
      TabIndex        =   324
      Top             =   9243
      Width           =   1080
   End
   Begin VB.Label vLab 
      Caption         =   "Carotiden"
      Height          =   238
      Index           =   160
      Left            =   11225
      TabIndex        =   326
      Top             =   8505
      Width           =   390
   End
   Begin VB.Label vLab 
      Caption         =   "Dm-Schlgn:"
      Height          =   256
      Index           =   161
      Left            =   14850
      TabIndex        =   328
      Top             =   10320
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "RR-Schlgn:"
      Height          =   256
      Index           =   162
      Left            =   16384
      TabIndex        =   330
      Top             =   10318
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "DMP hier:"
      Height          =   240
      Index           =   163
      Left            =   8957
      TabIndex        =   332
      Top             =   10601
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "Tkz"
      Height          =   225
      Index           =   164
      Left            =   6973
      TabIndex        =   334
      Top             =   56
      Width           =   390
   End
   Begin VB.Label vLab 
      Caption         =   "Therapieart"
      Height          =   224
      Index           =   165
      Left            =   4818
      TabIndex        =   336
      Top             =   626
      Width           =   900
   End
   Begin VB.Label vLab 
      Caption         =   "BZ-Meßger ausgew"
      Height          =   210
      Index           =   166
      Left            =   14630
      TabIndex        =   373
      Top             =   10830
      Width           =   1470
   End
   Begin VB.Label vLab 
      Caption         =   "orthop. SM aufgekl."
      Height          =   225
      Index           =   167
      Left            =   14630
      TabIndex        =   375
      Top             =   11282
      Width           =   1470
   End
   Begin VB.Label vLab 
      Caption         =   "üb.Podologie aufgek."
      Height          =   240
      Index           =   168
      Left            =   14630
      TabIndex        =   377
      Top             =   11532
      Width           =   1530
   End
   Begin VB.Label vLab 
      Caption         =   "&8: Merkbl DFS ausgehändigt"
      Height          =   240
      Index           =   169
      Left            =   14626
      TabIndex        =   379
      Top             =   11080
      Width           =   1470
   End
   Begin VB.Label vLab 
      Caption         =   " DMP-Beteiligung"
      Height          =   210
      Index           =   170
      Left            =   16611
      TabIndex        =   381
      Top             =   11508
      Width           =   1140
   End
   Begin VB.Label vLab 
      Caption         =   "&9: Schulungsaufkl"
      Height          =   210
      Index           =   171
      Left            =   16611
      TabIndex        =   383
      Top             =   11281
      Width           =   1170
   End
   Begin VB.Label vLab 
      Caption         =   "-&>"
      Height          =   227
      Index           =   172
      Left            =   6689
      TabIndex        =   390
      Top             =   623
      Width           =   178
   End
   Begin VB.Label vLab 
      Caption         =   "SchGr"
      Height          =   227
      Index           =   173
      Left            =   6746
      TabIndex        =   392
      Top             =   340
      Width           =   518
   End
   Begin VB.Label vLab 
      Caption         =   "Med.Netz:"
      Height          =   240
      Index           =   174
      Left            =   6633
      TabIndex        =   394
      Top             =   907
      Width           =   795
   End
   Begin VB.Label vLab 
      Caption         =   "-"
      Height          =   180
      Index           =   175
      Left            =   8047
      TabIndex        =   403
      Top             =   11735
      Width           =   240
   End
   Begin VB.Label vLab 
      Caption         =   "-"
      Height          =   180
      Index           =   176
      Left            =   11168
      TabIndex        =   405
      Top             =   11735
      Width           =   225
   End
   Begin VB.Label vLab 
      Caption         =   "&+"
      Height          =   285
      Index           =   177
      Left            =   16157
      TabIndex        =   409
      Top             =   10885
      Width           =   255
   End
   Begin VB.Label vLab 
      Caption         =   "Doppler"
      Height          =   240
      Index           =   178
      Left            =   3628
      TabIndex        =   412
      Top             =   11962
      Width           =   630
   End
   Begin VB.Label vLab 
      Caption         =   "Duplex"
      Height          =   226
      Index           =   179
      Left            =   8447
      TabIndex        =   414
      Top             =   11905
      Width           =   628
   End
   Begin VB.Label vLab 
      Caption         =   "Sono"
      Height          =   227
      Index           =   180
      Left            =   13776
      TabIndex        =   415
      Top             =   11848
      Width           =   462
   End
   Begin VB.Label vLab 
      Caption         =   "-"
      Height          =   226
      Index           =   181
      Left            =   3855
      TabIndex        =   418
      Top             =   340
      Width           =   134
   End
   Begin VB.Label vLab 
      Caption         =   "Daten&quelle:"
      Height          =   240
      Index           =   182
      Left            =   113
      TabIndex        =   420
      Top             =   12642
      Width           =   1005
   End
End
Attribute VB_Name = "AnBog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents adoRS As Recordset
Attribute adoRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Public mbDataChanged As Boolean
Dim altsuche$, SuchStringGeändert%
Private Sub adoRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hierdurch wird die aktuelle Datensatzposition für diese Datensatzgruppe angezeigt
 lblStatus.Caption = CStr(adoRS.AbsolutePosition)
 Call do_Form_Current(Me)
End Sub 'adoRS_MoveComplete

Private Sub adoRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hier können Sie Code zur Überprüfung einfügen
  'Dieses Ereignis wird aufgerufen, wenn die folgenden Aktionen eintreten
  Dim bCancel As Boolean
  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select
  If bCancel Then adStatus = adStatusCancel
End Sub 'adoRS_WillChangeRecord

Private Sub cmdAdd_Click()
  On Error GoTo fehler
  With adoRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Datensatz; hinzufügen"
    mbAddNewFlag = True
    SetButtons False
  End With
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdAdd_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'cmdAdd_Click()

Private Sub cmdDelete_Click()
  On Error GoTo fehler
  With adoRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdDelete_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Private Sub cmdDelete_Click()

Private Sub cmdRefresh_Click()
  'Dies wird nur für Mehrbenutzeranwendungen benötigt
  On Error GoTo fehler
  adoRS.Requery
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdRefresh_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Private Sub cmdRefresh_Click()"

Private Sub cmdEdit_Click()
  On Error GoTo fehler
  lblStatus.Caption = "Datensatz bearbeiten"
  mbEditFlag = True
  SetButtons False
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdEdit_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'cmdEdit_Click()

Private Sub cmdCancel_Click()
  On Error Resume Next
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoRS.CancelUpdate
 If mvBookMark > 0 Then
    adoRS.Bookmark = mvBookMark
  Else
    adoRS.MoveFirst
  End If
  mbDataChanged = False
End Sub ' cmdCancel_Click()

Private Sub cmdUpdate_Click()
  On Error GoTo fehler
  adoRS.UpdateBatch adAffectAll
  If mbAddNewFlag Then
    adoRS.MoveLast              'Zu neuem Datensatz gehen
  End If
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdUpdate_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdUpdate_Click()

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo fehler
  adoRS.MoveFirst
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdFirst_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdFirst_Click

Private Sub cmdLast_Click()
  On Error GoTo fehler
  adoRS.MoveLast
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdLast_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdLast_Click()

Private Sub cmdNext_Click()
  On Error GoTo fehler
  If Not adoRS.EOF Then adoRS.MoveNext
  If adoRS.EOF And adoRS.RecordCount > 0 Then
    Beep
     'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoRS.MoveLast
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdNext_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdNext_Click()

Private Sub cmdPrevious_Click()
  On Error GoTo fehler
  If Not adoRS.BOF Then adoRS.MovePrevious
  If adoRS.BOF And adoRS.RecordCount > 0 Then
    Beep
    'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoRS.MoveFirst
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdPrevious_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdPrevious_Click()

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub ' SetButtons

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Call doKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub Suchen_Click()
 On Error GoTo fehler
 If SuchStringGeändert Then
  altsuche = Me.suche
  If IsNumeric(Me.suche) Then
   adoRS.Find " Pat_id = " & Me.suche, 0, adSearchForward
  Else
   If adoRS.EOF Then adoRS.MoveFirst
   adoRS.Find "gesname like '" & Me.suche & "*'", 1, adSearchForward
  End If
 Else
  Me.suche.SetFocus
  SuchStringGeändert = True
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Suchen_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'Suchen_Click()

Private Sub Datenbank_Aufruf_Click()
 Call do_Datenbank_Aufruf_Click(Me)
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
 Call doForm_Load(Me)
End Sub
Sub Form_Close()
  Call do_Form_Close(Me)
End Sub

Private Sub Hirndurchblutungsstörung_Exit(Cancel%)
' Call do_Hirndurchblutungsstörung_Exit(Cancel, Me)
End Sub

Function do_u_Click(frm As Form, nr%)
 Call dodo_u_Click(frm, nr)
End Function ' do_u_Click(frm As Form, Nr%)



Private Sub vTextB_Click(Index As Integer)
 Select Case Index
  Case 140 ' HANr
   Call do_ha_click(HANr)
  Case 148 ' HA
   ' Hier Auswahl
  Case 175 ' HANr2
   Call do_ha_click(HANr2)
 End Select ' Index
End Sub 'Private Sub vTextB_Click(Index as Integer)

Private Sub vTextB_LostFocus(Index As Integer)
 Select Case Index
  Case 82 ' Hirndurchblutungsstörung
  ' Dim Cancel%
  ' Call do_Hirndurchblutungsstörung_Exit(Cancel, Me)
 End Select ' Index
End Sub 'Private Sub vTextB_LostFocus(Index as Integer)

Private Sub vTextB_Change(Index As Integer)
 Select Case Index
  Case 82 ' Hirndurchblutungsstörung
  ' Dim Cancel%
  ' Call do_Hirndurchblutungsstörung_Exit(Cancel, Me)
  Case 148 ' HA
   Call do_HAC(ActiveControl, 1)
   Call do_HAC(ActiveControl, 2)
 End Select ' Index
End Sub 'Private Sub vTextB_Change(Index as Integer)

Private Sub vCommandB_Click(Index As Integer)
 Select Case Index
  Case 1 ' Ausgabe
   Call do_Ausgabe_Click(Me)
  Case 2 ' An1Aufruf
   Call do_An1Aufruf_click(Me)
  Case 3 ' An2Aufruf
   Call do_An2Aufruf_click(Me)
  Case 4 ' AnAAufruf
   Call do_AnAAufruf_click(Me)
  Case 5 ' CheckAufruf
   Call do_CheckAufruf_click(Me)
  Case 6 ' Labor
   Call do_Labor_Click(Me)
  Case 7 ' Brief
   Call tu_brief(Me)
   Call WD
   Call Sound(WinDir + "\media\Windows XP-Batterie niedrig.wav")
  Case 8 ' Dokumente
   Call do_Dokumente_Click(Me)
  Case 9 ' Briefe
   Call do_Briefe_Click(Me)
  Case 10 ' DReset
   Call do_Diagnosen_Reset(Me)
  Case 11 ' BZKurven
   Call do_BZKurven_Click(Me)
  Case 12 ' DMPAusgeb
   Call do_DMPAusgeb(Me)
  Case 13 ' AugenBefunde
   Call do_AugenBefunde_Click(Me)
  Case 14 ' DokDown
   Call do_DokDown(Me)
  Case 15 ' HA1_Bez
   Call do_ha_click(HA_Auswahl)
  Case 16 ' HA2_Bez
   Call do_ha_click(HA_Auswahl2)
  Case 17 ' Unausgefülle
   Call DqC_U(Me)
  Case 18 ' alleDatensätze
   Call DqC_A(Me)
  Case 19 ' UnausgefNZ
   Call DqC_UNZ(Me)
 End Select ' Index
End Sub 'Private Sub vCommandB_Click(Index as Integer)

Private Sub vCommandB_KeyPress(Index As Integer, KeyAscii As Integer)
 Select Case Index
  Case 6 ' Labor
    Call Taste(KeyAscii, 0, Me)
  Case 27
   Stop
 End Select ' Index
End Sub 'Private Sub vCommandB_KeyPress(Index as Integer, KeyAscii As Integer)

Private Sub vComboB_Change(Index As Integer)
 Select Case Index
  Case 3 ' Datenquelle
   Call do_Datenquelle_Change(Me)
 End Select ' Index
End Sub 'Private Sub vComboB_Change(Index as Integer)

Private Sub vOptionB_Click(Index As Integer)
 Select Case Index
  Case 1 ' U1
   Call do_u_Click(Me, 1)
  Case 2 ' U2
   Call do_u_Click(Me, 2)
  Case 3 ' U3
   Call do_u_Click(Me, 3)
  Case 4 ' U4
   Call do_u_Click(Me, 4)
  Case 5 ' U5
   Call do_u_Click(Me, 5)
  Case 6 ' U6
   Call do_u_Click(Me, 6)
  Case 7 ' U7
   Call do_u_Click(Me, 7)
  Case 8 ' U8
   Call do_u_Click(Me, 8)
  Case 9 ' U9
   Call do_u_Click(Me, 9)
  Case 10 ' U10
   Call do_u_Click(Me, 10)
  Case 11 ' U11
   Call do_u_Click(Me, 11)
  Case 12 ' U12
   Call do_u_Click(Me, 12)
  Case 13 ' U13
   Call do_u_Click(Me, 13)
  Case 14 ' U14
   Call do_u_Click(Me, 14)
  Case 15 ' U15
   Call do_u_Click(Me, 15)
  Case 16 ' U16
   Call do_u_Click(Me, 16)
  Case 17 ' U17
   Call do_u_Click(Me, 17)
  Case 18 ' U18
   Call do_u_Click(Me, 18)
  Case 19 ' U19
   Call do_u_Click(Me, 19)
  Case 20 ' U20
   Call do_u_Click(Me, 20)
  Case 21 ' U21
   Call do_u_Click(Me, 21)
  Case 22 ' U22
   Call do_u_Click(Me, 22)
  Case 23 ' U23
   Call do_u_Click(Me, 23)
 End Select ' Index
End Sub 'Private Sub vOptionB_Click(Index as Integer)


Private Sub Form_Resize()
  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub 'Private Sub Form_Unload()
' vLab(1)        :GebDat_Bezeichnungsfeld
' vLab(2)        :Diabetestyp_Bezeichnungsfeld
' vLab(3)        :Diabetes seit_Bezeichnungsfeld
' vLab(4)        :Tabletten seit_Bezeichnungsfeld
' vLab(5)        :Insulin seit_Bezeichnungsfeld
' vLab(6)        :Grund für Vorstellung_Bezeichnungsfeld
' vLab(7)        :Familienanamnese_Bezeichnungsfeld
' vLab(8)        :Größe_Bezeichnungsfeld
' vLab(9)        :Gewicht_Bezeichnungsfeld
' vLab(10)       :Tendenz_Bezeichnungsfeld
' vLab(11)       :DiabetesMedikament 1_Bezeichnungsfeld
' vLab(12)       :DiabetesMedikamente 1 Menge_Bezeichnungsfeld
' vLab(13)       :DiabetesMedikament 2_Bezeichnungsfeld
' vLab(14)       :DiabetesMedikament 2 Menge_Bezeichnungsfeld
' vLab(15)       :DiabetesMedikament 3_Bezeichnungsfeld
' vLab(16)       :DiabetesMedikament 3 Menge_Bezeichnungsfeld
' vLab(17)       :DiabetesMedikament 4_Bezeichnungsfeld
' vLab(18)       :DiabetesMedikament 4 Menge_Bezeichnungsfeld
' vLab(19)       :Insulinpumpe_Bezeichnungsfeld
' vLab(20)       :Insulinpumpe seit_Bezeichnungsfeld
' vLab(21)       :Insulinpumpe Marke_Bezeichnungsfeld
' vLab(22)       :Broteinheiten gesamt_Bezeichnungsfeld
' vLab(23)       :Broteinheiten früh_Bezeichnungsfeld
' vLab(24)       :Broteinheiten ZM früh_Bezeichnungsfeld
' vLab(25)       :Broteinheiten mittags_Bezeichnungsfeld
' vLab(26)       :Broteinheiten nachmittags_Bezeichnungsfeld
' vLab(27)       :Broteinheiten abends_Bezeichnungsfeld
' vLab(28)       :Broteinheiten nachts_Bezeichnungsfeld
' vLab(29)       :Essenszeit  früh_Bezeichnungsfeld
' vLab(30)       :Essenszeit mittags_Bezeichnungsfeld
' vLab(31)       :Essenszeit abends_Bezeichnungsfeld
' vLab(32)       :Essenszeit spät_Bezeichnungsfeld
' vLab(33)       :Spritz-Eß-Abstand früh_Bezeichnungsfeld
' vLab(34)       :Spritz-Eß-Abstand mittags_Bezeichnungsfeld
' vLab(35)       :Spritz-Eß-Abstand abends_Bezeichnungsfeld
' vLab(36)       :Spritzstelle früh (B, O, A)_Bezeichnungsfeld
' vLab(37)       :Spritzstelle mittags_Bezeichnungsfeld
' vLab(38)       :Spritzstelle abends_Bezeichnungsfeld
' vLab(39)       :Spritzstelle nachts_Bezeichnungsfeld
' vLab(40)       :Jahr letzte Diabetesschulung_Bezeichnungsfeld
' vLab(41)       :Ort Schulung_Bezeichnungsfeld
' vLab(42)       :letztes HbA1c_Bezeichnungsfeld
' vLab(43)       :gemessen am_Bezeichnungsfeld
' vLab(44)       :vorherige Werte_Bezeichnungsfeld
' vLab(45)       :BZMessungen pW_Bezeichnungsfeld
' vLab(46)       :BZMessungen pW ndE_Bezeichnungsfeld
' vLab(47)       :BZMessungen p W nachts_Bezeichnungsfeld
' vLab(48)       :Aufschreiben_Bezeichnungsfeld
' vLab(49)       :BZWerte v d Essen_Bezeichnungsfeld
' vLab(50)       :BZWerte n d Essen_Bezeichnungsfeld
' vLab(51)       :Unterzucker pM_Bezeichnungsfeld
' vLab(52)       :UZ rechtzeitig_Bezeichnungsfeld
' vLab(53)       :Fremde Hilfe pa_Bezeichnungsfeld
' vLab(54)       :Bewußtlos pa_Bezeichnungsfeld
' vLab(55)       :Keto pa_Bezeichnungsfeld
' vLab(56)       :BZgr300 pM_Bezeichnungsfeld
' vLab(57)       :Bluthochdruck_Bezeichnungsfeld
' vLab(58)       :BHD seit_Bezeichnungsfeld
' vLab(59)       :BHD beh mit_Bezeichnungsfeld
' vLab(60)       :Blutdruckwerte_Bezeichnungsfeld
' vLab(61)       :BDselbst_Bezeichnungsfeld
' vLab(62)       :Schwanger_Bezeichnungsfeld
' vLab(63)       :Schwanger seit_Bezeichnungsfeld
' vLab(64)       :Augensp zuletzt_Bezeichnungsfeld
' vLab(65)       :Augensp Befund_Bezeichnungsfeld
' vLab(66)       :Netzhaut gelasert_Bezeichnungsfeld
' vLab(67)       :Sehminderung unbehebbar_Bezeichnungsfeld
' vLab(68)       :Diabet Nierenschaden_Bezeichnungsfeld
' vLab(69)       :Albumin zuletzt_Bezeichnungsfeld
' vLab(70)       :erhöht?_Bezeichnungsfeld
' vLab(71)       :Dialyse_Bezeichnungsfeld
' vLab(72)       :Dialyse seit_Bezeichnungsfeld
' vLab(73)       :andere Nierenerkrankung_Bezeichnungsfeld
' vLab(74)       :Herzkrankheit_Bezeichnungsfeld
' vLab(75)       :Angina pectoris_Bezeichnungsfeld
' vLab(76)       :Herzinfakt_Bezeichnungsfeld
' vLab(77)       :Herzinfarkt wann_Bezeichnungsfeld
' vLab(78)       :PTCA oder Stent_Bezeichnungsfeld
' vLab(79)       :Bypass kardial_Bezeichnungsfeld
' vLab(80)       :Bypass wann_Bezeichnungsfeld
' vLab(81)       :Herzschwäche_Bezeichnungsfeld
' vLab(82)       :Herzkrankheit Beschreibung_Bezeichnungsfeld
' vLab(83)       :Hirndurchblutungsstörung_Bezeichnungsfeld
' vLab(84)       :Schlaganfall_Bezeichnungsfeld
' vLab(85)       :Beindurchblutungsstörung_Bezeichnungsfeld
' vLab(86)       :Schaufensterkrankheit_Bezeichnungsfeld
' vLab(87)       :Bypaß peripher_Bezeichnungsfeld
' vLab(88)       :Geschwür_Bezeichnungsfeld
' vLab(89)       :Amputation_Bezeichnungsfeld
' vLab(90)       :pAVK Beschreibung_Bezeichnungsfeld
' vLab(91)       :Ameisenlaufen_Bezeichnungsfeld
' vLab(92)       :Ameisen Ausmaß_Bezeichnungsfeld
' vLab(93)       :Druckstellen_Bezeichnungsfeld
' vLab(94)       :Verformungen_Bezeichnungsfeld
' vLab(95)       :Verformungen Beschreibung_Bezeichnungsfeld
' vLab(96)       :Neue Fußkomplikationen_Bezeichnungsfeld
' vLab(97)       :Entleerungsstörungen Magen_Bezeichnungsfeld
' vLab(98)       :Entleerungsstörungen Harnblase_Bezeichnungsfeld
' vLab(99)       :Schwindel Aufstehen_Bezeichnungsfeld
' vLab(100)      :Folgeerkrankungen Haut_Bezeichnungsfeld
' vLab(101)      :Bewegungseinschränkungen_Bezeichnungsfeld
' vLab(102)      :Sexualstörung_Bezeichnungsfeld
' vLab(103)      :Sexualstörung seit_Bezeichnungsfeld
' vLab(104)      :Weitere Anamnese_Bezeichnungsfeld
' vLab(105)      :Alkohol_Bezeichnungsfeld
' vLab(106)      :Tabak_Bezeichnungsfeld
' vLab(107)      :Weitere Medikation_Bezeichnungsfeld
' vLab(108)      :Liphypertrophien Abdomen_Bezeichnungsfeld
' vLab(109)      :Liphypertrophien Beine_Bezeichnungsfeld
' vLab(110)      :Liphypertrophien Arme_Bezeichnungsfeld
' vLab(111)      :Hyperkeratosen_Bezeichnungsfeld
' vLab(112)      :Ulcera_Bezeichnungsfeld
' vLab(113)      :Kraft Zehenheber_Bezeichnungsfeld
' vLab(114)      :Kraft Zehenbeuger_Bezeichnungsfeld
' vLab(115)      :Kraft Knie_Bezeichnungsfeld
' vLab(116)      :ASR_Bezeichnungsfeld
' vLab(117)      :PSR_Bezeichnungsfeld
' vLab(118)      :Oberflächensensibilität_Bezeichnungsfeld
' vLab(119)      :Monofilamenttest_Bezeichnungsfeld
' vLab(120)      :Kalt-Warm_Bezeichnungsfeld
' vLab(121)      :Vibration IK_Bezeichnungsfeld
' vLab(122)      :Vibration Großzehe_Bezeichnungsfeld
' vLab(123)      :Puls Leiste_Bezeichnungsfeld
' vLab(124)      :Puls Kniekehle_Bezeichnungsfeld
' vLab(125)      :Puls Atp_Bezeichnungsfeld
' vLab(126)      :Puls Adp_Bezeichnungsfeld
' vLab(127)      :RR_Bezeichnungsfeld
' vLab(128)      :Herz_Bezeichnungsfeld
' vLab(129)      :Lunge_Bezeichnungsfeld
' vLab(130)      :Bauch_Bezeichnungsfeld
' vLab(131)      :WS_Bezeichnungsfeld
' vLab(132)      :NL_Bezeichnungsfeld
' vLab(133)      :SD_Bezeichnungsfeld
' vLab(134)      :NNH_Bezeichnungsfeld
' vLab(135)      :Zähne_Bezeichnungsfeld
' vLab(136)      :Mundhöhle_Bezeichnungsfeld
' vLab(137)      :LK_Bezeichnungsfeld
' vLab(138)      :Neuro sonst_Bezeichnungsfeld
' vLab(139)      :Weitere Befunde_Bezeichnungsfeld
' vLab(140)      :Schulung_Bezeichnungsfeld
' vLab(141)      :DMP_Bezeichnungsfeld
' vLab(142)      :HANr_Bezeichnungsfeld
' vLab(143)      :Beinbefund_Bezeichnungsfeld
' vLab(144)      :BMI-Bezeichnungsfeld
' vLab(145)      :Bezeichnungsfeld_BMI_Einheit
' vLab(146)      :letzte Änderung_Bezeichnung
' vLab(147)      :Fußpflege_Bezeichnungsfeld
' vLab(148)      :Podologie_Bezeichnungsfeld
' vLab(149)      :Einlagen_Bezeichnungsfeld
' vLab(150)      :Dioagnosen_Bezeichnungsfeld
' vLab(151)      :Vorgestellt_Bezeichnungsfeld
' vLab(152)      :RRTurboMed_Bezeichnungsfeld
' vLab(153)      :Versicherung_Bezeichnungsfeld
' vLab(154)      :Essenszeit vormittags_Bezeichnungsfeld
' vLab(155)      :Essenszeit nachmittags_Beichnungsfeld
' vLab(156)      :BZMessungen selbst_Bezeichnungsfeld
' vLab(157)      :Gerät_Bezeichnungsfeld
' vLab(158)      :UZTagesZeit_Bezeichnungsfeld
' vLab(159)      :BeiödVen_Bezeichnungsfeld
' vLab(160)      :Carotiden_Bezeichnungsfeld
' vLab(161)      :DMSchulz_Bezeichnungsfeld
' vLab(162)      :RR-Schulz-Bezeichnungsfeld
' vLab(163)      :DMPhier_Bezeichnungsfeld
' vLab(164)      :Tkz_Bezeichnungsfeld
' vLab(165)      :Therapieart_Bezeichnungsfeld
' vLab(166)      :obBZausgew_Bez
' vLab(167)      :obOSaufgek_Bez
' vLab(168)      :obPodAufgek_Bez
' vLab(169)      :obMBlAusgeh_Bez
' vLab(170)      :obDMPaufgekl_Bez
' vLab(171)      :Schulungsaufkl_Bez
' vLab(172)      :TherAkt_Bezeichnung
' vLab(173)      :SchGr_Bezeichnungsfeld
' vLab(174)      :obMedNetz_Bezeichnung
' vLab(175)      :HA1Nicht_Bez
' vLab(176)      :HA2Nicht_Bez
' vLab(177)      :HAFax_Bezeichnungsfeld
' vLab(178)      :DopplerBez
' vLab(179)      :Duplex_Bezeichnungsfeld
' vLab(180)      :SonoBez
' vLab(181)      :lebehBez
' vLab(182)      :Datenquelle_Bezeichnungsfeld
' vTextB(1)      :Pat_id
' vTextB(2)      :Vorname
' vTextB(3)      :GebDat
' vTextB(4)      :Diabetestyp
' vTextB(5)      :Diabetes seit
' vTextB(6)      :Tabletten seit
' vTextB(7)      :Insulin seit
' vTextB(8)      :Grund für Vorstellung
' vTextB(9)      :Familienanamnese
' vTextB(10)     :Größe
' vTextB(11)     :Gewicht
' vTextB(12)     :Tendenz
' vTextB(13)     :DiabetesMedikament 1
' vTextB(14)     :DiabetesMedikamente 1 Menge
' vTextB(15)     :DiabetesMedikament 2
' vTextB(16)     :DiabetesMedikament 2 Menge
' vTextB(17)     :DiabetesMedikament 3
' vTextB(18)     :DiabetesMedikament 3 Menge
' vTextB(19)     :DiabetesMedikament 4
' vTextB(20)     :DiabetesMedikament 4 Menge
' vTextB(21)     :Insulinpumpe seit
' vTextB(22)     :Insulinpumpe Marke
' vTextB(23)     :Broteinheiten gesamt
' vTextB(24)     :Broteinheiten früh
' vTextB(25)     :Broteinheiten ZM früh
' vTextB(26)     :Broteinheiten mittags
' vTextB(27)     :Broteinheiten nachmittags
' vTextB(28)     :Broteinheiten abends
' vTextB(29)     :Broteinheiten nachts
' vTextB(30)     :Essenszeit  früh
' vTextB(31)     :Essenszeit mittags
' vTextB(32)     :Essenszeit abends
' vTextB(33)     :Essenszeit spät
' vTextB(34)     :Spritz-Eß-Abstand früh
' vTextB(35)     :Spritz-Eß-Abstand mittags
' vTextB(36)     :Spritz-Eß-Abstand abends
' vTextB(37)     :Spritzstelle früh (B, O, A)
' vTextB(38)     :Spritzstelle mittags
' vTextB(39)     :Spritzstelle abends
' vTextB(40)     :Spritzstelle nachts
' vTextB(41)     :Jahr letzte Diabetesschulung
' vTextB(42)     :Ort Schulung
' vTextB(43)     :letztes HbA1c
' vTextB(44)     :gemessen am
' vTextB(45)     :vorherige Werte
' vTextB(46)     :BZMessungen pW
' vTextB(47)     :BZMessungen pW ndE
' vTextB(48)     :BZMessungen p W nachts
' vTextB(49)     :Aufschreiben
' vTextB(50)     :BZWerte v d Essen
' vTextB(51)     :BZWerte n d Essen
' vTextB(52)     :Unterzucker pM
' vTextB(53)     :UZ rechtzeitig
' vTextB(54)     :Fremde Hilfe pa
' vTextB(55)     :Bewußtlos pa
' vTextB(56)     :Keto pa
' vTextB(57)     :BZgr300 pM
' vTextB(58)     :Bluthochdruck
' vTextB(59)     :BHD seit
' vTextB(60)     :BHD beh mit
' vTextB(61)     :Blutdruckwerte
' vTextB(62)     :BDselbst
' vTextB(63)     :Schwanger
' vTextB(64)     :Schwanger seit
' vTextB(65)     :Augensp zuletzt
' vTextB(66)     :Augensp Befund
' vTextB(67)     :Netzhaut gelasert
' vTextB(68)     :Sehminderung unbehebbar
' vTextB(69)     :Diabet Nierenschaden
' vTextB(70)     :Albumin zuletzt
' vTextB(71)     :erhöht?
' vTextB(72)     :Dialyse seit
' vTextB(73)     :andere Nierenerkrankung
' vTextB(74)     :Herzkrankheit
' vTextB(75)     :Angina pectoris
' vTextB(76)     :Herzinfarkt
' vTextB(77)     :Herzinfarkt wann
' vTextB(78)     :PTCA oder Stent
' vTextB(79)     :Bypass wann
' vTextB(80)     :Herzschwäche
' vTextB(81)     :Herzkrankheit Beschreibung
' vTextB(82)     :Hirndurchblutungsstörung
' vTextB(83)     :Schlaganfall
' vTextB(84)     :Beindurchblutungsstörung
' vTextB(85)     :Schaufensterkrankheit
' vTextB(86)     :Geschwür
' vTextB(87)     :Amputation
' vTextB(88)     :pAVK Beschreibung
' vTextB(89)     :Ameisenlaufen
' vTextB(90)     :Ameisen Ausmaß
' vTextB(91)     :Druckstellen
' vTextB(92)     :Verformungen
' vTextB(93)     :Verformungen Beschreibung
' vTextB(94)     :Neue Fußkomplikationen
' vTextB(95)     :Entleerungsstörungen Magen
' vTextB(96)     :Entleerungsstörungen Harnblase
' vTextB(97)     :Schwindel Aufstehen
' vTextB(98)     :Folgeerkrankungen Haut
' vTextB(99)     :Bewegungseinschränkungen
' vTextB(100)    :Sexualstörung
' vTextB(101)    :Sexualstörung seit
' vTextB(102)    :Weitere Anamnese
' vTextB(103)    :Alkohol
' vTextB(104)    :Tabak
' vTextB(105)    :Weitere Medikation
' vTextB(106)    :Liphypertrophien Abdomen
' vTextB(107)    :Liphypertrophien Beine
' vTextB(108)    :Liphypertrophien Arme
' vTextB(109)    :Hyperkeratosen
' vTextB(110)    :Ulcera
' vTextB(111)    :Kraft Zehenheber
' vTextB(112)    :Kraft Zehenbeuger
' vTextB(113)    :Kraft Knie
' vTextB(114)    :ASR
' vTextB(115)    :PSR
' vTextB(116)    :Oberflächensensibilität
' vTextB(117)    :Monofilamenttest
' vTextB(118)    :Kalt-Warm
' vTextB(119)    :Vibration IK
' vTextB(120)    :Vibration Großzehe
' vTextB(121)    :Puls Leiste
' vTextB(122)    :Puls Kniekehle
' vTextB(123)    :Puls Atp
' vTextB(124)    :Puls Adp
' vTextB(125)    :RR
' vTextB(126)    :Herz
' vTextB(127)    :Lunge
' vTextB(128)    :Bauch
' vTextB(129)    :WS
' vTextB(130)    :NL
' vTextB(131)    :SD
' vTextB(132)    :NNH
' vTextB(133)    :Zähne
' vTextB(134)    :Mundhöhle
' vTextB(135)    :LK
' vTextB(136)    :Neuro sonst
' vTextB(137)    :Weitere Befunde
' vTextB(138)    :Schulung
' vTextB(139)    :DMP
' vTextB(140)    :HANr
' vTextB(141)    :Beinbefund
' vTextB(142)    :BMI
' vTextB(143)    :letzte Änderung
' vTextB(144)    :Fußpflege
' vTextB(145)    :Podologie
' vTextB(146)    :Einlagen
' vTextB(147)    :Diagnosen
' vTextB(148)    :HA
' vTextB(149)    :Vorgestellt
' vTextB(150)    :RRTurbomed
' vTextB(151)    :Versicherung
' vTextB(152)    :Titel
' vTextB(153)    :Anrede
' vTextB(154)    :Versicherungsart
' vTextB(155)    :Essenszeit vormittags
' vTextB(156)    :Essenszeit nachmittags
' vTextB(157)    :BZMessungen Selbst
' vTextB(158)    :Gerät
' vTextB(159)    :UZ Tageszeit
' vTextB(160)    :NachName
' vTextB(161)    :BeinödVen
' vTextB(162)    :Carotiden
' vTextB(163)    :DMSchulz
' vTextB(164)    :RRSchulz
' vTextB(165)    :DMPhier
' vTextB(166)    :Ther1
' vTextB(167)    :Gesamtname
' vTextB(168)    :MDi
' vTextB(169)    :HAimDMP
' vTextB(170)    :V1
' vTextB(171)    :obDMPaufgekl
' vTextB(172)    :obSchulaufgek
' vTextB(173)    :TherAkt
' vTextB(174)    :SchGr
' vTextB(175)    :HANr2
' vTextB(176)    :HAFax
' vTextB(177)    :Doppler
' vTextB(178)    :Duplex
' vTextB(179)    :Sono
' vTextB(180)    :Eintraege
' vTextB(181)    :leBeh
' vTextB(182)    :Datenbankname
' vCheckb(1)     :Insulinpumpe
' vCheckb(2)     :Dialyse
' vCheckb(3)     :Bypass kardial
' vCheckb(4)     :Bypaß peripher
' vCheckb(5)     :Tkz
' vCheckb(6)     :obAn1eing
' vCheckb(7)     :obAn2eing
' vCheckb(8)     :obAnAeing
' vCheckb(9)     :obCheck
' vCheckb(10)    :obBZausgew
' vCheckb(11)    :obOSaufgek
' vCheckb(12)    :obPodAufgek
' vCheckb(13)    :obMBlAusgeh
' vCheckb(14)    :obMedNetz
' vCheckb(15)    :HA1nicht
' vCheckb(16)    :HA2nicht
' vCommandB(1)   :Ausgabe
' vCommandB(2)   :An1Aufruf
' vCommandB(3)   :An2Aufruf
' vCommandB(4)   :AnAAufruf
' vCommandB(5)   :CheckAufruf
' vCommandB(6)   :Labor
' vCommandB(7)   :Brief
' vCommandB(8)   :Dokumente
' vCommandB(9)   :Briefe
' vCommandB(10)  :DReset
' vCommandB(11)  :BZKurven
' vCommandB(12)  :DMPAusgeb
' vCommandB(13)  :AugenBefunde
' vCommandB(14)  :DokDown
' vCommandB(15)  :HA1_Bez
' vCommandB(16)  :HA2_Bez
' vCommandB(17)  :Unausgefülle
' vCommandB(18)  :alleDatensätze
' vCommandB(19)  :UnausgefNZ
' vCommandB(20)  :Datenbank-Aufruf
' vComboB(1)     :HA_Auswahl
' vComboB(2)     :HA_Auswahl2
' vComboB(3)     :Datenquelle
' vOptionB(1)    :U1
' vOptionB(2)    :U2
' vOptionB(3)    :U3
' vOptionB(4)    :U4
' vOptionB(5)    :U5
' vOptionB(6)    :U6
' vOptionB(7)    :U7
' vOptionB(8)    :U8
' vOptionB(9)    :U9
' vOptionB(10)   :U10
' vOptionB(11)   :U11
' vOptionB(12)   :U12
' vOptionB(13)   :U13
' vOptionB(14)   :U14
' vOptionB(15)   :U15
' vOptionB(16)   :U16
' vOptionB(17)   :U17
' vOptionB(18)   :U18
' vOptionB(19)   :U19
' vOptionB(20)   :U20
' vOptionB(21)   :U21
' vOptionB(22)   :U22
' vOptionB(23)   :U23

'ASR:                                    vTextB(114)
'ASR_Bezeichnungsfeld:                   vLab(116)
'Albumin zuletzt:                        vTextB(70)
'Albumin zuletzt_Bezeichnungsfeld:       vLab(69)
'Alkohol:                                vTextB(103)
'Alkohol_Bezeichnungsfeld:               vLab(105)
'Ameisen Ausmaß:                         vTextB(90)
'Ameisen Ausmaß_Bezeichnungsfeld:        vLab(92)
'Ameisenlaufen:                          vTextB(89)
'Ameisenlaufen_Bezeichnungsfeld:         vLab(91)
'Amputation:                             vTextB(87)
'Amputation_Bezeichnungsfeld:            vLab(89)
'An1Aufruf:                              vCommandB(2)
'An2Aufruf:                              vCommandB(3)
'AnAAufruf:                              vCommandB(4)
'Angina pectoris:                        vTextB(75)
'Angina pectoris_Bezeichnungsfeld:       vLab(75)
'Anrede:                                 vTextB(153)
'Aufschreiben:                           vTextB(49)
'Aufschreiben_Bezeichnungsfeld:          vLab(48)
'AugenBefunde:                           vCommandB(13)
'Augensp Befund:                         vTextB(66)
'Augensp Befund_Bezeichnungsfeld:        vLab(65)
'Augensp zuletzt:                        vTextB(65)
'Augensp zuletzt_Bezeichnungsfeld:       vLab(64)
'Ausgabe:                                vCommandB(1)
'BDselbst:                               vTextB(62)
'BDselbst_Bezeichnungsfeld:              vLab(61)
'BHD beh mit:                            vTextB(60)
'BHD beh mit_Bezeichnungsfeld:           vLab(59)
'BHD seit:                               vTextB(59)
'BHD seit_Bezeichnungsfeld:              vLab(58)
'BMI-Bezeichnungsfeld:                   vLab(144)
'BMI:                                    vTextB(142)
'BZKurven:                               vCommandB(11)
'BZMessungen Selbst:                     vTextB(157)
'BZMessungen p W nachts:                 vTextB(48)
'BZMessungen p W nachts_Bezeichnungsfeld:vLab(47)
'BZMessungen pW ndE:                     vTextB(47)
'BZMessungen pW ndE_Bezeichnungsfeld:    vLab(46)
'BZMessungen pW:                         vTextB(46)
'BZMessungen pW_Bezeichnungsfeld:        vLab(45)
'BZMessungen selbst_Bezeichnungsfeld:    vLab(156)
'BZWerte n d Essen:                      vTextB(51)
'BZWerte n d Essen_Bezeichnungsfeld:     vLab(50)
'BZWerte v d Essen:                      vTextB(50)
'BZWerte v d Essen_Bezeichnungsfeld:     vLab(49)
'BZgr300 pM:                             vTextB(57)
'BZgr300 pM_Bezeichnungsfeld:            vLab(56)
'Bauch:                                  vTextB(128)
'Bauch_Bezeichnungsfeld:                 vLab(130)
'Beinbefund:                             vTextB(141)
'Beinbefund_Bezeichnungsfeld:            vLab(143)
'Beindurchblutungsstörung:               vTextB(84)
'Beindurchblutungsstörung_BezeichnungsfelvLab(85)
'BeinödVen:                              vTextB(161)
'BeiödVen_Bezeichnungsfeld:              vLab(159)
'Bewegungseinschränkungen:               vTextB(99)
'Bewegungseinschränkungen_BezeichnungsfelvLab(101)
'Bewußtlos pa:                           vTextB(55)
'Bewußtlos pa_Bezeichnungsfeld:          vLab(54)
'Bezeichnungsfeld_BMI_Einheit:           vLab(145)
'Blutdruckwerte:                         vTextB(61)
'Blutdruckwerte_Bezeichnungsfeld:        vLab(60)
'Bluthochdruck:                          vTextB(58)
'Bluthochdruck_Bezeichnungsfeld:         vLab(57)
'Brief:                                  vCommandB(7)
'Briefe:                                 vCommandB(9)
'Broteinheiten ZM früh:                  vTextB(25)
'Broteinheiten ZM früh_Bezeichnungsfeld: vLab(24)
'Broteinheiten abends:                   vTextB(28)
'Broteinheiten abends_Bezeichnungsfeld:  vLab(27)
'Broteinheiten früh:                     vTextB(24)
'Broteinheiten früh_Bezeichnungsfeld:    vLab(23)
'Broteinheiten gesamt:                   vTextB(23)
'Broteinheiten gesamt_Bezeichnungsfeld:  vLab(22)
'Broteinheiten mittags:                  vTextB(26)
'Broteinheiten mittags_Bezeichnungsfeld: vLab(25)
'Broteinheiten nachmittags:              vTextB(27)
'Broteinheiten nachmittags_BezeichnungsfevLab(26)
'Broteinheiten nachts:                   vTextB(29)
'Broteinheiten nachts_Bezeichnungsfeld:  vLab(28)
'Bypass kardial:                         vCheckb(3)
'Bypass kardial_Bezeichnungsfeld:        vLab(79)
'Bypass wann:                            vTextB(79)
'Bypass wann_Bezeichnungsfeld:           vLab(80)
'Bypaß peripher:                         vCheckb(4)
'Bypaß peripher_Bezeichnungsfeld:        vLab(87)
'Carotiden:                              vTextB(162)
'Carotiden_Bezeichnungsfeld:             vLab(160)
'CheckAufruf:                            vCommandB(5)
'DMP:                                    vTextB(139)
'DMPAusgeb:                              vCommandB(12)
'DMP_Bezeichnungsfeld:                   vLab(141)
'DMPhier:                                vTextB(165)
'DMPhier_Bezeichnungsfeld:               vLab(163)
'DMSchulz:                               vTextB(163)
'DMSchulz_Bezeichnungsfeld:              vLab(161)
'DReset:                                 vCommandB(10)
'Datenbank-Aufruf:                       vCommandB(20)
'Datenbankname:                          vTextB(182)
'Datenquelle:                            vComboB(3)
'Datenquelle_Bezeichnungsfeld:           vLab(182)
'Diabet Nierenschaden:                   vTextB(69)
'Diabet Nierenschaden_Bezeichnungsfeld:  vLab(68)
'Diabetes seit:                          vTextB(5)
'Diabetes seit_Bezeichnungsfeld:         vLab(3)
'DiabetesMedikament 1:                   vTextB(13)
'DiabetesMedikament 1_Bezeichnungsfeld:  vLab(11)
'DiabetesMedikament 2 Menge:             vTextB(16)
'DiabetesMedikament 2 Menge_BezeichnungsfvLab(14)
'DiabetesMedikament 2:                   vTextB(15)
'DiabetesMedikament 2_Bezeichnungsfeld:  vLab(13)
'DiabetesMedikament 3 Menge:             vTextB(18)
'DiabetesMedikament 3 Menge_BezeichnungsfvLab(16)
'DiabetesMedikament 3:                   vTextB(17)
'DiabetesMedikament 3_Bezeichnungsfeld:  vLab(15)
'DiabetesMedikament 4 Menge:             vTextB(20)
'DiabetesMedikament 4 Menge_BezeichnungsfvLab(18)
'DiabetesMedikament 4:                   vTextB(19)
'DiabetesMedikament 4_Bezeichnungsfeld:  vLab(17)
'DiabetesMedikamente 1 Menge:            vTextB(14)
'DiabetesMedikamente 1 Menge_BezeichnungsvLab(12)
'Diabetestyp:                            vTextB(4)
'Diabetestyp_Bezeichnungsfeld:           vLab(2)
'Diagnosen:                              vTextB(147)
'Dialyse seit:                           vTextB(72)
'Dialyse seit_Bezeichnungsfeld:          vLab(72)
'Dialyse:                                vCheckb(2)
'Dialyse_Bezeichnungsfeld:               vLab(71)
'Dioagnosen_Bezeichnungsfeld:            vLab(150)
'DokDown:                                vCommandB(14)
'Dokumente:                              vCommandB(8)
'Doppler:                                vTextB(177)
'DopplerBez:                             vLab(178)
'Druckstellen:                           vTextB(91)
'Druckstellen_Bezeichnungsfeld:          vLab(93)
'Duplex:                                 vTextB(178)
'Duplex_Bezeichnungsfeld:                vLab(179)
'Einlagen:                               vTextB(146)
'Einlagen_Bezeichnungsfeld:              vLab(149)
'Eintraege:                              vTextB(180)
'Entleerungsstörungen Harnblase:         vTextB(96)
'Entleerungsstörungen Harnblase_BezeichnuvLab(98)
'Entleerungsstörungen Magen:             vTextB(95)
'Entleerungsstörungen Magen_BezeichnungsfvLab(97)
'Essenszeit  früh:                       vTextB(30)
'Essenszeit  früh_Bezeichnungsfeld:      vLab(29)
'Essenszeit abends:                      vTextB(32)
'Essenszeit abends_Bezeichnungsfeld:     vLab(31)
'Essenszeit mittags:                     vTextB(31)
'Essenszeit mittags_Bezeichnungsfeld:    vLab(30)
'Essenszeit nachmittags:                 vTextB(156)
'Essenszeit nachmittags_Beichnungsfeld:  vLab(155)
'Essenszeit spät:                        vTextB(33)
'Essenszeit spät_Bezeichnungsfeld:       vLab(32)
'Essenszeit vormittags:                  vTextB(155)
'Essenszeit vormittags_Bezeichnungsfeld: vLab(154)
'Familienanamnese:                       vTextB(9)
'Familienanamnese_Bezeichnungsfeld:      vLab(7)
'Folgeerkrankungen Haut:                 vTextB(98)
'Folgeerkrankungen Haut_Bezeichnungsfeld:vLab(100)
'Fremde Hilfe pa:                        vTextB(54)
'Fremde Hilfe pa_Bezeichnungsfeld:       vLab(53)
'Fußpflege:                              vTextB(144)
'Fußpflege_Bezeichnungsfeld:             vLab(147)
'GebDat:                                 vTextB(3)
'GebDat_Bezeichnungsfeld:                vLab(1)
'Gerät:                                  vTextB(158)
'Gerät_Bezeichnungsfeld:                 vLab(157)
'Gesamtname:                             vTextB(167)
'Geschwür:                               vTextB(86)
'Geschwür_Bezeichnungsfeld:              vLab(88)
'Gewicht:                                vTextB(11)
'Gewicht_Bezeichnungsfeld:               vLab(9)
'Grund für Vorstellung:                  vTextB(8)
'Grund für Vorstellung_Bezeichnungsfeld: vLab(6)
'Größe:                                  vTextB(10)
'Größe_Bezeichnungsfeld:                 vLab(8)
'HA1Nicht_Bez:                           vLab(175)
'HA1_Bez:                                vCommandB(15)
'HA1nicht:                               vCheckb(15)
'HA2Nicht_Bez:                           vLab(176)
'HA2_Bez:                                vCommandB(16)
'HA2nicht:                               vCheckb(16)
'HA:                                     vTextB(148)
'HAFax:                                  vTextB(176)
'HAFax_Bezeichnungsfeld:                 vLab(177)
'HANr2:                                  vTextB(175)
'HANr:                                   vTextB(140)
'HANr_Bezeichnungsfeld:                  vLab(142)
'HA_Auswahl2:                            vComboB(2)
'HA_Auswahl:                             vComboB(1)
'HAimDMP:                                vTextB(169)
'Herz:                                   vTextB(126)
'Herz_Bezeichnungsfeld:                  vLab(128)
'Herzinfakt_Bezeichnungsfeld:            vLab(76)
'Herzinfarkt wann:                       vTextB(77)
'Herzinfarkt wann_Bezeichnungsfeld:      vLab(77)
'Herzinfarkt:                            vTextB(76)
'Herzkrankheit Beschreibung:             vTextB(81)
'Herzkrankheit Beschreibung_BezeichnungsfvLab(82)
'Herzkrankheit:                          vTextB(74)
'Herzkrankheit_Bezeichnungsfeld:         vLab(74)
'Herzschwäche:                           vTextB(80)
'Herzschwäche_Bezeichnungsfeld:          vLab(81)
'Hirndurchblutungsstörung:               vTextB(82)
'Hirndurchblutungsstörung_BezeichnungsfelvLab(83)
'Hyperkeratosen:                         vTextB(109)
'Hyperkeratosen_Bezeichnungsfeld:        vLab(111)
'Insulin seit:                           vTextB(7)
'Insulin seit_Bezeichnungsfeld:          vLab(5)
'Insulinpumpe Marke:                     vTextB(22)
'Insulinpumpe Marke_Bezeichnungsfeld:    vLab(21)
'Insulinpumpe seit:                      vTextB(21)
'Insulinpumpe seit_Bezeichnungsfeld:     vLab(20)
'Insulinpumpe:                           vCheckb(1)
'Insulinpumpe_Bezeichnungsfeld:          vLab(19)
'Jahr letzte Diabetesschulung:           vTextB(41)
'Jahr letzte Diabetesschulung_BezeichnungvLab(40)
'Kalt-Warm:                              vTextB(118)
'Kalt-Warm_Bezeichnungsfeld:             vLab(120)
'Keto pa:                                vTextB(56)
'Keto pa_Bezeichnungsfeld:               vLab(55)
'Kraft Knie:                             vTextB(113)
'Kraft Knie_Bezeichnungsfeld:            vLab(115)
'Kraft Zehenbeuger:                      vTextB(112)
'Kraft Zehenbeuger_Bezeichnungsfeld:     vLab(114)
'Kraft Zehenheber:                       vTextB(111)
'Kraft Zehenheber_Bezeichnungsfeld:      vLab(113)
'LK:                                     vTextB(135)
'LK_Bezeichnungsfeld:                    vLab(137)
'Labor:                                  vCommandB(6)
'Liphypertrophien Abdomen:               vTextB(106)
'Liphypertrophien Abdomen_BezeichnungsfelvLab(108)
'Liphypertrophien Arme:                  vTextB(108)
'Liphypertrophien Arme_Bezeichnungsfeld: vLab(110)
'Liphypertrophien Beine:                 vTextB(107)
'Liphypertrophien Beine_Bezeichnungsfeld:vLab(109)
'Lunge:                                  vTextB(127)
'Lunge_Bezeichnungsfeld:                 vLab(129)
'MDi:                                    vTextB(168)
'Monofilamenttest:                       vTextB(117)
'Monofilamenttest_Bezeichnungsfeld:      vLab(119)
'Mundhöhle:                              vTextB(134)
'Mundhöhle_Bezeichnungsfeld:             vLab(136)
'NL:                                     vTextB(130)
'NL_Bezeichnungsfeld:                    vLab(132)
'NNH:                                    vTextB(132)
'NNH_Bezeichnungsfeld:                   vLab(134)
'NachName:                               vTextB(160)
'Netzhaut gelasert:                      vTextB(67)
'Netzhaut gelasert_Bezeichnungsfeld:     vLab(66)
'Neue Fußkomplikationen:                 vTextB(94)
'Neue Fußkomplikationen_Bezeichnungsfeld:vLab(96)
'Neuro sonst:                            vTextB(136)
'Neuro sonst_Bezeichnungsfeld:           vLab(138)
'Oberflächensensibilität:                vTextB(116)
'Oberflächensensibilität_BezeichnungsfeldvLab(118)
'Ort Schulung:                           vTextB(42)
'Ort Schulung_Bezeichnungsfeld:          vLab(41)
'PSR:                                    vTextB(115)
'PSR_Bezeichnungsfeld:                   vLab(117)
'PTCA oder Stent:                        vTextB(78)
'PTCA oder Stent_Bezeichnungsfeld:       vLab(78)
'Pat_id:                                 vTextB(1)
'Podologie:                              vTextB(145)
'Podologie_Bezeichnungsfeld:             vLab(148)
'Puls Adp:                               vTextB(124)
'Puls Adp_Bezeichnungsfeld:              vLab(126)
'Puls Atp:                               vTextB(123)
'Puls Atp_Bezeichnungsfeld:              vLab(125)
'Puls Kniekehle:                         vTextB(122)
'Puls Kniekehle_Bezeichnungsfeld:        vLab(124)
'Puls Leiste:                            vTextB(121)
'Puls Leiste_Bezeichnungsfeld:           vLab(123)
'RR-Schulz-Bezeichnungsfeld:             vLab(162)
'RR:                                     vTextB(125)
'RRSchulz:                               vTextB(164)
'RRTurboMed_Bezeichnungsfeld:            vLab(152)
'RRTurbomed:                             vTextB(150)
'RR_Bezeichnungsfeld:                    vLab(127)
'SD:                                     vTextB(131)
'SD_Bezeichnungsfeld:                    vLab(133)
'SchGr:                                  vTextB(174)
'SchGr_Bezeichnungsfeld:                 vLab(173)
'Schaufensterkrankheit:                  vTextB(85)
'Schaufensterkrankheit_Bezeichnungsfeld: vLab(86)
'Schlaganfall:                           vTextB(83)
'Schlaganfall_Bezeichnungsfeld:          vLab(84)
'Schulung:                               vTextB(138)
'Schulung_Bezeichnungsfeld:              vLab(140)
'Schulungsaufkl_Bez:                     vLab(171)
'Schwanger seit:                         vTextB(64)
'Schwanger seit_Bezeichnungsfeld:        vLab(63)
'Schwanger:                              vTextB(63)
'Schwanger_Bezeichnungsfeld:             vLab(62)
'Schwindel Aufstehen:                    vTextB(97)
'Schwindel Aufstehen_Bezeichnungsfeld:   vLab(99)
'Sehminderung unbehebbar:                vTextB(68)
'Sehminderung unbehebbar_BezeichnungsfeldvLab(67)
'Sexualstörung seit:                     vTextB(101)
'Sexualstörung seit_Bezeichnungsfeld:    vLab(103)
'Sexualstörung:                          vTextB(100)
'Sexualstörung_Bezeichnungsfeld:         vLab(102)
'Sono:                                   vTextB(179)
'SonoBez:                                vLab(180)
'Spritz-Eß-Abstand abends:               vTextB(36)
'Spritz-Eß-Abstand abends_BezeichnungsfelvLab(35)
'Spritz-Eß-Abstand früh:                 vTextB(34)
'Spritz-Eß-Abstand früh_Bezeichnungsfeld:vLab(33)
'Spritz-Eß-Abstand mittags:              vTextB(35)
'Spritz-Eß-Abstand mittags_BezeichnungsfevLab(34)
'Spritzstelle abends:                    vTextB(39)
'Spritzstelle abends_Bezeichnungsfeld:   vLab(38)
'Spritzstelle früh (B, O, A):            vTextB(37)
'Spritzstelle früh (B, O, A)_BezeichnungsvLab(36)
'Spritzstelle mittags:                   vTextB(38)
'Spritzstelle mittags_Bezeichnungsfeld:  vLab(37)
'Spritzstelle nachts:                    vTextB(40)
'Spritzstelle nachts_Bezeichnungsfeld:   vLab(39)
'Tabak:                                  vTextB(104)
'Tabak_Bezeichnungsfeld:                 vLab(106)
'Tabletten seit:                         vTextB(6)
'Tabletten seit_Bezeichnungsfeld:        vLab(4)
'Tendenz:                                vTextB(12)
'Tendenz_Bezeichnungsfeld:               vLab(10)
'Ther1:                                  vTextB(166)
'TherAkt:                                vTextB(173)
'TherAkt_Bezeichnung:                    vLab(172)
'Therapieart_Bezeichnungsfeld:           vLab(165)
'Titel:                                  vTextB(152)
'Tkz:                                    vCheckb(5)
'Tkz_Bezeichnungsfeld:                   vLab(164)
'U10:                                    vOptionB(10)
'U11:                                    vOptionB(11)
'U12:                                    vOptionB(12)
'U13:                                    vOptionB(13)
'U14:                                    vOptionB(14)
'U15:                                    vOptionB(15)
'U16:                                    vOptionB(16)
'U17:                                    vOptionB(17)
'U18:                                    vOptionB(18)
'U19:                                    vOptionB(19)
'U1:                                     vOptionB(1)
'U20:                                    vOptionB(20)
'U21:                                    vOptionB(21)
'U22:                                    vOptionB(22)
'U23:                                    vOptionB(23)
'U2:                                     vOptionB(2)
'U3:                                     vOptionB(3)
'U4:                                     vOptionB(4)
'U5:                                     vOptionB(5)
'U6:                                     vOptionB(6)
'U7:                                     vOptionB(7)
'U8:                                     vOptionB(8)
'U9:                                     vOptionB(9)
'UZ Tageszeit:                           vTextB(159)
'UZ rechtzeitig:                         vTextB(53)
'UZ rechtzeitig_Bezeichnungsfeld:        vLab(52)
'UZTagesZeit_Bezeichnungsfeld:           vLab(158)
'Ulcera:                                 vTextB(110)
'Ulcera_Bezeichnungsfeld:                vLab(112)
'UnausgefNZ:                             vCommandB(19)
'Unausgefülle:                           vCommandB(17)
'Unterzucker pM:                         vTextB(52)
'Unterzucker pM_Bezeichnungsfeld:        vLab(51)
'V1:                                     vTextB(170)
'Verformungen Beschreibung:              vTextB(93)
'Verformungen Beschreibung_BezeichnungsfevLab(95)
'Verformungen:                           vTextB(92)
'Verformungen_Bezeichnungsfeld:          vLab(94)
'Versicherung:                           vTextB(151)
'Versicherung_Bezeichnungsfeld:          vLab(153)
'Versicherungsart:                       vTextB(154)
'Vibration Großzehe:                     vTextB(120)
'Vibration Großzehe_Bezeichnungsfeld:    vLab(122)
'Vibration IK:                           vTextB(119)
'Vibration IK_Bezeichnungsfeld:          vLab(121)
'Vorgestellt:                            vTextB(149)
'Vorgestellt_Bezeichnungsfeld:           vLab(151)
'Vorname:                                vTextB(2)
'WS:                                     vTextB(129)
'WS_Bezeichnungsfeld:                    vLab(131)
'Weitere Anamnese:                       vTextB(102)
'Weitere Anamnese_Bezeichnungsfeld:      vLab(104)
'Weitere Befunde:                        vTextB(137)
'Weitere Befunde_Bezeichnungsfeld:       vLab(139)
'Weitere Medikation:                     vTextB(105)
'Weitere Medikation_Bezeichnungsfeld:    vLab(107)
'Zähne:                                  vTextB(133)
'Zähne_Bezeichnungsfeld:                 vLab(135)
'alleDatensätze:                         vCommandB(18)
'andere Nierenerkrankung:                vTextB(73)
'andere Nierenerkrankung_BezeichnungsfeldvLab(73)
'erhöht?:                                vTextB(71)
'erhöht?_Bezeichnungsfeld:               vLab(70)
'gemessen am:                            vTextB(44)
'gemessen am_Bezeichnungsfeld:           vLab(43)
'leBeh:                                  vTextB(181)
'lebehBez:                               vLab(181)
'letzte Änderung:                        vTextB(143)
'letzte Änderung_Bezeichnung:            vLab(146)
'letztes HbA1c:                          vTextB(43)
'letztes HbA1c_Bezeichnungsfeld:         vLab(42)
'obAn1eing:                              vCheckb(6)
'obAn2eing:                              vCheckb(7)
'obAnAeing:                              vCheckb(8)
'obBZausgew:                             vCheckb(10)
'obBZausgew_Bez:                         vLab(166)
'obCheck:                                vCheckb(9)
'obDMPaufgekl:                           vTextB(171)
'obDMPaufgekl_Bez:                       vLab(170)
'obMBlAusgeh:                            vCheckb(13)
'obMBlAusgeh_Bez:                        vLab(169)
'obMedNetz:                              vCheckb(14)
'obMedNetz_Bezeichnung:                  vLab(174)
'obOSaufgek:                             vCheckb(11)
'obOSaufgek_Bez:                         vLab(167)
'obPodAufgek:                            vCheckb(12)
'obPodAufgek_Bez:                        vLab(168)
'obSchulaufgek:                          vTextB(172)
'pAVK Beschreibung:                      vTextB(88)
'pAVK Beschreibung_Bezeichnungsfeld:     vLab(90)
'vorherige Werte:                        vTextB(45)
'vorherige Werte_Bezeichnungsfeld:       vLab(44)
