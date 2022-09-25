VERSION 5.00
Begin VB.Form AnBog
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "AnBog"
   ClientHeight    =   13230
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   19110
   KeyPreview = -1         'True
   LinkTopic = "AnBog"
   MaxButton = 0           'False
   MinButton = 0           'False
   ScaleHeight = 465
   ScaleWidth = 1680
   Begin VB.TextBox vTextB
      DataField = "Pat_ID"
      Index           =   1
      TabIndex        =   0
      BackColor       =   12632256
      Height          =   238
      Left            =   113
      Top             =   57
      Width           =   516
   End
   Begin VB.TextBox vTextB
      DataField = "Vorname"
      Index           =   2
      TabIndex        =   2
      BackColor       =   65280
      Height          =   238
      Left            =   2834
      Top             =   56
      Width           =   2040
   End
   Begin VB.TextBox vTextB
      DataField = "GebDat"
      Index           =   3
      TabIndex        =   3
      BackColor       =   12632256
      Height          =   238
      Left            =   850
      Top             =   340
      Width           =   1017
   End
   Begin VB.Label vLab
      Caption         =   "GebDat"
      Index           =   1
      BackColor       =   65535
      Height          =   238
      Left            =   113
      Top             =   340
      Width           =   687
   End
   Begin VB.TextBox vTextB
      DataField = "Diabetestyp"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   4
      TabIndex        =   7
      BackColor       =   33023
      Height          =   285
      Left            =   114
      Top             =   879
      Width           =   1140
   End
   Begin VB.Label vLab
      Caption         =   "Diabetestyp&:"
      Index           =   2
      BackColor       =   65535
      Height          =   240
      Left            =   114
      Top             =   639
      Width           =   1140
   End
   Begin VB.TextBox vTextB
      DataField = "Diabetes seit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   5
      TabIndex        =   8
      BackColor       =   33023
      Height          =   285
      Left            =   1299
      Top             =   879
      Width           =   1125
   End
   Begin VB.Label vLab
      Caption         =   "Diabetes seit"
      Index           =   3
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1299
      Top             =   639
      Width           =   1125
   End
   Begin VB.TextBox vTextB
      DataField = "Tabletten seit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   6
      TabIndex        =   9
      BackColor       =   33023
      Height          =   285
      Left            =   2484
      Top             =   879
      Width           =   1140
   End
   Begin VB.Label vLab
      Caption         =   "Tabletten seit"
      Index           =   4
      BackColor       =   -2147483633
      Height          =   240
      Left            =   2484
      Top             =   639
      Width           =   1140
   End
   Begin VB.TextBox vTextB
      DataField = "Insulin seit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   7
      TabIndex        =   10
      BackColor       =   33023
      Height          =   285
      Left            =   3669
      Top             =   879
      Width           =   1023
   End
   Begin VB.Label vLab
      Caption         =   "Insulin seit"
      Index           =   5
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3669
      Top             =   639
      Width           =   1023
   End
   Begin VB.TextBox vTextB
      DataField = "Grund für Vorstellung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   8
      TabIndex        =   11
      BackColor       =   -2147483643
      Height          =   645
      Left            =   114
      Top             =   1404
      Width           =   7638
   End
   Begin VB.Label vLab
      Caption         =   "Grund für Vorstellung"
      Index           =   6
      BackColor       =   -2147483633
      Height          =   240
      Left            =   114
      Top             =   1164
      Width           =   1578
   End
   Begin VB.TextBox vTextB
      DataField = "Familienanamnese"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   9
      TabIndex        =   12
      BackColor       =   -2147483643
      Height          =   645
      Left            =   114
      Top             =   2289
      Width           =   7638
   End
   Begin VB.Label vLab
      Caption         =   "Familienanamnese"
      Index           =   7
      BackColor       =   -2147483633
      Height          =   240
      Left            =   114
      Top             =   2049
      Width           =   7638
   End
   Begin VB.TextBox vTextB
      DataField = "Größe"
      Index           =   10
      TabIndex        =   13
      BackColor       =   10314363
      Height          =   238
      Left            =   794
      Top             =   3005
      Width           =   840
   End
   Begin VB.Label vLab
      Caption         =   "Gr&öße"
      Index           =   8
      BackColor       =   65535
      Height          =   238
      Left            =   113
      Top             =   3005
      Width           =   570
   End
   Begin VB.TextBox vTextB
      DataField = "Gewicht"
      Index           =   11
      TabIndex        =   14
      BackColor       =   10314363
      Height          =   238
      Left            =   2438
      Top             =   3005
      Width           =   795
   End
   Begin VB.Label vLab
      Caption         =   "Gewicht"
      Index           =   9
      BackColor       =   -2147483633
      Height          =   238
      Left            =   1701
      Top             =   3005
      Width           =   690
   End
   Begin VB.TextBox vTextB
      DataField = "Tendenz"
      Index           =   12
      TabIndex        =   15
      BackColor       =   10314363
      Height          =   238
      Left            =   4081
      Top             =   3004
      Width           =   450
   End
   Begin VB.Label vLab
      Caption         =   "Tendenz"
      Index           =   10
      BackColor       =   -2147483633
      Height          =   238
      Left            =   3288
      Top             =   3005
      Width           =   750
   End
   Begin VB.TextBox vTextB
      DataField = "DiabetesMedikament 1"
      Index           =   13
      TabIndex        =   16
      BackColor       =   -2147483643
      Height          =   256
      Left            =   1019
      Top             =   3344
      Width           =   3633
   End
   Begin VB.Label vLab
      Caption         =   "Diab&.Med. 1"
      Index           =   11
      BackColor       =   65535
      Height          =   256
      Left            =   56
      Top             =   3344
      Width           =   963
   End
   Begin VB.TextBox vTextB
      DataField = "DiabetesMedikament 1 Menge"
      Index           =   14
      TabIndex        =   17
      BackColor       =   -2147483643
      Height          =   256
      Left            =   5385
      Top             =   3346
      Width           =   2490
   End
   Begin VB.Label vLab
      Caption         =   "Menge"
      Index           =   12
      BackColor       =   -2147483633
      Height          =   256
      Left            =   4762
      Top             =   3346
      Width           =   615
   End
   Begin VB.TextBox vTextB
      DataField = "DiabetesMedikament 2"
      Index           =   15
      TabIndex        =   18
      BackColor       =   -2147483643
      Height          =   256
      Left            =   737
      Top             =   3614
      Width           =   3915
   End
   Begin VB.Label vLab
      Caption         =   "2"
      Index           =   13
      BackColor       =   -2147483633
      Height          =   256
      Left            =   56
      Top             =   3614
      Width           =   630
   End
   Begin VB.TextBox vTextB
      DataField = "DiabetesMedikament 2 Menge"
      Index           =   16
      TabIndex        =   19
      BackColor       =   -2147483643
      Height          =   256
      Left            =   5385
      Top             =   3616
      Width           =   2478
   End
   Begin VB.Label vLab
      Caption         =   "Menge"
      Index           =   14
      BackColor       =   -2147483633
      Height          =   256
      Left            =   4705
      Top             =   3616
      Width           =   663
   End
   Begin VB.TextBox vTextB
      DataField = "DiabetesMedikament 3"
      Index           =   17
      TabIndex        =   20
      BackColor       =   -2147483643
      Height          =   256
      Left            =   737
      Top             =   3883
      Width           =   3915
   End
   Begin VB.Label vLab
      Caption         =   "3"
      Index           =   15
      BackColor       =   -2147483633
      Height          =   256
      Left            =   56
      Top             =   3883
      Width           =   630
   End
   Begin VB.TextBox vTextB
      DataField = "DiabetesMedikament 3 Menge"
      Index           =   18
      TabIndex        =   21
      BackColor       =   -2147483643
      Height          =   256
      Left            =   5385
      Top             =   3885
      Width           =   2490
   End
   Begin VB.Label vLab
      Caption         =   "Menge"
      Index           =   16
      BackColor       =   -2147483633
      Height          =   256
      Left            =   4762
      Top             =   3885
      Width           =   630
   End
   Begin VB.TextBox vTextB
      DataField = "DiabetesMedikament 4"
      Index           =   19
      TabIndex        =   22
      BackColor       =   -2147483643
      Height          =   256
      Left            =   737
      Top             =   4181
      Width           =   3018
   End
   Begin VB.Label vLab
      Caption         =   "4"
      Index           =   17
      BackColor       =   -2147483633
      Height          =   256
      Left            =   56
      Top             =   4181
      Width           =   633
   End
   Begin VB.TextBox vTextB
      DataField = "DiabetesMedikament 4 Menge"
      Index           =   20
      TabIndex        =   23
      BackColor       =   -2147483643
      Height          =   256
      Left            =   5385
      Top             =   4183
      Width           =   2490
   End
   Begin VB.Label vLab
      Caption         =   "Menge"
      Index           =   18
      BackColor       =   -2147483633
      Height          =   256
      Left            =   4648
      Top             =   4183
      Width           =   735
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "Insulinpumpe"
      TabIndex        =   24
      Index           =   1
      Height          =   165
      Left            =   1077
      Top             =   4592
      Width           =   180
   End
   Begin VB.Label vLab
      Caption         =   "Insulinpumpe"
      Index           =   19
      BackColor       =   -2147483633
      Height          =   240
      Left            =   56
      Top             =   4535
      Width           =   1005
   End
   Begin VB.TextBox vTextB
      DataField = "Insulinpumpe seit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   21
      TabIndex        =   25
      BackColor       =   -2147483643
      Height          =   285
      Left            =   1700
      Top             =   4535
      Width           =   1290
   End
   Begin VB.Label vLab
      Caption         =   "seit"
      Index           =   20
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1303
      Top             =   4535
      Width           =   345
   End
   Begin VB.TextBox vTextB
      DataField = "Insulinpumpe Marke"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   22
      TabIndex        =   26
      BackColor       =   -2147483643
      Height          =   285
      Left            =   3685
      Top             =   4479
      Width           =   4188
   End
   Begin VB.Label vLab
      Caption         =   "Marke"
      Index           =   21
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3061
      Top             =   4479
      Width           =   573
   End
   Begin VB.TextBox vTextB
      DataField = "Broteinheiten gesamt"
      Index           =   23
      TabIndex        =   27
      BackColor       =   -2147483643
      Height          =   244
      Left            =   907
      Top             =   4876
      Width           =   585
   End
   Begin VB.Label vLab
      Caption         =   "BE &gesamt"
      Index           =   22
      BackColor       =   65535
      Height          =   244
      Left            =   0
      Top             =   4876
      Width           =   855
   End
   Begin VB.TextBox vTextB
      DataField = "Broteinheiten früh"
      Index           =   24
      TabIndex        =   28
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1927
      Top             =   4875
      Width           =   795
   End
   Begin VB.Label vLab
      Caption         =   "früh"
      Index           =   23
      BackColor       =   -2147483633
      Height          =   244
      Left            =   1587
      Top             =   4875
      Width           =   345
   End
   Begin VB.TextBox vTextB
      DataField = "Broteinheiten ZM früh"
      Index           =   25
      TabIndex        =   29
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3117
      Top             =   4875
      Width           =   513
   End
   Begin VB.Label vLab
      Caption         =   "vormittags"
      Index           =   24
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2777
      Top             =   4875
      Width           =   348
   End
   Begin VB.TextBox vTextB
      DataField = "Broteinheiten mittags"
      Index           =   26
      TabIndex        =   30
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3912
      Top             =   4875
      Width           =   750
   End
   Begin VB.Label vLab
      Caption         =   "mittags"
      Index           =   25
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3685
      Top             =   4875
      Width           =   225
   End
   Begin VB.TextBox vTextB
      DataField = "Broteinheiten nachmittags"
      Index           =   27
      TabIndex        =   31
      BackColor       =   -2147483643
      Height          =   244
      Left            =   4989
      Top             =   4875
      Width           =   630
   End
   Begin VB.Label vLab
      Caption         =   "nachmittags"
      Index           =   26
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4705
      Top             =   4875
      Width           =   285
   End
   Begin VB.TextBox vTextB
      DataField = "Broteinheiten abends"
      Index           =   28
      TabIndex        =   32
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6009
      Top             =   4875
      Width           =   858
   End
   Begin VB.Label vLab
      Caption         =   "abends"
      Index           =   27
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5669
      Top             =   4875
      Width           =   348
   End
   Begin VB.TextBox vTextB
      DataField = "Broteinheiten nachts"
      Index           =   29
      TabIndex        =   33
      BackColor       =   -2147483643
      Height          =   244
      Left            =   7257
      Top             =   4762
      Width           =   615
   End
   Begin VB.Label vLab
      Caption         =   "spät"
      Index           =   28
      BackColor       =   -2147483633
      Height          =   244
      Left            =   6916
      Top             =   4762
      Width           =   345
   End
   Begin VB.TextBox vTextB
      DataField = "Essenszeit früh"
      Index           =   30
      TabIndex        =   34
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1927
      Top             =   5160
      Width           =   795
   End
   Begin VB.Label vLab
      Caption         =   "Essenszeit  fr&üh:"
      Index           =   29
      BackColor       =   65535
      Height          =   244
      Left            =   632
      Top             =   5160
      Width           =   1245
   End
   Begin VB.TextBox vTextB
      DataField = "Essenszeit mittags"
      Index           =   31
      TabIndex        =   36
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3968
      Top             =   5160
      Width           =   690
   End
   Begin VB.Label vLab
      Caption         =   "mittags"
      Index           =   30
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3685
      Top             =   5160
      Width           =   285
   End
   Begin VB.TextBox vTextB
      DataField = "Essenszeit abends"
      Index           =   32
      TabIndex        =   38
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6009
      Top             =   5160
      Width           =   870
   End
   Begin VB.Label vLab
      Caption         =   "abends"
      Index           =   31
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5669
      Top             =   5160
      Width           =   345
   End
   Begin VB.TextBox vTextB
      DataField = "Essenszeit spät"
      Index           =   33
      TabIndex        =   39
      BackColor       =   -2147483643
      Height          =   244
      Left            =   7256
      Top             =   5047
      Width           =   630
   End
   Begin VB.Label vLab
      Caption         =   "spät"
      Index           =   32
      BackColor       =   -2147483633
      Height          =   244
      Left            =   6916
      Top             =   5047
      Width           =   345
   End
   Begin VB.TextBox vTextB
      DataField = "Spritz-Eß-Abstand früh"
      Index           =   34
      TabIndex        =   40
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1927
      Top             =   5443
      Width           =   783
   End
   Begin VB.Label vLab
      Caption         =   "Spritz-Eß-Abstand früh:"
      Index           =   33
      BackColor       =   -2147483633
      Height          =   244
      Left            =   56
      Top             =   5442
      Width           =   1818
   End
   Begin VB.TextBox vTextB
      DataField = "Spritz-Eß-Abstand mittags"
      Index           =   35
      TabIndex        =   41
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3968
      Top             =   5443
      Width           =   720
   End
   Begin VB.Label vLab
      Caption         =   "mittags"
      Index           =   34
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3685
      Top             =   5443
      Width           =   225
   End
   Begin VB.TextBox vTextB
      DataField = "Spritz-Eß-Abstand abends"
      Index           =   36
      TabIndex        =   42
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6066
      Top             =   5443
      Width           =   810
   End
   Begin VB.Label vLab
      Caption         =   "abends"
      Index           =   35
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5669
      Top             =   5443
      Width           =   330
   End
   Begin VB.TextBox vTextB
      DataField = "Spritzstelle früh"
      Index           =   37
      TabIndex        =   43
      BackColor       =   12632256
      Height          =   244
      Left            =   1927
      Top             =   5727
      Width           =   795
   End
   Begin VB.Label vLab
      Caption         =   "Spritzstelle früh (B, O, A)"
      Index           =   36
      BackColor       =   -2147483633
      Height          =   244
      Left            =   56
      Top             =   5727
      Width           =   1785
   End
   Begin VB.TextBox vTextB
      DataField = "Spritzstelle mittags"
      Index           =   38
      TabIndex        =   44
      BackColor       =   12632256
      Height          =   244
      Left            =   3968
      Top             =   5727
      Width           =   723
   End
   Begin VB.Label vLab
      Caption         =   "mittags"
      Index           =   37
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3685
      Top             =   5727
      Width           =   288
   End
   Begin VB.TextBox vTextB
      DataField = "Spritzstelle abends"
      Index           =   39
      TabIndex        =   45
      BackColor       =   12632256
      Height          =   244
      Left            =   6066
      Top             =   5727
      Width           =   795
   End
   Begin VB.Label vLab
      Caption         =   "abends"
      Index           =   38
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5669
      Top             =   5727
      Width           =   345
   End
   Begin VB.TextBox vTextB
      DataField = "Spritzstelle nachts"
      Index           =   40
      TabIndex        =   46
      BackColor       =   12632256
      Height          =   244
      Left            =   7256
      Top             =   5614
      Width           =   705
   End
   Begin VB.Label vLab
      Caption         =   "nachts"
      Index           =   39
      BackColor       =   -2147483633
      Height          =   244
      Left            =   6859
      Top             =   5614
      Width           =   405
   End
   Begin VB.TextBox vTextB
      DataField = "Jahr letzte Diabetesschulung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   41
      TabIndex        =   47
      BackColor       =   -2147483643
      Height          =   285
      Left            =   2211
      Top             =   6122
      Width           =   2115
   End
   Begin VB.Label vLab
      Caption         =   "&Jahr letzte Diabetesschulung"
      Index           =   40
      BackColor       =   65535
      Height          =   240
      Left            =   56
      Top             =   6122
      Width           =   2115
   End
   Begin VB.TextBox vTextB
      DataField = "Ort Schulung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   42
      TabIndex        =   48
      BackColor       =   -2147483643
      Height          =   285
      Left            =   5102
      Top             =   6122
      Width           =   2778
   End
   Begin VB.Label vLab
      Caption         =   "Ort Schulung"
      Index           =   41
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4365
      Top             =   6122
      Width           =   738
   End
   Begin VB.TextBox vTextB
      DataField = "letztes HbA1c"
      Index           =   43
      TabIndex        =   49
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1077
      Top             =   6405
      Width           =   1065
   End
   Begin VB.Label vLab
      Caption         =   "letztes HbA1c"
      Index           =   42
      BackColor       =   -2147483633
      Height          =   244
      Left            =   56
      Top             =   6405
      Width           =   960
   End
   Begin VB.TextBox vTextB
      DataField = "gemessen am"
      Index           =   44
      TabIndex        =   50
      BackColor       =   -2147483643
      Height          =   244
      Left            =   2769
      Top             =   6405
      Width           =   1545
   End
   Begin VB.Label vLab
      Caption         =   "gem am"
      Index           =   43
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2097
      Top             =   6405
      Width           =   615
   End
   Begin VB.TextBox vTextB
      DataField = "vorherige Werte"
      Index           =   45
      TabIndex        =   51
      BackColor       =   -2147483643
      Height          =   244
      Left            =   5555
      Top             =   6405
      Width           =   2310
   End
   Begin VB.Label vLab
      Caption         =   "vorherige Werte"
      Index           =   44
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4308
      Top             =   6405
      Width           =   1170
   End
   Begin VB.TextBox vTextB
      DataField = "BZMessungen pW"
      Index           =   46
      TabIndex        =   55
      BackColor       =   -2147483643
      Height          =   244
      Left            =   850
      Top             =   7003
      Width           =   1365
   End
   Begin VB.Label vLab
      Caption         =   "Zahl pW"
      Index           =   45
      BackColor       =   -2147483633
      Height          =   244
      Left            =   113
      Top             =   7003
      Width           =   675
   End
   Begin VB.TextBox vTextB
      DataField = "BZMessungen pW ndE"
      Index           =   47
      TabIndex        =   56
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3174
      Top             =   7003
      Width           =   858
   End
   Begin VB.Label vLab
      Caption         =   "davon ndE"
      Index           =   46
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2324
      Top             =   7003
      Width           =   843
   End
   Begin VB.TextBox vTextB
      DataField = "BZMessungen p W nachts"
      Index           =   48
      TabIndex        =   57
      BackColor       =   -2147483643
      Height          =   244
      Left            =   5272
      Top             =   7003
      Width           =   915
   End
   Begin VB.Label vLab
      Caption         =   "davon nachts"
      Index           =   47
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4195
      Top             =   7003
      Width           =   1020
   End
   Begin VB.TextBox vTextB
      DataField = "Aufschreiben"
      Index           =   49
      TabIndex        =   54
      BackColor       =   16777215
      Height          =   244
      Left            =   6236
      Top             =   6720
      Width           =   570
   End
   Begin VB.Label vLab
      Caption         =   "Aufschreiben"
      Index           =   48
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5159
      Top             =   6720
      Width           =   1020
   End
   Begin VB.TextBox vTextB
      DataField = "BZWerte v d Essen"
      Index           =   50
      TabIndex        =   58
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1190
      Top             =   7371
      Width           =   1590
   End
   Begin VB.Label vLab
      Caption         =   "BZ &v d Essen"
      Index           =   49
      BackColor       =   65535
      Height          =   244
      Left            =   113
      Top             =   7371
      Width           =   1020
   End
   Begin VB.TextBox vTextB
      DataField = "BZWerte n d Essen"
      Index           =   51
      TabIndex        =   59
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3514
      Top             =   7370
      Width           =   1488
   End
   Begin VB.Label vLab
      Caption         =   "danach"
      Index           =   50
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2834
      Top             =   7371
      Width           =   633
   End
   Begin VB.TextBox vTextB
      DataField = "Unterzucker pM"
      Index           =   52
      TabIndex        =   61
      BackColor       =   -2147483643
      Height          =   244
      Left            =   680
      Top             =   7653
      Width           =   510
   End
   Begin VB.Label vLab
      Caption         =   "UZ pM"
      Index           =   51
      BackColor       =   -2147483633
      Height          =   244
      Left            =   113
      Top             =   7653
      Width           =   525
   End
   Begin VB.TextBox vTextB
      DataField = "UZ rechtzeitig"
      Index           =   53
      TabIndex        =   62
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1984
      Top             =   7653
      Width           =   330
   End
   Begin VB.Label vLab
      Caption         =   "rechtzeitig"
      Index           =   52
      BackColor       =   -2147483633
      Height          =   244
      Left            =   1246
      Top             =   7653
      Width           =   675
   End
   Begin VB.TextBox vTextB
      DataField = "Fremde Hilfe pa"
      Index           =   54
      TabIndex        =   63
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3004
      Top             =   7653
      Width           =   570
   End
   Begin VB.Label vLab
      Caption         =   "Fremde Hilfe pa"
      Index           =   53
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2381
      Top             =   7653
      Width           =   570
   End
   Begin VB.TextBox vTextB
      DataField = "Bewußtlos pa"
      Index           =   55
      TabIndex        =   64
      BackColor       =   -2147483643
      Height          =   244
      Left            =   4308
      Top             =   7653
      Width           =   450
   End
   Begin VB.Label vLab
      Caption         =   "Bewußtl pa"
      Index           =   54
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3628
      Top             =   7653
      Width           =   630
   End
   Begin VB.TextBox vTextB
      DataField = "Keto pa"
      Index           =   56
      TabIndex        =   65
      BackColor       =   -2147483643
      Height          =   244
      Left            =   5442
      Top             =   7653
      Width           =   453
   End
   Begin VB.Label vLab
      Caption         =   "Keto pa"
      Index           =   55
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4762
      Top             =   7653
      Width           =   633
   End
   Begin VB.TextBox vTextB
      DataField = "BZgr300 pM"
      Index           =   57
      TabIndex        =   66
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6746
      Top             =   7653
      Width           =   960
   End
   Begin VB.Label vLab
      Caption         =   "gr300 pM"
      Index           =   56
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5952
      Top             =   7653
      Width           =   735
   End
   Begin VB.TextBox vTextB
      DataField = "Bluthochdruck"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   58
      TabIndex        =   69
      BackColor       =   -2147483643
      Height          =   285
      Left            =   56
      Top             =   8687
      Width           =   735
   End
   Begin VB.Label vLab
      Caption         =   "Bl&uthochdruck"
      Index           =   57
      BackColor       =   65535
      Height          =   240
      Left            =   56
      Top             =   8447
      Width           =   735
   End
   Begin VB.TextBox vTextB
      DataField = "BHD seit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   59
      TabIndex        =   70
      BackColor       =   16777215
      Height          =   285
      Left            =   850
      Top             =   8687
      Width           =   1125
   End
   Begin VB.Label vLab
      Caption         =   "BHD seit"
      Index           =   58
      BackColor       =   -2147483633
      Height          =   240
      Left            =   850
      Top             =   8447
      Width           =   1125
   End
   Begin VB.TextBox vTextB
      DataField = "BHD beh mit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   60
      TabIndex        =   71
      BackColor       =   16777215
      Height          =   285
      Left            =   2035
      Top             =   8687
      Width           =   3000
   End
   Begin VB.Label vLab
      Caption         =   "BHD beh mit"
      Index           =   59
      BackColor       =   -2147483633
      Height          =   240
      Left            =   2035
      Top             =   8447
      Width           =   3000
   End
   Begin VB.TextBox vTextB
      DataField = "Blutdruckwerte"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   61
      TabIndex        =   73
      BackColor       =   -2147483643
      Height          =   285
      Left            =   6039
      Top             =   8687
      Width           =   1848
   End
   Begin VB.Label vLab
      Caption         =   "Blutdruckwerte"
      Index           =   60
      BackColor       =   -2147483633
      Height          =   240
      Left            =   6039
      Top             =   8447
      Width           =   1848
   End
   Begin VB.TextBox vTextB
      DataField = "BDselbst"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   62
      TabIndex        =   72
      BackColor       =   -2147483643
      Height          =   285
      Left            =   5102
      Top             =   8687
      Width           =   900
   End
   Begin VB.Label vLab
      Caption         =   "BDselbst"
      Index           =   61
      BackColor       =   -2147483633
      Height          =   240
      Left            =   5102
      Top             =   8447
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = "Schwanger"
      Index           =   63
      TabIndex        =   67
      BackColor       =   12632256
      Height          =   244
      Left            =   1077
      Top             =   8050
      Width           =   630
   End
   Begin VB.Label vLab
      Caption         =   "Schwanger"
      Index           =   62
      BackColor       =   -2147483633
      Height          =   244
      Left            =   113
      Top             =   8050
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = "Schwanger seit"
      Index           =   64
      TabIndex        =   68
      BackColor       =   12632256
      Height          =   244
      Left            =   3004
      Top             =   8050
      Width           =   1185
   End
   Begin VB.Label vLab
      Caption         =   "Schwanger seit"
      Index           =   63
      BackColor       =   -2147483633
      Height          =   244
      Left            =   1757
      Top             =   8050
      Width           =   1185
   End
   Begin VB.TextBox vTextB
      DataField = "Augensp zuletzt"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   65
      TabIndex        =   74
      BackColor       =   -2147483643
      Height          =   285
      Left            =   56
      Top             =   9310
      Width           =   1140
   End
   Begin VB.Label vLab
      Caption         =   "Augens&p zuletzt"
      Index           =   64
      BackColor       =   65535
      Height          =   240
      Left            =   60
      Top             =   9074
      Width           =   1140
   End
   Begin VB.TextBox vTextB
      DataField = "Augensp Befund"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   66
      TabIndex        =   75
      BackColor       =   -2147483643
      Height          =   285
      Left            =   1256
      Top             =   9310
      Width           =   3453
   End
   Begin VB.Label vLab
      Caption         =   "Augensp Befund"
      Index           =   65
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1256
      Top             =   9070
      Width           =   3453
   End
   Begin VB.TextBox vTextB
      DataField = "Netzhaut gelasert"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   67
      TabIndex        =   76
      BackColor       =   -2147483643
      Height          =   285
      Left            =   4762
      Top             =   9310
      Width           =   1260
   End
   Begin VB.Label vLab
      Caption         =   "Netzhaut gelasert"
      Index           =   66
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4762
      Top             =   9070
      Width           =   1260
   End
   Begin VB.TextBox vTextB
      DataField = "Sehminderung unbehebbar"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   68
      TabIndex        =   77
      BackColor       =   -2147483643
      Height          =   285
      Left            =   6055
      Top             =   9310
      Width           =   1800
   End
   Begin VB.Label vLab
      Caption         =   "Sehminderung unbehebbar"
      Index           =   67
      BackColor       =   -2147483633
      Height          =   240
      Left            =   6055
      Top             =   9070
      Width           =   1800
   End
   Begin VB.TextBox vTextB
      DataField = "Diabet Nierenschaden"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   69
      TabIndex        =   78
      BackColor       =   -2147483643
      Height          =   285
      Left            =   56
      Top             =   9933
      Width           =   900
   End
   Begin VB.Label vLab
      Caption         =   "Diabet Nierenschaden"
      Index           =   68
      BackColor       =   -2147483633
      Height          =   240
      Left            =   56
      Top             =   9693
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = "Albumin zuletzt"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   70
      TabIndex        =   79
      BackColor       =   -2147483643
      Height          =   285
      Left            =   1020
      Top             =   9933
      Width           =   1125
   End
   Begin VB.Label vLab
      Caption         =   "Albumin zuletzt"
      Index           =   69
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1020
      Top             =   9693
      Width           =   1125
   End
   Begin VB.TextBox vTextB
      DataField = "erhöht?"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   71
      TabIndex        =   80
      BackColor       =   -2147483643
      Height          =   285
      Left            =   2205
      Top             =   9933
      Width           =   1185
   End
   Begin VB.Label vLab
      Caption         =   "erhöht?"
      Index           =   70
      BackColor       =   -2147483633
      Height          =   240
      Left            =   2205
      Top             =   9693
      Width           =   1185
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "Dialyse"
      TabIndex        =   81
      Index           =   2
      Height          =   285
      Left            =   3390
      Top             =   9933
      Width           =   288
   End
   Begin VB.Label vLab
      Caption         =   "Dialyse"
      Index           =   71
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3390
      Top             =   9693
      Width           =   288
   End
   Begin VB.TextBox vTextB
      DataField = "Dialyse seit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   72
      TabIndex        =   82
      BackColor       =   -2147483643
      Height          =   285
      Left            =   3664
      Top             =   9933
      Width           =   1155
   End
   Begin VB.Label vLab
      Caption         =   "Dialyse seit"
      Index           =   72
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3664
      Top             =   9693
      Width           =   1155
   End
   Begin VB.TextBox vTextB
      DataField = "andere Nierenerkrankung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   73
      TabIndex        =   83
      BackColor       =   -2147483643
      Height          =   285
      Left            =   4822
      Top             =   9933
      Width           =   3030
   End
   Begin VB.Label vLab
      Caption         =   "andere Nierenerkrankung"
      Index           =   73
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4818
      Top             =   9689
      Width           =   3030
   End
   Begin VB.TextBox vTextB
      DataField = "Herzkrankheit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   74
      TabIndex        =   84
      BackColor       =   -2147483643
      Height          =   285
      Left            =   56
      Top             =   10500
      Width           =   795
   End
   Begin VB.Label vLab
      Caption         =   "&Herzkrankheit"
      Index           =   74
      BackColor       =   65535
      Height          =   240
      Left            =   60
      Top             =   10259
      Width           =   795
   End
   Begin VB.TextBox vTextB
      DataField = "Angina pectoris"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   75
      TabIndex        =   85
      BackColor       =   16777215
      Height          =   285
      Left            =   907
      Top             =   10500
      Width           =   795
   End
   Begin VB.Label vLab
      Caption         =   "Angina pectoris"
      Index           =   75
      BackColor       =   -2147483633
      Height          =   240
      Left            =   907
      Top             =   10260
      Width           =   855
   End
   Begin VB.TextBox vTextB
      DataField = "Herzinfarkt"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   76
      TabIndex        =   86
      BackColor       =   16777215
      Height          =   285
      Left            =   1757
      Top             =   10500
      Width           =   513
   End
   Begin VB.Label vLab
      Caption         =   "Herzinfakt"
      Index           =   76
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1757
      Top             =   10260
      Width           =   573
   End
   Begin VB.TextBox vTextB
      DataField = "Herzinfarkt wann"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   77
      TabIndex        =   87
      BackColor       =   16777215
      Height          =   285
      Left            =   2324
      Top             =   10500
      Width           =   675
   End
   Begin VB.Label vLab
      Caption         =   "wann"
      Index           =   77
      BackColor       =   -2147483633
      Height          =   240
      Left            =   2324
      Top             =   10260
      Width           =   675
   End
   Begin VB.TextBox vTextB
      DataField = "PTCA oder Stent"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   78
      TabIndex        =   88
      BackColor       =   16777215
      Height          =   285
      Left            =   3061
      Top             =   10500
      Width           =   510
   End
   Begin VB.Label vLab
      Caption         =   "PTCA oder Stent"
      Index           =   78
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3061
      Top             =   10260
      Width           =   510
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "Bypass kardial"
      TabIndex        =   89
      Index           =   3
      Height          =   285
      Left            =   3571
      Top             =   10500
      Width           =   1095
   End
   Begin VB.Label vLab
      Caption         =   "Bypass kardial"
      Index           =   79
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3571
      Top             =   10260
      Width           =   480
   End
   Begin VB.TextBox vTextB
      DataField = "Bypass wann"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   79
      TabIndex        =   90
      BackColor       =   16777215
      Height          =   285
      Left            =   3798
      Top             =   10500
      Width           =   795
   End
   Begin VB.Label vLab
      Caption         =   "wann"
      Index           =   80
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4081
      Top             =   10261
      Width           =   525
   End
   Begin VB.TextBox vTextB
      DataField = "Herzschwäche"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   80
      TabIndex        =   91
      BackColor       =   16777215
      Height          =   285
      Left            =   4648
      Top             =   10487
      Width           =   963
   End
   Begin VB.Label vLab
      Caption         =   "Herzschwäche"
      Index           =   81
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4648
      Top             =   10260
      Width           =   963
   End
   Begin VB.TextBox vTextB
      DataField = "Herzkrankheit Beschreibung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   81
      TabIndex        =   92
      BackColor       =   16777215
      Height          =   285
      Left            =   5585
      Top             =   10500
      Width           =   2265
   End
   Begin VB.Label vLab
      Caption         =   "Herzkrankheit Beschreibung"
      Index           =   82
      BackColor       =   -2147483633
      Height          =   240
      Left            =   5585
      Top             =   10260
      Width           =   2265
   End
   Begin VB.TextBox vTextB
      DataField = "Hirndurchblutungsstörung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   82
      TabIndex        =   93
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7880
      Top             =   240
      Width           =   975
   End
   Begin VB.Label vLab
      Caption         =   "H&irndurchblutungsstörung"
      Index           =   83
      BackColor       =   65535
      Height          =   240
      Left            =   7880
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox vTextB
      DataField = "Schlaganfall"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   83
      TabIndex        =   94
      BackColor       =   -2147483643
      Height          =   285
      Left            =   8900
      Top             =   240
      Width           =   2310
   End
   Begin VB.Label vLab
      Caption         =   "Schlaganfall"
      Index           =   84
      BackColor       =   -2147483633
      Height          =   240
      Left            =   8900
      Top             =   0
      Width           =   2310
   End
   Begin VB.TextBox vTextB
      DataField = "Beindurchblutungsstörung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   84
      TabIndex        =   95
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11225
      Top             =   240
      Width           =   858
   End
   Begin VB.Label vLab
      Caption         =   "B&eindurchblutungsstörung"
      Index           =   85
      BackColor       =   65535
      Height          =   240
      Left            =   11225
      Top             =   0
      Width           =   858
   End
   Begin VB.TextBox vTextB
      DataField = "Schaufensterkrankheit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   85
      TabIndex        =   96
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12132
      Top             =   240
      Width           =   900
   End
   Begin VB.Label vLab
      Caption         =   "Schaufensterkrankheit"
      Index           =   86
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12132
      Top             =   0
      Width           =   900
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "Bypaß peripher"
      TabIndex        =   97
      Index           =   4
      Height          =   285
      Left            =   13096
      Top             =   240
      Width           =   270
   End
   Begin VB.Label vLab
      Caption         =   "Bypaß peripher"
      Index           =   87
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13096
      Top             =   0
      Width           =   225
   End
   Begin VB.TextBox vTextB
      DataField = "Geschwür"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   86
      TabIndex        =   98
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13322
      Top             =   240
      Width           =   855
   End
   Begin VB.Label vLab
      Caption         =   "Geschwür"
      Index           =   88
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13322
      Top             =   0
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = "Amputation"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   87
      TabIndex        =   99
      BackColor       =   -2147483643
      Height          =   285
      Left            =   14222
      Top             =   240
      Width           =   2433
   End
   Begin VB.Label vLab
      Caption         =   "Amputation"
      Index           =   89
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14222
      Top             =   0
      Width           =   2433
   End
   Begin VB.TextBox vTextB
      DataField = "pAVK Beschreibung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   88
      TabIndex        =   100
      BackColor       =   -2147483643
      Height          =   645
      Left            =   7880
      Top             =   750
      Width           =   7638
   End
   Begin VB.Label vLab
      Caption         =   "pAVK Beschreibung"
      Index           =   90
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7880
      Top             =   510
      Width           =   7638
   End
   Begin VB.TextBox vTextB
      DataField = "Ameisenlaufen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   89
      TabIndex        =   101
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7880
      Top             =   1635
      Width           =   465
   End
   Begin VB.Label vLab
      Caption         =   "Ameisenlaufe&n"
      Index           =   91
      BackColor       =   65535
      Height          =   240
      Left            =   7880
      Top             =   1395
      Width           =   1140
   End
   Begin VB.TextBox vTextB
      DataField = "Ameisen Ausmaß"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   90
      TabIndex        =   102
      BackColor       =   -2147483643
      Height          =   285
      Left            =   8390
      Top             =   1657
      Width           =   2100
   End
   Begin VB.Label vLab
      Caption         =   "Ausmaß"
      Index           =   92
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9080
      Top             =   1417
      Width           =   675
   End
   Begin VB.TextBox vTextB
      DataField = "Druckstellen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   91
      TabIndex        =   103
      BackColor       =   -2147483643
      Height          =   285
      Left            =   10530
      Top             =   1657
      Width           =   525
   End
   Begin VB.Label vLab
      Caption         =   "Druckstellen"
      Index           =   93
      BackColor       =   -2147483633
      Height          =   240
      Left            =   10091
      Top             =   1417
      Width           =   960
   End
   Begin VB.TextBox vTextB
      DataField = "Verformungen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   92
      TabIndex        =   104
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11111
      Top             =   1657
      Width           =   510
   End
   Begin VB.Label vLab
      Caption         =   "Verformungen"
      Index           =   94
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11055
      Top             =   1417
      Width           =   1020
   End
   Begin VB.TextBox vTextB
      DataField = "Verformungen Beschreibung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   93
      TabIndex        =   105
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11678
      Top             =   1657
      Width           =   1413
   End
   Begin VB.Label vLab
      Caption         =   "Beschreibung"
      Index           =   95
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12132
      Top             =   1417
      Width           =   828
   End
   Begin VB.TextBox vTextB
      DataField = "Neue Fußkomplikationen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   94
      TabIndex        =   109
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7937
      Top             =   2281
      Width           =   2265
   End
   Begin VB.Label vLab
      Caption         =   "Neue Fußkomplikationen"
      Index           =   96
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7937
      Top             =   2041
      Width           =   2265
   End
   Begin VB.TextBox vTextB
      DataField = "Entleerungsstörungen Magen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   95
      TabIndex        =   110
      BackColor       =   -2147483643
      Height          =   285
      Left            =   10266
      Top             =   2281
      Width           =   1590
   End
   Begin VB.Label vLab
      Caption         =   "Entleerungsst Magen"
      Index           =   97
      BackColor       =   -2147483633
      Height          =   240
      Left            =   10261
      Top             =   2044
      Width           =   1575
   End
   Begin VB.TextBox vTextB
      DataField = "Entleerungsstörungen Harnblase"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   96
      TabIndex        =   111
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11853
      Top             =   2281
      Width           =   1155
   End
   Begin VB.Label vLab
      Caption         =   "Harnblase"
      Index           =   98
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11853
      Top             =   2041
      Width           =   1155
   End
   Begin VB.TextBox vTextB
      DataField = "Schwindel Aufstehen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   97
      TabIndex        =   112
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13044
      Top             =   2281
      Width           =   1248
   End
   Begin VB.Label vLab
      Caption         =   "Schwindel Aufstehen"
      Index           =   99
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13044
      Top             =   2041
      Width           =   1248
   End
   Begin VB.TextBox vTextB
      DataField = "Folgeerkrankungen Haut"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   98
      TabIndex        =   113
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7937
      Top             =   2848
      Width           =   1830
   End
   Begin VB.Label vLab
      Caption         =   "Folgeerkrankungen Haut"
      Index           =   100
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7937
      Top             =   2608
      Width           =   1830
   End
   Begin VB.TextBox vTextB
      DataField = "Bewegungseinschränkungen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   99
      TabIndex        =   114
      BackColor       =   9803263
      Height          =   285
      Left            =   9808
      Top             =   2848
      Width           =   2310
   End
   Begin VB.Label vLab
      Caption         =   "Bewegungseinschränkungen"
      Index           =   101
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9808
      Top             =   2608
      Width           =   2310
   End
   Begin VB.TextBox vTextB
      DataField = "Sexualstörung"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   100
      TabIndex        =   115
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12188
      Top             =   2848
      Width           =   1080
   End
   Begin VB.Label vLab
      Caption         =   "Sexualstörung"
      Index           =   102
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12188
      Top             =   2608
      Width           =   1080
   End
   Begin VB.TextBox vTextB
      DataField = "Sexualstörung seit"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   101
      TabIndex        =   116
      BackColor       =   16777215
      Height          =   285
      Left            =   13268
      Top             =   2848
      Width           =   1938
   End
   Begin VB.Label vLab
      Caption         =   "Sexualstörung seit"
      Index           =   103
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13268
      Top             =   2608
      Width           =   1938
   End
   Begin VB.TextBox vTextB
      DataField = "Weitere Anamnese"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   102
      TabIndex        =   117
      BackColor       =   -2147483643
      Height          =   645
      Left            =   7937
      Top             =   3415
      Width           =   7638
   End
   Begin VB.Label vLab
      Caption         =   "&Weitere Anamnese"
      Index           =   104
      BackColor       =   65535
      Height          =   240
      Left            =   7937
      Top             =   3175
      Width           =   7638
   End
   Begin VB.TextBox vTextB
      DataField = "Alkohol"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   103
      TabIndex        =   119
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11055
      Top             =   4082
      Width           =   4485
   End
   Begin VB.Label vLab
      Caption         =   "Alk&ohol"
      Index           =   105
      BackColor       =   65535
      Height          =   240
      Left            =   10318
      Top             =   4082
      Width           =   690
   End
   Begin VB.TextBox vTextB
      DataField = "Tabak"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   104
      TabIndex        =   118
      BackColor       =   -2147483643
      Height          =   285
      Left            =   8617
      Top             =   4082
      Width           =   1698
   End
   Begin VB.Label vLab
      Caption         =   "Tabak"
      Index           =   106
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7937
      Top             =   4082
      Width           =   633
   End
   Begin VB.TextBox vTextB
      DataField = "Weitere Medikation"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   105
      TabIndex        =   120
      BackColor       =   -2147483643
      Height          =   645
      Left            =   7937
      Top             =   4605
      Width           =   7638
   End
   Begin VB.Label vLab
      Caption         =   "Weitere Medikation"
      Index           =   107
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7937
      Top             =   4365
      Width           =   1653
   End
   Begin VB.TextBox vTextB
      DataField = "Liphypertrophien Abdomen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   106
      TabIndex        =   121
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7937
      Top             =   5512
      Width           =   2310
   End
   Begin VB.Label vLab
      Caption         =   "Liph&ypertrophien Abdomen"
      Index           =   108
      BackColor       =   65535
      Height          =   240
      Left            =   7937
      Top             =   5272
      Width           =   2310
   End
   Begin VB.TextBox vTextB
      DataField = "Liphypertrophien Beine"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   107
      TabIndex        =   122
      BackColor       =   -2147483643
      Height          =   285
      Left            =   10247
      Top             =   5512
      Width           =   2310
   End
   Begin VB.Label vLab
      Caption         =   "Liphypertrophien Beine"
      Index           =   109
      BackColor       =   -2147483633
      Height          =   240
      Left            =   10247
      Top             =   5272
      Width           =   2310
   End
   Begin VB.TextBox vTextB
      DataField = "Liphypertrophien Arme"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   108
      TabIndex        =   123
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12557
      Top             =   5512
      Width           =   3018
   End
   Begin VB.Label vLab
      Caption         =   "Liphypertrophien Arme"
      Index           =   110
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12557
      Top             =   5272
      Width           =   3018
   End
   Begin VB.TextBox vTextB
      DataField = "Hyperkeratosen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   109
      TabIndex        =   125
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7993
      Top             =   6419
      Width           =   3285
   End
   Begin VB.Label vLab
      Caption         =   "Hyper&keratosen"
      Index           =   111
      BackColor       =   65535
      Height          =   240
      Left            =   7993
      Top             =   6179
      Width           =   3285
   End
   Begin VB.TextBox vTextB
      DataField = "Ulcera"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   110
      TabIndex        =   126
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11338
      Top             =   6419
      Width           =   1485
   End
   Begin VB.Label vLab
      Caption         =   "Ulcera"
      Index           =   112
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11338
      Top             =   6179
      Width           =   1485
   End
   Begin VB.TextBox vTextB
      DataField = "Kraft Zehenheber"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   111
      TabIndex        =   127
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12869
      Top             =   6419
      Width           =   915
   End
   Begin VB.Label vLab
      Caption         =   "Kraft Zehenheber"
      Index           =   113
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12869
      Top             =   6179
      Width           =   915
   End
   Begin VB.TextBox vTextB
      DataField = "Kraft Zehenbeuger"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   112
      TabIndex        =   128
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13833
      Top             =   6419
      Width           =   843
   End
   Begin VB.Label vLab
      Caption         =   "Kraft Zehenbeuger"
      Index           =   114
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13833
      Top             =   6179
      Width           =   843
   End
   Begin VB.TextBox vTextB
      DataField = "Kraft Knie"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   113
      TabIndex        =   129
      BackColor       =   -2147483643
      Height          =   285
      Left            =   14740
      Top             =   6419
      Width           =   855
   End
   Begin VB.Label vLab
      Caption         =   "Kraft Knie"
      Index           =   115
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14740
      Top             =   6179
      Width           =   855
   End
   Begin VB.TextBox vTextB
      DataField = "ASR"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   114
      TabIndex        =   130
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7993
      Top             =   6986
      Width           =   1080
   End
   Begin VB.Label vLab
      Caption         =   "ASR"
      Index           =   116
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   6746
      Width           =   1080
   End
   Begin VB.TextBox vTextB
      DataField = "PSR"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   115
      TabIndex        =   131
      BackColor       =   -2147483643
      Height          =   285
      Left            =   9127
      Top             =   6986
      Width           =   990
   End
   Begin VB.Label vLab
      Caption         =   "PSR"
      Index           =   117
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9127
      Top             =   6746
      Width           =   990
   End
   Begin VB.TextBox vTextB
      DataField = "Oberflächensensibilität"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   116
      TabIndex        =   132
      BackColor       =   -2147483643
      Height          =   285
      Left            =   10148
      Top             =   6986
      Width           =   1638
   End
   Begin VB.Label vLab
      Caption         =   "Oberflächensensibilität"
      Index           =   118
      BackColor       =   -2147483633
      Height          =   240
      Left            =   10148
      Top             =   6746
      Width           =   1638
   End
   Begin VB.TextBox vTextB
      DataField = "Monofilamenttest"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   117
      TabIndex        =   133
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11848
      Top             =   6986
      Width           =   1560
   End
   Begin VB.Label vLab
      Caption         =   "Monofilamenttest"
      Index           =   119
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11848
      Top             =   6746
      Width           =   1560
   End
   Begin VB.TextBox vTextB
      DataField = "Kalt-Warm"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   118
      TabIndex        =   134
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13408
      Top             =   6986
      Width           =   1560
   End
   Begin VB.Label vLab
      Caption         =   "Kalt-Warm"
      Index           =   120
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13408
      Top             =   6746
      Width           =   1560
   End
   Begin VB.TextBox vTextB
      DataField = "Vibration IK"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   119
      TabIndex        =   135
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7993
      Top             =   7553
      Width           =   1560
   End
   Begin VB.Label vLab
      Caption         =   "Vibration IK"
      Index           =   121
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   7313
      Width           =   1560
   End
   Begin VB.TextBox vTextB
      DataField = "Vibration Großzehe"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   120
      TabIndex        =   136
      BackColor       =   -2147483643
      Height          =   285
      Left            =   9581
      Top             =   7553
      Width           =   1560
   End
   Begin VB.Label vLab
      Caption         =   "Vibration Großzehe"
      Index           =   122
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9581
      Top             =   7313
      Width           =   1560
   End
   Begin VB.TextBox vTextB
      DataField = "Puls Leiste"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   121
      TabIndex        =   137
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11225
      Top             =   7553
      Width           =   1023
   End
   Begin VB.Label vLab
      Caption         =   "Puls Leiste"
      Index           =   123
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11225
      Top             =   7313
      Width           =   1023
   End
   Begin VB.TextBox vTextB
      DataField = "Puls Kniekehle"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   122
      TabIndex        =   138
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12302
      Top             =   7553
      Width           =   900
   End
   Begin VB.Label vLab
      Caption         =   "Puls Kniekehle"
      Index           =   124
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12302
      Top             =   7313
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = "Puls Atp"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   123
      TabIndex        =   139
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13266
      Top             =   7553
      Width           =   1080
   End
   Begin VB.Label vLab
      Caption         =   "Puls Atp"
      Index           =   125
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13266
      Top             =   7313
      Width           =   1080
   End
   Begin VB.TextBox vTextB
      DataField = "Puls Adp"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   124
      TabIndex        =   140
      BackColor       =   -2147483643
      Height          =   285
      Left            =   14400
      Top             =   7553
      Width           =   1125
   End
   Begin VB.Label vLab
      Caption         =   "Puls Adp"
      Index           =   126
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14400
      Top             =   7313
      Width           =   1140
   End
   Begin VB.TextBox vTextB
      DataField = "RR"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   125
      TabIndex        =   141
      BackColor       =   -2147483643
      Height          =   285
      Left            =   8447
      Top             =   7937
      Width           =   7128
   End
   Begin VB.Label vLab
      Caption         =   "&RR"
      Index           =   127
      BackColor       =   65535
      Height          =   240
      Left            =   7993
      Top             =   7937
      Width           =   393
   End
   Begin VB.TextBox vTextB
      DataField = "Herz"
      Index           =   126
      TabIndex        =   143
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8503
      Top             =   8505
      Width           =   2655
   End
   Begin VB.Label vLab
      Caption         =   "Her&z"
      Index           =   128
      BackColor       =   65535
      Height          =   238
      Left            =   7993
      Top             =   8505
      Width           =   450
   End
   Begin VB.TextBox vTextB
      DataField = "Lunge"
      Index           =   127
      TabIndex        =   145
      BackColor       =   -2147483643
      Height          =   238
      Left            =   13548
      Top             =   8505
      Width           =   2040
   End
   Begin VB.Label vLab
      Caption         =   "Lunge"
      Index           =   129
      BackColor       =   -2147483633
      Height          =   238
      Left            =   12925
      Top             =   8505
      Width           =   555
   End
   Begin VB.TextBox vTextB
      DataField = "Bauch"
      Index           =   128
      TabIndex        =   146
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8616
      Top             =   8755
      Width           =   1968
   End
   Begin VB.Label vLab
      Caption         =   "Bauch"
      Index           =   130
      BackColor       =   -2147483633
      Height          =   238
      Left            =   7993
      Top             =   8755
      Width           =   573
   End
   Begin VB.TextBox vTextB
      DataField = "WS"
      Index           =   129
      TabIndex        =   147
      BackColor       =   -2147483643
      Height          =   238
      Left            =   10998
      Top             =   8755
      Width           =   1035
   End
   Begin VB.Label vLab
      Caption         =   "WS"
      Index           =   131
      BackColor       =   -2147483633
      Height          =   238
      Left            =   10658
      Top             =   8755
      Width           =   285
   End
   Begin VB.TextBox vTextB
      DataField = "NL"
      Index           =   130
      TabIndex        =   148
      BackColor       =   -2147483643
      Height          =   238
      Left            =   12302
      Top             =   8755
      Width           =   1080
   End
   Begin VB.Label vLab
      Caption         =   "NL"
      Index           =   132
      BackColor       =   -2147483633
      Height          =   238
      Left            =   12076
      Top             =   8755
      Width           =   240
   End
   Begin VB.TextBox vTextB
      DataField = "SD"
      Index           =   131
      TabIndex        =   149
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8220
      Top             =   8993
      Width           =   1200
   End
   Begin VB.Label vLab
      Caption         =   "SD"
      Index           =   133
      BackColor       =   -2147483633
      Height          =   238
      Left            =   7993
      Top             =   8993
      Width           =   225
   End
   Begin VB.TextBox vTextB
      DataField = "NNH"
      Index           =   132
      TabIndex        =   150
      BackColor       =   -2147483643
      Height          =   238
      Left            =   9874
      Top             =   8993
      Width           =   1290
   End
   Begin VB.Label vLab
      Caption         =   "NNH"
      Index           =   134
      BackColor       =   -2147483633
      Height          =   238
      Left            =   9467
      Top             =   8993
      Width           =   390
   End
   Begin VB.TextBox vTextB
      DataField = "Zähne"
      Index           =   133
      TabIndex        =   151
      BackColor       =   -2147483643
      Height          =   238
      Left            =   11792
      Top             =   8993
      Width           =   1023
   End
   Begin VB.Label vLab
      Caption         =   "Zähne"
      Index           =   135
      BackColor       =   -2147483633
      Height          =   238
      Left            =   11225
      Top             =   8993
      Width           =   573
   End
   Begin VB.TextBox vTextB
      DataField = "Mundhöhle"
      Index           =   134
      TabIndex        =   152
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8957
      Top             =   9243
      Width           =   1260
   End
   Begin VB.Label vLab
      Caption         =   "Mundhöhle"
      Index           =   136
      BackColor       =   -2147483633
      Height          =   238
      Left            =   7993
      Top             =   9243
      Width           =   915
   End
   Begin VB.TextBox vTextB
      DataField = "LK"
      Index           =   135
      TabIndex        =   154
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8390
      Top             =   9581
      Width           =   1380
   End
   Begin VB.Label vLab
      Caption         =   "LK"
      Index           =   137
      BackColor       =   -2147483633
      Height          =   238
      Left            =   7993
      Top             =   9581
      Width           =   330
   End
   Begin VB.TextBox vTextB
      DataField = "Neuro sonst"
      Index           =   136
      TabIndex        =   156
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8731
      Top             =   10035
      Width           =   2328
   End
   Begin VB.Label vLab
      Caption         =   "Neuro sonst"
      Index           =   138
      BackColor       =   -2147483633
      Height          =   238
      Left            =   8050
      Top             =   10035
      Width           =   618
   End
   Begin VB.TextBox vTextB
      DataField = "Weitere Befunde"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   137
      TabIndex        =   155
      BackColor       =   -2147483643
      Height          =   540
      Left            =   11162
      Top             =   9751
      Width           =   4428
   End
   Begin VB.Label vLab
      Caption         =   "Weitere Befunde"
      Index           =   139
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9807
      Top             =   9751
      Width           =   1308
   End
   Begin VB.TextBox vTextB
      DataField = "Schulung"
      Index           =   138
      TabIndex        =   158
      BackColor       =   -2147483643
      Height          =   256
      Left            =   13889
      Top             =   10320
      Width           =   855
   End
   Begin VB.Label vLab
      Caption         =   "S&chulung"
      Index           =   140
      BackColor       =   65535
      Height          =   256
      Left            =   13039
      Top             =   10320
      Width           =   840
   End
   Begin VB.TextBox vTextB
      DataField = "DMP"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   139
      TabIndex        =   159
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7993
      Top             =   10828
      Width           =   900
   End
   Begin VB.Label vLab
      Caption         =   "DMP"
      Index           =   141
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   10601
      Width           =   900
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "Ausgabe"
      TabIndex        =   160
      Index           =   1
      Height          =   291
      Left            =   15583
      Top             =   566
      Width           =   1551
   End
   Begin VB.TextBox vTextB
      DataField = "HANr"
      Index           =   140
      TabIndex        =   161
      BackColor       =   12632256
      Height          =   240
      Left            =   10148
      Top             =   10885
      Width           =   1176
   End
   Begin VB.Label vLab
      Caption         =   "HA"
      Index           =   142
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9977
      Top             =   10601
      Width           =   285
   End
   Begin VB.TextBox vTextB
      DataField = "Beinbefund"
      Index           =   141
      TabIndex        =   124
      BackColor       =   -2147483643
      Height          =   240
      Left            =   9015
      Top             =   5839
      Width           =   6576
   End
   Begin VB.Label vLab
      Caption         =   "Beinbefund"
      Index           =   143
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   5839
      Width           =   960
   End
   Begin VB.TextBox vTextB
      DataField = "=IIf([Größe]=0,"",[Gewicht]/[Größe]/[Größe]*IIf([Größe]>3,10000,1))"
      Index           =   142
      TabIndex        =   162
      BackColor       =   10314363
      Height          =   238
      Left            =   5272
      Top             =   3005
      Width           =   636
   End
   Begin VB.Label vLab
      Caption         =   "BMI"
      Index           =   144
      BackColor       =   -2147483633
      Height          =   238
      Left            =   4762
      Top             =   3005
      Width           =   450
   End
   Begin VB.Label vLab
      Caption         =   "kg/m²"
      Index           =   145
      BackColor       =   -2147483633
      Height          =   238
      Left            =   5953
      Top             =   3005
      Width           =   624
   End
   Begin VB.TextBox vTextB
      DataField = "letzte Änderung"
      Index           =   143
      TabIndex        =   163
      BackColor       =   12632256
      Height          =   227
      Left            =   17064
      Top             =   0
      Width           =   1349
   End
   Begin VB.Label vLab
      Caption         =   "letzte Änderung"
      Index           =   146
      BackColor       =   -2147483633
      Height          =   240
      Left            =   15817
      Top             =   0
      Width           =   1185
   End
   Begin VB.TextBox vTextB
      DataField = "Fußpflege"
      Index           =   144
      TabIndex        =   106
      BackColor       =   -2147483643
      Height          =   240
      Left            =   13152
      Top             =   1701
      Width           =   696
   End
   Begin VB.Label vLab
      Caption         =   "Fußplfege"
      Index           =   147
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13096
      Top             =   1417
      Width           =   795
   End
   Begin VB.TextBox vTextB
      DataField = "Podologie"
      Index           =   145
      TabIndex        =   107
      BackColor       =   -2147483643
      Height          =   240
      Left            =   13889
      Top             =   1701
      Width           =   696
   End
   Begin VB.Label vLab
      Caption         =   "Podologie"
      Index           =   148
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13889
      Top             =   1417
      Width           =   735
   End
   Begin VB.TextBox vTextB
      DataField = "Einlagen"
      Index           =   146
      TabIndex        =   108
      BackColor       =   -2147483643
      Height          =   240
      Left            =   14626
      Top             =   1701
      Width           =   906
   End
   Begin VB.Label vLab
      Caption         =   "Orthop.Einlg/Schuhe diab&'ger"
      Index           =   149
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14683
      Top             =   1417
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = ""
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   147
      TabIndex        =   164
      BackColor       =   12632256
      Height          =   4591
      Left            =   15647
      Top             =   850
      Width           =   13215
   End
   Begin VB.Label vLab
      Caption         =   "Diagnosen:"
      Index           =   150
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14400
      Top             =   1984
      Width           =   1125
   End
   Begin VB.TextBox vTextB
      DataField = ""
      Index           =   148
      TabIndex        =   165
      BackColor       =   12632256
      Height          =   227
      Left            =   10374
      Top             =   10601
      Width           =   8148
   End
   Begin VB.TextBox vTextB
      DataField = "[Vorgestellt]"
      Index           =   149
      TabIndex        =   4
      BackColor       =   12632256
      Height          =   238
      Left            =   2834
      Top             =   340
      Width           =   1011
   End
   Begin VB.Label vLab
      Caption         =   "vorgestellt"
      Index           =   151
      BackColor       =   -2147483633
      Height          =   238
      Left            =   1927
      Top             =   340
      Width           =   795
   End
   Begin VB.TextBox vTextB
      DataField = "RRTurboMed"
      Index           =   150
      TabIndex        =   142
      BackColor       =   12632256
      Height          =   240
      Left            =   8447
      Top             =   8220
      Width           =   7131
   End
   Begin VB.Label vLab
      Caption         =   "RR TurboMed:"
      Index           =   152
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   8221
      Width           =   450
   End
   Begin VB.TextBox vTextB
      DataField = "Versicherung"
      Index           =   151
      TabIndex        =   157
      BackColor       =   12632256
      Height          =   240
      Left            =   9070
      Top             =   10318
      Width           =   3846
   End
   Begin VB.Label vLab
      Caption         =   "Versicherung:"
      Index           =   153
      BackColor       =   -2147483633
      Height          =   240
      Left            =   8050
      Top             =   10318
      Width           =   960
   End
   Begin VB.TextBox vTextB
      DataField = "Titel"
      Index           =   152
      TabIndex        =   166
      BackColor       =   65280
      Height          =   240
      Left            =   4932
      Top             =   56
      Width           =   681
   End
   Begin VB.TextBox vTextB
      DataField = "Anrede"
      Index           =   153
      TabIndex        =   167
      BackColor       =   65280
      Height          =   240
      Left            =   5669
      Top             =   56
      Width           =   801
   End
   Begin VB.TextBox vTextB
      DataField = "Versicherungsart"
      Index           =   154
      TabIndex        =   168
      BackColor       =   16711808
      Height          =   240
      Left            =   6576
      Top             =   56
      Width           =   336
   End
   Begin VB.TextBox vTextB
      DataField = "Essenszeit vormittags"
      Index           =   155
      TabIndex        =   35
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3118
      Top             =   5160
      Width           =   516
   End
   Begin VB.Label vLab
      Caption         =   "vormittags"
      Index           =   154
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2721
      Top             =   5160
      Width           =   405
   End
   Begin VB.TextBox vTextB
      DataField = "Essenszeit nachmittags"
      Index           =   156
      TabIndex        =   37
      BackColor       =   -2147483643
      Height          =   244
      Left            =   4988
      Top             =   5160
      Width           =   681
   End
   Begin VB.Label vLab
      Caption         =   "nachmittags"
      Index           =   155
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4705
      Top             =   5160
      Width           =   270
   End
   Begin VB.TextBox vTextB
      DataField = "BZMessungen selbst"
      Index           =   157
      TabIndex        =   52
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1870
      Top             =   6720
      Width           =   336
   End
   Begin VB.Label vLab
      Caption         =   "BZ-&Messungen: selbst?"
      Index           =   156
      BackColor       =   65535
      Height          =   244
      Left            =   56
      Top             =   6720
      Width           =   1755
   End
   Begin VB.TextBox vTextB
      DataField = "Gerät"
      Index           =   158
      TabIndex        =   53
      BackColor       =   16777215
      Height          =   244
      Left            =   2890
      Top             =   6720
      Width           =   2091
   End
   Begin VB.Label vLab
      Caption         =   "Gerät:"
      Index           =   157
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2323
      Top             =   6720
      Width           =   525
   End
   Begin VB.TextBox vTextB
      DataField = "UZ Tageszeit"
      Index           =   159
      TabIndex        =   60
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6236
      Top             =   7258
      Width           =   1701
   End
   Begin VB.Label vLab
      Caption         =   "UZ Tageszeit"
      Index           =   158
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5102
      Top             =   7257
      Width           =   1065
   End
   Begin VB.TextBox vTextB
      DataField = "Nachname"
      Index           =   160
      TabIndex        =   1
      BackColor       =   65280
      Height          =   238
      Left            =   680
      Top             =   56
      Width           =   2100
   End
   Begin VB.TextBox vTextB
      DataField = "BeinödVen"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   161
      TabIndex        =   153
      BackColor       =   -2147483643
      Height          =   510
      Left            =   11399
      Top             =   9243
      Width           =   4191
   End
   Begin VB.Label vLab
      Caption         =   "Beinödeme&, Venenkrkht:"
      Index           =   159
      BackColor       =   65535
      Height          =   238
      Left            =   10261
      Top             =   9243
      Width           =   1080
   End
   Begin VB.TextBox vTextB
      DataField = "Carotiden"
      Index           =   162
      TabIndex        =   144
      BackColor       =   -2147483643
      Height          =   238
      Left            =   11632
      Top             =   8505
      Width           =   1230
   End
   Begin VB.Label vLab
      Caption         =   "Carotiden"
      Index           =   160
      BackColor       =   -2147483633
      Height          =   238
      Left            =   11225
      Top             =   8505
      Width           =   390
   End
   Begin VB.TextBox vTextB
      DataField = "DMSchulz"
      Index           =   163
      TabIndex        =   169
      BackColor       =   12632256
      Height          =   256
      Left            =   15817
      Top             =   10320
      Width           =   510
   End
   Begin VB.Label vLab
      Caption         =   "Dm-Schlgn:"
      Index           =   161
      BackColor       =   -2147483633
      Height          =   256
      Left            =   14850
      Top             =   10320
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = "RRSchulz"
      Index           =   164
      TabIndex        =   170
      BackColor       =   12632256
      Height          =   256
      Left            =   17351
      Top             =   10318
      Width           =   510
   End
   Begin VB.Label vLab
      Caption         =   "RR-Schlgn:"
      Index           =   162
      BackColor       =   -2147483633
      Height          =   256
      Left            =   16384
      Top             =   10318
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = "DMPhier"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   165
      TabIndex        =   171
      BackColor       =   12632256
      Height          =   285
      Left            =   8957
      Top             =   10828
      Width           =   900
   End
   Begin VB.Label vLab
      Caption         =   "DMP hier:"
      Index           =   163
      BackColor       =   -2147483633
      Height          =   240
      Left            =   8957
      Top             =   10601
      Width           =   900
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "Tkz"
      TabIndex        =   172
      Index           =   5
      Height          =   225
      Left            =   7426
      Top             =   56
      Width           =   270
   End
   Begin VB.Label vLab
      Caption         =   "Tkz"
      Index           =   164
      BackColor       =   -2147483633
      Height          =   225
      Left            =   6973
      Top             =   56
      Width           =   390
   End
   Begin VB.TextBox vTextB
      DataField = "Ther1"
      Index           =   166
      TabIndex        =   6
      BackColor       =   -2147483643
      Height          =   225
      Left            =   5782
      Top             =   626
      Width           =   750
   End
   Begin VB.Label vLab
      Caption         =   "Therapieart"
      Index           =   165
      BackColor       =   -2147483633
      Height          =   224
      Left            =   4818
      Top             =   626
      Width           =   900
   End
   Begin VB.TextBox vTextB
      DataField = "=[Nachname]+" "+[Vorname]"
      Index           =   167
      TabIndex        =   173
      BackColor       =   8421504
      Height          =   239
      Left            =   11338
      Top             =   4365
      Width           =   4202
   End
   Begin VB.TextBox vTextB
      DataField = ""
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   168
      TabIndex        =   174
      BackColor       =   12632256
      Height          =   4599
      Left            =   15760
      Top             =   5499
      Width           =   13094
   End
   Begin VB.TextBox vTextB
      DataField = ""
      Index           =   169
      TabIndex        =   175
      BackColor       =   12632256
      Height          =   227
      Left            =   12812
      Top             =   10885
      Width           =   1486
   End
   Begin VB.TextBox vTextB
      DataField = ""
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   170
      TabIndex        =   176
      BackColor       =   12632256
      Height          =   341
      Left            =   7993
      Top             =   11168
      Width           =   6353
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "An1Aufruf"
      TabIndex        =   177
      Index           =   2
      Height          =   276
      Left            =   56
      Top             =   10885
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "An2Aufruf"
      TabIndex        =   178
      Index           =   3
      Height          =   291
      Left            =   56
      Top             =   11221
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "AnAAufruf"
      TabIndex        =   179
      Index           =   4
      Height          =   321
      Left            =   3571
      Top             =   10840
      Width           =   2826
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "CheckAufruf"
      TabIndex        =   180
      Index           =   5
      Height          =   291
      Left            =   3571
      Top             =   11225
      Width           =   2826
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obAn1eing"
      TabIndex        =   181
      Index           =   6
      Height          =   240
      Left            =   2891
      Top             =   10885
      Width           =   260
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obAn2eing"
      TabIndex        =   182
      Index           =   7
      Height          =   240
      Left            =   2891
      Top             =   11245
      Width           =   260
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obAnAeing"
      TabIndex        =   183
      Index           =   8
      Height          =   240
      Left            =   6462
      Top             =   10885
      Width           =   260
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obCheck"
      TabIndex        =   184
      Index           =   9
      Height          =   240
      Left            =   6462
      Top             =   11225
      Width           =   260
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "Labor"
      TabIndex        =   185
      Index           =   6
      Height          =   336
      Left            =   6776
      Top             =   10828
      Width           =   1071
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "Brief"
      TabIndex        =   186
      Index           =   7
      Height          =   290
      Left            =   6803
      Top             =   11225
      Width           =   1011
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "Dokumente"
      TabIndex        =   187
      Index           =   8
      Height          =   336
      Left            =   56
      Top             =   11565
      Width           =   696
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "Briefe"
      TabIndex        =   188
      Index           =   9
      Height          =   336
      Left            =   1190
      Top             =   11565
      Width           =   2031
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   189
      Index           =   1
      Height          =   193
      Left            =   15649
      Top             =   5511
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   190
      Index           =   2
      Height          =   193
      Left            =   15649
      Top             =   5707
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   191
      Index           =   3
      Height          =   193
      Left            =   15649
      Top             =   5903
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   192
      Index           =   4
      Height          =   193
      Left            =   15649
      Top             =   6098
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   193
      Index           =   5
      Height          =   193
      Left            =   15649
      Top             =   6294
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   194
      Index           =   6
      Height          =   193
      Left            =   15649
      Top             =   6490
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   195
      Index           =   7
      Height          =   193
      Left            =   15649
      Top             =   6686
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   196
      Index           =   8
      Height          =   193
      Left            =   15649
      Top             =   6882
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   197
      Index           =   9
      Height          =   193
      Left            =   15649
      Top             =   7078
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   198
      Index           =   10
      Height          =   193
      Left            =   15649
      Top             =   7274
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   199
      Index           =   11
      Height          =   193
      Left            =   15649
      Top             =   7470
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   200
      Index           =   12
      Height          =   193
      Left            =   15649
      Top             =   7666
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   201
      Index           =   13
      Height          =   193
      Left            =   15649
      Top             =   7862
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   202
      Index           =   14
      Height          =   193
      Left            =   15649
      Top             =   8058
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   203
      Index           =   15
      Height          =   193
      Left            =   15649
      Top             =   8254
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   204
      Index           =   16
      Height          =   193
      Left            =   15649
      Top             =   8450
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   205
      Index           =   17
      Height          =   193
      Left            =   15649
      Top             =   8646
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   206
      Index           =   18
      Height          =   193
      Left            =   15647
      Top             =   8842
      Width           =   113
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "DReset"
      TabIndex        =   207
      Index           =   10
      Height          =   216
      Left            =   17971
      Top             =   10318
      Width           =   576
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obBZausgew"
      TabIndex        =   208
      Index           =   10
      Height          =   240
      Left            =   14400
      Top             =   10830
      Width           =   260
   End
   Begin VB.Label vLab
      Caption         =   "BZ-Meßger ausgew"
      Index           =   166
      BackColor       =   -2147483633
      Height          =   210
      Left            =   14630
      Top             =   10830
      Width           =   1470
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obOSaufgek"
      TabIndex        =   210
      Index           =   11
      Height          =   240
      Left            =   14400
      Top             =   11287
      Width           =   260
   End
   Begin VB.Label vLab
      Caption         =   "orthop. SM aufgekl."
      Index           =   167
      BackColor       =   -2147483633
      Height          =   225
      Left            =   14630
      Top             =   11282
      Width           =   1470
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obPodAufgek"
      TabIndex        =   211
      Index           =   12
      Height          =   240
      Left            =   14400
      Top             =   11532
      Width           =   260
   End
   Begin VB.Label vLab
      Caption         =   "üb.Podologie aufgek."
      Index           =   168
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14630
      Top             =   11532
      Width           =   1530
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obMBlAusgeh"
      TabIndex        =   209
      Index           =   13
      Height          =   240
      Left            =   14400
      Top             =   11080
      Width           =   260
   End
   Begin VB.Label vLab
      Caption         =   "&8: Merkbl DFS ausgehändigt"
      Index           =   169
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14626
      Top             =   11080
      Width           =   1470
   End
   Begin VB.TextBox vTextB
      DataField = "obDMPaufgekl"
      Index           =   171
      TabIndex        =   213
      BackColor       =   16777215
      Height          =   238
      Left            =   16157
      Top             =   11533
      Width           =   405
   End
   Begin VB.Label vLab
      Caption         =   " DMP-Beteiligung"
      Index           =   170
      BackColor       =   -2147483633
      Height          =   210
      Left            =   16611
      Top             =   11508
      Width           =   1140
   End
   Begin VB.TextBox vTextB
      DataField = "obSchulaufgek"
      Index           =   172
      TabIndex        =   212
      BackColor       =   16777215
      Height          =   238
      Left            =   16157
      Top             =   11281
      Width           =   405
   End
   Begin VB.Label vLab
      Caption         =   "&9: Schulungsaufkl"
      Index           =   171
      BackColor       =   -2147483633
      Height          =   210
      Left            =   16611
      Top             =   11281
      Width           =   1170
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   214
      Index           =   19
      Height          =   193
      Left            =   15647
      Top             =   9038
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   215
      Index           =   20
      Height          =   193
      Left            =   15647
      Top             =   9234
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   216
      Index           =   21
      Height          =   193
      Left            =   15647
      Top             =   9430
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   217
      Index           =   22
      Height          =   193
      Left            =   15647
      Top             =   9626
      Width           =   113
   End
      Begin VB.OptionButton vOptionB
         Caption         =   ""
      TabIndex        =   218
      Index           =   23
      Height          =   193
      Left            =   15647
      Top             =   9822
      Width           =   113
   End
   Begin VB.TextBox vTextB
      DataField = "TherAkt"
      Index           =   173
      TabIndex        =   219
      BackColor       =   16777215
      Height          =   238
      Left            =   6973
      Top             =   623
      Width           =   780
   End
   Begin VB.Label vLab
      Caption         =   "-&>"
      Index           =   172
      BackColor       =   -2147483633
      Height          =   227
      Left            =   6689
      Top             =   623
      Width           =   178
   End
   Begin VB.TextBox vTextB
      DataField = "Versicherungsart"
      Index           =   174
      TabIndex        =   220
      BackColor       =   12632256
      Height          =   226
      Left            =   7313
      Top             =   340
      Width           =   454
   End
   Begin VB.Label vLab
      Caption         =   "SchGr"
      Index           =   173
      BackColor       =   -2147483633
      Height          =   227
      Left            =   6746
      Top             =   340
      Width           =   518
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "obMedNetz"
      TabIndex        =   221
      Index           =   14
      Height          =   240
      Left            =   7483
      Top             =   907
      Width           =   260
   End
   Begin VB.Label vLab
      Caption         =   "Med.Netz:"
      Index           =   174
      BackColor       =   -2147483633
      Height          =   240
      Left            =   6633
      Top             =   907
      Width           =   795
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "BZKurven"
      TabIndex        =   222
      Index           =   11
      Height          =   336
      Left            =   3344
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "DMPAusgeb"
      TabIndex        =   223
      Index           =   12
      Height          =   290
      Left            =   17177
      Top             =   566
      Width           =   1296
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "AugenBefunde"
      TabIndex        =   224
      Index           =   13
      Height          =   337
      Left            =   5499
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "DokDown"
      TabIndex        =   225
      Index           =   14
      Height          =   291
      Left            =   793
      Top             =   11565
      Width           =   336
   End
   Begin VB.TextBox vTextB
      DataField = ""
      Index           =   175
      TabIndex        =   226
      BackColor       =   12632256
      Height          =   227
      Left            =   11735
      Top             =   10885
      Width           =   1032
   End
   Begin VB.ComboBox vComboB
      text = "
      Index           =   1
      BackColor       =   -2147483643
      Height          =   284
      Left            =   8569
      Top             =   11508
      Width           =   2599
   End
   Begin VB.ComboBox vComboB
      text = "
      Index           =   2
      BackColor       =   -2147483643
      Height          =   286
      Left            =   11791
      Top             =   11508
      Width           =   2551
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "HA1nicht"
      TabIndex        =   229
      Index           =   15
      Height          =   240
      Left            =   8333
      Top             =   11792
      Width           =   260
   End
   Begin VB.Label vLab
      Caption         =   "-"
      Index           =   175
      BackColor       =   -2147483633
      Height          =   180
      Left            =   8047
      Top             =   11735
      Width           =   240
   End
   Begin VB.CheckBox vCheckb
      Caption         =   "HA2nicht"
      TabIndex        =   230
      Index           =   16
      Height          =   240
      Left            =   11451
      Top             =   11735
      Width           =   260
   End
   Begin VB.Label vLab
      Caption         =   "-"
      Index           =   176
      BackColor       =   -2147483633
      Height          =   180
      Left            =   11168
      Top             =   11735
      Width           =   225
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "HA1_Bez"
      TabIndex        =   231
      Index           =   15
      Height          =   285
      Left            =   7990
      Top             =   11508
      Width           =   570
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "HA2_Bez"
      TabIndex        =   232
      Index           =   16
      Height          =   270
      Left            =   11168
      Top             =   11508
      Width           =   615
   End
   Begin VB.TextBox vTextB
      DataField = ""
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   176
      TabIndex        =   233
      BackColor       =   12632256
      Height          =   284
      Left            =   16440
      Top             =   10885
      Width           =   1984
   End
   Begin VB.Label vLab
      Caption         =   "&+"
      Index           =   177
      BackColor       =   -2147483633
      Height          =   285
      Left            =   16157
      Top             =   10885
      Width           =   255
   End
   Begin VB.Label vLab
      Caption         =   "Duplex"
      Index           =   178
      BackColor       =   -2147483633
      Height          =   283
      Left            =   6292
      Top             =   1026
      Width           =   850
   End
   Begin VB.TextBox vTextB
      DataField = ""
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   177
      TabIndex        =   234
      BackColor       =   12632256
      Height          =   680
      Left            =   4315
      Top             =   11962
      Width           =   4127
   End
   Begin VB.Label vLab
      Caption         =   "Doppler"
      Index           =   179
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3628
      Top             =   11962
      Width           =   630
   End
   Begin VB.TextBox vTextB
      DataField = ""
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   178
      TabIndex        =   235
      BackColor       =   12632256
      Height          =   725
      Left            =   9139
      Top             =   11905
      Width           =   4632
   End
   Begin VB.Label vLab
      Caption         =   "Duplex"
      Index           =   180
      BackColor       =   -2147483633
      Height          =   226
      Left            =   8447
      Top             =   11905
      Width           =   628
   End
   Begin VB.Label vLab
      Caption         =   "Sono"
      Index           =   181
      BackColor       =   -2147483633
      Height          =   227
      Left            =   13776
      Top             =   11848
      Width           =   462
   End
   Begin VB.TextBox vTextB
      DataField = ""
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   179
      TabIndex        =   236
      BackColor       =   12632256
      Height          =   737
      Left            =   14301
      Top             =   11905
      Width           =   4350
   End
   Begin VB.TextBox vTextB
      DataField = ""
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      Index           =   180
      TabIndex        =   237
      BackColor       =   12632256
      Height          =   680
      Left            =   0
      Top             =   11962
      Width           =   3635
   End
   Begin VB.Label vLab
      Caption         =   "-"
      Index           =   182
      BackColor       =   -2147483633
      Height          =   226
      Left            =   3855
      Top             =   340
      Width           =   134
   End
   Begin VB.TextBox vTextB
      DataField = ""
      Index           =   181
      TabIndex        =   5
      BackColor       =   12632256
      Height          =   226
      Left            =   4025
      Top             =   340
      Width           =   1190
   End
   Begin VB.ComboBox vComboB
      text = "
      Index           =   3
      BackColor       =   -2147483643
      Height          =   281
      Left            =   1231
      Top             =   56
      Width           =   3814
   End
   Begin VB.Label vLab
      Caption         =   "Daten&quelle:"
      Index           =   183
      BackColor       =   -2147483633
      Height          =   240
      Left            =   113
      Top             =   56
      Width           =   1005
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "Unausgefülle"
      TabIndex        =   1
      Index           =   17
      Height          =   291
      Left            =   5159
      Top             =   56
      Width           =   1251
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "alleDatensätze"
      TabIndex        =   3
      Index           =   18
      Height          =   291
      Left            =   8447
      Top             =   56
      Width           =   1251
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "UnausgefNZ"
      TabIndex        =   2
      Index           =   19
      Height          =   291
      Left            =   6463
      Top             =   56
      Width           =   1866
   End
   Begin VB.TextBox vTextB
      DataField = "=currentdb.Name"
      Index           =   182
      TabIndex        =   4
      BackColor       =   12632256
      Height          =   227
      Left            =   11225
      Top             =   0
      Width           =   7257
   End
   Begin VB.CommandButton vCommandB
      Caption         =   "Datenbank-Aufruf"
      TabIndex        =   5
      Index           =   20
      Height          =   291
      Left            =   11225
      Top             =   226
      Width           =   7311
   End
   Begin VB.PictureBox picStatBox
      Align = 2              'Unten ausrichten
      Appearance = 0         '2D
      BorderStyle = 0        'Kein
      ForeColor = &H80000008
      Height = 300
      Left = 0
      ScaleHeight = 300
      ScaleWidth = 7875
      TabIndex = 380
      Top = 10815
      Width = 7875
      Begin VB.CommandButton cmdLast
         Height = 300
         Left = 4545
         Picture         =   "AnTest.frx":0000
         Style = 1              'Grafisch
         TabIndex = 384
         Top = 0
         UseMaskColor = -1       'True
         Width = 345
      End
      Begin VB.CommandButton cmdNext
         Height = 300
         Left = 4200
         Picture         =   "AnTest.frx":0342
         Style = 1              'Grafisch
         TabIndex = 383
         Top = 0
         UseMaskColor = -1       'True
         Width = 345
      End
      Begin VB.CommandButton cmdPrevious
         Height = 300
         Left = 345
         Picture         =   "AnTest.frx":0684
         Style = 1              'Grafisch
         TabIndex = 382
         Top = 0
         UseMaskColor = -1       'True
         Width = 345
      End
      Begin VB.CommandButton cmdFirst
         Height = 300
         Left = 0
         Picture         =   "AnTest.frx":09C6
         Style = 1              'Grafisch
         TabIndex = 381
         Top = 0
         UseMaskColor = -1       'True
         Width = 345
      End
      Begin VB.Label lblStatus
         BackColor = &HFFFFFF
         BorderStyle = 1        'Fest Einfach
         Height = 285
         Left = 690
         TabIndex = 385
         Top = 0
         Width = 3360
      End
End
Attribute VB_Name = "AnBog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hierdurch wird die aktuelle Datensatzposition für diese Datensatzgruppe angezeigt
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub 'adoPrimaryRS_MoveComplete
Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
End Sub 'adoPrimaryRS_WillChangeRecord
Option Compare Database
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
 Call tu_brief(Me)
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
 Call do_DMPAusgeb(Me)
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
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=linux;uid=praxis;pwd=***REMOVED***;database=quelle;"
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select * from anamnesebogen", db, adOpenStatic, adLockOptimistic
  Dim oText As TextBox
  'Textfelder an Datenprovider binden
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Dim oCheck As CheckBox
  'Kontrollkästchen an Datenprovider binden
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = adoPrimaryRS
  Next
  mbDataChanged = False
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
Private Sub do_ha_click(ByVal HANr)
 Dim stelle%
 Dim rHa As DAO.Recordset
 If KVÄDatei = "" Then Call KVÄDateifind
 Set kvä = DBEngine.OpenDatabase(KVÄDatei, , True)
 Set rHa = kvä.OpenRecordset("HAe", dbOpenTable)
 rHa.index = "KVNr"
' Unterprogramm
 Dim teile$()
 teile = Split(HANr, "_")
 If UBound(teile) > -1 Then
  If InStr(teile(0), "/") = 0 Then
   teile(0) = Left(teile(0), 2) + "/" + Mid(teile(0), 3) 'Replace(teile(0), "/", "")
  End If
 End If
 If UBound(teile) = 2 Then rHa.Seek "=", teile(0), Trim(teile(1)), Trim(teile(2))
 If UBound(teile) = 1 Or rHa.NoMatch Then rHa.Seek "=", teile(0), Trim(teile(1))
 If UBound(teile) = 0 Or rHa.NoMatch Then rHa.Seek "=", teile(0)
 Dim stAppName$, dbnr$
 If Not rHa.NoMatch Then dbnr = rHa!dbnr
 If dbnr <> "" And Not IsNull(dbnr) Then
  stAppName = Environ("programfiles") + "\internet explorer\iexplore.exe http://www.kvb.de/servlet/PB/cmd/arztsuche3/index.html?&enr=" + CStr(dbnr)
  Call Shell(stAppName, vbNormalFocus)
 End If
 kvä.Close
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
 eRR.Clear
 rfDE.Seek "=", frm!Pat_id, MDIICD(nr), MDIDiag(nr)
 If eRR.Number > 0 Then Exit Function
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
'Private Sub Befehl465_Click()
'On Error GoTo Err_Befehl465_Click


'    Screen.PreviousControl.SetFocus
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

'Exit_Befehl465_Click:
'    Exit Sub

'Err_Befehl465_Click:
'    MsgBox eRR.description
'    Resume Exit_Befehl465_Click
    
'End Sub

'Private Sub Befehl466_Click()
'On Error GoTo Err_Befehl466_Click


'    Screen.PreviousControl.SetFocus
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

'Exit_Befehl466_Click:
'    Exit Sub

'Err_Befehl466_Click:
'    MsgBox eRR.description
'    Resume Exit_Befehl466_Click
    
'End Sub
Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub
  Private Sub Form_Unload(Cancel as Integer)
  Screen.MousePointer = vbDefault
  End Sub 'Private Sub Form_Unload()
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub
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
 Call tu_brief(Me)
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
 Call do_DMPAusgeb(Me)
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
Private Sub do_ha_click(ByVal HANr)
 Dim stelle%
 Dim rHa As DAO.Recordset
 If KVÄDatei = "" Then Call KVÄDateifind
 Set kvä = DBEngine.OpenDatabase(KVÄDatei, , True)
 Set rHa = kvä.OpenRecordset("HAe", dbOpenTable)
 rHa.index = "KVNr"
' Unterprogramm
 Dim teile$()
 teile = Split(HANr, "_")
 If UBound(teile) > -1 Then
  If InStr(teile(0), "/") = 0 Then
   teile(0) = Left(teile(0), 2) + "/" + Mid(teile(0), 3) 'Replace(teile(0), "/", "")
  End If
 End If
 If UBound(teile) = 2 Then rHa.Seek "=", teile(0), Trim(teile(1)), Trim(teile(2))
 If UBound(teile) = 1 Or rHa.NoMatch Then rHa.Seek "=", teile(0), Trim(teile(1))
 If UBound(teile) = 0 Or rHa.NoMatch Then rHa.Seek "=", teile(0)
 Dim stAppName$, dbnr$
 If Not rHa.NoMatch Then dbnr = rHa!dbnr
 If dbnr <> "" And Not IsNull(dbnr) Then
  stAppName = Environ("programfiles") + "\internet explorer\iexplore.exe http://www.kvb.de/servlet/PB/cmd/arztsuche3/index.html?&enr=" + CStr(dbnr)
  Call Shell(stAppName, vbNormalFocus)
 End If
 kvä.Close
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
 eRR.Clear
 rfDE.Seek "=", frm!Pat_id, MDIICD(nr), MDIDiag(nr)
 If eRR.Number > 0 Then Exit Function
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
'Private Sub Befehl465_Click()
'On Error GoTo Err_Befehl465_Click


'    Screen.PreviousControl.SetFocus
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

'Exit_Befehl465_Click:
'    Exit Sub

'Err_Befehl465_Click:
'    MsgBox eRR.description
'    Resume Exit_Befehl465_Click
    
'End Sub

'Private Sub Befehl466_Click()
'On Error GoTo Err_Befehl466_Click


'    Screen.PreviousControl.SetFocus
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

'Exit_Befehl466_Click:
'    Exit Sub

'Err_Befehl466_Click:
'    MsgBox eRR.description
'    Resume Exit_Befehl466_Click
    
'End Sub

