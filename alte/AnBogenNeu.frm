VERSION 5.00
Begin VB.Form Anamnesebogen
   BorderStyle     =   3  'Fester Dialog"
   Caption         =   "Anamnesebogen"
   ClientHeight    =   465
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   1680
   KeyPreview = -1         'True
   LinkTopic = "Anamnesebogen"
   MaxButton = 0           'False
   MinButton = 0           'False
   ScaleHeight = 465
   ScaleWidth = 1680
   Begin VB.TextBox Pat_id
      DataField = Pat_ID
      Index           =   1
      TabIndex        =   0
      BackColor       =   12632256
      Height          =   238
      Left            =   113
      Top             =   57
      Width           =   516
   End
   Begin VB.TextBox Vorname
      DataField = Vorname
      Index           =   1
      TabIndex        =   2
      BackColor       =   65280
      Height          =   238
      Left            =   2834
      Top             =   56
      Width           =   2040
   End
   Begin VB.TextBox GebDat
      DataField = GebDat
      Index           =   1
      TabIndex        =   3
      BackColor       =   12632256
      Height          =   238
      Left            =   850
      Top             =   340
      Width           =   1017
   End
   Begin VB.Label GebDat_Bezeichnungsfeld
      Caption         =   "GebDat"
      Index           =   1
      BackColor       =   65535
      Height          =   238
      Left            =   113
      Top             =   340
      Width           =   687
   End
   Begin VB.TextBox Diabetestyp
      DataField = Diabetestyp
      Index           =   2
      TabIndex        =   7
      BackColor       =   33023
      Height          =   285
      Left            =   114
      Top             =   879
      Width           =   1140
   End
   Begin VB.Label Diabetestyp_Bezeichnungsfeld
      Caption         =   "Diabetestyp&:"
      Index           =   2
      BackColor       =   65535
      Height          =   240
      Left            =   114
      Top             =   639
      Width           =   1140
   End
   Begin VB.TextBox Diabetes_seit
      DataField = Diabetes seit
      Index           =   3
      TabIndex        =   8
      BackColor       =   33023
      Height          =   285
      Left            =   1299
      Top             =   879
      Width           =   1125
   End
   Begin VB.Label Diabetes_seit_Bezeichnungsfeld
      Caption         =   "Diabetes seit"
      Index           =   3
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1299
      Top             =   639
      Width           =   1125
   End
   Begin VB.TextBox Tabletten_seit
      DataField = Tabletten seit
      Index           =   4
      TabIndex        =   9
      BackColor       =   33023
      Height          =   285
      Left            =   2484
      Top             =   879
      Width           =   1140
   End
   Begin VB.Label Tabletten_seit_Bezeichnungsfeld
      Caption         =   "Tabletten seit"
      Index           =   4
      BackColor       =   -2147483633
      Height          =   240
      Left            =   2484
      Top             =   639
      Width           =   1140
   End
   Begin VB.TextBox Insulin_seit
      DataField = Insulin seit
      Index           =   5
      TabIndex        =   10
      BackColor       =   33023
      Height          =   285
      Left            =   3669
      Top             =   879
      Width           =   1023
   End
   Begin VB.Label Insulin_seit_Bezeichnungsfeld
      Caption         =   "Insulin seit"
      Index           =   5
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3669
      Top             =   639
      Width           =   1023
   End
   Begin VB.TextBox Grund_fuer_Vorstellung
      DataField = Grund für Vorstellung
      Index           =   6
      TabIndex        =   11
      BackColor       =   -2147483643
      Height          =   645
      Left            =   114
      Top             =   1404
      Width           =   7638
   End
   Begin VB.Label Grund_fuer_Vorstellung_Bezeichnungsfeld
      Caption         =   "Grund für Vorstellung"
      Index           =   6
      BackColor       =   -2147483633
      Height          =   240
      Left            =   114
      Top             =   1164
      Width           =   1578
   End
   Begin VB.TextBox Familienanamnese
      DataField = Familienanamnese
      Index           =   7
      TabIndex        =   12
      BackColor       =   -2147483643
      Height          =   645
      Left            =   114
      Top             =   2289
      Width           =   7638
   End
   Begin VB.Label Familienanamnese_Bezeichnungsfeld
      Caption         =   "Familienanamnese"
      Index           =   7
      BackColor       =   -2147483633
      Height          =   240
      Left            =   114
      Top             =   2049
      Width           =   7638
   End
   Begin VB.TextBox Groesse
      DataField = Größe
      Index           =   8
      TabIndex        =   13
      BackColor       =   10314363
      Height          =   238
      Left            =   794
      Top             =   3005
      Width           =   840
   End
   Begin VB.Label Groesse_Bezeichnungsfeld
      Caption         =   "Gr&öße"
      Index           =   8
      BackColor       =   65535
      Height          =   238
      Left            =   113
      Top             =   3005
      Width           =   570
   End
   Begin VB.TextBox Gewicht
      DataField = Gewicht
      Index           =   9
      TabIndex        =   14
      BackColor       =   10314363
      Height          =   238
      Left            =   2438
      Top             =   3005
      Width           =   795
   End
   Begin VB.Label Gewicht_Bezeichnungsfeld
      Caption         =   "Gewicht"
      Index           =   9
      BackColor       =   -2147483633
      Height          =   238
      Left            =   1701
      Top             =   3005
      Width           =   690
   End
   Begin VB.TextBox Tendenz
      DataField = Tendenz
      Index           =   10
      TabIndex        =   15
      BackColor       =   10314363
      Height          =   238
      Left            =   4081
      Top             =   3004
      Width           =   450
   End
   Begin VB.Label Tendenz_Bezeichnungsfeld
      Caption         =   "Tendenz"
      Index           =   10
      BackColor       =   -2147483633
      Height          =   238
      Left            =   3288
      Top             =   3005
      Width           =   750
   End
   Begin VB.TextBox DiabetesMedikament_1
      DataField = DiabetesMedikament 1
      Index           =   11
      TabIndex        =   16
      BackColor       =   -2147483643
      Height          =   256
      Left            =   1019
      Top             =   3344
      Width           =   3633
   End
   Begin VB.Label DiabetesMedikament_1_Bezeichnungsfeld
      Caption         =   "Diab&.Med. 1"
      Index           =   11
      BackColor       =   65535
      Height          =   256
      Left            =   56
      Top             =   3344
      Width           =   963
   End
   Begin VB.TextBox DiabetesMedikamente_1_Menge
      DataField = DiabetesMedikament 1 Menge
      Index           =   12
      TabIndex        =   17
      BackColor       =   -2147483643
      Height          =   256
      Left            =   5385
      Top             =   3346
      Width           =   2490
   End
   Begin VB.Label DiabetesMedikamente_1_Menge_Bezeichnungsfeld
      Caption         =   "Menge"
      Index           =   12
      BackColor       =   -2147483633
      Height          =   256
      Left            =   4762
      Top             =   3346
      Width           =   615
   End
   Begin VB.TextBox DiabetesMedikament_2
      DataField = DiabetesMedikament 2
      Index           =   13
      TabIndex        =   18
      BackColor       =   -2147483643
      Height          =   256
      Left            =   737
      Top             =   3614
      Width           =   3915
   End
   Begin VB.Label DiabetesMedikament_2_Bezeichnungsfeld
      Caption         =   "2"
      Index           =   13
      BackColor       =   -2147483633
      Height          =   256
      Left            =   56
      Top             =   3614
      Width           =   630
   End
   Begin VB.TextBox DiabetesMedikament_2_Menge
      DataField = DiabetesMedikament 2 Menge
      Index           =   14
      TabIndex        =   19
      BackColor       =   -2147483643
      Height          =   256
      Left            =   5385
      Top             =   3616
      Width           =   2478
   End
   Begin VB.Label DiabetesMedikament_2_Menge_Bezeichnungsfeld
      Caption         =   "Menge"
      Index           =   14
      BackColor       =   -2147483633
      Height          =   256
      Left            =   4705
      Top             =   3616
      Width           =   663
   End
   Begin VB.TextBox DiabetesMedikament_3
      DataField = DiabetesMedikament 3
      Index           =   15
      TabIndex        =   20
      BackColor       =   -2147483643
      Height          =   256
      Left            =   737
      Top             =   3883
      Width           =   3915
   End
   Begin VB.Label DiabetesMedikament_3_Bezeichnungsfeld
      Caption         =   "3"
      Index           =   15
      BackColor       =   -2147483633
      Height          =   256
      Left            =   56
      Top             =   3883
      Width           =   630
   End
   Begin VB.TextBox DiabetesMedikament_3_Menge
      DataField = DiabetesMedikament 3 Menge
      Index           =   16
      TabIndex        =   21
      BackColor       =   -2147483643
      Height          =   256
      Left            =   5385
      Top             =   3885
      Width           =   2490
   End
   Begin VB.Label DiabetesMedikament_3_Menge_Bezeichnungsfeld
      Caption         =   "Menge"
      Index           =   16
      BackColor       =   -2147483633
      Height          =   256
      Left            =   4762
      Top             =   3885
      Width           =   630
   End
   Begin VB.TextBox DiabetesMedikament_4
      DataField = DiabetesMedikament 4
      Index           =   17
      TabIndex        =   22
      BackColor       =   -2147483643
      Height          =   256
      Left            =   737
      Top             =   4181
      Width           =   3018
   End
   Begin VB.Label DiabetesMedikament_4_Bezeichnungsfeld
      Caption         =   "4"
      Index           =   17
      BackColor       =   -2147483633
      Height          =   256
      Left            =   56
      Top             =   4181
      Width           =   633
   End
   Begin VB.TextBox DiabetesMedikament_4_Menge
      DataField = DiabetesMedikament 4 Menge
      Index           =   18
      TabIndex        =   23
      BackColor       =   -2147483643
      Height          =   256
      Left            =   5385
      Top             =   4183
      Width           =   2490
   End
   Begin VB.Label DiabetesMedikament_4_Menge_Bezeichnungsfeld
      Caption         =   "Menge"
      Index           =   18
      BackColor       =   -2147483633
      Height          =   256
      Left            =   4648
      Top             =   4183
      Width           =   735
   End
   Begin VB.CheckBox Insulinpumpe
      Caption         =   "Insulinpumpe"
      TabIndex        =   24
      Index           =   19
      Height          =   165
      Left            =   1077
      Top             =   4592
      Width           =   180
   End
   Begin VB.Label Insulinpumpe_Bezeichnungsfeld
      Caption         =   "Insulinpumpe"
      Index           =   20
      BackColor       =   -2147483633
      Height          =   240
      Left            =   56
      Top             =   4535
      Width           =   1005
   End
   Begin VB.TextBox Insulinpumpe_seit
      DataField = Insulinpumpe seit
      Index           =   21
      TabIndex        =   25
      BackColor       =   -2147483643
      Height          =   285
      Left            =   1700
      Top             =   4535
      Width           =   1290
   End
   Begin VB.Label Insulinpumpe_seit_Bezeichnungsfeld
      Caption         =   "seit"
      Index           =   21
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1303
      Top             =   4535
      Width           =   345
   End
   Begin VB.TextBox Insulinpumpe_Marke
      DataField = Insulinpumpe Marke
      Index           =   22
      TabIndex        =   26
      BackColor       =   -2147483643
      Height          =   285
      Left            =   3685
      Top             =   4479
      Width           =   4188
   End
   Begin VB.Label Insulinpumpe_Marke_Bezeichnungsfeld
      Caption         =   "Marke"
      Index           =   22
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3061
      Top             =   4479
      Width           =   573
   End
   Begin VB.TextBox Broteinheiten_gesamt
      DataField = Broteinheiten gesamt
      Index           =   23
      TabIndex        =   27
      BackColor       =   -2147483643
      Height          =   244
      Left            =   907
      Top             =   4876
      Width           =   585
   End
   Begin VB.Label Broteinheiten_gesamt_Bezeichnungsfeld
      Caption         =   "BE &gesamt"
      Index           =   23
      BackColor       =   65535
      Height          =   244
      Left            =   0
      Top             =   4876
      Width           =   855
   End
   Begin VB.TextBox Broteinheiten_frueh
      DataField = Broteinheiten früh
      Index           =   24
      TabIndex        =   28
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1927
      Top             =   4875
      Width           =   795
   End
   Begin VB.Label Broteinheiten_frueh_Bezeichnungsfeld
      Caption         =   "früh"
      Index           =   24
      BackColor       =   -2147483633
      Height          =   244
      Left            =   1587
      Top             =   4875
      Width           =   345
   End
   Begin VB.TextBox Broteinheiten_ZM_frueh
      DataField = Broteinheiten ZM früh
      Index           =   25
      TabIndex        =   29
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3117
      Top             =   4875
      Width           =   513
   End
   Begin VB.Label Broteinheiten_ZM_frueh_Bezeichnungsfeld
      Caption         =   "vormittags"
      Index           =   25
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2777
      Top             =   4875
      Width           =   348
   End
   Begin VB.TextBox Broteinheiten_mittags
      DataField = Broteinheiten mittags
      Index           =   26
      TabIndex        =   30
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3912
      Top             =   4875
      Width           =   750
   End
   Begin VB.Label Broteinheiten_mittags_Bezeichnungsfeld
      Caption         =   "mittags"
      Index           =   26
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3685
      Top             =   4875
      Width           =   225
   End
   Begin VB.TextBox Broteinheiten_nachmittags
      DataField = Broteinheiten nachmittags
      Index           =   27
      TabIndex        =   31
      BackColor       =   -2147483643
      Height          =   244
      Left            =   4989
      Top             =   4875
      Width           =   630
   End
   Begin VB.Label Broteinheiten_nachmittags_Bezeichnungsfeld
      Caption         =   "nachmittags"
      Index           =   27
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4705
      Top             =   4875
      Width           =   285
   End
   Begin VB.TextBox Broteinheiten_abends
      DataField = Broteinheiten abends
      Index           =   28
      TabIndex        =   32
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6009
      Top             =   4875
      Width           =   858
   End
   Begin VB.Label Broteinheiten_abends_Bezeichnungsfeld
      Caption         =   "abends"
      Index           =   28
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5669
      Top             =   4875
      Width           =   348
   End
   Begin VB.TextBox Broteinheiten_nachts
      DataField = Broteinheiten nachts
      Index           =   29
      TabIndex        =   33
      BackColor       =   -2147483643
      Height          =   244
      Left            =   7257
      Top             =   4762
      Width           =   615
   End
   Begin VB.Label Broteinheiten_nachts_Bezeichnungsfeld
      Caption         =   "spät"
      Index           =   29
      BackColor       =   -2147483633
      Height          =   244
      Left            =   6916
      Top             =   4762
      Width           =   345
   End
   Begin VB.TextBox Essenszeit__frueh
      DataField = Essenszeit früh
      Index           =   30
      TabIndex        =   34
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1927
      Top             =   5160
      Width           =   795
   End
   Begin VB.Label Essenszeit__frueh_Bezeichnungsfeld
      Caption         =   "Essenszeit  fr&üh:"
      Index           =   30
      BackColor       =   65535
      Height          =   244
      Left            =   632
      Top             =   5160
      Width           =   1245
   End
   Begin VB.TextBox Essenszeit_mittags
      DataField = Essenszeit mittags
      Index           =   31
      TabIndex        =   36
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3968
      Top             =   5160
      Width           =   690
   End
   Begin VB.Label Essenszeit_mittags_Bezeichnungsfeld
      Caption         =   "mittags"
      Index           =   31
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3685
      Top             =   5160
      Width           =   285
   End
   Begin VB.TextBox Essenszeit_abends
      DataField = Essenszeit abends
      Index           =   32
      TabIndex        =   38
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6009
      Top             =   5160
      Width           =   870
   End
   Begin VB.Label Essenszeit_abends_Bezeichnungsfeld
      Caption         =   "abends"
      Index           =   32
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5669
      Top             =   5160
      Width           =   345
   End
   Begin VB.TextBox Essenszeit_spät
      DataField = Essenszeit spät
      Index           =   33
      TabIndex        =   39
      BackColor       =   -2147483643
      Height          =   244
      Left            =   7256
      Top             =   5047
      Width           =   630
   End
   Begin VB.Label Essenszeit_spät_Bezeichnungsfeld
      Caption         =   "spät"
      Index           =   33
      BackColor       =   -2147483633
      Height          =   244
      Left            =   6916
      Top             =   5047
      Width           =   345
   End
   Begin VB.TextBox Spritz_Ess_Abstand_frueh
      DataField = Spritz-Eß-Abstand früh
      Index           =   34
      TabIndex        =   40
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1927
      Top             =   5443
      Width           =   783
   End
   Begin VB.Label Spritz_Ess_Abstand_frueh_Bezeichnungsfeld
      Caption         =   "Spritz-Eß-Abstand früh:"
      Index           =   34
      BackColor       =   -2147483633
      Height          =   244
      Left            =   56
      Top             =   5442
      Width           =   1818
   End
   Begin VB.TextBox Spritz_Ess_Abstand_mittags
      DataField = Spritz-Eß-Abstand mittags
      Index           =   35
      TabIndex        =   41
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3968
      Top             =   5443
      Width           =   720
   End
   Begin VB.Label Spritz_Ess_Abstand_mittags_Bezeichnungsfeld
      Caption         =   "mittags"
      Index           =   35
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3685
      Top             =   5443
      Width           =   225
   End
   Begin VB.TextBox Spritz_Ess_Abstand_abends
      DataField = Spritz-Eß-Abstand abends
      Index           =   36
      TabIndex        =   42
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6066
      Top             =   5443
      Width           =   810
   End
   Begin VB.Label Spritz_Ess_Abstand_abends_Bezeichnungsfeld
      Caption         =   "abends"
      Index           =   36
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5669
      Top             =   5443
      Width           =   330
   End
   Begin VB.TextBox Spritzstelle_frueh_B_O_A
      DataField = Spritzstelle früh
      Index           =   37
      TabIndex        =   43
      BackColor       =   12632256
      Height          =   244
      Left            =   1927
      Top             =   5727
      Width           =   795
   End
   Begin VB.Label Spritzstelle_frueh_B_O_A_Bezeichnungsfeld
      Caption         =   "Spritzstelle früh (B, O, A)"
      Index           =   37
      BackColor       =   -2147483633
      Height          =   244
      Left            =   56
      Top             =   5727
      Width           =   1785
   End
   Begin VB.TextBox Spritzstelle_mittags
      DataField = Spritzstelle mittags
      Index           =   38
      TabIndex        =   44
      BackColor       =   12632256
      Height          =   244
      Left            =   3968
      Top             =   5727
      Width           =   723
   End
   Begin VB.Label Spritzstelle_mittags_Bezeichnungsfeld
      Caption         =   "mittags"
      Index           =   38
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3685
      Top             =   5727
      Width           =   288
   End
   Begin VB.TextBox Spritzstelle_abends
      DataField = Spritzstelle abends
      Index           =   39
      TabIndex        =   45
      BackColor       =   12632256
      Height          =   244
      Left            =   6066
      Top             =   5727
      Width           =   795
   End
   Begin VB.Label Spritzstelle_abends_Bezeichnungsfeld
      Caption         =   "abends"
      Index           =   39
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5669
      Top             =   5727
      Width           =   345
   End
   Begin VB.TextBox Spritzstelle_nachts
      DataField = Spritzstelle nachts
      Index           =   40
      TabIndex        =   46
      BackColor       =   12632256
      Height          =   244
      Left            =   7256
      Top             =   5614
      Width           =   705
   End
   Begin VB.Label Spritzstelle_nachts_Bezeichnungsfeld
      Caption         =   "nachts"
      Index           =   40
      BackColor       =   -2147483633
      Height          =   244
      Left            =   6859
      Top             =   5614
      Width           =   405
   End
   Begin VB.TextBox Jahr_letzte_Diabetesschulung
      DataField = Jahr letzte Diabetesschulung
      Index           =   41
      TabIndex        =   47
      BackColor       =   -2147483643
      Height          =   285
      Left            =   2211
      Top             =   6122
      Width           =   2115
   End
   Begin VB.Label Jahr_letzte_Diabetesschulung_Bezeichnungsfeld
      Caption         =   "&Jahr letzte Diabetesschulung"
      Index           =   41
      BackColor       =   65535
      Height          =   240
      Left            =   56
      Top             =   6122
      Width           =   2115
   End
   Begin VB.TextBox Ort_Schulung
      DataField = Ort Schulung
      Index           =   42
      TabIndex        =   48
      BackColor       =   -2147483643
      Height          =   285
      Left            =   5102
      Top             =   6122
      Width           =   2778
   End
   Begin VB.Label Ort_Schulung_Bezeichnungsfeld
      Caption         =   "Ort Schulung"
      Index           =   42
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4365
      Top             =   6122
      Width           =   738
   End
   Begin VB.TextBox letztes_HbA1c
      DataField = letztes HbA1c
      Index           =   43
      TabIndex        =   49
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1077
      Top             =   6405
      Width           =   1065
   End
   Begin VB.Label letztes_HbA1c_Bezeichnungsfeld
      Caption         =   "letztes HbA1c"
      Index           =   43
      BackColor       =   -2147483633
      Height          =   244
      Left            =   56
      Top             =   6405
      Width           =   960
   End
   Begin VB.TextBox gemessen_am
      DataField = gemessen am
      Index           =   44
      TabIndex        =   50
      BackColor       =   -2147483643
      Height          =   244
      Left            =   2769
      Top             =   6405
      Width           =   1545
   End
   Begin VB.Label gemessen_am_Bezeichnungsfeld
      Caption         =   "gem am"
      Index           =   44
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2097
      Top             =   6405
      Width           =   615
   End
   Begin VB.TextBox vorherige_Werte
      DataField = vorherige Werte
      Index           =   45
      TabIndex        =   51
      BackColor       =   -2147483643
      Height          =   244
      Left            =   5555
      Top             =   6405
      Width           =   2310
   End
   Begin VB.Label vorherige_Werte_Bezeichnungsfeld
      Caption         =   "vorherige Werte"
      Index           =   45
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4308
      Top             =   6405
      Width           =   1170
   End
   Begin VB.TextBox BZMessungen_pW
      DataField = BZMessungen pW
      Index           =   46
      TabIndex        =   55
      BackColor       =   -2147483643
      Height          =   244
      Left            =   850
      Top             =   7003
      Width           =   1365
   End
   Begin VB.Label BZMessungen_pW_Bezeichnungsfeld
      Caption         =   "Zahl pW"
      Index           =   46
      BackColor       =   -2147483633
      Height          =   244
      Left            =   113
      Top             =   7003
      Width           =   675
   End
   Begin VB.TextBox BZMessungen_pW_ndE
      DataField = BZMessungen pW ndE
      Index           =   47
      TabIndex        =   56
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3174
      Top             =   7003
      Width           =   858
   End
   Begin VB.Label BZMessungen_pW_ndE_Bezeichnungsfeld
      Caption         =   "davon ndE"
      Index           =   47
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2324
      Top             =   7003
      Width           =   843
   End
   Begin VB.TextBox BZMessungen_p_W_nachts
      DataField = BZMessungen p W nachts
      Index           =   48
      TabIndex        =   57
      BackColor       =   -2147483643
      Height          =   244
      Left            =   5272
      Top             =   7003
      Width           =   915
   End
   Begin VB.Label BZMessungen_p_W_nachts_Bezeichnungsfeld
      Caption         =   "davon nachts"
      Index           =   48
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4195
      Top             =   7003
      Width           =   1020
   End
   Begin VB.TextBox Aufschreiben
      DataField = Aufschreiben
      Index           =   49
      TabIndex        =   54
      BackColor       =   16777215
      Height          =   244
      Left            =   6236
      Top             =   6720
      Width           =   570
   End
   Begin VB.Label Aufschreiben_Bezeichnungsfeld
      Caption         =   "Aufschreiben"
      Index           =   49
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5159
      Top             =   6720
      Width           =   1020
   End
   Begin VB.TextBox BZWerte_v_d_Essen
      DataField = BZWerte v d Essen
      Index           =   50
      TabIndex        =   58
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1190
      Top             =   7371
      Width           =   1590
   End
   Begin VB.Label BZWerte_v_d_Essen_Bezeichnungsfeld
      Caption         =   "BZ &v d Essen"
      Index           =   50
      BackColor       =   65535
      Height          =   244
      Left            =   113
      Top             =   7371
      Width           =   1020
   End
   Begin VB.TextBox BZWerte_n_d_Essen
      DataField = BZWerte n d Essen
      Index           =   51
      TabIndex        =   59
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3514
      Top             =   7370
      Width           =   1488
   End
   Begin VB.Label BZWerte_n_d_Essen_Bezeichnungsfeld
      Caption         =   "danach"
      Index           =   51
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2834
      Top             =   7371
      Width           =   633
   End
   Begin VB.TextBox Unterzucker_pM
      DataField = Unterzucker pM
      Index           =   52
      TabIndex        =   61
      BackColor       =   -2147483643
      Height          =   244
      Left            =   680
      Top             =   7653
      Width           =   510
   End
   Begin VB.Label Unterzucker_pM_Bezeichnungsfeld
      Caption         =   "UZ pM"
      Index           =   52
      BackColor       =   -2147483633
      Height          =   244
      Left            =   113
      Top             =   7653
      Width           =   525
   End
   Begin VB.TextBox UZ_rechtzeitig
      DataField = UZ rechtzeitig
      Index           =   53
      TabIndex        =   62
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1984
      Top             =   7653
      Width           =   330
   End
   Begin VB.Label UZ_rechtzeitig_Bezeichnungsfeld
      Caption         =   "rechtzeitig"
      Index           =   53
      BackColor       =   -2147483633
      Height          =   244
      Left            =   1246
      Top             =   7653
      Width           =   675
   End
   Begin VB.TextBox Fremde_Hilfe_pa
      DataField = Fremde Hilfe pa
      Index           =   54
      TabIndex        =   63
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3004
      Top             =   7653
      Width           =   570
   End
   Begin VB.Label Fremde_Hilfe_pa_Bezeichnungsfeld
      Caption         =   "Fremde Hilfe pa"
      Index           =   54
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2381
      Top             =   7653
      Width           =   570
   End
   Begin VB.TextBox Bewusstlos_pa
      DataField = Bewußtlos pa
      Index           =   55
      TabIndex        =   64
      BackColor       =   -2147483643
      Height          =   244
      Left            =   4308
      Top             =   7653
      Width           =   450
   End
   Begin VB.Label Bewusstlos_pa_Bezeichnungsfeld
      Caption         =   "Bewußtl pa"
      Index           =   55
      BackColor       =   -2147483633
      Height          =   244
      Left            =   3628
      Top             =   7653
      Width           =   630
   End
   Begin VB.TextBox Keto_pa
      DataField = Keto pa
      Index           =   56
      TabIndex        =   65
      BackColor       =   -2147483643
      Height          =   244
      Left            =   5442
      Top             =   7653
      Width           =   453
   End
   Begin VB.Label Keto_pa_Bezeichnungsfeld
      Caption         =   "Keto pa"
      Index           =   56
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4762
      Top             =   7653
      Width           =   633
   End
   Begin VB.TextBox BZgr300_pM
      DataField = BZgr300 pM
      Index           =   57
      TabIndex        =   66
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6746
      Top             =   7653
      Width           =   960
   End
   Begin VB.Label BZgr300_pM_Bezeichnungsfeld
      Caption         =   "gr300 pM"
      Index           =   57
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5952
      Top             =   7653
      Width           =   735
   End
   Begin VB.TextBox Bluthochdruck
      DataField = Bluthochdruck
      Index           =   58
      TabIndex        =   69
      BackColor       =   -2147483643
      Height          =   285
      Left            =   56
      Top             =   8687
      Width           =   735
   End
   Begin VB.Label Bluthochdruck_Bezeichnungsfeld
      Caption         =   "Bl&uthochdruck"
      Index           =   58
      BackColor       =   65535
      Height          =   240
      Left            =   56
      Top             =   8447
      Width           =   735
   End
   Begin VB.TextBox BHD_seit
      DataField = BHD seit
      Index           =   59
      TabIndex        =   70
      BackColor       =   16777215
      Height          =   285
      Left            =   850
      Top             =   8687
      Width           =   1125
   End
   Begin VB.Label BHD_seit_Bezeichnungsfeld
      Caption         =   "BHD seit"
      Index           =   59
      BackColor       =   -2147483633
      Height          =   240
      Left            =   850
      Top             =   8447
      Width           =   1125
   End
   Begin VB.TextBox BHD_beh_mit
      DataField = BHD beh mit
      Index           =   60
      TabIndex        =   71
      BackColor       =   16777215
      Height          =   285
      Left            =   2035
      Top             =   8687
      Width           =   3000
   End
   Begin VB.Label BHD_beh_mit_Bezeichnungsfeld
      Caption         =   "BHD beh mit"
      Index           =   60
      BackColor       =   -2147483633
      Height          =   240
      Left            =   2035
      Top             =   8447
      Width           =   3000
   End
   Begin VB.TextBox Blutdruckwerte
      DataField = Blutdruckwerte
      Index           =   61
      TabIndex        =   73
      BackColor       =   -2147483643
      Height          =   285
      Left            =   6039
      Top             =   8687
      Width           =   1848
   End
   Begin VB.Label Blutdruckwerte_Bezeichnungsfeld
      Caption         =   "Blutdruckwerte"
      Index           =   61
      BackColor       =   -2147483633
      Height          =   240
      Left            =   6039
      Top             =   8447
      Width           =   1848
   End
   Begin VB.TextBox BDselbst
      DataField = BDselbst
      Index           =   62
      TabIndex        =   72
      BackColor       =   -2147483643
      Height          =   285
      Left            =   5102
      Top             =   8687
      Width           =   900
   End
   Begin VB.Label BDselbst_Bezeichnungsfeld
      Caption         =   "BDselbst"
      Index           =   62
      BackColor       =   -2147483633
      Height          =   240
      Left            =   5102
      Top             =   8447
      Width           =   900
   End
   Begin VB.TextBox Schwanger
      DataField = Schwanger
      Index           =   63
      TabIndex        =   67
      BackColor       =   12632256
      Height          =   244
      Left            =   1077
      Top             =   8050
      Width           =   630
   End
   Begin VB.Label Schwanger_Bezeichnungsfeld
      Caption         =   "Schwanger"
      Index           =   63
      BackColor       =   -2147483633
      Height          =   244
      Left            =   113
      Top             =   8050
      Width           =   900
   End
   Begin VB.TextBox Schwanger_seit
      DataField = Schwanger seit
      Index           =   64
      TabIndex        =   68
      BackColor       =   12632256
      Height          =   244
      Left            =   3004
      Top             =   8050
      Width           =   1185
   End
   Begin VB.Label Schwanger_seit_Bezeichnungsfeld
      Caption         =   "Schwanger seit"
      Index           =   64
      BackColor       =   -2147483633
      Height          =   244
      Left            =   1757
      Top             =   8050
      Width           =   1185
   End
   Begin VB.TextBox Augensp_zuletzt
      DataField = Augensp zuletzt
      Index           =   65
      TabIndex        =   74
      BackColor       =   -2147483643
      Height          =   285
      Left            =   56
      Top             =   9310
      Width           =   1140
   End
   Begin VB.Label Augensp_zuletzt_Bezeichnungsfeld
      Caption         =   "Augens&p zuletzt"
      Index           =   65
      BackColor       =   65535
      Height          =   240
      Left            =   60
      Top             =   9074
      Width           =   1140
   End
   Begin VB.TextBox Augensp_Befund
      DataField = Augensp Befund
      Index           =   66
      TabIndex        =   75
      BackColor       =   -2147483643
      Height          =   285
      Left            =   1256
      Top             =   9310
      Width           =   3453
   End
   Begin VB.Label Augensp_Befund_Bezeichnungsfeld
      Caption         =   "Augensp Befund"
      Index           =   66
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1256
      Top             =   9070
      Width           =   3453
   End
   Begin VB.TextBox Netzhaut_gelasert
      DataField = Netzhaut gelasert
      Index           =   67
      TabIndex        =   76
      BackColor       =   -2147483643
      Height          =   285
      Left            =   4762
      Top             =   9310
      Width           =   1260
   End
   Begin VB.Label Netzhaut_gelasert_Bezeichnungsfeld
      Caption         =   "Netzhaut gelasert"
      Index           =   67
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4762
      Top             =   9070
      Width           =   1260
   End
   Begin VB.TextBox Sehminderung_unbehebbar
      DataField = Sehminderung unbehebbar
      Index           =   68
      TabIndex        =   77
      BackColor       =   -2147483643
      Height          =   285
      Left            =   6055
      Top             =   9310
      Width           =   1800
   End
   Begin VB.Label Sehminderung_unbehebbar_Bezeichnungsfeld
      Caption         =   "Sehminderung unbehebbar"
      Index           =   68
      BackColor       =   -2147483633
      Height          =   240
      Left            =   6055
      Top             =   9070
      Width           =   1800
   End
   Begin VB.TextBox Diabet_Nierenschaden
      DataField = Diabet Nierenschaden
      Index           =   69
      TabIndex        =   78
      BackColor       =   -2147483643
      Height          =   285
      Left            =   56
      Top             =   9933
      Width           =   900
   End
   Begin VB.Label Diabet_Nierenschaden_Bezeichnungsfeld
      Caption         =   "Diabet Nierenschaden"
      Index           =   69
      BackColor       =   -2147483633
      Height          =   240
      Left            =   56
      Top             =   9693
      Width           =   900
   End
   Begin VB.TextBox Albumin_zuletzt
      DataField = Albumin zuletzt
      Index           =   70
      TabIndex        =   79
      BackColor       =   -2147483643
      Height          =   285
      Left            =   1020
      Top             =   9933
      Width           =   1125
   End
   Begin VB.Label Albumin_zuletzt_Bezeichnungsfeld
      Caption         =   "Albumin zuletzt"
      Index           =   70
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1020
      Top             =   9693
      Width           =   1125
   End
   Begin VB.TextBox erhoeht
      DataField = erhöht?
      Index           =   71
      TabIndex        =   80
      BackColor       =   -2147483643
      Height          =   285
      Left            =   2205
      Top             =   9933
      Width           =   1185
   End
   Begin VB.Label erhoeht_Bezeichnungsfeld
      Caption         =   "erhöht?"
      Index           =   71
      BackColor       =   -2147483633
      Height          =   240
      Left            =   2205
      Top             =   9693
      Width           =   1185
   End
   Begin VB.CheckBox Dialyse
      Caption         =   "Dialyse"
      TabIndex        =   81
      Index           =   72
      Height          =   285
      Left            =   3390
      Top             =   9933
      Width           =   288
   End
   Begin VB.Label Dialyse_Bezeichnungsfeld
      Caption         =   "Dialyse"
      Index           =   73
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3390
      Top             =   9693
      Width           =   288
   End
   Begin VB.TextBox Dialyse_seit
      DataField = Dialyse seit
      Index           =   74
      TabIndex        =   82
      BackColor       =   -2147483643
      Height          =   285
      Left            =   3664
      Top             =   9933
      Width           =   1155
   End
   Begin VB.Label Dialyse_seit_Bezeichnungsfeld
      Caption         =   "Dialyse seit"
      Index           =   74
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3664
      Top             =   9693
      Width           =   1155
   End
   Begin VB.TextBox andere_Nierenerkrankung
      DataField = andere Nierenerkrankung
      Index           =   75
      TabIndex        =   83
      BackColor       =   -2147483643
      Height          =   285
      Left            =   4822
      Top             =   9933
      Width           =   3030
   End
   Begin VB.Label andere_Nierenerkrankung_Bezeichnungsfeld
      Caption         =   "andere Nierenerkrankung"
      Index           =   75
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4818
      Top             =   9689
      Width           =   3030
   End
   Begin VB.TextBox Herzkrankheit
      DataField = Herzkrankheit
      Index           =   76
      TabIndex        =   84
      BackColor       =   -2147483643
      Height          =   285
      Left            =   56
      Top             =   10500
      Width           =   795
   End
   Begin VB.Label Herzkrankheit_Bezeichnungsfeld
      Caption         =   "&Herzkrankheit"
      Index           =   76
      BackColor       =   65535
      Height          =   240
      Left            =   60
      Top             =   10259
      Width           =   795
   End
   Begin VB.TextBox Angina_pectoris
      DataField = Angina pectoris
      Index           =   77
      TabIndex        =   85
      BackColor       =   16777215
      Height          =   285
      Left            =   907
      Top             =   10500
      Width           =   795
   End
   Begin VB.Label Angina_pectoris_Bezeichnungsfeld
      Caption         =   "Angina pectoris"
      Index           =   77
      BackColor       =   -2147483633
      Height          =   240
      Left            =   907
      Top             =   10260
      Width           =   855
   End
   Begin VB.TextBox Herzinfarkt
      DataField = Herzinfarkt
      Index           =   78
      TabIndex        =   86
      BackColor       =   16777215
      Height          =   285
      Left            =   1757
      Top             =   10500
      Width           =   513
   End
   Begin VB.Label Herzinfakt_Bezeichnungsfeld
      Caption         =   "Herzinfakt"
      Index           =   78
      BackColor       =   -2147483633
      Height          =   240
      Left            =   1757
      Top             =   10260
      Width           =   573
   End
   Begin VB.TextBox Herzinfarkt_wann
      DataField = Herzinfarkt wann
      Index           =   79
      TabIndex        =   87
      BackColor       =   16777215
      Height          =   285
      Left            =   2324
      Top             =   10500
      Width           =   675
   End
   Begin VB.Label Herzinfarkt_wann_Bezeichnungsfeld
      Caption         =   "wann"
      Index           =   79
      BackColor       =   -2147483633
      Height          =   240
      Left            =   2324
      Top             =   10260
      Width           =   675
   End
   Begin VB.TextBox PTCA_oder_Stent
      DataField = PTCA oder Stent
      Index           =   80
      TabIndex        =   88
      BackColor       =   16777215
      Height          =   285
      Left            =   3061
      Top             =   10500
      Width           =   510
   End
   Begin VB.Label PTCA_oder_Stent_Bezeichnungsfeld
      Caption         =   "PTCA oder Stent"
      Index           =   80
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3061
      Top             =   10260
      Width           =   510
   End
   Begin VB.CheckBox Bypass_kardial
      Caption         =   "Bypass kardial"
      TabIndex        =   89
      Index           =   81
      Height          =   285
      Left            =   3571
      Top             =   10500
      Width           =   1095
   End
   Begin VB.Label Bypass_kardial_Bezeichnungsfeld
      Caption         =   "Bypass kardial"
      Index           =   82
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3571
      Top             =   10260
      Width           =   480
   End
   Begin VB.TextBox Bypass_wann
      DataField = Bypass wann
      Index           =   83
      TabIndex        =   90
      BackColor       =   16777215
      Height          =   285
      Left            =   3798
      Top             =   10500
      Width           =   795
   End
   Begin VB.Label Bypass_wann_Bezeichnungsfeld
      Caption         =   "wann"
      Index           =   83
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4081
      Top             =   10261
      Width           =   525
   End
   Begin VB.TextBox Herzschwäche
      DataField = Herzschwäche
      Index           =   84
      TabIndex        =   91
      BackColor       =   16777215
      Height          =   285
      Left            =   4648
      Top             =   10487
      Width           =   963
   End
   Begin VB.Label Herzschwäche_Bezeichnungsfeld
      Caption         =   "Herzschwäche"
      Index           =   84
      BackColor       =   -2147483633
      Height          =   240
      Left            =   4648
      Top             =   10260
      Width           =   963
   End
   Begin VB.TextBox Herzkrankheit_Beschreibung
      DataField = Herzkrankheit Beschreibung
      Index           =   85
      TabIndex        =   92
      BackColor       =   16777215
      Height          =   285
      Left            =   5585
      Top             =   10500
      Width           =   2265
   End
   Begin VB.Label Herzkrankheit_Beschreibung_Bezeichnungsfeld
      Caption         =   "Herzkrankheit Beschreibung"
      Index           =   85
      BackColor       =   -2147483633
      Height          =   240
      Left            =   5585
      Top             =   10260
      Width           =   2265
   End
   Begin VB.TextBox Hirndurchblutungsstoerung
      DataField = Hirndurchblutungsstörung
      Index           =   86
      TabIndex        =   93
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7880
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Hirndurchblutungsstoerung_Bezeichnungsfeld
      Caption         =   "H&irndurchblutungsstörung"
      Index           =   86
      BackColor       =   65535
      Height          =   240
      Left            =   7880
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Schlaganfall
      DataField = Schlaganfall
      Index           =   87
      TabIndex        =   94
      BackColor       =   -2147483643
      Height          =   285
      Left            =   8900
      Top             =   240
      Width           =   2310
   End
   Begin VB.Label Schlaganfall_Bezeichnungsfeld
      Caption         =   "Schlaganfall"
      Index           =   87
      BackColor       =   -2147483633
      Height          =   240
      Left            =   8900
      Top             =   0
      Width           =   2310
   End
   Begin VB.TextBox Beindurchblutungsstoerung
      DataField = Beindurchblutungsstörung
      Index           =   88
      TabIndex        =   95
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11225
      Top             =   240
      Width           =   858
   End
   Begin VB.Label Beindurchblutungsstoerung_Bezeichnungsfeld
      Caption         =   "B&eindurchblutungsstörung"
      Index           =   88
      BackColor       =   65535
      Height          =   240
      Left            =   11225
      Top             =   0
      Width           =   858
   End
   Begin VB.TextBox Schaufensterkrankheit
      DataField = Schaufensterkrankheit
      Index           =   89
      TabIndex        =   96
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12132
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Schaufensterkrankheit_Bezeichnungsfeld
      Caption         =   "Schaufensterkrankheit"
      Index           =   89
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12132
      Top             =   0
      Width           =   900
   End
   Begin VB.CheckBox Bypass_peripher
      Caption         =   "Bypaß peripher"
      TabIndex        =   97
      Index           =   90
      Height          =   285
      Left            =   13096
      Top             =   240
      Width           =   270
   End
   Begin VB.Label Bypass_peripher_Bezeichnungsfeld
      Caption         =   "Bypaß peripher"
      Index           =   91
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13096
      Top             =   0
      Width           =   225
   End
   Begin VB.TextBox Geschwuer
      DataField = Geschwür
      Index           =   92
      TabIndex        =   98
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13322
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Geschwuer_Bezeichnungsfeld
      Caption         =   "Geschwür"
      Index           =   92
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13322
      Top             =   0
      Width           =   900
   End
   Begin VB.TextBox Amputation
      DataField = Amputation
      Index           =   93
      TabIndex        =   99
      BackColor       =   -2147483643
      Height          =   285
      Left            =   14222
      Top             =   240
      Width           =   2433
   End
   Begin VB.Label Amputation_Bezeichnungsfeld
      Caption         =   "Amputation"
      Index           =   93
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14222
      Top             =   0
      Width           =   2433
   End
   Begin VB.TextBox pAVK_Beschreibung
      DataField = pAVK Beschreibung
      Index           =   94
      TabIndex        =   100
      BackColor       =   -2147483643
      Height          =   645
      Left            =   7880
      Top             =   750
      Width           =   7638
   End
   Begin VB.Label pAVK_Beschreibung_Bezeichnungsfeld
      Caption         =   "pAVK Beschreibung"
      Index           =   94
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7880
      Top             =   510
      Width           =   7638
   End
   Begin VB.TextBox Ameisenlaufen
      DataField = Ameisenlaufen
      Index           =   95
      TabIndex        =   101
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7880
      Top             =   1635
      Width           =   465
   End
   Begin VB.Label Ameisenlaufen_Bezeichnungsfeld
      Caption         =   "Ameisenlaufe&n"
      Index           =   95
      BackColor       =   65535
      Height          =   240
      Left            =   7880
      Top             =   1395
      Width           =   1140
   End
   Begin VB.TextBox Ameisen_Ausmass
      DataField = Ameisen Ausmaß
      Index           =   96
      TabIndex        =   102
      BackColor       =   -2147483643
      Height          =   285
      Left            =   8390
      Top             =   1657
      Width           =   2100
   End
   Begin VB.Label Ameisen_Ausmass_Bezeichnungsfeld
      Caption         =   "Ausmaß"
      Index           =   96
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9080
      Top             =   1417
      Width           =   675
   End
   Begin VB.TextBox Druckstellen
      DataField = Druckstellen
      Index           =   97
      TabIndex        =   103
      BackColor       =   -2147483643
      Height          =   285
      Left            =   10530
      Top             =   1657
      Width           =   525
   End
   Begin VB.Label Druckstellen_Bezeichnungsfeld
      Caption         =   "Druckstellen"
      Index           =   97
      BackColor       =   -2147483633
      Height          =   240
      Left            =   10091
      Top             =   1417
      Width           =   960
   End
   Begin VB.TextBox Verformungen
      DataField = Verformungen
      Index           =   98
      TabIndex        =   104
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11111
      Top             =   1657
      Width           =   510
   End
   Begin VB.Label Verformungen_Bezeichnungsfeld
      Caption         =   "Verformungen"
      Index           =   98
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11055
      Top             =   1417
      Width           =   1020
   End
   Begin VB.TextBox Verformungen_Beschreibung
      DataField = Verformungen Beschreibung
      Index           =   99
      TabIndex        =   105
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11678
      Top             =   1657
      Width           =   1413
   End
   Begin VB.Label Verformungen_Beschreibung_Bezeichnungsfeld
      Caption         =   "Beschreibung"
      Index           =   99
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12132
      Top             =   1417
      Width           =   828
   End
   Begin VB.TextBox Neue_Fusskomplikationen
      DataField = Neue Fußkomplikationen
      Index           =   100
      TabIndex        =   109
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7937
      Top             =   2281
      Width           =   2265
   End
   Begin VB.Label Neue_Fusskomplikationen_Bezeichnungsfeld
      Caption         =   "Neue Fußkomplikationen"
      Index           =   100
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7937
      Top             =   2041
      Width           =   2265
   End
   Begin VB.TextBox Entleerungsstoerungen_Magen
      DataField = Entleerungsstörungen Magen
      Index           =   101
      TabIndex        =   110
      BackColor       =   -2147483643
      Height          =   285
      Left            =   10266
      Top             =   2281
      Width           =   1590
   End
   Begin VB.Label Entleerungsstoerungen_Magen_Bezeichnungsfeld
      Caption         =   "Entleerungsst Magen"
      Index           =   101
      BackColor       =   -2147483633
      Height          =   240
      Left            =   10261
      Top             =   2044
      Width           =   1575
   End
   Begin VB.TextBox Entleerungsstoerungen_Harnblase
      DataField = Entleerungsstörungen Harnblase
      Index           =   102
      TabIndex        =   111
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11853
      Top             =   2281
      Width           =   1155
   End
   Begin VB.Label Entleerungsstoerungen_Harnblase_Bezeichnungsfeld
      Caption         =   "Harnblase"
      Index           =   102
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11853
      Top             =   2041
      Width           =   1155
   End
   Begin VB.TextBox Schwindel_Aufstehen
      DataField = Schwindel Aufstehen
      Index           =   103
      TabIndex        =   112
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13044
      Top             =   2281
      Width           =   1248
   End
   Begin VB.Label Schwindel_Aufstehen_Bezeichnungsfeld
      Caption         =   "Schwindel Aufstehen"
      Index           =   103
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13044
      Top             =   2041
      Width           =   1248
   End
   Begin VB.TextBox Folgeerkrankungen_Haut
      DataField = Folgeerkrankungen Haut
      Index           =   104
      TabIndex        =   113
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7937
      Top             =   2848
      Width           =   1830
   End
   Begin VB.Label Folgeerkrankungen_Haut_Bezeichnungsfeld
      Caption         =   "Folgeerkrankungen Haut"
      Index           =   104
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7937
      Top             =   2608
      Width           =   1830
   End
   Begin VB.TextBox Bewegungseinschränkungen
      DataField = Bewegungseinschränkungen
      Index           =   105
      TabIndex        =   114
      BackColor       =   9803263
      Height          =   285
      Left            =   9808
      Top             =   2848
      Width           =   2310
   End
   Begin VB.Label Bewegungseinschränkungen_Bezeichnungsfeld
      Caption         =   "Bewegungseinschränkungen"
      Index           =   105
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9808
      Top             =   2608
      Width           =   2310
   End
   Begin VB.TextBox Sexualstoerung
      DataField = Sexualstörung
      Index           =   106
      TabIndex        =   115
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12188
      Top             =   2848
      Width           =   1080
   End
   Begin VB.Label Sexualstoerung_Bezeichnungsfeld
      Caption         =   "Sexualstörung"
      Index           =   106
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12188
      Top             =   2608
      Width           =   1080
   End
   Begin VB.TextBox Sexualstoerung_seit
      DataField = Sexualstörung seit
      Index           =   107
      TabIndex        =   116
      BackColor       =   16777215
      Height          =   285
      Left            =   13268
      Top             =   2848
      Width           =   1938
   End
   Begin VB.Label Sexualstoerung_seit_Bezeichnungsfeld
      Caption         =   "Sexualstörung seit"
      Index           =   107
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13268
      Top             =   2608
      Width           =   1938
   End
   Begin VB.TextBox Weitere_Anamnese
      DataField = Weitere Anamnese
      Index           =   108
      TabIndex        =   117
      BackColor       =   -2147483643
      Height          =   645
      Left            =   7937
      Top             =   3415
      Width           =   7638
   End
   Begin VB.Label Weitere_Anamnese_Bezeichnungsfeld
      Caption         =   "&Weitere Anamnese"
      Index           =   108
      BackColor       =   65535
      Height          =   240
      Left            =   7937
      Top             =   3175
      Width           =   7638
   End
   Begin VB.TextBox Alkohol
      DataField = Alkohol
      Index           =   109
      TabIndex        =   119
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11055
      Top             =   4082
      Width           =   4485
   End
   Begin VB.Label Alkohol_Bezeichnungsfeld
      Caption         =   "Alk&ohol"
      Index           =   109
      BackColor       =   65535
      Height          =   240
      Left            =   10318
      Top             =   4082
      Width           =   690
   End
   Begin VB.TextBox Tabak
      DataField = Tabak
      Index           =   110
      TabIndex        =   118
      BackColor       =   -2147483643
      Height          =   285
      Left            =   8617
      Top             =   4082
      Width           =   1698
   End
   Begin VB.Label Tabak_Bezeichnungsfeld
      Caption         =   "Tabak"
      Index           =   110
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7937
      Top             =   4082
      Width           =   633
   End
   Begin VB.TextBox Weitere_Medikation
      DataField = Weitere Medikation
      Index           =   111
      TabIndex        =   120
      BackColor       =   -2147483643
      Height          =   645
      Left            =   7937
      Top             =   4605
      Width           =   7638
   End
   Begin VB.Label Weitere_Medikation_Bezeichnungsfeld
      Caption         =   "Weitere Medikation"
      Index           =   111
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7937
      Top             =   4365
      Width           =   1653
   End
   Begin VB.TextBox Liphypertrophien_Abdomen
      DataField = Liphypertrophien Abdomen
      Index           =   112
      TabIndex        =   121
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7937
      Top             =   5512
      Width           =   2310
   End
   Begin VB.Label Liphypertrophien_Abdomen_Bezeichnungsfeld
      Caption         =   "Liph&ypertrophien Abdomen"
      Index           =   112
      BackColor       =   65535
      Height          =   240
      Left            =   7937
      Top             =   5272
      Width           =   2310
   End
   Begin VB.TextBox Liphypertrophien_Beine
      DataField = Liphypertrophien Beine
      Index           =   113
      TabIndex        =   122
      BackColor       =   -2147483643
      Height          =   285
      Left            =   10247
      Top             =   5512
      Width           =   2310
   End
   Begin VB.Label Liphypertrophien_Beine_Bezeichnungsfeld
      Caption         =   "Liphypertrophien Beine"
      Index           =   113
      BackColor       =   -2147483633
      Height          =   240
      Left            =   10247
      Top             =   5272
      Width           =   2310
   End
   Begin VB.TextBox Liphypertrophien_Arme
      DataField = Liphypertrophien Arme
      Index           =   114
      TabIndex        =   123
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12557
      Top             =   5512
      Width           =   3018
   End
   Begin VB.Label Liphypertrophien_Arme_Bezeichnungsfeld
      Caption         =   "Liphypertrophien Arme"
      Index           =   114
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12557
      Top             =   5272
      Width           =   3018
   End
   Begin VB.TextBox Hyperkeratosen
      DataField = Hyperkeratosen
      Index           =   115
      TabIndex        =   125
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7993
      Top             =   6419
      Width           =   3285
   End
   Begin VB.Label Hyperkeratosen_Bezeichnungsfeld
      Caption         =   "Hyper&keratosen"
      Index           =   115
      BackColor       =   65535
      Height          =   240
      Left            =   7993
      Top             =   6179
      Width           =   3285
   End
   Begin VB.TextBox Ulcera
      DataField = Ulcera
      Index           =   116
      TabIndex        =   126
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11338
      Top             =   6419
      Width           =   1485
   End
   Begin VB.Label Ulcera_Bezeichnungsfeld
      Caption         =   "Ulcera"
      Index           =   116
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11338
      Top             =   6179
      Width           =   1485
   End
   Begin VB.TextBox Kraft_Zehenheber
      DataField = Kraft Zehenheber
      Index           =   117
      TabIndex        =   127
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12869
      Top             =   6419
      Width           =   915
   End
   Begin VB.Label Kraft_Zehenheber_Bezeichnungsfeld
      Caption         =   "Kraft Zehenheber"
      Index           =   117
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12869
      Top             =   6179
      Width           =   915
   End
   Begin VB.TextBox Kraft_Zehenbeuger
      DataField = Kraft Zehenbeuger
      Index           =   118
      TabIndex        =   128
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13833
      Top             =   6419
      Width           =   843
   End
   Begin VB.Label Kraft_Zehenbeuger_Bezeichnungsfeld
      Caption         =   "Kraft Zehenbeuger"
      Index           =   118
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13833
      Top             =   6179
      Width           =   843
   End
   Begin VB.TextBox Kraft_Knie
      DataField = Kraft Knie
      Index           =   119
      TabIndex        =   129
      BackColor       =   -2147483643
      Height          =   285
      Left            =   14740
      Top             =   6419
      Width           =   855
   End
   Begin VB.Label Kraft_Knie_Bezeichnungsfeld
      Caption         =   "Kraft Knie"
      Index           =   119
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14740
      Top             =   6179
      Width           =   855
   End
   Begin VB.TextBox ASR
      DataField = ASR
      Index           =   120
      TabIndex        =   130
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7993
      Top             =   6986
      Width           =   1080
   End
   Begin VB.Label ASR_Bezeichnungsfeld
      Caption         =   "ASR"
      Index           =   120
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   6746
      Width           =   1080
   End
   Begin VB.TextBox PSR
      DataField = PSR
      Index           =   121
      TabIndex        =   131
      BackColor       =   -2147483643
      Height          =   285
      Left            =   9127
      Top             =   6986
      Width           =   990
   End
   Begin VB.Label PSR_Bezeichnungsfeld
      Caption         =   "PSR"
      Index           =   121
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9127
      Top             =   6746
      Width           =   990
   End
   Begin VB.TextBox Oberflächensensibilität
      DataField = Oberflächensensibilität
      Index           =   122
      TabIndex        =   132
      BackColor       =   -2147483643
      Height          =   285
      Left            =   10148
      Top             =   6986
      Width           =   1638
   End
   Begin VB.Label Oberflächensensibilität_Bezeichnungsfeld
      Caption         =   "Oberflächensensibilität"
      Index           =   122
      BackColor       =   -2147483633
      Height          =   240
      Left            =   10148
      Top             =   6746
      Width           =   1638
   End
   Begin VB.TextBox Monofilamenttest
      DataField = Monofilamenttest
      Index           =   123
      TabIndex        =   133
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11848
      Top             =   6986
      Width           =   1560
   End
   Begin VB.Label Monofilamenttest_Bezeichnungsfeld
      Caption         =   "Monofilamenttest"
      Index           =   123
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11848
      Top             =   6746
      Width           =   1560
   End
   Begin VB.TextBox Kalt_Warm
      DataField = Kalt-Warm
      Index           =   124
      TabIndex        =   134
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13408
      Top             =   6986
      Width           =   1560
   End
   Begin VB.Label Kalt_Warm_Bezeichnungsfeld
      Caption         =   "Kalt-Warm"
      Index           =   124
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13408
      Top             =   6746
      Width           =   1560
   End
   Begin VB.TextBox Vibration_IK
      DataField = Vibration IK
      Index           =   125
      TabIndex        =   135
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7993
      Top             =   7553
      Width           =   1560
   End
   Begin VB.Label Vibration_IK_Bezeichnungsfeld
      Caption         =   "Vibration IK"
      Index           =   125
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   7313
      Width           =   1560
   End
   Begin VB.TextBox Vibration_Grosszehe
      DataField = Vibration Großzehe
      Index           =   126
      TabIndex        =   136
      BackColor       =   -2147483643
      Height          =   285
      Left            =   9581
      Top             =   7553
      Width           =   1560
   End
   Begin VB.Label Vibration_Grosszehe_Bezeichnungsfeld
      Caption         =   "Vibration Großzehe"
      Index           =   126
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9581
      Top             =   7313
      Width           =   1560
   End
   Begin VB.TextBox Puls_Leiste
      DataField = Puls Leiste
      Index           =   127
      TabIndex        =   137
      BackColor       =   -2147483643
      Height          =   285
      Left            =   11225
      Top             =   7553
      Width           =   1023
   End
   Begin VB.Label Puls_Leiste_Bezeichnungsfeld
      Caption         =   "Puls Leiste"
      Index           =   127
      BackColor       =   -2147483633
      Height          =   240
      Left            =   11225
      Top             =   7313
      Width           =   1023
   End
   Begin VB.TextBox Puls_Kniekehle
      DataField = Puls Kniekehle
      Index           =   128
      TabIndex        =   138
      BackColor       =   -2147483643
      Height          =   285
      Left            =   12302
      Top             =   7553
      Width           =   900
   End
   Begin VB.Label Puls_Kniekehle_Bezeichnungsfeld
      Caption         =   "Puls Kniekehle"
      Index           =   128
      BackColor       =   -2147483633
      Height          =   240
      Left            =   12302
      Top             =   7313
      Width           =   900
   End
   Begin VB.TextBox Puls_Atp
      DataField = Puls Atp
      Index           =   129
      TabIndex        =   139
      BackColor       =   -2147483643
      Height          =   285
      Left            =   13266
      Top             =   7553
      Width           =   1080
   End
   Begin VB.Label Puls_Atp_Bezeichnungsfeld
      Caption         =   "Puls Atp"
      Index           =   129
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13266
      Top             =   7313
      Width           =   1080
   End
   Begin VB.TextBox Puls_Adp
      DataField = Puls Adp
      Index           =   130
      TabIndex        =   140
      BackColor       =   -2147483643
      Height          =   285
      Left            =   14400
      Top             =   7553
      Width           =   1125
   End
   Begin VB.Label Puls_Adp_Bezeichnungsfeld
      Caption         =   "Puls Adp"
      Index           =   130
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14400
      Top             =   7313
      Width           =   1140
   End
   Begin VB.TextBox RR
      DataField = RR
      Index           =   131
      TabIndex        =   141
      BackColor       =   -2147483643
      Height          =   285
      Left            =   8447
      Top             =   7937
      Width           =   7128
   End
   Begin VB.Label RR_Bezeichnungsfeld
      Caption         =   "&RR"
      Index           =   131
      BackColor       =   65535
      Height          =   240
      Left            =   7993
      Top             =   7937
      Width           =   393
   End
   Begin VB.TextBox Herz
      DataField = Herz
      Index           =   132
      TabIndex        =   143
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8503
      Top             =   8505
      Width           =   2655
   End
   Begin VB.Label Herz_Bezeichnungsfeld
      Caption         =   "Her&z"
      Index           =   132
      BackColor       =   65535
      Height          =   238
      Left            =   7993
      Top             =   8505
      Width           =   450
   End
   Begin VB.TextBox Lunge
      DataField = Lunge
      Index           =   133
      TabIndex        =   145
      BackColor       =   -2147483643
      Height          =   238
      Left            =   13548
      Top             =   8505
      Width           =   2040
   End
   Begin VB.Label Lunge_Bezeichnungsfeld
      Caption         =   "Lunge"
      Index           =   133
      BackColor       =   -2147483633
      Height          =   238
      Left            =   12925
      Top             =   8505
      Width           =   555
   End
   Begin VB.TextBox Bauch
      DataField = Bauch
      Index           =   134
      TabIndex        =   146
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8616
      Top             =   8755
      Width           =   1968
   End
   Begin VB.Label Bauch_Bezeichnungsfeld
      Caption         =   "Bauch"
      Index           =   134
      BackColor       =   -2147483633
      Height          =   238
      Left            =   7993
      Top             =   8755
      Width           =   573
   End
   Begin VB.TextBox WS
      DataField = WS
      Index           =   135
      TabIndex        =   147
      BackColor       =   -2147483643
      Height          =   238
      Left            =   10998
      Top             =   8755
      Width           =   1035
   End
   Begin VB.Label WS_Bezeichnungsfeld
      Caption         =   "WS"
      Index           =   135
      BackColor       =   -2147483633
      Height          =   238
      Left            =   10658
      Top             =   8755
      Width           =   285
   End
   Begin VB.TextBox NL
      DataField = NL
      Index           =   136
      TabIndex        =   148
      BackColor       =   -2147483643
      Height          =   238
      Left            =   12302
      Top             =   8755
      Width           =   1080
   End
   Begin VB.Label NL_Bezeichnungsfeld
      Caption         =   "NL"
      Index           =   136
      BackColor       =   -2147483633
      Height          =   238
      Left            =   12076
      Top             =   8755
      Width           =   240
   End
   Begin VB.TextBox SD
      DataField = SD
      Index           =   137
      TabIndex        =   149
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8220
      Top             =   8993
      Width           =   1200
   End
   Begin VB.Label SD_Bezeichnungsfeld
      Caption         =   "SD"
      Index           =   137
      BackColor       =   -2147483633
      Height          =   238
      Left            =   7993
      Top             =   8993
      Width           =   225
   End
   Begin VB.TextBox NNH
      DataField = NNH
      Index           =   138
      TabIndex        =   150
      BackColor       =   -2147483643
      Height          =   238
      Left            =   9874
      Top             =   8993
      Width           =   1290
   End
   Begin VB.Label NNH_Bezeichnungsfeld
      Caption         =   "NNH"
      Index           =   138
      BackColor       =   -2147483633
      Height          =   238
      Left            =   9467
      Top             =   8993
      Width           =   390
   End
   Begin VB.TextBox Zähne
      DataField = Zähne
      Index           =   139
      TabIndex        =   151
      BackColor       =   -2147483643
      Height          =   238
      Left            =   11792
      Top             =   8993
      Width           =   1023
   End
   Begin VB.Label Zähne_Bezeichnungsfeld
      Caption         =   "Zähne"
      Index           =   139
      BackColor       =   -2147483633
      Height          =   238
      Left            =   11225
      Top             =   8993
      Width           =   573
   End
   Begin VB.TextBox Mundhoehle
      DataField = Mundhöhle
      Index           =   140
      TabIndex        =   152
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8957
      Top             =   9243
      Width           =   1260
   End
   Begin VB.Label Mundhoehle_Bezeichnungsfeld
      Caption         =   "Mundhöhle"
      Index           =   140
      BackColor       =   -2147483633
      Height          =   238
      Left            =   7993
      Top             =   9243
      Width           =   915
   End
   Begin VB.TextBox LK
      DataField = LK
      Index           =   141
      TabIndex        =   154
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8390
      Top             =   9581
      Width           =   1380
   End
   Begin VB.Label LK_Bezeichnungsfeld
      Caption         =   "LK"
      Index           =   141
      BackColor       =   -2147483633
      Height          =   238
      Left            =   7993
      Top             =   9581
      Width           =   330
   End
   Begin VB.TextBox Neuro_sonst
      DataField = Neuro sonst
      Index           =   142
      TabIndex        =   156
      BackColor       =   -2147483643
      Height          =   238
      Left            =   8731
      Top             =   10035
      Width           =   2328
   End
   Begin VB.Label Neuro_sonst_Bezeichnungsfeld
      Caption         =   "Neuro sonst"
      Index           =   142
      BackColor       =   -2147483633
      Height          =   238
      Left            =   8050
      Top             =   10035
      Width           =   618
   End
   Begin VB.TextBox Weitere_Befunde
      DataField = Weitere Befunde
      Index           =   143
      TabIndex        =   155
      BackColor       =   -2147483643
      Height          =   540
      Left            =   11162
      Top             =   9751
      Width           =   4428
   End
   Begin VB.Label Weitere_Befunde_Bezeichnungsfeld
      Caption         =   "Weitere Befunde"
      Index           =   143
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9807
      Top             =   9751
      Width           =   1308
   End
   Begin VB.TextBox Schulung
      DataField = Schulung
      Index           =   144
      TabIndex        =   158
      BackColor       =   -2147483643
      Height          =   256
      Left            =   13889
      Top             =   10320
      Width           =   855
   End
   Begin VB.Label Schulung_Bezeichnungsfeld
      Caption         =   "S&chulung"
      Index           =   144
      BackColor       =   65535
      Height          =   256
      Left            =   13039
      Top             =   10320
      Width           =   840
   End
   Begin VB.TextBox DMP
      DataField = DMP
      Index           =   145
      TabIndex        =   159
      BackColor       =   -2147483643
      Height          =   285
      Left            =   7993
      Top             =   10828
      Width           =   900
   End
   Begin VB.Label DMP_Bezeichnungsfeld
      Caption         =   "DMP"
      Index           =   145
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   10601
      Width           =   900
   End
   Begin VB.CommandButton Ausgabe
      Caption         =   "Ausgabe"
      TabIndex        =   160
      Index           =   146
      Height          =   291
      Left            =   15583
      Top             =   566
      Width           =   1551
   End
   Begin VB.TextBox HANr
      DataField = HANr
      Index           =   147
      TabIndex        =   161
      BackColor       =   12632256
      Height          =   240
      Left            =   10148
      Top             =   10885
      Width           =   1176
   End
   Begin VB.Label HANr_Bezeichnungsfeld
      Caption         =   "HA"
      Index           =   147
      BackColor       =   -2147483633
      Height          =   240
      Left            =   9977
      Top             =   10601
      Width           =   285
   End
   Begin VB.TextBox Beinbefund
      DataField = Beinbefund
      Index           =   148
      TabIndex        =   124
      BackColor       =   -2147483643
      Height          =   240
      Left            =   9015
      Top             =   5839
      Width           =   6576
   End
   Begin VB.Label Beinbefund_Bezeichnungsfeld
      Caption         =   "Beinbefund"
      Index           =   148
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   5839
      Width           =   960
   End
   Begin VB.TextBox BMI
      DataField = =IIf([Größe]=0,"",[Gewicht]/[Größe]/[Größe]*IIf([Größe]>3,10000,1))
      Index           =   149
      TabIndex        =   162
      BackColor       =   10314363
      Height          =   238
      Left            =   5272
      Top             =   3005
      Width           =   636
   End
   Begin VB.Label BMI_Bezeichnungsfeld
      Caption         =   "BMI"
      Index           =   149
      BackColor       =   -2147483633
      Height          =   238
      Left            =   4762
      Top             =   3005
      Width           =   450
   End
   Begin VB.Label Bezeichnungsfeld298
      Caption         =   "kg/m²"
      Index           =   150
      BackColor       =   -2147483633
      Height          =   238
      Left            =   5953
      Top             =   3005
      Width           =   624
   End
   Begin VB.TextBox letzte_Änderung
      DataField = letzte Änderung
      Index           =   151
      TabIndex        =   163
      BackColor       =   12632256
      Height          =   227
      Left            =   17064
      Top             =   0
      Width           =   1349
   End
   Begin VB.Label letzte_Änderung_Bezeichnung
      Caption         =   "letzte Änderung"
      Index           =   151
      BackColor       =   -2147483633
      Height          =   240
      Left            =   15817
      Top             =   0
      Width           =   1185
   End
   Begin VB.TextBox Fusspflege
      DataField = Fußpflege
      Index           =   152
      TabIndex        =   106
      BackColor       =   -2147483643
      Height          =   240
      Left            =   13152
      Top             =   1701
      Width           =   696
   End
   Begin VB.Label Fusspflege_Bezeichnungsfeld
      Caption         =   "Fußplfege"
      Index           =   152
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13096
      Top             =   1417
      Width           =   795
   End
   Begin VB.TextBox Podologie
      DataField = Podologie
      Index           =   153
      TabIndex        =   107
      BackColor       =   -2147483643
      Height          =   240
      Left            =   13889
      Top             =   1701
      Width           =   696
   End
   Begin VB.Label Podologie_Bezeichnungsfeld
      Caption         =   "Podologie"
      Index           =   153
      BackColor       =   -2147483633
      Height          =   240
      Left            =   13889
      Top             =   1417
      Width           =   735
   End
   Begin VB.TextBox Einlagen
      DataField = Einlagen
      Index           =   154
      TabIndex        =   108
      BackColor       =   -2147483643
      Height          =   240
      Left            =   14626
      Top             =   1701
      Width           =   906
   End
   Begin VB.Label Einlagen_Bezeichnungsfeld
      Caption         =   "Orthop.Einlg/Schuhe diab&'ger"
      Index           =   154
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14683
      Top             =   1417
      Width           =   900
   End
   Begin VB.TextBox Diagnosen
      DataField = 
      Index           =   155
      TabIndex        =   164
      BackColor       =   12632256
      Height          =   4591
      Left            =   15647
      Top             =   850
      Width           =   13215
   End
   Begin VB.Label Dioagnosen_Bezeichnungsfeld
      Caption         =   "Diagnosen:"
      Index           =   155
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14400
      Top             =   1984
      Width           =   1125
   End
   Begin VB.TextBox HA
      DataField = 
      Index           =   156
      TabIndex        =   165
      BackColor       =   12632256
      Height          =   227
      Left            =   10374
      Top             =   10601
      Width           =   8148
   End
   Begin VB.TextBox Vorgestellt
      DataField = [Vorgestellt]
      Index           =   156
      TabIndex        =   4
      BackColor       =   12632256
      Height          =   238
      Left            =   2834
      Top             =   340
      Width           =   1011
   End
   Begin VB.Label Vorgestellt_Bezeichnungsfeld
      Caption         =   "vorgestellt"
      Index           =   156
      BackColor       =   -2147483633
      Height          =   238
      Left            =   1927
      Top             =   340
      Width           =   795
   End
   Begin VB.TextBox RRTurbomed
      DataField = RRTurboMed
      Index           =   157
      TabIndex        =   142
      BackColor       =   12632256
      Height          =   240
      Left            =   8447
      Top             =   8220
      Width           =   7131
   End
   Begin VB.Label RRTurboMed_Bezeichnungsfeld
      Caption         =   "RR TurboMed:"
      Index           =   157
      BackColor       =   -2147483633
      Height          =   240
      Left            =   7993
      Top             =   8221
      Width           =   450
   End
   Begin VB.TextBox Versicherung
      DataField = Versicherung
      Index           =   158
      TabIndex        =   157
      BackColor       =   12632256
      Height          =   240
      Left            =   9070
      Top             =   10318
      Width           =   3846
   End
   Begin VB.Label Versicherung_Bezeichnungsfeld
      Caption         =   "Versicherung:"
      Index           =   158
      BackColor       =   -2147483633
      Height          =   240
      Left            =   8050
      Top             =   10318
      Width           =   960
   End
   Begin VB.TextBox Titel
      DataField = Titel
      Index           =   159
      TabIndex        =   166
      BackColor       =   65280
      Height          =   240
      Left            =   4932
      Top             =   56
      Width           =   681
   End
   Begin VB.TextBox Anrede
      DataField = Anrede
      Index           =   159
      TabIndex        =   167
      BackColor       =   65280
      Height          =   240
      Left            =   5669
      Top             =   56
      Width           =   801
   End
   Begin VB.TextBox Versicherungsart
      DataField = Versicherungsart
      Index           =   159
      TabIndex        =   168
      BackColor       =   16711808
      Height          =   240
      Left            =   6576
      Top             =   56
      Width           =   336
   End
   Begin VB.TextBox Essenszeit_vormittags
      DataField = Essenszeit vormittags
      Index           =   159
      TabIndex        =   35
      BackColor       =   -2147483643
      Height          =   244
      Left            =   3118
      Top             =   5160
      Width           =   516
   End
   Begin VB.Label Essenszeit_vormittags_Bezeichnungsfeld
      Caption         =   "vormittags"
      Index           =   159
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2721
      Top             =   5160
      Width           =   405
   End
   Begin VB.TextBox Essenszeit_nachmittags
      DataField = Essenszeit nachmittags
      Index           =   160
      TabIndex        =   37
      BackColor       =   -2147483643
      Height          =   244
      Left            =   4988
      Top             =   5160
      Width           =   681
   End
   Begin VB.Label Essenszeit_nachmittags_Beichnungsfeld
      Caption         =   "nachmittags"
      Index           =   160
      BackColor       =   -2147483633
      Height          =   244
      Left            =   4705
      Top             =   5160
      Width           =   270
   End
   Begin VB.TextBox BZMessungen_Selbst
      DataField = BZMessungen selbst
      Index           =   161
      TabIndex        =   52
      BackColor       =   -2147483643
      Height          =   244
      Left            =   1870
      Top             =   6720
      Width           =   336
   End
   Begin VB.Label BZMessungen_selbst_Bezeichnungsfeld
      Caption         =   "BZ-&Messungen: selbst?"
      Index           =   161
      BackColor       =   65535
      Height          =   244
      Left            =   56
      Top             =   6720
      Width           =   1755
   End
   Begin VB.TextBox Gerät
      DataField = Gerät
      Index           =   162
      TabIndex        =   53
      BackColor       =   16777215
      Height          =   244
      Left            =   2890
      Top             =   6720
      Width           =   2091
   End
   Begin VB.Label Gerät_Bezeichnungsfeld
      Caption         =   "Gerät:"
      Index           =   162
      BackColor       =   -2147483633
      Height          =   244
      Left            =   2323
      Top             =   6720
      Width           =   525
   End
   Begin VB.TextBox UZ_Tageszeit
      DataField = UZ Tageszeit
      Index           =   163
      TabIndex        =   60
      BackColor       =   -2147483643
      Height          =   244
      Left            =   6236
      Top             =   7258
      Width           =   1701
   End
   Begin VB.Label UZTagesZeit_Bezeichnungsfeld
      Caption         =   "UZ Tageszeit"
      Index           =   163
      BackColor       =   -2147483633
      Height          =   244
      Left            =   5102
      Top             =   7257
      Width           =   1065
   End
   Begin VB.TextBox NachName
      DataField = Nachname
      Index           =   164
      TabIndex        =   1
      BackColor       =   65280
      Height          =   238
      Left            =   680
      Top             =   56
      Width           =   2100
   End
   Begin VB.TextBox BeinoedVen
      DataField = BeinödVen
      Index           =   164
      TabIndex        =   153
      BackColor       =   -2147483643
      Height          =   510
      Left            =   11399
      Top             =   9243
      Width           =   4191
   End
   Begin VB.Label BeioedVen_Bezeichnungsfeld
      Caption         =   "Beinödeme&, Venenkrkht:"
      Index           =   164
      BackColor       =   65535
      Height          =   238
      Left            =   10261
      Top             =   9243
      Width           =   1080
   End
   Begin VB.TextBox Carotiden
      DataField = Carotiden
      Index           =   165
      TabIndex        =   144
      BackColor       =   -2147483643
      Height          =   238
      Left            =   11632
      Top             =   8505
      Width           =   1230
   End
   Begin VB.Label Carotiden_Bezeichnungsfeld
      Caption         =   "Carotiden"
      Index           =   165
      BackColor       =   -2147483633
      Height          =   238
      Left            =   11225
      Top             =   8505
      Width           =   390
   End
   Begin VB.TextBox DMSchulz
      DataField = DMSchulz
      Index           =   166
      TabIndex        =   169
      BackColor       =   12632256
      Height          =   256
      Left            =   15817
      Top             =   10320
      Width           =   510
   End
   Begin VB.Label DMSchulz_Bezeichnungsfeld
      Caption         =   "Dm-Schlgn:"
      Index           =   166
      BackColor       =   -2147483633
      Height          =   256
      Left            =   14850
      Top             =   10320
      Width           =   900
   End
   Begin VB.TextBox RRSchulz
      DataField = RRSchulz
      Index           =   167
      TabIndex        =   170
      BackColor       =   12632256
      Height          =   256
      Left            =   17351
      Top             =   10318
      Width           =   510
   End
   Begin VB.Label RR_Schulz_Bezeichnungsfeld
      Caption         =   "RR-Schlgn:"
      Index           =   167
      BackColor       =   -2147483633
      Height          =   256
      Left            =   16384
      Top             =   10318
      Width           =   900
   End
   Begin VB.TextBox DMPhier
      DataField = DMPhier
      Index           =   168
      TabIndex        =   171
      BackColor       =   12632256
      Height          =   285
      Left            =   8957
      Top             =   10828
      Width           =   900
   End
   Begin VB.Label DMPhier_Bezeichnungsfeld
      Caption         =   "DMP hier:"
      Index           =   168
      BackColor       =   -2147483633
      Height          =   240
      Left            =   8957
      Top             =   10601
      Width           =   900
   End
   Begin VB.CheckBox Tkz
      Caption         =   "Tkz"
      TabIndex        =   172
      Index           =   169
      Height          =   225
      Left            =   7426
      Top             =   56
      Width           =   270
   End
   Begin VB.Label Tkz_Bezeichnungsfeld
      Caption         =   "Tkz"
      Index           =   170
      BackColor       =   -2147483633
      Height          =   225
      Left            =   6973
      Top             =   56
      Width           =   390
   End
   Begin VB.TextBox Ther1
      DataField = Ther1
      Index           =   171
      TabIndex        =   6
      BackColor       =   -2147483643
      Height          =   225
      Left            =   5782
      Top             =   626
      Width           =   750
   End
   Begin VB.Label Therapieart_Bezeichnungsfeld
      Caption         =   "Therapieart"
      Index           =   171
      BackColor       =   -2147483633
      Height          =   224
      Left            =   4818
      Top             =   626
      Width           =   900
   End
   Begin VB.TextBox Gesamtname
      DataField = =[Nachname]+" "+[Vorname]
      Index           =   172
      TabIndex        =   173
      BackColor       =   8421504
      Height          =   239
      Left            =   11338
      Top             =   4365
      Width           =   4202
   End
   Begin VB.TextBox MDi
      DataField = 
      Index           =   172
      TabIndex        =   174
      BackColor       =   12632256
      Height          =   4599
      Left            =   15760
      Top             =   5499
      Width           =   13094
   End
   Begin VB.TextBox HAimDMP
      DataField = 
      Index           =   172
      TabIndex        =   175
      BackColor       =   12632256
      Height          =   227
      Left            =   12812
      Top             =   10885
      Width           =   1486
   End
   Begin VB.TextBox V1
      DataField = 
      Index           =   172
      TabIndex        =   176
      BackColor       =   12632256
      Height          =   341
      Left            =   7993
      Top             =   11168
      Width           =   6353
   End
   Begin VB.CommandButton An1Aufruf
      Caption         =   "An1Aufruf"
      TabIndex        =   177
      Index           =   172
      Height          =   276
      Left            =   56
      Top             =   10885
      Width           =   2826
   End
   Begin VB.CommandButton An2Aufruf
      Caption         =   "An2Aufruf"
      TabIndex        =   178
      Index           =   173
      Height          =   291
      Left            =   56
      Top             =   11221
      Width           =   2826
   End
   Begin VB.CommandButton AnAAufruf
      Caption         =   "AnAAufruf"
      TabIndex        =   179
      Index           =   174
      Height          =   321
      Left            =   3571
      Top             =   10840
      Width           =   2826
   End
   Begin VB.CommandButton CheckAufruf
      Caption         =   "CheckAufruf"
      TabIndex        =   180
      Index           =   175
      Height          =   291
      Left            =   3571
      Top             =   11225
      Width           =   2826
   End
   Begin VB.CheckBox obAn1eing
      Caption         =   "obAn1eing"
      TabIndex        =   181
      Index           =   176
      Height          =   240
      Left            =   2891
      Top             =   10885
      Width           =   260
   End
   Begin VB.CheckBox obAn2eing
      Caption         =   "obAn2eing"
      TabIndex        =   182
      Index           =   177
      Height          =   240
      Left            =   2891
      Top             =   11245
      Width           =   260
   End
   Begin VB.CheckBox obAnAeing
      Caption         =   "obAnAeing"
      TabIndex        =   183
      Index           =   178
      Height          =   240
      Left            =   6462
      Top             =   10885
      Width           =   260
   End
   Begin VB.CheckBox obCheck
      Caption         =   "obCheck"
      TabIndex        =   184
      Index           =   179
      Height          =   240
      Left            =   6462
      Top             =   11225
      Width           =   260
   End
   Begin VB.CommandButton Labor
      Caption         =   "Labor"
      TabIndex        =   185
      Index           =   180
      Height          =   336
      Left            =   6776
      Top             =   10828
      Width           =   1071
   End
   Begin VB.CommandButton Brief
      Caption         =   "Brief"
      TabIndex        =   186
      Index           =   181
      Height          =   290
      Left            =   6803
      Top             =   11225
      Width           =   1011
   End
   Begin VB.CommandButton Dokumente
      Caption         =   "Dokumente"
      TabIndex        =   187
      Index           =   182
      Height          =   336
      Left            =   56
      Top             =   11565
      Width           =   696
   End
   Begin VB.CommandButton Briefe
      Caption         =   "Briefe"
      TabIndex        =   188
      Index           =   183
      Height          =   336
      Left            =   1190
      Top             =   11565
      Width           =   2031
   End
      Begin VB.OptionButton U1
         Caption         =   ""
      TabIndex        =   189
      Index           =   184
      Height          =   193
      Left            =   15649
      Top             =   5511
      Width           =   113
   End
      Begin VB.OptionButton U2
         Caption         =   ""
      TabIndex        =   190
      Index           =   185
      Height          =   193
      Left            =   15649
      Top             =   5707
      Width           =   113
   End
      Begin VB.OptionButton U3
         Caption         =   ""
      TabIndex        =   191
      Index           =   186
      Height          =   193
      Left            =   15649
      Top             =   5903
      Width           =   113
   End
      Begin VB.OptionButton U4
         Caption         =   ""
      TabIndex        =   192
      Index           =   187
      Height          =   193
      Left            =   15649
      Top             =   6098
      Width           =   113
   End
      Begin VB.OptionButton U5
         Caption         =   ""
      TabIndex        =   193
      Index           =   188
      Height          =   193
      Left            =   15649
      Top             =   6294
      Width           =   113
   End
      Begin VB.OptionButton U6
         Caption         =   ""
      TabIndex        =   194
      Index           =   189
      Height          =   193
      Left            =   15649
      Top             =   6490
      Width           =   113
   End
      Begin VB.OptionButton U7
         Caption         =   ""
      TabIndex        =   195
      Index           =   190
      Height          =   193
      Left            =   15649
      Top             =   6686
      Width           =   113
   End
      Begin VB.OptionButton U8
         Caption         =   ""
      TabIndex        =   196
      Index           =   191
      Height          =   193
      Left            =   15649
      Top             =   6882
      Width           =   113
   End
      Begin VB.OptionButton U9
         Caption         =   ""
      TabIndex        =   197
      Index           =   192
      Height          =   193
      Left            =   15649
      Top             =   7078
      Width           =   113
   End
      Begin VB.OptionButton U10
         Caption         =   ""
      TabIndex        =   198
      Index           =   193
      Height          =   193
      Left            =   15649
      Top             =   7274
      Width           =   113
   End
      Begin VB.OptionButton U11
         Caption         =   ""
      TabIndex        =   199
      Index           =   194
      Height          =   193
      Left            =   15649
      Top             =   7470
      Width           =   113
   End
      Begin VB.OptionButton U12
         Caption         =   ""
      TabIndex        =   200
      Index           =   195
      Height          =   193
      Left            =   15649
      Top             =   7666
      Width           =   113
   End
      Begin VB.OptionButton U13
         Caption         =   ""
      TabIndex        =   201
      Index           =   196
      Height          =   193
      Left            =   15649
      Top             =   7862
      Width           =   113
   End
      Begin VB.OptionButton U14
         Caption         =   ""
      TabIndex        =   202
      Index           =   197
      Height          =   193
      Left            =   15649
      Top             =   8058
      Width           =   113
   End
      Begin VB.OptionButton U15
         Caption         =   ""
      TabIndex        =   203
      Index           =   198
      Height          =   193
      Left            =   15649
      Top             =   8254
      Width           =   113
   End
      Begin VB.OptionButton U16
         Caption         =   ""
      TabIndex        =   204
      Index           =   199
      Height          =   193
      Left            =   15649
      Top             =   8450
      Width           =   113
   End
      Begin VB.OptionButton U17
         Caption         =   ""
      TabIndex        =   205
      Index           =   200
      Height          =   193
      Left            =   15649
      Top             =   8646
      Width           =   113
   End
      Begin VB.OptionButton U18
         Caption         =   ""
      TabIndex        =   206
      Index           =   201
      Height          =   193
      Left            =   15647
      Top             =   8842
      Width           =   113
   End
   Begin VB.CommandButton DReset
      Caption         =   "DReset"
      TabIndex        =   207
      Index           =   202
      Height          =   216
      Left            =   17971
      Top             =   10318
      Width           =   576
   End
   Begin VB.CheckBox obBZausgew
      Caption         =   "obBZausgew"
      TabIndex        =   208
      Index           =   203
      Height          =   240
      Left            =   14400
      Top             =   10830
      Width           =   260
   End
   Begin VB.Label obBZausgew_Bez
      Caption         =   "BZ-Meßger ausgew"
      Index           =   204
      BackColor       =   -2147483633
      Height          =   210
      Left            =   14630
      Top             =   10830
      Width           =   1470
   End
   Begin VB.CheckBox obOSaufgek
      Caption         =   "obOSaufgek"
      TabIndex        =   210
      Index           =   205
      Height          =   240
      Left            =   14400
      Top             =   11287
      Width           =   260
   End
   Begin VB.Label obOSaufgek_Bez
      Caption         =   "orthop. SM aufgekl."
      Index           =   206
      BackColor       =   -2147483633
      Height          =   225
      Left            =   14630
      Top             =   11282
      Width           =   1470
   End
   Begin VB.CheckBox obPodAufgek
      Caption         =   "obPodAufgek"
      TabIndex        =   211
      Index           =   207
      Height          =   240
      Left            =   14400
      Top             =   11532
      Width           =   260
   End
   Begin VB.Label obPodAufgek_Bez
      Caption         =   "üb.Podologie aufgek."
      Index           =   208
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14630
      Top             =   11532
      Width           =   1530
   End
   Begin VB.CheckBox obMBlAusgeh
      Caption         =   "obMBlAusgeh"
      TabIndex        =   209
      Index           =   209
      Height          =   240
      Left            =   14400
      Top             =   11080
      Width           =   260
   End
   Begin VB.Label obMBlAusgeh_Bez
      Caption         =   "&8: Merkbl DFS ausgehändigt"
      Index           =   210
      BackColor       =   -2147483633
      Height          =   240
      Left            =   14626
      Top             =   11080
      Width           =   1470
   End
   Begin VB.TextBox obDMPaufgekl
      DataField = obDMPaufgekl
      Index           =   211
      TabIndex        =   213
      BackColor       =   16777215
      Height          =   238
      Left            =   16157
      Top             =   11533
      Width           =   405
   End
   Begin VB.Label obDMPaufgekl_Bez
      Caption         =   " DMP-Beteiligung"
      Index           =   211
      BackColor       =   -2147483633
      Height          =   210
      Left            =   16611
      Top             =   11508
      Width           =   1140
   End
   Begin VB.TextBox obSchulaufgek
      DataField = obSchulaufgek
      Index           =   212
      TabIndex        =   212
      BackColor       =   16777215
      Height          =   238
      Left            =   16157
      Top             =   11281
      Width           =   405
   End
   Begin VB.Label Schulungsaufkl_Bez
      Caption         =   "&9: Schulungsaufkl"
      Index           =   212
      BackColor       =   -2147483633
      Height          =   210
      Left            =   16611
      Top             =   11281
      Width           =   1170
   End
      Begin VB.OptionButton U19
         Caption         =   ""
      TabIndex        =   214
      Index           =   213
      Height          =   193
      Left            =   15647
      Top             =   9038
      Width           =   113
   End
      Begin VB.OptionButton U20
         Caption         =   ""
      TabIndex        =   215
      Index           =   214
      Height          =   193
      Left            =   15647
      Top             =   9234
      Width           =   113
   End
      Begin VB.OptionButton U21
         Caption         =   ""
      TabIndex        =   216
      Index           =   215
      Height          =   193
      Left            =   15647
      Top             =   9430
      Width           =   113
   End
      Begin VB.OptionButton U22
         Caption         =   ""
      TabIndex        =   217
      Index           =   216
      Height          =   193
      Left            =   15647
      Top             =   9626
      Width           =   113
   End
      Begin VB.OptionButton U23
         Caption         =   ""
      TabIndex        =   218
      Index           =   217
      Height          =   193
      Left            =   15647
      Top             =   9822
      Width           =   113
   End
   Begin VB.TextBox TherAkt
      DataField = TherAkt
      Index           =   218
      TabIndex        =   219
      BackColor       =   16777215
      Height          =   238
      Left            =   6973
      Top             =   623
      Width           =   780
   End
   Begin VB.Label TherAkt_Bezeichnung
      Caption         =   "-&>"
      Index           =   218
      BackColor       =   -2147483633
      Height          =   227
      Left            =   6689
      Top             =   623
      Width           =   178
   End
   Begin VB.TextBox SchGr
      DataField = Versicherungsart
      Index           =   219
      TabIndex        =   220
      BackColor       =   12632256
      Height          =   226
      Left            =   7313
      Top             =   340
      Width           =   454
   End
   Begin VB.Label Bezeichnungsfeld459
      Caption         =   "SchGr"
      Index           =   219
      BackColor       =   -2147483633
      Height          =   227
      Left            =   6746
      Top             =   340
      Width           =   518
   End
   Begin VB.CheckBox obMedNetz
      Caption         =   "obMedNetz"
      TabIndex        =   221
      Index           =   220
      Height          =   240
      Left            =   7483
      Top             =   907
      Width           =   260
   End
   Begin VB.Label obMedNetz_Bezeichnung
      Caption         =   "Med.Netz:"
      Index           =   221
      BackColor       =   -2147483633
      Height          =   240
      Left            =   6633
      Top             =   907
      Width           =   795
   End
   Begin VB.CommandButton BZKurven
      Caption         =   "BZKurven"
      TabIndex        =   222
      Index           =   222
      Height          =   336
      Left            =   3344
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CommandButton DMPAusgeb
      Caption         =   "DMPAusgeb"
      TabIndex        =   223
      Index           =   223
      Height          =   290
      Left            =   17177
      Top             =   566
      Width           =   1296
   End
   Begin VB.CommandButton AugenBefunde
      Caption         =   "AugenBefunde"
      TabIndex        =   224
      Index           =   224
      Height          =   337
      Left            =   5499
      Top             =   11565
      Width           =   2031
   End
   Begin VB.CommandButton DokDown
      Caption         =   "DokDown"
      TabIndex        =   225
      Index           =   225
      Height          =   291
      Left            =   793
      Top             =   11565
      Width           =   336
   End
   Begin VB.TextBox HANr2
      DataField = 
      Index           =   226
      TabIndex        =   226
      BackColor       =   12632256
      Height          =   227
      Left            =   11735
      Top             =   10885
      Width           =   1032
   End
   Begin VB.ComboBox HA_Auswahl
      text = "
      Index           =   226
      BackColor       =   -2147483643
      Height          =   284
      Left            =   8569
      Top             =   11508
      Width           =   2599
   End
   Begin VB.ComboBox HA_Auswahl2
      text = "
      Index           =   227
      BackColor       =   -2147483643
      Height          =   286
      Left            =   11791
      Top             =   11508
      Width           =   2551
   End
   Begin VB.CheckBox HA1nicht
      Caption         =   "HA1nicht"
      TabIndex        =   229
      Index           =   228
      Height          =   240
      Left            =   8333
      Top             =   11792
      Width           =   260
   End
   Begin VB.Label HA1Nicht_Bez
      Caption         =   "-"
      Index           =   229
      BackColor       =   -2147483633
      Height          =   180
      Left            =   8047
      Top             =   11735
      Width           =   240
   End
   Begin VB.CheckBox HA2nicht
      Caption         =   "HA2nicht"
      TabIndex        =   230
      Index           =   230
      Height          =   240
      Left            =   11451
      Top             =   11735
      Width           =   260
   End
   Begin VB.Label HA2Nicht_Bez
      Caption         =   "-"
      Index           =   231
      BackColor       =   -2147483633
      Height          =   180
      Left            =   11168
      Top             =   11735
      Width           =   225
   End
   Begin VB.CommandButton HA1_Bez
      Caption         =   "HA1_Bez"
      TabIndex        =   231
      Index           =   232
      Height          =   285
      Left            =   7990
      Top             =   11508
      Width           =   570
   End
   Begin VB.CommandButton HA2_Bez
      Caption         =   "HA2_Bez"
      TabIndex        =   232
      Index           =   233
      Height          =   270
      Left            =   11168
      Top             =   11508
      Width           =   615
   End
   Begin VB.TextBox HAFax
      DataField = 
      Index           =   234
      TabIndex        =   233
      BackColor       =   12632256
      Height          =   284
      Left            =   16440
      Top             =   10885
      Width           =   1984
   End
   Begin VB.Label Bezeichnungsfeld488
      Caption         =   "&+"
      Index           =   234
      BackColor       =   -2147483633
      Height          =   285
      Left            =   16157
      Top             =   10885
      Width           =   255
   End
   Begin VB.Label Bezeichnungsfeld501
      Caption         =   "Duplex"
      Index           =   235
      BackColor       =   -2147483633
      Height          =   283
      Left            =   6292
      Top             =   1026
      Width           =   850
   End
   Begin VB.TextBox Doppler
      DataField = 
      Index           =   236
      TabIndex        =   234
      BackColor       =   12632256
      Height          =   680
      Left            =   4315
      Top             =   11962
      Width           =   4127
   End
   Begin VB.Label DopplerBez
      Caption         =   "Doppler"
      Index           =   236
      BackColor       =   -2147483633
      Height          =   240
      Left            =   3628
      Top             =   11962
      Width           =   630
   End
   Begin VB.TextBox Duplex
      DataField = 
      Index           =   237
      TabIndex        =   235
      BackColor       =   12632256
      Height          =   725
      Left            =   9139
      Top             =   11905
      Width           =   4632
   End
   Begin VB.Label DuplexBez
      Caption         =   "Duplex"
      Index           =   237
      BackColor       =   -2147483633
      Height          =   226
      Left            =   8447
      Top             =   11905
      Width           =   628
   End
   Begin VB.Label SonoBez
      Caption         =   "Sono"
      Index           =   238
      BackColor       =   -2147483633
      Height          =   227
      Left            =   13776
      Top             =   11848
      Width           =   462
   End
   Begin VB.TextBox Sono
      DataField = 
      Index           =   239
      TabIndex        =   236
      BackColor       =   12632256
      Height          =   737
      Left            =   14301
      Top             =   11905
      Width           =   4350
   End
   Begin VB.TextBox Eintraege
      DataField = 
      Index           =   239
      TabIndex        =   237
      BackColor       =   12632256
      Height          =   680
      Left            =   0
      Top             =   11962
      Width           =   3635
   End
   Begin VB.Label lebehBez
      Caption         =   "-"
      Index           =   239
      BackColor       =   -2147483633
      Height          =   226
      Left            =   3855
      Top             =   340
      Width           =   134
   End
   Begin VB.TextBox leBeh
      DataField = 
      Index           =   240
      TabIndex        =   5
      BackColor       =   12632256
      Height          =   226
      Left            =   4025
      Top             =   340
      Width           =   1190
   End
   Begin VB.ComboBox Datenquelle
      text = "
      Index           =   240
      BackColor       =   -2147483643
      Height          =   281
      Left            =   1231
      Top             =   56
      Width           =   3814
   End
   Begin VB.Label Datenquelle_Bezeichnungsfeld
      Caption         =   "Daten&quelle:"
      Index           =   241
      BackColor       =   -2147483633
      Height          =   240
      Left            =   113
      Top             =   56
      Width           =   1005
   End
   Begin VB.CommandButton Unausgefuelle
      Caption         =   "Unausgefülle"
      TabIndex        =   1
      Index           =   242
      Height          =   291
      Left            =   5159
      Top             =   56
      Width           =   1251
   End
   Begin VB.CommandButton alleDatensätze
      Caption         =   "alleDatensätze"
      TabIndex        =   3
      Index           =   243
      Height          =   291
      Left            =   8447
      Top             =   56
      Width           =   1251
   End
   Begin VB.CommandButton UnausgefNZ
      Caption         =   "UnausgefNZ"
      TabIndex        =   2
      Index           =   244
      Height          =   291
      Left            =   6463
      Top             =   56
      Width           =   1866
   End
   Begin VB.TextBox Datenbankname
      DataField = =currentdb.Name
      Index           =   245
      TabIndex        =   4
      BackColor       =   12632256
      Height          =   227
      Left            =   11225
      Top             =   0
      Width           =   7257
   End
   Begin VB.CommandButton Datenbank_Aufruf
      Caption         =   "Datenbank-Aufruf"
      TabIndex        =   5
      Index           =   245
      Height          =   291
      Left            =   11225
      Top             =   226
      Width           =   7311
   End
End
Attribute VB_Name = "Anamnesebogen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

