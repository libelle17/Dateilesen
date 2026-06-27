VERSION 5.00
Begin VB.Form Medarten 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Medarten"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Medarten"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   10005
   Begin VB.CommandButton Googeln 
      Caption         =   "&Googeln"
      Height          =   255
      Left            =   1920
      TabIndex        =   138
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "GLP&1"
      DataField       =   "GLP1"
      Height          =   240
      Index           =   52
      Left            =   795
      TabIndex        =   24
      Top             =   3240
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SGLT&2"
      DataField       =   "SGLT2"
      Height          =   240
      Index           =   53
      Left            =   2115
      TabIndex        =   26
      Top             =   3240
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DPP&4"
      DataField       =   "DPP4"
      Height          =   240
      Index           =   51
      Left            =   2115
      TabIndex        =   22
      Top             =   2880
      Width           =   260
   End
   Begin VB.CommandButton inTM 
      Caption         =   "in MO an&z"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox vCheckb 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "&Sonst.AD"
      DataField       =   "SonstAD"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   27
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Medikament"
      Height          =   255
      Index           =   1
      Left            =   1095
      TabIndex        =   2
      Top             =   114
      Width           =   1860
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Langname"
      Height          =   450
      Index           =   2
      Left            =   858
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   4
      Top             =   456
      Width           =   2106
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Pat_ID"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   858
      TabIndex        =   6
      Top             =   969
      Width           =   780
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H8000000F&
      DataField       =   "Anzahl"
      Height          =   255
      Index           =   4
      Left            =   858
      TabIndex        =   10
      Top             =   1545
      Width           =   900
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Glib"
      DataField       =   "Glib"
      Height          =   240
      Index           =   1
      Left            =   858
      TabIndex        =   12
      Top             =   1890
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Metf"
      DataField       =   "Metf"
      Height          =   240
      Index           =   2
      Left            =   858
      TabIndex        =   14
      Top             =   2235
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "GlucI"
      DataField       =   "GlucI"
      Height          =   240
      Index           =   3
      Left            =   858
      TabIndex        =   16
      Top             =   2580
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SHGlin"
      DataField       =   "SHGlin"
      Height          =   240
      Index           =   4
      Left            =   2175
      TabIndex        =   18
      Top             =   2565
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Glit"
      DataField       =   "Glit"
      Height          =   240
      Index           =   5
      Left            =   858
      TabIndex        =   20
      Top             =   2895
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ins"
      DataField       =   "Ins"
      Height          =   240
      Index           =   6
      Left            =   858
      TabIndex        =   29
      Top             =   3960
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Anal"
      DataField       =   "Anal"
      Height          =   240
      Index           =   7
      Left            =   858
      TabIndex        =   31
      Top             =   4305
      Width           =   260
   End
   Begin VB.TextBox vTextB 
      DataField       =   "InsArt"
      Height          =   255
      Index           =   5
      Left            =   3822
      TabIndex        =   33
      Top             =   114
      Width           =   225
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HMG"
      DataField       =   "HMG"
      Height          =   240
      Index           =   8
      Left            =   3822
      TabIndex        =   35
      Top             =   456
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Hypt"
      DataField       =   "Hypt"
      Height          =   240
      Index           =   9
      Left            =   3822
      TabIndex        =   37
      Top             =   798
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Thro"
      DataField       =   "Thro"
      Height          =   240
      Index           =   10
      Left            =   3822
      TabIndex        =   39
      Top             =   1140
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Antib"
      DataField       =   "Antib"
      Height          =   240
      Index           =   11
      Left            =   3822
      TabIndex        =   41
      Top             =   1482
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "and"
      DataField       =   "and"
      Height          =   240
      Index           =   12
      Left            =   3822
      TabIndex        =   43
      Top             =   1824
      Width           =   260
   End
   Begin VB.TextBox vTextB 
      DataField       =   "hinzugefügt"
      Height          =   255
      Index           =   6
      Left            =   2445
      TabIndex        =   45
      Top             =   2211
      Width           =   1524
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Tstr"
      DataField       =   "Tstr"
      Height          =   240
      Index           =   13
      Left            =   3822
      TabIndex        =   47
      Top             =   2508
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Puzu"
      DataField       =   "Puzu"
      Height          =   240
      Index           =   14
      Left            =   3822
      TabIndex        =   49
      Top             =   2850
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "VMat"
      DataField       =   "VMat"
      Height          =   240
      Index           =   15
      Left            =   3822
      TabIndex        =   51
      Top             =   3192
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "PenN"
      DataField       =   "PenN"
      Height          =   240
      Index           =   16
      Left            =   3822
      TabIndex        =   53
      Top             =   3534
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Neurp"
      DataField       =   "Neurp"
      Height          =   240
      Index           =   17
      Left            =   5265
      TabIndex        =   55
      Top             =   114
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AutNP"
      DataField       =   "AutNP"
      Height          =   240
      Index           =   18
      Left            =   5265
      TabIndex        =   57
      Top             =   456
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Fetts"
      DataField       =   "Fetts"
      Height          =   240
      Index           =   19
      Left            =   5265
      TabIndex        =   59
      Top             =   798
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Hsre"
      DataField       =   "Hsre"
      Height          =   240
      Index           =   20
      Left            =   5265
      TabIndex        =   61
      Top             =   1140
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiMyk"
      DataField       =   "AntiMyk"
      Height          =   240
      Index           =   21
      Left            =   5265
      TabIndex        =   63
      Top             =   1482
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Glauk"
      DataField       =   "Glauk"
      Height          =   240
      Index           =   22
      Left            =   5265
      TabIndex        =   65
      Top             =   1824
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "COLD"
      DataField       =   "COLD"
      Height          =   240
      Index           =   23
      Left            =   5265
      TabIndex        =   67
      Top             =   2166
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Pros"
      DataField       =   "Pros"
      Height          =   240
      Index           =   24
      Left            =   5265
      TabIndex        =   69
      Top             =   2508
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Urä"
      DataField       =   "Urä"
      Height          =   240
      Index           =   25
      Left            =   5265
      TabIndex        =   71
      Top             =   2850
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HyThy"
      DataField       =   "HyThy"
      Height          =   240
      Index           =   26
      Left            =   5265
      TabIndex        =   73
      Top             =   3192
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Ostp"
      DataField       =   "Ostp"
      Height          =   240
      Index           =   27
      Left            =   5265
      TabIndex        =   75
      Top             =   3534
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "KHK"
      DataField       =   "KHK"
      Height          =   240
      Index           =   28
      Left            =   6825
      TabIndex        =   77
      Top             =   114
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HerzI"
      DataField       =   "HerzI"
      Height          =   240
      Index           =   29
      Left            =   6825
      TabIndex        =   79
      Top             =   456
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Stru"
      DataField       =   "Stru"
      Height          =   240
      Index           =   30
      Left            =   6825
      TabIndex        =   81
      Top             =   798
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AVK"
      DataField       =   "AVK"
      Height          =   240
      Index           =   31
      Left            =   6825
      TabIndex        =   83
      Top             =   1140
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "PanI"
      DataField       =   "PanI"
      Height          =   240
      Index           =   32
      Left            =   6825
      TabIndex        =   85
      Top             =   1482
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Vari"
      DataField       =   "Vari"
      Height          =   240
      Index           =   33
      Left            =   6825
      TabIndex        =   87
      Top             =   1824
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Östr"
      DataField       =   "Östr"
      Height          =   240
      Index           =   34
      Left            =   6825
      TabIndex        =   89
      Top             =   2166
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiDep"
      DataField       =   "AntiDep"
      Height          =   240
      Index           =   35
      Left            =   6825
      TabIndex        =   91
      Top             =   2508
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiDem"
      DataField       =   "AntiDem"
      Height          =   240
      Index           =   36
      Left            =   6825
      TabIndex        =   93
      Top             =   2850
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiEp"
      DataField       =   "AntiEp"
      Height          =   240
      Index           =   37
      Left            =   6825
      TabIndex        =   95
      Top             =   3192
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Park"
      DataField       =   "Park"
      Height          =   240
      Index           =   38
      Left            =   6825
      TabIndex        =   97
      Top             =   3534
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiPern"
      DataField       =   "AntiPern"
      Height          =   240
      Index           =   39
      Left            =   8400
      TabIndex        =   99
      Top             =   114
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Appet"
      DataField       =   "Appet"
      Height          =   240
      Index           =   40
      Left            =   8400
      TabIndex        =   101
      Top             =   456
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Anäm"
      DataField       =   "Anäm"
      Height          =   240
      Index           =   41
      Left            =   8400
      TabIndex        =   103
      Top             =   798
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Antiherp"
      DataField       =   "Antiherp"
      Height          =   240
      Index           =   42
      Left            =   8400
      TabIndex        =   105
      Top             =   1140
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "NSAR"
      DataField       =   "NSAR"
      Height          =   240
      Index           =   43
      Left            =   8400
      TabIndex        =   107
      Top             =   1482
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Antikoag"
      DataField       =   "Antikoag"
      Height          =   240
      Index           =   44
      Left            =   8400
      TabIndex        =   109
      Top             =   1824
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Betabl"
      DataField       =   "Betabl"
      Height          =   240
      Index           =   45
      Left            =   8400
      TabIndex        =   111
      Top             =   2166
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "ACEH"
      DataField       =   "ACEH"
      Height          =   240
      Index           =   46
      Left            =   8400
      TabIndex        =   113
      Top             =   2508
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AT1"
      DataField       =   "AT1"
      Height          =   240
      Index           =   47
      Left            =   8400
      TabIndex        =   115
      Top             =   2850
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "CalcA"
      DataField       =   "CalcA"
      Height          =   240
      Index           =   48
      Left            =   8400
      TabIndex        =   117
      Top             =   3192
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Diur"
      DataField       =   "Diur"
      Height          =   240
      Index           =   49
      Left            =   8400
      TabIndex        =   119
      Top             =   3534
      Width           =   211
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "falsch"
      DataField       =   "falsch"
      Height          =   240
      Index           =   50
      Left            =   8865
      TabIndex        =   121
      Top             =   113
      Width           =   260
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10005
      TabIndex        =   0
      Top             =   4920
      Width           =   10005
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Hinzufügen"
         Height          =   300
         Left            =   59
         TabIndex        =   122
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Bearbeiten"
         Height          =   300
         Left            =   1213
         TabIndex        =   123
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Löschen"
         Height          =   300
         Left            =   2367
         TabIndex        =   124
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "neu lesen"
         Height          =   300
         Left            =   3521
         TabIndex        =   125
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Schließen"
         Height          =   300
         Left            =   4675
         TabIndex        =   126
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Aktualisieren"
         Height          =   300
         Left            =   59
         TabIndex        =   127
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Abbrechen"
         Height          =   300
         Left            =   1213
         TabIndex        =   128
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
      ScaleWidth      =   10005
      TabIndex        =   129
      Top             =   5220
      Width           =   10005
      Begin VB.CommandButton Suchen 
         Caption         =   "&Suchen"
         Height          =   255
         Left            =   4920
         TabIndex        =   130
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox suche 
         Height          =   285
         Left            =   5640
         TabIndex        =   131
         Top             =   0
         Width           =   2655
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "Medarten geändert1.frx":0000
         Style           =   1  'Grafisch
         TabIndex        =   132
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&ü"
         Height          =   300
         Left            =   345
         Picture         =   "Medarten geändert1.frx":0342
         Style           =   1  'Grafisch
         TabIndex        =   133
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&ä"
         Height          =   300
         Left            =   4200
         Picture         =   "Medarten geändert1.frx":0684
         Style           =   1  'Grafisch
         TabIndex        =   134
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "Medarten geändert1.frx":09C6
         Style           =   1  'Grafisch
         TabIndex        =   135
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   690
         TabIndex        =   136
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label vLab 
      Caption         =   "GLP&1"
      DataField       =   "GLP1"
      Height          =   255
      Index           =   57
      Left            =   0
      TabIndex        =   23
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "SGLT&2"
      Height          =   255
      Index           =   58
      Left            =   1320
      TabIndex        =   25
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "DPP&4"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label PName 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Erklärung 
      Height          =   255
      Left            =   120
      TabIndex        =   137
      Top             =   4680
      Width           =   8295
   End
   Begin VB.Label vLab 
      Caption         =   "Medikament&+"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label vLab 
      Caption         =   "Langname"
      Height          =   255
      Index           =   2
      Left            =   57
      TabIndex        =   3
      Top             =   456
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Pat_ID"
      Height          =   255
      Index           =   3
      Left            =   57
      TabIndex        =   5
      Top             =   969
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Anzahl"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   9
      Top             =   1545
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Glib"
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   11
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "&Metformin"
      Height          =   255
      Index           =   6
      Left            =   60
      TabIndex        =   13
      Top             =   2235
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "GlucI"
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   15
      Top             =   2580
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "SHGlin"
      Height          =   255
      Index           =   8
      Left            =   1380
      TabIndex        =   17
      Top             =   2565
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Glit"
      Height          =   255
      Index           =   9
      Left            =   60
      TabIndex        =   19
      Top             =   2895
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "&Insuline"
      Height          =   255
      Index           =   10
      Left            =   60
      TabIndex        =   28
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "&Analoga"
      Height          =   255
      Index           =   11
      Left            =   60
      TabIndex        =   30
      Top             =   4425
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "InsArt"
      Height          =   255
      Index           =   12
      Left            =   3021
      TabIndex        =   32
      Top             =   114
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "&HMG"
      Height          =   255
      Index           =   13
      Left            =   3021
      TabIndex        =   34
      Top             =   456
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "H&yperton"
      Height          =   255
      Index           =   14
      Left            =   3021
      TabIndex        =   36
      Top             =   798
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Thro"
      Height          =   255
      Index           =   15
      Left            =   3021
      TabIndex        =   38
      Top             =   1140
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Anti&b"
      Height          =   255
      Index           =   16
      Left            =   3021
      TabIndex        =   40
      Top             =   1482
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "an&d"
      Height          =   255
      Index           =   17
      Left            =   3021
      TabIndex        =   42
      Top             =   1824
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "hinzugefügt"
      Height          =   255
      Index           =   18
      Left            =   1440
      TabIndex        =   44
      Top             =   2205
      Width           =   855
   End
   Begin VB.Label vLab 
      Caption         =   "Tstr"
      Height          =   255
      Index           =   19
      Left            =   3021
      TabIndex        =   46
      Top             =   2508
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Puzu"
      Height          =   255
      Index           =   20
      Left            =   3021
      TabIndex        =   48
      Top             =   2850
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "&VMat"
      Height          =   255
      Index           =   21
      Left            =   3021
      TabIndex        =   50
      Top             =   3192
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "PenNadeln"
      Height          =   255
      Index           =   22
      Left            =   3021
      TabIndex        =   52
      Top             =   3534
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "&Neurp"
      Height          =   255
      Index           =   23
      Left            =   4455
      TabIndex        =   54
      Top             =   120
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "AutNP"
      Height          =   255
      Index           =   24
      Left            =   4455
      TabIndex        =   56
      Top             =   450
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Fetts"
      Height          =   255
      Index           =   25
      Left            =   4455
      TabIndex        =   58
      Top             =   795
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Hsre"
      Height          =   255
      Index           =   26
      Left            =   4455
      TabIndex        =   60
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "AntiMyk"
      Height          =   255
      Index           =   27
      Left            =   4455
      TabIndex        =   62
      Top             =   1485
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Glauk"
      Height          =   255
      Index           =   28
      Left            =   4455
      TabIndex        =   64
      Top             =   1830
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "CO&LD"
      Height          =   255
      Index           =   29
      Left            =   4455
      TabIndex        =   66
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Pros"
      Height          =   255
      Index           =   30
      Left            =   4455
      TabIndex        =   68
      Top             =   2505
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "&Urämiesp"
      Height          =   255
      Index           =   31
      Left            =   4455
      TabIndex        =   70
      Top             =   2850
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "HyperThy"
      Height          =   255
      Index           =   32
      Left            =   4455
      TabIndex        =   72
      Top             =   3195
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "&Ostp"
      Height          =   255
      Index           =   33
      Left            =   4455
      TabIndex        =   74
      Top             =   3540
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "KHK"
      Height          =   255
      Index           =   34
      Left            =   6030
      TabIndex        =   76
      Top             =   120
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "HerzI"
      Height          =   255
      Index           =   35
      Left            =   6030
      TabIndex        =   78
      Top             =   450
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Struma&/Hy"
      Height          =   255
      Index           =   36
      Left            =   6030
      TabIndex        =   80
      Top             =   795
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "AVK"
      Height          =   255
      Index           =   37
      Left            =   6030
      TabIndex        =   82
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "PankrIns"
      Height          =   255
      Index           =   38
      Left            =   6030
      TabIndex        =   84
      Top             =   1485
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Varikose"
      Height          =   255
      Index           =   39
      Left            =   6030
      TabIndex        =   86
      Top             =   1830
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "&Östr"
      Height          =   255
      Index           =   40
      Left            =   6030
      TabIndex        =   88
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "AntiDe&p"
      Height          =   255
      Index           =   41
      Left            =   6030
      TabIndex        =   90
      Top             =   2505
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "AntiDem"
      Height          =   255
      Index           =   42
      Left            =   6030
      TabIndex        =   92
      Top             =   2850
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "An&tiEpil"
      Height          =   255
      Index           =   43
      Left            =   6030
      TabIndex        =   94
      Top             =   3195
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Park"
      Height          =   255
      Index           =   44
      Left            =   6030
      TabIndex        =   96
      Top             =   3540
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "AntiPern"
      Height          =   255
      Index           =   45
      Left            =   7590
      TabIndex        =   98
      Top             =   120
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Appet (&X)"
      Height          =   255
      Index           =   46
      Left            =   7590
      TabIndex        =   100
      Top             =   450
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Anäm"
      Height          =   255
      Index           =   47
      Left            =   7590
      TabIndex        =   102
      Top             =   795
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Antiherp"
      Height          =   255
      Index           =   48
      Left            =   7590
      TabIndex        =   104
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "NSA&R"
      Height          =   255
      Index           =   49
      Left            =   7590
      TabIndex        =   106
      Top             =   1485
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Anti&koag"
      Height          =   255
      Index           =   50
      Left            =   7590
      TabIndex        =   108
      Top             =   1830
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "B&etabl"
      Height          =   255
      Index           =   51
      Left            =   7590
      TabIndex        =   110
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "A&CE-H"
      Height          =   255
      Index           =   52
      Left            =   7590
      TabIndex        =   112
      Top             =   2505
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "AT&1"
      Height          =   255
      Index           =   53
      Left            =   7590
      TabIndex        =   114
      Top             =   2850
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "CalcAnta&g"
      Height          =   255
      Index           =   54
      Left            =   7590
      TabIndex        =   116
      Top             =   3195
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "Diur (&W)"
      Height          =   255
      Index           =   55
      Left            =   7590
      TabIndex        =   118
      Top             =   3540
      Width           =   735
   End
   Begin VB.Label vLab 
      Caption         =   "&falsch"
      Height          =   255
      Index           =   56
      Left            =   9150
      TabIndex        =   120
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "Medarten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents anaRS As Recordset
Attribute anaRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Public mbDataChanged As Boolean
Dim altsuche$, SuchStringGeändert%
Public obStumm As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
        lpOperation As String, ByVal lpFile As String, ByVal _
        lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Private Sub anaRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hierdurch wird die aktuelle Datensatzposition für diese Datensatzgruppe angezeigt
   lblStatus.Caption = CStr(anaRS.AbsolutePosition)
   Call do_Form_Current_Medarten(Me)
   Dim rs As New ADODB.Recordset
   If IsNumeric(anaRS!Pat_id) Then
    myFrag rs, "SELECT CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')), IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')), n.nachname,',',n.vorname) Name FROM `namen` n WHERE pat_id = " & anaRS!Pat_id
    If Not rs.BOF Then
     Me.PName = rs!name
    End If ' Not rs.BOF Then
   End If ' IsNumeric(anaRS!Pat_id) Then
   If Not anaRS.BOF And Not anaRS.EOF Then
    If anaRS!Falsch <> 0 Then
     Me.BackColor = "&H0080C0FF"
    ElseIf anaRS!glib <> 0 Or anaRS!metf <> 0 Or anaRS!gluci <> 0 Or anaRS!shglin <> 0 _
     Or anaRS!glit <> 0 Or anaRS!dpp4 Or anaRS!glp1 Or anaRS!sglt2 Or anaRS!sonstad <> 0 Or anaRS!InS <> 0 Or anaRS!anal <> 0 Or anaRS!insart <> "0" Or anaRS!hmg <> 0 Or anaRS!hypt <> 0 Or anaRS!Thro <> 0 Or anaRS!Antib <> 0 Or anaRS!And <> 0 Or anaRS!tStr <> 0 Or anaRS!puzu <> 0 Or anaRS!VMat <> 0 Or anaRS!PenN <> 0 Or anaRS!neurp <> 0 Or anaRS!autnp <> 0 Or anaRS!fetts <> 0 Or anaRS!Hsre <> 0 Or anaRS!antimyk <> 0 Or anaRS!glauk <> 0 Or anaRS!cold <> 0 Or anaRS!pros <> 0 Or anaRS!urä <> 0 Or anaRS!hythy <> 0 Or anaRS!ostp <> 0 Or anaRS!khk <> 0 Or anaRS!HerzI <> 0 Or anaRS!stru <> 0 Or anaRS!avk <> 0 Or anaRS!pani <> 0 Or anaRS!vari <> 0 Or anaRS!östr <> 0 Or anaRS!antidep <> 0 Or anaRS!antidem <> 0 Or anaRS!antiep <> 0 Or anaRS!park <> 0 Or anaRS!antipern <> 0 Or anaRS!Appet <> 0 Or anaRS!Anäm <> 0 Or anaRS!antiherp <> 0 Or anaRS!NSAR <> 0 Or anaRS!antikoag <> 0 Or anaRS!Betabl <> 0 Or anaRS!ACEH <> 0 Or anaRS!AT1 <> 0 Or anaRS!CalcA <> 0 Or anaRS!Diur <> 0 Then
     Me.BackColor = "&H0080FFFF"
    Else
     Me.BackColor = "&H8000000F" ' gelblichgrau
    End If
   End If ' Not anaRS.BOF And Not anaRS.EOF Then
End Sub 'anaRS_MoveComplete

Private Sub anaRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
End Sub 'anaRS_WillChangeRecord

Public Sub cmdAdd_Click()
  On Error GoTo fehler
  With anaRS
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
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdAdd_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'cmdAdd_Click()

Public Sub cmdDelete_Click()
  On Error GoTo fehler
  With anaRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdDelete_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Private Sub cmdDelete_Click()

Public Sub cmdRefresh_Click()
  'Dies wird nur für Mehrbenutzeranwendungen benötigt
  On Error GoTo fehler
  anaRS.Requery
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdRefresh_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Private Sub cmdRefresh_Click()"

Public Sub cmdEdit_Click()
  On Error GoTo fehler
  lblStatus.Caption = "Datensatz bearbeiten"
  mbEditFlag = True
  SetButtons False
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdEdit_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'cmdEdit_Click()

Public Sub cmdCancel_Click()
  On Error Resume Next
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  anaRS.CancelUpdate
 If mvBookMark > 0 Then
    anaRS.Bookmark = mvBookMark
  Else
    anaRS.MoveFirst
  End If
  mbDataChanged = False
End Sub ' cmdCancel_Click()

Public Sub cmdUpdate_Click()
  On Error GoTo fehler
  anaRS.UpdateBatch adAffectAll
  If mbAddNewFlag Then
    anaRS.MoveLast              'Zu neuem Datensatz gehen
  End If
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdUpdate_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdUpdate_Click()

Private Sub cmdClose_Click()
  Unload Me
End Sub ' cmdClose_Click()

Public Sub cmdFirst_Click()
  On Error GoTo fehler
  anaRS.MoveFirst
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdFirst_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdFirst_Click

Public Sub cmdLast_Click()
  On Error GoTo fehler
  anaRS.MoveLast
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdLast_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdLast_Click()

Public Sub cmdNext_Click()
  On Error GoTo fehler
  If Not anaRS.EOF Then anaRS.MoveNext
  If anaRS.EOF And anaRS.RecordCount > 0 Then
    Tüt 1760, 1000
     'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    anaRS.MoveLast
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdNext_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdNext_Click()

Public Sub cmdPrevious_Click()
  On Error GoTo fehler
  If Not anaRS.BOF Then anaRS.MovePrevious
  If anaRS.BOF And anaRS.RecordCount > 0 Then
    Tüt 1760, 1000
    'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    anaRS.MoveFirst
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdPrevious_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdPrevious_Click()

Public Sub SetButtons(bVal As Boolean)
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
End Sub ' Form_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Googeln_Click()
 Dim Result&
  Result = ShellExecute(Me.hwnd, "Open", "https://www.google.com/search?client=firefox-b-d&q=" + vTextB(1), "", App.path, 1)
End Sub ' Googeln_Click()

Private Sub inTM_Click()
  inMOAnz (vTextB(3))
End Sub ' inTM_Click()

Public Sub Suchen_Click()
 On Error GoTo fehler
 If SuchStringGeändert Then
  altsuche = Me.suche
  If IsNumeric(Me.suche) Then
   anaRS.Find " pat_id = " & Me.suche, 0, adSearchForward
  Else
   If anaRS.EOF Then anaRS.MoveFirst
   anaRS.Find "medikament LIKE '*" & Me.suche & "*'", 1, adSearchForward
  End If
 Else
  Me.suche.SetFocus
  SuchStringGeändert = True
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Suchen_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'Suchen_Click()

  Private Sub Form_Load()
  Call doForm_Load(Me)
  End Sub 'Private Sub Form_Load()

Private Sub Form_Resize()
  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Lese.Visible = True
  Screen.MousePointer = vbDefault
End Sub 'Private Sub Form_Unload()
' vLab(1)        :Medikament_Bezeichnungsfeld
' vLab(2)        :Langname_Bezeichnungsfeld
' vLab(3)        :Pat_ID_Bezeichnungsfeld
' vLab(4)        :Anzahl_Bezeichnungsfeld
' vLab(5)        :Glib_Bezeichnungsfeld
' vLab(6)        :Metf_Bezeichnungsfeld
' vLab(7)        :GlucI_Bezeichnungsfeld
' vLab(8)        :SHGlin_Bezeichnungsfeld
' vLab(9)        :Glit_Bezeichnungsfeld
' vLab(10)       :Ins_Bezeichnungsfeld
' vLab(11)       :Anal_Bezeichnungsfeld
' vLab(12)       :InsArt_Bezeichnungsfeld
' vLab(13)       :HMG_Bezeichnungsfeld
' vLab(14)       :Hypt_Bezeichnungsfeld
' vLab(15)       :Thro_Bezeichnungsfeld
' vLab(16)       :Antib_Bezeichnungsfeld
' vLab(17)       :and_Bezeichnungsfeld
' vLab(18)       :hinzugefügt_Bezeichnungsfeld
' vLab(19)       :Tstr_Bezeichnungsfeld
' vLab(20)       :Puzu_Bezeichnungsfeld
' vLab(21)       :VMat_Bezeichnungsfeld
' vLab(22)       :PenN_Bezeichnungsfeld
' vLab(23)       :Neurp_Bezeichnungsfeld
' vLab(24)       :AutNP_Bezeichnungsfeld
' vLab(25)       :Fetts_Bezeichnungsfeld
' vLab(26)       :Hsre_Bezeichnungsfeld
' vLab(27)       :AntiMyk_Bezeichnungsfeld
' vLab(28)       :Glauk_Bezeichnungsfeld
' vLab(29)       :COLD_Bezeichnungsfeld
' vLab(30)       :Pros_Bezeichnungsfeld
' vLab(31)       :Urä_Bezeichnungsfeld
' vLab(32)       :HyThy_Bezeichnungsfeld
' vLab(33)       :Ostp_Bezeichnungsfeld
' vLab(34)       :KHK_Bezeichnungsfeld
' vLab(35)       :HerzI_Bezeichnungsfeld
' vLab(36)       :Stru_Bezeichnungsfeld
' vLab(37)       :AVK_Bezeichnungsfeld
' vLab(38)       :PanI_Bezeichnungsfeld
' vLab(39)       :Vari_Bezeichnungsfeld
' vLab(40)       :Östr_Bezeichnungsfeld
' vLab(41)       :AntiDep_Bezeichnungsfeld
' vLab(42)       :AntiDem_Bezeichnungsfeld
' vLab(43)       :AntiEp_Bezeichnungsfeld
' vLab(44)       :Park_Bezeichnungsfeld
' vLab(45)       :AntiPern_Bezeichnungsfeld
' vLab(46)       :Appet_Bezeichnungsfeld
' vLab(47)       :Anäm_Bezeichnungsfeld
' vLab(48)       :Antiherp_Bezeichnungsfeld
' vLab(49)       :NSAR_Bezeichnungsfeld
' vLab(50)       :Antikoag_Bezeichnungsfeld
' vLab(51)       :Betabl_Bezeichnungsfeld
' vLab(52)       :ACEH_Bezeichnungsfeld
' vLab(53)       :AT1_Bezeichnungsfeld
' vLab(54)       :CalcA_Bezeichnungsfeld
' vLab(55)       :Diur_Bezeichnungsfeld
' vLab(56)       :falsch_Bezeichnungsfeld
' vTextB(1)      :Medikament
' vTextB(2)      :Langname
' vTextB(3)      :Pat_ID
' vTextB(4)      :Anzahl
' vTextB(5)      :InsArt
' vTextB(6)      :hinzugefügt
' vCheckb(1)     :Glib
' vCheckb(2)     :Metf
' vCheckb(3)     :GlucI
' vCheckb(4)     :SHGlin
' vCheckb(5)     :Glit
' vCheckb(6)     :Ins
' vCheckb(7)     :Anal
' vCheckb(8)     :HMG
' vCheckb(9)     :Hypt
' vCheckb(10)    :Thro
' vCheckb(11)    :Antib
' vCheckb(12)    :and
' vCheckb(13)    :Tstr
' vCheckb(14)    :Puzu
' vCheckb(15)    :VMat
' vCheckb(16)    :PenN
' vCheckb(17)    :Neurp
' vCheckb(18)    :AutNP
' vCheckb(19)    :Fetts
' vCheckb(20)    :Hsre
' vCheckb(21)    :AntiMyk
' vCheckb(22)    :Glauk
' vCheckb(23)    :COLD
' vCheckb(24)    :Pros
' vCheckb(25)    :Urä
' vCheckb(26)    :HyThy
' vCheckb(27)    :Ostp
' vCheckb(28)    :KHK
' vCheckb(29)    :HerzI
' vCheckb(30)    :Stru
' vCheckb(31)    :AVK
' vCheckb(32)    :PanI
' vCheckb(33)    :Vari
' vCheckb(34)    :Östr
' vCheckb(35)    :AntiDep
' vCheckb(36)    :AntiDem
' vCheckb(37)    :AntiEp
' vCheckb(38)    :Park
' vCheckb(39)    :AntiPern
' vCheckb(40)    :Appet
' vCheckb(41)    :Anäm
' vCheckb(42)    :Antiherp
' vCheckb(43)    :NSAR
' vCheckb(44)    :Antikoag
' vCheckb(45)    :Betabl
' vCheckb(46)    :ACEH
' vCheckb(47)    :AT1
' vCheckb(48)    :CalcA
' vCheckb(49)    :Diur
' vCheckb(50)    :falsch

'ACEH:                                   vCheckb(46)
'ACEH_Bezeichnungsfeld:                  vLab(52)
'AT1:                                    vCheckb(47)
'AT1_Bezeichnungsfeld:                   vLab(53)
'AVK:                                    vCheckb(31)
'AVK_Bezeichnungsfeld:                   vLab(37)
'Anal:                                   vCheckb(7)
'Anal_Bezeichnungsfeld:                  vLab(11)
'AntiDem:                                vCheckb(36)
'AntiDem_Bezeichnungsfeld:               vLab(42)
'AntiDep:                                vCheckb(35)
'AntiDep_Bezeichnungsfeld:               vLab(41)
'AntiEp:                                 vCheckb(37)
'AntiEp_Bezeichnungsfeld:                vLab(43)
'AntiMyk:                                vCheckb(21)
'AntiMyk_Bezeichnungsfeld:               vLab(27)
'AntiPern:                               vCheckb(39)
'AntiPern_Bezeichnungsfeld:              vLab(45)
'Antib:                                  vCheckb(11)
'Antib_Bezeichnungsfeld:                 vLab(16)
'Antiherp:                               vCheckb(42)
'Antiherp_Bezeichnungsfeld:              vLab(48)
'Antikoag:                               vCheckb(44)
'Antikoag_Bezeichnungsfeld:              vLab(50)
'Anzahl:                                 vTextB(4)
'Anzahl_Bezeichnungsfeld:                vLab(4)
'Anäm:                                   vCheckb(41)
'Anäm_Bezeichnungsfeld:                  vLab(47)
'Appet:                                  vCheckb(40)
'Appet_Bezeichnungsfeld:                 vLab(46)
'AutNP:                                  vCheckb(18)
'AutNP_Bezeichnungsfeld:                 vLab(24)
'Betabl:                                 vCheckb(45)
'Betabl_Bezeichnungsfeld:                vLab(51)
'COLD:                                   vCheckb(23)
'COLD_Bezeichnungsfeld:                  vLab(29)
'CalcA:                                  vCheckb(48)
'CalcA_Bezeichnungsfeld:                 vLab(54)
'Diur:                                   vCheckb(49)
'Diur_Bezeichnungsfeld:                  vLab(55)
'Fetts:                                  vCheckb(19)
'Fetts_Bezeichnungsfeld:                 vLab(25)
'Glauk:                                  vCheckb(22)
'Glauk_Bezeichnungsfeld:                 vLab(28)
'Glib:                                   vCheckb(1)
'Glib_Bezeichnungsfeld:                  vLab(5)
'Glit:                                   vCheckb(5)
'Glit_Bezeichnungsfeld:                  vLab(9)
'GlucI:                                  vCheckb(3)
'GlucI_Bezeichnungsfeld:                 vLab(7)
'HMG:                                    vCheckb(8)
'HMG_Bezeichnungsfeld:                   vLab(13)
'HerzI:                                  vCheckb(29)
'HerzI_Bezeichnungsfeld:                 vLab(35)
'Hsre:                                   vCheckb(20)
'Hsre_Bezeichnungsfeld:                  vLab(26)
'HyThy:                                  vCheckb(26)
'HyThy_Bezeichnungsfeld:                 vLab(32)
'Hypt:                                   vCheckb(9)
'Hypt_Bezeichnungsfeld:                  vLab(14)
'Ins:                                    vCheckb(6)
'InsArt:                                 vTextB(5)
'InsArt_Bezeichnungsfeld:                vLab(12)
'Ins_Bezeichnungsfeld:                   vLab(10)
'KHK:                                    vCheckb(28)
'KHK_Bezeichnungsfeld:                   vLab(34)
'Langname:                               vTextB(2)
'Langname_Bezeichnungsfeld:              vLab(2)
'Medikament:                             vTextB(1)
'Medikament_Bezeichnungsfeld:            vLab(1)
'Metf:                                   vCheckb(2)
'Metf_Bezeichnungsfeld:                  vLab(6)
'NSAR:                                   vCheckb(43)
'NSAR_Bezeichnungsfeld:                  vLab(49)
'Neurp:                                  vCheckb(17)
'Neurp_Bezeichnungsfeld:                 vLab(23)
'Ostp:                                   vCheckb(27)
'Ostp_Bezeichnungsfeld:                  vLab(33)
'PanI:                                   vCheckb(32)
'PanI_Bezeichnungsfeld:                  vLab(38)
'Park:                                   vCheckb(38)
'Park_Bezeichnungsfeld:                  vLab(44)
'Pat_ID:                                 vTextB(3)
'Pat_ID_Bezeichnungsfeld:                vLab(3)
'PenN:                                   vCheckb(16)
'PenN_Bezeichnungsfeld:                  vLab(22)
'Pros:                                   vCheckb(24)
'Pros_Bezeichnungsfeld:                  vLab(30)
'Puzu:                                   vCheckb(14)
'Puzu_Bezeichnungsfeld:                  vLab(20)
'SHGlin:                                 vCheckb(4)
'SHGlin_Bezeichnungsfeld:                vLab(8)
'Stru:                                   vCheckb(30)
'Stru_Bezeichnungsfeld:                  vLab(36)
'Thro:                                   vCheckb(10)
'Thro_Bezeichnungsfeld:                  vLab(15)
'Tstr:                                   vCheckb(13)
'Tstr_Bezeichnungsfeld:                  vLab(19)
'Urä:                                    vCheckb(25)
'Urä_Bezeichnungsfeld:                   vLab(31)
'VMat:                                   vCheckb(15)
'VMat_Bezeichnungsfeld:                  vLab(21)
'Vari:                                   vCheckb(33)
'Vari_Bezeichnungsfeld:                  vLab(39)
'and:                                    vCheckb(12)
'and_Bezeichnungsfeld:                   vLab(17)
'falsch:                                 vCheckb(50)
'falsch_Bezeichnungsfeld:                vLab(56)
'hinzugefügt:                            vTextB(6)
'hinzugefügt_Bezeichnungsfeld:           vLab(18)
'Östr:                                   vCheckb(34)
'Östr_Bezeichnungsfeld:                  vLab(40)

Private Sub vCheckb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim rs As New ADODB.Recordset
 If LVobMySQL Then
  myFrag rs, "SELECT column_comment FROM information_schema.`COLUMNS` WHERE table_schema = '" & CurDB(DBCn) & "' AND table_name = 'medarten' AND column_name = '" & Me.vCheckb(Index).DataField & "'"
  If Not rs.EOF Then
   Me.Erklärung = rs!column_comment
  End If
 End If ' LVobMySQL Then
End Sub ' vCheckb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x AS Single, Y AS Single)


Private Sub vTextB_GotFocus(Index As Integer)
 Me.vTextB(Index).SelStart = 0
 Me.vTextB(Index).SelLength = Len(Me.vTextB(Index).Text)
End Sub ' Sub vTextB_GotFocus(Index As Integer)

Private Sub vTextB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim rs As New ADODB.Recordset
 If LVobMySQL Then
  myFrag rs, "SELECT column_comment FROM information_schema.`COLUMNS` WHERE table_schema = '" & CurDB(DBCn) & "' AND table_name = 'medarten' AND column_name = '" & Me.vTextB(Index).DataField & "'"
  If Not rs.EOF Then
   Me.Erklärung = rs!column_comment
  End If
 End If
End Sub ' Sub vTextB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x AS Single, Y AS Single)
