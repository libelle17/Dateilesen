VERSION 5.00
Begin VB.Form Medarten 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Medarten"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Medarten"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8445
   Begin VB.TextBox vTextB 
      DataField       =   "Medikament"
      Height          =   255
      Index           =   1
      Left            =   858
      TabIndex        =   2
      Top             =   114
      Width           =   2106
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
      Height          =   255
      Index           =   3
      Left            =   858
      TabIndex        =   6
      Top             =   969
      Width           =   900
   End
   Begin VB.TextBox vTextB 
      DataField       =   "Anzahl"
      Height          =   255
      Index           =   4
      Left            =   858
      TabIndex        =   8
      Top             =   1311
      Width           =   900
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Glib"
      DataField       =   "Glib"
      Height          =   240
      Index           =   1
      Left            =   858
      TabIndex        =   10
      Top             =   1653
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Metf"
      DataField       =   "Metf"
      Height          =   240
      Index           =   2
      Left            =   858
      TabIndex        =   12
      Top             =   1995
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "GlucI"
      DataField       =   "GlucI"
      Height          =   240
      Index           =   3
      Left            =   858
      TabIndex        =   14
      Top             =   2337
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "SHGlin"
      DataField       =   "SHGlin"
      Height          =   240
      Index           =   4
      Left            =   858
      TabIndex        =   16
      Top             =   2679
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Glit"
      DataField       =   "Glit"
      Height          =   240
      Index           =   5
      Left            =   858
      TabIndex        =   18
      Top             =   3021
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Ins"
      DataField       =   "Ins"
      Height          =   240
      Index           =   6
      Left            =   858
      TabIndex        =   20
      Top             =   3363
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Anal"
      DataField       =   "Anal"
      Height          =   240
      Index           =   7
      Left            =   858
      TabIndex        =   22
      Top             =   3705
      Width           =   260
   End
   Begin VB.TextBox vTextB 
      DataField       =   "InsArt"
      Height          =   255
      Index           =   5
      Left            =   3822
      TabIndex        =   24
      Top             =   114
      Width           =   225
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HMG"
      DataField       =   "HMG"
      Height          =   240
      Index           =   8
      Left            =   3822
      TabIndex        =   26
      Top             =   456
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Hypt"
      DataField       =   "Hypt"
      Height          =   240
      Index           =   9
      Left            =   3822
      TabIndex        =   28
      Top             =   798
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Thro"
      DataField       =   "Thro"
      Height          =   240
      Index           =   10
      Left            =   3822
      TabIndex        =   30
      Top             =   1140
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Antib"
      DataField       =   "Antib"
      Height          =   240
      Index           =   11
      Left            =   3822
      TabIndex        =   32
      Top             =   1482
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "and"
      DataField       =   "and"
      Height          =   240
      Index           =   12
      Left            =   3822
      TabIndex        =   34
      Top             =   1824
      Width           =   260
   End
   Begin VB.TextBox vTextB 
      DataField       =   "hinzugefügt"
      Height          =   255
      Index           =   6
      Left            =   2445
      TabIndex        =   36
      Top             =   2211
      Width           =   1524
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Tstr"
      DataField       =   "Tstr"
      Height          =   240
      Index           =   13
      Left            =   3822
      TabIndex        =   38
      Top             =   2508
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Puzu"
      DataField       =   "Puzu"
      Height          =   240
      Index           =   14
      Left            =   3822
      TabIndex        =   40
      Top             =   2850
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "VMat"
      DataField       =   "VMat"
      Height          =   240
      Index           =   15
      Left            =   3822
      TabIndex        =   42
      Top             =   3192
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "PenN"
      DataField       =   "PenN"
      Height          =   240
      Index           =   16
      Left            =   3822
      TabIndex        =   44
      Top             =   3534
      Width           =   260
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Neurp"
      DataField       =   "Neurp"
      Height          =   240
      Index           =   17
      Left            =   5019
      TabIndex        =   46
      Top             =   114
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AutNP"
      DataField       =   "AutNP"
      Height          =   240
      Index           =   18
      Left            =   5019
      TabIndex        =   48
      Top             =   456
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Fetts"
      DataField       =   "Fetts"
      Height          =   240
      Index           =   19
      Left            =   5019
      TabIndex        =   50
      Top             =   798
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Hsre"
      DataField       =   "Hsre"
      Height          =   240
      Index           =   20
      Left            =   5019
      TabIndex        =   52
      Top             =   1140
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiMyk"
      DataField       =   "AntiMyk"
      Height          =   240
      Index           =   21
      Left            =   5019
      TabIndex        =   54
      Top             =   1482
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Glauk"
      DataField       =   "Glauk"
      Height          =   240
      Index           =   22
      Left            =   5019
      TabIndex        =   56
      Top             =   1824
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "COLD"
      DataField       =   "COLD"
      Height          =   240
      Index           =   23
      Left            =   5019
      TabIndex        =   58
      Top             =   2166
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Pros"
      DataField       =   "Pros"
      Height          =   240
      Index           =   24
      Left            =   5019
      TabIndex        =   60
      Top             =   2508
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Urä"
      DataField       =   "Urä"
      Height          =   240
      Index           =   25
      Left            =   5019
      TabIndex        =   62
      Top             =   2850
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HyThy"
      DataField       =   "HyThy"
      Height          =   240
      Index           =   26
      Left            =   5019
      TabIndex        =   64
      Top             =   3192
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Ostp"
      DataField       =   "Ostp"
      Height          =   240
      Index           =   27
      Left            =   5019
      TabIndex        =   66
      Top             =   3534
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "KHK"
      DataField       =   "KHK"
      Height          =   240
      Index           =   28
      Left            =   5988
      TabIndex        =   68
      Top             =   114
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "HerzI"
      DataField       =   "HerzI"
      Height          =   240
      Index           =   29
      Left            =   5988
      TabIndex        =   70
      Top             =   456
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Stru"
      DataField       =   "Stru"
      Height          =   240
      Index           =   30
      Left            =   5988
      TabIndex        =   72
      Top             =   798
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AVK"
      DataField       =   "AVK"
      Height          =   240
      Index           =   31
      Left            =   5988
      TabIndex        =   74
      Top             =   1140
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "PanI"
      DataField       =   "PanI"
      Height          =   240
      Index           =   32
      Left            =   5988
      TabIndex        =   76
      Top             =   1482
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Vari"
      DataField       =   "Vari"
      Height          =   240
      Index           =   33
      Left            =   5988
      TabIndex        =   78
      Top             =   1824
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Östr"
      DataField       =   "Östr"
      Height          =   240
      Index           =   34
      Left            =   5988
      TabIndex        =   80
      Top             =   2166
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiDep"
      DataField       =   "AntiDep"
      Height          =   240
      Index           =   35
      Left            =   5988
      TabIndex        =   82
      Top             =   2508
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiDem"
      DataField       =   "AntiDem"
      Height          =   240
      Index           =   36
      Left            =   5988
      TabIndex        =   84
      Top             =   2850
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiEp"
      DataField       =   "AntiEp"
      Height          =   240
      Index           =   37
      Left            =   5988
      TabIndex        =   86
      Top             =   3192
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Park"
      DataField       =   "Park"
      Height          =   240
      Index           =   38
      Left            =   5988
      TabIndex        =   88
      Top             =   3534
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AntiPern"
      DataField       =   "AntiPern"
      Height          =   240
      Index           =   39
      Left            =   6957
      TabIndex        =   90
      Top             =   114
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Appet"
      DataField       =   "Appet"
      Height          =   240
      Index           =   40
      Left            =   6957
      TabIndex        =   92
      Top             =   456
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Anäm"
      DataField       =   "Anäm"
      Height          =   240
      Index           =   41
      Left            =   6957
      TabIndex        =   94
      Top             =   798
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Antiherp"
      DataField       =   "Antiherp"
      Height          =   240
      Index           =   42
      Left            =   6957
      TabIndex        =   96
      Top             =   1140
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "NSAR"
      DataField       =   "NSAR"
      Height          =   240
      Index           =   43
      Left            =   6957
      TabIndex        =   98
      Top             =   1482
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Antikoag"
      DataField       =   "Antikoag"
      Height          =   240
      Index           =   44
      Left            =   6957
      TabIndex        =   100
      Top             =   1824
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Betabl"
      DataField       =   "Betabl"
      Height          =   240
      Index           =   45
      Left            =   6957
      TabIndex        =   102
      Top             =   2166
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "ACEH"
      DataField       =   "ACEH"
      Height          =   240
      Index           =   46
      Left            =   6957
      TabIndex        =   104
      Top             =   2508
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "AT1"
      DataField       =   "AT1"
      Height          =   240
      Index           =   47
      Left            =   6957
      TabIndex        =   106
      Top             =   2850
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "CalcA"
      DataField       =   "CalcA"
      Height          =   240
      Index           =   48
      Left            =   6957
      TabIndex        =   108
      Top             =   3192
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "Diur"
      DataField       =   "Diur"
      Height          =   240
      Index           =   49
      Left            =   6957
      TabIndex        =   110
      Top             =   3534
      Width           =   111
   End
   Begin VB.CheckBox vCheckb 
      Caption         =   "falsch"
      DataField       =   "falsch"
      Height          =   240
      Index           =   50
      Left            =   7426
      TabIndex        =   112
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
      ScaleWidth      =   8445
      TabIndex        =   0
      Top             =   3960
      Width           =   8445
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Hinzufügen"
         Height          =   300
         Left            =   59
         TabIndex        =   113
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Bearbeiten"
         Height          =   300
         Left            =   1213
         TabIndex        =   114
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Löschen"
         Height          =   300
         Left            =   2367
         TabIndex        =   115
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "neu lesen"
         Height          =   300
         Left            =   3521
         TabIndex        =   116
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "S&chließen"
         Height          =   300
         Left            =   4675
         TabIndex        =   117
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Aktualisieren"
         Height          =   300
         Left            =   59
         TabIndex        =   118
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Abbrechen"
         Height          =   300
         Left            =   1213
         TabIndex        =   119
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
      ScaleWidth      =   8445
      TabIndex        =   120
      Top             =   4260
      Width           =   8445
      Begin VB.CommandButton Suchen 
         Caption         =   "&Suchen"
         Height          =   255
         Left            =   4920
         TabIndex        =   121
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox suche 
         Height          =   285
         Left            =   5640
         TabIndex        =   122
         Top             =   0
         Width           =   2655
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "Medarten.frx":0000
         Style           =   1  'Grafisch
         TabIndex        =   123
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&ü"
         Height          =   300
         Left            =   345
         Picture         =   "Medarten.frx":0342
         Style           =   1  'Grafisch
         TabIndex        =   124
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&ä"
         Height          =   300
         Left            =   4200
         Picture         =   "Medarten.frx":0684
         Style           =   1  'Grafisch
         TabIndex        =   125
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "Medarten.frx":09C6
         Style           =   1  'Grafisch
         TabIndex        =   126
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   690
         TabIndex        =   127
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label vLab 
      Caption         =   "Medikament"
      Height          =   255
      Index           =   1
      Left            =   57
      TabIndex        =   1
      Top             =   114
      Width           =   741
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
      Left            =   57
      TabIndex        =   7
      Top             =   1311
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Glib"
      Height          =   255
      Index           =   5
      Left            =   57
      TabIndex        =   9
      Top             =   1653
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Metf"
      Height          =   255
      Index           =   6
      Left            =   57
      TabIndex        =   11
      Top             =   1995
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "GlucI"
      Height          =   255
      Index           =   7
      Left            =   57
      TabIndex        =   13
      Top             =   2337
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "SHGlin"
      Height          =   255
      Index           =   8
      Left            =   57
      TabIndex        =   15
      Top             =   2679
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Glit"
      Height          =   255
      Index           =   9
      Left            =   57
      TabIndex        =   17
      Top             =   3021
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Ins"
      Height          =   255
      Index           =   10
      Left            =   57
      TabIndex        =   19
      Top             =   3363
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Anal"
      Height          =   255
      Index           =   11
      Left            =   57
      TabIndex        =   21
      Top             =   3705
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "InsArt"
      Height          =   255
      Index           =   12
      Left            =   3021
      TabIndex        =   23
      Top             =   114
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "HMG"
      Height          =   255
      Index           =   13
      Left            =   3021
      TabIndex        =   25
      Top             =   456
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Hypt"
      Height          =   255
      Index           =   14
      Left            =   3021
      TabIndex        =   27
      Top             =   798
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Thro"
      Height          =   255
      Index           =   15
      Left            =   3021
      TabIndex        =   29
      Top             =   1140
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Antib"
      Height          =   255
      Index           =   16
      Left            =   3021
      TabIndex        =   31
      Top             =   1482
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "and"
      Height          =   255
      Index           =   17
      Left            =   3021
      TabIndex        =   33
      Top             =   1824
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "hinzugefügt"
      Height          =   255
      Index           =   18
      Left            =   1644
      TabIndex        =   35
      Top             =   2211
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Tstr"
      Height          =   255
      Index           =   19
      Left            =   3021
      TabIndex        =   37
      Top             =   2508
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Puzu"
      Height          =   255
      Index           =   20
      Left            =   3021
      TabIndex        =   39
      Top             =   2850
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "VMat"
      Height          =   255
      Index           =   21
      Left            =   3021
      TabIndex        =   41
      Top             =   3192
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "PenN"
      Height          =   255
      Index           =   22
      Left            =   3021
      TabIndex        =   43
      Top             =   3534
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Neurp"
      Height          =   255
      Index           =   23
      Left            =   4218
      TabIndex        =   45
      Top             =   114
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "AutNP"
      Height          =   255
      Index           =   24
      Left            =   4218
      TabIndex        =   47
      Top             =   456
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Fetts"
      Height          =   255
      Index           =   25
      Left            =   4218
      TabIndex        =   49
      Top             =   798
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Hsre"
      Height          =   255
      Index           =   26
      Left            =   4218
      TabIndex        =   51
      Top             =   1140
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "AntiMyk"
      Height          =   255
      Index           =   27
      Left            =   4218
      TabIndex        =   53
      Top             =   1482
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Glauk"
      Height          =   255
      Index           =   28
      Left            =   4218
      TabIndex        =   55
      Top             =   1824
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "COLD"
      Height          =   255
      Index           =   29
      Left            =   4218
      TabIndex        =   57
      Top             =   2166
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Pros"
      Height          =   255
      Index           =   30
      Left            =   4218
      TabIndex        =   59
      Top             =   2508
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Urä"
      Height          =   255
      Index           =   31
      Left            =   4218
      TabIndex        =   61
      Top             =   2850
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "HyThy"
      Height          =   255
      Index           =   32
      Left            =   4218
      TabIndex        =   63
      Top             =   3192
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Ostp"
      Height          =   255
      Index           =   33
      Left            =   4218
      TabIndex        =   65
      Top             =   3534
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "KHK"
      Height          =   255
      Index           =   34
      Left            =   5187
      TabIndex        =   67
      Top             =   114
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "HerzI"
      Height          =   255
      Index           =   35
      Left            =   5187
      TabIndex        =   69
      Top             =   456
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Stru"
      Height          =   255
      Index           =   36
      Left            =   5187
      TabIndex        =   71
      Top             =   798
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "AVK"
      Height          =   255
      Index           =   37
      Left            =   5187
      TabIndex        =   73
      Top             =   1140
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "PanI"
      Height          =   255
      Index           =   38
      Left            =   5187
      TabIndex        =   75
      Top             =   1482
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Vari"
      Height          =   255
      Index           =   39
      Left            =   5187
      TabIndex        =   77
      Top             =   1824
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Östr"
      Height          =   255
      Index           =   40
      Left            =   5187
      TabIndex        =   79
      Top             =   2166
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "AntiDep"
      Height          =   255
      Index           =   41
      Left            =   5187
      TabIndex        =   81
      Top             =   2508
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "AntiDem"
      Height          =   255
      Index           =   42
      Left            =   5187
      TabIndex        =   83
      Top             =   2850
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "AntiEp"
      Height          =   255
      Index           =   43
      Left            =   5187
      TabIndex        =   85
      Top             =   3192
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Park"
      Height          =   255
      Index           =   44
      Left            =   5187
      TabIndex        =   87
      Top             =   3534
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "AntiPern"
      Height          =   255
      Index           =   45
      Left            =   6156
      TabIndex        =   89
      Top             =   114
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Appet"
      Height          =   255
      Index           =   46
      Left            =   6156
      TabIndex        =   91
      Top             =   456
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Anäm"
      Height          =   255
      Index           =   47
      Left            =   6156
      TabIndex        =   93
      Top             =   798
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Antiherp"
      Height          =   255
      Index           =   48
      Left            =   6156
      TabIndex        =   95
      Top             =   1140
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "NSAR"
      Height          =   255
      Index           =   49
      Left            =   6156
      TabIndex        =   97
      Top             =   1482
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Antikoag"
      Height          =   255
      Index           =   50
      Left            =   6156
      TabIndex        =   99
      Top             =   1824
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Betabl"
      Height          =   255
      Index           =   51
      Left            =   6156
      TabIndex        =   101
      Top             =   2166
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "ACEH"
      Height          =   255
      Index           =   52
      Left            =   6156
      TabIndex        =   103
      Top             =   2508
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "AT1"
      Height          =   255
      Index           =   53
      Left            =   6156
      TabIndex        =   105
      Top             =   2850
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "CalcA"
      Height          =   255
      Index           =   54
      Left            =   6156
      TabIndex        =   107
      Top             =   3192
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "Diur"
      Height          =   255
      Index           =   55
      Left            =   6156
      TabIndex        =   109
      Top             =   3534
      Width           =   741
   End
   Begin VB.Label vLab 
      Caption         =   "falsch"
      Height          =   255
      Index           =   56
      Left            =   7710
      TabIndex        =   111
      Top             =   56
      Width           =   741
   End
End
Attribute VB_Name = "Medarten"
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
   Call do_Form_Current_Medarten(Me)
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

Public Sub cmdAdd_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdAdd_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'cmdAdd_Click()

Public Sub cmdDelete_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdDelete_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Private Sub cmdDelete_Click()

Public Sub cmdRefresh_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdRefresh_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdEdit_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'cmdEdit_Click()

Public Sub cmdCancel_Click()
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

Public Sub cmdUpdate_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdUpdate_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdUpdate_Click()

Private Sub cmdClose_Click()
  Unload Me
End Sub

Public Sub cmdFirst_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdFirst_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdFirst_Click

Public Sub cmdLast_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdLast_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdLast_Click()

Public Sub cmdNext_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdNext_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdNext_Click()

Public Sub cmdPrevious_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdPrevious_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
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
Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
 Call doKeyDown(Me, keyCode, Shift)
End Sub

Public Sub Suchen_Click()
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Suchen_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
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

Private Sub form_unload(Cancel As Integer)
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
