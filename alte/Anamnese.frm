VERSION 5.00
Begin VB.Form Anamnese 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "anamnesebogen"
   ClientHeight    =   39990
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   17895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   39990
   ScaleWidth      =   17895
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   10440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   412
      Text            =   "Anamnese.frx":0000
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1575
      Left            =   6120
      TabIndex        =   409
      Top             =   8160
      Width           =   4095
      Begin VB.OptionButton Option7 
         Caption         =   "Option7"
         Height          =   735
         Left            =   360
         TabIndex        =   411
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Left            =   360
         TabIndex        =   410
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Option5"
      Height          =   375
      Left            =   10080
      TabIndex        =   407
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   495
      Left            =   10080
      TabIndex        =   406
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   8400
      TabIndex        =   405
      Top             =   240
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   615
      Left            =   6840
      TabIndex        =   403
      Top             =   7440
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   6840
      TabIndex        =   402
      Top             =   6960
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   10080
      TabIndex        =   401
      Top             =   4800
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   7080
      TabIndex        =   400
      Top             =   2400
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   7560
      TabIndex        =   399
      Top             =   840
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   735
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5880
      TabIndex        =   396
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5880
      TabIndex        =   395
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   6000
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   394
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   17895
      TabIndex        =   386
      Top             =   39390
      Width           =   17895
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Abbrechen"
         Height          =   300
         Left            =   1213
         TabIndex        =   393
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Aktualisieren"
         Height          =   300
         Left            =   59
         TabIndex        =   392
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "S&chließen"
         Height          =   300
         Left            =   4675
         TabIndex        =   391
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Neu lesen"
         Height          =   300
         Left            =   3521
         TabIndex        =   390
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Löschen"
         Height          =   300
         Left            =   2367
         TabIndex        =   389
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Bearbeiten"
         Height          =   300
         Left            =   1213
         TabIndex        =   388
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Hinzufügen"
         Height          =   300
         Left            =   59
         TabIndex        =   387
         Top             =   0
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
      ScaleWidth      =   17895
      TabIndex        =   380
      Top             =   39690
      Width           =   17895
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "Anamnese.frx":0006
         Style           =   1  'Grafisch
         TabIndex        =   384
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "Anamnese.frx":0348
         Style           =   1  'Grafisch
         TabIndex        =   383
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "Anamnese.frx":068A
         Style           =   1  'Grafisch
         TabIndex        =   382
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "Anamnese.frx":09CC
         Style           =   1  'Grafisch
         TabIndex        =   381
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   720
         TabIndex        =   385
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "QT"
      Height          =   285
      Index           =   189
      Left            =   2040
      TabIndex        =   379
      Top             =   -4996
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "QS"
      Height          =   285
      Index           =   188
      Left            =   2040
      TabIndex        =   377
      Top             =   -5316
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "ob"
      Height          =   285
      Index           =   187
      Left            =   2040
      TabIndex        =   375
      Top             =   -5636
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Hausarzt"
      Height          =   285
      Index           =   186
      Left            =   2040
      TabIndex        =   373
      Top             =   -5956
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obMedNetz"
      Height          =   285
      Index           =   185
      Left            =   2040
      TabIndex        =   371
      Top             =   -6276
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "obDMPaufgekl"
      Height          =   285
      Index           =   184
      Left            =   2040
      TabIndex        =   369
      Top             =   -6596
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "obSchulaufgek"
      Height          =   285
      Index           =   183
      Left            =   2040
      TabIndex        =   367
      Top             =   -6916
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obMBlAusgeh"
      Height          =   285
      Index           =   182
      Left            =   2040
      TabIndex        =   365
      Top             =   -7236
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obPodAufgek"
      Height          =   285
      Index           =   181
      Left            =   2040
      TabIndex        =   363
      Top             =   -7556
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obOSaufgek"
      Height          =   285
      Index           =   180
      Left            =   2040
      TabIndex        =   361
      Top             =   -7876
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obBZausgew"
      Height          =   285
      Index           =   179
      Left            =   2040
      TabIndex        =   359
      Top             =   -8196
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obCheck"
      Height          =   285
      Index           =   178
      Left            =   2040
      TabIndex        =   357
      Top             =   -8516
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obAnAeing"
      Height          =   285
      Index           =   177
      Left            =   2040
      TabIndex        =   355
      Top             =   -8836
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obAn2eing"
      Height          =   285
      Index           =   176
      Left            =   2040
      TabIndex        =   353
      Top             =   -9156
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "obAn1eing"
      Height          =   285
      Index           =   175
      Left            =   2040
      TabIndex        =   351
      Top             =   -9476
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Prim"
      Height          =   285
      Index           =   174
      Left            =   2040
      TabIndex        =   349
      Top             =   -9796
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TherAkt"
      Height          =   285
      Index           =   173
      Left            =   2040
      TabIndex        =   347
      Top             =   -10116
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ther1"
      Height          =   285
      Index           =   172
      Left            =   2040
      TabIndex        =   345
      Top             =   -10436
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AktZeit"
      Height          =   285
      Index           =   171
      Left            =   2040
      TabIndex        =   343
      Top             =   -10756
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Versicherung"
      Height          =   285
      Index           =   170
      Left            =   2040
      TabIndex        =   341
      Top             =   -11076
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Vorgestellt"
      Height          =   285
      Index           =   169
      Left            =   2040
      TabIndex        =   339
      Top             =   -11396
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Diagnosen"
      Height          =   285
      Index           =   168
      Left            =   2040
      TabIndex        =   337
      Top             =   -11716
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "letzte Änderung"
      Height          =   285
      Index           =   167
      Left            =   2040
      TabIndex        =   335
      Top             =   -12036
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "HANr2"
      Height          =   285
      Index           =   166
      Left            =   2040
      TabIndex        =   333
      Top             =   -12356
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "HANr"
      Height          =   285
      Index           =   165
      Left            =   2040
      TabIndex        =   331
      Top             =   -12676
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DMPhier"
      Height          =   285
      Index           =   164
      Left            =   2040
      TabIndex        =   329
      Top             =   -12996
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RRSchulz"
      Height          =   285
      Index           =   163
      Left            =   2040
      TabIndex        =   327
      Top             =   -13316
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DMSchL"
      Height          =   285
      Index           =   162
      Left            =   2040
      TabIndex        =   325
      Top             =   -13636
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DMSchulz"
      Height          =   285
      Index           =   161
      Left            =   2040
      TabIndex        =   323
      Top             =   -13956
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DMP"
      Height          =   285
      Index           =   160
      Left            =   2040
      TabIndex        =   321
      Top             =   -14276
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Schulung"
      Height          =   285
      Index           =   159
      Left            =   2040
      TabIndex        =   319
      Top             =   -14596
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Weitere Befunde"
      Height          =   285
      Index           =   158
      Left            =   2040
      TabIndex        =   317
      Top             =   -14916
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Neuro sonst"
      Height          =   285
      Index           =   157
      Left            =   2040
      TabIndex        =   315
      Top             =   -15236
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BeinödVen"
      Height          =   285
      Index           =   156
      Left            =   2040
      TabIndex        =   313
      Top             =   -15556
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LK"
      Height          =   285
      Index           =   155
      Left            =   2040
      TabIndex        =   311
      Top             =   -15876
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Mundhöhle"
      Height          =   285
      Index           =   154
      Left            =   2040
      TabIndex        =   309
      Top             =   -16196
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Zähne"
      Height          =   285
      Index           =   153
      Left            =   2040
      TabIndex        =   307
      Top             =   -16516
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NNH"
      Height          =   285
      Index           =   152
      Left            =   2040
      TabIndex        =   305
      Top             =   -16836
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Carotiden"
      Height          =   285
      Index           =   151
      Left            =   2040
      TabIndex        =   303
      Top             =   -17156
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SD"
      Height          =   285
      Index           =   150
      Left            =   2040
      TabIndex        =   301
      Top             =   -17476
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NL"
      Height          =   285
      Index           =   149
      Left            =   2040
      TabIndex        =   299
      Top             =   -17796
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "WS"
      Height          =   285
      Index           =   148
      Left            =   2040
      TabIndex        =   297
      Top             =   -18116
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Bauch"
      Height          =   285
      Index           =   147
      Left            =   2040
      TabIndex        =   295
      Top             =   -18436
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Lunge"
      Height          =   285
      Index           =   146
      Left            =   2040
      TabIndex        =   293
      Top             =   -18756
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Herz"
      Height          =   285
      Index           =   145
      Left            =   2040
      TabIndex        =   291
      Top             =   -19076
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RRTurboMed"
      Height          =   285
      Index           =   144
      Left            =   2040
      TabIndex        =   289
      Top             =   -19396
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RR"
      Height          =   285
      Index           =   143
      Left            =   2040
      TabIndex        =   287
      Top             =   -19716
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Puls Adp"
      Height          =   285
      Index           =   142
      Left            =   2040
      TabIndex        =   285
      Top             =   -20036
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Puls Atp"
      Height          =   285
      Index           =   141
      Left            =   2040
      TabIndex        =   283
      Top             =   -20356
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Puls Kniekehle"
      Height          =   285
      Index           =   140
      Left            =   2040
      TabIndex        =   281
      Top             =   -20676
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Puls Leiste"
      Height          =   285
      Index           =   139
      Left            =   2040
      TabIndex        =   279
      Top             =   -20996
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Vibration Großzehe"
      Height          =   285
      Index           =   138
      Left            =   2040
      TabIndex        =   277
      Top             =   -21316
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Vibration IK"
      Height          =   285
      Index           =   137
      Left            =   2040
      TabIndex        =   275
      Top             =   -21636
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Kalt-Warm"
      Height          =   285
      Index           =   136
      Left            =   2040
      TabIndex        =   273
      Top             =   -21956
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Monofilamenttest"
      Height          =   285
      Index           =   135
      Left            =   2040
      TabIndex        =   271
      Top             =   -22276
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Oberflächensensibilität"
      Height          =   285
      Index           =   134
      Left            =   2040
      TabIndex        =   269
      Top             =   -22596
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PSR"
      Height          =   285
      Index           =   133
      Left            =   2040
      TabIndex        =   267
      Top             =   -22916
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ASR"
      Height          =   285
      Index           =   132
      Left            =   2040
      TabIndex        =   265
      Top             =   -23236
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Kraft Knie"
      Height          =   285
      Index           =   131
      Left            =   2040
      TabIndex        =   263
      Top             =   -23556
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Kraft Zehenbeuger"
      Height          =   285
      Index           =   130
      Left            =   2040
      TabIndex        =   261
      Top             =   -23876
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Kraft Zehenheber"
      Height          =   285
      Index           =   129
      Left            =   2040
      TabIndex        =   259
      Top             =   -24196
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ulcera"
      Height          =   285
      Index           =   128
      Left            =   2040
      TabIndex        =   257
      Top             =   -24516
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Hyperkeratosen"
      Height          =   285
      Index           =   127
      Left            =   2040
      TabIndex        =   255
      Top             =   -24836
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Beinbefund"
      Height          =   285
      Index           =   126
      Left            =   2040
      TabIndex        =   253
      Top             =   -25156
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Liphypertrophien Arme"
      Height          =   285
      Index           =   125
      Left            =   2040
      TabIndex        =   251
      Top             =   -25476
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Liphypertrophien Beine"
      Height          =   285
      Index           =   124
      Left            =   2040
      TabIndex        =   249
      Top             =   -25796
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Liphypertrophien Abdomen"
      Height          =   285
      Index           =   123
      Left            =   2040
      TabIndex        =   247
      Top             =   -26116
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Weitere Medikation"
      Height          =   285
      Index           =   122
      Left            =   2040
      TabIndex        =   245
      Top             =   -26436
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tabak"
      Height          =   285
      Index           =   121
      Left            =   2040
      TabIndex        =   243
      Top             =   -26756
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Alkohol"
      Height          =   285
      Index           =   120
      Left            =   2040
      TabIndex        =   241
      Top             =   -27076
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Weitere Anamnese"
      Height          =   285
      Index           =   119
      Left            =   2040
      TabIndex        =   239
      Top             =   -27396
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Sexualstörung seit"
      Height          =   285
      Index           =   118
      Left            =   2040
      TabIndex        =   237
      Top             =   -27716
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Sexualstörung"
      Height          =   285
      Index           =   117
      Left            =   2040
      TabIndex        =   235
      Top             =   -28036
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Bewegungseinschränkungen"
      Height          =   285
      Index           =   116
      Left            =   2040
      TabIndex        =   233
      Top             =   -28356
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Folgeerkrankungen Haut"
      Height          =   285
      Index           =   115
      Left            =   2040
      TabIndex        =   231
      Top             =   -28676
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Schwindel Aufstehen"
      Height          =   285
      Index           =   114
      Left            =   2040
      TabIndex        =   229
      Top             =   -28996
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Entleerungsstörungen Harnblase"
      Height          =   285
      Index           =   113
      Left            =   2040
      TabIndex        =   227
      Top             =   -29316
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Entleerungsstörungen Magen"
      Height          =   285
      Index           =   112
      Left            =   2040
      TabIndex        =   225
      Top             =   -29636
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Neue Fußkomplikationen"
      Height          =   285
      Index           =   111
      Left            =   2040
      TabIndex        =   223
      Top             =   -29956
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Einlagen"
      Height          =   285
      Index           =   110
      Left            =   2040
      TabIndex        =   221
      Top             =   -30276
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Podologie"
      Height          =   285
      Index           =   109
      Left            =   2040
      TabIndex        =   219
      Top             =   -30596
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fußpflege"
      Height          =   285
      Index           =   108
      Left            =   2040
      TabIndex        =   217
      Top             =   -30916
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Verformungen Beschreibung"
      Height          =   285
      Index           =   107
      Left            =   2040
      TabIndex        =   215
      Top             =   -31236
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Verformungen"
      Height          =   285
      Index           =   106
      Left            =   2040
      TabIndex        =   213
      Top             =   -31556
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Druckstellen"
      Height          =   285
      Index           =   105
      Left            =   2040
      TabIndex        =   211
      Top             =   -31876
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ameisen Ausmaß"
      Height          =   285
      Index           =   104
      Left            =   2040
      TabIndex        =   209
      Top             =   -32196
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ameisenlaufen"
      Height          =   285
      Index           =   103
      Left            =   2040
      TabIndex        =   207
      Top             =   -32516
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "pAVK Beschreibung"
      Height          =   285
      Index           =   102
      Left            =   2040
      TabIndex        =   205
      Top             =   32700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Amputation"
      Height          =   285
      Index           =   101
      Left            =   2040
      TabIndex        =   203
      Top             =   32380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Geschwür"
      Height          =   285
      Index           =   100
      Left            =   2040
      TabIndex        =   201
      Top             =   32060
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Bypaß peripher"
      Height          =   285
      Index           =   99
      Left            =   2040
      TabIndex        =   199
      Top             =   31740
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Schaufensterkrankheit"
      Height          =   285
      Index           =   98
      Left            =   2040
      TabIndex        =   197
      Top             =   31420
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Beindurchblutungsstörung"
      Height          =   285
      Index           =   97
      Left            =   2040
      TabIndex        =   195
      Top             =   31100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Schlaganfall"
      Height          =   285
      Index           =   96
      Left            =   2040
      TabIndex        =   193
      Top             =   30780
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Hirndurchblutungsstörung"
      Height          =   285
      Index           =   95
      Left            =   2040
      TabIndex        =   191
      Top             =   30460
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Herzkrankheit Beschreibung"
      Height          =   285
      Index           =   94
      Left            =   2040
      TabIndex        =   189
      Top             =   30140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Herzschwäche"
      Height          =   285
      Index           =   93
      Left            =   2040
      TabIndex        =   187
      Top             =   29820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Bypass wann"
      Height          =   285
      Index           =   92
      Left            =   2040
      TabIndex        =   185
      Top             =   29500
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Bypass kardial"
      Height          =   285
      Index           =   91
      Left            =   2040
      TabIndex        =   183
      Top             =   29180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PTCA oder Stent"
      Height          =   285
      Index           =   90
      Left            =   2040
      TabIndex        =   181
      Top             =   28860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Herzinfarkt wann"
      Height          =   285
      Index           =   89
      Left            =   2040
      TabIndex        =   179
      Top             =   28540
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Herzinfarkt"
      Height          =   285
      Index           =   88
      Left            =   2040
      TabIndex        =   177
      Top             =   28220
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Angina pectoris"
      Height          =   285
      Index           =   87
      Left            =   2040
      TabIndex        =   175
      Top             =   27900
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Herzkrankheit"
      Height          =   285
      Index           =   86
      Left            =   2040
      TabIndex        =   173
      Top             =   27580
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "andere Nierenerkrankung"
      Height          =   285
      Index           =   85
      Left            =   2040
      TabIndex        =   171
      Top             =   27260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Dialyse seit"
      Height          =   285
      Index           =   84
      Left            =   2040
      TabIndex        =   169
      Top             =   26940
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Dialyse"
      Height          =   285
      Index           =   83
      Left            =   2040
      TabIndex        =   167
      Top             =   26620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "erhöht?"
      Height          =   285
      Index           =   82
      Left            =   2040
      TabIndex        =   165
      Top             =   26300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Albumin zuletzt"
      Height          =   285
      Index           =   81
      Left            =   2040
      TabIndex        =   163
      Top             =   25980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Diabet Nierenschaden"
      Height          =   285
      Index           =   80
      Left            =   2040
      TabIndex        =   161
      Top             =   25660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Sehminderung unbehebbar"
      Height          =   285
      Index           =   79
      Left            =   2040
      TabIndex        =   159
      Top             =   25340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Netzhaut gelasert"
      Height          =   285
      Index           =   78
      Left            =   2040
      TabIndex        =   157
      Top             =   25020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Augensp Befund"
      Height          =   285
      Index           =   77
      Left            =   2040
      TabIndex        =   155
      Top             =   24700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Augensp zuletzt"
      Height          =   285
      Index           =   76
      Left            =   2040
      TabIndex        =   153
      Top             =   24380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Schwanger seit"
      Height          =   285
      Index           =   75
      Left            =   2040
      TabIndex        =   151
      Top             =   24060
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Schwanger"
      Height          =   285
      Index           =   74
      Left            =   2040
      TabIndex        =   149
      Top             =   23740
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BDselbst"
      Height          =   285
      Index           =   73
      Left            =   2040
      TabIndex        =   147
      Top             =   23420
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Blutdruckwerte"
      Height          =   285
      Index           =   72
      Left            =   2040
      TabIndex        =   145
      Top             =   23100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BHD beh mit"
      Height          =   285
      Index           =   71
      Left            =   2040
      TabIndex        =   143
      Top             =   22780
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BHD seit"
      Height          =   285
      Index           =   70
      Left            =   2040
      TabIndex        =   141
      Top             =   22460
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Bluthochdruck"
      Height          =   285
      Index           =   69
      Left            =   2040
      TabIndex        =   139
      Top             =   22140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BZgr300 pM"
      Height          =   285
      Index           =   68
      Left            =   2040
      TabIndex        =   137
      Top             =   21820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Keto pa"
      Height          =   285
      Index           =   67
      Left            =   2040
      TabIndex        =   135
      Top             =   21500
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Bewußtlos pa"
      Height          =   285
      Index           =   66
      Left            =   2040
      TabIndex        =   133
      Top             =   21180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fremde Hilfe pa"
      Height          =   285
      Index           =   65
      Left            =   2040
      TabIndex        =   131
      Top             =   20860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "UZ rechtzeitig"
      Height          =   285
      Index           =   64
      Left            =   2040
      TabIndex        =   129
      Top             =   20540
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Unterzucker pM"
      Height          =   285
      Index           =   63
      Left            =   2040
      TabIndex        =   127
      Top             =   20220
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "UZ Tageszeit"
      Height          =   285
      Index           =   62
      Left            =   2040
      TabIndex        =   125
      Top             =   19900
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BZWerte n d Essen"
      Height          =   285
      Index           =   61
      Left            =   2040
      TabIndex        =   123
      Top             =   19580
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BZWerte v d Essen"
      Height          =   285
      Index           =   60
      Left            =   2040
      TabIndex        =   121
      Top             =   19260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Aufschreiben"
      Height          =   285
      Index           =   59
      Left            =   2040
      TabIndex        =   119
      Top             =   18940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BZMessungen p W nachts"
      Height          =   285
      Index           =   58
      Left            =   2040
      TabIndex        =   117
      Top             =   18620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BZMessungen pW ndE"
      Height          =   285
      Index           =   57
      Left            =   2040
      TabIndex        =   115
      Top             =   18300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BZMessungen pW"
      Height          =   285
      Index           =   56
      Left            =   2040
      TabIndex        =   113
      Top             =   17980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Gerät"
      Height          =   285
      Index           =   55
      Left            =   2040
      TabIndex        =   111
      Top             =   17660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BZMessungen selbst"
      Height          =   285
      Index           =   54
      Left            =   2040
      TabIndex        =   109
      Top             =   17340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "vorherige Werte"
      Height          =   285
      Index           =   53
      Left            =   2040
      TabIndex        =   107
      Top             =   17020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "gemessen am"
      Height          =   285
      Index           =   52
      Left            =   2040
      TabIndex        =   105
      Top             =   16700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "letztes HbA1c"
      Height          =   285
      Index           =   51
      Left            =   2040
      TabIndex        =   103
      Top             =   16380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ort Schulung"
      Height          =   285
      Index           =   50
      Left            =   2040
      TabIndex        =   101
      Top             =   16060
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Jahr letzte Diabetesschulung"
      Height          =   285
      Index           =   49
      Left            =   2040
      TabIndex        =   99
      Top             =   15740
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Spritzstelle nachts"
      Height          =   285
      Index           =   48
      Left            =   2040
      TabIndex        =   97
      Top             =   15420
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Spritzstelle abends"
      Height          =   285
      Index           =   47
      Left            =   2040
      TabIndex        =   95
      Top             =   15100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Spritzstelle mittags"
      Height          =   285
      Index           =   46
      Left            =   2040
      TabIndex        =   93
      Top             =   14780
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Spritzstelle früh"
      Height          =   285
      Index           =   45
      Left            =   2040
      TabIndex        =   91
      Top             =   14460
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Spritz-Eß-Abstand abends"
      Height          =   285
      Index           =   44
      Left            =   2040
      TabIndex        =   89
      Top             =   14140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Spritz-Eß-Abstand mittags"
      Height          =   285
      Index           =   43
      Left            =   2040
      TabIndex        =   87
      Top             =   13820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Spritz-Eß-Abstand früh"
      Height          =   285
      Index           =   42
      Left            =   2040
      TabIndex        =   85
      Top             =   13500
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Essenszeit spät"
      Height          =   285
      Index           =   41
      Left            =   2040
      TabIndex        =   83
      Top             =   13180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Essenszeit abends"
      Height          =   285
      Index           =   40
      Left            =   2040
      TabIndex        =   81
      Top             =   12860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Essenszeit nachmittags"
      Height          =   285
      Index           =   39
      Left            =   2040
      TabIndex        =   79
      Top             =   12540
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Essenszeit mittags"
      Height          =   285
      Index           =   38
      Left            =   2040
      TabIndex        =   77
      Top             =   12220
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Essenszeit vormittags"
      Height          =   285
      Index           =   37
      Left            =   2040
      TabIndex        =   75
      Top             =   11900
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Essenszeit früh"
      Height          =   285
      Index           =   36
      Left            =   2040
      TabIndex        =   73
      Top             =   11580
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Broteinheiten nachts"
      Height          =   285
      Index           =   35
      Left            =   2040
      TabIndex        =   71
      Top             =   11260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Broteinheiten abends"
      Height          =   285
      Index           =   34
      Left            =   2040
      TabIndex        =   69
      Top             =   10940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Broteinheiten nachmittags"
      Height          =   285
      Index           =   33
      Left            =   2040
      TabIndex        =   67
      Top             =   10620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Broteinheiten mittags"
      Height          =   285
      Index           =   32
      Left            =   2040
      TabIndex        =   65
      Top             =   10300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Broteinheiten ZM früh"
      Height          =   285
      Index           =   31
      Left            =   2040
      TabIndex        =   63
      Top             =   9980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Broteinheiten früh"
      Height          =   285
      Index           =   30
      Left            =   2040
      TabIndex        =   61
      Top             =   9660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Broteinheiten gesamt"
      Height          =   285
      Index           =   29
      Left            =   2040
      TabIndex        =   59
      Top             =   9340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Insulinpumpe Marke"
      Height          =   285
      Index           =   28
      Left            =   2040
      TabIndex        =   57
      Top             =   9020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Insulinpumpe seit"
      Height          =   285
      Index           =   27
      Left            =   2040
      TabIndex        =   55
      Top             =   8700
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Insulinpumpe"
      Height          =   285
      Index           =   26
      Left            =   2040
      TabIndex        =   53
      Top             =   8380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DiabetesMedikament 4 Menge"
      Height          =   285
      Index           =   25
      Left            =   2040
      TabIndex        =   51
      Top             =   8060
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DiabetesMedikament 4"
      Height          =   285
      Index           =   24
      Left            =   2040
      TabIndex        =   49
      Top             =   7740
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DiabetesMedikament 3 Menge"
      Height          =   285
      Index           =   23
      Left            =   2040
      TabIndex        =   47
      Top             =   7420
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DiabetesMedikament 3"
      Height          =   285
      Index           =   22
      Left            =   2040
      TabIndex        =   45
      Top             =   7100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DiabetesMedikament 2 Menge"
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   43
      Top             =   6780
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DiabetesMedikament 2"
      Height          =   285
      Index           =   20
      Left            =   2040
      TabIndex        =   41
      Top             =   6460
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DiabetesMedikament 1 Menge"
      Height          =   285
      Index           =   19
      Left            =   2040
      TabIndex        =   39
      Top             =   6140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DiabetesMedikament 1"
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   37
      Top             =   5820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tendenz"
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   35
      Top             =   5500
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Gewicht"
      Height          =   285
      Index           =   16
      Left            =   2040
      TabIndex        =   33
      Top             =   5180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Größe"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   31
      Top             =   4860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Familienanamnese"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   29
      Top             =   4540
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Grund für Vorstellung"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   27
      Top             =   4220
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Insulin seit"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   25
      Top             =   3900
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tabletten seit"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Top             =   3580
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Diabetes seit"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   3260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Diabetestyp"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   2940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Versicherungsart"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2620
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Tkz"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GebDat"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Anrede"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Titel"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NVorsatz"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Vorname"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Nachname"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Pat_id"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   6480
      TabIndex        =   404
      Top             =   6720
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1815
      Left            =   9840
      TabIndex        =   408
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   7560
      X2              =   10800
      Y1              =   2160
      Y2              =   2880
   End
   Begin VB.OLE OLE2 
      Height          =   2055
      Left            =   6120
      TabIndex        =   398
      Top             =   4200
      Width           =   3375
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      Height          =   495
      Left            =   6000
      OleObjectBlob   =   "Anamnese.frx":0D0E
      TabIndex        =   397
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "QT:"
      Height          =   255
      Index           =   189
      Left            =   120
      TabIndex        =   378
      Top             =   -4996
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "QS:"
      Height          =   255
      Index           =   188
      Left            =   120
      TabIndex        =   376
      Top             =   -5316
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ob:"
      Height          =   255
      Index           =   187
      Left            =   120
      TabIndex        =   374
      Top             =   -5636
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Hausarzt:"
      Height          =   255
      Index           =   186
      Left            =   120
      TabIndex        =   372
      Top             =   -5956
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obMedNetz:"
      Height          =   255
      Index           =   185
      Left            =   120
      TabIndex        =   370
      Top             =   -6276
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obDMPaufgekl:"
      Height          =   255
      Index           =   184
      Left            =   120
      TabIndex        =   368
      Top             =   -6596
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obSchulaufgek:"
      Height          =   255
      Index           =   183
      Left            =   120
      TabIndex        =   366
      Top             =   -6916
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obMBlAusgeh:"
      Height          =   255
      Index           =   182
      Left            =   120
      TabIndex        =   364
      Top             =   -7236
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obPodAufgek:"
      Height          =   255
      Index           =   181
      Left            =   120
      TabIndex        =   362
      Top             =   -7556
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obOSaufgek:"
      Height          =   255
      Index           =   180
      Left            =   120
      TabIndex        =   360
      Top             =   -7876
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obBZausgew:"
      Height          =   255
      Index           =   179
      Left            =   120
      TabIndex        =   358
      Top             =   -8196
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obCheck:"
      Height          =   255
      Index           =   178
      Left            =   120
      TabIndex        =   356
      Top             =   -8516
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obAnAeing:"
      Height          =   255
      Index           =   177
      Left            =   120
      TabIndex        =   354
      Top             =   -8836
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obAn2eing:"
      Height          =   255
      Index           =   176
      Left            =   120
      TabIndex        =   352
      Top             =   -9156
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "obAn1eing:"
      Height          =   255
      Index           =   175
      Left            =   120
      TabIndex        =   350
      Top             =   -9476
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Prim:"
      Height          =   255
      Index           =   174
      Left            =   120
      TabIndex        =   348
      Top             =   -9796
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TherAkt:"
      Height          =   255
      Index           =   173
      Left            =   120
      TabIndex        =   346
      Top             =   -10116
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ther1:"
      Height          =   255
      Index           =   172
      Left            =   120
      TabIndex        =   344
      Top             =   -10436
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "AktZeit:"
      Height          =   255
      Index           =   171
      Left            =   120
      TabIndex        =   342
      Top             =   -10756
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Versicherung:"
      Height          =   255
      Index           =   170
      Left            =   120
      TabIndex        =   340
      Top             =   -11076
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vorgestellt:"
      Height          =   255
      Index           =   169
      Left            =   120
      TabIndex        =   338
      Top             =   -11396
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Diagnosen:"
      Height          =   255
      Index           =   168
      Left            =   120
      TabIndex        =   336
      Top             =   -11716
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "letzte Änderung:"
      Height          =   255
      Index           =   167
      Left            =   120
      TabIndex        =   334
      Top             =   -12036
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "HANr2:"
      Height          =   255
      Index           =   166
      Left            =   120
      TabIndex        =   332
      Top             =   -12356
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "HANr:"
      Height          =   255
      Index           =   165
      Left            =   120
      TabIndex        =   330
      Top             =   -12676
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DMPhier:"
      Height          =   255
      Index           =   164
      Left            =   120
      TabIndex        =   328
      Top             =   -12996
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RRSchulz:"
      Height          =   255
      Index           =   163
      Left            =   120
      TabIndex        =   326
      Top             =   -13316
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DMSchL:"
      Height          =   255
      Index           =   162
      Left            =   120
      TabIndex        =   324
      Top             =   -13636
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DMSchulz:"
      Height          =   255
      Index           =   161
      Left            =   120
      TabIndex        =   322
      Top             =   -13956
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DMP:"
      Height          =   255
      Index           =   160
      Left            =   120
      TabIndex        =   320
      Top             =   -14276
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Schulung:"
      Height          =   255
      Index           =   159
      Left            =   120
      TabIndex        =   318
      Top             =   -14596
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Weitere Befunde:"
      Height          =   255
      Index           =   158
      Left            =   120
      TabIndex        =   316
      Top             =   -14916
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Neuro sonst:"
      Height          =   255
      Index           =   157
      Left            =   120
      TabIndex        =   314
      Top             =   -15236
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BeinödVen:"
      Height          =   255
      Index           =   156
      Left            =   120
      TabIndex        =   312
      Top             =   -15556
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LK:"
      Height          =   255
      Index           =   155
      Left            =   120
      TabIndex        =   310
      Top             =   -15876
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Mundhöhle:"
      Height          =   255
      Index           =   154
      Left            =   120
      TabIndex        =   308
      Top             =   -16196
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Zähne:"
      Height          =   255
      Index           =   153
      Left            =   120
      TabIndex        =   306
      Top             =   -16516
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NNH:"
      Height          =   255
      Index           =   152
      Left            =   120
      TabIndex        =   304
      Top             =   -16836
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Carotiden:"
      Height          =   255
      Index           =   151
      Left            =   120
      TabIndex        =   302
      Top             =   -17156
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "SD:"
      Height          =   255
      Index           =   150
      Left            =   120
      TabIndex        =   300
      Top             =   -17476
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NL:"
      Height          =   255
      Index           =   149
      Left            =   120
      TabIndex        =   298
      Top             =   -17796
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "WS:"
      Height          =   255
      Index           =   148
      Left            =   120
      TabIndex        =   296
      Top             =   -18116
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bauch:"
      Height          =   255
      Index           =   147
      Left            =   120
      TabIndex        =   294
      Top             =   -18436
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Lunge:"
      Height          =   255
      Index           =   146
      Left            =   120
      TabIndex        =   292
      Top             =   -18756
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Herz:"
      Height          =   255
      Index           =   145
      Left            =   120
      TabIndex        =   290
      Top             =   -19076
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RRTurboMed:"
      Height          =   255
      Index           =   144
      Left            =   120
      TabIndex        =   288
      Top             =   -19396
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RR:"
      Height          =   255
      Index           =   143
      Left            =   120
      TabIndex        =   286
      Top             =   -19716
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Puls Adp:"
      Height          =   255
      Index           =   142
      Left            =   120
      TabIndex        =   284
      Top             =   -20036
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Puls Atp:"
      Height          =   255
      Index           =   141
      Left            =   120
      TabIndex        =   282
      Top             =   -20356
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Puls Kniekehle:"
      Height          =   255
      Index           =   140
      Left            =   120
      TabIndex        =   280
      Top             =   -20676
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Puls Leiste:"
      Height          =   255
      Index           =   139
      Left            =   120
      TabIndex        =   278
      Top             =   -20996
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vibration Großzehe:"
      Height          =   255
      Index           =   138
      Left            =   120
      TabIndex        =   276
      Top             =   -21316
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vibration IK:"
      Height          =   255
      Index           =   137
      Left            =   120
      TabIndex        =   274
      Top             =   -21636
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Kalt-Warm:"
      Height          =   255
      Index           =   136
      Left            =   120
      TabIndex        =   272
      Top             =   -21956
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Monofilamenttest:"
      Height          =   255
      Index           =   135
      Left            =   120
      TabIndex        =   270
      Top             =   -22276
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Oberflächensensibilität:"
      Height          =   255
      Index           =   134
      Left            =   120
      TabIndex        =   268
      Top             =   -22596
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PSR:"
      Height          =   255
      Index           =   133
      Left            =   120
      TabIndex        =   266
      Top             =   -22916
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ASR:"
      Height          =   255
      Index           =   132
      Left            =   120
      TabIndex        =   264
      Top             =   -23236
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Kraft Knie:"
      Height          =   255
      Index           =   131
      Left            =   120
      TabIndex        =   262
      Top             =   -23556
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Kraft Zehenbeuger:"
      Height          =   255
      Index           =   130
      Left            =   120
      TabIndex        =   260
      Top             =   -23876
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Kraft Zehenheber:"
      Height          =   255
      Index           =   129
      Left            =   120
      TabIndex        =   258
      Top             =   -24196
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ulcera:"
      Height          =   255
      Index           =   128
      Left            =   120
      TabIndex        =   256
      Top             =   -24516
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Hyperkeratosen:"
      Height          =   255
      Index           =   127
      Left            =   120
      TabIndex        =   254
      Top             =   -24836
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Beinbefund:"
      Height          =   255
      Index           =   126
      Left            =   120
      TabIndex        =   252
      Top             =   -25156
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Liphypertrophien Arme:"
      Height          =   255
      Index           =   125
      Left            =   120
      TabIndex        =   250
      Top             =   -25476
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Liphypertrophien Beine:"
      Height          =   255
      Index           =   124
      Left            =   120
      TabIndex        =   248
      Top             =   -25796
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Liphypertrophien Abdomen:"
      Height          =   255
      Index           =   123
      Left            =   120
      TabIndex        =   246
      Top             =   -26116
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Weitere Medikation:"
      Height          =   255
      Index           =   122
      Left            =   120
      TabIndex        =   244
      Top             =   -26436
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tabak:"
      Height          =   255
      Index           =   121
      Left            =   120
      TabIndex        =   242
      Top             =   -26756
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Alkohol:"
      Height          =   255
      Index           =   120
      Left            =   120
      TabIndex        =   240
      Top             =   -27076
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Weitere Anamnese:"
      Height          =   255
      Index           =   119
      Left            =   120
      TabIndex        =   238
      Top             =   -27396
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sexualstörung seit:"
      Height          =   255
      Index           =   118
      Left            =   120
      TabIndex        =   236
      Top             =   -27716
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sexualstörung:"
      Height          =   255
      Index           =   117
      Left            =   120
      TabIndex        =   234
      Top             =   -28036
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bewegungseinschränkungen:"
      Height          =   255
      Index           =   116
      Left            =   120
      TabIndex        =   232
      Top             =   -28356
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Folgeerkrankungen Haut:"
      Height          =   255
      Index           =   115
      Left            =   120
      TabIndex        =   230
      Top             =   -28676
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Schwindel Aufstehen:"
      Height          =   255
      Index           =   114
      Left            =   120
      TabIndex        =   228
      Top             =   -28996
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Entleerungsstörungen Harnblase:"
      Height          =   255
      Index           =   113
      Left            =   120
      TabIndex        =   226
      Top             =   -29316
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Entleerungsstörungen Magen:"
      Height          =   255
      Index           =   112
      Left            =   120
      TabIndex        =   224
      Top             =   -29636
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Neue Fußkomplikationen:"
      Height          =   255
      Index           =   111
      Left            =   120
      TabIndex        =   222
      Top             =   -29956
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Einlagen:"
      Height          =   255
      Index           =   110
      Left            =   120
      TabIndex        =   220
      Top             =   -30276
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Podologie:"
      Height          =   255
      Index           =   109
      Left            =   120
      TabIndex        =   218
      Top             =   -30596
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fußpflege:"
      Height          =   255
      Index           =   108
      Left            =   120
      TabIndex        =   216
      Top             =   -30916
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Verformungen Beschreibung:"
      Height          =   255
      Index           =   107
      Left            =   120
      TabIndex        =   214
      Top             =   -31236
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Verformungen:"
      Height          =   255
      Index           =   106
      Left            =   120
      TabIndex        =   212
      Top             =   -31556
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Druckstellen:"
      Height          =   255
      Index           =   105
      Left            =   120
      TabIndex        =   210
      Top             =   -31876
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ameisen Ausmaß:"
      Height          =   255
      Index           =   104
      Left            =   120
      TabIndex        =   208
      Top             =   -32196
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ameisenlaufen:"
      Height          =   255
      Index           =   103
      Left            =   120
      TabIndex        =   206
      Top             =   -32516
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "pAVK Beschreibung:"
      Height          =   255
      Index           =   102
      Left            =   120
      TabIndex        =   204
      Top             =   32700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Amputation:"
      Height          =   255
      Index           =   101
      Left            =   120
      TabIndex        =   202
      Top             =   32380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Geschwür:"
      Height          =   255
      Index           =   100
      Left            =   120
      TabIndex        =   200
      Top             =   32060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bypaß peripher:"
      Height          =   255
      Index           =   99
      Left            =   120
      TabIndex        =   198
      Top             =   31740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Schaufensterkrankheit:"
      Height          =   255
      Index           =   98
      Left            =   120
      TabIndex        =   196
      Top             =   31420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Beindurchblutungsstörung:"
      Height          =   255
      Index           =   97
      Left            =   120
      TabIndex        =   194
      Top             =   31100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Schlaganfall:"
      Height          =   255
      Index           =   96
      Left            =   120
      TabIndex        =   192
      Top             =   30780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Hirndurchblutungsstörung:"
      Height          =   255
      Index           =   95
      Left            =   120
      TabIndex        =   190
      Top             =   30460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Herzkrankheit Beschreibung:"
      Height          =   255
      Index           =   94
      Left            =   120
      TabIndex        =   188
      Top             =   30140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Herzschwäche:"
      Height          =   255
      Index           =   93
      Left            =   120
      TabIndex        =   186
      Top             =   29820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bypass wann:"
      Height          =   255
      Index           =   92
      Left            =   120
      TabIndex        =   184
      Top             =   29500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bypass kardial:"
      Height          =   255
      Index           =   91
      Left            =   120
      TabIndex        =   182
      Top             =   29180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PTCA oder Stent:"
      Height          =   255
      Index           =   90
      Left            =   120
      TabIndex        =   180
      Top             =   28860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Herzinfarkt wann:"
      Height          =   255
      Index           =   89
      Left            =   120
      TabIndex        =   178
      Top             =   28540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Herzinfarkt:"
      Height          =   255
      Index           =   88
      Left            =   120
      TabIndex        =   176
      Top             =   28220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Angina pectoris:"
      Height          =   255
      Index           =   87
      Left            =   120
      TabIndex        =   174
      Top             =   27900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Herzkrankheit:"
      Height          =   255
      Index           =   86
      Left            =   120
      TabIndex        =   172
      Top             =   27580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "andere Nierenerkrankung:"
      Height          =   255
      Index           =   85
      Left            =   120
      TabIndex        =   170
      Top             =   27260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Dialyse seit:"
      Height          =   255
      Index           =   84
      Left            =   120
      TabIndex        =   168
      Top             =   26940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Dialyse:"
      Height          =   255
      Index           =   83
      Left            =   120
      TabIndex        =   166
      Top             =   26620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "erhöht?:"
      Height          =   255
      Index           =   82
      Left            =   120
      TabIndex        =   164
      Top             =   26300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Albumin zuletzt:"
      Height          =   255
      Index           =   81
      Left            =   120
      TabIndex        =   162
      Top             =   25980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Diabet Nierenschaden:"
      Height          =   255
      Index           =   80
      Left            =   120
      TabIndex        =   160
      Top             =   25660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sehminderung unbehebbar:"
      Height          =   255
      Index           =   79
      Left            =   120
      TabIndex        =   158
      Top             =   25340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Netzhaut gelasert:"
      Height          =   255
      Index           =   78
      Left            =   120
      TabIndex        =   156
      Top             =   25020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Augensp Befund:"
      Height          =   255
      Index           =   77
      Left            =   120
      TabIndex        =   154
      Top             =   24700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Augensp zuletzt:"
      Height          =   255
      Index           =   76
      Left            =   120
      TabIndex        =   152
      Top             =   24380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Schwanger seit:"
      Height          =   255
      Index           =   75
      Left            =   120
      TabIndex        =   150
      Top             =   24060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Schwanger:"
      Height          =   255
      Index           =   74
      Left            =   120
      TabIndex        =   148
      Top             =   23740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BDselbst:"
      Height          =   255
      Index           =   73
      Left            =   120
      TabIndex        =   146
      Top             =   23420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Blutdruckwerte:"
      Height          =   255
      Index           =   72
      Left            =   120
      TabIndex        =   144
      Top             =   23100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BHD beh mit:"
      Height          =   255
      Index           =   71
      Left            =   120
      TabIndex        =   142
      Top             =   22780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BHD seit:"
      Height          =   255
      Index           =   70
      Left            =   120
      TabIndex        =   140
      Top             =   22460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bluthochdruck:"
      Height          =   255
      Index           =   69
      Left            =   120
      TabIndex        =   138
      Top             =   22140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BZgr300 pM:"
      Height          =   255
      Index           =   68
      Left            =   120
      TabIndex        =   136
      Top             =   21820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Keto pa:"
      Height          =   255
      Index           =   67
      Left            =   120
      TabIndex        =   134
      Top             =   21500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bewußtlos pa:"
      Height          =   255
      Index           =   66
      Left            =   120
      TabIndex        =   132
      Top             =   21180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fremde Hilfe pa:"
      Height          =   255
      Index           =   65
      Left            =   120
      TabIndex        =   130
      Top             =   20860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "UZ rechtzeitig:"
      Height          =   255
      Index           =   64
      Left            =   120
      TabIndex        =   128
      Top             =   20540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Unterzucker pM:"
      Height          =   255
      Index           =   63
      Left            =   120
      TabIndex        =   126
      Top             =   20220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "UZ Tageszeit:"
      Height          =   255
      Index           =   62
      Left            =   120
      TabIndex        =   124
      Top             =   19900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BZWerte n d Essen:"
      Height          =   255
      Index           =   61
      Left            =   120
      TabIndex        =   122
      Top             =   19580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BZWerte v d Essen:"
      Height          =   255
      Index           =   60
      Left            =   120
      TabIndex        =   120
      Top             =   19260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Aufschreiben:"
      Height          =   255
      Index           =   59
      Left            =   120
      TabIndex        =   118
      Top             =   18940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BZMessungen p W nachts:"
      Height          =   255
      Index           =   58
      Left            =   120
      TabIndex        =   116
      Top             =   18620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BZMessungen pW ndE:"
      Height          =   255
      Index           =   57
      Left            =   120
      TabIndex        =   114
      Top             =   18300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BZMessungen pW:"
      Height          =   255
      Index           =   56
      Left            =   120
      TabIndex        =   112
      Top             =   17980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Gerät:"
      Height          =   255
      Index           =   55
      Left            =   120
      TabIndex        =   110
      Top             =   17660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BZMessungen selbst:"
      Height          =   255
      Index           =   54
      Left            =   120
      TabIndex        =   108
      Top             =   17340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "vorherige Werte:"
      Height          =   255
      Index           =   53
      Left            =   120
      TabIndex        =   106
      Top             =   17020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "gemessen am:"
      Height          =   255
      Index           =   52
      Left            =   120
      TabIndex        =   104
      Top             =   16700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "letztes HbA1c:"
      Height          =   255
      Index           =   51
      Left            =   120
      TabIndex        =   102
      Top             =   16380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ort Schulung:"
      Height          =   255
      Index           =   50
      Left            =   120
      TabIndex        =   100
      Top             =   16060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Jahr letzte Diabetesschulung:"
      Height          =   255
      Index           =   49
      Left            =   120
      TabIndex        =   98
      Top             =   15740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Spritzstelle nachts:"
      Height          =   255
      Index           =   48
      Left            =   120
      TabIndex        =   96
      Top             =   15420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Spritzstelle abends:"
      Height          =   255
      Index           =   47
      Left            =   120
      TabIndex        =   94
      Top             =   15100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Spritzstelle mittags:"
      Height          =   255
      Index           =   46
      Left            =   120
      TabIndex        =   92
      Top             =   14780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Spritzstelle früh:"
      Height          =   255
      Index           =   45
      Left            =   120
      TabIndex        =   90
      Top             =   14460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Spritz-Eß-Abstand abends:"
      Height          =   255
      Index           =   44
      Left            =   120
      TabIndex        =   88
      Top             =   14140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Spritz-Eß-Abstand mittags:"
      Height          =   255
      Index           =   43
      Left            =   120
      TabIndex        =   86
      Top             =   13820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Spritz-Eß-Abstand früh:"
      Height          =   255
      Index           =   42
      Left            =   120
      TabIndex        =   84
      Top             =   13500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Essenszeit spät:"
      Height          =   255
      Index           =   41
      Left            =   120
      TabIndex        =   82
      Top             =   13180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Essenszeit abends:"
      Height          =   255
      Index           =   40
      Left            =   120
      TabIndex        =   80
      Top             =   12860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Essenszeit nachmittags:"
      Height          =   255
      Index           =   39
      Left            =   120
      TabIndex        =   78
      Top             =   12540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Essenszeit mittags:"
      Height          =   255
      Index           =   38
      Left            =   120
      TabIndex        =   76
      Top             =   12220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Essenszeit vormittags:"
      Height          =   255
      Index           =   37
      Left            =   120
      TabIndex        =   74
      Top             =   11900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Essenszeit früh:"
      Height          =   255
      Index           =   36
      Left            =   120
      TabIndex        =   72
      Top             =   11580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Broteinheiten nachts:"
      Height          =   255
      Index           =   35
      Left            =   120
      TabIndex        =   70
      Top             =   11260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Broteinheiten abends:"
      Height          =   255
      Index           =   34
      Left            =   120
      TabIndex        =   68
      Top             =   10940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Broteinheiten nachmittags:"
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   66
      Top             =   10620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Broteinheiten mittags:"
      Height          =   255
      Index           =   32
      Left            =   120
      TabIndex        =   64
      Top             =   10300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Broteinheiten ZM früh:"
      Height          =   255
      Index           =   31
      Left            =   120
      TabIndex        =   62
      Top             =   9980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Broteinheiten früh:"
      Height          =   255
      Index           =   30
      Left            =   120
      TabIndex        =   60
      Top             =   9660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Broteinheiten gesamt:"
      Height          =   255
      Index           =   29
      Left            =   120
      TabIndex        =   58
      Top             =   9340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Insulinpumpe Marke:"
      Height          =   255
      Index           =   28
      Left            =   120
      TabIndex        =   56
      Top             =   9020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Insulinpumpe seit:"
      Height          =   255
      Index           =   27
      Left            =   120
      TabIndex        =   54
      Top             =   8700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Insulinpumpe:"
      Height          =   255
      Index           =   26
      Left            =   120
      TabIndex        =   52
      Top             =   8380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DiabetesMedikament 4 Menge:"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   50
      Top             =   8060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DiabetesMedikament 4:"
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   48
      Top             =   7740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DiabetesMedikament 3 Menge:"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   46
      Top             =   7420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DiabetesMedikament 3:"
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   44
      Top             =   7100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DiabetesMedikament 2 Menge:"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   42
      Top             =   6780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DiabetesMedikament 2:"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   40
      Top             =   6460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DiabetesMedikament 1 Menge:"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   38
      Top             =   6140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DiabetesMedikament 1:"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   36
      Top             =   5820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tendenz:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   5500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Gewicht:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Top             =   5180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Größe:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   30
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Familienanamnese:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   4540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Grund für Vorstellung:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Insulin seit:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tabletten seit:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Diabetes seit:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Diabetestyp:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Versicherungsart:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tkz:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "GebDat:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Anrede:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Titel:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NVorsatz:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vorname:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nachname:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Pat_id:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "Anamnese"
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

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
'  db.Open "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=linux;uid=praxis;pwd=sonne;database=quelle;"
  db.Open DBCn

  Set adoPrimaryRS = New Recordset
  'adoPrimaryRS.Open "select Pat_id,Nachname,Vorname,NVorsatz,Titel,Anrede,GebDat,Tkz,Versicherungsart,Diabetestyp,`Diabetes seit`,`Tabletten seit`,`Insulin seit`,`Grund für Vorstellung`,Familienanamnese,Größe,Gewicht,Tendenz,`DiabetesMedikament 1`,`DiabetesMedikament 1 Menge`,`DiabetesMedikament 2`,`DiabetesMedikament 2 Menge`,`DiabetesMedikament 3`,`DiabetesMedikament 3 Menge`,`DiabetesMedikament 4`,`DiabetesMedikament 4 Menge`,Insulinpumpe,`Insulinpumpe seit`,`Insulinpumpe Marke`,`Broteinheiten gesamt`,`Broteinheiten früh`,`Broteinheiten ZM früh`,`Broteinheiten mittags`,`Broteinheiten nachmittags`,`Broteinheiten abends`,`Broteinheiten nachts`,`Essenszeit früh`,`Essenszeit vormittags`,`Essenszeit mittags`,`Essenszeit nachmittags`,`Essenszeit abends`,`Essenszeit spät`,`Spritz-Eß-Abstand früh`,`Spritz-Eß-Abstand mittags`,`Spritz-Eß-Abstand abends`,`Spritzstelle früh`,`Spritzstelle mittags`,`Spritzstelle abends`,`Spritzstelle nachts`,`Jahr letzte Diabetesschulung`,`Ort Schulung`,`letztes HbA1c`," & _
"`gemessen am`,`vorherige Werte`,`BZMessungen selbst`,Gerät,`BZMessungen pW`,`BZMessungen pW ndE`,`BZMessungen p W nachts`,Aufschreiben,`BZWerte v d Essen`,`BZWerte n d Essen`,`UZ Tageszeit`,`Unterzucker pM`,`UZ rechtzeitig`,`Fremde Hilfe pa`,`Bewußtlos pa`,`Keto pa`,`BZgr300 pM`,Bluthochdruck,`BHD seit`,`BHD beh mit`,Blutdruckwerte,BDselbst,Schwanger,`Schwanger seit`,`Augensp zuletzt`,`Augensp Befund`,`Netzhaut gelasert`,`Sehminderung unbehebbar`,`Diabet Nierenschaden`,`Albumin zuletzt`,erhöht?,Dialyse,`Dialyse seit`,`andere Nierenerkrankung`,Herzkrankheit,`Angina pectoris`,Herzinfarkt,`Herzinfarkt wann`,`PTCA oder Stent`,`Bypass kardial`,`Bypass wann`,Herzschwäche,`Herzkrankheit Beschreibung`,Hirndurchblutungsstörung,Schlaganfall,Beindurchblutungsstörung,Schaufensterkrankheit,`bypaß peripher`,Geschwür,Amputation,`pAVK Beschreibung`,Ameisenlaufen,`Ameisen Ausmaß`,Druckstellen,Verformungen,`Verformungen Beschreibung`,Fußpflege,Podologie,Einlagen,`Neue Fußkomplikationen`,`Entleerungsstörungen Magen`," & _
"`Entleerungsstörungen Harnblase`,`Schwindel Aufstehen`,`Folgeerkrankungen Haut`,Bewegungseinschränkungen,Sexualstörung,`Sexualstörung seit`,`Weitere Anamnese`,Alkohol,Tabak,`Weitere Medikation`,`Liphypertrophien Abdomen`,`Liphypertrophien Beine`,`Liphypertrophien Arme`,Beinbefund,Hyperkeratosen,Ulcera,`Kraft Zehenheber`,`Kraft Zehenbeuger`,`Kraft Knie`,ASR,PSR,Oberflächensensibilität,Monofilamenttest,Kalt-Warm,`Vibration IK`,`Vibration Großzehe`,`Puls Leiste`,`Puls Kniekehle`,`Puls Atp`,`Puls Adp`,RR,RRTurboMed,Herz,Lunge,Bauch,WS,NL,SD,Carotiden,NNH,Zähne,Mundhöhle,LK,BeinödVen,`Neuro sonst`,`Weitere Befunde`,Schulung,DMP,DMSchulz,DMSchL,RRSchulz,DMPhier,HANr,HANr2,`letzte Änderung`,Diagnosen,Vorgestellt,Versicherung,AktZeit,Ther1,TherAkt,Prim,obAn1eing,obAn2eing,obAnAeing,obCheck,obBZausgew,obOSaufgek,obPodAufgek,obMBlAusgeh,obSchulaufgek,obDMPaufgekl,obMedNetz,Hausarzt,ob,QS,QT from anamnesebogen", db, adOpenStatic, adLockOptimistic
  adoPrimaryRS.Open "select * from anamnesebogen", db, adOpenDynamic, adLockOptimistic
  Me.Combo1.DataField = NachName
  Set Me.Combo1.DataSource = adoPrimaryRS
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
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case keyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub form_unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hierdurch wird die aktuelle Datensatzposition für diese Datensatzgruppe angezeigt
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

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
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Datensatz hinzufügen"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Dies wird nur für Mehrbenutzeranwendungen benötigt
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Datensatz bearbeiten"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'Zu neuem Datensatz gehen
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoPrimaryRS.MoveLast
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoPrimaryRS.MoveFirst
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

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
End Sub

