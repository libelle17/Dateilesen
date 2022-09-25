VERSION 5.00
Begin VB.Form DMPBogenauswahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Name des Dialogfeldes"
   ClientHeight    =   3285
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "DMPBogenauswahl"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox DokuDatum 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMP Asthma Ver&lauf"
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMP &Asthma neu"
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMP KHK Ve&rlauf"
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMP K&HK neu"
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMP Typ 2 V&erlauf"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMP Typ 1 &Verlauf"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMP Typ &2 neu"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMP Typ &1 neu"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label DokudatumLab 
      Caption         =   "&Doku-Datum:"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "DMPBogenauswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vater As PatListe
Private Sub CancelButton_Click()
 vater.BogArtVar = 0
 Unload Me
End Sub ' CancelButton_Click

Private Sub OKButton_Click()
 Dim i&
 For i = 0 To Option1.COUNT
  If Me.Option1(i) Then
   vater.BogArtVar = i + 1
   Exit For
  End If
 Next i
 vater.DokuDat = Me.DokuDatum
 Unload Me
End Sub ' OKButton_Click
