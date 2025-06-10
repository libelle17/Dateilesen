VERSION 5.00
Begin VB.Form Übertragungsoptionen 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Übertragung von MO nach Diabdachau: Optionen"
   ClientHeight    =   2460
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check2 
      Caption         =   "&mit Labor"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&erzwinge Übertragung"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&diesen und alle darunter"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Nur diesen Patienten"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Pat_id 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Pat_id_Label 
      Caption         =   "&Patientennummer:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Übertragungsoptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
