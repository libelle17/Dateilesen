VERSION 5.00
Begin VB.Form MOSuch 
   Caption         =   "MOSuch"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2160
      Width           =   1575
   End
   Begin VB.OptionButton Istgleich 
      Caption         =   "&IstGleich"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Rlike"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Pat_ID 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Start"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox SuchString 
      Height          =   405
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   5295
   End
   Begin VB.ComboBox MOServer 
      Height          =   315
      ItemData        =   "MOSuch.frx":0000
      Left            =   1440
      List            =   "MOSuch.frx":000A
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Suchart 
      Caption         =   "&Suchart"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Pat_IDLabel 
      Caption         =   "&Pat_ID"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label SuchstringLabel 
      Caption         =   "S&uchstring:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label MOServerCapture 
      Caption         =   "&MOServer"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "MOSuch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Abbruch_Click()
 Unload Me
End Sub ' Abbruch_Click()

Private Sub Start_Click()
 If IsNumeric(Pat_ID) Then
  Call suchfi(Pat_ID, Me.SuchString, IIf(Me.MOServer = "szn4", True, False))
 Else ' IsNumeric(Pat_ID) Then else
  Call suchal(Me.SuchString, IIf(Me.Istgleich, True, False), IIf(Me.MOServer = "szn4", True, False))
 End If ' IsNumeric(Pat_ID) Then
End Sub ' Start_Click()
