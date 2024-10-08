VERSION 5.00
Begin VB.Form MOSuch 
   Caption         =   "MOSuch"
   ClientHeight    =   2340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton Istgleich 
      Caption         =   "&IstGleich"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   1440
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
      Top             =   1800
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
   Begin VB.Label Status 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Suchart 
      Caption         =   "S&uchart"
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
Option Explicit
Dim cR As New Registry
Const Pat_ID_Schl$ = "Pat_ID"
Const MOServer_Schl$ = "MOServer"
Const SuchString_Schl$ = "SuchString"
Const Option1_Schl$ = "Option1"
Const IstGleich_Schl$ = "IstGleich"
Const RegWurzel$ = "Software\GSProducts\"
Dim RegPos$

Private Sub Form_Load()
 RegPos = RegWurzel & App.EXEName & "\" & Me.name
 Pat_ID = cR.ReadKey(Pat_ID_Schl, RegPos, HKEY_CURRENT_USER)
 MOServer = cR.ReadKey(MOServer_Schl, RegPos, HKEY_CURRENT_USER)
 SuchString = cR.ReadKey(SuchString_Schl, RegPos, HKEY_CURRENT_USER)
 Option1 = cR.ReadKey(Option1_Schl, RegPos, HKEY_CURRENT_USER)
 Istgleich = cR.ReadKey(IstGleich_Schl, RegPos, HKEY_CURRENT_USER)
End Sub ' Form_Load()

Private Sub Form_Unload(Cancel As Integer)
 cR.WriteKey Pat_ID, Pat_ID_Schl, RegPos, HKEY_CURRENT_USER, REG_SZ
 cR.WriteKey MOServer, MOServer_Schl, RegPos, HKEY_CURRENT_USER, REG_SZ
 cR.WriteKey SuchString, SuchString_Schl, RegPos, HKEY_CURRENT_USER, REG_SZ
 cR.WriteKey Option1, Option1_Schl, RegPos, HKEY_CURRENT_USER, REG_SZ
 cR.WriteKey Istgleich, IstGleich_Schl, RegPos, HKEY_CURRENT_USER, REG_SZ
End Sub ' Form_Unload(Cancel As Integer)

Private Sub Option1_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub ' Option1_KeyDown(keycode As Integer, Shift As Integer)

Private Sub IstGleich_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub ' IstGleich_KeyDown(keycode As Integer, Shift As Integer)

Private Sub Abbruch_Click()
 Unload Me
End Sub ' Abbruch_Click()

Private Sub KeyDown(keycode%)
 If keycode = 27 Then
  Call Abbruch_Click
 End If
End Sub

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub

Private Sub MOServer_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub

Private Sub Pat_ID_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub


Private Sub SuchString_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub

Private Sub SuchArt_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub

Private Sub Start_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub

Private Sub Abbruch_KeyDown(keycode As Integer, Shift As Integer)
 KeyDown keycode
End Sub

Private Sub Start_Click()
 Me.Status = "Warte ..."
 Me.Refresh
 If IsNumeric(Pat_ID) Then
  Call suchfi(Pat_ID, Me.SuchString, IIf(Me.Istgleich, True, False), Me.MOServer)
 Else ' IsNumeric(Pat_ID) Then else
  Call suchal(Me.SuchString, IIf(Me.Istgleich, True, False), Me.MOServer)
 End If ' IsNumeric(Pat_ID) Then
 Me.Status = "Fertig"
End Sub ' Start_Click()

