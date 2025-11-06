VERSION 5.00
Begin VB.Form Übertragungsoptionen 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Übertragung von MO nach Diabdachau: Optionen"
   ClientHeight    =   2955
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox vorDatum 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton alleVor 
      Caption         =   "&alle vor:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton zuletzt 
      Caption         =   "&zuletzt übertragenen Pat. raussuchen"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Tage 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1060
      Width           =   855
   End
   Begin VB.CheckBox mitLabor 
      Caption         =   "&mit Labor, falls nicht aktuell oder erzwinge"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CheckBox erzwinge 
      Caption         =   "&erzwinge Übertragung, auch wenn schon aktuell"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   3855
   End
   Begin VB.OptionButton diesenundalledarunter 
      Caption         =   "&diesen und alle darunter."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.OptionButton nurdiesen 
      Caption         =   "&Nur diesen Patienten"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Pat_id 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label PatName 
      Height          =   260
      Left            =   120
      TabIndex        =   13
      Top             =   440
      Width           =   3375
   End
   Begin VB.Label Tagezurück 
      Caption         =   "Tage zurück"
      Height          =   255
      Left            =   3440
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Pat_id_Label 
      Caption         =   "&Patientennummer:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
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
Public Abbruch%

Private Sub Form_Activate()
 Abbruch = 0
 Me.Pat_ID.SetFocus
 Pat_ID.SelStart = 0
 Pat_ID.SelLength = Len(Pat_ID.Text)
End Sub

Private Sub Form_Load()
 If Me.Pat_ID = "" Then Me.Pat_ID = 99999
 If Me.Tage = "" Then Me.Tage = 99999
 Me.nurdiesen(0) = 1
 Me.mitLabor = 1
End Sub ' Form_Load()

Private Sub OKButton_Click()
 Me.Hide
End Sub ' OKButton_Click()

Private Sub CancelButton_Click()
 Abbruch = True
 Me.Hide
End Sub ' CancelButton_Click()

Private Sub Pat_id_Change()
 Call MOConInit(, "Pat_id_change()")
 If IsNumeric(Me.Pat_ID) Then
  Me.PatName = myEFrag("SELECT COALESCE((SELECT " & gesnamegmo & " FROM patstamm WHERE FSurogat=" & Me.Pat_ID & "),'')", , MOCon).Fields(0)
 End If
End Sub ' Pat_id_Change()

Private Sub Pat_ID_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub

Private Sub nurdiesen_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub

Private Sub Tage_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub

Private Sub diesenundalledarunter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub

Private Sub erzwinge_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub

Private Sub mitLabor_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub

Private Sub OKButton_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub

Private Sub CancelButton_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub

Private Sub zuletzt_Click()
 On Error Resume Next
 Me.Pat_ID = myEFrag("SELECT pat_id FROM namen ORDER BY aktzeit DESC LIMIT 1").Fields(0)
End Sub
