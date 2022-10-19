VERSION 5.00
Begin VB.Form Optionen 
   Caption         =   "Optionen"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   9000
      Width           =   615
   End
   Begin VB.CheckBox KReinCB 
      Caption         =   "&Kodierrichtlinien eingeschaltet"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Optionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 And Shift = 0 Then Unload Me
End Sub ' Form_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Form_Load()
 KReinCB.Tag = "x"
 KReinCB = KRein
 KReinCB.Tag = ""
End Sub ' Form_Load()

Private Sub KReinCB_Click()
 Dim rs As Adodb.Recordset
 If KReinCB.Tag = "" Then
  myFrag rs, "INSERT INTO koricht(Datum,IP,ob) VALUES('" & Format(Now(), "YYYY-MM-DD hh:mm:ss") & "','" & CptName & "','" & KReinCB & "')"
  KRein = KReinCB
 End If
End Sub ' KReinCB_Click

Private Sub KReinCB_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 And Shift = 0 Then Unload Me
End Sub ' KReinCB_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub OK_Click()
 Unload Me
End Sub ' OK_Click()

Private Sub Ok_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 And Shift = 0 Then Unload Me
End Sub ' Ok_KeyDown(KeyCode As Integer, Shift As Integer)
