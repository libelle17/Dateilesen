VERSION 5.00
Begin VB.Form Sonderpatientauswahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Sonderpatient-Auswahl"
   ClientHeight    =   8910
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   8640
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Sonderpatientauswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public aufRufer As Lese



Private Sub Form_Load()
 Dim rs As New Adodb.Recordset
 myFrag rs, "SELECT * FROM namen n WHERE nachname LIKE 'zutun%' OR straße LIKE 'mittermayer%13%'"
 If Not rs.BOF Then
  Do While Not rs.EOF
   Me.List1.AddItem rs!Nachname & ", " & rs!Vorname & "  (" & rs!Pat_id & ")"
   rs.MoveNext
  Loop
 End If
End Sub

Private Sub OKButton_Click()
 Dim spneu$(), spnk$
 SplitNeu Me.List1.Text, "(", spneu
 If UBound(spneu) > 0 Then
  spnk = REPLACE$(spneu(1), ")", vNS)
  If IsNumeric(spnk) Then
   aufRufer.SpPat_id = CLng(spnk)
   aufRufer.SpName = spneu(0)
  End If
 End If
 Me.Visible = False
End Sub ' OKButton_Click()
Private Sub CancelButton_Click()
 Me.Visible = False
End Sub

