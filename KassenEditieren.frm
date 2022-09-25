VERSION 5.00
Begin VB.Form KassenEditieren 
   Caption         =   "Form1"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   KeyPreview      =   -1  'True
   LinkTopic       =   "KassenEditieren"
   ScaleHeight     =   10740
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Aktion 
      Caption         =   "Aktion"
      Height          =   975
      Left            =   11760
      TabIndex        =   10
      Top             =   240
      Width           =   1575
      Begin VB.OptionButton löschen 
         Caption         =   "l&öschen"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton setzen 
         Caption         =   "se&tzen"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Kassen 
      Caption         =   "Kassen:"
      Height          =   975
      Left            =   10320
      TabIndex        =   7
      Top             =   240
      Width           =   1335
      Begin VB.OptionButton alle 
         Caption         =   "a&lle"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton ausgewählte 
         Caption         =   "ausgew&ählte"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Start"
      Height          =   375
      Left            =   13440
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Anzeigen 
      Caption         =   "An&zeigen"
      Height          =   375
      Left            =   9360
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   5
      Top             =   1560
      Width           =   6015
   End
   Begin VB.TextBox ergText 
      Height          =   9975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   720
      Width           =   9135
   End
   Begin VB.TextBox sqlText 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   7455
   End
   Begin VB.CommandButton Fertig 
      Caption         =   "&Fertig"
      Height          =   375
      Left            =   14040
      TabIndex        =   6
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label sqlTextLab 
      Caption         =   "&Kassenauswahl:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "KassenEditieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SichRück$ = "Wollen Sie wirklich alle Kassen für das/die ausgewählten Insulin(e) ändern?", SRTit$ = "Sicherheitsrückfrage"
Private Sub alle_Click()
 Dim erg&
 erg = MsgBox(SichRück, vbYesNo, SRTit)
 If erg = vbNo Then
  Me.ausgewählte = 1
 End If
End Sub

Private Sub Anzeigen_Click()
 Dim ergS As New CString, sql As New CString, i&, j&
 Dim rs As New Adodb.Recordset
 sql = "SELECT vk,ik,name,kateg"
 For i = 0 To Check1.COUNT - 1
  If Check1(i) <> 0 Then
   sql.AppVar Array(",", Check1(i).Tag)
  End If
 Next i
 sql.Append " FROM `kassenliste` k WHERE name LIKE '%" & Me.sqlText & "%'"
 myFrag rs, sql.Value
 TabAusgeb rs, Me, True
 Me.ausgewählte = 1
End Sub ' Anzeigen_Click()

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 27
   Unload Me
 End Select
End Sub

Private Sub Form_Load()
 Me.Caption = "Kassen editieren (Rabattverträge, Mehrwertverträge)"
 Me.ausgewählte = 1
 Me.setzen = 1
End Sub

Private Sub OK_Click()
 Unload Me
End Sub ' OK_Click

Function Ausgeb(txt$, Optional obWahr)
 Me.ergText = txt
End Function ' Ausgeb(txt$, Optional obWahr)

Private Sub sqlText_GotFocus()
 Me.sqlText.SelStart = 0
 Me.sqlText.SelLength = Len(Me.sqlText)
End Sub ' sqlText_GotFocus()

Private Sub Start_Click()
 Dim ergS As New CString, sql As New CString, i&, j&, erg&
 Dim rs As New Adodb.Recordset
 sql = "UPDATE `kassenliste` k SET geaen=" & Format(Now(), "yyyymmddHHMMSS") & ","
 For i = 0 To Check1.COUNT - 1
  If Check1(i) <> 0 Then
   sql.AppVar Array(Check1(i).Tag, " = " & IIf(Me.setzen = 0, "0", "1") & ",")
  End If
 Next i
 If sql.Right(1) = "," Then sql.Cut (sql.length - 1)
 If Me.alle <> 0 Then
   erg = MsgBox(SichRück, vbYesNo, SRTit)
   If erg = vbNo Then Me.ausgewählte = 1
 End If
 If Me.alle = 0 Then
  sql.Append " WHERE name LIKE '%" & Me.sqlText & "%'"
 End If
 DBCn.Execute sql
 Anzeigen_Click
 Me.sqlText.SetFocus
End Sub ' Start_Click()
