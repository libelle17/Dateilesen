VERSION 5.00
Begin VB.Form LANRauswahl 
   Caption         =   "Bitte wählen Sie die LANR aus"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   LinkTopic       =   "LANRauswahl"
   ScaleHeight     =   3090
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Ok 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label LDMPLbl 
      Caption         =   ".. DMP-Doku"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LFallLbl 
      Caption         =   "... Fall"
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LKartLbl 
      Caption         =   "Letzte(r): ... Karteieintrag"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Option1Lbl 
      Caption         =   "LANR und Arztname"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label LDMP 
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label LFall 
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label LKart 
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "LANRauswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Lanr&

' aufgerufen in doMachDMPBogen
Public Sub PrepPatid(Pat_id&)
 Dim i&, maxTag$
 Dim rfal As New ADODB.Recordset
 Dim rnam As New ADODB.Recordset
 myFrag rnam, "SELECT nachname, vorname, titel, nvorsatz FROM `namen` WHERE pat_id =  " & Pat_id
 If Not rnam.BOF Then
  Me.Caption = "Bitte LANR für " & rnam!Titel & " " & rnam!Vorname & " " & rnam!NVorsatz & " " & rnam!Nachname & " (Pat_id: " & Pat_id & ") auswählen!"
 End If
 For i = 0 To Me.LKart.COUNT - 1
  myFrag rfal, "SELECT zeitpunkt FROM `eintraege` e LEFT JOIN `eintragzulanr` ez ON e.art = ez.art LEFT JOIN `lanrpraxis` l ON ez.lanrid = l.id WHERE pat_id = " & Pat_id & " AND lanr = " & Me.Option1(i).Tag & " ORDER BY zeitpunkt DESC"
  If Not rfal.BOF Then
   LKart(i).Caption = ZQuart(rfal!Zeitpunkt)
   LKart(i).Tag = ZQSort(rfal!Zeitpunkt)
   If LKart(i).Tag > maxTag Then maxTag = LKart(i).Tag
  End If
  Set rfal = Nothing
 Next i
 For i = 0 To Me.LKart.COUNT - 1 ' Vorauswahl mangels besseren Wissens nach Fall richten
  If LKart(i).Tag = maxTag Then
   Option1(i) = True
   Exit For
  End If
 Next i
 For i = 0 To Me.LFall.COUNT - 1
  myFrag rfal, "SELECT lanr, quartal, bhfb FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_id & " AND lanr = " & Me.Option1(i).Tag & " ORDER BY bhfb DESC"
  If Not rfal.BOF Then
   LFall(i).Caption = rfal!Quartal
   LFall(i).Tag = ZQSort(rfal!BhFB)
  End If
  Set rfal = Nothing
 Next i
 For i = 0 To Me.LDMP.COUNT - 1
  myFrag rfal, "SELECT DokuDatum FROM `dmpreihe` dr LEFT JOIN `lanrpraxis` l ON dr.lanrid = l.id WHERE pat_id = " & Pat_id & " AND lanr = " & Me.Option1(i).Tag & " AND dr.abk RLIKE 'DM[12]|DTYP[12]'  AND DokuDatum <> 0 AND DokuDatum > '1899-12-30' ORDER BY DokuDatum DESC"
  If Not rfal.BOF Then
   LDMP(i).Caption = ZQuart(rfal!DokuDatum)
   LDMP(i).Tag = ZQSort(rfal!DokuDatum)
  End If
  Set rfal = Nothing
 Next i
End Sub ' PrepPatid(Pat_id&)

Private Sub Form_Load()
 Dim rlan As New ADODB.Recordset, sql$, i&
 sql = "SELECT a.lanr, a.nachname, a.vorname FROM `lanrpraxis` l LEFT JOIN `haerzte`.`arzt` a ON l.lanr=a.lanr WHERE NOT ISNULL(a.lanr) GROUP BY l.lanr ORDER BY l.lanr"
' myFrag rlan, "SELECT COUNT(0) zl FROM (" & sql & ")"
 myFrag rlan, sql
 If Not rlan.BOF Then
  i = 0
  Do While Not rlan.EOF
   On Error Resume Next
   Load Me.Option1(i)
   Me.Option1(i).Top = Me.Option1(i - 1).Top + 300
   Me.Option1(i).Left = Me.Option1(i - 1).Left
   Me.Option1(i).Visible = True
   Me.Option1(i).TabIndex = Me.Option1(i - 1).TabIndex + 1
   Load Me.LKart(i)
   Me.LKart(i).Top = Me.LKart(i - 1).Top + 300
   Me.LKart(i).Left = Me.LKart(i - 1).Left
   Me.LKart(i).Visible = True
   Me.LKart(i).Enabled = False
   Load Me.LFall(i)
   Me.LFall(i).Top = Me.LFall(i - 1).Top + 300
   Me.LFall(i).Left = Me.LFall(i - 1).Left
   Me.LFall(i).Visible = True
   Me.LFall(i).Enabled = False
   Load Me.LDMP(i)
   Me.LDMP(i).Top = Me.LDMP(i - 1).Top + 300
   Me.LDMP(i).Left = Me.LDMP(i - 1).Left
   Me.LDMP(i).Visible = True
   Me.LDMP(i).Enabled = False
   
   On Error GoTo fehler
   Me.Option1(i).Caption = rlan!Lanr & " &" & rlan!Vorname & " " & rlan!Nachname
   Me.Option1(i).Tag = rlan!Lanr
   i = i + 1
   rlan.MoveNext
  Loop
 End If
 Me.Height = MIN(Me.Option1.COUNT * 300& + 2000, 3675)
 Me.OK.Top = Me.Height - 1000
 Me.Abbruch.Top = Me.Height - 1000
 
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Form_Load/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Form_Load

Private Sub OK_Click()
 Dim i&
 For i = 0 To Option1.COUNT - 1
  If Option1(i).Value <> 0 Then
   Lanr = Option1(i).Tag
   Exit For
  End If
 Next i
 Unload Me
End Sub ' OK_Click

Private Sub Abbruch_Click()
 Lanr = 0
 Unload Me
End Sub

