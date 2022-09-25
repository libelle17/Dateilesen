VERSION 5.00
Begin VB.Form Hausarztanzeigen 
   Caption         =   "Hausarzt anzeigen"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton EOrt 
      Caption         =   "&Einzelner Ort:"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   3600
      Width           =   1455
   End
   Begin VB.OptionButton LkrMüNW 
      Caption         =   "&Landkreis München NW"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   3240
      Width           =   2295
   End
   Begin VB.OptionButton LkrSchrob 
      Caption         =   "&Landkreis Schrobenhausen"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
   End
   Begin VB.OptionButton LkrPfaff 
      Caption         =   "&Landkreis Pfaffenhofen"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
   End
   Begin VB.OptionButton LkrFreis 
      Caption         =   "&Landkreis Freising"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.OptionButton LkrFFB 
      Caption         =   "&Landkreis FFB"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.OptionButton LkrDachau 
      Caption         =   "&Landkreis Dachau"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ComboBox Fachricht 
      Height          =   315
      Left            =   1200
      TabIndex        =   14
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Felder leeren"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   12
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Frame Ortsauswahl 
      Caption         =   "Ortsauswahl"
      Height          =   2775
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "&Nachname:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "F&achrichtung:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   13
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "&Vorname:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Hausarztanzeigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public hlese As Lese

Private Sub Abbruch_Click()
 Unload Me
End Sub ' Abbruch_Click()

Private Sub Command1_Click()
 Dim i: For i = 0 To 2: Me.Text1(i) = vNS: Next i
End Sub ' Command1_Click()

Private Sub Form_KeyDown(KeyCode%, Shift%)
 Call Key(KeyCode, Shift, Me)
End Sub ' Form_KeyDown(KeyCode%, Shift%)

Private Sub Form_Load()
 Dim rs As New Adodb.Recordset
 Me.Fachricht.Clear
 myFrag rs, "SELECT fachrichtung FROM `haerzte`.fachrichtung GROUP BY fachrichtung"
 Do While Not rs.EOF
  Me.Fachricht.AddItem rs!fachrichtung
  rs.MoveNext
 Loop
End Sub ' Form_Load()

Private Sub Form_Unload(Cancel%)
' RegSpeichern
 On Error Resume Next
 Me.hlese.Show
End Sub ' Form_Unload(Cancel As Integer)


Private Sub OK_Click()
 Dim sql0$, sql1$, ostr$, sql$, rs As New Adodb.Recordset
 sql0 = "SELECT CONCAT_WS(' ',IF(a.obweibl<>0,'Frau','Herr'), t.titel, a.Namenszusatz, a.vorname, CONCAT(a.Nachname,', ',GROUP_CONCAT(DISTINCT Fachrichtung))) Name, a.LANR, bs.BSNR, CONCAT_WS(' ',bs.straße, CONCAT(bs.hausnr,','), bs.plz, ort) Adresse, Tel, Fax, Mail, DATE_FORMAT(a.Seit,'%e.%c.%Y') Seit FROM `haerzte`.arzt a LEFT JOIN `haerzte`.arzt_has_bs ahb ON a.idarzt = ahb.arzt_id LEFT JOIN `haerzte`.bs ON bs.idbs = ahb.bs_id LEFT JOIN `haerzte`.ort ON bs.ort_id = ort.idort LEFT JOIN `haerzte`.titel t ON a.titel_id = t.idtitel LEFT JOIN `haerzte`.arzt_has_fachrichtung ahf ON a.idarzt = ahf.arzt_id LEFT JOIN `haerzte`.fachrichtung fr ON ahf.fachrichtung_id = fr.idfachrichtung LEFT JOIN `haerzte`.tel ON bs.idbs = tel.bs_id LEFT JOIN `haerzte`.fax ON bs.idbs = fax.bs_id LEFT JOIN `haerzte`.mail ON bs.idbs = mail.bs_id WHERE nachname LIKE '" & Me.Text1(0) & "%' AND vorname LIKE '" & Me.Text1(1) & "%' "
 sql1 = "and fachrichtung LIKE '%" & Me.Fachricht & "%' " & _
       "GROUP BY lanr, REPLACE(straße,'str.','straße'),ort ORDER BY ort, REPLACE(straße,'str.','straße'), nachname"
'
 If Me.LkrDachau(0) <> 0 Then
  ostr = "and plz IN ('85250','85232','85221','85253','85778','85241','86567','85757','85229','85235','85238','85235','85244','85247','85254','85256','85258') "
 ElseIf Me.LkrFFB(1) <> 0 Then
  ostr = "and plz IN ('82276','82239','82278','82281','82223','82275','82256','82110','82284','82194','82285','82287','82288','82290','82216','82291','82293','82272','82294','82140','82178','82296','82299') "
 ElseIf Me.LkrFreis(0) <> 0 Then
  ostr = "and plz IN ('85391','85395','85386','85777','85354','85356','85408','85410','85399','85411','85413','85414','85402','85416','85417','85419','85368','85405','85375','85307','85406') "
 ElseIf Me.LkrPfaff(2) <> 0 Then
  ostr = "and plz IN ('85107','85119','85290','85302','85304','85305','85077','85126','85276','85309','85293','85084','85296','85298','85301','85088','85283') "
 ElseIf Me.LkrSchrob(3) <> 0 Then
  ostr = "and plz IN ('86561','86673','86562','86564','86666','86676','86565','86668','85123','86669','86571','86633','86697','86643','86701','86529','86579','86706') "
 ElseIf Me.LkrMüNW(0) <> 0 Then
  ostr = "and plz IN ('85748','82166','85764','85716') "
 ElseIf Me.EOrt(1) <> 0 Then
  ostr = "and ort LIKE '" & Me.Text1(2) & "%' "
 End If
 sql = sql0 & ostr & sql1
 DBCn.Execute "SET GROUP_CONCAT_MAX_LEN = 70"
 myFrag rs, sql
 If rs.EOF Then
  MsgBox "Nix gefunden"
 Else
  TabAusgeb rs, Me, , , , , Array(120), , "Hausarztergebnis für '" & Text1(0) & "', '" & Text1(1) & "' IN '" & Text1(2) & "'"
 End If
End Sub ' Ok_Click()

