VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PatListe 
   Caption         =   "Form1"
   ClientHeight    =   10425
   ClientLeft      =   510
   ClientTop       =   1245
   ClientWidth     =   18255
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   18255
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG 
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   17595
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "PatListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public hlese As Lese
Const MTBeg% = 1
Const MTMax% = 1
Dim MFGTyp%
Dim altFarbe&(MTBeg To MTMax)
Dim noenter%
Dim fgespei%(MTBeg To MTMax) ' Farbe gespeichert
Private Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal ms&)

Public Sub Pause(Millisekunden As Long)
sleep Millisekunden
End Sub

Private Sub Command1_Click()
 Const obstumm% = 0
 Dim hnd&, VorDoku$, Pat_id&, dtyp%
 Dim dt As dmptyp, j%
 Dim rTyp As New ADODB.Recordset
 MFG.col = 7
 VorDoku = MFG.Text
 MFG.col = 2
 Pat_id = MFG.Text
 Call DMPAusgeb(dt, Pat_id, Not obstumm)
 hnd = FensterHandle("TurboMed")
 If hnd <> 0 Then
  AppActivate "TurboMed", True
  Pause (100)
  SendKeys "{ESC}", True
  Pause (100)
  SendKeys "{ESC}", True
  Pause (100)
  SendKeys "{ESC}", True
  Pause (100)
'  SendKeys "{ESC}", True
'  Pause (100)
'  SendKeys "{ESC}", True
'  Pause (100)
  SendKeys "p", True
  Pause (100)
  SendKeys "" & "{bs}" & Pat_id & "", True
  Pause (100)
  SendKeys "{ENTER}", True
  Pause (1000)
'  SendKeys "{ENTER}", True
'  Pause (100)
'  SendKeys "{ENTER}", True
'  Pause (100)
  Dim SendStr$
  rTyp.Open "select icd from diagnosen where pat_id = " & Pat_id & " and diagsicherheit <> 'A' and icd like 'E1%' order by diagdatum desc", DBCn, adOpenStatic, adLockReadOnly
  If Not rTyp.EOF Then dtyp = Mid(rTyp.Fields(0), 3, 1) + 1
  If dtyp = 2 Then
   If VorDoku = "" Then
    SendStr = "%{F2}{TAB}+{TAB}"
    SendStr = SendStr & dt.daseit & IIf(dt.dspsy, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(dt.dspmed, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & "{TAB} {TAB 3} {TAB 2} {TAB 2}"
    For j = 0 To UBound(dt.FEn)
     SendStr = SendStr & IIf(dt.FEn(j), " ", "") & "{TAB}"
    Next j
    SendStr = SendStr & IIf(dt.tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(dt.kgr < 100, "0", "") & dt.kgr
    SendStr = SendStr & IIf(dt.gewi < 10, "00", IIf(dt.gewi < 100, "0", "")) & dt.gewi
    SendStr = SendStr & " {TAB 2}" ' altersgemäße körperliche Entwicklung
    SendStr = SendStr & IIf(dt.puls = ndok, " ", "") & "{TAB}" & IIf(dt.puls = unauff, " ", "") & "{TAB}" & IIf(dt.puls = auff, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(dt.sens = ndok, " ", "") & "{TAB}" & IIf(dt.sens = unauff, " ", "") & "{TAB}" & IIf(dt.sens = auff, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(dt.fußst = ndok, " ", "") & "{TAB}" & IIf(dt.fußst = unauff, " ", "") & "{TAB}" & IIf(dt.fußst = auff, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(dt.mSei = re Or dt.mSei = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(dt.mSei = li Or dt.mSei = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "0", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "1", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "2", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "3", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "4", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "5", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Right(dt.mWA, 1) = "A", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Right(dt.mWA, 1) = "B", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Right(dt.mWA, 1) = "C", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Right(dt.mWA, 1) = "D", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(dt.oap = re Or dt.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(dt.oap = li Or dt.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & "{TAB 8}"
    SendStr = SendStr & Format(dt.bekHb * 10, "000") & "062"
    SendStr = SendStr & "000" & Format(dt.Crea * 10, "000") & "{TAB}"
    SendStr = SendStr & IIf(dt.mau = ndok, " ", "") & "{TAB}" & IIf(dt.mau = auff, " ", "") & "{TAB}" & IIf(dt.mau = unauff, " ", "") & "{TAB 2}"
    SendStr = SendStr & Format(dt.fhilf, "00") & "0000"
    SendStr = SendStr & IIf(dt.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(dt.obIns, IIf(dt.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr = SendStr & IIf(dt.obAnal, IIf(dt.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr = SendStr & IIf(dt.obHMG, " {TAB}", "{TAB}")
    SendStr = SendStr & IIf(dt.obAntihyp, " {TAB}", "{TAB}")
    SendStr = SendStr & IIf(dt.obThro, " {TAB}", "{TAB}")
    SendStr = SendStr & "{TAB 2} {TAB}" & "{TAB 2} {TAB}"
    SendStr = SendStr & IIf(dt.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr = SendStr & IIf(dt.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
#If False Then
#End If
    SendKeys SendStr, True
    Pause (100)
   Else
    SendKeys "%{F3}{TAB}+{TAB}", True
    Pause (100)
   End If
  Else
   If VorDoku = "" Then
    SendStr = "%(MO){TAB}+{TAB}" & dt.daseit & " {TAB 2}" & IIf(dt.dspsy, " {TAB 2}", "{TAB} {TAB}") & "{TAB} {TAB}"
    For j = 1 To UBound(dt.FEn)
     SendStr = SendStr & IIf(dt.FEn(j), " ", "") & "{TAB}"
    Next j
    SendStr = SendStr & IIf(dt.tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(dt.kgr < 100, "0", "") & dt.kgr
    SendStr = SendStr & IIf(dt.gewi < 10, "00", IIf(dt.gewi < 100, "0", "")) & dt.gewi
    SendStr = SendStr & " {TAB 2}" ' altersgemäße körperliche Entwicklung
    SendStr = SendStr & IIf(dt.puls = ndok, " ", "") & "{TAB}" & IIf(dt.puls = unauff, " ", "") & "{TAB}" & IIf(dt.puls = auff, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(dt.sens = ndok, " ", "") & "{TAB}" & IIf(dt.sens = unauff, " ", "") & "{TAB}" & IIf(dt.sens = auff, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(dt.fußst = ndok, " ", "") & "{TAB}" & IIf(dt.fußst = unauff, " ", "") & "{TAB}" & IIf(dt.fußst = auff, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(dt.mSei = re Or dt.mSei = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(dt.mSei = li Or dt.mSei = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "0", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "1", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "2", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "3", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "4", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Left(dt.mWA, 1) = "5", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Right(dt.mWA, 1) = "A", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Right(dt.mWA, 1) = "B", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Right(dt.mWA, 1) = "C", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(Right(dt.mWA, 1) = "D", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(dt.oap = re Or dt.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(dt.oap = li Or dt.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & "{TAB 8}"
    SendStr = SendStr & Format(dt.bekHb * 10, "000") & "062"
    SendStr = SendStr & "000" & Format(dt.Crea * 10, "000") & "{TAB}"
    SendStr = SendStr & IIf(dt.mau = ndok, " ", "") & "{TAB}" & IIf(dt.mau = auff, " ", "") & "{TAB}" & IIf(dt.mau = unauff, " ", "") & "{TAB 2}"
    SendStr = SendStr & Format(dt.fhilf, "00") & "0000"
    SendStr = SendStr & IIf(dt.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(dt.obIns, IIf(dt.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr = SendStr & IIf(dt.obAnal, IIf(dt.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr = SendStr & IIf(dt.obHMG, " {TAB}", "{TAB}")
    SendStr = SendStr & IIf(dt.obAntihyp, " {TAB}", "{TAB}")
    SendStr = SendStr & IIf(dt.obThro, " {TAB}", "{TAB}")
    SendStr = SendStr & "{TAB 2} {TAB}" & "{TAB 2} {TAB}"
    SendStr = SendStr & IIf(dt.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr = SendStr & IIf(dt.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(dt.aug = durchg, " {TAB 3}", IIf(dt.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr = SendStr & " {TAB 16} {TAB 2}"
    Dim dmpdat$, dmpdatop$
    dmpdat = Format(min(QEnd(ZQuart(Now - Verspätung)), Now), "dd.mm.yy")
    dmpdatop = Format(min(QEnd(ZQuart(Now - Verspätung)), Now), "ddmmyyyy")
    Dim DatName$, DNn$
    DatName = "p:\" & dt.NachName & " " & dt.VorName & " (PID" & dt.Pat_id & "), DMP-Formular vom " & dmpdat & ".pdf"
    DNn = Dir(DatName)
    If DNn <> "" Then Kill DatName
    DatName = Replace(Replace(DatName, "(", "{(}"), ")", "{)}")
    SendStr = SendStr & dmpdatop & "{DEL 8}" & dmpdat & "%P %(DD)aa{ENTER}" & DatName & "{ENTER}n{ENTER}{F3}{ENTER}"
    SendKeys SendStr, True
   Else
    SendStr = "%(MU){TAB}+{TAB}"
    SendStr = SendStr & dt.daseit & IIf(dt.dspsy, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(dt.dspmed, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & "{TAB} {TAB 3} {TAB 2} {TAB 2}"
    SendKeys SendStr, True
   End If
  End If
 Else
  Stop
 End If
 MFG.SetFocus
End Sub

Private Sub Form_Load()
 Dim sql$
 MFGTyp = 1
 Dim rDPat As New ADODB.Recordset, lfdnr&, cbcol&, j%
 Dim rDok As New ADODB.Recordset
 sql = "select namen.*, icd from namen left join diagnosen on namen.pat_id = diagnosen.pat_id where notiz like '%DMP hier%' and icd like 'E1%' and diagsicherheit not in ('Z','A') order by pat_id"
 rDPat.Open "select count(*) from (" & sql & ") as innen", DBCn, adOpenStatic, adLockReadOnly
 MFG.Rows = rDPat.Fields(0) + 2
 MFG.Cols = 20
 MFG.Row = 1
 Set rDPat = Nothing
 lfdnr = 0
 rDPat.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rDPat.EOF
   If obhierdmp(rDPat!notiz) Then cbcol = 10790143 Else cbcol = vbWhite
   MFG.col = 1
   MFG.Text = lfdnr
   MFG.CellBackColor = cbcol
   lfdnr = lfdnr + 1
   MFG.col = 2
   MFG.Text = rDPat!Pat_id
   MFG.CellBackColor = cbcol
   MFG.col = 3
   MFG.Text = rDPat!NachName
   MFG.CellBackColor = cbcol
   MFG.col = 4
   MFG.Text = rDPat!VorName
   MFG.CellBackColor = cbcol
   MFG.col = 5
   MFG.Text = rDPat!notiz
   MFG.CellBackColor = cbcol
   MFG.col = 6
   MFG.Text = rDPat!ICD
   MFG.CellBackColor = IIf(rDPat!ICD Like "E11*", cbcol, vbYellow)
   rDok.Open "select * from dmpreihe where pat_id = " & rDPat!Pat_id & " order by exportiert desc;", DBCn, adOpenDynamic, adLockReadOnly
   For j = MFG.col + 1 To MFG.Cols - 1
    If rDok.EOF Then Exit For
    MFG.col = j
    MFG.Text = rDok!art & " " & rDok!DokuDatum
    If rDok!art = "ED" Then MFG.CellBackColor = 16767449
    rDok.Move 1
   Next j
   Set rDok = Nothing
  MFG.Row = MFG.Row + 1
  rDPat.Move 1
 Loop
 Call SizeColumns(MFG, Me)
 MFG.col = 2
 MFG.Row = 1
 Call mfg_entercell
 Command1.Caption = "&Erstellen"
End Sub

Private Sub mfg_entercell()
 If noenter = 0 Then
  If fgespei(MFGTyp) = 0 Then
   altFarbe(MFGTyp) = Me.MFG.CellBackColor
   Me.MFG.CellBackColor = vbYellow
   fgespei(MFGTyp) = -1
  End If
 End If
End Sub

Private Sub MFG_leavecell()
 If noenter = 0 Then
  Me.MFG.CellBackColor = IIf(Me.MFG.Row = 0, vbActiveBorder, altFarbe(MFGTyp))
'  If MFG.CellBackColor = 0 Then Stop
  fgespei(MFGTyp) = 0
 End If
End Sub
Private Sub MFG_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub
Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub
Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub
Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub
Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub
Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)
 Call key(KeyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub

