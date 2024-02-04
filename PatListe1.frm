VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PatListe 
   Caption         =   "Patientenliste"
   ClientHeight    =   10425
   ClientLeft      =   510
   ClientTop       =   1245
   ClientWidth     =   15120
   FillColor       =   &H80000005&
   Icon            =   "PatListe1.frx":0000
   LinkTopic       =   "PatListe"
   ScaleHeight     =   10425
   ScaleWidth      =   15120
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   18120
      TabIndex        =   33
      Top             =   10320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   15
      Left            =   12480
      TabIndex        =   20
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   14
      Left            =   11640
      TabIndex        =   19
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   13
      Left            =   10560
      TabIndex        =   18
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   13920
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   12
      Left            =   9480
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   11
      Left            =   8400
      TabIndex        =   14
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   10
      Left            =   7320
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   9
      Left            =   6240
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   8
      Left            =   5160
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   7
      Left            =   4440
      TabIndex        =   10
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.ListBox Li1 
      Height          =   255
      Left            =   17160
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cb1 
      Height          =   315
      Left            =   15960
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG 
      Height          =   9975
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   17595
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H00FFFF00&
      Caption         =   "and.DMP"
      Height          =   240
      Index           =   8
      Left            =   18195
      TabIndex        =   34
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label SpalteLbl 
      Caption         =   "Doku-Spalte"
      Height          =   255
      Index           =   3
      Left            =   18120
      TabIndex        =   32
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label SpalteLbl 
      Caption         =   "letzter Fall-Spalte"
      Height          =   255
      Index           =   2
      Left            =   18120
      TabIndex        =   31
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label SpalteLbl 
      Caption         =   "ICD-Spalte"
      Height          =   255
      Index           =   1
      Left            =   18120
      TabIndex        =   30
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label SpalteLbl 
      Caption         =   "Nachnamen-Spalte"
      Height          =   255
      Index           =   0
      Left            =   18120
      TabIndex        =   29
      Top             =   600
      Width           =   975
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H000000C0&
      Caption         =   "Erstdoku fehlt"
      Height          =   255
      Index           =   7
      Left            =   18120
      TabIndex        =   28
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H8000000D&
      Caption         =   "ED=Erstdoku"
      Height          =   255
      Index           =   6
      Left            =   18120
      TabIndex        =   27
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H00E0E0E0&
      Caption         =   "FD=Folgedoku"
      Height          =   255
      Index           =   5
      Left            =   18120
      TabIndex        =   26
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "zu alt"
      Height          =   255
      Index           =   4
      Left            =   18195
      TabIndex        =   25
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H0000FFFF&
      Caption         =   "Typ 1-D.m."
      Height          =   255
      Index           =   3
      Left            =   18195
      TabIndex        =   24
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cursor"
      Height          =   255
      Index           =   2
      Left            =   18195
      TabIndex        =   23
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H000000C0&
      Caption         =   "zu erstellen"
      Height          =   255
      Index           =   1
      Left            =   18195
      TabIndex        =   22
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label ErklärungLbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "tot"
      Height          =   255
      Index           =   0
      Left            =   18195
      TabIndex        =   21
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   13320
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "PatListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pRs As ADODB.Recordset ' zugewiesen in Tabausgeb, verwendet in TabAusFüll
Public hlese As Lese
Public AnfCode%
Public labxtb$
Public ohneTermine
' für LabordateiAnzeig und mfg_mousemove
' Const Pat_IDSp% = 0
' Const NamSp% = 1
' Const ParSp% = 2
' Const WertSp% = 3
' Const EinhSp% = 4
' Const VorWSp1% = 5
' Const VorWSp2% = 6
' Const NBSp% = 7
' Const MedSp% = 8
 Enum SpEnum
  Pat_IDSp
  namsp
  parsp
  wertsp
  einhsp
  vorwsp1
  vorwsp2
  nbsp
  medsp
  ficdsp
  terminsp
  labhwsp
 End Enum
 Enum DMPSpEnum
  DNrSp
  DPatIDSp
  DNamSp
 End Enum
 Const vorTage% = 5

Const lila = &HC000C0 ' als Reservefarbe

Const DunkelRosa& = &HA4A4FF
Const DunkelRot& = &HC0&
Const HellRot& = &H4040FF
Const Orange& = 33023 ' &H80FF&
Const HellOrange& = &H80C0FF ' (255,192,128)
' Const hellorange& = 8438015 ' &H80C0FF
Const MittigRosa& = &HC0C0FF ' (255,192,192)
' Const mittigrosa& = 12632319 ' &HC0C0FF
Const MittigAHRosa& = &HAAD2FF '( (255,210,170)
Const GelblichRosa& = &HC0E0FF ' (255,224,192)
Const Blau& = vbHighlight ' 0x8000000D  background colour of items selected in a control
'Const blau& = &H8000000D
Const HellBlau& = &HFFD9D9 '(217,217,255)
'Const hellblau& = 16767449 ' &HFFD9D9

Const vbHellGrau& = vb3DLight '0&80000016  Second lightest 3D colour after vb3DHighlight
' Const vbhellgrau& = &H80000016
Const vbHellGelb& = &HC0FFFF ' (255,255,192)
' Const vbhellgelb& = 12648447 ' &HC0FFFF
Const vbGelblichGrau& = vbButtonFace '  0x8000000F  colour of shading on the face of command buttons
' Const vbgelblichgrau& = &H8000000F
Const vbGräulich = &HE0E0E0 ' (224,224,224)
Const vbGelb& = &HFFFF ' (255,255,0) , gelb
' Const vbgelb& = 65535 ' &HFFFF&, gelb
Const vbMittelBlau = &HE16941 ' RGB(65, 105, 225) ' https://www.rapidtables.org/de/web/color/RGB_Color.html
' Const vbmittelblau& = 14772545 'RGB(65, 105, 225) ' http://www.am.uni-duesseldorf.de/de/Links/Tools/farbtabelle.html
Const vbWagnerBlau = &HE16941  'RGB(65,105,225)
' Const vbwagnerblau = 14772545  'RGB(65,105,225)
Const vbWagnerGrün = &H71B33C   'RGB(60,179,113)
' Const vbwagnergrün = 7451452   'RGB(60,179,113)
Const vbWagnerRot = &H6A6AFF   'RGB(255,106,106)
' Const vbwagnerrot = 6974207   'RGB(255,106,106)
Const vbWagnerAHRot = &H8282FF   'RGB(255,130,130)
Const vbMittelGrün& = &H228B22 'RGB(34, 139, 34)
' Const vbmittelgrün& = 2263842  'RGB(34, 139, 34)
Const vbHellLila& = &HD670DA  'rgb(218,112,214)
' Const vbhelllila& = 14053594  'rgb(218,112,214)
Const vbMittelLila& = &HD355BA  ' rgb(186,85,211)
' Const vbmittellila& = 13850042  ' rgb(186,85,211)
Const vbsiena& = &H80C0FF ' (255,192,128)
Const vborange& = &H80FF& ' (255,128,0)
Const vbbraun& = &H40C0 ' (192,64,0)

Const vbWeiss& = &HFFFFFF   ' Weiss
Const vbGrau& = &H808080   ' Grau
Const vbSchwarz& = &H0&       ' Schwarz
Const vbBlau& = &HFF0000   ' Blau
Const vbZyan& = &HFFFF00   ' Zyan (Türkis)
Const vbGrün& = &HFF00&    ' Grün
Const vbMagenta& = &HFF00FF   ' Magenta
Const vbRot& = &HFF&      ' Rot
Const vbDunkelrot& = &H800000  ' Dunkelblau
Const vbBlauGrün& = &H808000  ' Blaugrün
Const vbDunkelGrün& = &H8000&   ' Dunkelgrün
Const vbHellGrün& = &H80FF80  ' hellgrün
Const vbViolett& = &H800080  ' Violett
Const vbOcker& = &H8080&  ' Ocker
Const vbGoldenRod = 55295 ' RGB(255, 215, 0) ' helloliv
Const vbMittelBraun& = &HB86B8 ' RGB(184,134,11)
Const vbHellBraun& = &H20A5DA ' RGB(218,165,32)

Const MTBeg% = 1
Const MTMax% = 1
Const DMP_Import = "DMP_Import.BDT"
Public GesZl& ' Zahl der Patienten für die Gesamtzusammenstellung
Public GesColl As New Collection
Dim RegPos$
Dim MFGTyp%
Dim Cstumm%
Dim GruSp%, NachNameSp%, PIDSp%, ICDSp%, VorDokuSp%
Dim altFarbe&(MTBeg To MTMax)
Dim cRow&(MTBeg To MTMax), ccol&(MTBeg To MTMax), cRowSel&(MTBeg To MTMax) ' für Tb1, li1, Text1 und spätere Änderungen
Public obtb%
Dim noenter%, altnoenter%
Dim fgespei%(MTBeg To MTMax) ' Farbe gespeichert
Private Declare Sub Sleep Lib "kernel32" (ByVal ms&)
Dim KeyC%
Dim SuchStr$
Dim aktc&, aktr&, altr&, alttop&, altC&
Enum BogArtTyp
 typ1neu = 1
 typ2neu
 typ1alt
 typ2alt
End Enum
Enum ArtTyp
 artPat = 1 ' DMP hier Liste
 artDiag    ' Diagnosen sortieren
 artLab     ' pathologische Laborwerte anschauen
 artDMP     ' DMP-Infos an Hausärzte faxen
 artLPar    ' Laborparameter
 artTabAus  ' Tabausgeb
 artHA      ' Hausärzte (HAreal)
End Enum
Public PLArt As ArtTyp
Public Typisierung$ ' um die Zeilenhöhe für Motivationsgesprächskandidaten steuern zu können
Public BogArtVar As BogArtTyp ' nur für callMachDMPBogen
Public DokuDat As Date ' nur für callMachDMPBogen
Public LabDat$ ' nur für PathLabForm
Public HASL As New SortierListe
Public PatZuHASL As New SortierListe
Public HAS As SortierHA
Public PatZuHAS As SortierPatZuHA
Enum DMPCol
 nrcol
 zahlcol
 gewcol
 üwnrcol = 4
 FaxCol
 PlusCol
 gew2col = PlusCol + 2
End Enum
Dim altrow&, altCol&, altX&, altY&
Dim cb&, QuellZ&, IDS&()
' Konstante für Laboranzeige
Const MFGLabCols% = 9
' Zwischenspeicher für die Registry
Dim MFGCW&(), MFGTopRow&, MFGRow&, MFGLeftCol&, MFGCol&, MFGLabSort%
Public obMitZähler% ' 0 = ohne Zähler, 1 = mit Zähler
Dim obPidSp%
Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
        lpOperation As String, ByVal lpFile As String, ByVal _
        lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long


' aufgerufen in Form_Load
Private Sub TabAusFüll()
 Dim aktRow&, i&
 obPidSp = False
 aktRow = 1
 Me.MFG.Cols = pRs.Fields.COUNT + obMitZähler
 If Not pRs.BOF Then pRs.MoveFirst
 Do While Not pRs.EOF
  aktRow = aktRow + 1
  pRs.MoveNext
 Loop
 Me.MFG.Rows = aktRow
 aktRow = 1
 If Not pRs.BOF Then pRs.MoveFirst
 For i = 0 To pRs.Fields.COUNT - 1
  If obMitZähler <> 0 Then Me.MFG.TextMatrix(0, 0) = "Zähler"
  Me.MFG.TextMatrix(0, i + obMitZähler) = pRs.Fields(i).name
  If PIDSp = 0 Then If LCase$(pRs.Fields(i).name) = "pid" Or LCase$(pRs.Fields(i).name) = "pat_id" Then PIDSp = i + obMitZähler: obPidSp = True
 Next i
 
'
If Me.MFG.Cols < 6 Then Me.MFG.col = Me.MFG.Cols - 1 Else Me.MFG.col = 5
' Me.Label1.Width = Me.MFG.CellWidth
' Me.Label1.AutoSize = True
 Do While Not pRs.EOF
  Me.MFG.TextMatrix(aktRow, 0) = aktRow
  For i = 0 To pRs.Fields.COUNT - 1
   If IsNull(pRs.Fields(i)) Then
    Me.MFG.TextMatrix(aktRow, i + obMitZähler) = "Null"
   Else
    Me.MFG.TextMatrix(aktRow, i + obMitZähler) = pRs.Fields(i)
    
    If Me.Typisierung = "Motivationsgesprächskandidaten" Then
     If i = 5 Then
'     Me.Label1.Caption = pRs.Fields(i)
'     Me.MFG.RowHeight(aktRow) = Me.Label1.Height
     Me.MFG.RowHeight(aktRow) = Me.MFG.RowHeight(aktRow) * (Int(Len(Me.MFG.TextMatrix(aktRow, i + obMitZähler)) / 135) + 1)
     End If
    End If
    
   End If
  Next i
  aktRow = aktRow + 1
  pRs.MoveNext
 Loop
 MFG.WordWrap = True
 mfg_entercell
End Sub ' TabAusFüll()

' in Form_Load
Private Sub DMPFüll() ' für: Alle &DMP-Dokumente an Hausärzte faxen ' s. DMP_Dokumente_an_HA_Nachweis_Click
 Dim rs0 As New ADODB.Recordset, rs1 As New ADODB.Recordset, sql0$, sql1$, i&
 
 Static gafehler%
 On Error GoTo fehler
 FNr = 2
'' sql0 = "SELECT COUNT(0) Zahl, d0.adressat, d0.üwnnr, d0.fax FROM dmpausw d0 GROUP BY d0.üwnnr"
' sql0 = "SELECT COUNT(0) Zahl, n.getha0 ÜWNNr, h.fax, IF(h.anrede,'Herr','Frau') Anrede, CONCAT_WS(', ',h.nachname,h.vorname,LEFT(h.adressat,instr(h.adressat,h.vorname)-1)) Adressat, IF(innereallg,1,0) innereallg FROM `aktfaellev` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN `hareal` h ON n.getha0 = h.kvnr LEFT JOIN `desktop` dt ON n.pat_id = dt.pat_id AND dt.titel LIKE '%kein%bericht%' WHERE ISNULL(dt.titel) AND n.dmpklass = 2 AND f.icd RLIKE '^E1[0-4]\.' GROUP BY getha0" ' AND ((dmp1<>0 AND icd RLIKE '^E1[0234]') OR (dmp2<>0 AND icd RLIKE '^E1[1234]'))
sql0 = "SELECT COUNT(0) Zahl, i.* FROM (SELECT n.getha0 ÜWNNR, h.fax, h.Anrede, " & _
       "CONCAT_WS(', ',h.Name, h.Vorname, h.titelt) Adressat, IF(arzttyp='HA',1,0) innereallg " & _
       ", ISNULL(hk.kvnr) obdmpinfo " & _
       "FROM aktfaellev f " & _
       "LEFT JOIN namen n ON f.pat_id = n.pat_id " & _
       "LEFT JOIN liuez h ON h.kvnri=n.getha0 AND n.getha0<>0 " & _
       "LEFT JOIN `hakeinedmpinfo` hk ON n.getha0=hk.kvnr " & _
       "LEFT JOIN `desktop` dt ON n.pat_id = dt.pat_id AND dt.titel LIKE '%kein%bericht%' " & _
       "WHERE ISNULL(dt.titel) AND n.dmpklass = 2 AND f.icd RLIKE '^E1[0-4]\.' AND h.kvnr<>'' " & _
       "GROUP BY f.pat_id) i GROUP BY üwnnr ORDER BY adressat"

 myFrag rs0, sql0
 Do While Not rs0.EOF
  Set HAS = New SortierHA
  HAS.Zahl = rs0!Zahl
  HAS.fax = rs0!fax
  HAS.ÜwNm = rs0!Adressat
  HAS.ÜWNr = rs0!ÜWNNr
  HAS.obDMPInfo = rs0!obDMPInfo
  HAS.gewählt = -rs0!obDMPInfo
  HASL.sCAdd HAS
  rs0.MoveNext
 Loop
'' sql0 = "SELECT * FROM dmpausw"
' sql0 = "SELECT f.pat_id, CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')), IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')), n.nachname,',',n.vorname) Name, f.icd, n.getha0 ÜWNNr, h.Adressat FROM `aktfaellev` f LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN `hareal` h ON n.getha0 = h.kvnr LEFT JOIN `desktop` dt ON n.pat_id = dt.pat_id AND dt.titel LIKE '%kein%bericht%' WHERE ISNULL(dt.titel) AND n.dmpklass = 2 AND f.icd RLIKE '^E1[0-4]\.'" ' AND ((dmp1<>0 AND icd RLIKE '^E1[0234]') OR (dmp2<>0 AND icd RLIKE '^E1[1234]'))
 ''(SELECT pat_id FROM desktop dt WHERE pat_id = n.pat_id AND dt.titel LIKE '%kein%Bericht%') kb,
 sql0 = "SELECT f.pat_id, gesname(f.pat_id) name, " & vbCrLf & _
       "f.icd, n.getha0 Üwnnr, CONCAT_WS(', ',h.name, h.vorname, h.titelt) Adressat " & vbCrLf & _
       ", ISNULL(hk.kvnr) obdmpinfo " & vbCrLf & _
       "FROM aktfaellev f " & vbCrLf & _
       "LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
       "LEFT JOIN liuez h ON h.kvnri=n.getha0 AND n.getha0<>0 " & vbCrLf & _
       "LEFT JOIN `hakeinedmpinfo` hk ON n.getha0=hk.kvnr " & vbCrLf & _
       "LEFT JOIN `desktop` dt ON n.pat_id = dt.pat_id AND dt.titel LIKE '%kein%bericht%' " & vbCrLf & _
       "WHERE ISNULL(dt.titel) AND n.dmpklass = 2 AND f.icd RLIKE '^E1[0-4]\.' AND h.kvnr<>'' " & vbCrLf & _
       "GROUP BY f.pat_id;"
 Set rs0 = Nothing
 myFrag rs0, sql0
 Do While Not rs0.EOF
'  sql1 = "SELECT pat_id FROM `desktop` dt WHERE pat_id = " & rs0!Pat_id & " AND dt.titel LIKE '%kein%Bericht%'"
'  SET rs1 = Nothing
'  myFrag rs1, sql1
'  IF rs1.BOF THEN
   Set PatZuHAS = New SortierPatZuHA
   PatZuHAS.obDMPInfo = -rs0!obDMPInfo
   PatZuHAS.gewählt = -rs0!obDMPInfo
   PatZuHAS.name = rs0!name
   PatZuHAS.Pat_id = rs0!Pat_id
   PatZuHAS.ÜwNm = rs0!Adressat
   PatZuHAS.ÜWNr = rs0!ÜWNNr
   PatZuHASL.sCAdd PatZuHAS
'  END IF
  rs0.MoveNext
 Loop
 Me.MFG.Cols = PlusCol + 5
 Me.MFG.Rows = HASL.COUNT + 1 '  IF Me.MFG.Rows <= i THEN Me.MFG.Rows = Me.MFG.Rows + 1000
 For i = 1 To HASL.COUNT
  Set HAS = HASL.Item(i)
  Me.MFG.TextMatrix(i, nrcol) = i
  Me.MFG.TextMatrix(i, zahlcol) = HAS.Zahl
  If HAS.gewählt = -1 Then Me.MFG.TextMatrix(i, gewcol) = "X" Else If HAS.gewählt = -2 Then Me.MFG.TextMatrix(i, gewcol) = "(X)"
  Me.MFG.TextMatrix(i, 3) = HAS.ÜwNm
  Me.MFG.TextMatrix(i, üwnrcol) = HAS.ÜWNr
  Me.MFG.TextMatrix(i, FaxCol) = HAS.fax
  Me.MFG.TextMatrix(i, PlusCol) = "+"
'  Me.MFG.TextMatrix(i, 5) = HAS.Fax
 Next i
 Exit Sub
fehler:
If (Err.Number = -2147467259 Or Err.Number = -2147217887) And gafehler < 5 Then ' Server has gone away / ungültiger Eigenschaftswert
 Call DBCnOpen
 gafehler = gafehler + 1
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DMPFüll/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub 'DMPFüll

'Public Sub Pause(Millisekunden&)
'Sleep Millisekunden
'End Sub ' Pause(Millisekunden AS Long)

Private Function getBogArtTyp() As BogArtTyp
 FNr = 3
 With MFG
'  altC = .col
 If .TextMatrix(.Row, VorDokuSp) = "" Then ' Vorbefunde
  If .TextMatrix(.Row, ICDSp) Like "E10*" Then
   getBogArtTyp = typ1neu
  Else
   getBogArtTyp = typ2neu
  End If
 Else
  If .TextMatrix(.Row, ICDSp) Like "E10*" Then
   getBogArtTyp = typ1alt
  Else
   getBogArtTyp = typ2alt
  End If
 End If
'.col = altC
 End With
End Function ' getBogArtTyp() AS BogArtTyp

Public Sub TopAusricht()
  Dim i%, firstvis%
  FNr = 4
  With MFG
    On Error Resume Next
    If .Row - altr > (.Height / .CellHeight * 0.9) Then
     .TopRow = .Row - 1
    End If
  End With
  For i = Me.Command1.COUNT - 1 To 0 Step -1
   If Me.Command1(i).Visible Then
    firstvis = i
    Exit For
   End If
  Next i
  Me.Label1.Left = Me.Command1(firstvis).Left + Me.Command1(firstvis).Width + 100
'  Me.Text1.Left = Me.Label1.Left + MAX(Me.Label1.Width, Len(Me.Label1) * 100) + 50
  Me.Text1.Left = Me.Label1.Left + Me.Label1.Width + 50
End Sub ' TopAusricht

Public Sub verschieb(Ziel$)
 Call Shell("cmd.exe /c move """ & Me.LabDat & """ """ & Ziel & "")
End Sub ' verschieb(Ziel$)

Private Sub lösch()
 Dim erg&
 erg = MsgBox("Wollen Sie wirklich die Datei:" & vbCrLf & Me.LabDat & vbCrLf & "löschen?", vbDefaultButton2 + vbYesNo)
 If erg = vbYes Then
  Kill Me.LabDat
 End If
End Sub ' lösch()

Sub doStart() 'DMP-Infos an Hausärzte faxen starten
   Dim aktPatGefaxt$(), anfang%, docName$, dszahl&, i&, j&, gefunden%, ÜwNm$
   ReDim aktPatGefaxt(0)
   dszahl = 0
    For i = 1 To PatZuHASL.COUNT
     If PatZuHASL.Item(i).gewählt Then
      Dim fax1$, obdoppelt%
'      HASL.SuchItem PatZuHASL.Item(i).ÜWNr ' falsch sortiert
      gefunden = False
      For j = 1 To HASL.COUNT
       If HASL.Item(j).ÜWNr = PatZuHASL.Item(i).ÜWNr Then
        gefunden = True
        Exit For
       End If
      Next j
      If gefunden Then
         fax1 = HASL.Item(j).fax
         ÜwNm = HASL.Item(j).ÜwNm
         obdoppelt = False
'         For j = 0 To UBound(aktPatGefaxt) - 1
'          IF aktPatGefaxt(j) = fax1 THEN obdoppelt = True
'         Next j
         If Not obdoppelt Then
          syscmd 4, "Erstelle: " & " " & PatZuHASL.Item(i).Pat_id & " " & PatZuHASL.Item(i).name & " " & fax1 & " " & PatZuHASL.Item(i).ÜWNr
'          IF runde = 1 THEN
'          Else
'           IF pat_id = 1974 THEN
           anfang = True
'           END IF
           If anfang Then
            Call Ausgeb(PatZuHASL.Item(i).Pat_id & ": " & PatZuHASL.Item(i).name & ", " & fax1 & ", " & PatZuHASL.Item(i).ÜWNr)
            DoEvents
            docName = do_DMPAusgebStandAlone(PatZuHASL.Item(i).Pat_id, fax1, ÜwNm)
           End If ' 1 = 0
           aktPatGefaxt(UBound(aktPatGefaxt)) = fax1
           ReDim Preserve aktPatGefaxt(UBound(aktPatGefaxt) + 1)
'         Exit For ' Alternative zu obdoppelt
'          END IF
          dszahl = dszahl + 1
          Me.Label1 = dszahl
          Me.Label1.Visible = True
          DoEvents
         End If ' runde = 1
       End If ' gefunden
      End If ' gewählt
     Next i
     MsgBox "Fertig mit Erstellen der Faxe für alle markierten DMP-Patieten"
End Sub ' doStart

Public Sub Command1_Click(Index As Integer)
 FNr = 5
 Select Case PLArt
  Case artLPar
   Select Case Index
    Case 0
     Me.MFG.TopRow = MAXvb(Me.MFG.TopRow - 30, 1)
    Case 1
     Me.MFG.TopRow = MINvb(Me.MFG.TopRow + 30, Me.MFG.Rows)
    Case 2
     Unload Me
   End Select
   
  Case artDMP
   Select Case Index
    Case 0, 1, 4
     AlleMark Index - 1
    Case 2
     Call doStart
    Case 3
     Form_KeyDown 27, 0
   End Select
   
  Case artTabAus
   Select Case Index
    Case 2 - 2 * (obPidSp + 1): Call Nächster
    Case 3 - 2 * (obPidSp + 1): Call Voriger
    Case 0: Call FertigStellen(Me.MFG.Row, True) ' in Turbomed anzeigen
    Case 1: Call dodoPLZ(Me.MFG.TextMatrix(Me.MFG.Row, PIDSp), plzVz, Now, Now - Int(Now), True)  ' Patientenlaufzettel erstellen
   End Select
   
  Case artLab
   Select Case Index
    Case 0: Call verschieb(pVerz & "KothnyLabor")
    Case 1: Call verschieb(pVerz & "SchadeLabor")
    Case 2: Call verschieb(pVerz & "LaborAlt")
    Case 3: Call lösch
    Case 4: Call FertigStellen(Me.MFG.Row, True) ' in Turbomed anzeigen
    Case 5: Call dodoPLZ(Me.MFG.TextMatrix(Me.MFG.Row, Pat_IDSp), plzVz, Now, Now - Int(Now), True) ' Patientenlaufzettel erstellen
    Case 6: Call farberklärung ' hier Farberklärung
   End Select
   
  Case artPat ' DMP hier Liste
   Select Case Index
    Case 0: Call dokuErstelle
    Case 1: Call FertigStellen(Me.MFG.Row, True) ' in TurbomedAnzeigen
    Case 2: Call FertigStellen(Me.MFG.Row)
    Case 3: Call Suchen
    Case 4: Call Suchen(weiter:=True)
    Case 5: Call NächsterGleichartiger
  '  Case 6: Call Typ1ED
  '  Case 7: Call Typ1FD
  '  Case 8: Call Typ2ED
  '  Case 9: Call Typ2FD
    Case 6: Call DokuBeliebig
    Case 7: Call FertigStellenBeliebig
    Case 8: Call plz
    Case 9: Call htmlb
    Case 10: Call zurück
    Case 11: Call Auffrisch
    Case 12: Call alleFS
    Case 13: Call aktFS
    Case 14: Call GesZF  ' Gesamtzusammenstellung-Funktion
   End Select
  Case artLPar
  Case artHA
  Case artDiag
   If Index = 3 Then Call Suchen(alleSp:=True) Else If Index = 4 Then Call Suchen(weiter:=True, alleSp:=True)
 End Select
End Sub ' Command1_Click(Index As Integer)

Sub farberklärung()
'Shell uverz & "programmierung\Dateilesen\LaborFarblegende.htm", vbNormalFocus
Dim Result&
Result = ShellExecute(Me.hwnd, "Open", uVerz & "Programmierung\Dateilesen\LaborFarblegende.htm", "", App.path, 1)
#If False Then
 Dim fI As New FürIcon
 fI.Height = 16000
 fI.Width = 16000
 fI.Caption = "Farberklärung für pathologische Laborwerte"
Dim Cmd1 As Control
Set Cmd1 = fI.Controls.Add("vb.textbox", "Cmd1")

Cmd1.Width = 2000
Cmd1.Top = 10 ' Me.Height / 2 - Cmd1.Height / 2 - 100
Cmd1.Left = 10 ' Me.Width / 2 - Cmd1.Width / 2 - 100
Cmd1 = "Dynamic Button"
Cmd1.Visible = True
fI.Show
#End If
End Sub ' farberklärung

Sub GesZF()
 Dim i&, immeranhaeng% ' beim ersten Mal fragen, dann anhängen
' Sicherheitskopie
 Open uVerz & "tmimport\Umfang" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & ".txt " For Output As #335
 For i = 1 To GesColl.COUNT
  Print #335, "0193000" & GesColl(i)
 Next i
 Close #335
 
 Dim Datei$
 Datei = uVerz & "tmimport\" & "DMP_Import_" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & ".BDT"
 For i = 1 To GesColl.COUNT
  callMachDMPBogen GesColl(i), "", False, immeranhaeng, True, Datei
  immeranhaeng = True
 Next i
 MsgBox "Fertig mit DMP gesamt!"
End Sub ' GesZF

Sub aktFS()
 Call doFS(False)
End Sub ' aktFS

Sub alleFS()
 Dim pCol As New Collection, i&, zeile
' Dim altCol&, altRow& ' evtl. zu streichen 19.4.10
 Select Case PLArt
  Case artPat ' DMP hier Liste
   FNr = 6
   With Me.MFG
    .col = 0
    For i = .Row To .Rows - 1
     .Row = i
     If .CellBackColor = DunkelRosa Then
      pCol.Add i
     End If
    Next i
    .Row = altrow
    .col = altCol
   End With
   For Each zeile In pCol
    Call FertigStellen((zeile))
    Pause (Pausenlänge)
    SendKeys "{ESC}", True
   Next zeile
  Case artHA
 End Select
End Sub ' alleFS

Sub Auffrisch()
 Unload Me
 fgespei(MFGTyp) = 0
 Load Me
End Sub ' Auffrisch()

Private Sub plz()
 Dim Pat_id$
 Pat_id = MFG.TextMatrix(MFG.Row, PIDSp)
 Call dodoPLZ(Pat_id, plzVz, Now, Now - Int(Now), True)
End Sub ' Anleitung


Private Sub Form_Activate()
' IF Me.hlese.Visible THEN Me.hlese.Hide
End Sub ' Form_Activate()

Private Sub Form_Deactivate()
 On Error Resume Next
 If Not Me.hlese.Visible Then Me.hlese.Show
End Sub ' Form_Deactivate()

Private Sub Form_Unload(Cancel As Integer)
 RegSpeichern
 On Error Resume Next
 Me.hlese.Show
End Sub ' Form_Unload(Cancel As Integer)

Public Sub li1_Click()
 Dim rAf&, i&
 FNr = 7
 If Cstumm = 0 Then
  Li1.Visible = False
  If Li1 <> "" Then
   For i = cRow(MFGTyp) To cRowSel(MFGTyp)
    MFG.col = 0
    MFG.Row = i
    Call myEFrag("UPDATE `diagreihe` SET gi1 = (SELECT MIN(lfdnr) FROM diagg1 WHERE gruppe='" & Li1 & "') WHERE icd = '" & MFG.TextMatrix(i, 0) & "'", rAf)
    If rAf <> 1 Then
     MsgBox ("Fehler beim Diagnosenupdate: " & Li1 & " " & MFG.TextMatrix(i, 0))
    End If
    MFG.col = GruSp
    MFG.Text = Li1
   Next
  End If
'  Call MFGrefresh
  If Me.MFG.Row < Me.MFG.Rows Then
   Me.MFG.Row = Me.MFG.Row + 1
  End If
  cRow(MFGTyp) = Me.MFG.Row
  cRowSel(MFGTyp) = Me.MFG.RowSel
  Me.Li1.Clear
  Me.Li1.Visible = False
 End If
 Me.MFG.SetFocus
End Sub ' li1_click

Public Sub Text1_Fertig()
 Dim i&
 Dim rAf&
 FNr = 23
 If Cstumm = 0 Then
  Text1.Visible = False
  If Text1 <> "" Then
   For i = cRow(MFGTyp) To cRowSel(MFGTyp)
    MFG.col = 0
    MFG.Row = i
    Call myEFrag("UPDATE diagreihe SET gi2 = '" & Text1 & "' WHERE icd = '" & MFG.TextMatrix(i, 0) & "'", rAf)
    If rAf <> 1 Then
     MsgBox ("Fehler beim Diagnosenupdate: " & Text1 & " " & MFG.TextMatrix(i, 0))
    End If
    MFG.col = GruSp + 1
    MFG.Text = Text1
   Next
  End If ' text1<>""
  If Me.MFG.Row < Me.MFG.Rows Then
   Me.MFG.Row = Me.MFG.Row + 1
  End If
  cRow(MFGTyp) = Me.MFG.Row
  cRowSel(MFGTyp) = Me.MFG.RowSel
  Me.Text1.Visible = False
 End If
 Me.MFG.SetFocus
End Sub ' Text1_Click

Sub FertigStellenBeliebig()
 Dim erg As Variant
 FNr = 8
 Select Case PLArt
  Case artDMP
   Stop
  Case artDiag
  
  Case artPat ' DMP hier Liste
   erg = InputBox("Bitte Pat-ID eingeben:", "Eingabe der Patientennummer", Me.MFG.TextMatrix(Me.MFG.Row, 1))
   If IsNumeric(erg) Then
    Call FertigStellen(0, , CLng(erg))
   End If
   
  Case artLPar
  
  Case artHA
 End Select
End Sub ' FertigStellenBeliebig()

Private Sub DokuBeliebig() ' Doku zu beliebigem Patienten
 Dim erg As Variant, VorDoku$
 Select Case PLArt
 
  Case artDMP
   Stop
   
  Case artDiag
  
  Case artPat ' DMP hier Liste
   erg = InputBox("Bitte Pat-ID eingeben:", "Eingabe der Patientennummer", Me.MFG.TextMatrix(Me.MFG.Row, DPatIDSp))
   
   If IsNumeric(erg) Then
    alttop = MFG.TopRow
    altr = MFG.Row
    Call MFG_leavecell
    If erg = Me.MFG.TextMatrix(Me.MFG.Row, DPatIDSp) Then VorDoku = Me.MFG.TextMatrix(Me.MFG.Row, VorDokuSp)
    Call callMachDMPBogen(CLng(erg), VorDoku, True)
    FNr = 9
    MFG.SetFocus
    With MFG
     For aktr = altr + 1 To .Rows - 1
      .Row = aktr
      .col = NachNameSp
      If .CellBackColor = DunkelRot Then
       Call TopAusricht
       Call mfg_entercell
       .col = .col - 1
       Exit For
      End If
     Next
    End With
   End If
   
  Case artLPar
  
  Case artHA
  
 End Select
End Sub ' DokuBeliebig

Public Sub Nächster()
  MFG.SetFocus
  If MFG.Row < MFG.Rows - 1 Then
   alttop = MFG.TopRow
   altr = MFG.Row
   Call MFG_leavecell
   aktr = MFG.Row + 1
   MFG.Row = aktr
   Call TopAusricht
   mfg_entercell
  End If
End Sub ' Nächster()

Public Sub Voriger()
  MFG.SetFocus
  If MFG.Row > 1 Then
   alttop = MFG.TopRow
   altr = MFG.Row
   Call MFG_leavecell
   aktr = MFG.Row - 1
   MFG.Row = aktr
   Call TopAusricht
   mfg_entercell
  End If
End Sub ' Voriger()

Public Sub NächsterGleichartiger()
 Dim aktZA As BogArtTyp, i&
 FNr = 9
 MFG.SetFocus
 aktZA = getBogArtTyp
 With MFG
  alttop = .TopRow
  altr = .Row
  Call MFG_leavecell
  For aktr = .Row + 1 To .Rows - 1
   .Row = aktr
   If getBogArtTyp = aktZA Then
    Call TopAusricht
    Call mfg_entercell
    Exit Sub
   End If
  Next
  .Row = altr
  Call mfg_entercell
 End With
 
End Sub ' NächsterGleichartiger

Public Sub Suchen(Optional weiter%, Optional alleSp%) ' Suchen
 FNr = 10
 MFG.SetFocus
 Dim VonSp&, BisSp&
 If Not weiter Then SuchStr = LCase$(InputBox("Bitte Suchstring eingeben:", "Zeilensuche"))
 With MFG
  alttop = .TopRow
  VonSp = IIf(alleSp, 0, .col)
  BisSp = IIf(alleSp, .Cols - 1, .col)
'  altr = .Row
'  altC = .col
  Call MFG_leavecell
'  .col = 3
  For aktr = .Row + 1 To .Rows - 1
'   IF alleSp AND .TextMatrix(aktr, VonSp) = "" THEN Exit For
   For aktc = VonSp To BisSp
    If alleSp And .TextMatrix(aktr, aktc) = "" Then Exit For
    If (alleSp And InStrB(LCase$(.TextMatrix(aktr, aktc)), SuchStr)) Or (Not alleSp And LCase$(.TextMatrix(aktr, aktc)) Like SuchStr & "*") Then
     .Row = aktr
     .col = aktc
     Call TopAusricht
     GoTo Ende
    End If
   Next aktc
  Next aktr
'  .Row = altr
'  .col = altC
Ende:
  Call mfg_entercell
 End With
End Sub ' Suchen

Public Sub FertigStellen(zeile&, Optional nuranzeigen%, Optional PatID&) ' nachdem BDT-Datei(en) manuell importiert wurde(n)
 Const obStumm% = 0
 FNr = 11
 Dim VorDoku$, Pat_id&, dtyp%, rs As New Recordset
' Dim aktDC AS DMPClass
 Dim j%
 Dim rTyp As New ADODB.Recordset
 On Error GoTo fehler
 With MFG
  .SetFocus
  alttop = .TopRow
  altr = zeile
'  altC = .col
  If PatID <> 0 Then
   Pat_id = PatID
   myFrag rs, "SELECT quartal FROM `faelle` WHERE pat_id = " & PatID & " AND bhfb < now()- " & Verspätung & " ORDER BY bhfb DESC", , DBCn
  Else
   VorDoku = .TextMatrix(zeile, VorDokuSp)
   Pat_id = .TextMatrix(zeile, PIDSp)
  End If
'  .col = altC
 End With
' Call DMPAusgeb0(aktDC, Pat_id, Not obstumm)
 zwiFS Pat_id, nuranzeigen
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FertigStellen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' FertigStellen

Sub zwiFS(Pat_id&, nuranzeigen%)
 Dim hnd&
 On Error GoTo fehler
 hnd = FensterHandle("TurboMed")
 If hnd <> 0 Then
  On Error GoTo activatefehler
  AppActivate "TurboMed", True
  On Error GoTo fehler
  Pause (Pausenlänge)
  On Error Resume Next
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ENTER}", True
  Pause (Pausenlänge)
  SendKeys "{F12}", True
  Pause (Pausenlänge)
'  SendKeys "{ESC}", True
'  Pause (Pausenlänge)
'  SendKeys "{ESC}", True
'  Pause (Pausenlänge)
  SendKeys "p", True
  Pause (Pausenlänge)
  SendKeys "{bs}" & Pat_id & "", True
  
  Call doFS(nuranzeigen, True)
 End If
 Exit Sub
activatefehler:
 MsgBox "Fehler bei AppActivate 'TurboMed'"
 Resume Next
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in zwiFS/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' zwiFS(Pat_id&, nuranzeigen%)

Sub doFS(nuranzeigen%, Optional ohneakt% = False)
  FNr = 12
  Me.Text1.Enabled = False
  If Not ohneakt Then AppActivate "TurboMed", True
  Pause (Pausenlänge)
  SendKeys "{ENTER}", True
  Pause (Pausenlänge)
  SendKeys "{F3}", True
  Pause (Pausenlänge)
  If nuranzeigen <> 0 Then
'   AppActivate "TurboMed", True
   If Me.Text1.Visible <> 0 Then
    If IsNumeric(Me.Text1) Then
     If CDbl(Me.Text1) < 10 Then
      SendKeys "%{F" & Me.Text1 & "}", True
     Else
      SendKeys "{F5}", True
      Pause (Pausenlänge)
      SendKeys "{DOWN " & Me.Text1 - 2 & "}", True ' Filter DMP wählen
      Pause (Pausenlänge)
     End If
    Else
     SendKeys "%{F1}", True
    End If
   Else
    SendKeys "%{F1}", True
   End If
   Pause (Pausenlänge)
   Exit Sub
  End If
  SendKeys "{F5}", True
  Pause (Pausenlänge)
  SendKeys "{DOWN " & Me.Text1 - 2 & "}", True ' Filter DMP wählen
  Pause (Pausenlänge)
  SendKeys "{Enter}", True
  Pause (Pausenlänge)
  SendKeys "{Enter}", True
  Pause (Pausenlänge)
  SendKeys "%D", True
  Pause (Pausenlänge)
  SendKeys " ", True
  Pause (Pausenlänge)
  SendKeys "%O", True
  Pause (Pausenlänge)
  SendKeys " ", True
  Pause (Pausenlänge)
  SendKeys " " & Format(MINvb(Now(), QEnd(ZQuart(Now() - Verspätung))), "ddmmyyyy") & "{TAB}", True
  Pause (Pausenlänge)
  SendKeys "%O", True
  Pause (Pausenlänge)
  SendKeys " ", True
  Pause (Pausenlänge)
  SendKeys " ", True
  Pause (Pausenlänge)
  SendKeys " ", True
  Pause (Pausenlänge)
'  AppActivate "C:\TurboMed\Tmp", True
'  Pause (Pausenlänge)
'  Dim rs As New ADODB.Recordset
'  myFrag rs, "SELECT nachname, vorname, gebdat FROM `namen` WHERE pat_id = " & Pat_id
'  SendKeys "%DU", True
'  Pause (Pausenlänge)
'  SendKeys pverz & rs!Nachname & " " & rs!Vorname & ", geb. " & rs!GebDat & ", DMP-Doku " & qend(ZQuart(NOW() - Verspätung))
'  Pause (Pausenlänge)
'  SendKeys "{Enter}", True
  Call htmlb(obStumm:=True)
  Me.Text1.Enabled = True
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doFS/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub      ' doFS

Private Sub zurück()
 Dim erg$, erg1$, erg2$
 FNr = 13
 erg = Dir(hVerz & DMP_Import)
 If LenB(erg) <> 0 Then
  Open hVerz & DMP_Import & "kopie.BDT" For Output As #389
  Open hVerz & DMP_Import For Input As #390
  Do While Not EOF(390)
   Input #390, erg
Nächster:
   If Mid$(erg, 4, 4) = "8000" Then
    Input #390, erg1
    Input #390, erg2
    If Mid$(erg2, 4, 4) = "3000" And Mid$(erg2, 8, 4) = MFG.TextMatrix(MFG.Row, 1) Then
'     Stop
     Do While Not EOF(390)
      Input #390, erg
      If Mid$(erg, 4, 4) = "8000" Then GoTo Nächster
     Loop
    Else
'     Stop
     Print #389, erg
     Print #389, erg1
     Print #389, erg2
    End If
   Else
    Print #389, erg
   End If
  Loop
  Close #389
  Close #390
  Kill hVerz & DMP_Import
  Name hVerz & DMP_Import & "kopie.BDT" As hVerz & DMP_Import
  Call Auffrisch
 End If
End Sub ' zurück

Private Sub htmlb(Optional obStumm%) ' HTML-Dateien bewahren
  Dim TempDir$, erg$, UnterVerz$(), vz, lin$, lina$, dmpArt$, NN$, VN$, Geb$, aktdat$, Zahl&
  Dim Ziel$
  FNr = 14
  Ziel$ = pVerz
  ReDim UnterVerz(0)
  TempDir = getDokPfad("Temporärdateien")
  erg = Dir(TempDir & "\", vbDirectory)
  Do
   If LenB(erg) = 0 Then Exit Do
   If erg <> "." And erg <> ".." Then
    If (GetAttr(TempDir & "\" & erg) And vbDirectory) = vbDirectory Then
     UnterVerz(UBound(UnterVerz)) = erg
     ReDim Preserve UnterVerz(UBound(UnterVerz) + 1)
    End If
   End If
   erg = Dir
  Loop
  For Each vz In UnterVerz
   erg = Dir(TempDir & "\" & vz & "\")
   Do
    If LenB(erg) = 0 Then Exit Do
    If InStrB(erg, ".html") <> 0 Then
     lina = vNS
     dmpArt = vNS
     NN = vNS
     VN = vNS
     Geb = vNS
     aktdat = TempDir & "\" & vz & "\" & erg
     Open aktdat For Input As #329
     Do While Not EOF(329)
      lina = lin
      Input #329, lin
      If InStrB(lina, """DMP-Header""") <> 0 Then
       dmpArt = lin
       dmpArt = Mid$(dmpArt, InStr(dmpArt, "<h1>") + 4)
       dmpArt = Left$(dmpArt, InStr(dmpArt, "</h1>") - 1)
      ElseIf InStrB(lina, "Vorname") <> 0 Then
       NN = lin
       NN = Mid$(NN, InStr(NN, "<div class=""header_value"">") + Len("<div class=""header_value"">"))
       NN = Left$(NN, InStr(InStr(NN, "<br>") + 1, NN, "<br>") - 1)
       NN = REPLACE$(NN, "<br>", " ")
      ElseIf InStrB(lina, "geb. am") <> 0 Then
       Geb = lin
       Geb = Mid$(Geb, InStr(Geb, "<div class=""header_value"">") + Len("<div class=""header_value"">"))
       Geb = Left$(Geb, InStr(Geb, "</div>") - 1)
       Exit Do
      End If
     Loop
     Close #329
     On Error Resume Next
     Kill Ziel & NN & ", geb. " & Geb & ", DMP-" & dmpArt & ".html"
     On Error GoTo fehler
     FileCopy aktdat, Ziel & NN & ", geb. " & Geb & ", DMP-" & dmpArt & ".html"
     Zahl = Zahl + 1
     Kill aktdat
'     Debug.Print erg & " " & dmpArt & " " & NN & " " & Geb
    End If
    erg = Dir
   Loop
  Next vz
  If Not obStumm Then
   MsgBox "Fertig mit HTML bewahren, " & Zahl & " Dateien nach " & Ziel & " kopiert und umbenannt"
  End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in htmlb/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub

Public Sub dokuErstelle() ' Erstelle
 Dim VorDoku$, Pat_id&
 FNr = 15
 Select Case PLArt
 
  Case artDMP
  
  Case artDiag
   Call MFGrefresh
   
  Case artHA
  
  Case Else ' artPat ' DMP hier Liste : kommt ohnehin als einziges vor
   With MFG
    .SetFocus
    alttop = .TopRow
    altr = .Row
    VorDoku = .TextMatrix(.Row, VorDokuSp)
    Pat_id = .TextMatrix(.Row, PIDSp)
   End With
   Call callMachDMPBogen(Pat_id, VorDoku)
   
 End Select
End Sub ' dokuErstelle

Public Sub callMachDMPBogen(Pat_id&, Optional VorDoku$, Optional obmitauswahl%, Optional immeranhaeng%, Optional obStumm%, Optional Datei$)  ' Erstelle
 If Datei = "" Then Datei = uVerz & "tmimport\" & DMP_Import
 Dim rTyp As New ADODB.Recordset, dtyp%
 Dim dmpba As New DMPBogenauswahl
 FNr = 16
 myFrag rTyp, "SELECT icd FROM diagnosen d WHERE d.pat_id = " & Pat_id & " AND d.diagsicherheit <> 'A' AND d.icd LIKE 'E1%' ORDER BY icd", adOpenStatic, DBCn, adLockReadOnly ' AND COALESCE(d.f6010,0)=0
 If Not rTyp.EOF Then dtyp = Mid$(rTyp.Fields(0), 3, 1) + 1
 If VorDoku = "" Then
      Dim rDok As New ADODB.Recordset
      Dim begcol%, j%, AktCol%, obraus%
'      rDok.Open "SELECT `DokuDatum`, `Art`, `ausgedruckt`, `OK`, `exportiert` FROM `dmpreihe` dr WHERE pat_id = " & Pat_id & "  AND (dr.Abk LIKE 'eDMPDM%' OR dr.Abk LIKE 'DMPDTYP%') ORDER BY `DokuDatum` DESC", DBCn, adOpenDynamic, adLockReadOnly
      myFrag rDok, "SELECT `DokuDatum`, `Art`, `ausgedruckt`, `OK`, `exportiert` FROM `dmpreihe` dr WHERE pat_id = " & Pat_id & "  AND (dr.Abk LIKE 'eDMPDM%' OR dr.Abk LIKE 'DMPDTYP%') ORDER BY `DokuDatum` DESC"
      If Not rDok.BOF Then
       begcol = VorDokuSp - 1
'       IF ZQSort(NOW() - 120) > ZQSort(rDok!DokuDatum) THEN ' Now() - rDok!DokuDatum > 120 THEN ' 21. des übernä Monats
'  7.3.16: Eine Doku darf ausgelassen werden
       If ZQSort(Now() - 211) > ZQSort(rDok!DokuDatum) Then ' Now() - rDok!DokuDatum > 120 THEN ' 21. des übernä Monats
        obraus = True
        begcol = begcol + 1
       End If
       If ZQuart(Now() - Verspätung) <> ZQuart(rDok!DokuDatum) Then
        begcol = begcol + 1
       For j = begcol To VorDokuSp ' MFG.Cols - 1
        If rDok.EOF Then Exit For
        AktCol = j
'        IF VorDokuSp = 0 THEN VorDokuSp = AktCol
        VorDoku = rDok!art & " " & Format(rDok!DokuDatum, "dd.mm.yy")
        If rDok!Ok And rDok!ausgedruckt Then
         VorDoku = VorDoku & " ok"
        ElseIf j = begcol And Not obraus Then
'         obrot = True ' Wenn Pat. rausgeflogen, dann fehlt auch aktuelle Erstdoku
        End If
        If rDok!exportiert <> 0 Then
         VorDoku = VorDoku & " ex"
'         IF rDok!Art = "ED" THEN .CellBackColor = blau ELSE .CellBackColor = vbgelblichgrau ' blau / grau
        Else
'         IF rDok!Art = "ED" THEN .CellBackColor = hellblau
        End If
'        IF rDok!Art = "ED" THEN oberst = True
        rDok.Move 1
       Next j
      End If
     End If
 End If
 If dtyp = 2 Then If VorDoku = vNS Then BogArtVar = typ2neu Else BogArtVar = typ2alt Else If VorDoku = vNS Then BogArtVar = typ1neu Else BogArtVar = typ1alt
 DokuDat = MINvb(Now(), QEnd(ZQuart(Now() - Verspätung)))
 If obmitauswahl Then
  dmpba.Option1(BogArtVar - 1) = True
  Set dmpba.vater = Me
  Dim rs As New ADODB.Recordset
  myFrag rs, "SELECT nachname n, vorname v FROM `namen` WHERE pat_id = " & Pat_id
  If rs.EOF Then
   MsgBox "Pat_id " & Pat_id & " nicht gefunden!"
   Exit Sub
  Else
   dmpba.Caption = "Erstelle DMP-Bogen zu " & Pat_id & " (" & rs!n & ", " & rs!V & ")"
   dmpba.DokuDatum = DokuDat
   dmpba.Show vbModal, Me
'  BogArtVar = dmpba.Option1
   If BogArtVar = 0 Then Exit Sub
  End If
 End If
 Call domachDMPBogen(Pat_id, BogArtVar, DokuDat, immeranhaeng, Not obmitauswahl, obStumm, Datei)
End Sub ' callMachDMPBogen(Pat_id&, Optional VorDoku$, Optional obmitauswahl%, optional immeranhaeng)


'Public Sub Typ1ED()
' Dim Pat_id&
' With MFG
''  altC = .col
'  Pat_id = .TextMatrix(.Row, PIDSp)
''  .col = altC
' END With
' Call machDMPBogen(Pat_id, typ1neu)
'End Sub ' command6_click
'Public Sub Typ1FD()
' Dim Pat_id&
' With MFG
''  altC = .col
''  .col = 2
'  Pat_id = .TextMatrix(.Row, PIDSp)
''  .col = altC
' END With
' Call machDMPBogen(Pat_id, typ1alt)
'End Sub
'Public Sub Typ2ED()
' Dim Pat_id&
' With MFG
'  altC = .col
''  .col = 2
'  Pat_id = .TextMatrix(.Row, PIDSp)
''  .col = altC
'  END With
'  Call machDMPBogen(Pat_id, typ2neu)
'End Sub
'Public Sub Typ2FD()
' Dim Pat_id&
' With MFG
'  altC = .col
''  .col = 2
'  Pat_id = .TextMatrix(.Row, PIDSp)
''  .col = altC
'  END With
'  Call machDMPBogen(Pat_id, typ2alt)
'End Sub


' korrigiert was
Private Sub Command2_Click()
 Dim erg$, pos%
 Dim Datei$, Quelldatei$
 Datei = uVerz & "tmimport\" & "DMP_Import_" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & ".BDT"
 Dim plf As New PathLabForm
 plf.Titel = "Ursprungsdatei auswählen"
 plf.Filter = "BDT-Dateien(*.BDT)|*.BDT|txt-Dateien(*.txt)|*.txt|Alle Dateien (*.*)|*.*"
 plf.LabordateiLab.Caption = "&Ursprungsdatei"
 plf.initDir = uVerz & "tmimport"
 plf.Caption = "Ursprungsdatei auswählen"
 plf.weiter = False
 Set plf.eLese = Me.hlese
 plf.Show 1, Me
 If plf.erg = "" Then Exit Sub
   ' Dateiauswahl auswählen
   Open plf.erg For Input As #388
   Do While Not EOF(388)
    Input #388, erg
    pos = InStr(erg, "3000")
    If pos = 4 Then
     Dim Pat_id$
     Pat_id = Mid$(erg, pos + 4)
'     Debug.Print Pat_id
#If True Then
     callMachDMPBogen CLng(Pat_id), "", 0, True, True, Datei
#Else
    ElseIf InStr(erg, "-Dokumentation Diabetes") > 0 Then
     If Pat_id <> "" Then
      pos = InStr(10, erg, "Typ 2")
      Dim p2%
      p2 = InStr(erg, "Folge")
      Dim bog As BogArtTyp
      If pos = 0 And p2 = 0 Then
        bog = typ1neu
      ElseIf pos <> 0 And p2 = 0 Then
       bog = typ2neu
      ElseIf pos = 0 And p2 <> 0 Then
       bog = typ1alt
      ElseIf pos <> 0 And p2 <> 0 Then
       bog = typ2alt
      End If
       Dim pid&
       pid = CLng(Pat_id)
       Select Case pid
        Case 184, 349, 1445, 1259, 1510, 1954, 59884, 53126, 53381, 7822, 31850, 38909, 59693, 59649, 51594, 59492, 59487, 53617, 53619, 53895, 54045, 59347, 59347, 1530
        Case Else
         domachDMPBogen CLng(Pat_id), bog, CDate(Now() - 1), True, True, True, Datei
       End Select
      Pat_id = ""
     End If
#End If
     'machdmpbogen(pat_id,boga
    End If
   Loop
   Close #388
   MsgBox "Fertig mit Neuerstellung der DMP-Import-Datei"
End Sub ' Command2_Click()

' in callMachDMPBogen und auskommentiert in Command2_Click
' die Konstanten DokuVersion und Datenerfassung müssen jedes Quartal überprüft und ggf. geändert werden!
Public Sub domachDMPBogen(Pat_id&, BogArtlV As BogArtTyp, DokuDat As Date, Optional immeranhaeng%, Optional autolanr%, Optional obStumm%, Optional Datei)  ' Erstelle
 If Datei = "" Then Datei = uVerz & "tmimport\" & DMP_Import
 Const DokuVersion$ = "201410"
' Dim rsAnam As New ADODB.Recordset
' Dim rnam As New ADODB.Recordset
 Dim rfal As New ADODB.Recordset
 Dim rform As New ADODB.Recordset
 Dim hnd&
 Dim aktDC As DMPClass, j%, pos%
' Dim RRsyst%, RRdiast%
 Dim SendStr$, SendStr1a$, SendStr2$
 Dim BDT As New BDTSchreib, erg$, DokuDatS$ ' DokuDat As Date
 Dim ErstellDat$, i%, GewichtS$, GrößeS$, HbA1cS$, HypoZS$, KrZS$, KreaS$, eGFRs$, RRS$
  
 Dim WagSt$, ArmSt$
 Dim ab315%, ab317%, ab0418%, ab1118%, ab0619%, ab0921%, ab0423%, ab1023%
 ab1023 = (DokuDat > #9/30/2023#)
 ab0423 = (DokuDat > #3/31/2023#)
 ab0921 = (DokuDat > #6/30/2021#)
 ab0619 = (DokuDat > #4/1/2019#)
 ab1118 = (DokuDat > #10/1/2018#)
 ab0418 = (DokuDat > #3/31/2018#)
 ab317 = (DokuDat > #6/30/2017#)
 ab315 = (DokuDat > #6/30/2015#)
 Dim Datenerfassung$, FeldFormVersion$
 'Datenerfassung$ = IIf(ab1023, "Okt_2023", IIf(ab0423, "April_2023", IIf(ab0921, "July_2021", IIf(ab0619, "April_2019", IIf(ab1118, "Okt_2018", IIf(ab0418, "Jan_2018", IIf(ab317, "July_2017", "15_07")))))))
 Datenerfassung$ = Switch(ab1023, "Okt_2023", ab0423, "April_2023", ab0921, "July_2021", ab0619, "April_2019", ab1118, "Okt_2018", ab0418, "Jan_2018", ab317, "July_2017", True, "15_07")
 FeldFormVersion = Datenerfassung
 
 FNr = 17
 On Error GoTo fehler
 Dim iru%
 For iru = 1 To 2
  erg = Dir(Datei)
  If LenB(erg) <> 0 Then Exit For
  Datei = REPLACE$(Datei, "u:", LiServer & "Daten\eigene Dateien")
 Next iru
 If LenB(erg) <> 0 Then
  If Not immeranhaeng Then
   Select Case MsgBox("Datei '" & Datei & "' vorhanden. Umbenennen (sonst anhängen)?", vbYesNoCancel + vbDefaultButton2, "Rückfrage")
    Case vbYes
     On Error Resume Next
     FSO.MoveFile Datei, REPLACE(Datei, ".BDT", Format(Now(), "_yymmdd_HHMMSS.B\DT"))
'     Kill Datei
     Kill REPLACE$(Datei, "u:", LiServer & "Daten\eigene Dateien")
     On Error GoTo fehler
    Case vbNo
    Case vbCancel: Exit Sub
   End Select
  End If ' Not immeranhaeng Then
 End If ' LenB(erg) <> 0 Then
 Dim auswlanr As New LANRauswahl
 
 Dim rlanr As New ADODB.Recordset
 If autolanr Then
  myFrag rlanr, "SELECT p.lanr FROM lanrpraxis p WHERE id = IF((SELECT MIN(lanrid) FROM faelle WHERE pat_id = " & Pat_id & " AND qanf = (SELECT MAX(qanf) FROM faelle WHERE pat_id = " & Pat_id & "))>0,(SELECT MIN(lanrid) FROM faelle WHERE pat_id = " & Pat_id & " AND qanf = (SELECT MAX(qanf) FROM faelle WHERE pat_id = " & Pat_id & ")),1)"
  If Not rlanr.BOF Then
   auswlanr.Lanr = rlanr!Lanr
   Set rlanr = Nothing
  End If
 Else ' autolanr Then
  auswlanr.PrepPatid Pat_id
  auswlanr.Show 1, Me
  auswlanr.Visible = False
  If auswlanr.Lanr = 0 Then Exit Sub
 End If ' autolanr Then else
 
' DokuDat = qend(ZQuart(NOW() - Verspätung))
 DokuDatS = Format(DokuDat, "dd.mm.yyyy")
 
' myFrag rsAnam, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_id
' myFrag rnam, "SELECT * FROM `namen` WHERE pat_id = " & Pat_id
 Dim lVorl As Date
 myFrag rfal, "SELECT " & IIf(Not LVobMySQL, "top 1", "") & " * FROM `faelle` WHERE pat_id = " & Pat_id & " AND bhfb <= " & DatFor_k(MINvb(Now(), QEnd(ZQuart(Now() - Verspätung)))) & " ORDER BY bhfb DESC, schgr" & IIf(LVobMySQL, " LIMIT 1", "")
 lVorl = rfal!lVorl
 myFrag rform, "SELECT " & IIf(Not LVobMySQL, "top 1", "") & " `feldinh` FROM `formular` WHERE pat_id = " & Pat_id & " AND Feld = 'Kasse' AND `zeitpunkt` <= " & DatFor_k(MINvb(Now(), QEnd(ZQuart(Now - Verspätung)))) & " AND feldinh LIKE '%'" & " ORDER BY zeitpunkt DESC" & IIf(LVobMySQL, " LIMIT 1", "") ' & aktdc.vknr & "%'"  geht nicht gut: VKNr nicht unbedingt aktuell in `faelle` (s.Pat_id 51)
' Ermittlung der 'Kasse' aus Rezepten oder Überweisungen oder vorherigen DMP-Dokus usw.
 Call DMPAusgeb0(aktDC, CStr(Pat_id), Not obStumm, , DokuDat) ' dort wird DMPString aufgerufen
 Dim Kasse$
 If Not rform.BOF Then
  Kasse = rform!FeldInh ' Trim$(replace$(rform!FeldInh, aktdc.vknr, ""))
 Else
  Set rform = Nothing
  myFrag rform, "SELECT kurzname, rname FROM `kassenliste` WHERE vknr = '" & rfal!VKNr & "' AND ik = '" & rfal!IK & "'", adOpenStatic, DBCn, adLockReadOnly 'aktDC.VKNr
  If Not rform.BOF Then
   Kasse = rform!kurzname
   If InStrB(Kasse, " ") <> 0 Then
    pos = InStr(Kasse, " ")
    Kasse = Left$(Kasse, pos - 1)
   End If
  End If
  Kasse = Left$(Left$(Kasse, 25) & Space$(25), 25) & aktDC.VKNr
 End If ' Not rform.BOF Then
 Set rform = Nothing
 
 
' myFrag rform, "SELECT " & IIf(Not LVobMySQL, "top 1", "") & " feldinh FROM `formular` WHERE pat_id = " & Pat_id & " AND Feld = 'Status' AND zeitpunkt <= " & DatFor_k(min(NOW(), qend(ZQuart(Now - Verspätung)))) & " AND feldinh LIKE '% %'" & " ORDER BY zeitpunkt DESC" & IIf(LVobMySQL, " LIMIT 1", "")
' IF Not rform.BOF THEN
'  Status = rform!FeldInh
' Else
'  Status = rfal!KVKs & " " & rfal!KVKserg
' END IF
' SET rform = Nothing
 
 ' QAnf(ZQuart(NOW() - verspätung))
 
 #Const zumdebug = 0
 #If zumdebug = 1 Then
  Call DMPString$(Pat_id, aktDC, , , DokuDat, 0)
 #End If
 Call BDT.ImportFolderHerricht(hVerz, Mid$(Datei, Len(hVerz) + 1))
 
 erg = Dir(BDT.DMPImp)
 If LenB(erg) = 0 Then
  Call BDT.BDTKopf
  Call BDT.BDTKo2(auswlanr.Lanr)
 End If ' LenB(erg) = 0
 BDT.Satzart "6200" ' Falldaten
 BDT.PatID Pat_id
 If aktDC.NVorsatz <> "" Then
  BDT.NVorsatz aktDC.NVorsatz
 End If
 BDT.Nachname aktDC.Nachname
 BDT.Vorname aktDC.Vorname
 BDT.Geb aktDC.GebDat
 BDT.DAdd "6200", DokuDatS ' DokuDatum, Dokumentationsdatum, = Tagesdatum
 BDT.TAdd "6201", "2000"
 BDT.Add "6203" & "TM#?##"       ' ergänzt 1.1.5
 BDT.Add "3635TM#" & auswlanr.Lanr
 BDT.Add "3636TM#" & BSNR
 
 Select Case BogArtlV
  Case typ2alt, typ1alt, typ2neu, typ1neu
   Dim DmT$, obErstD%
   If BogArtlV = typ2alt Or BogArtlV = typ2neu Then DmT = "2" Else DmT = "1"
   If BogArtlV = typ1neu Or BogArtlV = typ2neu Then obErstD = True Else obErstD = 0
   BDT.Add "6295eDMPDM" & DmT
   BDT.Add "6296" & IIf(obErstD, "Ersteinschreibung", "Verlaufsdokumentation") & " Diabetes mellitus Typ " & DmT & " (ok/ausgedruckt)"
   BDT.Add "6297$\TurboMed\Formulare\Patientenmenue\eDMP Datenerfassung" & IIf(ab315, "_" & Datenerfassung, "") & ".tmf "
   If Not ab1023 Then
    BDT.FFAdd "ACEHemmerDM" & DmT & "#" & -(Not aktDC.obACEH)
    BDT.FIAdd "X"
    If aktDC.obAT1 Then
     BDT.FFAdd "ACEHemmerDM" & DmT & "#3"
     BDT.FIAdd "X"
    End If ' aktDC.obAT1 Then
   End If ' Not ab1023 Then
   If aktDC.obRauch Or aktDC.dmpKKTabakEmpf = "j" Then
    BDT.Add "6298Angebot#0"
    BDT.FIAdd "X"
   End If
   If aktDC.dmpKKErnEmpf = "j" Then
    BDT.Add "6298Angebot#1"
    BDT.FIAdd "X"
   End If
   If aktDC.dmpKKkTrainEmpf = "j" Then
    BDT.Add "6298Angebot#2"
    BDT.FIAdd "X"
   End If
   If DmT = "2" Then
    BDT.FFAdd "AntidiabetischDM2#" & IIf(aktDC.obSonstAD = adja Or aktDC.obGlit = adja Or aktDC.obGlucI = adja Or aktDC.obSHGlin = adja Or aktDC.obDpp4 Or aktDC.obGlp1 Or aktDC.obSglt2, "0", "1") ' "0" = ja, "1" = nein
    BDT.FIAdd "X"
   End If
   If Not ab1023 Then
    If ab315 Then
     BDT.FFAdd "AntihypertensivDM" & DmT & "#" & -(Not aktDC.obDiur)
    Else
     BDT.FFAdd "AntihypertensivDM" & DmT & "#" & -(Not aktDC.obAntihyp)
    End If
    BDT.FIAdd "X"
   End If ' Not ab1023 Then
   If Not ab317 Then
    If ab315 Then
     If obErstD Then
      If aktDC.dmpArztw = "j" Then
       BDT.FFAdd "ArztwechselDM" & DmT & "#0"
       BDT.FIAdd "X"
      End If ' aktDC.dmpArztw = "j" Then
     End If ' obErstD Then
    End If ' ab315 Then
   End If ' Not ab317 Then
   If Not ab1023 Then
    BDT.FFAdd "BetablockerDM" & DmT & "#" & -(Not aktDC.obBetabl)
    BDT.FIAdd "X"
   End If ' Not ab1023 Then
   BDT.FFAdd "Datum"
   BDT.FIAdd Format(DokuDat, "dd.mm.yy")
   Dim diabFussEinr$
   If ab315 Then ' neu: 0=ja, 1=nein, 2=veranlasst
    Select Case aktDC.dmpUewFuss
     Case "j"
      diabFussEinr = "0"
     Case "v"
      diabFussEinr = "2"
     Case Else ' "n", fehlend
      diabFussEinr = "1"
    End Select ' case aktDC.dmpUewFuss
   Else ' alt: 0 = nein, 1 = diabetologisch qualifizierter Arzt 2 = ja (Fußeinrichtung)
    Select Case aktDC.dmpUewFuss
     Case "j", "v"
      diabFussEinr = "2"
     Case Else ' "n", fehlend
      diabFussEinr = "0"
    End Select ' case aktDC.dmpUewFuss
   End If ' ab315 else
   BDT.FFAdd "DiabFussDM" & DmT & "#" & diabFussEinr
   BDT.FIAdd "X"
   ' vorangegangene Doku, s. SEmpfohlen (aktuelle Doku)
   If Not obErstD Then
    Dim DiabWahrgNr$
    Select Case aktDC.dmpDMSchulWahrg
     Case "j"
      DiabWahrgNr = "0"
     Case "n"
      DiabWahrgNr = "1"
     Case "n mögl"
      DiabWahrgNr = "2"
     Case Else ' "u", "n empf"
      DiabWahrgNr = "3"
    End Select ' cse aktDC.dmpDMSchulWahrg
    BDT.FFAdd "DiabWahrgenommenDM" & DmT & "#" & DiabWahrgNr
    BDT.FIAdd "X"
   End If ' Not obErstD
   ' Dokumentationsintervall: 0=quartalsweise, 1=halbjährlich
   BDT.FFAdd "DintervallABV2#0"
   BDT.FIAdd "X"
   BDT.FFAdd "DintervallCOPD#0"
   BDT.FIAdd "X"
   If DmT = 1 Then
    BDT.FFAdd "DintervallDM1#" & IIf(aktDC.dmphalbj = "j", "1", "0")
    BDT.FIAdd "X"
    BDT.FFAdd "DintervallDM2#0"
    BDT.FIAdd "X"
   Else ' DmT = 1 Then
    BDT.FFAdd "DintervallDM1#0"
    BDT.FIAdd "X"
    BDT.FFAdd "DintervallDM2#" & IIf(aktDC.dmphalbj = "j", "1", "0")
    BDT.FIAdd "X"
   End If ' DmT = 1 Then else
   BDT.FFAdd "DintervallKHK#0"
   BDT.FIAdd "X"
   BDT.FFAdd "DokuTyp"
   BDT.FIAdd "DM" & DmT & "." & IIf(obErstD, "E", "F") & "D"
   BDT.FFAdd "DokuVersion"
   ' 5.4.15: für 3/15 muss mit DokuVersion 10/14 dokumentiert werden, wer weiß, mit welcher künftig ...
   BDT.FIAdd DokuVersion ' Format(QAnf(QuartalStr(DokuDat - 92)), "yyyymm") '=> haut nicht ganz hin
   Dim EinschrWg$
   Select Case DmT
    Case 1: EinschrWg = "1"
    Case 2: EinschrWg = "0"
   End Select ' case DmT
   ' Einschreibung wegen: 0= DM Typ 2, 1=Typ 1, 2=KHK, 3=Asthma, 4=COPD
   BDT.FFAdd "Einschreibung#" & EinschrWg
   BDT.FIAdd "X"
   If ab315 Then
    BDT.FFAdd "EinweisungDM" & DmT & "#" & IIf(aktDC.dmpEinwDM = "j", "0", IIf(aktDC.dmpEinwDM = "v", "2", "1")) ' 0 = ja, 1 = nein, 2 = veranlasst
    BDT.FIAdd "X"
   End If
   ' Ereignisse: 0 = Nierenersatztherapie, 1 = Erblindung, 2 = Amputation, 3 = Herzinfarkt, 4 = keine
   If aktDC.neuDial Or aktDC.neuErbl Or aktDC.neuAmp Or aktDC.neuMI Or (ab315 And aktDC.neuApo) Then
    If aktDC.neuDial Then
     BDT.FFAdd "EreignisseDM" & DmT & "#0"
     BDT.FIAdd "X"
    End If
    If aktDC.neuErbl Then
     BDT.FFAdd "EreignisseDM" & DmT & "#1"
     BDT.FIAdd "X"
    End If
    If aktDC.neuAmp Then
     BDT.FFAdd "EreignisseDM" & DmT & "#2"
     BDT.FIAdd "X"
    End If
    If aktDC.neuMI Then
     BDT.FFAdd "EreignisseDM" & DmT & "#3"
     BDT.FIAdd "X"
    End If
    If ab315 Then
     If aktDC.neuApo Then
      BDT.FFAdd "EreignisseDM" & DmT & "#4"
      BDT.FIAdd "X"
     End If
    End If
   Else ' keine
    BDT.FFAdd "EreignisseDM" & DmT & "#" & IIf(ab315, "5", "4")
    BDT.FIAdd "X"
   End If
   ErstellDat = Format(DokuDat, "ddmmyyyy")
   For i = 0 To 7
    BDT.FFAdd "ErstellungsdatumDM" & DmT & "#" & i
    BDT.FIAdd Mid$(ErstellDat, i + 1, 1)
   Next i
   Dim testvar$
   testvar = Format(Pat_id, "0000000")
   For i = 7 - Len(CStr(Pat_id)) To 6
    BDT.FFAdd "FallNummerDM" & DmT & "#" & i
    BDT.FIAdd Mid$(testvar, i + 1, 1)
   Next i
   If ab315 Then
    BDT.FFAdd "FeldFormVersion"
    BDT.FIAdd FeldFormVersion
    BDT.FFAdd "FeldLayoutXML"
    BDT.FIAdd "Patientenmenue\eDMP Layout_" & Datenerfassung & ".xml"
   End If
' 13.8.15: folgt Folgeerkrankung
   If Not (aktDC.FEn(15) Or aktDC.FEn(1) Or aktDC.FEn(5) Or aktDC.FEn(17) Or aktDC.FEn(2) Or aktDC.FEn(6) Or aktDC.FEn(16) Or aktDC.FEn(3) Or aktDC.FEn(4)) Then
    BDT.FFAdd "Folgeerkrankung#0" '14.8.15 o.k.
    BDT.FIAdd "X"
   Else
    If aktDC.FEn(15) Then ' COPD '14.8.15 o.k.
     BDT.FFAdd "Folgeerkrankung#1"
     BDT.FIAdd "X"
    End If
    If aktDC.FEn(1) Then ' Art. Hypertonie '14.8.15 o.k.
     BDT.FFAdd "Folgeerkrankung#2"
     BDT.FIAdd "X"
    End If
    If DokuDat < #6/10/2015# Then
     If aktDC.FEn(5) Then ' Schlaganfall
      BDT.FFAdd "Folgeerkrankung#3"
      BDT.FIAdd "X"
     End If
    End If
    If aktDC.FEn(17) Then ' Herzinsuffizienz '14.8.15 o.k.
     BDT.FFAdd "Folgeerkrankung#4"
     BDT.FIAdd "X"
    End If
    ' 5 gibts nicht
    If aktDC.FEn(2) Then ' Fettstoffwechselstörung '14.8.15 o.k., 29.8.17 o.k.
     BDT.FFAdd "Folgeerkrankung#6"
     BDT.FIAdd "X"
    End If
    If aktDC.FEn(6) Then ' AVK '14.8.15 o.k., 29.8.17 o.k.
     BDT.FFAdd "Folgeerkrankung#7"
     BDT.FIAdd "X"
    End If
    If aktDC.FEn(16) Then ' Asthma '14.8.15 o.k.
     BDT.FFAdd "Folgeerkrankung#8"
     BDT.FIAdd "X"
    End If
    If aktDC.FEn(3) Or aktDC.FEn(4) Then ' KHK '14.8.15 o.k., 29.8.17 o.k.
     BDT.FFAdd "Folgeerkrankung#9"
     BDT.FIAdd "X"
    End If
   End If
   If Not ab317 Then
    Select Case aktDC.fußst
     Case ndok
      BDT.FFAdd "FussstatusDM" & DmT & "#0"
     Case unauff
      BDT.FFAdd "FussstatusDM" & DmT & "#1"
     Case auff
      BDT.FFAdd "FussstatusDM" & DmT & "#2"
    End Select ' case aktDC.fußst
    BDT.FIAdd "X"
   End If ' not ab317
   If ab1023 Then
    If DmT = "2" Then
     BDT.FFAdd "GLP1DM2#" & -(Not (aktDC.obGlp1 = adja))
     BDT.FIAdd "X"
    End If ' DmT = "2" Then
   End If ' ab1023 Then
   
   GewichtS = Format(aktDC.gewi, "000")
   For i = 3 - Len(CStr(GewichtS)) To 2
    BDT.FFAdd "Gewicht#" & i
    BDT.FIAdd Mid$(GewichtS, i + 1, 1)
   Next i
   If Not ab1023 Then
    If DmT = "2" Then
     BDT.FFAdd "GlibenclamidDM2#" & -(Not (aktDC.obGlib = adja))
     BDT.FIAdd "X"
    End If ' DmT = "2" Then
   End If ' Not ab1023 Then
   If Not ab317 Then
    If aktDC.fußst = auff Then
     BDT.FFAdd "Grad0bisVDM" & DmT & "#" & Left$(aktDC.mWA, 1) ' Wagner-Stadium, stimmt zufällig überein
     BDT.FIAdd "X"
     ArmSt = Mid$(aktDC.mWA, 2, 1)
     If Len(ArmSt) = 1 And ArmSt >= "A" And ArmSt <= "D" Then
      BDT.FFAdd "GradAbisDDM" & DmT & "#" & Asc(ArmSt) - 65
      BDT.FIAdd "X"
     End If
    End If
   End If ' not ab317
   GrößeS = Format(aktDC.kgr, "000")
   For i = 3 - Len(CStr(GrößeS)) To 2
    BDT.FFAdd "Groesse#" & i
    BDT.FIAdd Mid$(GrößeS, i + 1, 1)
   Next i
   If Not ab1023 Then
    BDT.FFAdd "HMGDM" & DmT & "#" & -(Not aktDC.obHMG)
    BDT.FIAdd "X"
   End If ' Not ab1023 Then
   Dim HbA1clen%
   HbA1clen = 3
'   HbA1clen = IIf(ab1023, 2, 3)
   HbA1cS = Format(aktDC.bekHb * 10 ^ (HbA1clen - 2), Left$("000", HbA1clen))
   For i = 3 - HbA1clen To 2
    BDT.FFAdd "Hba1cDM" & DmT & "#" & i
    BDT.FIAdd Mid$(HbA1cS, i + HbA1clen - 2, 1)
   Next i
   
   If Not obErstD Then
    Dim HypertWahrgNr$
    Select Case aktDC.dmpHypertSchulWahrg
     Case "j"
      HypertWahrgNr = "0"
     Case "n"
      HypertWahrgNr = "1"
     Case "n mögl"
      HypertWahrgNr = "2"
     Case Else ' "u", "n empf"
      HypertWahrgNr = "3"
    End Select ' case aktDC.dmpHypertSchulWahrg
    BDT.FFAdd "HyperWahrgenommenDM" & DmT & "#" & HypertWahrgNr
    BDT.FIAdd "X"
   End If
   
   Dim HypoLen%
   HypoLen = 2
'   HypoLen = IIf(ab1023, 1, 2)
   HypoZS = Format(aktDC.hypoZ, Left$("00", HypoLen))
   For i = 2 - HypoLen To 1
    BDT.FFAdd "HypoglykaemenDM" & DmT & "#" & i
    BDT.FIAdd Mid$(HypoZS, i + HypoLen - 1, 1)
   Next i
   
   BDT.FFAdd "Ik"
'   BDT.FIAdd "10" & rfal!IK    ' 1.1.15: Vorsatz von "10" empirisch bei Monterosso ermittelt, in IK-Liste scheint es bei den DMP-Kassen 10 zu sein, bei LKK auch 11
   Dim rik As New ADODB.Recordset, IKs$, ikname$, ikprae$
   myFrag rik, "SELECT ikprae,name FROM IKs WHERE ik = " & aktDC.IK
   If Not rik.BOF Then
    ikprae = rik!ikprae
    ikname = rik!name
    IKs = ikprae
   Else
    ikprae = ""
    ikname = ""
    IKs = ""
   End If
   IKs = IKs & aktDC.IK  ' 1.1.15: scheint aber auch so zu gehen, was mir momentan sicherer erscheint
   BDT.FIAdd IKs
   If ab315 Then
    If (aktDC.obIns Or aktDC.obAnal) Then
     BDT.FFAdd "InjektionsstellenDM" & DmT & "#" & aktDC.inj ' 0=nicht untersucht, 1=unauffällig, 2 = auffällig
     BDT.FIAdd "X"
'    Else
'     BDT.FIAdd ""
    End If
   End If ' ab315 Then
   If Not ab0921 Or DmT <> "1" Then
    BDT.FFAdd "InsulinDM" & DmT & "#" & -(Not (aktDC.obIns Or aktDC.obAnal))
    BDT.FIAdd "X"
   End If
   BDT.FFAdd "KVKGueltig"
   BDT.FIAdd auswlanr.Lanr ' eigene LANR
   BDT.FFAdd "Kasse"
   BDT.FIAdd Kasse
   
   If Not ab315 Then
    KreaS = Format(aktDC.Crea * 100, "0000")
    For i = 4 - Len(CStr(KreaS)) To 3
     BDT.FFAdd "Kreatinin2#" & i
     BDT.FIAdd Mid$(KreaS, i + 1, 1)
    Next i
   End If
   
   If DmT = "2" Then
    BDT.FFAdd "MetforminDM2#" & -(Not (aktDC.obmetf = adja))
    BDT.FIAdd "X"
   End If
   BDT.FFAdd "ModulKHKVisible"     ' 1.1.15
   BDT.FIAdd "y"                   ' 1.1.15
   BDT.FFAdd "Nachname"
   BDT.FIAdd aktDC.Nachname
   BDT.FFAdd "NetzhautDM" & DmT & "#" & IIf(aktDC.aug = durchg, "0", IIf(aktDC.aug = ndurch, "1", "2"))
   BDT.FIAdd "X"
   BDT.FFAdd "Postleitzahl"
   BDT.FIAdd aktDC.Postleitzahl
   BDT.FFAdd "PulsstatusDM" & DmT & "#" & IIf(aktDC.Puls = ndok, "0", IIf(aktDC.Puls = unauff, "1", "2"))
   BDT.FIAdd "X"
   RRS = Format(aktDC.RRdiast, "000")
   For i = 0 To 2 ' hier wirklich 0
    BDT.FFAdd "RRd#" & i
    BDT.FIAdd Mid$(RRS, i + 1, 1)
   Next i
   RRS = Format(aktDC.RRsyst, "000")
   For i = 0 To 2 ' hier wirklich 0
    BDT.FFAdd "RRs#" & i
    BDT.FIAdd Mid$(RRS, i + 1, 1)
   Next i
   BDT.FFAdd "Raucher#" & -(Not aktDC.Tabak)
   BDT.FIAdd "X"
   ' s. wahrgenommen
   ' hier nicht .SE(0) und .SE(1), da diese erweitere Bedingungen haben, also:
   ' - für hausärztliche Dokum Schulung z.T. automatisch zu empfehlen,
   ' - für eigene Doku nur über das Makro dmp
   If aktDC.dmpDMSchulEmpf <> "j" And aktDC.dmpHypertSchulEmpf <> "j" Then
    BDT.FFAdd "SEmpfohlenDM" & DmT & "#2"
    BDT.FIAdd "X"
   Else ' aktDC.dmpDMSchulEmpf <> "j" And aktDC.dmpHypertSchulEmpf <> "j" Then
    If aktDC.dmpDMSchulEmpf = "j" Then
     BDT.FFAdd "SEmpfohlenDM" & DmT & "#0"
     BDT.FIAdd "X"
    End If ' If aktDC.dmpDMSchulEmpf = "j" Then
    If aktDC.dmpHypertSchulEmpf = "j" Then
     BDT.FFAdd "SEmpfohlenDM" & DmT & "#1"
     BDT.FIAdd "X"
    End If ' aktDC.dmpHypertSchulEmpf = "j" Then
   End If ' aktDC.dmpDMSchulEmpf <> "j" And aktDC.dmpHypertSchulEmpf <> "j" Then else
   If ab1023 Then
    If DmT = "2" Then
     BDT.FFAdd "SGLT2DM2#" & -(Not (aktDC.obSglt2 = adja))
     BDT.FIAdd "X"
    End If ' DmT = "2" Then
   End If ' ab1023 Then
   BDT.FFAdd "SensibilitaetDM" & DmT & "#" & IIf(aktDC.sens = ndok, "0", IIf(aktDC.sens = unauff, "1", "2"))
   BDT.FIAdd "X"
   If aktDC.FEn(7) Then ' Nephropathie
    BDT.FFAdd "SpaetfolgenDM" & DmT & "#0"
    BDT.FIAdd "X"
   End If
   If aktDC.FEn(11) Then ' Neuropathie
    BDT.FFAdd "SpaetfolgenDM" & DmT & "#1"
    BDT.FIAdd "X"
   End If
   If aktDC.FEn(9) Then ' Retinopathie
    BDT.FFAdd "SpaetfolgenDM" & DmT & "#2"
    BDT.FIAdd "X"
   End If
   If ab317 And Not ab0921 Then
    If DmT = "1" Then
     For i = 0 To 1
      BDT.FFAdd "StaHba1cDM1#" & i
      BDT.FIAdd "0"
     Next i
    End If
   End If ' not ab317
   
   KrZS = Format(aktDC.krz, "00")
   ' stationäre Einweisungen wegen Diabetes seit der letzten Doku unter "relevante Ereignisse"
   ' zwar auch schon im Formular vor 30.6.15, aber dort (wohl durch einen Fehler) nicht in der Maske
   If Len(CStr(KrZS)) > 2 Then KrZS = "00" ' 28.6.17: "2014"
   For i = 2 - Len(CStr(KrZS)) To 1
    BDT.FFAdd "StationaereDM" & DmT & "#" & i
    BDT.FIAdd Mid$(KrZS, i + 1, 1)
   Next i
   
   BDT.FFAdd "Status"
   Dim Status$
   Status = aktDC.KVKStatus ' 1.1.15 nach Monterosso
   ' https://de.wikipedia.org/wiki/Versichertenstatus, ab 10/14
   If Status <> "1" And Status <> "3" And Status <> "5" Then Status = "1"
   If ab0619 Or ab1118 Then
     Status = Status & "000" & IIf(BogArtlV = typ2alt Or BogArtlV = typ2neu, "1", "4") & "00"
   ElseIf BogArtlV = typ1alt Or BogArtlV = typ1neu Then
     Status = Status & "   4"
   ElseIf BogArtlV = typ2alt Or BogArtlV = typ2neu Then
     Status = Status & "   1"
   End If
   BDT.FIAdd Status
   
   BDT.FFAdd "Strasse"
   BDT.FIAdd aktDC.strasse
'   ErstellDat = Format(min(NOW(), qend(ZQuart(qend(ZQuart(Now - Verspätung))) + 1)), "ddmmyyyy")
   'ErstellDat = Format(qend(ZQuart(qend(ZQuart(Now - Verspätung)) + 1)), "ddmmyyyy") ' 7.12.10
   If ab1023 Then
    BDT.FFAdd "TeilnahmeHinweisAngezeigt"
    BDT.FIAdd "DM" & DmT
   End If
   If Not ab315 Then
    ErstellDat = Format(QEnd(ZQuart(DokuDat)), "ddmmyyyy") ' 15.8.15
    For i = 0 To 7
     BDT.FFAdd "TerminDM" & DmT & "#" & i
     BDT.FIAdd Mid$(ErstellDat, i + 1, 1)
    Next i
   End If
   'SysGluABV2
   BDT.FFAdd "ThromDM" & DmT & "#" & -(Not aktDC.obThro) ' Thrombozytenaggregationshemmer
   BDT.FIAdd "X"
   'ThromKHK
   If aktDC.obOAK Then
   End If
   'UeberweisungABV2
   If Not ab315 Then
    BDT.FFAdd "UeberweisungDM" & DmT & "#0" ' zum diabetologisch qual Arzt bzw. Einrichtung
    BDT.FIAdd "X"
   End If
   'UeberweisungKHK
   BDT.FFAdd "UlcusDM" & DmT & "#" & aktDC.ulcus
   BDT.FIAdd "X"
   If Not ab1023 Or BogArtlV = typ1alt Or BogArtlV = typ1neu Then
    BDT.FFAdd "UrinDM" & DmT & "#" & IIf(aktDC.mau = ndok, "0", IIf(aktDC.mau = auff, "1", "2"))
    BDT.FIAdd "X"
   End If ' Not ab1023 Then
   BDT.FFAdd "Versichertennr"
   BDT.FIAdd Trim$(aktDC.Versichertennummer)
   BDT.FFAdd "Vertragsarztnr"
   BDT.FIAdd BSNR
   If ab315 Then
    If aktDC.dmpVertret = "j" Then
     BDT.FFAdd "VertretungDM" & DmT & "#0"
     BDT.FIAdd "X"
    End If
   End If
   BDT.FFAdd "Vorname"
   BDT.FIAdd aktDC.Vorname
   BDT.FFAdd "ZielHba1cDM" & DmT & "#" & IIf(aktDC.hbEmpf = senken, "1", "0")
   BDT.FIAdd "X"
   '5.4.15:
   BDT.FFAdd "ZusatzDaten"
   Dim ZsD$
   For i = 1 To 32
    ZsD = ZsD & Hex$(GetRandomNum(1, 15)) ' 29.6.15 von 16 auf 15 reduziert
   Next i
   ZsD = "<ZusatzDaten><Dokument_ID>" & ZsD & "</Dokument_ID><Dokumentationsdatum>" & Format(DokuDat, "yyyy-mm-dd") & "</Dokumentationsdatum>"
   ZsD = ZsD & "<Patient><publicID>"
   Dim publicID$
   Dim Zusatzdaten$, SuchStr$
   
   Dim rDMP As New ADODB.Recordset
   myFrag rDMP, "SELECT zusatzdaten z FROM dmpreihe WHERE pat_id = " & Pat_id & "  AND (Abk LIKE 'eDMPDM%' OR Abk LIKE 'DMPDTYP%') ORDER BY karteidatum"
   If Not rDMP.BOF Then
    Zusatzdaten = rDMP!z
    SuchStr = "<Patient><publicID>"
    pos = InStr(Zusatzdaten, SuchStr)
    If pos Then publicID = Mid$(Zusatzdaten, pos + Len(SuchStr), 18)
   End If
   If publicID = "" Then
    For i = 1 To 18
     publicID = publicID & GetRandomNum(0, 9)
    Next i
   End If
   ZsD = ZsD & publicID
   publicID = ""
   If Right$(aktDC.strasse, Len(aktDC.Hausnr)) = aktDC.Hausnr Then aktDC.strasse = Trim$(Left$(aktDC.strasse, Len(aktDC.strasse) - Len(aktDC.Hausnr))) ' 30.6.15: Laednderkennzeichen korrigiert (iif(laenderkz="","D",aktdc.Lkz) ) entfernt), Postfach korrigiert (& aktdc.plz geloescht), fax (& aktdc.PrivatFax geloescht)
   ZsD = ZsD & "</publicID><Titel>" & aktDC.Titel & "</Titel><Vorsatz>" & aktDC.NVorsatz & "</Vorsatz><Zusatz>" & aktDC.NVors & "</Zusatz><Vorname>" & aktDC.Vorname & "</Vorname><Nachname>" & aktDC.Nachname & "</Nachname><Anschriftenzusatz>" & aktDC.Anschrzus & "</Anschriftenzusatz><Strasse>" & aktDC.strasse & "</Strasse>"
   ZsD = ZsD & "<Hausnummer>" & aktDC.Hausnr & "</Hausnummer><PLZ>" & aktDC.plz & "</PLZ><Stadt>" & aktDC.ort & "</Stadt><Laenderkennzeichen>" & Trim$(aktDC.Lkz) & "</Laenderkennzeichen><Postfach></Postfach><PostfachPLZ>" & aktDC.PFPlz & "</PostfachPLZ><PostfachOrt>" & aktDC.PFOrt & "</PostfachOrt>"
   ZsD = ZsD & "<Geschlecht>" & aktDC.geschlecht & "</Geschlecht><GeburtsDatum>" & Format(aktDC.GebDat, "yyyy-mm-dd") & "</GeburtsDatum><Telefon>tel: " & aktDC.PrivatTel & "</Telefon><Telefax>fax: " & "</Telefax></Patient>"
   ZsD = ZsD & "<Behandelter><publicID>"
   If Not rDMP.BOF Then
    SuchStr = "<Behandelter><publicID>"
    pos = InStr(Zusatzdaten, SuchStr)
    If pos Then publicID = Mid$(Zusatzdaten, pos + Len(SuchStr), 19)
   End If
   Set rDMP = Nothing
   
   If publicID = "" Then
    For i = 1 To 19
     publicID = publicID & GetRandomNum(0, 9)
    Next i
   End If
   ZsD = ZsD & publicID
   publicID = ""
   ZsD = ZsD & "</publicID><Titel>" & aktDC.Titel & "</Titel><Vorsatz>" & aktDC.NVorsatz & "</Vorsatz><Zusatz>" & aktDC.NVors & "</Zusatz><Vorname>" & aktDC.Vorname & "</Vorname><Nachname>" & aktDC.Nachname & "</Nachname><Anschriftenzusatz>" & aktDC.Anschrzus & "</Anschriftenzusatz><Strasse>" & aktDC.strasse & "</Strasse>"
   ZsD = ZsD & "<Hausnummer>" & aktDC.Hausnr & "</Hausnummer><PLZ>" & aktDC.plz & "</PLZ><Stadt>" & aktDC.ort & "</Stadt><Laenderkennzeichen>" & Trim$(aktDC.Lkz) & "</Laenderkennzeichen><Postfach>" & "</Postfach><PostfachPLZ>" & aktDC.PFPlz & "</PostfachPLZ><PostfachOrt>" & aktDC.PFOrt & "</PostfachOrt>"
   ZsD = ZsD & "<Geschlecht>" & aktDC.geschlecht & "</Geschlecht><GeburtsDatum>" & Format(aktDC.GebDat, "yyyy-mm-dd") & "</GeburtsDatum><Telefon>tel: " & aktDC.PrivatTel & "</Telefon><Telefax>fax: " & "</Telefax></Behandelter>"
   
   myFrag rlanr, "SELECT * FROM lanrpraxis p WHERE lanr = " & auswlanr.Lanr
   ZsD = ZsD & "<Arzt><KVNummer>" & KVNr & "</KVNummer><Titel>" & rlanr!Titel & "</Titel><Vorsatz></Vorsatz><Zusatz></Zusatz><Vorname>" & rlanr!Vorname & "</Vorname><Nachname>" & rlanr!Nachname & "</Nachname><Anschriftenzusatz></Anschriftenzusatz><Strasse>Mittermayerstraße</Strasse><Hausnummer>13</Hausnummer><PLZ>85221</PLZ><Stadt>Dachau</Stadt><Laenderkennzeichen></Laenderkennzeichen><Postfach></Postfach><PostfachPLZ></PostfachPLZ><PostfachOrt></PostfachOrt><Telefon>tel: 08131 / 616 380</Telefon><Telefax>fax: </Telefax></Arzt>"
   Set rlanr = Nothing
   ' folgende Zeile war für beide Ärzte gleich
   ' 30.6.15 Anschriftenzusatz ergaenzt
   ZsD = ZsD & "<Krankenhaus><Abteilung>FA Innere und Allgemeinmedizin (Hausarzt)</Abteilung><Name>Gerald Schade</Name></Krankenhaus><Praxis><Praxisname>Gerald Schade</Praxisname><Anschriftenzusatz></Anschriftenzusatz><Strasse>Mittermayerstraße</Strasse><Hausnummer>13</Hausnummer><PLZ>85221</PLZ><Stadt>Dachau</Stadt><Laenderkennzeichen></Laenderkennzeichen><Postfach></Postfach><PostfachPLZ></PostfachPLZ><PostfachOrt></PostfachOrt><Telefon>tel: 08131 / 616 380</Telefon><Telefax>fax: </Telefax></Praxis><Betriebsstaette><BSNR>" & BSNR & "</BSNR><Betriebsstaettename>Diabetologische Gemeinschaftspraxis</Betriebsstaettename><Anschriftenzusatz></Anschriftenzusatz><Strasse>Mittermayerstraße</Strasse><Hausnummer>13</Hausnummer><PLZ>85221</PLZ><Stadt>Dachau</Stadt><Laenderkennzeichen></Laenderkennzeichen><Postfach></Postfach><PostfachPLZ></PostfachPLZ><PostfachOrt></PostfachOrt><Telefon>tel: 08131 / 616 380</Telefon><Telefax>fax: </Telefax></Betriebsstaette>"
   ZsD = ZsD & "<Krankenversicherung><Typ>GKV</Typ><Name>" & Trim$(ikname) & "</Name><ikPraefix>" & ikprae & "</ikPraefix><IKNummer>" & Trim$(aktDC.IK) & "</IKNummer><KostentraegerAbrechnungsbereich>" & aktDC.KtrAbrB & "</KostentraegerAbrechnungsbereich>"
   Set rik = Nothing
   
' Hier der KV-Bereich
   Dim kvbereich$, rkvber As New ADODB.Recordset
   myFrag rkvber, "SELECT kv FROM plz WHERE plz = " & aktDC.plz
   If rkvber.State = 0 Then
    kvbereich = "71"
   ElseIf Not rkvber.BOF Then
    kvbereich = rkvber!kv
   Else
    kvbereich = "71"
   End If
   Set rkvber = Nothing
   
   ZsD = ZsD & "<KVBereich>" & kvbereich & "</KVBereich>"
   ZsD = ZsD & "<AbrechnungsVKNR>" & aktDC.VKNr & "</AbrechnungsVKNR>"
   ZsD = ZsD & "<Versichertennummer>" & Trim$(aktDC.Versichertennummer) & "</Versichertennummer>"
   ' nehme ich erst mal nicht
'   Dim statuus$(), Stat$, Serg$
'   IF Len(rfal!Status) THEN
'    Split rfal!Status, " ", statuus
'    Stat = statuus(0)
'    IF UBound(statuus) > 0 THEN
'     Serg = statuus(1)
'    END IF
'   END IF
' nach http://de.wikipedia.org/wiki/Versichertenstatus sind ab 1.10.14 </VersichertenstatusKVK> und </VersichertenartMFR> idenstisch
   ZsD = ZsD & "<VersichertenstatusKVK>" & aktDC.KVKStatus & "000</VersichertenstatusKVK>"
   ZsD = ZsD & "<VersichertenartMFR>" & aktDC.KVKStatus & "</VersichertenartMFR>"
   ZsD = ZsD & "<BisDatumderGueltigkeit>" & Year(Now) + 1 & "-" & Right$("0" & Month(Now), 2) & "</BisDatumderGueltigkeit>"
' auskommentiert am 30.6.15, da bei Woltmann Fehler dadurch erzeugt ("Wenn eine Angabe im Feld KVK-Einlesedatum erfolgte, dann darf die Angabe der Statusergänzung nicht fehlen.")
   If ab0921 Then
    ZsD = ZsD & "<KVKEinlesedatum>" & Format(MAXvb(QAnf(ZQuart(DokuDat)), lVorl), "yyyy-mm-dd") & "</KVKEinlesedatum>"
   End If
   ZsD = ZsD & "<WohnsitzLaendercode>D</WohnsitzLaendercode>"
   If ab0921 Then
    ZsD = ZsD & "<BesPersonengruppe>00</BesPersonengruppe>"
   End If
   ' 30.6.15
   ZsD = ZsD & "<DMPKennzeichnung>" & IIf(ab0619 Or ab1118, "0", "") & IIf(BogArtlV = typ2alt Or BogArtlV = typ2neu, "1", "4") & "</DMPKennzeichnung>"
   ZsD = ZsD & "<VersSchutzBeginn>" & Format(aktDC.VschBeg, "yyyy-mm-dd") & "</VersSchutzBeginn>"
   ZsD = ZsD & "</Krankenversicherung><Software><KBVPruefnummer>F0312265</KBVPruefnummer><Name>TurboMed EDV GmbH</Name><Version>" & IIf(ab0423, "23.2.1.5780", IIf(ab0921, "21.3.1.5080", IIf(ab0619, "19.2.1.4087", IIf(ab1118, "18.4.1.3886", IIf(ab0418, "18.2.2.3686", "15.2.1.2774"))))) & "</Version></Software></ZusatzDaten>"
   BDT.FIAdd ZsD
   
   If ab315 Then
    eGFRs = Format(aktDC.eGFR, "000")
    For i = 3 - Len(CStr(eGFRs)) To 2
     BDT.FFAdd "eGFRDM" & DmT & "#" & i
     BDT.FIAdd Mid$(eGFRs, i + 1, 1)
    Next i
   End If
   
   BDT.FFAdd "geboren"
   BDT.FIAdd Format(aktDC.GebDat, "dd.mm.yy")
   BDT.FFAdd "intervallDM" & DmT & "#" & aktDC.NaeUs
   BDT.FIAdd "X"

   BDT.FFAdd "lblPatient"
   BDT.FIAdd aktDC.Vorname & " " & aktDC.Nachname & " geb. am: " & Format(aktDC.GebDat, "dd.mm.yyyy")
   If ab317 Then
    If obErstD Then
     If aktDC.dmpDMSchulWahrg = "j" Then
      BDT.FFAdd "schulungBereitsWahrgenommenDM" & DmT & "#0"
      BDT.FIAdd "X"
     End If ' aktDC.dmpDMSchulWahrg = "j" Then
     If aktDC.dmpHypertSchulWahrg = "j" Then
      BDT.FFAdd "schulungBereitsWahrgenommenDM" & DmT & "#1"
      BDT.FIAdd "X"
     End If ' aktDC.dmpHypertSchulWahrg = "j" Then
     If aktDC.dmpDMSchulWahrg <> "j" Or aktDC.dmpHypertSchulWahrg <> "j" Then
      BDT.FFAdd "schulungBereitsWahrgenommenDM" & DmT & "#2"
      BDT.FIAdd "X"
     End If ' aktDC.dmpDMSchulWahrg <> "j" Or aktDC.dmpHypertSchulWahrg <> "j" Then
    End If ' IF obErstD THEN
   End If ' IF ab317 THEN
   If aktDC.Deform Then ' Fußdeformität
    BDT.FFAdd "weiteresRisikoUlcusDM" & DmT & "#0"
    BDT.FIAdd "X"
   End If ' aktDC.Deform Then ' Fußdeformität
   If aktDC.Hyperk Then ' Hyperkeratose mit Einblutung
    BDT.FFAdd "weiteresRisikoUlcusDM" & DmT & "#1"
    BDT.FIAdd "X"
   End If
   If aktDC.ZnUlcus Then ' Z.n. Ulcus
    BDT.FFAdd "weiteresRisikoUlcusDM" & DmT & "#2"
    BDT.FIAdd "X"
   End If
   If aktDC.ZnAmput Then ' Z.n. Amputation
    BDT.FFAdd "weiteresRisikoUlcusDM" & DmT & "#3"
    BDT.FIAdd "X"
   End If
   BDT.FFAdd "weiteresRisikoUlcusDM" & DmT & "#" & IIf(aktDC.Deform Or aktDC.Hyperk Or aktDC.ZnUlcus Or aktDC.ZnAmput, "4", "5")
   BDT.FIAdd "X"
   BDT.FFAdd "wundInfektionDM" & DmT & "#" & aktDC.Infekt
   BDT.FIAdd "X"
 End Select ' Case BogArtlV
 Call BDT.Schreib(anhängen:=True)
 Exit Sub
 
 #If False Then
 Dim RRstr$
 If aktDC.RRsyst > 0 Then
  RRstr = Format$(aktDC.RRsyst, "000") & Format$(MAX(aktDC.RRdiast, 30), "000")
 Else
  RRstr = "{TAB 6}"
 End If
 Select Case BogArtlV
  Case typ2neu
    SendStr = "%{F2}{TAB}+{TAB}"
    SendStr = SendStr & aktDC.daseit & IIf(aktDC.dspsy, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(aktDC.dspmed, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & "{TAB} {TAB 3} {TAB 2} {TAB 2}"
    For j = 0 To UBound(aktDC.FEn)
     SendStr = SendStr & IIf(aktDC.FEn(j), " ", "") & "{TAB}"
    Next j
    SendStr = SendStr & IIf(aktDC.Tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(aktDC.kgr < 10, "00", IIf(aktDC.kgr < 100, "0", "")) & aktDC.kgr
    SendStr = SendStr & IIf(aktDC.gewi < 10, "00", IIf(aktDC.gewi < 100, "0", "")) & aktDC.gewi
'    SendStr = SendStr & " {TAB 2}" ' altersgemäße körperliche Entwicklung
    SendStr = SendStr & IIf(aktDC.Puls = ndok, " ", "") & "{TAB}" & IIf(aktDC.Puls = unauff, " ", "") & "{TAB}" & IIf(aktDC.Puls = auff, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(aktDC.sens = ndok, " ", "") & "{TAB}" & IIf(aktDC.sens = unauff, " ", "") & "{TAB}" & IIf(aktDC.sens = auff Or aktDC.sens = pathdok, " ", "") & "{TAB 2}"
    SendStr = SendStr & RRstr ' Blutdruck
    SendStr = SendStr & Format$(aktDC.bekHb * 10, "000") & "062"
    SendStr = SendStr & "000" & Format$(aktDC.Crea * 10, "000") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.fußst = ndok, " ", "") & "{TAB}" & IIf(aktDC.fußst = unauff, " ", "") & "{TAB}" & IIf(aktDC.fußst = auff, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.mSei = rE Or aktDC.mSei = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.mSei = li Or aktDC.mSei = gleich, " ", "") & "{TAB}"
    WagSt = Left$(aktDC.mWA, 1)
    ArmSt = Mid$(aktDC.mWA, 2, 1)
    If ArmSt = "" Then
     If WagSt = "0" Then
      If aktDC.Puls = unauff Then ArmSt = "A" Else ArmSt = "C" ' 0 eher nicht entzündet
     ElseIf WagSt > "0" Then
      If aktDC.Puls = unauff Then ArmSt = "B" Else ArmSt = "D"
     Else
      ' WagSt = ""
     End If
    End If
    SendStr = SendStr & IIf(WagSt = "0", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "1", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "2", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "3", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "4", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "5", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "A", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "B", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "C", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "D", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.oap = rE Or aktDC.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.oap = li Or aktDC.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & "{TAB 2}" ' Überweisung/ Einweisung
    SendStr = SendStr & IIf(aktDC.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & Format$(aktDC.hypoZAn, "00") & "00" ' stat Aufenthalte wegen schwerer Hypos
    If aktDC.dspmed Then
     SendStr = SendStr & IIf(aktDC.obGlib = adja, " {TAB 3}", IIf(aktDC.obGlib = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(aktDC.obmetf = adja, " {TAB 3}", IIf(aktDC.obmetf = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(aktDC.obGlucI = adja, " {TAB 3}", IIf(aktDC.obGlucI = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(aktDC.obSHGlin = adja, " {TAB 3}", IIf(aktDC.obSHGlin = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(aktDC.obGlit = adja, " {TAB 3}", IIf(aktDC.obGlit = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(aktDC.obIns, IIf(aktDC.insz > 2, "{TAB} {TAB 2}", "{TAB 2} {TAB}"), " {TAB 3}")
     SendStr = SendStr & IIf(aktDC.obAnal, IIf(aktDC.insz > 2, "{TAB} {TAB 2}", "{TAB 2} {TAB}"), " {TAB 3}")
    End If
    SendStr = SendStr & IIf(aktDC.obHMG, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.obAntihyp, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.obThro, " ", "") & "{TAB}"
    SendStr = SendStr & "{TAB 2} {TAB}" & "{TAB 2} {TAB}" ' Schulungen
    SendStr = SendStr & IIf(aktDC.Tabak, " {TAB 2}", "{TAB}{TAB}") ' 'X' steht bei "nein" schon drin
    SendStr = SendStr & IIf(aktDC.bmi >= 25, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(aktDC.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr = SendStr & IIf(aktDC.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(aktDC.aug = durchg, " {TAB 3}", IIf(aktDC.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr = SendStr & " {TAB 14} {TAB 2}"
   
   Case typ2alt
    SendStr = "%{F3}{TAB}+{TAB}"
    SendStr1a = "(%N){TAB}+{TAB}  "
    SendStr2 = " {TAB 2} {TAB}"
    For j = 1 To UBound(aktDC.FEn)
     SendStr2 = SendStr2 & IIf(aktDC.FEn(j) And Not aktDC.FE(j), " ", "") & "{TAB}"
    Next j
    SendStr2 = SendStr2 & IIf(aktDC.Tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(aktDC.gewi < 10, "00", IIf(aktDC.gewi < 100, "0", "")) & aktDC.gewi
'    sendstr2 = sendstr2 & " {TAB 2}" ' altersgemäße körperliche Entwicklung
    SendStr2 = SendStr2 & IIf(aktDC.Puls = ndok, " ", "") & "{TAB}" & IIf(aktDC.Puls = unauff, " ", "") & "{TAB}" & IIf(aktDC.Puls = auff, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & IIf(aktDC.sens = ndok, " ", "") & "{TAB}" & IIf(aktDC.sens = unauff, " ", "") & "{TAB}" & IIf(aktDC.sens = auff Or aktDC.sens = pathdok, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & RRstr ' Blutdruck
    SendStr2 = SendStr2 & Format$(aktDC.bekHb * 10, "000") & "062"
    SendStr2 = SendStr2 & "000" & Format$(aktDC.Crea * 10, "000") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.fußst = ndok, " ", "") & "{TAB}" & IIf(aktDC.fußst = unauff, " ", "") & "{TAB}" & IIf(aktDC.fußst = auff, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.mSei = rE Or aktDC.mSei = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.mSei = li Or aktDC.mSei = gleich, " ", "") & "{TAB}"
    WagSt = Left$(aktDC.mWA, 1)
    ArmSt = Mid$(aktDC.mWA, 2, 1)
    If ArmSt = "" And WagSt <> "" Then If aktDC.Puls = unauff Then ArmSt = "B" Else ArmSt = "D"
    SendStr2 = SendStr2 & IIf(WagSt = "0", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "1", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "2", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "3", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "4", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "5", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "A", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "B", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "C", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "D", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.oap = rE Or aktDC.oap = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.oap = li Or aktDC.oap = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & "{TAB 2}" ' Überweisung/ Einweisung
    SendStr2 = SendStr2 & IIf(aktDC.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & Format$(aktDC.hypoZAn, "00") & "0" ' stat Aufenthalte wegen schwerer Hypos
    SendStr2 = SendStr2 & IIf(aktDC.dspmed, " {TAB 2}", "{TAB} {TAB}")
    If aktDC.dspmed Then
     SendStr2 = SendStr2 & IIf(aktDC.obGlib = adja, " {TAB 3}", IIf(aktDC.obGlib = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(aktDC.obmetf = adja, " {TAB 3}", IIf(aktDC.obmetf = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(aktDC.obGlucI = adja, " {TAB 3}", IIf(aktDC.obGlucI = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(aktDC.obSHGlin = adja, " {TAB 3}", IIf(aktDC.obSHGlin = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(aktDC.obGlit = adja, " {TAB 3}", IIf(aktDC.obGlit = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(aktDC.obIns, IIf(aktDC.insz > 2, "{TAB} {TAB 2}", "{TAB 2} {TAB}"), " {TAB 3}")
     SendStr2 = SendStr2 & IIf(aktDC.obAnal, IIf(aktDC.insz > 2, "{TAB} {TAB 2}", "{TAB 2} {TAB}"), " {TAB 3}")
    Else
     SendStr2 = SendStr2 & "{TAB 21}"
    End If
    SendStr2 = SendStr2 & IIf(aktDC.obHMG, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.obAntihyp, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.obThro, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & "{TAB 6} {TAB 3} {TAB}" ' Schulungen
    SendStr2 = SendStr2 & IIf(aktDC.Tabak, " {TAB 2}", "{TAB}{TAB}") ' 'X' steht bei "nein" in der Folgedoku noch nicht drin
    SendStr2 = SendStr2 & IIf(aktDC.bmi >= 25, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(aktDC.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr2 = SendStr2 & IIf(aktDC.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(aktDC.aug = durchg, " {TAB 3}", IIf(aktDC.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr2 = SendStr2 & " {TAB 14} {TAB 2}"

  Case typ1neu ' Typ 1 neu
    SendStr = "%(MO){TAB}+{TAB}" & aktDC.daseit & " {TAB 2}"
    SendStr = SendStr & IIf(aktDC.dspsy, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(aktDC.obSchw, " ", "") & "{TAB} {TAB}" ' Schwangerschaft / keine Folgeerkrankungen
    For j = 1 To UBound(aktDC.FEn)
     SendStr = SendStr & IIf(aktDC.FEn(j), " ", "") & "{TAB}"
    Next j
    SendStr = SendStr & IIf(aktDC.Tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(aktDC.kgr < 10, "00", IIf(aktDC.kgr < 100, "0", "")) & aktDC.kgr
    SendStr = SendStr & IIf(aktDC.gewi < 10, "00", IIf(aktDC.gewi < 100, "0", "")) & aktDC.gewi
    SendStr = SendStr & " {TAB 2}" ' altersgemäße körperliche Entwicklung
    SendStr = SendStr & IIf(aktDC.Puls = ndok, " ", "") & "{TAB}" & IIf(aktDC.Puls = unauff, " ", "") & "{TAB}" & IIf(aktDC.Puls = auff, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(aktDC.sens = ndok, " ", "") & "{TAB}" & IIf(aktDC.sens = unauff, " ", "") & "{TAB}" & IIf(aktDC.sens = auff Or aktDC.sens = pathdok, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(aktDC.fußst = ndok, " ", "") & "{TAB}" & IIf(aktDC.fußst = unauff, " ", "") & "{TAB}" & IIf(aktDC.fußst = auff, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.mSei = rE Or aktDC.mSei = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.mSei = li Or aktDC.mSei = gleich, " ", "") & "{TAB}"
    WagSt = Left$(aktDC.mWA, 1)
    ArmSt = Mid$(aktDC.mWA, 2, 1)
    If ArmSt = "" And WagSt <> "" Then If aktDC.Puls = unauff Then ArmSt = "B" Else ArmSt = "D"
    SendStr = SendStr & IIf(WagSt = "0", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "1", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "2", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "3", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "4", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "5", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "A", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "B", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "C", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "D", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.oap = rE Or aktDC.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.oap = li Or aktDC.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & "{TAB 2}" & RRstr  ' Überweisung ausgestellt und Blutdruck
    SendStr = SendStr & Format$(aktDC.bekHb * 10, "000") & "062"
    SendStr = SendStr & "000" & Format$(aktDC.Crea * 10, "000") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.mau = ndok, " ", "") & "{TAB}" & IIf(aktDC.mau = auff, " ", "") & "{TAB}" & IIf(aktDC.mau = unauff, " ", "") & "{TAB 2}"
    SendStr = SendStr & Format$(aktDC.hypoZAn, "00") & "0000"
    SendStr = SendStr & IIf(aktDC.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(aktDC.obIns, IIf(aktDC.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.obAnal, IIf(aktDC.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr = SendStr & IIf(aktDC.obHMG, " {TAB}", "{TAB}")
    SendStr = SendStr & IIf(aktDC.obAntihyp, " {TAB}", "{TAB}")
    SendStr = SendStr & IIf(aktDC.obThro, " {TAB}", "{TAB}")
    SendStr = SendStr & "{TAB 2} {TAB}" & "{TAB 2} {TAB}"
    SendStr = SendStr & IIf(aktDC.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr = SendStr & IIf(aktDC.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(aktDC.aug = durchg, " {TAB 3}", IIf(aktDC.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr = SendStr & " {TAB 16} {TAB 2}"
   
   Case typ1alt
    SendStr = "%(MU){TAB}+{TAB}"
    SendStr1a = "(%N){TAB}+{TAB}  "
    SendStr2 = IIf(aktDC.dspsy, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(aktDC.obSchw, " ", "") & "{TAB} {TAB}" ' Schwangerschaft / keine Folgeerkrankungen
    For j = 1 To UBound(aktDC.FEn)
     SendStr2 = SendStr2 & IIf(aktDC.FEn(j), " ", "") & "{TAB}"
    Next j
    SendStr2 = SendStr2 & IIf(aktDC.Tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & " {TAB 2}" ' altersgerechte Entwicklung
    SendStr2 = SendStr2 & IIf(aktDC.gewi < 10, "00", IIf(aktDC.gewi < 100, "0", "")) & aktDC.gewi
    SendStr2 = SendStr2 & IIf(aktDC.Puls = ndok, " ", "") & "{TAB}" & IIf(aktDC.Puls = unauff, " ", "") & "{TAB}" & IIf(aktDC.Puls = auff, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & IIf(aktDC.sens = ndok, " ", "") & "{TAB}" & IIf(aktDC.sens = unauff, " ", "") & "{TAB}" & IIf(aktDC.sens = auff Or aktDC.sens = pathdok, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & IIf(aktDC.fußst = ndok, " ", "") & "{TAB}" & IIf(aktDC.fußst = unauff, " ", "") & "{TAB}" & IIf(aktDC.fußst = auff, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.mSei = rE Or aktDC.mSei = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.mSei = li Or aktDC.mSei = gleich, " ", "") & "{TAB}"
    WagSt = Left$(aktDC.mWA, 1)
    ArmSt = Mid$(aktDC.mWA, 2, 1)
    If ArmSt = "" And WagSt <> "" Then If aktDC.Puls = unauff Then ArmSt = "B" Else ArmSt = "D"
    SendStr2 = SendStr2 & IIf(WagSt = "0", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "1", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "2", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "3", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "4", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "5", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "A", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "B", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "C", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "D", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.oap = rE Or aktDC.oap = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.oap = li, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & "{TAB 2}" & RRstr   ' Überweisung ausgestellt und Blutdruck
    
    SendStr2 = SendStr2 & Format$(aktDC.bekHb * 10, "000") & "062"
    SendStr2 = SendStr2 & "000" & Format$(aktDC.Crea * 10, "000") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.mau = ndok, " ", "") & "{TAB}" & IIf(aktDC.mau = auff, " ", "") & "{TAB}" & IIf(aktDC.mau = unauff, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & Format$(aktDC.hypoZAn, "0") & "00"
    SendStr2 = SendStr2 & IIf(aktDC.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(aktDC.obIns, IIf(aktDC.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.obAnal, IIf(aktDC.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr2 = SendStr2 & IIf(aktDC.obHMG, " {TAB}", "{TAB}")
    SendStr2 = SendStr2 & IIf(aktDC.obAntihyp, " {TAB}", "{TAB}")
    SendStr2 = SendStr2 & IIf(aktDC.obThro, " {TAB}", "{TAB}")
    SendStr2 = SendStr2 & "{TAB 6} {TAB 3} {TAB}" ' Schulungen
    SendStr2 = SendStr2 & IIf(aktDC.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr2 = SendStr2 & IIf(aktDC.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(aktDC.aug = durchg, " {TAB 3}", IIf(aktDC.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr2 = SendStr2 & " {TAB 16} {TAB 2}"
    
 End Select
 hnd = FensterHandle("TurboMed")
 If hnd <> 0 Then
  On Error Resume Next
  Err.Clear
  AppActivate "TurboMed", True
  If Err.Number <> 0 Then
   MsgBox "Fehler beim Aktivieren von Turbomed"
   Exit Sub
  End If
  On Error GoTo fehler
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
'  hnd = FensterHandle("Hinweis")
'  IF hnd <> 0 THEN
'   ON Error Resume Next
'   AppActivate "Hinweis", True
'   IF Err.Number <> 0 THEN
'    SendKeys "%N", True
'    Pause (Pausenlänge)
'    SendKeys "{ESC}", True
'    Pause (Pausenlänge)
'   END IF
'   ON Error GoTo fehler
'  END IF
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  hnd = FensterHandle("TurboMed - `Kartei")
  If hnd <> 0 Then
   On Error Resume Next
   AppActivate "TurboMed - `Kartei", False
   Pause (Pausenlänge)
   SendKeys "{ESC}", True
   Pause (Pausenlänge)
   hnd = FensterHandle("TurboMed - `Fehler")
   If hnd <> 0 Then
    AppActivate "TurboMed - `Fehler", False
    Pause (Pausenlänge)
    SendKeys "{ENTER}", True
    Pause (Pausenlänge)
    SendKeys "{ESC}", True
    Pause (Pausenlänge)
   End If
   On Error GoTo fehler
  End If
'  MsgBox "Fertig!"
'  Exit Sub
'  SendKeys "{TAB}+{TAB}" & "%p" & "{BS}" & "%p" & "{BS}" & "%p" & "{BS}"
  SendKeys "p" & "{BS}", True
'  SendKeys "{ESC}", True
  Pause (Pausenlänge)
'  SendKeys "p", True
'  Pause (Pausenlänge)
  SendKeys Pat_id & vNS, True
  Pause (Pausenlänge)
  SendKeys "{ENTER}", True
  Pause (1000)
'  SendKeys "{ENTER}", True
'  Pause (Pausenlänge)
'  SendKeys "{ENTER}", True
'  Pause (Pausenlänge)
 Else
  MsgBox "Fehler beim Anwählen des aufgerufenen Turbomed!"
  Exit Sub
 End If
 If SendStr <> "" Then
    SendKeys SendStr, True
    If SendStr1a <> "" Then
     Pause (Pausenlänge)
     SendKeys SendStr1a, True
    End If
    If SendStr2 <> "" Then
     Pause (Pausenlänge)
     SendKeys SendStr2, True
    End If
'   IF 1 = 1 THEN
    Dim dmpdat$, dmpdatop$
    dmpdat = Format$(MIN(QEnd(ZQuart(Now - Verspätung)), Now), "dd.mm.yy")
    dmpdatop = Format$(MIN(QEnd(ZQuart(Now - Verspätung)), Now), "ddmmyyyy")
    Dim DatName$, DNn$
    DatName = pVerz & aktDC.Nachname & " " & aktDC.Vorname & " (PID" & aktDC.Pat_id & "), DMP-Formular vom " & dmpdat & ".pdf"
    DNn = Dir(DatName)
    If DNn <> "" Then Kill DatName
    DatName = REPLACE$(REPLACE$(DatName, "(", "{(}"), ")", "{)}")
    AppActivate "TurboMed", False
    Pause (Pausenlänge)
    SendKeys dmpdatop & "{DEL 8}" & dmpdat
    Pause (1000)
    SendKeys "{TAB}%P", True
    Pause (Pausenlänge)
    SendKeys "%O%(DD)aa{ENTER}", True
    Pause (Pausenlänge)
    SendKeys DatName & "{ENTER}n{ENTER}{F3}{ENTER}", True
'   END IF
 End If
 MFG.SetFocus
 On Error Resume Next
 AppActivate "TurboMed", True
 Exit Sub
#End If
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in domachDMPBogen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' domachDMPBogen

' wird nur für artDiag verwendet
Private Sub MFGrefresh()
 Dim i&
 Dim aicd$
 Dim rs As New ADODB.Recordset
 FNr = 18
 On Error GoTo fehler
 GruSp = 1
 Screen.MousePointer = vbHourglass
 Me.Hide
 Me.MFG.ColWidth(0) = 690
 Me.MFG.ColWidth(1) = 1185
 myFrag rs, "SELECT COUNT(0) ct FROM `diagreihe`", adOpenStatic, DBCn, adLockOptimistic
 MFG.Rows = rs!ct + 30
 Set rs = Nothing
 myFrag rs, "SELECT MAX(zahl) AS mz FROM (SELECT icd, COUNT(0) AS zahl FROM (SELECT d.icd AS icd,diagtext FROM `diagnosen` d LEFT JOIN `diagreihe` r ON d.icd = r.icd GROUP BY d.icd, diagtext) AS innen GROUP BY icd) AS innen;", adOpenStatic, DBCn, adLockOptimistic
 MFG.Cols = rs!mZ + 4
 Me.MFG.ColWidth(2) = 500
 For i = 3 To Me.MFG.Cols
  Me.MFG.ColWidth(i) = 3100
 Next i
 Set rs = Nothing
 myFrag rs, "SELECT d.icd AS icd,g1.gruppe dg1, gi2, diagtext FROM `diagnosen` d LEFT JOIN `diagreihe` r ON d.icd = r.icd LEFT JOIN `diagg1` g1 ON r.gi1 = g1.lfdnr GROUP BY d.icd, diagtext ORDER BY g1.rf ,gi2 DESC,icd", adOpenStatic, DBCn, adLockOptimistic
 aicd = vNS
 MFG.Row = 1
 Do While Not rs.EOF
  If rs!ICD <> aicd Then MFG.Row = MFG.Row + 1: MFG.col = 0: MFG = rs!ICD: MFG.col = GruSp: MFG = rs!dg1: MFG.col = MFG.col + 1: MFG = rs!gi2: aicd = rs!ICD
  MFG.col = MFG.col + 1
  MFG = rs!DiagText
  rs.Move 1
 Loop
 Me.Show
 Screen.MousePointer = vbNormal
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MFGrefresh/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' MFGrefresh

Public Sub csvLesen(Datei$) ' für PatListe Load (Laborwerte)
  On Error GoTo fehler
    Dim zeile$, z2$, Feld$(), f2$(), Tz$, pos%
    Dim angefangen%
    Dim rAf&
    Dim znr&
    Open Datei For Input Access Read As #389
    Do While Not EOF(389)
     znr = znr + 1
     Line Input #389, zeile$
     If Tz = "" Then
      If zeile <> "" Then
       pos = InStr(zeile, ",")
       If pos <> 0 Then
        Tz = ","
       Else
        pos = InStr(zeile, ";")
        If pos <> 0 Then
         Tz = ";"
        End If
       End If
      End If
     End If
     Feld = Split(zeile, Tz)
     If Not angefangen Then
       If UBound(Feld) > 0 Then
        If Feld(0) = "Patient" Then
         angefangen = True
         On Error Resume Next
         labxtb = "labor_xls" & Int(CDbl(Now()) * 1000000)
         myEFrag ("DROP TABLE `" & labxtb & "`")
         On Error GoTo fehler
         myEFrag "CREATE TABLE `" & labxtb & "` (id integer(10) auto_increment key, patient varchar(255), fehlerart varchar(3000))", rAf
'         myEFrag "BEGIN"
         BegTrans
        End If
       End If
     Else
     ' >20 führt zur Aufteilung
       If UBound(Feld) = 1 Then
        Line Input #389, z2
        ReDim Preserve Feld$(2)
        f2 = Split(z2, ",")
        Feld$(1) = Trim$(Feld$(1)) & f2(0)
        If Left$(Feld$(1), 1) = """" Then Feld$(1) = Mid$(Feld$(1), 2, Len(Feld$(1)) - 2)
       End If
       If UBound(Feld) > 1 Then
        InsKorr DBCn, DBCnS, "INSERT INTO `" & labxtb & "`(patient,fehlerart) VALUES('" & UmwfSQL(Feld(0)) & "','" & UmwfSQL(Feld(1)) & "')", rAf
       End If
     End If
    Loop
    Close #389
    ComTrans
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in CSVLesen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' csvLesen

Public Sub ExcelLesen(Datei$) ' für PatListe Load (Laborwerte)
 Const XStra = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
 Const XStrb = ";Extended Properties=""Excel 8.0;HDR=no;IMEX=1"""
 Dim rX As New ADOX.Catalog, sql$, ka%, ke%, runde%, angefangen%, obAnfang%, i&, rAf&
 Dim XCon As New ADODB.Connection
 Dim rEx As New ADODB.Recordset, rs As New ADODB.Recordset
 FNr = 19
 On Error GoTo fehler
 DBCnS = DBCn
 Set XCon = Nothing
 XCon.Open XStra & Datei & XStrb
 Set rX = Nothing
 rX.ActiveConnection = XCon
 rEx.Open "`" & rX.Tables(rX.Tables.COUNT - 1).name & "`", XCon ' Hier Excel, nicht obmysql = 0!
' Dim maxsp&
' maxsp = rEx.Fields.Count
 On Error Resume Next
 labxtb = "labor_xls" & Int(CDbl(Now()) * 1000000)
 myEFrag ("DROP TABLE `" & labxtb & "`")
 On Error GoTo fehler
 myEFrag "CREATE TABLE `" & labxtb & "` (id integer(10) auto_increment key, patient varchar(255), fehlerart varchar(3000))", rAf
 Do While Not rEx.EOF
  If obAnfang Then
    If IsNull(rEx.Fields(1)) Then Exit Do ' 8.1.08
    InsKorr DBCn, DBCnS, "INSERT INTO `" & labxtb & "`(patient,fehlerart) VALUES('" & UmwfSQL(rEx.Fields(0)) & "','" & UmwfSQL(rEx.Fields(1)) & "')", rAf
  ElseIf Not IsNull(rEx.Fields(0)) And Not IsNull(rEx.Fields(IIf(rEx.Fields.COUNT > 1, 1, 0))) And Not IsNull(rEx.Fields(IIf(rEx.Fields.COUNT > 2, 2, 0))) Then
   obAnfang = True
  End If
  runde = runde + 1
  rEx.Move 1
 Loop
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ExcelLesen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' ExcelLesen

' in Form_Load
Sub LabordateiAnzeig(Datei$)
 Dim rs As New ADODB.Recordset, i&, vorDp$, pos&, rs1 As New ADODB.Recordset, j&, DateiDatStr$, sql$
 Dim pid&, vorPID&, gschl$, DateiDat As Date, vorFarbe&, Sp1Farbe&, vorDPneu$, vorEinh$
 Dim fdt$, buch$, altPatient$ ' FileDateTime
 Dim rs2 As New ADODB.Recordset
 Dim pide$  ' PID von Pat. gleichen Namens
 Dim rsgl As New ADODB.Recordset
 
 FNr = 20
 On Error GoTo fehler
 vorFarbe = vbWhite
 With Me.MFG
 .Visible = False
 .Cols = labhwsp + 1
 .TextMatrix(0, namsp) = "Patient"
 .TextMatrix(0, parsp) = "Parameter"
 .TextMatrix(0, wertsp) = "Wert"
 .TextMatrix(0, nbsp) = "Normbereich"
 .TextMatrix(0, Pat_IDSp) = "Pat_ID"
 .TextMatrix(0, einhsp) = "Einheit"
 .TextMatrix(0, vorwsp1) = "Vorwert 1"
 .TextMatrix(0, vorwsp2) = "Vorwert 2"
 .TextMatrix(0, medsp) = "Hinweise"
 .TextMatrix(0, ficdsp) = "ICD"
 .TextMatrix(0, terminsp) = "Termine"
 .TextMatrix(0, labhwsp) = "Laborhinweise"
 Dim rc As New ADODB.Recordset
 myFrag rc, "SELECT COUNT(0) OVER() Zahl" & vbCrLf & _
 ",l.*,s.pat_id obsws, COALESCE(ityp,'') ityp, COALESCE(lp.langtext,l.parameter) lt" & vbCrLf & _
 "FROM labpatel p " & vbCrLf & _
 "LEFT JOIN labpath l ON l.elID=p.id" & vbCrLf & _
 "LEFT JOIN laborparameter lp ON l.Parameter=lp.`Abkü` AND l.Einheit=lp.Einheit " & vbCrLf & _
 "LEFT JOIN sws s ON s.pat_id=l.pat_id AND s.voret>qanf() AND s.voret>now()" & vbCrLf & _
 "LEFT JOIN dtypen dt ON dt.pat_id=l.pat_id" & vbCrLf & _
 "WHERE p.Name='" & FSO.GetFileName(Datei) & "' " & vbCrLf & _
 "ORDER BY l.pat_id,l.name,gruppe,reihe,id"
 '  ",IF(RANK() OVER(PARTITION BY pat_id ORDER BY l.pat_id,gruppe,reihe,id)=1,l.NAME,'') ueName" & vbCrLf & _

 Dim Pat_id&, dtyp$

 If Not rc.BOF And True Then
  .Rows = rc!Zahl + 1
  i = 1
  Do While Not rc.EOF
    If Not IsNull(rc!Pat_id) Then
     Pat_id = rc!Pat_id
     dtyp = rc!ityp
     If Pat_id <> vorPID Then vorFarbe = IIf(vorFarbe = vbWhite, vbGräulich, vbWhite) '&H8000000F&=vbgelblichgrau
     Sp1Farbe = vorFarbe
     If dtyp = "1" Then Sp1Farbe = &HCCCCFF Else If dtyp = "2" Then Sp1Farbe = &HFFCCCC
     'If Not IsNull(rc!obsws) Then If dtyp = "1" Or dtyp = "2" Then Sp1Farbe = vbGoldenRod Else Sp1Farbe = 2139610 ' RGB(218, 165, 32)
     If Not IsNull(rc!obsws) Then Sp1Farbe = vbGoldenRod
     vorPID = Pat_id
     .Row = i
     .col = Pat_IDSp:    .Text = rc!Pat_id:      .CellBackColor = Abs(Sp1Farbe)
     .col = namsp:       .Text = rc!name:        .CellBackColor = rc!namsp
     .col = parsp:       .Text = rc!LT:          .CellBackColor = rc!wertsp
     .col = wertsp:      .Text = rc!Wert:        .CellBackColor = rc!wertsp
     .col = einhsp:       .Text = rc!Einheit:     .CellBackColor = vorFarbe
     .col = vorwsp1:     .Text = rc!vorwert_1:   .CellBackColor = vorFarbe
     .col = vorwsp2:     .Text = rc!vorwert_2:   .CellBackColor = vorFarbe
     .col = nbsp:        .Text = rc!Normbereich: .CellBackColor = vorFarbe
     .col = medsp:       .Text = rc!Hinweise:    .CellBackColor = IIf(rc!hinwsp = vbWhite, vorFarbe, rc!hinwsp)
     .col = ficdsp:      .Text = rc!ficd:        .CellBackColor = IIf(rc!ficdsp = vbWhite, vorFarbe, rc!ficdsp)
     .col = terminsp:    .Text = rc!Termine:     .CellBackColor = IIf(rc!termsp = 0, vbWhite, rc!termsp)
     .col = labhwsp:     .Text = rc!labhinw:     .CellBackColor = vorFarbe
     i = i + 1
    End If
    rc.MoveNext
  Loop
  .ColAlignment(einhsp) = 7
  .ColAlignment(terminsp) = 0
  .ColAlignment(nbsp) = 4
'Constant Value Description
'flexAlignLeftTop 0 The column content is aligned left, top.
'flexAlignLeftCenter 1 Default for strings. The column content is aligned left, center.
'flexAlignLeftBottom 2 The column content is aligned left, bottom.
'flexAlignCenterTop 3 The column content is aligned center, top.
'flexAlignCenterCenter 4 The column content is aligned center, center.
'flexAlignCenterBottom 5 The column content is aligned center, bottom.
'flexAlignRightTop 6 The column content is aligned right, top.
'flexAlignRightCenter 7 Default for numbers. The column content is aligned right, center.
'flexAlignRightBottom 8 The column content is aligned right, bottom.
'flexAlignGeneral 9 The column content is of general alignment. This is "left, center" for strings AND "right, center" for numbers.
 Else
    Call csvLesen(Datei)
'   Call ExcelLesen(Datei)
 DateiDatStr = Datei
 Do
  pos = InStr(DateiDatStr, "\")
  If pos <> 0 Then DateiDatStr = Mid$(DateiDatStr, pos + 1) Else Exit Do
 Loop
 pos = InStr(DateiDatStr, " ")
 If pos <> 0 Then
  j = InStr(pos + 1, DateiDatStr, ".xls")
  If j = 0 Then j = InStr(pos + 1, DateiDatStr, ".csv")
'  i = InStr(pos + 1, DateiDatStr, " ")
'  IF i <> 0 AND i < j THEN j = i  ' Labor Kasse 18.10.18.csv
  DateiDatStr = Trim$(Mid$(DateiDatStr, pos, j - pos))
  For i = 0 To Len(DateiDatStr) - 1
   buch = Mid$(DateiDatStr, i + 1, 1)
   If InStrB("0123456789.,", buch) <> 0 Then Exit For
  Next i
  If (i <> 0) Then DateiDatStr = Mid$(DateiDatStr, i + 1)
  For i = 1 To Len(DateiDatStr)
   buch = Mid$(DateiDatStr, i, 1)
   If InStrB("0123456789.,", buch) = 0 Then
    DateiDatStr = Left$(DateiDatStr, i - 1)
    Exit For
   End If
  Next i
  If IsDate(DateiDatStr) Then
   DateiDat = CDate(DateiDatStr)
'  Else
'   Stop
  End If
 End If
 myFrag rs, "SELECT COUNT(0) ct FROM `" & labxtb & "`"
 If Not rs.BOF Then .Rows = rs!ct + 1
 Set rs = Nothing
 myFrag rs, "SELECT patient, fehlerart,abkü, " & _
 "SUBSTRING_INDEX(SUBSTRING_INDEX(fehlerart,' ',1),':',-1) wert, einheit,langtext, " & _
 "IF(INSTR(fehlerart,'(')=0,'',SUBSTRING_INDEX(SUBSTRING_INDEX(fehlerart,')',1),'(',-1)) nb " & _
 "FROM `" & labxtb & "` lx " & _
 "LEFT JOIN laborparameter lp ON abkü= SUBSTRING_INDEX(lx.fehlerart,':',1) " & _
 "AND einheit=SUBSTRING_INDEX(MID(fehlerart,LENGTH(SUBSTRING_INDEX(fehlerart,' ',1))+2),' (',1) " & _
 "ORDER BY patient, fehlerart", adOpenStatic
 i = 1
 Do While Not rs.EOF
'  .TextMatrix(i, wertsp) = rs!fehlerart
'  pos = InStr(rs!fehlerart, ":")
'  IF pos <> 0 THEN
'   vorDp = Left$(rs!fehlerart, pos - 1)
'   SET rs1 = Nothing
'   myFrag rs1, "SELECT * FROM `laborparameter` WHERE abkü = '" & vorDp & "'"
'   IF Not rs1.EOF THEN
    .TextMatrix(i, parsp) = IIf(IsNull(rs!Langtext), "", rs!Langtext)
    vorDp = IIf(IsNull(rs!Abkü), "", rs!Abkü)
    .TextMatrix(i, wertsp) = rs!Wert ' Mid$(rs!fehlerart, pos + 1)
    .Row = i
    .col = nbsp
    .CellAlignment = 0 'flexAlignLeftTop
'   END IF
'  END IF
  .TextMatrix(i, nbsp) = rs!Nb
  vorEinh = IIf(IsNull(rs!Einheit), "", rs!Einheit)
  .TextMatrix(i, einhsp) = IIf(IsNull(rs!Einheit), "", rs!Einheit)

'  pos = InStr(.TextMatrix(i, wertsp), "(")
'  IF pos <> 0 THEN
'   .TextMatrix(i, nbsp) = Mid$(.TextMatrix(i, wertsp), pos)
'   .TextMatrix(i, wertsp) = Left$(.TextMatrix(i, wertsp), pos - 1)
'  END IF
'  pos = InStr(.TextMatrix(i, wertsp), " ")
'  IF pos <> 0 THEN
'   .TextMatrix(i, einhsp) = Mid$(.TextMatrix(i, wertsp), pos + 1)
'   .TextMatrix(i, wertsp) = Left$(.TextMatrix(i, wertsp), pos - 1)
'  END IF
  
  If altPatient = rs!patient Then
   .TextMatrix(i, namsp) = .TextMatrix(i - 1, namsp)
   .TextMatrix(i, Pat_IDSp) = .TextMatrix(i - 1, Pat_IDSp)
   .TextMatrix(i, terminsp) = .TextMatrix(i - 1, terminsp)
   GoTo andSp
  Else ' altPatient = rs!patient THEN
   pid = 0
   .TextMatrix(i, namsp) = rs!patient
   Set rs1 = Nothing
   sql = "FROM `namen` WHERE TRIM(CONCAT(titel,' ',vorname,' ',nvors, IF(nvors='','',' '),nachname)) = '" & UmwfSQL(rs!patient) & "'"
   myFrag rs1, "SELECT COUNT(0) ct " & sql
   If rs1!ct > 0 Then
    Set rs2 = Nothing
    If rs1!ct > 1 Then ' Wenn mehrere Pat. gleichen Namens
     pide = vNS
     myFrag rs2, "SELECT pat_id " & sql
     If Not rs2.EOF Then
      Do While Not rs2.EOF
       pide = pide & IIf(pide = vNS, vNS, ",") & rs2!Pat_id
       rs2.MoveNext
      Loop
      Set rs2 = Nothing
      fdt = DatFor_k(FileDateTime(Datei)) ' dann rausfinden, zu welchem in letzter Zeit Labor eingetragen
      sql = "FROM `labor2a` l LEFT JOIN `namen` n ON l.pat_id = n.pat_id WHERE l.pat_id IN (" & pide & ") AND zeitpunkt BETWEEN DATE_SUB(" & fdt & ",INTERVAL 7 DAY) AND DATE_ADD(" & fdt & ",INTERVAL 1 DAY) ORDER BY zeitpunkt DESC"
      myFrag rs2, "SELECT COUNT(DISTINCT l.pat_id) ct " & sql
      If rs2!ct = 1 Then
       Set rs2 = Nothing
       myFrag rs2, "SELECT DISTINCT l.pat_id, GebDat, geschlecht " & sql
      Else
       GoTo mehrdeutig
      End If
     End If
    Else ' rs1!ct = 1
     myFrag rs2, "SELECT pat_id, gebdat, geschlecht " & sql
    End If
'    .TextMatrix(i, namsp) = .TextMatrix(i, namsp) & " (" & ROUND((NOW() - rs2!GebDat) * 2.73792574745373E-03, 0) & " a)" ' 1/365,24
    .TextMatrix(i, namsp) = .TextMatrix(i, namsp) & " (" & AlterBei(Now(), rs2!GebDat) & " a)"
    pid = rs2!Pat_id
    gschl = rs2!geschlecht
    .TextMatrix(i, Pat_IDSp) = pid
    Dim rTerm As ADODB.Recordset
    myFrag rTerm, "SELECT COALESCE(TRIM(GROUP_CONCAT(CONCAT(DATE_FORMAT(zp,'%d.%m.%y'),' ',LEFT(raum,3)) ORDER BY zp SEPARATOR '  ')),'') term FROM termine t WHERE zp >= DATE(NOW()) AND pid = " & CStr(pid) & ";", adOpenStatic
    If Not rTerm.BOF And Not IsNull(rTerm!term) Then
     .TextMatrix(i, terminsp) = rTerm!term
    End If ' not rTerm.bof
    If pid <> vorPID Then vorFarbe = IIf(vorFarbe = vbWhite, vbGräulich, vbWhite) '&H8000000F&=vbgelblichgrau
andSp:
    .Row = i
    .col = Pat_IDSp
    .CellBackColor = vorFarbe
    .col = einhsp
    .CellBackColor = vorFarbe
    .col = nbsp
    .CellBackColor = vorFarbe
    vorPID = pid
 ' Vorwerte: 10.7.11: abkü nochmal korrigieren, sonst fehlen welche
'    myFrag rsgl, "SELECT IF(ISNULL(n2.abkü),neu.abkü,n2.abkü) abkü, IF(ISNULL(n2.abkü),neu.einheit,n2.einheit) einheit FROM laborypneu neu " & _
'             "LEFT JOIN laborypgl gl ON neu.id = gl.idypneu " & _
'             "LEFT JOIN laborypneu n2 ON gl.idypbez = n2.id " & _
'             "WHERE neu.abkü = '" & vorDp & "'"
'    IF Not rsgl.EOF THEN
'     vorDPneu = rsgl!Abkü
'     vorEinh = rsgl!Einheit
'    END IF
    Set rsgl = Nothing
    Set rs1 = Nothing
    Dim vw1#, vw2#, aktw#, vw2s$
    vw1 = 0
    vw2 = 0
    vw2s = ""
    aktw = 0
   ' "SELECT * FROM `labor2a` WHERE pat_id = " & Pid & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pid
'   IF Pat_id = 52823 AND vorDp = "K" THEN Stop
    If pid <> 0 Then
'     myFrag rs1, "SELECT zeitpunkt,wert FROM (SELECT zeitpunkt,wert,pat_id,abkü,einheit,nb FROM `labor2a` WHERE pat_id = " & cstr(PID) & " AND abkü = '" & vorDPneu & "' AND einheit ='" & vorEinh & "' " & IIf(DateiDat = 0, vNS, "and zeitpunkt < " & DatFor_k(DateiDat - vorTage)) & " UNION SELECT zeitpunkt,wert,pat_id,abkü,einheit,nb FROM `labor1a` WHERE pat_id = " & cstr(Pid) & " AND abkü = '" & vorDPneu & "' AND einheit ='" & vorEinh & "' " & IIf(DateiDat = 0, vNS, "and zeitpunkt < " & DatFor_k(DateiDat - vorTage)) & ") i GROUP BY pat_id,zeitpunkt ORDER BY zeitpunkt DESC"
'If LEFT(vorDPneu, 1) = "K" AND Pat_id = 36833 THEN Stop
'     myFrag rs1, "SELECT zeitpunkt,wert FROM (SELECT zeitpunkt,wert,pat_id,abkü,einheit,nb FROM `labor2a` WHERE pat_id = " & cstr(PID) & " AND abkü = '" & vorDPneu & "' " & IIf(DateiDat = 0, vNS, "and zeitpunkt < " & DatFor_k(DateiDat - vorTage)) & " UNION SELECT zeitpunkt,wert,pat_id,abkü,einheit,nb FROM `labor1a` WHERE pat_id = " & Pat_id & " AND abkü = '" & vorDPneu & "' " & IIf(DateiDat = 0, vNS, "and zeitpunkt < " & DatFor_k(DateiDat - vorTage)) & ") i GROUP BY pat_id,zeitpunkt ORDER BY zeitpunkt DESC"
'#Const sqllangsam = True
#If sqllangsam Then ' Vergleich bei "Labor 4.5.20 lg tst.csv": 22s vs 8s
     Dim LErg As labtyp
     LErg = LetztLab(pid, rs!Abkü, rs!Einheit, IIf(DateiDat = 0, Now(), DateiDat - vorTage))
     If LErg.WertSg <> vNS And LErg.Zp <> -1 Then
      .TextMatrix(i, vorwsp1) = LErg.WertSg & " (" & Format(LErg.Zp, "d.m.yy") & ")"
      If LErg.WertSg = "" Then
       vw1 = 0
      Else
       Dim vw1s$
       vw1s = REPLACE$(REPLACE$(REPLACE$(LErg.WertSg, ">", ""), "<", ""), ".", ",")
       If IsNumeric(vw1s) Then
        vw1 = CDbl(vw1s)
       Else
        vw1 = 0
       End If
      End If
      LErg = LetztLab(pid, rs!Abkü, rs!Einheit, LErg.Zp)
      If LErg.WertSg <> vNS And LErg.Zp <> -1 Then
       .TextMatrix(i, vorwsp2) = LErg.WertSg & " (" & Format(LErg.Zp, "d.m.yy") & ")"
       vw2s = REPLACE$(LErg.WertSg, ".", ",")
       If IsNumeric(vw2s) Then vw2 = CDbl(vw2s) ' vorletzter Vorwert
      End If
     End If
#Else
     Set rs1 = hollabor(pid, vorDp, IIf(DateiDat = 0, Now(), DateiDat - vorTage), 0, 0, -1, vorEinh)
     If Not rs1.BOF Then
      .TextMatrix(i, vorwsp1) = rs1!Wert & " (" & Format(rs1!Zeitpunkt, "d.m.yy") & ")"
      If rs1!Wert = vNS Then
       vw1 = 0
      Else
       Dim vw1s$
       vw1s = REPLACE$(REPLACE$(REPLACE$(rs1!Wert, ">", ""), "<", ""), ".", ",")
       If IsNumeric(vw1s) Then
        vw1 = CDbl(vw1s)
       Else
        vw1 = 0
       End If
      End If
      rs1.Move 1
      If Not rs1.EOF Then
       .TextMatrix(i, vorwsp2) = rs1!Wert & " (" & Format(rs1!Zeitpunkt, "d.m.yy") & ")"
       If rs1!Wert <> vNS Then
        vw2s = REPLACE$(rs1!Wert, ".", ",")
        If IsNumeric(vw2s) Then vw2 = CDbl(vw2s) ' vorletzter Vorwert
       End If
      End If
     End If
#End If
    End If ' PID <> 0
    Dim aktDC As DMPClass, VorDat As Date
    If IsNumeric(.TextMatrix(i, wertsp)) Then aktw = CDbl(REPLACE$(.TextMatrix(i, wertsp), ".", ",")) Else aktw = 0
    If InStr(1, vorDp, "GFR", vbTextCompare) <> 0 Or InStr(1, vorDp, "GFC", vbTextCompare) <> 0 Or InStr(1, vorDp, "MDRD", vbTextCompare) <> 0 Then
     If aktw < 45 Then
      Call TherAuskunft(CStr(pid), 0, aktDC.insz, VorDat, aktDC.obIns, aktDC.obAnal, aktDC.obGlib, aktDC.obmetf, aktDC.obGlucI, aktDC.obSHGlin, aktDC.obGlit, aktDC.obDpp4, aktDC.obGlp1, aktDC.obSglt2, aktDC.obSonstAD, aktDC.obHMG, aktDC.obAntihyp, aktDC.obACEH, aktDC.obBetabl, aktDC.obThro, aktDC.obOAK, , , , aktDC.obDiur, aktDC.obAT1)
      If aktDC.obmetf Then
       Dim zuwarnen%
       zuwarnen = 0
       Dim metdos! ' Metformindosierung
       If aktw < 30 Then
        zuwarnen = True
       Else ' also aktw zw. 30 und 44
        On Error Resume Next
        ' metdos = myEFrag("SELECT (REPLACE(REPLACE(mo,'½',0.5),'1/2',0.5)+REPLACE(REPLACE(mi,'½',0.5),'1/2',0.5)+REPLACE(REPLACE(nm,'½',0.5),'1/2',0.5)+REPLACE(REPLACE(ab,'½',0.5),'1/2',0.5)+REPLACE(REPLACE(zn,'½',0.5),'1/2',0.5))*nr,i.* FROM " & _
              "(SELECT mp.medikament, zunr(mp.medikament) nr,mo,mi,nm,ab,zn FROM wmedplan mp " & _
              "LEFT JOIN medarten ma ON ma.medikament=mp.medanfang " & _
              "WHERE mp.pat_id = " & CStr(pid) & " AND zeitpunkt=(SELECT MAX(zeitpunkt) FROM wmedplan WHERE pat_id = " & CStr(pid) & " AND zeitpunkt=(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=" & CStr(pid) & ")) " & _
              "and metf " & _
              ") i;").Fields(0)
'        metdos = myEFrag("SELECT ROUND((zubruch(mo)+zubruch(mi)+zubruch(nm)+zubruch(ab)+zubruch(zn))*zunr(mp.medikament)) metdos " & vbCrLf & _
         "FROM wmedplan mp " & vbCrLf & _
         "LEFT JOIN medarten ma ON ma.medikament=mp.medanfang " & vbCrLf & _
         "WHERE mp.pat_id=" & CStr(pid) & " AND " & vbCrLf & _
         "zeitpunkt=(SELECT MAX(zeitpunkt) FROM wmedplan WHERE pat_id=" & CStr(pid) & ") AND metf" & vbCrLf & _
         ";").Fields(0) & vbCrLf
         metdos = myEFrag("SELECT metfdosis(" & CStr(pid) & ")").Fields(0)
        On Error GoTo fehler
        If metdos > 1000 Then zuwarnen = True
       End If ' aktw<30 else
       If zuwarnen Then
        .TextMatrix(i, medsp) = "Metformin" & IIf(metdos, "(" & metdos & " mg)", "") & "!"
        .Row = i
        .col = medsp
        .CellBackColor = vbRed
       End If ' zuwarnen
      End If ' aktdc.obmetf
     ElseIf aktw < 60 And aktw < vw1 * 0.9 Then .TextMatrix(i, medsp) = "GFR-Abfall!": .Row = i: .col = medsp: .CellBackColor = vbRed
     End If ' aktw<45
    Else ' GFR
     Select Case vorDp
      Case "CREAT", "KRE02", "KREA", "KREA02", "KRES"
       If aktw > 1.3 Then
       End If
      Case "CK"
       If aktw > 999 Then .TextMatrix(i, medsp) = "CK hoch": .Row = i: .col = medsp: .CellBackColor = vbRed
      Case "TSH", "TSBF", "TSBL"
       If aktw > 5 Then .TextMatrix(i, medsp) = "V.a. unzur.SD-Hormon-Substitution": .Row = i: .col = medsp: .CellBackColor = vbRed
       If aktw < 0.25 Then .TextMatrix(i, medsp) = "V.a. zu viel SD-Hormon": .Row = i: .col = medsp: .CellBackColor = vbRed
      Case "FT4"
       If aktw > 20 Then .TextMatrix(i, medsp) = "V.a. zu viel SD-Hormon": .Row = i: .col = medsp: .CellBackColor = vbRed
      Case "K"
       If aktw > 5.5 Or aktw < 3.5 Then .TextMatrix(i, medsp) = "V.a. Dyskaliämie": .Row = i: .col = medsp: .CellBackColor = vbRed
      Case "HB"
       If aktw < IIf(gschl = "w", 12, 14) Then
         .Row = i
         .col = medsp
         .TextMatrix(i, medsp) = "Anämie!"
         If .TextMatrix(i, 0) <> "" Then
         Dim radg As New ADODB.Recordset
         ' falls Anämie diagnostiziert
         Set radg = myEFrag("SELECT icd FROM `diagnosen` d WHERE d.pat_id = " & .TextMatrix(i, 0) & " AND d.diagtext LIKE '%anämie%' AND d.diagsicherheit NOT IN ('A','Z') AND d.obdauer<>0") ' AND COALESCE(d.f6010,0)=0
 '        myEFrag "SET GROUP_CONCAT_MAX_LEN = 255"
 '        SET radg = myEFrag("SELECT GROUP_CONCAT(diagtext) diag FROM `diagnosen` d WHERE d.pat_id = " & .TextMatrix(i, 0) & " AND d.diagsicherheit NOT IN ('A','Z') AND COALESCE(d.f6010,0)=0 AND d.obdauer<>0")
 '        IF InStrB(radg!Diag, "Anämie") <> 0 OR InStrB(radg!Diag, "anämie") <> 0 THEN
         If Not radg.BOF Then
          .CellBackColor = Orange
         Else
          .CellBackColor = vbRed
         End If
         End If
       End If
       If aktw < vw1 - 1.5 Then .TextMatrix(i, medsp) = "Hb-Abfall!": .Row = i: .col = medsp: .CellBackColor = vbRed
      Case Else
'       Debug.Print vorDp, i
     End Select ' vorDp
    End If ' GFR else
    Dim Tkz%, gsz%, wdz%, ahz%
    Tkz = 0: gsz = 0: wdz = 0: ahz = 0
    Set rs1 = Nothing
    ' Diane
    myFrag rs1, "SELECT zeitpunkt,art,inhalt FROM `eintraege` WHERE (art IN ('tk','gs','wd','ah') OR inhalt LIKE '%(gs)%' OR inhalt LIKE '%(tk)%' OR inhalt LIKE '%(wd)%' OR inhalt LIKE '%(ah)%') AND pat_id = " & CStr(pid) & " ORDER BY zeitpunkt DESC LIMIT 7"
    If Not rs1.BOF Then
     Do While Not rs1.EOF
      Select Case rs1!art
       Case "tk": Tkz = Tkz + 1
       Case "gs": gsz = gsz + 1
       Case "wd": wdz = wdz + 1
       Case "ah": ahz = ahz + 1
       Case Else: If InStrB(rs1!Inhalt, "(tk)") <> 0 Then Tkz = Tkz + 1 Else If InStrB(rs1!Inhalt, "(gs)") <> 0 Then gsz = gsz + 1 Else If InStrB(rs1!Inhalt, "(wd)") <> 0 Then wdz = wdz + 1 Else If InStrB(rs1!Inhalt, "(ah)") <> 0 Then ahz = ahz + 1
      End Select
      rs1.MoveNext
     Loop
    End If
    If Tkz <> 0 Or gsz <> 0 Or wdz <> 0 Or ahz <> 0 Then
     For j = namsp To wertsp
      .col = j
      If Tkz <> 0 And wdz = 0 And ahz = 0 And gsz = 0 Then
        .CellBackColor = IIf(j = namsp, vbWagnerBlau, HellBlau) ' vbmittelblau, hellblau) ' hellrot, mittigrosa)        ' rötlich
      ElseIf Tkz = 0 And wdz = 0 And ahz = 0 And gsz <> 0 Then
        .CellBackColor = IIf(j = namsp, vbGelb, vbHellGelb)     '
      ElseIf Tkz = 0 And (wdz <> 0) And gsz = 0 Then
        .CellBackColor = IIf(j = namsp, vbWagnerRot, MittigRosa) ' hellrot, mittigrosa)
      ElseIf Tkz = 0 And (ahz <> 0) And gsz = 0 Then
        .CellBackColor = IIf(j = namsp, vbWagnerAHRot, MittigAHRosa) ' hellrot, mittigrosa)
      ElseIf Tkz <> 0 And wdz = 0 And ahz = 0 And gsz <> 0 Then
        .CellBackColor = IIf(j = namsp, vbWagnerGrün, vbHellGrün) ' vbmittelgrün, vbhellgrün) ' orange, hellorange)     '
      ElseIf Tkz = 0 And (wdz <> 0 Or ahz <> 0) And gsz <> 0 Then
        .CellBackColor = IIf(j = namsp, Orange, HellOrange)     '
      ElseIf Tkz <> 0 And (wdz <> 0 Or ahz <> 0) And gsz = 0 Then
        .CellBackColor = IIf(j = namsp, vbMittelLila, vbHellLila)
      ElseIf Tkz <> 0 And (wdz <> 0 Or ahz <> 0) And gsz <> 0 Then
        .CellBackColor = IIf(j = namsp, vbMittelBraun, vbHellBraun)    '
      End If
     Next j
    End If ' Tkz <> 0 OR gsz <> 0 OR wdz <> 0 OR ahz <> 0 THEN
#If False Then
    If Not rs1.BOF Then
     .Row = i
'    .ColSel = .Cols - 1
     For j = namsp To wertsp ' .Cols - 1
      .col = j
      Select Case rs1!art
       Case "tk"
tk:
        .CellBackColor = IIf(j = namsp, HellRot, MittigRosa)        ' rötlich
       Case "gs"
gs:
        .CellBackColor = IIf(j = namsp, vbGelb, vbHellGelb)     'gelblich
       Case Else
        If InStrB(rs1!Inhalt, "(tk)") <> 0 Then GoTo tk Else If InStrB(rs1!Inhalt, "(gs)") <> 0 Then GoTo gs
      End Select
     Next j
'    .TextMatrix(i, ZPSp) = rs1!Zeitpunkt
'   Else
'    Stop
    End If ' not rs1.BOF
#End If
'  Else
'   Stop
   End If ' rs!ct = 0
mehrdeutig:
  End If ' altPatient = rs!patient THEN
  altPatient = rs!patient
  i = i + 1
  rs.Move 1
 Loop
 myEFrag "DROP TABLE IF EXISTS `" & labxtb & "`"
 End If
 .col = 1
 .Row = 1
 .Visible = True
 End With
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabordateiAnzeig/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' LabordateiAnzeig

' aufgerufen in Form_Load
Private Sub RegLaden()
 Dim neuS$, neuB&, i&
 Static angefangen%
 Dim cR As New Registry
 If Not angefangen Then
  On Error Resume Next
  RegPos = RegWurzel & App.EXEName & "\PatListe\" & Me.PLArt
  ReDim MFGCW(MFGLabCols - 1)
  For i = 0 To MFGLabCols - 1
   neuB = -1
   neuB = cR.ReadKey(CStr(i), RegPos, HKEY_CURRENT_USER)
   If neuB <> 0 Then
'    Me.MFG.ColWidth(i) = neuB
    MFGCW(i) = neuB
   End If
  Next i
  neuB = cR.ReadKey("TopRow", RegPos, HKEY_CURRENT_USER)
  MFGTopRow = neuB
  neuB = cR.ReadKey("LeftCol", RegPos, HKEY_CURRENT_USER)
  MFGLeftCol = neuB
  neuB = cR.ReadKey("Row", RegPos, HKEY_CURRENT_USER)
  MFGRow = neuB
  neuB = cR.ReadKey("Col", RegPos, HKEY_CURRENT_USER)
  MFGCol = neuB
  neuB = cR.ReadKey("LabSort", RegPos, HKEY_CURRENT_USER)
  MFGLabSort = neuB
 End If ' not nuranfangs OR not angefangen
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegLaden/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' regladen

Public Sub RegSpeichern()
 Dim cR As New Registry, i&
 On Error GoTo fehler
  RegPos = RegWurzel & App.EXEName & "\PatListe\" & Me.PLArt
  For i = 0 To Me.MFG.Cols - 1
   cR.WriteKey Me.MFG.ColWidth(i), CStr(i), RegPos, HKEY_CURRENT_USER, REG_DWORD
  Next i
 cR.WriteKey Me.MFG.TopRow, "TopRow", RegPos, HKEY_CURRENT_USER, REG_DWORD
 cR.WriteKey Me.MFG.LeftCol, "LeftCol", RegPos, HKEY_CURRENT_USER, REG_DWORD
 cR.WriteKey Me.MFG.Row, "Row", RegPos, HKEY_CURRENT_USER, REG_DWORD
 cR.WriteKey Me.MFG.col, "Col", RegPos, HKEY_CURRENT_USER, REG_DWORD
 cR.WriteKey MFGLabSort, "LabSort", RegPos, HKEY_CURRENT_USER
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegSpeichern/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' RegSpeichern

Private Sub invis(ab%)
 Dim i%
 For i = ab To Me.Command1.COUNT - 1
  Me.Command1(i).Visible = False
 Next i
End Sub ' invis

Private Sub Form_Load()
 Dim sql$, i%, erg$, j&, gehezu&, rAf&, pid$, BhFB As Date, dmpklass&, Tkz%
' #Const obpcol = True
#If obpcol Then
 Dim pCol As New Collection
#Else
 Dim psl As New SortierListe, pstr As SortierString
#End If
 
 Me.Visible = True
 FNr = 1
' syscmd 4, "bitte ca. 10 Sekunden warten ..."
 Me.WindowState = vbMaximized
 MFGTyp = 1
 Li1.Visible = False
 Text1.Visible = False
 cb1.Visible = False
 gehezu = 1
 obtb = -1
 Select Case Me.PLArt
  Case artPat ' DMP hier Liste
   For i = 0 To Me.ErklärungLbl.COUNT - 1
    Me.ErklärungLbl(i).Visible = True
   Next i
   For i = 0 To Me.SpalteLbl.COUNT - 1
    Me.SpalteLbl(i).Visible = True
   Next i
  Case Else ' DMP hier Liste => z.B. DMP-Infos an Hausärzte faxen
   For i = 0 To Me.ErklärungLbl.COUNT - 1
    Me.ErklärungLbl(i).Visible = False
   Next i
   For i = 0 To Me.SpalteLbl.COUNT - 1
    Me.SpalteLbl(i).Visible = False
   Next i
 End Select
 Select Case Me.PLArt
  Case artLPar ' Laborparameter
'   SET MFG.MouseIcon = LoadPicture(uverz & "Programmierung\Icons\Symbole\Hände\E - I411.ico", vbLPSmall, vbLPSmallShell)
   Me.Caption = "Laborparameter zuordnen"
   Call RegLaden
   Call LaborFüll(MFGLabSort)
   Me.Command1(0).Caption = "R&ückwärts"
   Me.Command1(1).Caption = "Vorw&ärts"
   Me.Command1(2).Caption = "Abbru&ch"
   invis 3
   Me.Label1.Visible = False
   Me.Text1.Visible = False
   Call SizeColumns(MFG, Me, True, 5000)
   For i = 0 To Me.MFG.Cols - 1
    If MFGCW(i) <> 0 Then Me.MFG.ColWidth(i) = MFGCW(i)
   Next i
   If MFGTopRow <> 0 Then Me.MFG.TopRow = MFGTopRow
   If MFGRow <> 0 Then Me.MFG.Row = MFGRow
   If MFGLeftCol <> 0 Then Me.MFG.LeftCol = MFGLeftCol
   If MFGCol <> 0 Then Me.MFG.col = MFGCol
   
  Case artDMP ' DMP-Infos an Hausärzte faxen
   Me.Caption = "DMP-Infos an Hausärzte faxen"
   Call DMPFüll
   Me.Command2.Caption = "&DMP-Datei nochmal"
   Me.Command1(0).Caption = "Alle &Markieren"
   Me.Command1(1).Caption = "Alle &Demarkieren"
   Me.Command1(2).Caption = "&Start"
   Me.Command1(3).Caption = "Abbru&ch"
   Me.Command1(4).Caption = "&Gesendete Demark."
   invis 5
   Me.Label1.Visible = False
   Me.Text1.Visible = False
   zähleDMPPat
   Call SizeColumns(MFG, Me, True, 5000)
   
  Case artTabAus ' Tabausgeb
   Call TabAusFüll
   If obPidSp Then
    Me.Command1(0).Caption = "in &Tm anzeigen" ' in TurbomedAnzeigen
    Me.Command1(1).Caption = "&Pat'laufz.anz"  ' Patientenlaufzettel
   End If
   Me.Command1(2 - 2 * (obPidSp + 1)).Caption = "&Nächster" ' Nächster Patient
   Me.Command1(3 - 2 * (obPidSp + 1)).Caption = "&Voriger" ' Nächster Patient
   invis 4 - 2 * (obPidSp + 1)
   Me.Label1.Visible = True
'   Me.Label1.Left = 2000
   Me.Text1.Visible = True
   Call SizeColumns(MFG, Me, True, 5000)
   
  Case artLab ' pathologische Laborwerte ausgeben
   Me.Caption = "Pathologische Laborwerte (" & Me.LabDat & ")"
   Me.Command1(0).Caption = "Datei zu &Kothny verschieben"
   Me.Command1(1).Caption = "Datei zu &Schade verschieben"
   Me.Command1(2).Caption = "Datei zu &Labor alt verschieben"
   Me.Command1(3).Caption = "Datei L&öschen"
   Me.Command1(4).Caption = "in &Tm anzeigen"  ' in TurbomedAnzeigen
   Me.Command1(5).Caption = "&Pat'laufz.anz"  ' Patientenlaufzettel
   Me.Command1(6).Caption = "&Farberklärung"
   Call invis(7)
   Me.Label1.Visible = False
   Call LabordateiAnzeig(Me.LabDat)
   Call SizeColumns(MFG, Me, True, 3200)
   
  Case Else ' artPat, artDiag
   On Error GoTo fehler
   erg = Dir(hVerz & DMP_Import)
   If LenB(erg) <> 0 Then
    Open hVerz & DMP_Import For Input As #388
    Do While Not EOF(388)
     Input #388, erg
     If Mid$(erg, 4, 4) = "3000" Then
#If obpcol Then
      pCol.Add Mid$(erg, 8)
#Else
      Set pstr = New SortierString
      pstr.Stri = Mid$(erg, 8)
      psl.sCAdd pstr, True
#End If
     End If
    Loop ' while not EOF(388)
    Close #388
   End If
#If obpcol Then
#Else
   Dim Fls As Files, Fl As File
   Const gelvz$ = "Papierkorb\eigene Dateien\TMImport"
   If FSO.FolderExists(zVerz & gelvz) Then
    Set Fls = FSO.GetFolder(zVerz & gelvz).Files
    For Each Fl In Fls
'     Debug.Print Fl.name, Fl.DateLastModified
     If Fl.name Like "*" & DMP_Import And Fl.DateLastModified > Now() - 14 Then
      Open Fl.path For Input As #388
      Do While Not EOF(388)
       Input #388, erg
       If Mid$(erg, 4, 4) = "3000" Then
        Set pstr = New SortierString
        pstr.Stri = Mid$(erg, 8)
        psl.sCAdd pstr, True
       End If
      Loop ' while not EOF(388)
      Close #388
     End If
    Next
   End If
#End If
   
   Select Case Me.PLArt
    Case artDiag
     Me.Caption = "Diagnosen sortieren"
     For i = Me.Command1.LBound To Me.Command1.UBound
      On Error Resume Next
      Me.Command1(i).Visible = False
      On Error GoTo fehler
     Next i
     Me.Command1(0).Caption = "&Refresh"
     myEFrag "INSERT INTO `diagreihe`(`icd`) SELECT `icd` FROM `diagnosen` WHERE not `icd` IN (SELECT icd FROM `diagreihe`) GROUP BY icd ", rAf
     Call MFGrefresh
     Me.MFG.Visible = False
     Dim z2&
     Dim gesetzt%
     For i = 2 To MFG.Rows
      Me.MFG.Row = i
      If z2 = 0 Then
       Me.MFG.col = 2
       If Me.MFG.Text = "0" Then
        z2 = Me.MFG.Row
       End If
       Me.MFG.col = 1
      End If
      If LenB(Me.MFG.Text) = 0 Then
       Me.MFG.col = 0
       If Me.MFG.Text <> "" Then
        Me.MFG.TopRow = Me.MFG.Row - 1
        gesetzt = True
       End If
       Me.MFG.LeftCol = 1
       Me.MFG.col = 1
       Exit For
      End If
     Next i
     If Not gesetzt Then
      Me.MFG.Row = MAXvb(z2, 1)
      Me.MFG.TopRow = MAXvb(Me.MFG.Row - 10, 1)
      Me.MFG.col = 2
     End If
     Me.MFG.Visible = True
    Case artHA ' Hausärzte korrigieren
     Dim rHa As New ADODB.Recordset
     sql = "SELECT"
    Case artPat ' DMP hier Liste
     Dim rDPat As New ADODB.Recordset, lfdnr&, cbcol&, k&
     Dim rDok As New ADODB.Recordset
    ' sql = "SELECT n.*, icd, kurzname FROM `namen` n LEFT JOIN `diagnosen` d ON n.pat_id = d.pat_id LEFT JOIN lfaelle f ON n.pat_id = f.pat_id LEFT JOIN `kassenliste` k ON k.ik = f.ik AND k.vknr = f.vknr WHERE notiz LIKE '%DMP hier%' AND (icd LIKE 'E1%' AND NOT icd LIKE 'E15%' AND NOT icd LIKE 'E16%') AND diagsicherheit NOT IN ('Z','A') AND NOT kurzname LIKE 'AOK%' ORDER BY n.pat_id"
'     sql = "SELECT f.bhfb, notiz, obhierdmp, tkz, n.pat_id, n.nachname, n.vorname, icd, kurzname FROM `namen` n LEFT JOIN `diagnosen` d ON n.pat_id = d.pat_id LEFT JOIN (SELECT f.pat_id, f.fid, f.schgr, f.bhfb, f.ik, f.vknr, f.quartal FROM (SELECT * FROM (SELECT BhFB AS mbhfb, bhfe1, Pat_ID AS pid, quartal FROM `faelle` f WHERE bhfb < " & DatFor_k(min(NOW(), qend(ZQuart(NOW() - Verspätung)))) & " ORDER BY pid, bhfb DESC) l GROUP BY pid DESC) l LEFT JOIN `faelle` f ON f.pat_id = l.pid AND f.bhfb = l.mbhfb GROUP BY pat_id) f ON n.pat_id = f.pat_id LEFT JOIN `anamnesebogen` a ON n.pat_id = a.pat_id LEFT JOIN `kassenliste` k ON k.ik = f.ik AND k.vknr = f.vknr WHERE (notiz LIKE '%DMP hier%' OR obhierdmp) AND icd RLIKE '^E1[0-4]' AND diagsicherheit NOT IN ('Z','A') AND NOT ISNULL (f.bhfb) GROUP BY n.pat_id ORDER BY n.pat_id"
    ' sql = "SELECT n.*, icd, kurzname FROM `namen` n LEFT JOIN `diagnosen` d ON n.pat_id = d.pat_id LEFT JOIN lfaelle f ON n.pat_id = f.pat_id LEFT JOIN `kassenliste` k ON k.ik = f.ik AND k.vknr = f.vknr WHERE notiz LIKE '%DMP hier%' AND (icd LIKE 'E1%' AND NOT icd LIKE 'E15%' AND NOT icd LIKE 'E16%') AND diagsicherheit NOT IN ('Z','A') AND kurzname LIKE 'AOK%' ORDER BY n.pat_id"
'     sql = "SELECT f.bhfb, dmpbeg AS notiz, dmpklass, tkz, n.pat_id, n.nachname, n.vorname, icd, kurzname FROM `namen` n LEFT JOIN `diagnosen` d ON n.pat_id = d.pat_id AND icd RLIKE '^E1[0-4]' AND diagsicherheit NOT IN ('Z','A') LEFT JOIN (SELECT f.pat_id, f.fid, f.schgr, f.bhfb, f.ik, f.vknr, f.quartal FROM (SELECT * FROM (SELECT BhFB AS mbhfb, bhfe1, Pat_ID AS pid, quartal FROM `faelle` f WHERE bhfb < " & DatFor_k(min(NOW(), qend(ZQuart(NOW() - Verspätung)))) & " ORDER BY pid, bhfb DESC) l GROUP BY pid DESC) l LEFT JOIN `faelle` f ON f.pat_id = l.pid AND f.bhfb = l.mbhfb GROUP BY pat_id) f ON n.pat_id = f.pat_id LEFT JOIN `anamnesebogen` a ON n.pat_id = a.pat_id LEFT JOIN `kassenliste` k ON k.ik = f.ik AND k.vknr = f.vknr WHERE (notiz LIKE '%DMP hier%' OR dmpklass = 3) AND NOT ISNULL (f.bhfb) GROUP BY n.pat_id ORDER BY n.pat_id"
' ' AND COALESCE(d.f6010,0)=0
 Me.WindowState = vbMinimized
     sql = "SELECT f.bhfb, dmpbeg notiz, dmpklass, tkz, n.pat_id, n.nachname, n.vorname, icd, kurzname " & _
      "FROM `namen` n " & _
      "LEFT JOIN `diagview` d ON n.pat_id = d.pat_id AND gicd RLIKE '^E1[0-4]' " & _
      "LEFT JOIN `faelle` f ON f.pat_id = n.pat_id AND bhfb = (SELECT MAX(bhfb) FROM `faelle` WHERE pat_id = n.pat_id AND bhfb < " & DatFor_k(MINvb(Now(), QEnd(ZQuart(Now() - Verspätung)))) & ") " & _
      "LEFT JOIN `anamnesebogen` a ON n.pat_id = a.pat_id " & _
      "LEFT JOIN `kassenliste` k ON k.ik = f.ik AND k.vknr = f.vknr "
      If Me.ohneTermine Then sql = sql & "LEFT JOIN termine t ON t.pid = n.pat_id AND DATE(t.zp) BETWEEN now() AND qende(NOW()) "
      sql = sql & "WHERE (notiz LIKE '%DMP hier%' OR dmpklass = 3) AND NOT ISNULL (f.bhfb) "
      If Me.ohneTermine Then sql = sql & "and ISNULL(t.zp) AND a.tkz = 0 "
      sql = sql & "GROUP BY n.pat_id ORDER BY n.pat_id"
    
     With MFG
     myFrag rDPat, "SELECT COUNT(0) FROM (" & sql & ") AS innen", adOpenStatic, DBCn, adLockReadOnly
     .Rows = rDPat.Fields(0) + 2
     .Cols = 21
     .Row = 1
     Set rDPat = Nothing
     lfdnr = 0
     myFrag rDPat, sql, adOpenStatic, DBCn, adLockReadOnly
     Do While Not rDPat.EOF
      pid = rDPat!Pat_id
      BhFB = rDPat!BhFB
      dmpklass = rDPat!dmpklass
      Tkz = rDPat!Tkz
      .col = DNrSp
  '   IF rDPat!Pat_id = 2901 THEN Stop
'      IF obhierdmp(rDPat!Notiz) AND rDPat!Tkz = 0 THEN cbcol = dunkelrosa ELSE cbcol = vbWhite
      If Tkz <> 0 Then
       cbcol = vbWhite ' tot
      ElseIf dmpklass = 3 Then
       cbcol = DunkelRosa
      Else
       cbcol = vbCyan ' and.DMP
      End If
      .Text = lfdnr
      .CellBackColor = vbWhite
' in der ersten Spalte die schon exportierten färben
#If obpcol Then
      For k = 1 To pCol.COUNT
       If pCol.Item(k) = pid Then
        .CellBackColor = DunkelRosa
        If cbcol = DunkelRosa Then gehezu = .Row
        Exit For
       End If
      Next k
#Else
      Set pstr = New SortierString
      pstr.Stri = pid
      If psl.SuchItem(pstr) Then
        .CellBackColor = DunkelRosa
        If cbcol = DunkelRosa Then gehezu = .Row
      End If
#End If
      If GruSp = 0 Then GruSp = .col
   '   .CellBackColor = cbcol
      lfdnr = lfdnr + 1
      .col = .col + 1
      If PIDSp = 0 Then PIDSp = .col
      .Text = pid
      .CellBackColor = cbcol
      .col = .col + 1
      If NachNameSp = 0 Then NachNameSp = .col
      .Text = rDPat!Nachname
      .CellBackColor = cbcol
      .col = .col + 1
      .Text = rDPat!Vorname
      .CellBackColor = cbcol
      .col = .col + 1
      .Text = IIf(IsNull(rDPat!kurzname), vNS, rDPat!kurzname)
      .CellBackColor = cbcol
      .col = .col + 1
      If IsNull(rDPat!Notiz) Then
       .Text = vNS
      Else
 '      .Text = replace$(replace$(replace$(rDPat!Notiz, "DMP", ""), &HA, " "), &HD, " ")
        .Text = rDPat!Notiz
      End If
      .CellBackColor = cbcol
      .col = .col + 1
      If ICDSp = 0 Then ICDSp = .col
      .Text = IIf(IsNull(rDPat!ICD), "!-!", rDPat!ICD)
      .CellBackColor = IIf(rDPat!ICD Like "E11*", cbcol, vbYellow)
      .col = .col + 1
      .Text = ZQuart(BhFB)
      If ZQuart(BhFB) = ZQuart(Now() - Verspätung) Then .CellBackColor = DunkelRosa Else .CellBackColor = vbWhite
      Dim begcol%, obraus%, obrot%, obhellrot%, oberst%
      obraus = 0 ' Patient rausgeflogen
      obrot = 0
      obhellrot = 0
      oberst = 0
'      rDok.Open "SELECT `DokuDatum`, `Art`, `ausgedruckt`, `OK`, `exportiert` FROM `dmpreihe` dr WHERE pat_id = " & pid & " AND (dr.Abk LIKE 'eDMPDM%' OR dr.Abk LIKE 'DMPDTYP%') ORDER BY `DokuDatum` DESC", DBCn, adOpenDynamic, adLockReadOnly
      myFrag rDok, "SELECT `DokuDatum`, `Art`, `ausgedruckt`, `OK`, `exportiert` FROM `dmpreihe` dr WHERE pat_id = " & pid & " AND (dr.Abk LIKE 'eDMPDM%' OR dr.Abk LIKE 'DMPDTYP%') ORDER BY `DokuDatum` DESC", adOpenStatic, DBCn, adLockReadOnly
      If Not rDok.BOF Then
       begcol = .col + 1
       If ZQSort(Now() - 120) > ZQSort(rDok!DokuDatum) Then ' Now() - rDok!DokuDatum > 120 THEN ' 21. des übernä Monats
        obraus = True
        begcol = begcol + 1
       End If
       If ZQuart(Now() - Verspätung) <> ZQuart(rDok!DokuDatum) Then
        begcol = begcol + 1
        If Not obraus And ZQuart(BhFB) = ZQuart(Now() - Verspätung) Then ' 2.Bedingung eingefügt: 31.12.15
         obrot = True
         .toolTipText = ZQuart(rDok!DokuDatum) & " <> " & ZQuart(Now() - Verspätung) & " (falsches Quartal)"
        End If
       End If
       For j = begcol To .Cols - 1
        If rDok.EOF Then Exit For
        .col = j
        If VorDokuSp = 0 Then VorDokuSp = .col
        .Text = rDok!art & " " & Format(rDok!DokuDatum, "dd.mm.yy")
        If rDok!Ok And rDok!ausgedruckt Then
         .Text = .Text & " ok"
        ElseIf j = begcol And Not obraus And ZQuart(BhFB) = ZQuart(Now() - Verspätung) Then ' letzte Bedingungen eingefügt 31.12.15
         .toolTipText = "Doku fehlt"
         obrot = True ' Wenn Pat. rausgeflogen, dann fehlt auch aktuelle Erstdoku
        End If
        If rDok!exportiert <> 0 Then
         .Text = .Text & " ex"
         If rDok!art = "ED" Then .CellBackColor = Blau Else .CellBackColor = vbGelblichGrau ' blau / grau
        Else
         If rDok!art = "ED" Then .CellBackColor = HellBlau
        End If
        If rDok!art = "ED" Then oberst = True
        rDok.Move 1
       Next j
       Do While Not rDok.EOF
        If rDok!art = "ED" Then oberst = True
        rDok.Move 1
       Loop
      End If ' not rdok!bof
      Set rDok = Nothing
      If Not oberst And Not obraus And ZQuart(BhFB) = ZQuart(Now() - Verspätung) Then ' letzte Bedingung eingefügt 31.12.15
       .col = begcol
       .CellBackColor = DunkelRot
       obhellrot = True
       .toolTipText = "Doku fehlt"
      End If
      If (obrot Or obhellrot) And dmpklass = 3 And Tkz = 0 Then
       Dim ttt$
       ttt = .toolTipText
       .col = DNamSp
       If obrot Then .CellBackColor = DunkelRot Else .CellBackColor = HellRot
       .toolTipText = ttt
      End If
      .Row = .Row + 1
      rDPat.Move 1
     Loop ' While Not rDPat.EOF
     
    .TextMatrix(0, GruSp) = "Lfdnr"
    .TextMatrix(0, PIDSp) = "Pat_id"
    .TextMatrix(0, NachNameSp) = "Nachname"
    .TextMatrix(0, 3) = "Vorname"
    .TextMatrix(0, 4) = "Kasse"
    .TextMatrix(0, 5) = "DMP"
    .TextMatrix(0, ICDSp) = "ICD"
    .TextMatrix(0, 7) = "lFall"
    .TextMatrix(0, VorDokuSp) = "lDoku"
    Call SizeColumns(MFG, Me)
    .col = NachNameSp
    .Row = gehezu
    End With
    Call mfg_entercell
    Me.Command1(0).Caption = "&Erstelle"
    Me.Command1(1).Caption = "In T&m anzeigen"
    Me.Command1(2).Caption = "-" ' "&In Tm fertigst."
    Me.Command1(3).Caption = "&Suchen"
    Me.Command1(4).Caption = "&Weitersuchen"
    Me.Command1(5).Caption = "&Nächster Gleichartiger"
  ' Me.Command1(6).Caption = "&Typ 1 ED"
  ' Me.Command1(7).Caption = "T&yp 1 FD"
  ' Me.Command1(8).Caption = "Ty&p 2 ED"
  ' Me.Command1(9).Caption = "Typ 2 &FD"
    Me.Command1(6).Caption = "Erstelle &zu Pat_id" ' Dokubeliebig
    Me.Command1(7).Caption = "-" ' "&Fertigst zu Pat_id"
    Me.Command1(8).Caption = "&PLZ"
    Me.Command1(9).Caption = "&HTML bewah"
    Me.Command1(10).Caption = "R&ückg"
    Me.Command1(11).Caption = "Auff&risch" ' Auffrisch
    Me.Command1(12).Caption = "&AlleFS" 'alleFS
    Me.Command1(13).Caption = "a&ktFS" ' aktFS
    Me.Command1(14).Caption = "&Gesamtzus." 'Gesamtzusammenstellung
    invis 15
    Me.Label1.Caption = "'&DMP'=Reiternr.:"
    Me.Text1 = 13 ' DMP-Spalte, DMP-Tab
    On Error GoTo fehler
    
   End Select ' art
  End Select ' me.art

 Dim ctl As Control, lctl As Control
 For i = Me.Command1.LBound To Me.Command1.UBound + 2 ' 0 to 10
  If i <= Me.Command1.UBound Then
   Set ctl = Me.Command1(i)
  ElseIf i = Me.Command1.UBound + 1 Then
   Set ctl = Me.Label1
  Else
   Set ctl = Me.Text1
  End If
  If i <= Me.Command1.UBound Then
   ctl.Width = Me.TextWidth(ctl.Caption) + 200
  ElseIf i = Me.Command1.UBound + 1 Then
   ctl.Width = Me.TextWidth(ctl.Caption)
  Else
   ctl.Width = 300
  End If
  If i = 0 Then
   ctl.Left = 100
  Else
   ctl.Left = lctl.Left + lctl.Width + 50
  End If
  Set lctl = ctl
 Next i

Me.WindowState = vbMaximized
Me.KeyPreview = True
Me.Visible = True
Me.MFG.SetFocus
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Form_LOAD (PatListe)/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
 Exit Sub
End Sub ' Form_load

Private Sub Form_Resize()
 FNr = 21
 On Error Resume Next
 MFG.Height = Me.Height - MFG.Top - 400
 Call Me.TopAusricht
 Exit Sub
End Sub ' Form_Resize

Sub zähleDMPPat()
 Dim i&, sum&
 For i = 1 To PatZuHASL.COUNT
  If PatZuHASL.Item(i).gewählt Then sum = sum + 1
 Next i
 Me.MFG.TextMatrix(0, 0) = sum
End Sub ' zähleDMPPat()

Sub AlleMark(ob%) ' Alle Markieren, alle Demarkieren
 Dim i&, sum&, j&, k&, übrig%, HaslZl&, rAf&
 Dim diffs$, diff&
 Dim rfax As New ADODB.Recordset
 On Error GoTo fehler
 If ob < 2 Then
  For i = 1 To PatZuHASL.COUNT
   PatZuHASL.Item(i).gewählt = IIf(ob, PatZuHASL.Item(i).obDMPInfo, 0)
  Next i
  For i = 1 To HASL.COUNT
   HASL.Item(i).gewählt = IIf(ob, HASL.Item(i).obDMPInfo, 0)
  Next i
 Else      ' Gesendete Demarkieren
  diffs = InputBox("Wie viele Tage zurück sollen die Faxe berücksichtigt werden?", "Rückfrage DMP-DeMarkierung", 60)
  If IsNumeric(diffs) Then diff = CDbl(diffs)
  Me.MFG.MousePointer = vbArrowHourglass ' vbCustom
  myEFrag "UPDATE `faxeinp`.`outa` SET pid = MID(docname,instr(docname,'PID ')+4,locate(', ',docname,instr(docname,'PID '))-instr(docname,'PID ')-4) WHERE instr(docname,'PID ')<>0 AND (ISNULL(pid) OR pid=0) AND MID(docname,instr(docname,'PID ')+4,locate(', ',docname,instr(docname,'PID '))-instr(docname,'PID ')-4)<>''", rAf
  For i = 1 To PatZuHASL.COUNT
   If PatZuHASL.Item(i).gewählt <> 0 Then
'    myFrag rfax, "SELECT docname, o.* FROM faxeinp.`outa` o WHERE docname LIKE '%pid " & PatZuHASL.Item(i).Pat_id & ", dmp-daten%' AND datediff(NOW(),submt) < " & Diff & " ORDER BY submt DESC"
    myFrag rfax, "SELECT docname, o.* FROM `faxeinp`.`outa` o WHERE erfolg<>0 AND pid = " & PatZuHASL.Item(i).Pat_id & " AND docname LIKE '%dmp-daten%' AND datediff(NOW(),submt) < " & diff & " ORDER BY submt DESC"
    If rfax.State <> 0 Then
     If Not rfax.EOF Then
      PatZuHASL.Item(i).gewählt = 0
     End If
     Set rfax = Nothing
    End If
   End If
  Next i
  HaslZl = 0
  For j = 1 To HASL.COUNT
    Do
     HaslZl = HaslZl + 1
     If Me.MFG.TextMatrix(HaslZl, üwnrcol) = HASL.Item(j).ÜWNr Then Exit Do
    Loop
    PatZuHASL.SuchItem HASL.Item(j).ÜWNr
    k = PatZuHASL.sCa
    übrig = 0
    Do
       If k > PatZuHASL.COUNT Then Exit Do
       If PatZuHASL.Item(k).ÜWNr <> HASL.Item(j).ÜWNr Then Exit Do
       If PatZuHASL.Item(k).gewählt <> 0 Then übrig = True
       k = k + 1
    Loop
    If übrig = 0 Then
     HASL.Item(j).gewählt = 0
     Me.MFG.TextMatrix(HaslZl, gewcol) = vNS
    End If
  Next j
 End If
 If ob < 2 Then
  For i = 1 To Me.MFG.Rows - 1
   If LenB(Me.MFG.TextMatrix(i, PlusCol)) = 0 Then Me.MFG.TextMatrix(i, gew2col) = IIf(IIf(ob, HASL.Item(i).obDMPInfo, 0), "X", vNS) Else Me.MFG.TextMatrix(i, gewcol) = IIf(IIf(ob, HASL.Item(i).obDMPInfo, 0), "X", vNS)
  Next i
 End If
 zähleDMPPat
 Me.MFG.MousePointer = flexDefault
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AlleMark/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
 Exit Sub
End Sub ' AlleMark(ob%)

Public Sub MFG_Click()
 Dim ctrl As Control, Fd0$, Fd1$, i%, rs As New ADODB.Recordset, j&, k&
 FNr = 22
 On Error GoTo fehler
 If cRow(MFGTyp) = cRowSel(MFGTyp) Then
  cRow(MFGTyp) = Me.MFG.Row
  cRowSel(MFGTyp) = Me.MFG.RowSel
 End If
 ccol(MFGTyp) = Me.MFG.col
 Select Case Me.PLArt
  Case artDMP
'   Stop
   Select Case ccol(MFGTyp)
    Case gew2col
     j = Me.MFG.TextMatrix(Me.MFG.Row, nrcol)
     PatZuHASL.Item(j).gewählt = Not PatZuHASL.Item(j).gewählt
     Me.MFG.Text = IIf(PatZuHASL.Item(j).gewählt, "X", vNS)
     k = Me.MFG.Row
     Do
      k = k - 1
      If LenB(Me.MFG.TextMatrix(k, PlusCol)) <> 0 Then Exit Do
     Loop
     j = Me.MFG.TextMatrix(k, nrcol)
     Me.MFG.TextMatrix(k, gewcol) = "(X)"
     HASL.Item(j).gewählt = -2
     Call zähleDMPPat
    Case gewcol
     j = Me.MFG.TextMatrix(Me.MFG.Row, nrcol)
     If Me.MFG.TextMatrix(Me.MFG.Row, gewcol + 1) = vNS Then ' doch Tochterzeile
      PatZuHASL.Item(j).gewählt = Not PatZuHASL.Item(j).gewählt
      Me.MFG.TextMatrix(Me.MFG.Row, gew2col) = IIf(PatZuHASL.Item(j).gewählt, "X", vNS)
      k = Me.MFG.Row
      Do
       k = k - 1
       If LenB(Me.MFG.TextMatrix(k, PlusCol)) <> 0 Then Exit Do
      Loop
      j = Me.MFG.TextMatrix(k, nrcol)
      Me.MFG.TextMatrix(k, gewcol) = "(X)"
      HASL.Item(j).gewählt = -2
     Else                                                ' Zeile ersten Ranges
      HASL.Item(j).gewählt = IIf(HASL.Item(j).gewählt <> -1, -1, 0)
      Me.MFG.Text = IIf(HASL.Item(j).gewählt, "X", "")
      PatZuHASL.SuchItem Me.MFG.TextMatrix(Me.MFG.Row, üwnrcol)
      k = PatZuHASL.sCa
      Do
       If k > PatZuHASL.COUNT Then Exit Do
       If PatZuHASL.Item(k).ÜWNr <> Me.MFG.TextMatrix(Me.MFG.Row, üwnrcol) Then Exit Do
       PatZuHASL.Item(k).gewählt = HASL.Item(j).gewählt
       k = k + 1
      Loop
      If Me.MFG.TextMatrix(Me.MFG.Row, PlusCol) = "-" Then
       For j = 1 To Me.MFG.TextMatrix(Me.MFG.Row, zahlcol)
        k = Me.MFG.TextMatrix(Me.MFG.Row + j, nrcol)
        Me.MFG.TextMatrix(Me.MFG.Row + j, gew2col) = IIf(PatZuHASL.Item(k).gewählt, "X", "")
       Next j
      End If ' Me.MFG.TextMatrix(Me.MFG.Row, PlusCol) = "-" THEN
     End If ' Me.MFG.TextMatrix(Me.MFG.Row, gewcol + 1) = vNS THEN ' doch Tochterzeile
     Call zähleDMPPat
    Case PlusCol
     Select Case Me.MFG.Text
      Case "+"
       j = Me.MFG.Row
       PatZuHASL.SuchItem Me.MFG.TextMatrix(j, üwnrcol)
       k = PatZuHASL.sCa
       Do
        If k > PatZuHASL.COUNT Then Exit Do
        If PatZuHASL.Item(k).ÜWNr <> Me.MFG.TextMatrix(Me.MFG.Row, üwnrcol) Then Exit Do
        Me.MFG.AddItem k & String(PlusCol + 1, vbTab) & PatZuHASL.Item(k).ÜWNr & vbTab & IIf(PatZuHASL.Item(k).gewählt, "X", "") & vbTab & PatZuHASL.Item(k).name & vbTab & PatZuHASL.Item(k).Pat_id & vbTab, j + 1
        k = k + 1
        j = j + 1
       Loop
       Me.MFG.Text = "-"
       altrow = Me.MFG.Row
       altCol = Me.MFG.col
       Me.MFG.col = nrcol
       For j = 1 To Me.MFG.TextMatrix(altrow, zahlcol)
        Me.MFG.Row = altrow + j
        Me.MFG.CellBackColor = vbWhite
       Next j
       Me.MFG.Row = altrow
       Me.MFG.col = altCol
      Case "-"
       For j = 1 To Me.MFG.TextMatrix(Me.MFG.Row, zahlcol)
        Me.MFG.RemoveItem Me.MFG.Row + 1
       Next j
       Me.MFG.Text = "+"
     End Select
     Call SizeColumns(MFG, Me, True, 5000, PlusCol + 1)
   End Select
  Case artDiag
   If MFG.col = GruSp Or MFG.col = GruSp + 1 Then
     obtb = 0
   End If
  Case artPat ' DMP hier Liste
   If Me.MFG.col = zahlcol Then
    Me.MFG.CellFontBold = Not Me.MFG.CellFontBold
    If MFG.CellBackColor = vbYellow Then
       MFG.CellBackColor = vbWhite
       GesZl = GesZl - 1
       Dim iz&
       For iz = 1 To GesColl.COUNT
        If GesColl(iz) = Me.MFG.Text Then
         GesColl.Remove (iz)
         Exit For
        End If
       Next iz
    Else
       MFG.CellBackColor = vbYellow
       GesZl = GesZl + 1
       GesColl.Add Me.MFG.Text
    End If
    Me.Command1(14).Caption = Me.GesZl & " &Ges"
   End If
  Case artLab
   If IsNumeric(Me.MFG.TextMatrix(Me.MFG.MouseRow, Pat_IDSp)) Then
    Select Case Me.MFG.MouseCol
     Case 0: Call FertigStellen(Me.MFG.MouseRow, True) ' in Turbomed anzeigen
     Case 1: Call dodoPLZ(Me.MFG.TextMatrix(Me.MFG.MouseRow, Pat_IDSp), plzVz, Now, Now - Int(Now), True, , , True)
    End Select
   End If
 End Select
 If obtb = -1 Then
  cRow(MFGTyp) = Me.MFG.Row
  cRowSel(MFGTyp) = Me.MFG.RowSel
 Else ' IF obtb = 0 THEN
  If MFG.col = GruSp Then
   Set ctrl = Me.Li1
   ctrl.Clear
   myFrag rs, "SELECT * FROM `diagg1` ORDER BY rf", adOpenStatic, DBCn, adLockOptimistic
   Do While Not rs.EOF
    ctrl.AddItem (rs!gruppe)
    rs.Move 1
   Loop
   If AnfCode <> 0 Then
    For i = 0 To Li1.ListCount - 1
     If Li1.List(i) <> "" Then
      If Asc(LCase$(Left$(Li1.List(i), 1))) = AnfCode Then
       Cstumm = True
       Li1.ListIndex = i
       Cstumm = False
       Exit For
      End If
     End If
    Next i
    AnfCode = 0
   End If
   ctrl.Height = ctrl.ListCount * 255
   Cstumm = True
   ctrl.Text = Me.MFG.Text
   Cstumm = False
  ElseIf MFG.col = GruSp + 1 Then
   If KeyC = 13 Then
    Call Text1_Fertig
    Exit Sub
   Else
    Set ctrl = Me.Text1
    Cstumm = True
    ctrl.Text = Chr(AnfCode)
    Cstumm = False
   End If
  End If
  ctrl.Left = Me.MFG.CellLeft + Me.MFG.Left
  ctrl.Width = 3000
  ctrl.Top = Me.MFG.CellTop + Me.MFG.Top
  ctrl.Visible = True
  Call MFG_leavecell
  ctrl.SetFocus
  On Error Resume Next
  ctrl.SelStart = Len(ctrl.Text)
  obtb = -1
 End If ' IF obtb = -1 else
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MFG_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
 Exit Sub
End Sub ' MFG_Click

Public Sub mfg_entercell()
 FNr = 24
 On Error GoTo fehler
 Select Case Me.PLArt
  Case artLPar
   If Me.MFG.col = 2 Then
    altnoenter = noenter
    noenter = True
   End If
 End Select
 If noenter = 0 Then
  If fgespei(MFGTyp) = 0 Then
   If Not (Me.PLArt = artPat And Me.MFG.col = zahlcol) Then
    altFarbe(MFGTyp) = Me.MFG.CellBackColor
    Me.MFG.CellBackColor = vbYellow
   End If
   fgespei(MFGTyp) = -1
  End If
 End If
 Select Case Me.PLArt
  Case artLPar
   If Me.MFG.col = 2 Then
    noenter = altnoenter
   End If
 End Select
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in mfg_entercell/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
 Exit Sub
End Sub ' mfg_entercell

Private Sub MFG_leavecell()
 FNr = 25
 Select Case Me.PLArt
  Case artLPar
   If Me.MFG.col = 2 Then
    altnoenter = noenter
    noenter = True
   End If
 End Select
 If noenter = 0 Then
  If altFarbe(MFGTyp) <> 0 Then ' am Anfang
   If Not (PLArt = artPat And Me.MFG.col = zahlcol) Then
    Me.MFG.CellBackColor = IIf(Me.MFG.Row = 0, vbActiveBorder, altFarbe(MFGTyp))
   End If
  End If
'  IF MFG.CellBackColor = 0 THEN
  fgespei(MFGTyp) = 0
 End If
 Select Case Me.PLArt
  Case artLPar
   If Me.MFG.col = 2 Then
    noenter = altnoenter
   End If
 End Select
End Sub ' MFG_leavecell

Private Sub MFG_KeyDown(KeyCode As Integer, Shift As Integer)
 FNr = 26
 If Me.hlese Is Nothing Then Set Me.hlese = lies
  Call Key(KeyCode, Shift, Me, Me.hlese.MyDB.name)
End Sub ' MFG_KeyDown

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 FNr = 27
 If Me.hlese Is Nothing Then Set Me.hlese = lies
 KeyC = KeyCode
 Call Key(KeyCode, Shift, Me, Me.hlese.MyDB.name)
End Sub ' Form_KeyDown

Private Sub Command1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 FNr = 28
 If Me.hlese Is Nothing Then Set Me.hlese = lies
 Call Key(KeyCode, Shift, Me, Me.hlese.MyDB.name)
End Sub ' Command1_KeyDown

Private Sub MFG_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim rs As New ADODB.Recordset, obbez%
 On Error GoTo fehler
 Select Case Me.PLArt
  Case artLPar
   altrow = Me.MFG.Row
   altCol = Me.MFG.col
   obbez = 0
   Select Case Me.MFG.MouseCol
    Case parsp
     If Me.MFG.TextMatrix(Me.MFG.MouseRow, 8) = "x" Then
      myFrag rs, "SELECT bez1.abkü, bez2.abkü, CONCAT(GROUP_CONCAT(' ',IF(ISNULL(bez1.abkü),''," & vbCrLf & _
                     "CONCAT(bez1.abkü, ' [',bez1.einheit, ']'))),', '," & vbCrLf & _
                     "GROUP_CONCAT(' ',IF(ISNULL(bez2.abkü),'',CONCAT(bez2.abkü, ' [',bez2.einheit, ']')))) Zuge " & vbCrLf & _
                "FROM `laborypneu` n " & vbCrLf & _
                "LEFT JOIN `laborypgl` gl ON n.id = gl.idypbez " & vbCrLf & _
                "LEFT JOIN `laborypneu` bez1 ON gl.idypneu = bez1.id " & vbCrLf & _
                "LEFT JOIN `laborparameter` bez2 ON gl.idpara = bez2.id " & vbCrLf & _
                "WHERE idypbez = " & IDS(1, MFG.MouseRow), adOpenStatic
      If Not rs.BOF Then
       If Not IsNull(rs!zuge) Then
        obbez = True
       End If
      End If
     End If
     If obbez Then
      Me.MFG.toolTipText = rs!zuge
     Else
      Me.MFG.toolTipText = MFG.MouseRow & " " & IDS(0, MFG.MouseRow) & " " & IDS(1, MFG.MouseRow)
     End If
   End Select
 
 Case artLab
  With Me.MFG
  .toolTipText = vNS
  If .toolTipText = "" Then
  If IsNumeric(.TextMatrix(.MouseRow, Pat_IDSp)) Then
   Select Case .MouseCol
     Case namsp
'      myFrag rs, "SELECT GROUP_CONCAT(CONCAT_WS(' ',DATE_FORMAT(diagdatum,'%d.%m.%y'), CONCAT('(',diagsicherheit,')'), diagtext, icd),char(13,10)) diag FROM `diagnosen` WHERE pat_id = " & Pat_id & " AND diagsicherheit <> 'A' ORDER BY DATE(diagdatum)"
vorher:
      myFrag rs, "SELECT GROUP_CONCAT(CONCAT_WS(' ',diagsicherheit, diagtext)) diag FROM `diagnosen` d WHERE d.pat_id = " & .TextMatrix(MFG.MouseRow, Pat_IDSp) & " AND d.diagsicherheit <> 'A' ORDER BY DATE(d.diagdatum)", adOpenStatic ' AND COALESCE(d.f6010,0)=0
      If Not rs.BOF Then
       If Not IsNull(rs!Diag) Then
        .toolTipText = rs!Diag
       End If
      End If
     Case parsp
      Dim pid$, hsql$
      pid = .TextMatrix(MFG.MouseRow, Pat_IDSp)
      hsql = "SELECT CONCAT_WS(' ',DATE_FORMAT(zeitpunkt,'%d.%m.%y'),DATE_FORMAT(datum,'%d.%m.%y'),GROUP_CONCAT(medikament)) med FROM `wmedplan` " & vbCrLf & _
      "WHERE pat_id = " & pid & " AND zeitpunkt = " & vbCrLf & _
      "(SELECT MAX(zeitpunkt) FROM `wmedplan` WHERE pat_id = " & pid & ")"
      myFrag rs, hsql
      If Not rs.BOF Then
       .toolTipText = rs.Fields(0)
      End If
     Case medsp
      .toolTipText = .TextMatrix(MFG.MouseRow, labhwsp) ' Zeilenumbruch statt '|' geht nicht nicht chr(10), chr(13), chr(10) & chr(13) oder chr(13) & chr(10)
     Case ficdsp
     ' s. labpath
      Select Case .TextMatrix(MFG.MouseRow, ficdsp)
       Case ""
       Case "D64.9": .toolTipText = "Anämie"
       Case "E05.0": .toolTipText = "M.Basedow"
       Case "E06.3": .toolTipText = "Autoimmunthyreoiditis"
       Case "E53.8": .toolTipText = "Vit.B12-Mangel"
       Case "E21.1":
        .toolTipText = "sek. Hyperparathyreoidismus"
       Case "E55.9": .toolTipText = "Vit.D-Mangel"
       Case "E78.0": .toolTipText = "Hypercholesterinämie"
       Case "E79.0": .toolTipText = "Hyperurikämie"
       Case "I50.19": .toolTipText = "Herzinsuffizienz"
       Case "N08.3": .toolTipText = "diabet. Nephropathie"
       Case "N18.3": .toolTipText = "Niereninsuffizienz Grad 3"
       Case "N18.4": .toolTipText = "Niereninsuffizienz Grad 4"
       Case "N18.5": .toolTipText = "Niereninsuffizienz Grad 5"
       
      End Select
   End Select
   End If
  End If ' isNumeric(.TextMatrix ...
  End With
 End Select ' case me.mfg.mousecol
 Exit Sub
fehler:
 Dim FNr&, FLastDLLError&, FSource$, FDescr$
 FNr = Err.Number
 FLastDLLError = Err.LastDllError
 FSource = Err.source
 FDescr = Err.Description
 If FNr = 3704 Then ' für ein geschlossene Objekt nicht zugelassen
  DBCnOpen
  Set rs = Nothing
  Resume vorher
 End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), vNS, CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in MouseMove/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' MFG_MouseMove

Private Sub MFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim MR&, MC&, altids&, altherk$, i&
 altX = x
 altY = Y
 Select Case Me.PLArt
  Case artLPar
   MR = Me.MFG.MouseRow
   MC = Me.MFG.MouseCol
   If MR = 0 Then
    altids = IDS(1, altrow)
    altherk = Me.MFG.TextMatrix(altrow, 8)
    Select Case MC
     Case 6
      MFGLabSort = True
     Case 2
      MFGLabSort = 0
    End Select
    Call LaborFüll(MFGLabSort)
    For i = 0 To UBound(IDS, 2)
     If IDS(1, i) = altids Then
      If Me.MFG.TextMatrix(i, 8) = altherk Then
       Me.MFG.Row = i
       Me.MFG.col = 2
       Me.MFG.TopRow = MAXvb(i - 10, 1)
       Exit For
      End If
     End If
    Next i
   Else ' MR = 0 Then
    QuellZ = MR
    If Me.MFG.TextMatrix(MR, 8) = "x" Then
     cb = IDS(1, MR)
     Me.MFG.MousePointer = vbArrowHourglass ' vbCustom
    Else
     cb = 0
    End If
   End If ' MR = 0 Then else
 End Select ' Case Me.PLArt
End Sub ' MFG_MouseDown

'vorausgehend:
'SET FOREIGN_KEY_CHECKS=0;
'INSERT INTO laborypgl(idypneu,idpara,idypbez,ergaenzt)
'SELECT xpneuzuypneu(x.idypneu),idpara,xpneuzuypneu(x.idypbez),ergaenzt
'from laborypgl x
'Where
'(NOT ISNULL(xpneuzuypneu(x.idypneu)) OR ISNULL(idypneu))
'and
'(NOT ISNULL(xpneuzuypneu(x.idypbez)) OR ISNULL(idypbez));
'SET FOREIGN_KEY_CHECKS=1;

Private Sub MFG_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim sqls$, sqld$, sqli$, MR&, rs As New ADODB.Recordset, i&, rAf&
 Me.MFG.MousePointer = flexDefault
 If x <> altX Or Y <> altY Then
 Select Case Me.PLArt
  Case artLPar
' Wird Parameter Pkind mit Parameter Pelter verknüpft, so muß
' 1) vorherige Verknüpfung Pelteralt zu Pkind gelöscht werden
' 2) vorherige Verknüfung Pelterelter zu Pelter gelöscht werden
' 3) vorherige Verknüfung Pkindelter zu Pkindkind überbrückt werden
   MR = Me.MFG.MouseRow
   If Me.MFG.col <> 2 Or Me.MFG.TextMatrix(MR, 8) <> "x" Then Me.MFG.CellBackColor = vbWhite
   Me.MFG.Row = Me.MFG.MouseRow
   Me.MFG.col = Me.MFG.MouseCol
   Select Case Me.MFG.TextMatrix(MR, 8)
    Case "x" 'laborxpneu
     sqls = "SELECT `idypbez` FROM `laborypgl` WHERE `idypneu` = " & IDS(1, MR)
     sqld = "DELETE FROM `laborypgl` WHERE idypneu = " & IDS(1, MR)
     sqli = "INSERT INTO `laborypgl`(`idypneu`,`idypbez`,`ergaenzt`) VALUES(" & IDS(1, MR) & "," & cb & "," & DatFor_k(Now) & ")"
    Case "p" ' laborparameter
     sqls = "SELECT `idypbez` FROM `laborypgl` WHERE `idpara` = " & IDS(1, MR)
     sqld = "DELETE FROM `laborypgl` WHERE `idpara` = " & IDS(1, MR)
     sqli = "INSERT INTO `laborypgl`(`idpara`,`idypbez`,`ergaenzt`) VALUES(" & IDS(1, MR) & "," & cb & "," & DatFor_k(Now) & ")"
    Case Else
     Exit Sub ' z.B. Kopfzeile
   End Select
   If IDS(0, MR) <> 0 Then
' Verknüpfungen 1) löschen
     myEFrag sqld, rAf
' Farben zu alternativen Verknüfungen nach 1) korrigieren
     For i = 0 To UBound(IDS, 2)
      If IDS(1, i) = IDS(0, MR) Then
       Call ZeigeZahl(i)
       Exit For
      End If
     Next i
   End If
   If cb <> 0 Then
    If QuellZ <> MR Then
' Verknüfung einfügen
     myEFrag "SET FOREIGN_KEY_CHECKS=0"
     myEFrag sqli
     myEFrag "SET FOREIGN_KEY_CHECKS=1"
' Verküpfung nach 2) suchen
     Set rs = Nothing
     myFrag rs, "SELECT `idypbez` FROM `laborypgl` WHERE `idypneu` = " & cb
     If Not rs.EOF Then
' Verknüpfung nach 2) löschen
      myEFrag "DELETE FROM `laborypgl` WHERE `idypneu` = " & cb, rAf
      If rAf <> 1 Then
       MsgBox "Fehler in MFG.Mouseup nach DELETE FROM `laborypgl` WHERE `idypneu` = " & cb
       Stop
      End If
      Me.MFG.TextMatrix(QuellZ, 1) = vNS
      IDS(0, QuellZ) = 0
' Farben zu alternativen Verknüfungen nach 2) korrigieren
      For i = 0 To UBound(IDS, 2)
       If IDS(1, i) = rs!idypbez Then
        Call ZeigeZahl(i)
        Exit For
       End If
      Next i
'      Me.MFG.TextMatrix(Cb, 1) = vNS
     End If
     sql = "SELECT CONCAT(abkü,' [',einheit,']') wert FROM laborypneu WHERE id = " & cb
     Set rs = Nothing
     myFrag rs, sql
     If rs.BOF Then
      Me.MFG.TextMatrix(MR, 1) = vNS
      IDS(0, MR) = 0
     Else
' Verknüfung nach 3) überbrücken
      If Me.MFG.TextMatrix(MR, 8) = "x" Then
       myEFrag "UPDATE `laborypgl` SET `idypbez` = " & cb & " WHERE `idypbez` = " & IDS(1, MR), rAf
       For i = 0 To UBound(IDS, 2)
        If IDS(0, i) = IDS(1, MR) Then
         Me.MFG.TextMatrix(i, 1) = rs!Wert
         IDS(0, i) = cb
        End If
       Next
       Call ZeigeZahl(MR)
      End If
' Verknüpfung markieren
      Me.MFG.TextMatrix(MR, 1) = rs!Wert
      Me.MFG.col = 2
      Me.MFG.Row = MR
      Me.MFG.CellBackColor = vbHellGrau
      IDS(0, MR) = cb
     End If
' Farbe zur aktuellen Verknüpfung korrigieren
     Call ZeigeZahl(QuellZ)
    Else
     Me.MFG.TextMatrix(MR, 1) = vNS
     IDS(0, MR) = 0
    End If
   Else
    Me.MFG.TextMatrix(MR, 1) = vNS
    IDS(0, MR) = 0
   End If
 End Select
 End If
End Sub ' mouseup


Private Sub LaborFüll(Optional nachLangtext%)
 Dim rs0 As New ADODB.Recordset, sql0$, sql1$, i&
 FNr = 2
 sql0 = "SELECT * FROM " & vbCrLf & _
        "(SELECT CONCAT(lp.abkü, ' [', lp.einheit,']') bezug, i.abkü,i.einheit,i.nb," & vbCrLf & _
        "DATE_FORMAT(eingang,'%d.%m.%y') Eingang,i.langtext,i.labor,herk,pid, oid " & vbCrLf & _
        "FROM (" & vbCrLf & _
        "SELECT 'x' herk, abkü,einheit, langtext,labor,ergaenzt eingang,nb, gl.idypbez pid,p.id oid " & vbCrLf & _
        "FROM `laborypneu` p " & vbCrLf & _
        "LEFT JOIN `laboryplab` l ON p.labid = l.id " & vbCrLf & _
        "LEFT JOIN `laborypnb` nb ON p.id = nb.pneuid " & vbCrLf & _
        "LEFT JOIN `laborypgl` gl ON p.id = gl.idypneu " & vbCrLf & _
        "UNION " & vbCrLf & _
        "SELECT 'p' herk, abkü,einheit, langtext,labor,aktzeit eingang, CONCAT(unw,'-',onw) nb, gl.idypbez pid, p.id oid " & vbCrLf & _
        "FROM `laborparameter` p " & vbCrLf & _
        "LEFT JOIN `laborypgl` gl ON p.id = gl.idpara) i " & vbCrLf & _
        "LEFT JOIN `laborypneu` lp ON i.pid = lp.id ORDER BY abkü DESC, herk, einheit,labor, eingang DESC) i " & vbCrLf & _
        "GROUP BY herk,labor,abkü,einheit "
 Me.MFG.Clear
 Me.MFG.Cols = MFGLabCols ' 9
 Me.MFG.Row = 0
 If nachLangtext Then
  sql0 = sql0 & "ORDER BY langtext,abkü,einheit DESC,nb DESC"
  Me.MFG.col = 2
  Me.MFG.CellBackColor = vbGelblichGrau
  Me.MFG.col = 6
  Me.MFG.CellBackColor = GelblichRosa
 Else
  sql0 = sql0 & "ORDER BY abkü,einheit DESC,nb DESC"
  Me.MFG.col = 6
  Me.MFG.CellBackColor = vbGelblichGrau
  Me.MFG.col = 2
  Me.MFG.CellBackColor = GelblichRosa
 End If
 myFrag rs0, "SELECT COUNT(0) ct FROM (" & sql0 & ") i"
 If Not rs0.BOF Then MFG.Rows = rs0!ct + 2
 ReDim IDS(1, Me.MFG.Rows)
 Set rs0 = Nothing
 myFrag rs0, sql0
 i = 1
' MFG.FillStyle = flexFillRepeat
 Do While Not rs0.EOF
  Me.MFG.TextMatrix(i, 0) = i
  Me.MFG.TextMatrix(i, 1) = IIf(IsNull(rs0!Bezug), vNS, rs0!Bezug)
  Me.MFG.TextMatrix(i, 2) = rs0!Abkü
  Me.MFG.TextMatrix(i, 3) = rs0!Einheit
  Me.MFG.TextMatrix(i, 4) = IIf(IsNull(rs0!Nb), vNS, rs0!Nb)
  Me.MFG.TextMatrix(i, 5) = IIf(IsNull(rs0!Eingang), vNS, rs0!Eingang)
  Me.MFG.TextMatrix(i, 6) = rs0!Langtext
  Me.MFG.TextMatrix(i, 7) = rs0!Labor
  Me.MFG.TextMatrix(i, 8) = rs0!herk
  If Not IsNull(rs0!pid) Then IDS(0, i) = rs0!pid
  IDS(1, i) = rs0!oid
  Me.MFG.Row = i
  Me.MFG.col = 0
  Select Case rs0!herk
   Case "x"
    Me.MFG.CellBackColor = vbYellow
    Call ZeigeZahl(i)
'   Case Else
'    Me.MFG.CellBackColor = vbgelblichgrau
  End Select
  Me.MFG.col = 2
  If LenB(Me.MFG.TextMatrix(i, 1)) <> 0 Then
   If rs0!herk <> "x" Or Me.MFG.CellBackColor = vbWhite Then Me.MFG.CellBackColor = vbHellGrau
  End If
  rs0.MoveNext
  i = i + 1
 Loop
 Me.MFG.TextMatrix(0, 1) = "Bezug"
 Me.MFG.TextMatrix(0, 2) = "Abkü"
 Me.MFG.TextMatrix(0, 3) = "Einheit"
 Me.MFG.TextMatrix(0, 4) = "NB"
 Me.MFG.TextMatrix(0, 5) = "Eingang"
 Me.MFG.TextMatrix(0, 6) = "Langtext"
 Me.MFG.TextMatrix(0, 7) = "Labor"
 Me.MFG.TextMatrix(0, 8) = "Herk"
' Me.MFG.TopRow = 80
 Me.MFG.Row = 1
 Me.MFG.col = 2
End Sub ' LaborFüll

Private Function ZeigeZahl(zeile&)
 Dim rs1 As New ADODB.Recordset
    Set rs1 = Nothing
    myFrag rs1, "SELECT COUNT(0) zl, `idypneu`,`idpara` FROM `laborypgl` WHERE `idypbez` = " & IDS(1, zeile)
    If Not rs1.BOF Then
     If zeile <> 0 Then Me.MFG.Row = zeile
     Me.MFG.col = 2
     Select Case rs1!zl
      Case 0
       Me.MFG.CellBackColor = vbWhite
      Case 1
       Me.MFG.CellBackColor = vbHellGelb
      Case 2
       Me.MFG.CellBackColor = vbGelb
      Case 3
       Me.MFG.CellBackColor = vbsiena
      Case 4
       Me.MFG.CellBackColor = vborange
      Case Is > 4
       Me.MFG.CellBackColor = vbbraun
     End Select
    End If
End Function ' ZeigeZahl(id&, Optional Zeile&)
