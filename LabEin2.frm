VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form LabEin2 
   Caption         =   "namen"
   ClientHeight    =   10350
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   14625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   14625
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "S&chließen"
      Height          =   300
      Left            =   13440
      TabIndex        =   0
      Top             =   9600
      Width           =   1080
   End
   Begin MSAdodcLib.Adodc dRS 
      Align           =   2  'Unten ausrichten
      Height          =   330
      Left            =   0
      Top             =   10020
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;dsn=MySQLpraxis;uid=praxis;pwd=sonne;database=quelle;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=MySQLpraxis;uid=praxis;pwd=sonne;database=quelle;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"LabEin2.frx":0000
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "LabEin2.frx":0130
      DragIcon        =   "LabEin2.frx":0142
      Height          =   9480
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   16722
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   19
      Cols            =   35
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   $"LabEin2.frx":044C
      _NumberOfBands  =   1
      _Band(0).Cols   =   35
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0)._NumMapCols=   34
      _Band(0)._MapCol(0)._Name=   "Pat_ID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(1)._Name=   "lfdnr"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(1)._Alignment=   7
      _Band(0)._MapCol(2)._Name=   "NVorsatz"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "Nachname"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "Vorname"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "GebDat"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "Straße"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "KVKStatus"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "Geschlecht"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(9)._Name=   "Plz"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(10)._Name=   "Ort"
      _Band(0)._MapCol(10)._RSIndex=   10
      _Band(0)._MapCol(11)._Name=   "Weggeldzone"
      _Band(0)._MapCol(11)._RSIndex=   11
      _Band(0)._MapCol(12)._Name=   "AufnDat"
      _Band(0)._MapCol(12)._RSIndex=   12
      _Band(0)._MapCol(13)._Name=   "intZOGP"
      _Band(0)._MapCol(13)._RSIndex=   13
      _Band(0)._MapCol(14)._Name=   "Titel"
      _Band(0)._MapCol(14)._RSIndex=   14
      _Band(0)._MapCol(15)._Name=   "Versichertennummer"
      _Band(0)._MapCol(15)._RSIndex=   15
      _Band(0)._MapCol(16)._Name=   "PrivatTel"
      _Band(0)._MapCol(16)._RSIndex=   16
      _Band(0)._MapCol(17)._Name=   "KVNr"
      _Band(0)._MapCol(17)._RSIndex=   17
      _Band(0)._MapCol(18)._Name=   "PrivatTel_2"
      _Band(0)._MapCol(18)._RSIndex=   18
      _Band(0)._MapCol(19)._Name=   "PrivatFax"
      _Band(0)._MapCol(19)._RSIndex=   19
      _Band(0)._MapCol(20)._Name=   "DienstTel"
      _Band(0)._MapCol(20)._RSIndex=   20
      _Band(0)._MapCol(21)._Name=   "PrivatMobil"
      _Band(0)._MapCol(21)._RSIndex=   21
      _Band(0)._MapCol(22)._Name=   "Email"
      _Band(0)._MapCol(22)._RSIndex=   22
      _Band(0)._MapCol(23)._Name=   "AnAllgda"
      _Band(0)._MapCol(23)._RSIndex=   23
      _Band(0)._MapCol(24)._Name=   "An1da"
      _Band(0)._MapCol(24)._RSIndex=   24
      _Band(0)._MapCol(25)._Name=   "An2da"
      _Band(0)._MapCol(25)._RSIndex=   25
      _Band(0)._MapCol(26)._Name=   "Checkda"
      _Band(0)._MapCol(26)._RSIndex=   26
      _Band(0)._MapCol(27)._Name=   "DMTypaD"
      _Band(0)._MapCol(27)._RSIndex=   27
      _Band(0)._MapCol(28)._Name=   "AktZeit"
      _Band(0)._MapCol(28)._RSIndex=   28
      _Band(0)._MapCol(29)._Name=   "absPos"
      _Band(0)._MapCol(29)._RSIndex=   29
      _Band(0)._MapCol(29)._Alignment=   7
      _Band(0)._MapCol(30)._Name=   "StByte"
      _Band(0)._MapCol(30)._RSIndex=   30
      _Band(0)._MapCol(30)._Alignment=   7
      _Band(0)._MapCol(31)._Name=   "Notiz"
      _Band(0)._MapCol(31)._RSIndex=   31
      _Band(0)._MapCol(32)._Name=   "zubenach"
      _Band(0)._MapCol(32)._RSIndex=   32
      _Band(0)._MapCol(33)._Name=   "Verwandt"
      _Band(0)._MapCol(33)._RSIndex=   33
   End
   Begin VB.Image imgSort 
      Height          =   480
      Index           =   1
      Left            =   2730
      Top             =   2355
      Width           =   1200
   End
End
Attribute VB_Name = "LabEin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MARGIN_SIZE = 60      ' in Twips
' Variablen für das Aktivieren der Spaltensortierung
Private m_iSortCol As Integer
Private m_iSortType As Integer

' Variablen für das Ziehen von Spalten
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer
Public hlese As Lese

Private Sub datPrimaryRS_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Form_Load()

    dpS.Visible = False

    With MSHFlexGrid1

        .Redraw = False
        ' Spaltenbreite für Tabelle festlegen
        .ColWidth(0) = 300
        .ColWidth(1) = -1
        .ColWidth(2) = -1
        .ColWidth(3) = -1
        .ColWidth(4) = -1
        .ColWidth(5) = -1
        .ColWidth(6) = -1
        .ColWidth(7) = -1
        .ColWidth(8) = -1
        .ColWidth(9) = -1
        .ColWidth(10) = -1
        .ColWidth(11) = -1
        .ColWidth(12) = -1
        .ColWidth(13) = -1
        .ColWidth(14) = -1
        .ColWidth(15) = -1
        .ColWidth(16) = -1
        .ColWidth(17) = -1
        .ColWidth(18) = -1
        .ColWidth(19) = -1
        .ColWidth(20) = -1
        .ColWidth(21) = -1
        .ColWidth(22) = -1
        .ColWidth(23) = -1
        .ColWidth(24) = -1
        .ColWidth(25) = -1
        .ColWidth(26) = -1
        .ColWidth(27) = -1
        .ColWidth(28) = -1
        .ColWidth(29) = -1
        .ColWidth(30) = -1
        .ColWidth(31) = -1
        .ColWidth(32) = -1
        .ColWidth(33) = -1
        .ColWidth(34) = -1

        ' Stil für Tabelle festlegen
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' Spaltenkopf Fett formatieren
        .Row = 0
        .col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With

End Sub

Private Sub MSHFlexGrid1_DragDrop(Source As Control, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' Code in den DragDrop-, MouseDown-, MouseMove- und MouseUp-Ereignissen der Tabelle aktiviert das Ziehen von Spalten
'-------------------------------------------------------------------------------------------

    If m_iDragCol = -1 Then Exit Sub    ' es wurde nicht gezogen
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub
    If MSHFlexGrid1.FixedCols = 1 And MSHFlexGrid1.MouseCol = 0 Then Exit Sub

    With MSHFlexGrid1
        .Redraw = False
        .ColPosition(m_iDragCol) = .MouseCol
        .Redraw = True
    End With

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' Code in den DragDrop-, MouseDown-, MouseMove- und MouseUp-Ereignissen der Tabelle aktiviert das Ziehen von Spalten
'-------------------------------------------------------------------------------------------

    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub
    If MSHFlexGrid1.MouseCol = 0 And MSHFlexGrid1.FixedCols = 1 Then Exit Sub

    xdn = X
    ydn = Y
    m_iDragCol = -1     ' Zieh-Attribut löschen
    m_bDragOK = True

End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' Code in den DragDrop-, MouseDown-, MouseMove- und MouseUp-Ereignissen der Tabelle aktiviert das Ziehen von Spalten
'-------------------------------------------------------------------------------------------

    ' testen, um zu sehen, ob zum dem Ziehen begonnen werden kann
    If Not m_bDragOK Then Exit Sub
    If Button <> 1 Then Exit Sub                        ' falsche Schaltfläche
    If m_iDragCol <> -1 Then Exit Sub                   ' es wird bereits gezogen
    If Abs(xdn - X) + Abs(ydn - Y) < 50 Then Exit Sub   ' noch nicht genug bewegt
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub         ' Spaltenkopf muß gezogen werden

    ' wenn Sie bis hierher gekommen sind, dann starten Sie den Ziehvorgang
    m_iDragCol = MSHFlexGrid1.MouseCol
    MSHFlexGrid1.Drag vbBeginDrag

End Sub

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' Code in den DragDrop-, MouseDown-, MouseMove- und MouseUp-Ereignissen der Tabelle aktiviert das Ziehen von Spalten
'-------------------------------------------------------------------------------------------

    m_bDragOK = False

End Sub

Private Sub MSHFlexGrid1_DblClick()
'-------------------------------------------------------------------------------------------
' Code in DblClick-Ereignis der Tabelle aktiviert Spaltensortierung
'-------------------------------------------------------------------------------------------

    Dim i As Integer

    ' nur dann sortieren, wenn eine feste Zeile angeklickt wurde
    If MSHFlexGrid1.MouseRow >= MSHFlexGrid1.FixedRows Then Exit Sub

    i = m_iSortCol                  ' alte Spalte speichern
    m_iSortCol = MSHFlexGrid1.col   ' neue Spalte festlegen

    ' Sortiertyp inkrementieren
    If i <> m_iSortCol Then
        ' wenn eine neue Spalte geklickt wird, mit aufsteigender Sortierung beginnen
        m_iSortType = 1
    Else
        ' wenn dieselbe Spalte geklickt wird, zwischen aufsteigender und absteigender Sortierung umschalten
        m_iSortType = m_iSortType + 1
    If m_iSortType = 3 Then m_iSortType = 1
    End If

    DoColumnSort

End Sub

Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' Führt Exchange-Sortierung von column m_iSortCol durch
'-------------------------------------------------------------------------------------------

    With MSHFlexGrid1
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .col = m_iSortCol
        .Sort = m_iSortType
        .Redraw = True
    End With

End Sub

Private Sub Form_Resize()

    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single

    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight

        ' Schaltfläche 'Schließen' in untere rechte Ecke verschieben
        With .cmdClose
                sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE)
                .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With

        .MSHFlexGrid1.Move MARGIN_SIZE, _
            MARGIN_SIZE, _
            sngScaleWidth - (2 * MARGIN_SIZE), _
            sngButtonTop - (2 * MARGIN_SIZE)

    End With
    Exit Sub

Form_Resize_Error:
    ' Fehler bei negativen Werten vermeiden
    Resume Next

End Sub
Private Sub cmdClose_Click()

    Unload Me

End Sub


