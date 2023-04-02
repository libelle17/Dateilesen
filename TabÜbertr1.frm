VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form TabÜbertr 
   Caption         =   "Tabellen übertragen"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14775
   KeyPreview      =   -1  'True
   LinkTopic       =   "TabÜbertr"
   ScaleHeight     =   5745
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton C1 
      Caption         =   "Labor&gruppen"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton C1 
      Caption         =   "&Laborparameter"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton C1 
      Caption         =   "&Medarten"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton C1 
      Caption         =   "&Kassenliste"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton C1 
      Caption         =   "&Dokumente abgehakt"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton C1 
      Caption         =   "&Anamnesebogen"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CheckBox ReplaceStattInsert 
      Caption         =   "&Replace statt Insert"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton Austauschen 
      Caption         =   "<-&>"
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   2895
      Left            =   8640
      TabIndex        =   11
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2895
      Left            =   3360
      TabIndex        =   10
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Start"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox Tabelle 
      Height          =   315
      Left            =   720
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton CsZ 
      Height          =   1455
      Left            =   8640
      TabIndex        =   3
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton CsQ 
      Height          =   1455
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label ZZ 
      Height          =   255
      Left            =   8640
      TabIndex        =   7
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label ZQ 
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Label TabelleLabel 
      Caption         =   "&Tabelle:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label CsZLabel 
      Caption         =   "ConnectionString &Ziel:"
      Height          =   375
      Left            =   10440
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label CsQLabel 
      Caption         =   "ConnectionString &Quelle:"
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "TabÜbertr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CnQ$, cnz$
Public F0 As Lese
Dim nichtTZahl%

Private Sub Austauschen_Click()
 Dim CnI$
 Screen.MousePointer = vbHourglass
 DoEvents
 CnI = Me.CnQ
 Me.CnQ = Me.cnz
 Me.cnz = CnI
 CnI = Me.CsQ.Caption
 Me.CsQ.Caption = Me.CsZ.Caption
 Me.CsZ.Caption = CnI
 CnI = Me.ZQ
 Me.ZQ = Me.ZZ
 Me.ZZ = CnI
 Call zeigtabellen
 Call TZahl
 Screen.MousePointer = vbNormal
 DoEvents
End Sub ' Austauschen_Click

Private Sub C1_Click(Index As Integer)
 Me.Tabelle = LCase$(REPLACE$(Me.C1(Index).Caption, "&", ""))
End Sub ' C1_Click

Private Sub CsQ_Click()
 Me.CnQ = F0.dbv.Auswahl("", "anamnesebogen", "Kopierquelle")
 If Me.CnQ <> "" Then Me.CsQ.Caption = F0.dbv.Constr
 zeigtabellen
End Sub ' CsQ_Click

Private Sub CsZ_Click()
 cnz = F0.dbv.Auswahl("", "anamnesebogen", "Kopierziel")
 If cnz <> "" Then Me.CsZ.Caption = F0.dbv.Constr
 Call TZahl(nurZiel:=True)
End Sub ' CsZ_Click

Private Sub Form_Activate()
 
' Call TZahl
End Sub ' Form_Activate

Private Sub Form_Load()
 CnQ = F0.dbv.cnVorb("", "anamnesebogen", "Kopierquelle")
 If CnQ <> "" Then Me.CsQ.Caption = F0.dbv.Constr
 zeigtabellen
 cnz = F0.dbv.cnVorb("", "anamnesebogen", "Kopierziel")
 If cnz <> "" Then Me.CsZ.Caption = F0.dbv.Constr
' Call TZahl
End Sub ' Form_Load()

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo fehler
 Select Case KeyCode
  Case 27
'  Unload Me
   Call Form_Unload(Cancel:=True)
   Me.Visible = False
'  Case 13
'   Call OK_Click
  Case Else
'
 End Select
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Form_KeyDown/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub

Sub zeigtabellen()
 Dim rs1 As New ADODB.Recordset, rCn As New ADODB.Connection
 Dim i%, obMySQL%, altDaBa$
 Screen.MousePointer = vbHourglass
 On Error GoTo fehler
 obMySQL = (InStrB(Me.CsQ.Caption, "MySQL") > 0) ' geht auch mit obMySQL = 0, aber viel langsamer
 Set rCn = Nothing
' rCn.CursorLocation = adUseServer
 On Error Resume Next
 Call rCn.Open(CnQ) '"PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=mitte;uid=praxis;database=quelle;pwd=***REMOVED***;option=3") ' (Me.CsQ.Caption)
 If Err.Number = 0 Then
'     ON Error Resume Next
     Set rs1 = Nothing
     Err.Clear
     If obMySQL Then
      Set rs1 = myEFrag("SHOW TABLES", , rCn)
     Else
      Set rs1 = rCn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))  'Array(, "" & rSch!catalog_name & "")
     End If
     Dim altTab$
     altTab = Me.Tabelle
     Me.Tabelle.Clear
      Do While Not rs1.EOF
       Dim Tb$
       Tb = vNS
       Tb = rs1!table_name
       If Tb = "" Then Tb = rs1.Fields(0)
       Me.Tabelle.AddItem Tb
       rs1.Move 1
      Loop
      Me.Tabelle = altTab
'     ON Error GoTo fehler
 End If
 Screen.MousePointer = vbNormal
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Verbind/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' zeigtabellen

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Private Sub Start_Click()
 Dim rQCn As New ADODB.Connection, rZCn As New ADODB.Connection, erg, CurCat$, Frage$
 Dim rsq As New ADODB.Recordset, InS$, AuiFd$, FListe$, rAF&
 On Error GoTo fehler
 rQCn.Open CnQ
 rZCn.Open cnz
 Dim FrageS$
 CurCat = rZCn.Properties("Current Catalog")
 If LenB(CurCat) = 0 Then CurCat = DefDB(rZCn) ' rZCn.Properties("Data Source Name")
 If Me.ReplaceStattInsert = 0 Then
  FrageS = "Wollen Sie wirklich alle " & Me.ZZ & " Datensätze aus " & vbCrLf & "" & Me.Tabelle & "" & " in " & vbCrLf & rZCn.Properties("dbms name") & " /" & vbCrLf & CurCat & vbCrLf & " löschen?"
  erg = MsgBox(FrageS, vbOKCancel)
 Else
  erg = 1
 End If
 LVobMySQL = (LCase$(cnz) Like "*mysql*") ' rZCn.ConnectionString' auch für DatFor_k(, umw(
 If erg = 1 Then
  If LVobMySQL Then
   myEFrag "SET foreign_key_checks = 0", , rZCn ' Kommentar 12.12.09
  End If
  rZCn.BeginTrans
  If rZCn.DefaultDatabase = DBCn.DefaultDatabase Then obTrans = 1
  If Me.ReplaceStattInsert = 0 Then
   myEFrag "DELETE FROM `" & LCase$(Me.Tabelle) & "`", , rZCn
   myEFrag "ALTER TABLE `" & LCase$(Me.Tabelle) & "` auto_increment=1", , rZCn
  End If
  Me.ZZ = 0
  Dim Felder$()
  myFrag rsq, "SELECT * FROM `" & LCase$(Me.Tabelle) & "`", adOpenStatic, rQCn, adLockReadOnly
'  rsq.Open "SELECT * FROM `" & LCase$(Me.Tabelle) & "`", rQCn, adOpenStatic, adLockReadOnly
  AuiFd = GetAutoFeld(rZCn, cnz, LCase$(Me.Tabelle), rsq, FListe, Felder)
  Do While Not rsq.EOF
   Call TIns(rZCn, LCase$(Me.Tabelle), rsq, AuiFd, FListe, Felder, Me.ReplaceStattInsert)
   If Me.ZZ = "" Then Me.ZZ = 0
   Me.ZZ = Me.ZZ + 1
   DoEvents
   rsq.Move 1
  Loop
  If LVobMySQL Then
   myEFrag "SET foreign_key_checks = 1", , rZCn
  End If
  If rZCn.DefaultDatabase <> DBCn.DefaultDatabase Or ((rZCn.DefaultDatabase = DBCn.DefaultDatabase) And obTrans <> 0) Then rZCn.CommitTrans: If rZCn.DefaultDatabase = DBCn.DefaultDatabase Then obTrans = 0
 End If
 Call TZahl(nurZiel:=True)
 Tüt 440, 1000
 Exit Sub
fehler:
 If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF, rZCn
  Resume
 End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Start_Click/" & AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select

End Sub ' Start_Click

Private Sub Tabelle_Change()
 If Not nichtTZahl Then
  If TZahl = 0 Then
   nichtTZahl = True
   Me.Tabelle = vNS
   nichtTZahl = False
  End If
 End If
End Sub

Private Function TZahl%(Optional nurZiel%)
 Dim rs As New ADODB.Recordset, i%, j%
 TZahl = True
 If nurZiel = 0 Then ZQ = vNS
 ZZ = vNS
 Screen.MousePointer = vbHourglass
' ON Error Resume Next
 If Me.Tabelle <> "" Then
  If CnQ <> "" And nurZiel = 0 Then
   myFrag rs, "SELECT * FROM information_schema.TABLEs t WHERE table_schema = '" & CurDB(CnQ) & "' AND table_name = '" & LCase$(Tabelle) & "'", adOpenStatic, CnQ, adLockReadOnly
   If rs.BOF Then Exit Function
'   Set rs = Nothing
   myFrag rs, "SELECT COUNT(0) ct FROM `" & LCase$(Tabelle) & "`", adOpenStatic, CnQ, adLockReadOnly
   ZQ = rs!ct
   LVobMySQL = (InStrB(LCase$(CnQ), "mysql") > 0)
   Set rs = Nothing
'   rs.Open "SELECT " & IIf(LVobMySQL, vNS, "top 20") & " * FROM `" & LCase$(Me.Tabelle) & "` " & IIf(LVobMySQL, " LIMIT 20", ""), CnQ, adOpenDynamic, adLockReadOnly
   myFrag rs, "SELECT " & IIf(LVobMySQL, vNS, "top 20") & " * FROM `" & LCase$(Me.Tabelle) & "` " & IIf(LVobMySQL, " LIMIT 20", ""), adOpenDynamic, CnQ
   With Me.MSHFlexGrid1
   .Clear
   .Cols = rs.Fields.COUNT
   .Rows = 21
   For i = 0 To rs.Fields.COUNT - 1
    .TextMatrix(0, i) = rs.Fields(i).name
   Next i
   j = 1
   Do While Not rs.EOF
    For i = 0 To rs.Fields.COUNT - 1
     .TextMatrix(j, i) = IIf(IsNull(rs.Fields(i)), vNS, rs.Fields(i))
    Next i
    j = j + 1
    rs.Move 1
   Loop
   End With
  End If
  If cnz <> "" Then
   Set rs = Nothing
   On Error Resume Next
   myFrag rs, "SELECT COUNT(0) ct FROM `" & LCase$(Me.Tabelle) & "`", adOpenStatic, cnz, adLockReadOnly
   On Error GoTo fehler
   If rs.State = 0 Then
    MsgBox "Tabelle `" & Me.Tabelle & "` existiert nicht unter '" & cnz
    TZahl = 0
    Screen.MousePointer = vbNormal
    Me.MSHFlexGrid2.Clear
    Exit Function
   End If
   ZZ = rs!ct
   LVobMySQL = (InStrB(LCase$(cnz), "mysql") > 0)
   Set rs = Nothing
'   rs.Open "SELECT " & IIf(LVobMySQL, vNS, "top 20") & " * FROM `" & LCase$(Me.Tabelle) & "` " & IIf(LVobMySQL, " LIMIT 20", ""), cnz, adOpenDynamic, adLockReadOnly
   myFrag rs, "SELECT " & IIf(LVobMySQL, vNS, "top 20") & " * FROM `" & LCase$(Me.Tabelle) & "` " & IIf(LVobMySQL, " LIMIT 20", ""), adOpenDynamic, cnz
   With Me.MSHFlexGrid2
   .Clear
   .Cols = rs.Fields.COUNT
   .Rows = 21
   For i = 0 To rs.Fields.COUNT - 1
    .TextMatrix(0, i) = rs.Fields(i).name
   Next i
   j = 1
   Do While Not rs.EOF
    For i = 0 To rs.Fields.COUNT - 1
     .TextMatrix(j, i) = IIf(IsNull(rs.Fields(i)), vNS, rs.Fields(i))
    Next i
    j = j + 1
    rs.Move 1
   Loop
   End With
  End If
 End If
 Screen.MousePointer = vbNormal
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetAutoFeld/" & AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' TZahl

Private Sub Tabelle_Click()
 If Not nichtTZahl Then
  If TZahl = 0 Then
   nichtTZahl = True
   Me.Tabelle = vNS
   nichtTZahl = False
  End If
 End If
End Sub ' Tabelle_Click

Public Function GetAutoFeld$(ZCn As ADODB.Connection, ZCnS$, TabN$, rsq, FListe$, ByRef Felder$()) ' ob Primärschlüssel
' s. http://support.microsoft.com/kb/q195910/
' by mysql ggf. mit lcase(tabn) aufrufen
   Dim rs As New ADODB.Recordset, i%
'   Call rs.Open("SELECT " & IIf(InStr(ZCnS, "Jet.OLEDB") > 0, "top 1", "") & " * FROM `" & TabN & "`" & IIf(LVobMySQL, " LIMIT 1", ""), ZCn, adOpenDynamic, adLockReadOnly)
   myFrag rs, "SELECT " & IIf(InStr(ZCnS, "Jet.OLEDB") > 0, "top 1", "") & " * FROM `" & TabN & "`" & IIf(LVobMySQL, " LIMIT 1", ""), adOpenDynamic, ZCn
   ReDim Felder(rs.Fields.COUNT - 1)
   FListe = "("
   For i = 0 To rs.Fields.COUNT - 1
    If rs.Fields(i).Properties("ISAUTOINCREMENT") Then
     GetAutoFeld = rs.Fields(i).name
'     Exit Function
    Else
     Dim QNam$
     QNam = vNS
     On Error Resume Next
     QNam = rsq(rs.Fields(i).name).name
     On Error GoTo fehler
     If QNam <> "" Then
      FListe = FListe & "`" & rs.Fields(i).name & "`,"
     End If
    End If
    Felder(i) = rs.Fields(i).name
   Next i
   FListe = Left$(FListe, Len(FListe) - 1) & ")"
   Exit Function
   
   On Error Resume Next
' zeigt nicht auto_increment an
   Dim catx As New ADOX.Catalog
   Set catx = Nothing
   catx.ActiveConnection = ZCnS
   For i = 1 To catx.Tables(TabN).Columns.COUNT
    rs.Move 1
   Next i
   
   Dim rsc As New ADODB.Recordset
'   SET rSc = ZCn.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, TabN))
   Set rsc = ZCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, TabN))
   If Err.Number <> 0 Then
    Exit Function
   End If
   Open uVerz & "Tabus.txt" For Append As #399
   Print #399, rsc.ActiveConnection, Now()
   Do While Not rsc.EOF
'    Debug.Print rSc!table_catalog, rSc!table_name, rSc!index_name, rSc!primary_key, rSc!Unique, rSc!column_name, rSc!column_guid, rSc!column_propid, rSc!collation, rSc!Clustered, rSc!Type, rSc!null_collation, rSc!ordinal_position
    Dim ausg$
    For i = 0 To rsc.Fields.COUNT - 1
     ausg = ausg & rsc.Fields(i).name & ":" & rsc.Fields(i).Value & " "
    Next i
    Debug.Print ausg
    Print #399, ausg
    ausg = vNS
    rsc.Move 1
   Loop
   Close #399
   zeigan uVerz & "tabus.txt"
   Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetAutoFeld/" & AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetAutoFeld
#Const turbomed = -1

' Tabelle kopieren
Public Function TIns&(ZCn As ADODB.Connection, TabN$, rq, AuiFd$, FListe$, ByRef Felder$(), RsI%, Optional schonAbgehakt) ' AutoIncrementFeld, ob Primärschlüssel
'     Dim rs2 As New ADODB.Recordset
'     SET rs2 = rCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, lcase(rs1!table_name), Empty))
'     Do While Not rs2.EOF
'' data_type: 130 = String, 7 = Datum, 2 = integer, 3 = long, 5 = double, 11 = ja/nein
'' column_flags: 106 = normal für Strings, 122 = normal für numerisch, 234 = bei Memo, 90 bei ja/nein
'      Debug.Print LEFT(rs2!column_name & space$(30), 30) & right$(space$(6) & rs2!data_type, 6) & right$(space$(6) & rs2!column_flags, 6)
'      rs2.Move 1
'     Loop
     Dim j&
 
 Dim i&, sql2$, sqldel$, Wert$, mprim&, obPS%
' Dim rZ As New ADODB.Recordset
'   Call rZ.Open("SELECT " & IIf(InStr(ZCn.ConnectionString, "Jet.OLEDB") > 0, "top 1", "") & " * FROM `" & TabN & "`" & IIf(LVobMySQL, " LIMIT 1", ""), ZCn, adOpenDynamic, adLockReadOnly)
   On Error GoTo fehler
   sql2 = IIf(RsI, "replace", "insert") & " INTO `" & TabN & "`" & FListe & " VALUES("
   
'   IF AuiFd <> "" THEN
'    SET rs = myEFrag("SELECT MAX(" & AuiFd & ") mprim FROM `" & "`" & TabN & "`",,ZCn)
'    mprim = IIf(ISNULL(rs!mprim), 0, rs!mprim) + 1
'   END IF
   For i = 0 To UBound(Felder) ' rQ.Fields.Count - 1
    Dim FTyp&
    FTyp = -1
    On Error Resume Next
    FTyp = rq.Fields(Felder(i)).Type
    On Error GoTo fehler
    If FTyp = -1 Then
    Else
     Select Case LCase$(Felder(i)) ' rQ.Fields(i).Name)
      Case LCase$(AuiFd)
      Case Else
       Select Case FTyp
         Case adBoolean, adUnsignedTinyInt ' 11, 17
           If IsNull(rq.Fields(Felder(i)).Value) Then
            Wert = "0"
           Else
            If rq.Fields(Felder(i)) Then
             Wert = "-1"
             If LVobMySQL Then
'              SET rZ = Nothing
'              rZ.Open "SHOW FULL COLUMNS FROM " & "`" & TabN & "`" & " WHERE field = '" & rQ.Fields(i).Name & "'", ZCn, adOpenStatic, adLockReadOnly
'              IF Not rZ.BOF THEN
'               IF rZ!Type = "bit(1)" THEN
                Wert = "1"
'               END IF
'              END IF
             End If
            Else
             Wert = "0"
            End If
           End If
'           IF Not IsError(schonAbgehakt) THEN
'            IF lcase(rQ.Fields(i).Name) = "abgehakt" AND schonAbgehakt THEN
'             IF lies.obMySQL THEN Wert = "1" ELSE Wert = "-1"
'            END IF
'           END IF
         Case adUnsignedSmallInt, adUnsignedInt, _
              adTinyInt, adSmallInt, adInteger, adSingle, adDouble, adBigInt, adUnsignedBigInt, _
              adNumeric, adVarNumeric, adCurrency, adDecimal
'         Case 18,19,  16, 2, 3, 4, 5, 20, 21, 131, 139, 6, 14
           If IsNull(rq.Fields(Felder(i))) Then
            Wert = "null"
           Else
            Wert = REPLACE$(REPLACE$(rq.Fields(Felder(i)), ".", vNS), ",", ".")
            If LenB(Wert) = 0 Then Wert = 0
           End If
         Case adDate, adFileTime, adDBDate, adDBTime, adDBTimeStamp
'         Case 7, 64, 133, 134, 135
           Wert = DatFor_k(rq.Fields(Felder(i)).Value)
           If Wert = "##" Or Wert = "''" Then Err.Raise 999, , "Fehler in TIns mit Wert = ""##"" OR wert = ""''"" in Datumsfeld"
         Case adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar, adEmpty, _
              adIDispatch, adVariant, adIUnknown, adGUID, adBinary, adUserDefined, adPropVariant, _
              adVarBinary, adLongVarBinary, adError, adArray
'         Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205, 10, 8192
           Wert = doUmwfSQL(IIf(IsNull(rq.Fields(Felder(i)).Value), vNS, rq.Fields(Felder(i)).Value), lies.obMySQL)
'            IF instrb(lcase(rQ.Fields(i).Name), "pfad") > 0 OR instrb(lcase(rQ.Fields(i).Name), "path") > 0 THEN
''            IF lenb(PcDokPfad) = 0 THEN getDokPfad
''            Wert = replace$(lcase(replace$(Wert, "\\", "\")), "$\turbomed\dokumente", PcDokPfad)
'            Wert = replace$(Wert, "\\", "\")
'           END IF
           Wert = "'" & Wert & "'"
'           IF InStrB(Wert, "??") <> 0 THEN Stop
         Case Else
           Err.Raise 999, , "Fehler in TIns: Falsche Zahl an Datensätzen aktualisiert: " & TIns
       End Select
       sql2 = sql2 & Wert & ","
     End Select
'    sql1 = sql1 & "`" & rQ.Fields(i).Name & "`" & ","
    End If
    Wert = vNS
   Next i
   sql2 = Left$(sql2, Len(sql2) - 1) & ")"
'   sql1 = LEFT(sql1, Len(sql1) - 1) & ") " & sql2
   Dim rAF&
'   zcn.CursorLocation = adUseClient
   Call myEFrag(sql2, rAF, ZCn)
   If rAF = 0 Or (rAF <> 1 And rAF < 100 And RsI = 0) Then Err.Raise 999, , "Fehler in TIns: Falsche Zahl an Datensätzen aktualisiert: " & rAF
   If rAF <> 1 Then TIns = 1
   Exit Function
fehler:
 If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF, ZCn
  Resume
 End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in TIns/" & AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' TIns(rq AS Recordset)

'Function umw$(q$) ' jetzt in quelledb
' IF LVobMySQL THEN
'  umw = replace$(replace$(replace$(Trim$(q), "\", "\\"), "'", "\'"), Chr$(0), "")
' Else
'  umw = replace$(replace$(replace$(Trim$(q), "\", "\"), "'", "''"), Chr$(0), "")
' END IF
'End FUNCTION ' umw

