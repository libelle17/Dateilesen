VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form TabÜbertr 
   Caption         =   "Tabellen übertragen"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14775
   DrawMode        =   0  'Schwarzintensität
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Austauschen 
      Caption         =   "<-&>"
      Height          =   255
      Left            =   7920
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   2895
      Left            =   8640
      TabIndex        =   10
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2895
      Left            =   3360
      TabIndex        =   9
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Start"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   480
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
Public CnQ$, CnZ$
Public F0 As Lese

Private Sub Austauschen_Click()
 Dim CnI$
 CnI = CnQ
 CnQ = CnZ
 CnZ = CnI
 CnI = CsQ.Caption
 CnQ.Caption = CsZ.Caption
 CnZ.Caption = CnI
 Call zeigtabellen
 Call TZahl
End Sub

Private Sub CsQ_Click()
 CnQ = F0.dbv.Auswahl("", "anamnesebogen", "Kopierquelle")
 If CnQ <> "" Then Me.CsQ.Caption = F0.dbv.ConStr
 zeigtabellen
End Sub

Private Sub CsZ_Click()
 CnZ = F0.dbv.Auswahl("", "anamnesebogen", "Kopierziel")
 If CnZ <> "" Then Me.CsZ.Caption = F0.dbv.ConStr
 Call TZahl(nurZiel:=True)
End Sub

Private Sub Form_Activate()
 Call TZahl
End Sub

Private Sub Form_Load()
 CnQ = F0.dbv.cnVorb("", "anamnesebogen", "Kopierquelle")
 If CnQ <> "" Then Me.CsQ.Caption = F0.dbv.ConStr
 zeigtabellen
 CnZ = F0.dbv.cnVorb("", "anamnesebogen", "Kopierziel")
 If CnZ <> "" Then Me.CsZ.Caption = F0.dbv.ConStr
' Call TZahl
End Sub
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
'   Stop
 End Select
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Form_KeyDown/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub
Sub zeigtabellen()
 Dim rs1 As New ADODB.Recordset, rCn As New ADODB.Connection
 Dim i%, obMySQL%, altDaBa$
 On Error GoTo fehler
 obMySQL = (InStr(Me.CsQ.Caption, "MySQL") > 0) ' geht auch mit obMySQL = 0, aber viel langsamer
 Set rCn = Nothing
' rCn.CursorLocation = adUseServer
 On Error Resume Next
 Call rCn.Open(CnQ) '"PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=mitte;uid=praxis;database=quelle;pwd=sonne;option=3") ' (Me.CsQ.Caption)
 If Err.Number = 0 Then
'     On Error Resume Next
     Set rs1 = Nothing
     Err.Clear
     If obMySQL Then
      Set rs1 = rCn.Execute("show tables")
     Else
      Set rs1 = rCn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))  'Array(, Chr(34) & rSch!catalog_name & Chr(34))
     End If
     Me.Tabelle.Clear
      Do While Not rs1.EOF
       Dim Tb$
       Tb = ""
       Tb = rs1!table_name
       If Tb = "" Then Tb = rs1.Fields(0)
       Me.Tabelle.AddItem Tb
       rs1.Move 1
      Loop
'     On Error GoTo fehler
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Verbind/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' zeigtabellen

Private Sub Form_Unload(Cancel As Integer)
' stop
End Sub

Private Sub Start_Click()
 Dim rQCn As New ADODB.Connection, rZCn As New ADODB.Connection, erg, CurCat$, Frage$
 Dim rs As New ADODB.Recordset, InS$, AuiFd$, FListe$
 rQCn.Open CnQ
 rZCn.Open CnZ
 Dim FrageS$
 CurCat = rZCn.Properties("Current Catalog")
 If CurCat = "" Then CurCat = rZCn.Properties("Data Source Name")
 FrageS = "Wollen Sie wirklich alle " & Me.ZZ & " Datensätze aus " & vbCrLf & Chr(34) & Me.Tabelle & Chr(34) & " in " & vbCrLf & rZCn.Properties("dbms name") & " /" & vbCrLf & CurCat & vbCrLf & " löschen?"
 erg = MsgBox(FrageS, vbOKCancel)
 LVobMySQL = (LCase(rZCn.ConnectionString) Like "*mysql*") ' auch für datform(, umw(
 If erg = 1 Then
  If LVobMySQL Then rZCn.Execute ("set foreign_key_checks = 0")
  rZCn.Execute "delete from `" & LCase(Me.Tabelle) & "`"
  AuiFd = GetAutoFeld(rZCn, LCase(Me.Tabelle), FListe)
  rs.Open "select * from `" & Me.Tabelle & "`", rQCn, adOpenStatic, adLockReadOnly
  Do While Not rs.EOF
   Call TIns(rZCn, LCase(Me.Tabelle), rs, AuiFd, FListe)
   rs.Move 1
  Loop
  If LVobMySQL Then rZCn.Execute ("set foreign_key_checks = 1")
 End If
 Call TZahl(nurZiel = True)
 Beep
End Sub ' Start_Click

Private Sub Tabelle_Change()
 Call TZahl
End Sub

Private Sub TZahl(Optional nurZiel%)
 Dim rs As New ADODB.Recordset
 If nurZiel = 0 Then ZQ = ""
 ZZ = ""
 Screen.MousePointer = vbHourglass
 On Error Resume Next
 If Tabelle <> "" Then
  If CnQ <> "" And nurZiel = 0 Then
   rs.Open "select count(*) as ct from `" & LCase(Tabelle) & "`", CnQ, adOpenStatic, adLockReadOnly
   ZQ = rs!ct
  End If
  If CnZ <> "" Then
   Set rs = Nothing
   rs.Open "select count(*) as ct from `" & LCase(Tabelle) & "`", CnZ, adOpenStatic, adLockReadOnly
   ZZ = rs!ct
  End If
 End If
 Screen.MousePointer = vbNormal
End Sub

Private Sub Tabelle_Click()
 Call TZahl
End Sub
Public Function GetAutoFeld$(ZCn As ADODB.Connection, TabN$, FListe$) ' ob Primärschlüssel
' s. http://support.microsoft.com/kb/q195910/
' by mysql ggf. mit lcase(tabn) aufrufen
   Dim rs As New ADODB.Recordset, i%
   Call rs.Open("select " & IIf(InStr(ZCn.ConnectionString, "Jet.OLEDB") > 0, "top 1", "") & " * from `" & TabN & "`" & IIf(LVobMySQL, " limit 1", ""), ZCn, adOpenDynamic, adLockReadOnly)
   FListe = "("
   For i = 0 To rs.Fields.Count - 1
    If rs.Fields(i).Properties("ISAUTOINCREMENT") Then
     GetAutoFeld = rs.Fields(i).Name
'     Exit Function
    Else
     FListe = FListe & "`" & rs.Fields(i).Name & "`,"
    End If
   Next i
   FListe = Left(FListe, Len(FListe) - 1) & ")"
   Exit Function
   On Error Resume Next
' zeigt nicht auto_increment an
   Dim catx As New ADOX.Catalog
   Set catx = Nothing
   catx.ActiveConnection = ZCn.ConnectionString
   For i = 1 To catx.Tables(TabN).Columns.Count
    rs.Move 1
   Next i
   
   Dim rSc As New ADODB.Recordset
'   Set rSc = ZCn.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, TabN))
   Set rSc = ZCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, TabN))
   If Err.Number <> 0 Then
    Exit Function
   End If
   Open "u:\Tabus.txt" For Append As #399
   Print #399, rSc.ActiveConnection, Now()
   Do While Not rSc.EOF
'    Debug.Print rSc!table_catalog, rSc!table_name, rSc!index_name, rSc!primary_key, rSc!Unique, rSc!column_name, rSc!column_guid, rSc!column_propid, rSc!collation, rSc!Clustered, rSc!Type, rSc!null_collation, rSc!ordinal_position
    Dim ausg$
    For i = 0 To rSc.Fields.Count - 1
     ausg = ausg & rSc.Fields(i).Name & ":" & rSc.Fields(i).Value & " "
    Next i
    Debug.Print ausg
    Print #399, ausg
    ausg = ""
    rSc.Move 1
   Loop
   Close #399
   Shell Environ("windir") & "\system32\notepad.exe u:\tabus.txt", vbMaximizedFocus
End Function
#Const turbomed = -1
Public Function TIns&(ZCn As ADODB.Connection, TabN$, rQ, AuiFd$, FListe$, Optional schonAbgehakt) ' AutoIncrementFeld, ob Primärschlüssel
'     Dim rs2 As New ADODB.Recordset
'     Set rs2 = rCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, LCase(rs1!table_name), Empty))
'     Do While Not rs2.EOF
'' data_type: 130 = String, 7 = Datum, 2 = integer, 3 = long, 5 = double, 11 = ja/nein
'' column_flags: 106 = normal für Strings, 122 = normal für numerisch, 234 = bei Memo, 90 bei ja/nein
'      Debug.Print Left(rs2!column_name & Space(30), 30) & Right(Space(6) & rs2!data_type, 6) & Right(Space(6) & rs2!column_flags, 6)
'      rs2.Move 1
'     Loop
     Dim j&
 
 Dim i&, sql2$, sqldel$, Wert$, mprim&, obPS%
 Dim rZ As New ADODB.Recordset, rs As ADODB.Recordset
   On Error GoTo fehler
   sql2 = "insert into `" & TabN & "`" & FListe & " values("
   
'   If AuiFd <> "" Then
'    Set rs = ZCn.Execute("select max(" & AuiFd & ") as mprim from `" & kla & TabN & "`")
'    mprim = IIf(IsNull(rs!mprim), 0, rs!mprim) + 1
'   End If
   For i = 0 To rQ.Fields.Count - 1
    Select Case LCase(rQ.Fields(i).Name)
     Case AuiFd
#If False Then
      Wert = mprim
      If obPS Then
       Wert = mprim
      Else
       Wert = rQ(sF)
       Wert = IIf(IsNull(rQ(sF)), "", rQ(sF))
#If turbomed Then
       If InStr(sF, "pfad") > 0 Then
        Dim ob2%, obnetz%
        If Wert Like "??*\\*" Then
         ob2 = True
         If Wert Like "\\\\*" Then obnetz = True
        Else
         If Wert Like "\\*" Then obnetz = True
        End If
'        If PcDokPfad = "" Then getDokPfad
'        Wert = umwfSQL(Replace(LCase(Replace(Wert, "\\", "\")), "$\turbomed\dokumente", PcDokPfad))
        Wert = umwfSQL(Replace(Wert, "\\", "\"))
        If obnetz And Not ob2 Then Wert = "\\" & Wert
       End If
#End If
       Wert = "'" & Wert & "'"
      End If
#End If
     Case Else
      Select Case rQ.Fields(i).Type
         Case adBoolean ' 11
           If IsNull(rQ.Fields(i).Value) Then
            Wert = "0"
           Else
            If rQ.Fields(i) Then
             Wert = "-1"
             If LVobMySQL Then
'              Set rZ = Nothing
'              rZ.Open "show full columns from " & kla & TabN & klz & " where field = '" & rQ.Fields(i).Name & "'", ZCn, adOpenStatic, adLockReadOnly
'              If Not rZ.BOF Then
'               If rZ!Type = "bit(1)" Then
                Wert = "1"
'               End If
'              End If
             End If
            Else
             Wert = "0"
            End If
           End If
'           If Not IsError(schonAbgehakt) Then
'            If LCase(rQ.Fields(i).Name) = "abgehakt" And schonAbgehakt Then
'             If lies.obMySQL Then Wert = "1" Else Wert = "-1"
'            End If
'           End If
         Case 16, 17, 2, 18, 3, 19, 4, 5, 20, 21, 131, 139, 6, 14
           If IsNull(rQ.Fields(i)) Then
            Wert = "null"
           Else
            Wert = Replace(Replace(rQ.Fields(i), ".", ""), ",", ".")
            If Wert = "" Then Wert = 0
           End If
         Case 7, 64, 133, 134, 135
           Wert = datform(rQ.Fields(i).Value)
           If Wert = "##" Or Wert = "''" Then Err.Raise 999, , "Fehler in TIns mit Wert = ""##"" or wert = ""''"" in Datumsfeld"
         Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
           Wert = umw(IIf(IsNull(rQ.Fields(i).Value), "", rQ.Fields(i).Value))
           If InStr(rQ.Fields(i).Name, "pfad") > 0 Or InStr(rQ.Fields(i).Name, "path") > 0 Then
'            If PcDokPfad = "" Then getDokPfad
'            Wert = Replace(LCase(Replace(Wert, "\\", "\")), "$\turbomed\dokumente", PcDokPfad)
            Stop
            Wert = Replace(Wert, "\\", "\")
           End If
           Wert = "'" & Wert & "'"
         Case Else
           Err.Raise 999, , "Fehler in TIns: Falsche Zahl an Datensätzen aktualisiert: " & TIns
       End Select
    End Select
'    sql1 = sql1 & kla & rQ.Fields(i).Name & klz & ","
    sql2 = sql2 & Wert & ","
    Wert = ""
   Next i
   sql2 = Left(sql2, Len(sql2) - 1) & ")"
'   sql1 = Left(sql1, Len(sql1) - 1) & ") " & sql2
   Dim rAF&
'   zcn.CursorLocation = adUseClient
   Call ZCn.Execute(sql2, rAF)
   If rAF = 0 Or (rAF <> 1 And rAF < 100) Then Err.Raise 999, , "Fehler in TIns: Falsche Zahl an Datensätzen aktualisiert: " & rAF
   If rAF <> 1 Then TIns = 1
   Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in TIns/" & AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' TIns(rq As Recordset)

