VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Lade 
   Caption         =   "Excel-DAtei laden"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Lade"
   ScaleHeight     =   2790
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Datenbankverbindung 
      Caption         =   "Datenbank&verbindung"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Auswahl 
      Caption         =   "&..."
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Datei 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   9255
   End
   Begin VB.CommandButton DateiBearbeiten 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Conn 
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   9255
   End
   Begin VB.Label DateiLabel 
      Caption         =   "Excel-Datei:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Lade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public obMySQL% ' für QuelleDB, das auch in DateiLese verwendet wird und dort in forms(0) diese Variable verlangt
Public obAcc%
Public MyDB$
Public MdB$
Public FxDB$
Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1
Public opt As Optionen
Attribute opt.VB_VarHelpID = -1
Public FSO As New FileSystemObject
Const RegWurzel$ = "Software\GSProducts\"
Public RegPos$, changeStill%
'Public Conn$
'Public ErrNumber$, ErrDescription$, ErrLastDllError$, ErrSource$ ' in RegComm
'Public DBCn As New ADODB.Connection
Enum Fensterarten
 BriefeZuListe = 1
 LaborArten = 2
End Enum
Public FenArt As Fensterarten  ' Fensterart: 1 = Briefe zu Liste aus Datei

Private Sub Auswahl_Click()
 Dim fileflags As FileOpenConstants
 Dim filefilter$
 On Error GoTo fehler
 Select Case FenArt
  Case 0
   'Set the text in the dialog title bar
   CommonDialog1.DialogTitle = "Exceldatei auswählen:"
   'Set the default file name AND filter
   CommonDialog1.initDir = uVerz
   CommonDialog1.Filename = "Listenausgabe*"
   filefilter = "Exceldateien (*.xls)|*.xls|Alle Dateien (*.*)|*.*"
   'Return the path AND file name SELECTed or
   'Return an empty string IF the user cancels the dialog
   ' Call ergVerarb(Nr, CommonDialog1.FileName)
   Case 1
    CommonDialog1.DialogTitle = "Datei mit Pat_id auswählen:"
    CommonDialog1.initDir = pVerz & "unkorrigiert\"
    CommonDialog1.Filename = "Nochzuschreiben*"
    filefilter = "Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
   Case 2
    CommonDialog1.DialogTitle = "Datei Listenausgabe_Labortests auswählen:"
    CommonDialog1.initDir = pügVerz ' pVerz & "Patientenübergreifendes\"
    CommonDialog1.Filename = "Listenausgabe_Labortests*"
    filefilter = "Exceldateien (*.xls)|*.xls|Alle Dateien (*.*)|*.*"
 End Select
 CommonDialog1.Filter = filefilter
 CommonDialog1.FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNHideReadOnly
 CommonDialog1.flags = fileflags
 'Show the Open common dialog box
 CommonDialog1.ShowOpen
 Me.Datei = CommonDialog1.Filename
 Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Auswahl_Click/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Auswahl_Click

' Datei -> Datenbankverbindung Patientendaten
Private Sub Datenbankverbindung_Click()
 Conn = dbv.Auswahl("quelle", "anamnesebogen", "Patientendaten")
 Call Me.RegSpeichern
End Sub ' Datenbankverbindung_Click

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
  Unload Me
  End
 End If
End Sub ' Form_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Form_Load()
 obMySQL = True
 MyDB = "quelle"
 MdB = QmdB ' uverz & "anamnese\quelle.mdb"
 FxDB = uVerz & "FaxeinP.mdb"
 Call Lese.ProgStart
' SET dbv = New DBVerb
' Call dbv.cnVorb("", "anamnesebogen", "Patientendaten")
 Select Case FenArt
  Case 0: Me.DateiBearbeiten.Caption = "&Datei bearbeiten"
  Case 1: Me.DateiBearbeiten.Caption = "&Briefe schreiben"
  Case 2: Me.DateiBearbeiten.Caption = "&Labore einlesen"
 End Select
 Call Me.RegLaden
 Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in form_load/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' form_load
#If False Then
'Function doBriefeBerichtspflicht(Optional ArztNr%)
' Dim QDat$, dszahl&, angefangen%, erg&, ergS$, FDT As Date
' ON Error Resume Next
'' QDat = getLDatei(uverz, "Listenausgabe_EBM-Ziffern*.xls")
'' IF Err.Number <> 0 THEN
''  MsgBox "Fehler beim Einladen der Datei " & uverz & "Listenausgabe_EBM-Ziffern*.xls"
''  Exit Function
'' END IF
' QDat = uVerz & "Listenausgabe_Listen.xls"
' ergS = Dir(QDat)
' IF ergS <> "" THEN
'  FDT = FileDateTime(QDat)
'  erg = MsgBox("Die Datei " & QDat & " stammt von: " & FDT & "." & vbCrLf & "Wollen Sie zu allen dort enthaltenen Patienten Briefe schreiben und abrechnen?", vbOKCancel, "Rückfrage")
'  IF erg = vbNo THEN Exit Function
' Else
'  MsgBox ("Datei " & QDat & " nicht gefunden, breche ab!")
'  Exit Function
' END IF
' ON Error GoTo fehler
'' aus FUNCTION EmailsImport(EmDatei$)
' Dim con As New ADODB.Connection  ' Connection
' Dim rNa As New ADODB.Recordset
' Dim rEx As New ADODB.Recordset
' Dim rX As New ADOX.Catalog
'' Const EmDatei$ = pverz & "Patientenübergreifendes\Emails.xls" ' Excel-Datei mit Suche aus Turbomed "*@*"
' ON Error GoTo fehler
' IF LenB(DBCnS) = 0 OR DBCn = "" THEN
''   Call frm.ConstrFestleg(2)
'   Call acon(quelleT, qDtb)
' END IF
' IF QDat <> "" THEN
' con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;HDR=No;IMEX=1"";Data Source=" & QDat & ";" ' TABLE=Adressen$"
' Dim runde%, i%, zFeld$, lFeld$, obAnfang%, pNr%, pRoh
'  rX.ActiveConnection = con
'  rEx.Open "`" & rX.Tables(rX.Tables.Count - 1).Name & "`", con ' Hier Excel, nicht lies.obmysql = 0!
'  Do While Not rEx.EOF
''  Debug.Print runde
'   IF obAnfang THEN
''    ON Error Resume Next
'     pRoh = rEx.Fields(0)
'     IF InStrB(rEx.Fields(2), "Es ist kein Bericht vorhanden") > 0 THEN
'' Hier jetzt Vorname, Name und Geburtsdatum extrahieren, um den Patienten zu finden!
'       Dim VN$, NN$, gd$, gesn$, leid$
'       Dim p1%, p2%, p3%
'       p1 = InStr(pRoh, " ")
'       p2 = InStr(pRoh, "(")
'       p3 = InStr(pRoh, " - ")
'       VN = LEFT(pRoh, p1 - 1)
'       gesn = replace$(LEFT(pRoh, p2 - 2), "zzz", "")
'       NN = Mid$(pRoh, p1 + 1, InStr(p1 + 1, pRoh, " ") - p1 - 1)
'       gd = Mid$(pRoh, p2 + 1, InStr(pRoh, ")") - p2 - 1)
'       leid = Mid$(pRoh, p3 + 3, InStr(pRoh, ";") - 3 - p3)
'       dszahl = dszahl + 1
'       IF NN LIKE "zzz*" THEN NN = Mid$(NN, 4)
'       SET rNa = Nothing
'       sql = "SELECT * FROM `namen` WHERE CONCAT(titel¡iif(titel='','',' ')¡nvorsatz¡iif(nvorsatz='','',' ')¡vorname¡' '¡nachname) = '" & gesn & "' AND gebdat = " & DatFor_k(gd) & ""
'       IF Forms(0).obMySQL THEN
'        sql = replace$(replace$(sql, "¡", ","), "iif(", "if(")
'       Else
'        sql = replace$(replace$(sql, "concat", ""), "¡", " & ")
'       END IF
'       Call rNa.Open("SELECT COUNT(0) FROM (" & sql & ") innen", DBCn, adOpenStatic, adLockOptimistic)
'
'       IF rNa.Fields(0) <> 1 THEN
'        MsgBox "Falsche Zahl an Datensätzen für " & VN & " " & NN & ", geb. " & gd & ": " & rNa.Fields(0)
'       Else
'        SET rNa = Nothing
'        Call rNa.Open(sql, DBCn, adOpenStatic, adLockOptimistic)
'        IF Not angefangen THEN
'          Call LeistungsExport0
'          angefangen = True
'        END IF
''        Call LeistungsExport1(rNa!Pat_id, "01601", CDate(leid), CDate("18:00"))
'        Call LeistungsExport1(rNa!Pat_id, "40120", CDate(leid), CDate("18:00"), , , ArztNr)
'        Call tubriefStandalone(rNa!Pat_id, True)
'       END IF
'     END IF
'   ElseIf NOT ISNULL(rEx.Fields(0)) AND NOT ISNULL(rEx.Fields(1)) AND NOT ISNULL(rEx.Fields(2)) THEN
'     obAnfang = True
'   END IF
'  runde = runde + 1
'  rEx.MoveNext
'  Loop
' END IF
' ON Error Resume Next
' rEx.Close
' Forms(0).Ausgeb "Fertig mit Nachtragen aus '" & QDat & "' von " & dszahl & " Datensätzen ", True
' IF angefangen THEN
'  Close #310
'  MsgBox "Datei " & LeistBDT & " neu mit den Leistungen zu den Briefen in " & QDat & " erstellt!"
' END IF
' Exit Function
'fehler:
'  Dim FNr&, FLastDLLError&, FSource$, FDescr$
'  FNr = Err.Number
'  FLastDLLError = Err.LastDllError
'  FSource = Err.source
'  FDescr = Err.Description
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(ISNULL(FSource), vNS, CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in doVorhandene/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' doBriefeBerichtspflicht
#End If

Private Sub DateiBearbeiten_Click()
' Const XStra = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
 Dim sql$, ka%, ke%, runde%, angefangen%, obAnfang%, rAF&
 Dim rEx As New ADODB.Recordset, rs As New ADODB.Recordset
 Select Case FenArt
  Case 1
   Open Me.Datei For Input As #394
   Dim pat_idS$
   Do While Not EOF(394)
    Input #394, pat_idS
    If IsNumeric(pat_idS) Then
     Call tubriefStandalone(CLng(pat_idS), True, pVerz & "unkorrigiert\")
    End If
   Loop
   Close #394
  Case 0, 2
' Do While Conn = ""
'  Conn = dbv.cnVorb("", "anamnesebogen", "Patientendaten")
' Loop
' DBCnS = Conn
' SET DBCn = Nothing
' DBCnOpen Conn
 xlsopen rEx, Me.Datei
 Dim kopfz%, pRoh$
 kopfz = 0
 Do While Not rEx.EOF
   If obAnfang Then
'  IF ISNULL(rEx.Fields(0)) OR rEx.Fields(0) = "0" OR rEx.Fields(0) = "" THEN kopfz = 1
'  IF kopfz >= 0 AND NOT ISNULL(rEx.Fields(0)) THEN kopfz = kopfz + 1
'  IF kopfz > 0 THEN
    If IsNull(rEx.Fields(1)) Then Exit Do ' 8.1.08
    Select Case FenArt
     Case 2
      Debug.Print rEx.Fields(0), rEx.Fields(1), rEx.Fields(2), rEx.Fields(3)
      Dim rsl As New ADODB.Recordset
      Set rsl = Nothing
      myFrag rsl, "SELECT COUNT(0) ct FROM `laborparameter` WHERE abkü = '" & UmwfSQL(rEx.Fields(0)) & "' AND langtext = '" & UmwfSQL(rEx.Fields(1)) & "' AND einheit = '" & UmwfSQL(rEx.Fields(2)) & "'"
      If rsl!ct = 1 Then
       myEFrag "UPDATE `laborparameter` SET labor = '" & UmwfSQL(rEx.Fields(3)) & "' WHERE abkü = '" & UmwfSQL(rEx.Fields(0)) & "' AND langtext = '" & UmwfSQL(rEx.Fields(1)) & "' AND einheit = '" & UmwfSQL(rEx.Fields(2)) & "'", rAF
      ElseIf rsl!ct > 1 Then
       Stop
      End If
     Case Else
    If rEx.Fields(0) <> "" And Not IsNumeric(rEx.Fields(0)) Then
     pRoh = rEx.Fields(0)
     If InStrB(rEx.Fields(2), "Es ist kein Bericht vorhanden") > 0 Then
      ka = InStr(pRoh, "(")
      ke = InStr(pRoh, ")")
      If ka > 0 And ke > 0 Then
       sql = "SELECT pat_id FROM `namen` WHERE CONCAT(vorname¡' '¡nvorsatz¡if(nvorsatz<>'',' ','')¡nachname¡' ('¡DATE_FORMAT(gebdat,'%d.%m.%Y')) = '" & Left$(pRoh, ke - 1) & "'"
       If InStrB(UCase$(dbv.ODBC), "MYSQL") > 0 Then
        sql = REPLACE$(sql, "¡", ",")
       Else
        sql = REPLACE$(REPLACE$(sql, "concat", ""), "¡", " & ")
        sql = REPLACE$(sql, "DATE_FORMAT(", "format$(")
        sql = REPLACE$(sql, "if(", "iif(")
        sql = REPLACE$(sql, "%d.%m.%Y", "dd.mm.yyyy")
       End If
       Set rs = Nothing
       myFrag rs, sql
       If rs.EOF Then
        MsgBox "Stop in DateiBearbeitenClick: " & vbCrLf & "rs.EOF" & vbCrLf & "sql: " & sql
        Stop
       Else
        If Not angefangen Then
'         Call LeistungsExport0
         angefangen = True
        End If
        Call teste(rs!Pat_id)
       End If
      End If
     End If
    End If
    End Select
  ElseIf Not IsNull(rEx.Fields(0)) And Not IsNull(rEx.Fields(1)) And Not IsNull(rEx.Fields(2)) Then
   obAnfang = True
  End If
  runde = runde + 1
  rEx.Move 1
 Loop
 If angefangen Then
'  Close #310
'  MsgBox "Datei " & LeistBDT & " neu mit den Leistungen zu den Briefen in " & Me.Datei & " erstellt (" & runde & " Runden)!"
 Else
  MsgBox "Fertig mit DateiBearbeiten ohne etwas zu tun"
 End If
 End Select
 Exit Sub
fehler:
  Dim FNr&, FLastDLLError&, FSource$, FDescr$
  FNr = Err.Number
  FLastDLLError = Err.LastDllError
  FSource = Err.source
  FDescr = Err.Description
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), vNS, CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in DateiBearbeiten_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DateiBearbeiten_Click()

Public Sub teste(Pat_id&)
 Dim rs As New ADODB.Recordset, rl As New ADODB.Recordset
 Dim BDT As New BDTSchreib
' Static obstart%
 myFrag rs, "SELECT 0 FROM `faelle` f LEFT JOIN `leistungen` l ON f.fid = l.fid WHERE f.pat_id = " & Pat_id & " AND leistung IN (01600,01601) AND quartal = '42007';"
 If rs.EOF Then
  Set rs = Nothing
  myFrag rs, "SELECT SchGr,FID FROM `faelle` f WHERE f.pat_id = " & Pat_id & " AND quartal = '42007'"
  If rs.EOF Then
   MsgBox "Stop in teste: " & vbCrLf & "rs.EOF" & vbCrLf & "sql: " & sql & vbCrLf & "Pat_id: " & Pat_id
   Stop
  ElseIf rs!SchGr <> "00" Then
   Debug.Print Pat_id
   Set rl = Nothing
   myFrag rl, "SELECT * FROM `leistungen` WHERE fid = " & rs!FID & " AND leistung IN (33012,33060,33061,33042) ORDER BY zeitpunkt DESC"
   Dim Datum As Date
   If rl.EOF Then Datum = #12/29/2007# Else Datum = rl!Zeitpunkt
   
'   IF Pat_id = 792 THEN obStart = True
'   Call tubriefStandalone(Pat_id, True, pverz & "unkorrigiert\")
   If Not BDT.Start(hVerz, "Leist") Then
    Exit Sub
   End If
   Call BDT.BDTKopf
   Open pVerz & "unkorrigiert\Nochzuschreiben.txt" For Append As #334
   Print #334, Pat_id
   Close #334
'   IF obstart THEN
'    Call LeistungsExport1(BDT, Pat_id, "01601", Datum, CDate("18:00"))
    Call LeistungsExport1(BDT, Pat_id, "40110", Datum, CDate("18:00"), , 0)
    '   END IF
  End If
 End If
End Sub ' teste

Public Sub RegLaden()
 Dim neuS$, neuB&
 On Error Resume Next
 changeStill = True
 RegPos = RegWurzel & App.EXEName & "\DBVerb"
 Select Case FenArt
  Case 0: neuS = fWertLesen(HCU, RegPos, "Datei")
  Case 1: neuS = fWertLesen(HCU, RegPos, "BDatei")
 End Select
 If neuS <> "" Then Me.Datei = neuS
 neuS = fWertLesen(HCU, RegPos, "Conn")
 
 If neuS <> "" Then Me.Conn = neuS
 changeStill = False
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegLaden/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub

Public Sub RegSpeichern()
 On Error GoTo fehler
 Select Case FenArt
  Case 0: Call fStSpei(HCU, RegPos, "Datei", Me.Datei)
  Case 1: Call fStSpei(HCU, RegPos, "BDatei", Me.Datei)
 End Select
 Call fStSpei(HCU, RegPos, "Conn", Me.Conn)
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegSpeichern/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Me.RegSpeichern
End Sub
