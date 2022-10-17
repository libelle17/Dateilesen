Attribute VB_Name = "HAAkt"
Option Explicit
' Public FSO As New FileSystemObject ' 15.5.22 auskommentiert, da hier doppelt
Const LÜDatei$ = "Listenausgabe_Überweiser.xls"
Const Verzeichnis$ = "p:\patientenübergreifendes\"
'Public Const XCon0 = "Driver={Microsoft Excel Driver (*.xls)};Dbq=" & LÜDatei & ";"
Public Const XStra = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public Const XStrb = ";Extended Properties=""Excel 8.0;HDR=no;IMEX=1"""
Const LIUEZ$ = "liuez" ' Listenausgabe Überweiser Ziel
Const LIUET$ = "liuet" ' Listenausgabe Überweiser temporär
Const aktVorw$ = "08131"
Const ausgD$ = "\HAAktAusg.txt"
' Public plzVz$, vVerz$, uVerz$, pVerz$ ' 15.5.22 auskommentiert, da hier doppelt
' Public Const vns$ = vbNullString ' 15.5.22 auskommentiert, da hier doppelt
' Public FPos&, FNr&, obStart% ' 15.5.22 auskommentiert, da hier doppelt
Dim fI As New FürIcon
' Public userprof$' 15.5.22 auskommentiert, da hier doppelt
' Public obTrans% ' ob BeginTrans für DBCn aufgerufen wurde' 15.5.22 auskommentiert, da hier doppelt
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwndOwner As Long, ByVal lpszPath As String, ByVal nFolder As Long, ByVal fCreate As Long) As Long
 
'SHGetFolderPath dwFlags-Konstanten
Private Const SHGFP_TYPE_CURRENT = 0 'Aktueller Verzeichnispfad des virtuellen Verzeichnisses
Private Const SHGFP_TYPE_DEFAULT = 1 'Standardpfad des virtuellen Verzeichnisses
 
'SHGetFolderPath Rückgabe-Konstanten
Private Const S_OK = &H0 'Die Funktion war erfolgreich
Private Const S_FALSE = &H1 'Die Funktion ist gescheitert
Private Const E_INVALIDARG = &H80070057 'Es wurden ungültige Parameter übergeben
 
'SHGetFolderPath/SHGetSpecialFolderPath nFolder-Konstanten
Private Const CSIDL_FLAG_CREATE = &H8000 '(Win 2000) Falls das Verzeichnis nicht Eistiert wird es erstellt
Private Const CSIDL_FLAG_DONT_VERIFY = &H4000 '(Win 2000) Benutzt den Ordner auch wenn der CSIDL nicht existiert
Private Const CSIDL_ADMINTOOLS = &H30 '(Win 2000) Ermittelt das Verzeichnis "Administrations Tools"
Private Const CSIDL_ALTSTARTUP = &H1D 'Ermittelt das Verzeichnis "Alternatives Startverzeichnis"
Private Const CSIDL_APPDATA = &H1A 'Ermittelt das Verzeichnis "AppData"
Private Const CSIDL_BITBUCKET = &HA 'Ermittelt das Verzeichnis "Papierkorb"
Private Const CSIDL_COMMON_ADMINTOOLS = &H2F '(Win 2000) Ermittelt das Verzeichnis "Administartions Tools für alle Benutzer"
Private Const CSIDL_COMMON_ALTSTARTUP = &H1D '(Win 2000) Ermittelt das Verzeichnis "Alternatives Startverzeichnis für alle Benutzer"
Private Const CSIDL_COMMON_APPDATA = &H23 '(Win NT/2000) Ermittelt das Verzeichnis "AppData für alle Benutzer"
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 '(Win NT/2000) Ermittelt das Verzeichnis "Desktop für alle Benutzer"
Private Const CSIDL_COMMON_DOCUMENTS = &H2E '(Win NT/2000) Ermittelt das Verzeichnis "Dokumente für alle Benutzer"
Private Const CSIDL_COMMON_FAVORITES = &H1F '(Win NT/2000) Ermittelt das Verzeichnis "Favoriten für alle Benutzer"
Private Const CSIDL_COMMON_PROGRAMS = &H17 '(Win NT/2000) Ermittelt das Verzeichnis "Startmenü/Programme für alle Benutzer"
Private Const CSIDL_COMMON_STARTMENU = &H16 '(Win NT/2000) Ermittelt das Verzeichnis "Startmenü für alle Benutzer"
Private Const CSIDL_COMMON_STARTUP = &H18 '(Win NT/2000) Ermittelt das Verzeichnis "Startmenü/Autostart für alle Benutzer"
Private Const CSIDL_COMMON_TEMPLATES = &H2D '(Win NT/2000) Ermittelt das Verzeichnis "Templates für alle Benutzer"
Private Const CSIDL_CONTROLS = &H3 'Ermittelt das Verzeichnis "Systemsteuerung"
Private Const CSIDL_COOKIES = &H21 'Ermittelt das Verzeichnis "Cookies"
Private Const CSIDL_DESKTOP = &H0 'Ermittelt das Verzeichnis "Desktop"
Private Const CSIDL_DESKTOPDIRECTORY = &H10 'Ermittelt das Verzeichnis "Desktop-Verzeichnis"
Private Const CSIDL_DRIVES = &H11 'Ermittelt das Verzeichnis "Arbeitsplatz"
Private Const CSIDL_FAVORITES = &H6 'Ermittelt das Verzeichnis "Favoriten"
Private Const CSIDL_FONTS = &H14 'Ermittelt das Verzeichnis "Fonts"
Private Const CSIDL_HISTORY = &H22 'Ermittelt das Verzeichnis "History"
Private Const CSIDL_INTERNET = &H1 'Ermittelt das Verzeichnis "Internet"
Private Const CSIDL_INTERNET_CACHE = &H20 'Ermittelt das Verzeichnis "Internet Cache"
Private Const CSIDL_LOCAL_APPDATA = &H1C '(nur IE ab version 5.0) Ermittelt das Verzeichnis "Local AppData"
Private Const CSIDL_MYPICTURES = &H27 '(nur IE ab version 5.0) Ermittelt das Verzeichnis "Eigene Bilder"
Private Const CSIDL_NETHOOD = &H13 'Ermittelt das Verzeichnis "Netzwerkumgebung"
Private Const CSIDL_NETWORK = &H12 'Ermittelt das Verzeichnis "Gesamtes Netzwerk"
Private Const CSIDL_PERSONAL = &H5 'Ermittelt das Verzeichnis "Eigene Dokumente"
Private Const CSIDL_PRINTERS = &H4 'Ermittelt das Verzeichnis "Drucker"
Private Const CSIDL_PRINTHOOD = &H1B 'Ermittelt das Verzeichnis "Drucker-Verzeichnis"
Private Const CSIDL_PROFILE = &H28 '(nur IE ab version 5.0) Ermittelt das Verzeichnis "Profile"
Private Const CSIDL_PROGRAM_FILES = &H26 '(nur IE ab version 5.0) Ermittelt das Verzeichnis "Programme"
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B '(Win NT/2000) Ermittelt das Verzeichnis "Programme für alle Benutzer"
Private Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C '(Win 2000 - x86) Ermittelt das Verzeichnis "Programme für alle Benutzer"
Private Const CSIDL_PROGRAM_FILESX86 = &H2A '(Win 2000 - RISC) Ermittelt das Verzeichnis "Programme für alle Benutzer"
Private Const CSIDL_PROGRAMS = &H2 'Ermittelt das Verzeichnis "Startmenü/Programme"
Private Const CSIDL_RECENT = &H8 'Ermittelt das Verzeichnis "Startmenü/Dokumente"
Private Const CSIDL_SENDTO = &H9 'Ermittelt das Verzeichnis "Senden zu"
Private Const CSIDL_STARTMENU = &HB 'Ermittelt das Verzeichnis "Startmenü"
Private Const CSIDL_STARTUP = &H7 'Ermittelt das Verzeichnis "Autostart"
Private Const CSIDL_SYSTEM = &H25 '(nur IE ab version 5.0) Ermittelt das Verzeichnis "System"
Private Const CSIDL_SYSTEMX86 = &H29 '(Win 2000) Ermittelt das Verzeichnis "Systemx86"
Private Const CSIDL_TEMPLATES = &H15 'Ermittelt das Verzeichnis "Templates"
Private Const CSIDL_WINDOWS = &H24 '(nur IE ab version 5.0) Ermittelt das Verzeichnis "Windows"

'Verzeichnis ermitteln
Public Function SHFolder(art&)
   Dim RetVal&, Buffer$
   On Error Resume Next
   Buffer = Space(256)
   RetVal = -1
   RetVal = SHGetFolderPath(0, art, 0&, SHGFP_TYPE_CURRENT, Buffer)
   'Fehlgeschlagen oder Falsches Betriebsystem ?
   If RetVal = S_OK Then
     SHFolder = Left$(Buffer, InStr(1, Buffer, vbNullChar) - 1)
   End If
'   MsgBox Left$(Buffer, InStr(1, Buffer, vbNullChar) - 1)
End Function ' SHFolder(art&)

'Public Const XCon2 = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;HDR=No;IMEX=2"";Data Source=" & LÜDatei & ";" ' TABLE=Adressen$"
Function datformZ(DaT, obMySQL%) ' for vb-Datumsformat oder vb-double (#)
 On Error GoTo fehler
 Dim obkurz%
 If IsNull(DaT) Then
  datformZ = "null"
 ElseIf (obMySQL <> 0) Then
  datformZ = "'" + Format$(DaT, "yyyy-mm-dd hh:mm:ss") + "'"
 Else
  If VarType(DaT) = vbString Then
   If Len(DaT) < 10 Then
    obkurz = True
   End If
  Else
   If DaT - Int(DaT) = 0 Then
    obkurz = True
   End If
  End If
  If obkurz Then
   datformZ = "#" + Format$(DaT, "mm\/dd\/yy") + "#"
  Else
   datformZ = "#" + Format$(DaT, "mm\/dd\/yy hh:mm:ss") + "#"
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in datFormZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' datFormZ

' in Hausärzte_aus_Listenausgabe_Ueberweiser_einlesen_Click
Public Sub doHAAkt(frm As Lese)
 Dim aktDatei$, HAStr$
 Dim XCon As New ADODB.Connection
 Dim rEx As New ADODB.Recordset
 Dim rX As New ADOX.Catalog
 Dim rV1 As New ADODB.Recordset
 Dim rV2 As New ADODB.Recordset
 Dim rs As New ADODB.Recordset
 Dim rAF&, obüberhaupt%
 Dim ausg$, lauf&
 Dim F0 As File, f1 As File
 On Error GoTo fehler
 ' 1. fetlegen, von wo die Datei Listenausgabe_Überweiser.xls gelesen wird
 userprof = Environ("userprofile")
 SetProgV
 AnamneseVerZeichnis1$ = uVerz & "Anamnese\"
 StACCDB$ = AnamneseVerZeichnis1 & "quelle.mdb"
 StFxDB$ = uVerz & "FaxeinP.mdb"
 StFtDB$ = uVerz & "FotosinP.mdb"
 StOffDB$ = uVerz & "office.mdb"
 
' Ausg = Environ("tmp") & ausgD ' 15.5.22
' Open Ausg For Output As #399  ' 15.5.22
 frm.Ausgeb "Fange an mit doHAAkt ...", True
 Load fI
 
 aktDatei = SHFolder(&H5) & "\" & LÜDatei ' eigene Dokumente
 If FSO.FileExists(aktDatei) Then
  Set F0 = FSO.GetFile(aktDatei)
  If FSO.FileExists(Verzeichnis & LÜDatei) Then
   Set f1 = FSO.GetFile(Verzeichnis & LÜDatei)
   If f1.DateLastModified > F0.DateLastModified Then
    aktDatei = Verzeichnis & LÜDatei
   End If
  End If
 Else
  aktDatei = Verzeichnis & LÜDatei
 End If
Dim fld$
Dim dbknr As ConDtb
' Dim LVobMySQL As Boolean
#Const vordef = False
#If Not vordef Then
' Provider=MSDASQL.1;Extended Properties="DATABASE=kvaerzte;DRIVER={MySQL ODBC 5.1 Driver};OPTION=3;PWD=***REMOVED***;PORT=0;SERVER=linux1;UID=praxis"
 Dim QCn() As ADODB.Connection
 Dim HCn() As ADODB.Connection
 Dim QCns$()
 Dim HCns$()
 Dim ccol As New Collection, Cpt, KVÄDB$
 If FSO.FileExists(aktDatei) Then
  ReDim QCn(0)
  ReDim HCn(0)
  ReDim QCns(0)
  ReDim HCns(0)
' das geht alles nicht mehr mit 64 bit oder wie auch immer: netServerEnum liefert Fehler 6118
'  HolServer
'  SET ccol = fI.dbv.ShowAllDomains
'  If ccol.Count = 0 Then
'    Debug.Print "Keine Domain gefunden"
    ccol.Add "LINUX1"
'  End If
  For Each Cpt In ccol
   If LCase(Cpt) = "linux1" Then GoTo gefunden
  Next Cpt
  ccol.Add "LINUX1"
gefunden:
  For Each Cpt In ccol
   If Cpt = "LINUX1" Then
   Dim MyCn As ADODB.Connection, MyCn2 As New ADODB.Connection, rdb As ADODB.Recordset, rHa As New ADODB.Recordset
   Set MyCn = New ADODB.Connection
   On Error Resume Next
   Err.Clear
   Dim Port&
   For Port = 0 To 3307 Step 3307
    Debug.Print Cpt
    Dim CnStr$
    CnStr = "Provider=MSDASQL.1;Extended Properties=DRIVER={" & ODBCStr() & "};OPTION=3;PWD=***REMOVED***;PORT=" & Port & ";SERVER=" & Cpt & ";UID=praxis;"
    Set MyCn = Nothing
    Set MyCn = New ADODB.Connection
    MyCn.Open CnStr
    If Err.Number <> 0 Then
     Debug.Print "Fehler beim Öffnen der Verbindung: " & CnStr & " " & Err.Description
     frm.Ausgeb "Fehler beim Öffnen der Verbindung: " & CnStr & " " & Err.Description, True
     Err.Clear
     CnStr = "DRIVER={" & ODBCStr() & "};server=" & Cpt & ";option=0;uid=praxis;pwd=***REMOVED***;"
     Set MyCn = Nothing
     Set MyCn = New ADODB.Connection
     MyCn.Open CnStr
    End If
    If Err.Number = 0 Then
     Dim obenthalten%
     Debug.Print MyCn.ConnectionString
     DoEvents
     obenthalten = False
#If langsam Then
     Set rdb = MyCn.Execute("SELECT schema_name FROM information_schema.`SCHEMATA` S")
#Else
     Set rdb = MyCn.Execute("SHOW DATABASES")
#End If
     Do While Not rdb.EOF
      fld = rdb.Fields(0)
      MyCn.Execute "USE `" & fld & "`"
      Debug.Print "USE `" & fld & "`"
      If Err.Number <> 0 Then
'       Print #399, "Computer: " & Cpt & ", Datenbank : " & fld & ": Fehler: " & Err.Description
       frm.Ausgeb "Computer: " & Cpt & ", Datenbank : " & fld & ": Fehler: " & Err.Description, True
      Else
       Set rHa = Nothing
 '      rha.Open "SELECT * FROM `hae` LIMIT 1", MyCn, adOpenStatic, adLockReadOnly
'       rHa.Open "SELECT table_name FROM information_schema.`TABLES` T WHERE table_schema = '" & fld & "' AND table_name = 'hae'", MyCn, adOpenStatic, adLockReadOnly
       myFrag rHa, "SELECT table_name FROM information_schema.`TABLES` T WHERE table_schema = '" & fld & "' AND table_name = 'hae'", adOpenStatic, MyCn
 '      If Err.Number = 0 Then
       If Not rHa.BOF Then
        KVÄDB = Left$(CnStr, Len(CnStr) - 1) & ";DATABASE=" & fld & ";"""
        obenthalten = True
        Exit Do
       Else
        Err.Clear
       End If
      End If
      rdb.Move 1
     Loop
     If obenthalten Then
      MyCn2.Open CnStr
      rdb.MoveFirst
      Dim runde&
      Do While Not rdb.EOF
       runde = runde + 1
       fld = rdb.Fields(0)
       Debug.Print "Runde: " & runde & ", rdb.fields(0): ", fld
       MyCn2.Execute "USE `" & fld & "`"
       If Err.Number = -2147467259 Then
        MyCn2.Close
        MyCn2.Open
        MyCn2.Execute "USE `" & fld & "`"
       End If
       If Err.Number = 0 Then
'        If fld = "quelle" Then Stop
        Set rHa = Nothing
 '       rha.Open "SELECT * FROM `anamnesebogen` LIMIT 1", MyCn, adOpenStatic, adLockReadOnly
'        rHa.Open "SELECT table_name FROM information_schema.`TABLES` WHERE table_schema = '" & fld & "' AND table_name = 'anamnesebogen'", MyCn, adOpenStatic, adLockReadOnly
        myFrag rHa, "SELECT table_name FROM information_schema.`TABLES` WHERE table_schema = '" & fld & "' AND table_name = 'anamnesebogen'", adOpenStatic, MyCn
        If Err.Number <> 0 Then
         Debug.Print Error.Description
         If Err.Number = -2147217887 Then
          MyCn.Close
          MyCn.Open
'          rHa.Open "SELECT table_name FROM information_schema.`TABLES` WHERE table_schema = '" & fld & "' AND table_name = 'anamnesebogen'", MyCn, adOpenStatic, adLockReadOnly
          myFrag rHa, "SELECT table_name FROM information_schema.`TABLES` WHERE table_schema = '" & fld & "' AND table_name = 'anamnesebogen'", adOpenStatic, MyCn
          If Err.Number <> 0 Then
           Debug.Print Error.Description
          End If
         End If
        End If
 '       If Err.Number = 0 Then
        If Not rHa.BOF Then
         Do While Not rHa.EOF
          obüberhaupt = True
          Set QCn(UBound(QCn)) = New ADODB.Connection
          Set HCn(UBound(HCn)) = New ADODB.Connection
          QCns(UBound(QCns)) = Left$(CnStr, Len(CnStr) - 1) & ";DATABASE=" & fld & ";"""
          HCns(UBound(HCns)) = KVÄDB
          QCn(UBound(QCn)).ConnectionString = QCns(UBound(QCns))
          HCn(UBound(HCn)).ConnectionString = KVÄDB
          QCn(UBound(QCn)).Open QCns(UBound(QCns))
          Dim zeigstr$
          If Err.Number <> 0 Then
           zeigstr = "Computer: " & Cpt & ", Datenbank : " & fld & ": Fehler beim 1. Verbinden"
           Err.Clear
          Else ' Err.Number <> 0 Then
           HCn(UBound(HCn)).Open
           If Err.Number <> 0 Then
            zeigstr = "Computer: " & Cpt & ", Datenbank : " & fld & ": Fehler beim 2. Verbinden"
            QCn(UBound(QCn)).Close
            Err.Clear
           Else ' Err.Number <> 0 Then
            zeigstr = "Computer: " & Cpt & ", Datenbank : " & fld & ": Verbunden (Nr." & UBound(HCn) & "):" & CurDB(HCn(UBound(HCn))) & " / " & CurDB(QCn(UBound(QCn)))
            QCn(UBound(HCn)).Close
            HCn(UBound(QCn)).Close
            ReDim Preserve HCn(UBound(HCn) + 1)
            ReDim Preserve QCn(UBound(QCn) + 1)
            ReDim Preserve HCns(UBound(HCns) + 1)
            ReDim Preserve QCns(UBound(QCns) + 1)
            obenthalten = True
           End If ' Err.Number <> 0 Then
          End If ' Err.Number <> 0 Then
          Debug.Print zeigstr
'          Print #399, zeigstr
          frm.Ausgeb zeigstr, True
          rHa.MoveNext
         Loop ' Do While Not rHa.EOF
        Else ' Not rHa.BOF Then
'        Print #399, "Computer: " & Cpt & ", Datenbank : " & fld & ": Keine geeignete Anamnesebogentabelle"
         Err.Clear
        End If ' Not rHa.BOF Then
       End If ' Err.Number = 0 nach use rdbfields(0) Then
       rdb.Move 1
      Loop ' While Not rdb.EOF
     Else ' obenthalten Then
'      Print #399, "Computer: " & Cpt & ", Keine Datenbank mit geeigneter Hausarzttabelle `hae`"
      frm.Ausgeb "Computer: " & Cpt & ", Keine Datenbank mit geeigneter Hausarzttabelle `hae`", True
     End If ' obenthalten Then
     Exit For
    End If ' If Err.Number = 0 Then
   Next Port
   On Error GoTo fehler
   End If
  Next Cpt
  If obüberhaupt Then
' 6.12.09: Access raus
'  SET QCn(UBound(QCn)) = acon(quelleT, accDtb)
  If QCn(UBound(QCn)) Is Nothing Then
   If UBound(HCn) > 0 Then
    ReDim Preserve HCn(UBound(HCn) - 1)
    ReDim Preserve HCns(UBound(HCns) - 1)
   End If
   If UBound(QCn) > 0 Then
    ReDim Preserve QCn(UBound(QCn) - 1)
    ReDim Preserve QCns(UBound(QCns) - 1)
   End If
  Else ' QCn(UBound(QCn)) Is Nothing Then
' 6.12.09: Access raus
'   SET HCn(UBound(QCn)) = acon(HaT, accDtb)
   If HCn(UBound(HCn)) Is Nothing Then
    ReDim Preserve HCn(UBound(HCn) - 1)
    ReDim Preserve QCn(UBound(QCn) - 1)
    ReDim Preserve HCns(UBound(HCns) - 1)
    ReDim Preserve QCns(UBound(QCns) - 1)
   Else ' HCn(UBound(HCn)) Is Nothing Then
'    Print #399, "Verbunden (Nr." & UBound(HCn) & "):" & DefDB(HCn(UBound(QCn))) & " / " & DefDB(QCn(UBound(HCn)))
    frm.Ausgeb "Verbunden (Nr." & UBound(HCn) & "):" & DefDB(HCn(UBound(QCn))) & " / " & DefDB(QCn(UBound(HCn))), True
    On Error Resume Next
    QCn(UBound(QCn)).Close
    HCn(UBound(HCn)).Close
    On Error GoTo fehler
   End If ' HCn(UBound(HCn)) Is Nothing Then
  End If ' QCn(UBound(QCn)) Is Nothing Then
#Else
 Dim QCn(accDtb To q2Dtb) As ADODB.Connection
 Dim HCn(accDtb To q2Dtb) As ADODB.Connection
 Dim QCns$(accDtb To q2Dtb)
 Dim HCns$(accDtb To q2Dtb)
 If FSO.FileExists(aktDatei) Then
  For dbknr = accDtb To q2Dtb ' 1 to 4
   Set QCn(dbknr) = acon(quelleT, dbknr)
   Set HCn(dbknr) = acon(HaT, dbknr)
   QCn(dbknr).Close
   HCn(dbknr).Close
  Next dbknr
#End If

#Const obdebug = False
#If Not obdebug Then
  
 For dbknr = LBound(QCn) To UBound(QCn)
  QCn(dbknr).Open QCns(dbknr)
  HCn(dbknr).Open HCns(dbknr)
  LVobMySQL = InStrB(UCase$(QCns(dbknr)), "MYSQL") <> 0  '(Not (cDtb = accDtb))
  LVobMySQL = True ' 11.10.15
  Call Zinit(LVobMySQL)
 
  Dim csql$
  If InStrB(Command, "neu") <> 0 Then
   csql = "DROP TABLE " & ifexists & " `" & LIUET & "`;"
   On Error Resume Next
   Call QCn(dbknr).Execute(csql)
   On Error GoTo fehler
  End If
  Dim kvnrlen$
  kvnrlen = QCn(dbknr).Execute("SELECT character_maximum_length FROM information_schema.columns WHERE table_schema='" & QCn(dbknr).DefaultDatabase & "' AND TABLE_NAME='faelle' AND COLUMN_NAME='übwvkvnr'").Fields(0)
  csql = "CREATE TABLE " & ifnotexists & " `" & LIUET & "` (name " & sqlText & "(53), vorname " & sqlText & "(35), titelt " & sqlText & "(40), fachgruppe " & sqlText & "(202), strasse " & sqlText & "(40), plz " & sqlText & "(7), ort " & sqlText & "(30), telefon " & sqlText & "(40)" & IIf((LVobMySQL), " collate utf8mb4_german2_ci", vNS) & ", fax " & sqlText & "(40) " & IIf((LVobMySQL), " collate utf8mb4_german2_ci", vNS) & ", kvnr " & sqlText & "(" & kvnrlen & ")" & IIf((LVobMySQL), " collate utf8mb4_german2_ci", vNS) & ", aktdat date, id " & sqlAutoIncr & " primary key," & _
         "überschrift " & sqlText & "(1), dbnr " & sqlText & "(7), bstelle " & sqlText & "(53), anrede " & sqlText & "(10), tel1 " & sqlText & "(40), tel2 " & sqlText & "(40), tel3 " & sqlText & "(40), tel4 " & sqlText & "(40), fax1 " & sqlText & "(40), fax2 " & sqlText & "(40), fax3 " & sqlText & "(40), email " & sqlText & "(70), zulg " & sqlText & "(150), " & _
         "arzttyp " & sqlText & "(2), gemmit " & sqlmemo & ", beme " & sqlmemo & ", dmpt2 " & sqlBool & ", dmpt1 " & sqlBool & ", geschlecht " & sqlText & "(1), titel " & sqlText & "(90), zusatz " & sqlText & "(20), ursp " & sqlText & "(30), aktzeit DATETIME" & _
         ", UNIQUE KEY `ident` (`name`,`vorname`,`plz`,`tel1`) " & vbCrLf & _
         ", KEY `telefon`(`telefon`))"
'         ", UNIQUE KEY `ident` (`name`,`vorname`,`plz`,`kvnr`))" ' ersetzt 1.1.17
  If (LVobMySQL) Then
   csql = csql & " Engine=InnoDB CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci"
  End If
  If Not LVobMySQL Then On Error Resume Next
  Call QCn(dbknr).Execute(csql)
  If Not LVobMySQL Then On Error GoTo fehler
  Call QCn(dbknr).Execute("DELETE FROM " & LIUET & " WHERE ISNULL(kvnr) OR kvnr=''", rAF) ' 28.12.15 hierher verschoben
  Call QCn(dbknr).Execute("DROP TABLE IF EXISTS `lauo`")
  Call QCn(dbknr).Execute("CREATE TABLE `lauo` (name " & sqlText & "(53),vorname " & sqlText & "(35),titelt " & sqlText & " (40),fachgruppe " & sqlText & "(202),strasse " & sqlText & "(40),plz " & sqlText & "(7),ort " & sqlText & "(30),telefon " & sqlText & "(40)" & IIf((LVobMySQL), " collate utf8_general_ci", vNS) & ",fax " & sqlText & "(40) " & IIf((LVobMySQL), " collate utf8_general_ci", vNS) & ", aktdat date, id " & sqlAutoIncr & " primary key)")
  QCn(dbknr).Close
  HCn(dbknr).Close
 Next ' dbknr = LBound(QCn) To UBound(QCn)
  
 Dim i%, v1sql$(5), pos%, posv%, rF0$, rF1$, ursp$
 Dim sqlakt$, f7$, f8$
 Dim HaeCon As New ADODB.Connection, rhae As New ADODB.Recordset
 Dim KVNr$, DBNr$, BStelle$, anrede$, tel1$, tel2$, tel3$, tel4$, fax1$, fax2$, fax3$, email$, zulg$, arzttyp$, gemmit$, beme$, dmpt2%, dmpt1%, geschlecht$, Titel$
 HaeCon.Open "Provider=MSDASQL.1;Extended Properties=""DATABASE=haerzte;DRIVER={" & ODBCStr() & "};OPTION=0;PWD=***REMOVED***;PORT=0;SERVER=linux1;UID=praxis"""
 
 Set XCon = Nothing
 XCon.Open XStra & aktDatei & XStrb
 rX.ActiveConnection = XCon
 Set rEx = Nothing
 rEx.Open "[" & rX.Tables(rX.Tables.COUNT - 1).name & "]", XCon ' Hier Excel, nicht obmysql = 0!
 Dim kopfz%
 kopfz = 0
 Dim zeile!
 frm.Ausgeb "Lese '" & aktDatei & "' ein ...", True
 
 Do While Not rEx.EOF
   zeile = zeile + 1
   frm.Ausgeb "Zeile: " & zeile, False
   KVNr = vNS
   rF0 = vNS
   rF1 = vNS
   f7 = vNS
   f8 = vNS
   
   If IsNull(rEx.Fields(0)) Or rEx.Fields(0) = "0" Then kopfz = 1
   If kopfz >= 1 And Not IsNull(rEx.Fields(0)) Then kopfz = kopfz + 1
   If kopfz > 2 Then
    If IsNull(rEx.Fields(7)) Then Exit Do ' 8.1.08
    If rEx.Fields(0) <> vNS And Not IsNumeric(rEx.Fields(0)) Then
     f7 = REPLACE(REPLACE(REPLACE(REPLACE(rEx.Fields(7), "-", vNS), "/", vNS), " ", vNS), "+", "00")
     If Left(f7, 1) <> "0" And f7 <> vNS Then
      If Len(f7) < 7 Then
       f7 = aktVorw & f7
      Else
       f7 = "0" & f7
      End If
     End If
     f8 = REPLACE(REPLACE(REPLACE(REPLACE(rEx.Fields(8), "-", vNS), "/", vNS), " ", vNS), "+", "00")
     If Left(f8, 1) <> "0" And f8 <> vNS Then
      If Len(f8) < 7 Then
       f8 = aktVorw & f8
      Else
       f8 = "0" & f8
      End If
     End If
      
     rF0 = rEx.Fields(0) ' Nachname
     rF1 = rEx.Fields(1) ' Vorname
     
     Set rhae = Nothing
'     If rF0 = "Winkhaus" Then Stop
     Dim sql$
     sql = "SELECT titel, Niederlassungsart arzttyp, group_concat(distinct a2.nachname) gemmit, mail email, bsnr kvnu, REPLACE(group_concat(distinct tel),'-','') tel, REPLACE(group_concat(distinct fax),'-','') fax, a.*, mail, IF(a.obweibl=1,'w','m') geschlecht, IF(a.obweibl=1,'Frau','Herrn') anrede, group_concat(distinct fachrichtung) zulg, " & _
           "MAX(IF(ISNULL((SELECT genehmigung FROM `haerzte`.`arzt_has_genehmigung` ahg1 LEFT JOIN `haerzte`.genehmigung g1 ON g1.idgenehmigung = ahg1.genehmigung_id WHERE ahg1.arzt_id = a2.idarzt AND genehmigung = 'DMP-DM1_Koordinierender Arzt_Hausarzt')),'0','1')) dmpt1, " & _
           "MAX(IF(ISNULL((SELECT genehmigung FROM `haerzte`.`arzt_has_genehmigung` ahg2 LEFT JOIN `haerzte`.genehmigung g2 ON g2.idgenehmigung = ahg2.genehmigung_id WHERE ahg2.arzt_id = a2.idarzt AND genehmigung = 'DMP-DM2_Koordinierender Arzt')),'0','1')) dmpt2 " & _
           "FROM `haerzte`.`arzt` a " & _
           "LEFT JOIN `haerzte`.`arzt_has_fachrichtung` ahf ON a.idarzt = ahf.arzt_id LEFT JOIN `haerzte`.fachrichtung fr ON ahf.fachrichtung_id = fr.idfachrichtung " & _
           "LEFT JOIN `haerzte`.`arzt_has_bs` ahb ON a.idarzt = ahb.arzt_id LEFT JOIN `haerzte`.bs ON bs.idbs = ahb.bs_id AND obneben=0 " & _
           "LEFT JOIN `haerzte`.`tel` ON `tel`.bs_id = bs.idbs " & _
           "LEFT JOIN `haerzte`.`fax` ON `fax`.bs_id = bs.idbs " & _
           "LEFT JOIN `haerzte`.`mail` ON `mail`.bs_id = bs.idbs " & _
           "LEFT JOIN `haerzte`.`arzt_has_bs` ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN `haerzte`.`arzt` a2 ON a2.idarzt = ahb2.arzt_id " & _
           "LEFT JOIN `haerzte`.`nlart` ON a.nlart_id = `nlart`.idnlart " & _
           "LEFT JOIN `haerzte`.`titel` t ON t.idtitel = a.titel_id " & _
           "WHERE a.nachname = '" & rF0 & "' AND a.vorname = '" & rF1 & "' AND plz = '" & rEx.Fields(5) & "'"
     rhae.Open sql, HaeCon, adOpenStatic, adLockReadOnly
     ursp = "+haerzte"
     If Not rhae.BOF And Not IsNull(rhae!kvnu) Then
      Dim telef$(), faxe$()
'      tel1 = telef(0)
      tel1 = vNS: tel2 = vNS: tel3 = vNS: tel4 = vNS
      If Not IsNull(rhae!tel) Then
       SplitNeu rhae!tel, ",", telef
       If UBound(telef) >= 0 Then tel1 = telef(0)
       If UBound(telef) > 0 Then tel2 = telef(1)
       If UBound(telef) > 1 Then tel3 = telef(2)
       If UBound(telef) > 2 Then tel4 = telef(3)
      End If
      fax1 = vNS: fax2 = vNS: fax3 = vNS
      If Not IsNull(rhae!fax) Then
       SplitNeu rhae!fax, ",", faxe
       If UBound(faxe) >= 0 Then fax1 = faxe(0)
       If UBound(faxe) > 0 Then fax2 = faxe(1)
       If UBound(faxe) > 1 Then fax3 = faxe(2)
      End If
     End If
     If rhae.BOF Or IsNull(rhae!kvnu) Then
      Set rhae = Nothing
      rhae.Open "SELECT * FROM kvaerzte.`hae` " & _
           "WHERE nachname = '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz = '" & rEx.Fields(5) & "'", HaeCon, adOpenStatic, adLockReadOnly
      ursp = "+kvaerzte.hae"
      If rhae.BOF Then
       Set rhae = Nothing
       rhae.Open "SELECT * FROM kvaerzte.`haealt` " & _
           "WHERE nachname = '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz = '" & rEx.Fields(5) & "'", HaeCon, adOpenStatic, adLockReadOnly
      ursp = "+kvaerzte.haealt"
       If rhae.BOF Then
        Set rhae = Nothing
        rhae.Open "SELECT if(geschlecht='w','Frau','Herrn') anrede, zulassungsgebiet zulg, kvnr kvnu, telefon tel1, '' tel2, '' tel3, '' tel4, telefax fax1, '' fax2, '' fax3, e_mail email, arzttyp, `Gemeinschaftspraxis mit` gemmit, schwerpunkt zulg, titel, geschlecht, -dmpt2 dmpt2, dmpt1 FROM quelle.`hausaerzte` " & _
           "WHERE nachname = '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz = '" & rEx.Fields(5) & "'", HaeCon, adOpenStatic, adLockReadOnly
        ursp = "+quelle.hausarzte"
       End If
      End If
      If Not rhae.BOF And Not IsNull(rhae!kvnu) Then
       If Not IsNull(rhae!tel2) Then tel2 = rhae!tel2 Else tel2 = vNS
       tel1 = rhae!tel1
       If Not IsNull(rhae!tel3) Then tel3 = rhae!tel3 Else tel3 = vNS
       If Not IsNull(rhae!tel4) Then tel4 = rhae!tel4 Else tel4 = vNS
       fax1 = rhae!fax1
       If Not IsNull(rhae!fax2) Then fax2 = rhae!fax2 Else fax2 = vNS
       If Not IsNull(rhae!fax3) Then fax3 = rhae!fax3 Else fax3 = vNS
      End If
     End If
     If Not rhae.BOF And Not IsNull(rhae!kvnu) Then
      If Not IsNull(rhae!zulg) Then zulg = rhae!zulg Else zulg = vNS
      anrede = rhae!anrede
      If Not IsNull(rhae!kvnu) Then KVNr = Left$(rhae!kvnu, 7) Else KVNr = vNS
      geschlecht = rhae!geschlecht
      If Not IsNull(rhae!gemmit) Then gemmit = rhae!gemmit Else gemmit = vNS
      If Not IsNull(rhae!email) Then email = rhae!email Else email = vNS
      If Not IsNull(rhae!arzttyp) Then arzttyp = rhae!arzttyp Else arzttyp = vNS
      If Not IsNull(rhae!dmpt1) Then dmpt1 = rhae!dmpt1 Else dmpt1 = 0
      If Not IsNull(rhae!dmpt2) Then dmpt2 = rhae!dmpt2 Else dmpt2 = 0
      If Not IsNull(rhae!Titel) Then Titel = rhae!Titel Else Titel = vNS
     End If
     If rhae.BOF Or IsNull(rhae!kvnu) Then
      anrede = IIf(InStrB(rEx.Fields(3), "FÄ") Or InStrB(rEx.Fields(3), "HÄ") Or InStrB(LCase$(rEx.Fields(3)), "ärztin"), "Frau", "Herrn")
      geschlecht = IIf(InStrB(rEx.Fields(3), "FÄ") Or InStrB(rEx.Fields(3), "HÄ"), "w", "m")
     End If
     For dbknr = LBound(QCn) To UBound(QCn)
      QCn(dbknr).Open QCns(dbknr)
      HCn(dbknr).Open HCns(dbknr)
      LVobMySQL = InStrB(UCase$(QCn(dbknr).ConnectionString), "MYSQL") <> 0  '(Not (cDtb = accDtb))
      LVobMySQL = True ' 11.10.15
      If LVobMySQL Then
      Call Zinit(LVobMySQL)
      sqlakt = "INSERT INTO `lauo` (name,vorname,titelt,fachgruppe,strasse,plz,ort,telefon,fax) VALUES (" & "'" & rF0 & "','" & rF1 & _
                "','" & rEx.Fields(2) & "','" & rEx.Fields(3) & "','" & rEx.Fields(4) & "','" & rEx.Fields(5) & "','" & _
                rEx.Fields(6) & "','" & f7 & "','" & f8 & "');"
      Call QCn(dbknr).Execute(sqlakt, rAF)
      
      pos = 0
      Set rs = Nothing
'      If rF0 = "Kasper" Then Stop
'      rs.Open "SELECT * FROM `" & liuet & "` WHERE name = '" & rF0 & "' AND vorname = '" & rF1 & "' AND kvnr = '" & KVNr & "'", QCn(dbknr), adOpenStatic, adLockReadOnly
      Dim Vorabfrage$
      Vorabfrage = "SELECT * FROM `" & LIUET & "` WHERE name = '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz ='" & rEx.Fields(5) & "' AND telefon = '" & f7 & "'"
      rs.Open Vorabfrage, QCn(dbknr), adOpenStatic, adLockReadOnly
      If rs.EOF Then
       sqlakt = "INSERT INTO `" & LIUET & "` (name,vorname,titelt,fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,DBNr, BStelle, anrede, tel1, tel2, tel3, tel4, fax1, fax2, fax3, email, zulg, arzttyp, gemmit, beme, dmpt2, dmpt1, geschlecht, titel, ursp, aktzeit) VALUES (" & "'" & rF0 & "','" & rF1 & _
                "','" & rEx.Fields(2) & "','" & rEx.Fields(3) & "','" & rEx.Fields(4) & "','" & rEx.Fields(5) & "','" & _
                rEx.Fields(6) & "','" & f7 & "','" & f8 & "'," & IIf(KVNr = "", "null", "'" & REPLACE$(KVNr, " ", "") & "'") & "," & datformZ(Now, True) & ",'" & DBNr & "','" & BStelle & "','" & anrede & "','" & tel1 & "','" & tel2 & "','" & tel3 & "','" & tel4 & "','" & fax1 & "','" & fax2 & "','" & fax3 & "','" & email & "','" & zulg & "','" & UCase$(Left$(arzttyp, 2)) & "','" & gemmit & "','" & beme & "'," & Abs(dmpt2) & "," & Abs(dmpt1) & ",'" & Left$(geschlecht, 1) & "','" & Titel & "','xls','" & Format(Now(), "yymmddhhmmss") & "');"
      Else
       If pos <> 0 Then
        rF0 = rF0 & "%"
       End If
       If KVNr = vNS Then
'        sqlakt = "UPDATE `" & liuet & "` SET titelt = '" & rEx.Fields(2) & "', fachgruppe='" & rEx.Fields(3) & "',strasse='" & rEx.Fields(4) & "',plz='" & rEx.Fields(5) & "',ort='" & rEx.Fields(6) & "',telefon='" & f7 & "',fax='" & f8 & "',aktdat=" & datformZ(Now, True) & ",DBNr='" & DBNr & "',BStelle='" & BStelle & "',Anrede='" & anrede & "',tel1='" & tel1 & "',tel2='" & tel2 & "',tel3='" & tel3 & "',tel4='" & tel4 & "',fax2='" & fax2 & "',fax3='" & fax3 & "',email='" & email & "',zulg='" & zulg & "',arzttyp='" & UCase$(Left$(arzttyp, 2)) & "',gemmit='" & gemmit & "',beme='" & beme & "',dmpt2=" & Abs(dmpt2) & ",dmpt1=" & Abs(dmpt1) & ",geschlecht='" & Left$(geschlecht, 1) & "',titel='" & Titel & "' WHERE name LIKE '" & rF0 & "' AND vorname = '" & rF1 & "'"
         sqlakt = "UPDATE `" & LIUET & "` SET titelt = '" & rEx.Fields(2) & "', fachgruppe='" & rEx.Fields(3) & "',strasse='" & rEx.Fields(4) & "',ort='" & rEx.Fields(6) & "',fax='" & f8 & "',aktdat=" & datformZ(Now, True) & ",DBNr='" & DBNr & "',BStelle='" & BStelle & "',Anrede='" & anrede & "',tel1='" & tel1 & "',tel2='" & tel2 & "',tel3='" & tel3 & "',tel4='" & tel4 & "',fax2='" & fax2 & "',fax3='" & fax3 & "',email='" & email & "',zulg='" & zulg & "',arzttyp='" & UCase$(Left$(arzttyp, 2)) & "',gemmit='" & gemmit & "',beme='" & beme & "',dmpt2=" & Abs(dmpt2) & ",dmpt1=" & Abs(dmpt1) & ",geschlecht='" & Left$(geschlecht, 1) & "',titel='" & Titel & "' WHERE name LIKE '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz ='" & rEx.Fields(5) & "' AND telefon = '" & f7 & "'"
       Else
'        sqlakt = "UPDATE `" & liuet & "` SET titelt = '" & rEx.Fields(2) & "', fachgruppe='" & rEx.Fields(3) & "',strasse='" & rEx.Fields(4) & "',plz='" & rEx.Fields(5) & "',ort='" & rEx.Fields(6) & "',telefon='" & f7 & "',fax='" & f8 & "',kvnr='" & KVNr & "',aktdat=" & datformZ(Now, True) & ",DBNr='" & DBNr & "',BStelle='" & BStelle & "',Anrede='" & anrede & "',tel1='" & tel1 & "',tel2='" & tel2 & "',tel3='" & tel3 & "',tel4='" & tel4 & "',fax2='" & fax2 & "',fax3='" & fax3 & "',email='" & email & "',zulg='" & zulg & "',arzttyp='" & UCase$(Left$(arzttyp, 2)) & "',gemmit='" & gemmit & "',beme='" & beme & "',dmpt2=" & Abs(dmpt2) & ",dmpt1=" & Abs(dmpt1) & ",geschlecht='" & Left$(geschlecht, 1) & "',titel='" & Titel & "' WHERE name LIKE '" & rF0 & "' AND vorname = '" & rF1 & "'"
        sqlakt = "UPDATE `" & LIUET & "` SET titelt = '" & rEx.Fields(2) & "', fachgruppe='" & rEx.Fields(3) & "',strasse='" & rEx.Fields(4) & "',ort='" & rEx.Fields(6) & "',fax='" & f8 & "',kvnr=" & IIf(KVNr = vNS, "null", "'" & REPLACE$(KVNr, " ", "") & "'") & ",aktdat=" & datformZ(Now, True) & ",DBNr='" & DBNr & "',BStelle='" & BStelle & "',Anrede='" & anrede & "',tel1='" & tel1 & "',tel2='" & tel2 & "',tel3='" & tel3 & "',tel4='" & tel4 & "',fax2='" & fax2 & "',fax3='" & fax3 & "',email='" & email & "',zulg='" & zulg & "',arzttyp='" & UCase$(Left$(arzttyp, 2)) & "',gemmit='" & gemmit & "',beme='" & beme & "',dmpt2=" & Abs(dmpt2) & ",dmpt1=" & Abs(dmpt1) & ",geschlecht='" & Left$(geschlecht, 1) & "',titel='" & Titel & "'" & IIf(KVNr = vNS, "", ",ursp=CONCAT(ursp,'" & ursp & "')") & " WHERE name LIKE '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz ='" & rEx.Fields(5) & "' AND telefon = '" & f7 & "'"
       End If
      End If ' rs.eof
      
      On Error Resume Next
      Call QCn(dbknr).Execute(sqlakt, rAF)
      If Err.Number <> 0 And LVobMySQL Then
       Debug.Print "Vorabfrage: ", Vorabfrage
       Debug.Print sqlakt
       Debug.Print Err.Number, Err.Description, "QCn(dbknr).ConnectionString = ", QCn(dbknr).ConnectionString
       frm.Ausgeb sqlakt, True
       frm.Ausgeb Err.Number & Err.Description & " QCn(dbknr).ConnectionString = '" & QCn(dbknr).ConnectionString & "'", True
      Else
      End If
#If alt Then
      Set rs = Nothing
      rs.Open "SELECT DISTINCT kvnu FROM kvaerzte.`hae` WHERE nachname = '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz = '" & rEx.Fields(5) & "' AND kvnu <> '" & REPLACE(KVNr, " ", "") & "'", QCn(dbknr), adOpenStatic, adLockReadOnly
      If rs.State <> 0 Then
       If Not rs.BOF Then
        Do While Not rs.EOF
         sqlakt = "INSERT INTO `" & LIUET & "` (name,vorname,titelt,fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,DBNr, BStelle, anrede, tel1, tel2, tel3, tel4, fax1, fax2, fax3, email, zulg, arzttyp, gemmit, beme, dmpt2, dmpt1, geschlecht, titel,ursp,aktzeit) VALUES (" & "'" & rF0 & "','" & rF1 & "','" & rEx.Fields(2) & "','" & rEx.Fields(3) & "','" & rEx.Fields(4) & "','" & rEx.Fields(5) & "','" & rEx.Fields(6) & "','" & f7 & "','" & f8 & "'," & IIf(rs!KVNr = "", "null", "'" & REPLACE$(rs!KVNr, " ", "") & "'") & "," & datformZ(Now, True) & ",'" & DBNr & "','" & BStelle & "','" & anrede & "','" & tel1 & "','" & tel2 & "','" & tel3 & "','" & tel4 & "','" & fax1 & "','" & fax2 & "','" & fax3 & "','" & email & "','" & zulg & "','" & UCase$(Left$(arzttyp, 2)) & "','" & gemmit & "','" & beme & "'," & Abs(dmpt2) & "," & Abs(dmpt1) & ",'" & Left$(geschlecht, 1) & "','" & Titel & "','kvaerzte.hae','" & Format(Now(), "yymmddhhmmss") & "');"
         Call QCn(dbknr).Execute(sqlakt, rAF)
         rs.Move 1
        Loop
       End If
      End If
      Set rs = Nothing
      rs.Open "SELECT DISTINCT kvnu FROM kvaerzte.`haealt` WHERE nachname = '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz = '" & rEx.Fields(5) & "' AND kvnu <> '" & REPLACE$(KVNr, " ", "") & "'", QCn(dbknr), adOpenStatic, adLockReadOnly
      If rs.State <> 0 Then
       If Not rs.BOF Then
        Do While Not rs.EOF
         sqlakt = "INSERT INTO `" & LIUET & "` (name,vorname,titelt,fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,DBNr, BStelle, anrede, tel1, tel2, tel3, tel4, fax1, fax2, fax3, email, zulg, arzttyp, gemmit, beme, dmpt2, dmpt1, geschlecht, titel) VALUES (" & "'" & rF0 & "','" & rF1 & "','" & rEx.Fields(2) & "','" & rEx.Fields(3) & "','" & rEx.Fields(4) & "','" & rEx.Fields(5) & "','" & rEx.Fields(6) & "','" & f7 & "','" & f8 & "'," & IIf(rs!KVNr = "", "null", "'" & REPLACE$(rs!KVNr, " ", "") & "'") & "," & datformZ(Now, True) & ",'" & DBNr & "','" & BStelle & "','" & anrede & "','" & tel1 & "','" & tel2 & "','" & tel3 & "','" & tel4 & "','" & fax1 & "','" & fax2 & "','" & fax3 & "','" & email & "','" & zulg & "','" & UCase$(Left$(arzttyp, 2)) & "','" & gemmit & "','" & beme & "'," & Abs(dmpt2) & "," & Abs(dmpt1) & ",'" & Left$(geschlecht, 1) & "','" & Titel & "','kvaerzte.haealt','" & Format(Now(), "yymmddhhmmss") & "');"
         Call QCn(dbknr).Execute(sqlakt, rAF)
         rs.Move 1
        Loop
       End If
      End If
      Set rs = Nothing
      rs.Open "SELECT DISTINCT kvnr FROM `hausaerzte` WHERE nachname = '" & rF0 & "' AND vorname = '" & rF1 & "' AND plz = '" & rEx.Fields(5) & "' AND kvnr <> '" & REPLACE(KVNr, " ", "") & "'", QCn(dbknr), adOpenStatic, adLockReadOnly
      If rs.State <> 0 Then
       If Not rs.BOF Then
        Do While Not rs.EOF
         sqlakt = "INSERT INTO `" & LIUET & "` (name,vorname,titelt,fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,DBNr, BStelle, anrede, tel1, tel2, tel3, tel4, fax1, fax2, fax3, email, zulg, arzttyp, gemmit, beme, dmpt2, dmpt1, geschlecht, titel) VALUES (" & "'" & rF0 & "','" & rF1 & "','" & rEx.Fields(2) & "','" & rEx.Fields(3) & "','" & rEx.Fields(4) & "','" & rEx.Fields(5) & "','" & rEx.Fields(6) & "','" & f7 & "','" & f8 & "'," & IIf(KVNr = "", "null", "'" & REPLACE$(rs!KVNr, " ", "") & "'") & "," & datformZ(Now, True) & ",'" & DBNr & "','" & BStelle & "','" & anrede & "','" & tel1 & "','" & tel2 & "','" & tel3 & "','" & tel4 & "','" & fax1 & "','" & fax2 & "','" & fax3 & "','" & email & "','" & zulg & "','" & UCase$(Left$(arzttyp, 2)) & "','" & gemmit & "','" & beme & "'," & Abs(dmpt2) & "," & Abs(dmpt1) & ",'" & Left$(geschlecht, 1) & "','" & Titel & "','hausaerzte','" & Format(Now(), "yymmddhhmmss") & "');"
         Call QCn(dbknr).Execute(sqlakt, rAF)
         rs.Move 1
        Loop
       End If
      End If
#End If
      On Error GoTo fehler
      DoEvents
      End If ' lvobmysql
      
      If QCn(dbknr).State = 1 Then QCn(dbknr).Close
      If HCn(dbknr).State = 1 Then HCn(dbknr).Close
     Next dbknr
    End If ' If rEx.Fields(0) <> vns AND NOT IsNumeric(rEx.Fields(0)) Then
   End If ' kopfz > 2
   rEx.Move 1
 Loop
 frm.Ausgeb "Fertig mit Einlesen von: '" & aktDatei & "', " & zeile & " Zeilen bearbeitet, aktualisiere jetzt ...", True
 Debug.Print zeile, " Zeilen bearbeitet"
' Exit Sub
#End If
 
For dbknr = LBound(QCn) To UBound(QCn)
 QCn(dbknr).Open
 HCn(dbknr).Open
 LVobMySQL = InStr(UCase(QCn(dbknr).ConnectionString), "MYSQL") > 0  '(Not (cDtb = accDtb))
 LVobMySQL = True ' 11.10.15
 Call Zinit(LVobMySQL)
 Set rV1 = Nothing
' rV1.Open "SELECT name,vorname, MAX(strasse) strasse, MAX(plz) plz, MAX(ort) ort, MAX(telefon) telefon, MAX(fax) fax, MAX(email) email FROM `" & LIUET & "` GROUP BY name, vorname", QCn(dbknr), adOpenDynamic, adLockReadOnly
 myFrag rV1, "SELECT name,vorname, MAX(strasse) strasse, MAX(plz) plz, MAX(ort) ort, MAX(telefon) telefon, MAX(fax) fax, MAX(email) email FROM `" & LIUET & "` GROUP BY name, vorname", adOpenDynamic, QCn(dbknr)
 Do While Not rV1.EOF
  Set rV2 = Nothing
'  If rV1!Name = "Basin" Then Stop
  On Error Resume Next
'  rV2.Open "SELECT * FROM `hausaerzte` WHERE nachname = '" & rV1!name & "' AND vorname = '" & rV1!Vorname & "'", QCn(dbknr), adOpenDynamic, adLockOptimistic
  myFrag rV2, "SELECT * FROM `hausaerzte` WHERE nachname = '" & rV1!name & "' AND vorname = '" & rV1!Vorname & "'", adOpenDynamic, QCn(dbknr)
  On Error GoTo fehler
  If rV2.State <> 0 Then
   If Not rV2.EOF Then
    If Not LVobMySQL Then On Error Resume Next
    Call QCn(dbknr).Execute("UPDATE `" & LIUET & "` SET überschrift = '" & rV2!Überschrift & "' WHERE name = '" & rV1!name & "' AND vorname = '" & rV1!Vorname & "'", rAF)
    If (Not IsNull(rV2!e_mail) And rV2!e_mail <> vNS) And (rV1!email = vNS Or IsNull(rV1!email)) Then
     Call QCn(dbknr).Execute("UPDATE `" & LIUET & "` SET email = '" & rV2!e_mail & "' WHERE name = '" & rV1!name & "' AND vorname = '" & rV1!Vorname & "'", rAF)
    End If
    On Error GoTo fehler
'   Debug.Print "nicht gefunden: ", rV1!Name, rV1!vorname
   End If
   Do While Not rV2.EOF
    If REPLACE$(REPLACE$(LCase(rV2!anschrift), " ", vNS), "str.", "straße") <> REPLACE$(REPLACE$(LCase(rV1!strasse & "," & rV1!plz & rV1!ort), " ", vNS), "str.", "straße") Then
'    Debug.Print rv2!nachname, rv2!vorname, rv2!anschrift, "<>", rV1!strasse & ", " & rV1!plz & " " & rV1!ort
    End If
   
    Dim obT%, v1$, v2$
    obT = False
    If IsNull(rV2!telefon) + (IsNull(rV1!telefon) Or rV1!telefon = vNS) = -1 Then
     obT = True
    ElseIf Not IsNull(rV2!telefon) Then
     v1 = REPLACE(REPLACE(REPLACE(rV1!telefon, " ", vNS), "/", vNS), "-", vNS)
     If Left(v1, 1) <> "0" Then
      If Len(v1) < 7 Then
       v1 = aktVorw & v1
      Else
       v1 = "0" & v1
      End If
     End If
     v2 = REPLACE(REPLACE(REPLACE(rV2!telefon, " ", vNS), "/", vNS), "-", vNS)
     If v1 <> v2 Then obT = True
    End If
    If obT Then
     Debug.Print rV2!Nachname, rV2!Vorname, rV2!telefon, "<>", rV1!telefon
    End If
   
    obT = False
    If IsNull(rV2!telefax) + (IsNull(rV1!fax) Or rV1!fax = vNS) = -1 Then
     obT = True
    ElseIf Not IsNull(rV2!telefax) Then
     v1 = REPLACE(REPLACE(REPLACE(rV1!fax, " ", vNS), "/", vNS), "-", vNS)
     If Left(v1, 1) <> "0" Then
      If Len(v1) < 7 Then
       v1 = aktVorw & v1
      Else
       v1 = "0" & v1
      End If
     End If
     v2 = REPLACE(REPLACE(REPLACE(rV2!telefax, " ", vNS), "/", vNS), "-", vNS)
     If v1 <> v2 Then obT = True
    End If
    If obT Then
     Debug.Print rV2!Nachname, rV2!Vorname, rV2!telefax, "<>", "'" & rV1!fax & "'"
    End If
    DoEvents
    rV2.Move 1
   Loop ' While Not rV2.EOF
  End If ' rV2.State <> 0 Then
  rV1.Move 1
 Loop ' While Not rV1.EOF
 Call QCn(dbknr).Execute("UPDATE IGNORE " & LIUET & " l " & _
 "LEFT JOIN (SELECT * FROM faelle f GROUP BY übwvbsnr, übwr, übwlanr, üwnnr, üwnan, üwtit, üwvor, üwvsw, üwvid) f " & _
 "ON l.name = f.üwnan AND (l.vorname = concat(f.üwvor,IF(ISNULL(f.üwvsw) OR f.üwvsw='','',' '),f.üwvsw) OR l.vorname = f.üwvsw) AND l.titelt = f.üwtit " & _
 "SET l.kvnr= REPLACE(f.übwvkvnr,' ',''),l.ursp=CONCAT(l.ursp,'+faelle.übwvknr') WHERE REPLACE(übwvkvnr,' ','')<>''" & _
 "", rAF) '"WHERE kvnr='' AND REPLACE(übwvkvnr,' ','')<>'' AND NOT ISNULL(übwvkvnr);", rAF)
 sql = "UPDATE " & LIUET & " SET anrede = 'Frau' WHERE vorname IN ('Andrea','Alexandra','Angelika','Anke','Anne','Annett','Annette','Astrid','Barbara','Bibiana','Birgit','Carolin','Christina','Christine','Dagmar','Dorothea','Edith','Elke','Eva','Gabriele','Gerlinde','Gesche','Gisela','Heidi','Heidrun','Ingrid','Isabella','Jana','Julia','Jutta','Katharina','Katrin','Kirsten','Kristina','Laima','Margot','Maria','Marianne','Marion','Mirjana','Monica','Nicole','Nina','Oana','Oskana','Rabia','Rita','Sabine','Silke','Sonja','Susanne','Swantje','Swetlana','Theodora','Tina','Ursula','Ute','Viktoria','Yvonne')"
 QCn(dbknr).Execute sql, rAF
 sql = "UPDATE " & LIUET & " SET anrede = 'Herrn' WHERE vorname IN ('Andreas','Axel','Bernhard','Christian','Christoph','Clemens','Dieter','Edgar','Elmer','Ernst','Ernst-Ulrich','Franz','Franz Egid','Günther','Guido','Hans','Hans-Hermann','Hans-Joachim','Hans-Jürgen','Heribert','Herrmann','Holger','Ioannis','Joachim','Johann','Johann de','Johannes','Karl','Ludwig','Malte','Mario','Meinrad','Nikolaus','Nikolaus von','Olaf','Peter','Rainer','Ramon','Reinhold','Roman','Rudolf','Rüdiger','Stefan','Theodor','Thomas','Ulrich','Volker','Wolf-Dieter','Wolfgang','Yves')"
 QCn(dbknr).Execute sql, rAF
 ' 29.12.15:
 On Error Resume Next
 QCn(dbknr).Execute "DROP TABLE " & LIUEZ, rAF
 On Error GoTo fehler
 QCn(dbknr).Execute "CREATE TABLE " & LIUEZ & " LIKE " & LIUET, rAF
 QCn(dbknr).Execute "INSERT INTO " & LIUEZ & " SELECT * FROM " & LIUET & " l GROUP BY kvnr, name, vorname, plz,ort, telefon, tel1,fax, tel2, email, dmpt2, dmpt1, geschlecht", rAF
 QCn(dbknr).Execute "DROP TABLE " & LIUET, rAF
 QCn(dbknr).Execute "ALTER TABLE " & LIUEZ & " ADD COLUMN kvnri INTEGER(20) AS (CAST(REPLACE(kvnr,' ','') As Integer));", rAF
 QCn(dbknr).Execute "ALTER TABLE " & LIUEZ & " ADD INDEX kvnr(kvnri) USING BTREE;", rAF
 QCn(dbknr).Close
 HCn(dbknr).Close
Next dbknr
End If ' obüberhaupt
'Print #399, Now() & ": Fertig mit Aktualisieren der Hausärzte!"
frm.Ausgeb Now() & ": Fertig mit Aktualisieren der Hausärzte!", True
Else
' Print #399, Now() & ": Keine Datei zum Aktualisieren der Hausärzte!"
 frm.Ausgeb Now() & ": Keine Datei zum Aktualisieren der Hausärzte!", True
End If ' Fileexists
'Close #399
'Shell Environ("windir") & "\system32\notepad.exe " & Environ("tmp") & ausgD, vbNormalFocus
'ProgEnde
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 And lauf < 200 Then
 lauf = lauf + 1
 Sleep 1000
 Resume
End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in main/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Main

Public Function hausaerztekomprimier() ' 6.12.09
 DBCn.Execute "CREATE TABLE quelle.`ha2` LIKE quelle.`hausaerzte`"
 DBCn.Execute "INSERT INTO quelle.`ha2` SELECT * FROM quelle.`hausaerzte` h GROUP BY überschrift, name, vorname, nachname, anschrift, kvnr, telefon, telefax, e_mail, zulassungsgebiet, arzttyp, `gemeinschaftspraxis mit`, schwerpunkt, zusatzbezeichnung, bemerkung, beme, sprechstunden, `von _ bis`, internetadressen, `behandlung IN fremdsprachen`, `rollstuhlgerechte praxis`, verkehrsmittel, linie, `haltestelle parkplätze`, wegbeschreibung, `entfernung zur praxis`, nichtmehr, titel, geschlecht, straße, plz, ort, dmpt2, dmpt1, gelöscht, zahl;"
 DBCn.Execute "set foreign_key_checks = 0;"
 DBCn.Execute "DROP TABLE quelle.`hausaerzte`;"
 DBCn.Execute "set foreign_key_checks = 1;"
 DBCn.Execute "CREATE TABLE quelle.`hausaerzte` LIKE quelle.`ha2`;"
 DBCn.Execute "INSERT INTO quelle.`hausaerzte` SELECT * FROM quelle.`ha2` h GROUP BY überschrift, name, vorname, nachname, anschrift, kvnr, telefon, telefax, e_mail, zulassungsgebiet, arzttyp, `gemeinschaftspraxis mit`, schwerpunkt, zusatzbezeichnung, bemerkung, beme, sprechstunden, `von _ bis`, internetadressen, `behandlung IN fremdsprachen`, `rollstuhlgerechte praxis`, verkehrsmittel, linie, `haltestelle parkplätze`, wegbeschreibung, `entfernung zur praxis`, nichtmehr, titel, geschlecht, straße, plz, ort, dmpt2, dmpt1, gelöscht, zahl;"
End Function ' hausaerztekomprimier()

' 15.5.22 auskommentiert, da hier doppelt
'Public FUNCTION syscmd(art%, Optional Inhalt$)
'End Function

'Sub ShowAllDomains(ccol As Collection, Optional neu%)
'  Dim oNameSpace  As Object
'  Dim oDomain     As Object
'  Static altCCol As Collection
'  ON Error GoTo fehler
'  If Not neu Then
'   If Not altCCol Is Nothing Then
'    Set ccol = altCCol
'    Exit Sub
'   End If
'  End If
'  Set oNameSpace = GetObject("WinNT:")
'  For Each oDomain IN oNameSpace
'    Call ShowAllComputers(oDomain.Name, ccol)
'  Next
'  Set altCCol = ccol
'  Exit Sub
'fehler:
'  ' vermutlich ist kein WMI installiert
' SELECT Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(fpos), vbAbortRetryIGNORE, "Aufgefangener Fehler in ShowAllDomains/" + App.Path)
'  Case vbAbort: Call MsgBox("Höre auf"): End
'  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
'  Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
' End Select
'End Sub
'Sub ShowAllComputers(ByVal strDomain As String, ccol As Collection)
'  Dim PrimDomainContr     As Object
'  Dim oComputer           As Object
'  ON Error GoTo fehler
'  Set PrimDomainContr = GetObject("WinNT://" & strDomain)
'  PrimDomainContr.Filter = Array("Computer")
'  ON Error Resume Next
'
'  For Each oComputer IN PrimDomainContr
'    ccol.Add oComputer.Name ' HostByName(oComputer.Name)
''    Debug.Print oComputer.Name
'  Next
'  ON Error GoTo fehler
'  Exit Sub
'
'fehler:
'  ' vermutlich ist kein WMI installiert
' SELECT Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(fpos), vbAbortRetryIGNORE, "Aufgefangener Fehler in ShowAllComputers/" + App.Path)
'  Case vbAbort: Call MsgBox("Höre auf"): End
'  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
'  Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
' End Select
'End Sub ' ShowAllComputers

' 15.5.22 auskommentiert, da hier doppelt
' Public Sub ProgEnde()
'  End
' End Sub ' Ende

' 15.5.22 auskommentiert, da hier doppelt
' FUNCTION fallzeig()
' End Function

Function testgetalldb()
 Dim acn() As ADODB.Connection
 Set fI = New FürIcon
 Set fI.dbv = New DBVerb
 Call fI.dbv.getAllDB("anamnesebogen", acn(), , , True)
 Stop
End Function

' hier das falsche von oben
#If False Then
nochmal:
      v1sql(0) = "SELECT * FROM `hae` WHERE nachname = '" & rF0 & "' and vorname = '" & rF1 & "' and plz = '" & rEx.Fields(5) & "' ORDER BY aktzeit desc"
      v1sql(1) = "SELECT * FROM `haealt` WHERE nachname = '" & rF0 & "' and vorname = '" & rF1 & "' and plz = '" & rEx.Fields(5) & "' ORDER BY aktzeit desc"
      v1sql(2) = "SELECT * FROM `hae` WHERE nachname = '" & rF0 & "' and vorname = '" & rF1 & "' ORDER BY aktzeit desc"
      v1sql(3) = "SELECT * FROM `haealt` WHERE nachname = '" & rF0 & "' and vorname = '" & rF1 & "' ORDER BY aktzeit desc"
      v1sql(4) = "SELECT * FROM `hae` WHERE nachname = '" & rF0 & "' ORDER BY aktzeit desc"
      v1sql(5) = "SELECT * FROM `haealt` WHERE nachname = '" & rF0 & "' ORDER BY aktzeit desc"
      
      For i = 0 To UBound(v1sql) + 1
       If i = UBound(v1sql) + 1 Then
        Debug.Print rF0
'        Stop
        pos = InStr(rF0, "/")
        If pos = 0 Then pos = InStr(rF0, ",")
        If pos = 0 Then pos = InStr(rF0, "(")
'        If pos = 0 Then pos = InStr(rF0, "-")
        If pos <> 0 Then
         rF0 = Trim$(Left$(rF0, pos - 1))
         posv = InStr(rF1, "/")
         If posv <> 0 Then
          rF1 = Left(rF1, posv - 1)
         End If
         GoTo nochmal:
        End If
        If Not rV1 Is Nothing Then
         If rV1.State <> 0 Then
          rV1.MoveLast
          rV1.Move 1
         End If
        End If
        Exit For
       End If
       Set rV1 = Nothing
       Err.Clear
       On Error Resume Next
       rV1.Open REPLACE(REPLACE(v1sql(i), "SELECT *", "SELECT count(0) as ct"), " ORDER BY aktzeit desc", vNS), HCn(dbknr), adOpenStatic, adLockReadOnly
       If Err.Number = 0 Then
        On Error GoTo fehler
        If rV1!ct >= 1 Then
         Set rV1 = Nothing
         rV1.Open v1sql(i), HCn(dbknr), adOpenStatic, adLockReadOnly
         Exit For
        Else
'        Stop
        End If
       Else
        On Error GoTo fehler
       End If
      Next i
      
      Dim obzu%
      obzu = 0
      If rV1 Is Nothing Then
       obzu = True
      ElseIf rV1.State = 0 Then
       obzu = True
      ElseIf rV1.EOF Then
       obzu = True
      End If
      If Not obzu Then
       KVNr = REPLACE(REPLACE(rV1!KVNr, "/", vNS), " ", "")
       anrede = IIf(IsNull(rV1!anrede), vNS, rV1!anrede)
       tel2 = IIf(IsNull(rV1!tel2), vNS, rV1!tel2)
       fax2 = IIf(IsNull(rV1!fax2), vNS, rV1!fax2)
       On Error Resume Next
       DBNr = rV1!DBNr
       BStelle = IIf(IsNull(rV1!BStelle), vNS, rV1!BStelle)
       tel3 = IIf(IsNull(rV1!tel3), vNS, rV1!tel3)
       tel4 = IIf(IsNull(rV1!tel4), vNS, rV1!tel4)
       fax3 = IIf(IsNull(rV1!fax3), vNS, rV1!fax3)
       gemmit = IIf(IsNull(rV1!gemmit), vNS, rV1!gemmit)
       beme = IIf(IsNull(rV1!beme), vNS, rV1!beme)
       geschlecht = IIf(IsNull(rV1!geschlecht), vNS, rV1!geschlecht)
       email = IIf(IsNull(rV1!email), vNS, rV1!email)
       On Error GoTo fehler
       zulg = IIf(IsNull(rV1!zulg), vNS, rV1!zulg)
       arzttyp = IIf(IsNull(rV1!arzttyp), vNS, rV1!arzttyp)
       dmpt2 = IIf(IsNull(rV1!dmpt2), 0, rV1!dmpt2)
       dmpt1 = IIf(IsNull(rV1!dmpt1), 0, rV1!dmpt1)
       Titel = IIf(IsNull(rV1!Titel), vNS, rV1!Titel)
      Else
       KVNr = vNS
       DBNr = vNS
       BStelle = vNS
       anrede = vNS
       tel2 = vNS
       tel3 = vNS
       tel4 = vNS
       fax2 = vNS
       fax3 = vNS
       email = vNS
       zulg = vNS
       arzttyp = vNS
       gemmit = vNS
       beme = vNS
       dmpt2 = 0
       dmpt1 = 0
       geschlecht = vNS
       Titel = vNS
      End If
#End If


