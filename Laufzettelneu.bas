Attribute VB_Name = "Laufzettel"
Option Explicit
' alle Variablen, die (nur) für tabelleinplz und plz gebraucht werden
#If mitcovid Then
 Const pzl% = 54     ' Zahl der Parameter
#Else
 Const pzl% = 53 ' Zahl der Parameter
#End If
' überprüft notwendig
 Dim MerkblattText$
 Dim sollz%
 Dim obnach% ' Nachschulungen festlegen
 Dim FLI#, FIB4#, NFS#
 Dim rund%(pzl) ' runden auf; default -1
 Dim rszmax& ' maximal dargestellte Zahl an Ausprägungen jedes Parameters=zmax%(pzl), Maximum davon je Großzeile
 Dim rs(pzl) As New ADODB.Recordset
 Dim sql$(pzl) ' werden hinter "' Recordsets öffnen" verwendet
 Dim kritnr(pzl) As LabArt ' Kriterium Nr. für die Filternr
 Dim rsz&(pzl) ' vorhandene Zahl
 Dim weite$(pzl) ' Angabe zur Weite
 Dim RowSp%(pzl) ' Zahl der Zeilen pro Kleinzeile
 Dim ÜS$(pzl), Titel$(pzl)  ' Überschrift
 Dim Fqmin!(pzl) ' Mindest-Untersuchungs-Zahl pro Jahr
 Dim Fqminu% 'Duplikat für die Füße
 Const GrenzeDiast! = 81
 Dim nurpath%(pzl) ' -1 = nicht anzeigen
                  ' 1 = Anzeige nur bei Fehlen eines Wertes und bei pathologischem Wert,
                  ' 2 = Anzeige nur bei pathologischem Wert
                  ' 3 = Anzeige bei Vorhandensein eines Wertes
 Dim pathol%() ' (pzl, zmaxges)
 Const maxzmaxges& = 100 ' nur in dodoplz
 Dim TI!(41) ' Timer
 Dim m%, p% ' für Timer
 Dim RRi% ' Parameter für Blutdruck
 Dim Neuri% ' Parameter für Neurostatus
 Dim Fußi% ' Parameter für Fußinspektion
 Dim LDLi% ' Parameter für LDL
 Dim Gwi% ' Paraemter für Gewicht
 Dim Schuli% ' Parameter für Schulungen
 Dim SchulEintri% ' Parameter für Schulungseinträge
 Dim obdm%, obweibl%, obdement%
 Dim therart$ ' letzte Therapieart
 Dim obreti% ' Retinopathie
 Dim aRisk As New Risk ' für UKPDS Risk engine, erfordert u:\programmierung\RiskEng.ocx
 Dim PAlter#, NName$ ' Patientenalter, Name
 Dim grenze!(pzl) ' Grenze für pathologische Werte
 Dim obFußmakro%
 Dim üdt As DMPClass
 Dim rsMB As New ADODB.Recordset
 Dim obMB% ' ob Merkblatt Fußsyndrom mitgegeben
 
'Public plzVerz$
'Const plzVerz$ = plzVz
Private Declare Function ShellExecuteEx& Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal lpOperation As String, _
  ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long
Declare Function WaitForSingleObject& Lib "kernel32.dll" (ByVal hHandle&, ByVal dwMilliseconds&)
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long ' ob saArray initialisiert

Private Const SYNCHRONIZE = &H100000
'Private Const WAIT_TIMEOUT AS Long = &H102&
 
Private Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hwnd As Long
  lpVerb As String
  lpFile As String
  lpParameters As String
  lpDirectory As String
  nShow As Long
  hInstApp As Long
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type
 
Type mptyp
 m As medplan
 metf As Integer
 sglt As Integer
End Type
Dim mdpl() As mptyp
 
' SHELLEXECUTEINFO fMask-Konstanten
Private Const SEE_MASK_CLASSKEY = &H3 ' Struktur wird mit dem Handle des Registry-
' Schlüssels für die Klasse gefüllt
Private Const SEE_MASK_CLASSNAME = &H1 ' Struktur wird mit dem Klassennamen oder GUID
' gefüllt der die Datei beschreibt
Private Const SEE_MASK_CONNECTNETDRV = &H80 ' Struktur verbindet den PC mit einem
' Netzlaufwerk falls nötig, lpFile muss der UNC-Pfadname im Netzwerk sein
Private Const SEE_MASK_DOENVSUBST = &H200 ' Struktur wird mit Umgebungsvariablen des
' Programms gefüllt, diese werden in lpDirectory oder lpFile gepackt
Private Const SEE_MASK_FLAG_DDEWAIT = &H100 ' die Funktion wartet, dass die DDE ihre
' Vorgänge abgeschlossen hat und kehrt dann erst zurück
Private Const SEE_MASK_FLAG_NO_UI = &H400 ' die Funktion zeigt keine Fehler-Dialogboxen an
Private Const SEE_MASK_HOTKEY = &H20 ' dwHotkey wird gefüllt
Private Const SEE_MASK_ICON = &H10 ' hIcon wird mit dem Icon Handle des
' Standardicons der Anwendung gefüllt
Private Const SEE_MASK_IDLIST = &H4 ' benutzt die lpIDLIST-Option, um das Programm zu
' starten, das in ihr beschrieben ist
Private Const SEE_MASK_INVOKEIDLIST = &HC ' benutzt die lpIDLIST-Option, um das
' Programm zu starten, das in ihr beschrieben ist
Private Const SEE_MASK_NOCLOSEPROCESS = &H40 ' füllt die Struktur-Option hProcess
' mit dem Process-Handle der gestarteten Anwendung
 
' SHELLEXECUTEINFO nShow-Konstanten
Private Const SW_HIDE = 0 ' versteckt das Fenster
Private Const SW_MAXIMIZE = 3 ' maximiert das Fenster
Private Const SW_MINIMIZE = 6 ' minimiert das Fenster
Private Const SW_RESTORE = 9 ' stellt das Fenster wieder her
Private Const SW_SHOW = 5 ' zeigt das Fenster an
Private Const SW_SHOWMAXIMIZED = 3 ' zeigt das Fenster maximiert an
Private Const SW_SHOWMINIMIZED = 2 ' zeigt das Fenster minimiert an
Private Const SW_SHOWMINNOACTIVE = 7 ' zeigt das Fenster minimiert an aber aktiviert
' es nicht
Private Const SW_SHOWNA = 8 ' zeigt das Fenster an, aber aktiviert es nicht
Private Const SW_SHOWNOACTIVATE = 4 ' zeigt das Fenster in der besten Größe und
' Position an aber aktiviert es nicht
Private Const SW_SHOWNORMAL = 1 ' zeigt das Fenster ganz normal an
 
' SHELLEXECUTEINFO hInstApp Rückgabe-Konstanten
Private Const SE_ERR_ACCESSDENIED = 5 ' Zugriff verweigert
Private Const SE_ERR_ASSOCINCOMPLETE = 27 ' Dateityp ist nicht ausreichend assoziiert
Private Const SE_ERR_DDEBUSY = 30 ' DDE konnte nicht gestartet werden
Private Const SE_ERR_DDEFAIL = 29 ' DDE ist gescheitert
Private Const SE_ERR_DDETIMEOUT = 28 ' DDE-Zeitlimit wurde ereicht
Private Const SE_ERR_DLLNOTFOUND = 32 ' eine benötigte Dll wurde nicht gefunden
Private Const SE_ERR_FNF = 2 ' Datei wurde nicht gefunden
Private Const SE_ERR_NOASSOC = 31 ' Dateityp ist nicht assoziiert
Private Const SE_ERR_OOM = 8 ' nicht genügend Speicher verfügbar
Private Const SE_ERR_PNF = 3 ' Pfad wurde nicht gefunden
Private Const SE_ERR_SHARE = 26 ' Datei konnte nicht geöffnet werden da sie bereits
' verwendet wird
 
' SHELLEXECUTEINFO dwHotKey-Konstanten
Private Const HOTKEYF_ALT = &H4 ' benutzt ALT für den Hotkey
Private Const HOTKEYF_CONTROL = &H2 ' benutzt STRG für den Hotkey
Private Const HOTKEYF_EXT = &H8 ' benutzt den Extendend-Key für den Hotkey
Private Const HOTKEYF_SHIFT = &H1 ' benutzt Shift für den Hotkey
 
' WaitForSingleObject dwMillisekond-Konstante
Private Const INFINITE = &HFFFF ' unendlich warten
 
' WaitForSingleObject Rückgabe-Konstanten
Private Const WAIT_ABANDONED = &H80 ' der Mutex der in hHanlde angegeben ist wird
' nicht freigegeben, so lange der Eltern-Thread zerstört ist. Der Mutex ist nun
' Bestandteil des aufrufenden Threads und ist nicht mehr im signalisierenden Status
Private Const WAIT_FAILED = &HFFFFFFFF ' die Funktion ist gescheitert
Private Const WAIT_OBJECT_0 = &H0 ' das Objekt, das in hHandle spezifiziert ist,
' ist in einem signalisierendem Status
Private Const WAIT_TIMEOUT = &H102 ' das Zeitlimit für eine Änderung des
' Thread-Status ist abgelaufen
Const phpVS$ = LiServer & "php\"
Const phpV$ = phpVS & "plz\"
Const phpVvorb$ = phpVS & "vorb\"
Const phpVfertig$ = phpVS & "fertig\"
Const phpVbehand$ = phpVS & "behand\"

Type plzDaten ' für doPatientenlaufzettel
 pid As Long
 Datum As Date
 Uhrzeit As Date
 Arzt As String
 zusatz As String
End Type ' plzDaten

Dim obHCT%, hctMed$

' #If False Then
' Editor mit der Autoexec.bat Datei öffnen
' kommentiert aufgerufen in doPatientenlaufzettel
Public Sub doplink(libefehl$, Optional passw$)
  Dim RetVal As Long, ShExInfo As SHELLEXECUTEINFO
  Static pwd$
  If passw <> "" Then pwd = passw
  If pwd = "" Then pwd = InputBox("bitte Passwort für root auf Linux eingeben", "Passworteingabe")
  ' Startoptionen festlegen
  With ShExInfo
    .cbSize = Len(ShExInfo)
    .fMask = SEE_MASK_FLAG_NO_UI Or SEE_MASK_CLASSNAME Or SEE_MASK_NOCLOSEPROCESS
    .hwnd = 0
    .lpVerb = "open"
    .lpFile = """C:\Program Files\PuTTY\plink.exe"""
    .lpParameters = "-pw " & pwd & " -batch root@" & LiName & " """ & libefehl & """" ' mariadb-dump --defaults-extra-file=~/.modbpwd medoff epraxis |mysql --defaults-extra-file=~/.mariadbpwd quelle"
'    .lpDirectory = "c:\program files\Putty"
'    .nShow = SW_SHOWMAXIMIZED
    .nShow = SW_SHOWNORMAL
  End With
 
  ' Programm ausführen
  RetVal = ShellExecuteEx(ShExInfo)
'  Retval = ShellExecute(0, "open", "plink", "root@" & liname & " /usr/bin/termine", "c:\", SW_SHOWNORMAL)
  If RetVal = 0 Then
 
    ' bei Fehler Text ausgeben
    Select Case ShExInfo.hInstApp
    Case SE_ERR_ACCESSDENIED
      Err.Raise 5, "ShellExecuteEx", "Zugriff verweigert"
    Case SE_ERR_FNF
      Err.Raise 5, "ShellExecuteEx", "Datei nicht gefunden"
    Case SE_ERR_NOASSOC
      Err.Raise 5, "ShellExecuteEx", "Datei ist mit keinem Programm verknüpft"
    End Select
  Else
    ' andernfalls Programmdaten ausgeben
'    Debug.Print "Datei gestartet: " & ShExInfo.lpFile
'    Debug.Print "Dateiklasse: " & ShExInfo.lpClass
'    Debug.Print "Prozess Handle: " & ShExInfo.hProcess
'    Debug.Print "Instanz Handle:" & ShExInfo.hInstApp
  End If
  ' warten, bis die Anwendung beendet wird
  Do
    DoEvents
  Loop Until WaitForSingleObject(ShExInfo.hProcess, 0) <> WAIT_TIMEOUT
 
  ' beenden des Programms signalisieren
  MsgBox "Das Programm wurde beendet"
End Sub ' doplink()
' #End If

'#If False Then
' kommentiert aufgerufen in doPatientenlaufzettel und rufThFestleg
Public Sub RunCommandLine(sCmdLine$)
    Dim nProcessID&
    Dim hProcess&
    Dim nResult&
    nProcessID = Shell(sCmdLine, vbNormalNoFocus)
    If nProcessID <> 0 Then
        hProcess = OpenProcess(SYNCHRONIZE, 0, nProcessID)
        If hProcess <> 0 Then
            Do
                DoEvents
                nResult = WaitForSingleObject(hProcess, 50)
            Loop Until nResult <> WAIT_TIMEOUT 'Or m_bStop
            CloseHandle hProcess
        End If
    End If
End Sub ' RunCommandLine
'#End If

Function doPatientenlaufzettel(Optional obohnerueckfrage% = 0, Optional obphp% = 0) ' aufgerufen aus: Patientenlaufzettel_Click() und MDIForm_Activate
    ' für PDF lesen
'#Const obadobe = 1
Dim Cn As ADODB.Connection
#If obadobe = 1 Then
 Dim PDTextS As Acrobat.CAcroPDTextSELECT
 Dim Result&, PDDoc As Acrobat.CAcroPDDoc
 Dim PDPage As Acrobat.CAcroPDPage
 Dim PDHili As Acrobat.CAcroHiliteList
 Dim b As Boolean
 Dim PDFStream$
#End If
 Dim Filename$, FTextname$, erg$
' Strings Falls Text von Anfang ausgelesen werden soll (bzw bis zum Ende)
 Dim FromFirst$, ToLast$
 Dim Str As New CString, NTL&
 Dim i%, j&, splitt$(), s2$(), pid$, Datum As Date
 Dim Uhrzeit$, Arzt$, erstZ&
 Dim obausdb%
 Dim plzVerz$

' IF pVerz = vNS THEN pVerz = "p:\"
' IF plzVerz = vNS THEN plzVerz = plzVz
' alte löschen
 
' Prüfen, ob Termine gespeichert wurden und mit dem C-Programm "termine" (auf <LiName>) ausgelesen wurden
 plzVerz = phpV
 If Lese.MOBetr <> 0 Then
  obausdb = True
 Else
  If obphp <> 0 Then plzVerz = phpV Else plzVerz = plzVz
  Filename = plzVz & "TMFTools.pdf"
  FTextname = plzVz & "TMFTools.pdf.txt"
' IF Dir(plzVerz) <> vNS AND Dir(FTextname) <> vNS THEN
  If FSO.FolderExists(plzVz) And FSO.FileExists(FTextname) Then
    If FileDateTime(Filename) <= FileDateTime(FTextname) Then
    obausdb = True
   End If
  End If
  If Not obausdb And FSO.FileExists(Filename) Then
'   doplink
' 22.10.22: führt bei Aufruf über Ado zumindest bis zur Mariadb-Version 10.9 immer wieder zum Server-Crash, s.ähnliche Bug-Hinweise früherer Versionen
   rufauf "ssh", "root@" & LiName & " termine", 2, "c:\windows\system32\openssh\", -1, 1
'  RunCommandLine ("plink root@" & LiName & " termine")
  End If ' Not obausdb And FSO.FileExists(Filename) Then
 End If ' Lese.MoBetr
 
' verschiedene Daten bestimmen
 Dim rTerm As New ADODB.Recordset
 Dim jetzt As Date
 Dim InetD As Date
 Dim ServD As Date
 Dim PCDate As Date
 InetD = InetZeit()
 myFrag rTerm, "SELECT now()"
 ServD = rTerm.Fields(0)
 Set rTerm = Nothing
 PCDate = Now()
 If IsDate(InetD) And InetD Then
  jetzt = InetD
 ElseIf IsDate(ServD) And ServD Then
  jetzt = ServD
 ElseIf IsDate(PCDate) And PCDate Then
  jetzt = PCDate
 End If
 
' Bestimmung des Termindatums für die Plz-Erstellung: s2(0)
 ReDim s2(0) ' verwendetes Datum
 If obausdb Then
  Dim dats$
  If obohnerueckfrage Then
   s2(0) = jetzt
  Else ' obohnerueckfrage THEN
   Dim ÜSchr$
   If Lese.MOBetr = 0 Then ÜSchr = "Terminkalender gespeichert um: " & FileDateTime(Filename) & vbCrLf Else ÜSchr = ""
   ÜSchr = ÜSchr & "aktuelle Internetzeit: " & InetD & vbCrLf & "aktuelle Serverzeit:   " & ServD & vbCrLf & "aktuelle PC-Zeit:       " & Date & " " & time & vbCrLf & vbCrLf & "bitte Patientenlaufzettel erstellen für: "
   dats = InputBox(ÜSchr, "Datumsfestlegung zur Patientenlaufzettelerstellung", Int(InetD))
   If Not IsDate(dats) Then Exit Function
   s2(0) = CDate(dats)
  End If ' obohnerueckfrage THEN
  If Not IsDate(s2(0)) Then Exit Function
 
 Else ' obausdb
#If obadobe = 1 Then
' Beginn mit Einlesen der PDF-Dati, u.a. um Termindatum für die PLZ-Erstellung zu bestimmen
  FromFirst = Chr(169) & Chr(170) & Chr(172)
  ToLast = Chr(163) & Chr(165) & Chr(164)
  Set PDDoc = CreateObject("AcroExch.pdDoc")
  Result = PDDoc.Open(Filename)
  If Not Result Then
   MsgBox "Fehler beim Öffnen von: " & Filename, vbCritical, "Patientenlaufzettel:"
   Exit Function
  End If ' not Result
' Nehme die Erste Seite - Index 0
  Do
   Set PDPage = PDDoc.AcquirePage(j)
   ' Erzeuge ein Highlight Objekt und weise ihm 2000 Elemente bei (keine Grenzprobleme)
   Set PDHili = CreateObject("AcroExch.HiliteList")
   b = PDHili.Add(0, 32000)
   ' Erzeuge eine Textauswahl aus dem gesamten Text
   On Error Resume Next
   Set PDTextS = PDPage.CreatePageHilite(PDHili)
   If Err.Number <> 0 Then Exit Do
   On Error GoTo fehler
   ' Hole Anzahl der "Textblöcke"
   NTL = PDTextS.GetNumText
   ' Gebe den Text der Textauswahl zurück
   For i = 0 To NTL - 1
    Str.Append PDTextS.GetText(i)
   Next i
   DoEvents
   Str.Append vbCrLf
   j = j + 1
  Loop
  Result = PDDoc.Close
  Set PDPage = Nothing
  Set PDHili = Nothing
  Set PDTextS = Nothing
  Set PDDoc = Nothing
  SplitNeu Str.Value, vbCrLf, splitt
  SplitNeu splitt(0), " ", s2 ' s0 = "Alle", s1 = "Terminarten", s2 = "1.1.13 - 31.3.2016"
  syscmd 4, "Erstelle Patientenlaufzettel für " & s2(2) & " in " & plzVerz & "..."
  DoEvents
  For j = 0 To UBound(s2)
   If InStrB(s2(j), "-") > 0 Then
    s2 = Split(s2(j), "-")
    Datum = CDate(s2(0)) ' s2(0) = "31.3.2016"
    Exit For
   End If
  Next j
#Else
 ReDim splitt(0)
#End If
 End If ' obausdb
 
' Löschen der Dateien, die nicht für heute oder das Terminerstellungsdatum sind
If s2(0) = "" Then s2(0) = 0
If obphp Then
 Dim Fil As File
 Dim loestr$(3), aktzls
 loestr(0) = phpV
 loestr(1) = phpVvorb
 loestr(2) = phpVfertig
 loestr(3) = phpVbehand
 For Each aktzls In loestr()
  For Each Fil In FSO.GetFolder(aktzls).Files
   If (InStrB(Fil, ",Pid_") <> 0 Or InStrB(Fil, "Patientenlaufzettel") <> 0) And Int(Fil.DateLastModified) <> Int(CDate(s2(0))) And Int(Fil.DateLastModified) <> Int(jetzt) Then
    On Error Resume Next
    FSO.DeleteFile (Fil)
    On Error GoTo fehler
   End If
  Next Fil
 Next aktzls
Else ' obphp Then
 erg = Dir$(plzVerz & "*")
 Do While LenB(erg) <> 0
  If (InStrB(erg, ",Pid_") <> 0 Or InStrB(erg, "Patientenlaufzettel") <> 0) And Int(FileDateTime(plzVerz & erg)) <> Int(CDate(s2(0))) And Int(FileDateTime(plzVerz & erg)) <> Int(jetzt) Then
   Kill plzVerz & erg
  End If
  erg = Dir
 Loop
End If ' obphp Then else
DoEvents
 
 If obausdb Then
'  myFrag rTerm, "SELECT PID, DATE(zp) Datum, time(zp) Uhrzeit, raum Arzt,Zusatz FROM termine t WHERE DATE(zp) = " & Format(s2(0), "yyyymmdd")
  Dim plzDat() As plzDaten
  Dim hinzu&, sql$, iru&, noch&, unoch&
  For hinzu = 0 To 45
   Datum = CDate(s2(0)) + hinzu
   If Lese.MOBetr = 0 Then
ohnemo:
    Set Cn = DBCn
    sql = "SELECT PID, DATE(zp) Datum, MIN(TIME(zp)) Uhrzeit, GROUP_CONCAT(DISTINCT raum SEPARATOR ' ') Arzt,GROUP_CONCAT(DISTINCT Zusatz SEPARATOR ' ') Zusatz FROM termine t WHERE " & SelDatum("zp", Datum) & " AND pid<>0 GROUP BY pid"
   Else ' Lese.MOBetr = 0 Then
    If MOConInit(, "Erstelle Patientenlaufzettel") Then GoTo ohnemo
    Set Cn = MOCon
    sql = "SELECT FPatnr pid, DATE(18900101 + INTERVAL t.FDatumvon DAY + INTERVAL t.FZeitvon SECOND)datum, CONVERT(SEC_TO_TIME(t.fzeitvon),CHAR)uhrzeit, t.fbemerkung zusatz, tz.fname arzt FROM termin t LEFT JOIN tzone tz ON t.FZonenid=tz.FSurogat WHERE DATE(18900101 + INTERVAL t.FDatumvon DAY + INTERVAL t.FZeitvon SECOND)=DATE(" & Format(Datum, "yyyymmdd") & ") AND fpatnr>0 GROUP BY fpatnr ORDER BY fpatnr"
   End If ' Lese.MOBetr = 0 Then Else
   myFrag rTerm, "SELECT COUNT(0)zl FROM (" & sql & ") i", adOpenStatic, Cn
   If Not rTerm.BOF Then
    noch = rTerm!zl
    unoch = noch
    If noch <> 0 Then
     syscmd 4, "zu erstellen: " & noch & " Patientenlaufzettel für " & s2(0) & " in " & plzVerz
     ReDim plzDat(noch - 1)
     myFrag rTerm, sql, adOpenStatic, Cn
     iru = 0
     Do While Not rTerm.EOF
      plzDat(iru).pid = rTerm!pid
      plzDat(iru).Datum = rTerm!Datum
      plzDat(iru).Uhrzeit = rTerm!Uhrzeit
      plzDat(iru).Arzt = rTerm!Arzt
      plzDat(iru).zusatz = rTerm!zusatz
      iru = iru + 1
      rTerm.MoveNext
     Loop
     For iru = 0 To UBound(plzDat)
      Call dodoplz(CStr(plzDat(iru).pid), plzVerz, plzDat(iru).Datum, CStr(plzDat(iru).Uhrzeit), , plzDat(iru).Arzt, , obphp)
      erstZ = erstZ + 1
      noch = noch - 1
      syscmd 4, "noch zu erstellen: " & noch & " Patientenlaufzettel für " & s2(0) & " in " & plzVerz
     Next iru
    End If ' noch <> 0 Then
   End If ' Not rTerm.EOF Then
   If noch < 1 And unoch > 5 Then Exit For ' wenn es ein Tag mit mehr als 5 Terminen war
'   If Not rTerm.BOF Then
'    Do While Not rTerm.EOF
'     DoEvents
'     PID = rTerm!PID
'     Datum = rTerm!Datum
'     Uhrzeit = rTerm!Uhrzeit
'     Arzt = rTerm!Arzt
'     Call dodoPLZ(CLng(PID), plzVerz, Datum, Uhrzeit, , Arzt, , obphp)
'     erstZ = erstZ + 1
'     rTerm.MoveNext
'    Loop
'    Exit For
'   End If
'   Set rTerm = Nothing
  Next hinzu
 Else ' obausdb
  For j = 0 To UBound(splitt)
   If InStrB(splitt(j), " ") <> 0 Then
    i = InStr(splitt(j), " ")
    pid = left$(splitt(j), i - 1)
    If IsNumeric(pid) Then
     If InStrB(pid, ".") = 0 And InStrB(pid, ",") = 0 And InStrB(pid, ";") = 0 Then
      Uhrzeit = vNS
      i = InStr(splitt(j), ":")
      If i <> 0 Then
       Arzt = Trim$(Mid$(splitt(j), i + 3))
       Uhrzeit = Mid$(splitt(j), i - 2, 5)
      End If ' i <> 0 THEN
      DoEvents
      Call dodoplz(CLng(pid), plzVerz, Datum, Uhrzeit, , Arzt, , obphp)
      erstZ = erstZ + 1
     End If ' InStrB(Pid, ".") = 0 AND ...
    End If ' IsNumeric(Pid) THEN
   End If ' InStrB(splitt(j), " ") <> 0 THEN
  Next j
' MsgBox "Fertig!"
 End If ' obausdb
 On Error Resume Next
 syscmd 4, erstZ & " Patientenlaufzettel für " & s2(0) & " in " & plzVerz & " erstellt."
 If Not obausdb Then
  syscmd 4, Filename + ": " + CStr(FileDateTime(Filename)) + ", " + FTextname + ": " + CStr(FileDateTime(FTextname)) + " => Programm 'termine' auf " & LiName & " starten!"
 End If
 If obphp <> 0 Then plzVerz = pVerz & "plz\"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doPatientenlaufzettel/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'doPatientenlaufzettel

Function getDiast$(Wert$)
 Dim psl&, psp&
 psl = InStr(Wert, "/")
 If psl <> 0 Then
  psp = InStr(psl + 2, Wert, " ") - 1
  If psp = -1 Then psp = Len(Wert)
 End If
 getDiast = Trim$(Mid$(Wert, psl + 1, psp - psl))
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getDiast/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getDiast
'Function testha1()
' Dim üwerg$()
' Call Lese.ProgStart
' Call getHausarzt(706, üwerg)
' Stop
'End Function

'Public Sub SpeichereStringAlsUTF8Textdatei(DerText AS String, Dateiname AS _
  String)
  
Private Function StringB(Number As Long, Character As String)
  Dim CharInt As String
 
  CharInt = MidB$(Character, 1, 1)
  StringB = String(Int(Number / 2), CharInt & CharInt)
  If Number Mod 2 = 1 Then StringB = StringB & CharInt
End Function ' StringB
  
Public Function UTF16NachUTF8B(Text As String) As String
  Dim temp1 As Long
  Dim Ausgabestr As String
  Dim TempBytes(2) As Byte
  Dim TempBytesUnicode(1) As Byte
  Dim TempArrayB() As Byte
  Dim VorhLaenge As Long
  Dim AktLaenge As Long
  Dim TempLngTextOrg As Long
 
  TempArrayB = Text
  TempLngTextOrg = Len(Text)
  Ausgabestr = StringB(TempLngTextOrg, ChrB(0))
  VorhLaenge = TempLngTextOrg
  For temp1 = 0 To UBound(TempArrayB) Step 2
    TempBytesUnicode(0) = TempArrayB(temp1)
    TempBytesUnicode(1) = TempArrayB(temp1 + 1)
    ' 1-Byte-Zeichn
    If CLng(TempBytesUnicode(1)) * 256& + CLng(TempBytesUnicode(0)) <= &H7F& Then
      If AktLaenge + 1 > VorhLaenge Then
        Ausgabestr = Ausgabestr & StringB(4 + Int(TempLngTextOrg / 4), ChrB(0))
        VorhLaenge = VorhLaenge + 4 + Int(TempLngTextOrg / 4)
      End If
      MidB(Ausgabestr, AktLaenge + 1) = ChrB(TempBytesUnicode(0))
      AktLaenge = AktLaenge + 1
 
    ' 2-Byte-Zeichn
    ElseIf CLng(TempBytesUnicode(1)) * 256& + CLng(TempBytesUnicode(0)) > &H7F& And _
      CLng(TempBytesUnicode(1)) * 256& + CLng(TempBytesUnicode(0)) <= &H7FF& Then
 
      If AktLaenge + 2 > VorhLaenge Then
        Ausgabestr = Ausgabestr & StringB(4 + Int(TempLngTextOrg / 4), ChrB(0))
        VorhLaenge = VorhLaenge + 4 + Int(TempLngTextOrg / 4)
      End If
      TempBytes(0) = &HC0 + TempBytesUnicode(1) * &H4 + Int(TempBytesUnicode(0) / 2 ^ 6)
      TempBytes(1) = &H80 + (TempBytesUnicode(0) And &H3F)
      MidB(Ausgabestr, AktLaenge + 1) = ChrB(TempBytes(0))
      MidB(Ausgabestr, AktLaenge + 2) = ChrB(TempBytes(1))
      AktLaenge = AktLaenge + 2
 
    ' 3-Byte-Zeichn
    ElseIf CLng(TempBytesUnicode(1)) * 256& + CLng(TempBytesUnicode(0)) > &H7FF& And _
      CLng(TempBytesUnicode(1)) * 256& + CLng(TempBytesUnicode(0)) <= &HFFFF& Then
 
      If AktLaenge + 3 > VorhLaenge Then
        Ausgabestr = Ausgabestr & StringB(4 + Int(TempLngTextOrg / 4), ChrB(0))
        VorhLaenge = VorhLaenge + 4 + Int(TempLngTextOrg / 4)
      End If
      TempBytes(0) = &HE0 + Int(TempBytesUnicode(1) / 2 ^ 4)
      TempBytes(1) = &H80 + (TempBytesUnicode(1) And &HF) * 2 ^ 2 + _
        Int(TempBytesUnicode(0) / 2 ^ 6)
      TempBytes(2) = &H80 + (TempBytesUnicode(0) And &H3F)
      MidB(Ausgabestr, AktLaenge + 1) = ChrB(TempBytes(0))
      MidB(Ausgabestr, AktLaenge + 2) = ChrB(TempBytes(1))
      MidB(Ausgabestr, AktLaenge + 3) = ChrB(TempBytes(2))
      AktLaenge = AktLaenge + 3
    End If
  Next
  UTF16NachUTF8B = MidB(Ausgabestr, 1, AktLaenge)
End Function ' UTF16NachUTF8B
  
    ' Speichern des Strings als UTF-8 in Textdatei
 Public Sub sUTF8(DerText$, DateiName$)
    Dim strBytes() As Byte
    Dim DatNr&
    DatNr = FreeFile
    Open DateiName For Binary As DatNr
    ' UTF-8 BOM speichern
    strBytes = ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF)
    Put DatNr, , strBytes
    ' String speichern
    strBytes = UTF16NachUTF8B(DerText & vbCrLf)
    Put DatNr, , strBytes
'    Close DatNr
End Sub ' Sub sUTF8(DerText$, DateiName$)

Public Sub sUTF(DatNr&, DerText$)
    Dim strBytes() As Byte
    strBytes = UTF16NachUTF8B(DerText & vbCrLf)
    Put DatNr, , strBytes
End Sub ' sUTF(DatNr&, DerText$)

Public Function UTFOpen&(DN$)
 UTFOpen = FreeFile
 If Dir(DN) <> "" Then Kill DN
 Open DN For Binary As UTFOpen
End Function '  UTFOpen&(DN$)

' letzte eGFR
Public Function letztGFR(pid&, Optional PAlter#, Optional obweibl%, Optional EthnicGroup% = 1) As labtyp
' letztGFR = hollwert(PID, "GFR','GFRCYS','GFRK','GFRM','GFRT','GFRW','CREACL','GFC','GFCK1','GFCK2','GFCK3','GFCK4','GFCM1','GFCW1','GFCW2','MDRD", "ml/m")
 Static altPID&, altPalter#, altobweibl%
 Dim LA As labtyp, lc As labtyp
 LA = LabPat(LA_eGFR, pid)
 lc = LabPat(LA_Krea, pid)
 If lc.Zp - LA.Zp > 1 Or (lc.WertSg <> "" And LA.WertSg = "") Then ' wenn letztes Kreatinin mehr als einen Tag jünger als die letzte eGFR ist
  If pid <> altPID And PAlter = 0 Then
   altPID = pid
   altPalter = myEFrag("SELECT patalter(" & pid & ")").Fields(0)
   altobweibl = myEFrag("SELECT IF(geschlecht='w',-1,0) FROM namen WHERE pat_id=" & CStr(pid)).Fields(0)
  ElseIf pid <> altPID Then
   altPID = pid
   altPalter = PAlter
   altobweibl = obweibl
  End If
  
  ' MDRD-Formel ohne Hautfarbe (bei schwarz müßte noch mit 1,212 multipliziert werden
  If lc.WertSg <> "" Then letztGFR.WertSg = 175 * (lc.WertSg ^ -1.154) * altPalter ^ -0.203 * IIf(altobweibl, 0.742, 1) * IIf(EthnicGroup = 2, 1.212, 1)
  letztGFR.Abkü = "MDRD"
  letztGFR.Einheit = "ml/min"
  letztGFR.Zp = lc.Zp
 Else
  letztGFR = LA
 End If
End Function ' letztGFR#(PID&)

' letzten Laborwert einer bestimmten Sorte holen
'  Public FUNCTION hollwert#(Pat_id&, Abk$, Einh$)
'   hollwert = myEFrag("call ltWert(" & Pat_id & ",""" & Abk & """,""" & Einh & """)")!ltWert
'  END Function
Public Function mplan(pid&)
' bei Auswahl aller Medikamentenpläne (Pat_id 1564: 4165 Einträge) Dauer: 0,64s
 Dim rTh As New ADODB.Recordset, mpz%, ru%
 Static altPID&
 On Error GoTo fehler
 If pid <> altPID Then
  obHCT = 0
  hctMed = ""
'  mpz = myEFrag("SELECT COUNT(0) Zahl FROM medplan mp WHERE pat_id = " & PID & "")!Zahl
  mpz = myEFrag("SELECT COUNT(0)Zahl FROM wmedplan mp WHERE pat_id = " & pid & " AND zeitpunkt=(SELECT MAX(zeitpunkt) FROM wmedplan WHERE pat_id=" & pid & ")")!Zahl
' mpz = myEFrag("SELECT COUNT(0) Zahl FROM medplan mp WHERE pat_id = " & Pid & " AND mpnr=(SELECT MAX(mpnr) FROM medplan mpi WHERE pat_id=mp.pat_id AND zeitpunkt=(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=mp.pat_id))")!Zahl
  Erase mdpl ' 30.5.20
  If mpz > 0 Then
   ru = 0
   ReDim mdpl(mpz - 1)
'            "WHERE mp.pat_id = " & PID & ""
'   myFrag rTh, "SELECT mp.*,ma.metf,ma.sglt2 FROM wmedplan mp " & _
            "LEFT JOIN medarten ma ON ma.medikament=mp.medanfang " & _
            "WHERE mp.pat_id = " & pid & " AND mp.zeitpunkt=(SELECT MAX(zeitpunkt) FROM wmedplan WHERE pat_id=" & pid & ")"
'           "WHERE mp.pat_id = " & Pid & " AND mpnr=(SELECT MAX(mpnr) FROM medplan mpi WHERE pat_id=mp.pat_id AND zeitpunkt=(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=mp.pat_id))"
' 27.12.25:
   myFrag rTh, "SELECT mp.*,ma.metf,ma.sglt2 FROM lmp mp " & _
            "LEFT JOIN medarten ma ON ma.medikament=mp.medanfang " & _
            "WHERE mp.pat_id = " & pid
   Do While Not rTh.EOF
    If ru > UBound(mdpl) Then ReDim Preserve mdpl(ru)
    mdpl(ru).m.ab = rTh!ab
    mdpl(ru).m.absPos = rTh!absPos
    mdpl(ru).m.aktZeit = rTh!aktZeit
    mdpl(ru).m.bBed = rTh!bBed
    mdpl(ru).m.Bemerkung = rTh!Bemerkung
    mdpl(ru).m.Datum = rTh!Datum
    mdpl(ru).m.Einheit = rTh!Einheit
    mdpl(ru).m.FeldNr = rTh!FeldNr
    mdpl(ru).m.FID = rTh!FID
    mdpl(ru).m.Form = rTh!Form
    mdpl(ru).m.Grund = rTh!Grund
    mdpl(ru).m.MedAnfang = rTh!MedAnfang
    mdpl(ru).m.Medikament = rTh!Medikament
    If InStrB(UCase$(mdpl(ru).m.Medikament), "HCT") <> 0 Or (InStrB(UCase$(mdpl(ru).m.Medikament), "COMP") <> 0 And InStrB(UCase$(mdpl(ru).m.Medikament), "COMPLEX") = 0 And InStrB(UCase$(mdpl(ru).m.Medikament), "COMPACT") = 0 And InStrB(UCase$(mdpl(ru).m.Medikament), "DORZOCOMP") = 0 And InStrB(UCase$(mdpl(ru).m.Medikament), "DORZOLAMID") = 0 And InStrB(UCase$(mdpl(ru).m.Medikament), "TILIDIN") = 0 And InStrB(UCase$(mdpl(ru).m.Medikament), "PARACETAMOL") = 0) Then
     obHCT = True
     hctMed = mdpl(ru).m.Medikament
    End If
    mdpl(ru).m.mi = rTh!mi
    mdpl(ru).m.mo = rTh!mo
    mdpl(ru).m.MPNr = rTh!MPNr
    mdpl(ru).m.nm = rTh!nm
    mdpl(ru).m.Pat_id = rTh!Pat_id
    mdpl(ru).m.PZN = rTh!PZN
    mdpl(ru).m.Stärke = rTh!Stärke
    mdpl(ru).m.StByte = rTh!StByte
    mdpl(ru).m.Wirkstoff = rTh!Wirkstoff
    mdpl(ru).m.Zeitpunkt = rTh!Zeitpunkt
    mdpl(ru).m.Zn = rTh!Zn
    mdpl(ru).metf = Not IsNull(rTh!metf) And rTh!metf <> 0
    mdpl(ru).sglt = Not IsNull(rTh!sglt2) And rTh!sglt2 <> 0
'    (medikament LIKE '%comp%' AND NOT medikament RLIKE 'complex|dorzocomp|tilidin')
    rTh.MoveNext
    ru = ru + 1
   Loop
  End If ' mpz>0 THEN
  altPID = pid
 End If ' pid<>altpid
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in mplan/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' mplan

' Medikamentenangabe zu Zahl
Function medzz!(ByVal ST$)
 Dim b$, z$, nachziffer%, pos%
 If left$(ST, 1) = "(" Then ST = Mid$(ST, 2)
 If Right$(ST, 1) = ")" Then ST = left$(ST, Len(ST) - 1)
 If ST = "1/2" Then
  medzz = 0.5
 ElseIf IsNumeric(ST) Then
  medzz = REPLACE$(ST, ".", ",")
 ElseIf ST = "" Then
  medzz = 0
 ElseIf ST = "½" Then
  medzz = 0.5
 ElseIf ST = "¼" Then
  medzz = 0.25
 Else
  nachziffer = 0
  For pos = 1 To Len(ST)
   If InStrB("0123456789" + IIf(nachziffer, " ,.", ""), Mid$(ST, pos, 1)) = 0 And Mid$(ST, pos, 1) <> vNS Then
    ST = left$(ST, pos - 1) & " " & Mid$(ST, pos + 1)
   Else
    nachziffer = True
   End If
  Next
  ST = Trim$(ST)
  If IsNumeric(ST) Then medzz = ST
 End If
End Function ' medzz

' aufgerufen in dodoPlz
Function UKPDS(ByRef aRisk As Risk, pid$, gbdt As Date, dmseit$, ByRef falDiabDau%, obweibl%)
' 1.11.14 vorgezogen
' UKPDS-Risk bestimmen (1)
  Dim diabseit As Date
  Dim rsDia As New ADODB.Recordset
  aRisk.AgeDiagDiabetes = MAXvb((erbe(pid) - gbdt) * 0.002737925747, 20) ' / 365.24 ' bei Active-X-Fehler: an der Kommandozeile "regsvr32 u:\programmierung\riskeng.ocx" laufen lassen
  aRisk.DurationDiagnosedDiabetes = 0
  falDiabDau = -1
  Dim DSnum As Date
  If dmseit <> "" Then
   DSnum = MachNumerisch(dmseit, erstDatum:=True)
   If Not IsDate(dmseit) And IsDate(DSnum) And DSnum > #1/1/1920# And DSnum < #1/1/2100# Then
    aRisk.DurationDiagnosedDiabetes = Round((IIf(DSnum >= Now(), 0, Now() - DSnum)) / 365.24, 1)
    On Error Resume Next
    aRisk.AgeDiagDiabetes = Round(MAXvb((Year(DSnum) - Year(gbdt)), 20), 1) ' PAlter - aRisk.DurationDiagnosedDiabetes
    On Error GoTo fehler
    falDiabDau = 0
'  IF isnumeric(DMseit)
   ElseIf (dmseit Like "*#/##" Or dmseit Like "*#/####") Then ' cdate( geht nicht bei "12/10"
    Dim dumw$
    dumw = "01/" & dmseit
    If IsDate(dumw) Then
'     aRisk.AgeDiagDiabetes = ROUND(MAXvb((CDate(dumw) - gbdt) * 2.7379257474573E-03, 20), 1)
     DSnum = CDate(dumw)
     aRisk.AgeDiagDiabetes = MAXvb(AlterBei(DSnum, gbdt), 20)
'     aRisk.DurationDiagnosedDiabetes = (NOW() - CDate(dumw)) * 2.7379257474573E-03
     aRisk.DurationDiagnosedDiabetes = Round(AlterBei(Now(), DSnum), 1)
     falDiabDau = 0
    End If
   ElseIf IsDate(dmseit) Then
    DSnum = CDate(dmseit)
    If DSnum > #1/1/1920# And DSnum < #1/1/2100# Then
'     aRisk.AgeDiagDiabetes = ROUND(MAXvb((CDate(DMseit) - gbdt) * 2.73792574745373E-03, 20), 1) ' / 365.24
     aRisk.AgeDiagDiabetes = MAXvb(AlterBei(DSnum, gbdt), 20)
     Dim durdiagdiab#
'     durdiagdiab = (NOW() - CDate(DMseit)) * 2.73792574745373E-03 ' / 365.24
     durdiagdiab = AlterBei(Now(), DSnum)
     If durdiagdiab > 0 Then
      aRisk.DurationDiagnosedDiabetes = durdiagdiab
     End If
     falDiabDau = 0
    End If
   ' Else
   End If ' Not IsDate(dmseit) And IsDate(DSnum) And DSnum > #1/1/1920# And DSnum < #1/1/2100# Then else else
  End If ' NOT ISNULL(DMseit)
' UKPDS-Risk bestimmen (2)
  aRisk.Female = obweibl
  Set rsDia = Nothing
  myFrag rsDia, "SELECT 0 FROM `diagview` d WHERE d.pat_id = " & pid & " AND (d.gicd LIKE 'I48%')" ' AND COALESCE(d.Dggel,0)=0
  aRisk.AtrialFibrillation = Not rsDia.BOF()
  Set rsDia = Nothing
  aRisk.EthnicGroup = 1 ' 2 = Afro-Caribbean, 3 = Asian-Indian
  myFrag rsDia, "SELECT diagsicherheit d FROM `diagnosen` d WHERE d.pat_id = " & pid & " AND (d.icd LIKE 'F17%') AND NOT d.diagsicherheit IN ('A')" '  AND COALESCE(d.Dggel,0)=0
  If rsDia.EOF Then
   aRisk.SmokingStatus = 0
  ElseIf rsDia!d = "Z" Then
   aRisk.SmokingStatus = 1
  Else
   aRisk.SmokingStatus = 2
  End If
  Set rsDia = Nothing
  myFrag rsDia, "SELECT COUNT(0) ct FROM (SELECT 0 FROM rrparse WHERE pat_id = " & pid & " ORDER BY zeitpunkt DESC LIMIT 6) i"
  If Not rsDia.BOF Then
   If Not IsNull(rsDia!ct) Then aRisk.BloodPressure_Precision = rsDia!ct
  End If
  Set rsDia = Nothing
  myFrag rsDia, "SELECT avg(rrsyst) s, avg(rrdiast) d FROM (SELECT rrsyst,rrdiast FROM rrparse WHERE pat_id = " & pid & " ORDER BY zeitpunkt DESC LIMIT 6) i"
  If Not rsDia.BOF Then
   If Not IsNull(rsDia!s) Then aRisk.BloodPressure = rsDia!s
  End If
  aRisk.HbA1c_Precision = 0
  If lwZahl Then
   For aktlwx = 0 To lwZahl
    If obLabI(LA_HbA1c, lab(aktlwx)) Then ' Kreai
'    aRisk.HbA1c = replace$(Lab(aktlwx).wertsg, ".", ",")
     If lab(aktlwx).WertSg <> vNS Then
      On Error Resume Next
      aRisk.HbA1c = MachNumerisch(lab(aktlwx).WertSg)
      On Error GoTo fehler
      aRisk.HbA1c_Precision = 1
      Exit For
     End If
    End If
   Next aktlwx
  End If
'  For i = 0 To pzl
'   IF Üs(i) = "HbA1c" THEN
'    IF Not rs(i).BOF THEN
'     rs(i).MoveFirst
'     IF NOT ISNULL(rs(i)!wert) THEN
'      IF rs(i)!wert <> 0 THEN
'       IF rs(i)!wert <> "" THEN
'        aRisk.HbA1c = replace$(rs(i)!wert, ".", ",")
'        aRisk.HbA1c_Precision = 1
'       END IF
'      END IF
'     END IF
'    END IF
'    Exit For
'   END IF
'  Next i
  
  aRisk.LipidRatio_Precision = 0
  Dim LipRat#
  LipRat = 0
  If lwZahl Then
   For aktlwx = 0 To lwZahl
    If obLabI(LA_Hdl, lab(aktlwx)) Then ' Kreai
     If IsNumeric(lab(aktlwx).WertSg) Then If lab(aktlwx).WertSg <> 0 Then LipRat = 1# / MachNumerisch(lab(aktlwx).WertSg)
     Exit For
    End If
   Next aktlwx
  End If
'  For i = 0 To pzl
'   IF Üs(i) = "HDL" THEN
'    IF Not rs(i).BOF THEN
'     rs(i).MoveFirst
'     IF NOT ISNULL(rs(i)!wert) THEN
'      IF rs(i)!wert <> 0 THEN
'       IF IsNumeric(rs(i)!wert) THEN IF rs(i)!wert <> 0 THEN LipRat = 1# / replace$(rs(i)!wert, ".", ",")
'      END IF
'     END IF
'    END IF
'    Exit For
'   END IF
'  Next i
  If LipRat <> 0 Then
   If lwZahl Then
    For aktlwx = 0 To lwZahl
     If obLabI(LA_Chol, lab(aktlwx)) Then ' Kreai
      LipRat = LipRat * MachNumerisch(lab(aktlwx).WertSg)
      Dim rLipRat#
      rLipRat = Round(LipRat, 2)
      If rLipRat <> 0 Then
       On Error Resume Next
       aRisk.LipidRatio = rLipRat
       On Error GoTo fehler
       aRisk.LipidRatio_Precision = 1
      End If
      Exit For
     End If
    Next aktlwx
   End If ' LipRat <> 0 THEN
'   For i = 0 To pzl
'    IF Üs(i) = "Chol" THEN
'     IF Not rs(i).BOF THEN
'      rs(i).MoveFirst
'      IF NOT ISNULL(rs(i)!wert) THEN
'       IF rs(i)!wert <> vNS THEN
'        LipRat = LipRat * replace$(rs(i)!wert, ".", ",")
'        aRisk.LipidRatio = ROUND(LipRat, 2)
'        aRisk.LipidRatio_Precision = 1
'       END IF
'      END IF
'     END IF
'     Exit For
'    END IF
'   Next i
  End If ' liprat<>0
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in UKPDS/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' UKPDS

' aufgerufen in dodoPLZ
Function pruefICD%(ICDs)
 Dim i%, j%
 If SafeArrayGetDim(ICD) <> 0 Then
 For i = 0 To UBound(ICDs)
  For j = 0 To UBound(ICD)
   If left$(ICD(j), Len(ICDs(i))) = ICDs(i) And ((DSic(j) = " " Or DSic(j) = "G") Or (left$(ICD(j), 1) = "Z" And DSic(j) = "Z")) Then
    pruefICD = True
    Exit Function
   End If
  Next j
 Next i
 End If
End Function ' pruefICD%(ICDs$())

' in dodoplz, aber nur ohne php!
Sub tukopf(ByVal AusS As CString)
  AusS.AppVar (Array("    <STYLE TYPE='text/css'>", vbCrLf))
  AusS.AppVar (Array("    <!--", vbCrLf))
  AusS.AppVar (Array("        @page { size: 21cm 29.7cm; margin: 1cm }", vbCrLf))  ' text-transform:capitalize
  AusS.AppVar (Array("    @media screen {", vbCrLf))
  AusS.AppVar (Array("      @-webkit-keyframes blinck {", vbCrLf))
  AusS.AppVar (Array("            0% { color: red; }", vbCrLf))
  AusS.AppVar (Array("            100% { color: orange; }", vbCrLf))
  AusS.AppVar (Array("        }", vbCrLf))
  AusS.AppVar (Array("      @-moz-keyframes blinck {", vbCrLf))
  AusS.AppVar (Array("          0% { color: red; }", vbCrLf))
  AusS.AppVar (Array("          100% { color: orange; }", vbCrLf))
  AusS.AppVar (Array("      }", vbCrLf))
  AusS.AppVar (Array("      @-ms-keyframes blinck {", vbCrLf))
  AusS.AppVar (Array("          0% { color: red; }", vbCrLf))
  AusS.AppVar (Array("          100% { color: orange; }", vbCrLf))
  AusS.AppVar (Array("      }", vbCrLf))
  AusS.AppVar (Array("      @-o-keyframes blinck {", vbCrLf))
  AusS.AppVar (Array("          0% { color: red; }", vbCrLf))
  AusS.AppVar (Array("          100% { color: orange; }", vbCrLf))
  AusS.AppVar (Array("      }", vbCrLf))
  AusS.AppVar (Array("      @keyframes blinck {", vbCrLf))
  AusS.AppVar (Array("          0% { color: red; }", vbCrLf))
  AusS.AppVar (Array("          100% { color: orange; }", vbCrLf))
  AusS.AppVar (Array("      }", vbCrLf))
  AusS.AppVar (Array("      .cave {", vbCrLf))
  AusS.AppVar (Array("          -webkit-animation: blinck 0.4s linear infinite;", vbCrLf))
  AusS.AppVar (Array("          -moz-animation: blinck 0.4s linear infinite;", vbCrLf))
  AusS.AppVar (Array("          -ms-animation: blinck 0.4s linear infinite;", vbCrLf))
  AusS.AppVar (Array("          -o-animation: blinck 0.4s linear infinite;", vbCrLf))
  AusS.AppVar (Array("          animation: blinck 0.4s linear infinite;", vbCrLf))
  AusS.AppVar (Array("      }", vbCrLf))
  AusS.AppVar (Array("      .cave { font-weight:bold;color:red;text-decoration:blink; }", vbCrLf))
  AusS.AppVar (Array("      .treat { font-weight:bold;color:blue;text-decoration:underline; }", vbCrLf))
  AusS.AppVar (Array("      .path { font-weight:bold;color:blue;background-color:yellow;font-size:2.8mm;}", vbCrLf))
  AusS.AppVar (Array("      .gruen { font-weight:normal;font-size:2.8mm;color:green }", vbCrLf))
  AusS.AppVar (Array("      .schwanger { font-weight:normal;font-size:5mm;color:olive }", vbCrLf))
  AusS.AppVar (Array("      .lila { font-weight:normal;font-size:2.8mm;color:663399 }", vbCrLf))
  AusS.AppVar (Array("    }", vbCrLf))
  AusS.AppVar (Array("    @media print {", vbCrLf))
  AusS.AppVar (Array("      .cave { font-weight:bold;color:red;text-transform:capitalize;text-decoration:underline;background-color:yellow }", vbCrLf))
  AusS.AppVar (Array("      .treat { font-weight:bold;color:blue;background-color:yellow }", vbCrLf))
  AusS.AppVar (Array("      .path { font-weight:bold;color:blue;font-size:2.8mm;background-color:yellow }", vbCrLf))
  AusS.AppVar (Array("    }", vbCrLf))
  AusS.AppVar (Array("      .abstand { font-size:2.5mm }", vbCrLf))
  AusS.AppVar (Array("      .unauff { font-weight:normal;font-size:2.8mm }", vbCrLf))
  AusS.AppVar (Array("      .gruen { font-weight:normal;font-size:2.8mm;color:green }", vbCrLf))
  AusS.AppVar (Array("      .schwanger { font-weight:normal;font-size:5mm;color:olive }", vbCrLf))
  AusS.AppVar (Array("      .lila { font-weight:normal;font-size:2.8mm;color:663399 }", vbCrLf))
  AusS.AppVar (Array("      .schmal { width:2.5em;border-left-width:1.5mm;padding:0cm; }", vbCrLf))
  AusS.AppVar (Array("      .med { color:teal;font-weight:bold }", vbCrLf))
  AusS.AppVar (Array("      .diag { color:#600000;font-weight:bold }", vbCrLf))
  AusS.AppVar (Array("      .lnn { width:2.5em;border-left-width:1.5mm;padding:0cm; }", vbCrLf))
  AusS.AppVar (Array("      .lne { min-width:25em;width:2.5em;border-left-width:1.5mm;padding:0cm; }", vbCrLf))
  AusS.AppVar (Array("      .lno { width:2.5em;border-left-width:1.5mm;padding:0cm;color:blue; }", vbCrLf))
  AusS.AppVar (Array("      .lpn { width:2.5em;border-left-width:1.5mm;padding:0cm;font-weight:bold;background-color:yellow;font-size:2.8mm; }", vbCrLf))
  AusS.AppVar (Array("      .lpo { width:2.5em;border-left-width:1.5mm;padding:0cm;font-weight:bold;color:blue;background-color:yellow;font-size:2.8mm; }", vbCrLf))
  AusS.AppVar (Array("      body { margin:0 }", vbCrLf))
  AusS.AppVar (Array("      h1 { margin-bottom:0.1cm;font-size:4.5mm;font:normal }", vbCrLf))
  AusS.AppVar (Array("      h2 { margin-bottom:0.1cm;font-size:4.5mm;font:normal }", vbCrLf))
  AusS.AppVar (Array("      th { border-style:none solid;overflow:auto;font:bold;font-size:3mm }", vbCrLf))  ' padding-left:0.5em;
  AusS.AppVar (Array("      td { border-style:solid;max-width:15em;border-width:0.1mm;margin:0 }", vbCrLf))
  AusS.AppVar (Array("      td p { text-align:left;overflow:auto;max-height:10mm;}", vbCrLf))
  AusS.AppVar (Array("      tr { vertical-align:top }", vbCrLf))
  AusS.AppVar (Array("      table { font-size:2.8mm; border-collapse:collapse; }", vbCrLf))
  AusS.AppVar (Array("    -->", vbCrLf))
  AusS.AppVar (Array("    </STYLE>", vbCrLf))
  AusS.AppVar (Array("</HEAD>", vbCrLf))
'  AusS.AppVar (Array("<BODY LANG='de-DE' DIR='LTR'", IIF(obphp <> 0, "onload='document.aufgabenform.aufgaben.focus();'", ""), ">", vbCrLf))
  AusS.AppVar (Array("<BODY LANG='de-DE' DIR='LTR'>", vbCrLf))
  AusS.AppVar (Array("<a name='AnkerTop' href='#AnkerTop' accesskey='t'></a>", vbCrLf))
  AusS.AppVar (Array("<h1>", vbCrLf))  ' , <span" & IIF(bmi > 24.9, " class='path'", "") & ">" & bmiS & "</span>
End Sub      ' tukopf



' in: doPLZeinzeln, PLZausListe_Click, los (Lese), Command1_Click (PatListe, 2x), PatListe.plz, MFG_Click (Patliste), doPatientenlaufzettel, Patientenlaufzettel_Click (PatAuswahl und AnBog)
Function dodoplz(Pat_id$, plzVerz$, Optional Datum As Date, Optional Uhrzeit$, Optional mitanzeig%, Optional Arzt$, Optional zmaxges% = 8, Optional obphp% = 0) ' Patientenlaufzettel
 Dim DN$, DatNr&
 Dim Pause%(pzl) ' Zahl der noch zu pausierenden Zeilen für die aktuelle Spalte wegen Rowspan vorher
 Dim üwerg$()
 Dim UKtip$ ' String für Tip-Tool zu UKPDS-Risk
 Dim dmtyp$
 Dim GebDat As Date, PrivatTel$, PrivatTel_2$, PrivatMobil$, PrivatFax$, DienstTel$, email$, HzV$
 Dim obart%, obherzi%, obneph% ' Arteriosklerose, Herzinsuffizienz, Nephropathie
 Dim dmseit$, icdh$ ' ICD-Hauptteil
 Dim VName$, Vorgestellt As Date
 Dim sql0$
 Dim obfalsch%, wiefalar$
 Dim ru% ' Runde
 Dim pKVNR&
 Dim obpath%(pzl)
 Dim falDiabDau% ' falsche Diabetesdauer (`Diabetes seit` in der der Anamnese)
 
 m = 0: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next
 syscmd 4, "Erstelle gleich Patientenlaufzettel zu Pat_id: " & Pat_id
#If printtest Then
 DN = LiServer & "php\plz\uml5.html"
 DatNr = UTFOpen(DN)
 sUTF DatNr, "<?php"
 sUTF DatNr, "echo ""äöüßÄÖÜ€µ"";"
 sUTF DatNr, "?>"
 Close DatNr
 Exit Function
#End If
On Error GoTo fehler
' alte Dateien des gleichen Patienten löschen
 If obphp Then
  Dim Fil As File
  Dim loestr$(3), aktzls
  loestr(0) = phpV
  loestr(1) = phpVvorb
  loestr(2) = phpVfertig
  loestr(3) = phpVbehand
  For Each aktzls In loestr()
   For Each Fil In FSO.GetFolder(aktzls).Files
    If InStrB(Fil, ",Pid_" & Pat_id & ",") <> 0 Or InStrB(Fil, ",Pid " & Pat_id & ", Patientenlaufzettel") <> 0 Then
     On Error Resume Next
     FSO.DeleteFile (Fil)
     On Error GoTo fehler
    End If
   Next Fil
  Next aktzls
 Else ' obphp
  Dim erg$
  For Each Fil In FSO.GetFolder(plzVerz).Files
   If InStrB(Fil, ",Pid_" & Pat_id & ",") <> 0 Or InStrB(Fil, ",Pid " & Pat_id & ", Patientenlaufzettel") <> 0 Then
    On Error Resume Next
    Kill Fil
    On Error GoTo fehler
   End If
  Next Fil
 End If ' obphp else
' bis hier 0,4s
 m = 1: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next

 Dim rnam As New ADODB.Recordset, rAn As New ADODB.Recordset, rFl As ADODB.Recordset, rFlb As New ADODB.Recordset
 Dim rTha As New ADODB.Recordset
 Dim rHAV As New ADODB.Recordset
 Dim i&, Gewicht!, Wert$, j&, k&, ianf&, iend&, FlbZ&, aktlwx&  ', lSQL$ ' bmi!, bmiS$,
 Dim gefunden&
 Dim AusS As New CString
 Dim iStri(3), iDN$(3), iob!(3)
 For i = 0 To 3
   iDN(i) = LiServer & "php\php\i" & i & "S.php"
   If GetFileTime(iDN(i), mftlastwritetime) + 1 < Now() Then
    Set iStri(i) = New CString
    iob(i) = True
   End If
 Next i
 Dim aktlz&(pzl), angef&(pzl) ' Ordinalzahl des in der aktuell bearbeiteten Zeile einzutragenden Wertes der aktuellen Spalte
 
 Const SpZahl% = 11
' Public Enum ParNr
'  RRi=3 ' Parameter für Blutdruck
' END Enum
 'Const zmaxges% = 8 ' Maximale Zeilenzahl pro Parameter
 Dim DateiNameRoh$, DateiName$, DateiNameAkt$, DateiNameAktRoh$
' Dim obAvan% ' ob Avandia verordnet wird
' Dim wasAvan$ ' was genau verordnet wurde
 On Error GoTo fehler
 
 Erase pathol
 Erase nurpath
 Erase Fqmin
 Erase ÜS
 Erase Titel, RowSp, weite, rsz, kritnr, sql, rs, rund
 If zmaxges = 0 Then zmaxges = maxzmaxges
 ReDim pathol(pzl, zmaxges)
 
' DMPString soll vor getHausarzt1 aufgerufen werden, da rNa(0) dort benötigt wird
  m = 14: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  Call DMPString(CLng(Pat_id), üdt, True)

' Hausärzte sollen vor myFrag rs(i) aufgerufen werden, da DBCn bei InsKorr unterbrochen werden kann
'  Call getHausarzt(CLng(Pat_ID), üwerg)
'  Dim rFa() As Faelle
'  Dim rKv1() As kvnrue
'  While True
  ReDim rFa(0)
  ReDim rKv(0)
  Call getHausarzt1(üwerg, rFa, rKv, , CLng(Pat_id), auchwir:=True, vonwo:="Erstelle Patientenlaufzettel")
  m = 25: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
'  Wend
  obpath(0) = -1
  For i = 0 To UBound(üwerg, 2)
   If InStrB(üwerg(10, i), "HA") <> 0 Then
    obpath(0) = 0
    Dim KVNrS$
    KVNrS = üwerg(12, i)
    If IsNumeric(KVNrS) Then
     pKVNR = KVNrS
    End If
    Exit For
   End If ' InStrB(üwerg(10, i), "HA") <> 0 Then
  Next i ' = 0 To UBound(üwerg, 2)
 
 
 On Error GoTo fehler
 If LenB(DBCn) = 0 Or LenB(DBCnS) = 0 Then Call acon(quelleT) ' DBCn.ConnectionString
 Call doPatvonMO(CStr(Pat_id), obmitFormularen:=True, obpruef:=True) ' 20.9.25
 Dim rnamBOF%, iru%
 Dim dmpbeg#, dmpklass As DMPEnum, dmpkhkklass As DMPEnum, dmpcopdklass As DMPEnum, dmpabklass As DMPEnum, dmpkhkbeg#, dmpcopdbeg#, dmpabbeg#
 For iru = 1 To 2
  Set rnam = Nothing
  sql0 = "SELECT n.*" & vbCrLf & _
  ",(obs AND lanrid<>1) OR (obk AND lanrid<>2) OR (obh AND lanrid<>5) obfalsch" & vbCrLf & _
  ",CONCAT('Desktop: ', IF(obs,'gelb,',''),IF(obk,'blau,',''),IF(obh,'gruen,',''),' Arztzuord.: ',IF(lanrid=1,'Schade',IF(lanrid=5,'Hammerschmidt','Kothny'))) wiefalar" & vbCrLf & _
  "FROM namen n LEFT JOIN faelle f USING(pat_id)" & vbCrLf & _
  "WHERE pat_id = " & Pat_id & " ORDER BY bhfb DESC LIMIT 1"
  myFrag rnam, sql0, adOpenStatic, DBCn, adLockReadOnly
  rnamBOF = rnam.BOF
  If rnamBOF Then
   If iru = 1 Then
    Call doPatvonMO(CStr(Pat_id), True)
   Else ' iru = 1 Then
'  MsgBox ("Keinen Patienten zu " & Pat_id & " gefunden!")
    Lese.Ausgeb "Keinen Patienten zu " & Pat_id & " gefunden!", True
    Exit Function
   End If ' iru = 1 Then Else
  Else
   Exit For
  End If ' rnamBOF
 Next iru
 On Error Resume Next ' 72910
 obfalsch = rnam!obfalsch
 On Error GoTo fehler
 wiefalar = rnam!wiefalar
 dmpklass = rnam!dmpklass
 dmpkhkklass = rnam!dmpkhkklass
 dmpcopdklass = rnam!dmpcopdklass
 dmpabklass = rnam!dmpabklass
 dmpbeg = rnam!dmpbeg
 dmpkhkbeg = rnam!dmpkhkbeg
 dmpcopdbeg = rnam!dmpcopdbeg
 dmpabbeg = rnam!dmpabbeg
 NName = rnam!Nachname
 VName = rnam!Vorname
 obweibl = (rnam!geschlecht = "w")
 GebDat = rnam!GebDat
 PrivatTel = rnam!PrivatTel
 PrivatTel_2 = rnam!PrivatTel_2
 PrivatMobil = rnam!PrivatMobil
 PrivatFax = rnam!PrivatFax
 DienstTel = rnam!DienstTel
 email = rnam!email
 HzV = rnam!HzV
 rNa(0).obk = rnam!obk
 rNa(0).obs = rnam!obs
 rNa(0).obh = rnam!obh
 PAlter = AlterBei(Now(), rnam!GebDat)
 
 myFrag rAn, "SELECT COALESCE(`diabetes seit`,'') dmseit,COALESCE(vorgestellt,(SELECT MIN(bhfb) FROM faelle WHERE pat_id=" & Pat_id & ")) vorg,Diabetestyp FROM `anamnesebogen` WHERE pat_id = " & Pat_id
 If rAn.State <> 0 Then
  If Not rAn.BOF Then
   dmseit = rAn!dmseit
'   Vorgestellt = rAn!vorg
   If Not IsNull(rAn!vorg) Then Vorgestellt = rAn!vorg
  End If
 End If
 
 Dim DiagTab() As CString
'anfang:
 m = 2: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
 Call DiagString(Pat_id, DiagTab, , , dmseit) ' 0,35s
 Call UKPDS(aRisk, Pat_id, GebDat, dmseit, falDiabDau, obweibl)
 m = 3: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
 
 dmtyp = "-"
 If SafeArrayGetDim(ICD) <> 0 Then
  For ru = 0 To UBound(ICD)
   If DSic(ru) = " " Or DSic(ru) = "G" Then
'"  SET typ = (SELECT CASE substr(icd,1,1) when 'E' THEN case substr(icd,3,1) when 1 THEN '2' when 0 THEN '1' when 3 THEN 's' ELSE 'u' END when 'O' THEN 'g' when 'R' THEN 'p' ELSE '-' END FROM " & vbCrLf
    icdh = left$(ICD(ru), 3)
    If dmtyp = "-" Then
     If left$(ICD(ru), 2) = "E1" Then
      Select Case Mid$(ICD(ru), 3, 1)
       Case "1": dmtyp = "2"
       Case "0": dmtyp = "1"
       Case "3": dmtyp = "s"
       Case Else: dmtyp = "u"
      End Select
      GoTo weiter
     ElseIf ICD(ru) = "O24.4" Then
      dmtyp = "g"
      GoTo weiter
     ElseIf ICD(ru) Like "R73.0*" Then
      dmtyp = "p"
      GoTo weiter
     End If ' Left$(icd(ru),2) elseif elseif
'    End If ' dmtyp = "-"
    ElseIf icdh = "F03" Then
     obdement = True
    ElseIf icdh = "I42" Or icdh = "I43" Or icdh = "I50" Or icdh = "I25" Then
     obherzi = True ' Herzschwäche
    ElseIf icdh = "I25" Then
     obart = True 'KHE
    ElseIf icdh = "I70" Or icdh = "I73" Or ICD(ru) = "I79.2" Then
     obart = True ' pAVK
    ElseIf ICD(ru) = "N08.3" Or icdh = "N18" Or icdh = "N19" Then
     obneph = True
    ElseIf ICD(ru) = "H36.0" Then
     obreti = True
    End If
   ElseIf DSic(ru) = "Z" Then
    If icdh = "I21" Then
     obherzi = True
     obart = True  ' Z.n. Myokardinarkt
    ElseIf icdh = "I63" Or icdh = "I64" Or icdh = "I65" Or icdh = "I66" Then
     obart = True ' Z.n. Schlaganfall
    ElseIf icdh = "I74" Then
     obart = True ' Z.n. Embolie
    End If ' icdh = "I21"
   End If ' DSic(ru) = " " OR DSic(ru) = "G" THEN elseif DSic(ru)= "Z"
weiter:
  Next ru
 End If ' safearraygetdim(icd) <> 0
 
' dmtyp = myEFrag("SELECT dmtyp(" & Pat_ID & ")").Fields(0)
 If dmtyp = "1" Or dmtyp = "2" Or dmtyp = "s" Or dmtyp = "u" Then obdm = -1 Else obdm = 0
 If Not rnamBOF Then
  syscmd 4, "Erstelle Patientenlaufzettel für: " & NName & ", " & VName & " (Pat_id: " & Pat_id & ")"
 Else
  syscmd 4, "Patient mit Pat_id " & Pat_id & " nicht in Datenbank: " & DefDB(DBCn) & "gefunden, Patientenlaufzettel daher nicht erstellbar!"
 End If
' IF Pat_ID = 2183 THEN Stop
 'myFrag rFl, "SELECT k.*, f.* FROM `faelle` f LEFT JOIN `kassenliste` k on(((f.VKNr = k.vknr) AND (f.IK = k.IK))) WHERE f.pat_id = " & Pat_ID & " GROUP BY f.pat_id ORDER BY qanf DESC"
 ' Feld !Kasse wird überlagert von gleichnamigem Feld aus Faelle, führt zu unvorsehbaren Ergebnissen (Fall 706)
 
 m = 4: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
 Dim rFlSchGr%, rFlKateg$, obDMPPrüf%, obLeist%, vorET As Date, rFlkrkasse$
 obLeist = True
 If obdm Then obDMPPrüf = True Else obDMPPrüf = False
' myFrag rFl, "SELECT k.name KrKasse, k.*, f.* FROM (SELECT f.* FROM faelle f WHERE f.pat_id = " & Pat_id & " AND qanf = (SELECT MAX(qanf) FROM faelle WHERE pat_id = " & Pat_id & ")) f LEFT JOIN `kassenliste` k ON f.VKNr=k.vknr AND f.IK=k.IK ORDER BY MID(qanf,2) DESC, qanf DESC" ' schgr 21.6.16
' myFrag rFl, "SELECT schgr,kateg,k.name krkasse" & vbCrLf & _
          ",(SELECT voret FROM sws WHERE pat_id=f.pat_id AND voret>qanf LIMIT 1) voret" & vbCrLf & _
          "FROM faelle f" & vbCrLf & _
          "LEFT JOIN kassenliste k ON f.VKNr=k.vknr AND f.ik = k.ik" & vbCrLf & _
          "WHERE f.pat_id=" & Pat_id & " AND qanf=(SELECT MAX(qanf) FROM faelle WHERE pat_id=f.pat_id)" & vbCrLf & _
          "ORDER BY MID(qanf,2) DESC, qanf DESC", adOpenStatic, DBCn, adLockReadOnly ' 11.4.22
sql0 = _
"SELECT schgr,kateg,k.name krkasse" & vbCrLf & _
",(SELECT voret FROM sws WHERE pat_id=f.pat_id AND voret>bhfb LIMIT 1) voret" & vbCrLf & _
"FROM faelle f" & vbCrLf & _
"LEFT JOIN kassenliste k ON k.id=f.kid" & vbCrLf & _
"WHERE f.Pat_id =" & Pat_id & vbCrLf & _
"ORDER BY bhfb DESC LIMIT 1;"
 myFrag rFl, sql0, adOpenStatic, DBCn, adLockReadOnly
 If rFl.State <> 0 Then
  If Not rFl.BOF Then
   rFlSchGr = rFl!SchGr
   On Error Resume Next
   rFlKateg = rFl!kateg
   On Error GoTo fehler
   If rFlSchGr = 90 Or rFlKateg = "PBe" Or rFlKateg = "SHV" Then obDMPPrüf = 0 ' rFlKateg = "LKK" Or
   If Not IsNull(rFl!vorET) Then vorET = rFl!vorET
   If rFl!SchGr = 90 Then obLeist = False
   If Not IsNull(rFl!krkasse) Then rFlkrkasse = rFl!krkasse
  End If ' rFl.BOF Then
 End If ' Not rFl.BOF Then
 m = 5: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  
  myFrag rTha, "SELECT therart,zp FROM therarten tha WHERE pat_id=" & Pat_id & " ORDER BY zp DESC,mpnr DESC LIMIT 1"
  If Not rTha.BOF Then therart = rTha!therart
  If Datum = 0 Then Datum = Now
  If obphp Then
   DateiNameRoh = NName & "_" & VName & ",Pid_" & Pat_id & IIf(LenB(Arzt) <> 0, "," & Arzt, vNS) ' replace$(CDate(Uhrzeit), ":", ".") & "_" & Format(Datum, "dd.mm.yyyy") &
   DateiNameRoh = REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(DateiNameRoh, "ä", "ae"), "ö", "oe"), "ü", "ue"), "Ä", "Ae"), "Ö", "Oe"), "Ü", "Ue"), "ß", "ss"), "'", "")
  Else
   DateiNameRoh = NName & " " & VName & ",   Pid " & Pat_id & IIf(LenB(Arzt) <> 0, ", " & Arzt, vNS) & ", Patientenlaufzettel" ' replace$(CDate(Uhrzeit), ":", ".") & " " & Format(Datum, "dd.mm.yyyy") &
  End If
 m = 6: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
'  Dim i%
'  PAlter = (NOW() - rnam!GebDat) * 2.73792574745373E-03 ' 1/365,24
  
  On Error GoTo fehler
  If (obphp <> 0) Then AusS.AppVar (Array("<?php session_start(); ?>", vbCrLf))
  AusS.AppVar (Array("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>", vbCrLf))
  AusS.AppVar (Array("<HTML>", vbCrLf))
  If obphp Then
   AusS.AppVar (Array("<HEAD>", vbCrLf))
   AusS.AppVar (Array("    <META HTTP-EQUIV='CONTENT-TYPE' CONTENT='text/html; charset=windows-1252'>", vbCrLf))
   AusS.AppVar (Array("    <TITLE>" & NName & " " & VName & " (Pat_id:" & Pat_id & "): Patientenlaufzettel vom " & Now & "</TITLE>", vbCrLf))
   AusS.AppVar (Array("    <META NAME='AUTHOR' CONTENT='Gerald Schade'>", vbCrLf))
   AusS.AppVar (Array("    <META NAME='CREATED' CONTENT=" & Format(Now, "yyyymmdd") & ";" & Format(Now, "hhmmss") & "00'>", vbCrLf))
   AusS.AppVar (Array("    <META NAME='CHANGEDBY' CONTENT='Gerald Schade'>", vbCrLf))
   AusS.AppVar (Array("    <META NAME='CHANGED' CONTENT=" & Format(Now, "yyyymmdd") & ";" & Format(Now, "hhmmss") & "00'>", vbCrLf))
   AusS.AppVar (Array("    <META NAME='CHANGEDBY' CONTENT='Gerald Schade'>", vbCrLf))
   AusS.AppVar Array("<?php include ""../php/kopf.html"" ?>", vbCrLf)
  Else ' obphp
   tukopf AusS
  End If ' obphp else
  
  ' Wenn keine Diagnose Hypertonie und Blutdruckwerte normal, dann keinen RR-Vergleich
 m = 7: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
 
 Call tabelleInplz(Pat_id, zmaxges, rFlSchGr)
  
  m = 18: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
' nurpath und Fqmin werden hier oberhalb nur festgelegt, unterhalb nur abgefragt
' pathol wird unterhalb nur abgefragt

'Beginn mit Schreiben der Daten
  Dim DmPStr As New CString, dmpfarbe$
  dmpfarbe = "class='unauff' style='color:blue'"
  If obDMPPrüf Then
   DmPStr = ",  DMP-Diab: "
   Select Case dmpklass
    Case unb: DmPStr.Append "unbek (": dmpfarbe = "class='cave'"
    Case nein: DmPStr.Append "nein (": dmpfarbe = "class='cave'"
    Case HA: DmPStr.Append "HA ("
    Case hier: DmPStr.Append "hier ("
    Case ausg: DmPStr.Append "ausgeschrieben ("
   End Select
   DmPStr.AppVar Array(IIf(dmpbeg = 0, "", Format(dmpbeg, "d.m.yy")), ")")
   If dmpkhkklass <> 0 Then
    DmPStr.Append ", DMP-KHK: "
    Select Case dmpkhkklass
     Case unb: DmPStr.Append "unbek ("
     Case nein: DmPStr.Append "nein ("
     Case HA: DmPStr.Append "HA ("
     Case hier: DmPStr.Append "hier ("
     Case ausg: DmPStr.Append "ausgeschrieben ("
    End Select
    DmPStr.AppVar Array(Format(dmpkhkbeg, "d.m.yy"), ")")
   End If
   If dmpcopdklass <> 0 Then
    DmPStr.Append ", DMP-COPD: "
    Select Case dmpcopdklass
     Case unb: DmPStr.Append "unbek ("
     Case nein: DmPStr.Append "nein ("
     Case HA: DmPStr.Append "HA ("
     Case hier: DmPStr.Append "hier ("
     Case ausg: DmPStr.Append "ausgeschrieben ("
    End Select
    DmPStr.AppVar Array(Format(dmpcopdbeg, "d.m.yy"), ")")
   End If
   If dmpabklass <> 0 Then
    DmPStr.Append ", DMP-AB: "
    Select Case dmpabklass
     Case unb: DmPStr.Append "unbek ("
     Case nein: DmPStr.Append "nein ("
     Case HA: DmPStr.Append "HA ("
     Case hier: DmPStr.Append "hier ("
     Case ausg: DmPStr.Append "ausgeschrieben ("
    End Select
   End If
  End If ' obDMPPrüf Then
  
  m = 19: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  If obdm Then
' 24.4.20 auch noch vorgezogen
   m = 20: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
   m = 21: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
   m = 22: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
'  vorverlegt wegen aRisk.hba1c 30.1.23
'   Call UKPDS(aRisk, Pat_id, GebDat, dmseit, falDiabDau, obweibl)
   m = 23: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
   UKtip = "D-Dauer: " & aRisk.DurationDiagnosedDiabetes & "a, weibl: " & aRisk.Female & ", VHFli: " & aRisk.AtrialFibrillation & ", Ethn: " & "kaukas" & ", Rauch: " & IIf(aRisk.SmokingStatus = 0, "nie", IIf(aRisk.SmokingStatus = 1, "früher", "ja")) & ", RRsyst: " & aRisk.BloodPressure & ", gemess. " & aRisk.BloodPressure_Precision & "mal, HbA1c: " & aRisk.HbA1c & ", gemess. " & aRisk.HbA1c_Precision & " mal, Chol/HDL: " & aRisk.LipidRatio & ", gemess. " & aRisk.LipidRatio_Precision & " mal"
   Dim cR$, CU$, CL$, CFR$, CFU$, CFL$, SR$, SU$, SL$, SFR$, SFU$, SFL$
   cR = Round(aRisk.CardiacRisk * 100)
   CU = Round(aRisk.CardiacUpper * 100)
   CL = Round(aRisk.CardiacLower * 100)
   CFR = Round(aRisk.CardiacFatalRisk * 100)
   CFU = Round(aRisk.CardiacUpper * 100)
   CFL = Round(aRisk.CardiacFatalLower * 100)
   SR = Round(aRisk.StrokeRisk * 100)
   SU = Round(aRisk.StrokeUpper * 100)
   SL = Round(aRisk.StrokeLower * 100)
   SFR = Round(aRisk.StrokeFatalRisk * 100)
   SFU = Round(aRisk.StrokeUpper * 100)
   SFL = Round(aRisk.StrokeFatalLower * 100)
  End If ' IF obdm
  
#If False Then
'  If Not obphp Then ' kommt also nie vor
'   If obphp Then
'    AusS.AppVar (Array("<a href=""http://" & LiName & "/vorb"">V</a>", vbCrLf))
'    AusS.AppVar (Array("<a href=""http://" & LiName & "/behand"">B</a>", vbCrLf))
'    AusS.AppVar (Array("<a href=""http://" & LiName & "/fertig"">F</a>", vbCrLf))
'    AusS.AppVar (Array("<a href=""http://" & LiName & "/plz"" accesskey='z'>&#8598;</a>", vbCrLf))
'   End If
'  End If
#End If
  ' Farbe je nach Diabetestyp
  ' 255,204,229 = #ffcce5 rötlich; 229,204,255 = #E5CCFF bläulich; #ffffcc gelblich
'  doMarkierungen CLng(Pat_id), True ' in doPatvonMO verschoben 20.9.25
  Dim notiz$
  notiz = Trim$(REPLACE$(REPLACE$(rnam!info & IIf(rnam!info = "", "", " ") & rnam!notiz, Chr(13), ""), Chr(10), ""))
  AusS.AppVar (Array(" ", IIf(dmtyp = "1" Or dmtyp = "2" Or dmtyp = "g", "<span style='background-color:" & IIf(dmtyp = "1", "#ff8fc7", IIf(dmtyp = "g", "#ffffde", "#efe0ff")) & "'", ""), "<B><span title='", VName, " ", NName, ", ", rnam!strasse, ", ", rnam!plz, " ", rnam!ort, ", Tel1: ", PrivatTel, ", Tel2: ", PrivatTel_2, ", Mobil:", PrivatMobil, ", Fax: ", PrivatFax, ", Diensttel: ", DienstTel & ", Email: ", email, "'>", IIf(vorET > Now(), "<span class='schwanger'>", ""), _
  GesNamFn(rnam), "</span></B>, *", Format(rnam!GebDat, "d.m.yy"), " (", PAlter, "a,&" & IIf(rnam!geschlecht = "w", "fe", "") & "male;), <span style='color:blue'><span class='unauff'>&nbsp;&nbsp;Pat_id: </span>", Pat_id, "</span><span id = 'unauff'>,", IIf(obdm, "&nbsp;&nbsp;D.m.: ", ""), IIf(obdm, dmseit, ""), ",&nbsp;&nbsp;<span style=""font-weight:normal"">vorgestellt: </span>", Format(Vorgestellt, "d.m.yy"), ",&nbsp;&nbsp;</span><span style='font-size:smaller;font-weight:normal'>für: ", Format(Datum, "d.m.yy"), " ", Format(Uhrzeit, "hh:mm"), ",</span>&nbsp;<span class='unauff'>", IIf(notiz = "", "", notiz & ",&nbsp;&nbsp;"), IIf(obdm, "Ther.zul: ", ""), "</span>", IIf(obdm, therart, ""), "<span " & dmpfarbe & ">", DmPStr, " </span><span style='background-color:black'>", IIf(rNa(0).obk <> 0, " &#x1F7E6;", ""), IIf(rNa(0).obs <> 0, "&#x1F7E8;", ""), IIf(rNa(0).obh <> 0, "&#x1F7E9;", ""), "<button type=""button"" onclick=location.href=""oeffneverz:" & Pat_id & _
  """>Da<u>t</u>eien</button></span></h1>", vbCrLf))
' TherapieArtEinzelnFestlegen(CLng(Pat_ID), rAn) & "</span></h1>" ' VName, " ", NName
  ' * 2.73792574745373E-03 ' 1/365,24
  AusS.AppVar (Array("</h1>", vbCrLf))
'  IF Not obhierdmp(rNam!Notiz) THEN
'  Call obhierdmp(rnam!Notiz, , , , obdmp)
'  IF Not obdmp THEN
  AusS.AppVar (Array("</h1>", vbCrLf))
  AusS.AppVar (Array("", vbCrLf))
  m = 24: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  
  
  AusS.AppVar (Array("<div class='unauff'", IIf(obpath(0), " class='cave'", ""), ">", vbCrLf))
  For i = 0 To UBound(üwerg, 2)
   If üwerg(10, i) <> vNS Then
    AusS.AppVar (Array("<span title='", üwerg(1, i), ", ", üwerg(2, i), ", ", üwerg(3, i), ", ", üwerg(8, i), ", Tel: ", üwerg(13, i), ", Fax: ", üwerg(4, i), "'>", vbCrLf))
    AusS.AppVar (Array(üwerg(10, i), ": ", üwerg(1, i), IIf(i < 2, " (T:" & üwerg(13, i) & " F: " & üwerg(4, i) & IIf(üwerg(6, i) = "X", ", DMP 2", "") & IIf(üwerg(7, i) = "X", ", DMP 1", "") & ")", ""), ", ", vbCrLf))
    AusS.AppVar (Array("</span>", vbCrLf))
   End If ' üwerg(10, i) <> vNS Then
  Next i
  AusS.AppVar (Array("</div>", vbCrLf))
  If MOtot Then
   AusS.AppVar (Array("<div class='cave'>Dieser Laufzettel wurde ohne Echtzeit-Verbindung zur MO-Datenbank erstellt!</div>", vbCrLf))
  Else ' MOtot Then Else
   If obpath(0) Then AusS.AppVar (Array("<div class='cave'>Hausarzt nicht richtig in Medical Office eingetragen</div>", vbCrLf))
  End If ' MOtot Then
  
  Dim obGU%
  Dim rLei As New ADODB.Recordset
  obGU = 0
  
  If obLeist Then
   If pKVNR = KVNr And PAlter > 35 Then
    myFrag rLei, "SELECT zeitpunkt FROM `leistungen` WHERE pat_id = " & Pat_id & " AND leistung = '01732' ORDER BY zeitpunkt DESC"
    If rLei.BOF Then
     obGU = True
    ElseIf fctQAnf(ZQuart(Now())) - fctQAnf(ZQuart(rLei!Zeitpunkt)) >= 2 * 365 Then
     obGU = True
    End If
    If obGU = True Then
     AusS.AppVar (Array("<div class='cave'>Gesundheitsuntersuchung fällig</div>", vbCrLf))
    End If ' obGU = True Then
   End If ' pKVNR = KVNr And PAlter > 35 Then
  End If ' if obLeist
  m = 26: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
'  IF Not rFl.BOF THEN ' 5.3.09
'   IF rFl!Kateg = "AOK" AND (pKVNr = kvnr OR obpath(0) = -1) THEN
'    SET rHAV = Nothing
'    myFrag rHAV, "SELECT * FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND art = 'hav'"
'    IF rHAV.BOF THEN
'     AusS.AppVar (Array("<div class='cave'>Hiesiger AOK-Patient ohne Hausarztvertrag?</div>", vbCrLf))
'    END IF
'   END IF
'  END IF
  
'  IF Not rFl.EOF THEN
'   SELECT CASE rnam!dmpklass
'    Case 0, 1
'     IF rFl!Kateg <> "LKK" THEN
'      AusS.AppVar (Array("<div class='unauff'><span id = 'cave'>Nicht im DMP-Diabetes</span>: ", IIF(rnam!dmpklass = unbek, "(&Oslash;&nbsp;Eintrag)", "DMP nein"), "</div>", vbCrLf))
'     END IF
'    Case 2
'      AusS.AppVar (Array("<div class='unauff'>DMP-Diabetes: HA ", "</div>", vbCrLf))
'    Case 3
'      AusS.AppVar (Array("<div class='unauff'>DMP-Diabetes: hier ", "</div>", vbCrLf))
'   END SELECT
'  END IF
  Dim rsauf As New ADODB.Recordset, BhFB$, Auftrag$, Verdacht$, Befund$
  BhFB = vNS
  Auftrag = vNS
  Verdacht = vNS
  Befund = vNS
  myFrag rsauf, "SELECT quartal, DATE_FORMAT(bhfb,'%d.%m.%y') bhfb, auftrag, verdacht, befund FROM `faelle` WHERE pat_id = " & Pat_id & " AND ((NOT ISNULL(auftrag) AND auftrag <> '') OR (NOT ISNULL(verdacht) AND verdacht <> '') OR (NOT ISNULL(befund) AND befund <> '')) ORDER BY bhfb DESC"
  If Not rsauf.BOF Then
   BhFB = rsauf!BhFB
   Auftrag = rsauf!Auftrag
   Verdacht = rsauf!Verdacht
   Befund = rsauf!Befund
  End If
  Set rsauf = Nothing
  AusS.AppVar (Array("<div class='lila'>" & BhFB & ": Auftrag:<span class='gruen'> " & Auftrag & "</span>" & " Verdacht:<span class='gruen'> " & Verdacht & "</span>" & " Befund:<span class='gruen'> " & Befund & "</span>" & "</div>", vbCrLf))
  If (obphp <> 0) Then
   AusS.AppVar (Array("<?php ", vbCrLf))
   AusS.AppVar (Array(" $pat_id=", Pat_id, ";", vbCrLf))
   AusS.AppVar Array(" $_SESSION['tel']=""T:", PrivatTel, ",T2:", PrivatTel_2, ",M:", PrivatMobil, ",Fx:", PrivatFax, ",D:", DienstTel, ",E:", email, """;", vbCrLf)
   AusS.AppVar Array(" $telnr=1;", vbCrLf)
   AusS.AppVar Array(" $_SESSION['ePA runterladen']=1;", vbCrLf)

' Liste der ganz fehlenden Parameter
   For i = 0 To 3
    If iob(i) Then iStri(i).AppVar Array("<?php ", vbCrLf)
   Next i
   AusS.AppVar Array(" $_SESSION['obtelnr']=1;", vbCrLf)
   If iob(3) Then iStri(3).AppVar Array(" $_SESSION['obtelnr']=0;", vbCrLf)
'  Dim razu As New ADODB.Recordset ' 20.9.25: nach oben verschoben
'  myFrag razu, "SELECT (obs AND lanrid<>1) OR (obk AND lanrid<>2) OR (obh AND lanrid<>5) obfalsch," & vbCrLf & _
     "CONCAT('Desktop: ', IF(obs,'gelb,',''),IF(obk,'blau,',''),IF(obh,'gruen,',''),' Arztzuord.: ',IF(lanrid=1,'Schade',IF(lanrid=5,'Hammerschmidt','Kothny'))) wiefalar FROM (" & vbCrLf & _
     "SELECT lanrid,f.pat_id,obk,obs,obh FROM namen n JOIN faelle f USING (pat_id) WHERE pat_id=" & Pat_ID & " ORDER BY  bhfb DESC LIMIT 1)i"
''     "COALESCE((SELECT 1 FROM desktop WHERE pat_id = f.pat_id AND iconpath RLIKE '4eckblau' AND showasnote=0 LIMIT 1),0) obk," & vbCrLf & _
     "COALESCE((SELECT 1 FROM desktop WHERE pat_id = f.pat_id AND iconpath RLIKE '4eckgelb' AND showasnote=0 LIMIT 1),0) obs," & vbCrLf & _
     "COALESCE((SELECT 1 FROM desktop WHERE pat_id = f.pat_id AND iconpath RLIKE '4EckHellgruen' AND showasnote=0 LIMIT 1),0) obh," & vbCrLf & _
     "lanrid,f.pat_id FROM faelle f  WHERE pat_id=" & Pat_ID & " ORDER BY bhfb DESC LIMIT 1) i"
'  If Not razu.EOF Then
   AusS.AppVar Array(" $_SESSION['falarzt']=", IIf(obfalsch, "1", "0"), ";", vbCrLf)
   If obfalsch <> 0 Then
    AusS.AppVar Array(" $_SESSION['wiefalar']='", wiefalar, "';", vbCrLf)
   End If ' obfalsch <> 0 Then
'  End If ' not razu.EOF
  For i = 0 To pzl
'   If i = 29 Or i = 53 Then Stop
   If ÜS(i) <> "" Then
    AusS.AppVar Array(" $_SESSION['ob", ÜS(i), "']=")
    If rsz(i) = 0 And nurpath(i) < 2 And Fqmin(i) <> 0 Then ' letzteres 15.11.08
     AusS.AppVar (Array("1;", vbCrLf))
    Else
     AusS.AppVar (Array("0;", vbCrLf))
    End If
    If iob(0) Then
     If i > 0 Then iStri(0).Append " } ELSE "
     iStri(0).AppVar Array(" IF(isset($_POST['", ÜS(i), "'])) {", vbCrLf)
     iStri(0).AppVar Array("    $_SESSION['", ÜS(i), "']=!$_SESSION['", ÜS(i), "'];", vbCrLf)
    End If
    If iob(1) Then
     iStri(1).AppVar Array("    IF ($_SESSION['ob", ÜS(i), "']) {", vbCrLf)
     iStri(1).AppVar Array("      IF ($_SESSION['", ÜS(i), "']) {", vbCrLf)
     iStri(1).AppVar Array("        $stil='cave';", vbCrLf)
     iStri(1).AppVar Array("        $text = ""&Oslash "";", vbCrLf)
     iStri(1).AppVar Array("      } ELSE {", vbCrLf)
     iStri(1).AppVar Array("        $stil='unauff';", vbCrLf)
     iStri(1).AppVar Array("        $text="""";", vbCrLf)
     iStri(1).AppVar Array("      }", vbCrLf)
     iStri(1).AppVar Array("      $text=$text.""", ÜS(i), """;", vbCrLf)
     iStri(1).AppVar Array("      echo ""<button class='"".$stil.""' name='", ÜS(i), "'>"".$text.""</button>"";", vbCrLf)
     iStri(1).AppVar Array("    }", vbCrLf)
    End If
    If iob(2) Then iStri(2).AppVar Array(" $_SESSION['", ÜS(i), "']=1;", vbCrLf)
    If iob(3) Then iStri(3).AppVar Array(" $_SESSION['ob", ÜS(i), "']=0;", vbCrLf)
   End If
  Next i
   If iob(0) Then iStri(0).AppVar Array(" }", vbCrLf, "?>", vbCrLf)
   For i = 1 To 3
    If iob(i) Then iStri(i).AppVar Array("?>", vbCrLf)
   Next
   
   AusS.AppVar (Array(" include '../php/anzeig.php';", vbCrLf))
   AusS.AppVar (Array("?>", vbCrLf))
   AusS.AppVar (Array("", vbCrLf))
  End If ' If (obphp <> 0) Then
#If alt Then
  AusS.AppVar (Array("<a name='AnkerEing' href='#AnkerEing' accesskey='e'></a>", vbCrLf))
  If obphp <> 0 Then
   AusS.AppVar Array("<?php ", vbCrLf)
   AusS.AppVar Array(" include '../php/form.php';", vbCrLf)
   AusS.AppVar Array("?>", vbCrLf)
  End If
#End If
  Set rFlb = Nothing
' rFlb.Open "SELECT `dokumente`.pat_id, `dokumente`.zeitpunkt AS zp, `dokumente`.dokpfad, `dokumente`.dokname, `dokumente`.quelldatum, `dokumente`.dokgroe, kvnr, LEFT(IF(übwvbsnr='',IF(übwvkvnr = '', andüw, übwvkvnr),übwvbsnr),7) übwv, IF(`br_abgehakt`.abgehakt,'X',' ') AS abgeh FROM ((`dokumente` LEFT JOIN `br_abgehakt` ON `dokumente`.DokPfad = `br_abgehakt`.DokPfad) LEFT JOIN `namen` ON `dokumente`.pat_id = `namen`.pat_id) LEFT JOIN `lfaelle` ON `dokumente`.pat_id = `lfaelle`.pat_id WHERE (((`dokumente`.DokName) LIKE '%Fremdlabor%')) AND (ISNULL(abgehakt) OR abgehakt = 0) AND `dokumente`.pat_id = " & Pat_ID & " ORDER BY `dokumente`.pat_id DESC, zp DESC", DBCn, adOpenDynamic, adLockReadOnly
'  rFlb.Open "SELECT `dokumente`.pat_id, `dokumente`.zeitpunkt AS zp, `dokumente`.dokpfad, `dokumente`.dokname, `dokumente`.quelldatum, `dokumente`.dokgroe, kvnr, übwr, IF(`br_abgehakt`.abgehakt,'X',' ') AS abgeh FROM ((`dokumente` LEFT JOIN `br_abgehakt` ON `dokumente`.DokPfad = `br_abgehakt`.DokPfad) LEFT JOIN `namen` ON `dokumente`.pat_id = `namen`.pat_id) LEFT JOIN `lfaelle` ON `dokumente`.pat_id = `lfaelle`.pat_id WHERE (((`dokumente`.DokName) LIKE '%Fremdlabor%')) AND (ISNULL(abgehakt) OR abgehakt = 0) AND `dokumente`.pat_id = " & Pat_id & " ORDER BY `dokumente`.pat_id DESC, zp DESC", DBCn, adOpenDynamic, adLockReadOnly
  myFrag rFlb, "SELECT b.pat_id, b.zeitpunkt zp, b.pfad, b.name, b.quelldatum, b.dokgroe, kvnr, übwr, IF(ba.abgehakt,'X',' ') abgeh FROM ((tmbrie b LEFT JOIN `br_abgehakt` ba ON b.Pfad = ba.DokPfad) LEFT JOIN `namen` n ON b.pat_id = n.pat_id) LEFT JOIN `lfaelle` lf ON b.pat_id = lf.pat_id WHERE (((b.Name) LIKE '%Fremdlabor%')) AND (ISNULL(abgehakt) OR abgehakt = 0) AND b.pat_id = " & Pat_id & " ORDER BY b.pat_id DESC, zp DESC"
  If Not rFlb.BOF Then
   AusS.AppVar (Array("<table>", vbCrLf))
   AusS.AppVar (Array("  <tr>", vbCrLf))
   AusS.AppVar (Array("   <p class='abstand'> </p>", vbCrLf))
   AusS.AppVar (Array("   <th class='cave'>Fehlende Fremdlabore:</th>", vbCrLf))
   FlbZ = 0
   Do
    AusS.AppVar (Array("    <td><p>", Format(rFlb!Zp, "d.m.yy"), "</p></td>", vbCrLf))
    rFlb.MoveNext
    If rFlb.EOF Then Exit Do
    FlbZ = FlbZ + 1
    If FlbZ > 4 Then Exit Do
   Loop
   AusS.AppVar (Array("  </tr>", vbCrLf))
   AusS.AppVar (Array("</table>", vbCrLf))
  End If ' not rFlb.bof
#If False Then
  m = 27: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
' Liste der ganz fehlenden Parameter
  For i = 0 To pzl
   Dim obschonabstand%
   If rsz(i) = 0 Then
    If Not obschonabstand Then
     AusS.AppVar (Array("<p class='abstand'> </p>", vbCrLf))
     obschonabstand = True
    End If
    If nurpath(i) < 2 And Fqmin(i) <> 0 Then ' letzteres 15.11.08
     AusS.AppVar (Array("     <span class='cave'>&Oslash;&nbsp;", ÜS(i), ",&nbsp;&nbsp;</span>", vbCrLf))
    End If
   End If
  Next i
#End If
  m = 28: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  If obMB Then AusS.AppVar (Array("    <span class='cave'>&Oslash;&nbsp;", MerkblattText, ",&nbsp;&nbsp;</span>", vbCrLf))
'  IF obAvan THEN AusS.AppVar (Array("    <span class='cave'>&nbsp;Avandia,&nbsp;&nbsp;</span>", vbCrLf))
'  IF obAvan THEN AusS.AppVar (Array("    <span class='cave'>&nbsp;", wasAvan, ",&nbsp;&nbsp;</span>", vbCrLf))
  If falDiabDau Then AusS.AppVar (Array("    <span class='cave'>&nbsp;bitte `Diabetes seit` im Anamnesemakro überprüfen: soll Jahreszahl oder Datum sein,&nbsp;&nbsp;</span>", vbCrLf))
' Barthel-Index
  Dim rsdd As New ADODB.Recordset, rse As ADODB.Recordset
  myFrag rsdd, "SELECT icd FROM diagview pfld WHERE pat_id = " & Pat_id & " AND (pfld.gicd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.gicd RLIKE '^F0[012]')"
  If rFlSchGr <> 90 And (PAlter >= 70 Or Not rsdd.BOF) Then
   Set rse = Nothing ' 13.4.21: wenn im Folgenden nicht _latin1 bei inhalt steht, kommt bei der virtuallen Spalte Unsinn raus
   myFrag rse, "SELECT DATE(zeitpunkt) datum, art, inhalt FROM `eintraege` WHERE zeitpunkt > DATE_SUB(NOW(), INTERVAL 150 DAY) AND Art IN('ADL','284','bar') AND pat_id = " & Pat_id & " ORDER BY zeitpunkt DESC"
   If rse.BOF Then
    AusS.AppVar (Array("    <span class='cave'>&nbsp;ADL(Barthel),&nbsp;&nbsp;</span>", vbCrLf))
   End If
   Set rse = Nothing
   ' ktag fehlerhaft
   myFrag rse, "SELECT DATE(zeitpunkt) datum, art, inhalt FROM `eintraege` WHERE zeitpunkt > DATE_SUB(NOW(), INTERVAL 150 DAY) AND Art IN('TUG','247') AND pat_id = " & Pat_id & " ORDER BY zeitpunkt DESC"
   If rse.BOF Then
    AusS.AppVar (Array("    <span class='cave'>&nbsp;TUG(Time up AND go),&nbsp;&nbsp;</span>", vbCrLf))
   End If
  End If ' rFlSchGr <> 90 AND (PAlter >= 70 OR Not rsdd.BOF) THEN
  Set rsdd = Nothing
  m = 29: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  If obdm Then
   ' bei der Techniker nicht, wenn Pat. in der hausarztzentrierten Versorgung
   If rFlkrkasse Like "DAK*" Or rFlkrkasse Like "KKH*" Or rFlkrkasse Like "Kaufmännische K*" Or (rFlKateg = "EK" And (((rFlkrkasse Like "*TK*" Or rFlkrkasse Like "*Techniker*") And HzV <> 1) Or rFlkrkasse Like "*hans*")) Then
    Dim rDF As ADODB.Recordset
    Set rDF = Nothing
    myFrag rDF, "SELECT CASE substr(inhalt,6,2) WHEN 'HA' THEN 'HA' WHEN 'hi' THEN 'hier' WHEN 'ne' THEN 'nein' ELSE '?' END wo," & _
     "DATE_FORMAT(zeitpunkt,'%d.%m.%y'), e.* FROM eintraege e WHERE art='dak' " & _
     "and zeitpunkt= (SELECT MAX(zeitpunkt) FROM eintraege ei WHERE pat_id = e.pat_id AND art='dak') " & _
     "and pat_id = " & Pat_id & _
     " ORDER BY inhalt"
    AusS.AppVar Array("<p class='abstand'> </p>", vbCrLf)
    
    Dim DAKUStr$
    If rDF.BOF Then DAKUStr = "??" Else DAKUStr = rDF.Fields(0) & " " & rDF.Fields(1)
    DAKUStr = DAKUStr & ", offene bzw. mögliche Module: "
    AusS.AppVar Array("<span style= 'color:rgb(51, 153, 255);font-size:4mm'>&nbsp;DAK/KKH/HEK/TK-Unterschrift " & DAKUStr)
     
     Dim NPmd%, LUTSmd%, Apmd%, LebMd%, NiMd%
' #Const einzeln = 1
' #Const aufzwei = 1
#If einzeln Then
#If aufzwei Then
     NPmd = myEFrag("SELECT 1 FROM diagview WHERE pat_id = " & Pat_id & " AND gICD RLIKE '^E1[0-4]\.4|^G59\.0|^G63\.2|^G99\.0'").EOF
     If NPmd Then NPmd = myEFrag("SELECT zeitpunkt FROM eintraege WHERE art='daknp' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)").EOF
#Else
     NPmd = myEFrag("SELECT 1 FROM diagview WHERE pat_id = " & Pat_id & " AND gICD RLIKE '^E1[0-4]\.4|^G59\.0|^G63\.2|^G99\.0'" & vbCrLf & _
     "UNION" & vbCrLf & _
     "SELECT 1 FROM eintraege WHERE art='daknp' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)").EOF
#End If
#If aufzwei Then
     LUTSmd = myEFrag("SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gicd RLIKE '^N31\.[12]'").BOF
     If LUTSmd Then LUTSmd = myEFrag("SELECT zeitpunkt FROM eintraege WHERE art='dakluts' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)").EOF
#Else
     LUTSmd = myEFrag("SELECT 1 FROM diagview WHERE pat_id = " & Pat_id & " AND gicd RLIKE '^N31\.[12]'" & vbCrLf & _
     "UNION" & vbCrLf & _
     "SELECT 1 FROM eintraege WHERE art='dakluts' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)").EOF
#End If
     If PAlter > 50 Then
#If aufzwei Then
     Apmd = myEFrag("SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gicd RLIKE '^E1.\.5|^I79.2|^I73.'").BOF
     If Apmd Then Apmd = myEFrag("SELECT zeitpunkt FROM eintraege WHERE art='dakap' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)").EOF
#Else
      Apmd = myEFrag("SELECT 1 FROM diagview WHERE pat_id = " & Pat_id & " AND gicd RLIKE '^E1.\.5|^I79.2|^I73.'" & vbCrLf & _
      "UNION" & vbCrLf & _
      "SELECT 1 FROM eintraege WHERE art='dakap' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)").EOF
     Else ' PAlter > 50 Then
      Apmd = 0
#End If
     End If ' PAlter > 50 Then Else
#If aufzwei Then
     LebMd = myEFrag("SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gicd = 'K77.8'").BOF ' COALESCE(Dggel,0)=0 AND
     If LebMd Then LebMd = myEFrag("SELECT zeitpunkt FROM eintraege WHERE art='sono' AND (inhalt LIKE '%Abdomen%' OR inhalt LIKE '%Bauch%') AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)").EOF
#Else
     LebMd = myEFrag("SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gicd = 'K77.8'" & vbCrLf & _
     "UNION" & vbCrLf & _
     "SELECT zeitpunkt FROM eintraege WHERE art='sono' AND (inhalt LIKE '%Abdomen%' OR inhalt LIKE '%Bauch%') AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)").EOF
#End If
     NiMd = myEFrag("SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gicd RLIKE '^E1.*\.2|^N18|^N19|^I12\.0|^I13\.1|^I13\.2|^Z49\.[012]|^Z99\.2'").BOF ' COALESCE(Dggel,0)=0 AND
#Else
     sql0 = "SELECT" & vbCrLf & _
     "COALESCE((SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gICD RLIKE '^G59\.0|^G63\.2|^G99\.0'" & vbCrLf & _
     "UNION" & vbCrLf & _
     "SELECT 0 FROM eintraege WHERE art='daknp' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)LIMIT 1),-1)NPmd," & vbCrLf & _
     "COALESCE((SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gICD RLIKE '^N31\.[12]'" & vbCrLf & _
     "UNION" & vbCrLf & _
     "SELECT 0 FROM eintraege WHERE art='dakluts' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)LIMIT 1),-1)LUTSmd," & vbCrLf & _
     "(PATALTER(" & Pat_id & ")>50)*COALESCE((SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gICD RLIKE '^I79.2|^I73.'" & vbCrLf & _
     "UNION" & vbCrLf & _
     "SELECT 0 FROM eintraege WHERE art='dakap' AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)LIMIT 1),-1)APmd," & vbCrLf & _
     "COALESCE((SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gICD RLIKE '^K77.8'" & vbCrLf & _
     "UNION" & vbCrLf & _
     "SELECT 0 FROM eintraege WHERE art='sono' AND (inhalt LIKE '%Abdomen%' OR inhalt LIKE '%Bauch%') AND pat_id = " & Pat_id & " AND zeitpunkt >= SUBDATE(qanf(),INTERVAL 9 MONTH)LIMIT 1),-1)LebMd," & vbCrLf & _
     "COALESCE((SELECT 0 FROM diagview WHERE pat_id = " & Pat_id & " AND gicd RLIKE '^N1[89]|^I12\.0|^I13\.1|^I13\.2|^Z49\.[012]|^Z99\.2'LIMIT 1),-1)NiMD"
     Set rse = myEFrag(sql0, rAf)
     NPmd = rse!NPmd
     LUTSmd = rse!LUTSmd
     Apmd = rse!Apmd
     LebMd = rse!LebMd
     NiMd = rse!NiMd
     Set rse = Nothing
#End If
     If (NPmd Or LUTSmd Or Apmd Or LebMd Or NiMd) Then
      AusS.AppVar Array(IIf(NPmd, "NPMd, ", ""), IIf(LUTSmd, "LUTSMd, ", ""), IIf(Apmd, "APMd, ", ""), IIf(LebMd, "LebMd ,", ""), IIf(NiMd, "NiMd ", ""))
     End If ' NPmd OR LUTSmd OR Apmd OR Lebmd OR NiMd
     AusS.AppVar Array("</span>", vbCrLf)
   End If ' IF rFl!Kasse LIKE "DAK*" THEN
  End If ' obdm
  m = 30: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  
  ' Therapiehinweise
  Dim thh$(), metdos%, obsglt%, obamio%
  m = 31: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  Call mplan(CLng(Pat_id))
  m = 32: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  
  m = 33: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  If SafeArrayGetDim(mdpl) <> 0 Then
   For ru = 0 To UBound(mdpl)
    If mdpl(ru).metf Then
     metdos = metdos + (medzz(mdpl(ru).m.mo) + medzz(mdpl(ru).m.mi) + medzz(mdpl(ru).m.nm) + medzz(mdpl(ru).m.ab) + medzz(mdpl(ru).m.Zn)) * MachNumerisch(mdpl(ru).m.Medikament)
    End If
    If mdpl(ru).sglt Then obsglt = True
    If left$(mdpl(ru).m.MedAnfang, 5) = "CORDA" Or left$(mdpl(ru).m.MedAnfang, 6) = "AMIODA" Then obamio = True
   Next ru
  End If
'  rTh As Adodb.Recordset
'  SET rTh = myEFrag("SELECT (REPLACE(REPLACE(mo,'½',0.5),'1/2',0.5)+REPLACE(REPLACE(mi,'½',0.5),'1/2',0.5)+REPLACE(REPLACE(nm,'½',0.5),'1/2',0.5)+REPLACE(REPLACE(ab,'½',0.5),'1/2',0.5)+REPLACE(REPLACE(zn,'½',0.5),'1/2',0.5))*nr mg,i.* FROM " & _
           "(SELECT mp.medikament, zunr(mp.medikament) nr,mo,mi,nm,ab,zn FROM medplan mp " & _
           "LEFT JOIN medarten ma ON ma.medikament=mp.medanfang " & _
           "WHERE mp.pat_id = " & Pat_id & " AND mpnr=(SELECT MAX(mpnr) FROM medplan WHERE pat_id = " & Pat_id & ") " & _
           "and metf " & _
           ") i;")
'  IF Not rTh.BOF THEN metdos = rTh!mg
  Dim lGFR As labtyp
  lGFR = letztGFR(CLng(Pat_id), PAlter, obweibl, aRisk.EthnicGroup)
  If lGFR.Abkü = "" And lGFR.WertSg = "" Then lGFR.WertSg = "500"
  ' beim längsten Pat. 0,015 - 0,023 s
'  obmetf = myEFrag("SELECT COUNT(0) zl FROM medplan mp LEFT JOIN medarten ma ON mp.medanfang = ma.medikament WHERE mp.pat_id=" & Pat_ID & " AND metf")!zl
  ' 0s
  If obdm And dmtyp = "2" And therart <> "Diät" Then
' TH:Metformin
   If (IIf(Not IsNumeric(lGFR.WertSg), 200, lGFR.WertSg)) > 45 Then
    Dim obmetf%
    ' ' convert(mp.medanfang, char CHARSET utf8 collate utf8_german2_ci) = convert(ma.medikament, char CHARSET utf8 COLLATE utf8_german2_ci) " & vbCrLf & _

    obmetf = myEFrag("SELECT COALESCE(-(SELECT 1 FROM medplan mp LEFT JOIN medarten ma ON mp.medanfang = ma.medikament " & vbCrLf & _
     " WHERE mp.pat_id=" & Pat_id & " AND metf LIMIT 1),0) zl")!zl
    If Not obmetf Then
     If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
     thh(UBound(thh)) = "Typ 2-Diabetes, Therapie " & therart & ", lGFR " & lGFR.WertSg & ", bisher noch kein Metformin => evtl. versuchen"
    End If
   End If ' lGFR.WertSg > 45
   If IIf(Not IsNumeric(lGFR.WertSg), 200, lGFR.WertSg) > 59 And (obneph Or obherzi) Then
' TH:SGLT2-Hemmer
    Dim sglt%
    sglt = myEFrag("SELECT COALESCE(" & vbCrLf & _
     "-(SELECT 1 FROM medplan mp LEFT JOIN medarten ma ON mp.medanfang = ma.medikament " & vbCrLf & _
     "LEFT JOIN eintraege e ON e.pat_id=mp.pat_id AND art IN ('KSP','KSÄ')" & vbCrLf & _
     "WHERE ISNULL(e.id) AND mp.pat_id=" & Pat_id & " AND sglt2 " & vbCrLf & _
     "AND mp.zeitpunkt=(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=" & Pat_id & ") LIMIT 1) " & vbCrLf & _
     ",0) zl")!zl
    If Not sglt Then
     If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
     thh(UBound(thh)) = "Typ 2-Diabetes, Therapie " & therart & IIf(obneph, ", Nephropathie", "") & IIf(obherzi, ", Herzinsuffizienz", "") & ", letzte eGFR: " & lGFR.WertSg & "; zuletzt kein SGLT-2-Hemmer im Med'plan, kein ksp/ksä-Eintrag => evtl. versuchen"
    End If
   End If
   If Not obdement And PAlter < 80 And (obart Or (üdt.bmi > 26.9 And (therart = "CSII" Or therart = "CT" Or therart = "ICT" Or therart = "Komb"))) Then
' TH:GLP1-Agonist
    Dim glp%
    glp = myEFrag("SELECT COALESCE(" & vbCrLf & _
     "-(SELECT 1 FROM medplan mp LEFT JOIN medarten ma ON mp.medanfang = ma.medikament " & vbCrLf & _
     "LEFT JOIN eintraege e ON e.pat_id=mp.pat_id AND art IN ('KGP','KGÄ')" & vbCrLf & _
     "WHERE ISNULL(e.id) AND mp.pat_id=" & Pat_id & " AND glp1 " & vbCrLf & _
     "AND mp.zeitpunkt=(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=" & Pat_id & ") LIMIT 1) " & vbCrLf & _
     ",0) zl")!zl
    If Not glp Then
     If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
     thh(UBound(thh)) = "Typ 2-Diabetes, Therapie " & therart & IIf(obart, ", Arteriosklerose", ", BMI: " & Round(üdt.bmi, 1)) & ", zuletzt kein GLP-1-Analogon im Medikamentenplan, kein kgp/kgä-Eintrag => evtl. versuchen"
    End If ' not glp THEN
   End If ' obart OR (üdt.bmi
  End If ' obdm AND dmtyp = "2" AND therart <> "Diät" THEN
  
  If rFlSchGr <> 90 Then ' nicht bei Privaten
' TH:Kombinationstherapie DPP4-Hemmer aufspalten
   Dim rdpp4 As New ADODB.Recordset
   myFrag rdpp4, "SELECT IF(therart IN ('ICT','CSII','CT','GLP1Ins','GLP1ICT','GLP1'),CONCAT('Therapieart ',therart,', bitte ',mp.medikament,' ',IF(metf,'um','ab'),'setzen!'),CONCAT('Kombinationstablette ',mp.medikament,' bitte aufspalten!')) Anw FROM namen f " & vbCrLf & _
   "LEFT JOIN therarten th ON f.pat_id=th.pat_id AND mpnr=(SELECT mpnr FROM therarten WHERE pat_id=f.pat_id ORDER BY zp DESC,mpnr DESC LIMIT 1) " & vbCrLf & _
   "LEFT JOIN lmp mp ON f.pat_id = mp.pat_id " & vbCrLf & _
   "LEFT JOIN medarten ma ON mp.medanfang=ma.medikament " & vbCrLf & _
   "WHERE (therart IN ('ICT','CSII','CT','GLP1Ins','GLP1ICT','GLP1') OR metf) AND DPP4 AND f.pat_id=" & Pat_id
   If Not rdpp4.BOF Then
    If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
    thh(UBound(thh)) = rdpp4!Anw
   End If
' TH:Kombinationstherapie SGLT-2-Hemmer aufspalten
   Set rdpp4 = Nothing
   myFrag rdpp4, "SELECT f.pat_id, CONCAT(mp.medikament,' bitte ab- oder umsetzen!') Anw FROM namen f " & vbCrLf & _
   "LEFT JOIN therarten th ON f.pat_id=th.pat_id AND mpnr=(SELECT mpnr FROM therarten WHERE pat_id=f.pat_id ORDER BY zp DESC,mpnr DESC LIMIT 1) " & vbCrLf & _
   "LEFT JOIN lmp mp ON f.pat_id = mp.pat_id " & vbCrLf & _
   "WHERE mp.medikament LIKE '%glyz%' AND f.pat_id=" & Pat_id
   If Not rdpp4.BOF Then
    If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
    thh(UBound(thh)) = rdpp4!Anw
   End If
' TH:Kombinationstherapie Fettsenker aufspalten
   Set rdpp4 = Nothing
   myFrag rdpp4, "SELECT f.pat_id, CONCAT('Kombipräparat ',mp.medikament,' bitte aufspalten!') Anw FROM namen f " & vbCrLf & _
   "LEFT JOIN therarten th ON f.pat_id=th.pat_id AND mpnr=(SELECT mpnr FROM therarten WHERE pat_id=f.pat_id ORDER BY zp DESC,mpnr DESC LIMIT 1) " & vbCrLf & _
   "LEFT JOIN lmp mp ON f.pat_id = mp.pat_id " & vbCrLf & _
   "LEFT JOIN medarten ma ON mp.medanfang = ma.medikament " & vbCrLf & _
   "WHERE fetts AND hmg AND f.pat_id=" & Pat_id
   If Not rdpp4.BOF Then
    If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
    thh(UBound(thh)) = rdpp4!Anw
   End If
' TH:kürzlich aut-idem-Kreuz
   Set rdpp4 = Nothing
   myFrag rdpp4, "SELECT pat_id, CONCAT('cave: kürzlich aut-idem-Kreuz bei: ',GROUP_CONCAT(LEFT(medikament,instr(medikament,' ')) SEPARATOR ',')) Anw FROM rezepteintraege WHERE auti=0 AND rezklkurz<>'prp' AND zeitpunkt> SUBDATE(NOW(),90) AND pat_id=" & Pat_id, adOpenStatic
   If Not rdpp4.BOF And Not IsNull(rdpp4!Anw) Then
    If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
    thh(UBound(thh)) = rdpp4!Anw ' .fields(1)
   End If
' TH:Brilique 60 mg
   Set rdpp4 = Nothing
   myFrag rdpp4, "SELECT pat_id, CONCAT('cave: ',medikament,' darf von uns nicht verordnet werden (nur 90 mg), bitte diesbezüglich zum Kardiologen gehen!') Anw FROM rezepteintraege rz WHERE medikament LIKE 'Brilique 60%' AND rezklkurz<>'prp' AND zeitpunkt> SUBDATE(NOW(),180) AND pat_id=" & Pat_id & " LIMIT 1"
   If Not rdpp4.BOF Then
    If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
    thh(UBound(thh)) = rdpp4!Anw ' .fields(1)
   End If ' Not rdpp4.BOF THEN
   Set rdpp4 = Nothing
  End If ' IF rFlSchGr <> 90 THEN ' nicht bei Privaten

' TH:Metformin-Dosierung
  If IIf(Not IsNumeric(lGFR.WertSg), 200, lGFR.WertSg) < 30 And metdos > 0 Then
   If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
   thh(UBound(thh)) = "eGFR " & lGFR.WertSg & " " & lGFR.Einheit & " und Metformindosis " & metdos & " mg/d => evtl. absetzen"
  ElseIf IIf(Not IsNumeric(lGFR.WertSg), 200, lGFR.WertSg) < 45 And metdos > 1000 Then
   If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
   thh(UBound(thh)) = "eGFR " & lGFR.WertSg & " " & lGFR.Einheit & " und Metformindosis " & metdos & " mg/d => evtl. vermindern"
  End If ' lGFR.WertSg < 30 AND metdos > 0 THEN else
  If obsglt And obamio Then
   If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
   thh(UBound(thh)) = "SGLT-2-Hemmer sowie Amiodaron: im Fall einer " & IIf(obweibl, "Vaginitis", "Balanitis") & " kein ~conazol!"
  End If ' obsglt AND obamio THEN
  
' TH:HCT
  If obHCT Then
   Set rsMB = Nothing
   myFrag rsMB, "SELECT 0 FROM `eintraege` WHERE pat_id = " & Pat_id & " AND art = 'notiz' AND zeitpunkt > SUBDATE(NOW(), INTERVAL 730 DAY) AND inhalt LIKE 'Aufklärung über erhöhtes Risiko für Spinaliom und Basaliom unter HCT%'"
   If rsMB.BOF Then
    If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
    thh(UBound(thh)) = "HCT (" & hctMed & ") ohne Makro HCT, evtl. aufklären, evtl. umsetzen"
   End If
   Set rsMB = Nothing
  End If ' obHCT THEN
  
' TH:Blutdruck
  Const frist% = 215 ' 7 Monate, damit bei den halbjährlichen Patienten nicht gleich jedesmal eine Messlücke gemeldet wird
  Dim lRRÄnd As Date, lRRs%, lRRd%, lRRz%, RRMedAktZ%, RRtxt$
  myFrag rsMB, "SELECT SUM(rrsyst*rrzahl)/SUM(rrzahl) Rs, SUM(rrdiast*rrzahl)/SUM(rrzahl) Rd, SUM(rrzahl) Rz FROM rr WHERE pat_id=" & Pat_id & " AND zeitpunkt> SUBDATE(NOW()," & frist & ")"
' BOF kann hier nicht auftreten
  If Not IsNull(rsMB!rs) Then lRRs = rsMB!rs
  If Not IsNull(rsMB!rD) Then lRRd = rsMB!rD
  If Not IsNull(rsMB!rz) Then lRRz = rsMB!rz
  Set rsMB = Nothing
' DKG/ESC/DHL/ESH-Leitlinien 2018, wenn wir diesen Wert als Praxis-Blutdruck annehmen
  If lRRz = 0 Then
   RRtxt = "Keine RR-Messung in den letzten " & frist & " Tagen gespeichert. Bitte messen und eintragen."
  Else ' => lRRz > 0
   If (lRRd <= 80 And (PAlter < 65 And lRRs < 130 And lRRs >= 120) Or (PAlter >= 65 And lRRs < 140 And lRRs >= 130)) Then
' => auch ohne Kenntnis von Medikation sowie Diagnosen alles in Ordnung
   Else ' wenn nicht, dann brauchen wir die Medikation
    Dim obKHE%, ICDs
    ICDs = Array("I20", "I21", "I22", "I23", "I24", "I25", "Z95.5")
    obKHE = pruefICD(ICDs)
'    obKHE = diI("I20 I21 I22 I23 I24 I25 Z95.5", Pat_ID, 0, True)
    myFrag rsMB, "SELECT IF(COUNT(0)<>0,IF(MAX(mizp)>MAX(mazp),MAX(mizp),MAX(mazp)),0) lÄnd FROM (" & vbCrLf & _
    "SELECT mp.Medikament Med" & vbCrLf & _
    ", (SELECT MIN(zeitpunkt) FROM medplan WHERE pat_id=mp.pat_id AND medikament=mp.Medikament AND mo=mp.mo AND mi=mp.mi AND nm=mp.nm AND ab=mp.ab AND zn=mp.zn AND bbed=mp.bbed) mizp " & vbCrLf & _
    ", (IF((SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=mp.pat_id AND medikament=mp.Medikament AND mo=mp.mo AND mi=mp.mi AND nm=mp.nm AND ab=mp.ab AND zn=mp.zn AND bbed=mp.bbed)=(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=" & Pat_id & "),0,(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=mp.pat_id AND medikament=mp.Medikament AND mo=mp.mo AND mi=mp.mi AND nm=mp.nm AND ab=mp.ab AND zn=mp.zn AND bbed=mp.bbed))) mazp" & vbCrLf & _
    ", mp.mo, mp.pat_id Pat_ID" & vbCrLf & _
    "FROM medplan mp LEFT JOIN medarten ma ON mp.medanfang = ma.medikament " & vbCrLf & _
    "WHERE hypt" & vbCrLf & _
    "GROUP BY mp.pat_id, mp.medikament,mo,mi,nm,ab,zn,bBed" & vbCrLf & _
    ") i WHERE pat_id=" & Pat_id
    lRRÄnd = rsMB!länd
    RRMedAktZ = myEFrag("SELECT COALESCE(" & vbCrLf & _
     "-(SELECT 1 FROM medplan mp " & vbCrLf & _
     "LEFT JOIN medarten ma ON mp.medanfang = ma.medikament " & vbCrLf & _
     "WHERE mp.pat_id=" & Pat_id & " AND hypt AND mp.zeitpunkt=(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=" & Pat_id & ") LIMIT 1) " & vbCrLf & _
     ",0) zl")!zl
    Set rsMB = Nothing
' wenn die letzte Änderung jünger als frist, dann Durchnitt neu berechnen
    If lRRÄnd > Now() - frist Then
     myFrag rsMB, "SELECT SUM(rrsyst*rrzahl)/SUM(rrzahl) Rs, SUM(rrdiast*rrzahl)/SUM(rrzahl) Rd, SUM(rrzahl) Rz FROM rr WHERE pat_id=" & Pat_id & " AND zeitpunkt> " & Format(lRRÄnd, "yyyymmdd")
' BOF kann hier nicht auftreten
     If Not IsNull(rsMB!rs) Then lRRs = rsMB!rs Else lRRs = 0
     If Not IsNull(rsMB!rD) Then lRRd = rsMB!rD Else lRRd = 0
     If Not IsNull(rsMB!rz) Then lRRz = rsMB!rz Else lRRz = 0
     Set rsMB = Nothing
    End If
    If (lRRz = 0 Or (lRRd <= 80 And (PAlter < 65 And lRRs < 130 And lRRs >= 120) Or (PAlter >= 65 And lRRs < 140 And lRRs >= 130))) Then
' Blutdruck nach letzter Änderung nicht wieder gemessen oder in Ordnung
    ElseIf RRMedAktZ = 0 Then ' keine antihypertensive Medikation
     If (PAlter < 65 And lRRs < 120) Or (PAlter >= 65 And lRRs < 130) Then
     ' Blutdruck ohne Medikation niedrig => kein Handlungsbedarf
     Else
     ' Blutdruck ohne Medikation zu hoch
      RRtxt = "Durchschnitt der letzten " & CStr(lRRz) & " Blutdruckwerte mit " & CStr(lRRs) & "/" & CStr(lRRd) & " höher als " & IIf(PAlter < 65, "120", "130") & "/80 mm Hg, evtl. Antihypertensiva zu erwägen."
     End If
    Else ' RRMedz>0: antihypertensive Medikation, RR evtl. außerhalb Zielbereich
     If (PAlter < 65 And lRRs < 120) Or (PAlter >= 65 And lRRs < 130) Then
      If lRRÄnd < Now() - 14 Then ' wenn nicht kürzlich die Medikation schon geändert
       RRtxt = "RRsyst mit " & lRRs & " evtl. zu niedrig (<" & IIf(PAlter < 65, "120", "130") & ") mm Hg. Letzte Änderung der antihypert.Med. " & Format(lRRÄnd, "dd.mm.yy") & ", evtl. Med. reduzieren."
      End If
     Else
     ' RR evtl. zu hoch
      RRtxt = "Durchschnitt der letzten " & CStr(lRRz) & " Blutdruckwerte mit " & CStr(lRRs) & "/" & CStr(lRRd) & " höher als " & IIf(PAlter < 65, "130", "140") & "/80 mm Hg. Letzte Änderung der antihypert.Med. " & Format(lRRÄnd, "dd.mm.yy") & ", evtl. Änderung angezeigt."
     End If ' (Palter < 65 AND lRRs < 120) OR (Palter >= 65 AND lRRs < 130) THEN else
    End If ' RRMedz=0 else
   End If ' (lRRd < 80 AND (Palter < 65 AND lRRs < 130 AND lRRs >= 120) OR (Palter >= 65 AND lRRs < 140 AND lRRs >= 130)) THEN else
  End If ' lRRz = 0 THEN else
  If RRtxt <> "" Then
   If SafeArrayGetDim(thh) = 0 Then ReDim thh(0) Else ReDim Preserve thh(UBound(thh) + 1)
   thh(UBound(thh)) = RRtxt
  End If ' RRtxt <> "" THEN
  
  ' Gesamtausgabe der Therapiehinweise
  If SafeArrayGetDim(thh) Then
   AusS.AppVar Array("<p class='abstand'> </p>", vbCrLf)
   AusS.AppVar Array("<span style= 'background-color:rgb(34, 139, 34);color:rgb(255,250,240);font-size:4mm'>&nbsp;Therapiehinweise: ")
   AusS.AppVar Array("<span style= 'font-size:4mm'>")
   AusS.AppVar Array(" " & thh(0), "<br>", vbCrLf)
   If UBound(thh) > 0 Then
    Const tabu$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Dim ith%
    For ith = 1 To UBound(thh)
     AusS.AppVar Array(tabu, thh(ith), "<br>", vbCrLf)
    Next ith
    AusS.AppVar Array("</span></ol>", vbCrLf)
   End If ' UBound(thh) > 1 THEN
   AusS.AppVar Array("</span>", vbCrLf)
  End If ' SafeArrayGetDim(thh) THEN
  
' Ende Therapiehinweise
  
  iend = -1
  k = 0
  AusS.AppVar (Array("<a name='AnkerKG' href='#AnkerKG' accesskey='g'></a>", vbCrLf))
  m = 34: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  Do ' Blöcke
   ianf = iend + 1
   iend = iend + SpZahl
   If iend > pzl Then iend = pzl
   AusS.AppVar (Array("<p class='abstand'> </p>", vbCrLf))
   AusS.AppVar (Array("<TABLE>", vbCrLf))
'Überschriften
   AusS.AppVar (Array("   <TR>", vbCrLf))
   i = ianf
'  IF üdt.fußst = auff OR üdt.sens = auff OR üdt.sens = pathdok OR üdt.Puls = auff OR üdt.FEn(12) = True THEN
   Do
    If i > iend Then Exit Do
'    If ((i <> SchulEintri Or Fqmin(Schuli) = 365) And ((nurpath(i) = 0 Or nurpath(i) = 3 Or pathol(i, 0)) And rsz(i) <> 0 And Titel(i) <> "")) Then
     Dim Schulzufuss%
     Schulzufuss = 0
     If Schuli <> -1 Then
      If i <> SchulEintri Or Fqmin(Schuli) = 365 Then Schulzufuss = True
     End If ' Schuli <> -1 Then
     If (Schulzufuss And ((nurpath(i) = 0 Or nurpath(i) = 3 Or pathol(i, 0)) And rsz(i) <> 0 And Titel(i) <> "")) Then
      Dim Fußstr$
      If i = Fußi Then
       Fußstr = ""
       If Fqmin(Fußi) > 1 Then
        If obFußmakro And Fqminu > 1 Then Fußstr = Fußstr + "Fußmakro"
        If üdt.Infekt = 1 Then
         If Fußstr <> "" Then Fußstr = Fußstr + ","
         Fußstr = Fußstr + "Wundinf"
        End If
        If üdt.fußst = auff Then
         If Fußstr <> "" Then Fußstr = Fußstr + ","
         Fußstr = Fußstr + "Fußst"
        End If
        If üdt.sens = auff Or üdt.sens = pathdok Then
         If Fußstr <> "" Then Fußstr = Fußstr + ","
         Fußstr = Fußstr + "Sens"
        End If
        If üdt.Puls = auff Then
         If Fußstr <> "" Then Fußstr = Fußstr + ","
         Fußstr = Fußstr + "Puls"
        End If
        If üdt.FEn(12) Then
         If Fußstr <> "" Then Fußstr = Fußstr + ","
         Fußstr = Fußstr + "ICD Fuß"
        End If
       End If ' Fqmin(Fußi) > 1
       If Fußstr <> "" Then Fußstr = " (" & Fußstr & ")"
      End If ' i = Fußi
      ' Begründungen für Intervalle
      AusS.AppVar (Array("     <th title='", Titel(i), "' colspan='2'", IIf(pathol(i, 0) Or (i = Gwi And üdt.bmi > 24.9), " class='path'", " class='schmal'"), ">", ÜS(i), _
      IIf(i = RRi, " (Ziel: &le;" & grenze(i) & "/" & GrenzeDiast & ")", ""), _
      IIf(i = Neuri, " (" & IIf(üdt.fußst = auff, "Fußst,", "") & IIf(üdt.sens = auff, "Sens,", "") & IIf(üdt.sens = pathdok, "ICD G63.2,", "") & IIf(üdt.Puls = auff, "Puls,", "") & IIf(üdt.FEn(12), "<span class='cave'>DFS</span>", "") & "&rArr;" & Fqmin(i) & "/a)", "")))
      If i = Fußi Then AusS.Append Fußstr
      AusS.AppVar (Array(IIf(i = Gwi, " (Gr " & üdt.kGR & ",<br>BMI " & Round(üdt.bmi, 1) & " kg/m²)", ""), _
      IIf(i = Schuli, " (Soll:" & sollz & IIf(obnach, "+", "") & ")", ""), "</th>", vbCrLf))
    Else
     If iend < pzl Then iend = iend + 1
    End If
    i = i + 1
   Loop ' Do
   
' folgendes noch nicht nötig, da nicht vorher verwendet:
'   For i = 0 To pzl
'    aktlz(i) = 0
'    angef(i)=0
'   Next i

' Tabelle der letzten (8) Ausprägungen jedes diabetologisch relevanten Parameters
   AusS.AppVar (Array("   </TR>", vbCrLf))
   For j = 1 To rszmax ' Durchlauf der Kleinzeilen innerhalb eines Blocks
'Werte(j)
    AusS.AppVar (Array("   <TR>", vbCrLf))
    For i = ianf To iend ' Parameter in jedem Block
'     IF i = GFRi THEN Stop
'     IF i = Hba1ci THEN Stop
'     If i = 12 Then Stop ' Kalium
     Schulzufuss = 0
     If Schuli <> -1 Then
      If i <> SchulEintri Or Fqmin(Schuli) = 365 Then Schulzufuss = True
     End If ' Schuli <> -1 Then
     If (Schulzufuss And ((nurpath(i) = 0 Or nurpath(i) = 3 Or pathol(i, 0)) And rsz(i) <> 0 And Titel(i) <> "")) Then
       If j <= rsz(i) * IIf(RowSp(i) = 0, 1, RowSp(i)) Then
        obpath(i) = 0
        Dim Zp As Date, zpu As Date
        If sql(i) = vNS Then
         gefunden = 0
         Do While True 'sollte am Ende fertig werden; gefunden hier nicht nötig da rsz(i)=0 schon ausgeschlossen
          If angef(i) Then aktlz(i) = aktlz(i) + 1
          angef(i) = True
          If obLabI(kritnr(i), lab(aktlz(i))) Then
           gefunden = True
           Exit Do
          End If
         Loop
         If gefunden Then
          zpu = lab(aktlz(i)).Zp
          Wert = lab(aktlz(i)).WertSg
         End If
        Else ' sql(i) = vNS Then
         If rs(i).EOF Then
          gefunden = 0
         Else
          gefunden = True
'          zpu = rs(i)!Zp
          If Not IsNull(rs(i)!Zp) Then zpu = rs(i)!Zp
          Wert = rs(i)!Wert
         End If ' rs(i).EOF Then else
        End If ' sql(i) = vNS Then else
        If gefunden And Not Wert Like "[<>]*" And rund(i) <> -1 Then ' 30.11.25: Albuminurie
         Wert = Round(MachNumerisch(Wert), rund(i))
        End If
        Zp = fctQAnf(ZQuart(zpu))
        If j = 1 Then
' wenn z.B. nicht in diesem Quartal (fqmin(i)=4) und auch nicht jünger als die Hälfte des entsprechenden Abstandes, dann rot
         If Fqmin(i) <> 0 Then
'          IF Now() - Zp + 30 > 365 / Fqmin(i) AND (Now - zpu) > 365 / Fqmin(i) * 0.5 THEN obpath(i) = True
          If Now() - Zp > 365 / Fqmin(i) And (Now - zpu) > 365 / Fqmin(i) * 0.5 Then obpath(i) = True
          If obpath(i) Then If i = Fußi Then If obpath(Neuri) Then obpath(Fußi) = False
         End If ' Fqmin(i) <> 0 THEN
        End If ' j = 1 THEN
        If Pause(i) = 0 Then
         Dim rsidat As Date
         
         If gefunden Then rsidat = zpu Else rsidat = CDate(0)
         AusS.AppVar (Array("    <td", IIf(RowSp(i) <> 0, " rowspan='" & MINvb(RowSp(i), rszmax - j + 1) & "' ", ""), " class='schmal'><P", IIf(obpath(i), " class='cave'", ""), ">", Format(rsidat, "d.m.yy"), "</td>", vbCrLf))
'         Debug.Print rs(i)!Wert
         Dim rsiwert
         On Error Resume Next
         If gefunden Then rsiwert = Wert Else rsiwert = 0
         On Error GoTo fehler
         AusS.AppVar (Array("    <td", IIf(RowSp(i) <> 0, " rowspan='" & MINvb(RowSp(i), rszmax - j + 1) & "' ", ""), IIf(pathol(i, j), " class='path'", ""), IIf(LenB(weite(i)) <> 0, " style='width:" & weite(i) & "'", ""), ">", IIf(RowSp(i) <> 0, vNS, "<p>"), rsiwert, IIf(RowSp(i) <> 0, vNS, "</p>"), "</td>", vbCrLf))
         If RowSp(i) <> 0 Then
          Pause(i) = RowSp(i) - 1
         End If
         If sql(i) <> vNS Then If rs(i).State <> 0 Then If Not rs(i).EOF Then rs(i).MoveNext
        Else
         Pause(i) = Pause(i) - 1
        End If
#Const mitLeerZellenMo = False
#If mitLeerZellenMo Then
       ElseIf Pause(i) = 0 Then
'        IF RowSp(i) = 0 OR j = 1 THEN ' Beim Neurostatus in die folgenden Zeilen nichts mehr schreiben
         AusS.AppVar (Array("    <td rowspan='", rszmax - j + 1, "' colspan = '2' class='schmal' style='border-style:none'></td>", vbCrLf))
         Pause(i) = -1
       ElseIf Pause(i) > 0 Then
         Pause(i) = Pause(i) - 1
#Else
      Else  ' rsz(i) >= j THEN
       If rsz(i) <> 0 Then
'        IF i <> Neuri OR j = 1 THEN ' Beim Neurostatus in die folgenden Zeilen nichts mehr schreiben
         AusS.AppVar (Array("    <td class='schmal'></td><td></td>", vbCrLf))
'        END IF ' rsz(i) <> 0 THEN
       End If
#End If
       End If ' rsz(i) >= j THEN
      End If ' nurpath(i) = 0 OR pathol(i, 0) THEN
    Next i
    AusS.AppVar (Array("   </TR>", vbCrLf))
    If j = rszmax Then
     If i = pzl Then
      Exit For ' fertig
     End If
    End If
   Next j
   AusS.AppVar (Array("</TABLE>", vbCrLf))
   If iend = pzl Then Exit Do
   k = k + 1
  Loop
  m = 35: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
' Diagnosen
'  AusS.AppVar (Array("<p class='abstand'> </p>", vbCrLf))
'  AusS.AppVar (Array("<div class='abstand'>", replace$(replace$(DiagString(CLng(Pat_ID)), vbVerticalTab & " ", "<br />&nbsp;"), vbVerticalTab, "<br />"), "</div>", vbCrLf))
  Dim TabZ&, obDiagnosen%, obmed%
  obDiagnosen = (SafeArrayGetDim(DiagTab) <> 0)
  obmed = (SafeArrayGetDim(mdpl) <> 0)
  If obDiagnosen Then TabZ = UBound(DiagTab)
  If obmed Then TabZ = MAXvb(TabZ, UBound(mdpl) + 2)
  If Not obDiagnosen And j <> -1 Then TabZ = j + 1
  ' UKPDS-Risk, s.o.
  If obdm Then TabZ = MAXvb(TabZ, 2)

' Termine eintragen
  Dim rTerm As New ADODB.Recordset
  Set rTerm = Nothing
  ' myFrag rTerm, "SELECT COUNT(0)-1 zahl FROM termine t WHERE zp >= DATE(NOW()) AND pid = " & Pat_id
  ' TabZ = MAX(TabZ, rTerm!Zahl)
  ' SET rTerm = Nothing
  myFrag rTerm, "SELECT COUNT(0) OVER()-1 zahl, DATE(zp) datum, zp uhrzeit, raum, zusatz FROM termine t WHERE zp >= DATE(NOW()) AND pid = " & Pat_id
  If Not rTerm.BOF Then TabZ = MAXvb(TabZ, rTerm!Zahl)
  
' Fettleberindices
  If NFS <> -1 Then TabZ = MAXvb(TabZ, 6) Else If FIB4 <> -1 Then TabZ = MAXvb(TabZ, 5) Else If FLI <> -1 Then TabZ = MAXvb(TabZ, 4)
  
' Diagnosen, Medikation und UKPDS Risk
  AusS.AppVar Array("<br><table border=""1""><thead align=""left""><tr><th>Diagnosen:</th><th>ICD</th><th bgcolor=""#CCCCCC"">_</th><th>Letzte Medikation:</th><th>fr</th><th>mi</th><th>nm</th><th>ab</th><th>zn</th><th>bBed</th><th bgcolor=""#CCCCCC"">_</th><th" & IIf(obdm, " bgcolor=""#FFFF00""", "") & ">", IIf(obdm, "<span title='" & UKtip & "'</span>UKPDS RE: </th><th>KHE</th><th>fatale KHE</th><th>Apoplex</th><th>fataler Apoplex</th><th bgcolor=""#CCCCCC"">_</th>", IIf(FLI <> -1 Or FIB4 <> -1 Or NFS <> -1, "</th><th></th><th></th><th></th><th></th><th></th>", "")), "<th>Termine</th>", vbCrLf)
  Dim medfertig%
  If obDiagnosen Or obmed Then
   For k = 0 To TabZ
    AusS.AppVar Array(" <tr>", vbCrLf)
    If obDiagnosen Then
     If k <= UBound(DiagTab) Then
      If Not DiagTab(k) Is Nothing Then
       Call DiagTab(k).REPLACE(vbTab, "</th><th class='diag'>")
       AusS.AppVar Array("  <th class='diag'>", DiagTab(k).Value, "</th>", vbCrLf)
      End If
     Else
      AusS.AppVar Array("  <th></th><th></th>", vbCrLf)
     End If
    Else ' obDiagnosen Then
     AusS.AppVar Array("  <th></th><th></th>", vbCrLf)
    End If ' obDiagnosen Then Else
    AusS.AppVar Array("  <th bgcolor=""#DDDDDD""></th>", vbCrLf)
    
    If obmed Then
     If k <= UBound(mdpl) Then
      AusS.AppVar Array("  <th class='med'>", mdpl(k).m.Medikament, "</th><th class='med'>", mdpl(k).m.mo, "</th><th class='med'>", mdpl(k).m.mi, "</th><th class='med'>", mdpl(k).m.nm, "</th><th class='med'>", mdpl(k).m.ab, "</th><th class='med'>", mdpl(k).m.Zn, "</th><th class='med'>", IIf(mdpl(k).m.bBed <> 0, "b.Bed.", vNS), "</th></span>", vbCrLf)
     ElseIf Not medfertig Then
      AusS.Append "  <th class='med' colspan=""7"" rowspan=""" & TabZ - k + 1 & """>"
      AusS.Append IIf(left$(mdpl(UBound(mdpl)).m.Bemerkung, 2) = vbCrLf, REPLACE$(Mid$(mdpl(UBound(mdpl)).m.Bemerkung, 2), vbCrLf, "<br>"), mdpl(UBound(mdpl)).m.Bemerkung)
      AusS.AppVar Array("</th>", vbCrLf)
      medfertig = True
     End If ' k <= UBound(mdpl) Then elseif
    End If ' obmed Then
    AusS.AppVar Array("  <th bgcolor=""#DDDDDD""></th>", vbCrLf)
    
    If obdm Or FLI <> -1 Or FIB4 <> -1 Or NFS <> -1 Then
     If (k < 3 And Not obdm) Or k = 3 Or (k = 4 And FLI = -1) Or (k = 5 And FIB4 = -1) Or (k = 6 And NFS = -1) Or k > 6 Then
      AusS.AppVar Array("  <th></th><th></th><th></th></th><th></th><th></th>")
     Else
      Select Case k
       Case 0
        If obdm Then AusS.AppVar Array("  <th>10a Risiko</th><th>", cR, "%</th><th>", CFR, "%</th><th>", SR, "%</th><th>", SFR, "%</th>")
       Case 1
        If obdm Then AusS.AppVar Array("  <th><span title='Vertrauensbereich-Untergrenze'>10a VB von</span></th><th>", CL, "%</th><th>", CFL, "%</th><th>", SL, "%</th><th>", SFL, "%</th>")
       Case 2
        If obdm Then AusS.AppVar Array("  <th><span title='Vertrauensbereich-Obergrenze'>10a VB bis</span></th><th>", CU, "%</th><th>", CFU, "%</th><th>", SU, "%</th><th>", SFU, "%</th>")
       Case Else
        If k = 4 And FLI <> -1 Then
         AusS.AppVar Array("  <th>Fettlb'ind.:</th><th></th><th", IIf(FLI < 30, "", " bgcolor=""" & IIf(FLI > 60, "yellow", "lightyellow") & """"), ">", FLI, "</th></th><th></th><th></th>")
        ElseIf k = 5 And FIB4 <> -1 Then
         AusS.AppVar Array("  <th>FIB-4'ind.:</th><th></th><th", IIf(FIB4 < IIf(PAlter < 66, 1.3, 2), "", " bgcolor=""" & IIf(FIB4 > 3.25, "yellow", "lightyellow") & """"), ">", CStr(FIB4), "</th></th><th></th><th></th>")
        ElseIf k = 6 And NFS <> -1 Then
         AusS.AppVar Array("  <th>NFScore:</th><th></th><th", IIf(NFS < -1.455, "", " bgcolor=""" & IIf(NFS > 0.676, "yellow", "lightyellow") & """"), ">", CStr(NFS), "</th></th><th></th><th></th>")
        End If
      End Select
     End If ' (k < 3 AND NOT obdm) OR (k = 4 And FLI = -1) OR (k = 5 And FIB4 = -1) OR (k = 6 And NFS = -1) OR k > 6 Then else
    End If ' obdm OR FLI <> -1 OR FIB4 <> -1 OR NFS <> -1 Then
    AusS.AppVar Array("  <th bgcolor=""#DDDDDD""></th>", vbCrLf)
    
    If Not rTerm.BOF And Not rTerm.EOF Then
     AusS.AppVar Array(" <th class='term'>", Format(rTerm!Uhrzeit, "d.m.yy hh:mm") & " " & rTerm!raum & " " & rTerm!zusatz & "</th>")
    End If
    If Not rTerm.BOF And Not rTerm.EOF Then rTerm.MoveNext
    
    AusS.AppVar Array(" </tr>", vbCrLf)
   Next k
  End If ' obDiagnosen OR obmed Then
  AusS.AppVar Array("</table><br>", vbCrLf)
  
' Hier kommt das Labor rein!
  Dim SelbstStatus%, raDatBOF%, Matr$() ' Matrix: 3. Dimension: 0 = Wert, 1 = Tip-Tool, 28.3.10
  Dim MForm%(), mBreiten$(), LDLZiel%
  LDLZiel = grenze(LDLi)
  Call LaborInsPLZ(Pat_id, SelbstStatus, raDatBOF, Matr, MForm, mBreiten, , LDLZiel)
  m = 37: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
'  Open "v:\matr.txt" For Append AS #336
'  For j = 0 To UBound(Matr, 3)
'   For i = 0 To UBound(Matr, 2)
'    Print #336, "i: " + Str$(i) + ", j: " + Str$(j) + " '" + Matr(0, i, j) + "'"
'   Next i
'  Next j
'  Close #336
  If SafeArrayGetDim(Matr) Then
   AusS.AppVar (Array("<p class='abstand'> </p>", vbCrLf))
   AusS.AppVar (Array("<Table Labor>", vbCrLf))
   AusS.AppVar (Array("<a name='AnkerLab' href='#AnkerLab' accesskey='q'></a>", vbCrLf))
   For j = 0 To UBound(Matr, 3)
    For i = 6 To UBound(Matr, 2)
     If (MForm(i, j) And 2) = 2 Then
      MForm(3, j) = 2 ' Parameterspalte mit pathologischem Wert fett auf gelbem Hintergrund
      Exit For
     End If
    Next i
    AusS.AppVar (Array(" <tr>", vbCrLf))
    Dim iumw& ' um Laborwerte in umgekehrter Datumsreihenfolge anzuzeigen
    For i = 3 To UBound(Matr, 2)
     If i > 5 Then
      iumw = UBound(Matr, 2) - i + 6
     Else
      iumw = i
     End If
     Dim Formg$ ' Formatierung / Farbe
     Select Case MForm(iumw, j)
      Case 0: Formg = "lnn" ' Labor normal eigen
      Case 1: Formg = "lno" ' Labor normal importiert
      Case 2: Formg = "lpn" ' Labor pathologisch eigen
      Case 3: Formg = "lpo" ' Labor pathologisch importiert
     End Select
     If i = 3 Then
      AusS.AppVar Array("  <td title=""", Matr(1, 0, j), """") ' ", IIf(j = 0, "h", "d"), "
     Else
      If LenB(Matr(1, iumw, j)) <> 0 Then ' Kommentar
       AusS.AppVar Array("  <td title=""", Matr(1, iumw, j), """") ' ", IIf(j = 0, "h", "d"), "
      Else
       AusS.AppVar Array("  <td") ' ", IIf(j = 0, "h", "d"), "
      End If
     End If
     AusS.AppVar Array(" class='", Formg, "'", Switch(j = 0 And iumw = 3, " style='min-width:12em'", j = 0 And iumw = 4, " style='min-width:6em'", True, ""), ">", Matr(0, iumw, j), "</td>", vbCrLf) ' ", IIf(j = 0, "h", "d"), "
    Next i
    AusS.AppVar (Array(" </tr>", vbCrLf))
   Next j
   AusS.AppVar (Array("</Table>", vbCrLf))
  End If ' SafeArrayGetDim
  m = 38: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  Call plzeintraege(Pat_id, AusS)
  
  On Error Resume Next
  For i = 0 To 100
   Err.Clear
   DateiNameAktRoh = DateiNameRoh & String(i, "_") & ".html"
   DateiNameAkt = IIf(obphp, "\\" & LiName & "\php\plz\", plzVerz) & DateiNameAktRoh ' replace$(CDate(Uhrzeit), ":", ".") & " " & Format(Datum, "dd.mm.yyyy") &
   If obphp <> 0 Then
    DatNr = UTFOpen(DateiNameAkt)
   Else
    Open DateiNameAkt For Output As #355
   End If
   If Err.Number = 0 Then Exit For
   If i = 100 Then MsgBox "Fehler beim Erstellen der Datei " & DateiNameAkt
  Next i
  On Error GoTo fehler
  m = 39: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  If obphp <> 0 Then
   sUTF DatNr, AusS.Value
   Close DatNr
  Else
   Print #355, AusS.Value
   Close #355
  End If
  Call SetFileTimeByDate(DateiNameAkt, mftlastwritetime, Datum + Uhrzeit)
  If obphp Then
   For i = 0 To 3
    If iob(i) Then
     DatNr = UTFOpen(iDN(i))
     sUTF DatNr, iStri(i).Value
     Close DatNr
    End If
   Next i
  End If
  m = 40: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
'  DoEvents
  Lese.Ausgeb "Patientenlaufzettel für " & NName & ", " & VName & " (Pat_id: " & Pat_id & ") erstellt.", True
  syscmd 4, "Fertig mit dem Patientenlaufzettel für: " & NName & ", " & VName & " (Pat_id: " & Pat_id & ")"
'  Call Shell("firefox.exe " & DateiName)
'  Dim reg0$, reg1$
 ' reg0 = getReg(HLM, "software\mozilla\Mozilla Firefox", "CurrentVersion")
  If mitanzeig Then
   If obphp <> 0 Then
    DateiName = "http://" & LiName & "/plz/" & DateiNameAktRoh
   Else
    DateiName = DateiNameAkt
   End If
   '' das geht nicht in sono1
   Dim DPfad$
   Dim dPFehler%
   DPfad = getReg(HLM, "software\mozilla\Mozilla Firefox\" & getReg(HLM, "software\mozilla\Mozilla Firefox", "CurrentVersion") & "\Main", "PathToExe") & " """ & DateiName & """"
   On Error Resume Next
   Shell DPfad
   dPFehler = Err.Number
   On Error GoTo fehler
   If dPFehler <> 0 Then
    rufauf "firefox.exe", DateiName
   End If
  End If
 m = 41: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = 0 Then
End If
Static runden%
If Err.Number = 429 Then
 runden = runden + 1
 If runden < 10 Then
   Dim QDat$, ZDat$
  If aktOSV >= win_vista Then
   QDat = uVerz & "Programmierung\riskeng.ocx"
   ZDat = Environ("programdata")
   Shell ("cmd /e:on /c copy " & Chr(34) & QDat & Chr(34) & " " & Chr(34) & ZDat & Chr(34))
   If AdminPwd = "" Then AdminPwd = holap("Administrator")
   Shell (vVerz & doalsAd & acceu & AdminPwd & " cmd /e:on /c regsvr32 " & Chr$(34) & ZDat & "\riskeng.ocx" & Chr$(34))
  Else
   Shell ("cmd /c regsvr32 """ & uVerz & "\programmierung\riskeng.ocx""")
  End If
  Sleep 1000
  Resume
 End If ' runden < 10 Then
End If ' Err.Number = 429 Then
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodoPLZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodoPLZ

' wegen Prozedurgröße zwangsweise aufgeteilt
' in dodoplz
Function tabelleInplz(Pat_id$, zmaxges%, rFlSchGr%)
 Dim i%, j&, WertNum!, gefunden%, Wert$
 Dim gef%(pzl)
 Dim normoton%, angefx&
 Dim pathAlbZ%, AlbZ%, pathKreaZ%, KreaZ%
 Dim obNP% ' ob Nephropathie vorliegt
 Dim rsDia As New ADODB.Recordset
 Dim UGrenze!(pzl)
 Dim pathz%(pzl), npathz%(pzl) ' Zahl der nicht/ Pathologischen
 Dim Caroi%, Caroai% ' Carotiduplex durchgeführt, alle
 Dim Hbi%, Ferri%, B12i%, Foli% ' Parameter für Hb, Ferritin, Vit B12, Folsre
 Dim Leuki%, CRPi%, TSHi%, fT4i%, fT3i%, TGi%, GGTi%, GPTi%
 Dim Kreai%, GFRi% ' Parameter für Kreatinin / GFR
 Dim Urini% ' Parameter für Urin
 Dim RRVgli% ' Parameter für RR-Vergleich
 Dim Tailli% ' Parameter für Taille
 Dim Hba1ci% ' Parameter für HbA1c (nur zum Debuggen)
 Dim Albi% ' Parameter für Albumin
 Dim loopct% ' , pathol%(pzl, zmaxges) ' pathol(pzl,0)= ' ob irgend einer der Werte pathologisch ist
 Dim rszahl As New ADODB.Recordset
 Dim zmax%(pzl) ' maximal dargestellte Zahl an Ausprägungen jedes Parameters
 On Error GoTo fehler

#If Not langsamer Then
 Dim ru% ' Runde
 If SafeArrayGetDim(ICD) <> 0 Then
  For ru = 0 To UBound(ICD)
   If left$(UCase$(ICD(ru)), 3) = "I10" And (DSic(ru) = " " Or DSic(ru) = "G") Then GoTo hyperton
  Next ru
  If pathz(RRi) < 1 Or (npathz(RRi) > 2 And pathz(RRi) < 2) Then
   normoton = True
  End If
 End If
hyperton:
#Else
  Set rsDia = Nothing
  myFrag rsDia, "SELECT 0 FROM diagview d WHERE d.pat_id = " & Pat_id & " AND d.gICD LIKE 'i10%' LIMIT 1" '  AND COALESCE(d.Dggel,0)=0
  If rsDia.EOF Then
   If pathz(RRi) < 1 Or (npathz(RRi) > 2 And pathz(RRi) < 2) Then
    normoton = True
   End If
  End If ' rsDia.EOF
  Set rsDia = Nothing
#End If
 m = 8: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  
  For i = 0 To pzl: UGrenze(i) = -1: grenze(i) = -1: rund(i) = -1: Next i
'  zMAX(0) = 6: zMAX(1) = 6: zMAX(2) = 4: zMAX(3) = 4: zMAX(4) = 6: zMAX(5) = 4: zMAX(6) = 8
  For i = 0 To pzl: zmax(i) = zmaxges: Next i '8
'  lsql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_ID & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_ID & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb" ' 29.10.18 ,einheit,nb"
  Dim rsl As New ADODB.Recordset ' Labor
  Dim lwZahl&
  Dim lab() As labtyp
  Dim obLabzugew%
  Set rsl = hollabor(CLng(Pat_id), "", 0, 0, 0, lwZahl)
  If Not rsl.BOF And lwZahl Then
   ReDim lab(rsl!dszahl)
   i = 0
   obLabzugew = True
   Do While Not rsl.EOF
    lab(i).Abkü = rsl!Abkü
    lab(i).WertSg = rsl!Wert
    lab(i).Einheit = rsl!Einheit
    lab(i).Zp = rsl!Zeitpunkt
    i = i + 1
    rsl.MoveNext
   Loop ' While Not rsl.EOF
  End If ' Not rsl.BOF And lwZahl Then
  Set rsl = Nothing
 m = 9: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  Dim obri%
'  lsql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_ID & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_ID & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb" ' 29.10.18 ,einheit,nb"
  
  ' lMP = "SELECT mp.* FROM (SELECT pat_id, mpnr FROM (SELECT mpi.pat_id, mpi.mpnr, mpi.zeitpunkt FROM (SELECT pat_id, MAX(zeitpunkt) AS zp FROM `medplan` mp WHERE pat_id = " & Pat_ID & " AND zeitpunkt <= now() GROUP BY pat_id) AS mpl LEFT JOIN `medplan` mpi ON mpl.pat_id = mpi.pat_id AND mpl.zp = mpi.zeitpunkt GROUP BY mpl.pat_id, mpl.zp, mpnr ORDER BY  mpl.pat_id, mpl.zp, mpnr DESC) AS mpii GROUP BY pat_id, zeitpunkt) i LEFT JOIN `medplan` mp ON i.pat_id = mp.pat_id AND i.mpnr = mp.mpnr"
  ' genauso schnell und kürzer wäre:
'  lMP = "SELECT pat_id,mpnr,zeitpunkt,medikament FROM medplan mp WHERE mpnr=(SELECT MAX(mpnr) FROM medplan mpi WHERE pat_id=mp.pat_id AND zeitpunkt=(SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=mp.pat_id)) AND pat_id = " & Pat_ID
  i = 0
  
  ÜS(i) = "Gw"
  Fqmin(i) = 4
  Titel(i) = "Körpergewicht [kg], Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  Gwi = i
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE art RLIKE '^gew$|^gw$|^gewi' AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
  
  ÜS(i) = "Taille"
  Fqmin(i) = 1
  If obweibl Then
   grenze(i) = 88
  Else
   grenze(i) = 102
  End If
  Titel(i) = "Taille (Mitte zwischen Hüfte und Beckenkamm) [cm], NB &#x2640;(-80)-88, &#x2642;(-94)-102,hier:" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp,IF(INSTR(inhalt,' '),LEFT(inhalt,INSTR(inhalt,' ')),inhalt) wert FROM `eintraege` WHERE art LIKE 'taille%' AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  Tailli = i
  i = i + 1
  
  ÜS(i) = "Bewegung"
  Fqmin(i) = 4
  Titel(i) = "Bewegung pro Woche (mit leichter Luftnot, NB > 2,5-5h/Wo), Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'beweg%' OR art IN ('bew','bewg','bewe')) AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1

  ÜS(i) = "RR"
  Fqmin(i) = 180
  grenze(i) = 130
  Titel(i) = "Blutdruck [mm Hg]<br>Ziel bei Nephropathie niedriger, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  RRi = i
  sql(i) = "SELECT zeitpunkt zp, rr wert FROM rr WHERE pat_id = " & Pat_id & " ORDER BY zeitpunkt DESC"
  i = i + 1
  
'  IF Not normoton THEN ' s.o.
    ÜS(i) = "OAU"
    Fqmin(i) = 1
    grenze(i) = 32
    Titel(i) = "Oberarmumfang [cm], grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
    sql(i) = "SELECT zeitpunkt zp, IF(instr(inhalt,' '),LEFT(inhalt,instr(inhalt,' ')),inhalt) wert FROM `eintraege` WHERE art LIKE 'oau%' AND pat_id = " & Pat_id & " ORDER BY zp DESC"
    nurpath(i) = 1
    i = i + 1
'  END IF
  
  ÜS(i) = "HbA1c"
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") i WHERE abkü LIKE 'HBA1C%' ORDER BY zp DESC"
'  sql(i) = "SELECT DATE(zeitpunkt) zp, wert FROM `laborneu` ln WHERE abkü RLIKE 'hba[c1]' AND einheit = '%' AND pat_id = " & Pat_ID & " UNION SELECT DATE(u.eingang) zp, w.wert FROM `" & vorsil & "us` u LEFT JOIN " & vorsil & "wert w ON u.refnr = w.refnr WHERE abkü RLIKE 'hba[c1]' AND einheit = '%' AND pat_id = " & Pat_ID & " GROUP BY zp ORDER BY zp DESC"
  Hba1ci = i
  kritnr(i) = LA_HbA1c
  sql(i) = vNS
'  sql(i) = LabEPatS(HbA1c, Pat_id)
  Fqmin(i) = 4
  grenze(i) = 6.5
  Titel(i) = "HbA1c [%], nach DCCT, NB [4,1-6,2],grenze(" & i & ")=" & grenze(i) & ", Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  If Not obdm Then nurpath(i) = 3
  i = i + 1
  
  Dim OADUzu As New ADODB.Recordset
  Select Case therart
   Case "CSII", "ICT", "CT", "Komb", "GLP1Ins", "GLP1ICT", "OAD", "GLP1"
    If therart = "OAD" Or therart = "GLP1" Then
    
     myFrag OADUzu, "SELECT mp.medikament FROM medplan mp LEFT JOIN medarten ma ON mp.medanfang=ma.medikament " & _
     "WHERE mp.pat_id = " & Pat_id & " AND (glib<>0 OR shglin<>0) AND zeitpunkt>DATE_SUB((SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id = " & Pat_id & "),INTERVAL 90 DAY)"
     If OADUzu.EOF Then
      nurpath(i) = 3
'      GoTo keinuzu
     End If
    End If
    ÜS(i) = "Uzu, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
    Fqmin(i) = 4
    Titel(i) = "Unterzucker seit letztem Besuch"
    sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'uzu%' OR art LIKE 'hypo%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
    i = i + 1
    
    ÜS(i) = "Hypo"
    Fqmin(i) = 0
    Titel(i) = "Schwere Hypoglykämie, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
    sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'hypo%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
    nurpath(i) = 3
    i = i + 1
    
   Case Else ' "Diät"
'    i = i + 2
  End Select
 m = 10: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  
'keinuzu:
  ÜS(i) = "Hyper"
  Fqmin(i) = 0
  Titel(i) = "Hyerglykämische Entgleisung, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'hyper%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
'  If Left$(rAn!Diabetestyp, 1) = "1*" Then
   nurpath(i) = 3 '1
'  Else
'   nurpath(i) = 3
'  End If
  i = i + 1
  
  ÜS(i) = "NeuSt"
  Fqmin(i) = 1
  Titel(i) = "Neuro-Status, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  Neuri = i
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE art LIKE 'usd%' AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  zmax(i) = 1
  weite(i) = "20%"
  RowSp(i) = 8
  If Not obdm Then nurpath(i) = 3
  i = i + 1
  
  ÜS(i) = "Fuß"
  Fqmin(i) = 0
  Titel(i) = "Fußuntersuchung, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  Fußi = i
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'fuß%' OR art like 'fuss%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  zmax(i) = 4
  RowSp(i) = 2
  If Not obdm Then nurpath(i) = 3
  i = i + 1
  
  ÜS(i) = "Beruf"
  Fqmin(i) = 0
  Titel(i) = "Berufsanamnese, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'beruf%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
  
  ÜS(i) = "Auto"
  Fqmin(i) = 0
  Titel(i) = "Aufklärung über Autofahrverbot ,Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'auto%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  If Not obdm Then nurpath(i) = 3
  i = i + 1
  
  ÜS(i) = "Keto"
  Fqmin(i) = 0
  Titel(i) = "Aufklärung über Ketonkörpermessung ,Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'keto%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  If Not obdm Then nurpath(i) = 3
  i = i + 1
  
  ÜS(i) = "K+"
  Fqmin(i) = 1
  Titel(i) = "Kalium [mmol/l], NB 3.5 - 5.6, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") i WHERE (abkü = 'K' OR abkü LIKE 'KALI%') ORDER BY zp DESC"
  kritnr(i) = LA_Kali
  sql(i) = vNS
'  sql(i) = LabEPatS(kali, Pat_id)
  UGrenze(i) = 3.5
  grenze(i) = 5.6
  nurpath(i) = 2
  i = i + 1
  
  ÜS(i) = "Chol"
  Fqmin(i) = 1
  grenze(i) = 200
  Titel(i) = "Gesamtcholesterin [mg/dl], grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") i WHERE abkü IN ('CHOL') ORDER BY zp DESC"
  kritnr(i) = LA_Chol
  sql(i) = vNS
'  sql(i) = LabEPatS(Chol, Pat_id)
  i = i + 1
  
  ÜS(i) = "HDL"
  Fqmin(i) = 1
  UGrenze(i) = 40
  Titel(i) = "HDL-Cholesterin [mg/dl], Ugrenze(" & i & ")=" & UGrenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") i WHERE abkü IN ('HDL','HDLC') ORDER BY zp DESC"
  kritnr(i) = LA_Hdl
  sql(i) = vNS
'  sql(i) = LabEPatS(Hdl, Pat_id)
  i = i + 1
  
  ÜS(i) = "LDL"
  LDLi = i
  Fqmin(i) = 1
  grenze(i) = 100
  Dim Manif$, RFakt$, RFz% ' Risikofaktorzahl
#If Not langsamer Then
 If SafeArrayGetDim(ICD) <> 0 Then
  For ru = 0 To UBound(ICD)
   If (ICD(ru) Like "I20.0*" Or ICD(ru) Like "I2[15]*" Or ICD(ru) Like "I46*" Or ICD(ru) Like "I6[3-6]*" Or _
   ICD(ru) Like "I67.[89]*" Or ICD(ru) Like "I69*" Or ICD(ru) Like "I70.[1289]*" Or _
   ICD(ru) Like "I73.[189]*" Or ICD(ru) Like "I7[479]*" Or ICD(ru) Like "I99*" Or _
   ICD(ru) Like "N18.4*") And DSic(ru) = "G" Then
    If Manif <> "" Then Manif = Manif & ","
    Manif = Manif & ICD(ru)
    If grenze(i) > 55 Then grenze(i) = 55
   ElseIf ICD(ru) Like "N18.3*" And DSic(ru) = "G" Then
    If Manif <> "" Then Manif = Manif & ","
    Manif = Manif & ICD(ru)
    If grenze(i) > 70 Then grenze(i) = 70
   ElseIf ICD(ru) Like "E10*" And DSic(ru) = "G" Then
    If aRisk.DurationDiagnosedDiabetes > 20 Then
     If Manif <> "" Then Manif = Manif & ","
     Manif = Manif & ICD(ru) & "& Diabdau=" & aRisk.DurationDiagnosedDiabetes
     If grenze(i) > 55 Then grenze(i) = 55
    End If
   End If ' (icd(ru) ...
   If ICD(ru) Like "E1[0-4]*" And DSic(ru) = "G" Then
    If RFakt <> "" Then RFakt = RFakt & ","
    RFakt = RFakt & ICD(ru)
    RFz = RFz + 1
    If aRisk.DurationDiagnosedDiabetes > 10 Then
     If grenze(i) > 70 Then
      If Manif <> "" Then Manif = Manif & ","
      Manif = Manif & ICD(ru) & "& Diabdau=" & aRisk.DurationDiagnosedDiabetes
      grenze(i) = 70
     End If
    End If
   ElseIf ((ICD(ru) Like "I10.9*" Or ICD(ru) Like "E78*") And DSic(ru) = "G") _
   Or (ICD(ru) Like "F17*" And (DSic(ru) = "G" Or DSic(ru) = "Z")) Then
    If RFakt <> "" Then RFakt = RFakt & ","
    RFakt = RFakt & ICD(ru) & DSic(ru)
    RFz = RFz + 1
   End If
  Next ru
  If RFz >= 3 Then
      If Manif <> "" Then Manif = Manif & ","
      Manif = Manif & RFz & " RF(" & RFakt & ")"
      If grenze(i) > 55 Then grenze(i) = 55
  ElseIf RFz > 1 Then
      If Manif <> "" Then Manif = Manif & ","
      Manif = Manif & RFz & " RF(" & RFakt & ")"
      If grenze(i) > 70 Then grenze(i) = 70
  End If
 End If ' SafeArrayGetDim(ICD) <> 0 Then
#Else
  Set rsDia = Nothing
  myFrag rsDia, "SELECT 0 FROM diagview d WHERE d.pat_id = " & Pat_id & vbCrLf & _
  "AND d.gICD RLIKE '^i(20.0|2[15]|46|6[3-6]|67.[89]|69|70.[1289]|73.[189]|7[479]|99)' LIMIT 1" '  AND COALESCE(d.Dggel,0)=0
  If Not rsDia.EOF Then
   If grenze(i) > 55 Then
    grenze(i) = 55
   End If ' grenze(i) > 55 Then
  End If ' rsDia.EOF
  Set rsDia = Nothing
#End If

  Titel(i) = "LDL-Cholesterin [mg/dl], grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a; " & Manif
 ' sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") innen WHERE abkü IN ('LDL','LDLB','LDLH','LDLC','LDLH01','LDLLG','LDLS') AND einheit = 'mg/dl' ORDER BY zp DESC"
  kritnr(i) = LA_LDL
  sql(i) = vNS
'  sql(i) = LabEPatS(LDL, Pat_id)
  i = i + 1
  
  ÜS(i) = "TG"
  Fqmin(i) = 1
  grenze(i) = 150
  Titel(i) = "Triglyceride [mg/dl], grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  TGi = i
  kritnr(i) = LA_TG
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") innen WHERE abkü IN ('TRI','TRIG','NTFE') ORDER BY zp DESC"
'  sql(i) = hollabor(Pat_id, "^TRI$|^TRIG$|^NTFE$").source
  nurpath(i) = 1
  i = i + 1
  
  ÜS(i) = "GGT"
  Fqmin(i) = 1
  If obweibl Then
   grenze(i) = 39
  Else
   grenze(i) = 66
  End If
  Titel(i) = "Gamma-GT [U/l], NB &#x2640;-39, &#x2642;-66, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  GGTi = i
  kritnr(i) = LA_GGT
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") innen WHERE abkü LIKE '%GT%' ORDER BY zp DESC"
  nurpath(i) = 1
  i = i + 1
  
  ÜS(i) = "GPT"
  Fqmin(i) = 1
  If obweibl Then
   grenze(i) = 35
  Else
   grenze(i) = 50
  End If
  Titel(i) = "GPT (ALAT) [U/l], NB &#x2640;-35, &#x2642;-50, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  GPTi = i
  kritnr(i) = LA_GPT
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") innen WHERE abkü LIKE 'GPT%' ORDER BY zp DESC"
  nurpath(i) = 1
  i = i + 1
  
  ÜS(i) = "Ferr"
  If obweibl Then
   UGrenze(i) = 15
   grenze(i) = 150
  Else
   UGrenze(i) = 30
   grenze(i) = 400
  End If
  Fqmin(i) = 0
  Titel(i) = "Ferritin [ng/ml], NB &#x2640;15-150, &#x2642;30-400, UGrenze(" & i & ")=" & UGrenze(i) & ",grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  kritnr(i) = LA_FERR
  sql(i) = vNS
  'sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") innen WHERE abkü LIKE 'FERR%' ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
  
  ÜS(i) = "Digo"
  UGrenze(i) = 0.8
  grenze(i) = 2#
  Fqmin(i) = 0
  Titel(i) = "Digoxin [ng/ml], NB 0,8-2,0, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  kritnr(i) = LA_DIGO
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") innen WHERE abkü LIKE 'DIGO%' ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
  
  ÜS(i) = "Digit"
  UGrenze(i) = 10#
  grenze(i) = 30#
  Fqmin(i) = 0
  Titel(i) = "Digitoxin [ng/ml], NB 10-30, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  kritnr(i) = LA_DIGI
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") innen WHERE abkü LIKE 'DIGI%' ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
  
  ÜS(i) = "Urin"
  Fqmin(i) = 1
  Titel(i) = "Urinbefund im Teststreifen, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  Urini = i
  ' natural JOIN verknüpft " & vorsil & "bakt.id mit laboyus.id
  'sql(i) = "SELECT * FROM (SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND art LIKE 'urin%') u1 UNION SELECT * FROM (SELECT eingang zp, keimzahl wert FROM `" & vorsil & "bakt` natural JOIN `" & vorsil & "us` WHERE pat_id = " & Pat_ID & " AND keimzahl<> '') u2 ORDER BY zp DESC"
'  sql(i) = "SELECT * FROM (SELECT zeitpunkt zp, CAST(inhalt AS char(100)) wert FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND art LIKE 'urin%' UNION SELECT eingang zp, CAST(keimzahl AS char(16)) wert FROM `" & vorsil & "us` us LEFT JOIN `" & vorsil & "bakt` b ON us.id=b.usid WHERE pat_id = " & Pat_ID & " AND keimzahl<> '')  u2 ORDER BY zp DESC"
  sql(i) = "SELECT * FROM (SELECT zeitpunkt zp, CAST(inhalt AS char(100)) wert FROM `eintraege` WHERE pat_id = " & Pat_id & " AND art LIKE 'urin%' UNION SELECT eingang zp, CAST(keimzahl AS char(16)) wert FROM labor2bakt WHERE pat_id = " & Pat_id & " AND keimzahl<> '')  u2 ORDER BY zp DESC"
  weite(i) = "7%"
  i = i + 1
  
  ÜS(i) = "Alb/U"
  grenze(i) = 30
  Fqmin(i) = 1
  Titel(i) = "Albuminurie [mg/g Kreatinin], NB < 30 mg/gCrea, bei Nichterrechnen Albumin [mg/l], grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  Albi = i
  kritnr(i) = LA_AlbCre
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBU') AND (abkü <> 'ALBU' OR wert LIKE '%<%') ORDER BY zp DESC"
  sql(i) = vNS ' LabEPatS(AlbCre, Pat_id)
  rund(i) = 0
  If Not obdm Then nurpath(i) = 3
  i = i + 1
  
  ÜS(i) = "Krea"
  Fqmin(i) = 2
  If obweibl Then
   grenze(i) = 1.1
  Else
   grenze(i) = 1.4
  End If
  Titel(i) = "Kreatinin im Serum [mg/dl], Jaffé-Methode, NB &#x2640; -1,1 mg/dl, &#x2642; -1,4 mg/dl, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  Kreai = i
  kritnr(i) = LA_Krea
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü IN ('CREAT','KRE02','KREA','KREA02','KRES') ORDER BY zp DESC"
  sql(i) = vNS ' LabEPatS(Krea, Pat_id)
  rund(i) = 1
  i = i + 1
  
  ÜS(i) = "GFR"
  Fqmin(i) = 0
  UGrenze(i) = 60
  Titel(i) = "Glomeruläre Filtrationsrate berechnet [ml/min], NB > 60 mg/dl, Ugrenze(" & i & ")=" & UGrenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  GFRi = i
  kritnr(i) = LA_eGFR
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü IN ('GFRT', 'CREACL') ORDER BY zp DESC"
  sql(i) = vNS
  rund(i) = 0
  nurpath(i) = 2
  i = i + 1
  
  ÜS(i) = "Augen-US"
  Dim RRS%, rrd%, rsr As ADODB.Recordset
  myFrag rsr, "SELECT COALESCE(SUM(if(RRzahl=0,1,RRzahl)*RRsyst)/SUM(if(RRsyst=0,0,if(RRzahl=0,1,RRzahl))),0)ds" & vbCrLf & _
  ",COALESCE(SUM(if(RRzahl=0,1,RRzahl)*RRdiast)/SUM(if(RRdiast=0,0,if(RRzahl=0,1,RRzahl))),0)dd" & vbCrLf & _
  "FROM rr" & vbCrLf & _
  "WHERE zeitpunkt>NOW()-INTERVAL 10 YEAR" & vbCrLf & _
  "AND pat_id=" & Pat_id
  RRS = rsr!Ds
  rrd = rsr!dd
  Set rsr = Nothing
  Fqmin(i) = IIf(RRS > 135 Or rrd > 85 Or obreti Or aRisk.HbA1c > 7.5, 1, 0.5) ' geändert 30.1.23 von 1
  Titel(i) = "Augenuntersuchung, Fqmin(" & i & ")=" & Fqmin(i) & "/a (rrs=" & RRS & IIf(RRS > 135, "(!)", "") & ",rrd=" & rrd & IIf(rrd > 85, "(!)", "") & ",obreti=" & obreti & IIf(obreti, "(!)", "") & ",aRisk.HbA1c=" & aRisk.HbA1c & IIf(aRisk.HbA1c > 7.5, "(!)", "") & ")"
  sql(i) = "SELECT * FROM (SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE pat_id = " & Pat_id & " AND ((art IN (" & artspezG & ") AND ((inhalt LIKE '%augenb%' AND NOT inhalt LIKE '%augenbl%' AND NOT inhalt LIKE '%augen') OR (inhalt LIKE '%augenarzt%' OR inhalt LIKE '%augenärzt%') OR (inhalt LIKE '% aa%' AND NOT inhalt LIKE '% aag%' AND NOT inhalt LIKE '% aa.%') OR art = 'aug')) OR (art = 'aa' OR art = 'augen'))" & "/* AND zeitpunkt > " & DatFor_k(Now() - 550) & " */" & ") u1 " & _
     "UNION SELECT * FROM (SELECT zeitpunkt zp, name wert FROM `tmbrie` WHERE pat_id = " & Pat_id & " AND name LIKE '%augen%' /* AND zeitpunkt > " & DatFor_k(Now() - 550) & " */" & ") u2 " & _
     "UNION SELECT * FROM (SELECT COALESCE(vorgestellt,(SELECT MIN(bhfb) FROM faelle WHERE pat_id=" & Pat_id & ")) zp, CONCAT(`Augensp zuletzt`,': ',`Augensp Befund`) wert FROM `anamnesebogen` WHERE pat_id = " & Pat_id & " AND NOT ISNULL(`Augensp zuletzt`)) u3 " & _
     "ORDER BY zp DESC"
  weite(i) = "7%"
  If Not obdm Then nurpath(i) = 3
  i = i + 1
  
  ÜS(i) = "RR-Vgl"
  Fqmin(i) = 1
  Titel(i) = "Vergleich des Blutdruckgerätes mit unserem Oberarmmanschettengerät,Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  RRVgli = i
  sql(i) = "SELECT zeitpunkt zp, rr wert FROM `rr` WHERE art LIKE 'rrvgl%' AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  zmax(i) = 3
  weite(i) = "7em"
  RowSp(i) = 2
  i = i + 1
  
  
  ÜS(i) = "BZ-Vgl"
  Fqmin(i) = 1
  Titel(i) = "Vergleich des Blutzuckergerätes mit Naßchemiegerät, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE art LIKE 'bzvgl%' AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  zmax(i) = 3
  weite(i) = "7em"
  RowSp(i) = 2
  If Not obdm Then nurpath(i) = 3
  i = i + 1
  
  Schuli = -1
' 20.9.25: Doch auch ohne Diabetes die Parameter anzeigen (Fall 68841, Meldung lf)
  If obdm Or True Then
   ÜS(i) = "Schul"
   Schuli = i
   Fqmin(i) = 1
   Titel(i) = "(Gruppen-)Schulungen, Nachschulungen, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
   sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'schul%' OR art LIKE 'sem%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
   weite(i) = "7%"
   i = i + 1
  End If
  
  ÜS(i) = "TSH"
  Fqmin(i) = 2
  UGrenze(i) = 0.27
  grenze(i) = IIf(PAlter > 60, 6, IIf(PAlter < 40, 2.5, 4.2))
  Titel(i) = "TSH [µU/ml], NB 0,27 - 2,5-4,2-6, UGrenze(" & i & ")=" & UGrenze(i) & ",grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  TSHi = i
  kritnr(i) = LA_TSH
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü IN ('TSH', 'TS1E01', 'TSH-L_K','TSBL','TSBF') ORDER BY zp DESC"
  nurpath(i) = 1
  i = i + 1
  
  ÜS(i) = "fT4 [pmol/l]"
  Fqmin(i) = 0
  UGrenze(i) = 9
  grenze(i) = 19
  Titel(i) = "fT4, UGrenze(" & i & ")=" & UGrenze(i) & ",grenze(" & i & ")=" & grenze(i) & ", Fqmin(" & i & ")=" & Fqmin(i) & "/a"
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü LIKE 'FT4%' ORDER BY zp DESC"
  fT4i = i
  kritnr(i) = LA_fT4
  sql(i) = vNS
  nurpath(i) = 2
  i = i + 1
  
  ÜS(i) = "fT3"
  Fqmin(i) = 0
  UGrenze(i) = 2.56
  grenze(i) = 6.36
  Titel(i) = ", UGrenze(" & i & ")=" & UGrenze(i) & ",grenze(" & i & ")=" & grenze(i) & ", Fqmin(" & i & ")=" & Fqmin(i) & "/a"
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü LIKE 'FT3%' ORDER BY zp DESC"
  kritnr(i) = LA_ft3
  sql(i) = vNS
  fT3i = i
  nurpath(i) = 2
  i = i + 1
  
  Const SchulPID& = 1789
  If obdm Then
   ÜS(i) = "Schuleintr"
   Titel(i) = "Schulungseintrag mit Nachname d.Pat., Fqmin(" & i & ")=" & Fqmin(i) & "/a"
   SchulEintri = i
   ' pat_id=1789: Patient "Seminar"
   sql(i) = "SELECT zeitpunkt zp, CONCAT(art, ' <- ', inhalt) wert FROM `eintraege` WHERE pat_id = " & SchulPID & " AND inhalt LIKE '%" & UmwfSQL(NName) & "%' ORDER BY zp DESC"
   nurpath(i) = 3
   i = i + 1
  End If ' obdm Then
  
  ÜS(i) = "Hb"
  Fqmin(i) = 2
  If obweibl Then
   UGrenze(i) = 11.5
   grenze(i) = 16
  Else
   UGrenze(i) = 13.5
   grenze(i) = 17.8
  End If
  Titel(i) = "Hämoglobin [g/dl], NB &#x2640;12-16, &#x2642;14-18, UGrenze(" & i & ")=" & UGrenze(i) & ",grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  Hbi = i
  kritnr(i) = LA_Hb
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü IN ('Hb') ORDER BY zp DESC"
  nurpath(i) = 1
  i = i + 1
  
  ÜS(i) = "Vit B12"
  B12i = i
  Fqmin(i) = 0
  UGrenze(i) = 211
  grenze(i) = 911
  Titel(i) = "Vitamin B12 [pg/ml], NB [211-911], UGrenze(" & i & ")=" & UGrenze(i) & ",grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  kritnr(i) = LA_B12
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü IN ('B12','VI1201','B12N') ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
  
  ÜS(i) = "Folsre"
  Foli = i
  Fqmin(i) = 0
  UGrenze(i) = 2.6
  grenze(i) = 14.6
  Titel(i) = "Folsäure [ng/ml], NB [2,6 - 14,6], UGrenze(" & i & ")=" & UGrenze(i) & ",grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  kritnr(i) = LA_FOL
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü LIKE 'FOLS%' ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
  
  ÜS(i) = "Leuko"
  Leuki = i
  Fqmin(i) = 1
  UGrenze(i) = 4#
  grenze(i) = 9.4
  Titel(i) = "Leukozyten im Blutbild [/nl], UGrenze(" & i & ")=" & UGrenze(i) & ",grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  kritnr(i) = LA_LEUK
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü IN ('LEUK', 'LEUKO') ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
    
  ÜS(i) = "CRP"
  CRPi = i
  Fqmin(i) = 0
  grenze(i) = 5#
  Titel(i) = "C-reaktives Protein [mg/dl], NB < 5, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  kritnr(i) = LA_CRP
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü IN ('CRP', 'CRPQ') ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
    
  ÜS(i) = "CK"
'  cki = i
  Fqmin(i) = 0
  grenze(i) = 190
  Titel(i) = "C(P)K [U/l], NB < 190, grenze(" & i & ")=" & grenze(i) & ",Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  kritnr(i) = LA_CK
  sql(i) = vNS
'  sql(i) = "SELECT zeitpunkt zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) wert FROM (" & lsql & ") AS innen WHERE abkü = 'CKNAC' ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
    
  ÜS(i) = "A.P."
  Fqmin(i) = 0
  Titel(i) = "Bericht über Angina pectoris, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art = 'ap') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
  
  ÜS(i) = "Carotis"
  Caroi = i
  Fqmin(i) = 0.5
  Titel(i) = "Duplexsonographie der Halsschlagadern, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE pat_id = " & Pat_id & " AND (art LIKE 'caro%' OR (art RLIKE '^(caro|duplex|dup$)' AND (inhalt LIKE '%Halsschlag%' OR inhalt LIKE '%Halsart%' OR inhalt LIKE '%Carotis%' OR inhalt LIKE '%ACC%' OR inhalt LIKE '%ACI%' OR inhalt LIKE '%subcl%'))) ORDER BY zp DESC"
  RowSp(i) = 2
  zmax(i) = 4
  i = i + 1
    
  ÜS(i) = "Car alle"
  Caroai = i
  Fqmin(i) = 0
  Titel(i) = "Duplexsonographie der Halsschlagadern alle Textstellen, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE pat_id = " & Pat_id & " AND (inhalt LIKE '%Halsschlag%' OR inhalt LIKE '%Halsart%' OR inhalt LIKE '%Carotis%') ORDER BY zp DESC"
  nurpath(i) = 2
  i = i + 1
    
  ÜS(i) = "Impf"
  Select Case Month(Now())
   Case 10, 11, 12
    Fqmin(i) = 1
   Case Else
    Fqmin(i) = 0
  End Select
  Titel(i) = "Impfaufklärung (Grippe, Pneumonie), Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'impf%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
    
#If mitcovid Then
  ÜS(i) = "ImpfCovid"
  Fqmin(i) = 1
  Titel(i) = "Impfaufklärung Covid, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE pat_id = " & Pat_id & " AND (" & vbCrLf & _
   " art IN ('cia','cib','c19i') OR " & vbCrLf & _
   "(art='vac' AND inhalt RLIKE 'covid|astra|biontech|comi') OR " & vbCrLf & _
   "(inhalt RLIKE 'covid|corona' AND inhalt RLIKE 'impf')) " & vbCrLf & _
   "ORDER BY zp DESC"
  i = i + 1
#End If
    
  ÜS(i) = "Colo"
  If PAlter > 54 And PAlter < 82 Then
   Fqmin(i) = 0.4
  Else
   Fqmin(i) = 0
  End If
  Titel(i) = "Aufklärung über Darmkrebsfrüherkennung, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'colo%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
    
  ÜS(i) = "Pros"
  If obweibl Or PAlter < 45 Or PAlter > 77 Then
   Fqmin(i) = 0
  Else
   Fqmin(i) = 1
  End If
  Titel(i) = "Aufklärung über Urologische Früherkennung, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'pros%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
    
  ÜS(i) = "Gyn"
  If obweibl And PAlter > 19 Then
   Fqmin(i) = 1
  Else
   Fqmin(i) = 0
  End If
  Titel(i) = "Aufklärung über gynäkologische Untersuchung, Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE (art LIKE 'gyn%') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
    
  ÜS(i) = "ADL"
  If rFlSchGr <> 90 And PAlter >= 70 Then
   If myEFrag("SELECT IF(EXISTS(SELECT 0 FROM diagview pfld WHERE (pfld.gicd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.gicd RLIKE '^F0[012]') AND pat_id=" & Pat_id & " LIMIT 1),1,0)").Fields(0) = 1 Then
    Fqmin(i) = 2
   Else
    Fqmin(i) = 0
   End If
  Else
   Fqmin(i) = 0
  End If
  Titel(i) = "Barthel-Index (ADL), Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp,/*CONCAT(*/" & vbCrLf & _
  " CASE WHEN inhalt LIKE 'Akt%' THEN" & vbCrLf & _
  "     LPAD(REGEXP_REPLACE(inhalt,'.*Gesamtpunktzahl [(]max. 100[)] *([0-9]+) *','\\1'),3,' ')" & vbCrLf & _
  " WHEN inhalt LIKE 'Barthel-Index bei%' THEN" & vbCrLf & _
  "     SUM(CAST(REGEXP_REPLACE(e.inhalt,CONCAT('^(?:.*?\\([0-9]+\\)){',nr.n,'}.*?\\(([0-9]+)\\).*$'),'\\1') AS UNSIGNED))" & vbCrLf & _
  "  ELSE '  ?' END /*,DATE_FORMAT(zeitpunkt,' (%d.%m.%Y)')),inhalt */wert" & vbCrLf & _
  "FROM eintraege e" & vbCrLf & _
  "JOIN (WITH RECURSIVE cte AS(SELECT 0 N UNION ALL SELECT N+1 FROM cte WHERE N<16) SELECT n FROM cte)nr" & vbCrLf & _
  "WHERE art IN('ADL','284','bar') AND pat_id=" & Pat_id & vbCrLf & _
  "GROUP BY e.id" & vbCrLf & _
  "ORDER BY zp DESC"
  i = i + 1
    
  ÜS(i) = "TUG"
  If rFlSchGr <> 90 And PAlter >= 70 Then
   If myEFrag("SELECT IF(EXISTS(SELECT 0 FROM diagview pfld WHERE (pfld.gicd IN ('Z74.9','G20.10','G20.20','R26.8','R29.6','R42','R32','R15','R13.9','R26.3','R41.0','R41.3','R41.8','R52.2','F45.41','G30.9','F29','F32.9','F69','F79.9','R63.4','R53','M62.50','R26.8','R68.8') OR pfld.gicd RLIKE '^F0[012]') AND pat_id=" & Pat_id & " LIMIT 1),1,0)").Fields(0) = 1 Then
    Fqmin(i) = 2
   Else
    Fqmin(i) = 0
   End If
  Else
   Fqmin(i) = 0
  End If
  Titel(i) = "Timed Up and Go (TUG), Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE art IN('ADL','284','bar') AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  sql(i) = "SELECT zeitpunkt zp," & vbCrLf & _
  "CONCAT(CASE WHEN inhalt='Time up and go'THEN'  ?'ELSE LPAD(REGEXP_REPLACE(inhalt,'.*enötigte Zeit\: *([0-9,.]+) [sS].*','\\1'),3,' ')END,' s') wert" & vbCrLf & _
  "FROM eintraege e" & vbCrLf & _
  "WHERE art IN('TUG','247') AND pat_id=" & Pat_id & vbCrLf & _
  "ORDER BY zp DESC"
  i = i + 1
    
    
  ÜS(i) = "gar:" ' Spaltenüberschrift
  Titel(i) = "Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  Fqmin(i) = 1
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE art IN (" & artSpezÄrzte & ") AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
    
  ÜS(i) = "Arzteintrag" ' Spaltenüberschrift
' nicht Anzuzeigende ausgrenzen
  Fqmin(i) = 1
  Titel(i) = "Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE art IN (" & artSpezÄrzte & ") AND zeitpunkt >= qanf() AND pat_id = " & Pat_id & " ORDER BY zp DESC"
  i = i + 1
  
  ÜS(i) = "ePA runterladen" ' Spaltenüberschrift
' nicht Anzuzeigende ausgrenzen
  Fqmin(i) = 1
  Titel(i) = "Epa runterladen Fqmin(" & i & ")=" & Fqmin(i) & "/a"
  sql(i) = "SELECT zeitpunkt zp, inhalt wert FROM `eintraege` WHERE " & IIf(rFlSchGr, "pat_id = " & Pat_id, "false") & " LIMIT 1"
  zmax(i) = 3
  weite(i) = "7em"
  RowSp(i) = 2
  i = i + 1

  
  If normoton Then
    nurpath(RRVgli) = 2 ' => RRVergleich wird nicht als fehlend angezeigt
    Fqmin(RRi) = 4 ' => Blutdruck muss nur vierteljährlich gemessen werden
  End If
  m = 11: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
 
  Dim obTaille%, Taille#
 
  m = 12: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
' Wenn therapieartbezogen zu wenig geschult, dann schulen
  Select Case therart
   Case "CSII", "ICT", "CT", "Komb", "GLP1Ins", "GLP1ICT": sollz = 6
   Case Else: sollz = 4 ' "Diät", "OAD", "GLP1", ""
  End Select
  
    
' für jeden Parameter i Wertezahl rsz(i) bestimmen
' in jeder Großzeile muß zunächst die Zahl der Unterzeilen ermittelt werden, diese wiederum hängt von der
' Menge der in der Großzeile dargestellten Spalten ab
  rszmax = 0
  For i = 0 To pzl
'   If i = 25 Then Stop
'   If i = 29 Or i = 53 Then Stop
   rsz(i) = 0
   If sql(i) <> vNS Then
' Recordsets öffnen
    myFrag rs(i), sql(i), adOpenStatic
'    myFrag rs(i), sql(i), adOpenStatic
    If Not rs(i).BOF Then ' sonst bleibt rsz(i)=0
     Select Case i
      Case Schuli
' 9.7.22: da das offenbar statt für Nachschulung für erneute Schulung verwendet wird, Erhöhung von 1 auf > 4 Jahre
       If fctQEnd(ZQuart(Now)) - fctQAnf(ZQuart(rs(Schuli)!Zp)) > 4 * 365 Then
        obnach = True
       End If ' fctQEnd(ZQuart(Now)) - ...
      Case Tailli
       On Error Resume Next
       Taille = MachNumerisch(rs(Tailli).Fields(1)): If Err.Number = 0 Then obTaille = True
       On Error GoTo fehler
     End Select ' Case i
     Set rszahl = Nothing
     myFrag rszahl, "SELECT COUNT(0)Zahl FROM (" & sql(i) & ") i"
     rsz(i) = rszahl!Zahl
    End If ' Not rs(i).BOF Then ' sonst bleibt rsz(i)=0
' Hier Zahl der Schulungen auswerten
    If i = Schuli Then
     If rsz(i) < sollz Or obnach Then
      Fqmin(Schuli) = 365
      Titel(i) = "(Gruppen-)Schulungen, Nachschulungen, Fqmin(" & Schuli & ")=" & Fqmin(Schuli) & "/a"
     End If ' rsz(i) < sollz
    End If ' i = Schuli Then
   ElseIf kritnr(i) <> LabArt0 And lwZahl Then
   
    For aktlwx = 0 To lwZahl
     If obLabI(kritnr(i), lab(aktlwx)) Then rsz(i) = rsz(i) + 1
    Next aktlwx
   End If ' sql(i) <> vNS then else
   If rsz(i) > zmax(i) Then rsz(i) = zmax(i)
   If rsz(i) > rszmax Then rszmax = rsz(i)
   
   
   
  Next i '   For i = 0 To pzl
 m = 13: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
 
' Grenze für Blutdruck korrigieren
  If rsz(Albi) Then
   For aktlwx = 0 To lwZahl
    If obLabI(kritnr(Albi), lab(aktlwx)) Then ' AlbCre
     WertNum = MachNumerisch(REPLACE$(lab(aktlwx).WertSg, ".", ","))
     If WertNum > grenze(Albi) Then pathAlbZ = pathAlbZ + 1
     AlbZ = AlbZ + 1
     If AlbZ = 3 And pathAlbZ = 0 Then Exit For ' Wenn die letzten 3 in Ordnung sind, dann o.k.
    End If
   Next aktlwx
'   Do While Not rs(Albi).EOF
'    wert = rs(Albi)!wert
'    WertNum = MachNumerisch(replace$(wert, ".", ","))
'    IF WertNum > grenze(Albi) THEN pathAlbZ = pathAlbZ + 1
'    AlbZ = AlbZ + 1
'    IF AlbZ = 3 AND pathAlbZ = 0 THEN Exit Do ' Wenn die letzten 3 in Ordnung sind, dann o.k.
'    rs(Albi).Move 1
'   Loop
'   rs(Albi).MoveFirst
   If pathAlbZ > 1 And pathAlbZ * 3 > AlbZ Then obNP = True
   If Not obNP Then
    If (rsz(Kreai)) Then
     For aktlwx = 0 To lwZahl
      If obLabI(kritnr(Kreai), lab(aktlwx)) Then ' Kreai
       WertNum = MachNumerisch(REPLACE$(lab(aktlwx).WertSg, ".", ","))
       If WertNum > grenze(Kreai) Then pathKreaZ = pathKreaZ + 1
       KreaZ = KreaZ + 1
       If KreaZ = 3 And pathKreaZ = 0 Then Exit For
      End If
     Next aktlwx
     If pathKreaZ > 1 And pathKreaZ * 3 > KreaZ Then obNP = True
    End If ' (rsz(Kreai)) Then
   End If ' Not obNP THEN
  End If ' rsz(Albi) Then
  If obNP Then
   grenze(RRi) = 120
  End If
' Merkblatt Fußsyndrom prüfen
  m = 15: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  
' Parameter für Fettleber-Index, FIB-4-Index und NFS (NAFLD-Fibrose-Score)
  Dim obGGT%, obGPT%, obGOT%, obTHR%, obTG%, obAlb%, GGT#, GPT#, GOT#, Thr#, TG#, Alb#
  FLI = -1
  FIB4 = -1
  NFS = -1
  If obLabzugew Then
   For aktlwx = 0 To lwZahl
    If lab(aktlwx).Zp < Now() - 365 Then Exit For
    If obLabI(LA_GGT, lab(aktlwx)) Then
     If Not obGGT Then GGT = MachNumerisch(lab(aktlwx).WertSg): obGGT = True
    ElseIf obLabI(LA_GPT, lab(aktlwx)) Then
     If Not obGPT Then GPT = MachNumerisch(lab(aktlwx).WertSg): obGPT = True
    ElseIf obLabI(LA_GOT, lab(aktlwx)) Then
     If Not obGOT Then GOT = MachNumerisch(lab(aktlwx).WertSg): obGOT = True
    ElseIf obLabI(LA_THR, lab(aktlwx)) Then
     If Not obTHR Then Thr = MachNumerisch(lab(aktlwx).WertSg): obTHR = True
    ElseIf obLabI(LA_TG, lab(aktlwx)) Then
     If Not obTG Then TG = MachNumerisch(lab(aktlwx).WertSg): obTG = True
    ElseIf obLabI(LA_AlbS, lab(aktlwx)) Then
     If Not obAlb Then Alb = MachNumerisch(lab(aktlwx).WertSg): obAlb = True
    End If
   Next ' aktlwx
  End If ' obLabzugew Then
  
  If obTG And TG <> 0 And obTaille And üdt.bmi <> 0 And obGGT Then
   If GGT <> 0 Then
    FLI = Round((Exp(0.953 * Log(TG) + 0.139 * üdt.bmi + 0.718 * Log(GGT) + 0.053 * Taille - 15.745) / (1 + Exp(0.953 * Log(TG) + 0.139 * üdt.bmi + 0.718 * Log(GGT) + 0.053 * Taille - 15.745))) * 100)
   End If ' GGT
  End If ' obTG And obTaille And üdt.bmi <> 0 And obGGT Then
  If obGPT And obGOT And obTHR Then
   If GPT And Thr Then
    FIB4 = Round(PAlter * GOT / Thr / Sqr(GPT), 2)
   End If
  End If ' obGPT And obGOT And obTHR Then
  If obGPT And obGOT And obTHR And obAlb And üdt.bmi <> 0 Then
   Dim obDMoPG%
   obDMoPG = myEFrag("SELECT IF(EXISTS(SELECT 0 FROM diagview WHERE gICD RLIKE '^E1[0-4]\.|^R73\.' AND pat_id=" & Pat_id & " LIMIT 1),1,0)").Fields(0)
   NFS = Round(-1.675 + 0.037 * PAlter + 0.094 * üdt.bmi + 1.13 * obDMoPG + 0.99 * GOT / GPT - 0.013 * Thr - 0.66 * Alb * 0.1, 3)
  End If ' obGPT And obGOT And obTHR And obAlb And üdt.bmi <> 0 Then
  
  ' 31.1.18: Angabe zum Intervall im Fußmakro berücksichtigen
  If rs(Fußi).State = 0 Then Set rs(Fußi) = Nothing: DBCn.Close: DBCn.Open: rs(Fußi).Open sql(Fußi), DBCn
  If Not rs(Fußi).BOF Then
   If rs(Fußi).Fields.COUNT > 1 Then
    Dim Intervall$, pos&, p2&
    Intervall = rs(Fußi).Fields(1)
    pos = InStr(Intervall, "nä US:")
    If pos > 0 Then
     pos = pos + 7
     p2 = InStr(pos, Intervall, ";")
     If p2 > 0 Then
      Intervall = Mid$(Intervall, pos, p2 - pos)
      Select Case Intervall
       Case "1 a"
        Fqminu = 1
       Case "6 Mo"
        Fqminu = 2
        obFußmakro = 1
       Case "<= 3Mo"
        Fqminu = 4
        obFußmakro = 2
      End Select
      Fqmin(Fußi) = Fqminu
      Titel(Fußi) = "Fußuntersuchung, Fqmin(" & Fußi & ")=" & Fqmin(Fußi) & "/a"
     End If ' p2 > 0 Then
    End If ' pos > 0 Then
   End If ' rs(Fußi).Fields.COUNT > 1 Then
  End If ' Not rs(Fußi).BOF Then
  m = 16: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
  
  If üdt.Infekt = 1 Or üdt.fußst = auff Or üdt.sens = auff Or üdt.sens = pathdok Or üdt.Puls = auff Or üdt.FEn(12) = True Then
   Fqmin(Fußi) = 4
' Wenn in diesem Quartal Neurostatus gemacht oder fällig, dann kein Fußstatus nötig
   If rs(Neuri).State = 0 Then Set rs(Neuri) = Nothing: DBCn.Close: DBCn.Open: rs(Neuri).Open sql(Neuri), DBCn
   If Not rs(Neuri).EOF Then
    If ZQuart(rs(Neuri)!Zp) = ZQuart(Now()) Then
     Fqmin(Fußi) = 0
     Titel(Fußi) = "Fußuntersuchung, Fqmin(" & Fußi & ")=" & Fqmin(Fußi) & "/a"
     Fqminu = Fqmin(Fußi)
    End If
   End If
   ' vielleicht alle 2 Jahre
   myFrag rsMB, "SELECT 0 FROM `eintraege` WHERE pat_id = " & Pat_id & " AND zeitpunkt > SUBDATE(NOW(), INTERVAL 730 DAY) AND art = 'notiz' AND inhalt LIKE 'Merkblatt Fußsyndrom%'"
   If rsMB.BOF Then
    obMB = True
    MerkblattText = "Mbl.Fußsy("
    If üdt.fußst = auff Then MerkblattText = MerkblattText & "Fußst.auff,"
    If üdt.sens = auff Then MerkblattText = MerkblattText & "Sens.auff,"
    If üdt.FEn(12) = True Or üdt.sens = pathdok Then
     MerkblattText = MerkblattText & "ICD("
     Set rsDia = Nothing
     myFrag rsDia, "SELECT DISTINCT icd FROM diagview d WHERE pat_id = " & Pat_id & " AND (d.gicd LIKE 'L89%' OR d.gicd LIKE 'M14.6%' OR d.gicd = 'G63.2') "
     Do While Not rsDia.EOF
      MerkblattText = MerkblattText & rsDia!ICD & ","
      rsDia.Move 1
     Loop
     MerkblattText = left$(MerkblattText, Len(MerkblattText) - 1) & ")"
    End If
    If üdt.Puls = auff Then MerkblattText = MerkblattText & "Puls auff,"
    MerkblattText = left$(MerkblattText, Len(MerkblattText) - 1) & ")"
   End If ' rsMB.BOF THEN
   Set rsMB = Nothing
  End If ' üdt.fußst = auff OR üdt.sens = auff OR üdt.sens = pathdok OR üdt.Puls = auff OR üdt.FEn(12) = True THEN
  m = 17: TI(m) = Timer: For p = 0 To m - 1: TI(m) = TI(m) - TI(p): Next p
'' ob noch Avandia gegeben wird
'  SET rsMB = Nothing
'  myFrag rsMB, "SELECT zeitpunkt zp, medikament wert FROM (" & lMP & ") i WHERE medikament LIKE 'Avan%'"
'  IF Not rsMB.EOF THEN obAvan = True
'  SET rsMB = Nothing
' ob noch HCT gegeben wird
'  SET rsMB = Nothing
'  myFrag rsMB, "SELECT zeitpunkt zp, medikament wert FROM (" & lMP & ") i WHERE medikament LIKE '%HCT%' OR (medikament LIKE '%comp%' AND NOT medikament RLIKE 'complex|dorzocomp|tilidin')"
'  IF Not rsMB.EOF THEN
'   Dim rsMBWert$
'   rsMBWert = rsMB!Wert
'   SET rsMB = Nothing
'   myFrag rsMB, "SELECT * FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND zeitpunkt > SUBDATE(NOW(), INTERVAL 730 DAY) AND inhalt LIKE 'Aufklärung über erhöhtes Risiko für Spinaliom und Basaliom unter HCT%'"
'   IF rsMB.EOF THEN
'    obAvan = True
'    wasAvan = "HCT (" & rsMBWert & ")"
'   END IF
'  END IF
'  SET rsMB = Nothing
  
' UGrenze und grenze werden oberhalb nur festgelegt, unterhalb nur abgefragt
  
  For i = 0 To pzl
'   IF i = Gwi THEN
'    IF Not rs(i).BOF THEN
'     IF rAn!Größe <> 0 THEN
'      bmi = MachNumerisch(rAn!Größe)
'      IF bmi > 3 THEN bmi = bmi / 100
'      bmi = ROUND(MachNumerisch(rs(i)!Wert) / (bmi * bmi), 1)
'      bmiS = bmi & " kg/m²"
'     END IF
'    END IF
'   END IF
'   If i = 29 Or i = 53 Then Stop
   loopct = 1
   
   aktlwx = 0
   angefx = 0
   For j = 0 To rsz(i)
    If sql(i) = vNS Then
     gefunden = 0
     If lwZahl Then
      Do While True 'sollte am Ende fertig werden
       If angefx Then aktlwx = aktlwx + 1
       If aktlwx > lwZahl Then Exit Do
       If obLabI(kritnr(i), lab(aktlwx)) Then
        gefunden = True
        Exit Do
       End If
       angefx = True
      Loop
     End If ' lwzahl
    Else
     gefunden = Not rs(i).BOF And Not rs(i).EOF
    End If
    If Not gefunden Then Exit For
    If UGrenze(i) <> -1 Or grenze(i) <> -1 Then
' allgemein (für alle Parameter) festlegen, ob pathologisch
     If sql(i) = vNS Then
      Wert = lab(aktlwx).WertSg
     Else
      Wert = rs(i)!Wert
     End If
     WertNum = MachNumerisch(REPLACE$(Wert, ".", ","))
     If (i = Urini And InStrB(Wert, "++") <> 0 Or InStrB(Wert, "000") <> 0) Or _
     (UGrenze(i) <> -1 And WertNum < UGrenze(i)) Or (grenze(i) <> -1 And WertNum > grenze(i)) Then
      pathol(i, loopct) = True
      pathol(i, 0) = True
     End If '
    End If ' UGrenze(i) <> -1 Or grenze(i) <> -1 Then
' bei pathologischen Laborwerten auch andere anzeigen
    If pathol(i, 0) Then
     If i = Hbi Then
      nurpath(Ferri) = 3
      nurpath(B12i) = 3
      nurpath(Foli) = 3
     End If
     If i = Leuki Then
      nurpath(CRPi) = 3
     End If
     If i = TSHi Then
      nurpath(fT4i) = 3
      nurpath(fT3i) = 3
     End If
'     IF i = TGi THEN
'      nurpath(GGTi) = 0
'      nurpath(GPTi) = 0
'     END IF
     If i = Kreai Then
      nurpath(GFRi) = 3
     End If
     If i = Caroi Then
      nurpath(Caroai) = 3
      nurpath(Caroi) = -1
     Else
      nurpath(Caroai) = -1
     End If
    End If ' pathol(i, 0) THEN
    
' auch diastolischen Wert berücksichtigen
    If i = RRi Then
     Wert = getDiast(rs(i)!Wert)
     If IsNumeric(Wert) Then
      If CDbl(Wert) >= GrenzeDiast Then
       pathol(i, loopct) = True
       pathol(i, 0) = True
      End If
     End If
    End If
    If pathol(i, loopct) Then
     pathz(i) = pathz(i) + 1
    Else
     npathz(i) = npathz(i) + 1
    End If
    
    If sql(i) = vNS Then
    Else
     rs(i).MoveNext
    End If
    loopct = loopct + 1
    If zmaxges <> 0 And loopct > zmaxges Then Exit For
   Next j
   If sql(i) <> vNS Then If Not rs(i).BOF Then rs(i).MoveFirst
  Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = 0 Then
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in tabelleinplz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' tabelleinplz


' in dodoPlz (ausgelagert)
Function plzeintraege(ByVal Pat_id$, ByRef AusS As CString)
' Einträge
  Dim rse As ADODB.Recordset
  Dim i&, DatNr&, p%
  On Error GoTo fehler
  ' ktag wieder korrumpiert
  myFrag rse, "SELECT DATE(zeitpunkt) datum, art, inhalt FROM `eintraege` WHERE Art IN (" & artSpezUS1 & ") AND pat_id = " & Pat_id & " ORDER BY zeitpunkt DESC"
  If Not rse.BOF Then
   AusS.AppVar Array("<p class='abstand'>" & vbCrLf & "</p>" & vbCrLf & "<b style=""color:blue;"">Untersuchungen:</b>" & vbCrLf & "<Table>" & vbCrLf & " <colgroup>" & vbCrLf & "  <col width=""10"">" & vbCrLf & "  <col width=""10"">" & vbCrLf & "  <col width=""1000"">" & vbCrLf & " </colgroup>", vbCrLf)
   AusS.AppVar (Array("<a name='AnkerUnters' href='#AnkerUnters' accesskey='u'></a>", vbCrLf))
   Do While Not rse.EOF
    AusS.AppVar Array(" <tr>", vbCrLf)
    AusS.AppVar Array("  <td class='lnn'>" & rse!Datum & "</td>", vbCrLf)
    AusS.AppVar Array("  <td class='lnn'>" & rse!art & "</td>", vbCrLf)
    AusS.AppVar Array("  <td>" & rse!Inhalt & "</td>", vbCrLf)
    AusS.AppVar Array(" </tr>", vbCrLf)
    rse.MoveNext
   Loop ' While Not rse.EOF
   AusS.AppVar Array("</Table>", vbCrLf)
  End If ' Not rse.BOF THEN
' Einträge
  Set rse = Nothing
  myFrag rse, "SELECT DATE(zeitpunkt) datum, art, inhalt FROM `eintraege` WHERE Art IN (" & artSpezEintr & ") AND pat_id = " & Pat_id & " ORDER BY zeitpunkt DESC"
  If Not rse.BOF Then
   AusS.AppVar Array("<p class='abstand'>" & vbCrLf & "</p>" & vbCrLf & "<b style=""color:blue;"">Einträge:</b>" & vbCrLf & "<Table>" & vbCrLf & " <colgroup>" & vbCrLf & "  <col width=""10"">" & vbCrLf & "  <col width=""10"">" & vbCrLf & "  <col width=""1000"">" & vbCrLf & " </colgroup>", vbCrLf)
   Do While Not rse.EOF
    AusS.AppVar Array(" <tr>", vbCrLf)
    AusS.AppVar Array("  <td class='lnn'>" & rse!Datum & "</td>", vbCrLf)
    AusS.AppVar Array("  <td class='lnn'>" & rse!art & "</td>", vbCrLf)
    AusS.AppVar Array("  <td>" & rse!Inhalt & "</td>", vbCrLf)
    AusS.AppVar Array(" </tr>", vbCrLf)
    rse.MoveNext
   Loop
   AusS.AppVar Array("</Table>", vbCrLf)
  End If
  
  AusS.AppVar (Array("</BODY>", vbCrLf))
  AusS.AppVar (Array("</HTML>", vbCrLf))
  
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = 0 Then
End If
Static runden%
If Err.Number = 429 Then
 runden = runden + 1
 If runden < 10 Then
   Dim QDat$, ZDat$
  If aktOSV >= win_vista Then
   QDat = uVerz & "Programmierung\riskeng.ocx"
   ZDat = Environ("programdata")
   Shell ("cmd /e:on /c copy " & Chr(34) & QDat & Chr(34) & " " & Chr(34) & ZDat & Chr(34))
    If AdminPwd = "" Then AdminPwd = holap("Administrator")
   Shell (vVerz & doalsAd & acceu & AdminPwd & " cmd /e:on /c regsvr32 " & Chr$(34) & ZDat & "\riskeng.ocx" & Chr$(34))
  Else ' aktOSV >= win_vista Then
   Shell ("cmd /c regsvr32 """ & uVerz & "programmierung\riskeng.ocx""")
  End If ' aktOSV >= win_vista Then else
  Sleep 1000
  Resume
 End If ' runden < 10 Then
End If ' Err.Number = 429 Then
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraege/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' eintraege

' for dodoPLZ und LaborIns1 <- tuBriefStandalone
Sub LaborInsPLZ(Pat_id$, SelbstStatus%, raDatBOF%, ByRef Matr$(), ByRef MForm%(), ByRef mBreiten$(), Optional ohnebr%, Optional ByVal LDLZiel%) 'Word.Document, Pat_id&)
 Dim raL As New ADODB.Recordset, raDat As New ADODB.Recordset, rLaU As New ADODB.Recordset, ls$
 Dim dSL As SortierListe
' Zeilenzahl bestimmen
 Dim ZZ&, gschl$, Vgl$, altGruppe%, Nb$ ' Normbereich
 Dim sql0$, sql1$, sqldat$ ' , sql2$
 Dim u!, o! ' oberer und unterer Grenzwert numerisch
 Dim uNG$, oNG$ ' obere und untere Normgrenze in Zeichen
 Dim uNm$, oNm$, unw$, onw$
 Dim pKz% ' Pathologisch-Kennzeichen
 Dim altgru&
 Dim ug$, og$
 Dim Kommentar$
 Dim diesselbst%
 Dim rlWS$, rlWZ#, obkl%, obgr%
 Dim buch$, pos&
' Dim Tj!(100), m%, p%
 'm = 0: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 
 On Error GoTo fehler
 If LDLZiel = 0 Then LDLZiel = 100
 SelbstStatus = 0
 Dim obselbst%, obnichtselbst%
 
#Const abnov = 1
#If abnov Then
 myFrag raL, "CALL geslabneu(" & CStr(Pat_id) & ",'" & IIf(ohnebr, "", "<br>") & "','')", adOpenStatic
 If Not raL.BOF Then
  ReDim Matr(1, raL!Zpz + 5, raL!zab + 2) ' Spaltenzahl, Zeilenzahl, matr(x,y)= Wert ' (1, = für Tooltip, 27.3.10, Umkehrung 6.4.11 ' 8.2.23 +6 statt +5 ' 9.2.23 wieder +5
  ReDim MForm(raL!Zpz + 5, raL!zab + 2) ' Formatierung (1 = schräg, 2 = fett) ' 8.2.23 +6 statt +5 ' 9.2. wieder +5
  ReDim mBreiten(raL!Zpz)
  Matr(0, 3, 0) = "Parameter"
  Matr(0, 4, 0) = "Einheit"
  Matr(0, 5, 0) = "Normbereich"
  Do While Not raL.EOF
   If raL!ezp <> 0 Then
    Matr(0, raL!zpnr + 5, 0) = Format(raL!Zeitpunkt, "dd.mm.yy")
    Matr(1, raL!zpnr + 5, 0) = raL!DatID & " " & raL!Pfad
   End If ' raL!ezp <> 0 Then
   If raL!eab <> 0 Then
    Matr(0, 0, raL!abknr) = raL!Abkü
    Matr(0, 0, raL!abknr) = "Gruppe: " & raL!Gruppe & "|Reihe: " & raL!Reihe & "|Abkü: " & raL!Abkü & "|NB: " & raL!uNG & "-" & raL!oNG
    Matr(1, 0, raL!abknr) = "Gruppe: " & raL!Gruppe & "|Reihe: " & raL!Reihe & "|Abkü: " & raL!Abkü & "|NB: " & raL!uNG & "-" & raL!oNG
    Matr(0, 3, raL!abknr) = IIf(raL!Langtext = "HbA1c Eigenlabor", "HbA1c", raL!Langtext)
    Matr(0, 4, raL!abknr) = raL!Einheit
    Matr(1, 4, raL!abknr) = raL!abknr
     
    ug = raL!uNG
    Matr(0, 5, raL!abknr) = ug & "-" & og
    ug = MachNumerisch(ug)
    If ug = "0" Then ug = ""
    If IsNumeric(ug) Then ug = CDbl(ug)
    og = raL!oNG
    og = MachNumerisch(og)
    If og = "0" Then og = ""
    If IsNumeric(og) Then og = CDbl(og)
    Matr(1, 5, raL!abknr) = ug & "-" & og
    Matr(0, 1, raL!abknr) = REPLACE$(raL!uNG, "1:", "")
    Matr(0, 2, raL!abknr) = REPLACE$(raL!oNG, "1:", "")
    
    Matr(1, 5, raL!abknr) = raL!Nb
'    Matr(0, 5, raL!abknr) = raL!uNG & "-" & raL!oNG
    Matr(0, 5, raL!abknr) = raL!uNG & "-" & raL!oNG
    Select Case raL!Abkü
     Case "LDL", "LDLB", "LDLH", "LDLC", "LDLH01", "LDLLG", "LDLS"
      If raL!Einheit = "mg/dl" Then
       og = LDLZiel
       Matr(0, 2, raL!abknr) = LDLZiel
       Matr(0, 5, raL!abknr) = "-" & LDLZiel
      End If
    End Select
   End If ' raL!eab <> 0 Then
   
   Kommentar = raL!Kom
   ' Hier wird der Wert eingetragen:
   Matr(0, raL!zpnr + 5, raL!abknr) = raL!Wert
   On Error Resume Next ' Pat. 52240 am 3.2.26
   Matr(1, raL!zpnr + 5, raL!abknr) = raL!info ' raL!Labor
   On Error GoTo fehler
   
'einf8
'#If später Then
   diesselbst = 0
   If LenB(Kommentar) <> 0 And Kommentar <> "; " And Kommentar <> "; :    /   *:" And Kommentar <> "; :    /  * :" And Kommentar <> "; :    / *  :" Then
    Matr(1, raL!zpnr + 5, raL!abknr) = Matr(1, raL!zpnr + 5, raL!abknr) & vbCrLf & Kommentar
    If InStrB(Kommentar, "manuell") <> 0 Then
     diesselbst = 1
'     Tabl.cell(raL!abknr + 1, raL!zpnr - 2).Range.Font.Italic = True
     MForm(raL!zpnr + 5, raL!abknr) = MForm(raL!zpnr + 5, raL!abknr) + 1
     
    ElseIf LenB(raL!Wert) = 0 Then
'     Matr(raL!zpnr + Diffbr, raL!abknr) = Matr(raL!zpnr, raL!abknr) & IIf(lenb(Matr(raL!zpnr, raL!abknr)) = 0, vns, " ") & kommentar
     Matr(0, raL!zpnr + 5, raL!abknr) = left$(Kommentar, 10)
    End If
   End If ' NOT ISNULL(kommentar) AND LenB(kommentar) <> 0 AND kommentar <> "; " AND kommentar <> "; :    ////////   *:" AND kommentar <> "; :    ////////  * :" AND kommentar <> "; :    //////// *  :" THEN
  If diesselbst Then obselbst = 1 Else obnichtselbst = 1
  pKz = 0 ' "n" normal
  If IsNumeric(raL!Wert) Or raL!Wert Like "1:*" Or raL!Wert Like "<*" Or raL!Wert Like ">*" Then
   obkl = 0
   obgr = 0
   rlWS = raL!Wert
   For pos = 1 To Len(raL!Wert)
    If InStrB("0123456789,.:-+<>", Mid$(rlWS, pos, 1)) = 0 Then
     rlWS = left$(rlWS, pos - 1) & Mid$(rlWS, pos + 1)
     pos = pos - 1
    End If
   Next pos
   If rlWS Like "1:<*" Then rlWS = "<1:" & Mid$(rlWS, 4)
   If left$(rlWS, 1) = "<" Then
    obkl = True
    rlWS = Mid$(rlWS, 2)
   ElseIf left$(rlWS, 1) = ">" Then
    obgr = True
    rlWS = Mid$(rlWS, 2)
   End If
   rlWS = REPLACE$(REPLACE$(REPLACE$(REPLACE$(rlWS, "^", ""), "1:", ""), ".", ","), ",,", ",")
   If rlWS = vNS Then rlWS = "0"
   If IsNumeric(rlWS) Then rlWZ = CDbl(rlWS)
   If Matr(0, 0, raL!abknr) = "KSM" Or Matr(0, 0, raL!abknr) = "LYATU" Then ' 9.2.22
     pKz = 1
   ElseIf LenB(Matr(0, 1, raL!abknr)) <> 0 Then
    If IsNumeric(Matr(0, 1, raL!abknr)) Then
     On Error Resume Next ' Pat. 53428: "5.11.11"; 11.12.11
     If rlWZ < CDbl(REPLACE$(Matr(0, 1, raL!abknr), ".", ",")) And Not obgr Then
      pKz = 1 ' "p"
     End If
     On Error GoTo fehler
    End If
   End If
   If LenB(Matr(0, 2, raL!abknr)) <> 0 Then
    On Error Resume Next
    If IsNumeric(Matr(0, 2, raL!abknr)) Then
     If rlWZ > CDbl(REPLACE$(Matr(0, 2, raL!abknr), ".", ",")) And Matr(0, 2, raL!abknr) <> "0" And Not obkl Then
      pKz = 1 ' "p"
     End If
    End If
    On Error GoTo fehler
   End If
'   IF o <> -99 AND rlwd > o THEN pKz = "p"
  End If ' IsNumeric(ral!Wert) OR ral!Wert LIKE "1:*" OR ral!Wert LIKE "<*" OR ral!Wert LIKE ">*" THEN
  On Error Resume Next
'  If pKz = 1 Then '"p" THEN 'If Tabl.Cell(raL!abknr + 1, raL!zpnr - 2) <> "" THEN
  If (raL!obpath <> "" And raL!obpath <> "0") Or pKz = 1 Then
   MForm(raL!zpnr + 5, raL!abknr) = MForm(raL!zpnr + 5, raL!abknr) + 2
  End If
  On Error GoTo fehler
'  SET rlpar = Nothing
'  End If 'NOT ISNULL(ral!Abkü) THEN
' #End If ' später
'einf8
   raL.MoveNext
  Loop ' While Not ral.EOF
 End If ' Not raL.BOF Then
#Else ' obnov
' sql1 = "SELECT kommentar FROM `labor1a` WHERE pat_id = " & Pat_id & " AND kommentar = ""Dieser Eintrag wurde manuell erzeugt."" LIMIT 1"
' raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
' IF Not raDat.BOF THEN SelbstStatus = SelbstStatus + 1
' sql1 = "SELECT kommentar FROM `labor1a` WHERE pat_id = " & Pat_id & " AND kommentar <> ""Dieser Eintrag wurde manuell erzeugt."" LIMIT 1"
' SET raDat = Nothing
' raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
' IF Not raDat.BOF THEN
'   SelbstStatus = SelbstStatus + 2
' ELSE ' 19.3.17, Pat. 62118: kein Wert aus Labor2a erscheint in Labor1a
'   Dim l2_1 As New ADODB.Recordset
'   l2_1.Open "SELECT wert FROM `labor2a` WHERE pat_id=" & Pat_id & " AND NOT ISNULL(wert) LIMIT 1", DBCn, adOpenDynamic, adLockReadOnly
'   IF Not l2_1.BOF THEN SelbstStatus = SelbstStatus + 2
'   l2_1.Close
' END IF
' raDat.Close
 
' gschl = vNS
' altGruppe = 0
' Dim rgschl As New ADODB.Recordset
' myFrag rgschl, "SELECT geschlecht FROM `namen` WHERE pat_id = " & Pat_id
' IF Not rgschl.EOF THEN
'  gschl = rgschl!Geschlecht
' END IF
' rgschl.Close
 
nochmal:

 ' 19.3.17: Cortisoltagesprofil wird noch nicht angezeigt (Wegnahme von Zeitpunkt aus Gruppierungsfelder in der nächsten Zeile nützt auch nichts)
' sql2 = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb"
' sql2 = "SELECT lab.* FROM (SELECT @patid:= " & Pat_id & " nix) nul, geslab lab"
 'm = 1: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 Dim lz&, gz&, dz&
 ' wenn hollabor vorher einmal mit dem selben Pat. ohne obnachgruppe aufgerufen wurde, geht diese Funktion 0,3s schneller
 Set raL = hollabor(CLng(Pat_id), "", 0, 0, obnachgruppe:=True, Zahl:=lz, Einheit:="", gz:=gz, dzz:=dz, dSL:=dSL, obneu:=True) ' , obUpdate:=True)
 'm = 2: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 ZZ = lz + gz
' ZZ = myEFrag("SELECT COUNT(0)+COUNT(DISTINCT gruppe) Zahl FROM (SELECT * FROM geslaba WHERE (reihe <> 999 OR ISNULL(reihe)) GROUP BY Gruppe, Reihe, Abkü, Einheit,uNg,oNg ORDER BY gruppe, reihe) i")!Zahl
' #If False THEN
'' sql0 = "SELECT Gruppe, Reihe, i.Abkü, i.Langtext,i.Einheit,i.nb, i.uNg,i.oNg FROM (" & sql2 & ") i LEFT JOIN `laborparameter` lp ON i.abkü = lp.`abkü` AND IF(ISNULL(i.einheit) OR i.einheit="""", ""kA"",i.einheit) = lp.`einheit` WHERE pat_id = " & CStr(Pat_id) & " AND (NOT ISNULL(wert) OR NOT ISNULL(kommentar)) AND (reihe <> 999 OR ISNULL(reihe)) GROUP BY Gruppe, Reihe, i.Abkü, i.Einheit,CAST(i.uNg * 100 AS decimal),CAST(i.oNg * 100 AS decimal) ORDER BY gruppe, reihe"
' IF lies.obMySQL THEN sql0 = replace$(sql0, "iif(", "if(")
' SET ral = Nothing
'' GoTo nochmal
' ral.Open sql0, DBCn, adOpenDynamic, adLockReadOnly
' altgru = vNS
' Open "v:\matrzeilen.txt" For Output AS #339
' Do While Not ral.EOF
'  IF ral!gruppe <> altgru THEN
'   Print #339, Str$(ZZ) + ": '" + Str$(ral!gruppe) + "'"
'   ZZ = ZZ + 1
'   altgru = ral!gruppe
'  END IF
'  Print #339, Str$(ZZ) + ": Abkü: '" + ral!abkü + "', Einheit: '" + ral!Einheit + "'"
'  ZZ = ZZ + 1
'  ral.MoveNext
' Loop
' Close #339
' ZZ = ZZ + 1
'#END IF 'endif
 
 syscmd 4, "Labor (1) nach Zeilenzahlbestimmung"
 'm = 3: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
nochmal1:
#If problematisch Then
 sqldat = "SELECT DATE(zeitpunkt) Datum FROM geslaba GROUP BY DATE(zeitpunkt) ORDER BY zeitpunkt"
 If dz = 0 Then dz = myEFrag("SELECT COUNT(0) ct FROM (" & sqldat & ") innen")!ct
#ElseIf False Then
 If dz = 0 Then dz = myEFrag("call geslabkatz(" & CStr(Pat_id) & ",'zeitpunkt')")!Zahl
#End If
' IF lies.obMySQL THEN
'  sql1 = replace$(replace$(sql1, "datevalue(", "date("), "iif(", "if(")
' END IF
' GoTo nochmal1
 
' Tabl.cell(1, 1).Range = "Parameter"
' Tabl.cell(1, 2).Range = "Einheit"
' Tabl.cell(1, 3).Range = "Normbereich"
 
 syscmd 4, "Labor (1) nach Überschrift"
 Dim i&, j&
 If False And Not dSL Is Nothing Then
 Else
  Set dSL = New SortierListe
  Set raDat = Nothing
#If problematisch Then
'  raDat.Open sqldat, DBCn, adOpenDynamic, adLockReadOnly
  myFrag raDat, sqldat
#Else
' raDat.Open "CALL geslabkatg(" & CStr(Pat_id) & ",'Zeitpunkt')", DBCn, adOpenDynamic, adLockReadOnly
' myFrag raDat, "CALL geslabkatg(" & CStr(Pat_ID) & ",'Zeitpunkt')"
' myFrag raDat, "CALL geslabdp" & IIf(ohnebr, "ohnebr", "") & "(" & CStr(Pat_ID) & ",'')", adOpenStatic
 myFrag raDat, "CALL geslabneu(" & CStr(Pat_id) & ",'" & IIf(ohnebr, "", "<br>") & "','')", adOpenStatic
#End If
  If raDat.BOF Then raDatBOF = True
  Dim SD As SortierDatum
  If Not raDat.BOF Then
   dz = 0
' zur Einsparung von hollabor, funktioniert aber noch nicht 3.4.25, da gz 15 statt 13 und lz 2048 statt 379:
'  Dim gSL As New SortierListe
'  Dim SG As SortierString
'   lz = raDat!Zahl
'   gz = 0
   Do While Not raDat.EOF
'   Matr(0, i + 5, 0) = Format$(raDat!Zp, "dd.mm.yy")
     Set SD = New SortierDatum
     SD.Datum = Int(raDat!Zeitpunkt)
     If dSL.GetItem(SD) Is Nothing Then
      dSL.sCAdd SD
      dz = dz + 1
     End If
'     Set SG = New SortierString
'     If Not IsNull(raDat!Gruppe) Then SG.Stri = raDat!Gruppe
'     If gSL.GetItem(SG) Is Nothing Then
'      gSL.sCAdd SG
'      gz = gz + 1
'     End If
'   If i > dz Then Exit Do
    raDat.Move 1
   Loop
   raDat.Close
  End If ' not radat.bof
'  ZZ = lz + gz
 End If
 syscmd 4, "Labor (1) nach Datumseintragung"
 'm = 4: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 If dz = 0 Then dz = 1
 ReDim Matr(1, dz + 5, ZZ + 2) ' Spaltenzahl, Zeilenzahl, matr(x,y)= Wert ' (1, = für Tooltip, 27.3.10, Umkehrung 6.4.11 ' 8.2.23 +6 statt +5 ' 9.2.23 wieder +5
 ReDim MForm(dz + 5, ZZ + 2) ' Formatierung (1 = schräg, 2 = fett) ' 8.2.23 +6 statt +5 ' 9.2. wieder +5
 ' links noch: 0:Abkü, 1:unterer NW, 2:oberer NW, dann 3:Langtext, 4:Einheit, 5:Normbereich
 ReDim mBreiten(dz)
 
 Matr(0, 3, 0) = "Parameter"
 Matr(0, 4, 0) = "Einheit"
 Matr(0, 5, 0) = "Normbereich"
 For i = 1 To dSL.COUNT
  Matr(0, i + 5, 0) = Format$(dSL.Item(i).Datum, "dd.mm.yy")
  If i > dz Then Exit For
 Next i
 
 
 Dim rlp As New ADODB.Recordset
 j = 1
 altgru = -1
 Dim ralSource$
 ralSource = raL.Source
 If raL.State = 0 Then
  Set raL = Nothing
  raL.Open ralSource, DBCn, adOpenStatic, adLockReadOnly
 End If
 If Not raL.BOF Then
 raL.MoveFirst
 Dim zeile&
 Do While Not raL.EOF
  zeile = zeile + 1
  If Not IsNull(raL!Abkü) Then
   If raL!Gruppe <> altgru Then
    Matr(0, 0, j) = "Gruppe: " & raL!Gruppe & "|Reihe: " & raL!Reihe & "|Abkü: " & raL!Abkü & "|NB: " & raL!uNG & "-" & raL!oNG
    altgru = raL!Gruppe
    j = j + 1
   End If
   Matr(0, 0, j) = raL!Abkü
   Matr(1, 0, j) = "Gruppe: " & raL!Gruppe & "|Reihe: " & raL!Reihe & "|Abkü: " & raL!Abkü & "|NB: " & raL!uNG & "-" & raL!oNG
   Matr(0, 4, j) = raL!Einheit
   Matr(1, 4, j) = zeile
'    IF NOT ISNULL(ral!Nb) THEN Matr(0,5, j) = ral!Nb
   ug = Trim$(IIf(IsNull(raL!uNG), vNS, raL!uNG))
   og = Trim$(IIf(IsNull(raL!oNG), vNS, raL!oNG))
   Matr(0, 5, j) = ug & "-" & og
   ug = MachNumerisch(ug)
   If ug = "0" Then ug = ""
   og = MachNumerisch(og)
   If og = "0" Then og = ""
   If IsNumeric(ug) Then ug = CDbl(ug)
   If IsNumeric(og) Then og = CDbl(og)
   Matr(1, 5, j) = ug & "-" & og
   If Not IsNull(raL!uNG) Then Matr(0, 1, j) = REPLACE$(IIf(IsNull(raL!uNG), vNS, raL!uNG), "1:", "")
   If Not IsNull(raL!oNG) Then Matr(0, 2, j) = REPLACE$(IIf(IsNull(raL!oNG), vNS, raL!oNG), "1:", "")
   On Error Resume Next
   Matr(0, 3, j) = raL!Langtext
   On Error GoTo fehler
   Select Case Matr(0, 0, j)
    Case "LDL", "LDLB", "LDLH", "LDLC", "LDLH01", "LDLLG", "LDLS"
     If Matr(0, 4, j) = "mg/dl" Then
      og = LDLZiel
      Matr(0, 5, j) = "-" & LDLZiel
     End If
   End Select
   j = j + 1
  End If ' Not IsNull(raL!Abkü) Then
  raL.Move 1
 Loop ' While Not raL.EOF
 End If
 syscmd 4, "Labor (1) nach Normbereichseingabe"
 
 'm = 5: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 Set raL = Nothing ' jetzt ganz ohne Gruppieren
#If problematisch Then
' ral.Open "SELECT * FROM geslaba", DBCn, adOpenDynamic, adLockReadOnly
 myFrag raL, "SELECT * FROM geslaba", adOpenStatic
#Else
' myFrag raL, "CALL geslabdp" & IIf(ohnebr, "ohnebr", "") & "(" & CStr(Pat_ID) & ",'')", adOpenStatic
 myFrag raDat, "CALL geslabneu(" & CStr(Pat_id) & ",'" & IIf(ohnebr, "", "<br>") & "','')", adOpenStatic
' myFrag ral, "CALL geslabdp(" & CStr(Pat_ID) & ",'WHERE (reihe <> 999 OR ISNULL(reihe)) GROUP BY zeitpunkt DESC,abkü,einheit')", adOpenStatic
#End If
 'm = 6: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
' ral.Open sql2, DBCn, adOpenDynamic, adLockReadOnly
 Do While Not raL.EOF
'  If ral!Wert = "6.9" Then Stop
  If Not IsNull(raL!Abkü) Then
   SD.Datum = Int(raL!Zeitpunkt)
   dSL.scinit
   dSL.Eingrenz SD
   i = dSL.sCa + 5
#If bewaehrt Then
   i = 6
   Do While CDate(IIf(LenB(Matr(0, i, 0)) = 0, 0, Matr(0, i, 0))) < DateValue(raL!Zeitpunkt)
    i = i + 1
    If i >= dz + 5 Then Exit Do
   Loop
   If i > UBound(Matr, 2) Then
    i = UBound(Matr, 2)
   End If
#End If
  j = 0
'  IF Left$(ral!Abkü, 2) = "HB" AND ral!Einheit = "%" THEN Stop
  Do While j <= UBound(Matr, 3)
   If Matr(0, 0, j) = raL!Abkü Then
    If LCase$(Matr(0, 4, j)) = LCase$(raL!Einheit) Then
     Dim NGVgl$
     ug = Trim$(IIf(IsNull(raL!uNG), vNS, raL!uNG))
     If ug = "0" Then ug = ""
     og = Trim$(IIf(IsNull(raL!oNG), vNS, raL!oNG))
     If og = "0" Then og = ""
     ug = MachNumerisch(ug)
     If ug = "0" Then ug = ""
     og = MachNumerisch(og)
     If og = "0" Then og = ""
     If IsNumeric(ug) Then ug = CDbl(ug)
     If IsNumeric(og) Then og = CDbl(og)
     NGVgl = ug & "-" & og
     If Matr(1, 5, j) = NGVgl Then
      Exit Do
     End If
    End If
   End If
   j = j + 1
  Loop
  If j > UBound(Matr, 3) Then
'   Dim rlpar As New ADODB.Recordset
'   myFrag rlpar, "SELECT gruppe, reihe, langtext FROM `laborparameter` lp WHERE abkü = '" & ral!abkü & "' AND einheit = '" & IIf(ISNULL(ral!Einheit) OR ral!Einheit = "", "kA", ral!Einheit) & "'"
'   IF Not rlpar.EOF THEN IF rlpar!reihe = 999 THEN GoTo näZei
   j = UBound(Matr, 3)
 '   ON Error Resume Next
   ReDim Preserve Matr(1, UBound(Matr, 2), UBound(Matr, 3) + 1)
   ReDim Preserve MForm(UBound(Matr, 2), UBound(Matr, 3) + 1)
   Matr(0, 0, j) = raL!Abkü
'   IF Not rlpar.EOF THEN
    Matr(1, 0, j) = "Gruppe: " & raL!Gruppe & "|Reihe: " & raL!Reihe & "|Abkü: " & raL!Abkü & "|NB: " & raL!uNG & "-" & raL!oNG
    Matr(0, 3, j) = IIf(IsNull(raL!Langtext), "", raL!Langtext) ' rlpar!Langtext
'   Else
'    Matr(1, 0, j) = ral!abkü
'    Matr(0, 3, j) = ral!abkü
'   END IF
   Matr(0, 4, j) = raL!Einheit
   Matr(0, 5, j) = Trim$(IIf(IsNull(raL!uNG), vNS, raL!uNG)) & "-" & Trim$(IIf(IsNull(raL!oNG), vNS, raL!oNG))
'    ON Error GoTo fehler
  End If ' j > UBound(Matr, 3) THEN
  If i <= dz + 5 Then
   If Matr(1, i, 0) = "" Then Matr(1, i, 0) = raL!DatID & " " & raL!Pfad
   If IsNull(raL!Kommentar) Then Kommentar = "" Else Kommentar = REPLACE$(raL!Kommentar, "//", "/")
   While InStrB(Kommentar, "//")
    Kommentar = REPLACE$(Kommentar, "//", "/")
   Wend
   ' Hier wird der Wert eingetragen:
   Matr(0, i, j) = IIf(IsNull(raL!Wert), IIf(Not IsNull(Kommentar), left$(Kommentar, 10), vNS), raL!Wert)
   Matr(1, i, j) = raL!Labor
   If InStrB(Kommentar, "//") <> 0 Then
    MsgBox " Fehler in LaborInsPLZ: InstrB(kommentar, ""//"") <> 0"
    Stop
   End If
   diesselbst = 0
   If Not IsNull(Kommentar) And LenB(Kommentar) <> 0 And Kommentar <> "; " And Kommentar <> "; :    ////////   *:" And Kommentar <> "; :    ////////  * :" And Kommentar <> "; :    //////// *  :" Then
    Matr(1, i, j) = Matr(1, i, j) & vbCrLf & Kommentar
    If InStrB(Kommentar, "manuell") <> 0 Then
     diesselbst = 1
'     Tabl.cell(j + 1, i - 2).Range.Font.Italic = True
     MForm(i, j) = MForm(i, j) + 1
     
    ElseIf LenB(raL!Wert) = 0 Then
'     Matr(i + Diffbr, j) = Matr(i, j) & IIf(lenb(Matr(i, j)) = 0, vns, " ") & kommentar
#If False Then
     Matr(0, i, j) = Matr(1, i, j) & vbCrLf & Kommentar
#Else
     Matr(0, i, j) = IIf(Not IsNull(Kommentar), left$(Kommentar, 10), vNS)
#End If
    End If
   End If ' NOT ISNULL(kommentar) AND LenB(kommentar) <> 0 AND kommentar <> "; " AND kommentar <> "; :    ////////   *:" AND kommentar <> "; :    ////////  * :" AND kommentar <> "; :    //////// *  :" THEN
  End If ' i <= dz + 5
  If diesselbst Then obselbst = 1 Else obnichtselbst = 1
  pKz = 0 ' "n" normal
  If IsNumeric(raL!Wert) Or raL!Wert Like "1:*" Or raL!Wert Like "<*" Or raL!Wert Like ">*" Then
   obkl = 0
   obgr = 0
   rlWS = raL!Wert
   For pos = 1 To Len(raL!Wert)
    If InStrB("0123456789,.:-+<>", Mid$(rlWS, pos, 1)) = 0 Then
     rlWS = left$(rlWS, pos - 1) & Mid$(rlWS, pos + 1)
     pos = pos - 1
    End If
   Next pos
   If rlWS Like "1:<*" Then rlWS = "<1:" & Mid$(rlWS, 4)
   If left$(rlWS, 1) = "<" Then
    obkl = True
    rlWS = Mid$(rlWS, 2)
   ElseIf left$(rlWS, 1) = ">" Then
    obgr = True
    rlWS = Mid$(rlWS, 2)
   End If
   rlWS = REPLACE$(REPLACE$(REPLACE$(REPLACE$(rlWS, "^", ""), "1:", ""), ".", ","), ",,", ",")
   If rlWS = vNS Then rlWS = "0"
   If IsNumeric(rlWS) Then rlWZ = CDbl(rlWS)
   If Matr(0, 0, j) = "KSM" Or Matr(0, 0, j) = "LYATU" Then ' 9.2.22
     pKz = 1
   ElseIf LenB(Matr(0, 1, j)) <> 0 Then
    If IsNumeric(Matr(0, 1, j)) Then
     On Error Resume Next ' Pat. 53428: "5.11.11"; 11.12.11
     If rlWZ < CDbl(REPLACE$(Matr(0, 1, j), ".", ",")) And Not obgr Then
      pKz = 1 ' "p"
     End If
     On Error GoTo fehler
    End If
   End If
   If LenB(Matr(0, 2, j)) <> 0 Then
    On Error Resume Next
    If IsNumeric(Matr(0, 2, j)) Then
     If rlWZ > CDbl(REPLACE$(Matr(0, 2, j), ".", ",")) And Not obkl Then
      pKz = 1 ' "p"
     End If
    End If
    On Error GoTo fehler
   End If ' LenB(Matr(0, 2, j)) <> 0 Then
'   IF o <> -99 AND rlwd > o THEN pKz = "p"
  End If ' IsNumeric(ral!Wert) OR ral!Wert LIKE "1:*" OR ral!Wert LIKE "<*" OR ral!Wert LIKE ">*" THEN
  On Error Resume Next
  If pKz = 1 Then '"p" THEN 'If Tabl.Cell(j + 1, i - 2) <> "" THEN
   MForm(i, j) = MForm(i, j) + 2
  End If
  On Error GoTo fehler
näZei:
'  SET rlpar = Nothing
  End If 'NOT ISNULL(ral!Abkü) THEN
  raL.MoveNext
 Loop ' While Not raL.EOF
 'm = 7: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 
 For j = 0 To ZrZ
  If Matr(0, 3, j) = "HbA1c Eigenlabor" Then
   Matr(0, 3, j) = "HbA1c"
  End If
 Next j
#End If ' obnov
 
 syscmd 4, "Labor (1) nach Dateneintragung "

 raL.Close
 SelbstStatus = obselbst + obnichtselbst + obnichtselbst
 'm = 8: Tj(m) = Timer: For p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborInsPLZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub      ' LaborIns

' 21.8.20: z.Zt. nicht aufgerufen
Function doPLZeinzeln()
 Dim pid$, ZeilZ$
 ZeilZ = "8"
 pid = InputBox("Bitte Pat-ID aus Turbomed eingeben:", "Patientenlaufzettel Eingabe")
' ZeilZ = InputBox("Bitte gewünschte Tabellenzeilenzahl eingeben:", "Patientenlaufzettel Eingabe", ZeilZ)
 If IsNumeric(pid) And IsNumeric(ZeilZ) Then
  Call dodoplz(pid, plzVz, Now, Now - Int(Now), True, "", CDbl(ZeilZ))
 End If
End Function ' doPLZeinzeln
