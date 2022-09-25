Attribute VB_Name = "Formular"
Option Explicit
Declare Function Tüt& Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long)
Const übertragen$ = "übertragen"
Dim obstumm%
#Const mitab = True ' auch noch in Lese
#If mitab Then
Public Const CStrMy$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;user=praxis;pwd=***REMOVED***;database="
#End If
'Public Const wdLineStyleSingle% = 1
'Public Const wdWindowStateMaximize% = 1
'Public Const wdFindContinue% = 1
'Public Const wdAlignTabLeft% = 0
'Public Const wdTabLeaderDots% = 1
'Public Const wdReplaceAll% = 2
'Public Const wdWord9TableBehavior% = 1
'Public Const wdAutoFitContent% = 1
'Public Const wdColorRed% = 255
'Public Const wdBorderLeft% = -2
'Public Const wdLineWidth050pt% = 4
'Public Const wdColorAutomatic& = -16777216
'Public Const wdBorderRight% = -4
'Public Const wdBorderTop% = -1
'Public Const wdBorderBottom% = -3
'Public Const wdBorderHorizontal% = -5
'Public Const wdBorderVertical% = -6
'Public Const wdLineStyleNone% = 0
'Public Const wdBorderDiagonalDown% = -7
'Public Const wdBorderDiagonalUp% = -8
'Public FNam$() ' FreigabeNamen
'Public FInh$() ' FreigabeInhalt
'Public FZ%     ' Freigabezahl
Public Const ArzSpez = "art in (""notiz"",""ni"",""telef"",""gs"",""rz"",""ga"",""ag"",""hz"",""ts"",""tst"",""cr"",""ep"",""bga"")"
'("notiz","ni","telef","gs","rz","ga","ag","hz","ts","cr")
Public Const MarkierFarbe$ = 9803263 ' rötlich
Public DatPfad$, DatPfadPI$, RDatPfad$
Public HANrBf$() ' Aktuell für den Brief verwendeter 1. und 2. Hausarzt
Public QMdbAkt$
'Public rafDE As ADODB.Recordset ' für Diagnosen Export
Dim raDok As New ADODB.Recordset, rsNa As New ADODB.Recordset ', rLaU As DAO.Recordset, rAnAlt As DAO.Recordset
Public raLau As New ADODB.Recordset
Public raDokab As New ADODB.Recordset ' Abgehakte Dokumente
'Dim aNaüm As New ADODB.Recordset ' Anamnesebogen in form_current
Dim HArst As New ADODB.Recordset
Dim raFa As New ADODB.Recordset ' für Anamnesebogen-Form_Current
'Public TMInterface As Object
'Public PcDokPfad$
Dim An1Pfad$, An2Pfad$, AnAPfad$, CheckPfad$, VorlagenPfad$
Dim Alter%, Pat_id&, GebDat As Date, fm As Form, obWeib As Boolean
Dim dsnr%, altDS(), altpat_id
Public MDIICD$(30) ' ICD-Nummern im Feld MDI
Public MDIDiag$(30) ' Diagnosentext im Feld MDI
Public MDIn% ' Fallzahl für die Zeile in MDI
' Dim i%
Public CGr%
'Dim nzl ' Neue Zeile -> vbcrlf
Dim nzw$ ' neue Zeile in Winword
Dim lTherZP As Date ' Zeitpunkt des letzten Therapieplans
Public obkNeph% ' ob keine Nephropathie zu bestehen scheint (-1 => vorliegende Diagnose wird ignoriert)
Public obMakroAlb% ' ob eine Makroalbuminurie vorliegt
' Aus Epikrise zur allgemeinen Verfügung:
    Public Const FEZ% = 27, maxIcd% = 5
    Public Const flNep% = 2
    Public Const flNiI% = 3
    Public Const flAVK% = 7
    Public Const flHyp% = 14
    Public Const flDFS% = 15
    Public Const flPNP% = 16
    Public Const flLiph% = 21
    Public Const flChar% = 23
    Public Const flUlc% = 26
    Public Const flMFF% = 27 ' höhergradiges Fußsyndrom
    Public flag%(FEZ)
    Public DiagSi%(FEZ)
    Public DiagNamen$(FEZ)
    Public ic$(FEZ, maxIcd)
    Public un%(FEZ) ' Unterdiagnosen
' Begleiterkrankungen
    Public Const BEZ% = 4, bmaxIcd% = 4
    Public bflag%(BEZ)
    Public bDiagsi$(BEZ)
    Public bDN$(BEZ)
    Public bic$(BEZ, bmaxIcd - 1)
    Public bun%(BEZ) ' Unterdiagnosen
    Public Const bflHyt% = 1
    Public Const bflRau% = 4
'Public Const KVÄDatei$ = aVerz + "\KV-Ärzte.mdb" ' u:\Anamnese
'Public KVÄDatei$
Public labPos% ' Laborwert positiv: -1 = positiv, 0 = negativ, 1 = V.a. positiv
'Public fpos&
'Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
'Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, lpData As Any, lpcbData&)         ' Note that if you declare the lpData parameter$, you must pass it By Value.
'Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&) ' unfertig, sollte ein Formular automatisch machen
'Declare Function RegEnumKey& Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey&, ByVal dwIndex&, ByVal lpName$, ByVal cbName&)
'Declare Function RegEnumValue& Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey&, ByVal dwIndex&, ByVal lpValueName$, lpcbValueName&, ByVal lpReserved&, lpType&, ByVal lpData$, lpcbData&)
' Deklarieren der nötigen API-Routinen:
' ... für das Excel-Beispiel
'Declare Function FindWindow Lib "user32" Alias _
'"FindWindowA" (ByVal lpClassName$, _
'                    ByVal lpWindowName&)&
Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
' für das Winword-Beispiel
Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
                    
'Store these declarations in the form wh
' ich is making the calls to Word
'Public Wapp As Object 'Word.Application
Dim WordWasNotRunning As Boolean ' Flag For final word unload
'The main sub which detects word, and cr
' eates the object, or starts a new Word
Declare Sub sleep Lib "Kernel32.dll" Alias "Sleep" (ByVal SleepTime&)
Public Declare Function sndPlaySound32& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName$, ByVal uFlags&)
Public Const Titel$ = "KV-Ärzte raussuchen lassen"
'Public WinDir$
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
Declare Function OpenClipboard& Lib "user32" (ByVal hWnd&)
Declare Function CloseClipboard& Lib "user32" ()
Declare Function GetClipboardData& Lib "user32" (ByVal wFormat&)
Declare Function GlobalLock& Lib "kernel32" (ByVal hMem&)
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem&)
Declare Function lstrcpy& Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any)
Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer$, ByVal nSize&)
 Declare Function RegEnumKey& Lib "advapi32.dll" _
                 Alias "RegEnumKeyA" ( _
                 ByVal Hkey&, _
                 ByVal dwIndex&, _
                 ByVal lpName$, _
                 ByVal cbName&)
'Public rest As DAO.Recordset
Dim Tabl ' Labortabelle, Funktionsübergreifend
Dim UStumm% ' Knöpfe neben den Diagnosen in Position setzen
Public Const laborAbfr$ = "select n.Pat_ID AS Pat_ID,n.ZeitPunkt AS ZeitPunkt,n.FertigStGrad AS FertigStGrad,n.Abkü AS Abkü,l.Langtext AS Langtext,n.Wert AS Wert,n.Einheit AS Einheit,k.Kommentar AS Kommentar,n.AbsPos AS AbsPos,n.AktZeit AS AktZeit from (laborlangtext l inner join (laborkommentar k inner join laborneu n on ((k.KommentarVW = n.KommentarVW))) on ((l.LangtextVW = n.LangtextVW)))"

'Public lies.obmysql%
#If False Then
Function ICDZeit() ' kommt nirgends vor
' Dim q As QueryDef ', raNa As DAO.Recordset
 Dim rAna As New ADODB.Recordset
 Const LDeP$ = FName
 On Error GoTo fehler
' Call dtbInit
 DoCmd.Close acForm, LDeP, acSaveYes
 DoCmd.Close acQuery, LDeP, acSaveYes
' DoCmd.DeleteObject acQuery, LDeP
' SELECT [Pat_id], [Nachname], [Vorname], [Diabetestyp], [Diabetes seit], "       " AS ICD, true AS ob, "                 " AS Begr1, "                 " AS Begr2, "                 " AS Begr3, "                 " AS Begr4 FROM _Anamnesebogen;
 DtbQuerydefsDelete LDeP
' Set raNa = TabÖff("Anamnesebogen")
 rAna.Open "select * from anamnesebogen", DBCn, adOpenDynamic, adLockReadOnly
 If Not rAna.BOF Then
  Do While Not rAna.EOF
'   raNa.Edit
   If rAna!Größe > 3 Then rAna!Größe = rAna!Größe / 100
   rAna!ob = 0
   If rAna!Größe <> 0 Then
    If rAna!Gewicht * IIf(rAna!Gewicht < 3, 100, 1) / rAna!Größe / rAna!Größe >= 29 And rAna!Gewicht / rAna!Größe / rAna!Größe <> 0 And InStrB(rAna!Diagnosen, "E66.9") = 0 Then
     rAna!ob = -1
    End If
   End If
   rAna.Update
   rAna.Move 1
  Loop
 End If
 '  order by gewicht/größe/größe*10000 desc
 Call DtbCreateQueryDef(LDeP, "select * from (select Pat_id, Nachname + "", "" + Vorname as Name,Diabetestyp,[Diabetes seit], ""E66.9"" as ICD,ob,round(gewicht * iif(gewicht<3,100,1)/größe/größe,2) as Begr1,"""" as begr2, """" as begr3, """" as begr4 from anamnesebogen) order by ob")
 DoCmd.OpenForm LDeP
 Forms(LDeP).Caption = "ICD-Auswahl Übergewicht"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ICDZeit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Labor_Click(frm As Form)
#End If

Function DtbCreateQueryDef$(QName$, sql$)
 If sql <> "" Then
 If Not LVobMySQL Then
  Dim cat As New ADOX.Catalog
  Dim Cmd$ ' As New ADODB.Command
  Dim V As ADOX.view
'  Set Cmd.ActiveConnection = DBCn
  Set cat.ActiveConnection = DBCn
  On Error GoTo fehler
'  Set Cmd = Nothing
'  Cmd.CommandText =
  Cmd = sql
  Cmd = Replace(Cmd, "`", "")
  Cmd = Replace(Cmd, "date_add", "")
  Cmd = Replace(Cmd, "date_sub", "")
  Cmd = Replace(Cmd, ",interval", "-")
  Cmd = Replace(Cmd, ", interval", "+")
  Cmd = Replace(Cmd, "day)", ")")
  Cmd = Replace(Cmd, "month)", "*30)")
  Cmd = Replace(Cmd, "cast(", "cdate(int(")
  Cmd = Replace(Cmd, "str_to_date(", "cdate(")
  Cmd = Replace(Cmd, ",'%d.%m.%Y'", "")
  Cmd = Replace(Cmd, "as date", ")")
  Cmd = Replace(Cmd, "_utf8", " ")
  Cmd = Replace(Cmd, "concat", "")
  Cmd = Replace(Cmd, "substr(", "mid$(")
  Cmd = Replace(Cmd, "¡", " & ")
  Cmd = Replace(Cmd, "intaccdatemy", "int")
  Cmd = Replace(Cmd, "intacc", "int")
  Cmd = Replace(Cmd, "selectmy", "((")
  Cmd = Replace(Cmd, "divmy", "/")
  Cmd = Replace(Cmd, "to_days", "int")
'  Cmd = Replace(Replace(Cmd, "cast(", ""), " as date)", "")
  Cmd = Replace(Replace(Cmd, "subdate(", "("), "adddate(", "(")
  
 '  Call cat.Views.Append(vn(i), Cmd)
  Cmd = Replace(Cmd, "str_to_date(", "")
  Cmd = Replace(Cmd, ",'#%m/%d/%Y#')", ")")
  Cmd = Replace(Cmd, ",""#%m/%d/%Y#"")", ")")
  Call Shell(App.path & "\viewserst.exe """ & Forms(0).dbv.Datei & """ """ & QName & """ """ & Cmd & """")
' hier ist wegen eines Fehlers in ADOX wirklich noch DAO nötig
#If False Then
  Dim q As dao.QueryDef
  Dim QTb As dao.Database
  Set QTb = dao.OpenDatabase(Forms(0).dbv.Datei)
  On Error Resume Next
'  Call cat.Views.Delete(QName)
  QTb.QueryDefs.Delete QName
  Call QTb.CreateQueryDef(QName, Cmd) '.CommandText)
  Set QTb = Nothing
#End If
 Else
  Cmd = sql
  Cmd = Replace(Cmd, "selectmy", "select")
  Cmd = Replace(Cmd, "divmy", "div")
  Cmd = Replace(Cmd, "¡", ",")
  Cmd = Replace(Cmd, "intaccdatemy", "date")
  Cmd = Replace(Cmd, "intacc", "")
  Cmd = Replace(Cmd, "cdate(", "(")
  Cmd = Replace(Cmd, "first(", "(")
  Call DBCn.Execute("drop view " & ifexists & " `" & QName & "`;")
  Call DBCn.Execute("CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`" & Forms(0).dbv.uid & "`@`%` SQL SECURITY DEFINER VIEW `" & QName & "` AS " & Cmd)
 End If
 DtbCreateQueryDef = Cmd
 End If ' sql = ""
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DtbCreateQueryDef/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DtbCreateQueryDef
Function do_Labor_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 'sql = "SELECT dl.[Pat_ID] AS Pat_id, [Namen].[Nachname] AS Nachname, [Namen].[Vorname] AS Vorname, dl.[ZeitPunkt], dl.[DokName], dl.[AbsPos], dl.[AktZeit], dl.[DokPfad], dl.[dokgroe], [Dokumente abgehakt].[abgehakt] " & _
 "FROM namen INNER JOIN ((select * from Dokumente where dokname like ""*labor*"") AS dl LEFT JOIN [Dokumente abgehakt] ON dl.[DokPfad]=[Dokumente abgehakt].[DokPfad]) ON [Namen].[Pat_ID]=[Dl].[Pat_ID] " & _
 "ORDER BY [Namen].[Nachname], [Namen].[Vorname], dl.[ZeitPunkt];"
 sql = "SELECT dl.Pat_ID AS Pat_id, Nachname, Vorname, ZeitPunkt, dl.DokName, dl.AbsPos, dl.AktZeit, dl.DokPfad, dokgroe, abgehakt " & _
 "FROM (Dokumente as dl LEFT JOIN [Dokumente abgehakt] as da ON dl.DokPfad=da.DokPfad) left join namen as na ON na.Pat_ID=dl.Pat_ID " & _
 "where DokName Like ""*labor*"" and dl.pat_id = " + CStr(frm!Pat_id) + _
 " ORDER BY Nachname, Vorname, ZeitPunkt desc;"

 Call DtbCreateQueryDef("LaborDokumente eP", sql)
 DoCmd.OpenForm "LaborDokumente eP"
 Forms![LaborDokumente ep].Caption = "LaborDokumente zu " & IIf(IsNull(frm!NachName), frm!Pat_id, frm!NachName) & " " & IIf(IsNull(frm.Vorname), "", frm.Vorname)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Labor_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Labor_Click(frm As Form)
#If mitab Then
Function do_DokDown(frm As Form) ' kommt nirgends vor
 Dim Oneu$, rDok As New ADODB.Recordset
 On Error GoTo fehler
' Call dtbInit
 If PcDokPfad = "" Then Call getDokPfad
 If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
 Oneu = frm.NachName + "_" + frm.Vorname + "_" + Format$(frm.GebDat, "dd/mm/yyyy")
 Call VerzPrüf("p:\" + Oneu)
 sql = "select """ + frm.NachName + """ as nachname,""" + frm.Vorname + """ as vorname, * from Dokumente where dokname not like ""*labor*"" and dokname not like ""*heckliste*"" and dokname not like ""*namnese*"" and pat_id = " + CStr(frm!Pat_id) + " order by zeitpunkt desc"
' Set rDok = Dtb.OpenRecordset(sql, dbOpenDynaset)
 rDok.Open sql, DBCn, adOpenStatic, adLockReadOnly
 On Error Resume Next
 Do While Not rDok.EOF
  FSO.CopyFile "" + Replace(Replace(LCase(rDok!DokPfad), "$\turbomed\dokumente", PcDokPfad), "^", "") + "", "" + "p:\" '+ Oneu + "\" + rDok!DokName + ""
  rDok.Move 1
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_DokDown/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_DokDown
#End If

Function do_Dokumente_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 sql = "select """ + frm.NachName + """ as nachname,""" + frm.Vorname + """ as vorname, * from " & kla & QMdbAkt & klz & ".Dokumente where dokname not like ""*labor*"" and dokname not like ""*heckliste*"" and dokname not like ""*namnese*"" and not (dokname like ""*BZ*"" or dokname like ""*RR*"" or dokname like ""*blutdr*"") and pat_id = " & CStr(frm!Pat_id) + " order by zeitpunkt desc"
 Call DtbCreateQueryDef("LaborDokumente eP", sql)
 DoCmd.OpenForm "LaborDokumente eP"
 Forms![LaborDokumente ep].Caption = "Sonstige Dokumente zu " + frm.NachName & " " & frm.Vorname
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Dokumente_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Dokumente_Click(frm As Form)
Function do_Briefe_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 sql = "select """ + frm.NachName + """ as nachname,""" + frm.Vorname + """ as vorname, pfad as DokPfad, name as DokName, art as DokArt, * from [" + QMdbAkt + "].Briefe where pat_id = " + CStr(frm!Pat_id) + " order by zeitpunkt desc"
 Call DtbCreateQueryDef("LaborDokumente eP", sql)
 DoCmd.OpenForm "LaborDokumente eP", acNormal, , "17=17"
 Forms![LaborDokumente ep].Caption = "Briefe zu " + frm.NachName & " " & frm.Vorname
 ' irfan
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Briefe_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Briefe_Click(frm As Form)
Function do_BZKurven_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 Call DtbCreateQueryDef("LaborDokumente eP", "select """ + frm.NachName + """ as nachname,""" + frm.Vorname + """ as vorname, * from [" + QMdbAkt + "].Dokumente where (dokname like ""*BZ*"" or dokname like ""*RR*"" or dokname like ""*blutdr*"") and pat_id = " + CStr(frm!Pat_id) + " order by zeitpunkt desc")
 DoCmd.OpenForm "LaborDokumente eP", acNormal, , "17=17"
 Forms![LaborDokumente ep].Caption = "Blutzuckerkurven zu " + frm.NachName & " " & frm.Vorname
 ' irfan
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_BZKurven_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Briefe_Click(frm As Form)
Function do_AugenBefunde_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 Call DtbCreateQueryDef("LaborDokumente eP", "select """ + frm.NachName + """ as nachname,""" + frm.Vorname + """ as vorname, zeitpunkt, aktzeit, dokname, dokpfad, dokart, pat_id from [" + QMdbAkt + "].Dokumente where (dokname like ""*augenb[!l]*"" or dokname like ""*augen[aä]rzt*"" or dokname like ""*aa*"") and pat_id = " + CStr(frm!Pat_id) + _
                                                " union select """ + frm.NachName + """ as nachname,""" + frm.Vorname + """ as vorname, zeitpunkt, aktzeit, inhalt as DokName, art as dokpfad, art as DokArt, pat_id from [" + QMdbAkt + "].eintraege where (inhalt like ""*augenb[!l]*"" or inhalt like ""*augen[aä]rzt*"" or inhalt like ""* aa[!g]*"") and pat_id = " + CStr(frm!Pat_id) + " order by zeitpunkt desc")
 DoCmd.OpenForm "LaborDokumente eP", acNormal, , "17=17"
 Forms![LaborDokumente ep].Caption = "Augendokumente zu " + frm.NachName & " " & frm.Vorname
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_AugenBefunde_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Function DtbQuerydefsDelete(view$)
 Dim cat As New ADOX.Catalog
 cat.ActiveConnection = DBCn
 Call cat.Views.Delete(view)
End Function ' DtbQuerydefsDelete
Function do_Click_Vorbereit()
 Const LDeP$ = "LaborDokumente eP"
 On Error GoTo fehler
#If obdao Then
 Call dtbInit
#End If
 DoCmd.Close acForm, LDeP, acSaveYes
 DoCmd.Close acQuery, LDeP, acSaveYes
' DoCmd.DeleteObject acQuery, LDeP
 DtbQuerydefsDelete LDeP
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Click_Vorbereit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Function do_HAC(HA_Ausw$, nr%)
 ReDim Preserve HANrBf(2)
 If HA_Ausw Like "##/#####*" Then
  If Replace(HA_Ausw, "/", "") <> HANrBf(nr) Then
   HANrBf(nr) = Replace(HA_Ausw, "/", "")
   Call Sound(WinDir + "\media\Windows XP-Batterie niedrig.wav")
  End If
 End If
End Function
Function dodo_u_Click(frm As AnBog, nr)
 Dim rfDE As New ADODB.Recordset, i&
 Dim rAF&
 On Error GoTo fehler
 If UStumm Then Exit Function
' Debug.Print frm.Controls("U" + CStr(nr)).Value
' rfDE.Seek "=", frm!Pat_id, MDIICD(Nr), MDIDiag(Nr)
 Dim icdlike$
 If MDIICD(nr) <> "" Then
  If MDIICD(nr) Like "E1*" Or MDIICD(nr) Like "L89*" Then
   icdlike = Left(MDIICD(nr), InStr(MDIICD(nr), ".")) & ".*"
  Else
   icdlike = MDIICD(nr)
   For i = 1 To 2
    If Right$(icdlike, 1) = "V" Then icdlike = Left(icdlike, Len(icdlike) - 1)
    If Right$(icdlike, 1) = "Z" Then icdlike = Left(icdlike, Len(icdlike) - 1)
   Next i
   icdlike = icdlike & "[VZ]?"
  End If
  Call rfDE.Open("select * from fuerdiagexp where pat_id = " & frm.adoRS!Pat_id & " and icd rlike '" & icdlike & "' and diagnose like '%" & LTrim$(Replace(LTrim$(Replace(MDIDiag(nr), "V.a.", "")), "Z.n.", "")) & "'", DBCn, adOpenDynamic, adLockOptimistic)
  UStumm = True
  frm.Va = 0
  frm.Zn = 0
  frm.Xtra = ""
  If frm.Controls("vOptionB")(nr).Value <> 0 Then
   If rfDE.BOF Then
    If MDIICD(nr) Like "*V*" Then frm.Va = 1 Else If MDIICD(nr) Like "*Z*" Then frm.Zn = 1
   Else
    If rfDE!ICD Like "*V*" Then frm.Va = 1 Else If rfDE!ICD Like "*Z*" Then frm.Zn = 1
   End If
   For i = frm.Xtra.ListCount - 1 To 0 Step -1
    frm.Xtra.RemoveItem i
   Next i
   If MDIICD(nr) Like "E1*" Then
    frm.Xtra.AddItem ".01 mit Koma"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 0
    frm.Xtra.AddItem ".11 mit Ketoazidose"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 1
    frm.Xtra.AddItem ".21 mit Nierenkomplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 2
    frm.Xtra.AddItem ".31 mit Augenkomplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 3
    frm.Xtra.AddItem ".41 mit neurologischen Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 4
    frm.Xtra.AddItem ".51 mit peripheren vaskulären Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 5
    frm.Xtra.AddItem ".61 mit sonstigen näher bezeichneten Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 6
    frm.Xtra.AddItem ".71 mit multiplen Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 7
    frm.Xtra.AddItem ".81 mit nicht näher bezeichneten Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 8
    frm.Xtra.AddItem ".91 ohne Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 9
   ElseIf MDIICD(nr) Like "L89*" Then
    frm.Xtra.AddItem ".17 Wagner 0 Ferse"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 0
    frm.Xtra.AddItem ".18 Wagner 0 ohne Ferse"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 1
    frm.Xtra.AddItem ".27 Wagner 1 Ferse (oberfl., Infektion nur im Wundbereich)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 2
    frm.Xtra.AddItem ".28 Wagner 1 ohne Ferse (oberfl., Infektion nur im Wundbereich)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 3
    frm.Xtra.AddItem ".37 Wagner 2 Ferse (tief, gelenknah)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 4
    frm.Xtra.AddItem ".38 Wagner 2 ohne Ferse (tief, gelenknah)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 5
    frm.Xtra.AddItem ".47 Wagner 3 Ferse (Teilnekrose)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 6
    frm.Xtra.AddItem ".48 Wagner 3 ohne Ferse (Teilnekrose)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 7
    frm.Xtra.AddItem ".97 Wagner nicht näher bezeichnet Ferse"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 8
    frm.Xtra.AddItem ".98 Wagner nicht näher bezeichnet ohne Ferse"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 9
   End If
  End If
  UStumm = False
  If rfDE.BOF And frm.Controls("vOptionB")(nr).Value <> 0 Then
   If MDIDiag(nr) = "" Then Stop
   DBCn.Execute "insert into fuerdiagexp(pat_id,name,icd,diagnose,nurquart) values(" & frm.adoRS!Pat_id & ",'" & frm.adoRS!GesName & "','" & MDIICD(nr) & "','" & MDIDiag(nr) & "'," & IIf(MDIICD(nr) = "O24.4" Or (MDIICD(nr) Like "L89*" And Not MDIICD(nr) Like "L89.1*"), 1, 0) & ")"
  ElseIf frm.Controls("vOptionB")(nr).Value = False Then
   DBCn.Execute "delete from fuerdiagexp where pat_id = " & frm.adoRS!Pat_id & " and icd = '" & rfDE!ICD & "' and diagnose = '" & rfDE!Diagnose & "'", rAF
  End If
 End If ' MDIICD(nr) <> "" Then
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodo_u_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodo_u_Click
Function do_Diagnosen_Reset(Optional frm As Form)
 Dim rfDE As New ADODB.Recordset, rfDEA As New ADODB.Recordset, F As ADODB.Field
 Dim archiviert As Date
 archiviert = Now()
 On Error GoTo fehler
' DoCmd.Close acTable, "fuerDiagExp"
' Set rfDEA = TabÖff("fuerDiagExpArchiv", "ID")
' Set rfDE = TabÖff("fuerDiagExp")
 Call rfDEA.Open("fuerdiagexparchiv", DBCn, adOpenDynamic, adLockOptimistic)
 Call rfDE.Open("fuerdiagexp", DBCn, adOpenDynamic, adLockOptimistic)
 If Not rfDE.BOF Then
  Do While Not rfDE.EOF
   rfDEA.AddNew
'   On Error Resume Next
   For Each F In rfDEA.Fields
    If F.Name <> "ID" Then
     F.Value = rfDE.Fields(F.Name).Value
    End If
   Next F
   rfDEA!archiviert = archiviert
   On Error GoTo rfdeaFehler
   rfDEA.Update
   On Error GoTo fehler
   rfDE.Delete
   rfDE.MoveNext
  Loop
 End If
 If Not frm Is Nothing Then Call knöpfeanpassen(frm)
 Exit Function
rfdeaFehler:
Dim tonRunde%
For tonRunde = 1 To 10
 Call Sound(WinDir + "\media\Windows XP-Standard.wav")
Next tonRunde
 If Err.Number = 3022 Then
  Resume Next
 Else
  MsgBox "Fehler beim Aktualisieren von fuerDiagExpArchiv!"
  MsgBox Err.Description
  Stop
  Resume
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Diagnosen_Reset/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Diagnosen_Reset
#If False Then
Function ÖffneAnamnesebogen()
 On Error GoTo fehler
 Call ÖffneFormular("Anamnesebogen")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ÖffneAnamnesebogen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
#End If
#If False Then
Function ÖffneFremdlabor()
 On Error GoTo fehler
 Call ÖffneFormular("LaborDokumente")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ÖffneFremdlabor/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
#End If
#If False Then
' wird in und für Commandbars benötigt
Function ÖffneFormular(FoName$)
 Dim Tbl
 On Error GoTo fehler
 Call dtbInit
 On Error Resume Next
 For Each Tbl In Dtb.TableDefs
  DoCmd.Close acTable, Tbl.Name, acSaveYes
 Next
 On Error GoTo fehler
 DoCmd.OpenForm FoName$ '"Anamnesebogen nach Name" dtb.Containers(2).Documents(0).Name
 DoCmd.Maximize
 On Error Resume Next
 Forms!Anamnesebogen.Recordset.FindFirst "ID = " + CStr(AktID) ' Wieder alten Datensatz wählen
 On Error GoTo fehler
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ÖffneFormular/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ÖffneFormular
#End If
#If False Then
' wird in und für Commandbars benötigt
Function ÖffneTabelle()
 Dim frm
 On Error Resume Next
 AktID = Forms(Anmnb)!Pat_id
 DoCmd.Close acForm, Anmnb, acSaveYes ' "Anamnesebogen nach Name"
 On Error GoTo fehler
 DoCmd.OpenTable Dtb.TableDefs!Anamnesebogen.Name, acViewDesign, acEdit
 '(dtb.TableDefs!Anamnesebogen!Vorname).SetFocus
' DoCmd.GoToControl "Vorname"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ÖffneTabelle/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
#End If
' wird in und für Commandbars benötigt
Function AnpassForm()
 Dim cb, i%
 Dim PatName$
 On Error GoTo fehler
 Err.Clear
 On Error Resume Next
 PatName = GesNam(Forms!Anamnesebogen!Recordset) ' Forms!anamnesebogen!Nachname & ", " & Forms!anamnesebogen.Vorname
 On Error GoTo fehler
 If Err.Number > 0 Then PatName = ""
 Err.Clear
 With CommandBars!Anamneseblatt.Controls(1)
 For i = 1 To .Controls.Count
  Debug.Print .Controls(i).OnAction
  If .Controls(i).OnAction = "wausgeb" Then
   .Controls(i).Caption = "&Werte für " + PatName + " ausgeben" ' wausgeb
   Exit For
  End If
 Next i
 End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AnpassForm/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
#If False Then
' Toolprogramm, um Allowzerolength zu korrigieren, wird nirgends aufgerufen
Function tabAnp()
Dim Tabl As dao.TableDef, i%  ', fld As Object
On Error GoTo fehler
Set Tabl = Dtb.TableDefs("Anamnesebogen")
For i = 1 To Dtb.TableDefs(1).Fields.Count
 With Dtb.TableDefs(1).Fields(i)
 If Dtb.TableDefs(1).Fields(i).Type = 10 Then
  If Not Dtb.TableDefs(1).Fields(i).AllowZeroLength Then
   Dtb.TableDefs(1).Fields(i).AllowZeroLength = True
  End If
 End If
 End With
Next
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in tabAnp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' tabanp()
#End If
Function do_Datenbank_Aufruf_Click(frm As Form)
 On Error GoTo fehler
 If QMdbAkt = "" Then QMdbAkt = LokPfad(QmdB)
 Call Shell(syscmd(9) + "msaccess.exe " & "" & QMdbAkt + "", vbNormalFocus)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Datenbank_Aufruf_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Datenbank_Aufruf_Click(frm As Form)
' wird aus dem Formularcode aufgerufen
Public Function do_Ausgabe_Click(frm As Form)
'On Error GoTo Err_Ausgabe_Click
 On Error GoTo fehler
 Call GibWerteAus '(frm!Pat_ID)

'    Call Shell("NOTEPAD.EXE ""u:\Ergeb.txt""", 1)

 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Ausgabe_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
#If mitab Then
' wird aus dem Formularcode aufgerufen
Public Function do_Datenquelle_Change(frm As Form)
 Dim NachName$, Vorname$, GebDat$, altpat_id
 On Error GoTo fehler
 Select Case frm.Name
  Case "Anamnesebogen"
   NachName = ""
   On Error Resume Next
   NachName = frm.Recordset!NachName
   Vorname = frm.Recordset!Vorname
   GebDat = frm.Recordset!GebDat
   frm.RecordSource = frm.Datenquelle
   On Error GoTo fehler
   If NachName <> "" Then
    On Error Resume Next
    frm.Recordset.FindFirst "NachName = '" + NachName + "' AND vorname = '" + Vorname + "' AND format$(gebdat,""dd/mm/yyyy"") = '" + GebDat + "'"
    On Error GoTo fehler
   End If
  Case "LaborDokumente"
   altpat_id = ""
   On Error Resume Next
   altpat_id = CStr(frm.Recordset!Pat_id)
   frm.RecordSource = frm.Datenquelle
   On Error GoTo fehler
   If altpat_id <> "" Then
    On Error Resume Next
    frm.Recordset.FindFirst "pat_id = " & altpat_id
    On Error GoTo fehler
   End If
 End Select
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Datenquelle_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function DqC_U(frm As Form)
 On Error GoTo fehler
 frm.Datenquelle = "Anamnesebogen Unausgefüllte"
Call do_Datenquelle_Change(frm)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DqC_U/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DqC_U(frm As Form)
Public Function DqC_UNZ(frm As Form)
 On Error GoTo fehler
 frm.Datenquelle = "Anamnesebogen Unausgefüllte nach Zeit"
 Call do_Datenquelle_Change(frm)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DqC_UNZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DqC_UNZ(frm As Form)
Public Function DqC_A(frm As Form)
 On Error GoTo fehler
 frm.Datenquelle = "Anamnesebogen alle Datensätze"
 Call do_Datenquelle_Change(frm)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DqC_A/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DqC_A(frm As Form)
#End If
#If mitab Then
' wird aus dem Formularcode aufgerufen
Public Function do_Form_Open(Cancel%, frm As Form)
#If False Then
 Dim rs As dao.Recordset
 Dim q As dao.QueryDef
 Dim geeignet%
 On Error GoTo fehler
 If QMdbAkt = "" Then QMdbAkt = LokPfad(QmdB)
 Select Case frm.Name
  Case "LaborDokumente eP"
   frm.RecordSource = "select * from [" + QMdbAkt + "].[" + frm.Name + "]"
   Exit Function
  Case "Checklisten"
   frm.RecordSource = "select * from [" + QMdbAkt + "].[" + frm.Name + "]"
   Exit Function
  Case "Augenbefunde"
   frm.RecordSource = "select * from [" + QMdbAkt + "].[" + frm.Name + "]"
  Case "Fremdlabor altes Konzept"
   frm.RecordSource = "select * from [" + QMdbAkt + "].[" + "Fremdlabor" + "]"
  Case "LaborDokumente"
   frm.RecordSource = "select * from [" + QMdbAkt + "].[" + "LaborDokumente fAB" + "]"
 End Select
 
 ' aus do_Fremdlabor_Form_Open
 Select Case frm.Name
  Case "Augenbefunde", "Fremdlabor altes Konzept"
   frm!FDatenquelle.RowSource = ""
   For Each q In Dtb.QueryDefs
    If InStr(q.Name, frm.Name) = 1 Then
     frm!FDatenquelle.RowSource = frm!FDatenquelle.RowSource + q.Name + ";"
    End If
   Next q
   Exit Function
 End Select
 ' hier sind also noch "Anamnesebogen", "Checklisten","LaborDokumente"
 frm.Datenquelle.RowSource = ""
 Call dtbInit
 For Each q In Dtb.QueryDefs
  geeignet = 0
  Select Case frm.Name
   Case "Anamnesebogen"
     If InStr(q.Name, "Anamnese") = 1 Then geeignet = 1
     frm.[Datenbank-Aufruf].Caption = QMdbAkt
   Case "LaborDokumente"
     If InStr(q.Name, "LaborDokumente") = 1 Then geeignet = -1
  End Select
  If geeignet Then
   frm.Datenquelle.RowSource = frm.Datenquelle.RowSource + q.Name + ";"
  End If
 Next q
' For Each q In CurrentDb.QueryDefs
'  geeignet = 0
'  Select Case frm.name
'   Case "Anamnesebogen"
'     If InStr(q.name, "Anamnese") = 1 Then geeignet = 1
'     frm.[datenbank-Aufruf].Caption = qmdbakt
'   Case "LaborDokumente"
'     If InStr(q.name, "LaborDokumente") = 1 Then geeignet = -1
'  End Select
'  If geeignet And instrb(frm.Datenquelle.RowSource, q.name) = 0 Then
'   frm.Datenquelle.RowSource = frm.Datenquelle.RowSource + q.name + ";"
'  End If
' Next q

' If instrb(frm.Name, "Anamnesebogen") > 0 And rfDE Is Nothing Then
'  Set rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
' End If

 'If rESt Is Nothing Then
' Set rest = TabÖff("Einstellungen", "Formular")
' On Error Resume Next
' rest.Seek "=", frm.Name
' If Not rest.NoMatch Then
'  frm.Datenquelle = rest![Abfrage für Formular]
'  frm.RecordSource = frm.Datenquelle
''  frm.Recordset.FindFirst "pat_ID = " + 97 'CStr(rESt![id für formular]) ' geht irgendwie nicht (mehr?)
'  DoCmd.GoToRecord acDataForm, "Anamnesebogen", acGoTo, rest!DatensatzNr
' End If
' On Error GoTo fehler
' rest.Close
#End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Open/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Open(Cancel%, frm As Form)
#End If
#If mitab Then
' wird aus dem Formularcode aufgerufen
Public Function do_Form_Close(frm As Form)
#If False Then
' If rESt Is Nothing Then
 On Error GoTo fehler
 Set rest = TabÖff("Einstellungen", "Formular")
 If Not rest Is Nothing Then
  rest.Seek "=", frm.Name
  If rest.NoMatch Then
   rest.AddNew
   rest!Formular = frm.Name
  Else
   rest.Edit
  End If
  rest![Abfrage für Formular] = frm.Datenquelle
  On Error Resume Next
  rest![ID für Formular] = frm.Recordset!Pat_id
  rest![DatensatzNr] = frm.Recordset.AbsolutePosition + 1
  rest.Update
  On Error GoTo fehler
  rest.Close
 End If
 'Call übrigeRecordsetsSchließen
 #End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Close/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Close(frm As Form)
#End If
'Function übrigeRecordsetsSchließen()
' On Error GoTo fehler
' If Not rfDE Is Nothing Then
'   rfDE.Close
'   Set rfDE = Nothing
' End If
' If Not aNaüm Is Nothing Then
'   aNaüm.Close
'   Set aNaüm = Nothing
' End If
' If Not rLaU Is Nothing Then
'  rLaU.Close
'  Set rLaU = Nothing
' End If
' If Not rDokab Is Nothing Then
'   rDokab.Close
'   Set rDokab = Nothing
' End If
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in übrigeRecordsetsSchließen/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): End
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function
Public Function do_form_beforeUpdate(frm)
 On Error GoTo fehler
 If Not frm.Recordset.BOF Then
  On Error Resume Next
  altpat_id = frm.Recordset!Pat_id
  On Error GoTo fehler
  ReDim altDS(frm.Recordset.Fields.Count - 1)
  On Error Resume Next
  For dsnr = 0 To frm.Recordset.Fields.Count - 1
   altDS(dsnr) = frm.Recordset.Fields(dsnr)
  Next dsnr
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_form_beforeUpdate/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_form_beforeUpdate(frm)
' wird aus dem Formularcode aufgerufen
Public Function do_Form_AfterUpdate(frm)
 Dim obGleich As Boolean
' If altID = frm!Pat_id Then
 On Error GoTo fehler
  obGleich = True
'  If Not rAnAlt Is Nothing Then
   On Error Resume Next
   If frm.Recordset!Pat_id = altpat_id Then
    For dsnr = 0 To frm.Recordset.Fields.Count - 1
'     If dsNr = 17 Then
'      Debug.Print "frm.Recordset.Fields(dsNr):", frm.Recordset.Fields(dsNr)
'      Debug.Print "altDS(dsNr):", altDS(dsNr)
'     End If
     If frm.Recordset.Fields(dsnr) <> altDS(dsnr) Or _
     (IsNull(frm.Recordset.Fields(dsnr)) And (Not IsNull(altDS(dsnr)))) Or _
     (Not (IsNull(frm.Recordset.Fields(dsnr))) And IsNull(altDS(dsnr))) Then 'frm(rAnAlt.Fields(dsNr).Name) Then
      If frm.Recordset.Fields(dsnr).Name <> "letzte Änderung" Then
       obGleich = False
      End If
      Exit For
     End If
    Next dsnr
    On Error GoTo fehler
'   End If
   If Not obGleich Then
    Call RRParse(frm!Pat_id)
    frm![letzte Änderung] = Now
   End If
  End If
' End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_AfterUpdate/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'do_Form_AfterUpdate(frm)
Function formtest()
 Dim i%
 On Error GoTo fehler
 For i = 1 To Forms.Count
  Debug.Print Forms(i).Name
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formtest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' formtest()
' wird aus dem Formularcode aufgerufen
#If mitab Then
Public Function do_Form_Current_Medarten(frm As Medarten)

End Function
Public Function do_Form_Current_AnBog(frm As AnBog)
 Dim farbe&, i&, tStr$
 Static rAna As ADODB.Recordset
' Set rAnAlt = frm.Recordset
' altPat_id = frm.Recordset!Pat_id
' ReDim altDS(frm.Recordset.Fields.Count - 1)
' For dsNr = 0 To frm.Recordset.Fields.Count - 1
'  altDS(dsNr) = frm.Recordset.Fields(dsNr)
' Next dsNr
' Debug.Print "Form_Current 1:" + format$(Now, "hh:mm:ss")
 Dim ctrl As Control

' Diagnosen
 Dim KZahlNeu%, pnpflNeu%, anpflNeu%, retflNeu%, katflNeu%, bgwflNeu%, _
 angflNeu%, dfsflNeu%, fieflNeu%, oblhNeu%
 If obstumm Then Exit Function
 KZahlNeu = 0
 pnpflNeu = 0
 anpflNeu = 0
 retflNeu = 0
 katflNeu = 0
 bgwflNeu = 0
 oblhNeu = 0
 angflNeu = 0
 dfsflNeu = 0
 fieflNeu = 0
 
 On Error GoTo fehler
 For Each ctrl In frm.Controls
  On Error Resume Next
  If ctrl.BackColor = MarkierFarbe Then ctrl.BackColor = -2147483643
  On Error GoTo fehler
 Next ctrl
' dtbInit
 Set rsNa = Nothing
'  Set rsNa = TabÖff("Anamnesebogen", "Pat_id")
 If frm.adoRS.EOF Or frm.adoRS.BOF Then
  Pat_id = 0
 Else
  Pat_id = frm.adoRS!Pat_id
 End If
 Set rsNa = Nothing
 rsNa.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 If rsNa.BOF Then Exit Function
 On Error Resume Next
 syscmd 4, "Beginn Formularvorbereitung " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 Err.Clear
' tStr = Pat_id
' If Err.Number <> 0 Then
'  MsgBox "Falsche Abfrage: " + frm!Datenquelle
'  MsgBox (CurrentDb.QueryDefs(frm!Datenquelle).sql)
'  MsgBox "Ein Feld muß mit 'pat_id' ohne Vorsätze ansprechbar sein"
'  Stop
' End If
' nzl = vbcr + vblf
' rsNa.Seek "=", pat_id
 If rsNa.BOF Then
  frm.vTextB(147) = "" ' Diagnosen
 Else
'  frm.vTextB(147) = Replace(Replace(rsNa!Diagnosen, vbverticaltab, vbcr + vblf), vbtab, chr$(32))
   frm.vTextB(147) = Replace(Replace(DiagString$(Pat_id), vbVerticalTab, vbCrLf), vbTab, Chr$(32))
 End If
 frm.vTextB(181) = leBe(Pat_id) ' leBeh
 If (IsNull(frm.adoRS!Ther1) Or frm.adoRS!Ther1 = "") And frm.adoRS!Diabetestyp <> "-" Then frm.adoRS!Ther1 = TherUmw(TherArt(Pat_id, -1))
 If (IsNull(frm.adoRS!TherAkt) Or frm.adoRS!TherAkt = "") And frm.adoRS!Diabetestyp <> "-" Then frm.adoRS!TherAkt = TherUmw(TherArt(Pat_id, 0))

' If rafDE Is Nothing Then Call rafDE.Open("select * from fuerdiagexp", DBCn, adOpenDynamic, adLockReadOnly) 'Set rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 If HArst Is Nothing Then Call HArst.Open("select * from listenausgabeuew", DBCn, adOpenDynamic, adLockReadOnly) 'Set HArst = TabÖff("Hausaerzte", "KVNr") ' isempty(harst), falls über dim harst statt dim harst as dao.recordset definiert
 On Error Resume Next
 Dim obHAimDMP%
 obHAimDMP = 0
 frm.vTextB(148) = "" ' HA
 frm.vTextB(176) = "" ' HAFax
 frm.vTextB(169) = "" ' HAimDMP
' Dim HaNr1$, HANr2$
 frm.vComboB(1) = "" ' HA_Auswahl
 frm.vComboB(2) = "" ' HA_auswahl2
 Dim HaErg$()
 frm.vTextB(183) = Üw12(Pat_id, HaErg) ' HANr, neues Feld
 frm.vTextB(175) = HaErg(0, 1) ' HANr2
 ReDim HANrBf(2)
 HANrBf(1) = frm.vTextB(183) ' HANr <- Ersatz schaffen
 HANrBf(2) = frm.vTextB(175) ' HANr2
 frm.vCheckb(15) = (HaErg(0, 0) = "") ' HA1nicht
 frm.vCheckb(16) = (HaErg(0, 1) = "") ' HA2nicht
 If Not IsNull(frm.vTextB(183)) And frm.vTextB(183) <> "" Then ' HANr
'  HArst.FindFirst "ltrim$(KVNr) = ltrim$('" + frm.HANr + "')"

' folgendes geht schlecht in ein Unterprogramm
   Dim teile$()
   teile = Split(HANrBf(1), "_")
       Dim gefunden%
       gefunden = 0
       If UBound(teile) >= 2 Then
'       rHa.Seek "=", teile(0), trim$(teile(1)), trim$(teile(2))
        Set HArst = Nothing
        Call HArst.Open("select * from listenausgabeuew where kvnr = '" & teile(0) & "' and name = '" & teile(1) & "' and vorname = '" & teile(2) & "'", DBCn, adOpenStatic, adLockReadOnly)
        If Not HArst.BOF Then gefunden = True
       End If
       If UBound(teile) = 1 Or (UBound(teile) > 1 And Not gefunden) Then
        gefunden = 0
'       rHa.Seek "=", teile(0), trim$(teile(1))
        Set HArst = Nothing
        Call HArst.Open("select * from listenausgabeuew where kvnr = '" & teile(0) & "' and name = '" & teile(1) & "'", DBCn, adOpenStatic, adLockReadOnly)
        If Not HArst.BOF Then gefunden = True
       End If
       If UBound(teile) = 0 Or (UBound(teile) > 0 And Not gefunden) Then
        gefunden = 0
 '      rHa.Seek "=", teile(0)
        Set HArst = Nothing
        Call HArst.Open("select * from listenausgabeuew where kvnr = '" & teile(0) & "'", DBCn, adOpenStatic, adLockReadOnly)
       End If
      
   If HArst.BOF Or IsNull(HArst!Name) Then
'    If IsNull(HArst!Name) Then
'     HArst.Delete
'    End If
    Select Case UBound(teile)
     Case 2: Call HAlokal(HArst, teile(0), teile(1), teile(2))
     Case 1: Call HAlokal(HArst, teile(0), teile(1))
     Case 0: Call HAlokal(HArst, teile(0))
    End Select
   End If
   If Not HArst.BOF And Not UBound(teile) = -1 Then ' wenn gar nicht gesucht wurde
    frm.vTextB(148) = IIf(IsNull(HArst!Name), "", HArst!Name) + ",T:" + IIf(IsNull(HArst!Telefon), "", HArst!Telefon) & ", " & IIf(IsNull(HArst!strasse), "", HArst!strasse) & " " & IIf(IsNull(HArst!Plz), "", HArst!Plz) & " " & IIf(IsNull(HArst!Ort), "", HArst!Ort) + ", Z:" + IIf(IsNull(HArst!zulg), "", HArst!zulg)
    frm.vTextB(176) = IIf(IsNull(HArst!Fax), "", HArst!Fax)
    obHAimDMP = fobHAimDMP(frm.vTextB(183)) ' HANr
   End If
   HArst.Find "HANr =", LTrim$(frm.vTextB(183))  ' HANr
  frm.vTextB(169) = IIf(obHAimDMP, "Hausarzt im DMP", "Hausarzt nicht im DMP") ' HAimDMP
 End If
' If rFa Is Nothing Then Set rFa = TabÖff("faelle", "aktF")
' rFa.Seek "=", pat_id
 Set raFa = Nothing
' Call raFa.Open("select * from faelle where pat_id = " & frm!Faelle, DBCn, adOpenDynamic, adLockReadOnly)
' frm.adoRS!SchGr = raFa!SchGr
 Select Case frm.adoRS!Versicherungsart ' Scheingruppe
  Case "00"
   farbe = 16711808 ' lila
  Case "24"
   farbe = 65280 ' grün
  Case "41"
   farbe = 12632256 ' hellgrau
  Case "90"
   farbe = 65535 ' gelb
  Case Else
   farbe = -2147483633  ' weiß
 End Select
 frm.vTextB(160).BackColor = farbe ' NachName
 frm.vTextB(2).BackColor = farbe ' Vorname
 frm.vTextB(152).BackColor = farbe ' Titel
 frm.vTextB(153).BackColor = farbe ' Anrede
 If frm.adoRS!Größe = 0 Or frm.adoRS!Gewicht = 0 Then
   farbe = -2147483633 ' weiß
 Else
   On Error Resume Next
   Dim BMI!, Gewicht!, Größe!
   Dim runde%
   runde = 0
   BMI = frm.vTextB(142)
   Gewicht = frm.vTextB(11)
   Größe = frm.vTextB(10)
   Do
    runde = runde + 1
    If BMI = 0 Then Exit Do ' BMI
    If BMI > 8 Then Exit Do ' BMI
    BMI = BMI * 10 ' BMI
    If Gewicht < 10 Then
     Gewicht = Gewicht * 10
    End If
   Loop Until runde >= 10
   runde = 0
   Do
    If BMI = 0 Then Exit Do ' BMI
    If BMI < 80 Then Exit Do ' BMI
'    frm.vtextb(142) = frm.vtextb(142) * 0.1 => geht nicht ,da Berechnungsfeld
    If Gewicht > 250 Then ' Gewicht     <- Ersatz schaffen
       Gewicht = Gewicht * 0.1 ' Gewicht
    ElseIf Größe < 3 Then ' Größe     <- Ersatz schaffen
      Größe = Größe * 10 ' Größe
    End If
    runde = runde + 1
   Loop Until runde >= 10
   Select Case BMI ' BMI
    Case Is < 20
     farbe = 15263952
    Case Is < 25
     farbe = 13619100
    Case Is < 30
     farbe = 11369810
    Case Is < 35
     farbe = 10314363
    Case Is < 40
     farbe = 8673192
    Case Else
     farbe = 5387467
   End Select
   On Error GoTo fehler
 End If
 frm.vTextB(142).BackColor = farbe ' BMI
 frm.vTextB(10).BackColor = farbe ' Größe
 frm.vTextB(11).BackColor = farbe ' Gewicht
 frm.vTextB(12).BackColor = farbe ' Tendenz
 
 syscmd 4, "Formularvorbereitung 1 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 If IsNumeric(frm.adoRS![letztes HbA1c]) Then ' letztes HbA1c
  Select Case Val(frm.adoRS![letztes HbA1c]) ' letztes HbA1c
   Case 0
    frm.vTextB(43).BackColor = -2147483643 ' letztes HbA1c
   Case Is < 6
    frm.vTextB(43).BackColor = 15263952 ' letztes HbA1c
   Case Is <= 6.5
    frm.vTextB(43).BackColor = 13619100 ' letztes HbA1c
   Case Is < 7
    frm.vTextB(43).BackColor = 11369810 ' letztes HbA1c
   Case Is < 8
    frm.vTextB(43).BackColor = 10314363 ' letztes HbA1c
   Case Is < 9
    frm.vTextB(43).BackColor = 8673192 ' letztes HbA1c
   Case Else
    frm.vTextB(43).BackColor = 5387467 ' letztes HbA1c
  End Select
 Else
   frm.vTextB(43).BackColor = -2147483643 ' weiß  ' letztes HbA1c
 End If
 Dim lsql$
 lsql = "select * from labor1 where pat_id = " & Pat_id & " union select * from labor2 where pat_id = " & Pat_id
 Set raLau = Nothing
 raLau.Open "select * from (" & lsql & ") as innen where abkü like '%BNP%' order by zeitpunkt desc", DBCn, adOpenDynamic, adLockReadOnly
 If Not raLau.EOF Then
  frm.ntPro = raLau!Wert
 Else
  frm.ntPro = ""
 End If
 frm.AlbCre = ""
 Set raLau = Nothing
 raLau.Open "select * from (" & lsql & ") as innen where abkü in ('ALBCRE','ALBKRE','ALBQ') order by zeitpunkt desc", DBCn, adOpenDynamic, adLockReadOnly
 Do While Not raLau.EOF
  frm.AlbCre = frm.AlbCre & raLau!Wert & "/"
  raLau.Move 1
 Loop
 
 If IsNull(frm.adoRS![Diabetestyp]) Or frm.adoRS![Diabetestyp] = "" Then ' Diabetestyp
  farbe = 16777215 ' weiß
 Else
  Select Case frm.adoRS![Diabetestyp] ' Diabetestyp
   Case "1"
    farbe = 255 'rot
   Case "2"
    farbe = 33023  ' orange
   Case "S", "s"
    farbe = 12615935 ' rosa
'   Case ""
'    farbe = 16777215 ' weiß
   Case "-"
    farbe = 12632256 ' hellgrau
   Case Else
    farbe = 65280 ' grün
  End Select
 End If
 frm.vTextB(4).BackColor = farbe ' Diabetestyp
 frm.vTextB(5).BackColor = farbe ' [Diabetes seit]
 frm.vTextB(6).BackColor = farbe ' [Tabletten seit]
 frm.vTextB(7).BackColor = farbe ' [Insulin seit]
' Schwangerschaftsfelder ggf. sperren
 Dim SAkt%
 SAkt = -1
 If Not IsNull(frm.adoRS!Anrede) Then
  If Not IsNull(frm.adoRS!GebDat) Then
   If frm.adoRS!Anrede = "Herr" Or Date - CDate(frm.adoRS!GebDat) > 55 * 365 Then
    SAkt = 0
   End If
  End If
 End If
 If SAkt Then
  Call Acti(frm.vTextB(63)) ' frm.adoRS!Schwanger
  Call Acti(frm.vTextB(64)) ' [Schwanger seit]
 Else
  Call inActi(frm.vTextB(63)) ' frm.adoRS!Schwanger
  Call inActi(frm.vTextB(64)) ' [Schwanger seit]
 End If
 If frm.adoRS!Ther1 = "OAD" Then
  Call inActi(frm.vTextB(37)) ' [Spritzstelle früh (B, O, A)]
  Call inActi(frm.vTextB(38))   ' [Spritzstelle mittags]
  Call inActi(frm.vTextB(39)) ' [Spritzstelle abends]
  Call inActi(frm.vTextB(40)) ' [Spritzstelle nachts]
 Else
  Call Acti(frm.vTextB(37)) ' [Spritzstelle früh (B, O, A)]
  Call Acti(frm.vTextB(38))   ' [Spritzstelle mittags]
  Call Acti(frm.vTextB(39)) ' [Spritzstelle abends]
  Call Acti(frm.vTextB(40)) ' [Spritzstelle nachts]
 End If
 If frm.adoRS![BZMessungen selbst] = "n" Or frm.adoRS![BZMessungen selbst] = "-" Then
  Call inActi(frm.vTextB(158))   ' Gerät
  Call inActi(frm.vTextB(49)) ' Aufschreiben
 Else
  Call Acti(frm.vTextB(158))   ' Gerät
  Call Acti(frm.vTextB(49)) ' Aufschreiben
 End If
 If frm.adoRS![Bluthochdruck] = "n" Or frm.adoRS!Bluthochdruck = "-" Then
  Call inActi(frm.vTextB(59)) ' BHD seit
  Call inActi(frm.vTextB(60)) ' BHD beh mit
 Else
  Call Acti(frm.vTextB(59)) ' BHD seit
  Call Acti(frm.vTextB(60)) ' BHD beh mit
 End If
 If frm.adoRS!Herzkrankheit = "n" Or frm.adoRS!Herzkrankheit = "-" Then
  Call inActi(frm.vTextB(75)) ' [Angina pectoris]
  Call inActi(frm.vTextB(76)) ' [Herzinfarkt]
  Call inActi(frm.vTextB(77)) ' [Herzinfarkt wann]
  Call inActi(frm.vTextB(78)) ' [PTCA oder Stent]
  Call inActi(frm.vCheckb(3)) ' [Bypass kardial]
  Call inActi(frm.vTextB(79)) ' [Bypass wann]
  Call inActi(frm.vTextB(80)) ' [Herzschwäche]
  Call inActi(frm.vTextB(81)) ' [Herzkrankheit Beschreibung]
 Else
  Call Acti(frm.vTextB(75)) ' [Angina pectoris]
  Call Acti(frm.vTextB(76)) ' [Herzinfarkt]
  Call Acti(frm.vTextB(77)) ' [Herzinfarkt wann]
  Call Acti(frm.vTextB(78)) ' [PTCA oder Stent]
  Call Acti(frm.vCheckb(3)) ' [Bypass kardial]
  Call Acti(frm.vTextB(79)) ' [Bypass wann]
  Call Acti(frm.vTextB(80)) ' [Herzschwäche]
  Call Acti(frm.vTextB(81)) ' [Herzkrankheit Beschreibung]
 End If
 If frm.adoRS![Sexualstörung] = "n" Or frm.adoRS!Sexualstörung = "-" Then
  Call inActi(frm.vTextB(101)) ' [Sexualstörung seit]
 Else
  Call Acti(frm.vTextB(101)) ' [Sexualstörung seit]
 End If
' Feld mit den Diagnosen der Medikamente füllen
syscmd 4, "Formularvorbereitung 2 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
Dim rsmd As New ADODB.Recordset
Dim hmg$, hypt$, neurp$, autnp$, fetts$, Hsre$, antimyk$, glauk$, cold$, pros$, urä$, hythy$, ostp$, _
    khk$, herzis$, stru$, avk$, pani$, park$, vari$, östr$, antidep$, antidem$, antiep$, antipern$, antiherp$, Antikoag$
    
If Pat_id <> 0 Then
' Set aNaüm = TabÖff("Anamnesebogen", "Pat_id")
' aNaüm.Seek "=", pat_id
' Set aNaüm = Nothing
' aNaüm.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
' Vorbelegungen für (zusätzliche) Diagnosen im Feld mdi
 frm.vTextB(168) = "" ' !MDi
 MDIn = 1
 For i = 0 To UBound(MDIICD)
  MDIICD(i) = ""
  MDIDiag(i) = ""
 Next
 Set fm = frm
 
 Dim TabakStatus%
 TabakStatus = doTabakSt(frm.adoRS!Tabak)
 Select Case TabakStatus
  Case 0
  Case 1
   Call KRAdd(frm, frm.adoRS!Tabak, "Z.n.Nikotinabusus", "F17.1", "F17.1", , , , "vTextB", 104) ' Tabak
  Case 3
   Call KRAdd(frm, frm.adoRS!Tabak, "Z.n.Nikotinabusus vor langem", "F17.1", "F17.1", , , , "vTextB", 104) ' Tabak
  Case 2
   Call KRAdd(frm, frm.adoRS!Tabak, "Nikotinabusus", "F17.1", "F17.1", , , , "vTextB", 104) ' "Tabak"
 End Select
 
 ' andersrum geordnet in EigTest
 'Set md = dtb.OpenRecordset("SELECT Fi.pat_id AS Pat_ID, Fi.zeitpunkt AS Zeitpunkt, Fi.med, medarten.HMG, medarten.Hypt, medarten.Neurp, medarten.AutNP, medarten.Fetts, medarten.Hsre, medarten.AntiMyk, medarten.Glauk, medarten.COLD, medarten.Pros, medarten.Urä, medarten.HyThy, medarten.Ostp, medarten.KHK, medarten.HerzI, medarten.Stru, medarten.AVK, medarten.PanI, medarten.Park, medarten.Vari, medarten.Vari, medarten.Östr, medarten.AntiDep, medarten.AntiDem, FROM [SELECT pat_id, zeitpunkt, (left(feldinh,iif(instrb(feldinh,"" "")>0,instr(feldinh,"" ""),len(feldinh)))) as med from forminhalt where form_abk = ""mp"" and feld = ""medikament""]. AS Fi LEFT JOIN medarten ON Fi.med = medarten.Medikament where fi.pat_id =" + CStr(frm.Pat_ID), dbOpenDynaset)
 sql = "select mp.*, ma.* from medplan mp left join medarten ma on mp.medanfang = ma.medikament where mp.pat_id =" & Pat_id
 Set rsmd = Nothing
 rsmd.Open sql, DBCn, adOpenStatic, adLockReadOnly
' Debug.Print "Form_Current 2:" + format$(Now, "hh:mm:ss")
 If Not rsmd.BOF Then
  Do While Not rsmd.EOF
   Call mdTest(rsmd, "hmg", hmg)
   Call mdTest(rsmd, "hypt", hypt)
   Call mdTest(rsmd, "neurp", neurp)
   Call mdTest(rsmd, "autnp", autnp)
   Call mdTest(rsmd, "fetts", fetts)
   Call mdTest(rsmd, "Hsre", Hsre)
   Call mdTest(rsmd, "antimyk", antimyk)
   Call mdTest(rsmd, "glauk", glauk)
   Call mdTest(rsmd, "cold", cold)
   Call mdTest(rsmd, "pros", pros)
   Call mdTest(rsmd, "urä", urä)
   Call mdTest(rsmd, "hythy", hythy)
   Call mdTest(rsmd, "ostp", ostp)
   Call mdTest(rsmd, "khk", khk)
   Call mdTest(rsmd, "herzi", herzis)
   Call mdTest(rsmd, "stru", stru)
   Call mdTest(rsmd, "avk", avk)
   Call mdTest(rsmd, "pani", pani)
   Call mdTest(rsmd, "park", park)
   Call mdTest(rsmd, "vari", vari)
   Call mdTest(rsmd, "östr", östr)
   Call mdTest(rsmd, "antidep", antidep)
   Call mdTest(rsmd, "antidem", antidem)
   Call mdTest(rsmd, "antiep", antiep)
   Call mdTest(rsmd, "antipern", antipern)
   Call mdTest(rsmd, "antiherp", antiherp)
   Call mdTest(rsmd, "antikoag", Antikoag)
   rsmd.MoveNext
  Loop ' not rsmd.eof
 End If ' not rsmd.bof
' Debug.Print "Form_Current 3:" + format$(Now, "hh:mm:ss")
 syscmd 4, "Formularvorbereitung 3 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 Call KRAdd(frm, neurp, "Polyneuropathie", "G63.2", "G63", "G62", , , "vTextB", 105) ' Weitere Medikation
 If labPos <> 0 Then pnpflNeu = -1
 Call KRAdd(frm, autnp, "Autonome Neuropathie", "G99.0", "G99", , , , "vTextB", 105) ' Weitere Medikation
 If labPos <> 0 Then anpflNeu = True
 Call KRAdd(frm, khk, "Koronare Herzerkrankung", "I25.9", "I25", "I21", , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, herzis, "Herzinsuffizienz", "I50.12", "I50", "", "", "", "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, avk, "AVK", "I73.9", "I74", "I70", "I73", , "vTextB", 105) ' Weitere Medikation
 If labPos <> 0 Then angflNeu = -1
 Call KRAdd(frm, urä, "Niereninsuffizienz", "N18.83", "N19", "N18", , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antimyk, "Pilzinfektion", "B39", "B39", "B35", "B37", , "vTextB", 105) ' Weitere Medikation  ' B49, N77 fehlt noch: Vaginalmykose, B37.2 Onychomykose
 Call KRAdd(frm, hmg, "Hypercholesterinämie", "E78.0", "E78", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, fetts, "Fettstoffwechselstörung", "E78.9", "E78", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, hypt, "Arterielle Hypertonie", "I10.90", "I10", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, Hsre, "Hyperurikämie", "E79.0", "E79", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, glauk, "Glaukom", "H40.9", "H40", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, cold, "COPD", "J42", "J42", "J44", "M06.99", , "vTextB", 105) ' Weitere Medikation ' Cortison auch bei Rheuma
 Call KRAdd(frm, pros, "Prostatahyperplasie", "N40", "N40", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, hythy, "Hyperthyreose", "E05.9", "E05", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, stru, "V.a.Struma (nach Med.)", "E04.9V", "E04", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, ostp, "Osteoporose", "M81.99", "M81", "M82", , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, pani, "Pankreasinsuffizienz", "K86.8", "K86", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, park, "Parkinson", "G20", "G20", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, vari, "Varikose", "I83.9", "I83", "I86", , , "vTextB", 105) ' Weitere Medikation
 GebDat = IIf(IsNull(frm.adoRS!GebDat), 0, frm.adoRS!GebDat)
 Alter = (Date - GebDat) * 2.73792574745373E-03 ' 1/365,24
 If Alter > 45 Then Call KRAdd(frm, östr, "Wechselbeschwerden", "N95.9", "N95", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antidep, "Depression", "F32.9", "F32", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antidem, "Demenz", "F03", "F0", "G30", , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antiep, "Epilepsie", "G40.9", "G40", "G41", , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antipern, "Vitamin-B12-Mangel", "E53.8", "G63.4", "D51", "E53.8", , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antiherp, "Herpes-Infektion", "B00.9", "B00", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, Antikoag, "Antikoagulantien-Therapie", "Z92.1", , , , "vTextB", 105) ' Weitere Medikation
 
 syscmd 4, "Formularvorbereitung 4 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
' Anamnesebogen
 Dim NPGru$
 NPGru = LegNPFest(Pat_id, frm)
 
 If NPGru Like "->*" Then Call KRAdd(frm, Mid$(NPGru, 3), "Polyneuropathie", "G63.2", "G62", "G63")
 If labPos <> 0 Then pnpflNeu = -1
'  If NPGrund = "KW" Or NPGrund = "MER" Then
 If NPGru Like "V.*" Then Call KRAdd(frm, Mid$(NPGru, 3), "V.a. Polyneuropathie", "G63.2V", "G62", "G63")
 Call LegNPFest(Pat_id, frm)
 
 Dim verform$, BewEinsch$, WSwo%
 verform = IIf(IsNull(frm.adoRS!Verformungen), "", LCase(frm.adoRS!Verformungen))
 BewEinsch = IIf(IsNull(frm.adoRS!Bewegungseinschränkungen), "", LCase(frm.adoRS!Bewegungseinschränkungen))
 Dim obBewEinsch%, obverform%
 obverform = 0
 obBewEinsch = 0
 If Left(verform, 1) = "j" Then
  obverform = -1
  Call doGilbE(frm, "M14", "vTextB", 92) ' Bewegungseinschränkungen
 End If
 If obNein(BewEinsch) = 0 Then
  obBewEinsch = -1
  Call doGilbE(frm, "M14", "vTextB", 99)
 End If
 If obverform Or obBewEinsch Then '(BewEinsch <> "" And BewEinsch <> "n" And BewEinsch <> "-" And instrb(BewEinsch, "WS-Syndrom") = 0 And (Not BewEinsch Like "o.B?") And (Not BewEinsch Like "o. B?")) Then
  Call KRAdd(frm, "Fragebogen", "V.a. Diab. Arthropathie", "M14.2V", "M14") ' Bewegungseinschränkungen
  If InStrB(BewEinsch, "Knie") <> 0 Or InStrB(BewEinsch, "Gonarth") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "V.a. Kniegelenksarthrose", "M17.9V", "M17", , , , "vTextB", 99) ' Bewegungseinschränkungen
  End If
  If InStrB(BewEinsch, "Hüft") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Hüftgelenksarthrose", "M16.9", "M16", , , , "vTextB", 99) ' Bewegungseinschränkungen
  End If
  If InStrB(BewEinsch, "Schulter") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Schultergelenksarthrose", "M19.21", "M19", , , , "vTextB", 99) ' Bewegungseinschränkungen
  End If
  WSwo = 0
  If InStrB(BewEinsch, "lws") <> 0 Or InStrB(BewEinsch, "Lend") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "V.a. LWS-Syndrom", "M47.96V", "M47.96", "M54.16", "M47.8", "M53", "vTextB", 99) ' Bewegungseinschränkungen
    WSwo = 1
  End If
  If InStrB(BewEinsch, "bws") <> 0 Or InStrB(BewEinsch, "Brustwirbel") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "V.a. BWS-Syndrom", "M47.94V", "M47.94", "M54.14", "M47.8", "M53", "vTextB", 99) ' Bewegungseinschränkungen
    WSwo = 2
  End If
  If InStrB(BewEinsch, "hws") <> 0 Or InStrB(BewEinsch, "Halswirbel") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "V.a. HWS-Syndrom", "M47.92V", "M47.92", "M54.2", "M47.8", "M53", "vTextB", 99) ' Bewegungseinschränkungen
    WSwo = 3
  End If
  If InStrB(BewEinsch, "WS") <> 0 And WSwo = 0 Then
    Call KRAdd(frm, "Fragebogen", "Wirbelsäulen-Syndrom", "M47.99V", "M47", "M54", "M53", , "vTextB", 99) ' Bewegungseinschränkungen
    WSwo = 3
  End If
  If InStrB(BewEinsch, "Bandsch") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Diskopathie", "M51.8V", "M51", , , , "vTextB", 99) ' Bewegungseinschränkungen
  End If
 End If
 Dim hkGr$
 hkGr = hkGrund(Pat_id)
 If hkGr <> "" Then
  Call KRAdd(frm, hkGr, "Hyperkeratose", "L84", "L85.9", "L84", , , "vTextB", IIf(hkGr = "Frgbg.", 91, 109)) ' Druckstellen / Hyperkeratosen
 End If ' left(druckst, 1) = "j" Then
 
 Dim EntlSt$, autGrund$, obAutNP, HarnB
 obAutNP = 0
 HarnB = 0
 EntlSt = IIf(IsNull(frm.adoRS![Entleerungsstörungen Magen]), "", LCase(frm.adoRS![Entleerungsstörungen Magen]))
 If Left(EntlSt, 1) = "j" Then
  obAutNP = obAutNP - 1
  Call doGilbE(frm, "G99", "vTextB", 95) ' Entleerungsstörungen Magen
 End If
 EntlSt = IIf(IsNull(frm.adoRS![Entleerungsstörungen Harnblase]), "", LCase(frm.adoRS![Entleerungsstörungen Harnblase]))
 If Left(EntlSt, 1) = "j" Then
  obAutNP = obAutNP - 1
  Call doGilbE(frm, "G99", "vTextB", 96) ' Entleerungsstörungen Harnblase
  HarnB = -1
 End If
 EntlSt = IIf(IsNull(frm.adoRS![Sexualstörung]), "", LCase(frm.adoRS![Sexualstörung]))
 If Left(EntlSt, 1) = "j" Then
  Call doGilbE(frm, "G99", "vTextB", 100) ' Sexualstörung
  obAutNP = obAutNP - 1
 End If
 If obAutNP <> 0 And Not (obAutNP And HarnB) Then ' Nur Harnblase lieber nicht
  Call KRAdd(frm, "Anamn", IIf(obAutNP = -1, "V.a.", "") + "Autonome Neuropathie", "G99.0" + IIf(obAutNP = -1, "V", ""), "G99")
  If labPos Then anpflNeu = True
 End If
 
 oblhNeu = obLH(Pat_id, frm)
 
 If oblhNeu Then
  Call KRAdd(frm, "Befund", "Liphypertrophien", "L94.8", "L94", "E88")
 End If
 
 Dim ulcer$
 ulcer = IIf(IsNull(frm.adoRS![Ulcera]), "", LCase(frm.adoRS!Ulcera))
 ulcer = Trim$(ulcer)
 If ulcer <> "" And ulcer <> "-" And ulcer <> "n" And ulcer <> "- | -" And ulcer <> "|" And ulcer <> "n | n" And ulcer <> "keine | keine" Then
  Call KRAdd(frm, "Unters", "Diabetisches Ulcus", "L89.18", "L89", , , , "vTextB", 110) ' "Ulcera"
  If labPos <> 0 Then dfsflNeu = True
 Else
  Dim Geschwür$
  Geschwür = IIf(IsNull(frm.adoRS![Geschwür]), "", LCase(frm.adoRS![Geschwür]))
  If Geschwür <> "" And Geschwür <> "-" And Geschwür <> "n" Then
   Call KRAdd(frm, "Anamn", "Diabetisches Ulcus", "L89.18", "L89", "", "", "", "vTextB", 86) ' "Geschwür"
   If labPos <> 0 Then dfsflNeu = True
  End If
 End If

 
 syscmd 4, "Formularvorbereitung 5 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 Dim PStatPath% ' Pulsstatus pathologisch
 PStatPath = 0
 PStatPath = PStatus(Pat_id, frm)
 
 If PStatPath Then
  Call KRAdd(frm, "Pulse", "Periphere art. Verschlußkrankheit", "I70.20", "I73", "I74", "I70", , "vTextB", 123) ' Puls Atp
  If labPos <> 0 Then angflNeu = -1
 End If
 
 Dim HSchw$, Lunge$, herzin%, HIGrund$
 HSchw = IIf(IsNull(frm.adoRS![Herzschwäche]), "", LCase(frm.adoRS![Herzschwäche]))
 If InStrB(HSchw, "j") <> 0 Or InStrB(HSchw, "dyspn") <> 0 Then
  herzin = True
  HIGrund = "Frgb"
 End If
 Lunge = IIf(IsNull(frm.adoRS![Lunge]), "", LCase(frm.adoRS![Lunge]))
 If InStrB(Lunge, "feucht") <> 0 Then
  herzin = True
  HIGrund = IIf(HIGrund = "", "", HIGrund + ", ") + "Ausk"
 End If
 If herzin Then
  Call KRAdd(frm, HIGrund, "V.a. Herzinsuffizienz", "I50.12V", "I50", "", "", "", "vTextB", IIf(HIGrund = "Frgb", 80, 127)) ' Herzschwäche / Lunge
 End If
 
 Dim WS$, obWS%, obBewEin%
 WS = IIf(IsNull(frm.adoRS!WS), "", LCase(frm.adoRS!WS))
 If (InStrB(WS, "ks") <> 0 And InStrB(WS, "kein") = 0) Then
  Call doGilbE(frm, "M53.99", "vTextB", 129) ' WS
  obWS = -1
 End If
 If obWS <> 0 Or obBewEin Then
  Call KRAdd(frm, IIf(obWS, "Unters", "Anamn"), "Wirbelsäulensyndrom", "M53.99", "M53", "M54", "M47")
 End If
 
 Dim Strum$
 If Not IsNull(frm.adoRS!SD) Then
  Strum = LCase$(frm.adoRS!SD)
 End If
 If InStrB(Strum, "stru") <> 0 Then
  If InStrB(Strum, "v.a") <> 0 Then
   Call KRAdd(frm, "Unters", "V.a. Struma", "E04.9V", "E04", , , , "vTextB", 131) ' SD
  Else
   Call KRAdd(frm, "Unters", "Struma", "E04.9", "E04", , , , "vTextB", 131) ' SD
  End If
 End If
 
 Dim weit$, lcWA$, lcAB$, lcSU$
 If IsNull(frm.adoRS![Weitere Anamnese]) Then
  lcWA = ""
 Else
  lcWA = LCase(frm.adoRS![Weitere Anamnese])
 End If
 If IsNull(frm.adoRS![Augensp Befund]) Then
  lcAB = ""
 Else
  lcAB = LCase(frm.adoRS![Augensp Befund])
 End If
 If IsNull(frm.adoRS![Sehminderung unbehebbar]) Then
  lcSU = ""
 Else
  lcSU = LCase(frm.adoRS![Sehminderung unbehebbar])
 End If
 If InStrB(lcWA, "psoriasis") <> 0 And InStrB(LCase(frm.vTextB(147)), "psoriasis") = 0 Then
  Call KRAdd(frm, "Anamn", "Psoriasis", "L40.9", "L40", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 
 If InStrB(lcWA, "osteopor") <> 0 And InStrB(LCase(frm.vTextB(147)), "m81") = 0 Then
  Call KRAdd(frm, "Anamn", "Osteoporose", "M81.99", "M81", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "spinalst") <> 0 And InStrB(LCase(frm.vTextB(147)), "m48") = 0 Then
  Call KRAdd(frm, "Anamn", "Spinalstenose", "M48.09", "M48", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "neurodermitis") <> 0 Or InStrB(LCase(frm.adoRS![Folgeerkrankungen Haut]), "neurodermitis") <> 0) And InStrB(LCase(frm.vTextB(147)), "l20.8") = 0 Then
  Call KRAdd(frm, "Anamn", "Neurodermitis", "L20.8", "L20.8", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "rhythmusstör") <> 0 Or InStrB(lcWA, "sm-imp") <> 0 Or InStrB(LCase(frm.adoRS![Herzkrankheit Beschreibung]), "rhythmusstö") <> 0) And InStrB(LCase(frm.vTextB(147)), "i49") = 0 And InStrB(LCase(frm.vTextB(147)), "i48") = 0 Then
  Call KRAdd(frm, "Anamn", "Herzrhythmusstörungen", "I49.9", "I49", "I48", , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "vorhoffli") <> 0 Or InStrB(LCase(frm.adoRS![Herzkrankheit Beschreibung]), "vorhoffli") <> 0) And InStrB(LCase(frm.vTextB(147)), "i48") = 0 Then
  Call KRAdd(frm, "Anamn", "Vorhofflimmern", "I48.19", "I48", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(LCase(frm.adoRS![Entleerungsstörungen Harnblase]), "j") <> 0 And InStrB(frm.vTextB(147), "N31") = 0 And InStrB(frm.vTextB(147), "N39") = 0 Then
  Call KRAdd(frm, "Anamn", "Störung der Harnblase", "N31.9", "N31", "N39", , , IIf(InStrB(LCase(frm.adoRS![Entleerungsstörungen Harnblase]), "j"), "Entleerungsstörungen Harnblase", ""))
 End If
 If (InStrB(lcWA, "cold") <> 0 Or InStrB(lcWA, "cole") <> 0 Or InStrB(lcWA, "copd") <> 0) And InStrB(LCase(frm.vTextB(147)), "j42") = 0 Then
  Call KRAdd(frm, "Anamn", "Chronische obstruktive Atemwegserkrankung", "J42", "J42", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "asthma") <> 0 And InStrB(LCase(frm.vTextB(147)), "j45") = 0 Then
  Call KRAdd(frm, "Anamn", "Asthma bronchiale", "J45.9", "J45", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(LCase(frm.adoRS![Herz]), "ystolikum") <> 0 And InStrB(LCase(frm.vTextB(147)), "R01") = 0 And InStrB(LCase(frm.vTextB(147)), "I34") = 0 And InStrB(LCase(frm.vTextB(147)), "I35") = 0 Then
  Call KRAdd(frm, "Befd", "Systolikum", "R01.1", "R01.1", "I34", "I35", , "Herz")
 End If
 If InStrB(lcWA, "allergie") <> 0 And InStrB(LCase(frm.vTextB(147)), "t78") = 0 Then
  Call KRAdd(frm, "Anamn", "Allergie", "T78.4", "T78", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "cts") <> 0 Or InStrB(lcWA, "karpaltunnel") <> 0) And InStrB(LCase(frm.vTextB(147)), "g56.0") = 0 Then
  Call KRAdd(frm, "Anamn", "Karpaltunnelsyndrom", "G56.0", "G56.0", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "kreuzschmerzen") <> 0 And InStrB(LCase(frm.vTextB(147)), "m54") = 0 Then
  Call KRAdd(frm, "Anamn", "LWS-Syndrom", "M54.16", "M54", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "byp") <> 0 And InStrB(LCase(frm.vTextB(147)), "z95") = 0 Then
  Call KRAdd(frm, "Anamn", "bypaß peripher", "Z95.88", "Z95.88", , , , "vTextB", 102) ' "Weitere Anamnese")
  Call KRAdd(frm, "Anamn", "Bypaß kardial", "Z95.1", "Z95.1", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "acvb") <> 0 And Not frm.adoRS![Bypaß kardial] And InStrB(LCase(frm.vTextB(147)), "z95.1") = 0 Then
  Call KRAdd(frm, "Anamn", "Bypaß kardial", "Z95.1", "Z95.1", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "glaukom") <> 0 Or InStrB(LCase(frm.adoRS![Sehminderung unbehebbar]), "glaukom") <> 0 Or InStrB(lcAB, "glaukom") <> 0 Or _
    InStrB(lcWA, "grüner star") <> 0 Or InStrB(LCase(frm.adoRS![Sehminderung unbehebbar]), "grüner star") <> 0 Or InStrB(lcAB, "grüner star") <> 0) And InStrB(LCase(frm.vTextB(147)), "glaukom") = 0 Then
  Call KRAdd(frm, "Anamn", "Glaukom", "H40.9", "H40", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 Dim catpos%, catzn%, obcatzn%
 catpos = 0
 catzn = 0
 obcatzn = 0
 Dim AugFeld$, AugNr%
 For i = 1 To 4
  Select Case i
   Case 1: AugFeld = "vTextB": AugNr = 102 ' AugFeld = "Weitere Anamnese"
   Case 2: AugFeld = "vTextB": AugNr = 68 '  "Sehminderung unbehebbar"
   Case 3: AugFeld = "vTextB": AugNr = 66 '"Augensp Befund"
   Case 4: AugFeld = "vTextB": AugNr = 67 '"Netzhaut gelasert"
  End Select
  If catpos = 0 And Not IsNull(frm(AugFeld)(AugNr)) Then
   catpos = InStr(frm(AugFeld)(AugNr).DataField, "atarac")
   If catpos = 0 Then catpos = InStr(frm(AugFeld)(AugNr).DataField, "atarak")
   If catpos = 0 Then catpos = InStr(frm(AugFeld)(AugNr).DataField, "grauer star")
   If catpos > 0 Then
    Call doGilbE(frm, "H26.9", AugFeld, AugNr)
    catzn = InStr(lcAB, "z.n.")
    If catzn > 0 And catzn < catpos Then
     obcatzn = -1
    Else
     catzn = InStr(frm(AugFeld), "op")
     If catzn > 0 And catzn < catpos + 15 Then
      obcatzn = -1
     End If
    End If
   End If
  End If
 Next i
 If catpos > 0 Then
  If catzn > 0 Then
   Call KRAdd(frm, "Anamn", "Z.n. Katarakt", "H26.9", "H26", "H28")
  Else
   Call KRAdd(frm, "Anamn", "Katarakt", "H26.9", "H26", "H28")
  End If
 End If
 
 Dim zähn$
 zähn = IIf(IsNull(frm.adoRS!Zähne), "", LCase(frm.adoRS!Zähne))
 If InStrB(zähn, "ungsbed") <> 0 Then
  Call KRAdd(frm, "Unters", "Kariöse Zähne", "K02.9", "K02", , , , "vTextB", 133) ' Zähne
 End If
 
 Dim beinöd$
 beinöd = IIf(IsNull(frm.adoRS!BeinödVen), "", LCase(frm.adoRS!BeinödVen))
 If (InStrB(beinöd, "ödem") <> 0 And Not beinöd Like "*kein ödem*" And Not beinöd Like "*keine ödem*") Or InStrB(beinöd, "gering") <> 0 Or InStrB(beinöd, "deutl") <> 0 Or InStrB(beinöd, "interm") <> 0 Or InStrB(beinöd, "stauungsderm") <> 0 Then
  Call KRAdd(frm, "Unters", "Beinödeme", "R60.0", "R60", , , , "vTextB", 166) ' BeinödVen
 End If
 
 Dim Beinbefund$
 If Not IsNull(frm.adoRS!Beinbefund) Then
  Beinbefund = LCase$(frm.adoRS!Beinbefund)
 End If
 If Not IsNull(frm.adoRS!Hyperkeratosen) Then
  Beinbefund = Beinbefund & " " & LCase$(frm.adoRS!Hyperkeratosen)
 End If
 If Not IsNull(frm.adoRS!Ulcera) Then
  Beinbefund = Beinbefund & " " & LCase$(frm.adoRS!Ulcera)
 End If
 
 If InStrB(Beinbefund, "hallux valgus") <> 0 Then
  Call KRAdd(frm, "Beinbefund", "Hallux valgus", "M20.1", "M20.1", , , , "vTextB", 141)
 End If

 If InStrB(Beinbefund, "myk") <> 0 Or InStrB(Beinbefund, "tinea") <> 0 Then
  If InStrB(Beinbefund, "nagel") <> 0 Then
   Call KRAdd(frm, "Unters", "V.a. Onychomykose", "B37.2V", "B37", , , , "vTextB", 141) ' Beinbefund
  Else
   Call KRAdd(frm, "Unters", "V.a. Fußmykose", "B35.3V", "B35", "B39", "B37", , "vTextB", 141) ' Beinbefund
  End If
 End If
 If InStrB(Beinbefund, "necrob") <> 0 Or InStrB(Beinbefund, "nekrob") <> 0 Then
  Call KRAdd(frm, "Unters", "Necrobiosis lipoidica", "L92.1", "L92.1", "L99.8", , , "vTextB", 141) ' Beinbefund
 End If
 If (InStrB(Beinbefund, "psoriasis") <> 0 Or InStrB(LCase(frm.adoRS![Folgeerkrankungen Haut]), "psoriasis") <> 0) And InStrB(LCase(frm.vTextB(147)), "psoriasis") = 0 Then
  Call KRAdd(frm, "Unters", "Psoriasis", "L40.9", "L40", , , , "vTextB", 141) ' Beinbefund
 End If
 If InStrB(Beinbefund, "senk-spreizfu") <> 0 Or InStrB(Beinbefund, "platt") <> 0 Or InStrB(Beinbefund, "planus") <> 0 Then
  Call KRAdd(frm, "Unters", "Senk-Spreizfuß", "M21.4", "M21.4", "Q66.8", , , "vTextB", 141) ' Beinbefund
 End If
 If InStrB(Beinbefund, "senkfu") <> 0 Then
  Call KRAdd(frm, "Unters", "Senkfuß", "M21.4", "M21.4", "", "", "", "vTextB", 141) ' Beinbefund
 End If
 If InStrB(Beinbefund, "krallenzeh") <> 0 Then
  Call KRAdd(frm, "Unters", "Krallenzehen", "Q66.8", "Q66.8", , , , "vTextB", 141) ' Beinbefund
 End If
 If InStrB(Beinbefund, "hohlfu") <> 0 Then
  Call KRAdd(frm, "Unters", "Hohlfuß", "Q66.7", "Q66.7", , , , "vTextB", 141) ' Beinbefund
 End If
 
 
 syscmd 4, "Formularvorbereitung 6 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
' If 1 = 0 Then
'  frm.MDi = frm.MDi & vbCrLf
'  MDIn = MDIn + 1
' End If
  
 Alter = 0
 Pat_id = Pat_id
 obWeib = IIf(frm.adoRS!Anrede = "Herr", False, True)
 'Set rlau = TabÖff("Labor", "WertSuch")
 
 Call do_Form_Current2(frm, Hsre, lcAB, PStatPath, pnpflNeu, KZahlNeu, anpflNeu, angflNeu, bgwflNeu, retflNeu, dfsflNeu, oblhNeu)
End If
Dim Ther1$
syscmd 4, "Formularvorbereitung 15 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
Ther1 = TherUmw(TherArt(Pat_id, -1))
If Ther1 <> frm.adoRS!Ther1 Then frm.adoRS!Ther1 = Ther1
'Debug.Print "Form_Current 5:" + format$(Now, "hh:mm:ss")
#If mitab Then
Call ddsono(frm)
#End If
frm.vTextB(182) = DBCn.ConnectionString
syscmd 5
 Call Tüt(440, 20)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Current_AnBog/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Current_AnBog
Function do_Form_Current2(frm As AnBog, Hsre$, lcAB$, PStatPath%, pnpflNeu%, KZahlNeu%, anpflNeu%, angflNeu%, bgwflNeu%, retflNeu%, dfsflNeu%, oblhNeu%)
 
 Dim Aldo_p!, Krea!, Krea02!, Hst!, Adrenu!, AlbCre!, AlbKre!, AlbQ!, Album!, Arq!, Amyl!, Lipase!, ANA!, Andro!, ap!, APCI_CP!, ATIII_K!, B12!, _
     BiliG!, Borg!, Ca!, Ca199_S!, cAnca!, Carigg!, Carigm!, Cea_k!, CHE!, Chltra!, Chltrg!, Chol!, Cicc1d!, _
     Cicc1q!, CKmb!, CKNAC!, CMVIGM!, COERUL!, CORS!, CPEP!, CREACL!, CRP!, CUS!, DIGIT!, DIGOX!, EBVM!, FERR!, FT3_K!, FT4_K!, GAGT!, GGT!, GPT!, GLU!, HAPTO!, HAVm_K!, Hb!, Hk!, HbA1c!, HBC!, HBSAG_K!, HCV!, HCYP!, HGEM_s!, HS!, MG_K!, IA2AK!, IAK_s!, Na!, IGE_K!, IGF1_S!, INSELA!, INSELQ!, k!, LACT_E!, LDL!, LDLH!, Leuko!, LPa!, MAK!, MCH!, NH4!, P_ANCQ_S!, PANC_s!, PANELF!, PCAK_cp!, PHOS!, PO4!, PROL_K!, PSA!, PSAK_cp!, rF!, RFQ!, ST24!, THRM!, THROMB!, TRAK_k!, TSH!, TSHL_k!, VITD25_s!, WACHSH_s!, CPEP_S!
 Dim ccp%
' Diagnosen
 Dim nieflNeu%, nasflNeu%
 On Error GoTo fehler
 nieflNeu = 0
 nasflNeu = 0
 Dim katfl%
 katfl = 0

 syscmd 4, "Formularvorbereitung 6a " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 sql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar, Normbereich as NB " + _
 "FROM laborxus LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " + _
 "WHERE pat_id = " + CStr(Pat_id) + " and not exists (select * from laborneu where pat_id = " + CStr(Pat_id) + " and abkü = laborxwert.Abkü and wert = laborxwert.wert and zeitpunkt > laborxus.Eingang-3 and zeitpunkt <  laborxus.Eingang+6) "
 sql1 = "ORDER BY zeitpunkt desc"
 
' Set raLau = Dtb.OpenRecordset(sql)
 Set raLau = Nothing
 raLau.Open LCase(sql) & sql1, DBCn, adOpenDynamic, adLockReadOnly
'Set rLaU = Dtb.OpenRecordset("select * from [" + QMdbAkt + "].LaborUnion where pat_id = " + CStr(pat_id) + " order by zeitpunkt desc")
 syscmd 4, "Formularvorbereitung 6a1 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 Krea = LabWert(frm, "Krea", ">", 1.3, "Nephropathie", "N08.3", "mg/dl", "N08.3", "N08.3", 1.8, 1.1)
 If labPos <> 0 Then nieflNeu = -1
 Krea02 = LabWert(frm, "Krea02", ">", 1.3, "Nephropathie", "N08.3", "mg/dl", "N08.3", "N08.3", 1.8, 1.1)
 If labPos <> 0 Then nieflNeu = -1
 Krea = LabWert(frm, "Krea", ">", 1.3, "Niereninsuffizienz", "N19", "mg/dl", "N18", "N19", 1.8, 1.1, "N17")
 Krea02 = LabWert(frm, "Krea02", ">", 1.3, "Niereninsuffizienz", "N19", "mg/dl", "N18", "N19", 1.8, 1.1, "N17")
 Hst = LabWert(frm, "hst")
 Adrenu = LabWert(frm, "adrenu") '20 µg/24h
 AlbCre = LabWert(frm, "albcre", ">", 20, "Nephropathie", "N08.3", "mg/gCrea", "N08.3", "N08.3", 50)
 AlbKre = LabWert(frm, "albkre", ">", 20, "Nephropathie", "N08.3", "mg/g Krea", "N08.3", "N08.3", 50)
 AlbQ = LabWert(frm, "albq", ">", 20, "Nephropathie", "N08.3", "mg/g Krea", "N08.3", "N08.3", 50)
 If labPos <> 0 Then nieflNeu = -1
 syscmd 4, "Formularvorbereitung 6b " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
' obkNeph = obKeineNephropathie(pat_id, obMakroAlb)
 Dim UZahl%
 UZahl = Urineintraege(Pat_id)
 If UZahl > 1 Then 'And Not obkNeph Then
   Call KRAdd(frm, CStr(UZahl) + " Stix-eintraege", "V.a. Nephropathie", "N08.3V", "N08.3")
   If labPos <> 0 Then nieflNeu = -1
 End If
 Album = LabWert0(frm, "Album", "<", 35, "Hypalbuminämie", "E46", "g/l", "E46", "N04")
 Aldo_p = LabWert(frm, "aldo_p") ' 15
 Arq = LabWert(frm, "arq") '3,0
 Amyl = LabWert(frm, "amyl", ">", 100, "V.a. Pankreatitis", "K85.0V", "U/l", "K85")
 Lipase = LabWert(frm, "lipase", ">", 60, "V.a. Pankreatitis", "K85.0V", "U/l", "K85")
 ANA = LabWert(frm, "ana", ">", 80, "V.a. Lupus erythematodes o.ä.", "L93.0V", "", "L93", "H01")
 Andro = LabWert(frm, "andro") 'ca. 0,6-3,2/1,1-2,8 s. Internet
 ap = LabWert(frm, "ap") '104/129
 APCI_CP = LabWert0(frm, "APCI_CP", "<", 2.1, "APC-Resistenz", "D68.8", "", "D68", "I82.9", "I82.9")
 ATIII_K = LabWert0(frm, "AT III_K", "<", 75, "Antithrombinmangel", "D68.8", "%", "D68", "I82.9", 40)
 B12 = LabWert0(frm, "b12", "<", 250, "Vit.B12-Mangel", "G63.4", "pg/ml", "D51", "G63.4", 211, 250, "E53.8") ' "G63.4", "D51", "E53.8"
 BiliG = LabWert(frm, "bilig", ">", 1.1, "V.a. Cholestase", "K83.1V", "mg/dl", "K83", "K74", 1.1, 1.1, "K71")
 If labPos <> 0 Then nasflNeu = -1
 Borg = LabWert(frm, "borg", ">", 6, "V.a. Borreliose", "A69.2V", "U/ml", "A68", "A69.2")
 Ca = LabWert0(frm, "ca", "<", 2.1, "Hypocalcämie", "E83.5", "mmol/l", "E83.5", "E83.5", 2)
 Ca = LabWert(frm, "ca", ">", 2.6, "Hypercalcämie", "E83.5", "mmol/l", "E83.5", "E83.5", 2.7)
 Ca199_S = LabWert(frm, "CA199_S", ">", 38, "Pankreas-Carcinom", "C25.-", "U/ml", "C25")
 cAnca = LabWert(frm, "C-ANCA", ">", 2, "V.a. Wegener/mPAN", "M31.3V", "", "M31.3", "M30") ' 2
 Carigg = LabWert(frm, "CARIGM", ">", 12, "Cardiolipin-Ak (IgG)", "D68.8", "U/ml", "A5", "I82.9", 12, 12, "D68.8")
 Carigm = LabWert(frm, "CARIGM", ">", 12, "Cardiolipin-Ak (IgG)", "D68.8", "U/ml", "A5", "I82.9", 12, 12, "D68.8")
 Cea_k = LabWert(frm, "CEA_K", ">", 5, "Carcinom", "C80", "ng/ml", "C80", "C80", 10)
 CHE = LabWert0(frm, "CHE", "<", 5.3, "Lebersynthesestörung", "K74.6", "kU/l", "K74", "K71", 5.3, 5.3, "K70")
 If labPos <> 0 Then nasflNeu = -1
 Chltra = LabWert(frm, "CHLTRA", ">", 1, "Trachom", "A71")
 Chltrg = LabWert(frm, "CHLTRG", ">", 1, "Trachom", "A71")
 Chol = LabWert(frm, "chol", ">", 200, "Hypercholesterinämie", "E78.0", "mg/dl", "E78", "E78")
 Cicc1d = LabWert(frm, "CIC-C1D", ">", 8, "Immunkomplexkranheit", "N05.8", "µgHAG/ml")
 Cicc1q = LabWert(frm, "CIC-C1Q", ">", 10, "Immunkomplexkranheit", "N05.8", "µgHAG/ml")
 CKmb = LabWert(frm, "CKMB", ">", 11.4, "Myokardinfarkt", "I25.2-", "U/l", "I25", "I21", 10)
 CKNAC = LabWert(frm, "CKNAC", ">", 190, "Muskelerkrankung", "G72.9", "U/l", "G72", 167)
 CMVIGM = LabWert(frm, "CMVIGM", ">", 0.9, "Cytomegalievirusinfektion", "", "B25")
 COERUL = LabWert0(frm, "COERUL", "<", 0.2, "M.Wilson", "E83.0", "g/l", "E83")
' rLa.Seek "=", Pat_id, "COR-S"
' raLau.FindFirst "abkü = ""COR-S"""
 Set raLau = Nothing
 raLau.Open sql & " and abkü = ""COR_S"" " & sql1, DBCn, adOpenDynamic, adLockReadOnly
 syscmd 4, "Formularvorbereitung 7 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
' Debug.Print "Form_Current 4:" + format$(Now, "hh:mm:ss")
 ' zwei Tage hintereinander, umgekehrte Zeitreihenfolge
 syscmd 4, "Formularvorbereitung 6c " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 If Not raLau.EOF Then
  If raLau!Wert < 4.3 Then
   raLau.Move 1
   If raLau!Pat_id = Pat_id And raLau!Abkü = "COR-S" Then
    If raLau!Wert < 4.3 Then
     CORS = LabWert0(frm, "COR-S", "<", 4.3, "M.Addison", "E27.1", "µg/dl", "E27")
    End If
   End If
 End If
' ralau.Seek "=", Pat_id, "COR-S"
' raLau.FindFirst "abkü = ""COR-S"""
 Set raLau = Nothing
 raLau.Open sql & " and abkü = ""COR_S"" " & sql1, DBCn, adOpenDynamic, adLockReadOnly

 Dim Zp As Date, VorWert!
 VorWert = 0
 If Not raLau.EOF Then
  Do While Not raLau.EOF
   If raLau!Pat_id <> Pat_id Or raLau!Abkü <> "COR-S" Then Exit Do
    If raLau!Wert > 5 Then
     If VorWert > 20 And Zp = raLau!Zeitpunkt - 1 Then
      CORS = LabWert(frm, "COR-S", ">", 22.4, "M.Cushing", "E24.-", "µg/dl", "E24")
     Else
      Zp = DateValue(raLau!Zeitpunkt)
      Dim VWStr$
      VWStr = Replace(raLau!Wert, "<", "")
      If IsNumeric(VWStr) Then VorWert = VWStr Else VorWert = 0
     End If
    End If
   raLau.Move 1
  Loop
 End If
 'CORT2 = labwert(frm,"CORT2", ">", 3, "M.Cushing", "µg/dl")
 'CORT2 = labwert(frm,"CORT2", "<", 20, "M.Addison", "µg/dl")
 CPEP_S = LabWert(frm, "C PEP_S", ">", 0.1, "Restsekretion", "", "ng/ml")
 CREACL = LabWert(frm, "CREACL")
 CRP = LabWert(frm, "CRP", ">", 5, "Entzündung", "", "mg/l", "B99")
 CUS = LabWert0(frm, "CU-S", "<", 70, "M.Wilson", "E83.0", "µg/dl", "E83.0", "E83.0", 85)
 DIGIT = LabWert(frm, "DIGIT", ">", 30, "Digitoxinintoxikation", "T46.0", "ng/ml")
 DIGOX = LabWert(frm, "DIGOX", ">", 2, "Digoxinintoxikation", "T46.0", "ng/ml")
 EBVM = LabWert(frm, "EBVM", ">", 1, "EBV-Virusinfektion", "B27.0", "")
 FERR = LabWert(frm, "FERR", ">", 600, "V.a. Hämochromatose", "E83.1V", "µg/l", "E83.1")
 FT3_K = LabWert(frm, "FT3_K", ">", 4.2, "Hyperthyreose", "E05.9", "pg/ml", "E05")
 FT4_K = LabWert(frm, "FT4_K", ">", 1.7, "Hyperthyreose", "E05.9", "ng/dl", "E05")
 FT4_K = LabWert(frm, "FT4_K", "<", 0.8, "Hypothyreose", "E03.9", "ng/dl", "E03", "E89", 0.8, 0.8, "E02")
 GAGT = LabWert(frm, "GAGT", ">", 66, "Hepatopathie", "K71.6", "U/l", "K71", "K76", 66, 39, "K77")
 If labPos <> 0 Then nasflNeu = -1
 GGT = LabWert(frm, "GGT", ">", 66, "Hepatopathie", "K71.7", "U/l", "K71", "K76", 66, 39, "K77")
 If labPos <> 0 Then nasflNeu = -1
 GPT = LabWert(frm, "GPT", ">", 50, "Hepatopathie", "K71.7", "U/l", "K71", "K76", 50, 35, "K77")
 If labPos <> 0 Then nasflNeu = -1
 GLU = LabWert(frm, "GLU", ">", 200, "Diabetes mellitus", "E11.9", "mg/dl", "E10", "E11", 130, 200, "E14")
 HAPTO = LabWert(frm, "HAPTO", "<", 0.3, "Hämolyse", "D59.9", "g/l", "D5")
 HAVm_K = LabWert(frm, "HAV M_K", ">", 0.1, "Hepatitis A", "B15.-", "", "B15")
 Hb = LabWert0(frm, "HB", "<", 14, "Anämie", "K92.2", "g/dl", "D64", "D5", 12.6, 12, "K9")
 Hb = LabWert(frm, "HB", ">", 17.4, "Polyglobulie", "J44.9", "g/dl", "D45", "J44.9", 17.4, 16)
 Hk = LabWert(frm, "HK", ">", 0.52, "Polycythaemia vera", "J44.9", "", "D45", 0.52, 0.46)
 HbA1c = LabWert(frm, "HBA1C", ">=", 6.5, "schlechte Einstellung", "", "%", ".71", ".61")
 HbA1c = LabWert0(frm, "HBA1C", "<", 6.6, "gute Einstellung", "", "%", ".90", ".40")
 HBC = LabWert(frm, "HBC", ">", 0.1, "Hepatitis B", "B16.-", "", "B16")
 HBSAG_K = LabWert(frm, "HBSAG_K", ">", 0.1, "Hepatitis B", "B16.-", "", "B16")
 HCV = LabWert(frm, "HCV", ">", 0, "(Z.n.)Hep.C", "B18.2", "", "B17", "B18")
 HCYP = LabWert(frm, "HCY-P", ">", 10.3, "Hyperhomozysteinämie", "E72.1", "µmol/l", "E72.1")
 HGEM_s = LabWert0(frm, "HGEM_S", "<", 20, "Ehrlichiose", "B99", "", "B99")
 HS = LabWert(frm, "HS", ">", 8.2, "Hyperurikämie", "E79.0", "mg/dl", "E79.0", "E79.0", 8.2, 6.1)
 Hsre = LabWert(frm, "HSRE", ">", 8.2, "Hyperurikämie", "E79.0", "mg/dl", "E79.0", "E79.0", 8.2, 6.1)
 MG_K = LabWert0(frm, "MG_K", "<", 0.66, "Magnesiummangel", "E61.2", "mmol/l", "E61.2")
 IA2AK = LabWert(frm, "IA2AK", ">", 0.1, "D.m.1 (IA2)", "E10.91", "U/ml", "E10")
 IA2AK = LabWert(frm, "IA2AK", "<", 0.1, "D.m.2 (IA2)", "E11.91", "U/ml", "E11", "R73")
 IAK_s = LabWert(frm, "IAK_S", ">", 0.1, "D.m.1 (IA)", "E10.91", "U/ml", "E10")
 IAK_s = LabWert(frm, "IAK_S", "<", 0.1, "D.m.2 (IA)", "E11.91", "U/ml", "E11", "R73")
 Na = LabWert0(frm, "NA", "<", 133, "Hyponatriämie", "E87.1", "mmol/l", "E87.1")
 Na = LabWert(frm, "NA", ">", 146, "Hypernatriämie", "E87.1", "mmol/l", "E87.0")
 IGE_K = LabWert(frm, "IGE_K", ">", 25, "Atopie", "T78.4", "U/ml", "T78", "T78", 25, 100)
 IGF1_S = LabWert(frm, "IGF-1_S", ">", 359, "V.a. Akromegalie", "E22.0V", "ng/ml", "E22", "E22", 316)
 INSELA = LabWert(frm, "INSELA", ">", 10, "D.m.1 (ICAa)", "E10.91", "", "E10")
 INSELA = LabWert(frm, "INSELA", "<", 10, "D.m.2 (ICAa)", "E11.91", "", "E11", "R73")
 INSELQ = LabWert(frm, "INSELQ", ">", 10, "D.m.1 (ICAq)", "E10.91", "", "E10")
 INSELQ = LabWert(frm, "INSELQ", "<", 10, "D.m.2 (ICAq)", "E11.91", "", "E11", "R73")
 k = LabWert(frm, "K", ">", 5.5, "Hyperkaliämie", "E87.5", "mmol/l", "E87.5", "E87.5", 5.5)
 k = LabWert0(frm, "K", "<", 3.6, "Hypokaliämie", "E87.6", "mmol/l", "E87.6", "E87.6", 3.6)
 LACT_E = LabWert(frm, "LACT_E", ">", 20, "Lactat-Azidose", "E87.2", "mg/dl", "E87.2")
 If Arq > 3 And Aldo_p > 15 Then _
  Call LabWert(frm, "ALDO_P", ">", 15, "V.a. Conn-Syndrom", "E26.0V", "", "E26.0")
 If Adrenu > 20 Or (Alter < 10 And Adrenu > 10) Then _
 Call LabWert(frm, "ADRENU", ">", 10, "V.a. Phäochromozytom", "D35.0V", "", "D35.0")
 LDL = LabWert(frm, "LDL", ">", 130, "Hypercholesterinämie", "E78.0", "mg/dl", "E78", "E78", 160)
 LDLH = LabWert(frm, "LDLH", ">", 130, "Hypercholesterinämie", "E78.0", "mg/dl", "E78", "E78", 160)
 Leuko = LabWert(frm, "LEUKO", ">", 10.5, "Entzündung", "B99", "/nl", "B99", "T8")
 LPa = LabWert(frm, "LP A", ">", 30, "Fettstoffwechselstörung", "E78.9", "mg/dl", "E78")
' ab hier muß weitergearbeitet werden
 MAK = LabWert(frm, "MAK", ">", 40, "Autoimmune SD-Krht", "E05.-", "IU/ml", "E05", "E06", 100)
 If (obWeib And Hb < 12.6) Or (Not obWeib And Hb < 14) Then
  MCH = LabWert(frm, "MCH", "<", 28, "Mikrozyt. Anämie", "D50.0", "pg", "D5", "D6", "K9")
  MCH = LabWert(frm, "MCH", ">", 34, "Makrozyt. Anämie", "D52.9", "pg", "D5", "D6", "K9")
 End If
 NH4 = LabWert(frm, "NH4", ">", 94, "Hepatische Enzephalopathie", "K72.7-", "µmol/l", "K72.7", "K74.6", 94, 82)
 P_ANCQ_S = LabWert(frm, "P_ANCQ_S", ">", 2, "V.a. mPAN/GN/RA/LE", "M31.-V", "", "M31", "N05")
 PANC_s = LabWert(frm, "P-ANC_S", ">", 2, "V.a. mPAN/GN/RA/LE", "M31.-V", "", "M31", "N05")
 PANELF = LabWert0(frm, "PANELF", "<", 300, "Pankreasinsuffizienz", "K86.8", "µg/g", "K86.8", "K86.1", 200)
 PCAK_cp = LabWert0(frm, "PC-AK_CP", "<", 70, "V.a. Protein-C-Mangel", "D68.8V", "%", "I82.9")
 PHOS = LabWert(frm, "PHOS", ">", 4.8, "N25.8", "sek. Hyperpara", "N25.8", "mg/dl", "N25.8", "E21")
 PO4 = LabWert(frm, "PO4", ">", 4.8, "N25.8", "sek. Hyperpara", "N25.8", "mg/dl", "N25.8", "E21")
 PROL_K = LabWert(frm, "PROL_K", ">", 619, "Prolaktinom", "D35.2", "mIU/l", "D35")
 PSA = LabWert(frm, "PSA", ">", 4, "Prostata-Ca", "C61", "ng/ml", "N40", "C61", 10)
 PSAK_cp = LabWert0(frm, "PS-AK_CP", "<", 70, "V.a. Protein-S-Mangel", "D68.8V", "%", "I82.9")
 rF = LabWert(frm, "RF", ">", 20, "Rheuma", "M05.-", "U/ml", "M05", "M06")
 RFQ = LabWert(frm, "RFQ", ">", 20, "Rheuma", "M05.-", "U/ml", "M05", "M06")
 ST24 = LabWert(frm, "ST 24", ">", 2, "Schwangerschaft", "", "IU/l", "", "", 10)
 THRM = LabWert0(frm, "THRM", "<", 130, "Thrombopenie", "D69.5-", "/nl", "D69", "K74", 50)
 THRM = LabWert(frm, "THRM", ">", 430, "Thrombozytose", "D75.2", "/nl", "D75", "D75", 1000)
 THROMB = LabWert0(frm, "THROMB", "<", 130, "Thrombopenie", , "D69.5-", "/nl", "D69", "K74", 50)
 THROMB = LabWert(frm, "THROMB", ">", 430, "Thrombozytose", "D75.2", "/nl", "D75.9", "D75.9", 1000)
 TRAK_k = LabWert(frm, "TRAK_K", ">", 1, "M.Basedow", "E05.0", "U/l", "E05", "E05", 2)
 TSH = LabWert0(frm, "TSH", "<", 0.3, "V.a. Hyperthyreose", "E05.9V", "µU/ml", "E05")
 TSH = LabWert(frm, "TSH", ">", 4.5, "Hypothyreose", "E03.9", "µU/ml", "E03", "E03", 10)
 TSHL_k = LabWert0(frm, "TSH-L_K", "<", 0.3, "V.a. Hyperthyreose", "E03.9V", "µU/ml", "E05.9")
 TSHL_k = LabWert(frm, "TSH-L_K", ">", 4.5, "Hypothyreose", "E03.9", "µU/ml", "E03", "E03", 10)
 VITD25_s = LabWert0(frm, "VITD25_S", "<", 10, "Vitamin-D-Mangel", "E55.9", "ng/ml", "E55", "E55")
 WACHSH_s = LabWert(frm, "WACHSH_S", ">", 7, "V.a. Acromegalie", "D35.2V", "ng/ml", "E22")
 syscmd 4, "Formularvorbereitung 8 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 ccp = False
 If CREACL <> 0 Then
  Select Case Alter
   Case Is < 20
   Case Is < 30: CGr = IIf(obWeib, 72, 94)
   Case Is < 40: CGr = IIf(obWeib, 71, 89)
   Case Is < 50: CGr = IIf(obWeib, 50, 76)
   Case Is < 60: CGr = IIf(obWeib, 45, 54)
   Case Is < 70: CGr = IIf(obWeib, 37, 49)
   Case Is < 80: CGr = IIf(obWeib, 27, 30)
   Case Is < 90: CGr = IIf(obWeib, 26, 26)
  End Select
  If CREACL < CGr Then ccp = True
   If ccp Then _
    CREACL = LabWert(frm, "CREACL", "<", CGr, "Niereninsuffizienz", "N19", "ml/min", "N19", "N18", CGr, CGr - 10, "N17")
'  frm.MDi = frm.MDi + "Nierensinffizienz (Crea-Clearance " + CStr(CREACL) + " ml/min)" & vbcrlf
  End If
 End If 'CreaCL <> 0

 ' Blutdruck auswerten
' Dim rsRR As DAO.Recordset
 Dim rsRR As New ADODB.Recordset
 Dim pZahl%
 pZahl = 0
' Set rsRR = TabÖff("RRParse", "ID")
' rsRR.Seek "=", Pat_id
' If Not rsRR.NoMatch Then
 frm.vTextB(150) = vbNullString
 rsRR.Open "select * from rrparse where pat_id = " & Pat_id & " order by zeitpunkt desc", DBCn, adOpenDynamic, adLockReadOnly
 If Not rsRR.BOF Then
  Do While Not rsRR.EOF
   If rsRR!RRsyst > 140 Or rsRR!RRdiast > 90 Then
    pZahl = pZahl + 1
   End If
   frm.vTextB(150) = frm.vTextB(150) & IIf(LenB(frm.vTextB(150)) = 0, "", ",") & rsRR!RRsyst & "/" & rsRR!RRdiast
   rsRR.MoveNext
  Loop
 End If ' not rsRR.bof then
 If pZahl = 1 Or pZahl = 2 Then
  Call KRAdd(frm, CStr(pZahl) + " Meßw.", "V.a. Arterielle Hypertonie", "I10.90V", "I10", , , , "vTextB", 150) ' RR
 ElseIf pZahl > 2 Then
  Call KRAdd(frm, CStr(pZahl) + " Meßw.", "Arterielle Hypertonie", "I10.90", "I10", , , , "vTextB", 150) ' R
 End If ' pzahl
 
 Dim BMI#
 If Not IsNull(frm.adoRS!Gewicht) And Not IsNull(frm.adoRS!Größe) Then
  If frm.adoRS!Größe > 0 Then
   BMI = frm.adoRS!Gewicht * IIf(frm.adoRS!Gewicht < 3, 100, 1) / frm.adoRS!Größe / frm.adoRS!Größe ' * 10000
   Do
    If BMI = 0 Then Exit Do
    If BMI > 8 Then Exit Do
    BMI = BMI * 10
   Loop
   If BMI >= 28 Then
    Call KRAdd(frm, "BMI = " + Format$(BMI, "##,#"), "Übergewicht", "E66.9", "E66", , , , "vTextB", 11) ' Gewicht
   End If
  End If
 End If
 
 If InStrB(frm.vTextB(168), "Hypertonie") = 0 And Left(frm.adoRS!Bluthochdruck, 1) = "j" Then ' MDi
  Call KRAdd(frm, "Ankreuzfeld", "Arterielle Hypertonie", "I10.90", "I10", , , , "vTextB", 58) ' Bluthochdruck
 End If
 
 If Not IsNull(frm.adoRS![Augensp Befund]) Then
  Dim rpfl As Boolean, glaufl As Boolean, t$
  katfl = 0
  rpfl = 0
  glaufl = 0
  t = lcAB
  If InStrB(t, "grau") <> 0 Or InStrB(t, "atara") <> 0 Then katfl = True
  If InStrB(t, "keine") = 0 And InStrB(t, "veränd") <> 0 Then rpfl = True
  If InStrB(t, "op ") <> 0 Or InStrB(t, "operier") <> 0 Then rpfl = True
  If InStrB(t, "path") <> 0 And InStrB(t, "path") <> InStrB(t, "opath") + 1 Then rpfl = True
  If InStrB(t, "gelasert") <> 0 Then rpfl = True
  If InStrB(t, "blut") <> 0 Then rpfl = True
  If InStrB(t, "schlecht") <> 0 Then rpfl = True
  If InStrB(t, "retinop") <> 0 And InStrB(t, "keine") = 0 Then rpfl = True
  If InStrB(t, "o.B") <> 0 Or InStrB(t, "o.b") <> 0 Or InStrB(t, "oB") <> 0 Then rpfl = False
  If rpfl Then Call doGilb(frm, "H36.0", "vTextB", 66)
  If InStrB(LCase(frm.adoRS![Netzhaut gelasert]), "j") <> 0 Then
   If InStrB(frm.adoRS![Netzhaut gelasert], "wegen") = 0 And _
      InStrB(frm.adoRS![Netzhaut gelasert], "wg") = 0 Then
      Call doGilb(frm, "H36.0", "vTextB", 67)
      rpfl = True
   End If
  End If
  If InStrB(t, "glau") <> 0 Or InStrB(t, "grün") <> 0 Then glaufl = True
 End If
 If rpfl Then
  Call KRAdd(frm, "Fragebogen", "Diab. Retinopathie", "H36.0") ' Augensp Befund
  If labPos <> 0 Then retflNeu = -1
 End If
 If katfl Then
  Call KRAdd(frm, "Fragebogen", "Diab. Katarakt", "H28.0", "H28", "H26", , , "vTextB", 66) ' Augensp Befund
 End If
 If glaufl Then
  Call KRAdd(frm, "Fragebogen", "Glaukom", "H40.9", "H40", , , , "vTextB", 66) ' Augensp Befund
 End If
 If Not obkNeph And ((InStrB(LCase(frm.adoRS![Diabet Nierenschaden]), "j") <> 0 Or frm.adoRS!Dialyse <> 0 Or diI("Z49.1", frm.adoRS!Pat_id)) And InStrB(frm.vTextB(168), "Nephrop") = 0) Then ' MDi
  Call KRAdd(frm, "Fragebogen", "Diab.Nephropathie", "N08.3", , , , , "vTextB", 69) ' Diabet Nierenschaden
  If labPos <> 0 Then nieflNeu = -1
 End If
 If InStrB(LCase(frm.adoRS![andere Nierenerkrankung]), "stein") <> 0 Then
  Call KRAdd(frm, "Fragebogen", "Z.n. Nierensteinleiden", "N20.0Z", "N20", , , , "vTextB", 73) ' andere Nierenerkrankung
 End If
 
 Dim Anginap$, obKHK%, mi$, obMI%, ptca$, obPTCA%, hkBschr$
 obKHK = 0
 obMI = 0
 obPTCA = 0
 Anginap = IIf(IsNull(frm.adoRS![Angina pectoris]), "", LCase(frm.adoRS![Angina pectoris]))
 If Left(Anginap, 1) = "j" Then
  Call doGilbE(frm, "I25.9", "vTextB", 75) ' Angina pectoris
  obKHK = 1
 ElseIf InStrB(Anginap, "leicht") <> 0 Or InStrB(Anginap, "z.n.") <> 0 Or InStrB(Anginap, "evtl") <> 0 Or InStrB(Anginap, "manchmal") <> 0 Then
  Call doGilbE(frm, "I25.9", "vTextB", 75) ' Angina pectoris
  obKHK = 1
 End If
 mi = IIf(IsNull(frm.adoRS![Herzinfarkt]), "", LCase(frm.adoRS![Herzinfarkt]))
 If Left(mi, 1) = "j" Or InStrB(mi, "stummer") <> 0 Or InStrB(mi, "vwi") <> 0 Or InStrB(mi, "hwi") <> 0 Then
  Call doGilbE(frm, "I21.9", "vTextB", 76) ' Herzinfarkt
  obMI = 1
  obKHK = 2
 End If
 ptca = IIf(IsNull(frm.adoRS![PTCA oder Stent]), "", LCase(frm.adoRS![PTCA oder Stent]))
 If InStrB(ptca, "j") <> 0 Then
  Call doGilbE(frm, "Z95.5Z", "vTextB", 78) ' PTCA oder Stent
  obPTCA = 1
  obKHK = 2
 End If
 If frm.adoRS![Bypass kardial] <> 0 Then
  Call doGilbE(frm, "Z95.1", "vCheckB", 3) ' Bypaß kardial
  obKHK = 2
 End If
 hkBschr = IIf(IsNull(frm.adoRS![Herzkrankheit Beschreibung]), "", LCase(frm.adoRS![Herzkrankheit Beschreibung]))
 If InStrB(hkBschr, "a.p.") <> 0 Or InStrB(hkBschr, "retrost") <> 0 Or InStrB(hkBschr, "-gef") <> 0 Or InStrB(hkBschr, "brustschmerz") <> 0 Then
  Call doGilbE(frm, "I25.9", "vTextB", 75) ' Angina pectoris
  obKHK = 2
 End If
 If obKHK = 2 Then
   If InStrB(hkBschr, "3-gef") <> 0 Or InStrB(hkBschr, "4-gef") <> 0 Or InStrB(hkBschr, "3-fach") <> 0 Or InStrB(hkBschr, "4-fach") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Koronare Dreigefäßerkrankung", "I25.13", "I25")
   ElseIf InStrB(hkBschr, "2-gef") <> 0 Or InStrB(hkBschr, "2-fach") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Koronare Zweigefäßerkrankung", "I25.12", "I25")
   ElseIf InStrB(hkBschr, "1-gef") <> 0 Or InStrB(hkBschr, "1-fach") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Koronare Eingefäßerkrankung", "I25.11", "I25")
   Else
    Call KRAdd(frm, "Fragebogen", "Koronare Herzerkrankung", "I25.9", "I25")
   End If
 ElseIf obKHK = 1 Then
  Call KRAdd(frm, "Fragebogen", "V.a. koronare Herzerkrankung", "I25.9V", "I25")
 End If
 If obPTCA = 1 Then
  Call KRAdd(frm, "Fragebogen", "Z.n. PTCA oder Stent", "Z95.5", "Z95.5")
 End If
 If frm.adoRS![Bypass kardial] <> 0 Then Call KRAdd(frm, "Fragebogen", "Z.n. Bypaß kardial", "Z95.1")
 If obMI = 1 Then Call KRAdd(frm, "Fragebogen", "Z.n. Myokardinfarkt", "I21.9", "I21", "I25.2")
 
 Dim hdb$
 hdb = IIf(IsNull(frm.adoRS![Hirndurchblutungsstörung]), "", LCase(frm.adoRS![Hirndurchblutungsstörung]))
 If InStrB(hdb, "j") <> 0 Or InStrB(hdb, "z.n.") <> 0 Then
  Call KRAdd(frm, "Fragebogen", "V.a. Atheromatose hirnzuf. Gefäße", "I65.9V", "I65", "I66", "I68", "vTextB", 82) ' Hirndurchblutungsstörung")
  If labPos <> 0 Then angflNeu = -1
 End If
 Dim Schlag$
 Schlag = IIf(IsNull(frm.adoRS![Schlaganfall]), "", LCase(frm.adoRS![Schlaganfall]))
 If Len(Schlag) > 0 And Schlag <> "-" And Schlag <> "n" Then
  If InStrB(Schlag, "gedächtnis") = 0 And InStrB(Schlag, "vergeßlich") = 0 And InStrB(Schlag, "vergesslich") = 0 And InStrB(Schlag, "entfällt") = 0 Then
   If InStrB(Schlag, "blutu") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Z.n. Hirnblutung", "I61.9Z", "I61", "I62", , , "vTextB", 83) ' Schlaganfall
   Else
    Call KRAdd(frm, "Fragebogen", "Z.n. Apoplex", "I63.9Z", "I64", "I63", , , "vTextB", 83) ' Schlaganfall
   End If
  End If
 End If
 
 syscmd 4, "Formularvorbereitung 10 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 Dim BDBSt%, SFKht%
 BDBSt = 0
 SFKht = 0
 If Left(LCase(frm.adoRS![Beindurchblutungsstörung]), 1) = "j" Then
  BDBSt = -1
  Call doGilb(frm, "I73.9", "vTextB", 84) ' Beindurchblutungsstörung
 End If
 If Left(LCase(frm.adoRS![Schaufensterkrankheit]), 1) = "j" Then
  SFKht = -1
  Call doGilb(frm, "I73.9", "vTextB", 85) ' Schaufensterkrankheit
 End If
 If (BDBSt Or SFKht) And Not (BDBSt And Not SFKht And Not PStatPath And frm.adoRS![bypaß peripher] = 0 And IsNull(frm.adoRS![Geschwür]) And IsNull(frm.adoRS![Amputation]) And IsNull(frm.adoRS![pAVK Beschreibung])) Then
  Call KRAdd(frm, "Fragebogen", "Periphere arterielle Verschlußkrankheit", "I73.9", "I73", "I74", "I70")
  If labPos <> 0 Then angflNeu = -1
 End If
 If frm.adoRS![bypaß peripher] <> 0 Then
  Call KRAdd(frm, "Fragebogen", "bypaß peripher", "Z95.88", "I73", "I74", "I70", "Z95", "vCheckB", 4) ' bypaß peripher
  If labPos <> 0 Then angflNeu = -1
 End If
 Dim amput$
 amput = IIf(IsNull(frm.adoRS![Amputation]), "", LCase(frm.adoRS![Amputation]))
 If Len(amput) > 1 And amput <> "-" And amput <> "n" And amput <> "entfällt" Then
  Call KRAdd(frm, "Fragebogen", "Z.n. Amputation", "Z44.1", "Z44", "Z97", , , "vTextB", 87) ' Amputation
 End If
#If mitab Then
 Call DiabetesDiagnose(frm, pnpflNeu, KZahlNeu, anpflNeu, nieflNeu, angflNeu, bgwflNeu, katfl, retflNeu, nasflNeu, dfsflNeu, oblhNeu)
#End If
syscmd 4, "Formularvorbereitung 11 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
Call knöpfeanpassen(frm)
 
syscmd 4, "Formularvorbereitung 12 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
'Call do_Hirndurchblutungsstörung_Exit(0, frm)
 Dim rsVerSi As New ADODB.Recordset, rsv1 As New ADODB.Recordset, rsV0 As New ADODB.Recordset
 Dim rsVK As New ADODB.Recordset
 Dim Versicherung$
 sql = "SELECT * from faelle INNER JOIN (SELECT min(pat_id) as pid, max(bhfb) as bb from faelle GROUP BY pat_id) AS sel ON (`faelle`.`bhfb`=sel.bb) AND (`faelle`.`pat_id`=sel.pid)"
' Set rVK = Dtb.OpenRecordset(sql)
 rsVK.Open sql & " where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
' rVK.FindFirst ("pat_id = " + CStr(Pat_id))
 Versicherung = ""
 If Not rsVK.BOF Then
  Versicherung = rsVK!VKNr
 End If
' Versicherung = IIf(IsNull(rVK!VKNr) Or rVK.NoMatch, "", rVK!VKNr)
''If Not versi.BOF Then frm.Versicherung = versi!Versicherung
''versi.Close
'' Aufrufe der Anamnese usw. belegen
 Call getDokPfad
 Dim rDok As ADODB.Recordset
 
 syscmd 4, "Formularvorbereitung 13 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 An1Pfad = PfadFestLeg("An1Aufruf", "%anamnese%1%", frm)
 An2Pfad = PfadFestLeg("An2Aufruf", "%anamnese%2%", frm)
 AnAPfad = PfadFestLeg("AnAAufruf", "%anamnese%allg%", frm, "%allg%anamnese%")
 CheckPfad = PfadFestLeg("CheckAufruf", "%che%kliste%", frm)
 
 Dim lz%, lzu%, rsLab As New ADODB.Recordset
 syscmd 4, "Formularvorbereitung 13a " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 'lz = Dtb.OpenRecordset("select count(*) as ct from [" + QMdbAkt + "].LaborDokumente where pat_id = " + CStr(frm.Pat_id))!ct
 sql = "SELECT count(0) as ct " & _
 "FROM (dokumente as dl LEFT JOIN `dokumente abgehakt` as da ON dl.DokPfad=da.DokPfad) " & _
 "where DokName Like '%labor%' and dl.pat_id = " & CStr(Pat_id)
 Set rsLab = Nothing
 rsLab.Open sql, DBCn, adOpenStatic, adLockReadOnly
 If Not rsLab.BOF Then lz = rsLab!ct
 syscmd 4, "Formularvorbereitung 13b " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
' lzu = Dtb.OpenRecordset("select count(*) as ct from [" + QMdbAkt + "].LaborDokumente where pat_id = " + CStr(frm.Pat_id) + " and abgehakt")!ct
 sql = "SELECT count(0) as ct " & _
 "FROM (dokumente as dl LEFT JOIN `dokumente abgehakt` as da ON dl.DokPfad=da.DokPfad) " & _
 "where DokName Like '%labor%' and abgehakt and dl.pat_id = " & CStr(Pat_id)
 Set rsLab = Nothing
 rsLab.Open sql, DBCn, adOpenStatic, adLockReadOnly
 If Not rsLab.BOF Then lzu = rsLab!ct
 syscmd 4, "Formularvorbereitung 13c " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 frm.vCommandB(6).Caption = CStr(lzu) + "&/" + CStr(lz) & " " & "LD" ' Labor
 frm.vCommandB(6).Enabled = True ' Labor
 frm.vCommandB(6).MaskColor = 0   ' Labor ' eigentlich ForeColor, gibt es aber nicht
 If lzu = lz Then
  If lz = 0 Then
   frm.vCommandB(6).Enabled = False ' Labor
  Else
   frm.vCommandB(6).MaskColor = RGB(100, 100, 100) ' Labor
  End If
 Else
 End If
 If 1 = 0 And lzu = 0 Then
   frm.vCommandB(6).Enabled = 0 ' Labor
 End If
' jetzt das gleiche für Briefe
 syscmd 4, "Formularvorbereitung 13d " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 sql = "select count(0) as ct from briefe where pat_id = " + CStr(Pat_id)
 Set rsLab = Nothing
 rsLab.Open sql, DBCn, adOpenStatic, adLockReadOnly
 If Not rsLab.BOF Then
  lz = rsLab!ct
  frm.vCommandB(9).Caption = CStr(lz) + " Briefe (7)" '(&7)
 End If
 sql = "select max(zeitpunkt) as zp from briefe where (name like 'Brief an Dr%' or name like 'Brief an F%Dr.' or name like '%Arztbrief%') and not name like '%Entwurf%' and pat_id = " + CStr(Pat_id)
 '!zp, "dd/mm/yy")
 Set rsLab = Nothing
 rsLab.Open sql, DBCn, adOpenStatic, adLockReadOnly
 If Not rsLab.BOF Then
  frm.vCommandB(9).Caption = frm.vCommandB(9).Caption & ", zul." & rsLab!Zp ' Briefe
 End If
' frm.briefe.Caption = CStr(lz) + " Briefe (&7), zul." + format$(Dtb.OpenRecordset("select max(zeitpunkt) as zp from [" + QMdbAkt + "].briefe where (name like ""Brief an Dr*"" or name like ""Brief an F*Dr."" or name like ""*Arztbrief*"") and not name like ""*Entwurf*"" and pat_id = " + CStr(Pat_id))!zp, "dd/mm/yy")
 frm.vCommandB(9).Enabled = True ' Briefe
 frm.vCommandB(9).MaskColor = 0 ' Briefe
 If lz = 0 Then
   frm.vCommandB(9).Enabled = 0 ' Briefe
 End If
 syscmd 4, "Formularvorbereitung 14 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname

' und für Blutzuckerkurven
 Dim maxzp As Date
 maxzp = 0
 sql = "select count(0) as ct, max(zeitpunkt) as mzp from dokumente where dokname like '%BZ%' and pat_id = " & CStr(Pat_id)
 Set rsLab = Nothing
 rsLab.Open sql, DBCn, adOpenStatic, adLockReadOnly
 If Not rsLab.BOF Then
  lz = rsLab!ct
  If lz > 0 Then
   maxzp = rsLab!mzp
  End If
 End If
' BZKurven
 frm.vCommandB(11).Caption = CStr(lz) + " BZKur.(8), zul." & Format$(maxzp, "dd/mm/yy") ' BZKurven ' (&8)
 frm.vCommandB(11).Enabled = True ' BZKurven
 frm.vCommandB(11).MaskColor = 0  ' BZKurven
 If lz = 0 Then
   frm.vCommandB(11).Enabled = 0 ' BZKurven
 End If
' und für Augenbefunde
' lz = Dtb.OpenRecordset("select count(*) as ct from [" + QMdbAkt + "].Dokumente where (dokname like ""*augenb[!l]*"" or dokname like ""*augen[aä]rzt*"" or dokname like ""*aa"") and pat_id = " + CStr(frm.Pat_id))!ct + _
      Dtb.OpenRecordset("select count(*) as ct from [" + QMdbAkt + "].eintraege where (inhalt like ""*augenb[!l]*"" or inhalt like ""*augen[aä]rzt*"" or inhalt like ""* aa[!g]*"") and pat_id = " + CStr(Pat_id))!ct
 lz = 0
 maxzp = 0
 Set rsLab = Nothing
 sql = "select count(0) as ct,max(zeitpunkt) as mzp from dokumente where ((dokname like '%augenb%' and not dokname like '%augenbl%') or dokname like '%augenarzt%' or dokname like '%augenärzt%' or dokname like '%aa') and pat_id = " & CStr(Pat_id) ' dokname regexp '.*augenb[^l].*'
 Set rsLab = Nothing
 rsLab.Open sql, DBCn, adOpenStatic, adLockReadOnly
 If Not rsLab.BOF Then
  lz = rsLab!ct
  If lz > 0 Then
   maxzp = rsLab!mzp
  End If
 End If
 sql = "select count(0) as ct,max(zeitpunkt) as mzp from eintraege where ((inhalt like '%augenb%' and not inhalt like '%augenbl%') or inhalt like '%augenarzt%' or inhalt like '%augenärzt%' or (inhalt like '% aa' and not inhalt like '% aag%')) and pat_id = " + CStr(Pat_id)
 Set rsLab = Nothing
 rsLab.Open sql, DBCn, adOpenStatic, adLockReadOnly
 If Not rsLab.BOF Then
  lz = lz + rsLab!ct
  If rsLab!mzp > maxzp Then maxzp = rsLab!mzp
 End If
 
 frm.vCommandB(13).Caption = CStr(lz) + " Aug'b.(9), zul." & Format$(maxzp, "dd/mm/yy") ' AugenBefunde ' (&9)
 frm.vCommandB(13).Enabled = True ' AugenBefunde
 frm.vCommandB(13).MaskColor = 0 ' AugenBefunde
 If lz = 0 Then
   frm.vCommandB(13).Enabled = 0 ' AugenBefunde
 End If
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Current2/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
#End If ' mitab
#If mitab Then
 Function DiabetesDiagnose(frm As AnBog, pnpflNeu%, KZahlNeu%, anpflNeu%, nieflNeu%, angflNeu%, bgwflNeu%, katfl%, retflNeu%, nasflNeu%, dfsflNeu%, oblhNeu%)
'Diabetesdiagnose
 Dim DICD$, DICDg$, DICDh$
 Dim KZahl%, anpfl As Boolean, pnpfl As Boolean, retfl As Boolean, angfl As Boolean, dfsfl As Boolean, _
 nasfl As Boolean, bgwfl As Boolean, niefl As Boolean, ketfl As Boolean, hypfl As Boolean
 Dim DiabetesName$
 DiabetesName = "Diabetes mellitus"
 Select Case frm.adoRS!Diabetestyp
  Case "1"
   DICD = "E10"
  Case "2"
   DICD = "E11"
  Case "-"
   DICD = "-"
  Case "?"
   DICD = "E14"
  Case "s"
   DICD = "E13"
  Case "g"
   DICD = "O24.4"
   DiabetesName = "Gestationsdiabetes"
  Case "p"
   DICD = "R73.0"
   DiabetesName = "Pathologische Glucosetoleranz"
 End Select
 KZahl = 0
 pnpfl = 0
 anpfl = 0
 retfl = 0
 nasfl = 0
 bgwfl = 0
 niefl = 0
 ketfl = 0
 hypfl = 0
 angfl = 0
 dfsfl = 0
 niefl = False
 If Not IsNull(frm.vTextB(147)) Then
 If InStrB(frm.vTextB(147), "G63") <> 0 Or InStrB(frm.vTextB(147), "G62") <> 0 Then
  pnpfl = True
  KZahl = KZahl + 1
 ElseIf pnpflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "G99") <> 0 Then
  anpfl = True
  KZahl = KZahl + 1
 ElseIf anpflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If Not obkNeph Then
  If (InStrB(frm.vTextB(147), "N1") <> 0 Or InStrB(frm.vTextB(147), "N08.3") <> 0) Then
   niefl = True
   KZahl = KZahl + 1
  ElseIf nieflNeu Then
   KZahlNeu = KZahlNeu + 1
  End If
 End If
 If InStrB(frm.vTextB(147), "I7") <> 0 Or InStrB(frm.vTextB(147), "I6") <> 0 Then
  angfl = True
  KZahl = KZahl + 1
 ElseIf angflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "M14.2") <> 0 Then
  bgwfl = True
  KZahl = KZahl + 1
 ElseIf bgwflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "H28.0") <> 0 Then
  katfl = True
  KZahl = KZahl + 1
 End If
 If InStrB(frm.vTextB(147), "H36.0") <> 0 Then
  retfl = True
  KZahl = KZahl + 1
 ElseIf retflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "K71") <> 0 Then
  nasfl = True
  KZahl = KZahl + 1
 ElseIf nasflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "E87.2") <> 0 Then
  ketfl = True
  KZahl = KZahl + 1
 End If
 If InStrB(frm.vTextB(147), "E16") <> 0 Then
  hypfl = True
  KZahl = KZahl + 1
 End If
 If InStrB(frm.vTextB(147), "L89") <> 0 Then
  dfsfl = True
  KZahl = KZahl + 1
 ElseIf dfsflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 
 Dim oblhAlt%
 oblhAlt = 0
 If InStrB(frm.vTextB(147), "E88") <> 0 Or InStrB(frm.vTextB(147), "L94") <> 0 Then
  oblhAlt = True
  KZahl = KZahl + 1
 ElseIf oblhNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 End If
 
 syscmd 4, "Formularvorbereitung 9 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 DICDg = DICD
 If Left(DICD, 2) = "E1" Then
  DICDg = DICDg + "."
  Select Case KZahl
   Case Is > 1
    DICDg = DICDg + "7"
   Case 1
    If ketfl Then
     DICDg = DICDg + "1"
    ElseIf niefl Then
     DICDg = DICDg + "2"
    ElseIf retfl Or katfl Then
     DICDg = DICDg + "3"
    ElseIf anpfl Or pnpfl Then
     DICDg = DICDg + "4"
    ElseIf angfl Then
     DICDg = DICDg + "5"
    ElseIf nasfl Or hypfl Or oblhAlt Or dfsfl Then
     DICDg = DICDg + "6"
    End If
   Case 0
    DICDg = DICDg + "9"
  End Select
  DICDg = DICDg + "1"
 End If
 If DICDg <> "-" Then Call KRAdd(frm, DICDg, DiabetesName, DICDg, "E1", "O24.4", "R73.0")

 syscmd 4, "Formularvorbereitung 10 " & frm.adoRS!NachName & ", " & frm.adoRS!Vorname
 DICDh = DICD
 If Left(DICD, 2) = "E1" Then
  DICDh = DICDh + "."
  Select Case KZahl + KZahlNeu
   Case Is > 1
    DICDh = DICDh + "7"
   Case 1
    If ketfl Then
     DICDh = DICDh + "1"
    ElseIf niefl Or nieflNeu Then
     DICDh = DICDh + "2"
    ElseIf retfl Or retflNeu Or katfl Then
     DICDh = DICDh + "3"
    ElseIf anpfl Or pnpfl Or anpflNeu Or pnpflNeu Then
     DICDh = DICDh + "4"
    ElseIf angfl Or angflNeu Then
     DICDh = DICDh + "5"
    ElseIf nasfl Or nasflNeu Or hypfl Or oblhAlt Or oblhNeu Or dfsfl Or dfsflNeu Then
     DICDh = DICDh + "6"
    End If
   Case 0
    DICDh = DICDh + "9"
  End Select
  DICDh = DICDh + "1"
 End If
 If DICDh <> "-" And DICDg <> DICDh Then Call KRAdd(frm, DICDh, DiabetesName, DICDh, DICDh, "O24.4", "R73.0")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DiabetesDiagnose/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Diabetesdiagnose
Function ddsono(frm As AnBog)
Dim rEintr As New ADODB.Recordset
  frm.vTextB(180) = "" ' eintraege = ""
  frm.vTextB(177) = "" ' Doppler = ""
  frm.vTextB(178) = "" ' Duplex = ""
  frm.vTextB(179) = "" ' Sono = ""
  Call rEintr.Open("select * from eintraege where pat_id = " + CStr(Pat_id) + " and art in (""notiz"",""gs"",""tst"",""hz"",""cr"",""wr"",""tele"") order by zeitpunkt desc", DBCn, adOpenStatic, adLockReadOnly)
  If Not rEintr.BOF Then
   frm.vTextB(180) = rEintr!art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:") ' eintraege
   Do While Not rEintr.EOF
    frm.vTextB(180) = rEintr!art + ": " + rEintr!Inhalt + vbCrLf ' eintraege
    rEintr.Move 1
   Loop
  End If
  Set rEintr = Nothing
  Call rEintr.Open("select * from eintraege where pat_id = " + CStr(Pat_id) + " and art in (""doppler"") order by zeitpunkt desc", DBCn, adOpenStatic, adLockReadOnly)
  If Not rEintr.BOF Then
   frm.vTextB(177) = rEintr!art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:")
   Do While Not rEintr.EOF
    frm.vTextB(177) = rEintr!art + ": " + rEintr!Inhalt + vbCrLf
    rEintr.Move 1
   Loop
  End If
  Set rEintr = Nothing
  Call rEintr.Open("select * from eintraege where pat_id = " + CStr(Pat_id) + " and art in (""duplex"") order by zeitpunkt desc", DBCn, adOpenStatic, adLockReadOnly)
  If Not rEintr.BOF Then
   frm.vTextB(178) = rEintr!art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:")
   Do While Not rEintr.EOF
    frm.vTextB(178) = rEintr!art + ": " + rEintr!Inhalt + vbCrLf
    rEintr.Move 1
   Loop
  End If
  Set rEintr = Nothing
  Call rEintr.Open("select * from eintraege where pat_id = " + CStr(Pat_id) + " and art in (""sono"",""sd"") order by zeitpunkt desc", DBCn, adOpenStatic, adLockReadOnly)
  If Not rEintr.BOF Then
   frm.vTextB(179) = rEintr!art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:")
   Do While Not rEintr.EOF
    frm.vTextB(179) = rEintr!art + ": " + rEintr!Inhalt + vbCrLf
    rEintr.Move 1
   Loop
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ddsono/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ddsono
 Function PStatus%(Pat_id&, Optional frm As Object)
  Dim pul$, puk$, putp$, pudp$, pulp%, pukp%, putpp%, pudpp%, tprp As Boolean, tplp As Boolean, dprp As Boolean, dplp As Boolean
  On Error GoTo fehler
  If frm Is Nothing Then
'   Set frm = TabÖff("Anamnesebogen", "Pat_id")
'   frm.Seek "=", Pat_id
'   If frm.NoMatch Then Exit Function
   frm.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
   If frm.BOF Then Exit Function
  End If
  pul = IIf(IsNull(frm.adoRS![Puls Leiste]), "", LCase(frm.adoRS![Puls Leiste]))
  puk = IIf(IsNull(frm.adoRS![Puls Kniekehle]), "", LCase(frm.adoRS![Puls Kniekehle]))
  putp = IIf(IsNull(frm.adoRS![Puls Atp]), "", LCase(frm.adoRS![Puls Atp]))
  pudp = IIf(IsNull(frm.adoRS![Puls Adp]), "", LCase(frm.adoRS![Puls Adp]))
  If InStrB(pul, "-") <> 0 Or InStrB(pul, "mono") <> 0 Or InStrB(pul, "pst") <> 0 Or InStrB(pul, "post") <> 0 Then
   Call doGilbE(frm, "I70.20", "vTextB", 121) ' Puls Leiste
   pulp = -1
   PStatus = -1
  End If
  If InStrB(puk, "-") <> 0 Or InStrB(puk, "mono") <> 0 Or InStrB(puk, "pst") <> 0 Or InStrB(puk, "post") <> 0 Then
   Call doGilbE(frm, "I70.20", "vTextB", 122) ' Puls Kniekehle
   pukp = -1
   PStatus = -1
  End If
  Dim ttprp As Boolean, tdprp As Boolean, ttplp As Boolean, tdplp As Boolean
  Call PulsParse(putp, ttprp, ttplp)
  Call PulsParse(pudp, tdprp, tdplp)
  If (ttprp And tdprp) Or (ttplp And tdplp) Then
   PStatus = -1
   Call doGilbE(frm, "I70.20", "vTextB", 123) ' Puls Atp")
   Call doGilbE(frm, "I70.20", "vTextB", 124) ' Puls Adp")
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PStatus/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PStatus
 
 Function obLH%(Pat_id&, Optional frm As Object)
  Dim liph$
  On Error GoTo fehler
  If frm Is Nothing Then
'   Set frm = TabÖff("Anamnesebogen", "Pat_id")
'   frm.Seek "=", Pat_id
'   If frm.NoMatch Then Exit Function
   frm.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
   If frm.BOF Then Exit Function
  End If
  liph = IIf(IsNull(frm.adoRS![Liphypertrophien Abdomen]), "", LCase(frm.adoRS![Liphypertrophien Abdomen]))
  obLH = Not obNein(liph)
  If Not obLH Then
   liph = IIf(IsNull(frm.adoRS![Liphypertrophien Beine]), "", LCase(frm.adoRS![Liphypertrophien Beine]))
   obLH = Not obNein(liph)
   If Not obLH Then
    liph = IIf(IsNull(frm.adoRS![Liphypertrophien Arme]), "", LCase(frm.adoRS![Liphypertrophien Arme]))
    obLH = Not obNein(liph)
    If obLH Then
     Call doGilbE(frm, "L94.8", "vTextB", 108) 'Liphypertrophien Arme")
    End If
   Else
    Call doGilbE(frm, "L94.8", "vTextB", 107) 'Liphypertrophien Beine")
   End If
  Else
   Call doGilbE(frm, "L94.8", "vTextB", 106) 'Liphypertrophien Abdomen")
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obLH/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
 End Function ' obLH
 Function LegNPFest$(Pat_id&, Optional frm As Object)
  Dim obNeurop%, obVNeurop%, NPGrund$
  Dim re$, li$
  On Error GoTo fehler
  If frm Is Nothing Then
'   Set frm = TabÖff("Anamnesebogen", "Pat_id")
'   frm.Seek "=", Pat_id
'   If frm.NoMatch Then Exit Function
   frm.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
   If frm.BOF Then Exit Function
  End If
 
 obNeurop = False
 obVNeurop = False
 NPGrund = ""
 If Not IsNull(frm.adoRS![Vibration IK]) Then
  Call Bruchteile(frm.adoRS![Vibration IK], re, li) ' [Vibration IK]
  If IsNumeric(re) Then
   If Val(re) < 5 Then
    If Val(re) < 4 Then
     obNeurop = True
    Else
     obVNeurop = True
    End If
    Call doGilbE(frm, "G63.2", "vTextB", 119) 'Vibration IK")
   End If
  End If
  If IsNumeric(li) Then
   If Val(li) < 5 Then
    If Val(li) < 4 Then
     obNeurop = True
    Else
     obVNeurop = True
    End If
    Call doGilbE(frm, "G63.2", "vTextB", 119) 'Vibration IK")
   End If
  End If
 End If
 If Not IsNull(frm.adoRS![Vibration Großzehe]) Then
  Call Bruchteile(frm.adoRS![Vibration Großzehe], re, li) ' [Vibration Großzehe]
  If IsNumeric(re) Then
   If Val(re) < 5 Then
    If Val(re) < 4 Then
     obNeurop = True
    Else
     obVNeurop = True
    End If
    Call doGilbE(frm, "G63.2", "vTextB", 120) 'Vibration Großzehe")
   End If
  End If
  If IsNumeric(li) Then
   If Val(li) < 5 Then
    If Val(li) < 4 Then
     obNeurop = True
    Else
     obVNeurop = True
    End If
    Call doGilbE(frm, "G63.2", "vTextB", 120) 'Vibration Großzehe")
   End If
  End If
 End If
 If obNeurop Or obVNeurop Then NPGrund = "Vibr"
 
 Dim ameis$
 ameis = IIf(IsNull(frm.adoRS!Ameisenlaufen), "", LCase(frm.adoRS!Ameisenlaufen))
 If Left(ameis, 1) = "j" Then
  obNeurop = True
  NPGrund = NPGrund + IIf(NPGrund = "", "", ",") + "Ameis"
  Call doGilbE(frm, "G63.2", "vTextB", 89) 'Ameisenlaufen")
 End If

 Dim kraft$, obkraft
 obkraft = 0
 kraft = IIf(IsNull(frm.adoRS![Kraft Zehenheber]), "", LCase(frm.adoRS![Kraft Zehenheber]))
 If kraft = "vermindert" Or InStrB(kraft, "schwach") <> 0 Or InStrB(kraft, "gering ") <> 0 Then
   obVNeurop = True
   obkraft = -1
   Call doGilbE(frm, "G63.2", "vTextB", 111) 'Kraft Zehenheber")
 End If
 kraft = IIf(IsNull(frm.adoRS![Kraft Zehenbeuger]), "", LCase(frm.adoRS![Kraft Zehenbeuger]))
 If kraft = "vermindert" Or InStrB(kraft, "schwach") <> 0 Or InStrB(kraft, "gering ") <> 0 Then
   obVNeurop = True
   obkraft = -1
   Call doGilbE(frm, "G63.2", "vTextB", 112) 'Kraft Zehenbeuger")
 End If
 kraft = IIf(IsNull(frm.adoRS![Kraft Knie]), "", LCase(frm.adoRS![Kraft Knie]))
 If kraft = "vermindert" Or InStrB(kraft, "schwach") <> 0 Or InStrB(kraft, "gering ") <> 0 Then
   obVNeurop = True
   obkraft = -1
   Call doGilbE(frm, "G63.2", "vTextB", 113) 'Kraft Knie")
 End If
 If obkraft Then NPGrund = NPGrund + IIf(NPGrund = "", "", ",") + "Kra"
 
 Dim mer$, obmer As Boolean
 obmer = 0
 mer = IIf(IsNull(frm.adoRS!ASR), "", LCase(frm.adoRS!ASR))
 If (InStrB(mer, "-") <> 0 And InStrB(mer, "?") <> 0) Or mer = "-" Or mer = "-/-" Then
  obVNeurop = True
  obmer = True
  Call doGilbE(frm, "G63.2", "vTextB", 114) 'ASR")
 End If
 mer = IIf(IsNull(frm.adoRS!PSR), "", LCase(frm.adoRS!PSR))
 If (InStrB(mer, "-") <> 0 And InStrB(mer, "?") <> 0) Or mer = "-" Or mer = "-/-" Then
  obVNeurop = True
  obmer = True
  Call doGilbE(frm, "G63.2", "vTextB", 115) 'PSR")
 End If
 If obmer Then NPGrund = NPGrund + IIf(NPGrund = "", "", ",") + "MER"
  
 Dim ofl$, obofl As Boolean
 ofl = IIf(IsNull(frm.adoRS!Oberflächensensibilität), "", LCase(frm.adoRS!Oberflächensensibilität))
 If InStrB(ofl, "gestört") <> 0 Or InStrB(ofl, "paret") <> 0 Or InStrB(ofl, "vermind") <> 0 Or InStrB(ofl, "path") <> 0 Or InStrB(ofl, "pelz") <> 0 Or InStrB(ofl, "strumpf") <> 0 Or InStrB(ofl, "schwäch") <> 0 Or InStrB(ofl, "eingeschr") <> 0 Then
  obNeurop = True
  NPGrund = NPGrund + IIf(NPGrund = "", "", ",") + "Ofl"
  Call doGilbE(frm, "G63.2", "vTextB", 116) 'Oberflächensensibilität")
 End If
 
 Dim Monf$
 Monf = IIf(IsNull(frm.adoRS!Monofilamenttest), "", LCase(frm.adoRS!Monofilamenttest))
 If Left(Monf, 1) = "0" Or InStrB(Monf, " 0") <> 0 Or InStrB(Monf, "1") <> 0 Or InStrB(Monf, "2") <> 0 Or (InStrB(Monf, "3") <> 0 And InStrB(Monf, "/3") = 0) Or InStrB(Monf, "4") <> 0 Or InStrB(Monf, "6/8") <> 0 Or InStrB(Monf, "5/8") <> 0 Or InStrB(Monf, "gestört") <> 0 Or InStrB(Monf, "5/10") <> 0 Or InStrB(Monf, "6/10") <> 0 Or InStrB(Monf, "7/10") <> 0 Or InStrB(Monf, "8/10") <> 0 Or InStrB(Monf, "neg") <> 0 Or InStrB(Monf, "path") <> 0 Or InStrB(Monf, " -") <> 0 Or InStrB(Monf, "zufäl") <> 0 Or InStrB(Monf, "erat") <> 0 Then
  obNeurop = True
  NPGrund = NPGrund + IIf(NPGrund = "", "", ",") + "MF"
  Call doGilbE(frm, "G63.2", "vTextB", 117) 'Monofilamenttest")
 End If
 
 Dim KW$
 KW = IIf(IsNull(frm.adoRS![Kalt-Warm]), "", LCase(frm.adoRS![Kalt-Warm]))
 If Left(KW, 1) = "0" Or InStrB(KW, " 0") <> 0 Or InStrB(KW, "1") <> 0 Or InStrB(KW, "2") <> 0 Then
' Or instrb(Monf, "6/8") <> 0 Or instrb(Monf, "5/8") <> 0 Or instrb(Monf, "6/8") <> 0 Or instrb(Monf, "5/10") <> 0 Or instrb(Monf, "6/10") <> 0 Or instrb(Monf, "7/10") <> 0 Or instrb(KW, "gestört") <> 0 Or instrb(KW, "neg") <> 0 Or instrb(KW, "path") <> 0 Or instrb(KW, " -") <> 0 Or instrb(KW, "zufäl") <> 0 Or instrb(KW, "erat") <> 0 Then ' Or instrb(KW, "4") <> 0 Or instrb(Monf, "8/10") <> 0
  obVNeurop = True
  NPGrund = NPGrund + IIf(NPGrund = "", "", ",") + "KW"
  Call doGilbE(frm, "G63.2", "vTextB", 118) 'Kalt-Warm")
 End If
 LegNPFest = NPGrund
' damit nur ein Parameter zugegeben zu werden braucht
 If obVNeurop Then
  LegNPFest = "V." + LegNPFest
 ElseIf obNeurop Then
  LegNPFest = "->" + LegNPFest
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LegNPFest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LegNPFest
#End If
' wird sowohl von form_current als auch von einer dynamischen SQL-Abfrage aufgerufen, deshalb zwei Aufrufmöglichkeiten
 Function hkGrund$(Pat_id&, Optional rsNa As Object)
  Dim druckst$, hyperkerat$, obDst%, obHk%
'  Dim rAna As DAO.Recordset
  On Error GoTo fehler
   If rsNa Is Nothing Then
'    Set rsNa = TabÖff("Anamnesebogen", "Pat_id")
'    rsNa.Seek "=", Pat_id
'    If rsNa.NoMatch Then Exit Function
    Set rsNa = New ADODB.Recordset
    rsNa.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
    If rsNa.BOF Then Exit Function
   End If
'  If Not rsNa.NoMatch Then
   obDst = 0
   obHk = 0
   hkGrund = ""
   druckst = IIf(IsNull(rsNa!Druckstellen), "", LCase(rsNa!Druckstellen))
   hyperkerat = IIf(IsNull(rsNa!Hyperkeratosen), "", LCase(rsNa!Hyperkeratosen))
   hyperkerat = Trim$(hyperkerat)
   If Left(druckst, 1) = "j" Then
    obDst = -1
    hkGrund = "Frgbg."
   End If
   If (hyperkerat <> "" And hyperkerat <> "-" And hyperkerat <> "n" And hyperkerat <> "n |" And hyperkerat <> "- | -" And hyperkerat <> "|" And hyperkerat <> "n | n") And hyperkerat <> "keine | keine" And hyperkerat <> "|" Then
    obHk = -1
    hkGrund = IIf(hkGrund = "", "", hkGrund + ",") + "Befd"
   End If
 ' End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in hkGrund/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
 End Function ' hkGrund
#If False Then
' wird von SQL-Abfragen aufgerufen
  Function EigTAlter$(Pat_id&, Eig$, vz$, Grenze$)
   Dim Alter$, GebDat#, obErfüllt%, rsNa As New ADODB.Recordset
   On Error GoTo fehler
'   Set rNa = TabÖff("Namen", "Pat_id")
'   rNa.Seek "=", Pat_id
   Call rsNa.Open("select * from namen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly)
   If Not rsNa.BOF Then
    GebDat = rsNa!GebDat
    Alter = (Date - GebDat) * 0.0027379
    Select Case vz
     Case ">"
      If Alter > Grenze Then obErfüllt = -1
     Case "<"
      If Alter < Grenze Then obErfüllt = -1
    End Select
    If obErfüllt Then
     EigTAlter = EigT(Pat_id, Eig)
    End If
   End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EigTAlter/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
  End Function
#End If
#If False Then
' wird von SQL-Abfragen aufgerufen
  Function EigT$(Pat_id&, Eig$) ' Eigenschaft testen
   Static sql$, altpat&
   Static md As dao.Recordset
   On Error GoTo fehler
   Call dtbInit
   If Pat_id = 0 Then Exit Function
   If Pat_id <> altpat Then
    sql = "select * from [" + QMdbAkt + "].[medikamente mit Arten] where medplan.pat_id =" + CStr(Pat_id)
    Set md = Dtb.OpenRecordset(sql, dbOpenDynaset)
    altpat = Pat_id
   End If
   If Not md.BOF Then
    md.MoveFirst
    Do While Not md.EOF
     Call mdTest(md, Eig, EigT)
     md.Move 1
    Loop
   End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EigT/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
  End Function
#End If
 Function PfadFestLeg$(art$, muster1$, frm As Form, Optional muster2$)
  On Error GoTo fehler
  sql = "SELECT * from dokumente where pat_id = " + CStr(frm.adoRS!Pat_id) + " and (dokname like " & "'" & muster1 & "'" & IIf(LenB(muster2) = 0, vbNullChar, " or dokname like " & "'" & muster2 + "'") + ") order by zeitpunkt desc"
  On Error GoTo fehler
'  Call dtbInit
'  Set rDok = Dtb.OpenRecordset(sql, dbOpenDynaset)
  Set raDok = Nothing
  Exit Function
  Call raDok.Open(sql, DBCn, adOpenStatic, adLockReadOnly)
  If PcDokPfad = "" Then Call getDokPfad
  Dim ctl, fehler&
  For Each ctl In frm.Controls
   Dim dataf$
   On Error Resume Next
   Err.Clear
   dataf = ctl.DataFields
   fehler = Err.Number
   On Error GoTo fehler
   If fehler = 0 Then
    If ctl.DataField Like art Then
     If raDok.BOF Or PcDokPfad = "" Then
      frm.Controls(ctl).Enabled = False
     Else
      frm.Controls(ctl).Enabled = True
  '   raDok.MoveFirst
      PfadFestLeg = raDok!DokPfad
      raDok.Close
     End If
    End If
    Exit For
   End If
  Next ctl
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PfadFestLeg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
 End Function
Function Urineintraege%(Pat_id&)
 Dim raUr As ADODB.Recordset, pZahlStix%, pZahlLab%, nZahlStix%, nZahlLab%, nZahlMic%, pZahlMic%, Ps$(), i%
 On Error GoTo fehler
 Set raUr = New ADODB.Recordset
 Call raUr.Open("Select * from eintraege where art = ""urin"" and pat_id = " & CStr(Pat_id), DBCn, adOpenDynamic, adLockReadOnly)
 Do While Not raUr.EOF
  Ps = Split(raUr!Inhalt, ",")
   For i = 0 To UBound(Ps())
    If InStrB(Ps(i), "Eiwei") <> 0 Then
     If InStrB(Ps(i), "+") <> 0 Then
      pZahlStix = pZahlStix + 1
     ElseIf InStrB(Ps(i), "-") <> 0 Then
      nZahlStix = nZahlStix + 1
     End If
    ElseIf InStrB(Ps(i), "Micral") <> 0 Then
     If (InStrB(Ps(i), "20") <> 0 Or InStrB(Ps(i), "50") <> 0) Then
      pZahlMic = pZahlMic + 1
     ElseIf InStrB(Ps(i), "neg") <> 0 Then
      nZahlMic = nZahlMic + 1
     End If
    End If
   Next i
  raUr.MoveNext
 Loop
 Dim LW&
' rlau.Seek "=", CStr(Pat_id), "albcre"
' raUr.Close
 Set raUr = New ADODB.Recordset
 Call raUr.Open("select * from laborneu left join laborkommentar using (kommentarvw) where abkü in (""ALBCRE"",""ALBKRE"",""ALBQ"") and pat_id = " & CStr(Pat_id), DBCn, adOpenDynamic, adLockReadOnly)
 If Not raUr.BOF Then
  Do While Not raUr.EOF
   LW = Val(Replace(Replace(IIf(IsNull(raUr!Wert), IIf(IsNull(raUr!Kommentar), "", raUr!Kommentar), raUr!Wert), ",", "."), "%", ""))
   If LW < 20 Then
    nZahlLab = nZahlLab + 1
   Else
    pZahlLab = pZahlLab + 1
   End If
   raUr.Move 1
  Loop
 End If
 Urineintraege = pZahlStix + pZahlMic - nZahlMic - nZahlMic - nZahlLab - nZahlLab - nZahlLab ' nzahlmic zählt doppelt, nzahllab dreifach
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Urineintraege/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Urineintraege
Function doXtra_Click(frm As AnBog)
 Dim rAF&
 On Error GoTo fehler
 If UStumm Then Exit Function
 If frm.Xtra <> "" Then
  Call DBCn.Execute("update fuerdiagexp set icd = concat(left(icd,instr(icd,'.')-1),'" & Left(frm.Xtra, 3) & "') where pat_id = " & frm.adoRS!Pat_id & " order by id desc limit 1", rAF)  ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call DBCn.Execute("update fuerdiagexp set diagnose = mid$(diagnose,6) where pat_id = " & frm.adoRS!Pat_id & " and (diagnose like 'Z.n.%' or diagnose like 'Z.n.%V.a.%')  order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call DBCn.Execute("update fuerdiagexp set diagnose = mid$(diagnose,6) where pat_id = " & frm.adoRS!Pat_id & " and (diagnose like 'V.a.%' or diagnose like 'V.a.%Z.n.%')  order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call DBCn.Execute("update fuerdiagexp set diagnose = mid$(diagnose,6) where pat_id = " & frm.adoRS!Pat_id & " and (diagnose like 'Z.n.%' or diagnose like 'Z.n.%V.a.%')  order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call DBCn.Execute("update fuerdiagexp set diagnose = mid$(diagnose,6) where pat_id = " & frm.adoRS!Pat_id & " and (diagnose like 'V.a.%' or diagnose like 'V.a.%Z.n.%')  order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  UStumm = True
  frm.Va = 0
  frm.Zn = 0
  UStumm = False
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doVa_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doXtra_Click
Function doVa_Click(frm As Form)
 Dim rAF&
 On Error GoTo fehler
 If UStumm Then Exit Function
 If frm.Va Then
  Call DBCn.Execute("update fuerdiagexp set icd = concat(icd,'V') where pat_id = " & frm.adoRS!Pat_id & " and (not icd like '%V' and not icd like '%VZ')  order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call DBCn.Execute("update fuerdiagexp set diagnose = concat('V.a. ',diagnose) where pat_id = " & frm.adoRS!Pat_id & " and (not diagnose like 'V.a.%' and not diagnose like 'V.a.%Z.n.%')  order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 Else
  Call DBCn.Execute("update fuerdiagexp set icd = left(icd,length(icd)-1) where pat_id = " & frm.adoRS!Pat_id & " and (icd like '%V' or icd like '%VZ') order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call DBCn.Execute("update fuerdiagexp set diagnose = ltrim(mid(diagnose,5)) where pat_id = " & frm.adoRS!Pat_id & " and (diagnose like 'V.a.%' or diagnose like 'V.a.%Z.n.%') order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doVa_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doVa_Click
Function doZn_Click(frm As Form)
 Dim rAF&
 On Error GoTo fehler
 If UStumm Then Exit Function
 If frm.Zn Then
  Call DBCn.Execute("update fuerdiagexp set icd = concat(icd,'Z') where pat_id = " & frm.adoRS!Pat_id & " and (not icd like '%Z' and not icd like '%ZV') order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call DBCn.Execute("update fuerdiagexp set diagnose = concat('Z.n. ',diagnose) where pat_id = " & frm.adoRS!Pat_id & " and (not diagnose like 'Z.n.%' and not diagnose like 'Z.n.%V.a.%') order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 Else
  Call DBCn.Execute("update fuerdiagexp set icd = left(icd,length(icd)-1) where pat_id = " & frm.adoRS!Pat_id & " and (icd like '%Z' or icd like '%ZV') order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call DBCn.Execute("update fuerdiagexp set diagnose = ltrim$(mid$(diagnose,5)) where pat_id = " & frm.adoRS!Pat_id & " and (diagnose like 'Z.n.%' or diagnose like 'Z.n.%V.a.%')  order by id desc limit 1", rAF) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doZn_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doZn_Click
Function knöpfeanpassen(frm As AnBog)
 Dim nr%, i%
 Dim rfDE As New ADODB.Recordset
 On Error GoTo fehler
 UStumm = True
 frm.Va = 0
 frm.Zn = 0
 frm.Xtra = ""
 For nr = 1 To 23
' Hier noch einfügen, wenn Tabelle rfDE abgefragt werden soll
  Dim icdlike$
  If MDIICD(nr) = "" Then
   If frm.Controls("vOptionB")(nr).Value <> 0 Then frm.Controls("vOptionB")(nr).Value = False      ' "u"
  Else
   If MDIICD(nr) Like "E1*" Or MDIICD(nr) Like "L89*" Then
    icdlike = Left(MDIICD(nr), InStr(MDIICD(nr), ".")) & ".*"
   Else
    icdlike = MDIICD(nr)
    For i = 1 To 2
     If Right$(icdlike, 1) = "V" Then icdlike = Left(icdlike, Len(icdlike) - 1)
     If Right$(icdlike, 1) = "Z" Then icdlike = Left(icdlike, Len(icdlike) - 1)
    Next i
    icdlike = icdlike & "[VZ]?"
   End If
   Set rfDE = Nothing
   Call rfDE.Open("select * from fuerdiagexp where pat_id = " & frm.adoRS!Pat_id & " and icd rlike '" & icdlike & "' and diagnose like '%" & LTrim$(Replace(LTrim$(Replace(MDIDiag(nr), "V.a.", "")), "Z.n.", "")) & "'", DBCn, adOpenDynamic, adLockReadOnly) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
'  rfDE.Seek "=", frm.adors!Pat_id, MDIICD(nr), MDIDiag(nr)
   If (rfDE.BOF Or Trim$(MDIICD(nr)) = "") Then
    If frm.Controls("vOptionB")(nr).Value <> 0 Then frm.Controls("vOptionB")(nr).Value = 0  ' "u"
   Else
    If frm.Controls("vOptionB")(nr).Value = 0 Then frm.Controls("vOptionB")(nr).Value = 1 ' "u"
   End If
  End If
 Next nr
 Dim diabzahl&, passnr%
 diabzahl = 0
 passnr = 0
 For nr = 1 To 23
  If MDIICD(nr) Like "E1*" Then
   diabzahl = diabzahl + 1
  End If
 Next nr
 If diabzahl > 1 Then
  For nr = 1 To 23
   If MDIICD(nr) Like "E1*" Then
    icdlike = Left(MDIICD(nr), 6) & "%"
    Set rfDE = Nothing
    Call rfDE.Open("select * from fuerdiagexp where pat_id = " & frm.adoRS!Pat_id & " and icd like '" & icdlike & "'", DBCn, adOpenDynamic, adLockReadOnly) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
    If Not rfDE.EOF Then
     passnr = nr
     Exit For
    End If
   End If
  Next nr
  If passnr = 0 Then
   For nr = 1 To 23
    If MDIICD(nr) Like "E1*" Then
     diabzahl = diabzahl - 1
     frm.Controls("vOptionB")(nr).Value = 0
     If diabzahl = 1 Then Exit For
    End If
   Next nr
  Else
   For nr = 1 To 23
    If MDIICD(nr) Like "E1*" Then
     If nr <> passnr Then
      diabzahl = diabzahl - 1
      frm.Controls("vOptionB")(nr).Value = 0
     End If
    End If
   Next nr
  End If
 End If ' diabzahl > 0
 UStumm = 0
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in knöpfeanpassen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' knöpfeanpassen
Function RRpAll()
Dim Pat_id&, rsNam As New ADODB.Recordset
On Error GoTo fehler
Call Lese.ProgStart
Call rsNam.Open("select pat_id from namen order by pat_id", DBCn, adOpenStatic, adLockReadOnly)
Do While Not rsNam.EOF
 Call RRParse(rsNam!Pat_id)
 rsNam.Move 1
Loop
Call Lese.ProgEnde
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in RRpAll/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'RRpAll
Function RRParse(Pat_id&)
 Dim rsRR As New ADODB.Recordset, rsP As New ADODB.Recordset, rsNa As New ADODB.Recordset
 On Error GoTo fehler
 Call DBCn.Execute("delete from rrparse where pat_id = " & Pat_id)
 rsRR.Open "select * from rr where pat_id = " & CStr(Pat_id), DBCn, adOpenDynamic, adLockReadOnly
 Do While Not rsRR.EOF
  Call do_RRParse(rsRR!RR, rsRR!Pat_id, rsRR!Zeitpunkt, "Tabelle RR")
  rsRR.MoveNext
 Loop
' Set raNa = TabÖff("Anamnesebogen", "Pat_id")
 Set rsNa = Nothing
 rsNa.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockOptimistic
' raNa.Seek "=", Pat_id
 If Not rsNa.EOF Then
  If Not IsNull(rsNa!Blutdruckwerte) Then
   Call do_RRParse(rsNa!Blutdruckwerte, Pat_id, IIf(IsNull(rsNa!Vorgestellt), 0, rsNa!Vorgestellt), "An'bg.BW")
  End If
  If Not IsNull(rsNa!RR) Then
   Call do_RRParse(rsNa!RR, Pat_id, rsNa!Vorgestellt, "An'bg.RR")
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in RRParse/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' RRParse()
Function do_RRParse(erg$, ByVal Pat_id, Zeitpunkt As Date, Quelle$)
  Dim RRsyst%, RRdiast%, Zp As Date
  Dim rsP As New ADODB.Recordset
  On Error GoTo fehler
  If dodoRRParse(erg, RRsyst, RRdiast, Zp) = 0 Then Exit Function
'  rsP.Seek "=", Pat_id, IIf(Zp = 0, Zeitpunkt, Zp), RRSyst, RRDiast
  rsP.Open "select 0 from rrparse where pat_id = " & Pat_id & " and zeitpunkt = " & datform(IIf(Zp = 0, Zeitpunkt, Zp)) & " and rrsyst = " & RRsyst & " and rrdiast = " & RRdiast, DBCn, adOpenStatic, adLockReadOnly
  If rsP.BOF Then
   Call DBCn.Execute("insert into rrparse(pat_id,zeitpunkt,rrsyst,rrdiast,quelle) values(" & Pat_id & "," & datform(IIf(Zp = 0, Zeitpunkt, Zp)) & "," & RRsyst & "," & RRdiast & ",'" & Replace(Quelle, "'", "''") & "')")
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_RRParse/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_RRParse(erg$, rsP as dao.recordset, Pat_id$, ZeitPunkt As Date)
Public Function dodoRRParse(ByVal erg$, RRsyst%, RRdiast%, Optional Zp As Date)
  Dim i%, T1$, T1a$, pos%, runde%
  On Error GoTo fehler
  dodoRRParse = -1
  On Error GoTo fehler
  runde = 0
  Do
   runde = runde + 1
   RRsyst = 0
   RRdiast = 0
   Zp = 0
   pos = InStr(erg, "/")
   If pos > 1 Then
    T1 = Left(erg, pos - 1)
    If Len(T1) > 2 Then
     If Mid$(T1, Len(T1) - 2, 1) = "," And InStrB("0123456789", Mid$(T1, Len(T1) - 1, 1)) <> 0 Then T1 = Left(T1, Len(T1) - 2)
    End If
    T1a = Right$(T1, 3)
    If IsNumeric(T1a) Then
     RRsyst = Val(T1a)
    Else
     T1a = Right$(T1a, 2)
     If IsNumeric(T1a) Then
      RRsyst = Val(T1a)
     End If
    End If
    erg = Mid$(erg, pos + 1)
    T1a = Left(erg, 3)
    If IsNumeric(T1a) Then
     RRdiast = Val(T1a)
     erg = Mid$(erg, 4)
    Else
     T1a = Left(T1a, 2)
     If IsNumeric(T1a) Then
      RRdiast = Val(T1a)
      erg = Mid$(erg, 2)
     End If
    End If
    If InStr(erg, "(") > InStr(erg, "/") Then
     T1a = Mid$(erg, InStr(erg, "("))
     If InStr(T1a, ")") > 1 Then
      T1a = Left(T1a, InStr(T1a, ")") - 1)
      If IsDate(T1a) Then Zp = CDate(T1a)
     End If
    End If
   Else ' pos > 1
    pos = InStr(erg, "syst")
    If pos > 1 Then
     T1 = Left(erg, pos - 1)
     For i = Len(T1) To 0 Step -1
      If InStrB("0123456789", Mid$(T1, i, 1)) <> 0 And Mid$(T1, i, 1) <> "" Then Exit For
      T1 = Left(T1, Len(T1) - 1)
      If T1 = "" Then Exit For
     Next
     If Len(T1) > 2 Then
      If Mid$(T1, Len(T1) - 2, 1) = "," And InStrB("0123456789", Mid$(T1, Len(T1) - 1, 1)) <> 0 And Mid$(T1, Len(T1) - 1, 1) <> "" Then T1 = Left(T1, Len(T1) - 2)
     End If
     If T1 = "" Then GoTo nix
     T1a = Right$(T1, 3)
     If IsNumeric(T1a) Then
      RRsyst = Val(T1a)
     Else
      T1a = Right$(T1a, 2)
      If IsNumeric(T1a) Then
       RRsyst = Val(T1a)
      End If
      If RRsyst < 50 Then
       T1a = Left(T1a, Len(T1a) - 2)
       For i = Len(T1a) To 0 Step -1
        If InStrB("0123456789", Mid$(T1a, i, 1)) <> 0 And Mid$(T1a, i, 1) <> "" Then Exit For
        T1 = Left(T1a, Len(T1a) - 1)
       Next
       If Len(T1a) > 2 Then
        If Mid$(T1a, Len(T1a) - 2, 1) = "," And InStrB("0123456789", Mid$(T1a, Len(T1a) - 1, 1)) <> 0 And Mid$(T1a, Len(T1a) - 1, 1) <> "" Then T1a = Left(T1a, Len(T1a) - 2)
       End If
       T1a = Right$(T1a, 3)
       If IsNumeric(T1a) Then
        RRsyst = Val(T1a)
       End If
      End If
     End If
     erg = Mid$(erg, pos + 4)
    Else ' pos > 1
     erg = ""
    End If ' pos > 1
   End If ' pos > 1
   If RRsyst > 0 Then Exit Do
   If erg = "" Then
nix:
     dodoRRParse = 0
     Exit Do
   End If
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodoRRParse/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function mdTest(md As ADODB.Recordset, Feld$, Optional ByRef bisher$)
 Dim medi$
 On Error GoTo fehler
 If Not IsNull(md.Fields(Feld).Value) Then
  If md.Fields(Feld).Value Then
   medi = Left(md![Medikament], 15)
   If IsNull(bisher) Or bisher = "" Then
     bisher = medi
   Else
    If Not InStrB(bisher, medi) <> 0 Then
     bisher = bisher & ", " & medi
    End If
   End If
  End If
 End If
 mdTest = bisher
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in mdTEst/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' mdTest
#If False Then
Function mdTest_alt$(md As dao.Recordset, Feld$)
 On Error GoTo fehler
 If Not IsNull(md.Fields(Feld$).Value) Then
  If md.Fields(Feld$).Value Then
   If IsNull(mdTest_alt) Or mdTest_alt = "" Then
     mdTest_alt = md!medi
   Else
     mdTest_alt = mdTest_alt & ", " & md!medi
   End If
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in mdTest_alt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' mdTest
#End If
Public Function do_An1Aufruf_click(frm As Form)
 On Error GoTo fehler
 If An1Pfad = "" Then An1Pfad = PfadFestLeg("An1Aufruf", "%anamnese%1%", frm)
 Call do_do_Aufruf(frm, An1Pfad)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_An1Aufruf_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_An1Aufruf_click(frm As Form)
Public Function do_An2Aufruf_click(frm As Form)
 On Error GoTo fehler
 If An2Pfad = "" Then An2Pfad = PfadFestLeg("An2Aufruf", "%anamnese%2%", frm)
 Call do_do_Aufruf(frm, An2Pfad)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_An2Aufruf_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function do_AnAAufruf_click(frm As Form)
 On Error GoTo fehler
 If AnAPfad = "" Then AnAPfad = PfadFestLeg("AnAAufruf", "%anamnese%allg%", frm, "*allg*anamnese*")
 Call do_do_Aufruf(frm, AnAPfad)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_AnAAufruf_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function do_CheckAufruf_click(frm As Form)
 On Error GoTo fehler
 If CheckPfad = "" Then CheckPfad = PfadFestLeg("CheckAufruf", "%che%kliste%", frm)
 Call do_do_Aufruf(frm, CheckPfad)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_CheckAufruf_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Function do_do_Aufruf(frm As Form, DPfad$)
 On Error GoTo fehler
 If (InStrB("$\", Left(DPfad, 1)) <> 0 And Left(DPfad, 1) <> "") Or Mid$(DPfad, 2, 1) = ":" Then
   If PcDokPfad = "" Then getDokPfad
   DatPfad = Environ("ProgramFiles") & "\irfanview\i_view32.exe " & "" & Stutz(Replace(LCase(DPfad), "$\turbomed\dokumente", PcDokPfad)) + ""
   Call Shell(DatPfad, vbMaximizedFocus)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_do_Aufruf/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function inActi(ct As Control)
  On Error Resume Next
  ct.BackColor = 12632256 ' grau
  On Error GoTo fehler
  ct.Enabled = False
  On Error Resume Next
  ct.Locked = True
  On Error GoTo fehler
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in inActi/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function Acti(ct As Control)
  On Error Resume Next
  ct.BackColor = 16777215 ' weiß
  ct.Locked = False
  On Error GoTo fehler
  ct.Enabled = True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Acti/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

Public Function do_Hirndurchblutungsstörung_Exit(Cancel%, frm As Form)
 On Error GoTo fehler
 If frm.adoRS!Hirndurchblutungsstörung = "n" Then ' Hirndurchblutungsstörung
  Call inActi(frm.vTextB(83))  ' Schlaganfall
 Else
  Call Acti(frm.vTextB(83))
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Hirndurchblutungsstörung_Exit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
#If mitab Then
Function doGilbE(frm As Object, ZielICD$, gilbFeld$, gilbNr%) ' gilb einzeln (ohne kradd)
 If InStrB(frm.vTextB(147), ZielICD) = 0 Then ' Diagnosen
  Call doGilb(frm, ZielICD, gilbFeld, gilbNr)
 End If
End Function ' doGilbE

Function doGilb(frm As Object, ZielICD$, gilbFeld$, gilbNr%)   ' As Form_Anamnesebogen
  Dim i%, rVorh As New ADODB.Recordset
  On Error GoTo fehler
     If gilbFeld <> "" Then
      Dim icdlike$ ' Replace(ZielICD, "V", "") & "[V]?"
      icdlike = Replace(ZielICD, "V", "") & "[V]?"
      If icdlike Like "E1*" Or icdlike Like "L89*" Then
       icdlike = Left(icdlike, InStr(icdlike, ".")) & ".*"
      Else
       icdlike = Replace(ZielICD, "V", "") & "[V]?"
      End If
      Call rVorh.Open("select * from fuerdiagexp where pat_id = " & frm.adoRS!Pat_id & " and icd rlike '" & icdlike & "'", DBCn, adOpenStatic, adLockReadOnly)
'      If rVorh.BOF Then
'       Set rVorh = Nothing
'       Call rVorh.Open("select * from fuerdiagexp where pat_id = " & frm.adoRS!Pat_id & " and icd rlike '" & ZielICD & "V'", DBCn, adOpenStatic, adLockReadOnly)
'      End If
''      If rVorh.NoMatch Then rVorh.Seek "=", frm.Pat_id, ZielICD + "Z" das wird zu viel
      If rVorh.BOF Then
       On Error Resume Next
       If frm.Controls(gilbFeld)(gilbNr).ControlType <> 106 Then
        frm.Controls(gilbFeld)(gilbNr).BackColor = MarkierFarbe
       End If
       On Error GoTo fehler
      End If
     End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doGilb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doGilb
 Function KRAdd(frm As Form, ByVal flag$, ByVal KrStr$, ZielICD$, Optional ICD1$, Optional ICD2$, Optional ICD3$, Optional ICD4$, Optional gilbFeld$, Optional gilbNr%)
  Dim rechts$
  On Error GoTo fehler
  rechts = ""
  labPos = 0
  If flag <> "" And Not IsNull(flag) Then
    If ICD1 = "" Then ICD1 = ZielICD
    If Not EinRueck(frm, ICD1, KrStr, ICD2, ICD3, ICD4) Then '  And flag = "Fragebogen") Then
     frm.vTextB(168) = frm.vTextB(168) + KrStr + " (" + flag + ")" & vbCrLf ' MDi
     Call doGilb(frm, ZielICD, gilbFeld, gilbNr)
     MDIICD(MDIn) = ZielICD
     rechts = Right$(ZielICD, 1)
     MDIDiag(MDIn) = KrStr
     If Left(KrStr, 3) = "V.a" Then
      If rechts <> "V" Then MDIICD(MDIn) = MDIICD(MDIn) + "V"
     ElseIf Left(KrStr, 3) = "Z.n" Then
      If rechts <> "Z" Then MDIICD(MDIn) = MDIICD(MDIn) + "Z"
     ElseIf Left(KrStr, 6) = "Aussch" Then
      If rechts <> "A" Then MDIICD(MDIn) = MDIICD(MDIn) + "A"
     End If
     If rechts = "V" Then
      labPos = 1
     ElseIf rechts <> "Z" And rechts <> "A" Then
      labPos = -1
     End If
     MDIn = MDIn + 1
    End If
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in KRAdd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
 End Function ' kradd(frm,ByVal ob$, KrStr$, Optional icd1$, Optional icd2$)
 
 ' Wenn Laborparameter mit dem Wert "0" nicht wirklich 0 bedeuten ...
Function LabWert0!(frm As Form, Name$, Optional sign, Optional Grenze!, Optional Diagnose$, Optional ZielICD$, Optional Einheit$, Optional ICD1$, Optional ICD2$, Optional SGrenze, Optional WGrenze, Optional ICD3$) ' sichere Grenze
 On Error GoTo fehler
 LabWert0 = 0
 'rLa.Seek "=", CStr(Pat_id), name
 raLau.Find "abkü = " & "'" & Name & "'"
 If Not raLau.EOF Then
  If Val(raLau!Wert) = 0 Then
  Else
   LabWert0 = LabWert(frm, Name$, sign, Grenze!, Diagnose$, ZielICD$, Einheit$, ICD1$, ICD2$, SGrenze, WGrenze, ICD3$)
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabWert0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LabWert0
' die folgenden Funktionen sind für Fremdlabor und Augenbefunde identisch
Function LabWert!(frm As Form, Name$, Optional sign, Optional ByVal Grenze!, Optional Diagnose$, Optional ZielICD$, Optional Einheit$, Optional ICD1$, Optional ICD2$, Optional SGrenze, Optional WGrenze, Optional ICD3$) ' sichere Grenze, weibliche Grenze
 Dim rbLau As New ADODB.Recordset
 sql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from (" & laborAbfr & " where pat_id = " & Pat_id & " and abkü = '" & Name & "' and wert <> '') as labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar, Normbereich as NB " + _
 "FROM laborxus LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " & _
 "WHERE pat_id = " & CStr(frm.adoRS!Pat_id) & "  and abkü = '" & Name & "' and wert <> '' and not exists (select * from laborneu where pat_id = " & CStr(frm.adoRS!Pat_id) & " and abkü = laborxwert.Abkü and wert = laborxwert.wert and zeitpunkt > laborxus.Eingang-3 and zeitpunkt <  laborxus.Eingang+6) "
 sql1 = "ORDER BY zeitpunkt desc"
 
' Set rblau = Dtb.OpenRecordset(sql)
 Set rbLau = Nothing
 rbLau.Open LCase(sql) & sql1, DBCn, adOpenDynamic, adLockReadOnly
 
 On Error GoTo fehler
 LabWert = 0
 'rlau.Seek "=", CStr(Pat_id), name
' rblau.Find "abkü = " & "" & Name & "" & " and wert <> """"", 0, adSearchForward, 1
 If Not rbLau.EOF Then
  LabWert = Val(Replace$(Replace(IIf(IsNull(rbLau!Wert), IIf(IsNull(rbLau!Kommentar), "", rbLau!Kommentar), rbLau!Wert), ",", "."), "%", ""))
  If Alter = 0 Then Alter = (rbLau!Zeitpunkt - GebDat) * 2.73792574745373E-03 ' 1/365,24
  If Not IsMissing(sign) Then
   If Not IsMissing(WGrenze) And obWeib Then Grenze = WGrenze
   If Not IsMissing(SGrenze) Then
    If Not ((sign = "<" And LabWert < SGrenze) Or (sign = ">" And LabWert > SGrenze)) Then
      Diagnose = "V.a. " + Diagnose
    End If
   End If
   If (sign = "<" And LabWert < Grenze) Or (sign = ">" And LabWert > Grenze) Then
    Call KRAdd(frm, UCase$(Name) & " " & CStr(LabWert) & " " & Einheit, Diagnose, ZielICD, ICD1, ICD2, ICD3)
   End If
  End If
 End If
' rlau.MoveFirst
' Do
'  If rlau!Pat_id = Pat_id Then
'   Debug.Print rlau!Abkü, rlau!Wert
'  End If
'  rlau.MoveNext
' Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabWert/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LabWert
#End If
Function EinRueck%(frm As Form, ICD1$, Diagnose$, Optional ICD2$, Optional ICD3$, Optional ICD4$)
  On Error GoTo fehler
' 15.6.: anam!diagnosen statt anam!diagnosen
  If ICD1 = "" Then Exit Function
'  On Error Resume Next
  If InStrB(frm.vTextB(147), ICD1) = 0 Then
   If ICD2 = "" Then Exit Function
   If ICD2 = ICD1 Then Exit Function
   If InStrB(frm.vTextB(147), ICD2) = 0 Then
    If ICD3 = "" Then Exit Function
    If ICD3 = ICD2 Or ICD3 = ICD1 Then Exit Function
    If InStrB(frm.vTextB(147), ICD3) = 0 Then
     If ICD4 = "" Then Exit Function
     If ICD4 = ICD3 Or ICD4 = ICD2 Or ICD4 = ICD1 Then Exit Function
     If InStrB(frm.vTextB(147), ICD4) = 0 Then Exit Function
    End If
   End If
  End If
  Diagnose = "  " + Diagnose
  EinRueck = -1
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EinRueck/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Einrück
Public Function do_Fremdlabor_Form_Current(frm As Form)
  On Error GoTo fehler
'  If raNa Is Nothing Then
'   Call do_Fremdlabor_Form_Load(frm)
'  End If
'  rsNa.Seek "=", frm.adors!Pat_id
  Set rsNa = Nothing
  Call rsNa.Open("select * from namen where pat_id = " & frm.adoRS!Pat_id, DBCn, adOpenStatic, adLockReadOnly)
  If Not rsNa.EOF Then
   frm.adoRS!PatName = GesNam(rsNa) ' rNa!Nachname & " " & rNa!Vorname
  End If
  If Not IsNull(frm.adoRS!DokPfad) Then
'   raDok.Seek "=", CLng(frm.adors!Pat_id), CStr(frm.adors!DokPfad)
'   If Not raDok.NoMatch Then
   raDok.Open "select * from dokumente where pat_id = " & frm.adoRS!Pat_id & " and dokpfad = '" & frm.adoRS!DokPfad & "'", DBCn, adOpenDynamic, adLockReadOnly
   If Not raDok.BOF Then
    frm.adoRS!Eintragsdatum = raDok!Zeitpunkt
   End If ' radok.NoMatch Then
   Call makeDatPfad(frm)
   frm.Controls![gescannte Datei].Caption = "&Zeig: " + frm.adoRS!DokName
  Else ' IsNull(frm.adors!DokPfad) Then
   DatPfad = ""
   frm.Controls![gescannte Datei].Caption = ""
  End If ' IsNull(frm.adors!DokPfad) Then
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Fremdlabor_Form_Current/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function makeDatPfad(frm As Form)
 On Error GoTo fehler
 If (InStrB("$\", Left(frm!DokPfad, 1)) <> 0 And Left(frm!DokPfad, 1) <> "") Or Mid$(frm!DokPfad, 2, 1) = ":" Then
   RDatPfad = Replace(Replace(LCase(IIf(IsNull(frm!DokPfad), "", frm!DokPfad)), "$\turbomed\dokumente", getDokPfad), "^", "")
   DatPfad = "" + RDatPfad + ""
   Dim cat As New ADOX.Catalog
   Dim Cmd As New ADODB.Command
   cat.ActiveConnection = DBCn
   Set Cmd = cat.Views("LaborDokumente ep").Command
   If InStrB(Cmd.CommandText, "from [" + QMdbAkt + "].Briefe") <> 0 Then
'   If instrb(Dtb.QueryDefs![LaborDokumente ep].sql, "from [" + QMdbAkt + "].Briefe") <> 0 Then
    Dim Appl$, Bef$
    Appl = getReg(0, ".doc", "")
    Bef = getReg(0, Appl + "\shell\open\command", "")
'    DatPfadPI = Bef & " " & DatPfad '"c:\programme\microsoft office\office10\winword.exe " + DatPfad
    DatPfadPI = "winword " + DatPfad
    Bef = Replace(Replace(getReg(0, "wordpad.document.1\shell\open\command", ""), "%1", ""), "%programfiles%", Environ("ProgramFiles"))
    DatPfad = Bef & " " & DatPfad '"c:\programme\windows nt\zubehör\wordpad.exe " + DatPfad
    frm.Anzeigen.Caption = "Word&Pad"
    frm.Photoimpact.Caption = "&Winword"
   Else
    DatPfadPI = Environ("ProgramFiles") & "\Ulead Systems\Ulead PhotoImpact 6\iedit.exe " + DatPfad
    DatPfad = Environ("ProgramFiles") & "\irfanview\i_view32.exe " + DatPfad
   End If
  frm!Anzeigen.Enabled = -1
  frm!Photoimpact.Enabled = -1
  frm!DokName.Enabled = 0
 Else
  DatPfad = ""
  DatPfadPI = ""
  frm!Anzeigen.Enabled = 0
  frm!Photoimpact.Enabled = 0
  frm!DokName.Enabled = -1
 End If
 
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in makeDatPfad/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' makeDatPfad
' 4.9.06: wird nirgends aufgerufen
Public Function do_LaborDokumente_Form_Current(frm As Form)
 Static Fs, F
 On Error GoTo fehler
 Call makeDatPfad(frm)
'frm.Controls!Nachname
' rDokab.Seek "=", frm!DokPfad
' If Not rDokab.NoMatch Then
'   frm!abgehakt = rDokab!abgehakt
'End If
' frm!DokGroe = CreateObject("Scripting.FileSystemObject").getfile(Replace(Replace(IIf(IsNull(frm!DokPfad), "", frm!DokPfad), "$\TurboMed\Dokumente\", getDokPfad), "^", "")).size
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_LaborDokumente_Form_Current/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function do_LaborDokumente_form_load(frm As Form)
 On Error GoTo fehler
' Set raDokab = TabÖff("Dokumente abgehakt", "DokPfad")
 Set raDokab = Nothing
 Call raDokab.Open("select -ob as j_ob, d.* from `Dokumente abgehakt` d", DBCn, adOpenDynamic, adLockReadOnly)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_LaborDokumente_form_load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

Public Sub Taste(keyCode%, Shift%, frm As Form)
On Error GoTo fehler
    If keyCode = 40 Then
        DoCmd.GoToRecord , , acNext
    ElseIf keyCode = 38 Then
        DoCmd.GoToRecord , , acPrevious
    ElseIf keyCode = 27 Then
        DoCmd.Close acForm, frm.Name, acSaveYes
    End If
Ende:
    Exit Sub
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Taste/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub

Public Function do_Fremdlabor_Form_Load(frm As Form)
' On Error GoTo fehler
' Set rDok = TabÖff("Dokumente", "PIDokPfad")
' Set rsNa = TabÖff("Namen", "pat_ID")
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Fremdlabor_Form_Load/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): End
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
End Function
Public Function do_Fremdlabor_Form_Unload(frm As Form, Cancel%)
' On Error GoTo fehler
' rDok.Close
' rsNa.Close
' rDokab.Close
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Fremdlabor_Form_Unload/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): End
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
End Function

Public Function do_gescannte_Datei_Click(frm As Form)
 On Error GoTo fehler
 Call Shell(DatPfad, vbMaximizedFocus)
 frm!datum.SetFocus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_gescannte_Datei_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function do_anzeigen_click(frm As Form)
 On Error GoTo fehler
 If InStrB(DatPfad, ".exe") <> 0 Then
  Call Shell(DatPfad, vbMaximizedFocus)
 Else
  If Left(DatPfad, InStr(DatPfad, " ")) = "winword" Then
   GetWord
   Wapp.Visible = True
   Wapp.WindowState = wdWindowStateMaximize
   Wapp.Documents.Open DatPfad
   Wapp.Activate
  Else
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_anzeigen_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function do_PhotoImpact_Click(frm As Form)
 On Error GoTo fehler
 If InStrB(DatPfadPI, ".exe") <> 0 Then
  Call Shell(DatPfadPI, vbMaximizedFocus)
 Else
  If Left(DatPfadPI, InStr(DatPfadPI, " ") - 1) = "winword" Then
   GetWord
   Wapp.Visible = True
   Wapp.WindowState = wdWindowStateMaximize
   Wapp.Documents.Open Mid$(DatPfadPI, InStrB(DatPfadPI, " "))
   Wapp.Activate
  Else
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_PhotoImpact_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

Public Function do_FDatenquelle_Change(frm As Form)
Dim Pat_id&
 On Error GoTo fehler
Pat_id = frm.Recordset!Pat_id
On Error Resume Next ' "nicht zulässsig" kam einmal
frm.RecordSource = frm!FDatenquelle
On Error GoTo fehler
frm.Recordset.FindFirst "Pat_id = " + CStr(Pat_id)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_FDatenquelle_Change/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_FDatenquelle_Change(frm As Form)

'Function ListFreigaben()
'Dim Nr&, erg$, Inhalt$
'Nr = 0
'FZ = 0
'Do
'' erg = ReadRegistryGetValues(&H80000002, "SYSTEM\ControlSet001\Services\lanmanserver\Shares", nr, Inhalt)
' erg = ReadRegistryGetValues(&H80000002, "SYSTEM\CurrentControlSet\Services\lanmanserver\Shares", Nr, Inhalt)
' If Inhalt <> "" Then
'  If mid$(Inhalt, 2, 1) = ":" Then
'   ReDim Preserve FNam(FZ)
'   ReDim Preserve FInh(FZ)
'   FNam(FZ) = erg
'   FInh(FZ) = Inhalt
'   FZ = FZ + 1
'   Debug.Print erg
'   Debug.Print Inhalt
'  End If
' End If
' Nr = Nr + 1
'Loop Until trim$(erg) = ""
'End Function
'Public Function ReadRegistryGetValues$(ByVal Group&, ByVal Section$, Idx&, Inhalt$)
'Dim lResult&, lKeyValue&, lValueLength&, sValue$, td#, res&, vTyp&
'Dim ValueN$, cbValueN&
'ValueN = Space$(512)
'cbValueN = Len(ValueN)
'sValue = Space$(512)
'lValueLength = Len(sValue)
'Inhalt = ""
'sValue = ""
'On Error Resume Next
'lResult = RegOpenKeyEx(Group, Section, 0, &HF003F, lKeyValue)
'If lResult = 0 Then
' lResult = RegEnumValue(lKeyValue, Idx, ValueN, cbValueN, 0&, vTyp, sValue, lValueLength)
' If (lResult = 0) And (eRR.Number = 0) Then
'   sValue = left(sValue, InStr(sValue, vbnullchar) - 1)
'   ValueN = left(ValueN, InStr(ValueN, vbnullchar) - 1)
'   ReadRegistryGetValues = ValueN
'   If (instrb(sValue, "Path") <> 0) Then
'    Inhalt = mid$(sValue, InStr(sValue, "Path") + 5)
'    Inhalt = left(Inhalt, InStr(Inhalt, "Permiss") - 2)
'   End If
'   lResult = RegCloseKey(lKeyValue)
' End If
'
'End If
'End Function ' ReadRegistryGetSubkey$(ByVal Group&, ByVal Section$, Idx&)
Public Function ReadRegistryGetSubkey$(ByVal Group&, ByVal Section$, idx&)
 Dim lResult&, lKeyValue&, lDataTypeValue&, lValueLength&, sValue$, td#
 On Error GoTo fehler
 sValue = Space$(2048)
 lValueLength = Len(sValue)
 On Error Resume Next
 lResult = RegOpenKeyEx(Group, Section, 0, &HF003F, lKeyValue)
 If lResult = 0 Then
  lResult = RegEnumKey(lKeyValue, idx, sValue, lValueLength)
  If (lResult = 0) And (Err.Number = 0) Then
   sValue = Left(sValue, InStr(sValue, vbNullChar) - 1)
  End If
  ReadRegistryGetSubkey = sValue
  lResult = RegCloseKey(lKeyValue)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ReadRegistryGetSubkey/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ReadRegistryGetSubkey$(ByVal Group&, ByVal Section$, Idx&)


Sub GetExcel()
    Dim XL1 As Excel.Application 'Object    ' Variable für Verweis auf
                            ' Microsoft Excel.
    Dim ExcelLiefNicht As Boolean    ' Attribut für Freigabe am Ende.

' Überprüfen, ob eine Kopie von Microsoft Excel bereits
' ausgeführt wird.
On Error Resume Next    ' Fehlerbehandlung zurückstellen.
' GetObject-Funktionsaufruf ohne erstes Argument gibt einen Verweis auf
' eine Instanz der Anwendung zurück. Wenn die Anwendung nicht
' ausgeführt wird, tritt ein Fehler auf.
    Set XL1 = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelLiefNicht = True
    Err.Clear    ' Err-Objekt im Fehlerfall löschen.

' Prüfen auf Microsoft Excel. Wenn Microsoft Excel ausgeführt wird, wird
' dies in die Tabelle ausgeführter Objekte eingetragen.
    DetectExcel

' Objektvariable so festlegen, daß sie auf die gewünschte Datei verweist.
    Set XL1 = GetObject("c:\vb4\TEST1.XLS")

' Microsoft Excel mit zugehöriger Application-Eigenschaft einblenden.
' Fenster mit der Datei unter Verwendung der Windows-Auflistung des
' XL1-Objektverweises anzeigen.
    XL1.Application.Visible = True
    XL1.Parent.Windows(1).Visible = True
' Dateiverarbeitung.
    ' ...
' Wenn diese Kopie von Microsoft Excel beim Starten des Beispiels
' nicht ausgeführt wurde, wird Excel mit der Quit-Methode des
' Application-Objekts beendet. Wenn Sie versuchen, Microsoft Excel zu
' beenden, blinkt die Titelleiste, und Sie werden
' in einer Meldung gefragt, ob Sie geladene Dateien speichern möchten.
    If ExcelLiefNicht = True Then
        XL1.Application.Quit
    End If

    Set XL1 = Nothing    ' Verweis auf Anwendung und
                            ' Tabelle freigeben.
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetExcel/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' GetExcel

Sub DetectExcel()
' Diese Prozedur erkennt, ob Excel ausgeführt wird, und registriert dies.
    Const WM_USER = 1024
    Dim hWnd&
' Wenn Excel ausgeführt wird, wird durch Ausführen dieses
' API-Aufrufs die zugehörige Zugriffsnummer zurückgegeben.
    On Error GoTo fehler
    hWnd = FindWindow("XLMAIN", 0)
    If hWnd = 0 Then    ' 0 bedeutet, daß Excel nicht ausgeführt wird.
        Exit Sub
    Else
    ' Excel wird ausgeführt. Verwenden der API-Funktion
    ' SendMessage, um Excel in der Tabelle ausgeführter
    ' Objekte einzutragen.
        SendMessage hWnd, WM_USER + 18, 0, 0
    End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DetectExcel/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DetectExcel


Public Sub Detectword(className$)
' Procedure detects a running Word and r
' egisters it.
 Const WM_USER = 1024
 Dim hWnd&
' If Word is running this API call retur
' ns its handle.
 On Error GoTo fehler
 hWnd = FindWindow(className, vbNullString)

 If hWnd = 0 Then ' 0 means Word Not running.
 Exit Sub
 Else
' Word is running so use the SendMessage
' API function to enter it in the Running
' Object Table.
 SendMessage hWnd, WM_USER + 18, 0, 0
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Detectword/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Detectword


Private Sub cmdPreview_Click(myfilepath$)
'Preview a word document
On Error GoTo fehler
Screen.MousePointer = 11 'vbHourglass
'Startup Word if not started, or switch to existing one
GetWord
'Open the document, maximised, and switch to word
Wapp.Visible = True
Wapp.Application.WindowState = wdWindowStateMaximize
Wapp.Documents.Open (myfilepath)
Screen.MousePointer = 0 ' vbDefault
Wapp.Application.Activate
Set Wapp = Nothing
'We can also use the 'WordWasNotRunning' flag to decide if we should quit word or not
'Users might get annoyed if docs they are currently editing aren't saved!!
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdPreview_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' cmdPreview_Click
#If mitab Then
Public Function do_DMPAusgeb(frm As Form)
 Call do_DMPAusgebStandAlone(frm.adoRS!Pat_id)
End Function
#End If
#If False Then
Public Function do_DMPAusgeb(frm As Form)
 Dim dc, VorString$, DT As dmptyp
 Dim mR1, mR2, mR3
 Dim rAna As dao.Recordset
 Const sverz$ = "p:\"
 On Error GoTo fehler
 Set rAna = TabÖff("Anamnesebogen", "Pat_id")
 rAna.Seek "=", frm!Pat_id
 VorString = ""
 nzw = vbCr
 If rAna!DMPhier <> CDate(0) And frm.vTextB(169) = "Hausarzt im DMP" Then ' HAimDMP
  VorString = "Bei mir fand am " + Format$(rAna!DMPhier, "dd/mm/yyyy") + " die Einschreibung ins DMP statt. Diese sollte jedoch auf Sie als Hausarzt übertragen werden." + nzw
 End If
 GetWord
 With Wapp
  .Options.SmartCutPaste = False
  On Error Resume Next
  .Options.SmartParaSelection = False
  On Error GoTo fehler
  .Documents.Add
  Set dc = .ActiveDocument
  dc.Range.InsertAfter VorString & DMPString$(frm!Pat_id, DT)
       Set mR1 = dc.content
       With mR1.Find
         .clearformatting
         .Text = "DMP-Informationen zu"
         .Replacement.Text = ""
         .wrap = wdFindContinue
         .Format = False
         .Execute
       End With
       If mR1.Find.found Then
        Set mR2 = dc.Range(mR1.Start, mR1.Start)
        mR2.Find.Text = ":"
        mR2.Find.Execute
        If mR2.Find.found Then
         Set mR3 = dc.Range(mR1.Start, mR2.End)
         mR3.Font.Bold = True
        On Error Resume Next
        Dim para
        Set para = mR3.Paragraphs.First.Range
        Do While Err.Number = 0
         Set para = para.Paragraphs.First.Next.Range
         dc.Range(para.Start, para.Start + InStr(para.Text, ":")).Font.Italic = True
        Loop
        End If
       End If
  On Error GoTo fehler
  dc.Range.ParagraphFormat.TabStops.ClearAll
  dc.Range.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5.5), _
      Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderDots
  dc.Range.ParagraphFormat.FirstLineIndent = -35
  dc.Range.ParagraphFormat.LeftIndent = 35
  dc.SaveAs FileName:=sverz + frm.NachName & " " & frm.Vorname + ", DMP-Daten vom " + Format$(Now, "DD/MM/YY hh.mm.ss") + ".doc"
  .Visible = True
  
  .Application.WindowState = wdWindowStateMaximize
  .Activate
 End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_DMPAusgeb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_DMPAusgeb
#End If
#If False Then ' wurde in ZielDBFunktionen kopiert
Public Function do_DMPAusgebStandAlone(Pat_id&)
 Dim dc, VorString$, docName$, DT As dmptyp
 Dim mR1, mR2, mR3
' Dim raNa As dao.Recordset
 Dim rAna As ADODB.Recordset
 Const sverz$ = "p:\"
 On Error GoTo fehler
' Set raNa = TabÖff("Anamnesebogen", "Pat_id")
' raNa.Seek "=", Pat_id
 rAna.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 VorString = ""
 nzw = vbCr
' If rAna!DMPhier <> CDate(0) And rAna!HAimDMP = "Hausarzt im DMP" Then
'  VorString = "Bei mir fand am " + format$(rAna!DMPhier, "dd/mm/yyyy") + " die Einschreibung ins DMP statt. Diese sollte jedoch auf Sie als Hausarzt übertragen werden." + nzw
' End If
 GetWord
 With Wapp
  .Options.SmartCutPaste = False
  On Error Resume Next
  .Options.SmartParaSelection = False
  On Error GoTo fehler
  Call .Documents.Add("u:\vorlagen\DMP-Vorlage.dot")
  Set dc = .ActiveDocument
  dc.Range.InsertAfter VorString + DMPString$(Pat_id, DT)
       Set mR1 = dc.content
       With mR1.Find
         .clearformatting
         .Text = "DMP-Informationen zu"
         .Replacement.Text = ""
         .wrap = wdFindContinue
         .Format = False
         .Execute
       End With
       If mR1.Find.found Then
        Set mR2 = dc.Range(mR1.Start, mR1.Start)
        mR2.Find.Text = ":"
        mR2.Find.Execute
        If mR2.Find.found Then
         Set mR3 = dc.Range(mR1.Start, mR2.End)
         mR3.Font.Bold = True
        On Error Resume Next
        Dim para
        Set para = mR3.Paragraphs.First.Range
        Do While Err.Number = 0
         Set para = para.Paragraphs.First.Next.Range
         dc.Range(para.Start, para.Start + InStr(para.Text, ":")).Font.Italic = True
        Loop
        End If
       End If
  On Error GoTo fehler
  
  dc.sections(2).Range.ParagraphFormat.TabStops.ClearAll
  dc.sections(2).Range.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5.5), _
      Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderDots
  dc.sections(2).Range.ParagraphFormat.FirstLineIndent = -35
  dc.sections(2).Range.ParagraphFormat.LeftIndent = 35
  docName = sverz + rAna!NachName & " " & rAna!Vorname + ", DMP-Daten vom " + Format$(Now, "DD/MM/YY hh.mm.ss") + ".doc"
  dc.SaveAs FileName:=docName
  dc.Close
'  .Visible = True
  
'  .Application.WindowState = wdWindowStateMaximize
'  .Activate
 End With
 do_DMPAusgebStandAlone = docName
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_DMPAusgebStandAlone/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_DMPAusgebStandalone
#End If
Sub FaxSend(docName$, RecName$, RecNum$)
Dim FaxServer As New FAXCOMLib.FaxServer
Dim FaxDoc As New FAXCOMLib.FaxDoc
Dim FaxTiff As New FAXCOMLib.FaxTiff
Dim strFaxJob As FAXCOMLib.FaxJobs
Dim strFaxStatus As FAXCOMLib.FaxJob
Dim DateiName$
Dim strFaxTiff As FAXCOMLib.FaxTiff
Dim strJobID&
On Error GoTo fehler
 DateiName = "u:\test1.txt"
 Err.Clear
 FaxServer.Connect ("anmeld2")

Set FaxDoc = FaxServer.CreateDocument(DateiName)
   
    FaxDoc.BillingCode = "Rechnungsnummer 381"
    FaxDoc.CoverpageName = ""
    FaxDoc.CoverpageNote = "Insulinanaloga"
    FaxDoc.CoverpageSubject = "Insulinanaloga"
    FaxDoc.DiscountSend = 0
    FaxDoc.DisplayName = "G.Schade"
    FaxDoc.EmailAddress = "diabetologie@dachau-mail.de"
    FaxDoc.FaxNumber = RecNum
    FaxDoc.RecipientAddress = ""
    FaxDoc.RecipientCity = ""
    FaxDoc.RecipientCompany = "Praxis"
    FaxDoc.RecipientCountry = "D"
    FaxDoc.RecipientDepartment = ""
    FaxDoc.RecipientHomePhone = ""
    FaxDoc.RecipientName = RecName
    FaxDoc.RecipientOffice = ""
    FaxDoc.RecipientOfficePhone = ""
    FaxDoc.RecipientState = "Bayern"
    FaxDoc.RecipientTitle = ""
    FaxDoc.RecipientZip = ""
    FaxDoc.SendCoverpage = 0
    FaxDoc.SenderAddress = "Mittermayerstraße 13"
    FaxDoc.SenderCompany = "Diabetologische Schwerpunktpraxis"
    FaxDoc.SenderDepartment = "Schreibbüro"
    FaxDoc.SenderFax = "08131 616381"
    FaxDoc.SenderHomePhone = "616380"
    FaxDoc.SenderName = "Gerald Schade"
    FaxDoc.SenderOffice = "Praxis"
    FaxDoc.SenderOfficePhone = "616380"
    FaxDoc.SenderTitle = ""
    FaxDoc.ServerCoverpage = 1
    FaxDoc.FileName = docName
    FaxDoc.DisplayName = docName
    strJobID = FaxDoc.Send
'    MsgBox FaxServer.ArchiveDirectory
  
Set strFaxJob = FaxServer.GetJobs()
Set strFaxStatus = strFaxJob.Item(1)
    
On Error Resume Next

Set FaxServer = Nothing
Set FaxDoc = Nothing
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FaxSend/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' FaxSend

Function tha12()
 Call Lese.ProgStart
 Dim h$()
 Call getHausarzt(1576, h)
 Call Lese.ProgEnde
End Function

Function getHausarzt(pid&, Infos$(), Optional obHAimDMP%)
' Dim rNa As New ADODB.Recordset
 Dim rHa As New ADODB.Recordset
    Dim gefunden%, runde%, irunde%, HACngef%
    Dim rListena As New ADODB.Recordset
    Dim rHae As New ADODB.Recordset
    Dim faxfehlt%
' Dim rFa As New ADODB.Recordset
' Dim i%, j%
 Dim InfRoh$()
 Dim Üw12Erg$() ' HaErg$()
 On Error GoTo fehler
 ReDim Infos(12, 2)
 '  0: Frau/Herrn
 '  1: Titel+Vorn+Nachn,
 '  2: Straße,
 '  3: PLZ+Ort,
 '  4: Faxnr,
 '  5: S.g./Liebe,
 '  6: DMPTyp2,
 '  7: DMPTyp1,
 '  8: Niederlassungsgebiet 3. Feld für einmal austauschen,
 '  9: Vorname,
 ' 10: Funktion ("Üw 207, HA"),
 ' 11: Fachrichtung
 ' 12: KV-Nummer
 
 Call acon(HaT)

' in Faelle steht der als Überweiser eingetragene
 If Üw12(pid, Üw12Erg) = "" Then Exit Function
    
    ReDim InfRoh(12, UBound(Üw12Erg, 2))
    ReDim Infos(12, UBound(Üw12Erg, 2))
    For runde = 0 To UBound(Üw12Erg, 2)
     If Üw12Erg(0, runde) <> "" Then
     
      gefunden = 0
      If Üw12Erg(1, runde) <> "" And Üw12Erg(2, runde) <> "" Then
       Set rListena = Nothing
       rListena.Open "select * from listenausgabeuew where kvnr = '" & Üw12Erg(0, runde) & "' and name = '" & Üw12Erg(1, runde) & "' and vorname = '" & Üw12Erg(2, runde) & "' order by fax desc", DBCn, adOpenStatic, adLockReadOnly
       If Not rListena.BOF Then gefunden = True
      End If
      If Not gefunden And Üw12Erg(1, runde) <> "" Then
       Set rListena = Nothing
       rListena.Open "select * from listenausgabeuew where kvnr = '" & Üw12Erg(0, runde) & "' and name = '" & Üw12Erg(1, runde) & "' order by fax desc", DBCn, adOpenStatic, adLockReadOnly
       If Not rListena.BOF Then gefunden = True
      End If
      If Not gefunden Then
       Set rListena = Nothing
       rListena.Open "select * from listenausgabeuew where kvnr = '" & Üw12Erg(0, runde) & "' order by fax desc", DBCn, adOpenStatic, adLockReadOnly
       If Not rListena.BOF Then gefunden = True
      End If
      If Not gefunden Then
       Set rListena = Nothing
       rListena.Open "select * from listenausgabeuew where name = '" & Üw12Erg(1, runde) & "' and vorname = '" & Üw12Erg(2, runde) & "' order by fax desc", DBCn, adOpenStatic, adLockReadOnly
       If Not rListena.BOF Then gefunden = True
      End If
      If gefunden Then
       InfRoh(1, runde) = IIf(IsNull(rListena!Titel) Or rListena!Titel = "", "Dr.med.", rListena!Titel) & " " & rListena!Vorname & " " & rListena!Name
       InfRoh(2, runde) = rListena!strasse
       InfRoh(3, runde) = rListena!Plz & " " & rListena!Ort
       InfRoh(4, runde) = rListena!Fax
       InfRoh(8, runde) = rListena!fachgruppe
       InfRoh(9, runde) = rListena!Vorname
       InfRoh(10, runde) = Üw12Erg(3, runde)
       InfRoh(12, runde) = rListena!KVNr
      End If
      
      gefunden = 0
      HACngef = 0
      For irunde = 0 To 1 ' aktuelle oder alte Hausärzte
       If Not gefunden And Üw12Erg(1, runde) <> "" And Üw12Erg(2, runde) <> "" Then
        Set rHa = Nothing
        rHa.Open "select  -dmpt1 as j_dmpt1, -dmpt2 as j_dmpt2, h.* from hae" & IIf(irunde = 1, "alt", "") & " h where kvnr = '" & Left(Üw12Erg(0, runde), 2) & "/" & Right$(Üw12Erg(0, runde), 5) & "' and nachname = '" & Üw12Erg(1, runde) & "' and vorname = '" & Üw12Erg(2, runde) & "'", HAECn, adOpenStatic, adLockReadOnly
        If Not rHa.BOF Then gefunden = True
       End If
       If Not gefunden And Üw12Erg(1, runde) <> "" Then
        Set rHa = Nothing
        rHa.Open "select -dmpt1 as j_dmpt1, -dmpt2 as j_dmpt2, h.* from hae" & IIf(irunde = 1, "alt", "") & " h where kvnr = '" & Left(Üw12Erg(0, runde), 2) & "/" & Right$(Üw12Erg(0, runde), 5) & "' and nachname = '" & Üw12Erg(1, runde) & "'", HAECn, adOpenStatic, adLockReadOnly
        If Not rHa.BOF Then gefunden = True
       End If
       If Not gefunden Then
        Set rHa = Nothing
        rHa.Open "select -dmpt1 as j_dmpt1, -dmpt2 as j_dmpt2, h.* from hae" & IIf(irunde = 1, "alt", "") & " h where kvnr = '" & Left(Üw12Erg(0, runde), 2) & "/" & Right$(Üw12Erg(0, runde), 5) & "'", HAECn, adOpenStatic, adLockReadOnly
        If Not rHa.BOF Then gefunden = True
       End If
      Next irunde
      If gefunden Then
       InfRoh(0, runde) = rHa!Anrede
       InfRoh(1, runde) = IIf(IsNull(rHa!Titel) Or rHa!Titel = "", "Dr.med.", rHa!Titel) & " " & rHa!Vorname & " " & rHa!NachName
       InfRoh(2, runde) = rHa!Straße
       InfRoh(3, runde) = rHa!Plz & " " & rHa!Ort
       If InfRoh(4, runde) = "" And Not IsNull(rHa!fax1) Then InfRoh(4, runde) = rHa!fax1
       InfRoh(6, runde) = IIf(rHa!j_dmpt2 <> 0, "X", "") ' geht auch für null
       InfRoh(7, runde) = IIf(rHa!j_dmpt1 <> 0, "X", "")
       Dim rDMP As ADODB.Recordset
       Set rDMP = Nothing
       Set rDMP = New ADODB.Recordset
       rDMP.Open "select * from hae where kvnr = '" & Left(Üw12Erg(0, runde), 2) & "/" & Right$(Üw12Erg(0, runde), 5) & "' and dmpt2 = 1", HAECn, adOpenStatic, adLockReadOnly
       InfRoh(6, runde) = IIf(rDMP.BOF, "", "X")
       Set rDMP = Nothing
       Set rDMP = New ADODB.Recordset
       rDMP.Open "select * from hae where kvnr = '" & Left(Üw12Erg(0, runde), 2) & "/" & Right$(Üw12Erg(0, runde), 5) & "' and dmpt1 = 1", HAECn, adOpenStatic, adLockReadOnly
       InfRoh(7, runde) = IIf(rDMP.BOF, "", "X")
       If InfRoh(8, runde) = "" Then InfRoh(8, runde) = rHa!zulg
       If InfRoh(9, runde) = "" Then InfRoh(9, runde) = rHa!Vorname
       InfRoh(10, runde) = Üw12Erg(3, runde)
       InfRoh(12, runde) = rHa!kvnu
       HACngef = -1
      End If
      
      gefunden = 0
      If Üw12Erg(1, runde) <> "" And Üw12Erg(2, runde) <> "" Then
       Set rHae = Nothing
       rHae.Open "select * from hausaerzte where kvnr = '" & Üw12Erg(0, runde) & "' and nachname = '" & Üw12Erg(1, runde) & "' and vorname = '" & Üw12Erg(2, runde) & "'", DBCn, adOpenStatic, adLockReadOnly
       If Not rHae.BOF Then gefunden = True
      End If
      If Not gefunden And Üw12Erg(1, runde) <> "" Then
       Set rHae = Nothing
       rHae.Open "select * from hausaerzte where kvnr = '" & Üw12Erg(0, runde) & "' and nachname = '" & Üw12Erg(1, runde) & "'", DBCn, adOpenStatic, adLockReadOnly
       If Not rHae.BOF Then gefunden = True
      End If
      If Not gefunden Then
       Set rHae = Nothing
       rHae.Open "select * from hausaerzte where kvnr = '" & Üw12Erg(0, runde) & "'", DBCn, adOpenStatic, adLockReadOnly
       If Not rHae.BOF Then gefunden = True
      End If
      If gefunden Then
       If InfRoh(0, runde) = "" Then InfRoh(0, runde) = IIf(rHae!Geschlecht = "w", "Frau", "Herr")
       If InfRoh(1, runde) = "" Then InfRoh(1, runde) = IIf(IsNull(rHae!Titel) Or rHae!Titel = "", "Dr.med.", rHae!Titel) & " " & rHae!Vorname & " " & rHae!NachName
       If InfRoh(2, runde) = "" Then InfRoh(2, runde) = rHae!Straße
       If InfRoh(3, runde) = "" Then InfRoh(3, runde) = rHae!Plz & " " & rHae!Ort
       If InfRoh(9, runde) = "" Then InfRoh(9, runde) = rHae!Vorname
       If InfRoh(4, runde) = "" And Not IsNull(rHae!Telefax) Then InfRoh(4, runde) = rHae!Telefax
        Select Case rHae!Überschrift
         Case "L": InfRoh(5, runde) = IIf(InfRoh(0, runde) = "Frau", "Liebe", "Lieber") & " " & InfRoh(9, runde) ' nicht: rHae!Geschlecht = "w"
         Case "H": InfRoh(5, runde) = "Hallo " + InfRoh(9, runde)
         Case Else: InfRoh(5, runde) = IIf(InfRoh(0, runde) = "Frau", "Sehr geehrte Frau Kollegin", "Sehr geehrter Herr Kollege")
        End Select
'       infroh(6, runde) = IIf(rHae!dmpt2 = -1, "X", "")
'       infroh(7, runde) = IIf(rHae!dmpt1 = -1, "X", "")
       If InfRoh(8, runde) = "" Then InfRoh(8, runde) = rHae!Zulassungsgebiet
       InfRoh(10, runde) = Üw12Erg(3, runde)
       InfRoh(12, runde) = rHae!KVNr
      End If
      Dim obInsert%
      If Not gefunden Then
       obInsert = True
      Else
       obInsert = False
       If Not rHae.BOF Then
        If rHae!NachName <> Üw12Erg(1, runde) Or rHae!Vorname <> Üw12Erg(2, runde) Then obInsert = True
       End If
      End If
      If obInsert And Not rHae.BOF Then
       If HACngef Then
        Set rHae = Nothing
'        Call DBCn.Execute("set foreign_key_checks = 0")
'        Call DBCn.Execute("delete from hausaerzte where id >= 753")
'        Call DBCn.Execute("set foreign_key_checks = 1")
        Call rHae.Open("insert into hausaerzte (name, vorname, nachname, anschrift, kvnr, telefon, telefax, e_mail, zulassungsgebiet, arzttyp, " & _
        kla & "gemeinschaftspraxis mit" & klz & ", beme, geschlecht,titel,straße,ort,plz,überschrift,dmpt2,dmpt1,zahl,nichtmehr,schwerpunkt,zusatzbezeichnung,bemerkung,sprechstunden) values('" & rHa!Anrede & " " & IIf(rHa!Titel <> "" And Not IsNull(rHa!Titel), rHa!Titel, "Dr.med.") & " " & rHa!Vorname & " " & rHa!NachName & "','" & rHa!Vorname & "','" & rHa!NachName & "','" & rHa!Straße & ", " & rHa!Plz & " " & rHa!Ort & "','" & rHa!kvnu & "','" & rHa!tel1 & "','" & rHa!fax1 & "','" & rHa!Email & "','" & rHa!zulg & "','" & rHa!Arzttyp & "','" & rHa!gemmit & "','" & rHa!Beme & "','" & IIf(rHa!Anrede = "w", "Frau", "Herr") & "','" & rHa!Titel & "','" & rHa!Straße & "','" & rHa!Ort & "','" & rHa!Plz & "',''," & IIf(InfRoh(6, runde) = "X", "1", "0") & "," & IIf(InfRoh(7, runde) = "X", "1", "0") & ",0,0,'','','','')", DBCn, adOpenStatic, adLockOptimistic)
       End If
      End If
     End If ' Üw12Erg(0,runde)<> ""
     If InfRoh(5, runde) = "" Then InfRoh(5, runde) = IIf(InfRoh(0, runde) = "Frau", "Sehr geehrte Frau Kollegin", IIf(InfRoh(0, runde) = "Herr", "Sehr geehrter Herr Kollege", "Sehr geehrte Frau Kollegin, sehr geehrter Herr Kollege"))
     InfRoh(11, runde) = IIf(InStrB(LCase(InfRoh(8, runde)), "intern") <> 0 Or InStrB(LCase(InfRoh(8, runde)), "allg") <> 0, "x", "")
    Next runde
    Dim pos&, i%, aktrunde%, gelöscht%
    pos = UBound(InfRoh, 2)
' zuletzt kommen die Überweiser, die nicht als Allgemeinarzt/Internist sind und nicht als Hausarzt eingetragen
    For runde = UBound(InfRoh, 2) To 0 Step -1
     If InfRoh(11, runde) = "" And InStrB(InfRoh(10, runde), "HA") = 0 Then
      For i = 0 To UBound(InfRoh, 1)
       Infos(i, pos) = InfRoh(i, runde)
'       InfRoh(i, runde) = ""
      Next i
      pos = pos - 1
     End If
    Next runde
    aktrunde = UBound(InfRoh, 2) + 1 ' der letzte hier berücksichtigte
' dann kommen bis zur 3. Stelle die Überweiser, die als Allgemeinarzt/Internist sind aber nicht als Hausarzt eingetragen
    For runde = UBound(InfRoh, 2) To 0 Step -1
     If pos <= 1 Then Exit For
     If InfRoh(11, runde) = "x" And InStrB(InfRoh(10, runde), "HA") = 0 Then
      gelöscht = 0
      For irunde = UBound(InfRoh, 2) To 0 Step -1
       If irunde <> runde Then
        If InfRoh(1, irunde) = InfRoh(1, runde) Then ' doppelter Arzt
         For i = 0 To UBound(InfRoh, 1)
          InfRoh(i, runde) = ""
         Next i
         gelöscht = True
         Exit For
        End If
       End If
      Next irunde
      If Not gelöscht Then ' war ja oben so, also falsch nicht gelöscht
       For i = 0 To UBound(InfRoh, 1)
        Infos(i, pos) = InfRoh(i, runde)
        InfRoh(i, runde) = "" ' schon eingetragen
       Next i
      End If
      pos = pos - 1
     End If
     aktrunde = runde
    Next runde
' dann kommt der Hausarzt, falls bisher nicht berücksichtigt
    For runde = UBound(InfRoh, 2) To aktrunde Step -1
     If InStrB(InfRoh(10, runde), "HA") <> 0 Then
      gelöscht = 0
      For irunde = UBound(InfRoh, 2) To 0 Step -1
       If irunde <> runde Then
        If InfRoh(1, irunde) = InfRoh(1, runde) Then ' doppelter Arzt
         For i = 0 To UBound(InfRoh, 1)
          InfRoh(i, runde) = ""
         Next i
         gelöscht = True
         Exit For
        End If
       End If
      Next irunde
      If Not gelöscht Then ' war ja oben so, also falsch nicht gelöscht
       For i = 0 To UBound(InfRoh, 1)
        Infos(i, pos) = InfRoh(i, runde)
        InfRoh(i, runde) = ""
       Next i
      End If
      pos = pos - 1
      Exit For
     End If
    Next runde
' dann der Rest in der Reihenfolge
    For runde = aktrunde - 1 To 0 Step -1
     If InfRoh(11, runde) = "x" Or InStrB(InfRoh(10, runde), "HA") <> 0 Then
      gelöscht = 0
      For irunde = UBound(InfRoh, 2) To 0 Step -1
       If irunde <> runde Then
        If InfRoh(1, irunde) = InfRoh(1, runde) Then ' doppelter Arzt
         For i = 0 To UBound(InfRoh, 1)
          InfRoh(i, runde) = ""
         Next i
         gelöscht = True
         Exit For
        End If
       End If
      Next irunde
      If Not gelöscht Then  ' war ja oben so, also falsch nicht gelöscht
       For i = 0 To UBound(InfRoh, 1)
        Infos(i, pos) = InfRoh(i, runde)
        InfRoh(i, runde) = ""
       Next i
      End If
      pos = pos - 1
     End If
    Next runde
' Wenn erste Position leer, schaut's unschön aus
    If Infos(1, 0) = "" Then
     If UBound(Infos, 2) > 0 Then
      If Infos(1, 1) <> "" Then
       For i = 0 To UBound(Infos, 1)
        Infos(i, 0) = Infos(i, 1)
        Infos(i, 1) = ""
       Next i
      Else
       If UBound(Infos, 2) > 1 Then
        If Infos(1, 2) <> "" Then
         For i = 0 To UBound(Infos, 1)
          Infos(i, 0) = Infos(i, 1)
          Infos(i, 1) = ""
         Next i
        End If
       End If
      End If
     End If
    End If
    Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getHausarzt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getHausarzt
Function doVorhandene&()
 Dim Fil As File, pid&, pos&, p2&, BriefZiel$
 Dim VorMüll$, FilName$, FilPath$
 On Error GoTo fehler
 BriefZiel = InputBox("Verzeichnis:", "wo sollen vorhandene Briefe korrigiert werden?", AutoBriefZiel)
 If BriefZiel <> "" Then
  For Each Fil In FSO.GetFolder(BriefZiel).Files
   FilPath = Fil.path
   pos = InStr(FilPath, "PID")
   Debug.Print FilPath
   If pos > 0 Then
    p2 = InStr(pos, FilPath, ",")
    If p2 > 0 Then
     pid = CLng(Mid$(FilPath, pos + 4, p2 - pos - 4))
     If pid > 0 Then
      VorMüll = FilPath & "zulöschen"
      FilName = Fil.Name
      Name Fil.path As VorMüll
      Forms(0).Ausgabe = "Erneuere: " & FilName & vbCrLf & altAusgabe
      altAusgabe = Forms(0).Ausgabe
      Call tubriefStandalone(pid, True, BriefZiel)
      If BrichAb Then Exit For
      Kill VorMüll
      doVorhandene = doVorhandene + 1
nächstedatei:
     End If
    End If
   End If
  Next Fil
 End If
 Exit Function
fehler:
  Dim FNr&, FLastDLLError&, FSource$, FDescr$
  FNr = Err.Number
  FLastDLLError = Err.LastDllError
  FSource = Err.source
  FDescr = Err.Description
  Call meld("Achtung Fehler beim Erneuern von: " & FilPath, 0)
  PiepKurz
  If FNr = 75 Then Resume nächstedatei
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), "", CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in doVorhandene/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doVorhandene
Function doDuplexkontrollieren()
 Dim rs As New ADODB.Recordset, begD As Date, sql$
 begD = CDate(InputBox("Ab welchem Datum?"))
 sql = "select pat_id, zeitpunkt, min(dokname) as DokName, quelldatum, count(*) as zahl from dokumente where dokname like '%SonoBild%' and quelldatum >= " & datform(begD) & " group by pat_id, zeitpunkt,quelldatum order by quelldatum"
 rs.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 Do While Not rs.EOF
  Forms(0).Ausgabe = Right$(Space$(5) & rs!Pat_id, 5) & " " & rs!Zeitpunkt & " " & Right$(Space$(3) & rs!Zahl, 3) & " " & rs!DokName & " " & rs!Quelldatum & vbCrLf & altAusgabe
  altAusgabe = Forms(0).Ausgabe
  rs.Move 1
 Loop
End Function ' doDuplexkontrollieren

Function doBriefeBerichtspflicht()
 Dim QDat$, DSZahl&, angefangen%, erg&, ergS$, FDT As Date
 On Error Resume Next
' QDat = getLDatei("u:\", "Listenausgabe_EBM-Ziffern*.xls")
' If Err.Number <> 0 Then
'  MsgBox "Fehler beim Einladen der Datei " & "u:\" & "Listenausgabe_EBM-Ziffern*.xls"
'  Exit Function
' End If
 QDat = "u:\Listenausgabe_Listen.xls"
 ergS = Dir(QDat)
 If ergS <> "" Then
  FDT = FileDateTime(QDat)
  erg = MsgBox("Die Datei " & QDat & " stammt von: " & FDT & "." & vbCrLf & "Wollen Sie zu allen dort enthaltenen Patienten Briefe schreiben und abrechnen?", vbOKCancel, "Rückfrage")
  If erg = vbNo Then Exit Function
 Else
  MsgBox ("Datei " & QDat & " nicht gefunden, breche ab!")
  Exit Function
 End If
 On Error GoTo fehler
' aus Function EmailsImport(EmDatei$)
 Dim con As New ADODB.Connection  ' Connection
 Dim rNa As New ADODB.Recordset
 Dim rEx As New ADODB.Recordset
 Dim rX As New ADOX.Catalog
' Const EmDatei$ = "p:\Patientenübergreifendes\Emails.xls" ' Excel-Datei mit Suche aus Turbomed "*@*"
 On Error GoTo fehler
 If DBCn Is Nothing Or DBCn = "" Then
   Dim dlg As New Dialog
'   Call frm.ConstrFestleg(2)
   Call acon(quelleT, qDtb)
 End If
 If QDat <> "" Then
 con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;HDR=No;IMEX=1"";Data Source=" & QDat & ";" ' TABLE=Adressen$"
 Dim runde%, i%, zFeld$, lFeld$, obAnfang%, pNr%, pRoh
  rX.ActiveConnection = con
  rEx.Open "[" & rX.Tables(rX.Tables.Count - 1).Name & "]", con ' Hier Excel, nicht lies.obmysql = 0!
  Do While Not rEx.EOF
'  Debug.Print runde
   If obAnfang Then
'    On Error Resume Next
     pRoh = rEx.Fields(0)
     If InStrB(rEx.Fields(2), "Es ist kein Bericht vorhanden") > 0 Then
' Hier jetzt Vorname, Name und Geburtsdatum extrahieren, um den Patienten zu finden!
       Dim vn$, nn$, gd$, gesn$, leid$
       Dim p1%, p2%, p3%
       p1 = InStr(pRoh, " ")
       p2 = InStr(pRoh, "(")
       p3 = InStr(pRoh, " - ")
       vn = Left(pRoh, p1 - 1)
       gesn = Replace(Left(pRoh, p2 - 2), "zzz", "")
       nn = Mid$(pRoh, p1 + 1, InStr(p1 + 1, pRoh, " ") - p1 - 1)
       gd = Mid$(pRoh, p2 + 1, InStr(pRoh, ")") - p2 - 1)
       leid = Mid$(pRoh, p3 + 3, InStr(pRoh, ";") - 3 - p3)
       DSZahl = DSZahl + 1
       If nn Like "zzz*" Then nn = Mid$(nn, 4)
       Set rNa = Nothing
       sql = "select * from namen where concat(titel¡iif(titel='','',' ')¡nvorsatz¡iif(nvorsatz='','',' ')¡vorname¡' '¡nachname) = '" & gesn & "' and gebdat = " & datform(gd) & ""
       If Forms(0).obMySQL Then
        sql = Replace(Replace(sql, "¡", ","), "iif(", "if(")
       Else
        sql = Replace(Replace(sql, "concat", ""), "¡", " & ")
       End If
       Call rNa.Open("select count(*) from (" & sql & ") as innen", DBCn, adOpenStatic, adLockOptimistic)
       
       If rNa.Fields(0) <> 1 Then
        MsgBox "Falsche Zahl an Datensätzen für " & vn & " " & nn & ", geb. " & gd & ": " & rNa.Fields(0)
       Else
        Set rNa = Nothing
        Call rNa.Open(sql, DBCn, adOpenStatic, adLockOptimistic)
        If Not angefangen Then
          Call LeistungsExport0
          angefangen = True
        End If
        Call LeistungsExport1(rNa!Pat_id, "01601", CDate(leid), CDate("18:00"))
        Call LeistungsExport1(rNa!Pat_id, "40120", CDate(leid), CDate("18:00"))
        Call tubriefStandalone(rNa!Pat_id, True)
       End If
     End If
   ElseIf Not IsNull(rEx.Fields(0)) And Not IsNull(rEx.Fields(1)) And Not IsNull(rEx.Fields(2)) Then
     obAnfang = True
   End If
  runde = runde + 1
  rEx.MoveNext
  Loop
 End If
 On Error Resume Next
 rEx.Close
 Forms(0).Ausgabe = "Fertig mit Nachtragen aus '" & QDat & "' von " & DSZahl & " Datensätzen " & vbCrLf + altAusgabe
 altAusgabe = Forms(0).Ausgabe
 If angefangen Then
  Close #310
  MsgBox "Datei " & LeistBDT & " neu mit den Leistungen zu den Briefen in " & QDat & " erstellt!"
 End If
 Exit Function
fehler:
  Dim FNr&, FLastDLLError&, FSource$, FDescr$
  FNr = Err.Number
  FLastDLLError = Err.LastDllError
  FSource = Err.source
  FDescr = Err.Description
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), "", CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in doVorhandene/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doBriefeBerichtspflicht
'lfaelle: SELECT *
'FROM faelle INNER JOIN [SELECT min(pat_id) as pid, max(bhfb) as bb FROM faelle GROUP BY pat_id]. AS sel ON ([faelle].[bhfb]=sel.bb) AND ([faelle].[pat_id]=sel.pid)
'ORDER BY [pat_id];
' Vorlage: getDokPfad("Vorlagen") + "\AccessBrief.dot" ' \\linux\turbomed\vorlagen\AccessBrief.dot (s.u.)
#If mitab Then
Public Sub tu_brief(frm As Form)
 Call tubriefStandalone(frm.adoRS!Pat_id, 0)
End Sub
#End If
Public Sub tubriefStandalone(pid&, obstumm%, Optional Zielverz$)
  Dim myRange, Docu, Inh, dc As Object ' Word.Document, wegen unbekannter Wordversion als Object
  Dim raHa As New ADODB.Recordset
  Dim raAn As New ADODB.Recordset
  Dim raHae As New ADODB.Recordset
  Dim rsNa As New ADODB.Recordset
  raAn.Open "select * from anamnesebogen where pat_id = " & pid, DBCn, adOpenStatic, adLockReadOnly
  Dim obHAimDMP%
  Dim j&
' dim fs
  Dim runde%, KVNr$
  Dim vordat As Date
     Dim ÜwNr$, ÜWeingef As Boolean
  Dim pos%
     ÜwNr = ""
     ÜWeingef = False
  Dim sverz$
  On Error GoTo fehler
'  sverz = "p:\"
  If Zielverz = "" Then
   sverz = BriefZiel
   If obstumm Then
'    sverz$ = sverz & "unkorrigiert\"
     sverz = AutoBriefZiel
   End If
  Else
   sverz = Zielverz
  End If
'  lies.obmysql = False
'  Call Zinit(Lies.obMySQL)
  nzw = vbCr
  
'  Set rsNa = TabÖff("Namen", "pat_ID")
'  rsNa.Seek "=", raAn!Pat_id
  rsNa.Open "select * from namen where pat_id = " & pid, DBCn, adOpenStatic, adLockReadOnly
  If rsNa.EOF Then Exit Sub
  syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 1) Hausarzt ..."
  Dim Faxnr$, Infos$() ' Frau/Herrn, Vorn+Nachn, Straße, PLZ+Ort, Faxnr, S.g./Liebe, DMPTyp2, DMPTyp1
  Call getHausarzt(raAn!Pat_id, Infos())
  
'  Set rHa = TabÖff("Hausaerzte", "KVNR")
  
'  On Error Resume Next
  Call GetWord
  With Wapp
   .Options.SmartCutPaste = False
   If WappBuild > 9 Then '.options("SmartParaSelection") = 0 '.Options.SmartParaSelection = False
    With .Options
     .SmartParaSelection = 0
    End With
   End If
'   .Documents.Add Template:="c:\turbomed\vorlagen\AccessBriefa.dot" ' Replace(Replace(An1Pfad, "$\TurboMed\Dokumente\", PCDokPfad), "^", "")
     Dim tpl$
    Set dc = Nothing
    While dc Is Nothing
     tpl = getDokPfad("Vorlagen") + "\AccessBrief.dot" ' \\linux\turbomed\vorlagen\AccessBrief.dot
     On Error Resume Next
     Set dc = .Documents.Add(Template:=tpl)
     On Error GoTo fehler
     If dc Is Nothing Then
      MsgBox "Fehler beim Anlegen eines Winword-Dokumentes auf der Grundlage von " + tpl + ", welches aus den Turbomed Grundeinstellungen gelesen wurde!"
     End If
    Wend
    .ScreenUpdating = False
    .Visible = False
    With dc.Bookmarks
'    Set dc = .ActiveDocument
    vordat = GetVorDat(raAn!Pat_id, obstumm)
    
    Call BMRd(dc, !ha1a, Infos(0, 0))
    Call BMRd(dc, !ha2a, Infos(1, 0))
    Call BMRd(dc, !ha3a, Infos(2, 0))
    Call BMRd(dc, !ha4a, Infos(3, 0))
    Call BMRd(dc, !ha5a, Infos(4, 0)) ' Faxnummer
    Call BMRd(dc, !ha6a, Infos(5, 0)) ' Anrede
    Call BMRd(dc, !ha10a, Infos(10, 0)) ' Überweiserinfos
    
    If UBound(Infos, 2) > 0 Then
     Call BMRd(dc, !ha1b, Infos(0, 1)) '  IIf(raHacall bmrd(dc,!Geschlecht = "w", "Frau", "Herrn")) '+ vblf) + nzw
     Call BMRd(dc, !ha2b, Infos(1, 1)) ' "Dr.med. " + raHacall bmrd(dc,!VorName & " " & raHacall bmrd(dc,!NachName
     Call BMRd(dc, !ha3b, Infos(2, 1)) ' raHacall bmrd(dc,!Straße
     Call BMRd(dc, !ha4b, Infos(3, 1)) ' raHacall bmrd(dc,!Plz & " " & raHacall bmrd(dc,!Ort
     Call BMRd(dc, !ha5b, Infos(4, 1)) ' Faxnummer
     Call BMRd(dc, !ha6b, Infos(5, 1)) ' Anrede
     Call BMRd(dc, !ha10b, Infos(10, 1)) ' Überweiserinfos
    End If
'weiter:
'    Next i
    
    !PVorn.Range = IIf(IsNull(rsNa!Titel), "", rsNa!Titel + " ") + rsNa!Vorname
    !pnachn.Range = rsNa!NachName
    !PGeb.Range = Format$(rsNa!GebDat, "d.M.YYYY")
    
'    Dim obAnredeSpez%
'    obAnredeSpez = 0
'    If raHa.State = 1 Then
'     If Not (raHa.BOF Or raHa.EOF) Then
'      obAnredeSpez = -1
'     End If
'    End If
'    If obAnredeSpez Then
'     Select Case raHa!Überschrift
'      Case "L":  !haanr.Range = IIf(raHa!Geschlecht = "w", "Liebe", "Lieber") & " " & raHa!VorName
'      Case "H":  !haanr.Range = "Hallo " + raHa!VorName
'      Case Else: !haanr.Range = IIf(raHa!Geschlecht = "w", "Sehr geehrte Frau Kollegin", "Sehr geehrter Herr Kollege")
'     End Select
'    Else
''    If !haanr.Range = "$1313$" Then
'     !haanr.Range = "Sehr geehrte Frau Kollegin, sehr geehrter Herr Kollege"
'    End If
    
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 2)a) Diagnosen prüfen ..."
    obkNeph = obKeineNephropathie(raAn!Pat_id, obMakroAlb)
    Dim raFa As New ADODB.Recordset
    Dim lddat As Date
    raFa.Open "select max(bhfb) as lddat, max(fanf) as fanf, max(ausgst) as ausgst from faelle where pat_id = " & raAn!Pat_id, DBCn, adOpenStatic, adLockReadOnly
    If Not raFa.BOF Then
     lddat = raFa!lddat
     If raFa!Fanf > lddat Then lddat = raFa!Fanf
     If raFa!ausgst > lddat Then lddat = raFa!ausgst
    End If
    Dim iDiag As New ADODB.Recordset, licd$, ldiag$
    iDiag.Open "Select * from diagnosen where pat_id = " & CStr(raAn!Pat_id) & " and (obdauer or diagdatum > " & datform(lddat) & ") order by icd", DBCn, adOpenDynamic, adLockReadOnly
    If Not iDiag.BOF Then
     iDiag.MoveFirst
     Do While Not iDiag.EOF
       Dim tonRunde%
'       MsgBox "Achtung doppelte Diagnose: " + nzw + _
              ldiag & " " & licd + nzw + _
              iDiag!DiagText & " " & iDiag!ICD
'      Call Shell("cmd /c echo ""Achtung doppelte Diagnose"" ", vbNormalNoFocus)
      If iDiag!ICD = licd Then
       For tonRunde = 1 To 5
        Call Sound(WinDir + "\media\Windows XP-Standard.wav")
       Next tonRunde
       Call meld("Achtung doppelte Diagnose: " & vbCrLf & ldiag & " " & licd & vbCrLf & iDiag!DiagText & " " & iDiag!ICD, obstumm)
      End If
      licd = iDiag!ICD
      ldiag = iDiag!DiagText
      iDiag.Move 1
     Loop
    End If
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 2)b) Diagnosen einfügen ..."
    !Diagnosen.Range = DiagString$(raAn!Pat_id, vordat, True)
    !pdieder.Range = IIf(rsNa!Geschlecht = "w", "die", "der")
    !zeitraum.Range = BehDauerStr(raAn!Pat_id)
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 3) Anamnese ..."
    !Anamnese.Range = machWertString$(raAn!Pat_id) '+ nzw + vblf
'    !Verlauf.Range = Verlauf(raan!Pat_id)
'End If
    Dim Ereih%, TMn$, TZMn$, TMnErg$, Krit$, Zerleg$(), Zerl2$(), Inhalt$
    Dim rlAb As New ADODB.Recordset, Komm$, lauf%, KSpl$()
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 4) Weitere Daten ..."
    rlAb.Open "SELECT distinct quelle,Eingang,kommentar from laborxus INNER join laborxbakt ON laborxus.RefNr = laborxbakt.RefNr WHERE pat_id = " & CStr(raAn!Pat_id) & " and not isnull(keimzahl) and eingang > " & datform(vordat), DBCn, adOpenDynamic, adLockReadOnly
    Const bmUrin% = 7
    Const bmFuß% = 9
    Const bmEKG% = 13
    Const bmTaille = 14
    For Ereih = 1 To 14
     TMn = "e" + CStr(Ereih)
     TZMn = "ez" + CStr(Ereih)
     Krit = ""
     On Error Resume Next
     Zerleg = Null
     Zerleg = Split(dc.Bookmarks(TMn).Range, "[")
     On Error GoTo fehler
     If Not IsNull(Zerleg) Then
      For j = 0 To UBound(Zerleg)
       If InStrB(Zerleg(j), "]") > 0 Then
        Zerl2 = Split(Zerleg(j), "]")
        Krit = Krit + """" + Zerl2(0) + """" + ","
       End If
      Next j
      Krit = Left(Krit, Len(Krit) - 1) '+ """"
      Inhalt = eintraege(raAn!Pat_id, Krit, vordat)
     End If
     Select Case Ereih
      Case bmTaille
       Dim rTail As New ADODB.Recordset
       sql = "select zeitpunkt as zp, inhalt from eintraege WHERE pat_id = " + CStr(raAn!Pat_id) + " and trim(art) = ""taille"" order by zeitpunkt"
       Set rTail = Nothing
       rTail.Open sql, DBCn, adOpenDynamic, adLockReadOnly
       If rTail.BOF Then
       Else
        Inhalt = ""
        Do While Not rTail.EOF
         Inhalt = Inhalt + CStr(DateValue(rTail!Zp)) + vbTab + rTail!Inhalt + nzw
         rTail.Move 1
        Loop
       End If
      Case bmUrin ' Urinuntersuchung
       Inhalt = Replace(Replace(Inhalt, "hts:", "Harnteststreifen:"), "hts.:", "Harnteststreifen:")
       If Not rlAb.BOF Then
        Do While Not rlAb.EOF
         KSpl = Split(rlAb!Kommentar, nzw)
         For lauf = 0 To UBound(KSpl)
          If InStrB(KSpl(lauf), ":") > 0 And InStrB(KSpl(lauf), "  ") > 0 Then
           KSpl(lauf) = Replace(KSpl(lauf), " ", "")
          End If
          If Mid$(KSpl(lauf), 2, 4) = "Keim" Then If Right$(KSpl(lauf), 1) <> ":" Then KSpl(lauf) = KSpl(lauf) + ":"
         Next lauf
         For lauf = 0 To UBound(KSpl) - 1
          If Right$(KSpl(lauf), 1) = ":" Then
            KSpl(lauf) = KSpl(lauf) + IIf(Asc(Left(KSpl(lauf + 1), 1)) < 32, Mid$(KSpl(lauf + 1), 2), KSpl(lauf + 1))
            KSpl(lauf + 1) = ""
          End If
         Next lauf
         Komm = ""
         For lauf = 0 To UBound(KSpl)
          KSpl(lauf) = Replace(KSpl(lauf), ":", ":" + vbTab)
          If KSpl(lauf) <> "" Then
             Komm = Komm + nzw + KSpl(lauf)
          End If
         Next lauf
         If Komm <> "" Then
          Inhalt = Inhalt + nzw + rlAb!Quelle + " (" & rlAb!Eingang & "):" + vbTab + Komm + nzw
         End If
         rlAb.Move 1
        Loop
       End If ' Not rLab.BOF Then
      Case bmFuß
       Inhalt = Replace(Inhalt, nzw, ", ")
       sql = "select zeitpunkt as zp, inhalt from eintraege WHERE pat_id = " + CStr(raAn!Pat_id) + " and trim(art) = ""usdm"" and zeitpunkt > " & datform(vordat) & " order by zeitpunkt"
'       Set rTail = Dtb.OpenRecordset(sql)
       rTail.Open sql, DBCn, adOpenDynamic, adLockReadOnly
       If rTail.BOF Then
       Else
        Inhalt = ""
        Do While Not rTail.EOF
         Dim InhVorl$
         pos = InStr(rTail!Inhalt, "aktuellen Blutdruck")
         If pos > 0 Then InhVorl = Left(rTail!Inhalt, pos - 1) Else InhVorl = rTail!Inhalt
         Inhalt = Inhalt + CStr(DateValue(rTail!Zp)) + vbTab + InhVorl + nzw
         rTail.Move 1
        Loop
        Inhalt = Left(Inhalt, Len(Inhalt) - 1)
       End If
      Case bmEKG
'       sql = "SELECT FormInhaltNeu.ZeitPunkt, FormInhaltFeld.Feld, FormInhaltFeldInh.FeldInh, FormInhaltNeu.AbsPos " + _
             "FROM FormInhaltForm_Abk INNER JOIN (FormInhaltFeldInh INNER JOIN (FormInhaltFeld INNER JOIN FormInhaltNeu ON FormInhaltFeld.FeldVW = FormInhaltNeu.FeldVW) ON FormInhaltFeldInh.FeldInhVW = FormInhaltNeu.FeldInhVW) ON FormInhaltForm_Abk.Form_AbkVW = FormInhaltNeu.Form_AbkVW " + _
             "WHERE (Feld =""Zeile"" or isnull(Feld)) AND FeldInh Not Like ""-*"" AND FeldInh Not Like ""Ruhe-EKG Messwerte*"" AND Form_Abk=""GDT"" and pat_id = " & CStr(Pat_id)
'       sql = " SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, FeldVW, FeldInhVW FROM ((forminhfeld left join forminhkopf on forminhfeld.foid = forminhkopf.foid) left join formulare on formulare.formid = forminhkopf.form_id) left join forminhaltform_abk on forminhaltform_abk.form_abkvw = formulare.form_abkvw WHERE (Feld =""Zeile"" or isnull(Feld)) AND FeldInh Not Like ""-*"" AND FeldInh Not Like ""Ruhe-EKG Messwerte*"" AND Form_Abk=""GDT"" and pat_id = " & CStr(Pat_id)
       Dim rEkg As New ADODB.Recordset
       sql = "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh FROM (((forminhfeld left join forminhkopf on forminhfeld.foid = forminhkopf.foid) left join formulare on formulare.formid = forminhkopf.form_id) left join forminhaltfeld on forminhfeld.feldvw = forminhaltfeld.feldvw) left join forminhaltfeldinh on forminhfeld.feldinhvw = forminhaltfeldinh.feldinhvw WHERE (Feld =""Zeile"" or isnull(Feld)) AND FeldInh Not Like ""-*"" AND FeldInh Not Like ""Ruhe-EKG Messwerte*"" AND Form_Abk=""GDT"" and pat_id = " & CStr(Pat_id) & " and Zeitpunkt > " & datform(vordat)
       rEkg.Open sql, DBCn, adOpenDynamic, adLockReadOnly
       Dim altzeitp As Date, FeldInh$, nfi$, aktZ%
       Inhalt = "": runde = 0
       Do While Not rEkg.EOF
        aktZ = aktZ + 1
        FeldInh = Replace(Replace(Replace(Replace(rEkg!FeldInh, "|", ","), " ¦", " °"), " ,", ","), ",,", ",")
        Do
         nfi = Replace(FeldInh, "  ", " ")
         If nfi = FeldInh Then
          Exit Do
         Else
          FeldInh = nfi
         End If
        Loop
        FeldInh = Left(FeldInh, Len(FeldInh) - IIf(Right$(FeldInh, 1) = ",", 1, 0))
        Select Case aktZ
         Case 2, 4
          Inhalt = Inhalt & ", " & FeldInh
         Case Else
          Inhalt = Inhalt + IIf(Inhalt = "", "", nzw) + IIf(rEkg!Zeitpunkt <> altzeitp, Format$(rEkg!Zeitpunkt, "dd/mm/yy"), "") + vbTab + FeldInh
        End Select
        altzeitp = rEkg!Zeitpunkt
        rEkg.Move 1
       Loop
      End Select
     If Inhalt = "" Then
'      If instrb(dc.Bookmarks(TMn).Range, "$2004") > 0 Then
       On Error Resume Next
       dc.Range(dc.Bookmarks(TMn).Range.Start, dc.Bookmarks(TMn).Range.End + 1) = Inhalt
       dc.Bookmarks(TZMn).Range = ""
       On Error GoTo fehler
'      End If
     Else
      On Error GoTo fehler
'      If instrb(dc.Bookmarks(TMn).Range, "$2004") > 0 Then
       dc.Bookmarks(TMn).Range = Inhalt
'      End If
     End If
    Next Ereih
    End With
    If Not rlAb.BOF Then
     Dim aktr As Range
     Dim aktRa
     Set aktRa = dc.Bookmarks("Urin").Range
'     With aktra.Find
'     .ClearFormatting
'     .text = nzw + nzw '"^p^p"
'     .Execute
'     End With
'     If aktra.Find.Found Then
'      Set aktra = dc.Range(aktra.Start + 2, dc.Bookmarks("Urin").End)
     pos = InStr(aktRa, nzw + nzw)
     If pos > 0 Then
      Set aktRa = dc.Range(aktRa.Start + pos + 1, dc.Bookmarks("Urin").End)
      aktRa.ParagraphFormat.TabStops.ClearAll
      aktRa.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(4)
      aktRa.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
      aktRa.ParagraphFormat.LeftIndent = 0
     End If
    End If
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 5) Labor..."
    Call LaborIns(dc, raAn!Pat_id)
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 6) Medikation..."
    Call letzteMed(dc, raAn!Pat_id)
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 7) Epikrise..."
    Call Epikrise(dc, raAn!Pat_id, vordat, lddat, obstumm)
#If False Then
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 8) DMP-Infos..."
    Dim rAna As ADODB.Recordset
'    Set raNa = New ADODB.Recordset
'    Set raNa = TabÖff("Anamnesebogen", "Pat_id")
'    raNa.Seek "=", raAn!Pat_id
'    Set raNa = Nothing
'    Call raNa.Open("select * from anamnesebogen where pat_id = " & raAn!Pat_id, DBCn, adOpenStatic, adLockReadOnly)
    Dim mR1, mR2, mR3
    Set raFa = Nothing
    Call raFa.Open("select * from faelle where pat_id = " & rsNa!Pat_id & " order by bhfb desc", DBCn, adOpenStatic, adLockReadOnly)
    Dim SchGr$
    If Not raFa.BOF Then
     SchGr = raFa!SchGr
    Else
     SchGr = ""
    End If
    If Not rsNa.BOF And obHAimDMP And SchGr <> "90" And (raAn!Diabetestyp = "2" Or raAn!Diabetestyp = "s") Then
      Dim VorString$
      VorString = ""
      If rsNa!DMPhier <> CDate(0) Then
       VorString = "Bei mir fand am " + Format$(rsNa!DMPhier, "dd/mm/yyyy") + " die Einschreibung ins DMP statt. Diese sollte jedoch auf Sie als Hausarzt übertragen werden." + nzw
      End If
'      If (rAna!Diabetestyp = "2" Or rAna!Diabetestyp = "s") Then ' Not raan!HAimDMP = "Hausarzt nicht im DMP")
       dc.Bookmarks("DMP").Range = vbFormFeed + VorString + DMPString$(raAn!Pat_id)
       Set mR1 = dc.content
       With mR1.Find
         .clearformatting
         .Text = "DMP-Informationen zu"
         .Replacement.Text = ""
         .wrap = wdFindContinue
         .Format = False
         .Execute
       End With
       If mR1.Find.found Then
        Set mR2 = dc.Range(mR1.Start, mR1.Start)
        mR2.Find.Text = ":"
        mR2.Find.Execute
        If mR2.Find.found Then
         Set mR3 = dc.Range(mR1.Start, mR2.End)
         mR3.Font.Bold = True
        On Error Resume Next
        Dim para
        Set para = mR3.Paragraphs.First.Range
        Do While Err.Number = 0
         Set para = para.Paragraphs.First.Next.Range
         dc.Range(para.Start, para.Start + InStr(para.Text, ":")).Font.Italic = True
        Loop
        End If
       End If
    Else
#End If
'      dc.bookmarks("DMP").Range = ""
      If Now - lddat > 90 Then
       If Now - lddat > 365 Then
        dc.Bookmarks("DMPText").Range = "Ich bedauere die beträchtliche Verspätung des Arztbriefes." & vbCrLf
       Else
        dc.Bookmarks("DMPText").Range = "Ich bedauere die Verspätung des Arztbriefes." & vbCrLf
       End If
      Else
       dc.Bookmarks("DMPText").Range = ""
      End If
'      dc.bookmarks("DMPText").Range = ""
#If False Then
    End If
#End If
    On Error GoTo fehler
' bei Wiederholungsbriefen
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 9) Prüfen von Vorberichten..."
    Call WiederHolungsBrief(rsNa!NachName, vordat, Wapp, dc)
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 10) Korrekturen..."
    Call dc.Range.Find.Execute("Markt^t", , , , , , , , , "Markt Indersdorf^t", wdReplaceAll)
    Call dc.Range.Find.Execute("IDA", True, , , , , , , , "Insulindosisanpassung", wdReplaceAll)
    Call dc.Range.Find.Execute("ID-Anpassung", True, , , , , , , , "Insulindosisanpassung", wdReplaceAll)
    Call dc.Range.Find.Execute("Hypo`s", True, , , , , , , , "Hypoglykämien", wdReplaceAll)
    Call dc.Range.Find.Execute("GPD", True, , , , , , , , "Gesundheitspaß Diabetes", wdReplaceAll)
    Call dc.Range.Find.Execute("DFS", True, , , , , , , , "Diabetisches Fußsyndrom", wdReplaceAll)
    Call dc.Range.Find.Execute("bettunge,", True, , , , , , , , "bettungen,", wdReplaceAll)
    Call dc.Range.Find.Execute("bettunge.", True, , , , , , , , "bettungen.", wdReplaceAll)
    Call dc.Range.Find.Execute("$2032$", True, , , , , , , , "", wdReplaceAll)
    Dim suchr
    Set suchr = dc.Range
    Call suchr.Find.Execute("Verlaufsnotizen")
    dc.Range(suchr.Start, suchr.Start).Select
   If Not obstumm Then .Visible = True
   .Application.WindowState = wdWindowStateMaximize
   .ScreenUpdating = True
   If Not obstumm Then .Activate
    syscmd acSysCmdSetStatus, "Erstelle Brief für " + raAn!NachName & ", " & raAn!Vorname + ": 11) Speichern..."
    Call VerzPrüf(sverz)
    Call ZwischenSpeichern(dc, sverz, rsNa!NachName, rsNa!Vorname, rsNa!Pat_id, Infos(4, 0)) ' Faxnr
  End With ' wapp
  On Error Resume Next
  If obstumm Then
   Call Wapp.Quit(0)  '  = wdDoNotSaveChanges
  Else
   Call snie
  End If
  syscmd acSysCmdClearStatus
  Exit Sub
fehler:
  Dim FNr&, FLastDLLError&, FSource$, FDescr$
  FNr = Err.Number
  FLastDLLError = Err.LastDllError
  FSource = Err.source
  FDescr = Err.Description
  PiepKurz
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), "", CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in tubriefStandalone/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' tubriefStandalone(frm As Form)
' Ersetzt Bookmark mit Text, ohne es zu löschen
Function BMRd(dc As Object, Bm, Text$)
 Dim Start&, Ende&, NeuLen&, BMName$
 With Bm
  Start = .Start
  Ende = .End
 End With
 NeuLen = Len(Text)
 BMName = Bm.Name
 Bm.Range = Text
 Call dc.Bookmarks.Add(BMName, dc.Range(Start, Start + NeuLen))
 Exit Function
fehler:
  Dim FNr&, FLastDLLError&, FSource$, FDescr$
  FNr = Err.Number
  FLastDLLError = Err.LastDllError
  FSource = Err.source
  FDescr = Err.Description
  PiepKurz
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), "", CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in BMRd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' bmrd
Function fobHAimDMP%(HANr$)
 Dim rKVA As New ADODB.Recordset
 Dim obDMP2%, obDMP1%
 On Error GoTo fehler
' Call KVÄVorb
 Call acon(HaT)
 rKVA.Open "select -dmpt1 as j_dmpt1, -dmpt2 as j_dmpt2, hae.* from hae where kvnr = '" & Left(HANr, 2) & "/" & Right$(HANr, 5) & "'", HAECn, adOpenDynamic, adLockReadOnly
 If Not rKVA.BOF Then
  Do While Not rKVA.EOF
   If rKVA!j_dmpt2 <> 0 Then obDMP2 = True
   If rKVA!j_dmpt1 <> 0 Then obDMP1 = True
   rKVA.Move 1
  Loop
 End If
 fobHAimDMP = obDMP2
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in fobHAimDMP/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' fobHAimDMP(HANr$)
#If mitab Then
Sub HAlokal(rHa As ADODB.Recordset, KVNr$, Optional NaN, Optional VoN)
#If False Then
' in der örtlichen Datei nicht gefundenen Hausarzt aus der systemischen übertragen
  Dim rKv As New ADODB.Recordset
  On Error GoTo fehler
       Set rKv = TabÖff("HAE")
       If Not IsMissing(NaN) And Not IsMissing(VoN) Then
        Call rKv.Open("select * from hae where kvnr = '" & Left(KVNr, 2) & "/" & Right$(KVNr, 5) & "' and nachname = '" & NaN & "' and vorname = '" & VoN & "'", HAECn, adOpenStatic, adLockReadOnly)
       End If
       If rKv Is Nothing Or rKv.BOF Then
        Set rKv = Nothing
        Call rKv.Open("select * from hae where kvnr = '" & Left(KVNr, 2) & "/" & Right$(KVNr, 5) & "' and nachname = '" & NaN & "'", HAECn, adOpenStatic, adLockReadOnly)
       End If
       If rKv Is Nothing Or rKv.BOF Then
        Set rKv = Nothing
        Call rKv.Open("select * from hae where kvnr = '" & Left(KVNr, 2) & "/" & Right$(KVNr, 5) & "'", HAECn, adOpenStatic, adLockReadOnly)
       End If
       If Not rKv.BOF Then
        rHa.AddNew
        rHa!NachName = rKv!NachName
        rHa!Vorname = rKv!Vorname
        rHa!KVNr = Left(Replace(rKv!KVNr, "/", ""), 7)
        rHa!Anschrift = rKv!Straße & ", " & rKv!Plz & " " & rKv!Ort
        rHa!Plz = rKv!Plz
        rHa!Ort = rKv!Ort
        rHa!Straße = rKv!Straße
        rHa!Telefon = rKv!tel1
        rHa!Telefax = rKv!fax1
        rHa!E_Mail = rKv!Email
        rHa!Zulassungsgebiet = rKv!zulg
        rHa!Arzttyp = rKv!Arzttyp
        rHa![Gemeinschaftspraxis mit] = rKv!gemmit
        rHa!Bemerkung = rKv!Beme
        rHa!Geschlecht = IIf(rKv!Anrede = "Frau", "w", IIf(rKv!Anrede = "Herr", "m", ""))
        rHa.Update
'        rHa.Bookmark = rHa.LastModified
        Call do_haakt(rHa)
       End If
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HAlokal/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
#End If
End Sub
#End If
Function Üw12$(Pat_id&, Üw1$()) ' Saubere Funktion zum Ermitteln der Überweisernummern aus Faelle sowie des Hausarztes aus Namen
' zunächst für das Anamneseblatt, später auch für den Brief
  Const maxFZ% = 50 ' Maximalzahl der verschiedenen Fälle
  Dim rsNa As New ADODB.Recordset
  Dim raFa As New ADODB.Recordset
  Dim stand%, runde%, Feld$, i%, obalt%, bhb#(), obneueFallZeit%
  ReDim bhb(0)
  On Error GoTo fehler
  stand = -1
'  Üw2 = ""
    
  raFa.Open "select * from faelle where pat_id = " & Pat_id & " order by bhfb desc", DBCn, adOpenDynamic, adLockReadOnly
'  Set rFa = TabÖff("faelle", "AktF") ' pat_id aufsteigend, BHFB absteigend
'  rFa.Seek "=", Pat_id
  If Not raFa.EOF Then
' Fälle werden nach zeitlicher Relevanz durchsucht
' Wenn ein neuer Überweiser auftaucht, wird Üw1 um 1 auf (stand-1) erweitert und zuletzt der neue Überweiser eingetragen
' wenn ein alter Überweiser, der zuvor evtl. nur aus AndÜW bekannt war, jetzt aus Übwv bekannt wird, so werden Name und Vorname ergänzt
   Do
' Zahl der verschiedenen Fälle bestimmen, ggf Abbruch
    For runde = 0 To UBound(bhb)
     obneueFallZeit = True
     If bhb(runde) = 0 Then ' nur am Anfang
      obneueFallZeit = False
      bhb(runde) = raFa!BhFB
      Exit For
     ElseIf bhb(runde) = raFa!BhFB Then
      obneueFallZeit = False
      Exit For
     End If
    Next runde
    If obneueFallZeit Then
     ReDim Preserve bhb(UBound(bhb) + 1)
     bhb(UBound(bhb)) = raFa!BhFB
    End If
    If UBound(bhb) = maxFZ Then Exit Do
    
    For runde = 1 To 2
     Select Case runde
      Case 1: Feld = "AndÜw"
      Case 2: Feld = "Übwv"
     End Select
     If Not IsNull(raFa.Fields(Feld)) Then
      If raFa.Fields(Feld) <> "0000000" And raFa.Fields(Feld) <> "" Then
       obalt = 0
       For i = 0 To stand
        If Üw1(0, i) = raFa.Fields(Feld) Then
         If runde = 2 Then
          Üw1(1, i) = raFa!ÜWNaN
          Üw1(2, i) = raFa!ÜWVor
         End If
         obalt = True
         Exit For
        End If
       Next i
       If Not obalt Then
        stand = stand + 1
        ReDim Preserve Üw1(3, stand) ' 0 = KV-Nr, 1 = Nachname, 2 = Vorname, 3 = Position
        Üw1(0, stand) = raFa.Fields(Feld)
        Üw1(3, stand) = "Üw " & Left(raFa!Quartal, 1) & Right$(raFa!Quartal, 2)
        If runde = 2 Then
         Üw1(1, stand) = raFa!ÜWNaN
         Üw1(2, stand) = raFa!ÜWVor
        End If
       End If ' not obalt
      End If
     End If
    Next runde
'    If stand >= maxÜwZ - 1 Then Exit Do
    raFa.MoveNext
    If raFa.EOF Then Exit Do
'    If raFa!Pat_id <> Pat_id Then Exit Do
   Loop
  End If
  
'  Set rsNa = TabÖff("Namen", "pat_ID")
'  rsNa.Seek "=", Pat_id
  rsNa.Open "select * from namen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
  If Not rsNa.EOF Then
   If Not IsNull(rsNa!KVNr) Then
    If rsNa!KVNr <> "0000000" And rsNa!KVNr <> "" Then
     obalt = 0
     For i = 0 To stand
      If Üw1(0, stand) = rsNa!KVNr Then
       Üw1(3, stand) = Üw1(3, stand) & ", HA"
       obalt = True
       Exit For
      End If
     Next i
     If Not obalt Then
      stand = stand + 1
      ReDim Preserve Üw1(3, stand) ' 0 = KV-Nr, 1 = Nachname, 2 = Vorname, 3 = Position
      Üw1(0, stand) = rsNa!KVNr
      Üw1(3, stand) = "HA"
     End If
    End If
   End If
  End If
  
  If Not raFa.EOF Or rsNa!KVNr <> "" Then
   Üw12 = Üw1(0, 0)
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Üw12/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Üw12
Function Testüb5()
 Dim Üw()
 Stop
 ReDim Üw(0)
 Stop
 ReDim Üw(1)
End Function
Function testWied()
 Dim vordat As Date
 On Error GoTo fehler
 Call GetWord
 vordat = GetVorDat(Forms![Anamnesebogen]!Pat_id, False)
 Call WiederHolungsBrief(Forms![Anamnesebogen], vordat, Wapp)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in testWied/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' testWied()
Function GetVorDat(Pat_id&, obstumm%) As Date
  Const obsichtbar = -1 ' 0 ergibt das aktuelle Datum
  Dim BRz As New ADODB.Recordset
  Dim lBrNam$, WAlt As Object, US$, Spl$(), j%
  On Error GoTo fehler
  If Wapp Is Nothing Then GetWord
  If Not Wapp Is Nothing Then
'   Call dtbInit
'   sql = "select count(*) as ct from [" + QMdbAkt + "].[briefe relevant] where pat_id = " & CStr(Pat_id)
   sql = "SELECT count(*) as ct From briefe WHERE (name Like ""%Brief an %Dr%"" Or name Like ""%Arztbrief%"" Or name Like ""Brief an HA%"" Or name Like ""Brief an HAe%"") And name Not Like ""%Entwurf%"" and pat_id = " & Pat_id
   BRz.Open sql, DBCn, adOpenStatic, adLockReadOnly
'   If Dtb.OpenRecordset(sql)!ct > 0 Then
   If BRz!ct > 0 Then
'    lBrNam = Dtb.OpenRecordset("select pfad from [" + QMdbAkt + "].[briefe relevant] where pat_id = " & CStr(Pat_id))!Pfad
    sql = "SELECT pfad From briefe WHERE (name Like ""%Brief an %Dr%"" Or name Like ""%Arztbrief%"" Or name Like ""Brief an HA%"" Or name Like ""Brief an HAe%"") And name Not Like ""%Entwurf%"" and pat_id = " & Pat_id & " order by zeitpunkt desc"
    Set BRz = Nothing
    BRz.Open sql, DBCn, adOpenStatic, adLockReadOnly
    lBrNam = BRz!Pfad
    lBrNam = Replace(Replace(lBrNam, "$\TurboMed\Dokumente", getDokPfad), "^", "")
    Select Case Wapp.Version
     Case "7.0", "8.0", "9.0"
     On Error Resume Next
     Set WAlt = Wapp.Documents.Open(lBrNam, ReadOnly:=-1, Visible:=obsichtbar) ', -1, -1, -1, , , 0, , , , , obsichtbar)
     On Error GoTo fehler
     Case "10.0", "11.0", "12.0" ' ab XP
      On Error Resume Next
' Ausdruck.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenConflictDocument, OpenAndRepair , DocumentDirection, NoEncodingDialog)
      
      Set WAlt = Wapp.Documents.Open(lBrNam, 0, -1, -1, , , 0, , , , -1, obsichtbar, -1, -1)
      If Err.Number > 0 Then
       Call meld("Fehler beim Öffnen des Vorbefundes: " & lBrNam, obstumm)
      End If
      On Error GoTo fehler
    End Select
    If Not WAlt Is Nothing Then
     US$ = Replace(Replace(Replace(WAlt.Range, vbTab, Chr$(32)), vbCr, Chr$(32)), vbVerticalTab, Chr$(32))
     Spl = Split(US)
     For j = 1 To UBound(Spl)
      If IsDate(Spl(j)) Then
       GetVorDat = CDate(Spl(j))
       Exit For
      End If
     Next j
    End If
   End If ' not isempty(walt)
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetVorDat/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetVorDat(pat_id&, Optional WApp) As Date

Function WiederHolungsBrief(NachName$, vordat As Date, Wapp, Optional dc)
 Dim i%, lBriefDat As Date
 Dim r1, r2, Rn
 Dim dcName$
 
 On Error GoTo fehler
 If vordat = CDate(0) Then GoTo schluss
 
 With Wapp
  On Error Resume Next
  dcName = dc.Name
  If Err.Number <> 0 Then
   For i = 1 To .Windows.Count
    If InStrB(.Windows(i), NachName) > 0 Then
     Set dc = .Windows(i)
     GoTo habDC
    End If
   Next i
   GoTo schluss
  End If
  On Error GoTo fehler

habDC:
   On Error GoTo fehler
   Set r1 = dc.Range ' WApp.Documents(dc).Content
   Set r2 = dc.Range
   Dim r3, r0
   Set r3 = dc.Range
   With r3.Find
    .clearformatting
    .Execute ("Ich berichte")
   End With
   If r3.Find.found Then
    r3.InsertAfter (" erneut")
   End If
   With r1.Find
    .clearformatting
    .Execute ("Angaben auf dem Anamnese- und Untersuchungsbogen")
   End With
   If r1.Find.found Then
    Set r0 = dc.Range(r1.Start - 1, r1.Start - 1)
    r0.InsertAfter (vbCr + "Zur Vorgeschichte s. Bericht vom " + Format$(vordat, "dd/mm/yy") + ".")
   End If
   With r2.Find
    .clearformatting
    .Execute ("Verlaufsnotizen")
   End With
   If r1.Find.found And r2.Find.found Then
    Set Rn = dc.Range(r1.Start, r2.Start - 1)
    Rn.Font.Hidden = True
    Call dc.Bookmarks.Add("Anamnese", Rn)
   End If
   Dim pAkt ' Aktueller Absatz
   If r2.Find.found Then
    Set pAkt = r2.Paragraphs(1).Next
    Do While pAkt.Range.Text <> vbCr
     If Left(pAkt.Range, 1) <> vbTab And Not IsDate(Left(pAkt.Range, 8)) Then GoTo fertig
     If IsDate(Left(pAkt.Range, 8)) And IsNumeric(Left(pAkt.Range, 8)) Then ' Ergänzung isnumeric 27.4.08 Jörger
      If CDate(Left(pAkt.Range, 8)) > vordat Then
        GoTo fertig
      End If
     End If
     Set pAkt = pAkt.Next
    Loop
fertig:
    On Error Resume Next
    dc.Range(r2.Paragraphs(1).Next.Range.Start, pAkt.Range.Start - 1).Font.Hidden = True
    On Error GoTo fehler
   End If

schluss:
  End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WiederHolungsBrief/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' WiederHolungsBrief(frm As Form, WApp, Optional dc)
#If False Then
Public Function obKeineNephropathie%(Pat_id&, Optional obMakroAlb%)
 Dim lAlbS$, raLab As New ADODB.Recordset, lKreS
 On Error GoTo fehler
 obMakroAlb = 0
' Call dtbInit
 sql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich as NB " + _
  "FROM laborxus LEFT JOIN laborxwert ON laborxus.RefNr= laborxwert.RefNr " & _
  "WHERE pat_id = " + CStr(Pat_id) + " and not exists (select * from laborneu where pat_id = " + CStr(Pat_id) + " and abkü = laborxwert.Abkü and wert = laborxwert.wert and zeitpunkt > laborxus.Eingang-3 and zeitpunkt < laborxus.Eingang+6)"
' lAlbS = Dtb.OpenRecordset("select iif(isnull(wert),iif(isnull(kommentar),"""",kommentar),wert) as erg from (" + sql + ") as sql1 where abkü in (""ALBCRE"",""ALBKRE"",""ALBQ"") order by zeitpunkt desc")!erg
 sql1 = "select iif(isnull(wert),iif(isnull(kommentar),"""",kommentar),wert) as erg from (" + sql + ") as sql1 where abkü in (""ALBCRE"",""ALBKRE"",""ALBQ"") order by zeitpunkt desc"
 If lies.obMySQL Then sql1 = Replace(sql1, "iif(", "if(")
 raLab.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 If Not raLab.BOF Then
  lAlbS = raLab!erg
 End If
 If lAlbS = "" Then Exit Function
 If InStrB(lAlbS, "<") = 0 Then
  lAlbS = Replace(lAlbS, ".", ",")
  If Not IsNumeric(lAlbS) Then Exit Function
  If CDbl(lAlbS) >= 200 Then obMakroAlb = True
  If CDbl(lAlbS) >= 10 Then Exit Function ' setzen wir mal 10 als Grenze an
 End If
 On Error Resume Next
' lKreS = Dtb.OpenRecordset("select iif(isnull(wert),iif(isnull(kommentar),"""",kommentar),wert) as erg from (" + sql + ") as sql1 where abkü in (""KREA"",""KREA02"") order by zeitpunkt desc")!erg
 sql1 = "select iif(isnull(wert),iif(isnull(kommentar),"""",kommentar),wert) as erg from (" + sql + ") as sql1 where abkü in (""KREA"", ""KREA02"", ""KRE02"", ""CREAT"") order by zeitpunkt desc"
 If lies.obMySQL Then sql1 = Replace(sql1, "iif(", "if(")
 Set raLab = Nothing
 raLab.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 If Not raLab.BOF Then
  lKreS = raLab!erg
 End If
 If IsNumeric(lKreS) And lKreS <> 0 Then
  If CDbl(Replace(lKreS, ".", ",")) >= 1.3 Then
   Exit Function
  End If
  obKeineNephropathie = -1
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obKeineNephropathie/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' obKeineNephropathie%(Pat_id&, Optional obMakroAlb%)
#End If
Function ZwischenSpeichern(dc, sverz, NachName$, Vorname$, Pat_id&, Optional Fax$)
 Const ABv$ = "Arztbrief vom"
 Dim ABName$, FNr&
 On Error GoTo fehler
 ABName = NachName & " " & Vorname & ", PID " & Pat_id & ", " & ABv & " " & Format$(Now, "DD/MM/YY hhmmss")
 If Fax <> "" Then
  ABName = ABName & " an Fax " & Fax
 End If
 ABName = ABName & ".doc"
 If InStrB(Wapp.ActiveDocument.Name, ABv) > 0 Then
  Wapp.ActiveDocument.Save
 Else
  On Error GoTo fehler
  Wapp.ActiveDocument.SaveAs Replace(sverz + ABName, "/", "")
  If FNr <> 0 Then
   On Error GoTo fehler
   Wapp.ActiveDocument.SaveAs ABName
  End If
 End If
 Exit Function
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "beim Speichern unter: " + ABName + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZwischenSpeichern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZwischenSpeichern

Function letzteMed(dc, Pat_id&)
 Dim raDat As New ADODB.Recordset, sql$, ZZ&, aktZ&, runde%
 On Error GoTo fehler
 sql = "SELECT  max(zeitpunkt) AS zp FROM medplan WHERE pat_id = " + CStr(Pat_id)
 raDat.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 If IsNull(raDat!Zp) Then
  dc.Bookmarks("Therapiedat").Range = ""
  Exit Function
 End If
 lTherZP = raDat!Zp
 On Error GoTo fehler
 Dim Bemerkung
 dc.Bookmarks.Add Name:="MedBemerkung", Range:=dc.Range(dc.Bookmarks!Therapie.Start + 1, dc.Bookmarks!Therapie.Start + 1)
 dc.Bookmarks("Therapiedat").Range = Format$(lTherZP, "dd.mm.yy") + ":"
 ''sql = "SELECT * FROM MedPlan WHERE pat_id = " + CStr(Pat_id) + " and zeitpunkt = (SELECT  max([zeitpunkt]) AS zp FROM MedPlan WHERE pat_id = " + CStr(Pat_id) + ") ORDER BY [zeitpunkt] DESC , [feld], [feldnr]"
'' sql = "select * from [" + QMdbAkt + "].[Medikamente aktuell] where medplan.pat_id = " + CStr(Pat_id) + " order by anal, ins, metf, glib, gluci, shglin, hypt, hmg, thro, antib desc"
' sql = "SELECT *, medplan.medikament as mmedikament from medplan left join medarten on medplan.medikament = medarten.medikament WHERE mpnr = (select max(mpnr) from medplan where zeitpunkt = (select max(zeitpunkt) from medplan where pat_id = " + CStr(Pat_id) + ")) order by anal, ins, metf, glib, gluci, shglin, hypt, hmg, thro, antib desc"
 sql = "select medarten.*, medplan.medikament as mmedikament,bemerkung,mo,mi,nm,ab,zn,-bbed as j_bbed from medplan left join medarten on medplan.medanfang = medarten.medikament where medplan.mpnr = (select max(mpnr) from medplan where pat_id = " & Pat_id & " and zeitpunkt = (select max(zeitpunkt) from medplan where pat_id = " & CStr(Pat_id) & ")) and not isnull(medplan.medikament) and medplan.medikament <> '' and medplan.pat_id = " & Pat_id
' Set radat = Dtb.OpenRecordset(sql)
 Set raDat = Nothing
 raDat.Open "select count(*) as ct from (" & sql & ") as innen", DBCn, adOpenStatic, adLockReadOnly
 If Not raDat.BOF Then
   ZZ = raDat!ct
 End If
 raDat.Close
 raDat.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 
 Dim Tabl
 Set Tabl = dc.Tables.Add(Range:=dc.Bookmarks!Therapie.Range, NumRows:=ZZ + 1, NumColumns:=7, _
               DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitContent)
' raDat.MoveFirst
 runde = 2
 Do While Not raDat.EOF
  If IsNull(raDat!Bemerkung) Or raDat!Bemerkung = "" Then
  Else
   dc.Bookmarks!MedBemerkung.Range = raDat!Bemerkung
  End If
  Tabl.cell(runde, 1).Range = IIf(IsNull(raDat!mmedikament.Value), "", raDat!mmedikament.Value)
  Tabl.cell(runde, 2).Range = IIf(IsNull(raDat!mo), "", Replace(Replace(raDat!mo, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 3).Range = IIf(IsNull(raDat!mi), "", Replace(Replace(raDat!mi, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 4).Range = IIf(IsNull(raDat!nm), "", Replace(Replace(raDat!nm, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 5).Range = IIf(IsNull(raDat!ab), "", Replace(Replace(raDat!ab, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 6).Range = IIf(IsNull(raDat!Zn), "", Replace(Replace(raDat!Zn, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 7).Range = IIf(raDat!j_bBed = 0, "", "X")
  runde = runde + 1
  raDat.Move 1
 Loop
 Tabl.cell(1, 1) = "Medikament"
 Tabl.cell(1, 2) = "früh"
 Tabl.cell(1, 3) = "mittags"
 Tabl.cell(1, 4) = "nachm"
 Tabl.cell(1, 5) = "abends"
 Tabl.cell(1, 6) = "spät"
 Tabl.cell(1, 7) = "b.B."
 Tabl.Rows(1).Range.Font.Bold = True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in letzteMed/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' letzteMed(dc, Pat_id&)
#If False Then
Sub MedArtenUnausgefüllt()
  Const Abfr$ = "Medarten unausgefüllt"
  Dim qd$, i%, Feld1%, rs
  On Error GoTo fehler
  qd = "select * from [" + QMdbAkt + "].Medarten where not ("
   For i = 0 To Dtb.TableDefs!Medarten.Fields.Count - 1
     Select Case Dtb.TableDefs!Medarten.Fields(i).Name
      Case "falsch"
       Debug.Print "ausgespart: " + Dtb.TableDefs!Medarten.Fields(i).Name
      Case Else
       If Dtb.TableDefs!Medarten.Fields(i).Type = 1 Then
        qd = qd + IIf(Feld1, " or", "") & "" & "[" + Dtb.TableDefs!Medarten.Fields(i).Name + "]"
        Feld1 = -1
       End If
     End Select
   Next i
 qd = qd + ") order by hinzugefügt desc"
 Call DoCmd.Close(acQuery, Abfr, acSavePrompt)
 On Error Resume Next
' DoCmd.DeleteObject acQuery, Abfr
 DtbQuerydefsDelete Abfr
 On Error GoTo fehler
 Call DtbCreateQueryDef(Abfr, qd)
#If False Then
 Set rs = Dtb.OpenRecordset(Abfr)
 Do While Not rs.EOF
   For i = 0 To Dtb.TableDefs!Medarten.Fields.Count - 1
     Select Case Dtb.TableDefs!Medarten.Fields(i).Name
      Case "falsch"
      Case Else
       If Dtb.TableDefs!Medarten.Fields(i).Type = 1 Then
        If rs.Fields(Dtb.TableDefs!Medarten.Fields(i).Name) Then
         Debug.Print rs!Medikament, rs.Fields(Dtb.TableDefs!Medarten.Fields(i).Name).Name + " positiv"
         GoTo weiter
        End If
       End If
     End Select
   Next i
   Debug.Print rs!Medikament, " keines positiv"
weiter:
  rs.MoveNext
 Loop
#End If
 DoCmd.OpenQuery Abfr
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MedArtenUnausgefüllt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' MedArtenUnausgefüllt
#End If
Function FlagInit()
 Dim j%, k%
 On Error GoTo fehler
 For j = 0 To FEZ ' 1 würde reichen
  flag(j) = 0
  DiagSi(j) = 0
  DiagNamen(j) = ""
  For k = 0 To maxIcd
   ic(j, k) = ""
  Next k
  un(j) = 0
 Next j
 For j = 0 To BEZ ' 1 würde reichen, zur Sicherheit
  bflag(j) = 0
  bDiagsi(j) = ""
  bDN(j) = ""
  For k = 0 To bmaxIcd - 1
   bic(j, k) = ""
  Next k
  bun(j) = 0
 Next j
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FlagInit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FlagInit
Function DSeit(ranam)
 Dim rFanf As New ADODB.Recordset
 On Error GoTo fehler
    If IsNull(ranam("Diabetes seit")) Then
      DSeit = "?"
    ElseIf LCase(ranam("Diabetes seit")) = "bu" Then
''      DSeit = format$(rAnam!Vorgestellt, "mm\/yy")
'      DSeit = format$(Dtb.OpenRecordset("select fanf from faelle where fid = (select min(fid) from faelle where pat_id = " + CStr(rAnam!Pat_id) + ")")!Fanf, "mm\/yy")
      Call rFanf.Open("select fanf from faelle where fid = (select min(fid) from faelle where pat_id = " + CStr(ranam!Pat_id) + ")", DBCn, adOpenStatic, adLockReadOnly)
      If Not rFanf.BOF Then
       DSeit = Format$(rFanf!Fanf, "mm\/yy")
      End If
    Else: DSeit = CStr(ranam("Diabetes seit"))
    End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DSeit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DSeit(rAnam)

Sub Epikrise(dc, Pat_id&, vordat As Date, lddat As Date, obstumm%)
  Dim i%
  On Error GoTo fehler
  nzw = vbCr
  Dim Epi$, Titel$, DT$, Folge$, FZ%, Begl$, bglz%, Akkusat$, Akklang$, Nominat$
  Dim rNam As New ADODB.Recordset, ranam As New ADODB.Recordset, rDT As New ADODB.Recordset
  Dim obkomma%
  Dim lKrea!, GFR!, Alter%, DialZt%, DialAlter%
  Dim RRsyst%, RRdiast%
  Dim rALz As New ADODB.Recordset
  
  sql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar, Normbereich as NB " + _
   "FROM laborxus LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " + _
   "WHERE pat_id = " + CStr(Pat_id) + " and not exists (select * from laborneu where pat_id = " + CStr(Pat_id) + " and abkü = laborxwert.Abkü and wert = laborxwert.wert and zeitpunkt > laborxus.Eingang -3 and zeitpunkt < laborxus.Eingang+6)"
  On Error GoTo fehler
  Call FlagInit
'  Set rNam = TabÖff("Namen", "Pat_id")
  rNam.Open "select * from namen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
'  Set rAnam = TabÖff("Anamnesebogen", "Pat_id")
'  rAnam.Seek "=", Pat_id
'  ranam.Open "select -obmblausgeh as j_obmblausgeh, a.* from anamnesebogen a where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockOptimistic
  ranam.Open "select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockOptimistic
'  Set rDT = TabÖff("Diagnosen", "DiagSuch")
'  rNam.Seek "=", Pat_ID
  If rNam.BOF Then Exit Sub
  Titel = IIf(IsNull(rNam!Titel), "", rNam!Titel + " ")
  Akkusat = IIf(rNam!Geschlecht = "w", "Frau ", "Herrn ") + Titel + rNam!NachName
  Akklang = IIf(rNam!Geschlecht = "w", "Frau ", "Herrn ") + Titel + rNam!Vorname & " " & rNam!NachName
  Nominat = IIf(rNam!Geschlecht = "w", "Frau ", "Herr ") + Titel + rNam!NachName
  Alter = (Now - CDate(rNam!GebDat)) / 365.24
  Epi = "Bei " + Akkusat
  Select Case ranam!Diabetestyp
   Case "1": DT = " Typ-1-Diabetes"
   Case "2": DT = " Typ-2-Diabetes"
   Case "s": DT = " sekundärer Diabetes"
   Case "g": DT = " Gestationsdiabetes"
   Case "p": DT = "e pathologische Glucosetoleranz"
   Case "-": DT = " "
   Case "?": DT = " Diabetes mellitus"
   Case Else: DT = " Diabetes mellitus Typ " + ranam!Diabetestyp
  End Select
  Select Case ranam!Diabetestyp
   Case "g", "p"
    Epi = Epi & " ist ein" & DT & " diagnostiziert worden"
   Case "-"
    Dim GfV$
    If IsNull(ranam("Grund für Vorstellung")) Then
     GfV = ""
    Else
     GfV = ranam("Grund für Vorstellung") '![Grund für Vorstellung]
    End If
   Case Else
    Epi = Epi & " ist ein" & DT & " seit " & DSeit(ranam) & " bekannt "
  End Select
  FZ = 0
  Select Case ranam!Diabetestyp
   Case "1", "2", "s", "p"
    Dim j%, k%, obunter%
    DiagNamen(1) = "Retinopathie" 'H36.0
    DiagNamen(flNep) = "Nephropathie" 'N08.3
    DiagNamen(flNiI) = "Niereninsuffizienz" 'N18*, N19
    DiagNamen(4) = "koronare Herzerkrankung" 'I25.*
    DiagNamen(5) = "Z.n. Myokardinfarkt" 'I25.2, I21*, I22*
    DiagNamen(6) = "Z.n. ACVB" 'Z95.1
    DiagNamen(flAVK) = "periphere Angiopathie" 'I79.2, I79.9
    DiagNamen(8) = "periphere arterielle Verschlusskrankheit" '  "I73.9", "I74", "I70", "I73"
    DiagNamen(9) = "Z.n. Bypassoperation" ' Z95.88
    DiagNamen(10) = "Z.n. Amputation" 'Z44.1 (=Amputation), Z97 (=Prothese)
    DiagNamen(11) = "Arteriosklerose der hirnzuführenden Arterien" '"I65.9", "I65", "I66", "I68"
    DiagNamen(12) = "Z.n. Apoplex" ' "I63.9Z", "I64", "I63")
    DiagNamen(13) = "Mesenterialischämie" ' K55.*
    DiagNamen(flHyp) = "Hyperkeratosen" 'L84, L85.9
    DiagNamen(flDFS) = "diabetisches Fußsyndrom" 'L89.*
    DiagNamen(flPNP) = "periphere Polyneuropathie" ' G63.2
    DiagNamen(17) = "diabetische Mononeuropathie" ' G59.0
    DiagNamen(18) = "autonome Neuropathie" ' G99.0
    DiagNamen(19) = "nicht-alkoholische Steatohepatose" ' K71.6, K71.7, K76.9, K77.8, K76.0
    DiagNamen(20) = "Katarakt" 'H28.0
    DiagNamen(flLiph) = "Liphypertrophie" ' "L94.8", "L94", "E88"
    DiagNamen(22) = "Nekrobiose" ' L92.1
    DiagNamen(flChar) = "Diabetische Osteoarthropathie" ' "Charcot-Fuß" 'M14.6
    DiagNamen(24) = "Diabetische Arthropathie" ' M14.2
    DiagNamen(25) = "Diabetische Adhäsiopathie" 'M68.8
    DiagNamen(26) = "Erkrankung d.gastroint.Schleimhaut"
    DiagNamen(27) = "Manifestes Fußsyndrom"
    un(flNiI) = -1: un(5) = -1: un(6) = -1: un(8) = -1: un(9) = -1: un(10) = -1: un(12) = -1
    ic(1, 0) = "H36.0":
    ic(flNep, 0) = "N08.3": ic(flNiI, 0) = "N18": ic(flNiI, 1) = "N19"
    ic(4, 0) = "I25": ic(5, 0) = "I25.2": ic(5, 1) = "I21": ic(5, 2) = "I22"
    ic(6, 0) = "Z95.1"
    ic(flAVK, 0) = "I79.9": ic(flAVK, 1) = "I79.2": ic(flAVK, 2) = "I73"
    ic(8, 0) = "I73.9": ic(8, 1) = "I74": ic(8, 2) = "I770": ic(8, 3) = "I73"
    ic(9, 0) = "K55": ic(10, 0) = "Z44.1": ic(10, 1) = "Z97": ic(11, 0) = "I65": ic(11, 1) = "I66": ic(11, 2) = "I68"
    ic(12, 0) = "I64": ic(12, 1) = "I65"
    ic(13, 0) = "K55": ic(14, 0) = "L85.9": ic(flDFS, 0) = "L89"
    ic(flPNP, 0) = "G63.2": ic(17, 0) = "G59.0": ic(18, 0) = "G99.0"
    ic(19, 0) = "K71.6": ic(19, 1) = "K71.7": ic(19, 2) = "K76.0": ic(19, 3) = "K77.8": ic(19, 4) = "K76.9"
    ic(20, 0) = "H28.0"
    ic(flLiph, 0) = "L94.8": ic(flLiph, 1) = "E88": ic(22, 0) = "L92.1"
    ic(flChar, 0) = "M14.6": ic(24, 0) = "M14.2": ic(25, 0) = "M68.8"
    ic(flUlc, 0) = "K23": ic(flUlc, 1) = "K25": ic(flUlc, 2) = "K26": ic(flUlc, 3) = "K27": ic(flUlc, 4) = "K28": ic(flUlc, 5) = "K29"
    ic(flMFF, 0) = "L89.2": ic(flMFF, 1) = "L89.3": ic(flMFF, 2) = "L89.4"
'    rDT.index = "DiagSuch"
'    rDT.Seek "=", Pat_ID
'    If Not rDT.NoMatch Then
'     Do Until rDT.EOF
     Set rDT = Nothing
     rDT.Open "select * from diagnosen where pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockReadOnly
     Do While Not rDT.EOF
'      If rDT!Pat_id <> Pat_id Then Exit Do
      For j = 1 To FEZ
       For k = 0 To maxIcd
        If IsNull(ic(j, k)) Or ic(j, k) = "" Then GoTo w1:
        If (InStrB(rDT!ICD, ic(j, k)) > 0 And ic(j, k) <> "") And Not (InStrB("ZA", rDT!DiagSicherheit) > 0 And un(j) = 0) Then ' bei UnterDtagnosen ist Z.n. relevant
         flag(j) = -1
         If rDT!DiagSicherheit = "V" Then DiagSi(j) = True
        End If
       Next
w1:
      Next j
      rDT.MoveNext
     Loop
'    End If
    obunter = 0
    For j = 1 To FEZ
     If Not un(j) And obunter Then
      Folge = Folge + ")"
      obunter = 0
     End If
     If flag(j) And Not (j = flNep And obkNeph) Then
      If un(j) And Not un(j - 1) And flag(j - 1) Then
       Folge = Folge + " ("
       obunter = -1
      Else
       Folge = Folge + ", "
      End If
      If Not un(j) Then obunter = 0
      Folge = Folge + IIf(DiagSi(j), "V.a. ", "") + Replace(DiagNamen(j), "Diabet", "diabet")
      FZ = FZ + 1
     End If
    Next j
    If Folge <> "" Then Folge = Mid$(Folge, 3)
    
    If flag(flNep) And Not flag(flNiI) And obkNeph Then
       Dim tonRunde
      If Not obstumm Then
       For tonRunde = 1 To 10
        Call Sound(WinDir + "\media\Windows XP-Standard.wav")
       Next tonRunde
      End If
'     MsgBox "Nephropathie gestrichen!"
     flag(flNep) = 0 ' Unzutreffende Nephropathie streichen
     Call meld("Bei Pat." & Pat_id & " wurde im Brief die Nephropathie gestrichen!", obstumm)
    End If
    
    bDN(bflHyt) = "arterielle Hypertonie" ' I1*
    bDN(2) = "Übergewicht" ' E65*, E66*, E67*, E68*
    bDN(3) = "Fettstoffwechselstörung" ' E78*
    bDN(bflRau) = "Nikotinabusus" 'F17.1
    bic(bflHyt, 0) = "I1*"
    bic(2, 0) = "E65*": bic(2, 1) = "E66*": bic(2, 2) = "E67*": bic(2, 3) = "E68*"
    bic(3, 0) = "E78*"
    bic(bflRau, 0) = "F17.1"
'    rDT.index = "DiagSuch"
'    rDT.Seek "=", Pat_ID
'    Set rDT = Nothing
    Set rDT = Nothing
    rDT.Open " select * from diagnosen where pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockReadOnly
    If Not rDT.BOF Then
    Do Until rDT.EOF
      If rDT!Pat_id <> Pat_id Then Exit Do
      For j = 1 To BEZ
       For k = 0 To bmaxIcd - 1
        If IsNull(bic(j, k)) Or bic(j, k) = "" Then GoTo w2:
        If rDT!ICD Like bic(j, k) Then
'         If rDt!diagsicherheit <> "Z" Then
          bflag(j) = -1
          Select Case rDT!DiagSicherheit
'           Case Null, "", "G":  bDiagsi(j) = ""
           Case "V":            bDiagsi(j) = "V.a. "
           Case "Z":            bDiagsi(j) = "Z.n. "
           Case "A":            bDiagsi(j) = "Ausschluss "
           Case Else
            bDiagsi(j) = ""
            If InStrB(rDT!DiagText, "V.a.") > 0 Then
             bDiagsi(j) = "V.a. "
            ElseIf InStrB(rDT!DiagText, "Z.n.") > 0 Then
             bDiagsi(j) = "Z.n. "
            ElseIf InStrB(rDT!DiagText, "Ausschl") > 0 Then
             bDiagsi(j) = "Ausschluss "
            End If
          End Select
'         End If ' rDt!diagsicherheit <> "Z" Then
        End If ' rDt!ICD Like bic(j, K) Then
       Next
w2:
      Next j
      rDT.MoveNext
     Loop
    End If
    bglz = 0
    obunter = 0
    For j = 1 To BEZ
     If Not bun(j) And obunter Then
      Begl = Begl + ")"
     End If
     If bflag(j) Then
      If bun(j) And Not bun(j - 1) And bflag(j - 1) Then
       Begl = Begl + " ("
       obunter = -1
      Else
       Begl = Begl + ", "
      End If
      If Not bun(j) Then obunter = 0
      Begl = Begl + bDiagsi(j) + bDN(j)
      bglz = bglz + 1
     End If
    Next j
    If Begl <> "" Then Begl = Mid$(Begl, 3)
    '         Folge = Folge + IIf(rDt!diagsicherheit = "V", "V.a.", "") + DiagNamen(j)
    If FZ <> 0 Then
     Epi = Epi + IIf(FZ = 1, "mit der Folgeerkrankung ", "mit den Folgeerkrankungen ") + Folge
    Else
     Select Case ranam!Diabetestyp
      Case "g", "p"
      Case Else
       Epi = Epi + "ohne Hinweis auf Folgeerkrankungen"
     End Select
    End If
   Case Else ' Diabetestyp
  End Select
  If bglz <> 0 Then
   Epi = Epi + IIf(FZ = 0, ", mit ", " und ") + IIf(bglz = 1, "der Begleiterkrankung ", "den Begleiterkrankungen ") + Begl
  End If
  Epi = Epi + "." + nzw
' Therapieart
  Dim Therapie$, rang%(1), lHbA1c!, loopHbA1c!, maxHbA1c!, lHbA1cTag As Date, rAktLab As New ADODB.Recordset, TherLang$(1)
  For i = 0 To 1
   Select Case i
    Case 0
     If vordat = CDate(0) Then
      If IsNull(ranam!Ther1) Then
      Else
       Therapie = ranam!Ther1
      End If
     Else
      Therapie = TherUmw(TherArt(ranam!Pat_id, 0, ranam, vordat))
     End If
    Case 1
    Therapie = TherUmw(TherArt(ranam!Pat_id, 0, ranam, vordat)) 'rAnam!TherAkt
   End Select
   Select Case Therapie ' rAnam.Fields("Ther" + IIf(i = 0, "1", "Akt"))
    Case "Diät", "Diät?": rang(i) = 0: TherLang(i) = "diätetische und physikalische Therapie"
    Case "OAD":           rang(i) = 1: TherLang(i) = "Therapie mit oralen Antidiabetika"
    Case "Komb":          rang(i) = 2: TherLang(i) = "Kombinationstherapie mit OAD und Insulin"
    Case "CT", "(I)CT?", "I/CT?": rang(i) = 3: TherLang(i) = "einfache Insulintherapie"
    Case "ICT":           rang(i) = 4: TherLang(i) = IIf(ranam!Diabetestyp = "2", "prandiale Insulintherapie", "intensivierte Insulintherapie")
    Case "CSII":          rang(i) = 5: TherLang(i) = "Insulinpumpentherapie"
   End Select
  Next i
  lHbA1c = 0
  lHbA1cTag = CDate(0)
'  Set rAktLab = Dtb.OpenRecordset("Select Wert,Zeitpunkt from (" + sql + ") as sql1 where abkü = ""HBA1C"" order by zeitpunkt desc")
  rAktLab.Open "Select Wert,Zeitpunkt from (" + sql + ") as sql1 where abkü = ""HBA1C"" order by zeitpunkt desc", DBCn, adOpenDynamic, adLockReadOnly
  If Not rAktLab.BOF Then
   Do While Not rAktLab.EOF
    If Not IsNull(rAktLab!Wert) Then
     Dim vorHbA1c$
     vorHbA1c = Replace(Replace(rAktLab!Wert, ".", ","), "%", "")
     loopHbA1c = CDbl(IIf(IsNumeric(vorHbA1c), vorHbA1c, 0))
     If lHbA1c = 0 Then
      lHbA1c = loopHbA1c
      lHbA1cTag = rAktLab!Zeitpunkt
     End If
     If loopHbA1c > maxHbA1c Then maxHbA1c = loopHbA1c
    End If
    rAktLab.Move 1
   Loop
  End If
  
  Dim npLabZahl%, aktRund%
  If flag(flPNP) Then
   Set rAktLab = Nothing
   rAktLab.Open "select max(wert),max(zeitpunkt),langtext from (" + sql + ") as sql1 where abkü in (""B12"",""BORG"",""BORM"",""BLOG"",""BLOM"",""IEPN_S"",""IEPN_U"",""C-ANCA"",""P-ANC_S"",""TPHA"",""HIV1/2_K"",""TSH"",""TSH-L_K"",""T3"",""T4"") group by langtext", DBCn, adOpenDynamic, adLockReadOnly
   If Not rAktLab.BOF Then
'    rAktLab.MoveLast
'    npLabZahl = rAktLab.RecordCount
    Set rALz = Nothing
    Call rALz.Open("select count(*) as ct from (" & rAktLab.source & ") as innen", DBCn, adOpenDynamic, adLockReadOnly)
    If Not rALz.BOF Then
     npLabZahl = rALz!ct
    Else
     npLabZahl = 0
    End If
    If npLabZahl > 5 Or (npLabZahl > 3 And lHbA1c < 7.5) Then
     Epi = Epi + "Zur Differentialdiagnostik der Neuropathie wurden mit untenstehenden Ergebnissen folgende LaborWerte untersucht: "
'     rAktLab.MoveFirst
     aktRund = 0
     Do While Not rAktLab.EOF
      Epi = Epi + IIf(aktRund > 0, ", ", "") + rAktLab!Langtext
      rAktLab.Move 1
      aktRund = aktRund + 1
     Loop
     Epi = Epi + "." + nzw
    End If
   End If
  End If
  
  Dim neLabZahl%
  If flag(flNiI) Or flag(flNep) Then
   Set rAktLab = Nothing
   Call rAktLab.Open("select max(wert),max(zeitpunkt),langtext from (" + sql + ") as sql1 where abkü in (""C3"",""C4"",""IEPN_S"",""IEPN_U"",""C-ANCA"",""P-ANC_S"",""HIV1/2_K"",""GLOMQL"",""LUBARL"",""TUBAQL"",""CIC"") group by langtext", DBCn, adOpenDynamic, adLockReadOnly)
'   Set rAktLab = Dtb.OpenRecordset("select max(wert),max(zeitpunkt),langtext from [" + QMdbAkt + "].laborunion where pat_id = " + CStr(Pat_id) + " and abkü in (""C3"",""C4"",""IEPN_S"",""IEPN_U"",""C-ANCA"",""P-ANC_S"",""HIV1/2_K"",""GLOMQL"",""LUBARL"",""TUBAQL"",""CIC"") group by langtext")
   If Not rAktLab.BOF Then
 '   rAktLab.MoveLast
 '   neLabZahl = rAktLab.RecordCount
    Set rALz = Nothing
    Call rALz.Open("select count(*) as ct from (" & rAktLab.source & ") as innen", DBCn, adOpenDynamic, adLockReadOnly)
    If Not rALz.BOF Then
     neLabZahl = rALz!ct
    Else
     neLabZahl = 0
    End If
    If neLabZahl > 2 Then
     Epi = Epi + "Um andere Nephropathieursachen auszuschließen, wurden speziell bestimmt (Ergebnisse s.u.): "
'     rAktLab.MoveFirst
     aktRund = 0
     Do While Not rAktLab.EOF
      Epi = Epi + IIf(aktRund > 0, ", ", "") + rAktLab!Langtext
      rAktLab.Move 1
      aktRund = aktRund + 1
     Loop
     Epi = Epi + "." + nzw
    End If
   End If
  End If
' ""C3"",""C4"",
' Grenze zwischen Diagnostik und Therapie
  Epi = Epi + nzw
  If Not IsNull(TherLang(1)) And TherLang(1) <> "" Then
   If rang(1) > rang(0) Then
    Epi = Epi + Nominat + " ließ sich auf eine " + TherLang(1) + " einstellen."
   ElseIf rang(1) < rang(0) Then
    Epi = Epi + "Zuletzt wurde eine " + TherLang(1) + " gewählt."
   ElseIf rang(1) = rang(0) And rang(1) < 3 And lHbA1c >= 6.5 Then
    Epi = Epi + "Zuletzt war " + Akkusat + " noch auf eine " + TherLang(1) + " eingestellt."
   Else
    Epi = Epi + "Die " + TherLang(1) + " wurde fortgeführt und angepasst."
   End If
   Epi = Epi + nzw
  End If
  
  If ranam!Diabetestyp <> "-" Then
   Epi = Epi + "Eine sinnvolle Verteilung und individuell bemessene Häufigkeit der prä- und postprandialen Blutzuckermessungen"
   If rang(1) >= 2 Then
    Epi = Epi + " sowie die Vermeidung von Therapiefehlern bei der Insulinapplikation (Spritzorte, -technik, -zeitpunkte usw.)"
    If rang(1) >= 3 Then Epi = Epi + " und Insulindosisanpassung an Mahlzeiten, Bewegung und Blutzuckerentgleisungen"
   End If
   Epi = Epi + " wurde angestrebt." + nzw
'  If flag(flLiph) Then Epi = Epi + "Durch häufiges Spritzen in die selben Stellen waren Liphypertrophien aufgetreten, die in einigen Monaten wieder verschwinden und möglichst nicht an neuer Stelle auftreten sollten." + nzw
   If vordat <> CDate(0) And flag(flLiph) Then Epi = Epi + "Liphypertrophien als Folge zu häufigen Spritzens in dieselben Stellen sollten wegen der dort ungleichmäßigen Insulinwirkung vermieden werden." + nzw
  End If
' Gestationsdiabetes
  If ranam!Diabetestyp = "g" Then
   Epi = Epi + "Für den weiteren Verlauf gelten nach Leitlinien folgende Empfehlungen: " & vbCrLf
   If Now - lddat < 60 Then
    Epi = Epi + "Monatliche Ultraschalluntersuchungen ab der 24. SSW, besondere Vorsicht mit Tokolytika etc., Entbindung in geeigneter Klinik"
    If rang(1) >= 2 Then
     Epi = Epi + " mit Neonatologie, BZ perinatal 2-stündlich messen, Ziel 70 - 110 mg/dl"
    End If
    Epi = Epi + ", besondere Überwachung des Kindes"
    If rang(1) >= 2 Then
     Epi = Epi + " (mit Bestimmung von Hb, Hkt, Ca, Mg, Bili, auch ohne klinische Auffälligkeit, zwischen 3.u.5. Tag), "
    End If
    Epi = Epi + " naßchemische Blutglucosebestimmung im Kreißsaal, Frühestfütterung, Stillen, Information des Kinderarztes über die positive Familienanamnese." & vbCrLf
    Epi = Epi + "Für die Patientin wird weiterhin empfohlen: " & vbCrLf
    Epi = Epi + "Nüchterne und postprandiale Blutzuckermessungen am 2. Tag nach der Entbindung, bei pathologischem Ausfall diabetologische Weiterbetreuung, ansonsten oraler Glucosetoleranztest nach 6-12 Wochen und dann spätestens alle 2 Jahre, "
    Epi = Epi + "bei erneuter Schwangerschaft oraler Glucosetoleranztest im 1. Trimenon, in der 24.-28. Woche sowie in der 32.-34. Woche, jeweils falls zuvor negativ."
   Else
    Epi = Epi + "Oraler Glucosetoleranztest spätestens alle 2 Jahre, "
    Epi = Epi + "bei erneuter Schwangerschaft oraler Glucosetoleranztest im 1. Trimenon, in der 24.-28. Woche sowie in der 32.-34. Woche, jeweils falls zuvor negativ."
   End If
  End If ' rAnam!Diabetestyp = "g" Then
' Schulungen
  Dim zp1$, zpl$, schulz%
  schulz = SchulzBest(Pat_id, zp1, zpl, vordat)
  If schulz > 1 Then
   Epi = Epi + Nominat + " nahm bei uns zwischen " + zp1 + " und " + zpl + " an " + CStr(schulz) + " Gruppenschulungen teil." + nzw
  End If
  Dim rlar As New ADODB.Recordset, Vero$, VeroZ%
  sql1 = "select * from eintraege where pat_id = " + CStr(Pat_id) + " and art = ""lar"" and not inhalt like ""%umpe%"" and not inhalt like ""%nsul%"" "
  If vordat <> CDate(0) Then sql1 = sql1 + "and Zeitpunkt >= " & datform(vordat)
  sql1 = sql1 + " order by zeitpunkt"
  sql1 = "select distinct Inhalt from (" + sql1 + ") as innen"
'  Set rLar = Dtb.OpenRecordset(sql1)
  rlar.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
  If Not rlar.BOF Then Epi = Epi + "Ich verordnete: "
  VeroZ = 0
  Do While Not rlar.EOF And Not rlar.BOF
   Dim VeroI$
   VeroI = rlar!Inhalt
   VeroI = Replace(VeroI, "diabetes Therapie", "Diabetes-Therapie")
   VeroI = Replace(VeroI, "diabetes adapt", "diabetesadapt")
   VeroI = Replace(VeroI, "diabetes adapie", "diabetesadaptie")
   VeroI = Replace(VeroI, "undAbsatzrolle", "und Absatzrolle")
   If InStrB(VeroI, "D:") > 0 Then
    Vero = Left(VeroI, InStr(rlar!Inhalt, "D:") - 1)
   Else
    Vero = VeroI
   End If
   If Vero <> "" Then
     Epi = Epi + IIf(VeroZ > 0, ", ", "") + Vero
     VeroZ = VeroZ + 1
   End If
   rlar.Move 1
  Loop
  If Not rlar.BOF Then Epi = Epi + "." + nzw
  
' Antibiose
  Dim obHWI%, AntibText$
  AntibText = ""
  Set rDT = Nothing
  rDT.Open "select * from diagnosen where pat_id = " & Pat_id & " and (icd like ""N12%"" or icd like ""N30%"" or icd like ""N39%"")", DBCn, adOpenDynamic, adLockReadOnly
  If rDT.BOF Then obHWI = 0 Else obHWI = -1
  Dim rAntib As New ADODB.Recordset
  If obHWI Or flag(flMFF) <> 0 Then
   sql1 = "select distinct medplan.medikament, min(zeitpunkt) as zeitpunkt from medplan left join medarten on medplan.medanfang = medarten.medikament where medplan.pat_id = " & Pat_id & " and antib<>0"
'  sql1 = "SELECT * FROM medikamente mit arten] WHERE antib and medplan.pat_id = " & CStr(Pat_ID) & " "
   If vordat <> CDate(0) Then sql1 = sql1 + " and Zeitpunkt >= " & datform(vordat)
   sql1 = sql1 & " group by medplan.medikament;"
'   Set rAntib = Dtb.OpenRecordset(sql1)
   Call rAntib.Open(sql1, DBCn, adOpenDynamic, adLockReadOnly)
   obkomma = 0
   Do While Not rAntib.EOF
    If obkomma Then
     AntibText = AntibText + ", "
    Else
     If obHWI And Not flag(flMFF) Then
      AntibText = AntibText + "Bei der Diagnostik der diabetischen Nephropathie fiel ein Harnwegsinfekt auf, ich gab "
'     ElseIf (Not obHWI And flag(flMFF)) Then
'      AntibText = AntibText + "Zur Behandlung des Fußsyndroms gab ich "
     Else
      AntibText = AntibText + "Als Antibiose gab ich "
     End If
    End If
    On Error Resume Next
    Err.Clear
'    If lies.obMySQL Then
     AntibText = AntibText + rAntib![Medikament] + " ab " + Format$(rAntib!Zeitpunkt, "dd/mm/yy")
     If Err.Number <> 0 Then
'    Else
     AntibText = AntibText + rAntib![medplan.medikament] + " ab " + Format$(rAntib!Zeitpunkt, "dd/mm/yy")
    End If
    On Error GoTo fehler
    obkomma = -1
    rAntib.Move 1
   Loop
   If obkomma Then AntibText = AntibText + "." + nzw
  End If
  
  If flag(flMFF) Then
   Epi = Epi & "Das offene diabetische Fußsyndrom wurde nach den üblichen Regeln behandelt (Druckentlastung, Überprüfung der Durchblutung" & IIf(flag(flMFF) And AntibText <> "", ", Antibiose", "") & ")." & vbCrLf
   If AntibText <> "" Then
    Epi = Epi & AntibText & vbCrLf
   End If
  End If
  
  Dim obEintr%, mbl As New ADODB.Recordset
  Set mbl = Nothing
  mbl.Open "select art from eintraege where pat_id = " + CStr(Pat_id) + " and inhalt like ""%merkblatt%"" and (inhalt like ""%dfs%"" or inhalt like ""%fuß%"")", DBCn, adOpenDynamic, adLockReadOnly
'  If ranam!j_obMBlAusgeh <> 0 Or Not mbl.BOF Then
  If ranam!obMBlAusgeh <> 0 Or Not mbl.BOF Then
   If flag(flDFS) Then
     Epi = Epi + "Wegen des diabetischen Fußsyndroms wurde " + Akkusat + " in der Vorbeugung weiterer Läsionen unterwiesen."
   Else
    Epi = Epi + "Wegen "
    If flag(flAVK) Then Epi = Epi + "der peripheren Angiopathie "
    If flag(flPNP) Then
     If flag(flAVK) Then Epi = Epi + "und "
     Epi = Epi + "der peripheren Polyneuropathie "
     If flag(flChar) Then Epi = Epi + "und der vorliegenden Diabetischen Osteoarthropathie "
     If flag(flHyp) Then Epi = Epi + "sowie Hyperkeratosen "
     If flag(flAVK) Then
      Epi = Epi + "besteht eine erhöhte Gefährdung für das diabetische Fußsyndrom. Wir haben "
     Else
      Epi = Epi + "wurde "
     End If
    End If
    If Not flag(flAVK) And Not flag(flPNP) Then
     Epi = Epi + "der Gefahr eines diabetischen Fußsyndroms wurde "
    End If
    Epi = Epi + Akkusat + " in der Vorbeugung von " + IIf(flAVK Xor flPNP, "Fußl", "L") + "äsionen unterwiesen."
   End If
   Epi = Epi + nzw
  End If
  Call GetPrRR(ranam, RRsyst, RRdiast)
  If Pat_id <> ranam!Pat_id Then
   MsgBox "Programmfehler Identität!"
   Stop
  End If
  If diI("I10", Pat_id) Then
   Epi = Epi + nzw + "Blutdruckziel nach Leitlinien ("
'   Epi = Epi + "Für den Blutdruck gilt nach den aktuellen Leitlinien "
   If flag(flNep) Or flag(flNiI) Then
    Epi = Epi + "für bestehende Nephropathie "
    If flag(flNiI) Then Epi = Epi + "bzw. Niereninsuffizienz "
   Else
    Epi = Epi + "für nicht nachweisbare Nephropathie"
   End If
'   Epi = Epi + " ein Zielwert von "
   Epi = Epi + "): "
   If flag(flNep) Or flag(flNiI) Then
    Epi = Epi + "120/80 mm Hg, falls verträglich, andernfalls "
   End If
   Epi = Epi + "130/85 mm Hg (Praxismessung)."
   If rrEmpf(RRsyst, RRdiast, ranam!Pat_id) = "halten" Then
    Epi = Epi + " Dieses schien zuletzt erreicht worden zu sein."
   End If
   Epi = Epi + nzw
  End If
  If obMakroAlb And Alter < 61 Then
   Epi = Epi + "Dieses ist angesichts der bereits vorliegenden Makroalbuminurie und des noch relativ geringen Lebensalters besonders wichtig." + nzw
  End If
  
  Dim lLDL!, lLDLTag As Date
  lLDL = 0
  lLDLTag = CDate(0)
  Set rAktLab = Nothing
  Call rAktLab.Open("Select Wert,Zeitpunkt from (" + sql + ") as sql1 where (abkü = ""LDL"" or abkü = ""LDLH"") order by zeitpunkt desc", DBCn, adOpenDynamic, adLockReadOnly)
'  Set rAktLab = Dtb.OpenRecordset("Select Wert,Zeitpunkt from [" + QMdbAkt + "].Laborunion where pat_id = " + CStr(Pat_id) + " and (abkü = ""LDL"" or abkü = ""LDLH"") order by zeitpunkt desc")
  If Not rAktLab.EOF Then
   Dim rAktLabZWI$
   rAktLabZWI = Trim$(Replace(rAktLab!Wert, ".", ","))
   If rAktLabZWI = "" Then
    lLDL = 0
   Else
    lLDL = 0
    On Error Resume Next
    lLDL = CDbl(rAktLabZWI)
    On Error GoTo fehler
   End If
   lLDLTag = rAktLab!Zeitpunkt
  End If
' LDL
  If lLDL > 100 And Alter < 76 Then
'   Epi = Epi + "Für das LDL gilt nach den noch nicht aktualisierten Leitlinien ein Zielwert von 100 mg/dl." + nzw
   Epi = Epi + "LDL-Zielwert nach Leitlinien: 100 mg/dl (Untergrenze nicht belegbar)." + nzw
  End If
' ASS
  Dim AssA%, AssE%, AntikoagA, AntikoagE, keinAss
  Call TherAuskunft(Pat_id, 1, , , , , , , , , , , , AssA, AntikoagA)
  Call TherAuskunft(Pat_id, 0, , , , , , , , , , , , AssE, AntikoagE)
  keinAss = 0
  If flag(flUlc) Then keinAss = True
  If AssA Then keinAss = True
  If AssE Then keinAss = True
  If AntikoagE Then keinAss = True
  Set rDT = Nothing
  rDT.Open "select d.icd as icd, diagtext,dg1 from diagnosen d left join diagreihe dr on d.icd = dr.icd where pat_id = " & Pat_id & " and diagsicherheit <> 'A'", DBCn, adOpenDynamic, adLockReadOnly
  Do While Not rDT.EOF
  ' Leberkrankheiten außer Leberzirrhose, Obstipation
   If rDT!ICD Like "C*" Or (rDT!dg1 = "Magen-Darm-Leber" And Not (rDT!ICD Like "K7*" And Not rDT!ICD Like "K74*") And Not rDT!ICD Like "K59*") Or InStrB(rDT!DiagText, "lutu") > 0 Or InStrB(rDT!DiagText, "sthma") > 0 Then keinAss = True
   rDT.Move 1
  Loop
  If ranam!Diabetestyp <> "-" And (Now - ranam!GebDat) > 365 * 40 And Not keinAss Then
   Epi = Epi + Nominat + " gehört zum Personenkreis derer, die rein statistisch gesehen von einer Thrombozytenhemmung profitieren könnten." + nzw
  End If
  
  If flag(flNiI) Then
   Set rAktLab = Nothing
   Call rAktLab.Open("Select Wert,Zeitpunkt from (" + sql + ") as sql1 where abkü in (""KREA"", ""KREA02"", ""KRE02"", ""CREAT"") order by zeitpunkt desc", DBCn, adOpenDynamic, adLockReadOnly)
'   Set rAktLab = Dtb.OpenRecordset("Select Wert,Zeitpunkt from [" + QMdbAkt + "].laborunion where pat_id = " + CStr(Pat_id) + " and abkü = ""KREA"" order by zeitpunkt desc")
   If rAktLab.EOF Then GoTo keineGFR
   Dim sKrea$, obSchonDial%, obSchonNephr%
   Dim rDial As New ADODB.Recordset
   Dim obNiereWichtig%, obNephrologe%
     obSchonDial = 0
     obSchonNephr = 0
   sKrea = Replace(rAktLab!Wert, ".", ",")
   If sKrea = "" Then sKrea = "0"
   lKrea = CDbl(sKrea)
   If lKrea <> 0 Then
    GFR = (140 - Alter) * ranam!Gewicht * IIf(ranam!Gewicht < 3, 100, 1) * IIf(rNam!Geschlecht = "w", 0.85, 1) / lKrea / 72
    DialZt = (GFR - 10) / 3
    DialAlter = Alter + DialZt
    obNiereWichtig = (DialAlter < 80)
    obNephrologe = (GFR < 50)
    If obNiereWichtig Or obNephrologe Then
     Set rDial = Nothing
     rDial.Open "SELECT * FROM eintraege where pat_id = " & Pat_id & " and instrb(inhalt,'dialyse')>0", DBCn, adOpenStatic, adLockReadOnly
     If Not rDial.BOF Then
      obSchonDial = True
      obNephrologe = False
     Else
      Set rDial = Nothing
      rDial.Open "SELECT * FROM diagnosen where pat_id = " & Pat_id & " and icd like 'Z49%'", DBCn, adOpenStatic, adLockReadOnly
      If Not rDial.BOF Then obSchonDial = True: obNephrologe = False
     End If
    End If
    If obNephrologe Then
     Set rDial = Nothing
     rDial.Open "SELECT * FROM dokumente where pat_id = " & Pat_id & " and (dokname like '%Arenz%' or dokname like '%Stenglein%' or dokname like '%Habersetzer%' or dokname like '%Nephrol%')", DBCn, adOpenStatic, adLockReadOnly
     If Not rDial.BOF Then obNephrologe = False
    End If
   
    If Not obSchonDial Then
     If obNiereWichtig Then
      Epi = Epi + "Bei einer geschätzten GFR von ca. " + CStr(Int(GFR)) + " ml/min "
'    If DialAlter < 60 Then
'     Epi = Epi + "ist mit dem Erleben von Nierenversagen zu rechnen."
'    Else
'     Epi = Epi + "könnte mit dem Erleben von Nierenversagen zu rechnen sein."
'    End If
      Epi = Epi + "gilt es, deren Abnahmegeschwindigkeit durch Optimalhalten " + IIf(bflag(bflHyt), "von Blutdruck und Blutzucker", "des Blutzuckers") + IIf(bflag(bflRau) And bDiagsi(bflRau) = "", " und Einstellen des Rauchens", "") + " zu minimieren."
      Epi = Epi + IIf(ranam!Diabetestyp = "1", " Bei", " Nur bei") + " Typ-1-Diabetikern ist außerdem der Nutzen einer Eiweißrestriktion belegt."
     Else
      Epi = Epi + "Die GFR-Schätzung nach Cockcroft-Gault beträgt (mit einem Kreatinin von " + CStr(lKrea) + " mg/dl) ca. " + CStr(Int(GFR)) + " ml/min."
     End If
    End If

   Epi = Epi + nzw
   If obNephrologe Then
    If GFR < 30 Then
     Epi = Epi + "Eine weitere nephrologische Mitbehandlung halte ich für sinnvoll."
    Else
     Epi = Epi + "Eine nephrologische Mitbehandlung könnte sinnvoll sein."
    End If
   End If
   Epi = Epi + nzw
   End If ' Krea > 0
keineGFR:
  End If ' flag(flNiI) Then
  
  Dim rUebw As New ADODB.Recordset
  sql1 = "select feldinh,zeitpunkt from (" & forminhalt & " where pat_id = " & CStr(Pat_id) & ") as forminhalt where form_abk = ""uew"" and feld = ""Ueberweisung_an"" "
  If vordat <> CDate(0) Then sql1 = sql1 + "and Zeitpunkt >= " & datform(vordat)
  sql1 = sql1 + " order by zeitpunkt"
  rUebw.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
  
  If Not rUebw.BOF Then
   obkomma = 0
   Do While Not rUebw.EOF
    If Not obkomma Then Epi = Epi + "Weitere Überweisungen wurden von mir ausgestellt: "
    If obkomma Then Epi = Epi + ", "
    Epi = Epi + rUebw!FeldInh + " (" + Format$(rUebw!Zeitpunkt, "dd/mm/yy")
    Dim rBegrü As New ADODB.Recordset
    'Set rBegrü = Dtb.OpenRecordset("select feldinh from FormInhaltForm_Abk INNER JOIN " + _
                 "(FormInhaltFeldInh INNER JOIN (FormInhaltFeld INNER JOIN FormInhaltNeu ON " + _
                 "FormInhaltFeld.FeldVW = FormInhaltNeu.FeldVW) ON FormInhaltFeldInh.FeldInhVW = " + _
                 "FormInhaltNeu.FeldInhVW) ON FormInhaltForm_Abk.Form_AbkVW = FormInhaltNeu.Form_AbkVW " + _
                 "WHERE Pat_id = " + CStr(Pat_id) + " And ZeitPunkt = CDate(""" + _
                 format$(rUebw!ZeitPunkt, "DD/MM/YY hh:mm") + """) And form_abk = ""uew"" And Feld = ""Diagnose"" " + _
                 "order by feldnr")
'    Set rBegrü = Dtb.OpenRecordset("SELECT [Pat_ID], [FID], [Form_ID], [ZeitPunkt], [Nr], [FeldNr], [Feld], [FeldInh], form_abk FROM (((forminhfeld LEFT JOIN forminhkopf ON [forminhfeld].[foid]=[forminhkopf].[foid]) LEFT JOIN formulare ON [formulare].[formid]=[forminhkopf].[form_id]) LEFT JOIN forminhaltfeld ON [forminhfeld].[feldvw]=[forminhaltfeld].[feldvw]) LEFT JOIN forminhaltfeldinh ON [forminhfeld].[feldinhvw]=[forminhaltfeldinh].[feldinhvw] WHERE Pat_id = " + CStr(Pat_ID) + " And ZeitPunkt = CDate(""" + format$(rUebw!Zeitpunkt, "DD/MM/YY hh:mm") + """) And form_abk = ""uew"" And Feld = ""Diagnose"" " + " order by feldnr")
    Set rBegrü = Nothing
    rBegrü.Open "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((forminhfeld LEFT JOIN forminhkopf ON forminhfeld.foid=forminhkopf.foid) LEFT JOIN formulare ON formulare.formid=forminhkopf.form_id) LEFT JOIN forminhaltfeld ON forminhfeld.feldvw=forminhaltfeld.feldvw) LEFT JOIN forminhaltfeldinh ON forminhfeld.feldinhvw=forminhaltfeldinh.feldinhvw WHERE Pat_id = " + CStr(Pat_id) + " And ZeitPunkt = " & datform(rUebw!Zeitpunkt) & " And form_abk = ""uew"" And Feld = ""Diagnose"" " + " order by feldnr", DBCn, adOpenDynamic, adLockReadOnly
    If Not rBegrü.BOF Then
     Epi = Epi + ", """
     Do While Not rBegrü.EOF
      Epi = Epi + rBegrü!FeldInh
      rBegrü.Move 1
      If Not rBegrü.EOF Then Epi = Epi + " "
     Loop
     Epi = Epi + """"
    End If
    Epi = Epi + ")"
    obkomma = -1
    rUebw.Move 1
   Loop
   If obkomma Then Epi = Epi + nzw
  End If ' Not rUebw.BOF Then
  
  Dim rKh As New ADODB.Recordset
'  Set rKH = TabÖff("KHEinweis", "Auswahl")
  rKh.Open "select * from kheinweis where pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockReadOnly
'  rKH.Seek "=", Pat_ID
  If Not rKh.BOF Then
   obkomma = 0
   Do While Not rKh.EOF
    If rKh!Pat_id <> Pat_id Then Exit Do
    If rKh!Zeitpunkt > vordat Then
     If Not obkomma Then Epi = Epi + "Folgende Einweisungen wurden von mir ausgestellt: "
     If obkomma Then Epi = Epi + ", "
     Epi = Epi + Format$(rKh!Zeitpunkt, "dd/mm/yy") & " " & IIf(InStrB(rKh!Ziel, "einweisung") = 0, rKh!Ziel + " ", "") + _
           IIf(IsNull(rKh!Diagnose), "", "(" + rKh!Diagnose + ")")
     obkomma = -1
    End If ' rkh!ZeitPunkt > VorDat Then
    rKh.Move 1
   Loop
   If obkomma Then Epi = Epi + nzw
  End If ' Not rkh.NoMatch Then
  
  Dim rAu As New ADODB.Recordset
'  Set rAu = TabÖff("AU", "Auswahl")
  rAu.Open "select * from au where pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockReadOnly
'  rAu.Seek "=", Pat_ID
  If Not rAu.EOF Then
   obkomma = 0
   Do
    If rAu.EOF Then Exit Do
    If rAu!Pat_id <> Pat_id Then Exit Do
    If rAu!Pat_id <> Pat_id Then Exit Do
    If rAu!Zeitpunkt > vordat Then
     If Not obkomma Then Epi = Epi + "Für folgende Zeiträume wurden Arbeitsunfähigkeitsbescheinigungen ausgestellt: "
     If obkomma Then Epi = Epi + ", "
     Epi = Epi + IIf(rAu!Beginn = "00000000", "", Format$(Left(rAu!Beginn, 4) + Right$(rAu!Beginn, 2), "##\.##\.##")) + "-" + Format$(Left(rAu!Ende, 4) + Right$(rAu!Ende, 2), "##\.##\.##")
     obkomma = -1
    End If
    rAu.Move 1
   Loop
   If obkomma Then Epi = Epi + nzw
  End If ' Not rAu.NoMatch Then
  
  
  If obHWI Then
   Epi = Epi + AntibText
  End If
  
  If ranam!Diabetestyp <> "-" And ranam!Diabetestyp <> "g" Then Epi = Epi + nzw + "Kontrolluntersuchungen gemäß dem Gesundheitspass Diabetes werden empfohlen."
  Epi = Epi + nzw
  
  dc.Bookmarks("Epikrise").Range = Epi
  Dim rF 'As Range
  Dim SuchText$(), STakt
  ReDim SuchText(6)
  SuchText(0) = "Eine sinnvolle Verteilung"
  SuchText(1) = "Liphypertrophien als Folge zu häufigen"
  SuchText(2) = "Ich verordnete"
  SuchText(3) = "Blutdruckziel nach Leitlinien"
  SuchText(4) = "LDL-Zielwert nach Leitlinien"
  SuchText(5) = "Weitere Überweisungen"
  SuchText(6) = "Folgende Einweisungen"
  For Each STakt In SuchText
   Set rF = dc.Range
   With rF.Find
    .clearformatting
    .Text = STakt
    .Execute
    If .found Then
     rF.Paragraphs(1).Range.Font.Size = 9
    End If
   End With
  Next

  dc.Bookmarks!kopfgrenze.Range = Replace(Replace(IIf(ranam!Diabetestyp = "-", dc.Bookmarks!kopfgrenze.Range, Replace(dc.Bookmarks!kopfgrenze.Range, "Diabetologischer ", "")), "Patname", Akklang), "Patgeb", Format$(ranam!GebDat, "dd/mm/yyyy"))
  Exit Sub
  Dim ii%, jj%
  For ii = 1 To dc.sections.Count
    For jj = 1 To dc.sections(ii).Headers.Count
     Dim kra As Range
     Set kra = dc.Range(dc.sections(ii).Headers(jj).Range.Start, dc.Bookmarks("KopfGrenze").Start)
     kra = Replace(kra, "Patname", Akklang)
     kra = Replace(kra, "Patgeb", Format$(rNam!GebDat, "dd/mm/yyyy"))
     If ranam!Diabetestyp = "-" Then dc.sections(ii).Headers(jj).Range = Replace(dc.sections(ii).Headers(jj).Range, "Diabetologischer ", "")
    Next
  Next ii
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Epikrise/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Epikrise
Function SchulzBest%(Pat_id&, Optional zp1$, Optional zpl$, Optional abDat)
  Dim sqls$, lapp As New ADODB.Recordset
  On Error GoTo fehler
  sqls$ = "SELECT * FROM eintraege where pat_id = " & CStr(Pat_id) + " and art = ""schul"" "
  If Not IsMissing(abDat) Then If Now > CDate("15.10.05") Then sqls = sqls + "and Zeitpunkt >= " & datform(abDat)
  sqls = sqls + " order by zeitpunkt"
'  Set lpp = Dtb.OpenRecordset(sqls, dbOpenDynaset)
  lapp.Open sqls, DBCn, adOpenDynamic, adLockReadOnly
  If Not lapp.BOF Then
   lapp.MoveLast
   zpl = Format$(lapp!Zeitpunkt, "dd/mm/yy")
   lapp.MoveFirst
   zp1 = Format$(lapp!Zeitpunkt, "dd/mm/yy")
   Set lapp = Nothing
   lapp.Open "select count(*) as ct from (" & sqls & ") as innen", DBCn, adOpenStatic, adLockReadOnly
   SchulzBest = lapp!ct
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SchutzBest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'SchulzBest
Public Sub doViewsErstellen()
'On Error Resume Next
Const frist% = 14  ' Tage
Const vz& = 16
Dim i&
Dim vn$(vz)
Dim vsql$(vz)
syscmd 4, "Erstelle Views für " & DBCn
vn(0) = "_lfaelle" ' letzte Fälle bis zum Quartal der Frist
vn(1) = "lfaelle"
vn(2) = "labor1"
vn(3) = "labor2"
vn(4) = "formular"
vn(5) = "_kontakttage"
vn(6) = "kontaktzahl"
vn(7) = "aktfaelle"
vn(8) = "letzte faelle"
vn(9) = "_faellenachschgr"
vn(10) = "faelleverschieden"
vn(11) = "aktfaellev"
vn(12) = "lfaellev"
vn(13) = "DMPInkonsistenzen"
vn(14) = "faelleverschiedenneu"
vn(15) = "aktf"
vn(16) = "aktfv"

vsql(0) = "select max(bhfb) as mbhfb, pat_id, min(schgr) as mschgr from faelle where str_to_date(concat('01.'¡(left(quartal,1)-1)*3+1¡'.'¡substr(quartal,2,5)),'%d.%m.%Y') < subdate(now(),interval " & frist & " day) group by pat_id order by min(schgr)"
vsql(1) = "select f.* from faelle as f inner join _lfaelle as i on f.pat_id = i.pat_id and f.bhfb = i.mbhfb;"
vsql(2) = "select n.Pat_ID AS Pat_ID,cast(n.ZeitPunkt as date) AS ZeitPunkt,n.FertigStGrad AS FertigStGrad,n.Abkü AS Abkü,l.Langtext AS Langtext,n.Wert AS Wert,n.Einheit AS Einheit,k.Kommentar AS Kommentar,_utf8'' AS nb from (laborneu n left join laborlangtext l on l.langtextvw = n.langtextvw) left join laborkommentar k on k.KommentarVW = n.kommentarvw  where wert <> '' and not isnull(wert) ;"
vsql(3) = "select u.Pat_id AS Pat_ID,cast(u.Eingang as date) AS zeitpunkt,u.BefArt AS FertigStGrad,w.Abkü AS Abkü,w.Langname AS Langtext,w.Wert AS Wert,w.Einheit AS Einheit,w.Kommentar AS Kommentar,w.Normbereich AS NB from (laborxus u left join laborxwert w on((u.RefNr = w.RefNr))) where (not(exists(select 1 AS Not_used from laborneu where ((laborneu.Pat_ID = u.Pat_id) and (laborneu.Abkü = w.Abkü) and (laborneu.Wert = w.Wert) and (laborneu.ZeitPunkt > subdate(u.Eingang,interval 3 day)) and (laborneu.ZeitPunkt < adddate(u.Eingang, interval 6 day))))))  and wert <> '' and not isnull(wert) ;"
vsql(4) = "select forminhkopf.FoID AS foid,forminhkopf.Pat_ID AS Pat_ID,forminhkopf.FID AS FID,forminhkopf.Form_ID AS Form_ID,forminhkopf.ZeitPunkt AS ZeitPunkt,forminhfeld.Nr AS Nr,forminhfeld.FeldNr AS FeldNr,forminhaltfeld.Feld AS Feld,forminhaltfeldinh.FeldInh AS FeldInh,formulare.Form_Abk AS form_abk,formulare.FormVorl AS FormVorl from ((((forminhfeld left join forminhkopf on((forminhfeld.FoID = forminhkopf.FoID))) left join formulare on((formulare.FormID = forminhkopf.Form_ID))) left join forminhaltfeld on((forminhfeld.FeldVW = forminhaltfeld.FeldVW))) left join forminhaltfeldinh on((forminhfeld.FeldInhVW = forminhaltfeldinh.FeldInhVW))) order by forminhkopf.FoID;"
If LVobMySQL And 1 = 0 Then
 vsql(5) = "select pat_id,first(zeitpunkt) from eintraege e where ZeitPunkt between cdate(concat(year(subdate(now(),interval " & frist & " day))¡'-'¡intacc((month(subdate(now(),interval " & frist & " day))-1)divmy 3)*3+1¡'-01')) and cdate(adddate(concat(year(subdate(now(),interval " & frist & " day))¡'-'¡intacc((month(subdate(now(),interval " & frist & " day))-1)divmy 3)*3+1¡'-01'), interval 3 month)) and e.Art in ('notiz','ni','telef','gs','rz','ga','ag','hz','ts','cr','ep','tst','wr','ek','ph','bga','gstel','rz','ga','ag','hz','ts','ke','uew','rp','sono','duplex','andm','anal')  group by e.Pat_ID,intaccdatemy(zeitpunkt)"
Else
 vsql(5) = "select pat_id,first(zeitpunkt) from eintraege e where ZeitPunkt between cdate(concat(year(subdate(now(),interval " & frist & " day))¡'-'¡intacc((month(subdate(now(),interval " & frist & " day))-1)divmy 3)*3+1¡'-01')) and cdate(       (concat(year(subdate(now(),interval " & frist & " day)) + round((intacc((month(subdate(now(),interval " & frist & " day))-1)divmy 3)*3+4) / 12)¡'-'¡(intacc((month(subdate(now(),interval " & frist & " day))-1)divmy 3)*3+4) mod 12¡'-01')                  )) and e.Art in ('notiz','ni','telef','gs','rz','ga','ag','hz','ts','cr','ep','tst','wr','ek','ph','bga','gstel','rz','ga','ag','hz','ts','ke','uew','rp','sono','duplex','andm','anal')  group by e.Pat_ID,intaccdatemy(zeitpunkt)"
End If
'vsql(5) = "select pat_id,zeitpunkt from eintraege e where ZeitPunkt between str_to_date(concat('#'¡intacc((month(subdate(now(),interval " & frist & " day))-1)divmy 3)*3+1¡'/01/'¡year(subdate(now(),interval " & frist & " day))¡'#'),'#%m/%d/%Y#') and adddate(str_to_date(concat('#'¡intacc((month(subdate(now(),interval " & frist & " day))-1)divmy 3)*3+1¡'/01/'¡year(subdate(now(),interval " & frist & " day))¡'#'),'#%m/%d/%Y#'), interval 3 month) and e.Art in ('notiz','ni','telef','gs','rz','ga','ag','hz','ts','cr','ep','tst','wr','ek','ph','bga','gstel','rz','ga','ag','hz','ts','ke','uew','rp','sono','duplex','andm','anal')  group by e.Pat_ID,date(zeitpunkt)"
vsql(6) = "select count(0) as ct,pat_id from _kontakttage group by pat_id;"
If LVobMySQL Then
 vsql(7) = "select f.pat_id as pid, notiz, stru.leistung as stru, chron.leistung as chron, kt.ct as kt, ebm.leistung as verspau, icd, f.*, id, vk, k.name as kname, kateg, anzahlik, anzahlktug, gültigvon, gültigbis, go, kurzname from ((faelle f left join kassenliste k on f.vknr = k.vk and f.ik = k.ik) left join diagnosen d on f.pat_id = d.pat_id and ((icd like 'E1%' and icd not like 'E16%' and icd not like 'E15%') or icd = 'O24.4') and diagsicherheit <> 'A' and (obdauer = 1 or d.fid = f.fid)) left join leistungen ebm on f.fid = ebm.fid and (ebm.leistung like '031%' or ebm.leistung like '01210') left join leistungen chron on f.fid = chron.fid and (chron.leistung = '03212') " & _
"left join leistungen stru on f.fid = stru.fid and (stru.leistung like '973%') " & _
"left join kontaktzahl kt on kt.pat_id = f.pat_id " & _
"left join namen n on n.pat_id = f.pat_id " & _
"where schgr <> '90' and quartal = (selectmy  concat(intacc(((month(subdate(now(),interval " & frist & " day))-1) divmy 3) + 1) ¡ year(subdate(now(),interval " & frist & " day)))) " & _
"group by f.fid order by pid, schgr, icd"
End If
vsql(16) = "select pat_id, fid,schgr,vknr from faelle where schgr <> '90' and quartal = (selectmy  concat(intacc(((month(subdate(now(),interval " & frist & " day))-1) divmy 3) + 1) ¡ year(subdate(now(),interval " & frist & " day)))) " & _
           "group by pat_id order by pat_id, schgr;"
vsql(15) = "select pat_id, fid,schgr,vknr from faelle where schgr <> '90' and quartal = (selectmy  concat(intacc(((month(subdate(now(),interval " & frist & " day))-1) divmy 3) + 1) ¡ year(subdate(now(),interval " & frist & " day)))) " & _
           "order by pat_id, fid desc, schgr;"
vsql(8) = vsql(1)
vsql(9) = "select * from faelle order by schgr"
If LVobMySQL Then
 vsql(10) = "select * from _faellenachschgr as f group by pat_id,quartal"
Else
 vsql(10) = "select faelle.* from (SELECT first(fid) as mfid, pat_id, quartal FROM _faellenachschgr f GROUP BY f.Pat_ID, f.Quartal) as f inner join faelle on f.mfid = faelle.fid"
End If
vsql(11) = "select * from aktfaelle group by pat_id"
vsql(12) = "select * from lfaelle group by pat_id"
If 1 = 0 Then
vsql(13) = "select f.quartal, count(0) as ct, o.pat_id, o.fid, o.zeitpunkt, feldinh, form_abk from formular o left join faelle f using (fid)" & _
           "where form_abk like 'dmp%' and " & _
           "feld = 'Nachname' and " & _
           "(`f`.`Quartal` = " & _
           "(select concat((((month((now() - interval 23 day)) - 1) DIV 3) + 1),year((now() - interval 23 day))) )) " & _
           "group by pat_id " & _
           "having ct > 1;"
End If
If LVobMySQL Then
 vsql(14) = "select *,(select min(fanf) from faelle f1 where pat_id = f.pat_id and f1.fanf < f.fanf) as erst  from _faellenachschgr as f group by pat_id,quartal"
Else
 vsql(14) = "select (select min(fanf) from faelle f1 where pat_id = f0.pat_id and f1.fanf < ffanf) as erst, faelle.* from (SELECT first(fid) as mfid, first(fanf) as ffanf, pat_id, quartal FROM _faellenachschgr f GROUP BY f.Pat_ID, f.Quartal) as f0 inner join faelle on f0.mfid = faelle.fid"
End If


For i = 0 To vz
' If i = 7 Then Stop
 Call DtbCreateQueryDef(vn(i), vsql(i))
Next i
Forms(0).Ausgabe = "Views für " & DBCn & " erstellt" & vbCrLf + altAusgabe
altAusgabe = Forms(0).Ausgabe
syscmd 5
Exit Sub

End Sub
Sub LaborIns(dc As Object, Pat_id&) 'Word.Document, Pat_id&)
 Dim raLW As New ADODB.Recordset, raDat As New ADODB.Recordset, rLaU As New ADODB.Recordset, ls$
' Zeilenzahl bestimmen
 Dim ZZ&, rz$, gschl$, vgl$, altGruppe%, NB$ ' Normbereich
 Dim sql1$
 Dim u!, o! ' oberer und unterer Grenzwert numerisch
 Dim uNG$, oNG$ ' obere und untere Normgrenze in Zeichen
 Dim unm$, onm$, unw$, onw$
 Dim pKz$ ' Pathologisch-Kennzeichen
 Dim SelbstStatus
 Dim Matr$(), mBreiten$()
 Dim altgru$
 On Error GoTo fehler
 sql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor where not (Wert like ""%eine Berechnung%"") " & _
  "UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich as NB " + _
  "FROM laborxus LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " + _
  "WHERE pat_id = " + CStr(Pat_id) + " and not (Wert like ""%eine Berechnung%"") and not exists (select * from laborneu where pat_id = " + CStr(Pat_id) + " and abkü = laborxwert.Abkü and wert = laborxwert.wert and zeitpunkt > laborxus.Eingang-3 and zeitpunkt < laborxus.Eingang+6)"
' sql = "select * from labor1 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' order by zeitpunkt desc union select * from labor2 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' order by zeitpunkt desc"
 sql = "select * from labor1 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' order by zeitpunkt desc union select * from labor2 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' order by zeitpunkt desc"
 sql = "select * from labor1 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' union select * from labor2 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%'"
 On Error GoTo fehler
 SelbstStatus = 0
' Set rDat = Dtb.OpenRecordset("Select * from [" + QMdbAkt + "].labor where pat_id = " & CStr(Pat_id) & " and kommentar = ""Dieser Eintrag wurde manuell erzeugt.""")
 sql1 = "Select * from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor where kommentar = ""Dieser Eintrag wurde manuell erzeugt."""
 raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 If Not raDat.BOF Then SelbstStatus = SelbstStatus + 1
' Set rDat = Dtb.OpenRecordset("Select * from [" + QMdbAkt + "].labor where pat_id = " & CStr(Pat_id) & " and kommentar <> ""Dieser Eintrag wurde manuell erzeugt.""")
 sql1 = "Select * from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor where kommentar <> ""Dieser Eintrag wurde manuell erzeugt."""
 Set raDat = Nothing
 raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 If Not raDat.BOF Then SelbstStatus = SelbstStatus + 2
 Select Case SelbstStatus
  Case 0
   Dim rH
   Set rH = dc.Range(dc.Bookmarks("LaborÜS").Range.Start - 1, dc.Bookmarks("LaborÜS").Range.Start - 1)
   dc.Bookmarks("LaborÜS").Range = ""
   dc.Bookmarks.Add Name:="LaborÜS", Range:=dc.Range(rH.Start + 1, rH.Start + 1)
  Case 1
   dc.Bookmarks("LaborÜS").Range = "Laborbefunde (nur mitgeteilte, woanders bestimmte Werte):"
  Case 2
   dc.Bookmarks("LaborÜS").Range = "Bei mir erhobene Laborbefunde:"
 End Select
 
 syscmd 4, "Labor nach Überschrift"
 
 gschl = ""
 altGruppe = 0
' Set rLW = TabÖff("Namen", "Pat_id")
' ralw.Seek "=", Pat_id
 raLW.Open "select geschlecht from namen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 If Not raLW.EOF Then
  gschl = raLW!Geschlecht
 End If
 raLW.Close
 Dim sql0$
 Dim sql2$
nochmal:
 sql2 = "select * from labor1 where pat_id = " & Pat_id & " union select * from labor2 where pat_id = " & Pat_id
 sql0 = "SELECT distinct gruppe, reihe, sql1.abkü, sql1.einheit from (" + sql2 + ") as sql1 left join laborparameter on sql1.abkü = laborparameter.abkü and iif(isnull(sql1.einheit) or sql1.einheit="""", ""kA"",sql1.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar)) and reihe <> 999 order by gruppe, reihe"
 If lies.obMySQL Then sql0 = Replace(sql0, "iif(", "if(")
 Set raLW = Nothing
' GoTo nochmal
 raLW.Open sql0, DBCn, adOpenDynamic, adLockReadOnly
 altgru = ""
 Do While Not raLW.EOF
  If raLW!gruppe <> altgru Then
   ZZ = ZZ + 1
   altgru = raLW!gruppe
  End If
  ZZ = ZZ + 1
  raLW.MoveNext
 Loop
 
 syscmd 4, "Labor nach Zeilenzahlbestimmung"
' sql1 = "select count(*) from (SELECT distinct abkü from (" + sql + ") as sql1 where (not isnull(wert) or not isnull(kommentar))) as innen"
'' sql = "select count(*) from (SELECT distinct abkü from [" + QMdbAkt + "].laborunion where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar)))"
'' Set raLW = Dtb.OpenRecordset(sql1)
' Set raLW = Nothing
' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
' zZ = raLW.Fields(0)
' raLW.Close
' sql1 = "select count(*) from (SELECT distinct gruppe from (" + sql + ") as sql1 left join laborparameter on sql1.abkü = laborparameter.abkü and iif(isnull(sql1.einheit) or sql1.einheit="""", ""kA"",sql1.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar))) as innen"
'' sql = "select count(*) from (select distinct gruppe from (SELECT * from [" + QMdbAkt + "].laborunion left join [" + QMdbAkt + "].laborparameter on laborunion.abkü = laborparameter.abkü and iif(isnull(laborunion.einheit) or laborunion.einheit="""", ""kA"",laborunion.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar))))"
'' Set rLW = Dtb.OpenRecordset(sql1)
' If lies.obmysql Then sql1 = Replace(sql1, "iif(", "if(")
' Set raLW = Nothing
' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
' zZ = zZ + raLW.Fields(0)
' raLW.Close
nochmal1:
 sql1 = "select distinct * from (SELECT distinct datevalue(zeitpunkt) as Datum from (" + sql + ") as sql1 left join laborparameter on sql1.abkü = laborparameter.abkü and iif(isnull(sql1.einheit) or sql1.einheit="""", ""kA"",sql1.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar)) and reihe <> 999) as i0 order by datum"
' das erste distinct ist für Access nötig
' Set rDat = Dtb.OpenRecordset("select count(*) from (" + sql1 + ")")
' rZ = radat.Fields(0) ' dateserial(year(zeitpunkt),month(zeitpunkt),day(zeitpunkt))
 Set raDat = Nothing
 If lies.obMySQL Then
  sql1 = Replace(Replace(sql1, "datevalue(", "date("), "iif(", "if(")
 End If
' GoTo nochmal1
 raDat.Open "select count(*) as ct from (" + sql1 + ") as innen", DBCn, adOpenDynamic, adLockReadOnly
 If Not raDat.BOF Then
  rz = raDat!ct
 End If
 Const MaxBreite% = 17
 Dim DiffBr%
 DiffBr = 0
 If rz > MaxBreite Then
   DiffBr = rz - MaxBreite
 End If
 raDat.Close
 ReDim Matr(rz + 5, ZZ + 1) ' links noch: 0:Abkü, 1:unterer NW, 2:oberer NW, dann 3:Langtext, 4:Einheit, 5:Normbereich
 ReDim mBreiten(rz)
 
 syscmd 4, "Labor nach Diffbr-Bestimmung"
 
 Set raDat = Nothing
 raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
' Set rDat = Dtb.OpenRecordset(sql1)
 If raDat.BOF Then
    dc.Bookmarks!LaborÜS.Range = ""
    dc.Bookmarks!LabHinw.Range = ""
    Exit Sub
 End If
 Dim i&, j&
 i = 1
' raDat.Move DiffBr
 Do While Not raDat.EOF
  Matr(i + 5, 0) = Format$(raDat!datum, "dd.mm.yy")
  i = i + 1
  raDat.Move 1
  If i > rz Then Exit Do
 Loop
 
 syscmd 4, "Labor nach Datumseintragung"
 
 Dim rLP As New ADODB.Recordset
 j = 1
 altgru = ""
 raLW.MoveFirst
 Do While Not raLW.EOF
  If raLW!gruppe <> altgru Then
   altgru = raLW!gruppe
   Matr(0, j) = "Gruppe " & raLW!gruppe & "|" & raLW!reihe
'   Debug.Print Matr(0, j) & "," & Matr(1, j) & "," & Matr(2, j)
   j = j + 1
  End If
  Matr(0, j) = raLW!Abkü
  Matr(4, j) = raLW!Einheit
  Set rLP = Nothing
  Dim rLpgeht%
  rLpgeht = -1
  rLP.Open "select * from laborparameter where abkü = '" & raLW!Abkü & "' and einheit = '" & raLW!Einheit & "' and (unm <> """" or unw <> """" or onm <> """" or onw <> """")", DBCn, adOpenStatic, adLockReadOnly
  If rLP.BOF Then rLpgeht = 0
  If rLpgeht Then If IsNull(rLP!Langtext) Then rLpgeht = 0
  If rLpgeht Then If rLP!Langtext = "" Then rLpgeht = 0
  If rLpgeht Then
    Matr(3, j) = rLP!Langtext
  Else
   Dim rLangt As New ADODB.Recordset
   Set rLangt = Nothing
   rLangt.Open "select LangText from (" & laborAbfr & " where abkü = '" & raLW!Abkü & "' and einheit = '" & raLW!Einheit & "' and not isnull(langtext) and langtext <> """") as labor", DBCn, adOpenStatic, adLockReadOnly
   If rLangt.BOF Then
    Set rLangt = Nothing
    rLangt.Open "select LangText from (" & laborAbfr & " where abkü = '" & raLW!Abkü & "' and not isnull(langtext) and langtext <> """") as labor", DBCn, adOpenStatic, adLockReadOnly
   End If
   If Not rLangt.BOF Then
    Matr(3, j) = rLangt!Langtext
   End If
  End If
  If Not rLP.EOF Then
   If gschl = "m" Or (rLP!unw = "" And rLP!onw = "") Then
    Matr(5, j) = rLP!unm & "-" & rLP!onm
    Matr(1, j) = Replace(IIf(IsNull(rLP!unm), "", rLP!unm), "1:", "")
    Matr(2, j) = Replace(IIf(IsNull(rLP!onm), "", rLP!onm), "1:", "")
   Else
    Matr(5, j) = rLP!unw & "-" & rLP!onw
    Matr(1, j) = Replace(IIf(IsNull(rLP!unw), "", rLP!unw), "1:", "")
    Matr(2, j) = Replace(IIf(IsNull(rLP!onw), "", rLP!onw), "1:", "")
   End If
   If Matr(0, j) Like "LDL*" Then Matr(5, j) = "-100"
  End If
'  Debug.Print Matr(0, j) & "," & Matr(1, j) & "," & Matr(2, j)
  j = j + 1
  raLW.Move 1
 Loop
 
 syscmd 4, "Labor nach Normbereichseingabe"
 
' Stop
' dc.Application.Visible = True
 Wapp.ScreenUpdating = False
 Set Tabl = dc.Tables.Add(Range:=dc.Bookmarks!laborTab.Range, NumRows:=ZZ + 1, NumColumns:=rz + 3 - DiffBr, _
               DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitContent)
 syscmd 4, "Labor nach Tabellen-Addition"
 dc.Bookmarks.Add Name:="DMP", Range:=dc.Range(Tabl.Range.End, dc.Bookmarks!DMP.Range.End)
 syscmd 4, "Labor nach DMP-Markenlöschung"
 With Tabl
  If WappBuild > 9 Then
   If .Style <> "Tabellengitternetz" Then
     .Style = "Tabellengitternetz"
   End If
  .ApplyStyleHeadingRows = True
  .ApplyStyleLastRow = True
  .ApplyStyleFirstColumn = True
  .ApplyStyleLastColumn = True
  End If
  .Rows(1).Range.Orientation = 1
  .Rows(1).Range.Font.Bold = True
  syscmd 4, "Labor nach TabellenStil"
  For j = 1 To ZZ
   'If .Cell(j, 1) <> "" Then
   .cell(j, 1).Range.Font.Bold = True
  Next j
   .Rows.LeftIndent = CentimetersToPoints(-0.7)
 End With
 
 syscmd 4, "Labor nach Tabellenformatierung"
 
 Set raLW = Nothing
 raLW.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 Do While Not raLW.EOF
  i = 6
  Do While CDate(Matr(i, 0)) < DateValue(raLW!Zeitpunkt)
   i = i + 1
   If i >= rz + 5 Then Exit Do
  Loop
  j = 0
  If i > UBound(Matr, 1) Then
   i = UBound(Matr, 1)
   Stop
  End If
  Do While j <= UBound(Matr, 2)
   If Matr(0, j) = raLW!Abkü Then
    If Matr(4, j) = raLW!Einheit Then
     Exit Do
    End If
   End If
   j = j + 1
  Loop
  If j > UBound(Matr, 2) Then j = UBound(Matr, 2)
  Matr(i, j) = raLW!Wert
  If Not IsNull(raLW!Kommentar) And raLW!Kommentar <> "" Then
   If InStrB(raLW!Kommentar, "manuell") > 0 Then
     On Error Resume Next
     Tabl.cell(j + 1, i - 2).Range.Font.Italic = True
     On Error GoTo fehler
   ElseIf raLW!Wert = "" Then
'     Matr(i + Diffbr, j) = Matr(i, j) & IIf(Matr(i, j) = "", "", " ") & raLW!Kommentar
     Matr(i, j) = raLW!Kommentar
   End If
  End If
  pKz = "n" ' normal
  If IsNumeric(raLW!Wert) Or raLW!Wert Like "1:*" Or raLW!Wert Like "<*" Or raLW!Wert Like ">*" Then
   Dim rlWS$, rlWZ#, obkl%, obgr%
   Dim buch$, pos&
   obkl = 0
   obgr = 0
   rlWS = raLW!Wert
   For pos = 1 To Len(raLW!Wert)
    If InStrB("0123456789,.:-+<>", Mid$(rlWS, pos, 1)) = 0 Then
     rlWS = Left(rlWS, pos - 1) & Mid$(rlWS, pos + 1)
     pos = pos - 1
    End If
   Next pos
   If rlWS Like "1:<*" Then rlWS = "<1:" & Mid$(rlWS, 4)
   If Left(rlWS, 1) = "<" Then
    obkl = True
    rlWS = Mid$(rlWS, 2)
   ElseIf Left(rlWS, 1) = ">" Then
    obgr = True
    rlWS = Mid$(rlWS, 2)
   End If
   rlWZ = CDbl(Replace(Replace(Replace(Replace(rlWS, "^", ""), "1:", ""), ".", ","), ",,", ","))
   If Matr(1, j) <> "" Then
    If rlWZ < CDbl(Replace(Matr(1, j), ".", ",")) And Not obgr Then
     pKz = "p"
    End If
   End If
   If Matr(2, j) <> "" Then
    On Error Resume Next
    If rlWZ > CDbl(Replace(Matr(2, j), ".", ",")) And Not obkl Then
     pKz = "p"
    End If
    On Error GoTo fehler
   End If
'   If o <> -99 And rlwd > o Then pKz = "p"
  End If
  On Error Resume Next
  If pKz = "p" Then 'If Tabl.Cell(j + 1, i - 2) <> "" Then
   Tabl.cell(j + 1, i - 2).Range.Font.Bold = True
  End If
  On Error GoTo fehler
'  If IsNull(raLW!gruppe) Then
'   altgruppe = 0
'  Else
'   altgruppe = Val(raLW!gruppe)
'  End If
  raLW.MoveNext
 Loop
' Stop
' dc.Bookmarks!laborTab.Range
 'sql1 = "select * from (SELECT distinct(cdate(int(zeitpunkt))) as Datum from (" + sql + ") as sql left join [" + QMdbAkt + "].laborparameter on sql.abkü = laborparameter.abkü and iif(isnull(sql.einheit) or sql.einheit="""", ""kA"",sql.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar))) order by datum"
' Set rLW = Dtb.OpenRecordset(sql1, dbOpenDynaset)
 
syscmd 4, "Labor nach Dateneintragung "
 
' Exit Sub
 For i = 3 To rz + 5 - DiffBr
  For j = 0 To ZZ
   If i = 3 Then
    If Matr(3, j) = "HbA1c Eigenlabor" Then
     Matr(3, j) = "HbA1c"
    End If
   End If
   Tabl.cell(j + 1, i - 2) = IIf(Len(Matr(i + IIf(i < 6, 0, DiffBr), j)) < 11 Or i + DiffBr < 6, Matr(i + IIf(i < 6, 0, DiffBr), j), Left(Matr(i + IIf(i < 6, 0, DiffBr), j), 8) & "..")
  Next j
 Next
 
 For j = 2 To Tabl.Rows.Count
  For i = 1 To Tabl.Columns.Count
   If i <> 2 And Trim$(Replace(Replace(Tabl.cell(j, i).Range.Text, vbCr, ""), Chr$(7), "")) <> "" Then GoTo nr
  Next i
  Tabl.Rows(j).Delete
  j = j - 1
 Next j
nr:
' Stop
 ' ls = CStr(rZ) & vbcrlf
' raDat.MoveFirst
' Dim CNr%
' CNr = rZ + 3
' raDat.MoveLast
' Do While Not raDat.BOF
'  ' ls = LS + format$(radat!Datum, "dd.mm.yyyy") & vbcrlf
'  If lies.obmysql Then
'   Tabl.cell(1, CNr).Range = raDat!Datum 'BDTtoDate(mid$(raDat!Datum, 7, 2) & mid$(raDat!Datum, 5, 2) & left(raDat!Datum, 4))
'  Else
'   Tabl.cell(1, CNr).Range = format$(raDat!Datum, "dd.mm.yy")
'  End If
'  CNr = CNr - 1
'  raDat.Move -1
' Loop
 Tabl.cell(1, 1).Range = "Parameter"
 Tabl.cell(1, 2).Range = "Einheit"
 Tabl.cell(1, 3).Range = "Normbereich"
' Exit Sub
' ' ls = LS + CStr(zZ) & vbcrlf
' sql1 = "select * from (SELECT reihe, gruppe, unw, onw, unm, onm, sql1.*, datevalue(zeitpunkt) as Datum, sql1.abkü as sabkü, sql1.einheit as seinheit, sql1.langtext as slangtext from (" + sql + ") as sql1 left join laborparameter on sql1.abkü = laborparameter.abkü and iif(isnull(sql1.einheit) or sql1.einheit="""", ""kA"",sql1.einheit) = iif(isnull(laborparameter.einheit),"""",laborparameter.einheit) where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar))) as i0 order by gruppe,reihe,datum" ' dateserial(year(zeitpunkt),month(zeitpunkt),day(zeitpunkt))
' If lies.obmysql Then sql1 = Replace(Replace(sql1, "datevalue(", "date("), "iif(", "if(")
' Set raLW = Nothing
'' Set rLW = Dtb.OpenRecordset(sql1)
' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
' Dim aktZ%, VorWert$, AktGru$, AktSp%, AktAbkü$, obgeändert%
' Do While Not raLW.EOF
'  If Not IsNull(raLW!sabkü) Then
'   If IsNull(raLW!gruppe) Then
''    If rLaU Is Nothing Then Set rLaU = Dtb.OpenRecordset("select * from (" + sql + ") as sql order by zeitpunkt")
'    Set raLau = Nothing
'    raLau.Open "select *,sql.abkü as sabkü, sql.langtext as slangtext from (" & sql & ") as sql order by zeitpunkt", dbcn, adOpenDynamic, adLockReadOnly
'    Debug.Print "Kein Parametereintrag für: " + CStr(raLW!sabkü) & " " & CStr(IIf(IsNull(raLW!slangtext), "", raLW!slangtext)) + " [" + IIf(IsNull(raLW!seinheit), "", raLW!seinheit) + "]"
'    Call raLau.Find("Pat_id = " & CStr(raLW!Pat_id) & " and zeitpunkt = cdate(""" & CStr(raLW!Zeitpunkt) & """) and fertigstgrad = """ & CStr(raLW!FertigStGrad) & """ and abkü = """ & CStr(raLW!sabkü) & """", 0, adSearchForward, 1)
    
'    If Not rLaU.BOF Then
'     Call LaborParameter(raLau)
'     obgeändert = True
'    End If
'   End If
'  End If
'  raLW.Move 1
' Loop
' If obgeändert Then
'  Set raLW = Nothing
''  Set raLW = Dtb.OpenRecordset(sql)
'  raLW.Open sql, dbcn, adOpenDynamic, adLockReadOnly
' Else
' raLW.MoveFirst
' End If
' aktZ = 1
'  Dim TestFeld$
'  Dim TestInh
'  On Error Resume Next
'  Err.Clear
'  TestFeld = "sql."
'  TestInh = raLW.Fields(TestFeld + "abkü")
'  If Err.Number <> 0 Then
'   TestFeld = ""
'   Err.Clear
'  End If
'  On Error GoTo fehler
' Do While Not raLW.EOF
'  If Not IsNull(raLW.Fields(TestFeld + "abkü")) Then
''   If Val(ralw!Gruppe) <> altGruppe Then
'    ' ls = LS & vbcrlf
''   End If
'   If IsNull(VorWert) Or VorWert = "" Or raLW.Fields(TestFeld + "LangText") <> VorWert Then
'    aktZ = aktZ + 1
'    AktSp = 4
'    On Error GoTo gruppenfehler
'    If Not IsNull(AktGru) And AktGru <> "" And raLW.Fields("Gruppe") <> AktGru Then
'    On Error GoTo fehler
'      aktZ = aktZ + 1
'    End If
'    Do While aktZ > Tabl.Rows.Count
'     Tabl.Rows.Add
'     zZ = zZ + 1
'    Loop
'    AktGru = IIf(IsNull(raLW.Fields("Gruppe")), "", raLW.Fields("Gruppe"))
'    Tabl.cell(aktZ, 1).Range = raLW.Fields("slangtext") 'IIf(IsNull(raLW.Fields("sql.LangText")), "", raLW.Fields("sql.LangText"))
'    Tabl.cell(aktZ, 2).Range = raLW.Fields("seinheit") 'IIf(IsNull(raLW.Fields("sql.Einheit")), "", raLW.Fields("sql.Einheit"))
'    u = -99
'    o = -99
'    uNG = ""
'    oNG = ""
'    unw = IIf(IsNull(raLW!unw), "", raLW!unw)
'    onw = IIf(IsNull(raLW!onw), "", raLW!onw)
'    unm = IIf(IsNull(raLW!unm), "", raLW!unm)
'    onm = IIf(IsNull(raLW!onm), "", raLW!onm)
''    If instrb(raLW.Fields("sql.Langtext"), "LDL") > 0 Then
'    If instrb(raLW.Fields("slangtext"), "LDL") > 0 Then
'     onm = "100"
'    End If
'    If gschl = "w" Then
'     uNG = Replace(unw, "1:", "")
'     oNG = Replace(onw, "1:", "")
'    End If
'    If uNG = "" And oNG = "" Then
'     uNG = Replace(unm, "1:", "")
'     oNG = Replace(onm, "1:", "")
'    End If
'    If uNG <> "" Then u = CDbl(uNG)
'    If oNG <> "" Then o = CDbl(oNG)
''   If IsNumeric(raLw!Wert) Or raLw!Wert Like "1:*" Then
''    If gschl = "w" Then
''     If Not ralw!unw = "" Then u = CDbl(Replace(ralw!unw, "1:", ""))
''     If Not ralw!onw = "" Then o = CDbl(Replace(ralw!onw, "1:", ""))
''    End If
''    If gschl = "m" Or u = -99 Then
''     If Not ralw!unm = "" Then u = CDbl(Replace(ralw!unm, "1:", ""))
''    End If
''    If gschl = "m" Or o = -99 Then
''     If Not ralw!onm = "" Then o = CDbl(Replace(ralw!onm, "1:", ""))
''    End If
''   End If
'    ' ls = LS + pKz & vbcrlf
'    NB = ""
'    If uNG <> "" Or oNG <> "" Then
'     If gschl = "w" And (unw <> "" Or onw <> "") Then
'      NB = unw + "-" + onw
'     End If
'    End If
'    If NB = "" And (unm <> "" Or onm <> "") Then
'     NB = unm + "-" + onm
'    End If
'    Tabl.cell(aktZ, 3) = NB
'    ' ls = LS + NB & vbcrlf
'   End If
''   VorWert = IIf(IsNull(raLW.Fields("sql.LangText")), "", raLW.Fields("sql.LangText"))
'   VorWert = raLW.Fields("slangtext") 'IIf(IsNull(raLW.Fields("sql.LangText")), "", raLW.Fields("sql.LangText"))
'   ' ls = LS + ralw.Fields("sql.LangText") & vbcrlf
'   ' ls = LS + ralw.Fields("sql.Einheit") & vbcrlf
'   ' ls = LS + format$(ralw!Datum, "dd.mm.yyyy") & vbcrlf
'   ' ls = LS + ralw!Wert & vbcrlf
   
'   Do
'    If DateValue(raLW!Datum) > CDate(left(Tabl.cell(1, AktSp).Range, 8)) Then
'     AktSp = AktSp + 1
'    Else
'     Exit Do
'    End If
'    If AktSp > rZ + 3 Then
'     Exit Do
'    End If
'   Loop
   
'   Tabl.cell(aktZ, AktSp).Range = IIf(IsNull(raLW!Wert), IIf(IsNull(raLW!Kommentar), "", format$(raLW!Zeitpunkt, "dd/mm/yy") + ": " + raLW!Kommentar), Replace(Replace(trim$(IIf(IsNull(raLW!Wert), "", raLW!Wert)), vbcr, ""), vbtab, ""))
'   If instrb(raLW!Kommentar, "manuell") > 0 Then Tabl.cell(aktZ, AktSp).Range.Font.Italic = True
   
'  End If ' not isnull(ralw![sql!abkü]
'  raLW.Move 1
' Loop '  While Not ralw.EOF
 
' Open "c:\wwtest.txt" For Output As #34
' Print #34, LS
' Close #34
 syscmd 4, "Labor vor Endformatierung"

 With Tabl
  Select Case rz
   Case 1, 2, 3, 4, 5, 6
     .Range.Font.Size = 9
'     .Columns(1).Width =
     .Columns(2).Width = 53
     .Columns(3).Width = 44
   Case 7, 8, 9, 10, 11
     .Range.Font.Size = 8
     .Columns(1).Width = 96.6
     .Columns(2).Width = 49
     .Columns(3).Width = 41
   Case 12, 13, 14
     .Range.Font.Size = 7
     .Columns(1).Width = 85.6
     .Columns(2).Width = 43.5
     .Columns(3).Width = 36.5
   Case Else
     .Range.Font.Size = 6
     .Columns(1).Width = 73
     .Columns(2).Width = 38
     .Columns(3).Width = 33
  End Select
' Weiten:
' 6: 73   / 38   / 33
' 7: 85,6 / 43,5 /36,5
' 8: 96,6 / 49 / 41
' 9:      / 53 / 44
 syscmd 4, "Labor nach Endformatierung"
 raDat.Close
 raLW.Close
 End With
 
 #If False Then
  Dim adn$, dcn$, Tn$
  adn = Wapp.ActiveDocument.Name
  dcn = dc.Name
  Tabl.id = "Labor"
  Set Wapp = Nothing
  GetWord
  Set dc = Wapp.Documents(dcn)
  Dim tabli
  For Each tabli In dc.Tables
   If tabli.id = "Labor" Then
    Set Tabl = tabli
    Exit For
   End If
  Next tabli
 #End If

syscmd 4, "Labor vor LaborTabAnp"
 Call LaborTabAnp(Wapp, dc)
syscmd 4, "Labor nach LaborTabAnp"
 
 Exit Sub
gruppenfehler:
 If Err.Number = 3265 Then ' Gruppe fehlt in Auflistung, 18.4.06
  Stop
  Resume
  Resume Next
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborIns/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub      ' LaborIns
Sub LaborTabAnp(Wapp As Object, dc As Object) 'Word.Application, dc As Word.Document)
 Dim i%, j%, k%, tz%, wi#, wineu!, mk%, buch$
 On Error GoTo fehler
 Wapp.ScreenUpdating = False
 Dim FontSize%
 With dc.Tables(dc.Tables.Count)
 FontSize = .cell(1, 1).Range.Font.Size
  .AllowAutoFit = False
  For i = 1 To .Columns.Count
   wi = 0
   mk = 0
   For j = 2 To .Rows.Count
'    wineu = (.Cell(j, i).Range.Characters.Count)
'    For k = wineu To 0 Step -1
'     tz = Asc(mid$(.Cell(j, i).Range, k, 1))
'     If tz < 32 Then
'      wineu = wineu - 1
'     Else
'      Exit For
'     End If
'    Next
'   wineu = wineu+1 ' das Viereck (chr$(7)) am Schluß
' schneller, wenngleich nicht ganz so robust:
    wineu = .cell(j, i).Range.Characters.Count
    wineu = FontSize * wineu
    If wineu > wi Then
     If i > 3 Then
      If InStrB(.cell(j, i).Range, ",") = 0 And InStrB(.cell(j, i).Range, ".") = 0 Then
        mk = 0
      Else
        mk = -1
      End If
     End If
     wi = wineu
 '    Debug.Print "Längenrekord Spalte ", i, ":", .Cell(j, i).Range.Characters.Count, "(" + .Cell(j, i).Range + ")"
    ElseIf wineu = wi And mk = -1 Then
      If InStrB(.cell(j, i).Range, ",") = 0 And InStrB(.cell(j, i).Range, ".") = 0 Then
       mk = 0
      End If
    End If
   Next j
   Debug.Print "Spalte: ", i, wi
   Select Case i
    Case 1: wi = wi * 0.52  ' 0.5 * (3.41 / 3.21) ' paßt
    Case 2: wi = wi * 0.59  ' 0.5 * (1.74 / 1.36) ' paßt
    Case 3: wi = wi * 0.53  ' 0.5 * (1.47 / 1.23) ' paßt
    Case Is > 3
      If mk = -1 Then
       wi = 0# + ((wi - 1) * 0.81)
      Else
       wi = wi * 0.81 '1 * 0.5 + (wi - 1) * 1
      End If
   End Select
   If i > 3 Then
    If wi > 40 Then wi = 40
   End If
   Select Case i
'    Case 1: wi = 82.5
'    Case 2: wi = 32.55
'    Case 3: wi = 44.2
'    Case Else: wi = 30
   End Select
'   .Columns.AutoFit
   If wi < 7.2 Then wi = 7.2
   .Columns(i).Width = wi
'   For j = 1 To .Rows.Count
'    .Cell(j, i).Range.FitTextWidth = 0
'   Next j
  Next
  .Rows.LeftIndent = CentimetersToPoints(0)
  .Columns.AutoFit
  .Rows.LeftIndent = CentimetersToPoints(0)
End With
 Wapp.ScreenUpdating = True
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborTabAnp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' LaborTabAnp

Sub LaborTabAnp_alt(dc)
 Dim Diff%, erl%, aend%, i%, d1, d2
 Dim ColuNr%
 On Error GoTo fehler
 With Tabl
  For ColuNr = 1 To .Columns.Count
'   Debug.Print "vorher:  " + CStr(ColuNr) & " " & CStr(.Columns(ColuNr).Width)
  Next ColuNr
'  d1 = Now + 2 / 60 / 60 / 24
'  Do While d2 < d1
'   d2 = Now
'  Loop
' Das folgende funktioniert leider nicht im Gegensatz zum Obigen
  Tabl.Select
  Do While .Columns(.Columns.Count - 1).Width = .Columns(.Columns.Count).Width And .Columns(.Columns.Count - 1).Width = .Columns(1).Width
  Loop
  For ColuNr = 1 To .Columns.Count
   Debug.Print "vorher2: " + CStr(ColuNr) & " " & CStr(.Columns(ColuNr).Width)
  Next ColuNr
  For ColuNr = 4 To .Columns.Count
   If .Columns(ColuNr).Width > 80 Then GoTo ender
  Next ColuNr
'  Stop
  For ColuNr = 1 To .Columns.Count
   Debug.Print "nachher: " + CStr(ColuNr) & " " & CStr(.Columns(ColuNr).Width)
  Next ColuNr
  GoTo nichtender
ender:
  For i = 2 To 3
   erl = .Columns(ColuNr).Width - 80
   aend = IIf(i = 2, 58, 56) - .Columns(i).Width
   If erl < aend Then aend = erl
   .Columns(i).Width = .Columns(i).Width + aend
   .Columns(ColuNr).Width = .Columns(ColuNr).Width - aend
  Next
 End With
nichtender:
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborTabAnp_alt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' LaborTabAnp

Sub Tabelle()
Dim i%, zZStr$, ZZ%, rZStr$, rz%, Diff%, ct$ ' Container
Dim aktZ% ' aktuelle Zeile
Dim Par$, Einh$, Datu$, Wert$, path$, NB$, altPar$
Dim F As File
With Wapp.ActiveDocument
 If .Tables.Count > 0 Then
  With .Tables(.Tables.Count)
    Stop
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    Set F = FSO.GetFile(aVerz + "\Labor.txt") ' u:\Anamnese
    If Not F Is Nothing Then
      Close #16
      Open F.path For Input As #16
      Line Input #16, rZStr
      rz = Val(rZStr)
      Diff = rz - .Columns.Count + 3
      If Diff > 0 Then
        For i = 1 To Diff
         .Columns.Add
        Next
      ElseIf Diff < 0 Then
       For i = 1 To -Diff
        .Columns(.Columns.Count).Delete
       Next
      End If
      For i = 1 To rz ' Datum
       Line Input #16, ct
       .cell(1, 3 + i).Range = Left(ct, 10) ' mit ganzem Jahr
      Next
      Line Input #16, zZStr
      ZZ = Val(zZStr)
      Diff = ZZ - .Rows.Count
      If Diff > 0 Then
       For i = 1 To Diff
        .Rows.Add
       Next
      ElseIf Diff < 0 Then
       For i = 1 To -Diff
        .Rows(.Rows.Count).Delete
       Next
      End If
      With .Range
        .Font.Bold = False
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0) ' wdfontblack
      End With
      altPar = ""
      Line Input #16, ct ' obligatorische Leerzeile
      aktZ = 1
      Do While Not EOF(16)
       Line Input #16, Par
       If Par = "" Then ' neue Gruppe
        aktZ = aktZ + 1
        For i = 1 To .Columns.Count
         .cell(aktZ, i).Range = ""
        Next i
        Line Input #16, Par
       End If
       If Par <> altPar Then
        aktZ = aktZ + 1
        For i = 4 To .Columns.Count
         .cell(aktZ, i).Range = ""
        Next i
        altPar = Par
       End If
       .cell(aktZ, 1).Range = Par
       Line Input #16, Einh
       .cell(aktZ, 2).Range = Einh
       Line Input #16, Datu
       For i = 4 To .Columns.Count
        If Left(.cell(1, i).Range, 10) = Datu Then
         Exit For
        End If
       Next i
       Line Input #16, Wert
       .cell(aktZ, i).Range = Wert
       Line Input #16, path
       If path = "p" Then
        .cell(aktZ, i).Range.Font.Bold = True
        .cell(aktZ, i).Range.Font.Italic = True
        .cell(aktZ, i).Range.Font.Color = wdColorRed
       Else
       End If
       Line Input #16, NB
       .cell(aktZ, 3).Range = NB
      Loop
      Close #16
    End If ' not f is nothing
    
'    Exit Sub
  
        .Columns(1).Width = 92
        .Columns(2).Width = 47
        .Columns(3).Width = 42
        For i = 4 To .Columns.Count
         .cell(1, i).Range = Format$(Left(.cell(1, i).Range, 10), "dd.mm.yyyy")
         .Columns(i).Width = 30
        Next i
        For i = 2 To .Rows.Count
         If .cell(i, 3).Range.Paragraphs.Count > 2 Then
          .cell(i, 3).Range = ""
         End If
        Next i
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderHorizontal)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderVertical)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    With .Range.Find
        .clearformatting
        .Replacement.clearformatting
        .Text = "Eiwei¯"
        .Replacement.Text = "Eiweiß"
        .Execute Replace:=wdReplaceAll
        .clearformatting
        .Replacement.clearformatting
        .Text = "Á"
        .Replacement.Text = "µ"
        .Execute Replace:=wdReplaceAll
        .clearformatting
        .Font.Color = wdColorRed
        .Text = ""
        With .Replacement
          .clearformatting
          .Font.Bold = True
          .Font.Italic = True
          .Text = ""
        End With
       .Execute Format:=True, Replace:=wdReplaceAll
    End With ' .range.find
  .Rows(1).Range.Font.Bold = True
  For i = 1 To .Rows.Count
   .cell(i, 1).Range.Font.Bold = True
  Next i
'  With .Columns(1)
'   .Range.Font.Bold = True
''   .Select
''   Selection.Font.Bold = True
''   ActiveDocument.Range(Start:=.Cells(1).Range.Start, End:=.Cells(.Cells.Count).Range.End).Font.Bold = True
 ' End With
'  If 1 = 0 Then
'   Dim wrkJet As Workspace, Adr As Database, rLab As DAO.Recordset, rNam As DAO.Recordset
'   Dim fehler ' wg. Kompilierfehler bei untenstehender Zeile trotz Einbzug aller Verweise wie in winword
   ' Public fehler As New C_Fehlerbehandlung
'   If fehler.Ist_TMAktiv Then
'    Set wrkJet = CreateWorkspace$("", "Admin", "", dbUseJet)
'    Set Adr = wrkJet.OpenDatabase("u:\Anamneseblatt.mdb", False, False)
'    Set rNam = Adr.OpenRecordset("Namen", dbOpenTable)
'    rNam.index = "Auswahl"
'    Set TMInterface.TMTxt.Daten.AktivesDokument = ActiveDocument
'    rNam.Seek "=", TMInterface.TMTxt.Daten.AktuellerPatient.Namensdaten.Nachname, TMInterface.TMTxt.Daten.AktuellerPatient.Namensdaten.Vorname, TMInterface.TMTxt.Daten.AktuellerPatient.Geburtsdaten.Datum
'    If Not rNam.NoMatch Then
'     Set rLab = Adr.OpenRecordset("select * from [" + QMdbAkt + "].Labor where pat_id = " + CStr(rNam!Pat_id) + ";", dbOpenDynaset)
'     Do While Not rLab.EOF
''     If Not IsNumeric(rLab!Wert) And Not IsNull(rLab!Wert) And rLab!Wert <> "" Then
'       Debug.Print rLab!LangText & " " & CStr(rLab!Wert) & " " & rLab!Einheit
''     End If
'      rLab.Move 1
'     Loop
'    End If
'   End If
'  End If
  End With ' .tables
  End If
 End With
 With Wapp.Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
 End With

 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Tabelle/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Tabelle
'Function Verlauf$(Pat_id&)
' Verlauf = eintraege(Pat_id, """notiz"",""telef"",""med""")
'End Function ' Verlauf$(Pat_id&)
Function eintraege$(Pat_id&, Krit$, Optional vordat As Date)
#Const obAlte = True
 Dim raVL As New ADODB.Recordset, lzp As Date, aktDat As Date
 On Error GoTo fehler
' Call dtbInit
 Select Case LCase(Krit)
  Case """rr"""
'   Set raVL = Dtb.OpenRecordset("select zeitpunkt,rr as inhalt from [" + QMdbAkt + "].rr where pat_id = " & CStr(Pat_id) + " order by zeitpunkt")
   raVL.Open "select zeitpunkt,rr as inhalt from rr where pat_id = " & CStr(Pat_id) + " order by zeitpunkt", DBCn, adOpenDynamic, adLockReadOnly
  Case Else
'   Set raVL = Dtb.OpenRecordset("select zeitpunkt,inhalt,art from [" + QMdbAkt + "].eintraege where pat_id = " & CStr(Pat_id) + " and art in (" & Krit & ") order by zeitpunkt")
#If Not obAlte Then
   raVL.Open "select zeitpunkt,inhalt,art from eintraege where pat_id = " & CStr(Pat_id) + " and art in (" & Krit & ") and zeitpunkt > " & datform(vordat) & " order by zeitpunkt", DBCn, adOpenDynamic, adLockReadOnly
#Else
   raVL.Open "select zeitpunkt,inhalt,art from eintraege where pat_id = " & CStr(Pat_id) + " and art in (" & Krit & ") order by zeitpunkt", DBCn, adOpenDynamic, adLockReadOnly
#End If
 End Select
 Do While Not raVL.EOF
  If raVL!Inhalt <> "Blutentnahme" And raVL!Inhalt <> "Blutabnahme" Then
   aktDat = DateValue(raVL!Zeitpunkt)
   If aktDat <> lzp Then
    eintraege = eintraege + Format$(aktDat, "DD.MM.YY")
   End If
   If LCase(Krit) <> """rr""" Then
    If raVL!art = "usdm" Then
     eintraege = eintraege + vbTab + Replace(Replace(raVL!Inhalt, "^", "") & vbCrLf, "aktuellen Blutdruck und ggf. Puls bitte extra eingeben", "")
     GoTo hamsscho
    End If
   End If
   eintraege = eintraege + vbTab + Replace(raVL!Inhalt, "^", "") & vbCrLf
hamsscho:
   lzp = DateValue(raVL!Zeitpunkt)
  End If
  raVL.Move 1
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraege/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' eintraege$(Pat_id&, Krit$)

'Function haakt() ' Hauärzte aus dem Internet ergänzen (Knopf)
' Dim rHa As DAO.Recordset, geht%
' On Error GoTo fehler
' Set rHa = TabÖff("Hausaerzte", "KVNr")
' Do While Not rHa.EOF
''  If rHA!Nachname = "Licht" Or geht Then
'   Call do_haakt(rHa)
''   geht = -1
''  End If
'  rHa.Move 1
' Loop
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in haakt/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): End
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function ' haakt()
#If mitab Then
Function do_ha_click(ByVal HANr)
#If False Then
 Dim stelle%
 Dim rHa As dao.Recordset
 If KVÄDatei = "" Then Call KVÄDateifind
 Set kvä = DBEngine.OpenDatabase(KVÄDatei, , True)
 Set rHa = kvä.OpenRecordset("HAe", dbOpenTable)
 rHa.Index = "KVNr"
' Unterprogramm
 Dim teile$()
 teile = Split(HANr, "_")
 If UBound(teile) > -1 Then
  If InStrB(teile(0), "/") = 0 Then
   teile(0) = Left(teile(0), 2) + "/" + Mid$(teile(0), 3) 'Replace(teile(0), "/", "")
  End If
 End If
 If UBound(teile) = 2 Then rHa.Seek "=", teile(0), Trim$(teile(1)), Trim$(teile(2))
 If UBound(teile) = 1 Or rHa.NoMatch Then rHa.Seek "=", teile(0), Trim$(teile(1))
 If UBound(teile) = 0 Or rHa.NoMatch Then rHa.Seek "=", teile(0)
 Dim stAppName$, dbnr$
 If Not rHa.NoMatch Then dbnr = rHa!dbnr
 If dbnr <> "" And Not IsNull(dbnr) Then
  stAppName = Environ("programfiles") + "\internet explorer\iexplore.exe http://www.kvb.de/servlet/PB/cmd/arztsuche3/index.html?&enr=" + CStr(dbnr)
  Call Shell(stAppName, vbNormalFocus)
 End If
 kvä.Close
#End If
End Function
#End If
'Function KVÄDateifind$()
' Dim Fil As File
' Dim FGef As File
' Dim db As DAO.Database
' Dim rs As DAO.Recordset, Sta As DAO.Recordset
' Dim altStD#
' On Error GoTo fehler
' If FSO Is Nothing Then Set FSO = New FileSystemObject
' For Each Fil In FSO.GetFolder(aVerz).Files
'  If Fil.Name Like "KV*.mdb" Then
'   On Error Resume Next
'   Set db = DAO.OpenDatabase(Fil.path)
'   Set rs = db.OpenRecordset("HAE")
'   rs.MoveLast
'   rs.MoveFirst
'   If rs.RecordCount > 14000 Then
'    Set Sta = db.OpenRecordset("Select * from Stand order by stand desc")
'    Sta.Move 30
'    If altStD = 0 Or (altStD > 0 And Not Sta.EOF And Not Sta.BOF And Sta!datum > altStD) Then
'     altStD = Sta!datum
'     Set FGef = Fil
'    End If
'   End If
'   On Error GoTo fehler
'  End If
' Next Fil
' KVÄDatei = FGef.path
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = Currentdb.Name
'#Else
' AnwPfad = App.path
'#End If
'Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in KVÄDateifind/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): End
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function ' KVÄDateiFind
Function do_haakt(rHa As ADODB.Recordset)
 Dim ars As New ADODB.Recordset, aRsx As New ADODB.Recordset
' Call KVÄVorb
 Call acon(HaT)
 ars.Open "select * from hae where kvnr = '" & Replace(rHa!KVNr, "/", "") & "' and nachname = '" & rHa!NachName & "' and vorname = '" & rHa!Vorname & "'", HAECn, adOpenStatic, adLockReadOnly
 If ars.BOF Then
  ars.Close
  ars.Open "select * from hae where kvnr = '" & Replace(rHa!KVNr, "/", "") & "' and nachname = '" & rHa!NachName & "'", HAECn, adOpenStatic, adLockReadOnly
  If ars.BOF Then
   ars.Open "select * from hae where nachname = '" & rHa!NachName & "' and vorname = '" & rHa!Vorname & "' and ort = '" & rHa!Ort & "'", HAECn, adOpenStatic, adLockReadOnly
  End If
 End If
 If ars.BOF Then
  MsgBox "Vermutlich veraltete KV-Nummer in Turbomed: " & rHa!KVNr & " " & rHa!NachName & " " & rHa!Vorname
  Stop
 Else
  Stop
'     rHa.Edit
     rHa!Name = IIf(IsNull(ars!Anrede), "", ars!Anrede + " ") + IIf(IsNull(ars!Titel), "", ars!Titel + " ") + IIf(IsNull(ars!Vorname), "", ars!Vorname + " ") + IIf(IsNull(ars!NachName), "", ars!NachName)
     rHa!Anschrift = ars!Straße & ", " & ars!Plz & " " & ars!Ort
     rHa!Telefon = ars!tel1
     If Not IsNull(ars!fax1) And ars!fax1 <> "" Then
      rHa!Telefax = ars!fax1
     Else
      aRsx.Open "select * from haxls where name = '" & ars!NachName & "' and vorname = '" & ars!Vorname & "' and kvnr = '" & Replace(ars!KVNr, "/", "") & "'", HAECn, adOpenStatic, adLockReadOnly
      If aRsx.BOF Then
       aRsx.Open "select * from haxls where name = '" & ars!NachName & "' and vorname = '" & ars!Vorname & "'", HAECn, adOpenStatic, adLockReadOnly
      End If
      If aRsx.BOF Then
       MsgBox "Vermutlich Schreibfehler in Turbomedtabelle bei: " & ars!KVNr & " " & ars!NachName & " " & ars!Vorname
      Else
       If Not IsNull(aRsx!Fax) And ars!Fax <> "" Then
        rHa!Telefax = ars!Fax
       End If
      End If
     End If
     If Not (IsNull(ars!Email) Or ars!Email = "") Then rHa!E_Mail = ars!Email
     rHa!Zulassungsgebiet = ars!zulg
     rHa!Arzttyp = ars!Arzttyp
     rHa!dmpt2 = ars!dmpt2
     rHa!Beme = ars!Beme
     If IsNull(rHa!Vorname) Or rHa!Vorname = "" Then rHa!Vorname = ars!Vorname
     If IsNull(rHa!NachName) Or rHa!NachName = "" Then rHa!NachName = ars!NachName
     rHa!Titel = ars!Titel
     rHa!Geschlecht = ars!Geschlecht
     rHa!Ort = ars!Ort
     rHa!Plz = ars!Plz
     rHa!Straße = ars!Straße
     rHa!KVNr = Left(Replace(ars!KVNr, "/", ""), 7)
     rHa.Update
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_haakt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_haakt(rHa as dao.recordset)

Public Function PiepKurz()
 On Error GoTo fehler
 Call WD
 Call Sound(WinDir + "\media\Windows XP-Hinweis.wav")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PiepKurz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PiepKurz()
Public Function Piep()
 On Error GoTo fehler
 Call WD
 Call Sound(WinDir + "\media\ringout.wav")
 sleep 1000
 Call Sound(WinDir + "\media\recycle.wav")
 sleep 1000
 Call Sound(WinDir + "\media\ringout.wav")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Piep/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Piep()

Function GetZA(dbnr&)
         Dim hClipMemory&
         Dim lpClipMemory&
         Dim MyString$
         Dim RetVal&
 On Error GoTo fehler
         If OpenClipboard(0&) = 0 Then
            Call PiepKurz
            Call MsgBox("Konnte Zwischenspeicher nicht öffnen. Könnte besetzt sein", , Titel)
            Exit Function
         End If

         ' Obtain the handle to the global memory
         ' block that is referencing the text.
         hClipMemory = GetClipboardData(CF_TEXT)
         If IsNull(hClipMemory) Then
            Call PiepKurz
            Call MsgBox("Konnte keinen Speicher reservieren", , Titel)
            GoTo OutOfHere
         End If

         ' Lock Clipboard memory so we can reference
         ' the actual data string.
         lpClipMemory = GlobalLock(hClipMemory)

         If Not IsNull(lpClipMemory) Then
            MyString = Space$(MAXSIZE)
            RetVal = lstrcpy(MyString, lpClipMemory)
            RetVal = GlobalUnlock(hClipMemory)

            ' Peel off the null terminating character.
            If InStrB(MyString, vbNullChar) > 0 Then
             MyString = Mid$(MyString, 1, InStr(1, MyString, vbNullChar, 0) - 1)
            End If
         Else
            Call PiepKurz
            Call MsgBox("Konnte Speicher nicht reservieren, um den String zu kopieren", , Titel)
         End If

OutOfHere:

         RetVal = CloseClipboard()
         GetZA = MyString
         
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Call PiepKurz
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + " bei DS Nr. " + CStr(IIf(IsNull(dbnr), "", dbnr)) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetZA/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetZA

'schneidet Neue-Zeile-Zeichen ab
Function z$(rHa, FeName$, ByVal q, dbnr&)
 On Error GoTo fehler
 Select Case rHa.Fields(FeName).Type
  Case 10
   z = Left(q, rHa.Fields(FeName).Size)
  Case Else ' 12 = Memo
   z = q
 End Select
 
  GoTo schluss:
schluss:
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Call PiepKurz
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + " bei DS Nr. " + CStr(dbnr) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Z(neueZeileZeichen abschneiden)/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Z$

Function snie()
  Call WD
  Call Sound(WinDir + "\media\Windows XP-Batterie niedrig.wav")
End Function
Function WD()
 On Error GoTo fehler
 WinDir = Space$(144)
 Call GetWindowsDirectory(WinDir, 144)
 WinDir = Replace(Trim$(WinDir), vbNullChar, "")
 WD = WinDir
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WD/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' WD
#If False Then
Public Function testdb()
 Dim anam As dao.Recordset
 On Error GoTo fehler
 Set testdb = OpenDatabase(QMdbAkt)
 Set anam = testdb.OpenRecordset("Anamnesebogen", dbOpenTable)
 anam.Index = "Pat_id"
 anam.Seek "=", 97
 anam.Edit
 anam.Update
 testdb.Close
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in testdb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' testdb
#End If
Function Datenbankkontrolle()
 Dim obGleich%, pText$
' Dim haz As DAO.Recordset, hae As DAO.Recordset, hae1 As DAO.Recordset
 Dim haz As ADODB.Recordset, hae As ADODB.Recordset, hae1 As ADODB.Recordset
' Dim kvä As DAO.Database
 On Error GoTo fehler
' Call dtbInit
 ' if kvä is nothing then
' If Not lies.obmysql And KVÄDatei = "" Then Call KVAccSuch
' Call KVÄVorb
  Call acon(HaT)
' Set kvä = DBEngine.OpenDatabase(KVÄDatei, , True)
 kvä.Open "hae", HAECn, adOpenDynamic, adLockReadOnly
 Open aVerz + "\HAFehler.txt" For Output As #32 ' u:\Anamnese
' Set haz = TabÖff("Hausaerzte")
 haz.Open "hausaerzte", DBCn, adOpenDynamic, adLockReadOnly
' Set hae = kvä.OpenRecordset("Hae", dbOpenTable)
 hae.Open "hae", HAECn, adOpenDynamic, adLockReadOnly
' Set hae1 = kvä.OpenRecordset("Hae", dbOpenDynaset)
 hae1.Open "hae", HAECn, adOpenDynamic, adLockReadOnly
' hae.index = "KVNr"
 Do While Not haz.EOF
'  hae.Seek "=", left(haz!KVNr, 2) + "/" + right$(haz!KVNr, 5) ' "kvnr = " +
  Set hae = Nothing
  hae.Open "select * from hae where kvnr = " & haz!kvnu
  If hae.BOF Then
   pText = CStr(haz!id) + ": " + CStr(IIf(IsNull(haz!KVNr), "KV-Nr: (Null)", haz!KVNr)) & " " & IIf(IsNull(haz!NachName), "Nachname: (Null)", haz!NachName) & " " & IIf(IsNull(haz!Vorname), "Vorname: (Null)", haz!Vorname) + " in HAE nicht gefunden"
'   hae1.FindFirst "instr(HAName, """ + haz!Nachname + """) > 0 and instr(HAName, """ + haz!Vorname + """) > 0 and ort = """ + haz!Ort + """"
   Set hae1 = Nothing
   hae1.Open "select * from hae where haname like '%" & haz!NachName & "%' and haname like '%" & haz!Vorname & "%' and ort like '%" & haz!Ort & "%'", HAECn, adOpenStatic, adLockReadOnly
   If Not hae1.EOF Then
    pText = pText + ",        dort aber " + haz!NachName & ", " & haz!Vorname + " in " + haz!Ort + " mit KV-Nr: " + hae1!KVNr + "!"
   End If
   Print #32, pText
  Else
   obGleich = 0
   Do
    If hae!NachName = haz!NachName And hae!Vorname = haz!Vorname Then
     obGleich = -1
     Exit Do
    End If
    hae.Move 1
    If hae.EOF Then Exit Do
    If hae!KVNr <> Replace(haz!KVNr, "/", "") Then Exit Do
   Loop
   If Not obGleich Then
    Print #32, CStr(haz!id) + ": " + CStr(IIf(IsNull(haz!KVNr), "KV-Nr: (Null)", haz!KVNr)) & " " & IIf(IsNull(haz!NachName), "Nachname: (Null)", haz!NachName) & " " & IIf(IsNull(haz!Vorname), "Vorname: (Null)", haz!Vorname) + " in HAE nicht mit gleichem Namen nicht gefunden"
   End If
  End If
  haz.Move 1
 Loop
 Close #32
 kvä.Close
' Dtb.Close
 MsgBox "Fertig mit Datenbankkontrolle!"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Datenbankkontrolle/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Datenbankkontrolle

Function AlleBriefe()
' Dim q As DAO.QueryDef
 Call do_Click_Vorbereit
 sql = "select anamnesebogen.nachname +"",""+anamnesebogen.vorname as Nachname, cstr(anamnesebogen.diabetestyp) as vorname, briefe.pfad as DokPfad, briefe.name as DokName, briefe.art as DokArt, briefe.* from [" + QMdbAkt + "].Briefe left join [" + QMdbAkt + "].anamnesebogen on briefe.pat_id = anamnesebogen.pat_id where briefe.name like ""Brief an*"" and not cstr(diabetestyp) in (""1"", ""2"",""g"",""s"") order by zeitpunkt desc"
 Call DtbCreateQueryDef("LaborDokumente eP", sql)
 DoCmd.OpenForm "LaborDokumente eP", acNormal, , "17=17"
 Forms![LaborDokumente ep].Caption = "Alle Briefe"
 ' irfan
End Function ' do_Briefe_Click(frm As Form)


Function FormAusg()
 Dim i%, ct As Control, j%
 Open "u:\Anambog.txt" For Output As #293
 With Application.Forms("Anamnesebogen")
 For i = 0 To .Controls.Count - 1
  Set ct = .Controls(i)
  Debug.Print ct.Properties.Count
  For j = 0 To ct.Properties.Count - 1
   Debug.Print " " & ct.Properties(j).Name
   Debug.Print "  " & ct.Properties(j).Value
  Next j
 Next i
' For i = 1 To .Properties.Count
'  Debug.Print i & " " & .Properties(i).Name & " "; .Properties(i).Value
' Next i
End With
 Close #293
End Function

#If False Then

DROP TABLE IF EXISTS `quelle`.`medarten`;
CREATE TABLE  `quelle`.`medarten` (
  `Medikament` varchar(50) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL,
  `Langname` varchar(150) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Beispiel-Langname',
  `Pat_ID` int(10) DEFAULT NULL COMMENT 'Beispiel-PatID',
  `Anzahl` int(10) DEFAULT NULL COMMENT 'Anzahl der Vorkommen',
  `Glib` bit(1) DEFAULT NULL COMMENT 'Glibenclamid',
  `Metf` bit(1) DEFAULT NULL COMMENT 'Metformin',
  `GlucI` bit(1) DEFAULT NULL COMMENT 'Glucosidase-Inhibitoren',
  `SHGlin` bit(1) DEFAULT NULL COMMENT 'andere Sulfonylharnstoffe oder Glinide',
  `Glit` bit(1) DEFAULT NULL COMMENT 'Glitanzone',
  `Ins` bit(1) DEFAULT NULL COMMENT 'Insulin',
  `Anal` bit(1) DEFAULT NULL COMMENT 'Insulin-Analoga',
  `InsArt` varchar(1) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL COMMENT '1= schnell, 2 = langsam, 3 = Misch',
  `HMG` bit(1) DEFAULT NULL COMMENT 'HMG-CoA-Reduktase-Inhibitoren',
  `Hypt` bit(1) DEFAULT NULL COMMENT 'Hypertonie-Mittel',
  `Thro` bit(1) DEFAULT NULL COMMENT 'Thrombozyten-Hemmer',
  `Antib` bit(1) DEFAULT NULL COMMENT 'Antibiotika',
  `and` bit(1) DEFAULT NULL COMMENT 'andere',
  `hinzugefügt` datetime DEFAULT NULL,
  `Tstr` bit(1) DEFAULT NULL COMMENT 'Teststreifen',
  `Puzu` bit(1) DEFAULT NULL COMMENT 'Pumpenzubehör',
  `VMat` bit(1) DEFAULT NULL COMMENT 'Verbandsmaterial',
  `PenN` bit(1) DEFAULT NULL COMMENT 'Pennadeln',
  `Neurp` bit(1) DEFAULT NULL COMMENT 'Neuropathie-Behandlungsmittel',
  `AutNP` bit(1) DEFAULT NULL COMMENT 'Autonome Neuropathie',
  `Fetts` bit(1) DEFAULT NULL COMMENT 'Fibrate, Ezetrol, Niaspan u.a.',
  `Hsre` bit(1) DEFAULT NULL COMMENT 'Hyperuriämie-Mittel',
  `AntiMyk` bit(1) DEFAULT NULL COMMENT 'Antimykotika',
  `Glauk` bit(1) DEFAULT NULL COMMENT 'Glaukom',
  `COLD` bit(1) DEFAULT NULL COMMENT 'COLD und Asthma',
  `Pros` bit(1) DEFAULT NULL COMMENT 'Prostatahypertrophie',
  `Urä` bit(1) DEFAULT NULL COMMENT 'Urämie-spezifische',
  `HyThy` bit(1) DEFAULT NULL COMMENT 'thyreostatische Mittel',
  `Ostp` bit(1) DEFAULT NULL COMMENT 'Osteoporosemittel',
  `KHK` bit(1) DEFAULT NULL COMMENT 'KHK-spezifisch',
  `HerzI` bit(1) DEFAULT NULL COMMENT 'Herzinsuffizienz-spezifische',
  `Stru` bit(1) DEFAULT NULL COMMENT 'Struma- und Hypothyreosemittel',
  `AVK` bit(1) DEFAULT NULL COMMENT 'AVK-Mittel',
  `PanI` bit(1) DEFAULT NULL COMMENT 'Pankreasinsuffizienz',
  `Vari` bit(1) DEFAULT NULL COMMENT 'Varikosemittel',
  `Östr` bit(1) DEFAULT NULL COMMENT 'Östrogene, Gestagene usw.',
  `AntiDep` bit(1) DEFAULT NULL COMMENT 'Antidepressiva',
  `AntiDem` bit(1) DEFAULT NULL COMMENT 'Antidementika',
  `AntiEp` bit(1) DEFAULT NULL COMMENT 'Antiepileptika',
  `Park` bit(1) DEFAULT NULL COMMENT 'Parkinson-Medikament',
  `AntiPern` bit(1) DEFAULT NULL COMMENT 'Antiperniziosa',
  `Appet` bit(1) DEFAULT NULL COMMENT 'Appetitzügler',
  `Anäm` bit(1) DEFAULT NULL COMMENT 'Anämiebehandlungsmittel',
  `Antiherp` bit(1) DEFAULT NULL COMMENT 'Anti-Herpes-Mittel',
  `NSAR` bit(1) DEFAULT NULL COMMENT 'NSAR',
  `Antikoag` bit(1) DEFAULT NULL COMMENT 'Antikoagulatien',
  `Betabl` bit(1) DEFAULT NULL COMMENT 'Betablocker',
  `ACEH` bit(1) DEFAULT NULL COMMENT 'ACE-Hemmer',
  `AT1` bit(1) DEFAULT NULL COMMENT 'AT-1-Blocker',
  `CalcA` bit(1) DEFAULT NULL COMMENT 'Calcium-Antagonist',
  `Diur` bit(1) DEFAULT NULL COMMENT 'Diuretikum',
  `falsch` bit(1) DEFAULT NULL COMMENT 'falsch geschrieben oder nicht erkennbar',
  UNIQUE KEY `Medikament` (`Medikament`),
  KEY `pat_id` (`Pat_ID`),
  CONSTRAINT `NamenMedArten_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;


DROP TABLE IF EXISTS `quelle`.`medplan`;
CREATE TABLE  `quelle`.`medplan` (
  `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug',
  `Pat_ID` int(10) DEFAULT NULL COMMENT '3000',
  `MPNr` int(10) DEFAULT NULL COMMENT 'Ordnungsziffer für Medikamentenplan',
  `ZeitPunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt, der Speicherung im Turbomed',
  `Datum` datetime DEFAULT NULL COMMENT 'Zeitpunkt aus dem Kopf des Medikamentenplans',
  `Medikament` varchar(75) COLLATE latin1_german2_ci DEFAULT NULL,
  `MedAnfang` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL,
  `FeldNr` smallint(6) DEFAULT NULL,
  `mo` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `mi` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `nm` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `ab` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `zn` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `bBed` bit(1) DEFAULT NULL,
  `Bemerkung` longtext COLLATE latin1_german2_ci,
  `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei',
  `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit',
  `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung',
  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`FeldNr`),
  KEY `FälleMedPlan` (`FID`),
  KEY `FID` (`FID`),
  KEY `MedPlanMedikament` (`Medikament`),
  KEY `NamenMedPlan` (`Pat_ID`),
  KEY `MedArtenMedPlan_AccRel` (`MedAnfang`),
  CONSTRAINT `FälleMedPlan_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`),
  CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`),
  CONSTRAINT `NamenMedPlan_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;


alter table medplan drop foreign key MedArtenMedPlan_AccRel;
ALTER TABLE medplan MODIFY COLUMN `MedAnfang` VARCHAR(35) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL;
delete from medarten;
alter table medarten modify column Medikament varchar(50) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL;
insert into medarten (medikament) SELECT distinct medanfang FROM medplan m;
alter TABLE  medplan add CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`);

#End If

Function asctest()
Dim i%, j%
For i = 0 To 255
 Debug.Print i, Chr$(i)
Next i
End Function
Public Function zeigviews()
 Dim rV As New ADODB.Recordset, rCn As New ADODB.Connection, i%, Stri$
' Call rCn.Open(CStrAcc & StACCDB)
 Call rCn.Open(CStrMy & "quelle")
' Set rV = rCn.OpenSchema(adSchemaViews)
 Set rV = rCn.OpenSchema(adSchemaTables)
 Do While Not rV.EOF
  If rV!table_type = "VIEW" Then
   Stri = ""
   For i = 0 To rV.Fields.Count - 1
    Stri = Stri & " " & rV.Fields(i).Name & ":" & rV.Fields(i)
   Next i
   Debug.Print Stri
  End If
  rV.Move 1
 Loop
End Function

Public Function btest()
 Dim rV As New ADODB.Recordset, rCn As New ADODB.Connection, i%, Stri$, rAF&
' Call rCn.Open(CStrAcc & StACCDB)
 Call rCn.Open(CStrMy & "quelle")
 rV.Open "select tkz,pat_id from anamnesebogen where pat_id=2", rCn, adOpenStatic, adLockOptimistic
 rV!Tkz = 1
 If rV.EditMode = adEditInProgress Then rV.Update
 Debug.Print rV!Tkz
 rV!Tkz = 0
 If rV.EditMode = adEditInProgress Then rV.Update
 Debug.Print rV!Tkz
 rV!Tkz = 1
 If rV.EditMode = adEditInProgress Then rV.Update
 Debug.Print rV!Tkz
 Call rCn.Execute("update anamnesebogen set tkz = 1 where pat_id = 2", rAF)
 rV.Close
 rV.Open "select tkz,pat_id from anamnesebogen where pat_id=2", rCn, adOpenStatic, adLockOptimistic
 Debug.Print "rAf=", rAF, rV!Tkz
 Call rCn.Execute("update anamnesebogen set tkz = 0 where pat_id = 2", rAF)
 rV.Close
 rV.Open "select tkz,pat_id from anamnesebogen where pat_id=2", rCn, adOpenStatic, adLockOptimistic
 Debug.Print "rAf=", rAF, rV!Tkz
 Call rCn.Execute("update anamnesebogen set tkz = -1 where pat_id = 2", rAF)
 rV.Close
 rV.Open "select tkz,pat_id from anamnesebogen where pat_id=2", rCn, adOpenStatic, adLockOptimistic
 Debug.Print "rAf=", rAF, rV!Tkz
End Function

Function dovCommandB_Click(Index%)
' Stop
End Function


Public Function do_abgehakt_Click(frm As Form) ' für Labordokumente eP
 Static Pat_id&, cR
 On Error GoTo fehler
 Dim rs As New ADODB.Recordset
 rs.Open frm.adoRS.source, DBCn, adOpenStatic, adLockReadOnly
 If Not rs.BOF Then
'  dbcn.Execute("update ")
 Else
'
 End If
' rDokab.Seek "=", frm.DokPfad
' If Not rDokab.NoMatch Then
'  rDokab.Edit
' Else
'  rDokab.AddNew
'  rDokab!DokPfad = frm.DokPfad
' End If
' If rDokab!abgehakt Then
'  rDokab!abgehakt = False
' Else
'  rDokab!abgehakt = True
' End If
' rDokab!AktZeit = Now
' rDokab.Update
' Pat_id = frm.Pat_id
' cR = frm.CurrentRecord
  frm.Requery
 On Error Resume Next
' DoCmd.GoToRecord acDataForm, frm.Name, acGoTo, cR
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in do_abgehakt_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_abgehakt_Click(frm as Form)

' Labordokumente eP
Public Function doDatensatzPosition$(frm As Form)
  Dim rs As New ADODB.Recordset
  If Not IsNull(frm.adoRS!DokPfad) Then 'vTextB(6)) Then ' DokPfad
    rs.Open frm.adoRS.source, DBCn, adOpenStatic, adLockReadOnly
    rs.Find "dokpfad = '" & frm.adoRS!DokPfad & "'"
    If Not rs.EOF Then
     doDatensatzPosition = CStr(rs.AbsolutePosition + 1)
    End If
'    Set rs = frm.RecordsetClone
'    rs.FindFirst ("LaborDokumente.Dokpfad=" & "" & frm!DokPfad & "")
'    DatensatzPosition = CStr(rs.AbsolutePosition + 1)
'    rs.Close
  End If
End Function ' doDatensatzPosition

Public Function doAktAbgehakt(frm As Form)
 frm.adoRS.Find "dokpfad = '" & frm.DokPfad & "'"
 If frm.adoRS.EOF Then
  doAktAbgehakt = False
 Else
  doAktAbgehakt = frm.adoRS!abgehakt
 End If
' rDokab.Seek "=", frm.DokPfad
' If rDokab.NoMatch Then
'  AktAbgehakt = False
' Else
'  AktAbgehakt = rDokab!abgehakt
' End If
End Function
Public Function do_obAbgehakt(frm As Form)
 frm.adoRS.Find "dokpfad = '" & frm.DokPfad & "'"
 If frm.adoRS.EOF Then
  do_obAbgehakt = ""
 Else
  If frm.adoRS!abgehakt <> 0 Then
   do_obAbgehakt = "ja"
  Else
   do_obAbgehakt = ""
  End If
  do_obAbgehakt = frm.adoRS!abgehakt
 End If
' rDokab.Seek "=", frm.DokPfad
' If rDokab.NoMatch Then
'  obAbgehakt = ""
' Else
'  If rDokab!abgehakt Then
'   obAbgehakt = "ja"
'  Else
'   obAbgehakt = ""
'  End If
' End If
End Function

Public Function doAltName(frm As Form)
 Static AltN$, NeuN$, altMerk$
 NeuN = frm.NachName + frm.Vorname
 If NeuN <> AltN Then
  If altMerk = "neu" Then
   altMerk = ""
  Else
   altMerk = "neu"
  End If
 End If
 doAltName = altMerk
 AltN = NeuN
End Function ' doAltName

Public Function doForm_Load(frm As Form)
  Dim db As New Connection
  Dim oText As TextBox
  Dim oCheck As CheckBox
  Dim Pat_id&
   db.CursorLocation = adUseClient
   db.Open DBCn
   Set frm.adoRS = New Recordset
  
  If frm.Name Like "AnBog*" Then
   syscmd 4, "Lade Datensätze von der Datenbank ..."
   obstumm = True
   If LVobMySQL Then
'    frm.adoRS.Open "select *,if(größe=0,'',gewicht/größe/größe*if(größe>3,10000,1)) as bmi, concat(nachname,' ', vorname) as gesname  from anamnesebogen where pat_id in (select pat_id from aktfv) order by pat_id", db, adOpenStatic, adLockOptimistic
    frm.adoRS.Open "select *,if(größe=0,'',gewicht/größe/größe*if(größe>3,10000,1)) as bmi, concat(nachname,' ', vorname) as gesname  from anamnesebogen order by pat_id", db, adOpenDynamic, adLockOptimistic
   Else
    frm.adoRS.Open "select *,iif(größe=0,'',gewicht/größe/größe*iif(größe>3,10000,1)) as bmi, nachname & ' ' & vorname as gesname  from anamnesebogen  order by pat_id", db, adOpenDynamic, adLockOptimistic
   End If
   obstumm = False
   On Error Resume Next
   Pat_id = fWertLesen(HCU, RegWurzel & App.EXEName, "pat_id")
   On Error GoTo 0
   frm.adoRS.Find "pat_id >= " & Pat_id
  'Kontrollkästchen an Datenprovider binden
   For Each oCheck In frm.vCheckb
    Set oCheck.DataSource = frm.adoRS
   Next
  ElseIf frm.Name = "Medarten" Then
   syscmd 4, "Lade Datensätze von der Datenbank ..."
   obstumm = True
   frm.adoRS.Open "select * from medarten order by hinzugefügt desc", db, adOpenStatic, adLockOptimistic
  'Kontrollkästchen an Datenprovider binden
   For Each oCheck In frm.vCheckb
    Set oCheck.DataSource = frm.adoRS
   Next
  End If
  'Textfelder an Datenprovider binden
   For Each oText In frm.vTextB
    Set oText.DataSource = frm.adoRS
   Next
  frm.mbDataChanged = False
  Lese.Hide
 'Me.Controls!Diagnosen.ControlSource = "=replace(replace(recordset!Diagnosen,vbverticaltab,vbcr+vblf),vbtab,"" "")"
End Function ' doFormLoad

Public Function doKeyDown(frm As Form, keyCode%, Shift%)
 If frm.Name Like "AnBog*" Then
   Select Case keyCode
    Case 27
     Call AnBogUnload(frm)
   End Select
 ElseIf frm.Name Like "Medarten*" Then
   Select Case keyCode
    Case 27
     Unload frm
   End Select
 End If
End Function
Public Function AnBogUnload(frm As AnBog)
  Call fdwSpei(HCU, RegWurzel & App.EXEName, "pat_id", frm.adoRS!Pat_id)       ' für MySQL
 Unload frm
End Function
Public Function diagexpHerricht()
 On Error GoTo fehler
 DBCn.Execute ("delete from diagnosenexport")
 DBCn.Execute ("insert into diagnosenexport select * from fuerdiagexp")
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in diagexpHerricht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Public Function GesDiagExp(Optional obTest%) ' Im Menü unter "Diagnosen exportieren"
 On Error GoTo fehler
 Call diagexpHerricht

' On Error Resume Next
' Dim test$
' test = Forms!Anamnesebogen.Name
' If Err.Number = 0 Then
'  On Error GoTo fehler
'  Call do_Diagnosen_Reset(Forms!Anamnesebogen) ' muß offen sein
' Else
'  On Error GoTo fehler
'  Call do_Diagnosen_Reset
' End If
 Call doDiagnosenexport(obTest:=obTest)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GesDiagExp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DiagExpGesamt()

Public Function doDiagnosenexport(Optional obTest%)
 On Error GoTo fehler
 Dim Quartal$, erg$, DZahl&
' Dim Q As DAO.Recordset, rFa As DAO.Recordset, n As DAO.Recordset, ü As DAO.Recordset
 Dim q As New ADODB.Recordset, rFa As New ADODB.Recordset, n As New ADODB.Recordset, Ü As New ADODB.Recordset
 Dim rDT As New ADODB.Recordset
 Dim op$
 Dim z$, ZDat$ ', DDat As Date, rohdat$
 Dim aktDat As Date
' Const quartal$ = "22005"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Diagnosen exportieren ...")
 Call ImportFolderHerricht
' Do While Not IsDate(rohdat)
'  rohdat = InputBox("Für welches Datum sollen die Diagnosen eingetragen werden?")
' Loop
' DDat = CDate(rohdat)
 ZDat = "DIAG " + Format$(Now, "dd/mm/yy HH.MM") + ".BDT"
 z = hVerz & ZDat
 Open z For Output As #327
 Close #327
 erg = Dir(hVerz & ZDat)
 If LenB(erg) = 0 Then
  MsgBox "Fehler beim Erstellen der Exportdatei '" & hVerz & ZDat & "'"
  Exit Function
 End If
' Set Q = Dtb.OpenRecordset("DiagnosenExport", dbOpenDynaset)
 q.Open "select * from diagnosenexport", DBCn, adOpenDynamic, adLockOptimistic
 If q Is Nothing Then
  MsgBox "Fehler beim  Öffnen der Tabelle `diagnosenexport`"
  Exit Function
 ElseIf q.BOF Then
  MsgBox "Keine Diagnosen zum Exportieren!"
  Exit Function
 End If
 
' Set ü = Dtb.OpenRecordset("Diagnosen exportiert", dbOpenTable)
 Ü.Open "select * from `diagnosen exportiert`", DBCn, adOpenDynamic, adLockOptimistic
 If Ü Is Nothing Then
  MsgBox "Fehler beim  Öffnen der Tabelle `diagnosen exportiert`"
  Exit Function
 End If
 
 Open z For Append As #327
' Set rFa = TabÖff("faelle", "Auswahl")
' Set n = TabÖff("Namen", "Pat_ID")
 
 If Not q.BOF Then
  q.MoveFirst
  Print #327, "01380000020" ' Satzart
  Print #327, "014810000082" ' Satzlänge
  Print #327, "01691006419153" ' Arztnummer des Absenders
  Print #327, "0179103" & Format$(Now, "ddmmyyyy")  ' Erstellungsdatum
  Print #327, "0129105001" ' Ordnungsnummer Datenträger (Header) des DP
  Print #327, "01091062" ' verwendeter Zeichensatz (2=IBM-Code)
  Print #327, "01380000022" ' Satzart
  Print #327, "014810000107" ' Satzlänge
  Print #327, "014921001/99" ' Version ADT-Satzbeschreibung
  Print #327, "014921302/94" ' Version BDT-Satzbeschreibung
  Print #327, "01096001" ' Archivierungsart
  Print #327, "025960101011980" & Format$(Now, "ddmmyyyy") 'Zeitraum der Spiegelung
  Print #327, "0179602" & Format(Now, "hhmmss") & "00" ' Beginn der Übertragung
  Print #327, "01380000010" 'Satzarzt
  Print #327, "014810000315" ' Satzlänge
  Print #327, "0160101A001011" ' KBV-Prüfnummer
  Print #327, "0260102TurboMed EDV GmbH" ' Softwareverantwortlicher
  Print #327, "0250103TurboMed@Windows" ' Software
  Print #327, "0180104IBM PC/AT"
  Print #327, "01002021" ' Praxistyp (1=Einzelpraxis)
  Print #327, "0138000besa"
  Print #327, "014810000213"
  Print #327, "0160201641915300" ' Arztnummer
  Print #327, "0220203Schade" ' Arztname
  Print #327, "0160212889690003" ' Lebenslange Arztnummer
  Print #327, "0220211Gerald Schade" 'Arztname für Leistungsdifferenzierung
  Print #327, "0500204TG Diabetologie" ' FA Innere und Allgemeinmedizin (Hausarzt)"
  Print #327, ZSU("0280205Mittermayerstraße 13") ' Straße der Praxis
  Print #327, "014021585221" ' PLZ der Praxisadresse
  Print #327, "0150216Dachau" ' Ort der Praxisadresse
  Print #327, "024020808131 / 616 380" ' Telefonnummer der Praxis
  Print #327, "024020908131 / 616 381" ' Telefaxnummer der Praxis
  Dim altpat_id&
  altpat_id = 0
'  Set rDT = TabÖff("Diagnosen", "DiagSuch")
  Set rDT = Nothing
'  rDT.Open "select * from diagnosen", DBCn, adOpenDynamic, adLockOptimistic
  Do While Not q.EOF
'   If q!Pat_id = 77 Then
   If Not IsNull(q!Pat_id) And q!Pat_id <> 0 Then
    If IsNull(q!Status) Or (Not IsNull(q!Status) And q!Status <> übertragen) Then
     If IsNull(q!Zeitpunkt) Then aktDat = lFDat(q!Pat_id) Else aktDat = q!Zeitpunkt
     Quartal = ZQuart(aktDat)
     If q!Pat_id <> altpat_id Then
      Set rFa = Nothing
      rFa.Open "select * from faelle where pat_id = " & q!Pat_id & " and quartal = '" & Quartal & "' order by bhfb desc", DBCn, adOpenDynamic, adLockOptimistic
'      rFa.Seek "=", Q!Pat_id, quartal
'      If rFa.NoMatch Then
      If rFa.BOF Then
       Set rFa = Nothing
       rFa.Open "select * from faelle where pat_id = " & q!Pat_id & " order by bhfb desc", DBCn, adOpenDynamic, adLockOptimistic
'       rFa.Seek "=", Q!Pat_id
      End If
      If Not rFa.BOF Then
'      If Not rFa.NoMatch Then
       Set n = Nothing
       n.Open "select * from namen where pat_id = " & q!Pat_id, DBCn, adOpenDynamic, adLockOptimistic
'       n.Seek "=", Q!Pat_id
       If Not n.BOF Then
'       If Not n.NoMatch Then
        Print #327, "0138000" + "6200" ' IIf(rFa!SchGr = "90", "0190", "0102") ' Satzidentifikation (s.geslies)
' bei 0101 entstehen bei zwei Aufrufen fehlerfrei zwei neue Kassenfaelle, jeder mit der Leistung
' bei 190 entsteht ein neuer Privatfall, bei 6100 läuft alles ohne Fehler durch, aber keine Leistung steht drin
' bei 6200 entsteht ein neuer Kassenfall
        Print #327, "014810000845" ' Satzlänge, noch nicht entschlüsselt"
'     op = format$(3 + 4 + 4, "000") + "8000" + CStr(f!s8000)
'     Print #327, zsu(op)
'     op = format$(3 + 4 + 5, "000") + "8100" + CStr(f!s8100)
'     Print #327, zsu(op)
        op = Format$(3 + 4 + Len(CStr(q!Pat_id)), "000") + "3000" + CStr(q!Pat_id)
        Print #327, ZSU(op)
'        If 1 = 0 Then 'Auswirkung bisher nicht geprüft 31.7.05 (3x)
        If Not IsNull(n!NVorsatz) Then
         If n!NVorsatz <> "" Then
          op = Format$(3 + 4 + Len(n!NVorsatz), "000") + "3100" + CStr(n!NVorsatz)
          Print #327, ZSU(op)
         End If
        End If
        op = Format$(3 + 4 + Len(n!NachName), "000") + "3101" + CStr(n!NachName)
        Print #327, ZSU(op)
        op = Format$(3 + 4 + Len(n!Vorname), "000") + "3102" + CStr(n!Vorname)
        Print #327, ZSU(op)
        Dim FGeb$
        FGeb = Format$(n!GebDat, "ddmmyyyy")
        If Mid$(FGeb, 5, 2) = "20" Then
         If n!GebDat > Date Then FGeb = Left(FGeb, 4) & "19" & Right$(FGeb, 2)
        End If
        op = Format$(3 + 4 + 8, "000") + "3103" + FGeb
        Print #327, ZSU(op)
        
        If 1 = 0 Then ' 26.6.08: Für Diagnosen nicht nötig
         If Not IsNull(n!Versichertennummer) Then
          op = Format$(3 + 4 + Len(n!Versichertennummer), "000") + "3105" + n!Versichertennummer
          Print #327, ZSU(op)
         End If
'     op = format$(3 + 4 + Len(n!Straße), "000") + "3107" + n!Straße
'     Print #327, ZSU(op)
'     op = format$(3 + 4 + Len(n!Plz), "000") + "3112" + n!Plz
'     Print #327, ZSU(op)
'     op = format$(3 + 4 + Len(n!Ort), "000") + "3113" + n!Ort
'     Print #327, ZSU(op)
         If Not IsNull(n![KVKStatus]) And n![KVKStatus] > 0 And Trim$(n![KVKStatus]) <> "" Then
          op = Format$(3 + 4 + Len(n![KVKStatus]), "000") + "3108" + n![KVKStatus]
          Print #327, ZSU(op)
         End If
         op = Format$(3 + 4 + 1, "000") + "3110" + IIf(n!Geschlecht = "m", "1", IIf(n!Geschlecht = "w", "2", ""))
         Print #327, ZSU(op)
         If rFa!SchGr <> "90" Then
          op = Format$(3 + 4 + 5, "000") + "4101" + ZQuart(rFa!ausgst) 'quartal
          Print #327, ZSU(op)
          op = Format$(3 + 4 + 8, "000") + "4102" + Format$(rFa!ausgst, "ddmmyyyy")
          Print #327, ZSU(op)
          op = Format$(3 + 4 + Len(rFa!VKNr), "000") + "4104" + rFa!VKNr
          Print #327, ZSU(op)
         End If
         If Not IsNull(rFa!KtrAbrB) Then  ' bei Privaten
          op = Format$(3 + 4 + Len(rFa!KtrAbrB), "000") + "4106" + rFa!KtrAbrB
          Print #327, ZSU(op)
         End If
         If Not IsNull(rFa!AbrAr) Then  ' bei Privaten
          op = Format$(3 + 4 + Len(rFa!AbrAr), "000") + "4107" + rFa!AbrAr
          Print #327, ZSU(op)
         End If
         If rFa!SchGr <> "90" Then
          op = Format$(3 + 4 + 8, "000") + "4109" + Format$(rFa!lVorl, "ddmmyyyy")
          Print #327, ZSU(op)
          op = Format$(3 + 4 + 4, "000") + "4110" + Format$(rFa!lVorl, "hhmm")
          Print #327, ZSU(op)
          op = Format$(3 + 4 + Len(rFa!IK), "000") + "4111" + rFa!IK
          Print #327, ZSU(op)
          If Not IsNull(rFa!KVKs) Then ' bei Pat_id 43
           op = Format$(3 + 4 + Len(rFa!KVKs), "000") + "4112" + rFa!KVKs
           Print #327, ZSU(op)
          End If
          If Not IsNull(rFa!KVKserg) Then
           op = Format$(3 + 4 + Len(rFa!KVKserg), "000") + "4113" + rFa!KVKserg
           Print #327, ZSU(op)
          End If
         End If
'         FGeb = format$(rFa!GebOr, "ddmmyyyy")
'         If mid$(FGeb, 5, 2) = "20" Then
'          If n!GebDat > Date Then FGeb = left(FGeb, 4) & "19" & right$(FGeb, 2)
'         End If
'         If Not IsNull(FGeb) Then ' bei Privaten
         If Not IsNull(rFa!GebOr) Then
'          op = format$(3 + 4 + Len(FGeb), "000") + "4121" & FGeb
          op = Format$(3 + 4 + Len(rFa!GebOr), "000") + "4121" & rFa!FGeb
          Print #327, ZSU(op)
         End If
         If Not IsNull(rFa!AbrGb) Then ' bei Privaten
          op = Format$(3 + 4 + Len(rFa!AbrGb), "000") + "4122" + rFa!AbrGb
          Print #327, ZSU(op)
         End If
         op = Format$(3 + 4 + Len("TM#" + IIf(IsNull(rFa!TMFNr), "", rFa!TMFNr)), "000") + "4144" + "TM#" + IIf(IsNull(rFa!TMFNr), "", rFa!TMFNr)
         Print #327, ZSU(op)
         op = Format$(3 + 4 + 8, "000") + "4150" + IIf(rFa!BhFB = 0, "00000000", Format$(rFa!BhFB, "ddmmyyyy"))
         Print #327, ZSU(op)
         op = Format$(3 + 4 + 8, "000") + "4151" + IIf(rFa!BhFE1 = 0, "00000000", Format$(rFa!BhFE1, "ddmmyyyy"))
         Print #327, ZSU(op)
         op = Format$(3 + 4 + 8, "000") + "4152" + IIf(rFa!BhFE2 = 0, "00000000", Format$(rFa!BhFE2, "ddmmyyyy"))
         Print #327, ZSU(op)
         If 1 = 1 Then
          Dim Üw$
          Üw = IIf(IsNull(rFa!ÜbwV), "", rFa!ÜbwV)
          If Üw <> "" Then
           op = Format$(3 + 4 + Len(Üw), "000") + "4218" + Üw
           Print #327, ZSU(op)
          End If
          Üw = IIf(IsNull(rFa!AndÜw), "", rFa!AndÜw)
          If Üw <> "" Then
           op = Format$(3 + 4 + Len(Üw), "000") + "4219" + Üw
           Print #327, ZSU(op)
          End If
          If Not IsNull(rFa!ÜWZiel) And rFa!ÜWZiel <> "" Then
           op = Format$(3 + 4 + Len(rFa!ÜWZiel), "000") + "4220" + rFa!ÜWZiel
           Print #327, ZSU(op)
          End If
         End If ' 1 = 1
         op = Format$(3 + 4 + 2, "000") + "4239" + rFa!SchGr
         Print #327, ZSU(op)
         If Not IsNull(rFa!PGeb) Then
          FGeb = Format$(CStr(rFa!PGeb), "ddmmyyyy")
          If Mid$(FGeb, 5, 2) = "20" Then
           If n!GebDat > Date Then FGeb = Left(FGeb, 4) & "19" & Right$(FGeb, 2)
          End If
          Üw = "TM#" + FGeb
          op = Format$(3 + 4 + Len(Üw), "000") + "4401" + Üw
         End If
         FGeb = Format$(rFa!PGebErg, "ddmmyyyy")
         If Mid$(FGeb, 5, 2) = "20" Then
          If n!GebDat > Date Then FGeb = Left(FGeb, 4) & "19" & Right$(FGeb, 2)
         End If
         If Not IsNull(FGeb) Then
          Üw = "TM#" & FGeb
          op = Format$(3 + 4 + Len(Üw), "000") + "4402" + Üw
         End If
         If (Not IsNull(rFa!GOÄKatNr) And rFa!GOÄKatNr <> "") Or (Not IsNull(rFa!GOÄKatName) And rFa!GOÄKatName <> "") Then
          Üw = "TM#" + CStr(rFa!GOÄKatNr) + "#" + CStr(rFa!GOÄKatName)
          op = Format$(3 + 4 + Len(Üw), "000") + "4580" + Üw
          Print #327, ZSU(op)
         End If
' if 1 = 0
         Print #327, "01380006200"
         Print #327, "014810000039"
         op = Format$(3 + 4 + Len(CStr(q!Pat_id)), "000") + "3000" + CStr(q!Pat_id)
         Print #327, ZSU(op)
         Print #327, "01380006100"
         Print #327, "014810000369"
         op = Format$(3 + 4 + Len(CStr(q!Pat_id)), "000") + "3000" + CStr(q!Pat_id)
         Print #327, ZSU(op)
         If 1 = 0 Then 'Auswirkung bisher nicht geprüft 31.7.05 (3x)
          op = Format$(3 + 4 + Len(n!NVorsatz), "000") + "3100" + CStr(n!NVorsatz)
          Print #327, ZSU(op)
         End If
         op = Format$(3 + 4 + Len(n!NachName), "000") + "3101" + CStr(n!NachName)
         Print #327, ZSU(op)
         op = Format$(3 + 4 + Len(n!Vorname), "000") + "3102" + CStr(n!Vorname)
         Print #327, ZSU(op)
         FGeb = Format$(n!GebDat, "ddmmyyyy")
         If Mid$(FGeb, 5, 2) = "20" Then
          If n!GebDat > Date Then FGeb = Left(FGeb, 4) & "19" & Right$(FGeb, 2)
         End If
         op = Format$(3 + 4 + 8, "000") + "3103" + FGeb
         Print #327, ZSU(op)
         If Not IsNull(n!Versichertennummer) Then
          op = Format$(3 + 4 + Len(n!Versichertennummer), "000") + "3105" + n!Versichertennummer
          Print #327, ZSU(op)
         End If
'     op = format$(3 + 4 + Len(n!Straße), "000") + "3107" + n!Straße
'     Print #327, ZSU(op)
'     op = format$(3 + 4 + Len(n!Plz), "000") + "3112" + n!Plz
'     Print #327, ZSU(op)
'     op = format$(3 + 4 + Len(n!Ort), "000") + "3113" + n!Ort
'     Print #327, ZSU(op)
         If Not IsNull(n![KVKStatus]) And n![KVKStatus] > 0 And Trim$(n![KVKStatus]) <> "" Then
          op = Format$(3 + 4 + Len(n![KVKStatus]), "000") + "3108" + n![KVKStatus]
          Print #327, ZSU(op)
         End If
         op = Format$(3 + 4 + 1, "000") + "3110" + IIf(n!Geschlecht = "m", "1", IIf(n!Geschlecht = "w", "2", ""))
         Print #327, ZSU(op)
' 3610
         op = Format$(3 + 4 + 8, "000") + "3610" + Format$(n!AufnDat, "ddmmyyyy")
         Print #327, ZSU(op)
'( 3626
' 3629
' 3626
' 3629)
' 3630
         If 1 = 0 Then ' führt zu doppelten HausarztEintraegen
          If Not IsNull(n!KVNr) Then
           op = Format$(3 + 4 + Len(n!KVNr), "000") + "3630" + CStr(n!KVNr)
           Print #327, ZSU(op)
          End If
         End If ' 1 = 0
' 3631
         op = Format$(3 + 4 + Len(Trim$(CStr(IIf(IsNull(n!Weggeldzone), "", n!Weggeldzone)))), "000") + "3631" + Trim$(CStr(IIf(IsNull(n!Weggeldzone), "", n!Weggeldzone)))
         Print #327, ZSU(op)
        End If ' 1 = 0
       End If ' not n.bof
       altpat_id = q!Pat_id
      End If ' Not f.NoMatch Then
     End If ' q!Pat_id <> altpat_id Then
     Dim DiagText$
     op = Format$(3 + 4 + 8, "000") + "3649" + Format$(aktDat, "ddmmyyyy")
     Print #327, ZSU(op)
     op = Format$(3 + 4 + 4, "000") + "6201" + Format$(aktDat, "HHMM")
     Print #327, ZSU(op)
     Dim ICD$, DiagSi$
     ICD = q!ICD
     DiagSi = "G"
     DiagText = q!Diagnose
     If Not IsNull(q!ICD) Then
      If q!ICD <> "" Then
       If InStrB("VGZA", Right$(q!ICD, 1)) > 0 And Right$(q!ICD, 1) <> "" Then
        DiagSi = Right$(q!ICD, 1)
        ICD = Left(q!ICD, Len(q!ICD) - 1)
        Select Case DiagSi
         Case "V": DiagText = LTrim$(Replace(DiagText, "V.a.", ""))
         Case "Z": DiagText = LTrim$(Replace(DiagText, "Z.n.", ""))
         Case "A": DiagText = LTrim$(Replace(DiagText, "Ausschluss ", ""))
        End Select
       End If
      End If
     End If
     Dim BDTFeld$
     If IsNull(q!nurquart) Then
      BDTFeld = "3650"
     ElseIf q!nurquart = 0 Then
      BDTFeld = "3650"
     Else
      BDTFeld = "6000"
     End If
     op = Format$(3 + 4 + Len(DiagText), "000") & BDTFeld & DiagText
     Print #327, ZSU(op)
     op = Format$(3 + 4 + Len(ICD), "000") + "6001" + ICD
     Print #327, ZSU(op)
     op = Format$(3 + 4 + Len(DiagSi), "000") + "6003" + DiagSi
     Print #327, ZSU(op)
     DZahl = DZahl + 1
' jetzt kommt der Diagnoseneintrag in die Tabelle
'     rDT.Index = "DiagSuch" ' pat_id, icd, diagsicherheit, diagseite
'     rDT.Seek "=", n!Pat_id, Q!ICD
'#If False Then
     Set rDT = Nothing
'     rDT.Open "select * from diagnosen where pat_id = " & q!Pat_id & " and icd = " & q!ICD, DBCn, adOpenDynamic, adLockOptimistic
'     If Not rDT.BOF Then
     Dim rAF&, rAfL&
     If Not obTest Then
      Call DBCn.Execute("insert into diagnosen(pat_id,gesName,ICD,diagdatum,diagsicherheit,diagtext,obdauer,aktzeit) values(" & q!Pat_id & ",'" & q!Name & "','" & q!ICD & "'," & datform(aktDat) & ",'" & DiagSi & "','" & DiagText & "'," & IIf(BDTFeld = "3650", 1, 0) & "," & datform(Now) & ")", rAF)
      If rAF <> 1 Then
       MsgBox "Fehler beim Diagnoseneeinfügen für Pat. " & q!Pat_id & " (" & q!GesName & ")" & vbCrLf & "ICD: " & q!ICD & vbCrLf & "Diagtext:" & q!DiagText & vbCrLf & "Datum: " & datform(aktDat) & rAF & " Datensätze eingefügt"
      End If
      If LenB(q!ICD) <> 0 And LenB(q!Diagnose) <> 0 Then
       Call DBCn.Execute("update diagnosenexport set status = '" & übertragen & "' where id = " & q!id, rAF)
       If rAF <> 1 Then
        MsgBox "Fehler beim Statussetzen in `diagnosenexport` für ID: " & q!id & rAF & " Datensätze gesetzt"
       End If
       Call DBCn.Execute("insert into `diagnosen exportiert`(pat_id,datum,icd,diagnose,übertragen) values(" & q!Pat_id & "," & datform(aktDat) & ",'" & q!ICD & "','" & DiagText & "'," & datform(Now) & ")", rAF)
       If rAF > 0 Then
        Call DBCn.Execute("delete from fuerdiagexp where id = " & q!id, rAfL)
        If rAfL <> 1 Then
         MsgBox "Fehler beim Löschen aus `fuerdiagexp` von " & q!Pat_id & " (" & q!GesName & ")" & vbCrLf & "ICD: " & q!ICD & vbCrLf & "Diagtext:" & q!DiagText & vbCrLf & "Datum: " & datform(aktDat) & rAfL & " Datensätze gelöscht"
        End If
       End If
       If rAF <> 1 Then
        MsgBox "Fehler beim Eintragen in `diagnosen exportiert` von " & q!Pat_id & " (" & q!GesName & ")" & vbCrLf & "ICD: " & q!ICD & vbCrLf & "Diagtext:" & q!DiagText & vbCrLf & "Datum: " & datform(aktDat) & rAF & " Datensätze eingetragen"
       End If
      End If ' q!icd <> "" And q!Diagnose <> "" Then
     End If ' obTEst
'End If
'#End If
    End If ' q!status <> "übertragen"
   End If ' q!pat_id <> 0
'   End If ' q!pat_id = 77
   q.MoveNext
  Loop ' do while not q.eof
'  Dim Bm$
'  On Error Resume Next
'  Bm = Forms("Anamnesebogen").Recordset.Bookmark
  On Error GoTo fehler
'  Set Q = Dtb.OpenRecordset("Select distinct(pat_id) from [" + QMdbAkt + "].Diagnosenexport")
  Set q = Nothing
  q.Open "select distinct(pat_id) from diagnosenexport", DBCn, adOpenStatic, adLockReadOnly
  Do While Not q.EOF
   Call dynDiag(q!Pat_id)
   q.MoveNext
'   On Error Resume Next
'   Forms("Anamnesebogen").Recordset.FindFirst ("pat_id = " & CStr(q!Pat_id))
'   Forms("Anamnesebogen").Requery
'   On Error GoTo fehler
  Loop ' do while not q.eof
  On Error Resume Next
'  Forms("Anamnesebogen").Repaint
'  Forms("Anamnesebogen").Recordset.Bookmark = Bm
  On Error GoTo fehler
 End If ' Not q.BOF Then
 Close #327
 MsgBox DZahl & " Diagnosen übertragen in die Datei '" & hVerz & ZDat & "'"
' DoCmd.Close acTable, "X00D"
' If FSO Is Nothing Then Set FSO = CreateObject("scripting.FileSystemObject")
 On Error Resume Next
 FSO.DeleteFile Replace(z, ".BDT", ".txt"), True
 On Error GoTo fehler
 FSO.CopyFile z, Replace(z, ".BDT", ".txt"), True
 'Shell "cmd /c copy /y " + """" + z + """" & " " & """" + Replace(z, ".BDT", ".txt") + """"
 'Shell "notepad " & "" & z + ""
' DoCmd.OpenQuery "Leistexp 2 Abfrage"
' DoCmd.OpenQuery "LeistExp übertragen Abfrage"
' Call syscmd(4, "Verknüpfe Exportat")
' Dim con$, td As DAO.TableDef, Fil As File
' Set Fil = FSO.GetFile(Replace(Z, ".BDT", ".txt"))
' con = "Text;DSN=Labor;FMT=Fixed;HDR=NO;IMEX=2;CharacterSet=1252;DATABASE=" + hVerz + ";TABLE=" + Replace(Fil.ShortName, ".TXT", "") + "#txt"
' Set td = Dtb.CreateTableDef("DiagExpVerknue", dbAttachExclusive, Fil.ShortName, con)
' On Error Resume Next
' DoCmd.Close acTable, "DiagExpVerknue", acSaveNo
' Dtb.Execute "drop table [" + "DiagExpVerknue" + "];"
' On Error GoTo fehler
' Dtb.TableDefs.Append td
' Call BDTVerschönere
' Call syscmd(acSysCmdSetStatus, "Fertig mit Diagnosen exportieren")
 GoTo schluss
eintragsfehler:
 Dim tonRunde%
 For tonRunde = 1 To 10
  Call Sound(WinDir + "\media\Windows XP-Standard.wav")
 Next tonRunde
 Call VerzPrüf(üVerz)
 Open üVerz + "Eintragsfehler" For Append As #26
 Print #26, CStr(rDT!Pat_id) & " " & CStr(rDT!GesName) & " " & IIf(IsNull(rDT!ICD), "", rDT!ICD) & " " & IIf(IsNull(rDT!DiagText), "", rDT!DiagText)
 Close #26
 Resume Next
schluss:
  MsgBox "Fertig mit DiagnosenExport!"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DiagnosenExport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DiagnosenExport()
Function dynDiag$(Pat_id&) ' für DiagnosenExport
 On Error GoTo fehler
 Dim DBDiagString$
 Dim rAF&
 DBDiagString = DiagString$(Pat_id)
 If InStrB(DBDiagString, "'") <> 0 Then
  DBDiagString = Replace(DBDiagString, "'", "''")
 End If
 Call DBCn.Execute("update anamnesebogen set diagnosen = '" & DBDiagString & "' where pat_id = " & Pat_id, rAF)
' Set rana = TabÖff("Anamnesebogen", "pat_id")
' If Not rana.BOF Then
'  rana.Seek "=", Pat_id
'  If Not rana.NoMatch Then
'   rana.Edit
'   rana!Diagnosen = Diagstring$(Pat_id)
'   If rana!Pat_id = 0 Then
'    rana.CancelUpdate ' kommt auf mir noch unbekanntem Weg zustande
'    Stop
'   Else
'    rana.Update
'   End If
'  End If
' End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dyndiag/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dyndiag$

Function lFDat(Pat_id&) As Date ' letztes Falldatum, falls später als heute, dann heute
' dtbInit
' lFDat = Dtb.OpenRecordset("select max(bhfe1) as bhfe from faelle where pat_id = " + CStr(Pat_id))!BhFE
 Dim rs As New ADODB.Recordset
 rs.Open "select max(bhfe1) as bhfe from faelle where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 If Not rs.BOF Then
  lFDat = rs!BhFE
 End If
 If lFDat > Now Then lFDat = Now
End Function ' lFDat(pat_id&) As Date

Function test_fdübertrag()
 Dim q As New ADODB.Recordset, z As New ADODB.Connection, i%
 Const JCn$ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & StACCDB & "';"
 z.Open JCn
 Call Lese.ProgStart
 q.Open "select * from fuerdiagexp", DBCn, adOpenStatic, adLockReadOnly
 Do While Not q.EOF
  z.Execute ("insert into fuerdiagexp(name,pat_id,icd,diagnose) values('" & q!Name & "'," & q!Pat_id & ",'" & q!ICD & "','" & q!Diagnose & "')")
  q.Move 1
 Loop
End Function
