Attribute VB_Name = "HausärzteHier"
Const haz$ = "HAz" 'Zwischenname
Const hae$ = "Hausaerzte" ' endgültig
'Const Datei$ = uVerz + "HA.doc"
Public Const opti& = 1 + 2 '+ 2048  ' 4 + 8 '+ 32 + 131118 ' 131118, 32 ' + 2048 + 16384
'Public Const CSStr$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux1;uid=praxis;pwd=***REMOVED***;database=kvaerzte;OPTION=" & opti
#If False Then
Function dbtest()
 Dim i%, rs As DAO.Recordset, name$, FNr%
 On Error GoTo fehler
 For i = 0 To Dtb.TableDefs.COUNT - 1
  name = Dtb.TableDefs(i).name
  On Error Resume Next
  Set rs = Dtb.OpenRecordset(name, dbOpenTable)
  FNr = Err.Number
  On Error GoTo fehler
  Debug.Print i, name, FNr
 Next
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dbtest/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' dbtest
#End If
Function HANrS$(HANr) ' wird z.Zt. nicht verwendet, 7.1.06
 If IsNull(HANr) Then
  HANrS = vNS
 ElseIf InStrB(CStr(HANr), "/") = 0 Then
  HANrS = Left(HANr, 2) + "/" + Right$(HANr, 5)
 Else
  HANrS = CStr(HANr)
 End If
End Function ' HANrs$
Sub HAzählen()
 Dim HA As New ADODB.Recordset, pa As New ADODB.Recordset, ng%, Na%, FText$
' SET HA = Dtb.OpenRecordset(hae, dbOpenTable)
' Call KVÄVorb
'  Call acon(HaT)
 If LenB(DBCn) = 0 Or DBCn = "" Then Call acon(quelleT)
' Call HA.Open("SELECT * FROM `kvaerzte`.`hae`", DBCn, adOpenDynamic, adLockOptimistic) 'haecn
 myFrag HA, "SELECT zahl FROM `kvaerzte`.`hae`"
 If Not HA.BOF Then
'  HA.MoveFirst
  While Not HA.EOF
'   HA.Edit
   HA!Zahl = 0
   HA.Update
   HA.MoveNext
  Wend
 End If
 ng = 0
 Na = 0
 FText = vNS
' SET pa = Dtb.OpenRecordset("Anamnesebogen", dbOpenDynaset)
' pa.Open "anamnesebogen", DBCn, adOpenDynamic, adLockOptimistic
 myFrag pa, "SELECT * FROM anamnesebogen"
 If Not pa.BOF Then
'  pa.MoveFirst
'  HA.index = "KVNr"
  While Not pa.EOF
 '  HA.FindFirst "KVNr = '" & pa!HANr & "'"
   Set HA = Nothing
'   Call HA.Open("SELECT * FROM `hae` WHERE kvnr = '" & Left(REPLACE$(pa!KVNr, "/", ""), 7) & "'", DBCn, adOpenDynamic, adLockOptimistic)
   myFrag HA, "SELECT * FROM `hae` WHERE kvnr = '" & Left(REPLACE$(pa!KVNr, "/", ""), 7) & "'"
   If IsNull(pa!HANr) Or pa!HANr = vNS Then
    Na = Na + 1
   Else
'    HA.Seek "=", LEFT(replace$(pa!HANr, "/", ""), 7)
    If Not HA.BOF Then
'     HA.Edit
     HA!Zahl = HA!Zahl + 1
     HA.Update
    Else
     ng = ng + 1
     FText = FText + vbCr + "  " + GesNamFn(pa) & " " & Left(REPLACE$(pa!HANr, "/", ""), 7)
    End If
   End If
   pa.MoveNext
  Wend
  MsgBox "Hausaerzte gezählt: " + CStr(ng) + " nicht gefunden, " + CStr(Na) + " nicht angegeben." + FText
 End If
End Sub ' HAzählen()

' Achtung: die Abfrage der offenen Briefe wird von Access selbst immer so verändert, daß ein Syntaxfehler entsteht (durch Ersetzen runder klammern der letzten Abfrage durch ``.).
' Hier der originale Text (natürlich ohne Kommentar):
'SELECT `anamnesebogen`.`pat_id` AS id, `nachname`, `vorname`, `gebdat`, `diabetestyp`, `versicherungsart`, `diabetes seit`, `vorgestellt`, format$(br.zeitpunkt,"dd/mm/yy") AS zp, br.name AS brief, `hanr`, hanam AS HA, zahl, haanschrift AS anschrift FROM (`anamnesebogen` LEFT JOIN `SELECT pat_id, zeitpunkt, name FROM `briefe` WHERE instr(name,"rief")>0`. AS br ON `anamnesebogen`.`pat_id`=br.pat_id) LEFT JOIN (SELECT `KVNr`, MIN(`name`) AS HANam, COUNT(name) AS Zahl, MIN(anschrift) AS HAAnschrift FROM `hausaerzte` GROUP BY `KVNr` ORDER BY KVNr) AS ha ON `anamnesebogen`.`hanr`=ha.KVNr ORDER BY `anamnesebogen`.`pat_id`;

' ha, so geht's auch ohne Fehler (1)im inneren SELECT vornamen angegeben, 2)doppelte eckige klammern vermieden):
'SELECT `anamnesebogen`.`pat_id` AS id, `nachname`, `vorname`, `gebdat`, `diabetestyp`, `versicherungsart`, `diabetes seit`, `vorgestellt`, format$(br.zeitpunkt,"dd/mm/yy") AS zp, br.name AS brief, `hanr`, hanam AS HA, zahl, haanschrift AS anschrift FROM (`anamnesebogen` LEFT JOIN `SELECT pat_id, zeitpunkt, name FROM `briefe` WHERE instr(name,"rief")>0`. AS br ON `anamnesebogen`.`pat_id`=br.pat_id) LEFT JOIN `SELECT `hausaerzte`.KVNr, MIN(`hausaerzte`.name) AS HANam, COUNT(`hausaerzte`.name) AS Zahl, MIN(`hausaerzte`.anschrift) AS HAAnschrift FROM `hausaerzte` GROUP BY `hausaerzte`.KVNr`. AS ha ON `anamnesebogen`.`hanr`=ha.KVNr ORDER BY `anamnesebogen`.`pat_id`;
