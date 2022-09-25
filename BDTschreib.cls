VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BDTSchreib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private t$(), zl%, ZZ&
Private i&, lenge&
Public mVerz$, DMPImp$, ZDat$
Public ¸zpt As Date ' ‹bertragungszeitpunkt

Private Sub Class_Initialize()
    ' create the t() array
    ZZ = -1
    ¸zpt = Now()
End Sub
Private Function callVerzPr¸f(Verz$)
 Dim erg$
 On Error GoTo fehler
 mVerz = Verz
 erg = VerzPr¸f(mVerz)
 If LenB(erg) <> 0 Then
  MsgBox "Konnte Verzeichnis '" & erg & "' nicht erstellen!"
  Exit Function
 End If
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler IN callVerzPr¸f/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' callVerzPr¸f
'ˆffnet BDT-Exportdatei mit Handle 327 und Namen mit Datum und praxisinternem Arzt
Public Function Start%(ByRef Verz$, ByRef Vorsilbe$, Optional Arztnr&)
 Dim erg$
 Call callVerzPr¸f(Verz)
 ZDat = Vorsilbe & " " & Format$(¸zpt, "dd/mm/yy HH.MM") & IIf(Arztnr = 1, " Schade", IIf(Arztnr = 2, " Kothny", "")) & ".BDT"
 DMPImp = hVerz & ZDat
 Open DMPImp For Output As #327
 Close #327
 erg = Dir(hVerz & ZDat)
 If LenB(erg) = 0 Then
  MsgBox "Fehler beim Erstellen der Exportdatei '" & hVerz & ZDat & "'"
  Exit Function
 End If
 Start = -1
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler IN Start/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function 'Start
Public Function ImportFolderHerricht(Optional Verz$, Optional DateiName$)
 Dim fld As Folder, Fil As File, FPath$
 Dim errn&
 On Error GoTo fehler
 If LenB(Verz) <> 0 Then
  Call callVerzPr¸f(Verz)
 End If
 If LenB(DateiName) <> 0 Then
  DMPImp = hVerz & DateiName
 End If
 On Error Resume Next
 Set fld = FSO.GetFolder(mVerz)
 errn = Err.Number
 On Error GoTo fehler
 If errn <> 0 Then
  Set fld = FSO.GetFolder(REPLACE$(LCase$(mVerz), "u:", "\\linux1\daten\eigene Dateien"))
 End If
 For Each Fil In fld.Files
  If Right$(Fil.name, 4) = ".BDT" And UCase$(Fil.path) <> UCase$(DMPImp) Then
   FPath = Fil.path
   On Error Resume Next
   Name FPath As REPLACE$(FPath, ".BDT", ".txt")
   Kill FPath
   On Error GoTo fehler
  End If
 Next Fil
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler IN ImportFolderHerricht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ImportFolderHerricht()
Public Sub Add(ByRef Schl¸sselundInhalt$)
 ZZ = ZZ + 1
 If ZZ = 0 Then
  ReDim t(0)
 Else
  ReDim Preserve t(ZZ)
 End If
 zl = 5 + Len(Schl¸sselundInhalt)
 t(ZZ) = Format(zl, "000")
 If Len(t(ZZ)) > 3 Then
  t(ZZ) = Right(t(ZZ), 3)
 End If
 t(ZZ) = t(ZZ) & ZSU(Schl¸sselundInhalt)
 If Left$(Schl¸sselundInhalt, 4) = "8000" Then
  ZZ = ZZ + 1
  ReDim Preserve t(ZZ)
 End If
End Sub ' Add(ByRef Schl¸sselundInhalt$)
Public Sub DAdd(ByRef Schl¸ssel$, ByRef Datum)
 On Error GoTo fehler
 If Datum = 0 Then
  Call Add(Schl¸ssel & "00000000")
 Else
  Call Add(Schl¸ssel & Format$(Datum, "ddmmyyyy"))
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler IN Add/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Sub DAdd(ByRef Schl¸ssel$, ByRef Datum)
Public Sub TAdd(ByRef Schl¸ssel$, ByRef Datum)
 Call Add(Schl¸ssel & Format$(Datum, "HHMM"))
End Sub
Public Sub FFAdd(ByRef Schl¸ssel$) ' Formularfeld
 Call Add("6298" & Schl¸ssel)
End Sub
Public Sub FIAdd(ByRef Schl¸ssel$) ' FormularfeldInhalt
 Call Add("6299" & Schl¸ssel)
End Sub
Public Sub SAdd(ByRef Schl¸ssel$, ByRef Roh$)
' Dim Verarb$
' Verarb = ZSU(Roh)
 Call Add(Schl¸ssel & Roh)
End Sub
Public Sub Satzart(ByRef Satzart$)
 Call Add("8000" & Satzart)
End Sub
Public Sub PatID(ByRef PatID&)
 Call Add("3000" & CStr(PatID))
End Sub
Public Sub NVorsatz(ByRef NV$)
 Call SAdd("3100", NV)
End Sub
Public Sub Nachname(ByRef NN$)
 Call SAdd("3101", NN)
End Sub
Public Sub Vorname(ByRef VN$)
 Call SAdd("3102", VN)
End Sub
Public Sub Geb(ByRef GebDat As Date)
 Dim FGeb$
 FGeb = Format$(GebDat, "ddmmyyyy")
 If Mid$(FGeb, 5, 2) = "20" Then
  If GebDat > Date Then FGeb = Left(FGeb, 4) & "19" & Right$(FGeb, 2)
 End If
 Add ("3103" & FGeb)
End Sub
Public Sub Arztnummer(ByRef nr$)
 Call Add("9100" & nr)
End Sub
Public Sub Erstellungsdatum(ByRef Datum As Date)
 Call DAdd("9103", Datum)
End Sub

Public Sub Schreib(Optional anh‰ngen%)
 Dim errn&
 For i = ZZ To 0 Step -1
  lenge = lenge + Len(t(i)) + 2
  If Left(t(i), 7) = "0138000" Then
   lenge = lenge + 12
   t(i + 1) = "0148100" & Format(lenge, "00000")
   lenge = 0
  End If
 Next i
 If anh‰ngen Then
  On Error Resume Next
  Open DMPImp For Append As #327
  errn = Err.Number
  On Error GoTo 0
  If errn <> 0 Then Open REPLACE$(LCase$(DMPImp), "u:", "\\linux1\daten\eigene Dateien") For Append As #327
 Else
  Open DMPImp For Output As #327
 End If
 For i = 0 To ZZ
  Print #327, t(i)
 Next i
 Close #327
End Sub ' Schreib

Public Sub BDTKopf()
  Call Satzart("0020") ' Satzart
  Call Arztnummer("6419153")  ' Arztnummer des Absenders
  Call Erstellungsdatum(¸zpt) ' Erstellungsdatum
  Call Add("9105001") ' Ordnungsnummer Datentr‰ger (Header) des DP
  Call Add("91062") ' verwendeter Zeichensatz (2=IBM-Code)
  Call Satzart("0022") ' Satzart
  Call Add("921001/99") ' Version ADT-Satzbeschreibung
  Call Add("921302/94") ' Version BDT-Satzbeschreibung
  Call Add("96001") ' Archivierungsart
  Call Add("960101011980" & Format$(¸zpt, "ddmmyyyy")) 'Zeitraum der Spiegelung
  Call Add("9602" & Format(¸zpt, "hhmmss") & "00") ' Beginn der ‹bertragung
  Call Satzart("besa") ' ("0010") 'Satzarzt
'  Call Add("0101A001011") ' KBV-Pr¸fnummer
'  Call Add("0102TurboMed EDV GmbH") ' Softwareverantwortlicher
'  Call Add("0103TurboMed@Windows") ' Software
'  Call Add("0104IBM PC/AT")
'  Call Add("02021") ' Praxistyp (1=Einzelpraxis)
''  Call Add("8000besa")
  Call Add("0201641915300") ' Arztnummer
  Call Add("0203Diabetologische Gemeinschaftspraxis") ' Arztname
  Call Add("0212889690003") ' Lebenslange Arztnummer
  Call Add("0211Gerald Schade") 'Arztname f¸r Leistungsdifferenzierung
  Call Add("0971TM#0") '
  Call Add("0212933284903") ' Lebenslange Arztnummer
  Call Add("0211Dr.med. Thomas Kothny") 'Arztname f¸r Leistungsdifferenzierung
  Call Add("0971TM#0") '
'  Call Add("0204TG Diabetologie") ' FA Innere und Allgemeinmedizin (Hausarzt)"
  Call SAdd("0205", "Mittermayerstraﬂe 13") ' Straﬂe der Praxis
  Call Add("021585221") ' PLZ der Praxisadresse
  Call Add("0216Dachau") ' Ort der Praxisadresse
  Call Add("020808131 / 616 380") ' Telefonnummer der Praxis
  Call Add("020908131 / 616 381") ' Telefaxnummer der Praxis
End Sub ' BDTKopf

Public Function ZSU$(ByRef Quelle$)
Dim i%, ZS$, bq$, bz$
 On Error GoTo fehler
ZS = vNS
For i = 1 To Len(Quelle)
 bq = Mid$(Quelle, i, 1)
 Select Case bq
  Case "ﬂ"
   bz = "·"
  Case "‰"
   bz = "Ñ"
  Case "ˆ"
   bz = "î"
  Case "¸"
   bz = "Å"
  Case "ƒ"
   bz = "é"
  Case "÷"
   bz = "ô"
  Case "‹"
   bz = "ö"
  Case "·"
   bz = "†"
  Case "‡"
   bz = "Ö"
  Case "È"
   bz = "Ç"
  Case "Ë"
   bz = "ä"
  Case "Ì"
   bz = "°"
  Case "Ï"
   bz = "ç"
  Case "Û"
   bz = "¢"
  Case "Ú"
   bz = "ï"
  Case "˙"
   bz = "£"
  Case "˘"
   bz = "ó"
  Case "‚"
   bz = "É"
  Case "Í"
   bz = "à"
  Case "Ó"
   bz = "å"
  Case "Ù"
   bz = "ì"
  Case "˚"
   bz = "ñ"
  Case "µ"
   bz = "Ê"
  Case "≤"
   bz = "˝"
  Case "≥"
   bz = "¸"
  Case "ß"
   bz = "ı"
  Case "Ä"
   bz = "_"
  Case Else
   bz = bq
 End Select
 ZS = ZS + bz
Next i
ZSU = ZS
If ZSU <> Quelle Then
 Debug.Print "ZSU: " + ZSU + " -> " + Quelle
End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler IN ZSU/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZSU
' liefert vns, falls Erfolg, oder das nicht zu erstellende Verzeichnis
Public Function VerzPr¸f$(ByRef Verz$)
 Dim Bstd$(), i%, j%, k%, tStr$
 Dim FSO As New FileSystemObject
 On Error GoTo fehler
 tStr = FSO.GetAbsolutePathName(Verz)
 Bstd = Split(tStr, "\")
 For j = 0 To UBound(Bstd)
  tStr = vNS
  For i = 0 To j
   tStr = tStr + IIf(i = 0, vNS, "\") + Bstd(i)
  Next i
  If tStr <> "" Then
   If Not FSO.FolderExists(tStr) Then
    If FSO.FolderExists(REPLACE$(LCase$(tStr), "u:", "\\linux1\daten\eigene Dateien")) Then VerzPr¸f = "": GoTo weiter
    If InStrB(Mid$(tStr, 3), "\") <> 0 Then 'And Not LEFT(TStr, 2) = "\\" AND InStr(mid$(TStr, 3), "\") = 1 THEN
     Call FSO.CreateFolder(tStr)
    End If
    If Not (Left$(tStr, 2) = "\\" And Not InStr(Mid(tStr, 3), "\")) Then
     If Not FSO.FolderExists(tStr) Then
      VerzPr¸f = tStr
     End If
    End If
   End If
  End If
weiter:
 Next
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler IN VerzPr¸f/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' VerzPr¸f(ByVal Verz$)

