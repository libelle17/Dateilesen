Attribute VB_Name = "Kompilierbarkeit"
' aus Importiere
Function doFAnfFuell%(raFa As ADODB.Recordset, Optional obgroß%)
 Dim lpid&, aktpid&, Bm$, lAusgst As Date
 Dim raSQL As New ADODB.Recordset
 On Error GoTo fehler
'   raFa.Edit
   If obgroß Then
    raFa!Fanf = CDate(0)
   End If
   If IsNull(raFa!Fanf) Or raFa!Fanf = CDate(0) Then
    If raFa!SchGr = "90" Or raFa!SchGr = "41" Then
     raFa!Fanf = raFa!BhFB
    ElseIf Not IsNull(raFa!ausgst) And raFa!ausgst <> CDate(0) Then
     raFa!Fanf = raFa!ausgst
    End If
    If IsNull(raFa!Fanf) Or raFa!Fanf = CDate(0) Then
     If Not IsNull(raFa!lVorl) And raFa!lVorl <> CDate(0) And DateValue(raFa!lVorl) <= lAusgst Then
      raFa!Fanf = raFa!lVorl
     End If
    End If
    If IsNull(raFa!Fanf) Or raFa!Fanf = CDate(0) Then
     sql = "SELECT * FROM `eintraege` WHERE pat_id = " & CStr(raFa!Pat_id) & " AND zeitpunkt >= " & DatFor_k(Format$(raFa!BhFB, "DD.MM.YY")) & " "
     If Not IsNull(raFa!BhFE1) And raFa!BhFE1 <> CDate(0) Then
      sql = sql + " AND zeitpunkt <= " + DatFor_k(Format$(raFa!BhFE1, "DD.MM.YY")) & " "
     End If
     sql = sql + "ORDER BY zeitpunkt"
'     SET raSQL = Dtb.OpenRecordset(sql)
     Set raSQL = Nothing
'     raSQL.Open sql, DBCn, adOpenDynamic, adLockReadOnly
     myFrag raSQL, sql
     If Not raSQL.BOF Then
      If raSQL!Zeitpunkt > raFa!lVorl And Not IsNull(raFa!lVorl) And raFa!lVorl <> CDate(0) Then
       raFa!Fanf = raFa!lVorl
      Else
       raFa!Fanf = raSQL!Zeitpunkt
      End If
     Else
      If Not IsNull(raFa!lVorl) And raFa!lVorl <> CDate(0) Then
       raFa!Fanf = raFa!lVorl
      ElseIf raFa!Pat_id = 2 And raFa!Nachname Like "Muster*" Then
       raFa!Fanf = #7/1/2004#
      Else
'       Stop ' hartnäckiger Fall
       syscmd acSysCmdSetStatus, "doFAnfFuell: für " & CStr(raFa!Pat_id) & " nur BhFB zu nehmen, hartnäckiger Fall"
       raFa!Fanf = raFa!BhFB
      End If
     End If
    End If
    If raFa.EditMode <> 0 Then raFa.Update
   End If ' ISNULL(rafa!fanf) OR rafa!fand = CDate(0) THEN
   lAusgst = raFa!ausgst
 Exit Function
fehler:
 If Err.Number = 3021 Then doFAnfFuell = 1: Exit Function
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doFAnfFuell/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doFAnfFuell

Function FAnfFuell(Optional obgroß%) ' Fallanfangsfeld befüllen
 Dim raFa As New ADODB.Recordset
 On Error GoTo fehler
' Call dtbInit
' Call raFa.Open("faelle", DBCn, adOpenDynamic, adLockOptimistic)
 myFrag raFa, "SELECT * FROM faelle"
 If Not raFa.BOF Then
  raFa.MoveLast
  Do While Not raFa.BOF
   Call doFAnfFuell(raFa, obgroß)
   raFa.Move -1
  Loop
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FAnfFuell/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FAnfFuell

