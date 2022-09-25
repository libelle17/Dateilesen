Attribute VB_Name = "ZielDBDAO"
Option Explicit
Public Dtb As DAO.Database
Function Tab÷ff(Name$, Optional Ind$, Optional DBs) As DAO.Recordset
 Dim dbsName$
 On Error GoTo fehler
 Call dtbInit
 On Error Resume Next
 dbsName = DBs.Name
 If dbsName = "" Then
  DoCmd.Close acTable, Name
  Err.Clear
  Set Tab÷ff = Dtb.OpenRecordset(Name, dbOpenTable)
  If Err.Number <> 0 Then
   On Error GoTo fehler
   Set Tab÷ff = Dtb.OpenRecordset(Name)
  End If
 Else
  Err.Clear
  Set Tab÷ff = DBs.OpenRecordset(Name, dbOpenTable)
  If Err.Number <> 0 Then
   On Error GoTo fehler
   Set Tab÷ff = DBs.OpenRecordset(Name)
  End If
 End If ' eRR.Number <> 0 Then
 On Error GoTo fehler
 If Ind <> "" Then
  On Error Resume Next
  Tab÷ff.Index = Ind
  On Error GoTo fehler
 End If ' Ind <> 0 Then
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Tab÷ff/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Tab÷ff()


Function dtbInit()
 Dim DtbName$
 Dim dtbCt&
 Dim i%
 On Error GoTo fehler
 If QMdbAkt = "" Then QMdbAkt = LokPfad(QmdB)
' If InStr(CurrentDb.Name, "Anamneseblatt") > 0 Then
'  Set Dtb = CurrentDb
' Else
  On Error Resume Next
  DtbName = Dtb.Name
  dtbCt = Dtb.TableDefs.Count
  If Err.Number <> 0 Or dtbCt = 0 Then
   Err.Clear
   For i = DBEngine.Workspaces(0).Databases.Count - 1 To i Step -1
    If UCase(QMdbAkt) = UCase(DBEngine.Workspaces(0).Databases(i).Name) Then
     Set Dtb = DBEngine.Workspaces(0).Databases(i)
     GoTo hamma
    End If
   Next i
   Set Dtb = DBEngine.OpenDatabase(QMdbAkt)
hamma:
   If Err.Number = 3043 Then ' Datentr‰ger- oder Netzwerkfehler
    Debug.Print "Datentr‰ger- oder Netzwerkfehler in dtbInit"
    End
   End If
'  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dtbInit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dtbInit

