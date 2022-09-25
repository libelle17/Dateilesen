Attribute VB_Name = "ViewsErstellen"
Option Explicit
Public numArgs%, ArgFeld$()
Sub Main()
 Dim erg$
 Dim q As DAO.QueryDef
 Dim QTb As DAO.Database
 On Error GoTo fehler
' Database = argfeld(0)
' QName = argfeld(1)
 Call BefZeileAbrufen
 If UBound(ArgFeld) <> 3 Then End
 erg = Dir(ArgFeld(1))
 If erg = "" Then End
  Set QTb = DAO.OpenDatabase(ArgFeld(1))
  If QTb Is Nothing Then End
  On Error Resume Next
'  Call cat.Views.Delete(QName)
  QTb.QueryDefs.Delete ArgFeld(2) 'QName
  Call QTb.CreateQueryDef(ArgFeld(2), ArgFeld(3)) '.CommandText)
  Set QTb = Nothing
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Main/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' main

Function BefZeileAbrufen(Optional MaxArgs)
   ' Variablen deklarieren.
   Dim C$, BefZl$, BefZlLng&, InArg%, I&, ZahlArgs&, inAnfü%
   ' Feststellen, ob MaxArgs angegeben wurde.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   ' Datenfeld passender Größe erstellen.
   ReDim ArgFeld(MaxArgs)
   numArgs = 0: InArg = False
   ' Befehlszeilenargumente abrufen.
   BefZl = Command()
   BefZlLng = Len(BefZl)
   ' Befehlszeile Zeichen für Zeichen
   ' durchgehen.
   For I = 1 To BefZlLng
      C = Mid(BefZl, I, 1)
      ' Auf Leerzeichen oder Tabulatoren prüfen.
      If C = Chr(34) Then
       If I < BefZlLng Then
        If inAnfü And Mid(BefZl, I + 1, 1) = Chr(34) Then
         I = I + 1
         GoTo zähle
        End If
       End If
       inAnfü = Not inAnfü
      Else
       If (C = " " Or C = vbTab) And Not inAnfü Then
         ' Tabulator oder Leerzeichen gefunden.
         ' Das InArg-Attribut auf False festlegen.
         InArg = False
       Else
         ' Weder Leerzeichen noch Tabulatoren.
         ' Überprüfen ob bereits in Argument enthalten.
         If Not InArg Then
         ' Anfang des neuen Arguments.
         ' Überprüfen, ob zu viele Argumente verwendet wurden.
            If ZahlArgs = MaxArgs Then Exit For
            ZahlArgs = ZahlArgs + 1
            InArg = True
         End If
         ' Zeichen an aktuelles Argument anfügen.
zähle:
         ArgFeld(ZahlArgs) = ArgFeld(ZahlArgs) + C
       End If
      End If ' c = chr(34)
   Next I
   ' Größe des Datenfeldes neu bestimmen, so daß es
   ' gerade alle Argumente aufnehmen kann.
   ReDim Preserve ArgFeld(ZahlArgs)
   ' Datenfeld in Funktionsnamen zurückgeben.
   BefZeileAbrufen = ArgFeld()
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BefZeileAbrufen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' BefZeileAbrufen
