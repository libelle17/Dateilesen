Attribute VB_Name = "AccDBMach"
Option Explicit
#Const ohneDAO = True
' folgendes wurde manuell hinzugefügt:
'Public WS As DAO.Workspace, DBE As DAO.DBEngine
' folgendes auskommentiert:
'Const ZielDbS$ = "u:\Anamnese\Quelle.mdb"
Const ß$ = ""
Const tRn$ = "tmpRelationen"
'Achtung: Dies ist ein Kompilat. Geändert werden sollte allenfalls, was über diesen Zeilen steht.
'Zweck ist das (Wieder)herstellen von Tabllenstrukturen, Indices und Beziehungen.
'Im Steuermodul sollten zum Customizen nur die Kopfkonstanten und die Function rufCreate() geändert werden
'Durch Aufruf derselben wird dieses Modul erstellt, was wiederholbar bleiben sollte.
'Durch Aufruf von Function rufAlleCreate() wird ein Modul "MachAlleTab" für alle Datenbanktabellen gefüllt.
'Aufrufbar sind dort u.a. alle Funktionen, die mit 'mach' beginnen.
'Achtung: Tabelleninhalte werden ohne Rückfrage gelöscht!!!!
'Autor Gerald Schade, 14.8.05
Dim mitBez% ' True heißt mit Löschen und Wiederherstellen von Beziehungen
Dim gesfehler%
Dim SammelAufruf%
#If Not ohneDAO Then
Function MachAVLassen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "VLassen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Datum]  DATE"
 sql = sql + ", [Name]  TEXT(255)"
 sql = sql + ", [Größe]  INTEGER"
 sql = sql + ", [pfad]  TEXT(255)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "", 0, Null, 4080, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Größe", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Pfad].AllowZeroLength = True
 Call CPropA(QTb + ß, "pfad", "", 0, Null, 2760, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Datum].Name) = "DATUM" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Datum] on [" + QTb + ß + "] ("
  sql = sql + "[Datum] DESC"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Name].Name) = "NAME" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Name] on [" + QTb + ß + "] ("
  sql = sql + "[Name]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAVLassen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAVLassen

Function MachAUnbekannte_Kennungen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Unbekannte Kennungen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Kennung]  TEXT(4) CONSTRAINT [Kennung] UNIQUE"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 .Fields![Kennung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kennung", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "zugehöriger Patient für spätere Ermittlungen", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Kennung].Name) = "KENNUNG" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [Kennung] on [" + QTb + ß + "] ("
  sql = sql + "[Kennung]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Pat_ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAUnbekannte_Kennungen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAUnbekannte_Kennungen

Function MachAtmpRelationen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "tmpRelationen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ccolumn]  INTEGER"
 sql = sql + ", [grbit]  INTEGER"
 sql = sql + ", [icolumn]  INTEGER"
 sql = sql + ", [szColumn]  TEXT(255)"
 sql = sql + ", [szObject]  TEXT(255) NOT NULL"
 sql = sql + ", [szReferencedColumn]  TEXT(255)"
 sql = sql + ", [szReferencedObject]  TEXT(255) NOT NULL"
 sql = sql + ", [szRelationship]  TEXT(255) NOT NULL"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 .Fields![szcolumn].AllowZeroLength = True
 Call PropA(QTb + ß, "szColumn", "UniCodeCompression", 1, 0, ZielDB)
 .Fields![szobject].AllowZeroLength = True
 Call PropA(QTb + ß, "szObject", "UniCodeCompression", 1, 0, ZielDB)
 .Fields![szreferencedcolumn].AllowZeroLength = True
 Call PropA(QTb + ß, "szReferencedColumn", "UniCodeCompression", 1, 0, ZielDB)
 .Fields![szreferencedobject].AllowZeroLength = True
 Call PropA(QTb + ß, "szReferencedObject", "UniCodeCompression", 1, 0, ZielDB)
 .Fields![szrelationship].AllowZeroLength = True
 Call PropA(QTb + ß, "szRelationship", "UniCodeCompression", 1, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAtmpRelationen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAtmpRelationen

Function MachATabTeil()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "TabTeil"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[pat_id]  INTEGER"
 sql = sql + ", [Datum]  TEXT(8)"
 sql = sql + ", [Dokument]  TEXT(128)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "pat_id", "3000", Null, 109, -1, 0, 0, ZielDB)
 .Fields![Datum].AllowZeroLength = True
 Call CPropA(QTb + ß, "Datum", "6200", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Dokument].AllowZeroLength = True
 Call CPropA(QTb + ß, "Dokument", "6325", 0, 109, 7395, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [pat_id] on [" + QTb + ß + "] ("
  sql = sql + "[pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachATabTeil/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachATabTeil

Function MachATabProp()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "TabProp"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Prop]  TEXT(50)"
 sql = sql + ", [ReiheVorher]  INTEGER"
 sql = sql + ", [TypVorher]  INTEGER"
 sql = sql + ", [WertVorher]  TEXT(50)"
 sql = sql + ", [ReiheNachher]  INTEGER"
 sql = sql + ", [TypNachher]  INTEGER"
 sql = sql + ", [WertNachher]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 .Fields![Prop].AllowZeroLength = True
 Call CPropA(QTb + ß, "Prop", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ReiheVorher", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "TypVorher", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields![WertVorher].AllowZeroLength = True
 Call CPropA(QTb + ß, "WertVorher", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ReiheNachher", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "TypNachher", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields![WertNachher].AllowZeroLength = True
 Call CPropA(QTb + ß, "WertNachher", "", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Prop].Name) = "PROP" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [prop] on [" + QTb + ß + "] ("
  sql = sql + "[Prop]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachATabProp/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachATabProp

Function MachATabGanz()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "TabGanz"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[pat_id]  INTEGER"
 sql = sql + ", [Datum]  TEXT(8)"
 sql = sql + ", [Dokument]  TEXT(128)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "pat_id", "3000", Null, 109, -1, 0, 0, ZielDB)
 .Fields![Datum].AllowZeroLength = True
 Call CPropA(QTb + ß, "Datum", "6200", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Dokument].AllowZeroLength = True
 Call CPropA(QTb + ß, "Dokument", "6325", 0, 109, 7905, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [pat_id] on [" + QTb + ß + "] ("
  sql = sql + "[pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachATabGanz/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachATabGanz

Function MachATabellenfuellung()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Tabellenfuellung"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Tabname]  TEXT(50) CONSTRAINT [Tabname] PRIMARY KEY"
 sql = sql + ", [T0501050310]  INTEGER"
 sql = sql + ", [T0501090240]  INTEGER"
 sql = sql + ", [T0504161552]  INTEGER"
 sql = sql + ", [T0507240436]  INTEGER"
 sql = sql + ", [T0507311225]  INTEGER"
 sql = sql + ", [T0508141516]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 .Fields![TabName].AllowZeroLength = True
 Call CPropA(QTb + ß, "Tabname", "", 0, 109, 2490, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "T0501050310", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "T0501090240", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "T0504161552", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "T0507240436", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "T0507311225", "", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachATabellenfuellung/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachATabellenfuellung

Function MachAStand()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Stand"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Stand]  TEXT(7) CONSTRAINT [Stand] PRIMARY KEY"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Stand", "", 0, 109, 885, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "", Null, Null, 1920, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAStand/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAStand

Function MachARRParse()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "RRParse"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Pat_id]  INTEGER"
 sql = sql + ", [Zeitpunkt]  DATE"
 sql = sql + ", [RRSyst]  SMALLINT"
 sql = sql + ", [RRDiast]  SMALLINT"
 sql = sql + ", [Quelle]  TEXT(10)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zeitpunkt", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "RRSyst", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "RRDiast", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields![Quelle].AllowZeroLength = True
 Call CPropA(QTb + ß, "Quelle", "", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![id].Name) = "ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ",[Zeitpunkt]"
  sql = sql + ",[RRSyst]"
  sql = sql + ",[RRDiast]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachARRParse/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachARRParse

Function MachARR()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "RR"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [RR]  MEMO"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 465, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "6200 + 6201", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![RR].AllowZeroLength = True
 Call CPropA(QTb + ß, "RR", "6230", 0, Null, 12510, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in BDT-Datei", Null, 109, 825, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt] DESC"
  sql = sql + ",[RR]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachARR/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachARR

Function MachARezeptEintraege()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "RezeptEintraege"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [Rezept]  TEXT(3)"
 sql = sql + ", [Rezeptklasse]  TEXT(5)"
 sql = sql + ", [Medikament]  TEXT(50)"
 sql = sql + ", [PZN]  TEXT(20)"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [QS]  TEXT(5)"
 sql = sql + ", [QT]  TEXT(5)"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "6200 + 6201", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Rezept].AllowZeroLength = True
 Call CPropA(QTb + ß, "Rezept", "6210, 3652(1), 6218(1)", 0, 109, 810, 0, 0, ZielDB)
 .Fields![Rezeptklasse].AllowZeroLength = True
 Call CPropA(QTb + ß, "Rezeptklasse", "6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)", 0, 109, 1395, 0, 0, ZielDB)
 .Fields![Medikament].AllowZeroLength = True
 Call CPropA(QTb + ß, "Medikament", "3652(2), 6218(4)", 0, 109, 5445, 0, 0, ZielDB)
 .Fields![PZN].AllowZeroLength = True
 Call CPropA(QTb + ß, "PZN", "6210(2), 6218(3)", 0, 109, 885, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in BDT-Datei", Null, 109, 825, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![QS].AllowZeroLength = True
 Call CPropA(QTb + ß, "QS", "Quartal des Behandlungsfallbeginns sortiert", 0, 109, -1, 0, 0, ZielDB)
 .Fields![QT].AllowZeroLength = True
 Call CPropA(QTb + ß, "QT", "Quartal des Behandlungsfallbeginns", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Statusbyte", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt]"
  sql = sql + ",[Rezept]"
  sql = sql + ",[Medikament]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachARezeptEintraege/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachARezeptEintraege

Function MachARelationen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Relationen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ccolumn]  INTEGER"
 sql = sql + ", [grbit]  INTEGER"
 sql = sql + ", [icolumn]  INTEGER"
 sql = sql + ", [szColumn]  TEXT(255)"
 sql = sql + ", [szObject]  TEXT(255) CONSTRAINT [szObject] NOT NULL"
 sql = sql + ", [szReferencedColumn]  TEXT(255)"
 sql = sql + ", [szReferencedObject]  TEXT(255) CONSTRAINT [szReferencedObject] NOT NULL"
 sql = sql + ", [szRelationship]  TEXT(255) CONSTRAINT [szRelationship] NOT NULL"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 .Fields![szcolumn].AllowZeroLength = True
 Call PropA(QTb + ß, "szColumn", "UniCodeCompression", 1, 0, ZielDB)
 .Fields![szobject].AllowZeroLength = True
 Call PropA(QTb + ß, "szObject", "UniCodeCompression", 1, 0, ZielDB)
 .Fields![szreferencedcolumn].AllowZeroLength = True
 Call PropA(QTb + ß, "szReferencedColumn", "UniCodeCompression", 1, 0, ZielDB)
 .Fields![szreferencedobject].AllowZeroLength = True
 Call PropA(QTb + ß, "szReferencedObject", "UniCodeCompression", 1, 0, ZielDB)
 .Fields![szrelationship].AllowZeroLength = True
 Call PropA(QTb + ß, "szRelationship", "UniCodeCompression", 1, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![szobject].Name) = "SZOBJECT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [szObject] on [" + QTb + ß + "] ("
  sql = sql + "[szObject]"
  sql = sql + ") WITH IGNORE NULL"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![szreferencedobject].Name) = "SZREFERENCEDOBJECT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [szReferencedObject] on [" + QTb + ß + "] ("
  sql = sql + "[szReferencedObject]"
  sql = sql + ") WITH IGNORE NULL"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![szrelationship].Name) = "SZRELATIONSHIP" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [szRelationship] on [" + QTb + ß + "] ("
  sql = sql + "[szRelationship]"
  sql = sql + ") WITH IGNORE NULL"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachARelationen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachARelationen

Function MachARel2()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Rel2"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ccolumn]  INTEGER"
 sql = sql + ", [grbit]  INTEGER"
 sql = sql + ", [icolumn]  INTEGER"
 sql = sql + ", [szColumn]  TEXT(255)"
 sql = sql + ", [szObject]  TEXT(255)"
 sql = sql + ", [szReferencedColumn]  TEXT(255)"
 sql = sql + ", [szReferencedObject]  TEXT(255)"
 sql = sql + ", [szRelationship]  TEXT(255)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ccolumn", "", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "grbit", "", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "icolumn", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![szcolumn].AllowZeroLength = True
 Call CPropA(QTb + ß, "szColumn", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![szobject].AllowZeroLength = True
 Call CPropA(QTb + ß, "szObject", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![szreferencedcolumn].AllowZeroLength = True
 Call CPropA(QTb + ß, "szReferencedColumn", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![szreferencedobject].AllowZeroLength = True
 Call CPropA(QTb + ß, "szReferencedObject", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![szrelationship].AllowZeroLength = True
 Call CPropA(QTb + ß, "szRelationship", "", 0, Null, 2700, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![szobject].Name) = "SZOBJECT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [szObject] on [" + QTb + ß + "] ("
  sql = sql + "[szObject]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![szreferencedobject].Name) = "SZREFERENCEDOBJECT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [szReferencedObject] on [" + QTb + ß + "] ("
  sql = sql + "[szReferencedObject]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![szrelationship].Name) = "SZRELATIONSHIP" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [szRelationship] on [" + QTb + ß + "] ("
  sql = sql + "[szRelationship]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachARel2/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachARel2

Function MachAQueries()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Queries"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Name]  TEXT(100)"
 sql = sql + ", [sql]  MEMO"
 sql = sql + ", [gespeichert]  DATE"
 sql = sql + ", CONSTRAINT [NamGesp] PRIMARY KEY ( Name, gespeichert)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Name der SQL-Abfrage", 0, 109, 6030, 0, 0, ZielDB)
 .Fields![sql].AllowZeroLength = True
 Call CPropA(QTb + ß, "sql", "SQL-String, der geht", 0, Null, 17550, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "gespeichert", "Zeitpunkt, zu der er geht", Null, Null, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![NamGesp].Name) = "NAMGESP" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [NamGesp] on [" + QTb + ß + "] ("
  sql = sql + "[Name]"
  sql = sql + ",[gespeichert] DESC"
  sql = sql + ") WITH PRIMARY"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![sql].Name) = "SQL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [sql] on [" + QTb + ß + "] ("
  sql = sql + "[sql]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAQueries/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAQueries

Function MachAPauschalen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Pauschalen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Leistung]  TEXT(50)"
 sql = sql + ", [Betreuung]  BIT"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 .Fields![Leistung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leistung", "", 0, 109, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Betreuung", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Betreuung", "", Null, 106, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAPauschalen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAPauschalen

Function MachANamen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Namen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Pat_ID]  INTEGER CONSTRAINT [PAT_ID] UNIQUE"
 sql = sql + ", [lfdnr]  INTEGER"
 sql = sql + ", [NVorsatz]  TEXT(10)"
 sql = sql + ", [Nachname]  TEXT(50)"
 sql = sql + ", [Vorname]  TEXT(50)"
 sql = sql + ", [GebDat]  DATE"
 sql = sql + ", [Straße]  TEXT(50)"
 sql = sql + ", [KVKStatus]  TEXT(1)"
 sql = sql + ", [Geschlecht]  TEXT(1)"
 sql = sql + ", [Plz]  TEXT(20)"
 sql = sql + ", [Ort]  TEXT(70)"
 sql = sql + ", [Weggeldzone]  TEXT(7)"
 sql = sql + ", [AufnDat]  DATE"
 sql = sql + ", [intZOGP]  TEXT(50)"
 sql = sql + ", [Titel]  TEXT(39)"
 sql = sql + ", [Versichertennummer]  TEXT(30)"
 sql = sql + ", [PrivatTel]  TEXT(100)"
 sql = sql + ", [KVNr]  TEXT(10)"
 sql = sql + ", [PrivatTel_2]  TEXT(50)"
 sql = sql + ", [PrivatFax]  TEXT(50)"
 sql = sql + ", [DienstTel]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "lfdnr", "laufende Patientennummer", Null, 109, 585, 0, 0, ZielDB)
 .Fields![NVorsatz].AllowZeroLength = True
 Call CPropA(QTb + ß, "NVorsatz", "3100", 0, 109, 990, 0, 0, ZielDB)
 .Fields![Nachname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Nachname", "3101", 0, 109, 1365, 0, 0, ZielDB)
 .Fields![Vorname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vorname", "3102", 0, 109, 1845, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "GebDat", "3103", Null, Null, 1110, 0, 0, ZielDB)
 .Fields![Straße].AllowZeroLength = True
 Call CPropA(QTb + ß, "Straße", "3107", 0, 109, 2460, 0, 0, ZielDB)
 .Fields![KVKStatus].AllowZeroLength = True
 Call CPropA(QTb + ß, "KVKStatus", "3108", 0, 109, 1200, 0, 0, ZielDB)
 .Fields![Geschlecht].AllowZeroLength = True
 Call CPropA(QTb + ß, "Geschlecht", "3110", 0, 109, 1200, 0, 0, ZielDB)
 .Fields![Plz].AllowZeroLength = True
 Call CPropA(QTb + ß, "Plz", "3112", 0, 109, 675, 0, 0, ZielDB)
 .Fields![Ort].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ort", "3113", 0, 109, 2190, 0, 0, ZielDB)
 .Fields![Weggeldzone].AllowZeroLength = True
 Call CPropA(QTb + ß, "Weggeldzone", "3631", 0, 109, 1470, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AufnDat", "3610", Null, Null, 1110, 0, 0, ZielDB)
 .Fields![intZOGP].AllowZeroLength = True
 Call CPropA(QTb + ß, "intZOGP", "3635, interne Zuordnung Arzt bei GP", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Titel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Titel", "3104", 0, 109, 570, 0, 0, ZielDB)
 .Fields![Versichertennummer].AllowZeroLength = True
 Call CPropA(QTb + ß, "Versichertennummer", "3105", 0, 109, 2145, 0, 0, ZielDB)
 .Fields![PrivatTel].AllowZeroLength = True
 Call CPropA(QTb + ß, "PrivatTel", "3629", 0, 109, 2985, 0, 0, ZielDB)
 .Fields![KVNr].AllowZeroLength = True
 Call CPropA(QTb + ß, "KVNr", "3630", 0, 109, 885, 0, 0, ZielDB)
 .Fields![PrivatTel_2].AllowZeroLength = True
 Call CPropA(QTb + ß, "PrivatTel_2", "3629", 0, 109, 1905, 0, 0, ZielDB)
 .Fields![PrivatFax].AllowZeroLength = True
 Call CPropA(QTb + ß, "PrivatFax", "3629", 0, 109, 1365, 0, 0, ZielDB)
 .Fields![DienstTel].AllowZeroLength = True
 Call CPropA(QTb + ß, "DienstTel", "3629", 0, 109, 1425, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("PrivatMobil", 10, 50)
 .Fields![PrivatMobil].AllowZeroLength = True
 Call CPropA(QTb + ß, "PrivatMobil", "3629", 0, 109, 1365, 0, 0, ZielDB)
 .Fields.Append .CreateField("Email", 10, 100)
 .Fields![Email].AllowZeroLength = True
 Call CPropA(QTb + ß, "Email", "Email", 0, 109, 3330, 0, 0, ZielDB)
 .Fields.Append .CreateField("AnAllgda", 1, 1)
 Call PropA(QTb + ß, "AnAllgda", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AnAllgda", "Anamnese allgemein da", Null, 109, 1020, 0, 0, ZielDB)
 .Fields.Append .CreateField("An1da", 1, 1)
 Call PropA(QTb + ß, "An1da", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "An1da", "Anamnese S.1 da", Null, 109, 750, 0, 0, ZielDB)
 .Fields.Append .CreateField("An2da", 1, 1)
 Call PropA(QTb + ß, "An2da", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "An2da", "Anamnese S.2 da", Null, 109, 750, 0, 0, ZielDB)
 .Fields.Append .CreateField("Checkda", 1, 1)
 Call PropA(QTb + ß, "Checkda", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Checkda", "Checkliste da", Null, 109, 975, 0, 0, ZielDB)
 .Fields.Append .CreateField("DMTypaD", 10, 1)
 .Fields![DMTypaD].AllowZeroLength = True
 Call CPropA(QTb + ß, "DMTypaD", "aus Diagnosen", 0, 109, 1050, 0, 0, ZielDB)
 .Fields.Append .CreateField("AktZeit", 8, 8)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 .Fields.Append .CreateField("absPos", 4, 4)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("StByte", 4, 4)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Cave", 10, 100)
 .Fields![Cave].AllowZeroLength = True
 Call CPropA(QTb + ß, "Cave", "3654", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Notiz", 10, 150)
 .Fields![Notiz].AllowZeroLength = True
 Call CPropA(QTb + ß, "Notiz", "3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>'", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("zubenach", 10, 45)
 .Fields![zubenach].AllowZeroLength = True
 Call CPropA(QTb + ß, "zubenach", "3633", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Verwandt", 10, 77)
 .Fields![Verwandt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Verwandt", "3632", 0, 109, -1, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Nachname]"
  sql = sql + ",[Vorname]"
  sql = sql + ",[GebDat]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [PAT_ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachANamen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachANamen

Function MachAMedPlan()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "MedPlan"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [MPNr]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Medikament]  TEXT(75)"
 sql = sql + ", [MedAnfang]  TEXT(35)"
 sql = sql + ", [FeldNr]  SMALLINT"
 sql = sql + ", [mo]  TEXT(10)"
 sql = sql + ", [mi]  TEXT(10)"
 sql = sql + ", [nm]  TEXT(10)"
 sql = sql + ", [ab]  TEXT(10)"
 sql = sql + ", [zn]  TEXT(10)"
 sql = sql + ", [bBed]  BIT"
 sql = sql + ", [Bemerkung]  MEMO"
 sql = sql + ", [AbsPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 675, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "MPNr", "Ordnungsziffer für Medikamentenplan", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "Zeitpunkt, der Speicherung im Turbomed", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Zeitpunkt aus dem Kopf des Medikamentenplans", Null, Null, 765, 0, 0, ZielDB)
 .Fields![Medikament].AllowZeroLength = True
 Call CPropA(QTb + ß, "Medikament", "", 0, 109, 5595, 0, 0, ZielDB)
 .Fields![MedAnfang].AllowZeroLength = True
 Call CPropA(QTb + ß, "MedAnfang", "", 0, 109, 1245, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "FeldNr", "", Null, 109, 780, 0, 0, ZielDB)
 .Fields![mo].AllowZeroLength = True
 Call CPropA(QTb + ß, "mo", "", 0, 109, 450, 0, 0, ZielDB)
 .Fields![mi].AllowZeroLength = True
 Call CPropA(QTb + ß, "mi", "", 0, 109, 525, 0, 0, ZielDB)
 .Fields![nm].AllowZeroLength = True
 Call CPropA(QTb + ß, "nm", "", 0, 109, 450, 0, 0, ZielDB)
 .Fields![ab].AllowZeroLength = True
 Call CPropA(QTb + ß, "ab", "", 0, 109, 390, 0, 0, ZielDB)
 .Fields![Zn].AllowZeroLength = True
 Call CPropA(QTb + ß, "zn", "", 0, 109, 375, 0, 0, ZielDB)
 Call PropA(QTb + ß, "bBed", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "bBed", "", Null, 109, 645, 0, 0, ZielDB)
 .Fields![Bemerkung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bemerkung", "", 0, Null, 1245, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AbsPos", "Zeile in der BDT-Datei", Null, 109, 840, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt] DESC"
  sql = sql + ",[FeldNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![MedPlanMedikament].Name) = "MEDPLANMEDIKAMENT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [MedPlanMedikament] on [" + QTb + ß + "] ("
  sql = sql + "[Medikament]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![MPNr].Name) = "MPNR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [MPNr] on [" + QTb + ß + "] ("
  sql = sql + "[MPNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAMedPlan/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAMedPlan

Function MachAMedArten()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "MedArten"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Medikament]  TEXT(50) CONSTRAINT [Medikament] UNIQUE"
 sql = sql + ", [Langname]  TEXT(150)"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [Anzahl]  INTEGER"
 sql = sql + ", [Glib]  BIT"
 sql = sql + ", [Metf]  BIT"
 sql = sql + ", [GlucI]  BIT"
 sql = sql + ", [SHGlin]  BIT"
 sql = sql + ", [Glit]  BIT"
 sql = sql + ", [Ins]  BIT"
 sql = sql + ", [Anal]  BIT"
 sql = sql + ", [InsArt]  TEXT(1)"
 sql = sql + ", [HMG]  BIT"
 sql = sql + ", [Hypt]  BIT"
 sql = sql + ", [Thro]  BIT"
 sql = sql + ", [Antib]  BIT"
 sql = sql + ", [and]  BIT"
 sql = sql + ", [hinzugefügt]  DATE"
 sql = sql + ", [Tstr]  BIT"
 sql = sql + ", [Puzu]  BIT"
 sql = sql + ", [VMat]  BIT"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call PropTabA(QTb + ß, "FrozenColumns", 3, 2, ZielDB)
 .Fields![Medikament].AllowZeroLength = True
 Call CPropA(QTb + ß, "Medikament", "", 0, 109, 3030, 1, 0, ZielDB)
 .Fields![Langname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Langname", "Beispiel-Langname", 0, 109, 3645, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "Beispiel-PatID", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Anzahl", "Anzahl der Vorkommen", Null, 109, 810, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Glib", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Glib", "Glibenclamid", Null, 106, 540, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Metf", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Metf", "Metformin", Null, 106, 555, 0, 0, ZielDB)
 Call PropA(QTb + ß, "GlucI", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "GlucI", "Glucosidase-Inhibitoren", Null, 106, 600, 0, 0, ZielDB)
 Call PropA(QTb + ß, "SHGlin", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "SHGlin", "andere Sulfonylharnstoffe oder Glinide", Null, 106, 675, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Glit", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Glit", "Glitanzone", Null, 106, 480, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Ins", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Ins", "Insulin", Null, 106, 420, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Anal", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Anal", "Insulin-Analoga", Null, 106, 585, 0, 0, ZielDB)
 .Fields![InsArt].AllowZeroLength = True
 Call CPropA(QTb + ß, "InsArt", "1= schnell, 2 = langsam, 3 = Misch", 0, 109, 690, 0, 0, ZielDB)
 Call PropA(QTb + ß, "HMG", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "HMG", "HMG-CoA-Reduktase-Inhibitoren", Null, 106, 540, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Hypt", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Hypt", "Hypertonie-Mittel", Null, 106, 525, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Thro", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Thro", "Thrombozyten-Hemmer", Null, 106, 525, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Antib", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Antib", "Antibiotika", Null, 106, 615, 0, 0, ZielDB)
 Call PropA(QTb + ß, "and", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "and", "andere", Null, 106, 510, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "hinzugefügt", "", Null, Null, 1920, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Tstr", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Tstr", "Teststreifen", Null, 106, 495, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Puzu", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Puzu", "Pumpenzubehör", Null, 106, 630, 0, 0, ZielDB)
 Call PropA(QTb + ß, "VMat", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "VMat", "Verbandsmaterial", Null, 106, 630, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("PenN", 1, 1)
 Call PropA(QTb + ß, "PenN", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "PenN", "Pennadeln", Null, 106, 660, 0, 0, ZielDB)
 .Fields.Append .CreateField("Neurp", 1, 1)
 Call PropA(QTb + ß, "Neurp", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Neurp", "Neuropathie-Behandlungsmittel", Null, 106, 720, 0, 0, ZielDB)
 .Fields.Append .CreateField("AutNP", 1, 1)
 Call PropA(QTb + ß, "AutNP", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AutNP", "Autonome Neuropathie", Null, 106, 735, 0, 0, ZielDB)
 .Fields.Append .CreateField("Fetts", 1, 1)
 Call PropA(QTb + ß, "Fetts", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Fetts", "Fibrate, Ezetrol, Niaspan u.a.", Null, 106, 555, 0, 0, ZielDB)
 .Fields.Append .CreateField("Hsre", 1, 1)
 Call PropA(QTb + ß, "Hsre", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Hsre", "Hyperuriämie-Mittel", Null, 106, 495, 0, 0, ZielDB)
 .Fields.Append .CreateField("AntiMyk", 1, 1)
 Call PropA(QTb + ß, "AntiMyk", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AntiMyk", "Antimykotika", Null, 106, 810, 0, 0, ZielDB)
 .Fields.Append .CreateField("Glauk", 1, 1)
 Call PropA(QTb + ß, "Glauk", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Glauk", "Glaukom", Null, 106, 630, 0, 0, ZielDB)
 .Fields.Append .CreateField("COLD", 1, 1)
 Call PropA(QTb + ß, "COLD", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "COLD", "COLD und Asthma", Null, 106, 705, 0, 0, ZielDB)
 .Fields.Append .CreateField("Pros", 1, 1)
 Call PropA(QTb + ß, "Pros", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Pros", "Prostatahypertrophie", Null, 106, 645, 0, 0, ZielDB)
 .Fields.Append .CreateField("Urä", 1, 1)
 Call PropA(QTb + ß, "Urä", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Urä", "Urämie-spezifische", Null, 106, 480, 0, 0, ZielDB)
 .Fields.Append .CreateField("HyThy", 1, 1)
 Call PropA(QTb + ß, "HyThy", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "HyThy", "thyreostatische Mittel", Null, 106, 735, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ostp", 1, 1)
 Call PropA(QTb + ß, "Ostp", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Ostp", "Osteoporosemittel", Null, 106, 570, 0, 0, ZielDB)
 .Fields.Append .CreateField("KHK", 1, 1)
 Call PropA(QTb + ß, "KHK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "KHK", "KHK-spezifisch", Null, 106, 555, 0, 0, ZielDB)
 .Fields.Append .CreateField("HerzI", 1, 1)
 Call PropA(QTb + ß, "HerzI", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "HerzI", "Herzinsuffizienz-spezifische", Null, 106, 600, 0, 0, ZielDB)
 .Fields.Append .CreateField("Stru", 1, 1)
 Call PropA(QTb + ß, "Stru", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Stru", "Struma- und Hypothyreosemittel", Null, 106, 495, 0, 0, ZielDB)
 .Fields.Append .CreateField("AVK", 1, 1)
 Call PropA(QTb + ß, "AVK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AVK", "AVK-Mittel", Null, 106, 555, 0, 0, ZielDB)
 .Fields.Append .CreateField("PanI", 1, 1)
 Call PropA(QTb + ß, "PanI", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "PanI", "Pankreasinsuffizienz", Null, 106, 585, 0, 0, ZielDB)
 .Fields.Append .CreateField("Vari", 1, 1)
 Call PropA(QTb + ß, "Vari", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Vari", "Varikosemittel", Null, 106, 570, 0, 0, ZielDB)
 .Fields.Append .CreateField("Östr", 1, 1)
 Call PropA(QTb + ß, "Östr", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Östr", "Östrogene, Gestagene usw.", Null, 106, 525, 0, 0, ZielDB)
 .Fields.Append .CreateField("AntiDep", 1, 1)
 Call PropA(QTb + ß, "AntiDep", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AntiDep", "Antidepressiva", Null, 106, 855, 0, 0, ZielDB)
 .Fields.Append .CreateField("AntiDem", 1, 1)
 Call PropA(QTb + ß, "AntiDem", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AntiDem", "Antidementika", Null, 106, 960, 0, 0, ZielDB)
 .Fields.Append .CreateField("AntiEp", 1, 1)
 Call PropA(QTb + ß, "AntiEp", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AntiEp", "Antiepileptika", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Park", 1, 1)
 Call PropA(QTb + ß, "Park", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Park", "Parkinson-Medikament", Null, 106, 585, 0, 0, ZielDB)
 .Fields.Append .CreateField("AntiPern", 1, 1)
 Call PropA(QTb + ß, "AntiPern", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AntiPern", "Antiperniziosa", Null, 106, 555, 0, 0, ZielDB)
 .Fields.Append .CreateField("Appet", 1, 1)
 Call PropA(QTb + ß, "Appet", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Appet", "Appetitzügler", Null, 106, 690, 0, 0, ZielDB)
 .Fields.Append .CreateField("Anäm", 1, 1)
 Call PropA(QTb + ß, "Anäm", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Anäm", "Anämiebehandlungsmittel", Null, 106, 900, 0, 0, ZielDB)
 .Fields.Append .CreateField("Antiherp", 1, 1)
 Call PropA(QTb + ß, "Antiherp", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Antiherp", "Anti-Herpes-Mittel", Null, 106, 840, 0, 0, ZielDB)
 .Fields.Append .CreateField("NSAR", 1, 1)
 Call PropA(QTb + ß, "NSAR", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "NSAR", "NSAR", Null, 106, 690, 0, 0, ZielDB)
 .Fields.Append .CreateField("Antikoag", 1, 1)
 Call PropA(QTb + ß, "Antikoag", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Antikoag", "Antikoagulatien", Null, 106, 990, 0, 0, ZielDB)
 .Fields.Append .CreateField("Betabl", 1, 1)
 Call PropA(QTb + ß, "Betabl", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Betabl", "Betablocker", Null, 106, 765, 0, 0, ZielDB)
 .Fields.Append .CreateField("ACEH", 1, 1)
 Call PropA(QTb + ß, "ACEH", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "ACEH", "ACE-Hemmer", Null, 106, 675, 0, 0, ZielDB)
 .Fields.Append .CreateField("AT1", 1, 1)
 Call PropA(QTb + ß, "AT1", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AT1", "AT-1-Blocker", Null, 106, 510, 0, 0, ZielDB)
 .Fields.Append .CreateField("CalcA", 1, 1)
 Call PropA(QTb + ß, "CalcA", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "CalcA", "Calcium-Antagonist", Null, 106, 705, 0, 0, ZielDB)
 .Fields.Append .CreateField("Diur", 1, 1)
 Call PropA(QTb + ß, "Diur", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Diur", "Diuretikum", Null, 106, 540, 0, 0, ZielDB)
 .Fields.Append .CreateField("falsch", 1, 1)
 Call PropA(QTb + ß, "falsch", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "falsch", "falsch geschrieben oder nicht erkennbar", Null, 106, 675, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Medikament].Name) = "MEDIKAMENT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [Medikament] on [" + QTb + ß + "] ("
  sql = sql + "[Medikament]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [pat_id] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAMedArten/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAMedArten

Function MachALkat()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Lkat"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Kat]  TEXT(3)"
 sql = sql + ", [Ziffer]  TEXT(20)"
 sql = sql + ", [Text]  MEMO"
 sql = sql + ", [Euro]  MONEY"
 sql = sql + ", [Punkte]  INTEGER"
 sql = sql + ", [Reihe]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 .Fields![kat].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kat", "", 0, 109, -1, 1, 0, ZielDB)
 .Fields![Ziffer].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ziffer", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Text].AllowZeroLength = True
 Call CPropA(QTb + ß, "Text", "", 0, Null, 5790, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Euro", "", Null, 109, 795, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Punkte", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Reihe", "", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Kat]"
  sql = sql + ",[Ziffer]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALkat/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALkat

Function MachAlistenausgabeuew()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "listenausgabeuew"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[name]  TEXT(40)"
 sql = sql + ", [vorname]  TEXT(35)"
 sql = sql + ", [titel]  TEXT(40)"
 sql = sql + ", [fachgruppe]  TEXT(100)"
 sql = sql + ", [strasse]  TEXT(40)"
 sql = sql + ", [plz]  TEXT(5)"
 sql = sql + ", [ort]  TEXT(25)"
 sql = sql + ", [telefon]  TEXT(40)"
 sql = sql + ", [fax]  TEXT(40)"
 sql = sql + ", [kvnr]  TEXT(7)"
 sql = sql + ", [aktdat]  DATE"
 sql = sql + ", [id]  COUNTER CONSTRAINT [Index_37603BD8_E1B5_4FBF] PRIMARY KEY"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAlistenausgabeuew/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAlistenausgabeuew

Function MachALex()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Lex"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [SchGr]  TEXT(2)"
 sql = sql + ", [Leistung]  TEXT(10)"
 sql = sql + ", [übertragen]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Reihenfolge", Null, Null, -1, 0, -1, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Leistungsdatum", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!Pat_id", Null, 109, -1, 0, 0, ZielDB)
 .Fields![SchGr].AllowZeroLength = True
 Call CPropA(QTb + ß, "SchGr", "Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Leistung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leistung", "Ziffer aus EBM / GOÄ", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "übertragen", """"", übertragen", Null, Null, 2115, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![id].Name) = "ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALex/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALex

Function MachALeistungsexport__Kopie()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Leistungsexport, Kopie"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Prim]  COUNTER"
 sql = sql + ", [PatID]  INTEGER"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Leistung]  TEXT(50)"
 sql = sql + ", [SchGr]  TEXT(2)"
 sql = sql + ", [Status]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Prim", "Primärschlüssel wegen Reihenfolge", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "PatID", "-> Namen!Pat_id", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Leistungsdatum", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Leistung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leistung", "Ziffer aus EBM / GOÄ", 0, 109, -1, 0, 0, ZielDB)
 .Fields![SchGr].AllowZeroLength = True
 Call CPropA(QTb + ß, "SchGr", "Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Status].AllowZeroLength = True
 Call CPropA(QTb + ß, "Status", """"", übertragen", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![PatID].Name) = "PATID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [PatID] on [" + QTb + ß + "] ("
  sql = sql + "[PatID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALeistungsexport__Kopie/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALeistungsexport__Kopie

Function MachALeistungsexport()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Leistungsexport"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Prim]  COUNTER"
 sql = sql + ", [PatID]  INTEGER"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [UZeit]  DATE"
 sql = sql + ", [Leistung]  TEXT(50)"
 sql = sql + ", [SchGr]  TEXT(2)"
 sql = sql + ", [Status]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Prim", "Primärschlüssel wegen Reihenfolge", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "PatID", "-> Namen!Pat_id", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Leistungsdatum", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "UZeit", "Leistungszeit, falls gewünscht", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Leistung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leistung", "Ziffer aus EBM / GOÄ", 0, 109, -1, 0, 0, ZielDB)
 .Fields![SchGr].AllowZeroLength = True
 Call CPropA(QTb + ß, "SchGr", "Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Status].AllowZeroLength = True
 Call CPropA(QTb + ß, "Status", """"", übertragen", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![PatID].Name) = "PATID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [PatID] on [" + QTb + ß + "] ("
  sql = sql + "[PatID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALeistungsexport/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALeistungsexport

Function MachALeistungen_exportiert()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Leistungen exportiert"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [SchGr]  TEXT(2)"
 sql = sql + ", [Leistung]  TEXT(50)"
 sql = sql + ", [übertragen]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Reihenfolge", Null, Null, -1, 0, -1, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Leistungsdatum", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!Pat_id", Null, 109, -1, 0, 0, ZielDB)
 .Fields![SchGr].AllowZeroLength = True
 Call CPropA(QTb + ß, "SchGr", "Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Leistung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leistung", "Ziffer aus EBM / GOÄ", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "übertragen", """"", übertragen", Null, Null, 2115, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![id].Name) = "ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALeistungen_exportiert/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALeistungen_exportiert

Function MachALeistungen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Leistungen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [Leistung]  TEXT(10)"
 sql = sql + ", [f5002]  TEXT(32)"
 sql = sql + ", [f5005]  TEXT(2)"
 sql = sql + ", [f5006]  TEXT(20)"
 sql = sql + ", [f5009]  TEXT(80)"
 sql = sql + ", [Med]  TEXT(30)"
 sql = sql + ", [f5015]  TEXT(30)"
 sql = sql + ", [f5016]  TEXT(28)"
 sql = sql + ", [f5021]  TEXT(20)"
 sql = sql + ", [f5026]  TEXT(20)"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [QS]  TEXT(5)"
 sql = sql + ", [QT]  TEXT(5)"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 720, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "5000 + 6201", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Leistung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leistung", "5001", 0, 109, 960, 0, 0, ZielDB)
 .Fields![f5002].AllowZeroLength = True
 Call CPropA(QTb + ß, "f5002", "5002", 0, 109, 1860, 0, 0, ZielDB)
 .Fields![f5005].AllowZeroLength = True
 Call CPropA(QTb + ß, "f5005", "5005", 0, 109, 630, 0, 0, ZielDB)
 .Fields![f5006].AllowZeroLength = True
 Call CPropA(QTb + ß, "f5006", "5006", 0, 109, 630, 0, 0, ZielDB)
 .Fields![f5009].AllowZeroLength = True
 Call CPropA(QTb + ß, "f5009", "5009", 0, 109, 5595, 0, 0, ZielDB)
 .Fields![Med].AllowZeroLength = True
 Call CPropA(QTb + ß, "Med", "5010", 0, 109, -1, 0, 0, ZielDB)
 .Fields![f5015].AllowZeroLength = True
 Call CPropA(QTb + ß, "f5015", "5015", 0, 109, 2055, 0, 0, ZielDB)
 .Fields![f5016].AllowZeroLength = True
 Call CPropA(QTb + ß, "f5016", "5016", 0, 109, 2760, 0, 0, ZielDB)
 .Fields![f5021].AllowZeroLength = True
 Call CPropA(QTb + ß, "f5021", "5021", 0, 109, 630, 0, 0, ZielDB)
 .Fields![f5026].AllowZeroLength = True
 Call CPropA(QTb + ß, "f5026", "5026", 0, 109, 630, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, 825, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![QS].AllowZeroLength = True
 Call CPropA(QTb + ß, "QS", "Quartal des Behandlungsfallbeginns sortiert", 0, 109, -1, 0, 0, ZielDB)
 .Fields![QT].AllowZeroLength = True
 Call CPropA(QTb + ß, "QT", "Quartal des Behandlungsfallbeginns", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt]"
  sql = sql + ",[Leistung]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Leistung].Name) = "LEISTUNG" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Leistung] on [" + QTb + ß + "] ("
  sql = sql + "[Leistung]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALeistungen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALeistungen

Function MachALbAnforderungen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LbAnforderungen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [AnfText]  MEMO"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 465, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "6200 + 6201", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![AnfText].AllowZeroLength = True
 Call CPropA(QTb + ß, "AnfText", "6280", 0, Null, 12525, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, 825, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Statusbyte", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt]"
  sql = sql + ",[AnfText]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALbAnforderungen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALbAnforderungen

Function MachALauf()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Lauf"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Programm]  TEXT(30)"
 sql = sql + ", [abbruch]  BIT"
 sql = sql + ", [Computer]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALauf/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALauf

Function MachALaborXWert()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborXWert"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[RefNr]  INTEGER"
 sql = sql + ", [Abkü]  TEXT(16)"
 sql = sql + ", [Langname]  TEXT(50)"
 sql = sql + ", [Quelle]  TEXT(41)"
 sql = sql + ", [QSpez]  TEXT(20)"
 sql = sql + ", [AbnDat]  DATE"
 sql = sql + ", [Wert]  TEXT(10)"
 sql = sql + ", [Einheit]  TEXT(15)"
 sql = sql + ", [Grenzwerti]  TEXT(2)"
 sql = sql + ", [Kommentar]  MEMO"
 sql = sql + ", [Teststatus]  TEXT(5)"
 sql = sql + ", [Erklärung]  MEMO"
 sql = sql + ", [Normbereich]  TEXT(65)"
 sql = sql + ", [NormU]  TEXT(10)"
 sql = sql + ", [NormO]  TEXT(50)"
 sql = sql + ", [AuftrHinw]  TEXT(180)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "RefNr", "Bezug auf LaborUS", Null, 109, 675, 0, 0, ZielDB)
 .Fields![Abkü].AllowZeroLength = True
 Call CPropA(QTb + ß, "Abkü", "8410 Test-Ident  (Turbomed)", 0, 109, 1155, 0, 0, ZielDB)
 .Fields![Langname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Langname", "8411 Testbezeichnung (Turbomed)", 0, 109, 3525, 0, 0, ZielDB)
 .Fields![Quelle].AllowZeroLength = True
 Call CPropA(QTb + ß, "Quelle", "8430 Probenmaterial-Bezeichnung (Turbomed)", 0, 109, 780, 0, 0, ZielDB)
 .Fields![QSpez].AllowZeroLength = True
 Call CPropA(QTb + ß, "QSpez", "8431 Probenmaterial-Spezifikation (Turbomed)", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AbnDat", "8432 Abnahmedatum (Turbomed)", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Wert].AllowZeroLength = True
 Call CPropA(QTb + ß, "Wert", "8420 Ergebniswert (Turbomed)", 0, 109, 600, 0, 0, ZielDB)
 .Fields![Einheit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Einheit", "8421 Einheit (Turbomed)", 0, 109, 1110, 0, 0, ZielDB)
 .Fields![Grenzwerti].AllowZeroLength = True
 Call CPropA(QTb + ß, "Grenzwerti", "8422 Grenzwertindikator (Turbomed)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Kommentar].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kommentar", "8480 Ergebnistext (Turbomed)", 0, Null, 4860, 0, 0, ZielDB)
 .Fields![Teststatus].AllowZeroLength = True
 Call CPropA(QTb + ß, "Teststatus", "8418 Teststatus (Turbomed)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Erklärung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Erklärung", "8470 Testbezogene Hinweise (Turbomed)", 0, Null, 4500, 0, 0, ZielDB)
 .Fields![Normbereich].AllowZeroLength = True
 Call CPropA(QTb + ß, "Normbereich", "8460 Normalwert-Text (Turbomed)", 0, 109, 1380, 0, 0, ZielDB)
 .Fields![NormU].AllowZeroLength = True
 Call CPropA(QTb + ß, "NormU", "8461 Normuntergrenze", 0, 109, -1, 0, 0, ZielDB)
 .Fields![NormO].AllowZeroLength = True
 Call CPropA(QTb + ß, "NormO", "8461 Normobergrenze", 0, 109, -1, 0, 0, ZielDB)
 .Fields![AuftrHinw].AllowZeroLength = True
 Call CPropA(QTb + ß, "AuftrHinw", "8490 Auftragsbezogene Hinweise (Turbomed)", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![LaborXWertAbkü].Name) = "LABORXWERTABKÜ" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [LaborXWertAbkü] on [" + QTb + ß + "] ("
  sql = sql + "[Abkü]"
  sql = sql + ",[Einheit]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![RefNr].Name) = "REFNR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [RefNr] on [" + QTb + ß + "] ("
  sql = sql + "[RefNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborXWert/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborXWert

Function MachALaborXUS()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborXUS"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[RefNr]  COUNTER CONSTRAINT [RefNr] PRIMARY KEY"
 sql = sql + ", [DatID]  INTEGER"
 sql = sql + ", [SatzID]  INTEGER"
 sql = sql + ", [Satzart]  TEXT(10)"
 sql = sql + ", [Satzlänge]  TEXT(10)"
 sql = sql + ", [Auftragsnummer]  TEXT(11)"
 sql = sql + ", [Auftragsschlüssel]  TEXT(15)"
 sql = sql + ", [Eingang]  DATE"
 sql = sql + ", [Berichtsdatum]  TEXT(13)"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [Nachname]  TEXT(30)"
 sql = sql + ", [Vorname]  TEXT(20)"
 sql = sql + ", [GebDat]  TEXT(10)"
 sql = sql + ", [Titel]  TEXT(15)"
 sql = sql + ", [NVorsatz]  TEXT(15)"
 sql = sql + ", [BefArt]  TEXT(10)"
 sql = sql + ", [Abrechnungstyp]  TEXT(10)"
 sql = sql + ", [GebüOrd]  TEXT(10)"
 sql = sql + ", [Patienteninformation]  TEXT(10)"
 sql = sql + ", [Geschlecht]  TEXT(10)"
 sql = sql + ", [AuftrHinw]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "RefNr", "Bezug auf LaborWert", Null, Null, 555, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "DatID", "Bezug auf LaborEingelesen", Null, 109, 420, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "SatzID", "Bezug auf LaborXSätze", Null, 109, 540, 0, 0, ZielDB)
 .Fields![Satzart].AllowZeroLength = True
 Call CPropA(QTb + ß, "Satzart", "8000 Satzart (Turbomed)", 0, 109, 600, 0, 0, ZielDB)
 .Fields![Satzlänge].AllowZeroLength = True
 Call CPropA(QTb + ß, "Satzlänge", "8100 Satzlänge (Turbomed)", 0, 109, 750, 0, 0, ZielDB)
 .Fields![Auftragsnummer].AllowZeroLength = True
 Call CPropA(QTb + ß, "Auftragsnummer", "8310 Anforderungsident (Turbomed)", 0, 109, 735, 0, 0, ZielDB)
 .Fields![Auftragsschlüssel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Auftragsschlüssel", "8311 Anforderungsnr d Labors (Turbomed)", 0, 109, 1065, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Eingang", "8301 Eingangsdatum in Datumsform", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Berichtsdatum].AllowZeroLength = True
 Call CPropA(QTb + ß, "Berichtsdatum", "8302 Berichtsdatum", 0, 109, 990, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "", Null, 109, 555, 0, 0, ZielDB)
 .Fields![Nachname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Nachname", "3101", 0, 109, 1080, 0, 0, ZielDB)
 .Fields![Vorname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vorname", "3102", 0, 109, 1095, 0, 0, ZielDB)
 .Fields![gebdat].AllowZeroLength = True
 Call CPropA(QTb + ß, "GebDat", "3103", 0, 109, 990, 0, 0, ZielDB)
 .Fields![Titel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Titel", "3104", 0, 109, -1, 0, 0, ZielDB)
 .Fields![NVorsatz].AllowZeroLength = True
 Call CPropA(QTb + ß, "NVorsatz", "3100", 0, 109, -1, 0, 0, ZielDB)
 .Fields![BefArt].AllowZeroLength = True
 Call CPropA(QTb + ß, "BefArt", "8401 Befundart (Turbomed) / Fertigstellungsgrad (""E""=Endbefund, ""T"" = Teilbefund)", 0, 109, 420, 0, 0, ZielDB)
 .Fields![Abrechnungstyp].AllowZeroLength = True
 Call CPropA(QTb + ß, "Abrechnungstyp", "8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)", 0, 109, 435, 0, 0, ZielDB)
 .Fields![GebüOrd].AllowZeroLength = True
 Call CPropA(QTb + ß, "GebüOrd", "8403 Gebührenordnung (Turbomed)", 0, 109, 330, 0, 0, ZielDB)
 .Fields![Patienteninformation].AllowZeroLength = True
 Call CPropA(QTb + ß, "Patienteninformation", "8405 Patienteninformation (Turbomed)", 0, 109, 825, 0, 0, ZielDB)
 .Fields![Geschlecht].AllowZeroLength = True
 Call CPropA(QTb + ß, "Geschlecht", "8407 Geschlecht (Turbomed)", 0, 109, 390, 0, 0, ZielDB)
 .Fields![AuftrHinw].AllowZeroLength = True
 Call CPropA(QTb + ß, "AuftrHinw", "8490 Auftragsbezogene Hinweise (Turbomed)", 0, 109, 300, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("Pat_idUrsp", 10, 1)
 .Fields![Pat_idUrsp].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_idUrsp", "Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor", 0, 109, 330, 0, 0, ZielDB)
 .Fields.Append .CreateField("Pat_idErwVNG", 10, 20)
 .Fields![Pat_idErwVNG].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_idErwVNG", "erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag", 0, 109, 705, 0, 0, ZielDB)
 .Fields.Append .CreateField("Pat_idErwVN", 10, 20)
 .Fields![Pat_idErwVN].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_idErwVN", "erwogene Pat_id mit gleichem Vornamen und Nachnamen", 0, 109, 555, 0, 0, ZielDB)
 .Fields.Append .CreateField("Pat_idErwG", 10, 20)
 .Fields![Pat_idErwG].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_idErwG", "erwogene Pat_id mit gleichem Geburtstag", 0, 109, 570, 0, 0, ZielDB)
 .Fields.Append .CreateField("Pat_idErwGB", 10, 20)
 .Fields![Pat_idErwGB].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_idErwGB", "erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung", 0, 109, 540, 0, 0, ZielDB)
 .Fields.Append .CreateField("Pat_idErwGL", 10, 20)
 .Fields![Pat_idErwGL].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_idErwGL", "erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor", 0, 109, 570, 0, 0, ZielDB)
 .Fields.Append .CreateField("Pat_idLaborNeu", 10, 50)
 .Fields![Pat_idLaborNeu].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_idLaborNeu", "Pat_ids von in Laborneu zuordnbaren Patienten", 0, 109, 1500, 0, 0, ZielDB)
 .Fields.Append .CreateField("ZeitpunktLaborneu", 8, 8)
 Call CPropA(QTb + ß, "ZeitpunktLaborneu", "Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde", Null, Null, 1935, 0, 0, ZielDB)
 .Fields.Append .CreateField("ZdüP", 3, 2)
 Call CPropA(QTb + ß, "ZdüP", "Zahl der verglichenen Parameter", Null, 109, 630, 0, 0, ZielDB)
 .Fields.Append .CreateField("ZdiP", 4, 4)
 Call CPropA(QTb + ß, "ZdiP", "Zahl der infragekommenden Patienten", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("LWerte", 12, 0)
 .Fields![LWerte].AllowZeroLength = True
 Call CPropA(QTb + ß, "LWerte", "Laborwerte, die zur Zuordnung geführt haben", 0, Null, 18375, 0, 0, ZielDB)
 .Fields.Append .CreateField("verglichen", 8, 8)
 Call CPropA(QTb + ß, "verglichen", "Datum, zu dem Datensatz zuletzt verglichen wurde", Null, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("AfN", 3, 2)
 Call CPropA(QTb + ß, "AfN", "Affected Number: Zahl der zugehörigen Datensätze in Laborneu", Null, 109, -1, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![DatID].Name) = "DATID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [DatID] on [" + QTb + ß + "] ("
  sql = sql + "[DatID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Name].Name) = "NAME" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Name] on [" + QTb + ß + "] ("
  sql = sql + "[Nachname]"
  sql = sql + ",[Vorname]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Pat_id] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![SatzID].Name) = "SATZID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [SatzID] on [" + QTb + ß + "] ("
  sql = sql + "[SatzID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborXUS/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborXUS

Function MachALaborXSaetze()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborXSaetze"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[SatzID]  COUNTER CONSTRAINT [SatzID] UNIQUE"
 sql = sql + ", [DatID]  INTEGER"
 sql = sql + ", [Satzart]  TEXT(20)"
 sql = sql + ", [Satzlänge]  TEXT(10)"
 sql = sql + ", [SatzlängeSchluss]  TEXT(10)"
 sql = sql + ", [VersionSatzb]  TEXT(50)"
 sql = sql + ", [Arztnr]  TEXT(10)"
 sql = sql + ", [Arztname]  TEXT(30)"
 sql = sql + ", [StraßePraxis]  TEXT(30)"
 sql = sql + ", [PLZPraxis]  TEXT(10)"
 sql = sql + ", [OrtPraxis]  TEXT(20)"
 sql = sql + ", [Labor]  TEXT(30)"
 sql = sql + ", [StraßeLabor]  TEXT(50)"
 sql = sql + ", [PLZLabor]  TEXT(30)"
 sql = sql + ", [OrtLabor]  TEXT(20)"
 sql = sql + ", [KBVPrüfnr]  TEXT(10)"
 sql = sql + ", [Zeichensatz]  TEXT(10)"
 sql = sql + ", [Kundenarztnr]  TEXT(10)"
 sql = sql + ", [Erstellungsdatum]  TEXT(25)"
 sql = sql + ", [Gesamtlänge]  TEXT(10)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "SatzID", "zum Bezug für LaborUS", Null, Null, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "DatID", "Bezug zu LaborEingelesen", Null, 109, -1, 0, 0, ZielDB)
 .Fields![Satzart].AllowZeroLength = True
 Call CPropA(QTb + ß, "Satzart", "8000 Satzart (Turbomed)", 0, 109, 825, 0, 0, ZielDB)
 .Fields![Satzlänge].AllowZeroLength = True
 Call CPropA(QTb + ß, "Satzlänge", "8100 Satzlänge (Turbomed)", 0, 109, 1110, 0, 0, ZielDB)
 .Fields![SatzlängeSchluss].AllowZeroLength = True
 Call CPropA(QTb + ß, "SatzlängeSchluss", "8100 Satzlänge (Turbomed), nach 8221 in Feld 8000", 0, 109, -1, 0, 0, ZielDB)
 .Fields![VersionSatzb].AllowZeroLength = True
 Call CPropA(QTb + ß, "VersionSatzb", "9212 Version der Satzbeschreibung (Turbomed)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Arztnr].AllowZeroLength = True
 Call CPropA(QTb + ß, "Arztnr", "201 Arztnummer (Turbomed)", 0, 109, 885, 0, 0, ZielDB)
 .Fields![Arztname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Arztname", "203 Arztname (Turbomed)", 0, 109, 2760, 0, 0, ZielDB)
 .Fields![StraßePraxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "StraßePraxis", "205 Straße der Praxis (Turbomed)", 0, 109, 1575, 0, 0, ZielDB)
 .Fields![PLZPraxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "PLZPraxis", "215 PLZ der Praxis (Turbomed)", 0, 109, 1110, 0, 0, ZielDB)
 .Fields![OrtPraxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "OrtPraxis", "216 Ort der Praxis (Turbomed)", 0, 109, 1035, 0, 0, ZielDB)
 .Fields![Labor].AllowZeroLength = True
 Call CPropA(QTb + ß, "Labor", "8320 Labor", 0, 109, 2595, 0, 0, ZielDB)
 .Fields![StraßeLabor].AllowZeroLength = True
 Call CPropA(QTb + ß, "StraßeLabor", "8321 Straße der Laboradresse (Turbomed)", 0, 109, 1575, 0, 0, ZielDB)
 .Fields![PLZLabor].AllowZeroLength = True
 Call CPropA(QTb + ß, "PLZLabor", "8322 PLZ der Laboradresse (Turbomed)", 0, 109, 1065, 0, 0, ZielDB)
 .Fields![OrtLabor].AllowZeroLength = True
 Call CPropA(QTb + ß, "OrtLabor", "8323 Ort der Laboradresse (Turbomed)", 0, 109, 1620, 0, 0, ZielDB)
 .Fields![KBVPrüfnr].AllowZeroLength = True
 Call CPropA(QTb + ß, "KBVPrüfnr", "101 KBV-Prüfnummer (Turbomed)", 0, 109, 1140, 0, 0, ZielDB)
 .Fields![Zeichensatz].AllowZeroLength = True
 Call CPropA(QTb + ß, "Zeichensatz", "9106 verwendeter Zeichensatz (Turbomed)", 0, 109, 1275, 0, 0, ZielDB)
 .Fields![Kundenarztnr].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kundenarztnr", "8312 Kundenarztnummer (Turbomed)", 0, 109, 1440, 0, 0, ZielDB)
 .Fields![Erstellungsdatum].AllowZeroLength = True
 Call CPropA(QTb + ß, "Erstellungsdatum", "9103 Erstellungsdatum (Turbomed)", 0, 109, 1785, 0, 0, ZielDB)
 .Fields![Gesamtlänge].AllowZeroLength = True
 Call CPropA(QTb + ß, "Gesamtlänge", "9202 Gesamtlänge des Datenpaketes (Turbomed)", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![DatID].Name) = "DATID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [DatID] on [" + QTb + ß + "] ("
  sql = sql + "[DatID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Name].Name) = "NAME" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Name] on [" + QTb + ß + "] ("
  sql = sql + "[PLZLabor]"
  sql = sql + ",[OrtLabor]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![SatzID].Name) = "SATZID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [SatzID] on [" + QTb + ß + "] ("
  sql = sql + "[SatzID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborXSaetze/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborXSaetze

Function MachALaborXLeist()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborXLeist"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[RefNr]  INTEGER"
 sql = sql + ", [Abkü]  TEXT(15)"
 sql = sql + ", [Verf]  TEXT(42)"
 sql = sql + ", [EBM]  TEXT(8)"
 sql = sql + ", [goä]  TEXT(8)"
 sql = sql + ", [Anzahl]  TEXT(2)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "RefNr", "Bezug auf LaborUS", Null, 109, 675, 0, 0, ZielDB)
 .Fields![Abkü].AllowZeroLength = True
 Call CPropA(QTb + ß, "Abkü", "8410 Test-Ident (Turbomed)", 0, 109, 960, 0, 0, ZielDB)
 .Fields![Verf].AllowZeroLength = True
 Call CPropA(QTb + ß, "Verf", "8434", 0, 109, 540, 0, 0, ZielDB)
 .Fields![EBM].AllowZeroLength = True
 Call CPropA(QTb + ß, "EBM", "5001 GNR (Turbomed)", 0, 109, 675, 0, 0, ZielDB)
 .Fields![goä].AllowZeroLength = True
 Call CPropA(QTb + ß, "goä", "8406", 0, 109, 570, 0, 0, ZielDB)
 .Fields![Anzahl].AllowZeroLength = True
 Call CPropA(QTb + ß, "Anzahl", "5005", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![RefNr].Name) = "REFNR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [RefNr] on [" + QTb + ß + "] ("
  sql = sql + "[RefNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborXLeist/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborXLeist

Function MachALaborXEingel()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborXEingel"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[DatID]  COUNTER CONSTRAINT [DatID] PRIMARY KEY"
 sql = sql + ", [Pfad]  TEXT(100)"
 sql = sql + ", [Name]  TEXT(50)"
 sql = sql + ", [Zp]  DATE"
 sql = sql + ", [fertig]  BIT"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "DatID", "Bezug auf LaborEingelesen", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Pfad].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pfad", "Pfadname", 0, 109, 4470, 0, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Name der eingelesenen Labordatei ohne Endung", 0, 109, 2550, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp", "Einlesezeitpunkt", Null, Null, 1920, 0, 0, ZielDB)
 Call PropA(QTb + ß, "fertig", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "fertig", "ob Einlesen fertig", Null, 106, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![NamePfad].Name) = "NAMEPFAD" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [NamePfad] on [" + QTb + ß + "] ("
  sql = sql + "[Name]"
  sql = sql + ",[Pfad]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborXEingel/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborXEingel

Function MachALaborXBakt()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborXBakt"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[RefNr]  INTEGER"
 sql = sql + ", [Verf]  TEXT(42)"
 sql = sql + ", [KuQu]  TEXT(50)"
 sql = sql + ", [Quelle]  TEXT(50)"
 sql = sql + ", [QSpez]  TEXT(20)"
 sql = sql + ", [AbnDat]  DATE"
 sql = sql + ", [Kommentar]  MEMO"
 sql = sql + ", [Erklärung]  MEMO"
 sql = sql + ", [Keimzahl]  TEXT(30)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "RefNr", "", Null, 109, 675, 0, 0, ZielDB)
 .Fields![Verf].AllowZeroLength = True
 Call CPropA(QTb + ß, "Verf", "", 0, 109, 2085, 0, 0, ZielDB)
 .Fields![KuQu].AllowZeroLength = True
 Call CPropA(QTb + ß, "KuQu", "8428 Probenmaterial-Ident (Turbomed)", 0, 109, 675, 0, 0, ZielDB)
 .Fields![Quelle].AllowZeroLength = True
 Call CPropA(QTb + ß, "Quelle", "8430 Probenmaterial-Bezeichnung (Turbomed)", 0, 109, 1095, 0, 0, ZielDB)
 .Fields![QSpez].AllowZeroLength = True
 Call CPropA(QTb + ß, "QSpez", "8431 Probenmaterial-Spezifikation (Turbomed)", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AbnDat", "8432 Abnahmedatum (Turbomed)", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Kommentar].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kommentar", "8480 Ergebnistest (Turbomed)", 0, Null, 4620, 0, 0, ZielDB)
 .Fields![Erklärung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Erklärung", "", 0, Null, 1065, 0, 0, ZielDB)
 .Fields![Keimzahl].AllowZeroLength = True
 Call CPropA(QTb + ß, "Keimzahl", "", 0, 109, 1590, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![RefNr].Name) = "REFNR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [RefNr] on [" + QTb + ß + "] ("
  sql = sql + "[RefNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborXBakt/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborXBakt

Function MachALaborParameter()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborParameter"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Abkü]  TEXT(46)"
 sql = sql + ", [Labor]  TEXT(50)"
 sql = sql + ", [Langtext]  TEXT(60)"
 sql = sql + ", [Einheit]  TEXT(20)"
 sql = sql + ", [Gruppe]  INTEGER"
 sql = sql + ", [Reihe]  INTEGER"
 sql = sql + ", [uNm]  TEXT(7)"
 sql = sql + ", [oNm]  TEXT(7)"
 sql = sql + ", [uNw]  TEXT(7)"
 sql = sql + ", [oNw]  TEXT(7)"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ", CONSTRAINT [Abkü] UNIQUE ( Abkü, Einheit)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 .Fields![Abkü].AllowZeroLength = True
 Call CPropA(QTb + ß, "Abkü", "8410(1)", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Labor", "8410(2)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Langtext].AllowZeroLength = True
 Call CPropA(QTb + ß, "Langtext", "8411", 0, 109, 3225, 0, 0, ZielDB)
 .Fields![Einheit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Einheit", "8421", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Gruppe", "Ordnungsgruppe", Null, 109, 1620, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Reihe", "Reihenfolge innerhalb der Gruppe", Null, 109, -1, 0, 0, ZielDB)
 .Fields![unm].AllowZeroLength = True
 Call CPropA(QTb + ß, "uNm", "unterer Normwert männlich", 0, 109, -1, 0, 0, ZielDB)
 .Fields![onm].AllowZeroLength = True
 Call CPropA(QTb + ß, "oNm", "oberer Normwert männlich", 0, 109, -1, 0, 0, ZielDB)
 .Fields![unw].AllowZeroLength = True
 Call CPropA(QTb + ß, "uNw", "unterer Normwert weiblich", 0, 109, -1, 0, 0, ZielDB)
 .Fields![onw].AllowZeroLength = True
 Call CPropA(QTb + ß, "oNw", "oberer Normwert weiblich", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Abkü].Name) = "ABKÜ" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [Abkü] on [" + QTb + ß + "] ("
  sql = sql + "[Abkü]"
  sql = sql + ",[Einheit]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Fehlende].Name) = "FEHLENDE" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Fehlende] on [" + QTb + ß + "] ("
  sql = sql + "[Gruppe] DESC"
  sql = sql + ",[AktZeit]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![LaborParameterAbkü].Name) = "LABORPARAMETERABKÜ" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [LaborParameterAbkü] on [" + QTb + ß + "] ("
  sql = sql + "[Abkü]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![reihe].Name) = "REIHE" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Reihe] on [" + QTb + ß + "] ("
  sql = sql + "[Gruppe]"
  sql = sql + ",[Abkü]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborParameter/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborParameter

Function MachALaborNeu()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborNeu"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [FertigStGrad]  TEXT(1)"
 sql = sql + ", [Abkü]  TEXT(46)"
 sql = sql + ", [LangtextVW]  INTEGER"
 sql = sql + ", [Wert]  TEXT(39)"
 sql = sql + ", [Einheit]  TEXT(20)"
 sql = sql + ", [KommentarVW]  INTEGER"
 sql = sql + ", [AbsPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [Refnr]  INTEGER"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 720, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![FertigStGrad].AllowZeroLength = True
 Call CPropA(QTb + ß, "FertigStGrad", "8401", 0, 109, 1365, 0, 0, ZielDB)
 .Fields![Abkü].AllowZeroLength = True
 Call CPropA(QTb + ß, "Abkü", "8410", 0, 109, 930, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "LangtextVW", "8411", Null, 109, 1320, 0, 0, ZielDB)
 .Fields![Wert].AllowZeroLength = True
 Call CPropA(QTb + ß, "Wert", "8420", 0, 109, 600, 0, 0, ZielDB)
 .Fields![Einheit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Einheit", "8421", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "KommentarVW", "8480", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AbsPos", "Zeile in der BDT-Datei", Null, 109, 840, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Refnr", "Bezug auf LaborXUS", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![AbküWert].Name) = "ABKÜWERT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [AbküWert] on [" + QTb + ß + "] ("
  sql = sql + "[Abkü]"
  sql = sql + ",[Wert]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt]"
  sql = sql + ",[FertigStGrad]"
  sql = sql + ",[Abkü]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Prüf].Name) = "PRÜF" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Prüf] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[Abkü]"
  sql = sql + ",[Wert]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborNeu/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborNeu

Function MachALaborLangtext()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborLangtext"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[LangtextVW]  COUNTER CONSTRAINT [LangtextVW] PRIMARY KEY"
 sql = sql + ", [Langtext]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "LangtextVW", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Langtext].AllowZeroLength = True
 Call CPropA(QTb + ß, "Langtext", "", 0, Null, 4575, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Langtext].Name) = "LANGTEXT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Langtext] on [" + QTb + ß + "] ("
  sql = sql + "[Langtext]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborLangtext/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborLangtext

Function MachALaborKommentar()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborKommentar"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[KommentarVW]  COUNTER CONSTRAINT [KommentarVW] PRIMARY KEY"
 sql = sql + ", [Kommentar]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "KommentarVW", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Kommentar].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kommentar", "", 0, Null, 17325, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Kommentar].Name) = "KOMMENTAR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Kommentar] on [" + QTb + ß + "] ("
  sql = sql + "[Kommentar]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborKommentar/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborKommentar

Function MachALaborgruppen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Laborgruppen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Laborgruppe]  INTEGER CONSTRAINT [Gruppe] PRIMARY KEY"
 sql = sql + ", [Erklärung]  TEXT(100)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Laborgruppe", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields![Erklärung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Erklärung", "", 0, 109, 9795, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborgruppen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborgruppen

Function MachALaborEintraege()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "LaborEintraege"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER NOT NULL CONSTRAINT [id] PRIMARY KEY"
 sql = sql + ", [Länge]  TEXT(4)"
 sql = sql + ", [Kennung]  TEXT(4)"
 sql = sql + ", [Inhalt]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, 720, 1, 0, ZielDB)
 .Fields![Länge].AllowZeroLength = True
 Call CPropA(QTb + ß, "Länge", "", 0, 109, 750, 0, 0, ZielDB)
 .Fields![Kennung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kennung", "", 0, 109, 1005, 0, 0, ZielDB)
 .Fields![Inhalt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Inhalt", "", 0, Null, 5145, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Index_3E84F79A_81DC_4D48].Name) = "INDEX_3E84F79A_81DC_4D48" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [Index_3E84F79A_81DC_4D48] on [" + QTb + ß + "] ("
  sql = sql + "[ID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachALaborEintraege/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachALaborEintraege

Function MachAKVNrUe()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "KVNrUe"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[lfdnr]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [KVNr]  TEXT(7)"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ", CONSTRAINT [zuord] UNIQUE ( Pat_ID, KVNr)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "lfdnr", "", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "KVNr", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in BDT-Datei", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Zeit der Aktualisuerung aus der BDT-Datei", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Zuord].Name) = "ZUORD" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [zuord] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[KVNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAKVNrUe/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAKVNrUe

Function MachAKopie_von_Hausaerzte()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Kopie von Hausaerzte"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(100)"
 sql = sql + ", [Anschrift]  TEXT(200)"
 sql = sql + ", [KVNr]  TEXT(20)"
 sql = sql + ", [Telefon]  TEXT(50)"
 sql = sql + ", [Telefax]  TEXT(50)"
 sql = sql + ", [E_Mail]  TEXT(70)"
 sql = sql + ", [Zulassungsgebiet]  TEXT(50)"
 sql = sql + ", [Arzttyp]  TEXT(50)"
 sql = sql + ", [Gemeinschaftspraxis mit]  MEMO"
 sql = sql + ", [Schwerpunkt]  TEXT(50)"
 sql = sql + ", [Zusatzbezeichnung]  TEXT(70)"
 sql = sql + ", [Bemerkung]  MEMO"
 sql = sql + ", [Sprechstunden]  TEXT(255)"
 sql = sql + ", [von _ bis]  TEXT(150)"
 sql = sql + ", [Internetadressen]  MEMO"
 sql = sql + ", [Behandlung in Fremdsprachen]  TEXT(50)"
 sql = sql + ", [Rollstuhlgerechte Praxis]  TEXT(20)"
 sql = sql + ", [Verkehrsmittel]  TEXT(50)"
 sql = sql + ", [Linie]  TEXT(20)"
 sql = sql + ", [Haltestelle Parkplätze]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "", 0, 109, 5310, 0, 0, ZielDB)
 .Fields![Anschrift].AllowZeroLength = True
 Call CPropA(QTb + ß, "Anschrift", "", 0, 109, 4365, 0, 0, ZielDB)
 .Fields![KVNr].AllowZeroLength = True
 Call CPropA(QTb + ß, "KVNr", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Telefon].AllowZeroLength = True
 Call CPropA(QTb + ß, "Telefon", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Telefax].AllowZeroLength = True
 Call CPropA(QTb + ß, "Telefax", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![E_Mail].AllowZeroLength = True
 Call CPropA(QTb + ß, "E_Mail", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Zulassungsgebiet].AllowZeroLength = True
 Call CPropA(QTb + ß, "Zulassungsgebiet", "", 0, 109, 2805, 0, 0, ZielDB)
 .Fields![Arzttyp].AllowZeroLength = True
 Call CPropA(QTb + ß, "Arzttyp", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Gemeinschaftspraxis mit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Gemeinschaftspraxis mit", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![Schwerpunkt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schwerpunkt", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Zusatzbezeichnung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Zusatzbezeichnung", "", 0, 109, 4080, 0, 0, ZielDB)
 .Fields![Bemerkung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bemerkung", "", 0, Null, 17550, 0, 0, ZielDB)
 .Fields![Sprechstunden].AllowZeroLength = True
 Call CPropA(QTb + ß, "Sprechstunden", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![von _ bis].AllowZeroLength = True
 Call CPropA(QTb + ß, "von _ bis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Internetadressen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Internetadressen", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![Behandlung in Fremdsprachen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Behandlung in Fremdsprachen", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Rollstuhlgerechte Praxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "Rollstuhlgerechte Praxis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Verkehrsmittel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Verkehrsmittel", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Linie].AllowZeroLength = True
 Call CPropA(QTb + ß, "Linie", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Haltestelle Parkplätze].AllowZeroLength = True
 Call CPropA(QTb + ß, "Haltestelle Parkplätze", "", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("Wegbeschreibung", 12, 0)
 .Fields![Wegbeschreibung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Wegbeschreibung", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Entfernung zur Praxis", 10, 50)
 .Fields![Entfernung zur Praxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "Entfernung zur Praxis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Zahl", 4, 4)
 Call CPropA(QTb + ß, "Zahl", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("nichtmehr", 1, 1)
 Call PropA(QTb + ß, "nichtmehr", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "nichtmehr", "Arzt nicht mehr im Verzeichnis", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Vorname", 10, 50)
 .Fields![Vorname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vorname", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Nachname", 10, 50)
 .Fields![Nachname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Nachname", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Titel", 10, 50)
 .Fields![Titel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Titel", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Geschlecht", 10, 50)
 .Fields![Geschlecht].AllowZeroLength = True
 Call CPropA(QTb + ß, "Geschlecht", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Straße", 10, 50)
 .Fields![Straße].AllowZeroLength = True
 Call CPropA(QTb + ß, "Straße", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("PLZ", 10, 10)
 .Fields![Plz].AllowZeroLength = True
 Call CPropA(QTb + ß, "PLZ", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ort", 10, 50)
 .Fields![Ort].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ort", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DMPT2", 3, 2)
 Call CPropA(QTb + ß, "DMPT2", "", Null, 109, -1, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[KVNr]"
  sql = sql + ",[Name]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![KVNr].Name) = "KVNR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [KVNr] on [" + QTb + ß + "] ("
  sql = sql + "[KVNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAKopie_von_Hausaerzte/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAKopie_von_Hausaerzte

Function MachAKHEinweis()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "KHEinweis"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [Ziel]  TEXT(100)"
 sql = sql + ", [Diagnose]  MEMO"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 465, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Ziel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ziel", "6291", 0, 109, 3555, 0, 0, ZielDB)
 .Fields![Diagnose].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diagnose", "6230", 0, Null, 18600, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, 885, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt] DESC"
  sql = sql + ",[Ziel] DESC"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAKHEinweis/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAKHEinweis

Function MachAKassenliste()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Kassenliste"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [VK]  TEXT(6)"
 sql = sql + ", [IK]  TEXT(8)"
 sql = sql + ", [Name]  TEXT(100)"
 sql = sql + ", [Kateg]  TEXT(3)"
 sql = sql + ", [AnzahlIK]  INTEGER"
 sql = sql + ", [AnzahlKTUG]  INTEGER"
 sql = sql + ", [GültigVon]  TEXT(10) NOT NULL"
 sql = sql + ", [GültigBis]  TEXT(10) NOT NULL"
 sql = sql + ", [GO]  TEXT(15) NOT NULL"
 sql = sql + ", [Kurzname]  TEXT(60)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![VK].AllowZeroLength = True
 Call CPropA(QTb + ß, "VK", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![IK].AllowZeroLength = True
 Call CPropA(QTb + ß, "IK", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "", 0, 109, 5565, 0, 0, ZielDB)
 .Fields![Kateg].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kateg", "Kategorie", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AnzahlIK", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AnzahlKTUG", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "GültigVon", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "GültigBis", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "GO", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Kurzname", "", 0, 109, 3180, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![VK].Name) = "VK" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [VK] on [" + QTb + ß + "] ("
  sql = sql + "[VK]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAKassenliste/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAKassenliste

Function MachAICD2006()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "ICD2006"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[klassebene]  TEXT(1)"
 sql = sql + ", [OrtdSchlüNr]  TEXT(1)"
 sql = sql + ", [ArtdVierst]  TEXT(1)"
 sql = sql + ", [KapitelNr]  TEXT(2)"
 sql = sql + ", [Dreist]  TEXT(3)"
 sql = sql + ", [SchlüssNrRoh]  TEXT(6)"
 sql = sql + ", [klassTit]  TEXT(255)"
 sql = sql + ", [VerwdP295]  TEXT(1)"
 sql = sql + ", [VerwdP311]  TEXT(1)"
 sql = sql + ", [ME1]  TEXT(5)"
 sql = sql + ", [ME2]  TEXT(5)"
 sql = sql + ", [ME3]  TEXT(5)"
 sql = sql + ", [ME4]  TEXT(5)"
 sql = sql + ", [MbListe]  TEXT(5)"
 sql = sql + ", [Geschl]  TEXT(1)"
 sql = sql + ", [FlrBGschl]  TEXT(1)"
 sql = sql + ", [untAltG]  TEXT(4)"
 sql = sql + ", [untAltGalt]  TEXT(4)"
 sql = sql + ", [obAltG]  TEXT(4)"
 sql = sql + ", [obAltGalt]  TEXT(4)"
 sql = sql + ", [ArtdFbAlt]  TEXT(1)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("KiMs", 10, 4)
 .Fields.Append .CreateField("SNrmInh", 10, 1)
 .Fields.Append .CreateField("IfsGMeld", 10, 1)
 .Fields.Append .CreateField("IfsGLab", 10, 1)
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAICD2006/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAICD2006

Function MachAHAz()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "HAz"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(100)"
 sql = sql + ", [Anschrift]  TEXT(200)"
 sql = sql + ", [KV_Nr]  TEXT(20)"
 sql = sql + ", [Telefon]  TEXT(50)"
 sql = sql + ", [Telefax]  TEXT(50)"
 sql = sql + ", [E_Mail]  TEXT(70)"
 sql = sql + ", [Zulassungsgebiet]  TEXT(50)"
 sql = sql + ", [Arzttyp]  TEXT(50)"
 sql = sql + ", [Gemeinschaftspraxis mit]  MEMO"
 sql = sql + ", [Schwerpunkt]  TEXT(50)"
 sql = sql + ", [Zusatzbezeichnung]  TEXT(70)"
 sql = sql + ", [Bemerkung]  MEMO"
 sql = sql + ", [Sprechstunden]  TEXT(255)"
 sql = sql + ", [von _ bis]  TEXT(150)"
 sql = sql + ", [Internetadressen]  MEMO"
 sql = sql + ", [Behandlung in Fremdsprachen]  TEXT(50)"
 sql = sql + ", [Rollstuhlgerechte Praxis]  TEXT(20)"
 sql = sql + ", [Verkehrsmittel]  TEXT(50)"
 sql = sql + ", [Linie]  TEXT(20)"
 sql = sql + ", [Haltestelle Parkplätze]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "", 0, 109, 5310, 0, 0, ZielDB)
 .Fields![Anschrift].AllowZeroLength = True
 Call CPropA(QTb + ß, "Anschrift", "", 0, 109, 4365, 0, 0, ZielDB)
 .Fields![KV_Nr].AllowZeroLength = True
 Call CPropA(QTb + ß, "KV_Nr", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Telefon].AllowZeroLength = True
 Call CPropA(QTb + ß, "Telefon", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Telefax].AllowZeroLength = True
 Call CPropA(QTb + ß, "Telefax", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![E_Mail].AllowZeroLength = True
 Call CPropA(QTb + ß, "E_Mail", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Zulassungsgebiet].AllowZeroLength = True
 Call CPropA(QTb + ß, "Zulassungsgebiet", "", 0, 109, 2805, 0, 0, ZielDB)
 .Fields![Arzttyp].AllowZeroLength = True
 Call CPropA(QTb + ß, "Arzttyp", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Gemeinschaftspraxis mit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Gemeinschaftspraxis mit", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![Schwerpunkt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schwerpunkt", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Zusatzbezeichnung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Zusatzbezeichnung", "", 0, 109, 4080, 0, 0, ZielDB)
 .Fields![Bemerkung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bemerkung", "", 0, Null, 17550, 0, 0, ZielDB)
 .Fields![Sprechstunden].AllowZeroLength = True
 Call CPropA(QTb + ß, "Sprechstunden", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![von _ bis].AllowZeroLength = True
 Call CPropA(QTb + ß, "von _ bis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Internetadressen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Internetadressen", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![Behandlung in Fremdsprachen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Behandlung in Fremdsprachen", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Rollstuhlgerechte Praxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "Rollstuhlgerechte Praxis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Verkehrsmittel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Verkehrsmittel", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Linie].AllowZeroLength = True
 Call CPropA(QTb + ß, "Linie", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Haltestelle Parkplätze].AllowZeroLength = True
 Call CPropA(QTb + ß, "Haltestelle Parkplätze", "", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("Wegbeschreibung", 12, 0)
 .Fields![Wegbeschreibung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Wegbeschreibung", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Entfernung zur Praxis", 10, 50)
 .Fields![Entfernung zur Praxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "Entfernung zur Praxis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Zahl", 4, 4)
 Call CPropA(QTb + ß, "Zahl", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("nichtmehr", 1, 1)
 Call PropA(QTb + ß, "nichtmehr", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "nichtmehr", "Arzt nicht mehr im Verzeichnis", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Vorname", 10, 50)
 .Fields![Vorname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vorname", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Nachname", 10, 50)
 .Fields![Nachname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Nachname", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Titel", 10, 50)
 .Fields![Titel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Titel", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Geschlecht", 10, 50)
 .Fields![Geschlecht].AllowZeroLength = True
 Call CPropA(QTb + ß, "Geschlecht", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Straße", 10, 50)
 .Fields![Straße].AllowZeroLength = True
 Call CPropA(QTb + ß, "Straße", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("PLZ", 10, 10)
 .Fields![Plz].AllowZeroLength = True
 Call CPropA(QTb + ß, "PLZ", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ort", 10, 50)
 .Fields![Ort].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ort", "", 0, 109, -1, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[KV_Nr]"
  sql = sql + ",[Name]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![KVNr].Name) = "KVNR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [KVNr] on [" + QTb + ß + "] ("
  sql = sql + "[KV_Nr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAHAz/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAHAz

Function MachAHausaerzte()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Hausaerzte"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Überschrift]  TEXT(1)"
 sql = sql + ", [Name]  TEXT(100)"
 sql = sql + ", [Vorname]  TEXT(50)"
 sql = sql + ", [Nachname]  TEXT(50)"
 sql = sql + ", [Anschrift]  TEXT(200)"
 sql = sql + ", [KVNr]  TEXT(20)"
 sql = sql + ", [Telefon]  TEXT(50)"
 sql = sql + ", [Telefax]  TEXT(50)"
 sql = sql + ", [E_Mail]  TEXT(70)"
 sql = sql + ", [Zulassungsgebiet]  TEXT(200)"
 sql = sql + ", [Arzttyp]  TEXT(50)"
 sql = sql + ", [Gemeinschaftspraxis mit]  MEMO"
 sql = sql + ", [Schwerpunkt]  TEXT(50)"
 sql = sql + ", [Zusatzbezeichnung]  TEXT(70)"
 sql = sql + ", [Bemerkung]  MEMO"
 sql = sql + ", [Beme]  MEMO"
 sql = sql + ", [Sprechstunden]  TEXT(255)"
 sql = sql + ", [von _ bis]  TEXT(150)"
 sql = sql + ", [Internetadressen]  MEMO"
 sql = sql + ", [Behandlung in Fremdsprachen]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, -1, 1, 0, ZielDB)
 .Fields![Überschrift].AllowZeroLength = True
 Call CPropA(QTb + ß, "Überschrift", """L"" = Liebe(r), ""H"" = Hallo", 0, 109, 315, 0, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "", 0, 109, 5310, 0, 0, ZielDB)
 .Fields![Vorname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vorname", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Nachname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Nachname", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Anschrift].AllowZeroLength = True
 Call CPropA(QTb + ß, "Anschrift", "", 0, 109, 4365, 0, 0, ZielDB)
 .Fields![KVNr].AllowZeroLength = True
 Call CPropA(QTb + ß, "KVNr", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Telefon].AllowZeroLength = True
 Call CPropA(QTb + ß, "Telefon", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Telefax].AllowZeroLength = True
 Call CPropA(QTb + ß, "Telefax", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![E_Mail].AllowZeroLength = True
 Call CPropA(QTb + ß, "E_Mail", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Zulassungsgebiet].AllowZeroLength = True
 Call CPropA(QTb + ß, "Zulassungsgebiet", "", 0, 109, 2805, 0, 0, ZielDB)
 .Fields![Arzttyp].AllowZeroLength = True
 Call CPropA(QTb + ß, "Arzttyp", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Gemeinschaftspraxis mit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Gemeinschaftspraxis mit", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![Schwerpunkt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schwerpunkt", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Zusatzbezeichnung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Zusatzbezeichnung", "", 0, 109, 4080, 0, 0, ZielDB)
 .Fields![Bemerkung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bemerkung", "", 0, Null, 17550, 0, 0, ZielDB)
 .Fields![Beme].AllowZeroLength = True
 Call CPropA(QTb + ß, "Beme", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![Sprechstunden].AllowZeroLength = True
 Call CPropA(QTb + ß, "Sprechstunden", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![von _ bis].AllowZeroLength = True
 Call CPropA(QTb + ß, "von _ bis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Internetadressen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Internetadressen", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields![Behandlung in Fremdsprachen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Behandlung in Fremdsprachen", "", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("Rollstuhlgerechte Praxis", 10, 20)
 .Fields![Rollstuhlgerechte Praxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "Rollstuhlgerechte Praxis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Verkehrsmittel", 10, 50)
 .Fields![Verkehrsmittel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Verkehrsmittel", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Linie", 10, 20)
 .Fields![Linie].AllowZeroLength = True
 Call CPropA(QTb + ß, "Linie", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Haltestelle Parkplätze", 10, 50)
 .Fields![Haltestelle Parkplätze].AllowZeroLength = True
 Call CPropA(QTb + ß, "Haltestelle Parkplätze", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Wegbeschreibung", 12, 0)
 .Fields![Wegbeschreibung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Wegbeschreibung", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Entfernung zur Praxis", 10, 50)
 .Fields![Entfernung zur Praxis].AllowZeroLength = True
 Call CPropA(QTb + ß, "Entfernung zur Praxis", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Zahl", 4, 4)
 Call CPropA(QTb + ß, "Zahl", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("nichtmehr", 1, 1)
 Call PropA(QTb + ß, "nichtmehr", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "nichtmehr", "Arzt nicht mehr im Verzeichnis", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Titel", 10, 50)
 .Fields![Titel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Titel", "", 0, 109, 1230, 0, 0, ZielDB)
 .Fields.Append .CreateField("Geschlecht", 10, 50)
 .Fields![Geschlecht].AllowZeroLength = True
 Call CPropA(QTb + ß, "Geschlecht", "", 0, 109, 1200, 0, 0, ZielDB)
 .Fields.Append .CreateField("Straße", 10, 50)
 .Fields![Straße].AllowZeroLength = True
 Call CPropA(QTb + ß, "Straße", "", 0, 109, 2805, 0, 0, ZielDB)
 .Fields.Append .CreateField("PLZ", 10, 10)
 .Fields![Plz].AllowZeroLength = True
 Call CPropA(QTb + ß, "PLZ", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ort", 10, 50)
 .Fields![Ort].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ort", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DMPT2", 3, 2)
 Call CPropA(QTb + ß, "DMPT2", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DMPT1", 3, 2)
 Call CPropA(QTb + ß, "DMPT1", "", Null, 109, -1, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[KVNr]"
  sql = sql + ",[Name]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![KVNr].Name) = "KVNR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [KVNr] on [" + QTb + ß + "] ("
  sql = sql + "[KVNr]"
  sql = sql + ",[Nachname]"
  sql = sql + ",[Vorname]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAHausaerzte/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAHausaerzte

Function MachAfuerLeistExpArchiv()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "fuerLeistExpArchiv"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(50)"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [SchGr]  TEXT(2)"
 sql = sql + ", [archiviert]  DATE"
 sql = sql + ", [Protokoll]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos", Null, Null, -1, 1, -1, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Freitext mit Anfangswert(en) für Nach- und Vorname", 0, 109, 1965, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!pat_id, wird vom Programm gesucht", Null, 109, -1, 0, 0, ZielDB)
 .Fields![SchGr].AllowZeroLength = True
 Call CPropA(QTb + ß, "SchGr", "Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "archiviert", "Datum der Übertragung in Tabelle Leistungsexport", Null, Null, 2115, 0, 0, ZielDB)
 .Fields![Protokoll].AllowZeroLength = True
 Call CPropA(QTb + ß, "Protokoll", "eingetragene Leistungen", 0, Null, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![id].Name) = "ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAfuerLeistExpArchiv/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAfuerLeistExpArchiv

Function MachAfuerLeistExp_Kopie()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "fuerLeistExp Kopie"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(50)"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [SchGr]  TEXT(2)"
 sql = sql + ", [Status]  TEXT(50)"
 sql = sql + ", [Protokoll]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos", Null, Null, -1, 1, -1, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Freitext mit Anfangswert(en) für Nach- und Vorname", 0, 109, 1350, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!pat_id", Null, 109, -1, 0, 0, ZielDB)
 .Fields![SchGr].AllowZeroLength = True
 Call CPropA(QTb + ß, "SchGr", "Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Status].AllowZeroLength = True
 Call CPropA(QTb + ß, "Status", "gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Protokoll].AllowZeroLength = True
 Call CPropA(QTb + ß, "Protokoll", "eingetragene Leistungen", 0, Null, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![id].Name) = "ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAfuerLeistExp_Kopie/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAfuerLeistExp_Kopie

Function MachAfuerLeistExp()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "fuerLeistExp"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(50)"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [SchGr]  TEXT(2)"
 sql = sql + ", [Status]  TEXT(50)"
 sql = sql + ", [Protokoll]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos", Null, Null, -1, 0, -1, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Freitext mit Anfangswert(en) für Nach- und Vorname", 0, 109, 6345, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!pat_id, wird vom Programm gesucht", Null, 109, -1, 0, 0, ZielDB)
 .Fields![SchGr].AllowZeroLength = True
 Call CPropA(QTb + ß, "SchGr", "Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Status].AllowZeroLength = True
 Call CPropA(QTb + ß, "Status", "gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat", 0, 109, 1530, 0, 0, ZielDB)
 .Fields![Protokoll].AllowZeroLength = True
 Call CPropA(QTb + ß, "Protokoll", "eingetragene Leistungen", 0, Null, 3000, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![id].Name) = "ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAfuerLeistExp/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAfuerLeistExp

Function MachAfuerDiagExpArchiv()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "fuerDiagExpArchiv"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [ID] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(50)"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [ICD]  TEXT(10)"
 sql = sql + ", [Diagnose]  TEXT(90)"
 sql = sql + ", [archiviert]  DATE"
 sql = sql + ", [Protokoll]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos", Null, Null, -1, 1, -1, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Freitext mit Anfangswert(en) für Nach- und Vorname", 0, 109, 2640, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!pat_id, wird vom Programm gesucht", Null, 109, 750, 0, 0, ZielDB)
 .Fields![ICD].AllowZeroLength = True
 Call CPropA(QTb + ß, "ICD", "ICD 10", 0, 109, 840, 0, 0, ZielDB)
 .Fields![Diagnose].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diagnose", "Diagnose Text", 0, 109, 3075, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "archiviert", "Datum der Übertragung in Tabelle Leistungsexport", Null, Null, 1050, 0, 0, ZielDB)
 .Fields![Protokoll].AllowZeroLength = True
 Call CPropA(QTb + ß, "Protokoll", "eingetragene Leistungen", 0, Null, 1005, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [pat_ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![suche].Name) = "SUCHE" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Suche] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ",[ICD]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAfuerDiagExpArchiv/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAfuerDiagExpArchiv

Function MachAfuerDiagExp_Kopie()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "fuerDiagExp Kopie"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [ID] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(50)"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [ICD]  TEXT(10)"
 sql = sql + ", [Diagnose]  TEXT(255)"
 sql = sql + ", [Status]  TEXT(50)"
 sql = sql + ", [Protokoll]  MEMO"
 sql = sql + ", [nurQuart]  BIT"
 sql = sql + ", [Zeitpunkt]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos", Null, Null, -1, 1, -1, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Freitext mit Anfangswert(en) für Nach- und Vorname", 0, 109, 2565, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!pat_id, wird vom Programm gesucht", Null, 109, -1, 0, 0, ZielDB)
 .Fields![ICD].AllowZeroLength = True
 Call CPropA(QTb + ß, "ICD", "ICD 10", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Diagnose].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diagnose", "Diagnose Text", 0, 109, 5595, 0, 0, ZielDB)
 .Fields![Status].AllowZeroLength = True
 Call CPropA(QTb + ß, "Status", "gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat", 0, 109, 1530, 0, 0, ZielDB)
 .Fields![Protokoll].AllowZeroLength = True
 Call CPropA(QTb + ß, "Protokoll", "eingetragene Leistungen", 0, Null, 3000, 0, 0, ZielDB)
 Call PropA(QTb + ß, "nurQuart", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "nurQuart", "ja = nur für ein Quartal", Null, 106, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zeitpunkt", "Zeitpunkt der gewünschten Diagnose", Null, Null, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [pat_ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![suche].Name) = "SUCHE" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Suche] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ",[ICD]"
  sql = sql + ",[Diagnose]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAfuerDiagExp_Kopie/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAfuerDiagExp_Kopie

Function MachAfuerDiagExp()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "fuerDiagExp"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [ID] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(50)"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [ICD]  TEXT(10)"
 sql = sql + ", [Diagnose]  TEXT(255)"
 sql = sql + ", [Status]  TEXT(50)"
 sql = sql + ", [Protokoll]  MEMO"
 sql = sql + ", [nurQuart]  BIT"
 sql = sql + ", [Zeitpunkt]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos", Null, Null, -1, 1, -1, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Freitext mit Anfangswert(en) für Nach- und Vorname", 0, 109, 2565, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!pat_id, wird vom Programm gesucht", Null, 109, -1, 0, 0, ZielDB)
 .Fields![ICD].AllowZeroLength = True
 Call CPropA(QTb + ß, "ICD", "ICD 10", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Diagnose].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diagnose", "Diagnose Text", 0, 109, 5595, 0, 0, ZielDB)
 .Fields![Status].AllowZeroLength = True
 Call CPropA(QTb + ß, "Status", "gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat", 0, 109, 1530, 0, 0, ZielDB)
 .Fields![Protokoll].AllowZeroLength = True
 Call CPropA(QTb + ß, "Protokoll", "eingetragene Leistungen", 0, Null, 3000, 0, 0, ZielDB)
 Call PropA(QTb + ß, "nurQuart", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "nurQuart", "ja = nur für ein Quartal", Null, 106, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zeitpunkt", "Zeitpunkt der gewünschten Diagnose", Null, Null, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [pat_ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![suche].Name) = "SUCHE" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Suche] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ",[ICD]"
  sql = sql + ",[Diagnose]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAfuerDiagExp/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAfuerDiagExp

Function MachAFremdlabor()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Fremdlabor"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [HbA1c]  TEXT(20)"
 sql = sql + ", [Glucose]  TEXT(20)"
 sql = sql + ", [Crea]  TEXT(20)"
 sql = sql + ", [Hst]  TEXT(20)"
 sql = sql + ", [GGT]  TEXT(20)"
 sql = sql + ", [GPT]  TEXT(20)"
 sql = sql + ", [INR]  TEXT(20)"
 sql = sql + ", [Leuko]  TEXT(20)"
 sql = sql + ", [Hb]  TEXT(20)"
 sql = sql + ", [MCH]  TEXT(20)"
 sql = sql + ", [Thrombo]  TEXT(20)"
 sql = sql + ", [CRP]  TEXT(20)"
 sql = sql + ", [K]  TEXT(20)"
 sql = sql + ", [Na]  TEXT(20)"
 sql = sql + ", [TSH]  TEXT(20)"
 sql = sql + ", [fT4]  TEXT(20)"
 sql = sql + ", [fT3]  TEXT(20)"
 sql = sql + ", [Chol]  TEXT(20)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![HbA1c].AllowZeroLength = True
 Call CPropA(QTb + ß, "HbA1c", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Glucose].AllowZeroLength = True
 Call CPropA(QTb + ß, "Glucose", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Crea].AllowZeroLength = True
 Call CPropA(QTb + ß, "Crea", "", 0, 109, 600, 0, 0, ZielDB)
 .Fields![Hst].AllowZeroLength = True
 Call CPropA(QTb + ß, "Hst", "", 0, 109, 435, 0, 0, ZielDB)
 .Fields![GGT].AllowZeroLength = True
 Call CPropA(QTb + ß, "GGT", "", 0, 109, 570, 0, 0, ZielDB)
 .Fields![GPT].AllowZeroLength = True
 Call CPropA(QTb + ß, "GPT", "", 0, 109, 555, 0, 0, ZielDB)
 .Fields![INR].AllowZeroLength = True
 Call CPropA(QTb + ß, "INR", "", 0, 109, 480, 0, 0, ZielDB)
 .Fields![Leuko].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leuko", "", 0, 109, 735, 0, 0, ZielDB)
 .Fields![Hb].AllowZeroLength = True
 Call CPropA(QTb + ß, "Hb", "", 0, 109, 630, 0, 0, ZielDB)
 .Fields![MCH].AllowZeroLength = True
 Call CPropA(QTb + ß, "MCH", "", 0, 109, 585, 0, 0, ZielDB)
 .Fields![Thrombo].AllowZeroLength = True
 Call CPropA(QTb + ß, "Thrombo", "", 0, 109, 1005, 0, 0, ZielDB)
 .Fields![CRP].AllowZeroLength = True
 Call CPropA(QTb + ß, "CRP", "", 0, 109, 555, 0, 0, ZielDB)
 .Fields![k].AllowZeroLength = True
 Call CPropA(QTb + ß, "K", "", 0, 109, 525, 0, 0, ZielDB)
 .Fields![Na].AllowZeroLength = True
 Call CPropA(QTb + ß, "Na", "", 0, 109, 705, 0, 0, ZielDB)
 .Fields![TSH].AllowZeroLength = True
 Call CPropA(QTb + ß, "TSH", "", 0, 109, 540, 0, 0, ZielDB)
 .Fields![fT4].AllowZeroLength = True
 Call CPropA(QTb + ß, "fT4", "", 0, 109, 435, 0, 0, ZielDB)
 .Fields![fT3].AllowZeroLength = True
 Call CPropA(QTb + ß, "fT3", "", 0, 109, 435, 0, 0, ZielDB)
 .Fields![Chol].AllowZeroLength = True
 Call CPropA(QTb + ß, "Chol", "", 0, 109, 585, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("LDL", 10, 20)
 .Fields![LDL].AllowZeroLength = True
 Call CPropA(QTb + ß, "LDL", "", 0, 109, 525, 0, 0, ZielDB)
 .Fields.Append .CreateField("HDL", 10, 20)
 .Fields![hdl].AllowZeroLength = True
 Call CPropA(QTb + ß, "HDL", "", 0, 109, 540, 0, 0, ZielDB)
 .Fields.Append .CreateField("TG", 10, 20)
 .Fields![TG].AllowZeroLength = True
 Call CPropA(QTb + ß, "TG", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Hsre", 10, 20)
 .Fields![Hsre].AllowZeroLength = True
 Call CPropA(QTb + ß, "Hsre", "", 0, 109, 570, 0, 0, ZielDB)
 .Fields.Append .CreateField("LeuU", 10, 20)
 .Fields![LeuU].AllowZeroLength = True
 Call CPropA(QTb + ß, "LeuU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("EryU", 10, 20)
 .Fields![EryU].AllowZeroLength = True
 Call CPropA(QTb + ß, "EryU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("EiwU", 10, 20)
 .Fields![EiwU].AllowZeroLength = True
 Call CPropA(QTb + ß, "EiwU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BakU", 10, 20)
 .Fields![BakU].AllowZeroLength = True
 Call CPropA(QTb + ß, "BakU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("NitU", 10, 20)
 .Fields![NitU].AllowZeroLength = True
 Call CPropA(QTb + ß, "NitU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BiliU", 10, 20)
 .Fields![BiliU].AllowZeroLength = True
 Call CPropA(QTb + ß, "BiliU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("UrobU", 10, 20)
 .Fields![UrobU].AllowZeroLength = True
 Call CPropA(QTb + ß, "UrobU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("SedU", 10, 50)
 .Fields![SedU].AllowZeroLength = True
 Call CPropA(QTb + ß, "SedU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("KetU", 10, 20)
 .Fields![KetU].AllowZeroLength = True
 Call CPropA(QTb + ß, "KetU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("ZucU", 10, 20)
 .Fields![ZucU].AllowZeroLength = True
 Call CPropA(QTb + ß, "ZucU", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("AlbU", 10, 20)
 .Fields![AlbU].AllowZeroLength = True
 Call CPropA(QTb + ß, "AlbU", "", 0, 109, 600, 0, 0, ZielDB)
 .Fields.Append .CreateField("GEiwU", 10, 20)
 .Fields![GEiwU].AllowZeroLength = True
 Call CPropA(QTb + ß, "GEiwU", "", 0, 109, 780, 0, 0, ZielDB)
 .Fields.Append .CreateField("Micral", 10, 20)
 .Fields![Micral].AllowZeroLength = True
 Call CPropA(QTb + ß, "Micral", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DokName", 12, 0)
 .Fields![DokName].AllowZeroLength = True
 Call CPropA(QTb + ß, "DokName", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DokPfad", 12, 0)
 .Fields![DokPfad].AllowZeroLength = True
 Call CPropA(QTb + ß, "DokPfad", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DokGroe", 4, 4)
 Call CPropA(QTb + ß, "DokGroe", "Größe in Bytes", Null, 109, -1, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[DokPfad]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Pat_ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFremdlabor/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFremdlabor

Function MachAFormulare()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Formulare"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FormID]  INTEGER CONSTRAINT [FormID] PRIMARY KEY"
 sql = sql + ", [Form_Abk]  TEXT(8)"
 sql = sql + ", [FormBez]  MEMO"
 sql = sql + ", [FormVorl]  MEMO"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FormID", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields![Form_Abk].AllowZeroLength = True
 Call CPropA(QTb + ß, "Form_Abk", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![FormBez].AllowZeroLength = True
 Call CPropA(QTb + ß, "FormBez", "", 0, Null, 7665, 0, 0, ZielDB)
 .Fields![FormVorl].AllowZeroLength = True
 Call CPropA(QTb + ß, "FormVorl", "", 0, Null, 8565, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Zeitpunkt der Aktualisierung", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in BDT-Datei", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Form_Abk]"
  sql = sql + ",[FormBez]"
  sql = sql + ",[FormVorl]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FormBez].Name) = "FORMBEZ" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FormBez] on [" + QTb + ß + "] ("
  sql = sql + "[FormBez]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFormulare/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFormulare

Function MachAFormInhKopf()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "FormInhKopf"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FoID]  INTEGER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [Form_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [AbsPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ", [Satzart]  TEXT(4)"
 sql = sql + ", [Satzlänge]  TEXT(7)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FoID", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 675, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Form_ID", "", Null, 109, 945, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AbsPos", "Zeile in der BDT-Datei", Null, 109, 840, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 .Fields![Satzart].AllowZeroLength = True
 Call CPropA(QTb + ß, "Satzart", "8000", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Satzlänge].AllowZeroLength = True
 Call CPropA(QTb + ß, "Satzlänge", "8100", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[Form_ID]"
  sql = sql + ",[ZeitPunkt]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFormInhKopf/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFormInhKopf

Function MachAFormInhFeld()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "FormInhFeld"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FoID]  INTEGER"
 sql = sql + ", [Nr]  SMALLINT"
 sql = sql + ", [FeldNr]  SMALLINT"
 sql = sql + ", [FeldVW]  INTEGER"
 sql = sql + ", [FeldInhVW]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "FoID", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Nr", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "FeldNr", "", Null, 109, 780, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "FeldVW", "", Null, 109, 900, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "FeldInhVW", "", Null, 109, 1200, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FoID].Name) = "FOID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FoID] on [" + QTb + ß + "] ("
  sql = sql + "[FoID]"
  sql = sql + ",[FeldVW]"
  sql = sql + ",[FeldNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFormInhFeld/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFormInhFeld

Function MachAFormInhaltForm_Abk()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "FormInhaltForm_Abk"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Form_AbkVW]  COUNTER CONSTRAINT [Form_AbkVW] PRIMARY KEY"
 sql = sql + ", [Form_Abk]  TEXT(20) CONSTRAINT [Form_Abk] UNIQUE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Form_AbkVW", "", Null, Null, 1440, 0, 0, ZielDB)
 .Fields![Form_Abk].AllowZeroLength = True
 Call CPropA(QTb + ß, "Form_Abk", "", 0, 109, 1200, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Form_Abk].Name) = "FORM_ABK" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [Form_Abk] on [" + QTb + ß + "] ("
  sql = sql + "[Form_Abk]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFormInhaltForm_Abk/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFormInhaltForm_Abk

Function MachAFormInhaltFeldInh()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "FormInhaltFeldInh"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FeldInhVW]  COUNTER CONSTRAINT [FeldInhVW] PRIMARY KEY"
 sql = sql + ", [FeldInh]  MEMO"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "FeldInhVW", "", Null, Null, 1200, 0, 0, ZielDB)
 .Fields![FeldInh].AllowZeroLength = True
 Call CPropA(QTb + ß, "FeldInh", "", 0, Null, 16995, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordinalziffer der Einlesung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FeldInh].Name) = "FELDINH" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FeldInh] on [" + QTb + ß + "] ("
  sql = sql + "[FeldInh]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFormInhaltFeldInh/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFormInhaltFeldInh

Function MachAFormInhaltFeld()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "FormInhaltFeld"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FeldVW]  COUNTER CONSTRAINT [PRIMARY] PRIMARY KEY"
 sql = sql + ", [Feld]  MEMO"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FeldVW", "", Null, Null, 900, 0, 0, ZielDB)
 .Fields![Feld].AllowZeroLength = True
 Call CPropA(QTb + ß, "Feld", "", 0, Null, 2835, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordinalziffer der Einlesung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Feld].Name) = "FELD" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Feld] on [" + QTb + ß + "] ("
  sql = sql + "[Feld]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFormInhaltFeld/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFormInhaltFeld

Function MachAFeldProp()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "FeldProp"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Prop]  TEXT(50)"
 sql = sql + ", [ReiheVorher]  INTEGER"
 sql = sql + ", [TypVorher]  INTEGER"
 sql = sql + ", [WertVorher]  TEXT(50)"
 sql = sql + ", [ReiheNachher]  INTEGER"
 sql = sql + ", [TypNachher]  INTEGER"
 sql = sql + ", [WertNachher]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 .Fields![Prop].AllowZeroLength = True
 Call CPropA(QTb + ß, "Prop", "", 0, 109, 1575, 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ReiheVorher", "", Null, 109, 1350, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "TypVorher", "", Null, 109, 1140, 0, 0, ZielDB)
 .Fields![WertVorher].AllowZeroLength = True
 Call CPropA(QTb + ß, "WertVorher", "", 0, 109, 1245, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ReiheNachher", "", Null, 109, 1500, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "TypNachher", "", Null, 109, 1290, 0, 0, ZielDB)
 .Fields![WertNachher].AllowZeroLength = True
 Call CPropA(QTb + ß, "WertNachher", "", 0, 109, 2070, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Prop].Name) = "PROP" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [prop] on [" + QTb + ß + "] ("
  sql = sql + "[Prop]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFeldProp/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFeldProp

Function MachAFeldLaengen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "FeldLaengen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[TabName]  TEXT(100)"
 sql = sql + ", [FeldName]  TEXT(100)"
 sql = sql + ", [Typ]  TEXT(1)"
 sql = sql + ", [Verkn]  BIT"
 sql = sql + ", [Größe]  INTEGER"
 sql = sql + ", [MaxFüll]  INTEGER"
 sql = sql + ", [Diff]  INTEGER"
 sql = sql + ", [Zahl]  INTEGER"
 sql = sql + ", [ProdLeer]  INTEGER"
 sql = sql + ", [ProdDaten]  INTEGER"
 sql = sql + ", [ProdCont]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 .Fields![TabName].AllowZeroLength = True
 Call CPropA(QTb + ß, "TabName", "", 0, Null, 1605, 1, 0, ZielDB)
 .Fields![FeldName].AllowZeroLength = True
 Call CPropA(QTb + ß, "FeldName", "", 0, Null, 2790, 0, 0, ZielDB)
 .Fields![Typ].AllowZeroLength = True
 Call CPropA(QTb + ß, "Typ", "", 0, Null, 495, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Verkn", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Verkn", "", Null, Null, 705, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Größe", "", Null, Null, 735, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "MaxFüll", "", Null, Null, 915, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Diff", "", Null, Null, 630, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zahl", "", Null, Null, 675, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ProdLeer", "", Null, Null, 1155, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ProdDaten", "", Null, Null, 1155, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ProdCont", "", Null, Null, 1035, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFeldLaengen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFeldLaengen

Function MachAFaelle()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Faelle"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [Quartal]  TEXT(5)"
 sql = sql + ", [Nachname]  TEXT(50)"
 sql = sql + ", [Vorname]  TEXT(50)"
 sql = sql + ", [lfdnr]  INTEGER"
 sql = sql + ", [TMFNr]  TEXT(20)"
 sql = sql + ", [VKNr]  TEXT(10)"
 sql = sql + ", [BhFB]  DATE"
 sql = sql + ", [BhFE1]  DATE"
 sql = sql + ", [BhFE2]  DATE"
 sql = sql + ", [f4202]  TEXT(5)"
 sql = sql + ", [ausgst]  DATE"
 sql = sql + ", [KtrAbrB]  TEXT(3)"
 sql = sql + ", [AbrAr]  TEXT(2)"
 sql = sql + ", [lVorl]  DATE"
 sql = sql + ", [IK]  TEXT(50)"
 sql = sql + ", [KVKs]  TEXT(10)"
 sql = sql + ", [KVKserg]  TEXT(5)"
 sql = sql + ", [GebOr]  TEXT(2)"
 sql = sql + ", [AbrGb]  TEXT(5)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Primärschlüssel", Null, Null, 675, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 .Fields![Quartal].AllowZeroLength = True
 Call CPropA(QTb + ß, "Quartal", "4101", 0, 109, 855, 0, 0, ZielDB)
 .Fields![Nachname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Nachname", "3101", 0, 109, 1365, 0, 0, ZielDB)
 .Fields![Vorname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vorname", "3102", 0, 109, 1170, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "lfdnr", "laufende Fallnummer", Null, 109, 585, 0, 0, ZielDB)
 .Fields![TMFNr].AllowZeroLength = True
 Call CPropA(QTb + ß, "TMFNr", "4144 Fallnummer in Turbomed", 0, 109, 1305, 0, 0, ZielDB)
 .Fields![VKNr].AllowZeroLength = True
 Call CPropA(QTb + ß, "VKNr", "4104", 0, 109, 975, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "BhFB", "4150", Null, Null, 1110, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "BhFE1", "4151", Null, Null, 1110, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "BhFE2", "4152", Null, Null, 1110, 0, 0, ZielDB)
 .Fields![f4202].AllowZeroLength = True
 Call CPropA(QTb + ß, "f4202", "4202", 0, 109, 630, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ausgst", "4102 ('ausgestellt am')", Null, Null, 1110, 0, 0, ZielDB)
 .Fields![KtrAbrB].AllowZeroLength = True
 Call CPropA(QTb + ß, "KtrAbrB", "4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))", 0, 109, 630, 0, 0, ZielDB)
 .Fields![AbrAr].AllowZeroLength = True
 Call CPropA(QTb + ß, "AbrAr", "4107, Abrechnungsart (1 = Primärkassen)", 0, 109, 630, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "lVorl", "4109, letzte Vorlage", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![IK].AllowZeroLength = True
 Call CPropA(QTb + ß, "IK", "4111 (auch patientenbezogen)", 0, 109, 885, 0, 0, ZielDB)
 .Fields![KVKs].AllowZeroLength = True
 Call CPropA(QTb + ß, "KVKs", "4112", 0, 109, 645, 0, 0, ZielDB)
 .Fields![KVKserg].AllowZeroLength = True
 Call CPropA(QTb + ß, "KVKserg", "4113", 0, 109, 960, 0, 0, ZielDB)
 .Fields![GebOr].AllowZeroLength = True
 Call CPropA(QTb + ß, "GebOr", "4121, Gebührenordnung (1 = BMÄ, 2)", 0, 109, 630, 0, 0, ZielDB)
 .Fields![AbrGb].AllowZeroLength = True
 Call CPropA(QTb + ß, "AbrGb", "4122, Abrechnungsgebiet (07 = Diabetes)", 0, 109, 630, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("f4206", 10, 15)
 .Fields![f4206].AllowZeroLength = True
 Call CPropA(QTb + ß, "f4206", "4206, ?", 0, 109, 630, 0, 0, ZielDB)
 .Fields.Append .CreateField("ÜwText", 12, 0)
 .Fields![ÜwText].AllowZeroLength = True
 Call CPropA(QTb + ß, "ÜwText", "4209: Auftrags- / erläuternder Text zur Überweisung", 0, Null, 870, 0, 0, ZielDB)
 .Fields.Append .CreateField("f4210", 10, 5)
 .Fields![f4210].AllowZeroLength = True
 Call CPropA(QTb + ß, "f4210", "4210, unbekannt", 0, 109, 630, 0, 0, ZielDB)
 .Fields.Append .CreateField("statNuller", 10, 20)
 .Fields![statNuller].AllowZeroLength = True
 Call CPropA(QTb + ß, "statNuller", "4216, nu bei Musterfrau 16 Nuller", 0, 109, 1830, 0, 0, ZielDB)
 .Fields.Append .CreateField("ÜbwV", 10, 10)
 .Fields![ÜbwV].AllowZeroLength = True
 Call CPropA(QTb + ß, "ÜbwV", "4218, überwiesen von", 0, 109, 885, 0, 0, ZielDB)
 .Fields.Append .CreateField("AndÜw", 10, 10)
 .Fields![AndÜw].AllowZeroLength = True
 Call CPropA(QTb + ß, "AndÜw", "4219, anderer Überweiser", 0, 109, 885, 0, 0, ZielDB)
 .Fields.Append .CreateField("Übw", 10, 8)
 .Fields![Übw].AllowZeroLength = True
 Call CPropA(QTb + ß, "Übw", "4218 oder 4219, je nachdem, was befüllt", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("ÜWZiel", 10, 50)
 .Fields![ÜWZiel].AllowZeroLength = True
 Call CPropA(QTb + ß, "ÜWZiel", "4220", 0, 109, 3915, 0, 0, ZielDB)
 .Fields.Append .CreateField("ÜWNNr", 10, 10)
 .Fields![ÜWNNr].AllowZeroLength = True
 Call CPropA(QTb + ß, "ÜWNNr", "4231(4): KV-Nummer des Überweisers", 0, 109, 885, 0, 0, ZielDB)
 .Fields.Append .CreateField("ÜWNaN", 10, 50)
 .Fields![ÜWNaN].AllowZeroLength = True
 Call CPropA(QTb + ß, "ÜWNaN", "4231(3): Nachname des Überweisers", 0, 109, 1185, 0, 0, ZielDB)
 .Fields.Append .CreateField("ÜWTit", 10, 39)
 .Fields![ÜWTit].AllowZeroLength = True
 Call CPropA(QTb + ß, "ÜWTit", "4231(3): Titel des Überweisers", 0, 109, 900, 0, 0, ZielDB)
 .Fields.Append .CreateField("ÜWVor", 10, 35)
 .Fields![ÜWVor].AllowZeroLength = True
 Call CPropA(QTb + ß, "ÜWVor", "4231(2): Vorname des Überweisers", 0, 109, 1365, 0, 0, ZielDB)
 .Fields.Append .CreateField("ÜWVsw", 10, 10)
 .Fields![ÜWVsw].AllowZeroLength = True
 Call CPropA(QTb + ß, "ÜWVsw", "4231(2b): Vorsatzwort des Überweisers", 0, 109, 870, 0, 0, ZielDB)
 .Fields.Append .CreateField("statklasse", 10, 5)
 .Fields![statKlasse].AllowZeroLength = True
 Call CPropA(QTb + ß, "statklasse", "4236", 0, 109, 1095, 0, 0, ZielDB)
 .Fields.Append .CreateField("f4237", 10, 5)
 .Fields![f4237].AllowZeroLength = True
 Call CPropA(QTb + ß, "f4237", "4237, ? (nur bei Musterw)", 0, 109, 630, 0, 0, ZielDB)
 .Fields.Append .CreateField("statBehTage", 4, 4)
 Call CPropA(QTb + ß, "statBehTage", "4238", Null, 109, 1335, 0, 0, ZielDB)
 .Fields.Append .CreateField("SchGr", 10, 5)
 .Fields![SchGr].AllowZeroLength = True
 Call CPropA(QTb + ß, "SchGr", "4239, Scheingruppe", 0, 109, 735, 0, 0, ZielDB)
 .Fields.Append .CreateField("Weiterbeh", 10, 70)
 .Fields![Weiterbeh].AllowZeroLength = True
 Call CPropA(QTb + ß, "Weiterbeh", "4243, Weiterbehandelnder", 0, 109, 1140, 0, 0, ZielDB)
 .Fields.Append .CreateField("PGeb", 10, 70)
 .Fields![PGeb].AllowZeroLength = True
 Call CPropA(QTb + ß, "PGeb", "4401, Praxisgebühr", 0, 109, 3180, 0, 0, ZielDB)
 .Fields.Append .CreateField("PGebErg", 10, 50)
 .Fields![PGebErg].AllowZeroLength = True
 Call CPropA(QTb + ß, "PGebErg", "4402, ?", 0, 109, 990, 0, 0, ZielDB)
 .Fields.Append .CreateField("Mahnfrist", 10, 8)
 .Fields![Mahnfrist].AllowZeroLength = True
 Call CPropA(QTb + ß, "Mahnfrist", "4403, Mahnfrist bis", 0, 109, 1020, 0, 0, ZielDB)
 .Fields.Append .CreateField("GOÄKatNr", 10, 10)
 .Fields![GOÄKatNr].AllowZeroLength = True
 Call CPropA(QTb + ß, "GOÄKatNr", "4580 (1): Katalog-Nummer", 0, 109, 1110, 0, 0, ZielDB)
 .Fields.Append .CreateField("GOÄKatName", 10, 30)
 .Fields![GOÄKatName].AllowZeroLength = True
 Call CPropA(QTb + ß, "GOÄKatName", "4580 (2): Privat-Abrechnungskatalog", 0, 109, 1455, 0, 0, ZielDB)
 .Fields.Append .CreateField("AdNam", 10, 28)
 .Fields![AdNam].AllowZeroLength = True
 Call CPropA(QTb + ß, "AdNam", "4602(1) Name bei Musterfrau", 0, 109, 1485, 0, 0, ZielDB)
 .Fields.Append .CreateField("AdStr", 10, 40)
 .Fields![AdStr].AllowZeroLength = True
 Call CPropA(QTb + ß, "AdStr", "4602(2) Straße bei Musterfrau", 0, 109, 1590, 0, 0, ZielDB)
 .Fields.Append .CreateField("AdPlz", 10, 10)
 .Fields![AdPlz].AllowZeroLength = True
 Call CPropA(QTb + ß, "AdPlz", "4602(3) PLZ bei Musterfrau", 0, 109, 735, 0, 0, ZielDB)
 .Fields.Append .CreateField("AdOrt", 10, 30)
 .Fields![AdOrt].AllowZeroLength = True
 Call CPropA(QTb + ß, "AdOrt", "4602(4) Ort bei Musterfrau", 0, 109, 1605, 0, 0, ZielDB)
 .Fields.Append .CreateField("BhFE", 8, 8)
 Call CPropA(QTb + ß, "BhFE", "4604, Behandlungsfall: Ende, bei Privatpatienten", Null, Null, 1110, 0, 0, ZielDB)
 .Fields.Append .CreateField("s8000", 10, 4)
 .Fields![s8000].AllowZeroLength = True
 Call CPropA(QTb + ß, "s8000", "8000, ???", 0, 109, 660, 0, 0, ZielDB)
 .Fields.Append .CreateField("s8100", 10, 7)
 .Fields![s8100].AllowZeroLength = True
 Call CPropA(QTb + ß, "s8100", "8100, ???", 0, 109, 675, 0, 0, ZielDB)
 .Fields.Append .CreateField("AktZeit", 8, 8)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 .Fields.Append .CreateField("Fanf", 8, 8)
 Call CPropA(QTb + ß, "Fanf", "Fallanfang", Null, Null, 1920, 0, 0, ZielDB)
 .Fields.Append .CreateField("altQuart", 10, 5)
 .Fields![altQuart].AllowZeroLength = True
 Call CPropA(QTb + ß, "altQuart", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("QAnf", 8, 8)
 Call CPropA(QTb + ß, "QAnf", "Quartalsanfang", Null, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("QEnd", 8, 8)
 Call CPropA(QTb + ß, "QEnd", "Quartalsende", Null, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("QS", 10, 5)
 .Fields![QS].AllowZeroLength = True
 Call CPropA(QTb + ß, "QS", "Quartal des Behandlungsfallbeginns sortiert", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("QT", 10, 5)
 .Fields![QT].AllowZeroLength = True
 Call CPropA(QTb + ß, "QT", "Quartal des Behandlungsfallbeginns", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("StByte", 4, 4)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("absPos", 4, 4)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, -1, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![aktF].Name) = "AKTF" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [AktF] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[BhFB] DESC"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[Quartal] DESC"
  sql = sql + ",[BhFB]"
  sql = sql + ",[BhFE1]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![BhFB].Name) = "BHFB" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [BhFB] on [" + QTb + ß + "] ("
  sql = sql + "[BhFB]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Fanf].Name) = "FANF" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FanF] on [" + QTb + ß + "] ("
  sql = sql + "[Fanf]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![pQ].Name) = "PQ" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [pQ] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[Quartal]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Quartal].Name) = "QUARTAL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Quartal] on [" + QTb + ß + "] ("
  sql = sql + "[Quartal]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![SchGr].Name) = "SCHGR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [SchGr] on [" + QTb + ß + "] ("
  sql = sql + "[SchGr]"
  sql = sql + ",[Nachname]"
  sql = sql + ",[Vorname]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![VKNr].Name) = "VKNR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [vknr] on [" + QTb + ß + "] ("
  sql = sql + "[VKNr]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAFaelle/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAFaelle

Function MachAEintragsZahlen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "EintragsZahlen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Beginn]  DATE CONSTRAINT [Beginn] PRIMARY KEY"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ", [Zp1]  DATE"
 sql = sql + ", [Zp2]  DATE"
 sql = sql + ", [Zp3]  DATE"
 sql = sql + ", [Zp4]  DATE"
 sql = sql + ", [Zp5]  DATE"
 sql = sql + ", [Zp6]  DATE"
 sql = sql + ", [Zp7]  DATE"
 sql = sql + ", [Zp8]  DATE"
 sql = sql + ", [Fallzahl]  INTEGER"
 sql = sql + ", [Sekunden]  INTEGER"
 sql = sql + ", [Datei]  TEXT(120)"
 sql = sql + ", [DateiAend]  DATE"
 sql = sql + ", [SpeicherZt]  DATE"
 sql = sql + ", [TabellenLöschen]  BIT"
 sql = sql + ", [ZurücksetzenLAktDat]  BIT"
 sql = sql + ", [Pat_IDVon]  TEXT(6)"
 sql = sql + ", [Pat_IDbis]  TEXT(6)"
 sql = sql + ", [VorladenFFI]  BIT"
 sql = sql + ", [SammelInsert]  BIT"
 sql = sql + ", [LaborDirektEinlesen]  BIT"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Beginn", "", Null, Null, 1920, 1, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Statusbyte", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp1", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp2", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp3", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp4", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp5", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp6", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp7", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zp8", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Fallzahl", "", Null, 109, 915, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Sekunden", "", Null, 109, 1110, 0, 0, ZielDB)
 .Fields![Datei].AllowZeroLength = True
 Call CPropA(QTb + ß, "Datei", "", 0, 109, 3285, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "DateiAend", "", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "SpeicherZt", "", Null, Null, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "TabellenLöschen", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "TabellenLöschen", "", Null, 106, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "ZurücksetzenLAktDat", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "ZurücksetzenLAktDat", "", Null, 106, -1, 0, 0, ZielDB)
 .Fields![Pat_IDVon].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_IDVon", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Pat_IDBis].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pat_IDbis", "", 0, 109, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "VorladenFFI", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "VorladenFFI", "", Null, 106, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "SammelInsert", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "SammelInsert", "", Null, 106, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "LaborDirektEinlesen", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "LaborDirektEinlesen", "", Null, 106, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("LaborDirektNeu", 1, 1)
 Call PropA(QTb + ß, "LaborDirektNeu", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "LaborDirektNeu", "", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("LaborQuerVerb", 1, 1)
 Call PropA(QTb + ß, "LaborQuerVerb", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "LaborQuerVerb", "", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("LaborQuerNeu", 1, 1)
 Call PropA(QTb + ß, "LaborQuerNeu", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "LaborQuerNeu", "", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("AlterTab", 1, 1)
 Call PropA(QTb + ß, "AlterTab", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AlterTab", "", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obmitEmails", 1, 1)
 Call PropA(QTb + ß, "obmitEmails", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obmitEmails", "", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("LaborPfadBeispiel", 12, 0)
 .Fields![LaborPfadBeispiel].AllowZeroLength = True
 Call CPropA(QTb + ß, "LaborPfadBeispiel", "", 0, Null, 2280, 0, 0, ZielDB)
 .Fields.Append .CreateField("obVglMitLetzterEinlesung", 1, 1)
 .Fields.Append .CreateField("Email-Datei", 10, 120)
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAEintragsZahlen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAEintragsZahlen

Function MachAEintraege_Arten()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Eintraege Arten"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Art]  TEXT(50)"
 sql = sql + ", [obAutom]  BIT"
 sql = sql + ", [Erklärung]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, -1, 0, 0, ZielDB)
 .Fields![art].AllowZeroLength = True
 Call CPropA(QTb + ß, "Art", "", 0, 109, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "obAutom", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obAutom", "ob Art auch automatisch aus Formular entsteht", Null, 106, -1, 0, 0, ZielDB)
 .Fields![Erklärung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Erklärung", "Erklärung für die Art", 0, Null, 11880, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAEintraege_Arten/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAEintraege_Arten

Function MachAEintraege()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Eintraege"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [Art]  TEXT(15)"
 sql = sql + ", [Inhalt]  MEMO"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [QS]  TEXT(5)"
 sql = sql + ", [QT]  TEXT(5)"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 720, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![art].AllowZeroLength = True
 Call CPropA(QTb + ß, "Art", "6330", 0, 109, 810, 0, 0, ZielDB)
 .Fields![Inhalt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Inhalt", "8480", 0, Null, 13080, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, 840, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![QS].AllowZeroLength = True
 Call CPropA(QTb + ß, "QS", "Quartal des Behandlungsfallbeginns sortiert", 0, 109, -1, 0, 0, ZielDB)
 .Fields![QT].AllowZeroLength = True
 Call CPropA(QTb + ß, "QT", "Quartal des Behandlungsfallbeginns", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[Art]"
  sql = sql + ",[ZeitPunkt] DESC"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAEintraege/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAEintraege

Function MachAEinstellungen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Einstellungen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Formular]  TEXT(50)"
 sql = sql + ", [Abfrage für Formular]  TEXT(100)"
 sql = sql + ", [ID für Formular]  INTEGER"
 sql = sql + ", [DatensatzNr]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, -1, 1, 0, ZielDB)
 .Fields![Formular].AllowZeroLength = True
 Call CPropA(QTb + ß, "Formular", "Name des Formulars", 0, 109, 1605, 0, 0, ZielDB)
 .Fields![Abfrage für Formular].AllowZeroLength = True
 Call CPropA(QTb + ß, "Abfrage für Formular", "Name der Abfrage, die zuletzt für das Formular ""Anamnesebogen"" verwendet wurde", 0, 109, 4710, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ID für Formular", "Pat_ID in dieser Abfrage", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "DatensatzNr", "Datensatz-Nr. in dieser Abfrage", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Formular].Name) = "FORMULAR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Formular] on [" + QTb + ß + "] ("
  sql = sql + "[Formular]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![ID für Formular].Name) = "ID FÜR FORMULAR" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ID für Formular] on [" + QTb + ß + "] ("
  sql = sql + "[ID für Formular]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAEinstellungen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAEinstellungen

Function MachAEinfuegefehler()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Einfuegefehler"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Name]  TEXT(255)"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [archiviert]  DATE"
 sql = sql + ", [Protokoll]  MEMO"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "", 0, Null, 1095, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "", Null, Null, 1110, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "", Null, Null, 750, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "archiviert", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Protokoll].AllowZeroLength = True
 Call CPropA(QTb + ß, "Protokoll", "", 0, Null, 2745, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAEinfuegefehler/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAEinfuegefehler

Function MachAEBM2000plus1()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "EBM2000plus1"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Leistung]  TEXT(50)"
 sql = sql + ", [Titel]  TEXT(200)"
 sql = sql + ", [Punktwert]  MONEY"
 sql = sql + ", [Euro]  MONEY"
 sql = sql + ", [Bericht]  BIT"
 sql = sql + ", [Text]  MEMO"
 sql = sql + ", [Betr]  BIT"
 sql = sql + ", [Schul]  BIT"
 sql = sql + ", [Typ1]  BIT"
 sql = sql + ", [Typ2]  BIT"
 sql = sql + ", [Gest]  BIT"
 sql = sql + ", [DFS]  BIT"
 sql = sql + ", [DMP]  BIT"
 sql = sql + ", [AOK]  BIT"
 sql = sql + ", [BKK]  BIT"
 sql = sql + ", [BKN]  BIT"
 sql = sql + ", [EK]  BIT"
 sql = sql + ", [IKK]  BIT"
 sql = sql + ", [LKK]  BIT"
 sql = sql + ", [Üw]  BIT"
 sql = sql + ", [Insulin]  BIT"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call PropTabA(QTb + ß, "FrozenColumns", 3, 2, ZielDB)
 .Fields![Leistung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leistung", "Leistungsziffer", 0, 109, 825, 0, 0, ZielDB)
 .Fields![Titel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Titel", "Kurztext", 0, 109, 7470, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Punktwert", "Punktwert", Null, 109, 720, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Euro", "Format", dbText, "#,##0.00 €;-#,##0.00 €", ZielDB)
 Call CPropA(QTb + ß, "Euro", "€", Null, 109, 750, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Bericht", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Bericht", "Berichtspflicht", Null, 106, 600, 0, 0, ZielDB)
 .Fields![Text].AllowZeroLength = True
 Call CPropA(QTb + ß, "Text", "restlicher Leistungstext", 0, Null, 14805, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Betr", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Betr", "Betreuung", Null, 106, 540, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Schul", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Schul", "Schulung", Null, 106, 690, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Typ1", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Typ1", "", Null, 106, 600, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Typ2", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Typ2", "", Null, 106, 600, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Gest", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Gest", "", Null, 106, 570, 0, 0, ZielDB)
 Call PropA(QTb + ß, "DFS", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "DFS", "", Null, 106, 540, 0, 0, ZielDB)
 Call PropA(QTb + ß, "DMP", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "DMP", "", Null, 106, 585, 0, 0, ZielDB)
 Call PropA(QTb + ß, "AOK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AOK", "", Null, 106, 570, 0, 0, ZielDB)
 Call PropA(QTb + ß, "BKK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "BKK", "", Null, 106, 555, 0, 0, ZielDB)
 Call PropA(QTb + ß, "BKN", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "BKN", "", Null, 106, 555, 0, 0, ZielDB)
 Call PropA(QTb + ß, "EK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "EK", "", Null, 106, 405, 0, 0, ZielDB)
 Call PropA(QTb + ß, "IKK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "IKK", "", Null, 106, 480, 0, 0, ZielDB)
 Call PropA(QTb + ß, "LKK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "LKK", "", Null, 106, 540, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Üw", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Üw", "Überweisung durch HA nötig", Null, 106, 450, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Insulin", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Insulin", "", Null, 106, 555, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("ICT", 1, 1)
 Call PropA(QTb + ß, "ICT", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "ICT", "", Null, 106, 465, 0, 0, ZielDB)
 .Fields.Append .CreateField("CSII", 1, 1)
 Call PropA(QTb + ß, "CSII", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "CSII", "", Null, 106, 540, 0, 0, ZielDB)
 .Fields.Append .CreateField("Erst", 1, 1)
 Call PropA(QTb + ß, "Erst", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Erst", "", Null, 106, 495, 0, 0, ZielDB)
 .Fields.Append .CreateField("Folge", 1, 1)
 Call PropA(QTb + ß, "Folge", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Folge", "", Null, 106, 690, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Leistung].Name) = "LEISTUNG" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Leistung] on [" + QTb + ß + "] ("
  sql = sql + "[Leistung]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Titel].Name) = "TITEL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Titel] on [" + QTb + ß + "] ("
  sql = sql + "[Titel]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAEBM2000plus1/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAEBM2000plus1

Function MachAEBM2000plus()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "EBM2000plus"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Leistung]  TEXT(50)"
 sql = sql + ", [Titel]  TEXT(200)"
 sql = sql + ", [Punktwert]  MONEY"
 sql = sql + ", [Euro]  MONEY"
 sql = sql + ", [Bericht]  BIT"
 sql = sql + ", [Text]  MEMO"
 sql = sql + ", [Betr]  BIT"
 sql = sql + ", [Schul]  BIT"
 sql = sql + ", [Typ1]  BIT"
 sql = sql + ", [Typ2]  BIT"
 sql = sql + ", [Gest]  BIT"
 sql = sql + ", [DFS]  BIT"
 sql = sql + ", [DMP]  BIT"
 sql = sql + ", [AOK]  BIT"
 sql = sql + ", [BKK]  BIT"
 sql = sql + ", [BKN]  BIT"
 sql = sql + ", [EK]  BIT"
 sql = sql + ", [IKK]  BIT"
 sql = sql + ", [LKK]  BIT"
 sql = sql + ", [Üw]  BIT"
 sql = sql + ", [Insulin]  BIT"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call PropTabA(QTb + ß, "FrozenColumns", 3, 2, ZielDB)
 .Fields![Leistung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Leistung", "Leistungsziffer", 0, 109, 825, 0, 0, ZielDB)
 .Fields![Titel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Titel", "Kurztext", 0, 109, 7470, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Punktwert", "Punktwert", Null, 109, 720, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Euro", "Format", dbText, "#,##0.00 €;-#,##0.00 €", ZielDB)
 Call CPropA(QTb + ß, "Euro", "€", Null, 109, 750, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Bericht", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Bericht", "Berichtspflicht", Null, 106, 600, 0, 0, ZielDB)
 .Fields![Text].AllowZeroLength = True
 Call CPropA(QTb + ß, "Text", "restlicher Leistungstext", 0, Null, 14805, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Betr", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Betr", "Betreuung", Null, 106, 540, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Schul", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Schul", "Schulung", Null, 106, 690, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Typ1", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Typ1", "", Null, 106, 600, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Typ2", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Typ2", "", Null, 106, 600, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Gest", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Gest", "", Null, 106, 570, 0, 0, ZielDB)
 Call PropA(QTb + ß, "DFS", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "DFS", "", Null, 106, 540, 0, 0, ZielDB)
 Call PropA(QTb + ß, "DMP", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "DMP", "", Null, 106, 585, 0, 0, ZielDB)
 Call PropA(QTb + ß, "AOK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "AOK", "", Null, 106, 570, 0, 0, ZielDB)
 Call PropA(QTb + ß, "BKK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "BKK", "", Null, 106, 555, 0, 0, ZielDB)
 Call PropA(QTb + ß, "BKN", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "BKN", "", Null, 106, 555, 0, 0, ZielDB)
 Call PropA(QTb + ß, "EK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "EK", "", Null, 106, 405, 0, 0, ZielDB)
 Call PropA(QTb + ß, "IKK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "IKK", "", Null, 106, 480, 0, 0, ZielDB)
 Call PropA(QTb + ß, "LKK", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "LKK", "", Null, 106, 540, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Üw", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Üw", "Überweisung durch HA nötig", Null, 106, 450, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Insulin", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Insulin", "", Null, 106, 555, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("ICT", 1, 1)
 Call PropA(QTb + ß, "ICT", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "ICT", "", Null, 106, 465, 0, 0, ZielDB)
 .Fields.Append .CreateField("CSII", 1, 1)
 Call PropA(QTb + ß, "CSII", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "CSII", "", Null, 106, 540, 0, 0, ZielDB)
 .Fields.Append .CreateField("Erst", 1, 1)
 Call PropA(QTb + ß, "Erst", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Erst", "", Null, 106, 495, 0, 0, ZielDB)
 .Fields.Append .CreateField("Folge", 1, 1)
 Call PropA(QTb + ß, "Folge", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Folge", "", Null, 106, 690, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Leistung].Name) = "LEISTUNG" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Leistung] on [" + QTb + ß + "] ("
  sql = sql + "[Leistung]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Titel].Name) = "TITEL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Titel] on [" + QTb + ß + "] ("
  sql = sql + "[Titel]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAEBM2000plus/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAEBM2000plus

Function MachADokumente_abgehakt()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Dokumente abgehakt"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[DokPfad]  TEXT(140)"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [abgehakt]  BIT"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 .Fields![DokPfad].AllowZeroLength = True
 Call CPropA(QTb + ß, "DokPfad", "", 0, 109, 10035, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call PropA(QTb + ß, "abgehakt", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "abgehakt", "", Null, 106, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![DokPfad].Name) = "DOKPFAD" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [DokPfad] on [" + QTb + ß + "] ("
  sql = sql + "[DokPfad]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachADokumente_abgehakt/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachADokumente_abgehakt

Function MachADokumente()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Dokumente"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [DokPfad]  TEXT(150)"
 sql = sql + ", [DokArt]  TEXT(40)"
 sql = sql + ", [DokName]  TEXT(255)"
 sql = sql + ", [Quelldatum]  DATE"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [DokGroe]  INTEGER"
 sql = sql + ", [QS]  TEXT(5)"
 sql = sql + ", [QT]  TEXT(5)"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 720, 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![DokPfad].AllowZeroLength = True
 Call CPropA(QTb + ß, "DokPfad", "", 0, 109, 9570, 0, 0, ZielDB)
 .Fields![DokArt].AllowZeroLength = True
 Call CPropA(QTb + ß, "DokArt", "", 0, 109, 780, 0, 0, ZielDB)
 .Fields![DokName].AllowZeroLength = True
 Call CPropA(QTb + ß, "DokName", "", 0, 109, 5970, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Quelldatum", "Datum, auf das sich das Dokument bezieht", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, 840, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "DokGroe", "Dokument-Größe", Null, 109, 975, 0, 0, ZielDB)
 .Fields![QS].AllowZeroLength = True
 Call CPropA(QTb + ß, "QS", "Quartal des Behandlungsfallbeginns sortiert", 0, 109, -1, 0, 0, ZielDB)
 .Fields![QT].AllowZeroLength = True
 Call CPropA(QTb + ß, "QT", "Quartal des Behandlungsfallbeginns", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt]"
  sql = sql + ",[DokArt]"
  sql = sql + ",[DokName]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![DokName].Name) = "DOKNAME" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [DokName] on [" + QTb + ß + "] ("
  sql = sql + "[DokName]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![DokPfad].Name) = "DOKPFAD" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [DokPfad] on [" + QTb + ß + "] ("
  sql = sql + "[DokPfad]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![PIDokPfad].Name) = "PIDOKPFAD" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [PIDokPfad] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[DokPfad]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Quelldatum].Name) = "QUELLDATUM" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Quelldatum] on [" + QTb + ß + "] ("
  sql = sql + "[Quelldatum]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Zeitpunkt].Name) = "ZEITPUNKT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ZeitPunkt] on [" + QTb + ß + "] ("
  sql = sql + "[ZeitPunkt]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachADokumente/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachADokumente

Function MachADMP_Uschr$()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "DMP-Uschr"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Pat_id]  INTEGER"
 sql = sql + ", [U1]  DATE"
 sql = sql + ", [U2]  DATE"
 sql = sql + ", [U3]  DATE"
 sql = sql + ", [Arztwechsel]  BIT"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "Bezug auf Namen", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "U1", "vorliegendes Blatt mit DMP-Unterschrift", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "U2", "vorliegendes 2. Blatt mit DMP-Unterschrift", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "U3", "vorliegendes 3. Blatt mit DMP-Unterschrift", Null, Null, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Arztwechsel", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Arztwechsel", "danach Arztwechsel", Null, 106, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Pat_id] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachADMP_Uschr/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachADMP_Uschr

Function MachAdmpreihe()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "dmpreihe"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Abk]  TEXT(30)"
 sql = sql + ", [Art]  TEXT(2)"
 sql = sql + ", [KarteiDatum]  DATE"
 sql = sql + ", [exportiert]  DATE"
 sql = sql + ", [DokuDatum]  DATE"
 sql = sql + ", [obvoll]  BIT"
 sql = sql + ", [NachName]  TEXT(20)"
 sql = sql + ", [VorName]  TEXT(20)"
 sql = sql + ", [GebDat]  DATE"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Abk", "Abkürzung der DMP-Art", 0, 109, -1, 0, 0, ZielDB)
 .Fields![art].AllowZeroLength = True
 Call CPropA(QTb + ß, "Art", "ED = Erstdoku, FD = Folgedoku", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "KarteiDatum", "Datum des Karteikarteneintrags der Dokumentation", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "exportiert", "Datum des Exports", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "DokuDatum", "Datum der Dokumentation", Null, Null, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "obvoll", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obvoll", "ob vollständig", Null, 106, -1, 0, 0, ZielDB)
 .Fields![Nachname].AllowZeroLength = True
 Call CPropA(QTb + ß, "NachName", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Vorname].AllowZeroLength = True
 Call CPropA(QTb + ß, "VorName", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "GebDat", "", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungzeit", Null, Null, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAdmpreihe/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAdmpreihe

Function MachADiagnosenExport()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "DiagnosenExport"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [ID] PRIMARY KEY"
 sql = sql + ", [Name]  TEXT(50)"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [ICD]  TEXT(10)"
 sql = sql + ", [Diagnose]  TEXT(255)"
 sql = sql + ", [Status]  TEXT(50)"
 sql = sql + ", [Protokoll]  MEMO"
 sql = sql + ", [nurQuart]  BIT"
 sql = sql + ", [Zeitpunkt]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos", Null, Null, -1, 0, -1, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "Freitext mit Anfangswert(en) für Nach- und Vorname", 0, 109, 6345, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!pat_id, wird vom Programm gesucht", Null, 109, -1, 0, 0, ZielDB)
 .Fields![ICD].AllowZeroLength = True
 Call CPropA(QTb + ß, "ICD", "ICD 10", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Diagnose].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diagnose", "Diagnose Text", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Status].AllowZeroLength = True
 Call CPropA(QTb + ß, "Status", "gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat", 0, 109, 1530, 0, 0, ZielDB)
 .Fields![Protokoll].AllowZeroLength = True
 Call CPropA(QTb + ß, "Protokoll", "eingetragene Leistungen", 0, Null, 3000, 0, 0, ZielDB)
 Call PropA(QTb + ß, "nurQuart", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "nurQuart", "ja = nur für ein Quartal", Null, 106, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Zeitpunkt", "Zeitpunkt der gewünschten Diagnose", Null, Null, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [pat_ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![suche].Name) = "SUCHE" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Suche] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ",[ICD]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachADiagnosenExport/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachADiagnosenExport

Function MachADiagnosen_exportiert()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Diagnosen exportiert"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [ICD]  TEXT(15)"
 sql = sql + ", [Diagnose]  TEXT(70)"
 sql = sql + ", [übertragen]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "Primärschlüssel wegen Reihenfolge", Null, Null, -1, 1, -1, ZielDB)
 Call CPropA(QTb + ß, "Datum", "Leistungsdatum", Null, Null, 1545, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "-> Namen!Pat_id", Null, 109, 750, 0, 0, ZielDB)
 .Fields![ICD].AllowZeroLength = True
 Call CPropA(QTb + ß, "ICD", "ICD-Nummer der Diagnose", 0, 109, 1050, 0, 0, ZielDB)
 .Fields![Diagnose].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diagnose", "Text der Diagnose", 0, 109, 1665, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "übertragen", """"", übertragen", Null, Null, 2115, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![id].Name) = "ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachADiagnosen_exportiert/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachADiagnosen_exportiert

Function MachADiagnosen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Diagnosen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID1]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [FID]  INTEGER"
 sql = sql + ", [Pat_id]  INTEGER"
 sql = sql + ", [GesName]  TEXT(50)"
 sql = sql + ", [DiagDatum]  DATE"
 sql = sql + ", [DiagSicherheit]  TEXT(1)"
 sql = sql + ", [DiagText]  MEMO"
 sql = sql + ", [DiagSeite]  TEXT(1)"
 sql = sql + ", [DiagAttr]  TEXT(60)"
 sql = sql + ", [ICD]  TEXT(10)"
 sql = sql + ", [obDauer]  BIT"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID1", "", Null, Null, 780, 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 675, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "Bezug auf Anamneseblattt", Null, 109, 750, 0, 0, ZielDB)
 .Fields![GesName].AllowZeroLength = True
 Call CPropA(QTb + ß, "GesName", "", 0, 109, 2790, 0, 0, ZielDB)
 Call PropA(QTb + ß, "DiagDatum", "Format", dbText, "General Date", ZielDB)
 Call CPropA(QTb + ß, "DiagDatum", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![DiagSicherheit].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiagSicherheit", "6003", 0, 109, 780, 0, 0, ZielDB)
 .Fields![DiagText].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiagText", "", 0, Null, 5670, 0, 0, ZielDB)
 .Fields![DiagSeite].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiagSeite", "6004", 0, 109, 1080, 0, 0, ZielDB)
 .Fields![DiagAttr].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiagAttr", "6006", 0, 109, 915, 0, 0, ZielDB)
 .Fields![ICD].AllowZeroLength = True
 Call CPropA(QTb + ß, "ICD", "", 0, 109, 765, 0, 0, ZielDB)
 Call PropA(QTb + ß, "obDauer", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obDauer", "ob Dauerdiagnose", Null, 109, 960, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ",[DiagDatum]"
  sql = sql + ",[DiagSicherheit]"
  sql = sql + ",[DiagSeite]"
  sql = sql + ",[DiagAttr]"
  sql = sql + ",[DiagText]"
  sql = sql + ",[ICD]"
  sql = sql + ",[obDauer]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![DiagSuch].Name) = "DIAGSUCH" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [DiagSuch] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ",[ICD]"
  sql = sql + ",[DiagSicherheit]"
  sql = sql + ",[DiagSeite]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![DiagText].Name) = "DIAGTEXT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [DiagText] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ",[DiagText]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachADiagnosen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachADiagnosen

Function MachADateien()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Dateien"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Dateiname]  TEXT(128)"
 sql = sql + ", [Pfad]  TEXT(150)"
 sql = sql + ", [Größe]  FLOAT"
 sql = sql + ", [Änderung]  DATE"
 sql = sql + ", [Endung]  TEXT(20)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, 570, 0, 0, ZielDB)
 .Fields![DateiName].AllowZeroLength = True
 Call CPropA(QTb + ß, "Dateiname", "", 0, 109, 6720, 0, 0, ZielDB)
 .Fields![Pfad].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pfad", "Absolute Pfad", 0, 109, 3780, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Größe", "Dateigröße in Bytes", Null, 109, 1200, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Änderung", "letzte Änderung", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Endung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Endung", "Dateiendung", 0, 109, 870, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Name].Name) = "NAME" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Name] on [" + QTb + ß + "] ("
  sql = sql + "[Dateiname]"
  sql = sql + ",[Pfad]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachADateien/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachADateien

Function MachABriefe()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Briefe"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [Pfad]  TEXT(134)"
 sql = sql + ", [Art]  TEXT(40)"
 sql = sql + ", [Name]  TEXT(150)"
 sql = sql + ", [Typ]  TEXT(50)"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [DokGroe]  INTEGER"
 sql = sql + ", [QS]  TEXT(5)"
 sql = sql + ", [QT]  TEXT(5)"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 720, 1, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Pfad].AllowZeroLength = True
 Call CPropA(QTb + ß, "Pfad", "", 0, 109, 10005, 0, 0, ZielDB)
 .Fields![art].AllowZeroLength = True
 Call CPropA(QTb + ß, "Art", "", 0, 109, 525, 0, 0, ZielDB)
 .Fields![Name].AllowZeroLength = True
 Call CPropA(QTb + ß, "Name", "", 0, 109, 5655, 0, 0, ZielDB)
 .Fields![Typ].AllowZeroLength = True
 Call CPropA(QTb + ß, "Typ", "", 0, 109, 540, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "DokGroe", "Größe der Datei", Null, 109, -1, 0, 0, ZielDB)
 .Fields![QS].AllowZeroLength = True
 Call CPropA(QTb + ß, "QS", "Quartal des Behandlungsfallbeginns sortiert", 0, 109, -1, 0, 0, ZielDB)
 .Fields![QT].AllowZeroLength = True
 Call CPropA(QTb + ß, "QT", "Quartal des Behandlungsfallbeginns", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt]"
  sql = sql + ",[Name]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachABriefe/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachABriefe

Function MachABezirke()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Bezirke"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[id]  INTEGER CONSTRAINT [id] PRIMARY KEY"
 sql = sql + ", [Bezirk]  TEXT(15)"
 sql = sql + ", [Nr]  INTEGER"
 sql = sql + ", [MaxNr]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 .Fields![Bezirk].AllowZeroLength = True
 Dim fld As DAO.Field
 End With
 With ZielDB
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachABezirke/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachABezirke

Function MachAAugenbefunde()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Augenbefunde"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[ID]  COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [Datum]  DATE"
 sql = sql + ", [Retinopathie]  MEMO"
 sql = sql + ", [klassifikation]  TEXT(150)"
 sql = sql + ", [Maculopathie]  TEXT(255)"
 sql = sql + ", [Sonstiges]  MEMO"
 sql = sql + ", [Visus re]  TEXT(10)"
 sql = sql + ", [Visus li]  TEXT(50)"
 sql = sql + ", [KontrolleinMonaten]  SINGLE"
 sql = sql + ", [Augenarzt]  TEXT(70)"
 sql = sql + ", [DokName]  TEXT(150)"
 sql = sql + ", [DokPfad]  MEMO"
 sql = sql + ", [verglichen]  DATE"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, 0, ZielDB)
 Call CPropA(QTb + ß, "ID", "", Null, Null, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Datum", "", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Retinopathie].AllowZeroLength = True
 Call CPropA(QTb + ß, "Retinopathie", "", 0, Null, 17550, 0, 0, ZielDB)
 .Fields![klassifikation].AllowZeroLength = True
 Call CPropA(QTb + ß, "klassifikation", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Maculopathie].AllowZeroLength = True
 Call CPropA(QTb + ß, "Maculopathie", "", 0, 109, 600, 0, 0, ZielDB)
 .Fields![Sonstiges].AllowZeroLength = True
 Call CPropA(QTb + ß, "Sonstiges", "", 0, Null, 435, 0, 0, ZielDB)
 .Fields![Visus re].AllowZeroLength = True
 Call CPropA(QTb + ß, "Visus re", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Visus li].AllowZeroLength = True
 Call CPropA(QTb + ß, "Visus li", "", 0, 109, -1, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "KontrolleinMonaten", "", Null, 109, -1, 0, 0, ZielDB)
 .Fields![Augenarzt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Augenarzt", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![DokName].AllowZeroLength = True
 Call CPropA(QTb + ß, "DokName", "", 0, 109, 5040, 0, 0, ZielDB)
 .Fields![DokPfad].AllowZeroLength = True
 Call CPropA(QTb + ß, "DokPfad", "", 0, Null, 10065, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "verglichen", "", Null, Null, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[DokPfad]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Text].Name) = "TEXT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Text] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[Retinopathie]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAAugenbefunde/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAAugenbefunde

Function MachAAU()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "AU"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[FID]  INTEGER"
 sql = sql + ", [Pat_ID]  INTEGER"
 sql = sql + ", [ZeitPunkt]  DATE"
 sql = sql + ", [Beginn]  TEXT(8)"
 sql = sql + ", [Ende]  TEXT(8)"
 sql = sql + ", [ICDs]  TEXT(100)"
 sql = sql + ", [absPos]  INTEGER"
 sql = sql + ", [AktZeit]  DATE"
 sql = sql + ", [StByte]  INTEGER"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "FID", "Fall-Bezug", Null, 109, 720, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "Pat_ID", "3000", Null, 109, 765, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "ZeitPunkt", "6200 + 6201", Null, Null, 1920, 0, 0, ZielDB)
 .Fields![Beginn].AllowZeroLength = True
 Call CPropA(QTb + ß, "Beginn", "6285 1. Hälfte", 0, 109, 990, 0, 0, ZielDB)
 .Fields![Ende].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ende", "6285 2. Hälfte", 0, 109, 990, 0, 0, ZielDB)
 .Fields![ICDs].AllowZeroLength = True
 Call CPropA(QTb + ß, "ICDs", "6286", 0, 109, 4830, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "absPos", "Zeile in der BDT-Datei", Null, 109, 885, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "StByte", "Ordnungsnummer der Datenübertragung", Null, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_ID]"
  sql = sql + ",[ZeitPunkt]"
  sql = sql + ",[Beginn]"
  sql = sql + ",[Ende]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![FID].Name) = "FID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [FID] on [" + QTb + ß + "] ("
  sql = sql + "[FID]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAAU/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAAU

Function MachAAnamnesebogen()
 Dim sql$, ZielDB As DAO.Database, gefunden%
 Dim rel As DAO.Relation, RelLöschRunde%, aktBez%
 Const QTb$ = "Anamnesebogen"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Erstelle die Tabelle " + QTb + ß)
 Set ZielDB = OpenDatabase(ZielDbS)
 sql = "create table [" + ZielDbS + "].[" + QTb + ß + "] ("
 sql = sql + "[Pat_id]  INTEGER"
 sql = sql + ", [Nachname]  TEXT(21)"
 sql = sql + ", [Vorname]  TEXT(19)"
 sql = sql + ", [NVorsatz]  TEXT(5)"
 sql = sql + ", [Titel]  TEXT(12)"
 sql = sql + ", [Anrede]  TEXT(4)"
 sql = sql + ", [GebDat]  DATE"
 sql = sql + ", [Tkz]  BIT"
 sql = sql + ", [Versicherungsart]  TEXT(2)"
 sql = sql + ", [Diabetestyp]  TEXT(25)"
 sql = sql + ", [Diabetes seit]  TEXT(152)"
 sql = sql + ", [Tabletten seit]  TEXT(66)"
 sql = sql + ", [Insulin seit]  TEXT(115)"
 sql = sql + ", [Grund für Vorstellung]  MEMO"
 sql = sql + ", [Familienanamnese]  MEMO"
 sql = sql + ", [Größe]  FLOAT"
 sql = sql + ", [Gewicht]  FLOAT"
 sql = sql + ", [Tendenz]  TEXT(3)"
 sql = sql + ", [DiabetesMedikament 1]  TEXT(48)"
 sql = sql + ", [DiabetesMedikament 1 Menge]  TEXT(112)"
 sql = sql + ", [DiabetesMedikament 2]  TEXT(50)"
 sql = sql + ")"
 On Error Resume Next
 DoCmd.Close acTable, QTb + ß
 With ZielDB
  gefunden = True
  On Error Resume Next
  If Not UCase$(ZielDB.TableDefs(QTb + ß).Name) = UCase$(QTb + ß) Then gefunden = 0
  On Error GoTo fehler
  If gefunden Then
   If ZielDB.Name = "." Then If Application.CurrentData.AllTables(QTb + ß).IsLoaded Then DoCmd.Close acTable, QTb + ß
   If Not SammelAufruf Then
    For RelLöschRunde = 1 To 10
     For Each rel In ZielDB.Relations
      If rel.Table = QTb + ß Or rel.ForeignTable = QTb + ß Then
       If Not aktBez Then
        mitBez = True ' alle Aufrufe
        aktBez = True ' aktuelle function
       End If
       Err.Clear
       ZielDB.Relations.Delete rel.Name
      End If
     Next rel
    Next RelLöschRunde
   End If ' not SammelAufruf
   Err.Clear
   Call .Execute("drop Table [" + ZielDbS + "].[" + QTb + ß + "]")
  End If
  Call .Execute(Replace(sql, "()", ""))
  If Not gefunden Then
   Dim ZName$
   ZName = ZielDB.Name
   ZielDB.Close
   Set ZielDB = OpenDatabase(ZName)
  End If
 End With
 With ZielDB.TableDefs(QTb + ß)
 Call PropTabA(QTb + ß, "OrderByOn", 1, -1, ZielDB)
 Call CPropA(QTb + ß, "Pat_id", "", Null, 109, 510, 1, 0, ZielDB)
 .Fields![Nachname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Nachname", "-", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Vorname].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vorname", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![NVorsatz].AllowZeroLength = True
 Call CPropA(QTb + ß, "NVorsatz", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields![Titel].AllowZeroLength = True
 Call CPropA(QTb + ß, "Titel", "", 0, 109, 300, 0, 0, ZielDB)
 .Fields![Anrede].AllowZeroLength = True
 Call CPropA(QTb + ß, "Anrede", "", 0, 109, 495, 0, 0, ZielDB)
 Call CPropA(QTb + ß, "GebDat", ", geb.", 0, Null, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Tkz", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Tkz", "Tod-Kennzeichen", Null, 106, -1, 0, 0, ZielDB)
 .Fields![Versicherungsart].AllowZeroLength = True
 Call CPropA(QTb + ß, "Versicherungsart", "", 0, 109, 1050, 0, 0, ZielDB)
 .Fields![Diabetestyp].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diabetestyp", "^Diabetes Typ", 0, 109, 420, 0, 0, ZielDB)
 .Fields![Diabetes seit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diabetes seit", "<seit", 0, 109, 600, 0, 0, ZielDB)
 .Fields![Tabletten seit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Tabletten seit", ", Tabletten seit", 0, 109, 555, 0, 0, ZielDB)
 .Fields![Insulin seit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Insulin seit", ", Insulin seit", 0, 109, 600, 0, 0, ZielDB)
 .Fields![Grund für Vorstellung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Grund für Vorstellung", "^:", 0, Null, 1665, 0, 0, ZielDB)
 .Fields![Familienanamnese].AllowZeroLength = True
 Call CPropA(QTb + ß, "Familienanamnese", "^:", 0, Null, -1, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Größe", "Format", dbText, "Fixed", ZielDB)
 Call PropA(QTb + ß, "Größe", "InputMask", dbText, "#,##", ZielDB)
 Call CPropA(QTb + ß, "Größe", "^:", Null, 109, 735, 0, 0, ZielDB)
 Call PropA(QTb + ß, "Gewicht", "Format", dbText, "General Number", ZielDB)
 Call PropA(QTb + ß, "Gewicht", "InputMask", dbText, "###.#", ZielDB)
 Call PropA(QTb + ß, "Gewicht", "DecimalPlaces", 2, 1, ZielDB)
 Call CPropA(QTb + ß, "Gewicht", ",:", Null, 109, 930, 0, 0, ZielDB)
 .Fields![Tendenz].AllowZeroLength = True
 Call CPropA(QTb + ß, "Tendenz", "<, Tendenz", 0, 109, 330, 0, 0, ZielDB)
 .Fields![DiabetesMedikament 1].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiabetesMedikament 1", "^Letzte Diabetesmedikation:", 0, 109, -1, 0, 0, ZielDB)
 .Fields![DiabetesMedikament 1 Menge].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiabetesMedikament 1 Menge", "<", 0, 109, -1, 0, 0, ZielDB)
 .Fields![DiabetesMedikament 2].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiabetesMedikament 2", "<,", 0, 109, -1, 0, 0, ZielDB)
 Dim fld As DAO.Field
 .Fields.Append .CreateField("DiabetesMedikament 2 Menge", 10, 126)
 .Fields![DiabetesMedikament 2 Menge].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiabetesMedikament 2 Menge", "<", 0, 109, 3015, 0, 0, ZielDB)
 .Fields.Append .CreateField("DiabetesMedikament 3", 10, 37)
 .Fields![DiabetesMedikament 3].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiabetesMedikament 3", "<,", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DiabetesMedikament 3 Menge", 10, 38)
 .Fields![DiabetesMedikament 3 Menge].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiabetesMedikament 3 Menge", "<", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DiabetesMedikament 4", 10, 35)
 .Fields![DiabetesMedikament 4].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiabetesMedikament 4", "<,", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DiabetesMedikament 4 Menge", 10, 35)
 .Fields![DiabetesMedikament 4 Menge].AllowZeroLength = True
 Call CPropA(QTb + ß, "DiabetesMedikament 4 Menge", "<,", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Insulinpumpe", 1, 1)
 Call PropA(QTb + ß, "Insulinpumpe", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Insulinpumpe", "^:", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Insulinpumpe seit", 10, 250)
 .Fields![Insulinpumpe seit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Insulinpumpe seit", "<seit", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Insulinpumpe Marke", 10, 73)
 .Fields![Insulinpumpe Marke].AllowZeroLength = True
 Call CPropA(QTb + ß, "Insulinpumpe Marke", "<, Marke:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Broteinheiten gesamt", 10, 49)
 .Fields![Broteinheiten gesamt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Broteinheiten gesamt", "^Broteinheiten:gesamt", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Broteinheiten früh", 10, 30)
 .Fields![Broteinheiten früh].AllowZeroLength = True
 Call CPropA(QTb + ß, "Broteinheiten früh", "<, früh", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Broteinheiten ZM früh", 10, 15)
 .Fields![Broteinheiten ZM früh].AllowZeroLength = True
 Call CPropA(QTb + ß, "Broteinheiten ZM früh", "<, Zwischenmahlzeit vormittags", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Broteinheiten mittags", 10, 86)
 .Fields![Broteinheiten mittags].AllowZeroLength = True
 Call CPropA(QTb + ß, "Broteinheiten mittags", "<, mittags", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Broteinheiten nachmittags", 10, 29)
 .Fields![Broteinheiten nachmittags].AllowZeroLength = True
 Call CPropA(QTb + ß, "Broteinheiten nachmittags", "<, nachmittags", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Broteinheiten abends", 10, 13)
 .Fields![Broteinheiten abends].AllowZeroLength = True
 Call CPropA(QTb + ß, "Broteinheiten abends", "<, abends", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Broteinheiten nachts", 10, 24)
 .Fields![Broteinheiten nachts].AllowZeroLength = True
 Call CPropA(QTb + ß, "Broteinheiten nachts", "<, nachts", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Essenszeit früh", 10, 145)
 .Fields![Essenszeit früh].AllowZeroLength = True
 Call CPropA(QTb + ß, "Essenszeit früh", "^Essenszeiten:früh", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Essenszeit vormittags", 10, 34)
 .Fields![Essenszeit vormittags].AllowZeroLength = True
 Call CPropA(QTb + ß, "Essenszeit vormittags", "<, vormittags", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Essenszeit mittags", 10, 17)
 .Fields![Essenszeit mittags].AllowZeroLength = True
 Call CPropA(QTb + ß, "Essenszeit mittags", "<, mittags", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Essenszeit nachmittags", 10, 24)
 .Fields![Essenszeit nachmittags].AllowZeroLength = True
 Call CPropA(QTb + ß, "Essenszeit nachmittags", "<, nachmittags", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Essenszeit abends", 10, 18)
 .Fields![Essenszeit abends].AllowZeroLength = True
 Call CPropA(QTb + ß, "Essenszeit abends", "<, abends", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Essenszeit spät", 10, 21)
 .Fields![Essenszeit spät].AllowZeroLength = True
 Call CPropA(QTb + ß, "Essenszeit spät", "<, spät", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Spritz-Eß-Abstand früh", 10, 54)
 .Fields![Spritz-Eß-Abstand früh].AllowZeroLength = True
 Call CPropA(QTb + ß, "Spritz-Eß-Abstand früh", "^Spritz-Eß-Abstand:früh", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Spritz-Eß-Abstand mittags", 10, 27)
 .Fields![Spritz-Eß-Abstand mittags].AllowZeroLength = True
 Call CPropA(QTb + ß, "Spritz-Eß-Abstand mittags", "<, mittags", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Spritz-Eß-Abstand abends", 10, 62)
 .Fields![Spritz-Eß-Abstand abends].AllowZeroLength = True
 Call CPropA(QTb + ß, "Spritz-Eß-Abstand abends", "<, abends", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Spritzstelle früh", 10, 8)
 .Fields![Spritzstelle früh].AllowZeroLength = True
 Call CPropA(QTb + ß, "Spritzstelle früh", "^Spritzstellen:früh", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Spritzstelle mittags", 10, 17)
 .Fields![Spritzstelle mittags].AllowZeroLength = True
 Call CPropA(QTb + ß, "Spritzstelle mittags", "<, mittags", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Spritzstelle abends", 10, 37)
 .Fields![Spritzstelle abends].AllowZeroLength = True
 Call CPropA(QTb + ß, "Spritzstelle abends", "<, abends", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Spritzstelle nachts", 10, 94)
 .Fields![Spritzstelle nachts].AllowZeroLength = True
 Call CPropA(QTb + ß, "Spritzstelle nachts", "<, nachts", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Jahr letzte Diabetesschulung", 10, 129)
 .Fields![Jahr letzte Diabetesschulung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Jahr letzte Diabetesschulung", "^Letzte Diabetesschulung:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ort Schulung", 10, 104)
 .Fields![Ort Schulung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ort Schulung", "<in", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("letztes HbA1c", 10, 29)
 .Fields![letztes HbA1c].AllowZeroLength = True
 Call CPropA(QTb + ß, "letztes HbA1c", "^Letztes HbA1c:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("gemessen am", 8, 8)
 Call CPropA(QTb + ß, "gemessen am", "<, gemessen", Null, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("vorherige Werte", 10, 156)
 .Fields![vorherige Werte].AllowZeroLength = True
 Call CPropA(QTb + ß, "vorherige Werte", "<, vorher:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BZMessungen selbst", 10, 41)
 .Fields![BZMessungen selbst].AllowZeroLength = True
 Call CPropA(QTb + ß, "BZMessungen selbst", "^Blutzuckermessung:Selbstmessung?", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Gerät", 10, 63)
 .Fields![Gerät].AllowZeroLength = True
 Call CPropA(QTb + ß, "Gerät", "<:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BZMessungen pW", 10, 112)
 .Fields![BZMessungen pW].AllowZeroLength = True
 Call CPropA(QTb + ß, "BZMessungen pW", "<Zahl d.Messungen pro Woche:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BZMessungen pW ndE", 10, 59)
 .Fields![BZMessungen pW ndE].AllowZeroLength = True
 Call CPropA(QTb + ß, "BZMessungen pW ndE", "<, davon nach dem Essen:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BZMessungen p W nachts", 10, 108)
 .Fields![BZMessungen p W nachts].AllowZeroLength = True
 Call CPropA(QTb + ß, "BZMessungen p W nachts", "<, nachts:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Aufschreiben", 10, 93)
 .Fields![Aufschreiben].AllowZeroLength = True
 Call CPropA(QTb + ß, "Aufschreiben", "<, Dokumentation:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BZWerte v d Essen", 10, 96)
 .Fields![BZWerte v d Essen].AllowZeroLength = True
 Call CPropA(QTb + ß, "BZWerte v d Essen", "^Blutzuckerwerte vor dem Essen:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BZWerte n d Essen", 10, 79)
 .Fields![BZWerte n d Essen].AllowZeroLength = True
 Call CPropA(QTb + ß, "BZWerte n d Essen", "<, nach dem Essen:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("UZ Tageszeit", 10, 104)
 .Fields![UZ Tageszeit].AllowZeroLength = True
 Call CPropA(QTb + ß, "UZ Tageszeit", "^Unterzucker:Bevorzugte Tages-/Uhrzeit", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Unterzucker pM", 10, 84)
 .Fields![Unterzucker pM].AllowZeroLength = True
 Call CPropA(QTb + ß, "Unterzucker pM", "<Zahl der schweren (<50 mg/dl) pro Monat:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("UZ rechtzeitig", 10, 101)
 .Fields![UZ rechtzeitig].AllowZeroLength = True
 Call CPropA(QTb + ß, "UZ rechtzeitig", "<, rechtzeitig bemerkt:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Fremde Hilfe pa", 10, 59)
 .Fields![Fremde Hilfe pa].AllowZeroLength = True
 Call CPropA(QTb + ß, "Fremde Hilfe pa", "<, fremde Hilfe deshalb nötig:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Bewußtlos pa", 10, 70)
 .Fields![Bewußtlos pa].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bewußtlos pa", "<, bewußtlos deshalb:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Keto pa", 10, 66)
 .Fields![Keto pa].AllowZeroLength = True
 Call CPropA(QTb + ß, "Keto pa", "^Zahl der Ketoazidosen pro Jahr:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BZgr300 pM", 10, 122)
 .Fields![BZgr300 pM].AllowZeroLength = True
 Call CPropA(QTb + ß, "BZgr300 pM", ", Zahl der Blutzucker > 300 mg/dl pro Monat:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Bluthochdruck", 10, 65)
 .Fields![Bluthochdruck].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bluthochdruck", "^Bluthochdruck:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BHD seit", 10, 84)
 .Fields![BHD seit].AllowZeroLength = True
 Call CPropA(QTb + ß, "BHD seit", "<seit:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BHD beh mit", 10, 50)
 .Fields![BHD beh mit].AllowZeroLength = True
 Call CPropA(QTb + ß, "BHD beh mit", "<, behandelt mit:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Blutdruckwerte", 10, 186)
 .Fields![Blutdruckwerte].AllowZeroLength = True
 Call CPropA(QTb + ß, "Blutdruckwerte", "^Blutdruckwerte:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BDselbst", 10, 92)
 .Fields![BDselbst].AllowZeroLength = True
 Call CPropA(QTb + ß, "BDselbst", "^Blutdruckselbstmessung:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Schwanger", 10, 43)
 .Fields![Schwanger].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schwanger", "^Aktuelle Schwangerschaft:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Schwanger seit", 10, 8)
 .Fields![Schwanger seit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schwanger seit", "<, seit:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Augensp zuletzt", 10, 107)
 .Fields![Augensp zuletzt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Augensp zuletzt", "^Letzte Augenspiegelung:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Augensp Befund", 10, 122)
 .Fields![Augensp Befund].AllowZeroLength = True
 Call CPropA(QTb + ß, "Augensp Befund", "<, Befund:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Netzhaut gelasert", 10, 142)
 .Fields![Netzhaut gelasert].AllowZeroLength = True
 Call CPropA(QTb + ß, "Netzhaut gelasert", ", Netzhaut schon gelasert:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Sehminderung unbehebbar", 10, 151)
 .Fields![Sehminderung unbehebbar].AllowZeroLength = True
 Call CPropA(QTb + ß, "Sehminderung unbehebbar", ", mit Brille nicht behebbare Sehminderung:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Diabet Nierenschaden", 10, 96)
 .Fields![Diabet Nierenschaden].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diabet Nierenschaden", "^Diabetischer Nierenschaden:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Albumin zuletzt", 10, 73)
 .Fields![Albumin zuletzt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Albumin zuletzt", ", letztes Albumin:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("erhöht?", 10, 52)
 .Fields![erhöht?].AllowZeroLength = True
 Call CPropA(QTb + ß, "erhöht?", "<, Befund:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Dialyse", 1, 1)
 Call PropA(QTb + ß, "Dialyse", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Dialyse", ",:", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Dialyse seit", 10, 31)
 .Fields![Dialyse seit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Dialyse seit", "<seit", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("andere Nierenerkrankung", 10, 50)
 .Fields![andere Nierenerkrankung].AllowZeroLength = True
 Call CPropA(QTb + ß, "andere Nierenerkrankung", ", andere Nierenerkrankung:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Herzkrankheit", 10, 1)
 .Fields![Herzkrankheit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Herzkrankheit", "^Herzkrankheit:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Angina pectoris", 10, 134)
 .Fields![Angina pectoris].AllowZeroLength = True
 Call CPropA(QTb + ß, "Angina pectoris", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Herzinfarkt", 10, 122)
 .Fields![Herzinfarkt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Herzinfarkt", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Herzinfarkt wann", 10, 104)
 .Fields![Herzinfarkt wann].AllowZeroLength = True
 Call CPropA(QTb + ß, "Herzinfarkt wann", "<, wann:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("PTCA oder Stent", 10, 94)
 .Fields![PTCA oder Stent].AllowZeroLength = True
 Call CPropA(QTb + ß, "PTCA oder Stent", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Bypass kardial", 1, 1)
 Call PropA(QTb + ß, "Bypass kardial", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "Bypass kardial", ",:", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Bypass wann", 12, 0)
 .Fields![Bypass wann].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bypass wann", "<, wann:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Herzschwäche", 10, 140)
 .Fields![Herzschwäche].AllowZeroLength = True
 Call CPropA(QTb + ß, "Herzschwäche", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Herzkrankheit Beschreibung", 10, 213)
 .Fields![Herzkrankheit Beschreibung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Herzkrankheit Beschreibung", ", Beschreibung:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Hirndurchblutungsstörung", 10, 65)
 .Fields![Hirndurchblutungsstörung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Hirndurchblutungsstörung", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Schlaganfall", 10, 74)
 .Fields![Schlaganfall].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schlaganfall", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Beindurchblutungsstörung", 10, 136)
 .Fields![Beindurchblutungsstörung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Beindurchblutungsstörung", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Schaufensterkrankheit", 10, 175)
 .Fields![Schaufensterkrankheit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schaufensterkrankheit", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("bypaß peripher", 1, 1)
 Call PropA(QTb + ß, "bypaß peripher", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "bypaß peripher", ",:", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Geschwür", 10, 60)
 .Fields![Geschwür].AllowZeroLength = True
 Call CPropA(QTb + ß, "Geschwür", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Amputation", 10, 98)
 .Fields![Amputation].AllowZeroLength = True
 Call CPropA(QTb + ß, "Amputation", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("pAVK Beschreibung", 10, 97)
 .Fields![pAVK Beschreibung].AllowZeroLength = True
 Call CPropA(QTb + ß, "pAVK Beschreibung", ", Beschreibung der Beinbeschwerden:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ameisenlaufen", 10, 115)
 .Fields![Ameisenlaufen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ameisenlaufen", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ameisen Ausmaß", 10, 93)
 .Fields![Ameisen Ausmaß].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ameisen Ausmaß", "<, Ausmaß:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Druckstellen", 10, 105)
 .Fields![Druckstellen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Druckstellen", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Verformungen", 10, 60)
 .Fields![Verformungen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Verformungen", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Verformungen Beschreibung", 10, 70)
 .Fields![Verformungen Beschreibung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Verformungen Beschreibung", "<Beschreibung:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Fußpflege", 10, 86)
 .Fields![Fußpflege].AllowZeroLength = True
 Call CPropA(QTb + ß, "Fußpflege", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Podologie", 10, 70)
 .Fields![Podologie].AllowZeroLength = True
 Call CPropA(QTb + ß, "Podologie", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Einlagen", 10, 91)
 .Fields![Einlagen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Einlagen", ", diabetesgerechte orthopädische Einlagen/Schuhe:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Neue Fußkomplikationen", 10, 103)
 .Fields![Neue Fußkomplikationen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Neue Fußkomplikationen", "^Neue Fußkomplikationen in den letzten 12 Monaten:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Entleerungsstörungen Magen", 10, 130)
 .Fields![Entleerungsstörungen Magen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Entleerungsstörungen Magen", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Entleerungsstörungen Harnblase", 10, 157)
 .Fields![Entleerungsstörungen Harnblase].AllowZeroLength = True
 Call CPropA(QTb + ß, "Entleerungsstörungen Harnblase", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Schwindel Aufstehen", 10, 121)
 .Fields![Schwindel Aufstehen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schwindel Aufstehen", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Folgeerkrankungen Haut", 10, 156)
 .Fields![Folgeerkrankungen Haut].AllowZeroLength = True
 Call CPropA(QTb + ß, "Folgeerkrankungen Haut", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Bewegungseinschränkungen", 10, 119)
 .Fields![Bewegungseinschränkungen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bewegungseinschränkungen", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Sexualstörung", 10, 69)
 .Fields![Sexualstörung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Sexualstörung", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Sexualstörung seit", 10, 149)
 .Fields![Sexualstörung seit].AllowZeroLength = True
 Call CPropA(QTb + ß, "Sexualstörung seit", "<seit", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Weitere Anamnese", 12, 0)
 .Fields![Weitere Anamnese].AllowZeroLength = True
 Call CPropA(QTb + ß, "Weitere Anamnese", "^:", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Alkohol", 10, 148)
 .Fields![Alkohol].AllowZeroLength = True
 Call CPropA(QTb + ß, "Alkohol", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Tabak", 10, 160)
 .Fields![Tabak].AllowZeroLength = True
 Call CPropA(QTb + ß, "Tabak", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Weitere Medikation", 12, 0)
 .Fields![Weitere Medikation].AllowZeroLength = True
 Call CPropA(QTb + ß, "Weitere Medikation", "^:", 0, Null, 17550, 0, 0, ZielDB)
 .Fields.Append .CreateField("Liphypertrophien Abdomen", 10, 135)
 .Fields![Liphypertrophien Abdomen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Liphypertrophien Abdomen", "^Liphypertrophien:Abdomen", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Liphypertrophien Beine", 10, 49)
 .Fields![Liphypertrophien Beine].AllowZeroLength = True
 Call CPropA(QTb + ß, "Liphypertrophien Beine", "<, Beine:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Liphypertrophien Arme", 10, 4)
 .Fields![Liphypertrophien Arme].AllowZeroLength = True
 Call CPropA(QTb + ß, "Liphypertrophien Arme", "<, Arme:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Beinbefund", 12, 0)
 .Fields![Beinbefund].AllowZeroLength = True
 Call CPropA(QTb + ß, "Beinbefund", "^:", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Hyperkeratosen", 12, 0)
 .Fields![Hyperkeratosen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Hyperkeratosen", ",:", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ulcera", 10, 103)
 .Fields![Ulcera].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ulcera", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Kraft Zehenheber", 10, 51)
 .Fields![Kraft Zehenheber].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kraft Zehenheber", "^Kraft:Zehenheber", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Kraft Zehenbeuger", 10, 50)
 .Fields![Kraft Zehenbeuger].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kraft Zehenbeuger", "<, Zehenbeuger:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Kraft Knie", 10, 59)
 .Fields![Kraft Knie].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kraft Knie", "<, Knie:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("ASR", 10, 50)
 .Fields![ASR].AllowZeroLength = True
 Call CPropA(QTb + ß, "ASR", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("PSR", 10, 67)
 .Fields![PSR].AllowZeroLength = True
 Call CPropA(QTb + ß, "PSR", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Oberflächensensibilität", 10, 92)
 .Fields![Oberflächensensibilität].AllowZeroLength = True
 Call CPropA(QTb + ß, "Oberflächensensibilität", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Monofilamenttest", 10, 70)
 .Fields![Monofilamenttest].AllowZeroLength = True
 Call CPropA(QTb + ß, "Monofilamenttest", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Kalt-Warm", 10, 53)
 .Fields![Kalt-Warm].AllowZeroLength = True
 Call CPropA(QTb + ß, "Kalt-Warm", ", Kalt-Warm-Diskrimination:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Vibration IK", 10, 27)
 .Fields![Vibration IK].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vibration IK", ", Vibrationsempfinden Innenknöchel:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Vibration Großzehe", 10, 21)
 .Fields![Vibration Großzehe].AllowZeroLength = True
 Call CPropA(QTb + ß, "Vibration Großzehe", "<, Großzehe:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Puls Leiste", 10, 33)
 .Fields![Puls Leiste].AllowZeroLength = True
 Call CPropA(QTb + ß, "Puls Leiste", "^Pulse:Leiste", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Puls Kniekehle", 10, 26)
 .Fields![Puls Kniekehle].AllowZeroLength = True
 Call CPropA(QTb + ß, "Puls Kniekehle", "<,Kniekehle:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Puls Atp", 10, 53)
 .Fields![Puls Atp].AllowZeroLength = True
 Call CPropA(QTb + ß, "Puls Atp", "<,Innenknöchel:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Puls Adp", 10, 59)
 .Fields![Puls Adp].AllowZeroLength = True
 Call CPropA(QTb + ß, "Puls Adp", "<,Fußrücken:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("RR", 12, 0)
 .Fields![RR].AllowZeroLength = True
 Call CPropA(QTb + ß, "RR", "^Blutdruck:", 0, Null, 4095, 0, 0, ZielDB)
 .Fields.Append .CreateField("RRTurboMed", 12, 0)
 .Fields![RRTurboMed].AllowZeroLength = True
 Call CPropA(QTb + ß, "RRTurboMed", "", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Herz", 10, 145)
 .Fields![Herz].AllowZeroLength = True
 Call CPropA(QTb + ß, "Herz", "^:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Lunge", 10, 79)
 .Fields![Lunge].AllowZeroLength = True
 Call CPropA(QTb + ß, "Lunge", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Bauch", 10, 106)
 .Fields![Bauch].AllowZeroLength = True
 Call CPropA(QTb + ß, "Bauch", ", Abdomen:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("WS", 10, 56)
 .Fields![WS].AllowZeroLength = True
 Call CPropA(QTb + ß, "WS", ", Wirbelsäule:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("NL", 10, 41)
 .Fields![NL].AllowZeroLength = True
 Call CPropA(QTb + ß, "NL", ", Nierenlager:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("SD", 10, 88)
 .Fields![SD].AllowZeroLength = True
 Call CPropA(QTb + ß, "SD", ", Schilddrüse:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Carotiden", 10, 70)
 .Fields![Carotiden].AllowZeroLength = True
 Call CPropA(QTb + ß, "Carotiden", ", Halsschlagadern:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("NNH", 10, 21)
 .Fields![NNH].AllowZeroLength = True
 Call CPropA(QTb + ß, "NNH", ", Nasennebenhöhlen:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Zähne", 10, 58)
 .Fields![Zähne].AllowZeroLength = True
 Call CPropA(QTb + ß, "Zähne", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Mundhöhle", 10, 64)
 .Fields![Mundhöhle].AllowZeroLength = True
 Call CPropA(QTb + ß, "Mundhöhle", ",:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("LK", 10, 58)
 .Fields![LK].AllowZeroLength = True
 Call CPropA(QTb + ß, "LK", ", Lymphknoten:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("BeinödVen", 12, 0)
 .Fields![BeinödVen].AllowZeroLength = True
 Call CPropA(QTb + ß, "BeinödVen", ", Beinödeme/ Venenkrankheiten:", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Neuro sonst", 10, 74)
 .Fields![Neuro sonst].AllowZeroLength = True
 Call CPropA(QTb + ß, "Neuro sonst", "^Sonstige neurologische Befunde:", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Weitere Befunde", 12, 0)
 .Fields![Weitere Befunde].AllowZeroLength = True
 Call CPropA(QTb + ß, "Weitere Befunde", ", weitere Befunde:", 0, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Schulung", 10, 1)
 .Fields![Schulung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Schulung", "ob Schulungsbedarf", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DMP", 10, 85)
 .Fields![DMP].AllowZeroLength = True
 Call CPropA(QTb + ß, "DMP", "ob Pat. bei HA im DMP", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DMSchulz", 3, 2)
 Call CPropA(QTb + ß, "DMSchulz", "Zahl der DMP-Schulungen hier", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DMSchL", 3, 2)
 Call CPropA(QTb + ß, "DMSchL", "Zahl der abgerechneten DMP-Schulungen hier", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("RRSchulz", 3, 2)
 Call CPropA(QTb + ß, "RRSchulz", "Zahl der Hypertonie-Schulungen hier", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("DMPhier", 8, 8)
 Call CPropA(QTb + ß, "DMPhier", "ob Pat hier im DMP", Null, Null, 1920, 0, 0, ZielDB)
 .Fields.Append .CreateField("HANr", 10, 7)
 .Fields![HANr].AllowZeroLength = True
 Call CPropA(QTb + ß, "HANr", "mit ""/""", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("HANr2", 10, 7)
 .Fields![HANr2].AllowZeroLength = True
 Call CPropA(QTb + ß, "HANr2", "mit ""/""", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("letzte Änderung", 8, 8)
 Call CPropA(QTb + ß, "letzte Änderung", "Datum der letzten Änderung", Null, Null, 1920, 0, 0, ZielDB)
 .Fields.Append .CreateField("Diagnosen", 12, 0)
 .Fields![Diagnosen].AllowZeroLength = True
 Call CPropA(QTb + ß, "Diagnosen", "", 0, Null, 17445, 0, 0, ZielDB)
 .Fields.Append .CreateField("Vorgestellt", 8, 8)
 Call CPropA(QTb + ß, "Vorgestellt", "Erstvorstellung", Null, Null, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Versicherung", 10, 5)
 .Fields![Versicherung].AllowZeroLength = True
 Call CPropA(QTb + ß, "Versicherung", "", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("AktZeit", 8, 8)
 Call CPropA(QTb + ß, "AktZeit", "Aktualisierungszeit", Null, Null, 1920, 0, 0, ZielDB)
 .Fields.Append .CreateField("Ther1", 10, 6)
 .Fields![Ther1].AllowZeroLength = True
 Call CPropA(QTb + ß, "Ther1", "Diät, OAD, CT, ICT, CSII", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("TherAkt", 10, 6)
 .Fields![TherAkt].AllowZeroLength = True
 Call CPropA(QTb + ß, "TherAkt", "Diät, OAD, CT, ICT, CSII", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Prim", 4, 4)
 Call PropA(QTb + ß, "Prim", "DecimalPlaces", 2, 0, ZielDB)
 Call CPropA(QTb + ß, "Prim", "Primärschlüssel", Null, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obAn1eing", 1, 1)
 Call PropA(QTb + ß, "obAn1eing", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obAn1eing", "ob Anamneseblatt S. 1 eingegeben wurde", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obAn2eing", 1, 1)
 Call PropA(QTb + ß, "obAn2eing", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obAn2eing", "ob Anamneseblatt S. 2 eingegeben wurde", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obAnAeing", 1, 1)
 Call PropA(QTb + ß, "obAnAeing", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obAnAeing", "ob Anamneseblatt allgemein eingegeben wurde", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obCheck", 1, 1)
 Call PropA(QTb + ß, "obCheck", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obCheck", "ob Checkliste vorliegt", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obBZausgew", 1, 1)
 Call PropA(QTb + ß, "obBZausgew", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obBZausgew", "ob Blutzuckergerät ausgewechselt", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obOSaufgek", 1, 1)
 Call PropA(QTb + ß, "obOSaufgek", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obOSaufgek", "ob über orthopäd Schuhmacher aufgeklärt", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obPodAufgek", 1, 1)
 Call PropA(QTb + ß, "obPodAufgek", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obPodAufgek", "ob über Podologie aufgeklärt", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obMBlAusgeh", 1, 1)
 Call PropA(QTb + ß, "obMBlAusgeh", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obMBlAusgeh", "ob Merkblatt Fußsyndrom ausgehändigt", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obSchulaufgek", 10, 14)
 .Fields![obSchulaufgek].AllowZeroLength = True
 Call CPropA(QTb + ß, "obSchulaufgek", "ob über Schulung aufgeklärt", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obDMPaufgekl", 10, 17)
 .Fields![obDMPaufgekl].AllowZeroLength = True
 Call PropA(QTb + ß, "obDMPaufgekl", "Caption", dbText, "DMP: j,n,u", ZielDB)
 Call CPropA(QTb + ß, "obDMPaufgekl", "ob über DMP aufgeklärt", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("obMedNetz", 1, 1)
 Call PropA(QTb + ß, "obMedNetz", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "obMedNetz", "ob von Med. Netz geschickt", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("Hausarzt", 10, 98)
 .Fields![Hausarzt].AllowZeroLength = True
 Call CPropA(QTb + ß, "Hausarzt", "Hausarzt laut Anamnesebogen", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("ob", 1, 1)
 Call PropA(QTb + ß, "ob", "Format", dbText, "Yes/No", ZielDB)
 Call CPropA(QTb + ß, "ob", "für verschiedene Aktionen", Null, 106, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("QS", 10, 5)
 .Fields![QS].AllowZeroLength = True
 Call CPropA(QTb + ß, "QS", "Quartal sortiert von vorgestellt", 0, 109, -1, 0, 0, ZielDB)
 .Fields.Append .CreateField("QT", 10, 5)
 .Fields![QT].AllowZeroLength = True
 Call CPropA(QTb + ß, "QT", "Quartal sortiert von vorgestellt", 0, 109, -1, 0, 0, ZielDB)
 End With
 With ZielDB
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Auswahl].Name) = "AUSWAHL" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Auswahl] on [" + QTb + ß + "] ("
  sql = sql + "[Nachname]"
  sql = sql + ",[Vorname]"
  sql = sql + ",[GebDat]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![DMPhier].Name) = "DMPHIER" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [DMPhier] on [" + QTb + ß + "] ("
  sql = sql + "[DMPhier]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![gebdat].Name) = "GEBDAT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [GebDat] on [" + QTb + ß + "] ("
  sql = sql + "[GebDat]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![jlD].Name) = "JLD" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [jlD] on [" + QTb + ß + "] ("
  sql = sql + "[Jahr letzte Diabetesschulung]"
  sql = sql + ",[GebDat]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![lÄnd].Name) = "LÄND" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [lÄnd] on [" + QTb + ß + "] ("
  sql = sql + "[letzte Änderung]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Pat_id].Name) = "PAT_ID" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Pat_ID] on [" + QTb + ß + "] ("
  sql = sql + "[Pat_id]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![PrimaryKey].Name) = "PRIMARYKEY" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE UNIQUE index [PrimaryKey] on [" + QTb + ß + "] ("
  sql = sql + "[Prim]"
  sql = sql + ") WITH PRIMARY"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Ther1].Name) = "THER1" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Ther1] on [" + QTb + ß + "] ("
  sql = sql + "[Ther1]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(QTb + ß).Indexes![Vorgestellt].Name) = "VORGESTELLT" Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  sql = "CREATE index [Vorgestellt] on [" + QTb + ß + "] ("
  sql = sql + "[Vorgestellt]"
  sql = sql + ",[Nachname]"
  sql = sql + ",[Vorname]"
  sql = sql + ",[GebDat]"
  sql = sql + ")"
  Call .Execute(sql)
 End If
 End With
 On Error GoTo fehler
 ' if aktBez then call BezWHerstA(1, 0)
 If Not SammelAufruf Then
 End If ' not Sammelaufruf
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachAAnamnesebogen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MachAAnamnesebogen

Sub CPropA(QT$, F$, üDes$, üUni, üDC, vw, vo, vh, ZDB As DAO.Database)
 On Error GoTo fehler
 If üDes <> "" Then
  Call PropA(QT, F, "Description", 10, üDes, ZDB)
 End If
 If Not IsNull(üUni) Then
  Call PropA(QT, F, "UnicodeCompression", 1, üUni, ZDB)
 End If
 If Not IsNull(üDC) Then
  Call PropA(QT, F, "DisplayControl", 3, üDC, ZDB)
 End If
 Call PropA(QT, F, "ColumnWidth", 3, vw, ZDB)
 Call PropA(QT, F, "ColumnOrder", 3, vo, ZDB)
 Call PropA(QT, F, "ColumnHidden", 1, vh, ZDB)
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in CPropA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' CPropA(QT$, F$, üDes$, üUni, üDC, vw, vo, vh, ZDB as DAO.Database)

Sub PropTabA(Tabl$, PName$, dbTy%, Inh, db As DAO.Database)
 On Error GoTo fehler
 With db.TableDefs(Tabl)
 On Error Resume Next
 Err.Clear
 Properties(PName) = Inh
 If Err.Number <> 0 Then
  On Error GoTo fehler
  .Properties.Append .CreateProperty(PName, dbTy, Inh)
 End If
 End With
 On Error GoTo fehler
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PropTabA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Sub PropTabA(Tabl$, PName$, dbTy%, inh, db As DAO.Database)

Sub PropA(Tabl$, Feld$, PName$, dbTy%, Inh, db As DAO.Database)
 On Error GoTo fehler
 With db.TableDefs(Tabl)
 On Error Resume Next
 Err.Clear
 .Fields(Feld).Properties(PName) = Inh
 If Err.Number <> 0 Then
  On Error GoTo fehler
  .Fields(Feld).Properties.Append .Fields(Feld).CreateProperty(PName, dbTy, Inh)
 End If
 End With
 On Error GoTo fehler
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PropA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Sub PropA(Tabl$, Feld$, PName$, dbTy%, Inh)

Public Function BezRettA()
 Dim ZielDB As DAO.Database, gefunden%
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Rette die Beziehungen ... ")
 Set ZielDB = OpenDatabase(ZielDbS)
 gefunden = True
 On Error Resume Next
 If Not UCase$(ZielDB.TableDefs(tRn).Name) = UCase$(tRn) Then gefunden = 0
 On Error GoTo fehler
 If Not gefunden Then
  DoCmd.TransferDatabase acImport, "Microsoft Access", ZielDbS, acTable, "MSysRelationships", tRn
  DoCmd.TransferDatabase acExport, "Microsoft Access", ZielDbS, acTable, tRn, tRn
  CurrentDb.Execute "drop Table " + tRn
 End If
 ZielDB.Close
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BezRettA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' BezRettA()
#End If ' not ohnedao
Public Sub BezLöschA()
 Dim rel, i%, j% ', ZielDB As DAO.Database
 On Error GoTo fehler
#If Not ohneDAO Then
 Dim ZielDB As DAO.Database
 Set ZielDB = OpenDatabase(ZielDbS)
 For i = 1 To 10
  For Each rel In ZielDB.Relations
   ZielDB.Relations.Delete rel.Name
  Next rel
 Next i
 ZielDB.Close
#Else
 Dim cat As New ADOX.Catalog
 Dim Tbl As ADOX.Table
' Dim fk As New ADOX.Key
 Set DBCn = Nothing
 DBCn.Open aCStr(quelleT, accDtb, "u:\anamnese\quelle.mdb", "")
 Call Zinit(obMySQL:=False)
 cat.ActiveConnection = DBCn ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & ZielDbS & "'"
 For Each Tbl In cat.Tables
  If Tbl.Type = "TABLE" Then
   For j = Tbl.Keys.Count To 1 Step -1
    If Tbl.Keys(j - 1).Type = adKeyForeign Then
     Call Tbl.Keys.Delete(j - 1)
    End If
   Next j
  End If
 Next Tbl
#End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BezLöschA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' BezLöschA()

Public Function BezHerstA()
 ' 1: 1:1-Beziehung (dbRelationUnique)
 ' 2: ref.Integr.nicht erzwungen (DbRelationDontEnforce)
 ' 4: Bez zw 2 verkn Tab aus and DB (dbRelationInherited)
 ' 256: Aktualisierungen weitergegeben (dbRelationUpdateCascade)
 ' 4096: Löschvorgänge weitergegebn (dbRelationDeleteCascade)
 ' 16777216: Left Join standardm angez (dbRelationLeft)
 ' 33554432: Righ Join standardm angez (dbRelationRight)
 
' Dim rel, ZielDB As DAO.Database
' On Error GoTo fehler
' Call syscmd(acSysCmdSetStatus, "Beziehungen herstellen")
' Set ZielDB = OpenDatabase(ZielDbS)
' ZielDB.Close
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BezHerstA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' BezHerstA()
Public Function BezUnters()
 Dim cat As New ADOX.Catalog
 Dim Tbl As New ADOX.Table
 Dim aBez As New ADODB.Recordset, aBezFeld As New ADODB.Recordset, fk As New ADOX.key
 Dim CnStr$
 CnStr = aCStr(quelleT, accDtb, "u:\anamnese\quelle6.mdb", "")
 Call Zinit(obMySQL:=False)
 DBCn.Open CnStr
 Set cat.ActiveConnection = DBCn
 Set Tbl = cat.Tables("LaborXWert")
 For Each fk In Tbl.Keys
  Stop
 Next fk
End Function

Public Function BezWHerstA(art%, Optional obmitFehlerMeld%)
 Dim sql$
 sql = "select distinct(szRelationship) as rel from " & tRn & " where not exists (select * from MSysRelationships where szrelationship = " & tRn & ".szRelationship)"
 Call syscmd(acSysCmdSetStatus, "Beziehungen wiederherstellen")
#If Not DAO Then
 Dim cat As New ADOX.Catalog
 Dim Tbl As New ADOX.Table
 Dim aBez As New ADODB.Recordset, aBezFeld As New ADODB.Recordset, fk As New ADOX.key
 On Error GoTo fehler
 Dim CnStr$
 CnStr = aCStr(quelleT, accDtb, "u:\anamnese\quelle.mdb", "")
 Call Zinit(obMySQL:=False)
 DBCn.Open CnStr
' Call DBCn.Execute("insert into tmprelationen select * from relationen")
 Dim rs As ADODB.Recordset
 Set rs = DBCn.Execute("select * from relationen")
 Do While Not rs.EOF
  Dim rAF&
  Call DBCn.Execute("insert into tmprelationen values(" & rs!ccolumn & "," & rs!grbit & "," & rs!icolumn & ",'" & rs!szcolumn & "','" & Replace(rs!szobject, "ä", "ae") & "','" & rs!szreferencedcolumn & "','" & Replace(rs!szreferencedobject, "ä", "ae") & "','" & Replace(rs!szrelationship, "ä", "ae") & "')", rAF)
  Debug.Print rAF
  rs.Move 1
 Loop
   On Error Resume Next
   aBez.Open sql, DBCn, adOpenStatic, adLockReadOnly
   If Err.Number <> 0 Then
' Berechtigungsfehler, s. http://www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_20263306.html unten,
' muss offenbar nur einmal korrigiert werden
    Dim ds$
    ds = DBCn.Properties("Data Source")
    DBCn.Close
    DBCn.Provider = "Microsoft.Jet.OLEDB.4.0"
    DBCn.Open "data source=" & ds & ";jet oledb:system database=" & Environ("commonprogramfiles") & "\system\system.mdw"
    Const SysTable = "MSysRelationships"
    With cat
    .ActiveConnection = DBCn
    'Set the table to allow Read
    .Users("Admin").SetPermissions SysTable, adPermObjTable, adAccessSet, adRightRead
    End With
    aBez.Open sql, DBCn, adOpenStatic, adLockReadOnly
   End If
   On Error GoTo fehler
   Set cat.ActiveConnection = DBCn
   Do While Not aBez.EOF
     sql = "select * from " & tRn & " where szrelationship = """ + aBez!rel + """ order by icolumn"
'     Set rBezFeld = ZielDB.OpenRecordset(sql)
     Set aBezFeld = Nothing
     aBezFeld.Open sql, DBCn, adOpenStatic, adLockReadOnly
     Dim TName$
     TName = aBezFeld!szobject
     On Error Resume Next
     Err.Clear
     Set Tbl = cat.Tables(TName)
     If Err.Number = 0 Then
      On Error GoTo fehler
      Set fk = Nothing
      fk.Name = aBezFeld!szrelationship
      fk.Type = adKeyForeign
      fk.RelatedTable = aBezFeld!szreferencedobject
' grbit: 1 = 1:1-Beziehung, 2 = ohne ref. Integrität, 256 = Aktualisierungsweitergabe, 4096 = Löschweitergabe,
'        16777216 = alle 1.Tabelle statt beide gleich, 33554433 = alle 2.Tabelle statt beide gleich
     Dim lUpdateRule As ADOX.RuleEnum, lDeleteRule As ADOX.RuleEnum
     
      If (aBezFeld!grbit And 256) <> 0 Then fk.UpdateRule = adRICascade
      If (aBezFeld!grbit And 4096) <> 0 Then fk.DeleteRule = adRICascade
      Do While Not aBezFeld.EOF
      Dim colu$
       colu = aBezFeld!szcolumn
       fk.Columns.Append colu
'      fk.Columns(colu).Attributes = adColFixed
      Dim refcol$
       refcol = aBezFeld!szreferencedcolumn
       fk.Columns(colu).RelatedColumn = refcol
       fk.Columns(colu).ParentCatalog = cat
       aBezFeld.Move 1
      Loop
      On Error Resume Next
      Err.Clear
neu:
      Tbl.Keys.Append fk
      If Err.Number = -2147467259 Then ' kein eindeutiger Index
      Dim idx As ADOX.Index
       Set idx = New ADOX.Index
       idx.PrimaryKey = False
       idx.Name = fk.Columns(0).RelatedColumn
       idx.Unique = True
       idx.Columns.Append fk.Columns(0).RelatedColumn ', cat.Tables(fk.RelatedTable).Columns(fk.Columns(0).RelatedColumn).Type, cat.Tables(fk.RelatedTable).Columns(fk.Columns(0).RelatedColumn).DefinedSize
       Tbl.Indexes.Append idx
       Set idx = New ADOX.Index
       idx.PrimaryKey = False
       idx.Name = fk.Columns(0).Name
       idx.Unique = True
       idx.Columns.Append fk.Columns(0).Name ', cat.Tables(fk.RelatedTable).Columns(fk.Columns(0).Name).Type, cat.Tables(fk.RelatedTable).Columns(fk.Columns(0).Name).DefinedSize
       cat.Tables(fk.RelatedTable).Indexes.Append idx
       Err.Clear
       Tbl.Keys.Append fk
'      If Err.Number <> 0 Then
'      End If
      ElseIf Err.Number <> 0 Then
       Stop
       GoTo neu
      End If
      On Error GoTo fehler
     End If
    aBez.Move 1
   Loop
#Else
   ZielDbS = "u:\anamnese\quelle.mdb"
   Dim rBez As DAO.Recordset, rBezFeld As DAO.Recordset, rel, ZielDB As DAO.Database
   On Error GoTo fehler
   Set ZielDB = OpenDatabase(ZielDbS)
   Select Case art
    Case 1
     Set rBez = ZielDB.OpenRecordset(sql)
     Do While Not rBez.EOF
      sql = "select * from " & tRn & " where szrelationship = """ + rBez!rel + """ order by icolumn"
      Set rBezFeld = ZielDB.OpenRecordset(sql)
      Set rel = ZielDB.CreateRelation(rBezFeld!szrelationship, rBezFeld!szreferencedobject, rBezFeld!szobject, rBezFeld!grbit)
      Do While Not rBezFeld.EOF
       rel.Fields.Append rel.CreateField(rBezFeld!szreferencedcolumn)
       rel.Fields(rBezFeld!szreferencedcolumn).ForeignName = rBezFeld!szcolumn
       rBezFeld.Move 1
      Loop
     Dim k%, l%
      ZielDB.Relations.Append rel
      If Err.Number <> 0 Then
       gesfehler = gesfehler + 1
       If obmitFehlerMeld Then
        MsgBox "Fehler bei Beziehung: " + CStr(rel.Properties(0)) & " " & CStr(rel.Properties(1)) & " " & CStr(rel.Properties(2)) & " " & CStr(rel.Properties(3))
       End If
       For k = 0 To ZielDB.TableDefs.Count - 1
        For l = 0 To ZielDB.TableDefs(k).Indexes.Count - 1
         If ZielDB.TableDefs(k).Indexes(l).Name = rel.Name Then
          If obmitFehlerMeld Then
           MsgBox CStr(k) & " " & CStr(l)
           MsgBox ZielDB.TableDefs(k).Indexes(l).Name & " " & rel.Name
           MsgBox ZielDB.TableDefs(k).Name
           MsgBox "Muß Index löschen!"
          End If
          ZielDB.TableDefs(k).Indexes.Delete rel.Name
          On Error Resume Next
          ZielDB.Relations.Append rel
          If Err.Number <> 0 And obmitFehlerMeld Then
           MsgBox Err.Description
           Stop
          End If
          GoTo wi
         End If
        Next l
       Next k
wi:
      End If
      On Error GoTo fehler
     rBez.Move 1
    Loop
    rBez.Close
    If Not rBezFeld Is Nothing Then rBezFeld.Close
  Case 2
   If InStr(tRn, "tmp") = 1 Then ZielDB.Execute " drop Table [" + tRn + "]; "
   mitBez = 0
 End Select
 ZielDB.Close
#End If
 syscmd acSysCmdClearStatus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BezWHerstA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' BezWHerstA(art%, optional obmitFehlerMeld%)

Public Sub MachSammelTabA()
#If Not ohneDAO Then
 On Error GoTo fehler
 SammelAufruf = True
 Call BezLöschA
 Call MachAAnamnesebogen
 Call MachAAU
 Call MachAAugenbefunde
 Call MachABezirke
 Call MachABriefe
 Call MachADateien
 Call MachADiagnosen
 Call MachADiagnosen_exportiert
 Call MachADiagnosenExport
 Call MachAdmpreihe
 Call MachADMP_Uschr
 Call MachADokumente
 Call MachADokumente_abgehakt
 Call MachAEBM2000plus
 Call MachAEBM2000plus1
 Call MachAEinfuegefehler
 Call MachAEinstellungen
 Call MachAEintraege
 Call MachAEintraege_Arten
 Call MachAEintragsZahlen
 Call MachAFaelle
 Call MachAFeldLaengen
 Call MachAFeldProp
 Call MachAFormInhaltFeld
 Call MachAFormInhaltFeldInh
 Call MachAFormInhaltForm_Abk
 Call MachAFormInhFeld
 Call MachAFormInhKopf
 Call MachAFormulare
 Call MachAFremdlabor
 Call MachAfuerDiagExp
 Call MachAfuerDiagExp_Kopie
 Call MachAfuerDiagExpArchiv
 Call MachAfuerLeistExp
 Call MachAfuerLeistExp_Kopie
 Call MachAfuerLeistExpArchiv
 Call MachAHausaerzte
 Call MachAHAz
 Call MachAICD2006
 Call MachAKassenliste
 Call MachAKHEinweis
 Call MachAKopie_von_Hausaerzte
 Call MachAKVNrUe
 Call MachALaborEintraege
 Call MachALaborgruppen
 Call MachALaborKommentar
 Call MachALaborLangtext
 Call MachALaborNeu
 Call MachALaborParameter
 Call MachALaborXBakt
 Call MachALaborXEingel
 Call MachALaborXLeist
 Call MachALaborXSaetze
 Call MachALaborXUS
 Call MachALaborXWert
 Call MachALauf
 Call MachALbAnforderungen
 Call MachALeistungen
 Call MachALeistungen_exportiert
 Call MachALeistungsexport
 Call MachALeistungsexport__Kopie
 Call MachALex
 Call MachAlistenausgabeuew
 Call MachALkat
 Call MachAMedArten
 Call MachAMedPlan
 Call MachANamen
 Call MachAPauschalen
 Call MachAQueries
 Call MachARel2
 Call MachARelationen
 Call MachARezeptEintraege
 Call MachARR
 Call MachARRParse
 Call MachAStand
 Call MachATabellenfuellung
 Call MachATabGanz
 Call MachATabProp
 Call MachATabTeil
 Call MachAtmpRelationen
 Call MachAUnbekannte_Kennungen
 Call MachAVLassen
 Call MachAAU
 Call MachAAugenbefunde
 Call MachABezirke
 Call MachABriefe
 Call MachADateien
 Call MachADiagnosen
 Call MachADiagnosen_exportiert
 Call MachADiagnosenExport
 Call MachAdmpreihe
 Call MachADMP_Uschr
 Call MachADokumente
 Call MachADokumente_abgehakt
 Call MachAEBM2000plus
 Call MachAEBM2000plus1
 Call MachAEinfuegefehler
 Call MachAEinstellungen
 Call MachAEintraege
 Call MachAEintraege_Arten
 Call MachAEintragsZahlen
 Call MachAFaelle
 Call MachAFeldLaengen
 Call MachAFeldProp
 Call MachAFormInhaltFeld
 Call MachAFormInhaltFeldInh
 Call MachAFormInhaltForm_Abk
 Call MachAFormInhFeld
 Call MachAFormInhKopf
 Call MachAFormulare
 Call MachAFremdlabor
 Call MachAfuerDiagExp
 Call MachAfuerDiagExp_Kopie
 Call MachAfuerDiagExpArchiv
 Call MachAfuerLeistExp
 Call MachAfuerLeistExp_Kopie
 Call MachAfuerLeistExpArchiv
 Call MachAHausaerzte
 Call MachAHAz
 Call MachAICD2006
 Call MachAKassenliste
 Call MachAKHEinweis
 Call MachAKopie_von_Hausaerzte
 Call MachAKVNrUe
 Call MachALaborEintraege
 Call MachALaborgruppen
 Call MachALaborKommentar
 Call MachALaborLangtext
 Call MachALaborNeu
 Call MachALaborParameter
 Call MachALaborXBakt
 Call MachALaborXEingel
 Call MachALaborXLeist
 Call MachALaborXSaetze
 Call MachALaborXUS
 Call MachALaborXWert
 Call MachALauf
 Call MachALbAnforderungen
 Call MachALeistungen
 Call MachALeistungen_exportiert
 Call MachALeistungsexport
 Call MachALeistungsexport__Kopie
 Call MachALex
 Call MachAlistenausgabeuew
 Call MachALkat
 Call MachAMedArten
 Call MachAMedPlan
 Call MachANamen
 Call MachAPauschalen
 Call MachAQueries
 Call MachARel2
 Call MachARelationen
 Call MachARezeptEintraege
 Call MachARR
 Call MachARRParse
 Call MachAStand
 Call MachATabellenfuellung
 Call MachATabGanz
 Call MachATabProp
 Call MachATabTeil
 Call MachAtmpRelationen
 Call MachAUnbekannte_Kennungen
 Call MachAVLassen
 Call BezHerstA
' call BezWHerstA(1, 1)
' if mitBez then if gesfehler = 0 then call BezWHerstA(2)
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachSammelTabA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
#End If
End Sub ' MachSammelTabA
