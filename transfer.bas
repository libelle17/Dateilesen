Attribute VB_Name = "transfer"
#Const ohneDAO = False
#If Not ohneDAO Then
Option Explicit
'Option Compare Database
'Public Const DBn$ = "quelle3"
Public Const opti& = 1 + 2 + 4 + 8 + 131072 ' 131118, 32 ' + 2048 + 16384
' Const CS$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;uid=mysql;pwd=***REMOVED***;OPTION=" & opti
' Public Const CP$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;database=" + DBn + ";uid=praxis;pwd=***REMOVED***;OPTION=" & opti
Const ZwiAnhg$ = "_Zwi"
Const CSStr$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;uid=mysql;pwd=***REMOVED***;OPTION=" & opti
Const tmpRel$ = "tmpRelAbfragen" ' Zwischenspeicher für die Relationen in Access
Const accRel$ = "_AccRel" ' Anhang für Relationenname aus Access

'Function nachMySQL(DBn$, Optional obmitDaten%, Optional anschließendverknüpfen%) ' Transferiert die nicht verknüpften Tabellen aus der aktuellen Accessdatenbank in eine neu zu erstellende MySQL-Datenbank
' Dim sql$, mft$, Comment$, i&
' Dim td As dao.TableDef, fld As dao.Field, f2 As dao.Field, ind As dao.Index, rsq As dao.Recordset
' Dim db As ADODB.Connection
' Dim rs As ADODB.Recordset
' On Error Resume Next
' Close #256
' CurrentDb.TableDefs(tmpRel).Attributes = (CurrentDb.TableDefs(tmpRel).Attributes Or dbSystemObject) - dbSystemObject
' Call CurrentDb.Execute("delete from " + tmpRel)
' Call CurrentDb.Execute("create table " + tmpRel + " (ccolumn long, grbit long, icolumn long, szColumn text(255), szObject text(255), szReferencedColumn text(255), szReferencedObject text(255), szRelationship text(255))")
' On Error GoTo fehler
' Open "u:\nachMySQL.txt" For Output As #256
' Dim rel As Relation
' For Each rel In CurrentDb.Relations
'  For Each fld In rel.Fields
'   Call CurrentDb.Execute("insert into " + tmpRel + " values (" + CStr(rel.Fields.Count) + "," + CStr(rel.Attributes) + "," + CStr(i) + ",'" + fld.ForeignName + "','" + rel.ForeignTable + "','" + fld.Name + "','" + rel.Table + "','" + rel.Name + "')")
'  Next fld
' Next rel
' For Each td In CurrentDb.TableDefs
'  If td.Name = tmpRel Then
'   CurrentDb.TableDefs(tmpRel).Attributes = (CurrentDb.TableDefs(tmpRel).Attributes Or dbSystemObject)
'   Exit For
'  End If
' Next td
' Exit Function
' Set db = New ADODB.Connection
' Set rs = New ADODB.Recordset
' Call db.Open(CSStr)
' Set rs = db.Execute("drop database if exists " + DBn + ";")
' Set rs = db.Execute("create database if not exists " + DBn + " character set latin1 collate latin1_german2_ci;") ' Telefonbuchsortierung Ä=AE, ß = ss; wörterbuchsortiung german1 wäre Ä=A, ß=s
' Set rs = db.Execute("grant all privileges on " + DBn + ".* to praxis with grant option")
' Set rs = db.Execute("use " + DBn + ";")
' Dim aiFName$, obai%
' For Each td In CurrentDb.TableDefs
'  aiFName = ""
'  obai = 0
'  If (td.Attributes And dbAttachedTable) = 0 And (td.Attributes And dbAttachedODBC) = 0 And (td.Attributes And dbSystemObject) = 0 Then
'    Set rs = db.Execute("drop table if exists `" & td.Name & "`")
'    sql = "create table `" + LCase(td.Name) + "` (" + vbCrLf
'    For Each fld In td.Fields
'     sql = sql + " `" + fld.Name + "` " + MySqlTyp(fld.Type, fld.size, fld.Attributes)
'     If Not obai And (fld.Attributes And dbAutoIncrField) Then
'      aiFName = fld.Name
'      obai = True
'     End If
'     On Error Resume Next
'     Comment = fld.Properties!Description
'     If Err.Number = 0 Then sql = sql + " COMMENT '" + Replace(Comment, "'", "''") + "'"
'     On Error GoTo fehler
'     sql = sql + "," + vbCrLf
'    Next fld
'    Dim runde%
'    For Each ind In td.Indexes
'     For runde = 1 To 2
'      If Not ind.Primary Then runde = 2
'      Select Case runde
'       Case 1: sql = sql + " PRIMARY KEY ("
'       Case 2: sql = sql + IIf(ind.Unique, " UNIQUE", "") + " INDEX `" & ind.Name & IIf(ind.Name = "PRIMARY", "KEY", "") & "` (" ' And Not ind.Primary
'      End Select
'      For Each fld In ind.Fields
'       sql = sql + "`" + fld.Name & "`"
'       For Each f2 In td.Fields ' 19.9.08
'        If LCase(f2.Name) = LCase(fld.Name) Then
'         If f2.Type = 12 Then sql = sql & "(255)"
'         Exit For
'        End If
'       Next f2
'       sql = sql & IIf(fld.Attributes And dbDescending, " DESC", "") + ","
'      Next fld
'      If obai = True And runde = 2 And ind.Fields.Count = 1 And ind.Fields(0).Name = aiFName Then obai = False
'      sql = Left(sql, Len(sql) - 1)
'      sql = sql + ")," + vbCrLf
'     Next runde
'    Next ind
'    If obai Then sql = sql & " KEY `" & aiFName & "` (`" & aiFName & "`)," & vbCrLf
'    sql = Left(sql, Len(sql) - 3) + vbCrLf + ") ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;"
'    Print #256, """" + LCase(td.Name); """:" + vbCrLf + sql + vbCrLf + vbCrLf
'    Set rs = db.Execute(sql)
'    Debug.Print td.Name + " fertig!"
'   End If
' Next td
' For Each rel In CurrentDb.Relations
'  On Error Resume Next
'  sql = "alter table " + rel.ForeignTable + " drop foreign key " + rel.Name
'  Set rs = db.Execute(sql)
'  sql = "alter table " + rel.Table + " drop foreign key " + rel.Name
'  Set rs = db.Execute(sql)
'  On Error GoTo fehler
'  sql = "ALTER TABLE `" + LCase(rel.ForeignTable) + "`" + vbCrLf
'  sql = sql + " ADD CONSTRAINT `" + rel.Name + accRel + "` FOREIGN KEY ("
'  For Each fld In rel.Fields
'   sql = sql + "`" + fld.ForeignName + "` ,"
'  Next fld
'  sql = Left(sql, Len(sql) - 2) + ")" + vbCrLf + " REFERENCES `" + LCase(rel.Table) + "`("
'  For Each fld In rel.Fields
'   sql = sql + "`" + fld.Name + "` ,"
'  Next fld
'  sql = Left(sql, Len(sql) - 2) + ")" + vbCrLf
'  If rel.Attributes And dbRelationUpdateCascade Then
'   sql = sql + " ON UPDATE CASCADE" + vbCrLf
''  ElseIf rel.Attributes And dbRelationDontEnforce Then
'' gibts leider nichts passendes
'  End If
'  If rel.Attributes And dbRelationDeleteCascade Then
'   sql = sql + " ON DELETE CASCADE" + vbCrLf
'  End If
'  Set rs = db.Execute(sql)
' Next rel
'' Werte in die SQL-Datenbank einfügen
' ' #if false then
' If obmitDaten Then
'  Set rs = db.Execute("set FOREIGN_KEY_CHECKS = 0")
' For Each td In CurrentDb.TableDefs
'  If (td.Attributes And dbAttachedTable) = 0 And (td.Attributes And dbAttachedODBC) = 0 And (td.Attributes And dbSystemObject) = 0 Then
'    Set rsq = CurrentDb.OpenRecordset(td.Name, dbOpenTable)
'    Do While Not rsq.EOF
'     sql = "insert into `" + LCase(td.Name) + "` values ("
'     For Each fld In rsq.Fields
'      If IsNull(rsq.Fields(fld.Name)) Then
'       sql = sql + "NULL, "
'      ElseIf IsDate(rsq.Fields(fld.Name)) Then
'       sql = sql + """" + Format(rsq.Fields(fld.Name), "YYYY-MM-DD hh:mm:ss") + """, "
'      ElseIf fld.Type = dbBoolean Then
'       sql = sql + IIf(rsq.Fields(fld.Name), "-1", "0") + ", "
'      ElseIf IsNumeric(rsq.Fields(fld.Name)) Then
'       sql = sql + """" + CStr(rsq.Fields(fld.Name)) + """, "
'      Else
'       sql = sql + """" + Replace(Replace(CStr(rsq.Fields(fld.Name)), "\", "\\"), """", "\""") + """, "
'      End If
'     Next fld
'     sql = Left(sql, Len(sql) - 2) + ")"
'     Set rs = db.Execute(sql)
'     DoEvents
'     rsq.Move 1
'    Loop
'    rsq.Close
'   End If
' Next td
' Set rs = db.Execute("set FOREIGN_KEY_CHECKS = 1")
'' #End If
'' Verknüpfte Tabellen erstellen und formatieren
' End If ' obmitdaten
' If anschließendverknüpfen Then Call machVerknTab(DBn)
' Close #256
' MsgBox "Fertig"
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in nachMySQL/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): End
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function ' nachMySQL
Function JetTyp$(Typ%, Optional size&, Optional obauto%)
 Select Case Typ
   Case 1, 11, 17: JetTyp = "BIT" ' dbboolean
   Case 2: JetTyp = "SMALLINT" ' dbbyte
   Case 3:
    If obauto Then ' 16
     JetTyp = "AUTOINCREMENT"
    Else
     JetTyp = "INTEGER"
    End If
   Case 4, 19, 131: JetTyp = "INTEGER" ' dbInteger ' dblong
   Case 5: JetTyp = "double" ' "DECIMAL(15,4)" ' dbcurrency
   Case 6: JetTyp = "float" 'double(10) ' dbSingle
   Case 7: JetTyp = "DATETIME" ' double(20) ' dbDouble
   Case 133, 134, 135: JetTyp = "DATETIME" ' dbDate
'   Case 9 ' dbbinary
   Case 8, 10, 129, 200, 202:
    If size < 256 Then
     JetTyp = "VARCHAR(" + CStr(size) + ")" ' CHAR macht attributes = 1 => dbfixedfield, mit Leerzeichen am Schluß
    Else
     JetTyp = "LONGTEXT"
    End If
'   Case 11: JetTyp = "VARBINARY(255)"
   Case 12, 201, 203: JetTyp = "LONGTEXT"

   Case 15: JetTyp = "BINARY(16)"
'   Case 16 ' dbbigint
'   Case 18 ' dbchar
   Case 20: JetTyp = "DECIMAL(15,4)" ' dbdecimal"
   Case 23, 135: JetTyp = "TIMESTAMP"
   Case 205: JetTyp = "LONGBINARY"
   Case Else: JetTyp = "?"
 End Select
 If Typ = 10 Then If InStr(JetTyp, "NULL") = 0 Then JetTyp = JetTyp + " NULL"

End Function
Function MySqlTyp$(Typ%, Optional size&, Optional obauto%, Optional obQuelleNichtMySQL%)
 Select Case Typ
   Case 1, 11, 17: MySqlTyp = "TINYINT(1) UNSIGNED" 'MySqlTyp = "BOOL" ' dbboolean
   Case 2: MySqlTyp = "SMALLINT" ' dbbyte
   Case 3: MySqlTyp = " INT(10)" ' dbInteger
   Case 4, 131 ' dblong
    If obauto Then ' 16
     MySqlTyp = "INT(10) NOT NULL AUTO_INCREMENT KEY"
    Else
     MySqlTyp = "INT(10)"
    End If
   Case 5:
    If obQuelleNichtMySQL Then
     MySqlTyp = "DOUBLE"
    Else
     MySqlTyp = "DATETIME" ' "DECIMAL(15,4)" ' dbcurrency
    End If
   Case 6: MySqlTyp = "float" 'double(10) ' dbSingle
   Case 7: MySqlTyp = "DOUBLE" ' double(20) ' dbDouble
   Case 133: MySqlTyp = "DATE"
   Case 134: MySqlTyp = "TIME"
   Case 135: MySqlTyp = "DATETIME" ' dbDate
'   Case 9 ' dbbinary
   Case 8, 10, 129, 200, 202: MySqlTyp = "VARCHAR(" + CStr(size) + ")" ' CHAR macht attributes = 1 => dbfixedfield, mit Leerzeichen am Schluß
   Case 11: MySqlTyp = "VARBINARY(255)"
   Case 12, 201, 203: MySqlTyp = "LONGTEXT"

   Case 15: MySqlTyp = "BINARY(16)"
'   Case 16 ' dbbigint
'   Case 18 ' dbchar
   Case 19: MySqlTyp = "INT(2) UNSIGNED"
   Case 20: MySqlTyp = "DECIMAL(15,4)" ' dbdecimal"
   Case 23: MySqlTyp = "TIMESTAMP"
   Case 205: MySqlTyp = "LONGBINARY"
   Case Else: MySqlTyp = "?"
 End Select
' If InStr(MySqlTyp, "NULL") = 0 Then MySqlTyp = MySqlTyp + " NULL"

End Function ' SQLTyp
Function machVerknTab(DBn$) ' Dienstprogramm (erstellt Verknüpfungen und stellt zwischen diesen Beziehungen her)
' Dim Akt$
' Dim fso As New FileSystemObject
' Akt = fso.getfile(CurrentDb.Name).Name
' Akt = LCase(Left(Akt, Len(Akt) - 4))
 Dim db As ADODB.Connection
 Dim rs As ADODB.Recordset
 On Error GoTo fehler
 Set db = New ADODB.Connection
 Set rs = New ADODB.Recordset
 Call db.Open(CSStr)
 Set rs = db.Execute("show tables from " + DBn)
'#If False Then
 Do While Not rs.EOF
  If rs.Fields(0) <> tmpRel Then
   Call domachVerknTab(rs.Fields(0), DBn)
   Call doFormatier(rs.Fields(0), rs.Fields(0) + ZwiAnhg)
   DoEvents
  End If
  rs.Move 1
 Loop
'#End If
 rs.Close
 db.Close
 Call restoreRelations
' Dim td As TableDef
' For Each td In CurrentDb.TableDefs
'   If (td.Attributes And dbAttachedTable) = 0 And (td.Attributes And dbAttachedODBC) = 0 And (td.Attributes And dbSystemObject) = 0 Then
'    Call domachVerknTab(td.Name, dbn$) ' Akt)
'   End If
'  DoEvents
' Next td
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachVerknTab/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' machVerknTab
Function restoreRelations()
 Dim rsDAO1 As dao.Recordset, rsDAO2 As dao.Recordset
 Set rsDAO1 = CurrentDb.OpenRecordset("select distinctrow szrelationship,szObject,szReferencedObject,grbit from " + tmpRel)
 Set rsDAO2 = CurrentDb.OpenRecordset(tmpRel)
 Dim rel As dao.Relation
 If Not rsDAO1.BOF Then
  Do While Not rsDAO1.EOF
   Set rel = CurrentDb.CreateRelation(rsDAO1!szrelationship + "_V", rsDAO1!szreferencedobject, rsDAO1!szobject, rsDAO1!grbit)
   rsDAO2.MoveFirst
   Do While Not rsDAO2.EOF
    If rsDAO2!szrelationship = rsDAO1!szrelationship Then
     rel.Fields.Append rel.CreateField(Trim(rsDAO2!szreferencedcolumn))
     rel.Fields(rel.Fields.Count - 1).ForeignName = Trim(rsDAO2!szcolumn)
    End If
    rsDAO2.Move 1
   Loop
' dbrelationunique        =        1
' DbRelationDontEnforce   =        2
' dbRelationInherited     =        4
' dbRelationUpdateCascade =      256
' dbRelationDeleteCascade =     4096
' dbRelationLeft          = 16777216
' dbRelationRight         = 33554432
   rel.Attributes = rel.Attributes Or dbRelationDontEnforce
   rel.Attributes = (rel.Attributes Or dbRelationUpdateCascade) - dbRelationUpdateCascade
   rel.Attributes = (rel.Attributes Or dbRelationDeleteCascade) - dbRelationDeleteCascade
   CurrentDb.Relations.Append rel
   rsDAO1.Move 1
  Loop
 End If
End Function
Function domachVerknTab(Tabl$, db$) ' Dienstprogramm (erstellt eine Verknüpfung)
 Dim td As dao.TableDef, ZielTab$, tdname$, testDef As dao.TableDef
' On Error Resume Next
 For Each td In CurrentDb.TableDefs
  If td.Name Like Tabl Then
   On Error Resume Next
   ZielTab = td.Name + ZwiAnhg
   Set testDef = CurrentDb.TableDefs(ZielTab)
   If Err.Number <> 0 Then
    On Error GoTo 0
    Call DoCmd.Rename(ZielTab, acTable, td.Name)
   End If
   tdname = td.Name
   GoTo gefunden
  End If
 Next td
 tdname = Tabl
gefunden:
   On Error GoTo 0
   Set td = CurrentDb.CreateTableDef(tdname)
   td.Connect = "ODBC;DRIVER={MySQL ODBC 3.51 Driver};server=linux;database=" + db + ";OPTION=" + CStr(opti) + ";uid=praxis;pwd=***REMOVED***"
' td.Connect = "ODBC;Provider=MSDASQL.1;DRIVER={MySQL ODBC 3.51 Driver};Extended Properties=""DATABASE=" + DB + ";OPTION=131118;PWD=***REMOVED***;PORT=0;SERVER=linux;UID=praxis""" 'geht auch
' td.Connect = "ODBC;DSN=My1"
' td.Name = td.Name + CStr(Len(td.Connect))
   td.SourceTableName = LCase(Tabl)
   On Error Resume Next
   CurrentDb.TableDefs.Append td
' Set td = CurrentDb.CreateTableDef("conn1", dbAttachedODBC, "codes", )
End Function 'domachVerknTab

Private Function doFormatier(Ziel$, Quelle$) ' Dienstprogramm (formatiert verknüpfte Tabellen)
 Dim j&
 On Error GoTo fehler
     Debug.Print Quelle, "->", Ziel
     For j = 0 To CurrentDb.TableDefs(Ziel).Fields.Count - 1
      Debug.Print " " + CurrentDb.TableDefs(Ziel).Fields(j).Name
      On Error Resume Next
      Dim Eig
      Eig = Null
      Eig = CurrentDb.TableDefs(Quelle).Fields(j).Properties("ColumnWidth")
      If Not IsNull(Eig) Then
       CurrentDb.TableDefs(Ziel).Fields(j).Properties.Append CurrentDb.TableDefs(Ziel).Fields(j).CreateProperty("ColumnWidth", 3, Eig)
      End If
      Eig = Null
      Eig = CurrentDb.TableDefs(Quelle).Fields(j).Properties("ColumnOrder")
      If Not IsNull(Eig) Then
       CurrentDb.TableDefs(Ziel).Fields(j).Properties.Append CurrentDb.TableDefs(Ziel).Fields(j).CreateProperty("ColumnOrder", 3, Eig)
      End If
      Eig = Null
      Eig = CurrentDb.TableDefs(Quelle).Fields(j).Properties("ColumnHidden")
      If Not IsNull(Eig) Then
       CurrentDb.TableDefs(Ziel).Fields(j).Properties.Append CurrentDb.TableDefs(Ziel).Fields(j).CreateProperty("ColumnHidden", 1, Eig)
      End If
      Eig = CurrentDb.TableDefs(Quelle).Fields(j).Properties("DisplayControl")
      If Not IsNull(Eig) Then
       CurrentDb.TableDefs(Ziel).Fields(j).Properties("DisplayControl") = Eig
       CurrentDb.TableDefs(Ziel).Fields(j).Properties.Append CurrentDb.TableDefs(Ziel).Fields(j).CreateProperty("DisplayControl", 3, Eig)
      End If
      On Error GoTo fehler
      Debug.Print Quelle
      DoEvents
     Next
  DoEvents
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doFormatier/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' machVerknTabB

#End If
