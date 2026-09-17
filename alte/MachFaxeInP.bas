Attribute VB_Name = "MachFaxeinP"
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.Connection, lErrNr& ' letzter Fehler bei doEx
Const DBn$ = "faxeinp"

Function doEx%(sql$, obtolerant%)
 Dim rAF&, FMeld$
 If obtolerant Then On Error Resume Next Else On Error GoTo fehler
 Call cnz.Execute(sql, rAF)
 lErrNr = Err.Number
 FMeld = "Err.Nr " & lErrNr & ", rAf: " & rAF & " bei " & sql
 On Error GoTo fehler
 Debug.Print FMeld
 Print #302, FMeld
 DoEvents
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doEx/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function ' doEx

Function doGenMachDB()
 On Error GoTo fehler
 Dim zCat As New ADOX.Catalog
 Open "u:\programmierung\dateilesen\MachFaxeInP.bas_prot.txt" For Output As #302
 cnzCStr = "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=Linux;uid=mysql;pwd=97a5o6;"
 cnz.Open cnzCStr
 Call doEx("create database if not exists `" & DBn & "` character set latin1 collate latin1_german2_ci;", 0)
 Call doEx("grant all privileges on `" & DBn & "`.* to 'praxis'@'%' with grant option", 0)
 Call doEx("grant all privileges on `" & DBn & "`.* to 'praxis'@'localhost' with grant option", 0)
 Call doEx("use `" & DBn & "`", 0)
 Call doEx("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", 0)
 cnz.BeginTrans
 Call doEx("set FOREIGN_KEY_CHECKS = 0", 0)
 Call doEx("create table if not exists `dateien` (dummerl char(0)) ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;", 0)
 Call doEx("alter table `dateien`  add `id` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug zu faxe' FIRST, add `pfad` VARCHAR(500) NOT NULL COLLATE latin1_german2_ci COMMENT 'Dateiname samt Pfad' AFTER `id`, add `geändert` DATETIME NOT NULL COMMENT 'Zeile eingetragen' AFTER `pfad`, add `geprüft` TINYINT(1) UNSIGNED NOT NULL COMMENT 'Zeile geprüft' AFTER `geändert`", -1)
 If lErrNr <> 0 Then
 Call doEx("  alter table `dateien`  add `id` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug zu faxe' FIRST", -1)
 Call doEx("  alter table `dateien`  add `pfad` VARCHAR(500) NOT NULL COLLATE latin1_german2_ci COMMENT 'Dateiname samt Pfad' AFTER `id`", -1)
 Call doEx("  alter table `dateien`  add `geändert` DATETIME NOT NULL COMMENT 'Zeile eingetragen' AFTER `pfad`", -1)
 Call doEx("  alter table `dateien`  add `geprüft` TINYINT(1) UNSIGNED NOT NULL COMMENT 'Zeile geprüft' AFTER `geändert`", -1)
 End If ' if lErrNr <> 0 then
 Call doEx("alter table `dateien`  modify `id` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug zu faxe' FIRST, modify `pfad` VARCHAR(500) NOT NULL COLLATE latin1_german2_ci COMMENT 'Dateiname samt Pfad' AFTER `id`, modify `geändert` DATETIME NOT NULL COMMENT 'Zeile eingetragen' AFTER `pfad`, modify `geprüft` TINYINT(1) UNSIGNED NOT NULL COMMENT 'Zeile geprüft' AFTER `geändert`", 0)
 Call doEx("alter table `dateien` drop column dummerl", -1)
 Call doEx("create table if not exists `faxe` (dummerl char(0)) ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;", 0)
 Call doEx("alter table `faxe`  add `Name` VARCHAR(100) NOT NULL COLLATE latin1_german2_ci COMMENT 'Name der Faxdatei' FIRST, add `erstellt` DATETIME NULL COMMENT 'Datum, an dem die Datei zuletzt erstellt wurde' AFTER `Name`, add `geändert` DATETIME NULL COMMENT 'Datum, an dem die Datei zuletzt geändert wurde' AFTER `erstellt`, add `Größe` INT(10) NULL COMMENT 'Größe in Bytes' AFTER `geändert`, add `NameInP` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT 'Neuer Name im Patientenordner' AFTER `Größe`, add `Telefonnummer` VARCHAR(50) NULL COLLATE latin1_german2_ci COMMENT 'Telefonnummer des Senders' AFTER `NameInP`, add `Absender` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT 'Absender, falls bekannt' AFTER `Telefonnummer`, add `kopiert` ? NULL COMMENT 'ob Fax schon kopiert wurde' AFTER `Absender`, add `altNameInP` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT 'Alter Name im Patienten" & _
"ordner' AFTER `kopiert`, add `ID` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT KEY COMMENT '' AFTER `altNameInP`, add `TMStart` DATETIME NULL COMMENT 'TransmissionStart' AFTER `ID`, add `TMEnd` DATETIME NULL COMMENT 'TransmissionEnd' AFTER `TMStart`, add `fFNr` TINYINT(1) UNSIGNED NOT NULL COMMENT 'faxFolder Nr. (1= IncomingQueue, 2 = IncomingArchive' AFTER `TMEnd`", -1)
 If lErrNr <> 0 Then
 Call doEx("  alter table `faxe`  add `Name` VARCHAR(100) NOT NULL COLLATE latin1_german2_ci COMMENT 'Name der Faxdatei' FIRST", -1)
 Call doEx("  alter table `faxe`  add `erstellt` DATETIME NULL COMMENT 'Datum, an dem die Datei zuletzt erstellt wurde' AFTER `Name`", -1)
 Call doEx("  alter table `faxe`  add `geändert` DATETIME NULL COMMENT 'Datum, an dem die Datei zuletzt geändert wurde' AFTER `erstellt`", -1)
 Call doEx("  alter table `faxe`  add `Größe` INT(10) NULL COMMENT 'Größe in Bytes' AFTER `geändert`", -1)
 Call doEx("  alter table `faxe`  add `NameInP` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT 'Neuer Name im Patientenordner' AFTER `Größe`", -1)
 Call doEx("  alter table `faxe`  add `Telefonnummer` VARCHAR(50) NULL COLLATE latin1_german2_ci COMMENT 'Telefonnummer des Senders' AFTER `NameInP`", -1)
 Call doEx("  alter table `faxe`  add `Absender` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT 'Absender, falls bekannt' AFTER `Telefonnummer`", -1)
 Call doEx("  alter table `faxe`  add `kopiert` ? NULL COMMENT 'ob Fax schon kopiert wurde' AFTER `Absender`", -1)
 Call doEx("  alter table `faxe`  add `altNameInP` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT 'Alter Name im Patientenordner' AFTER `kopiert`", -1)
 Call doEx("  alter table `faxe`  add `ID` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT KEY COMMENT '' AFTER `altNameInP`", -1)
 Call doEx("  alter table `faxe`  add `TMStart` DATETIME NULL COMMENT 'TransmissionStart' AFTER `ID`", -1)
 Call doEx("  alter table `faxe`  add `TMEnd` DATETIME NULL COMMENT 'TransmissionEnd' AFTER `TMStart`", -1)
 Call doEx("  alter table `faxe`  add `fFNr` TINYINT(1) UNSIGNED NOT NULL COMMENT 'faxFolder Nr. (1= IncomingQueue, 2 = IncomingArchive' AFTER `TMEnd`", -1)
 End If ' if lErrNr <> 0 then
 Call doEx("alter table `faxe`  modify `Name` VARCHAR(100) NOT NULL COLLATE latin1_german2_ci COMMENT 'Name der Faxdatei' FIRST, modify `erstellt` DATETIME NULL COMMENT 'Datum, an dem die Datei zuletzt erstellt wurde' AFTER `Name`, modify `geändert` DATETIME NULL COMMENT 'Datum, an dem die Datei zuletzt geändert wurde' AFTER `erstellt`, modify `Größe` INT(10) NULL COMMENT 'Größe in Bytes' AFTER `geändert`, modify `NameInP` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT 'Neuer Name im Patientenordner' AFTER `Größe`, modify `Telefonnummer` VARCHAR(50) NULL COLLATE latin1_german2_ci COMMENT 'Telefonnummer des Senders' AFTER `NameInP`, modify `Absender` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT 'Absender, falls bekannt' AFTER `Telefonnummer`, modify `kopiert` ? NULL COMMENT 'ob Fax schon kopiert wurde' AFTER `Absender`, modify `altNameInP` VARCHAR(255) NULL COLLATE latin1_german2_ci COMME" & _
"NT 'Alter Name im Patientenordner' AFTER `kopiert`, modify `ID` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '' AFTER `altNameInP`, modify `TMStart` DATETIME NULL COMMENT 'TransmissionStart' AFTER `ID`, modify `TMEnd` DATETIME NULL COMMENT 'TransmissionEnd' AFTER `TMStart`, modify `fFNr` TINYINT(1) UNSIGNED NOT NULL COMMENT 'faxFolder Nr. (1= IncomingQueue, 2 = IncomingArchive' AFTER `TMEnd`", 0)
 Call doEx("alter table `faxe` drop column dummerl", -1)
 Call doEx("create table if not exists `inca` (dummerl char(0)) ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;", 0)
 Call doEx("alter table `inca`  add `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST, add `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`, add `transe` DATETIME NULL COMMENT '' AFTER `tsid`, add `transs` DATETIME NULL COMMENT '' AFTER `transe`, add `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `transs`, add `fsize` INT(10) NULL COMMENT '' AFTER `id`, add `pages` INT(10) NULL COMMENT '' AFTER `fsize`, add `devname` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `pages`, add `retries` INT(10) NULL COMMENT '' AFTER `devname`, add `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `retries`, add `routi` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `csid`, add `callerid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `routi`", -1)
 If lErrNr <> 0 Then
 Call doEx("  alter table `inca`  add `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST", -1)
 Call doEx("  alter table `inca`  add `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`", -1)
 Call doEx("  alter table `inca`  add `transe` DATETIME NULL COMMENT '' AFTER `tsid`", -1)
 Call doEx("  alter table `inca`  add `transs` DATETIME NULL COMMENT '' AFTER `transe`", -1)
 Call doEx("  alter table `inca`  add `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `transs`", -1)
 Call doEx("  alter table `inca`  add `fsize` INT(10) NULL COMMENT '' AFTER `id`", -1)
 Call doEx("  alter table `inca`  add `pages` INT(10) NULL COMMENT '' AFTER `fsize`", -1)
 Call doEx("  alter table `inca`  add `devname` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `pages`", -1)
 Call doEx("  alter table `inca`  add `retries` INT(10) NULL COMMENT '' AFTER `devname`", -1)
 Call doEx("  alter table `inca`  add `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `retries`", -1)
 Call doEx("  alter table `inca`  add `routi` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `csid`", -1)
 Call doEx("  alter table `inca`  add `callerid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `routi`", -1)
 End If ' if lErrNr <> 0 then
 Call doEx("alter table `inca`  modify `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST, modify `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`, modify `transe` DATETIME NULL COMMENT '' AFTER `tsid`, modify `transs` DATETIME NULL COMMENT '' AFTER `transe`, modify `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `transs`, modify `fsize` INT(10) NULL COMMENT '' AFTER `id`, modify `pages` INT(10) NULL COMMENT '' AFTER `fsize`, modify `devname` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `pages`, modify `retries` INT(10) NULL COMMENT '' AFTER `devname`, modify `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `retries`, modify `routi` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `csid`, modify `callerid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `routi`", 0)
 Call doEx("alter table `inca` drop column dummerl", -1)
 Call doEx("create table if not exists `incq` (dummerl char(0)) ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;", 0)
 Call doEx("alter table `incq`  add `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST, add `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`, add `transe` DATETIME NULL COMMENT '' AFTER `tsid`, add `transs` DATETIME NULL COMMENT '' AFTER `transe`, add `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `transs`, add `fsize` INT(10) NULL COMMENT '' AFTER `id`, add `curp` INT(10) NULL COMMENT '' AFTER `fsize`, add `devid` INT(10) NULL COMMENT '' AFTER `curp`, add `status` INT(10) NULL COMMENT '' AFTER `devid`, add `exts` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `status`, add `extsc` INT(10) NULL COMMENT '' AFTER `exts`, add `jobt` INT(10) NULL COMMENT '' AFTER `extsc`, add `retries` INT(10) NULL COMMENT '' AFTER `jobt`, add `routi` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `retries`, add `csid` VARCHAR(30) NU" & _
"LL COLLATE latin1_german2_ci COMMENT '' AFTER `routi`, add `avop` INT(10) NULL COMMENT '' AFTER `csid`, add `callerid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `avop`", -1)
 If lErrNr <> 0 Then
 Call doEx("  alter table `incq`  add `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST", -1)
 Call doEx("  alter table `incq`  add `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`", -1)
 Call doEx("  alter table `incq`  add `transe` DATETIME NULL COMMENT '' AFTER `tsid`", -1)
 Call doEx("  alter table `incq`  add `transs` DATETIME NULL COMMENT '' AFTER `transe`", -1)
 Call doEx("  alter table `incq`  add `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `transs`", -1)
 Call doEx("  alter table `incq`  add `fsize` INT(10) NULL COMMENT '' AFTER `id`", -1)
 Call doEx("  alter table `incq`  add `curp` INT(10) NULL COMMENT '' AFTER `fsize`", -1)
 Call doEx("  alter table `incq`  add `devid` INT(10) NULL COMMENT '' AFTER `curp`", -1)
 Call doEx("  alter table `incq`  add `status` INT(10) NULL COMMENT '' AFTER `devid`", -1)
 Call doEx("  alter table `incq`  add `exts` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `status`", -1)
 Call doEx("  alter table `incq`  add `extsc` INT(10) NULL COMMENT '' AFTER `exts`", -1)
 Call doEx("  alter table `incq`  add `jobt` INT(10) NULL COMMENT '' AFTER `extsc`", -1)
 Call doEx("  alter table `incq`  add `retries` INT(10) NULL COMMENT '' AFTER `jobt`", -1)
 Call doEx("  alter table `incq`  add `routi` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `retries`", -1)
 Call doEx("  alter table `incq`  add `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `routi`", -1)
 Call doEx("  alter table `incq`  add `avop` INT(10) NULL COMMENT '' AFTER `csid`", -1)
 Call doEx("  alter table `incq`  add `callerid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `avop`", -1)
 End If ' if lErrNr <> 0 then
 Call doEx("alter table `incq`  modify `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST, modify `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`, modify `transe` DATETIME NULL COMMENT '' AFTER `tsid`, modify `transs` DATETIME NULL COMMENT '' AFTER `transe`, modify `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `transs`, modify `fsize` INT(10) NULL COMMENT '' AFTER `id`, modify `curp` INT(10) NULL COMMENT '' AFTER `fsize`, modify `devid` INT(10) NULL COMMENT '' AFTER `curp`, modify `status` INT(10) NULL COMMENT '' AFTER `devid`, modify `exts` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `status`, modify `extsc` INT(10) NULL COMMENT '' AFTER `exts`, modify `jobt` INT(10) NULL COMMENT '' AFTER `extsc`, modify `retries` INT(10) NULL COMMENT '' AFTER `jobt`, modify `routi` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' " & _
"AFTER `retries`, modify `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `routi`, modify `avop` INT(10) NULL COMMENT '' AFTER `csid`, modify `callerid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `avop`", 0)
 Call doEx("alter table `incq` drop column dummerl", -1)
 Call doEx("create table if not exists `outa` (dummerl char(0)) ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;", 0)
 Call doEx("alter table `outa`  add `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST, add `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`, add `submt` DATETIME NULL COMMENT '' AFTER `tsid`, add `submid` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `submt`, add `oscht` DATETIME NULL COMMENT '' AFTER `submid`, add `subject` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `oscht`, add `docname` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `subject`, add `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `docname`, add `fsize` INT(10) NULL COMMENT '' AFTER `id`, add `pages` INT(10) NULL COMMENT '' AFTER `fsize`, add `devname` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `pages`, add `retries` INT(10) NULL COMMENT '' AFTER `devname`, add `prio` INT(10) NULL COMMENT '' AFTER `retries`, " & _
"add `rcfax` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `prio`, add `rcname` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcfax`, add `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcname`, add `sender` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `csid`, add `transs` DATETIME NULL COMMENT '' AFTER `sender`, add `transe` DATETIME NULL COMMENT '' AFTER `transs`", -1)
 If lErrNr <> 0 Then
 Call doEx("  alter table `outa`  add `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST", -1)
 Call doEx("  alter table `outa`  add `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`", -1)
 Call doEx("  alter table `outa`  add `submt` DATETIME NULL COMMENT '' AFTER `tsid`", -1)
 Call doEx("  alter table `outa`  add `submid` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `submt`", -1)
 Call doEx("  alter table `outa`  add `oscht` DATETIME NULL COMMENT '' AFTER `submid`", -1)
 Call doEx("  alter table `outa`  add `subject` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `oscht`", -1)
 Call doEx("  alter table `outa`  add `docname` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `subject`", -1)
 Call doEx("  alter table `outa`  add `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `docname`", -1)
 Call doEx("  alter table `outa`  add `fsize` INT(10) NULL COMMENT '' AFTER `id`", -1)
 Call doEx("  alter table `outa`  add `pages` INT(10) NULL COMMENT '' AFTER `fsize`", -1)
 Call doEx("  alter table `outa`  add `devname` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `pages`", -1)
 Call doEx("  alter table `outa`  add `retries` INT(10) NULL COMMENT '' AFTER `devname`", -1)
 Call doEx("  alter table `outa`  add `prio` INT(10) NULL COMMENT '' AFTER `retries`", -1)
 Call doEx("  alter table `outa`  add `rcfax` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `prio`", -1)
 Call doEx("  alter table `outa`  add `rcname` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcfax`", -1)
 Call doEx("  alter table `outa`  add `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcname`", -1)
 Call doEx("  alter table `outa`  add `sender` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `csid`", -1)
 Call doEx("  alter table `outa`  add `transs` DATETIME NULL COMMENT '' AFTER `sender`", -1)
 Call doEx("  alter table `outa`  add `transe` DATETIME NULL COMMENT '' AFTER `transs`", -1)
 End If ' if lErrNr <> 0 then
 Call doEx("alter table `outa`  modify `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST, modify `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`, modify `submt` DATETIME NULL COMMENT '' AFTER `tsid`, modify `submid` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `submt`, modify `oscht` DATETIME NULL COMMENT '' AFTER `submid`, modify `subject` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `oscht`, modify `docname` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `subject`, modify `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `docname`, modify `fsize` INT(10) NULL COMMENT '' AFTER `id`, modify `pages` INT(10) NULL COMMENT '' AFTER `fsize`, modify `devname` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `pages`, modify `retries` INT(10) NULL COMMENT '' AFTER `devname`, modify `prio` IN" & _
"T(10) NULL COMMENT '' AFTER `retries`, modify `rcfax` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `prio`, modify `rcname` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcfax`, modify `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcname`, modify `sender` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `csid`, modify `transs` DATETIME NULL COMMENT '' AFTER `sender`, modify `transe` DATETIME NULL COMMENT '' AFTER `transs`", 0)
 Call doEx("alter table `outa` drop column dummerl", -1)
 Call doEx("create table if not exists `outq` (dummerl char(0)) ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;", 0)
 Call doEx("alter table `outq`  add `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST, add `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`, add `submt` DATETIME NULL COMMENT '' AFTER `tsid`, add `submid` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `submt`, add `scht` DATETIME NULL COMMENT '' AFTER `submid`, add `oscht` DATETIME NULL COMMENT '' AFTER `scht`, add `subject` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `oscht`, add `docname` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `subject`, add `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `docname`, add `fsize` INT(10) NULL COMMENT '' AFTER `id`, add `pages` INT(10) NULL COMMENT '' AFTER `fsize`, add `curp` INT(10) NULL COMMENT '' AFTER `pages`, add `devid` INT(10) NULL COMMENT '' AFTER `curp`, add `status` INT(10) NULL COMMENT '' AFTE" & _
"R `devid`, add `exts` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `status`, add `extsc` INT(10) NULL COMMENT '' AFTER `exts`, add `retries` INT(10) NULL COMMENT '' AFTER `extsc`, add `prio` INT(10) NULL COMMENT '' AFTER `retries`, add `rct` INT(10) NULL COMMENT '' AFTER `prio`, add `rcfax` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rct`, add `rcname` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcfax`, add `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcname`, add `avop` INT(10) NULL COMMENT '' AFTER `csid`, add `gbr` ? NULL COMMENT '' AFTER `avop`, add `sender` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `gbr`, add `obalt` ? NULL COMMENT '' AFTER `sender`", -1)
 If lErrNr <> 0 Then
 Call doEx("  alter table `outq`  add `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST", -1)
 Call doEx("  alter table `outq`  add `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`", -1)
 Call doEx("  alter table `outq`  add `submt` DATETIME NULL COMMENT '' AFTER `tsid`", -1)
 Call doEx("  alter table `outq`  add `submid` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `submt`", -1)
 Call doEx("  alter table `outq`  add `scht` DATETIME NULL COMMENT '' AFTER `submid`", -1)
 Call doEx("  alter table `outq`  add `oscht` DATETIME NULL COMMENT '' AFTER `scht`", -1)
 Call doEx("  alter table `outq`  add `subject` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `oscht`", -1)
 Call doEx("  alter table `outq`  add `docname` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `subject`", -1)
 Call doEx("  alter table `outq`  add `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `docname`", -1)
 Call doEx("  alter table `outq`  add `fsize` INT(10) NULL COMMENT '' AFTER `id`", -1)
 Call doEx("  alter table `outq`  add `pages` INT(10) NULL COMMENT '' AFTER `fsize`", -1)
 Call doEx("  alter table `outq`  add `curp` INT(10) NULL COMMENT '' AFTER `pages`", -1)
 Call doEx("  alter table `outq`  add `devid` INT(10) NULL COMMENT '' AFTER `curp`", -1)
 Call doEx("  alter table `outq`  add `status` INT(10) NULL COMMENT '' AFTER `devid`", -1)
 Call doEx("  alter table `outq`  add `exts` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `status`", -1)
 Call doEx("  alter table `outq`  add `extsc` INT(10) NULL COMMENT '' AFTER `exts`", -1)
 Call doEx("  alter table `outq`  add `retries` INT(10) NULL COMMENT '' AFTER `extsc`", -1)
 Call doEx("  alter table `outq`  add `prio` INT(10) NULL COMMENT '' AFTER `retries`", -1)
 Call doEx("  alter table `outq`  add `rct` INT(10) NULL COMMENT '' AFTER `prio`", -1)
 Call doEx("  alter table `outq`  add `rcfax` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rct`", -1)
 Call doEx("  alter table `outq`  add `rcname` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcfax`", -1)
 Call doEx("  alter table `outq`  add `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcname`", -1)
 Call doEx("  alter table `outq`  add `avop` INT(10) NULL COMMENT '' AFTER `csid`", -1)
 Call doEx("  alter table `outq`  add `gbr` ? NULL COMMENT '' AFTER `avop`", -1)
 Call doEx("  alter table `outq`  add `sender` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `gbr`", -1)
 Call doEx("  alter table `outq`  add `obalt` ? NULL COMMENT '' AFTER `sender`", -1)
 End If ' if lErrNr <> 0 then
 Call doEx("alter table `outq`  modify `titel` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' FIRST, modify `tsid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `titel`, modify `submt` DATETIME NULL COMMENT '' AFTER `tsid`, modify `submid` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `submt`, modify `scht` DATETIME NULL COMMENT '' AFTER `submid`, modify `oscht` DATETIME NULL COMMENT '' AFTER `scht`, modify `subject` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `oscht`, modify `docname` VARCHAR(255) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `subject`, modify `id` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `docname`, modify `fsize` INT(10) NULL COMMENT '' AFTER `id`, modify `pages` INT(10) NULL COMMENT '' AFTER `fsize`, modify `curp` INT(10) NULL COMMENT '' AFTER `pages`, modify `devid` INT(10) NULL COMMENT '' AFTER `curp`, mo" & _
"dify `status` INT(10) NULL COMMENT '' AFTER `devid`, modify `exts` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `status`, modify `extsc` INT(10) NULL COMMENT '' AFTER `exts`, modify `retries` INT(10) NULL COMMENT '' AFTER `extsc`, modify `prio` INT(10) NULL COMMENT '' AFTER `retries`, modify `rct` INT(10) NULL COMMENT '' AFTER `prio`, modify `rcfax` VARCHAR(40) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rct`, modify `rcname` VARCHAR(150) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcfax`, modify `csid` VARCHAR(30) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `rcname`, modify `avop` INT(10) NULL COMMENT '' AFTER `csid`, modify `gbr` ? NULL COMMENT '' AFTER `avop`, modify `sender` VARCHAR(100) NULL COLLATE latin1_german2_ci COMMENT '' AFTER `gbr`, modify `obalt` ? NULL COMMENT '' AFTER `sender`", 0)
 Call doEx("alter table `outq` drop column dummerl", -1)
 Call doEx("create table if not exists `verze` (dummerl char(0)) ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;", 0)
 Call doEx("alter table `verze`  add `id` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT KEY COMMENT 'Bezug auf Verzeichnis' FIRST, add `verz` VARCHAR(255) NOT NULL COLLATE latin1_german1_ci COMMENT 'Verzeichnis (eindeutig)' AFTER `id`", -1)
 If lErrNr <> 0 Then
 Call doEx("  alter table `verze`  add `id` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT KEY COMMENT 'Bezug auf Verzeichnis' FIRST", -1)
 Call doEx("  alter table `verze`  add `verz` VARCHAR(255) NOT NULL COLLATE latin1_german1_ci COMMENT 'Verzeichnis (eindeutig)' AFTER `id`", -1)
 End If ' if lErrNr <> 0 then
 Call doEx("alter table `verze`  modify `id` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf Verzeichnis' FIRST, modify `verz` VARCHAR(255) NOT NULL COLLATE latin1_german1_ci COMMENT 'Verzeichnis (eindeutig)' AFTER `id`", 0)
 Call doEx("alter table `verze` drop column dummerl", -1)
 Call doEx("create table if not exists `vzo` (dummerl char(0)) ENGINE = INNODB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;", 0)
 Call doEx("alter table `vzo`  add `zid` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT KEY COMMENT 'eindeutige ID dieser Tabelle' FIRST, add `fid` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug auf Faxe' AFTER `zid`, add `vid` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug auf verze' AFTER `fid`, add `geändert` DATETIME NOT NULL COMMENT 'Datum des Eintrags' AFTER `vid`", -1)
 If lErrNr <> 0 Then
 Call doEx("  alter table `vzo`  add `zid` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT KEY COMMENT 'eindeutige ID dieser Tabelle' FIRST", -1)
 Call doEx("  alter table `vzo`  add `fid` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug auf Faxe' AFTER `zid`", -1)
 Call doEx("  alter table `vzo`  add `vid` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug auf verze' AFTER `fid`", -1)
 Call doEx("  alter table `vzo`  add `geändert` DATETIME NOT NULL COMMENT 'Datum des Eintrags' AFTER `vid`", -1)
 End If ' if lErrNr <> 0 then
 Call doEx("alter table `vzo`  modify `zid` INT(2) UNSIGNED NOT NULL AUTO_INCREMENT COMMENT 'eindeutige ID dieser Tabelle' FIRST, modify `fid` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug auf Faxe' AFTER `zid`, modify `vid` INT(2) UNSIGNED NOT NULL COMMENT 'Bezug auf verze' AFTER `fid`, modify `geändert` DATETIME NOT NULL COMMENT 'Datum des Eintrags' AFTER `vid`", 0)
 Call doEx("alter table `vzo` drop column dummerl", -1)
Set zCat = Nothing
Set zCat.ActiveConnection = cnz
 Call doEx("ALTER TABLE `dateien` ADD INDEX `id`(`id`),ADD INDEX `pfad`(`pfad`)", -1)
 Call doEx("ALTER TABLE `faxe` ADD INDEX `kopiertgeändertGröße`(`kopiert`,`geändert`,`Größe`),ADD INDEX `kopiertName`(`kopiert`,`Name`),ADD INDEX `NameinP`(`NameInP`),ADD INDEX `S2`(`erstellt`,`Größe`),ADD INDEX `Telefonnummer`(`Telefonnummer`),ADD INDEX `geändertGröße`(`geändert`,`Größe`),ADD INDEX `Name`(`Name`)", -1)
 Call doEx("ALTER TABLE `inca` ADD INDEX `Id`(`id`),ADD INDEX `transe`(`transe`)", -1)
 Call doEx("ALTER TABLE `incq` ADD INDEX `Id`(`id`),ADD INDEX `transs`(`transs`),ADD INDEX `transe`(`transe`)", -1)
 Call doEx("ALTER TABLE `outa` ADD INDEX `Id`(`id`),ADD INDEX `submt`(`submt`),ADD INDEX `oscht`(`oscht`)", -1)
 Call doEx("ALTER TABLE `outq` ADD INDEX `Id`(`id`),ADD INDEX `submt`(`submt`),ADD INDEX `scht`(`scht`)", -1)
 Call doEx("ALTER TABLE `verze` ADD UNIQUE INDEX `verz`(`verz`)", -1)
 Call doEx("ALTER TABLE `vzo` ADD INDEX `vzo`(`vid`),ADD INDEX `geändert`(`geändert`),ADD INDEX `fid`(`fid`)", -1)
Set zCat = Nothing
Set zCat.ActiveConnection = cnz
Set zCat = Nothing
Set zCat.ActiveConnection = cnz
 Call doEx("drop table `versorgungsamt oberfranken`", -1)
 Call doEx("drop view `versorgungsamt oberfranken`", -1)
 Call doEx("create view `versorgungsamt oberfranken` as select `i`.`titel` AS `titel`,`i`.`tsid` AS `tsid`,`i`.`transe` AS `transe`,`i`.`transs` AS `transs`,`i`.`id` AS `id`,`i`.`fsize` AS `fsize`,`i`.`pages` AS `pages`,`i`.`devname` AS `devname`,`i`.`retries` AS `retries`,`i`.`csid` AS `csid`,`i`.`routi` AS `routi`,`i`.`callerid` AS `callerid` from `inca` `i` where ((`i`.`tsid` like '%803599%') and (`i`.`pages` > 1)) order by `i`.`transe` desc", -1)
 Call doEx("drop table `verz`", -1)
 Call doEx("drop view `verz`", -1)
 Call doEx("create view `verz` as select `vzo`.`zid` AS `id`,`vzo`.`fid` AS `fid`,`verze`.`verz` AS `verz`,`vzo`.`geändert` AS `geändert` from (`vzo` left join `verze` on((`vzo`.`vid` = `verze`.`id`)))", -1)
 Call doEx("set FOREIGN_KEY_CHECKS = 1", 0)
 cnz.CommitTrans
 Close #302
 MsgBox "Fertig mit doGenMachDB!"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doGenmachDB/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doGenmachDB

