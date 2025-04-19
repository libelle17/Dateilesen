#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=DokImport.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
Opt("MustDeclareVars", 1)
#include <GUIConstantsEx.au3>
#include <ColorConstants.au3>
#include <MsgBoxConstants.au3>
#include <StaticConstants.au3> ; für $STM_SETIMAGE
#include <WindowsConstants.au3>
#include <GuiListBox.au3>
#include <File.au3>
#include <word.au3>
#include <Excel.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include "mysql.au3"
#include <WinAPIFiles.au3>
; #include "v:\autoit\excel.au3"
#include <GuiListView.au3> ; für _GUICtrlListView_SetColumnWidth
#include <Timers.au3> ; für _Timer_Init(), _Timer_Diff($hStarttime)
#include <GuiComboBox.au3> ; für _GUICtrlComboBox_ResetContent

Const $RegPath = "HKEY_CURRENT_USER\SOFTWARE\GSProducts\Dokimport"
Const $strich = MultStr(" .", 250)
Const $muster = "*.*"
Const $lMaxZahl = 20 ; Zahl der erinnerten letzten Pfade
Const $mitImport = 0; Archivierung in p:dok\<pid>\ statt Import ins Arztpraxisprogramm
Dim $arrlPf[$lMaxZahl] ; Array der letzten Pfade
Const $lUnZahl = 20 ; Zahl der erinnerten letzten patientenunabhängigen Speicherpfade
Const $clArr = ["Namen aufsteigend", "Namen absteigend", "Datum aufsteigend", "Datum absteigend", "Größe aufsteigend", "Größe absteigend", "Typ aufsteigend", "Typ absteigend"]
Dim $aktSort = 0
Dim $idUnPf[$lUnZahl] ; Buttons für alte Pfade
Dim $idNeuPf ; Button für neuen Pfad
Dim $arrUnPf[0] ; Array der letzten patientenunabhängigen Speicherpfade
Dim $unFrage ; Überschrift für die Frage nach der patientenunabhängigen Speicherung
Dim $ExitFnr ; Fehlernummer beim Programmende
Dim $g_bSortSense, $idDtn, $altdatei, $oPDF, $GUI_PDF, $oExcel, $GUI_Excel, $oWord, $GUI_Word, $aPos, $aDUR, $aML, $Weite, $Hoehe ; aDUR = Datei-Unterrand, aML = Mitellinie
Dim $aFL, $aDL[0], $aUb, $typ[0], $itm[0], $dgroe[0], $dlaend[0]
Dim $Pfad = EnvGet("userprofile") & "\Documents" ; $Pfad="p:\"
Dim $idPic, $idFile, $idFileitem, $idRecentfilesmenu
Dim $idHaupt, $idExit, $excelneu
Dim $idVPat ; Fenster für letzte Verzeichnisse der Speicherung ohne Patientennamen
Dim $iLPf ; Ordnungsziffer des letzten Pfades in der Registry
Dim $iLUnZg ; Ordnungsziffer des letzten patientenunabhängigen Speicherpfades
Dim $idTest, $iOpt, $idOpt, $optaktiv, $PatUnAktiv = 0, $idFertig ; optaktiv: Optionendialog aktiv
Dim $obPDF, $obWd, $obEx, $idCbPDF, $idCbWd, $idCbEx, $idMO, $idTM, $pSystem ; PraxisSystem
Dim $idRei, $Reit ; Reiter für Karteikarte in Turbomed
Dim $idItv, $Itv ; Intervall (ms zwischen zwischen Sendkey-Befehlen)
Dim $idurspNm, $oburspNm = 0, $idurspNVgb, $oburspNVgb = 0 ; ursprünglichen Namen anhängen
Dim $idkPatN, $obkPatN = 0, $idkPatVgb, $obkPatVgb = 0 ; keinen Patientenamen zuteilen
Dim $idntUm, $obntum = 0, $idntUVgb, $obntUVgb = 0 ; nicht umbenennen
Dim $idnimp, $obnImp = 0, $idnimVgb, $obnimVgb = 0 ; nicht importieren
Dim $PatCn = 0 ; Datenbankverbindung zur Patientenauswahl
Dim $ProtCn = 0 ; Datenbankverbindung zur Protokollierung
Dim $connected, $idPatWLb, $idPatWahl, $idPatUn, $idVerarb ; id des Patientenauswahlfeldes
Dim $idVALb, $idVA ; verantwortlich
Dim $irefr ; zum Aktulisieren der Dateiliste
Dim $idlzeig ; zum Fokussieren auf Dateiliste
Const $Vbzl = 5
Dim $iRd[$Vbzl - 1], $iNr[$Vbzl][7], $iSvr[$Vbzl], $iUsr[$Vbzl], $iPwd[$Vbzl], $iDb[$Vbzl], $iPrt[$Vbzl], $iSQL[$Vbzl - 1], $iMA[$Vbzl - 1] ; i = iD, iVA = verantwortlich
Dim $cRd, $cSvr[$Vbzl], $cUsr[$Vbzl], $cPwd[$Vbzl], $cDb[$Vbzl], $cPrt[$Vbzl], $cSQL[$Vbzl], $cMA[$Vbzl] ; c = Contents
Const $BZ = 6
Dim $iILb[$BZ], $iInh[$BZ], $iD1Lb[$BZ], $iD1[$BZ], $iD2Lb[$BZ], $iD2[$BZ]
Dim $urspDat[0], $Kuerzel, $Jahr[0], $datName[0], $gesPfad[0], $patName[0]; für die Benennung: ursprünglicher/späterer Name, Patientenname
Dim $svz[0], $vvz[0] ; Sicherungsverzeichnis, Verschiebeverzeichnis
Dim $spPfad ; Pfad zum Verschieben patientenunabhängiger Dateien
Dim $gwPt ; ausgewählter Patient
; Dim $gwIx; Index der gewählten Datei
; Dim $gwDa; gewählte Datei
Dim $gwVA ; gewählter Mitarbeiter
; Dim $gwNr; Nummer der Datei in der Reihenfolge
; Dim $gwInh; gewählter Inhalt der Zeile 0
; Dim $gwDt1[$BZ]; gewähltes Datum 1
; Dim $gwDt2[$BZ]; gewähltes Datum 2
Dim $iDtAufr ; Datei aufrufen
Dim $iProt ; Protokoll
Dim $iEing, $iDokK, $idnR, $idnL
Dim $BegZ, $Begr[0] ; Zahl der Begriffe, Begriffe
Dim $KopieIn   ; Kopie der Datei in
Dim $iKopIn, $iKopAusw, $KopIn
Dim $idokakt ; zum Fokussieren auf Dokumentenliste
Dim $iSort ; Dateisortierung
Dim $sql
Dim $usr[0], $pwd[0], $passwd
Dim $lgwFl = 0 ;
Dim $markD ; Array markierter Dateien
Dim $ursn, $dtnn ; gesammelte ursprüngliche und zu verarbeitende Dateien
Dim $PData ; Ergebnis des SQL-Befehls zur Patientendatenauswahl

Haupt()

Func Haupt()
	LiesReg()
	$Pfad = RegRead($RegPath, "Pfad")
	If $Pfad = "" Then
		waehlepfad()
	Else ; $Pfad="" Then
		machflstr()
	EndIf ; $Pfad="" Then Else
	zeichneHaupt() ; für $idHaupt

	;	HotKeySet("!l","iDokAct")
	;	HotKeySet("!p","PatAct")
		HotKeySet("!{ESC}","TuComZu")
	;;	HotKeySet("!i","InhAct")
	Schleife()
EndFunc   ;==>Haupt

; in Liesreg, Haupt
Func Schleife()
	Local $aMsg
	While 1
		$aMsg = GUIGetMsg(1) ; Use advanced parameter to get array
		$gwPt = ControlGetText($idHaupt, "", $idPatWahl) ; GUICtrlRead($idPatWahl) bringt nur den Listeneintrag
		$gwVA = GUICtrlRead($idVA)
		Local $gwNr = GUICtrlRead($idDtn) ; Reihenfolgenummer des ggf. gewählten Datei-Eintrags
		Local $gwInh = GUICtrlRead($iInh[0]) ; gewählter Inhalt der Zeile 0
		;		For $i = 0 To $BZ-1
		;			$gwInh[$i]=GUICtrlRead($iInh[$i])
		;			$gwDt1[$i]=GUICtrlRead($iD1[$i])
		;			$gwDt2[$i]=GUICtrlRead($iD2[$i])
		;		Next; $i
		If $obntum = 0 And (($gwPt = "" And $obkPatN = 0) Or $gwNr = "" Or $gwVA = "" Or $gwInh = "" Or GUICtrlRead($iD1[0]) = "") Then
			If BitAND(GUICtrlGetState($idVerarb), $GUI_DISABLE) <> $GUI_DISABLE Then GUICtrlSetState($idVerarb, $GUI_DISABLE)
		Else ; $obntum=0 And (($gwPt="" ...
			If BitAND(GUICtrlGetState($idVerarb), $GUI_ENABLE) <> $GUI_ENABLE Then GUICtrlSetState($idVerarb, $GUI_ENABLE)
		EndIf ; $obntum=0 And (($gwPt="" ... Else
		;		ConsoleWrite("gwPt: " & $gwPt & ", gwNr: " & $gwNr & ", gwVA: " & $gwVA & ", gwInh: " & $gwInh & " obntum: " & $obntum & @LF)
		If $gwPt <> "" Or $gwNr = "" Or $gwVA = "" Or ($gwInh = "" And Not $obntum) Then
			If BitAND(GUICtrlGetState($idPatUn), $GUI_DISABLE) <> $GUI_DISABLE Then GUICtrlSetState($idPatUn, $GUI_DISABLE)
		Else ; $obntum=0 And ($gwPt<>"" Or ...
			If BitAND(GUICtrlGetState($idPatUn), $GUI_ENABLE) <> $GUI_ENABLE Then GUICtrlSetState($idPatUn, $GUI_ENABLE)
		EndIf ; $obntum=0 And ($gwPt<>"" Or ... Else
		If Not IsHWnd($aMsg[1]) Then ContinueLoop ; preventing subsequent lines from processing when nothing happens

		Local $gEntr
		Switch $aMsg[1] ; check which GUI sent the message
			Case $idHaupt
				Switch $aMsg[0] ; Now check for the messages for $g_hGUI2
					Case $GUI_EVENT_MOUSEMOVE, $GUI_EVENT_RESIZED ; -11, -12
					Case Else
						ConsoleWrite("aMsg[1]=" & $aMsg[1] & ", aMsg[0]: " & $aMsg[0] & @LF)
				EndSwitch
				Switch $aMsg[0] ; Now check for the messages for $g_hGUI1
					Case $iDtAufr, $idPatUn, $idVerarb ; Datei gleich umbenennen und importieren
						;		$gwIx = _GUICtrlListView_GetSelectedIndices($idDtn)
						;		$gwDa=GUICtrlRead($gwNr); Datei-Eintrag
						;		$gwIx=StringMid($gwDa,StringInStr($gwDa,"|",0,-2)+1)
						;		$gwIx=StringLeft($gwIx,StringLen($gwIx)-1)
						$gEntr = _GetEntries($idDtn)
				EndSwitch ; $aMsg[0]
				Switch $aMsg[0] ; Now check for the messages for $g_hGUI1
					Case $GUI_EVENT_CLOSE, $idCancel, $idExit ; -3 (auch: esc-Taste), 2, 13
						ComZu("Schleife idHaupt, Msg[0]=" & $aMsg[0])
						ExitLoop
					Case $irefr
						machfilelist(1) ; für $idDtn
						ContinueCase ; danach muss $idlzeig kommen
					Case $idlzeig ; soll nach $irefr kommen
						;						ConsoleWrite("ListAct() über Accelerator" & @LF)
						ListAct() ; für $idDtn
					Case $idPatUn
						$PatUnAktiv = 1
						$markD = $gEntr
						ConsoleWrite("markD: " & $markD & @lf)
						VPatUn() ; rufe patientenunabhängige Verarbeitung auf
					Case $idVerarb ; Datei gleich umbenennen und importieren
						;						ConsoleWrite("!!!!!!! Neufestlegung $markD" & @LF)
						$markD = $gEntr
						ConsoleWrite("markD: " & $markD & @lf)
						Verarbeit()
					Case $idFileitem ; 4
						;						ConsoleWrite("!!!!!!! Neufestlegung $markD" & @LF)
						waehlepfad()
					Case $iOpt
						GUICtrlSetState($idHaupt, $GUI_DISABLE)
						;						GUISetState(@SW_DISABLE, $idHaupt)
						;						GUISetState(@SW_HIDE, $idHaupt)
						$optaktiv = 1
						Optionen()
						;						GUISetState(@SW_ENABLE,$idHaupt)
					Case $iDtAufr ; Datei aufrufen
						;						ConsoleWrite("idDtn: " & GUICtrlRead($idDtn) & ", UBound($aDL): " & UBound($aDL) & ",selind: " & $gwIx & @lf);						ConsoleWrite(_GUICtrlListView_GetSelectedIndices($idDtn) & @LF)
						;						Local $indx=_GUICtrlListView_GetSelectedIndices($idDtn)
						;						ConsoleWrite("indx: " & $indx & @LF)
						;						 MsgBox(0,"Ctrl",GUICtrlRead(GUICtrlRead($idDtn)) & ", indx: " & $indx & ", indxu: " & $indxu & @LF)
						If UBound($gEntr) Then
							For $ix = 0 To UBound($gEntr) - 1
								Local $gesPfd = $Pfad & (StringRight($Pfad, 1) = "\" ? "" : "\") & $aDL[$gEntr[$ix]]
								ConsoleWrite("Führe aus: " & $gesPfd & @LF)
								ShellExecute($gesPfd, "", $Pfad)
							Next ; $ix
						EndIf ; UBound($gEntr) Then
					Case $idDtn ; 15
						;                                MsgBox($MB_SYSTEMMODAL, "listview", "clicked=" & GUICtrlGetState($idDtn), 2)
						;                           if $g_bSortSense = false Then $g_bSortSense = true else $g_bSortSense = false
						_GUICtrlListView_SimpleSort($idDtn, $g_bSortSense, GUICtrlGetState($idDtn), True)
					Case $idurspNm
						; $oburspNm=($oburspNm?0:1); ursprünglichen Namen anhängen
						$oburspNm = (BitAND(GUICtrlRead($idurspNm), $GUI_CHECKED) = $GUI_CHECKED) ; ursprünglichen Namen anhängen
					Case $idntUm
						; $obntum=($obntum?0:1); nicht umbenennen
						$obntum = (BitAND(GUICtrlRead($idntUm), $GUI_CHECKED) = $GUI_CHECKED) ; nicht umbenennen
					Case $idkPatN
						; $obkPatN=($obkPatN?0:1); keinen Patientennamen zuteilen
						$obkPatN = (BitAND(GUICtrlRead($idkPatN), $GUI_CHECKED) = $GUI_CHECKED) ; keinen Patientenamen zuteilen
						ControlSetText($idHaupt, "", $idPatWahl, "")
					Case $idPatWahl
						;						If GUICtrlRead($idPatWahl) Then GUICtrlSetState($idkPatN,$GUI_UNCHECKED)
						If ControlGetText($idHaupt, "", $idPatWahl) Then GUICtrlSetState($idkPatN, $GUI_UNCHECKED)
					Case $idnimp
						; $obnImp=($obnImp?0:1); nicht importieren
						$obnImp = (BitAND(GUICtrlRead($idnimp), $GUI_CHECKED) = $GUI_CHECKED) ; nicht importieren
						;					Case $GUI_EVENT_MOUSEMOVE; -11 ; hier waren anfangs auch noch Com-Objekte drin
					Case 0
						ListAct() ; für $idDtn
					Case Else
						Local $sMenutext = GUICtrlRead($aMsg[0], $GUI_READ_EXTENDED)
						If StringMid($sMenutext, 2, 1) = ":" Then
							;							ConsoleWrite("Vor Aufruf von waehlepfad mit " & $sMenutext & @LF)
							waehlepfad($sMenutext)
						EndIf ; StringMid($sMenutext,2,1)=":" Then
				EndSwitch ; $aMsg[0] ; Now check for the messages for $g_hGUI1
				;				ConsoleWrite("Ende Hauptfensterauswahl" & @LF)
			Case $idVPat
				Switch $aMsg[0] ; Now check for the messages for $g_hGUI2
					Case $GUI_EVENT_CLOSE, $idCancel, $idExit ; -3 (auch: esc-Taste), 2, 13
						ComZu("Schleife idVPat, Msg[0]=" & $aMsg[0])
						;                        ExitLoop
					Case $idNeuPf
						$spPfad = NeuPfad()
						unVerarb($spPfad)
					Case Else
						For $i = 0 To UBound($arrUnPf) - 1
							If $aMsg[0] = $idUnPf[$i] Then
								$spPfad = GUICtrlRead($idUnPf[$i]) ; Speicherpfad
								ConsoleWrite("Treffer: " & $spPfad & @LF)
								unVerarb($spPfad)
								ComZu("Schleife idVPat, Msg[0]=" & $aMsg[0])
								ExitLoop
							EndIf ; $aMsg[0]=$idUnPf[$i] Then
						Next ; $i
				EndSwitch ; $aMsg[0]
			Case $idOpt
				If $optaktiv Then ; Optionendialog aktiv
					Switch $aMsg[0] ; Now check for the messages for $g_hGUI2
						Case $GUI_EVENT_MOUSEMOVE, $GUI_EVENT_RESIZED ; -11, -12
						Case Else
							ConsoleWrite("Hier Auswahl Optionen mit aMsg[0]=" & $aMsg[0] & @LF)
					EndSwitch ; $aMsg[0]
					Switch $aMsg[0] ; Now check for the messages for $g_hGUI2
						Case $GUI_EVENT_CLOSE, $idCancel, $idExit ; -3 (auch: esc-Taste), 2, 13
							If ComZu("Schleife idOpt, Msg[0]=" & $aMsg[0]) Then ExitLoop
						Case $idokakt ; Dummy zum Aktivieren von iDokk
							iDokAct()
							;	HotKeySet("!l","iDokAct")
						Case $idnR
							Local $aktBegr = GUICtrlRead($iEing)
							GUICtrlCreateListViewItem($aktBegr, $iDokK)
							GUICtrlSetData($iEing, "")
							If $BegZ = "" Then $BegZ = 0
							ReDim $Begr[$BegZ + 1]
							$Begr[$BegZ] = $aktBegr
							RegWrite($RegPath, "Begr" & $BegZ, "REG_SZ", $aktBegr)
							$BegZ = $BegZ + 1
							RegWrite($RegPath, "BegZ", "REG_SZ", $BegZ)
						Case $idnL
							GUICtrlSetData($iEing, GUICtrlRead(GUICtrlRead($iDokK)))
							Local $zul = _GUICtrlListView_GetSelectedIndices($iDokK)
							_GUICtrlListView_DeleteItem($iDokK, $zul)
							$BegZ = $BegZ - 1
							RegWrite($RegPath, "BegZ", "REG_SZ", $BegZ)
							For $i = $zul To $BegZ - 1
								$Begr[$i] = $Begr[$i + 1]
								RegWrite($RegPath, "Begr" & $i, "REG_SZ", $Begr[$i])
							Next ; $i
						Case $GUI_EVENT_MOUSEMOVE ; -11
						Case $GUI_EVENT_CLOSE, $idCancel, $idExit, $idOk, $idFertig ; If we get the CLOSE message from this GUI :
							ConsoleWrite("Vor Optionenschluss" & @LF)
							For $i = 0 To $Vbzl - 1
								Local $CtArr[7]
								$cSvr[$i] = GUICtrlRead($iSvr[$i])
								$cUsr[$i] = GUICtrlRead($iUsr[$i])
								$cPwd[$i] = GUICtrlRead($iPwd[$i])
								$cDb[$i] = GUICtrlRead($iDb[$i])
								$cPrt[$i] = GUICtrlRead($iPrt[$i])
								$CtArr[0] = $iSvr[$i]
								$CtArr[1] = $iUsr[$i]
								$CtArr[2] = $iPwd[$i]
								$CtArr[3] = $iDb[$i]
								$CtArr[4] = $iPrt[$i]
								If $i < $Vbzl - 1 Then
									$cMA[$i] = GUICtrlRead($iMA[$i])
									$cSQL[$i] = GUICtrlRead($iSQL[$i])
									If BitAND(GUICtrlRead($iRd[$i]), $GUI_CHECKED) = $GUI_CHECKED Then
										RegWrite($RegPath, "VbdgRadio", "REG_SZ", $i)
									EndIf
									$CtArr[5] = $iSQL[$i]
									$CtArr[6] = $iMA[$i]
								EndIf ; $i<$Vbzl-1 Then
								For $iCtl In $CtArr
									RegWrite($RegPath, OptName($iCtl), "REG_SZ", GUICtrlRead($iCtl))
								Next ; $iCtl
							Next ; $i
							$Reit = GUICtrlRead($idRei)
							RegWrite($RegPath, "Reiter", "REG_SZ", $Reit)
							$Itv = GUICtrlRead($idItv)
							$Itv = 50 ; mal passager
							RegWrite($RegPath, "Intervall", "REG_SZ", $Itv)
							$KopIn = GUICtrlRead($iKopIn)
							RegWrite($RegPath, "KopIn", "REG_SZ", $KopIn)
							$aktSort = _ArraySearch($clArr, GUICtrlRead($iSort))
							RegWrite($RegPath, "Sort", "REG_SZ", $aktSort)
							If ComZu("Schleife idOpt, Msg[0]=" & $aMsg[0]) Then ExitLoop
						Case $idCbPDF
							ConsoleWrite("hier in idCbPdf" & @LF)
							If $obPDF <> "" Then
								$obPDF = ""
							Else
								$obPDF = "1"
								If Not IsObj($oPDF) Then machPdf()
							EndIf ; $obPdf<>""
							RegWrite($RegPath, "obPDF", "REG_SZ", $obPDF)
						Case $idCbWd
							ConsoleWrite("hier in idCbWd" & @LF)
							If $obWd <> "" Then
								$obWd = ""
							Else
								$obWd = "1"
								If Not IsObj($oWord) Then machWord()
							EndIf ; $obWd<>""
							RegWrite($RegPath, "obWord", "REG_SZ", $obWd)
						Case $idCbEx
							If $obEx <> "" Then
								$obEx = ""
							Else
								$obEx = "1"
								If Not IsObj($oExcel) Then machExcel()
							EndIf ; $obEx<>""
							RegWrite($RegPath, "obExcel", "REG_SZ", $obEx)
						Case $idurspNVgb
							$oburspNVgb = ($oburspNVgb ? 0 : 1)
							RegWrite($RegPath, "oburspNVgb", "REG_SZ", $oburspNVgb)
						Case $idkPatVgb
							$obkPatVgb = ($obkPatVgb ? 0 : 1)
							RegWrite($RegPath, "obkPatVgb", "REG_SZ", $obkPatVgb)
						Case $idntUVgb
							$obntUVgb = ($obntUVgb ? 0 : 1)
							RegWrite($RegPath, "obntuVgb", "REG_SZ", $obntUVgb)
						Case $idnimVgb
							$obnimVgb = ($obnimVgb ? 0 : 1)
							RegWrite($RegPath, "obnimVgb", "REG_SZ", $obnimVgb)
						Case $idMO ; Medial Office
							$pSystem = 0
							RegWrite($RegPath, "pSystem", "REG_SZ", $pSystem)
						Case $idTM ; Turbomed
							$pSystem = 1
							RegWrite($RegPath, "pSystem", "REG_SZ", $pSystem)
						Case $iKopAusw
							Local $sFileSelectFolder1 = FileSelectFolder("Bitte Verzeichnis auswählen:", "")
							Local $_Path = _WinAPI_GetFullPathName($sFileSelectFolder1)
							ConsoleWrite("_Path: " & $_Path & @LF)
							If $_Path <> "" Then GUICtrlSetData($iKopIn, $_Path)
						Case Else
							For $i = 0 To $Vbzl - 1
								;							If $aMsg[0] = $iSvr[$i] Then ConsoleWrite('Event from input control, value: ' & GUICtrlRead($iSvr[$i]) & @CRLF)
								If $i < $Vbzl - 1 Then
									If $aMsg[0] = $iRd[$i] Then
										If BitAND(GUICtrlRead($iRd[$i]), $GUI_CHECKED) = $GUI_CHECKED Then
											$cRd = $i
											RegWrite($RegPath, "VbdgRadio", "REG_SZ", $i)
											ContinueLoop
										EndIf ; BitAND(GUICtrlRead($iRd[
									Else ; $aMsg[0] = $iRd[$i] Then
										Switch $aMsg[0] ; WM_COMMAND mit en_change führte z.T. zum Löschen der Inputbox-Inhalte in der Registry
											Case $iSQL[$i]
												RegWrite($RegPath, OptName($iSQL[$i]), "REG_SZ", GUICtrlRead($iSQL[$i]))
												ContinueLoop
										EndSwitch ; $aMsg[0]
									EndIf ; $aMsg[0] = $iRd[$i] Then Else
									Switch $aMsg[0] ; WM_COMMAND mit en_change führte z.T. zum Löschen der Inputbox-Inhalte in der Registry
										Case $iSvr[$i], $iUsr[$i], $iPwd[$i], $iDb[$i], $iPrt[$i]
											RegWrite($RegPath, OptName($aMsg[0]), "REG_SZ", GUICtrlRead($aMsg[0]))
									EndSwitch
								EndIf ; $i<$Vbzl-1 Then
							Next ; $i
							ConsoleWrite("sonstiges Event " & $aMsg[0] & @LF)
					EndSwitch ; $aMsg[0]
				EndIf ; $optaktiv
		EndSwitch ; $aMsg[1] ; check which GUI sent the message
	WEnd ; 1
	ConsoleWrite("Ende Schleife()" & @LF)
	Exit (0)
EndFunc   ;==>Schleife

; aufgerufen in Schleife()
Func NeuPfad() ; s. fuelleUnPfade()
	Local $Pfad = FileSelectFolder($unFrage, "", @DocumentsCommonDir)
	ConsoleWrite($Pfad & @LF)
	If $Pfad Then
		RegWrite($RegPath, "un" & $iLUnZg, "REG_SZ", $Pfad)
		$iLUnZg = $iLUnZg + 1
		If $iLUnZg > $lUnZahl Then $iLUnZg = 1
		RegWrite($RegPath, "iLUnZg", "REG_SZ", $iLUnZg)
	EndIf ; $Pfad
	Return $Pfad
EndFunc   ;==>NeuPfad

; aufgerufen in Verarbeit
Func patNamen($ix, ByRef $gebDat, ByRef $Ort, ByRef $Pid) ; erstellt $urspDat[$ix], $Jahr[$ix], $datName[$ix], $patName[$ix], $gebDat, $Ort, $pid
	$patName[$ix] = ""
	$datName[$ix] = ""
	Local $pos = StringInStr($gwPt, ", *")
	If $pos <> 0 Then
		If $obkPatN = 0 Then $datName[$ix] = StringLeft($gwPt, $pos + 1) & "" & StringMid($gwPt, $pos + 3, 10) & "," ; Name + Geburtsdatum; "geb. "; StringLeft($gwPt,$pos+14)
		$patName[$ix] = StringLeft($gwPt, $pos - 1) ; alles vor ", *"
		$gebDat = StringMid($gwPt, $pos + 9, 4) & StringMid($gwPt, $pos + 6, 2) & StringMid($gwPt, $pos + 3, 2)
		$Ort = StringMid($gwPt, $pos + 15)
		$pos = StringInStr($Ort, ",")
		If $pos > 0 Then $Ort = StringLeft($Ort, $pos - 1)
		$pos = StringInStr($gwPt, "Pat_ID", 0, -1)
		If $pos <> 0 Then
			If $obkPatN = 0 Then $datName[$ix] = $datName[$ix] & StringMid($datName[$ix], $pos) ; Name + Geburtstdatum + " Pat_ID " + pat_id
			$Pid = StringMid($gwPt, $pos + 7)
			;								ConsoleWrite("pos: " & $pos & ", gwPt: " & $gwPt & @LF)
			;								ConsoleWrite("pid: " & $pid & @LF)
		EndIf ; $pos<>0
	Else ; $pos<>0 Then
		If $obkPatN = 0 Then $patName[$ix] = $gwPt
		$datName[$ix] = $gwPt
	EndIf ; $pos<>0
	If $datName[$ix] Then $datName[$ix] = $datName[$ix] & " "
	;						 Local $patName[$ix]=StringLeft($gwPt,StringInStr($gwPt,", *")) & " "
	DNamen2($ix) ; erstellt $urspDat[$ix], $Jahr[$ix] und den hinteren Teil von $datName[$ix]
	;						ConsoleWrite($gwPt & " Name: " & $patName[$ix] & ", PatID: " & $Pid & ", Gebdat: " & $gebDat & ", Ort: " & $Ort & ";" & @LF)
	;						ConsoleWrite("aDL: " & $aDL[$gwIx] & ", dgroe: " & $dgroe[$gwIx] & ", dlaend: " & $dlaend[$gwIx] & @LF)
	;						ConsoleWrite("NowTime:" & _NowDate() & ", NowTime:" & _NowTime() & @LF)
	ConsoleWrite("datName: " & $datName[$ix] & @LF)
	If $Pid = 0 Then
		Local $pos = StringInStr($datName[$ix], "PID")
		If $pos Then
			Local $p2 = StringInStr($datName[$ix], " ", 0, 1, $pos + 4)
			If $p2 Then
				$Pid = StringMid($datName[$ix], $pos + 4, $p2 - $pos - 5)
				;				ConsoleWrite("pid: " & $pid & @LF)
				If Not IsNumber($Pid) Then $Pid = 0
			EndIf ; $p2
		EndIf ; $pos
	EndIf ; $pid=0 Then
	Local $obvb = 1 ;
	If $Pid = 0 Then
		$obvb = MsgBox($MB_OKCANCEL + $MB_DEFBUTTON2, "Verarbeitung", $datName[$ix] & " enthält keine Pat_ID. Wurde er richtig ausgewählt?")
		If $obvb = 2 Then GUICtrlSetState($idPatWahl, $GUI_FOCUS)
	EndIf ; $pid=0 Then
	If $obvb = 1 And $gebDat = '' Then
		$obvb = MsgBox($MB_OKCANCEL + $MB_DEFBUTTON2, "Verarbeitung", $datName[$ix] & " enthält kein Geburtsdatum. Wurde er richtig ausgewählt?")
		If $obvb = 2 Then
			GUICtrlSetState($idPatWahl, $GUI_FOCUS)
		Else ; $obvb=2 Then
			$gebDat = "30.12.1899" ; in MySQL 0-Datum
		EndIf ; $obvb=2 Then Else
	EndIf ; $obvb=1 And $gebDat='' Then
	Return ($obvb = 1)
EndFunc   ;==>patNamen

; aufgerufen in Verarbeit und unVerarb
Func bestaetig($ohneBei = 0)
	$ursn = ""
	$dtnn = ""
	For $ix = 0 To UBound($markD) - 1
		$ursn = $ursn & " '" & $urspDat[$ix] & "'" & @LF
		$dtnn = $dtnn & " '" & $datName[$ix] & "'" & @LF
	Next ; $ix=0 to UBound($markD)-1
	Local $obweiter = MsgBox($MB_OKCANCEL, "Verarbeitung", "Soll" & (UBound($markD) > 1 ? "en" : "") & " die Datei" & (UBound($markD) > 1 ? "en" : "") & @LF & $ursn & ($ohneBei ? "in" & @LF & " '" & $spPfad & "'" : "bei" & @LF & " " & $gwPt) & @LF & "als" & @LF & $dtnn & ($ohneBei ? "abgespeichert" : "eingetragen") & " werden?")
	If $obweiter = 2 Then $obweiter = 0
	Return $obweiter
EndFunc   ;==>bestaetig

; aufgerufen in Verarbeit und unVerarb
Func tuBenenn($ix) ; benennt $urspDat[$ix] in $datName[$ix] bzw. prüft bei Gleichnamigkeit die Bearbeitbarkeit
	;							ConsoleWrite("idntum: " & $idntUm & @LF)
	Dim $obweiter
	$gesPfad[$ix] = $Pfad & (StringRight($Pfad, 1) = "\" ? "" : "\") & $urspDat[$ix]
	If $obntum Or $urspDat[$ix] = $datName[$ix] Then
		if false then
		 Local $benanntd = $urspDat[$ix] & "_" & Random(0, 1000000000, 1)
 		 $obweiter = Umbenenn($urspDat[$ix], $benanntd, $Pfad)
		 If $obweiter Then $obweiter = Umbenenn($benanntd, $urspDat[$ix], $Pfad)
		else
		 $obweiter=true
		endif
		If $obweiter Then
			GUICtrlSetData($iProt, "1) Bearbeitbarkeit bestätigt: '" & $gesPfad[$ix] & "'" & @CRLF, 1)
		Else ; $obweiter Then
			MsgBox($MB_ICONERROR, "Fehler beim Bewegen", $gesPfad[$ix] & @LF & " konnte nicht bewegt werden. Breche hier ab.")
		EndIf ; $obweiter Then Else
		$datName[$ix] = $urspDat[$ix]
	Else ; $obntum
		;								ConsoleWrite("Ziel: " & $benannt & @LF)
		;								$obweiter = FileMove('"' & $gesPfad[$ix] & '"', '"' & $benannt & '"')
		$obweiter = Umbenenn($urspDat[$ix], $datName[$ix], $Pfad)
		If $obweiter Then
			GUICtrlSetData($iProt, "1) umbenannt: '" & $gesPfad[$ix] & "' -> '" & $datName[$ix] & "'" & @CRLF, 1)
			$gesPfad[$ix] = $Pfad & (StringRight($Pfad, 1) = "\" ? "" : "\") & $datName[$ix]
		Else ; $obweiter Then
			MsgBox($MB_ICONERROR, "Fehler beim Umbenennen", $gesPfad[$ix] & @LF & " konnte nicht in" & @LF & $datName[$ix] & @LF & " umbenannt werden. Breche hier ab.")
		EndIf ; $obweiter Then Else
		;								MsgBox($MB_ICONERROR,"Erfolg beim Umbenennen",$gesPfad[$ix] & @LF & " konnte in" & @LF & $benannt & @LF & " umbenannt werden.")
	EndIf ; $obntum Else
	Return $obweiter
EndFunc   ;==>tuBenenn

; aufgerufen in patNamen und VPatUn
Func DNamen2($ix, $nurersteZeile = 0) ; erstellt $urspDat[$ix], $Jahr[$ix] und den hinteren Teil von $datName[$ix]
	$urspDat[$ix] = $aDL[$markD[$ix]]
	$Jahr[$ix] = StringLeft($dlaend[$markD[$ix]], 4)
	For $i = 0 To ($nurersteZeile ? 0 : $BZ - 1)
		Local $aktInh = GUICtrlRead($iInh[$i])
		If $aktInh <> "" Then
			$datName[$ix] = $datName[$ix] & $aktInh
			Local $aktD1 = GUICtrlRead($iD1[$i])
			If $aktD1 <> "" Then
				$datName[$ix] = $datName[$ix] & " " & $aktD1
				Local $aktD2 = GUICtrlRead($iD2[$i]), $njahr, $orig
				If $aktD2 <> "" Then
					$datName[$ix] = $datName[$ix] & "-" & $aktD2
				EndIf ; $aktD2<>"" Then
				For $k = 1 To 2
					If $k = 1 Then
						$orig = $aktD2
					Else ; $k = 1
						$orig = $aktD1
					EndIf ; $k = 1
					ConsoleWrite("k: " & $k & ", orig: " & $orig & @LF)
					If $orig <> "" Then
						$njahr = StringRegExpReplace(StringRegExpReplace($orig, "\A(\d*)[:.,](\d*)[:.,](\d{2,4}).*", "$3"), "\A(\d{1,2}[./])?(\d{2})\z", "20$2")
						ConsoleWrite("njahr: " & $njahr & @LF)
						If $njahr > 1900 And $njahr < 2100 Then
							$Jahr[$ix] = $njahr
							ExitLoop
						EndIf ; $njahr>1900 And $njahr<2100 Then
					EndIf ; $orig<>"" Then
				Next ; $k
			EndIf ; $aktD1<>""
			$datName[$ix] = $datName[$ix] & ","
		EndIf ; $akz<>""
	Next ; $i

	Local $pos = StringInStr($gwVA, ",")
	If $pos <> 0 Then
		$Kuerzel = StringLeft($gwVA, $pos - 1)
	Else ; $pos<>0 Then
		$Kuerzel = $gwVA
	EndIf ; $pos<>0 Then Else

	If $Kuerzel Then $datName[$ix] = $datName[$ix] & " " & $Kuerzel
	If $oburspNm <> 0 Then
		$datName[$ix] = $datName[$ix] & "; " & $urspDat[$ix]
	Else ; $oburspNm<>0 Then
		$datName[$ix] = $datName[$ix] & "." & $typ[$markD[$ix]] ; (StringInStr($urspDat[$ix],".",-1) ? StringRight($urspDat[$ix],StringLen($urspDat[$ix])-StringInStr($urspDat[$ix],".",0,-1)) : "")
	EndIf ; $oburspNm<>0 Then Else
	If $obntum Then
		$datName[$ix] = $urspDat[$ix]
	EndIf ; $obntum<>0 Then Else
EndFunc   ;==>DNamen2


; aufgerufen in Schleife bei $idVPat
Func unVerarb($spPfad) ; (patienten)unabhängige Verarbeitung
	Local $obweiter, $gebDat = "18991230", $Ort = "", $Pid = 0
	While UBound($markD) ; Pseudo-Schleife
		For $ix = 0 To UBound($markD) - 1
			; 1) benennen
			$urspDat[$ix] = $aDL[$markD[$ix]]
		Next ; $ix
		If Not bestaetig(1) Then ExitLoop
		For $ix = 0 To UBound($markD) - 1
			$obweiter = tuBenenn($ix)
			; 2) in die Datenbank eintragen
			If $obweiter Then
				SicherVZ($ix)
				VerschbVZ($ix, $spPfad)
				$obweiter = DBEintrag($ix, $gebDat, $Ort, $Pid, 2)
			EndIf ; $obweiter Then
			; 3) Sicherheitskopie erstellen
			If $obweiter Then ; der Protkolleintrag entweder nicht vorgesehen war oder gelang
				$obweiter = SicherheitsKopie($ix)
			EndIf; $obweiter
			; 4) verschieben
			If $obweiter Then ;
				$obweiter = Verschieb($ix)
			EndIf ; $obweiter
		Next; $ix
		; 5) flags für nächste Datei zurücksetzen
		If $obweiter Then
			GUICtrlSetState($idurspNm, $GUI_UNCHECKED)
			$oburspNm = 0
			GUICtrlSetState($idkPatN, $GUI_UNCHECKED)
			$obkPatN = 0
			GUICtrlSetState($idntUm, $GUI_UNCHECKED)
			$obntum = 0
			GUICtrlSetState($idnimp, $GUI_UNCHECKED)
			$obnImp = 0
			MsgBox($MB_OK, "Verarbeitung", "Fertig mit Abspeichern von: " & @LF & $ursn & "in" & @LF & " '" & $spPfad & "'" & @LF & "als" & @LF & $dtnn)
		EndIf    ; $obweiter
		ExitLoop
	WEnd ; UBound($markD)
EndFunc   ;==>unVerarb

; aufgerufen in Verarbeit, unVerarb
Func dimensionier($zahl)
	ReDim $urspDat[$zahl]
	ReDim $datName[$zahl]
	ReDim $patName[$zahl]
	ReDim $gesPfad[$zahl]
	ReDim $Jahr[$zahl]
	ReDim $svz[$zahl]
	ReDim $vvz[$zahl]
EndFunc   ;==>dimensionier

; aufgerufen in Schleife()
Func Verarbeit() ; Ok-Knopf => verarbeitet die markierten Dateien patientenabhängig
	Local $obweiter, $gebDat = "", $Ort = "", $Pid = 0
	$oburspNm = (BitAND(GUICtrlRead($idurspNm), $GUI_CHECKED) = $GUI_CHECKED) ; ursprünglichen Namen anhängen
	$obkPatN = (BitAND(GUICtrlRead($idkPatN), $GUI_CHECKED) = $GUI_CHECKED) ; keinen Patientenamen zuteilen
	$obntum = (BitAND(GUICtrlRead($idntUm), $GUI_CHECKED) = $GUI_CHECKED) ; nicht umbenennen
	$obnImp = (BitAND(GUICtrlRead($idnimp), $GUI_CHECKED) = $GUI_CHECKED) ; nicht importieren
	ConsoleWrite("oburspNm: " & $oburspNm & ", obkPatN: " & $obkPatN & ", ontum: " & $obntum & ", obnImp: " & $obnImp & @LF)
	While UBound($markD) ; Pseudo-Schleife
		dimensionier(UBound($markD))
		For $ix = 0 To UBound($markD) - 1
			; 1) benennen; Dateien werden entweder nicht umbenannt oder sind schon im Rumpf gleichnamig
;			ConsoleWrite("UBound($markD): " & UBound($markD) & ", ix: " & $ix & ", markD[$ix]: " & $markD[$ix] & @LF)
			$obweiter = patNamen($ix, $gebDat, $Ort, $Pid) ; erstellt $urspDat[$ix], $Jahr[$ix], $datName[$ix], $patName[$ix], $gebDat, $Ort, $pid
			If Not $obweiter Then ExitLoop
		Next ; $ix
		If Not $obweiter Or Not bestaetig() Then ExitLoop
		For $ix = 0 To UBound($markD) - 1
			If $ix Then Sleep(10 * $Itv)
			If $obweiter Then $obweiter = tuBenenn($ix) ; benennt $urspDat[$ix] in $datName[$ix] bzw. prüft bei Gleichnamigkeit die Bearbeitbarkeit
			; 2) in die Datenbank eintragen
			If $obweiter Then
				SicherVZ($ix)
				$obweiter = DBEintrag($ix, $gebDat, $Ort, $Pid, $KopIn ? 1 : 0)
			EndIf ; $obweiter Then
			; 3) Sicherheitskopie erstellen
			If $obweiter Then ; der Protkolleintrag entweder nicht vorgesehen war oder gelang
				$obweiter = SicherheitsKopie($ix)
			EndIf ; $obweiter
			; 4) ins Arztpraxisprogramm importieren
			If $obweiter And Not $obnImp Then ; der Protkolleintrag entweder nicht vorgesehen war oder gelang und die Sicherheitskopie gelang und nicht kein Import
				If $mitImport Then
				 $obweiter = insArztPraxisProgramm($ix, $Pid)
				Else
				 $obweiter = Aufraeum($ix,$Pid)
				EndIf
			EndIf ; $obweiter
		Next ; $ix
		machfilelist(1) ; für $idDtn
;		ListAct()
		; 5) flags für nächste Datei zurücksetzen
		If $obweiter Then
			GUICtrlSetState($idurspNm, $GUI_UNCHECKED)
			$oburspNm = 0
			GUICtrlSetState($idkPatN, $GUI_UNCHECKED)
			$obkPatN = 0
			GUICtrlSetState($idntUm, $GUI_UNCHECKED)
			$obntum = 0
			GUICtrlSetState($idnimp, $GUI_UNCHECKED)
			$obnImp = 0
			MsgBox($MB_OK, "Verarbeitung", "Fertig mit Eintragen von: " & @LF & $ursn & "bei" & @LF & " '" & $gwPt & "'" & @LF & "als" & @LF & $dtnn)
		EndIf ; $obweiter
		ExitLoop
	WEnd ; UBound($markD)
EndFunc   ;==>Verarbeit

; aufgerufen in unVerarb
Func VerschbVZ($ix, $spPfad)
	If FileExists($gesPfad[$ix]) Then
		$vvz[$ix] = $spPfad
	EndIf ; FileExists($gesPfad[$ix]) Then
EndFunc   ;==>VerschbVZ

; aufgerufen in unVerarb
Func Verschieb($ix)
	;	$gesPfad[$ix] = $Pfad & (StringRight($Pfad,1)="\" ? "" : "\") & $urspDat[$ix]
	If FileExists($gesPfad[$ix]) Then
		Local $erg
		For $iru = 1 To 100
			If $iru > 1 Then
				Local $tpos = StringInStr($datName[$ix], ".", 0, -1)
				Local $rumpf = ($tpos > 0 ? StringLeft($datName[$ix], $tpos - 1) : $datName[$ix])
				Local $btyp = ($tpos > 0 ? StringMid($datName[$ix], $tpos) : "")
				$spPfad = $vvz[$ix] & "\" & $rumpf & " (" & $iru & ")" & $btyp
				;										ConsoleWrite("tpos: " & $tpos & ", rumpf: " & $rumpf & ", btyp: " & $btyp & ", spPfad: " & $spPfad & @LF)
			EndIf ; $iru
			ConsoleWrite("gesPfad[$ix]: '" & $gesPfad[$ix] & "', spPfad: '" & $spPfad & "'" & @LF)
			$erg = FileMove($gesPfad[$ix], $spPfad)
			If $erg Then ExitLoop
		Next ; $iru
		If $erg Then
			GUICtrlSetData($iProt, "  4) verschoben in: '" & $spPfad & "'" & @CRLF, 1)
			; noch zu tun: Datenbankeintrag und Verschiebung/Sicherheitskopie nachträglich in Übereinstimmung bringen
			;			$kopart=2
		Else ; $erg
			$lgwFl = $markD[$ix]
			MsgBox($MB_ICONERROR, "Verschiebefehler", "'" & $gesPfad[$ix] & "' konnte nicht nach '" & $vvz[$ix] & "' verschoben werden, breche ab. ")
			Return 0
		EndIf ; $erg Else
	Else ; FileExists($gesPfad[$ix]) Then
		MsgBox($MB_ICONERROR, "Existenzfehler", "Datei '" & $gesPfad[$ix] & "' nicht gefunden. Breche ab.")
		Return 0
	EndIf ; FileExists($gesPfad[$ix]) Then Else
	Return $erg
EndFunc   ;==>Verschieb

; aufgerufen in unVerarb und Verarbeit
Func SicherVZ($ix)
	$svz[$ix] = ""
	If FileExists($gesPfad[$ix]) Then
		ConsoleWrite("Kopin: " & $KopIn & @LF)
		If $KopIn Then
			Local $zvz = $KopIn
			If $Jahr[$ix] Then $zvz = $zvz & "\" & $Jahr[$ix]
			If Not FileExists($zvz) Then DirCreate($zvz)
			ConsoleWrite("zvz: '" & $zvz & "', von: '" & $gesPfad[$ix] & "'" & @LF)
			$svz[$ix] = $zvz
		EndIf ; $KopIn<>"" Then
	EndIf ; FileExists($gesPfad[$ix]) Then Else
EndFunc   ;==>SicherVZ

; aufgerufen in unVerarb und Verarbeit
Func SicherheitsKopie($ix)
	Local $LSc="p:\Laborscan\"
	If FileExists($gesPfad[$ix]) Then
		If $svz[$ix] Then
			Local $zvz = $svz[$ix] & "\", $erg
			For $iru = 1 To 100
				If $iru > 1 Then
					Local $tpos = StringInStr($datName[$ix], ".", 0, -1)
					Local $rumpf = ($tpos > 0 ? StringLeft($datName[$ix], $tpos - 1) : $datName[$ix])
					Local $btyp = ($tpos > 0 ? StringMid($datName[$ix], $tpos) : "")
					$zvz = $svz[$ix] & "\" & $rumpf & " (" & $iru & ")" & $btyp
					;										ConsoleWrite("tpos: " & $tpos & ", rumpf: " & $rumpf & ", btyp: " & $btyp & ", zvz: " & $zvz & @LF)
				EndIf ; $iru
				ConsoleWrite("gesPfad: '" & $gesPfad[$ix] & "', zvz: '" & $zvz & "'" & @LF)
				$erg = FileCopy($gesPfad[$ix], $zvz)
				If $erg Then ExitLoop
			Next ; $iru
			If $erg Then
				GUICtrlSetData($iProt, "  3) gesichert in: '" & $zvz & "'" & @CRLF, 1)
				; noch zu tun: Datenbankeintrag und Sicherheitskopie nachträglich in Übereinstimmung bringen
				;				$kopart=1
			Else
				$lgwFl = $markD[$ix]
				MsgBox($MB_ICONERROR, "Kopierfehler", "'" & $gesPfad[$ix] & "' konnte nicht nach '" & $svz[$ix] & "' kopiert werden, breche ab. ")
				Return 0
			EndIf ; Not FileCopy($benannt,$zvz) Then
			If Not FileExists($LSc & $datname[$ix]) Then
			 If StringInStr($datname[$ix],"Fremdlabor") Then
				$erg = FileCopy($gespfad[$ix],$LSc)
				If $erg Then
					GUICtrlSetData($iProt, "  3b) gesichert in: '" & $LSc & "'" & @CRLF, 1)
					; noch zu tun: Datenbankeintrag und Sicherheitskopie nachträglich in Übereinstimmung bringen
					;				$kopart=1
				Else
					$lgwFl = $markD[$ix]
					MsgBox($MB_ICONERROR, "Kopierfehler", "'" & $gesPfad[$ix] & "' konnte nicht nach '" & $LSc & "' kopiert werden, breche ab. ")
					Return 0
				EndIf ; Not FileCopy($benannt,$zvz) Then
			 EndIf
			EndIf
		EndIf ; $KopIn<>"" Then
	Else ; FileExists($gesPfad[$ix]) Then
		MsgBox($MB_ICONERROR, "Existenzfehler", "Datei '" & $gesPfad[$ix] & "' nicht gefunden. Breche ab.")
		Return 0
	EndIf ; FileExists($gesPfad[$ix]) Then Else
	Return $erg
EndFunc   ;==>SicherheitsKopie

; aufgerufen in Verarbeit
Func Aufraeum($ix, $Pid)
	Local $DPf="p:\dok\", $PIDPf=$DPf & $Pid
	If Not FileExists($DPf) Then DirCreate($DPf)
	If Not FileExists($PIDPf) Then DirCreate($PIDPf)
	If FileExists($gesPfad[$ix]) Then
		If $PIDPf Then
			Local $zvz = $PIDPf & "\", $erg
			For $iru = 1 To 100
				If $iru > 1 Then
					Local $tpos = StringInStr($datName[$ix], ".", 0, -1)
					Local $rumpf = ($tpos > 0 ? StringLeft($datName[$ix], $tpos - 1) : $datName[$ix])
					Local $btyp = ($tpos > 0 ? StringMid($datName[$ix], $tpos) : "")
					$zvz = $PIDPf & "\" & $rumpf & " (" & $iru & ")" & $btyp
					;										ConsoleWrite("tpos: " & $tpos & ", rumpf: " & $rumpf & ", btyp: " & $btyp & ", zvz: " & $zvz & @LF)
				EndIf ; $iru
				ConsoleWrite("gesPfad: '" & $gesPfad[$ix] & "', PIDPf: '" & $PIDPf & "'" & @LF)
				$erg = FileMove($gesPfad[$ix], $zvz)
				If $erg Then ExitLoop
			Next ; $iru
			If $erg Then
				GUICtrlSetData($iProt, "  4) aufgeräumt in: '" & $zvz & "'" & @CRLF, 1)
				; noch zu tun: Datenbankeintrag und Sicherheitskopie nachträglich in Übereinstimmung bringen
				;				$kopart=1
			Else
				$lgwFl = $markD[$ix]
				MsgBox($MB_ICONERROR, "Kopierfehler", "'" & $gesPfad[$ix] & "' konnte nicht nach '" & $PIDPf & "' aufgeräumt werden. ")
				Return 0
			EndIf ; Not FileCopy($benannt,$zvz) Then
		EndIf ; $KopIn<>"" Then
	Else ; FileExists($gesPfad[$ix]) Then
		MsgBox($MB_ICONERROR, "Existenzfehler", "Datei '" & $gesPfad[$ix] & "' nicht gefunden. Breche ab.")
		Return 0
	EndIf ; FileExists($gesPfad[$ix]) Then Else
	Return $erg
EndFunc   ;==>SicherheitsKopie


Func GetAllWindowsControls($hCallersWindow, $bOnlyVisible = Default, $sStringIncludes = Default, $sClass = Default)
	If Not IsHWnd($hCallersWindow) Then
		ConsoleWrite("$hCallersWindow must be a handle...provided=[" & $hCallersWindow & "]" & @CRLF)
		Return False
	EndIf
	; Get all list of controls
	If $bOnlyVisible = Default Then $bOnlyVisible = False
	If $sStringIncludes = Default Then $sStringIncludes = ""
	If $sClass = Default Then $sClass = ""
	Local $sClassList = WinGetClassList($hCallersWindow)

	; Create array
	Local $aClassList = StringSplit($sClassList, @CRLF, 2)

	; Sort array
	_ArraySort($aClassList)
	_ArrayDelete($aClassList, 0)

	; Loop
	Local $iCurrentClass = ""
	Local $iCurrentCount = 1
	Local $iTotalCounter = 1

	If StringLen($sClass) > 0 Then
		For $i = UBound($aClassList) - 1 To 0 Step -1
			If $aClassList[$i] <> $sClass Then
				_ArrayDelete($aClassList, $i)
			EndIf
		Next
	EndIf

	For $i = 0 To UBound($aClassList) - 1
		If $aClassList[$i] = $iCurrentClass Then
			$iCurrentCount += 1
		Else
			$iCurrentClass = $aClassList[$i]
			$iCurrentCount = 1
		EndIf

		Local $hControl = ControlGetHandle($hCallersWindow, "", "[CLASSNN:" & $iCurrentClass & $iCurrentCount & "]")
		Local $text = StringRegExpReplace(ControlGetText($hCallersWindow, "", $hControl), "[\n\r]", "{@CRLF}")
		$aPos = ControlGetPos($hCallersWindow, "", $hControl)
		Local $sControlID = _WinAPI_GetDlgCtrlID($hControl)

		Local $bIsVisible = ControlCommand($hCallersWindow, "", $hControl, "IsVisible")
		If $bOnlyVisible And Not $bIsVisible Then
			$iTotalCounter += 1
			ContinueLoop
		EndIf

		If StringLen($sStringIncludes) > 0 Then
			If Not StringInStr($text, $sStringIncludes) Then
				$iTotalCounter += 1
				ContinueLoop
			EndIf
		EndIf

		If IsArray($aPos) Then
			ConsoleWrite("Func=[GetAllWindowsControls]: ControlCounter=[" & StringFormat("%3s", $iTotalCounter) & "] ControlID=[" & StringFormat("%5s", $sControlID) & "] Handle=[" & StringFormat("%10s", $hControl) & "] ClassNN=[" & StringFormat("%19s", $iCurrentClass & $iCurrentCount) & "] XPos=[" & StringFormat("%4s", $aPos[0]) & "] YPos=[" & StringFormat("%4s", $aPos[1]) & "] Width=[" & StringFormat("%4s", $aPos[2]) & "] Height=[" & StringFormat("%4s", $aPos[3]) & "] IsVisible=[" & $bIsVisible & "] Text=[" & $text & "]." & @CRLF)
		Else
			ConsoleWrite("Func=[GetAllWindowsControls]: ControlCounter=[" & StringFormat("%3s", $iTotalCounter) & "] ControlID=[" & StringFormat("%5s", $sControlID) & "] Handle=[" & StringFormat("%10s", $hControl) & "] ClassNN=[" & StringFormat("%19s", $iCurrentClass & $iCurrentCount) & "] XPos=[winclosed] YPos=[winclosed] Width=[winclosed] Height=[winclosed] Text=[" & $text & "]." & @CRLF)
		EndIf

		If Not WinExists($hCallersWindow) Then ExitLoop
		$iTotalCounter += 1
	Next
EndFunc   ;==>GetAllWindowsControls

Func _ProcessGetHWnd($PID)
    local $WinList = WinList()
    local $Return = 0
    for $i = 1 to $WinList[0][0]
        if WinGetProcess($WinList[$i][1]) = $PID Then
            $Return = $WinList[$i][1]
            ExitLoop
        EndIf
    Next
    Return $Return
EndFunc

; aufgerufen in insArztPraxisProgramm
Func Transportiere($ix, $wthd, $pSystem, ByRef $gefunden)
	Local $zustarten = @ComSpec & " /c explorer.exe /select,""" & $gesPfad[$ix] & """"
	ConsoleWrite("Starte gleich: " & $zustarten & @LF)
	Local $procID = Run($zustarten, $Pfad, @SW_HIDE)
	If $procID = 0 Then
		Switch @extended
			Case 267
				MsgBox(16, "Fehler bei run(" & $zustarten & ")", "Der Verzeichnisname ist ungültig")
			Case Else
				MsgBox(16, "Fehler", "run(" & $zustarten & ") misslang, @error = " & @error & ", @extended = " & @extended)
		EndSwitch            ; @extended
	EndIf            ; $procID=0 Then
	ConsoleWrite("ProcID: " & $procID & ", gesPfad: '" & $gesPfad[$ix] & "'" & @LF)
	;											Opt("WinTitleMatchMode",3)
	ConsoleWrite("vor WinGetHandle, Pfad: " & $Pfad & @LF)
	Local $pWindow = WinGetHandle($Pfad)
	If $pwindow=0 Then
		Local $Owt=Opt("WinTitleMatchMode", 2) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase
		local $subs=StringLeft($pfad,StringLen($pfad)-(StringRight($pfad,1)="\"?1:0))
		$subs=" (" & $subs & ")"
		$pWindow = WinGetHandle($subs)
		ConsoleWrite("PWindow nach match substring " & $subs & ": " & $pWindow & @LF)
        Opt("WinTitleMatchMode",$Owt)
	EndIf
;	If $pwindow=0 then
;		$pWindow=_ProcessGetHWnd($procID)
;		ConsoleWrite("PWindow nach _ProcessGetHWnd: " & $pWindow & @LF)
;	EndIf
	If $pWindow = 0 Then
		ConsoleWrite("vor WinWaitActive, Pfad: " & $Pfad & @LF)
		$pWindow = WinWaitActive($Pfad, "", 20)
	EndIf            ; $pwindow=0 Then
	ConsoleWrite("nach WinWaitActive " & @LF)
	;											Opt("WinTitleMatchMode",1)
	;											Local $pWindow=_GetHwndFromPID($procID)
	ConsoleWrite("pWindow: " & $pWindow & @LF)
	If $pWindow Then
		ConsoleWrite("Handle P" & @LF)
		If WinActivate($pWindow) Then
			ConsoleWrite("Handle P aktiv" & @LF)
			WinWaitActive($pWindow)
			Send("! x")            ; Maximieren
			Local $Color
			For $x = 400 To 1100 Step 100            ; von li nach re
				For $y = 0 To 950 Step 5            ; in Bildschirmmitte
					$Color = PixelGetColor(490, $y, $pWindow)
					ConsoleWrite("y: " & $y & ", Color: " & Hex($Color) & @LF)
					If $Color = 0xCCE8FF Then          ; 1342899110
						ConsoleWrite("!!!!!!!! Farbe " & $Color & " gefunden!" & @LF)
						$gefunden = 1
						MouseMove(490, $y, 0)
						MouseDown("left")
						;															MouseMove(500,$y,5)
						;			{TAB}{TAB}{ALTUP}{ALTUP}")
						;															Local $hMO = WinGetHandle( "[CLASS:OWL_Window]" )
						MouseMove(510, $y, 3)
						WinActivate($wthd)            ; $hMO)
						WinWaitActive($wthd)            ; $hMO)
						If $pSystem Then; Turbomed
							Local $mmo=Opt("MouseCoordMode",2)
							Local $acp = ControlGetPos($wthd,"","AfxFrameOrView1401")
							MouseMove($acp[0]+$acp[2]/2, $acp[1]+10)
							Opt("MouseCoordMode",$mmo)
						Else; Medical Office; TmoStringGrid3 oder TmoStringGrid20, zu ermitteln s. CtrlGetPos.au3
							MouseMove(412, 400)
						EndIf; $pSystem Then
 					    Sleep(2000)
						MouseUp("left")
						If $pSystem = 0 Then
							Sleep(2000)
							If $typ[$ix] = "pdf" Then Send("p")
							Sleep(1000)
							Send("{TAB}")
							Sleep(2000)
							Send($datName[$ix] & "{Enter}")
						EndIf            ; $pSystem
						GUICtrlSetData($iProt, "  4) importiert in " & ($pSystem ? "Turbomed" : "Medical Office") & ": '" & $gesPfad[$ix] & "'" & @CRLF, 1)
						machfilelist(1)            ; für $idDtn; wenn Datei entfernt, dann Dateiliste neu anzeigen
						ExitLoop            ; For $y
					EndIf            ; $Color=0xCCE8FF Then; 1342899110
				Next            ; $y
				If $gefunden Then ExitLoop            ; For $x
			Next            ; $x
			If $gefunden Then
				ConsoleWrite("gefunden an y: " & $y & @LF)
			Else            ; $gefunden
				MsgBox(64, "Dateisuchfehler", "Datei anhand der Farbe nicht im Explorer gefunden, bitte Explorerfenster schließen und Programm neu starten", 5)
			EndIf            ; $gefunden
		EndIf            ; WinActivate($pWindow) Then
		;	ConsoleWrite("y: " & $y & @LF)
	EndIf            ; $pWindow Then
	ProcessClose($procID)
	;				ExitLoop; For $iru = 1 To 6
	Return 1            ;
EndFunc   ;==>Transportiere


Func insArztPraxisProgramm($ix, $Pid) ; Rückgabe 1 = vermutlich Erfolg
	Local $hAct = WinGetHandle("[ACTIVE]")
	Local $gefunden = 0
	Local $procID, $wthd, $hd, $titel
	If $pSystem = 0 Then ; Medical Office
		If UBound($pwd) = 0 Then ReDim $pwd[1]
		$titel = "Medical Office - Zentrale"
		Local $titi = "MEDICAL OFFICE Login"
		Local $wthi = WinGetHandle($titi)
		$wthd = WinGetHandle($titel)
		ConsoleWrite("Stelle -1, $wthi: " & $wthi & ", wthd: " & $wthd & @LF)
		If $wthi = 0 And $wthd = 0 Then
			$procID = Run("C:\INDAMED\medoff.exe")
			$wthi = WinWaitActive($titi)
		EndIf ; $procID=0 Then
;		ConsoleWrite("Stelle 0" & @LF)
		If $wthi Then
;			ConsoleWrite("Stelle 1" & @LF)
			If WinActivate($titi) Then ; Passworteingabe
				ConsoleWrite("Stelle 2" & @LF)
				$wthi = WinWaitActive($titi)
				Local $hTim = TimerInit()
				$passwd = $pwd[0]
				While $passwd = ""
					$passwd = InputBox("Medical Office Passwortabfrage", "Medical Office Passwort")
					If $passwd <> "" Then
						$pwd[0] = $passwd
					EndIf ; $passwd<>"" Then
				WEnd ; $passwd=""
				ControlSetText($wthi, "", "[CLASS:TEdit; INSTANCE:1]", $passwd)
				ControlClick($wthi, "", "[CLASS:TNewBitBtn; INSTANCE:1]")
				;  Send("{TAB}87?oisOk__{TAB 2} !o")
				; 	  ConsoleWrite("Stelle 3" & @LF)
				WinActivate($wthd)
				; 	  ConsoleWrite("Stelle 4, wthd: " & $wthd & @LF)
				WinWaitActive($wthd, "", 1)
				; 	  ConsoleWrite("Stelle 5" & @LF)
				GetAllWindowsControls($wthd)
				; 	  ConsoleWrite("Stelle 6" & @LF)
			EndIf ; WinActivate($titel) Then
		EndIf ; $wthd Then
		If WinActivate($titel) Then
			Local $pname[] = ["Fallauswahl für", "Vertragsinformationen", "Patienteninfo", "Erinnerungen", "Externe Datei"]
			Local $pbutt[] = ["TmoButton2", "TmoButton1", "TNewBitBtn1", "", ""]
			For $aru = 1 To 2
				While 1
					Local $sauber = 1
					For $i = 0 To UBound($pname) - 1
						Local $wthf = WinGetHandle($pname[$i])
						If $wthf Then
							If BitAND(WinGetState($wthf), $WIN_STATE_VISIBLE) Then
								WinActivate($wthf)
								; 	  ConsoleWrite("Stelle 7" & @LF)
								$wthf = WinWaitActive($pname[$i])
								; 	  ConsoleWrite("Stelle 8" & @LF)
								If $wthf Then
									If $pbutt[$i] Then
										ControlClick($wthf, "", $pbutt[$i])
									Else
										Send("{ESC}")
									EndIf
								EndIf
								ConsoleWrite($pname[$i] & " gefunden" & @LF)
								$sauber = 0
							Else
								ConsoleWrite($pname[$i] & " nicht gefunden" & @LF)
							EndIf
						Else
							ConsoleWrite($pname[$i] & " nicht gefunden" & @LF)
						EndIf
					Next ; $i
					If $sauber Then ExitLoop
					Sleep(20)
				WEnd
				If $aru = 2 Then ExitLoop
				Send("{F4}")
				Sleep(20)
				Send($Pid)
				Send("{ENTER 2}")
				Sleep(5000)
			Next ; $aru
			Local $kkpos
			For $j = 0 To 1000
				$kkpos = ControlGetPos($wthd, "", "#327702")
				If @error = 0 Then
					ConsoleWrite("ControlGetPos gelungen, j=" & $j & ", kkpos[1]=" & $kkpos[1] & @LF)
					ExitLoop
				EndIf
				Sleep(10)
			Next ; $j
			If @error Then
				ConsoleWrite("ControlGetPos(" & $wthd & ") misslang, @error = " & @error & ", @extended = " & @extended & @LF)
				MsgBox(16, "Fehler", "ControlGetPos(" & $wthd & ") misslang, @error = " & @error & ", @extended = " & @extended)
				Return @error
			EndIf
			If $kkpos[1] < 101 Then
				ConsoleWrite("schalte um" & @LF)
				Send("{F6}")
			EndIf
			;		Sleep(20)
			;		GetAllWindowsControls($wthd)
			;	    ConsoleWrite("wthd: " & $wthd & ", Achtung: " & ControlGetText($wthd,"","[CLASS:OWL_Window:F; INSTANCE:2]") & @LF)
		EndIf ; WinActivate($titel) Then
		Transportiere($ix, $wthd, $pSystem, $gefunden)
		ConsoleWrite("Ende" & @LF)

	ElseIf $pSystem = 1 Then ; Turbomed
		Local $list = ProcessList()
		$titel = "Benutzeranmeldung - CGM TURBOMED"
		For $i = 1 To $list[0][0]
			;	  ConsoleWrite($list[$i][0] & " " & $list[$i][1] & @LF)
			If $list[$i][0] = "TurboMed.exe" Then
				$procID = $list[$i][1]
				$wthd = WinGetHandle($titel)
				ExitLoop
			EndIf ; $list[$i][0]="TurboMed.exe" Then
		Next ; $i = 1 To $list[0][0]
		If $procID = 0 Then
			$procID = Run("c:\turbomed\programm\turbomed.exe")
			$wthd = WinWaitActive($titel)
		EndIf ; $procID=0 Then
		If $wthd Then
			If WinActivate($titel) Then
				$wthd = WinWaitActive($titel)
				;  ControlFocus($wthd, "", "Edit1")
				;  Local $ben=ControlGetText($wthh,"","Edit1")
				Local $hTim = TimerInit()
				;                                   Benutzeranmeldung - CGM TURBOMED
				While 1
					If ControlGetHandle($wthd, "", "Edit1") Then ExitLoop
					If TimerDiff($hTim) > 120000 Then ExitLoop
				WEnd ; 1
				Local $ben = ControlGetText($wthd, "", "Edit1")
				While $ben = ""
					$ben = InputBox("Turbomed Benutzerabfrage", "Turbomedbenutzer")
				WEnd ; $ben=""
				$passwd = ""
				For $i = 0 To UBound($usr) - 1
					If $usr[$i] = $ben Then
						$passwd = $pwd[$i]
						ExitLoop
					EndIf
				Next ; $i=0 To UBound($usr)-1
				;	ConsoleWrite("Stelle 450, passwd: " & $passwd & @LF)
				While $passwd = ""
					$passwd = InputBox("Turbomed Passwortabfrage", "Turbomedpasswort für Benutzer " & $ben, "", "*")
					If $passwd <> "" Then
						ReDim $usr[UBound($usr) + 1]
						$usr[UBound($usr) - 1] = $ben
						ReDim $pwd[UBound($pwd) + 1]
						$pwd[UBound($pwd) - 1] = $passwd
					EndIf ; $passwd<>"" Then
				WEnd ; $passwd=""
				;  ConsoleWrite("Benutzer: " & $usr[0] & ", pwd: " & $pwd[0] & @LF)
				ControlSetText($wthd, "", "Edit2", $passwd)
				ControlClick($wthd, "", "Button2")
				;  Send("{TAB}87?oisOk__{TAB 2} !o")
				Sleep(1000)
			EndIf ; WinActivate($titel) Then
		EndIf ; $wthd Then

		$titel = "TURBOMED - [Bitte wählen Sie die Ärzte aus"
		If WinActivate($titel) Then
;			ConsoleWrite("Stelle 460a" & @LF)
			$wthd = WinWaitActive($titel)
;			ConsoleWrite("Stelle 460b" & @LF)
			;	ConsoleWrite("wthd: " & $wthd & @LF)
			ControlClick($wthd, "", "Static7")
			Sleep(8000)
			;									Send("{ESC}!w{ESC}"); neue Oberfläche
		EndIf ; WinActivate($titel) Then
		;								Send("!{TAB}")
;		MsgBox(64,"Zeitanzeige","Stelle 1")
		$titel = "TURBOMED" ; "[CLASS:Qt5153QWindow]"
		ConsoleWrite("hAct: " & $hAct & ", Stelle 454___" & @LF)
		If WinActivate($titel) Then ; $hAct)
			$wthd = WinWaitActive($titel) ; $hAct)
		EndIf ; WinActivate($titel)
;		MsgBox(64,"Zeitanzeige","Stelle 1a")
		;                               wenn Turbomed, aber nicht Turbomed - {, dann p drucken
		;								Local $pid=1618
		;								$wthd=0
		$titel = "TURBOMED - [" & $Pid
		For $iru = 1 To 6
			Local $akttit = WinGetTitle("[active]")
			If StringLeft($akttit, StringLen($titel)) = $titel Then
				$wthd = $akttit
			Else
;				ConsoleWrite("Stelle 500, iru: " & $iru & @LF)
;		MsgBox(64,"Zeitanzeige","Stelle 1b")
				WinActivate($titel)
				For $iiru = 0 To 3
					$wthd = WinWaitActive($titel, "", 1) ; wenn der neue Patient (ggf. schon von vorher) aktiv ist
;					ConsoleWrite("Stelle 501, iru: " & $iru & @LF)
					If $wthd Then ExitLoop
				Next ; $iiru
;		MsgBox(64,"Zeitanzeige","Stelle 1c")
			EndIf
;			ConsoleWrite("Stelle 501a, iru: " & $iru & @LF)
;		MsgBox(64,"Zeitanzeige","Stelle 2, iru: " & $iru)
			If $wthd Then
				Sleep(100)
				Local $karteikarte = ControlGetHandle($wthd, "", "AfxFrameOrView1401")
				If $karteikarte Then
					If $Reit Then
						Sleep($Itv)
						;					ConsoleWrite("Reiter: " & $Reit & @LF)
						Send("!{" & (StringLeft($Reit, 1) = "F" ? "" : "F") & $Reit & "}") ; Reiter "Gescanntes"
					EndIf ; $Reit
					; vorbestehende Explorerfenster mit vielleicht falscher Dateimarkierung schließen
					Do
						Local $pWindow = WinGetHandle($Pfad)
						If @error Then ExitLoop
						WinClose($pWindow)
					Until 0
					If Transportiere($ix, $wthd, $pSystem, $gefunden) Then ExitLoop
				Else ; $karteikarte Then
					;			ConsoleWrite("keine Karteikarte" & @lf)
				EndIf ; $karteikarte Then Else
			Else ; $wthd Then
				Local $ktit = "TURBOMED"
				If WinActivate($ktit) Then
					$wthd = WinWaitActive($ktit)
					Send("{Esc}")
					Sleep($Itv)
					Send("{Esc}")
					Sleep($Itv)
					Send("{Esc}")
					Sleep($Itv)
					Send("{Enter}")
					Sleep($Itv)
					Send("{F12}")
					Sleep($Itv)
					Send("p")
					Sleep(3 * $Itv)
					Send("{BS}" & $Pid)
					ConsoleWrite("nach {BS}pid: " & $Pid & @LF)
					Sleep($Itv)
					Send("{Enter}")
					Sleep(2 * $Itv)
					Send("{F3}")
					For $k=0 To 1000
						Local $karteikarte = ControlGetHandle($wthd, "", "AfxFrameOrView1401")
						If $karteikarte Then
;							MsgBox(64,"Zeitanzeige","Karteikkarte offen!, iru: " & $iru)
							ExitLoop
						endif
						Sleep(10)
					Next
;					Sleep(1000)
					For $brunde = 1 To 2 ; Hinweis, der Patient ist verstorben
						If ControlGetHandle($wthd, "", "AfxFrameOrView1401") Then ExitLoop
						If $brunde = 1 Then
							MouseMove(670, 570, 0)
							MouseClick("left")
						Else
							Send("!j")
						EndIf
						Sleep($Itv)
					Next ; $brunde
				EndIf ; WinActivate($ktit) Then; neue Oberfläche
				ConsoleWrite("Stelle 1, pid: " & $Pid & @LF)
;		MsgBox(64,"Zeitanzeige","Stelle 3")
			EndIf ; $wthd Then Else
		Next ; $iru
	EndIf ; $pSystem=0 Else
	Return $gefunden
EndFunc   ;==>insArztPraxisProgramm



; aufgerufen in Verarbeit und unVerarb
Func DBEintrag($ix, $gebDat, $Ort, $Pid, $kopart) ; Rückgabe: 1 = Erfolg; $kopart: 0= keine Kopie erstellt; $kopart: 0=keine, 1=nur Sicherheitskopie, 2=auch Verschiebung
	; 1= Sicherheitskopie, 2= Verschiebung
	ConsoleWrite("1 Kopart: " & $kopart & ", ix: " & $ix & @LF)
	;	_ArrayDisplay($svz)
	;	_ArrayDisplay($vvz)
	Local $mDix = $markD[$ix]
	If $cSvr[$Vbzl - 1] <> "" And $cUsr[$Vbzl - 1] And $cDb[$Vbzl - 1] Then
		$sql = _
				"CREATE TABLE IF NOT EXISTS `dokprotlist` (" & @CRLF & _
				"	`id` INT(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'," & @CRLF & _
				"	`kPatN` BIT(1) NOT NULL DEFAULT b'0' COMMENT '1=kein Patientenname'," & @CRLF & _
				"	`ntum` BIT(1) NOT NULL DEFAULT b'0' COMMENT '1=nicht umbenennen'," & @CRLF & _
				"	`nImp` BIT(1) NOT NULL DEFAULT b'0' COMMENT '1=nicht importieren'," & @CRLF & _
				"	`urspnm` BIT(1) NOT NULL DEFAULT b'0' COMMENT '1=ursprünglichen Dateinamen anhängen'," & @CRLF & _
				"	`Patientenname` VARCHAR(50) NOT NULL DEFAULT '\'\'' COMMENT 'Patientenname nach Auswahl' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
				"	`Gebdat` DATE NOT NULL DEFAULT '0000-00-00' COMMENT 'Geburtsdatum'," & @CRLF & _
				"	`Ort` VARCHAR(50) NULL DEFAULT '\'\'' COMMENT 'Wohnort' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
				"	`Pat_ID` INT(10) NOT NULL DEFAULT '0' COMMENT 'Patientennummer im Patientendatenverwaltungsprogramm'," & @CRLF & _
				"	`urspName` VARCHAR(200) NOT NULL DEFAULT '\'\'' COMMENT 'ursprünglicher Dateiname' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
				"	`datName` VARCHAR(300) NOT NULL DEFAULT '\'\'' COMMENT 'neuer Dateiname' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
				"	`lAend` DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00' COMMENT 'letzte Dateiänderung'," & @CRLF & _
				"	`groesse` INT(10) NOT NULL DEFAULT '0' COMMENT 'Dateigröße'," & @CRLF & _
				"	`Typ` VARCHAR(7) NOT NULL DEFAULT '\'\'' COMMENT 'Dateityp' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
				"   `PC` VARCHAR(10) NOT NULL DEFAULT '\'\'' COMMENT 'Computername' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
				"   `Benutzer` VARCHAR(10) NOT NULL DEFAULT '\'\'' COMMENT 'PC-Benutzername' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
				"	`Mitarbeiter` VARCHAR(4) NOT NULL DEFAULT '\'\'' COMMENT 'Mitarbeiter (Kürzel)' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
				"	`SicherVzID` INT(10) NOT NULL DEFAULT '0' COMMENT 'Sicherheitsverzeichnis-ID (dokprotvz)'," & @CRLF & _
				"	`ZielVzID` INT(10) NOT NULL DEFAULT '0' COMMENT 'Zielverzeichnis-ID (dokprotvz)'," & @CRLF & _
				"	`eingetragen` DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00' COMMENT 'Eintagsdatum in diese Tabelle'," & @CRLF & _
				"	PRIMARY KEY (`id`) USING BTREE," & @CRLF & _
				"	INDEX `Pat_ID` (`Pat_ID`) USING BTREE," & @CRLF & _
				"	INDEX `Mitarbeiter` (`Mitarbeiter`) USING BTREE," & @CRLF & _
				"	INDEX `urspName` (`urspName`) USING BTREE," & @CRLF & _
				"	INDEX `datName` (`datName`) USING BTREE," & @CRLF & _
				"	INDEX `lAend` (`lAend`) USING BTREE," & @CRLF & _
				"	INDEX `Ort` (`Ort`) USING BTREE," & @CRLF & _
				"	INDEX `Gebdat` (`Gebdat`) USING BTREE," & @CRLF & _
				"	INDEX `groesse` (`groesse`) USING BTREE," & @CRLF & _
				"	INDEX `Typ` (`Typ`) USING BTREE," & @CRLF & _
				"	INDEX `eingetragen` (`eingetragen`, `PC`, `Benutzer`) USING BTREE," & @CRLF & _
				"	INDEX `kPatNam` (`kPatN`) USING BTREE," & @CRLF & _
				"	INDEX `nUmben` (`ntum`) USING BTREE," & @CRLF & _
				"	INDEX `nImport` (`nImp`) USING BTREE," & @CRLF & _
				"	INDEX `urspnm` (`urspnm`) USING BTREE," & @CRLF & _
				"	INDEX `SicherVzID` (`SicherVzID`) USING BTREE," & @CRLF & _
				"	INDEX `ZielVzID` (`ZielVzID`) USING BTREE" & @CRLF & _
				")" & @CRLF & _
				"COMMENT='Protokolldatei der importierten Dokumente'" & @CRLF & _
				"COLLATE='utf8mb3_unicode_ci'" & @CRLF & _
				"ENGINE=InnoDB;"
		If FragProt($sql) = 0 Then
			Return 0
		Else ; FragProt
			$sql = _
					"CREATE TABLE IF NOT EXISTS `dokprotinh` (" & @CRLF & _
					"	`id` INT(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärindex'," & @CRLF & _
					"	`dplbez` INT(10) NOT NULL COMMENT 'Bezug auf id von dokprotlist'," & @CRLF & _
					"	`nr` INT(2) NOT NULL COMMENT 'laufende Nummer pro Bezug'," & @CRLF & _
					"	`Inhalt` VARCHAR(50) NOT NULL DEFAULT '' COMMENT 'Inhalt' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
					"	`von` DATE NOT NULL DEFAULT '0000-00-00' COMMENT 'von'," & @CRLF & _
					"	`bis` DATE NOT NULL DEFAULT '0000-00-00' COMMENT 'bis'," & @CRLF & _
					"	PRIMARY KEY (`id`) USING BTREE," & @CRLF & _
					"	INDEX `dplbez` (`dplbez`, `nr`) USING BTREE," & @CRLF & _
					"	INDEX `Inhalt` (`Inhalt`) USING BTREE," & @CRLF & _
					"	INDEX `von` (`von`) USING BTREE," & @CRLF & _
					"	INDEX `bis` (`bis`) USING BTREE" & @CRLF & _
					")" & @CRLF & _
					"COMMENT='Inhaltsangaben  und Datum der importieren Dokumente'" & @CRLF & _
					"COLLATE='utf8mb3_unicode_ci'" & @CRLF & _
					"ENGINE=InnoDB"
			If FragProt($sql) = 0 Then
				Return 0
			Else ; FragProt
				$sql = _
						"CREATE TABLE IF NOT EXISTS `dokprotvz` (" & @CRLF & _
						"	`id` INT(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'," & @CRLF & _
						"	`Verzeichnis` VARCHAR(300) NOT NULL DEFAULT '\'\'' COMMENT 'Verzeichnisname' COLLATE 'utf8mb3_unicode_ci'," & @CRLF & _
						"	PRIMARY KEY (`id`) USING BTREE," & @CRLF & _
						"	INDEX `Verzeichnis` (`Verzeichnis`) USING BTREE" & @CRLF & _
						")" & @CRLF & _
						"COMMENT='Verzeichnisse für Sicherheitskopie und Verschiebung'" & @CRLF & _
						"COLLATE='utf8mb3_unicode_ci'" & @CRLF & _
						"ENGINE=InnoDB" & @CRLF & _
						"ROW_FORMAT=DYNAMIC"
				If FragProt($sql) = 0 Then
					Return 0
				Else ; FragProt
					Local $jetzt = StringReplace(StringReplace(StringReplace(_NowCalc(), ":", ""), "/", ""), " ", ""), $resu, $idArr, $idstr[2]
					$idstr[0] = ""
					$idstr[1] = ""
					If $kopart Then
						ConsoleWrite("3 Kopart: " & $kopart & @LF)
						For $aru = 0 To $kopart - 1
							For $i = 1 To 2
								$sql = "SELECT id FROM dokprotvz WHERE Verzeichnis = '" & doUmwfSQL($aru ? $vvz[$ix] : $svz[$ix]) & "'"
								ConsoleWrite("sql: " & $sql & @LF)
								If MyProtAnf($ProtCn) Then
									If _MySQL_Real_Query($ProtCn, $sql) = 0 Then
										$resu = _MySQL_Store_Result($ProtCn)
										If $resu Then
											$idArr = _MySQL_Fetch_Result_StringArray($resu)
											_MySQL_Free_Result($resu)
											If UBound($idArr, 1) > 0 Then
												$idstr[$aru] = $idArr[1][0]
												ExitLoop
											EndIf ; UBound($idArr,1)>0 Then
										EndIf ; $res
									EndIf ; _MySQL_Real_Query
								EndIf ; MyProtAnf($ProtCn) Then
								$sql = "INSERT INTO dokprotvz(Verzeichnis) VALUES('" & doUmwfSQL($aru ? $vvz[$ix] : $svz[$ix]) & "')"
								ConsoleWrite("sql: " & $sql & @LF)
								FragProt($sql)
							Next ; $i = 1 To 2
						Next ; $aru
					EndIf ; $kopart Then
					$sql = _
"INSERT INTO tmbrie(fid,pat_id,tm_pat_id,zeitpunkt,quelldatum,Pfad,Art,NAME,Autor,AktZeit,DokGroe,DokAenD,StByte,qs,qt)" & @CRLF & _
"VALUES(COALESCE((SELECT MAX(fid) FROM faelle WHERE pat_id=" & $pid & " AND laend BETWEEN bhfb AND bhfe1),COALESCE((SELECT MAX(fid) FROM faelle WHERE pat_id=" & $pid & " AND laend > bhfb),(SELECT MAX(fid) FROM faelle WHERE pat_id=" & $pid & ")))," & @CRLF & _
 $pid & ",pidtm(" & $pid & "),'" & $jetzt & "',quelldat('" & doUmwfSQL($patName[$ix]) & "','" & $dlaend[$mDix] & "'),CONCAT('p:\\eingelesen\\',YEAR('" & $dlaend[$mDix] & "'),'\\','" & doUmwfSQL($patName[$ix]) & "'),'" & $typ[$mDix] & "','" & doUmwfSQL($patName[$ix]) & "','" & $Kuerzel & "',NOW()," & StringReplace($dgroe[$mDix], ".", "") & ",'" & $dlaend[$mDix] & "',-11,zqsort('" & $dlaend[$mDix] & "'),zquart('" & $dlaend[$mDix] & "'))"
					FragProt($sql)
					$sql = _
							"INSERT INTO dokprotlist(kPatN,ntum,nImp,urspnm,Patientenname,Gebdat,Ort,Pat_ID,urspName,datName,lAend,groesse,Typ,Mitarbeiter,PC,Benutzer,eingetragen" & (($kopart And $idstr[0]) ? ",ZielVZID" : "") & (($kopart = 2 And $idstr[1]) ? ",SicherVZID" : "") & ")" & @CRLF & _
							"VALUES(" & $obkPatN & "," & $obntum & "," & $obnImp & "," & $oburspNm & ",'" & doUmwfSQL($patName[$ix]) & "','" & $gebDat & "','" & $Ort & "','" & $Pid & "','" & doUmwfSQL($urspDat[$ix]) & "','" & doUmwfSQL($datName[$ix]) & "','" & $dlaend[$mDix] & "'," & StringReplace($dgroe[$mDix], ".", "") & ",'" & $typ[$mDix] & "','" & $Kuerzel & "','" & @ComputerName & "','" & @UserName & "','" & $jetzt & (($kopart And $idstr[0]) ? "','" & $idstr[0] : "") & (($kopart = 2 And $idstr[1]) ? "','" & $idstr[1] : "") & "')"
					ConsoleWrite("dpl sql: " & $sql & @LF)
					If FragProt($sql) = 0 Then
						Return 0
					Else ; FragProt
						$sql = "SELECT id FROM dokprotlist WHERE PC='" & @ComputerName & "' AND Benutzer='" & @UserName & "' AND eingetragen = '" & $jetzt & "'"
						If MyProtAnf($ProtCn) Then
							If _MySQL_Real_Query($ProtCn, $sql) = 0 Then
								$resu = _MySQL_Store_Result($ProtCn)
								If $resu Then
									Local $array = _MySQL_Fetch_Result_StringArray($resu)
									_MySQL_Free_Result($resu)
									If UBound($array, 1) > 0 Then
										For $i = 0 To $BZ - 1
											Local $inh = GUICtrlRead($iInh[$i])
											If $inh <> "" Then
												$sql = "INSERT INTO dokprotinh(dplbez,nr,Inhalt,von,bis) VALUES(" & $array[1][0] & "," & $i & ",'" & $inh & _
														"',STR_TO_DATE('" & GUICtrlRead($iD1[$i]) & "','%e.%c.%y'),STR_TO_DATE('" & GUICtrlRead($iD2[$i]) & "','%e.%c.%y'))"
												;										  ConsoleWrite("sql: " & $sql & @LF)
												If FragProt($sql) = 0 Then
													;														ConsoleWrite("Stelle 107" & @LF)
													Return 0
												EndIf ; FragProt
											EndIf ; $inh<>"" Then
										Next ; $i
									Else ; UBound($array,1)>0 Then
										Return 0
									EndIf ; UBound($array,1)>0 Then Else
								Else ; $resu
									Return 0
								EndIf ; $resu Else
							Else ; _MySQL_Real_Query($ProtCn, $sql)=0 Then
								Return 0
							EndIf ; _MySQL_Real_Query($ProtCn, $sql)=0 Then Else
							_MySQL_Close($ProtCn)
							$ProtCn = 0
						Else ; MyProtAnf($ProtCn) Then
							Return 0
						EndIf ; MyProtAnf($ProtCn) Then
					EndIf ; FragProt($sql)=0 Then
				EndIf ; FragProt($sql)=0 Then
			EndIf ; ; FragProt($sql)=0 Then
		EndIf ; FragProt($sql)=0 Then
		GUICtrlSetData($iProt, "  2) In Datenbank eingetragen: '" & $datName[$ix] & "'" & @CRLF, 1)
	EndIf ; $obweiter And $cSvr[$Vbzl-1]<>"" And $cUsr[$Vbzl-1] And $cDb[$Vbzl-1] Then
	Return 1 ; wenn kein Datenbankeintrag vorgesehen, dann auch o.k.
EndFunc   ;==>DBEintrag

; aufgerufen in Schleife
Func VPatUn()
	Local $von, $nach
	If UBound($markD) Then
		dimensionier(UBound($markD))
		For $ix = 0 To UBound($markD) - 1
			Local $top = 10
			$patName[$ix] = ""
			$datName[$ix] = ""
			DNamen2($ix, 1)
			$von = $von & ($ix ? ($ix < UBound($markD) - 1 ? ", " : " und ") : "") & $aDL[$markD[$ix]]
			$nach = $nach & ($ix ? ($ix < UBound($markD) - 1 ? ", " : " und ") : "") & $datName[$ix]
		Next ; $ix
		$unFrage = "Wo soll" & (UBound($markD) > 1 ? "en" : "") & "  '" & $von & "' als '" & $nach & "' gespeichert werden?"
		$idVPat = GUICreate($unFrage, 600, 900)
		fuelleUnPfade()
		If $lUnZahl >= 1 Then
			For $i = 0 To UBound($arrUnPf) - 1
				$idUnPf[$i] = GUICtrlCreateButton($arrUnPf[$i], 5, $top)
				$top = $top + 30
			Next ; $i
		EndIf ; $lUnZahl>=1 Then
		$idNeuPf = GUICtrlCreateButton("neuen Pfad wählen", 5, $top)
		GUISetState()
	EndIf ; UBound($markD)
EndFunc   ;==>VPatUn

; aufgerufen und Haupt und ComZu
Func zeichneHaupt() ; für $idHaupt
	$idHaupt = GUICreate("", 220, 250, 100, 200, BitOR($WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_SYSMENU, $WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_TABSTOP, $WS_EX_CONTROLPARENT), $WS_EX_ACCEPTFILES)
	GUISetState(@SW_HIDE, $idHaupt)
	WinSetTitle($idHaupt, "", "Dateien in '" & $Pfad & "' benennen (Pat. aus " & $cSvr[$cRd] & ", User: " & $cUsr[$cRd] & ", Db: " & $cDb[$cRd] & ", Port: " & $cPrt[$cRd] & ")")
	GUIRegisterMsg($WM_SIZE, 'WM_SIZE')
	If $mitImport then
		GUISetBkColor(0x00E0FFFF) ; will change background color
	Else
		GUISetBkColor($COLOR_LIGHTGREEN) ; 0x90EE90); will change background color
	EndIf; $mitImport
	$idFile = GUICtrlCreateMenu("Men&ü")
	$idFileitem = GUICtrlCreateMenuItem("&Pfad wählen", $idFile)
	GUICtrlSetState(-1, $GUI_DEFBUTTON)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED) ; to allow drag and dropping
	fuelleLetztePfade()
	$iOpt = GUICtrlCreateMenuItem("&Optionen", $idFile)
	$idExit = GUICtrlCreateMenuItem("&Beenden", $idFile)
	GUISetState(@SW_MAXIMIZE)
	$aPos = WinGetPos("[ACTIVE]")
	ConsoleWrite("Position: " & $aPos[0] & ", " & $aPos[1] & ", " & $aPos[2] & ", " & $aPos[3] & @LF)
	$aML = $aPos[2] / 2
	$aDUR = $aPos[3] * 0.667 - 50 + (8 - $BZ) * 30
	ConsoleWrite("aML: " & $aML & ", aDUR: " & $aDUR & @LF)

	If $obPDF <> "" Then machPdf()
	If $obWd <> "" Then machWord()
	If $obEx <> "" Then machExcel()
	machfilelist(1) ; legt $idDtn fest

	;	Opt("GUIDataSeparatorChar",";")
	Sleep(20)
	$iDtAufr = GUICtrlCreateButton("Datei &aufrufen", $aML, $aDUR, 80, 20) ; Position s. $idDtn
	$iProt = GUICtrlCreateEdit("", 3, $aDUR, $aML - 5, $aPos[3] - $aDUR - 50, $ES_READONLY + $ES_AUTOVSCROLL + $WS_VSCROLL)
	GUICtrlSetFont($iProt, 9, 400, $GUI_FONTNORMAL)
	$idurspNm = GUICtrlCreateCheckbox("+ &urspr.Namen", $aML + 87, $aDUR)
	$oburspNm = $oburspNVgb
	If $oburspNm <> 0 Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idkPatN = GUICtrlCreateCheckbox("&k.Pat'nam", $aML + 182, $aDUR) ; keinen Patientenamen zuteilen
	$obkPatN = $obkPatVgb
	If $obkPatN <> 0 Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idntUm = GUICtrlCreateCheckbox("&nicht umben.", $aML + 256, $aDUR)
	$obntum = $obkPatVgb
	If $obntUVgb <> 0 Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idnimp = GUICtrlCreateCheckbox("n.impo&rt.", $aML + 344, $aDUR) ; nicht importieren
	$obnImp = $obnimVgb
	If $obnImp <> 0 Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idVALb = GUICtrlCreateLabel("&Mitarbeiter:", $aML + 450, $aDUR + 5, 52, 20)
	$idVA = GUICtrlCreateCombo("Krz, Vorname Nachname", $aML + 505, $aDUR, 125, 25)
	fuellMA() ; für $idVA
	$irefr = GUICtrlCreateDummy()
	$idlzeig = GUICtrlCreateDummy()

	$idPatWLb = GUICtrlCreateLabel("&Patient:", $aML, $aDUR + 35, 39)
	$idPatWahl = GUICtrlCreateCombo("Nachname, Vorname, Gebdat, Ort, PID", $aML + 40, $aDUR + 30, $Weite - $aML - 40 - 2, 20)
	;    $query = "SELECT GROUP_CONCAT(CONCAT(FNachname,', ',FVorname,', ',DATE_FORMAT(FGeburtsdatum,'%d.%m.%Y'),', ',FOrt,', ',FSurogat) ORDER BY fnachname,fvorname SEPARATOR '|') FROM patstamm;"
	;    $query = "SELECT GROUP_CONCAT(CONCAT(nachname, ', ',vorname,', ',DATE_FORMAT(gebdat,'%d.%m.%Y'),', ',ort,', ',pat_id) ORDER BY Nachname,Vorname SEPARATOR '|') FROM namen;"
	fuellPatWahl() ; für $idPatWahl

	_ArraySort($Begr)
	Local $Begn ; Begriffe
	For $i = 0 To $BegZ - 1
		If $i > 0 Then $Begn = $Begn & "|"
		$Begn = $Begn & $Begr[$i]
	Next ; $i
	For $i = 0 To $BZ - 1
		$iILb[$i] = GUICtrlCreateLabel("&Inhalt" & ($i + 1) & ":", $aML, $aDUR + 69 + 30 * $i, 39, 30)
		;		$iInh[$i]=GUICtrlCreateInput("",$aML+40,$aDUR+60+30*$i,$Weite-165-$aML-40,25)
		$iInh[$i] = GUICtrlCreateCombo("", $aML + 40, $aDUR + 60 + 30 * $i, $Weite - 165 - $aML - 40, 25)
		GUICtrlSetData(-1, $Begn)
		$iD1Lb[$i] = GUICtrlCreateLabel("Da&tum:", $Weite - 165, $aDUR + 69 + 30 * $i, 38, 30)
		$iD1[$i] = GUICtrlCreateInput("", $Weite - 129, $aDUR + 60 + 30 * $i, 60, 25)
		$iD2Lb[$i] = GUICtrlCreateLabel("-", $Weite - 67, $aDUR + 69 + 30 * $i, 9, 30)
		$iD2[$i] = GUICtrlCreateInput("", $Weite - 62, $aDUR + 60 + 30 * $i, 60, 25)
		;       Combo mit Mitarbeiterauswahl
	Next ; $i
	$idPatUn = GUICtrlCreateButton("Pati&entenunabhängig speichern", $Weite - 200, $Hoehe - 25)
	$idVerarb = GUICtrlCreateButton("&Ok", $Weite - 30, $Hoehe - 25)
	GUICtrlSetState($idVerarb, $GUI_DISABLE)

	Local $aAccelKeys[2][2] = [["!d", $idlzeig], ["{F5}", $irefr]]
	GUISetAccelerators($aAccelKeys)
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
	;;    GUIRegisterMsg($WM_COMMAND, "WM_COMMAND") ; WM_COMMAND mit en_change führte z.T. zum Löschen der Inputbox-Inhalte in der Registry
	;	HotKeySet("!d","ListAct")
	;	HotKeySet("{F5}","flneu"); machfilelist(1)
	GUISetState(@SW_MAXIMIZE)

	;	GUICtrlSetState($idDtn ,$GUI_SHOW)
	GUICtrlSetState($idDtn, $GUI_FOCUS) ; sonst funktionieren die Hotkeys nicht
	_GUICtrlListView_SetItemState($idDtn, 0, BitOR($LVIS_SELECTED, $LVIS_FOCUSED), BitOR($LVIS_SELECTED, $LVIS_FOCUSED))
	MouseMove(635, 675, 1) ; die Maus aufräumen, damit sie nicht bei F4-Auswahlen stört
EndFunc   ;==>zeichneHaupt

; aufgerufen in tuBenenn (3x)
Func Umbenenn($von, $nach, $wo)
	Local $qlw, $qvz, $lw, $vz, $rumpf, $ext, $ergd, $ergges
	_PathSplit($wo, $qlw, $qvz, $rumpf, $ext)
	_PathSplit($nach, $lw, $vz, $rumpf, $ext)
	$ergd = $rumpf & $ext
	$ergges = $qlw & $qvz & $ergd
	Const $cmd = "ren """ & $wo & "\" & $von & """ """ & $ergd & """"
	ConsoleWrite("Benenne um in " & $wo & ": " & $von & " -> " & $ergd & @LF)
	;	ConsoleWrite("cmd: " & $cmd & @LF)
	;	RunWait(@ComSpec & " /c chcp 1252 & cmd /c " & $cmd,$wo)
	RunWait(@ComSpec & " /c " & $cmd) ; ,$wo)
	For $i = 0 To 300
		If FileExists($ergges) Then Return (@error = 0)
		Sleep(10)
	Next ; $i=0
	Return 0
EndFunc   ;==>Umbenenn

; aufgerufen in Haupt (2x)
Func doUmwfSQL($str)
	If StringInStr($str, """") Then $str = StringReplace($str, """", "\""")
	If StringInStr($str, "\") Then $str = StringReplace($str, "\", "\\")
	If StringInStr($str, "'") Then $str = StringReplace($str, "'", "\'")
	Return $str
EndFunc   ;==>doUmwfSQL

; beinahe aufgerufen in Verarbeit
Func _GetHwndFromPID($procID)
	Local $hWnd = 0
	Local $stPID = DllStructCreate("int")
	Do
		Local $winlist2 = WinList()
		For $i = 1 To $winlist2[0][0]
			If $winlist2[$i][0] <> "" Then
				DllCall("user32.dll", "int", "GetWindowThreadProcessId", "hwnd", $winlist2[$i][1], "ptr", DllStructGetPtr($stPID))
				If DllStructGetData($stPID, 1) = $procID Then
					$hWnd = $winlist2[$i][1]
					ExitLoop
				EndIf
			EndIf
		Next
		ConsoleWrite("hWnd in GetHwndFromPID: " & $hWnd & @LF)
		If $hWnd Then ExitLoop
		Sleep(100)
	Until 0
	Return $hWnd
EndFunc   ;==>_GetHwndFromPID

; aufgerufen in Haupt
Func FragProt(ByRef $sql)
	If MyProtAnf($ProtCn) Then
		If _MySQL_Real_Query($ProtCn, $sql) Then ; =$MYSQL_ERROR Then
			MsgBox($MB_ICONERROR, "MariaDB-Fehler bei _MySQL_Real_Query mit " & $sql, _MySQL_Error($ProtCn))
			Return 0
		EndIf ; _MySQL_Real_Query
		_MySQL_Close($ProtCn)
		$ProtCn = 0
		Return 1
	EndIf ; MyProtAnf($ProtCn) Then
	Return 0
EndFunc   ;==>FragProt


; aufgerufen in fuellPatWahl, fuellWahl
Func MyPatAnf(ByRef $PatCn, $obneu = 0)
	Return MyVbd("Pat", $PatCn, $cSvr[$cRd], $cUsr[$cRd], $cPwd[$cRd], $cDb[$cRd], $cPrt[$cRd], $obneu)
EndFunc   ;==>MyPatAnf

; aufgerufen in Haupt
Func MyProtAnf(ByRef $ProtCn, $obneu = 0)
	Return MyVbd("Prot", $ProtCn, $cSvr[$Vbzl - 1], $cUsr[$Vbzl - 1], $cPwd[$Vbzl - 1], $cDb[$Vbzl - 1], $cPrt[$Vbzl - 1], $obneu)
EndFunc   ;==>MyProtAnf

; aufgerufen in MyPatAnf, MyProtAnf
Func MyVbd($uNm, ByRef $MySQL_ptr, $Host, $User, $Pass, $Db, $Port, $obneu)
	Static $anzufangen = 1
	If $obneu Then $anzufangen = 1
	If $anzufangen Then
		_MySQL_InitLibrary()
		$PatCn = 0
		$ProtCn = 0
		If @error Then
			MsgBox($MB_ICONERROR, $uNm & ", Fehler in My" & $uNm & "Anf bei _MySQL_InitLibrary", "Error: " & @error)
		Else
			$anzufangen = 0
		EndIf ; @error Then Else
		$MySQL_ptr = 0
	EndIf ; $anzufangen
	ConsoleWrite($uNm & ", MySQL_ptr: " & $MySQL_ptr & @LF)
	If $MySQL_ptr = 0 Then
		$MySQL_ptr = _MySQL_Init()
		If @error Then
			Local $errno = _MySQL_errno($MySQL_ptr)
			MsgBox($MB_ICONERROR, $uNm & ", Fehler in My" & $uNm & "Anf bei _MySQL_Init", "Errno: " & $errno & @LF & _MySQL_error($MySQL_ptr))
			Return 0
		EndIf ; @error=0
	EndIf ; $MySQL_ptr=0 Then
	; Func _MySQL_Real_Connect($MySQL_ptr, $Host, $User, $Pass, $Database = "", $Port = 0, $unix_socket = "", $Client_Flag = 0)
	If $Pass = "" Then $Pass = InputBox($uNm & ", Rückfrage", "Bitte Passwort für Benutzer '" & $User & "' auf Host '" & $Host & "' für Datenbank '" & $Db & "' eingeben.")
	ConsoleWrite($uNm & ", MySQL_ptr: " & $MySQL_ptr & ", Host: " & $Host & ", User: " & $User & ", Pass: " & $Pass & ", Db: " & $Db & ", Port: " & $Port & @LF)
	Local $verbunden = _MySQL_Real_Connect($MySQL_ptr, $Host, $User, $Pass, $Db, $Port)
	If $verbunden Then
		Return $verbunden
	Else
		Local $errno = _MySQL_errno($MySQL_ptr)
		MsgBox($MB_ICONERROR, $uNm & ", Fehler in My" & $uNm & "Anf bei _MySQL_Real_Connect", "Errno: " & $errno & @LF & _MySQL_error($MySQL_ptr))
	EndIf ; $verbunden = 0 Then
	Return 0
EndFunc   ;==>MyVbd

; aufgerufen in Haupt
Func fuellPatWahl() ; für $idPatWahl
	Local $res, $fields, $rows, $mysqlrow, $lenthsStruct, $length, $fieldPtr ; ,$row1,$field
	Static $altSQL
	;	_GUICtrlComboBox_ResetContent($idPatWahl)
	GUICtrlSetData($idPatWahl, '')
	If $cSQL[$cRd] <> $altSQL Then
		If MyPatAnf($PatCn) Then
			;  	  ConsoleWrite("SQL Pat: " & $cSQL[$cRd] & @LF)
			_MySQL_Real_Query($PatCn,"SET SESSION group_concat_max_len = 10000000;")
			If _MySQL_Real_Query($PatCn, $cSQL[$cRd]) = 0 Then
				_MySQL_Data_Seek($res, 0) ; nur zum zum Zurücksetzen an den Anfang der Abfrage
				$res = _MySQL_Store_Result($PatCn)
				If $res Then
					$fields = _MySQL_Num_Fields($res)
					$rows = _MySQL_Num_Rows($res)
					;     MsgBox(0, "", $rows & "-" & $fields)
					;     dim $array[$rows][$fields]
					For $k = 1 To $rows
						$mysqlrow = _MySQL_Fetch_Row($res, $fields)
						$lenthsStruct = _MySQL_Fetch_Lengths($res)
						For $i = 1 To $fields
							$length = DllStructGetData($lenthsStruct, 1, $i)
							$fieldPtr = DllStructGetData($mysqlrow, 1, $i)
							$PData = DllStructGetData(DllStructCreate("char[" & $length & "]", $fieldPtr), 1)
							;		   ConsoleWrite("data: " & $PData & @LF)
							;           GUICtrlSetData($idPatWahl,$PData)
							;		 $array[$k - 1][$i - 1] = $PData
						Next ; $i = 1 To $fields
					Next ; $k = 1 To $rows
					_MySQL_Free_Result($res)
				EndIf ; $res Then
				$altSQL = $cSQL[$cRd]
			EndIf ; _MySQL_Real_Query($PatCn, $cSQL[$cRd])=0 Then
			_MySQL_Close($PatCn)
			$PatCn = 0
		EndIf ; MyPatAnf($PatCn) Then
	Else ; $cSQL[$cRd]<>$altSQL Then
		;	  ConsoleWrite("cRd: " & $cRd & ", cSQL[$cRd]: " & $cSQL[$cRd] & ", " & @lf & "altsql: " & $altSQL & @LF)
	EndIf ; $cSQL[$cRd]<>$altSQL Else
;	ConsoleWrite("data: " & $PData & @LF)
	GUICtrlSetData($idPatWahl, $PData)
EndFunc   ;==>fuellPatWahl

; aufgerufen in ComZu, Haupt
Func fuellMA() ; für $idVA
	Local $res, $fields, $rows, $mysqlrow, $lenthsStruct, $length, $fieldPtr, $data ; ,$row1,$field
	Static $altSQL
	If $cMA[$cRd] <> $altSQL Then
		_GUICtrlComboBox_ResetContent($idVA)
		GUICtrlSetData($idVA, '')
		If MyPatAnf($PatCn) Then
			;	ConsoleWrite("SQL MA: " & $cMA[$cRd] & @lf)
			$res = _MySQL_Real_Query($PatCn, $cMA[$cRd])
			_MySQL_Data_Seek($res, 0) ; nur zum zum Zurücksetzen an den Anfang der Abfrage
			$res = _MySQL_Store_Result($PatCn)
			If $res Then
				$fields = _MySQL_Num_Fields($res)
				$rows = _MySQL_Num_Rows($res)
				;     MsgBox(0, "", $rows & "-" & $fields)
				;     dim $array[$rows][$fields]
				For $k = 1 To $rows
					$mysqlrow = _MySQL_Fetch_Row($res, $fields)
					$lenthsStruct = _MySQL_Fetch_Lengths($res)
					For $i = 1 To $fields
						$length = DllStructGetData($lenthsStruct, 1, $i)
						$fieldPtr = DllStructGetData($mysqlrow, 1, $i)
						$data = DllStructGetData(DllStructCreate("char[" & $length & "]", $fieldPtr), 1)
						;		 ConsoleWrite("data: " & $data & @lf)
						GUICtrlSetData($idVA, $data)
						;		 $array[$k - 1][$i - 1] = $data
					Next ; $i = 1 To $fields
				Next ; $k = 1 To $rows
				_MySQL_Free_Result($res)
			EndIf ; $res Then
			$altSQL = $cMA[$cRd]
			_MySQL_Close($PatCn)
			$PatCn = 0
		EndIf ; MyPatAnf($PatCn) then
	EndIf ; $cSQL[$cRd]<>$altSQL
EndFunc   ;==>fuellMA

; aufgerufen in WM_COMMAND
Func OptName($Nr)
	; Namen zu koordinieren mit LiesReg
	For $i = 0 To $Vbzl - 1
		If $i < $Vbzl - 1 Then
			Switch $Nr
				Case $iRd[$i]
					Return "Rd" & $i
				Case $iSQL[$i]
					Return "SQL" & $i
				Case $iMA[$i]
					Return "MA" & $i
			EndSwitch ; $Nr
		EndIf ; $i<$Vbzl-1 Then
		Switch $Nr
			Case $iSvr[$i]
				Return "Svr" & $i
			Case $iUsr[$i]
				Return "Usr" & $i
			Case $iPwd[$i]
				Return "Pwd" & $i
			Case $iDb[$i]
				Return "Db" & $i
			Case $iPrt[$i]
				Return "Prt" & $i
		EndSwitch ; $Nr
	Next ; $i
	Return ""
EndFunc   ;==>OptName

; aufgerufen in Haupt
Func LiesReg()
	While 1
		Local $regErr = 0
		;		ConsoleWrite("Anfang regErr: " & $regErr & @LF)
		$obPDF = RegRead($RegPath, "obPDF")
		$obWd = RegRead($RegPath, "obWord")
		$obEx = RegRead($RegPath, "obExcel")
		$oburspNVgb = RegRead($RegPath, "oburspNVgb")
		If $oburspNVgb = "" Then $oburspNVgb = 0
		$pSystem = RegRead($RegPath, "pSystem") ; PraxisSystem
		If @error > 0 Then $regErr = $regErr + @error
		;		ConsoleWrite("pSystem: " & $pSystem & ", 2 regErr: " & $regErr & @LF)
		; Namen zu koordinieren mit OptName
		$cRd = RegRead($RegPath, "VbdgRadio")
		If @error > 0 Then $regErr = $regErr + @error
		For $i = 0 To $Vbzl - 1
			$cSvr[$i] = RegRead($RegPath, "Svr" & $i)
			If $i = 0 And @error > 0 Then $regErr = $regErr + @error
			$cUsr[$i] = RegRead($RegPath, "Usr" & $i)
			If $i = 0 And @error > 0 Then $regErr = $regErr + @error
			$cPwd[$i] = RegRead($RegPath, "Pwd" & $i)
			If $i = 0 And @error > 0 Then $regErr = $regErr + @error
			$cDb[$i] = RegRead($RegPath, "Db" & $i)
			If $cDb[$i] = "" Then $cDb[$i] = "medoff"
			$cPrt[$i] = RegRead($RegPath, "Prt" & $i)
			If $cPrt[$i] = "" Then $cPrt[$i] = "2020"
			$cSQL[$i] = RegRead($RegPath, "SQL" & $i)
			If $cSQL[$i] = "" Then
				If $cDb[$i] = "medoff" Then
					$cSQL[$i] = "SELECT GROUP_CONCAT(CONCAT(FNachname,', ',FVorname,' (',COUNT(0) OVER(PARTITION BY FNachname,FVorname),'), *',DATE_FORMAT(FGeburtsdatum,'%d.%m.%Y'),', ',FOrt,', Pat_ID ',FSurogat) ORDER BY fnachname,fvorname SEPARATOR '|') FROM patstamm;"
				Else ; $cDb[$i]="medoff" Then
					$cSQL[$i] = "SELECT GROUP_CONCAT(CONCAT(nachname, ', ',vorname,', *',DATE_FORMAT(gebdat,'%d.%m.%Y'),', ',ort,', Pat_id ',pat_id) ORDER BY Nachname,Vorname SEPARATOR '|') FROM namen;"
				EndIf ; $cDb[$i]="medoff" Then Else
			EndIf ; $cSQL[$i]="" Then
			$cMA[$i] = RegRead($RegPath, "MA" & $i)
			If $cMA[$i] = "" Then
				If $cDb[$i] = "medoff" Then
					$cMA[$i] = "SELECT GROUP_CONCAT(CONCAT(FInitialen, ', ',FUsername) ORDER BY FInitialen SEPARATOR '|') FROM nutzerneu WHERE FInitialen<>'';"
				Else ; $cDb[$i]="medoff" Then
					$cMA[$i] = "SELECT GROUP_CONCAT(CONCAT(kuerzel,', ',Vorname,' ',Nachname) ORDER BY Kuerzel SEPARATOR '|') FROM dp.mitarbeiter WHERE aus=0 OR aus IS NULL;"
				EndIf ; $cDb[$i]="medoff" Then Else
			EndIf ; $cMA[$i]="" Then
		Next ; $i
		$BegZ = RegRead($RegPath, "BegZ")
		If @error > 0 Then $regErr = $regErr + @error
		If $BegZ <> "" Then
			ReDim $Begr[$BegZ]
			For $i = 0 To $BegZ - 1
				$Begr[$i] = RegRead($RegPath, "Begr" & $i)
				If $i = 0 And @error > 0 Then $regErr = $regErr + @error
			Next ; $i = 0 To $BegZ -1
		EndIf ; $Begz<>"" Then
		$Reit = RegRead($RegPath, "Reiter")
		If Not $Reit Then $Reit = "4"
		$Itv = RegRead($RegPath, "Intervall")
		If Not $Itv Then $Itv = 200
		$KopIn = RegRead($RegPath, "KopIn")
		If $i = 0 And @error > 0 Then $regErr = $regErr + @error
		If $cSvr[0] = "" Or $cUsr[0] = "" Then $regErr = $regErr + 1
		$aktSort = RegRead($RegPath, "Sort")
		If $regErr Then
			$optaktiv = 2
			Optionen()
			Schleife()
		Else ; $regErr
			ExitLoop
		EndIf ; $regErr Else
	WEnd ; While 1
; Passagere Festlegung 17.3.25
	$pSystem=0
	$cSvr[0]="wser"
	$cRd=0
    $cSQL[0]="SELECT GROUP_CONCAT(el SEPARATOR '|') FROM (SELECT CONCAT(FNachname,', ',FVorname,' (',COUNT(0) OVER(PARTITION BY FNachname,FVorname),'), *',DATE_FORMAT(FGeburtsdatum,'%d.%m.%Y'),', ',FOrt,', Pat_ID ',FSurogat)el FROM patstamm ORDER BY FNachname,FVorname)i;"
	$cSQL[1]= "SELECT GROUP_CONCAT(el SEPARATOR '|') FROM (SELECT CONCAT(nachname, ', ',vorname,' (',COUNT(0) OVER(PARTITION BY nachname,vorname),'), *',DATE_FORMAT(gebdat,'%d.%m.%Y'),', ',ort,', Pat_id ',pat_id)el FROM namen ORDER BY Nachname,Vorname)i;"
;    $cSQL[0]="SELECT GROUP_CONCAT(CONCAT(FNachname,', ',FVorname,' , *',DATE_FORMAT(FGeburtsdatum,'%d.%m.%Y'),', ',FOrt,', Pat_ID ',FSurogat) ORDER BY fnachname,fvorname SEPARATOR '|') FROM patstamm;"
;	$cSQL[1]= "SELECT GROUP_CONCAT(CONCAT(nachname, ', ',vorname,', *',DATE_FORMAT(gebdat,'%d.%m.%Y'),', ',ort,', Pat_id ',pat_id) ORDER BY Nachname,Vorname SEPARATOR '|') FROM namen;"
EndFunc   ;==>LiesReg

; aufgerufen in Haupt()
Func Optionen()
	; Dim $iNr[$Vbzl][6], $iSvr[$Vbzl], $iUsr[$Vbzl], $iPwd[$Vbzl], $iDb[$Vbzl], $iPrt[$Vbzl],$iSQL[$Vbzl]
	$idOpt = GUICreate("Optionen", 1620, 750, 0, 100, BitOR($WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_SYSMENU, $WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_EX_CONTROLPARENT))
	$idCbPDF = GUICtrlCreateCheckbox("&PDF-Vorschau", 10, 10, 100, 20)
	If $obPDF <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idCbWd = GUICtrlCreateCheckbox("&Word-Vorschau", 110, 10, 100, 20)
	If $obWd <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idCbEx = GUICtrlCreateCheckbox("&Excel-Vorschau", 210, 10, 100, 20)
	If $obEx <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idMO = GUICtrlCreateRadio("&für Medical Office?", 600, 10)
	$idTM = GUICtrlCreateRadio("für Tur&bomed?", 720, 10)
	GUICtrlCreateLabel("Re&iter:", 810, 14)
	$idRei = GUICtrlCreateInput($Reit, 845, 10, 15)
	GUICtrlCreateLabel("Intervall[ms]&:", 870, 14)
	$idItv = GUICtrlCreateInput($Itv, 930, 10, 30)
	Switch $pSystem
		Case 0
			GUICtrlSetState($idMO, $GUI_CHECKED)
		Case 1
			GUICtrlSetState($idTM, $GUI_CHECKED)
	EndSwitch ; $pSystem
	$idFertig = GUICtrlCreateButton("&Ok", 911, 690, 60, 30)
	GUIStartGroup()
	For $i = 0 To $Vbzl - 1
		If $i < $Vbzl - 1 Then
			$iRd[$i] = GUICtrlCreateRadio("&Vbdg.", 2, 16 + 25 * ($i + 1), 35, 23)
			If $i = $cRd Then GUICtrlSetState($iRd[$i], $GUI_CHECKED)
		EndIf ; $i<$Vbzl-1 Then
		;	Next; $i
		;	For $i= 0 to $Vbzl-1
		$iNr[$i][0] = GUICtrlCreateLabel(($i = $Vbzl - 1 ? "&Protok'vbg. " : "" & $i & ":") & " &Server:", ($i = $Vbzl - 1 ? 0 : 40), 21 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), ($i = $Vbzl - 1 ? 100 : 80), 23)
		$iSvr[$i] = GUICtrlCreateInput($cSvr[$i], 90, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 50, 23)
		$iNr[$i][1] = GUICtrlCreateLabel("&User:", 147, 21 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 29, 23)
		$iUsr[$i] = GUICtrlCreateInput($cUsr[$i], 177, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 50, 23)
		$iNr[$i][2] = GUICtrlCreateLabel("P&asswort:", 233, 21 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 49, 23)
		$iPwd[$i] = GUICtrlCreateInput($cPwd[$i], 281, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 49, 23)
		$iNr[$i][3] = GUICtrlCreateLabel("Date&nbank:", 334, 21 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 55, 23)
		$iDb[$i] = GUICtrlCreateInput($cDb[$i], 392, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 69, 23)
		$iNr[$i][4] = GUICtrlCreateLabel("Por&t:", 465, 21 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 24, 23)
		$iPrt[$i] = GUICtrlCreateInput($cPrt[$i], 490, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 39, 23)
		If $i < $Vbzl - 1 Then
			$iNr[$i][5] = GUICtrlCreateLabel("S&QL-Pat:", 533, 21 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 45, 23)
			$iSQL[$i] = GUICtrlCreateInput($cSQL[$i], 579, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 454, 23)
			$iNr[$i][6] = GUICtrlCreateLabel("SQL-&MA:", 1038, 21 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 42, 23)
			$iMA[$i] = GUICtrlCreateInput($cMA[$i], 1084, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 185, 23)
		Else
			$iKopAusw = GUICtrlCreateButton("&KopVz ohne Jahreszahlen:", 550, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 131, 23)
			$iKopIn = GUICtrlCreateInput($KopIn, 685, 17 + 25 * ($i + 1 + ($i = $Vbzl - 1 ? 1 : 0)), 584, 23)
		EndIf ; $i<$Vbzl-1 Then
	Next ; $i
	GUICtrlCreateLabel("Vo&rgabetext für Dokumentenkategorie:", 5, 25 + 25 * ($Vbzl + 2.5))
	$iEing = GUICtrlCreateInput("", 190, 21 + 25 * ($Vbzl + 2.5), 205)
	$iDokK = GUICtrlCreateListView("Begriff", 400, 21 + 25 * ($Vbzl + 2.5), 400, 500)
	GUICtrlSetBkColor(-1, 0xDDEEFF) ; Set the background color for the listview
	_ArraySort($Begr)
	For $i = 0 To $BegZ - 1
		GUICtrlCreateListViewItem($Begr[$i], $iDokK)
	Next ; $i = 0 To $BegZ -1

	_GUICtrlListView_SetColumnWidth($iDokK, 0, 500)
	$idnR = GUICtrlCreateButton("&>", 336, 25 + 25 * ($Vbzl + 3.5), 60, 30)
	$idnL = GUICtrlCreateButton("&<", 336, 25 + 25 * ($Vbzl + 4.5), 60, 30)
	$idurspNVgb = GUICtrlCreateCheckbox("urspr&üngl. Namen meist anhängen", 5, 65 + 25 * ($Vbzl + 2.5))
	If $oburspNVgb <> 0 Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idkPatVgb = GUICtrlCreateCheckbox("meist keinen Patientennamen einset&zen", 5, 85 + 25 * ($Vbzl + 2.5))
	If $obkPatVgb <> 0 Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idntUVgb = GUICtrlCreateCheckbox("&meist nicht umbenennen", 5, 105 + 25 * ($Vbzl + 2.5))
	If $obntUVgb <> 0 Then GUICtrlSetState(-1, $GUI_CHECKED)
	$idnimVgb = GUICtrlCreateCheckbox("meist nicht impo&rtieren", 5, 125 + 25 * ($Vbzl + 2.5))
	If $obnimVgb <> 0 Then GUICtrlSetState(-1, $GUI_CHECKED)

	$idokakt = GUICtrlCreateDummy()
	GUICtrlCreateLabel("&Dateisortierung:", 820, 25 + 25 * ($Vbzl + 2.5))
	; Const $clArr=["Namen aufsteigend","Namen absteigend","Datum aufsteigend","Datum absteigend","Größe aufsteigend","Größe absteigend","Typ aufsteigend","Typ absteigend"]
	$iSort = GUICtrlCreateCombo("", 900, 21 + 25 * ($Vbzl + 2.5), 200, 450, $CBS_DROPDOWNLIST) ; "Sortierung                         "
	GUICtrlSetData($iSort, _ArrayToString($clArr), $clArr[$aktSort])

	Local $aAccelKeys[1][2] = [["!l", $idokakt]]
	GUISetAccelerators($aAccelKeys)
	;	HotKeySet("!l","iDokAct")
	GUISetState() ; = @SW_SHOW,$idOpt)
EndFunc   ;==>Optionen

; aufgerufen in WM_NOTIFY, Haupt (Hotkey)
Func ListAct() ; für $idDtn
	;   Local $hHaupt = IsHWnd($idHaupt)?$idHaupt:GUICtrlGetHandle($idHaupt);
	;	Local $ctlref = ControlGetFocus($hHaupt)
	;	$ctlref = ControlGetFocus($idHaupt)
	;    Local $hwnd = ControlGetHandle($idHaupt, "", ControlGetFocus($idHaupt))
	;    Local $id = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hwnd)
	;    Local $result = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hwnd)
	Static $lCtlid
	;	ConsoleWrite("setze fokus auf listview" & @LF)
	; ggf das vorher markierte Element wieder markieren
	Local $itnr = _GUICtrlListView_GetSelectedIndices($idDtn) ; GUICtrlRead($idDtn)
	;	ConsoleWrite("ausgewähltes Item: " & $itnr & @LF)
	If $itnr = "" Then
		_GUICtrlListView_SetItemSelected($idDtn, $lgwFl)
		;		_GUICtrlListView_ClickItem($idDtn, 0)
	EndIf; $itnr = ""
	If ControlGetFocus($idHaupt) <> $lctlid Then
		GUICtrlSetState($idDtn, $GUI_FOCUS) ; wird nach PDF-Befüllung hier aufgerufen
		;	 ConsoleWrite("ausgewähltes Item: " & GUICtrlRead($idDtn) & @lf)
		$lctlid = ControlGetFocus($idHaupt)
	EndIf ; ControlGetFocus($idHaupt)<>$lctlid Then
EndFunc   ;==>ListAct

; aufgerufen in HotKeySet
Func iDokAct()
	GUICtrlSetState($iDokK, $GUI_FOCUS)
EndFunc   ;==>iDokAct

; aufgerufen in Haupt (Hotkey)
Func PatAct()
	GUICtrlSetState($idPatWahl, $GUI_FOCUS)
EndFunc   ;==>PatAct

; aufgerufen auskommentiert in Haupt (Hotkey)
Func InhAct()
	GUICtrlSetState($iILb[0], $GUI_FOCUS)
EndFunc   ;==>InhAct

; aufgerufen nirgends
Func _GuiCtrlGetFocus($GuiRef)
	Local $hWnd = ControlGetHandle($GuiRef, "", ControlGetFocus($GuiRef))
	;    Return $result[0]
	;	return $hWnd
	Return ControlGetFocus($GuiRef)
EndFunc   ;==>_GuiCtrlGetFocus

; aufgerufen in Haupt
Func machPdf()
	$oPDF = ObjCreate("AcroPDF.PDF.1") ;
	If IsObj($oPDF) Then
		; 		$oPDF = ObjCreate("Preview.Preview.1");
		; 		Local $Version=$oPDF.GetVersions
		$oPDF.src = "gibtsnicht.pdf" ;PDF FILE 1
		; 		GuiCreate("PDF Object", 802, 590,(@DesktopWidth-802)/2, (@DesktopHeight-590)/2 , $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS)
		$GUI_PDF = GUICtrlCreateObj($oPDF, 1, 0, $aML - 5, $aDUR)
		; 		GUICtrlSetStyle ( $GUI_PDF,  $WS_VISIBLE )
		GUICtrlSetResizing($GUI_PDF, $GUI_DOCKAUTO)    ; Auto Resize Object
		GUICtrlSetState($GUI_PDF, $GUI_HIDE)
	EndIf ; IsObj($oPDF) Then
EndFunc   ;==>machPdf

; aufgerufen in Haupt (2x), WM_NOTIFY
Func machExcel()
	;	$oExcel = ObjCreate("Excel.Application") ; Create an Excel Object
	$oExcel = _Excel_Open(False) ; false entfernt 29.12.24
	If IsObj($oExcel) Then
		;		$oExcel.WindowStyle = 0
		;		$oExcel.Visible = 1 ; Let Excel show itself
		; 		$oExcel.open "gibtsnicht.xls"
		$GUI_Excel = GUICtrlCreateObj($oExcel, 1, 0, $aML - 5, $aDUR)
		GUICtrlSetStyle($GUI_Excel, $WS_VISIBLE)
		GUICtrlSetResizing($GUI_Excel, $GUI_DOCKAUTO)     ; Auto Resize Object
		GUICtrlSetState($GUI_Excel, $GUI_HIDE)
	Else
		ConsoleWrite("oExcel konnte nicht erstellt werden " & @LF)
	EndIf ; IsObj($oExcel) Then
EndFunc   ;==>machExcel

; aufgerufen in Haupt (2x), WM_NOTIFY
Func machWord()
	;	$oWord = ObjCreate("Word.Application") ; Create a Word Object
	$oWord = _Word_Create(False)
	If IsObj($oWord) Then
		;		$oWord.Visible = 1 ; Let Word show itself
		;		$oWord.WindowStyle = 0 ; wdWindowStateNormal 1 = wdWindowStateMaximize
		;		$oWord.Documents.open "gibtsnicht.xls"
		$GUI_Word = GUICtrlCreateObj($oWord, 1, 0, $aML - 5, $aDUR)
		GUICtrlSetStyle($GUI_Word, $WS_VISIBLE)
		GUICtrlSetResizing($GUI_Word, $GUI_DOCKAUTO)    ; Auto Resize Object
		GUICtrlSetState($GUI_Word, $GUI_HIDE)
	Else
		ConsoleWrite("oWord konnte nicht erstellt werden " & @LF)
	EndIf ; IsObj($oWord) Then
EndFunc   ;==>machWord

; aufgerufen in Haupt, machfilelist
Func machflstr()
	Local $zwdat = EnvGet("temp") & "\get2.txt"
	; Const $clArr=["Namen aufsteigend","Namen absteigend","Datum aufsteigend","Datum absteigend","Größe aufsteigend","Größe absteigend","Typ aufsteigend","Typ absteigend"]
	Local $Sortue
	Switch $aktSort
		Case 0
			$Sortue = "N"
		Case 1
			$Sortue = "-N"
		Case 2
			$Sortue = "D"
		Case 3
			$Sortue = "-D"
		Case 4
			$Sortue = "S"
		Case 5
			$Sortue = "-S"
		Case 6
			$Sortue = "E"
		Case 7
			$Sortue = "-E"
	EndSwitch ; $aktSort
	ShellExecuteWait("cmd.exe", "/c ""chcp 1252 & cmd /c dir " & $Pfad & " /a-d /O" & $Sortue & " >" & $zwdat & """", "", "", @SW_HIDE)
	Local $hFO = FileOpen($zwdat)
	Local $sFileRead = FileRead($hFO)
	;	ConsoleWrite("sFileRead: " & $sFileRead & @LF)
	$aFL = StringSplit($sFileRead, @LF)
	$aUb = UBound($aFL)
	ReDim $aDL[$aUb]
	ReDim $typ[$aUb]
	ReDim $itm[$aUb]
	ReDim $dgroe[$aUb]
	ReDim $dlaend[$aUb]
	Dim $j = 0
	For $i = 1 To $aUb - 1
		Local $sl = StringLeft($aFL[$i], 1)
		;	 ConsoleWrite("i: " & $i & ", aFL[i]: " & $aFL[$i] & ", s1: " & $sl & ";" & @LF)
		If $sl <> "" And $sl <> " " And $sl <> Chr(13) Then
			$dgroe[$j] = StringMid($aFL[$i], 18, 18)
			Local $goff = 0
			While StringLeft($dgroe[$j], $goff + 1) = " "
				$goff = $goff + 1
			WEnd
			If $goff <> 0 Then $dgroe[$j] = StringMid($dgroe[$j], $goff)
			Local $Dnam = StringTrimRight(StringMid($aFL[$i], 37), 1)
			If StringInStr($Dnam, "Schutzdatei_bitte_belassen") = 0 Then
				$typ[$j] = StringInStr($Dnam, ".", 0, -1) ? StringLower(StringRight($Dnam, StringLen($Dnam) - StringInStr($Dnam, ".", 0, -1))) : ""
				$dlaend[$j] = StringMid($aFL[$i], 7, 4) & "." & StringMid($aFL[$i], 4, 2) & "." & StringLeft($aFL[$i], 2) & " " & StringMid($aFL[$i], 13, 2) & ":" & StringMid($aFL[$i], 16, 2)
				$itm[$j] = StringLeft($Dnam & $strich, 500) & "|" & $dlaend[$j] & "|" & $dgroe[$j] & "|" & $typ[$j] & "|" & $j
				;		ConsoleWrite("Erstelle Item: " & $j & ": " & $itm[$j] & @LF)
				$aDL[$j] = $Dnam
				$j = $j + 1
			EndIf ; StringInStr($Dnam,"Schutzdatei_bitte_belassen")=0 Then
			;		ConsoleWrite("j: " & $j & ", aDL: " & $aDL[$j] & ", i: " & $i & ", aFL[" & $i & "]: " & $aFL[$i] & @lf)
		EndIf ; $sl<>"" And $sl<>" " And $sl<>chr(13) Then
	Next ; $i=1 To $aUb-1
	If $j > 0 Then
		ReDim $aDL[$j]
		ReDim $typ[$j]
		ReDim $itm[$j]
		ReDim $dgroe[$j]
		ReDim $dlaend[$j]
	EndIf ; $j>0 Then
EndFunc   ;==>machflstr

; aufgerufen in Verarbeit, Schleife, Haupt, waehlepfad
Func machfilelist($obflstr = 0) ; für $idDtn
	Local $hStarttime = _Timer_Init()
	If $idDtn <> 0 Then GUICtrlDelete($idDtn)
	$idDtn = GUICtrlCreateListView("Dateien '" & $muster & "' in '" & $Pfad & "' (Alt+D = fokussieren, F5 = auffrischen)|geändert|Größe|Typ|Index", $aML - ($aML * 0.333), 4, $aML + ($aML * 0.333) - 18, $aDUR - 10, $LVS_REPORT) ; ,$LVS_SORTDESCENDING)
	GUICtrlSetState(-1, $GUI_HIDE)    ; zum schnelleren Bildaufbau
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_LV_ALTERNATE)     ; alternate between the listview background color and the listview item background color
	GUICtrlSetBkColor(-1, 0xDDEEFF)     ; Set the background color for the listview
	Local Const $methode = 1
	Switch $methode
		Case 1     ; Schnellstes: 4,7 s bei 10000 Dateien
			If $obflstr Then machflstr()
			Local $idIt[$aUb]
			For $j = 0 To UBound($itm) - 1
				If $itm[$j] <> "" Then
					;						ConsoleWrite("j: " & $j & ", itm[$j]: " & $itm[$j] & @LF)
					$idIt[$j] = GUICtrlCreateListViewItem($itm[$j], $idDtn)
					GUICtrlSetBkColor(-1, 0xFCF6EA)     ; Set the background color for the listview item
				EndIf    ; $itm[$j]<>"" Then
			Next    ; $j=0 To UBound($itm)-1
		Case 2     ; Zweitschnellstes: 6,4 s bei 10000 Dateien
			ShellExecuteWait("u:\programmierung\getfolder.exe", $Pfad)    ; ,"","",@SW_HIDE)
			Local $hFO = FileOpen(EnvGet("temp") & "\getFolderAusg.txt")
			Local $sFileRead = FileRead($hFO)
			$aFL = StringSplit($sFileRead, @LF)
			$aUb = UBound($aFL)
			Local $idIt[$aUb]
			For $i = 0 To $aUb - 1
				$idIt[$i] = GUICtrlCreateListViewItem($aFL[$i], $idDtn)
				$aFL[$i] = StringLeft(StringLeft($aFL[$i], 500), StringInStr($aFL[$i], " . .") - 1)
				;					ConsoleWrite($aFL[$i] & @lf)
				GUICtrlSetBkColor(-1, 0xFCF6EA)     ; Set the background color for the listview item
			Next
		Case 3     ; Drittschnellstes: 63 s bei 10000 Dateien
			$aFL = _FileListToArray($Pfad, $muster, $FLTA_FILES, False)
			ConsoleWrite("1 Brauchte: " & _Timer_Diff($hStarttime) & @LF)
			Const $FM = "_FileListToArray Fehler"
			If FileListFehler($FM, $Pfad, $muster) Then
				$ExitFnr = 17
				ComZu("Schleife machfilelist, Methode 3")
			EndIf
			$aUb = UBound($aFL)
			Local $elem
			Local $idIt[$aUb]
			For $i = 0 To $aUb - 1
				If True Then
					Local $gesPfd = $Pfad & (StringRight($Pfad, 1) = "\" ? "" : "\") & $aFL[$i]
					Local $groe = FileGetSize($gesPfd)
					Local $groelen = StringLen($groe)
					$groe = StringLeft("__________________", 12 - $groelen) & $groe
					$groe = StringLeft($groe, 3) & "." & StringMid($groe, 4, 3) & "." & StringMid($groe, 7, 3) & "." & StringMid($groe, 10, 3)
					Local $typ = StringRight($aFL[$i], StringLen($aFL[$i]) - StringInStr($aFL[$i], ".", 0, -1))
					Local $aend = FileGetTime($gesPfd, 0, 1)
					$elem = StringLeft($aFL[$i] & $strich, 500) & "|" & StringLeft($aend, 4) & "." & StringMid($aend, 5, 2) & "." & StringMid($aend, 7, 2) & " " & StringMid($aend, 9, 2) & ":" & StringMid($aend, 11, 2) & ":" & StringRight($aend, 2) & "|" & $groe & "|" & $typ
				Else     ; True
					$elem = $aFL[$i]
				EndIf     ; True Else
				$idIt[$i] = GUICtrlCreateListViewItem($elem, $idDtn)
				GUICtrlSetBkColor(-1, 0xFCF6EA)     ; Set the background color for the listview item
			Next     ; $i
		Case 4     ; Drittschnellstes: 600 s bei 10000 Dateien
			Local $oFSO = ObjCreate("Scripting.FileSystemObject")
			Local $oFolder = $oFSO.GetFolder($Pfad)
			Local $idIt[$oFolder.Files.count], $au[$oFolder.Files.count]
			Local $Dtn = $oFolder.Files
			Local $i = 0
			For $DatElem In $Dtn
				$au[$i] = $DatElem.Name
				Local $groe = $datelem.size
				Local $groelen = StringLen($groe)
				$groe = StringLeft("__________________", (12 - $groelen)) & $groe
				$groe = StringLeft($groe, 3) & "." & StringMid($groe, 4, 3) & "." & StringMid($groe, 7, 3) & "." & StringMid($groe, 10, 3)
				;					Local $typ=StringRight($aFL[$i],stringlen($aFL[$i])-stringinstr($aFL[$i],".",0,-1))
				Local $typ = $datelem.type
				Local $elem = StringLeft($datelem.name & $strich, 500) & "|" & StringLeft($datelem.DateLastModified, 4) & "." & StringMid($datelem.DateLastModified, 5, 2) & "." & StringMid($datelem.DateLastModified, 7, 2) & " " & StringMid($datelem.DateLastModified, 9, 2) & ":" & StringMid($datelem.DateLastModified, 11, 2) & ":" & StringRight($datelem.DateLastModified, 2) & "|" & $groe & "|" & $typ
				$idIt[$i] = GUICtrlCreateListViewItem($elem, $idDtn)
				GUICtrlSetBkColor(-1, 0xFCF6EA)     ; Set the background color for the listview item
				$i = $i + 1
			Next    ; $DatElem In $Dtn
			$aFL = $au
	EndSwitch
	_GUICtrlListView_SetColumnWidth($idDtn, 0, 590)   ; $LVSCW_AUTOSIZE); 480)
	GUICtrlSetState($idDtn, $GUI_SHOW)    ; zum schnelleren Bildaufbau
	ConsoleWrite("2 Brauchte: " & _Timer_Diff($hStarttime) & @LF)
EndFunc   ;==>machfilelist

; aufgerufen auskommentiert in Haupt (GUIRegisterMsg)
Func WM_SIZE($hWnd, $iMsg, $iWParam, $iLParam)
	#forceref $hWnd, $iMsg, $iWParam
	$Weite = BitAND($iLParam, 0xFFFF)
	Local $iGUIWidth1 = BitAND($iLParam, 0xFFFF)
	$Hoehe = BitShift($iLParam, 16)
	Local $iGUIWidth2 = WinGetPos($idHaupt)[2]
	Local $iGUIHeight2 = WinGetPos($idHaupt)[3]
	;   ConsoleWrite("GUIWidth : " & $Weite & "  " & $iGUIWidth2 & "  " & $Weite - $iGUIWidth2  & @LF )
	;    ConsoleWrite("GUIHeight: " & $Hoehe & "  " & $iGUIHeight2 & "  " & $Hoehe - $iGUIHeight2  & @LF  & @LF )
	Return ($GUI_RUNDEFMSG)
EndFunc   ;==>WM_SIZE

; aufgerufen auskommentiert in Haupt (GUIRegisterMsg)
Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam
	Local $hdlWindowFrom, _
			$intMessageCode, _
			$intControlID_From

	$intControlID_From = BitAND($wParam, 0xFFFF)  ; _WinAPI_HiWord(
	$intMessageCode = BitShift($wParam, 16) ; _WinAPI_LoWord(
	Local $inhalt = GUICtrlRead($intControlID_From)
	ConsoleWrite("Schreibe: hWnd: " & $hWnd & ", iMsg: " & $iMsg & ", wParam: " & $wParam & ", $intControlID_From: " & $intControlID_From & ", lParam: " & $lParam & ", Inhalt: " & $inhalt & ", intMessageCode: " & $intMessageCode & @LF)
	RegWrite($RegPath, OptName($intControlID_From), "REG_SZ", $inhalt)
	Switch $intMessageCode
		Case $EN_CHANGE ; $EN_UPDATE
			;			Local $inhalt = GUICtrlRead($intControlID_From)
			; Schreibe: hWnd: 0x003A17F4, iMsg: 273, wParam: 0x03000048, $intControlID_From: 72, Optname(): Usr0, lParam: 0x002E1CE0, Inhalt: 4
			; Schreibe: hWnd: 0x003A17F4, iMsg: 273, wParam: 0x03000048, $intControlID_From: 72, Optname(): Usr0, lParam: 0x002E1CE0, Inhalt: medof
			;           nicht trennbar, ob Radiobutton oder Inputbox betätigt wurde
			;			RegWrite($RegPath, Optname($intControlID_From),"REG_SZ", $inhalt)
			;        Switch $intControlID_From
			;          Case $iSvr[0]
			;                    ConsoleWrite("[" & _Now() & "] - The text in the $iSvr[0] control has changed! Text = " & GUICtrlRead($iSvr[0]) & @LF)
			;      EndSwitch ; $EN_CHANGE
	EndSwitch ; $intMessageCode
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func _GetEntries($idDtn)
	Local $aIndex = _GUICtrlListView_GetSelectedIndices($idDtn, True)
	If $aIndex[0] < 1 Then Return False
	Local $aData[UBound($aIndex) - 1] ; [6]
	$aData[0] = $aIndex[0] ; $aData[0][0]
	For $i = 1 To $aIndex[0]
		$aData[$i - 1] = _GUICtrlListView_GetItemText($idDtn, $aIndex[$i], 4) ; Index
	Next ; $i = 1 To $aIndex[0]
	Return $aData
EndFunc   ;==>_GetEntries

; aufgerufen in Haupt (GUIRegisterMsg)
Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam
	Static $hDC, $iItem, $hBrush = DllCall("gdi32.dll", "handle", "CreateSolidBrush", "int", _WinAPI_GetSysColor($COLOR_HIGHLIGHT))[0]   ; Selected back color, _WinAPI_CreateSolidBrush
	Static $tRect = DllStructCreate($tagRECT), $pRect = DllStructGetPtr($tRect), $tLVitem = DllStructCreate($tagLVITEM), $pLVitem = DllStructGetPtr($tLVitem), $tBuffer = DllStructCreate("wchar Text[4096]"), $pBuffer = DllStructGetPtr($tBuffer)
	;  Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)

	Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR, $hWndDtn, $obneupdf
	$tNMHDR = DllStructCreate("hwnd;uint_ptr;int_ptr;int;int", $lParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, 1))
	If $hWndFrom <> 0 Then
		$iCode = DllStructGetData($tNMHDR, 3)
		$hWndDtn = $idDtn
		If Not IsHWnd($idDtn) Then $hWndDtn = GUICtrlGetHandle($idDtn)
		Switch $hWndFrom
			Case $hWndDtn
				Switch $iCode
					Case $NM_CUSTOMDRAW
						Local $tNMLVCUSTOMDRAW = DllStructCreate($tagNMLVCUSTOMDRAW, $lParam)
						Switch DllStructGetData($tNMLVCUSTOMDRAW, "dwDrawStage")
							Case $CDDS_PREPAINT                          ; Before the paint cycle begins
								$hDC = DllStructGetData($tNMLVCUSTOMDRAW, "hdc")                          ; Device context
								DllCall("gdi32.dll", "int", "SetBkMode", "handle", $hDC, "int", $TRANSPARENT) ; Transparent background, _WinAPI_SetBkMode
								Return $CDRF_NOTIFYITEMDRAW              ; Notify the parent window before an item is painted

							Case $CDDS_ITEMPREPAINT                      ; Before an item is painted
								$iItem = DllStructGetData($tNMLVCUSTOMDRAW, "dwItemSpec") ; Item index
								If GUICtrlSendMsg($idDtn, $LVM_GETITEMSTATE, $iItem, $LVIS_SELECTED) Then _ ; Selected item?
										Return $CDRF_NOTIFYPOSTPAINT     ; Custom drawing of selected items
								Return $CDRF_NEWFONT                     ; Default drawing of other items

							Case $CDDS_ITEMPOSTPAINT                     ; After an item has been painted
								DllCall("gdi32.dll", "int", "SetTextColor", "handle", $hDC, "int", 0xFFFFFF) ; Selected fore color, white, _WinAPI_SetTextColor
								; For each subitem
								For $iSubItem = 0 To 4
									; Subitem rectangle
									DllStructSetData($tRect, "Top", $iSubItem)
									DllStructSetData($tRect, "Left", $LVIR_LABEL)
									GUICtrlSendMsg($idDtn, $LVM_GETSUBITEMRECT, $iItem, $pRect)
									DllStructSetData($tRect, "Left", DllStructGetData($tRect, "Left"))

									; Selected back color
									DllCall("user32.dll", "int", "FillRect", "handle", $hDC, "struct*", $tRect, "handle", $hBrush) ; _WinAPI_FillRect

									; Left and top margin of subitem text
									DllStructSetData($tRect, "Left", DllStructGetData($tRect, "Left") + ($iSubItem ? 6 : 2))
									DllStructSetData($tRect, "Top", DllStructGetData($tRect, "Top") + 2)

									; Extract subitem text from listview
									DllStructSetData($tLVitem, "Mask", $LVIF_TEXT)
									DllStructSetData($tLVitem, "SubItem", $iSubItem)
									DllStructSetData($tLVitem, "Text", $pBuffer)
									DllStructSetData($tLVitem, "TextMax", 4096)
									GUICtrlSendMsg($idDtn, $LVM_GETITEMTEXTW, $iItem, $pLVitem)

									; Draw subitem text with selected fore color and selected text font (through the device context, $hDC)
									DllCall("user32.dll", "int", "DrawTextW", "handle", $hDC, "wstr", DllStructGetData($tBuffer, "Text"), "int", -1, "struct*", $tRect, "uint", $DT_WORD_ELLIPSIS) ; _WinAPI_DrawText
								Next
								Return $CDRF_NEWFONT                     ; $CDRF_NEWFONT must be returned after changing font or colors
						EndSwitch ; Switch DllStructGetData( $tNMLVCUSTOMDRAW, "dwDrawStage" )
					Case $NM_DBLCLK, $NM_CLICK, -101 ; -3, -2
						Local $gwNr = GUICtrlRead($idDtn) ; Reihenfolgenummer des ggf. gewählten Datei-Eintrags
						If $gwNr Then
							ConsoleWrite("gwNr: " & $gwNr & @LF)
							Local $gwDa = GUICtrlRead($gwNr) ; Datei-Eintrag
							ConsoleWrite("gwDA: " & $gwDa & @LF)

							ControlSetText($idHaupt, "", $idPatWahl, "")
							Local $TzA[] = [" ","_",", ",","], $gefunden=0
							local $gEntr
							$gEntr = _GetEntries($idDtn)
							local $merkname
							$merkname=""
							For $iy = 0 to UBound($gEntr)-1
								consolewrite("$aDL[$gEntr[" & $iy & "]]: " & $aDL[$gEntr[$iy]] & @lf)
								For $Tz In $TzA
									Local $pos = StringInStr($aDL[$gEntr[$iy]], $Tz, 0, 2)
									If $pos Then
										Local $pna = StringSplit(StringLeft($aDL[$gEntr[$iy]], $pos), $Tz)
										;							ConsoleWrite(">>>>>>>>>> " & $pna[1] & ", " & $pna[2] & @LF)
										For $j = 0 To 1
											Local $suchstr=StringLeft($pna[1+$j],StringLen($pna[1+$j])-(StringRight($pna[1+$j],1)=","?1:0)) & ", " & StringLeft($pna[2-$j],StringLen($pna[2-$j])-(StringRight($pna[2-$j],1)=","?1:0)) & " (1)"
;                                            ConsoleWrite("Namensuche: " & $suchstr & @LF)
											If _GUICtrlComboBox_SelectString($idpatwahl,$suchstr) <> -1 Then
;												consolewrite("gefunden!!!!!!!!!" & @lf)
												if $merkname = "" then
													$merkname = GUICtrlRead($idpatwahl)
													$gefunden = 1
													ExitLoop
												ElseIf $merkname <> GUICtrlRead($idpatwahl) and GUICtrlRead($idpatwahl)<>"" Then
													ConsoleWrite("Merkname: " & $merkname & ", GUICtrlRead($idpatwahl): " & GUICtrlRead($idpatwahl) & @lf)
													MsgBox(64,"Warnung","Wohl Dateien verschiedener Patienten ausgewählt!")
													_GUICtrlListView_SetItemSelected($idDtn, -1, False)
													Return 0
												endif
											EndIf
										Next; $j
									EndIf; $pos
								Next ; $Tz
							Next; $iy
;							$gefunden=0
;							For $Tz In $TzA
;								Local $pos = StringInStr($gwDa, $Tz, 0, 2)
;								If $pos Then
;									Local $pna = StringSplit(StringLeft($gwDa, $pos), $Tz)
;									;							ConsoleWrite(">>>>>>>>>> " & $pna[1] & ", " & $pna[2] & @LF)
;									For $j = 0 To 1
;										If _GUICtrlComboBox_SelectString($idPatWahl, $pna[1+$j] & ", " & $pna[2-$j]) <> -1 Then
;											ConsoleWrite("fand: " & GUICtrlRead($idpatwahl) & @lf)
;											$gefunden=1
;											ExitLoop
;										EndIf
;									Next; $j
;									If $gefunden Then ExitLoop
;								EndIf; $pos
;							Next ; $Tz
							If Not $gefunden Then ControlSetText($idHaupt, "", $idPatWahl, "")

							Local $gwIx = StringMid($gwDa, StringInStr($gwDa, "|", 0, -2) + 1)
							$gwIx = StringLeft($gwIx, StringLen($gwIx) - 1)
							;							ConsoleWrite(@LF & "gwNr: " & $gwNr & ", gwIx: " & $gwIx & ", gwDa: " & $gwDa & @LF)
							Local $verschieden = 0 ; wenn Dateien mit verschiedenem Rumpf gleichzeitig markiert, dann nicht unbenennen
							Local $gEntr = _GetEntries($idDtn)
							If UBound($gEntr) > 1 Then
								Local $drv, $dir, $tru, $ext, $t2
								_PathSplit($aDL[$gEntr[0]], $drv, $dir, $tru, $ext)
								For $ix = 1 To UBound($gEntr) - 1
									_PathSplit($aDL[$gEntr[$ix]], $drv, $dir, $t2, $ext)
									If $t2 <> $tru Then
										$verschieden = 1
										ExitLoop
									EndIf ; $t2<>$tru Then
									;									ConsoleWrite("markD ix: " & $ix & " " & $gEntr[$ix] & " " & $aDL[$gEntr[$ix]] & @LF)
								Next ; $i
							EndIf ; UBound($aGetData) Then
							If $verschieden Then
								GUICtrlSetState($idntUm, $GUI_CHECKED)
								$obntum = 1
							EndIf ; $verschieden Then
							If $gwIx >= 0 And $gwIx < UBound($aDL) - 1 Then ; das zuletzt Markierte
								Local $udat = $aDL[$gwIx]
								If $udat <> $altdatei Then
									Local $gesdatei = $Pfad & (StringRight($Pfad, 1) = "\" ? "" : "\") & $udat
									;	ConsoleWrite("hWndFrom: " & $hWndFrom & ", hWndListView: " & $hWndDtn & ", idDtn: " & $idDtn & ", iCode: " & $iCode & ", ind: " & $gwIx & ", ubound(aDL): " & UBound($aDL) & ", datei: " & $udat & @lf)
									$altdatei = $udat
									If IsObj($oPDF) Then
										$oPDF.src = ""
									EndIf ; IsObj($oPDF) Then
									If IsObj($oExcel) Then
										Local $eerg = _Excel_BookList($oExcel)
										If $eerg <> 0 Then
											;						    ConsoleWrite("eerg: " & $eerg & @lf)
											If $oExcel.workbooks.count > 0 Then
												$oExcel.workbooks(1).close
											EndIf ; $oExcel.workbooks.count>0 Then
										Else ; $eerg<>0 Then
											machExcel()
										EndIf ; $eerg<>0 Then Else
										$oExcel.Visible = 0
										GUICtrlSetState($GUI_Excel, $GUI_HIDE)
									EndIf ; If IsObj($oExcel) Then
									If IsObj($oWord) Then
										Local $werg = _Word_DocGet($oWord)
										If $werg <> 0 Then
											;						    ConsoleWrite("werg: " & $werg & @lf)
											If $oWord.documents.count > 0 Then
												$oWord.Documents(1).close
											EndIf ; $oWord.documents.count>0 Then
										Else ; $werg<>0 Then
											machWord()
										EndIf ; $werg<>0 Then Else
										$oWord.Visible = 0
										GUICtrlSetState($GUI_Word, $GUI_HIDE)
									EndIf ; IsObj($oWord) Then Else
									GUICtrlDelete($idPic)
									;						  ConsoleWrite("Datei: '" & $udat & "' " & StringLower(StringRight($udat,StringLen($udat)-StringInStr($udat,".",0,-1))) & @lf)
									Switch StringLower(StringRight($udat, StringLen($udat) - StringInStr($udat, ".", 0, -1)))
										Case "jpg", "jpeg", "tga", "tif", "tiff", "png", "bmp", "gif", "raw", "wbmp", "wmf", "rgb", "sff", "ico", "b3d", "crw", "cr2", "cr3", "dcm", "acr", "ima", "dds", "djvu", "iw44", "dxf", "ecw", "emf", "eps", "ps", "exr", "g3", "hdp", "jxr", "wdp", "heic", "iff", "lbm", "jls", "jp2", "jpc", "j2k", "jpf", "jfif", "jpe", "jpm", "jxl", "mng", "jng", "pbm", "pcd", "pcx", "dcx", "pgm", "ppm", "psd", "psb", "psp", "qoi", "ras", "sun", "rle", "sgi", "sid", "webp", "xbm", "xpm"
											$idPic = GUICtrlCreatePic($gesdatei, 1, 0, $aML * 0.667 - 5, $aDUR * 0.667)
											;							if $idPic=0 Then $idPic = GUICtrlCreatePic($gesdatei, 1,0 , 200,100)
											;							ConsoleWrite("versuchte das Bild '" & $gesdatei & "' zu malen, idPic: " & $idPic & ", aPos[2]: " & $apos[2] & ", aPos[3]: " & $apos[3] & @LF)
											If $idPic <> 0 Then GUICtrlSetState($idPic, $GUI_SHOW)
										Case "xls", "xlsx", "csv"
											If $obEx And IsObj($oExcel) Then
												;							   ConsoleWrite("will jetzt "  & $gesdatei & " in Excel anzeigen " & @LF)
												;							   $oExcel.Workbooks.Open($gesdatei,0,-1)
												_Excel_BookOpen($oExcel, $gesdatei, True) ; Word könnte "\\" ignorieren, Excel 2000 nicht
												$oExcel.Visible = 1
												GUICtrlSetState($GUI_Excel, $GUI_SHOW)
											EndIf ; $obEx And IsObj($oExcel) Then
										Case "doc", "docx", "txt", "qbquery", "sql", "bas", "rtf", "odt", "dot", "html", "htm", "ini", "log", "tmp" ; ,"" geht nicht
											If $obWd And IsObj($oWord) Then
												;							  $oWord.Documents.Open($gesdatei,0,-1)
												_Word_DocOpen($oWord, $gesdatei, Default, Default, True)
												ConsoleWrite("zeige jetzt " & $gesdatei & " in Word an " & @LF)
												$oWord.Visible = 1
											EndIf ; $obWd And IsObj($oWord) Then
										Case "pdf"
											;						 	 ConsoleWrite(" vor pdf, mit " & $gesdatei & @lf)
											If $obPDF And IsObj($oPDF) Then
												;							  ConsoleWrite(" in pdf, mit " & $gesdatei & @lf)
												Local $hFileOpen = FileOpen($gesdatei, $FO_READ + $FO_UTF8)
												;							  ConsoleWrite("datei: '" & $gesdatei & "', hFileOpen: " & $hFileOpen & @LF)
												If $hFileOpen <> -1 Then
													Local $sFileRead = FileRead($hFileOpen, 5)
													If $sFileRead = "%PDF-" Then
														;								ConsoleWrite("zeige PDF an: " & $gesdatei & @lf)
														$oPDF.src = $gesdatei ;PDF FILE 1
														$obneupdf = 1
														GUICtrlSetState($GUI_PDF, $GUI_SHOW)
													EndIf ; $sFileRead="%PDF-" Then
													FileClose($hFileOpen)
													Sleep(5 * $Itv)
												EndIf ; $hFileOpen <> -1 Then
												ListAct() ; für $idDtn
											EndIf ; IsObj($oPDF) Then
									EndSwitch ; StringLower(StringRight($udat,StringLen($udat)-StringInStr($udat,".",0,-1)))
									If $obneupdf = 0 Then GUICtrlSetState($GUI_PDF, $GUI_HIDE)
								EndIf ; $udat <> $altdatei
							EndIf ; If $gwIx >= 0 And $gwIx < UBound($aDL)-1 Then
							;                    ConsoleWrite("Row: " & DllStructGetData($tNMHDR, 4) & " - Col: " & DllStructGetData($tNMHDR, 5) & @LF)
							;                Case -12
							;                Case else
							;                    ConsoleWrite("iCode:" & $iCode & ", Row: " & DllStructGetData($tNMHDR, 4) & ", Col: " & DllStructGetData($tNMHDR, 5) & @LF)
						EndIf ; $gwNr
				EndSwitch ; $iCode
			Case Else
				Switch $iCode
					Case $GUI_EVENT_MOUSEMOVE ; -11
					Case $GUI_EVENT_RESIZED ; -12
					Case -320, -321 ; ??
					Case -1249 ; "Datei aufrufen"
					Case Else
						ConsoleWrite("in WM_NOTIFY, hWndFrom: " & $hWndFrom & ", iCode: " & $iCode & @LF)
				EndSwitch ; $iCode
		EndSwitch ; $hWndFrom
	EndIf ; $hwndfrom<>0 Then ; sonst HWnd() gescheitert
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

; aufgerufen in waehlepfad, zeichneHaupt
Func fuelleLetztePfade()
	GUICtrlDelete($idRecentfilesmenu)
	$idRecentfilesmenu = GUICtrlCreateMenu("&Letzte Pfade", $idFile, 1)
	;	ConsoleWrite("vor Schleife in fuelleLetztePfade " & @LF)
	$iLPf = RegRead($RegPath, "iLPfad")
	If $iLPf = "" Then $iLPf = 1
	Local $i = $iLPf
	For $j = 1 To $lMaxZahl
		;		ConsoleWrite("in Schleife " & $j & " in fuelleLetztePfade " & @LF)
		$i = $i - 1
		If $i = 0 Then $i = $lMaxZahl - 1
		Local $lPf = RegRead($RegPath, $i)
		If $lPf = "" Then ExitLoop
		$arrlPf[$i] = GUICtrlCreateMenuItem($lPf, $idRecentfilesmenu)
		;		ConsoleWrite("arrlpf[" & $j & "]: " & $arrlpf[$j] & @LF)
	Next ; $j
EndFunc   ;==>fuelleLetztePfade

Func fuelleUnPfade()
	; Const $lUnZahl=15; Zahl der erinnerten letzten patientenunabhängigen Speicherpfade
	; Dim $arrUnPf[0]; Array der letzten patientenunabhängigen Speicherpfade
	; Dim $iLUnZg; Ordnungsziffer des letzten patientenunabhängigen Speicherpfades
	$iLUnZg = RegRead($RegPath, "iLUnZg")
	If $iLUnZg = "" Then $iLUnZg = 1
	Local $i = $iLUnZg, $k = 0
	For $j = 1 To $lUnZahl
		$i = $i - 1
		If $i = 0 Then $i = $lUnZahl - 1
		Local $unPf = RegRead($RegPath, "un" & $i)
		If $unPf = "" Then ExitLoop
		ReDim $arrUnPf[$k + 1]
		$arrUnPf[$k] = $unPf
		$k = $k + 1
	Next ; $j
EndFunc   ;==>fuelleUnPfade

; aufgerufen in Haupt (3x)
Func waehlepfad($wpfad = "") ; bestimmt $Pfad
	ConsoleWrite("Wähle Pfad: " & $wpfad & @LF)
	;	$Pfad = RegRead($RegPath, "Pfad")
	Local $altPfad = $Pfad
	;	  ConsoleWrite(" pfad: " & $Pfad & ", altPfad: " & $altPfad & @LF)
	If $wpfad = "" Then
		$Pfad = FileSelectFolder("Wähle Pfad, dessen Dateien benannt werden sollen", "", @DocumentsCommonDir)
		If $Pfad = "" Then $Pfad = $altPfad
	Else ; $wpfad = "" Then
		$Pfad = $wpfad
	EndIf ; $wpfad = "" Then Else
	If $Pfad <> $altPfad Then
		;		ConsoleWrite("waehlepfad, schreibe: " & $Pfad & ", altPfad: " & $altPfad & ", iLPf: " & $iLPf & @LF)
		RegWrite($RegPath, "Pfad", "REG_SZ", $Pfad)
		If $altPfad <> "" Then
			RegWrite($RegPath, $iLPf, "REG_SZ", $altPfad)
			$iLPf = $iLPf + 1
			If $iLPf > $lMaxZahl Then $iLPf = 1
			RegWrite($RegPath, "iLPfad", "REG_SZ", $iLPf)
		EndIf ; $altPfad<>"" Then
		machfilelist(1)
	EndIf ; $Pfad <> $altPfad Then
	;	  If @error <> 1 Then GUICtrlCreateMenuItem($Pfad, $idRecentfilesmenu)
	fuelleLetztePfade()
EndFunc   ;==>waehlepfad

; aufgerufen in HotKey
Func TuComZu()
	Return ComZu("Comzu über Hotkey")
EndFunc; TuComZu()

; aufgerufen in Haupt (2x + 1 x HotKeySet), machfilelist
Func ComZu($quelle)
	ConsoleWrite("in ComZu, " & $quelle & ", optaktiv = " & $optaktiv & @LF)
	Switch $PatUnAktiv
		Case 1
			GUIDelete($idVPat)
			$PatUnAktiv = 0
			machfilelist(1) ; für $idDtn; wenn Datei entfernt, dann Dateiliste neu anzeigen
			Return
		Case Else
	EndSwitch ; $PatUnAktiv
	Switch $optaktiv
		Case 1
			ConsoleWrite("Comzu, " & $quelle & ", mit $optaktiv " & $optaktiv & @LF)
			;		If True Then
			GUIDelete($idHaupt)
			zeichneHaupt() ; für $idHaupt
			GUIDelete($idOpt) ; ... just delete the GUI ...
			GUICtrlSetState($iOpt, $GUI_ENABLE) ; ... enable button (previously disabled)
			$optaktiv = 0
		Case 2
			ConsoleWrite("Comzu, " & $quelle & ", mit $optaktiv " & $optaktiv & @LF)
			GUIDelete($idOpt) ; ... just delete the GUI ...
			GUICtrlSetState($iOpt, $GUI_ENABLE) ; ... enable button (previously disabled)
			$optaktiv = 0
			Return 1
		Case Else
			If IsObj($oPDF) Then GUICtrlDelete($oPDF)
			If IsObj($oWord) Then GUICtrlDelete($oWord)
			If __Word_CloseOnQuit() Then
				_Word_Quit($oWord, False, $WdWordDocument, True)
				ShellExecute("taskkill.exe", "/im winword.exe /f", "", "", @SW_HIDE)
			EndIf ; __Word_CloseOnQuit() Then
			If IsObj($oExcel) Then GUICtrlDelete($oExcel)
			If __Excel_CloseOnQuit($oExcel) Then
				_Excel_Close($oExcel, False, True)
				ShellExecute("taskkill.exe", "/im excel.exe /f", "", "", @SW_HIDE)
			EndIf ; __Excel_CloseOnQuit($oExcel)	Then
			GUIDelete($idHaupt)
			myEnd($PatCn)
			ConsoleWrite("Exit in Comzu, " & $quelle & ", mit " & $ExitFnr & @LF)
			Exit ($ExitFnr)
	EndSwitch ; $optaktiv
EndFunc   ;==>ComZu

; aufgerufen in machfilelist
Func FileListFehler($FM, $Pfad, $muster)
	Switch @error
		Case 0
		Case 1
			MsgBox($MB_ICONERROR, $FM, "Pfad ungültig : '" & $Pfad & "'")
		Case 4
			MsgBox($MB_ICONWARNING, $FM, "Keine Dateien '" & $muster & "' in '" & $Pfad & "' gefunden")
		Case Else
			MsgBox($MB_ICONERROR, $FM, @error)
	EndSwitch    ; @error
	Return @error
EndFunc   ;==>FileListFehler

; aufgerufen in machfilelist
Func MultStr($buch, $zahl)
	Local $stri
	For $i = 0 To $zahl - 1
		$stri = $stri & $buch
	Next
	Return $stri
EndFunc   ;==>MultStr

; aufgerufen in ComZu
Func myEnd(ByRef $PatCn)
	; Abfrage freigeben
	; Verbindung beenden
	; _MySQL_Close($PatCn)
	; MYSQL beenden
	_MySQL_EndLibrary()
EndFunc   ;==>myEnd

