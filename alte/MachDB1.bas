Attribute VB_Name = "Module1"
'Bauanleitung für eine Datenbank wie `//linux/quelle` vom 19.9.09 21:44:54
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.Connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 113, 212) As New CString, ArtZ&(3, 113)
Dim hDBn$ ' hiesiger Datenbankname


Function FüllStr0()
 Str(0, 0, 0) = "DMPInkonsistenzen"
 Str(1, 0, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `DMPInkonsistenzen` AS select `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr` from (`_lfaelle` `l` left join `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`)))) group by `f`.`Pat_ID`"
End Function ' FüllStr0

Function FüllStr1()
 Str(0, 1, 0) = "DiabetesICDErsetzung"
 Str(1, 1, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `DiabetesICDErsetzung` AS select `d1`.`Pat_id` AS `pat_id`,`d1`.`ICD` AS `icd`,if(((`d2`.`ICD` is not null) or (`d`.`DokName` is not null) or (`e`.`Inhalt` is not null) or (`fr`.`Inhalt` is not null)),concat('E',(substr(replace(`d1`.`ICD`,'-','1'),2) + 0.04)),concat('E',(substr(replace(`d1`.`ICD`,'-','1'),2) + 0.02))) AS `neu`,`d1`.`GesName` AS `gesname`,`d2`.`ICD` AS `fußicd`,`d`.`DokName` AS `dokname`,`e`.`Inhalt` AS `inhalt`,`fr`.`Inhalt` AS `frinhalt` from ((((`diagnosen` `d1` left join `diagnosen` `d2` on(((`d1`.`Pat_id` = `d2`.`Pat_id`) and (`d2`.`ICD` like 'L89%')))) left join `dokumente` `d` on(((`d1`.`Pat_id` = `d`.`Pat_ID`) and (`d`.`DokName` regexp 'Foto.*WA[^-]')))) left join `eintraege` `e` on(((`d1`.`Pat_id` = `e`.`Pat_ID`) and (`e`.`Art` = 'debr')))) left join `eintraege` `fr` on(((`d1" & _
  "`.`Pat_id` = `fr`.`Pat_ID`) and (`fr`.`Art` = 'htxt') and (`fr`.`Inhalt` regexp ' df[abc]')))) where (`d1`.`ICD` regexp 'E...7[01-]') group by `d1`.`Pat_id` order by substr(`neu`,6),`d1`.`Pat_id` desc"
End Function ' FüllStr1

Function FüllStr2()
 Str(0, 2, 0) = "DiagEingabe"
 Str(1, 2, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `DiagEingabe` AS select `anamnesebogen`.`Prim` AS `Prim`,`anamnesebogen`.`Pat_id` AS `Pat_id`,`anamnesebogen`.`Nachname` AS `Nachname`,`anamnesebogen`.`Vorname` AS `Vorname`,`anamnesebogen`.`NVorsatz` AS `NVorsatz`,`anamnesebogen`.`Titel` AS `Titel`,`anamnesebogen`.`Anrede` AS `Anrede`,`anamnesebogen`.`GebDat` AS `GebDat`,`anamnesebogen`.`Tkz` AS `Tkz`,`anamnesebogen`.`Versicherungsart` AS `Versicherungsart`,`anamnesebogen`.`Diabetestyp` AS `Diabetestyp`,`anamnesebogen`.`Diabetes seit` AS `Diabetes seit`,`anamnesebogen`.`Tabletten seit` AS `Tabletten seit`,`anamnesebogen`.`Insulin seit` AS `Insulin seit`,`anamnesebogen`.`Grund für Vorstellung` AS `Grund für Vorstellung`,`anamnesebogen`.`Familienanamnese` AS `Familienanamnese`,`anamnesebogen`.`Größe` AS `Größe`,`anamnesebogen`.`Gewicht` AS `Gewicht`,`" & _
  "anamnesebogen`.`Tendenz` AS `Tendenz`,`anamnesebogen`.`DiabetesMedikament 1` AS `DiabetesMedikament 1`,`anamnesebogen`.`DiabetesMedikament 1 Menge` AS `DiabetesMedikament 1 Menge`,`anamnesebogen`.`DiabetesMedikament 2` AS `DiabetesMedikament 2`,`anamnesebogen`.`DiabetesMedikament 2 Menge` AS `DiabetesMedikament 2 Menge`,`anamnesebogen`.`DiabetesMedikament 3` AS `DiabetesMedikament 3`,`anamnesebogen`.`DiabetesMedikament 3 Menge` AS `DiabetesMedikament 3 Menge`,`anamnesebogen`.`DiabetesMedikament 4` AS `DiabetesMedikament 4`,`anamnesebogen`.`DiabetesMedikament 4 Menge` AS `DiabetesMedikament 4 Menge`,`anamnesebogen`.`Insulinpumpe` AS `Insulinpumpe`,`anamnesebogen`.`Insulinpumpe seit` AS `Insulinpumpe seit`,`anamnesebogen`.`Insulinpumpe Marke` AS `Insulinpumpe Marke`,`anamnesebogen`.`Broteinheiten gesamt` AS `Broteinheiten gesamt`,`anamnesebogen`.`Broteinheiten früh` AS `Broteinheiten früh`" & _
  ",`anamnesebogen`.`Broteinheiten ZM früh` AS `Broteinheiten ZM früh`,`anamnesebogen`.`Broteinheiten mittags` AS `Broteinheiten mittags`,`anamnesebogen`.`Broteinheiten nachmittags` AS `Broteinheiten nachmittags`,`anamnesebogen`.`Broteinheiten abends` AS `Broteinheiten abends`,`anamnesebogen`.`Broteinheiten nachts` AS `Broteinheiten nachts`,`anamnesebogen`.`Essenszeit früh` AS `Essenszeit früh`,`anamnesebogen`.`Essenszeit vormittags` AS `Essenszeit vormittags`,`anamnesebogen`.`Essenszeit mittags` AS `Essenszeit mittags`,`anamnesebogen`.`Essenszeit nachmittags` AS `Essenszeit nachmittags`,`anamnesebogen`.`Essenszeit abends` AS `Essenszeit abends`,`anamnesebogen`.`Essenszeit spät` AS `Essenszeit spät`,`anamnesebogen`.`Spritz-Eß-Abstand früh` AS `Spritz-Eß-Abstand früh`,`anamnesebogen`.`Spritz-Eß-Abstand mittags` AS `Spritz-Eß-Abstand mittags`,`anamnesebogen`.`Spritz-Eß-Abstand abends` AS `Spr" & _
  "itz-Eß-Abstand abends`,`anamnesebogen`.`Spritzstelle früh` AS `Spritzstelle früh`,`anamnesebogen`.`Spritzstelle mittags` AS `Spritzstelle mittags`,`anamnesebogen`.`Spritzstelle abends` AS `Spritzstelle abends`,`anamnesebogen`.`Spritzstelle nachts` AS `Spritzstelle nachts`,`anamnesebogen`.`Jahr letzte Diabetesschulung` AS `Jahr letzte Diabetesschulung`,`anamnesebogen`.`Ort Schulung` AS `Ort Schulung`,`anamnesebogen`.`letztes HbA1c` AS `letztes HbA1c`,`anamnesebogen`.`gemessen am` AS `gemessen am`,`anamnesebogen`.`vorherige Werte` AS `vorherige Werte`,`anamnesebogen`.`BZMessungen selbst` AS `BZMessungen selbst`,`anamnesebogen`.`Gerät` AS `Gerät`,`anamnesebogen`.`BZMessungen pW` AS `BZMessungen pW`,`anamnesebogen`.`BZMessungen pW ndE` AS `BZMessungen pW ndE`,`anamnesebogen`.`BZMessungen p W nachts` AS `BZMessungen p W nachts`,`anamnesebogen`.`Aufschreiben` AS `Aufschreiben`,`anamnesebogen`." & _
  "`BZWerte v d Essen` AS `BZWerte v d Essen`,`anamnesebogen`.`BZWerte n d Essen` AS `BZWerte n d Essen`,`anamnesebogen`.`UZ Tageszeit` AS `UZ Tageszeit`,`anamnesebogen`.`Unterzucker pM` AS `Unterzucker pM`,`anamnesebogen`.`UZ rechtzeitig` AS `UZ rechtzeitig`,`anamnesebogen`.`Fremde Hilfe pa` AS `Fremde Hilfe pa`,`anamnesebogen`.`Bewußtlos pa` AS `Bewußtlos pa`,`anamnesebogen`.`Keto pa` AS `Keto pa`,`anamnesebogen`.`BZgr300 pM` AS `BZgr300 pM`,`anamnesebogen`.`Bluthochdruck` AS `Bluthochdruck`,`anamnesebogen`.`BHD seit` AS `BHD seit`,`anamnesebogen`.`BHD beh mit` AS `BHD beh mit`,`anamnesebogen`.`Blutdruckwerte` AS `Blutdruckwerte`,`anamnesebogen`.`BDselbst` AS `BDselbst`,`anamnesebogen`.`Schwanger` AS `Schwanger`,`anamnesebogen`.`Schwanger seit` AS `Schwanger seit`,`anamnesebogen`.`Augensp zuletzt` AS `Augensp zuletzt`,`anamnesebogen`.`Augensp Befund` AS `Augensp Befund`,`anamnesebogen`.`N" & _
  "etzhaut gelasert` AS `Netzhaut gelasert`,`anamnesebogen`.`Sehminderung unbehebbar` AS `Sehminderung unbehebbar`,`anamnesebogen`.`Diabet Nierenschaden` AS `Diabet Nierenschaden`,`anamnesebogen`.`Albumin zuletzt` AS `Albumin zuletzt`,`anamnesebogen`.`erhöht?` AS `erhöht?`,`anamnesebogen`.`Dialyse` AS `Dialyse`,`anamnesebogen`.`Dialyse seit` AS `Dialyse seit`,`anamnesebogen`.`andere Nierenerkrankung` AS `andere Nierenerkrankung`,`anamnesebogen`.`Herzkrankheit` AS `Herzkrankheit`,`anamnesebogen`.`Angina pectoris` AS `Angina pectoris`,`anamnesebogen`.`Herzinfarkt` AS `Herzinfarkt`,`anamnesebogen`.`Herzinfarkt wann` AS `Herzinfarkt wann`,`anamnesebogen`.`PTCA oder Stent` AS `PTCA oder Stent`,`anamnesebogen`.`Bypass kardial` AS `Bypass kardial`,`anamnesebogen`.`Bypass wann` AS `Bypass wann`,`anamnesebogen`.`Herzschwäche` AS `Herzschwäche`,`anamnesebogen`.`Herzkrankheit Beschreibung` AS `Herzkra" & _
  "nkheit Beschreibung`,`anamnesebogen`.`Hirndurchblutungsstörung` AS `Hirndurchblutungsstörung`,`anamnesebogen`.`Schlaganfall` AS `Schlaganfall`,`anamnesebogen`.`Beindurchblutungsstörung` AS `Beindurchblutungsstörung`,`anamnesebogen`.`Schaufensterkrankheit` AS `Schaufensterkrankheit`,`anamnesebogen`.`Bypaß peripher` AS `Bypaß peripher`,`anamnesebogen`.`Geschwür` AS `Geschwür`,`anamnesebogen`.`Amputation` AS `Amputation`,`anamnesebogen`.`pAVK Beschreibung` AS `pAVK Beschreibung`,`anamnesebogen`.`Ameisenlaufen` AS `Ameisenlaufen`,`anamnesebogen`.`Ameisen Ausmaß` AS `Ameisen Ausmaß`,`anamnesebogen`.`Druckstellen` AS `Druckstellen`,`anamnesebogen`.`Verformungen` AS `Verformungen`,`anamnesebogen`.`Verformungen Beschreibung` AS `Verformungen Beschreibung`,`anamnesebogen`.`Fußpflege` AS `Fußpflege`,`anamnesebogen`.`Podologie` AS `Podologie`,`anamnesebogen`.`Einlagen` AS `Einlagen`,`anamnesebogen`" & _
  ".`Neue Fußkomplikationen` AS `Neue Fußkomplikationen`,`anamnesebogen`.`Entleerungsstörungen Magen` AS `Entleerungsstörungen Magen`,`anamnesebogen`.`Entleerungsstörungen Harnblase` AS `Entleerungsstörungen Harnblase`,`anamnesebogen`.`Schwindel Aufstehen` AS `Schwindel Aufstehen`,`anamnesebogen`.`Folgeerkrankungen Haut` AS `Folgeerkrankungen Haut`,`anamnesebogen`.`Bewegungseinschränkungen` AS `Bewegungseinschränkungen`,`anamnesebogen`.`Sexualstörung` AS `Sexualstörung`,`anamnesebogen`.`Sexualstörung seit` AS `Sexualstörung seit`,`anamnesebogen`.`Weitere Anamnese` AS `Weitere Anamnese`,`anamnesebogen`.`Alkohol` AS `Alkohol`,`anamnesebogen`.`Tabak` AS `Tabak`,`anamnesebogen`.`tabakex` AS `tabakex`,`anamnesebogen`.`tabakbis` AS `tabakbis`,`anamnesebogen`.`tabakakt` AS `tabakakt`,`anamnesebogen`.`tabakmenge` AS `tabakmenge`,`anamnesebogen`.`Weitere Medikation` AS `Weitere Medikation`,`anamnese" & _
  "bogen`.`Liphypertrophien Abdomen` AS `Liphypertrophien Abdomen`,`anamnesebogen`.`Liphypertrophien Beine` AS `Liphypertrophien Beine`,`anamnesebogen`.`Liphypertrophien Arme` AS `Liphypertrophien Arme`,`anamnesebogen`.`Beinbefund` AS `Beinbefund`,`anamnesebogen`.`Hyperkeratosen` AS `Hyperkeratosen`,`anamnesebogen`.`Ulcera` AS `Ulcera`,`anamnesebogen`.`Kraft Zehenheber` AS `Kraft Zehenheber`,`anamnesebogen`.`Kraft Zehenbeuger` AS `Kraft Zehenbeuger`,`anamnesebogen`.`Kraft Knie` AS `Kraft Knie`,`anamnesebogen`.`ASR` AS `ASR`,`anamnesebogen`.`PSR` AS `PSR`,`anamnesebogen`.`Oberflächensensibilität` AS `Oberflächensensibilität`,`anamnesebogen`.`Monofilamenttest` AS `Monofilamenttest`,`anamnesebogen`.`Kalt-Warm` AS `Kalt-Warm`,`anamnesebogen`.`Vibration IK` AS `Vibration IK`,`anamnesebogen`.`Vibration Großzehe` AS `Vibration Großzehe`,`anamnesebogen`.`Puls Leiste` AS `Puls Leiste`,`anamnesebogen" & _
  "`.`Puls Kniekehle` AS `Puls Kniekehle`,`anamnesebogen`.`Puls Atp` AS `Puls Atp`,`anamnesebogen`.`Puls Adp` AS `Puls Adp`,`anamnesebogen`.`RR` AS `RR`,`anamnesebogen`.`RRTurboMed` AS `RRTurboMed`,`anamnesebogen`.`Herz` AS `Herz`,`anamnesebogen`.`Lunge` AS `Lunge`,`anamnesebogen`.`Bauch` AS `Bauch`,`anamnesebogen`.`WS` AS `WS`,`anamnesebogen`.`NL` AS `NL`,`anamnesebogen`.`SD` AS `SD`,`anamnesebogen`.`Carotiden` AS `Carotiden`,`anamnesebogen`.`NNH` AS `NNH`,`anamnesebogen`.`Zähne` AS `Zähne`,`anamnesebogen`.`Mundhöhle` AS `Mundhöhle`,`anamnesebogen`.`LK` AS `LK`,`anamnesebogen`.`BeinödVen` AS `BeinödVen`,`anamnesebogen`.`Neuro sonst` AS `Neuro sonst`,`anamnesebogen`.`Weitere Befunde` AS `Weitere Befunde`,`anamnesebogen`.`Schulung` AS `Schulung`,`anamnesebogen`.`DMP` AS `DMP`,`anamnesebogen`.`DMSchulz` AS `DMSchulz`,`anamnesebogen`.`DMSchL` AS `DMSchL`,`anamnesebogen`.`RRSchulz` AS `RRSchulz" & _
  "`,`anamnesebogen`.`DMPhier` AS `DMPhier`,`anamnesebogen`.`HANr` AS `HANr`,`anamnesebogen`.`HANr2` AS `HANr2`,`anamnesebogen`.`letzte Änderung` AS `letzte Änderung`,`anamnesebogen`.`Diagnosen` AS `Diagnosen`,`anamnesebogen`.`Vorgestellt` AS `Vorgestellt`,`anamnesebogen`.`Versicherung` AS `Versicherung`,`anamnesebogen`.`AktZeit` AS `AktZeit`,`anamnesebogen`.`Ther1` AS `Ther1`,`anamnesebogen`.`TherAkt` AS `TherAkt`,`anamnesebogen`.`obAn1eing` AS `obAn1eing`,`anamnesebogen`.`obAn2eing` AS `obAn2eing`,`anamnesebogen`.`obAnAeing` AS `obAnAeing`,`anamnesebogen`.`obCheck` AS `obCheck`,`anamnesebogen`.`obBZausgew` AS `obBZausgew`,`anamnesebogen`.`obOSaufgek` AS `obOSaufgek`,`anamnesebogen`.`obPodAufgek` AS `obPodAufgek`,`anamnesebogen`.`obMBlAusgeh` AS `obMBlAusgeh`,`anamnesebogen`.`obSchulaufgek` AS `obSchulaufgek`,`anamnesebogen`.`obDMPaufgekl` AS `obDMPaufgekl`,`anamnesebogen`.`obMedNetz` AS `" & _
  "obMedNetz`,`anamnesebogen`.`Hausarzt` AS `Hausarzt`,`anamnesebogen`.`ob` AS `ob`,`anamnesebogen`.`QS` AS `QS`,`anamnesebogen`.`QT` AS `QT`,if((`anamnesebogen`.`Größe` = 0),'',(((`anamnesebogen`.`Gewicht` / `anamnesebogen`.`Größe`) / `anamnesebogen`.`Größe`) * if((`anamnesebogen`.`Größe` > 3),10000,1))) AS `bmi`,concat(`anamnesebogen`.`Nachname`,' ',`anamnesebogen`.`Vorname`) AS `gesname` from `anamnesebogen` where `anamnesebogen`.`Pat_id` in (select `aktf`.`pat_id` AS `pat_id` from `aktf`) order by `anamnesebogen`.`Pat_id` desc"
End Function ' FüllStr2

Function FüllStr3()
 Str(0, 3, 0) = "Trop-Tests"
 Str(1, 3, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `Trop-Tests` AS select `e`.`FID` AS `FID`,`e`.`Pat_ID` AS `Pat_ID`,`e`.`ZeitPunkt` AS `ZeitPunkt`,`e`.`Art` AS `Art`,`e`.`Inhalt` AS `Inhalt`,`e`.`absPos` AS `absPos`,`e`.`AktZeit` AS `AktZeit`,`e`.`QS` AS `QS`,`e`.`QT` AS `QT`,`e`.`StByte` AS `StByte` from `eintraege` `e` where ((`e`.`Art` like 'trop%') or ((`e`.`Inhalt` like '%trop%') and (not((`e`.`Inhalt` like '%troph%'))) and (not((`e`.`Inhalt` like '%tropa%'))) and (not((`e`.`Inhalt` like '%tropf%'))) and (not((`e`.`Inhalt` like '%trope%'))) and (not((`e`.`Inhalt` like '%tropos%'))))) order by `e`.`Pat_ID`,`e`.`ZeitPunkt` desc"
End Function ' FüllStr3

Function FüllStr4()
 Str(0, 4, 0) = "__fuerlmp"
 Str(1, 4, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `__fuerlmp` AS select `mp`.`Pat_ID` AS `pat_id`,`mp`.`ZeitPunkt` AS `zeitpunkt`,`mp`.`MPNr` AS `mpnr` from `medplan` `mp` group by `mp`.`Pat_ID`,`mp`.`ZeitPunkt`,`mp`.`MPNr` order by `mp`.`Pat_ID`,`mp`.`ZeitPunkt` desc,`mp`.`MPNr` desc"
End Function ' FüllStr4

Function FüllStr5()
 Str(0, 5, 0) = "__kontakttage"
 Str(1, 5, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `__kontakttage` AS select `e`.`Pat_ID` AS `pat_id`,`e`.`ZeitPunkt` AS `zeitpunkt` from `eintraege` `e` where ((`e`.`ZeitPunkt` between concat(year((now() - interval 14 day)),'-',((((month((now() - interval 14 day)) - 1) DIV 3) * 3) + 1),'-01') and concat((year((now() - interval 14 day)) + round((((((month((now() - interval 14 day)) - 1) DIV 3) * 3) + 4) / 12),0)),'-',(((((month((now() - interval 14 day)) - 1) DIV 3) * 3) + 4) % 12),'-01')) and (`e`.`Art` in ('notiz','telef','ni','gstel','gs','rz','ep','bga','tk','APK','wr','ga','tst','cr','ke','hz','mh','ag','ph','pq','er','ds','st','eb','fa','bz','rp','uzu','hypo','colo','aug','beweg','pros','impf','gyn','caro','beruf','ap','mu','rauch','alko','fams','schula','ass','kra','proc','au','GPD','ba','ARCHIE2','gewicht','gewi','rrvgl','bzvgl','bzm','bztp','" & _
  "bks','anal','andm','usal','usdm','doppler','duplex','sono','sd','UKG','Größe','HbA1c','hyper','fuß','keto','wv','ulc','kv','debr','EKG','LZRR','Lufu','lactoset','trop','temp','oGTT','gpt','bmi','urin','taille','hüfte','puls','GDT','bef'))) group by `e`.`Pat_ID`,cast(`e`.`ZeitPunkt` as date)"
End Function ' FüllStr5

Function FüllStr6()
 Str(0, 6, 0) = "__lfaelle"
 Str(1, 6, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `__lfaelle` AS select `f`.`BhFB` AS `mbhfb`,`f`.`BhFE1` AS `bhfe1`,`f`.`Pat_ID` AS `pid` from `faelle` `f` order by `f`.`Pat_ID`,`f`.`BhFB`"
End Function ' FüllStr6

Function FüllStr7()
 Str(0, 7, 0) = "_f1"
 Str(1, 7, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `_f1` AS select `faelle`.`BhFB` AS `bhfb`,`faelle`.`Pat_ID` AS `pat_id` from `faelle` order by `faelle`.`Pat_ID`,`faelle`.`BhFB` desc,`faelle`.`SchGr`"
End Function ' FüllStr7

Function FüllStr8()
 Str(0, 8, 0) = "_faellenachschgr"
 Str(1, 8, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `_faellenachschgr` AS select `faelle`.`FID` AS `FID`,`faelle`.`Pat_ID` AS `Pat_ID`,`faelle`.`Quartal` AS `Quartal`,`faelle`.`Nachname` AS `Nachname`,`faelle`.`Vorname` AS `Vorname`,`faelle`.`lfdnr` AS `lfdnr`,`faelle`.`TMFNr` AS `TMFNr`,`faelle`.`VKNr` AS `VKNr`,`faelle`.`BhFB` AS `BhFB`,`faelle`.`BhFE1` AS `BhFE1`,`faelle`.`BhFE2` AS `BhFE2`,`faelle`.`f4202` AS `f4202`,`faelle`.`ausgst` AS `ausgst`,`faelle`.`KtrAbrB` AS `KtrAbrB`,`faelle`.`AbrAr` AS `AbrAr`,`faelle`.`lVorl` AS `lVorl`,`faelle`.`IK` AS `IK`,`faelle`.`KVKs` AS `KVKs`,`faelle`.`KVKserg` AS `KVKserg`,`faelle`.`Kasse` AS `Kasse`,`faelle`.`GebOr` AS `GebOr`,`faelle`.`AbrGb` AS `AbrGb`,`faelle`.`PersKreis` AS `PersKreis`,`faelle`.`SKtZusatz` AS `SKtZusatz`,`faelle`.`f4206` AS `f4206`,`faelle`.`ÜwText` AS `ÜwText`,`faelle`.`f4210` AS `f4210`" & _
  ",`faelle`.`AkfHAH` AS `AkfHAH`,`faelle`.`AkfAB0` AS `AkfAB0`,`faelle`.`AkfAK` AS `AkfAK`,`faelle`.`statNuller` AS `statNuller`,`faelle`.`ÜbwV` AS `ÜbwV`,`faelle`.`AndÜw` AS `AndÜw`,`faelle`.`Übw` AS `Übw`,`faelle`.`ÜbwLANR` AS `ÜbwLANR`,`faelle`.`ÜWZiel` AS `ÜWZiel`,`faelle`.`ÜWNNr` AS `ÜWNNr`,`faelle`.`ÜWNaN` AS `ÜWNaN`,`faelle`.`ÜWTit` AS `ÜWTit`,`faelle`.`ÜWVor` AS `ÜWVor`,`faelle`.`ÜWVsw` AS `ÜWVsw`,`faelle`.`üwvid` AS `üwvid`,`faelle`.`statKlasse` AS `statKlasse`,`faelle`.`f4237` AS `f4237`,`faelle`.`statBehTage` AS `statBehTage`,`faelle`.`SchGr` AS `SchGr`,`faelle`.`Weiterbeh` AS `Weiterbeh`,`faelle`.`PGeb` AS `PGeb`,`faelle`.`PGebErg` AS `PGebErg`,`faelle`.`Mahnfrist` AS `Mahnfrist`,`faelle`.`GOÄKatNr` AS `GOÄKatNr`,`faelle`.`GOÄKatName` AS `GOÄKatName`,`faelle`.`abrArzt` AS `abrArzt`,`faelle`.`privVers` AS `privVers`,`faelle`.`AdNam` AS `AdNam`,`faelle`.`AdStr` AS `AdStr`,`faelle" & _
  "`.`AdPlz` AS `AdPlz`,`faelle`.`AdOrt` AS `AdOrt`,`faelle`.`BhFE` AS `BhFE`,`faelle`.`s8000` AS `s8000`,`faelle`.`s8100` AS `s8100`,`faelle`.`AktZeit` AS `AktZeit`,`faelle`.`Fanf` AS `Fanf`,`faelle`.`altQuart` AS `altQuart`,`faelle`.`QAnf` AS `QAnf`,`faelle`.`QEnd` AS `QEnd`,`faelle`.`QS` AS `QS`,`faelle`.`QT` AS `QT`,`faelle`.`TherArt` AS `TherArt`,`faelle`.`StByte` AS `StByte`,`faelle`.`absPos` AS `absPos` from `faelle` order by `faelle`.`SchGr`"
End Function ' FüllStr8

Function FüllStr9()
 Str(0, 9, 0) = "_fuerlmp"
 Str(1, 9, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `_fuerlmp` AS select `__fuerlmp`.`pat_id` AS `pat_id`,`__fuerlmp`.`mpnr` AS `mpnr` from `__fuerlmp` group by `__fuerlmp`.`pat_id`"
End Function ' FüllStr9

Function FüllStr10()
 Str(0, 10, 0) = "_kontakttage"
 Str(0, 10, 1) = "`pat_id`"
 Str(0, 10, 2) = "`zeitpunkt`"
 ArtZ(0, 10) = 2
 Str(1, 10, 0) = "CREATE TABLE `_kontakttage` ("
 Str(1, 10, 1) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 10, 2) = " `zeitpunkt` datetime DEFAULT NULL"
 Str(1, 10, 3) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr10

Function FüllStr11()
 Str(0, 11, 0) = "_kontaktzahl"
 Str(1, 11, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `_kontaktzahl` AS select count(0) AS `ct`,`__kontakttage`.`pat_id` AS `pat_id` from `__kontakttage` group by `__kontakttage`.`pat_id`"
End Function ' FüllStr11

Function FüllStr12()
 Str(0, 12, 0) = "_lfaelle"
 Str(1, 12, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `_lfaelle` AS select `__lfaelle`.`mbhfb` AS `mbhfb`,`__lfaelle`.`bhfe1` AS `bhfe1`,`__lfaelle`.`pid` AS `pid` from `__lfaelle` group by `__lfaelle`.`pid` desc"
End Function ' FüllStr12

Function FüllStr13()
 Str(0, 13, 0) = "aktf"
 Str(1, 13, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `aktf` AS select `faelle`.`Pat_ID` AS `pat_id`,`faelle`.`FID` AS `fid`,`faelle`.`SchGr` AS `schgr`,`faelle`.`VKNr` AS `vknr` from `faelle` where ((`faelle`.`SchGr` <> '90') and (`faelle`.`Quartal` = (select concat((((month((now() - interval 14 day)) - 1) DIV 3) + 1),year((now() - interval 14 day))) AS `lq`))) order by `faelle`.`Pat_ID`,`faelle`.`FID` desc,`faelle`.`SchGr`"
End Function ' FüllStr13

Function FüllStr14()
 Str(0, 14, 0) = "aktfaelle"
 Str(1, 14, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `aktfaelle` AS select `f`.`Pat_ID` AS `pid`,`n`.`Notiz` AS `notiz`,`stru`.`Leistung` AS `stru`,`chron`.`Leistung` AS `chron`,`kt`.`ct` AS `kt`,`ebm`.`Leistung` AS `verspau`,`d`.`ICD` AS `icd`,`f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`f4206` AS `f4206`,`f`.`ÜwText` AS `Ü" & _
  "wText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`ÜbwV` AS `ÜbwV`,`f`.`AndÜw` AS `AndÜw`,`f`.`Übw` AS `Übw`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8100` AS `s8100`,`f`.`AktZe" & _
  "it` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`TherArt` AS `TherArt`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos`,`k`.`ID` AS `id`,`k`.`VK` AS `vk`,`k`.`Name` AS `kname`,`k`.`Kateg` AS `kateg`,`k`.`AnzahlIK` AS `anzahlik`,`k`.`AnzahlKTUG` AS `anzahlktug`,`k`.`GültigVon` AS `gültigvon`,`k`.`GültigBis` AS `gültigbis`,`k`.`GO` AS `go`,`k`.`Kurzname` AS `kurzname` from (((((((`faelle` `f` left join `kassenliste` `k` on(((`f`.`VKNr` = `k`.`VK`) and (`f`.`IK` = `k`.`IK`)))) left join `diagnosen` `d` on(((`f`.`Pat_ID` = `d`.`Pat_id`) and (((`d`.`ICD` like 'E1%') and (not((`d`.`ICD` like 'E16%'))) and (not((`d`.`ICD` like 'E15%')))) or (`d`.`ICD` = 'O24.4')) and (`d`.`DiagSicherheit` <> 'A') and ((`d`.`obDauer` = 1) or (`d`.`FID` = `f`.`FID`))))) left join `leistungen` `ebm` on(((`f`.`FID` =" & _
  " `ebm`.`FID`) and ((`ebm`.`Leistung` like '031%') or (`ebm`.`Leistung` like '01210'))))) left join `leistungen` `chron` on(((`f`.`FID` = `chron`.`FID`) and (`chron`.`Leistung` = '03212')))) left join `leistungen` `stru` on(((`f`.`FID` = `stru`.`FID`) and (`stru`.`Leistung` like '973%')))) left join `_kontaktzahl` `kt` on((`kt`.`pat_id` = `f`.`Pat_ID`))) left join `namen` `n` on((`n`.`Pat_ID` = `f`.`Pat_ID`))) where ((`f`.`SchGr` <> '90') and (`f`.`Quartal` = (select concat((((month((now() - interval 14 day)) - 1) DIV 3) + 1),year((now() - interval 14 day))) AS `lq`))) group by `f`.`FID` order by `f`.`Pat_ID`,`f`.`SchGr`,`d`.`ICD`"
End Function ' FüllStr14

Function FüllStr15()
 Str(0, 15, 0) = "aktfaellev"
 Str(1, 15, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `aktfaellev` AS select `aktfaelle`.`pid` AS `pid`,`aktfaelle`.`notiz` AS `notiz`,`aktfaelle`.`stru` AS `stru`,`aktfaelle`.`chron` AS `chron`,`aktfaelle`.`kt` AS `kt`,`aktfaelle`.`verspau` AS `verspau`,`aktfaelle`.`icd` AS `icd`,`aktfaelle`.`FID` AS `FID`,`aktfaelle`.`Pat_ID` AS `Pat_ID`,`aktfaelle`.`Quartal` AS `Quartal`,`aktfaelle`.`Nachname` AS `Nachname`,`aktfaelle`.`Vorname` AS `Vorname`,`aktfaelle`.`lfdnr` AS `lfdnr`,`aktfaelle`.`TMFNr` AS `TMFNr`,`aktfaelle`.`VKNr` AS `VKNr`,`aktfaelle`.`BhFB` AS `BhFB`,`aktfaelle`.`BhFE1` AS `BhFE1`,`aktfaelle`.`BhFE2` AS `BhFE2`,`aktfaelle`.`f4202` AS `f4202`,`aktfaelle`.`ausgst` AS `ausgst`,`aktfaelle`.`KtrAbrB` AS `KtrAbrB`,`aktfaelle`.`AbrAr` AS `AbrAr`,`aktfaelle`.`lVorl` AS `lVorl`,`aktfaelle`.`IK` AS `IK`,`aktfaelle`.`KVKs` AS `KVKs`,`aktfaelle`.`KVKser" & _
  "g` AS `KVKserg`,`aktfaelle`.`Kasse` AS `Kasse`,`aktfaelle`.`GebOr` AS `GebOr`,`aktfaelle`.`AbrGb` AS `AbrGb`,`aktfaelle`.`PersKreis` AS `PersKreis`,`aktfaelle`.`SKtZusatz` AS `SKtZusatz`,`aktfaelle`.`f4206` AS `f4206`,`aktfaelle`.`ÜwText` AS `ÜwText`,`aktfaelle`.`f4210` AS `f4210`,`aktfaelle`.`AkfHAH` AS `AkfHAH`,`aktfaelle`.`AkfAB0` AS `AkfAB0`,`aktfaelle`.`AkfAK` AS `AkfAK`,`aktfaelle`.`statNuller` AS `statNuller`,`aktfaelle`.`ÜbwV` AS `ÜbwV`,`aktfaelle`.`AndÜw` AS `AndÜw`,`aktfaelle`.`Übw` AS `Übw`,`aktfaelle`.`ÜbwLANR` AS `ÜbwLANR`,`aktfaelle`.`ÜWZiel` AS `ÜWZiel`,`aktfaelle`.`ÜWNNr` AS `ÜWNNr`,`aktfaelle`.`ÜWNaN` AS `ÜWNaN`,`aktfaelle`.`ÜWTit` AS `ÜWTit`,`aktfaelle`.`ÜWVor` AS `ÜWVor`,`aktfaelle`.`ÜWVsw` AS `ÜWVsw`,`aktfaelle`.`üwvid` AS `üwvid`,`aktfaelle`.`statKlasse` AS `statKlasse`,`aktfaelle`.`f4237` AS `f4237`,`aktfaelle`.`statBehTage` AS `statBehTage`,`aktfaelle`.`SchGr` AS `" & _
  "SchGr`,`aktfaelle`.`Weiterbeh` AS `Weiterbeh`,`aktfaelle`.`PGeb` AS `PGeb`,`aktfaelle`.`PGebErg` AS `PGebErg`,`aktfaelle`.`Mahnfrist` AS `Mahnfrist`,`aktfaelle`.`GOÄKatNr` AS `GOÄKatNr`,`aktfaelle`.`GOÄKatName` AS `GOÄKatName`,`aktfaelle`.`abrArzt` AS `abrArzt`,`aktfaelle`.`privVers` AS `privVers`,`aktfaelle`.`AdNam` AS `AdNam`,`aktfaelle`.`AdStr` AS `AdStr`,`aktfaelle`.`AdPlz` AS `AdPlz`,`aktfaelle`.`AdOrt` AS `AdOrt`,`aktfaelle`.`BhFE` AS `BhFE`,`aktfaelle`.`s8000` AS `s8000`,`aktfaelle`.`s8100` AS `s8100`,`aktfaelle`.`AktZeit` AS `AktZeit`,`aktfaelle`.`Fanf` AS `Fanf`,`aktfaelle`.`altQuart` AS `altQuart`,`aktfaelle`.`QAnf` AS `QAnf`,`aktfaelle`.`QEnd` AS `QEnd`,`aktfaelle`.`QS` AS `QS`,`aktfaelle`.`QT` AS `QT`,`aktfaelle`.`TherArt` AS `TherArt`,`aktfaelle`.`StByte` AS `StByte`,`aktfaelle`.`absPos` AS `absPos`,`aktfaelle`.`id` AS `id`,`aktfaelle`.`vk` AS `vk`,`aktfaelle`.`kname` AS `kn" & _
  "ame`,`aktfaelle`.`kateg` AS `kateg`,`aktfaelle`.`anzahlik` AS `anzahlik`,`aktfaelle`.`anzahlktug` AS `anzahlktug`,`aktfaelle`.`gültigvon` AS `gültigvon`,`aktfaelle`.`gültigbis` AS `gültigbis`,`aktfaelle`.`go` AS `go`,`aktfaelle`.`kurzname` AS `kurzname` from `aktfaelle` group by `aktfaelle`.`Pat_ID`"
End Function ' FüllStr15

Function FüllStr16()
 Str(0, 16, 0) = "aktfv"
 Str(1, 16, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `aktfv` AS select `faelle`.`Pat_ID` AS `pat_id`,`faelle`.`FID` AS `fid`,`faelle`.`SchGr` AS `schgr`,`faelle`.`VKNr` AS `vknr` from `faelle` where ((`faelle`.`SchGr` <> '90') and (`faelle`.`Quartal` = (select concat((((month((now() - interval 14 day)) - 1) DIV 3) + 1),year((now() - interval 14 day))) AS `lq`))) group by `faelle`.`Pat_ID` order by `faelle`.`Pat_ID`,`faelle`.`SchGr`"
End Function ' FüllStr16

Function FüllStr17()
 Str(0, 17, 0) = "anamnesebogen"
 Str(0, 17, 1) = "`Prim`"
 Str(0, 17, 2) = "`Pat_id`"
 Str(0, 17, 3) = "`Nachname`"
 Str(0, 17, 4) = "`Vorname`"
 Str(0, 17, 5) = "`NVorsatz`"
 Str(0, 17, 6) = "`Titel`"
 Str(0, 17, 7) = "`Anrede`"
 Str(0, 17, 8) = "`GebDat`"
 Str(0, 17, 9) = "`Tkz`"
 Str(0, 17, 10) = "`Versicherungsart`"
 Str(0, 17, 11) = "`Diabetestyp`"
 Str(0, 17, 12) = "`Diabetes seit`"
 Str(0, 17, 13) = "`Tabletten seit`"
 Str(0, 17, 14) = "`Insulin seit`"
 Str(0, 17, 15) = "`Grund für Vorstellung`"
 Str(0, 17, 16) = "`Familienanamnese`"
 Str(0, 17, 17) = "`Größe`"
 Str(0, 17, 18) = "`Gewicht`"
 Str(0, 17, 19) = "`bmi`"
 Str(0, 17, 20) = "`Tendenz`"
 Str(0, 17, 21) = "`DiabetesMedikament 1`"
 Str(0, 17, 22) = "`DiabetesMedikament 1 Menge`"
 Str(0, 17, 23) = "`DiabetesMedikament 2`"
 Str(0, 17, 24) = "`DiabetesMedikament 2 Menge`"
 Str(0, 17, 25) = "`DiabetesMedikament 3`"
 Str(0, 17, 26) = "`DiabetesMedikament 3 Menge`"
 Str(0, 17, 27) = "`DiabetesMedikament 4`"
 Str(0, 17, 28) = "`DiabetesMedikament 4 Menge`"
 Str(0, 17, 29) = "`Insulinpumpe`"
 Str(0, 17, 30) = "`Insulinpumpe seit`"
 Str(0, 17, 31) = "`Insulinpumpe Marke`"
 Str(0, 17, 32) = "`Broteinheiten gesamt`"
 Str(0, 17, 33) = "`Broteinheiten früh`"
 Str(0, 17, 34) = "`Broteinheiten ZM früh`"
 Str(0, 17, 35) = "`Broteinheiten mittags`"
 Str(0, 17, 36) = "`Broteinheiten nachmittags`"
 Str(0, 17, 37) = "`Broteinheiten abends`"
 Str(0, 17, 38) = "`Broteinheiten nachts`"
 Str(0, 17, 39) = "`Essenszeit früh`"
 Str(0, 17, 40) = "`Essenszeit vormittags`"
 Str(0, 17, 41) = "`Essenszeit mittags`"
 Str(0, 17, 42) = "`Essenszeit nachmittags`"
 Str(0, 17, 43) = "`Essenszeit abends`"
 Str(0, 17, 44) = "`Essenszeit spät`"
 Str(0, 17, 45) = "`Spritz-Eß-Abstand früh`"
 Str(0, 17, 46) = "`Spritz-Eß-Abstand mittags`"
 Str(0, 17, 47) = "`Spritz-Eß-Abstand abends`"
 Str(0, 17, 48) = "`Spritzstelle früh`"
 Str(0, 17, 49) = "`Spritzstelle mittags`"
 Str(0, 17, 50) = "`Spritzstelle abends`"
 Str(0, 17, 51) = "`Spritzstelle nachts`"
 Str(0, 17, 52) = "`Jahr letzte Diabetesschulung`"
 Str(0, 17, 53) = "`Ort Schulung`"
 Str(0, 17, 54) = "`letztes HbA1c`"
 Str(0, 17, 55) = "`gemessen am`"
 Str(0, 17, 56) = "`vorherige Werte`"
 Str(0, 17, 57) = "`BZMessungen selbst`"
 Str(0, 17, 58) = "`Gerät`"
 Str(0, 17, 59) = "`BZMessungen pW`"
 Str(0, 17, 60) = "`BZMessungen pW ndE`"
 Str(0, 17, 61) = "`BZMessungen p W nachts`"
 Str(0, 17, 62) = "`Aufschreiben`"
 Str(0, 17, 63) = "`BZWerte v d Essen`"
 Str(0, 17, 64) = "`BZWerte n d Essen`"
 Str(0, 17, 65) = "`UZ Tageszeit`"
 Str(0, 17, 66) = "`Unterzucker pM`"
 Str(0, 17, 67) = "`UZ rechtzeitig`"
 Str(0, 17, 68) = "`Fremde Hilfe pa`"
 Str(0, 17, 69) = "`Bewußtlos pa`"
 Str(0, 17, 70) = "`Keto pa`"
 Str(0, 17, 71) = "`BZgr300 pM`"
 Str(0, 17, 72) = "`Bluthochdruck`"
 Str(0, 17, 73) = "`BHD seit`"
 Str(0, 17, 74) = "`BHD beh mit`"
 Str(0, 17, 75) = "`Blutdruckwerte`"
 Str(0, 17, 76) = "`BDselbst`"
 Str(0, 17, 77) = "`Schwanger`"
 Str(0, 17, 78) = "`Schwanger seit`"
 Str(0, 17, 79) = "`Augensp zuletzt`"
 Str(0, 17, 80) = "`Augensp Befund`"
 Str(0, 17, 81) = "`Netzhaut gelasert`"
 Str(0, 17, 82) = "`Sehminderung unbehebbar`"
 Str(0, 17, 83) = "`Diabet Nierenschaden`"
 Str(0, 17, 84) = "`Albumin zuletzt`"
 Str(0, 17, 85) = "`erhöht?`"
 Str(0, 17, 86) = "`Dialyse`"
 Str(0, 17, 87) = "`Dialyse seit`"
 Str(0, 17, 88) = "`andere Nierenerkrankung`"
 Str(0, 17, 89) = "`Herzkrankheit`"
 Str(0, 17, 90) = "`Angina pectoris`"
 Str(0, 17, 91) = "`Herzinfarkt`"
 Str(0, 17, 92) = "`Herzinfarkt wann`"
 Str(0, 17, 93) = "`PTCA oder Stent`"
 Str(0, 17, 94) = "`Bypass kardial`"
 Str(0, 17, 95) = "`Bypass wann`"
 Str(0, 17, 96) = "`Herzschwäche`"
 Str(0, 17, 97) = "`Herzkrankheit Beschreibung`"
 Str(0, 17, 98) = "`Hirndurchblutungsstörung`"
 Str(0, 17, 99) = "`Schlaganfall`"
 Str(0, 17, 100) = "`Beindurchblutungsstörung`"
 Str(0, 17, 101) = "`Schaufensterkrankheit`"
 Str(0, 17, 102) = "`Bypaß peripher`"
 Str(0, 17, 103) = "`Geschwür`"
 Str(0, 17, 104) = "`Amputation`"
 Str(0, 17, 105) = "`pAVK Beschreibung`"
 Str(0, 17, 106) = "`Ameisenlaufen`"
 Str(0, 17, 107) = "`Ameisen Ausmaß`"
 Str(0, 17, 108) = "`Druckstellen`"
 Str(0, 17, 109) = "`Verformungen`"
 Str(0, 17, 110) = "`Verformungen Beschreibung`"
 Str(0, 17, 111) = "`Fußpflege`"
 Str(0, 17, 112) = "`Podologie`"
 Str(0, 17, 113) = "`Einlagen`"
 Str(0, 17, 114) = "`Neue Fußkomplikationen`"
 Str(0, 17, 115) = "`Entleerungsstörungen Magen`"
 Str(0, 17, 116) = "`Entleerungsstörungen Harnblase`"
 Str(0, 17, 117) = "`Schwindel Aufstehen`"
 Str(0, 17, 118) = "`Folgeerkrankungen Haut`"
 Str(0, 17, 119) = "`Bewegungseinschränkungen`"
 Str(0, 17, 120) = "`Sexualstörung`"
 Str(0, 17, 121) = "`Sexualstörung seit`"
 Str(0, 17, 122) = "`Weitere Anamnese`"
 Str(0, 17, 123) = "`Alkohol`"
 Str(0, 17, 124) = "`Tabak`"
 Str(0, 17, 125) = "`tabakex`"
 Str(0, 17, 126) = "`tabakbis`"
 Str(0, 17, 127) = "`tabakakt`"
 Str(0, 17, 128) = "`tabakmenge`"
 Str(0, 17, 129) = "`Weitere Medikation`"
 Str(0, 17, 130) = "`Liphypertrophien Abdomen`"
 Str(0, 17, 131) = "`Liphypertrophien Beine`"
 Str(0, 17, 132) = "`Liphypertrophien Arme`"
 Str(0, 17, 133) = "`Beinbefund`"
 Str(0, 17, 134) = "`Hyperkeratosen`"
 Str(0, 17, 135) = "`Ulcera`"
 Str(0, 17, 136) = "`Kraft Zehenheber`"
 Str(0, 17, 137) = "`Kraft Zehenbeuger`"
 Str(0, 17, 138) = "`Kraft Knie`"
 Str(0, 17, 139) = "`ASR`"
 Str(0, 17, 140) = "`PSR`"
 Str(0, 17, 141) = "`Oberflächensensibilität`"
 Str(0, 17, 142) = "`Monofilamenttest`"
 Str(0, 17, 143) = "`Kalt-Warm`"
 Str(0, 17, 144) = "`Vibration IK`"
 Str(0, 17, 145) = "`Vibration Großzehe`"
 Str(0, 17, 146) = "`Puls Leiste`"
 Str(0, 17, 147) = "`Puls Kniekehle`"
 Str(0, 17, 148) = "`Puls Atp`"
 Str(0, 17, 149) = "`Puls Adp`"
 Str(0, 17, 150) = "`RR`"
 Str(0, 17, 151) = "`RRTurboMed`"
 Str(0, 17, 152) = "`Herz`"
 Str(0, 17, 153) = "`Lunge`"
 Str(0, 17, 154) = "`Bauch`"
 Str(0, 17, 155) = "`WS`"
 Str(0, 17, 156) = "`NL`"
 Str(0, 17, 157) = "`SD`"
 Str(0, 17, 158) = "`Carotiden`"
 Str(0, 17, 159) = "`NNH`"
 Str(0, 17, 160) = "`Zähne`"
 Str(0, 17, 161) = "`Mundhöhle`"
 Str(0, 17, 162) = "`LK`"
 Str(0, 17, 163) = "`BeinödVen`"
 Str(0, 17, 164) = "`Neuro sonst`"
 Str(0, 17, 165) = "`Weitere Befunde`"
 Str(0, 17, 166) = "`Schulung`"
 Str(0, 17, 167) = "`DMP`"
 Str(0, 17, 168) = "`DMSchulz`"
 Str(0, 17, 169) = "`DMSchL`"
 Str(0, 17, 170) = "`RRSchulz`"
 Str(0, 17, 171) = "`DMPhier`"
 Str(0, 17, 172) = "`HANr`"
 Str(0, 17, 173) = "`HANr2`"
 Str(0, 17, 174) = "`letzte Änderung`"
 Str(0, 17, 175) = "`Diagnosen`"
 Str(0, 17, 176) = "`Vorgestellt`"
 Str(0, 17, 177) = "`Versicherung`"
 Str(0, 17, 178) = "`AktZeit`"
 Str(0, 17, 179) = "`Ther1`"
 Str(0, 17, 180) = "`TherAkt`"
 Str(0, 17, 181) = "`obAn1eing`"
 Str(0, 17, 182) = "`obAn2eing`"
 Str(0, 17, 183) = "`obAnAeing`"
 Str(0, 17, 184) = "`obCheck`"
 Str(0, 17, 185) = "`obBZausgew`"
 Str(0, 17, 186) = "`obOSaufgek`"
 Str(0, 17, 187) = "`obPodAufgek`"
 Str(0, 17, 188) = "`obMBlAusgeh`"
 Str(0, 17, 189) = "`obSchulaufgek`"
 Str(0, 17, 190) = "`obDMPaufgekl`"
 Str(0, 17, 191) = "`obMedNetz`"
 Str(0, 17, 192) = "`Hausarzt`"
 Str(0, 17, 193) = "`ob`"
 Str(0, 17, 194) = "`QS`"
 Str(0, 17, 195) = "`QT`"
 Str(0, 17, 196) = "`Prim`"
 Str(0, 17, 197) = "`PrimaryKey`"
 Str(0, 17, 198) = "`Auswahl`"
 Str(0, 17, 199) = "`DMPhier`"
 Str(0, 17, 200) = "`GebDat`"
 Str(0, 17, 201) = "`jlD`"
 Str(0, 17, 202) = "`lÄnd`"
 Str(0, 17, 203) = "`Pat_ID`"
 Str(0, 17, 204) = "`Ther1`"
 Str(0, 17, 205) = "`Vorgestellt`"
 Str(0, 17, 206) = "`HausärzteAnamnesebogen_AccRel`"
 Str(0, 17, 207) = "`KassenlisteAnamnesebogen_AccRel`"
 Str(0, 17, 208) = "`HausärzteAnamnesebogen_AccRel`"
 Str(0, 17, 209) = "`KassenlisteAnamnesebogen_AccRel`"
 Str(0, 17, 210) = "`NamenAnamnesebogen_AccRel`"
 ArtZ(0, 17) = 195
 ArtZ(1, 17) = 12
 ArtZ(2, 17) = 3
 Str(1, 17, 0) = "CREATE TABLE `anamnesebogen` ("
 Str(1, 17, 1) = " `Prim` int(10) unsigned NOT NULL COMMENT 'Primärschlüssel'"
 Str(1, 17, 2) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 17, 3) = " `Nachname` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '-'"
 Str(1, 17, 4) = " `Vorname` varchar(19) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 5) = " `NVorsatz` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 6) = " `Titel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 7) = " `Anrede` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 8) = " `GebDat` datetime DEFAULT NULL COMMENT ', geb.'"
 Str(1, 17, 9) = " `Tkz` tinyint(1) unsigned DEFAULT NULL COMMENT 'Tod-Kennzeichen'"
 Str(1, 17, 10) = " `Versicherungsart` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 11) = " `Diabetestyp` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetes Typ'"
 Str(1, 17, 12) = " `Diabetes seit` varchar(152) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 17, 13) = " `Tabletten seit` varchar(66) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Tabletten seit'"
 Str(1, 17, 14) = " `Insulin seit` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Insulin seit'"
 Str(1, 17, 15) = " `Grund für Vorstellung` varchar(721) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 16) = " `Familienanamnese` varchar(291) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 17) = " `Größe` double DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 18) = " `Gewicht` double DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 19) = " `bmi` decimal(5,1) DEFAULT '0.0'"
 Str(1, 17, 20) = " `Tendenz` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Tendenz'"
 Str(1, 17, 21) = " `DiabetesMedikament 1` varchar(48) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesmedikation:'"
 Str(1, 17, 22) = " `DiabetesMedikament 1 Menge` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 17, 23) = " `DiabetesMedikament 2` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 17, 24) = " `DiabetesMedikament 2 Menge` varchar(126) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 17, 25) = " `DiabetesMedikament 3` varchar(37) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 17, 26) = " `DiabetesMedikament 3 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 17, 27) = " `DiabetesMedikament 4` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 17, 28) = " `DiabetesMedikament 4 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 17, 29) = " `Insulinpumpe` tinyint(1) unsigned DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 30) = " `Insulinpumpe seit` varchar(250) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 17, 31) = " `Insulinpumpe Marke` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Marke:'"
 Str(1, 17, 32) = " `Broteinheiten gesamt` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Broteinheiten:gesamt'"
 Str(1, 17, 33) = " `Broteinheiten früh` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, früh'"
 Str(1, 17, 34) = " `Broteinheiten ZM früh` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zwischenmahlzeit vormittags'"
 Str(1, 17, 35) = " `Broteinheiten mittags` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 17, 36) = " `Broteinheiten nachmittags` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 17, 37) = " `Broteinheiten abends` varchar(13) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 17, 38) = " `Broteinheiten nachts` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 17, 39) = " `Essenszeit früh` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Essenszeiten:früh'"
 Str(1, 17, 40) = " `Essenszeit vormittags` varchar(34) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vormittags'"
 Str(1, 17, 41) = " `Essenszeit mittags` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 17, 42) = " `Essenszeit nachmittags` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 17, 43) = " `Essenszeit abends` varchar(18) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 17, 44) = " `Essenszeit spät` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, spät'"
 Str(1, 17, 45) = " `Spritz-Eß-Abstand früh` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritz-Eß-Abstand:früh'"
 Str(1, 17, 46) = " `Spritz-Eß-Abstand mittags` varchar(27) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 17, 47) = " `Spritz-Eß-Abstand abends` varchar(62) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 17, 48) = " `Spritzstelle früh` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritzstellen:früh'"
 Str(1, 17, 49) = " `Spritzstelle mittags` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 17, 50) = " `Spritzstelle abends` varchar(139) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 17, 51) = " `Spritzstelle nachts` varchar(141) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 17, 52) = " `Jahr letzte Diabetesschulung` varchar(129) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesschulung:'"
 Str(1, 17, 53) = " `Ort Schulung` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<in'"
 Str(1, 17, 54) = " `letztes HbA1c` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letztes HbA1c:'"
 Str(1, 17, 55) = " `gemessen am` datetime DEFAULT NULL COMMENT '<, gemessen'"
 Str(1, 17, 56) = " `vorherige Werte` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vorher:'"
 Str(1, 17, 57) = " `BZMessungen selbst` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckermessung:Selbstmessung?'"
 Str(1, 17, 58) = " `Gerät` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<:'"
 Str(1, 17, 59) = " `BZMessungen pW` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl d.Messungen pro Woche:'"
 Str(1, 17, 60) = " `BZMessungen pW ndE` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, davon nach dem Essen:'"
 Str(1, 17, 61) = " `BZMessungen p W nachts` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts:'"
 Str(1, 17, 62) = " `Aufschreiben` varchar(93) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Dokumentation:'"
 Str(1, 17, 63) = " `BZWerte v d Essen` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckerwerte vor dem Essen:'"
 Str(1, 17, 64) = " `BZWerte n d Essen` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nach dem Essen:'"
 Str(1, 17, 65) = " `UZ Tageszeit` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Unterzucker:Bevorzugte Tages-/Uhrzeit'"
 Str(1, 17, 66) = " `Unterzucker pM` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl der schweren (<50 mg/dl) pro Monat:'"
 Str(1, 17, 67) = " `UZ rechtzeitig` varchar(101) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, rechtzeitig bemerkt:'"
 Str(1, 17, 68) = " `Fremde Hilfe pa` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, fremde Hilfe deshalb nötig:'"
 Str(1, 17, 69) = " `Bewußtlos pa` varchar(81) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, bewußtlos deshalb:'"
 Str(1, 17, 70) = " `Keto pa` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Zahl der Ketoazidosen pro Jahr:'"
 Str(1, 17, 71) = " `BZgr300 pM` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Zahl der Blutzucker > 300 mg/dl pro Monat:'"
 Str(1, 17, 72) = " `Bluthochdruck` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Bluthochdruck:'"
 Str(1, 17, 73) = " `BHD seit` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit:'"
 Str(1, 17, 74) = " `BHD beh mit` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, behandelt mit:'"
 Str(1, 17, 75) = " `Blutdruckwerte` varchar(186) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckwerte:'"
 Str(1, 17, 76) = " `BDselbst` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckselbstmessung:'"
 Str(1, 17, 77) = " `Schwanger` varchar(43) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Aktuelle Schwangerschaft:'"
 Str(1, 17, 78) = " `Schwanger seit` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, seit:'"
 Str(1, 17, 79) = " `Augensp zuletzt` varchar(107) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Augenspiegelung:'"
 Str(1, 17, 80) = " `Augensp Befund` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 17, 81) = " `Netzhaut gelasert` varchar(142) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Netzhaut schon gelasert:'"
 Str(1, 17, 82) = " `Sehminderung unbehebbar` varchar(157) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', mit Brille nicht behebbare Sehminderung:'"
 Str(1, 17, 83) = " `Diabet Nierenschaden` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetischer Nierenschaden:'"
 Str(1, 17, 84) = " `Albumin zuletzt` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', letztes Albumin:'"
 Str(1, 17, 85) = " `erhöht?` varchar(52) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 17, 86) = " `Dialyse` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 87) = " `Dialyse seit` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 17, 88) = " `andere Nierenerkrankung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', andere Nierenerkrankung:'"
 Str(1, 17, 89) = " `Herzkrankheit` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Herzkrankheit:'"
 Str(1, 17, 90) = " `Angina pectoris` varchar(134) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 91) = " `Herzinfarkt` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 92) = " `Herzinfarkt wann` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 17, 93) = " `PTCA oder Stent` varchar(94) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 94) = " `Bypass kardial` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 95) = " `Bypass wann` varchar(388) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 17, 96) = " `Herzschwäche` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 97) = " `Herzkrankheit Beschreibung` varchar(213) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung:'"
 Str(1, 17, 98) = " `Hirndurchblutungsstörung` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 99) = " `Schlaganfall` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 100) = " `Beindurchblutungsstörung` varchar(136) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 101) = " `Schaufensterkrankheit` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 102) = " `Bypaß peripher` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 103) = " `Geschwür` varchar(174) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 104) = " `Amputation` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 105) = " `pAVK Beschreibung` varchar(97) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung der Beinbeschwerden:'"
 Str(1, 17, 106) = " `Ameisenlaufen` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 107) = " `Ameisen Ausmaß` varchar(123) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Ausmaß:'"
 Str(1, 17, 108) = " `Druckstellen` varchar(173) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 109) = " `Verformungen` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 110) = " `Verformungen Beschreibung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Beschreibung:'"
 Str(1, 17, 111) = " `Fußpflege` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 112) = " `Podologie` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 113) = " `Einlagen` varchar(91) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', diabetesgerechte orthopädische Einlagen/Schuhe:'"
 Str(1, 17, 114) = " `Neue Fußkomplikationen` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Neue Fußkomplikationen in den letzten 12 Monaten:'"
 Str(1, 17, 115) = " `Entleerungsstörungen Magen` varchar(130) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 116) = " `Entleerungsstörungen Harnblase` varchar(157) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 117) = " `Schwindel Aufstehen` varchar(121) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 118) = " `Folgeerkrankungen Haut` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 119) = " `Bewegungseinschränkungen` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 120) = " `Sexualstörung` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 121) = " `Sexualstörung seit` varchar(149) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 17, 122) = " `Weitere Anamnese` varchar(989) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 123) = " `Alkohol` varchar(148) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 124) = " `Tabak` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 125) = " `tabakex` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 126) = " `tabakbis` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 127) = " `tabakakt` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 128) = " `tabakmenge` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 129) = " `Weitere Medikation` varchar(298) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 130) = " `Liphypertrophien Abdomen` varchar(176) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Liphypertrophien:Abdomen'"
 Str(1, 17, 131) = " `Liphypertrophien Beine` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Beine:'"
 Str(1, 17, 132) = " `Liphypertrophien Arme` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Arme:'"
 Str(1, 17, 133) = " `Beinbefund` varchar(272) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 134) = " `Hyperkeratosen` varchar(207) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 135) = " `Ulcera` varchar(103) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 136) = " `Kraft Zehenheber` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Kraft:Zehenheber'"
 Str(1, 17, 137) = " `Kraft Zehenbeuger` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zehenbeuger:'"
 Str(1, 17, 138) = " `Kraft Knie` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Knie:'"
 Str(1, 17, 139) = " `ASR` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 140) = " `PSR` varchar(67) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 141) = " `Oberflächensensibilität` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 142) = " `Monofilamenttest` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 143) = " `Kalt-Warm` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Kalt-Warm-Diskrimination:'"
 Str(1, 17, 144) = " `Vibration IK` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Vibrationsempfinden Innenknöchel:'"
 Str(1, 17, 145) = " `Vibration Großzehe` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Großzehe:'"
 Str(1, 17, 146) = " `Puls Leiste` varchar(33) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Pulse:Leiste'"
 Str(1, 17, 147) = " `Puls Kniekehle` varchar(26) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Kniekehle:'"
 Str(1, 17, 148) = " `Puls Atp` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Innenknöchel:'"
 Str(1, 17, 149) = " `Puls Adp` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Fußrücken:'"
 Str(1, 17, 150) = " `RR` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruck:'"
 Str(1, 17, 151) = " `RRTurboMed` varchar(1362) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 152) = " `Herz` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 17, 153) = " `Lunge` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 154) = " `Bauch` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Abdomen:'"
 Str(1, 17, 155) = " `WS` varchar(56) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Wirbelsäule:'"
 Str(1, 17, 156) = " `NL` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nierenlager:'"
 Str(1, 17, 157) = " `SD` varchar(88) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Schilddrüse:'"
 Str(1, 17, 158) = " `Carotiden` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Halsschlagadern:'"
 Str(1, 17, 159) = " `NNH` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nasennebenhöhlen:'"
 Str(1, 17, 160) = " `Zähne` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 161) = " `Mundhöhle` varchar(64) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 17, 162) = " `LK` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Lymphknoten:'"
 Str(1, 17, 163) = " `BeinödVen` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beinödeme/ Venenkrankheiten:'"
 Str(1, 17, 164) = " `Neuro sonst` varchar(74) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Sonstige neurologische Befunde:'"
 Str(1, 17, 165) = " `Weitere Befunde` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', weitere Befunde:'"
 Str(1, 17, 166) = " `Schulung` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Schulungsbedarf'"
 Str(1, 17, 167) = " `DMP` varchar(85) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Pat. bei HA im DMP'"
 Str(1, 17, 168) = " `DMSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der DMP-Schulungen hier'"
 Str(1, 17, 169) = " `DMSchL` smallint(6) DEFAULT NULL COMMENT 'Zahl der abgerechneten DMP-Schulungen hier'"
 Str(1, 17, 170) = " `RRSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der Hypertonie-Schulungen hier'"
 Str(1, 17, 171) = " `DMPhier` datetime DEFAULT NULL COMMENT 'ob Pat hier im DMP'"
 Str(1, 17, 172) = " `HANr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 17, 173) = " `HANr2` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 17, 174) = " `letzte Änderung` datetime DEFAULT NULL COMMENT 'Datum der letzten Änderung'"
 Str(1, 17, 175) = " `Diagnosen` varchar(1071) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 176) = " `Vorgestellt` datetime DEFAULT NULL COMMENT 'Erstvorstellung'"
 Str(1, 17, 177) = " `Versicherung` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 178) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 17, 179) = " `Ther1` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, ICT, CSII'"
 Str(1, 17, 180) = " `TherAkt` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, ICT, CSII'"
 Str(1, 17, 181) = " `obAn1eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 1 eingegeben wurde'"
 Str(1, 17, 182) = " `obAn2eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 2 eingegeben wurde'"
 Str(1, 17, 183) = " `obAnAeing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt allgemein eingegeben wurde'"
 Str(1, 17, 184) = " `obCheck` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Checkliste vorliegt'"
 Str(1, 17, 185) = " `obBZausgew` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Blutzuckergerät ausgewechselt'"
 Str(1, 17, 186) = " `obOSaufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über orthopäd Schuhmacher aufgeklärt'"
 Str(1, 17, 187) = " `obPodAufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 17, 188) = " `obMBlAusgeh` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 17, 189) = " `obSchulaufgek` varchar(14) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 17, 190) = " `obDMPaufgekl` varchar(17) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 17, 191) = " `obMedNetz` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob von Med. Netz geschickt'"
 Str(1, 17, 192) = " `Hausarzt` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Hausarzt laut Anamnesebogen'"
 Str(1, 17, 193) = " `ob` tinyint(1) unsigned DEFAULT NULL COMMENT 'für verschiedene Aktionen'"
 Str(1, 17, 194) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 17, 195) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 17, 196) = "  PRIMARY KEY (`Prim`)"
 Str(1, 17, 197) = "  UNIQUE KEY `PrimaryKey` (`Prim`)"
 Str(1, 17, 198) = "  KEY `Auswahl` (`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 17, 199) = "  KEY `DMPhier` (`DMPhier`)"
 Str(1, 17, 200) = "  KEY `GebDat` (`GebDat`)"
 Str(1, 17, 201) = "  KEY `jlD` (`Jahr letzte Diabetesschulung`,`GebDat`)"
 Str(1, 17, 202) = "  KEY `lÄnd` (`letzte Änderung`)"
 Str(1, 17, 203) = "  KEY `Pat_ID` (`Pat_id`)"
 Str(1, 17, 204) = "  KEY `Ther1` (`Ther1`)"
 Str(1, 17, 205) = "  KEY `Vorgestellt` (`Vorgestellt`,`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 17, 206) = "  KEY `HausärzteAnamnesebogen_AccRel` (`HANr`)"
 Str(1, 17, 207) = "  KEY `KassenlisteAnamnesebogen_AccRel` (`Versicherung`)"
 Str(1, 17, 208) = "  CONSTRAINT `HausärzteAnamnesebogen_AccRel` FOREIGN KEY (`HANr`) REFERENCES `hausaerzte` (`KVNr`)"
 Str(1, 17, 209) = "  CONSTRAINT `KassenlisteAnamnesebogen_AccRel` FOREIGN KEY (`Versicherung`) REFERENCES `kassenliste` (`VK`)"
 Str(1, 17, 210) = "  CONSTRAINT `NamenAnamnesebogen_AccRel` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`)"
 Str(1, 17, 211) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Function ' FüllStr17

Function FüllStr18()
 Str(0, 18, 0) = "anbogalt"
 Str(0, 18, 1) = "`Prim`"
 Str(0, 18, 2) = "`Pat_id`"
 Str(0, 18, 3) = "`Nachname`"
 Str(0, 18, 4) = "`Vorname`"
 Str(0, 18, 5) = "`NVorsatz`"
 Str(0, 18, 6) = "`Titel`"
 Str(0, 18, 7) = "`Anrede`"
 Str(0, 18, 8) = "`GebDat`"
 Str(0, 18, 9) = "`Tkz`"
 Str(0, 18, 10) = "`Versicherungsart`"
 Str(0, 18, 11) = "`Diabetestyp`"
 Str(0, 18, 12) = "`Diabetes seit`"
 Str(0, 18, 13) = "`Tabletten seit`"
 Str(0, 18, 14) = "`Insulin seit`"
 Str(0, 18, 15) = "`Grund für Vorstellung`"
 Str(0, 18, 16) = "`Familienanamnese`"
 Str(0, 18, 17) = "`Größe`"
 Str(0, 18, 18) = "`Gewicht`"
 Str(0, 18, 19) = "`Tendenz`"
 Str(0, 18, 20) = "`DiabetesMedikament 1`"
 Str(0, 18, 21) = "`DiabetesMedikament 1 Menge`"
 Str(0, 18, 22) = "`DiabetesMedikament 2`"
 Str(0, 18, 23) = "`DiabetesMedikament 2 Menge`"
 Str(0, 18, 24) = "`DiabetesMedikament 3`"
 Str(0, 18, 25) = "`DiabetesMedikament 3 Menge`"
 Str(0, 18, 26) = "`DiabetesMedikament 4`"
 Str(0, 18, 27) = "`DiabetesMedikament 4 Menge`"
 Str(0, 18, 28) = "`Insulinpumpe`"
 Str(0, 18, 29) = "`Insulinpumpe seit`"
 Str(0, 18, 30) = "`Insulinpumpe Marke`"
 Str(0, 18, 31) = "`Broteinheiten gesamt`"
 Str(0, 18, 32) = "`Broteinheiten früh`"
 Str(0, 18, 33) = "`Broteinheiten ZM früh`"
 Str(0, 18, 34) = "`Broteinheiten mittags`"
 Str(0, 18, 35) = "`Broteinheiten nachmittags`"
 Str(0, 18, 36) = "`Broteinheiten abends`"
 Str(0, 18, 37) = "`Broteinheiten nachts`"
 Str(0, 18, 38) = "`Essenszeit früh`"
 Str(0, 18, 39) = "`Essenszeit vormittags`"
 Str(0, 18, 40) = "`Essenszeit mittags`"
 Str(0, 18, 41) = "`Essenszeit nachmittags`"
 Str(0, 18, 42) = "`Essenszeit abends`"
 Str(0, 18, 43) = "`Essenszeit spät`"
 Str(0, 18, 44) = "`Spritz-Eß-Abstand früh`"
 Str(0, 18, 45) = "`Spritz-Eß-Abstand mittags`"
 Str(0, 18, 46) = "`Spritz-Eß-Abstand abends`"
 Str(0, 18, 47) = "`Spritzstelle früh`"
 Str(0, 18, 48) = "`Spritzstelle mittags`"
 Str(0, 18, 49) = "`Spritzstelle abends`"
 Str(0, 18, 50) = "`Spritzstelle nachts`"
 Str(0, 18, 51) = "`Jahr letzte Diabetesschulung`"
 Str(0, 18, 52) = "`Ort Schulung`"
 Str(0, 18, 53) = "`letztes HbA1c`"
 Str(0, 18, 54) = "`gemessen am`"
 Str(0, 18, 55) = "`vorherige Werte`"
 Str(0, 18, 56) = "`BZMessungen selbst`"
 Str(0, 18, 57) = "`Gerät`"
 Str(0, 18, 58) = "`BZMessungen pW`"
 Str(0, 18, 59) = "`BZMessungen pW ndE`"
 Str(0, 18, 60) = "`BZMessungen p W nachts`"
 Str(0, 18, 61) = "`Aufschreiben`"
 Str(0, 18, 62) = "`BZWerte v d Essen`"
 Str(0, 18, 63) = "`BZWerte n d Essen`"
 Str(0, 18, 64) = "`UZ Tageszeit`"
 Str(0, 18, 65) = "`Unterzucker pM`"
 Str(0, 18, 66) = "`UZ rechtzeitig`"
 Str(0, 18, 67) = "`Fremde Hilfe pa`"
 Str(0, 18, 68) = "`Bewußtlos pa`"
 Str(0, 18, 69) = "`Keto pa`"
 Str(0, 18, 70) = "`BZgr300 pM`"
 Str(0, 18, 71) = "`Bluthochdruck`"
 Str(0, 18, 72) = "`BHD seit`"
 Str(0, 18, 73) = "`BHD beh mit`"
 Str(0, 18, 74) = "`Blutdruckwerte`"
 Str(0, 18, 75) = "`BDselbst`"
 Str(0, 18, 76) = "`Schwanger`"
 Str(0, 18, 77) = "`Schwanger seit`"
 Str(0, 18, 78) = "`Augensp zuletzt`"
 Str(0, 18, 79) = "`Augensp Befund`"
 Str(0, 18, 80) = "`Netzhaut gelasert`"
 Str(0, 18, 81) = "`Sehminderung unbehebbar`"
 Str(0, 18, 82) = "`Diabet Nierenschaden`"
 Str(0, 18, 83) = "`Albumin zuletzt`"
 Str(0, 18, 84) = "`erhöht?`"
 Str(0, 18, 85) = "`Dialyse`"
 Str(0, 18, 86) = "`Dialyse seit`"
 Str(0, 18, 87) = "`andere Nierenerkrankung`"
 Str(0, 18, 88) = "`Herzkrankheit`"
 Str(0, 18, 89) = "`Angina pectoris`"
 Str(0, 18, 90) = "`Herzinfarkt`"
 Str(0, 18, 91) = "`Herzinfarkt wann`"
 Str(0, 18, 92) = "`PTCA oder Stent`"
 Str(0, 18, 93) = "`Bypass kardial`"
 Str(0, 18, 94) = "`Bypass wann`"
 Str(0, 18, 95) = "`Herzschwäche`"
 Str(0, 18, 96) = "`Herzkrankheit Beschreibung`"
 Str(0, 18, 97) = "`Hirndurchblutungsstörung`"
 Str(0, 18, 98) = "`Schlaganfall`"
 Str(0, 18, 99) = "`Beindurchblutungsstörung`"
 Str(0, 18, 100) = "`Schaufensterkrankheit`"
 Str(0, 18, 101) = "`Bypaß peripher`"
 Str(0, 18, 102) = "`Geschwür`"
 Str(0, 18, 103) = "`Amputation`"
 Str(0, 18, 104) = "`pAVK Beschreibung`"
 Str(0, 18, 105) = "`Ameisenlaufen`"
 Str(0, 18, 106) = "`Ameisen Ausmaß`"
 Str(0, 18, 107) = "`Druckstellen`"
 Str(0, 18, 108) = "`Verformungen`"
 Str(0, 18, 109) = "`Verformungen Beschreibung`"
 Str(0, 18, 110) = "`Fußpflege`"
 Str(0, 18, 111) = "`Podologie`"
 Str(0, 18, 112) = "`Einlagen`"
 Str(0, 18, 113) = "`Neue Fußkomplikationen`"
 Str(0, 18, 114) = "`Entleerungsstörungen Magen`"
 Str(0, 18, 115) = "`Entleerungsstörungen Harnblase`"
 Str(0, 18, 116) = "`Schwindel Aufstehen`"
 Str(0, 18, 117) = "`Folgeerkrankungen Haut`"
 Str(0, 18, 118) = "`Bewegungseinschränkungen`"
 Str(0, 18, 119) = "`Sexualstörung`"
 Str(0, 18, 120) = "`Sexualstörung seit`"
 Str(0, 18, 121) = "`Weitere Anamnese`"
 Str(0, 18, 122) = "`Alkohol`"
 Str(0, 18, 123) = "`Tabak`"
 Str(0, 18, 124) = "`tabakex`"
 Str(0, 18, 125) = "`tabakbis`"
 Str(0, 18, 126) = "`tabakakt`"
 Str(0, 18, 127) = "`tabakmenge`"
 Str(0, 18, 128) = "`Weitere Medikation`"
 Str(0, 18, 129) = "`Liphypertrophien Abdomen`"
 Str(0, 18, 130) = "`Liphypertrophien Beine`"
 Str(0, 18, 131) = "`Liphypertrophien Arme`"
 Str(0, 18, 132) = "`Beinbefund`"
 Str(0, 18, 133) = "`Hyperkeratosen`"
 Str(0, 18, 134) = "`Ulcera`"
 Str(0, 18, 135) = "`Kraft Zehenheber`"
 Str(0, 18, 136) = "`Kraft Zehenbeuger`"
 Str(0, 18, 137) = "`Kraft Knie`"
 Str(0, 18, 138) = "`ASR`"
 Str(0, 18, 139) = "`PSR`"
 Str(0, 18, 140) = "`Oberflächensensibilität`"
 Str(0, 18, 141) = "`Monofilamenttest`"
 Str(0, 18, 142) = "`Kalt-Warm`"
 Str(0, 18, 143) = "`Vibration IK`"
 Str(0, 18, 144) = "`Vibration Großzehe`"
 Str(0, 18, 145) = "`Puls Leiste`"
 Str(0, 18, 146) = "`Puls Kniekehle`"
 Str(0, 18, 147) = "`Puls Atp`"
 Str(0, 18, 148) = "`Puls Adp`"
 Str(0, 18, 149) = "`RR`"
 Str(0, 18, 150) = "`RRTurboMed`"
 Str(0, 18, 151) = "`Herz`"
 Str(0, 18, 152) = "`Lunge`"
 Str(0, 18, 153) = "`Bauch`"
 Str(0, 18, 154) = "`WS`"
 Str(0, 18, 155) = "`NL`"
 Str(0, 18, 156) = "`SD`"
 Str(0, 18, 157) = "`Carotiden`"
 Str(0, 18, 158) = "`NNH`"
 Str(0, 18, 159) = "`Zähne`"
 Str(0, 18, 160) = "`Mundhöhle`"
 Str(0, 18, 161) = "`LK`"
 Str(0, 18, 162) = "`BeinödVen`"
 Str(0, 18, 163) = "`Neuro sonst`"
 Str(0, 18, 164) = "`Weitere Befunde`"
 Str(0, 18, 165) = "`Schulung`"
 Str(0, 18, 166) = "`DMP`"
 Str(0, 18, 167) = "`DMSchulz`"
 Str(0, 18, 168) = "`DMSchL`"
 Str(0, 18, 169) = "`RRSchulz`"
 Str(0, 18, 170) = "`DMPhier`"
 Str(0, 18, 171) = "`HANr`"
 Str(0, 18, 172) = "`HANr2`"
 Str(0, 18, 173) = "`letzte Änderung`"
 Str(0, 18, 174) = "`Diagnosen`"
 Str(0, 18, 175) = "`Vorgestellt`"
 Str(0, 18, 176) = "`Versicherung`"
 Str(0, 18, 177) = "`AktZeit`"
 Str(0, 18, 178) = "`Ther1`"
 Str(0, 18, 179) = "`TherAkt`"
 Str(0, 18, 180) = "`obAn1eing`"
 Str(0, 18, 181) = "`obAn2eing`"
 Str(0, 18, 182) = "`obAnAeing`"
 Str(0, 18, 183) = "`obCheck`"
 Str(0, 18, 184) = "`obBZausgew`"
 Str(0, 18, 185) = "`obOSaufgek`"
 Str(0, 18, 186) = "`obPodAufgek`"
 Str(0, 18, 187) = "`obMBlAusgeh`"
 Str(0, 18, 188) = "`obSchulaufgek`"
 Str(0, 18, 189) = "`obDMPaufgekl`"
 Str(0, 18, 190) = "`obMedNetz`"
 Str(0, 18, 191) = "`Hausarzt`"
 Str(0, 18, 192) = "`ob`"
 Str(0, 18, 193) = "`QS`"
 Str(0, 18, 194) = "`QT`"
 Str(0, 18, 195) = "`Prim`"
 Str(0, 18, 196) = "`PrimaryKey`"
 Str(0, 18, 197) = "`Auswahl`"
 Str(0, 18, 198) = "`DMPhier`"
 Str(0, 18, 199) = "`GebDat`"
 Str(0, 18, 200) = "`jlD`"
 Str(0, 18, 201) = "`lÄnd`"
 Str(0, 18, 202) = "`Pat_ID`"
 Str(0, 18, 203) = "`Ther1`"
 Str(0, 18, 204) = "`Vorgestellt`"
 Str(0, 18, 205) = "`HausärzteAnamnesebogen_AccRel`"
 Str(0, 18, 206) = "`KassenlisteAnamnesebogen_AccRel`"
 ArtZ(0, 18) = 194
 ArtZ(1, 18) = 12
 Str(1, 18, 0) = "CREATE TABLE `anbogalt` ("
 Str(1, 18, 1) = " `Prim` int(2) unsigned NOT NULL COMMENT 'Primärschlüssel'"
 Str(1, 18, 2) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 18, 3) = " `Nachname` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '-'"
 Str(1, 18, 4) = " `Vorname` varchar(19) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 5) = " `NVorsatz` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 6) = " `Titel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 7) = " `Anrede` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 8) = " `GebDat` datetime DEFAULT NULL COMMENT ', geb.'"
 Str(1, 18, 9) = " `Tkz` tinyint(1) unsigned DEFAULT NULL COMMENT 'Tod-Kennzeichen'"
 Str(1, 18, 10) = " `Versicherungsart` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 11) = " `Diabetestyp` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetes Typ'"
 Str(1, 18, 12) = " `Diabetes seit` varchar(152) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 18, 13) = " `Tabletten seit` varchar(66) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Tabletten seit'"
 Str(1, 18, 14) = " `Insulin seit` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Insulin seit'"
 Str(1, 18, 15) = " `Grund für Vorstellung` varchar(721) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 16) = " `Familienanamnese` varchar(291) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 17) = " `Größe` double DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 18) = " `Gewicht` double DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 19) = " `Tendenz` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Tendenz'"
 Str(1, 18, 20) = " `DiabetesMedikament 1` varchar(48) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesmedikation:'"
 Str(1, 18, 21) = " `DiabetesMedikament 1 Menge` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 18, 22) = " `DiabetesMedikament 2` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 18, 23) = " `DiabetesMedikament 2 Menge` varchar(126) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 18, 24) = " `DiabetesMedikament 3` varchar(37) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 18, 25) = " `DiabetesMedikament 3 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 18, 26) = " `DiabetesMedikament 4` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 18, 27) = " `DiabetesMedikament 4 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 18, 28) = " `Insulinpumpe` tinyint(1) unsigned DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 29) = " `Insulinpumpe seit` varchar(250) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 18, 30) = " `Insulinpumpe Marke` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Marke:'"
 Str(1, 18, 31) = " `Broteinheiten gesamt` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Broteinheiten:gesamt'"
 Str(1, 18, 32) = " `Broteinheiten früh` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, früh'"
 Str(1, 18, 33) = " `Broteinheiten ZM früh` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zwischenmahlzeit vormittags'"
 Str(1, 18, 34) = " `Broteinheiten mittags` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 18, 35) = " `Broteinheiten nachmittags` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 18, 36) = " `Broteinheiten abends` varchar(13) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 18, 37) = " `Broteinheiten nachts` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 18, 38) = " `Essenszeit früh` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Essenszeiten:früh'"
 Str(1, 18, 39) = " `Essenszeit vormittags` varchar(34) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vormittags'"
 Str(1, 18, 40) = " `Essenszeit mittags` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 18, 41) = " `Essenszeit nachmittags` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 18, 42) = " `Essenszeit abends` varchar(18) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 18, 43) = " `Essenszeit spät` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, spät'"
 Str(1, 18, 44) = " `Spritz-Eß-Abstand früh` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritz-Eß-Abstand:früh'"
 Str(1, 18, 45) = " `Spritz-Eß-Abstand mittags` varchar(27) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 18, 46) = " `Spritz-Eß-Abstand abends` varchar(62) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 18, 47) = " `Spritzstelle früh` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritzstellen:früh'"
 Str(1, 18, 48) = " `Spritzstelle mittags` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 18, 49) = " `Spritzstelle abends` varchar(139) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 18, 50) = " `Spritzstelle nachts` varchar(141) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 18, 51) = " `Jahr letzte Diabetesschulung` varchar(129) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesschulung:'"
 Str(1, 18, 52) = " `Ort Schulung` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<in'"
 Str(1, 18, 53) = " `letztes HbA1c` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letztes HbA1c:'"
 Str(1, 18, 54) = " `gemessen am` datetime DEFAULT NULL COMMENT '<, gemessen'"
 Str(1, 18, 55) = " `vorherige Werte` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vorher:'"
 Str(1, 18, 56) = " `BZMessungen selbst` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckermessung:Selbstmessung?'"
 Str(1, 18, 57) = " `Gerät` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<:'"
 Str(1, 18, 58) = " `BZMessungen pW` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl d.Messungen pro Woche:'"
 Str(1, 18, 59) = " `BZMessungen pW ndE` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, davon nach dem Essen:'"
 Str(1, 18, 60) = " `BZMessungen p W nachts` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts:'"
 Str(1, 18, 61) = " `Aufschreiben` varchar(93) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Dokumentation:'"
 Str(1, 18, 62) = " `BZWerte v d Essen` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckerwerte vor dem Essen:'"
 Str(1, 18, 63) = " `BZWerte n d Essen` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nach dem Essen:'"
 Str(1, 18, 64) = " `UZ Tageszeit` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Unterzucker:Bevorzugte Tages-/Uhrzeit'"
 Str(1, 18, 65) = " `Unterzucker pM` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl der schweren (<50 mg/dl) pro Monat:'"
 Str(1, 18, 66) = " `UZ rechtzeitig` varchar(101) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, rechtzeitig bemerkt:'"
 Str(1, 18, 67) = " `Fremde Hilfe pa` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, fremde Hilfe deshalb nötig:'"
 Str(1, 18, 68) = " `Bewußtlos pa` varchar(81) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, bewußtlos deshalb:'"
 Str(1, 18, 69) = " `Keto pa` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Zahl der Ketoazidosen pro Jahr:'"
 Str(1, 18, 70) = " `BZgr300 pM` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Zahl der Blutzucker > 300 mg/dl pro Monat:'"
 Str(1, 18, 71) = " `Bluthochdruck` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Bluthochdruck:'"
 Str(1, 18, 72) = " `BHD seit` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit:'"
 Str(1, 18, 73) = " `BHD beh mit` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, behandelt mit:'"
 Str(1, 18, 74) = " `Blutdruckwerte` varchar(186) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckwerte:'"
 Str(1, 18, 75) = " `BDselbst` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckselbstmessung:'"
 Str(1, 18, 76) = " `Schwanger` varchar(43) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Aktuelle Schwangerschaft:'"
 Str(1, 18, 77) = " `Schwanger seit` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, seit:'"
 Str(1, 18, 78) = " `Augensp zuletzt` varchar(107) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Augenspiegelung:'"
 Str(1, 18, 79) = " `Augensp Befund` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 18, 80) = " `Netzhaut gelasert` varchar(142) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Netzhaut schon gelasert:'"
 Str(1, 18, 81) = " `Sehminderung unbehebbar` varchar(151) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', mit Brille nicht behebbare Sehminderung:'"
 Str(1, 18, 82) = " `Diabet Nierenschaden` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetischer Nierenschaden:'"
 Str(1, 18, 83) = " `Albumin zuletzt` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', letztes Albumin:'"
 Str(1, 18, 84) = " `erhöht?` varchar(52) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 18, 85) = " `Dialyse` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 86) = " `Dialyse seit` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 18, 87) = " `andere Nierenerkrankung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', andere Nierenerkrankung:'"
 Str(1, 18, 88) = " `Herzkrankheit` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Herzkrankheit:'"
 Str(1, 18, 89) = " `Angina pectoris` varchar(134) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 90) = " `Herzinfarkt` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 91) = " `Herzinfarkt wann` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 18, 92) = " `PTCA oder Stent` varchar(94) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 93) = " `Bypass kardial` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 94) = " `Bypass wann` varchar(388) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 18, 95) = " `Herzschwäche` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 96) = " `Herzkrankheit Beschreibung` varchar(213) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung:'"
 Str(1, 18, 97) = " `Hirndurchblutungsstörung` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 98) = " `Schlaganfall` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 99) = " `Beindurchblutungsstörung` varchar(136) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 100) = " `Schaufensterkrankheit` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 101) = " `Bypaß peripher` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 102) = " `Geschwür` varchar(174) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 103) = " `Amputation` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 104) = " `pAVK Beschreibung` varchar(97) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung der Beinbeschwerden:'"
 Str(1, 18, 105) = " `Ameisenlaufen` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 106) = " `Ameisen Ausmaß` varchar(123) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Ausmaß:'"
 Str(1, 18, 107) = " `Druckstellen` varchar(173) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 108) = " `Verformungen` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 109) = " `Verformungen Beschreibung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Beschreibung:'"
 Str(1, 18, 110) = " `Fußpflege` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 111) = " `Podologie` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 112) = " `Einlagen` varchar(91) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', diabetesgerechte orthopädische Einlagen/Schuhe:'"
 Str(1, 18, 113) = " `Neue Fußkomplikationen` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Neue Fußkomplikationen in den letzten 12 Monaten:'"
 Str(1, 18, 114) = " `Entleerungsstörungen Magen` varchar(130) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 115) = " `Entleerungsstörungen Harnblase` varchar(157) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 116) = " `Schwindel Aufstehen` varchar(121) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 117) = " `Folgeerkrankungen Haut` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 118) = " `Bewegungseinschränkungen` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 119) = " `Sexualstörung` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 120) = " `Sexualstörung seit` varchar(149) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 18, 121) = " `Weitere Anamnese` varchar(989) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 122) = " `Alkohol` varchar(148) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 123) = " `Tabak` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 124) = " `tabakex` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 125) = " `tabakbis` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 126) = " `tabakakt` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 127) = " `tabakmenge` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 128) = " `Weitere Medikation` varchar(298) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 129) = " `Liphypertrophien Abdomen` varchar(176) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Liphypertrophien:Abdomen'"
 Str(1, 18, 130) = " `Liphypertrophien Beine` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Beine:'"
 Str(1, 18, 131) = " `Liphypertrophien Arme` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Arme:'"
 Str(1, 18, 132) = " `Beinbefund` varchar(272) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 133) = " `Hyperkeratosen` varchar(207) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 134) = " `Ulcera` varchar(103) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 135) = " `Kraft Zehenheber` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Kraft:Zehenheber'"
 Str(1, 18, 136) = " `Kraft Zehenbeuger` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zehenbeuger:'"
 Str(1, 18, 137) = " `Kraft Knie` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Knie:'"
 Str(1, 18, 138) = " `ASR` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 139) = " `PSR` varchar(67) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 140) = " `Oberflächensensibilität` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 141) = " `Monofilamenttest` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 142) = " `Kalt-Warm` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Kalt-Warm-Diskrimination:'"
 Str(1, 18, 143) = " `Vibration IK` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Vibrationsempfinden Innenknöchel:'"
 Str(1, 18, 144) = " `Vibration Großzehe` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Großzehe:'"
 Str(1, 18, 145) = " `Puls Leiste` varchar(33) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Pulse:Leiste'"
 Str(1, 18, 146) = " `Puls Kniekehle` varchar(26) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Kniekehle:'"
 Str(1, 18, 147) = " `Puls Atp` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Innenknöchel:'"
 Str(1, 18, 148) = " `Puls Adp` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Fußrücken:'"
 Str(1, 18, 149) = " `RR` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruck:'"
 Str(1, 18, 150) = " `RRTurboMed` varchar(1362) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 151) = " `Herz` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 18, 152) = " `Lunge` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 153) = " `Bauch` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Abdomen:'"
 Str(1, 18, 154) = " `WS` varchar(56) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Wirbelsäule:'"
 Str(1, 18, 155) = " `NL` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nierenlager:'"
 Str(1, 18, 156) = " `SD` varchar(88) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Schilddrüse:'"
 Str(1, 18, 157) = " `Carotiden` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Halsschlagadern:'"
 Str(1, 18, 158) = " `NNH` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nasennebenhöhlen:'"
 Str(1, 18, 159) = " `Zähne` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 160) = " `Mundhöhle` varchar(64) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 18, 161) = " `LK` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Lymphknoten:'"
 Str(1, 18, 162) = " `BeinödVen` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beinödeme/ Venenkrankheiten:'"
 Str(1, 18, 163) = " `Neuro sonst` varchar(74) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Sonstige neurologische Befunde:'"
 Str(1, 18, 164) = " `Weitere Befunde` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', weitere Befunde:'"
 Str(1, 18, 165) = " `Schulung` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Schulungsbedarf'"
 Str(1, 18, 166) = " `DMP` varchar(85) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Pat. bei HA im DMP'"
 Str(1, 18, 167) = " `DMSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der DMP-Schulungen hier'"
 Str(1, 18, 168) = " `DMSchL` smallint(6) DEFAULT NULL COMMENT 'Zahl der abgerechneten DMP-Schulungen hier'"
 Str(1, 18, 169) = " `RRSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der Hypertonie-Schulungen hier'"
 Str(1, 18, 170) = " `DMPhier` datetime DEFAULT NULL COMMENT 'ob Pat hier im DMP'"
 Str(1, 18, 171) = " `HANr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 18, 172) = " `HANr2` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 18, 173) = " `letzte Änderung` datetime DEFAULT NULL COMMENT 'Datum der letzten Änderung'"
 Str(1, 18, 174) = " `Diagnosen` varchar(1071) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 175) = " `Vorgestellt` datetime DEFAULT NULL COMMENT 'Erstvorstellung'"
 Str(1, 18, 176) = " `Versicherung` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 177) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 18, 178) = " `Ther1` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, ICT, CSII'"
 Str(1, 18, 179) = " `TherAkt` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, ICT, CSII'"
 Str(1, 18, 180) = " `obAn1eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 1 eingegeben wurde'"
 Str(1, 18, 181) = " `obAn2eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 2 eingegeben wurde'"
 Str(1, 18, 182) = " `obAnAeing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt allgemein eingegeben wurde'"
 Str(1, 18, 183) = " `obCheck` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Checkliste vorliegt'"
 Str(1, 18, 184) = " `obBZausgew` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Blutzuckergerät ausgewechselt'"
 Str(1, 18, 185) = " `obOSaufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über orthopäd Schuhmacher aufgeklärt'"
 Str(1, 18, 186) = " `obPodAufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 18, 187) = " `obMBlAusgeh` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 18, 188) = " `obSchulaufgek` varchar(14) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 18, 189) = " `obDMPaufgekl` varchar(17) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 18, 190) = " `obMedNetz` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob von Med. Netz geschickt'"
 Str(1, 18, 191) = " `Hausarzt` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Hausarzt laut Anamnesebogen'"
 Str(1, 18, 192) = " `ob` tinyint(1) unsigned DEFAULT NULL COMMENT 'für verschiedene Aktionen'"
 Str(1, 18, 193) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 18, 194) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 18, 195) = "  PRIMARY KEY (`Prim`)"
 Str(1, 18, 196) = "  UNIQUE KEY `PrimaryKey` (`Prim`)"
 Str(1, 18, 197) = "  KEY `Auswahl` (`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 18, 198) = "  KEY `DMPhier` (`DMPhier`)"
 Str(1, 18, 199) = "  KEY `GebDat` (`GebDat`)"
 Str(1, 18, 200) = "  KEY `jlD` (`Jahr letzte Diabetesschulung`,`GebDat`)"
 Str(1, 18, 201) = "  KEY `lÄnd` (`letzte Änderung`)"
 Str(1, 18, 202) = "  KEY `Pat_ID` (`Pat_id`)"
 Str(1, 18, 203) = "  KEY `Ther1` (`Ther1`)"
 Str(1, 18, 204) = "  KEY `Vorgestellt` (`Vorgestellt`,`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 18, 205) = "  KEY `HausärzteAnamnesebogen_AccRel` (`HANr`)"
 Str(1, 18, 206) = "  KEY `KassenlisteAnamnesebogen_AccRel` (`Versicherung`)"
 Str(1, 18, 207) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr18

Function FüllStr19()
 Str(0, 19, 0) = "au"
 Str(0, 19, 1) = "`FID`"
 Str(0, 19, 2) = "`Pat_ID`"
 Str(0, 19, 3) = "`ZeitPunkt`"
 Str(0, 19, 4) = "`Beginn`"
 Str(0, 19, 5) = "`Ende`"
 Str(0, 19, 6) = "`ICDs`"
 Str(0, 19, 7) = "`absPos`"
 Str(0, 19, 8) = "`AktZeit`"
 Str(0, 19, 9) = "`StByte`"
 Str(0, 19, 10) = "`Auswahl`"
 Str(0, 19, 11) = "`FälleAU`"
 Str(0, 19, 12) = "`FID`"
 Str(0, 19, 13) = "`NamenAU`"
 Str(0, 19, 14) = "`FälleAU_AccRel`"
 Str(0, 19, 15) = "`NamenAU_AccRel`"
 ArtZ(0, 19) = 9
 ArtZ(1, 19) = 4
 ArtZ(2, 19) = 2
 Str(1, 19, 0) = "CREATE TABLE `au` ("
 Str(1, 19, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 19, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 19, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 19, 4) = " `Beginn` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6285 1. Hälfte'"
 Str(1, 19, 5) = " `Ende` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6285 2. Hälfte'"
 Str(1, 19, 6) = " `ICDs` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6286'"
 Str(1, 19, 7) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 19, 8) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 19, 9) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 19, 10) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Beginn`,`Ende`)"
 Str(1, 19, 11) = "  KEY `FälleAU` (`FID`)"
 Str(1, 19, 12) = "  KEY `FID` (`FID`)"
 Str(1, 19, 13) = "  KEY `NamenAU` (`Pat_ID`)"
 Str(1, 19, 14) = "  CONSTRAINT `FälleAU_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON UPDATE CASCADE"
 Str(1, 19, 15) = "  CONSTRAINT `NamenAU_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 19, 16) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr19

Function FüllStr20()
 Str(0, 20, 0) = "augenbefunde"
 Str(0, 20, 1) = "`ID`"
 Str(0, 20, 2) = "`Pat_ID`"
 Str(0, 20, 3) = "`Datum`"
 Str(0, 20, 4) = "`Retinopathie`"
 Str(0, 20, 5) = "`Klassifikation`"
 Str(0, 20, 6) = "`Maculopathie`"
 Str(0, 20, 7) = "`Sonstiges`"
 Str(0, 20, 8) = "`Visus re`"
 Str(0, 20, 9) = "`Visus li`"
 Str(0, 20, 10) = "`KontrolleinMonaten`"
 Str(0, 20, 11) = "`Augenarzt`"
 Str(0, 20, 12) = "`DokName`"
 Str(0, 20, 13) = "`DokPfad`"
 Str(0, 20, 14) = "`verglichen`"
 Str(0, 20, 15) = "`ID`"
 Str(0, 20, 16) = "`PrimaryKey`"
 Str(0, 20, 17) = "`Auswahl`"
 Str(0, 20, 18) = "`NamenAugenbefunde`"
 Str(0, 20, 19) = "`Text`"
 Str(0, 20, 20) = "`NamenAugenbefunde_AccRel`"
 ArtZ(0, 20) = 14
 ArtZ(1, 20) = 5
 ArtZ(2, 20) = 1
 Str(1, 20, 0) = "CREATE TABLE `augenbefunde` ("
 Str(1, 20, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 20, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 20, 3) = " `Datum` datetime DEFAULT NULL"
 Str(1, 20, 4) = " `Retinopathie` longtext COLLATE latin1_german2_ci"
 Str(1, 20, 5) = " `Klassifikation` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 6) = " `Maculopathie` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 7) = " `Sonstiges` longtext COLLATE latin1_german2_ci"
 Str(1, 20, 8) = " `Visus re` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 9) = " `Visus li` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 10) = " `KontrolleinMonaten` float DEFAULT NULL"
 Str(1, 20, 11) = " `Augenarzt` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 12) = " `DokName` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 13) = " `DokPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 20, 14) = " `verglichen` datetime DEFAULT NULL"
 Str(1, 20, 15) = "  PRIMARY KEY (`ID`)"
 Str(1, 20, 16) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 20, 17) = "  KEY `Auswahl` (`Pat_ID`,`DokPfad`(255))"
 Str(1, 20, 18) = "  KEY `NamenAugenbefunde` (`Pat_ID`)"
 Str(1, 20, 19) = "  KEY `Text` (`Pat_ID`,`Retinopathie`(255))"
 Str(1, 20, 20) = "  CONSTRAINT `NamenAugenbefunde_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 20, 21) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr20

Function FüllStr21()
 Str(0, 21, 0) = "briefe"
 Str(0, 21, 1) = "`FID`"
 Str(0, 21, 2) = "`Pat_ID`"
 Str(0, 21, 3) = "`ZeitPunkt`"
 Str(0, 21, 4) = "`Pfad`"
 Str(0, 21, 5) = "`Art`"
 Str(0, 21, 6) = "`Name`"
 Str(0, 21, 7) = "`Typ`"
 Str(0, 21, 8) = "`AktZeit`"
 Str(0, 21, 9) = "`DokGroe`"
 Str(0, 21, 10) = "`QS`"
 Str(0, 21, 11) = "`QT`"
 Str(0, 21, 12) = "`absPos`"
 Str(0, 21, 13) = "`StByte`"
 Str(0, 21, 14) = "`Auswahl`"
 Str(0, 21, 15) = "`FälleBriefe`"
 Str(0, 21, 16) = "`FID`"
 Str(0, 21, 17) = "`NamenBriefe`"
 Str(0, 21, 18) = "`FälleBriefe_AccRel`"
 Str(0, 21, 19) = "`NamenBriefe_AccRel`"
 ArtZ(0, 21) = 13
 ArtZ(1, 21) = 4
 ArtZ(2, 21) = 2
 Str(1, 21, 0) = "CREATE TABLE `briefe` ("
 Str(1, 21, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 21, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 21, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 21, 4) = " `Pfad` varchar(128) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 5) = " `Art` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 6) = " `Name` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 7) = " `Typ` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 21, 8) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 21, 9) = " `DokGroe` int(10) DEFAULT NULL COMMENT 'Größe der Datei'"
 Str(1, 21, 10) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 21, 11) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 21, 12) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 21, 13) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 21, 14) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Name`)"
 Str(1, 21, 15) = "  KEY `FälleBriefe` (`FID`)"
 Str(1, 21, 16) = "  KEY `FID` (`FID`)"
 Str(1, 21, 17) = "  KEY `NamenBriefe` (`Pat_ID`)"
 Str(1, 21, 18) = "  CONSTRAINT `FälleBriefe_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 21, 19) = "  CONSTRAINT `NamenBriefe_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 21, 20) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr21

Function FüllStr22()
 Str(0, 22, 0) = "diagnosen"
 Str(0, 22, 1) = "`ID1`"
 Str(0, 22, 2) = "`FID`"
 Str(0, 22, 3) = "`Pat_id`"
 Str(0, 22, 4) = "`GesName`"
 Str(0, 22, 5) = "`DiagDatum`"
 Str(0, 22, 6) = "`DiagSicherheit`"
 Str(0, 22, 7) = "`DiagText`"
 Str(0, 22, 8) = "`DiagSeite`"
 Str(0, 22, 9) = "`DiagAttr`"
 Str(0, 22, 10) = "`ICD`"
 Str(0, 22, 11) = "`obDauer`"
 Str(0, 22, 12) = "`Ausnahme`"
 Str(0, 22, 13) = "`intBemerk`"
 Str(0, 22, 14) = "`absPos`"
 Str(0, 22, 15) = "`AktZeit`"
 Str(0, 22, 16) = "`StByte`"
 Str(0, 22, 17) = "`ID1`"
 Str(0, 22, 18) = "`PrimaryKey`"
 Str(0, 22, 19) = "`Auswahl`"
 Str(0, 22, 20) = "`DiagSuch`"
 Str(0, 22, 21) = "`DiagText`"
 Str(0, 22, 22) = "`FälleDiagnosen`"
 Str(0, 22, 23) = "`FID`"
 Str(0, 22, 24) = "`NamenDiagnosen`"
 Str(0, 22, 25) = "`ICD`"
 Str(0, 22, 26) = "`FälleDiagnosen_AccRel`"
 Str(0, 22, 27) = "`NamenDiagnosen_AccRel`"
 ArtZ(0, 22) = 16
 ArtZ(1, 22) = 9
 ArtZ(2, 22) = 2
 Str(1, 22, 0) = "CREATE TABLE `diagnosen` ("
 Str(1, 22, 1) = " `ID1` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 22, 2) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 22, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT 'Bezug auf Anamneseblattt'"
 Str(1, 22, 4) = " `GesName` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 22, 5) = " `DiagDatum` datetime DEFAULT NULL"
 Str(1, 22, 6) = " `DiagSicherheit` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6003'"
 Str(1, 22, 7) = " `DiagText` longtext COLLATE latin1_german2_ci"
 Str(1, 22, 8) = " `DiagSeite` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6004'"
 Str(1, 22, 9) = " `DiagAttr` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6006'"
 Str(1, 22, 10) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 22, 11) = " `obDauer` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Dauerdiagnose'"
 Str(1, 22, 12) = " `Ausnahme` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3677 Ausnahme / Begründung für abweichendes Geschlecht'"
 Str(1, 22, 13) = " `intBemerk` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6009 interne Bemerkung'"
 Str(1, 22, 14) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 22, 15) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 22, 16) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 22, 17) = "  PRIMARY KEY (`ID1`)"
 Str(1, 22, 18) = "  UNIQUE KEY `PrimaryKey` (`ID1`)"
 Str(1, 22, 19) = "  KEY `Auswahl` (`Pat_id`,`DiagDatum`,`DiagSicherheit`,`DiagSeite`,`DiagAttr`,`DiagText`(255),`ICD`,`obDauer`)"
 Str(1, 22, 20) = "  KEY `DiagSuch` (`Pat_id`,`ICD`,`DiagSicherheit`,`DiagSeite`)"
 Str(1, 22, 21) = "  KEY `DiagText` (`Pat_id`,`DiagText`(255))"
 Str(1, 22, 22) = "  KEY `FälleDiagnosen` (`FID`)"
 Str(1, 22, 23) = "  KEY `FID` (`FID`)"
 Str(1, 22, 24) = "  KEY `NamenDiagnosen` (`Pat_id`)"
 Str(1, 22, 25) = "  KEY `ICD` (`ICD`)"
 Str(1, 22, 26) = "  CONSTRAINT `FälleDiagnosen_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 22, 27) = "  CONSTRAINT `NamenDiagnosen_AccRel` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 22, 28) = " ENGINE=InnoDB AUTO_INCREMENT=34884 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr22

Function FüllStr23()
 Str(0, 23, 0) = "diagnosen exportiert"
 Str(0, 23, 1) = "`ID`"
 Str(0, 23, 2) = "`Datum`"
 Str(0, 23, 3) = "`Pat_id`"
 Str(0, 23, 4) = "`ICD`"
 Str(0, 23, 5) = "`Diagnose`"
 Str(0, 23, 6) = "`übertragen`"
 Str(0, 23, 7) = "`ID`"
 Str(0, 23, 8) = "`PrimaryKey`"
 Str(0, 23, 9) = "`ID`"
 ArtZ(0, 23) = 6
 ArtZ(1, 23) = 3
 Str(1, 23, 0) = "CREATE TABLE `diagnosen exportiert` ("
 Str(1, 23, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Reihenfolge'"
 Str(1, 23, 2) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum'"
 Str(1, 23, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!Pat_id'"
 Str(1, 23, 4) = " `ICD` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD-Nummer der Diagnose'"
 Str(1, 23, 5) = " `Diagnose` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Text der Diagnose'"
 Str(1, 23, 6) = " `übertragen` datetime DEFAULT NULL COMMENT '"""", übertragen'"
 Str(1, 23, 7) = "  PRIMARY KEY (`ID`)"
 Str(1, 23, 8) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 23, 9) = "  KEY `ID` (`Pat_id`)"
 Str(1, 23, 10) = " ENGINE=InnoDB AUTO_INCREMENT=5911 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr23

Function FüllStr24()
 Str(0, 24, 0) = "diagnosenexport"
 Str(0, 24, 1) = "`ID`"
 Str(0, 24, 2) = "`Name`"
 Str(0, 24, 3) = "`Pat_id`"
 Str(0, 24, 4) = "`ICD`"
 Str(0, 24, 5) = "`Diagnose`"
 Str(0, 24, 6) = "`Status`"
 Str(0, 24, 7) = "`Protokoll`"
 Str(0, 24, 8) = "`nurQuart`"
 Str(0, 24, 9) = "`Zeitpunkt`"
 Str(0, 24, 10) = "`ID`"
 Str(0, 24, 11) = "`ID`"
 Str(0, 24, 12) = "`pat_ID`"
 Str(0, 24, 13) = "`Suche`"
 ArtZ(0, 24) = 9
 ArtZ(1, 24) = 4
 Str(1, 24, 0) = "CREATE TABLE `diagnosenexport` ("
 Str(1, 24, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 24, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 24, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 24, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 24, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 24, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 24, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 24, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 24, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 24, 10) = "  PRIMARY KEY (`ID`)"
 Str(1, 24, 11) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 24, 12) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 24, 13) = "  KEY `Suche` (`Pat_id`,`ICD`)"
 Str(1, 24, 14) = " ENGINE=InnoDB AUTO_INCREMENT=2267 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr24

Function FüllStr25()
 Str(0, 25, 0) = "diagnoseng1"
 Str(0, 25, 1) = "`lfdnr`"
 Str(0, 25, 2) = "`gruppe`"
 Str(0, 25, 3) = "`rf`"
 Str(0, 25, 4) = "`lfdnr`"
 Str(0, 25, 5) = "`gruppe`"
 Str(0, 25, 6) = "`rf`"
 ArtZ(0, 25) = 3
 ArtZ(1, 25) = 3
 Str(1, 25, 0) = "CREATE TABLE `diagnoseng1` ("
 Str(1, 25, 1) = " `lfdnr` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 25, 2) = " `gruppe` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Gruppenname'"
 Str(1, 25, 3) = " `rf` int(10) unsigned NOT NULL DEFAULT '0' COMMENT 'Reihenfolge'"
 Str(1, 25, 4) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 25, 5) = "  KEY `gruppe` (`gruppe`) USING BTREE"
 Str(1, 25, 6) = "  KEY `rf` (`rf`)"
 Str(1, 25, 7) = " ENGINE=InnoDB AUTO_INCREMENT=23 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Diagnosengruppierung 1'"
End Function ' FüllStr25

Function FüllStr26()
 Str(0, 26, 0) = "diagreihe"
 Str(0, 26, 1) = "`lfdnr`"
 Str(0, 26, 2) = "`ICD`"
 Str(0, 26, 3) = "`dg1`"
 Str(0, 26, 4) = "`dg2`"
 Str(0, 26, 5) = "`rf`"
 Str(0, 26, 6) = "`lfdnr`"
 Str(0, 26, 7) = "`ICD`"
 Str(0, 26, 8) = "`rf`"
 Str(0, 26, 9) = "`dg1`"
 Str(0, 26, 10) = "`dg1`"
 Str(0, 26, 11) = "`dg1_rel`"
 Str(0, 26, 12) = "`ICD_rel`"
 ArtZ(0, 26) = 5
 ArtZ(1, 26) = 4
 ArtZ(2, 26) = 3
 Str(1, 26, 0) = "CREATE TABLE `diagreihe` ("
 Str(1, 26, 1) = " `lfdnr` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 26, 2) = " `ICD` varchar(10) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 26, 3) = " `dg1` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Diagnosenguppierung 1'"
 Str(1, 26, 4) = " `dg2` varchar(45) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Diagnosenguppierung 2'"
 Str(1, 26, 5) = " `rf` int(10) unsigned NOT NULL DEFAULT '0' COMMENT 'Reihenfolge für Arztbrief'"
 Str(1, 26, 6) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 26, 7) = "  UNIQUE KEY `ICD` (`ICD`) USING BTREE"
 Str(1, 26, 8) = "  KEY `rf` (`rf`)"
 Str(1, 26, 9) = "  KEY `dg1` (`dg1`)"
 Str(1, 26, 10) = "  CONSTRAINT `dg1` FOREIGN KEY (`dg1`) REFERENCES `diagnoseng1` (`gruppe`) ON UPDATE CASCADE"
 Str(1, 26, 11) = "  CONSTRAINT `dg1_rel` FOREIGN KEY (`dg1`) REFERENCES `diagnoseng1` (`gruppe`) ON UPDATE CASCADE"
 Str(1, 26, 12) = "  CONSTRAINT `ICD_rel` FOREIGN KEY (`ICD`) REFERENCES `diagnosen` (`ICD`)"
 Str(1, 26, 13) = " ENGINE=InnoDB AUTO_INCREMENT=1021 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Reihenfolge der Diagnosen für Arztbrief'"
End Function ' FüllStr26

Function FüllStr27()
 Str(0, 27, 0) = "dmp-uschr"
 Str(0, 27, 1) = "`Pat_id`"
 Str(0, 27, 2) = "`U1`"
 Str(0, 27, 3) = "`U2`"
 Str(0, 27, 4) = "`U3`"
 Str(0, 27, 5) = "`Arztwechsel`"
 Str(0, 27, 6) = "`Pat_id`"
 ArtZ(0, 27) = 5
 ArtZ(1, 27) = 1
 Str(1, 27, 0) = "CREATE TABLE `dmp-uschr` ("
 Str(1, 27, 1) = " `Pat_id` int(10) DEFAULT NULL COMMENT 'Bezug auf Namen'"
 Str(1, 27, 2) = " `U1` datetime DEFAULT NULL COMMENT 'vorliegendes Blatt mit DMP-Unterschrift'"
 Str(1, 27, 3) = " `U2` datetime DEFAULT NULL COMMENT 'vorliegendes 2. Blatt mit DMP-Unterschrift'"
 Str(1, 27, 4) = " `U3` datetime DEFAULT NULL COMMENT 'vorliegendes 3. Blatt mit DMP-Unterschrift'"
 Str(1, 27, 5) = " `Arztwechsel` bit(1) DEFAULT NULL COMMENT 'danach Arztwechsel'"
 Str(1, 27, 6) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 27, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr27

Function FüllStr28()
 Str(0, 28, 0) = "dmpreihe"
 Str(0, 28, 1) = "`Abk`"
 Str(0, 28, 2) = "`Art`"
 Str(0, 28, 3) = "`KarteiDatum`"
 Str(0, 28, 4) = "`exportiert`"
 Str(0, 28, 5) = "`DokuDatum`"
 Str(0, 28, 6) = "`obvoll`"
 Str(0, 28, 7) = "`NachName`"
 Str(0, 28, 8) = "`VorName`"
 Str(0, 28, 9) = "`GebDat`"
 Str(0, 28, 10) = "`Pat_id`"
 Str(0, 28, 11) = "`StByte`"
 Str(0, 28, 12) = "`AktZeit`"
 ArtZ(0, 28) = 12
 Str(1, 28, 0) = "CREATE TABLE `dmpreihe` ("
 Str(1, 28, 1) = " `Abk` varchar(30) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Abkürzung der DMP-Art'"
 Str(1, 28, 2) = " `Art` varchar(2) CHARACTER SET latin1 NOT NULL COMMENT 'ED = Erstdoku, FD = Folgedoku'"
 Str(1, 28, 3) = " `KarteiDatum` date DEFAULT NULL COMMENT 'Datum des Karteikarteneintrags der Dokumentation'"
 Str(1, 28, 4) = " `exportiert` datetime DEFAULT NULL COMMENT 'Datum des Exports'"
 Str(1, 28, 5) = " `DokuDatum` datetime DEFAULT NULL COMMENT 'Datum der Dokumentation'"
 Str(1, 28, 6) = " `obvoll` bit(1) DEFAULT NULL COMMENT 'ob vollständig'"
 Str(1, 28, 7) = " `NachName` varchar(20) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 28, 8) = " `VorName` varchar(20) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 28, 9) = " `GebDat` date DEFAULT NULL"
 Str(1, 28, 10) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 28, 11) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 28, 12) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungzeit'"
 Str(1, 28, 13) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr28

Function FüllStr29()
 Str(0, 29, 0) = "doc1"
 Str(0, 29, 1) = "`pat_id`"
 Str(0, 29, 2) = "`nachname`"
 Str(0, 29, 3) = "`vorname`"
 Str(0, 29, 4) = "`gebdat`"
 Str(0, 29, 5) = "`pat_id`"
 Str(0, 29, 6) = "`such`"
 ArtZ(0, 29) = 4
 ArtZ(1, 29) = 2
 Str(1, 29, 0) = "CREATE TABLE `doc1` ("
 Str(1, 29, 1) = " `pat_id` int(11) NOT NULL"
 Str(1, 29, 2) = " `nachname` tinytext COLLATE latin1_german2_ci"
 Str(1, 29, 3) = " `vorname` tinytext COLLATE latin1_german2_ci"
 Str(1, 29, 4) = " `gebdat` datetime DEFAULT NULL"
 Str(1, 29, 5) = "  PRIMARY KEY (`pat_id`)"
 Str(1, 29, 6) = "  KEY `such` (`nachname`(30),`vorname`(30))"
 Str(1, 29, 7) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr29

Function FüllStr30()
 Str(0, 30, 0) = "doc2"
 Str(0, 30, 1) = "`pat_id`"
 Str(0, 30, 2) = "`nachname`"
 Str(0, 30, 3) = "`vorname`"
 Str(0, 30, 4) = "`gebdat`"
 Str(0, 30, 5) = "`pat_id`"
 Str(0, 30, 6) = "`such`"
 ArtZ(0, 30) = 4
 ArtZ(1, 30) = 2
 Str(1, 30, 0) = "CREATE TABLE `doc2` ("
 Str(1, 30, 1) = " `pat_id` int(11) NOT NULL"
 Str(1, 30, 2) = " `nachname` tinytext COLLATE latin1_german2_ci"
 Str(1, 30, 3) = " `vorname` tinytext COLLATE latin1_german2_ci"
 Str(1, 30, 4) = " `gebdat` datetime DEFAULT NULL"
 Str(1, 30, 5) = "  PRIMARY KEY (`pat_id`)"
 Str(1, 30, 6) = "  KEY `such` (`nachname`(30),`vorname`(30))"
 Str(1, 30, 7) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr30

Function FüllStr31()
 Str(0, 31, 0) = "doce"
 Str(0, 31, 1) = "`pat_id`"
 Str(0, 31, 2) = "`nachname`"
 Str(0, 31, 3) = "`vorname`"
 Str(0, 31, 4) = "`gebdat`"
 Str(0, 31, 5) = "`pat_id`"
 Str(0, 31, 6) = "`such`"
 ArtZ(0, 31) = 4
 ArtZ(1, 31) = 2
 Str(1, 31, 0) = "CREATE TABLE `doce` ("
 Str(1, 31, 1) = " `pat_id` int(11) NOT NULL"
 Str(1, 31, 2) = " `nachname` tinytext COLLATE latin1_german2_ci"
 Str(1, 31, 3) = " `vorname` tinytext COLLATE latin1_german2_ci"
 Str(1, 31, 4) = " `gebdat` datetime DEFAULT NULL"
 Str(1, 31, 5) = "  PRIMARY KEY (`pat_id`)"
 Str(1, 31, 6) = "  KEY `such` (`nachname`(30),`vorname`(30))"
 Str(1, 31, 7) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr31

Function FüllStr32()
 Str(0, 32, 0) = "dokabkop"
 Str(0, 32, 1) = "`DokPfad`"
 Str(0, 32, 2) = "`AktZeit`"
 Str(0, 32, 3) = "`abgehakt`"
 Str(0, 32, 4) = "`DokPfad`"
 ArtZ(0, 32) = 3
 ArtZ(1, 32) = 1
 Str(1, 32, 0) = "CREATE TABLE `dokabkop` ("
 Str(1, 32, 1) = " `DokPfad` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 2) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 32, 3) = " `abgehakt` bit(1) DEFAULT NULL"
 Str(1, 32, 4) = "  KEY `DokPfad` (`DokPfad`)"
 Str(1, 32, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr32

Function FüllStr33()
 Str(0, 33, 0) = "dokumente"
 Str(0, 33, 1) = "`FID`"
 Str(0, 33, 2) = "`Pat_ID`"
 Str(0, 33, 3) = "`ZeitPunkt`"
 Str(0, 33, 4) = "`DokPfad`"
 Str(0, 33, 5) = "`DokArt`"
 Str(0, 33, 6) = "`DokName`"
 Str(0, 33, 7) = "`Quelldatum`"
 Str(0, 33, 8) = "`absPos`"
 Str(0, 33, 9) = "`AktZeit`"
 Str(0, 33, 10) = "`DokGroe`"
 Str(0, 33, 11) = "`QS`"
 Str(0, 33, 12) = "`QT`"
 Str(0, 33, 13) = "`StByte`"
 Str(0, 33, 14) = "`Auswahl`"
 Str(0, 33, 15) = "`DokName`"
 Str(0, 33, 16) = "`DokPfad`"
 Str(0, 33, 17) = "`FälleDokumente`"
 Str(0, 33, 18) = "`FID`"
 Str(0, 33, 19) = "`NamenDokumente`"
 Str(0, 33, 20) = "`PIDokPfad`"
 Str(0, 33, 21) = "`Quelldatum`"
 Str(0, 33, 22) = "`ZeitPunkt`"
 Str(0, 33, 23) = "`FälleDokumente_AccRel`"
 Str(0, 33, 24) = "`NamenDokumente_AccRel`"
 ArtZ(0, 33) = 13
 ArtZ(1, 33) = 9
 ArtZ(2, 33) = 2
 Str(1, 33, 0) = "CREATE TABLE `dokumente` ("
 Str(1, 33, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 33, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 33, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 33, 4) = " `DokPfad` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 5) = " `DokArt` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 6) = " `DokName` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 33, 7) = " `Quelldatum` datetime DEFAULT NULL COMMENT 'Datum, auf das sich das Dokument bezieht'"
 Str(1, 33, 8) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 33, 9) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 33, 10) = " `DokGroe` int(10) DEFAULT NULL COMMENT 'Dokument-Größe'"
 Str(1, 33, 11) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 33, 12) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 33, 13) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 33, 14) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`DokArt`,`DokName`)"
 Str(1, 33, 15) = "  KEY `DokName` (`DokName`)"
 Str(1, 33, 16) = "  KEY `DokPfad` (`DokPfad`)"
 Str(1, 33, 17) = "  KEY `FälleDokumente` (`FID`)"
 Str(1, 33, 18) = "  KEY `FID` (`FID`)"
 Str(1, 33, 19) = "  KEY `NamenDokumente` (`Pat_ID`)"
 Str(1, 33, 20) = "  KEY `PIDokPfad` (`Pat_ID`,`DokPfad`)"
 Str(1, 33, 21) = "  KEY `Quelldatum` (`Quelldatum`)"
 Str(1, 33, 22) = "  KEY `ZeitPunkt` (`ZeitPunkt`)"
 Str(1, 33, 23) = "  CONSTRAINT `FälleDokumente_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 33, 24) = "  CONSTRAINT `NamenDokumente_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 33, 25) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr33

Function FüllStr34()
 Str(0, 34, 0) = "dokumente abgehakt"
 Str(0, 34, 1) = "`DokPfad`"
 Str(0, 34, 2) = "`AktZeit`"
 Str(0, 34, 3) = "`abgehakt`"
 Str(0, 34, 4) = "`ungueltig`"
 Str(0, 34, 5) = "`DokPfad`"
 ArtZ(0, 34) = 4
 ArtZ(1, 34) = 1
 Str(1, 34, 0) = "CREATE TABLE `dokumente abgehakt` ("
 Str(1, 34, 1) = " `DokPfad` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 34, 2) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 34, 3) = " `abgehakt` bit(1) DEFAULT NULL"
 Str(1, 34, 4) = " `ungueltig` bit(1) DEFAULT NULL"
 Str(1, 34, 5) = "  KEY `DokPfad` (`DokPfad`)"
 Str(1, 34, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr34

Function FüllStr35()
 Str(0, 35, 0) = "ebm2000plus"
 Str(0, 35, 1) = "`Leistung`"
 Str(0, 35, 2) = "`Titel`"
 Str(0, 35, 3) = "`Punktwert`"
 Str(0, 35, 4) = "`Euro`"
 Str(0, 35, 5) = "`Bericht`"
 Str(0, 35, 6) = "`Text`"
 Str(0, 35, 7) = "`Betr`"
 Str(0, 35, 8) = "`Schul`"
 Str(0, 35, 9) = "`Typ1`"
 Str(0, 35, 10) = "`Typ2`"
 Str(0, 35, 11) = "`Gest`"
 Str(0, 35, 12) = "`DFS`"
 Str(0, 35, 13) = "`DMP`"
 Str(0, 35, 14) = "`AOK`"
 Str(0, 35, 15) = "`BKK`"
 Str(0, 35, 16) = "`BKN`"
 Str(0, 35, 17) = "`EK`"
 Str(0, 35, 18) = "`IKK`"
 Str(0, 35, 19) = "`LKK`"
 Str(0, 35, 20) = "`Üw`"
 Str(0, 35, 21) = "`Insulin`"
 Str(0, 35, 22) = "`ICT`"
 Str(0, 35, 23) = "`CSII`"
 Str(0, 35, 24) = "`Erst`"
 Str(0, 35, 25) = "`Folge`"
 Str(0, 35, 26) = "`fid`"
 Str(0, 35, 27) = "`Leistung`"
 Str(0, 35, 28) = "`Titel`"
 ArtZ(0, 35) = 26
 ArtZ(1, 35) = 2
 Str(1, 35, 0) = "CREATE TABLE `ebm2000plus` ("
 Str(1, 35, 1) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Leistungsziffer'"
 Str(1, 35, 2) = " `Titel` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kurztext'"
 Str(1, 35, 3) = " `Punktwert` decimal(10,1) DEFAULT NULL COMMENT 'Punktwert'"
 Str(1, 35, 4) = " `Euro` decimal(15,4) DEFAULT NULL COMMENT '€'"
 Str(1, 35, 5) = " `Bericht` bit(1) DEFAULT NULL COMMENT 'Berichtspflicht'"
 Str(1, 35, 6) = " `Text` longtext COLLATE latin1_german2_ci COMMENT 'restlicher Leistungstext'"
 Str(1, 35, 7) = " `Betr` bit(1) DEFAULT NULL COMMENT 'Betreuung'"
 Str(1, 35, 8) = " `Schul` bit(1) DEFAULT NULL COMMENT 'Schulung'"
 Str(1, 35, 9) = " `Typ1` bit(1) DEFAULT NULL"
 Str(1, 35, 10) = " `Typ2` bit(1) DEFAULT NULL"
 Str(1, 35, 11) = " `Gest` bit(1) DEFAULT NULL"
 Str(1, 35, 12) = " `DFS` bit(1) DEFAULT NULL"
 Str(1, 35, 13) = " `DMP` bit(1) DEFAULT NULL"
 Str(1, 35, 14) = " `AOK` bit(1) DEFAULT NULL"
 Str(1, 35, 15) = " `BKK` bit(1) DEFAULT NULL"
 Str(1, 35, 16) = " `BKN` bit(1) DEFAULT NULL"
 Str(1, 35, 17) = " `EK` bit(1) DEFAULT NULL"
 Str(1, 35, 18) = " `IKK` bit(1) DEFAULT NULL"
 Str(1, 35, 19) = " `LKK` bit(1) DEFAULT NULL"
 Str(1, 35, 20) = " `Üw` bit(1) DEFAULT NULL COMMENT 'Überweisung durch HA nötig'"
 Str(1, 35, 21) = " `Insulin` bit(1) DEFAULT NULL"
 Str(1, 35, 22) = " `ICT` bit(1) DEFAULT NULL"
 Str(1, 35, 23) = " `CSII` bit(1) DEFAULT NULL"
 Str(1, 35, 24) = " `Erst` bit(1) DEFAULT NULL"
 Str(1, 35, 25) = " `Folge` bit(1) DEFAULT NULL"
 Str(1, 35, 26) = " `fid` int(11) DEFAULT NULL"
 Str(1, 35, 27) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 35, 28) = "  KEY `Titel` (`Titel`)"
 Str(1, 35, 29) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr35

Function FüllStr36()
 Str(0, 36, 0) = "einstellungen"
 Str(0, 36, 1) = "`ID`"
 Str(0, 36, 2) = "`Formular`"
 Str(0, 36, 3) = "`Abfrage für Formular`"
 Str(0, 36, 4) = "`ID für Formular`"
 Str(0, 36, 5) = "`DatensatzNr`"
 Str(0, 36, 6) = "`ID`"
 Str(0, 36, 7) = "`PrimaryKey`"
 Str(0, 36, 8) = "`Formular`"
 Str(0, 36, 9) = "`ID für Formular`"
 ArtZ(0, 36) = 5
 ArtZ(1, 36) = 4
 Str(1, 36, 0) = "CREATE TABLE `einstellungen` ("
 Str(1, 36, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 36, 2) = " `Formular` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name des Formulars'"
 Str(1, 36, 3) = " `Abfrage für Formular` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name der Abfrage, die zuletzt für das Formular ""Anamnesebogen"" verwendet wurde'"
 Str(1, 36, 4) = " `ID für Formular` int(10) DEFAULT NULL COMMENT 'Pat_ID in dieser Abfrage'"
 Str(1, 36, 5) = " `DatensatzNr` int(10) DEFAULT NULL COMMENT 'Datensatz-Nr. in dieser Abfrage'"
 Str(1, 36, 6) = "  PRIMARY KEY (`ID`)"
 Str(1, 36, 7) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 36, 8) = "  KEY `Formular` (`Formular`)"
 Str(1, 36, 9) = "  KEY `ID für Formular` (`ID für Formular`)"
 Str(1, 36, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr36

Function FüllStr37()
 Str(0, 37, 0) = "eintraege"
 Str(0, 37, 1) = "`FID`"
 Str(0, 37, 2) = "`Pat_ID`"
 Str(0, 37, 3) = "`ZeitPunkt`"
 Str(0, 37, 4) = "`Art`"
 Str(0, 37, 5) = "`Inhalt`"
 Str(0, 37, 6) = "`absPos`"
 Str(0, 37, 7) = "`AktZeit`"
 Str(0, 37, 8) = "`QS`"
 Str(0, 37, 9) = "`QT`"
 Str(0, 37, 10) = "`StByte`"
 Str(0, 37, 11) = "`Auswahl`"
 Str(0, 37, 12) = "`FälleEinträge`"
 Str(0, 37, 13) = "`NamenEinträge`"
 Str(0, 37, 14) = "`Art`"
 Str(0, 37, 15) = "`FälleEinträge_AccRel`"
 Str(0, 37, 16) = "`NamenEinträge_AccRel`"
 ArtZ(0, 37) = 10
 ArtZ(1, 37) = 4
 ArtZ(2, 37) = 2
 Str(1, 37, 0) = "CREATE TABLE `eintraege` ("
 Str(1, 37, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 37, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 37, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 37, 4) = " `Art` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6330'"
 Str(1, 37, 5) = " `Inhalt` longtext COLLATE latin1_german2_ci COMMENT '8480'"
 Str(1, 37, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 37, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 37, 8) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 37, 9) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 37, 10) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 37, 11) = "  KEY `Auswahl` (`Pat_ID`,`Art`,`ZeitPunkt`)"
 Str(1, 37, 12) = "  KEY `FälleEinträge` (`FID`)"
 Str(1, 37, 13) = "  KEY `NamenEinträge` (`Pat_ID`)"
 Str(1, 37, 14) = "  KEY `Art` (`Art`)"
 Str(1, 37, 15) = "  CONSTRAINT `FälleEinträge_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 37, 16) = "  CONSTRAINT `NamenEinträge_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 37, 17) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr37

Function FüllStr38()
 Str(0, 38, 0) = "eintraege arten"
 Str(0, 38, 1) = "`ID`"
 Str(0, 38, 2) = "`Art`"
 Str(0, 38, 3) = "`obAutom`"
 Str(0, 38, 4) = "`Erklärung`"
 Str(0, 38, 5) = "`ID`"
 Str(0, 38, 6) = "`PrimaryKey`"
 ArtZ(0, 38) = 4
 ArtZ(1, 38) = 2
 Str(1, 38, 0) = "CREATE TABLE `eintraege arten` ("
 Str(1, 38, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 38, 2) = " `Art` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 38, 3) = " `obAutom` bit(1) DEFAULT NULL COMMENT 'ob Art auch automatisch aus Formular entsteht'"
 Str(1, 38, 4) = " `Erklärung` longtext COLLATE latin1_german2_ci COMMENT 'Erklärung für die Art'"
 Str(1, 38, 5) = "  PRIMARY KEY (`ID`)"
 Str(1, 38, 6) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 38, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr38

Function FüllStr39()
 Str(0, 39, 0) = "eintragszahlen"
 Str(0, 39, 1) = "`Beginn`"
 Str(0, 39, 2) = "`StByte`"
 Str(0, 39, 3) = "`Zp1`"
 Str(0, 39, 4) = "`Zp2`"
 Str(0, 39, 5) = "`Zp3`"
 Str(0, 39, 6) = "`Zp4`"
 Str(0, 39, 7) = "`Zp5`"
 Str(0, 39, 8) = "`Zp6`"
 Str(0, 39, 9) = "`Zp7`"
 Str(0, 39, 10) = "`Zp8`"
 Str(0, 39, 11) = "`Fallzahl`"
 Str(0, 39, 12) = "`Sekunden`"
 Str(0, 39, 13) = "`Datei`"
 Str(0, 39, 14) = "`DateiAend`"
 Str(0, 39, 15) = "`SpeicherZt`"
 Str(0, 39, 16) = "`TabellenEntleeren`"
 Str(0, 39, 17) = "`ZurücksetzenLAktDat`"
 Str(0, 39, 18) = "`Pat_IDVon`"
 Str(0, 39, 19) = "`Pat_IDbis`"
 Str(0, 39, 20) = "`VorladenFFI`"
 Str(0, 39, 21) = "`ÜberTabelle`"
 Str(0, 39, 22) = "`SammelInsert`"
 Str(0, 39, 23) = "`bereinigeFormInhFeld`"
 Str(0, 39, 24) = "`LaborDirektEinlesen`"
 Str(0, 39, 25) = "`LaborDirektNeu`"
 Str(0, 39, 26) = "`LaborQuerVerb`"
 Str(0, 39, 27) = "`LaborQuerNeu`"
 Str(0, 39, 28) = "`AlterTab`"
 Str(0, 39, 29) = "`obmitEmails`"
 Str(0, 39, 30) = "`LaborPfadBeispiel`"
 Str(0, 39, 31) = "`obVglMitLetzterEinlesung`"
 Str(0, 39, 32) = "`Beginn`"
 Str(0, 39, 33) = "`Beginn`"
 Str(0, 39, 34) = "`stbyte`"
 ArtZ(0, 39) = 31
 ArtZ(1, 39) = 3
 Str(1, 39, 0) = "CREATE TABLE `eintragszahlen` ("
 Str(1, 39, 1) = " `Beginn` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 39, 2) = " `StByte` int(10) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 39, 3) = " `Zp1` datetime DEFAULT NULL"
 Str(1, 39, 4) = " `Zp2` datetime DEFAULT NULL"
 Str(1, 39, 5) = " `Zp3` datetime DEFAULT NULL"
 Str(1, 39, 6) = " `Zp4` datetime DEFAULT NULL"
 Str(1, 39, 7) = " `Zp5` datetime DEFAULT NULL"
 Str(1, 39, 8) = " `Zp6` datetime DEFAULT NULL"
 Str(1, 39, 9) = " `Zp7` datetime DEFAULT NULL"
 Str(1, 39, 10) = " `Zp8` datetime DEFAULT NULL"
 Str(1, 39, 11) = " `Fallzahl` int(10) DEFAULT NULL"
 Str(1, 39, 12) = " `Sekunden` int(10) DEFAULT NULL"
 Str(1, 39, 13) = " `Datei` varchar(120) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 39, 14) = " `DateiAend` datetime DEFAULT NULL"
 Str(1, 39, 15) = " `SpeicherZt` datetime DEFAULT NULL"
 Str(1, 39, 16) = " `TabellenEntleeren` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 17) = " `ZurücksetzenLAktDat` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 18) = " `Pat_IDVon` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 39, 19) = " `Pat_IDbis` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 39, 20) = " `VorladenFFI` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 21) = " `ÜberTabelle` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 22) = " `SammelInsert` tinyint(1) unsigned NOT NULL DEFAULT '0'"
 Str(1, 39, 23) = " `bereinigeFormInhFeld` tinyint(1) unsigned NOT NULL COMMENT 'ob FormInhFeld bereinigt wird'"
 Str(1, 39, 24) = " `LaborDirektEinlesen` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 25) = " `LaborDirektNeu` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 26) = " `LaborQuerVerb` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 27) = " `LaborQuerNeu` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 28) = " `AlterTab` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 29) = " `obmitEmails` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 30) = " `LaborPfadBeispiel` longtext COLLATE latin1_german2_ci"
 Str(1, 39, 31) = " `obVglMitLetzterEinlesung` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 39, 32) = "  PRIMARY KEY (`Beginn`)"
 Str(1, 39, 33) = "  UNIQUE KEY `Beginn` (`Beginn`)"
 Str(1, 39, 34) = "  UNIQUE KEY `stbyte` (`StByte`) USING BTREE"
 Str(1, 39, 35) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Function ' FüllStr39

Function FüllStr40()
 Str(0, 40, 0) = "eintrhist"
 Str(1, 40, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `eintrhist` AS select `eintrhist1`.`ID` AS `ID`,`eintrhist1`.`Pat_ID` AS `Pat_ID`,`eintrhist1`.`ZeitPunkt` AS `ZeitPunkt`,`eintrhist1`.`Art` AS `Art`,`eintrhist1`.`Inhalt` AS `Inhalt`,`eintrhist1`.`QS` AS `QS`,`eintrhist1`.`QT` AS `QT`,`eintrhist2`.`FID` AS `FID`,`eintrhist2`.`absPos` AS `absPos`,`eintrhist2`.`AktZeit` AS `AktZeit`,`eintrhist2`.`StByte` AS `StByte` from (`eintrhist1` join `eintrhist2` on((`eintrhist1`.`ID` = `eintrhist2`.`ID`)))"
End Function ' FüllStr40

Function FüllStr41()
 Str(0, 41, 0) = "eintrhist1"
 Str(0, 41, 1) = "`Pat_ID`"
 Str(0, 41, 2) = "`ZeitPunkt`"
 Str(0, 41, 3) = "`Art`"
 Str(0, 41, 4) = "`Inhalt`"
 Str(0, 41, 5) = "`QS`"
 Str(0, 41, 6) = "`QT`"
 Str(0, 41, 7) = "`ID`"
 Str(0, 41, 8) = "`ID`"
 Str(0, 41, 9) = "`Auswahl`"
 Str(0, 41, 10) = "`NamenEinträge`"
 Str(0, 41, 11) = "`Art`"
 ArtZ(0, 41) = 7
 ArtZ(1, 41) = 4
 Str(1, 41, 0) = "CREATE TABLE `eintrhist1` ("
 Str(1, 41, 1) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 41, 2) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 41, 3) = " `Art` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6330'"
 Str(1, 41, 4) = " `Inhalt` longtext COLLATE latin1_german2_ci COMMENT '8480'"
 Str(1, 41, 5) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 41, 6) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 41, 7) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 41, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 41, 9) = "  KEY `Auswahl` (`Pat_ID`,`Art`,`ZeitPunkt`)"
 Str(1, 41, 10) = "  KEY `NamenEinträge` (`Pat_ID`)"
 Str(1, 41, 11) = "  KEY `Art` (`Art`)"
 Str(1, 41, 12) = " ENGINE=InnoDB AUTO_INCREMENT=1017 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr41

Function FüllStr42()
 Str(0, 42, 0) = "eintrhist2"
 Str(0, 42, 1) = "`FID`"
 Str(0, 42, 2) = "`absPos`"
 Str(0, 42, 3) = "`AktZeit`"
 Str(0, 42, 4) = "`StByte`"
 Str(0, 42, 5) = "`ID`"
 Str(0, 42, 6) = "`FälleEinträge`"
 Str(0, 42, 7) = "`FK_id`"
 Str(0, 42, 8) = "`FK_id`"
 ArtZ(0, 42) = 5
 ArtZ(1, 42) = 2
 ArtZ(2, 42) = 1
 Str(1, 42, 0) = "CREATE TABLE `eintrhist2` ("
 Str(1, 42, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 42, 2) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 42, 3) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 42, 4) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 42, 5) = " `ID` int(10) unsigned NOT NULL DEFAULT '0'"
 Str(1, 42, 6) = "  KEY `FälleEinträge` (`FID`)"
 Str(1, 42, 7) = "  KEY `FK_id` (`ID`)"
 Str(1, 42, 8) = "  CONSTRAINT `FK_id` FOREIGN KEY (`ID`) REFERENCES `eintrhist1` (`ID`)"
 Str(1, 42, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr42

Function FüllStr43()
 Str(0, 43, 0) = "faelle"
 Str(0, 43, 1) = "`FID`"
 Str(0, 43, 2) = "`Pat_ID`"
 Str(0, 43, 3) = "`Quartal`"
 Str(0, 43, 4) = "`Nachname`"
 Str(0, 43, 5) = "`Vorname`"
 Str(0, 43, 6) = "`lfdnr`"
 Str(0, 43, 7) = "`TMFNr`"
 Str(0, 43, 8) = "`VKNr`"
 Str(0, 43, 9) = "`BhFB`"
 Str(0, 43, 10) = "`BhFE1`"
 Str(0, 43, 11) = "`BhFE2`"
 Str(0, 43, 12) = "`f4202`"
 Str(0, 43, 13) = "`ausgst`"
 Str(0, 43, 14) = "`KtrAbrB`"
 Str(0, 43, 15) = "`AbrAr`"
 Str(0, 43, 16) = "`lVorl`"
 Str(0, 43, 17) = "`IK`"
 Str(0, 43, 18) = "`KVKs`"
 Str(0, 43, 19) = "`KVKserg`"
 Str(0, 43, 20) = "`Kasse`"
 Str(0, 43, 21) = "`GebOr`"
 Str(0, 43, 22) = "`AbrGb`"
 Str(0, 43, 23) = "`PersKreis`"
 Str(0, 43, 24) = "`SKtZusatz`"
 Str(0, 43, 25) = "`f4206`"
 Str(0, 43, 26) = "`ÜwText`"
 Str(0, 43, 27) = "`f4210`"
 Str(0, 43, 28) = "`AkfHAH`"
 Str(0, 43, 29) = "`AkfAB0`"
 Str(0, 43, 30) = "`AkfAK`"
 Str(0, 43, 31) = "`statNuller`"
 Str(0, 43, 32) = "`ÜbwV`"
 Str(0, 43, 33) = "`AndÜw`"
 Str(0, 43, 34) = "`Übw`"
 Str(0, 43, 35) = "`ÜbwLANR`"
 Str(0, 43, 36) = "`ÜWZiel`"
 Str(0, 43, 37) = "`ÜWNNr`"
 Str(0, 43, 38) = "`ÜWNaN`"
 Str(0, 43, 39) = "`ÜWTit`"
 Str(0, 43, 40) = "`ÜWVor`"
 Str(0, 43, 41) = "`ÜWVsw`"
 Str(0, 43, 42) = "`üwvid`"
 Str(0, 43, 43) = "`statKlasse`"
 Str(0, 43, 44) = "`f4237`"
 Str(0, 43, 45) = "`statBehTage`"
 Str(0, 43, 46) = "`SchGr`"
 Str(0, 43, 47) = "`Weiterbeh`"
 Str(0, 43, 48) = "`PGeb`"
 Str(0, 43, 49) = "`PGebErg`"
 Str(0, 43, 50) = "`Mahnfrist`"
 Str(0, 43, 51) = "`GOÄKatNr`"
 Str(0, 43, 52) = "`GOÄKatName`"
 Str(0, 43, 53) = "`abrArzt`"
 Str(0, 43, 54) = "`privVers`"
 Str(0, 43, 55) = "`AdNam`"
 Str(0, 43, 56) = "`AdStr`"
 Str(0, 43, 57) = "`AdPlz`"
 Str(0, 43, 58) = "`AdOrt`"
 Str(0, 43, 59) = "`BhFE`"
 Str(0, 43, 60) = "`s8000`"
 Str(0, 43, 61) = "`s8100`"
 Str(0, 43, 62) = "`AktZeit`"
 Str(0, 43, 63) = "`Fanf`"
 Str(0, 43, 64) = "`altQuart`"
 Str(0, 43, 65) = "`QAnf`"
 Str(0, 43, 66) = "`QEnd`"
 Str(0, 43, 67) = "`QS`"
 Str(0, 43, 68) = "`QT`"
 Str(0, 43, 69) = "`TherArt`"
 Str(0, 43, 70) = "`StByte`"
 Str(0, 43, 71) = "`absPos`"
 Str(0, 43, 72) = "`FID`"
 Str(0, 43, 73) = "`PrimaryKey`"
 Str(0, 43, 74) = "`AktF`"
 Str(0, 43, 75) = "`Auswahl`"
 Str(0, 43, 76) = "`BhFB`"
 Str(0, 43, 77) = "`FanF`"
 Str(0, 43, 78) = "`NamenFälle`"
 Str(0, 43, 79) = "`pQ`"
 Str(0, 43, 80) = "`Quartal`"
 Str(0, 43, 81) = "`SchGr`"
 Str(0, 43, 82) = "`vknr`"
 Str(0, 43, 83) = "`KassenlisteFälle_AccRel`"
 Str(0, 43, 84) = "`NamenFälle_AccRel`"
 ArtZ(0, 43) = 71
 ArtZ(1, 43) = 11
 ArtZ(2, 43) = 2
 Str(1, 43, 0) = "CREATE TABLE `faelle` ("
 Str(1, 43, 1) = " `FID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'"
 Str(1, 43, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 43, 3) = " `Quartal` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4101'"
 Str(1, 43, 4) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 43, 5) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 43, 6) = " `lfdnr` int(10) DEFAULT NULL COMMENT 'laufende Fallnummer'"
 Str(1, 43, 7) = " `TMFNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4144 Fallnummer in Turbomed'"
 Str(1, 43, 8) = " `VKNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4104'"
 Str(1, 43, 9) = " `BhFB` datetime DEFAULT NULL COMMENT '4150'"
 Str(1, 43, 10) = " `BhFE1` datetime DEFAULT NULL COMMENT '4151'"
 Str(1, 43, 11) = " `BhFE2` datetime DEFAULT NULL COMMENT '4152'"
 Str(1, 43, 12) = " `f4202` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4202'"
 Str(1, 43, 13) = " `ausgst` datetime DEFAULT NULL COMMENT '4102 (''ausgestellt am'')'"
 Str(1, 43, 14) = " `KtrAbrB` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))'"
 Str(1, 43, 15) = " `AbrAr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4107, Abrechnungsart (1 = Primärkassen)'"
 Str(1, 43, 16) = " `lVorl` datetime DEFAULT NULL COMMENT '4109, letzte Vorlage'"
 Str(1, 43, 17) = " `IK` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4111 Krankenkassennummer (IK)'"
 Str(1, 43, 18) = " `KVKs` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4112 Versichertenstatus VK'"
 Str(1, 43, 19) = " `KVKserg` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4113 Ost/West-Status VK'"
 Str(1, 43, 20) = " `Kasse` varchar(70) COLLATE latin1_german2_ci NOT NULL COMMENT '6299 Kasse (aus Formularen)'"
 Str(1, 43, 21) = " `GebOr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4121, Gebührenordnung (1 = BMÄ, 2)'"
 Str(1, 43, 22) = " `AbrGb` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4122, Abrechnungsgebiet (07 = Diabetes)'"
 Str(1, 43, 23) = " `PersKreis` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4123 Personenkreis/Untersuchungskategorie'"
 Str(1, 43, 24) = " `SKtZusatz` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4124 SKT-Zusatzangaben'"
 Str(1, 43, 25) = " `f4206` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4206, mutmasslicher Tag der Entbindung'"
 Str(1, 43, 26) = " `ÜwText` longtext COLLATE latin1_german2_ci COMMENT '4209: Auftrags- / erläuternder Text zur Überweisung'"
 Str(1, 43, 27) = " `f4210` tinyint(1) unsigned DEFAULT NULL COMMENT '4210, Ankreuzfeld LSR'"
 Str(1, 43, 28) = " `AkfHAH` tinyint(1) unsigned DEFAULT NULL COMMENT '4211 Ankreuzfeld HAH'"
 Str(1, 43, 29) = " `AkfAB0` tinyint(1) unsigned DEFAULT NULL COMMENT '4212 Ankreuzfeld AB0.RH'"
 Str(1, 43, 30) = " `AkfAK` tinyint(1) unsigned DEFAULT NULL COMMENT '4213 Ankreuzfeld AK'"
 Str(1, 43, 31) = " `statNuller` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4216, nu bei Musterfrau 16 Nuller'"
 Str(1, 43, 32) = " `ÜbwV` varchar(22) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4218, überwiesen von Arztnummer'"
 Str(1, 43, 33) = " `AndÜw` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4219, anderer Überweiser'"
 Str(1, 43, 34) = " `Übw` varchar(8) COLLATE latin1_german2_ci NOT NULL COMMENT '4218 oder 4219, je nachdem, was befüllt'"
 Str(1, 43, 35) = " `ÜbwLANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4242 LANR des Überweisers'"
 Str(1, 43, 36) = " `ÜWZiel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4220 Überweisung an'"
 Str(1, 43, 37) = " `ÜWNNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(4): KV-Nummer des Überweisers'"
 Str(1, 43, 38) = " `ÜWNaN` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(3): Nachname des Überweisers'"
 Str(1, 43, 39) = " `ÜWTit` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(3): Titel des Überweisers'"
 Str(1, 43, 40) = " `ÜWVor` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(2): Vorname des Überweisers'"
 Str(1, 43, 41) = " `ÜWVsw` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(2b): Vorsatzwort des Überweisers'"
 Str(1, 43, 42) = " `üwvid` int(10) unsigned NOT NULL COMMENT '4247 Bezug auf ueberwvon'"
 Str(1, 43, 43) = " `statKlasse` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4236 Klasse bei Behandlung'"
 Str(1, 43, 44) = " `f4237` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4237 Krankenhausname'"
 Str(1, 43, 45) = " `statBehTage` int(10) DEFAULT NULL COMMENT '4238 Krankenhausaufenthalt'"
 Str(1, 43, 46) = " `SchGr` decimal(2,0) DEFAULT NULL COMMENT '4239, Schein(unter)gruppe'"
 Str(1, 43, 47) = " `Weiterbeh` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4243, Weiterbehandelnder'"
 Str(1, 43, 48) = " `PGeb` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4401, Praxisgebühr'"
 Str(1, 43, 49) = " `PGebErg` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4402, ?'"
 Str(1, 43, 50) = " `Mahnfrist` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4403, Mahnfrist bis'"
 Str(1, 43, 51) = " `GOÄKatNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4580 (1): Katalog-Nummer'"
 Str(1, 43, 52) = " `GOÄKatName` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4580 (2): Privat-Abrechnungskatalog'"
 Str(1, 43, 53) = " `abrArzt` varchar(30) COLLATE latin1_german2_ci NOT NULL COMMENT '4585 abrechnender Arzt'"
 Str(1, 43, 54) = " `privVers` varchar(45) COLLATE latin1_german2_ci NOT NULL COMMENT '4586 private Versicherung'"
 Str(1, 43, 55) = " `AdNam` varchar(28) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT '4602(1) Name Rechnungsanschrift'"
 Str(1, 43, 56) = " `AdStr` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(2) Straße Rechnungsanschrift'"
 Str(1, 43, 57) = " `AdPlz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(3) PLZ Rechnungsanschrift'"
 Str(1, 43, 58) = " `AdOrt` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(4) Ort Rechnungsanschrift'"
 Str(1, 43, 59) = " `BhFE` datetime DEFAULT NULL COMMENT '4604, Behandlungsfall: Ende, bei Privatpatienten'"
 Str(1, 43, 60) = " `s8000` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000, Satzidentifikation'"
 Str(1, 43, 61) = " `s8100` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge'"
 Str(1, 43, 62) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 43, 63) = " `Fanf` datetime DEFAULT NULL COMMENT 'Fallanfang'"
 Str(1, 43, 64) = " `altQuart` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 43, 65) = " `QAnf` datetime DEFAULT NULL COMMENT 'Quartalsanfang'"
 Str(1, 43, 66) = " `QEnd` datetime DEFAULT NULL COMMENT 'Quartalsende'"
 Str(1, 43, 67) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 43, 68) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 43, 69) = " `TherArt` int(2) unsigned DEFAULT NULL COMMENT 'Therapieart: (0 = offen,  1= diät,  2= oad, 3= komb,  4= ct, 5= ict, 6 = csii)'"
 Str(1, 43, 70) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 43, 71) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 43, 72) = "  PRIMARY KEY (`FID`)"
 Str(1, 43, 73) = "  UNIQUE KEY `PrimaryKey` (`FID`)"
 Str(1, 43, 74) = "  KEY `AktF` (`Pat_ID`,`BhFB`)"
 Str(1, 43, 75) = "  KEY `Auswahl` (`Pat_ID`,`Quartal`,`BhFB`,`BhFE1`)"
 Str(1, 43, 76) = "  KEY `BhFB` (`BhFB`)"
 Str(1, 43, 77) = "  KEY `FanF` (`Fanf`)"
 Str(1, 43, 78) = "  KEY `NamenFälle` (`Pat_ID`)"
 Str(1, 43, 79) = "  KEY `pQ` (`Pat_ID`,`Quartal`)"
 Str(1, 43, 80) = "  KEY `Quartal` (`Quartal`)"
 Str(1, 43, 81) = "  KEY `SchGr` (`SchGr`,`Nachname`,`Vorname`)"
 Str(1, 43, 82) = "  KEY `vknr` (`VKNr`)"
 Str(1, 43, 83) = "  CONSTRAINT `KassenlisteFälle_AccRel` FOREIGN KEY (`VKNr`) REFERENCES `kassenliste` (`VK`)"
 Str(1, 43, 84) = "  CONSTRAINT `NamenFälle_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 43, 85) = " ENGINE=InnoDB AUTO_INCREMENT=26078 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr43

Function FüllStr44()
 Str(0, 44, 0) = "faelleverschieden"
 Str(1, 44, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `faelleverschieden` AS select `f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`f4206` AS `f4206`,`f`.`ÜwText` AS `ÜwText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`ÜbwV` AS `ÜbwV`,`f`.`" & _
  "AndÜw` AS `AndÜw`,`f`.`Übw` AS `Übw`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8100` AS `s8100`,`f`.`AktZeit` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`TherArt` AS `T" & _
  "herArt`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos` from `_faellenachschgr` `f` group by `f`.`Pat_ID`,`f`.`Quartal`"
End Function ' FüllStr44

Function FüllStr45()
 Str(0, 45, 0) = "faelleverschiedenneu"
 Str(1, 45, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `faelleverschiedenneu` AS select `f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`f4206` AS `f4206`,`f`.`ÜwText` AS `ÜwText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`ÜbwV` AS `ÜbwV`,`f" & _
  "`.`AndÜw` AS `AndÜw`,`f`.`Übw` AS `Übw`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8100` AS `s8100`,`f`.`AktZeit` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`TherArt` AS" & _
  " `TherArt`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos`,(select min(`f1`.`Fanf`) AS `min(fanf)` from `faelle` `f1` where ((`f1`.`Pat_ID` = `f`.`Pat_ID`) and (`f1`.`Fanf` < `f`.`Fanf`))) AS `erst` from `_faellenachschgr` `f` group by `f`.`Pat_ID`,`f`.`Quartal`"
End Function ' FüllStr45

Function FüllStr46()
 Str(0, 46, 0) = "fallzahlen"
 Str(0, 46, 1) = "`Zahl`"
 Str(0, 46, 2) = "`DZahl`"
 Str(0, 46, 3) = "`NDZahl`"
 Str(0, 46, 4) = "`Quartal`"
 ArtZ(0, 46) = 4
 Str(1, 46, 0) = "CREATE TABLE `fallzahlen` ("
 Str(1, 46, 1) = " `Zahl` bigint(21) DEFAULT NULL"
 Str(1, 46, 2) = " `DZahl` decimal(32,0) DEFAULT NULL"
 Str(1, 46, 3) = " `NDZahl` decimal(33,0) DEFAULT NULL"
 Str(1, 46, 4) = " `Quartal` varchar(5) DEFAULT NULL"
 Str(1, 46, 5) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Function ' FüllStr46

Function FüllStr47()
 Str(0, 47, 0) = "fallzahlstand"
 Str(0, 47, 1) = "`zahl`"
 Str(0, 47, 2) = "`z0`"
 Str(0, 47, 3) = "`quartal`"
 ArtZ(0, 47) = 3
 Str(1, 47, 0) = "CREATE TABLE `fallzahlstand` ("
 Str(1, 47, 1) = " `zahl` bigint(21) DEFAULT NULL"
 Str(1, 47, 2) = " `z0` bigint(22) DEFAULT NULL"
 Str(1, 47, 3) = " `quartal` varchar(5) DEFAULT NULL"
 Str(1, 47, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Function ' FüllStr47

Function FüllStr48()
 Str(0, 48, 0) = "fallzahlstand 1"
 Str(1, 48, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 1` AS select count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` from `faelleverschiedenneu` `f` where ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 1 day)) - to_days(concat(year((now() - interval 1 day)),'-',((((month((now() - interval 1 day)) - 1) DIV 3) * 3) + 1),'-01')))) and ((`f`.`Pat_ID` < 3044) or (`f`.`Pat_ID` > 50000))) group by `f`.`Quartal` order by substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Function ' FüllStr48

Function FüllStr49()
 Str(0, 49, 0) = "fallzahlstand 2"
 Str(1, 49, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 2` AS select count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` from `faelleverschiedenneu` `f` where ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 1 day)) - to_days(concat(year((now() - interval 1 day)),'-',((((month((now() - interval 1 day)) - 1) DIV 3) * 3) + 1),'-01')))) and (`f`.`Pat_ID` > 3044)) group by `f`.`Quartal` order by substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Function ' FüllStr49

Function FüllStr50()
 Str(0, 50, 0) = "fallzahlstand 3"
 Str(1, 50, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 3` AS select count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` from `faelleverschiedenneu` `f` where ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 1 day)) - to_days(concat(year((now() - interval 1 day)),'-',((((month((now() - interval 1 day)) - 1) DIV 3) * 3) + 1),'-01')))) and 1) group by `f`.`Quartal` order by substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Function ' FüllStr50

Function FüllStr51()
 Str(0, 51, 0) = "forminhaltfeld"
 Str(0, 51, 1) = "`FeldVW`"
 Str(0, 51, 2) = "`Feld`"
 Str(0, 51, 3) = "`StByte`"
 Str(0, 51, 4) = "`FeldVW`"
 Str(0, 51, 5) = "`Feld`"
 ArtZ(0, 51) = 3
 ArtZ(1, 51) = 2
 Str(1, 51, 0) = "CREATE TABLE `forminhaltfeld` ("
 Str(1, 51, 1) = " `FeldVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 51, 2) = " `Feld` longtext COLLATE latin1_german2_ci"
 Str(1, 51, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 51, 4) = "  PRIMARY KEY (`FeldVW`)"
 Str(1, 51, 5) = "  KEY `Feld` (`Feld`(255))"
 Str(1, 51, 6) = " ENGINE=InnoDB AUTO_INCREMENT=13315 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr51

Function FüllStr52()
 Str(0, 52, 0) = "forminhaltfeldinh"
 Str(0, 52, 1) = "`FeldInhVW`"
 Str(0, 52, 2) = "`FeldInh`"
 Str(0, 52, 3) = "`StByte`"
 Str(0, 52, 4) = "`FeldInhVW`"
 Str(0, 52, 5) = "`FeldInhVW`"
 Str(0, 52, 6) = "`FeldInh`"
 ArtZ(0, 52) = 3
 ArtZ(1, 52) = 3
 Str(1, 52, 0) = "CREATE TABLE `forminhaltfeldinh` ("
 Str(1, 52, 1) = " `FeldInhVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 52, 2) = " `FeldInh` longtext COLLATE latin1_german2_ci"
 Str(1, 52, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 52, 4) = "  PRIMARY KEY (`FeldInhVW`)"
 Str(1, 52, 5) = "  UNIQUE KEY `FeldInhVW` (`FeldInhVW`)"
 Str(1, 52, 6) = "  KEY `FeldInh` (`FeldInh`(255))"
 Str(1, 52, 7) = " ENGINE=InnoDB AUTO_INCREMENT=50114 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr52

Function FüllStr53()
 Str(0, 53, 0) = "forminhaltform_abk"
 Str(0, 53, 1) = "`Form_AbkVW`"
 Str(0, 53, 2) = "`Form_Abk`"
 Str(0, 53, 3) = "`Form_AbkVW`"
 Str(0, 53, 4) = "`Form_AbkVW`"
 Str(0, 53, 5) = "`Form_Abk`"
 ArtZ(0, 53) = 2
 ArtZ(1, 53) = 3
 Str(1, 53, 0) = "CREATE TABLE `forminhaltform_abk` ("
 Str(1, 53, 1) = " `Form_AbkVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 53, 2) = " `Form_Abk` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 53, 3) = "  PRIMARY KEY (`Form_AbkVW`)"
 Str(1, 53, 4) = "  UNIQUE KEY `Form_AbkVW` (`Form_AbkVW`)"
 Str(1, 53, 5) = "  UNIQUE KEY `Form_Abk` (`Form_Abk`)"
 Str(1, 53, 6) = " ENGINE=InnoDB AUTO_INCREMENT=41 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr53

Function FüllStr54()
 Str(0, 54, 0) = "forminhfeld"
 Str(0, 54, 1) = "`FoID`"
 Str(0, 54, 2) = "`Nr`"
 Str(0, 54, 3) = "`FeldNr`"
 Str(0, 54, 4) = "`FeldVW`"
 Str(0, 54, 5) = "`FeldInhVW`"
 Str(0, 54, 6) = "`FoID`"
 Str(0, 54, 7) = "`FormInhaltFeldFormInhFeld`"
 Str(0, 54, 8) = "`FormInhaltFeldInhFormInhFeld`"
 Str(0, 54, 9) = "`FormInhFeldFormInhaltFeld`"
 Str(0, 54, 10) = "`FormInhFeldFormInhaltFeldInhalt`"
 ArtZ(0, 54) = 5
 ArtZ(1, 54) = 3
 ArtZ(2, 54) = 2
 Str(1, 54, 0) = "CREATE TABLE `forminhfeld` ("
 Str(1, 54, 1) = " `FoID` int(10) DEFAULT NULL"
 Str(1, 54, 2) = " `Nr` smallint(6) DEFAULT NULL"
 Str(1, 54, 3) = " `FeldNr` smallint(6) DEFAULT NULL"
 Str(1, 54, 4) = " `FeldVW` int(10) DEFAULT NULL"
 Str(1, 54, 5) = " `FeldInhVW` int(10) DEFAULT NULL"
 Str(1, 54, 6) = "  KEY `FoID` (`FoID`,`FeldVW`,`FeldNr`)"
 Str(1, 54, 7) = "  KEY `FormInhaltFeldFormInhFeld` (`FeldVW`)"
 Str(1, 54, 8) = "  KEY `FormInhaltFeldInhFormInhFeld` (`FeldInhVW`)"
 Str(1, 54, 9) = "  CONSTRAINT `FormInhFeldFormInhaltFeld` FOREIGN KEY (`FeldVW`) REFERENCES `forminhaltfeld` (`FeldVW`)"
 Str(1, 54, 10) = "  CONSTRAINT `FormInhFeldFormInhaltFeldInhalt` FOREIGN KEY (`FeldInhVW`) REFERENCES `forminhaltfeldinh` (`FeldInhVW`)"
 Str(1, 54, 11) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Function ' FüllStr54

Function FüllStr55()
 Str(0, 55, 0) = "forminhkopf"
 Str(0, 55, 1) = "`FoID`"
 Str(0, 55, 2) = "`FID`"
 Str(0, 55, 3) = "`Pat_ID`"
 Str(0, 55, 4) = "`Form_ID`"
 Str(0, 55, 5) = "`ZeitPunkt`"
 Str(0, 55, 6) = "`AbsPos`"
 Str(0, 55, 7) = "`AktZeit`"
 Str(0, 55, 8) = "`StByte`"
 Str(0, 55, 9) = "`Satzart`"
 Str(0, 55, 10) = "`Satzlänge`"
 Str(0, 55, 11) = "`FoID`"
 Str(0, 55, 12) = "`PrimaryKey`"
 Str(0, 55, 13) = "`Auswahl`"
 Str(0, 55, 14) = "`FälleFormInhKopf`"
 Str(0, 55, 15) = "`FID`"
 Str(0, 55, 16) = "`FormulareFormInhKopf`"
 Str(0, 55, 17) = "`NamenFormInhKopf`"
 Str(0, 55, 18) = "`FälleFormInhKopf_AccRel`"
 Str(0, 55, 19) = "`FormulareFormInhKopf_AccRel`"
 Str(0, 55, 20) = "`NamenFormInhKopf_AccRel`"
 ArtZ(0, 55) = 10
 ArtZ(1, 55) = 7
 ArtZ(2, 55) = 3
 Str(1, 55, 0) = "CREATE TABLE `forminhkopf` ("
 Str(1, 55, 1) = " `FoID` int(10) NOT NULL DEFAULT '0'"
 Str(1, 55, 2) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 55, 3) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 55, 4) = " `Form_ID` int(10) DEFAULT NULL"
 Str(1, 55, 5) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 55, 6) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 55, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 55, 8) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 55, 9) = " `Satzart` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000'"
 Str(1, 55, 10) = " `Satzlänge` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100'"
 Str(1, 55, 11) = "  PRIMARY KEY (`FoID`)"
 Str(1, 55, 12) = "  UNIQUE KEY `PrimaryKey` (`FoID`)"
 Str(1, 55, 13) = "  KEY `Auswahl` (`Pat_ID`,`Form_ID`,`ZeitPunkt`)"
 Str(1, 55, 14) = "  KEY `FälleFormInhKopf` (`FID`)"
 Str(1, 55, 15) = "  KEY `FID` (`FID`)"
 Str(1, 55, 16) = "  KEY `FormulareFormInhKopf` (`Form_ID`)"
 Str(1, 55, 17) = "  KEY `NamenFormInhKopf` (`Pat_ID`)"
 Str(1, 55, 18) = "  CONSTRAINT `FälleFormInhKopf_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 55, 19) = "  CONSTRAINT `FormulareFormInhKopf_AccRel` FOREIGN KEY (`Form_ID`) REFERENCES `formulare` (`FormID`) ON UPDATE CASCADE"
 Str(1, 55, 20) = "  CONSTRAINT `NamenFormInhKopf_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 55, 21) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr55

Function FüllStr56()
 Str(0, 56, 0) = "formular"
 Str(1, 56, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `formular` AS select `forminhkopf`.`FoID` AS `foid`,`forminhkopf`.`Pat_ID` AS `Pat_ID`,`forminhkopf`.`FID` AS `FID`,`forminhkopf`.`Form_ID` AS `Form_ID`,`forminhkopf`.`ZeitPunkt` AS `ZeitPunkt`,`forminhfeld`.`Nr` AS `Nr`,`forminhfeld`.`FeldNr` AS `FeldNr`,`forminhaltfeld`.`Feld` AS `Feld`,`forminhaltfeldinh`.`FeldInh` AS `FeldInh`,`formulare`.`Form_Abk` AS `form_abk`,`formulare`.`FormVorl` AS `FormVorl` from ((((`forminhfeld` left join `forminhkopf` on((`forminhfeld`.`FoID` = `forminhkopf`.`FoID`))) left join `formulare` on((`formulare`.`FormID` = `forminhkopf`.`Form_ID`))) left join `forminhaltfeld` on((`forminhfeld`.`FeldVW` = `forminhaltfeld`.`FeldVW`))) left join `forminhaltfeldinh` on((`forminhfeld`.`FeldInhVW` = `forminhaltfeldinh`.`FeldInhVW`))) order by `forminhkopf`.`FoID`"
End Function ' FüllStr56

Function FüllStr57()
 Str(0, 57, 0) = "formulare"
 Str(0, 57, 1) = "`FormID`"
 Str(0, 57, 2) = "`Form_Abk`"
 Str(0, 57, 3) = "`FormBez`"
 Str(0, 57, 4) = "`FormVorl`"
 Str(0, 57, 5) = "`AktZeit`"
 Str(0, 57, 6) = "`absPos`"
 Str(0, 57, 7) = "`StByte`"
 Str(0, 57, 8) = "`FormID`"
 Str(0, 57, 9) = "`FormID`"
 Str(0, 57, 10) = "`Auswahl`"
 Str(0, 57, 11) = "`FormBez`"
 Str(0, 57, 12) = "`FormInhaltForm_AbkFormulare`"
 Str(0, 57, 13) = "`FormInhaltForm_AbkFormulare_AccRel`"
 ArtZ(0, 57) = 7
 ArtZ(1, 57) = 5
 ArtZ(2, 57) = 1
 Str(1, 57, 0) = "CREATE TABLE `formulare` ("
 Str(1, 57, 1) = " `FormID` int(10) NOT NULL DEFAULT '0'"
 Str(1, 57, 2) = " `Form_Abk` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 57, 3) = " `FormBez` longtext COLLATE latin1_german2_ci"
 Str(1, 57, 4) = " `FormVorl` varchar(114) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 57, 5) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Aktualisierung'"
 Str(1, 57, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 57, 7) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 57, 8) = "  PRIMARY KEY (`FormID`)"
 Str(1, 57, 9) = "  UNIQUE KEY `FormID` (`FormID`)"
 Str(1, 57, 10) = "  KEY `Auswahl` (`Form_Abk`,`FormBez`(255),`FormVorl`)"
 Str(1, 57, 11) = "  KEY `FormBez` (`FormBez`(255))"
 Str(1, 57, 12) = "  KEY `FormInhaltForm_AbkFormulare` (`Form_Abk`)"
 Str(1, 57, 13) = "  CONSTRAINT `FormInhaltForm_AbkFormulare_AccRel` FOREIGN KEY (`Form_Abk`) REFERENCES `forminhaltform_abk` (`Form_Abk`)"
 Str(1, 57, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Function ' FüllStr57

Function FüllStr58()
 Str(0, 58, 0) = "fuerdiagexp"
 Str(0, 58, 1) = "`ID`"
 Str(0, 58, 2) = "`Name`"
 Str(0, 58, 3) = "`Pat_id`"
 Str(0, 58, 4) = "`ICD`"
 Str(0, 58, 5) = "`Diagnose`"
 Str(0, 58, 6) = "`Status`"
 Str(0, 58, 7) = "`Protokoll`"
 Str(0, 58, 8) = "`nurQuart`"
 Str(0, 58, 9) = "`Zeitpunkt`"
 Str(0, 58, 10) = "`ID`"
 Str(0, 58, 11) = "`ID`"
 Str(0, 58, 12) = "`pat_ID`"
 Str(0, 58, 13) = "`Suche`"
 ArtZ(0, 58) = 9
 ArtZ(1, 58) = 4
 Str(1, 58, 0) = "CREATE TABLE `fuerdiagexp` ("
 Str(1, 58, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 58, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 58, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 58, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 58, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 58, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 58, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 58, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 58, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 58, 10) = "  PRIMARY KEY (`ID`)"
 Str(1, 58, 11) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 58, 12) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 58, 13) = "  KEY `Suche` (`Pat_id`,`ICD`,`Diagnose`)"
 Str(1, 58, 14) = " ENGINE=InnoDB AUTO_INCREMENT=2158 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr58

Function FüllStr59()
 Str(0, 59, 0) = "fuerdiagexparchiv"
 Str(0, 59, 1) = "`ID`"
 Str(0, 59, 2) = "`Name`"
 Str(0, 59, 3) = "`Pat_id`"
 Str(0, 59, 4) = "`ICD`"
 Str(0, 59, 5) = "`Diagnose`"
 Str(0, 59, 6) = "`Status`"
 Str(0, 59, 7) = "`Protokoll`"
 Str(0, 59, 8) = "`nurQuart`"
 Str(0, 59, 9) = "`Zeitpunkt`"
 Str(0, 59, 10) = "`archiviert`"
 Str(0, 59, 11) = "`ID`"
 Str(0, 59, 12) = "`ID`"
 Str(0, 59, 13) = "`pat_ID`"
 Str(0, 59, 14) = "`Suche`"
 ArtZ(0, 59) = 10
 ArtZ(1, 59) = 4
 Str(1, 59, 0) = "CREATE TABLE `fuerdiagexparchiv` ("
 Str(1, 59, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 59, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 59, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 59, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 59, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 59, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 59, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 59, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 59, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 59, 10) = " `archiviert` datetime DEFAULT NULL COMMENT 'Datum der Übertragung in Tabelle Leistungsexport'"
 Str(1, 59, 11) = "  PRIMARY KEY (`ID`)"
 Str(1, 59, 12) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 59, 13) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 59, 14) = "  KEY `Suche` (`Pat_id`,`ICD`,`Diagnose`)"
 Str(1, 59, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr59

Function FüllStr60()
 Str(0, 60, 0) = "fuerleistexp"
 Str(0, 60, 1) = "`ID`"
 Str(0, 60, 2) = "`Name`"
 Str(0, 60, 3) = "`Datum`"
 Str(0, 60, 4) = "`Pat_id`"
 Str(0, 60, 5) = "`SchGr`"
 Str(0, 60, 6) = "`Status`"
 Str(0, 60, 7) = "`Protokoll`"
 Str(0, 60, 8) = "`ID`"
 Str(0, 60, 9) = "`PrimaryKey`"
 Str(0, 60, 10) = "`ID`"
 Str(0, 60, 11) = "`NamenfürLeistExp`"
 Str(0, 60, 12) = "`NamenfürLeistExp_AccRel`"
 ArtZ(0, 60) = 7
 ArtZ(1, 60) = 4
 ArtZ(2, 60) = 1
 Str(1, 60, 0) = "CREATE TABLE `fuerleistexp` ("
 Str(1, 60, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 60, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 60, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen'"
 Str(1, 60, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 60, 5) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 60, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 60, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 60, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 60, 9) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 60, 10) = "  KEY `ID` (`Pat_id`)"
 Str(1, 60, 11) = "  KEY `NamenfürLeistExp` (`Pat_id`)"
 Str(1, 60, 12) = "  CONSTRAINT `NamenfürLeistExp_AccRel` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 60, 13) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr60

Function FüllStr61()
 Str(0, 61, 0) = "fuerleistexparchiv"
 Str(0, 61, 1) = "`ID`"
 Str(0, 61, 2) = "`Name`"
 Str(0, 61, 3) = "`Datum`"
 Str(0, 61, 4) = "`Pat_id`"
 Str(0, 61, 5) = "`SchGr`"
 Str(0, 61, 6) = "`archiviert`"
 Str(0, 61, 7) = "`Protokoll`"
 Str(0, 61, 8) = "`ID`"
 Str(0, 61, 9) = "`PrimaryKey`"
 Str(0, 61, 10) = "`ID`"
 ArtZ(0, 61) = 7
 ArtZ(1, 61) = 3
 Str(1, 61, 0) = "CREATE TABLE `fuerleistexparchiv` ("
 Str(1, 61, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 61, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 61, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen'"
 Str(1, 61, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 61, 5) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 61, 6) = " `archiviert` datetime DEFAULT NULL COMMENT 'Datum der Übertragung in Tabelle Leistungsexport'"
 Str(1, 61, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 61, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 61, 9) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 61, 10) = "  KEY `ID` (`Pat_id`)"
 Str(1, 61, 11) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr61

Function FüllStr62()
 Str(0, 62, 0) = "hausaerzte"
 Str(0, 62, 1) = "`ID`"
 Str(0, 62, 2) = "`Überschrift`"
 Str(0, 62, 3) = "`Name`"
 Str(0, 62, 4) = "`Vorname`"
 Str(0, 62, 5) = "`Nachname`"
 Str(0, 62, 6) = "`Anschrift`"
 Str(0, 62, 7) = "`KVNr`"
 Str(0, 62, 8) = "`Telefon`"
 Str(0, 62, 9) = "`Telefax`"
 Str(0, 62, 10) = "`E_Mail`"
 Str(0, 62, 11) = "`Zulassungsgebiet`"
 Str(0, 62, 12) = "`Arzttyp`"
 Str(0, 62, 13) = "`Gemeinschaftspraxis mit`"
 Str(0, 62, 14) = "`Schwerpunkt`"
 Str(0, 62, 15) = "`Zusatzbezeichnung`"
 Str(0, 62, 16) = "`Bemerkung`"
 Str(0, 62, 17) = "`Beme`"
 Str(0, 62, 18) = "`Sprechstunden`"
 Str(0, 62, 19) = "`von _ bis`"
 Str(0, 62, 20) = "`Internetadressen`"
 Str(0, 62, 21) = "`Behandlung in Fremdsprachen`"
 Str(0, 62, 22) = "`Rollstuhlgerechte Praxis`"
 Str(0, 62, 23) = "`Verkehrsmittel`"
 Str(0, 62, 24) = "`Linie`"
 Str(0, 62, 25) = "`Haltestelle Parkplätze`"
 Str(0, 62, 26) = "`Wegbeschreibung`"
 Str(0, 62, 27) = "`Entfernung zur Praxis`"
 Str(0, 62, 28) = "`Zahl`"
 Str(0, 62, 29) = "`nichtmehr`"
 Str(0, 62, 30) = "`Titel`"
 Str(0, 62, 31) = "`Geschlecht`"
 Str(0, 62, 32) = "`Straße`"
 Str(0, 62, 33) = "`PLZ`"
 Str(0, 62, 34) = "`Ort`"
 Str(0, 62, 35) = "`DMPT2`"
 Str(0, 62, 36) = "`DMPT1`"
 Str(0, 62, 37) = "`gelöscht`"
 Str(0, 62, 38) = "`ID`"
 Str(0, 62, 39) = "`PrimaryKey`"
 Str(0, 62, 40) = "`Auswahl`"
 Str(0, 62, 41) = "`KVNr`"
 ArtZ(0, 62) = 37
 ArtZ(1, 62) = 4
 Str(1, 62, 0) = "CREATE TABLE `hausaerzte` ("
 Str(1, 62, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 62, 2) = " `Überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '""L"" = Liebe(r), ""H"" = Hallo'"
 Str(1, 62, 3) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 4) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 5) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 6) = " `Anschrift` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 7) = " `KVNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 8) = " `Telefon` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 9) = " `Telefax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 10) = " `E_Mail` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 11) = " `Zulassungsgebiet` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 12) = " `Arzttyp` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 13) = " `Gemeinschaftspraxis mit` longtext COLLATE latin1_german2_ci"
 Str(1, 62, 14) = " `Schwerpunkt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 15) = " `Zusatzbezeichnung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 16) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 62, 17) = " `Beme` longtext COLLATE latin1_german2_ci"
 Str(1, 62, 18) = " `Sprechstunden` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 19) = " `von _ bis` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 20) = " `Internetadressen` longtext COLLATE latin1_german2_ci"
 Str(1, 62, 21) = " `Behandlung in Fremdsprachen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 22) = " `Rollstuhlgerechte Praxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 23) = " `Verkehrsmittel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 24) = " `Linie` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 25) = " `Haltestelle Parkplätze` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 26) = " `Wegbeschreibung` longtext COLLATE latin1_german2_ci"
 Str(1, 62, 27) = " `Entfernung zur Praxis` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 28) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 62, 29) = " `nichtmehr` bit(1) DEFAULT NULL COMMENT 'Arzt nicht mehr im Verzeichnis'"
 Str(1, 62, 30) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 31) = " `Geschlecht` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 32) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 33) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 34) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 62, 35) = " `DMPT2` smallint(6) DEFAULT NULL"
 Str(1, 62, 36) = " `DMPT1` bit(1) DEFAULT NULL"
 Str(1, 62, 37) = " `gelöscht` bit(1) DEFAULT NULL"
 Str(1, 62, 38) = "  PRIMARY KEY (`ID`)"
 Str(1, 62, 39) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 62, 40) = "  KEY `Auswahl` (`KVNr`,`Name`)"
 Str(1, 62, 41) = "  KEY `KVNr` (`KVNr`,`Nachname`,`Vorname`)"
 Str(1, 62, 42) = " ENGINE=InnoDB AUTO_INCREMENT=22275 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr62

Function FüllStr63()
 Str(0, 63, 0) = "inl"
 Str(0, 63, 1) = "`id`"
 Str(0, 63, 2) = "`StByte`"
 Str(0, 63, 3) = "`breite`"
 Str(0, 63, 4) = "`kennung`"
 Str(0, 63, 5) = "`inhalt`"
 Str(0, 63, 6) = "`id`"
 Str(0, 63, 7) = "`stbyte`"
 Str(0, 63, 8) = "`kennung`"
 ArtZ(0, 63) = 5
 ArtZ(1, 63) = 3
 Str(1, 63, 0) = "CREATE TABLE `inl` ("
 Str(1, 63, 1) = " `id` int(15) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 63, 2) = " `StByte` int(10) unsigned NOT NULL"
 Str(1, 63, 3) = " `breite` char(3) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 63, 4) = " `kennung` char(4) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 63, 5) = " `inhalt` varchar(10000) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 63, 6) = "  PRIMARY KEY (`id`)"
 Str(1, 63, 7) = "  KEY `stbyte` (`StByte`) USING BTREE"
 Str(1, 63, 8) = "  KEY `kennung` (`kennung`,`inhalt`(6)) USING BTREE"
 Str(1, 63, 9) = " ENGINE=MyISAM AUTO_INCREMENT=51933093 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr63

Function FüllStr64()
 Str(0, 64, 0) = "kassenliste"
 Str(0, 64, 1) = "`ID`"
 Str(0, 64, 2) = "`VK`"
 Str(0, 64, 3) = "`IK`"
 Str(0, 64, 4) = "`Name`"
 Str(0, 64, 5) = "`Kateg`"
 Str(0, 64, 6) = "`AnzahlIK`"
 Str(0, 64, 7) = "`AnzahlKTUG`"
 Str(0, 64, 8) = "`GültigVon`"
 Str(0, 64, 9) = "`GültigBis`"
 Str(0, 64, 10) = "`GO`"
 Str(0, 64, 11) = "`Kurzname`"
 Str(0, 64, 12) = "`rName`"
 Str(0, 64, 13) = "`ID`"
 Str(0, 64, 14) = "`PrimaryKey`"
 Str(0, 64, 15) = "`VK`"
 Str(0, 64, 16) = "`IK`"
 Str(0, 64, 17) = "`VKIK`"
 ArtZ(0, 64) = 12
 ArtZ(1, 64) = 5
 Str(1, 64, 0) = "CREATE TABLE `kassenliste` ("
 Str(1, 64, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 64, 2) = " `VK` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 3) = " `IK` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 4) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 5) = " `Kateg` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kategorie'"
 Str(1, 64, 6) = " `AnzahlIK` int(4) unsigned DEFAULT NULL"
 Str(1, 64, 7) = " `AnzahlKTUG` int(4) unsigned DEFAULT NULL"
 Str(1, 64, 8) = " `GültigVon` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 9) = " `GültigBis` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 10) = " `GO` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 11) = " `Kurzname` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 12) = " `rName` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kurzname, falls nicht verfügbar: Name'"
 Str(1, 64, 13) = "  PRIMARY KEY (`ID`)"
 Str(1, 64, 14) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 64, 15) = "  KEY `VK` (`VK`)"
 Str(1, 64, 16) = "  KEY `IK` (`IK`)"
 Str(1, 64, 17) = "  KEY `VKIK` (`VK`,`IK`)"
 Str(1, 64, 18) = " ENGINE=InnoDB AUTO_INCREMENT=5327 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr64

Function FüllStr65()
 Str(0, 65, 0) = "kheinweis"
 Str(0, 65, 1) = "`FID`"
 Str(0, 65, 2) = "`Pat_ID`"
 Str(0, 65, 3) = "`ZeitPunkt`"
 Str(0, 65, 4) = "`Ziel`"
 Str(0, 65, 5) = "`Diagnose`"
 Str(0, 65, 6) = "`absPos`"
 Str(0, 65, 7) = "`AktZeit`"
 Str(0, 65, 8) = "`StByte`"
 Str(0, 65, 9) = "`Auswahl`"
 Str(0, 65, 10) = "`FälleKHEinweis`"
 Str(0, 65, 11) = "`FID`"
 Str(0, 65, 12) = "`NamenKHEinweis`"
 Str(0, 65, 13) = "`FälleKHEinweis_AccRel`"
 Str(0, 65, 14) = "`NamenKHEinweis_AccRel`"
 ArtZ(0, 65) = 8
 ArtZ(1, 65) = 4
 ArtZ(2, 65) = 2
 Str(1, 65, 0) = "CREATE TABLE `kheinweis` ("
 Str(1, 65, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 65, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 65, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 65, 4) = " `Ziel` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6291'"
 Str(1, 65, 5) = " `Diagnose` longtext COLLATE latin1_german2_ci COMMENT '6230'"
 Str(1, 65, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 65, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 65, 8) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 65, 9) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Ziel`)"
 Str(1, 65, 10) = "  KEY `FälleKHEinweis` (`FID`)"
 Str(1, 65, 11) = "  KEY `FID` (`FID`)"
 Str(1, 65, 12) = "  KEY `NamenKHEinweis` (`Pat_ID`)"
 Str(1, 65, 13) = "  CONSTRAINT `FälleKHEinweis_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON UPDATE CASCADE"
 Str(1, 65, 14) = "  CONSTRAINT `NamenKHEinweis_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 65, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr65

Function FüllStr66()
 Str(0, 66, 0) = "kvnrue"
 Str(0, 66, 1) = "`lfdnr`"
 Str(0, 66, 2) = "`Pat_ID`"
 Str(0, 66, 3) = "`KVNr`"
 Str(0, 66, 4) = "`absPos`"
 Str(0, 66, 5) = "`AktZeit`"
 Str(0, 66, 6) = "`StByte`"
 Str(0, 66, 7) = "`lfdnr`"
 Str(0, 66, 8) = "`PrimaryKey`"
 Str(0, 66, 9) = "`zuord`"
 ArtZ(0, 66) = 6
 ArtZ(1, 66) = 3
 Str(1, 66, 0) = "CREATE TABLE `kvnrue` ("
 Str(1, 66, 1) = " `lfdnr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 66, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 66, 3) = " `KVNr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 66, 4) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 66, 5) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Zeit der Aktualisuerung aus der BDT-Datei'"
 Str(1, 66, 6) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 66, 7) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 66, 8) = "  UNIQUE KEY `PrimaryKey` (`lfdnr`)"
 Str(1, 66, 9) = "  UNIQUE KEY `zuord` (`Pat_ID`,`KVNr`)"
 Str(1, 66, 10) = " ENGINE=InnoDB AUTO_INCREMENT=3993 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr66

Function FüllStr67()
 Str(0, 67, 0) = "labor1"
 Str(1, 67, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `labor1` AS select `n`.`Pat_ID` AS `Pat_ID`,cast(`n`.`ZeitPunkt` as date) AS `ZeitPunkt`,`n`.`FertigStGrad` AS `FertigStGrad`,`n`.`Abkü` AS `Abkü`,`l`.`Langtext` AS `Langtext`,`n`.`Wert` AS `Wert`,`n`.`Einheit` AS `Einheit`,`k`.`Kommentar` AS `Kommentar`,_utf8'' AS `nb` from ((`laborneu` `n` left join `laborlangtext` `l` on((`l`.`LangtextVW` = `n`.`LangtextVW`))) left join `laborkommentar` `k` on((`k`.`KommentarVW` = `n`.`KommentarVW`))) where ((`n`.`Wert` <> '') and (`n`.`Wert` is not null))"
End Function ' FüllStr67

Function FüllStr68()
 Str(0, 68, 0) = "labor2"
 Str(1, 68, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `labor2` AS select `u`.`Pat_id` AS `Pat_ID`,cast(`u`.`Eingang` as date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,`w`.`Abkü` AS `Abkü`,`w`.`Langname` AS `Langtext`,`w`.`Wert` AS `Wert`,`w`.`Einheit` AS `Einheit`,`w`.`Kommentar` AS `Kommentar`,`w`.`Normbereich` AS `NB` from (`laborxus` `u` left join `laborxwert` `w` on((`u`.`RefNr` = `w`.`RefNr`))) where ((not(exists(select 1 AS `Not_used` from `laborneu` where ((`laborneu`.`Pat_ID` = `u`.`Pat_id`) and (`laborneu`.`Abkü` = `w`.`Abkü`) and (`laborneu`.`Wert` = `w`.`Wert`) and (`laborneu`.`ZeitPunkt` > (`u`.`Eingang` - interval 3 day)) and (`laborneu`.`ZeitPunkt` < (`u`.`Eingang` + interval 6 day)))))) and (`w`.`Wert` <> '') and (`w`.`Wert` is not null))"
End Function ' FüllStr68

Function FüllStr69()
 Str(0, 69, 0) = "laborgruppen"
 Str(0, 69, 1) = "`Laborgruppe`"
 Str(0, 69, 2) = "`Erklärung`"
 Str(0, 69, 3) = "`Laborgruppe`"
 Str(0, 69, 4) = "`Gruppe`"
 ArtZ(0, 69) = 2
 ArtZ(1, 69) = 2
 Str(1, 69, 0) = "CREATE TABLE `laborgruppen` ("
 Str(1, 69, 1) = " `Laborgruppe` int(10) NOT NULL DEFAULT '0'"
 Str(1, 69, 2) = " `Erklärung` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 69, 3) = "  PRIMARY KEY (`Laborgruppe`)"
 Str(1, 69, 4) = "  UNIQUE KEY `Gruppe` (`Laborgruppe`)"
 Str(1, 69, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr69

Function FüllStr70()
 Str(0, 70, 0) = "laborkommentar"
 Str(0, 70, 1) = "`KommentarVW`"
 Str(0, 70, 2) = "`Kommentar`"
 Str(0, 70, 3) = "`KommentarVW`"
 Str(0, 70, 4) = "`KommentarVW`"
 Str(0, 70, 5) = "`Kommentar`"
 ArtZ(0, 70) = 2
 ArtZ(1, 70) = 3
 Str(1, 70, 0) = "CREATE TABLE `laborkommentar` ("
 Str(1, 70, 1) = " `KommentarVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 70, 2) = " `Kommentar` longtext COLLATE latin1_german2_ci"
 Str(1, 70, 3) = "  PRIMARY KEY (`KommentarVW`)"
 Str(1, 70, 4) = "  UNIQUE KEY `KommentarVW` (`KommentarVW`)"
 Str(1, 70, 5) = "  KEY `Kommentar` (`Kommentar`(255))"
 Str(1, 70, 6) = " ENGINE=InnoDB AUTO_INCREMENT=27487 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr70

Function FüllStr71()
 Str(0, 71, 0) = "laborlangtext"
 Str(0, 71, 1) = "`LangtextVW`"
 Str(0, 71, 2) = "`Langtext`"
 Str(0, 71, 3) = "`LangtextVW`"
 Str(0, 71, 4) = "`LangtextVW`"
 Str(0, 71, 5) = "`Langtext`"
 ArtZ(0, 71) = 2
 ArtZ(1, 71) = 3
 Str(1, 71, 0) = "CREATE TABLE `laborlangtext` ("
 Str(1, 71, 1) = " `LangtextVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 71, 2) = " `Langtext` longtext COLLATE latin1_german2_ci"
 Str(1, 71, 3) = "  PRIMARY KEY (`LangtextVW`)"
 Str(1, 71, 4) = "  UNIQUE KEY `LangtextVW` (`LangtextVW`)"
 Str(1, 71, 5) = "  KEY `Langtext` (`Langtext`(255))"
 Str(1, 71, 6) = " ENGINE=InnoDB AUTO_INCREMENT=931 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr71

Function FüllStr72()
 Str(0, 72, 0) = "laborneu"
 Str(0, 72, 1) = "`FID`"
 Str(0, 72, 2) = "`Pat_ID`"
 Str(0, 72, 3) = "`ZeitPunkt`"
 Str(0, 72, 4) = "`FertigStGrad`"
 Str(0, 72, 5) = "`Abkü`"
 Str(0, 72, 6) = "`LangtextVW`"
 Str(0, 72, 7) = "`Wert`"
 Str(0, 72, 8) = "`Einheit`"
 Str(0, 72, 9) = "`KommentarVW`"
 Str(0, 72, 10) = "`AbsPos`"
 Str(0, 72, 11) = "`AktZeit`"
 Str(0, 72, 12) = "`Refnr`"
 Str(0, 72, 13) = "`StByte`"
 Str(0, 72, 14) = "`AbküWert`"
 Str(0, 72, 15) = "`Auswahl`"
 Str(0, 72, 16) = "`FälleLaborNeu`"
 Str(0, 72, 17) = "`LaborKommentarLaborNeu`"
 Str(0, 72, 18) = "`LaborLangtextLaborNeu`"
 Str(0, 72, 19) = "`LaborParameterLaborNeu`"
 Str(0, 72, 20) = "`NamenLaborNeu`"
 Str(0, 72, 21) = "`Prüf`"
 Str(0, 72, 22) = "`FälleLaborNeu_AccRel`"
 Str(0, 72, 23) = "`LaborKommentarLaborNeu_AccRel`"
 Str(0, 72, 24) = "`LaborLangtextLaborNeu_AccRel`"
 Str(0, 72, 25) = "`LaborParameterLaborNeu_AccRel`"
 Str(0, 72, 26) = "`NamenLaborNeu_AccRel`"
 ArtZ(0, 72) = 13
 ArtZ(1, 72) = 8
 ArtZ(2, 72) = 5
 Str(1, 72, 0) = "CREATE TABLE `laborneu` ("
 Str(1, 72, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 72, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 72, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 72, 4) = " `FertigStGrad` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8401'"
 Str(1, 72, 5) = " `Abkü` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410'"
 Str(1, 72, 6) = " `LangtextVW` int(10) DEFAULT NULL COMMENT '8411'"
 Str(1, 72, 7) = " `Wert` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8420'"
 Str(1, 72, 8) = " `Einheit` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421'"
 Str(1, 72, 9) = " `KommentarVW` int(10) DEFAULT NULL COMMENT '8480'"
 Str(1, 72, 10) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 72, 11) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 72, 12) = " `Refnr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborXUS'"
 Str(1, 72, 13) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 72, 14) = "  KEY `AbküWert` (`Abkü`,`Wert`)"
 Str(1, 72, 15) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`FertigStGrad`,`Abkü`)"
 Str(1, 72, 16) = "  KEY `FälleLaborNeu` (`FID`)"
 Str(1, 72, 17) = "  KEY `LaborKommentarLaborNeu` (`KommentarVW`)"
 Str(1, 72, 18) = "  KEY `LaborLangtextLaborNeu` (`LangtextVW`)"
 Str(1, 72, 19) = "  KEY `LaborParameterLaborNeu` (`Abkü`,`Einheit`)"
 Str(1, 72, 20) = "  KEY `NamenLaborNeu` (`Pat_ID`)"
 Str(1, 72, 21) = "  KEY `Prüf` (`Pat_ID`,`Abkü`,`Wert`)"
 Str(1, 72, 22) = "  CONSTRAINT `FälleLaborNeu_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 72, 23) = "  CONSTRAINT `LaborKommentarLaborNeu_AccRel` FOREIGN KEY (`KommentarVW`) REFERENCES `laborkommentar` (`KommentarVW`)"
 Str(1, 72, 24) = "  CONSTRAINT `LaborLangtextLaborNeu_AccRel` FOREIGN KEY (`LangtextVW`) REFERENCES `laborlangtext` (`LangtextVW`) ON UPDATE CASCADE"
 Str(1, 72, 25) = "  CONSTRAINT `LaborParameterLaborNeu_AccRel` FOREIGN KEY (`Abkü`, `Einheit`) REFERENCES `laborparameter` (`Abkü`, `Einheit`)"
 Str(1, 72, 26) = "  CONSTRAINT `NamenLaborNeu_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 72, 27) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Function ' FüllStr72

Function FüllStr73()
 Str(0, 73, 0) = "laborparameter"
 Str(0, 73, 1) = "`Abkü`"
 Str(0, 73, 2) = "`Labor`"
 Str(0, 73, 3) = "`Langtext`"
 Str(0, 73, 4) = "`Einheit`"
 Str(0, 73, 5) = "`Gruppe`"
 Str(0, 73, 6) = "`Reihe`"
 Str(0, 73, 7) = "`uNm`"
 Str(0, 73, 8) = "`oNm`"
 Str(0, 73, 9) = "`uNw`"
 Str(0, 73, 10) = "`oNw`"
 Str(0, 73, 11) = "`AktZeit`"
 Str(0, 73, 12) = "`StByte`"
 Str(0, 73, 13) = "`Abkü`"
 Str(0, 73, 14) = "`Fehlende`"
 Str(0, 73, 15) = "`LaborParameterAbkü`"
 Str(0, 73, 16) = "`Reihe`"
 Str(0, 73, 17) = "`LaborgruppenLaborParameter_AccRel`"
 ArtZ(0, 73) = 12
 ArtZ(1, 73) = 4
 ArtZ(2, 73) = 1
 Str(1, 73, 0) = "CREATE TABLE `laborparameter` ("
 Str(1, 73, 1) = " `Abkü` varchar(70) COLLATE latin1_german2_ci NOT NULL COMMENT '8410(1)'"
 Str(1, 73, 2) = " `Labor` varchar(40) COLLATE latin1_german2_ci NOT NULL COMMENT '8410(2)'"
 Str(1, 73, 3) = " `Langtext` varchar(60) COLLATE latin1_german2_ci NOT NULL COMMENT '8411'"
 Str(1, 73, 4) = " `Einheit` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421'"
 Str(1, 73, 5) = " `Gruppe` int(10) DEFAULT NULL COMMENT 'Ordnungsgruppe'"
 Str(1, 73, 6) = " `Reihe` int(10) DEFAULT NULL COMMENT 'Reihenfolge innerhalb der Gruppe'"
 Str(1, 73, 7) = " `uNm` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'unterer Normwert männlich'"
 Str(1, 73, 8) = " `oNm` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'oberer Normwert männlich'"
 Str(1, 73, 9) = " `uNw` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'unterer Normwert weiblich'"
 Str(1, 73, 10) = " `oNw` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'oberer Normwert weiblich'"
 Str(1, 73, 11) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 73, 12) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 73, 13) = "  UNIQUE KEY `Abkü` (`Abkü`,`Einheit`)"
 Str(1, 73, 14) = "  KEY `Fehlende` (`Gruppe`,`AktZeit`)"
 Str(1, 73, 15) = "  KEY `LaborParameterAbkü` (`Abkü`)"
 Str(1, 73, 16) = "  KEY `Reihe` (`Gruppe`,`Abkü`)"
 Str(1, 73, 17) = "  CONSTRAINT `LaborgruppenLaborParameter_AccRel` FOREIGN KEY (`Gruppe`) REFERENCES `laborgruppen` (`Laborgruppe`)"
 Str(1, 73, 18) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Function ' FüllStr73

Function FüllStr74()
 Str(0, 74, 0) = "laborxbakt"
 Str(0, 74, 1) = "`RefNr`"
 Str(0, 74, 2) = "`Verf`"
 Str(0, 74, 3) = "`KuQu`"
 Str(0, 74, 4) = "`Quelle`"
 Str(0, 74, 5) = "`QSpez`"
 Str(0, 74, 6) = "`AbnDat`"
 Str(0, 74, 7) = "`Kommentar`"
 Str(0, 74, 8) = "`Erklärung`"
 Str(0, 74, 9) = "`Keimzahl`"
 Str(0, 74, 10) = "`LaborXBaktUS`"
 Str(0, 74, 11) = "`LaborXUSLaborXBakt`"
 Str(0, 74, 12) = "`RefNr`"
 Str(0, 74, 13) = "`LaborXBaktUS_AccRel`"
 Str(0, 74, 14) = "`LaborXUSLaborXBakt_AccRel`"
 ArtZ(0, 74) = 9
 ArtZ(1, 74) = 3
 ArtZ(2, 74) = 2
 Str(1, 74, 0) = "CREATE TABLE `laborxbakt` ("
 Str(1, 74, 1) = " `RefNr` int(10) DEFAULT NULL"
 Str(1, 74, 2) = " `Verf` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 3) = " `KuQu` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8428 Probenmaterial-Ident (Turbomed)'"
 Str(1, 74, 4) = " `Quelle` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8430 Probenmaterial-Bezeichnung (Turbomed)'"
 Str(1, 74, 5) = " `QSpez` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8431 Probenmaterial-Spezifikation (Turbomed)'"
 Str(1, 74, 6) = " `AbnDat` datetime DEFAULT NULL COMMENT '8432 Abnahmedatum (Turbomed)'"
 Str(1, 74, 7) = " `Kommentar` longtext COLLATE latin1_german2_ci COMMENT '8480 Ergebnistest (Turbomed)'"
 Str(1, 74, 8) = " `Erklärung` longtext COLLATE latin1_german2_ci"
 Str(1, 74, 9) = " `Keimzahl` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 10) = "  KEY `LaborXBaktUS` (`RefNr`)"
 Str(1, 74, 11) = "  KEY `LaborXUSLaborXBakt` (`RefNr`)"
 Str(1, 74, 12) = "  KEY `RefNr` (`RefNr`)"
 Str(1, 74, 13) = "  CONSTRAINT `LaborXBaktUS_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 74, 14) = "  CONSTRAINT `LaborXUSLaborXBakt_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 74, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr74

Function FüllStr75()
 Str(0, 75, 0) = "laborxeingel"
 Str(0, 75, 1) = "`DatID`"
 Str(0, 75, 2) = "`Pfad`"
 Str(0, 75, 3) = "`Name`"
 Str(0, 75, 4) = "`Zp`"
 Str(0, 75, 5) = "`fertig`"
 Str(0, 75, 6) = "`DatID`"
 Str(0, 75, 7) = "`DatID`"
 Str(0, 75, 8) = "`NamePfad`"
 ArtZ(0, 75) = 5
 ArtZ(1, 75) = 3
 Str(1, 75, 0) = "CREATE TABLE `laborxeingel` ("
 Str(1, 75, 1) = " `DatID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf LaborEingelesen'"
 Str(1, 75, 2) = " `Pfad` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Pfadname'"
 Str(1, 75, 3) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name der eingelesenen Labordatei ohne Endung'"
 Str(1, 75, 4) = " `Zp` datetime DEFAULT NULL COMMENT 'Einlesezeitpunkt'"
 Str(1, 75, 5) = " `fertig` bit(1) DEFAULT NULL COMMENT 'ob Einlesen fertig'"
 Str(1, 75, 6) = "  PRIMARY KEY (`DatID`)"
 Str(1, 75, 7) = "  UNIQUE KEY `DatID` (`DatID`)"
 Str(1, 75, 8) = "  KEY `NamePfad` (`Name`,`Pfad`)"
 Str(1, 75, 9) = " ENGINE=InnoDB AUTO_INCREMENT=269 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr75

Function FüllStr76()
 Str(0, 76, 0) = "laborxleist"
 Str(0, 76, 1) = "`RefNr`"
 Str(0, 76, 2) = "`Abkü`"
 Str(0, 76, 3) = "`Verf`"
 Str(0, 76, 4) = "`EBM`"
 Str(0, 76, 5) = "`goä`"
 Str(0, 76, 6) = "`Anzahl`"
 Str(0, 76, 7) = "`LaborXLeistUS`"
 Str(0, 76, 8) = "`RefNr`"
 Str(0, 76, 9) = "`LaborXLeistUS_AccRel`"
 ArtZ(0, 76) = 6
 ArtZ(1, 76) = 2
 ArtZ(2, 76) = 1
 Str(1, 76, 0) = "CREATE TABLE `laborxleist` ("
 Str(1, 76, 1) = " `RefNr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborUS'"
 Str(1, 76, 2) = " `Abkü` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410 Test-Ident (Turbomed)'"
 Str(1, 76, 3) = " `Verf` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8434'"
 Str(1, 76, 4) = " `EBM` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5001 GNR (Turbomed)'"
 Str(1, 76, 5) = " `goä` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8406'"
 Str(1, 76, 6) = " `Anzahl` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5005'"
 Str(1, 76, 7) = "  KEY `LaborXLeistUS` (`RefNr`)"
 Str(1, 76, 8) = "  KEY `RefNr` (`RefNr`)"
 Str(1, 76, 9) = "  CONSTRAINT `LaborXLeistUS_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 76, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr76

Function FüllStr77()
 Str(0, 77, 0) = "laborxls"
 Str(0, 77, 1) = "`id`"
 Str(0, 77, 2) = "`patient`"
 Str(0, 77, 3) = "`fehlerart`"
 Str(0, 77, 4) = "`id`"
 ArtZ(0, 77) = 3
 ArtZ(1, 77) = 1
 Str(1, 77, 0) = "CREATE TABLE `laborxls` ("
 Str(1, 77, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 77, 2) = " `patient` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 77, 3) = " `fehlerart` varchar(3000) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 77, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 77, 5) = " ENGINE=MyISAM AUTO_INCREMENT=30 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr77

Function FüllStr78()
 Str(0, 78, 0) = "laborxsaetze"
 Str(0, 78, 1) = "`SatzID`"
 Str(0, 78, 2) = "`DatID`"
 Str(0, 78, 3) = "`Satzart`"
 Str(0, 78, 4) = "`Satzlänge`"
 Str(0, 78, 5) = "`SatzlängeSchluss`"
 Str(0, 78, 6) = "`VersionSatzb`"
 Str(0, 78, 7) = "`Arztnr`"
 Str(0, 78, 8) = "`Arztname`"
 Str(0, 78, 9) = "`StraßePraxis`"
 Str(0, 78, 10) = "`Arzt`"
 Str(0, 78, 11) = "`LANR`"
 Str(0, 78, 12) = "`PLZPraxis`"
 Str(0, 78, 13) = "`OrtPraxis`"
 Str(0, 78, 14) = "`Labor`"
 Str(0, 78, 15) = "`StraßeLabor`"
 Str(0, 78, 16) = "`PLZLabor`"
 Str(0, 78, 17) = "`OrtLabor`"
 Str(0, 78, 18) = "`KBVPrüfnr`"
 Str(0, 78, 19) = "`Zeichensatz`"
 Str(0, 78, 20) = "`Kundenarztnr`"
 Str(0, 78, 21) = "`Erstellungsdatum`"
 Str(0, 78, 22) = "`Gesamtlänge`"
 Str(0, 78, 23) = "`SatzID`"
 Str(0, 78, 24) = "`SatzID`"
 Str(0, 78, 25) = "`DatID`"
 Str(0, 78, 26) = "`Name`"
 ArtZ(0, 78) = 22
 ArtZ(1, 78) = 4
 Str(1, 78, 0) = "CREATE TABLE `laborxsaetze` ("
 Str(1, 78, 1) = " `SatzID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'zum Bezug für LaborUS'"
 Str(1, 78, 2) = " `DatID` int(10) DEFAULT NULL COMMENT 'Bezug zu LaborEingelesen'"
 Str(1, 78, 3) = " `Satzart` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000 Satzart (Turbomed)'"
 Str(1, 78, 4) = " `Satzlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed)'"
 Str(1, 78, 5) = " `SatzlängeSchluss` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000'"
 Str(1, 78, 6) = " `VersionSatzb` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9212 Version der Satzbeschreibung (Turbomed)'"
 Str(1, 78, 7) = " `Arztnr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '201 Arztnummer (Turbomed)'"
 Str(1, 78, 8) = " `Arztname` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '203 Arztname (Turbomed)'"
 Str(1, 78, 9) = " `StraßePraxis` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '205 Straße der Praxis (Turbomed)'"
 Str(1, 78, 10) = " `Arzt` varchar(45) COLLATE latin1_german2_ci NOT NULL COMMENT ' 211 Ausführender Arzt'"
 Str(1, 78, 11) = " `LANR` varchar(11) COLLATE latin1_german2_ci NOT NULL COMMENT ' 212 LANR'"
 Str(1, 78, 12) = " `PLZPraxis` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '215 PLZ der Praxis (Turbomed)'"
 Str(1, 78, 13) = " `OrtPraxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '216 Ort der Praxis (Turbomed)'"
 Str(1, 78, 14) = " `Labor` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8320 Labor'"
 Str(1, 78, 15) = " `StraßeLabor` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8321 Straße der Laboradresse (Turbomed)'"
 Str(1, 78, 16) = " `PLZLabor` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8322 PLZ der Laboradresse (Turbomed)'"
 Str(1, 78, 17) = " `OrtLabor` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8323 Ort der Laboradresse (Turbomed)'"
 Str(1, 78, 18) = " `KBVPrüfnr` varchar(16) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '101 KBV-Prüfnummer (Turbomed)'"
 Str(1, 78, 19) = " `Zeichensatz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9106 verwendeter Zeichensatz (Turbomed)'"
 Str(1, 78, 20) = " `Kundenarztnr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8312 Kundenarztnummer (Turbomed)'"
 Str(1, 78, 21) = " `Erstellungsdatum` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9103 Erstellungsdatum (Turbomed)'"
 Str(1, 78, 22) = " `Gesamtlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9202 Gesamtlänge des Datenpaketes (Turbomed)'"
 Str(1, 78, 23) = "  PRIMARY KEY (`SatzID`)"
 Str(1, 78, 24) = "  UNIQUE KEY `SatzID` (`SatzID`)"
 Str(1, 78, 25) = "  KEY `DatID` (`DatID`)"
 Str(1, 78, 26) = "  KEY `Name` (`PLZLabor`,`OrtLabor`)"
 Str(1, 78, 27) = " ENGINE=InnoDB AUTO_INCREMENT=685 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr78

Function FüllStr79()
 Str(0, 79, 0) = "laborxus"
 Str(0, 79, 1) = "`RefNr`"
 Str(0, 79, 2) = "`DatID`"
 Str(0, 79, 3) = "`SatzID`"
 Str(0, 79, 4) = "`Satzart`"
 Str(0, 79, 5) = "`Satzlänge`"
 Str(0, 79, 6) = "`Auftragsnummer`"
 Str(0, 79, 7) = "`Auftragsschlüssel`"
 Str(0, 79, 8) = "`Eingang`"
 Str(0, 79, 9) = "`Berichtsdatum`"
 Str(0, 79, 10) = "`Pat_id`"
 Str(0, 79, 11) = "`Nachname`"
 Str(0, 79, 12) = "`Vorname`"
 Str(0, 79, 13) = "`GebDat`"
 Str(0, 79, 14) = "`Titel`"
 Str(0, 79, 15) = "`NVorsatz`"
 Str(0, 79, 16) = "`BefArt`"
 Str(0, 79, 17) = "`Abrechnungstyp`"
 Str(0, 79, 18) = "`GebüOrd`"
 Str(0, 79, 19) = "`Patienteninformation`"
 Str(0, 79, 20) = "`Geschlecht`"
 Str(0, 79, 21) = "`AuftrHinw`"
 Str(0, 79, 22) = "`Pat_idUrsp`"
 Str(0, 79, 23) = "`Pat_idErwVNG`"
 Str(0, 79, 24) = "`Pat_idErwVN`"
 Str(0, 79, 25) = "`Pat_idErwG`"
 Str(0, 79, 26) = "`Pat_idErwGB`"
 Str(0, 79, 27) = "`Pat_idErwGL`"
 Str(0, 79, 28) = "`Pat_idLaborNeu`"
 Str(0, 79, 29) = "`ZeitpunktLaborneu`"
 Str(0, 79, 30) = "`ZdüP`"
 Str(0, 79, 31) = "`ZdiP`"
 Str(0, 79, 32) = "`LWerte`"
 Str(0, 79, 33) = "`verglichen`"
 Str(0, 79, 34) = "`AfN`"
 Str(0, 79, 35) = "`RefNr`"
 Str(0, 79, 36) = "`RefNr`"
 Str(0, 79, 37) = "`DatID`"
 Str(0, 79, 38) = "`LaborXEingelLaborXUS`"
 Str(0, 79, 39) = "`LaborXSätzeLaborXUS`"
 Str(0, 79, 40) = "`Name`"
 Str(0, 79, 41) = "`Pat_id`"
 Str(0, 79, 42) = "`SatzID`"
 Str(0, 79, 43) = "`LaborXEingelLaborXUS_AccRel`"
 Str(0, 79, 44) = "`LaborXSätzeLaborXUS_AccRel`"
 ArtZ(0, 79) = 34
 ArtZ(1, 79) = 8
 ArtZ(2, 79) = 2
 Str(1, 79, 0) = "CREATE TABLE `laborxus` ("
 Str(1, 79, 1) = " `RefNr` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf LaborWert'"
 Str(1, 79, 2) = " `DatID` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborEingelesen'"
 Str(1, 79, 3) = " `SatzID` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborXSätze'"
 Str(1, 79, 4) = " `Satzart` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000 Satzart (Turbomed)'"
 Str(1, 79, 5) = " `Satzlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed)'"
 Str(1, 79, 6) = " `Auftragsnummer` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8310 Anforderungsident (Turbomed)'"
 Str(1, 79, 7) = " `Auftragsschlüssel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8311 Anforderungsnr d Labors (Turbomed)'"
 Str(1, 79, 8) = " `Eingang` datetime DEFAULT NULL COMMENT '8301 Eingangsdatum in Datumsform'"
 Str(1, 79, 9) = " `Berichtsdatum` varchar(13) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8302 Berichtsdatum'"
 Str(1, 79, 10) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 79, 11) = " `Nachname` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 79, 12) = " `Vorname` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 79, 13) = " `GebDat` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3103'"
 Str(1, 79, 14) = " `Titel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3104'"
 Str(1, 79, 15) = " `NVorsatz` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3100'"
 Str(1, 79, 16) = " `BefArt` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8401 Befundart (Turbomed) / Fertigstellungsgrad (""E""=Endbefund, ""T"" = Teilbefund)'"
 Str(1, 79, 17) = " `Abrechnungstyp` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)'"
 Str(1, 79, 18) = " `GebüOrd` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8403 Gebührenordnung (Turbomed)'"
 Str(1, 79, 19) = " `Patienteninformation` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8405 Patienteninformation (Turbomed)'"
 Str(1, 79, 20) = " `Geschlecht` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8407 Geschlecht (Turbomed)'"
 Str(1, 79, 21) = " `AuftrHinw` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8490 Auftragsbezogene Hinweise (Turbomed)'"
 Str(1, 79, 22) = " `Pat_idUrsp` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor'"
 Str(1, 79, 23) = " `Pat_idErwVNG` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag'"
 Str(1, 79, 24) = " `Pat_idErwVN` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Vornamen und Nachnamen'"
 Str(1, 79, 25) = " `Pat_idErwG` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag'"
 Str(1, 79, 26) = " `Pat_idErwGB` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung'"
 Str(1, 79, 27) = " `Pat_idErwGL` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor'"
 Str(1, 79, 28) = " `Pat_idLaborNeu` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Pat_ids von in Laborneu zuordnbaren Patienten'"
 Str(1, 79, 29) = " `ZeitpunktLaborneu` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde'"
 Str(1, 79, 30) = " `ZdüP` smallint(6) DEFAULT NULL COMMENT 'Zahl der verglichenen Parameter'"
 Str(1, 79, 31) = " `ZdiP` int(10) DEFAULT NULL COMMENT 'Zahl der infragekommenden Patienten'"
 Str(1, 79, 32) = " `LWerte` longtext COLLATE latin1_german2_ci COMMENT 'Laborwerte, die zur Zuordnung geführt haben'"
 Str(1, 79, 33) = " `verglichen` datetime DEFAULT NULL COMMENT 'Datum, zu dem Datensatz zuletzt verglichen wurde'"
 Str(1, 79, 34) = " `AfN` smallint(6) DEFAULT NULL COMMENT 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu'"
 Str(1, 79, 35) = "  PRIMARY KEY (`RefNr`)"
 Str(1, 79, 36) = "  UNIQUE KEY `RefNr` (`RefNr`)"
 Str(1, 79, 37) = "  KEY `DatID` (`DatID`)"
 Str(1, 79, 38) = "  KEY `LaborXEingelLaborXUS` (`DatID`)"
 Str(1, 79, 39) = "  KEY `LaborXSätzeLaborXUS` (`SatzID`)"
 Str(1, 79, 40) = "  KEY `Name` (`Nachname`,`Vorname`)"
 Str(1, 79, 41) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 79, 42) = "  KEY `SatzID` (`SatzID`)"
 Str(1, 79, 43) = "  CONSTRAINT `LaborXEingelLaborXUS_AccRel` FOREIGN KEY (`DatID`) REFERENCES `laborxeingel` (`DatID`) ON UPDATE CASCADE"
 Str(1, 79, 44) = "  CONSTRAINT `LaborXSätzeLaborXUS_AccRel` FOREIGN KEY (`SatzID`) REFERENCES `laborxsaetze` (`SatzID`) ON UPDATE CASCADE"
 Str(1, 79, 45) = " ENGINE=InnoDB AUTO_INCREMENT=3619 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr79

Function FüllStr80()
 Str(0, 80, 0) = "laborxwert"
 Str(0, 80, 1) = "`RefNr`"
 Str(0, 80, 2) = "`Abkü`"
 Str(0, 80, 3) = "`Langname`"
 Str(0, 80, 4) = "`Quelle`"
 Str(0, 80, 5) = "`QSpez`"
 Str(0, 80, 6) = "`AbnDat`"
 Str(0, 80, 7) = "`Wert`"
 Str(0, 80, 8) = "`Einheit`"
 Str(0, 80, 9) = "`Grenzwerti`"
 Str(0, 80, 10) = "`Kommentar`"
 Str(0, 80, 11) = "`Teststatus`"
 Str(0, 80, 12) = "`Erklärung`"
 Str(0, 80, 13) = "`Normbereich`"
 Str(0, 80, 14) = "`NormU`"
 Str(0, 80, 15) = "`NormO`"
 Str(0, 80, 16) = "`AuftrHinw`"
 Str(0, 80, 17) = "`LaborXUSLaborXWert`"
 Str(0, 80, 18) = "`LaborXWertAbkü`"
 Str(0, 80, 19) = "`LaborXWertUS`"
 Str(0, 80, 20) = "`RefNr`"
 Str(0, 80, 21) = "`LaborParameterLaborXWert_AccRel`"
 Str(0, 80, 22) = "`LaborXUSLaborXWert_AccRel`"
 Str(0, 80, 23) = "`LaborXWertUS_AccRel`"
 ArtZ(0, 80) = 16
 ArtZ(1, 80) = 4
 ArtZ(2, 80) = 3
 Str(1, 80, 0) = "CREATE TABLE `laborxwert` ("
 Str(1, 80, 1) = " `RefNr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborUS'"
 Str(1, 80, 2) = " `Abkü` varchar(16) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410 Test-Ident  (Turbomed)'"
 Str(1, 80, 3) = " `Langname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8411 Testbezeichnung (Turbomed)'"
 Str(1, 80, 4) = " `Quelle` varchar(43) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8430 Probenmaterial-Bezeichnung (Turbomed)'"
 Str(1, 80, 5) = " `QSpez` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8431 Probenmaterial-Spezifikation (Turbomed)'"
 Str(1, 80, 6) = " `AbnDat` datetime DEFAULT NULL COMMENT '8432 Abnahmedatum (Turbomed)'"
 Str(1, 80, 7) = " `Wert` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8420 Ergebniswert (Turbomed)'"
 Str(1, 80, 8) = " `Einheit` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421 Einheit (Turbomed)'"
 Str(1, 80, 9) = " `Grenzwerti` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8422 Grenzwertindikator (Turbomed)'"
 Str(1, 80, 10) = " `Kommentar` varchar(1385) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8480 Ergebnistext (Turbomed)'"
 Str(1, 80, 11) = " `Teststatus` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8418 Teststatus (Turbomed)'"
 Str(1, 80, 12) = " `Erklärung` varchar(759) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8470 Testbezogene Hinweise (Turbomed)'"
 Str(1, 80, 13) = " `Normbereich` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8460 Normalwert-Text (Turbomed)'"
 Str(1, 80, 14) = " `NormU` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8461 Normuntergrenze'"
 Str(1, 80, 15) = " `NormO` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8461 Normobergrenze'"
 Str(1, 80, 16) = " `AuftrHinw` varchar(180) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8490 Auftragsbezogene Hinweise (Turbomed)'"
 Str(1, 80, 17) = "  KEY `LaborXUSLaborXWert` (`RefNr`)"
 Str(1, 80, 18) = "  KEY `LaborXWertAbkü` (`Abkü`,`Einheit`)"
 Str(1, 80, 19) = "  KEY `LaborXWertUS` (`RefNr`)"
 Str(1, 80, 20) = "  KEY `RefNr` (`RefNr`)"
 Str(1, 80, 21) = "  CONSTRAINT `LaborParameterLaborXWert_AccRel` FOREIGN KEY (`Abkü`, `Einheit`) REFERENCES `laborparameter` (`Abkü`, `Einheit`)"
 Str(1, 80, 22) = "  CONSTRAINT `LaborXUSLaborXWert_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 80, 23) = "  CONSTRAINT `LaborXWertUS_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 80, 24) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr80

Function FüllStr81()
 Str(0, 81, 0) = "lbanforderungen"
 Str(0, 81, 1) = "`FID`"
 Str(0, 81, 2) = "`Pat_ID`"
 Str(0, 81, 3) = "`ZeitPunkt`"
 Str(0, 81, 4) = "`AnfText`"
 Str(0, 81, 5) = "`absPos`"
 Str(0, 81, 6) = "`AktZeit`"
 Str(0, 81, 7) = "`StByte`"
 Str(0, 81, 8) = "`Auswahl`"
 Str(0, 81, 9) = "`FälleLbAnforderungen`"
 Str(0, 81, 10) = "`FID`"
 Str(0, 81, 11) = "`NamenLbAnforderungen`"
 Str(0, 81, 12) = "`FälleLbAnforderungen_AccRel`"
 Str(0, 81, 13) = "`NamenLbAnforderungen_AccRel`"
 ArtZ(0, 81) = 7
 ArtZ(1, 81) = 4
 ArtZ(2, 81) = 2
 Str(1, 81, 0) = "CREATE TABLE `lbanforderungen` ("
 Str(1, 81, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 81, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 81, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 81, 4) = " `AnfText` longtext COLLATE latin1_german2_ci COMMENT '6280'"
 Str(1, 81, 5) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 81, 6) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 81, 7) = " `StByte` int(11) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 81, 8) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`AnfText`(255))"
 Str(1, 81, 9) = "  KEY `FälleLbAnforderungen` (`FID`)"
 Str(1, 81, 10) = "  KEY `FID` (`FID`)"
 Str(1, 81, 11) = "  KEY `NamenLbAnforderungen` (`Pat_ID`)"
 Str(1, 81, 12) = "  CONSTRAINT `FälleLbAnforderungen_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 81, 13) = "  CONSTRAINT `NamenLbAnforderungen_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 81, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr81

Function FüllStr82()
 Str(0, 82, 0) = "leistungen"
 Str(0, 82, 1) = "`FID`"
 Str(0, 82, 2) = "`Pat_ID`"
 Str(0, 82, 3) = "`ZeitPunkt`"
 Str(0, 82, 4) = "`Leistung`"
 Str(0, 82, 5) = "`f5002`"
 Str(0, 82, 6) = "`f5005`"
 Str(0, 82, 7) = "`f5006`"
 Str(0, 82, 8) = "`f5009`"
 Str(0, 82, 9) = "`Med`"
 Str(0, 82, 10) = "`f5015`"
 Str(0, 82, 11) = "`f5016`"
 Str(0, 82, 12) = "`f5021`"
 Str(0, 82, 13) = "`f5026`"
 Str(0, 82, 14) = "`Faktor`"
 Str(0, 82, 15) = "`f5098`"
 Str(0, 82, 16) = "`LANR`"
 Str(0, 82, 17) = "`letzVorg`"
 Str(0, 82, 18) = "`Ausn`"
 Str(0, 82, 19) = "`Beme`"
 Str(0, 82, 20) = "`absPos`"
 Str(0, 82, 21) = "`AktZeit`"
 Str(0, 82, 22) = "`QS`"
 Str(0, 82, 23) = "`QT`"
 Str(0, 82, 24) = "`StByte`"
 Str(0, 82, 25) = "`Auswahl`"
 Str(0, 82, 26) = "`FälleLeistungen`"
 Str(0, 82, 27) = "`FID`"
 Str(0, 82, 28) = "`Leistung`"
 Str(0, 82, 29) = "`NamenLeistungen`"
 Str(0, 82, 30) = "`FälleLeistungen_AccRel`"
 Str(0, 82, 31) = "`NamenLeistungen_AccRel`"
 ArtZ(0, 82) = 24
 ArtZ(1, 82) = 5
 ArtZ(2, 82) = 2
 Str(1, 82, 0) = "CREATE TABLE `leistungen` ("
 Str(1, 82, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 82, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 82, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '5000 + 6201'"
 Str(1, 82, 4) = " `Leistung` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5001'"
 Str(1, 82, 5) = " `f5002` varchar(32) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5002'"
 Str(1, 82, 6) = " `f5005` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5005'"
 Str(1, 82, 7) = " `f5006` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5006'"
 Str(1, 82, 8) = " `f5009` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5009'"
 Str(1, 82, 9) = " `Med` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5010'"
 Str(1, 82, 10) = " `f5015` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5015'"
 Str(1, 82, 11) = " `f5016` varchar(28) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5016'"
 Str(1, 82, 12) = " `f5021` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5021'"
 Str(1, 82, 13) = " `f5026` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5026'"
 Str(1, 82, 14) = " `Faktor` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5062 Multiplikator für GOÄ-Rechnung'"
 Str(1, 82, 15) = " `f5098` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5098 0000000000'"
 Str(1, 82, 16) = " `LANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5099 LANR'"
 Str(1, 82, 17) = " `letzVorg` datetime DEFAULT NULL COMMENT '5101 letzter Vorgang'"
 Str(1, 82, 18) = " `Ausn` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3677 Ausnahme/Begründung für abweichendes Geschlecht'"
 Str(1, 82, 19) = " `Beme` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '         Bemerkung'"
 Str(1, 82, 20) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 82, 21) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 82, 22) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 82, 23) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 82, 24) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 82, 25) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Leistung`)"
 Str(1, 82, 26) = "  KEY `FälleLeistungen` (`FID`)"
 Str(1, 82, 27) = "  KEY `FID` (`FID`)"
 Str(1, 82, 28) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 82, 29) = "  KEY `NamenLeistungen` (`Pat_ID`)"
 Str(1, 82, 30) = "  CONSTRAINT `FälleLeistungen_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 82, 31) = "  CONSTRAINT `NamenLeistungen_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 82, 32) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr82

Function FüllStr83()
 Str(0, 83, 0) = "leistungen exportiert"
 Str(0, 83, 1) = "`ID`"
 Str(0, 83, 2) = "`Datum`"
 Str(0, 83, 3) = "`Pat_id`"
 Str(0, 83, 4) = "`SchGr`"
 Str(0, 83, 5) = "`Leistung`"
 Str(0, 83, 6) = "`übertragen`"
 Str(0, 83, 7) = "`ID`"
 Str(0, 83, 8) = "`PrimaryKey`"
 Str(0, 83, 9) = "`ID`"
 ArtZ(0, 83) = 6
 ArtZ(1, 83) = 3
 Str(1, 83, 0) = "CREATE TABLE `leistungen exportiert` ("
 Str(1, 83, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Reihenfolge'"
 Str(1, 83, 2) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum'"
 Str(1, 83, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!Pat_id'"
 Str(1, 83, 4) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 83, 5) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ziffer aus EBM / GOÄ'"
 Str(1, 83, 6) = " `übertragen` datetime DEFAULT NULL COMMENT '"""", übertragen'"
 Str(1, 83, 7) = "  PRIMARY KEY (`ID`)"
 Str(1, 83, 8) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 83, 9) = "  KEY `ID` (`Pat_id`)"
 Str(1, 83, 10) = " ENGINE=InnoDB AUTO_INCREMENT=3970 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr83

Function FüllStr84()
 Str(0, 84, 0) = "leistungsexport"
 Str(0, 84, 1) = "`Prim`"
 Str(0, 84, 2) = "`PatID`"
 Str(0, 84, 3) = "`Datum`"
 Str(0, 84, 4) = "`UZeit`"
 Str(0, 84, 5) = "`Leistung`"
 Str(0, 84, 6) = "`SchGr`"
 Str(0, 84, 7) = "`Status`"
 Str(0, 84, 8) = "`Prim`"
 Str(0, 84, 9) = "`PatID`"
 ArtZ(0, 84) = 7
 ArtZ(1, 84) = 2
 Str(1, 84, 0) = "CREATE TABLE `leistungsexport` ("
 Str(1, 84, 1) = " `Prim` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Reihenfolge'"
 Str(1, 84, 2) = " `PatID` int(10) DEFAULT NULL COMMENT '-> Namen!Pat_id'"
 Str(1, 84, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum'"
 Str(1, 84, 4) = " `UZeit` datetime DEFAULT NULL COMMENT 'Leistungszeit, falls gewünscht'"
 Str(1, 84, 5) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ziffer aus EBM / GOÄ'"
 Str(1, 84, 6) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 84, 7) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '"""", übertragen'"
 Str(1, 84, 8) = "  PRIMARY KEY (`Prim`)"
 Str(1, 84, 9) = "  KEY `PatID` (`PatID`)"
 Str(1, 84, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr84

Function FüllStr85()
 Str(0, 85, 0) = "letze faelle"
 Str(1, 85, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `letze faelle` AS select `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik` from (`_lfaelle` `l` left join `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`))))"
End Function ' FüllStr85

Function FüllStr86()
 Str(0, 86, 0) = "letzte faelle"
 Str(0, 86, 1) = "`pat_id`"
 Str(0, 86, 2) = "`fid`"
 Str(0, 86, 3) = "`schgr`"
 Str(0, 86, 4) = "`bhfb`"
 Str(0, 86, 5) = "`ik`"
 ArtZ(0, 86) = 5
 Str(1, 86, 0) = "CREATE TABLE `letzte faelle` ("
 Str(1, 86, 1) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 86, 2) = " `fid` int(10) DEFAULT NULL"
 Str(1, 86, 3) = " `schgr` decimal(2,0) DEFAULT NULL"
 Str(1, 86, 4) = " `bhfb` datetime DEFAULT NULL"
 Str(1, 86, 5) = " `ik` varchar(50) DEFAULT NULL"
 Str(1, 86, 6) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Function ' FüllStr86

Function FüllStr87()
 Str(0, 87, 0) = "letztefaelleverschieden"
 Str(1, 87, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `letztefaelleverschieden` AS select `i`.`bhfb` AS `bhfb`,`i`.`pat_id` AS `pat_id` from `_f1` `i` group by `i`.`pat_id`"
End Function ' FüllStr87

Function FüllStr88()
 Str(0, 88, 0) = "lfaelle"
 Str(1, 88, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `lfaelle` AS select `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr`,`f`.`ÜbwV` AS `übwv` from (`_lfaelle` `l` left join `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`))))"
End Function ' FüllStr88

Function FüllStr89()
 Str(0, 89, 0) = "lfaellev"
 Str(1, 89, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `lfaellev` AS select `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr` from (`_lfaelle` `l` left join `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`)))) group by `f`.`Pat_ID`"
End Function ' FüllStr89

Function FüllStr90()
 Str(0, 90, 0) = "listenausgabeuew"
 Str(0, 90, 1) = "`name`"
 Str(0, 90, 2) = "`vorname`"
 Str(0, 90, 3) = "`titelt`"
 Str(0, 90, 4) = "`fachgruppe`"
 Str(0, 90, 5) = "`strasse`"
 Str(0, 90, 6) = "`plz`"
 Str(0, 90, 7) = "`ort`"
 Str(0, 90, 8) = "`telefon`"
 Str(0, 90, 9) = "`fax`"
 Str(0, 90, 10) = "`kvnr`"
 Str(0, 90, 11) = "`aktdat`"
 Str(0, 90, 12) = "`id`"
 Str(0, 90, 13) = "`QZ`"
 Str(0, 90, 14) = "`überschrift`"
 Str(0, 90, 15) = "`dbnr`"
 Str(0, 90, 16) = "`bstelle`"
 Str(0, 90, 17) = "`anrede`"
 Str(0, 90, 18) = "`tel2`"
 Str(0, 90, 19) = "`tel3`"
 Str(0, 90, 20) = "`tel4`"
 Str(0, 90, 21) = "`fax2`"
 Str(0, 90, 22) = "`fax3`"
 Str(0, 90, 23) = "`email`"
 Str(0, 90, 24) = "`zulg`"
 Str(0, 90, 25) = "`arzttyp`"
 Str(0, 90, 26) = "`gemmit`"
 Str(0, 90, 27) = "`beme`"
 Str(0, 90, 28) = "`dmpt2`"
 Str(0, 90, 29) = "`dmpt1`"
 Str(0, 90, 30) = "`geschlecht`"
 Str(0, 90, 31) = "`titel`"
 Str(0, 90, 32) = "`id`"
 Str(0, 90, 33) = "`kvnr`"
 Str(0, 90, 34) = "`name`"
 Str(0, 90, 35) = "`fax`"
 ArtZ(0, 90) = 31
 ArtZ(1, 90) = 4
 Str(1, 90, 0) = "CREATE TABLE `listenausgabeuew` ("
 Str(1, 90, 1) = " `name` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 2) = " `vorname` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 3) = " `titelt` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 4) = " `fachgruppe` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 5) = " `strasse` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 6) = " `plz` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 7) = " `ort` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 8) = " `telefon` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 90, 9) = " `fax` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 90, 10) = " `kvnr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 11) = " `aktdat` datetime DEFAULT NULL"
 Str(1, 90, 12) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 90, 13) = " `QZ` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Mitglied im Qualitätszirkel Stoffwechsel- und Gefäßerkrankungen'"
 Str(1, 90, 14) = " `überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 15) = " `dbnr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 16) = " `bstelle` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 17) = " `anrede` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 18) = " `tel2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 19) = " `tel3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 20) = " `tel4` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 21) = " `fax2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 22) = " `fax3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 23) = " `email` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 24) = " `zulg` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 25) = " `arzttyp` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 26) = " `gemmit` longtext COLLATE latin1_german2_ci"
 Str(1, 90, 27) = " `beme` longtext COLLATE latin1_german2_ci"
 Str(1, 90, 28) = " `dmpt2` bit(1) DEFAULT NULL"
 Str(1, 90, 29) = " `dmpt1` bit(1) DEFAULT NULL"
 Str(1, 90, 30) = " `geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 31) = " `titel` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 90, 32) = "  PRIMARY KEY (`id`)"
 Str(1, 90, 33) = "  KEY `kvnr` (`kvnr`)"
 Str(1, 90, 34) = "  KEY `name` (`name`)"
 Str(1, 90, 35) = "  KEY `fax` (`fax`)"
 Str(1, 90, 36) = " ENGINE=InnoDB AUTO_INCREMENT=976 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr90

Function FüllStr91()
 Str(0, 91, 0) = "lmp"
 Str(1, 91, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`mysql`@`%` SQL SECURITY DEFINER VIEW `lmp` AS select `mp`.`FID` AS `FID`,`mp`.`Pat_ID` AS `Pat_ID`,`mp`.`MPNr` AS `MPNr`,`mp`.`ZeitPunkt` AS `ZeitPunkt`,`mp`.`Datum` AS `Datum`,`mp`.`Medikament` AS `Medikament`,`mp`.`MedAnfang` AS `MedAnfang`,`mp`.`FeldNr` AS `FeldNr`,`mp`.`mo` AS `mo`,`mp`.`mi` AS `mi`,`mp`.`nm` AS `nm`,`mp`.`ab` AS `ab`,`mp`.`zn` AS `zn`,`mp`.`bBed` AS `bBed`,`mp`.`Bemerkung` AS `Bemerkung`,`mp`.`AbsPos` AS `AbsPos`,`mp`.`AktZeit` AS `AktZeit`,`mp`.`StByte` AS `StByte` from (`_fuerlmp` `i` join `medplan` `mp` on(((`i`.`pat_id` = `mp`.`Pat_ID`) and (`i`.`mpnr` = `mp`.`MPNr`))))"
End Function ' FüllStr91

Function FüllStr92()
 Str(0, 92, 0) = "medarten"
 Str(0, 92, 1) = "`Medikament`"
 Str(0, 92, 2) = "`Langname`"
 Str(0, 92, 3) = "`Pat_ID`"
 Str(0, 92, 4) = "`Anzahl`"
 Str(0, 92, 5) = "`Glib`"
 Str(0, 92, 6) = "`Metf`"
 Str(0, 92, 7) = "`GlucI`"
 Str(0, 92, 8) = "`SHGlin`"
 Str(0, 92, 9) = "`Glit`"
 Str(0, 92, 10) = "`SonstAD`"
 Str(0, 92, 11) = "`Ins`"
 Str(0, 92, 12) = "`Anal`"
 Str(0, 92, 13) = "`InsArt`"
 Str(0, 92, 14) = "`HMG`"
 Str(0, 92, 15) = "`Hypt`"
 Str(0, 92, 16) = "`Thro`"
 Str(0, 92, 17) = "`Antib`"
 Str(0, 92, 18) = "`and`"
 Str(0, 92, 19) = "`hinzugefügt`"
 Str(0, 92, 20) = "`Tstr`"
 Str(0, 92, 21) = "`Puzu`"
 Str(0, 92, 22) = "`VMat`"
 Str(0, 92, 23) = "`PenN`"
 Str(0, 92, 24) = "`Neurp`"
 Str(0, 92, 25) = "`AutNP`"
 Str(0, 92, 26) = "`Fetts`"
 Str(0, 92, 27) = "`Hsre`"
 Str(0, 92, 28) = "`AntiMyk`"
 Str(0, 92, 29) = "`Glauk`"
 Str(0, 92, 30) = "`COLD`"
 Str(0, 92, 31) = "`Pros`"
 Str(0, 92, 32) = "`Urä`"
 Str(0, 92, 33) = "`HyThy`"
 Str(0, 92, 34) = "`Ostp`"
 Str(0, 92, 35) = "`KHK`"
 Str(0, 92, 36) = "`HerzI`"
 Str(0, 92, 37) = "`Stru`"
 Str(0, 92, 38) = "`AVK`"
 Str(0, 92, 39) = "`PanI`"
 Str(0, 92, 40) = "`Vari`"
 Str(0, 92, 41) = "`Östr`"
 Str(0, 92, 42) = "`AntiDep`"
 Str(0, 92, 43) = "`AntiDem`"
 Str(0, 92, 44) = "`AntiEp`"
 Str(0, 92, 45) = "`Park`"
 Str(0, 92, 46) = "`AntiPern`"
 Str(0, 92, 47) = "`Appet`"
 Str(0, 92, 48) = "`Anäm`"
 Str(0, 92, 49) = "`Antiherp`"
 Str(0, 92, 50) = "`NSAR`"
 Str(0, 92, 51) = "`Antikoag`"
 Str(0, 92, 52) = "`Betabl`"
 Str(0, 92, 53) = "`ACEH`"
 Str(0, 92, 54) = "`AT1`"
 Str(0, 92, 55) = "`CalcA`"
 Str(0, 92, 56) = "`Diur`"
 Str(0, 92, 57) = "`falsch`"
 Str(0, 92, 58) = "`ID`"
 Str(0, 92, 59) = "`ID`"
 Str(0, 92, 60) = "`Medikament`"
 Str(0, 92, 61) = "`pat_id`"
 Str(0, 92, 62) = "`NamenMedArten_AccRel`"
 ArtZ(0, 92) = 58
 ArtZ(1, 92) = 3
 ArtZ(2, 92) = 1
 Str(1, 92, 0) = "CREATE TABLE `medarten` ("
 Str(1, 92, 1) = " `Medikament` varchar(50) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL"
 Str(1, 92, 2) = " `Langname` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Beispiel-Langname'"
 Str(1, 92, 3) = " `Pat_ID` int(10) DEFAULT NULL COMMENT 'Beispiel-PatID'"
 Str(1, 92, 4) = " `Anzahl` int(10) DEFAULT NULL COMMENT 'Anzahl der Vorkommen'"
 Str(1, 92, 5) = " `Glib` tinyint(1) unsigned NOT NULL COMMENT 'Glibenclamid'"
 Str(1, 92, 6) = " `Metf` tinyint(1) unsigned NOT NULL COMMENT 'Metformin'"
 Str(1, 92, 7) = " `GlucI` tinyint(1) unsigned NOT NULL COMMENT 'Glucosidase-Inhibitoren'"
 Str(1, 92, 8) = " `SHGlin` tinyint(1) unsigned NOT NULL COMMENT 'andere Sulfonylharnstoffe oder Glinide'"
 Str(1, 92, 9) = " `Glit` tinyint(1) unsigned NOT NULL COMMENT 'Glitanzone'"
 Str(1, 92, 10) = " `SonstAD` tinyint(1) unsigned NOT NULL COMMENT 'Sonstige antidiabetische Medikation'"
 Str(1, 92, 11) = " `Ins` tinyint(1) unsigned NOT NULL COMMENT 'Insulin'"
 Str(1, 92, 12) = " `Anal` tinyint(1) unsigned NOT NULL COMMENT 'Insulin-Analoga'"
 Str(1, 92, 13) = " `InsArt` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '1= schnell, 2 = langsam, 3 = Misch'"
 Str(1, 92, 14) = " `HMG` tinyint(1) unsigned NOT NULL COMMENT 'HMG-CoA-Reduktase-Inhibitoren'"
 Str(1, 92, 15) = " `Hypt` tinyint(1) unsigned NOT NULL COMMENT 'Hypertonie-Mittel'"
 Str(1, 92, 16) = " `Thro` tinyint(1) unsigned NOT NULL COMMENT 'Thrombozyten-Hemmer'"
 Str(1, 92, 17) = " `Antib` tinyint(1) unsigned NOT NULL COMMENT 'Antibiotika'"
 Str(1, 92, 18) = " `and` tinyint(1) unsigned NOT NULL COMMENT 'andere'"
 Str(1, 92, 19) = " `hinzugefügt` datetime DEFAULT NULL"
 Str(1, 92, 20) = " `Tstr` tinyint(1) unsigned NOT NULL COMMENT 'Teststreifen'"
 Str(1, 92, 21) = " `Puzu` tinyint(1) unsigned NOT NULL COMMENT 'Pumpenzubehör'"
 Str(1, 92, 22) = " `VMat` tinyint(1) unsigned NOT NULL COMMENT 'Verbandsmaterial'"
 Str(1, 92, 23) = " `PenN` tinyint(1) unsigned NOT NULL COMMENT 'Pennadeln'"
 Str(1, 92, 24) = " `Neurp` tinyint(1) unsigned NOT NULL COMMENT 'Neuropathie-Behandlungsmittel'"
 Str(1, 92, 25) = " `AutNP` tinyint(1) unsigned NOT NULL COMMENT 'Autonome Neuropathie'"
 Str(1, 92, 26) = " `Fetts` tinyint(1) unsigned NOT NULL COMMENT 'Fibrate, Ezetrol, Niaspan u.a.'"
 Str(1, 92, 27) = " `Hsre` tinyint(1) unsigned NOT NULL COMMENT 'Hyperuriämie-Mittel'"
 Str(1, 92, 28) = " `AntiMyk` tinyint(1) unsigned NOT NULL COMMENT 'Antimykotika'"
 Str(1, 92, 29) = " `Glauk` tinyint(1) unsigned NOT NULL COMMENT 'Glaukom'"
 Str(1, 92, 30) = " `COLD` tinyint(1) unsigned NOT NULL COMMENT 'COLD und Asthma'"
 Str(1, 92, 31) = " `Pros` tinyint(1) unsigned NOT NULL COMMENT 'Prostatahypertrophie'"
 Str(1, 92, 32) = " `Urä` tinyint(1) unsigned NOT NULL COMMENT 'Urämie-spezifische'"
 Str(1, 92, 33) = " `HyThy` tinyint(1) unsigned NOT NULL COMMENT 'thyreostatische Mittel'"
 Str(1, 92, 34) = " `Ostp` tinyint(1) unsigned NOT NULL COMMENT 'Osteoporosemittel'"
 Str(1, 92, 35) = " `KHK` tinyint(1) unsigned NOT NULL COMMENT 'KHK-spezifisch'"
 Str(1, 92, 36) = " `HerzI` tinyint(1) unsigned NOT NULL COMMENT 'Herzinsuffizienz-spezifische'"
 Str(1, 92, 37) = " `Stru` tinyint(1) unsigned NOT NULL COMMENT 'Struma- und Hypothyreosemittel'"
 Str(1, 92, 38) = " `AVK` tinyint(1) unsigned NOT NULL COMMENT 'AVK-Mittel'"
 Str(1, 92, 39) = " `PanI` tinyint(1) unsigned NOT NULL COMMENT 'Pankreasinsuffizienz'"
 Str(1, 92, 40) = " `Vari` tinyint(1) unsigned NOT NULL COMMENT 'Varikosemittel'"
 Str(1, 92, 41) = " `Östr` tinyint(1) unsigned NOT NULL COMMENT 'Östrogene, Gestagene usw.'"
 Str(1, 92, 42) = " `AntiDep` tinyint(1) unsigned NOT NULL COMMENT 'Antidepressiva'"
 Str(1, 92, 43) = " `AntiDem` tinyint(1) unsigned NOT NULL COMMENT 'Antidementika'"
 Str(1, 92, 44) = " `AntiEp` tinyint(1) unsigned NOT NULL COMMENT 'Antiepileptika'"
 Str(1, 92, 45) = " `Park` tinyint(1) unsigned NOT NULL COMMENT 'Parkinson-Medikament'"
 Str(1, 92, 46) = " `AntiPern` tinyint(1) unsigned NOT NULL COMMENT 'Antiperniziosa'"
 Str(1, 92, 47) = " `Appet` tinyint(1) unsigned NOT NULL COMMENT 'Appetitzügler'"
 Str(1, 92, 48) = " `Anäm` tinyint(1) unsigned NOT NULL COMMENT 'Anämiebehandlungsmittel'"
 Str(1, 92, 49) = " `Antiherp` tinyint(1) unsigned NOT NULL COMMENT 'Anti-Herpes-Mittel'"
 Str(1, 92, 50) = " `NSAR` tinyint(1) unsigned NOT NULL COMMENT 'NSAR'"
 Str(1, 92, 51) = " `Antikoag` tinyint(1) unsigned NOT NULL COMMENT 'Antikoagulatien'"
 Str(1, 92, 52) = " `Betabl` tinyint(1) unsigned NOT NULL COMMENT 'Betablocker'"
 Str(1, 92, 53) = " `ACEH` tinyint(1) unsigned NOT NULL COMMENT 'ACE-Hemmer'"
 Str(1, 92, 54) = " `AT1` tinyint(1) unsigned NOT NULL COMMENT 'AT-1-Blocker'"
 Str(1, 92, 55) = " `CalcA` tinyint(1) unsigned NOT NULL COMMENT 'Calcium-Antagonist'"
 Str(1, 92, 56) = " `Diur` tinyint(1) unsigned NOT NULL COMMENT 'Diuretikum'"
 Str(1, 92, 57) = " `falsch` tinyint(1) unsigned NOT NULL COMMENT 'falsch geschrieben oder nicht erkennbar'"
 Str(1, 92, 58) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'"
 Str(1, 92, 59) = "  PRIMARY KEY (`ID`)"
 Str(1, 92, 60) = "  UNIQUE KEY `Medikament` (`Medikament`)"
 Str(1, 92, 61) = "  KEY `pat_id` (`Pat_ID`)"
 Str(1, 92, 62) = "  CONSTRAINT `NamenMedArten_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`)"
 Str(1, 92, 63) = " ENGINE=InnoDB AUTO_INCREMENT=3546 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr92

Function FüllStr93()
 Str(0, 93, 0) = "medplan"
 Str(0, 93, 1) = "`FID`"
 Str(0, 93, 2) = "`Pat_ID`"
 Str(0, 93, 3) = "`MPNr`"
 Str(0, 93, 4) = "`ZeitPunkt`"
 Str(0, 93, 5) = "`Datum`"
 Str(0, 93, 6) = "`Medikament`"
 Str(0, 93, 7) = "`MedAnfang`"
 Str(0, 93, 8) = "`FeldNr`"
 Str(0, 93, 9) = "`mo`"
 Str(0, 93, 10) = "`mi`"
 Str(0, 93, 11) = "`nm`"
 Str(0, 93, 12) = "`ab`"
 Str(0, 93, 13) = "`zn`"
 Str(0, 93, 14) = "`bBed`"
 Str(0, 93, 15) = "`Bemerkung`"
 Str(0, 93, 16) = "`AbsPos`"
 Str(0, 93, 17) = "`AktZeit`"
 Str(0, 93, 18) = "`StByte`"
 Str(0, 93, 19) = "`FälleMedPlan`"
 Str(0, 93, 20) = "`MedPlanMedikament`"
 Str(0, 93, 21) = "`NamenMedPlan`"
 Str(0, 93, 22) = "`MedArtenMedPlan_AccRel`"
 Str(0, 93, 23) = "`MPNr`"
 Str(0, 93, 24) = "`MPNrPat_ID`"
 Str(0, 93, 25) = "`Auswahl`"
 Str(0, 93, 26) = "`FälleMedPlan_AccRel`"
 Str(0, 93, 27) = "`MedArtenMedPlan_AccRel`"
 Str(0, 93, 28) = "`NamenMedPlan_AccRel`"
 ArtZ(0, 93) = 18
 ArtZ(1, 93) = 7
 ArtZ(2, 93) = 3
 Str(1, 93, 0) = "CREATE TABLE `medplan` ("
 Str(1, 93, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 93, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 93, 3) = " `MPNr` int(10) DEFAULT NULL COMMENT 'Ordnungsziffer für Medikamentenplan'"
 Str(1, 93, 4) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt, der Speicherung im Turbomed'"
 Str(1, 93, 5) = " `Datum` datetime DEFAULT NULL COMMENT 'Zeitpunkt aus dem Kopf des Medikamentenplans'"
 Str(1, 93, 6) = " `Medikament` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 7) = " `MedAnfang` varchar(35) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL"
 Str(1, 93, 8) = " `FeldNr` smallint(6) DEFAULT NULL"
 Str(1, 93, 9) = " `mo` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 10) = " `mi` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 11) = " `nm` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 12) = " `ab` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 13) = " `zn` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 93, 14) = " `bBed` bit(1) DEFAULT NULL"
 Str(1, 93, 15) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 93, 16) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 93, 17) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 93, 18) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 93, 19) = "  KEY `FälleMedPlan` (`FID`)"
 Str(1, 93, 20) = "  KEY `MedPlanMedikament` (`Medikament`)"
 Str(1, 93, 21) = "  KEY `NamenMedPlan` (`Pat_ID`)"
 Str(1, 93, 22) = "  KEY `MedArtenMedPlan_AccRel` (`MedAnfang`)"
 Str(1, 93, 23) = "  KEY `MPNr` (`MPNr`)"
 Str(1, 93, 24) = "  KEY `MPNrPat_ID` (`Pat_ID`,`MPNr`)"
 Str(1, 93, 25) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`MPNr`) USING BTREE"
 Str(1, 93, 26) = "  CONSTRAINT `FälleMedPlan_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 93, 27) = "  CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`)"
 Str(1, 93, 28) = "  CONSTRAINT `NamenMedPlan_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 93, 29) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr93

Function FüllStr94()
 Str(0, 94, 0) = "namen"
 Str(0, 94, 1) = "`Pat_ID`"
 Str(0, 94, 2) = "`lfdnr`"
 Str(0, 94, 3) = "`NVorsatz`"
 Str(0, 94, 4) = "`Nachname`"
 Str(0, 94, 5) = "`Vorname`"
 Str(0, 94, 6) = "`GebDat`"
 Str(0, 94, 7) = "`Straße`"
 Str(0, 94, 8) = "`KVKStatus`"
 Str(0, 94, 9) = "`Geschlecht`"
 Str(0, 94, 10) = "`Plz`"
 Str(0, 94, 11) = "`Ort`"
 Str(0, 94, 12) = "`Weggeldzone`"
 Str(0, 94, 13) = "`WeggzZahl`"
 Str(0, 94, 14) = "`AufnDat`"
 Str(0, 94, 15) = "`LANR`"
 Str(0, 94, 16) = "`BStNr`"
 Str(0, 94, 17) = "`Titel`"
 Str(0, 94, 18) = "`Versichertennummer`"
 Str(0, 94, 19) = "`PrivatTel`"
 Str(0, 94, 20) = "`KVNr`"
 Str(0, 94, 21) = "`PrivatTel_2`"
 Str(0, 94, 22) = "`PrivatFax`"
 Str(0, 94, 23) = "`DienstTel`"
 Str(0, 94, 24) = "`PrivatMobil`"
 Str(0, 94, 25) = "`Email`"
 Str(0, 94, 26) = "`Arbeitgeber`"
 Str(0, 94, 27) = "`AnAllgda`"
 Str(0, 94, 28) = "`An1da`"
 Str(0, 94, 29) = "`An2da`"
 Str(0, 94, 30) = "`Checkda`"
 Str(0, 94, 31) = "`DMTypaD`"
 Str(0, 94, 32) = "`AktZeit`"
 Str(0, 94, 33) = "`absPos`"
 Str(0, 94, 34) = "`StByte`"
 Str(0, 94, 35) = "`Cave`"
 Str(0, 94, 36) = "`Notiz`"
 Str(0, 94, 37) = "`zubenach`"
 Str(0, 94, 38) = "`Verwandt`"
 Str(0, 94, 39) = "`Sprache`"
 Str(0, 94, 40) = "`lAktTM`"
 Str(0, 94, 41) = "`PAT_ID`"
 Str(0, 94, 42) = "`Auswahl`"
 Str(0, 94, 43) = "`HausärzteNamen_AccRel`"
 Str(0, 94, 44) = "`weggszahl`"
 Str(0, 94, 45) = "`weggeldzone`"
 Str(0, 94, 46) = "`HausärzteNamen_AccRel`"
 ArtZ(0, 94) = 40
 ArtZ(1, 94) = 5
 ArtZ(2, 94) = 1
 Str(1, 94, 0) = "CREATE TABLE `namen` ("
 Str(1, 94, 1) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 94, 2) = " `lfdnr` int(10) DEFAULT NULL COMMENT 'laufende Patientennummer'"
 Str(1, 94, 3) = " `NVorsatz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3100'"
 Str(1, 94, 4) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 94, 5) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 94, 6) = " `GebDat` datetime DEFAULT NULL COMMENT '3103'"
 Str(1, 94, 7) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3107'"
 Str(1, 94, 8) = " `KVKStatus` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3108'"
 Str(1, 94, 9) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3110'"
 Str(1, 94, 10) = " `Plz` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3112'"
 Str(1, 94, 11) = " `Ort` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3113'"
 Str(1, 94, 12) = " `Weggeldzone` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3631 (1) Weggeldzone mit Z'"
 Str(1, 94, 13) = " `WeggzZahl` decimal(1,0) DEFAULT NULL COMMENT '3631 (2) Weggeldzone, Zahl in Feld 2'"
 Str(1, 94, 14) = " `AufnDat` datetime DEFAULT NULL COMMENT '3610'"
 Str(1, 94, 15) = " `LANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP'"
 Str(1, 94, 16) = " `BStNr` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3536 Betriebsstättennummer'"
 Str(1, 94, 17) = " `Titel` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3104'"
 Str(1, 94, 18) = " `Versichertennummer` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3105'"
 Str(1, 94, 19) = " `PrivatTel` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 94, 20) = " `KVNr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3630'"
 Str(1, 94, 21) = " `PrivatTel_2` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 94, 22) = " `PrivatFax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 94, 23) = " `DienstTel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 94, 24) = " `PrivatMobil` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 94, 25) = " `Email` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Email'"
 Str(1, 94, 26) = " `Arbeitgeber` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3625'"
 Str(1, 94, 27) = " `AnAllgda` bit(1) DEFAULT NULL COMMENT 'Anamnese allgemein da'"
 Str(1, 94, 28) = " `An1da` bit(1) DEFAULT NULL COMMENT 'Anamnese S.1 da'"
 Str(1, 94, 29) = " `An2da` bit(1) DEFAULT NULL COMMENT 'Anamnese S.2 da'"
 Str(1, 94, 30) = " `Checkda` bit(1) DEFAULT NULL COMMENT 'Checkliste da'"
 Str(1, 94, 31) = " `DMTypaD` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'aus Diagnosen'"
 Str(1, 94, 32) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 94, 33) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 94, 34) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 94, 35) = " `Cave` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3654'"
 Str(1, 94, 36) = " `Notiz` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>'"
 Str(1, 94, 37) = " `zubenach` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3633'"
 Str(1, 94, 38) = " `Verwandt` varchar(77) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3632'"
 Str(1, 94, 39) = " `Sprache` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3628'"
 Str(1, 94, 40) = " `lAktTM` datetime NOT NULL COMMENT 'letzte Aktualisierung in Turbomed'"
 Str(1, 94, 41) = "  UNIQUE KEY `PAT_ID` (`Pat_ID`)"
 Str(1, 94, 42) = "  KEY `Auswahl` (`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 94, 43) = "  KEY `HausärzteNamen_AccRel` (`KVNr`)"
 Str(1, 94, 44) = "  KEY `weggszahl` (`WeggzZahl`)"
 Str(1, 94, 45) = "  KEY `weggeldzone` (`Weggeldzone`)"
 Str(1, 94, 46) = "  CONSTRAINT `HausärzteNamen_AccRel` FOREIGN KEY (`KVNr`) REFERENCES `hausaerzte` (`KVNr`)"
 Str(1, 94, 47) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr94

Function FüllStr95()
 Str(0, 95, 0) = "pauschalen"
 Str(0, 95, 1) = "`Leistung`"
 Str(0, 95, 2) = "`Betreuung`"
 ArtZ(0, 95) = 2
 Str(1, 95, 0) = "CREATE TABLE `pauschalen` ("
 Str(1, 95, 1) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 95, 2) = " `Betreuung` bit(1) DEFAULT NULL"
 Str(1, 95, 3) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr95

Function FüllStr96()
 Str(0, 96, 0) = "pumpenträger"
 Str(0, 96, 1) = "`fall`"
 Str(0, 96, 2) = "`anrede`"
 Str(0, 96, 3) = "`nachname`"
 Str(0, 96, 4) = "`vorname`"
 Str(0, 96, 5) = "`gebdat`"
 Str(0, 96, 6) = "`privattel`"
 Str(0, 96, 7) = "`privattel_2`"
 Str(0, 96, 8) = "`privatfax`"
 Str(0, 96, 9) = "`diensttel`"
 Str(0, 96, 10) = "`straße`"
 Str(0, 96, 11) = "`plz`"
 Str(0, 96, 12) = "`ort`"
 Str(0, 96, 13) = "`mail1`"
 Str(0, 96, 14) = "`mail2`"
 Str(0, 96, 15) = "`email`"
 ArtZ(0, 96) = 15
 Str(1, 96, 0) = "CREATE TABLE `pumpenträger` ("
 Str(1, 96, 1) = " `fall` varchar(2) DEFAULT NULL"
 Str(1, 96, 2) = " `anrede` varchar(4) DEFAULT NULL"
 Str(1, 96, 3) = " `nachname` varchar(21) DEFAULT NULL"
 Str(1, 96, 4) = " `vorname` varchar(19) DEFAULT NULL"
 Str(1, 96, 5) = " `gebdat` datetime DEFAULT NULL"
 Str(1, 96, 6) = " `privattel` varchar(100) DEFAULT NULL"
 Str(1, 96, 7) = " `privattel_2` varchar(50) DEFAULT NULL"
 Str(1, 96, 8) = " `privatfax` varchar(50) DEFAULT NULL"
 Str(1, 96, 9) = " `diensttel` varchar(50) DEFAULT NULL"
 Str(1, 96, 10) = " `straße` varchar(50) DEFAULT NULL"
 Str(1, 96, 11) = " `plz` varchar(20) DEFAULT NULL"
 Str(1, 96, 12) = " `ort` varchar(70) DEFAULT NULL"
 Str(1, 96, 13) = " `mail1` varchar(100) DEFAULT NULL"
 Str(1, 96, 14) = " `mail2` varchar(100) DEFAULT NULL"
 Str(1, 96, 15) = " `email` varchar(100) DEFAULT NULL"
 Str(1, 96, 16) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Function ' FüllStr96

Function FüllStr97()
 Str(0, 97, 0) = "queries"
 Str(0, 97, 1) = "`Name`"
 Str(0, 97, 2) = "`sql`"
 Str(0, 97, 3) = "`gespeichert`"
 Str(0, 97, 4) = "`Name`"
 Str(0, 97, 5) = "`NamGesp`"
 Str(0, 97, 6) = "`sql`"
 ArtZ(0, 97) = 3
 ArtZ(1, 97) = 3
 Str(1, 97, 0) = "CREATE TABLE `queries` ("
 Str(1, 97, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Name der SQL-Abfrage'"
 Str(1, 97, 2) = " `sql` longtext COLLATE latin1_german2_ci COMMENT 'SQL-String, der geht'"
 Str(1, 97, 3) = " `gespeichert` datetime NOT NULL DEFAULT '0000-00-00 00:00:00' COMMENT 'Zeitpunkt, zu der er geht'"
 Str(1, 97, 4) = "  PRIMARY KEY (`Name`,`gespeichert`)"
 Str(1, 97, 5) = "  UNIQUE KEY `NamGesp` (`Name`,`gespeichert`)"
 Str(1, 97, 6) = "  KEY `sql` (`sql`(255))"
 Str(1, 97, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr97

Function FüllStr98()
 Str(0, 98, 0) = "rel2"
 Str(0, 98, 1) = "`ccolumn`"
 Str(0, 98, 2) = "`grbit`"
 Str(0, 98, 3) = "`icolumn`"
 Str(0, 98, 4) = "`szColumn`"
 Str(0, 98, 5) = "`szObject`"
 Str(0, 98, 6) = "`szReferencedColumn`"
 Str(0, 98, 7) = "`szReferencedObject`"
 Str(0, 98, 8) = "`szRelationship`"
 Str(0, 98, 9) = "`szObject`"
 Str(0, 98, 10) = "`szReferencedObject`"
 Str(0, 98, 11) = "`szRelationship`"
 ArtZ(0, 98) = 8
 ArtZ(1, 98) = 3
 Str(1, 98, 0) = "CREATE TABLE `rel2` ("
 Str(1, 98, 1) = " `ccolumn` int(10) DEFAULT NULL"
 Str(1, 98, 2) = " `grbit` int(10) DEFAULT NULL"
 Str(1, 98, 3) = " `icolumn` int(10) DEFAULT NULL"
 Str(1, 98, 4) = " `szColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 98, 5) = " `szObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 98, 6) = " `szReferencedColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 98, 7) = " `szReferencedObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 98, 8) = " `szRelationship` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 98, 9) = "  KEY `szObject` (`szObject`)"
 Str(1, 98, 10) = "  KEY `szReferencedObject` (`szReferencedObject`)"
 Str(1, 98, 11) = "  KEY `szRelationship` (`szRelationship`)"
 Str(1, 98, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr98

Function FüllStr99()
 Str(0, 99, 0) = "relationen"
 Str(0, 99, 1) = "`ccolumn`"
 Str(0, 99, 2) = "`grbit`"
 Str(0, 99, 3) = "`icolumn`"
 Str(0, 99, 4) = "`szColumn`"
 Str(0, 99, 5) = "`szObject`"
 Str(0, 99, 6) = "`szReferencedColumn`"
 Str(0, 99, 7) = "`szReferencedObject`"
 Str(0, 99, 8) = "`szRelationship`"
 Str(0, 99, 9) = "`szObject`"
 Str(0, 99, 10) = "`szReferencedObject`"
 Str(0, 99, 11) = "`szRelationship`"
 ArtZ(0, 99) = 8
 ArtZ(1, 99) = 3
 Str(1, 99, 0) = "CREATE TABLE `relationen` ("
 Str(1, 99, 1) = " `ccolumn` int(10) DEFAULT NULL"
 Str(1, 99, 2) = " `grbit` int(10) DEFAULT NULL"
 Str(1, 99, 3) = " `icolumn` int(10) DEFAULT NULL"
 Str(1, 99, 4) = " `szColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 99, 5) = " `szObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 99, 6) = " `szReferencedColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 99, 7) = " `szReferencedObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 99, 8) = " `szRelationship` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 99, 9) = "  KEY `szObject` (`szObject`)"
 Str(1, 99, 10) = "  KEY `szReferencedObject` (`szReferencedObject`)"
 Str(1, 99, 11) = "  KEY `szRelationship` (`szRelationship`)"
 Str(1, 99, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr99

Function FüllStr100()
 Str(0, 100, 0) = "rezepteintraege"
 Str(0, 100, 1) = "`FID`"
 Str(0, 100, 2) = "`Pat_ID`"
 Str(0, 100, 3) = "`ZeitPunkt`"
 Str(0, 100, 4) = "`Rezept`"
 Str(0, 100, 5) = "`Rezeptklasse`"
 Str(0, 100, 6) = "`Medikament`"
 Str(0, 100, 7) = "`PZN`"
 Str(0, 100, 8) = "`absPos`"
 Str(0, 100, 9) = "`AktZeit`"
 Str(0, 100, 10) = "`QS`"
 Str(0, 100, 11) = "`QT`"
 Str(0, 100, 12) = "`StByte`"
 Str(0, 100, 13) = "`Auswahl`"
 Str(0, 100, 14) = "`FälleRezeptEinträge`"
 Str(0, 100, 15) = "`FID`"
 Str(0, 100, 16) = "`NamenRezeptEinträge`"
 Str(0, 100, 17) = "`FälleRezeptEinträge_AccRel`"
 Str(0, 100, 18) = "`NamenRezeptEinträge_AccRel`"
 ArtZ(0, 100) = 12
 ArtZ(1, 100) = 4
 ArtZ(2, 100) = 2
 Str(1, 100, 0) = "CREATE TABLE `rezepteintraege` ("
 Str(1, 100, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 100, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 100, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 100, 4) = " `Rezept` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6210, 3652(1), 6218(1)'"
 Str(1, 100, 5) = " `Rezeptklasse` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)'"
 Str(1, 100, 6) = " `Medikament` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3652(2), 6218(4)'"
 Str(1, 100, 7) = " `PZN` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6210(2), 6218(3)'"
 Str(1, 100, 8) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 100, 9) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 100, 10) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 100, 11) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 100, 12) = " `StByte` int(11) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 100, 13) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Rezept`,`Medikament`)"
 Str(1, 100, 14) = "  KEY `FälleRezeptEinträge` (`FID`)"
 Str(1, 100, 15) = "  KEY `FID` (`FID`)"
 Str(1, 100, 16) = "  KEY `NamenRezeptEinträge` (`Pat_ID`)"
 Str(1, 100, 17) = "  CONSTRAINT `FälleRezeptEinträge_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 100, 18) = "  CONSTRAINT `NamenRezeptEinträge_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 100, 19) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr100

Function FüllStr101()
 Str(0, 101, 0) = "rr"
 Str(0, 101, 1) = "`FID`"
 Str(0, 101, 2) = "`Pat_ID`"
 Str(0, 101, 3) = "`ZeitPunkt`"
 Str(0, 101, 4) = "`RR`"
 Str(0, 101, 5) = "`absPos`"
 Str(0, 101, 6) = "`AktZeit`"
 Str(0, 101, 7) = "`StByte`"
 Str(0, 101, 8) = "`Auswahl`"
 Str(0, 101, 9) = "`FälleRR`"
 Str(0, 101, 10) = "`FID`"
 Str(0, 101, 11) = "`NamenRR`"
 Str(0, 101, 12) = "`FälleRR_AccRel`"
 Str(0, 101, 13) = "`NamenRR_AccRel`"
 ArtZ(0, 101) = 7
 ArtZ(1, 101) = 4
 ArtZ(2, 101) = 2
 Str(1, 101, 0) = "CREATE TABLE `rr` ("
 Str(1, 101, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 101, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 101, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 101, 4) = " `RR` longtext COLLATE latin1_german2_ci COMMENT '6230'"
 Str(1, 101, 5) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 101, 6) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 101, 7) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 101, 8) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`RR`(255))"
 Str(1, 101, 9) = "  KEY `FälleRR` (`FID`)"
 Str(1, 101, 10) = "  KEY `FID` (`FID`)"
 Str(1, 101, 11) = "  KEY `NamenRR` (`Pat_ID`)"
 Str(1, 101, 12) = "  CONSTRAINT `FälleRR_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 101, 13) = "  CONSTRAINT `NamenRR_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 101, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr101

Function FüllStr102()
 Str(0, 102, 0) = "rrparse"
 Str(0, 102, 1) = "`Pat_id`"
 Str(0, 102, 2) = "`Zeitpunkt`"
 Str(0, 102, 3) = "`RRSyst`"
 Str(0, 102, 4) = "`RRDiast`"
 Str(0, 102, 5) = "`Quelle`"
 Str(0, 102, 6) = "`ID`"
 ArtZ(0, 102) = 5
 ArtZ(1, 102) = 1
 Str(1, 102, 0) = "CREATE TABLE `rrparse` ("
 Str(1, 102, 1) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 102, 2) = " `Zeitpunkt` datetime DEFAULT NULL"
 Str(1, 102, 3) = " `RRSyst` smallint(6) DEFAULT NULL"
 Str(1, 102, 4) = " `RRDiast` smallint(6) DEFAULT NULL"
 Str(1, 102, 5) = " `Quelle` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 102, 6) = "  KEY `ID` (`Pat_id`,`Zeitpunkt`,`RRSyst`,`RRDiast`)"
 Str(1, 102, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr102

Function FüllStr103()
 Str(0, 103, 0) = "schulleistzahl"
 Str(0, 103, 1) = "`quartal`"
 Str(0, 103, 2) = "`lzahl`"
 ArtZ(0, 103) = 2
 Str(1, 103, 0) = "CREATE TABLE `schulleistzahl` ("
 Str(1, 103, 1) = " `quartal` varchar(5) DEFAULT NULL"
 Str(1, 103, 2) = " `lzahl` bigint(21) DEFAULT NULL"
 Str(1, 103, 3) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Function ' FüllStr103

Function FüllStr104()
 Str(0, 104, 0) = "schulzahl"
 Str(0, 104, 1) = "`quartal`"
 Str(0, 104, 2) = "`lzahl`"
 Str(0, 104, 3) = "`ezahl`"
 ArtZ(0, 104) = 3
 Str(1, 104, 0) = "CREATE TABLE `schulzahl` ("
 Str(1, 104, 1) = " `quartal` varchar(5) DEFAULT NULL"
 Str(1, 104, 2) = " `lzahl` bigint(21) DEFAULT NULL"
 Str(1, 104, 3) = " `ezahl` bigint(21) DEFAULT NULL"
 Str(1, 104, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Function ' FüllStr104

Function FüllStr105()
 Str(0, 105, 0) = "tabelle2"
 Str(0, 105, 1) = "`id`"
 Str(0, 105, 2) = "`ob`"
 Str(0, 105, 3) = "`id`"
 ArtZ(0, 105) = 2
 ArtZ(1, 105) = 1
 Str(1, 105, 0) = "CREATE TABLE `tabelle2` ("
 Str(1, 105, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 105, 2) = " `ob` bit(1) DEFAULT NULL"
 Str(1, 105, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 105, 4) = " ENGINE=MyISAM AUTO_INCREMENT=3 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr105

Function FüllStr106()
 Str(0, 106, 0) = "tmpfif"
 Str(0, 106, 1) = "`FeldVW`"
 Str(0, 106, 2) = "`Feld`"
 Str(0, 106, 3) = "`StByte`"
 Str(0, 106, 4) = "`FeldVW`"
 Str(0, 106, 5) = "`Feld`"
 ArtZ(0, 106) = 3
 ArtZ(1, 106) = 2
 Str(1, 106, 0) = "CREATE TABLE `tmpfif` ("
 Str(1, 106, 1) = " `FeldVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 106, 2) = " `Feld` longtext COLLATE latin1_german2_ci"
 Str(1, 106, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 106, 4) = "  PRIMARY KEY (`FeldVW`)"
 Str(1, 106, 5) = "  KEY `Feld` (`Feld`(255))"
 Str(1, 106, 6) = " ENGINE=InnoDB AUTO_INCREMENT=13279 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr106

Function FüllStr107()
 Str(0, 107, 0) = "tmpkassenliste"
 Str(0, 107, 1) = "`ID`"
 Str(0, 107, 2) = "`VK`"
 Str(0, 107, 3) = "`IK`"
 Str(0, 107, 4) = "`Name`"
 Str(0, 107, 5) = "`Kateg`"
 Str(0, 107, 6) = "`AnzahlIK`"
 Str(0, 107, 7) = "`AnzahlKTUG`"
 Str(0, 107, 8) = "`GültigVon`"
 Str(0, 107, 9) = "`GültigBis`"
 Str(0, 107, 10) = "`GO`"
 Str(0, 107, 11) = "`Kurzname`"
 Str(0, 107, 12) = "`rName`"
 Str(0, 107, 13) = "`ID`"
 Str(0, 107, 14) = "`PrimaryKey`"
 Str(0, 107, 15) = "`VK`"
 Str(0, 107, 16) = "`IK`"
 Str(0, 107, 17) = "`VKIK`"
 ArtZ(0, 107) = 12
 ArtZ(1, 107) = 5
 Str(1, 107, 0) = "CREATE TABLE `tmpkassenliste` ("
 Str(1, 107, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 107, 2) = " `VK` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 3) = " `IK` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 4) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 5) = " `Kateg` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kategorie'"
 Str(1, 107, 6) = " `AnzahlIK` int(4) unsigned DEFAULT NULL"
 Str(1, 107, 7) = " `AnzahlKTUG` int(4) unsigned DEFAULT NULL"
 Str(1, 107, 8) = " `GültigVon` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 9) = " `GültigBis` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 10) = " `GO` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 11) = " `Kurzname` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 107, 12) = " `rName` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kurzname, falls nicht verfügbar: Name'"
 Str(1, 107, 13) = "  PRIMARY KEY (`ID`)"
 Str(1, 107, 14) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 107, 15) = "  KEY `VK` (`VK`)"
 Str(1, 107, 16) = "  KEY `IK` (`IK`)"
 Str(1, 107, 17) = "  KEY `VKIK` (`VK`,`IK`)"
 Str(1, 107, 18) = " ENGINE=InnoDB AUTO_INCREMENT=5274 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr107

Function FüllStr108()
 Str(0, 108, 0) = "tmprelationen"
 Str(0, 108, 1) = "`ccolumn`"
 Str(0, 108, 2) = "`grbit`"
 Str(0, 108, 3) = "`icolumn`"
 Str(0, 108, 4) = "`szColumn`"
 Str(0, 108, 5) = "`szObject`"
 Str(0, 108, 6) = "`szReferencedColumn`"
 Str(0, 108, 7) = "`szReferencedObject`"
 Str(0, 108, 8) = "`szRelationship`"
 ArtZ(0, 108) = 8
 Str(1, 108, 0) = "CREATE TABLE `tmprelationen` ("
 Str(1, 108, 1) = " `ccolumn` int(10) DEFAULT NULL"
 Str(1, 108, 2) = " `grbit` int(10) DEFAULT NULL"
 Str(1, 108, 3) = " `icolumn` int(10) DEFAULT NULL"
 Str(1, 108, 4) = " `szColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 5) = " `szObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 6) = " `szReferencedColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 7) = " `szReferencedObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 8) = " `szRelationship` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr108

Function FüllStr109()
 Str(0, 109, 0) = "ueberwvon"
 Str(0, 109, 1) = "`ID`"
 Str(0, 109, 2) = "`KVNr`"
 Str(0, 109, 3) = "`Titel`"
 Str(0, 109, 4) = "`Vorname`"
 Str(0, 109, 5) = "`Zusatz`"
 Str(0, 109, 6) = "`Nachname`"
 Str(0, 109, 7) = "`ID`"
 Str(0, 109, 8) = "`KVNr`"
 Str(0, 109, 9) = "`Name`"
 ArtZ(0, 109) = 6
 ArtZ(1, 109) = 3
 Str(1, 109, 0) = "CREATE TABLE `ueberwvon` ("
 Str(1, 109, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 109, 2) = " `KVNr` varchar(7) DEFAULT NULL"
 Str(1, 109, 3) = " `Titel` varchar(30) DEFAULT NULL"
 Str(1, 109, 4) = " `Vorname` varchar(45) DEFAULT NULL"
 Str(1, 109, 5) = " `Zusatz` varchar(30) DEFAULT NULL"
 Str(1, 109, 6) = " `Nachname` varchar(45) DEFAULT NULL"
 Str(1, 109, 7) = "  PRIMARY KEY (`ID`) USING BTREE"
 Str(1, 109, 8) = "  KEY `KVNr` (`KVNr`)"
 Str(1, 109, 9) = "  KEY `Name` (`Nachname`,`Vorname`)"
 Str(1, 109, 10) = " ENGINE=InnoDB AUTO_INCREMENT=5334 DEFAULT CHARSET=latin1 COMMENT='4247 Überwiesen von'"
End Function ' FüllStr109

Function FüllStr110()
 Str(0, 110, 0) = "unbekannte kennungen"
 Str(0, 110, 1) = "`Kennung`"
 Str(0, 110, 2) = "`absPos`"
 Str(0, 110, 3) = "`StByte`"
 Str(0, 110, 4) = "`Pat_id`"
 Str(0, 110, 5) = "`Inhalt`"
 Str(0, 110, 6) = "`Kennung`"
 ArtZ(0, 110) = 5
 ArtZ(1, 110) = 1
 Str(1, 110, 0) = "CREATE TABLE `unbekannte kennungen` ("
 Str(1, 110, 1) = " `Kennung` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 110, 2) = " `absPos` int(10) DEFAULT NULL"
 Str(1, 110, 3) = " `StByte` int(10) DEFAULT NULL"
 Str(1, 110, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT 'zugehöriger Patient für spätere Ermittlungen'"
 Str(1, 110, 5) = " `Inhalt` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Inhalt Zeile zum Wiederauffinden'"
 Str(1, 110, 6) = "  UNIQUE KEY `Kennung` (`Kennung`)"
 Str(1, 110, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr110

Function FüllStr111()
 Str(0, 111, 0) = "vlassen"
 Str(0, 111, 1) = "`Datum`"
 Str(0, 111, 2) = "`Name`"
 Str(0, 111, 3) = "`Größe`"
 Str(0, 111, 4) = "`pfad`"
 Str(0, 111, 5) = "`Datum`"
 Str(0, 111, 6) = "`Name`"
 ArtZ(0, 111) = 4
 ArtZ(1, 111) = 2
 Str(1, 111, 0) = "CREATE TABLE `vlassen` ("
 Str(1, 111, 1) = " `Datum` datetime DEFAULT NULL"
 Str(1, 111, 2) = " `Name` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 111, 3) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 111, 4) = " `pfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 111, 5) = "  KEY `Datum` (`Datum`)"
 Str(1, 111, 6) = "  KEY `Name` (`Name`)"
 Str(1, 111, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Function ' FüllStr111

Function FüllStr112()
 Str(0, 112, 0) = "werte_scheingruppen"
 Str(0, 112, 1) = "`schgr`"
 Str(0, 112, 2) = "`Erklärung`"
 Str(0, 112, 3) = "`schgr`"
 ArtZ(0, 112) = 2
 ArtZ(1, 112) = 1
 Str(1, 112, 0) = "CREATE TABLE `werte_scheingruppen` ("
 Str(1, 112, 1) = " `schgr` decimal(2,0) NOT NULL DEFAULT '0' COMMENT '4239'"
 Str(1, 112, 2) = " `Erklärung` varchar(70) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 112, 3) = "  PRIMARY KEY (`schgr`) USING BTREE"
 Str(1, 112, 4) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Scheingruppen in Turbomed'"
End Function ' FüllStr112

Function FüllStr113()
 Str(0, 113, 0) = "werte_weggeldzonen"
 Str(0, 113, 1) = "`Weggeldzone`"
 Str(0, 113, 2) = "`Zonennr`"
 Str(0, 113, 3) = "`Bereich`"
 Str(0, 113, 4) = "`Weggeldzone`"
 Str(0, 113, 5) = "`Zonennr`"
 ArtZ(0, 113) = 3
 ArtZ(1, 113) = 2
 Str(1, 113, 0) = "CREATE TABLE `werte_weggeldzonen` ("
 Str(1, 113, 1) = " `Weggeldzone` varchar(2) CHARACTER SET latin1 NOT NULL DEFAULT '' COMMENT '3631 (1) Weggeldzone'"
 Str(1, 113, 2) = " `Zonennr` decimal(1,0) DEFAULT NULL COMMENT '3631 (2) Weggeldzonenziffer'"
 Str(1, 113, 3) = " `Bereich` varchar(30) CHARACTER SET latin1 NOT NULL DEFAULT '' COMMENT 'Kilometer-Bereich'"
 Str(1, 113, 4) = "  PRIMARY KEY (`Weggeldzone`) USING BTREE"
 Str(1, 113, 5) = "  KEY `Zonennr` (`Zonennr`)"
 Str(1, 113, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Weggeldzonen'"
End Function ' FüllStr113

Function doEx&(sql$, obtolerant%) ' SQL-Befehl ausführen, Fehler anzeigen
 Dim rAF&, FMeld$
 If obtolerant Then On Error Resume Next Else On Error GoTo fehler
 Call cnz.Execute(sql, rAF)
 lErrNr = Err.Number
 FMeld = "Err.Nr " & lErrNr & ", rAf: " & rAF & " bei " & sql
 On Error GoTo fehler
 Debug.Print FMeld
 If obProt Then Print #302, FMeld
 DoEvents
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case Err.Number
 Case -2147467259
  If InStrB(Err.Description, "nicht erzeugen") Then ' 'Kann Tabelle 'testDB1.faxe' nicht erzeugen (Fehler: 150)
   doEx = 150
   Exit Function
  ElseIf InStrB(Err.Description, "is not BASE TABLE") <> 0 Then
   doEx = 151
   Exit Function
  ElseIf InStrB(Err.Description, "MySQL server has gone away") Then
   cnz.Close
   cnz.Open
   Call doEx("use `" & hDBn & "`", 0)
   Resume
  End If
End Select
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doEx/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function ' doEx

Function SplitN&(ByVal q$, Sep$, erg$()) ' da Split() Speicher fraß
 Dim p1&, p2&, Slen&, obExit%, runde&
 On Error GoTo fehler
 If Not IsNull(q) Then
  Slen = Len(Sep)
  For runde = 1 To 2
   p2 = 0
   Do
    p1 = p2
    p2 = InStr(p1 + Slen, q, Sep)
    If p2 = 0 Then p2 = Len(q) + 1: obExit = True
    If p2 <> 0 Then
     If runde = 2 Then
      erg(SplitN) = Mid$(q, p1 + Slen, p2 - p1 - Slen)
     End If
     SplitN = SplitN + 1
    End If
    If obExit Then Exit Do
   Loop
   If runde = 1 Then
    ReDim erg(SplitN - 1)
    SplitN = 0
    obExit = 0
   End If
  Next runde
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SplitN/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' aufSplit

Function doMach_quelle(DBn$, Server$) ' Datenbankname
 Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$
 Dim i&, p1&, p2&, p3&, CLen&, CLen1&, obLT%
 Dim index$()
 On Error Resume Next
 hDBn = DBn
 Open App.path & "\MachDB.bas_prot.txt" For Output As #302
 obProt = (Err.Number = 0)
 On Error GoTo fehler
 cnzCStr = "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=" & Server & ";uid=mysql;pwd=97a5o6;"
 Set cnz = Nothing
 cnz.Open cnzCStr
 Call doEx("create database if not exists `" & DBn & "` character set latin1 collate latin1_german2_ci;", 0)
 Call doEx("grant all privileges on `" & DBn & "`.* to 'praxis'@'%' identified by 'sonne' with grant option", 0)
 Call doEx("grant all privileges on `" & DBn & "`.* to 'praxis'@'localhost' identified by 'sonne' with grant option", 0)
 Call doEx("use `" & DBn & "`", 0)
 Call doEx("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", 0)
 FüllStr0
 FüllStr1
 FüllStr2
 FüllStr3
 FüllStr4
 FüllStr5
 FüllStr6
 FüllStr7
 FüllStr8
 FüllStr9
 FüllStr10
 FüllStr11
 FüllStr12
 FüllStr13
 FüllStr14
 FüllStr15
 FüllStr16
 FüllStr17
 FüllStr18
 FüllStr19
 FüllStr20
 FüllStr21
 FüllStr22
 FüllStr23
 FüllStr24
 FüllStr25
 FüllStr26
 FüllStr27
 FüllStr28
 FüllStr29
 FüllStr30
 FüllStr31
 FüllStr32
 FüllStr33
 FüllStr34
 FüllStr35
 FüllStr36
 FüllStr37
 FüllStr38
 FüllStr39
 FüllStr40
 FüllStr41
 FüllStr42
 FüllStr43
 FüllStr44
 FüllStr45
 FüllStr46
 FüllStr47
 FüllStr48
 FüllStr49
 FüllStr50
 FüllStr51
 FüllStr52
 FüllStr53
 FüllStr54
 FüllStr55
 FüllStr56
 FüllStr57
 FüllStr58
 FüllStr59
 FüllStr60
 FüllStr61
 FüllStr62
 FüllStr63
 FüllStr64
 FüllStr65
 FüllStr66
 FüllStr67
 FüllStr68
 FüllStr69
 FüllStr70
 FüllStr71
 FüllStr72
 FüllStr73
 FüllStr74
 FüllStr75
 FüllStr76
 FüllStr77
 FüllStr78
 FüllStr79
 FüllStr80
 FüllStr81
 FüllStr82
 FüllStr83
 FüllStr84
 FüllStr85
 FüllStr86
 FüllStr87
 FüllStr88
 FüllStr89
 FüllStr90
 FüllStr91
 FüllStr92
 FüllStr93
 FüllStr94
 FüllStr95
 FüllStr96
 FüllStr97
 FüllStr98
 FüllStr99
 FüllStr100
 FüllStr101
 FüllStr102
 FüllStr103
 FüllStr104
 FüllStr105
 FüllStr106
 FüllStr107
 FüllStr108
 FüllStr109
 FüllStr110
 FüllStr111
 FüllStr112
 FüllStr113
 Call doEx("SET FOREIGN_KEY_CHECKS = 0", 0)

 Dim j&, ZZ&, Tbl$, sql As New CString
 For i = 0 To 113
  If InStr(Str(1, i, 0), "CREATE TABLE") <> 0 Then
   Tbl = Str(0, i, 0)
   ZZ = ArtZ(0, i) + ArtZ(1, i)
   sql = "CREATE TABLE IF NOT EXISTS `" & Tbl & "` (" & vbLf
   For j = 1 To ZZ
    sql.Append Str(1, i, j)
    If j < ZZ Then sql.Append "," & vbLf
   Next j
   ZZ = ZZ + ArtZ(2, i) + 1
   sql.Append vbLf & ")"
   sql.Append Str(1, i, ZZ)
   FNr = doEx(sql.Value, 0)
   Do
    Set rsc = Nothing
    rsc.Open "show create table `" & Tbl & "`", cnz, adOpenStatic, adLockReadOnly
    sct = rsc.Fields(1)
    If InStrB(sct, "CREATE ALGORITHM") = 1 Then
     FNr = doEx("drop view `" & Tbl & "`", 0)
     FNr = doEx(sql.Value, 0)
    Else
     Exit Do
    End If
   Loop
   If InStrB(sct, Str(1, i, ZZ)) = 0 Then
    Call doEx("alter table `" & Tbl & "`" & Str(1, i, ZZ), 0)
   End If
   TMt.Clear
   SplitN sct, vbLf, Spli
   For j = 1 To ArtZ(0, i) ' Tabellenfelder
    Dim k&, enthalten%, genau%, Posi$
    enthalten = 0
    genau = 0
    k = 0
    Set rsc = Nothing
    rsc.Open "show columns from `" & Tbl & "` where field = '" & Mid$(Str(0, i, j), 2, Len(Str(0, i, j)) - 2) & "'", cnz, adOpenStatic, adLockReadOnly
    enthalten = Not rsc.BOF
    If enthalten Then
     genau = (InStrB(sct, Str(1, i, j)) <> 0)
     If Not genau Then
      CLen = -1 ' Column-Length nicht kürzen
      obLT = (InStrB(sct, Str(0, i, j) & " longtext") <> 0)
      If Not obLT Then
       p1 = InStr(sct, "(")
       p2 = InStr(p1, sct, Str(0, i, j)) 'zCat.Tables(Tbl).Columns(k).Name & "`")
       If p2 = 0 Then p2 = InStr(p1, LCase$(sct), LCase(Str(0, i, j)))
       p1 = InStr(p2, sct, "(")
       p3 = InStr(p2, sct, ",")
       If p3 = 0 Then p3 = InStr(p2, sct, vbLf & ")")
       If p1 <> 0 And p1 < p3 Then
        p2 = InStr(p1, sct, ")")
        CLen = Mid(sct, p1 + 1, p2 - p1 - 1)
       End If
      End If
     End If
    End If
    If Not enthalten Or Not genau Then
     If j = 1 Then
      Posi = " FIRST,"
     Else
      Posi = " AFTER " & Str(1, i, j - 1) & ","
     End If
     If Not enthalten Then
      TMt.AppVar (Array(" add ", Str(1, i, j), Posi))
     ElseIf Not genau Then
      If CLen <> -1 Or obLT Then
       p1 = InStr(Str(1, i, j), "(")
       If p1 <> 0 Then
        p2 = InStr(p1, Str(1, i, j), ")")
        If p2 <> 0 Then
         CLen1 = Mid(Str(1, i, j), p1 + 1, p2 - p1 - 1)
         If obLT Then
          Str(1, i, j).Replace "varchar(" & CLen1 & ")", "longtext"
         ElseIf CLen1 < CLen Then
          Str(1, i, j).Replace "(" & CLen1 & ")", "(" & CLen & ")"
         End If
         genau = (InStrB(sct, Str(1, i, j)) <> 0)
        End If
       End If
      End If
      If Not genau Then
       TMt.AppVar (Array(" modify ", Str(1, i, j), Posi))
      End If
     End If
    End If
   Next j
   For j = ArtZ(0, i) + 1 To ArtZ(0, i) + ArtZ(1, i) ' Indices
    If InStrB(sct, Str(1, i, j)) = 0 Then
     If InStrB(Str(1, i, j).Value, "PRIMARY") <> 0 Then
      If InStrB(sct, "PRIMARY KEY (") <> 0 Then
       TMt.Append (" DROP PRIMARY KEY,")
      End If
     Else
      If InStrB(sct, "KEY " & Str(0, i, j).Value) <> 0 Then
       TMt.AppVar Array(" DROP KEY ", Str(0, i, j), ",")
      End If
     End If
     TMt.AppVar Array(" add ", Str(1, i, j), ",")
    End If
   Next j
   If TMt.Length <> 0 Then
    TMt.Cut (TMt.Length - 1)
    Call doEx("Alter Table `" & Tbl & "` " & TMt.Value, -1)
   End If
   For j = ArtZ(0, i) + ArtZ(1, i) + 1 To ZZ - 1 'Constraints
    If InStrB(sct, Str(1, i, j)) = 0 Then
     If InStrB(sct, "CONSTRAINT " & Str(0, i, j)) <> 0 Then
      Call doEx("ALTER TABLE `" & Tbl & "` DROP FOREIGN KEY " & Str(0, i, j), 0)
     End If
     Call doEx("ALTER TABLE `" & Tbl & "` ADD" & Str(1, i, j), 0)
    End If
   Next j
  End If ' InStr(Str(1, i, 0), "CREATE TABLE") <> 0 Then
 Next i
 Dim runde%
 For runde = 0 To 4
  For i = 0 To 113
   If InStrB(Str(1, i, 0), "DEFINER VIEW") <> 0 Then
    Dim obCr%
    obCr = 0
    Set rsc = Nothing
    rsc.Open "show tables from `" & DBn & "` where `tables_in_" & DBn & "` = """ & Str(0, i, 0) & """", cnz, adOpenStatic, adLockReadOnly
    If rsc.BOF Then
     obCr = True
    Else
     Set rsc = Nothing
     rsc.Open "show create table `" & Str(0, i, 0) & "`", cnz, adOpenStatic, adLockReadOnly
     If rsc.Fields(1) <> Str(1, i, 0) Then
      Call doEx("DROP TABLE IF EXISTS `" & Str(0, i, 0) & "`", 0)
      Call doEx("DROP VIEW IF EXISTS `" & Str(0, i, 0) & "`", 0)
      obCr = True
     End If
    End If
    If obCr Then
     Call doEx(Str(1, i, 0).Value, True)
    End If
   End If
  Next i
 Next runde
 Call doEx("set FOREIGN_KEY_CHECKS = 1", 0)
 If obProt Then Close #302
 MsgBox "Fertig mit doMach_testDB!"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_testDB/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_testDB

