'Bauanleitung für eine Datenbank wie `//linux/quelle1` vom 15.11.09 14:26:50
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 126, 212) As new CString, ArtZ&(3, 126)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "DMPInkonsistenzen"
 Str(1, 0, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `DMPInkonsistenzen` AS select `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr` from (`_lfaelle` `l` left join `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`)))) group by `f`.`Pat_ID`"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "__fuerlmp"
 Str(1, 1, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `__fuerlmp` AS select `mp`.`Pat_ID` AS `pat_id`,`mp`.`ZeitPunkt` AS `zeitpunkt`,`mp`.`MPNr` AS `mpnr` from `medplan` `mp` group by `mp`.`Pat_ID`,`mp`.`ZeitPunkt`,`mp`.`MPNr` order by `mp`.`Pat_ID`,`mp`.`ZeitPunkt` desc,`mp`.`MPNr` desc"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "__kontakttage"
 Str(1, 2, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `__kontakttage` AS select `e`.`Pat_ID` AS `pat_id`,`e`.`ZeitPunkt` AS `zeitpunkt` from `eintraege` `e` where ((`e`.`ZeitPunkt` between concat(year((now() - interval 14 day)),'-',((((month((now() - interval 14 day)) - 1) DIV 3) * 3) + 1),'-01') and concat((year((now() - interval 14 day)) + round((((((month((now() - interval 14 day)) - 1) DIV 3) * 3) + 4) / 12),0)),'-',(((((month((now() - interval 14 day)) - 1) DIV 3) * 3) + 4) % 12),'-01')) and (`e`.`Art` in ('notiz','telef','ni','gstel','gs','rz','ep','bga','tk','APK','wr','ga','tst','cr','ke','hz','mh','ag','ph','pq','er','ds','st','eb','fa','bz','rp','uzu','hypo','colo','aug','beweg','pros','impf','gyn','caro','beruf','ap','mu','rauch','alko','fams','schula','ass','kra','proc','au','GPD','ba','ARCHIE2','gewicht','gewi','rrvgl','bzvgl','bzm','bztp" & _ 
  "','bks','anal','andm','usal','usdm','doppler','duplex','sono','sd','UKG','Größe','HbA1c','hyper','fuß','keto','wv','ulc','kv','debr','EKG','LZRR','Lufu','lactoset','trop','temp','oGTT','gpt','bmi','urin','taille','hüfte','puls','GDT','bef'))) group by `e`.`Pat_ID`,cast(`e`.`ZeitPunkt` as date)"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "__lfaelle"
 Str(1, 3, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `__lfaelle` AS select `f`.`BhFB` AS `mbhfb`,`f`.`BhFE1` AS `bhfe1`,`f`.`Pat_ID` AS `pid` from `faelle` `f` order by `f`.`Pat_ID`,`f`.`BhFB`"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "_f1"
 Str(1, 4, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_f1` AS select `faelle`.`BhFB` AS `bhfb`,`faelle`.`Pat_ID` AS `pat_id` from `faelle` order by `faelle`.`Pat_ID`,`faelle`.`BhFB` desc,`faelle`.`SchGr`"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "_faellenachschgr"
 Str(1, 5, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_faellenachschgr` AS select `faelle`.`FID` AS `FID`,`faelle`.`Pat_ID` AS `Pat_ID`,`faelle`.`Quartal` AS `Quartal`,`faelle`.`Nachname` AS `Nachname`,`faelle`.`Vorname` AS `Vorname`,`faelle`.`lfdnr` AS `lfdnr`,`faelle`.`TMFNr` AS `TMFNr`,`faelle`.`VKNr` AS `VKNr`,`faelle`.`BhFB` AS `BhFB`,`faelle`.`BhFE1` AS `BhFE1`,`faelle`.`BhFE2` AS `BhFE2`,`faelle`.`f4202` AS `f4202`,`faelle`.`ausgst` AS `ausgst`,`faelle`.`KtrAbrB` AS `KtrAbrB`,`faelle`.`AbrAr` AS `AbrAr`,`faelle`.`lVorl` AS `lVorl`,`faelle`.`IK` AS `IK`,`faelle`.`KVKs` AS `KVKs`,`faelle`.`KVKserg` AS `KVKserg`,`faelle`.`Kasse` AS `Kasse`,`faelle`.`GebOr` AS `GebOr`,`faelle`.`AbrGb` AS `AbrGb`,`faelle`.`PersKreis` AS `PersKreis`,`faelle`.`SKtZusatz` AS `SKtZusatz`,`faelle`.`f4206` AS `f4206`,`faelle`.`ÜwText` AS `ÜwText`,`faelle`.`f4210` AS `f42" & _ 
  "10`,`faelle`.`AkfHAH` AS `AkfHAH`,`faelle`.`AkfAB0` AS `AkfAB0`,`faelle`.`AkfAK` AS `AkfAK`,`faelle`.`statNuller` AS `statNuller`,`faelle`.`ÜbwV` AS `ÜbwV`,`faelle`.`AndÜw` AS `AndÜw`,`faelle`.`Übw` AS `Übw`,`faelle`.`ÜbwLANR` AS `ÜbwLANR`,`faelle`.`ÜWZiel` AS `ÜWZiel`,`faelle`.`ÜWNNr` AS `ÜWNNr`,`faelle`.`ÜWNaN` AS `ÜWNaN`,`faelle`.`ÜWTit` AS `ÜWTit`,`faelle`.`ÜWVor` AS `ÜWVor`,`faelle`.`ÜWVsw` AS `ÜWVsw`,`faelle`.`üwvid` AS `üwvid`,`faelle`.`statKlasse` AS `statKlasse`,`faelle`.`f4237` AS `f4237`,`faelle`.`statBehTage` AS `statBehTage`,`faelle`.`SchGr` AS `SchGr`,`faelle`.`Weiterbeh` AS `Weiterbeh`,`faelle`.`PGeb` AS `PGeb`,`faelle`.`PGebErg` AS `PGebErg`,`faelle`.`Mahnfrist` AS `Mahnfrist`,`faelle`.`GOÄKatNr` AS `GOÄKatNr`,`faelle`.`GOÄKatName` AS `GOÄKatName`,`faelle`.`abrArzt` AS `abrArzt`,`faelle`.`privVers` AS `privVers`,`faelle`.`AdNam` AS `AdNam`,`faelle`.`AdStr` AS `AdStr`,`fae" & _ 
  "lle`.`AdPlz` AS `AdPlz`,`faelle`.`AdOrt` AS `AdOrt`,`faelle`.`BhFE` AS `BhFE`,`faelle`.`s8000` AS `s8000`,`faelle`.`s8100` AS `s8100`,`faelle`.`AktZeit` AS `AktZeit`,`faelle`.`Fanf` AS `Fanf`,`faelle`.`altQuart` AS `altQuart`,`faelle`.`QAnf` AS `QAnf`,`faelle`.`QEnd` AS `QEnd`,`faelle`.`QS` AS `QS`,`faelle`.`QT` AS `QT`,`faelle`.`TherArt` AS `TherArt`,`faelle`.`StByte` AS `StByte`,`faelle`.`absPos` AS `absPos` from `faelle` order by `faelle`.`SchGr`"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "_fuerlmp"
 Str(1, 6, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_fuerlmp` AS select `__fuerlmp`.`pat_id` AS `pat_id`,`__fuerlmp`.`mpnr` AS `mpnr` from `__fuerlmp` group by `__fuerlmp`.`pat_id`"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "_kontakttage"
 Str(0, 7, 1) = "`pat_id`"
 Str(0, 7, 2) = "`zeitpunkt`"
 ArtZ(0, 7) = 2
 Str(1, 7, 0) = "CREATE TABLE `_kontakttage` ("
 Str(1, 7, 1) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 7, 2) = " `zeitpunkt` datetime DEFAULT NULL"
 Str(1, 7, 3) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "_kontaktzahl"
 Str(1, 8, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_kontaktzahl` AS select count(0) AS `ct`,`__kontakttage`.`pat_id` AS `pat_id` from `__kontakttage` group by `__kontakttage`.`pat_id`"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "_lfaelle"
 Str(1, 9, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_lfaelle` AS select `__lfaelle`.`mbhfb` AS `mbhfb`,`__lfaelle`.`bhfe1` AS `bhfe1`,`__lfaelle`.`pid` AS `pid` from `__lfaelle` group by `__lfaelle`.`pid` desc"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "aktf"
 Str(1, 10, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktf` AS select `faelle`.`Pat_ID` AS `pat_id`,`faelle`.`FID` AS `fid`,`faelle`.`SchGr` AS `schgr`,`faelle`.`VKNr` AS `vknr` from `faelle` where ((`faelle`.`SchGr` <> '90') and (`faelle`.`Quartal` = (select concat((((month((now() - interval 14 day)) - 1) DIV 3) + 1),year((now() - interval 14 day))) AS `lq`))) order by `faelle`.`Pat_ID`,`faelle`.`FID` desc,`faelle`.`SchGr`"
End Sub ' FüllStr10

Sub FüllStr11()
 Str(0, 11, 0) = "aktfaelle"
 Str(1, 11, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktfaelle` AS select `f`.`Pat_ID` AS `pid`,`n`.`Notiz` AS `notiz`,`stru`.`Leistung` AS `stru`,`chron`.`Leistung` AS `chron`,`kt`.`ct` AS `kt`,`ebm`.`Leistung` AS `verspau`,`d`.`ICD` AS `icd`,`f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`f4206` AS `f4206`,`f`.`ÜwText` AS" & _ 
  " `ÜwText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`ÜbwV` AS `ÜbwV`,`f`.`AndÜw` AS `AndÜw`,`f`.`Übw` AS `Übw`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8100` AS `s8100`,`f`.`Ak" & _ 
  "tZeit` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`TherArt` AS `TherArt`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos`,`k`.`ID` AS `id`,`k`.`VK` AS `vk`,`k`.`Name` AS `kname`,`k`.`Kateg` AS `kateg`,`k`.`AnzahlIK` AS `anzahlik`,`k`.`AnzahlKTUG` AS `anzahlktug`,`k`.`GültigVon` AS `gültigvon`,`k`.`GültigBis` AS `gültigbis`,`k`.`GO` AS `go`,`k`.`Kurzname` AS `kurzname` from (((((((`faelle` `f` left join `kassenliste` `k` on(((`f`.`VKNr` = `k`.`VK`) and (`f`.`IK` = `k`.`IK`)))) left join `diagnosen` `d` on(((`f`.`Pat_ID` = `d`.`Pat_id`) and (((`d`.`ICD` like 'E1%') and (not((`d`.`ICD` like 'E16%'))) and (not((`d`.`ICD` like 'E15%')))) or (`d`.`ICD` = 'O24.4')) and (`d`.`DiagSicherheit` <> 'A') and ((`d`.`obDauer` = 1) or (`d`.`FID` = `f`.`FID`))))) left join `leistungen` `ebm` on(((`f`.`FID" & _ 
  "` = `ebm`.`FID`) and ((`ebm`.`Leistung` like '031%') or (`ebm`.`Leistung` like '01210'))))) left join `leistungen` `chron` on(((`f`.`FID` = `chron`.`FID`) and (`chron`.`Leistung` = '03212')))) left join `leistungen` `stru` on(((`f`.`FID` = `stru`.`FID`) and (`stru`.`Leistung` like '973%')))) left join `_kontaktzahl` `kt` on((`kt`.`pat_id` = `f`.`Pat_ID`))) left join `namen` `n` on((`n`.`Pat_ID` = `f`.`Pat_ID`))) where ((`f`.`SchGr` <> '90') and (`f`.`Quartal` = (select concat((((month((now() - interval 14 day)) - 1) DIV 3) + 1),year((now() - interval 14 day))) AS `lq`))) group by `f`.`FID` order by `f`.`Pat_ID`,`f`.`SchGr`,`d`.`ICD`"
End Sub ' FüllStr11

Sub FüllStr12()
 Str(0, 12, 0) = "aktfaellev"
 Str(1, 12, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktfaellev` AS select `aktfaelle`.`pid` AS `pid`,`aktfaelle`.`notiz` AS `notiz`,`aktfaelle`.`stru` AS `stru`,`aktfaelle`.`chron` AS `chron`,`aktfaelle`.`kt` AS `kt`,`aktfaelle`.`verspau` AS `verspau`,`aktfaelle`.`icd` AS `icd`,`aktfaelle`.`FID` AS `FID`,`aktfaelle`.`Pat_ID` AS `Pat_ID`,`aktfaelle`.`Quartal` AS `Quartal`,`aktfaelle`.`Nachname` AS `Nachname`,`aktfaelle`.`Vorname` AS `Vorname`,`aktfaelle`.`lfdnr` AS `lfdnr`,`aktfaelle`.`TMFNr` AS `TMFNr`,`aktfaelle`.`VKNr` AS `VKNr`,`aktfaelle`.`BhFB` AS `BhFB`,`aktfaelle`.`BhFE1` AS `BhFE1`,`aktfaelle`.`BhFE2` AS `BhFE2`,`aktfaelle`.`f4202` AS `f4202`,`aktfaelle`.`ausgst` AS `ausgst`,`aktfaelle`.`KtrAbrB` AS `KtrAbrB`,`aktfaelle`.`AbrAr` AS `AbrAr`,`aktfaelle`.`lVorl` AS `lVorl`,`aktfaelle`.`IK` AS `IK`,`aktfaelle`.`KVKs` AS `KVKs`,`aktfaelle`.`KVK" & _ 
  "serg` AS `KVKserg`,`aktfaelle`.`Kasse` AS `Kasse`,`aktfaelle`.`GebOr` AS `GebOr`,`aktfaelle`.`AbrGb` AS `AbrGb`,`aktfaelle`.`PersKreis` AS `PersKreis`,`aktfaelle`.`SKtZusatz` AS `SKtZusatz`,`aktfaelle`.`f4206` AS `f4206`,`aktfaelle`.`ÜwText` AS `ÜwText`,`aktfaelle`.`f4210` AS `f4210`,`aktfaelle`.`AkfHAH` AS `AkfHAH`,`aktfaelle`.`AkfAB0` AS `AkfAB0`,`aktfaelle`.`AkfAK` AS `AkfAK`,`aktfaelle`.`statNuller` AS `statNuller`,`aktfaelle`.`ÜbwV` AS `ÜbwV`,`aktfaelle`.`AndÜw` AS `AndÜw`,`aktfaelle`.`Übw` AS `Übw`,`aktfaelle`.`ÜbwLANR` AS `ÜbwLANR`,`aktfaelle`.`ÜWZiel` AS `ÜWZiel`,`aktfaelle`.`ÜWNNr` AS `ÜWNNr`,`aktfaelle`.`ÜWNaN` AS `ÜWNaN`,`aktfaelle`.`ÜWTit` AS `ÜWTit`,`aktfaelle`.`ÜWVor` AS `ÜWVor`,`aktfaelle`.`ÜWVsw` AS `ÜWVsw`,`aktfaelle`.`üwvid` AS `üwvid`,`aktfaelle`.`statKlasse` AS `statKlasse`,`aktfaelle`.`f4237` AS `f4237`,`aktfaelle`.`statBehTage` AS `statBehTage`,`aktfaelle`.`SchGr` A" & _ 
  "S `SchGr`,`aktfaelle`.`Weiterbeh` AS `Weiterbeh`,`aktfaelle`.`PGeb` AS `PGeb`,`aktfaelle`.`PGebErg` AS `PGebErg`,`aktfaelle`.`Mahnfrist` AS `Mahnfrist`,`aktfaelle`.`GOÄKatNr` AS `GOÄKatNr`,`aktfaelle`.`GOÄKatName` AS `GOÄKatName`,`aktfaelle`.`abrArzt` AS `abrArzt`,`aktfaelle`.`privVers` AS `privVers`,`aktfaelle`.`AdNam` AS `AdNam`,`aktfaelle`.`AdStr` AS `AdStr`,`aktfaelle`.`AdPlz` AS `AdPlz`,`aktfaelle`.`AdOrt` AS `AdOrt`,`aktfaelle`.`BhFE` AS `BhFE`,`aktfaelle`.`s8000` AS `s8000`,`aktfaelle`.`s8100` AS `s8100`,`aktfaelle`.`AktZeit` AS `AktZeit`,`aktfaelle`.`Fanf` AS `Fanf`,`aktfaelle`.`altQuart` AS `altQuart`,`aktfaelle`.`QAnf` AS `QAnf`,`aktfaelle`.`QEnd` AS `QEnd`,`aktfaelle`.`QS` AS `QS`,`aktfaelle`.`QT` AS `QT`,`aktfaelle`.`TherArt` AS `TherArt`,`aktfaelle`.`StByte` AS `StByte`,`aktfaelle`.`absPos` AS `absPos`,`aktfaelle`.`id` AS `id`,`aktfaelle`.`vk` AS `vk`,`aktfaelle`.`kname` AS " & _ 
  "`kname`,`aktfaelle`.`kateg` AS `kateg`,`aktfaelle`.`anzahlik` AS `anzahlik`,`aktfaelle`.`anzahlktug` AS `anzahlktug`,`aktfaelle`.`gültigvon` AS `gültigvon`,`aktfaelle`.`gültigbis` AS `gültigbis`,`aktfaelle`.`go` AS `go`,`aktfaelle`.`kurzname` AS `kurzname` from `aktfaelle` group by `aktfaelle`.`Pat_ID`"
End Sub ' FüllStr12

Sub FüllStr13()
 Str(0, 13, 0) = "aktfv"
 Str(1, 13, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktfv` AS select `faelle`.`Pat_ID` AS `pat_id`,`faelle`.`FID` AS `fid`,`faelle`.`SchGr` AS `schgr`,`faelle`.`VKNr` AS `vknr` from `faelle` where ((`faelle`.`SchGr` <> '90') and (`faelle`.`Quartal` = (select concat((((month((now() - interval 14 day)) - 1) DIV 3) + 1),year((now() - interval 14 day))) AS `lq`))) group by `faelle`.`Pat_ID` order by `faelle`.`Pat_ID`,`faelle`.`SchGr`"
End Sub ' FüllStr13

Sub FüllStr14()
 Str(0, 14, 0) = "anamnesebogen"
 Str(0, 14, 1) = "`Prim`"
 Str(0, 14, 2) = "`Pat_id`"
 Str(0, 14, 3) = "`Nachname`"
 Str(0, 14, 4) = "`Vorname`"
 Str(0, 14, 5) = "`NVorsatz`"
 Str(0, 14, 6) = "`Titel`"
 Str(0, 14, 7) = "`Anrede`"
 Str(0, 14, 8) = "`GebDat`"
 Str(0, 14, 9) = "`Tkz`"
 Str(0, 14, 10) = "`Versicherungsart`"
 Str(0, 14, 11) = "`Diabetestyp`"
 Str(0, 14, 12) = "`Diabetes seit`"
 Str(0, 14, 13) = "`Tabletten seit`"
 Str(0, 14, 14) = "`Insulin seit`"
 Str(0, 14, 15) = "`Grund für Vorstellung`"
 Str(0, 14, 16) = "`Familienanamnese`"
 Str(0, 14, 17) = "`Größe`"
 Str(0, 14, 18) = "`Gewicht`"
 Str(0, 14, 19) = "`bmi`"
 Str(0, 14, 20) = "`Tendenz`"
 Str(0, 14, 21) = "`DiabetesMedikament 1`"
 Str(0, 14, 22) = "`DiabetesMedikament 1 Menge`"
 Str(0, 14, 23) = "`DiabetesMedikament 2`"
 Str(0, 14, 24) = "`DiabetesMedikament 2 Menge`"
 Str(0, 14, 25) = "`DiabetesMedikament 3`"
 Str(0, 14, 26) = "`DiabetesMedikament 3 Menge`"
 Str(0, 14, 27) = "`DiabetesMedikament 4`"
 Str(0, 14, 28) = "`DiabetesMedikament 4 Menge`"
 Str(0, 14, 29) = "`Insulinpumpe`"
 Str(0, 14, 30) = "`Insulinpumpe seit`"
 Str(0, 14, 31) = "`Insulinpumpe Marke`"
 Str(0, 14, 32) = "`Broteinheiten gesamt`"
 Str(0, 14, 33) = "`Broteinheiten früh`"
 Str(0, 14, 34) = "`Broteinheiten ZM früh`"
 Str(0, 14, 35) = "`Broteinheiten mittags`"
 Str(0, 14, 36) = "`Broteinheiten nachmittags`"
 Str(0, 14, 37) = "`Broteinheiten abends`"
 Str(0, 14, 38) = "`Broteinheiten nachts`"
 Str(0, 14, 39) = "`Essenszeit früh`"
 Str(0, 14, 40) = "`Essenszeit vormittags`"
 Str(0, 14, 41) = "`Essenszeit mittags`"
 Str(0, 14, 42) = "`Essenszeit nachmittags`"
 Str(0, 14, 43) = "`Essenszeit abends`"
 Str(0, 14, 44) = "`Essenszeit spät`"
 Str(0, 14, 45) = "`Spritz-Eß-Abstand früh`"
 Str(0, 14, 46) = "`Spritz-Eß-Abstand mittags`"
 Str(0, 14, 47) = "`Spritz-Eß-Abstand abends`"
 Str(0, 14, 48) = "`Spritzstelle früh`"
 Str(0, 14, 49) = "`Spritzstelle mittags`"
 Str(0, 14, 50) = "`Spritzstelle abends`"
 Str(0, 14, 51) = "`Spritzstelle nachts`"
 Str(0, 14, 52) = "`Jahr letzte Diabetesschulung`"
 Str(0, 14, 53) = "`Ort Schulung`"
 Str(0, 14, 54) = "`letztes HbA1c`"
 Str(0, 14, 55) = "`gemessen am`"
 Str(0, 14, 56) = "`vorherige Werte`"
 Str(0, 14, 57) = "`BZMessungen selbst`"
 Str(0, 14, 58) = "`Gerät`"
 Str(0, 14, 59) = "`BZMessungen pW`"
 Str(0, 14, 60) = "`BZMessungen pW ndE`"
 Str(0, 14, 61) = "`BZMessungen p W nachts`"
 Str(0, 14, 62) = "`Aufschreiben`"
 Str(0, 14, 63) = "`BZWerte v d Essen`"
 Str(0, 14, 64) = "`BZWerte n d Essen`"
 Str(0, 14, 65) = "`UZ Tageszeit`"
 Str(0, 14, 66) = "`Unterzucker pM`"
 Str(0, 14, 67) = "`UZ rechtzeitig`"
 Str(0, 14, 68) = "`Fremde Hilfe pa`"
 Str(0, 14, 69) = "`Bewußtlos pa`"
 Str(0, 14, 70) = "`Keto pa`"
 Str(0, 14, 71) = "`BZgr300 pM`"
 Str(0, 14, 72) = "`Bluthochdruck`"
 Str(0, 14, 73) = "`BHD seit`"
 Str(0, 14, 74) = "`BHD beh mit`"
 Str(0, 14, 75) = "`Blutdruckwerte`"
 Str(0, 14, 76) = "`BDselbst`"
 Str(0, 14, 77) = "`Schwanger`"
 Str(0, 14, 78) = "`Schwanger seit`"
 Str(0, 14, 79) = "`Augensp zuletzt`"
 Str(0, 14, 80) = "`Augensp Befund`"
 Str(0, 14, 81) = "`Netzhaut gelasert`"
 Str(0, 14, 82) = "`Sehminderung unbehebbar`"
 Str(0, 14, 83) = "`Diabet Nierenschaden`"
 Str(0, 14, 84) = "`Albumin zuletzt`"
 Str(0, 14, 85) = "`erhöht?`"
 Str(0, 14, 86) = "`Dialyse`"
 Str(0, 14, 87) = "`Dialyse seit`"
 Str(0, 14, 88) = "`andere Nierenerkrankung`"
 Str(0, 14, 89) = "`Herzkrankheit`"
 Str(0, 14, 90) = "`Angina pectoris`"
 Str(0, 14, 91) = "`Herzinfarkt`"
 Str(0, 14, 92) = "`Herzinfarkt wann`"
 Str(0, 14, 93) = "`PTCA oder Stent`"
 Str(0, 14, 94) = "`Bypass kardial`"
 Str(0, 14, 95) = "`Bypass wann`"
 Str(0, 14, 96) = "`Herzschwäche`"
 Str(0, 14, 97) = "`Herzkrankheit Beschreibung`"
 Str(0, 14, 98) = "`Hirndurchblutungsstörung`"
 Str(0, 14, 99) = "`Schlaganfall`"
 Str(0, 14, 100) = "`Beindurchblutungsstörung`"
 Str(0, 14, 101) = "`Schaufensterkrankheit`"
 Str(0, 14, 102) = "`Bypaß peripher`"
 Str(0, 14, 103) = "`Geschwür`"
 Str(0, 14, 104) = "`Amputation`"
 Str(0, 14, 105) = "`pAVK Beschreibung`"
 Str(0, 14, 106) = "`Ameisenlaufen`"
 Str(0, 14, 107) = "`Ameisen Ausmaß`"
 Str(0, 14, 108) = "`Druckstellen`"
 Str(0, 14, 109) = "`Verformungen`"
 Str(0, 14, 110) = "`Verformungen Beschreibung`"
 Str(0, 14, 111) = "`Fußpflege`"
 Str(0, 14, 112) = "`Podologie`"
 Str(0, 14, 113) = "`Einlagen`"
 Str(0, 14, 114) = "`Neue Fußkomplikationen`"
 Str(0, 14, 115) = "`Entleerungsstörungen Magen`"
 Str(0, 14, 116) = "`Entleerungsstörungen Harnblase`"
 Str(0, 14, 117) = "`Schwindel Aufstehen`"
 Str(0, 14, 118) = "`Folgeerkrankungen Haut`"
 Str(0, 14, 119) = "`Bewegungseinschränkungen`"
 Str(0, 14, 120) = "`Sexualstörung`"
 Str(0, 14, 121) = "`Sexualstörung seit`"
 Str(0, 14, 122) = "`Weitere Anamnese`"
 Str(0, 14, 123) = "`Alkohol`"
 Str(0, 14, 124) = "`Tabak`"
 Str(0, 14, 125) = "`TabakEx`"
 Str(0, 14, 126) = "`TabakBis`"
 Str(0, 14, 127) = "`TabakAkt`"
 Str(0, 14, 128) = "`TabakMenge`"
 Str(0, 14, 129) = "`Weitere Medikation`"
 Str(0, 14, 130) = "`Liphypertrophien Abdomen`"
 Str(0, 14, 131) = "`Liphypertrophien Beine`"
 Str(0, 14, 132) = "`Liphypertrophien Arme`"
 Str(0, 14, 133) = "`Beinbefund`"
 Str(0, 14, 134) = "`Hyperkeratosen`"
 Str(0, 14, 135) = "`Ulcera`"
 Str(0, 14, 136) = "`Kraft Zehenheber`"
 Str(0, 14, 137) = "`Kraft Zehenbeuger`"
 Str(0, 14, 138) = "`Kraft Knie`"
 Str(0, 14, 139) = "`ASR`"
 Str(0, 14, 140) = "`PSR`"
 Str(0, 14, 141) = "`Oberflächensensibilität`"
 Str(0, 14, 142) = "`Monofilamenttest`"
 Str(0, 14, 143) = "`Kalt-Warm`"
 Str(0, 14, 144) = "`Vibration IK`"
 Str(0, 14, 145) = "`Vibration Großzehe`"
 Str(0, 14, 146) = "`Puls Leiste`"
 Str(0, 14, 147) = "`Puls Kniekehle`"
 Str(0, 14, 148) = "`Puls Atp`"
 Str(0, 14, 149) = "`Puls Adp`"
 Str(0, 14, 150) = "`RR`"
 Str(0, 14, 151) = "`RRTurboMed`"
 Str(0, 14, 152) = "`Herz`"
 Str(0, 14, 153) = "`Lunge`"
 Str(0, 14, 154) = "`Bauch`"
 Str(0, 14, 155) = "`WS`"
 Str(0, 14, 156) = "`NL`"
 Str(0, 14, 157) = "`SD`"
 Str(0, 14, 158) = "`Carotiden`"
 Str(0, 14, 159) = "`NNH`"
 Str(0, 14, 160) = "`Zähne`"
 Str(0, 14, 161) = "`Mundhöhle`"
 Str(0, 14, 162) = "`LK`"
 Str(0, 14, 163) = "`BeinödVen`"
 Str(0, 14, 164) = "`Neuro sonst`"
 Str(0, 14, 165) = "`Weitere Befunde`"
 Str(0, 14, 166) = "`Schulung`"
 Str(0, 14, 167) = "`DMP`"
 Str(0, 14, 168) = "`DMSchulz`"
 Str(0, 14, 169) = "`DMSchL`"
 Str(0, 14, 170) = "`RRSchulz`"
 Str(0, 14, 171) = "`DMPhier`"
 Str(0, 14, 172) = "`HANr`"
 Str(0, 14, 173) = "`HANr2`"
 Str(0, 14, 174) = "`letzte Änderung`"
 Str(0, 14, 175) = "`Diagnosen`"
 Str(0, 14, 176) = "`Vorgestellt`"
 Str(0, 14, 177) = "`Versicherung`"
 Str(0, 14, 178) = "`AktZeit`"
 Str(0, 14, 179) = "`Ther1`"
 Str(0, 14, 180) = "`TherAkt`"
 Str(0, 14, 181) = "`obAn1eing`"
 Str(0, 14, 182) = "`obAn2eing`"
 Str(0, 14, 183) = "`obAnAeing`"
 Str(0, 14, 184) = "`obCheck`"
 Str(0, 14, 185) = "`obBZausgew`"
 Str(0, 14, 186) = "`obOSaufgek`"
 Str(0, 14, 187) = "`obPodAufgek`"
 Str(0, 14, 188) = "`obMBlAusgeh`"
 Str(0, 14, 189) = "`obSchulaufgek`"
 Str(0, 14, 190) = "`obDMPaufgekl`"
 Str(0, 14, 191) = "`obMedNetz`"
 Str(0, 14, 192) = "`Hausarzt`"
 Str(0, 14, 193) = "`ob`"
 Str(0, 14, 194) = "`QS`"
 Str(0, 14, 195) = "`QT`"
 Str(0, 14, 196) = "`Prim`"
 Str(0, 14, 197) = "`PrimaryKey`"
 Str(0, 14, 198) = "`Auswahl`"
 Str(0, 14, 199) = "`DMPhier`"
 Str(0, 14, 200) = "`GebDat`"
 Str(0, 14, 201) = "`jlD`"
 Str(0, 14, 202) = "`lÄnd`"
 Str(0, 14, 203) = "`Pat_ID`"
 Str(0, 14, 204) = "`Ther1`"
 Str(0, 14, 205) = "`Vorgestellt`"
 Str(0, 14, 206) = "`HausärzteAnamnesebogen_AccRel`"
 Str(0, 14, 207) = "`KassenlisteAnamnesebogen_AccRel`"
 Str(0, 14, 208) = "`HausärzteAnamnesebogen_AccRel`"
 Str(0, 14, 209) = "`KassenlisteAnamnesebogen_AccRel`"
 Str(0, 14, 210) = "`NamenAnamnesebogen_AccRel`"
 ArtZ(0, 14) = 195
 ArtZ(1, 14) = 12
 ArtZ(2, 14) = 3
 Str(1, 14, 0) = "CREATE TABLE `anamnesebogen` ("
 Str(1, 14, 1) = " `Prim` int(2) unsigned NOT NULL COMMENT 'Primärschlüssel'"
 Str(1, 14, 2) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 14, 3) = " `Nachname` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '-'"
 Str(1, 14, 4) = " `Vorname` varchar(19) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 5) = " `NVorsatz` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 6) = " `Titel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 7) = " `Anrede` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 8) = " `GebDat` datetime DEFAULT NULL COMMENT ', geb.'"
 Str(1, 14, 9) = " `Tkz` tinyint(1) unsigned DEFAULT NULL COMMENT 'Tod-Kennzeichen'"
 Str(1, 14, 10) = " `Versicherungsart` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 11) = " `Diabetestyp` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetes Typ'"
 Str(1, 14, 12) = " `Diabetes seit` varchar(152) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 14, 13) = " `Tabletten seit` varchar(66) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Tabletten seit'"
 Str(1, 14, 14) = " `Insulin seit` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Insulin seit'"
 Str(1, 14, 15) = " `Grund für Vorstellung` varchar(721) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 16) = " `Familienanamnese` varchar(291) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 17) = " `Größe` double DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 18) = " `Gewicht` double DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 19) = " `bmi` decimal(5,1) DEFAULT '0.0'"
 Str(1, 14, 20) = " `Tendenz` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Tendenz'"
 Str(1, 14, 21) = " `DiabetesMedikament 1` varchar(48) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesmedikation:'"
 Str(1, 14, 22) = " `DiabetesMedikament 1 Menge` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 14, 23) = " `DiabetesMedikament 2` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 14, 24) = " `DiabetesMedikament 2 Menge` varchar(126) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 14, 25) = " `DiabetesMedikament 3` varchar(44) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 14, 26) = " `DiabetesMedikament 3 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<'"
 Str(1, 14, 27) = " `DiabetesMedikament 4` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 14, 28) = " `DiabetesMedikament 4 Menge` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,'"
 Str(1, 14, 29) = " `Insulinpumpe` tinyint(1) unsigned DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 30) = " `Insulinpumpe seit` varchar(250) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 14, 31) = " `Insulinpumpe Marke` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Marke:'"
 Str(1, 14, 32) = " `Broteinheiten gesamt` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Broteinheiten:gesamt'"
 Str(1, 14, 33) = " `Broteinheiten früh` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, früh'"
 Str(1, 14, 34) = " `Broteinheiten ZM früh` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zwischenmahlzeit vormittags'"
 Str(1, 14, 35) = " `Broteinheiten mittags` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 14, 36) = " `Broteinheiten nachmittags` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 14, 37) = " `Broteinheiten abends` varchar(13) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 14, 38) = " `Broteinheiten nachts` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 14, 39) = " `Essenszeit früh` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Essenszeiten:früh'"
 Str(1, 14, 40) = " `Essenszeit vormittags` varchar(34) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vormittags'"
 Str(1, 14, 41) = " `Essenszeit mittags` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 14, 42) = " `Essenszeit nachmittags` varchar(24) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachmittags'"
 Str(1, 14, 43) = " `Essenszeit abends` varchar(18) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 14, 44) = " `Essenszeit spät` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, spät'"
 Str(1, 14, 45) = " `Spritz-Eß-Abstand früh` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritz-Eß-Abstand:früh'"
 Str(1, 14, 46) = " `Spritz-Eß-Abstand mittags` varchar(27) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 14, 47) = " `Spritz-Eß-Abstand abends` varchar(62) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 14, 48) = " `Spritzstelle früh` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Spritzstellen:früh'"
 Str(1, 14, 49) = " `Spritzstelle mittags` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, mittags'"
 Str(1, 14, 50) = " `Spritzstelle abends` varchar(139) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, abends'"
 Str(1, 14, 51) = " `Spritzstelle nachts` varchar(141) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts'"
 Str(1, 14, 52) = " `Jahr letzte Diabetesschulung` varchar(129) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Diabetesschulung:'"
 Str(1, 14, 53) = " `Ort Schulung` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<in'"
 Str(1, 14, 54) = " `letztes HbA1c` varchar(29) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letztes HbA1c:'"
 Str(1, 14, 55) = " `gemessen am` datetime DEFAULT NULL COMMENT '<, gemessen'"
 Str(1, 14, 56) = " `vorherige Werte` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, vorher:'"
 Str(1, 14, 57) = " `BZMessungen selbst` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckermessung:Selbstmessung?'"
 Str(1, 14, 58) = " `Gerät` varchar(63) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<:'"
 Str(1, 14, 59) = " `BZMessungen pW` varchar(112) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl d.Messungen pro Woche:'"
 Str(1, 14, 60) = " `BZMessungen pW ndE` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, davon nach dem Essen:'"
 Str(1, 14, 61) = " `BZMessungen p W nachts` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nachts:'"
 Str(1, 14, 62) = " `Aufschreiben` varchar(93) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Dokumentation:'"
 Str(1, 14, 63) = " `BZWerte v d Essen` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutzuckerwerte vor dem Essen:'"
 Str(1, 14, 64) = " `BZWerte n d Essen` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, nach dem Essen:'"
 Str(1, 14, 65) = " `UZ Tageszeit` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Unterzucker:Bevorzugte Tages-/Uhrzeit'"
 Str(1, 14, 66) = " `Unterzucker pM` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Zahl der schweren (<50 mg/dl) pro Monat:'"
 Str(1, 14, 67) = " `UZ rechtzeitig` varchar(101) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, rechtzeitig bemerkt:'"
 Str(1, 14, 68) = " `Fremde Hilfe pa` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, fremde Hilfe deshalb nötig:'"
 Str(1, 14, 69) = " `Bewußtlos pa` varchar(81) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, bewußtlos deshalb:'"
 Str(1, 14, 70) = " `Keto pa` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Zahl der Ketoazidosen pro Jahr:'"
 Str(1, 14, 71) = " `BZgr300 pM` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Zahl der Blutzucker > 300 mg/dl pro Monat:'"
 Str(1, 14, 72) = " `Bluthochdruck` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Bluthochdruck:'"
 Str(1, 14, 73) = " `BHD seit` varchar(84) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit:'"
 Str(1, 14, 74) = " `BHD beh mit` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, behandelt mit:'"
 Str(1, 14, 75) = " `Blutdruckwerte` varchar(186) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckwerte:'"
 Str(1, 14, 76) = " `BDselbst` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruckselbstmessung:'"
 Str(1, 14, 77) = " `Schwanger` varchar(43) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Aktuelle Schwangerschaft:'"
 Str(1, 14, 78) = " `Schwanger seit` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, seit:'"
 Str(1, 14, 79) = " `Augensp zuletzt` varchar(107) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Letzte Augenspiegelung:'"
 Str(1, 14, 80) = " `Augensp Befund` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 14, 81) = " `Netzhaut gelasert` varchar(142) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Netzhaut schon gelasert:'"
 Str(1, 14, 82) = " `Sehminderung unbehebbar` varchar(157) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', mit Brille nicht behebbare Sehminderung:'"
 Str(1, 14, 83) = " `Diabet Nierenschaden` varchar(96) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Diabetischer Nierenschaden:'"
 Str(1, 14, 84) = " `Albumin zuletzt` varchar(73) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', letztes Albumin:'"
 Str(1, 14, 85) = " `erhöht?` varchar(52) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Befund:'"
 Str(1, 14, 86) = " `Dialyse` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 87) = " `Dialyse seit` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 14, 88) = " `andere Nierenerkrankung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', andere Nierenerkrankung:'"
 Str(1, 14, 89) = " `Herzkrankheit` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Herzkrankheit:'"
 Str(1, 14, 90) = " `Angina pectoris` varchar(134) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 91) = " `Herzinfarkt` varchar(122) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 92) = " `Herzinfarkt wann` varchar(104) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 14, 93) = " `PTCA oder Stent` varchar(94) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 94) = " `Bypass kardial` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 95) = " `Bypass wann` varchar(388) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, wann:'"
 Str(1, 14, 96) = " `Herzschwäche` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 97) = " `Herzkrankheit Beschreibung` varchar(213) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung:'"
 Str(1, 14, 98) = " `Hirndurchblutungsstörung` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 99) = " `Schlaganfall` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 100) = " `Beindurchblutungsstörung` varchar(136) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 101) = " `Schaufensterkrankheit` varchar(175) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 102) = " `Bypaß peripher` tinyint(1) unsigned DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 103) = " `Geschwür` varchar(174) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 104) = " `Amputation` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 105) = " `pAVK Beschreibung` varchar(97) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beschreibung der Beinbeschwerden:'"
 Str(1, 14, 106) = " `Ameisenlaufen` varchar(115) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 107) = " `Ameisen Ausmaß` varchar(123) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Ausmaß:'"
 Str(1, 14, 108) = " `Druckstellen` varchar(173) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 109) = " `Verformungen` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 110) = " `Verformungen Beschreibung` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<Beschreibung:'"
 Str(1, 14, 111) = " `Fußpflege` varchar(86) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 112) = " `Podologie` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 113) = " `Einlagen` varchar(91) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', diabetesgerechte orthopädische Einlagen/Schuhe:'"
 Str(1, 14, 114) = " `Neue Fußkomplikationen` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Neue Fußkomplikationen in den letzten 12 Monaten:'"
 Str(1, 14, 115) = " `Entleerungsstörungen Magen` varchar(130) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 116) = " `Entleerungsstörungen Harnblase` varchar(157) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 117) = " `Schwindel Aufstehen` varchar(121) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 118) = " `Folgeerkrankungen Haut` varchar(156) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 119) = " `Bewegungseinschränkungen` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 120) = " `Sexualstörung` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 121) = " `Sexualstörung seit` varchar(149) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<seit'"
 Str(1, 14, 122) = " `Weitere Anamnese` varchar(989) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 123) = " `Alkohol` varchar(148) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 124) = " `Tabak` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 125) = " `TabakEx` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, früher:'"
 Str(1, 14, 126) = " `TabakBis` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, bis:'"
 Str(1, 14, 127) = " `TabakAkt` varchar(55) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, aktuell:'"
 Str(1, 14, 128) = " `TabakMenge` varchar(108) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<:'"
 Str(1, 14, 129) = " `Weitere Medikation` varchar(298) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 130) = " `Liphypertrophien Abdomen` varchar(176) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Liphypertrophien:Abdomen'"
 Str(1, 14, 131) = " `Liphypertrophien Beine` varchar(49) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Beine:'"
 Str(1, 14, 132) = " `Liphypertrophien Arme` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Arme:'"
 Str(1, 14, 133) = " `Beinbefund` varchar(272) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 134) = " `Hyperkeratosen` varchar(207) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 135) = " `Ulcera` varchar(103) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 136) = " `Kraft Zehenheber` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Kraft:Zehenheber'"
 Str(1, 14, 137) = " `Kraft Zehenbeuger` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Zehenbeuger:'"
 Str(1, 14, 138) = " `Kraft Knie` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Knie:'"
 Str(1, 14, 139) = " `ASR` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 140) = " `PSR` varchar(67) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 141) = " `Oberflächensensibilität` varchar(92) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 142) = " `Monofilamenttest` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 143) = " `Kalt-Warm` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Kalt-Warm-Diskrimination:'"
 Str(1, 14, 144) = " `Vibration IK` varchar(31) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Vibrationsempfinden Innenknöchel:'"
 Str(1, 14, 145) = " `Vibration Großzehe` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<, Großzehe:'"
 Str(1, 14, 146) = " `Puls Leiste` varchar(33) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Pulse:Leiste'"
 Str(1, 14, 147) = " `Puls Kniekehle` varchar(26) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Kniekehle:'"
 Str(1, 14, 148) = " `Puls Atp` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Innenknöchel:'"
 Str(1, 14, 149) = " `Puls Adp` varchar(59) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '<,Fußrücken:'"
 Str(1, 14, 150) = " `RR` varchar(245) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Blutdruck:'"
 Str(1, 14, 151) = " `RRTurboMed` varchar(1362) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 152) = " `Herz` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^:'"
 Str(1, 14, 153) = " `Lunge` varchar(79) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 154) = " `Bauch` varchar(106) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Abdomen:'"
 Str(1, 14, 155) = " `WS` varchar(56) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Wirbelsäule:'"
 Str(1, 14, 156) = " `NL` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nierenlager:'"
 Str(1, 14, 157) = " `SD` varchar(88) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Schilddrüse:'"
 Str(1, 14, 158) = " `Carotiden` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Halsschlagadern:'"
 Str(1, 14, 159) = " `NNH` varchar(21) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Nasennebenhöhlen:'"
 Str(1, 14, 160) = " `Zähne` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 161) = " `Mundhöhle` varchar(64) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ',:'"
 Str(1, 14, 162) = " `LK` varchar(58) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Lymphknoten:'"
 Str(1, 14, 163) = " `BeinödVen` varchar(145) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', Beinödeme/ Venenkrankheiten:'"
 Str(1, 14, 164) = " `Neuro sonst` varchar(74) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '^Sonstige neurologische Befunde:'"
 Str(1, 14, 165) = " `Weitere Befunde` varchar(160) COLLATE latin1_german2_ci DEFAULT NULL COMMENT ', weitere Befunde:'"
 Str(1, 14, 166) = " `Schulung` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Schulungsbedarf'"
 Str(1, 14, 167) = " `DMP` varchar(85) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Pat. bei HA im DMP'"
 Str(1, 14, 168) = " `DMSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der DMP-Schulungen hier'"
 Str(1, 14, 169) = " `DMSchL` smallint(6) DEFAULT NULL COMMENT 'Zahl der abgerechneten DMP-Schulungen hier'"
 Str(1, 14, 170) = " `RRSchulz` smallint(6) DEFAULT NULL COMMENT 'Zahl der Hypertonie-Schulungen hier'"
 Str(1, 14, 171) = " `DMPhier` datetime DEFAULT NULL COMMENT 'ob Pat hier im DMP'"
 Str(1, 14, 172) = " `HANr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 14, 173) = " `HANr2` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'mit ""/""'"
 Str(1, 14, 174) = " `letzte Änderung` datetime DEFAULT NULL COMMENT 'Datum der letzten Änderung'"
 Str(1, 14, 175) = " `Diagnosen` varchar(1071) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 176) = " `Vorgestellt` datetime DEFAULT NULL COMMENT 'Erstvorstellung'"
 Str(1, 14, 177) = " `Versicherung` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 14, 178) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 14, 179) = " `Ther1` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, ICT, CSII'"
 Str(1, 14, 180) = " `TherAkt` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diät, OAD, CT, ICT, CSII'"
 Str(1, 14, 181) = " `obAn1eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 1 eingegeben wurde'"
 Str(1, 14, 182) = " `obAn2eing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt S. 2 eingegeben wurde'"
 Str(1, 14, 183) = " `obAnAeing` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Anamneseblatt allgemein eingegeben wurde'"
 Str(1, 14, 184) = " `obCheck` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Checkliste vorliegt'"
 Str(1, 14, 185) = " `obBZausgew` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Blutzuckergerät ausgewechselt'"
 Str(1, 14, 186) = " `obOSaufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über orthopäd Schuhmacher aufgeklärt'"
 Str(1, 14, 187) = " `obPodAufgek` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 14, 188) = " `obMBlAusgeh` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 14, 189) = " `obSchulaufgek` varchar(14) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob über Podologie aufgeklärt'"
 Str(1, 14, 190) = " `obDMPaufgekl` varchar(17) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ob Merkblatt Fußsyndrom ausgehändigt'"
 Str(1, 14, 191) = " `obMedNetz` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob von Med. Netz geschickt'"
 Str(1, 14, 192) = " `Hausarzt` varchar(98) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Hausarzt laut Anamnesebogen'"
 Str(1, 14, 193) = " `ob` tinyint(1) unsigned DEFAULT NULL COMMENT 'für verschiedene Aktionen'"
 Str(1, 14, 194) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 14, 195) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal sortiert von vorgestellt'"
 Str(1, 14, 196) = "  PRIMARY KEY (`Prim`)"
 Str(1, 14, 197) = "  UNIQUE KEY `PrimaryKey` (`Prim`)"
 Str(1, 14, 198) = "  KEY `Auswahl` (`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 14, 199) = "  KEY `DMPhier` (`DMPhier`)"
 Str(1, 14, 200) = "  KEY `GebDat` (`GebDat`)"
 Str(1, 14, 201) = "  KEY `jlD` (`Jahr letzte Diabetesschulung`,`GebDat`)"
 Str(1, 14, 202) = "  KEY `lÄnd` (`letzte Änderung`)"
 Str(1, 14, 203) = "  KEY `Pat_ID` (`Pat_id`)"
 Str(1, 14, 204) = "  KEY `Ther1` (`Ther1`)"
 Str(1, 14, 205) = "  KEY `Vorgestellt` (`Vorgestellt`,`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 14, 206) = "  KEY `HausärzteAnamnesebogen_AccRel` (`HANr`)"
 Str(1, 14, 207) = "  KEY `KassenlisteAnamnesebogen_AccRel` (`Versicherung`)"
 Str(1, 14, 208) = "  CONSTRAINT `HausärzteAnamnesebogen_AccRel` FOREIGN KEY (`HANr`) REFERENCES `hausaerzte` (`KVNr`)"
 Str(1, 14, 209) = "  CONSTRAINT `KassenlisteAnamnesebogen_AccRel` FOREIGN KEY (`Versicherung`) REFERENCES `kassenliste` (`VK`)"
 Str(1, 14, 210) = "  CONSTRAINT `NamenAnamnesebogen_AccRel` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`)"
 Str(1, 14, 211) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr14

Sub FüllStr15()
 Str(0, 15, 0) = "angiozahl"
 Str(0, 15, 1) = "`count(0)`"
 Str(0, 15, 2) = "`quartal`"
 ArtZ(0, 15) = 2
 Str(1, 15, 0) = "CREATE TABLE `angiozahl` ("
 Str(1, 15, 1) = " `count(0)` decimal(15,4) DEFAULT NULL"
 Str(1, 15, 2) = " `quartal` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 15, 3) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr15

Sub FüllStr16()
 Str(0, 16, 0) = "au"
 Str(0, 16, 1) = "`FID`"
 Str(0, 16, 2) = "`Pat_ID`"
 Str(0, 16, 3) = "`ZeitPunkt`"
 Str(0, 16, 4) = "`Beginn`"
 Str(0, 16, 5) = "`Ende`"
 Str(0, 16, 6) = "`ICDs`"
 Str(0, 16, 7) = "`absPos`"
 Str(0, 16, 8) = "`AktZeit`"
 Str(0, 16, 9) = "`StByte`"
 Str(0, 16, 10) = "`Auswahl`"
 Str(0, 16, 11) = "`FälleAU`"
 Str(0, 16, 12) = "`FID`"
 Str(0, 16, 13) = "`NamenAU`"
 Str(0, 16, 14) = "`FälleAU_AccRel`"
 Str(0, 16, 15) = "`NamenAU_AccRel`"
 ArtZ(0, 16) = 9
 ArtZ(1, 16) = 4
 ArtZ(2, 16) = 2
 Str(1, 16, 0) = "CREATE TABLE `au` ("
 Str(1, 16, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 16, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 16, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 16, 4) = " `Beginn` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6285 1. Hälfte'"
 Str(1, 16, 5) = " `Ende` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6285 2. Hälfte'"
 Str(1, 16, 6) = " `ICDs` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6286'"
 Str(1, 16, 7) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 16, 8) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 16, 9) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 16, 10) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Beginn`,`Ende`)"
 Str(1, 16, 11) = "  KEY `FälleAU` (`FID`)"
 Str(1, 16, 12) = "  KEY `FID` (`FID`)"
 Str(1, 16, 13) = "  KEY `NamenAU` (`Pat_ID`)"
 Str(1, 16, 14) = "  CONSTRAINT `FälleAU_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON UPDATE CASCADE"
 Str(1, 16, 15) = "  CONSTRAINT `NamenAU_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 16, 16) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr16

Sub FüllStr17()
 Str(0, 17, 0) = "augenbefunde"
 Str(0, 17, 1) = "`ID`"
 Str(0, 17, 2) = "`Pat_ID`"
 Str(0, 17, 3) = "`Datum`"
 Str(0, 17, 4) = "`Retinopathie`"
 Str(0, 17, 5) = "`Klassifikation`"
 Str(0, 17, 6) = "`Maculopathie`"
 Str(0, 17, 7) = "`Sonstiges`"
 Str(0, 17, 8) = "`Visus re`"
 Str(0, 17, 9) = "`Visus li`"
 Str(0, 17, 10) = "`KontrolleinMonaten`"
 Str(0, 17, 11) = "`Augenarzt`"
 Str(0, 17, 12) = "`DokName`"
 Str(0, 17, 13) = "`DokPfad`"
 Str(0, 17, 14) = "`verglichen`"
 Str(0, 17, 15) = "`ID`"
 Str(0, 17, 16) = "`PrimaryKey`"
 Str(0, 17, 17) = "`Auswahl`"
 Str(0, 17, 18) = "`NamenAugenbefunde`"
 Str(0, 17, 19) = "`Text`"
 Str(0, 17, 20) = "`NamenAugenbefunde_AccRel`"
 ArtZ(0, 17) = 14
 ArtZ(1, 17) = 5
 ArtZ(2, 17) = 1
 Str(1, 17, 0) = "CREATE TABLE `augenbefunde` ("
 Str(1, 17, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 17, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 17, 3) = " `Datum` datetime DEFAULT NULL"
 Str(1, 17, 4) = " `Retinopathie` longtext COLLATE latin1_german2_ci"
 Str(1, 17, 5) = " `Klassifikation` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 6) = " `Maculopathie` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 7) = " `Sonstiges` longtext COLLATE latin1_german2_ci"
 Str(1, 17, 8) = " `Visus re` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 9) = " `Visus li` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 10) = " `KontrolleinMonaten` int(10) DEFAULT NULL"
 Str(1, 17, 11) = " `Augenarzt` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 12) = " `DokName` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 17, 13) = " `DokPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 17, 14) = " `verglichen` datetime DEFAULT NULL"
 Str(1, 17, 15) = "  PRIMARY KEY (`ID`)"
 Str(1, 17, 16) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 17, 17) = "  KEY `Auswahl` (`Pat_ID`,`DokPfad`(255))"
 Str(1, 17, 18) = "  KEY `NamenAugenbefunde` (`Pat_ID`)"
 Str(1, 17, 19) = "  KEY `Text` (`Pat_ID`,`Retinopathie`(255))"
 Str(1, 17, 20) = "  CONSTRAINT `NamenAugenbefunde_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 17, 21) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr17

Sub FüllStr18()
 Str(0, 18, 0) = "bezirke"
 Str(0, 18, 1) = "`id`"
 Str(0, 18, 2) = "`Bezirk`"
 Str(0, 18, 3) = "`Nr`"
 Str(0, 18, 4) = "`MaxNr`"
 Str(0, 18, 5) = "`id`"
 Str(0, 18, 6) = "`id`"
 ArtZ(0, 18) = 4
 ArtZ(1, 18) = 2
 Str(1, 18, 0) = "CREATE TABLE `bezirke` ("
 Str(1, 18, 1) = " `id` int(10) NOT NULL"
 Str(1, 18, 2) = " `Bezirk` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 18, 3) = " `Nr` int(10) DEFAULT NULL"
 Str(1, 18, 4) = " `MaxNr` int(10) DEFAULT NULL"
 Str(1, 18, 5) = "  PRIMARY KEY (`id`)"
 Str(1, 18, 6) = "  UNIQUE KEY `id` (`id`)"
 Str(1, 18, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr18

Sub FüllStr19()
 Str(0, 19, 0) = "briefe"
 Str(0, 19, 1) = "`FID`"
 Str(0, 19, 2) = "`Pat_ID`"
 Str(0, 19, 3) = "`ZeitPunkt`"
 Str(0, 19, 4) = "`Pfad`"
 Str(0, 19, 5) = "`Art`"
 Str(0, 19, 6) = "`Name`"
 Str(0, 19, 7) = "`Typ`"
 Str(0, 19, 8) = "`AktZeit`"
 Str(0, 19, 9) = "`DokGroe`"
 Str(0, 19, 10) = "`QS`"
 Str(0, 19, 11) = "`QT`"
 Str(0, 19, 12) = "`absPos`"
 Str(0, 19, 13) = "`StByte`"
 Str(0, 19, 14) = "`Auswahl`"
 Str(0, 19, 15) = "`FälleBriefe`"
 Str(0, 19, 16) = "`FID`"
 Str(0, 19, 17) = "`NamenBriefe`"
 Str(0, 19, 18) = "`FälleBriefe_AccRel`"
 Str(0, 19, 19) = "`NamenBriefe_AccRel`"
 ArtZ(0, 19) = 13
 ArtZ(1, 19) = 4
 ArtZ(2, 19) = 2
 Str(1, 19, 0) = "CREATE TABLE `briefe` ("
 Str(1, 19, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 19, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 19, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 19, 4) = " `Pfad` varchar(128) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 5) = " `Art` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 6) = " `Name` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 7) = " `Typ` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 19, 8) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 19, 9) = " `DokGroe` int(10) DEFAULT NULL COMMENT 'Größe der Datei'"
 Str(1, 19, 10) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 19, 11) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 19, 12) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 19, 13) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 19, 14) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Name`)"
 Str(1, 19, 15) = "  KEY `FälleBriefe` (`FID`)"
 Str(1, 19, 16) = "  KEY `FID` (`FID`)"
 Str(1, 19, 17) = "  KEY `NamenBriefe` (`Pat_ID`)"
 Str(1, 19, 18) = "  CONSTRAINT `FälleBriefe_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 19, 19) = "  CONSTRAINT `NamenBriefe_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 19, 20) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr19

Sub FüllStr20()
 Str(0, 20, 0) = "dateien"
 Str(0, 20, 1) = "`ID`"
 Str(0, 20, 2) = "`Dateiname`"
 Str(0, 20, 3) = "`Pfad`"
 Str(0, 20, 4) = "`Größe`"
 Str(0, 20, 5) = "`Änderung`"
 Str(0, 20, 6) = "`Endung`"
 Str(0, 20, 7) = "`ID`"
 Str(0, 20, 8) = "`PrimaryKey`"
 Str(0, 20, 9) = "`Name`"
 ArtZ(0, 20) = 6
 ArtZ(1, 20) = 3
 Str(1, 20, 0) = "CREATE TABLE `dateien` ("
 Str(1, 20, 1) = " `ID` int(10) NOT NULL"
 Str(1, 20, 2) = " `Dateiname` varchar(128) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 20, 3) = " `Pfad` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Absolute Pfad'"
 Str(1, 20, 4) = " `Größe` double DEFAULT NULL COMMENT 'Dateigröße in Bytes'"
 Str(1, 20, 5) = " `Änderung` datetime DEFAULT NULL COMMENT 'letzte Änderung'"
 Str(1, 20, 6) = " `Endung` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Dateiendung'"
 Str(1, 20, 7) = "  PRIMARY KEY (`ID`)"
 Str(1, 20, 8) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 20, 9) = "  KEY `Name` (`Dateiname`,`Pfad`)"
 Str(1, 20, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr20

Sub FüllStr21()
 Str(0, 21, 0) = "diabetestypkontrolle"
 Str(0, 21, 1) = "`diabetestyp`"
 Str(0, 21, 2) = "`FID`"
 Str(0, 21, 3) = "`Pat_ID`"
 Str(0, 21, 4) = "`Quartal`"
 Str(0, 21, 5) = "`Nachname`"
 Str(0, 21, 6) = "`Vorname`"
 Str(0, 21, 7) = "`lfdnr`"
 Str(0, 21, 8) = "`TMFNr`"
 Str(0, 21, 9) = "`VKNr`"
 Str(0, 21, 10) = "`BhFB`"
 Str(0, 21, 11) = "`BhFE1`"
 Str(0, 21, 12) = "`BhFE2`"
 Str(0, 21, 13) = "`f4202`"
 Str(0, 21, 14) = "`ausgst`"
 Str(0, 21, 15) = "`KtrAbrB`"
 Str(0, 21, 16) = "`AbrAr`"
 Str(0, 21, 17) = "`lVorl`"
 Str(0, 21, 18) = "`IK`"
 Str(0, 21, 19) = "`KVKs`"
 Str(0, 21, 20) = "`KVKserg`"
 Str(0, 21, 21) = "`GebOr`"
 Str(0, 21, 22) = "`AbrGb`"
 Str(0, 21, 23) = "`f4206`"
 Str(0, 21, 24) = "`ÜwText`"
 Str(0, 21, 25) = "`f4210`"
 Str(0, 21, 26) = "`statNuller`"
 Str(0, 21, 27) = "`ÜbwV`"
 Str(0, 21, 28) = "`AndÜw`"
 Str(0, 21, 29) = "`Übw`"
 Str(0, 21, 30) = "`ÜWZiel`"
 Str(0, 21, 31) = "`ÜWNNr`"
 Str(0, 21, 32) = "`ÜWNaN`"
 Str(0, 21, 33) = "`ÜWTit`"
 Str(0, 21, 34) = "`ÜWVor`"
 Str(0, 21, 35) = "`ÜWVsw`"
 Str(0, 21, 36) = "`statKlasse`"
 Str(0, 21, 37) = "`f4237`"
 Str(0, 21, 38) = "`statBehTage`"
 Str(0, 21, 39) = "`SchGr`"
 Str(0, 21, 40) = "`Weiterbeh`"
 Str(0, 21, 41) = "`PGeb`"
 Str(0, 21, 42) = "`PGebErg`"
 Str(0, 21, 43) = "`Mahnfrist`"
 Str(0, 21, 44) = "`GOÄKatNr`"
 Str(0, 21, 45) = "`GOÄKatName`"
 Str(0, 21, 46) = "`AdNam`"
 Str(0, 21, 47) = "`AdStr`"
 Str(0, 21, 48) = "`AdPlz`"
 Str(0, 21, 49) = "`AdOrt`"
 Str(0, 21, 50) = "`BhFE`"
 Str(0, 21, 51) = "`s8000`"
 Str(0, 21, 52) = "`s8100`"
 Str(0, 21, 53) = "`AktZeit`"
 Str(0, 21, 54) = "`Fanf`"
 Str(0, 21, 55) = "`altQuart`"
 Str(0, 21, 56) = "`QAnf`"
 Str(0, 21, 57) = "`QEnd`"
 Str(0, 21, 58) = "`QS`"
 Str(0, 21, 59) = "`QT`"
 Str(0, 21, 60) = "`StByte`"
 Str(0, 21, 61) = "`absPos`"
 Str(0, 21, 62) = "`id`"
 Str(0, 21, 63) = "`vk`"
 Str(0, 21, 64) = "`kname`"
 Str(0, 21, 65) = "`kateg`"
 Str(0, 21, 66) = "`anzahlik`"
 Str(0, 21, 67) = "`anzahlktug`"
 Str(0, 21, 68) = "`gültigvon`"
 Str(0, 21, 69) = "`gültigbis`"
 Str(0, 21, 70) = "`go`"
 Str(0, 21, 71) = "`kurzname`"
 ArtZ(0, 21) = 71
 Str(1, 21, 0) = "CREATE TABLE `diabetestypkontrolle` ("
 Str(1, 21, 1) = " `diabetestyp` varchar(1) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 2) = " `FID` int(10) DEFAULT NULL"
 Str(1, 21, 3) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 21, 4) = " `Quartal` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 5) = " `Nachname` varchar(50) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 6) = " `Vorname` varchar(50) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 7) = " `lfdnr` int(10) DEFAULT NULL"
 Str(1, 21, 8) = " `TMFNr` varchar(20) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 9) = " `VKNr` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 10) = " `BhFB` datetime DEFAULT NULL"
 Str(1, 21, 11) = " `BhFE1` datetime DEFAULT NULL"
 Str(1, 21, 12) = " `BhFE2` datetime DEFAULT NULL"
 Str(1, 21, 13) = " `f4202` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 14) = " `ausgst` datetime DEFAULT NULL"
 Str(1, 21, 15) = " `KtrAbrB` varchar(3) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 16) = " `AbrAr` varchar(2) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 17) = " `lVorl` datetime DEFAULT NULL"
 Str(1, 21, 18) = " `IK` varchar(50) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 19) = " `KVKs` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 20) = " `KVKserg` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 21) = " `GebOr` varchar(2) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 22) = " `AbrGb` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 23) = " `f4206` varchar(15) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 24) = " `ÜwText` longtext CHARACTER SET utf8"
 Str(1, 21, 25) = " `f4210` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 26) = " `statNuller` varchar(20) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 27) = " `ÜbwV` varchar(22) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 28) = " `AndÜw` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 29) = " `Übw` varchar(8) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 30) = " `ÜWZiel` varchar(50) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 31) = " `ÜWNNr` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 32) = " `ÜWNaN` varchar(53) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 33) = " `ÜWTit` varchar(39) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 34) = " `ÜWVor` varchar(35) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 35) = " `ÜWVsw` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 36) = " `statKlasse` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 37) = " `f4237` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 38) = " `statBehTage` int(10) DEFAULT NULL"
 Str(1, 21, 39) = " `SchGr` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 40) = " `Weiterbeh` varchar(70) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 41) = " `PGeb` varchar(70) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 42) = " `PGebErg` varchar(50) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 43) = " `Mahnfrist` varchar(8) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 44) = " `GOÄKatNr` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 45) = " `GOÄKatName` varchar(30) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 46) = " `AdNam` varchar(28) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 47) = " `AdStr` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 48) = " `AdPlz` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 49) = " `AdOrt` varchar(30) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 50) = " `BhFE` datetime DEFAULT NULL"
 Str(1, 21, 51) = " `s8000` varchar(4) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 52) = " `s8100` varchar(7) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 53) = " `AktZeit` datetime DEFAULT NULL"
 Str(1, 21, 54) = " `Fanf` datetime DEFAULT NULL"
 Str(1, 21, 55) = " `altQuart` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 56) = " `QAnf` datetime DEFAULT NULL"
 Str(1, 21, 57) = " `QEnd` datetime DEFAULT NULL"
 Str(1, 21, 58) = " `QS` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 59) = " `QT` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 60) = " `StByte` int(10) DEFAULT NULL"
 Str(1, 21, 61) = " `absPos` int(10) DEFAULT NULL"
 Str(1, 21, 62) = " `id` int(10) DEFAULT NULL"
 Str(1, 21, 63) = " `vk` varchar(6) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 64) = " `kname` varchar(100) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 65) = " `kateg` varchar(3) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 66) = " `anzahlik` int(2) unsigned DEFAULT NULL"
 Str(1, 21, 67) = " `anzahlktug` int(2) unsigned DEFAULT NULL"
 Str(1, 21, 68) = " `gültigvon` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 69) = " `gültigbis` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 70) = " `go` varchar(15) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 71) = " `kurzname` varchar(60) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 21, 72) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr21

Sub FüllStr22()
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
 Str(1, 22, 28) = " ENGINE=InnoDB AUTO_INCREMENT=26107 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr22

Sub FüllStr23()
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
 Str(1, 23, 10) = " ENGINE=InnoDB AUTO_INCREMENT=798 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr23

Sub FüllStr24()
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
 Str(1, 24, 14) = " ENGINE=InnoDB AUTO_INCREMENT=893 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr24

Sub FüllStr25()
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
 Str(1, 25, 1) = " `lfdnr` int(2) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 25, 2) = " `gruppe` varchar(45) COLLATE latin1_german2_ci NOT NULL COMMENT 'Gruppenname'"
 Str(1, 25, 3) = " `rf` int(2) unsigned NOT NULL COMMENT 'Reihenfolge'"
 Str(1, 25, 4) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 25, 5) = "  KEY `gruppe` (`gruppe`)"
 Str(1, 25, 6) = "  KEY `rf` (`rf`)"
 Str(1, 25, 7) = " ENGINE=InnoDB AUTO_INCREMENT=23 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Diagnosengruppierung 1'"
End Sub ' FüllStr25

Sub FüllStr26()
 Str(0, 26, 0) = "diagnoseng2"
 Str(0, 26, 1) = "`id`"
 Str(0, 26, 2) = "`gruppe`"
 Str(0, 26, 3) = "`rf`"
 Str(0, 26, 4) = "`id`"
 Str(0, 26, 5) = "`gruppe`"
 Str(0, 26, 6) = "`rf`"
 ArtZ(0, 26) = 3
 ArtZ(1, 26) = 3
 Str(1, 26, 0) = "CREATE TABLE `diagnoseng2` ("
 Str(1, 26, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 26, 2) = " `gruppe` varchar(45) COLLATE latin1_german2_ci NOT NULL COMMENT 'Gruppenname'"
 Str(1, 26, 3) = " `rf` int(10) unsigned NOT NULL COMMENT 'Reihenfolge'"
 Str(1, 26, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 26, 5) = "  UNIQUE KEY `gruppe` (`gruppe`) USING BTREE"
 Str(1, 26, 6) = "  KEY `rf` (`rf`)"
 Str(1, 26, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Diagnosengruppierung 2'"
End Sub ' FüllStr26

Sub FüllStr27()
 Str(0, 27, 0) = "diagreihe"
 Str(0, 27, 1) = "`lfdnr`"
 Str(0, 27, 2) = "`ICD`"
 Str(0, 27, 3) = "`dg1`"
 Str(0, 27, 4) = "`dg2`"
 Str(0, 27, 5) = "`rf`"
 Str(0, 27, 6) = "`lfdnr`"
 Str(0, 27, 7) = "`ICD`"
 Str(0, 27, 8) = "`rf`"
 Str(0, 27, 9) = "`dg1`"
 Str(0, 27, 10) = "`dg1`"
 Str(0, 27, 11) = "`dg1_rel`"
 Str(0, 27, 12) = "`ICD_rel`"
 ArtZ(0, 27) = 5
 ArtZ(1, 27) = 4
 ArtZ(2, 27) = 3
 Str(1, 27, 0) = "CREATE TABLE `diagreihe` ("
 Str(1, 27, 1) = " `lfdnr` int(2) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 27, 2) = " `ICD` varchar(10) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 27, 3) = " `dg1` varchar(45) COLLATE latin1_german2_ci NOT NULL COMMENT 'Diagnosenguppierung 1'"
 Str(1, 27, 4) = " `dg2` varchar(45) COLLATE latin1_german2_ci NOT NULL COMMENT 'Diagnosenguppierung 2'"
 Str(1, 27, 5) = " `rf` int(2) unsigned NOT NULL COMMENT 'Reihenfolge für Arztbrief'"
 Str(1, 27, 6) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 27, 7) = "  UNIQUE KEY `ICD` (`ICD`)"
 Str(1, 27, 8) = "  KEY `rf` (`rf`)"
 Str(1, 27, 9) = "  KEY `dg1` (`dg1`)"
 Str(1, 27, 10) = "  CONSTRAINT `dg1` FOREIGN KEY (`dg1`) REFERENCES `diagnoseng1` (`gruppe`) ON UPDATE CASCADE"
 Str(1, 27, 11) = "  CONSTRAINT `dg1_rel` FOREIGN KEY (`dg1`) REFERENCES `diagnoseng1` (`gruppe`) ON UPDATE CASCADE"
 Str(1, 27, 12) = "  CONSTRAINT `ICD_rel` FOREIGN KEY (`ICD`) REFERENCES `diagnosen` (`ICD`)"
 Str(1, 27, 13) = " ENGINE=InnoDB AUTO_INCREMENT=810 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Reihenfolge der Diagnosen für Arztbrief'"
End Sub ' FüllStr27

Sub FüllStr28()
 Str(0, 28, 0) = "dmp-uschr"
 Str(0, 28, 1) = "`Pat_id`"
 Str(0, 28, 2) = "`U1`"
 Str(0, 28, 3) = "`U2`"
 Str(0, 28, 4) = "`U3`"
 Str(0, 28, 5) = "`Arztwechsel`"
 Str(0, 28, 6) = "`Pat_id`"
 ArtZ(0, 28) = 5
 ArtZ(1, 28) = 1
 Str(1, 28, 0) = "CREATE TABLE `dmp-uschr` ("
 Str(1, 28, 1) = " `Pat_id` int(10) DEFAULT NULL COMMENT 'Bezug auf Namen'"
 Str(1, 28, 2) = " `U1` datetime DEFAULT NULL COMMENT 'vorliegendes Blatt mit DMP-Unterschrift'"
 Str(1, 28, 3) = " `U2` datetime DEFAULT NULL COMMENT 'vorliegendes 2. Blatt mit DMP-Unterschrift'"
 Str(1, 28, 4) = " `U3` datetime DEFAULT NULL COMMENT 'vorliegendes 3. Blatt mit DMP-Unterschrift'"
 Str(1, 28, 5) = " `Arztwechsel` tinyint(1) unsigned DEFAULT NULL COMMENT 'danach Arztwechsel'"
 Str(1, 28, 6) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 28, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr28

Sub FüllStr29()
 Str(0, 29, 0) = "dmpreihe"
 Str(0, 29, 1) = "`Abk`"
 Str(0, 29, 2) = "`Art`"
 Str(0, 29, 3) = "`KarteiDatum`"
 Str(0, 29, 4) = "`exportiert`"
 Str(0, 29, 5) = "`DokuDatum`"
 Str(0, 29, 6) = "`obvoll`"
 Str(0, 29, 7) = "`NachName`"
 Str(0, 29, 8) = "`VorName`"
 Str(0, 29, 9) = "`GebDat`"
 Str(0, 29, 10) = "`Pat_id`"
 Str(0, 29, 11) = "`StByte`"
 Str(0, 29, 12) = "`AktZeit`"
 ArtZ(0, 29) = 12
 Str(1, 29, 0) = "CREATE TABLE `dmpreihe` ("
 Str(1, 29, 1) = " `Abk` varchar(30) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Abkürzung der DMP-Art'"
 Str(1, 29, 2) = " `Art` varchar(2) CHARACTER SET latin1 NOT NULL COMMENT 'ED = Erstdoku, FD = Folgedoku'"
 Str(1, 29, 3) = " `KarteiDatum` date DEFAULT NULL COMMENT 'Datum des Karteikarteneintrags der Dokumentation'"
 Str(1, 29, 4) = " `exportiert` datetime DEFAULT NULL COMMENT 'Datum des Exports'"
 Str(1, 29, 5) = " `DokuDatum` datetime DEFAULT NULL COMMENT 'Datum der Dokumentation'"
 Str(1, 29, 6) = " `obvoll` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob vollständig'"
 Str(1, 29, 7) = " `NachName` varchar(20) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 29, 8) = " `VorName` varchar(20) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 29, 9) = " `GebDat` date DEFAULT NULL"
 Str(1, 29, 10) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 29, 11) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 29, 12) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungzeit'"
 Str(1, 29, 13) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr29

Sub FüllStr30()
 Str(0, 30, 0) = "dokab"
 Str(0, 30, 1) = "`ob`"
 Str(0, 30, 2) = "`dokname`"
 Str(0, 30, 3) = "`dokpfad`"
 Str(0, 30, 4) = "`id`"
 Str(0, 30, 5) = "`id`"
 ArtZ(0, 30) = 4
 ArtZ(1, 30) = 1
 Str(1, 30, 0) = "CREATE TABLE `dokab` ("
 Str(1, 30, 1) = " `ob` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 30, 2) = " `dokname` varchar(256) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 30, 3) = " `dokpfad` varchar(255) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 30, 4) = " `id` int(10) NOT NULL"
 Str(1, 30, 5) = "  PRIMARY KEY (`id`)"
 Str(1, 30, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr30

Sub FüllStr31()
 Str(0, 31, 0) = "dokumente"
 Str(0, 31, 1) = "`FID`"
 Str(0, 31, 2) = "`Pat_ID`"
 Str(0, 31, 3) = "`ZeitPunkt`"
 Str(0, 31, 4) = "`DokPfad`"
 Str(0, 31, 5) = "`DokArt`"
 Str(0, 31, 6) = "`DokName`"
 Str(0, 31, 7) = "`Quelldatum`"
 Str(0, 31, 8) = "`absPos`"
 Str(0, 31, 9) = "`AktZeit`"
 Str(0, 31, 10) = "`DokGroe`"
 Str(0, 31, 11) = "`QS`"
 Str(0, 31, 12) = "`QT`"
 Str(0, 31, 13) = "`StByte`"
 Str(0, 31, 14) = "`Auswahl`"
 Str(0, 31, 15) = "`DokName`"
 Str(0, 31, 16) = "`DokPfad`"
 Str(0, 31, 17) = "`FälleDokumente`"
 Str(0, 31, 18) = "`FID`"
 Str(0, 31, 19) = "`NamenDokumente`"
 Str(0, 31, 20) = "`PIDokPfad`"
 Str(0, 31, 21) = "`Quelldatum`"
 Str(0, 31, 22) = "`ZeitPunkt`"
 Str(0, 31, 23) = "`FälleDokumente_AccRel`"
 Str(0, 31, 24) = "`NamenDokumente_AccRel`"
 ArtZ(0, 31) = 13
 ArtZ(1, 31) = 9
 ArtZ(2, 31) = 2
 Str(1, 31, 0) = "CREATE TABLE `dokumente` ("
 Str(1, 31, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 31, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 31, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 31, 4) = " `DokPfad` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 31, 5) = " `DokArt` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 31, 6) = " `DokName` varchar(1557) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 31, 7) = " `Quelldatum` datetime DEFAULT NULL COMMENT 'Datum, auf das sich das Dokument bezieht'"
 Str(1, 31, 8) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 31, 9) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 31, 10) = " `DokGroe` int(10) DEFAULT NULL COMMENT 'Dokument-Größe'"
 Str(1, 31, 11) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 31, 12) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 31, 13) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 31, 14) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`DokArt`,`DokName`(767))"
 Str(1, 31, 15) = "  KEY `DokName` (`DokName`(767))"
 Str(1, 31, 16) = "  KEY `DokPfad` (`DokPfad`)"
 Str(1, 31, 17) = "  KEY `FälleDokumente` (`FID`)"
 Str(1, 31, 18) = "  KEY `FID` (`FID`)"
 Str(1, 31, 19) = "  KEY `NamenDokumente` (`Pat_ID`)"
 Str(1, 31, 20) = "  KEY `PIDokPfad` (`Pat_ID`,`DokPfad`)"
 Str(1, 31, 21) = "  KEY `Quelldatum` (`Quelldatum`)"
 Str(1, 31, 22) = "  KEY `ZeitPunkt` (`ZeitPunkt`)"
 Str(1, 31, 23) = "  CONSTRAINT `FälleDokumente_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 31, 24) = "  CONSTRAINT `NamenDokumente_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 31, 25) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr31

Sub FüllStr32()
 Str(0, 32, 0) = "dokumente abgehakt"
 Str(0, 32, 1) = "`DokPfad`"
 Str(0, 32, 2) = "`AktZeit`"
 Str(0, 32, 3) = "`abgehakt`"
 Str(0, 32, 4) = "`ungueltig`"
 Str(0, 32, 5) = "`DokPfad`"
 ArtZ(0, 32) = 4
 ArtZ(1, 32) = 1
 Str(1, 32, 0) = "CREATE TABLE `dokumente abgehakt` ("
 Str(1, 32, 1) = " `DokPfad` varchar(140) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 32, 2) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 32, 3) = " `abgehakt` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 32, 4) = " `ungueltig` bit(1) DEFAULT NULL"
 Str(1, 32, 5) = "  KEY `DokPfad` (`DokPfad`)"
 Str(1, 32, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr32

Sub FüllStr33()
 Str(0, 33, 0) = "dopplerzahl"
 Str(0, 33, 1) = "`count(0)`"
 Str(0, 33, 2) = "`quartal`"
 ArtZ(0, 33) = 2
 Str(1, 33, 0) = "CREATE TABLE `dopplerzahl` ("
 Str(1, 33, 1) = " `count(0)` decimal(15,4) DEFAULT NULL"
 Str(1, 33, 2) = " `quartal` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 33, 3) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr33

Sub FüllStr34()
 Str(0, 34, 0) = "duplexzahl"
 Str(0, 34, 1) = "`count(0)`"
 Str(0, 34, 2) = "`quartal`"
 ArtZ(0, 34) = 2
 Str(1, 34, 0) = "CREATE TABLE `duplexzahl` ("
 Str(1, 34, 1) = " `count(0)` decimal(15,4) DEFAULT NULL"
 Str(1, 34, 2) = " `quartal` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 34, 3) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr34

Sub FüllStr35()
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
 Str(0, 35, 26) = "`Leistung`"
 Str(0, 35, 27) = "`Titel`"
 ArtZ(0, 35) = 25
 ArtZ(1, 35) = 2
 Str(1, 35, 0) = "CREATE TABLE `ebm2000plus` ("
 Str(1, 35, 1) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Leistungsziffer'"
 Str(1, 35, 2) = " `Titel` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kurztext'"
 Str(1, 35, 3) = " `Punktwert` int(10) DEFAULT NULL COMMENT 'Punktwert'"
 Str(1, 35, 4) = " `Euro` int(10) DEFAULT NULL COMMENT '€'"
 Str(1, 35, 5) = " `Bericht` tinyint(1) unsigned DEFAULT NULL COMMENT 'Berichtspflicht'"
 Str(1, 35, 6) = " `Text` longtext COLLATE latin1_german2_ci COMMENT 'restlicher Leistungstext'"
 Str(1, 35, 7) = " `Betr` tinyint(1) unsigned DEFAULT NULL COMMENT 'Betreuung'"
 Str(1, 35, 8) = " `Schul` tinyint(1) unsigned DEFAULT NULL COMMENT 'Schulung'"
 Str(1, 35, 9) = " `Typ1` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 10) = " `Typ2` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 11) = " `Gest` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 12) = " `DFS` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 13) = " `DMP` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 14) = " `AOK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 15) = " `BKK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 16) = " `BKN` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 17) = " `EK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 18) = " `IKK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 19) = " `LKK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 20) = " `Üw` tinyint(1) unsigned DEFAULT NULL COMMENT 'Überweisung durch HA nötig'"
 Str(1, 35, 21) = " `Insulin` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 22) = " `ICT` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 23) = " `CSII` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 24) = " `Erst` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 25) = " `Folge` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 35, 26) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 35, 27) = "  KEY `Titel` (`Titel`)"
 Str(1, 35, 28) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr35

Sub FüllStr36()
 Str(0, 36, 0) = "ebm2000plus1"
 Str(0, 36, 1) = "`Leistung`"
 Str(0, 36, 2) = "`Titel`"
 Str(0, 36, 3) = "`Punktwert`"
 Str(0, 36, 4) = "`Euro`"
 Str(0, 36, 5) = "`Bericht`"
 Str(0, 36, 6) = "`Text`"
 Str(0, 36, 7) = "`Betr`"
 Str(0, 36, 8) = "`Schul`"
 Str(0, 36, 9) = "`Typ1`"
 Str(0, 36, 10) = "`Typ2`"
 Str(0, 36, 11) = "`Gest`"
 Str(0, 36, 12) = "`DFS`"
 Str(0, 36, 13) = "`DMP`"
 Str(0, 36, 14) = "`AOK`"
 Str(0, 36, 15) = "`BKK`"
 Str(0, 36, 16) = "`BKN`"
 Str(0, 36, 17) = "`EK`"
 Str(0, 36, 18) = "`IKK`"
 Str(0, 36, 19) = "`LKK`"
 Str(0, 36, 20) = "`Üw`"
 Str(0, 36, 21) = "`Insulin`"
 Str(0, 36, 22) = "`ICT`"
 Str(0, 36, 23) = "`CSII`"
 Str(0, 36, 24) = "`Erst`"
 Str(0, 36, 25) = "`Folge`"
 Str(0, 36, 26) = "`Leistung`"
 Str(0, 36, 27) = "`Titel`"
 ArtZ(0, 36) = 25
 ArtZ(1, 36) = 2
 Str(1, 36, 0) = "CREATE TABLE `ebm2000plus1` ("
 Str(1, 36, 1) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Leistungsziffer'"
 Str(1, 36, 2) = " `Titel` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kurztext'"
 Str(1, 36, 3) = " `Punktwert` int(10) DEFAULT NULL COMMENT 'Punktwert'"
 Str(1, 36, 4) = " `Euro` int(10) DEFAULT NULL COMMENT '€'"
 Str(1, 36, 5) = " `Bericht` tinyint(1) unsigned DEFAULT NULL COMMENT 'Berichtspflicht'"
 Str(1, 36, 6) = " `Text` longtext COLLATE latin1_german2_ci COMMENT 'restlicher Leistungstext'"
 Str(1, 36, 7) = " `Betr` tinyint(1) unsigned DEFAULT NULL COMMENT 'Betreuung'"
 Str(1, 36, 8) = " `Schul` tinyint(1) unsigned DEFAULT NULL COMMENT 'Schulung'"
 Str(1, 36, 9) = " `Typ1` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 10) = " `Typ2` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 11) = " `Gest` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 12) = " `DFS` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 13) = " `DMP` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 14) = " `AOK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 15) = " `BKK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 16) = " `BKN` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 17) = " `EK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 18) = " `IKK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 19) = " `LKK` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 20) = " `Üw` tinyint(1) unsigned DEFAULT NULL COMMENT 'Überweisung durch HA nötig'"
 Str(1, 36, 21) = " `Insulin` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 22) = " `ICT` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 23) = " `CSII` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 24) = " `Erst` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 25) = " `Folge` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 36, 26) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 36, 27) = "  KEY `Titel` (`Titel`)"
 Str(1, 36, 28) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr36

Sub FüllStr37()
 Str(0, 37, 0) = "einfuegefehler"
 Str(0, 37, 1) = "`Name`"
 Str(0, 37, 2) = "`Datum`"
 Str(0, 37, 3) = "`Pat_id`"
 Str(0, 37, 4) = "`archiviert`"
 Str(0, 37, 5) = "`Protokoll`"
 ArtZ(0, 37) = 5
 Str(1, 37, 0) = "CREATE TABLE `einfuegefehler` ("
 Str(1, 37, 1) = " `Name` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 37, 2) = " `Datum` datetime DEFAULT NULL"
 Str(1, 37, 3) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 37, 4) = " `archiviert` datetime DEFAULT NULL"
 Str(1, 37, 5) = " `Protokoll` longtext COLLATE latin1_german2_ci"
 Str(1, 37, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr37

Sub FüllStr38()
 Str(0, 38, 0) = "einstellungen"
 Str(0, 38, 1) = "`ID`"
 Str(0, 38, 2) = "`Formular`"
 Str(0, 38, 3) = "`Abfrage für Formular`"
 Str(0, 38, 4) = "`ID für Formular`"
 Str(0, 38, 5) = "`DatensatzNr`"
 Str(0, 38, 6) = "`ID`"
 Str(0, 38, 7) = "`PrimaryKey`"
 Str(0, 38, 8) = "`Formular`"
 Str(0, 38, 9) = "`ID für Formular`"
 ArtZ(0, 38) = 5
 ArtZ(1, 38) = 4
 Str(1, 38, 0) = "CREATE TABLE `einstellungen` ("
 Str(1, 38, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 38, 2) = " `Formular` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name des Formulars'"
 Str(1, 38, 3) = " `Abfrage für Formular` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name der Abfrage, die zuletzt für das Formular ""Anamnesebogen"" verwendet wurde'"
 Str(1, 38, 4) = " `ID für Formular` int(10) DEFAULT NULL COMMENT 'Pat_ID in dieser Abfrage'"
 Str(1, 38, 5) = " `DatensatzNr` int(10) DEFAULT NULL COMMENT 'Datensatz-Nr. in dieser Abfrage'"
 Str(1, 38, 6) = "  PRIMARY KEY (`ID`)"
 Str(1, 38, 7) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 38, 8) = "  KEY `Formular` (`Formular`)"
 Str(1, 38, 9) = "  KEY `ID für Formular` (`ID für Formular`)"
 Str(1, 38, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr38

Sub FüllStr39()
 Str(0, 39, 0) = "eintraege"
 Str(0, 39, 1) = "`FID`"
 Str(0, 39, 2) = "`Pat_ID`"
 Str(0, 39, 3) = "`ZeitPunkt`"
 Str(0, 39, 4) = "`Art`"
 Str(0, 39, 5) = "`Inhalt`"
 Str(0, 39, 6) = "`absPos`"
 Str(0, 39, 7) = "`AktZeit`"
 Str(0, 39, 8) = "`QS`"
 Str(0, 39, 9) = "`QT`"
 Str(0, 39, 10) = "`StByte`"
 Str(0, 39, 11) = "`Auswahl`"
 Str(0, 39, 12) = "`FälleEinträge`"
 Str(0, 39, 13) = "`NamenEinträge`"
 Str(0, 39, 14) = "`Art`"
 Str(0, 39, 15) = "`FälleEinträge_AccRel`"
 Str(0, 39, 16) = "`NamenEinträge_AccRel`"
 ArtZ(0, 39) = 10
 ArtZ(1, 39) = 4
 ArtZ(2, 39) = 2
 Str(1, 39, 0) = "CREATE TABLE `eintraege` ("
 Str(1, 39, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 39, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 39, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 39, 4) = " `Art` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6330'"
 Str(1, 39, 5) = " `Inhalt` longtext COLLATE latin1_german2_ci COMMENT '8480'"
 Str(1, 39, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 39, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 39, 8) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 39, 9) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 39, 10) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 39, 11) = "  KEY `Auswahl` (`Pat_ID`,`Art`,`ZeitPunkt`)"
 Str(1, 39, 12) = "  KEY `FälleEinträge` (`FID`)"
 Str(1, 39, 13) = "  KEY `NamenEinträge` (`Pat_ID`)"
 Str(1, 39, 14) = "  KEY `Art` (`Art`)"
 Str(1, 39, 15) = "  CONSTRAINT `FälleEinträge_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 39, 16) = "  CONSTRAINT `NamenEinträge_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 39, 17) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr39

Sub FüllStr40()
 Str(0, 40, 0) = "eintraege arten"
 Str(0, 40, 1) = "`ID`"
 Str(0, 40, 2) = "`Art`"
 Str(0, 40, 3) = "`obAutom`"
 Str(0, 40, 4) = "`Erklärung`"
 Str(0, 40, 5) = "`ID`"
 Str(0, 40, 6) = "`PrimaryKey`"
 ArtZ(0, 40) = 4
 ArtZ(1, 40) = 2
 Str(1, 40, 0) = "CREATE TABLE `eintraege arten` ("
 Str(1, 40, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 40, 2) = " `Art` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 40, 3) = " `obAutom` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Art auch automatisch aus Formular entsteht'"
 Str(1, 40, 4) = " `Erklärung` longtext COLLATE latin1_german2_ci COMMENT 'Erklärung für die Art'"
 Str(1, 40, 5) = "  PRIMARY KEY (`ID`)"
 Str(1, 40, 6) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 40, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr40

Sub FüllStr41()
 Str(0, 41, 0) = "eintragszahlen"
 Str(0, 41, 1) = "`Beginn`"
 Str(0, 41, 2) = "`StByte`"
 Str(0, 41, 3) = "`Zp1`"
 Str(0, 41, 4) = "`Zp2`"
 Str(0, 41, 5) = "`Zp3`"
 Str(0, 41, 6) = "`Zp4`"
 Str(0, 41, 7) = "`Zp5`"
 Str(0, 41, 8) = "`Zp6`"
 Str(0, 41, 9) = "`Zp7`"
 Str(0, 41, 10) = "`Zp8`"
 Str(0, 41, 11) = "`Fallzahl`"
 Str(0, 41, 12) = "`Sekunden`"
 Str(0, 41, 13) = "`Datei`"
 Str(0, 41, 14) = "`DateiAend`"
 Str(0, 41, 15) = "`SpeicherZt`"
 Str(0, 41, 16) = "`TabellenEntleeren`"
 Str(0, 41, 17) = "`ZurücksetzenLAktDat`"
 Str(0, 41, 18) = "`Pat_IDVon`"
 Str(0, 41, 19) = "`Pat_IDbis`"
 Str(0, 41, 20) = "`VorladenFFI`"
 Str(0, 41, 21) = "`ÜberTabelle`"
 Str(0, 41, 22) = "`SammelInsert`"
 Str(0, 41, 23) = "`bereinigeFormInhFeld`"
 Str(0, 41, 24) = "`LaborDirektEinlesen`"
 Str(0, 41, 25) = "`LaborDirektNeu`"
 Str(0, 41, 26) = "`LaborQuerVerb`"
 Str(0, 41, 27) = "`LaborQuerNeu`"
 Str(0, 41, 28) = "`AlterTab`"
 Str(0, 41, 29) = "`obmitEmails`"
 Str(0, 41, 30) = "`LaborPfadBeispiel`"
 Str(0, 41, 31) = "`obVglMitLetzterEinlesung`"
 Str(0, 41, 32) = "`Beginn`"
 Str(0, 41, 33) = "`Beginn`"
 ArtZ(0, 41) = 31
 ArtZ(1, 41) = 2
 Str(1, 41, 0) = "CREATE TABLE `eintragszahlen` ("
 Str(1, 41, 1) = " `Beginn` datetime NOT NULL"
 Str(1, 41, 2) = " `StByte` int(10) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 41, 3) = " `Zp1` datetime DEFAULT NULL"
 Str(1, 41, 4) = " `Zp2` datetime DEFAULT NULL"
 Str(1, 41, 5) = " `Zp3` datetime DEFAULT NULL"
 Str(1, 41, 6) = " `Zp4` datetime DEFAULT NULL"
 Str(1, 41, 7) = " `Zp5` datetime DEFAULT NULL"
 Str(1, 41, 8) = " `Zp6` datetime DEFAULT NULL"
 Str(1, 41, 9) = " `Zp7` datetime DEFAULT NULL"
 Str(1, 41, 10) = " `Zp8` datetime DEFAULT NULL"
 Str(1, 41, 11) = " `Fallzahl` int(10) DEFAULT NULL"
 Str(1, 41, 12) = " `Sekunden` int(10) DEFAULT NULL"
 Str(1, 41, 13) = " `Datei` varchar(120) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 41, 14) = " `DateiAend` datetime DEFAULT NULL"
 Str(1, 41, 15) = " `SpeicherZt` datetime DEFAULT NULL"
 Str(1, 41, 16) = " `TabellenEntleeren` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 17) = " `ZurücksetzenLAktDat` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 18) = " `Pat_IDVon` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 41, 19) = " `Pat_IDbis` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 41, 20) = " `VorladenFFI` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 21) = " `ÜberTabelle` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 22) = " `SammelInsert` tinyint(1) unsigned NOT NULL"
 Str(1, 41, 23) = " `bereinigeFormInhFeld` tinyint(1) unsigned NOT NULL COMMENT 'ob FormInhFeld bereinigt wird'"
 Str(1, 41, 24) = " `LaborDirektEinlesen` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 25) = " `LaborDirektNeu` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 26) = " `LaborQuerVerb` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 27) = " `LaborQuerNeu` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 28) = " `AlterTab` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 29) = " `obmitEmails` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 30) = " `LaborPfadBeispiel` longtext COLLATE latin1_german2_ci"
 Str(1, 41, 31) = " `obVglMitLetzterEinlesung` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 41, 32) = "  PRIMARY KEY (`Beginn`)"
 Str(1, 41, 33) = "  UNIQUE KEY `Beginn` (`Beginn`)"
 Str(1, 41, 34) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr41

Sub FüllStr42()
 Str(0, 42, 0) = "eintrhist"
 Str(1, 42, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `eintrhist` AS select `eintrhist1`.`ID` AS `ID`,`eintrhist1`.`Pat_ID` AS `Pat_ID`,`eintrhist1`.`ZeitPunkt` AS `ZeitPunkt`,`eintrhist1`.`Art` AS `Art`,`eintrhist1`.`Inhalt` AS `Inhalt`,`eintrhist1`.`QS` AS `QS`,`eintrhist1`.`QT` AS `QT`,`eintrhist2`.`FID` AS `FID`,`eintrhist2`.`absPos` AS `absPos`,`eintrhist2`.`AktZeit` AS `AktZeit`,`eintrhist2`.`StByte` AS `StByte` from (`eintrhist1` join `eintrhist2` on((`eintrhist1`.`ID` = `eintrhist2`.`ID`)))"
End Sub ' FüllStr42

Sub FüllStr43()
 Str(0, 43, 0) = "eintrhist1"
 Str(0, 43, 1) = "`Pat_ID`"
 Str(0, 43, 2) = "`ZeitPunkt`"
 Str(0, 43, 3) = "`Art`"
 Str(0, 43, 4) = "`Inhalt`"
 Str(0, 43, 5) = "`QS`"
 Str(0, 43, 6) = "`QT`"
 Str(0, 43, 7) = "`ID`"
 Str(0, 43, 8) = "`ID`"
 Str(0, 43, 9) = "`Auswahl`"
 Str(0, 43, 10) = "`NamenEinträge`"
 Str(0, 43, 11) = "`Art`"
 ArtZ(0, 43) = 7
 ArtZ(1, 43) = 4
 Str(1, 43, 0) = "CREATE TABLE `eintrhist1` ("
 Str(1, 43, 1) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 43, 2) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 43, 3) = " `Art` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6330'"
 Str(1, 43, 4) = " `Inhalt` longtext COLLATE latin1_german2_ci COMMENT '8480'"
 Str(1, 43, 5) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 43, 6) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 43, 7) = " `ID` int(2) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 43, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 43, 9) = "  KEY `Auswahl` (`Pat_ID`,`Art`,`ZeitPunkt`)"
 Str(1, 43, 10) = "  KEY `NamenEinträge` (`Pat_ID`)"
 Str(1, 43, 11) = "  KEY `Art` (`Art`)"
 Str(1, 43, 12) = " ENGINE=InnoDB AUTO_INCREMENT=788 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr43

Sub FüllStr44()
 Str(0, 44, 0) = "eintrhist2"
 Str(0, 44, 1) = "`FID`"
 Str(0, 44, 2) = "`absPos`"
 Str(0, 44, 3) = "`AktZeit`"
 Str(0, 44, 4) = "`StByte`"
 Str(0, 44, 5) = "`ID`"
 Str(0, 44, 6) = "`FälleEinträge`"
 Str(0, 44, 7) = "`FK_id`"
 Str(0, 44, 8) = "`FK_id`"
 ArtZ(0, 44) = 5
 ArtZ(1, 44) = 2
 ArtZ(2, 44) = 1
 Str(1, 44, 0) = "CREATE TABLE `eintrhist2` ("
 Str(1, 44, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 44, 2) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 44, 3) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 44, 4) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 44, 5) = " `ID` int(2) unsigned NOT NULL"
 Str(1, 44, 6) = "  KEY `FälleEinträge` (`FID`)"
 Str(1, 44, 7) = "  KEY `FK_id` (`ID`)"
 Str(1, 44, 8) = "  CONSTRAINT `FK_id` FOREIGN KEY (`ID`) REFERENCES `eintrhist1` (`ID`)"
 Str(1, 44, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr44

Sub FüllStr45()
 Str(0, 45, 0) = "faelle"
 Str(0, 45, 1) = "`FID`"
 Str(0, 45, 2) = "`Pat_ID`"
 Str(0, 45, 3) = "`Quartal`"
 Str(0, 45, 4) = "`Nachname`"
 Str(0, 45, 5) = "`Vorname`"
 Str(0, 45, 6) = "`lfdnr`"
 Str(0, 45, 7) = "`TMFNr`"
 Str(0, 45, 8) = "`VKNr`"
 Str(0, 45, 9) = "`BhFB`"
 Str(0, 45, 10) = "`BhFE1`"
 Str(0, 45, 11) = "`BhFE2`"
 Str(0, 45, 12) = "`f4202`"
 Str(0, 45, 13) = "`ausgst`"
 Str(0, 45, 14) = "`KtrAbrB`"
 Str(0, 45, 15) = "`AbrAr`"
 Str(0, 45, 16) = "`lVorl`"
 Str(0, 45, 17) = "`IK`"
 Str(0, 45, 18) = "`KVKs`"
 Str(0, 45, 19) = "`KVKserg`"
 Str(0, 45, 20) = "`Kasse`"
 Str(0, 45, 21) = "`GebOr`"
 Str(0, 45, 22) = "`AbrGb`"
 Str(0, 45, 23) = "`PersKreis`"
 Str(0, 45, 24) = "`SKtZusatz`"
 Str(0, 45, 25) = "`f4206`"
 Str(0, 45, 26) = "`ÜwText`"
 Str(0, 45, 27) = "`f4210`"
 Str(0, 45, 28) = "`AkfHAH`"
 Str(0, 45, 29) = "`AkfAB0`"
 Str(0, 45, 30) = "`AkfAK`"
 Str(0, 45, 31) = "`statNuller`"
 Str(0, 45, 32) = "`ÜbwV`"
 Str(0, 45, 33) = "`AndÜw`"
 Str(0, 45, 34) = "`Übw`"
 Str(0, 45, 35) = "`ÜbwLANR`"
 Str(0, 45, 36) = "`ÜWZiel`"
 Str(0, 45, 37) = "`ÜWNNr`"
 Str(0, 45, 38) = "`ÜWNaN`"
 Str(0, 45, 39) = "`ÜWTit`"
 Str(0, 45, 40) = "`ÜWVor`"
 Str(0, 45, 41) = "`ÜWVsw`"
 Str(0, 45, 42) = "`üwvid`"
 Str(0, 45, 43) = "`statKlasse`"
 Str(0, 45, 44) = "`f4237`"
 Str(0, 45, 45) = "`statBehTage`"
 Str(0, 45, 46) = "`SchGr`"
 Str(0, 45, 47) = "`Weiterbeh`"
 Str(0, 45, 48) = "`PGeb`"
 Str(0, 45, 49) = "`PGebErg`"
 Str(0, 45, 50) = "`Mahnfrist`"
 Str(0, 45, 51) = "`GOÄKatNr`"
 Str(0, 45, 52) = "`GOÄKatName`"
 Str(0, 45, 53) = "`abrArzt`"
 Str(0, 45, 54) = "`privVers`"
 Str(0, 45, 55) = "`AdNam`"
 Str(0, 45, 56) = "`AdStr`"
 Str(0, 45, 57) = "`AdPlz`"
 Str(0, 45, 58) = "`AdOrt`"
 Str(0, 45, 59) = "`BhFE`"
 Str(0, 45, 60) = "`s8000`"
 Str(0, 45, 61) = "`s8100`"
 Str(0, 45, 62) = "`AktZeit`"
 Str(0, 45, 63) = "`Fanf`"
 Str(0, 45, 64) = "`altQuart`"
 Str(0, 45, 65) = "`QAnf`"
 Str(0, 45, 66) = "`QEnd`"
 Str(0, 45, 67) = "`QS`"
 Str(0, 45, 68) = "`QT`"
 Str(0, 45, 69) = "`TherArt`"
 Str(0, 45, 70) = "`StByte`"
 Str(0, 45, 71) = "`absPos`"
 Str(0, 45, 72) = "`FID`"
 Str(0, 45, 73) = "`PrimaryKey`"
 Str(0, 45, 74) = "`AktF`"
 Str(0, 45, 75) = "`Auswahl`"
 Str(0, 45, 76) = "`BhFB`"
 Str(0, 45, 77) = "`FanF`"
 Str(0, 45, 78) = "`NamenFälle`"
 Str(0, 45, 79) = "`pQ`"
 Str(0, 45, 80) = "`Quartal`"
 Str(0, 45, 81) = "`SchGr`"
 Str(0, 45, 82) = "`vknr`"
 Str(0, 45, 83) = "`KassenlisteFälle_AccRel`"
 Str(0, 45, 84) = "`NamenFälle_AccRel`"
 ArtZ(0, 45) = 71
 ArtZ(1, 45) = 11
 ArtZ(2, 45) = 2
 Str(1, 45, 0) = "CREATE TABLE `faelle` ("
 Str(1, 45, 1) = " `FID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'"
 Str(1, 45, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 45, 3) = " `Quartal` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4101'"
 Str(1, 45, 4) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 45, 5) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 45, 6) = " `lfdnr` int(10) DEFAULT NULL COMMENT 'laufende Fallnummer'"
 Str(1, 45, 7) = " `TMFNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4144 Fallnummer in Turbomed'"
 Str(1, 45, 8) = " `VKNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4104'"
 Str(1, 45, 9) = " `BhFB` datetime DEFAULT NULL COMMENT '4150'"
 Str(1, 45, 10) = " `BhFE1` datetime DEFAULT NULL COMMENT '4151'"
 Str(1, 45, 11) = " `BhFE2` datetime DEFAULT NULL COMMENT '4152'"
 Str(1, 45, 12) = " `f4202` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4202'"
 Str(1, 45, 13) = " `ausgst` datetime DEFAULT NULL COMMENT '4102 (''ausgestellt am'')'"
 Str(1, 45, 14) = " `KtrAbrB` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))'"
 Str(1, 45, 15) = " `AbrAr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4107, Abrechnungsart (1 = Primärkassen)'"
 Str(1, 45, 16) = " `lVorl` datetime DEFAULT NULL COMMENT '4109, letzte Vorlage'"
 Str(1, 45, 17) = " `IK` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4111 Krankenkassennummer (IK)'"
 Str(1, 45, 18) = " `KVKs` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4112 Versichertenstatus VK'"
 Str(1, 45, 19) = " `KVKserg` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4113 Ost/West-Status VK'"
 Str(1, 45, 20) = " `Kasse` varchar(70) COLLATE latin1_german2_ci NOT NULL COMMENT '6299 Kasse (aus Formularen)'"
 Str(1, 45, 21) = " `GebOr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4121, Gebührenordnung (1 = BMÄ, 2)'"
 Str(1, 45, 22) = " `AbrGb` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4122, Abrechnungsgebiet (07 = Diabetes)'"
 Str(1, 45, 23) = " `PersKreis` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4123 Personenkreis/Untersuchungskategorie'"
 Str(1, 45, 24) = " `SKtZusatz` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4124 SKT-Zusatzangaben'"
 Str(1, 45, 25) = " `f4206` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4206, mutmasslicher Tag der Entbindung'"
 Str(1, 45, 26) = " `ÜwText` longtext COLLATE latin1_german2_ci COMMENT '4209: Auftrags- / erläuternder Text zur Überweisung'"
 Str(1, 45, 27) = " `f4210` tinyint(1) unsigned DEFAULT NULL COMMENT '4210, Ankreuzfeld LSR'"
 Str(1, 45, 28) = " `AkfHAH` tinyint(1) unsigned DEFAULT NULL COMMENT '4211 Ankreuzfeld HAH'"
 Str(1, 45, 29) = " `AkfAB0` tinyint(1) unsigned DEFAULT NULL COMMENT '4212 Ankreuzfeld AB0.RH'"
 Str(1, 45, 30) = " `AkfAK` tinyint(1) unsigned DEFAULT NULL COMMENT '4213 Ankreuzfeld AK'"
 Str(1, 45, 31) = " `statNuller` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4216, nu bei Musterfrau 16 Nuller'"
 Str(1, 45, 32) = " `ÜbwV` varchar(22) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4218, überwiesen von Arztnummer'"
 Str(1, 45, 33) = " `AndÜw` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4219, anderer Überweiser'"
 Str(1, 45, 34) = " `Übw` varchar(8) COLLATE latin1_german2_ci NOT NULL COMMENT '4218 oder 4219, je nachdem, was befüllt'"
 Str(1, 45, 35) = " `ÜbwLANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4242 LANR des Überweisers'"
 Str(1, 45, 36) = " `ÜWZiel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4220 Überweisung an'"
 Str(1, 45, 37) = " `ÜWNNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(4): KV-Nummer des Überweisers'"
 Str(1, 45, 38) = " `ÜWNaN` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(3): Nachname des Überweisers'"
 Str(1, 45, 39) = " `ÜWTit` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(3): Titel des Überweisers'"
 Str(1, 45, 40) = " `ÜWVor` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(2): Vorname des Überweisers'"
 Str(1, 45, 41) = " `ÜWVsw` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4231(2b): Vorsatzwort des Überweisers'"
 Str(1, 45, 42) = " `üwvid` int(2) unsigned NOT NULL COMMENT '4247 Bezug auf ueberwvon'"
 Str(1, 45, 43) = " `statKlasse` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4236 Klasse bei Behandlung'"
 Str(1, 45, 44) = " `f4237` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4237 Krankenhausname'"
 Str(1, 45, 45) = " `statBehTage` int(10) DEFAULT NULL COMMENT '4238 Krankenhausaufenthalt'"
 Str(1, 45, 46) = " `SchGr` int(10) DEFAULT NULL COMMENT '4239, Schein(unter)gruppe'"
 Str(1, 45, 47) = " `Weiterbeh` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4243, Weiterbehandelnder'"
 Str(1, 45, 48) = " `PGeb` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4401, Praxisgebühr'"
 Str(1, 45, 49) = " `PGebErg` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4402, ?'"
 Str(1, 45, 50) = " `Mahnfrist` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4403, Mahnfrist bis'"
 Str(1, 45, 51) = " `GOÄKatNr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4580 (1): Katalog-Nummer'"
 Str(1, 45, 52) = " `GOÄKatName` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4580 (2): Privat-Abrechnungskatalog'"
 Str(1, 45, 53) = " `abrArzt` varchar(30) COLLATE latin1_german2_ci NOT NULL COMMENT '4585 abrechnender Arzt'"
 Str(1, 45, 54) = " `privVers` varchar(45) COLLATE latin1_german2_ci NOT NULL COMMENT '4586 private Versicherung'"
 Str(1, 45, 55) = " `AdNam` varchar(28) COLLATE latin1_german2_ci NOT NULL COMMENT '4602(1) Name Rechnungsanschrift'"
 Str(1, 45, 56) = " `AdStr` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(2) Straße Rechnungsanschrift'"
 Str(1, 45, 57) = " `AdPlz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(3) PLZ Rechnungsanschrift'"
 Str(1, 45, 58) = " `AdOrt` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '4602(4) Ort Rechnungsanschrift'"
 Str(1, 45, 59) = " `BhFE` datetime DEFAULT NULL COMMENT '4604, Behandlungsfall: Ende, bei Privatpatienten'"
 Str(1, 45, 60) = " `s8000` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000, Satzidentifikation'"
 Str(1, 45, 61) = " `s8100` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge'"
 Str(1, 45, 62) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 45, 63) = " `Fanf` datetime DEFAULT NULL COMMENT 'Fallanfang'"
 Str(1, 45, 64) = " `altQuart` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 45, 65) = " `QAnf` datetime DEFAULT NULL COMMENT 'Quartalsanfang'"
 Str(1, 45, 66) = " `QEnd` datetime DEFAULT NULL COMMENT 'Quartalsende'"
 Str(1, 45, 67) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 45, 68) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 45, 69) = " `TherArt` int(2) unsigned DEFAULT NULL COMMENT 'Therapieart: (0 = offen,  1= diät,  2= oad, 3= komb,  4= ct, 5= ict, 6 = csii)'"
 Str(1, 45, 70) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 45, 71) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 45, 72) = "  PRIMARY KEY (`FID`)"
 Str(1, 45, 73) = "  UNIQUE KEY `PrimaryKey` (`FID`)"
 Str(1, 45, 74) = "  KEY `AktF` (`Pat_ID`,`BhFB`)"
 Str(1, 45, 75) = "  KEY `Auswahl` (`Pat_ID`,`Quartal`,`BhFB`,`BhFE1`)"
 Str(1, 45, 76) = "  KEY `BhFB` (`BhFB`)"
 Str(1, 45, 77) = "  KEY `FanF` (`Fanf`)"
 Str(1, 45, 78) = "  KEY `NamenFälle` (`Pat_ID`)"
 Str(1, 45, 79) = "  KEY `pQ` (`Pat_ID`,`Quartal`)"
 Str(1, 45, 80) = "  KEY `Quartal` (`Quartal`)"
 Str(1, 45, 81) = "  KEY `SchGr` (`SchGr`,`Nachname`,`Vorname`)"
 Str(1, 45, 82) = "  KEY `vknr` (`VKNr`)"
 Str(1, 45, 83) = "  CONSTRAINT `KassenlisteFälle_AccRel` FOREIGN KEY (`VKNr`) REFERENCES `kassenliste` (`VK`)"
 Str(1, 45, 84) = "  CONSTRAINT `NamenFälle_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 45, 85) = " ENGINE=InnoDB AUTO_INCREMENT=18957 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr45

Sub FüllStr46()
 Str(0, 46, 0) = "faelleverschieden"
 Str(1, 46, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `faelleverschieden` AS select `f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`f4206` AS `f4206`,`f`.`ÜwText` AS `ÜwText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`ÜbwV` AS `ÜbwV`,`f" & _ 
  "`.`AndÜw` AS `AndÜw`,`f`.`Übw` AS `Übw`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8100` AS `s8100`,`f`.`AktZeit` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`TherArt` AS" & _ 
  " `TherArt`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos` from `_faellenachschgr` `f` group by `f`.`Pat_ID`,`f`.`Quartal`"
End Sub ' FüllStr46

Sub FüllStr47()
 Str(0, 47, 0) = "faelleverschiedenneu"
 Str(1, 47, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `faelleverschiedenneu` AS select `f`.`FID` AS `FID`,`f`.`Pat_ID` AS `Pat_ID`,`f`.`Quartal` AS `Quartal`,`f`.`Nachname` AS `Nachname`,`f`.`Vorname` AS `Vorname`,`f`.`lfdnr` AS `lfdnr`,`f`.`TMFNr` AS `TMFNr`,`f`.`VKNr` AS `VKNr`,`f`.`BhFB` AS `BhFB`,`f`.`BhFE1` AS `BhFE1`,`f`.`BhFE2` AS `BhFE2`,`f`.`f4202` AS `f4202`,`f`.`ausgst` AS `ausgst`,`f`.`KtrAbrB` AS `KtrAbrB`,`f`.`AbrAr` AS `AbrAr`,`f`.`lVorl` AS `lVorl`,`f`.`IK` AS `IK`,`f`.`KVKs` AS `KVKs`,`f`.`KVKserg` AS `KVKserg`,`f`.`Kasse` AS `Kasse`,`f`.`GebOr` AS `GebOr`,`f`.`AbrGb` AS `AbrGb`,`f`.`PersKreis` AS `PersKreis`,`f`.`SKtZusatz` AS `SKtZusatz`,`f`.`f4206` AS `f4206`,`f`.`ÜwText` AS `ÜwText`,`f`.`f4210` AS `f4210`,`f`.`AkfHAH` AS `AkfHAH`,`f`.`AkfAB0` AS `AkfAB0`,`f`.`AkfAK` AS `AkfAK`,`f`.`statNuller` AS `statNuller`,`f`.`ÜbwV` AS `ÜbwV`" & _ 
  ",`f`.`AndÜw` AS `AndÜw`,`f`.`Übw` AS `Übw`,`f`.`ÜbwLANR` AS `ÜbwLANR`,`f`.`ÜWZiel` AS `ÜWZiel`,`f`.`ÜWNNr` AS `ÜWNNr`,`f`.`ÜWNaN` AS `ÜWNaN`,`f`.`ÜWTit` AS `ÜWTit`,`f`.`ÜWVor` AS `ÜWVor`,`f`.`ÜWVsw` AS `ÜWVsw`,`f`.`üwvid` AS `üwvid`,`f`.`statKlasse` AS `statKlasse`,`f`.`f4237` AS `f4237`,`f`.`statBehTage` AS `statBehTage`,`f`.`SchGr` AS `SchGr`,`f`.`Weiterbeh` AS `Weiterbeh`,`f`.`PGeb` AS `PGeb`,`f`.`PGebErg` AS `PGebErg`,`f`.`Mahnfrist` AS `Mahnfrist`,`f`.`GOÄKatNr` AS `GOÄKatNr`,`f`.`GOÄKatName` AS `GOÄKatName`,`f`.`abrArzt` AS `abrArzt`,`f`.`privVers` AS `privVers`,`f`.`AdNam` AS `AdNam`,`f`.`AdStr` AS `AdStr`,`f`.`AdPlz` AS `AdPlz`,`f`.`AdOrt` AS `AdOrt`,`f`.`BhFE` AS `BhFE`,`f`.`s8000` AS `s8000`,`f`.`s8100` AS `s8100`,`f`.`AktZeit` AS `AktZeit`,`f`.`Fanf` AS `Fanf`,`f`.`altQuart` AS `altQuart`,`f`.`QAnf` AS `QAnf`,`f`.`QEnd` AS `QEnd`,`f`.`QS` AS `QS`,`f`.`QT` AS `QT`,`f`.`TherArt`" & _ 
  " AS `TherArt`,`f`.`StByte` AS `StByte`,`f`.`absPos` AS `absPos`,(select min(`f1`.`Fanf`) AS `min(fanf)` from `faelle` `f1` where ((`f1`.`Pat_ID` = `f`.`Pat_ID`) and (`f1`.`Fanf` < `f`.`Fanf`))) AS `erst` from `_faellenachschgr` `f` group by `f`.`Pat_ID`,`f`.`Quartal`"
End Sub ' FüllStr47

Sub FüllStr48()
 Str(0, 48, 0) = "fallzahlen"
 Str(0, 48, 1) = "`Zahl`"
 Str(0, 48, 2) = "`DZahl`"
 Str(0, 48, 3) = "`NDZahl`"
 Str(0, 48, 4) = "`Quartal`"
 ArtZ(0, 48) = 4
 Str(1, 48, 0) = "CREATE TABLE `fallzahlen` ("
 Str(1, 48, 1) = " `Zahl` decimal(15,4) DEFAULT NULL"
 Str(1, 48, 2) = " `DZahl` int(10) DEFAULT NULL"
 Str(1, 48, 3) = " `NDZahl` int(10) DEFAULT NULL"
 Str(1, 48, 4) = " `Quartal` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 48, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr48

Sub FüllStr49()
 Str(0, 49, 0) = "fallzahlstand"
 Str(0, 49, 1) = "`zahl`"
 Str(0, 49, 2) = "`z0`"
 Str(0, 49, 3) = "`quartal`"
 ArtZ(0, 49) = 3
 Str(1, 49, 0) = "CREATE TABLE `fallzahlstand` ("
 Str(1, 49, 1) = " `zahl` bigint(21) DEFAULT NULL"
 Str(1, 49, 2) = " `z0` bigint(22) DEFAULT NULL"
 Str(1, 49, 3) = " `quartal` varchar(5) DEFAULT NULL"
 Str(1, 49, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr49

Sub FüllStr50()
 Str(0, 50, 0) = "fallzahlstand 1"
 Str(1, 50, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 1` AS select count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` from `faelleverschiedenneu` `f` where ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 1 day)) - to_days(concat(year((now() - interval 1 day)),'-',((((month((now() - interval 1 day)) - 1) DIV 3) * 3) + 1),'-01')))) and ((`f`.`Pat_ID` < 3044) or (`f`.`Pat_ID` > 50000))) group by `f`.`Quartal` order by substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Sub ' FüllStr50

Sub FüllStr51()
 Str(0, 51, 0) = "fallzahlstand 2"
 Str(1, 51, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 2` AS select count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` from `faelleverschiedenneu` `f` where ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 1 day)) - to_days(concat(year((now() - interval 1 day)),'-',((((month((now() - interval 1 day)) - 1) DIV 3) * 3) + 1),'-01')))) and (`f`.`Pat_ID` > 3044)) group by `f`.`Quartal` order by substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Sub ' FüllStr51

Sub FüllStr52()
 Str(0, 52, 0) = "fallzahlstand 3"
 Str(1, 52, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `fallzahlstand 3` AS select count(0) AS `zahl`,(count(0) - count(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` from `faelleverschiedenneu` `f` where ((`f`.`SchGr` <> '90') and ((to_days(`f`.`Fanf`) - to_days(concat(substr(`f`.`Quartal`,2,4),'-',(((left(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) between 0 and (to_days((now() - interval 1 day)) - to_days(concat(year((now() - interval 1 day)),'-',((((month((now() - interval 1 day)) - 1) DIV 3) * 3) + 1),'-01')))) and 1) group by `f`.`Quartal` order by substr(`f`.`Quartal`,2,4),left(`f`.`Quartal`,1)"
End Sub ' FüllStr52

Sub FüllStr53()
 Str(0, 53, 0) = "falsche dokpfade"
 Str(0, 53, 1) = "`pat_id`"
 Str(0, 53, 2) = "`dokpfad`"
 Str(0, 53, 3) = "`aktzeit`"
 ArtZ(0, 53) = 3
 Str(1, 53, 0) = "CREATE TABLE `falsche dokpfade` ("
 Str(1, 53, 1) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 53, 2) = " `dokpfad` varchar(150) DEFAULT NULL"
 Str(1, 53, 3) = " `aktzeit` datetime DEFAULT NULL"
 Str(1, 53, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr53

Sub FüllStr54()
 Str(0, 54, 0) = "fehlende dmp-notizen"
 Str(0, 54, 1) = "`pat_id`"
 Str(0, 54, 2) = "`fid`"
 Str(0, 54, 3) = "`quartal`"
 Str(0, 54, 4) = "`notiz`"
 Str(0, 54, 5) = "`icd`"
 ArtZ(0, 54) = 5
 Str(1, 54, 0) = "CREATE TABLE `fehlende dmp-notizen` ("
 Str(1, 54, 1) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 54, 2) = " `fid` int(10) DEFAULT NULL"
 Str(1, 54, 3) = " `quartal` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 54, 4) = " `notiz` varchar(150) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 54, 5) = " `icd` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 54, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr54

Sub FüllStr55()
 Str(0, 55, 0) = "fehlerversichertenpauschale"
 Str(0, 55, 1) = "`schgr`"
 Str(0, 55, 2) = "`verspau`"
 Str(0, 55, 3) = "`icd`"
 Str(0, 55, 4) = "`fid`"
 Str(0, 55, 5) = "`pat_id`"
 Str(0, 55, 6) = "`quartal`"
 Str(0, 55, 7) = "`nachname`"
 Str(0, 55, 8) = "`vorname`"
 ArtZ(0, 55) = 8
 Str(1, 55, 0) = "CREATE TABLE `fehlerversichertenpauschale` ("
 Str(1, 55, 1) = " `schgr` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 55, 2) = " `verspau` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 55, 3) = " `icd` varchar(10) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 55, 4) = " `fid` int(10) DEFAULT NULL"
 Str(1, 55, 5) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 55, 6) = " `quartal` varchar(5) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 55, 7) = " `nachname` varchar(50) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 55, 8) = " `vorname` varchar(50) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 55, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr55

Sub FüllStr56()
 Str(0, 56, 0) = "feldlaengen"
 Str(0, 56, 1) = "`TabName`"
 Str(0, 56, 2) = "`FeldName`"
 Str(0, 56, 3) = "`Typ`"
 Str(0, 56, 4) = "`Verkn`"
 Str(0, 56, 5) = "`Größe`"
 Str(0, 56, 6) = "`MaxFüll`"
 Str(0, 56, 7) = "`Diff`"
 Str(0, 56, 8) = "`Zahl`"
 Str(0, 56, 9) = "`ProdLeer`"
 Str(0, 56, 10) = "`ProdDaten`"
 Str(0, 56, 11) = "`ProdCont`"
 ArtZ(0, 56) = 11
 Str(1, 56, 0) = "CREATE TABLE `feldlaengen` ("
 Str(1, 56, 1) = " `TabName` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 56, 2) = " `FeldName` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 56, 3) = " `Typ` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 56, 4) = " `Verkn` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 56, 5) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 56, 6) = " `MaxFüll` int(10) DEFAULT NULL"
 Str(1, 56, 7) = " `Diff` int(10) DEFAULT NULL"
 Str(1, 56, 8) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 56, 9) = " `ProdLeer` int(10) DEFAULT NULL"
 Str(1, 56, 10) = " `ProdDaten` int(10) DEFAULT NULL"
 Str(1, 56, 11) = " `ProdCont` int(10) DEFAULT NULL"
 Str(1, 56, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr56

Sub FüllStr57()
 Str(0, 57, 0) = "feldprop"
 Str(0, 57, 1) = "`Prop`"
 Str(0, 57, 2) = "`ReiheVorher`"
 Str(0, 57, 3) = "`TypVorher`"
 Str(0, 57, 4) = "`WertVorher`"
 Str(0, 57, 5) = "`ReiheNachher`"
 Str(0, 57, 6) = "`TypNachher`"
 Str(0, 57, 7) = "`WertNachher`"
 Str(0, 57, 8) = "`prop`"
 ArtZ(0, 57) = 7
 ArtZ(1, 57) = 1
 Str(1, 57, 0) = "CREATE TABLE `feldprop` ("
 Str(1, 57, 1) = " `Prop` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 57, 2) = " `ReiheVorher` int(10) DEFAULT NULL"
 Str(1, 57, 3) = " `TypVorher` int(10) DEFAULT NULL"
 Str(1, 57, 4) = " `WertVorher` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 57, 5) = " `ReiheNachher` int(10) DEFAULT NULL"
 Str(1, 57, 6) = " `TypNachher` int(10) DEFAULT NULL"
 Str(1, 57, 7) = " `WertNachher` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 57, 8) = "  KEY `prop` (`Prop`)"
 Str(1, 57, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr57

Sub FüllStr58()
 Str(0, 58, 0) = "forminhaltfeld"
 Str(0, 58, 1) = "`FeldVW`"
 Str(0, 58, 2) = "`Feld`"
 Str(0, 58, 3) = "`StByte`"
 Str(0, 58, 4) = "`FeldVW`"
 Str(0, 58, 5) = "`Feld`"
 ArtZ(0, 58) = 3
 ArtZ(1, 58) = 2
 Str(1, 58, 0) = "CREATE TABLE `forminhaltfeld` ("
 Str(1, 58, 1) = " `FeldVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 58, 2) = " `Feld` longtext COLLATE latin1_german2_ci"
 Str(1, 58, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 58, 4) = "  PRIMARY KEY (`FeldVW`)"
 Str(1, 58, 5) = "  KEY `Feld` (`Feld`(255))"
 Str(1, 58, 6) = " ENGINE=InnoDB AUTO_INCREMENT=462 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr58

Sub FüllStr59()
 Str(0, 59, 0) = "forminhaltfeldinh"
 Str(0, 59, 1) = "`FeldInhVW`"
 Str(0, 59, 2) = "`FeldInh`"
 Str(0, 59, 3) = "`StByte`"
 Str(0, 59, 4) = "`FeldInhVW`"
 Str(0, 59, 5) = "`FeldInhVW`"
 Str(0, 59, 6) = "`FeldInh`"
 ArtZ(0, 59) = 3
 ArtZ(1, 59) = 3
 Str(1, 59, 0) = "CREATE TABLE `forminhaltfeldinh` ("
 Str(1, 59, 1) = " `FeldInhVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 59, 2) = " `FeldInh` longtext COLLATE latin1_german2_ci"
 Str(1, 59, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 59, 4) = "  PRIMARY KEY (`FeldInhVW`)"
 Str(1, 59, 5) = "  UNIQUE KEY `FeldInhVW` (`FeldInhVW`)"
 Str(1, 59, 6) = "  KEY `FeldInh` (`FeldInh`(255))"
 Str(1, 59, 7) = " ENGINE=InnoDB AUTO_INCREMENT=51471 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr59

Sub FüllStr60()
 Str(0, 60, 0) = "forminhaltform_abk"
 Str(0, 60, 1) = "`Form_AbkVW`"
 Str(0, 60, 2) = "`Form_Abk`"
 Str(0, 60, 3) = "`Form_AbkVW`"
 Str(0, 60, 4) = "`Form_AbkVW`"
 Str(0, 60, 5) = "`Form_Abk`"
 ArtZ(0, 60) = 2
 ArtZ(1, 60) = 3
 Str(1, 60, 0) = "CREATE TABLE `forminhaltform_abk` ("
 Str(1, 60, 1) = " `Form_AbkVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 60, 2) = " `Form_Abk` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 60, 3) = "  PRIMARY KEY (`Form_AbkVW`)"
 Str(1, 60, 4) = "  UNIQUE KEY `Form_AbkVW` (`Form_AbkVW`)"
 Str(1, 60, 5) = "  UNIQUE KEY `Form_Abk` (`Form_Abk`)"
 Str(1, 60, 6) = " ENGINE=InnoDB AUTO_INCREMENT=47 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr60

Sub FüllStr61()
 Str(0, 61, 0) = "forminhfeld"
 Str(0, 61, 1) = "`FoID`"
 Str(0, 61, 2) = "`Nr`"
 Str(0, 61, 3) = "`FeldNr`"
 Str(0, 61, 4) = "`FeldVW`"
 Str(0, 61, 5) = "`FeldInhVW`"
 Str(0, 61, 6) = "`FormInhKopfFormInhFeld`"
 Str(0, 61, 7) = "`FoID`"
 Str(0, 61, 8) = "`FormInhaltFeldFormInhFeld`"
 Str(0, 61, 9) = "`FormInhaltFeldInhFormInhFeld`"
 Str(0, 61, 10) = "`FormInhaltFeldFormInhFeld_AccRel`"
 Str(0, 61, 11) = "`FormInhaltFeldInhFormInhFeld_AccRel`"
 Str(0, 61, 12) = "`FormInhFeldFormInhKopf_AccRel`"
 ArtZ(0, 61) = 5
 ArtZ(1, 61) = 4
 ArtZ(2, 61) = 3
 Str(1, 61, 0) = "CREATE TABLE `forminhfeld` ("
 Str(1, 61, 1) = " `FoID` int(10) DEFAULT NULL"
 Str(1, 61, 2) = " `Nr` smallint(6) DEFAULT NULL"
 Str(1, 61, 3) = " `FeldNr` smallint(6) DEFAULT NULL"
 Str(1, 61, 4) = " `FeldVW` int(10) DEFAULT NULL"
 Str(1, 61, 5) = " `FeldInhVW` int(10) DEFAULT NULL"
 Str(1, 61, 6) = "  KEY `FormInhKopfFormInhFeld` (`FoID`)"
 Str(1, 61, 7) = "  KEY `FoID` (`FoID`,`FeldVW`,`FeldNr`)"
 Str(1, 61, 8) = "  KEY `FormInhaltFeldFormInhFeld` (`FeldVW`)"
 Str(1, 61, 9) = "  KEY `FormInhaltFeldInhFormInhFeld` (`FeldInhVW`)"
 Str(1, 61, 10) = "  CONSTRAINT `FormInhaltFeldFormInhFeld_AccRel` FOREIGN KEY (`FeldVW`) REFERENCES `forminhaltfeld` (`FeldVW`) ON UPDATE CASCADE"
 Str(1, 61, 11) = "  CONSTRAINT `FormInhaltFeldInhFormInhFeld_AccRel` FOREIGN KEY (`FeldInhVW`) REFERENCES `forminhaltfeldinh` (`FeldInhVW`) ON UPDATE CASCADE"
 Str(1, 61, 12) = "  CONSTRAINT `FormInhFeldFormInhKopf_AccRel` FOREIGN KEY (`FoID`) REFERENCES `forminhkopf` (`FoID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 61, 13) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr61

Sub FüllStr62()
 Str(0, 62, 0) = "forminhkopf"
 Str(0, 62, 1) = "`FoID`"
 Str(0, 62, 2) = "`FID`"
 Str(0, 62, 3) = "`Pat_ID`"
 Str(0, 62, 4) = "`Form_ID`"
 Str(0, 62, 5) = "`ZeitPunkt`"
 Str(0, 62, 6) = "`AbsPos`"
 Str(0, 62, 7) = "`AktZeit`"
 Str(0, 62, 8) = "`StByte`"
 Str(0, 62, 9) = "`Satzart`"
 Str(0, 62, 10) = "`Satzlänge`"
 Str(0, 62, 11) = "`FoID`"
 Str(0, 62, 12) = "`PrimaryKey`"
 Str(0, 62, 13) = "`Auswahl`"
 Str(0, 62, 14) = "`FälleFormInhKopf`"
 Str(0, 62, 15) = "`FID`"
 Str(0, 62, 16) = "`FormulareFormInhKopf`"
 Str(0, 62, 17) = "`NamenFormInhKopf`"
 Str(0, 62, 18) = "`FälleFormInhKopf_AccRel`"
 Str(0, 62, 19) = "`FormulareFormInhKopf_AccRel`"
 Str(0, 62, 20) = "`NamenFormInhKopf_AccRel`"
 ArtZ(0, 62) = 10
 ArtZ(1, 62) = 7
 ArtZ(2, 62) = 3
 Str(1, 62, 0) = "CREATE TABLE `forminhkopf` ("
 Str(1, 62, 1) = " `FoID` int(10) NOT NULL"
 Str(1, 62, 2) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 62, 3) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 62, 4) = " `Form_ID` int(10) DEFAULT NULL"
 Str(1, 62, 5) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 62, 6) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 62, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 62, 8) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 62, 9) = " `Satzart` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000'"
 Str(1, 62, 10) = " `Satzlänge` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100'"
 Str(1, 62, 11) = "  PRIMARY KEY (`FoID`)"
 Str(1, 62, 12) = "  UNIQUE KEY `PrimaryKey` (`FoID`)"
 Str(1, 62, 13) = "  KEY `Auswahl` (`Pat_ID`,`Form_ID`,`ZeitPunkt`)"
 Str(1, 62, 14) = "  KEY `FälleFormInhKopf` (`FID`)"
 Str(1, 62, 15) = "  KEY `FID` (`FID`)"
 Str(1, 62, 16) = "  KEY `FormulareFormInhKopf` (`Form_ID`)"
 Str(1, 62, 17) = "  KEY `NamenFormInhKopf` (`Pat_ID`)"
 Str(1, 62, 18) = "  CONSTRAINT `FälleFormInhKopf_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 62, 19) = "  CONSTRAINT `FormulareFormInhKopf_AccRel` FOREIGN KEY (`Form_ID`) REFERENCES `formulare` (`FormID`) ON UPDATE CASCADE"
 Str(1, 62, 20) = "  CONSTRAINT `NamenFormInhKopf_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 62, 21) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr62

Sub FüllStr63()
 Str(0, 63, 0) = "formular"
 Str(1, 63, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `formular` AS select `forminhkopf`.`FoID` AS `foid`,`forminhkopf`.`Pat_ID` AS `Pat_ID`,`forminhkopf`.`FID` AS `FID`,`forminhkopf`.`Form_ID` AS `Form_ID`,`forminhkopf`.`ZeitPunkt` AS `ZeitPunkt`,`forminhfeld`.`Nr` AS `Nr`,`forminhfeld`.`FeldNr` AS `FeldNr`,`forminhaltfeld`.`Feld` AS `Feld`,`forminhaltfeldinh`.`FeldInh` AS `FeldInh`,`formulare`.`Form_Abk` AS `form_abk`,`formulare`.`FormVorl` AS `FormVorl` from ((((`forminhfeld` left join `forminhkopf` on((`forminhfeld`.`FoID` = `forminhkopf`.`FoID`))) left join `formulare` on((`formulare`.`FormID` = `forminhkopf`.`Form_ID`))) left join `forminhaltfeld` on((`forminhfeld`.`FeldVW` = `forminhaltfeld`.`FeldVW`))) left join `forminhaltfeldinh` on((`forminhfeld`.`FeldInhVW` = `forminhaltfeldinh`.`FeldInhVW`))) order by `forminhkopf`.`FoID`"
End Sub ' FüllStr63

Sub FüllStr64()
 Str(0, 64, 0) = "formulare"
 Str(0, 64, 1) = "`FormID`"
 Str(0, 64, 2) = "`Form_Abk`"
 Str(0, 64, 3) = "`FormBez`"
 Str(0, 64, 4) = "`FormVorl`"
 Str(0, 64, 5) = "`AktZeit`"
 Str(0, 64, 6) = "`absPos`"
 Str(0, 64, 7) = "`StByte`"
 Str(0, 64, 8) = "`FormID`"
 Str(0, 64, 9) = "`FormID`"
 Str(0, 64, 10) = "`Auswahl`"
 Str(0, 64, 11) = "`FormBez`"
 Str(0, 64, 12) = "`FormInhaltForm_AbkFormulare`"
 Str(0, 64, 13) = "`FormInhaltForm_AbkFormulare_AccRel`"
 ArtZ(0, 64) = 7
 ArtZ(1, 64) = 5
 ArtZ(2, 64) = 1
 Str(1, 64, 0) = "CREATE TABLE `formulare` ("
 Str(1, 64, 1) = " `FormID` int(10) NOT NULL"
 Str(1, 64, 2) = " `Form_Abk` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 3) = " `FormBez` longtext COLLATE latin1_german2_ci"
 Str(1, 64, 4) = " `FormVorl` varchar(114) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 64, 5) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Aktualisierung'"
 Str(1, 64, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 64, 7) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 64, 8) = "  PRIMARY KEY (`FormID`)"
 Str(1, 64, 9) = "  UNIQUE KEY `FormID` (`FormID`)"
 Str(1, 64, 10) = "  KEY `Auswahl` (`Form_Abk`,`FormBez`(255),`FormVorl`)"
 Str(1, 64, 11) = "  KEY `FormBez` (`FormBez`(255))"
 Str(1, 64, 12) = "  KEY `FormInhaltForm_AbkFormulare` (`Form_Abk`)"
 Str(1, 64, 13) = "  CONSTRAINT `FormInhaltForm_AbkFormulare_AccRel` FOREIGN KEY (`Form_Abk`) REFERENCES `forminhaltform_abk` (`Form_Abk`)"
 Str(1, 64, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr64

Sub FüllStr65()
 Str(0, 65, 0) = "fremdlabor"
 Str(0, 65, 1) = "`ID`"
 Str(0, 65, 2) = "`Pat_ID`"
 Str(0, 65, 3) = "`Datum`"
 Str(0, 65, 4) = "`HbA1c`"
 Str(0, 65, 5) = "`Glucose`"
 Str(0, 65, 6) = "`Crea`"
 Str(0, 65, 7) = "`Hst`"
 Str(0, 65, 8) = "`GGT`"
 Str(0, 65, 9) = "`GPT`"
 Str(0, 65, 10) = "`INR`"
 Str(0, 65, 11) = "`Leuko`"
 Str(0, 65, 12) = "`Hb`"
 Str(0, 65, 13) = "`MCH`"
 Str(0, 65, 14) = "`Thrombo`"
 Str(0, 65, 15) = "`CRP`"
 Str(0, 65, 16) = "`K`"
 Str(0, 65, 17) = "`Na`"
 Str(0, 65, 18) = "`TSH`"
 Str(0, 65, 19) = "`fT4`"
 Str(0, 65, 20) = "`fT3`"
 Str(0, 65, 21) = "`Chol`"
 Str(0, 65, 22) = "`LDL`"
 Str(0, 65, 23) = "`HDL`"
 Str(0, 65, 24) = "`TG`"
 Str(0, 65, 25) = "`Hsre`"
 Str(0, 65, 26) = "`LeuU`"
 Str(0, 65, 27) = "`EryU`"
 Str(0, 65, 28) = "`EiwU`"
 Str(0, 65, 29) = "`BakU`"
 Str(0, 65, 30) = "`NitU`"
 Str(0, 65, 31) = "`BiliU`"
 Str(0, 65, 32) = "`UrobU`"
 Str(0, 65, 33) = "`SedU`"
 Str(0, 65, 34) = "`KetU`"
 Str(0, 65, 35) = "`ZucU`"
 Str(0, 65, 36) = "`AlbU`"
 Str(0, 65, 37) = "`GEiwU`"
 Str(0, 65, 38) = "`Micral`"
 Str(0, 65, 39) = "`DokName`"
 Str(0, 65, 40) = "`DokPfad`"
 Str(0, 65, 41) = "`DokGroe`"
 Str(0, 65, 42) = "`ID`"
 Str(0, 65, 43) = "`PrimaryKey`"
 Str(0, 65, 44) = "`Auswahl`"
 Str(0, 65, 45) = "`Pat_ID`"
 ArtZ(0, 65) = 41
 ArtZ(1, 65) = 4
 Str(1, 65, 0) = "CREATE TABLE `fremdlabor` ("
 Str(1, 65, 1) = " `ID` int(10) NOT NULL"
 Str(1, 65, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 65, 3) = " `Datum` datetime DEFAULT NULL"
 Str(1, 65, 4) = " `HbA1c` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 5) = " `Glucose` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 6) = " `Crea` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 7) = " `Hst` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 8) = " `GGT` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 9) = " `GPT` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 10) = " `INR` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 11) = " `Leuko` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 12) = " `Hb` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 13) = " `MCH` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 14) = " `Thrombo` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 15) = " `CRP` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 16) = " `K` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 17) = " `Na` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 18) = " `TSH` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 19) = " `fT4` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 20) = " `fT3` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 21) = " `Chol` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 22) = " `LDL` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 23) = " `HDL` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 24) = " `TG` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 25) = " `Hsre` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 26) = " `LeuU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 27) = " `EryU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 28) = " `EiwU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 29) = " `BakU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 30) = " `NitU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 31) = " `BiliU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 32) = " `UrobU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 33) = " `SedU` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 34) = " `KetU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 35) = " `ZucU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 36) = " `AlbU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 37) = " `GEiwU` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 38) = " `Micral` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 65, 39) = " `DokName` longtext COLLATE latin1_german2_ci"
 Str(1, 65, 40) = " `DokPfad` longtext COLLATE latin1_german2_ci"
 Str(1, 65, 41) = " `DokGroe` int(10) DEFAULT NULL COMMENT 'Größe in Bytes'"
 Str(1, 65, 42) = "  PRIMARY KEY (`ID`)"
 Str(1, 65, 43) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 65, 44) = "  KEY `Auswahl` (`Pat_ID`,`DokPfad`(20))"
 Str(1, 65, 45) = "  KEY `Pat_ID` (`Pat_ID`)"
 Str(1, 65, 46) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr65

Sub FüllStr66()
 Str(0, 66, 0) = "fuerdiagexp"
 Str(0, 66, 1) = "`ID`"
 Str(0, 66, 2) = "`Name`"
 Str(0, 66, 3) = "`Pat_id`"
 Str(0, 66, 4) = "`ICD`"
 Str(0, 66, 5) = "`Diagnose`"
 Str(0, 66, 6) = "`Status`"
 Str(0, 66, 7) = "`Protokoll`"
 Str(0, 66, 8) = "`nurQuart`"
 Str(0, 66, 9) = "`Zeitpunkt`"
 Str(0, 66, 10) = "`ID`"
 Str(0, 66, 11) = "`ID`"
 Str(0, 66, 12) = "`pat_ID`"
 Str(0, 66, 13) = "`Suche`"
 ArtZ(0, 66) = 9
 ArtZ(1, 66) = 4
 Str(1, 66, 0) = "CREATE TABLE `fuerdiagexp` ("
 Str(1, 66, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 66, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 66, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 66, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 66, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 66, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 66, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 66, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 66, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 66, 10) = "  PRIMARY KEY (`ID`)"
 Str(1, 66, 11) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 66, 12) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 66, 13) = "  KEY `Suche` (`Pat_id`,`ICD`,`Diagnose`)"
 Str(1, 66, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr66

Sub FüllStr67()
 Str(0, 67, 0) = "fuerdiagexp kopie"
 Str(0, 67, 1) = "`ID`"
 Str(0, 67, 2) = "`Name`"
 Str(0, 67, 3) = "`Pat_id`"
 Str(0, 67, 4) = "`ICD`"
 Str(0, 67, 5) = "`Diagnose`"
 Str(0, 67, 6) = "`Status`"
 Str(0, 67, 7) = "`Protokoll`"
 Str(0, 67, 8) = "`nurQuart`"
 Str(0, 67, 9) = "`Zeitpunkt`"
 Str(0, 67, 10) = "`ID`"
 Str(0, 67, 11) = "`ID`"
 Str(0, 67, 12) = "`pat_ID`"
 Str(0, 67, 13) = "`Suche`"
 ArtZ(0, 67) = 9
 ArtZ(1, 67) = 4
 Str(1, 67, 0) = "CREATE TABLE `fuerdiagexp kopie` ("
 Str(1, 67, 1) = " `ID` int(10) NOT NULL COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 67, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 67, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 67, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 67, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 67, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 67, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 67, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 67, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 67, 10) = "  PRIMARY KEY (`ID`)"
 Str(1, 67, 11) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 67, 12) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 67, 13) = "  KEY `Suche` (`Pat_id`,`ICD`,`Diagnose`)"
 Str(1, 67, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr67

Sub FüllStr68()
 Str(0, 68, 0) = "fuerdiagexparchiv"
 Str(0, 68, 1) = "`ID`"
 Str(0, 68, 2) = "`Name`"
 Str(0, 68, 3) = "`Pat_id`"
 Str(0, 68, 4) = "`ICD`"
 Str(0, 68, 5) = "`Diagnose`"
 Str(0, 68, 6) = "`Status`"
 Str(0, 68, 7) = "`Protokoll`"
 Str(0, 68, 8) = "`nurQuart`"
 Str(0, 68, 9) = "`Zeitpunkt`"
 Str(0, 68, 10) = "`archiviert`"
 Str(0, 68, 11) = "`ID`"
 Str(0, 68, 12) = "`ID`"
 Str(0, 68, 13) = "`pat_ID`"
 Str(0, 68, 14) = "`Suche`"
 ArtZ(0, 68) = 10
 ArtZ(1, 68) = 4
 Str(1, 68, 0) = "CREATE TABLE `fuerdiagexparchiv` ("
 Str(1, 68, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 68, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 68, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 68, 4) = " `ICD` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'ICD 10'"
 Str(1, 68, 5) = " `Diagnose` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Diagnose Text'"
 Str(1, 68, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 68, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 68, 8) = " `nurQuart` tinyint(1) unsigned DEFAULT NULL COMMENT 'ja = nur für ein Quartal'"
 Str(1, 68, 9) = " `Zeitpunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt der gewünschten Diagnose'"
 Str(1, 68, 10) = " `archiviert` datetime DEFAULT NULL COMMENT 'Datum der Übertragung in Tabelle Leistungsexport'"
 Str(1, 68, 11) = "  PRIMARY KEY (`ID`)"
 Str(1, 68, 12) = "  UNIQUE KEY `ID` (`ID`)"
 Str(1, 68, 13) = "  KEY `pat_ID` (`Pat_id`)"
 Str(1, 68, 14) = "  KEY `Suche` (`Pat_id`,`ICD`,`Diagnose`)"
 Str(1, 68, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr68

Sub FüllStr69()
 Str(0, 69, 0) = "fuerleistexp"
 Str(0, 69, 1) = "`ID`"
 Str(0, 69, 2) = "`Name`"
 Str(0, 69, 3) = "`Datum`"
 Str(0, 69, 4) = "`Pat_id`"
 Str(0, 69, 5) = "`SchGr`"
 Str(0, 69, 6) = "`Status`"
 Str(0, 69, 7) = "`Protokoll`"
 Str(0, 69, 8) = "`ID`"
 Str(0, 69, 9) = "`PrimaryKey`"
 Str(0, 69, 10) = "`ID`"
 Str(0, 69, 11) = "`NamenfürLeistExp`"
 Str(0, 69, 12) = "`NamenfürLeistExp_AccRel`"
 ArtZ(0, 69) = 7
 ArtZ(1, 69) = 4
 ArtZ(2, 69) = 1
 Str(1, 69, 0) = "CREATE TABLE `fuerleistexp` ("
 Str(1, 69, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 69, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 69, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen'"
 Str(1, 69, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 69, 5) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 69, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 69, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 69, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 69, 9) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 69, 10) = "  KEY `ID` (`Pat_id`)"
 Str(1, 69, 11) = "  KEY `NamenfürLeistExp` (`Pat_id`)"
 Str(1, 69, 12) = "  CONSTRAINT `NamenfürLeistExp_AccRel` FOREIGN KEY (`Pat_id`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 69, 13) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr69

Sub FüllStr70()
 Str(0, 70, 0) = "fuerleistexp kopie"
 Str(0, 70, 1) = "`ID`"
 Str(0, 70, 2) = "`Name`"
 Str(0, 70, 3) = "`Datum`"
 Str(0, 70, 4) = "`Pat_id`"
 Str(0, 70, 5) = "`SchGr`"
 Str(0, 70, 6) = "`Status`"
 Str(0, 70, 7) = "`Protokoll`"
 Str(0, 70, 8) = "`ID`"
 Str(0, 70, 9) = "`PrimaryKey`"
 Str(0, 70, 10) = "`ID`"
 ArtZ(0, 70) = 7
 ArtZ(1, 70) = 3
 Str(1, 70, 0) = "CREATE TABLE `fuerleistexp kopie` ("
 Str(1, 70, 1) = " `ID` int(10) NOT NULL COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 70, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 70, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen'"
 Str(1, 70, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id'"
 Str(1, 70, 5) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 70, 6) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'gefunden, nicht gefunden, nicht eindeutig, nicht im Quartal, / privat'"
 Str(1, 70, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 70, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 70, 9) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 70, 10) = "  KEY `ID` (`Pat_id`)"
 Str(1, 70, 11) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr70

Sub FüllStr71()
 Str(0, 71, 0) = "fuerleistexparchiv"
 Str(0, 71, 1) = "`ID`"
 Str(0, 71, 2) = "`Name`"
 Str(0, 71, 3) = "`Datum`"
 Str(0, 71, 4) = "`Pat_id`"
 Str(0, 71, 5) = "`SchGr`"
 Str(0, 71, 6) = "`archiviert`"
 Str(0, 71, 7) = "`Protokoll`"
 Str(0, 71, 8) = "`ID`"
 Str(0, 71, 9) = "`PrimaryKey`"
 Str(0, 71, 10) = "`ID`"
 ArtZ(0, 71) = 7
 ArtZ(1, 71) = 3
 Str(1, 71, 0) = "CREATE TABLE `fuerleistexparchiv` ("
 Str(1, 71, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Datensatzreihenfolge, sonst bedeutungslos'"
 Str(1, 71, 2) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitext mit Anfangswert(en) für Nach- und Vorname'"
 Str(1, 71, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum, wird bei Folgedatensätzen aus dem vorherigen Feld übernommen'"
 Str(1, 71, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!pat_id, wird vom Programm gesucht'"
 Str(1, 71, 5) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 71, 6) = " `archiviert` datetime DEFAULT NULL COMMENT 'Datum der Übertragung in Tabelle Leistungsexport'"
 Str(1, 71, 7) = " `Protokoll` longtext COLLATE latin1_german2_ci COMMENT 'eingetragene Leistungen'"
 Str(1, 71, 8) = "  PRIMARY KEY (`ID`)"
 Str(1, 71, 9) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 71, 10) = "  KEY `ID` (`Pat_id`)"
 Str(1, 71, 11) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr71

Sub FüllStr72()
 Str(0, 72, 0) = "hausaerzte"
 Str(0, 72, 1) = "`ID`"
 Str(0, 72, 2) = "`Überschrift`"
 Str(0, 72, 3) = "`Name`"
 Str(0, 72, 4) = "`Vorname`"
 Str(0, 72, 5) = "`Nachname`"
 Str(0, 72, 6) = "`Anschrift`"
 Str(0, 72, 7) = "`KVNr`"
 Str(0, 72, 8) = "`Telefon`"
 Str(0, 72, 9) = "`Telefax`"
 Str(0, 72, 10) = "`E_Mail`"
 Str(0, 72, 11) = "`Zulassungsgebiet`"
 Str(0, 72, 12) = "`Arzttyp`"
 Str(0, 72, 13) = "`Gemeinschaftspraxis mit`"
 Str(0, 72, 14) = "`Schwerpunkt`"
 Str(0, 72, 15) = "`Zusatzbezeichnung`"
 Str(0, 72, 16) = "`Bemerkung`"
 Str(0, 72, 17) = "`Beme`"
 Str(0, 72, 18) = "`Sprechstunden`"
 Str(0, 72, 19) = "`von _ bis`"
 Str(0, 72, 20) = "`Internetadressen`"
 Str(0, 72, 21) = "`Behandlung in Fremdsprachen`"
 Str(0, 72, 22) = "`Rollstuhlgerechte Praxis`"
 Str(0, 72, 23) = "`Verkehrsmittel`"
 Str(0, 72, 24) = "`Linie`"
 Str(0, 72, 25) = "`Haltestelle Parkplätze`"
 Str(0, 72, 26) = "`Wegbeschreibung`"
 Str(0, 72, 27) = "`Entfernung zur Praxis`"
 Str(0, 72, 28) = "`Zahl`"
 Str(0, 72, 29) = "`nichtmehr`"
 Str(0, 72, 30) = "`Titel`"
 Str(0, 72, 31) = "`Geschlecht`"
 Str(0, 72, 32) = "`Straße`"
 Str(0, 72, 33) = "`PLZ`"
 Str(0, 72, 34) = "`Ort`"
 Str(0, 72, 35) = "`DMPT2`"
 Str(0, 72, 36) = "`DMPT1`"
 Str(0, 72, 37) = "`gelöscht`"
 Str(0, 72, 38) = "`ID`"
 Str(0, 72, 39) = "`PrimaryKey`"
 Str(0, 72, 40) = "`Auswahl`"
 Str(0, 72, 41) = "`KVNr`"
 ArtZ(0, 72) = 37
 ArtZ(1, 72) = 4
 Str(1, 72, 0) = "CREATE TABLE `hausaerzte` ("
 Str(1, 72, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 72, 2) = " `Überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '""L"" = Liebe(r), ""H"" = Hallo'"
 Str(1, 72, 3) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 4) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 5) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 6) = " `Anschrift` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 7) = " `KVNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 8) = " `Telefon` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 9) = " `Telefax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 10) = " `E_Mail` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 11) = " `Zulassungsgebiet` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 12) = " `Arzttyp` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 13) = " `Gemeinschaftspraxis mit` longtext COLLATE latin1_german2_ci"
 Str(1, 72, 14) = " `Schwerpunkt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 15) = " `Zusatzbezeichnung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 16) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 72, 17) = " `Beme` longtext COLLATE latin1_german2_ci"
 Str(1, 72, 18) = " `Sprechstunden` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 19) = " `von _ bis` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 20) = " `Internetadressen` longtext COLLATE latin1_german2_ci"
 Str(1, 72, 21) = " `Behandlung in Fremdsprachen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 22) = " `Rollstuhlgerechte Praxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 23) = " `Verkehrsmittel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 24) = " `Linie` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 25) = " `Haltestelle Parkplätze` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 26) = " `Wegbeschreibung` longtext COLLATE latin1_german2_ci"
 Str(1, 72, 27) = " `Entfernung zur Praxis` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 28) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 72, 29) = " `nichtmehr` tinyint(1) unsigned DEFAULT NULL COMMENT 'Arzt nicht mehr im Verzeichnis'"
 Str(1, 72, 30) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 31) = " `Geschlecht` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 32) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 33) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 34) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 72, 35) = " `DMPT2` smallint(6) DEFAULT NULL"
 Str(1, 72, 36) = " `DMPT1` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 72, 37) = " `gelöscht` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 72, 38) = "  PRIMARY KEY (`ID`)"
 Str(1, 72, 39) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 72, 40) = "  KEY `Auswahl` (`KVNr`,`Name`)"
 Str(1, 72, 41) = "  KEY `KVNr` (`KVNr`,`Nachname`,`Vorname`)"
 Str(1, 72, 42) = " ENGINE=InnoDB AUTO_INCREMENT=11080 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr72

Sub FüllStr73()
 Str(0, 73, 0) = "hausaerzte2"
 Str(0, 73, 1) = "`ID`"
 Str(0, 73, 2) = "`Überschrift`"
 Str(0, 73, 3) = "`Name`"
 Str(0, 73, 4) = "`Vorname`"
 Str(0, 73, 5) = "`Nachname`"
 Str(0, 73, 6) = "`Anschrift`"
 Str(0, 73, 7) = "`KVNr`"
 Str(0, 73, 8) = "`Telefon`"
 Str(0, 73, 9) = "`Telefax`"
 Str(0, 73, 10) = "`E_Mail`"
 Str(0, 73, 11) = "`Zulassungsgebiet`"
 Str(0, 73, 12) = "`Arzttyp`"
 Str(0, 73, 13) = "`Gemeinschaftspraxis mit`"
 Str(0, 73, 14) = "`Schwerpunkt`"
 Str(0, 73, 15) = "`Zusatzbezeichnung`"
 Str(0, 73, 16) = "`Bemerkung`"
 Str(0, 73, 17) = "`Beme`"
 Str(0, 73, 18) = "`Sprechstunden`"
 Str(0, 73, 19) = "`von _ bis`"
 Str(0, 73, 20) = "`Internetadressen`"
 Str(0, 73, 21) = "`Behandlung in Fremdsprachen`"
 Str(0, 73, 22) = "`Rollstuhlgerechte Praxis`"
 Str(0, 73, 23) = "`Verkehrsmittel`"
 Str(0, 73, 24) = "`Linie`"
 Str(0, 73, 25) = "`Haltestelle Parkplätze`"
 Str(0, 73, 26) = "`Wegbeschreibung`"
 Str(0, 73, 27) = "`Entfernung zur Praxis`"
 Str(0, 73, 28) = "`Zahl`"
 Str(0, 73, 29) = "`nichtmehr`"
 Str(0, 73, 30) = "`Titel`"
 Str(0, 73, 31) = "`Geschlecht`"
 Str(0, 73, 32) = "`Straße`"
 Str(0, 73, 33) = "`PLZ`"
 Str(0, 73, 34) = "`Ort`"
 Str(0, 73, 35) = "`DMPT2`"
 Str(0, 73, 36) = "`DMPT1`"
 Str(0, 73, 37) = "`gelöscht`"
 Str(0, 73, 38) = "`ID`"
 Str(0, 73, 39) = "`PrimaryKey`"
 Str(0, 73, 40) = "`Auswahl`"
 Str(0, 73, 41) = "`KVNr`"
 ArtZ(0, 73) = 37
 ArtZ(1, 73) = 4
 Str(1, 73, 0) = "CREATE TABLE `hausaerzte2` ("
 Str(1, 73, 1) = " `ID` int(10) NOT NULL"
 Str(1, 73, 2) = " `Überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '""L"" = Liebe(r), ""H"" = Hallo'"
 Str(1, 73, 3) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 4) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 5) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 6) = " `Anschrift` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 7) = " `KVNr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 8) = " `Telefon` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 9) = " `Telefax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 10) = " `E_Mail` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 11) = " `Zulassungsgebiet` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 12) = " `Arzttyp` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 13) = " `Gemeinschaftspraxis mit` longtext COLLATE latin1_german2_ci"
 Str(1, 73, 14) = " `Schwerpunkt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 15) = " `Zusatzbezeichnung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 16) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 73, 17) = " `Beme` longtext COLLATE latin1_german2_ci"
 Str(1, 73, 18) = " `Sprechstunden` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 19) = " `von _ bis` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 20) = " `Internetadressen` longtext COLLATE latin1_german2_ci"
 Str(1, 73, 21) = " `Behandlung in Fremdsprachen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 22) = " `Rollstuhlgerechte Praxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 23) = " `Verkehrsmittel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 24) = " `Linie` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 25) = " `Haltestelle Parkplätze` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 26) = " `Wegbeschreibung` longtext COLLATE latin1_german2_ci"
 Str(1, 73, 27) = " `Entfernung zur Praxis` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 28) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 73, 29) = " `nichtmehr` tinyint(1) unsigned DEFAULT NULL COMMENT 'Arzt nicht mehr im Verzeichnis'"
 Str(1, 73, 30) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 31) = " `Geschlecht` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 32) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 33) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 34) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 73, 35) = " `DMPT2` smallint(6) DEFAULT NULL"
 Str(1, 73, 36) = " `DMPT1` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 73, 37) = " `gelöscht` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 73, 38) = "  PRIMARY KEY (`ID`)"
 Str(1, 73, 39) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 73, 40) = "  KEY `Auswahl` (`KVNr`,`Name`)"
 Str(1, 73, 41) = "  KEY `KVNr` (`KVNr`,`Nachname`,`Vorname`)"
 Str(1, 73, 42) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr73

Sub FüllStr74()
 Str(0, 74, 0) = "haz"
 Str(0, 74, 1) = "`ID`"
 Str(0, 74, 2) = "`Name`"
 Str(0, 74, 3) = "`Anschrift`"
 Str(0, 74, 4) = "`KV_Nr`"
 Str(0, 74, 5) = "`Telefon`"
 Str(0, 74, 6) = "`Telefax`"
 Str(0, 74, 7) = "`E_Mail`"
 Str(0, 74, 8) = "`Zulassungsgebiet`"
 Str(0, 74, 9) = "`Arzttyp`"
 Str(0, 74, 10) = "`Gemeinschaftspraxis mit`"
 Str(0, 74, 11) = "`Schwerpunkt`"
 Str(0, 74, 12) = "`Zusatzbezeichnung`"
 Str(0, 74, 13) = "`Bemerkung`"
 Str(0, 74, 14) = "`Sprechstunden`"
 Str(0, 74, 15) = "`von _ bis`"
 Str(0, 74, 16) = "`Internetadressen`"
 Str(0, 74, 17) = "`Behandlung in Fremdsprachen`"
 Str(0, 74, 18) = "`Rollstuhlgerechte Praxis`"
 Str(0, 74, 19) = "`Verkehrsmittel`"
 Str(0, 74, 20) = "`Linie`"
 Str(0, 74, 21) = "`Haltestelle Parkplätze`"
 Str(0, 74, 22) = "`Wegbeschreibung`"
 Str(0, 74, 23) = "`Entfernung zur Praxis`"
 Str(0, 74, 24) = "`Zahl`"
 Str(0, 74, 25) = "`nichtmehr`"
 Str(0, 74, 26) = "`Vorname`"
 Str(0, 74, 27) = "`Nachname`"
 Str(0, 74, 28) = "`Titel`"
 Str(0, 74, 29) = "`Geschlecht`"
 Str(0, 74, 30) = "`Straße`"
 Str(0, 74, 31) = "`PLZ`"
 Str(0, 74, 32) = "`Ort`"
 Str(0, 74, 33) = "`ID`"
 Str(0, 74, 34) = "`PrimaryKey`"
 Str(0, 74, 35) = "`Auswahl`"
 Str(0, 74, 36) = "`KVNr`"
 ArtZ(0, 74) = 32
 ArtZ(1, 74) = 4
 Str(1, 74, 0) = "CREATE TABLE `haz` ("
 Str(1, 74, 1) = " `ID` int(10) NOT NULL"
 Str(1, 74, 2) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 3) = " `Anschrift` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 4) = " `KV_Nr` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 5) = " `Telefon` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 6) = " `Telefax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 7) = " `E_Mail` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 8) = " `Zulassungsgebiet` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 9) = " `Arzttyp` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 10) = " `Gemeinschaftspraxis mit` longtext COLLATE latin1_german2_ci"
 Str(1, 74, 11) = " `Schwerpunkt` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 12) = " `Zusatzbezeichnung` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 13) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 74, 14) = " `Sprechstunden` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 15) = " `von _ bis` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 16) = " `Internetadressen` longtext COLLATE latin1_german2_ci"
 Str(1, 74, 17) = " `Behandlung in Fremdsprachen` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 18) = " `Rollstuhlgerechte Praxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 19) = " `Verkehrsmittel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 20) = " `Linie` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 21) = " `Haltestelle Parkplätze` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 22) = " `Wegbeschreibung` longtext COLLATE latin1_german2_ci"
 Str(1, 74, 23) = " `Entfernung zur Praxis` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 24) = " `Zahl` int(10) DEFAULT NULL"
 Str(1, 74, 25) = " `nichtmehr` tinyint(1) unsigned DEFAULT NULL COMMENT 'Arzt nicht mehr im Verzeichnis'"
 Str(1, 74, 26) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 27) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 28) = " `Titel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 29) = " `Geschlecht` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 30) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 31) = " `PLZ` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 32) = " `Ort` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 74, 33) = "  PRIMARY KEY (`ID`)"
 Str(1, 74, 34) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 74, 35) = "  KEY `Auswahl` (`KV_Nr`,`Name`)"
 Str(1, 74, 36) = "  KEY `KVNr` (`KV_Nr`)"
 Str(1, 74, 37) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr74

Sub FüllStr75()
 Str(0, 75, 0) = "icd2006"
 Str(0, 75, 1) = "`KlassEbene`"
 Str(0, 75, 2) = "`OrtdSchlüNr`"
 Str(0, 75, 3) = "`ArtdVierst`"
 Str(0, 75, 4) = "`KapitelNr`"
 Str(0, 75, 5) = "`Dreist`"
 Str(0, 75, 6) = "`SchlüssNrRoh`"
 Str(0, 75, 7) = "`KlassTit`"
 Str(0, 75, 8) = "`VerwdP295`"
 Str(0, 75, 9) = "`VerwdP311`"
 Str(0, 75, 10) = "`ME1`"
 Str(0, 75, 11) = "`ME2`"
 Str(0, 75, 12) = "`ME3`"
 Str(0, 75, 13) = "`ME4`"
 Str(0, 75, 14) = "`MbListe`"
 Str(0, 75, 15) = "`Geschl`"
 Str(0, 75, 16) = "`FlrBGschl`"
 Str(0, 75, 17) = "`untAltG`"
 Str(0, 75, 18) = "`untAltGalt`"
 Str(0, 75, 19) = "`obAltG`"
 Str(0, 75, 20) = "`obAltGalt`"
 Str(0, 75, 21) = "`ArtdFbAlt`"
 Str(0, 75, 22) = "`KiMs`"
 Str(0, 75, 23) = "`SNrmInh`"
 Str(0, 75, 24) = "`IfsGMeld`"
 Str(0, 75, 25) = "`IfsGLab`"
 ArtZ(0, 75) = 25
 Str(1, 75, 0) = "CREATE TABLE `icd2006` ("
 Str(1, 75, 1) = " `KlassEbene` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 2) = " `OrtdSchlüNr` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 3) = " `ArtdVierst` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 4) = " `KapitelNr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 5) = " `Dreist` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 6) = " `SchlüssNrRoh` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 7) = " `KlassTit` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 8) = " `VerwdP295` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 9) = " `VerwdP311` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 10) = " `ME1` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 11) = " `ME2` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 12) = " `ME3` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 13) = " `ME4` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 14) = " `MbListe` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 15) = " `Geschl` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 16) = " `FlrBGschl` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 17) = " `untAltG` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 18) = " `untAltGalt` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 19) = " `obAltG` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 20) = " `obAltGalt` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 21) = " `ArtdFbAlt` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 22) = " `KiMs` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 23) = " `SNrmInh` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 24) = " `IfsGMeld` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 25) = " `IfsGLab` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 75, 26) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr75

Sub FüllStr76()
 Str(0, 76, 0) = "inl"
 Str(0, 76, 1) = "`id`"
 Str(0, 76, 2) = "`breite`"
 Str(0, 76, 3) = "`kennung`"
 Str(0, 76, 4) = "`inhalt`"
 Str(0, 76, 5) = "`id`"
 Str(0, 76, 6) = "`kennunginhalt`"
 ArtZ(0, 76) = 4
 ArtZ(1, 76) = 2
 Str(1, 76, 0) = "CREATE TABLE `inl` ("
 Str(1, 76, 1) = " `id` int(2) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 76, 2) = " `breite` varchar(3) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 76, 3) = " `kennung` varchar(4) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 76, 4) = " `inhalt` varchar(10000) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 76, 5) = "  PRIMARY KEY (`id`)"
 Str(1, 76, 6) = "  KEY `kennunginhalt` (`kennung`,`inhalt`(6))"
 Str(1, 76, 7) = " ENGINE=MyISAM AUTO_INCREMENT=10224275 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr76

Sub FüllStr77()
 Str(0, 77, 0) = "inlalt"
 Str(0, 77, 1) = "`id`"
 Str(0, 77, 2) = "`breite`"
 Str(0, 77, 3) = "`kennung`"
 Str(0, 77, 4) = "`inhalt`"
 Str(0, 77, 5) = "`id`"
 Str(0, 77, 6) = "`kennunginhalt`"
 ArtZ(0, 77) = 4
 ArtZ(1, 77) = 2
 Str(1, 77, 0) = "CREATE TABLE `inlalt` ("
 Str(1, 77, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 77, 2) = " `breite` char(3) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 77, 3) = " `kennung` char(4) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 77, 4) = " `inhalt` varchar(10000) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 77, 5) = "  PRIMARY KEY (`id`)"
 Str(1, 77, 6) = "  KEY `kennunginhalt` (`kennung`,`inhalt`(6))"
 Str(1, 77, 7) = " ENGINE=MyISAM AUTO_INCREMENT=10206987 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr77

Sub FüllStr78()
 Str(0, 78, 0) = "kassenliste"
 Str(0, 78, 1) = "`ID`"
 Str(0, 78, 2) = "`VK`"
 Str(0, 78, 3) = "`IK`"
 Str(0, 78, 4) = "`Name`"
 Str(0, 78, 5) = "`Kateg`"
 Str(0, 78, 6) = "`AnzahlIK`"
 Str(0, 78, 7) = "`AnzahlKTUG`"
 Str(0, 78, 8) = "`GültigVon`"
 Str(0, 78, 9) = "`GültigBis`"
 Str(0, 78, 10) = "`GO`"
 Str(0, 78, 11) = "`Kurzname`"
 Str(0, 78, 12) = "`rName`"
 Str(0, 78, 13) = "`ID`"
 Str(0, 78, 14) = "`PrimaryKey`"
 Str(0, 78, 15) = "`VK`"
 Str(0, 78, 16) = "`IK`"
 Str(0, 78, 17) = "`VKIK`"
 ArtZ(0, 78) = 12
 ArtZ(1, 78) = 5
 Str(1, 78, 0) = "CREATE TABLE `kassenliste` ("
 Str(1, 78, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 78, 2) = " `VK` varchar(6) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 78, 3) = " `IK` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 78, 4) = " `Name` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 78, 5) = " `Kateg` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kategorie'"
 Str(1, 78, 6) = " `AnzahlIK` int(4) unsigned DEFAULT NULL"
 Str(1, 78, 7) = " `AnzahlKTUG` int(4) unsigned DEFAULT NULL"
 Str(1, 78, 8) = " `GültigVon` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 78, 9) = " `GültigBis` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 78, 10) = " `GO` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 78, 11) = " `Kurzname` varchar(60) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 78, 12) = " `rName` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kurzname, falls nicht verfügbar: Name'"
 Str(1, 78, 13) = "  PRIMARY KEY (`ID`)"
 Str(1, 78, 14) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 78, 15) = "  KEY `VK` (`VK`)"
 Str(1, 78, 16) = "  KEY `IK` (`IK`)"
 Str(1, 78, 17) = "  KEY `VKIK` (`VK`,`IK`)"
 Str(1, 78, 18) = " ENGINE=InnoDB AUTO_INCREMENT=10578 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr78

Sub FüllStr79()
 Str(0, 79, 0) = "kheinweis"
 Str(0, 79, 1) = "`FID`"
 Str(0, 79, 2) = "`Pat_ID`"
 Str(0, 79, 3) = "`ZeitPunkt`"
 Str(0, 79, 4) = "`Ziel`"
 Str(0, 79, 5) = "`Diagnose`"
 Str(0, 79, 6) = "`absPos`"
 Str(0, 79, 7) = "`AktZeit`"
 Str(0, 79, 8) = "`StByte`"
 Str(0, 79, 9) = "`Auswahl`"
 Str(0, 79, 10) = "`FälleKHEinweis`"
 Str(0, 79, 11) = "`FID`"
 Str(0, 79, 12) = "`NamenKHEinweis`"
 Str(0, 79, 13) = "`FälleKHEinweis_AccRel`"
 Str(0, 79, 14) = "`NamenKHEinweis_AccRel`"
 ArtZ(0, 79) = 8
 ArtZ(1, 79) = 4
 ArtZ(2, 79) = 2
 Str(1, 79, 0) = "CREATE TABLE `kheinweis` ("
 Str(1, 79, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 79, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 79, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 79, 4) = " `Ziel` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6291'"
 Str(1, 79, 5) = " `Diagnose` longtext COLLATE latin1_german2_ci COMMENT '6230'"
 Str(1, 79, 6) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 79, 7) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 79, 8) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 79, 9) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Ziel`)"
 Str(1, 79, 10) = "  KEY `FälleKHEinweis` (`FID`)"
 Str(1, 79, 11) = "  KEY `FID` (`FID`)"
 Str(1, 79, 12) = "  KEY `NamenKHEinweis` (`Pat_ID`)"
 Str(1, 79, 13) = "  CONSTRAINT `FälleKHEinweis_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON UPDATE CASCADE"
 Str(1, 79, 14) = "  CONSTRAINT `NamenKHEinweis_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON UPDATE CASCADE"
 Str(1, 79, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr79

Sub FüllStr80()
 Str(0, 80, 0) = "kontaktzahl"
 Str(0, 80, 1) = "`ct`"
 Str(0, 80, 2) = "`pat_id`"
 ArtZ(0, 80) = 2
 Str(1, 80, 0) = "CREATE TABLE `kontaktzahl` ("
 Str(1, 80, 1) = " `ct` bigint(21) DEFAULT NULL"
 Str(1, 80, 2) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 80, 3) = " ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr80

Sub FüllStr81()
 Str(0, 81, 0) = "kvnrue"
 Str(0, 81, 1) = "`lfdnr`"
 Str(0, 81, 2) = "`Pat_ID`"
 Str(0, 81, 3) = "`KVNr`"
 Str(0, 81, 4) = "`absPos`"
 Str(0, 81, 5) = "`AktZeit`"
 Str(0, 81, 6) = "`StByte`"
 Str(0, 81, 7) = "`lfdnr`"
 Str(0, 81, 8) = "`PrimaryKey`"
 Str(0, 81, 9) = "`zuord`"
 ArtZ(0, 81) = 6
 ArtZ(1, 81) = 3
 Str(1, 81, 0) = "CREATE TABLE `kvnrue` ("
 Str(1, 81, 1) = " `lfdnr` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 81, 2) = " `Pat_ID` int(10) DEFAULT NULL"
 Str(1, 81, 3) = " `KVNr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 81, 4) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 81, 5) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Zeit der Aktualisuerung aus der BDT-Datei'"
 Str(1, 81, 6) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 81, 7) = "  PRIMARY KEY (`lfdnr`)"
 Str(1, 81, 8) = "  UNIQUE KEY `PrimaryKey` (`lfdnr`)"
 Str(1, 81, 9) = "  UNIQUE KEY `zuord` (`Pat_ID`,`KVNr`)"
 Str(1, 81, 10) = " ENGINE=InnoDB AUTO_INCREMENT=3168 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr81

Sub FüllStr82()
 Str(0, 82, 0) = "labor1"
 Str(1, 82, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `labor1` AS select `n`.`Pat_ID` AS `Pat_ID`,cast(`n`.`ZeitPunkt` as date) AS `ZeitPunkt`,`n`.`FertigStGrad` AS `FertigStGrad`,`n`.`Abkü` AS `Abkü`,`l`.`Langtext` AS `Langtext`,`n`.`Wert` AS `Wert`,`n`.`Einheit` AS `Einheit`,`k`.`Kommentar` AS `Kommentar`,_utf8'' AS `nb` from ((`laborneu` `n` left join `laborlangtext` `l` on((`l`.`LangtextVW` = `n`.`LangtextVW`))) left join `laborkommentar` `k` on((`k`.`KommentarVW` = `n`.`KommentarVW`))) where ((`n`.`Wert` <> '') and (`n`.`Wert` is not null))"
End Sub ' FüllStr82

Sub FüllStr83()
 Str(0, 83, 0) = "labor2"
 Str(1, 83, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `labor2` AS select `u`.`Pat_id` AS `Pat_ID`,cast(`u`.`Eingang` as date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,`w`.`Abkü` AS `Abkü`,`w`.`Langname` AS `Langtext`,`w`.`Wert` AS `Wert`,`w`.`Einheit` AS `Einheit`,`w`.`Kommentar` AS `Kommentar`,`w`.`Normbereich` AS `NB` from (`laborxus` `u` left join `laborxwert` `w` on((`u`.`RefNr` = `w`.`RefNr`))) where ((not(exists(select 1 AS `Not_used` from `laborneu` where ((`laborneu`.`Pat_ID` = `u`.`Pat_id`) and (`laborneu`.`Abkü` = `w`.`Abkü`) and (`laborneu`.`Wert` = `w`.`Wert`) and (`laborneu`.`ZeitPunkt` > (`u`.`Eingang` - interval 3 day)) and (`laborneu`.`ZeitPunkt` < (`u`.`Eingang` + interval 6 day)))))) and (`w`.`Wert` <> '') and (`w`.`Wert` is not null))"
End Sub ' FüllStr83

Sub FüllStr84()
 Str(0, 84, 0) = "laborgruppen"
 Str(0, 84, 1) = "`Laborgruppe`"
 Str(0, 84, 2) = "`Erklärung`"
 Str(0, 84, 3) = "`Laborgruppe`"
 Str(0, 84, 4) = "`Gruppe`"
 ArtZ(0, 84) = 2
 ArtZ(1, 84) = 2
 Str(1, 84, 0) = "CREATE TABLE `laborgruppen` ("
 Str(1, 84, 1) = " `Laborgruppe` int(10) NOT NULL DEFAULT '0'"
 Str(1, 84, 2) = " `Erklärung` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 84, 3) = "  PRIMARY KEY (`Laborgruppe`)"
 Str(1, 84, 4) = "  UNIQUE KEY `Gruppe` (`Laborgruppe`)"
 Str(1, 84, 5) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr84

Sub FüllStr85()
 Str(0, 85, 0) = "laborkommentar"
 Str(0, 85, 1) = "`KommentarVW`"
 Str(0, 85, 2) = "`Kommentar`"
 Str(0, 85, 3) = "`KommentarVW`"
 Str(0, 85, 4) = "`KommentarVW`"
 Str(0, 85, 5) = "`Kommentar`"
 ArtZ(0, 85) = 2
 ArtZ(1, 85) = 3
 Str(1, 85, 0) = "CREATE TABLE `laborkommentar` ("
 Str(1, 85, 1) = " `KommentarVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 85, 2) = " `Kommentar` longtext COLLATE latin1_german2_ci"
 Str(1, 85, 3) = "  PRIMARY KEY (`KommentarVW`)"
 Str(1, 85, 4) = "  UNIQUE KEY `KommentarVW` (`KommentarVW`)"
 Str(1, 85, 5) = "  KEY `Kommentar` (`Kommentar`(255))"
 Str(1, 85, 6) = " ENGINE=InnoDB AUTO_INCREMENT=28561 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr85

Sub FüllStr86()
 Str(0, 86, 0) = "laborlangtext"
 Str(0, 86, 1) = "`LangtextVW`"
 Str(0, 86, 2) = "`Langtext`"
 Str(0, 86, 3) = "`LangtextVW`"
 Str(0, 86, 4) = "`LangtextVW`"
 Str(0, 86, 5) = "`Langtext`"
 ArtZ(0, 86) = 2
 ArtZ(1, 86) = 3
 Str(1, 86, 0) = "CREATE TABLE `laborlangtext` ("
 Str(1, 86, 1) = " `LangtextVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 86, 2) = " `Langtext` longtext COLLATE latin1_german2_ci"
 Str(1, 86, 3) = "  PRIMARY KEY (`LangtextVW`)"
 Str(1, 86, 4) = "  UNIQUE KEY `LangtextVW` (`LangtextVW`)"
 Str(1, 86, 5) = "  KEY `Langtext` (`Langtext`(255))"
 Str(1, 86, 6) = " ENGINE=InnoDB AUTO_INCREMENT=943 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr86

Sub FüllStr87()
 Str(0, 87, 0) = "laborneu"
 Str(0, 87, 1) = "`FID`"
 Str(0, 87, 2) = "`Pat_ID`"
 Str(0, 87, 3) = "`ZeitPunkt`"
 Str(0, 87, 4) = "`FertigStGrad`"
 Str(0, 87, 5) = "`Abkü`"
 Str(0, 87, 6) = "`LangtextVW`"
 Str(0, 87, 7) = "`Wert`"
 Str(0, 87, 8) = "`Einheit`"
 Str(0, 87, 9) = "`KommentarVW`"
 Str(0, 87, 10) = "`AbsPos`"
 Str(0, 87, 11) = "`AktZeit`"
 Str(0, 87, 12) = "`Refnr`"
 Str(0, 87, 13) = "`StByte`"
 Str(0, 87, 14) = "`AbküWert`"
 Str(0, 87, 15) = "`Auswahl`"
 Str(0, 87, 16) = "`FälleLaborNeu`"
 Str(0, 87, 17) = "`LaborKommentarLaborNeu`"
 Str(0, 87, 18) = "`LaborLangtextLaborNeu`"
 Str(0, 87, 19) = "`LaborParameterLaborNeu`"
 Str(0, 87, 20) = "`NamenLaborNeu`"
 Str(0, 87, 21) = "`Prüf`"
 Str(0, 87, 22) = "`FälleLaborNeu_AccRel`"
 Str(0, 87, 23) = "`LaborKommentarLaborNeu_AccRel`"
 Str(0, 87, 24) = "`LaborLangtextLaborNeu_AccRel`"
 Str(0, 87, 25) = "`LaborParameterLaborNeu_AccRel`"
 Str(0, 87, 26) = "`NamenLaborNeu_AccRel`"
 ArtZ(0, 87) = 13
 ArtZ(1, 87) = 8
 ArtZ(2, 87) = 5
 Str(1, 87, 0) = "CREATE TABLE `laborneu` ("
 Str(1, 87, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 87, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 87, 3) = " `ZeitPunkt` datetime DEFAULT NULL"
 Str(1, 87, 4) = " `FertigStGrad` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8401'"
 Str(1, 87, 5) = " `Abkü` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410'"
 Str(1, 87, 6) = " `LangtextVW` int(10) DEFAULT NULL COMMENT '8411'"
 Str(1, 87, 7) = " `Wert` varchar(69) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8420'"
 Str(1, 87, 8) = " `Einheit` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421'"
 Str(1, 87, 9) = " `KommentarVW` int(10) DEFAULT NULL COMMENT '8480'"
 Str(1, 87, 10) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 87, 11) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 87, 12) = " `Refnr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborXUS'"
 Str(1, 87, 13) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 87, 14) = "  KEY `AbküWert` (`Abkü`,`Wert`)"
 Str(1, 87, 15) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`FertigStGrad`,`Abkü`)"
 Str(1, 87, 16) = "  KEY `FälleLaborNeu` (`FID`)"
 Str(1, 87, 17) = "  KEY `LaborKommentarLaborNeu` (`KommentarVW`)"
 Str(1, 87, 18) = "  KEY `LaborLangtextLaborNeu` (`LangtextVW`)"
 Str(1, 87, 19) = "  KEY `LaborParameterLaborNeu` (`Abkü`,`Einheit`)"
 Str(1, 87, 20) = "  KEY `NamenLaborNeu` (`Pat_ID`)"
 Str(1, 87, 21) = "  KEY `Prüf` (`Pat_ID`,`Abkü`,`Wert`)"
 Str(1, 87, 22) = "  CONSTRAINT `FälleLaborNeu_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 87, 23) = "  CONSTRAINT `LaborKommentarLaborNeu_AccRel` FOREIGN KEY (`KommentarVW`) REFERENCES `laborkommentar` (`KommentarVW`)"
 Str(1, 87, 24) = "  CONSTRAINT `LaborLangtextLaborNeu_AccRel` FOREIGN KEY (`LangtextVW`) REFERENCES `laborlangtext` (`LangtextVW`) ON UPDATE CASCADE"
 Str(1, 87, 25) = "  CONSTRAINT `LaborParameterLaborNeu_AccRel` FOREIGN KEY (`Abkü`, `Einheit`) REFERENCES `laborparameter` (`Abkü`, `Einheit`)"
 Str(1, 87, 26) = "  CONSTRAINT `NamenLaborNeu_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 87, 27) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr87

Sub FüllStr88()
 Str(0, 88, 0) = "laborparameter"
 Str(0, 88, 1) = "`Abkü`"
 Str(0, 88, 2) = "`Labor`"
 Str(0, 88, 3) = "`Langtext`"
 Str(0, 88, 4) = "`Einheit`"
 Str(0, 88, 5) = "`Gruppe`"
 Str(0, 88, 6) = "`Reihe`"
 Str(0, 88, 7) = "`uNm`"
 Str(0, 88, 8) = "`oNm`"
 Str(0, 88, 9) = "`uNw`"
 Str(0, 88, 10) = "`oNw`"
 Str(0, 88, 11) = "`AktZeit`"
 Str(0, 88, 12) = "`StByte`"
 Str(0, 88, 13) = "`Abkü`"
 Str(0, 88, 14) = "`Fehlende`"
 Str(0, 88, 15) = "`LaborParameterAbkü`"
 Str(0, 88, 16) = "`Reihe`"
 Str(0, 88, 17) = "`LaborgruppenLaborParameter_AccRel`"
 ArtZ(0, 88) = 12
 ArtZ(1, 88) = 4
 ArtZ(2, 88) = 1
 Str(1, 88, 0) = "CREATE TABLE `laborparameter` ("
 Str(1, 88, 1) = " `Abkü` varchar(70) COLLATE latin1_german2_ci NOT NULL COMMENT '8410(1)'"
 Str(1, 88, 2) = " `Labor` varchar(40) COLLATE latin1_german2_ci NOT NULL COMMENT '8410(2)'"
 Str(1, 88, 3) = " `Langtext` varchar(60) COLLATE latin1_german2_ci NOT NULL COMMENT '8411'"
 Str(1, 88, 4) = " `Einheit` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421'"
 Str(1, 88, 5) = " `Gruppe` int(10) DEFAULT NULL COMMENT 'Ordnungsgruppe'"
 Str(1, 88, 6) = " `Reihe` int(10) DEFAULT NULL COMMENT 'Reihenfolge innerhalb der Gruppe'"
 Str(1, 88, 7) = " `uNm` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'unterer Normwert männlich'"
 Str(1, 88, 8) = " `oNm` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'oberer Normwert männlich'"
 Str(1, 88, 9) = " `uNw` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'unterer Normwert weiblich'"
 Str(1, 88, 10) = " `oNw` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'oberer Normwert weiblich'"
 Str(1, 88, 11) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 88, 12) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 88, 13) = "  UNIQUE KEY `Abkü` (`Abkü`,`Einheit`)"
 Str(1, 88, 14) = "  KEY `Fehlende` (`Gruppe`,`AktZeit`)"
 Str(1, 88, 15) = "  KEY `LaborParameterAbkü` (`Abkü`)"
 Str(1, 88, 16) = "  KEY `Reihe` (`Gruppe`,`Abkü`)"
 Str(1, 88, 17) = "  CONSTRAINT `LaborgruppenLaborParameter_AccRel` FOREIGN KEY (`Gruppe`) REFERENCES `laborgruppen` (`Laborgruppe`)"
 Str(1, 88, 18) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci ROW_FORMAT=DYNAMIC"
End Sub ' FüllStr88

Sub FüllStr89()
 Str(0, 89, 0) = "laborxbakt"
 Str(0, 89, 1) = "`RefNr`"
 Str(0, 89, 2) = "`Verf`"
 Str(0, 89, 3) = "`KuQu`"
 Str(0, 89, 4) = "`Quelle`"
 Str(0, 89, 5) = "`QSpez`"
 Str(0, 89, 6) = "`AbnDat`"
 Str(0, 89, 7) = "`Kommentar`"
 Str(0, 89, 8) = "`Erklärung`"
 Str(0, 89, 9) = "`Keimzahl`"
 Str(0, 89, 10) = "`LaborXBaktUS`"
 Str(0, 89, 11) = "`LaborXUSLaborXBakt`"
 Str(0, 89, 12) = "`RefNr`"
 Str(0, 89, 13) = "`LaborXBaktUS_AccRel`"
 Str(0, 89, 14) = "`LaborXUSLaborXBakt_AccRel`"
 ArtZ(0, 89) = 9
 ArtZ(1, 89) = 3
 ArtZ(2, 89) = 2
 Str(1, 89, 0) = "CREATE TABLE `laborxbakt` ("
 Str(1, 89, 1) = " `RefNr` int(10) DEFAULT NULL"
 Str(1, 89, 2) = " `Verf` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 3) = " `KuQu` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8428 Probenmaterial-Ident (Turbomed)'"
 Str(1, 89, 4) = " `Quelle` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8430 Probenmaterial-Bezeichnung (Turbomed)'"
 Str(1, 89, 5) = " `QSpez` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8431 Probenmaterial-Spezifikation (Turbomed)'"
 Str(1, 89, 6) = " `AbnDat` datetime DEFAULT NULL COMMENT '8432 Abnahmedatum (Turbomed)'"
 Str(1, 89, 7) = " `Kommentar` longtext COLLATE latin1_german2_ci COMMENT '8480 Ergebnistest (Turbomed)'"
 Str(1, 89, 8) = " `Erklärung` longtext COLLATE latin1_german2_ci"
 Str(1, 89, 9) = " `Keimzahl` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 89, 10) = "  KEY `LaborXBaktUS` (`RefNr`)"
 Str(1, 89, 11) = "  KEY `LaborXUSLaborXBakt` (`RefNr`)"
 Str(1, 89, 12) = "  KEY `RefNr` (`RefNr`)"
 Str(1, 89, 13) = "  CONSTRAINT `LaborXBaktUS_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 89, 14) = "  CONSTRAINT `LaborXUSLaborXBakt_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 89, 15) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr89

Sub FüllStr90()
 Str(0, 90, 0) = "laborxeingel"
 Str(0, 90, 1) = "`DatID`"
 Str(0, 90, 2) = "`Pfad`"
 Str(0, 90, 3) = "`Name`"
 Str(0, 90, 4) = "`Zp`"
 Str(0, 90, 5) = "`fertig`"
 Str(0, 90, 6) = "`DatID`"
 Str(0, 90, 7) = "`DatID`"
 Str(0, 90, 8) = "`NamePfad`"
 ArtZ(0, 90) = 5
 ArtZ(1, 90) = 3
 Str(1, 90, 0) = "CREATE TABLE `laborxeingel` ("
 Str(1, 90, 1) = " `DatID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf LaborEingelesen'"
 Str(1, 90, 2) = " `Pfad` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Pfadname'"
 Str(1, 90, 3) = " `Name` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Name der eingelesenen Labordatei ohne Endung'"
 Str(1, 90, 4) = " `Zp` datetime DEFAULT NULL COMMENT 'Einlesezeitpunkt'"
 Str(1, 90, 5) = " `fertig` bit(1) DEFAULT NULL COMMENT 'ob Einlesen fertig'"
 Str(1, 90, 6) = "  PRIMARY KEY (`DatID`)"
 Str(1, 90, 7) = "  UNIQUE KEY `DatID` (`DatID`)"
 Str(1, 90, 8) = "  KEY `NamePfad` (`Name`,`Pfad`)"
 Str(1, 90, 9) = " ENGINE=InnoDB AUTO_INCREMENT=269 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr90

Sub FüllStr91()
 Str(0, 91, 0) = "laborxleist"
 Str(0, 91, 1) = "`RefNr`"
 Str(0, 91, 2) = "`Abkü`"
 Str(0, 91, 3) = "`Verf`"
 Str(0, 91, 4) = "`EBM`"
 Str(0, 91, 5) = "`goä`"
 Str(0, 91, 6) = "`Anzahl`"
 Str(0, 91, 7) = "`LaborXLeistUS`"
 Str(0, 91, 8) = "`RefNr`"
 Str(0, 91, 9) = "`LaborXLeistUS_AccRel`"
 ArtZ(0, 91) = 6
 ArtZ(1, 91) = 2
 ArtZ(2, 91) = 1
 Str(1, 91, 0) = "CREATE TABLE `laborxleist` ("
 Str(1, 91, 1) = " `RefNr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborUS'"
 Str(1, 91, 2) = " `Abkü` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410 Test-Ident (Turbomed)'"
 Str(1, 91, 3) = " `Verf` varchar(42) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8434'"
 Str(1, 91, 4) = " `EBM` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5001 GNR (Turbomed)'"
 Str(1, 91, 5) = " `goä` varchar(8) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8406'"
 Str(1, 91, 6) = " `Anzahl` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5005'"
 Str(1, 91, 7) = "  KEY `LaborXLeistUS` (`RefNr`)"
 Str(1, 91, 8) = "  KEY `RefNr` (`RefNr`)"
 Str(1, 91, 9) = "  CONSTRAINT `LaborXLeistUS_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 91, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr91

Sub FüllStr92()
 Str(0, 92, 0) = "laborxls"
 Str(0, 92, 1) = "`id`"
 Str(0, 92, 2) = "`patient`"
 Str(0, 92, 3) = "`fehlerart`"
 Str(0, 92, 4) = "`id`"
 ArtZ(0, 92) = 3
 ArtZ(1, 92) = 1
 Str(1, 92, 0) = "CREATE TABLE `laborxls` ("
 Str(1, 92, 1) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 92, 2) = " `patient` varchar(255) DEFAULT NULL"
 Str(1, 92, 3) = " `fehlerart` varchar(3000) DEFAULT NULL"
 Str(1, 92, 4) = "  PRIMARY KEY (`id`)"
 Str(1, 92, 5) = " ENGINE=MyISAM AUTO_INCREMENT=51 DEFAULT CHARSET=latin1"
End Sub ' FüllStr92

Sub FüllStr93()
 Str(0, 93, 0) = "laborxsaetze"
 Str(0, 93, 1) = "`SatzID`"
 Str(0, 93, 2) = "`DatID`"
 Str(0, 93, 3) = "`Satzart`"
 Str(0, 93, 4) = "`Satzlänge`"
 Str(0, 93, 5) = "`SatzlängeSchluss`"
 Str(0, 93, 6) = "`VersionSatzb`"
 Str(0, 93, 7) = "`Arztnr`"
 Str(0, 93, 8) = "`Arztname`"
 Str(0, 93, 9) = "`StraßePraxis`"
 Str(0, 93, 10) = "`Arzt`"
 Str(0, 93, 11) = "`LANR`"
 Str(0, 93, 12) = "`PLZPraxis`"
 Str(0, 93, 13) = "`OrtPraxis`"
 Str(0, 93, 14) = "`Labor`"
 Str(0, 93, 15) = "`StraßeLabor`"
 Str(0, 93, 16) = "`PLZLabor`"
 Str(0, 93, 17) = "`OrtLabor`"
 Str(0, 93, 18) = "`KBVPrüfnr`"
 Str(0, 93, 19) = "`Zeichensatz`"
 Str(0, 93, 20) = "`Kundenarztnr`"
 Str(0, 93, 21) = "`Erstellungsdatum`"
 Str(0, 93, 22) = "`Gesamtlänge`"
 Str(0, 93, 23) = "`SatzID`"
 Str(0, 93, 24) = "`SatzID`"
 Str(0, 93, 25) = "`DatID`"
 Str(0, 93, 26) = "`Name`"
 ArtZ(0, 93) = 22
 ArtZ(1, 93) = 4
 Str(1, 93, 0) = "CREATE TABLE `laborxsaetze` ("
 Str(1, 93, 1) = " `SatzID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'zum Bezug für LaborUS'"
 Str(1, 93, 2) = " `DatID` int(10) DEFAULT NULL COMMENT 'Bezug zu LaborEingelesen'"
 Str(1, 93, 3) = " `Satzart` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000 Satzart (Turbomed)'"
 Str(1, 93, 4) = " `Satzlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed)'"
 Str(1, 93, 5) = " `SatzlängeSchluss` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000'"
 Str(1, 93, 6) = " `VersionSatzb` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9212 Version der Satzbeschreibung (Turbomed)'"
 Str(1, 93, 7) = " `Arztnr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '201 Arztnummer (Turbomed)'"
 Str(1, 93, 8) = " `Arztname` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '203 Arztname (Turbomed)'"
 Str(1, 93, 9) = " `StraßePraxis` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '205 Straße der Praxis (Turbomed)'"
 Str(1, 93, 10) = " `Arzt` varchar(45) COLLATE latin1_german2_ci NOT NULL COMMENT ' 211 Ausführender Arzt'"
 Str(1, 93, 11) = " `LANR` varchar(11) COLLATE latin1_german2_ci NOT NULL COMMENT ' 212 LANR'"
 Str(1, 93, 12) = " `PLZPraxis` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '215 PLZ der Praxis (Turbomed)'"
 Str(1, 93, 13) = " `OrtPraxis` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '216 Ort der Praxis (Turbomed)'"
 Str(1, 93, 14) = " `Labor` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8320 Labor'"
 Str(1, 93, 15) = " `StraßeLabor` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8321 Straße der Laboradresse (Turbomed)'"
 Str(1, 93, 16) = " `PLZLabor` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8322 PLZ der Laboradresse (Turbomed)'"
 Str(1, 93, 17) = " `OrtLabor` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8323 Ort der Laboradresse (Turbomed)'"
 Str(1, 93, 18) = " `KBVPrüfnr` varchar(16) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '101 KBV-Prüfnummer (Turbomed)'"
 Str(1, 93, 19) = " `Zeichensatz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9106 verwendeter Zeichensatz (Turbomed)'"
 Str(1, 93, 20) = " `Kundenarztnr` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8312 Kundenarztnummer (Turbomed)'"
 Str(1, 93, 21) = " `Erstellungsdatum` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9103 Erstellungsdatum (Turbomed)'"
 Str(1, 93, 22) = " `Gesamtlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '9202 Gesamtlänge des Datenpaketes (Turbomed)'"
 Str(1, 93, 23) = "  PRIMARY KEY (`SatzID`)"
 Str(1, 93, 24) = "  UNIQUE KEY `SatzID` (`SatzID`)"
 Str(1, 93, 25) = "  KEY `DatID` (`DatID`)"
 Str(1, 93, 26) = "  KEY `Name` (`PLZLabor`,`OrtLabor`)"
 Str(1, 93, 27) = " ENGINE=InnoDB AUTO_INCREMENT=685 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr93

Sub FüllStr94()
 Str(0, 94, 0) = "laborxus"
 Str(0, 94, 1) = "`RefNr`"
 Str(0, 94, 2) = "`DatID`"
 Str(0, 94, 3) = "`SatzID`"
 Str(0, 94, 4) = "`Satzart`"
 Str(0, 94, 5) = "`Satzlänge`"
 Str(0, 94, 6) = "`Auftragsnummer`"
 Str(0, 94, 7) = "`Auftragsschlüssel`"
 Str(0, 94, 8) = "`Eingang`"
 Str(0, 94, 9) = "`Berichtsdatum`"
 Str(0, 94, 10) = "`Pat_id`"
 Str(0, 94, 11) = "`Nachname`"
 Str(0, 94, 12) = "`Vorname`"
 Str(0, 94, 13) = "`GebDat`"
 Str(0, 94, 14) = "`Titel`"
 Str(0, 94, 15) = "`NVorsatz`"
 Str(0, 94, 16) = "`BefArt`"
 Str(0, 94, 17) = "`Abrechnungstyp`"
 Str(0, 94, 18) = "`GebüOrd`"
 Str(0, 94, 19) = "`Patienteninformation`"
 Str(0, 94, 20) = "`Geschlecht`"
 Str(0, 94, 21) = "`AuftrHinw`"
 Str(0, 94, 22) = "`Pat_idUrsp`"
 Str(0, 94, 23) = "`Pat_idErwVNG`"
 Str(0, 94, 24) = "`Pat_idErwVN`"
 Str(0, 94, 25) = "`Pat_idErwG`"
 Str(0, 94, 26) = "`Pat_idErwGB`"
 Str(0, 94, 27) = "`Pat_idErwGL`"
 Str(0, 94, 28) = "`Pat_idLaborNeu`"
 Str(0, 94, 29) = "`ZeitpunktLaborneu`"
 Str(0, 94, 30) = "`ZdüP`"
 Str(0, 94, 31) = "`ZdiP`"
 Str(0, 94, 32) = "`LWerte`"
 Str(0, 94, 33) = "`verglichen`"
 Str(0, 94, 34) = "`AfN`"
 Str(0, 94, 35) = "`RefNr`"
 Str(0, 94, 36) = "`RefNr`"
 Str(0, 94, 37) = "`DatID`"
 Str(0, 94, 38) = "`LaborXEingelLaborXUS`"
 Str(0, 94, 39) = "`LaborXSätzeLaborXUS`"
 Str(0, 94, 40) = "`Name`"
 Str(0, 94, 41) = "`Pat_id`"
 Str(0, 94, 42) = "`SatzID`"
 Str(0, 94, 43) = "`LaborXEingelLaborXUS_AccRel`"
 Str(0, 94, 44) = "`LaborXSätzeLaborXUS_AccRel`"
 ArtZ(0, 94) = 34
 ArtZ(1, 94) = 8
 ArtZ(2, 94) = 2
 Str(1, 94, 0) = "CREATE TABLE `laborxus` ("
 Str(1, 94, 1) = " `RefNr` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Bezug auf LaborWert'"
 Str(1, 94, 2) = " `DatID` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborEingelesen'"
 Str(1, 94, 3) = " `SatzID` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborXSätze'"
 Str(1, 94, 4) = " `Satzart` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8000 Satzart (Turbomed)'"
 Str(1, 94, 5) = " `Satzlänge` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8100 Satzlänge (Turbomed)'"
 Str(1, 94, 6) = " `Auftragsnummer` varchar(11) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8310 Anforderungsident (Turbomed)'"
 Str(1, 94, 7) = " `Auftragsschlüssel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8311 Anforderungsnr d Labors (Turbomed)'"
 Str(1, 94, 8) = " `Eingang` datetime DEFAULT NULL COMMENT '8301 Eingangsdatum in Datumsform'"
 Str(1, 94, 9) = " `Berichtsdatum` varchar(13) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8302 Berichtsdatum'"
 Str(1, 94, 10) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 94, 11) = " `Nachname` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 94, 12) = " `Vorname` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 94, 13) = " `GebDat` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3103'"
 Str(1, 94, 14) = " `Titel` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3104'"
 Str(1, 94, 15) = " `NVorsatz` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3100'"
 Str(1, 94, 16) = " `BefArt` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8401 Befundart (Turbomed) / Fertigstellungsgrad (""E""=Endbefund, ""T"" = Teilbefund)'"
 Str(1, 94, 17) = " `Abrechnungstyp` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)'"
 Str(1, 94, 18) = " `GebüOrd` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8403 Gebührenordnung (Turbomed)'"
 Str(1, 94, 19) = " `Patienteninformation` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8405 Patienteninformation (Turbomed)'"
 Str(1, 94, 20) = " `Geschlecht` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8407 Geschlecht (Turbomed)'"
 Str(1, 94, 21) = " `AuftrHinw` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8490 Auftragsbezogene Hinweise (Turbomed)'"
 Str(1, 94, 22) = " `Pat_idUrsp` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor'"
 Str(1, 94, 23) = " `Pat_idErwVNG` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag'"
 Str(1, 94, 24) = " `Pat_idErwVN` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Vornamen und Nachnamen'"
 Str(1, 94, 25) = " `Pat_idErwG` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag'"
 Str(1, 94, 26) = " `Pat_idErwGB` varchar(23) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung'"
 Str(1, 94, 27) = " `Pat_idErwGL` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor'"
 Str(1, 94, 28) = " `Pat_idLaborNeu` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Pat_ids von in Laborneu zuordnbaren Patienten'"
 Str(1, 94, 29) = " `ZeitpunktLaborneu` datetime DEFAULT NULL COMMENT 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde'"
 Str(1, 94, 30) = " `ZdüP` smallint(6) DEFAULT NULL COMMENT 'Zahl der verglichenen Parameter'"
 Str(1, 94, 31) = " `ZdiP` int(10) DEFAULT NULL COMMENT 'Zahl der infragekommenden Patienten'"
 Str(1, 94, 32) = " `LWerte` longtext COLLATE latin1_german2_ci COMMENT 'Laborwerte, die zur Zuordnung geführt haben'"
 Str(1, 94, 33) = " `verglichen` datetime DEFAULT NULL COMMENT 'Datum, zu dem Datensatz zuletzt verglichen wurde'"
 Str(1, 94, 34) = " `AfN` smallint(6) DEFAULT NULL COMMENT 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu'"
 Str(1, 94, 35) = "  PRIMARY KEY (`RefNr`)"
 Str(1, 94, 36) = "  UNIQUE KEY `RefNr` (`RefNr`)"
 Str(1, 94, 37) = "  KEY `DatID` (`DatID`)"
 Str(1, 94, 38) = "  KEY `LaborXEingelLaborXUS` (`DatID`)"
 Str(1, 94, 39) = "  KEY `LaborXSätzeLaborXUS` (`SatzID`)"
 Str(1, 94, 40) = "  KEY `Name` (`Nachname`,`Vorname`)"
 Str(1, 94, 41) = "  KEY `Pat_id` (`Pat_id`)"
 Str(1, 94, 42) = "  KEY `SatzID` (`SatzID`)"
 Str(1, 94, 43) = "  CONSTRAINT `LaborXEingelLaborXUS_AccRel` FOREIGN KEY (`DatID`) REFERENCES `laborxeingel` (`DatID`) ON UPDATE CASCADE"
 Str(1, 94, 44) = "  CONSTRAINT `LaborXSätzeLaborXUS_AccRel` FOREIGN KEY (`SatzID`) REFERENCES `laborxsaetze` (`SatzID`) ON UPDATE CASCADE"
 Str(1, 94, 45) = " ENGINE=InnoDB AUTO_INCREMENT=3619 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr94

Sub FüllStr95()
 Str(0, 95, 0) = "laborxwert"
 Str(0, 95, 1) = "`RefNr`"
 Str(0, 95, 2) = "`Abkü`"
 Str(0, 95, 3) = "`Langname`"
 Str(0, 95, 4) = "`Quelle`"
 Str(0, 95, 5) = "`QSpez`"
 Str(0, 95, 6) = "`AbnDat`"
 Str(0, 95, 7) = "`Wert`"
 Str(0, 95, 8) = "`Einheit`"
 Str(0, 95, 9) = "`Grenzwerti`"
 Str(0, 95, 10) = "`Kommentar`"
 Str(0, 95, 11) = "`Teststatus`"
 Str(0, 95, 12) = "`Erklärung`"
 Str(0, 95, 13) = "`Normbereich`"
 Str(0, 95, 14) = "`NormU`"
 Str(0, 95, 15) = "`NormO`"
 Str(0, 95, 16) = "`AuftrHinw`"
 Str(0, 95, 17) = "`LaborXUSLaborXWert`"
 Str(0, 95, 18) = "`LaborXWertAbkü`"
 Str(0, 95, 19) = "`LaborXWertUS`"
 Str(0, 95, 20) = "`RefNr`"
 Str(0, 95, 21) = "`LaborParameterLaborXWert_AccRel`"
 Str(0, 95, 22) = "`LaborXUSLaborXWert_AccRel`"
 Str(0, 95, 23) = "`LaborXWertUS_AccRel`"
 ArtZ(0, 95) = 16
 ArtZ(1, 95) = 4
 ArtZ(2, 95) = 3
 Str(1, 95, 0) = "CREATE TABLE `laborxwert` ("
 Str(1, 95, 1) = " `RefNr` int(10) DEFAULT NULL COMMENT 'Bezug auf LaborUS'"
 Str(1, 95, 2) = " `Abkü` varchar(16) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8410 Test-Ident  (Turbomed)'"
 Str(1, 95, 3) = " `Langname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8411 Testbezeichnung (Turbomed)'"
 Str(1, 95, 4) = " `Quelle` varchar(43) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8430 Probenmaterial-Bezeichnung (Turbomed)'"
 Str(1, 95, 5) = " `QSpez` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8431 Probenmaterial-Spezifikation (Turbomed)'"
 Str(1, 95, 6) = " `AbnDat` datetime DEFAULT NULL COMMENT '8432 Abnahmedatum (Turbomed)'"
 Str(1, 95, 7) = " `Wert` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8420 Ergebniswert (Turbomed)'"
 Str(1, 95, 8) = " `Einheit` varchar(15) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8421 Einheit (Turbomed)'"
 Str(1, 95, 9) = " `Grenzwerti` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8422 Grenzwertindikator (Turbomed)'"
 Str(1, 95, 10) = " `Kommentar` varchar(1385) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8480 Ergebnistext (Turbomed)'"
 Str(1, 95, 11) = " `Teststatus` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8418 Teststatus (Turbomed)'"
 Str(1, 95, 12) = " `Erklärung` varchar(759) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8470 Testbezogene Hinweise (Turbomed)'"
 Str(1, 95, 13) = " `Normbereich` varchar(65) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8460 Normalwert-Text (Turbomed)'"
 Str(1, 95, 14) = " `NormU` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8461 Normuntergrenze'"
 Str(1, 95, 15) = " `NormO` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8461 Normobergrenze'"
 Str(1, 95, 16) = " `AuftrHinw` varchar(180) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '8490 Auftragsbezogene Hinweise (Turbomed)'"
 Str(1, 95, 17) = "  KEY `LaborXUSLaborXWert` (`RefNr`)"
 Str(1, 95, 18) = "  KEY `LaborXWertAbkü` (`Abkü`,`Einheit`)"
 Str(1, 95, 19) = "  KEY `LaborXWertUS` (`RefNr`)"
 Str(1, 95, 20) = "  KEY `RefNr` (`RefNr`)"
 Str(1, 95, 21) = "  CONSTRAINT `LaborParameterLaborXWert_AccRel` FOREIGN KEY (`Abkü`, `Einheit`) REFERENCES `laborparameter` (`Abkü`, `Einheit`)"
 Str(1, 95, 22) = "  CONSTRAINT `LaborXUSLaborXWert_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 95, 23) = "  CONSTRAINT `LaborXWertUS_AccRel` FOREIGN KEY (`RefNr`) REFERENCES `laborxus` (`RefNr`) ON UPDATE CASCADE"
 Str(1, 95, 24) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr95

Sub FüllStr96()
 Str(0, 96, 0) = "lbanforderungen"
 Str(0, 96, 1) = "`FID`"
 Str(0, 96, 2) = "`Pat_ID`"
 Str(0, 96, 3) = "`ZeitPunkt`"
 Str(0, 96, 4) = "`AnfText`"
 Str(0, 96, 5) = "`absPos`"
 Str(0, 96, 6) = "`AktZeit`"
 Str(0, 96, 7) = "`StByte`"
 Str(0, 96, 8) = "`Auswahl`"
 Str(0, 96, 9) = "`FälleLbAnforderungen`"
 Str(0, 96, 10) = "`FID`"
 Str(0, 96, 11) = "`NamenLbAnforderungen`"
 Str(0, 96, 12) = "`FälleLbAnforderungen_AccRel`"
 Str(0, 96, 13) = "`NamenLbAnforderungen_AccRel`"
 ArtZ(0, 96) = 7
 ArtZ(1, 96) = 4
 ArtZ(2, 96) = 2
 Str(1, 96, 0) = "CREATE TABLE `lbanforderungen` ("
 Str(1, 96, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 96, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 96, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 96, 4) = " `AnfText` longtext COLLATE latin1_german2_ci COMMENT '6280'"
 Str(1, 96, 5) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 96, 6) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 96, 7) = " `StByte` int(11) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 96, 8) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`AnfText`(255))"
 Str(1, 96, 9) = "  KEY `FälleLbAnforderungen` (`FID`)"
 Str(1, 96, 10) = "  KEY `FID` (`FID`)"
 Str(1, 96, 11) = "  KEY `NamenLbAnforderungen` (`Pat_ID`)"
 Str(1, 96, 12) = "  CONSTRAINT `FälleLbAnforderungen_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 96, 13) = "  CONSTRAINT `NamenLbAnforderungen_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 96, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr96

Sub FüllStr97()
 Str(0, 97, 0) = "leistungen"
 Str(0, 97, 1) = "`FID`"
 Str(0, 97, 2) = "`Pat_ID`"
 Str(0, 97, 3) = "`ZeitPunkt`"
 Str(0, 97, 4) = "`Leistung`"
 Str(0, 97, 5) = "`f5002`"
 Str(0, 97, 6) = "`f5005`"
 Str(0, 97, 7) = "`f5006`"
 Str(0, 97, 8) = "`f5009`"
 Str(0, 97, 9) = "`Med`"
 Str(0, 97, 10) = "`f5015`"
 Str(0, 97, 11) = "`f5016`"
 Str(0, 97, 12) = "`f5021`"
 Str(0, 97, 13) = "`f5026`"
 Str(0, 97, 14) = "`Faktor`"
 Str(0, 97, 15) = "`f5098`"
 Str(0, 97, 16) = "`LANR`"
 Str(0, 97, 17) = "`letzVorg`"
 Str(0, 97, 18) = "`Ausn`"
 Str(0, 97, 19) = "`Beme`"
 Str(0, 97, 20) = "`absPos`"
 Str(0, 97, 21) = "`AktZeit`"
 Str(0, 97, 22) = "`QS`"
 Str(0, 97, 23) = "`QT`"
 Str(0, 97, 24) = "`StByte`"
 Str(0, 97, 25) = "`Auswahl`"
 Str(0, 97, 26) = "`FälleLeistungen`"
 Str(0, 97, 27) = "`FID`"
 Str(0, 97, 28) = "`Leistung`"
 Str(0, 97, 29) = "`NamenLeistungen`"
 Str(0, 97, 30) = "`FälleLeistungen_AccRel`"
 Str(0, 97, 31) = "`NamenLeistungen_AccRel`"
 ArtZ(0, 97) = 24
 ArtZ(1, 97) = 5
 ArtZ(2, 97) = 2
 Str(1, 97, 0) = "CREATE TABLE `leistungen` ("
 Str(1, 97, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 97, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 97, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '5000 + 6201'"
 Str(1, 97, 4) = " `Leistung` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5001'"
 Str(1, 97, 5) = " `f5002` varchar(32) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5002'"
 Str(1, 97, 6) = " `f5005` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5005'"
 Str(1, 97, 7) = " `f5006` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5006'"
 Str(1, 97, 8) = " `f5009` varchar(80) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5009'"
 Str(1, 97, 9) = " `Med` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5010'"
 Str(1, 97, 10) = " `f5015` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5015'"
 Str(1, 97, 11) = " `f5016` varchar(28) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5016'"
 Str(1, 97, 12) = " `f5021` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5021'"
 Str(1, 97, 13) = " `f5026` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5026'"
 Str(1, 97, 14) = " `Faktor` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5062 Multiplikator für GOÄ-Rechnung'"
 Str(1, 97, 15) = " `f5098` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5098 0000000000'"
 Str(1, 97, 16) = " `LANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '5099 LANR'"
 Str(1, 97, 17) = " `letzVorg` datetime DEFAULT NULL COMMENT '5101 letzter Vorgang'"
 Str(1, 97, 18) = " `Ausn` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3677 Ausnahme/Begründung für abweichendes Geschlecht'"
 Str(1, 97, 19) = " `Beme` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '         Bemerkung'"
 Str(1, 97, 20) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 97, 21) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 97, 22) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 97, 23) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 97, 24) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 97, 25) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Leistung`)"
 Str(1, 97, 26) = "  KEY `FälleLeistungen` (`FID`)"
 Str(1, 97, 27) = "  KEY `FID` (`FID`)"
 Str(1, 97, 28) = "  KEY `Leistung` (`Leistung`)"
 Str(1, 97, 29) = "  KEY `NamenLeistungen` (`Pat_ID`)"
 Str(1, 97, 30) = "  CONSTRAINT `FälleLeistungen_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 97, 31) = "  CONSTRAINT `NamenLeistungen_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 97, 32) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr97

Sub FüllStr98()
 Str(0, 98, 0) = "leistungen exportiert"
 Str(0, 98, 1) = "`ID`"
 Str(0, 98, 2) = "`Datum`"
 Str(0, 98, 3) = "`Pat_id`"
 Str(0, 98, 4) = "`SchGr`"
 Str(0, 98, 5) = "`Leistung`"
 Str(0, 98, 6) = "`übertragen`"
 Str(0, 98, 7) = "`ID`"
 Str(0, 98, 8) = "`PrimaryKey`"
 Str(0, 98, 9) = "`ID`"
 ArtZ(0, 98) = 6
 ArtZ(1, 98) = 3
 Str(1, 98, 0) = "CREATE TABLE `leistungen exportiert` ("
 Str(1, 98, 1) = " `ID` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Reihenfolge'"
 Str(1, 98, 2) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum'"
 Str(1, 98, 3) = " `Pat_id` int(10) DEFAULT NULL COMMENT '-> Namen!Pat_id'"
 Str(1, 98, 4) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 98, 5) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ziffer aus EBM / GOÄ'"
 Str(1, 98, 6) = " `übertragen` datetime DEFAULT NULL COMMENT '"""", übertragen'"
 Str(1, 98, 7) = "  PRIMARY KEY (`ID`)"
 Str(1, 98, 8) = "  UNIQUE KEY `PrimaryKey` (`ID`)"
 Str(1, 98, 9) = "  KEY `ID` (`Pat_id`)"
 Str(1, 98, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr98

Sub FüllStr99()
 Str(0, 99, 0) = "leistungsexport"
 Str(0, 99, 1) = "`Prim`"
 Str(0, 99, 2) = "`PatID`"
 Str(0, 99, 3) = "`Datum`"
 Str(0, 99, 4) = "`UZeit`"
 Str(0, 99, 5) = "`Leistung`"
 Str(0, 99, 6) = "`SchGr`"
 Str(0, 99, 7) = "`Status`"
 Str(0, 99, 8) = "`Prim`"
 Str(0, 99, 9) = "`PatID`"
 ArtZ(0, 99) = 7
 ArtZ(1, 99) = 2
 Str(1, 99, 0) = "CREATE TABLE `leistungsexport` ("
 Str(1, 99, 1) = " `Prim` int(10) NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel wegen Reihenfolge'"
 Str(1, 99, 2) = " `PatID` int(10) DEFAULT NULL COMMENT '-> Namen!Pat_id'"
 Str(1, 99, 3) = " `Datum` datetime DEFAULT NULL COMMENT 'Leistungsdatum'"
 Str(1, 99, 4) = " `UZeit` datetime DEFAULT NULL COMMENT 'Leistungszeit, falls gewünscht'"
 Str(1, 99, 5) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Ziffer aus EBM / GOÄ'"
 Str(1, 99, 6) = " `SchGr` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Scheingruppe (""24"" = Mitbehandlung, ""00"" = eigener Patient, ""41"" = Vertretung, ""90""=privat)'"
 Str(1, 99, 7) = " `Status` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '"""", übertragen'"
 Str(1, 99, 8) = "  PRIMARY KEY (`Prim`)"
 Str(1, 99, 9) = "  KEY `PatID` (`PatID`)"
 Str(1, 99, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr99

Sub FüllStr100()
 Str(0, 100, 0) = "letze faelle"
 Str(1, 100, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `letze faelle` AS select `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik` from (`_lfaelle` `l` left join `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`))))"
End Sub ' FüllStr100

Sub FüllStr101()
 Str(0, 101, 0) = "letzte faelle"
 Str(0, 101, 1) = "`pat_id`"
 Str(0, 101, 2) = "`fid`"
 Str(0, 101, 3) = "`schgr`"
 Str(0, 101, 4) = "`bhfb`"
 Str(0, 101, 5) = "`ik`"
 ArtZ(0, 101) = 5
 Str(1, 101, 0) = "CREATE TABLE `letzte faelle` ("
 Str(1, 101, 1) = " `pat_id` int(10) DEFAULT NULL"
 Str(1, 101, 2) = " `fid` int(10) DEFAULT NULL"
 Str(1, 101, 3) = " `schgr` int(10) DEFAULT NULL"
 Str(1, 101, 4) = " `bhfb` datetime DEFAULT NULL"
 Str(1, 101, 5) = " `ik` varchar(50) DEFAULT NULL"
 Str(1, 101, 6) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr101

Sub FüllStr102()
 Str(0, 102, 0) = "letztefaelleverschieden"
 Str(1, 102, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `letztefaelleverschieden` AS select `i`.`bhfb` AS `bhfb`,`i`.`pat_id` AS `pat_id` from `_f1` `i` group by `i`.`pat_id`"
End Sub ' FüllStr102

Sub FüllStr103()
 Str(0, 103, 0) = "lfaelle"
 Str(1, 103, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `lfaelle` AS select `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr`,`f`.`ÜbwV` AS `übwv` from (`_lfaelle` `l` left join `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`))))"
End Sub ' FüllStr103

Sub FüllStr104()
 Str(0, 104, 0) = "lfaellev"
 Str(1, 104, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `lfaellev` AS select `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik`,`f`.`VKNr` AS `vknr` from (`_lfaelle` `l` left join `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) and (`f`.`BhFB` = `l`.`mbhfb`)))) group by `f`.`Pat_ID`"
End Sub ' FüllStr104

Sub FüllStr105()
 Str(0, 105, 0) = "listenausgabeuew"
 Str(0, 105, 1) = "`name`"
 Str(0, 105, 2) = "`vorname`"
 Str(0, 105, 3) = "`titelt`"
 Str(0, 105, 4) = "`fachgruppe`"
 Str(0, 105, 5) = "`strasse`"
 Str(0, 105, 6) = "`plz`"
 Str(0, 105, 7) = "`ort`"
 Str(0, 105, 8) = "`telefon`"
 Str(0, 105, 9) = "`fax`"
 Str(0, 105, 10) = "`kvnr`"
 Str(0, 105, 11) = "`aktdat`"
 Str(0, 105, 12) = "`id`"
 Str(0, 105, 13) = "`QZ`"
 Str(0, 105, 14) = "`überschrift`"
 Str(0, 105, 15) = "`dbnr`"
 Str(0, 105, 16) = "`bstelle`"
 Str(0, 105, 17) = "`anrede`"
 Str(0, 105, 18) = "`tel2`"
 Str(0, 105, 19) = "`tel3`"
 Str(0, 105, 20) = "`tel4`"
 Str(0, 105, 21) = "`fax2`"
 Str(0, 105, 22) = "`fax3`"
 Str(0, 105, 23) = "`email`"
 Str(0, 105, 24) = "`zulg`"
 Str(0, 105, 25) = "`arzttyp`"
 Str(0, 105, 26) = "`gemmit`"
 Str(0, 105, 27) = "`beme`"
 Str(0, 105, 28) = "`dmpt2`"
 Str(0, 105, 29) = "`dmpt1`"
 Str(0, 105, 30) = "`geschlecht`"
 Str(0, 105, 31) = "`titel`"
 Str(0, 105, 32) = "`id`"
 Str(0, 105, 33) = "`kvnr`"
 Str(0, 105, 34) = "`name`"
 Str(0, 105, 35) = "`fax`"
 ArtZ(0, 105) = 31
 ArtZ(1, 105) = 4
 Str(1, 105, 0) = "CREATE TABLE `listenausgabeuew` ("
 Str(1, 105, 1) = " `name` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 2) = " `vorname` varchar(35) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 3) = " `titelt` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 4) = " `fachgruppe` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 5) = " `strasse` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 6) = " `plz` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 7) = " `ort` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 8) = " `telefon` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 105, 9) = " `fax` varchar(40) CHARACTER SET utf8 DEFAULT NULL"
 Str(1, 105, 10) = " `kvnr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 11) = " `aktdat` datetime DEFAULT NULL"
 Str(1, 105, 12) = " `id` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 105, 13) = " `QZ` tinyint(1) unsigned DEFAULT NULL COMMENT 'ob Mitglied im Qualitätszirkel Stoffwechsel- und Gefäßerkrankungen'"
 Str(1, 105, 14) = " `überschrift` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 15) = " `dbnr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 16) = " `bstelle` varchar(53) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 17) = " `anrede` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 18) = " `tel2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 19) = " `tel3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 20) = " `tel4` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 21) = " `fax2` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 22) = " `fax3` varchar(40) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 23) = " `email` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 24) = " `zulg` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 25) = " `arzttyp` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 26) = " `gemmit` longtext COLLATE latin1_german2_ci"
 Str(1, 105, 27) = " `beme` longtext COLLATE latin1_german2_ci"
 Str(1, 105, 28) = " `dmpt2` bit(1) DEFAULT NULL"
 Str(1, 105, 29) = " `dmpt1` bit(1) DEFAULT NULL"
 Str(1, 105, 30) = " `geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 31) = " `titel` varchar(90) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 105, 32) = "  PRIMARY KEY (`id`)"
 Str(1, 105, 33) = "  KEY `kvnr` (`kvnr`)"
 Str(1, 105, 34) = "  KEY `name` (`name`)"
 Str(1, 105, 35) = "  KEY `fax` (`fax`)"
 Str(1, 105, 36) = " ENGINE=InnoDB AUTO_INCREMENT=979 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr105

Sub FüllStr106()
 Str(0, 106, 0) = "lmp"
 Str(1, 106, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `lmp` AS select `mp`.`FID` AS `FID`,`mp`.`Pat_ID` AS `Pat_ID`,`mp`.`MPNr` AS `MPNr`,`mp`.`ZeitPunkt` AS `ZeitPunkt`,`mp`.`Datum` AS `Datum`,`mp`.`Medikament` AS `Medikament`,`mp`.`MedAnfang` AS `MedAnfang`,`mp`.`FeldNr` AS `FeldNr`,`mp`.`mo` AS `mo`,`mp`.`mi` AS `mi`,`mp`.`nm` AS `nm`,`mp`.`ab` AS `ab`,`mp`.`zn` AS `zn`,`mp`.`bBed` AS `bBed`,`mp`.`Bemerkung` AS `Bemerkung`,`mp`.`AbsPos` AS `AbsPos`,`mp`.`AktZeit` AS `AktZeit`,`mp`.`StByte` AS `StByte` from (`_fuerlmp` `i` join `medplan` `mp` on(((`i`.`pat_id` = `mp`.`Pat_ID`) and (`i`.`mpnr` = `mp`.`MPNr`))))"
End Sub ' FüllStr106

Sub FüllStr107()
 Str(0, 107, 0) = "medarten"
 Str(0, 107, 1) = "`Medikament`"
 Str(0, 107, 2) = "`Langname`"
 Str(0, 107, 3) = "`Pat_ID`"
 Str(0, 107, 4) = "`Anzahl`"
 Str(0, 107, 5) = "`Glib`"
 Str(0, 107, 6) = "`Metf`"
 Str(0, 107, 7) = "`GlucI`"
 Str(0, 107, 8) = "`SHGlin`"
 Str(0, 107, 9) = "`Glit`"
 Str(0, 107, 10) = "`SonstAD`"
 Str(0, 107, 11) = "`Ins`"
 Str(0, 107, 12) = "`Anal`"
 Str(0, 107, 13) = "`InsArt`"
 Str(0, 107, 14) = "`HMG`"
 Str(0, 107, 15) = "`Hypt`"
 Str(0, 107, 16) = "`Thro`"
 Str(0, 107, 17) = "`Antib`"
 Str(0, 107, 18) = "`and`"
 Str(0, 107, 19) = "`hinzugefügt`"
 Str(0, 107, 20) = "`Tstr`"
 Str(0, 107, 21) = "`Puzu`"
 Str(0, 107, 22) = "`VMat`"
 Str(0, 107, 23) = "`PenN`"
 Str(0, 107, 24) = "`Neurp`"
 Str(0, 107, 25) = "`AutNP`"
 Str(0, 107, 26) = "`Fetts`"
 Str(0, 107, 27) = "`Hsre`"
 Str(0, 107, 28) = "`AntiMyk`"
 Str(0, 107, 29) = "`Glauk`"
 Str(0, 107, 30) = "`COLD`"
 Str(0, 107, 31) = "`Pros`"
 Str(0, 107, 32) = "`Urä`"
 Str(0, 107, 33) = "`HyThy`"
 Str(0, 107, 34) = "`Ostp`"
 Str(0, 107, 35) = "`KHK`"
 Str(0, 107, 36) = "`HerzI`"
 Str(0, 107, 37) = "`Stru`"
 Str(0, 107, 38) = "`AVK`"
 Str(0, 107, 39) = "`PanI`"
 Str(0, 107, 40) = "`Vari`"
 Str(0, 107, 41) = "`Östr`"
 Str(0, 107, 42) = "`AntiDep`"
 Str(0, 107, 43) = "`AntiDem`"
 Str(0, 107, 44) = "`AntiEp`"
 Str(0, 107, 45) = "`Park`"
 Str(0, 107, 46) = "`AntiPern`"
 Str(0, 107, 47) = "`Appet`"
 Str(0, 107, 48) = "`Anäm`"
 Str(0, 107, 49) = "`Antiherp`"
 Str(0, 107, 50) = "`NSAR`"
 Str(0, 107, 51) = "`Antikoag`"
 Str(0, 107, 52) = "`Betabl`"
 Str(0, 107, 53) = "`ACEH`"
 Str(0, 107, 54) = "`AT1`"
 Str(0, 107, 55) = "`CalcA`"
 Str(0, 107, 56) = "`Diur`"
 Str(0, 107, 57) = "`falsch`"
 Str(0, 107, 58) = "`ID`"
 Str(0, 107, 59) = "`ID`"
 Str(0, 107, 60) = "`Medikament`"
 Str(0, 107, 61) = "`pat_id`"
 Str(0, 107, 62) = "`NamenMedArten_AccRel`"
 ArtZ(0, 107) = 58
 ArtZ(1, 107) = 3
 ArtZ(2, 107) = 1
 Str(1, 107, 0) = "CREATE TABLE `medarten` ("
 Str(1, 107, 1) = " `Medikament` varchar(50) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL"
 Str(1, 107, 2) = " `Langname` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Beispiel-Langname'"
 Str(1, 107, 3) = " `Pat_ID` int(10) DEFAULT NULL COMMENT 'Beispiel-PatID'"
 Str(1, 107, 4) = " `Anzahl` int(10) DEFAULT NULL COMMENT 'Anzahl der Vorkommen'"
 Str(1, 107, 5) = " `Glib` tinyint(1) unsigned NOT NULL COMMENT 'Glibenclamid'"
 Str(1, 107, 6) = " `Metf` tinyint(1) unsigned NOT NULL COMMENT 'Metformin'"
 Str(1, 107, 7) = " `GlucI` tinyint(1) unsigned NOT NULL COMMENT 'Glucosidase-Inhibitoren'"
 Str(1, 107, 8) = " `SHGlin` tinyint(1) unsigned NOT NULL COMMENT 'andere Sulfonylharnstoffe oder Glinide'"
 Str(1, 107, 9) = " `Glit` tinyint(1) unsigned NOT NULL COMMENT 'Glitanzone'"
 Str(1, 107, 10) = " `SonstAD` tinyint(1) unsigned NOT NULL COMMENT 'Sonstige antidiabetische Medikation'"
 Str(1, 107, 11) = " `Ins` tinyint(1) unsigned NOT NULL COMMENT 'Insulin'"
 Str(1, 107, 12) = " `Anal` tinyint(1) unsigned NOT NULL COMMENT 'Insulin-Analoga'"
 Str(1, 107, 13) = " `InsArt` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '1= schnell, 2 = langsam, 3 = Misch'"
 Str(1, 107, 14) = " `HMG` tinyint(1) unsigned NOT NULL COMMENT 'HMG-CoA-Reduktase-Inhibitoren'"
 Str(1, 107, 15) = " `Hypt` tinyint(1) unsigned NOT NULL COMMENT 'Hypertonie-Mittel'"
 Str(1, 107, 16) = " `Thro` tinyint(1) unsigned NOT NULL COMMENT 'Thrombozyten-Hemmer'"
 Str(1, 107, 17) = " `Antib` tinyint(1) unsigned NOT NULL COMMENT 'Antibiotika'"
 Str(1, 107, 18) = " `and` tinyint(1) unsigned NOT NULL COMMENT 'andere'"
 Str(1, 107, 19) = " `hinzugefügt` datetime DEFAULT NULL"
 Str(1, 107, 20) = " `Tstr` tinyint(1) unsigned NOT NULL COMMENT 'Teststreifen'"
 Str(1, 107, 21) = " `Puzu` tinyint(1) unsigned NOT NULL COMMENT 'Pumpenzubehör'"
 Str(1, 107, 22) = " `VMat` tinyint(1) unsigned NOT NULL COMMENT 'Verbandsmaterial'"
 Str(1, 107, 23) = " `PenN` tinyint(1) unsigned NOT NULL COMMENT 'Pennadeln'"
 Str(1, 107, 24) = " `Neurp` tinyint(1) unsigned NOT NULL COMMENT 'Neuropathie-Behandlungsmittel'"
 Str(1, 107, 25) = " `AutNP` tinyint(1) unsigned NOT NULL COMMENT 'Autonome Neuropathie'"
 Str(1, 107, 26) = " `Fetts` tinyint(1) unsigned NOT NULL COMMENT 'Fibrate, Ezetrol, Niaspan u.a.'"
 Str(1, 107, 27) = " `Hsre` tinyint(1) unsigned NOT NULL COMMENT 'Hyperuriämie-Mittel'"
 Str(1, 107, 28) = " `AntiMyk` tinyint(1) unsigned NOT NULL COMMENT 'Antimykotika'"
 Str(1, 107, 29) = " `Glauk` tinyint(1) unsigned NOT NULL COMMENT 'Glaukom'"
 Str(1, 107, 30) = " `COLD` tinyint(1) unsigned NOT NULL COMMENT 'COLD und Asthma'"
 Str(1, 107, 31) = " `Pros` tinyint(1) unsigned NOT NULL COMMENT 'Prostatahypertrophie'"
 Str(1, 107, 32) = " `Urä` tinyint(1) unsigned NOT NULL COMMENT 'Urämie-spezifische'"
 Str(1, 107, 33) = " `HyThy` tinyint(1) unsigned NOT NULL COMMENT 'thyreostatische Mittel'"
 Str(1, 107, 34) = " `Ostp` tinyint(1) unsigned NOT NULL COMMENT 'Osteoporosemittel'"
 Str(1, 107, 35) = " `KHK` tinyint(1) unsigned NOT NULL COMMENT 'KHK-spezifisch'"
 Str(1, 107, 36) = " `HerzI` tinyint(1) unsigned NOT NULL COMMENT 'Herzinsuffizienz-spezifische'"
 Str(1, 107, 37) = " `Stru` tinyint(1) unsigned NOT NULL COMMENT 'Struma- und Hypothyreosemittel'"
 Str(1, 107, 38) = " `AVK` tinyint(1) unsigned NOT NULL COMMENT 'AVK-Mittel'"
 Str(1, 107, 39) = " `PanI` tinyint(1) unsigned NOT NULL COMMENT 'Pankreasinsuffizienz'"
 Str(1, 107, 40) = " `Vari` tinyint(1) unsigned NOT NULL COMMENT 'Varikosemittel'"
 Str(1, 107, 41) = " `Östr` tinyint(1) unsigned NOT NULL COMMENT 'Östrogene, Gestagene usw.'"
 Str(1, 107, 42) = " `AntiDep` tinyint(1) unsigned NOT NULL COMMENT 'Antidepressiva'"
 Str(1, 107, 43) = " `AntiDem` tinyint(1) unsigned NOT NULL COMMENT 'Antidementika'"
 Str(1, 107, 44) = " `AntiEp` tinyint(1) unsigned NOT NULL COMMENT 'Antiepileptika'"
 Str(1, 107, 45) = " `Park` tinyint(1) unsigned NOT NULL COMMENT 'Parkinson-Medikament'"
 Str(1, 107, 46) = " `AntiPern` tinyint(1) unsigned NOT NULL COMMENT 'Antiperniziosa'"
 Str(1, 107, 47) = " `Appet` tinyint(1) unsigned NOT NULL COMMENT 'Appetitzügler'"
 Str(1, 107, 48) = " `Anäm` tinyint(1) unsigned NOT NULL COMMENT 'Anämiebehandlungsmittel'"
 Str(1, 107, 49) = " `Antiherp` tinyint(1) unsigned NOT NULL COMMENT 'Anti-Herpes-Mittel'"
 Str(1, 107, 50) = " `NSAR` tinyint(1) unsigned NOT NULL COMMENT 'NSAR'"
 Str(1, 107, 51) = " `Antikoag` tinyint(1) unsigned NOT NULL COMMENT 'Antikoagulatien'"
 Str(1, 107, 52) = " `Betabl` tinyint(1) unsigned NOT NULL COMMENT 'Betablocker'"
 Str(1, 107, 53) = " `ACEH` tinyint(1) unsigned NOT NULL COMMENT 'ACE-Hemmer'"
 Str(1, 107, 54) = " `AT1` tinyint(1) unsigned NOT NULL COMMENT 'AT-1-Blocker'"
 Str(1, 107, 55) = " `CalcA` tinyint(1) unsigned NOT NULL COMMENT 'Calcium-Antagonist'"
 Str(1, 107, 56) = " `Diur` tinyint(1) unsigned NOT NULL COMMENT 'Diuretikum'"
 Str(1, 107, 57) = " `falsch` tinyint(1) unsigned NOT NULL COMMENT 'falsch geschrieben oder nicht erkennbar'"
 Str(1, 107, 58) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Primärschlüssel'"
 Str(1, 107, 59) = "  PRIMARY KEY (`ID`)"
 Str(1, 107, 60) = "  UNIQUE KEY `Medikament` (`Medikament`)"
 Str(1, 107, 61) = "  KEY `pat_id` (`Pat_ID`)"
 Str(1, 107, 62) = "  CONSTRAINT `NamenMedArten_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`)"
 Str(1, 107, 63) = " ENGINE=InnoDB AUTO_INCREMENT=16557 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr107

Sub FüllStr108()
 Str(0, 108, 0) = "medplan"
 Str(0, 108, 1) = "`FID`"
 Str(0, 108, 2) = "`Pat_ID`"
 Str(0, 108, 3) = "`MPNr`"
 Str(0, 108, 4) = "`ZeitPunkt`"
 Str(0, 108, 5) = "`Datum`"
 Str(0, 108, 6) = "`Medikament`"
 Str(0, 108, 7) = "`MedAnfang`"
 Str(0, 108, 8) = "`FeldNr`"
 Str(0, 108, 9) = "`mo`"
 Str(0, 108, 10) = "`mi`"
 Str(0, 108, 11) = "`nm`"
 Str(0, 108, 12) = "`ab`"
 Str(0, 108, 13) = "`zn`"
 Str(0, 108, 14) = "`bBed`"
 Str(0, 108, 15) = "`Bemerkung`"
 Str(0, 108, 16) = "`AbsPos`"
 Str(0, 108, 17) = "`AktZeit`"
 Str(0, 108, 18) = "`StByte`"
 Str(0, 108, 19) = "`FälleMedPlan`"
 Str(0, 108, 20) = "`MedPlanMedikament`"
 Str(0, 108, 21) = "`NamenMedPlan`"
 Str(0, 108, 22) = "`MedArtenMedPlan_AccRel`"
 Str(0, 108, 23) = "`MPNr`"
 Str(0, 108, 24) = "`MPNrPat_ID`"
 Str(0, 108, 25) = "`Auswahl`"
 Str(0, 108, 26) = "`FälleMedPlan_AccRel`"
 Str(0, 108, 27) = "`MedArtenMedPlan_AccRel`"
 Str(0, 108, 28) = "`NamenMedPlan_AccRel`"
 ArtZ(0, 108) = 18
 ArtZ(1, 108) = 7
 ArtZ(2, 108) = 3
 Str(1, 108, 0) = "CREATE TABLE `medplan` ("
 Str(1, 108, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 108, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 108, 3) = " `MPNr` int(10) DEFAULT NULL COMMENT 'Ordnungsziffer für Medikamentenplan'"
 Str(1, 108, 4) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT 'Zeitpunkt, der Speicherung im Turbomed'"
 Str(1, 108, 5) = " `Datum` datetime DEFAULT NULL COMMENT 'Zeitpunkt aus dem Kopf des Medikamentenplans'"
 Str(1, 108, 6) = " `Medikament` varchar(54) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 7) = " `MedAnfang` varchar(35) CHARACTER SET latin1 COLLATE latin1_general_ci DEFAULT NULL"
 Str(1, 108, 8) = " `FeldNr` smallint(6) DEFAULT NULL"
 Str(1, 108, 9) = " `mo` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 10) = " `mi` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 11) = " `nm` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 12) = " `ab` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 13) = " `zn` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 108, 14) = " `bBed` bit(1) DEFAULT NULL"
 Str(1, 108, 15) = " `Bemerkung` longtext COLLATE latin1_german2_ci"
 Str(1, 108, 16) = " `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 108, 17) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 108, 18) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 108, 19) = "  KEY `FälleMedPlan` (`FID`)"
 Str(1, 108, 20) = "  KEY `MedPlanMedikament` (`Medikament`)"
 Str(1, 108, 21) = "  KEY `NamenMedPlan` (`Pat_ID`)"
 Str(1, 108, 22) = "  KEY `MedArtenMedPlan_AccRel` (`MedAnfang`)"
 Str(1, 108, 23) = "  KEY `MPNr` (`MPNr`)"
 Str(1, 108, 24) = "  KEY `MPNrPat_ID` (`Pat_ID`,`MPNr`)"
 Str(1, 108, 25) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`MPNr`)"
 Str(1, 108, 26) = "  CONSTRAINT `FälleMedPlan_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 108, 27) = "  CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`)"
 Str(1, 108, 28) = "  CONSTRAINT `NamenMedPlan_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 108, 29) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr108

Sub FüllStr109()
 Str(0, 109, 0) = "namen"
 Str(0, 109, 1) = "`Pat_ID`"
 Str(0, 109, 2) = "`lfdnr`"
 Str(0, 109, 3) = "`NVorsatz`"
 Str(0, 109, 4) = "`Nachname`"
 Str(0, 109, 5) = "`Vorname`"
 Str(0, 109, 6) = "`GebDat`"
 Str(0, 109, 7) = "`Straße`"
 Str(0, 109, 8) = "`KVKStatus`"
 Str(0, 109, 9) = "`Geschlecht`"
 Str(0, 109, 10) = "`Plz`"
 Str(0, 109, 11) = "`Ort`"
 Str(0, 109, 12) = "`Weggeldzone`"
 Str(0, 109, 13) = "`WeggzZahl`"
 Str(0, 109, 14) = "`AufnDat`"
 Str(0, 109, 15) = "`LANR`"
 Str(0, 109, 16) = "`BStNr`"
 Str(0, 109, 17) = "`Titel`"
 Str(0, 109, 18) = "`Versichertennummer`"
 Str(0, 109, 19) = "`PrivatTel`"
 Str(0, 109, 20) = "`KVNr`"
 Str(0, 109, 21) = "`PrivatTel_2`"
 Str(0, 109, 22) = "`PrivatFax`"
 Str(0, 109, 23) = "`DienstTel`"
 Str(0, 109, 24) = "`PrivatMobil`"
 Str(0, 109, 25) = "`Email`"
 Str(0, 109, 26) = "`Arbeitgeber`"
 Str(0, 109, 27) = "`AnAllgda`"
 Str(0, 109, 28) = "`An1da`"
 Str(0, 109, 29) = "`An2da`"
 Str(0, 109, 30) = "`Checkda`"
 Str(0, 109, 31) = "`DMTypaD`"
 Str(0, 109, 32) = "`AktZeit`"
 Str(0, 109, 33) = "`absPos`"
 Str(0, 109, 34) = "`StByte`"
 Str(0, 109, 35) = "`Cave`"
 Str(0, 109, 36) = "`Notiz`"
 Str(0, 109, 37) = "`zubenach`"
 Str(0, 109, 38) = "`Verwandt`"
 Str(0, 109, 39) = "`Sprache`"
 Str(0, 109, 40) = "`lAktTM`"
 Str(0, 109, 41) = "`PAT_ID`"
 Str(0, 109, 42) = "`Auswahl`"
 Str(0, 109, 43) = "`HausärzteNamen_AccRel`"
 Str(0, 109, 44) = "`weggszahl`"
 Str(0, 109, 45) = "`weggeldzone`"
 Str(0, 109, 46) = "`HausärzteNamen_AccRel`"
 ArtZ(0, 109) = 40
 ArtZ(1, 109) = 5
 ArtZ(2, 109) = 1
 Str(1, 109, 0) = "CREATE TABLE `namen` ("
 Str(1, 109, 1) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 109, 2) = " `lfdnr` int(10) DEFAULT NULL COMMENT 'laufende Patientennummer'"
 Str(1, 109, 3) = " `NVorsatz` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3100'"
 Str(1, 109, 4) = " `Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3101'"
 Str(1, 109, 5) = " `Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3102'"
 Str(1, 109, 6) = " `GebDat` datetime DEFAULT NULL COMMENT '3103'"
 Str(1, 109, 7) = " `Straße` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3107'"
 Str(1, 109, 8) = " `KVKStatus` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3108'"
 Str(1, 109, 9) = " `Geschlecht` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3110'"
 Str(1, 109, 10) = " `Plz` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3112'"
 Str(1, 109, 11) = " `Ort` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3113'"
 Str(1, 109, 12) = " `Weggeldzone` varchar(2) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3631 (1) Weggeldzone mit Z'"
 Str(1, 109, 13) = " `WeggzZahl` decimal(1,0) DEFAULT NULL COMMENT '3631 (2) Weggeldzone, Zahl in Feld 2'"
 Str(1, 109, 14) = " `AufnDat` datetime DEFAULT NULL COMMENT '3610'"
 Str(1, 109, 15) = " `LANR` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP'"
 Str(1, 109, 16) = " `BStNr` varchar(9) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3536 Betriebsstättennummer'"
 Str(1, 109, 17) = " `Titel` varchar(39) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3104'"
 Str(1, 109, 18) = " `Versichertennummer` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3105'"
 Str(1, 109, 19) = " `PrivatTel` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 109, 20) = " `KVNr` varchar(7) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3630'"
 Str(1, 109, 21) = " `PrivatTel_2` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 109, 22) = " `PrivatFax` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 109, 23) = " `DienstTel` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 109, 24) = " `PrivatMobil` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3629'"
 Str(1, 109, 25) = " `Email` varchar(100) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Email'"
 Str(1, 109, 26) = " `Arbeitgeber` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3625'"
 Str(1, 109, 27) = " `AnAllgda` bit(1) DEFAULT NULL COMMENT 'Anamnese allgemein da'"
 Str(1, 109, 28) = " `An1da` bit(1) DEFAULT NULL COMMENT 'Anamnese S.1 da'"
 Str(1, 109, 29) = " `An2da` bit(1) DEFAULT NULL COMMENT 'Anamnese S.2 da'"
 Str(1, 109, 30) = " `Checkda` bit(1) DEFAULT NULL COMMENT 'Checkliste da'"
 Str(1, 109, 31) = " `DMTypaD` varchar(1) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'aus Diagnosen'"
 Str(1, 109, 32) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 109, 33) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei'"
 Str(1, 109, 34) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 109, 35) = " `Cave` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3654'"
 Str(1, 109, 36) = " `Notiz` varchar(150) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>'"
 Str(1, 109, 37) = " `zubenach` varchar(41) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3633'"
 Str(1, 109, 38) = " `Verwandt` varchar(102) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3632'"
 Str(1, 109, 39) = " `Sprache` varchar(45) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3628'"
 Str(1, 109, 40) = " `lAktTM` datetime NOT NULL COMMENT 'letzte Aktualisierung in Turbomed'"
 Str(1, 109, 41) = "  UNIQUE KEY `PAT_ID` (`Pat_ID`)"
 Str(1, 109, 42) = "  KEY `Auswahl` (`Nachname`,`Vorname`,`GebDat`)"
 Str(1, 109, 43) = "  KEY `HausärzteNamen_AccRel` (`KVNr`)"
 Str(1, 109, 44) = "  KEY `weggszahl` (`WeggzZahl`)"
 Str(1, 109, 45) = "  KEY `weggeldzone` (`Weggeldzone`)"
 Str(1, 109, 46) = "  CONSTRAINT `HausärzteNamen_AccRel` FOREIGN KEY (`KVNr`) REFERENCES `hausaerzte` (`KVNr`)"
 Str(1, 109, 47) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr109

Sub FüllStr110()
 Str(0, 110, 0) = "pauschalen"
 Str(0, 110, 1) = "`Leistung`"
 Str(0, 110, 2) = "`Betreuung`"
 ArtZ(0, 110) = 2
 Str(1, 110, 0) = "CREATE TABLE `pauschalen` ("
 Str(1, 110, 1) = " `Leistung` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 110, 2) = " `Betreuung` bit(1) DEFAULT NULL"
 Str(1, 110, 3) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr110

Sub FüllStr111()
 Str(0, 111, 0) = "pumpenträger"
 Str(0, 111, 1) = "`fall`"
 Str(0, 111, 2) = "`anrede`"
 Str(0, 111, 3) = "`nachname`"
 Str(0, 111, 4) = "`vorname`"
 Str(0, 111, 5) = "`gebdat`"
 Str(0, 111, 6) = "`privattel`"
 Str(0, 111, 7) = "`privattel_2`"
 Str(0, 111, 8) = "`privatfax`"
 Str(0, 111, 9) = "`diensttel`"
 Str(0, 111, 10) = "`straße`"
 Str(0, 111, 11) = "`plz`"
 Str(0, 111, 12) = "`ort`"
 Str(0, 111, 13) = "`mail1`"
 Str(0, 111, 14) = "`mail2`"
 Str(0, 111, 15) = "`email`"
 ArtZ(0, 111) = 15
 Str(1, 111, 0) = "CREATE TABLE `pumpenträger` ("
 Str(1, 111, 1) = " `fall` varchar(2) DEFAULT NULL"
 Str(1, 111, 2) = " `anrede` varchar(4) DEFAULT NULL"
 Str(1, 111, 3) = " `nachname` varchar(21) DEFAULT NULL"
 Str(1, 111, 4) = " `vorname` varchar(19) DEFAULT NULL"
 Str(1, 111, 5) = " `gebdat` datetime DEFAULT NULL"
 Str(1, 111, 6) = " `privattel` varchar(100) DEFAULT NULL"
 Str(1, 111, 7) = " `privattel_2` varchar(50) DEFAULT NULL"
 Str(1, 111, 8) = " `privatfax` varchar(50) DEFAULT NULL"
 Str(1, 111, 9) = " `diensttel` varchar(50) DEFAULT NULL"
 Str(1, 111, 10) = " `straße` varchar(50) DEFAULT NULL"
 Str(1, 111, 11) = " `plz` varchar(20) DEFAULT NULL"
 Str(1, 111, 12) = " `ort` varchar(70) DEFAULT NULL"
 Str(1, 111, 13) = " `mail1` varchar(100) DEFAULT NULL"
 Str(1, 111, 14) = " `mail2` varchar(100) DEFAULT NULL"
 Str(1, 111, 15) = " `email` varchar(100) DEFAULT NULL"
 Str(1, 111, 16) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr111

Sub FüllStr112()
 Str(0, 112, 0) = "queries"
 Str(0, 112, 1) = "`Name`"
 Str(0, 112, 2) = "`sql`"
 Str(0, 112, 3) = "`gespeichert`"
 Str(0, 112, 4) = "`Name`"
 Str(0, 112, 5) = "`NamGesp`"
 Str(0, 112, 6) = "`sql`"
 ArtZ(0, 112) = 3
 ArtZ(1, 112) = 3
 Str(1, 112, 0) = "CREATE TABLE `queries` ("
 Str(1, 112, 1) = " `Name` varchar(100) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Name der SQL-Abfrage'"
 Str(1, 112, 2) = " `sql` longtext COLLATE latin1_german2_ci COMMENT 'SQL-String, der geht'"
 Str(1, 112, 3) = " `gespeichert` datetime NOT NULL DEFAULT '0000-00-00 00:00:00' COMMENT 'Zeitpunkt, zu der er geht'"
 Str(1, 112, 4) = "  PRIMARY KEY (`Name`,`gespeichert`)"
 Str(1, 112, 5) = "  UNIQUE KEY `NamGesp` (`Name`,`gespeichert`)"
 Str(1, 112, 6) = "  KEY `sql` (`sql`(255))"
 Str(1, 112, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr112

Sub FüllStr113()
 Str(0, 113, 0) = "rel2"
 Str(0, 113, 1) = "`ccolumn`"
 Str(0, 113, 2) = "`grbit`"
 Str(0, 113, 3) = "`icolumn`"
 Str(0, 113, 4) = "`szColumn`"
 Str(0, 113, 5) = "`szObject`"
 Str(0, 113, 6) = "`szReferencedColumn`"
 Str(0, 113, 7) = "`szReferencedObject`"
 Str(0, 113, 8) = "`szRelationship`"
 Str(0, 113, 9) = "`szObject`"
 Str(0, 113, 10) = "`szReferencedObject`"
 Str(0, 113, 11) = "`szRelationship`"
 ArtZ(0, 113) = 8
 ArtZ(1, 113) = 3
 Str(1, 113, 0) = "CREATE TABLE `rel2` ("
 Str(1, 113, 1) = " `ccolumn` int(10) DEFAULT NULL"
 Str(1, 113, 2) = " `grbit` int(10) DEFAULT NULL"
 Str(1, 113, 3) = " `icolumn` int(10) DEFAULT NULL"
 Str(1, 113, 4) = " `szColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 113, 5) = " `szObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 113, 6) = " `szReferencedColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 113, 7) = " `szReferencedObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 113, 8) = " `szRelationship` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 113, 9) = "  KEY `szObject` (`szObject`)"
 Str(1, 113, 10) = "  KEY `szReferencedObject` (`szReferencedObject`)"
 Str(1, 113, 11) = "  KEY `szRelationship` (`szRelationship`)"
 Str(1, 113, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr113

Sub FüllStr114()
 Str(0, 114, 0) = "relationen"
 Str(0, 114, 1) = "`ccolumn`"
 Str(0, 114, 2) = "`grbit`"
 Str(0, 114, 3) = "`icolumn`"
 Str(0, 114, 4) = "`szColumn`"
 Str(0, 114, 5) = "`szObject`"
 Str(0, 114, 6) = "`szReferencedColumn`"
 Str(0, 114, 7) = "`szReferencedObject`"
 Str(0, 114, 8) = "`szRelationship`"
 Str(0, 114, 9) = "`szObject`"
 Str(0, 114, 10) = "`szReferencedObject`"
 Str(0, 114, 11) = "`szRelationship`"
 ArtZ(0, 114) = 8
 ArtZ(1, 114) = 3
 Str(1, 114, 0) = "CREATE TABLE `relationen` ("
 Str(1, 114, 1) = " `ccolumn` int(10) DEFAULT NULL"
 Str(1, 114, 2) = " `grbit` int(10) DEFAULT NULL"
 Str(1, 114, 3) = " `icolumn` int(10) DEFAULT NULL"
 Str(1, 114, 4) = " `szColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 114, 5) = " `szObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 114, 6) = " `szReferencedColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 114, 7) = " `szReferencedObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 114, 8) = " `szRelationship` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 114, 9) = "  KEY `szObject` (`szObject`)"
 Str(1, 114, 10) = "  KEY `szReferencedObject` (`szReferencedObject`)"
 Str(1, 114, 11) = "  KEY `szRelationship` (`szRelationship`)"
 Str(1, 114, 12) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr114

Sub FüllStr115()
 Str(0, 115, 0) = "rezepteintraege"
 Str(0, 115, 1) = "`FID`"
 Str(0, 115, 2) = "`Pat_ID`"
 Str(0, 115, 3) = "`ZeitPunkt`"
 Str(0, 115, 4) = "`Rezept`"
 Str(0, 115, 5) = "`Rezeptklasse`"
 Str(0, 115, 6) = "`Medikament`"
 Str(0, 115, 7) = "`PZN`"
 Str(0, 115, 8) = "`absPos`"
 Str(0, 115, 9) = "`AktZeit`"
 Str(0, 115, 10) = "`QS`"
 Str(0, 115, 11) = "`QT`"
 Str(0, 115, 12) = "`StByte`"
 Str(0, 115, 13) = "`Auswahl`"
 Str(0, 115, 14) = "`FälleRezeptEinträge`"
 Str(0, 115, 15) = "`FID`"
 Str(0, 115, 16) = "`NamenRezeptEinträge`"
 Str(0, 115, 17) = "`FälleRezeptEinträge_AccRel`"
 Str(0, 115, 18) = "`NamenRezeptEinträge_AccRel`"
 ArtZ(0, 115) = 12
 ArtZ(1, 115) = 4
 ArtZ(2, 115) = 2
 Str(1, 115, 0) = "CREATE TABLE `rezepteintraege` ("
 Str(1, 115, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 115, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 115, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 115, 4) = " `Rezept` varchar(3) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6210, 3652(1), 6218(1)'"
 Str(1, 115, 5) = " `Rezeptklasse` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)'"
 Str(1, 115, 6) = " `Medikament` varchar(70) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '3652(2), 6218(4)'"
 Str(1, 115, 7) = " `PZN` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL COMMENT '6210(2), 6218(3)'"
 Str(1, 115, 8) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 115, 9) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 115, 10) = " `QS` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns sortiert'"
 Str(1, 115, 11) = " `QT` varchar(5) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Quartal des Behandlungsfallbeginns'"
 Str(1, 115, 12) = " `StByte` int(11) DEFAULT NULL COMMENT 'Statusbyte'"
 Str(1, 115, 13) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`Rezept`,`Medikament`)"
 Str(1, 115, 14) = "  KEY `FälleRezeptEinträge` (`FID`)"
 Str(1, 115, 15) = "  KEY `FID` (`FID`)"
 Str(1, 115, 16) = "  KEY `NamenRezeptEinträge` (`Pat_ID`)"
 Str(1, 115, 17) = "  CONSTRAINT `FälleRezeptEinträge_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 115, 18) = "  CONSTRAINT `NamenRezeptEinträge_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 115, 19) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr115

Sub FüllStr116()
 Str(0, 116, 0) = "rr"
 Str(0, 116, 1) = "`FID`"
 Str(0, 116, 2) = "`Pat_ID`"
 Str(0, 116, 3) = "`ZeitPunkt`"
 Str(0, 116, 4) = "`RR`"
 Str(0, 116, 5) = "`absPos`"
 Str(0, 116, 6) = "`AktZeit`"
 Str(0, 116, 7) = "`StByte`"
 Str(0, 116, 8) = "`Auswahl`"
 Str(0, 116, 9) = "`FälleRR`"
 Str(0, 116, 10) = "`FID`"
 Str(0, 116, 11) = "`NamenRR`"
 Str(0, 116, 12) = "`FälleRR_AccRel`"
 Str(0, 116, 13) = "`NamenRR_AccRel`"
 ArtZ(0, 116) = 7
 ArtZ(1, 116) = 4
 ArtZ(2, 116) = 2
 Str(1, 116, 0) = "CREATE TABLE `rr` ("
 Str(1, 116, 1) = " `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug'"
 Str(1, 116, 2) = " `Pat_ID` int(10) DEFAULT NULL COMMENT '3000'"
 Str(1, 116, 3) = " `ZeitPunkt` datetime DEFAULT NULL COMMENT '6200 + 6201'"
 Str(1, 116, 4) = " `RR` longtext COLLATE latin1_german2_ci COMMENT '6230'"
 Str(1, 116, 5) = " `absPos` int(10) DEFAULT NULL COMMENT 'Zeile in BDT-Datei'"
 Str(1, 116, 6) = " `AktZeit` datetime DEFAULT NULL COMMENT 'Aktualisierungszeit'"
 Str(1, 116, 7) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung'"
 Str(1, 116, 8) = "  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`RR`(255))"
 Str(1, 116, 9) = "  KEY `FälleRR` (`FID`)"
 Str(1, 116, 10) = "  KEY `FID` (`FID`)"
 Str(1, 116, 11) = "  KEY `NamenRR` (`Pat_ID`)"
 Str(1, 116, 12) = "  CONSTRAINT `FälleRR_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 116, 13) = "  CONSTRAINT `NamenRR_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`) ON DELETE CASCADE ON UPDATE CASCADE"
 Str(1, 116, 14) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr116

Sub FüllStr117()
 Str(0, 117, 0) = "rrparse"
 Str(0, 117, 1) = "`Pat_id`"
 Str(0, 117, 2) = "`Zeitpunkt`"
 Str(0, 117, 3) = "`RRSyst`"
 Str(0, 117, 4) = "`RRDiast`"
 Str(0, 117, 5) = "`Quelle`"
 Str(0, 117, 6) = "`ID`"
 ArtZ(0, 117) = 5
 ArtZ(1, 117) = 1
 Str(1, 117, 0) = "CREATE TABLE `rrparse` ("
 Str(1, 117, 1) = " `Pat_id` int(10) DEFAULT NULL"
 Str(1, 117, 2) = " `Zeitpunkt` datetime DEFAULT NULL"
 Str(1, 117, 3) = " `RRSyst` smallint(6) DEFAULT NULL"
 Str(1, 117, 4) = " `RRDiast` smallint(6) DEFAULT NULL"
 Str(1, 117, 5) = " `Quelle` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 117, 6) = "  KEY `ID` (`Pat_id`,`Zeitpunkt`,`RRSyst`,`RRDiast`)"
 Str(1, 117, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr117

Sub FüllStr118()
 Str(0, 118, 0) = "schulleistzahl"
 Str(0, 118, 1) = "`quartal`"
 Str(0, 118, 2) = "`lzahl`"
 ArtZ(0, 118) = 2
 Str(1, 118, 0) = "CREATE TABLE `schulleistzahl` ("
 Str(1, 118, 1) = " `quartal` varchar(5) DEFAULT NULL"
 Str(1, 118, 2) = " `lzahl` bigint(21) DEFAULT NULL"
 Str(1, 118, 3) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr118

Sub FüllStr119()
 Str(0, 119, 0) = "schulzahl"
 Str(0, 119, 1) = "`quartal`"
 Str(0, 119, 2) = "`lzahl`"
 Str(0, 119, 3) = "`ezahl`"
 ArtZ(0, 119) = 3
 Str(1, 119, 0) = "CREATE TABLE `schulzahl` ("
 Str(1, 119, 1) = " `quartal` varchar(5) DEFAULT NULL"
 Str(1, 119, 2) = " `lzahl` bigint(21) DEFAULT NULL"
 Str(1, 119, 3) = " `ezahl` bigint(21) DEFAULT NULL"
 Str(1, 119, 4) = " ENGINE=MyISAM DEFAULT CHARSET=utf8"
End Sub ' FüllStr119

Sub FüllStr120()
 Str(0, 120, 0) = "tmpfif"
 Str(0, 120, 1) = "`FeldVW`"
 Str(0, 120, 2) = "`Feld`"
 Str(0, 120, 3) = "`StByte`"
 Str(0, 120, 4) = "`FeldVW`"
 Str(0, 120, 5) = "`Feld`"
 ArtZ(0, 120) = 3
 ArtZ(1, 120) = 2
 Str(1, 120, 0) = "CREATE TABLE `tmpfif` ("
 Str(1, 120, 1) = " `FeldVW` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 120, 2) = " `Feld` longtext COLLATE latin1_german2_ci"
 Str(1, 120, 3) = " `StByte` int(10) DEFAULT NULL COMMENT 'Ordinalziffer der Einlesung'"
 Str(1, 120, 4) = "  PRIMARY KEY (`FeldVW`)"
 Str(1, 120, 5) = "  KEY `Feld` (`Feld`(255))"
 Str(1, 120, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr120

Sub FüllStr121()
 Str(0, 121, 0) = "tmprelationen"
 Str(0, 121, 1) = "`ccolumn`"
 Str(0, 121, 2) = "`grbit`"
 Str(0, 121, 3) = "`icolumn`"
 Str(0, 121, 4) = "`szColumn`"
 Str(0, 121, 5) = "`szObject`"
 Str(0, 121, 6) = "`szReferencedColumn`"
 Str(0, 121, 7) = "`szReferencedObject`"
 Str(0, 121, 8) = "`szRelationship`"
 ArtZ(0, 121) = 8
 Str(1, 121, 0) = "CREATE TABLE `tmprelationen` ("
 Str(1, 121, 1) = " `ccolumn` int(10) DEFAULT NULL"
 Str(1, 121, 2) = " `grbit` int(10) DEFAULT NULL"
 Str(1, 121, 3) = " `icolumn` int(10) DEFAULT NULL"
 Str(1, 121, 4) = " `szColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 121, 5) = " `szObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 121, 6) = " `szReferencedColumn` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 121, 7) = " `szReferencedObject` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 121, 8) = " `szRelationship` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 121, 9) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr121

Sub FüllStr122()
 Str(0, 122, 0) = "ueberwvon"
 Str(0, 122, 1) = "`ID`"
 Str(0, 122, 2) = "`KVNr`"
 Str(0, 122, 3) = "`Titel`"
 Str(0, 122, 4) = "`Vorname`"
 Str(0, 122, 5) = "`Zusatz`"
 Str(0, 122, 6) = "`Nachname`"
 Str(0, 122, 7) = "`ID`"
 Str(0, 122, 8) = "`KVNr`"
 Str(0, 122, 9) = "`Name`"
 ArtZ(0, 122) = 6
 ArtZ(1, 122) = 3
 Str(1, 122, 0) = "CREATE TABLE `ueberwvon` ("
 Str(1, 122, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 122, 2) = " `KVNr` varchar(7) DEFAULT NULL"
 Str(1, 122, 3) = " `Titel` varchar(30) DEFAULT NULL"
 Str(1, 122, 4) = " `Vorname` varchar(45) DEFAULT NULL"
 Str(1, 122, 5) = " `Zusatz` varchar(30) DEFAULT NULL"
 Str(1, 122, 6) = " `Nachname` varchar(45) DEFAULT NULL"
 Str(1, 122, 7) = "  PRIMARY KEY (`ID`) USING BTREE"
 Str(1, 122, 8) = "  KEY `KVNr` (`KVNr`)"
 Str(1, 122, 9) = "  KEY `Name` (`Nachname`,`Vorname`)"
 Str(1, 122, 10) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='4247 Überwiesen von'"
End Sub ' FüllStr122

Sub FüllStr123()
 Str(0, 123, 0) = "unbekannte kennungen"
 Str(0, 123, 1) = "`Kennung`"
 Str(0, 123, 2) = "`absPos`"
 Str(0, 123, 3) = "`StByte`"
 Str(0, 123, 4) = "`Pat_id`"
 Str(0, 123, 5) = "`Inhalt`"
 Str(0, 123, 6) = "`Kennung`"
 ArtZ(0, 123) = 5
 ArtZ(1, 123) = 1
 Str(1, 123, 0) = "CREATE TABLE `unbekannte kennungen` ("
 Str(1, 123, 1) = " `Kennung` varchar(4) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 123, 2) = " `absPos` int(10) DEFAULT NULL"
 Str(1, 123, 3) = " `StByte` int(10) DEFAULT NULL"
 Str(1, 123, 4) = " `Pat_id` int(10) DEFAULT NULL COMMENT 'zugehöriger Patient für spätere Ermittlungen'"
 Str(1, 123, 5) = " `Inhalt` varchar(200) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Inhalt Zeile zum Wiederauffinden'"
 Str(1, 123, 6) = "  UNIQUE KEY `Kennung` (`Kennung`)"
 Str(1, 123, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr123

Sub FüllStr124()
 Str(0, 124, 0) = "vlassen"
 Str(0, 124, 1) = "`Datum`"
 Str(0, 124, 2) = "`Name`"
 Str(0, 124, 3) = "`Größe`"
 Str(0, 124, 4) = "`pfad`"
 Str(0, 124, 5) = "`Datum`"
 Str(0, 124, 6) = "`Name`"
 ArtZ(0, 124) = 4
 ArtZ(1, 124) = 2
 Str(1, 124, 0) = "CREATE TABLE `vlassen` ("
 Str(1, 124, 1) = " `Datum` datetime DEFAULT NULL"
 Str(1, 124, 2) = " `Name` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 124, 3) = " `Größe` int(10) DEFAULT NULL"
 Str(1, 124, 4) = " `pfad` varchar(255) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 124, 5) = "  KEY `Datum` (`Datum`)"
 Str(1, 124, 6) = "  KEY `Name` (`Name`)"
 Str(1, 124, 7) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr124

Sub FüllStr125()
 Str(0, 125, 0) = "werte_scheingruppen"
 Str(0, 125, 1) = "`schgr`"
 Str(0, 125, 2) = "`Erklärung`"
 Str(0, 125, 3) = "`schgr`"
 ArtZ(0, 125) = 2
 ArtZ(1, 125) = 1
 Str(1, 125, 0) = "CREATE TABLE `werte_scheingruppen` ("
 Str(1, 125, 1) = " `schgr` decimal(2,0) NOT NULL DEFAULT '0' COMMENT '4239'"
 Str(1, 125, 2) = " `Erklärung` varchar(70) CHARACTER SET latin1 DEFAULT NULL"
 Str(1, 125, 3) = "  PRIMARY KEY (`schgr`) USING BTREE"
 Str(1, 125, 4) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Scheingruppen in Turbomed'"
End Sub ' FüllStr125

Sub FüllStr126()
 Str(0, 126, 0) = "werte_weggeldzonen"
 Str(0, 126, 1) = "`Weggeldzone`"
 Str(0, 126, 2) = "`Zonennr`"
 Str(0, 126, 3) = "`Bereich`"
 Str(0, 126, 4) = "`Weggeldzone`"
 Str(0, 126, 5) = "`Zonennr`"
 ArtZ(0, 126) = 3
 ArtZ(1, 126) = 2
 Str(1, 126, 0) = "CREATE TABLE `werte_weggeldzonen` ("
 Str(1, 126, 1) = " `Weggeldzone` varchar(2) CHARACTER SET latin1 NOT NULL DEFAULT '' COMMENT '3631 (1) Weggeldzone'"
 Str(1, 126, 2) = " `Zonennr` decimal(1,0) DEFAULT NULL COMMENT '3631 (2) Weggeldzonenziffer'"
 Str(1, 126, 3) = " `Bereich` varchar(30) CHARACTER SET latin1 NOT NULL DEFAULT '' COMMENT 'Kilometer-Bereich'"
 Str(1, 126, 4) = "  PRIMARY KEY (`Weggeldzone`) USING BTREE"
 Str(1, 126, 5) = "  KEY `Zonennr` (`Zonennr`)"
 Str(1, 126, 6) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Weggeldzonen'"
End Sub ' FüllStr126

Function doEx&(sql$, obtolerant%) ' SQL-Befehl ausführen, Fehler anzeigen
 Dim rAF&, FMeld$
 If obtolerant Then On Error Resume Next Else On Error GoTo fehler
 call cnz.Execute(sql, rAf)
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
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
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
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SplitN/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' aufSplit

Public Function doMach_quelle1(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
 Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$
 Dim i&, p1&, p2&, p3&, CLen&, CLen1&, obLT%
 Dim Index$()
 On Error Resume Next
 hDBn = DBn
 Open App.Path & "\MachDB.bas_prot.txt" For Output As #302
 obProt = (Err.Number = 0)
 On Error GoTo fehler
 If LenB(server) = 0 Then Server = GetServer(DbCn)
 cnzCStr = "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=" & Server & ";uid=mysql;pwd=97a5o6;"
 set cnz = Nothing
 cnz.open cnzCStr
 call doex("create database if not exists `" & DBN & "` character set latin1 collate latin1_german2_ci;",0)
 call doex("grant all privileges on `" & DBN & "`.* to 'praxis'@'%' identified by 'sonne' with grant option",0)
 call doex("grant all privileges on `" & DBN & "`.* to 'praxis'@'localhost' identified by 'sonne' with grant option",0)
 call doex("use `" & DBN & "`",0)
 call doex("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ",0)
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
 FüllStr114
 FüllStr115
 FüllStr116
 FüllStr117
 FüllStr118
 FüllStr119
 FüllStr120
 FüllStr121
 FüllStr122
 FüllStr123
 FüllStr124
 FüllStr125
 FüllStr126
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 126
  If Instr(Str(1, i, 0),"CREATE TABLE")<>0 then
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
    set rsc = nothing
    rsc.Open "show create table `" & tbl & "`", cnz, adOpenStatic, adLockReadOnly
    sct = rsc.Fields(1)
    If InStrB(sct, "CREATE ALGORITHM") = 1 Then
     FNr = doEx("drop view `" & Tbl & "`", 0)
     FNr = doEx(sql.Value, 0)
    Else
     Exit Do
    End If
   Loop
   If InStrB(AIoZ(sct), AIoZ(Str(1, i, ZZ))) = 0 Then
    Call doEx("alter table `" & tbl & "`" & Str(1, i, ZZ), 0)
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
      posi = " FIRST,"
     Else
      posi = " AFTER " & Str(0, i, j - 1) & ","
     End If
     If Not enthalten Then
      TMt.AppVar (Array(" add ", Str(1, i, j), posi))
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
       TMt.AppVar (Array(" modify ", Str(1, i, j), posi))
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
    Call doEx("Alter Table `" & tbl & "` " & TMt.Value, -1)
   End If
  End If ' InStr(Str(1, i, 0), "CREATE TABLE") <> 0 Then
 Next i
 For i = 0 To 126
  If Instr(Str(1, i, 0),"CREATE TABLE")<>0 then
   Tbl = Str(0, i, 0)
   ZZ = ArtZ(0, i) + ArtZ(1, i)
   set rsc = nothing
   rsc.Open "show create table `" & tbl & "`", cnz, adOpenStatic, adLockReadOnly
   sct = rsc.Fields(1)
   ZZ = ZZ + ArtZ(2, i) + 1
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
 For runde = 0 to 4
  For i = 0 To 126
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
 call doex("set FOREIGN_KEY_CHECKS = 1",0)
 If obProt then Close #302
 If not obstumm then
  MsgBox "Fertig mit doMach_quelle1(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_quelle1/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_quelle1

Function GetServer$(DBCn As ADODB.Connection)
Dim spos&, sp2&
spos = InStr(LCase$(DBCn), "server=")
If spos <> 0 Then
 sp2 = InStr(spos, DBCn, ";")
 If sp2 = 0 Then sp2 = Len(DBCn)
 GetServer = Mid$(DBCn, spos + 7, sp2 - spos - 7)
End If
End Function ' GetServer

Function AIoZ(Ursp) As CString ' Ursp kann $ oder CString sein
 Const Such$ = "AUTO_INCREMENT="
 Set AIoZ = New CString
 AIoZ = Ursp
 Dim p0&, p1&
 p0 = AIoZ.Instr(Such)
 If p0 <> 0 Then
  p1 = AIoZ.Instr(" ", p0)
  AIoZ.Cut (p0 - 2)
  AIoZ.Append Mid(Ursp, p1)
 End If
End Function ' AIoZ(Ursp$) As CString
