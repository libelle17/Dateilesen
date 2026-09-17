'Bauanleitung für eine Datenbank wie `//linux/mails` vom 15.11.09 14:26:24
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 19, 71) As new CString, ArtZ&(3, 19)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "address"
 Str(0, 0, 1) = "`iid`"
 Str(0, 0, 2) = "`Address`"
 Str(0, 0, 3) = "`Class`"
 Str(0, 0, 4) = "`DisplayType`"
 Str(0, 0, 5) = "`ID`"
 Str(0, 0, 6) = "`Name`"
 Str(0, 0, 7) = "`Type`"
 Str(0, 0, 8) = "`Cpt`"
 Str(0, 0, 9) = "`MailEntryID`"
 Str(0, 0, 10) = "`iid`"
 Str(0, 0, 11) = "`ID`"
 ArtZ(0, 0) = 9
 ArtZ(1, 0) = 2
 Str(1, 0, 0) = "CREATE TABLE `address` ("
 Str(1, 0, 1) = " `iid` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 0, 2) = " `Address` varchar(4781) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 0, 3) = " `Class` int(10) NOT NULL DEFAULT '0'"
 Str(1, 0, 4) = " `DisplayType` int(10) NOT NULL DEFAULT '0'"
 Str(1, 0, 5) = " `ID` text COLLATE latin1_german2_ci"
 Str(1, 0, 6) = " `Name` varchar(3103) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 0, 7) = " `Type` varchar(4) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 0, 8) = " `Cpt` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Computername'"
 Str(1, 0, 9) = " `MailEntryID` varchar(51) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 0, 10) = "  PRIMARY KEY (`iid`)"
 Str(1, 0, 11) = "  UNIQUE KEY `ID` (`ID`(255),`Cpt`)"
 Str(1, 0, 12) = " ENGINE=InnoDB AUTO_INCREMENT=147390 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "attachments"
 Str(0, 1, 1) = "`Class`"
 Str(0, 1, 2) = "`DisplayName`"
 Str(0, 1, 3) = "`FileName`"
 Str(0, 1, 4) = "`Index`"
 Str(0, 1, 5) = "`PathName`"
 Str(0, 1, 6) = "`Position`"
 Str(0, 1, 7) = "`Type`"
 Str(0, 1, 8) = "`ParentEntryID`"
 Str(0, 1, 9) = "`Cpt`"
 Str(0, 1, 10) = "`Zuordnung`"
 ArtZ(0, 1) = 9
 ArtZ(1, 1) = 1
 Str(1, 1, 0) = "CREATE TABLE `attachments` ("
 Str(1, 1, 1) = " `Class` int(10) NOT NULL DEFAULT '0'"
 Str(1, 1, 2) = " `DisplayName` varchar(255) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 1, 3) = " `FileName` varchar(255) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 1, 4) = " `Index` int(10) NOT NULL DEFAULT '0'"
 Str(1, 1, 5) = " `PathName` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 1, 6) = " `Position` int(10) NOT NULL DEFAULT '0'"
 Str(1, 1, 7) = " `Type` int(10) NOT NULL DEFAULT '0'"
 Str(1, 1, 8) = " `ParentEntryID` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 1, 9) = " `Cpt` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Computername'"
 Str(1, 1, 10) = "  UNIQUE KEY `Zuordnung` (`ParentEntryID`,`Index`,`Cpt`)"
 Str(1, 1, 11) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "mailbcc"
 Str(0, 2, 1) = "`id`"
 Str(0, 2, 2) = "`BCC`"
 Str(0, 2, 3) = "`id`"
 Str(0, 2, 4) = "`bcc`"
 ArtZ(0, 2) = 2
 ArtZ(1, 2) = 2
 Str(1, 2, 0) = "CREATE TABLE `mailbcc` ("
 Str(1, 2, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 2, 2) = " `BCC` varchar(1548) COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 2, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 2, 4) = "  KEY `bcc` (`BCC`(100))"
 Str(1, 2, 5) = " ENGINE=MyISAM AUTO_INCREMENT=59 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='BCC von mails'"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "mailbody"
 Str(0, 3, 1) = "`id`"
 Str(0, 3, 2) = "`Body`"
 Str(0, 3, 3) = "`id`"
 Str(0, 3, 4) = "`body`"
 ArtZ(0, 3) = 2
 ArtZ(1, 3) = 2
 Str(1, 3, 0) = "CREATE TABLE `mailbody` ("
 Str(1, 3, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 3, 2) = " `Body` mediumtext COLLATE latin1_german2_ci NOT NULL"
 Str(1, 3, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 3, 4) = "  KEY `body` (`Body`(100))"
 Str(1, 3, 5) = " ENGINE=MyISAM AUTO_INCREMENT=60843 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Body von mails'"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "mailcategories"
 Str(0, 4, 1) = "`id`"
 Str(0, 4, 2) = "`Categories`"
 Str(0, 4, 3) = "`id`"
 Str(0, 4, 4) = "`Categories`"
 ArtZ(0, 4) = 2
 ArtZ(1, 4) = 2
 Str(1, 4, 0) = "CREATE TABLE `mailcategories` ("
 Str(1, 4, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 4, 2) = " `Categories` varchar(165) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 4, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 4, 4) = "  KEY `Categories` (`Categories`(20))"
 Str(1, 4, 5) = " ENGINE=MyISAM AUTO_INCREMENT=16 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Categories von mails'"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "mailcc"
 Str(0, 5, 1) = "`id`"
 Str(0, 5, 2) = "`CC`"
 Str(0, 5, 3) = "`id`"
 Str(0, 5, 4) = "`cc`"
 ArtZ(0, 5) = 2
 ArtZ(1, 5) = 2
 Str(1, 5, 0) = "CREATE TABLE `mailcc` ("
 Str(1, 5, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 5, 2) = " `CC` text COLLATE latin1_german2_ci NOT NULL"
 Str(1, 5, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 5, 4) = "  KEY `cc` (`CC`(100))"
 Str(1, 5, 5) = " ENGINE=MyISAM AUTO_INCREMENT=1525 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='CC von mails'"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "mailconversationtopic"
 Str(0, 6, 1) = "`id`"
 Str(0, 6, 2) = "`conversationtopic`"
 Str(0, 6, 3) = "`id`"
 Str(0, 6, 4) = "`conversationtopic`"
 ArtZ(0, 6) = 2
 ArtZ(1, 6) = 2
 Str(1, 6, 0) = "CREATE TABLE `mailconversationtopic` ("
 Str(1, 6, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 6, 2) = " `conversationtopic` varchar(497) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 6, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 6, 4) = "  KEY `conversationtopic` (`conversationtopic`(20))"
 Str(1, 6, 5) = " ENGINE=MyISAM AUTO_INCREMENT=64989 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='ConversationTopic von mails'"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "mailhtmlbody"
 Str(0, 7, 1) = "`id`"
 Str(0, 7, 2) = "`HTMLbody`"
 Str(0, 7, 3) = "`id`"
 Str(0, 7, 4) = "`HTMLbody`"
 ArtZ(0, 7) = 2
 ArtZ(1, 7) = 2
 Str(1, 7, 0) = "CREATE TABLE `mailhtmlbody` ("
 Str(1, 7, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 7, 2) = " `HTMLbody` mediumtext COLLATE latin1_german2_ci NOT NULL"
 Str(1, 7, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 7, 4) = "  KEY `HTMLbody` (`HTMLbody`(20))"
 Str(1, 7, 5) = " ENGINE=MyISAM AUTO_INCREMENT=49136 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='HTMLbody von mails'"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "mailreplyrecipientnames"
 Str(0, 8, 1) = "`id`"
 Str(0, 8, 2) = "`replyrecipientnames`"
 Str(0, 8, 3) = "`id`"
 Str(0, 8, 4) = "`replyrecipientnames`"
 ArtZ(0, 8) = 2
 ArtZ(1, 8) = 2
 Str(1, 8, 0) = "CREATE TABLE `mailreplyrecipientnames` ("
 Str(1, 8, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 8, 2) = " `replyrecipientnames` varchar(246) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 8, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 8, 4) = "  KEY `replyrecipientnames` (`replyrecipientnames`(20))"
 Str(1, 8, 5) = " ENGINE=MyISAM AUTO_INCREMENT=7801 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='ReplyRecipientNames von mails'"
End Sub ' FüllStr8

Sub FüllStr9()
 Str(0, 9, 0) = "mails"
 Str(1, 9, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `mails` AS select `mailsroh`.`ReceivedTime` AS `ReceivedTime`,`mailsendername`.`SenderName` AS `SenderName`,`mailsubject`.`subject` AS `Subject`,`mailbody`.`Body` AS `Body`,`mailsentonbehalfofname`.`SentOnBehalfOfName` AS `SentOnBehalfOfName`,`mailreplyrecipientnames`.`replyrecipientnames` AS `ReplyRecipientNames`,`mailcategories`.`Categories` AS `Categories`,`mailconversationtopic`.`conversationtopic` AS `ConversationTopic`,`mailbcc`.`BCC` AS `BCC`,`mailcc`.`CC` AS `CC`,`mailto`.`to` AS `To`,`mailhtmlbody`.`HTMLbody` AS `HTMLBody`,`mailsroh`.`SentOn` AS `SentOn`,`mailsroh`.`ATZahl` AS `ATZahl`,`mailsroh`.`LastModificationTime` AS `LastModificationTime`,`mailsroh`.`OrdnerEntryID` AS `OrdnerEntryID`,`mailsroh`.`MID` AS `MID`,`mailsroh`.`AlternateRecipientAllowed` AS `AlternateRecipientAllowed`,`mail" & _ 
  "sroh`.`AutoForwarded` AS `AutoForwarded`,`mailsroh`.`Class` AS `Class`,`mailsroh`.`CreationTime` AS `CreationTime`,`mailsroh`.`DeferredDeliveryTime` AS `DeferredDeliveryTime`,`mailsroh`.`DeleteAfterSubmit` AS `DeleteAfterSubmit`,`mailsroh`.`EntryID` AS `EntryID`,`mailsroh`.`ExpiryTime` AS `ExpiryTime`,`mailsroh`.`FlagDueBy` AS `FlagDueBy`,`mailsroh`.`FlagRequest` AS `FlagRequest`,`mailsroh`.`Importance` AS `Importance`,`mailsroh`.`MessageClass` AS `MessageClass`,`mailsroh`.`NoAging` AS `NoAging`,`mailsroh`.`ReceivedByEntryID` AS `ReceivedByEntryID`,`mailsroh`.`ReceivedByName` AS `ReceivedByName`,`mailsroh`.`Saved` AS `Saved`,`mailsroh`.`Sensitivity` AS `Sensitivity`,`mailsroh`.`Sent` AS `Sent`,`mailsroh`.`Size` AS `Size`,`mailsroh`.`Submitted` AS `Submitted`,`mailsroh`.`Unread` AS `Unread`,`mailsroh`.`Cpt` AS `Cpt`,`mailsroh`.`Application` AS `Application`,`mailsroh`.`BillingInformation`" & _ 
  " AS `BillingInformation`,`mailsroh`.`Companies` AS `Companies`,`mailsroh`.`ConversationIndex` AS `ConversationIndex`,`mailsroh`.`FlagStatus` AS `FlagStatus`,`mailsroh`.`FormDescription` AS `FormDescription`,`mailsroh`.`Mileage` AS `Mileage`,`mailsroh`.`OriginatorDeliveryReportRequested` AS `OriginatorDeliveryReportRequested`,`mailsroh`.`OutlookInternalVersion` AS `OutlookInternalVersion`,`mailsroh`.`OutlookVersion` AS `OutlookVersion`,`mailsroh`.`Parent` AS `Parent`,`mailsroh`.`ReadReceiptRequested` AS `ReadReceiptRequested`,`mailsroh`.`ReceivedOnBehalfOfEntryID` AS `ReceivedOnBehalfOfEntryID`,`mailsroh`.`ReceivedOnBehalfOfName` AS `ReceivedOnBehalfOfName`,`mailsroh`.`RecipientReassignmentProhibited` AS `RecipientReassignmentProhibited`,`mailsroh`.`ReminderOverrideDefault` AS `ReminderOverrideDefault`,`mailsroh`.`ReminderPlaySound` AS `ReminderPlaySound`,`mailsroh`.`ReminderSet` AS `Remi" & _ 
  "nderSet`,`mailsroh`.`ReminderSoundFile` AS `ReminderSoundFile`,`mailsroh`.`ReminderTime` AS `ReminderTime`,`mailsroh`.`RemoteStatus` AS `RemoteStatus`,`mailsroh`.`SaveSentMessageFolder` AS `SaveSentMessageFolder`,`mailsroh`.`Session` AS `Session`,`mailsroh`.`VotingOptions` AS `VotingOptions`,`mailsroh`.`VotingResponse` AS `VotingResponse` from (((((((((((`mailsroh` left join `mailsendername` on((`mailsroh`.`idSenderName` = `mailsendername`.`id`))) left join `mailsubject` on((`mailsroh`.`idSubject` = `mailsubject`.`id`))) left join `mailbody` on((`mailsroh`.`idBody` = `mailbody`.`id`))) left join `mailsentonbehalfofname` on((`mailsroh`.`idSentOnBehalfOfName` = `mailsentonbehalfofname`.`id`))) left join `mailreplyrecipientnames` on((`mailsroh`.`idReplyRecipientNames` = `mailreplyrecipientnames`.`id`))) left join `mailcategories` on((`mailsroh`.`idCategories` = `mailcategories`.`id`))) left" & _ 
  " join `mailconversationtopic` on((`mailsroh`.`idConversationTopic` = `mailconversationtopic`.`id`))) left join `mailbcc` on((`mailsroh`.`idBCC` = `mailbcc`.`id`))) left join `mailcc` on((`mailsroh`.`idCC` = `mailcc`.`id`))) left join `mailto` on((`mailsroh`.`idTo` = `mailto`.`id`))) left join `mailhtmlbody` on((`mailsroh`.`idHTMLBody` = `mailhtmlbody`.`id`)))"
End Sub ' FüllStr9

Sub FüllStr10()
 Str(0, 10, 0) = "mails_abfrage"
 Str(1, 10, 0) = "CREATE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `mails_abfrage` AS select `mails`.`ReceivedTime` AS `ReceivedTime`,`mails`.`SenderName` AS `SenderName`,`mails`.`ATZahl` AS `ATZahl`,`mails`.`Subject` AS `Subject`,`mails`.`LastModificationTime` AS `LastModificationTime`,`mails`.`AlternateRecipientAllowed` AS `AlternateRecipientAllowed`,`mails`.`AutoForwarded` AS `AutoForwarded`,`mails`.`BCC` AS `BCC`,left(`mails`.`Body`,100) AS `Body`,`mails`.`Categories` AS `Categories`,left(`mails`.`CC`,100) AS `CC`,`mails`.`Class` AS `Class`,`mails`.`CreationTime` AS `CreationTime`,`mails`.`DeferredDeliveryTime` AS `DeferredDeliveryTime`,`mails`.`DeleteAfterSubmit` AS `DeleteAfterSubmit`,`mails`.`EntryID` AS `EntryId`,`mails`.`ExpiryTime` AS `ExpiryTime`,`mails`.`FlagDueBy` AS `FlagDueBy`,`mails`.`FlagRequest` AS `FlagRequest`,left(`mails`.`HTMLBody`,100) AS `" & _ 
  "HTMLbody`,`mails`.`Importance` AS `Importance`,`mails`.`MessageClass` AS `MessageClass`,`mails`.`NoAging` AS `NoAging`,`mails`.`ReceivedByEntryID` AS `ReceivedByEntryID`,`mails`.`ReceivedByName` AS `ReceivedByName`,`mails`.`Saved` AS `Saved`,`mails`.`Sensitivity` AS `Sensitivity`,`mails`.`Sent` AS `Sent`,`mails`.`Size` AS `Size`,`mails`.`Submitted` AS `Submitted`,`mails`.`To` AS `To`,`mails`.`Unread` AS `Unread`,`mails`.`OrdnerEntryID` AS `OrdnerEntryId`,`mails`.`Application` AS `Application`,`mails`.`BillingInformation` AS `BillingInformation`,`mails`.`Companies` AS `Companies`,`mails`.`ConversationIndex` AS `ConvInd`,left(`mails`.`ConversationTopic`,100) AS `ConversationTopic`,`mails`.`FlagStatus` AS `FlagStatus`,`mails`.`FormDescription` AS `FormDescription`,`mails`.`Mileage` AS `Mileage`,`mails`.`OriginatorDeliveryReportRequested` AS `ODRR`,`mails`.`OutlookInternalVersion` AS `OIV`,`" & _ 
  "mails`.`OutlookVersion` AS `OutlookVersion`,`mails`.`Parent` AS `Parent`,`mails`.`ReadReceiptRequested` AS `ReadReceiptRequested`,`mails`.`ReceivedOnBehalfOfEntryID` AS `ReceivedOnBehalfofEntryid`,`mails`.`ReceivedOnBehalfOfName` AS `ReceivedOnBehalfofName`,`mails`.`RecipientReassignmentProhibited` AS `RecipientReassignmentProhibited`,`mails`.`ReminderOverrideDefault` AS `ReminderOverrideDefault`,`mails`.`ReminderPlaySound` AS `ReminderPlaySound`,`mails`.`ReminderSet` AS `ReminderSet`,`mails`.`ReminderSoundFile` AS `ReminderSoundFile`,`mails`.`ReminderTime` AS `ReminderTime`,`mails`.`RemoteStatus` AS `RemoteStatus`,`mails`.`ReplyRecipientNames` AS `ReplyRecipientNames`,`mails`.`SaveSentMessageFolder` AS `SaveSentMessageFolder`,`mails`.`SentOn` AS `SentOn`,`mails`.`SentOnBehalfOfName` AS `SentOnBehalfOfName`,`mails`.`Session` AS `Session`,`mails`.`VotingOptions` AS `VotingOptions`,`mails`" & _ 
  ".`VotingResponse` AS `VotingResponse` from `mails`"
End Sub ' FüllStr10

Sub FüllStr11()
 Str(0, 11, 0) = "mailsendername"
 Str(0, 11, 1) = "`id`"
 Str(0, 11, 2) = "`SenderName`"
 Str(0, 11, 3) = "`id`"
 Str(0, 11, 4) = "`SenderName`"
 ArtZ(0, 11) = 2
 ArtZ(1, 11) = 2
 Str(1, 11, 0) = "CREATE TABLE `mailsendername` ("
 Str(1, 11, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 11, 2) = " `SenderName` varchar(483) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 11, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 11, 4) = "  KEY `SenderName` (`SenderName`(20))"
 Str(1, 11, 5) = " ENGINE=MyISAM AUTO_INCREMENT=26214 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='SenderName von mails'"
End Sub ' FüllStr11

Sub FüllStr12()
 Str(0, 12, 0) = "mailsentonbehalfofname"
 Str(0, 12, 1) = "`id`"
 Str(0, 12, 2) = "`SentOnBehalfOfName`"
 Str(0, 12, 3) = "`id`"
 Str(0, 12, 4) = "`SentOnBehalfOfName`"
 ArtZ(0, 12) = 2
 ArtZ(1, 12) = 2
 Str(1, 12, 0) = "CREATE TABLE `mailsentonbehalfofname` ("
 Str(1, 12, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 12, 2) = " `SentOnBehalfOfName` varchar(483) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 12, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 12, 4) = "  KEY `SentOnBehalfOfName` (`SentOnBehalfOfName`(20))"
 Str(1, 12, 5) = " ENGINE=MyISAM AUTO_INCREMENT=24938 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='SentOnBehalfOfName von mails'"
End Sub ' FüllStr12

Sub FüllStr13()
 Str(0, 13, 0) = "mailsroh"
 Str(0, 13, 1) = "`ReceivedTime`"
 Str(0, 13, 2) = "`idSenderName`"
 Str(0, 13, 3) = "`idSubject`"
 Str(0, 13, 4) = "`idBody`"
 Str(0, 13, 5) = "`idSentOnBehalfOfName`"
 Str(0, 13, 6) = "`idReplyRecipientNames`"
 Str(0, 13, 7) = "`idCategories`"
 Str(0, 13, 8) = "`idConversationTopic`"
 Str(0, 13, 9) = "`idBCC`"
 Str(0, 13, 10) = "`idCC`"
 Str(0, 13, 11) = "`idTo`"
 Str(0, 13, 12) = "`idHTMLBody`"
 Str(0, 13, 13) = "`SentOn`"
 Str(0, 13, 14) = "`ATZahl`"
 Str(0, 13, 15) = "`LastModificationTime`"
 Str(0, 13, 16) = "`OrdnerEntryID`"
 Str(0, 13, 17) = "`MID`"
 Str(0, 13, 18) = "`AlternateRecipientAllowed`"
 Str(0, 13, 19) = "`AutoForwarded`"
 Str(0, 13, 20) = "`Class`"
 Str(0, 13, 21) = "`CreationTime`"
 Str(0, 13, 22) = "`DeferredDeliveryTime`"
 Str(0, 13, 23) = "`DeleteAfterSubmit`"
 Str(0, 13, 24) = "`EntryID`"
 Str(0, 13, 25) = "`ExpiryTime`"
 Str(0, 13, 26) = "`FlagDueBy`"
 Str(0, 13, 27) = "`FlagRequest`"
 Str(0, 13, 28) = "`Importance`"
 Str(0, 13, 29) = "`MessageClass`"
 Str(0, 13, 30) = "`NoAging`"
 Str(0, 13, 31) = "`ReceivedByEntryID`"
 Str(0, 13, 32) = "`ReceivedByName`"
 Str(0, 13, 33) = "`Saved`"
 Str(0, 13, 34) = "`Sensitivity`"
 Str(0, 13, 35) = "`Sent`"
 Str(0, 13, 36) = "`Size`"
 Str(0, 13, 37) = "`Submitted`"
 Str(0, 13, 38) = "`Unread`"
 Str(0, 13, 39) = "`Cpt`"
 Str(0, 13, 40) = "`Application`"
 Str(0, 13, 41) = "`BillingInformation`"
 Str(0, 13, 42) = "`Companies`"
 Str(0, 13, 43) = "`ConversationIndex`"
 Str(0, 13, 44) = "`FlagStatus`"
 Str(0, 13, 45) = "`FormDescription`"
 Str(0, 13, 46) = "`Mileage`"
 Str(0, 13, 47) = "`OriginatorDeliveryReportRequested`"
 Str(0, 13, 48) = "`OutlookInternalVersion`"
 Str(0, 13, 49) = "`OutlookVersion`"
 Str(0, 13, 50) = "`Parent`"
 Str(0, 13, 51) = "`ReadReceiptRequested`"
 Str(0, 13, 52) = "`ReceivedOnBehalfOfEntryID`"
 Str(0, 13, 53) = "`ReceivedOnBehalfOfName`"
 Str(0, 13, 54) = "`RecipientReassignmentProhibited`"
 Str(0, 13, 55) = "`ReminderOverrideDefault`"
 Str(0, 13, 56) = "`ReminderPlaySound`"
 Str(0, 13, 57) = "`ReminderSet`"
 Str(0, 13, 58) = "`ReminderSoundFile`"
 Str(0, 13, 59) = "`ReminderTime`"
 Str(0, 13, 60) = "`RemoteStatus`"
 Str(0, 13, 61) = "`SaveSentMessageFolder`"
 Str(0, 13, 62) = "`Session`"
 Str(0, 13, 63) = "`VotingOptions`"
 Str(0, 13, 64) = "`VotingResponse`"
 Str(0, 13, 65) = "`MID`"
 Str(0, 13, 66) = "`Zuordnung`"
 Str(0, 13, 67) = "`Ident`"
 Str(0, 13, 68) = "`idSubject`"
 Str(0, 13, 69) = "`idbody`"
 ArtZ(0, 13) = 64
 ArtZ(1, 13) = 5
 Str(1, 13, 0) = "CREATE TABLE `mailsroh` ("
 Str(1, 13, 1) = " `ReceivedTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 13, 2) = " `idSenderName` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 3) = " `idSubject` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 4) = " `idBody` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 5) = " `idSentOnBehalfOfName` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 6) = " `idReplyRecipientNames` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 7) = " `idCategories` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 8) = " `idConversationTopic` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 9) = " `idBCC` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 10) = " `idCC` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 11) = " `idTo` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 12) = " `idHTMLBody` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 13) = " `SentOn` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 13, 14) = " `ATZahl` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 15) = " `LastModificationTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 13, 16) = " `OrdnerEntryID` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 17) = " `MID` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 13, 18) = " `AlternateRecipientAllowed` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 19) = " `AutoForwarded` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 20) = " `Class` varchar(2) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 21) = " `CreationTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 13, 22) = " `DeferredDeliveryTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 13, 23) = " `DeleteAfterSubmit` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 24) = " `EntryID` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 25) = " `ExpiryTime` varchar(10) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 26) = " `FlagDueBy` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 13, 27) = " `FlagRequest` varchar(54) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 13, 28) = " `Importance` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 29) = " `MessageClass` varchar(22) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 30) = " `NoAging` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 31) = " `ReceivedByEntryID` varchar(42) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 32) = " `ReceivedByName` varchar(19) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 13, 33) = " `Saved` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 34) = " `Sensitivity` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 35) = " `Sent` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 36) = " `Size` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 37) = " `Submitted` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 38) = " `Unread` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 39) = " `Cpt` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Computername'"
 Str(1, 13, 40) = " `Application` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 41) = " `BillingInformation` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 42) = " `Companies` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 43) = " `ConversationIndex` varchar(381) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 13, 44) = " `FlagStatus` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 45) = " `FormDescription` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 46) = " `Mileage` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 47) = " `OriginatorDeliveryReportRequested` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 48) = " `OutlookInternalVersion` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 49) = " `OutlookVersion` varchar(4) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 13, 50) = " `Parent` varchar(48) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 13, 51) = " `ReadReceiptRequested` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 52) = " `ReceivedOnBehalfOfEntryID` varchar(42) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 53) = " `ReceivedOnBehalfOfName` varchar(19) COLLATE latin1_german2_ci DEFAULT ''"
 Str(1, 13, 54) = " `RecipientReassignmentProhibited` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 55) = " `ReminderOverrideDefault` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 56) = " `ReminderPlaySound` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 57) = " `ReminderSet` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 13, 58) = " `ReminderSoundFile` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 59) = " `ReminderTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 13, 60) = " `RemoteStatus` int(10) NOT NULL DEFAULT '0'"
 Str(1, 13, 61) = " `SaveSentMessageFolder` varchar(17) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 62) = " `Session` varchar(4) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 63) = " `VotingOptions` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 64) = " `VotingResponse` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 13, 65) = "  PRIMARY KEY (`MID`)"
 Str(1, 13, 66) = "  UNIQUE KEY `Zuordnung` (`EntryID`,`Cpt`)"
 Str(1, 13, 67) = "  KEY `Ident` (`SentOn`,`idSenderName`)"
 Str(1, 13, 68) = "  KEY `idSubject` (`idSubject`)"
 Str(1, 13, 69) = "  KEY `idbody` (`idBody`)"
 Str(1, 13, 70) = " ENGINE=InnoDB AUTO_INCREMENT=661261 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr13

Sub FüllStr14()
 Str(0, 14, 0) = "mailsubject"
 Str(0, 14, 1) = "`id`"
 Str(0, 14, 2) = "`subject`"
 Str(0, 14, 3) = "`id`"
 Str(0, 14, 4) = "`subject`"
 ArtZ(0, 14) = 2
 ArtZ(1, 14) = 2
 Str(1, 14, 0) = "CREATE TABLE `mailsubject` ("
 Str(1, 14, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 14, 2) = " `subject` varchar(255) COLLATE latin1_german2_ci NOT NULL"
 Str(1, 14, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 14, 4) = "  KEY `subject` (`subject`(20))"
 Str(1, 14, 5) = " ENGINE=MyISAM AUTO_INCREMENT=66625 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Subject von mails'"
End Sub ' FüllStr14

Sub FüllStr15()
 Str(0, 15, 0) = "mailto"
 Str(0, 15, 1) = "`id`"
 Str(0, 15, 2) = "`to`"
 Str(0, 15, 3) = "`id`"
 Str(0, 15, 4) = "`to`"
 ArtZ(0, 15) = 2
 ArtZ(1, 15) = 2
 Str(1, 15, 0) = "CREATE TABLE `mailto` ("
 Str(1, 15, 1) = " `id` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 15, 2) = " `to` text COLLATE latin1_german2_ci NOT NULL"
 Str(1, 15, 3) = "  PRIMARY KEY (`id`)"
 Str(1, 15, 4) = "  KEY `to` (`to`(20))"
 Str(1, 15, 5) = " ENGINE=MyISAM AUTO_INCREMENT=8087 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='`to` von mails'"
End Sub ' FüllStr15

Sub FüllStr16()
 Str(0, 16, 0) = "ordner"
 Str(0, 16, 1) = "`Tiefe`"
 Str(0, 16, 2) = "`Class`"
 Str(0, 16, 3) = "`DefaultItemType`"
 Str(0, 16, 4) = "`DefaultMessageClass`"
 Str(0, 16, 5) = "`Description`"
 Str(0, 16, 6) = "`EntryID`"
 Str(0, 16, 7) = "`Name`"
 Str(0, 16, 8) = "`Parent`"
 Str(0, 16, 9) = "`Session`"
 Str(0, 16, 10) = "`StoreID`"
 Str(0, 16, 11) = "`UnReadItemCount`"
 Str(0, 16, 12) = "`WebViewAllowNavigation`"
 Str(0, 16, 13) = "`WebViewOn`"
 Str(0, 16, 14) = "`WebViewURL`"
 Str(0, 16, 15) = "`Cpt`"
 Str(0, 16, 16) = "`EntryID`"
 ArtZ(0, 16) = 15
 ArtZ(1, 16) = 1
 Str(1, 16, 0) = "CREATE TABLE `ordner` ("
 Str(1, 16, 1) = " `Tiefe` int(10) NOT NULL DEFAULT '0'"
 Str(1, 16, 2) = " `Class` int(10) NOT NULL DEFAULT '0'"
 Str(1, 16, 3) = " `DefaultItemType` int(10) NOT NULL DEFAULT '0' COMMENT '1=Termin,2=Kontakt,4=Journal,0=Mail,5=Notiz,6=olPostItem,3=olTaskItem'"
 Str(1, 16, 4) = " `DefaultMessageClass` varchar(15) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 16, 5) = " `Description` varchar(26) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 16, 6) = " `EntryID` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 16, 7) = " `Name` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 16, 8) = " `Parent` varchar(23) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 16, 9) = " `Session` varchar(4) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 16, 10) = " `StoreID` varchar(342) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 16, 11) = " `UnReadItemCount` int(10) NOT NULL DEFAULT '0'"
 Str(1, 16, 12) = " `WebViewAllowNavigation` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 16, 13) = " `WebViewOn` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 16, 14) = " `WebViewURL` varchar(71) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 16, 15) = " `Cpt` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Computername'"
 Str(1, 16, 16) = "  UNIQUE KEY `EntryID` (`EntryID`,`Cpt`)"
 Str(1, 16, 17) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr16

Sub FüllStr17()
 Str(0, 17, 0) = "recipients"
 Str(0, 17, 1) = "`iid`"
 Str(0, 17, 2) = "`MailEntryID`"
 Str(0, 17, 3) = "`AddressEntryIID`"
 Str(0, 17, 4) = "`AutoResponse`"
 Str(0, 17, 5) = "`Index`"
 Str(0, 17, 6) = "`MeetingResponseStatus`"
 Str(0, 17, 7) = "`Resolved`"
 Str(0, 17, 8) = "`TrackingStatus`"
 Str(0, 17, 9) = "`TrackingStatusTime`"
 Str(0, 17, 10) = "`Type`"
 Str(0, 17, 11) = "`Nr`"
 Str(0, 17, 12) = "`Class`"
 Str(0, 17, 13) = "`Cpt`"
 Str(0, 17, 14) = "`iid`"
 Str(0, 17, 15) = "`MailEntryID`"
 Str(0, 17, 16) = "`Zuordnung`"
 ArtZ(0, 17) = 13
 ArtZ(1, 17) = 3
 Str(1, 17, 0) = "CREATE TABLE `recipients` ("
 Str(1, 17, 1) = " `iid` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 17, 2) = " `MailEntryID` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 17, 3) = " `AddressEntryIID` int(10) NOT NULL DEFAULT '0'"
 Str(1, 17, 4) = " `AutoResponse` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 17, 5) = " `Index` int(10) NOT NULL DEFAULT '0'"
 Str(1, 17, 6) = " `MeetingResponseStatus` int(10) NOT NULL DEFAULT '0'"
 Str(1, 17, 7) = " `Resolved` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 17, 8) = " `TrackingStatus` int(10) NOT NULL DEFAULT '0'"
 Str(1, 17, 9) = " `TrackingStatusTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 17, 10) = " `Type` int(10) NOT NULL DEFAULT '0'"
 Str(1, 17, 11) = " `Nr` int(10) NOT NULL DEFAULT '0'"
 Str(1, 17, 12) = " `Class` int(10) NOT NULL DEFAULT '0'"
 Str(1, 17, 13) = " `Cpt` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Computername'"
 Str(1, 17, 14) = "  PRIMARY KEY (`iid`)"
 Str(1, 17, 15) = "  KEY `MailEntryID` (`MailEntryID`)"
 Str(1, 17, 16) = "  KEY `Zuordnung` (`MailEntryID`,`AddressEntryIID`,`Cpt`)"
 Str(1, 17, 17) = " ENGINE=InnoDB AUTO_INCREMENT=825624 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr17

Sub FüllStr18()
 Str(0, 18, 0) = "replyrecipients"
 Str(0, 18, 1) = "`iid`"
 Str(0, 18, 2) = "`MailEntryID`"
 Str(0, 18, 3) = "`AddressEntryIID`"
 Str(0, 18, 4) = "`AutoResponse`"
 Str(0, 18, 5) = "`Index`"
 Str(0, 18, 6) = "`MeetingResponseStatus`"
 Str(0, 18, 7) = "`Resolved`"
 Str(0, 18, 8) = "`TrackingStatus`"
 Str(0, 18, 9) = "`TrackingStatusTime`"
 Str(0, 18, 10) = "`Type`"
 Str(0, 18, 11) = "`Nr`"
 Str(0, 18, 12) = "`Class`"
 Str(0, 18, 13) = "`Cpt`"
 Str(0, 18, 14) = "`AddressEntryID`"
 Str(0, 18, 15) = "`iid`"
 Str(0, 18, 16) = "`Zuordnung`"
 Str(0, 18, 17) = "`MainEntryID`"
 ArtZ(0, 18) = 14
 ArtZ(1, 18) = 3
 Str(1, 18, 0) = "CREATE TABLE `replyrecipients` ("
 Str(1, 18, 1) = " `iid` int(10) NOT NULL AUTO_INCREMENT"
 Str(1, 18, 2) = " `MailEntryID` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 18, 3) = " `AddressEntryIID` varchar(255) COLLATE latin1_german2_ci NOT NULL DEFAULT '0'"
 Str(1, 18, 4) = " `AutoResponse` varchar(1) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 18, 5) = " `Index` int(10) NOT NULL DEFAULT '0'"
 Str(1, 18, 6) = " `MeetingResponseStatus` int(10) NOT NULL DEFAULT '0'"
 Str(1, 18, 7) = " `Resolved` tinyint(1) NOT NULL DEFAULT '0'"
 Str(1, 18, 8) = " `TrackingStatus` int(10) NOT NULL DEFAULT '0'"
 Str(1, 18, 9) = " `TrackingStatusTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 18, 10) = " `Type` int(10) NOT NULL DEFAULT '0'"
 Str(1, 18, 11) = " `Nr` int(10) NOT NULL DEFAULT '0'"
 Str(1, 18, 12) = " `Class` int(10) NOT NULL DEFAULT '0'"
 Str(1, 18, 13) = " `Cpt` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Computername'"
 Str(1, 18, 14) = " `AddressEntryID` int(10) NOT NULL"
 Str(1, 18, 15) = "  PRIMARY KEY (`iid`)"
 Str(1, 18, 16) = "  UNIQUE KEY `Zuordnung` (`MailEntryID`,`AddressEntryIID`,`Cpt`)"
 Str(1, 18, 17) = "  KEY `MainEntryID` (`MailEntryID`)"
 Str(1, 18, 18) = " ENGINE=InnoDB AUTO_INCREMENT=128198 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr18

Sub FüllStr19()
 Str(0, 19, 0) = "subfolders"
 Str(0, 19, 1) = "`EntryID`"
 Str(0, 19, 2) = "`subEntryID`"
 Str(0, 19, 3) = "`Nr`"
 Str(0, 19, 4) = "`Cpt`"
 Str(0, 19, 5) = "`Zuordnung`"
 Str(0, 19, 6) = "`EntryID`"
 Str(0, 19, 7) = "`subEntryID`"
 ArtZ(0, 19) = 4
 ArtZ(1, 19) = 3
 Str(1, 19, 0) = "CREATE TABLE `subfolders` ("
 Str(1, 19, 1) = " `EntryID` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 19, 2) = " `subEntryID` varchar(48) COLLATE latin1_german2_ci NOT NULL DEFAULT ''"
 Str(1, 19, 3) = " `Nr` int(10) NOT NULL DEFAULT '0'"
 Str(1, 19, 4) = " `Cpt` varchar(7) COLLATE latin1_german2_ci NOT NULL DEFAULT '' COMMENT 'Computername'"
 Str(1, 19, 5) = "  UNIQUE KEY `Zuordnung` (`EntryID`,`subEntryID`,`Cpt`)"
 Str(1, 19, 6) = "  KEY `EntryID` (`EntryID`)"
 Str(1, 19, 7) = "  KEY `subEntryID` (`subEntryID`)"
 Str(1, 19, 8) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci"
End Sub ' FüllStr19

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

Public Function doMach_mails(DBn$, Optional Server$, Optional obStumm%=True) ' Datenbankname
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
 call doex("SET FOREIGN_KEY_CHECKS = 0",0)

 Dim j&, ZZ&, Tbl$, sql As new CString
 For i = 0 To 19
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
 For i = 0 To 19
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
  For i = 0 To 19
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
  MsgBox "Fertig mit doMach_mails(" & DBn & "," & Server & ")!
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_mails/" + AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): Ende
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_mails

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
