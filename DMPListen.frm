VERSION 5.00
Begin VB.Form DMPListen 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "DMP-Listen"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox erfBis 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox erfVon 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox eingBis 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox eingVon 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label erfasst 
      Caption         =   "er&fasst:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label eingereicht 
      Caption         =   "ein&gereicht:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "DMPListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mLese As Lese
Option Explicit
Private Type DMPEin
 Pat_id As Long
 zp As Date
 Abk As String
 ErstDS As String
 ErstD As Date
 Datum As Date
 NachName As String
 VorName As String
 GebDat As Date
End Type

Private Sub CancelButton_Click()
  Me.Hide
  Call mLese.ProgEnde
End Sub

Public Sub init(varmLese As Lese)
 Set mLese = varmLese
End Sub

Private Sub Form_Load()
 Me.erfVon = "1.6.04"
 Me.erfBis = Format(Now + 1, "d.m.yy")
 Me.eingVon = "1.6.04"
 Me.eingBis = Format(Now + 1, "d.m.yy")
End Sub
Private Sub OKButton_Click()
  Const AusgDat$ = "p:\DMPListe.txt"
  Const AusgDat1$ = "p:\DMPListeRein.txt"
  Const FLen& = 30
    
  Dim AuS$, i%, AktL& ' , altPat_id&, obDruck%,
  Dim lfdnr&, lfdf&
  Dim DE As DMPEin
  Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
  Dim OkSql$
  Dim FormFrisch%, FormFeld$
  #Const debugs = 0
  
  On Error GoTo fehler
  If Lies.obMySQL = 0 Then On Error Resume Next
  Set rs = DBCn.Execute("drop table" & ifexists & "dmpreihe")
  sql = "create table" & ifnotexists & "dmpreihe (ZP date, Abk " & sqlTEXT & "(30), ErstD date, Datum date, NachName " & sqlTEXT & "(20), VorName " & sqlTEXT & "(20), GebDat Date, Pat_id " & sqlLong & ")"
  Set rs = DBCn.Execute(sql)  ' Telefonbuchsortierung Ä=AE, ß = ss; wörterbuchsortiung german1 wäre Ä=A, ß=s
  If Lies.obMySQL Then On Error GoTo fehler
'  Set rs = DBCn.Execute("grant all privileges on dmpreihe to praxis with grant option")
  #If debugs Then
   Open AusgDat For Output As #307
  #End If
  Open AusgDat1 For Output As #308
  OkSql = "SELECT namen.pat_id as pat_id,zeitpunkt,form_abk,nachname,vorname,gebdat,feld,feldnr,feldinh  FROM ((((forminhkopf left join formulare on forminhkopf.form_id = formulare.formid) left join forminhfeld on forminhfeld.foid = forminhkopf.foid) left join forminhaltfeld on forminhaltfeld.feldvw = forminhfeld.feldvw) left join forminhaltfeldinh on forminhaltfeldinh.feldinhvw = forminhfeld.feldinhvw) left join namen on forminhkopf.pat_id = namen.pat_id where form_abk like ""DMPDTYP" & "%" & """ and (feldinh like ""_"" or feldinh like ""%.%.%"") order by namen.pat_id, forminhkopf.form_id, forminhfeld.feldvw, nr"
'  If Not lies.obmysql Then OkSql = Replace(OkSql, "%", "*")
  rs1.Open "select distinct pat_id from (" & OkSql & ") as innen", DBCn, adOpenDynamic, adLockReadOnly
  Do While Not rs1.EOF
  Set rs = Nothing
  rs.Open "select * from (" & OkSql & ") as innen where pat_id = " & rs1!Pat_id, DBCn, adOpenDynamic, adLockReadOnly
'  rs.Open "SELECT * FROM forminhkopf left join formulare on forminhkopf.form_id = formulare.formid left join forminhfeld on forminhfeld.foid = forminhkopf.foid where form_abk like ""DMPDTYP%"" limit 1000;", DBCn, adOpenDynamic, adLockReadOnly
' SELECT * FROM forminhkopf left join formulare on forminhkopf.form_id = formulare.formid left join forminhfeld on forminhfeld.foid = forminhkopf.foid where form_abk like "DMPDTYP%" limit 1000;
' left join forminhaltfeld on forminhaltfeld.feldvw = forminhfeld.feldvw left join forminhaltfeldinh on forminhaltfeldinh.feldinhvw = forminhfeld.feldinhvw
  #If debugs Then
   For i = 0 To rs.Fields.Count - 1
    AktL = rs.Fields(i).DefinedSize + 3
    AuS = AuS & Right(Space(AktL) & rs.Fields(i).Name, AktL) & " "
   Next
   Print #307, AuS
  #End If
  
  Do
'   If rs!Pat_id = 138 Then Stop
'   altPat_id = rs.Fields(2) 'rs!Pat_id
   #If debugs Then
    AuS = ""
    For i = 0 To rs.Fields.Count - 1
     AktL = rs.Fields(i).DefinedSize + 3
     AuS = AuS & Right(Space(AktL) & rs.Fields(i).Value, AktL) & " "
    Next
    Print #307, AuS
   #End If
' Eigentlich heißen die relevanten Felder "Datum" bzw. "Erstellungsdatum". Die Namen können aber bei Download von falschem Computer aus vertauscht sein:
' -> allgemeinere Methode muß gewählt werden
' das erste Datenfeld mit Datumsformatierung entspricht bei richtigem Formular dem Feld "Datum"
   If DE.Datum = 0 Then
    If Len(rs!FeldInh) > 1 Then
     If IsDate(rs!FeldInh) Then
      DE.Datum = rs!FeldInh
      FormFrisch = False
     End If
    End If
   End If
   If Not FormFrisch And rs!FeldNr = 0 And Len(rs!FeldInh) = 1 Then
    DE.Pat_id = rs!Pat_id
    DE.zp = rs!Zeitpunkt
    DE.Abk = rs!Form_Abk
    DE.ErstDS = ""
    DE.NachName = rs!NachName
    DE.VorName = rs!VorName
    DE.GebDat = rs!GebDat
    DE.Pat_id = rs!Pat_id
    DE.Datum = 0 ' nach dem Erstellungsdatumfeld kommt kein Datumsfeld mehr
    FormFeld = rs!Feld
   End If
   If rs!Feld = FormFeld Then
    DE.ErstDS = DE.ErstDS & rs!FeldInh
    If rs!FeldNr = 7 Then
     DE.ErstD = BDTtoDate(DE.ErstDS)
     Print #308, DE.Pat_id & " " & DE.NachName & " " & DE.VorName & " " & DE.GebDat & " " & DE.zp & " " & DE.Abk & " " & DE.ErstD & " " & DE.Datum
     DBCn.Execute ("insert into dmpreihe values(" & datform(DE.zp) & ",'" & DE.Abk & "'," & datform(DE.ErstD) & "," & datform(DE.Datum) & ",'" & DE.NachName & "','" & DE.VorName & "'," & datform(DE.GebDat) & "," & DE.Pat_id & ")")
     Debug.Print DE.Pat_id & " " & DE.NachName & " " & DE.VorName & " " & DE.GebDat & " " & DE.zp & " " & DE.Abk & " " & DE.ErstD & " " & DE.Datum
     DBCn.Execute ("insert into dmpreihe values(" & datform(DE.zp) & ",'" & DE.Abk & "'," & datform(DE.ErstD) & "," & datform(DE.Datum) & ",'" & DE.NachName & "','" & DE.VorName & "'," & datform(DE.GebDat) & "," & DE.Pat_id & ")")
     DoEvents
     FormFrisch = True
     FormFeld = ""
    End If
   End If
   
#If False Then
    Select Case rs!Feld
     Case "Datum", "Vertragsarztnr" ' durch falsches Formular
      If FormFrisch Then
       DE.Datum = rs!FeldInh
       FormFrisch = False
      End If
     Case "Erstellungsdatum", "Check2" ' durch falsches Formular
      If rs!FeldNr = 0 And FormFeld = "" Then
       DE.Pat_id = rs!Pat_id
       DE.zp = rs!Zeitpunkt
       DE.Abk = rs!Form_Abk
       DE.ErstDS = ""
       DE.NachName = rs!NachName
       DE.VorName = rs!VorName
       DE.GebDat = rs!GebDat
       DE.Pat_id = rs!Pat_id
       FormFeld = rs!Feld
      End If
      If rs!Feld = FormFeld Then
       DE.ErstDS = DE.ErstDS & rs!FeldInh
       If rs!FeldNr = 7 Then
' Fehler hier nicht abfangen, da sie nur bei falschem SQL-Befehl auftreten
'        If Len(DE.ErstDS) = 8 Then
'         If Left(DE.ErstDS, 2) > 0 And Left(DE.ErstDS, 2) < 31 Then
          DE.ErstD = BDTtoDate(DE.ErstDS)
'         End If
'        End If
        Print #308, DE.Pat_id & " " & DE.NachName & " " & DE.VorName & " " & DE.GebDat & " " & DE.zp & " " & DE.Abk & " " & DE.ErstD & " " & DE.Datum
        DBCn.Execute ("insert into dmpreihe values(" & datform(DE.zp) & ",'" & DE.Abk & "'," & datform(DE.ErstD) & "," & datform(DE.Datum) & ",'" & DE.NachName & "','" & DE.VorName & "'," & datform(DE.GebDat) & "," & DE.Pat_id & ")")
        FormFrisch = True
        FormFeld = ""
       End If
      End If
    End Select
#End If
'   obDruck = 0
'   Dim altod%
'   altod = obDruck
'   If rs!FeldNr = 7 Then obDruck = True
   rs.Move 1
'   If rs.EOF Then obDruck = True Else If rs.Fields(2) <> altPat_id Then DoEvents: Debug.Print lfdnr, rs.Fields(2), rs!NachName, rs!VorName: lfdnr = lfdnr + 1: obDruck = True ' forminhkopf.pat_id
'   If obDruck Then
'    If Not altod Then Debug.Print "Fehler bei Pat_id: ", lfdf, DE.Pat_id: lfdf = lfdf + 1
'   End If
   If rs.EOF Then Exit Do
  Loop
  rs1.Move 1
  Loop
  #If debugs Then
   Close #307
   Call Shell("notepad " & AusgDat, vbMaximizedFocus)
  #End If
  Close #308
  Call Shell("notepad " & AusgDat1, vbMaximizedFocus)
  Me.Hide
  Call mLese.ProgEnde
  Unload Me
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in okbutton_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub
