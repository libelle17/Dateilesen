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
 Pat_ID As Long
 Zp As Date
 Abk As String
 ErstDS As String
 ErstD As Date
 Datum As Date
 Nachname As String
 Vorname As String
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
  
  Dim AuS$, i%, AktL&
  Dim DE As DMPEin
  Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
  
  On Error GoTo fehler
  If obMySQL = 0 Then On Error Resume Next
  Set rs = DBCn.Execute("drop table" & ifexists & "dmpreihe")
  Set rs = DBCn.Execute("create table" & ifnotexists & "dmpreihe (ZP date, Abk varchar(30), ErstD date, Datum date, NachName varchar(20), VorName varchar(20), GebDat Date);")  ' Telefonbuchsortierung Ä=AE, ß = ss; wörterbuchsortiung german1 wäre Ä=A, ß=s
  If obMySQL Then On Error GoTo fehler
'  Set rs = DBCn.Execute("grant all privileges on dmpreihe to praxis with grant option")
  
  Open AusgDat For Output As #307
  Open AusgDat1 For Output As #308
  rs.Open "SELECT * FROM ((((forminhkopf left join formulare on forminhkopf.form_id = formulare.formid) left join forminhfeld on forminhfeld.foid = forminhkopf.foid) left join forminhaltfeld on forminhaltfeld.feldvw = forminhfeld.feldvw) left join forminhaltfeldinh on forminhaltfeldinh.feldinhvw = forminhfeld.feldinhvw) left join namen on forminhkopf.pat_id = namen.pat_id where form_abk like ""DMPDTYP" & sqlStern & """ and feld like """ & sqlStern & "datum"";", DBCn, adOpenDynamic, adLockReadOnly
'  rs.Open "SELECT * FROM forminhkopf left join formulare on forminhkopf.form_id = formulare.formid left join forminhfeld on forminhfeld.foid = forminhkopf.foid where form_abk like ""DMPDTYP%"" limit 1000;", DBCn, adOpenDynamic, adLockReadOnly
' SELECT * FROM forminhkopf left join formulare on forminhkopf.form_id = formulare.formid left join forminhfeld on forminhfeld.foid = forminhkopf.foid where form_abk like "DMPDTYP%" limit 1000;
' left join forminhaltfeld on forminhaltfeld.feldvw = forminhfeld.feldvw left join forminhaltfeldinh on forminhaltfeldinh.feldinhvw = forminhfeld.feldinhvw
  For i = 0 To rs.Fields.Count - 1
   AktL = rs.Fields(i).DefinedSize + 3
   AuS = AuS & Right(Space(AktL) & rs.Fields(i).Name, AktL) & " "
  Next
  Print #307, AuS
  Do
   AuS = ""
   For i = 0 To rs.Fields.Count - 1
    AktL = rs.Fields(i).DefinedSize + 3
    AuS = AuS & Right(Space(AktL) & rs.Fields(i).Value, AktL) & " "
   Next
   Select Case rs!Feld
    Case "Datum"
     DE.Datum = rs!FeldInh
    Case "Erstellungsdatum"
     If rs!FeldNr = 0 Then
      DE.Pat_ID = rs!Pat_ID
      DE.Zp = rs!Zeitpunkt
      DE.Abk = rs!Form_Abk
      DE.ErstDS = ""
      DE.Nachname = rs!Nachname
      DE.Vorname = rs!Vorname
      DE.GebDat = rs!GebDat
     End If
     DE.ErstDS = DE.ErstDS & rs!FeldInh
     If rs!FeldNr = 7 Then
      DE.ErstD = BDTtoDate(DE.ErstDS)
      Print #308, DE.Pat_ID & " " & DE.Nachname & " " & DE.Vorname & " " & DE.GebDat & " " & DE.Zp & " " & DE.Abk & " " & DE.ErstD & " " & DE.Datum
      DBCn.Execute ("insert into dmpreihe values(" & datForm(DE.Zp) & ",'" & DE.Abk & "'," & datForm(DE.ErstD) & "," & datForm(DE.Datum) & ",'" & DE.Nachname & "','" & DE.Vorname & "'," & datForm(DE.GebDat) & ")")
     End If
   End Select
   Print #307, AuS
   rs.Move 1
   If rs.EOF Then Exit Do
  Loop
  Close #307
  Close #308
  Call Shell("notepad " & AusgDat, vbMaximizedFocus)
  Call Shell("notepad " & AusgDat1, vbMaximizedFocus)
  Me.Hide
  Call mLese.ProgEnde
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
