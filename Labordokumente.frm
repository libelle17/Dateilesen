VERSION 5.00
Begin VB.Form Labordokumente 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Labordokumente"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Labordokumente"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   19620
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Pat_id"
      Height          =   256
      Index           =   1
      Left            =   57
      TabIndex        =   9
      Top             =   0
      Width           =   456
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H0046A3FF&
      DataField       =   "Nachname"
      Height          =   256
      Index           =   2
      Left            =   570
      TabIndex        =   10
      Top             =   0
      Width           =   1083
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Vorname"
      Height          =   256
      Index           =   3
      Left            =   1710
      TabIndex        =   11
      Top             =   0
      Width           =   1083
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "ZeitPunkt"
      Height          =   256
      Index           =   4
      Left            =   2850
      TabIndex        =   12
      Top             =   0
      Width           =   1629
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "DokName"
      Height          =   256
      Index           =   5
      Left            =   6236
      TabIndex        =   13
      Top             =   0
      Width           =   5727
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "DokPfad"
      Height          =   256
      Index           =   6
      Left            =   17631
      TabIndex        =   14
      Top             =   0
      Width           =   1992
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Anzeigen"
      Height          =   313
      Index           =   1
      Left            =   14117
      TabIndex        =   15
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "PhotoImpact"
      Height          =   314
      Index           =   2
      Left            =   15760
      TabIndex        =   16
      Top             =   0
      Width           =   786
   End
   Begin VB.TextBox vTextB 
      Height          =   256
      Index           =   7
      Left            =   17178
      TabIndex        =   17
      Top             =   0
      Width           =   164
   End
   Begin VB.TextBox vTextB 
      Height          =   256
      Index           =   8
      Left            =   16951
      TabIndex        =   18
      Top             =   0
      Width           =   167
   End
   Begin VB.TextBox vTextB 
      Height          =   256
      Index           =   9
      Left            =   17405
      TabIndex        =   19
      Top             =   0
      Width           =   170
   End
   Begin VB.TextBox vTextB 
      Height          =   256
      Index           =   10
      Left            =   16611
      TabIndex        =   20
      Top             =   0
      Width           =   227
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "AktZeit"
      Height          =   256
      Index           =   11
      Left            =   4535
      TabIndex        =   21
      Top             =   0
      Width           =   1644
   End
   Begin VB.CommandButton vCommandB 
      Caption         =   "Abhak"
      Height          =   314
      Index           =   3
      Left            =   15023
      TabIndex        =   22
      Top             =   0
      Width           =   666
   End
   Begin VB.TextBox vTextB 
      BackColor       =   &H00C0C0C0&
      DataField       =   "DokGroe"
      Height          =   256
      Index           =   12
      Left            =   12075
      TabIndex        =   23
      Top             =   0
      Width           =   1992
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   19620
      TabIndex        =   0
      Top             =   615
      Width           =   19620
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Hinzufügen"
         Height          =   300
         Left            =   59
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Bearbeiten"
         Height          =   300
         Left            =   1213
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Löschen"
         Height          =   300
         Left            =   2367
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "neu lesen"
         Height          =   300
         Left            =   3521
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "S&chließen"
         Height          =   300
         Left            =   4675
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Aktualisieren"
         Height          =   300
         Left            =   59
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Abbrechen"
         Height          =   300
         Left            =   1213
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   19620
      TabIndex        =   31
      Top             =   915
      Width           =   19620
      Begin VB.CommandButton Suchen 
         Caption         =   "&Suchen"
         Height          =   255
         Left            =   4920
         TabIndex        =   32
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox suche 
         Height          =   285
         Left            =   5640
         TabIndex        =   33
         Top             =   0
         Width           =   2655
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "Labordokumente.frx":0000
         Style           =   1  'Grafisch
         TabIndex        =   34
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&ü"
         Height          =   300
         Left            =   345
         Picture         =   "Labordokumente.frx":0342
         Style           =   1  'Grafisch
         TabIndex        =   35
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&ä"
         Height          =   300
         Left            =   4200
         Picture         =   "Labordokumente.frx":0684
         Style           =   1  'Grafisch
         TabIndex        =   36
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "Labordokumente.frx":09C6
         Style           =   1  'Grafisch
         TabIndex        =   37
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   690
         TabIndex        =   38
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label vLab 
      Caption         =   "Pat_id"
      Height          =   240
      Index           =   1
      Left            =   57
      TabIndex        =   1
      Top             =   57
      Width           =   456
   End
   Begin VB.Label vLab 
      Caption         =   "Nachname"
      Height          =   240
      Index           =   2
      Left            =   570
      TabIndex        =   2
      Top             =   57
      Width           =   1083
   End
   Begin VB.Label vLab 
      Caption         =   "Vorname"
      Height          =   240
      Index           =   3
      Left            =   1710
      TabIndex        =   3
      Top             =   57
      Width           =   1083
   End
   Begin VB.Label vLab 
      Caption         =   "ZeitPunkt"
      Height          =   240
      Index           =   4
      Left            =   2834
      TabIndex        =   4
      Top             =   1
      Width           =   1569
   End
   Begin VB.Label vLab 
      Caption         =   "DokName"
      Height          =   240
      Index           =   5
      Left            =   5993
      TabIndex        =   5
      Top             =   0
      Width           =   5727
   End
   Begin VB.Label vLab 
      Caption         =   "Dok-Größe"
      Height          =   240
      Index           =   6
      Left            =   11776
      TabIndex        =   6
      Top             =   0
      Width           =   1992
   End
   Begin VB.Label vLab 
      Caption         =   "AktZeit"
      Height          =   240
      Index           =   7
      Left            =   4406
      TabIndex        =   7
      Top             =   0
      Width           =   1539
   End
   Begin VB.Label vLab 
      Caption         =   "DokPfad"
      Height          =   240
      Index           =   8
      Left            =   17858
      TabIndex        =   8
      Top             =   0
      Width           =   1422
   End
End
Attribute VB_Name = "Labordokumente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents adoRS As Recordset
Attribute adoRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim altsuche$, SuchStringGeändert%
Private Sub adoRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hierdurch wird die aktuelle Datensatzposition für diese Datensatzgruppe angezeigt
   lblStatus.Caption = CStr(adoRS.AbsolutePosition)
   Call do_Form_Current(Me)
End Sub 'adoRS_MoveComplete

Private Sub adoRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hier können Sie Code zur Überprüfung einfügen
  'Dieses Ereignis wird aufgerufen, wenn die folgenden Aktionen eintreten
  Dim bCancel As Boolean
  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select
  If bCancel Then adStatus = adStatusCancel
End Sub 'adoRS_WillChangeRecord

Private Sub cmdAdd_Click()
  On Error GoTo fehler
  With adoRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Datensatz; hinzufügen"
    mbAddNewFlag = True
    SetButtons False
  End With
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdAdd_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'cmdAdd_Click()

Private Sub cmdDelete_Click()
  On Error GoTo fehler
  With adoRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdDelete_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Private Sub cmdDelete_Click()

Private Sub cmdRefresh_Click()
  'Dies wird nur für Mehrbenutzeranwendungen benötigt
  On Error GoTo fehler
  adoRS.Requery
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdRefresh_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Private Sub cmdRefresh_Click()"

Private Sub cmdEdit_Click()
  On Error GoTo fehler
  lblStatus.Caption = "Datensatz bearbeiten"
  mbEditFlag = True
  SetButtons False
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdEdit_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'cmdEdit_Click()

Private Sub cmdCancel_Click()
  On Error Resume Next
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoRS.CancelUpdate
 If mvBookMark > 0 Then
    adoRS.Bookmark = mvBookMark
  Else
    adoRS.MoveFirst
  End If
  mbDataChanged = False
End Sub ' cmdCancel_Click()

Private Sub cmdUpdate_Click()
  On Error GoTo fehler
  adoRS.UpdateBatch adAffectAll
  If mbAddNewFlag Then
    adoRS.MoveLast              'Zu neuem Datensatz gehen
  End If
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdUpdate_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdUpdate_Click()

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo fehler
  adoRS.MoveFirst
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdFirst_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdFirst_Click

Private Sub cmdLast_Click()
  On Error GoTo fehler
  adoRS.MoveLast
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdLast_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdLast_Click()

Private Sub cmdNext_Click()
  On Error GoTo fehler
  If Not adoRS.EOF Then adoRS.MoveNext
  If adoRS.EOF And adoRS.RecordCount > 0 Then
    Beep
     'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoRS.MoveLast
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdNext_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdNext_Click()

Private Sub cmdPrevious_Click()
  On Error GoTo fehler
  If Not adoRS.BOF Then adoRS.MovePrevious
  If adoRS.BOF And adoRS.RecordCount > 0 Then
    Beep
    'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoRS.MoveFirst
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdPrevious_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdPrevious_Click()

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub ' SetButtons

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Call doKeyDown(Me, KeyCode, Shift)
' End Sub

Private Sub Suchen_Click()
 On Error GoTo fehler
 If SuchStringGeändert Then
  altsuche = Me.suche
  If IsNumeric(Me.suche) Then
   adoRS.Find " Pat_id = " & Me.suche, 0, adSearchForward
  Else
   If adoRS.EOF Then adoRS.MoveFirst
   adoRS.Find "gesname like '" & Me.suche & "*'", 1, adSearchForward
  End If
 Else
  Me.suche.SetFocus
  SuchStringGeändert = True
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Suchen_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'Suchen_Click()









Private Sub Form_KeyPress(KeyCode%)
 Call Taste(KeyCode, 0, Me)
End Sub

Private Sub Form_KeyUp(KeyCode%, Shift%)
 Call Taste(KeyCode, Shift, Me)
End Sub

Private Sub Form_Load()
  Call doForm_Load(Me)
 Call do_LaborDokumente_form_load(Me)
End Sub

Private Sub Form_Open(Cancel%)
 Call do_Form_Open(Cancel, Me)
End Sub






Private Sub Form_KeyDown(KeyCode%, Shift%)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub
  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
 Call Taste(KeyCode, Shift, Me)
End Sub


Private Sub Form_Current()
 Me.Datensatz = Me.CurrentRecord
 Call do_LaborDokumente_Form_Current(Me)
End Sub

'Datensatz und Feld farbig hervorheben
'Autor: Günther Ritter - gritter@gmx.de, April 2002

Private Sub Form_AfterDelConfirm(Status%)
  Me.DatensatzPos.Requery
End Sub

Private Sub Form_AfterUpdate()
  Me.DatensatzPos.Requery
End Sub

Private Sub Detailbereich_Click()
  Me.Datensatz = Me.CurrentRecord
End Sub

Private Function DatensatzPosition() As Variant
 DatensatzPosition = doDatensatzPosition(Me)
End Function
Private Function AktAbgehakt()
 AktAbgehakt = doAktAbgehakt(Me)
End Function
Private Function obAbgehakt()
 obAbgehakt = do_obAbgehakt(Me)
End Function
Private Function AltName() As Variant
 AltName = doAltName(Me)
End Function
Private Sub Zeitpunkt_KeyPress(KeyCode%)
  Call Taste(KeyCode, 0, Me)
End Sub
Private Sub vTextB_Click(Index As Integer)
 Select Case Index
  Case 10 ' abgehakt
   Call do_abgehakt_Click(Me)
 End Select ' Index
End Sub 'Private Sub vTextB_Click(Index as Integer)

Private Sub vTextB_KeyPress(Index As Integer, KeyAscii As Integer)
 Select Case Index
  Case 1 ' Pat_id
    Call Taste(KeyCode, 0, Me)
  Case 2 ' Nachname
    Call Taste(KeyCode, 0, Me)
  Case 3 ' Vorname
    Call Taste(KeyCode, 0, Me)
  Case 5 ' DokName
    Call Taste(KeyCode, 0, Me)
  Case 6 ' DokPfad
    Call Taste(KeyCode, 0, Me)
  Case 10 ' abgehakt
    Call Taste(KeyCode, 0, Me)
  Case 11 ' AktZeit
    Call Taste(KeyCode, 0, Me)
 End Select ' Index
End Sub 'Private Sub vTextB_KeyPress(Index as Integer, KeyAscii As Integer)

Private Sub vCommandB_Click(Index As Integer)
 Select Case Index
  Case 1 ' Anzeigen
   Call do_anzeigen_click(Me)
  Case 2 ' PhotoImpact
   Call do_PhotoImpact_Click(Me)
  Case 3 ' Abhak
   Call abgehakt_Click
 End Select ' Index
End Sub 'Private Sub vCommandB_Click(Index as Integer)

Private Sub vCommandB_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Select Case Index
  Case 1 ' Anzeigen
    Call Taste(KeyCode, Shift, Me)
  Case 2 ' PhotoImpact
   Call Taste(KeyCode, Shift, Me)
 End Select ' Index
End Sub 'Private Sub vCommandB_KeyDown(Index as Integer, KeyCode As Integer, Shift As Integer)

Private Sub vCommandB_KeyPress(Index As Integer, KeyAscii As Integer)
 Select Case Index
  Case 1 ' Anzeigen
    Call Taste(KeyCode, 0, Me)
  Case 2 ' PhotoImpact
    Call Taste(KeyCode, 0, Me)
 End Select ' Index
End Sub 'Private Sub vCommandB_KeyPress(Index as Integer, KeyAscii As Integer)


Private Sub Form_Resize()
  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub 'Private Sub Form_Unload()
' vLab(1)        :Pat_id_Bezeichnungsfeld
' vLab(2)        :Nachname_Bezeichnungsfeld
' vLab(3)        :Vorname_Bezeichnungsfeld
' vLab(4)        :ZeitPunkt_Bezeichnungsfeld
' vLab(5)        :DokName_Bezeichnungsfeld
' vLab(6)        :DokGroe_Bezeichnungsfeld
' vLab(7)        :AktZeit_Bezeichnungsfeld
' vLab(8)        :DokPfad-Bezeichnungsfeld
' vTextB(1)      :Pat_id
' vTextB(2)      :Nachname
' vTextB(3)      :Vorname
' vTextB(4)      :ZeitPunkt
' vTextB(5)      :DokName
' vTextB(6)      :DokPfad
' vTextB(7)      :Datensatz
' vTextB(8)      :DatensatzPos
' vTextB(9)      :MerkName
' vTextB(10)     :abgehakt
' vTextB(11)     :AktZeit
' vTextB(12)     :DokGroe
' vCommandB(1)   :Anzeigen
' vCommandB(2)   :PhotoImpact
' vCommandB(3)   :Abhak

'Abhak:                                  vCommandB(3)
'AktZeit:                                vTextB(11)
'AktZeit_Bezeichnungsfeld:               vLab(7)
'Anzeigen:                               vCommandB(1)
'Datensatz:                              vTextB(7)
'DatensatzPos:                           vTextB(8)
'DokGroe:                                vTextB(12)
'DokGroe_Bezeichnungsfeld:               vLab(6)
'DokName:                                vTextB(5)
'DokName_Bezeichnungsfeld:               vLab(5)
'DokPfad-Bezeichnungsfeld:               vLab(8)
'DokPfad:                                vTextB(6)
'MerkName:                               vTextB(9)
'Nachname:                               vTextB(2)
'Nachname_Bezeichnungsfeld:              vLab(2)
'Pat_id:                                 vTextB(1)
'Pat_id_Bezeichnungsfeld:                vLab(1)
'PhotoImpact:                            vCommandB(2)
'Vorname:                                vTextB(3)
'Vorname_Bezeichnungsfeld:               vLab(3)
'ZeitPunkt:                              vTextB(4)
'ZeitPunkt_Bezeichnungsfeld:             vLab(4)
'abgehakt:                               vTextB(10)
