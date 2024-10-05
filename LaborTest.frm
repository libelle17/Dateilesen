VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Begin VB.Form LaborTest 
   Caption         =   "dokumente"
   ClientHeight    =   10470
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   14025
   KeyPreview      =   -1  'True
   LinkTopic       =   "LaborTest"
   ScaleHeight     =   10470
   ScaleWidth      =   14025
   Begin MSDataRepeaterLib.DataRepeater DataRepeater1 
      Height          =   1410
      Left            =   0
      TabIndex        =   14
      Top             =   330
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   2487
      _StreamID       =   -1412567295
      _Version        =   393216
      ScrollBars      =   3
      RowDividerStyle =   0
      Caption         =   "Labordokumente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         Name            =   "DokumenteCtl.CtrlDokumente"
      EndProperty
      RepeaterBindings=   9
      BeginProperty RepeaterBinding(0) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "DokName"
         DataField       =   "DOKNAME"
         Key             =   "DOKNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty RepeaterBinding(1) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Dokgroe"
         DataField       =   "dokgroe"
      EndProperty
      BeginProperty RepeaterBinding(2) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "abgehakt"
         DataField       =   "abgehakt"
      EndProperty
      BeginProperty RepeaterBinding(3) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Aktzeit"
         DataField       =   "aktzeit"
      EndProperty
      BeginProperty RepeaterBinding(4) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "DokPfad"
         DataField       =   "dokpfad"
      EndProperty
      BeginProperty RepeaterBinding(5) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Nachname"
         DataField       =   "nachname"
      EndProperty
      BeginProperty RepeaterBinding(6) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Pat_ID"
         DataField       =   "pat_id"
      EndProperty
      BeginProperty RepeaterBinding(7) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Zeitpunkt"
         DataField       =   "zeitpunkt"
      EndProperty
      BeginProperty RepeaterBinding(8) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Vorname"
         DataField       =   "vorname"
      EndProperty
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   14025
      TabIndex        =   6
      Top             =   9870
      Width           =   14025
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Abbrechen"
         Height          =   300
         Left            =   1213
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Aktualisieren"
         Height          =   300
         Left            =   59
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "S&chließen"
         Height          =   300
         Left            =   4675
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Neu lesen"
         Height          =   300
         Left            =   3521
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Löschen"
         Height          =   300
         Left            =   2367
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Bearbeiten"
         Height          =   300
         Left            =   1213
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Hinzufügen"
         Height          =   300
         Left            =   59
         TabIndex        =   7
         Top             =   0
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
      ScaleWidth      =   14025
      TabIndex        =   0
      Top             =   10170
      Width           =   14025
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "LaborTest.frx":0000
         Style           =   1  'Grafisch
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "LaborTest.frx":0342
         Style           =   1  'Grafisch
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "LaborTest.frx":0684
         Style           =   1  'Grafisch
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "LaborTest.frx":09C6
         Style           =   1  'Grafisch
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
End
Attribute VB_Name = "LaborTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoRS As Recordset
Attribute adoRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Private Declare Function Tüt& Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long)
Public dlg As Dialog

Private Sub Form_Load()
  Dim erg&
'  Dim db As New ADODB.Connection
  On Error GoTo fehler
'  Call doConstrFestleg(2, 0)
  Call aCStr(quelleT, qDtb)
  DBCn.CursorLocation = adUseClient
'  db.CursorLocation = adUseClient
'  db.Open "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=linux1;uid=praxis;pwd=sonne;database=quelle;"
  Set adoRS = New ADODB.Recordset
  erg = InputBox("Bitte Pat_id eingeben:", "Pat_id-Eingabe", "0")
  If erg <> 0 Then
   myFrag adoRS, "SELECT FID,`dokumente`.Pat_ID AS pat_id,`namen`.nachname AS nachname, `namen`.vorname AS vorname, ZeitPunkt,`dokumente`.DokPfad AS dokpfad,DokArt,DokName,Quelldatum, -abgehakt AS abgehakt, `dokumente`.absPos AS abspos,`dokumente`.AktZeit AS aktzeit,DokGroe,QS,QT,`dokumente`.StByte AS stbyte FROM ((`dokumente` LEFT JOIN `namen` ON `dokumente`.pat_id = `namen`.pat_id) LEFT JOIN `dokumente abgehakt` ON `dokumente`.dokpfad = " & "`" & "dokumente abgehakt" & "`" & ".dokpfad) WHERE `dokumente`.pat_id = " & erg & " ORDER BY ZeitPunkt DESC", adOpenStatic, DBCn, adLockOptimistic
  Else
   myFrag adoRS, "SELECT FID,`dokumente`.Pat_ID AS pat_id,`namen`.nachname AS nachname, `namen`.vorname AS vorname, ZeitPunkt,`dokumente`.DokPfad AS dokpfad,DokArt,DokName,Quelldatum, abgehakt, `dokumente`.absPos AS abspos,`dokumente`.AktZeit AS aktzeit,DokGroe,QS,QT,`dokumente`.StByte AS stbyte FROM ((`dokumente` LEFT JOIN `namen` ON `dokumente`.pat_id = `namen`.pat_id) LEFT JOIN `dokumente abgehakt` ON `dokumente`.dokpfad = `dokumente abgehakt`.dokpfad) WHERE abgehakt <> 0 ORDER BY ZeitPunkt DESC", adOpenStatic, DBCn, adLockOptimistic
  End If
'  SET grdDataGrid.DataSource = adoRS
  Set Me.DataRepeater1.DataSource = adoRS
  Me.DataRepeater1.RepeatedControl.PcDokPfad = getDokPfad
  mbDataChanged = False
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in form_load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Form_Load()

Private Sub Form_Resize()
  On Error Resume Next
  'Hierdurch wird die Größe der Tabelle geändert, wenn die Größe des Formulars geändert wird
  grdDataGrid.Height = Me.ScaleHeight - 30 - picButtons.Height - picStatBox.Height
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub ' Form_Resize()

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
End Sub ' Form_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub ' Form_Unload(Cancel As Integer)

Private Sub adoRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hierdurch wird die aktuelle Datensatzposition für diese Datensatzgruppe angezeigt
  lblStatus.Caption = CStr(adoRS.AbsolutePosition)
End Sub ' adoRS_MoveComplete

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
End Sub ' adoRS_WillChangeRecord

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  adoRS.MoveLast
  adoRS.AddNew
  grdDataGrid.SetFocus

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub ' cmdAdd_Click()

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub ' cmdDelete_Click()

Private Sub cmdRefresh_Click()
  'Dies wird nur für Mehrbenutzeranwendungen benötigt
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoRS.Requery
  Set grdDataGrid.DataSource = adoRS

  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub ' cmdRefresh_Click()

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Datensatz bearbeiten"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
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

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoRS.MoveLast              'Zu neuem Datensatz gehen
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoRS.EOF Then adoRS.MoveNext
  If adoRS.EOF And adoRS.RecordCount > 0 Then
    Tüt 1760, 1000
     'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoRS.MoveLast
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoRS.BOF Then adoRS.MovePrevious
  If adoRS.BOF And adoRS.RecordCount > 0 Then
    Tüt 1760, 1000
    'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoRS.MoveFirst
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

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
End Sub

