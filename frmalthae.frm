VERSION 5.00
Begin VB.Form frmalthae 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "althae"
   ClientHeight    =   12570
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmalthae"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12570
   ScaleWidth      =   5775
   Begin VB.TextBox pat_idFeld 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2040
      TabIndex        =   87
      Top             =   11610
      Width           =   1335
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   78
      Top             =   11970
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Abbrechen"
         Height          =   300
         Left            =   1213
         TabIndex        =   85
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Aktualisieren"
         Height          =   300
         Left            =   59
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "S&chließen"
         Height          =   300
         Left            =   4675
         TabIndex        =   83
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Neu lesen"
         Height          =   300
         Left            =   3521
         TabIndex        =   82
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Löschen"
         Height          =   300
         Left            =   2367
         TabIndex        =   81
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Bearbeiten"
         Height          =   300
         Left            =   1213
         TabIndex        =   80
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Hinzufügen"
         Height          =   300
         Left            =   59
         TabIndex        =   79
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
      ScaleWidth      =   5775
      TabIndex        =   72
      Top             =   12270
      Width           =   5775
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmalthae.frx":0000
         Style           =   1  'Grafisch
         TabIndex        =   76
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmalthae.frx":0342
         Style           =   1  'Grafisch
         TabIndex        =   75
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmalthae.frx":0684
         Style           =   1  'Grafisch
         TabIndex        =   74
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmalthae.frx":09C6
         Style           =   1  'Grafisch
         TabIndex        =   73
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   690
         TabIndex        =   77
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AktZeit"
      Height          =   285
      Index           =   35
      Left            =   2040
      TabIndex        =   71
      Top             =   11260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "bis"
      Height          =   285
      Index           =   34
      Left            =   2040
      TabIndex        =   69
      Top             =   10940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "seit"
      Height          =   285
      Index           =   33
      Left            =   2040
      TabIndex        =   67
      Top             =   10620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "gelöscht"
      Height          =   285
      Index           =   32
      Left            =   2040
      TabIndex        =   65
      Top             =   10300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PLZ"
      Height          =   285
      Index           =   31
      Left            =   2040
      TabIndex        =   63
      Top             =   9980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Straße"
      Height          =   285
      Index           =   30
      Left            =   2040
      TabIndex        =   61
      Top             =   9660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Nachname"
      Height          =   285
      Index           =   29
      Left            =   2040
      TabIndex        =   59
      Top             =   9340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Vorname"
      Height          =   285
      Index           =   28
      Left            =   2040
      TabIndex        =   57
      Top             =   9020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Titel"
      Height          =   285
      Index           =   27
      Left            =   2040
      TabIndex        =   55
      Top             =   8700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Geschlecht"
      Height          =   285
      Index           =   26
      Left            =   2040
      TabIndex        =   53
      Top             =   8380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DMPT1"
      Height          =   285
      Index           =   25
      Left            =   2040
      TabIndex        =   51
      Top             =   8060
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DMPT2"
      Height          =   285
      Index           =   24
      Left            =   2040
      TabIndex        =   49
      Top             =   7740
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "beme"
      Height          =   285
      Index           =   23
      Left            =   2040
      TabIndex        =   47
      Top             =   7420
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GemMit"
      Height          =   285
      Index           =   22
      Left            =   2040
      TabIndex        =   45
      Top             =   7100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Arzttyp"
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   43
      Top             =   6780
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ZulG"
      Height          =   285
      Index           =   20
      Left            =   2040
      TabIndex        =   41
      Top             =   6460
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Email"
      Height          =   285
      Index           =   19
      Left            =   2040
      TabIndex        =   39
      Top             =   6140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax3k"
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   37
      Top             =   5820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax3"
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   35
      Top             =   5500
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax2k"
      Height          =   285
      Index           =   16
      Left            =   2040
      TabIndex        =   33
      Top             =   5180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax2"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   31
      Top             =   4860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax1k"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   29
      Top             =   4540
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax1"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   27
      Top             =   4220
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel4"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   25
      Top             =   3900
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel3"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Top             =   3580
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel2"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   3260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel1"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   2940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LANR"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "KVNu"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "KVNR"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ort"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "HAName"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Anrede"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BStelle"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "diff"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DBNr"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Beispiel-Pat_id:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   86
      Top             =   11640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "AktZeit:"
      Height          =   255
      Index           =   35
      Left            =   120
      TabIndex        =   70
      Top             =   11260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "bis:"
      Height          =   255
      Index           =   34
      Left            =   120
      TabIndex        =   68
      Top             =   10940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "seit:"
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   66
      Top             =   10620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "gelöscht:"
      Height          =   255
      Index           =   32
      Left            =   120
      TabIndex        =   64
      Top             =   10300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PLZ:"
      Height          =   255
      Index           =   31
      Left            =   120
      TabIndex        =   62
      Top             =   9980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Straße:"
      Height          =   255
      Index           =   30
      Left            =   120
      TabIndex        =   60
      Top             =   9660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nachname:"
      Height          =   255
      Index           =   29
      Left            =   120
      TabIndex        =   58
      Top             =   9340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vorname:"
      Height          =   255
      Index           =   28
      Left            =   120
      TabIndex        =   56
      Top             =   9020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Titel:"
      Height          =   255
      Index           =   27
      Left            =   120
      TabIndex        =   54
      Top             =   8700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Geschlecht:"
      Height          =   255
      Index           =   26
      Left            =   120
      TabIndex        =   52
      Top             =   8380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DMPT1:"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   50
      Top             =   8060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DMPT2:"
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   48
      Top             =   7740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "beme:"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   46
      Top             =   7420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "GemMit:"
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   44
      Top             =   7100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Arzttyp:"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   42
      Top             =   6780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ZulG:"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   40
      Top             =   6460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Email:"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   38
      Top             =   6140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fax3k:"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   36
      Top             =   5820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fax3:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   5500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fax2k:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Top             =   5180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fax2:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   30
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fax1k:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   4540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fax1:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tel4:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tel3:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tel2:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tel1:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LANR:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "KVNu:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "KVNR:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ort:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "HAName:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Anrede:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BStelle:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "diff:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DBNr:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmalthae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

' nur in HausärztemitalterKVNrergänzen_Click
Public Function Vorbereit(Umfang$)
  Set adoPrimaryRS = New Recordset
'  adoPrimaryRS.Open "SELECT DBNr,diff,BStelle,Anrede,HAName,ort,KVNR,KVNu,LANR,Tel1,Tel2,Tel3,Tel4,Fax1,Fax1k,Fax2,Fax2k,Fax3,Fax3k,Email,ZulG,Arzttyp,GemMit,beme,DMPT2,DMPT1,Geschlecht,Titel,Vorname,Nachname,Straße,PLZ,gelöscht,seit,bis,AktZeit FROM althae ORDER BY kvnu", DBCn, adOpenStatic, adLockOptimistic
'  DBCn.CursorLocation = adUseClient
  If Right$(Umfang, 1) = "," Then Umfang = Left$(Umfang, Len(Umfang) - 1)
  adoPrimaryRS.Open "SELECT neuAI,DBNr,diff,BStelle,Anrede,HAName,ort,KVNR,KVNu,LANR,Tel1,Tel2,Tel3,Tel4,Fax1,Fax1k,Fax2,Fax2k,Fax3,Fax3k,Email,ZulG,Arzttyp,GemMit,beme,DMPT2,DMPT1,Geschlecht,Titel,Vorname,Nachname,Straße,PLZ,gelöscht,seit,bis,AktZeit FROM althae WHERE kvnu IN (" & Umfang & ")", DBCn, adOpenStatic, adLockOptimistic
  If adoPrimaryRS.BOF Then Exit Function
  Dim oText As TextBox
  'Textfelder an Datenprovider binden
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  mbDataChanged = False
  Vorbereit = True
End Function ' Vorbereit

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub ' Form_Resize()

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
End Sub ' Form_KeyDown(KeyCode%, Shift%)

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub ' Form_Unload(Cancel As Integer)

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As Adodb.EventReasonEnum, ByVal pError As Adodb.Error, adStatus As Adodb.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
  Dim rs As New Adodb.Recordset
  'Hierdurch wird die aktuelle Datensatzposition für diese Datensatzgruppe angezeigt
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
  If Not adoPrimaryRS.BOF And Not adoPrimaryRS.EOF Then
   myFrag rs, "SELECT pat_id FROM `namen` WHERE kvnr = '" & adoPrimaryRS!kvnu & "'"
   If Not rs.BOF Then
    Me.pat_idFeld = rs!Pat_id
   End If
  End If
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As Adodb.EventReasonEnum, ByVal cRecords As Long, adStatus As Adodb.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)
  'Hier können Sie Code zur Überprüfung einfügen
  'Dieses Ereignis wird aufgerufen, wenn die folgenden Aktionen eintreten
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
   Me.txtFields(35).Text = CDate(Format(Now, "D.M.YYYY hh:mm:ss"))
'   adoPrimaryRS.Requery
'   adoPrimaryRS.Resync
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
   Me.txtFields(35).Text = CDate(Format(Now, "D.M.YYYY hh:mm:ss"))
'   adoPrimaryRS.Requery
'   adoPrimaryRS.Resync
  End Select
  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Datensatz hinzufügen"
    mbAddNewFlag = True
    SetButtons False
  End With
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub ' cmdAdd_Click()

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
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
  adoPrimaryRS.Requery
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
End Sub ' cmdEdit_Click()

Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
End Sub ' cmdCancel_Click()

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  adoPrimaryRS.UpdateBatch adAffectAll
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'Zu neuem Datensatz gehen
  End If
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub ' cmdUpdate_Click()

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
  adoPrimaryRS.MoveFirst
  mbDataChanged = False
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub ' cmdFirst_Click()

Private Sub cmdLast_Click()
  On Error GoTo GoLastError
  adoPrimaryRS.MoveLast
  mbDataChanged = False
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub ' cmdLast_Click()

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep 1000, 100
     'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoPrimaryRS.MoveLast
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub ' cmdNext_Click()

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError
  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep 1000, 100
    'Ende der Zeile wurde erreicht; zurück zum Zeilenanfang
    adoPrimaryRS.MoveFirst
  End If
  'Aktuellen Datensatz anzeigen
  mbDataChanged = False
  Exit Sub
GoPrevError:
  MsgBox Err.Description
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
End Sub ' SetButtons(bVal As Boolean)

