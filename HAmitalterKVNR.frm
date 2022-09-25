VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form HAmitalterKVNR 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "althae"
   ClientHeight    =   12150
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12150
   ScaleWidth      =   5775
   Begin VB.PictureBox picButtons 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   72
      Top             =   11520
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         Caption         =   "S&chließen"
         Height          =   300
         Left            =   4675
         TabIndex        =   77
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Neu lesen"
         Height          =   300
         Left            =   3521
         TabIndex        =   76
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Löschen"
         Height          =   300
         Left            =   2367
         TabIndex        =   75
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Aktualisieren"
         Height          =   300
         Left            =   1213
         TabIndex        =   74
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Hinzufügen"
         Height          =   300
         Left            =   59
         TabIndex        =   73
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AktZeit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   35
      Left            =   2040
      TabIndex        =   71
      Top             =   11260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "bis"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   34
      Left            =   2040
      TabIndex        =   69
      Top             =   10940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "seit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   33
      Left            =   2040
      TabIndex        =   67
      Top             =   10620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "gelöscht"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   32
      Left            =   2040
      TabIndex        =   65
      Top             =   10300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PLZ"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   31
      Left            =   2040
      TabIndex        =   63
      Top             =   9980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Straße"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   30
      Left            =   2040
      TabIndex        =   61
      Top             =   9660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Nachname"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   29
      Left            =   2040
      TabIndex        =   59
      Top             =   9340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Vorname"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   28
      Left            =   2040
      TabIndex        =   57
      Top             =   9020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Titel"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   27
      Left            =   2040
      TabIndex        =   55
      Top             =   8700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Geschlecht"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   26
      Left            =   2040
      TabIndex        =   53
      Top             =   8380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DMPT1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   25
      Left            =   2040
      TabIndex        =   51
      Top             =   8060
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DMPT2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   24
      Left            =   2040
      TabIndex        =   49
      Top             =   7740
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "beme"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   23
      Left            =   2040
      TabIndex        =   47
      Top             =   7420
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GemMit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   22
      Left            =   2040
      TabIndex        =   45
      Top             =   7100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Arzttyp"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   43
      Top             =   6780
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ZulG"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   20
      Left            =   2040
      TabIndex        =   41
      Top             =   6460
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Email"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   19
      Left            =   2040
      TabIndex        =   39
      Top             =   6140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax3k"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   37
      Top             =   5820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   35
      Top             =   5500
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax2k"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   16
      Left            =   2040
      TabIndex        =   33
      Top             =   5180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   31
      Top             =   4860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax1k"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   29
      Top             =   4540
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   27
      Top             =   4220
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   25
      Top             =   3900
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Top             =   3580
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   3260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   2940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LANR"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "KVNu"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "KVNR"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ort"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "HAName"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Anrede"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BStelle"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "diff"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DBNr"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Unten ausrichten
      Height          =   330
      Left            =   0
      Top             =   11820
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=linux;uid=praxis;pwd=***REMOVED***;database=quelle;"
      OLEDBString     =   "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=linux;uid=praxis;pwd=***REMOVED***;database=quelle;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"HAmitalterKVNR.frx":0000
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
Attribute VB_Name = "HAmitalterKVNR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Public Sub vorbereit()
' datprimaryRS.Recordset.Resync adAffectCurrent
' datprimaryRS.RecordSource = "select DBNr,diff,BStelle,Anrede,HAName,ort,KVNR,KVNu,LANR,Tel1,Tel2,Tel3,Tel4,Fax1,Fax1k,Fax2,Fax2k,Fax3,Fax3k,Email,ZulG,Arzttyp,GemMit,beme,DMPT2,DMPT1,Geschlecht,Titel,Vorname,Nachname,Straße,PLZ,gelöscht,seit,bis,AktZeit from (select n.kvnr nkvnu, HAName hHA, concat_ws(', ',l.name, l.vorname) lHA, h.* from aktfv f left join namen n on f.pat_id = n.pat_id left join listenausgabeuew l on n.kvnr = l.kvnr left join althae h on n.kvnr = h.kvnu group by n.pat_id) innen where (isnull(lha) or lha='') and (isnull(hha) or hha='') and kvnr <> ''"
 datPrimaryRS.ConnectionString = DBCn
 datPrimaryRS.RecordSource = "select DBNr,diff,BStelle,Anrede,HAName,ort,KVNR,KVNu,LANR,Tel1,Tel2,Tel3,Tel4,Fax1,Fax1k,Fax2,Fax2k,Fax3,Fax3k,Email,ZulG,Arzttyp,GemMit,beme,DMPT2,DMPT1,Geschlecht,Titel,Vorname,Nachname,Straße,PLZ,gelöscht,seit,bis,AktZeit from (select n.kvnr nkvnu, HAName hHA, concat_ws(', ',l.name, l.vorname) lHA, h.* from aktfv f left join namen n on f.pat_id = n.pat_id left join listenausgabeuew l on n.kvnr = l.kvnr left join althae h on n.kvnr = h.kvnu group by n.pat_id) innen where (isnull(lha) or lha='') and (isnull(hha) or hha='') and kvnr <> ''"
' Set datprimaryRS.Recordset = New ADODB.Recordset
' Set rs = New ADODB.Recordset
' rs.Open "select DBNr,diff,BStelle,Anrede,HAName,ort,KVNR,KVNu,LANR,Tel1,Tel2,Tel3,Tel4,Fax1,Fax1k,Fax2,Fax2k,Fax3,Fax3k,Email,ZulG,Arzttyp,GemMit,beme,DMPT2,DMPT1,Geschlecht,Titel,Vorname,Nachname,Straße,PLZ,gelöscht,seit,bis,AktZeit from (select n.kvnr nkvnu, HAName hHA, concat_ws(', ',l.name, l.vorname) lHA, h.* from aktfv f left join namen n on f.pat_id = n.pat_id left join listenausgabeuew l on n.kvnr = l.kvnr left join althae h on n.kvnr = h.kvnu group by n.pat_id) innen where (isnull(lha) or lha='') and (isnull(hha) or hha='') and kvnr <> ''", DBCn, adOpenDynamic, adLockOptimistic
' Set datprimaryRS.Recordset = rs
' Set datprimaryRS.Recordset = New ADODB.Recordset
' datprimaryRS.Recordset.Open "select DBNr,diff,BStelle,Anrede,HAName,ort,KVNR,KVNu,LANR,Tel1,Tel2,Tel3,Tel4,Fax1,Fax1k,Fax2,Fax2k,Fax3,Fax3k,Email,ZulG,Arzttyp,GemMit,beme,DMPT2,DMPT1,Geschlecht,Titel,Vorname,Nachname,Straße,PLZ,gelöscht,seit,bis,AktZeit from (select n.kvnr nkvnu, HAName hHA, concat_ws(', ',l.name, l.vorname) lHA, h.* from aktfv f left join namen n on f.pat_id = n.pat_id left join listenausgabeuew l on n.kvnr = l.kvnr left join althae h on n.kvnr = h.kvnu group by n.pat_id) innen where (isnull(lha) or lha='') and (isnull(hha) or hha='') and kvnr <> ''", DBCn, adOpenStatic, adLockReadOnly
' datprimaryRS.Recordset.Open
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datprimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Hier können Sie Ihre Fehlerbehandlungsroutine einfügen
  'Wenn Sie Fehler ignorieren möchten, kommentieren Sie die nächste Zeile aus
  'Wenn Sie die Fehler auffangen möchten, fügen Sie an dieser Stelle Code ein, um Fehler zu behandeln
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datprimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Hierdurch wird die aktuelle Datensatzposition für diese Datensatzgruppe angezeigt
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datprimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Dies wird nur für Mehrbenutzeranwendungen benötigt
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

