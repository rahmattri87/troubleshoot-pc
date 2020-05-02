VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormMaintenance 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   Picture         =   "FormMaintenance.frx":0000
   ScaleHeight     =   8940
   ScaleMode       =   0  'User
   ScaleWidth      =   26658.95
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5940
      Picture         =   "FormMaintenance.frx":C304
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   8160
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5940
      Picture         =   "FormMaintenance.frx":F5CD
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   7500
      Width           =   1515
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   4080
      Top             =   1980
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormMaintenance.frx":128F3
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   8520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "IDRAWAT"
         Caption         =   "IDRAWAT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NM"
         Caption         =   "NM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "RAWAT"
         Caption         =   "RAWAT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1735,21
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3870,175
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3870,175
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8220
      Top             =   8460
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\trobleshooting\ahli.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\trobleshooting\ahli.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from rawat"
      Caption         =   "Adodc1"
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
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6840
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2475
      Left            =   7440
      TabIndex        =   3
      Top             =   180
      Width           =   4275
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   $"FormMaintenance.frx":12908
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5175
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.TextBox rwt 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6195
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2700
      Width           =   3975
   End
   Begin MSDataListLib.DataCombo combojns 
      Bindings        =   "FormMaintenance.frx":12A60
      DataSource      =   "Adodc1"
      Height          =   465
      Left            =   4680
      TabIndex        =   1
      Top             =   1140
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   820
      _Version        =   393216
      ForeColor       =   49152
      ListField       =   "NM"
      BoundColumn     =   "NM"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Form Tips Perawatan Komponen-Komponen Komputer Anda"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   360
      TabIndex        =   6
      Top             =   300
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Komponen"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   660
      Width           =   2115
   End
End
Attribute VB_Name = "FormMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UMI As IAgentCtlCharacterEx
Const DATAPATH = "merlin.acs"
Private Sub combojns_Change()
With Adodc1.Recordset
    .Find "nm='" & combojns & "'"
    rwt.Text = .Fields("rawat")
    Adodc1.Refresh
End With
End Sub
Private Sub Command1_Click()
FormMenuUtama.Show
FormMaintenance.Hide
End Sub
Private Sub Form_Activate()
Adodc1.Visible = False
DataGrid1.Visible = False
End Sub

Private Sub Picture1_Click()
frmlayar.Show
frmlayar.FontName = "comic sans MS"
frmlayar.CurrentX = 0
frmlayar.CurrentY = 0
frmlayar.Print ""
frmlayar.Label1.ForeColor = vbBlue
frmlayar.ForeColor = vbBlue
frmlayar.FontSize = 33
frmlayar.FontBold = True
frmlayar.FontUnderline = True
frmlayar.Print Tab(8); "SISTEM PAKAR"
'frmlayar.FontBold = False
frmlayar.FontUnderline = False
frmlayar.FontSize = 24
frmlayar.Print Tab(12); "Kategori Perawatan "
frmlayar.FontSize = 16
frmlayar.FontBold = True
frmlayar.Print String(115, "=")
frmlayar.Print Tab(2); "Komponen :"; Tab(15); combojns.Text
frmlayar.Print String(115, "=")
frmlayar.FontSize = 20
frmlayar.Print Tab(2); "CARA PERAWATAN :";
frmlayar.Label1.Caption = FormMaintenance.rwt.Text
End Sub

Private Sub Picture2_Click()
x = MsgBox("Jangan lupa selalu menjaga KEBERSIHAN setiap komponen-komponen PC anda!!!", vbInformation + vbOKOnly, "Himbauan..!!")
Unload Me
FormMenuUtama.Show
End Sub

Private Sub Timer1_Timer()
If Label2.Top + Label2.Height < 0 Then
Label2.Top = Label2.Height
Else
Label2.Top = Label2.Top - 200
End If
End Sub

Private Sub Timer2_Timer()
If Label3.Visible = True Then
Label3.Visible = False
Label3.ForeColor = RGB(Rnd * 500, Rnd * 500, Rnd * 500)
Else
Label3.Visible = True
Label3.ForeColor = RGB(Rnd * 500, Rnd * 500, Rnd * 500)
End If
End Sub


