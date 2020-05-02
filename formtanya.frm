VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FormTroubleshooting 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Pertanyaan"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "formtanya.frx":0000
   ScaleHeight     =   9000
   ScaleMode       =   0  'User
   ScaleWidth      =   12470.15
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture8 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   10020
      Picture         =   "formtanya.frx":D089
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   27
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   8400
      Picture         =   "formtanya.frx":10411
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   26
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   10020
      Picture         =   "formtanya.frx":1396E
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   25
      Top             =   7320
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   8400
      Picture         =   "formtanya.frx":16CE2
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   24
      Top             =   7320
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   1860
      Picture         =   "formtanya.frx":1A108
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   23
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   240
      Picture         =   "formtanya.frx":1D3F7
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   22
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   1860
      Picture         =   "formtanya.frx":209AC
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   21
      Top             =   7320
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   240
      Picture         =   "formtanya.frx":23D50
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   20
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdbatal 
      BackColor       =   &H0000C000&
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdtutup 
      BackColor       =   &H00008080&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   3555
   End
   Begin VB.CommandButton cmdkembali 
      BackColor       =   &H0000C000&
      Caption         =   "&Kembali"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdindikasi 
      BackColor       =   &H0080FF80&
      Caption         =   "&Indikasi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5100
      Width           =   2655
   End
   Begin VB.CommandButton cmdlanjut 
      BackColor       =   &H0000C000&
      Caption         =   "&Lanjut"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdmulai 
      BackColor       =   &H0000C000&
      Caption         =   "&Mulai"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   9120
      TabIndex        =   4
      Top             =   5880
      Width           =   2535
      Begin VB.CommandButton cmdcetak 
         BackColor       =   &H0000C000&
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox solusi 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2595
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2400
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tidak"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   2
      Top             =   5460
      Width           =   1455
   End
   Begin VB.OptionButton optya 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ya"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   5460
      Width           =   975
   End
   Begin VB.TextBox txttny 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "formtanya.frx":26F6C
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1508
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "formtanya.frx":26F81
      Height          =   2550
      Left            =   3360
      TabIndex        =   16
      ToolTipText     =   "Pilih Indikasi yang mungkin terjadi pada komputer Anda...!!!"
      Top             =   2400
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4498
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "TABEL INDIKASI & SOLUSI"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "idsolusi"
         Caption         =   "idsolusi"
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
         DataField       =   "indikasi"
         Caption         =   "indikasi"
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
         DataField       =   "solusi"
         Caption         =   "solusi"
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
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7506,792
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   3720
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9840
      Top             =   4440
      Visible         =   0   'False
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
      CommandType     =   2
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
      RecordSource    =   "tanya"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9840
      Top             =   4080
      Visible         =   0   'False
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
      CommandType     =   2
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
      RecordSource    =   "solusi"
      Caption         =   "Adodc2"
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   9120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   3960
      Width           =   2175
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   600
      Top             =   600
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      BorderWidth     =   8
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   5340
      Width           =   3195
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FORM KONSULTASI MASALAH PADA KOMPUTER ANDA :"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1215
      Left            =   2580
      TabIndex        =   15
      Top             =   180
      Width           =   6735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "SOLUSI :"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9540
      TabIndex        =   14
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PERTANYAAN :"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1860
      Width           =   2835
   End
End
Attribute VB_Name = "FormTroubleshooting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Control
Sub MATI()
For Each x In Me
    If TypeName(x) = "CommandButton" Then
    x.Enabled = False
    End If
Next
optya.Enabled = False
Option1.Enabled = False
cmdtutup.Enabled = True
End Sub
Sub aktif()
For Each x In Me
    If TypeName(x) = "PictureBox" Then
        x.Enabled = True
    End If
Next
End Sub
Sub pasif()
For Each x In Me
    If TypeName(x) = "PictureBox" Then
        x.Enabled = False
    End If
Next
Label17.Enabled = True
Label18.Enabled = True
Label19.Enabled = True
End Sub

Private Sub cmdbatal_Click()
aktif
txttny.Text = ""
solusi.Text = ""
optya.Value = False
Option1.Value = False
DataGrid1.Visible = False
MATI
Adodc2.Refresh
End Sub
Private Sub cmdcetak_Click()
frmlayar.Show
frmlayar.FontName = "monotype corsiva"
frmlayar.CurrentX = 0
frmlayar.CurrentY = 0
frmlayar.Print ""
frmlayar.Label1.ForeColor = vbBlue
frmlayar.ForeColor = vbBlue
frmlayar.FontSize = 33
frmlayar.FontBold = True
frmlayar.Print Tab(12); "SISTEM PAKAR"
frmlayar.FontUnderline = True
frmlayar.FontSize = 24
frmlayar.Print Tab(12); "       Kategori TroubleShooting       "
frmlayar.FontSize = 18
frmlayar.FontBold = True
frmlayar.FontUnderline = False
frmlayar.Print String(115, "=")
frmlayar.Print Tab(2); "Masalah Anda :"; Tab(20); txttny.Text
frmlayar.Print String(115, "=")
frmlayar.FontSize = 16
frmlayar.Print Tab(2); "Indikasi :"; Tab(15); Text3.Text
frmlayar.Print ""
frmlayar.Print Tab(2); "Solusi Dari Kami :";
frmlayar.Label1.FontName = "Monotype Corsiva"
frmlayar.Label1.FontSize = 20
frmlayar.Label1.Caption = FormTroubleshooting.solusi.Text
End Sub

Private Sub cmdindikasi_Click()
DataGrid1.Visible = True
If Picture1.Enabled = True Then
    If Text1.Text = "001KEY" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='KEY-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "002KEY" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='KEY-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "003KEY" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='KEY-3'"
    Adodc1.Refresh
    ElseIf Text1.Text = "004KEY" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='KEY-4'"
    Adodc1.Refresh
    End If
ElseIf Picture2.Enabled = True Then
    If Text1.Text = "005MOS" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MOS-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "006MOS" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MOS-2'"
    Adodc1.Refresh
    ElseIf Text1.Text = "007MOS" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MOS-3'"
    Adodc1.Refresh
    ElseIf Text1.Text = "008MOS" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MOS-4'"
    Adodc1.Refresh
    End If
ElseIf Picture3.Enabled = True Then
    If Text1.Text = "009MTH" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTH-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "010MTH" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTH-2'"
    Adodc1.Refresh
    ElseIf Text1.Text = "011MTH" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTH-3'"
    Adodc1.Refresh
    ElseIf Text1.Text = "012MTH" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTH-4'"
    Adodc1.Refresh
    ElseIf Text1.Text = "013MTH" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTH-5'"
    Adodc1.Refresh
    ElseIf Text1.Text = "014MTH" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTH-6'"
    Adodc1.Refresh
    End If
ElseIf Picture4.Enabled = True Then
    If Text1.Text = "015CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "016CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-2'"
    Adodc1.Refresh
    ElseIf Text1.Text = "017CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-3'"
    Adodc1.Refresh
    ElseIf Text1.Text = "018CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-4'"
    Adodc1.Refresh
    ElseIf Text1.Text = "019CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-5'"
    Adodc1.Refresh
    ElseIf Text1.Text = "020CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-6'"
    Adodc1.Refresh
    ElseIf Text1.Text = "021CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-7'"
    Adodc1.Refresh
    ElseIf Text1.Text = "022CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-8'"
    Adodc1.Refresh
    ElseIf Text1.Text = "023CDR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='CDR-9'"
    Adodc1.Refresh
    End If
ElseIf Picture5.Enabled = True Then
    If Text1.Text = "024FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "025FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-2'"
    Adodc1.Refresh
    ElseIf Text1.Text = "026FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-3'"
    Adodc1.Refresh
    ElseIf Text1.Text = "027FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-4'"
    Adodc1.Refresh
    ElseIf Text1.Text = "028FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-5'"
    Adodc1.Refresh
    ElseIf Text1.Text = "029FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-6'"
    Adodc1.Refresh
    ElseIf Text1.Text = "030FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-7'"
    Adodc1.Refresh
    ElseIf Text1.Text = "031FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-8'"
    Adodc1.Refresh
    ElseIf Text1.Text = "032FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='FDS-9'"
    Adodc1.Refresh
    ElseIf Text1.Text = "033FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,6)='FDS-10'"
    Adodc1.Refresh
    ElseIf Text1.Text = "034FDD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,6)='FDS-11'"
    Adodc1.Refresh
    End If
ElseIf Picture6.Enabled = True Then
    If Text1.Text = "039FAN" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='PWS-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "040FAN" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='PWS-2'"
    Adodc1.Refresh
    ElseIf Text1.Text = "041FAN" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='PWS-3'"
    Adodc1.Refresh
    ElseIf Text1.Text = "042FAN" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='PWS-4'"
    Adodc1.Refresh
    End If
ElseIf Picture7.Enabled = True Then
    If Text1.Text = "035HRD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='HDD-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "036HRD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='HDD-2'"
    Adodc1.Refresh
    ElseIf Text1.Text = "037HRD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='HDD-3'"
    Adodc1.Refresh
    ElseIf Text1.Text = "038HRD" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='HDD-4'"
    Adodc1.Refresh
    End If
ElseIf Picture8.Enabled = True Then
     If Text1.Text = "043MTR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTR-1'"
    Adodc1.Refresh
    ElseIf Text1.Text = "044MTR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTR-2'"
    Adodc1.Refresh
    ElseIf Text1.Text = "045MTR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTR-3'"
    Adodc1.Refresh
    ElseIf Text1.Text = "06MTR" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from solusi where left(idsolusi,5)='MTR-4'"
    Adodc1.Refresh
    End If
End If
solusi.Height = 2595
optya.Visible = False
Option1.Visible = False
cmdindikasi.Enabled = False
End Sub
Private Sub cmdkembali_Click()
Adodc2.Recordset.MovePrevious
Text1.Text = Adodc2.Recordset.Fields("idtanya")
Option1.Value = False
optya.Value = False
cmdlanjut.Enabled = False
optya.Visible = True
Option1.Visible = True
Shape9.Visible = True
solusi.Height = 2595
solusi.Text = ""
If Picture1.Enabled = True Then
If Text1.Text = "001KEY" Then
    aktif
    Adodc2.Refresh
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture2.Enabled = True Then
If Text1.Text = "005MOS" Then
    aktif
    Adodc2.Refresh
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture3.Enabled = True Then
If Text1.Text = "009MTH" Then
    aktif
    Adodc2.Refresh
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture4.Enabled = True Then
If Text1.Text = "015CDR" Then
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture5.Enabled = True Then
If Text1.Text = "024FDD" Then
    aktif
    Adodc2.Refresh
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture7.Enabled = True Then
If Text1.Text = "035HRD" Then
    aktif
    Adodc2.Refresh
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture6.Enabled = True Then
If Text1.Text = "039FAN" Then
    aktif
    Adodc2.Refresh
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture8.Enabled = True Then
If Text1.Text = "043MTR" Then
    aktif
    Adodc2.Refresh
    cmdkembali.Enabled = False
    Exit Sub
End If
End If
End Sub
Private Sub cmdlanjut_Click()
Adodc2.Recordset.MoveNext
Text1.Text = Adodc2.Recordset.Fields("idtanya")
Option1.Value = False
optya.Value = False
cmdlanjut.Enabled = False
cmdcetak.Enabled = False
optya.Visible = True
Option1.Visible = True
solusi.Height = 2595
solusi.Text = ""
cmdkembali.Enabled = True

End Sub
Private Sub cmdmulai_Click()
With Adodc2.Recordset
    .Find "idtanya='" & Text1.Text & "'"
txttny.Text = .Fields("tanya")
 End With
 Option1.Value = False
 cmdmulai.Enabled = False
optya.Visible = True
Option1.Visible = True
optya.Enabled = True
Option1.Enabled = True
End Sub
Private Sub cmdtutup_Click()
pesan = MsgBox("Semoga Komponen Dalam PC anda Tidak Ada Masalah!!", vbCritical + vbOKOnly, "Informasi..")
Unload Me
FormMenuUtama.Show
End Sub
Private Sub DataGrid1_DblClick()
MsgBox "Berikut ini Adalah Solusi dari Kami...!!!", 32, "Solusi"
If Adodc1.Recordset.RecordCount > 0 Then
  solusi.Text = Adodc1.Recordset.Fields("solusi")
  Text2.Text = Adodc1.Recordset.Fields("idsolusi")
  Text3.Text = Adodc1.Recordset.Fields("indikasi")
End If
DataGrid1.Visible = False
solusi.Height = 2595
cmdcetak.Enabled = True
cmdlanjut.Enabled = True
If Text1.Text = "004KEY" Then
    cmdlanjut.Enabled = False
    cmdkembali.Enabled = True
ElseIf Text1.Text = "008MOS" Then
    cmdlanjut.Enabled = False
    cmdkembali.Enabled = True
ElseIf Text1.Text = "014MTH" Then
    cmdlanjut.Enabled = False
    cmdkembali.Enabled = True
ElseIf Text1.Text = "019CDR" Then
    cmdlanjut.Enabled = False
    cmdkembali.Enabled = True
ElseIf Text1.Text = "034FDD" Then
    cmdlanjut.Enabled = False
    cmdkembali.Enabled = True
ElseIf Text1.Text = "038HRD" Then
    cmdlanjut.Enabled = False
    cmdkembali.Enabled = True
ElseIf Text1.Text = "042FAN" Then
    cmdlanjut.Enabled = False
    cmdkembali.Enabled = True
ElseIf Text1.Text = "046MTR" Then
    cmdlanjut.Enabled = False
    cmdkembali.Enabled = True
End If
End Sub
Private Sub Form_Load()
MATI
DataGrid1.Visible = False

End Sub
Private Sub Option1_Click()
cmdlanjut.Enabled = True
cmdindikasi.Enabled = False
If Picture1.Enabled = True Then
If Text1.Text = "004KEY" Then
    'MsgBox "KeyBoard Anda Tidak Bermasalah...!!!", 48, "STOP"
    cmdlanjut.Enabled = False
    Text1.Text = ""
    txttny.Text = ""
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture2.Enabled = True Then
If Text1.Text = "008MOS" Then
    'MsgBox "Mouse Anda Tidak Bermasalah...!!!", 48, "STOP"
    cmdlanjut.Enabled = False
    Text1.Text = ""
    txttny.Text = ""
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture3.Enabled = True Then
If Text1.Text = "014MTH" Then
    'MsgBox "MotheBoard Anda Tidak Bermasalah...!!!", 48, "STOP"
    cmdlanjut.Enabled = False
    Text1.Text = ""
    txttny.Text = ""
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture4.Enabled = True Then
If Text1.Text = "019CDR" Then
    'MsgBox "CD Room Anda Tidak Bermasalah...!!!", 48, "STOP"
    cmdlanjut.Enabled = False
    Text1.Text = ""
    txttny.Text = ""
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture5.Enabled = True Then
If Text1.Text = "034FDD" Then
    'MsgBox "Floppy Disk Anda Tidak Bermasalah...!!!", 48, "STOP"
    cmdlanjut.Enabled = False
    Text1.Text = ""
    txttny.Text = ""
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture7.Enabled = True Then
If Text1.Text = "038HRD" Then
    'MsgBox "HardDisk Anda Tidak Bermasalah...!!!", 48, "STOP"
    cmdlanjut.Enabled = False
    Text1.Text = ""
    txttny.Text = ""
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture6.Enabled = True Then
If Text1.Text = "042FAN" Then
    'MsgBox "Power Supply Anda Tidak Bermasalah...!!!", 48, "STOP"
    cmdlanjut.Enabled = False
    Text1.Text = ""
    txttny.Text = ""
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
ElseIf Picture8.Enabled = True Then
If Text1.Text = "046MTR" Then
    'MsgBox "Monitor Anda Tidak Bermasalah...!!!", 48, "STOP"
    cmdlanjut.Enabled = False
    Text1.Text = ""
    txttny.Text = ""
    aktif
    Adodc2.Refresh
    optya.Enabled = False
    Option1.Enabled = False
    cmdkembali.Enabled = False
    Exit Sub
End If
End If
End Sub
Private Sub optya_Click()
cmdindikasi.Enabled = True
cmdlanjut.Enabled = False
cmdkembali.Enabled = False

End Sub
Private Sub Picture1_Click()
pasif
Picture1.Enabled = True
Text1.Text = Adodc2.Recordset.Fields("idtanya")
cmdmulai.Enabled = True
cmdbatal.Enabled = True
txttny.Text = ""
End Sub
Private Sub Picture2_Click()
Adodc2.Recordset.GetRows (4)
pasif
Picture2.Enabled = True
Text1.Text = Adodc2.Recordset.Fields("idtanya")
cmdmulai.Enabled = True
cmdbatal.Enabled = True
txttny.Text = ""
End Sub

Private Sub Picture3_Click()
Adodc2.Recordset.GetRows (8)
pasif
Picture3.Enabled = True
Text1.Text = Adodc2.Recordset.Fields("idtanya")
cmdmulai.Enabled = True
cmdbatal.Enabled = True
txttny.Text = ""
End Sub
Private Sub Picture4_Click()
Adodc2.Recordset.GetRows (14)
pasif
Picture4.Enabled = True
Text1.Text = Adodc2.Recordset.Fields("idtanya")
cmdmulai.Enabled = True
cmdbatal.Enabled = True
txttny.Text = ""
End Sub
Private Sub Picture5_Click()
Adodc2.Recordset.GetRows (23)
pasif
Picture5.Enabled = True
Text1.Text = Adodc2.Recordset.Fields("idtanya")
cmdmulai.Enabled = True
cmdbatal.Enabled = True
txttny.Text = ""
End Sub
Private Sub Picture6_Click()
Adodc2.Recordset.GetRows (38)
pasif
Picture6.Enabled = True
Text1.Text = Adodc2.Recordset.Fields("idtanya")
cmdmulai.Enabled = True
cmdbatal.Enabled = True
txttny.Text = ""
End Sub
Private Sub Picture7_Click()
Adodc2.Recordset.GetRows (34)
pasif
Picture7.Enabled = True
Text1.Text = Adodc2.Recordset.Fields("idtanya")
cmdmulai.Enabled = True
cmdbatal.Enabled = True
txttny.Text = ""
End Sub
Private Sub Picture8_Click()
Adodc2.Recordset.GetRows (42)
pasif
Picture8.Enabled = True
Text1.Text = Adodc2.Recordset.Fields("idtanya")
cmdmulai.Enabled = True
cmdbatal.Enabled = True
txttny.Text = ""
End Sub
Private Sub Text1_Change()
If Len(Text1.Text) = 0 Then Exit Sub
If Len(Text1.Text) < 7 Then
Adodc2.Recordset.Find "idtanya='" & Text1.Text & "'"
txttny.Text = Adodc2.Recordset.Fields("tanya")
Else
Exit Sub
End If
End Sub
Private Sub txttny_Change()
If txttny.Text = "" Then
    Option1.Value = False
End If
End Sub
