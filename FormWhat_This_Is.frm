VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormWhat_This_Is 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "FormWhat_This_Is.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   3960
      Top             =   600
   End
   Begin VB.PictureBox Picture8 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9540
      Picture         =   "FormWhat_This_Is.frx":7FD8
      ScaleHeight     =   615
      ScaleWidth      =   1635
      TabIndex        =   13
      Top             =   1860
      Width           =   1635
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   7860
      Picture         =   "FormWhat_This_Is.frx":B5AD
      ScaleHeight     =   615
      ScaleWidth      =   1635
      TabIndex        =   12
      Top             =   1860
      Width           =   1635
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6180
      Picture         =   "FormWhat_This_Is.frx":EDD3
      ScaleHeight     =   615
      ScaleWidth      =   1635
      TabIndex        =   11
      Top             =   1860
      Width           =   1635
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9540
      Picture         =   "FormWhat_This_Is.frx":127EA
      ScaleHeight     =   615
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   1140
      Width           =   1635
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   7860
      Picture         =   "FormWhat_This_Is.frx":15EAD
      ScaleHeight     =   615
      ScaleWidth      =   1635
      TabIndex        =   9
      Top             =   1140
      Width           =   1635
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6180
      Picture         =   "FormWhat_This_Is.frx":1909A
      ScaleHeight     =   615
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   1140
      Width           =   1635
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4500
      Picture         =   "FormWhat_This_Is.frx":1C965
      ScaleHeight     =   615
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   1860
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4500
      Picture         =   "FormWhat_This_Is.frx":1FF46
      ScaleHeight     =   615
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   1140
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Back"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7500
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   6060
      Top             =   6180
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Gambar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   780
      TabIndex        =   4
      Top             =   240
      Width           =   2955
      Begin VB.Image gambarkey 
         Height          =   1110
         Left            =   480
         Picture         =   "FormWhat_This_Is.frx":23298
         Top             =   840
         Width           =   1935
      End
      Begin VB.Image gambarmouse 
         Height          =   1605
         Left            =   600
         Picture         =   "FormWhat_This_Is.frx":23E4D
         Top             =   420
         Width           =   1905
      End
      Begin VB.Image gambarmotherboard 
         Height          =   1620
         Left            =   480
         Picture         =   "FormWhat_This_Is.frx":24949
         Top             =   480
         Width           =   2025
      End
      Begin VB.Image gambarcdrom 
         Height          =   1110
         Left            =   480
         Picture         =   "FormWhat_This_Is.frx":261BC
         Top             =   600
         Width           =   2010
      End
      Begin VB.Image gambarfloppy 
         Height          =   1815
         Left            =   480
         Picture         =   "FormWhat_This_Is.frx":26B81
         Top             =   360
         Width           =   1815
      End
      Begin VB.Image gambarpower 
         Height          =   1470
         Left            =   360
         Picture         =   "FormWhat_This_Is.frx":27123
         Top             =   480
         Width           =   2145
      End
      Begin VB.Image gambarharddisk 
         Height          =   1740
         Left            =   360
         Picture         =   "FormWhat_This_Is.frx":27E7F
         Top             =   360
         Width           =   2220
      End
      Begin VB.Image gambarmonitor 
         Height          =   1935
         Left            =   480
         Picture         =   "FormWhat_This_Is.frx":28B21
         Top             =   360
         Width           =   1995
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   7560
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormWhat_This_Is.frx":293D7
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   7440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   7680
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Keluar"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox istilah 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5235
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2760
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mengenal Komponen Dalam Kompoter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   4440
      TabIndex        =   14
      Top             =   420
      Width           =   6795
   End
End
Attribute VB_Name = "FormWhat_This_Is"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Control
Sub gambar()
gambarkey.Visible = False
gambarmouse.Visible = False
gambarmotherboard.Visible = False
gambarcdrom.Visible = False
gambarfloppy.Visible = False
gambarpower.Visible = False
gambarharddisk.Visible = False
gambarmonitor.Visible = False
End Sub
Sub aktif()
For Each x In Me
If TypeName(x) = "Label" Then
    x.Enabled = True
End If
Next
End Sub
Sub pasif()
For Each x In Me
If TypeName(x) = "Label" Then
    x.Enabled = False
End If
Next
End Sub
Sub APA()
With Adodc1
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\trobleshooting\ahli.mdb;Persist Security Info=False"
.CommandType = adCmdText
.RecordSource = "Select * from arti where id like '%" & Text1.Text & "%'order by id asc"
.Refresh
End With
istilah = DataGrid1.Columns(1)
End Sub
Private Sub Command1_Click()
aktif
istilah.Text = ""
gambar
End Sub
Private Sub Command2_Click()
a = MsgBox("Terima Kasih Telah Menggunakan Form Ini", vbExclamation + vbOKOnly, "Ucapan!!")
Unload Me
FormMenuUtama.Show
End Sub
Private Sub Form_Activate()
gambar
End Sub
Private Sub Form_Load()
Adodc1.Visible = False
DataGrid1.Visible = False
Text1.Visible = False
End Sub
Private Sub Label2_Click()
pasif
Label2.Enabled = True
Text1.Text = "02"
APA
gambar
gambarmouse.Visible = True
End Sub
Private Sub Label3_Click()

End Sub
Private Sub Label4_Click()
pasif
Label4.Enabled = True
Text1.Text = "04"
APA
gambar
gambarcdrom.Visible = True
End Sub
Private Sub Label5_Click()

End Sub
Private Sub Label6_Click()

End Sub
Private Sub Label7_Click()
pasif
Label7.Enabled = True
Text1.Text = "06"
APA
gambar
gambarharddisk.Visible = True
End Sub
Private Sub Label8_Click()

End Sub
Private Sub Picture1_Click()
pasif
'Label1.Enabled = True
Text1.Text = "01"
APA
gambar
gambarkey.Visible = True
End Sub
Private Sub Picture2_Click()
pasif
'Label2.Enabled = True
Text1.Text = "02"
APA
gambar
gambarmouse.Visible = True
End Sub
Private Sub Picture3_Click()
pasif
'Label3.Enabled = True
Text1.Text = "03"
APA
gambar
gambarmotherboard.Visible = True
End Sub

Private Sub Picture4_Click()
pasif
'Label4.Enabled = True
Text1.Text = "04"
APA
gambar
gambarcdrom.Visible = True
End Sub

Private Sub Picture5_Click()
pasif
'Label5.Enabled = True
Text1.Text = "05"
APA
gambar
gambarfloppy.Visible = True
End Sub

Private Sub Picture6_Click()
pasif
'Label6.Enabled = True
Text1.Text = "07"
APA
gambar
gambarpower.Visible = True
End Sub

Private Sub Picture7_Click()
pasif
'Label7.Enabled = True
Text1.Text = "06"
APA
gambar
gambarharddisk.Visible = True
End Sub

Private Sub Picture8_Click()
pasif
'Label8.Enabled = True
Text1.Text = "08"
APA
gambar
gambarmonitor.Visible = True
End Sub

Private Sub Timer1_Timer()
If Frame1.Visible = True Then
Frame1.Visible = False
Else
Frame1.Visible = True
End If
End Sub

Private Sub Timer2_Timer()
If Label1.Visible = True Then
Label1.Visible = False
Else
Label1.Visible = True
End If
End Sub
