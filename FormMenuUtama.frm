VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FormMenuUtama 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormMenuUtama.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   8520
      Picture         =   "FormMenuUtama.frx":10791
      ScaleHeight     =   555
      ScaleWidth      =   1455
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   8520
      Picture         =   "FormMenuUtama.frx":134DA
      ScaleHeight     =   555
      ScaleWidth      =   1455
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   3240
      Picture         =   "FormMenuUtama.frx":16288
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   8
      Top             =   7560
      Width           =   2115
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   8520
      Picture         =   "FormMenuUtama.frx":1A100
      ScaleHeight     =   675
      ScaleWidth      =   2175
      TabIndex        =   7
      Top             =   300
      Width           =   2175
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   6240
      Picture         =   "FormMenuUtama.frx":1D791
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   300
      Width           =   2235
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   3960
      Picture         =   "FormMenuUtama.frx":213D9
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   300
      Width           =   2235
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1680
      Picture         =   "FormMenuUtama.frx":251DD
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   300
      Width           =   2235
   End
   Begin MSComDlg.CommonDialog informasi 
      Left            =   5640
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Orientation     =   2
   End
   Begin VB.Timer Timer6 
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   300
      Left            =   4560
      Top             =   5280
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1860
      Top             =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   3495
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   1980
         Top             =   480
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   $"FormMenuUtama.frx":28BB7
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   6795
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3015
      End
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   840
      Top             =   3480
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Created by muslimah_sulung@yahoo.com {juli2007}"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5820
      TabIndex        =   3
      Top             =   8280
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   1740
      X2              =   11280
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Suatu Sistem Pakar adalah suatu sistem komputer yang menyamai (emulates) kemampuan pengambilan keputusan dari seorang pakar."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1395
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   7095
   End
End
Attribute VB_Name = "FormMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UMI As IAgentCtlCharacterEx
Const DATAPATH = "merlin.acs"
Private Sub Form_Activate()
Picture6.Visible = False
Picture7.Visible = False
Picture4.Visible = True
End Sub
Private Sub Form_Load()
Agent1.Characters.Load "merlin.acs", DATAPATH
    Set UMI = Agent1.Characters("merlin.acs")
    UMI.LanguageID = &H409
    UMI.Show
    UMI.Speak "Helllooo My Name SULUNG"
    UMI.MoveTo 700, 10
    'UMI.MoveTo 0, 10
    'UMI.MoveTo 300, 500
    UMI.Play "suggest"
    UMI.Speak "Selamat Datang di Applikasi ini..."
    'UMI.Play "confused"
    UMI.MoveTo 300, 200
     UMI.Speak "Sulung akan membantu anda...!!!"
    UMI.Play "read"
    UMI.MoveTo 300, 400
    UMI.Play "search"
    UMI.MoveTo 700, 200
    'UMI.MoveTo 700, 10
    'UMI.Play "acknowledge"

   
    'UMI.Play "read"
    ' ***************************************************
    ' Anda bisa menambahkan perintah lain yang anda
    ' dapatkan dari program MERLIN in Action 1.0 disini.
    ' ***************************************************
   
    'merlin.Hide n
End Sub
Private Sub Picture1_Click()
FormWhat_This_Is.Show
FormMenuUtama.Hide
End Sub

Private Sub Picture2_Click()
FormTroubleshooting.Show
FormMenuUtama.Hide
End Sub

Private Sub Picture3_Click()
FormMaintenance.Show
FormMenuUtama.Hide
End Sub

Private Sub Picture4_Click()
Picture6.Visible = True
Picture7.Visible = True
Picture4.Visible = False
End Sub

Private Sub Picture5_Click()

Unload Me
End Sub
Private Sub Picture6_Click()
informasi.HelpFile = "C:\trobleshooting\HELP\pakar.hlp"
informasi.HelpCommand = cdlHelpContents
informasi.ShowHelp
Picture6.Visible = False
Picture7.Visible = False
Picture4.Visible = True
End Sub
Private Sub Picture7_Click()
informasi.HelpFile = "C:\trobleshooting\HELP\saya.hlp"
informasi.HelpCommand = cdlHelpContents
informasi.ShowHelp
Picture6.Visible = False
Picture7.Visible = False
Picture4.Visible = True
End Sub
Private Sub Timer1_Timer()
If Label5.Top + Label5.Height < 0 Then
Label5.Top = Label5.Height
Else
Label5.Top = Label5.Top - 200
End If
End Sub
Private Sub Timer2_Timer()
If Label6.Visible = True Then
Label6.Visible = False
Label6.ForeColor = RGB(Rnd * 500, Rnd * 500, Rnd * 500)
Else
Label6.Visible = True
Label6.ForeColor = RGB(Rnd * 500, Rnd * 500, Rnd * 500)
End If
End Sub

