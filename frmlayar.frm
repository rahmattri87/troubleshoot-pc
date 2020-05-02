VERSION 5.00
Begin VB.Form frmlayar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   9735
      Left            =   0
      Top             =   0
      Width           =   10095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   3
      Height          =   5415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   9855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   4440
      Width           =   9255
   End
End
Attribute VB_Name = "frmlayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormTroubleshooting.DataGrid1.Visible = False
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
