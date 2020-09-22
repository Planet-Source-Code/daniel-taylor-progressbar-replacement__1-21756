VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin PBar_Ctl.PBar PBar4 
      Height          =   495
      Left            =   360
      Top             =   1560
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      ShowPercent     =   0   'False
      TextColor       =   0
      Value           =   37
      Max             =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReverseColor    =   0   'False
      BarEnd          =   1
   End
   Begin PBar_Ctl.PBar PBar3 
      Height          =   375
      Left            =   360
      Top             =   1080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      ShowPercent     =   0   'False
      ForeColor       =   8421631
      TextColor       =   16711680
      Value           =   450
      Max             =   500
      BorderStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReverseColor    =   0   'False
      BarLength       =   10
      Text            =   "BarLength = 10"
      TextAlign       =   2
   End
   Begin PBar_Ctl.PBar PBar2 
      Height          =   375
      Left            =   360
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      ShowPercent     =   0   'False
      ForeColor       =   65280
      TextColor       =   0
      Value           =   350
      Max             =   1000
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Working..."
      TextAlign       =   0
      BarEnd          =   2
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   3960
      Top             =   2520
   End
   Begin PBar_Ctl.PBar PBar1 
      Height          =   375
      Left            =   360
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      TextColor       =   192
      Value           =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Loading... "
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PBar1_MaxReached()
    PBar1.Value = 0
End Sub

Private Sub PBar2_MaxReached()
    PBar2.Value = 0
End Sub

Private Sub PBar3_MaxReached()
    PBar3.Value = 0
End Sub

Private Sub PBar4_MaxReached()
    PBar4.Value = 0
End Sub

Private Sub Timer1_Timer()
    PBar1.Value = PBar1.Value + 1
    PBar2.Value = PBar2.Value + 1
    PBar3.Value = PBar3.Value + 1
    PBar4.Value = PBar4.Value + 1
End Sub
