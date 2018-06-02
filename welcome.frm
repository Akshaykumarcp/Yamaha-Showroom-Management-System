VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   -120
      Picture         =   "welcome.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   10
      Top             =   0
      Width           =   2535
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   360
      Top             =   6360
   End
   Begin VB.FileListBox File1 
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9600
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Text            =   "c:\welcome.mp3"
      Top             =   7800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   360
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Interval        =   102
      Left            =   360
      Top             =   5040
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   9240
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   102
      Scrolling       =   1
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   8400
      Width           =   3255
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4683
      _cy             =   661
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   18360
      TabIndex        =   5
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   17640
      TabIndex        =   4
      Top             =   8520
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME "
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   19215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   3720
      TabIndex        =   2
      Top             =   8280
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   1920
      TabIndex        =   0
      Top             =   8280
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -1920
      Picture         =   "welcome.frx":0CD2
      Top             =   -3600
      Width           =   24000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim x As Integer
Option Explicit


Private Sub Form_Load()
File1.FileName = App.Path
x = File1.ListCount
WindowsMediaPlayer1.URL = Text1.Text
End Sub


Private Sub Timer1_Timer()
i = ProgressBar1.Value
ProgressBar1.Value = i + 1
Label4.Caption = i
If ProgressBar1.Value = 102 Then
Label4.Caption = i
Unload Me
Form2.Show
End If
End Sub

Private Sub Timer2_Timer()
Label2.Caption = "." + Label2.Caption
If (Len(Label2.Caption) > 5) Then
Label2.Caption = ""
End If
If (i <= x) Then
Label6.Caption = File1.List(i)
i = i + 1
Else
Load Form2
Form2.Show
Unload Me
End If
End Sub

Private Sub Timer3_Timer()
If Label3.Visible = True Then
Label3.Visible = False
Else
Label3.Visible = True
End If
End Sub

