VERSION 5.00
Begin VB.Form Form24 
   Caption         =   "Form24"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form24"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Main menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   22
      Top             =   9960
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   7695
      Left            =   720
      Picture         =   "FAZER.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   13875
      TabIndex        =   0
      Top             =   1320
      Width           =   13935
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   16320
      Picture         =   "FAZER.frx":39DA3
      ScaleHeight     =   2715
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture4 
      Height          =   2775
      Left            =   16320
      Picture         =   "FAZER.frx":3ED16
      ScaleHeight     =   2715
      ScaleWidth      =   3915
      TabIndex        =   12
      Top             =   7560
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   16320
      Picture         =   "FAZER.frx":43F0A
      ScaleHeight     =   2715
      ScaleWidth      =   3915
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture10 
      Height          =   4935
      Left            =   12480
      Picture         =   "FAZER.frx":49186
      ScaleHeight     =   4875
      ScaleWidth      =   6195
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture9 
      Height          =   5055
      Left            =   6240
      Picture         =   "FAZER.frx":57868
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture8 
      Height          =   5055
      Left            =   0
      Picture         =   "FAZER.frx":6B9C5
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture7 
      Height          =   5055
      Left            =   12480
      Picture         =   "FAZER.frx":8A7F2
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture6 
      Height          =   5055
      Left            =   6240
      Picture         =   "FAZER.frx":A387B
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture5 
      Height          =   5055
      Left            =   0
      Picture         =   "FAZER.frx":B6F67
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BOOK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   6
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Height          =   855
      Left            =   7800
      Picture         =   "FAZER.frx":CE079
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9840
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Height          =   855
      Left            =   5760
      Picture         =   "FAZER.frx":CE79A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   3600
      Picture         =   "FAZER.frx":CEEDF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   1560
      Picture         =   "FAZER.frx":CF5E3
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9840
      Width           =   2055
   End
   Begin VB.PictureBox picture11 
      Height          =   8895
      Left            =   3480
      Picture         =   "FAZER.frx":CFFD7
      ScaleHeight     =   8835
      ScaleWidth      =   12795
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton Command6 
         Caption         =   "RETURN"
         Height          =   615
         Left            =   5160
         TabIndex        =   21
         Top             =   8280
         Width           =   2295
      End
   End
   Begin VB.Label Label5 
      Caption         =   "TERRAIN WHITE"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   13
      Top             =   7080
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "RAVINE RED"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "WILDERNESS BLACK"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14640
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   $"FAZER.frx":EE98F
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   14760
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   $"FAZER.frx":EF02B
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   14760
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Visible = True
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
Picture10.Visible = False
End Sub

Private Sub Command2_Click()
Label2.Visible = True
Label3.Visible = False
Label4.Visible = False
Label4.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
Picture10.Visible = False
End Sub

Private Sub Command3_Click()
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
Picture10.Visible = True
Label1.Visible = False
Label2.Visible = False
End Sub

Private Sub Command4_Click()
Picture1.Visible = True
Label1.Visible = False
Label2.Visible = False
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
Picture10.Visible = False
End Sub

Private Sub Command5_Click()
Form17.Show
End Sub

Private Sub Command6_Click()
picture11.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Picture1.Visible = True
End Sub

Private Sub Command7_Click()
Form28.Show
End Sub

Private Sub Picture2_Click()
picture11.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
Picture10.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
End Sub
