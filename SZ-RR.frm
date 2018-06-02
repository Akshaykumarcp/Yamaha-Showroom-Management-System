VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H80000008&
   Caption         =   "Form20"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form20"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture7 
      Height          =   5055
      Left            =   12480
      Picture         =   "SZ-RR.frx":0000
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   18
      Top             =   0
      Width           =   6255
   End
   Begin VB.PictureBox Picture6 
      Height          =   5055
      Left            =   6240
      Picture         =   "SZ-RR.frx":C1F7
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   17
      Top             =   0
      Width           =   6255
   End
   Begin VB.PictureBox Picture5 
      Height          =   5055
      Left            =   0
      Picture         =   "SZ-RR.frx":2284A
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   16
      Top             =   0
      Width           =   6255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "RETURN"
      Height          =   495
      Left            =   8880
      TabIndex        =   15
      Top             =   9240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   16320
      Picture         =   "SZ-RR.frx":310EC
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture3 
      Height          =   3015
      Left            =   16320
      Picture         =   "SZ-RR.frx":36AB0
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   16320
      Picture         =   "SZ-RR.frx":3BBF5
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   9
      Top             =   7800
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MAIN MENU"
      Height          =   735
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10200
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BOOK"
      Height          =   855
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Height          =   855
      Left            =   6720
      Picture         =   "SZ-RR.frx":40EF3
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Height          =   855
      Left            =   4680
      Picture         =   "SZ-RR.frx":41614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   2520
      Picture         =   "SZ-RR.frx":41D59
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10080
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   360
      Picture         =   "SZ-RR.frx":4245D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10080
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   11175
      Left            =   360
      Picture         =   "SZ-RR.frx":42E51
      ScaleHeight     =   11115
      ScaleWidth      =   14955
      TabIndex        =   0
      Top             =   -960
      Width           =   15015
   End
   Begin VB.Image Image3 
      Height          =   8820
      Left            =   3480
      Picture         =   "SZ-RR.frx":74940
      Top             =   960
      Visible         =   0   'False
      Width           =   12750
   End
   Begin VB.Image Image2 
      Height          =   8820
      Left            =   3600
      Picture         =   "SZ-RR.frx":92219
      Top             =   960
      Visible         =   0   'False
      Width           =   12750
   End
   Begin VB.Image Image1 
      Height          =   8820
      Left            =   3600
      Picture         =   "SZ-RR.frx":B12F1
      Top             =   960
      Visible         =   0   'False
      Width           =   12750
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "QUALITY BLACK"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15480
      TabIndex        =   14
      Top             =   7200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "QUALITY RED"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   15480
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "QUALITY BLUE"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   15480
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   $"SZ-RR.frx":CF339
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   11055
      Left            =   15480
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   $"SZ-RR.frx":CFA66
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   10695
      Left            =   15480
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Visible = True
LABEL2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture2.Visible = False
End Sub

Private Sub Command2_Click()
LABEL2.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture2.Visible = False
End Sub

Private Sub Command4_Click()
Label1.Visible = False
LABEL2.Visible = False
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture2.Visible = True
Image1.Visible = False
Image2.Visible = False
End Sub

Private Sub Command5_Click()
Form17.Show
End Sub

Private Sub Command6_Click()
Form28.Show
End Sub

Private Sub Command7_Click()
Picture1.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Command7.Visible = False
End Sub

Private Sub Picture2_Click()
Image1.Visible = False
Image3.Visible = True
Image2.Visible = False
Command7.Visible = True
Label1.Visible = False
LABEL2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture2.Visible = False
Picture1.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
End Sub

Private Sub Picture3_Click()
Image1.Visible = True
Command7.Visible = True
Image2.Visible = False
Label1.Visible = False
LABEL2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture2.Visible = False
Picture1.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
End Sub

Private Sub Picture4_Click()
Image1.Visible = False
Image2.Visible = True
Command7.Visible = True
Label1.Visible = False
LABEL2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture2.Visible = False
Picture1.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
End Sub
