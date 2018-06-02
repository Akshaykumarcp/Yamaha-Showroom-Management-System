VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   0
      Picture         =   "Login.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   8
      Top             =   0
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   17760
      Picture         =   "Login.frx":0CD2
      ScaleHeight     =   5175
      ScaleWidth      =   4215
      TabIndex        =   7
      Top             =   5400
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   0
      Left            =   17400
      Picture         =   "Login.frx":15F2B
      ScaleHeight     =   5175
      ScaleWidth      =   3615
      TabIndex        =   6
      Top             =   -960
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      TabIndex        =   4
      Top             =   9000
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   10440
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SERVICES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   3
      Top             =   7440
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PURCHASE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   2
      Top             =   5880
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADMIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   1
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE LOGIN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   975
      Left            =   7440
      TabIndex        =   0
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   17775
      Left            =   -120
      Picture         =   "Login.frx":2B184
      Stretch         =   -1  'True
      Top             =   -1680
      Width           =   20280
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
StatusBar1.Panels(1).Text = "Loading...."
StatusBar1.Panels(2).Text = "ADMIN LOADING...."
Unload Me
Form3.Show
End Sub

Private Sub Command2_Click()
StatusBar1.Panels(1).Text = "Loading...."
StatusBar1.Panels(2).Text = "RECORDS LOADING...."
Form16.Show
End Sub

Private Sub Command3_Click()
StatusBar1.Panels(1).Text = "Loading...."
StatusBar1.Panels(2).Text = "USER DETAILS...."
Form10.Show
End Sub

Private Sub Command4_Click()
If MsgBox("Sure to exit!!!!", vbYesNo) = vbYes Then
End
Else
Exit Sub
End If
End Sub
