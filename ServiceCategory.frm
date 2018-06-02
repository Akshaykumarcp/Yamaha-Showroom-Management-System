VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Finance"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   1
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Spare Parts"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   0
      Top             =   5880
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   16260
      Index           =   0
      Left            =   0
      Picture         =   "ServiceCategory.frx":0000
      ScaleHeight     =   16200
      ScaleWidth      =   28800
      TabIndex        =   2
      Top             =   0
      Width           =   28860
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   18600
         TabIndex        =   5
         Top             =   10200
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "ServiceCategory.frx":48B30
         ScaleHeight     =   855
         ScaleWidth      =   2535
         TabIndex        =   4
         Top             =   0
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Service"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4920
         TabIndex        =   3
         Top             =   3600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form7.Show
End Sub

Private Sub Command2_Click()
Form15.Show
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Command3_Click()
Form8.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
