VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form16"
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
      Picture         =   "PurchaseCategories.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18360
      TabIndex        =   4
      Top             =   10080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Waiting list"
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
      Left            =   12240
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmd_book 
      Caption         =   "Booking"
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
      Left            =   9360
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Item 
      Caption         =   "Item"
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
      Left            =   6360
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton cmd_bike 
      Caption         =   "BIKE"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -1800
      Picture         =   "PurchaseCategories.frx":0CD2
      Top             =   -3000
      Width           =   24000
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_bike_Click()
Form28.Show
End Sub

Private Sub cmd_book_Click()
Form17.Show
End Sub

Private Sub cmd_Item_Click()
Form27.Show
End Sub

Private Sub Command1_Click()
Form18.Show
End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me
End Sub
