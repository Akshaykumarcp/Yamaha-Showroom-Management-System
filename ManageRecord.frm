VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_Sales 
      Caption         =   "Sales Master"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   7
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Vender 
      Caption         =   "Vendor"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   6
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Vehicle 
      Caption         =   "Vehicle "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Purchase 
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14280
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Item 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Customer 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Company 
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   1
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Vehicle_Sales 
      Caption         =   "Vehicle Sales"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   0
      Top             =   4800
      Width           =   2895
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Company_Click()
Form13.Show
End Sub

Private Sub cmd_Customer_Click()
Form10.Show
End Sub

Private Sub cmd_Item_Click()
Form9.Show
End Sub

Private Sub cmd_Purchase_Click()
Form12.Show
End Sub

Private Sub cmd_Sales_Click()
Form8.Show
End Sub

Private Sub cmd_Vehicle_Click()
Form11.Show
End Sub

Private Sub cmd_Vehicle_Sales_Click()
Form15.Show
End Sub

Private Sub cmd_Vender_Click()
Form14.Show
End Sub
