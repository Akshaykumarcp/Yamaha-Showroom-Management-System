VERSION 5.00
Begin VB.Form Form19 
   Caption         =   "Form19"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form19"
   Picture         =   "SearchRecord.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
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
      Left            =   9240
      TabIndex        =   7
      Top             =   5640
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
      Left            =   6360
      TabIndex        =   6
      Top             =   5040
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
      Left            =   12120
      TabIndex        =   5
      Top             =   3480
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
      Left            =   6240
      TabIndex        =   4
      Top             =   3600
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
      Left            =   14880
      TabIndex        =   3
      Top             =   4200
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
      Left            =   3600
      TabIndex        =   2
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Vendor 
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
      Left            =   12120
      TabIndex        =   1
      Top             =   4920
      Width           =   2895
   End
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
      Left            =   9120
      TabIndex        =   0
      Top             =   2760
      Width           =   2895
   End
End
Attribute VB_Name = "Form19"
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

Private Sub cmd_Vendor_Click()
Form14.Show
End Sub
