VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000F&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1455
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   Picture         =   "ADMIN MDI.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   9015
      Left            =   0
      Picture         =   "ADMIN MDI.frx":1C6A
      ScaleHeight     =   8955
      ScaleWidth      =   20190
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      Begin VB.Label Label1 
         Caption         =   "Designed And Developed By:-  Akshay Kumar.C.P.  Dilip.G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   15600
         TabIndex        =   1
         Top             =   6960
         Width           =   4335
      End
   End
   Begin VB.Menu EMPLOYEE 
      Caption         =   "EMPLOYEE"
      Begin VB.Menu AddEmployee 
         Caption         =   "AddEmployee"
         Shortcut        =   ^C
      End
      Begin VB.Menu UpdateEmployee 
         Caption         =   "UpdateEmployee"
         Shortcut        =   ^O
      End
      Begin VB.Menu DeleteEmployee 
         Caption         =   "DeleteEmployee"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu IMPORT 
      Caption         =   "IMPORT"
      Begin VB.Menu Item 
         Caption         =   "Item"
      End
      Begin VB.Menu Vehicle 
         Caption         =   "Vehicle"
      End
   End
   Begin VB.Menu REPORT 
      Caption         =   "REPORT"
      Begin VB.Menu CustomerReport 
         Caption         =   "CustomerReport"
      End
      Begin VB.Menu AllEmployeeReport 
         Caption         =   "AllEmployeeReport"
      End
      Begin VB.Menu VehicleImported 
         Caption         =   "VehicleImportrd"
      End
      Begin VB.Menu BikeSaled 
         Caption         =   "BikeSaled"
      End
   End
   Begin VB.Menu CHANGEPASSWORD 
      Caption         =   "CHANGEPASSWORD"
   End
   Begin VB.Menu ABOUTSHOWROOM 
      Caption         =   "ABOUTSHOWROOM"
   End
   Begin VB.Menu ABOUT 
      Caption         =   "ABOUT"
   End
   Begin VB.Menu EXIT 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ABOUT_Click()
Form30.Show
End Sub

Private Sub ABOUTSHOWROOM_Click()
Form14.Show
End Sub

Private Sub AddEmployee_Click()
Form4.Show
End Sub

Private Sub AllEmployeeReport_Click()
DataReport2.Show
End Sub

Private Sub BikeSaled_Click()
DataReport4.Show
End Sub

Private Sub CHANGEPASSWORD_Click()
Form9.Show
End Sub

Private Sub CustomerReport_Click()
DataReport3.Show
End Sub

Private Sub DeleteEmployee_Click()
Form6.Show
End Sub

Private Sub MANAGERECORD_Click()
Form7.Show
End Sub

Private Sub EXIT_Click()
Form2.Show
End Sub

Private Sub Item_Click()
Form11.Show
End Sub

Private Sub SparePartsInvoice_Click()
DataReport1.Show
End Sub

Private Sub UpdateEmployee_Click()
Form5.Show
End Sub

Private Sub Vehicle_Click()
Form12.Show
End Sub

Private Sub VehicleImported_Click()
DataReport6.Show
End Sub
