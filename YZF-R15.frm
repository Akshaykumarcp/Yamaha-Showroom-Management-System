VERSION 5.00
Begin VB.Form Form22 
   Caption         =   "Form22"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form22"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picture1 
      Height          =   9135
      Left            =   840
      Picture         =   "YZF-R15.frx":0000
      ScaleHeight     =   9075
      ScaleWidth      =   12795
      TabIndex        =   12
      Top             =   360
      Width           =   12855
   End
   Begin VB.PictureBox Picture8 
      Height          =   9735
      Left            =   0
      Picture         =   "YZF-R15.frx":22F00
      ScaleHeight     =   9675
      ScaleWidth      =   12795
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton Command7 
         Caption         =   "RETURN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5280
         TabIndex        =   16
         Top             =   8880
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Main menu"
      Height          =   735
      Left            =   11280
      TabIndex        =   21
      Top             =   10080
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Height          =   3015
      Left            =   16320
      Picture         =   "YZF-R15.frx":42A32
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   7680
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   16320
      Picture         =   "YZF-R15.frx":49E04
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture3 
      Height          =   3135
      Left            =   16320
      Picture         =   "YZF-R15.frx":50CF5
      ScaleHeight     =   3075
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Book"
      Height          =   735
      Left            =   8760
      TabIndex        =   3
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Height          =   855
      Left            =   6240
      Picture         =   "YZF-R15.frx":57FE1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   3960
      Picture         =   "YZF-R15.frx":58702
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   1440
      Picture         =   "YZF-R15.frx":58E06
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9960
      Width           =   2175
   End
   Begin VB.PictureBox picture9 
      Height          =   9735
      Left            =   4200
      Picture         =   "YZF-R15.frx":597FA
      ScaleHeight     =   9675
      ScaleWidth      =   12795
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton Command9 
         Caption         =   "RETURN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5280
         TabIndex        =   20
         Top             =   8880
         Width           =   2535
      End
   End
   Begin VB.PictureBox picture7 
      Height          =   9735
      Left            =   4440
      Picture         =   "YZF-R15.frx":788CB
      ScaleHeight     =   9675
      ScaleWidth      =   12795
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton Command8 
         Caption         =   "RETURN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5160
         TabIndex        =   18
         Top             =   9000
         Width           =   2535
      End
   End
   Begin VB.PictureBox picture6 
      Height          =   9735
      Left            =   4440
      Picture         =   "YZF-R15.frx":930C2
      ScaleHeight     =   9675
      ScaleWidth      =   12795
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton Command6 
         Caption         =   "RETURN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5280
         TabIndex        =   14
         Top             =   8880
         Width           =   2535
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Racing Bliue"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   13800
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Invincible Black"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   2655
      Left            =   13800
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Raring Red"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2655
      Left            =   13800
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   $"YZF-R15.frx":B2BF4
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Left            =   13920
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   $"YZF-R15.frx":B36A3
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   13920
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Visible = True
Label2.Visible = False
'Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
'Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
End Sub

Private Sub Command10_Click()
Form28.Show
End Sub

Private Sub Command2_Click()
Label2.Visible = True
Label1.Visible = False
'Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
'Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
End Sub

Private Sub Command3_Click()
Label1.Visible = False
Label2.Visible = False
'Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
'Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
End Sub

Private Sub Command4_Click()
Label1.Visible = False
Label2.Visible = False
'Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
'Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
End Sub

Private Sub Command5_Click()
Form17.Show
End Sub

Private Sub Command6_Click()
Form22.Show
Picture6.Visible = False
Command1.Visible = True
Picture1.Visible = True
End Sub

Private Sub Command8_Click()
Form22.Show
Picture7.Visible = False
Command1.Visible = True
Picture1.Visible = True
End Sub

Private Sub Command9_Click()
Form22.Show
Picture9.Visible = False
Command1.Visible = True
Picture1.Visible = True
End Sub

Private Sub Picture3_Click()
Picture6.Visible = True
Label1.Visible = False
Label2.Visible = False
'Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Picture1.Visible = False
'Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Command1.Visible = False
End Sub

Private Sub Picture4_Click()
Picture6.Visible = False
Label1.Visible = False
Label2.Visible = False
'Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Picture1.Visible = False
'Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Command1.Visible = False
Picture7.Visible = True
End Sub

Private Sub Picture5_Click()
Picture6.Visible = False
Label1.Visible = False
Label2.Visible = False
'Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Picture1.Visible = False
'Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Command1.Visible = False
Picture9.Visible = True
End Sub

