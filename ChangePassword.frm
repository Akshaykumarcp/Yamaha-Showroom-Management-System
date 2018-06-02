VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "ChangePassword.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   11
      Top             =   0
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   15000
      Top             =   8040
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\database3.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\database3.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "admin_login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17280
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   8
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   17040
      TabIndex        =   7
      Top             =   5760
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   17040
      TabIndex        =   6
      Top             =   4920
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "Password"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   17040
      TabIndex        =   5
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataField       =   "User"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17040
      TabIndex        =   0
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   12000
      TabIndex        =   10
      Top             =   2400
      Width           =   8055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Enter the new Password"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   11880
      TabIndex        =   4
      Top             =   5880
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   11880
      TabIndex        =   3
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   11880
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   11880
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -3480
      Picture         =   "ChangePassword.frx":0CD2
      Top             =   -1200
      Width           =   24000
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim prev As String
prev = Text2.Text
If Text3.Text = Text4.Text Then
Adodc1.Recordset("Password") = Text3.Text
Adodc1.Recordset.Update
MsgBox "Changed Password Successfully"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
ElseIf Text4.Text <> Text3.Text Then
Adodc1.Recordset.CancelUpdate
MsgBox "Re-Enter the New Password correctly"
Text4.Text = ""
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_LostFocus()
On Error GoTo errmsg
Dim username As String
username = Text1.Text
If username = "" Then
Exit Sub
Else
Adodc1.Refresh
Adodc1.Recordset.Find "Username='" + username + "'"
Text3.SetFocus
If Adodc1.Recordset.EOF Then
MsgBox "Username dosen't exists"
Text1.Text = ""
Exit Sub
End If
End If
Exit Sub
errmsg:
MsgBox ("Sorry error occured")
End Sub
