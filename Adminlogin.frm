VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00400000&
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "Adminlogin.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   7
      Top             =   0
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3360
      Top             =   6960
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\database3.accdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\database3.accdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
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
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "Password"
      DataSource      =   "Adodc1"
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5280
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      DataField       =   "User"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "LOG IN"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   2160
      TabIndex        =   6
      Top             =   2160
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   5160
      Picture         =   "Adminlogin.frx":0CD2
      Top             =   -600
      Width           =   15360
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection

Private Sub Command1_Click()
Dim str As String
'validation ......
If Text1.Text = "" Then
MsgBox "Enter Username", vbInformation + vbOKOnly, "Admin_login"
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "Enter Password", vbInformation + vbOKOnly, "Admin_login"
Text2.SetFocus
Exit Sub
End If
'verification ......
If Text1.Text <> "" And Text2.Text <> "" Then
If rs.State = 1 Then
rs.Close
Else
str = "Select * from Admin_login where User='" & Text1.Text & "'and Password='" & Text2.Text & "'"
rs.Open str, con, adOpenDynamic, adLockOptimistic, adCmdText
End If
If rs.EOF Then
MsgBox "Invalid Username or Password", vbInformation + vbOKOnly, "Admin_login"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Else
MsgBox "You logged in successfully", vbInformation + vbOKOnly, "Admin_login"
MDIForm1.Show
Text1.Text = ""
Text2.Text = ""
End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form2.Show
End Sub

Private Sub Form_Load()
Set con = Nothing
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\database3.accdb;Persist Security Info=False"
con.CursorLocation = adUseClient
Text1.Text = ""
Text2.Text = ""
End Sub
