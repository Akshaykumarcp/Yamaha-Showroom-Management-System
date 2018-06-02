VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "DeleteEmployee.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12360
      Top             =   8400
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\database3.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\database3.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employee"
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
   Begin VB.TextBox Text1 
      DataField       =   "EMPLOYEE_ID"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   15000
      TabIndex        =   3
      Top             =   5160
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14880
      TabIndex        =   2
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11400
      TabIndex        =   1
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "  DELETE     ACCOUNT"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   12120
      TabIndex        =   4
      Top             =   2640
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "           ENTER  EMPLOYEE ID"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   9360
      TabIndex        =   0
      Top             =   5160
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -120
      Picture         =   "DeleteEmployee.frx":0CD2
      Top             =   -6000
      Width           =   24000
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errmsg
Dim flag As String
flag = MsgBox("Sure to delete the record?", vbYesNo + vbInformation)
If flag = vbYes Then
Adodc1.Recordset.Delete
MsgBox " Deleted successfully "
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveLast
End If
Exit Sub
End If
errmsg:
MsgBox Err.Description
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub



