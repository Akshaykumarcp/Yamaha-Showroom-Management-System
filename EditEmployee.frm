VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "EditEmployee.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   32
      Top             =   0
      Width           =   2535
   End
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
      Left            =   14880
      TabIndex        =   31
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "EDIT  APPLICATION"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10335
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   12135
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   735
         Left            =   8160
         Top             =   6480
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1296
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
         Height          =   495
         Left            =   3720
         TabIndex        =   17
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         DataField       =   "EMPLOYEE_NAME"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   16
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         DataField       =   "emp_DOB"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   15
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox Text4 
         DataField       =   "EMPLOYEE ADRESS"
         DataSource      =   "Adodc1"
         Height          =   975
         Left            =   3720
         TabIndex        =   14
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         DataField       =   "emp_CITY"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   13
         Top             =   3720
         Width           =   3975
      End
      Begin VB.TextBox Text6 
         DataField       =   "emp_STATE"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   12
         Top             =   4440
         Width           =   3975
      End
      Begin VB.TextBox Text7 
         DataField       =   "emp_COUNTRY"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   11
         Top             =   5160
         Width           =   3975
      End
      Begin VB.TextBox Text8 
         DataField       =   "emp_PINCODE"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   10
         Top             =   6000
         Width           =   3975
      End
      Begin VB.TextBox Text9 
         DataField       =   "DATE OF JOIN"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   9
         Top             =   6720
         Width           =   3975
      End
      Begin VB.TextBox Text10 
         DataField       =   "emp_PINCODE"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   8
         Top             =   7440
         Width           =   3975
      End
      Begin VB.TextBox Text11 
         DataField       =   "emp_MOBILE"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   7
         Top             =   8160
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BROWSE"
         Height          =   375
         Left            =   9120
         TabIndex        =   6
         Top             =   4920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         DataField       =   "emp_EMAIL"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   5
         Top             =   8880
         Width           =   3975
      End
      Begin VB.TextBox Text13 
         DataField       =   "SAVINGS ACCOUNT"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   9480
         Width           =   4095
      End
      Begin VB.TextBox txtpicture 
         DataField       =   "IMAGE"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   8640
         TabIndex        =   3
         Top             =   5520
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9720
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label11 
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   30
         Top             =   8280
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "TELEPHONE"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   29
         Top             =   7560
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "DATE OF JOIN"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   28
         Top             =   6840
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "PINCODE"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   27
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "COUNTRY"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   26
         Top             =   5280
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "STATE"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   25
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "CITY"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   24
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "EMPLOYEE ADDRESS"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   23
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   22
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "EMPLOYEE NAME"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "EMPLOYEE ID"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   2775
         Left            =   8640
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   19
         Top             =   8880
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "ACCOUNT NUMBER"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   18
         Top             =   9480
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
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
      Left            =   14880
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
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
      Left            =   14880
      TabIndex        =   0
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Line Line3 
      X1              =   14520
      X2              =   14880
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line2 
      X1              =   14520
      X2              =   14880
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line1 
      X1              =   14640
      X2              =   14880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Image Image2 
      Height          =   18000
      Left            =   -2280
      Picture         =   "EditEmployee.frx":0CD2
      Top             =   -480
      Width           =   24000
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.ShowOpen
txtpicture.Text = CommonDialog1.FileName
Image1.Picture = LoadPicture(txtpicture.Text)
End Sub

Private Sub Command2_Click()
Dim id As Integer
On Error GoTo errmsg
Adodc1.Refresh
Adodc1.Recordset.MoveLast
id = Adodc1.Recordset("Employee_id") + 1
Adodc1.Recordset.AddNew
Text1.Text = id
Text1.SetFocus
Exit Sub
errmsg:
MsgBox Err.Description
Text1.SetFocus
Command1.Visible = True
End Sub

Private Sub Command3_Click()
If Text1.Text = "" _
Or Text2.Text = "" _
Or Text3.Text = "" _
Or Text4.Text = "" _
Or Text5.Text = "" _
Or Text6.Text = "" _
Or Text7.Text = "" _
Or Text8.Text = "" _
Or Text9.Text = "" _
Or Text11.Text = "" _
Or Text12.Text = "" _
Or txtpicture.Text = "" _
Or Text13.Text = "" Then
MsgBox "Please fill all the fields"
Else
Adodc1.Recordset.Fields(0) = Text1.Text
Adodc1.Recordset.Fields(1) = Text2.Text
Adodc1.Recordset.Fields(2) = Text3.Text
Adodc1.Recordset.Fields(3) = Text4.Text
Adodc1.Recordset.Fields(4) = Text5.Text
Adodc1.Recordset.Fields(5) = Text6.Text
Adodc1.Recordset.Fields(6) = Text7.Text
Adodc1.Recordset.Fields(7) = Text8.Text
Adodc1.Recordset.Fields(8) = Text9.Text
Adodc1.Recordset.Fields(9) = Text10.Text
Adodc1.Recordset.Fields(10) = Text11.Text
Adodc1.Recordset.Fields(11) = Text12.Text
Adodc1.Recordset.Fields(12) = Text13.Text
Adodc1.Recordset.Fields(13) = txtpicture.Text
Image1.Picture = LoadPicture(txtpicture.Text)
Adodc1.Recordset.Update
MsgBox "Record Saved Successfully"
Exit Sub
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

