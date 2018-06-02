VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form17 
   BackColor       =   &H80000007&
   Caption         =   "Form17"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form17"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Personal Details"
      Height          =   3975
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   6495
      Begin VB.TextBox Text5 
         DataField       =   "cust_phone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   3360
         Width           =   4335
      End
      Begin VB.TextBox Text4 
         DataField       =   "cust_pin"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         DataField       =   "cust_city"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         DataField       =   "cust_address"
         DataSource      =   "Adodc1"
         Height          =   975
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         DataField       =   "cust_name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Pin"
         Height          =   615
         Left            =   3840
         TabIndex        =   8
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Phone"
         Height          =   615
         Left            =   960
         TabIndex        =   4
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "City"
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6360
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "Book"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   0
      Picture         =   "Booking.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   38
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   33
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   32
      Top             =   9120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   31
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   30
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Mode of payment"
      Height          =   3735
      Left            =   7680
      TabIndex        =   23
      Top             =   5040
      Width           =   5535
      Begin VB.TextBox Text11 
         DataField       =   "bank"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3120
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         DataField       =   "cheque"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3120
         TabIndex        =   36
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cheque"
         Height          =   495
         Left            =   720
         TabIndex        =   29
         Top             =   2520
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cash"
         Height          =   315
         Left            =   720
         TabIndex        =   28
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         DataField       =   "amt_due"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2280
         TabIndex        =   27
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         DataField       =   "amt_paid"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2280
         TabIndex        =   26
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label14 
         Caption         =   "Bank"
         Height          =   375
         Left            =   1920
         TabIndex        =   35
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Cheque No"
         Height          =   375
         Left            =   1920
         TabIndex        =   34
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Amount Due"
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Amount Paid"
         Height          =   495
         Left            =   720
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Other Details"
      Height          =   3975
      Left            =   7680
      TabIndex        =   18
      Top             =   1080
      Width           =   5535
      Begin VB.TextBox Text7 
         DataField       =   "allot_no"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         DataField       =   "reg_date"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Allot no"
         Height          =   495
         Left            =   720
         TabIndex        =   21
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Reg Date"
         Height          =   375
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vehicle Details"
      Height          =   3735
      Left            =   1200
      TabIndex        =   11
      Top             =   5040
      Width           =   6495
      Begin VB.ComboBox Combo3 
         DataField       =   "veh_color"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Booking.frx":0CD2
         Left            =   2280
         List            =   "Booking.frx":0CE5
         TabIndex        =   17
         Top             =   2520
         Width           =   3375
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "veh_model"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Booking.frx":0D0A
         Left            =   2280
         List            =   "Booking.frx":0D0C
         TabIndex        =   16
         Top             =   1560
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "veh_name"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Booking.frx":0D0E
         Left            =   2280
         List            =   "Booking.frx":0D30
         TabIndex        =   15
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Vehicle Color"
         Height          =   495
         Left            =   960
         TabIndex        =   14
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Vehicle Model"
         Height          =   495
         Left            =   960
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Vehicle Name"
         Height          =   615
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   6465
      Left            =   13200
      Picture         =   "Booking.frx":0D74
      Top             =   1680
      Width           =   7035
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = Text1.Text
Adodc1.Recordset.Fields(1) = Text2.Text
Adodc1.Recordset.Fields(2) = Text3.Text
Adodc1.Recordset.Fields(3) = Text4.Text
Adodc1.Recordset.Fields(4) = Text5.Text
Adodc1.Recordset.Fields(5) = Text6.Text
Adodc1.Recordset.Fields(6) = Text7.Text
Adodc1.Recordset.Fields(7) = Combo1.Text
Adodc1.Recordset.Fields(8) = Combo2.Text
Adodc1.Recordset.Fields(9) = Combo3.Text
Adodc1.Recordset.Fields(10) = Text8.Text
Adodc1.Recordset.Fields(11) = Text9.Text
Adodc1.Recordset.Fields(12) = Text10.Text
Adodc1.Recordset.Fields(13) = Text11.Text
Adodc1.Recordset.Update
Option1.Value = False
Option2.Value = False
MsgBox "Booked successfully"
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
End Sub

Private Sub Command4_Click()
Form20.Show
End Sub

Private Sub Command3_Click()
DataReport5.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Option2_Click()
Label13.Visible = True
Label14.Visible = True
Text10.Visible = True
Text11.Visible = True
End Sub
