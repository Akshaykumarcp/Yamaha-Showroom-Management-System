VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form15"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "Loandetails.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   39
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "colour"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11040
      TabIndex        =   38
      Top             =   8160
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      DataField       =   "showroom_price"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11040
      TabIndex        =   37
      Top             =   7440
      Width           =   2415
   End
   Begin VB.TextBox Text14 
      DataField       =   "e_no"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11040
      TabIndex        =   36
      Top             =   6720
      Width           =   2415
   End
   Begin VB.TextBox Text13 
      DataField       =   "chasis_no"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11040
      TabIndex        =   35
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "model"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11040
      TabIndex        =   34
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text12 
      DataField       =   "v_id"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11040
      TabIndex        =   33
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "c_name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11040
      TabIndex        =   32
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataField       =   "c_id"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11040
      TabIndex        =   31
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataField       =   "d_o_p"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   17400
      TabIndex        =   30
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9000
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   7080
      Picture         =   "Loandetails.frx":0CD2
      ScaleHeight     =   615
      ScaleWidth      =   2535
      TabIndex        =   10
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9000
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9000
      Width           =   1935
   End
   Begin VB.TextBox Text20 
      DataField       =   "remainl_amt"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   17400
      TabIndex        =   7
      Top             =   7560
      Width           =   2415
   End
   Begin VB.TextBox Text15 
      DataField       =   "l_duration"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   17400
      MaxLength       =   5
      TabIndex        =   6
      ToolTipText     =   "LOAN  DURATION"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text16 
      DataField       =   "l_amount"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   17400
      MaxLength       =   5
      TabIndex        =   5
      ToolTipText     =   "LOAN  ALOUNT"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      DataField       =   "instal_amt"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   17400
      MaxLength       =   5
      TabIndex        =   4
      ToolTipText     =   "INSTALLEMET  AMOUNT"
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      DataField       =   "down_pay"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   17400
      MaxLength       =   5
      TabIndex        =   3
      ToolTipText     =   "DOWN  PAYMENT"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      DataField       =   "total_amt_paid"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   17400
      MaxLength       =   5
      TabIndex        =   2
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&CALCULATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7920
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "Loan"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   LOAN    DETAILS   "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   12840
      TabIndex        =   29
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER   ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   8640
      TabIndex        =   28
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER   NAME"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   8160
      TabIndex        =   27
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "VEHICLE  ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   9000
      TabIndex        =   26
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SHOWROOM   PRICE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   8040
      TabIndex        =   25
      Top             =   7560
      Width           =   2775
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CHASIS  NUMBER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   8400
      TabIndex        =   24
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MODEL"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   9720
      TabIndex        =   23
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ENGINE  NUMBER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   8280
      TabIndex        =   22
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "REMAINING LOAN AMOUNT "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   13560
      TabIndex        =   21
      Top             =   7560
      Width           =   3735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE   OF   PURCHASE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   13920
      TabIndex        =   20
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "LOAN  DURATION"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   14640
      TabIndex        =   19
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "LOAN  AMOUNT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   14880
      TabIndex        =   18
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "INSTALLMENT  AMOUNT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   13800
      TabIndex        =   17
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "DOWN  PAYMENT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   14640
      TabIndex        =   16
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL  AMOUNT PAID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   14040
      TabIndex        =   15
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHORISED SIGNATORY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19200
      TabIndex        =   14
      ToolTipText     =   "SIGANTORY AFTER   PRINTED"
      Top             =   10440
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00D1B3AB&
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18480
      TabIndex        =   13
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "COLOR"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   9720
      TabIndex        =   12
      Top             =   8280
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -7680
      Picture         =   "Loandetails.frx":19A4
      Top             =   -2640
      Width           =   28800
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Adodc1.Refresh
  Adodc1.Recordset.AddNew
  Text1.SetFocus
  Adodc1.Recordset.Fields(0) = Text6.Text
  Adodc1.Recordset.Fields(1) = Text3.Text
  Adodc1.Recordset.Fields(2) = Text12.Text
  Adodc1.Recordset.Fields(3) = Text1.Text
  Adodc1.Recordset.Fields(4) = Text13.Text
  Adodc1.Recordset.Fields(5) = Text14.Text
  Adodc1.Recordset.Fields(6) = Text4.Text
  Adodc1.Recordset.Fields(7) = Text2.Text
  Adodc1.Recordset.Fields(8) = Text16.Text
  Adodc1.Recordset.Fields(9) = Text7.Text
  Adodc1.Recordset.Fields(10) = Text15.Text
  Adodc1.Recordset.Fields(11) = Text5.Text
  Adodc1.Recordset.Fields(12) = Text11.Text
  Adodc1.Recordset.Fields(13) = Text8.Text
  Adodc1.Recordset.Fields(14) = Text20.Text
  'Adodc1.Recordset.Fields(15) = Text4.Text
  'Adodc1.Recordset.Fields(19) = Text2.Text
  'Adodc1.Recordset.Fields(14) = Text19.Text
  Adodc1.Recordset.Update
  Adodc1.Refresh
  'Call disable
  MsgBox "RECORD SAVED", vbOKCancel, "SAVED"
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()
On Error GoTo errmsg
Adodc1.Refresh
Adodc1.Recordset.MoveLast
Adodc1.Recordset.AddNew
Text6.SetFocus
errmsg:
MsgBox Err.Description
Text6.SetFocus
End Sub

Private Sub Command9_Click()
a = (Val(Text16.Text) * 2.5) / 100
b = Val(Text15.Text) * a
c = b + Val(Text16.Text)
Text20.Text = c
d = Val(Text7.Text)
Text11.Text = d
e = (Val(Text20.Text) / (Text15.Text))
Text8.Text = e
End Sub

