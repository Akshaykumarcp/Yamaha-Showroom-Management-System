VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BackColor       =   &H80000009&
   Caption         =   "Form8"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   8160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
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
   Begin VB.TextBox Text6 
      DataField       =   "payingl_amt"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10200
      MaxLength       =   5
      TabIndex        =   17
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Text20 
      DataField       =   "totall_amtpaid"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10200
      MaxLength       =   8
      TabIndex        =   16
      Top             =   6360
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataField       =   "c_name"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      ToolTipText     =   "CUSTOMER NAME"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataField       =   "r_no"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   14
      ToolTipText     =   "REGISTER  NUMBER"
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&BACK"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8280
      Width           =   1935
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8280
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "Loan.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   2535
      TabIndex        =   11
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FFFFFF&
      DataField       =   "chasis_no"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   10
      ToolTipText     =   "CHASIS  NUMBER"
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text14 
      DataField       =   "E_no"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   9
      ToolTipText     =   "ENGINE  NUMBER"
      Top             =   4200
      Width           =   2415
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Text15 
      DataField       =   "remainingl_months"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10200
      MaxLength       =   2
      TabIndex        =   7
      ToolTipText     =   "LOAN  DURATION"
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text16 
      DataField       =   "l_amount"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10200
      MaxLength       =   5
      TabIndex        =   6
      ToolTipText     =   "LOAN  ALOUNT"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      DataField       =   "install_amt"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10200
      MaxLength       =   5
      TabIndex        =   5
      ToolTipText     =   "INSTALLEMET  AMOUNT"
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      DataField       =   "remaining_amt"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10200
      MaxLength       =   5
      TabIndex        =   4
      Top             =   5640
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "l_no"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "model"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   LOAN  PAYMENT  DETAILS   "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   3240
      TabIndex        =   29
      Top             =   1440
      Width           =   7935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYING LOAN AMOUNT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6840
      TabIndex        =   31
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL LOAN AMOUNT PAID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6240
      TabIndex        =   30
      Top             =   6360
      Width           =   3855
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   840
      TabIndex        =   28
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTER  NUMBER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Top             =   6360
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1080
      TabIndex        =   26
      Top             =   3480
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2520
      TabIndex        =   25
      Top             =   4920
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1080
      TabIndex        =   24
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "REMAINING LOAN MONTHS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6360
      TabIndex        =   23
      Top             =   4200
      Width           =   3975
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   7800
      TabIndex        =   22
      Top             =   2760
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "REMAINING  AMOUNT "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   11280
      TabIndex        =   19
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "LOAN  NO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   5640
      Width           =   1455
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()

On Error GoTo err1
a = MsgBox("Do you want to save the record", vbOKCancel + vbDefaultButton1, "SAVE")
If a = vbOK Then
  Adodc1.Recordset.Fields(0) = Text3.Text
  Adodc1.Recordset.Fields(2) = Text13.Text
  Adodc1.Recordset.Fields(3) = Text14.Text
  Adodc1.Recordset.Fields(1) = Text2.Text
  Adodc1.Recordset.Fields(12) = Text4.Text
  Adodc1.Recordset.Fields(5) = Text16.Text
  Adodc1.Recordset.Fields(6) = Text8.Text
  Adodc1.Recordset.Fields(7) = Text15.Text
   Adodc1.Recordset.Fields(11) = Text19.Text
   'Adodc1.Recordset.Fields(17) = Text20.Text
   Adodc1.Recordset.Fields(17) = Text11.Text
 ' Adodc1.Recordset.Fields(8) = Text6.Text
  'Adodc1.Recordset.Fields(9) = Text11.Text
  'Adodc1.Recordset.Fields(10) = Text20.Text
  Adodc2.Refresh
  Adodc2.Recordset.MoveFirst
  Do Until Adodc2.Recordset.EOF
    If Adodc2.Recordset.Fields(0) = Text4.Text Then
      Adodc2.Recordset.Fields(5) = Text15.Text
      Adodc2.Recordset.Update
      Adodc2.Refresh
      GoTo 1
    End If
    Adodc2.Recordset.MoveNext
  Loop
1:
  Adodc2.Refresh
  Adodc2.Recordset.MoveFirst
  Do Until Adodc2.Recordset.EOF
    If Adodc2.Recordset.Fields(0) = Text4.Text Then
      Adodc2.Recordset.Fields(7) = Text11.Text
      Adodc2.Recordset.Update
      Adodc2.Refresh
      GoTo 2
    End If
    Adodc2.Recordset.MoveNext
  Loop
2:
  Adodc1.Recordset.Update
  Adodc1.Refresh
  'Call disable
  MsgBox "RECORD SAVED", vbOKCancel, "SAVED"
Else
  MsgBox "Record NOT SAVED"
  Adodc1.Recordset.CancelUpdate
End If
err1:
MsgBox Err.Description, vbQuestion, "Primary Key Error"
End Sub

Private Sub Command8_Click()
Adodc1.Refresh
'Call enable
If Adodc1.Recordset.RecordCount = 0 Then
  a = 1
Else
  Adodc1.Recordset.MoveLast
'  a = Adodc1.Recordset.Fields(16) + 1
End If
Adodc1.Recordset.AddNew
'Text18.Text = a
'Text17.SetFocus
End Sub

