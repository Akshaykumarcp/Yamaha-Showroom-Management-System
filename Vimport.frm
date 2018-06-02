VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form12"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      DataField       =   "total_amount"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   14280
      TabIndex        =   23
      Top             =   5520
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      DataField       =   "vat_amount"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   14280
      TabIndex        =   22
      Top             =   4800
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      DataField       =   "total_bikep"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   14280
      TabIndex        =   21
      Top             =   4080
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12360
      Top             =   8400
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "Import"
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
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Back"
      Top             =   7320
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   -120
      Picture         =   "Vimport.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   10
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
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
      Height          =   615
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Save"
      Top             =   6600
      Width           =   2295
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
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Save"
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "CLEAR  ALL  THE  DETAILS  FROM  CURRENT  PAGE"
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "CLEAR  ALL  THE  DETAILS  FROM  CURRENT  PAGE"
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataField       =   "s_price"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   5
      ToolTipText     =   "SHOWROOM PRICE"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "no_of_bikes"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   4
      ToolTipText     =   "ENTER NUMBER OF BIKES  "
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17280
      TabIndex        =   3
      Text            =   "6%"
      ToolTipText     =   "ADVANCED  AMOUNT   PAID"
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "CLEAR  ALL  THE  DETAILS  FROM  CURRENT  PAGE"
      Top             =   7320
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "model"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Vimport.frx":0CD2
      Left            =   14280
      List            =   "Vimport.frx":0CF1
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "colour"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "Vimport.frx":0D31
      Left            =   14280
      List            =   "Vimport.frx":0D44
      TabIndex        =   0
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   VEHICLE  IMPORT   DETAILS    "
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
      Left            =   9120
      TabIndex        =   20
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "COLOUR"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   19
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL  AMOUNT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   18
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "VAT AMOUNT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   17
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SHOWROOM PRICE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   16
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Left            =   12840
      TabIndex        =   15
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NO  OF  BIKES"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   14
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL BIKE PRICE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   13
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHORISED  SIGNATORY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18000
      TabIndex        =   12
      ToolTipText     =   "SIGANTORY AFTER   PRINTED"
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -5880
      Picture         =   "Vimport.frx":0D68
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command13_Click()
End Sub

Private Sub Command2_Click()
If Adodc1.Recordset.EOF = True Then
  MsgBox "NO RECORDS"
Else
  ans = MsgBox("DO you want to delete the record", vbOKCancel + vbExclamation + vbDefaultButton2, "DELETE")
  If ans = vbOK Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    MsgBox "Record Deleted", vbInformation, "Record Deleted"
  End If
End If
DataReport6.Hide
End Sub

Private Sub Command1_Click()
MDIForm1.Show
DataReport6.Hide
End Sub

Private Sub Command4_Click()
Unload Me
DataReport6.Hide
End Sub

Private Sub Command5_Click()
'Call enable
Adodc1.Refresh
'If Adodc1.Recordset.RecordCount = 0 Then
'no = 1101
'Else
'Adodc1.Recordset.MoveLast
'no = Adodc1.Recordset.Fields(1) + 1
'End If
Adodc1.Recordset.AddNew
'Text13.Text = no
Combo1.SetFocus
DataReport6.Hide
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = Combo1.Text
Adodc1.Recordset.Fields(1) = Combo2.Text
Adodc1.Recordset.Fields(2) = Text1.Text
Adodc1.Recordset.Fields(3) = Text6.Text
Adodc1.Recordset.Fields(4) = Text3.Text
Adodc1.Recordset.Fields(5) = Text7.Text
Adodc1.Recordset.Fields(6) = Text8.Text
Adodc1.Recordset.Update
  'Call disable
  MsgBox "RECORD SAVED", vbOK, "SAVED"
  DataReport6.Hide
End Sub

Private Sub Command8_Click()
DataReport6.Hide
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
'Text2.Text = ""
Text3.Text = ""
'Text4.Text = ""
'Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
'Text9.Text = ""
'Text10.Text = ""
'Text11.Text = ""
'Text12.Text = ""
'Text13.Text = ""
'Text14.Text = ""
'Text15.Text = ""
'Text16.Text = ""
End Sub

Private Sub Command9_Click()
DataReport6.Show
End Sub

Private Sub Form_Load()
DataReport6.Hide
End Sub

Private Sub Text6_Change()
'Call charvalid(KeyAscii)
DataReport6.Show
End Sub

Private Sub Text6_LostFocus()
DataReport6.Hide
Text3.Text = Val(Text1.Text) * Val(Text6.Text)
Text7.Text = (Val(Text3.Text) * 6) / 100
Text8.Text = Val(Text3.Text) + Val(Text7.Text)
End Sub
