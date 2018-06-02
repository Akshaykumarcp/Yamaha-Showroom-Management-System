VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form11"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   0
      Picture         =   "ItemIndend.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   12
      Top             =   0
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   2280
      Top             =   4920
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
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
      RecordSource    =   "Item_indend"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ItemIndend.frx":0CD2
      Height          =   1695
      Left            =   6480
      TabIndex        =   11
      Top             =   7440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "item_name"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "ItemIndend.frx":0CE7
      Left            =   9480
      List            =   "ItemIndend.frx":0CFD
      TabIndex        =   10
      Text            =   "Select"
      Top             =   4560
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "item_no"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "ItemIndend.frx":0D3B
      Left            =   9480
      List            =   "ItemIndend.frx":0D51
      TabIndex        =   9
      Text            =   "Select"
      Top             =   3480
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   615
      Left            =   12000
      Picture         =   "ItemIndend.frx":0D79
      TabIndex        =   8
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   615
      Left            =   9720
      TabIndex        =   7
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Indend"
      Height          =   615
      Left            =   7200
      TabIndex        =   6
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   6840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "u_o_m"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   9480
      TabIndex        =   3
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Indend"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   6960
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit of measure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5760
      TabIndex        =   2
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   16200
      Left            =   -4800
      Picture         =   "ItemIndend.frx":4F5E
      Top             =   -3240
      Width           =   28800
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveLast
Combo1.SetFocus
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
Text1.SetFocus
Adodc1.Recordset.Fields(0) = Combo1.Text
Adodc1.Recordset.Fields(1) = Combo2.Text
Adodc1.Recordset.Fields(2) = Text1.Text
Adodc1.Recordset.Update
End Sub

Private Sub Command3_Click()
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
End Sub
