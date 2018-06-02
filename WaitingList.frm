VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form18 
   BackColor       =   &H8000000D&
   Caption         =   "Form18"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form18"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15360
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      Picture         =   "WaitingList.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "WaitingList.frx":0CD2
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   20175
      _ExtentX        =   35586
      _ExtentY        =   12938
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
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   735
      Left            =   18120
      TabIndex        =   0
      Top             =   8760
      Width           =   1935
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Form29.Show
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.AddNew
Text1.SetFocus
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
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
