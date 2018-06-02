VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form yloandetails 
   BackColor       =   &H00D1B3AB&
   Caption         =   "y.loandetails"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9735
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   57
      Top             =   7680
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Left            =   13080
      Top             =   5640
      Width           =   2175
      _ExtentX        =   3836
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=import"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "import"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "mysql"
      RecordSource    =   "import"
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   13080
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=loan_temp"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "loan_temp"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "mysql"
      RecordSource    =   "loan_temp"
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   13080
      Top             =   3840
      Width           =   2175
      _ExtentX        =   3836
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=customer"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "customer"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "mysql"
      RecordSource    =   "cust_info"
      Caption         =   "Adodc2"
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
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   56
      Top             =   4800
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
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
      Height          =   480
      Left            =   4080
      TabIndex        =   55
      Top             =   2640
      Width           =   2415
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   8160
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   615
      Left            =   13080
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=vehicle"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "vehicle"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "mysql"
      RecordSource    =   "vehicle"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   13080
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=loan"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "loan"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "mysql"
      RecordSource    =   "loan"
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
   Begin MSComCtl2.DTPicker DTPicker3 
      DataField       =   "DOP"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10440
      TabIndex        =   51
      Top             =   4800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   90046465
      CurrentDate     =   41145
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&EDIT"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8880
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      DataField       =   "TAP"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10440
      MaxLength       =   5
      TabIndex        =   9
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      DataField       =   "down_payment"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10440
      MaxLength       =   5
      TabIndex        =   6
      ToolTipText     =   "DOWN  PAYMENT"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      DataField       =   "insta_a"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10440
      MaxLength       =   5
      TabIndex        =   7
      ToolTipText     =   "INSTALLEMET  AMOUNT"
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox Text16 
      DataField       =   "loan_amount"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10440
      MaxLength       =   5
      TabIndex        =   5
      ToolTipText     =   "LOAN  ALOUNT"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text15 
      DataField       =   "loan_duration"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10440
      MaxLength       =   5
      TabIndex        =   8
      ToolTipText     =   "LOAN  DURATION"
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text20 
      DataField       =   "TCAP"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10440
      TabIndex        =   10
      Top             =   6960
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8880
      Width           =   1935
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00BC8F83&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   11760
      TabIndex        =   31
      Top             =   240
      Width           =   3495
      Begin VB.TextBox Text18 
         BackColor       =   &H00FFFFFF&
         DataField       =   "loan_no"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         ToolTipText     =   "INVOICE NUMBER"
         Top             =   600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "time"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90046466
         CurrentDate     =   41145.6458333333
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "date"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         MousePointer    =   99
         CheckBox        =   -1  'True
         Format          =   90046465
         CurrentDate     =   41145.6444444444
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Text            =   "Y6540"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00FFFFFF&
         DataField       =   "e_id"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "EMPLOYEE-ID"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label24 
         BackColor       =   &H00BC8F83&
         Caption         =   "SHOWROOM-ID"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackColor       =   &H00BC8F83&
         Caption         =   "LOAN  NO"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label22 
         BackColor       =   &H00BC8F83&
         Caption         =   "EMPLOYEE-ID"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label21 
         BackColor       =   &H00BC8F83&
         Caption         =   "TIME"
         Height          =   255
         Left            =   840
         TabIndex        =   34
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label20 
         BackColor       =   &H00BC8F83&
         Caption         =   "DATE"
         Height          =   255
         Left            =   840
         TabIndex        =   33
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.TextBox Text14 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   14
      ToolTipText     =   "ENGINE  NUMBER"
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   13
      ToolTipText     =   "CHASIS  NUMBER"
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   15
      ToolTipText     =   "SHOWROOM  PRICE "
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox Text12 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   12
      ToolTipText     =   "VEHICLE  ID"
      Top             =   4080
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "yloandetails.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   2535
      TabIndex        =   25
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
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
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8880
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9840
      Width           =   1935
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9840
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      ToolTipText     =   "CUSTOMER NAME"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BC8F83&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   3960
      TabIndex        =   16
      Top             =   840
      Width           =   7455
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   54
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&SEARCH"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Vehicle Id->"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER     LOAN       DETAILS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Top             =   120
         Width           =   5535
      End
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
      Height          =   495
      Left            =   2760
      TabIndex        =   58
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00D1B3AB&
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
      Left            =   11520
      TabIndex        =   53
      Top             =   4080
      Width           =   1455
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
      Left            =   12240
      TabIndex        =   50
      ToolTipText     =   "SIGANTORY AFTER   PRINTED"
      Top             =   9840
      Width           =   3615
   End
   Begin VB.Label Label12 
      BackColor       =   &H00BC8F83&
      Caption         =   $"yloandetails.frx":0CD2
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   47
      Top             =   720
      Width           =   3255
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
      Height          =   375
      Left            =   7080
      TabIndex        =   46
      Top             =   5520
      Width           =   3015
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
      Height          =   495
      Left            =   7680
      TabIndex        =   45
      Top             =   3360
      Width           =   2415
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
      Height          =   495
      Left            =   6840
      TabIndex        =   44
      Top             =   6240
      Width           =   3375
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
      Height          =   495
      Left            =   7920
      TabIndex        =   43
      Top             =   2640
      Width           =   2175
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
      Height          =   495
      Left            =   7680
      TabIndex        =   42
      Top             =   4080
      Width           =   2415
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
      Height          =   375
      Left            =   6960
      TabIndex        =   41
      Top             =   4800
      Width           =   3135
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
      Height          =   375
      Left            =   6600
      TabIndex        =   40
      Top             =   6960
      Width           =   3735
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
      Height          =   495
      Left            =   1320
      TabIndex        =   30
      Top             =   6240
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
      Height          =   495
      Left            =   2760
      TabIndex        =   29
      Top             =   4800
      Width           =   975
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
      Height          =   495
      Left            =   1440
      TabIndex        =   28
      Top             =   5520
      Width           =   2415
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
      Height          =   495
      Left            =   1080
      TabIndex        =   27
      Top             =   6960
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
      Height          =   495
      Left            =   2040
      TabIndex        =   26
      Top             =   4080
      Width           =   1695
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
      Height          =   495
      Left            =   1200
      TabIndex        =   21
      Top             =   3360
      Width           =   2775
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
      Height          =   495
      Left            =   1680
      TabIndex        =   20
      Top             =   2640
      Width           =   2295
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
      Height          =   855
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "yloandetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_LostFocus()
On Error GoTo errmsg
If Combo2.Text = "" Then
'Combo2.SetFocus
MsgBox " Plese enter the CUSTOMER id ", vbInformation, " Warning"
Exit Sub
End If
Adodc2.Refresh
Do Until Adodc2.Recordset.EOF
If Combo2.Text = Adodc2.Recordset.Fields(0) Then
MsgBox " Valid CUSTOMER id ", vbInformation, "CUSTOMER"
Text3.Text = Adodc2.Recordset.Fields(1)
GoTo label111
End If
Adodc2.Recordset.MoveNext
Loop
'Combo2.SetFocus
MsgBox " Invalid CUSTOMER id , Please Enter a valid CUSTOMER id ", vbInformation, "Warning "
Combo2.Text = ""
Exit Sub
label111:
'MsgBox "RECORD DOESNOT EXISTS"
errmsg:
  MsgBox Err.Description
End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Then
MsgBox " Plese enter the VEHICLE id ", vbInformation, " Warning"
Exit Sub
End If
Adodc3.Refresh
Do Until Adodc3.Recordset.EOF
If Combo1.Text = Adodc3.Recordset.Fields(6) Then
MsgBox " Valid VEHICLE id ", vbInformation, "VEHICLE"
'Combo1.Text = Adodc2.Recordset.Fields(1)
'Combo3.Text = Adodc2.Recordset.Fields(2)
Text1.Text = Adodc3.Recordset.Fields(0)
Text2.Text = Adodc3.Recordset.Fields(1)
Text13.Text = Adodc3.Recordset.Fields(3)
Text14.Text = Adodc3.Recordset.Fields(4)
Text4.Text = Adodc3.Recordset.Fields(2)
Text12.Text = Combo1.Text
'Text10.Text = Adodc2.Recordset.Fields(5)
GoTo label111
End If
Adodc3.Recordset.MoveNext
Loop
MsgBox " Invalid VEHICLE id , Please Enter a valid VEHICLE id ", vbInformation, "Warning "
Combo1.Text = ""
Exit Sub
label111:
End Sub

Private Sub Command2_Click()
Me.Hide
alldetails.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
yloandetails.CommonDialog1.ShowPrinter
Command2.Visible = False
Command6.Visible = False
Command4.Visible = False
Command5.Visible = False
Command8.Visible = False
Command7.Visible = False
Frame1.Visible = False
End Sub

Private Sub Command5_Click()
yservicedetails.Show
Me.Hide
End Sub


Private Sub Command6_Click()
On Error GoTo err1
a = MsgBox("Do you want to save the record", vbOKCancel + vbDefaultButton1, "SAVE")
If a = vbOK Then
Adodc5.Refresh
   Adodc5.Recordset.MoveFirst
   Do Until Adodc5.Recordset.EOF
     If Text1.Text = Adodc5.Recordset.Fields(5) And Text2.Text = Adodc5.Recordset.Fields(6) Then
       If Adodc5.Recordset.Fields(7) > 0 Then
         Adodc5.Recordset.Fields(7) = Adodc5.Recordset.Fields(7) - 1
         Adodc5.Recordset.Update
         Adodc5.Refresh
         GoTo 1
       Else
         MsgBox "Selected Bike Sold OUT"
         GoTo 4
       End If
     End If
     Adodc5.Recordset.MoveNext
   Loop
1:
  'Adodc1.Recordset.Fields(3) = List1.Text
  'Adodc1.Recordset.Fields(12) = DTPicker3.Value
  'Adodc1.Recordset.Fields(19) = DTPicker1.Value
  'Adodc1.Recordset.Fields(19) = DTPicker2.Value
  Adodc1.Recordset.Fields(0) = Combo2.Text
  Adodc1.Recordset.Fields(1) = Text3.Text
  Adodc1.Recordset.Fields(2) = Text12.Text
  Adodc1.Recordset.Fields(3) = Text1.Text
  Adodc1.Recordset.Fields(4) = Text13.Text
  Adodc1.Recordset.Fields(5) = Text14.Text
  Adodc1.Recordset.Fields(6) = Text4.Text
  Adodc1.Recordset.Fields(19) = Text2.Text
  Adodc1.Recordset.Fields(14) = Text19.Text
  Adodc4.Refresh
  Adodc4.Recordset.AddNew
  Adodc4.Recordset.Fields(0) = Text18.Text
  Adodc4.Recordset.Fields(1) = Text4.Text
  Adodc4.Recordset.Fields(2) = Text16.Text
  Adodc4.Recordset.Fields(3) = Text7.Text
  Adodc4.Recordset.Fields(4) = Text8.Text
  Adodc4.Recordset.Fields(5) = Text15.Text
  Adodc4.Recordset.Fields(6) = Text11.Text
  Adodc4.Recordset.Fields(7) = Text20.Text
  Adodc4.Recordset.Update
  Adodc4.Refresh
  'Adodc1.Recordset.Fields(6) = Text.Text
  Adodc1.Recordset.Update
  Adodc1.Refresh
  Call disable
  MsgBox "RECORD SAVED", vbOKCancel, "SAVED"
Else
  MsgBox "Record NOT SAVED"
  Adodc1.Recordset.CancelUpdate
End If
4:
err1:
MsgBox Err.Description, vbQuestion, "Primary Key Error"
End Sub

Private Sub Command7_Click()
Call enable
End Sub

Private Sub Command8_Click()
 Adodc1.Refresh
Call enable
If Adodc1.Recordset.RecordCount = 0 Then
  a = 1
Else
  Adodc1.Recordset.MoveLast
  a = Adodc1.Recordset.Fields(0) + 1
End If
Adodc1.Recordset.AddNew
Text18.Text = a
Text17.SetFocus
End Sub

Private Sub Command9_Click()
a = (Val(Text16.Text) * 2.5) / 100
b = Val(Text15.Text) * a
c = b + Val(Text16.Text)
Text20.Text = c
End Sub

Private Sub Form_Load()
Call disable
Adodc3.Refresh
Do Until Adodc3.Recordset.EOF
Combo1.AddItem Adodc3.Recordset.Fields(6)
Adodc3.Recordset.MoveNext
Loop
Adodc2.Refresh
Do Until Adodc2.Recordset.EOF
Combo2.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End Sub

Private Sub List1_LostFocus()
If List1.Text = "" Then
MsgBox "ENTER THE VEHICLE  NAME ", vbCritical
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text11_LostFocus()
If Text11.Text = "" Then
MsgBox "ENTER THE AMOUNT  PAID ", vbCritical
End If
End Sub
Private Sub Text11KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text12_LostFocus()
If Text12.Text = "" Then
MsgBox "ENTER THE VEHICLE ID ", vbCritical
End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
Call bothvalid(KeyAscii)
End Sub

Private Sub Text13_LostFocus()
If Text13.Text = "" Then
MsgBox "ENTER THE CHASIS NUMBER ", vbCritical
End If
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text14_LostFocus()
If Text14.Text = "" Then
MsgBox "ENTER THE ENGINE NUMBER ", vbCritical
End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text15_LostFocus()
If Text15.Text = "" Then
MsgBox "ENTER THE LOAN  DURATION ", vbCritical
Else
a = (Val(Text16.Text) * 2.5) / 100
b = Val(Text15.Text) * a
c = b + Val(Text16.Text)
Text8.Text = c / Val(Text15.Text)
Text11.Text = Val(Text7.Text) + Val(Text16.Text)
'Text8.Text = ((Val(Text16.Text) * 500) / Val(Text15.Text))
End If
End Sub
Private Sub Text15_KeyPress(KeyAscii As Integer)
Call bothvalid(KeyAscii)
End Sub

Private Sub Text16_LostFocus()
If Text16.Text = "" Then
MsgBox "ENTER THE LOAN  AMOUNT ", vbCritical
End If
End Sub
Private Sub Text16_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub text2_LostFocus()
If Text2.Text = "" Then
MsgBox "ENTER THE CUSTOMER ID ", vbCritical
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text3_LostFocus()
If Text3.Text = "" Then
MsgBox "ENTER THE CUSTOMER NAME ", vbCritical
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
Call alphavalid(KeyAscii)
End Sub

Private Sub Text4_LostFocus()
If Text4.Text = "" Then
MsgBox "ENTER THE SHOWROOM PRICE ", vbCritical
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
If Text5.Text = "" Then
MsgBox "ENTER THE REGISTER NUMBER ", vbCritical
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text6_LostFocus()
If Text6.Text = "" Then
MsgBox "ENTER THE ON ROAD  AMOUNT ", vbCritical
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub
Private Sub Text7_LostFocus()
If Text7.Text = "" Then
MsgBox "ENTER THE DOWN  PAYMENT ", vbCritical
End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text8_LostFocus()
If Text8.Text = "" Then
MsgBox "ENTER THE INSTALLMENT AMOUNT ", vbCritical
End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub

Private Sub Text9_LostFocus()
If Text9.Text = "" Then
MsgBox "ENTER THE DATE OF  PURCHASE ", vbCritical
End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
Call charvalid(KeyAscii)
End Sub
Public Sub disable()
'Text6.Enabled = False
Text16.Enabled = False
Text7.Enabled = False
'Text8.Enabled = False
Text15.Enabled = False
'Text11.Enabled = False
'Text18.Enabled = False
Text17.Enabled = False
DTPicker1.Enabled = False
DTPicker2.Enabled = False
DTPicker3.Enabled = False
End Sub
Public Sub enable()
'Text6.Enabled = True
Text16.Enabled = True
Text7.Enabled = True
'Text8.Enabled = True
Text15.Enabled = True
'Text11.Enabled = True
'Text18.Enabled = True
Text17.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Enabled = True
DTPicker3.Enabled = True
End Sub
