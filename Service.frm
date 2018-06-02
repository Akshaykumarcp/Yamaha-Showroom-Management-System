VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "Service.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   69
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
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
      Height          =   1575
      Left            =   19320
      TabIndex        =   61
      Top             =   4200
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Service.frx":0CD2
      Height          =   2175
      Left            =   4800
      TabIndex        =   59
      Top             =   8400
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   3836
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
   Begin VB.TextBox Text3 
      DataField       =   "total_amount"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   15000
      TabIndex        =   58
      Top             =   7560
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6000
      Top             =   7800
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
      RecordSource    =   "Service"
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Save"
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Add"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ca&lculate"
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
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Frame Frame6 
      Caption         =   "WORK    DONE  DURING    SERVICE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   3975
      Left            =   10320
      TabIndex        =   23
      Top             =   2520
      Width           =   9015
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   1095
         Left            =   9000
         TabIndex        =   60
         Top             =   1680
         Width           =   75
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "WASHING"
         DataField       =   "store"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "BRAKE  PADS"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   37
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "NORMAL  POLISHING"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   36
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "COMSUMABLE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   35
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         Caption         =   "OILING"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   34
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         Caption         =   "SPARK  PLUG"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         TabIndex        =   33
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         Caption         =   "DRIVE  CHAIN"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         Caption         =   "SHOCK ABSORBER"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         TabIndex        =   31
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox Check9 
         Appearance      =   0  'Flat
         Caption         =   "BRAKE  PADS"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         TabIndex        =   30
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         Caption         =   "WHEEL  BENDING"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         TabIndex        =   29
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check11 
         Appearance      =   0  'Flat
         Caption         =   "ENGINE   BOX"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6240
         TabIndex        =   28
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CheckBox Check12 
         Appearance      =   0  'Flat
         Caption         =   "ACCELARATOR  CABLE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6240
         TabIndex        =   27
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox Check13 
         Appearance      =   0  'Flat
         Caption         =   "NORMAL  POLISHING"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6240
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox Check14 
         Appearance      =   0  'Flat
         Caption         =   "CARBURETOR"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6240
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check15 
         Appearance      =   0  'Flat
         Caption         =   "CLUTCH"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6240
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "Rs:-500"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   53
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "Rs:-1000"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   52
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label38 
         Caption         =   "Rs:-50"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   51
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Rs:-100"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   50
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Rs:-1000"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   49
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "Rs:-200"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   48
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "Rs:-800"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   47
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Rs:-600"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   46
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Rs:-300"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   45
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Rs:-50"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   44
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Rs:-100"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   43
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Rs:-150"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Rs:-230"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   41
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Rs:-250"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   40
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Rs:-750"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   39
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CUSTOMER  DETAILS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3600
      TabIndex        =   17
      Top             =   1680
      Width           =   6735
      Begin VB.TextBox Text20 
         DataField       =   "address"
         DataSource      =   "Adodc1"
         Height          =   1455
         Left            =   1200
         TabIndex        =   68
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text19 
         DataField       =   "mob_no"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   4200
         TabIndex        =   67
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text18 
         DataField       =   "c_name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   4200
         TabIndex        =   66
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text17 
         DataField       =   "cust_id"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1200
         TabIndex        =   65
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER-ID"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER  NAME"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MOBILE  NO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "VECHICLE  DETAILS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3600
      TabIndex        =   0
      Top             =   4440
      Width           =   6735
      Begin VB.TextBox Text5 
         DataField       =   "e_no"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4080
         TabIndex        =   64
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         DataField       =   "chasis_no"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4080
         TabIndex        =   63
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         DataField       =   "v_no"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   1200
         TabIndex        =   62
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         DataField       =   "reg_no"
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
         Height          =   405
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "ENTER  PERMANENT  ADDRESS"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         DataField       =   "kms"
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
         Height          =   405
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         ToolTipText     =   "PINCODE"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text6 
         DataField       =   "no_of_service"
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
         Height          =   405
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   5
         ToolTipText     =   "LANDLINE  NUMBER"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "FREE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   2880
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "PAYMENT"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   3
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         DataField       =   "model"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         DataField       =   "colour"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTER-NO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHASIS-NO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MODEL"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "VEHICLE  ID"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ENGINE -NO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NO OF SERVICE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "COLOUR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "KMS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT SERVICE  MODE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
      End
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL AMOUNT"
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
      Left            =   12480
      TabIndex        =   57
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  SERVICES   DETAILS   "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   9720
      TabIndex        =   22
      Top             =   480
      Width           =   6255
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
Text17.SetFocus
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields(0) = Text17.Text
  Adodc1.Recordset.Fields(1) = Text18.Text
  Adodc1.Recordset.Fields(2) = Text20.Text
  Adodc1.Recordset.Fields(3) = Text19.Text
  Adodc1.Recordset.Fields(5) = Text7.Text
  Adodc1.Recordset.Fields(4) = Text4.Text
  'Adodc1.Recordset.Fields(6) = Text7.Text
  Adodc1.Recordset.Fields(6) = Text9.Text
  Adodc1.Recordset.Fields(7) = Text2.Text
  'Adodc1.Recordset.Fields(9) = Text1.Text
  Adodc1.Recordset.Fields(9) = Text5.Text
    Adodc1.Recordset.Fields(10) = Text8.Text
 Adodc1.Recordset.Fields(11) = Text6.Text
  'Adodc1.Recordset.Fields(12) = Text8.Text
  Adodc1.Recordset.Fields(12) = check1check
  Adodc1.Recordset.Fields(13) = Text3.Text
   Adodc1.Recordset.Update
  'Call disable
  MsgBox "RECORD SAVED", vbOKCancel, "SAVED"
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command9_Click()
Dim a As Long
  a = 0
  If Check1.Value = 1 - Checked Then
    a = 500
  End If
  If Check2.Value = 1 - Checked Then
    a = (a + 750)
  End If
  If Check3.Value = 1 - Checked Then
    a = (a + 250)
  End If
  If Check4.Value = 1 - Checked Then
     a = (a + 230)
  End If
  If Check5.Value = 1 - Checked Then
    a = (a + 150)
  End If
   If Check6.Value = 1 - Checked Then
    a = (a + 800)
  End If
   If Check7.Value = 1 - Checked Then
    a = (a + 600)
  End If
   If Check8.Value = 1 - Checked Then
    a = (a + 300)
  End If
   If Check9.Value = 1 - Checked Then
    a = (a + 50)
  End If
   If Check10.Value = 1 - Checked Then
    a = (a + 100)
  End If
   If Check11.Value = 1 - Checked Then
    a = (a + 1000)
  End If
   If Check12.Value = 1 - Checked Then
    a = (a + 50)
  End If
   If Check13.Value = 1 - Checked Then
    a = (a + 100)
  End If
   If Check14.Value = 1 - Checked Then
    a = (a + 1000)
  End If
   If Check15.Value = 1 - Checked Then
    a = (a + 200)
  End If
 Text3.Text = 6080 - Val(a)
End Sub

