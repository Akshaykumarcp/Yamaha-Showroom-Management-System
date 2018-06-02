VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Tag             =   "S"
   WindowState     =   2  'Maximized
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6840
      Width           =   1935
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Back"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Clear"
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
      TabIndex        =   38
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
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
      TabIndex        =   37
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECT  THE  SPARE  PARTS  "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   8520
      TabIndex        =   7
      Top             =   1680
      Width           =   10335
      Begin VB.CheckBox Check1 
         Caption         =   "Wheel alloy"
         DataField       =   "Storin"
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
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   1335
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
         Left            =   6600
         TabIndex        =   21
         Top             =   480
         Width           =   1215
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
         Left            =   6600
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox Check13 
         Appearance      =   0  'Flat
         Caption         =   "LEVER"
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
         Left            =   6600
         TabIndex        =   19
         Top             =   1800
         Width           =   2295
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
         Left            =   6600
         TabIndex        =   18
         Top             =   2520
         Width           =   2415
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
         Left            =   6600
         TabIndex        =   17
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         Caption         =   "MIRROR"
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
         TabIndex        =   16
         Top             =   480
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
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
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
         TabIndex        =   14
         Top             =   1800
         Width           =   2175
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
         TabIndex        =   13
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         Caption         =   "SEAT COVER"
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
         TabIndex        =   12
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         Caption         =   "ENGINE OIL"
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
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "INDICATOR"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
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
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "TYRE"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label30 
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
         Left            =   1920
         TabIndex        =   36
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label29 
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
         Left            =   1920
         TabIndex        =   35
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label28 
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
         TabIndex        =   34
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label27 
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
         Left            =   1920
         TabIndex        =   33
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label26 
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
         Left            =   5400
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label25 
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
         Left            =   5400
         TabIndex        =   31
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label21 
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
         Left            =   5400
         TabIndex        =   30
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label20 
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
         Left            =   5400
         TabIndex        =   29
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   5400
         TabIndex        =   28
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label17 
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
         Left            =   9120
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label16 
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
         Left            =   9120
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label14 
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
         Left            =   9120
         TabIndex        =   25
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label12 
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
         Left            =   9120
         TabIndex        =   24
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label11 
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
         Left            =   9120
         TabIndex        =   23
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         Left            =   1920
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   0
      Picture         =   "Services.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "cust_id"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "cust_name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "cust_address"
      DataSource      =   "Adodc1"
      Height          =   1695
      Left            =   5280
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "cust_phone"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   5160
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "total_amt"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   15000
      TabIndex        =   0
      Top             =   7440
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Services.frx":0CD2
      Height          =   2055
      Left            =   6120
      TabIndex        =   5
      Top             =   8640
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   3625
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   690
      Left            =   1080
      Top             =   8280
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1217
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
      RecordSource    =   "Spareparts"
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
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2160
      X2              =   19440
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label1 
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   48
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   47
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   46
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   45
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Total amount"
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
      Left            =   15000
      TabIndex        =   44
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "YAMAHA SPARE PARTS INVOICE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6600
      TabIndex        =   43
      Top             =   360
      Width           =   8295
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Form10.Show
End Sub

Private Sub Command3_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
Text1.SetFocus
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = Text1.Text
Adodc1.Recordset.Fields(1) = Text2.Text
Adodc1.Recordset.Fields(2) = Text3.Text
Adodc1.Recordset.Fields(3) = Text4.Text
Adodc1.Recordset.Fields(4) = check1check
Adodc1.Recordset.Fields(5) = Text5.Text
Adodc1.Recordset.Update
MsgBox "RECORD SAVED", vbOKCancel
End Sub

Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command7_Click()
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
 Text5.Text = 6080 - Val(a)
End Sub


