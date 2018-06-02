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
   Begin VB.PictureBox Picture1 
      Height          =   11055
      Left            =   0
      Picture         =   "Vehicle.frx":0000
      ScaleHeight     =   10995
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   -120
      Width           =   20295
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "VEHICLE DETAILS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5895
         Left            =   4680
         TabIndex        =   12
         Top             =   1800
         Width           =   9495
         Begin VB.TextBox txt_Vehicle_Code 
            Alignment       =   2  'Center
            DataField       =   "VEHICLECODE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   540
            Left            =   4800
            TabIndex        =   16
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox txt_Discription 
            Alignment       =   2  'Center
            DataField       =   "DESCRIPTION"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   570
            Left            =   4800
            TabIndex        =   15
            Top             =   4440
            Width           =   3615
         End
         Begin VB.TextBox txt_Company_Code 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "COMPANYCODE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   570
            Left            =   4800
            TabIndex        =   14
            Top             =   3240
            Width           =   3615
         End
         Begin VB.TextBox txt_Vehicle_Name 
            Alignment       =   2  'Center
            DataField       =   "VEHICLENAME"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   570
            Left            =   4800
            TabIndex        =   13
            Top             =   2160
            Width           =   3615
         End
         Begin VB.Label lbl_Vehicle_Code 
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Code"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   20
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label lbl_Company_Code 
            BackStyle       =   0  'Transparent
            Caption         =   "Company Code"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   19
            Top             =   3360
            Width           =   3975
         End
         Begin VB.Label lbl_Discription 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   18
            Top             =   4560
            Width           =   3975
         End
         Begin VB.Label lbl_Vehicle_Name 
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Name"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   17
            Top             =   2280
            Width           =   3975
         End
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "NEW"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14520
         TabIndex        =   11
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Delete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14520
         TabIndex        =   10
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Save 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14520
         TabIndex        =   9
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Exit 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14520
         TabIndex        =   8
         Top             =   6720
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Back 
         Caption         =   "<< BACK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14520
         TabIndex        =   7
         Top             =   5280
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Next 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14040
         TabIndex        =   6
         Top             =   9000
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Last 
         Caption         =   "LAST"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   15960
         TabIndex        =   5
         Top             =   8160
         Width           =   1935
      End
      Begin VB.CommandButton cmd_First 
         Caption         =   "FIRST"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14040
         TabIndex        =   4
         Top             =   8160
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Modify 
         Caption         =   "MODIFY"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14520
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Previous 
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   15960
         TabIndex        =   2
         Top             =   9000
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Clear 
         BackColor       =   &H00C0C0FF&
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14520
         TabIndex        =   1
         Top             =   6000
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Vehicle.frx":1B5D2
         Height          =   2415
         Left            =   4800
         TabIndex        =   21
         Top             =   7680
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4260
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
         Height          =   375
         Left            =   15000
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
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
         Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "system"
         Password        =   "9844219074"
         RecordSource    =   "vehicle"
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
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
