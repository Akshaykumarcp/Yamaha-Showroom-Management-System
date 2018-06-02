VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form15"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox frm11 
      Height          =   10935
      Left            =   0
      Picture         =   "VehicleSales.frx":0000
      ScaleHeight     =   10875
      ScaleWidth      =   20115
      TabIndex        =   0
      Top             =   0
      Width           =   20175
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
         Left            =   11880
         TabIndex        =   24
         Top             =   8880
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
         Left            =   11160
         TabIndex        =   23
         Top             =   480
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
         Left            =   9960
         TabIndex        =   22
         Top             =   8040
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
         Left            =   11880
         TabIndex        =   21
         Top             =   8040
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
         Left            =   9960
         TabIndex        =   20
         Top             =   8880
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Back 
         BackColor       =   &H00C0C0FF&
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
         Left            =   10920
         TabIndex        =   19
         Top             =   4560
         Width           =   2535
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
         Height          =   855
         Left            =   10920
         TabIndex        =   18
         Top             =   5280
         Width           =   2535
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
         Left            =   11160
         TabIndex        =   17
         Top             =   2160
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
         Left            =   11160
         TabIndex        =   16
         Top             =   3000
         Width           =   1935
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
         Left            =   11160
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "VEHICLE SALES TRANSACTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   7095
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   9495
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
            TabIndex        =   8
            Top             =   1680
            Width           =   3615
         End
         Begin VB.TextBox txt_Customer_Code 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "CUSTOMERCODE"
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
            TabIndex        =   7
            Top             =   3240
            Width           =   3615
         End
         Begin VB.TextBox txt_Sales_Date 
            Alignment       =   2  'Center
            DataField       =   "SALESDATE"
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
            Height          =   570
            Left            =   4800
            TabIndex        =   6
            Top             =   2400
            Width           =   3615
         End
         Begin VB.TextBox txt_Quantity 
            Alignment       =   2  'Center
            DataField       =   "QUANTITY"
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
            TabIndex        =   5
            Top             =   4920
            Width           =   3615
         End
         Begin VB.TextBox txt_Price 
            Alignment       =   2  'Center
            DataField       =   "PRICE"
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
            TabIndex        =   4
            Top             =   4080
            Width           =   3615
         End
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
            TabIndex        =   3
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label lbl_Sales_Date 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Date"
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
            TabIndex        =   14
            Top             =   2520
            Width           =   3975
         End
         Begin VB.Label lbl_Item_Name 
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
            Left            =   600
            TabIndex        =   13
            Top             =   1800
            Width           =   3975
         End
         Begin VB.Label lbl_Price 
            BackStyle       =   0  'Transparent
            Caption         =   "Price"
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
            TabIndex        =   12
            Top             =   4200
            Width           =   3975
         End
         Begin VB.Label lbl_Quantity 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
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
            TabIndex        =   11
            Top             =   5160
            Width           =   3975
         End
         Begin VB.Label lbl_Customer_Code 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Code"
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
            TabIndex        =   10
            Top             =   3360
            Width           =   3975
         End
         Begin VB.Label lbl_Sales_Code 
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
            Left            =   600
            TabIndex        =   9
            Top             =   960
            Width           =   3255
         End
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
         Left            =   10920
         TabIndex        =   1
         Top             =   3840
         Width           =   2535
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "VehicleSales.frx":14A9F
         Height          =   2415
         Left            =   720
         TabIndex        =   25
         Top             =   7560
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
         Connect         =   "Provider=MSDAORA.1;Password=9844219074;User ID=system;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;Password=9844219074;User ID=system;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "system"
         Password        =   "9844219074"
         RecordSource    =   "vehiclesales"
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
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
