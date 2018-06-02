VERSION 5.00
Begin VB.Form Form30 
   Caption         =   "Form30"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form30"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "About.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   17
      Top             =   0
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Product Info"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   7440
      TabIndex        =   11
      Top             =   2880
      Width           =   6975
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "YAMAHA SHOWROOM MANAGEMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   1320
         TabIndex        =   16
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5760
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Product:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "License"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   7440
      TabIndex        =   6
      Top             =   6840
      Width           =   6975
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Registered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Freeware"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Contributor"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   7440
      TabIndex        =   0
      Top             =   4680
      Width           =   6975
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Developers:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "AKSHAY KUMAR.C.P(12XWSB6004)"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "DILIP.G(12XWSB6015)"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Guided by:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Mrs.Amalorpavam.G"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   1440
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
