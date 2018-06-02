VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form13"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   405
      Left            =   9480
      TabIndex        =   3
      Top             =   6120
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "ENTER"
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
      Left            =   9480
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
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
      Left            =   12480
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR"
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
      Left            =   10920
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Customer allot no"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   5640
      Width           =   3615
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
