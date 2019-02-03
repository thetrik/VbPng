VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icons cursors test"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   3600
      MouseIcon       =   "frmMain.frx":14501
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":35E09
      ScaleHeight     =   1875
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   2040
      Width           =   3315
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "From ANI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         TabIndex        =   7
         Top             =   600
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   1920
         Index           =   3
         Left            =   -120
         Picture         =   "frmMain.frx":57711
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3480
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   120
      MouseIcon       =   "frmMain.frx":5A696
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":67026
      ScaleHeight     =   1875
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   2040
      Width           =   3435
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "From ANI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   60
         TabIndex        =   6
         Top             =   600
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   1920
         Index           =   2
         Left            =   0
         Picture         =   "frmMain.frx":739B6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3480
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   3600
      MouseIcon       =   "frmMain.frx":7693B
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":779FD
      ScaleHeight     =   1875
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   120
      Width           =   3315
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "From CUR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -60
         TabIndex        =   5
         Top             =   600
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   1920
         Index           =   1
         Left            =   -120
         Picture         =   "frmMain.frx":78ABF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3480
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   120
      MouseIcon       =   "frmMain.frx":7BA44
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":8CF51
      ScaleHeight     =   1875
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   120
      Width           =   3435
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "From ICO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   60
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   1920
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":9E45E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3480
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "All icons are downloaded from http://www.rw-designer.com/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   6795
   End
   Begin VB.Image Image2 
      Height          =   5100
      Left            =   0
      Picture         =   "frmMain.frx":A13E3
      Stretch         =   -1  'True
      Top             =   -60
      Width           =   7020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

