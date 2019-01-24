VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   Caption         =   "VbPng test by The trick"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5880
      Picture         =   "frmMain.frx":C0F2
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Index           =   1
      Left            =   5220
      Picture         =   "frmMain.frx":E5D2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2580
      Width           =   1875
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FF00&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      Left            =   5220
      Picture         =   "frmMain.frx":122EC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4500
      Width           =   1875
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   7
      Top             =   6420
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   767
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":15F2F
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":164AD
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2460
      Picture         =   "frmMain.frx":169F3
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      Picture         =   "frmMain.frx":178CA
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "Command1"
      DownPicture     =   "frmMain.frx":19597
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      Picture         =   "frmMain.frx":1A557
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2235
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Option:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5220
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Statusbar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   2175
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3780
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   49344
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1B630
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1BBA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1C136
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1C756
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1CB8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D0E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D662
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1DAB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1DD8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1E2D1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Picturebox:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Buttons:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Images:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1860
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   900
      Picture         =   "frmMain.frx":1E7EC
      Top             =   2160
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   480
      Picture         =   "frmMain.frx":1F6C3
      Top             =   2160
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "frmMain.frx":20683
      Top             =   2160
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   3210
      Left            =   120
      Picture         =   "frmMain.frx":21653
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   3390
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Using VbPng.dll test
' // By The trick
' // 2019

Option Explicit

Private Declare Function InitCommonControlsEx Lib "comctl32" ( _
                         ByRef icc As Any) As Long
Private Declare Function Initialize Lib "..\..\VbPngLibCpp\ReleaseDll\VBPng.dll" () As Long
Private Declare Sub Uninitialize Lib "..\..\VbPngLibCpp\ReleaseDll\VBPng.dll" ()

Private Sub Form_Initialize()

    InitCommonControlsEx 3435973.8623@
    
    If Initialize() = 0 Then
        MsgBox "Unable to initialize png dll", vbCritical
    End If
    
End Sub

Private Sub Form_Terminate()
    Uninitialize
End Sub


