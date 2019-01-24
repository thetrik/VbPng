VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Loading and saving png file"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   DrawWidth       =   4
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1260
      TabIndex        =   0
      Top             =   1740
      Width           =   1635
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Saving PNG using VbPng
' // By The trick
' // 2019
' //
Option Explicit

Private Declare Function SelectObject Lib "gdi32" ( _
                         ByVal hdc As Long, _
                         ByVal hObject As Long) As Long

Dim cPic    As IPicture

' // Draw yellow circle on PNG
Private Sub Command1_Click()
    Dim hPrevDc     As Long
    Dim hPrevBmp    As Long
    
    cPic.SelectPicture Me.hdc, hPrevDc, hPrevBmp
    
    ' // Set 255 alpha
    ForeColor = (Not vbYellow) And &HFFFFFF
    DrawMode = vbBlackness
    Circle (50, 50), 40
    DrawMode = vbMergeNotPen
    Circle (50, 50), 40
    
    cPic.SelectPicture hPrevDc, 0, 0
    
    SelectObject Me.hdc, hPrevBmp
    
    SavePicture cPic, App.Path & "\out.png"
    
End Sub

Private Sub Form_Load()
    
    Set cPic = LoadPicture(App.Path & "\..\pngs\110x110_26.Png")
    
    cPic.KeepOriginalFormat = False
    
End Sub
