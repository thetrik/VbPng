VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim c As Object

Private Sub Form_Load()
   
    Set c = CreateObject("VbPngAXDll_Test.CShowFormWithPng")
    
    c.ShowForm
    
End Sub
