VERSION 5.00
Begin VB.UserControl ctlTest 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ctlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Set Picture( _
                    ByVal cPic As StdPicture)
    Set UserControl.Picture = cPic
    PropertyChanged ("Picture")
End Property

Public Property Get Picture() As StdPicture
    Set Picture = UserControl.Picture
End Property

Private Sub UserControl_ReadProperties( _
            ByRef PropBag As PropertyBag)
    Set UserControl.Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties( _
            ByRef PropBag As PropertyBag)
    PropBag.WriteProperty "Picture", UserControl.Picture, Nothing
End Sub


