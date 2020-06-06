VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Button2 
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm_Button2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Button2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Option1_Click()

UserForm_Button2.Hide
Call direct(Label2.caption, CommandButton_Option1.caption)

End Sub
Private Sub CommandButton_Option2_Click()

UserForm_Button2.Hide
Call direct(Label2.caption, CommandButton_Option2.caption)

End Sub
