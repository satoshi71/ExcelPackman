VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Start"
   ClientHeight    =   375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1965
   OleObjectBlob   =   "UserForm1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastposx As Integer
Dim lastposy As Integer


Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   End
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   End
End Sub

Sub setPos(x, y)
   lastposx = x
   lastposy = y
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   
   If KeyCode = vbKeyUp Then
      Call packman(lastposx, lastposy - 1, "UP")
   ElseIf KeyCode = vbKeyDown Then
      Call packman(lastposx, lastposy + 1, "DOWN")
   ElseIf KeyCode = vbKeyLeft Then
      Call packman(lastposx - 1, lastposy, "LEFT")
   ElseIf KeyCode = vbKeyRight Then
      Call packman(lastposx + 1, lastposy, "RIGHT")
   End If
End Sub


