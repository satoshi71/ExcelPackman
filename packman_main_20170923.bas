Attribute VB_Name = "packman_main"
Sub start()

   Range("A2:ZZ200").Interior.Pattern = xlNone
   Range("AM1") = 0

   Call setStage
   Call packman(7, 7, "RIGHT")
   Call UserForm1.setPos(7, 7)
   UserForm1.Show
End Sub

Sub packman(posx, posy, direct)
   Dim pack As Variant

   If direct = "RIGHT" Then
      If posx > 150 Then Exit Sub
      For i = -2 To 2
         If Cells(posy + i, posx + 2).Interior.ThemeColor = xlThemeColorAccent4 Then Exit Sub
      Next i
      If Cells(posy, posx + 2).Interior.ThemeColor = xlThemeColorAccent5 Then Call getPoint
      pack = Array( _
              Array(0, 0, 1, 1, 0), _
              Array(0, 1, 1, 1, 1), _
              Array(1, 1, 1, 0, 0), _
              Array(0, 1, 1, 1, 1), _
              Array(0, 0, 1, 1, 0) _
          )
      Range(Cells(posy - 2, posx - 3), Cells(posy + 2, posx - 3)).Interior.Pattern = xlNone
   ElseIf direct = "LEFT" Then
      If posx < 7 Then Exit Sub
      For i = -2 To 2
         If Cells(posy + i, posx - 2).Interior.ThemeColor = xlThemeColorAccent4 Then Exit Sub
      Next i
      If Cells(posy, posx - 2).Interior.ThemeColor = xlThemeColorAccent5 Then Call getPoint
      pack = Array( _
              Array(0, 1, 1, 0, 0), _
              Array(1, 1, 1, 1, 0), _
              Array(0, 0, 1, 1, 1), _
              Array(1, 1, 1, 1, 0), _
              Array(0, 1, 1, 0, 0) _
          )
      Range(Cells(posy - 2, posx + 3), Cells(posy + 2, posx + 3)).Interior.Pattern = xlNone

   ElseIf direct = "UP" Then
      If posy < 7 Then Exit Sub
      For i = -2 To 2
         If Cells(posy - 2, posx + i).Interior.ThemeColor = xlThemeColorAccent4 Then Exit Sub
      Next i
      If Cells(posy - 2, posx).Interior.ThemeColor = xlThemeColorAccent5 Then Call getPoint
      pack = Array( _
              Array(0, 1, 0, 1, 0), _
              Array(1, 1, 0, 1, 1), _
              Array(1, 1, 1, 1, 1), _
              Array(0, 1, 1, 1, 0), _
              Array(0, 0, 1, 0, 0) _
          )
      Range(Cells(posy + 3, posx - 2), Cells(posy + 3, posx + 2)).Interior.Pattern = xlNone

   ElseIf direct = "DOWN" Then
      If posy > 70 Then Exit Sub
      For i = -2 To 2
         If Cells(posy + 2, posx + i).Interior.ThemeColor = xlThemeColorAccent4 Then Exit Sub
      Next i
      If Cells(posy + 2, posx).Interior.ThemeColor = xlThemeColorAccent5 Then Call getPoint
      pack = Array( _
              Array(0, 0, 1, 0, 0), _
              Array(0, 1, 1, 1, 0), _
              Array(1, 1, 1, 1, 1), _
              Array(1, 1, 0, 1, 1), _
              Array(0, 1, 0, 1, 0) _
          )
      Range(Cells(posy - 3, posx - 2), Cells(posy - 3, posx + 2)).Interior.Pattern = xlNone
   End If

   For i = 0 To UBound(pack)
      For j = 0 To UBound(pack(i))
         If pack(i)(j) = 1 Then
            Cells(posy - 2 + i, posx - 2 + j).Interior.Color = 49407
         Else
            Cells(posy - 2 + i, posx - 2 + j).Interior.Pattern = xlNone
         End If
      Next j
   Next i

   Call UserForm1.setPos(posx, posy)
End Sub


Sub setStage()
   sname = Range("V1")
   If sname = "" Then
      MsgBox ("ステージ名が選択されていません。")
      Range("V1").Select
      End
   End If

   On Error GoTo ERROR_CATCH

   Application.ScreenUpdating = False
      Sheets(sname).Select
      Range("A2:ZZ100").Select
      Selection.Copy
      Sheets("Main").Select
      Range("A2").Select
      ActiveSheet.Paste
      Range("A2").Select
   Application.ScreenUpdating = True
   Exit Sub

ERROR_CATCH:
   MsgBox ("ERROR!! 選択したステージ名に対するシートが存在していない可能性があります。")
   End
End Sub



Sub getPoint()
   Range("AM1") = Range("AM1") + 1
End Sub


Sub setStageList()

   Dim stageList As String
   For Each sheet_name In Worksheets
      sname = sheet_name.Name
      If InStr(sname, "Stage") > 0 Then
         stageList = sname & "," & stageList
      End If
   Next

   Range("V1") = ""
   Range("V1").Select
   With Selection.Validation
      .Delete
      .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
      xlBetween, Formula1:=stageList
   End With

End Sub


Sub StageClear()
   Range("A2:ZZ200").Interior.Pattern = xlNone
   Call setStageList
   Range("AM1") = 0
End Sub

