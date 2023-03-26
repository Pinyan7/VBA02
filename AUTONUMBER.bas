Attribute VB_Name = "AUTONUMBER"
Sub 插入D句號()
    Dim myRange As Range
    Set myRange = ActiveDocument.content
    Call 換行
    
    With myRange.find
        .ClearFormatting
        .Text = "(D)"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
 Do While myRange.find.Execute
    ' Check if next paragraph contains "(E)"
    If InStr(myRange.Next(wdParagraph).Text, "(E)") = 0 Then
        ' Insert period after paragraph if next paragraph doesn't contain "(E)"
        myRange.MoveEnd wdParagraph, 1
        myRange.MoveEnd wdCharacter, -1
        myRange.Collapse wdCollapseEnd
        myRange.InsertAfter "。"
        myRange.Collapse wdCollapseEnd
    End If
Loop
Call 插入E句號
Call 取代五四三的
Call 將題目自動編號
Call 更改選項的樣式
End Sub

Sub 插入E句號()
    Dim myRange As Range
    Set myRange = ActiveDocument.content
    
    With myRange.find
        .ClearFormatting
        .Text = "(E)"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
 Do While myRange.find.Execute
        myRange.MoveEnd wdParagraph, 1
        myRange.MoveEnd wdCharacter, -1
        myRange.Collapse wdCollapseEnd
        myRange.InsertAfter "。"
        myRange.Collapse wdCollapseEnd
Loop
End Sub
Sub 取代五四三的()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If
 mystr = area.find.Execute(findtext:="^#^#.", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="^#.", ReplaceWith:="", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="^t", ReplaceWith:="", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="()", ReplaceWith:="", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:=" ", ReplaceWith:="", Replace:=wdReplaceAll)
     mystr = area.find.Execute(findtext:="★", ReplaceWith:="", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="「」", ReplaceWith:="「　」", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="：(A)", ReplaceWith:="：" & Chr(13) & "(A)", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="。」。", ReplaceWith:="。」", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="！」。", ReplaceWith:="！」", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="？」。", ReplaceWith:="？」", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="(^#)()", ReplaceWith:="", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="(共0分,每題0分)", ReplaceWith:="", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="※", ReplaceWith:="【篇章】", Replace:=wdReplaceAll)
       ActiveDocument.content.Select
   Selection.Style = ActiveDocument.Styles("00")
       
End Sub
Sub 將題目自動編號()
  Dim oPara As Paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(A)") > 0 Then
     oPara.Previous.Range.Style = ActiveDocument.Styles("1")
    End If
  Next
End Sub

Sub 更改選項的樣式()
  Dim oPara As Paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(A)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(B)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(C)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(D)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
   For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(E)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
End Sub
