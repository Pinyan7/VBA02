Attribute VB_Name = "AUTONUMBER"
Sub ���JD�y��()
    Dim myRange As Range
    Set myRange = ActiveDocument.content
    Call ����
    
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
        myRange.InsertAfter "�C"
        myRange.Collapse wdCollapseEnd
    End If
Loop
Call ���JE�y��
Call ���N���|�T��
Call �N�D�ئ۰ʽs��
Call ���ﶵ���˦�
End Sub

Sub ���JE�y��()
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
        myRange.InsertAfter "�C"
        myRange.Collapse wdCollapseEnd
Loop
End Sub
Sub ���N���|�T��()
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
     mystr = area.find.Execute(findtext:="��", ReplaceWith:="", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="�u�v", ReplaceWith:="�u�@�v", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="�G(A)", ReplaceWith:="�G" & Chr(13) & "(A)", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="�C�v�C", ReplaceWith:="�C�v", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="�I�v�C", ReplaceWith:="�I�v", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="�H�v�C", ReplaceWith:="�H�v", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="(^#)()", ReplaceWith:="", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="(�@0��,�C�D0��)", ReplaceWith:="", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="��", ReplaceWith:="�i�g���j", Replace:=wdReplaceAll)
       ActiveDocument.content.Select
   Selection.Style = ActiveDocument.Styles("00")
       
End Sub
Sub �N�D�ئ۰ʽs��()
  Dim oPara As Paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(A)") > 0 Then
     oPara.Previous.Range.Style = ActiveDocument.Styles("1")
    End If
  Next
End Sub

Sub ���ﶵ���˦�()
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
