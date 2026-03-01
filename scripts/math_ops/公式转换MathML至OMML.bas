Attribute VB_Name = "公式转换MathML至OMML"
Sub PlainMathMLToEquation()
    Dim xmlns As String, mathText As String
    xmlns = "xmlns=""http://www.w3.org/1998/Math/MathML"""

    ' set up find parameters
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "\<math?*\</math\>^13"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With

    Debug.Print "-------------<Start>-------------------"
    Selection.SetRange 0, 0

    While Selection.Find.Execute()
        mathText = Selection.Text

        ' 1) 复制到剪贴板（带重试）
        If CopyPlainText(mathText, 20, 0.06) Then
            ' 2) 删除原文
            Selection.Delete

            ' 3) 短等待（避免立刻粘贴时剪贴板状态未稳定）
            Waiting 0.05

            ' 4) 粘贴（带重试）
            If Not PastePlainText(20, 0.06) Then
                ' 粘贴失败时回退原文本，避免内容丢失
                Selection.TypeText Text:=mathText
                Debug.Print "[PasteFail] 已回退原文本。"
            End If
        Else
            ' 复制失败时跳过当前匹配，避免死循环
            Debug.Print "[CopyFail] 跳过一次匹配。"
            Selection.Collapse wdCollapseEnd
        End If
    Wend

    RemoveMathTypeTailsExact
    Debug.Print "-------------<Finished>-------------------"
End Sub

Sub Waiting(Optional ByVal seconds As Double = 0.95)
    If seconds <= 0 Then Exit Sub

    Dim t As Double
    t = Timer

    Do
        DoEvents
        ' 跨午夜保护：Timer 会在 0 点归零
        If Timer < t Then t = t - 86400#
    Loop While Timer - t < seconds
End Sub

Function CopyPlainText(ByVal srcText As String, _
                       Optional ByVal retryCount As Long = 20, _
                       Optional ByVal waitSeconds As Double = 0.06) As Boolean
    Dim i As Long
    Dim dobj As Object

    For i = 1 To retryCount
        On Error Resume Next
        Set dobj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dobj.SetText srcText
        dobj.PutInClipboard

        If Err.Number = 0 Then
            CopyPlainText = True
            On Error GoTo 0
            Exit Function
        End If

        Debug.Print "[CopyRetry " & i & "] Err=" & Err.Number & " " & Err.Description
        Err.Clear
        Set dobj = Nothing
        On Error GoTo 0

        Waiting waitSeconds
    Next i
End Function

Function PastePlainText(Optional ByVal retryCount As Long = 20, _
                        Optional ByVal waitSeconds As Double = 0.06) As Boolean
    Dim i As Long

    For i = 1 To retryCount
        On Error Resume Next
        Selection.PasteSpecial Link:=False, DataType:=wdPasteText, _
                               Placement:=wdInLine, DisplayAsIcon:=False

        If Err.Number = 0 Then
            PastePlainText = True
            On Error GoTo 0
            Exit Function
        End If

        Debug.Print "[PasteRetry " & i & "] Err=" & Err.Number & " " & Err.Description
        Err.Clear
        On Error GoTo 0

        Waiting waitSeconds
    Next i
End Function

Sub RemoveMathTypeTailsExact()
    ReplaceExactInAllStories "<!-- MathType@Translator@5@5@MathML2 (namespace attr).tdl@MathML 2.0 (namespace attr)@ -->"
    ReplaceExactInAllStories "<!-- MathType@End@5@5@ -->"
End Sub

Sub ReplaceExactInAllStories(ByVal findText As String)
    Dim rng As Range
    Set rng = ActiveDocument.StoryRanges(wdMainTextStory)

    Do While Not rng Is Nothing
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = findText
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
        End With
        rng.Find.Execute Replace:=wdReplaceAll
        Set rng = rng.NextStoryRange
    Loop
End Sub

