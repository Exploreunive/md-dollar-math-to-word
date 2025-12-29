Attribute VB_Name = "模块11"
Option Explicit

Sub ConvertDollarLatexToWordEquations_Robust()
    Dim doc As Document
    Set doc = ActiveDocument

    Application.ScreenUpdating = False

    ' 1) 处理 $$...$$（块公式，允许多行）
    ProcessDoubleDollarBlocks_Robust doc

    ' 2) 处理 $...$（行内公式，允许换行）
    ProcessSingleDollarBlocks_Multiline doc

    ' 3) 全篇 BuildUp
    On Error Resume Next
    doc.OMaths.BuildUp
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "完成：已将 $...$ / $$...$$（含换行）转为 Word 公式。", vbInformation
End Sub


' =========================
' $$...$$（块公式，多行）
' =========================
Private Sub ProcessDoubleDollarBlocks_Robust(ByVal doc As Document)
    Dim story As Range, nxt As Range
    For Each story In doc.StoryRanges
        ConvertDoubleDollarInRange_Robust doc, story
        Set nxt = story.NextStoryRange
        Do While Not nxt Is Nothing
            ConvertDoubleDollarInRange_Robust doc, nxt
            Set nxt = nxt.NextStoryRange
        Loop
    Next story
End Sub

Private Sub ConvertDoubleDollarInRange_Robust(ByVal doc As Document, ByVal rngStory As Range)
    ConvertDollarPairsInRange doc, rngStory, "$$"
End Sub


' =========================
' $...$（行内公式，允许换行）
' =========================
Private Sub ProcessSingleDollarBlocks_Multiline(ByVal doc As Document)
    Dim story As Range, nxt As Range
    For Each story In doc.StoryRanges
        ConvertDollarPairsInRange doc, story, "$"
        Set nxt = story.NextStoryRange
        Do While Not nxt Is Nothing
            ConvertDollarPairsInRange doc, nxt, "$"
            Set nxt = nxt.NextStoryRange
        Loop
    Next story
End Sub


' =========================
' 核心：成对查找 $ 或 $$
' =========================
Private Sub ConvertDollarPairsInRange(ByVal doc As Document, ByVal rngStory As Range, ByVal token As String)
    Dim searchRng As Range
    Set searchRng = rngStory.Duplicate

    With searchRng.Find
        .ClearFormatting
        .Text = token
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    Do While searchRng.Find.Execute
        Dim startPos As Long
        startPos = searchRng.Start

        Dim endSearch As Range
        Set endSearch = rngStory.Duplicate
        endSearch.SetRange Start:=searchRng.End, End:=rngStory.End

        With endSearch.Find
            .ClearFormatting
            .Text = token
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
        End With

        If Not endSearch.Find.Execute Then Exit Do

        Dim endPos As Long
        endPos = endSearch.End

        Dim blockRng As Range
        Set blockRng = rngStory.Duplicate
        blockRng.SetRange Start:=startPos, End:=endPos

        Dim latex As String
        latex = blockRng.Text

        ' 去掉首尾 token
        latex = Mid$(latex, Len(token) + 1, Len(latex) - 2 * Len(token))

        ' 把换行压成空格（关键）
        latex = Replace(latex, vbCr, " ")
        latex = Replace(latex, vbLf, " ")
        latex = CollapseSpaces(Trim$(latex))

        blockRng.Text = latex
        MakeRangeMathAndBuildUp blockRng

        ' 继续
        searchRng.SetRange Start:=blockRng.End, End:=rngStory.End
        searchRng.Find.Text = token
    Loop
End Sub


' =========================
' 把 Range 变成 Word 公式
' =========================
Private Sub MakeRangeMathAndBuildUp(ByVal r As Range)
    On Error Resume Next
    r.OMaths.Add r
    If r.OMaths.Count > 0 Then
        r.OMaths(1).BuildUp
    End If
    On Error GoTo 0
End Sub


' =========================
' 压缩空白
' =========================
Private Function CollapseSpaces(ByVal s As String) As String
    Dim i As Long, out As String, prevSpace As Boolean
    out = ""
    prevSpace = False

    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = vbTab Then
            If Not prevSpace Then
                out = out & " "
                prevSpace = True
            End If
        Else
            out = out & ch
            prevSpace = False
        End If
    Next i
    CollapseSpaces = out
End Function

