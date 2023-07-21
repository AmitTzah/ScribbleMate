# ScribbleMate

It is recommended to use following VBA Word macros when working with this add-in. These Macros will allow you to select text quickly:

Sub SelectXParagraphsBackward1()
    Dim NumParagraphs As Integer
    NumParagraphs = 10 ' Change this number to the desired number of paragraphs to select backward from cursor position
    
    Dim i As Integer
    Application.ScreenUpdating = False ' Turn off screen updating to prevent the view from jumping
    
    For i = 1 To NumParagraphs
        Selection.MoveUp Unit:=wdParagraph, Count:=1, Extend:=wdExtend
    Next i
    
    Selection.Document.ActiveWindow.ScrollIntoView Selection.Characters(Selection.Characters.Count), True
    
    Application.ScreenUpdating = True ' Turn on screen updating
End Sub


Sub SelectXParagraphsBackward2()
    Dim NumParagraphs As Integer
    NumParagraphs = 10 ' Change this number to the desired number of paragraphs to select backward from cursor position
    
    Dim i As Integer
    
    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    Application.ScreenUpdating = False ' Turn off screen updating to prevent the view from jumping
   
    For i = 1 To NumParagraphs - 1
'        Selection.MoveUp Unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Set rng = rng.Paragraphs(1).Previous.Range
    Next i

    rng.End = Selection.Start
    rng.Select
    
    Application.ScreenUpdating = True ' Turn on screen updating
End Sub

Sub SelectParagraphsBackwardUntilHeading()
    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    Application.ScreenUpdating = False ' Turn off screen updating to prevent the view from jumping

    Do While Not rng.Paragraphs.First.Style Like "Heading*"
        If rng.Start <= 1 Then Exit Do ' Exit loop if the beginning of the document is reached
        Set rng = rng.Paragraphs(1).Previous.Range
    Loop

    rng.End = Selection.Start
    rng.Select

    Application.ScreenUpdating = True ' Turn on screen updating
End Sub

Sub SelectParagraphsBackwardUntilSceneBreak()
    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    Application.ScreenUpdating = False ' Turn off screen updating to prevent the view from jumping
    
    Dim paraText As String
    
    Do
        paraText = Trim(rng.Paragraphs.First.Range.Text)
        paraText = Replace(paraText, vbCr, "") ' Remove hidden carriage return characters
        paraText = Replace(paraText, vbLf, "") ' Remove hidden line feed characters

        If paraText = "***" Then Exit Do

        If rng.Start <= 1 Then Exit Do ' Exit loop if the beginning of the document is reached
        Set rng = rng.Paragraphs(1).Previous.Range
    Loop

    rng.End = Selection.Start
    rng.Select

    Application.ScreenUpdating = True ' Turn on screen updating
End Sub


