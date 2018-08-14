Attribute VB_Name = "Module1"
Option Explicit



Function findMatch(selectedRange As Range, pattern As String, Optional number As Variant) As String
    
    Dim regEx As New RegExp
        
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = pattern
    End With

    If regEx.Test(selectedRange.Value) Then
        
        Dim matches As MatchCollection
        Set matches = regEx.Execute(selectedRange.Value)
        
        If IsMissing(number) Then
            findMatch = matches.Item(0).Value
        Else
            findMatch = matches.Item(number - 1).Value
        End If
        
    End If
    
End Function


Sub registerDescription()
    Application.MacroOptions Macro:="findMatch", Description:="findMatch(selectedRange As Range, pattern As String, Optional number As Variant) As String", Category:=10
End Sub

