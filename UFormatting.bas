' Requirements taken from:
' http://www.modeloff.com/wp-content/uploads/2013/10/ModelOff-2013-Round-1-Question-Breakdown-and-Style-Guide.pdf

Sub normal_numbers()
' One decimal place
' Contain thousands separators
' Negative numbers should appear in parentheses
' Font colour should be black

    ' Number format copied from document
    Selection.NumberFormat = "#,##0.0_);(#,##0.0);0.0_);@_)"
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With

End Sub

Sub percentages()
' One decimal place
' Negative number in parentheses

    ' Number format copied from document
    Selection.NumberFormat = "0.0%_);(0.0%)"

End Sub

Sub assumption_inputs()
' Blue font color (0, 0, 255)
' Outline cell border
' Yellow fill colour (255, 255, 204)

    ' Border
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    ' Font colour
    With Selection.Font
        .Color = -65536
        .TintAndShade = 0
    End With

    ' Fill colour
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Sub non_assumption_inputs()
' Blue font colour (0, 0, 255)

    ' Font colour
    With Selection.Font
        .Color = -65536
        .TintAndShade = 0
    End With
    
    ' No fill colour
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' No border
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub
