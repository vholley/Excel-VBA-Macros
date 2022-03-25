Attribute VB_Name = "ColorPicker"
'These macros are used to change the highlight color of selected cells.
'Import these macros into a module in your PERSONAL.XLSB workbook.
'Then you can assign these macros to keyboard shortcuts.

' This macro removes all formatting (colors, highlighting, text styles, etc.) from the selected cells.
Sub ClearFormat()
Attribute ClearFormat.VB_ProcData.VB_Invoke_Func = "C\n14"
    Selection.ClearFormats
End Sub

Sub ColorPickerStandard() 'Standard
Attribute ColorPickerStandard.VB_ProcData.VB_Invoke_Func = "D\n14"

If Selection.Interior.ColorIndex = xlColorIndexNone Then
    Selection.Interior.ColorIndex = 6 'Yellow
ElseIf Selection.Interior.ColorIndex = 6 Then
    Selection.Interior.ColorIndex = 4 'Green
ElseIf Selection.Interior.ColorIndex = 4 Then
    Selection.Interior.ColorIndex = 3 'Red
ElseIf Selection.Interior.ColorIndex = 3 Then
    Selection.Interior.ColorIndex = 8 'Cyan
ElseIf Selection.Interior.ColorIndex = 8 Then
    Selection.Interior.ColorIndex = 7 'Magenta
ElseIf Selection.Interior.ColorIndex = 7 Then
    Selection.Interior.ColorIndex = 45 'Orange
ElseIf Selection.Interior.ColorIndex = 45 Then
    Selection.Interior.ColorIndex = 5 'Blue
Else:
    Selection.Interior.ColorIndex = xlColorIndexNone
End If


End Sub

Sub ColorPickerPastel() 'Pastel
Attribute ColorPickerPastel.VB_ProcData.VB_Invoke_Func = "S\n14"

If Selection.Interior.ColorIndex = xlColorIndexNone Then
    Selection.Interior.ColorIndex = 36
ElseIf Selection.Interior.ColorIndex = 36 Then
    Selection.Interior.ColorIndex = 35
ElseIf Selection.Interior.ColorIndex = 35 Then
    Selection.Interior.ColorIndex = 38
ElseIf Selection.Interior.ColorIndex = 38 Then
    Selection.Interior.ColorIndex = 34
ElseIf Selection.Interior.ColorIndex = 34 Then
    Selection.Interior.ColorIndex = 39
ElseIf Selection.Interior.ColorIndex = 39 Then
    Selection.Interior.ColorIndex = 40
ElseIf Selection.Interior.ColorIndex = 40 Then
    Selection.Interior.ColorIndex = 37
Else:
    Selection.Interior.ColorIndex = xlColorIndexNone
End If


End Sub

Sub ColorPickerGreys()
Attribute ColorPickerGreys.VB_ProcData.VB_Invoke_Func = "G\n14"
If Selection.Interior.ColorIndex = xlColorIndexNone Then
    Selection.Interior.ColorIndex = 15
ElseIf Selection.Interior.ColorIndex = 15 Then
    Selection.Interior.ColorIndex = 48
ElseIf Selection.Interior.ColorIndex = 48 Then
    Selection.Interior.ColorIndex = 16
ElseIf Selection.Interior.ColorIndex = 16 Then
    Selection.Interior.ColorIndex = 56
ElseIf Selection.Interior.ColorIndex = 56 Then
    Selection.Interior.ColorIndex = 1
Else:
    Selection.Interior.ColorIndex = xlColorIndexNone
End If

End Sub


