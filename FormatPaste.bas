Attribute VB_Name = "FormatPaste"
'This macro is used to paste only the formatting of copied cells into selected cells.
'Import this macro into a module in your PERSONAL.XLSB workbook.
'Then you can assign this macro to a keyboard shortcut.
'NOTE: The clipboard is automatically cleared when running macros through the
'Macro Dialog Box (Alt + F8).  Run this macro by assigning it to an object,
'creating a keyboard shortcut, or in the VB Editor (shortcut F5).

Sub PasteFormat()
Attribute PasteFormat.VB_ProcData.VB_Invoke_Func = "V\n14"
'''Pastes the format of what is currently copied in the clipboard'''
    If Application.CutCopyMode = False Then
        Beep
    Else
        Application.Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    End If
End Sub
