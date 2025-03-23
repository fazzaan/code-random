' Insert a new row in a table with a keyboard shortcut. I assign this to Ctrl+Shift+Enter.
Sub Table_New_Row()
'
' Table_New_Row Macro
'
'
    Selection.InsertRowsBelow 1
End Sub

' Make a table autofit its contents and then autofit the width of the document. I assign this to Ctrl+Shift+Backspace.
Sub Table_Fit_Content_Wide()
'
' Table_Fit_Content_Wide Macro
'
'
    Dim table As table
    
    ' Check if there are any tables in the selection
    If Selection.Tables.Count > 0 Then
    
      ' Loop through each table in the selection
        For Each table In Selection.Tables
        
        ' Set to AutoFit Content
            table.AutoFitBehavior (wdAutoFitContent)
            DoEvents
            
            table.AutoFitBehavior (wdAutoFitWindow)
            DoEvents
        
        Next table
    Else
        MsgBox "No tables found in the current selection."
    End If
    
    Set table = Nothing ' Clean up object reference

End Sub

