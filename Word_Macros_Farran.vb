Sub Table_New_Row()
'
' Table_New_Row Macro
'
'
    Selection.InsertRowsBelow 1
End Sub
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

