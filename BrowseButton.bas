Attribute VB_Name = "BrowseButton"
Sub OpenDialog(loc As String)
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    dialogBox.AllowMultiSelect = False
    
    dialogBox.Title = "Select a file"
    
'    dialogBox.InitialFileName = Range("A12").Value
    
    dialogBox.Filters.Clear
    
    dialogBox.Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls"
    
    If dialogBox.Show = -1 Then
        ActiveSheet.Range(loc).Value = dialogBox.SelectedItems(1)
    End If
End Sub

Sub TemplateFile()
    OpenDialog ("B1")
End Sub

Sub M0File()
    OpenDialog ("B6")
End Sub

Sub SourceFile()
    OpenDialog ("B11")
End Sub


