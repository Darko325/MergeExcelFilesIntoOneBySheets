Attribute VB_Name = "Module1"
Sub MergeExcelFilesInFolder()
    Dim FolderPath As String
    Dim Filename As String
    Dim wbSource As Workbook
    Dim wbDestination As Workbook
    Dim ws As Worksheet
    
    ' Set the folder path where the Excel files are located
    FolderPath = "C:\Users\gis\Desktop\CVD_DEATHS_2020\"
    
    ' Create a new workbook to merge the files
    Set wbDestination = Workbooks.Add
    
    ' Disable updates and events to improve performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Loop through all files in the folder
    Filename = Dir(FolderPath & "*.xlsx")
    Do While Filename <> ""
        ' Open each source workbook
        Set wbSource = Workbooks.Open(FolderPath & Filename)
        
        ' Copy each sheet from the source workbook to the destination workbook
        For Each ws In wbSource.Worksheets
            ws.Copy After:=wbDestination.Sheets(wbDestination.Sheets.Count)
        Next ws
        
        ' Close the source workbook without saving changes
        wbSource.Close SaveChanges:=False
        
        ' Get the next file in the folder
        Filename = Dir
    Loop
    
    ' Enable updates and events
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Save and close the destination workbook
    wbDestination.SaveAs FolderPath & "MergedFile.xlsx"
    wbDestination.Close SaveChanges:=False
    
    MsgBox "Files merged successfully!"
End Sub

