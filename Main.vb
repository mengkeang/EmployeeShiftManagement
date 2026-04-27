Option Explicit

Sub btnArchiveData_Click()
    
    ' Disable Screen Update to boost performance
    Application.ScreenUpdating = False
    
    ' --------------------------------------------------------------------------------
    ' Step: Copy Column: Sun to Column: LW_Sun
    ' --------------------------------------------------------------------------------
    ' This is for Clopen Detection Conditional Formatting
    Dim tblSchedule As ListObject
    Set tblSchedule = SheetSchedule.ListObjects("tblSchedule")
    
    ' Copy Data from Column: Sun to Column: LW_Sun
    tblSchedule.ListColumns("Sun").DataBodyRange.Copy
    tblSchedule.ListColumns("LW_Sun").DataBodyRange.PasteSpecial xlPasteValues
    
    ' End Copy and Paste
    Application.CutCopyMode = False
    
    ' --------------------------------------------------------------------------------
    ' Step: Refresh tblWeeklyData
    ' --------------------------------------------------------------------------------
    Dim tblWeeklyData As ListObject
    Set tblWeeklyData = SheettblWeeklyData.ListObjects("tblWeeklyData")
    
    ' Refresh data on Sheet tblWeeklyData via Power Query
    tblWeeklyData.Refresh

    ' --------------------------------------------------------------------------------
    ' Step: Copy Data from tblWeeklyData and Append to tblHistory
    ' --------------------------------------------------------------------------------
    Dim tblHistory As ListObject
    Set tblHistory = SheettblHistory.ListObjects("tblHistory")
    
    ' Add new row to tblHistory
    ' This is the row where will paste the data from tblWeeklyData
    Dim newRow As ListRow
    Set newRow = tblHistory.ListRows.Add
    
    ' Copy data from Sheet tblWeeklyData
    ' Paste data on Sheet tblHistory > New Row
    tblWeeklyData.DataBodyRange.Copy
    newRow.Range(tblHistory.ListColumns("StaffName").Index).PasteSpecial xlPasteValues
    
    ' End Copy and Paste
    Application.CutCopyMode = False
    ' --------------------------------------------------------------------------------
    
    ' Show successful message
    MsgBox "Action Completed!", vbInformation, "AbcFunWithExcel"
    
    ' Enable Screen Update after we complete our task
    Application.ScreenUpdating = True
    
End Sub

Sub btnNewStaff_Click()
    ' Disable Screen Update to boost performance
    Application.ScreenUpdating = False
    
    ' --------------------------------------------------------------------------------
    ' Step: Add New Row to tblStaff
    ' --------------------------------------------------------------------------------
    Dim tblStaff As ListObject
    Set tblStaff = SheetStaff.ListObjects("tblStaff")
    
    ' Add a new row on tblStaff
    Dim newRow As ListRow
    Set newRow = tblStaff.ListRows.Add
    
    ' Select Column: StaffName of the newly inserted row
    newRow.Range(tblStaff.ListColumns("StaffName").Index).Select
    
    ' --------------------------------------------------------------------------------
    ' Step: Add New Row to tblSchedule and Insert an empty row below the newly added row
    ' --------------------------------------------------------------------------------
    Dim tblSchedule As ListObject
    Set tblSchedule = SheetSchedule.ListObjects("tblSchedule")
    
    ' Add a new row on tblSchedule
    Dim newRowSchedule As ListRow
    Set newRowSchedule = tblSchedule.ListRows.Add
    
    ' Find the Row Number of the newly inserted row
    '   Plus 2 so that there are gaps between tblSchedule and staff head count summary section
    Dim nextRow As Long
    nextRow = newRowSchedule.Range().Row + 2
    
    ' Insert a row at index: nextRow
    SheetSchedule.Range("A" & nextRow).EntireRow.Insert
    
    ' Enable Screen Update after we complete our task
    Application.ScreenUpdating = True
End Sub
