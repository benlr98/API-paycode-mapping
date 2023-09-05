Attribute VB_Name = "HandleTableRefresh"


Sub Run_Mapping_Updates()
Attribute Run_Mapping_Updates.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Run_Mapping_Updates Macro
'
    If hasAccessExpired Then
        Exit Sub
    End If
    ActiveWorkbook.Connections("Query - DDU Import(1)").Refresh
    ActiveWorkbook.Connections("Query - Data Hub Import").Refresh

End Sub


Sub Validation_Update()
'
' Run_Mapping_Updates Macro
'
    If hasAccessExpired Then
        Exit Sub
    End If
    ActiveWorkbook.Connections("Query - Validation - Main").Refresh

End Sub



Sub Update_Mapping_Review()
'
' Update_Mapping_Review Macro
'
    If hasAccessExpired Then
        Exit Sub
    End If
    
    ActiveWorkbook.Connections("Query - Validation - Compare Mapping").Refresh
    
End Sub


Sub Get_Paycodes_Table()
'
' Get_Paycodes_Table
'

    ' Call MsgBox("hasAccessExpired: " & hasAccessExpired)
    
    If hasAccessExpired = True Then
        Exit Sub
    End If

    
    ActiveWorkbook.Connections("Query - WFM Paycodes Table").Refresh
    
End Sub

Sub Get_Dataview_Profiles_Table()
'
' Get_Dataview_Profiles_Table
'
    
    If hasAccessExpired Then
        Exit Sub
    End If
    
    ActiveWorkbook.Connections("Query - WFM Dataview Profiles").Refresh
    
End Sub


Sub Get_Report_Profiles_Table()
'
' Get_Report_Profiles_Table
'
    If hasAccessExpired Then
        Exit Sub
    End If
    
    ActiveWorkbook.Connections("Query - WFM Report Profiles").Refresh
    
End Sub

Sub Get_Location_Types_Table()
'
' Get_Location_Types_Table
'
    If hasAccessExpired Then
        Exit Sub
    End If
    
    ActiveWorkbook.Connections("Query - WFM Location Types").Refresh
    
End Sub

Sub Get_Dataviews_and_RDOs_Table()
'
' Get_Dataviews_and_RDOs_Table
'
    If hasAccessExpired Then
        Exit Sub
    End If
    
    ActiveWorkbook.Connections("Query - WFM Dataviews and RDOs(1)").Refresh
    
End Sub


Sub CopySheetsToNamedWorkbook()
    Dim newWorkbook As Workbook
    Dim saveName As Variant
    
    ' Create a new workbook and copy the sheets to it
    Set newWorkbook = Workbooks.Add
    ThisWorkbook.Sheets(Array("Analytics Pay Code Mapping", "Lookups")).Copy Before:=newWorkbook.Sheets(1)
    
    'Remove the button from the new workbook
    newWorkbook.Sheets("Analytics Pay Code Mapping").Shapes("Button 1").Delete
    
    'Prompt the user to enter the file name
    Dim fileName As String
    custName = InputBox("Enter the customer name:")
    If custName = "" Then Exit Sub ' Exit the macro if the user cancels or enters a blank name
    
    'Combine the file name with the current directory path to create the full save path
    Dim savePath As String
    savePath = ThisWorkbook.Path & "\" & custName & " - Analytics Pay Code Mapping" & ".xlsx"
    
    'Save and close the new workbook
    newWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWorkbook.Close SaveChanges:=False
    
End Sub

