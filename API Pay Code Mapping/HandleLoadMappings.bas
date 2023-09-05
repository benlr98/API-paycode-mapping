Attribute VB_Name = "HandleLoadMappings"

Sub MappingAPIRequestByRow()
    Dim tbl As ListObject
    Dim rng As Range
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Set wb = ThisWorkbook
    Set DDUSheet = wb.Worksheets("DDU Load")
    Set tbl = DDUSheet.ListObjects("DDU_Import")
    Dim expirationDate As Date
    Dim userConfirmation As String
    
    
    ' Handle expired refreshToken
    expirationDate = Sheets("WFM Paycodes Table").Cells(13, 10).Value
    Dim now As Date
    
    now = DateTime.now
    If now >= expirationDate Or IsEmpty(expirationDate) Then
        MsgBox "Please refresh access token"
        Exit Sub
    End If
    
    ' Confirm with user about making API requests
    userConfirmation = ShowConfirmation()
    If userConfirmation = "abort" Then
        Exit Sub
    End If
    
    For Each Row In tbl.ListRows
        Dim id As String
        Dim name As String
        Dim description As String
        Dim includes As Variant
        Dim PaycodeJsonString As String
        
        id = Row.Range.Cells(1).Value
        name = Row.Range.Cells(2).Value
        description = Row.Range.Cells(3).Value
        includes = Split(Row.Range.Cells(4).Value, ",")
        excludes = Split(Row.Range.Cells(5).Value, ",")
        
        ' IMPORTANT to clear string for each mapping category
        PaycodeJsonString = ""
        
        For Each paycodeName In includes
            ' append paycode string with comma at the end for each paycode
            ' If found in excludes, create a cost only json string
            If IsInArray(paycodeName, excludes) Then
                PaycodeJsonString = PaycodeJsonString & BuildAttributesJson(paycodeName, True) & ","
            Else
                ' Else create a normal paycode json string
                PaycodeJsonString = PaycodeJsonString & BuildAttributesJson(paycodeName, False) & ","
            End If
        Next paycodeName
        
        If name <> "" And description <> "" Then
        
            ' Remove the last comma from the paycode array
            PaycodeJsonString = Left$(PaycodeJsonString, Len(PaycodeJsonString) - 1)
            
            ' build full json string for mapping category
            ' Json string is different between PUT and POST => (includes "id": "123")
            If Row.Range.Cells(1).Value = "" Then
                Json = BuildMappingJson("POST", id, name, description, PaycodeJsonString)
                Row.Range.Cells(6).Value = MappingAPIRequest(Json, id)
            Else
                Json = BuildMappingJson("PUT", id, name, description, PaycodeJsonString)
                Row.Range.Cells(6).Value = MappingAPIRequest(Json, id)
            End If
        Else
            Row.Range.Cells(6).Value = "Must have both Name and Description values filled out"
        End If
        
        
    Next Row
    
End Sub

Function ShowConfirmation() As String
    ' Display a confirmation window with Yes and No buttons
    Dim response As Integer
    Dim tenant As String
    
    tenant = Sheets("WFM Paycodes Table").Cells(7, 10).Value
    response = MsgBox("Are you sure you want to proceed with posting mappings to " & tenant & "?", vbQuestion + vbYesNo, "Confirmation")
    
    ' Check the user's response
    If response = vbYes Then
        ' User clicked Yes, proceed with the action
        ' Add your code here for the action you want to perform
        ShowConfirmation = "proceed"
    Else
        ' User clicked No, cancel the action
        MsgBox "Action canceled!"
        ' Add any additional code to handle the cancellation
        ShowConfirmation = "abort"
    End If
End Function


Function BuildMappingJson(ByVal method As String, id As String, name As String, description As String, PaycodeJsonString As String)
    Dim Json As Object
    ' MappingCategoryAttributes is set in while parsing with converter, set to PaycodeJsonString
    ' Is is set with string concatenation on next line
    Set Json = JsonConverter.ParseJson("{'name':'zzBrTest','description':'zzBrTest','mappingCategoryType':{'id':1,'name':'PAYCODE','description':'Paycode mapping category type'},'mappingCategoryAttributes':[" & PaycodeJsonString & "]}")
    If method = "PUT" And id <> "" Then
        Json("id") = id
    End If
    Json("name") = name
    Json("description") = description
    
    BuildMappingJson = JsonConverter.ConvertToJson(Json)
End Function

Public Function BuildAttributesJson(ByVal paycodeName As String, costOnly As Boolean) As String
    Set Json = JsonConverter.ParseJson("{'name':'REPLACE','attributes':[{'id':0,'name':'PayCodeId'}],'customAttributes':[{'customAttributeValue':'REPLACEBOOLEAN','customAttributeCtx':{'id':1,'name':'Cost Only','dataType':{'id':4,'name':'BOOLEAN','description':'Boolean Data Type'},'attributeType':{'id':2}}}]}")
    Json("name") = paycodeName
    Json("customAttributes")(1)("customAttributeValue") = costOnly
    
    BuildAttributesJson = JsonConverter.ConvertToJson(Json)
End Function

Public Function IsInArray(ByVal stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Public Function MappingAPIRequest(ByVal reqBody As String, mId As String)
    Dim req As New MSXML2.ServerXMLHTTP60
    
    Dim ServiceURL As String
    Dim reqUrl As String
    Dim appkey As String
    Dim token As String
    Dim response As String
    Dim endPoint As String
    Dim method As String
    
    If mId = "" Then
        method = "POST"
    Else
        method = "PUT"
    End If
    
    ServiceURL = Sheets("WFM Paycodes Table").Cells(7, 10).Text
    appkey = Sheets("WFM Paycodes Table").Cells(10, 10).Value
    token = Sheets("WFM Paycodes Table").Cells(11, 10).Value
    ' don't forget forward slash at the end of endpoint
    endPoint = "/api/v1/platform/analytics/mapping_categories/"
    reqUrl = ServiceURL & endPoint & mId
    
    req.Open method, reqUrl, False
    req.setRequestHeader "Appkey", appkey
    req.setRequestHeader "Authorization", token
    req.setRequestHeader "Content-Type", "application/json"
    
    req.send reqBody
    response = req.responseText
    
    MappingAPIRequest = response
    
End Function


