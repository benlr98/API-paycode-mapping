Attribute VB_Name = "HandleRefreshToken"
Option Explicit
Const DblQuote = """"



Function GetKey(ByVal sBuffer As String, ByVal sKey As String) As String
    
    Dim iKeyLen As Long
    Dim sValue As String
    Dim iLast(1) As Long
    Dim iStart As Long
    Dim sTemp As String
    Dim iEnd As Long
    
    GetKey = ""
    sKey = sKey & DblQuote
    iKeyLen = Len(sKey)
    iStart = InStr(sBuffer, sKey)
    
    If iStart > 0 Then
        iStart = iStart + iKeyLen + 1
        
        If mId(sBuffer, iStart, 1) = DblQuote Then
            iEnd = InStr(iStart + 1, sBuffer, DblQuote)
        Else
            iLast(0) = InStr(iStart + 1, sBuffer & ",", ",")
            iLast(1) = InStr(iStart + 1, sBuffer & "}", "}")
            iEnd = Application.Min(iLast)
        End If
        
        sTemp = mId(sBuffer, iStart, iEnd - iStart)
        GetKey = Replace(sTemp, """", "")
    End If
    
End Function

Function hasAccessExpired() As Boolean
    Dim expirationTime As Date
    Dim testExpirationTime As String
    Dim testCurrentTime As String
    
    hasAccessExpired = False
    
    expirationTime = Sheets("WFM Paycodes Table").Cells(13, 10)
    
    ' Access Token has expired
    If now > expirationTime Then
        Call MsgBox("Please refresh access token")
        hasAccessExpired = True
    End If

End Function

Sub RefreshWFD_accesKey()

    Dim oXmlHttp As MSXML2.XMLHTTP60
    Set oXmlHttp = New MSXML2.XMLHTTP60

    Dim ServiceURL As String
    Dim appkey As String
    Dim client_id As String
    Dim client_secret As String
    Dim refresh_token As String
    Dim response As String
    Dim iNumSeconds As Integer
    
    Sheets("WFM Paycodes Table").Cells(13, 10) = ""
    
    ServiceURL = Sheets("WFM Paycodes Table").Cells(7, 10).Text
    appkey = Sheets("WFM Paycodes Table").Cells(10, 10).Value
    refresh_token = Sheets("WFM Paycodes Table").Cells(12, 10).Value
    client_id = Sheets("WFM Paycodes Table").Cells(8, 10).Value
    client_secret = Sheets("WFM Paycodes Table").Cells(9, 10).Value
    
    ServiceURL = ServiceURL & "/api/authentication/access_token" & _
                "?refresh_token=" & refresh_token & _
                "&client_id=" & client_id & _
                "&client_secret=" & client_secret & _
                "&grant_type=refresh_token&auth_chain=OAuthLdapService"

    oXmlHttp.Open "POST", ServiceURL, False, "", ""
    oXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    oXmlHttp.setRequestHeader "appkey", appkey

    oXmlHttp.send           'Body is optional
    
    response = oXmlHttp.responseText
    
    Sheets("WFM Paycodes Table").Cells(11, 10) = GetKey(response, "access_token")
    
    If GetKey(response, "error_description") <> "" Then
      Call MsgBox(GetKey(response, "error"))
    Else
        iNumSeconds = GetKey(response, "expires_in")
        Sheets("WFM Paycodes Table").Cells(13, 10) = now + (iNumSeconds / 86400)
        Call MsgBox("New token expires at: " & vbCrLf & now + (iNumSeconds / 86400))
    End If
    
End Sub

