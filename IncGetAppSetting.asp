<%
'------------------------------------------------------------------------------
'       Name: GetAppSetting()
'    Purpose: This procedure is used to look up an application setting value
'             the setting name parameter passed to the function - from the 
'             table tblApplicationSettings.
'             The page that includes this function must already have defined 
'             and set a connection object for gadoCon.
' Programmer: Brian Wieczorek 10/9/3
'    Updated: 
'------------------------------------------------------------------------------
Function GetAppSetting(strSettingName)
    Dim adCmd
    
    Set adCmd = Server.CreateObject("ADODB.Command")
    With adCmd
        Set .ActiveConnection = gadoCon
        .CommandText = "spGetAppSetting"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@SettingName", adVarChar, adParamInput, 50, strSettingName)
        .Parameters.Append .CreateParameter("@SettingValue", adVarChar, adParamOutput, 255, Null)
        .Execute
        If IsNull(.Parameters("@SettingValue").Value) Then
            GetAppSetting = ""
        Else
            GetAppSetting = .Parameters("@SettingValue").Value
        End If
    End With
    Set adCmd = Nothing
End Function
%>
