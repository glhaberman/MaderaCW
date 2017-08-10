<%
Function BuildList(strListName, strListTitle, intLeft, intTop, intWidth)
	Dim adCmd
    Dim adRs
    Dim strOptions

	If intLeft <> 0 Then
		Response.Write "<SPAN id=lbl" & strListName & " class=DefLabel style=""LEFT:" & intLeft & "; TOP:" & intTop & """>" & vbCrLf
		Response.Write strListTitle
		Response.Write "</SPAN>" & vbCrLf

		Response.Write "<SELECT id=cbo" & strListName & " title=""" & GetAppSetting("TitleCaseWorker") & """ style=""WIDTH:" & intWidth & "; LEFT:" & intLeft & "; TOP:" & intTop + 15 & """ tabIndex=1>" & vbCrLf
		Response.Write "<OPTION VALUE=0 SELECTED>" & vbCrLf
	End IF

	Set adCmd = Server.CreateObject("ADODB.Command")
	Set adRs = Server.CreateObject("ADODB.Recordset")
	With adCmd
		.ActiveConnection = gadoCon
		.CommandType = adCmdStoredProc
		.CommandText = "spListValuesGet"
		.Parameters.Append .CreateParameter("@LstName", adVarChar, adParamInput, 50, strListName)
	End With

	adRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
	strOptions = ""
	Do While Not adRs.EOF
		strOptions = strOptions & "<OPTION VALUE=" & adRs.Fields(0).Value & ">" & adRs.Fields(1).Value & "</OPTION>"
		adRs.MoveNext 
	Loop 
	adRs.Close
	Set adRs = Nothing
	Set adCmd = Nothing
	If intLeft <> 0 Then
		Response.Write strOptions
		Response.Write vbCrLf & "</SELECT>" & vbCrLf
	End If
	BuildList = strOptions
	
End Function
%>