<%
Function WriteOptionList(strListName)
	Dim adCmd
    Dim adRs

	Set adCmd = GetAdoCmd("spGetOptListValues")
        AddParmIn adCmd, "@LstName", adVarChar, 50, strListName
    Set adRs = GetAdoRs(adCmd)
	Do While Not adRs.EOF
		Response.Write adRs.Fields("OptionValue").Value
		adRs.MoveNext 
	Loop 
	adRs.Close
	Set adRs = Nothing
	Set adCmd = Nothing
End Function
%>