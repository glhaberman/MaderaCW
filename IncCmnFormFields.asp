<% 
'------------------------------------------------------------------------------
' Name:     CommonFormFields
' Purpose:  This called from the Form section of each page to incorporate the
'           fields that need to be on every page.
'------------------------------------------------------------------------------
Sub CommonFormFields()
    Dim strTmp
    
    strTmp = Space(8) & "<INPUT TYPE=""hidden"" Name="
    Response.Write strTmp & """UserID"" VALUE=""" & gstrUserID & """ ID=UserID>" & vbCrLf
    Response.Write strTmp & """Password"" VALUE=""" & gstrPassword & """ ID=Password>" & vbCrLf
    Response.Write strTmp & """CalledFrom"" VALUE=""" & Request.Form("CalledFrom") & """ ID=CalledFrom>" & vbCrLf
    Response.Write strTmp & """ProgramsSelected"" VALUE=""" & Request.Form("ProgramsSelected") & """ ID=ProgramsSelected>" & vbCrLf
    
End Sub
%>
