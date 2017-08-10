<% 
'------------------------------------------------------------------------------
' Name:     WriteFormField
' Purpose:  Used to build the HTML fields of an HTML form.  It receives the
'           name of the field and writes the correct HTML string back out to 
'           the response object.
'------------------------------------------------------------------------------
Sub WriteFormField(strFieldName, strFieldValue)
    Dim strPc1
    Dim strPc2
    Dim strPc3
    Dim strPc4
    
    strPc1 = Space(8) & "<INPUT TYPE=""hidden"" Name="""
    strPc2 = """ VALUE="""
    strPc3 = """ ID="
    strPc4 = ">"

    Response.Write strPc1 & strFieldName & strPc2 & strFieldValue & strPc3 & strFieldName & strPc4 & vbCrLf

End Sub
%>
