<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncRptExpParms.asp                                               '
' Purpose: This include file contains code to package up report parameters  '
'          to be used in the export page.                                   '
'                                                                           '
'==========================================================================='
Dim mintExportI
Dim mstrExportSPName
Dim mstrExportParameterList
Dim mstrURL

mstrExportSPName = adCmd.CommandText
mintExportI = InStr(mstrExportSPName,"(?")
mstrExportSPName = Mid(mstrExportSPName, 8, mintExportI - 8)
mstrURL = Request.ServerVariables("SCRIPT_NAME")
mstrURL = Mid(mstrURL, InStrRev(mstrURL,"/") + 1, Len(mstrURL) - InStrRev(mstrURL,"/"))
If InStr(UCase(mstrURL),".ASP") > 0 Then
    mstrURL = Mid(mstrURL, 1, Len(mstrURL) - 4)
End If
mstrExportParameterList = ""

For mintExportI = 0 To adCmd.Parameters.Count - 1
    mstrExportParameterList = mstrExportParameterList & _
        adCmd.Parameters(mintExportI).Name & "^" & _
        adCmd.Parameters(mintExportI).Type & "^" & _
        adCmd.Parameters(mintExportI).Size & "^" & _
        adCmd.Parameters(mintExportI).Value & "|"
Next
%>

<SCRIPT ID=clientExportEventHandlersVBS LANGUAGE=vbscript>
Sub cmdExport_onclick()
    Dim strResults
    Dim strURL

    PageBody.style.cursor = "wait"
    
    strURL = "ReportExport.asp?RName=<% = mstrURL %>&SPName=<% = mstrExportSPName %>&Parameters=<% = mstrExportParameterList %>"
    strResults = window.showModalDialog(strURL)

    PageBody.style.cursor = "auto"
    
    MsgBox "Results copied to clipboard.", ,"Copy Results"    
End Sub
</SCRIPT>