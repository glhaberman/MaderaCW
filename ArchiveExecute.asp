<%@ LANGUAGE="VBScript" %>
<%Option Explicit%> 
<!--#include file="IncCnn.asp"-->
<%
Dim madoCmd
Dim mstrAction
Dim mstrError

Server.ScriptTimeout = 360
If Request.QueryString("UserID") <> "" Then
    Set madoCmd = GetAdoCmd("spArchiveReviews")
        AddParmIn madoCmd, "@LoginID", adVarchar, 255, Request.QueryString("UserID")
        AddParmIn madoCmd, "@ArchiveDate", adDBTimeStamp, 0, Request.QueryString("ArchiveDate")

    madoCmd.CommandTimeout = 360
    On Error Resume Next
    Call madoCmd.Execute
    If Err.number <> 0 Then
        mstrError = Err.Description
    End If
    On Error Goto 0
    
End If
%>
<HTML>
<HEAD>
    <META name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE>
        Execute Command
    </TITLE>
</HEAD>
<SCRIPT LANGUAGE="vbscript">
Sub window_onload
    If "<%=Request.QueryString("UserID")%>" <> "" And "<%=mstrError%>" = "" Then
        window.parent.document.all("PageFrame").disabled = False
        window.parent.document.all("PageFrame").style.cursor = "default"
        window.parent.document.all("lblExecuting").innerText = "Archiving has completed."
        window.parent.document.all("divExecute").style.left = "-1000"
        window.close
    ElseIf "<%=mstrError%>" <> "" Then
        If InStr("<%=mstrError%>","Timeout") > 0 Then
            divMessage.innerText = "Maximum time alloted for the archive to run has expired.  The archive has failed.  Please enter a date that is older to reduce the amount of records to archive or contact The Rushmore Group to set up an archive procedure to be performed by IT."
            window.parent.document.all("PageFrame").disabled = False
            window.parent.document.all("PageFrame").style.cursor = "default"
            window.parent.document.all("lblExecuting").innerText = "Error"
            PageBody.style.cursor = "default"
        Else
            divMessage.innerText = "Archiving Failed:  <%=mstrError%>"
            window.parent.document.all("PageFrame").disabled = False
            window.parent.document.all("PageFrame").style.cursor = "default"
            window.parent.document.all("lblExecuting").innerText = "Error"
            PageBody.style.cursor = "default"
        End If
    End If
End Sub
</SCRIPT>

<BODY id="PageBody" style="cursor:wait">
    <div id=divMessage>
        Executing...
    </div>
</BODY>
</HTML>
<%
%>
