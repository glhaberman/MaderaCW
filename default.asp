<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: default.asp                                                     '
'  Purpose: This is the initial page from which the application window is   '
'   launched.  After the application launches, this page remains up only    '
'   to avoid the message box that occurs when attempting to close in code.  '
'==========================================================================='
'If Request.ServerVariables("SERVER_PORT_SECURE") <> "1" Then
'    Response.Write "<BR><BR>&nbsp;&nbsp;This site requires a secure connection.  Click on the link below for the secure connection.<br><br>&nbsp;&nbsp;"
'    Response.Write "<a href=""https://secure.rushmore-group.com/MaderaCW/"">https://secure.rushmore-group.com/MaderaCW/</a>"
'    Response.End
'End If

%><!--#include file="IncCnn.asp"-->
  <!--#include file="IncDefStyles.asp"-->

<HTML>
<HEAD>
    <META name=vs_targetSchema content="HTML 4.0">
    <META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
    <TITLE><%=Trim(gstrLocationName & " " & gstrOrgAbbr & " " & gstrAppName)%></TITLE>
</HEAD>

<SCRIPT ID=ClientScript LANGUAGE=vbscript>
<!--
Option Explicit

Sub window_onload
    'Display the application logon form:
    Call ShowLogon
End Sub

Sub ShowLogon()
    Window.ResizeTo 400, 335
    Window.MoveTo 25, 25
    Call Window.open("Logon.asp", Null, "directories=no,fullscreen=no,location=no,menubar=no,status=no,resizable=yes,toolbar=no,height=300,width=425,scrollbars=yes")
End Sub
-->
</SCRIPT>

<BODY style="COLOR:<%=gstrTitleColor%>; BACKGROUND-COLOR: <%=gstrPageColor%>">
    <%
    Response.Write "<DIV id=lblLocationName CLASS=DefTitleText style=""LEFT:140; TOP:5"">"
    Response.Write gstrLocationName & "<br>"

    If gstrOrgName <> "" Then
        Response.Write gstrOrgName & "<br>"
    End If
    If gstrAppName <> "" Then
        Response.Write gstrAppName & "<br>"
    End If
    Response.Write "<br>"
    Response.Write "<span style=""LEFT:0; FONT-WEIGHT:normal; FONT-SIZE:10pt"">This window is for display only and may be closed at any time after successfully logging in.</span><BR>"
    Response.Write "</DIV>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<img border=0 src=""rg2.jpg"" style=""position:absolute;top:5"">" & vbCrLf
    %>
</BODY>
</HTML>