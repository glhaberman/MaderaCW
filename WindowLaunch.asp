<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: WindowLaunch.asp                                                '
'  Purpose: This page is used to open child pages.  This page is opened     '
'           with window.open, allowing for multiple pages to be opened at   '
'           one time.                                                       '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<HTML><HEAD>
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <STYLE id=ThisPageStyles type="text/css" rel="stylesheet">
        BODY
            {
            margin:1;
            position: absolute; 
            FONT-SIZE: 10pt; 
            FONT-FAMILY: Tahoma; 
            OVERFLOW: auto; 
            BACKGROUND-COLOR: #FFFFCC
            }
    </STYLE>
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Option Explicit

Sub window_onload
    PageBody.style.cursor="wait"

    ' All pages that open ReportsLaunch must have the following values in the hidden form    
    Form.UserID.value = window.opener.document.Form.UserID.value
    Form.Password.value = window.opener.document.Form.Password.value
    Form.CalledFrom.value = window.opener.document.Form.CalledFrom.value
    Form.ProgramsSelected.value = window.opener.document.Form.ProgramsSelected.value
    Form.casID.value = window.opener.document.Form.casID.value
    Form.rvwID.value = window.opener.document.Form.rvwID.value
    Form.FormAction.value = window.opener.document.Form.FormAction.value
    Form.ReReviewID.value = window.opener.document.Form.ReReviewID.value
    Form.ReReviewTypeID.value = window.opener.document.Form.ReReviewTypeID.value
    Form.WhoCalled.value = window.opener.document.Form.WhoCalled.value
    
    Form.Action = window.opener.document.Form.Action
    Form.Submit
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:#white; overflow:scroll" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
<BR><BR><BR>
<div id=divDisplay>
    Accessing Database...
</div>
<%
Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY: hidden"" ACTION="""" ID=Form>" & vbCrLf
    ' Generic values used on all reports
	WriteFormField "UserID", ""
	WriteFormField "Password", ""
	WriteFormField "CalledFrom", ""
	WriteFormField "ProgramsSelected", ""
	WriteFormField "casID", ""
	WriteFormField "rvwID", ""
	WriteFormField "FormAction", ""
    WriteFormField "ReReviewID", 0
    WriteFormField "ReReviewTypeID", 0
    WriteFormField "WhoCalled",""
Response.Write "</FORM>"
%>

</BODY>
</HTML>
<!--#include file="IncWriteFormField.asp"-->