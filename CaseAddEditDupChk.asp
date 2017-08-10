<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: CaseAddEditDupChk.asp                                           '
'  Purpose: This page is used to check for a duplicate review.  A flag will '
'           be passed back to CaseAddEdit indicating if a dup was found.    '
' Includes:                                                                 '
'   IncCnn.asp          - ADO database connection                           '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim adCmd
Dim adRs
Dim lngID
Dim strMonthYear
Dim lngReviewClassID
Dim strCaseNumber
Dim strDupFlag

lngID = Request.QueryString("ID")
If Len(lngID) = 0 Then lngID = 0
strMonthYear = Request.QueryString("MonthYear")
lngReviewClassID = Request.QueryString("ReviewClassID")
strCaseNumber = Request.QueryString("CaseNumber")

Set adCmd = GetAdoCmd("spReviewDupChk")

adCmd.CommandTimeout = 180
    AddParmIn adCmd, "@ID", adInteger, 0, lngID
    AddParmIn adCmd, "@MonthYear", adVarChar, 7, strMonthYear
    AddParmIn adCmd, "@ReviewClassID", adInteger, 0, lngReviewClassID
    AddParmIn adCmd, "@CaseNumber", adVarChar, 25, strCaseNumber
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
If Not adRs.EOF Then
	strDupFlag = adRs.Fields("DupFlag").Value
End If

%>
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
    window.returnvalue = "<% = strDupFlag %>"
    window.close
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:#white; overflow:scroll" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
<BR><BR><BR>
<div id=divDisplay>
    Checking for duplicate Review, please wait...
</div>

</BODY>
</HTML>
