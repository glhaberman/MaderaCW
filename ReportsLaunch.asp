<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ReportsLaunch.asp                                               '
'  Purpose: This page is used to call a report ASP.  This page is opened    '
'           with window.open, allowing for multiple reports to be opened at '
'           one time.  Also, Reports.asp does not have to be reloaded after '
'           closing a report.                                               '
' Includes:                                                                 '
'   IncCnn.asp          - ADO database connection                           '
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
    ' In addition to the above fields, ReportName and ReportType must also be included on the calling page

    ' If calling page passed in a name for the report, display it.
    If window.opener.document.Form.ReportName.value <> "" Then    
        divDisplay.innerHTML = "Building <B>" & window.opener.document.Form.ReportName.value & "</B> report, please wait..."
    End If

    'Report.asp
    If window.opener.document.Form.ReportType.value = "Reports.asp" Then
        Form.ProgramsSelected.value = window.opener.document.Form.ProgramsSelected.value
	    Form.ReportTitle.value = window.opener.document.Form.ReportTitle.value
        Form.ReportIndex.value = window.opener.document.Form.ReportIndex.value
        Form.StartDate.value = window.opener.document.Form.StartDate.value
        Form.EndDate.value = window.opener.document.Form.EndDate.value
        Form.ReportingMode.value = window.opener.document.Form.ReportingMode.value
        Form.ReportNum.value = window.opener.document.Form.ReportNum.value
        Form.DirectorID.value = window.opener.document.Form.DirectorID.value
        Form.Director.value = window.opener.document.Form.Director.value
        Form.OfficeID.value = window.opener.document.Form.OfficeID.value
        Form.Office.value = window.opener.document.Form.Office.value
        Form.ProgramManagerID.value = window.opener.document.Form.ProgramManagerID.value
        Form.ProgramManager.value = window.opener.document.Form.ProgramManager.value
        Form.SupervisorID.value = window.opener.document.Form.SupervisorID.value
        Form.Supervisor.value = window.opener.document.Form.Supervisor.value
        Form.WorkerID.value = window.opener.document.Form.WorkerID.value
        Form.Worker.value = window.opener.document.Form.Worker.value
        Form.ReviewerID.value = window.opener.document.Form.ReviewerID.value
        Form.Reviewer.value = window.opener.document.Form.Reviewer.value
        Form.ReReviewerID.value = window.opener.document.Form.ReReviewerID.value
        Form.ReReviewer.value = window.opener.document.Form.ReReviewer.value
        Form.ReviewTypeID.value = window.opener.document.Form.ReviewTypeID.value
        Form.ReviewTypeText.value = window.opener.document.Form.ReviewTypeText.value
        Form.ReviewClassID.value = window.opener.document.Form.ReviewClassID.value
        Form.ReviewClassText.value = window.opener.document.Form.ReviewClassText.value
        Form.ProgramID.value = window.opener.document.Form.ProgramID.value
        Form.ProgramText.value = window.opener.document.Form.ProgramText.value
        Form.EligElementID.value = window.opener.document.Form.EligElementID.value
        Form.EligElementText.value = window.opener.document.Form.EligElementText.value
        Form.CaseActionID.value = window.opener.document.Form.CaseActionID.value
        Form.CaseActionText.value = window.opener.document.Form.CaseActionText.value
        Form.ResponseID.value = window.opener.document.Form.ResponseID.value
        Form.ResponseText.value = window.opener.document.Form.ResponseText.value
        Form.CaseNumber.value = window.opener.document.Form.CaseNumber.value
        Form.MinAvgDays.value = window.opener.document.Form.MinAvgDays.value
        Form.ShowNonComOnlyID.value = window.opener.document.Form.ShowNonComOnlyID.value
        Form.ShowNonComOnly.value = window.opener.document.Form.ShowNonComOnly.value
        Form.Submitted.value = window.opener.document.Form.Submitted.value
        Form.ShowDetail.value = window.opener.document.Form.ShowDetail.value
        Form.IncludeCorrect.value = window.opener.document.Form.IncludeCorrect.value
        Form.DaysPastDue.value = window.opener.document.Form.DaysPastDue.value
	    Form.StartReviewMonth.value = window.opener.document.Form.StartReviewMonth.value
	    Form.EndReviewMonth.value = window.opener.document.Form.EndReviewMonth.value
	    Form.RespDueBasedOn.value = window.opener.document.Form.RespDueBasedOn.value
        Form.SSProgramID.value = window.opener.document.Form.SSProgramID.value
        Form.RepLAliasPosID.value = window.opener.document.Form.RepLAliasPosID.value
        Form.RepLUserAdmin.value = window.opener.document.Form.RepLUserAdmin.value
        Form.RepLUserQA.value = window.opener.document.Form.RepLUserQA.value
        Form.RepLUserID.value = window.opener.document.Form.RepLUserID.value
        Form.FactorID.value = window.opener.document.Form.FactorID.value
        Form.FactorText.value = window.opener.document.Form.FactorText.value
        Form.ReReviewTypeID.value = window.opener.document.Form.ReReviewTypeID.value
    End If
    'RptPositionChartPrint.asp
    If window.opener.document.Form.ReportType.value = "RptPositionChartPrint.asp" Then
        Form.RecordID.value = window.opener.document.Form.RecordID.value
    End If

    Form.Action = window.opener.document.Form.Action
    Form.Submit
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:#white; overflow:scroll" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
<BR><BR><BR>
<div id=divDisplay>
    Building report, please wait...
</div>
<%
Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY: hidden"" ACTION="""" ID=Form>" & vbCrLf
	WriteFormField "UserID", ""
	WriteFormField "Password", ""
	WriteFormField "CalledFrom", ""

    ' Values used on reports called from Reports.asp        
	WriteFormField "ProgramsSelected", ""
        
	WriteFormField "ReportIndex", ""
	WriteFormField "ReportTitle", ""
	WriteFormField "StartDate", ""
	WriteFormField "EndDate", ""
	WriteFormField "ReportingMode", ""
	WriteFormField "ReportNum", ""
	
	WriteFormField "DirectorID", ""
	WriteFormField "Director", ""
	WriteFormField "OfficeID", ""
	WriteFormField "Office", ""
	WriteFormField "ProgramManagerID", ""
	WriteFormField "ProgramManager", ""
	WriteFormField "SupervisorID", ""
	WriteFormField "Supervisor", ""
	WriteFormField "WorkerID", ""
	WriteFormField "Worker", ""
	WriteFormField "ReviewerID", ""
	WriteFormField "Reviewer", ""
	WriteFormField "ReReviewerID", ""
	WriteFormField "ReReviewer", ""
	
	WriteFormField "ReviewTypeID", ""
	WriteformField "ReviewTypeText", ""
	WriteFormField "ReviewClassID", ""
	WriteformField "ReviewClassText", ""
	
	WriteFormField "ProgramID", ""
	WriteFormField "ProgramText", ""
	WriteFormField "EligElementID", ""
	WriteFormField "EligElementText", ""
	
	WriteFormField "CaseActionID", ""
	WriteFormField "CaseActionText", ""
	WriteFormField "ResponseID", ""
	WriteFormField "ResponseText", ""
    WriteFormField "CaseNumber", ""
	WriteFormField "MinAvgDays", ""
	WriteFormField "ShowNonComOnlyID", ""
	WriteFormField "ShowNonComOnly", ""
	WriteFormField "Submitted", ""
	WriteFormField "ShowDetail", ""
	WriteFormField "IncludeCorrect", ""
	WriteFormField "StartReviewMonth", ""
	WriteFormField "EndReviewMonth", ""
	WriteFormField "FactorID", ""
	WriteFormField "FactorText", ""

	' Values used on RptPositionChartPrint.asp
	WriteFormField "RecordID", ""
	WriteFormField "DaysPastDue", ""
	WriteFormField "RespDueBasedOn", ""
	WriteFormField "SSProgramID", ""
	WriteFormField "RepLAliasPosID", ReqForm("RepLAliasPosID")
    WriteFormField "RepLUserAdmin", ReqForm("RepLUserAdmin")
    WriteFormField "RepLUserQA", ReqForm("RepLUserQA")
    WriteFormField "RepLUserID", ReqForm("RepLUserID")
    WriteFormField "ReReviewTypeID", ReqForm("ReReviewTypeID")

Response.Write "</FORM>"
%>

</BODY>
</HTML>
<!--#include file="IncWriteFormField.asp"-->