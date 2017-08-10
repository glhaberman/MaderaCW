<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncFormsReportDef.asp                                              '
' Purpose: This include file contains the definition of the input elements  '
'          for the HTML form used to post Case data entry.  This version is '
'          for a blank form.                                                '
'                                                                           '
'==========================================================================='
Sub ReportFormDef()
	'ReportIndex is currently being used as a temporary solution for keeping
	'track of which specific report was selected, now that some reports have
	'the same criteria screen:
	WriteFormField "ReportIndex", ReqForm("ReportIndex")
	WriteFormField "ReportTitle", ReqForm("ReportTitle")
	
	WriteFormField "StartDate", ReqForm("StartDate")
	WriteFormField "EndDate", ReqForm("EndDate")
	WriteFormField "ReportingMode", ReqForm("ReportingMode")
	WriteFormField "ReportNum", ReqForm("ReportNum")
	
	WriteFormField "SupBackGround", ReqForm("SupBackGround")
	WriteFormField "SupFontColor", ReqForm("SupFontColor")
	WriteFormField "WkrBackGround", ReqForm("WkrBackGround")
	WriteFormField "WkrFontColor", ReqForm("WkrFontColor")
	WriteFormField "ColBackGround", ReqForm("ColBackGround")
	WriteFormField "ColFontColor", ReqForm("ColFontColor")
	
	WriteFormField "DirectorID", ReqForm("DirectorID")
	WriteFormField "Director", ReqForm("Director")
	WriteFormField "OfficeID", ReqForm("OfficeID")
	WriteFormField "Office", ReqForm("Office")
	WriteFormField "ProgramManagerID", ReqForm("ProgramManagerID")
	WriteFormField "ProgramManager", ReqForm("ProgramManager")
	WriteFormField "SupervisorID", ReqForm("SupervisorID")
	WriteFormField "Supervisor", ReqForm("Supervisor")
	WriteFormField "WorkerID", ReqForm("WorkerID")
	WriteFormField "Worker", ReqForm("Worker")
	WriteFormField "ReviewerID", ReqForm("ReviewerID")
	WriteFormField "Reviewer", ReqForm("Reviewer")
	
	WriteFormField "ReviewTypeID", ReqForm("ReviewTypeID")
	WriteformField "ReviewTypeText", ReqForm("ReviewTypeText")
	WriteFormField "ReviewClassID", ReqForm("ReviewClassID")
	WriteformField "ReviewClassText", ReqForm("ReviewClassText")
	
	WriteFormField "ProgramID", ReqForm("ProgramID")
	WriteFormField "ProgramText", ReqForm("ProgramText")
	WriteFormField "EligElementID", ReqForm("EligElementID")
	WriteFormField "EligElementText", ReqForm("EligElementText")
	
	WriteFormField "CaseActionID", ReqForm("CaseActionID")
	WriteFormField "CaseActionText", ReqForm("CaseActionText")
	WriteFormField "DiscoveryID", ReqForm("DiscoveryID")
	WriteFormField "DiscoveryText", ReqForm("DiscoveryText")
	WriteFormField "BenErrorTypeID", ReqForm("BenErrorTypeID")
	WriteFormField "BenErrorTypeText", ReqForm("BenErrorTypeText")
	
	WriteFormField "ResponseID", ReqForm("ResponseID")
	WriteFormField "ResponseText", ReqForm("ResponseText")
	WriteFormField "HouseholdParents", ReqForm("HouseholdParents")
	WriteFormField "HouseholdText", ReqForm("HouseholdText")
	WriteFormField "PartHours", ReqForm("PartHours")
	WriteFormField "PartHoursText", ReqForm("PartHoursText")
    WriteFormField "CaseNumber", ReqForm("CaseNumber")
	WriteFormField "MinAvgDays", ReqForm("MinAvgDays")
	WriteFormField "ShowNonComOnlyID", ReqForm("ShowNonComOnlyID")
	WriteFormField "ShowNonComOnly", ReqForm("ShowNonComOnly")
	WriteFormField "Submitted", ReqForm("Submitted")
	WriteFormField "ShowDetail", ReqForm("ShowDetail")
	WriteFormField "CountyOfficesID", ReqForm("CountyOfficesID")
	WriteFormField "CountyOfficesText", ReqForm("CountyOfficesText")
End Sub
%>