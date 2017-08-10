<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncFormDefBlank.asp                                              '
' Purpose: This include file contains the definition of the input elements  '
'          for the HTML form used to post Case data entry.  This version is '
'          for a blank form.                                                '
'                                                                           '
'==========================================================================='
Sub FormDefBlank()
    WriteFormField "rvwID", ""
    WriteFormField "rvwDateEntered", ""
    WriteFormField "rvwMonthYear", ""
    WriteFormField "rvwReviewerName", ""
    WriteFormField "rvwReviewerEmpID", ""
    WriteFormField "rvwReviewClassID", "0"
    WriteFormField "rvwWorkerID", ""
    WriteFormField "rvwWorkerName", ""
    WriteFormField "rvwWorkerEmpID", ""
    WriteFormField "rvwSupervisorName", ""
    WriteFormField "rvwSupervisorEmpID", ""
    WriteFormField "rvwManagerName", ""
    WriteFormField "rvwOfficeName", ""
    WriteFormField "rvwDirectorName", ""
    WriteFormField "rvwCaseLastName", ""
    WriteFormField "rvwCaseFirstName", ""
    WriteFormField "rvwCaseNumber", ""
    WriteFormField "rvwResponseDueDate", ""
    WriteFormField "rvwWorkerResponseID", "0"
    WriteFormField "rvwWorkerSigResponseID", "0"
    WriteFormField "rvwSubmitted", "N"
    WriteFormField "rvwSupSig", ""
    WriteFormField "rvwSupComments", ""
    WriteFormField "rvwWrkSig", ""
    WriteFormField "rvwWrkComments", ""
    WriteFormField "rvwStatusID", ""
    WriteFormField "ReviewElementData", ""
    WriteFormField "ReviewCommentData", ""
    WriteFormField "ReviewProgramData", "" 
    WriteFormField "FormAction", "AddRecord"
    WriteFormField "Changed", ""
    WriteFormField "DeleteCode", ""
    WriteFormField "rvwUserID", gstrUserID
End Sub
%>