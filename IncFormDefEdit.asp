<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncFormDefEdit.asp                                               '
' Purpose: This include file contains the definition of the input elements  '
'          for the HTML form used to post Case data entry.  This version is '
'          for a form filled in from the case recordset.                    '
' Changes:                                                                  '
'  06/01/03 tjs - code for Programs that apply check boxes.                 '
'  07/04/03 bw - rewritten to to use response.write for efficiency.         '
'                                                                           '
'==========================================================================='
Sub FormDefEdit()
    Dim strFieldName
    Dim adRsElm, adRsComments, adRsPrgs
    Dim strRecord
	
    WriteFormField "rvwID", madoRsRvw.Fields("rvwID").Value
    WriteFormField "rvwDateEntered", madoRsRvw.Fields("rvwDateEntered").Value
    WriteFormField "rvwMonthYear", madoRsRvw.Fields("rvwMonthYear").Value
    WriteFormField "rvwReviewerName", madoRsRvw.Fields("rvwReviewerName").Value
    WriteFormField "rvwReviewClassID", madoRsRvw.Fields("rvwReviewClassID").Value
	WriteFormField "rvwWorkerID", madoRsRvw.Fields("rvwWorkerID").Value
    WriteFormField "rvwWorkerEmpID", madoRsRvw.Fields("rvwWorkerEmpID").Value
    WriteFormField "rvwWorkerName", madoRsRvw.Fields("rvwWorkerName").Value
    WriteFormField "rvwSupervisorEmpID", madoRsRvw.Fields("rvwSupervisorEmpID").Value
    WriteFormField "rvwSupervisorName", madoRsRvw.Fields("rvwSupervisorName").Value
    WriteFormField "rvwManagerName", madoRsRvw.Fields("rvwManagerName").Value
    WriteFormField "rvwOfficeName", madoRsRvw.Fields("rvwOfficeName").Value
    WriteFormField "rvwDirectorName", madoRsRvw.Fields("rvwDirectorName").Value
    WriteFormField "rvwCaseLastName", madoRsRvw.Fields("rvwCaseLastName").Value
    WriteFormField "rvwCaseFirstName", madoRsRvw.Fields("rvwCaseFirstName").Value
    WriteFormField "rvwCaseNumber", madoRsRvw.Fields("rvwCaseNumber").Value
    WriteFormField "rvwResponseDueDate", madoRsRvw.Fields("rvwResponseDueDate").Value
    WriteFormField "rvwWorkerResponseID", madoRsRvw.Fields("rvwWorkerResponseID").Value
    WriteFormField "rvwWorkerSigResponseID", madoRsRvw.Fields("rvwWorkerSigResponseID").Value
    WriteFormField "rvwSubmitted", madoRsRvw.Fields("rvwSubmitted").Value
    WriteFormField "rvwStatusID", madoRsRvw.Fields("rvwStatusID").Value
    WriteFormField "rvwSupSig", madoRsRvw.Fields("rvwSupSig").Value
    WriteFormField "rvwSupComments", madoRsRvw.Fields("rvwSupComments").Value
    WriteFormField "rvwWrkSig", madoRsRvw.Fields("rvwWrkSig").Value
    WriteFormField "rvwWrkComments", madoRsRvw.Fields("rvwWrkComments").Value
    WriteFormField "rvwUserID", madoRsRvw.Fields("rvwUserID").Value
    WriteFormField "FormAction", ""
    WriteFormField "Changed", ""
    WriteFormField "DeleteCode", ""
    
    'Build the program-element delimited string to pass to the form:
    Set adRsElm = Server.CreateObject("ADODB.Recordset")
    Set gadoCmd = GetAdoCmd("spReviewElementsGet")
        AddParmIn gadoCmd, "@rveReviewID", adInteger, 0, madoRsRvw.Fields("rvwID").Value
        adRsElm.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
    Set gadoCmd = Nothing
    Set adRsComments = adRsElm.NextRecordset
    Set adRsPrgs = adRsElm.NextRecordset
    adRsElm.Sort = "ProgramID, TypeID, ElementID"
    
    strRecord = ""
    Do While Not adRsElm.EOF
        strRecord = strRecord & adRsElm.Fields("ProgramID").Value & "^" & _
                                              adRsElm.Fields("TypeID").Value & "^" & _
                                              adRsElm.Fields("ElementID").Value & "^" & _
                                              adRsElm.Fields("StatusID").Value & "*" & _
                                              adRsElm.Fields("TimeframeID").Value & "*" & _
                                              adRsElm.Fields("Comments").Value & "*" & _
                                              adRsElm.Fields("FactorList").Value & "|"
        adRsElm.MoveNext
    Loop
    adRsElm.Close
    Set adRsElm = Nothing
    WriteFormField "ReviewElementData", strRecord
    strRecord = ""
    Do While Not adRsComments.EOF
        strRecord = strRecord & adRsComments.Fields("rvcScreenName").Value & "^" & _
                                              adRsComments.Fields("rvcComments").Value & "|"
        adRsComments.MoveNext
    Loop
    adRsComments.Close
    Set adRsComments = Nothing
    WriteFormField "ReviewCommentData", strRecord
    strRecord = ""
    Do While Not adRsPrgs.EOF
        strRecord = strRecord & adRsPrgs.Fields("rvpProgramID").Value & "^" & _
                                              adRsPrgs.Fields("rvpReviewTypeID").Value & "|"
        adRsPrgs.MoveNext
    Loop
    adRsPrgs.Close
    Set adRsPrgs = Nothing
    WriteFormField "ReviewProgramData", strRecord 
End Sub
%>
