<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: FormDefEditInc.asp                                               '
' Purpose: This include file contains the definition of the input elements  '
'          for the HTML form used to post Case data entry.  This version is '
'          for a form filled in from the case recordset.                    '
' Changes:                                                                  '
'                                                                           '
'==========================================================================='
%>

    <INPUT TYPE="hidden" name="rrvorgReviewID" value=<%=madoReReview("rrvOrgReviewID").Value%> id=rrvorgReviewID>
    <INPUT TYPE="hidden" name="rrvDateEntered" value="<%=madoReReview("rrvDateEntered").Value%>" id=rrvDateEntered>
    <INPUT TYPE="hidden" name="rrvSubmitted" value="<%=madoReReview("rrvSubmitted").Value%>" id=rrvSubmitted>
    <INPUT type="hidden" name="rrvStatusID" value="<%=madoReReview("rrvStatusID").Value%>" id=rrvStatusID>
    <INPUT type="hidden" name="rrvResponseID" value="<%=madoReReview("rrvResponseID").Value%>" id=rrvResponseID>
    <INPUT type="hidden" name="rrvResponseDue" value="<%=madoReReview("rrvResponseDue").Value%>" id=rrvResponseDue>
    <INPUT TYPE="hidden" name="rrvEvaluater" value="<%=madoReReview("rrvReReviewer").Value%>" id=rrvEvaluater>
    <INPUT TYPE="hidden" name="rrvRrvSig" value="<%=madoReReview("rrvRrvSig").Value%>" id=rrvRrvSig>
    <INPUT TYPE="hidden" name="rrvRvwSig" value="<%=madoReReview("rrvRvwSig").Value%>" id=rrvRvwSig>
    <INPUT TYPE="hidden" name="ProgramsReviewed" value="<%=madoReReview("ProgramsReviewed").Value%>" id=ProgramsReviewed>
    <INPUT TYPE="hidden" name="ProgramsReviewedValue" value="<%=madoReReview("ProgramsReviewed").Value%>" id=ProgramsReviewedValue>
    <INPUT TYPE="hidden" name="ProgramsReReviewed" value="<%=madoReReview("rrvEvalPrograms").Value%>" id=ProgramsReReviewed>
    <INPUT TYPE="hidden" name="ProgramsReReviewedValue" value="<%=madoReReview("rrvEvalPrograms").Value%>" id=ProgramsReReviewedValue>
    <INPUT TYPE="hidden" name="ReviewMonthValue" value="<%=madoReReview("rvwMonthYear").Value%>" id=ReviewMonthValue>
    <INPUT TYPE="hidden" name="ReviewDateValue" value="<%=madoReReview("rvwDateEntered").Value%>" id=ReviewDateValue>
    <INPUT TYPE="hidden" name="ReviewClassValue" value="<%=madoReReview("ReviewClass").Value%>" id=ReviewClassValue>
    <INPUT TYPE="hidden" name="CaseNameValue" value="<%=madoReReview("ClientName").Value%>" id=CaseNameValue>
    <INPUT TYPE="hidden" name="CaseNumberValue" value="<%=madoReReview("rvwCaseNumber").Value%>" id=CaseNumberValue>
    <INPUT TYPE="hidden" name="ReviewStatusValue" value="<%=madoReReview("ReviewStatus").Value%>" id=ReviewStatusValue>
    <INPUT TYPE="hidden" name="ReviewerNameValue" value="<%=madoReReview("rvwReviewerName").Value%>" id=ReviewerNameValue>
    <INPUT TYPE="hidden" name="WorkerNameValue" value="<%=madoReReview("rvwWorkerName").Value%>" id=WorkerNameValue>
    <INPUT TYPE="hidden" name="WorkerResponseValue" value="<%=madoReReview("WorkerResponse").Value%>" id=WorkerResponseValue>
    <INPUT TYPE="hidden" name="ElementsChanged" value="Y" id=ElementsChanged>
<%
mstrReReviewElements = ""
Do While Not madoReReviewElms.EOF
    mstrReReviewElements = mstrReReviewElements & _
        madoReReviewElms.Fields("Program").value & "^" & _
        madoReReviewElms.Fields("Element").value & "^" & _
        madoReReviewElms.Fields("ItemStatus").value & "^" & _
        madoReReviewElms.Fields("FactorName").value & "^" & _
        madoReReviewElms.Fields("GroupID").value & "^" & _
        madoReReviewElms.Fields("GroupName").value & "^" & _
        ConvertCRLFToBR(madoReReviewElms.Fields("rveComments").value) & "^" & _
        madoReReviewElms.Fields("rveProgramID").value & "^" & _
        madoReReviewElms.Fields("rveElementID").value & "^" & _
        madoReReviewElms.Fields("rreStatusID").value & "^" & _
        ConvertCRLFToBR(madoReReviewElms.Fields("rreComments").value) & "^" & _
        madoReReviewElms.Fields("ReviewType").value & "^" & _
        madoReReviewElms.Fields("rveTypeID").value & "^" & _
        madoReReviewElms.Fields("FactorID").value & "|"
    
    madoReReviewElms.MoveNext
Loop

Function ConvertCRLFToBR(strText)
    Dim strTemp
    Dim intI
    
    strTemp = ""
    For intI = 1 To Len(strText)
        If Asc(Mid(strText, intI, 1)) = 13 Then
            strTemp = strTemp & "[linebreak]"
        Else
            If Asc(Mid(strText, intI, 1)) <> 10 Then
                strTemp = strTemp & Mid(strText, intI, 1)
            End If
        End If
    Next
    ConvertCRLFToBR = strTemp
End Function

%>