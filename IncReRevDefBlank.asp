<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: FormDefBlankInc.asp                                              '
' Purpose: This include file contains the definition of the INPUT elements  '
'          for the HTML form used to post Case data entry.  This version is '
'          for a blank form.                                                '
'                                                                           '
'==========================================================================='
%>

    <INPUT TYPE="hidden" name="rrvorgReviewID" value=0 id=rrvorgReviewID>
    <INPUT TYPE="hidden" name="rrvDateEntered" value="" id=rrvDateEntered>
    <INPUT TYPE="hidden" name="rrvSubmitted" value="" id=rrvSubmitted>
    <INPUT type="hidden" name="rrvStatusID" value="" id=rrvStatusID>
    <INPUT type="hidden" name="rrvResponseID" value="0" id=rrvResponseID>
    <INPUT type="hidden" name="rrvResponseDue" value="" id=rrvResponseDue>
    <INPUT TYPE="hidden" name="rrvEvaluater" value="<%=gstrUserName%>" id=rrvEvaluater>
    <INPUT TYPE="hidden" name="ProgramsReviewed" value="" id=ProgramsReviewed>
    <INPUT TYPE="hidden" name="ProgramsReviewedValue" value="" id=ProgramsReviewedValue>
    <INPUT TYPE="hidden" name="ProgramsReReviewed" value="" id=ProgramsReReviewed>
    <INPUT TYPE="hidden" name="ProgramsReReviewedValue" value="" id=ProgramsReReviewedValue>
    <INPUT TYPE="hidden" name="ReviewMonthValue" value="" id=ReviewMonthValue>
    <INPUT TYPE="hidden" name="ReviewDateValue" value="" id=ReviewDateValue>
    <INPUT TYPE="hidden" name="ReviewClassValue" value="" id=ReviewClassValue>
    <INPUT TYPE="hidden" name="CaseNameValue" value="" id=CaseNameValue>
    <INPUT TYPE="hidden" name="CaseNumberValue" value="" id=CaseNumberValue>
    <INPUT TYPE="hidden" name="ReviewStatusValue" value="" id=ReviewStatusValue>
    <INPUT TYPE="hidden" name="ReviewerNameValue" value="" id=ReviewerNameValue>
    <INPUT TYPE="hidden" name="WorkerNameValue" value="" id=WorkerNameValue>
    <INPUT TYPE="hidden" name="WorkerResponseValue" value="" id=WorkerResponseValue>
    <INPUT TYPE="hidden" name="rrvRrvSig" value="" id=rrvRrvSig>
    <INPUT TYPE="hidden" name="rrvRvwSig" value="" id=rrvRvwSig>
    <INPUT TYPE="hidden" name="ElementsChanged" value="Y" id=ElementsChanged>
