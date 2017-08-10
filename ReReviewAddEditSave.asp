<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: CaseAddEditSave.asp                                             '
'  Purpose: This page is used to save tblReview records                     '
'==========================================================================='
Dim madoRs              'Generic recordset reused for various tasks.
Dim mstrPageTitle       'Title used at the top of the page.
Dim mstrAction          'The action from the post-back (add, update, etc).
Dim madoCmdRvw          'ADO command object for updating And getting review.
Dim madoRsRvw           'ADO recordset for updating And getting a review.
Dim mintI               'Generic loop counter.
Dim mblnChangesSaved    'Set to True if page is being loaded after saving changes.
Dim mlngReReviewID
Dim strRecord
Dim strComments
Dim madoReReview
Dim madoReReviewElms
Dim mstrReReviewElements
%>
<!--#include file="IncCnn.asp"-->
<%
mlngReReviewID = 0

'Instantiate the recordset that is reused for temporary results or queries:
Set madoRs = Server.CreateObject("ADODB.Recordset")

'Set the page title:
mstrPageTitle = Trim(gstrTitle & " " & gstrAppName)
mstrAction = ReqForm("FormAction")

'The form indicates that the user clicked Print will adding Or editing
'a review by appending "Print" to the end of the current action.  The
'record is saved, And when posting back the form will finish printing
'in the window_onload.  The "Print" keyword is removed here to restore
'the indicator to the previous action.
If Instr(mstrAction, "Print") Then
    'Remove the Print text from the form action
    mstrAction = Left(mstrAction, Len(mstrAction) - 5)
End If
mblnChangesSaved = False
Select Case mstrAction
    Case "AddRecord"
        'Insert the new review record:
        Set gadoCmd = GetAdoCmd("spReReviewAdd")
            AddParmIn gadoCmd, "@UserID", adVarChar, 20, ReqForm("UserID")
            AddParmIn gadoCmd, "@rrvReReviewer", adVarChar, 255, ReqForm("rrvEvaluater")
            AddParmIn gadoCmd, "@rrvDateEntered", adDBTimeStamp, 0, ReqIsDate("rrvDateEntered")
            AddParmIn gadoCmd, "@rrvOrgReviewID", adInteger, 0, ReqForm("casID")
            AddParmIn gadoCmd, "@rrvComment", adVarChar, 5000, NULL
            AddParmIn gadoCmd, "@rrvStatusID", adInteger, 0, ReqForm("rrvStatusID")
            AddParmIn gadoCmd, "@rrvSubmitted", adChar, 1, ReqForm("rrvSubmitted")
            AddParmIn gadoCmd, "@EvalPrograms", adVarChar, 255, ReqForm("ProgramsReReviewed")
            AddParmIn gadoCmd, "@rrvTypeID", adInteger, 0, ReqForm("ReReviewTypeID")
            AddParmIn gadoCmd, "@rrvRrvSig", adChar, 1, ReqForm("rrvRrvSig")
            AddParmIn gadoCmd, "@rrvRvwSig", adChar, 1, ReqForm("rrvRvwSig")
            AddParmOut gadoCmd, "@rrvID", adInteger, 0
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mlngReReviewID = gadoCmd.Parameters("@rrvID").Value
        Set gadoCmd = Nothing
        If mlngReReviewID <> -1 Then
            For mintI = 1 To 1000
                strRecord = Parse(ReqForm("ReReviewElementsWrite"),"|",mintI)
                If strRecord = "" Then Exit For
                strComments = Parse(strRecord,"^",6)
                If Len(strComments) = 0 Then strComments = ""
                Set gadoCmd = GetAdoCmd("spRRVElementsAdd")
                    AddParmIn gadoCmd, "@rreEvaluationID", adInteger, 0, mlngReReviewID
                    AddParmIn gadoCmd, "@rreProgramID", adInteger, 0, CLng(Parse(strRecord,"^",1))
                    AddParmIn gadoCmd, "@rreTypeID", adInteger, 0, CLng(Parse(strRecord,"^",2))
                    AddParmIn gadoCmd, "@rreElementID", adInteger, 0, CLng(Parse(strRecord,"^",3))
                    If Parse(strRecord,"^",4) = "" Then
                        AddParmIn gadoCmd, "@rreFactorID", adInteger, 0, 0
                    Else
                        AddParmIn gadoCmd, "@rreFactorID", adInteger, 0, CLng(Parse(strRecord,"^",4))
                    End If
                    AddParmIn gadoCmd, "@rreStatusID", adInteger, 0, CLng(Parse(strRecord,"^",5))
                    AddParmIn gadoCmd, "@rreComments", adVarChar, 5000, strComments
                    'ShowCmdParms(gadoCmd) '***DEBUG
                    gadoCmd.Execute
            Next
        End If
        mblnChangesSaved = True
    Case "ChangeRecord"
        mlngReReviewID = ReqForm("rrvID")
        Set gadoCmd = GetAdoCmd("spReReviewUpd")
            AddParmIn gadoCmd, "@UserID", adVarChar, 20, ReqForm("UserID")
            AddParmIn gadoCmd, "@rrvID", adInteger, 0, mlngReReviewID
            AddParmIn gadoCmd, "@rrvSubmitted", adVarChar, 1, ReqForm("rrvSubmitted")
            AddParmIn gadoCmd, "@rrvDateEntered", adDBTimeStamp, 0, ReqIsDate("rrvDateEntered")
            AddParmIn gadoCmd, "@rrvComment", adVarChar, 5000, NULL
            AddParmIn gadoCmd, "@rrvStatusID", adInteger, 0, ReqForm("rrvStatusID")
            AddParmIn gadoCmd, "@EvalPrograms", adVarChar, 255, ReqForm("ProgramsReReviewed")
            AddParmIn gadoCmd, "@rrvReReviewer", adVarChar, 255, ReqForm("rrvEvaluater")
            AddParmIn gadoCmd, "@rrvRrvSig", adChar, 1, ReqForm("rrvRrvSig")
            AddParmIn gadoCmd, "@rrvRvwSig", adChar, 1, ReqForm("rrvRvwSig")
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
        Set gadoCmd = Nothing
        If ReqForm("ElementsChanged") = "Y" Then
            ' Delete all existing ReReview Elements
            Set gadoCmd = GetAdoCmd("spReReviewDelElements")
                AddParmIn gadoCmd, "@ReReviewID", adInteger, 0, mlngReReviewID
                gadoCmd.Execute
            Set gadoCmd = Nothing

            For mintI = 1 To 1000
                strRecord = Parse(ReqForm("ReReviewElementsWrite"),"|",mintI)
                If strRecord = "" Then Exit For
                strComments = Parse(strRecord,"^",6)
                If Len(strComments) = 0 Then strComments = ""
                Set gadoCmd = GetAdoCmd("spRRVElementsAdd")
                    AddParmIn gadoCmd, "@rreEvaluationID", adInteger, 0, mlngReReviewID
                    AddParmIn gadoCmd, "@rreProgramID", adInteger, 0, CLng(Parse(strRecord,"^",1))
                    AddParmIn gadoCmd, "@rreTypeID", adInteger, 0, CLng(Parse(strRecord,"^",2))
                    AddParmIn gadoCmd, "@rreElementID", adInteger, 0, CLng(Parse(strRecord,"^",3))
                    If Parse(strRecord,"^",4) = "" Then
                        AddParmIn gadoCmd, "@rreFactorID", adInteger, 0, 0
                    Else
                        AddParmIn gadoCmd, "@rreFactorID", adInteger, 0, CLng(Parse(strRecord,"^",4))
                    End If
                    AddParmIn gadoCmd, "@rreStatusID", adInteger, 0, CLng(Parse(strRecord,"^",5))
                    AddParmIn gadoCmd, "@rreComments", adVarChar, 5000, strComments
                    'ShowCmdParms(gadoCmd) '***DEBUG
                    gadoCmd.Execute
            Next
        End If
    Case "DeleteRecord"
        'Delete an existing re-review:
        mlngReReviewID = ReqForm("rrvID")
        Set gadoCmd = GetAdoCmd("spReReviewDel")
            AddParmIn gadoCmd, "@rrvID", adInteger, 0, mlngReReviewID
            gadoCmd.Execute
        Set gadoCmd = Nothing
End Select

If mstrAction <> "DeleteRecord" Then
    ' Rebuild the ReReviewElements string to ensure blank place holders are included
    ' when not all reviewed programs are re-reviewed.

    'Retrieve the Re-Review to display:
    Set madoReReview = Server.CreateObject("ADODB.Recordset")
    Set madoReReviewElms = Server.CreateObject("ADODB.Recordset")
    Set gadoCmd = GetAdoCmd("spReReviewGet")
        AddParmIn gadoCmd, "@ReReviewID", adInteger, 0, mlngReReviewID
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, ReqForm("UserID")
        'Call ShowCmdParms(adCmd) '***DEBUG
        madoReReview.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly

    Set madoReReviewElms = madoReReview.NextRecordset
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
End If

Function ConvertCRLFToBR(strText)
    Dim strTemp
    Dim intI
    
    If IsNull(strText) Then strText = ""
    If strText = "" Then
        strTemp = ""
    Else
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
    End If
    ConvertCRLFToBR = strTemp
End Function

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
    window.parent.SaveWindow.style.left = -1000
    window.parent.divCaseBody.style.left = 1

    ' Update parent Form with save form.    
    window.parent.Form.rrvID.value = <%=mlngReReviewID%>
    window.parent.Form.SaveCompleted.value = "Y"
    window.parent.Form.ElementsChanged.value = ""
    window.parent.Form.ReReviewElements.value = "<%=mstrReReviewElements%>"
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:<%=gstrBackColor%>; overflow:scroll" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
<BR><BR><BR>
<div id=divDisplay>
    Saving Record...
</div>

</BODY>
</HTML>
