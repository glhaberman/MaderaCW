<SCRIPT id=DrillDownClientScript language=vbscript>
Dim mstrWeight

Sub ColMouseEvent(intDir, intColID, intRowID)
    If intDir = 0 Then
        ' Mouse over
        mstrWeight = document.all("lblCol" & intColID & intRowID).style.fontweight
        document.all("lblCol" & intColID & intRowID).style.fontweight = "bold"
    Else
        ' Mouse out
        document.all("lblCol" & intColID & intRowID).style.fontweight = mstrWeight
    End If
End Sub

Sub DrillDownColClickEvent(strSPName, intColID, intRowID, blnUseOfficeLevel, intOtherID)
    Dim intI, intJ
    Dim strStaffParms
    Dim aStaffItemIDs(5)
    Dim strSelectedName
    Dim strModalParms
    Dim strSortColumn
    Dim ctlClicked, strCtlValue
    
    Set ctlClicked = document.all("lblCol" & intColID & intRowID)
    <%'Find first staffing level with a name selected and capture name for drilldown heading %>
    strSelectedName = ""
    For intI = 1 To 5
        If Parse(document.all("txtDrillDownNames" & intRowID).value,"^",intI) <> "" Then
            strSelectedName = Parse(document.all("txtDrillDownNames" & intRowID).value,"^",intI)
            Exit For
        End If
    Next
    If strSelectedName = "" Then strSelectedName = "All"
    If blnUseOfficeLevel = True Then
        intJ = 7
    Else
        intJ = 6
    End If
    <%'Build the staffing argument list.  The array below is used to tie the staffing item from 
    ' the hidden text box to the correct parameter in the stored procedure call.
    ' For example, the worker is item 1 in the hidden text box and it is parameter 11 in the stored proc. %>
    aStaffItemIDs(1) = 10
    aStaffItemIDs(2) = 9
    aStaffItemIDs(3) = 8
    aStaffItemIDs(4) = intJ
    aStaffItemIDs(5) = 6
    
    strStaffParms = ""
    For intI = 1 To intJ-2
        strStaffParms = strStaffParms & "&A" & aStaffItemIDs(intI) & "=" & Parse(document.all("txtDrillDownNames" & intRowID).value,"^",intI)
    Next

    strModalParms = "<%=mstrStaticParms%>" & strStaffParms
    strModalParms = strModalParms & "&DD=" & intColID
    strModalParms = strModalParms & "&SN=" & strSelectedName
    strModalParms = strModalParms & "&DD2=" & intOtherID
    strSortColumn = "A0"
    
    strCtlValue = ctlClicked.innerText
    Do While strSortColumn <> ""
        ctlClicked.innerText = strCtlValue & "..."
        strURL = "RptReportDetails.asp?SC=" & strSortColumn & "&Rpt=" & "<%=Request.Form("ReportTitle")%>" & "&Proc=" & strSPName & strModalParms
        strReturnValue = window.showModalDialog(strURL, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
        ctlClicked.innerText = strCtlValue
        strSortColumn = strReturnValue
    Loop
End Sub

Sub DrillDownColClickEventNoStaff(strSPName, intColID, intRowID, intDD2, strSelectedName)
    Dim intI, intJ
    Dim strModalParms
    Dim strSortColumn
    Dim ctlClicked, strCtlValue, strReportTitle
    
    Set ctlClicked = document.all("lblCol" & intColID & intRowID)
    Select Case strSPName
        Case "spRptCausalFactor"
            Select Case strSPName
                Case "spRptCausalFactor"
                    strReportTitle = "Causal Factor Summary"
            End Select
            strModalParms = ""
            <%
                For intI = 1 To 17
                    If intI <> 15 Then
                        Response.Write "strModalParms = strModalParms & ""&A" & intI & "=" & maCriteria(intI) & """" & vbCrLf
                    End If
                Next
            %>
            strModalParms = strModalParms & "&A15=" & Parse(intDD2,"^",2)
            strModalParms = strModalParms & "&SN=" & Parse(intDD2,"^",3)
            intDD2 = Parse(intDD2,"^",1)
        Case Else
            strReportTitle = "<%=Request.Form("ReportTitle")%>"
            strModalParms = ""
            <%
            incIntJ = 4
            For intI = 6 To 10
                Response.Write "strModalParms = strModalParms & ""&A" & intI & "=" & maStaffParms(incIntJ) & """" & vbCrLf
                incIntJ = incIntJ - 1
            Next
            %>
            strModalParms = strModalParms & "<%=mstrStaticParms%>"
    End Select
    strModalParms = strModalParms & "&DD=" & intColID
    strModalParms = strModalParms & "&SN2=" & strSelectedName
    'intDD2 is an optional parameter, used on reports where drilldown requires 2 parameters
    strModalParms = strModalParms & "&DD2=" & intDD2
    strSortColumn = "A0"
    
    strCtlValue = Trim(ctlClicked.innerText)
    Do While strSortColumn <> ""
        ctlClicked.innerText = strCtlValue & "..."
        strURL = "RptReportDetails.asp?SC=" & strSortColumn & "&Rpt=" & strReportTitle & "&Proc=" & strSPName & strModalParms
        strReturnValue = window.showModalDialog(strURL, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
        ctlClicked.innerText = strCtlValue
        strSortColumn = strReturnValue
    Loop
End Sub

</SCRIPT>
