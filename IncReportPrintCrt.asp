<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncReportPrintCrt.asp                                            '
' Purpose: This include file has the functions to print selected report     '
'	criteria, the form fields, and report visibility						'
'                                                                           '
'==========================================================================='
Sub WriteCriteria()
	Dim strReport
	Dim strReviewClassTitle
	Dim strSelectedCriteria
	Dim intColCnt
	Dim strBasedOn
	Dim strText
	Dim strReviewMonth
	
    strReport = Request.Form("ReportNum")
    
    'Check if Criteria was selected to be printed
    strSelectedCriteria = ""
    If strReport = "77" Then
        strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("RRDirector")) & Trim(Request.Form("RROffice"))
        strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("ReReviewer")) 
    Else
        'strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("MngLvl5")) & Trim(Request.Form("Director")) 
        'strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("Office")) & Trim(Request.Form("ProgramManager"))
        If Trim(Request.Form("Supervisor")) <> "0" Then
            strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("Supervisor"))
        End If
        If Trim(Request.Form("Worker")) <> "0" Then
            strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("Worker"))
        End If
        If Trim(Request.Form("Submitted")) <> "0" Then
            strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("Submitted"))
        End If
        If Trim(Request.Form("Reviewer")) <> "0" Then
            strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("Reviewer"))
        End If
        strSelectedCriteria = strSelectedCriteria & Trim(Request.Form("ReviewTypeText")) '& Trim(Request.Form("CaseActionText"))
    End If
    If strSelectedCriteria = "" Then
        'Nothing was specified, so no need to print a criteria section.
		Exit Sub
    End IF
'response.Write "<br><br>xxxsssxxxxxxxxxx" & strSelectedCriteria & "yyyyyyyyyyy<br><br>"
'response.Flush

    If (Request.Form("StartReviewMonth") = "" And Request.Form("EndReviewMonth") = "") Or strReport = "77" Then
        strReviewMonth = ""
    Else
        strReviewMonth = "From:&nbsp;"
        If Request.Form("StartReviewMonth") <> "" Then
            strReviewMonth = strReviewMonth & Left(Request.Form("StartReviewMonth"),2) & "/" & Right(Request.Form("StartReviewMonth"),4) 
        Else
            strReviewMonth = strReviewMonth & "&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        strReviewMonth = strReviewMonth & "&nbsp;&nbsp;&nbsp;To: "
        If Request.Form("EndReviewMonth") <> "" Then
            strReviewMonth = strReviewMonth & Left(Request.Form("EndReviewMonth"),2) & "/" & Right(Request.Form("EndReviewMonth"),4)
        End If
    End If

    strReviewClassTitle = GetAppSetting("Review Class")
    
    '----- Selected Criteria section ------------------------------------------
    Response.Write "<TABLE id=tabCriteria rules=none cellspacing=0 border=0 bordercolor=#C0C0C0 bordercolorlight=#C0C0C0 bordercolordark=#C0C0C0 style=""TABLE-LAYOUT:auto"">"
    ' Column Headers:
    Response.Write "<THEAD>"
    Response.Write "<TR valign=top>"
    Response.Write "<TH width=110></TH>"
    Response.Write "<TH width=215></TH>"
    Response.Write "<TH width=110></TH>"
    Response.Write "<TH width=215></TH>"
    Response.Write "</TR>"
    Response.Write "</THEAD>"
    ' Title Cell (seperated in a tbody):
    Response.Write "<TBODY>"
    Response.Write "<TR>"
    Response.Write "<TD class=RptCriteriaCell colspan=4 style=""text-align:center; padding:5; FONT-SIZE:11pt"">"
	Response.Write "<b>Selected Criteria</b>"
    Response.Write "</TD>"
    Response.Write "</TR>"
    Response.Write "</TBODY>"
    ' Main body of criteria table:
    Response.Write "<TBODY>"
    'Print Upper Management Levels -- Common to all Reports 
    intColCnt = 0
    If strReport = "77" Then
        If Request.Form("RRDirector") <> "" Then
            intColCnt = WriteTableCell(intColCnt, gstrDirTitle, Parse(Request.Form("RRDirector"), "--", 1))
        End If
        If Request.Form("RROffice") <> "" Then
            intColCnt = WriteTableCell(intColCnt, gstrOffTitle, Parse(Request.Form("RROffice"), "--", 1))
        End If
	    If Request.Form("ReReviewer") <> "" Then
            intColCnt = WriteTableCell(intColCnt, gstrEvaTitle, Parse(Request.Form("ReReviewer"), "--", 1))
        End If
    Else
        If Request.Form("Director") <> "" Then
            intColCnt = WriteTableCell(intColCnt, gstrDirTitle, Parse(Request.Form("Director"), "--", 1))
        End If
        If Request.Form("Office") <> "" Then
            intColCnt = WriteTableCell(intColCnt, gstrOffTitle, Parse(Request.Form("Office"), "--", 1))
        End If
	    If Request.Form("ProgramManager") <> "" Then
            intColCnt = WriteTableCell(intColCnt, gstrMgrTitle, Parse(Request.Form("ProgramManager"), "--", 1))
        End If
	End If
    Select Case strReport
        Case 26,57 'Case Accuracy Summary and Detail
            intColCnt = PrintWorkerManagement(intColCnt)
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewTypeText"), "Review Type", intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewClassText"), gstrReviewClassTitle, intColCnt)
        Case 35 'Reviewer Case Count
            intColCnt = PrintReviewerManagement(intColCnt)
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
        Case 29,30,74,75 'Element Overview
            intColCnt = PrintWorkerManagement(intColCnt)
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
			intColCnt = PrintOptional("Element", Replace(Request.Form("EligElementText"),"[AMP]","&"), intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewTypeText"), "Review Type", intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewClassText"), gstrReviewClassTitle, intColCnt)
        Case 28 'Causal Factor Summary
            intColCnt = PrintWorkerManagement(intColCnt)
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
			intColCnt = PrintOptional("Element", Replace(Request.Form("EligElementText"),"[AMP]","&"), intColCnt)
			intColCnt = PrintOptional("Causal Factor", Request.Form("FieldName"), intColCnt)
			If Request.Form("ShowDetail") = "Y" Then
    			intColCnt = PrintOptional("Include All Factors", "Yes", intColCnt)
            Else
    			intColCnt = PrintOptional("Include All Factors", "No", intColCnt)
            End If
            intColCnt = PrintReview(Request.Form("ReviewTypeText"), "Review Type", intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewClassText"), gstrReviewClassTitle, intColCnt)
        Case 55 'Re-Review Element Accuracy Summary
            intColCnt = PrintWorkerManagement(intColCnt)
            intColCnt = PrintReviewerManagement(intColCnt)
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
			intColCnt = PrintOptional("Element", Replace(Request.Form("EligElementText"),"[AMP]","&"), intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewTypeText"), "Review Type", intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewClassText"), gstrReviewClassTitle, intColCnt)
        Case 56 'Re-Review Accuracy Summary
            intColCnt = PrintWorkerManagement(intColCnt)
            intColCnt = PrintReviewerManagement(intColCnt)
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
        Case 127 'Employee Performance
            intColCnt = PrintWorkerManagement(intColCnt)
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
			
        Case 27,138 'Case Review Detail
            intColCnt = PrintWorkerManagement(intColCnt)
            intColCnt = PrintOptional("Case Action", Request.Form("CaseActionText"), intColCnt)
            intColCnt = PrintOptional("Case Number", Request.Form("CaseNumber"), intColCnt)
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewTypeText"), "Review Type", intColCnt)
            intColCnt = PrintReview(Request.Form("ReviewClassText"), gstrReviewClassTitle, intColCnt)
			
        Case 33 'Unsubmitted Reviews
            intColCnt = PrintWorkerManagement(intColCnt)
            If Request.Form("Submitted") = "1" Then
                intColCnt = PrintOptional("Submitted", "No Supervisor Signature", intColCnt)
            ElseIf Request.Form("Submitted") = "2" Then
                intColCnt = PrintOptional("Submitted", "No Worker Acknowledgement", intColCnt)
            ElseIf Request.Form("Submitted") = "3" Then
                intColCnt = PrintOptional("Submitted", "Not Submitted To Reports", intColCnt)
            Else
                intColCnt = PrintOptional("Submitted", "All", intColCnt)
            End If
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
        Case 34 'Response Due
			intColCnt = PrintWorkerManagement(intColCnt)
			intColCnt = PrintOptional("Response", Request.Form("ResponseText"), intColCnt)
			If Request.Form("Resposne") = 1 Then
                intColCnt = PrintOptional("Min Days Past Due", Request.Form("DaysPastDue"), intColCnt)
            Else
                intColCnt = PrintOptional("Min Days Pending", Request.Form("DaysPastDue"), intColCnt)
            End If
			intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
			intColCnt = PrintReview(Request.Form("ReviewClassText"), gstrReviewClassTitle, intColCnt)
			
    End Select
   
    If intColCnt <> 0 Then
        'Fill out the remaining columns to complete the table, and end the row:
        Response.Write "<TD class=RptCriteriaCell>&nbsp</TD>"
        Response.Write "<TD class=RptCriteriaCell>&nbsp</TD>"
        Response.Write "</TR>"
    End If
    Response.Write "<TR><TD colspan=4 style=""border-top:1 solid #c0c0c0"">&nbsp</TR>"
    Response.Write "</TBODY>"
    Response.Write "</TABLE>"
End Sub

Function WriteTableCell(intColCnt, strLabel, strValue)
    If intColCnt = 0 Then
        Response.Write "<TR valign=top>"
    End If
    Response.Write "<TD class=RptCriteriaCell>"
    Response.Write "<b>" & strLabel & ":</b>"
    Response.Write "</TD>"
    intColCnt = intColCnt + 1
    Response.Write "<TD class=RptCriteriaCell>"
    Response.Write strValue
    Response.Write "</TD>"
    intColCnt = intColCnt + 1
    If intColCnt = 4 Then
        Response.Write "</TR>"
        intColCnt = 0
    End If
    
    WriteTableCell = intColCnt
End Function

Function PrintWorkerManagement(intColCnt)
    Dim strSupHeading
    Dim strSupValue
    
    strSupHeading = gstrSupTitle
    strSupValue = Parse(Request.Form("Supervisor"), "--", 1)
    If CInt(ReqForm("ReportNum")) = 34 And ReqForm("RespDueBasedOn") = "R" Then
        strSupHeading = "Reviewer"
        strSupValue = Parse(Request.Form("Reviewer"), "--", 1)
    End If
    
    If strSupValue = "0" Then strSupValue = ""
	If strSupValue <> "" Then
        intColCnt = WriteTableCell(intColCnt, strSupHeading, strSupValue)
    End If
	
	strSupValue = Request.Form("Worker")
    If strSupValue = "0" Then strSupValue = ""
    If strSupValue <> "" Then
        intColCnt = WriteTableCell(intColCnt, gstrWkrTitle, Parse(Request.Form("Worker"), "--", 1))
    End If    
    PrintWorkerManagement = intColCnt
End Function
 
Function PrintReviewerManagement(intColCnt)
	If Request.Form("Reviewer") <> "" Then
        intColCnt = WriteTableCell(intColCnt, gstrRvwTitle, Parse(Request.Form("Reviewer"), "--", 1))
    End If
    PrintReviewerManagement = intColCnt
End Function

Function PrintReview(strText, strLabel, intColCnt)
    Dim strPrint
    Dim strHTML
    Dim intI
    Dim strStyle

    If Trim(strText) = "" Then
        PrintReview = intColCnt
        Exit Function
    End If
    
	strPrint = "x"
	intI = 1
	If strText <> "" Then
	    strHTML = ""
		Do While strPrint <> ""
			strPrint = Parse(strText, "||", intI)
			If strHTML <> "" Then
			    strHTML = strHTML & "&nbsp &nbsp &nbsp &nbsp"
			End If
			strHTML = strHTML & strPrint
			intI = intI + 1
		Loop
    End If
    
    'When PrintReview is called, end the previous tbody and start a new one:
    strStyle = ""
    If intColCnt <> 0 Then
        'Need to close previous row - fill out the remaining 
        'columns to complete the row, and end the row:
        Response.Write "<TD colspan=2 class=RptCriteriaCell>&nbsp</TD>"
        Response.Write "</TR>"
        intColcnt = 0
        Response.Write "</TBODY>"
        Response.Write "<TBODY>"
    End If
    If strLabel = "Review Type" Then
        'strStyle = " style=""border-top:1 solid #C0C0C0"""
    Else
        strStyle = " style=""border-bottom:1 solid #C0C0C0"""
    End If
    Response.Write "<TR>"
    Response.Write "<TD class=RptCriteriaCell" & strStyle & ">"
    Response.Write "<b>" & strLabel & ":</b>"
    Response.Write "</TD>"

    Response.Write "<TD colspan=3 class=RptCriteriaCell" & strStyle & ">"
    Response.Write strHTML
    Response.Write "</TD>"
    Response.Write "</TR>"

    PrintReview = intColCnt
End Function

Function PrintOptional(strLabel, strText, intColCnt)
	If strText <> "" Then
        intColCnt = WriteTableCell(intColCnt, strLabel, strText)
	End If
	PrintOptional = intColCnt
End Function

Sub ReqQS_WriteCriteria()
	Dim strSelectedCriteria
	Dim intColCnt
	Dim strText
	Dim strReviewMonth
    
    'Check if Criteria was selected to be printed
    strText = "N"
    For intColCnt = 7 To 18
        Select Case intColCnt
            Case 14, 15, 18
                If CLng(maCriteria(intColCnt)) > 0 Then
                    strText = "Y"
                    Exit For
                End If
            Case Else
                If maCriteria(intColCnt) <> "" Then
                    strText = "Y"
                    Exit For
                End If
        End Select 
    Next
    
    If strText = "N" Then
        'Nothing was specified, so no need to print a criteria section.
		Exit Sub
    End If

    If maCriteria(16) <> "" Or maCriteria(17) <> "" Then
        strReviewMonth = "From:&nbsp;"
        If maCriteria(16) <> "" Then
            strReviewMonth = strReviewMonth & Left(maCriteria(16),2) & "/" & Right(maCriteria(16),4) 
        Else
            strReviewMonth = strReviewMonth & "&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        strReviewMonth = strReviewMonth & "&nbsp;&nbsp;&nbsp;To: "
        If maCriteria(17) <> "" Then
            strReviewMonth = strReviewMonth & Left(maCriteria(17),2) & "/" & Right(maCriteria(17),4)
        End If
    Else
        strReviewMonth = ""
    End If
    '----- Selected Criteria section ------------------------------------------
    Response.Write "<TABLE id=tabCriteria rules=none cellspacing=0 border=0 bordercolor=#C0C0C0 bordercolorlight=#C0C0C0 bordercolordark=#C0C0C0 style=""TABLE-LAYOUT:auto"">"
    ' Column Headers:
    Response.Write "<THEAD>"
    Response.Write "<TR valign=top>"
    Response.Write "<TH width=110></TH>"
    Response.Write "<TH width=215></TH>"
    Response.Write "<TH width=110></TH>"
    Response.Write "<TH width=215></TH>"
    Response.Write "</TR>"
    Response.Write "</THEAD>"
    ' Title Cell (seperated in a tbody):
    Response.Write "<TBODY>"
    Response.Write "<TR>"
    Response.Write "<TD class=RptCriteriaCell colspan=4 style=""text-align:center; padding:5; FONT-SIZE:11pt"">"
	Response.Write "<b>Selected Criteria</b>"
    Response.Write "</TD>"
    Response.Write "</TR>"
    Response.Write "</TBODY>"
    ' Main body of criteria table:
    Response.Write "<TBODY>"
    'Print Upper Management Levels -- Common to all Reports 
    intColCnt = 0
    If maCriteria(7) <> "" Then
        intColCnt = WriteTableCell(intColCnt, gstrDirTitle, maCriteria(7))
    End If
    If maCriteria(8) <> "" Then
        intColCnt = WriteTableCell(intColCnt, gstrOffTitle, maCriteria(8))
    End If
    If maCriteria(9) <> "" Then
        intColCnt = WriteTableCell(intColCnt, gstrMgrTitle, maCriteria(9))
    End If
    intColCnt = ReqQS_PrintWorkerManagement(intColCnt)
	intColCnt = PrintOptional("Review Month", strReviewMonth, intColCnt)
	intColCnt = PrintOptional("Element", Replace(maCriteriaText(15),"[AMP]","&"), intColCnt)
	intColCnt = PrintOptional("Causal Factor", maCriteriaText(18), intColCnt)
    intColCnt = PrintReview(maCriteriaText(12), "Review Type", intColCnt)
    intColCnt = PrintReview(maCriteriaText(13), "Review Class", intColCnt)
    If intColCnt <> 0 Then
        'Fill out the remaining columns to complete the table, and end the row:
        Response.Write "<TD class=RptCriteriaCell>&nbsp</TD>"
        Response.Write "<TD class=RptCriteriaCell>&nbsp</TD>"
        Response.Write "</TR>"
    End If
    Response.Write "<TR><TD colspan=4 style=""border-top:1 solid #c0c0c0"">&nbsp</TR>"
    Response.Write "</TBODY>"
    Response.Write "</TABLE>"
End Sub

Function ReqQS_PrintWorkerManagement(intColCnt)
	If maCriteria(10) <> "" And maCriteria(10) <> "0" Then
        intColCnt = WriteTableCell(intColCnt, gstrSupTitle, maCriteria(10))
    End If
	
    If maCriteria(11) <> "" And maCriteria(11) <> "0" Then
        intColCnt = WriteTableCell(intColCnt, gstrWkrTitle, maCriteria(11))
    End If    
    ReqQS_PrintWorkerManagement = intColCnt
End Function

%>