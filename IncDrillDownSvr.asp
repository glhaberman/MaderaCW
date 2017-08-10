<%
Dim mstrStaticParms
Dim maStaffParms(4)
Dim maCriteria(20), maCriteriaText(20)
Dim incIntJ
mstrStaticParms = "&AAL=" & glngAliasPosID
mstrStaticParms = mstrStaticParms & "&AUA=" & gblnUserAdmin
mstrStaticParms = mstrStaticParms & "&AQA=" & gblnUserQA
mstrStaticParms = mstrStaticParms & "&AUI=" & gstrUserID
mstrStaticParms = mstrStaticParms & "&ASD=" & ReqIsDate("StartDate")
mstrStaticParms = mstrStaticParms & "&AED=" & ReqIsDate("EndDate")
mstrStaticParms = mstrStaticParms & "&ART=" & ReqIsBlank("ReviewTypeID")
mstrStaticParms = mstrStaticParms & "&ARC=" & ReqIsBlank("ReviewClassID")
mstrStaticParms = mstrStaticParms & "&APR=" & ReqZeroToNull("ProgramID")
mstrStaticParms = mstrStaticParms & "&EID=" & ReqZeroToNull("EligElementID")
mstrStaticParms = mstrStaticParms & "&ASR=" & ReqIsDate("StartReviewMonth")
mstrStaticParms = mstrStaticParms & "&AER=" & ReqIsDate("EndReviewMonth")
mstrStaticParms = mstrStaticParms & "&ARN=" & ReqIsBlank("Reviewer")
mstrStaticParms = mstrStaticParms & "&ARRN=" & ReqIsBlank("ReReviewer")
mstrStaticParms = mstrStaticParms & "&ARM=" & ReqIsNumeric("ReportingMode")
mstrStaticParms = mstrStaticParms & "&ADPD=" & ReqZeroToNull("DaysPastDue")
mstrStaticParms = mstrStaticParms & "&ATID=" & ReqZeroToNull("TabID")
mstrStaticParms = mstrStaticParms & "&AFID=" & ReqZeroToNull("FactorID")
mstrStaticParms = mstrStaticParms & "&ARRT=" & ReqIsBlank("ReReviewTypeID")
mstrStaticParms = mstrStaticParms & "&ASDT=" & ReqForm("ShowDetail")

maStaffParms(0) = ReqIsBlank("Worker")
maStaffParms(1) = ReqIsBlank("Supervisor")
maStaffParms(2) = ReqIsBlank("ProgramManager")
maStaffParms(3) = ReqIsBlank("Office")
maStaffParms(4) = ReqIsBlank("Director")

Sub WriteColumn(intColID, strClass, intValue, intLeft, strBackColor, strStyleAdd, intWidth, intNumAfterDecimal)
    Response.Write "<SPAN id=lblCol" & intColID & oCounts.RowID & " class=" & strClass & " " & vbCrLf
    If intValue > 0 Then
        Response.Write "onmouseover=""Call ColMouseEvent(0," & intColID & "," & oCounts.RowID & ")"" onmouseout=""Call ColMouseEvent(1," & intColID & "," & oCounts.RowID & ")"" onclick=""Call ColClickEvent(" & intColID & "," & oCounts.RowID & ")"" " & vbCrLf
        Response.Write "style=""cursor:hand;WIDTH:" & intWidth & "; LEFT:" & intLeft & "; TEXT-ALIGN:center; COLOR:blue; BACKGROUND:" & strBackCOlor & strStyleAdd & """>" & vbCrLf
    Else
        Response.Write "style=""WIDTH:" & intWidth & "; LEFT:" & intLeft & "; TEXT-ALIGN:center; BACKGROUND:" & strBackCOlor & strStyleAdd & """>" & vbCrLf
    End If
    Response.Write FormatNumber(intValue,intNumAfterDecimal,True,False,True) & "</SPAN>" & vbCrLf
End Sub

Sub WriteColumnNoClass(intColID, strClass, intValue, intLeft, strBackColor, strStyleAdd, intRowID)
    Response.Write "<SPAN id=lblCol" & intColID & intRowID & " class=" & strClass & " " & vbCrLf
    If intValue > 0 Then
        Response.Write "onmouseover=""Call ColMouseEvent(0," & intColID & "," & intRowID & ")"" onmouseout=""Call ColMouseEvent(1," & intColID & "," & intRowID & ")"" onclick=""Call ColClickEvent(" & intColID & "," & intRowID & ")"" " & vbCrLf
        Response.Write "style=""cursor:hand;WIDTH:90; LEFT:" & intLeft & "; TEXT-ALIGN:center; COLOR:blue; BACKGROUND:" & strBackCOlor & strStyleAdd & """>" & vbCrLf
    Else
        Response.Write "style=""WIDTH:90; LEFT:" & intLeft & "; TEXT-ALIGN:center; COLOR:black; BACKGROUND:" & strBackCOlor & strStyleAdd & """>" & vbCrLf
    End If
    Response.Write FormatNumber(intValue,0,True,False,True) & "</SPAN>" & vbCrLf
End Sub

Sub WriteColumnNoClassPercent(intColID, strClass, intValue, intLeft, strBackColor, strStyleAdd, intRowID, blnDrillDown)
    Response.Write "<SPAN id=lblCol" & intColID & intRowID & " class=" & strClass & " " & vbCrLf
    If intValue > 0 And blnDrillDown = True Then
        Response.Write "onmouseover=""Call ColMouseEvent(0," & intColID & "," & intRowID & ")"" onmouseout=""Call ColMouseEvent(1," & intColID & "," & intRowID & ")"" onclick=""Call ColClickEvent(" & intColID & "," & intRowID & ")"" " & vbCrLf
        Response.Write "style=""cursor:hand;WIDTH:90; LEFT:" & intLeft & "; TEXT-ALIGN:center; COLOR:blue; BACKGROUND:" & strBackCOlor & strStyleAdd & """>" & vbCrLf
    Else
        Response.Write "style=""WIDTH:90; LEFT:" & intLeft & "; TEXT-ALIGN:center; COLOR:black; BACKGROUND:" & strBackCOlor & strStyleAdd & """>" & vbCrLf
    End If
    Response.Write FormatNumber(intValue,2,True,False,True) & "%</SPAN>" & vbCrLf
End Sub

Function CreateDrillDownCell(intColID, strValue, intWidth, intRowID, blnNumeric, strClass, strStyle)
    Dim strCell, strCursor
    
    strCell = "<TD id=lblCol" & intColID & intRowID & " class=" & strClass & " "
    If blnNumeric = True Then
        If CLng(strValue) > 0 Then
            strCell = strCell & "onmouseover=""Call ColMouseEvent(0," & intColID & "," & intRowID & ")"" onmouseout=""Call ColMouseEvent(1," & intColID & "," & intRowID & ")"" onclick=""Call ColClickEvent(" & intColID & "," & intRowID & ")"" " & vbCrLf
            strCursor = ";cursor:hand;color:blue;"
        Else
            strCursor = ""
        End If
        strCell = strCell & "style=""" & strStyle & ";width:" & intWidth & strCursor & """>" & FormatNumber(strValue,0,True,False,True) & "</TD>" & vbCrLf
    Else
        strCell = strCell & "onmouseover=""Call ColMouseEvent(0," & intColID & "," & intRowID & ")"" onmouseout=""Call ColMouseEvent(1," & intColID & "," & intRowID & ")"" onclick=""Call ColClickEvent(" & intColID & "," & intRowID & ")"" " & vbCrLf
        strCursor = ";cursor:hand;color:blue;"
        strCell = strCell & "style=""" & strStyle & ";width:" & intWidth & strCursor & """>" & strValue & "</TD>" & vbCrLf
    End If
    CreateDrillDownCell = strCell
End Function

Sub WriteNames(intWhich)
    Dim intI
    Dim strNames
    
    strNames = ""
    For intI = 0 To icDir
        If intI < intWhich Then
            strNames = strNames & maStaffParms(intI) & "^"
        Else
            strNames = strNames & Replace(Replace(oCounts.CurrentName(intI),"[",""),"]","") & "^"
        End If
    Next
    Response.Write vbCrLf & "<INPUT type=""hidden"" id=txtDrillDownNames" & oCounts.RowID & " value=""" & strNames & """></INPUT>" & vbCrLf
End Sub

Sub WriteColumnHeader(strText, intLeft, intWidth, strBorders, strBackColor, intColumnID)
    If strText = "[BLANK]" Then
        Response.Write "<SPAN id=lblBlankRow class=ColumnHeading " & vbCrLf
        Response.Write "style=""LEFT:10; WIDTH:630; BORDER-BOTTOM-STYLE:none; Background:" & strBackColor & """>" & vbCrLf
        Response.Write "</SPAN>" & vbCrLf
    Else
        Response.Write "<SPAN id=lblColumnHeader" & intColumnID & " class=ColumnHeading " & vbCrLf
        Response.Write "style=""LEFT:" & intLeft & ";WIDTH:" & intWidth & ";" & strBorders & "; Background:" & strBackColor & """>" & vbCrLf
        Response.Write "" & strText & "" & "</SPAN>" & vbCrLf
    End If
End Sub
%>