<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: EligElemSum.asp                                                 '
'  Purpose: Displays the Eligibility Element Summary report, based on the   '
'           criteria passed to this page by the previous criteria screen.   '
'==========================================================================='
Dim mstrPageTitle 
Dim adRs
Dim adCmd
Dim intTotalScreens, intTotalYes, intTotalNo, intTotalNA
Dim intTotalNonNA
Dim intI
Dim dblPercent
Dim intShadeCount
Dim strColor
Dim mstrPaymentRate
Dim mintRowID

%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<!--#include file="IncDrillDownSvr.asp"-->
<%
mstrPaymentRate = GetAppSetting("AccuracyErrorRate") 
If IsNull(mstrPaymentRate) Then mstrPaymentRate = "Error"
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
'Retrieve the records that match the report criteria:
Set adCmd = GetAdoCmd("spRptEligElemSum")
    AddParmIn adCmd, "@AliasID", adInteger, 0, glngAliasPosID
    AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
    AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
    AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
    AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, ReqIsDate("StartDate")
    AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
    AddParmIn adCmd, "@Director", adVarchar, 50, ReqIsBlank("Director")
    AddParmIn adCmd, "@Office", adVarchar, 50, ReqIsBlank("Office")
    AddParmIn adCmd, "@Manager", adVarchar, 50, ReqIsBlank("ProgramManager")
    AddParmIn adCmd, "@Supervisor", adVarchar, 50, ReqIsBlank("Supervisor")
    AddParmIn adCmd, "@WorkerName", adVarchar, 50, ReqIsBlank("Worker")
    AddParmIn adCmd, "@ReviewTypeID", adVarChar, 100, ReqIsBlank("ReviewTypeID")
    AddParmIn adCmd, "@ReviewClassID", adVarChar, 100, ReqIsBlank("ReviewClassID")
    AddParmIn adCmd, "@ProgramID", adInteger, 0, ReqZeroToNull("ProgramID")
    AddParmIn adCmd, "@ElementID", adInteger, 0, ReqZeroToNull("EligElementID")
    AddParmIn adCmd, "@StartReviewMonth", adDBTimeStamp, 0, ReqIsDate("StartReviewMonth")
    AddParmIn adCmd, "@EndReviewMonth", adDBTimeStamp, 0, ReqIsDate("EndReviewMonth")
    AddParmIn adCmd, "@DrillDownID", adInteger, 0, Null
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
%>

<HTML>
<HEAD>
    <TITLE><%=ReqForm("ReportTitle")%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncRptStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim blnCloseClicked

<!--
Sub window_onload
	Call FormShow("none")
	PageBody.style.cursor = "wait"
    If Form.UserID.Value = "" Then
        MsgBox "User not recognized.  Logon failed, please try again.", vbinformation, "Log On"
        window.navigate "Logon.asp"
    End If
	Call SizeAndCenterWindow(767, 520, True)
    Call FormShow("")
    PageBody.style.cursor = "default"
    cmdPrint1.focus
End Sub

Sub cmdClose_onclick()	
	Window.close
End Sub

Sub FormShow(strVis)
	cmdPrint1.style.display = strVis
    cmdClose1.style.display = strVis
    cmdPrint2.style.display = strVis
    cmdClose2.style.display = strVis
    cmdExport1.style.display = strVis
    cmdExport2.style.display = strVis
    Header.style.display = strVis
    PageFrame.style.display = strVis
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub window_onbeforeprint()
    cmdPrint1.style.visibility = "hidden"
    cmdClose1.style.visibility = "hidden"
    cmdPrint2.style.visibility = "hidden"
    cmdClose2.style.visibility = "hidden"
    cmdExport1.style.visibility = "hidden"
    cmdExport2.style.visibility = "hidden"
End Sub

Sub window_onafterprint()
    cmdPrint1.style.visibility = "visible"
    cmdClose1.style.visibility = "visible"
    cmdPrint2.style.visibility = "visible"
    cmdClose2.style.visibility = "visible"
    cmdExport1.style.visibility = "visible"
    cmdExport2.style.visibility = "visible"
End Sub

Sub CallFieldReport(intRowID)
    intActionID = document.all("txtElementID" & intRowID).value
    Dim strParms
    Dim strURL
    Dim strReturnValue
    Dim intElementID
    
    intElementID = Document.all("txtElementID" & intRowID).value
    
    strParms = "?RT=Causal Factor Summary"
    strParms = strParms & "&PT=<%=ReqForm("ProgramText")%>"
    strParms = strParms & "&A1=<%=glngAliasPosID%>"
    strParms = strParms & "&A2=<%=gblnUserAdmin%>"
    strParms = strParms & "&A3=<%=gblnUserQA%>"
    strParms = strParms & "&A4=<%=gstrUserID%>"
    strParms = strParms & "&A5=<%=ReqForm("StartDate")%>"
    strParms = strParms & "&A6=<%=ReqForm("EndDate")%>"
    strParms = strParms & "&A7=<%=ReqForm("Director")%>"
    strParms = strParms & "&A8=<%=ReqForm("Office")%>"
    strParms = strParms & "&A9=<%=ReqForm("ProgramManager")%>"
    strParms = strParms & "&A10=<%=ReqForm("Supervisor")%>"
    strParms = strParms & "&A11=<%=ReqForm("Worker")%>"
    strParms = strParms & "&A12=<%=ReqForm("ReviewTypeID")%>"
    strParms = strParms & "&A13=<%=ReqForm("ReviewClassID")%>"
    strParms = strParms & "&A14=<%=ReqForm("ProgramID")%>"
    strParms = strParms & "&A15=" & intElementID
    strParms = strParms & "&A16=<%=ReqForm("StartReviewMonth")%>"
    strParms = strParms & "&A17=<%=ReqForm("EndReviewMonth")%>"
    strParms = strParms & "&A19=<%=ReqForm("ShowDetail")%>"
    strParms = strParms & "&AT12=<%=ReqForm("ReviewTypeText")%>"
    strParms = strParms & "&AT13=<%=ReqForm("ReviewClassText")%>"
    strParms = strParms & "&AT15=" & Replace(document.all("lblCol0" & intRowID).innerText,"&","[AMP]")
    
    strURL = "RptCausalFactorSummary.asp" & strParms
    strReturnValue = window.showModalDialog(strURL, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
End Sub

Sub ColClickEvent(intColID, intRowID)
    Dim intActionID
    Dim strField

    If CInt(intRowID) < 999 Then
        intActionID = document.all("txtElementID" & intRowID).value
        Select Case intColID
            Case 1
                strField = "All"
            Case 2
                strField = "NA"
            Case 3
                strField = "Yes"
            Case 4
                strField = "No"
        End Select
        strField = document.all("lblCol0" & intRowID).innerText & "&SN2=Status: " & strField
    Else
        intActionID = 0
        Select Case intColID
            Case 5
                strField = "Total Screens"
            Case 6
                strField = "Total Screens NA"
            Case 7
                strField = "Total Screens Yes"
            Case 8
                strField = "Total Screens No"
        End Select
    End If
    Call DrillDownColClickEventNoStaff("spRptEligElemSum", intColID, intRowID, intActionID, strField)
End Sub 

-->
</SCRIPT>
<!--#include file="IncRptExpParms.asp"-->
<!--#include file="IncCmnCliFunctions.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<!--#include file="IncRptHeader.asp"-->
<!--#include file="IncDrillDownCli.asp"-->
<DIV id=PageFrame style="HEIGHT:225; WIDTH:650; TOP:116; LEFT:10; FONT-SIZE:10pt; padding-top:5">
<BR>
<%
Call WriteCriteria()
Dim maColumns(6)
maColumns(1) = 240
maColumns(2) = 310
maColumns(3) = 380
maColumns(4) = 450
maColumns(5) = 520
maColumns(6) = 590

strColor = "#FFEFD5"

Response.Write "<BR>"
Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
Response.Write "style=""WIDTH:630; LEFT:10; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "</SPAN>"

Response.Write "<SPAN id=lblErrorPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(1) & "; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Total</SPAN>"

Response.Write "<SPAN id=lblCorrectCntHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(2) & "; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Total</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(3) & "; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Total</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(4) & "; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Percent</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(5) & "; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Total</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(6) & "; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Percent</SPAN>"

Response.Write "<BR>"
Response.Write "<SPAN id=lblHdr class=ColumnHeading "
Response.Write "style=""WIDTH:630; LEFT:10;BORDER-TOP-STYLE:none;background:" & strColor & """></SPAN>"

Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
Response.Write "style=""overflow:visible;text-align:left;WIDTH:375; LEFT:10;BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "Element Name</SPAN>"

Response.Write "<SPAN id=lblErrorPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:95; LEFT:" & maColumns(1) & ";BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "Elements</SPAN>"

Response.Write "<SPAN id=lblCorrectCntHdr class=ColumnHeading "
Response.Write "style=""WIDTH:70; LEFT:" & maColumns(2)+10 & ";BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "NA</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(3) & ";BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "Yes</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(4) & ";BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "Yes</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(5) & ";BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "No</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:" & maColumns(6) & ";BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "No</SPAN>"

Response.Write "<br><br>"

If adRs.EOF Then
    Response.Write "<BR><BR>"
    Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
    Response.Write " * No reviews matched the report criteria *"
End If

mintRowID = 0
intTotalScreens = 0
intTotalYes = 0
intTotalNo = 0
intTotalNA = 0
Do While Not adRs.EOF

    intTotalScreens = intTotalScreens + adRs.Fields("TotalYes").Value + adRs.Fields("TotalNo").Value + adRs.Fields("TotalNA").Value
    intTotalYes = intTotalYes + adRs.Fields("TotalYes").Value
    intTotalNo = intTotalNo + adRs.Fields("TotalNo").Value
    intTotalNA = intTotalNA + adRs.Fields("TotalNA").Value
    If intShadeCount MOD 2 = 0 Then
        strColor = "#ffffff"
    Else 
        strColor = "#FFEFD5"
    End If
    Call WriteLine(adRs.Fields("elmShortName").Value, strColor, _
        adRs.Fields("TotalYes").Value, adRs.Fields("TotalNo").Value, adRs.Fields("TotalNA").Value, adRs.Fields("elmID").value)
    intShadeCount = intShadeCount + 1

    adRs.MoveNext
Loop
Response.Write "<BR><BR>"

If intTotalScreens - intTotalNA > 0 Then
    intTotalNonNA = intTotalScreens - intTotalNA
    Call WriteTotalLine(5, "Total Elements", intTotalScreens, False, False, False)
    Call WriteTotalLine(6, "Total NA", intTotalNA, False, False, False)
    Call WriteTotalLine(7, "Total Yes", intTotalYes, False, False, False)
    Call WriteTotalLine(17, "Percent Yes", (intTotalYes/intTotalNonNA)*100, False, True, False)
    Call WriteTotalLine(8, "Total No", intTotalNo, False, False, False)
    Call WriteTotalLine(18, "Percent No", (intTotalNo/intTotalNonNA)*100, False, True, False)

    Response.Write "<BR><BR>"
End If

Sub WriteTotalLine(intColID, strText, strValue, blnDrillDown, blnPercent, blnHeading)
    Dim intTextWidth
    
    intTextWidth = 250
    If blnHeading Then
        Response.Write "<BR>"
        Response.Write "<SPAN id=lblRowText" & intColID & " Class=ManagementText "
        Response.Write "style=""WIDTH:700; LEFT:5; BORDER-STYLE:none"">"
        Response.Write "<B>" &strText & "</B></SPAN>"
    Else
        Response.Write "<SPAN id=lblRowText" & intColID & " Class=ManagementText "
        Response.Write "style=""WIDTH:" & intTextWidth & "; LEFT:10; BORDER-STYLE:solid;BORDER-WIDTH:1"">"
        Response.Write strText & ": </SPAN>"

        Response.Write "<SPAN id=lblCol" & intColID & "999 Class=ManagementText "
        If blnDrillDown And CLng(strValue) <> 0 Then
            Response.Write "style=""cursor:hand;color:blue;border-color:black;LEFT:" & intTextWidth+10 & ";width:80; text-align:center;BORDER-STYLE:solid;BORDER-WIDTH:1""" & vbCrLf
            Response.Write "onmouseover=""Call ColMouseEvent(0," & intColID & ",999)"" onmouseout=""Call ColMouseEvent(1," & intColID & ",999)"" onclick=""Call ColClickEvent(" & intColID & ",999)"">" & vbCrLf
        Else
            Response.Write "style=""LEFT:" & intTextWidth+10 & ";width:80; text-align:center;BORDER-STYLE:solid;BORDER-WIDTH:1"">" & vbCrLf
        End If
        If blnPercent Then
            If CDbl(strValue) > 0 Then
                Response.Write FormatNumber(strValue, 2, True, True, True) & "%</B></SPAN>"
            Else
                Response.Write "---</B></SPAN>"
            End If
        Else
            Response.Write FormatNumber(strValue, 0, True, True, True) & "</B></SPAN>"
        End If
    End If
    Response.Write "<BR>"
End Sub

Sub WriteLine(strElement, strColor, intTotalYes, intTotalNo, intTotalNA, intElementID)
    Dim dblPercent
    Dim intTotalScreens
    
    mintRowID = mintRowID + 1
    intTotalScreens = intTotalYes+intTotalNo '+intTotalNA
	Response.Write "<SPAN id=lblElement class=ReportText "
    Response.Write "style=""WIDTH:630; LEFT:10;TEXT-ALIGN:left;background:" & strColor & """></SPAN>"

    Response.Write "<INPUT id=txtElementID" & mintRowID & " type=hidden value=" & intElementID & ">"

    Response.Write "<SPAN id=lblCol0" & mintRowID & " class=ReportText " & vbCrLf
    If intTotalYes+intTotalNo+intTotalNA > 0 Then
        Response.Write "onmouseover=""Call ColMouseEvent(0,0," & mintRowID & ")"" onmouseout=""Call ColMouseEvent(1,0," & mintRowID & ")"" onclick=""Call CallFieldReport(" & mintRowID & ")"" " & vbCrLf
        Response.Write "style=""cursor:hand;COLOR:blue;WIDTH:240; LEFT:10;OverFlow:hidden; TEXT-ALIGN:left;background:" & strColor & """>" & vbCrLf
    Else
        Response.Write "style=""WIDTH:240; LEFT:10;OverFlow:hidden; TEXT-ALIGN:left;background:" & strColor & """>" & vbCrLf
    End If
    Response.Write strElement & "</SPAN>"
    
    Call WriteColumnNoClass(1,"ReportText",intTotalYes+intTotalNo+intTotalNA,maColumns(1),strColor,"",mintRowID)
    Call WriteColumnNoClass(2,"ReportText",intTotalNA,maColumns(2),strColor,"",mintRowID)
    Call WriteColumnNoClass(3,"ReportText",intTotalYes,maColumns(3),strColor,"",mintRowID)
    If intTotalScreens > 0 Then
        dblPercent = (intTotalYes / intTotalScreens) * 100
    Else
        dblPercent = 0
    End If
    Call WriteColumnNoClassPercent(5,"ReportText",dblPercent,maColumns(4),strColor,"",mintRowID, False)
    If intTotalScreens > 0 Then
        dblPercent = (intTotalNo / intTotalScreens) * 100
    Else
        dblPercent = 0
    End If
    Call WriteColumnNoClass(4,"ReportText",intTotalNo,maColumns(5),strColor,"",mintRowID)
    Call WriteColumnNoClassPercent(6,"ReportText",dblPercent,maColumns(6),strColor,"",mintRowID, False)
    
    Response.Write "<BR>"
End Sub
%>
<!--#include file="IncRptFooter.asp"-->
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->