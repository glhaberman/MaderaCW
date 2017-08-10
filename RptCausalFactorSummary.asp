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
Dim intTotalDecisions, intTotalYes, intTotalNo, intTotalNA
Dim intTotalNon30
Dim intI
Dim dblPercent
Dim intShadeCount
Dim strColor
Dim mstrPaymentRate
Dim mintRowID
Dim mstrReportTitle
Dim mstrProgramText

%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncDrillDownSvr.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
If Request.Form("ReportTitle") <> "" Then
    mstrReportTitle = ReqForm("ReportTitle")
    mstrProgramText = ReqForm("ProgramText")
    maCriteria(1) = ReqForm("RepLAliasPosID")
    maCriteria(2) = ReqForm("RepLUserAdmin")
    maCriteria(3) = ReqForm("RepLUserQA")
    maCriteria(4) = ReqForm("RepLUserID")
    maCriteria(5) = ReqForm("StartDate")
    maCriteria(6) = ReqForm("EndDate")
    maCriteria(7) = ReqForm("Director")
    maCriteria(8) = ReqForm("Office")
    maCriteria(9) = ReqForm("ProgramManager")
    maCriteria(10) = ReqForm("Supervisor")
    maCriteria(11) = ReqForm("Worker")
    maCriteria(12) = ReqForm("ReviewTypeID")
    maCriteriaText(12) = ReqForm("ReviewTypeText")
    maCriteria(13) = ReqForm("ReviewClassID")
    maCriteriaText(13) = ReqForm("ReviewClassText")
    maCriteria(14) = ReqForm("ProgramID")
    maCriteria(15) = ReqForm("EligElementID")
    maCriteriaText(15) = ReqForm("EligElementText")
    maCriteria(16) = ReqForm("StartReviewMonth")
    maCriteria(17) = ReqForm("EndReviewMonth")
    maCriteria(18) = ReqForm("FactorID")
    maCriteriaText(18) = ReqForm("FactorText")
    maCriteria(19) = ReqForm("ShowDetail")
Else
    mstrReportTitle = Request.QueryString("RT")
    mstrProgramText = Request.QueryString("PT")
    maCriteria(1) = Request.QueryString("A1")
    maCriteria(2) = Request.QueryString("A2")
    maCriteria(3) = Request.QueryString("A3")
    maCriteria(4) = Request.QueryString("A4")
    maCriteria(5) = Request.QueryString("A5")
    maCriteria(6) = Request.QueryString("A6")
    maCriteria(7) = Request.QueryString("A7")
    maCriteria(8) = Request.QueryString("A8")
    maCriteria(9) = Request.QueryString("A9")
    maCriteria(10) = Request.QueryString("A10")
    maCriteria(11) = Request.QueryString("A11")
    maCriteria(12) = Request.QueryString("A12")
    maCriteria(13) = Request.QueryString("A13")
    maCriteria(14) = Request.QueryString("A14")
    maCriteria(15) = Request.QueryString("A15")
    maCriteria(16) = Request.QueryString("A16")
    maCriteria(17) = Request.QueryString("A17")
    maCriteria(18) = "0" 'Request.QueryString("A18")
    maCriteria(19) = Request.QueryString("A19")
    maCriteriaText(12) = Request.QueryString("AT12")
    maCriteriaText(13) = Request.QueryString("AT13")
    maCriteriaText(15) = Request.QueryString("AT15")
End If

'Retrieve the records that match the report criteria:
Set adCmd = GetAdoCmd("spRptCausalFactor")
    AddParmIn adCmd, "@AliasID", adInteger, 0, maCriteria(1)
    AddParmIn adCmd, "@Admin", adBoolean, 0, maCriteria(2)
    AddParmIn adCmd, "@QA", adBoolean, 0, maCriteria(3)
    AddParmIn adCmd, "@UserID", adVarchar, 20, maCriteria(4)
    AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(maCriteria(5))
    AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(maCriteria(6))
    AddParmIn adCmd, "@Director", adVarChar, 50, IsBlank(maCriteria(7))
    AddParmIn adCmd, "@Office", adVarChar, 50, IsBlank(maCriteria(8))
    AddParmIn adCmd, "@Manager", adVarChar, 50, IsBlank(maCriteria(9))
    AddParmIn adCmd, "@Supervisor", adVarChar, 50, IsBlank(maCriteria(10))
    AddParmIn adCmd, "@WorkerName", adVarchar, 50, IsBlank(maCriteria(11))
    AddParmIn adCmd, "@ReviewTypeID", adVarChar, 100, IsBlank(maCriteria(12))
    AddParmIn adCmd, "@ReviewClassID", adVarChar, 100, IsBlank(maCriteria(13))
    AddParmIn adCmd, "@ProgramID", adInteger, 0, IsBlank(maCriteria(14))
    AddParmIn adCmd, "@ElementID", adInteger, 0, ZeroToNull(maCriteria(15))
    AddParmIn adCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(maCriteria(16))
    AddParmIn adCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(maCriteria(17))
    AddParmIn adCmd, "@FactorID", adInteger, 0, ZeroToNull(maCriteria(18))
    AddParmIn adCmd, "@DrillDownID", adInteger, 0, Null
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
If maCriteria(19) <> "Y" Then
    'If "Include All Factors" is not checked, filter out causal factors that are all NA.
    adRs.Filter = "TotYes>0 Or TotNo>0"
End If
adRs.Sort = "elmSortOrder, TotNo DESC, fctShortName"
%>

<HTML>
<HEAD>
    <TITLE><%=mstrReportTitle%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncRptStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim blnCloseClicked

<!--
Sub window_onload
	Call FormShow("none")
	PageBody.style.cursor = "wait"
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

Sub ColClickEvent(intColID, intRowID)
    Dim intFieldID
    Dim strField

    If CInt(intRowID) < 999 Then
        intFieldID = document.all("txtFactorID" & intRowID).value
        Select Case intColID
            Case 1
                strField = "All"
            Case 2
                strField = "NA"
            Case 3
                strField = "Yes"
        End Select
        strField = document.all("lblFactor" & intRowID).innerText & "&SN3=Status: " & strField
    Else
        intFieldID = 0
        Select Case intColID
            Case 5
                strField = "Total Decisions"
            Case 6
                strField = "Total Decisions - Yes"
            Case 7
                strField = "Total Decisions - No"
        End Select
    End If

    Call DrillDownColClickEventNoStaff("spRptCausalFactor", intColID, intRowID, intFieldID, strField)
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
If Request.Form("ReportTitle") <> "" Then
    Call WriteCriteria()
Else
    Call ReqQS_WriteCriteria()
End If
Dim maColumns(4)
Dim mstrHoldElement
maColumns(1) = 360
maColumns(2) = 415
maColumns(3) = 496
maColumns(4) = 570

strColor = "#FFEFD5"

If adRs.EOF Then
    Response.Write "<BR><BR>"
    Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
    Response.Write " * No reviews matched the report criteria *"
End If

mintRowID = 0
intTotalDecisions = 0
intTotalYes = 0
intTotalNo = 0
intTotalNA = 0
mstrHoldElement = ""
Do While Not adRs.EOF
    If mstrHoldElement <> adRs.Fields("elmShortName").value Then
        intShadeCount = 0
        If mstrHoldElement <> "" Then
            Response.Write "<BR><BR>"
        End If
	    Response.Write "<SPAN id=lblElement class=ReportText "
        Response.Write "style=""font-size:14;WIDTH:630; LEFT:0;TEXT-ALIGN:left;""><B>" & adRs.Fields("elmShortName").value & "</B></SPAN><BR>"
        mstrHoldElement = adRs.Fields("elmShortName").value
        Call WriteColumnHeaders()
    End If
    intTotalDecisions = intTotalDecisions + adRs.Fields("TotYes").Value + adRs.Fields("TotNo").Value + adRs.Fields("TotNA").Value
    intTotalYes = intTotalYes + adRs.Fields("TotYes").Value
    intTotalNo = intTotalNo + adRs.Fields("TotNo").Value
    intTotalNA = intTotalNA + adRs.Fields("TotNA").Value
    If intShadeCount MOD 2 = 0 Then
        strColor = "#ffffff"
    Else 
        strColor = "#FFEFD5"
    End If
    Call WriteLine(adRs.Fields("fctShortName").Value, strColor, _
        adRs.Fields("TotYes").Value, adRs.Fields("TotNo").Value, adRs.Fields("TotNA").Value, adRs.Fields("fctID").value & "^" & adRs.Fields("elmID").value & "^" & adRs.Fields("elmShortName").value)
    intShadeCount = intShadeCount + 1

    adRs.MoveNext
Loop

If intTotalDecisions - intTotalNA > 0 Then
    Response.Write "<BR><BR>"
    intTotalDecisions = intTotalDecisions - intTotalNA
    Call WriteTotalLine(5, "Total Causal Factors (NA Excluded)", intTotalDecisions, False, False, False)
    Call WriteTotalLine(6, "Total Yes", intTotalYes, False, False, False)
    Call WriteTotalLine(11, "Percent Yes", (intTotalYes/intTotalDecisions)*100, False, True, False)
    Call WriteTotalLine(7, "Total No", intTotalNo, False, False, False)
    Call WriteTotalLine(11, "Percent No", (intTotalNo/intTotalDecisions)*100, False, True, False)

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

Sub WriteLine(strFactor, strColor, intTotYes, intTotNo, intTotNA, intFactorID)
    Dim dblPercent
    Dim intTotalDecisions
    
    mintRowID = mintRowID + 1
    intTotalDecisions = intTotYes+intTotNo '+intTotNA
	Response.Write "<SPAN id=lblElement class=ReportText "
    Response.Write "style=""WIDTH:630; LEFT:10;TEXT-ALIGN:left;background:" & strColor & """></SPAN>"

    Response.Write "<INPUT id=txtFactorID" & mintRowID & " type=hidden value=""" & intFactorID & """>"

    Response.Write "<SPAN id=lblFactor" & mintRowID & " class=ReportText "
    Response.Write " style=""WIDTH:370; LEFT:10;OverFlow:hidden; TEXT-ALIGN:left;background:" & strColor & """>"
    Response.Write strFactor & "</SPAN>"
    
    Call WriteColumnNoClass(1,"ReportText", intTotYes+intTotNo, maColumns(2), strColor,"",mintRowID)
    'Call WriteColumnNoClass(2,"ReportText", intTotNA, maColumns(2), strColor,"",mintRowID)
    If intTotalDecisions > 0 Then
        dblPercent = (intTotYes / intTotalDecisions) * 100
    Else
        dblPercent = 0
    End If
    Call WriteColumnNoClassPercent(3,"ReportText", dblPercent, maColumns(3), strColor,"",mintRowID, True)
    If intTotalDecisions > 0 Then
        dblPercent = (intTotNo / intTotalDecisions) * 100
    Else
        dblPercent = 0
    End If
    Call WriteColumnNoClassPercent(4,"ReportText", dblPercent, maColumns(4), strColor,"",mintRowID, True)
    Response.Write "<BR>"
End Sub

Sub WriteColumnHeaders()
    Dim strBColor 
    strBColor = "#FFEFD5"
    Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
    Response.Write "style=""WIDTH:620; LEFT:10; BORDER-BOTTOM-STYLE:none;background:" & strBColor & """>"
    Response.Write "</SPAN>"

    Response.Write "<SPAN id=lblErrorPercentHdr class=ColumnHeading "
    Response.Write "style=""WIDTH:90; LEFT:" & maColumns(2) & "; BORDER-BOTTOM-STYLE:none;background:" & strBColor & """>"
    Response.Write "Factors</SPAN>"

    'Response.Write "<SPAN id=lblCorrectCntHdr class=ColumnHeading "
    'Response.Write "style=""WIDTH:90; LEFT:" & maColumns(2) & "; BORDER-BOTTOM-STYLE:none;background:" & strBColor & """>"
    'Response.Write "Total</SPAN>"

    Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
    Response.Write "style=""WIDTH:90; LEFT:" & maColumns(3) & "; BORDER-BOTTOM-STYLE:none;background:" & strBColor & """>"
    Response.Write "Percent</SPAN>"

    Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
    Response.Write "style=""WIDTH:90; LEFT:" & maColumns(4) & "; BORDER-BOTTOM-STYLE:none;background:" & strBColor & """>"
    Response.Write "Percent</SPAN>"

    Response.Write "<BR>"
    Response.Write "<SPAN id=lblHdr class=ColumnHeading "
    Response.Write "style=""WIDTH:620; LEFT:10;BORDER-TOP-STYLE:none;background:" & strBColor & """></SPAN>"

    Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
    Response.Write "style=""text-align:left;WIDTH:375; LEFT:10;BORDER-TOP-STYLE:none;background:" & strBColor & """>"
    Response.Write "Causal Factor</SPAN>"

    Response.Write "<SPAN id=lblErrorPercentHdr class=ColumnHeading "
    Response.Write "style=""WIDTH:90; LEFT:" & maColumns(2) & ";BORDER-TOP-STYLE:none;background:" & strBColor & """>"
    Response.Write "Reviewed</SPAN>"

    'Response.Write "<SPAN id=lblCorrectCntHdr class=ColumnHeading "
    'Response.Write "style=""WIDTH:90; LEFT:" & maColumns(2) & ";BORDER-TOP-STYLE:none;background:" & strBColor & """>"
    'Response.Write "NA</SPAN>"

    Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
    Response.Write "style=""font-size:11;WIDTH:90; LEFT:" & maColumns(3) & ";BORDER-TOP-STYLE:none;background:" & strBColor & """>"
    Response.Write "Yes</SPAN>"

    Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
    Response.Write "style=""font-size:11;WIDTH:90; LEFT:" & maColumns(4) & ";BORDER-TOP-STYLE:none;background:" & strBColor & """>"
    Response.Write "No</SPAN><BR><BR style=""font-size:6"">"
End Sub
%>
<!--#include file="IncRptFooter.asp"-->
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->