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
Dim intHoldTabID, intHoldProgramID, blnDI
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
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
'Retrieve the records that match the report criteria:
Set adCmd = GetAdoCmd("spRptReReviewEligElemSum")
    AddParmIn adCmd, "@AliasID", adInteger, 0, glngAliasPosID
    AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
    AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
    AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
    AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, ReqIsDate("StartDate")
    AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
    AddParmIn adCmd, "@Director", adVarChar, 50, ReqIsBlank("Director")
    AddParmIn adCmd, "@Office", adVarChar, 50, ReqIsBlank("Office")
    AddParmIn adCmd, "@PrgMgr", adVarChar, 50, ReqIsBlank("ProgramManager")
    AddParmIn adCmd, "@Reviewer", adVarChar, 50, ReqIsBlank("Reviewer")
    AddParmIn adCmd, "@ProgramID", adInteger, 0, ReqZeroToNull("ProgramID")
    AddParmIn adCmd, "@ElementID", adInteger, 0, Null 'ReqZeroToNull("EligElementID")
    AddParmIn adCmd, "@FactorID", adInteger, 0, Null ' ReqZeroToNull("FactorID")
    AddParmIn adCmd, "@ReReviewTypeID", adInteger, 0, ReqForm("ReReviewTypeID")
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

Sub ColClickEvent(intColID, intRowID)
    Dim intElementID, intFactorID
    Dim strField

    intElementID = Parse(document.all("txtElementID" & intRowID).value,"^",1)
    intFactorID = Parse(document.all("txtElementID" & intRowID).value,"^",2)
    strField = document.all("lblElement" & intRowID).innerText & "&DD2=" & intElementID & "&DD3=" & intFactorID
    Call DrillDownColClickEventNoStaff("spRptReReviewEligElemSum", intColID, intRowID, intFieldID, strField)
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

strColor = "#FFEFD5"

Response.Write "<BR>"
Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
Response.Write "style=""WIDTH:630; LEFT:10; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "</SPAN>"

Response.Write "<SPAN id=lblErrorPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:385; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Total</SPAN>"

Response.Write "<SPAN id=lblCorrectCntHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:475; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Percent</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:565; BORDER-BOTTOM-STYLE:none;background:" & strColor & """>"
Response.Write "Percent</SPAN>"

Response.Write "<BR>"
Response.Write "<SPAN id=lblHdr class=ColumnHeading "
Response.Write "style=""WIDTH:630; LEFT:10;BORDER-TOP-STYLE:none;background:" & strColor & """></SPAN>"

Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
Response.Write "style=""text-align:left;WIDTH:375; LEFT:10;BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "</SPAN>"

Response.Write "<SPAN id=lblErrorPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:385;BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "Cases</SPAN>"

Response.Write "<SPAN id=lblCorrectCntHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:475;BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "Accurate</SPAN>"

Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
Response.Write "style=""WIDTH:90; LEFT:565;BORDER-TOP-STYLE:none;background:" & strColor & """>"
Response.Write "Inaccurate</SPAN>"

Response.Write "<br><br>"

If adRs.EOF Then
    Response.Write "<BR><BR>"
    Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
    Response.Write " * No reviews matched the report criteria *"
End If

mintRowID = 0
intHoldTabID = 0
Do While Not adRs.EOF
    If intShadeCount MOD 2 = 0 Then
        strColor = "#ffffff"
    Else 
        strColor = "#FFEFD5"
    End If
    If intHoldTabID <> adRs.Fields("TypeID").Value Then
        Call WriteHeading(GetTabName(adRs.Fields("TypeID").Value), 12, True)
        intHoldTabID = adRs.Fields("TypeID").Value
        
        If intHoldTabID = 2 Then
            intHoldProgramID = 0
        End If
    End If
    If intHoldProgramID <> adRs.Fields("ProgramID").Value And intHoldTabID = 2 And ReqZeroToNull("ProgramID") = 6 Then
        Call WriteHeading(adRs.Fields("DIProgram").Value, 11, True)
        intHoldProgramID = adRs.Fields("ProgramID").Value
    End If
    If intHoldTabID = 2 Then
        Call WriteLine(adRs.Fields("elmShortName").Value & " - " & adRs.Fields("fctShortName").Value, strColor, _
            adRs.Fields("TotalCases").Value, adRs.Fields("TotalCorrect").Value, adRs.Fields("elmID").value & "^" & adRs.Fields("FactorID").value)
    Else
        Call WriteLine(adRs.Fields("elmShortName").Value, strColor, _
            adRs.Fields("TotalCases").Value, adRs.Fields("TotalCorrect").Value, adRs.Fields("elmID").value & "^0")
    End If
    intShadeCount = intShadeCount + 1

    adRs.MoveNext
Loop
Response.Write "<BR><BR>"

Function GetTabName(intTabID)
    Select Case intTabID
        Case 1
            GetTabName = "Action Integrity"
        Case 2
            GetTabName = "Data Integrity"
        Case 3
            GetTabName = "Information Gathering"
        Case Else
            GetTabName = "unknown"
    End Select
End Function

Sub WriteHeading(strText, intFontSize, blnBold)
    Response.Write "<BR>"
    Response.Write "<SPAN Class=ManagementText "
    Response.Write "style=""WIDTH:700; LEFT:5; BORDER-STYLE:none;font-size:" & intFontSize & """>"
    If blnBold Then
        Response.Write "<B>" & strText & "</B></SPAN><BR>"
    Else
        Response.Write strText & "</SPAN><BR>"
    End If
End Sub

Sub WriteLine(strElement, strColor, intTotalActions, intTotalCorrect, intElementID)
    Dim dblPercent
    
    mintRowID = mintRowID + 1
    
	Response.Write "<SPAN id=lblElement class=ReportText "
    Response.Write "style=""WIDTH:630; LEFT:10;TEXT-ALIGN:left;background:" & strColor & """></SPAN>"

    Response.Write "<INPUT id=txtElementID" & mintRowID & " type=hidden value=" & intElementID & ">"

    Response.Write "<SPAN id=lblElement" & mintRowID & " class=ReportText "
    Response.Write " style=""WIDTH:375; LEFT:10;OverFlow:hidden; TEXT-ALIGN:left;background:" & strColor & """>"
    Response.Write strElement & "</SPAN>"
    
    If intTotalActions > 0 Then
        dblPercent = (intTotalCorrect / intTotalActions) * 100
    Else
        dblPercent = 0
    End If
    Call WriteColumnNoClass(1,"ReportText",intTotalActions,385,strColor,"",mintRowID)
    Call WriteColumnNoClassPercent(2,"ReportText",dblPercent,475,strColor,"",mintRowID, True)
    If intTotalActions > 0 Then
        dblPercent = ((intTotalActions-intTotalCorrect) / intTotalActions) * 100
    Else
        dblPercent = 0
    End If
    Call WriteColumnNoClassPercent(3,"ReportText",dblPercent,565,strColor,"",mintRowID, True)
    
    Response.Write "<BR>"
End Sub
%>
<!--#include file="IncRptFooter.asp"-->
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->