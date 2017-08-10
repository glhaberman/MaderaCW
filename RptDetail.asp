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
Dim adRs, adRsElms, adRsFacs, adRsComs, adRsAuds
Dim adCmd
Dim intI, intJ
Dim intShadeCount
Dim strColor, strRecord
Dim mintRowID, mstrShowDetail

%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<!--#include file="IncDrillDownSvr.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
'Retrieve the records that match the report criteria:
Set adCmd = GetAdoCmd("spRptCaseDetail")
    AddParmIn adCmd, "@AliasID", adInteger, 0, glngAliasPosID
    AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
    AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
    AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
    AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, ReqIsDate("StartDate")
    AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
    AddParmIn adCmd, "@Director", adVarchar, 50, ReqIsBlank("Director")
    AddParmIn adCmd, "@Office", adVarchar, 50, ReqIsBlank("Office")
    AddParmIn adCmd, "@Manager", adVarchar, 50, ReqIsBlank("ProgramManager")
    AddParmIn adCmd, "@Supervisor", adVarchar, 100, ReqIsBlank("Supervisor")
    AddParmIn adCmd, "@WorkerName", adVarchar, 100, ReqIsBlank("Worker")
    AddParmIn adCmd, "@ReviewTypeID", adVarChar, 100, ReqIsBlank("ReviewTypeID")
    AddParmIn adCmd, "@ReviewClassID", adVarChar, 100, ReqIsBlank("ReviewClassID")
    AddParmIn adCmd, "@CaseNumber", adVarChar, 15, ReqIsBlank("CaseNumber")
    AddParmIn adCmd, "@Programs", adInteger, 0, ReqZeroToNull("ProgramID")
    AddParmIn adCmd, "@ShowDetail", adChar, 1, ReqForm("ShowDetail")
    AddParmIn adCmd, "@IncludeCorrect", adChar, 1, ReqForm("IncludeCorrect")
    AddParmIn adCmd, "@StartReviewMonth", adDBTimeStamp, 0, ReqIsDate("StartReviewMonth")
    AddParmIn adCmd, "@EndReviewMonth", adDBTimeStamp, 0, ReqIsDate("EndReviewMonth")
    AddParmOut adCmd, "@ReturnVal", adInteger, 0 
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)    

mstrShowDetail = ReqForm("ShowDetail")
If mstrShowDetail = "Y" And adRs.RecordCount > 0 Then
    Set adRsElms = adRs.NextRecordSet
    Set adRsFacs = adRs.NextRecordSet
    Set adRsComs = adRs.NextRecordSet
    Set adRsAuds = adRs.NextRecordSet
End If
%>

<HTML>
<HEAD>
    <TITLE><%=ReqForm("ReportTitle")%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncRptStyles.asp"-->        
    <STYLE>
    .RowHeader
        {
        height: 15;
        padding-left: 5px;
        padding-right: 5px;
        border-style: solid;
        border-width: 1px;
        font-family: Tahoma;
        font-size:8pt;
        font-weight: bold;
        text-align: left;
        color:black;
        background-color: beige;
        overflow: hidden
        }
    .GroupHeader
        {
        height: 15;
        padding-left: 5px;
        padding-right: 5px;
        border-style: none;
        font-family: Tahoma;
        font-size:8pt;
        font-weight: bold;
        text-align: left;
        color:black;
        background-color: white;
        overflow: hidden
        }
    .TableRow
        {
        font-family: Tahoma;
        font-size:10pt;
        background-color: white;
        color: black;
        overflow: visible;
        }
    </STYLE>
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
    If "<%=mstrShowDetail%>" = "N" Then
        If "<%=ReqForm("ProgramID")%>" = "0" Then
            lblAppTitle.innerText = "Case Review Listing: All Programs"
            lblAppTitle.style.fontweight = "bold"
        End If
    Else
        If "<%=ReqForm("ProgramID")%>" = "0" Then
            lblAppTitle.innerText = "Worker Case Review Listing: All Programs"
            lblAppTitle.style.fontweight = "bold"
        Else
            lblAppTitle.innerText = "Worker Case Review Listing: <%=ReqForm("ProgramText")%>"
            lblAppTitle.style.fontweight = "bold"
        End If
    End If
    Header.style.width = 700
    lblAppTitle.style.width = 690
    lblDate.style.left = 265
    tabCriteria.style.width = 700
    cmdClose1.style.left = 630
    cmdClose2.style.left = 630
    cmdExport1.style.left = -1000
    cmdExport2.style.left = -1000
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
    Dim strReturnValue, strType
    
    strType = ""
    On Error Resume Next
    strType = document.all("hidRowInfo" & intRowID).value
    If Err.number <> 0 Then
        MsgBox "Report is still building.  Click OK."
        strType = "Error"
    Else
        strType = ""
    End If
    On Error Goto 0
    If strType <> "" Then Exit Sub

    strReturnValue = window.showModalDialog("PrintReview.asp?UserID=<%=gstrUserID%>&ReviewID=" & document.all("hidRowInfo" & intRowID).value, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
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
Dim strHoldWorker, strHoldReviewID, strHoldTab, strHoldScreen
Dim mstrHidden, intPrgCount, strValue, strValue2, strHoldSup
Dim maAuditRecords(1000,3)
Dim mstrAuditsPrinted
Dim blnFirstWorker

strColor = "#FFEFD5"

If adRs.EOF Then
    Response.Write "<BR><BR>"
    Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
    Response.Write " * No reviews matched the report criteria *"
End If

mintRowID = 0
strHoldWorker = ""
strHoldSup = ""
mstrHidden = ""
mstrAuditsPrinted = ""
Do While Not adRs.EOF
    If strHoldSup <> adRs.Fields("rvwSupervisorName").value Then
        If mintRowID > 0 Then
            Response.Write "<BR>"
        End If
        Call SingleLineTable(adRs.Fields("rvwSupervisorName").value,"GroupHeader^font-size:11pt;",0)
        blnFirstWorker = True
        strHoldSup = adRs.Fields("rvwSupervisorName").value
        strHoldReviewID = ""
    End If
    If strHoldWorker <> adRs.Fields("Worker").value Or blnFirstWorker = True Then
        If blnFirstWorker = False Then
            Response.Write "<BR>"
        Else
            blnFirstWorker = False
        End If
        Call SingleLineTable(adRs.Fields("Worker").value,"GroupHeader^font-size:10pt;",0)
        Call PrintHeaders()
        strHoldWorker = adRs.Fields("Worker").value
        strHoldReviewID = "" 
    End If
    If strHoldReviewID <> adRs.Fields("rvwID").value Or mstrShowDetail = "N" Then
        If strHoldReviewID <> "" And strHoldReviewID <> adRs.Fields("rvwID").value And mstrShowDetail = "Y" Then
            Call PrintAuditHeaders()
            Call PrintAuditRecords()
        End If
        If strHoldReviewID <> "" And mstrShowDetail = "Y" Then
            Call SingleLineTable("<HR>","GroupHeader",0)
        End If
        Call StartTable()
        Call StartTableBody()
        mintRowID = mintRowID + 1
        Call AddTableColumn(100,adRs.Fields("rvwDateEntered").value,False,"TableRow")
        Call AddTableColumn(90,adRs.Fields("rvwCaseNumber").value,True,"TableRow")
        Call AddTableColumn(90,adRs.Fields("rvwMonthYear").value,False,"TableRow")
        Call AddTableColumn(130,adRs.Fields("Reviewer").value,False,"TableRow")
        Call AddTableColumn(130,adRs.Fields("ReviewClass").value,False,"TableRow")
        Call AddTableColumn(50,adRs.Fields("rvwSupSig").value,False,"TableRow")
        Call AddTableColumn(50,adRs.Fields("rvwWrkSig").value,False,"TableRow")
        Call AddTableColumn(50,adRs.Fields("rvwSubmitted").value,False,"TableRow")
        Call EndTableBody()
        Call EndTable()
        strHoldReviewID = adRs.Fields("rvwID").value
        intPrgCount = 0
    End If
    
    If mstrShowDetail = "Y" Then
        strHoldTab = ""
        intPrgCount = intPrgCount + 1
        If intPrgCount > 1 Then
            Call SingleLineTable("&nbsp;", "GroupHeader", 20)
        End If
        Call SingleLineTable(adRs.Fields("ProgramName").value & " -- " & adRs.Fields("ReviewType").value, "GroupHeader", 10)
        adRsElms.Filter = "rveReviewID=" & adRs.Fields("rvwID").value & " And rveProgramID=" & adRs.Fields("ProgramID").value
        If adRsElms.RecordCount > 0 Then
            Do While Not adRsElms.Eof
                If strHoldScreen <> adRsElms.Fields("ElementName").value Then
                    Call SingleLineTable(adRsElms.Fields("ElementName").value, "GroupHeader", 20)
                    Call PrintDIHeaders()
                    strHoldScreen = adRsElms.Fields("ElementName").value
                End If
                adRsFacs.Filter = "rvfReviewID=" & adRs.Fields("rvwID").value & _
                                  " And rvfProgramID=" & adRs.Fields("ProgramID").value & _
                                  " And rvfElementID=" & adRsElms.Fields("rveElementID").value
                If adRsFacs.RecordCount > 0 Then
                    Do While Not adRsFacs.Eof
                        Call StartTable()
                        Call StartTableBody()
                        mintRowID = mintRowID + 1
                        Call AddTableColumn(20,"&nbsp;",False,"TableRow")
                        Call AddTableColumn(400,adRsFacs.Fields("FactorName").value,False,"TableRow")
                        Call AddTableColumn(50,adRsFacs.Fields("FactorStatus").value,False,"TableRow")
                        Call EndTableBody()
                        Call EndTable()
                        adRsFacs.MoveNext
                    Loop
                End If
                adRsComs.Filter = "rvcReviewID=" & adRs.Fields("rvwID").value & _
                                  " And rvcScreenName='" & adRsElms.Fields("ElementName").value & "'"
                If adRsComs.RecordCount > 0 Then
                    Do While Not adRsComs.Eof
                        'Call SingleLineTable(" And rvcScreenName='" & adRsElms.Fields("ElementName").value & "'", "GroupHeader", 30)
                        Call SingleLineTable("Comments", "GroupHeader", 20)
                        Call StartTable()
                        Call StartTableBody()
                        mintRowID = mintRowID + 1
                        Call AddTableColumn(30,"&nbsp;",False,"TableRow")
                        Call AddTableColumn(450,adRsComs.Fields("rvcComments").value,False,"TableRow")
                        Call EndTableBody()
                        Call EndTable()
                        adRsComs.MoveNext
                    Loop
                End If
                adRsElms.MoveNext
            Loop
        End If
        adRsAuds.Filter = "audTableRecordID=" & adRs.Fields("rvwID").value
        adRsAuds.Sort = "audDateOfAction"
        intI = 0
        If adRsAuds.RecordCount > 0 Then
            Call Clear_maAuditRecords()
            Do While Not adRsAuds.Eof
                For intJ = 1 To 100
                    strRecord = Parse(adRsAuds.Fields("audActionComments").value,"|",intJ)
                    If strRecord = "" Then Exit For
                    intI = intI + 1
                    maAuditRecords(intI,0) = FormatDateTime(adRsAuds.Fields("audDateOfAction").value,2) & " " & FormatDateTime(adRsAuds.Fields("audDateOfAction").value,3)
                    maAuditRecords(intI,1) = adRsAuds.Fields("audUserLogin").value
                    If InStr(strRecord,"^") > 0 Then
                        strValue = Parse(strRecord,"^",2)
                        If strValue = "" Then strValue="[BLANK]"
                        strValue2 = Parse(strRecord,"^",3)
                        If strValue2 = "" Then strValue2="[BLANK]"
                        maAuditRecords(intI,2) = Parse(strRecord,"^",1) & " CHANGED FROM: " & strValue & " TO: " & strValue2
                    Else
                        maAuditRecords(intI,2) = strRecord
                    End If
                Next
                adRsAuds.MoveNext
            Loop
        End If
    End If
    adRs.MoveNext
Loop
Response.Write mstrHidden
'Response.Write "<BR><BR>"

Sub Clear_maAuditRecords()
    Dim intI
    For intI = 1 To 1000
        maAuditRecords(intI,0) = ""
        maAuditRecords(intI,1) = ""
        maAuditRecords(intI,2) = ""
    Next
End Sub
Sub PrintAuditRecords()
    Dim intI
    
    For intI = 1 To 1000
        If maAuditRecords(intI,0) = "" Then Exit For
        Call StartTable()
        Call StartTableBody()
        mintRowID = mintRowID + 1
        Call AddTableColumn(20,"&nbsp;",False,"TableRow")
        Call AddTableColumn(145,Trim(maAuditRecords(intI,0)),False,"TableRow")
        Call AddTableColumn(80,maAuditRecords(intI,1),False,"TableRow")
        Call AddTableColumn(305,maAuditRecords(intI,2),False,"TableRow")
        
        Call EndTableBody()
        Call EndTable()
    Next
End Sub

Function GetTabName(intTabID)
    Select Case CInt(intTabID)
        Case 1
            GetTabName = "Action Integrity"
        Case 2
            GetTabName = "Data Integrity"
        Case 3
            GetTabName = "Information Gathering"
    End Select
End Function

Sub PrintAuditHeaders()
    Call SingleLineTable("&nbsp;", "GroupHeader", 20)
    Call SingleLineTable("Audit History", "GroupHeader", 20)
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(20,"&nbsp;", False, "TableRow")
    Call AddTableColumn(145,"Date",False,"RowHeader")
    Call AddTableColumn(80,"User ID",False,"RowHeader")
    Call AddTableColumn(305,"Description",False,"RowHeader")
    Call EndTableBody()
    Call EndTable()
End Sub

Sub PrintAIHeaders()
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(30,"&nbsp;", False, "TableRow")
    Call AddTableColumn(200,"Action",False,"RowHeader")
    Call AddTableColumn(250,"Status",False,"RowHeader")
    Call AddTableColumn(220,"Comments",False,"RowHeader")
    Call EndTableBody()
    Call EndTable()
End Sub

Sub PrintDIHeaders()
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(20,"&nbsp;", False, "TableRow")
    Call AddTableColumn(400,"Causal Factor",False,"RowHeader")
    Call AddTableColumn(50,"Status",False,"RowHeader")
    Call EndTableBody()
    Call EndTable()
End Sub

Sub PrintHeaders()
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(100,"Review Date",False,"RowHeader")
    Call AddTableColumn(90,"Case Number",False,"RowHeader")
    Call AddTableColumn(90,"Review Month",False,"RowHeader")
    Call AddTableColumn(130,"Reviewer",False,"RowHeader")
    Call AddTableColumn(130,"Review Class",False,"RowHeader")
    Call AddTableColumn(50,"Sup Sig",False,"RowHeader")
    Call AddTableColumn(50,"Wkr Sig",False,"RowHeader")
    Call AddTableColumn(50,"Sub Rpt",False,"RowHeader")
    Call EndTableBody()
    Call EndTable()
End Sub

Sub SingleLineTable(strText, strClassInfo, intBlankCellWidth)
    Call StartTable()
    Call StartTableBody()
    If intBlankCellWidth > 0 Then
        Call AddTableColumn(intBlankCellWidth,"&nbsp;",False,strClassInfo)
    End If
    Call AddTableColumn(700-intBlankCellWidth,strText,False,strClassInfo)
    Call EndTableBody()
    Call EndTable()
End Sub

Sub StartTable()
    Response.Write vbCrLf
    Response.Write "<TABLE Border=0 Rules=none Width=700 CellSpacing=0" & vbCrLf
    Response.Write "Style=""overflow: hidden; TOP:0"">" & vbCrLf
End Sub

Sub StartTableHeader()
    Response.Write "<THEAD id=tbhPrint style=""height:17"">" & vbCrLf
    Response.Write "<TR>"
End Sub

Sub AddTableColumn(intWidth, strText, blnDrillDown, strClassInfo)
    Dim strClass, strStyle
    
    strClass = Parse(strClassInfo,"^",1)
    strStyle = Parse(strClassInfo,"^",2)
    If strStyle = "" Then strStyle = ""

    If blnDrillDown = True Then
        mstrHidden = mstrHidden & "<INPUT type=hidden id=hidRowInfo" & mintRowID & " value=""" & adRs.Fields("rvwID").value & """>" & vbCrLf
        Response.Write "<TD id=lblCol0" & mintRowID & " class=" & strClass & " style=""" & strStyle & ";color:blue;cursor:hand;width:" & intWidth & ";padding-left:0;padding-right:0""" & vbCrLf
        Response.Write "onmouseover=""Call ColMouseEvent(0,0," & mintRowID & ")"" onmouseout=""Call ColMouseEvent(1,0," & mintRowID & ")"" onclick=""Call ColClickEvent(0," & mintRowID & ")"">" & strText & "</TD>" & vbCrLf
    Else
        Response.Write "<TD class=" & strClass & " style=""" & strStyle & ";width:" & intWidth & ";padding-left:0;padding-right:0"">" & strText & "</TD>"
    End If
End Sub
Sub EndTableHeader()
    Response.Write "</TR></THEAD>"
End Sub
Sub StartTableBody()
    Response.Write "<TBODY><TR>"
End Sub
Sub EndTableBody()
    Response.Write "</TR></TBODY>"
End Sub
Sub EndTable()
    Response.Write "</TABLE>"
End Sub

%>
<!--#include file="IncRptFooter.asp"-->
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->