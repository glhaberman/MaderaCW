<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
Dim adRs, adCmd
Dim adRsElms, adRsFacs, adRsComs, adRsAuds, adRsPrgs
Dim mlngReviewID
Dim mstrUserID, mstrPageTitle, mstrHTML
Dim intI, intJ, mintRowID, strRecord
Dim strHoldWorker, strHoldReviewID, strHoldTab, strHoldScreen
Dim mstrHidden, intPrgCount, strValue, strValue2, strColor
Dim maAuditRecords(1000,3)
Dim strER

mstrUserID = Request.QueryString("UserID")
If Len(mstrUserID) = 0 Then mstrUserID = "unknown"
mlngReviewID = Request.QueryString("ReviewID")
If Len(mlngReviewID) = 0 Then mlngReviewID = 0
%>
<!--#include file="IncCnn.asp"-->
<%
mstrPageTitle = "Printed Review"
'Retrieve the records that match the report criteria:
Set adCmd = GetAdoCmd("spReviewGet4Print")
    AddParmIn adCmd, "@ReviewID", adInteger, 0, mlngReviewID
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)    

Set adRsElms = adRs.NextRecordSet
Set adRsFacs = adRs.NextRecordSet
Set adRsComs = adRs.NextRecordSet
Set adRsAuds = adRs.NextRecordSet
Set adRsPrgs = adRs.NextRecordSet
%>

<HTML>
<HEAD>
    <TITLE>Printed Review</TITLE>
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
    .FunctionHeader
        {
        height: 15;
        padding-left: 5px;
        padding-right: 5px;
        border-style: none;
        font-family: Tahoma;
        font-size:12pt;
        font-weight: bold;
        text-align: left;
        color:black;
        background-color: white;
        overflow: hidden
        }
    .GroupHeader
        {
        height: 15;
        padding-left: 5px;
        padding-right: 5px;
        border-style: none;
        font-family: Tahoma;
        font-size:10pt;
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
    .TableRowBold
        {
        font-family: Tahoma;
        font-size:10pt;
        background-color: white;
        color: black;
        overflow: visible;
        font-weight:bold;
        }
    .TableRowBoldRight
        {
        font-family: Tahoma;
        font-size:10pt;
        text-align:right;
        background-color: white;
        color: black;
        overflow: visible;
        font-weight:bold;
        }
    </STYLE>
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>

<!--
Sub window_onload
	Call FormShow("none")
	PageBody.style.cursor = "wait"
	Call SizeAndCenterWindow(767, 520, True)
    Call FormShow("")
    PageBody.style.cursor = "default"
    cmdPrint.focus
End Sub

Sub cmdClose_onclick()	
	Window.close
End Sub

Sub FormShow(strVis)
	cmdPrint.style.display = strVis
    cmdClose.style.display = strVis
    cmdExport.style.display = strVis
    PageFrame.style.display = strVis
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub window_onbeforeprint()
    cmdPrint.style.visibility = "hidden"
    cmdClose.style.visibility = "hidden"
    cmdExport.style.visibility = "hidden"
    divSignature.style.display = "inline"
End Sub

Sub window_onafterprint()
    cmdPrint.style.visibility = "visible"
    cmdClose.style.visibility = "visible"
    cmdExport.style.visibility = "visible"
    divSignature.style.display = "none"
End Sub
Sub cmdExport_onclick()
    Dim CtlRng
    'If the results div is not empty, copy it's contents to the clipboard:
    If PageFrame.children.length > 0 Then
        'A controlRange object is used to select the results div, then copy it:
        Set CtlRng = PageBody.createControlRange()
        CtlRng.AddElement(PageFrame)
        CtlRng.Select
        CtlRng.execCommand("Copy")
        Set CtlRng = Nothing
        'Clear the selection:
        document.selection.empty
        MsgBox "Results copied to clipboard.", ,"Copy Results"
    End If
End Sub

-->
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<BODY id=Pagebody style="background-color:white" bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style=overflow:auto>
    <BUTTON id=cmdPrint title="Send to the printer" style="LEFT:20; WIDTH:65; TOP:10; HEIGHT:20" tabIndex=1>
        Print
    </BUTTON>
    
     <BUTTON id=cmdExport title="Copy review form to clipboard" style="LEFT:95; WIDTH:65; TOP:10; HEIGHT:20" onclick="cmdExport_onclick" tabIndex=55>
        Copy
    </BUTTON>

    <BUTTON id=cmdClose title="Close window and return to case review" style="LEFT:595; WIDTH:65; TOP:10; HEIGHT:20" tabIndex=55>
        Close
    </BUTTON>
<%
Response.Write "<DIV id=PageFrame style=""LEFT:10;visibility:visible; WIDTH:700; FONT-SIZE:12pt; HEIGHT:225"">"
Response.Write "<BR><BR>"
Response.Write "<TABLE id=PageHeader cellspacing=0 border=0 width=700 style=""TABLE-LAYOUT:auto"">"
Response.Write "<THEAD>"
Response.Write "<TR valign=top>"
Response.Write "<TD width=325 style=""COLOR:#C0C0C0; FONT-SIZE:8pt; PADDING-LEFT:8; TEXT-ALIGN:Left; FONT-WEIGHT:bold; BORDER-TOP:1 solid #C0C0C0"">"
Response.Write gstrTitle & " " & gstrAppName
Response.Write "</TD>"
Response.Write "<TD width=325 style=""COLOR:#C0C0C0; FONT-SIZE:8pt; PADDING-RIGHT:8; TEXT-ALIGN:right; FONT-WEIGHT:bold; BORDER-TOP:1 solid #C0C0C0"">"
Response.Write "Date Printed: " & FormatDateTime(Now(),vbGeneralDate) & "(" & mstrUserID & ")"
Response.Write "</TD></TR></THEAD></TABLE>"
Response.Write "<TABLE id=PageTitle" & intI & " cellspacing=0 border=0 width=700 style=""TABLE-LAYOUT:auto"">"
Response.Write "<THEAD>"
Response.Write "<TR valign=middle>"
Response.Write "<TD style=""FONT-SIZE:14pt; HEIGHT:30; TEXT-ALIGN:center; BORDER-BOTTOM:1 solid black"">"
Response.Write "<B>Print Case Review</B>"
Response.Write "</TD></TR></THEAD></TABLE>"
Response.Write ""

strColor = "#FFEFD5"

mintRowID = 0
strHoldWorker = ""
mstrHidden = ""
Call StartTable()
Call StartTableBody()
Call AddTableColumn(350,"<B>Review ID:</B> " & adRs.Fields("rvwID").value,False,"TableRow")
Call AddTableColumn(350,"",False,"TableRow")
Call EndTableBody()
Call StartTableBody()
Call AddTableColumn(350,"<B>Case Number:</B> " & adRs.Fields("rvwCaseNumber").value,False,"TableRow")
Call AddTableColumn(350,"<B>Case Name:</B> " & adRs.Fields("CaseName").value,False,"TableRow")
Call EndTableBody()
Call StartTableBody()
Call AddTableColumn(350,"<B>Review Date:</B> " & adRs.Fields("rvwDateEntered").value,False,"TableRow")
Call AddTableColumn(350,"<B>Review Month:</B> " & adRs.Fields("rvwMonthYear").value,False,"TableRow")
Call EndTableBody()
Call StartTableBody()
Call AddTableColumn(350,"<B>Reviewer:</B> " & adRs.Fields("Reviewer").value,False,"TableRow")
Call AddTableColumn(350,"<B>Review Class:</B> " & adRs.Fields("ReviewClass").value,False,"TableRow")
Call EndTableBody()
Call StartTableBody()
Call AddTableColumn(350,"<B>Supervisor:</B> " & adRs.Fields("Supervisor").value,False,"TableRow")
Call AddTableColumn(350,"<B>Supervisor Signature:</B> " & adRs.Fields("rvwSupSig").value,False,"TableRow")
Call EndTableBody()
Call StartTableBody()
Call AddTableColumn(350,"<B>Worker Response Requirement:</B> " & adRs.Fields("WorkerResponseRequirement").value,False,"TableRow")
Call AddTableColumn(350,"<B>Worker Response Due Date:</B> " & adRs.Fields("ResponseDueDate").value,False,"TableRow")
Call EndTableBody()
Call StartTableBody()
Call AddTableColumn(350,"<B>Worker:</B> " & adRs.Fields("Worker").value,False,"TableRow")
Call AddTableColumn(350,"<B>Worker Signature:</B> " & adRs.Fields("rvwWrkSig").value,False,"TableRow")
Call EndTableBody()
Call StartTableBody()
Call AddTableColumn(350,"<B>Worker Response:</B> " & adRs.Fields("WorkerResponse").value,False,"TableRow")
Call AddTableColumn(350,"<B>Submitted To Reports:</B> " & adRs.Fields("rvwSubmitted").value,False,"TableRow")
Call EndTableBody()
Response.Write "<BR>"

strHoldReviewID = adRs.Fields("rvwID").value
intPrgCount = 0

strER = ""

Do While Not adRsPrgs.EOF
    strHoldTab = ""
    intPrgCount = intPrgCount + 1
    If intPrgCount > 1 Then
        Call SingleLineTable("&nbsp;", "GroupHeader", 20)
    End If
    If adRsPrgs.Fields("ProgramID").value >= 50 And strER = "" Then
        Call SingleLineTable("Enforcement Remedies", "FunctionHeader", 5)
        strER = "Y"
    End If
    If adRsPrgs.Fields("ProgramID").value >= 50 Or adRsPrgs.Fields("ProgramID").value = 6 Then
        Call SingleLineTable(adRsPrgs.Fields("FunctionName").value, "FunctionHeader", 10)
    Else
        Call SingleLineTable(adRsPrgs.Fields("FunctionName").value & " -- " & adRsPrgs.Fields("ReviewType").value, "FunctionHeader", 10)
    End If
    adRsElms.Filter = "rveReviewID=" & mlngReviewID & " And rveProgramID=" & adRsPrgs.Fields("ProgramID").value
    If adRsElms.RecordCount > 0 Then
        Do While Not adRsElms.Eof
            If adRsElms.Fields("rveTypeID").value <> strHoldTab Then
                If strHoldTab <> "" Then
                    Call SingleLineTable("&nbsp;", "GroupHeader", 20)
                End If
                strHoldTab = adRsElms.Fields("rveTypeID").value
                Call SingleLineTable("Elements", "GroupHeader", 20)
                Select Case CInt(adRsElms.Fields("rveTypeID").value)
                    Case 1, 3
                        Call PrintAIHeaders()
                    Case 2
                        strHoldScreen = ""
                End Select
            End If
            Select Case CInt(adRsElms.Fields("rveTypeID").value)
                Case 1,3
                    Call StartTable()
                    Call StartTableBody()
                    mintRowID = mintRowID + 1
                    Call AddTableColumn(30,"&nbsp;",False,"TableRow")
                    Call AddTableColumn(200,adRsElms.Fields("ElementName").value,False,"TableRow")
                    Call AddTableColumn(250,adRsElms.Fields("ElementStatus").value,False,"TableRow")
                    Call AddTableColumn(220,adRsElms.Fields("rveComments").value,False,"TableRow")
                    Call EndTableBody()
                    Call EndTable()
                Case 2
                    If strHoldScreen <> adRsElms.Fields("ElementName").value Then
                        Call SingleLineTable(adRsElms.Fields("ElementName").value, "GroupHeader", 30)
                        Call PrintDIHeaders()
                        strHoldScreen = adRsElms.Fields("ElementName").value
                    End If
                    adRsFacs.Filter = "rvfReviewID=" & mlngReviewID & _
                                      " And rvfProgramID=" & adRsPrgs.Fields("ProgramID").value & _
                                      " And rvfElementID=" & adRsElms.Fields("rveElementID").value
                    If adRsFacs.RecordCount > 0 Then
                        Do While Not adRsFacs.Eof
                            Call StartTable()
                            Call StartTableBody()
                            mintRowID = mintRowID + 1
                            Call AddTableColumn(30,"&nbsp;",False,"TableRow")
                            Call AddTableColumn(400,adRsFacs.Fields("FactorName").value,False,"TableRow")
                            Call AddTableColumn(50,adRsFacs.Fields("FactorStatus").value,False,"TableRow")
                            Call EndTableBody()
                            Call EndTable()
                            adRsFacs.MoveNext
                        Loop
                    End If
                    adRsComs.Filter = "rvcReviewID=" & mlngReviewID & _
                                      " And rvcScreenName='" & adRsElms.Fields("ElementName").value & "'"
                    If adRsComs.RecordCount > 0 Then
                        Do While Not adRsComs.Eof
                            'Call SingleLineTable(" And rvcScreenName='" & adRsElms.Fields("ElementName").value & "'", "GroupHeader", 30)
                            Call SingleLineTable("Comments", "GroupHeader", 30)
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
                    
            End Select
            adRsElms.MoveNext
        Loop
    End If
    adRsPrgs.MoveNext
Loop
adRsAuds.Sort = "audDateOfAction"
intI = 0
If adRsAuds.RecordCount > 0 Then
    Call PrintAuditHeaders()
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
    Call PrintAuditRecords()
End If
Response.Write mstrHidden

Response.Write "<DIV id=divSignature style=""LEFT:10;display:none; WIDTH:700; FONT-SIZE:12pt; HEIGHT:225""><BR><BR>"
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(10,"&nbsp;", False, "TableRow")
    Call AddTableColumn(300,"Worker Agrees with Review &nbsp;[ ]",False,"TableRowBold")
    Call AddTableColumn(40,"&nbsp;",False,"TableRowBold")
    Call AddTableColumn(300,"Worker Disagrees with Review &nbsp;[ ]",False,"TableRowBoldRight")
    Call EndTableBody()
    Call EndTable()
    Response.Write "<BR>"
    Call SingleLineTable("Comments / Disagree:", "TableRowBold", 10)
    Response.Write "<BR><BR><BR><BR>"
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(10,"&nbsp;", False, "TableRow")
    Call AddTableColumn(450,"Worker Signature: ______________________________________",False,"TableRowBold")
    Call AddTableColumn(5,"&nbsp;", False, "TableRow")
    Call AddTableColumn(185,"Date: ________________",False,"TableRowBold")
    Response.Write "</TR><TR><TD>&nbsp;</TD></TR><TR>"
    Call AddTableColumn(10,"&nbsp;", False, "TableRow")
    Call AddTableColumn(450,"Supervisor Initials: ____________________________________",False,"TableRowBold")
    Call AddTableColumn(5,"&nbsp;", False, "TableRow")
    Call AddTableColumn(185,"Date: ________________",False,"TableRowBold")
    Call EndTableBody()
    Call EndTable()
Response.Write "</DIV>"
'Response.Write mstrHTML

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
        Call AddTableColumn(10,"&nbsp;",False,"TableRow")
        Call AddTableColumn(150,Trim(maAuditRecords(intI,0)),False,"TableRow")
        Call AddTableColumn(80,maAuditRecords(intI,1),False,"TableRow")
        Call AddTableColumn(310,maAuditRecords(intI,2),False,"TableRow")
        
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
    Call SingleLineTable("&nbsp;", "GroupHeader", 10)
    Call SingleLineTable("Audit History", "FunctionHeader", 10)
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(10,"&nbsp;", False, "TableRow")
    Call AddTableColumn(150,"Date",False,"RowHeader")
    Call AddTableColumn(80,"User ID",False,"RowHeader")
    Call AddTableColumn(310,"Description",False,"RowHeader")
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
    Call AddTableColumn(30,"&nbsp;", False, "TableRow")
    Call AddTableColumn(400,"Causal Factor",False,"RowHeader")
    Call AddTableColumn(50,"Status",False,"RowHeader")
    Call EndTableBody()
    Call EndTable()
End Sub

Sub PrintHeaders()
    Call StartTable()
    Call StartTableBody()
    'Call AddTableColumn(100,"Review Date",False,"RowHeader")
    Call AddTableColumn(90,"Case Number",False,"RowHeader")
    Call AddTableColumn(190,"Reviewer",False,"RowHeader")
    Call AddTableColumn(170,"Review Class",False,"RowHeader")
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
</HTML>
