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
Dim adRs, adRsFull, adRsFactors
Dim adCmd
Dim intI, intJ
Dim intShadeCount
Dim strColor, strRecord
Dim mintRowID, mstrHidden

%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<!--#include file="IncDrillDownSvr.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
'Retrieve the records that match the report criteria:
Set adCmd = GetAdoCmd("spRptArcEmployeePerformance")
    AddParmIn adCmd, "@AliasID", adInteger, 0, glngAliasPosID
    AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
    AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
    AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
    AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, ReqIsDate("StartDate")
    AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
    AddParmIn adCmd, "@WorkerName", adVarchar, 50, ReqIsBlank("Worker")
    AddParmIn adCmd, "@StartReviewMonth", adDBTimeStamp, 0, ReqIsDate("StartReviewMonth")
    AddParmIn adCmd, "@EndReviewMonth", adDBTimeStamp, 0, ReqIsDate("EndReviewMonth")
    AddParmIn adCmd, "@DrillDownID", adInteger, 0, Null
    AddParmIn adCmd, "@DrillDownKeyID", adInteger, 0, Null
    AddParmIn adCmd, "@ProgramID", adInteger, 0, Null
    AddParmIn adCmd, "@TabID", adInteger, 0, Null
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)    
Set adRsFull = adRs.NextRecordSet()
Set adRsFactors = adRs.NextRecordSet()
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
        font-size:10pt;
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
        padding-left: 5px;
        padding-right: 5px;
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
    lblAppTitle.innerText = "Employee Performance: All Functions"
    lblAppTitle.style.fontweight = "bold"
    Header.style.width = 700
    lblAppTitle.style.width = 690
    lblDate.style.left = 265
    tabCriteria.style.width = 700
    cmdClose1.style.left = 630
    cmdClose2.style.left = 630
    cmdExport1.style.left = -1100
    cmdExport2.style.left = -1100
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
    Dim strType
    
    strType = ""
    On Error Resume Next
    strType = document.all("hidRowInfo" & intColID & intRowID).value
    If Err.number <> 0 Then
        MsgBox "Report is still building.  Click OK."
        strType = "Error"
    Else
        strType = ""
    End If
    On Error Goto 0
    If strType <> "" Then Exit Sub

    strSelectedName = "<%=ReqIsBlank("Worker")%>"
    strSelectedName = strSelectedName & "&PID=" & Parse(document.all("hidRowInfo" & intColID & intRowID).value,"^",1)
    strSelectedName = strSelectedName & "&TID=" & Parse(document.all("hidRowInfo" & intColID & intRowID).value,"^",2)
    strType = Parse(document.all("hidRowInfo" & intColID & intRowID).value,"^",4)
    Select Case strType
        Case "PFT" 'Factor Totals for a Program
            'Select Case Parse(document.all("hidRowInfo" & intColID & intRowID).value,"^",2)
            '    Case 1
            '        strSelectedName = strSelectedName & "&SN2="
            '    Case 2
            '    Case 3
            'End Select
            strSelectedName = strSelectedName & "&CIDAdd=40"
            intDD2 = 0
        Case "EFT" 'Factor Totals for an Element
            'intColID = CInt(intColID) + 20
            strSelectedName = strSelectedName & "&CIDAdd=20"
            intDD2 = Parse(document.all("hidRowInfo" & intColID & intRowID).value,"^",3)
        Case "PET" 'Element Totals for a Program
            'intColID = CInt(intColID) + 30
            strSelectedName = strSelectedName & "&CIDAdd=30"
            intDD2 = 0
        Case ""    'Element line
            intDD2 = Parse(document.all("hidRowInfo" & intColID & intRowID).value,"^",3)
        Case Else  'Factor line
            'intColID = CInt(intColID) + 10
            strSelectedName = strSelectedName & "&CIDAdd=10"
            intDD2 = Parse(document.all("hidRowInfo" & intColID & intRowID).value,"^",4)
    End Select
    Call DrillDownColClickEventNoStaff("spRptEmployeePerformance", intColID, intRowID, intDD2, strSelectedName)
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
Dim oReport, oElement, dctElements, oFactor, dctFactors
Dim intHoldFunctionID, intHoldTabID, intTotalElms, dblPercent, intHoldElementID
Dim aTotals(2,3), strHoldFunctionName
Dim blnFirstFactor

strColor = "#FFEFD5"

If adRs.EOF Then
    Response.Write "<BR><BR>"
    Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
    Response.Write " * No reviews matched the report criteria *"
End If

Set oReport = New clsReport

Set dctElements = oReport.Elements
Set dctFactors = oReport.Factors
Call SingleLineTable(ReqIsBlank("Worker"), "MangementHeading^font-weight:bold;font-size:14;", 0)
intHoldTabID = 0
For Each oElement In dctElements
    adRsFull.Filter = "prgID=" & Parse(oElement,"^",1) & " AND elmTypeID=" & Parse(oElement,"^",2) & " AND elmShortName='" & Parse(oElement,"^",3) & "'"
    adRsFull.MoveFirst
    If CInt(Parse(oElement,"^",2)) <> CInt(intHoldTabID) Then
        If CInt(intHoldTabID) > 0 Then
            If CLng(aTotals(2,1)) + CLng(aTotals(2,2)) + CLng(aTotals(2,3)) > 0 Then
                'Call SingleLineTable("&nbsp;", "", 0)
                Call WriteTotals("ElementTotal", strHoldFunctionName, intHoldTabID, aTotals(2,1), aTotals(2,2), aTotals(2,3), intHoldFunctionID & "^" & intHoldTabID & "^^PFT") ' & "^" & intHoldElementID)
            End If
            Call WriteTotals("Function", strHoldFunctionName, intHoldTabID, aTotals(0,1), aTotals(0,2), aTotals(0,3), intHoldFunctionID & "^" & intHoldTabID & "^^PET")
            Call SingleLineTable("&nbsp;", "", 0)
        End If
        Call SingleLineTable(GetTabName(Parse(oElement,"^",2),1), "MangementHeading^font-weight:bold;", 0)
        intHoldTabID = Parse(oElement,"^",2)
        intHoldFunctionID = 0
        intHoldElementID = -1
    End If
    If CInt(Parse(oElement,"^",1)) <> CInt(intHoldFunctionID) Then
        If CInt(intHoldFunctionID) > 0 Then
            If CLng(aTotals(2,1)) + CLng(aTotals(2,2)) + CLng(aTotals(2,3)) > 0 Then
                'Call SingleLineTable("&nbsp;", "", 0)
                Call WriteTotals("ElementTotal", strHoldFunctionName, intHoldTabID, aTotals(2,1), aTotals(2,2), aTotals(2,3), intHoldFunctionID & "^" & intHoldTabID & "^^PFT") ' & "^" & intHoldElementID)
            End If
            Call WriteTotals("Function", strHoldFunctionName, intHoldTabID, aTotals(0,1), aTotals(0,2), aTotals(0,3), intHoldFunctionID & "^" & intHoldTabID & "^^PET")
            Call SingleLineTable("&nbsp;", "", 0)
        End If
        Call SingleLineTable(adRsFull.Fields("prgShortTitle").value, "MangementHeading^font-weight:bold;", 10)
        intHoldFunctionID = Parse(oElement,"^",1)
        strHoldFunctionName = adRsFull.Fields("prgShortTitle").value
        Call PrintHeaders(intHoldTabID)
        For intI = 1 To 3
            aTotals(0,intI) = 0
            aTotals(2,intI) = 0
        Next
        intHoldElementID = -2
    End If
    For intI = 1 To 3
        aTotals(1,intI) = 0
    Next
    'oReport.WriteLine oElement & " -- " & dctElements(oElement)
    intTotalElms = CInt(Parse(dctElements(oElement),"^",1)) + CInt(Parse(dctElements(oElement),"^",2))
    intHoldElementID = Parse(oElement,"^",3)
    Call StartTable()
    Call StartTableBody()
    mintRowID = mintRowID + 1
    Call AddTableColumn(10,"&nbsp;",False,"TableRow")
    'Call DebugWrite("<BR>oElement:<BR>" & oElement,False)
    If intHoldTabID = 2 Then
        Call AddTableColumn(320,adRsFull.Fields("elmShortName").value,False,"TableRow")
        Call AddTableColumn(90,Parse(dctElements(oElement),"^",3),False,"TableRow^text-align:center;")
    Else
        Call AddTableColumn(410,adRsFull.Fields("elmShortName").value,False,"TableRow")
    End If
    Call AddTableColumn(90,Parse(dctElements(oElement),"^",1),oElement,"TableRow^text-align:center;^1")
    
    If intTotalElms > 0 Then
        dblPercent = (CInt(Parse(dctElements(oElement),"^",1)) / CInt(intTotalElms)) * 100
    Else
        dblPercent = 0
    End If
    Call AddTableColumn(90,FormatNumber(dblPercent,1,True,True,True) & "%",False,"TableRow^text-align:center;")
    Call AddTableColumn(90,Parse(dctElements(oElement),"^",2),oElement,"TableRow^text-align:center;^2")
    Call EndTableBody()
    Call EndTable()
    For intI = 1 To 3
        aTotals(0,intI) = CLng(aTotals(0,intI)) + CLng(Parse(dctElements(oElement),"^",intI))
    Next
    blnFirstFactor = True
    For Each oFactor In dctFactors
        If Parse(oFactor,"^",1) & "^" & Parse(oFactor,"^",2) & "^" & Parse(oFactor,"^",3) = oElement Then
            If blnFirstFactor Then
                Call PrintFactorHeaders(intHoldTabID)
                blnFirstFactor = False
            End If
            Call StartTable()
            Call StartTableBody()
            mintRowID = mintRowID + 1
            Call AddTableColumn(10,"&nbsp;",False,"TableRow")
            If intHoldTabID = 2 Then
                Call AddTableColumn(320,Parse(oFactor,"^",4),False,"TableRow")
                Call AddTableColumn(90,Parse(dctFactors(oFactor),"^",3),False,"TableRow^text-align:center;^")
            Else
                Call AddTableColumn(410,Parse(oFactor,"^",4),False,"TableRow")
            End If
            Call AddTableColumn(90,Parse(dctFactors(oFactor),"^",1),oFactor,"TableRow^text-align:center;^1")
            
            If intTotalElms > 0 Then
                dblPercent = (CInt(Parse(dctFactors(oFactor),"^",1)) / CInt(intTotalElms)) * 100
            Else
                dblPercent = 0
            End If
            Call AddTableColumn(90,FormatNumber(dblPercent,1,True,True,True) & "%",False,"TableRow^text-align:center;")
            Call AddTableColumn(90,Parse(dctFactors(oFactor),"^",2),oFactor,"TableRow^text-align:center;^2")
            Call EndTableBody()
            Call EndTable()
            For intI = 1 To 3
                aTotals(1,intI) = CLng(aTotals(1,intI)) + CLng(Parse(dctFactors(oFactor),"^",intI))
                aTotals(2,intI) = CLng(aTotals(2,intI)) + CLng(Parse(dctFactors(oFactor),"^",intI))
            Next
        End If
    Next
    If blnFirstFactor = False Then
        Call WriteTotals("Element", adRsFull.Fields("elmShortName").value, intHoldTabID, aTotals(1,1), aTotals(1,2), aTotals(1,3), intHoldFunctionID & "^" & intHoldTabID & "^" & intHoldElementID & "^EFT")
        Call SingleLineTable("&nbsp;", "", 0)
    End If
Next
If CInt(intHoldFunctionID) > 0 Then
    If CLng(aTotals(2,1)) + CLng(aTotals(2,2)) + CLng(aTotals(2,3)) > 0 Then
        'Call SingleLineTable("&nbsp;", "", 0)
        Call WriteTotals("ElementTotal", strHoldFunctionName, intHoldTabID, aTotals(2,1), aTotals(2,2), aTotals(2,3), intHoldFunctionID & "^" & intHoldTabID & "^^PFT")
    End If
    Call WriteTotals("Function", strHoldFunctionName, intHoldTabID, aTotals(0,1), aTotals(0,2), aTotals(0,3), intHoldFunctionID & "^" & intHoldTabID & "^^PET")
    Call SingleLineTable("&nbsp;", "", 0)
End If

Response.Write "<BR>"
Response.Write mstrHidden

Function DebugWrite(strText, blnEnd)
    Response.Write strText
    If blnEnd Then
        Response.End
    Else
        Response.Flush
    End If
End Function

Function GetTabName(intTabID, intTypeID)
    Select Case CInt(intTabID)
        Case 1
            Select Case intTypeID
                Case 1
                    GetTabName = "Action Integrity"
                Case 2
                    GetTabName = "Action"
                Case 3
                    GetTabName = "Decision"
                Case 4
                    GetTabName = "Correct"
                Case 5
                    GetTabName = "Incorrect"
            End Select
        Case 2
            Select Case intTypeID
                Case 1
                    GetTabName = "Data Integrity"
                Case 2
                    GetTabName = "Screen"
                Case 3
                    GetTabName = "Field"
                Case 4
                    GetTabName = "Yes"
                Case 5
                    GetTabName = "No"
            End Select
        Case 3
            Select Case intTypeID
                Case 1
                    GetTabName = "Information Gathering"
                Case 2
                    GetTabName = "Question"
                Case 3
                    GetTabName = "Answer"
                Case 4
                    GetTabName = "Yes"
                Case 5
                    GetTabName = "No"
            End Select
    End Select
End Function

Sub PrintFactorHeaders(intTabID)
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(10,"",False,"")
    If intTabID = 2 Then
        Call AddTableColumn(320,GetTabName(intTabID,3),False,"RowHeader")
        Call AddTableColumn(90,"NA",False,"RowHeader^text-align:center;")
    Else
        Call AddTableColumn(410,GetTabName(intTabID,3),False,"RowHeader")
    End If
    Call AddTableColumn(90,GetTabName(intTabID,4),False,"RowHeader^text-align:center;")
    Call AddTableColumn(90,"% " & GetTabName(intTabID,4),False,"RowHeader^text-align:center;")
    Call AddTableColumn(90,GetTabName(intTabID,5),False,"RowHeader^text-align:center;")
    Call EndTableBody()
    Call EndTable()
End Sub

Sub WriteTotals(strType, strTotalName, intTabID, intTotal1, intTotal2, intTotal3, strDDKey)
    Dim intTotal, dblPercent
    Dim intCallID
    
    If strType = "Function" Then
        intCallID = 2
    Else
        intCallID = 3
    End If
    'If strType = "ElementTotal" Then strTotalName = strTotalName & "
    intTotal = intTotal1 + intTotal2
    mintRowID = mintRowID + 1
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(10,"",False,"")
    If intTabID = 2 Then
        Call AddTableColumn(320,strTotalName & " " & GetTabName(intTabID,intCallID) & " Total:",False,"RowHeader^background-color:lightgray")
        Call AddTableColumn(90,intTotal3,False,"RowHeader^text-align:center;background-color:lightgray")
    Else
        Call AddTableColumn(410,strTotalName & " " & GetTabName(intTabID,intCallID) & " Total:",False,"RowHeader^background-color:lightgray")
    End If
    Call AddTableColumn(90,intTotal1,strDDKey,"RowHeader^text-align:center;background-color:lightgray^1")
    If intTotal > 0 Then
        dblPercent = (CInt(intTotal1) / CInt(intTotal)) * 100
    Else
        dblPercent = 0
    End If
    Call AddTableColumn(90,FormatNumber(dblPercent,1,True,True,True) & "%",False,"RowHeader^text-align:center;background-color:lightgray")
    Call AddTableColumn(90,intTotal2,strDDKey,"RowHeader^text-align:center;background-color:lightgray^2")
    Call EndTableBody()
    Call EndTable()
End Sub

Sub PrintHeaders(intTabID)
    Call StartTable()
    Call StartTableBody()
    Call AddTableColumn(10,"",False,"")
    If intTabID = 2 Then
        Call AddTableColumn(320,GetTabName(intTabID,2),False,"RowHeader")
        Call AddTableColumn(90,"NA",False,"RowHeader^text-align:center;")
    Else
        Call AddTableColumn(410,GetTabName(intTabID,2),False,"RowHeader")
    End If
    Call AddTableColumn(90,GetTabName(intTabID,4),False,"RowHeader^text-align:center;")
    Call AddTableColumn(90,"% " & GetTabName(intTabID,4),False,"RowHeader^text-align:center;")
    Call AddTableColumn(90,GetTabName(intTabID,5),False,"RowHeader^text-align:center;")
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

Sub AddTableColumn(intWidth, strText, strDrillDownKey, strClassInfo)
    Dim strClass, strStyle, intColID
    Dim blnAllowDrillDown
    
    blnAllowDrillDown = False 'For now, disable drilldown on Archive report
    
    strClass = Parse(strClassInfo,"^",1)
    strStyle = Parse(strClassInfo,"^",2)
    intColID = Parse(strClassInfo,"^",3)
    If strStyle = "" Then strStyle = ""
    If intColID = "" Then intColID = "0"

    If strDrillDownKey = "False" Then
        Response.Write "<TD class=" & strClass & " style=""" & strStyle & ";width:" & intWidth & ";padding-left:0;padding-right:0"">" & strText & "</TD>"
    Else
        If CDbl(strText) > 0 And blnAllowDrillDown = True Then
            mstrHidden = mstrHidden & "<INPUT type=hidden id=hidRowInfo" & intColID & mintRowID & " value=""" & strDrillDownKey & """>" & vbCrLf
            Response.Write "<TD id=lblCol" & intColID & mintRowID & " class=" & strClass & " style=""" & strStyle & ";color:blue;cursor:hand;width:" & intWidth & ";padding-left:0;padding-right:0""" & vbCrLf
            Response.Write "onmouseover=""Call ColMouseEvent(0," & intColID & "," & mintRowID & ")"" onmouseout=""Call ColMouseEvent(1," & intColID & "," & mintRowID & ")"" onclick=""Call ColClickEvent(" & intColID & "," & mintRowID & ")"">" & strText & "</TD>" & vbCrLf
        Else
            Response.Write "<TD class=" & strClass & " style=""" & strStyle & ";width:" & intWidth & ";padding-left:0;padding-right:0"">" & strText & "</TD>"
        End If
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

Class clsReport
    Private dctTabs
    Private dctElements
    Private dctFactors
    
    Private Sub Class_Initialize()
        Set dctTabs = CreateObject("Scripting.Dictionary")
        Set dctElements = CreateObject("Scripting.Dictionary")
        Set dctFactors = CreateObject("Scripting.Dictionary")
        
        Call ProcessRecordset()
    End Sub
    
    Public Property Get Elements()
        Set Elements = dctElements
    End Property
    
    Public Property Get Factors()
        Set Factors = dctFactors
    End Property
    
    Public Sub WriteLine(strText)
        Response.Write strText & "<BR>"
    End Sub
    
    Private Sub AddFactor(strKey, strFactorStatus)
        Dim intI, strRecord, strNewItem
        Dim aCols(3)
        
        'Dictionary object stores the Yes, No and NA in a string
        If dctFactors.Exists(strKey) Then
            strRecord = dctFactors(strKey)
            For intI = 1 To 3
                aCols(intI) = Parse(strRecord,"^",intI)
            Next
        Else
            dctFactors.Add strKey, ""
            For intI = 1 To 3
                aCols(intI) = "0"
            Next
        End If
        strNewItem = ""
        For intI = 1 To 3
            If CInt(intI) = CInt(GetStatusID(strFactorStatus)) - 21 Then
                aCols(intI) =  CInt(aCols(intI)) + 1
            End If
            strNewItem = strNewItem & aCols(intI) & "^"
        Next
        dctFactors(strKey) = strNewItem
    End Sub
    
    Private Sub AddElement(strKey, intElementStatus)
        Dim intI, strRecord, strNewItem
        Dim aCols(3)
        
        'Dictionary object stores the Correct and Not Correct in a string
        If dctElements.Exists(strKey) Then
            strRecord = dctElements(strKey)
            For intI = 1 To 3
                aCols(intI) = Parse(strRecord,"^",intI)
            Next
        Else
            dctElements.Add strKey, ""
            For intI = 1 To 3
                aCols(intI) = "0"
            Next
        End If
        strNewItem = ""
        Select Case GetStatusID(intElementStatus)
            Case 30, 60, 22 'Correct and Yes (For Arrearage, Correct is 60, all other use 30.  IG uses Yes/No, DI uses Yes/No/NA)
                intI = 1
            Case 24 'NA
                intI = 3
            Case Else
                intI = 2
        End Select
        
        aCols(intI) =  CInt(aCols(intI)) + 1
        
        strNewItem = strNewItem & aCols(1) & "^" & aCols(2) & "^" & aCols(3) & "^"
        dctElements(strKey) = strNewItem
        Call AddTab(strKey, intI)
    End Sub
    
    Private Function GetStatusID(strElementStatus)
        Select Case strElementStatus
            Case "Correct"
                GetStatusID = 30
            Case "Yes"
                GetStatusID = 22
            Case "N/A"
                GetStatusID = 24
            Case "N/R"
                GetStatusID = 25
            Case "No"
                GetStatusID = 23
            Case Else
                GetStatusID = 100
        End Select
    End Function
    
    Private Sub AddTab(strKey, intColID)
        Dim strTabKey, strRecord, strNewRecord
        Dim intI
        
        strTabKey = Parse(strKey,"^",1) & "^" & Parse(strKey,"^",2)
        If dctTabs.Exists(strTabKey) Then
            strRecord = dctTabs(strTabKey)
        Else
            dctTabs.Add strTabKey, ""
            strRecord = "0^0^0"
        End If
        strNewRecord = ""
        For intI = 1 To 3
            If CInt(intI) = CInt(intColID) Then
                strNewRecord = strNewRecord & CInt(Parse(strRecord,"^",intI)) + 1 & "^"
            Else
                strNewRecord = strNewRecord & Parse(strRecord,"^",intI) & "^"
            End If
        Next
        dctTabs(strTabKey) = strNewRecord
    End Sub
    
    Public Function GetTabID(strTabName)
        GetTabID = -1
        Select Case strTabName
            Case "Action Integrity"
                GetTabID = 1
            Case "Data Integrity"
                GetTabID = 2
            Case "Information Gathering"
                GetTabID = 3
        End Select
    End Function
    
    Public Function RollingScreenStatus(intArrayStatus, intRSStatus)
        If intRSStatus = 23 Or intArrayStatus = 23 Then
            RollingScreenStatus = 23
                    'Response.Write "here 23<br>" 
            Exit Function
        End If
        'If you get this far, niether status is 23, so if passing in 22, status is 22
        If intRSStatus = 22 Then
            RollingScreenStatus = 22
                    'Response.Write "here 22<br>" 
            Exit Function
        End If
        'If you get this far, status is either 22, 24 or 0
        If intRSStatus = 24 Then
            If intArrayStatus = 22 Then
                RollingScreenStatus = 22
            Else
                RollingScreenStatus = 24
            End If
                    'Response.Write "here 24<br>" 
            Exit Function
        End If
    End Function
    
    Public Sub ProcessRecordset()
        Dim strKey, oItem, strNewItem
        Dim aCols(3)
        Dim intI, intCol
        Dim aDIElmStatus(4)
        Dim dctAIKeys
        Dim blnLastDI
        
        aDIElmStatus(1) = 0
        aDIElmStatus(2) = 0
        aDIElmStatus(3) = 0
        aDIElmStatus(4) = 0
        blnLastDI = False
        
        Set dctAIKeys = CreateObject("Scripting.Dictionary")
        Do While Not adRs.EOF
            If GetTabID(adRs.Fields("rveType").value) = 3 And blnLastDI = True Then
                'When first Info Gathering record is found, insert last DI record
                Call AddElement(aDIElmStatus(1) & "^2^" & aDIElmStatus(2), aDIElmStatus(4))
                blnLastDI = False
            End If
            If adRs.Fields("rvfFactor").value <> "" Then
                Call AddFactor(adRs.Fields("rveProgramID").value & "^" & GetTabID(adRs.Fields("rveType").value) & "^" & adRs.Fields("rveElement").value & "^" & adRs.Fields("rvfFactor").value, adRs.Fields("rvfStatus").value)
            End If
            Select Case GetTabID(adRs.Fields("rveType").value)
                Case 1
                    'Action Integrity elements can have decisions, but Action status is derived from element
                    'status.  Only add element 1 time per ReviewID+key
                    If Not dctAIKeys.Exists(adRs.Fields("rveReviewID").value & "^" & adRs.Fields("rveProgramID").value & "^" & GetTabID(adRs.Fields("rveType").value) & "^" & adRs.Fields("rveElement").value) Then
                        Call AddElement(adRs.Fields("rveProgramID").value & "^" & GetTabID(adRs.Fields("rveType").value) & "^" & adRs.Fields("rveElement").value, adRs.Fields("rveStatus").value)
                        dctAIKeys.Add adRs.Fields("rveReviewID").value & "^" & adRs.Fields("rveProgramID").value & "^" & GetTabID(adRs.Fields("rveType").value) & "^" & adRs.Fields("rveElement").value, ""
                    End If
                Case 2
                    'For Data Integrity, element status is determined from the factors
                    If aDIElmStatus(1) & "^" & aDIElmStatus(2) & "^" & aDIElmStatus(3) <> adRs.Fields("rveProgramID").value & "^" & adRs.Fields("rveElement").value & "^" & adRs.Fields("rveReviewID").value Then
                        'Screen changed, add Element status to object
                        If aDIElmStatus(1) > 0 Then
                         'rESPONSE.Write  aDIElmStatus(1) & "^" & aDIElmStatus(2) & "^" & aDIElmStatus(3) & "  ==  " & adRs.Fields("rveProgramID").value & "^" & adRs.Fields("rveElement").value & "^" & adRs.Fields("rveReviewID").value & "<br>"
                         '
                            Call AddElement(aDIElmStatus(1) & "^2^" & aDIElmStatus(2), aDIElmStatus(4))
                        End If                   
                        aDIElmStatus(1) = adRs.Fields("rveProgramID").value
                        aDIElmStatus(2) = adRs.Fields("rveElement").value
                        aDIElmStatus(3) = adRs.Fields("rveReviewID").value
                        aDIElmStatus(4) = 0
                    End If
                    aDIElmStatus(4) = RollingScreenStatus(aDIElmStatus(4), GetStatusID(adRs.Fields("rvfStatus").value))
                    blnLastDI = True
                Case 3
                    Call AddElement(adRs.Fields("rveProgramID").value & "^" & GetTabID(adRs.Fields("rveType").value) & "^" & adRs.Fields("rveElement").value, adRs.Fields("rveStatus").value)
            End Select
            adRs.MoveNext
        Loop            
        If blnLastDI = True Then
            'Check last DI record, incase there were no IG records
            Call AddElement(aDIElmStatus(1) & "^2^" & aDIElmStatus(2), aDIElmStatus(4))
            blnLastDI = False
        End If
    End Sub
    
End Class
%>
<!--#include file="IncRptFooter.asp"-->
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->