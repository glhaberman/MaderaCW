<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: RptReviewsReturnedDet.asp                                            '
'==========================================================================='
Const icWRK = 0
Const icSUP = 1
Const icMGR = 2
Const icOFF = 3
Const icDIR = 4
Const icFIN = 5
				
Dim intMglLevel
Dim adCmd
Dim adRs
Dim adMultiPosIDs

Dim oCounts
Dim aDoTotals(6)
Dim aShowLevel(6)
Dim intI
Dim intJ
Dim dblPercent
Dim mstrPageTitle
Dim blnTempTot

Dim intHeader
Dim intleft
Dim intwidth
Dim strClass
Dim strAlign
Dim intReqLevel
Dim intTempReqLevel
Dim strTempName

Dim intShadeLevel
Dim intShadeCount
Dim strWkrColor
Dim strSupColor
Dim strBackColor
Dim strFontColor

Dim blnNewTable
Dim mintTableID
Dim mstrTableMode
Dim strComments
Dim strTemp

%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%

mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
Set adCmd = GetAdoCmd("spRptEligElemComments")
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
    AddParmIn adCmd, "@ElementName", adVarChar, 255, ReqIsBlank("EligElementText")
    AddParmIn adCmd, "@StartReviewMonth", adDBTimeStamp, 0, ReqIsDate("StartReviewMonth")
    AddParmIn adCmd, "@EndReviewMonth", adDBTimeStamp, 0, ReqIsDate("EndReviewMonth")
	'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
%>
<HTML>
    <HEAD>
        <META name="vs_targetSchema" content="HTML 4.0">
        <TITLE>
            <%=Request.Form("ReportTitle")%>
        </TITLE>
        <!--#include file="IncDefStyles.asp"-->
        <!--#include file="IncRptStyles.asp"-->
    </HEAD>
    <SCRIPT ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
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
-->
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->
<!--#include file="IncRptExpParms.asp"-->
<!--#include file="IncSvrFunctions.asp"-->

<!--=== Start of Report Definition and Layout ============================= -->
<!--#include file="IncRptHeader.asp"-->

    <DIV id=PageFrame class=RptPageFrame>
        <%
        Response.Write "<BR><BR><BR><BR><BR><BR><BR><BR>"
        Call WriteCriteria()
        
		If adRs.EOF Then
            Response.Write "<BR><BR>"
            Response.Write "<SPAN id=lblNoResults class=ReportText "
            Response.Write "style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
            Response.Write " * No reviews matched the report criteria *"
        End If
        
        Set oCounts = New clsCounters
		
		strSupColor = "#FFEFD5"
		
        'Initialize the flags for showing totals:
        For intI = 0 To 5
			aDoTotals(intI) = False
		Next
		
		intReqLevel = icWRK
		
		intShadeCount = 0
		strWkrColor = "#ffffff"
		strSupColor = "#FFEFD5"
		If intReqLevel = icDIR Then
			Call WriteColumnHeaders("")
		End If
        
        Response.Write vbCrLf
        Response.Write vbCrLf
        mintTableID = 0
        mstrTableMode = "End"
        Do While Not adRs.EOF
			For intI = icWRK To icDIR
				oCounts.CurrentName(intI) = adRs.Fields(oCounts.Field(intI)).Value
			Next
            
            'Place a break between Totals
			For intI = intReqLevel to icDIR
				If ocounts.EmployeeCount(intI) >= 1 AND oCounts.Changed(intI) AND intI >= intReqLevel AND oCounts.PreviousName(intReqLevel) <> "XXX" Then
					'Response.Write "<BR style=""FONT-SIZE:8"">"
					'Response.Write "<SPAN id=lblWkr style=""LEFT:10; WIDTH:640;HEIGHT:20;FONT-SIZE:14;BORDER-COLOR:#C0C0C0; BORDER-TOP-STYLE:double; BORDER-TOP-WIDTH:3""> </SPAN>"
					'Response.Write "<BR style=""FONT-SIZE:20"">"
					Exit For
				End IF
			Next
		
            'Reset values -- specialized for RptResponseDue
            intHeader = intReqLevel
			          
			For intI = intTempReqLevel To icDIR
				aShowLevel(intI) = False
			Next	
			oCounts.ShowLevel(icDIR) = False
			intTempReqLevel = intReqLevel
			If intTempReqLevel > icWRK Then
				'Find the highest valid management level
				Do While oCounts.CurrentName(intTempReqLevel) = oCounts.CurrentName(intTempReqLevel + 1)
					intTempReqLevel = intTempReqLevel - 1
				Loop
			End If
			
			oCounts.ShowLevel(intTempReqLevel) = True
			
			'Do not print repeated names
            For intI = intTempReqLevel To icDIR
				If oCounts.CurrentName(intI) <> oCounts.CurrentName(intI + 1) Then
					aShowLevel(intI) = True
				End If
			Next
			
			'reset intshadecount if upper management has changed
			If oCounts.Changed(intReqLevel) Then
				intShadeCount = 0 
			End If
			
            If intShadeCount MOD 2 = 0 Then
				strWkrColor = "#ffffff"
			Else 
				strWkrColor = ReqForm("WkrBackGround")
			End If
			      
			blnNewTable = True
            For intI = icSup To intReqLevel Step -1
				If oCounts.Changed(intI) AND oCounts.ShowLevel(intI) Then
					If aShowLevel(intI) Then
					    If blnNewTable = True Then
                            Call WriteTable("Start")
                            blnNewTable = False
					    End If
						strClass = oCounts.HeaderClass(intI)
						strFontColor = "#000000" 
						If intI = intTempReqLevel Then
							strClass = "ManagementText"
							strBackColor = ReqForm("SupBackGround")
							strFontColor = ReqForm("SupFontColor")
						Else
							strBackColor = strWkrColor
						End If
						If intI < intReqLevel Then
							Call WriteName("*&nbsp" & oCounts.NameOnly(intI), intI)
						Else
							Call WriteName(oCounts.NameOnly(intI), intI)
						End If
					End If
					
					If intI = intHeader Then
                        Call WriteTable("Start")
						Call WriteColumnHeaders("")
					End If
				End If
			Next  
			If blnNewTable = False Then
    			Call WriteTable("Start")
			End If

            'Translate CrLf's into HTML BR tags:
            strComments = adRs.Fields("rvcComments").Value
            strTemp = ""
            For intI = 1 To Len(strComments)
                If Asc(Mid(strComments, intI, 1)) = 13 Then
                    strTemp = strTemp & "<BR>"
                End If
                If Asc(Mid(strComments, intI, 1)) <> 10 Then
                    strTemp = strTemp & Mid(strComments, intI, 1)
                End If
            Next
            strComments = Replace(strTemp, Chr(9) & "#dq#", """")
            strComments = Replace(strComments, Chr(9) & "#sq#", "'")
            strComments = Replace(strComments, Chr(9) & "#ca#", "^")
            strComments = Replace(strComments, Chr(9) & "#ba#", "|")
            Call WriteRow(adRs.Fields("rvwCaseNumber").Value, strComments)
            Call WriteRow("", "")

            intShadeCount = intShadeCount + 1
            adRs.MoveNext
        Loop 
        If mstrTableMode = "Start" Then
            Call WriteRow("", "")
            Call WriteTable("End")
        End If
    %>
<!--#include file="IncRptFooter.asp"-->

</HTML>
<%

Sub WriteTable(strAction)
    If strAction = "Start" Then
        If mstrTableMode = "Start" Then
            Response.Write "</TABLE>"
            Response.Write vbCrLf
        End If
        Response.Write "<TABLE id=lstComments" & mintTableID & " rules=all cellspacing=0 border=0 width=650 style=""FONT-SIZE:10pt; TABLE-LAYOUT:auto; overflow:visible; VISIBILITY:visible"">"
        Response.Write vbCrLf
        mintTableID = mintTableID + 1
        mstrTableMode = "Start"
    Else
        Response.Write "</TABLE>"
        Response.Write vbCrLf
        mstrTableMode = "End"
    End If
End Sub

Sub WriteRow(strCol1, strCol2)
    Response.Write "<tr>"
    Response.Write "<TD class=RptGenericCell width=100>"
    Response.Write "&nbsp" & strCol1
    Response.Write "</TD>"
    Response.Write "<TD class=RptGenericCell width=550>"
    Response.Write strCol2
    Response.Write "</TD>"
    Response.Write "</tr>"
    Response.Write vbCrLf
End Sub

Sub WriteName(strName, intLevel)
    Dim strWeight, strWeightClose
    Dim strBackColor, strFontColor
						
	strFontColor = "#000000" 
    strBackColor = "#ffffff"
    
    strWeight = ""
    strWeightClose = ""
    If intLevel >= 1 Then
        strWeight = "<B>"
        strWeightClose = "</B>"
    End If
    If intLevel = 0 Then
		strBackColor = ReqForm("SupBackGround")
		strFontColor = ReqForm("SupFontColor")
    End If
    strName = "&nbsp" & strName
    Response.Write "<tr>"
    Response.Write "<TD class=RptGenericCell width=200 colspan=2 nowrap=true style=""overflow:visible;padding-left:1;color:" & strFontColor & ";background-color:" & strBackColor & """>"
    Response.Write strWeight & strName & strWeightClose
    Response.Write "</TD>"
    Response.Write "<TD class=RptGenericCell width=450 style=""color:" & strFontColor & ";background-color:" & strBackColor & """>"
    Response.Write "&nbsp"
    Response.Write "</TD>"
    Response.Write "</tr>"
    Response.Write vbCrLf
End Sub

Sub WriteColumnHeaders(strNameTitle) 
    Dim strBackColor, strFontColor
    strBackColor = ReqForm("ColBackGround")
    strFontColor = ReqForm("ColFontColor")
    Response.Write "<tr>"
    Response.Write "<TD class=RptGenericCell width=100 nowrap=true style=""padding-left:1;color:" & strFontColor & ";background-color:" & strBackColor & """>"
    Response.Write "<B>Case Number</B>"
    Response.Write "</TD>"
    Response.Write "<TD class=RptGenericCell width=550 style=""color:" & strFontColor & ";background-color:" & strBackColor & """>"
    Response.Write "<B>Comments</B>"
    Response.Write "</TD>"
    Response.Write "</tr>"
    Response.Write vbCrLf
End Sub            
                
Class clsCounters
    'Arrays hold counters.  Indexes 0-5:
    '   0 = worker,    1 = supervisor, 
    '   2 = manager,   3 = office,
    '   4 = director   5 = final total
    Private aFormNames(6)	   'Form Names
    Private aCurrentNames(6)   'Current Names
    Private aPreviousNames(6)  'Previous Names
    Private aEmployeeCount(6)  'Count of Employees for each level
    Private aNameChanged(6)    'Current Differs from Last
    Private aShowLevel(6)      'Keeps track of what levels are displayed
    Private aCFieldNames(6)    'Keeps track of the level's SQL field name
    Private aCTotalClass(6)    'Keeps track of the level's style class for totals
    Private aCHeaderClass(6)   'Keeps track of the level's style class for headings
    
    Private Sub Class_Initialize()
        Dim intI
       'Initialize property holders:
        aCFieldNames(0) = "Worker"
        aCFieldNames(1) = "Supervisor"
        aCFieldNames(2) = "Manager"
        aCFieldNames(3) = "Office"
        aCFieldNames(4) = "Director"
        
        aCHeaderClass(0) = "WorkerHeading"
        aCHeaderClass(1) = "SupervisorHeading"
        aCHeaderClass(2) = "ManagerHeading"
        aCHeaderClass(3) = "OfficeHeading"
        aCHeaderClass(4) = "DirectorHeading"
        
        aCTotalClass(0) = "WorkerTotals"
        aCTotalClass(1) = "SupervisorTotals"
        aCTotalClass(2) = "ManagerTotals"
        aCTotalClass(3) = "OfficeTotals"
        aCTotalClass(4) = "DirectorTotals"
        
        aFormNames(0) = ReqForm("Worker")
        aFormNames(1) = ReqForm("Supervisor")
        aFormNames(2) = ReqForm("ProgramManager")
        aFormNames(3) = ReqForm("Office")
        aFormNames(4) = ReqForm("Director")
        For intI = icWRK To icFIN
            aCurrentNames(intI) = "XXX"
            aPreviousNames(intI) = "XXX"
            aEmployeeCount(intI) = 1
            aNameChanged(intI) = False
            aShowLevel(intI) = False
        Next
    End Sub

    Public Property Let CurrentName(intWhich, strVal)
		Dim intI
		
        'Move current value to holder for previous name:
        aPreviousNames(intWhich) = aCurrentNames(intWhich)
        If instr(1, strVal, "*") > 0 Then
			aCurrentNames(intWhich) = Parse(strVal, "*", 2)
		Else
			aCurrentNames(intWhich) = strVal
		End If
        'If the name is changing, set the changed flag:
        If aPreviousNames(intWhich) <> aCurrentNames(intWhich) Then
            aNameChanged(intWhich) = True
            'Reset counters for level of the name:
            aEmployeeCount(intWhich) = aEmployeeCount(intWhich) + 1
            'Reset Count of Employees in lower Management levels. 
            For intI = icWRK To intWhich - 1
				aEmployeeCount(intI) = 1
			Next
        Else
            aNameChanged(intWhich) = False
        End If
    End Property
    
    Public Property Get Field(intWhich)
		Field = aCFieldNames(intWhich)
    End Property
    
    Public Property Get HeaderClass(intWhich)
		HeaderClass = aCHeaderClass(intWhich)
    End Property
    
    Public Property Get TotalClass(intWhich)
		TotalClass = aCTotalClass(intWhich)
    End Property
    
    Public Property Get FormName(intWhich)
		FormName = aFormNames(intWhich)
    End Property
    
    Public Property Get Changed(intWhich)
        Changed = aNameChanged(intWhich)
    End Property

    Public Property Get CurrentName(intWhich)
        CurrentName = aCurrentNames(intWhich)
    End Property

    Public Property Get PreviousName(intWhich)
        PreviousName = aPreviousNames(intWhich)
    End Property
        
    Public Property Let ShowLevel(intWhich, blnVal)
        Dim intI
        Dim intBeg
        Dim intEnd
        Dim intStep
        
        intBeg = intWhich
        If blnVal Then
            intEnd = icFIN
            intStep = 1
        Else
            intEnd = icWRK
            intStep = -1
        End If
        
        For intI = intBeg To intEnd Step intStep
            aShowLevel(intI) = blnVal
        Next
    End Property
    
    Public Property Get ShowLevel(intWhich)
        ShowLevel = aShowLevel(intWhich)
    End Property
    
    Public Property Get EmployeeCount(intWhich)
        EmployeeCount = aEmployeeCount(intWhich)
    End Property
    
    Public Property Get NameOnly(intWhich)
        NameOnly = Parse(aCurrentNames(intWhich), "--", 1)
    End Property
End Class

%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->