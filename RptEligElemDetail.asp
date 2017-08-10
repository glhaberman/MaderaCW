<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: EligElemDetail.asp                                              '
'  Purpose: Displays the Eligibility Element Detail report, based on the    '
'           criteria passed to this page by the previous criteria screen.   '
'==========================================================================='
Const icWRK = 0
Const icSUP = 1
Const icMGR = 2
Const icDIR = 3
Const icFIN = 4
				
Dim intMglLevel
Dim adCmd
Dim adRs

Dim oCounts
Dim aDoTotals(5)
Dim aShowLevel(5)
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
Dim mstrTableName
Dim mstrCurrentName
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<!--#include file="IncDrillDownSvr.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
Set adCmd = GetAdoCmd("spRptEligElemDet")
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
    AddParmIn adCmd, "@DetailStatusID", adInteger, 0, Null ' Only used for detail report
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
%>

<HTML>
<HEAD>
    <TITLE><%=Request.Form("ReportTitle")%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncRptStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim blnCloseClicked

Sub window_onload
	Call FormShow("none")
	PageBody.style.cursor = "wait"
    If Form.UserID.Value = "" Then
        MsgBox "User not recognized.  Logon failed, please try again.", vbinformation, "Log On"
        window.navigate "Logon.asp"
    End If
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
    Call DrillDownColClickEvent("spRptEligElemDet", intColID, intRowID, False, 0)
End Sub 

-->
</SCRIPT>
<!--#include file="IncRptExpParms.asp"-->
<!--#include file="IncCmnCliFunctions.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<!--#include file="IncRptHeader.asp"-->
<!--#include file="IncDrillDownCli.asp"-->
<DIV id=PageFrame
    style="HEIGHT:225; WIDTH:650; TOP:116; LEFT:10; FONT-SIZE:10pt; padding-top:5">
    <%
    
    Call WriteCriteria()
    
	If adRs.EOF Then
        Response.Write "<BR><BR>"
        Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
        Response.Write " * No reviews matched the report criteria *"
    End If
    
    Set oCounts = New clsCounters
    
    'Initialize the flags for showing totals:
    For intI = icWRK To icFIN
		aDoTotals(intI) = False
	Next
	
	If ReqForm("Worker") <> "" Then
		'Show worker and above:
		intReqLevel = icWRK
	ElseIf ReqForm("Supervisor") <> "" Then
		'Show worker and above:
		intReqLevel = icWRK
	'ElseIf ReqForm("ProgramManager") <> "" Then
	'	'Show supervisor and above:
	'	intReqLevel = icSUP
	'ElseIf ReqForm("Director") <> "" Then
	'	'Show offices and above:
	'	intReqLevel = icSUP
	Else
		'Show directors:
		intReqLevel = icSUP
	End If
	
	intTempReqLevel = intReqLevel
	intShadeCount = 0
	strWkrColor = "#ffffff"
	strSupColor = "#FFEFD5"
	If intReqLevel = icDIR Then
		Call WriteColumnHeaders("")
	End If
    
    Do While Not adRs.EOF
		For intI = intReqLevel To icDIR
			oCounts.CurrentName(intI) = adRs.Fields(oCounts.Field(intI)).Value
		Next 
		
        'Place a break between Totals
		For intI = intReqLevel to icDIR
			If ocounts.EmployeeCount(intI) = 1 And aDoTotals(intI + 1) And intI >= intReqLevel Then
				Response.Write "<BR style=""FONT-SIZE:8"">"
				Response.Write "<SPAN id=lblWkr style=""LEFT:10; WIDTH:640;HEIGHT:20;FONT-SIZE:14;BORDER-COLOR:#C0C0C0; BORDER-TOP-STYLE:double; BORDER-TOP-WIDTH:3""> </SPAN>"
				Response.Write "<BR style=""FONT-SIZE:20"">"
				Exit For
			End If
		Next
	
        'Reset values
        For intI = icDIR To intReqLevel + 1 Step -1
			If oCounts.Changed(intI) And oCounts.CurrentName(intI) <> oCounts.CurrentName(intI + 1) Then
				intHeader = intI
			End If
        Next
        
		For intI = intTempReqLevel To icDIR
			aShowLevel(intI) = False
		Next	
		oCounts.ShowLevel(icDIR) = False
		intTempReqLevel = intReqLevel
		
		'Find the highest valid management level
        Do While oCounts.CurrentName(intTempReqLevel) = oCounts.CurrentName(intTempReqLevel + 1)
			intTempReqLevel = intTempReqLevel - 1
		Loop
		oCounts.ShowLevel(intTempReqLevel) = True
		
		intShadeLevel = icDIR
		'Do not print repeated names
            For intI = intTempReqLevel To icSup
			If oCounts.CurrentName(intI) <> oCounts.CurrentName(intI + 1) Then
				If intShadeLevel > intI And intI > 0 Or intI = intReqLevel + 1 Then
					intShadeLevel = intI
				End If
				aShowLevel(intI) = True
			End If
		Next
		
		'reset intshadecount if upper management has changed
		If oCounts.Changed(intReqLevel + 1) Then
			intShadeCount = 0 
		End If
		
        If intShadeCount MOD 2 = 0 Then
			strWkrColor = "#ffffff"
			strSupColor = "#FFEFD5"
		Else 
			strWkrColor = "#FFEFD5"
			strSupColor = "#ffffff"
		End If
		
        For intI = icDir To intTempReqLevel Step -1
			If oCounts.Changed(intI) And oCounts.ShowLevel(intI) Then
				If aShowLevel(intI) Then
					strClass = oCounts.HeaderClass(intI)
					strFontColor = "#000000" 
					If intI = intTempReqLevel Then
						strClass = "ManagementText"
						strBackColor = strWkrColor
					Else
						strClass = oCounts.HeaderClass(intI) 
						strBackColor = "#ffffff"
					End If
					'Spacer
					Response.Write "<SPAN id=lblWkr class=" & strClass & " "
					Response.Write "style=""WIDTH:640; BackGround:" & strBackColor & """>"
					Response.Write "</SPAN>"
					
					'Display the Director Name:	
					Response.Write "<SPAN id=lblDir class=" & strClass & " "
					Response.Write "style=""WIDTH:200; Color:" & strFontCOlor & "; BackGround:" & strBackColor & """>"
					If intI < intReqLevel Then
						Response.Write "*&nbsp" & oCounts.NameOnly(intI) & "</SPAN>"
					Else
						Response.Write oCounts.NameOnly(intI) & "</SPAN>"
					End If
					
					If intI > intTempReqLevel Then
						Response.Write "<BR style=""FONT-SIZE:11pt"">"
					End If
				End If
				
				If intI = intHeader Then
					Call WriteColumnHeaders("")
				End If
			End If
		Next  
		
		oCounts.TotalNotApp(icWRK) = adRs.Fields("TotalNA").Value
		oCounts.TotalCorrect(icWRK) = adRs.Fields("TotalCorrect").Value
		oCounts.TotalError(icWRK) = adRs.Fields("TotalError").Value
		oCounts.TotalCount(icWRK) = adRs.Fields("TotalCases").Value ' + adRs.Fields("Prg1CorrectCnt").Value + adRs.Fields("Prg1NotAppCnt").Value
		
		adRs.MoveNext
		If adRs.EOF Then
			For intI = intTempReqLevel To icFIN
				If aShowLevel(intI) Or intI = intTempReqLevel Then
					aDoTotals(intI) = True
				End If
			Next
			aDoTotals(icFIN) = True
			
		Else
			For intI = intTempReqLevel To 5
				aDoTotals(intI) = False
			Next
			For intI = intTempReqLevel To icDIR
                mstrTableName = ""
                mstrCurrentName = ""
                For intJ = intI To icDir
                    If InStr(1, adRs.Fields(oCounts.Field(intJ)).Value, "*") > 0 Then
                        strTempName = Parse(adRs.Fields(oCounts.Field(intJ)).Value, "*", 2)
                    Else
                        strTempName = adRs.Fields(oCounts.Field(intJ)).Value
                    End If
                    mstrCurrentName = mstrCurrentName & "[" & strTempName & "]"
                    mstrTableName = mstrTableName & "[" & oCounts.CurrentName(intJ) & "]"
                Next                    
				
                If mstrCurrentName <> mstrTableName Then
					IF aShowLevel(intI) OR intI = intTempReqLevel Then
						aDoTotals(intI) = True
					ElseIF intI > intTempReqLevel Then
						For intJ = intI + 1 to icDIR 
							If aShowLevel(intJ) AND oCounts.CurrentName(intI) = oCounts.CurrentName(intJ) Then
								aDoTotals(intI) = True
								Exit For
							End If
						Next
					End If
				End If
			Next			
		End If

		'Do not print totals for levels above the lowest criteria level
		For intI = intTempReqLevel + 1 To icDIR
			If aDoTotals(intI) Then
				For IntJ = intI + 1 To icFIN
					If oCounts.FormName(intI) <> "" Then
						aDoTotals(intJ) = False
					End If
				Next
			End If
		Next
		
		'Print total for the lowest level, usually the Worker
		If aDoTotals(intTempReqLevel) Then
			'Place Totals
			intShadeCount = intShadeCount + 1
			Call WriteTotals(intTempReqLevel, "ManagementText", strWkrColor, "#000000" )
			If aDoTotals(intTempReqLevel + 1) Then
				Response.Write "<BR style=""FONT-SIZE:8"">"
			Else
				Response.Write "<BR style=""FONT-SIZE:15"">"
			End If
		End If
		
		For intI = intTempReqLevel + 1 To icDIR
			'Place totals for previous Supervisor: 
			If aDoTotals(intI) and aShowLevel(intI)Then
				strClass = oCounts.TotalClass(intI)
				strFontColor = "#000000"
				'Place totals for previous Supervisor:
				If oCounts.ShowLevel(intI - 1) Then
					strBackColor = "#ffffff"
					strFontColor = "#000000"
					strClass = oCounts.TotalClass(intI)
					Response.Write "<BR style=""FONT-SIZE:10	"">"
					Response.Write "<span id=lblSup class=" & strClass & " "
					Response.Write "style=""LEFT:10; WIDTH:640; BACKGROUND:" & strBackColor & """></SPAN>"
					
					Response.Write "<span id=lblSup class=" & strClass & " "
					Response.Write "style=""Color:" & strFontColor & "; BACKGROUND:" & strBackColor & "; TEXT-ALIGN:LEFT "">"
					Response.Write oCounts.NameOnly(intI) & "&nbspTotal:" & "</Span>"
					
				Else
					strBackColor = strWkrColor
					intShadeCount = intShadeCount + 1
					strClass = "ManagementText"
				End If
				Call WriteTotals(intI, strClass, strBackColor, strFontColor)
				
				Response.Write "<BR style=""FONT-SIZE:8"">"
			End If
		Next
    Loop
	If oCounts.EmployeeCount(intTempReqLevel) >= 2 Then
		aDoTotals(icFin) = True
	End If
	For inti = intTempReqLevel To icDIR 
		If oCounts.FormName(intI) <> "" Then
			aDoTotals(icFin) = False
		End If
	Next
	If aDoTotals(icFin) Then
		Response.Write "<BR style=""FONT-SIZE:16"">"

		Response.Write "<SPAN id=lblAvgLabel class=DirectorTotals "
		Response.Write "style=""WIDTH:640; LEFT:10; BACKGROUND:#FFFFFF; TEXT-ALIGN:left"">"
		Response.Write "Final Total:</SPAN>"
		
		'Place Totals
		intShadeCount = intShadeCount + 1
		Call WriteTotals(icFIN, "DirectorTotals", "#FFFFFF", "#000000" )
		
		Response.Write "<BR style=""FONT-SIZE:8"">"
	End If
			
	Response.Write "<BR style=""FONT-SIZE:8"">"
	Response.Write "<SPAN id=lblWkr style=""LEFT:10; WIDTH:640;HEIGHT:20;FONT-SIZE:14;BORDER-COLOR:#C0C0C0; BORDER-TOP-STYLE:double; BORDER-TOP-WIDTH:3""> </SPAN>"
	Response.Write "<BR style=""FONT-SIZE:20"">"
    %>

<!--#include file="IncRptFooter.asp"-->
</HTML>
<%
Sub WriteTotals(intWhich, strClass, strBackColor, strFontColor)
    Dim sngAccuracyRate
    Dim strHidden, strName
    Dim intQ
    
    Call oCounts.NewRow()

    'Total
    Call WriteNames(intWhich)
    Call WriteColumn(1,strClass,oCounts.TotalCount(intWhich),200,strBackColor,"",80,0)
    Call WriteColumn(2,strClass,oCounts.TotalCorrect(intWhich),490,strBackColor,"",80,0)
    Call WriteColumn(3,strClass,oCounts.TotalError(intWhich),340,strBackColor,"",80,0)

	Response.Write "<SPAN id=lblWrk class=" & strClass & " "
	Response.Write "style=""LEFT:265; WIDTH:80; COLOR:" & strFontColor & "; background:" & strBackColor & "; TEXT-ALIGN:Center""> "
	Response.Write oCounts.TotalNotApp(intWhich) & "</SPAN>"

	Response.Write "<SPAN id=lblErrorPercent class=" & strClass & " "
	Response.Write "style=""LEFT:415; WIDTH:80; COLOR:" & strFontColor & "; BACKGROUND:" & strBackColor & "; TEXT-ALIGN:right"">"
	Response.Write oCounts.PercentError(intWhich) & "</SPAN>"
	
	Response.Write "<SPAN id=lblPMPercentCorrect1 class=" & strClass & " "
	Response.Write "style=""LEFT:560; WIDTH:80; COLOR:" & strFontColor & "; BACKGROUND:" & strBackColor & ";TEXT-ALIGN:right"">"
	Response.Write oCounts.PercentCorrect(intWhich) & "</SPAN>"
End Sub
Sub WriteColumnHeaders(strNameTitle)
	'Layout row 1 of the column headings:
	Response.Write "<SPAN id=lblPrmHdr1 class=ColumnHeading "
	Response.Write "style=""LEFT:10; WIDTH:640; BORDER-BOTTOM-STYLE:none; Background:" & strSupColor & """>"
	Response.Write "</SPAN>"
	
	Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
	Response.Write "style=""WIDTH:200; LEFT:10; BORDER-BOTTOM-STYLE:none; Background:" & strSupColor & """>"
	Response.Write "</SPAN>"

	Response.Write "<SPAN id=lblTotalCasesHdr class=ColumnHeading "
	Response.Write "style=""LEFT:200; BORDER-BOTTOM-STYLE:none"">"
	Response.Write "Total </SPAN>"

	Response.Write "<SPAN id=lblErrorCntHdr class=ColumnHeading "
	Response.Write "style=""LEFT:265; BORDER-BOTTOM-STYLE:none"">"
	Response.Write "Number </SPAN>"

	Response.Write "<SPAN id=lblErrorCntHdr class=ColumnHeading "
	Response.Write "style=""LEFT:340; BORDER-BOTTOM-STYLE:none"">"
	Response.Write "Number </SPAN>"
	
	Response.Write "<SPAN id=lblErrorPercentHdr class=ColumnHeading "
	Response.Write "style=""LEFT:415; BORDER-BOTTOM-STYLE:none"">"
	Response.Write "Percent </SPAN>"
	
	Response.Write "<SPAN id=lblCorrectCntHdr class=ColumnHeading "
	Response.Write "style=""LEFT:490; BORDER-BOTTOM-STYLE:none"">"
	Response.Write "Number </SPAN>"
	
	Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
	Response.Write "style=""LEFT:560; BORDER-BOTTOM-STYLE:none"">"
	Response.Write "Percent </SPAN>"
	
	Response.Write "<BR>"
	'Layout second row of the column headings:			
	Response.Write "<SPAN id=lblPrmHdr1 class=ColumnHeading "
	Response.Write "style=""LEFT:10; WIDTH:640; BORDER-TOP-STYLE:none; Background:" & strSupColor & """>"
	Response.Write "</SPAN>"
	
	Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
	Response.Write "style=""WIDTH:200; LEFT:10; BORDER-TOP-STYLE:none"">"
	Response.Write strNameTitle & "&nbspName</SPAN>"

	Response.Write "<SPAN id=lblTotalCasesHdr class=ColumnHeading "
	Response.Write "style=""LEFT:200; BORDER-TOP-STYLE:none"">"
	Response.Write "Cases </SPAN>"

	Response.Write "<SPAN id=lblErrorCntHdr class=ColumnHeading "
	Response.Write "style=""LEFT:265; BORDER-TOP-STYLE:none"">"
	Response.Write "N/A </SPAN>"

	Response.Write "<SPAN id=lblErrorCntHdr class=ColumnHeading "
	Response.Write "style=""LEFT:340; BORDER-TOP-STYLE:none"">"
	Response.Write "Incorrect </SPAN>"
	
	Response.Write "<SPAN id=lblErrorPercentHdr class=ColumnHeading "
	Response.Write "style=""LEFT:415; BORDER-TOP-STYLE:none"">"
	Response.Write "Incorrect </SPAN>"
	
	Response.Write "<SPAN id=lblCorrectCntHdr class=ColumnHeading "
	Response.Write "style=""LEFT:490; BORDER-TOP-STYLE:none"">"
	Response.Write "Correct </SPAN>"
	
	Response.Write "<SPAN id=lblCorrectPercentHdr class=ColumnHeading "
	Response.Write "style=""LEFT:560; BORDER-TOP-STYLE:none"">"
	Response.Write "Correct </SPAN>"
	Response.Write "<br>"  

	Response.Write  "<br>"
End Sub

Class clsCounters
    'Arrays hold counters.  Indexes 0-5:
    '   0 = worker,    1 = supervisor, 
    '   2 = manager,   3 = office,
    '   4 = director   5 = final total
    Private aFormNames(5)	   'Form Names
    Private aCurrentNames(5)   'Current Names
    Private aPreviousNames(5)  'Previous Names
    Private aEmployeeCount(5)  'Count of Employees for each level
    Private aTotalNotApp(5)     'Total Not Applicable
    Private aTotalCount(5)     'Total Count
    Private aTotalError(5)	   'Error Count
    Private aTotalCorrect(5)   'Correct Count
    Private aNameChanged(5)    'Current Differs from Last
    Private aShowLevel(5)      'Keeps track of what levels are displayed
    Private aCFieldNames(5)    'Keeps track of the level's SQL field name
    Private aCTotalClass(5)    'Keeps track of the level's style class for totals
    Private aCHeaderClass(5)   'Keeps track of the level's style class for headings
    Private mintRowID
    
    Private Sub Class_Initialize()
        Dim intI
       'Initialize property holders:
        aCFieldNames(0) = "Worker"
        aCFieldNames(1) = "Supervisor"
        aCFieldNames(2) = "Manager"
        aCFieldNames(3) = "Director"
        aCFieldNames(4) = "Total"
        
        aCHeaderClass(0) = "WorkerHeading"
        aCHeaderClass(1) = "SupervisorHeading"
        aCHeaderClass(2) = "ManagerHeading"
        aCHeaderClass(3) = "DirectorHeading"
        
        aCTotalClass(0) = "WorkerTotals"
        aCTotalClass(1) = "SupervisorTotals"
        aCTotalClass(2) = "ManagerTotals"
        aCTotalClass(3) = "DirectorTotals"
        
        aFormNames(0) = ReqForm("Worker")
        aFormNames(1) = ReqForm("Supervisor")
        aFormNames(2) = ReqForm("ProgramManager")
        aFormNames(3) = ReqForm("Director")
        For intI = icWRK To icFIN
            aCurrentNames(intI) = "XXX"
            aPreviousNames(intI) = "XXX"
            aEmployeeCount(intI) = 0
            aTotalNotApp(intI) = 0
            aTotalCount(intI) = 0
            aTotalError(intI) = 0
            aTotalCorrect(intI) = 0 
            aNameChanged(intI) = False
            aShowLevel(intI) = False
        Next
        mintRowID = -1
    End Sub

    Public Sub NewRow()
        mintRowID = mintRowID + 1
    End Sub
    
    Public Property Get RowID()
        RowID = mintRowID
    End Property

    Public Property Let CurrentName(intWhich, strVal)
		Dim intI
        Dim strPreviousName, strCurrentName
		
        'Move current value to holder for previous name:
        aPreviousNames(intWhich) = aCurrentNames(intWhich)
        If instr(1, strVal, "*") > 0 Then
			aCurrentNames(intWhich) = Parse(strVal, "*", 2)
		Else
			aCurrentNames(intWhich) = strVal
		End If
		
        ' After the highest level has been set, check each level to see if anything has changed
        If intWhich = icDir Then
            For intWhich = icWrk To icDir
                'If the name is changing, set the changed flag:
                strPreviousName = ""
                strCurrentName = ""
                For intI = intWhich To icDIR
                    strPreviousName = strPreviousName & "[" & aPreviousNames(intI) & "]"
                    strCurrentName = strCurrentName & "[" & aCurrentNames(intI) & "]"
                Next
                If strPreviousName <> strCurrentName Then
                    aNameChanged(intWhich) = True
                    'Reset counters for level of the name:
                    aTotalNotApp(intWhich) = 0
                    aTotalCount(intWhich) = 0
                    aTotalError(intWhich) = 0
                    aTotalCorrect(intWhich) = 0
                    aEmployeeCount(intWhich) = aEmployeeCount(intWhich) + 1
                    'Reset Count of Employees in lower Management levels. 
                    For intI = icWRK To intWhich - 1
                        aEmployeeCount(intI) = 1
                    Next
                Else
                    aNameChanged(intWhich) = False
                End If
            Next
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
    
    Public Property Let TotalNotApp(intWhich, intVal)
        Dim intI
        
        For intI = intWhich To icFIN
            aTotalNotApp(intI) = CLng(aTotalNotApp(intI)) + CLng(intVal)
        Next
    End Property
    
    Public Property Get TotalNotApp(intWhich)
        TotalNotApp = aTotalNotApp(intWhich)
    End Property
    
    Public Property Let TotalCount(intWhich, intVal)
        Dim intI
        
        For intI = intWhich To icFIN
            aTotalCount(intI) = CLng(aTotalCount(intI)) + CLng(intVal)
        Next
    End Property
    
    Public Property Get TotalCount(intWhich)
        TotalCount = aTotalCount(intWhich)
    End Property
    
    Public Property Let TotalError(intWhich, intVal)
        Dim intI
        For intI = intWhich To icFIN
            aTotalError(intI) = aTotalError(intI) + intVal
        Next
    End Property
    
    Public Property Get TotalError(intWhich)
        TotalError = aTotalError(intWhich)
    End Property
    
    Public Property Let TotalCorrect(intWhich, dblVal)
        Dim intI
        
        If Not IsNumeric(dblVal) Then
            dblVal = 0
        End If
        For intI = intWhich To icFIN
            aTotalCorrect(intI) = aTotalCorrect(intI) + dblVal
        Next
    End Property
    
    Public Property Get TotalCorrect(intWhich)
        TotalCorrect = FormatNumber(aTotalCorrect(intWhich), 0, True, False, True)
    End Property
    
    Public Property Get PercentCorrect(intWhich)
        Dim dblPercent

        If aTotalCorrect(intWhich) > 0 And aTotalCount(intWhich) > 0 Then 
			dblPercent = CDbl(aTotalCorrect(intWhich)/(aTotalCount(intWhich) - aTotalNotApp(intWhich))) * 100
		Else
			dblPercent = "0.00"
		End If

        PercentCorrect = FormatNumber(dblPercent, 2, True, False, True) & " %"
    End Property
    
    Public Property Get PercentError(intWhich)
        Dim dblPercent

        If aTotalError(intWhich) > 0 And aTotalCount(intWhich) > 0 Then 
			dblPercent = CDbl(aTotalError(intWhich)/(aTotalCount(intWhich) - aTotalNotApp(intWhich))) * 100
		Else
			dblPercent = "0.00"
		End If

        PercentError = FormatNumber(dblPercent, 2, True, False, True) & " %"
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

End Class%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->