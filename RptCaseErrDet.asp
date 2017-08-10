<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: RptCaseErrDet.asp                                            '
'  Purpose:																	'
'==========================================================================='
Const icWRK = 0
Const icSUP = 1
Const icMGR = 2
Const icOFF = 3
Const icDIR = 4
Const icFIN = 5
				
Dim adCmd
Dim adRs

Dim oCounts
Dim aDoTotals(6)
Dim aShowLevel(6)
Dim intI
Dim intJ
Dim mstrPageTitle
Dim intHeader

Dim intleft
Dim strClass
Dim intReqLevel
Dim intTempReqLevel
Dim strTempName

Dim intShadeCount
Dim strWkrColor
Dim strSupColor
Dim strBackColor
%>

<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<!--#include file="IncDrillDownSvr.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
Set adCmd = GetAdoCmd("spRptCaseErrSum")
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
    AddParmIn adCmd, "@StartReviewMonth", adDBTimeStamp, 0, ReqIsDate("StartReviewMonth")
    AddParmIn adCmd, "@EndReviewMonth", adDBTimeStamp, 0, ReqIsDate("EndReviewMonth")
    AddParmIn adCmd, "@DrillDownID", adInteger, 0, Null
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)

%>

<HTML>
<HEAD>
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Request.Form("ReportTitle")%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncRptStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
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
    If "<%=ReqForm("ProgramText")%>" = "" Or "<%=ReqForm("ProgramText")%>" = "<All>" Then
        lblAppTitle.innerText = "<%=ReqForm("ReportTitle")%> : All Functions"
        'lblAppTitle.style.font-weight = "bold"
        lblAppTitle.style.fontweight = "bold"
    End If
    PageBody.style.cursor = "default"
    cmdPrint1.focus
End Sub

Sub cmdClose_onclick()	
    window.close
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
    Call DrillDownColClickEvent("spRptCaseErrSum", intColID, intRowID, True, 0)
End Sub 
-->
</SCRIPT>
<!--#include file="IncDrillDownCli.asp"-->
<!--#include file="IncRptExpParms.asp"-->
<!--#include file="IncCmnCliFunctions.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<!--#include file="IncRptHeader.asp"-->
    <DIV id=PageFrame
        style="HEIGHT:225; WIDTH:650; TOP:116; LEFT:10; FONT-SIZE:10pt; padding-top:5">
        <%
        
        Call WriteCriteria()
        Dim mstrTableName
        Dim mstrCurrentName
        
        If adRs.EOF Then
            Response.Write "<BR><BR>"
            Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
            Response.Write " * No reviews matched the report criteria *"
        End If
        Set oCounts = New clsCounters
        
        'Initialize the flags for showing totals:
        For intI = 0 To 5
			aDoTotals(intI) = False
		Next
		
		'Show worker and above:
		intReqLevel = icWRK
		
		intTempReqLevel = intReqLevel
		intShadeCount = 0
		strWkrColor = "#ffffff"
		strSupColor = "#FFEFD5"
		If intReqLevel = icDIR Then
			Call WriteColumnHeaders("")
		End If

		Do While Not adRs.EOF
		    For intI = icWrk To icDir
    			oCounts.CurrentName(intI) = adRs.Fields(oCounts.Field(intI)).Value
		    Next
            
            'Place a break between Totals
			For intI = icWRK to icDIR
				If ocounts.EmployeeCount(intI) = 1 AND aDoTotals(intI + 1) AND intI >= intReqLevel Then
					Response.Write "<BR style=""FONT-SIZE:8"">"
					Response.Write "<SPAN id=lblWkr style=""LEFT:10; WIDTH:640;HEIGHT:20;FONT-SIZE:14;BORDER-COLOR:#C0C0C0; BORDER-TOP-STYLE:double; BORDER-TOP-WIDTH:3""> </SPAN>"
					Response.Write "<BR style=""FONT-SIZE:20"">"
					Exit For
				End If
			Next
		
            'Reset values
            For intI = icDIR To intReqLevel + 1 Step -1
				If oCounts.Changed(intI) AND oCounts.CurrentName(intI) <> oCounts.CurrentName(intI + 1) Then
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
			
			'Do not print repeated names
            For intI = intTempReqLevel To icSUP
				If oCounts.CurrentName(intI) <> oCounts.CurrentName(intI + 1) Then
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
			      
            For intI = icSUP To icWRK Step -1
				If oCounts.Changed(intI) AND oCounts.ShowLevel(intI) Then
					If aShowLevel(intI) Then
						If intI = intTempReqLevel Then
							strClass = "ManagementText"
							strBackColor = strWkrColor
						Else
							strClass = oCounts.HeaderClass(intI) 
							strBackColor = "#ffffff"
						End If
						'Spacer
						Response.Write "<SPAN id=lblWkr class=" & strClass & " "
						Response.Write "style=""WIDTH:630; BackGround:" & strBackColor & """>"
						Response.Write "</SPAN>"
						
						'Display the Director Name:	
						Response.Write "<SPAN id=lblDir class=" & strClass & " "
						Response.Write "style=""WIDTH:400; BackGround:" & strBackColor & """>"
						If icDIR < intReqLevel Then
							Response.Write "*&nbsp" & oCounts.NameOnly(intI) & "</SPAN>"
						Else
						    If intI = 1 Then
							    Response.Write oCounts.CurrentName(intI) & "</SPAN>"
						    Else
							    Response.Write oCounts.NameOnly(intI) & "</SPAN>"
							End If
						End If
						
						If intI > intTempReqLevel Then
							Response.Write "<BR style=""FONT-SIZE:16"">"
						End If
					End If
					
					If intI = intHeader Then
						Call WriteColumnHeaders("")
					End If
				End If
			Next  
			
			oCounts.TotalCorrect(icWRK) = adRs.Fields("TotalCorrect").Value
			oCounts.TotalError(icWRK) = adRs.Fields("TotalError").Value
			oCounts.TotalCount(icWRK) = adRs.Fields("TotalError").Value + adRs.Fields("TotalCorrect").Value
			
			adRs.MoveNext
			
			If adRs.EOF Then
				For intI = intTempReqLevel To icFIN
					If aShowLevel(intI) Or intI = intTempReqLevel Then
						aDoTotals(intI) = True
					End If
				Next
				aDoTotals(icFIN) = TRUE
				
			Else
				For intI = 0 To 5
					aDoTotals(intI) = False
				Next
				For intI = intTempReqLevel To icSUP
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
					
					If mstrTableName <> mstrCurrentName Then
						If aShowLevel(intI) Or intI = intTempReqLevel Then
							aDoTotals(intI) = True
						ElseIf intI > intTempReqLevel Then
							For intJ = intI + 1 To icDIR 
								If aShowLevel(intJ) And oCounts.CurrentName(intI) = oCounts.CurrentName(intJ) Then
									aDoTotals(intI) = True
									Exit For
								End If
							Next
						End If
					End If
				Next			
			End If
			
			'Do not print totals for levels above the lowest criteria level
			For intI = 1 To icDIR
				If aDoTotals(intI) Then
					For IntJ = intI + 1 To icFIN
						If oCounts.FormName(intI) <> "" Then
							aDoTotals(intJ) = False
						End If
					Next
				End IF
			Next
			
			If aDoTotals(icWRK) Then
				'Place Totals
				intShadeCount = intShadeCount + 1
				Call WriteTotals(icWRK, "ReportText", strWkrColor, "#000000" )
				
				Response.Write "<BR style=""FONT-SIZE:20	"">"
			End If
			
			For intI = icSUP To icSUP
				'Place totals for previous Supervisor: 
				If aDoTotals(intI) Then
					'Place totals for previous Supervisor:
					If oCounts.ShowLevel(intI - 1) Then
						strClass = oCounts.TotalClass(intI)
						strBackColor = "#ffffff"
						
						Response.Write "<small><br></small>"
						Response.Write "<span id=lblSup class=" & strClass & " "
						Response.Write "style=""LEFT:10; WIDTH:630; BACKGROUND:" & strBackColor & """></SPAN>"
						
						Response.Write "<span id=lblSup class=" & strClass & " "
						Response.Write "style=""width:280;BACKGROUND:" & strBackColor & "; TEXT-ALIGN:LEFT "">"
						If intI = 1 Then
							Response.Write oCounts.CurrentName(intI) & "&nbspTotal:" & "</Span>"
						Else
	    					Response.Write oCounts.NameOnly(intI) & "&nbspTotal:" & "</Span>"
	    		        End If
    						
					Else
						strClass = "ManagementText"
						strBackColor = strWkrColor
						intShadeCount = intShadeCount + 1
					End If
					
					Call WriteTotals(intI, strClass, strBackColor, "#000000")
					If intI = intTempReqLevel Then
						Response.Write "<BR style=""FONT-SIZE:16"">"
					ELse
						Response.Write "<BR style=""FONT-SIZE:8"">"
					End If
				End If
			Next
	Loop
	If (oCounts.FormName(icWRK) = "" Or oCounts.FormName(icSUP) = "" Or oCounts.FormName(icMGR) = "" Or oCounts.FormName(icDIR) = "") AND oCounts.EmployeeCount(icWRK) >= 2 Then
		Response.Write "<BR style=""FONT-SIZE:16"">"

		Response.Write "<SPAN id=lblAvgLabel class=DirectorTotals "
		Response.Write "style=""WIDTH:640; LEFT:10; BACKGROUND:#FFEFD5; TEXT-ALIGN:left"">"
		Response.Write "Final Total:</SPAN>"
		
				'Place Totals
		intShadeCount = intShadeCount + 1
		Call WriteTotals(icFIN, "DirectorTotals", "#FFEFD5", "#000000" )
		
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
    Call oCounts.NewRow()
    
    Call WriteNames(intWhich)
    Call WriteColumn(1,strClass,oCounts.TotalCount(intWhich),250,strBackColor,"",90,0)
    Call WriteColumn(3,strClass,oCounts.TotalError(intWhich),330,strBackColor,"",90,0)
    Call WriteColumn(2,strClass,oCounts.TotalCorrect(intWhich),495,strBackColor,"",90,0)

	Response.Write "<SPAN id=lblErrorPercent class=" & strClass & " "
	Response.Write "style=""WIDTH:95; LEFT:400; BACKGROUND:" & strBackColor & "; TEXT-ALIGN:right"">"
	Response.Write oCounts.PercentError(intWhich) & "</SPAN>"
	
	Response.Write "<SPAN id=lblPMPercentCorrect1 class=" & strClass & " "
	Response.Write "style=""WIDTH:90; LEFT:560; BACKGROUND:" & strBackColor & ";TEXT-ALIGN:right"">"
	Response.Write oCounts.PercentCorrect(intWhich) & "</SPAN>"
End Sub

Sub WriteColumnHeaders(strNameTitle)
    Call WriteColumnHeader("[BLANK]",0,0,"",strSupColor,0)
    Call WriteColumnHeader("Total",250,85,"BORDER-BOTTOM-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Number",335,85,"BORDER-BOTTOM-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Percent",415,85,"BORDER-BOTTOM-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Number",495,85,"BORDER-BOTTOM-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Percent",575,85,"BORDER-BOTTOM-STYLE:none",strSupColor,1)
    Response.Write "<BR>"
    Call WriteColumnHeader("[BLANK]",0,0,"",strSupColor,0)
    Call WriteColumnHeader(strNameTitle & "&nbspName",10,240,"BORDER-BOTTOM-STYLE:none;BORDER-TOP-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Cases",250,85,"BORDER-BOTTOM-STYLE:none;BORDER-TOP-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Incorrect",335,85,"BORDER-BOTTOM-STYLE:none;BORDER-TOP-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Incorrect",415,85,"BORDER-BOTTOM-STYLE:none;BORDER-TOP-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Correct",495,85,"BORDER-BOTTOM-STYLE:none;BORDER-TOP-STYLE:none",strSupColor,1)
    Call WriteColumnHeader("Correct",575,85,"BORDER-BOTTOM-STYLE:none;BORDER-TOP-STYLE:none",strSupColor,1)
	
    Response.Write  "<big><br></big>"
End Sub

Class clsCounters
     'Arrays hold counters.  Indexes 0-5:
    '   0 = worker,    1 = supervisor, 
    '   2 = manager,   3 = director,
    '   4 = final total
    Private aFormNames(6)	   'Form Names
    Private aCurrentNames(6)   'Current Names
    Private aPreviousNames(6)  'Previous Names
    Private aEmployeeCount(6)  'Count of Employees for each level
    Private aTotalCount(6)     'Total Count
    Private aTotalError(6)	   'Error Count
    Private aTotalCorrect(6)   'Correct Count
    Private aNameChanged(6)    'Current Differs from Last
    Private aShowLevel(6)      'Keeps track of what levels are displayed
    Private aCFieldNames(6)    'Keeps track of the level's SQL field name
    Private aCTotalClass(6)    'Keeps track of the level's style class for totals
    Private aCHeaderClass(6)   'Keeps track of the level's style class for headings
    Private intRowID
    
    Private Sub Class_Initialize()
        Dim intI
       'Initialize property holders:
        aCFieldNames(0) = "Worker"
        aCFieldNames(1) = "Supervisor"
        aCFieldNames(2) = "Manager"
        aCFieldNames(3) = "Office"
        aCFieldNames(4) = "Director"
        aCFieldNames(6) = "Total"
        
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
        aFormNames(6) = "X"
        For intI = 0 To 5
            aCurrentNames(intI) = "XXX"
            aPreviousNames(intI) = "XXX"
            aEmployeeCount(intI) = 0
            aTotalCount(intI) = 0
            aTotalError(intI) = 0
            aTotalCorrect(intI) = 0 
            aNameChanged(intI) = False
            aShowLevel(intI) = False
        Next
        intRowID = -1
    End Sub

    Public Sub NewRow()
        intRowID = intRowID + 1
    End Sub
    
    Public Property Get RowID()
        RowID = intRowID
    End Property

    Public Property Let CurrentName(intWhich, strVal)
		Dim intI
		Dim strPreviousName, strCurrentName
		
        'Move current value to holder for previous name:
        aPreviousNames(intWhich) = aCurrentNames(intWhich)
        If InStr(1, strVal, "*") > 0 Then
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
                    aTotalCount(intWhich) = 0
                    aTotalError(intWhich) = 0
                    aTotalCorrect(intWhich) = 0
                    aEmployeeCount(intWhich) = aEmployeeCount(intWhich) + 1
                    'Reset Count of Employees in lower Management levels. 
                    For intI = 0 To intWhich - 1
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
        CurrentName = aPreviousNames(intWhich)
    End Property
    
    Public Property Let TotalCount(intWhich, intVal)
        Dim intI
        
        For intI = intWhich To 5
            aTotalCount(intI) = CLng(aTotalCount(intI)) + CLng(intVal)
        Next
    End Property
    
    Public Property Get TotalCount(intWhich)
        TotalCount = aTotalCount(intWhich)
    End Property
    
    Public Property Let TotalError(intWhich, intVal)
        Dim intI
        For intI = intWhich To 5
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
        For intI = intWhich To 5
            aTotalCorrect(intI) = aTotalCorrect(intI) + dblVal
        Next
    End Property
    
    Public Property Get TotalCorrect(intWhich)
        TotalCorrect = FormatNumber(aTotalCorrect(intWhich), 0, True, False, True)
    End Property
    
    Public Property Get PercentCorrect(intWhich)
        Dim dblPercent

        If aTotalCorrect(intWhich) > 0 And aTotalCount(intWhich) > 0 Then 
			dblPercent = CDbl(aTotalCorrect(intWhich)/aTotalCount(intWhich)) * 100
		Else
			dblPercent = "0.00"
		End If

        PercentCorrect = FormatNumber(dblPercent, 2, True, False, True) & " %"
    End Property
    
    Public Property Get PercentError(intWhich)
        Dim dblPercent

        If aTotalError(intWhich) > 0 And aTotalCount(intWhich) > 0 Then 
			dblPercent = CDbl(aTotalError(intWhich)/aTotalCount(intWhich)) * 100
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
            intEnd = 5
            intStep = 1
        Else
            intEnd = 0
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