<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: Reports.asp                                                     '
'  Purpose: Screen for selecting a report to view.                          '
' Includes:                                                                 '
'   IncCnn.asp          - ADO database connection                           '
'   IncValidUser.asp    - Validates the login userid                        '
'   IncDefStyles.asp    - Contains DHTML styles common in the application.  '
'==========================================================================='
Dim mstrPageTitle   'Sets the title at the top of the form.
Dim madoRs
Dim mstrHTML
Dim mstrOptions
Dim intCnt
Dim mstrUseAuthBy
Dim mstrTestResult
Dim mstrReportName
Dim mstrRvwTypOpts
Dim mstrRvw
Dim mstrRvwClass
Dim intCountClass
Dim mintFullID
Dim strList, strRteList, strPrgList
Dim mintMaxCaseNumLen
Dim mstrResponseDueMode, mstrRespDueBasedOn
Dim mblnShowReport
Dim adRsPrg
Dim mstrAllowFutureReviewDates
Dim mstrItem, mlngTabIndex, mstrKey
Dim moRevType, mdctRevTypes
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->

<%
mstrPageTitle = "Available Reports"
mstrUseAuthBy = UCase(GetAppSetting("UseAuthBy"))
mintMaxCaseNumLen = GetAppSetting("MaxCaseNumberLength")
mstrResponseDueMode = GetAppSetting("ResponseDueMode")
mstrRespDueBasedOn = GetAppSetting("ResponseDueOwner")

Dim adCmd
Dim adRs

Set adRsPrg = Server.CreateObject("ADODB.Recordset")
Set adCmd = GetAdoCmd("spGetProgramList")
    AddParmIn adCmd, "@PrgID", adVarchar, 255, Null
    'Call ShowCmdParms(adCmdPrg) '***DEBUG
    adRsPrg.Open adCmd, , adOpenForwardOnly, adLockReadOnly
Set adCmd = Nothing
strPrgList = "|"
Do While Not adRsPrg.EOF
    strPrgList = strPrgList & adRsPrg.Fields("prgID").Value & "^" & _
        adRsPrg.Fields("prgShortTitle").Value & "^" & _
        "1^" & _
        "x^" & _
        "N|"
    adRsPrg.MoveNext
Loop

' Retreive all review types
Set adCmd = GetAdoCmd("spGetReviewTypeElms")
    AddParmIn adCmd, "@Programs", adVarChar, 100, Null
    AddParmIn adCmd, "@EffectiveDate", adDBTimeStamp, 0, Null
Set madoRs = GetAdoRs(adCmd)
Set mdctRevTypes = CreateObject("Scripting.Dictionary")
'mdctRevTypes.Add "55","55^Full 2014^0^^^^"
Do While Not madoRs.EOF
    mdctRevTypes.Add madoRs.Fields("ReviewTypeID").Value, madoRs.Fields("ReviewTypeID").Value & "^" & _
        madoRs.Fields("ReviewTypeText").Value & "^" & _
        madoRs.Fields("rteProgramID").Value & "^" & _
        madoRs.Fields("rteStartDate").Value & "^" & _
        madoRs.Fields("rteEndDate").Value & "^" & _
        madoRs.Fields("rteElementID").Value
    madoRs.MoveNext
Loop
madoRs.Close
%>

<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Option Explicit
Dim mblnCloseClicked
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>
Dim mdctPrograms
Dim mdctRevTypes
Dim mstrCboBoxDefaultText
Dim mstrCboBoxDefaultTextHTML
Dim mdctSupervisors, mdctWorkers, mdctReviewers
Dim mblnMainClosed

Sub window_onload
    Dim intDelim
    Dim intLast
    Dim strStaff
    Dim strArrayValue
    Dim intI

    Call CheckForValidUser()
    If Trim(Form.UserID.Value) = "" Then Exit Sub
    'This variable is used to set the default first item in the combobox
    'options list.  Some clients may use blank, most prefer to display <All>:
    mstrCboBoxDefaultText = "<All>"
    mstrCboBoxDefaultTextHTML = "&lt;All&gt;"
    mblnSetFocusToMain = True
    mblnMainClosed = False
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>
    
    Call SizeAndCenterWindow(767, 520, True)
    lblStatusChange.style.visibility = "visible"
    lblViewReport.style.visibility = "hidden"
    
    Set mdctRevTypes = CreateObject("Scripting.Dictionary")
    Set mdctPrograms = CreateObject("Scripting.Dictionary")
    Set mdctSupervisors = CreateObject("Scripting.Dictionary")
    Set mdctWorkers = CreateObject("Scripting.Dictionary")
    Set mdctReviewers = CreateObject("Scripting.Dictionary")
    Set mdctPrograms = LoadDictionaryObject("<% = strPrgList %>")

    Call FillStaffDictionaries()

    Call ValidateStaffInList("All")

    Call lstReports_onchange()
    Call FillControls()
    
    If lstProgram.options.length = 2 Then
        lstProgram.selectedIndex = 1
        lstProgram.disabled = True
    End If
    
    Call GetReviewTypesFromForm()
    Call GetReviewClassFromForm()

    Call ChangeBlankOptionToAll    
 <%
    For Each moRevType In mdctRevTypes
        Response.Write "mdctRevTypes.Add CLng(" & moRevType & "), """ & mdctRevTypes(moRevType) & """" & vbCrLf
    Next
%>   
    HideShowFrames("visible")
    lblStatusChange.style.visibility = "hidden"
    lblViewReport.style.visibility = "hidden"
    
    Call lstReports_onchange()
    PageBody.style.cursor = "default"
    Call CheckOptionsLength()
End Sub

Sub FillStaffDictionaries()
    <%
    'Supervisor------------------
    Set adCmd = GetAdoCmd("spGetStaffFromReviews")
        AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
        AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
        AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
        AddParmIn adCmd, "@WhichGroup", adVarchar, 2, "SW"  'Sups
        'Call ShowCmdParms(adCmdPrg) '***DEBUG
        Set adRs = GetAdoRs(adCmd)
    Set adCmd = Nothing
    Do While Not adRs.EOF
        mstrKey = adRs.Fields("PersonName").Value & " -- " & adRs.Fields("PersonNumber").Value
        mstrItem = adRs.Fields("PersonName").Value & "^SUPHOLDER^" & adRs.Fields("StartDate").Value & "^" & adRs.Fields("EndDate").Value & "^" & adRs.Fields("PersonNumber").Value
        Response.Write vbTab & "mdctSupervisors.Add """ & mstrKey & """, """ & mstrItem & """" & vbCrLf
        adRs.MoveNext
    Loop
    adRs.Close
    Set adRs = Nothing
    'Reviewer------------------
    Set adCmd = GetAdoCmd("spGetStaffFromReviews")
        AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
        AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
        AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
        AddParmIn adCmd, "@WhichGroup", adVarchar, 2, "R"  'Reviewers
        'Call ShowCmdParms(adCmdPrg) '***DEBUG
        Set adRs = GetAdoRs(adCmd)
    Set adCmd = Nothing
    Do While Not adRs.EOF
        mstrKey = adRs.Fields("PersonName").Value & " -- " & adRs.Fields("PersonNumber").Value
        mstrItem = adRs.Fields("PersonName").Value & "^SUPHOLDER^" & adRs.Fields("StartDate").Value & "^" & adRs.Fields("EndDate").Value & "^" & adRs.Fields("PersonNumber").Value
        Response.Write vbTab & "mdctReviewers.Add """ & mstrKey & """, """ & mstrItem & """" & vbCrLf
        adRs.MoveNext
    Loop
    adRs.Close
    Set adRs = Nothing
    'Worker ------------------
    Set adCmd = GetAdoCmd("spGetStaffFromReviews")
        AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
        AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
        AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
        AddParmIn adCmd, "@WhichGroup", adVarchar, 2, "W"  'Workers
        'Call ShowCmdParms(adCmdPrg) '***DEBUG
        Set adRs = GetAdoRs(adCmd)
    Set adCmd = Nothing
    Do While Not adRs.EOF
        mstrKey = adRs.Fields("PersonName").Value & " -- " & adRs.Fields("PersonNumber").Value & "^" & adRs.Fields("SupervisorNumber").Value
        mstrItem = adRs.Fields("PersonName").Value & "^" & adRs.Fields("SupervisorName").Value & " -- " & adRs.Fields("SupervisorNumber").Value & "^" & adRs.Fields("StartDate").Value & "^" & adRs.Fields("EndDate").Value & "^" & adRs.Fields("PersonNumber").Value
        Response.Write vbTab & "mdctWorkers.Add """ & mstrKey & """, """ & mstrItem & """" & vbCrLf
        adRs.MoveNext
    Loop
    adRs.Close
    Set adRs = Nothing
    %>
End Sub

<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    mblnMainClosed = True
    mblnCloseClicked = True
    mblnSetFocusToMain = False
    window.close
End Sub

Sub document_onclick()
    Select Case window.event.srcElement.id 
        Case "txtWorker", "txtSupervisor", "txtReviewer", "txtReReviewer"
        Case Else
            If window.event.srcElement.id = "lblDetail" Then
                chkDetail.checked = Not chkDetail.checked
            End If
            If window.event.srcElement.id = "lblInclude" Then
                chkInclude.checked = Not chkInclude.checked
            End If
    End Select
End Sub

Sub ChangeBlankOptionToAll()
    Dim objCtrl
    Dim oOption
    
    For Each objCtrl In Document.all
        Select Case Left(objCtrl.ID,3)
            Case "cbo","lst"
                If objCtrl.ID <> "lstProgram" Then
                    For Each oOption In objCtrl.options
                        If IsNumeric(oOption.value) Then
                            If oOption.value = 0 And (oOption.Text = "All" Or Trim(oOption.Text) = "") Then
                                oOption.Text = mstrCboBoxDefaultText
                                Exit For
                            End If
                        End If
                    Next
                End If
        End Select
    Next
End Sub

Sub HideShowFrames(strVisibility)
    Header.style.visibility = strVisibility
    PageFrame.style.visibility = strVisibility
    LeftFrame.style.visibility = strVisibility
    RightFrame.style.visibility = strVisibility
    CriteriaFrame.style.visibility = strVisibility
End Sub

'Display the description for each report
Sub lstReports_onchange()
    Dim strDescr
    Dim PrgList

    PrgList = parse(lstReports.value, ":", 3)
    Call PutReviewTypesToForm
    Call Display_Criteria()
    Call ProgramSelection(PrgList)
    If lstProgram.options.length = 2 Then
        lstProgram.selectedIndex = 1
        lstProgram.disabled = True
    Else
        lstProgram.disabled = False
    End If
    If lstReports.selectedIndex >= 0 Then
        strDescr = Trim(lstReportDescriptions.options(lstReports.selectedIndex).Text)
        If strDescr = "" Then
            strDescr = "&ltno description available for selected report&gt"
        End If
        lblDescription.innerHtml = strDescr
    End If
    lblCriteriaFrame.innerHTML = "<B>" & lstReports.Options(LstReports.selectedIndex).Text & " Criteria</B>"
    lblViewReport.innerHTML = "Building " & lstReports.Options(LstReports.selectedIndex).Text & ", please wait..."
    Call GetReviewTypesFromForm()
    Call FillElementDropDown(lstProgram.value)
End Sub

Sub ProgramSelection(PrgList)
    Dim intI
    Dim intPrg
    Dim strPrgVal
    Dim oOption
    Dim strAll
    Dim oPrg
    Dim strRecord
    Dim aGroups(10)
    Dim blnFound
    Dim intLastUsed
    
    strAll = mstrCboBoxDefaultText
    lstProgram.options.length = Null
    Set oOption = Document.createElement("OPTION")
    oOption.value = 0
    oOption.Text = strAll
    lstProgram.options.add oOption
    
    For Each oPrg In mdctPrograms
        strPrgVal = "[" & oPrg & "]"
        If InStr(PrgList, strPrgVal) > 0 Then
            strRecord = mdctPrograms(oPrg)
            If Parse(strRecord,"^",5) = "Y" Then
                ' Program is a sub-program - Check if Group name has been entered into drop down.  If
                ' it has not, enter it now
                blnFound = False
                intLastUsed = -1
                For intI = 0 To 10
                    If aGroups(intI) = Parse(strRecord,"^",4) Then
                        blnFound = True
                        Exit For
                    ElseIf aGroups(intI) <> "" Then
                        intLastUsed = intI
                    End If
                Next
                If blnFound = False Then
                    Set oOption = Document.createElement("OPTION")
                    oOption.Value = (-1)*(Parse(strRecord,"^",4))
                    oOption.Text = Parse(strRecord,"^",3)
                    lstProgram.options.add oOption
                    Set oOption = Nothing
                    aGroups(intLastUsed+1) = Parse(strRecord,"^",4)
                End If
            End If
            Set oOption = Document.createElement("OPTION")
            oOption.Value = oPrg
            If Parse(strRecord,"^",5) = "Y" Then
                oOption.Text = "  " & Parse(strRecord,"^",2)
            Else
                oOption.Text = Parse(strRecord,"^",2)
            End If
            lstProgram.options.add oOption
            Set oOption = Nothing
        End If
    Next
        
    lstProgram.value = 0
    For intI = 0 to lstProgram.options.length - 1 
        If lstProgram.options(intI).Text = Form.ProgramText.Value Then
            If lstProgram.options(intI).value = Form.ProgramID.Value Then
                lstProgram.selectedindex = intI
                Exit For
            End If
        End If
    Next
End Sub

'Double click cmd on a report causes the report to be viewed
Sub lstReports_ondblclick()
    If cmdViewReport.disabled = True Then Exit Sub
    Call cmdViewReport_onclick()
End Sub

'Load the selected report for veiwing
Sub cmdViewReport_onclick()
    Dim strFeatures
    Dim objNewWindow
    Dim lngWindowID
    
    If IsNull(lstReports.value) Or Trim(lstReports.value) = "" Then
        Exit Sub
    End If
    If Not FillForms(True) Then
        Exit Sub
    End If
    
    mblnCloseClicked = True
    Form.ReportName.value = lstReports.Options(lstReports.selectedIndex).Text
    strFeatures = "directories=no,fullscreen=no,location=no,menubar=no,status=no,resizable=yes,toolbar=no,height=480,width=750,scrollbars=yes, left=1, top=1"
    lngWindowID = Int(Timer())
    Set objNewWindow = window.open("ReportsLaunch.asp", "ReportWindow" & lngWindowID , strFeatures)
    Call window.opener.AddReportWindow(lngWindowID, objNewWindow)
End Sub

'Return to main.asp on the close button cmd
Sub cmdClose_onclick
    mblnCloseClicked = true
    Call window.opener.ManageWindows(3,"Close")
End Sub

<%'If Main has not been closed, set focus back to it.%>
Sub window_onbeforeunload()
    If mblnSetFocusToMain = True And mblnMainClosed = False Then
        window.opener.focus
    End If
End Sub

'Set Date txt box blank on button click
Sub GenDate_onkeypress(txtDate)
    If txtDate.value = "(MM/DD/YYYY)" Then
        txtDate.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub

'Refills fillers for Start Date if blank
Sub txtStartDate_onfocus()
    If Trim(txtStartDate.value) = "" Then
        txtStartDate.value = "(MM/DD/YYYY)"
    End If
    txtStartDate.select
End Sub

'Clears Start Date if left blank, Checks that a valid date was entered
Sub txtStartDate_onblur()
    Dim blnCompare
    
    If Trim(txtStartDate.value) = "(MM/DD/YYYY)" Or Trim(txtStartDate.value) = "" Then
        txtStartDate.value = ""
        Exit Sub
    End If
    
    blnCompare = True
    If Not ValidDate(txtStartDate.value) Then
        MsgBox "The Start Date must be a valid date - MM/DD/YYYY.", vbInformation, "Enter Report Criteria"
        txtStartDate.focus
        blnCompare = False
    Else
        txtStartDate.value = FormatDateTime(txtStartDate.value,2)
        Select Case Parse(lstReports.value,":",1)
            Case 76,75
                txtEndDate.value = DateAdd("m",6,txtStartDate.value)
                txtEndDate.value = DateAdd("d",-1,txtEndDate.value)
            Case 47,60,35
                txtEndDate.value = DateAdd("m",1,txtStartDate.value)
                txtEndDate.value = DateAdd("d",-1,txtEndDate.value)
        End Select
    End If
    If blnCompare And IsDate(txtEndDate.value) Then
        If CDate(txtStartDate.value) > CDate(txtEndDate.value) Then
            MsgBox "The Start Date cannot be after End Date.", vbInformation, "Enter Report Criteria"
            txtStartDate.value = ""
            txtStartDate.focus
            Exit Sub
        End If
    End If
    ValidateStaffInList("All")
    Call ListReviewTyps(lstProgram.value)
End Sub

'Refills fillers for End Date if blank
Sub txtEndDate_onfocus()
    If Trim(txtEndDate.value) = "" Then
        txtEndDate.value = "(MM/DD/YYYY)"
    End If
    txtEndDate.select
End Sub

'Clears End Date if left blank, Checks that a valid date was entered
Sub txtEndDate_onblur()
    
    If Trim(txtEndDate.value) = "(MM/DD/YYYY)" Or Trim(txtEndDate.value) = "" Then
        txtEndDate.value = ""
        Exit Sub
    End If
    
    If Not ValidDate(txtEndDate.value) Then
        MsgBox "The End Date must be a valid date - MM/DD/YYYY.", vbInformation, "Enter Report Criteria"
        txtEndDate.focus
    Else
        txtEndDate.value = FormatDateTime(txtEndDate.value,2)
        If IsDate(txtStartDate.value) Then
            If CDate(txtStartDate.value) > CDate(txtEndDate.value) Then
                MsgBox "The Start Date cannot be after End Date.", vbInformation, "Enter Report Criteria"
                txtEndDate.value = ""
                txtEndDate.focus
                Exit Sub
            End If
        End If
    End If
    ValidateStaffInList("All")
    Call ListReviewTyps(lstProgram.value)
End Sub

Sub cboProgram_onchange()
    Dim intPrg

    If IsNumeric(lstProgram.Value) Then
        If CInt(lstProgram.Value) < 0 Then
            ' A group was selected - Groups in the program drop down 
            ' are stored as negative - set GroupID to Abs of dropdown value.
            txtGroupID.value = Abs(CInt(lstProgram.Value))
        Else
            txtGroupID.value = -1
        End If
    End If
    intPrg = lstProgram.Value

    Call ListReviewTyps(intPrg)
    Call FillElementDropDown(intPrg)
End Sub

Sub FillElementDropDown(intPrg)
    Dim oOption
    Dim intI
    Dim strRecord
    Dim oDictObj, strElementList
    Dim blnInRevType
        
    cboElement.options.length = Null
    
    Set oOption = Document.createElement("OPTION")
    oOption.value = 0
    oOption.Text = "<All>"
    cboElement.options.add oOption
    If intPrg = "" Or lblReviewCount.value = "" Or lblReviewCount.value = "0" Then Exit Sub

    'Build list of elements for the selected tab and function
    If chkReviewType0.checked = True Or CInt(lblReviewCount.value) <= 1 Then
        'If Full is checked, include all elements
        strElementList = "ALL"
    Else
        strElementList = ""
        'Start at 1, since 0=Full review type
        For intI = 1 To CInt(lblReviewCount.value) - 1
            If document.all("chkReviewType" & intI).checked = True Then
                strElementList = strElementList & Parse(mdctRevTypes(CLng(document.all("chkReviewType" & intI).value)),"^",6)
            End If
        Next
    End If

    For Each oDictObj In window.opener.mdctElements
        strRecord = window.opener.mdctElements(oDictObj)
        blnInRevType = False
        If InStr(strElementList,"[" & oDictObj & "]") > 0 Or strElementList = "ALL" Then
            blnInRevType = True
        End If
        If CLng(intPrg) = CLng(Parse(strRecord,"^",4)) And blnInRevType Then
            If CheckEndDate(Parse(strRecord,"^",3)) Then
                Set oOption = Document.createElement("OPTION")
                oOption.Value = oDictObj
                oOption.Text = Parse(strRecord, "^", 1)
                cboElement.options.add oOption
                Set oOption = Nothing
            End If
        End If
    Next
    cboElement.value = 0
    For intI = 0 to cboElement.options.length - 1
        If cboElement.options(intI).Text = Form.EligElementText.value Then
            If cboElement.options(intI).value = Form.EligElementID.Value Then
                cboElement.selectedIndex = intI
                Exit For
            End If
        End If
    Next
    Call FillFactorDropDown()
End Sub

Function CheckEndDate(dtmEndDate)
    If dtmEndDate = "" Then
        CheckEndDate = True
        Exit Function
    End If
    CheckEndDate = True
    If IsDate(txtStartDate.value) Then
        If CDate(dtmEndDate) >= CDate(txtStartDate.value) Then
            CheckEndDate = True
        Else
            CheckEndDate = False
        End If
    End If
End Function

Sub cboElement_onchange()
    Call FillFactorDropDown()
End Sub

Sub FillFactorDropDown()
    Dim oOption
    Dim intI, intElementID, intFactorID
    Dim strRecord, strFactorList
    Dim oDictObj

    cboFactor.options.length = Null
    Set oOption = Document.createElement("OPTION")
    oOption.value = 0
    oOption.Text = "<All>"
    cboFactor.options.add oOption

    If cboElement.value = 0 Then Exit Sub
    intElementID = cboElement.value

    strFactorList = Parse(window.opener.mdctElements(CLng(intElementID)),"^",8)
    
    For intI = 1 To 100
        strRecord = Parse(strFactorList,"*",intI)
        If strRecord = "" Then Exit For
        If CheckEndDate(Parse(strRecord,".",2)) Then
            intFactorID = Parse(strRecord,".",1)
            Set oOption = Document.createElement("OPTION")
            oOption.Value = intFactorID
            oOption.Text = Parse(window.opener.mdctFactors(CLng(intFactorID)),"^",1)
            cboFactor.options.add oOption
            Set oOption = Nothing
        End If
    Next
End Sub

Sub ListReviewTyps(intPrg)
    Dim intI 
    Dim intCount
    Dim intPrgID
    Dim intRvwID
    Dim strRvw
    Dim strReviewType
    Dim strChecked
    Dim blnSelectAllChecked
    Dim strRecord
    Dim oRevType
    Dim dtmRteEnd, dtmRteStart, dtmParmEnd, dtmParmStart
    Dim strGroupPrgIDs
    
    If IsNumeric(intPrg) Then
        If CInt(intPrg) = 0 And Parse(lstReports.options(lstReports.selectedIndex).value,":",1) = "27" Then
            intPrg = -1
        End If
        
        If CInt(intPrg) >= 50 Or CInt(intPrg) = 6 Then
            ' For Information Remedy sub programs, do not show Review Types
            lblReviewType.style.visibility="hidden"
            divReviewTypeDefs.style.visibility="hidden"
        Else
            lblReviewType.style.visibility="visible"
            divReviewTypeDefs.style.visibility="visible"
        End If
    End If
    
    <% 'If a GroupID is passed (value <-1) then include review types for all programs in group%>
    If IsNumeric(intPrg) Then
        strGroupPrgIDs = "[" & intPrg & "]"
        If CInt(intPrg) < -1 Then
            strGroupPrgIDs = ""
            For Each oRevType In mdctPrograms
                strRecord = mdctPrograms(oRevType)
                If CInt(Parse(strRecord,"^",4)) = CInt(Abs(intPrg)) Then
                    strGroupPrgIDs = strGroupPrgIDs & "[" & Parse(strRecord,"^",1) & "]"
                End If
            Next
        End If
    End If
        
    intCount = 0
    strChecked = " "
    blnSelectAllChecked = False
    If InStr(divReviewTypeDefs.innerHTML,"chkCheckAllRT") > 0 Then
        If document.all("chkCheckAllRT").checked = True Then
            strChecked = " checked "
            blnSelectAllChecked = True
        End If
    End If
    strRvw = "<INPUT type=""checkbox"" ID=chkCheckAllRT onclick=CheckAll(2)" & strChecked & " style=""left:5"" NAME=chkCheckAllRT>"
    strRvw = strRvw & "<SPAN id=lblCheckAllRT class=DefLabel onclick=CheckAll(3) style=""WIDTH:50;left:28"">Select All</SPAN>"
    strRvw = strRvw & "<BR>"

    For Each oRevType In mdctRevTypes
        strRecord = mdctRevTypes(oRevType)
        
        intPrgID = Parse(strRecord, "^", 3)
        strReviewType = Parse(strRecord, "^", 2)
        intRvwID = Parse(strRecord, "^", 1)
        ' If intRvwID = 54 or 55 (Targeted General or Full), check if currently checked and
        ' if it is currently checked, keep it checked. 
        strChecked = " "
        If blnSelectAllChecked = False Then
            If intRvwID = 54 Or intRvwID = 55 Then
                If InStr(divReviewTypeDefs.innerHTML,"chkReviewType" & intCount) > 0 Then
                    If document.all("chkReviewType" & intCount).checked = True Then strChecked = " checked "
                End If
            End If
        Else
            ' If Select all was checked, check all subsequent RTs
            strChecked = " checked "
        End If
        dtmRteEnd = Parse(strRecord, "^", 5)
        If dtmRteEnd = "" Then dtmRteEnd = "12/31/3100"
        dtmParmEnd = txtEndDate.value
        If dtmParmEnd = "" Then dtmParmEnd = "12/31/3100"
        dtmParmStart = txtStartDate.value
        If dtmParmStart = "" Then dtmParmStart = "01/01/1900"
        dtmRteStart = Parse(strRecord, "^", 4)
        If dtmRteStart = "" Then dtmRteStart = "01/01/1900"
        If (InStr(strGroupPrgIDs,"[" & intPrgID & "]") > 0 Or intPrg = -1 Or intPrgID = "0") _
            And (CDate(dtmRteStart)<= CDate(dtmParmEnd)) _
            And (CDate(dtmRteEnd) >= CDate(dtmParmStart)) Then
            
            strRvw = strRvw & "<INPUT TYPE=checkbox ID=chkReviewType" & intCount & " VALUE=" & intRvwID & strChecked & "STYLE=""LEFT:5"" tabIndex=11 onclick="" ChkReviewType_onclick(" & intCount & ")"">"
            strRvw = strRvw & "<SPAN id=lblChkReviewType" & intCount & " onclick=""lblChkReviewType_onclick(" & intCount & ")"" STYLE=""LEFT:30;WIDTH:250"">" & strReviewType & "</SPAN><BR>" & vbCrLf 
            intCount = intCount + 1
        End If
    Next
    
    divReviewTypeDefs.innerHTML = strRvw
    lblReviewCount.value = intCount
End Sub

Sub cboSupervisor_onchange()
    Call ValidateStaffInList("SW")
End Sub

Sub cboWorker_onchange()
End Sub

Sub ValidateStaffInList(strType)
    Dim oKey, strRecord
    Dim dtmReportStartDate, dtmReportEndDate
    Dim intI, strSelected
    Dim dctOptions, oOption
    Dim strComboValue

    dtmReportStartDate = txtStartDate.value
    If dtmReportStartDate = "" Then dtmReportStartDate = "01/01/2000"
    dtmReportEndDate = txtEndDate.value
    If dtmReportEndDate = "" Then dtmReportEndDate = "12/31/2100"

    Set dctOptions = CreateObject("Scripting.Dictionary")
    
    If strType = "SW" Or strType = "All" Then
        '--Update Supervisor combo
        strSelected = cboSupervisor.value
        If strSelected = "" Then strSelected = "0"
        For Each oKey In mdctSupervisors
            strRecord = mdctSupervisors(oKey)
            If CDate(Parse(strRecord,"^",3)) <= CDate(dtmReportEndDate) And CDate(Parse(strRecord,"^",4)) >= CDate(dtmReportStartDate) Then
                dctOptions.Add oKey, oKey
            End If
        Next
        
        cboSupervisor.options.length = Null
        Set oOption = Document.createElement("OPTION")
        oOption.value = 0
        oOption.Text = mstrCboBoxDefaultText
        cboSupervisor.options.add oOption
        Set oOption = Nothing
        For Each oKey In dctOptions
            Set oOption = Document.createElement("OPTION")
            oOption.value = oKey
            oOption.Text = dctOptions(oKey)
            cboSupervisor.options.add oOption
            Set oOption = Nothing
        Next
        cboSupervisor.value = strSelected
    End If

    dctOptions.RemoveAll
    If strType = "R" Or strType = "All" Then
        '--Update Reviewer combo
        strSelected = cboReviewer.value
        If strSelected = "" Then strSelected = "0"
        For Each oKey In mdctReviewers
            strRecord = mdctReviewers(oKey)
            If CDate(Parse(strRecord,"^",3)) <= CDate(dtmReportEndDate) And CDate(Parse(strRecord,"^",4)) >= CDate(dtmReportStartDate) Then
                dctOptions.Add oKey, oKey
            End If
        Next
        
        cboReviewer.options.length = Null
        Set oOption = Document.createElement("OPTION")
        oOption.value = 0
        oOption.Text = mstrCboBoxDefaultText
        cboReviewer.options.add oOption
        Set oOption = Nothing
        For Each oKey In dctOptions
            Set oOption = Document.createElement("OPTION")
            oOption.value = oKey
            oOption.Text = dctOptions(oKey)
            cboReviewer.options.add oOption
            Set oOption = Nothing
        Next
        cboReviewer.value = strSelected
    End If

    If strType = "W" Or strType = "SW" Or strType = "All" Then
        '--Update Worker combo
        dctOptions.RemoveAll
        For Each oKey In mdctWorkers
            strRecord = mdctWorkers(oKey)
            strComboValue = Parse(oKey,"^",1)
            If (Parse(strRecord,"^",2) = cboSupervisor.value Or cboSupervisor.value = "0") And _
                CDate(Parse(strRecord,"^",3)) <= CDate(dtmReportEndDate) And CDate(Parse(strRecord,"^",4)) >= CDate(dtmReportStartDate) Then
                If Not dctOptions.Exists(strComboValue) Then
                    dctOptions.Add strComboValue, strComboValue
                End If
            End If
        Next
        
        cboWorker.options.length = Null
        Set oOption = Document.createElement("OPTION")
        oOption.value = 0
        oOption.Text = mstrCboBoxDefaultText
        cboWorker.options.add oOption
        Set oOption = Nothing
        For Each oKey In dctOptions
            Set oOption = Document.createElement("OPTION")
            oOption.value = oKey
            oOption.Text = dctOptions(oKey)
            cboWorker.options.add oOption
            Set oOption = Nothing
        Next
    End If
End Sub

Sub CheckOptionsLength()
    If cboSupervisor.options.length = 2 Then
        If "<%=glngAliasPosID%>" = "2" Or "<%=gblnUserAdmin%>" = "True" Then
            cboSupervisor.selectedIndex = 0
            cboSupervisor.disabled=False
        Else
            cboSupervisor.selectedIndex = 1
            cboSupervisor.disabled=True
        End If
    End If
    If cboWorker.options.length = 2 Then
        cboWorker.selectedIndex = 1
        cboWorker.disabled = True
    Else
        cboWorker.selectedIndex = 0
        cboWorker.disabled = False
    End If
    If cboReviewer.options.length = 2 Then
        cboReviewer.selectedIndex = 1
        cboReviewer.disabled = True
    Else
        cboReviewer.selectedIndex = 0
        cboReviewer.disabled = False
    End If
End Sub

Sub CheckAll(intWho)
    Dim intI
    
    If intWho <= 1 Then
        If intWho = 1 Then chkCheckAllRC.checked = Not chkCheckAllRC.checked
        For intI = 0 To lblReviewClassCount.innerText - 1
                If chkCheckAllRC.checked = True Then
                    document.all("chkReviewClass" & intI).checked = True
                Else
                    document.all("chkReviewClass" & intI).checked = False
                End If
        Next
    Else
        If intWho = 3 Then chkCheckAllRT.checked = Not chkCheckAllRT.checked
        For intI = 0 To lblReviewCount.innerText - 1
                If chkCheckAllRT.checked = True Then
                    document.all("chkReviewType" & intI).checked = True
                Else
                    document.all("chkReviewType" & intI).checked = False
                End If
        Next
        Call FillElementDropDown(lstProgram.value)
    End If
End Sub

Sub cboResponse_onchange()
    If cboResponse.value = 1 Then
        txtDaysPastDue.style.visibility = "visible"
        lblDaysPastDue.style.visibility = "visible"
        lblDaysPastDue.innerText = "Minimum Days Past Due"
    ElseIf cboResponse.value = 2 Then
        txtDaysPastDue.style.visibility = "visible"
        lblDaysPastDue.style.visibility = "visible"
        lblDaysPastDue.innerText = "Minimum Days Pending"
    Else
        txtDaysPastDue.style.visibility = "hidden"
        lblDaysPastDue.style.visibility = "hidden"
    End If
End Sub

Sub ReviewMonthYear_onfocus(ctlTextBox)
    If Trim(ctlTextBox.value) = "" Then
        ctlTextBox.value = "(MM/YYYY)"
    End If
    ctlTextBox.select
End Sub

Sub ReviewMonthYear_onkeypress(ctlTextBox)
    If ctlTextBox.value = "(MM/YYYY)" Then
        ctlTextBox.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub

Sub ReviewMonthYear_onblur(ctlTextBox)
    Dim strMonth
    Dim strYear
    Dim intPos
    Dim blnErr
    Dim blnFutureDate
    Dim dtmReviewMonth, dtmEndReviewMonth
    
    If Trim(ctlTextBox.value) = "" Then
        Exit Sub
    End If
    If Trim(ctlTextBox.value) = "(MM/YYYY)" Then
        ctlTextBox.value = ""
        Exit Sub
    End If
    
    intPos = Instr(ctlTextBox.value, "/")
    If intPos = 0 Then
        blnErr = True
    Else
        strMonth = Trim(Mid(ctlTextBox.value, 1, intPos -1))
        strYear = Trim(Mid(ctlTextBox.value, intPos + 1))
        
        If Not IsNumeric(strMonth) Or Not IsNumeric(strYear) Then
            blnErr = True
        ElseIf Len(strMonth) > 2 Or Len(strYear) > 4 Or Len(strYear) = 3 Then
            blnErr = True
        ElseIf Cint(strMonth) = 0 Or CInt(strMonth) > 12 Then
            blnErr = True
        ElseIf Len(strYear) > 2 And CInt(strYear) < 1601 Then
            blnErr = True
        Else
            If Len(strMonth) < 2 Then
                strMonth = "0" & strMonth
            End If
            If Len(strYear) <= 2 Then
                strYear = 2000 + CInt(strYear) 
            End If
            ctlTextBox.value = strMonth & "/" & strYear

            ' Do not allow Month/Year in the future
            blnFutureDate = False
            If "<% = mstrAllowFutureReviewDates %>" <> "Yes" Then
                dtmReviewMonth = strMonth & "/01/" & strYear
                If CDate(FormatDateTime(dtmReviewMonth,2)) > CDate(FormatDateTime(Now(),2)) Then blnFutureDate = True
            End If
        End If
    End If

    If blnErr Then
        MsgBox "Review Month Year must be in the format MM/YYYY.", vbInformation, "View Report"
        ctlTextBox.focus
    End If
    
    If blnFutureDate Then
        MsgBox "Review Month Year cannot be after " & Right("00" & Month(Now()),2) & "/" & Year(Now()) & ".", vbInformation, "View Report"
        ctlTextBox.focus
    End If
    If Not blnErr And Not blnFutureDate And ctlTextBox.ID = "txtStartReviewMonth" Then
        ' If end review month is blank, default to same month/year as start
        If txtEndReviewMonth.value = "" Or txtEndReviewMonth.value = "(MM/YYYY)" Then
            txtEndReviewMonth.value = txtStartReviewMonth.value
        End If
    End If
End Sub

Sub NavigateFix(strAction)
    If strAction = "Open" Then
        lblSelectReport.style.top = 40
        lstReports.style.top = 55
        lstReports.style.height = 145
    Else
        lblSelectReport.style.top = 5
        lstReports.style.top = 20
        lstReports.style.height = 180
    End If
End Sub
-->
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style="cursor:wait">
   <SPAN id=lblStatusChange class=DefLabel style="LEFT:15; TOP:80; visibility:visible; cursor:wait">
        Building report criteria form, please wait...
    </SPAN>
    <SPAN id=lblViewReport class=DefLabel style="LEFT:15; TOP:80; visibility:hidden; cursor:wait">
        Building report, please wait...
    </SPAN>
    <DIV id=blankFrame class=DefPageFrame style="HEIGHT:400; WIDTH:737; TOP:51; visibility:hidden; cursor:auto"></div>   
    <%
    Dim strAuthLabel
    If Instr(ReqForm("ProgramsSelected"), "4") > 0 Then 
        strAuthLabel = "Causer"
    Else
        strAuthLabel = "Benefit Authorized By"
    End If
    %>
    
    <DIV id=Header class=DefTitleArea style="WIDTH:737;visibility:hidden">
        <SPAN id=lblAppTitle class=DefTitleText style="WIDTH:737">
            <%=mstrPageTitle%>
        </SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>

    <% Call WriteNavigateControls(3,30,gstrBackColor) %>

    <DIV id=PageFrame class=DefPageFrame style="HEIGHT:47; WIDTH:737; TOP:51; border-bottom-style:none; visibility:hidden">
       <DIV id=divDateRange class=DefPageFrame style="LEFT:10 ; HEIGHT:20; WIDTH:360; TOP:10; border:none">
            <SPAN id=lblStartDate class=DefLabel style="LEFT:0; WIDTH:56">
                Start Date
            </SPAN>
            <INPUT id=txtStartDate type=text title="Start Date for reporting range" 
                style="LEFT:55; WIDTH:80; background:<%=gstrCtrlBackColor%>"
                onkeypress="GenDate_onkeypress(txtStartDate)"
                tabIndex=11 maxlength=10 NAME="txtStartDate">

            <SPAN id=lblEndDate class=DefLabel style="LEFT:150; WIDTH:50">
                End Date
            </SPAN>
            <INPUT id=txtEndDate type=text title="End Date for reporting range" 
                style="LEFT:200; WIDTH:80; background:<%=gstrCtrlBackColor%>" 
                onkeypress="GenDate_onkeypress(txtEndDate)"
                tabIndex=11 maxlength=10 NAME="txtEndDate">
        </DIV>
     </DIV>

     <DIV id=LeftFrame class=DefPageFrame style="HEIGHT:385; WIDTH:280; TOP:90; visibility:hidden"> 
        <SPAN id="lblSelectReport" class=DefLabel style="LEFT:10;WIDTH:120;TOP:5">
            Select a report:
        </SPAN> 
        <!--Select box for selecting specific report to be generated-->
        <SELECT id=lstReports TYPE="select-one"
            style="LEFT:10; WIDTH:255; TOP:20; HEIGHT:180; BACKGROUND-COLOR: <%=gstrPageColor%>" 
            tabIndex=11 size=2 NAME="lstReports">
            <% 
            mstrOptions = ""
            Set gadoCmd = GetAdoCmd("spGetReportList")
                AddParmIn gadoCmd, "@PrgID", adVarchar, 255, Null 'ReqForm("ProgramsSelected")
                'Call ShowCmdParms(gadoCmd) '***DEBUG
                Set madoRs = GetAdoRs(gadoCmd)
            Do While Not madoRs.EOF
                Select Case madoRs.Fields("rptRecordID").Value
                    Case 72, 73
                        'These reports are only visible if the user has 
                        'the Quality Assurance security role, role 21:
                        '72 - CAR Elig Elem Summary
                        '73 - CAR Accuracy Summary
                        If Instr(gstrRoles, "[1]") > 0 Then
                            mblnShowReport = True
                        Else
                            mblnShowReport = False
                        End If
                    Case 118, 139, 140 
                        ' Archive reports
                        mblnShowReport = False
                    Case Else
                        mblnShowReport = True
                End Select
                If Instr(madoRs.Fields("rptProgramList").Value, "[999]") <= 0 Then '[999] is a disabled report.
                    If mblnShowReport Then
                        If mstrOptions = "" Then
                            'Have the first report selected automatically at load time:
                            Response.Write "<OPTION VALUE=" & madoRs.Fields("rptRecordID").Value & ":" & madoRs.Fields("rptReportSource").Value & ":" & madoRs.Fields("rptProgramList").Value & " SELECTED>" & madoRs.Fields("rptReportTitle").Value
                            mstrOptions = mstrOptions & "<OPTION VALUE=" & madoRs.Fields("rptRecordID").Value & ":" & madoRs.Fields("rptReportSource").Value & ":" & madoRs.Fields("rptProgramList").Value & " SELECTED>" & madoRs.Fields("rptDescription").Value
                        Else
                            Response.Write "<OPTION VALUE=" & madoRs.Fields("rptRecordID").Value & ":" & madoRs.Fields("rptReportSource").Value & ":" & madoRs.Fields("rptProgramList").Value & ">" & madoRs.Fields("rptReportTitle").Value
                            mstrOptions = mstrOptions & "<OPTION VALUE=" & madoRs.Fields("rptRecordID").Value & ":" & madoRs.Fields("rptReportSource").Value & ":" & madoRs.Fields("rptProgramList").Value & ">" & madoRs.Fields("rptDescription").Value
                        End If
                    End If
                End If
                madoRs.MoveNext 
            Loop 
            madoRs.Close
            Set madoRs = Nothing
            Set gadoCmd = Nothing
            %>
        </SELECT>
        
        <SELECT id=lstReportDescriptions
            style="LEFT:0; WIDTH:0; TOP:0; VISIBILITY:hidden" 
            tabIndex=0 NAME="lstReportDescriptions">
            <%
            Response.Write mstrOptions
            mstrOptions = ""
            %>
        </SELECT>

        <DIV id=divReportDescription class=DefPageFrame style="LEFT:10 ; HEIGHT:120; WIDTH:255; TOP:210">
            <SPAN id=lblCapDescription class=DefLabel
                style="Left:10; TOP:-7; WIDTH:100; background-color:<%=gstrBackColor%>; TEXT-ALIGN:center">
                Report Description
            </SPAN>
            <SPAN id=lblDescription class=DefLabel
                style="Left:10; TOP:15; WIDTH:240; HEIGHT:92; overflow:auto">
            </SPAN>
        </DIV>

        <BUTTON id=cmdClearCriteria class=DefBUTTON title="Clear the report criteria" 
            style="LEFT:8; TOP:345; WIDTH:80"
            onclick="cmdClearCriteria_onclick()"
            accessKey=V tabIndex=20>
            <U>C</U>lear
            </BUTTON>

        <BUTTON id=cmdViewReport class=DefBUTTON title="View the selected report" 
            style="LEFT:97; TOP:345; WIDTH:80"
            accessKey=V tabIndex=20>
            <U>V</U>iew Report
        </BUTTON>
        
        <BUTTON id=cmdClose class=DefBUTTON title="Return to main screen" 
            style="LEFT:185; TOP:345; WIDTH:80"
            accessKey=C tabIndex=20>
            <U>C</U>lose
        </BUTTON>
        
    </DIV>
    <DIV id=RightFrame class=DefPageFrame style="LEFT:289;HEIGHT:385; WIDTH:458;TOP:90;visibility:hidden">
        <SPAN id=lblCriteriaFrame class=DefLabel style="LEFT:10;WIDTH:300;TOP:5">
        </SPAN> 
            <!-- IFRAME used for searching for staff----------------------------->
        <DIV ID=CriteriaFrame class=DefPageFrame 
            style="Left:5; HEIGHT:340; Width:448;Top:20; border-style:none; visibility:hidden">

            <SPAN id=lblDirector class=DefLabel style="LEFT:-1010; WIDTH:200; TOP:5">
                <%=gstrDirTitle%>
            </SPAN>
            <SELECT id=cboDirector style="LEFT:-1010; WIDTH:200; TOP:20" tabIndex=-1 NAME="cboDirector">
            </SELECT>

            <SPAN id=lblOffice class=DefLabel style="LEFT:-1010; WIDTH:200; TOP:50">
                <%=gstrOffTitle%>
            </SPAN>
            <SELECT id=cboOffice style="LEFT:-1010; WIDTH:200; TOP:65" tabIndex=-1 NAME="cboOffice">
            </SELECT>
            
            <SPAN id=lblManager class=DefLabel style="LEFT:-1010; WIDTH:200; TOP:95">
                <%=gstrMgrTitle%>
            </SPAN>
            <SELECT id=cboManager style="LEFT:-1010; WIDTH:200; TOP:110" tabIndex=-1 NAME="cboManager">
            </SELECT>
            
            <SPAN id=lblSupervisor class=DefLabel style="LEFT:10; WIDTH:200; TOP:5">
                <%=gstrSupTitle%>
            </SPAN>
            <SELECT id=cboSupervisor style="LEFT:10; WIDTH:200; TOP:20" tabIndex=<%=GetTabIndex%> NAME="cboSupervisor">
            </SELECT>
            
            <SPAN id=lblReviewer class=DefLabel style="LEFT:10; WIDTH:200; TOP:5">
                <%=gstrRvwTitle%>
            </SPAN>
            <SELECT id=cboReviewer style="LEFT:10; WIDTH:200; TOP:20" tabIndex=<%=GetTabIndex%> NAME="cboReviewer">
            </SELECT>
            
            <SPAN id=lblWorker class=DefLabel style="LEFT:10; WIDTH:200; TOP:50">
                <%=gstrWkrTitle%>
            </SPAN>
            <SELECT id=cboWorker style="LEFT:10; WIDTH:200; TOP:65" tabIndex=<%=GetTabIndex%> NAME="cboWorker">
            </SELECT>

            <SPAN id=lblReReviewer class=DefLabel style="LEFT:10; WIDTH:200; TOP:95">
                <%=gstrEvaTitle%>
            </SPAN>
            <INPUT type="text" ID=txtReReviewer NAME="txtReReviewer" 
                onkeydown="Gen_onkeydown(txtReReviewer)"  tabIndex=<%=GetTabIndex%>
                onfocus="CmnTxt_onfocus(txtReReviewer)" maxlength=50
                style="LEFT:10; WIDTH:200; TOP:110;text-align:left">
            <input type=hidden id=txtReReviewerStartDate>
            <input type=hidden id=txtReReviewerEndDate>

                
            <DIV id=divReviewMonth class=DefPageFrame 
                style="Left:222;HEIGHT:50;Width:215;Top:270;border-style:none;visibility:hidden;background-color:transparent;z-index:1000">
                <SPAN id=lblReviewMonth class=DefLabel style="LEFT:12; WIDTH:200; TOP:0">
                    Review Month
                </SPAN>
                <SPAN id=lblStartReviewMonth class=DefLabel style="LEFT:20; WIDTH:70; TOP:15">
                    From
                </SPAN>
                <INPUT type=text id=txtStartReviewMonth title="Starting Review Month/Year" style="LEFT:20; WIDTH:70; TOP: 30"
                    maxlength=7 NAME="txtStartReviewMonth" 
                    onblur="ReviewMonthYear_onblur(txtStartReviewMonth)"
                    onfocus="ReviewMonthYear_onfocus(txtStartReviewMonth)"
                    onkeypress="ReviewMonthYear_onkeypress(txtStartReviewMonth)">
                <SPAN id=lblEndReviewMonth class=DefLabel style="LEFT:100; WIDTH:70; TOP:15">
                    To
                </SPAN>
                <INPUT type=text id=txtEndReviewMonth title="Ending Review Month/Year" style="LEFT:100; WIDTH:70; TOP: 30"
                    maxlength=7 NAME="txtEndReviewMonth" 
                    onblur="ReviewMonthYear_onblur(txtEndReviewMonth)"
                    onfocus="ReviewMonthYear_onfocus(txtEndReviewMonth)"
                    onkeypress="ReviewMonthYear_onkeypress(txtEndReviewMonth)">
            </DIV>

            <DIV id=divReviewType  class=DefPageFrame 
                style="LEFT:10; WIDTH:425; Top:230; HEIGHT:120;BORDER: none CURSOR:default">
                    
                <% Call WriteReviewClass() %>
                
                <SPAN id=lblReviewClass
                    class=DefLabel
                    style="WIDTH:80">
                    <% = gstrReviewClassTitle %>
                </SPAN>
                
                <DIV id=divReviewClass
                    style="FONT-SIZE: 8pt; 
                            FONT-FAMILY: <%=gstrTextFont%>;
                            LINE-HEIGHT: 12pt;
                            PADDING-TOP: 2px;
                            PADDING-BOTTOM: 2px;
                            LEFT:0; 
                            WIDTH:200; 
                            TOP:20;
                            HEIGHT:100;
                            Overflow:auto;
                            BORDER-STYLE:solid; 
                            Border-Width:1; 
                            Border-Color:7f9db9; 
                            BACKGROUND-COLOR: <%=gstrPageColor%>;
                            CURSOR:default">
                    <INPUT type="checkbox" ID=chkCheckAllRC onclick=CheckAll(0) style="top:3;left:5" NAME=chkCheckAllRC>
                    <SPAN id=lblCheckAllRC class=DefLabel onclick=CheckAll(1) style="top:3;WIDTH:50;left:28">Select All</SPAN>
                    <% = mstrRvwClass %>
                </DIV>
                
                <% Call WriteReviewControls() %>
                
                <SPAN id=lblReviewType
                    class=DefLabel
                    style="LEFT:225; WIDTH:100">
                    Review Type
                </SPAN>
                <TEXTAREA id=lblReviewCount style="visibility:hidden" NAME="lblReviewCount"></textarea>
                <DIV id=divReviewTypeDefs
                    style="FONT-SIZE: 8pt; 
                            FONT-FAMILY: <%=gstrTextFont%>;
                            LINE-HEIGHT: 12pt;
                            PADDING-TOP: 2px;
                            PADDING-BOTTOM: 2px;
                            LEFT:225; 
                            WIDTH:200; 
                            TOP:20;
                            HEIGHT:100;
                            Overflow:auto; 
                            BORDER-STYLE:solid; 
                            Border-Width:1; 
                            Border-Color:7f9db9; 
                            BACKGROUND-COLOR: <%=gstrPageColor%>;
                            CURSOR:default">
                    
                </DIV>
                                
            </DIV>    
            <SPAN id=lblSubmitted class=DefLabel style="LEFT:235; TOP:5">
                Show Unsubmitted
            </SPAN>
            <SELECT ID=cboSubmitted 
                style="z-index:-1; Left:235; WIDTH:200; Top:20; OVERFLOW:auto; visibility:Hidden" 
                tabindex=<%=GetTabIndex%> NAME="cboSubmitted">
                <OPTION value=0>All</OPTION>
                <OPTION value=1>No Supervisor Signature</OPTION>
                <OPTION value=2>No Worker Acknowledgement</OPTION>
                <OPTION value=3>Not Submitted To Reports</OPTION>
            </SELECT>
            
            <SPAN id=lblResponse class=DefLabel style="LEFT:235; WIDTH:200; Top:5">
                Response
            </SPAN>
            <SELECT ID=cboResponse 
                style="z-index:-1; Left:235; WIDTH:200; Top:20; OVERFLOW:auto; visibility:Hidden" 
                tabindex=<%=GetTabIndex%> NAME="cboResponse">
                <OPTION value=0>All</OPTION>
                <OPTION value=1>Worker Past Due</OPTION>
                <OPTION value=2>Worker Pending</OPTION>
                <OPTION value=3>Supervisor Pending</OPTION>
            </SELECT>
            <SPAN id=lblDaysPastDue class=DefLabel style="LEFT:235; WIDTH:200; Top:50">
                Days Past Due
            </SPAN>
            <INPUT type=text id=txtDaysPastDue style="LEFT:235;TOP:65;WIDTH:40"
                onkeypress="Call TextBoxOnKeyPress(window.event.keyCode,'N')"
                tabIndex=<%=GetTabIndex%> maxlength=3 NAME="txtDaysPastDue">
            
            <SPAN id=lblProgram class=DefLabel style="LEFT:235; WIDTH:200; TOP:5">
                Program
            </SPAN>
            <INPUT type="hidden" id=txtGroupID NAME="txtGroupID">
            <SELECT id=lstProgram 
                style="z-index:-1; LEFT:235; WIDTH:200; Top:20; OVERFLOW:auto; visibility:visible"
                onchange="cboProgram_onchange"
                tabindex=<%=GetTabIndex%> NAME="lstProgram">
            </SELECT>
            
            <SPAN id=lblEligElement 
                class=DefLabel
                Style="LEFT:235; TOP:50; Width:200">
                Element
            </SPAN>
            <SELECT id=cboElement title="Screen"
                style="z-index:-1; LEFT:235; TOP:65; WIDTH:200; OVERFLOWL:auto; visibility:hidden"
                onchange=""
                tabIndex=<%=GetTabIndex%> NAME="cboElement">
            </SELECT>
            
            <SPAN id=lblFactor 
                class=DefLabel
                Style="LEFT:235; TOP:95; Width:200">
                Causal Factor
            </SPAN>
            <SELECT id=cboFactor 
                style="z-index:-1; LEFT:235; TOP:110; WIDTH:200; OVERFLOWL:auto; visibility:hidden"
                onchange=""
                tabIndex=<%=GetTabIndex%> NAME="cboFactor">
            </SELECT>

            <SPAN id=lblCaseAction
                class=DefLabel
                style="Left:-1235; Top:50; Width:200">
                Case Action
            </SPAN>
            <SELECT id=cboCaseAction
                style="z-index:-1; LEFT:-1235; TOP:65; Width:200; OVERFLOW:auto; visibility:hidden"    
                tabIndex=<%=GetTabIndex%> NAME="cboCaseAction">
                <OPTION value=0 selected></OPTION>
                <%=BuildList("CaseAction","",0,0,0)%>
            </SELECT>
            
            <SPAN id=lblCaseNumber class=DefLabel style="LEFT:235;TOP:50">
                Case Number
            </SPAN>
            <INPUT type=text id=txtCaseNumber style="LEFT:235;TOP:65;WIDTH:200"
                tabIndex=11 maxlength=<%=mintMaxCaseNumLen%> NAME="txtCaseNumber">
            
            <SPAN id=lblMinDays class=DefLabel style="LEFT:235; TOP:50; WIDTH:200">
                Minimum Processing Days
            </SPAN>
            <TEXTAREA id=txtMinDays style="Left:235;TOP:65;WIDTH:200" tabIndex=11 NAME="txtMinDays"></TEXTAREA>
            
             <SPAN id=lblDetail class=DefLabel style="cursor:hand;LEFT:235; WIDTH:200; TOP:185">
                    Show Detail
                    <INPUT type=checkbox id=chkDetail
                        style="LEFT:160; WIDTH:20; HEIGHT:20; TOP:-2" tabIndex=5 NAME="chkDetail">
            </SPAN>
             <SPAN id=lblInclude class=DefLabel style="cursor:hand;LEFT:235; WIDTH:200; TOP:185">
                    Include Correct Cases
                    <INPUT type=checkbox id=chkInclude
                        style="LEFT:160; WIDTH:20; HEIGHT:20; TOP:-2" tabIndex=5 NAME="chkInclude">
            </SPAN>
            <SPAN id=lblSupVsReviewer class=DefLabel style="LEFT:235; WIDTH:120; TOP:170">
                Based on
            </SPAN>
            
            <DIV id=divSupVsReviewer class=DefPageFrame style="Left:235;TOP:185;WIDTH:200;height:60;BACKGROUND-COLOR: <%=gstrPageColor%>" >
                <INPUT type="radio" ID=optSupervisor onclick=SupVsReviewer(-1) style="LEFT:5;TOP:5" NAME="SupVsRevGroup" VALUE="optSupervisor"><BR>
                <SPAN id=lblRadioSupervisor class=DefLabel onclick=SupVsReviewer(0)
                    style="LEFT:25; WIDTH:120; TOP:5">
                    Supervisor of <%=gstrWkrTitle%>
                </SPAN>
                <INPUT type="radio" ID=optReviewer onclick=SupVsReviewer(-1) style="LEFT:5;TOP:25" NAME="SupVsRevGroup" VALUE="optReviewer">
                <SPAN id=lblRadioReviewer class=DefLabel onclick=SupVsReviewer(1)
                    style="LEFT:25; WIDTH:120; TOP:26">
                    Reviewer
                </SPAN>
            </DIV>
        </DIV>
    </DIV>
</BODY>
<%
Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY: hidden"" ACTION=""CausalFactorSummaryParms.asp"" ID=Form>" & vbCrLf
    Call CommonFormFields()
    Call ReportFormDef()
    Call WriteFormField("onchangeFlag", "Y")
    Call WriteFormField("ReportName", "")
    Call WriteFormField("TabID", "")
    Call WriteFormField("TabName", "")
    Call WriteFormField("FactorID", "")
    Call WriteFormField("FactorText", "")
    Call WriteFormField("ReportType", "Reports.asp")
Response.Write "</FORM>"
%>
</HTML>
<SCRIPT LANGUAGE=vbscript>
<!--

'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: CrtAppProcessing.asp                                            '
'  Purpose: The report criteria selection screen for the Application        '
'           Processing report (medicaid).                                   '
'==========================================================================='
'Hides all criteria
Sub Hide_Criteria()
    lblWorker.style.top = 50
    cboWorker.style.top = 65
        
    lblSupervisor.style.visibility="hidden"
    cboSupervisor.style.visibility="hidden"
    lblWorker.style.visibility="hidden"
    cboWorker.style.visibility="hidden"
    lblReviewer.style.visibility="hidden"
    cboReviewer.style.visibility="hidden"
    lblReReviewer.style.visibility="hidden"
    txtReReviewer.style.visibility="hidden"
    
    lblProgram.style.visibility="hidden"
    lstProgram.style.visibility="hidden"
    lblEligElement.style.visibility="hidden"
    cboElement.style.visibility="hidden"
    divSupVsReviewer.style.left = -1000
    lblSupVsReviewer.style.left = -1000
    lblFactor.style.top = 95
    cboFactor.style.top = 110
    lblFactor.style.visibility="hidden"
    cboFactor.style.visibility="hidden"
        
    lblCaseAction.style.visibility="hidden"
    cboCaseAction.style.visibility="hidden"
    lblSubmitted.style.visibility="hidden"
    cboSubmitted.style.visibility="hidden"
    lblCaseNumber.style.visibility="hidden"
    txtCaseNumber.style.visibility="hidden"
    lblDetail.style.visibility="hidden"
    lblInclude.style.visibility="hidden"
    lblMinDays.style.visibility="hidden"
    txtMinDays.style.visibility="hidden"
    lblResponse.style.visibility="hidden"
    cboResponse.style.visibility="hidden"
    lblDaysPastDue.style.visibility="hidden"
    txtDaysPastDue.style.visibility="hidden"
    
    divReviewType.style.visibility="hidden"
    lblReviewType.style.visibility="hidden"
    divReviewTypeDefs.style.visibility="hidden"
    lblReviewer.innerText="Reviewer"
    divReviewMonth.style.visibility = "hidden"
    lblDetail.style.visibility = "hidden"
End Sub

'Display the appropriate Criteria for the selected Report
Sub Display_Criteria()
    Dim strReport
    Dim intI
    Dim intLength
    
    strReport = Parse(lstReports.value, ":", 1)
    Call Hide_Criteria()
    
    divReviewMonth.style.visibility = "visible"
    lblProgram.style.visibility="visible"
    lstProgram.style.visibility="visible"
    
	divReviewType.style.left = 10
    chkDetail.checked = False
    Select Case strReport
        Case 26,57 ' Case Accuracy Summary
            Call SetMgrSupWkrAbyRT()
            divReviewMonth.style.top = 55
            If lstProgram.selectedIndex = 0 Then
                lstProgram.selectedIndex = 1
            End If
            Call cboProgram_onchange
        Case 27,127, 138 'Case Review Detail
            Call SetMgrSupWkrAbyRT()
            lblCaseNumber.style.visibility="visible"
            txtCaseNumber.style.visibility="visible"
            lblProgram.style.visibility="visible"
            lstProgram.style.visibility="visible"
            If lstProgram.selectedIndex = 0 Then
                lstProgram.selectedIndex = 1
            End If
            If strReport = 27 Then
                chkDetail.checked = False
            Else
                chkDetail.checked = True
            End If
            divReviewMonth.style.top = 95
            Call ListReviewTyps(-1)
        Case 28 'Causal Factor Summary
            Call SetMgrSupWkrAbyRT()
            lblEligElement.style.visibility="visible"
            cboElement.style.visibility="visible"
            divReviewMonth.style.top = 135
			lblDetail.innerHTML="Include All Factors" & "<INPUT type=checkbox id=chkDetail style=""LEFT:100; WIDTH:20; HEIGHT:20; TOP:-2"" tabIndex=1>"
			lblDetail.style.top=195
			lblDetail.style.left = 235
			lblDetail.style.visibility="visible"
            If lstProgram.selectedIndex = 0 Then
                lstProgram.selectedIndex = 1
            End If
            Call cboProgram_onchange
            lblFactor.style.visibility="visible"
            cboFactor.style.visibility="visible"
        Case 29 ' Element Summary
            Call SetMgrSupWkrAbyRT()
            lblEligElement.innerText = "Element (Required)"
            lblEligElement.style.visibility="visible"
            cboElement.style.visibility="visible"
            divReviewMonth.style.top = 95
            If lstProgram.selectedIndex = 0 Then
                lstProgram.selectedIndex = 1
            End If
            Call cboProgram_onchange
        Case 30,75 'Element Overview
            Call SetMgrSupWkrAbyRT()
            lblEligElement.innerText = "Element"
            lblEligElement.style.visibility="visible"
            cboElement.style.visibility="visible"
            divReviewMonth.style.top = 95
            If lstProgram.selectedIndex = 0 Then
                lstProgram.selectedIndex = 1
            End If
            Call cboProgram_onchange
        Case 33 'Unsubmitted Reviews
            Call SetMgrSupWkrAbyRT()
            lblSubmitted.style.visibility = "visible"
            cboSubmitted.style.visibility = "visible"
            divReviewType.style.visibility="hidden"
            lblReviewType.style.visibility="hidden"
            divReviewTypeDefs.style.visibility="hidden"
            lblProgram.style.visibility="hidden"
            lstProgram.style.visibility="hidden"
            divReviewMonth.style.top = 55
        Case 34 'Response Due
            Call SetMgrSupWkrAbyRT()
            lblReviewType.style.visibility="hidden"
            divReviewTypeDefs.style.visibility="hidden"
            lblResponse.style.visibility="visible"
            cboResponse.style.visibility="visible"
            lblDaysPastDue.style.visibility="visible"
            txtDaysPastDue.style.visibility="visible"
            lblProgram.style.visibility="hidden"
            lstProgram.style.visibility="hidden"
            Call cboResponse_onchange()
            divReviewMonth.style.top = 85
        Case 35 'Reviewer Case Count
            Call SetMgrRvw()
	        lblReviewer.style.top = 5
	        cboReviewer.style.top = 20
	        lblReviewer.style.visibility="visible"
	        cboReviewer.style.visibility="visible"
            divReviewMonth.style.top = 5
            lblProgram.style.visibility="hidden"
            lstProgram.style.visibility="hidden"
		Case 55,72 'ReReview Elig Element overview
			Call SetMgrSupWkrAbyRT()
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
			Call cboProgram_onchange
			lblEligElement.style.top = 50
			cboElement.style.top = 65
			lblEligElement.style.visibility="visible"
			cboElement.style.visibility="visible"
	        If strReport = 55 Then
	            Form.ReReviewTypeID.value = 0
	        Else
	            Form.ReReviewTypeID.value = 1
	        End If
	        lblReviewer.style.visibility="visible"
	        cboReviewer.style.visibility="visible"
	        lblReviewer.style.top = 95
	        cboReviewer.style.top = 110
            divReviewMonth.style.top = 95
		Case 56,73 'Re-Review accuracy summary
			Call SetMgrSupWkrAbyRT()
			lblProgram.style.visibility="visible"
			lstProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
	        lblReviewer.style.visibility="visible"
	        cboReviewer.style.visibility="visible"
	        lblReviewer.style.top = 95
	        cboReviewer.style.top = 110
            divReviewMonth.style.top = 95
	        If strReport = 56 Then
	            Form.ReReviewTypeID.value = 0
	        Else
	            Form.ReReviewTypeID.value = 1
	        End If
	        divReviewType.style.left = -1000
            divReviewMonth.style.top = 55
			Call cboProgram_onchange
        Case 74 'Element Comments
            Call SetMgrSupWkrAbyRT()
            lblEligElement.style.visibility="visible"
            cboElement.style.visibility="visible"
            divReviewMonth.style.top = 95
            If lstProgram.selectedIndex = 0 Then
                lstProgram.selectedIndex = 1
            End If
            Call cboProgram_onchange
    End Select
End Sub

Sub SupVsReviewer(intWhich)
    If intWhich = 0 Then
        optSupervisor.checked = Not optSupervisor.checked
    ElseIf intWhich = 1 Then
        optReviewer.checked = Not optReviewer.checked
    End If
    
    Call ResponseDueBasedOn()
End Sub

Sub ResponseDueBasedOn()
    If optReviewer.checked = True Then
        lblWorker.style.visibility="hidden"
        cboWorker.style.visibility="hidden"
        lblSupervisor.style.visibility="hidden"
        cboSupervisor.style.visibility="hidden"
        lblReviewer.style.visibility="visible"
        cboReviewer.style.visibility="visible"
    Else
        lblWorker.style.visibility="visible"
        cboWorker.style.visibility="visible"
        lblSupervisor.style.visibility="visible"
        cboSupervisor.style.visibility="visible"
        lblReviewer.style.visibility="hidden"
        cboReviewer.style.visibility="hidden"
    End If
End Sub

Sub SetMgrSupWkrAbyRT()
    lblSupervisor.style.visibility="visible"
    lblWorker.style.visibility="visible"
    
    divReviewType.style.visibility="visible"
    lblReviewType.style.visibility="visible"
    divReviewTypeDefs.style.visibility="visible"
    
    cboSupervisor.style.visibility="visible"
    cboWorker.style.visibility="visible"
    
    lblWorker.innerText = "<%=gstrWkrTitle%>"
    divReviewType.style.visibility="visible"
    lblReviewType.style.visibility="visible"
    divReviewTypeDefs.style.visibility="visible"
End Sub

Sub SetMgrRvw()
    lblReviewer.style.visibility="visible"
    cboReviewer.style.visibility="visible"
End Sub

'Clears all criteria
Sub cmdClearCriteria_onclick()
    Dim intI
    Dim oCtl
    
    ' If any of the combos that filter down are not 0, rebuild them
    cboSupervisor.value = "0"
    cboWorker.value = "0"
    Call CheckOptionsLength()
    If lstProgram.options.length > 2 Then lstProgram.value = 0
    cboElement.value = 0
    
    txtDaysPastDue.value = ""
    
    Set oCtl = Nothing
    cboSubmitted.value = 0
    If chkInclude.checked Then chkInclude.checked = False
    txtCaseNumber.value = ""
    txtMinDays.value = 0
    cboResponse.value = 0
    cboSubmitted.value = 0
    txtStartDate.value = ""
    txtEndDate.value = ""
    txtStartReviewMonth.value = ""
    txtEndReviewMonth.value = ""
    cboElement.value = 0
    chkDetail.checked = False
    cboFactor.value = 0
    
    Call cboProgram_onchange()
    For intI = 0 To lblReviewCount.innerText - 1
        If document.all("chkReviewType" & intI).Checked Then
            Call lblChkReviewType_onclick(intI)
        End If
    Next
    
    For intI = 0 To lblReviewClassCount.innerText - 1
        If document.all("chkReviewClass" & intI).Checked Then    
            Call lblChkReviewClass_onclick(intI)
        End If
    Next
    Call lstReports_onchange()
End Sub
-->
</Script>
<SCRIPT LANGUAGE=vbscript> 
<!--
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: IncFillControls.asp                                                '
'  Purpose: Fills the controls for the Reports screen                        '
'==========================================================================='
Sub FillControls()
    Dim intI
    
    If Trim(Form.ReportIndex.Value) <> "" Then
        lstReports.selectedIndex = Form.ReportIndex.Value
    End If

    If IsDate(Form.StartDate.value) Then
        txtStartDate.value = Form.StartDate.value
    End If
    
    If IsDate(Form.EndDate.value) Then
        txtEndDate.value = Form.EndDate.value
    End If
    
    If Trim(Form.ShowDetail.value) = "Y" Then
        chkDetail.checked = True
    Else
        chkDetail.checked = False
    End If
    
    If Trim(Form.IncludeCorrect.value) = "Y" Then
        chkInclude.checked = True
    Else
        chkInclude.checked = False
    End If

    'need to default staffing here if logged in as a sup????????
        
    lstProgram.value = 0
    For intI = 0 to lstProgram.options.length - 1 
        If lstProgram.options(intI).Text = Form.ProgramText.Value Then
            If Parse(lstProgram.options(intI).value, ":", 1) = Form.ProgramID.Value Then
                lstProgram.selectedindex = intI
                Exit For
            End If
        End If
    Next
    
    For intI = 0 To cboElement.options.length - 1
        If cboElement.options(intI).Text = Form.EligElementText.value Then
            If cboElement.options(intI).value = Form.EligElementID.Value Then
                cboElement.selectedIndex = intI
                Exit For
            End If
        End If
    Next
    
    If lstProgram.options.length = 1 Then
        lblProgram.disabled = True
        lstProgram.disabled = True
    End If
    
    If Form.CaseActionID.value = "" Then
        cboCaseAction.value = 0
    Else
        cboCaseAction.value = Form.CaseActionID.value
    End If
    
    If Trim(Form.ResponseID.value) = "" Then
        cboResponse.value = 0
    Else
        cboResponse.value = Trim(Form.ResponseID.value)
    End If
    
    If Trim(Form.Submitted.value) = "" Then
        cboSubmitted.value = 0
    Else
        cboSubmitted.value = Trim(Form.Submitted.value)
    End If
        
    If Trim(Form.MinAvgDays.value) = "" Then
        txtMinDays.value = 0
    Else
        txtMinDays.value = Form.MinAvgDays.value
    End If

    txtCaseNumber.value = Form.CaseNumber.value
End Sub

Sub SelectStaffCboValue(cboCtl, strVal)
    'First attempt to set the cboCtl to the passed in value:
    cboCtl.value = strVal
    If IsNull(cboCtl.value) Or cboCtl.value = "" Then
        'Cbo was not set, so set combo to default value:
        If cboCtl.Id = "cboSupervisor" Then
            cboCtl.value = ""
            If IsNull(cboCtl.value) Or cboCtl.value = "" Then
                'Cbo still not set, so select the default <All> entry:
                cboCtl.value = 0
            End If
        Else
            cboCtl.Value = 0
        End If
    End If
End Sub

Function FillForms(blnEditCriteria)
    Dim strExReviewType
    Dim strSpecificProg
    Dim blnSelected
    Dim intJ
    Dim strAllowGroup

    strAllowGroup = "[52]"
    strSpecificProg = "[26][28][29][30][31][45][49][50][51][52][53][55][56][57][72][73][75]"
    strExReviewType = "[32][33][34][35][36][47][55][56][72][73]"
    strExReviewClass = "[33][35][47][55][56][72][73]"
    Call ClearForms()
    
    'Edit Criteria only if viewing a report, not if saving a date change
    If blnEditCriteria Then
        If Not IsNumeric(txtGroupID.value) Then txtGroupID.value = -1
        If lstProgram.value = 0 Then txtGroupID.value = -1
        If Instr(strAllowGroup, Parse(lstReports.value, ":", 1)) = 0 Then
            If lstProgram.value < 0 And txtGroupID.value > 0 And Instr(strSpecificProg, Parse(lstReports.value, ":", 1)) > 0 Then
                MsgBox "Please select a Specific Program", vbInformation, "View Report"
                lstProgram.focus
                Exit Function
            End If
        End If
        If InStr(strSpecificProg, "[" & Parse(lstReports.value, ":", 1) & "]") > 0 Then
            If lstProgram.value <= 0 And (txtGroupID.value <= 0 Or Instr(strAllowGroup, Parse(lstReports.value, ":", 1)) = 0) Then
                If Instr(strAllowGroup, Parse(lstReports.value, ":", 1)) = 0 Then
                    MsgBox "Please select a Specific Program", vbInformation, "View Report"
                Else
                    MsgBox "Please select a Specific Program or Program Group", vbInformation, "View Report"
                End If
                lstProgram.focus
                Exit Function
            End If
        End If

	    If Parse(lstReports.value, ":", 1) = "75" Or Parse(lstReports.value, ":", 1) = "76" Then
			If Not IsDate(txtStartDate.value) Or Not IsDate(txtEndDate.value) Then
				MsgBox "This report is designed to be run for a 6 month to 1 year time period." & vbCrlf & vbCrLf & "Please enter a start and end for the report date range." & Space(5), vbInformation, "View Report"
				If Not IsDate(txtStartDate.value) Then
					txtStartDate.focus
				Else
					txtEndDate.focus
				End If
				Exit Function
			Else
				txtStartDate.value = Right("00" & Month(txtStartDate.value),2) & "/01/" & Year(txtStartDate.value)
			    dtmMin = DateAdd("m",6,txtStartDate.value)
			    dtmMin = DateAdd("d",-1,dtmMin)
			    dtmMax = DateAdd("m",12,txtStartDate.value)
			    dtmMax = DateAdd("d",-1,dtmMax)
			    
				If CDate(txtEndDate.value) < CDate(dtmMin) Or CDate(txtEndDate.value) > CDate(dtmMax) Then
					MsgBox "This report is designed to be run for a six month to one year time period." & vbCrLf & vbCrLf & "Please enter a six month to one year time period for the report date range." & Space(5), vbInformation, "View Report"
					If Not IsDate(txtStartDate.value) Then
						txtStartDate.focus
					Else
						txtEndDate.focus
					End If
					Exit Function
				End If
			End If
		End If    

        If Parse(lstReports.value, ":", 1) = "127" Or Parse(lstReports.value, ":", 1) = "138" Then
            If cboWorker.Value = "0" Then
                MsgBox "This report is designed to be run by Worker." & vbCrlf & vbCrLf & "Please Select a Worker." & Space(5), vbInformation, "View Report"
                cboWorker.focus
                Exit Function
            End If
        End If
        
		If Parse(lstReports.value, ":", 1) = "29" Or Parse(lstReports.value, ":", 1) = "74" Then
			If cboElement.Value = 0 Then
				MsgBox "This report is designed to be run by Element." & vbCrlf & vbCrLf & "Please Select Element." & Space(5), vbInformation, "View Report"
				cboElement.focus
				Exit Function
			End If
		End If
	    
        If Instr(strExReviewType, "[" & Parse(lstReports.value, ":", 1) & "]") = 0 And Not PutReviewTypesToForm Then
            'Specific Full Review not Selected
            If lstProgram.value < 50 And lstProgram.value <> 6 Then
                MsgBox "This report is designed to be run by Review Type." & vbCrlf & vbCrLf & "Please Select a Review Type." & Space(5), vbInformation, "View Report"
                Exit Function
            End If
        End If
        
        If Instr(strExReviewClass, "[" & Parse(lstReports.value, ":", 1) & "]") = 0 And Not PutReviewClassToForm(Parse(lstReports.value, ":", 1)) Then
            MsgBox "This report is designed to be run by <% = gstrReviewClassTitle %>." & vbCrlf & vbCrLf & "Please Select a <% = gstrReviewClassTitle %>." & Space(5), vbInformation, "View Report"
            Exit Function
        End If
    End If    
    If lstProgram.value < 0 Then
        ' Group selected
        Form.ProgramID.value = CInt(txtGroupID.value) * -1
    Else
        Form.ProgramID.Value = lstProgram.value
    End If
    Form.ProgramText.Value = GetComboText(lstProgram)
 
    Form.FactorID.Value = cboFactor.value
    Form.FactorText.Value = GetComboText(cboFactor)
    
    Form.StartDate.value = txtStartDate.value
    Form.EndDate.value = txtEndDate.value

    Form.Director.value = cboDirector.value
    Form.Office.value = cboOffice.value
    Form.ProgramManager.value = cboManager.value
    If Form.Director.value = "0" Then Form.Director.value = ""
    If Form.Office.value = "0" Then Form.Office.value = ""
    If Form.ProgramManager.value = "0" Then Form.ProgramManager.value = ""
    
    Form.Worker.value = cboWorker.value
    Form.Supervisor.value = cboSupervisor.value
    If Form.Worker.value = "0" Then Form.Worker.value = ""
    If Form.Supervisor.value = "0" Then Form.Supervisor.value = ""
    
    Form.Reviewer.Value = cboReviewer.value
    If Form.Reviewer.value = "0" Then Form.Reviewer.value = ""
    Form.ReReviewer.value = txtReReviewer.value

    Form.EligElementID.Value = cboElement.value
    Form.EligElementText.Value = GetComboText(cboElement)
    
    Form.ResponseID.value = cboResponse.value
    Form.ResponseText.value = GetComboText(cboResponse)
    If cboResponse.value = 0 Or cboResponse.value = 3 Then
        Form.DaysPastDue.value = ""
    Else
        Form.DaysPastDue.value = txtDaysPastDue.value
    End If
    
    Form.CaseNumber.Value = txtCaseNumber.value
    Form.MinAvgDays.Value = txtMinDays.Value
    Form.Submitted.Value = cboSubmitted.value
        
    If chkDetail.checked Then
        Form.ShowDetail.value = "Y"
    Else
        Form.ShowDetail.value = "N"
    End If
    If chkInclude.checked Then
        Form.IncludeCorrect.value = "Y"
    Else
        Form.IncludeCorrect.value = "N"
    End If

    If divReviewMonth.style.visibility = "visible" Then
        If txtStartReviewMonth.value <> "" And txtEndReviewMonth.value <> "" And txtStartReviewMonth.value <> "(MM/YYYY)" And txtEndReviewMonth.value <> "(MM/YYYY)" Then
            ' Verify that the start is before end month
            intPos = InStr(txtStartReviewMonth.value, "/")
            If intPos > 0 Then
                strMonth = Trim(Mid(txtStartReviewMonth.value, 1, intPos -1))
                strYear = Trim(Mid(txtStartReviewMonth.value, intPos + 1))
            End If
            dtmReviewMonth = strMonth & "/01/" & strYear
            
            intPos = InStr(txtEndReviewMonth.value, "/")
            If intPos > 0 Then
                strMonth = Trim(Mid(txtEndReviewMonth.value, 1, intPos -1))
                strYear = Trim(Mid(txtEndReviewMonth.value, intPos + 1))
            End If
            dtmEndReviewMonth = strMonth & "/01/" & strYear
            
            If CDate(dtmReviewMonth) > CDate(dtmEndReviewMonth) Then
                MsgBox "Start Review Month/Year cannot be after End Review Month/Year.", vbInformation, "View Report"
                txtStartReviewMonth.focus
                Exit Function
            End If
        End If

        ' Convert Review Month/Year fields to dates
        If txtStartReviewMonth.value <> "" Then
            intPos = Instr(txtStartReviewMonth.value, "/")
            If intPos > 0 Then
                dtmReviewMonth = Trim(Mid(txtStartReviewMonth.value, 1, intPos -1)) & "/01/" & Trim(Mid(txtStartReviewMonth.value, intPos + 1))
            Else
                dtmReviewMonth = ""
            End If
            Form.StartReviewMonth.value = dtmReviewMonth
        Else
            Form.StartReviewMonth.value = txtStartReviewMonth.value
        End If

        If txtEndReviewMonth.value <> "" Then
            intPos = Instr(txtEndReviewMonth.value, "/")
            If intPos > 0 Then
                dtmReviewMonth = Trim(Mid(txtEndReviewMonth.value, 1, intPos -1)) & "/01/" & Trim(Mid(txtEndReviewMonth.value, intPos + 1))
            Else
                dtmReviewMonth = ""
            End If
            Form.EndReviewMonth.value = dtmReviewMonth
        Else
            Form.EndReviewMonth.value = txtEndReviewMonth.value
        End If
    Else
        Form.StartReviewMonth.value = ""
        Form.EndReviewMonth.value = ""
    End If

    Form.ReportIndex.value = lstReports.selectedIndex
    If mstrCboBoxDefaultText <> "" Then
        'Strip out the "<All>" values before passing over to report:
           Form.Director.value = Replace(Form.Director.value,mstrCboBoxDefaultText,"")
           Form.ProgramManager.value = Replace(Form.ProgramManager.value,mstrCboBoxDefaultText,"")
           Form.Supervisor.value = Replace(Form.Supervisor.value,mstrCboBoxDefaultText,"")
           Form.Worker.value = Replace(Form.Worker.value,mstrCboBoxDefaultText,"")
           Form.Reviewer.value = Replace(Form.Reviewer.value,mstrCboBoxDefaultText,"")
           Form.ReReviewer.value = Replace(Form.ReReviewer.value,mstrCboBoxDefaultText,"")
           Form.EligElementText.value = Replace(Form.EligElementText.value,mstrCboBoxDefaultText,"")
           Form.CaseActionText.value = Replace(Form.CaseActionText.value,mstrCboBoxDefaultText,"")
           Form.ResponseText.value = Replace(Form.ResponseText.value,mstrCboBoxDefaultText,"")
           Form.FactorText.value = Replace(Form.FactorText.value,mstrCboBoxDefaultText,"")
    End If
    Form.ReportNum.value = Parse(lstReports.value, ":", 1)
    Form.ReportTitle.Value = lstReports.Options(LstReports.selectedIndex).Text
    Form.action = Parse(lstReports.value, ":", 2)
    FillForms = True
End Function

Function GetNameWithPosNum(oCbo)
    'This function combines the position number back with the name, if it
    'is not already part of the name, before it is sent on as a parameter
    'for the reports.  This was done because the names stored in the review
    'fields have the "name - number" format, so that the parameter and the
    'field value could be compared without the previous method of CHARINDEX.
    Dim intJ
    Dim strPosNum
    Dim strName

    strName = GetComboText(oCbo)
    GetNameWithPosNum = strName
End Function

Sub ClearForms()
    Form.StartDate.value = ""
    Form.EndDate.value = ""
    Form.DirectorID.value = 0
    Form.Director.value = ""
    Form.OfficeID.value = 0
    Form.Office.value = ""
    Form.ProgramManagerID.value = 0
    Form.ProgramManager.value = ""
    Form.SupervisorID.value = 0
    Form.Supervisor.value = ""
    Form.WorkerID.value = 0
    Form.Worker.value = ""
    Form.ReviewerID.Value = 0
    Form.Reviewer.Value = ""
    Form.ReviewTypeID.Value = 0
    Form.ReviewTypeText.Value = ""
    Form.ReviewClassID.Value = 0
    Form.ReviewClassText.Value = ""
    Form.ProgramID.Value = 0
    Form.ProgramText.Value = ""
    Form.EligElementID.Value = 0
    Form.EligElementText.Value = ""
    Form.CaseActionID.value = 0
    Form.CaseActionText.value = ""
    Form.CaseNumber.Value = ""
    Form.MinAvgDays.Value = 0
    Form.ShowNonComOnly.Value = ""
    Form.ShowNonComOnlyID.Value = 0
    Form.Submitted.Value = ""
    Form.ShowDetail.value = ""
    Form.IncludeCorrect.value = ""
    Form.ReportingMode.value = 0
End Sub
-->
</SCRIPT>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncFormsReportDef.asp                                              '
' Purpose: This include file contains the definition of the input elements  '
'          for the HTML form used to post Case data entry.  This version is '
'          for a blank form.                                                '
'                                                                           '
'==========================================================================='
Sub ReportFormDef()
    'ReportIndex is currently being used as a temporary solution for keeping
    'track of which specific report was selected, now that some reports have
    'the same criteria screen:
    WriteFormField "ReportIndex", ReqForm("ReportIndex")
    WriteFormField "ReportTitle", ReqForm("ReportTitle")
    WriteFormField "StartDate", ReqForm("StartDate")
    WriteFormField "EndDate", ReqForm("EndDate")
    WriteFormField "ReportingMode", ReqForm("ReportingMode")
    WriteFormField "ReportNum", ReqForm("ReportNum")
    
    WriteFormField "DirectorID", ReqForm("DirectorID")
    WriteFormField "Director", ReqForm("Director")
    WriteFormField "OfficeID", ReqForm("OfficeID")
    WriteFormField "Office", ReqForm("Office")
    WriteFormField "ProgramManagerID", ReqForm("ProgramManagerID")
    WriteFormField "ProgramManager", ReqForm("ProgramManager")
    WriteFormField "SupervisorID", ReqForm("SupervisorID")
    WriteFormField "Supervisor", ReqForm("Supervisor")
    WriteFormField "WorkerID", ReqForm("WorkerID")
    WriteFormField "Worker", ReqForm("Worker")
    WriteFormField "ReviewerID", ReqForm("ReviewerID")
    WriteFormField "Reviewer", ReqForm("Reviewer")
    WriteFormField "ReReviewerID", ReqForm("ReReviewerID")
    WriteFormField "ReReviewer", ReqForm("ReReviewer")
    
    WriteFormField "ReviewTypeID", ReqForm("ReviewTypeID")
    WriteformField "ReviewTypeText", ReqForm("ReviewTypeText")
    WriteFormField "ReviewClassID", ReqForm("ReviewClassID")
    WriteformField "ReviewClassText", ReqForm("ReviewClassText")
    
    WriteFormField "ProgramID", ReqForm("ProgramID")
    WriteFormField "ProgramText", ReqForm("ProgramText")
    WriteFormField "EligElementID", ReqForm("EligElementID")
    WriteFormField "EligElementText", ReqForm("EligElementText")
    
    WriteFormField "CaseActionID", ReqForm("CaseActionID")
    WriteFormField "CaseActionText", ReqForm("CaseActionText")
    WriteFormField "ResponseID", ReqForm("ResponseID")
    WriteFormField "ResponseText", ReqForm("ResponseText")
    WriteFormField "CaseNumber", ReqForm("CaseNumber")
    WriteFormField "MinAvgDays", ReqForm("MinAvgDays")
    WriteFormField "ShowNonComOnlyID", ReqForm("ShowNonComOnlyID")
    WriteFormField "ShowNonComOnly", ReqForm("ShowNonComOnly")
    WriteFormField "Submitted", ReqForm("Submitted")
    WriteFormField "ShowDetail", ReqForm("ShowDetail")
    WriteFormField "IncludeCorrect", ReqForm("IncludeCorrect")
    WriteFormField "DaysPastDue", ReqForm("DaysPastDue")
    WriteFormField "StartReviewMonth", ReqForm("StartReviewMonth")
    WriteFormField "EndReviewMonth", ReqForm("EndReviewMonth")
    WriteFormField "RespDueBasedOn", ReqForm("RespDueBasedOn")
    WriteFormField "SSProgramID", ReqForm("SSProgramID")
    WriteFormField "ParentLoading", "Y"
    WriteFormField "RepLUserAdmin", gblnUserAdmin
    WriteFormField "RepLUserQA", gblnUserQA
    WriteFormField "RepLAliasPosID", glngAliasPosID
    WriteFormField "RepLUserID", gstrUserID
    WriteFormField "StaffInformation", ""
    WriteFormField "ReReviewTypeID", ""
End Sub
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncCrtReviewType.asp                                             '
' Purpose: This include file contains functions used to build the review    '
'          type screen elements used on report criteria (Crt) pages.        '
'==========================================================================='

Sub WriteReviewControls()
    '-----------------------------------------------------------------------'
    'Retrieve the list of review types from the database and build HTML     '
    'elements for each review type.                                         '
    '-----------------------------------------------------------------------'
    Dim intCount
    
    intCount = 0
    
    Set gadoCmd = GetAdoCmd("spGetReviewTypeElms")
        AddParmIn gadoCmd, "@Programs", adVarChar, 100, ReqForm("ProgramsSelected")
        AddParmIn gadoCmd, "@EffectiveDate", adDBTimeStamp, 0, null
    Set madoRs = GetAdoRs(gadoCMd)
    
    'Writes out checkboxes and supporting controls for review type on report criteria screens:
    mintFULLID = intCount
    
    intCount = intCount + 1
    Do While Not madoRs.EOF
        'Write out visible Checkbox control for each review type.
        mstrRvw = mstrRvw & "<OPTION VALUE=" & madoRs.Fields("ReviewTypeID").Value & ">" & madoRs.Fields("ReviewTypeText").Value & "|" & madoRs.Fields("rteElementID").Value & "|" & madoRs.Fields("rteProgramID").Value & "|" & madoRs.Fields("rteStartDate").Value & "|" & madoRs.Fields("rteEndDate").Value & "</OPTION>"
        intCount = intCount + 1
        madoRs.MoveNext
    Loop
        
    madoRs.Close
    Set gadoCmd = Nothing
End Sub

Sub WriteReviewClass()
    Set adCmd = nothing
    Set adRs = nothing
    Set adCmd = Server.CreateObject("ADODB.Command")
    Set adRs = Server.CreateObject("ADODB.RecordSet")
    With adCmd
        .ActiveConnection = gadoCon
        .CommandType = adCmdStoredProc
        .CommandText = "spGetReviewClass"
        .CommandTimeout = 180
        
    End With
    adRs.Open adCmd,, adOpenForwardOnly, adLockReadOnly
    intCountClass = 0
    mstrRvwClass = "<BR>"
    Do While Not adRs.EOF
        mstrRvwClass = mstrRvwClass & "<INPUT TYPE=checkbox ID=chkReviewClass" & intCountClass & " onclick=chkReviewClass_onclick(" & intCountClass & ") VALUE=" & adRs.Fields("ClassID").Value & " STYLE=""LEFT:5"" tabIndex=11>"
        mstrRvwClass = mstrRvwClass & "<SPAN id=lblChkReviewClass" & intCountClass & " onclick=""lblChkReviewClass_onclick(" & intCountClass & ")"" STYLE=""LEFT:30"">" & adRs.Fields("ClassText").value & "</SPAN><BR>" & vbCrLf 
        intCountClass = intCountClass + 1
        adRs.MoveNext
    Loop
    Response.Write "<SPAN id=lblReviewClassCount style=""visibility:hidden"">" & intCountClass & "</SPAN>" & vbCrLf
End SUb
'Client side functions:
'----------------------------------------------------------------------------%>
<SCRIPT id=ReviewTypeFunctions language=vbscript>
<!--
Sub GetReviewTypesFromForm()
    Dim intI
    Dim oCtl
    
    If Not IsNumeric(lblReviewCount.innerText) Then
        Exit Sub
    End If
    For intI = 0 To lblReviewCount.innerText - 1
        Set oCtl = document.all("chkReviewType" & intI)
      
        If Instr(Form.ReviewTypeID.Value, "[" & oCtl.value & "]") > 0 Then
            oCtl.checked = True
        End If
    Next
    Set oCtl = Nothing
End Sub

Function PutReviewTypesToForm()
    Dim intI
    Dim oCtl
    Dim oLbl
    
    PutReviewTypesToForm = False

    If Not IsNumeric(lblReviewCount.innerText) Then
        Exit Function
    End If
    Form.ReviewTypeID.Value = ""
    For intI = 0 To lblReviewCount.innerText - 1
        Set oCtl = document.all("chkReviewType" & intI)
        Set oLbl = document.all("lblChkReviewType" & intI)
        If oCtl.Checked Then
            PutReviewTypesToForm = True
            Form.ReviewTypeID.Value = Form.ReviewTypeID.Value & "[" & oCtl.value & "]"
            Form.ReviewTypeText.Value = oLbl.innerText & "||" & Form.ReviewTypeText.Value 
        End If        
    Next
   
    Set oCtl = Nothing
    Set oLbl = Nothing
End Function

Sub GetReviewClassFromForm()
    Dim intI
    Dim oCtl

    For intI = 0 To lblReviewClassCount.innerText - 1
        Set oCtl = document.all("chkReviewClass" & intI)
       
        If Instr(Form.ReviewClassID.Value, "[" & oCtl.value & "]") > 0 Then
            oCtl.checked = True
        End If
    Next
    Set oCtl = Nothing
End Sub

Function PutReviewClassToForm(strReportID)
    Dim intI
    Dim oCtl
    Dim oLbl
    
    PutReviewClassToForm = False
    Select Case strReportID
        Case "67","69" ' Error Type and Screen Error reports have QA auto selected
            For intI = 0 To lblReviewClassCount.innerText - 1
                Set oCtl = document.all("chkReviewClass" & intI)
                Set oLbl = document.all("lblChkReviewClass" & intI)
                If oCtl.value = 258 Then
                    PutReviewClassToForm = True
                    Form.ReviewClassID.Value = Form.ReviewClassID.Value & "[" & oCtl.value & "]"
                    Form.ReviewClassText.Value = oLbl.innerText & "||" & Form.ReviewClassText.Value
                End If        
            Next
        Case Else
            For intI = 0 To lblReviewClassCount.innerText - 1
                Set oCtl = document.all("chkReviewClass" & intI)
                Set oLbl = document.all("lblChkReviewClass" & intI)
                If oCtl.Checked Then
                    PutReviewClassToForm = True
                    Form.ReviewClassID.Value = Form.ReviewClassID.Value & "[" & oCtl.value & "]"
                    Form.ReviewClassText.Value = oLbl.innerText & "||" & Form.ReviewClassText.Value
                End If        
            Next
    End Select
    Set oCtl = Nothing
    Set oLbl = Nothing
End Function

Sub lblChkReviewType_onclick(intChk)
    Dim intI
 
    document.all("chkReviewType" & intChk).checked = Not document.all("chkReviewType" & intChk).checked
    If document.all("chkReviewType" & intChk).checked = False And chkCheckAllRT.checked = True Then
        chkCheckAllRT.checked = False
    End If
    Call FillElementDropDown(lstProgram.value)
End Sub

Sub chkReviewClass_onclick(intChk)
    If document.all("chkReviewClass" & intChk).value <> 278 Then 'Informal is excluded from all
        If document.all("chkReviewClass" & intChk).checked = False And chkCheckAllRC.checked = True Then
            chkCheckAllRC.checked = False
        End If
    End If
End Sub

Sub lblChkReviewClass_onclick(intChk)
    document.all("chkReviewClass" & intChk).checked = Not document.all("chkReviewClass" & intChk).checked
    Call chkReviewClass_onclick(intChk)
End Sub

Sub ChkReviewType_onclick(intValue)
    If document.all("chkReviewType" & intValue).checked = False And chkCheckAllRT.checked = True Then
        chkCheckAllRT.checked = False
    End If

    Call FillElementDropDown(lstProgram.value)
End Sub
-->
</script>

<!--#include file="IncBuildList.asp"-->
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncNavigateControls.asp"-->
