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
Dim mstrItem, mlngTabIndex
Dim moRevType, mdctRevTypes, mdctPrograms
Dim intRT
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

Set mdctRevTypes = CreateObject("Scripting.Dictionary")
Set mdctPrograms = CreateObject("Scripting.Dictionary")

Set adRsPrg = Server.CreateObject("ADODB.Recordset")
Set adCmd = GetAdoCmd("spArchiveLists")
    AddParmIn adCmd, "@ListType", adVarchar, 255, NULL
    'Call ShowCmdParms(adCmdPrg) '***DEBUG
    adRsPrg.Open adCmd, , adOpenForwardOnly, adLockReadOnly
Set adCmd = Nothing
intRT = 1
adRsPrg.Sort = "ListType, ListItem"
mdctRevTypes.Add 1, "1^Full^0^01/01/2000^^" 
Do While Not adRsPrg.EOF
    Select Case adRsPrg.Fields("ListType").Value
        Case "Program"
            mdctPrograms.Add Parse(adRsPrg.Fields("ListItem").Value,"^",3), Parse(adRsPrg.Fields("ListItem").Value,"^",3) & "^" & Parse(adRsPrg.Fields("ListItem").Value,"^",1) & "^1^x^N"
        Case "Review Type"
            If Parse(adRsPrg.Fields("ListItem").Value,"^",1) <> "Full" Then
                intRT = intRT + 1
                mdctRevTypes.Add intRT, intRT & "^" & _
                    Parse(adRsPrg.Fields("ListItem").Value,"^",1) & "^" & _
                    Parse(adRsPrg.Fields("ListItem").Value,"^",4) & "^01/01/2000^^"
            End If
    End Select
    adRsPrg.MoveNext
Loop

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
Dim mblnUseAuthby       
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>
Dim mdctPrograms
Dim mdctRevTypes
Dim mstrCboBoxDefaultText
Dim mstrCboBoxDefaultTextHTML
Dim mdctOffices, mdctDirectors, mdctManagers
Dim mblnMainClosed
Dim maParentNames(4), mintStaffTimer
Dim mblnStaffLoaded, mintStaffTimer_W
Dim mstrHoldStaff, mctlStaff

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
    For intI = 1 To 4
        maParentNames(intI) = "^"
    Next
    mblnStaffLoaded = True
    mblnSetFocusToMain = True
    mblnMainClosed = False
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>
    
    Call SizeAndCenterWindow(767, 520, True)
    lblStatusChange.style.visibility = "visible"
    lblViewReport.style.visibility = "hidden"
    
    Set mdctRevTypes = CreateObject("Scripting.Dictionary")
    Set mdctPrograms = CreateObject("Scripting.Dictionary")
    Set mdctOffices = CreateObject("Scripting.Dictionary")
    Set mdctDirectors = CreateObject("Scripting.Dictionary")
    Set mdctManagers = CreateObject("Scripting.Dictionary")
 <%
    For Each moRevType In mdctRevTypes
        Response.Write "mdctRevTypes.Add CLng(" & moRevType & "), """ & mdctRevTypes(moRevType) & """" & vbCrLf
    Next
    For Each moRevType In mdctPrograms
        Response.Write "mdctPrograms.Add CLng(" & moRevType & "), """ & mdctPrograms(moRevType) & """" & vbCrLf
    Next
%>   
    Call FillStaffDictionaries()

    Call FillStaffCboLists()
    For intI = 1 To 4
        maParentNames(intI) = cboDirector.value & "^" & cboOffice.value
    Next
    Call lstReports_onchange()
    Call FillControls()
    
    If lstProgram.options.length = 2 Then
        lstProgram.selectedIndex = 1
        lstProgram.disabled = True
    End If
    Call GetReviewTypesFromForm()
    Call GetReviewClassFromForm()

    Call ChangeBlankOptionToAll    

    HideShowFrames("visible")
    lblStatusChange.style.visibility = "hidden"
    lblViewReport.style.visibility = "hidden"
    
    'Remove the "No Error" choice from the Benefit Error Type
    'list, since it is not applicable for the reports:
    For intI = 0 To cboBenErrorType.options.length - 1
        If cboBenErrorType.options(intI).Text = "No Error" Then
            cboBenErrorType.options.Remove intI
            Exit For
        End If
    Next
    
    PageBody.style.cursor = "default"
    Call CheckOptionsLength()
    
    optReportMode2.checked = True
    Call optReportMode_onclick(1)
    If "<% = mstrRespDueBasedOn %>" = "Reviewer" Then
        optReviewer.checked = True
    Else
        optSupervisor.checked = True
    End If
    txtSupervisor.value = "<All>"
    txtReviewer.value = "<All>"
    txtReReviewer.value = "<All>"
    txtWorker.value = "<All>"
End Sub

Sub FillStaffDictionaries()
    <%
    'Director(Missouri=Region) ------------------
    Set adCmd = GetAdoCmd("spArchiveGetStaffing")
        AddParmIn adCmd, "@RoleName", adVarchar, 100, "Director"
        'Call ShowCmdParms(adCmdPrg) '***DEBUG
        Set adRs = GetAdoRs(adCmd)
    Set adCmd = Nothing
    Do While Not adRs.EOF
        mstrItem = ""
        mstrItem = mstrItem & adRs.Fields("StaffName").Value
        Response.Write vbTab & "mdctDirectors.Add """ & mstrItem & """, """ & mstrItem & """" & vbCrLf
        adRs.MoveNext
    Loop
    adRs.Close
    Set adRs = Nothing
    'Office(Missouri=FIPs) ------------------
    Set adCmd = GetAdoCmd("spArchiveGetStaffing")
        AddParmIn adCmd, "@RoleName", adVarchar, 100, "Office"
        'Call ShowCmdParms(adCmdPrg) '***DEBUG
        Set adRs = GetAdoRs(adCmd)
    Set adCmd = Nothing
    Do While Not adRs.EOF
        mstrItem = ""
        mstrItem = mstrItem & adRs.Fields("StaffName").Value & "^" & adRs.Fields("Supervisor").Value
        Response.Write vbTab & "mdctOffices.Add """ & mstrItem & """, """ & mstrItem & """" & vbCrLf
        adRs.MoveNext
    Loop
    adRs.Close
    Set adRs = Nothing
    'Manager(Missouri=Office Manager) ------------------
    Set adCmd = GetAdoCmd("spArchiveGetStaffing")
        AddParmIn adCmd, "@RoleName", adVarchar, 100, "Manager"
        'Call ShowCmdParms(adCmdPrg) '***DEBUG
        Set adRs = GetAdoRs(adCmd)
    Set adCmd = Nothing
    Do While Not adRs.EOF
        mstrItem = ""
        mstrItem = mstrItem & adRs.Fields("StaffName").Value & "^" & adRs.Fields("Supervisor").Value
        Response.Write vbTab & "mdctManagers.Add """ & mstrItem & """, """ & mstrItem & """" & vbCrLf
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
            If divStaffSearch.style.left <> "-1000px" Then
                Call fraStaffSearch.LostFocus()
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
    Form.Action = "ArchiveMenu.asp"
    Form.Submit
End Sub

<%'If Main has not been closed, set focus back to it.%>
Sub window_onbeforeunload()
    If mblnSetFocusToMain = True And mblnMainClosed = False Then
        'window.opener.focus
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
Sub txtStartDate_onblur
    If Trim(txtStartDate.value) = "(MM/DD/YYYY)" Or Trim(txtStartDate.value) = ""Then
        txtStartDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtStartDate.value) Then
        MsgBox "The Start Date must be a valid date - MM/DD/YYYY.", vbInformation, "Enter Report Criteria"
        txtStartDate.focus
    Else
        Select Case Parse(lstReports.value,":",1)
            Case 76,75
                txtEndDate.value = DateAdd("m",6,txtStartDate.value)
                txtEndDate.value = DateAdd("d",-1,txtEndDate.value)
            Case 47,60,35
                txtEndDate.value = DateAdd("m",1,txtStartDate.value)
                txtEndDate.value = DateAdd("d",-1,txtEndDate.value)
            Case Else
        End Select
    End If
    ValidStaffInDateRange("txtSupervisor")
    ValidStaffInDateRange("txtWorker")
    ValidStaffInDateRange("txtReviewer")
End Sub

'Refills fillers for End Date if blank
Sub txtEndDate_onfocus()
    If Trim(txtEndDate.value) = "" Then
        txtEndDate.value = "(MM/DD/YYYY)"
    End If
    txtEndDate.select
End Sub

'Clears End Date if left blank, Checks that a valid date was entered
Sub txtEndDate_onblur
    If Trim(txtEndDate.value) = "(MM/DD/YYYY)" Then
        txtEndDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtEndDate.value) Then
        MsgBox "The Start Date must be a valid date - MM/DD/YYYY.", vbInformation, "Enter Report Criteria"
        txtEndDate.focus
    End If
    ValidStaffInDateRange("txtSupervisor")
    ValidStaffInDateRange("txtWorker")
    ValidStaffInDateRange("txtReviewer")
End Sub

Sub ValidStaffInDateRange(strTextBox)
    Dim ctlName, ctlStartDate, ctlEndDate
    Dim dtmReportStartDate, dtmReportEndDate
    
    Set ctlName = document.all(strTextBox)
    Set ctlStartDate = document.all(strTextBox & "StartDate")
    Set ctlEndDate = document.all(strTextBox & "EndDate")
    
    If ctlName.value = "<All>" Then
        Exit Sub
    End If
    
    If ctlStartDate.value = "" Then ctlStartDate.value = "01/01/1900"
    If ctlEndDate.value = "" Then ctlEndDate.value = "12/31/3100"
    dtmReportStartDate = txtStartDate.value
    If dtmReportStartDate = "" Then dtmReportStartDate = "01/01/2000"
    dtmReportEndDate = txtEndDate.value
    If dtmReportEndDate = "" Then dtmReportEndDate = "12/31/2100"
    
    If CDate(ctlEndDate.value) < CDate(dtmReportStartDate) Or CDate(ctlStartDate.value) > CDate(dtmReportEndDate) Then
        ctlStartDate.value = ""
        ctlEndDate.value = ""
        ctlName.value = "<All>"
    End If
End Sub

Sub optReportMode_onClick(intMode)
    txtSupervisor.value = "<All>"
    txtReviewer.value = "<All>"
    txtReReviewer.value = "<All>"
    txtWorker.value = "<All>"
    document.all("txtSupervisorStartDate").value = ""
    document.all("txtSupervisorEndDate").value = ""
    document.all("txtReviewerStartDate").value = ""
    document.all("txtReviewerEndDate").value = ""
    document.all("txtWorkerStartDate").value = ""
    document.all("txtWorkerEndDate").value = ""
    document.all("txtReReviewerStartDate").value = ""
    document.all("txtReReviewerEndDate").value = ""

    Call Display_Criteria
End Sub

Sub lblReportMode_onclick(intMode)
    optReportMode(intMode).checked = true
    Call optReportMode_onclick(intMode)
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

Sub cboTab_onchange()
    Call FillElementDropDown(lstProgram.Value)
End Sub

Sub StaffComboOnChange(intMngLvlID)
    Dim intParentID
    Dim blnReload
    Dim intI
    Dim blnSupOk
 
    If Form.onchangeFlag.value = "N" Then Exit Sub
    If intMngLvlID >= 4 Then 
        Call FillStaffCbo(cboOffice)
        Call FillStaffCbo(cboManager)
    ElseIf intMngLvlID >= 3 Then 
        Call FillStaffCbo(cboManager)
    End If
    Call ClearStaffingFields("All")
End Sub

Sub ClearStaffingFields(strLevel)
    If strLevel = "S" Or strLevel = "All" Then
        txtSupervisor.value = "<All>"
        txtSupervisorStartDate.value = ""
        txtSupervisorEndDate.value = ""
    End If
    If strLevel = "W" Or strLevel = "All" Then
        txtWorker.value = "<All>"
        txtWorkerStartDate.value = ""
        txtWorkerEndDate.value = ""
    End If
    If strLevel = "R" Or strLevel = "All" Then
        txtReviewer.value = "<All>"
        txtReviewerStartDate.value = ""
        txtReviewerEndDate.value = ""
    End If
    If strLevel = "RR" Or strLevel = "All" Then
        txtReReviewer.value = "<All>"
        txtReReviewerStartDate.value = ""
        txtReReviewerEndDate.value = ""
    End If
End Sub

Sub cboDirector_onchange()
    Call StaffComboOnChange(4)
End Sub
Sub cboDirector_onkeypress()
    window.event.returnValue = False
End Sub

Sub cboOffice_onchange()
    Call StaffComboOnChange(3)
End Sub
Sub cboOffice_onkeypress()
    window.event.returnValue = False
End Sub

Sub cboManager_onchange()
    Call StaffComboOnChange(2)
End Sub
Sub cboManager_onkeypress()
    window.event.returnValue = False
End Sub

Sub FillElementDropDown(intPrg)
    Dim oOption
    Dim intI
    Dim strRecord
    Dim oDictObj, strElementList
        
    cboElement.options.length = Null
    
    Set oOption = Document.createElement("OPTION")
    oOption.value = 0
    oOption.Text = "<All>"
    cboElement.options.add oOption
    If cboTab.value = 0 Or intPrg = "" Or lblReviewCount.value = "" Or lblReviewCount.value = "0" Then Exit Sub

    For Each oDictObj In window.opener.mdctElements
        If CInt(oDictObj) <> CInt(window.opener.mintArrearageID) Then
            strRecord = window.opener.mdctElements(oDictObj)
            If CLng(intPrg) = 6 And CLng(cboTab.value) = 1 Then
                'For Enf Rem, Action Integrity, the actions must be taken from program IDs 50-70
                If CLng(Parse(strRecord,"^",4)) >= 50 And CLng(Parse(strRecord,"^",5)) = 1 Then
                    If CheckEndDate(Parse(strRecord,"^",3)) Then
                        Set oOption = Document.createElement("OPTION")
                        oOption.Value = oDictObj
                        oOption.Text = Parse(strRecord, "^", 1)
                        cboElement.options.add oOption
                        Set oOption = Nothing
                    End If
                End If
            Else
                If CLng(intPrg) = CLng(Parse(strRecord,"^",4)) And CLng(cboTab.value) = CLng(Parse(strRecord,"^",5)) Then
                    If CheckEndDate(Parse(strRecord,"^",3)) Then
                        Set oOption = Document.createElement("OPTION")
                        oOption.Value = oDictObj
                        oOption.Text = Parse(strRecord, "^", 1)
                        cboElement.options.add oOption
                        Set oOption = Nothing
                    End If
                End If
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

Sub ArrearageOption(strAction)
    Dim intI, intID, oOption, strRecord

    intID = -1
    For intI = 0 To cboElement.options.length - 1
        If CInt(cboElement.options(intI).value) = CInt(window.opener.mintArrearageID) Then
            intID = intI
            Exit For
        End If
    Next
    
    If strAction = "Remove" Then
        If intID > 0 Then
            cboElement.remove(intID)
        End If
    Else
        If intID = -1 Then
            strRecord = window.opener.mdctElements(window.opener.mintArrearageID)
            Set oOption = Document.createElement("OPTION")
            oOption.Value = window.opener.mintArrearageID
            oOption.Text = Parse(strRecord, "^", 1)
            cboElement.options.add oOption
            Set oOption = Nothing
        End If
   End If
End Sub

Function CheckEndDate(dtmEndDate)
    If dtmEndDate = "" Then
        CheckEndDate = True
        Exit Function
    End If
    
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

    strFactorList = Parse(window.opener.mdctElements(CLng(intElementID)),"^",7)
    
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
        ' If intRvwID = 1 (Full), check if currently checked and
        ' if it is currently checked, keep it checked. 
        strChecked = " "
        If blnSelectAllChecked = False Then
            If intRvwID = 1 Then
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

Sub FillStaffCbo(cboCtl)
    Dim intI
    Dim strRoles
    Dim dteTxtStart
    Dim dteTxtEnd
    Dim dteEmpStart
    Dim dteEmpEnd
    Dim strParentIDs
    Dim strParentPosID
    Dim strOptions
    Dim strRebuild
    Dim strOnXEvents
    Dim intCnt
    Dim strOuterHTML
    Dim strItem
    Dim strKey
    Dim strRegionFilter
    Dim dteItemStartDate
    Dim dteItemEndDate

    ' When coming back from viewing a report, the text boxes may not have been reset.
    ' If they need to be reset, do it now.
    If IsDate(Form.StartDate.value) And txtStartDate.value = "" Then txtStartDate.value = Form.StartDate.value
    If IsDate(Form.EndDate.value) And txtEndDate.value = "" Then txtEndDate.value = Form.EndDate.value

    intCnt = 1
    If txtStartDate.value = "" Then
        dteTxtStart = CDate("1/1/1900")
    Else
        dteTxtStart = CDate(txtStartDate.Value)
    End If
    If txtEndDate.value = "" Then
        dteTxtEnd = Now
    Else
        dteTxtEnd = CDate(txtEndDate.Value)
    End If
    strOptions = "<OPTION Value = 0>" & mstrCboBoxDefaultTextHTML & "</OPTION>"
    If cboCtl.Id = "cboDirector" Then
        For Each strItem In mdctDirectors
            strOptions = strOptions & "<OPTION value=""" & strItem & """>" & strItem & "</OPTION>"
        Next
    ElseIf cboCtl.Id = "cboOffice" Then
        For Each strItem In mdctOffices
            If (Parse(strItem, "^", 2) = cboDirector.value Or cboDirector.value = "0") And Parse(strItem, "^", 2) <> "State Of Missouri" Then
                strOptions = strOptions & "<OPTION value=""" & Parse(strItem, "^", 1) & """>" & Parse(strItem, "^", 1) & "</OPTION>"
            End If
        Next
    ElseIf cboCtl.Id = "cboManager" Then
        For Each strItem In mdctManagers
            'If (Parse(strItem, "^", 2) = cboOffice.value Or cboOffice.value = "0") Then
            If InStr(cboOffice.innerHTML,Parse(strItem, "^", 2)) > 0 Then
                strOptions = strOptions & "<OPTION value=""" & Parse(strItem, "^", 1) & """>" & Parse(strItem, "^", 1) & "</OPTION>"
            End If
        Next
    End If
    
    strOuterHTML = cboCtl.outerHTML
    intI = InStr(strOuterHTML,">")
    Select Case cboCtl.ID
        Case "cboReviewer"
            strOnXEvents = " onkeypress=" & Chr(34) & cboCtl.ID & "_onkeypress" & Chr(34)
        Case Else
            strOnXEvents = " onchange=" & Chr(34) & cboCtl.ID & "_onchange" & Chr(34) & _
                           " onkeypress=" & Chr(34) & cboCtl.ID & "_onkeypress" & Chr(34)
    End Select

    strRebuild = Left(strOuterHTML, intI - 1) & strOnXEvents & ">" & strOptions & "</SELECT>"
    cboCtl.outerHTML = strRebuild
End Sub

Sub CheckOptionsLength()
    If cboDirector.options.length = 2 Then
        If "<%=glngAliasPosID%>" = "2" Or "<%=gblnUserAdmin%>" = "True" Then
            cboDirector.selectedIndex = 0
            cboDirector.disabled=False
        Else
            cboDirector.selectedIndex = 1
            cboDirector.disabled=True
        End If
    End If
    If cboOffice.options.length = 2 Then
        cboOffice.selectedIndex = 1
        cboOffice.disabled = True
    Else
        cboOffice.selectedIndex = 0
        cboOffice.disabled = False
    End If
    If cboManager.options.length = 2 Then
        cboManager.selectedIndex = 1
        cboManager.disabled = True
    Else
        cboManager.selectedIndex = 0
        cboManager.disabled = False
    End If
End Sub

Sub FillStaffCboLists()

    Form.onchangeFlag.value = "N"

    Call FillStaffCbo(cboDirector)
    Call SelectStaffCboValue(cboDirector, Form.Director.Value)

    Call FillStaffCbo(cboOffice)
    Call SelectStaffCboValue(cboOffice, Form.Office.Value)
    
    Call FillStaffCbo(cboManager)
    Call SelectStaffCboValue(cboManager, Form.ProgramManager.Value)
    
    Form.onchangeFlag.value = "Y"
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

Sub txtReReviewer_onfocus()
    mstrHoldStaff = txtReReviewer.value
End Sub

Sub txtReReviewer_onblur()
    If txtReReviewer.value = "" Then
        Call ClearStaffingFields("RR")
    Else
        If txtReReviewer.value <> mstrHoldStaff Then
            Call StaffLookUp(txtReReviewer)
        End If
    End If
End Sub

Sub txtReviewer_onfocus()
    mstrHoldStaff = txtReviewer.value
End Sub

Sub txtReviewer_onblur()
    If txtReviewer.value = "" Then
        Call ClearStaffingFields("R")
    Else
        If txtReviewer.value <> mstrHoldStaff Then
            Call StaffLookUp(txtReviewer)
        End If
    End If
End Sub

Sub txtSupervisor_onfocus()
    mstrHoldStaff = txtSupervisor.value
End Sub

Sub txtSupervisor_onblur()
    If txtSupervisor.value = "" Then
        Call ClearStaffingFields("S")
        If mstrHoldStaff <> "<All>" Then
            Call ClearStaffingFields("W")
        End If
    Else
        If txtSupervisor.value <> mstrHoldStaff Then
            Call StaffLookUp(txtSupervisor)
        End If
    End If
End Sub

Sub txtWorker_onfocus()
    mstrHoldStaff = txtWorker.value
End Sub

Sub txtWorker_onblur()
    If txtWorker.value = "" Then
        Call ClearStaffingFields("W")
    Else
        If txtWorker.value <> mstrHoldStaff Then
            Call StaffLookUp(txtWorker)
        End If
    End If
End Sub

Sub Gen_onkeydown(ctlFrom)
    If window.event.keyCode = 13 Then
        Call StaffLookUp(ctlFrom)
    End If
End Sub

Sub StaffLookUp(ctlStaffID)
    Dim strType
    Dim strID
    Dim strSupervisor, strManager, strOffice, strDirector
    
    <%'Attempt to select the reviewer for the passed ID:%>
    <%'Fill in the reviewer from ID of the logged in user:%>
    If Len(ctlStaffID.value) > 0 And ctlStaffID.value <> mstrHoldStaff Then
        Form.StaffInformation.value = ""
        divStaffSearch.style.top = 150
        strID = ctlStaffID.value
        Set mctlStaff = ctlStaffID
        Select Case ctlStaffID.ID
            Case "txtWorker"
                strType = "txtWorker"
                Select Case Parse(lstReports.value,":",1)
                    Case "64"
                        divStaffSearch.style.top = 60
                        strType = "txtEmployee"
                    Case "118","139","140"
                        strType = "txtArcWorker"
                        divStaffSearch.style.top = 240
                    Case Else
                        divStaffSearch.style.top = 240
                End Select
                divStaffSearch.style.left = 15
            Case "txtSupervisor"
                strType = "txtSupervisor"
                divStaffSearch.style.top = 195
                divStaffSearch.style.left = 15
            Case "txtReviewer"
                strType = "txtReviewer"
                'Select Case Parse(lstReports.value,":",1)
                '    Case "35","32","33","34"
                '        divStaffSearch.style.top = 150
                '    Case "111","56","55","110"
                        divStaffSearch.style.top = 195
                '    Case Else
                '        divStaffSearch.style.top = 240
                ''End Select
                divStaffSearch.style.left = 15
            Case "txtReReviewer"
                strType = "txtReReviewer"
                divStaffSearch.style.top = 195
                divStaffSearch.style.left = 15
        End Select
        If strID = "?" Then strID = "%"
        If ctlStaffID.ID = "txtWorker" Then
            If txtSupervisor.value = "<All>" Then
                strSupervisor = ""
            Else
                strSupervisor = Trim(txtSupervisor.value)
            End If
        Else
            strSupervisor = ""
            strManager = ""
        End If
        'strManager = Trim(txtProgramManager.value)
        strOffice = cboOffice.value
        strDirector = cboDirector.value
        strManager = cboManager.value
        If strType = "txtEmployee" Then
            'Remove all filters when searching for employees
            strManager = ""
            strOffice = ""
            strDirector = ""
            strSupervisor = ""
        End If
        fraStaffSearch.frameElement.src = "ReportsStaffSearch.asp?" & _
            "AliasID=<%=glngAliasPosID%>" & _
            "&UserAdmin=<%=gblnUserAdmin%>" & _
            "&UserQA=<%=gblnUserQA%>" & _
            "&UserID=<%=gstrUserID%>" & _
            "&Type=" & strType & _
            "&StaffName=" & strID & _
            "&Width=314" & _
            "&Supervisor=" & strSupervisor & _
            "&Manager=" & strManager & _
            "&Office=" & strOffice & _
            "&Director=" & strDirector & _
            "&StartDate=" & txtStartDate.value & _
            "&EndDate=" & txtEndDate.value
            
    End If
End Sub

Sub StaffLookUpClose(strStaffInfo)
    Dim strName, dtmStartDate, dtmEndDate
    
    strName = Parse(strStaffInfo, "^", 1)
    dtmStartDate = Parse(strStaffInfo, "^", 2)
    dtmEndDate = Parse(strStaffInfo, "^", 3)
    divStaffSearch.style.left = -1000
    If strStaffInfo = "[CANCEL]" Then
        mctlStaff.value = mstrHoldStaff
    Else
        If strName <> "no matches [Close]" Then
            mctlStaff.value = strName
            If mctlStaff.ID = "txtSupervisor" Then
                txtWorker.value = "<All>"
            End If
        Else
            mctlStaff.value = "<All>"
            If mctlStaff.value <> mstrHoldStaff And mctlStaff.ID = "txtSupervisor" Then
                txtWorker.value = "<All>"
            End If
        End If
    End If
    
    document.all(mctlStaff.ID & "StartDate").value = dtmStartDate
    document.all(mctlStaff.ID & "EndDate").value = dtmEndDate
    If mctlStaff.disabled = False Then mctlStaff.focus
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
        
        <DIV id=divReportMode class=DefPageFrame style="LEFT:-1425; HEIGHT:20; WIDTH:305; TOP:10; border:none">
            <INPUT type=radio ID=optReportMode2 name="optReportMode" onclick="optReportMode_onclick(1)" style="LEFT:0" CHECKED tabIndex=11 VALUE="optReportMode2">
            <SPAN id=lblReportMode2 class=DefLabel style="LEFT:25; WIDTH:120"
                onclick="lblReportMode_onclick(0)">
                <%=gstrAuthTitle%>
            </SPAN>
            
            <INPUT type=radio ID=optReportMode1 name="optReportMode" onclick="optReportMode_onclick(0)" style="LEFT:155" tabIndex=11 VALUE="optReportMode1">
            <SPAN id=lblReportMode1 class=DefLabel style="LEFT:180; WIDTH:120"
                onclick="lblReportMode_onclick(1)">
                <%=gstrWkrTitle%>
            </SPAN>
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
                    Case 0
                        'These reports are only visible if the user has 
                        'the Quality Assurance security role, role 21:
                        If Instr(gstrRoles, "[1]") > 0 Then
                            mblnShowReport = True
                        Else
                            mblnShowReport = False
                        End If
                    Case 118,139,140
                        mblnShowReport = True
                    Case Else
                        mblnShowReport = False
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
    <DIV id=divStaffSearch class=ControlDiv
        style="width:320;height:150;left:-1000;top:40;z-index:101">
        <IFRAME id=fraStaffSearch tabindex=-1 style="width:318;height:148;left:1;top:1"></IFRAME>
    </DIV>
        <DIV ID=CriteriaFrame class=DefPageFrame 
            style="Left:5; HEIGHT:340; Width:448;Top:20; border-style:none; visibility:hidden">

            <SPAN id=lblDirector class=DefLabel style="LEFT:10; WIDTH:200; TOP:5">
                <%=gstrDirTitle%>
            </SPAN>
            <SELECT id=cboDirector style="LEFT:10; WIDTH:200; TOP:20" tabIndex=<%=GetTabIndex%> NAME="cboDirector">
            </SELECT>

            <SPAN id=lblOffice class=DefLabel style="LEFT:10; WIDTH:200; TOP:50">
                <%=gstrOffTitle%>
            </SPAN>
            <SELECT id=cboOffice style="LEFT:10; WIDTH:200; TOP:65" tabIndex=<%=GetTabIndex%> NAME="cboOffice">
            </SELECT>
            
            <SPAN id=lblManager class=DefLabel style="LEFT:10; WIDTH:200; TOP:95">
                <%=gstrMgrTitle%>
            </SPAN>
            <SELECT id=cboManager style="LEFT:10; WIDTH:200; TOP:110" tabIndex=<%=GetTabIndex%> NAME="cboManager">
            </SELECT>
            
            <SPAN id=lblSupervisor class=DefLabel title="Enter ? To Return All" style="LEFT:10; WIDTH:200; TOP:140">
                <%=gstrSupTitle%>
            </SPAN>
            <INPUT type="text" ID=txtSupervisor NAME="txtSupervisor" title="Enter ? To Return All" 
                onkeydown="Gen_onkeydown(txtSupervisor)"  tabIndex=<%=GetTabIndex%>
                onfocus="CmnTxt_onfocus(txtSupervisor)" maxlength=50
                style="LEFT:10; WIDTH:200; TOP:155;text-align:left">
            <input type=hidden id=txtSupervisorStartDate>
            <input type=hidden id=txtSupervisorEndDate>
            
            <SPAN id=lblReviewer class=DefLabel title="Enter ? To Return All" style="LEFT:10; WIDTH:200; TOP:140">
                <%=gstrRvwTitle%>
            </SPAN>
            <INPUT type="text" ID=txtReviewer NAME="txtReviewer" title="Enter ? To Return All" 
                onkeydown="Gen_onkeydown(txtReviewer)"  tabIndex=<%=GetTabIndex%>
                onfocus="CmnTxt_onfocus(txtReviewer)" maxlength=50
                style="LEFT:10; WIDTH:200; TOP:155;text-align:left">
            <input type=hidden id=txtReviewerStartDate>
            <input type=hidden id=txtReviewerEndDate>
            
            <SPAN id=lblWorker class=DefLabel title="Enter ? To Return All" style="LEFT:10; WIDTH:200; TOP:185">
                <%=gstrWkrTitle%>
            </SPAN>
            <INPUT type="text" ID=txtWorker NAME="txtWorker" title="Enter ? To Return All" 
                onkeydown="Gen_onkeydown(txtWorker)"  tabIndex=<%=GetTabIndex%>
                onfocus="CmnTxt_onfocus(txtWorker)" maxlength=50
                style="LEFT:10; WIDTH:200; TOP:200;text-align:left">
            <input type=hidden id=txtWorkerStartDate>
            <input type=hidden id=txtWorkerEndDate>

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
                style="Left:222;HEIGHT:50;Width:215;Top:107;border-style:none;visibility:hidden;background-color:transparent;z-index:1000">
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
                Functions
            </SPAN>
            <INPUT type="hidden" id=txtGroupID NAME="txtGroupID">
            <SELECT id=lstProgram 
                style="z-index:-1; LEFT:235; WIDTH:200; Top:20; OVERFLOW:auto; visibility:visible"
                onchange="cboProgram_onchange"
                tabindex=<%=GetTabIndex%> NAME="lstProgram">
            </SELECT>
            
            <SPAN id=lblEligElement 
                class=DefLabel
                Style="LEFT:235; TOP:95; Width:200">
                Screens
            </SPAN>
            <SELECT id=cboElement title="Screen"
                style="z-index:-1; LEFT:235; TOP:110; WIDTH:200; OVERFLOWL:auto; visibility:hidden"
                onchange=""
                tabIndex=<%=GetTabIndex%> NAME="cboElement">
            </SELECT>
            
            <SPAN id=lblTab
                class=DefLabel
                style="Left:235; Top:50; Width:200">
                Tab
            </SPAN>
            <SELECT id=cboTab
                style="z-index:-1; LEFT:235; TOP:65; Width:200; OVERFLOW:auto; visibility:hidden"    
                tabIndex=<%=GetTabIndex%> NAME="cboTab">
                <OPTION value=0 selected>All</OPTION>
                <OPTION value=1>Action Integrity</OPTION>
                <OPTION value=2>Data Integrity</OPTION>
                <OPTION value=3>Information Gathering</OPTION>
                
            </SELECT>

            <SPAN id=lblFactor 
                class=DefLabel
                Style="LEFT:235; TOP:140; Width:200">
                Field Name
            </SPAN>
            <SELECT id=cboFactor title="Field Name"
                style="z-index:-1; LEFT:235; TOP:155; WIDTH:200; OVERFLOWL:auto; visibility:hidden"
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
            
            <SPAN id=lblErrorDiscovery
                class=DefLabel
                style="Left:235; Top:95; Width:200">
                Discovery
            </SPAN>
            <SELECT id=cboErrorDiscovery
                style="z-index:-1; LEFT:235; TOP:110; Width:200; OVERFLOW:auto; visibility:hidden"
                tabIndex=<%=GetTabIndex%> NAME="cboErrorDiscovery">
                <OPTION value=0 selected></OPTION>
                <%=BuildList("ErrorDiscovery","",0,0,0)%>
            </SELECT>
            <SPAN id=lblBenErrorType
                class=DefLabel
                style="Left:235; Top:140; Width:200">
                Benefit Error Type
            </SPAN>
            <SELECT id=cboBenErrorType
                style="z-index:-1; LEFT:235; TOP:155; Width:200; Height:100; OVERFLOW:auto; visibility:hidden"
                tabIndex=<%=GetTabIndex%> NAME="BenErrorType">
                <OPTION value=0 selected></OPTION>
                <%=BuildList("BenErrorType","",0,0,0)%>
            </SELECT>
            
            <SPAN id=lblCompliance class=DefLabel style="LEFT:235; TOP:5">
                Display Compliance
            </SPAN>
            <SELECT id=cboCompliance
                style="z-index:-1; WIDTH:200; LEFT:235; TOP:20; OVERFLOW:auto; visibility:hidden" 
                tabIndex=-1 NAME="cboCompliance">
                <OPTION VALUE=0 SELECTED>Display All (Compliant and non-compliant)</option>
                <OPTION value=1>Display Non-Compliant Only</option>
            </SELECT>
            
            <SPAN id=lblHousehold class=DefLabel style="LEFT:235; WIDTH:46; TOP:95">
                Household
            </SPAN>
            <SELECT id=cboHouseHold title="Household Type" 
                style="z-index:-1; LEFT:235; WIDTH:200; TOP:110; OVERFLOW:auto; visibility:hidden" 
                onchange=""
                tabIndex=11 NAME="cboHouseHold">
                <OPTION VALUE=0></option>
                <OPTION VALUE=1>Single Parent Household</option>
                <OPTION VALUE=2>Two Parent Household</option>
            </SELECT>

            <SPAN id=lblPartHours class=DefLabel style="LEFT:-1235; WIDTH:120; TOP:140">
                Participation Hours
            </SPAN>
            <SELECT id=cboPartHours title="Status of Participation Hours" 
                style="z-index:-1; LEFT:-1235; WIDTH:200; TOP:175; Height:75; OVERFLOW:auto; visibility:hidden" 
                onclick="cbo_onclick lstPartHours, txtPartHours"
                onchange=""
                tabIndex=111 size=2 NAME="lstPartHours">
                <OPTION VALUE=0></option>
                <OPTION VALUE=214>02 - State hours not met</option>
                <OPTION VALUE=215>03 - Federal hours not met</option>
                <OPTION VALUE=216>04 - Federal hours not met/State hours met</option>
                <OPTION VALUE=295>05 - Federal hours not met/State hours not met</option>
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
    lblWorker.style.top = 185
    txtWorker.style.top = 200
        
    lblDirector.style.visibility="hidden"
    cboDirector.style.visibility="hidden"
    lblOffice.style.visibility="hidden"
    cboOffice.style.visibility="hidden"
    lblManager.style.visibility="hidden"
    cboManager.style.visibility="hidden"
    lblSupervisor.style.visibility="hidden"
    txtSupervisor.style.visibility="hidden"
    lblWorker.style.visibility="hidden"
    txtWorker.style.visibility="hidden"
    lblReviewer.style.visibility="hidden"
    txtReviewer.style.visibility="hidden"
    lblReReviewer.style.visibility="hidden"
    txtReReviewer.style.visibility="hidden"
    
    lblProgram.style.visibility="hidden"
    lstProgram.style.visibility="hidden"
    lblEligElement.style.top = 50
    cboElement.style.top = 65
    lblEligElement.innerText="Screen"
    lblEligElement.style.visibility="hidden"
    cboElement.style.visibility="hidden"
    divSupVsReviewer.style.left = -1000
    lblSupVsReviewer.style.left = -1000
    lblFactor.style.top = 95
    cboFactor.style.top = 110
    lblFactor.innerText="Field Name"
    lblFactor.style.visibility="hidden"
    cboFactor.style.visibility="hidden"
        
    lblCaseAction.style.visibility="hidden"
    cboCaseAction.style.visibility="hidden"
    lblErrorDiscovery.style.visibility="hidden"
    cboErrorDiscovery.style.visibility="hidden"
    lblBenErrorType.style.visibility="hidden"
    cboBenErrorType.style.visibility="hidden"
    lblCompliance.style.visibility="hidden"
    cboCompliance.style.visibility="hidden"
    lblSubmitted.style.visibility="hidden"
    cboSubmitted.style.visibility="hidden"
    lblCaseNumber.style.visibility="hidden"
    txtCaseNumber.style.visibility="hidden"
    lblDetail.style.visibility="hidden"
    lblInclude.style.visibility="hidden"
    lblMinDays.style.visibility="hidden"
    txtMinDays.style.visibility="hidden"
    lblHousehold.style.visibility="hidden"
    cboHousehold.style.visibility="hidden"
    lblPartHours.style.visibility="hidden"
    cboPartHours.style.visibility="hidden"
    lblResponse.style.visibility="hidden"
    cboResponse.style.visibility="hidden"
    lblDaysPastDue.style.visibility="hidden"
    txtDaysPastDue.style.visibility="hidden"
    
    divReviewType.style.visibility="hidden"
    lblReviewType.style.visibility="hidden"
    divReviewTypeDefs.style.visibility="hidden"
    lblReviewer.innerText="Reviewer"
    divReviewMonth.style.visibility = "hidden"
End Sub

'Display the appropriate Criteria for the selected Report
Sub Display_Criteria()
    Dim strReport
    Dim intI
    Dim intLength
    
    strReport = Parse(lstReports.value, ":", 1)
    Call Hide_Criteria()
    
    optReportMode1.disabled = False
    optReportMode2.disabled = False
    
    If strReport <> 64 And lblWorker.innerText = "Employee" Then
        txtWorker.value = "<All>"
        txtWorkerStartDate.value = ""
        txtWorkerEndDate.value = ""
    ElseIf strReport = 64 Then
        txtWorker.value = "<All>"
        txtWorkerStartDate.value = ""
        txtWorkerEndDate.value = ""
    End If

    mblnUseAuthBy = False
    
    lblTab.style.top = 50
    cboTab.style.top = 65
    lblTab.style.visibility = "hidden"
    cboTab.style.visibility = "hidden"
    
	divReviewType.style.left = 10
    chkDetail.checked = False
    Select Case strReport
        Case 139,118, 140 'Case Review Detail
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
            lblCaseNumber.style.visibility="visible"
            txtCaseNumber.style.visibility="visible"
            lblProgram.style.visibility="visible"
            lstProgram.style.visibility="visible"
            If lstProgram.selectedIndex = 0 Then
                lstProgram.selectedIndex = 1
            End If
            If strReport = 139 Then
                chkDetail.checked = False
            Else
                chkDetail.checked = True
            End If
            Call ListReviewTyps(-1)
            divReviewMonth.style.visibility = "visible"
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
        txtWorker.style.visibility="hidden"
        lblSupervisor.style.visibility="hidden"
        txtSupervisor.style.visibility="hidden"
        lblReviewer.style.visibility="visible"
        txtReviewer.style.visibility="visible"
    Else
        lblWorker.style.visibility="visible"
        txtWorker.style.visibility="visible"
        lblSupervisor.style.visibility="visible"
        txtSupervisor.style.visibility="visible"
        lblReviewer.style.visibility="hidden"
        txtReviewer.style.visibility="hidden"
    End If
End Sub

Sub SetMgrSupWkrAbyRT(mblnUseAuthby)
    lblDirector.style.visibility="visible"
    cboDirector.style.visibility="visible"
    
    lblOffice.style.visibility="visible"
    cboOffice.style.visibility="visible"
    lblManager.style.visibility="visible"
    cboManager.style.visibility="visible"
    lblSupervisor.style.visibility="visible"
    lblWorker.style.visibility="visible"
    
    divReviewType.style.visibility="visible"
    lblReviewType.style.visibility="visible"
    divReviewTypeDefs.style.visibility="visible"
    
    txtSupervisor.style.visibility="visible"
    txtWorker.style.visibility="visible"
    
    lblWorker.innerText = "<%=gstrWkrTitle%>"
    divReviewType.style.visibility="visible"
    lblReviewType.style.visibility="visible"
    divReviewTypeDefs.style.visibility="visible"
End Sub

Sub SetMgrRvw()
    lblDirector.style.visibility="visible"
    cboDirector.style.visibility="visible"

    lblOffice.style.visibility="visible"
    cboOffice.style.visibility="visible"
    lblManager.style.visibility="visible"
    cboManager.style.visibility="visible"
    
    lblReviewer.style.visibility="visible"
    txtReviewer.style.visibility="visible"
End Sub

'Clears all criteria
Sub cmdClearCriteria_onclick()
    Dim intI
    Dim oCtl
    
    ' If any of the combos that filter down are not 0, rebuild them
    If cboDirector.value <> "0" Then
        cboDirector.value = "0"
        Call StaffComboOnChange(4)
    ElseIf cboOffice.value <> "0" Then
        cboOffice.value = "0"
        Call StaffComboOnChange(3)
    ElseIf cboManager.value <> "0" Then
        cboManager.value = "0"
        Call StaffComboOnChange(2)
    End If
    Call ClearStaffingFields("All")
    Call CheckOptionsLength()
    cboTab.value = 0
    cboErrorDiscovery.value = 0
    cboCompliance.value = 0
    If lstProgram.options.length > 2 Then lstProgram.value = 0
    cboElement.value = 0
    
    txtDaysPastDue.value = ""
    
    Set oCtl = Nothing
    cboSubmitted.value = 0
    If chkDetail.checked Then
        chkDetail.checked = False
    End If
    If chkInclude.checked Then chkInclude.checked = False
    txtCaseNumber.value = ""
    txtMinDays.value = 0
    cboHousehold.value = 0
    cboPartHours.value = 0
    cboResponse.value = 0
    cboSubmitted.value = 0
    txtStartDate.value = ""
    txtEndDate.value = ""
    txtStartReviewMonth.value = ""
    txtEndReviewMonth.value = ""
    
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

    If Form.ReportingMode.value = "" Then Form.ReportingMode.value = "1"
    If Form.ReportingMode.value <> "1" Then
        Form.ReportingMode.value = "0"
        mblnUseAuthby = False
        lblReportMode_onclick(0)
    Else
        Form.ReportingMode.value = "1"
        mblnUseAuthby = True
        lblReportMode_onclick(1)
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
    
    If Trim(Form.HouseholdParents.value) = "" Then
        cboHousehold.value = 0
    Else
        cboHousehold.value = Trim(Form.HouseholdParents.value)
    End If

    If Trim(Form.PartHours.value) = "" Then
        cboPartHours.value = 0
    Else
        cboPartHours.value = Trim(Form.PartHours.value)
    End If
    
    If Trim(Form.DiscoveryID.value) = "" Then
        cboErrorDiscovery.value = 0
    Else
        cboErrorDiscovery.value = Form.DiscoveryID.value
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
        If cboCtl.Id = "cboDirector" Then
            cboCtl.value = ""
            If IsNull(cboCtl.value) Or cboCtl.value = "" Then
                'Cbo still not set, so select the default <All> entry:
                cboCtl.value = 0
            End If
        ElseIf cboCtl.Id = "cboOffice" Then
            'Cbo still not set, so select the default <All> entry:
            cboCtl.value = 0
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

    strAllowGroup = "[26][52][57]"
    strSpecificProg = "[28][29][30][31][45][49][50][51][52][53][55][56][72][73][131][132][133][134][135][136][137]"
    strExReviewType = "[32][33][34][35][36][47][55][56][72][73]"
    strExReviewClass = "[33][35][47][55][56][72][73]"
    Call ClearForms()
    
    'Edit Criteria only if viewing a report, not if saving a date change
    If blnEditCriteria Then
        If Not IsNumeric(txtGroupID.value) Then txtGroupID.value = -1
        If lstProgram.value = 0 Then txtGroupID.value = -1
        If Instr(strAllowGroup, Parse(lstReports.value, ":", 1)) = 0 Then
            If lstProgram.value < 0 And txtGroupID.value > 0 And Instr(strSpecificProg, Parse(lstReports.value, ":", 1)) > 0 Then
                MsgBox "Please select a Specific Function", vbInformation, "View Report"
                lstProgram.focus
                Exit Function
            End If
        End If
        If InStr(strSpecificProg, "[" & Parse(lstReports.value, ":", 1) & "]") > 0 Then
            If lstProgram.value <= 0 And (txtGroupID.value <= 0 Or Instr(strAllowGroup, Parse(lstReports.value, ":", 1)) = 0) Then
                If Instr(strAllowGroup, Parse(lstReports.value, ":", 1)) = 0 Then
                    MsgBox "Please select a Specific Function", vbInformation, "View Report"
                Else
                    MsgBox "Please select a Specific Function or Function Group", vbInformation, "View Report"
                End If
                lstProgram.focus
                Exit Function
            End If
        End If

        If Parse(lstReports.value, ":", 1) = "118" Or Parse(lstReports.value, ":", 1) = "140" Then
            If txtWorker.Value = "<All>" Then
                MsgBox "This report is designed to be run by Worker." & vbCrlf & vbCrLf & "Please Select a Worker." & Space(5), vbInformation, "View Report"
                txtWorker.focus
                Exit Function
            End If
        End If
        
        'If (Parse(lstReports.value, ":", 1) = "131" Or Parse(lstReports.value, ":", 1) = "132") And lstProgram.value >= 50 Then
        If lstProgram.value >= 50 Or lstProgram.value = 6 Then
            chkCheckAllRT.checked = False
            For intI = 0 To lblReviewCount.innerText - 1
                document.all("chkReviewType" & intI).checked = False
            Next
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

    Form.TabID.Value = cboTab.value
    Form.TabName.Value = GetComboText(cboTab) '.options(cboTab.selectedIndex).text

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
    
    Form.Worker.value = txtWorker.value
    Form.Supervisor.value = txtSupervisor.value
    
    Form.Reviewer.Value = txtReviewer.value
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
    If optReportMode1.checked Then
        Form.ReportingMode.value = 0
    Else
        Form.ReportingMode.value = 1
    End If
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
           Form.DiscoveryText.value = Replace(Form.DiscoveryText.value,mstrCboBoxDefaultText,"")
           Form.ResponseText.value = Replace(Form.ResponseText.value,mstrCboBoxDefaultText,"")
           Form.HouseholdText.value = Replace(Form.HouseholdText.value,mstrCboBoxDefaultText,"")
           Form.TabName.value = Replace(Form.TabName.value,mstrCboBoxDefaultText,"")
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
    Form.DiscoveryID.Value = 0
    Form.DiscoveryText.Value = ""
    Form.HouseholdParents.value = 0
    Form.HouseholdText.value = ""
    Form.PartHours.value = 0
    Form.PartHoursText.value = ""
    Form.CaseNumber.Value = ""
    Form.MinAvgDays.Value = 0
    Form.ShowNonComOnly.Value = ""
    Form.ShowNonComOnlyID.Value = 0
    Form.Submitted.Value = ""
    Form.ShowDetail.value = ""
    Form.IncludeCorrect.value = ""
    Form.ReportingMode.value = 0
    Form.CountyOfficesID.value = 0
    Form.CountyOfficesText.value = ""
    Form.WTWReasonID.value = 0
    Form.WTWReasonText.value = ""
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
    WriteFormField "DiscoveryID", ReqForm("DiscoveryID")
    WriteFormField "DiscoveryText", ReqForm("DiscoveryText")
    WriteFormField "ResponseID", ReqForm("ResponseID")
    WriteFormField "ResponseText", ReqForm("ResponseText")
    WriteFormField "HouseholdParents", ReqForm("HouseholdParents")
    WriteFormField "HouseholdText", ReqForm("HouseholdText")
    WriteFormField "PartHours", ReqForm("PartHours")
    WriteFormField "PartHoursText", ReqForm("PartHoursText")
    WriteFormField "CaseNumber", ReqForm("CaseNumber")
    WriteFormField "MinAvgDays", ReqForm("MinAvgDays")
    WriteFormField "ShowNonComOnlyID", ReqForm("ShowNonComOnlyID")
    WriteFormField "ShowNonComOnly", ReqForm("ShowNonComOnly")
    WriteFormField "Submitted", ReqForm("Submitted")
    WriteFormField "ShowDetail", ReqForm("ShowDetail")
    WriteFormField "IncludeCorrect", ReqForm("IncludeCorrect")
    WriteFormField "CountyOfficesID", ReqForm("CountyOfficesID")
    WriteFormField "CountyOfficesText", ReqForm("CountyOfficesText")
    WriteFormField "WTWReasonID", ReqForm("WTWReasonID")
    WriteFormField "WTWReasonText", ReqForm("WTWReasonText")
    WriteFormField "DaysPastDue", ReqForm("DaysPastDue")
    WriteFormField "Under26", ReqForm("Under26")
    WriteFormField "StartReviewMonth", ReqForm("StartReviewMonth")
    WriteFormField "EndReviewMonth", ReqForm("EndReviewMonth")
    WriteFormField "CTR", "0" 'ReqForm("CTR")
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
    
    If Len(lblReviewCount.innerText) = 0 Then Exit Sub

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
