<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: Main.asp                                                        '
'  Purpose: The main menu or switchboard for the application.               '
'==========================================================================='
Dim mstrPageTitle   'Sets the title at the top of the form.
Dim madoCmd         'ADO command object used for this page.
Dim mlngRespDueRvws 'Indicates whether to alert for worker response past due.
Dim mstrTmp         'Temporary string holder for building prompts, etc.
Dim mstrPrgSelected 'The programs last selected by the user.
Dim mintLeft
Dim mintTop
Dim strResizeScreen	'Holds value of Screen Resize flag.
Dim mdctAliasIDs, mdctLinks, mdctElements, mdctFactors, moDictObj
Dim madoRs, mdctElementIDs, mstrUserType
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<%
mstrPageTitle = Trim(gstrTitle & " " & gstrAppName)

If gstrUserID <> vbNullstring Then
    'If the user entered a new password, update the user record:
    If request.Form("NewPassword") <> vbNullstring Then
        Set madoCmd = Server.CreateObject("ADODB.Command")
        With madoCmd
            .CommandType = adCmdStoredProc 
            .CommandText = "spChangePassword"
            Set .ActiveConnection = gadoCon
            .Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, 60, Encrypt(LCase(Request.Form("NewPassword")), UCase(Request.Form("UserID"))))
            .Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 20, Request.Form("UserID"))
            .Execute
        End With
        Set madoCmd = Nothing
        gstrPassword = Encrypt(Request.Form("NewPassword"), UCase(Request.Form("UserID")))
    End If
End If
Set mdctAliasIDs = CreateObject("Scripting.Dictionary")

Set madoCmd = GetAdoCmd("spGetAlaisIDs")
    AddParmIn madoCmd, "@ID", adInteger, 0, Null
    AddParmIn madoCmd, "@TypeID", adInteger, 0, Null
    AddParmIn madoCmd, "@ParentID", adInteger, 0, Null
    AddParmIn madoCmd, "@Name", adVarchar, 20, Null
Set madoRs = Server.CreateObject("ADODB.Recordset")
Call madoRs.Open(madoCmd, , adOpenForwardOnly, adLockReadOnly)
madoRs.Sort = "alsName"
Do While Not madoRs.Eof
    If mdctAliasIDs.Exists(CStr(madoRs.Fields("alsID").value)) Then
        mdctAliasIDs(CStr(madoRs.Fields("alsID").value)) = mdctAliasIDs(CStr(madoRs.Fields("alsID").value)) & madoRs.Fields("ParentID").value & "*"
    Else
        mdctAliasIDs.Add CStr(madoRs.Fields("alsID").value), madoRs.Fields("alsTypeID").value & "^" & madoRs.Fields("alsName").value & "^" & madoRs.Fields("ParentID").value & "*"
    End If
    madoRs.MoveNext
Loop
madoRs.Close

Set mdctElements = CreateObject("Scripting.Dictionary")
Set mdctElementIDs = CreateObject("Scripting.Dictionary")
Set mdctFactors = CreateObject("Scripting.Dictionary")
Set madoCmd = GetAdoCmd("spElementsFactors")
    AddParmIn madoCmd, "@PrgID", adInteger, 0, Null
Set madoRs = Server.CreateObject("ADODB.Recordset")
Call madoRs.Open(madoCmd, , adOpenForwardOnly, adLockReadOnly)
madoRs.Filter = "elmStartDate<='" & FormatDateTime(Now(),2) & "'"
Do While Not madoRs.Eof
    If Not mdctElementIDs.Exists(madoRs.Fields("ProgramID").value & "^" & madoRs.Fields("ElementName").value) Then
        mdctElementIDs.Add madoRs.Fields("ProgramID").value & "^" & madoRs.Fields("ElementName").value, madoRs.Fields("ElementID").value
    End If
    If Not mdctElements.Exists(CLng(madoRs.Fields("ElementID").value)) Then
        mdctElements.Add CLng(madoRs.Fields("ElementID").value), madoRs.Fields("ElementName").value & "^" & _
            madoRs.Fields("ElementDescr").value & "^" & _
            madoRs.Fields("ElmEndDate").value & "^" & _
            madoRs.Fields("ProgramID").value & "^" & _
            madoRs.Fields("TabID").value & "^" & _
            madoRs.Fields("InFull").value & "^" & _
            madoRs.Fields("elmStartDate").value & "^"
    End If
    If CLng(madoRs.Fields("FactorID").value) > 0 Then
        If Not mdctFactors.Exists(CLng(madoRs.Fields("FactorID").value)) Then
            mdctFactors.Add CLng(madoRs.Fields("FactorID").value), madoRs.Fields("FactorName").value & "^" & _
                Replace(madoRs.Fields("FactorDescr").value,Chr(13) & Chr(10),"[vbCrLf]")
        End If
    End If
    
    mdctElements(CLng(madoRs.Fields("ElementID").value)) = mdctElements(CLng(madoRs.Fields("ElementID").value)) & madoRs.Fields("FactorID").value & "." & madoRs.Fields("ElfEndDate").value & "*"
    madoRs.MoveNext
Loop
If InStr(gstrRoles,"[6]") > 0 Then
    mstrUserType = "W" 'Worker
End If
If InStr(gstrRoles,"[2]") > 0 Then
    mstrUserType = "S" 'Sup
End If
If InStr(gstrRoles,"[1]") > 0 Then
    mstrUserType = "O" 'Office manager
End If
If InStr(gstrRoles,"[3]") > 0 Or InStr(gstrRoles,"[4]") > 0 Then
    mstrUserType = "A" 'Admin
End If
%>

<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Dim maProgramIDs(2)
Dim mdctWindows
Dim mlngTimerID
Dim mstrFeatures
Dim mstrLastPage
Dim mintWindow
Dim mdctWindowTable
Dim mdctReportWindows     ' Will hold all report windows opened
Dim mblnCloseClicked
Dim mdctRegions
Dim mdctDivisions
Dim mdctManagers
Dim mdctAliasIDs, mdctElements, mdctFactors, moDictObj, mdctElementIDs 
Dim mintArrearageID

Sub window_onload
    Dim intI

    ' Load array to hold Program IDs for the check boxes
    maProgramIDs(0) = 1  'FS
    maProgramIDs(1) = 2  'TAF
    maProgramIDs(2) = 3  'MC

    If Form.ResizeScreen.value = "Y" Then
	    Window.MoveTo 25, 25
        If "<%=Request.ServerVariables("SERVER_NAME")%>" = "secure.rushmore-group.com" Then
	        Window.ResizeTo 770, 560
        Else
	        Window.ResizeTo 770, 520
	    End If
    End If
    
    mblnCloseClicked = False
    Set mdctWindows = CreateObject("Scripting.Dictionary")
    Set mdctWindowTable = CreateObject("Scripting.Dictionary")
    Call LoadWindowTable()
    Set mdctReportWindows = CreateObject("Scripting.Dictionary")
    Set mdctRegions = CreateObject("Scripting.Dictionary")
    Set mdctDivisions = CreateObject("Scripting.Dictionary")
    Set mdctAliasIDs = CreateObject("Scripting.Dictionary")
    Set mdctElementIDs = CreateObject("Scripting.Dictionary")
    Set mdctElements = CreateObject("Scripting.Dictionary")
    Set mdctFactors = CreateObject("Scripting.Dictionary")
    Set mdctManagers = CreateObject("Scripting.Dictionary")
    
    Call LoadDictionaries
   
    mstrFeatures = "directories=no,fullscreen=no,location=no,menubar=no,status=no,resizable=yes,toolbar=no,left=1,top=1,height=495,width=760,scrollbars=yes"
    
    'If the user validation failed or the user is inactive, then display the 
    'warning and redirect the browser to the login screen.
    If Trim(Form.UserID.Value) = "" Then
        MsgBox "User not recognized.  Logon failed, please try again.", vbinformation, "Case Review Log On"
        Call ButtonClick(cmdLogOff)
    ElseIf UCase(Trim(Form.UserID.Value)) = "*INACTIVE*" Then
        MsgBox "User is not active.  Logon failed, please try again.", vbinformation, "Case Review Log On"
        Call ButtonClick(cmdLogOff)
    Else
        PageFrame.style.visibility = "visible"
    End If

    Call LoadReviewList("")
    <% '--- Server-side ---
    'If the NewPassword field is not blank, it means we are coming from the
    'login screen and the user has changed their password (handled by the
    'IncValidUser include file), so display the success notification:
    If request.Form("NewPassword") <> vbNullstring Then
        Response.Write "MsgBox ""Password has been changed."", vbInformation, ""Change Password""" & vbCrLf
    End If
    'If the Main form is being loaded by the login screen, check for the number
    'of reviews that have a response due, and if it is greater than zero then
    'give the user a message:
    If Request.Form("CalledFrom") = "Logon" Then
        Set madoCmd = Server.CreateObject("ADODB.Command")
        With madoCmd
            .CommandType = adCmdStoredProc 
            .CommandText = "spCheckForResponseDue"
            Set .ActiveConnection = gadoCon
            .Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 50, gstrUserID)
            .Parameters.Append .CreateParameter("@WorkerRspDue", adInteger, adParamOutput, 0, NULL)
            .Execute
            mlngRespDueRvws = .Parameters("@WorkerRspDue").Value
        End With
        Set madoCmd = Nothing
        'TODO list replaces the need for this message box.  The stored procedure above is still
        'called because it also truncates the audit table.
        
        'If mlngRespDueRvws > 0 Then
        '    If mlngRespDueRvws = 1 Then
        '        mstrTmp = " review "
        '    Else
        '        mstrTmp = " reviews "
        '    End If
        '    Response.Write "MsgBox ""You have " & mlngRespDueRvws & mstrTmp & "waiting for a " & gstrWkrTitle & " response."" & vbcrlf & vbcrlf & ""Please check the Response Due Report for a detailed list of reviews."" & Space(10), vbInformation, """ & gstrWkrTitle & " Response Past Due""" & vbCrLf
        'End If
    End If

    'Retrieve the last programs that were selected by the user:
    Set madoCmd = Server.CreateObject("ADODB.Command")
    With madoCmd
        .CommandType = adCmdStoredProc 
        .CommandText = "spProfileSettingGet"
        Set .ActiveConnection = gadoCon
        .Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 20, gstrUserID)
        .Parameters.Append .CreateParameter("@SettingName", adVarchar, adParamInput, 50, "ProgramsSelected")
        .Parameters.Append .CreateParameter("@SettingValue", adVarchar, adParamOutput, 255, NULL)
        .Execute
        mstrPrgSelected = .Parameters("@SettingValue").Value
    End With
    Set madoCmd = Nothing
    'Store selected programs string in the form field during the load event:
    Response.Write "Form.ProgramsSelected.value = """ & mstrPrgSelected & """"
    %>

    If Trim(Form.ProgramsSelected.value) <> "" Then
        For intI = 0 To UBound(maProgramIDs)
            If Instr(Form.ProgramsSelected.value, "[" & maProgramIDs(intI) & "]") > 0 Then
                document.all("chkProgram" & maProgramIDs(intI)).checked = True
            End If
        Next
    End If
    Call SecurityRoleOptions
End Sub

Sub LoadDictionaries()
    <%
    For Each moDictObj In mdctAliasIDs
        Select Case Parse(mdctAliasIDs(moDictObj),"^",1)
            Case 125
                Response.Write vbTab & "If Not mdctManagers.Exists(""" & Parse(mdctAliasIDs(moDictObj), "^", 2) & """) Then mdctManagers.Add """ & Parse(mdctAliasIDs(moDictObj), "^", 2) & """, """ & Parse(mdctAliasIDs(moDictObj), "^", 2) & "^" & Parse(mdctAliasIDs(moDictObj), "^", 3) & """" & vbCrLf
            Case 126
                Response.Write vbTab & "If Not mdctDivisions.Exists(""" & Parse(mdctAliasIDs(moDictObj), "^", 2) & """) Then mdctDivisions.Add """ & Parse(mdctAliasIDs(moDictObj), "^", 2) & """, """ & Parse(mdctAliasIDs(moDictObj), "^", 2) & """" & vbCrLf
            Case 250
                Response.Write vbTab & "If Not mdctRegions.Exists(""" & Parse(mdctAliasIDs(moDictObj), "^", 2) & """) Then mdctRegions.Add """ & Parse(mdctAliasIDs(moDictObj), "^", 2) & """, """ & Parse(mdctAliasIDs(moDictObj), "^", 2) & """" & vbCrLf
        End Select
        Response.Write vbTab & "If Not mdctAliasIDs.Exists(""" & moDictObj & """) Then mdctAliasIDs.Add """ & moDictObj & """, """ & Parse(mdctAliasIDs(moDictObj), "^", 2) & """" & vbCrLf
    Next
    For Each moDictObj In mdctElementIDs
        If moDictObj = "1^Arrearage" Then
            Response.Write vbTab & "mintArrearageID=" & mdctElementIDs(moDictObj) & vbCrLf
        End If
        Response.Write vbTab & "mdctElementIDs.Add """ & moDictObj & """, """ & mdctElementIDs(moDictObj) & """" & vbCrLf
    Next
    For Each moDictObj In mdctElements
        Response.Write vbTab & "mdctElements.Add CLng(" & moDictObj & "), """ & mdctElements(moDictObj) & """" & vbCrLf
    Next
    For Each moDictObj In mdctFactors
        Response.Write vbTab & "mdctFactors.Add CLng(" & moDictObj & "), """ & mdctFactors(moDictObj) & """" & vbCrLf
    Next
    %>
End Sub

Sub LoadWindowTable()
    <%' This dictionary object will be used when closing child windows from Main.
    ' The key will be the name of the page that is open in the window.  The item
    ' contains 3 elements: 
    '   1)Y/N indicating if page requires the variable mblnCloseClicked
    '     to be set to True to allow window to be closed without user input.
    '   2)If page can be in EDIT mode, the name of the Save button.  Leave blank if 
    '     page does not have an EDIT mode.
    '   3)Name of page to display in message box if page is found to be in EDIT mode.%>
    mdctWindowTable.Add UCase("CaseAddEdit"),"Y^cmdSaveRecord^Enter Case Review"
    mdctWindowTable.Add UCase("FindCase"),"N^^Find Case Review For Edit"
    mdctWindowTable.Add UCase("Reports"),"Y^^View Reports"
    mdctWindowTable.Add UCase("EmployeeSelect"),"N^cmdSave^Employees"
    mdctWindowTable.Add UCase("UsersSelect"),"N^cmdSave^User Logins"
    mdctWindowTable.Add UCase("FactorAddEditAssign"),"N^cmdSave^Causal Factors"
    mdctWindowTable.Add UCase("ListSelect"),"N^cmdSave^Drowdown Lists"
    mdctWindowTable.Add UCase("PosEmpSelect"),"Y^cmdOkToClose^Position/ Employees"
    mdctWindowTable.Add UCase("GuideQuestionSelect"),"N^cmdSave^Review Guide"
    mdctWindowTable.Add UCase("ReviewTypeSelect"),"N^cmdSave^Review Types"
    mdctWindowTable.Add UCase("AppOptionSelect"),"N^cmdSave^Application Options"
    mdctWindowTable.Add UCase("SQLQueries"),"Y^cmdSave^SQL Queries"
    mdctWindowTable.Add UCase("MngLvlAddEdit"),"N^cmdSave^Management Levels"
    mdctWindowTable.Add UCase("ReReviewAddEdit"),"Y^cmdSaveRecord^Re-Review"
    mdctWindowTable.Add UCase("FindReReview"),"Y^^Find Re-Review"
    mdctWindowTable.Add UCase("CAReReviewAddEdit"),"Y^cmdSaveRecord^CAR Re-Review"
    mdctWindowTable.Add UCase("CAFindRReReview"),"Y^^Find CAR Re-Review"
    mdctWindowTable.Add UCase("ReportEdit"),"N^cmdSave^Reports Maintenance"  
    mdctWindowTable.Add UCase("Admin"),"Y^cmdSave^System Administration"
End Sub

Sub LoadReviewList(strSortOrder)
    Dim objWindow
    
    strCaller = ""
    If strSortOrder = "CASEADDEDIT" Or strSortOrder = "REREVIEWADDEDIT" Or strSortOrder = "CARREREVIEWADDEDIT" Then
        strCaller = strSortOrder
        strSortOrder = "" 
    End If
    If strSortOrder <> "" Then
        If strSortOrder = Parse(mstrLastSortOrder,"^",1) Then
            If Parse(mstrLastSortOrder,"^",2) = "ASC" Then
                'Previous sort was ASC, so make this sort DESC
                strSortOrder = strSortOrder & "^DESC"
            Else
                strSortOrder = strSortOrder & "^ASC"
            End If
        Else
            strSortOrder = strSortOrder & "^ASC"
        End If
        mstrLastSortOrder = strSortOrder
    Else
        strSortOrder = mstrLastSortOrder
    End If
    strSortOrder = Replace(strSortOrder,"^"," ")
    fraToDoList.frameElement.src = "MainReviewList.asp?SetFocus=" & strCaller & "&Load=Y&UserType=<%=mstrUserType%>&UserID=<%=gstrUserID%>&SortOrder=" & strSortOrder
End Sub

Sub ShowReviewList(lngRecordCount)
    If lngRecordCount > 0 Then
        divToDoListHdr.style.left = 33
        divToDoList.style.left = 35
    Else
        divToDoListHdr.style.left = -3000
        divToDoList.style.left = -3000
    End If
End Sub

Sub ColClick(intRowID)
    Select Case intRowID
        Case 1
            Call LoadReviewList("rvwID")
        Case 2
            Call LoadReviewList("rvwCaseNumber")
        Case 3
            Call LoadReviewList("CaseName")
        Case 4
            Call LoadReviewList("rvwDateEntered")
        Case 5
            Call LoadReviewList("rvwResponseDueDate")
        Case 6
            Call LoadReviewList("ReviewClass")
        Case Else
            Call LoadReviewList("")
    End Select
End Sub

Sub ShowPage(blnShow)
    If blnShow Then
        Header.style.visibility="visible"
        PageFrame.style.visibility="visible"
        lblDatabaseStatus.style.visibility="hidden"
        PageBody.style.cursor = "default"
    Else
        Header.style.visibility="hidden"
        PageFrame.style.visibility="hidden"
        lblDatabaseStatus.style.visibility="visible"
        PageBody.style.cursor = "wait"
    End If
End Sub

Sub ButtonClick(cmdButton)
    Dim strPage
    Dim intWindow
    Dim intResponse
    Dim strMessage
    
    intWindow = -1
    Form.CalledFrom.Value = "Main"
    Form.WhoCalled.value = ""
    Select Case cmdButton.id
        Case "cmdAddEdit"
            strPage = "CaseAddEdit.asp"
            intWindow = 1
            Form.rvwID.value = 0
            Form.LastRvwID.value = 0
            Form.CalledFrom.Value = "Main"

        Case "cmdViewReports"
            strPage = "Reports.Asp"
            Form.CalledFrom.Value = "Main"
            intWindow = 3

        Case "cmdFindCase"
            strPage = "FindCase.asp"
            intWindow = 2

        Case "cmdUsers"
            strPage = "UsersSelect.asp"
            intWindow = 6

        Case "cmdFactors"
            strPage = "ElementsCausalFactors.asp"
            intWindow = 6
            Form.WhoCalled.value = "Factors"
            
        Case "cmdElements"
            strPage = "ElementsCausalFactors.asp"
            intWindow = 6
            Form.WhoCalled.value = "Elements"
        
        Case "cmdLists"
            strPage = "ListSelect.asp"
            intWindow = 6

        Case "cmdManagers"
            strPage = "AliasSelect.asp"
            Form.CalledFrom.Value = ""
            intWindow = 6

        Case "cmdReviewTypes"
            strPage = "ReviewTypeSelect.asp"
            intWindow = 6

        Case "cmdAdminMenu"
            strPage = "Admin.asp"
            intWindow = 6
        
        Case "cmdLogOff"
            Form.CalledFrom.Value = "ManualLogon"
            strPage = "Logon.asp"
            
        Case "cmdReReviewAddEdit"
            strPage = "ReReviewAddEdit.asp"
            Form.ReReviewID.value = 0
            Form.ReReviewTypeID.value = 0
            Form.CalledFrom.Value = "Main"
            intWindow = 4

        Case "cmdFindReReview"
			strPage = "FindReReview.asp"
            Form.ReReviewTypeID.value = 0
            intWindow = 5
		
		Case "cmdCARReReviewAddEdit"
            strPage = "ReReviewAddEdit.asp"
            Form.ReReviewTypeID.value = 1
            Form.ReReviewID.value = 0
            Form.CalledFrom.Value = "Main"
            intWindow = 7
		
        Case "cmdFindCARReReview"
			strPage = "FindReReview.asp"
            Form.ReReviewTypeID.value = 1
            intWindow = 8

		Case "cmdReportEdit"
			strPage = "ReportEdit.asp"
            intWindow = 6

        Case "cmdArchive"           
            strPage = "ArchiveMenu.asp"
            intWindow = 6
    End Select

    Form.Action = strPage
    If intWindow >= 0 Then
        Call ManageWindows(intWindow, "Open")
    Else
        strMessage = CheckChildWindowsForClosing(False)
        If strMessage <> "" Then
            strMessage = "If you proceed with Log Off, changes made on the" & Space(10) & vbCrLf & _
                         "follwing page(s) will be lost:" & vbCrLf & vbCrLf & _
                         strMessage & vbCrLf & vbCrLf & _
                         "Do you wish to proceed?"
                         
            intResponse = MsgBox(strMessage,vbInformation + vbYesNo,"Case Review Log Off")
            
            If intResponse = vbNo Then Exit Sub
            
            strMessage = CheckChildWindowsForClosing(True)
        End If
        mblnCloseClicked = True
        <% If (gblnUserAdmin Or gblnUserQA Or Instr(gstrUserID, "ADMIN") > 0) Or gblnUseLogon = True Then %>
            Form.submit
        <% Else %>
            window.close
        <% End If %>
    End If
End Sub

Sub window_onbeforeunload()
    If mblnCloseClicked = False Then
        window.event.returnValue = "Closing this page will exit the application." & Space(10)
    End If
End Sub

Function CheckChildWindowsForClosing(blnConfirmed)
    Dim intI
    Dim intJ
    Dim aItems
    Dim aReportItems
    Dim strRecord
    Dim strPages
    Dim oItem
    
    strPages = ""
    <%' blnConfirmed is passed from the Log Off button click. A TRUE value
    ' indicates the user was displayed a message stating 1 or more windows
    ' are in EDIT mode and changes will be lost.  User confirmed the close. %>
    If Not mdctWindows Is Nothing Then
        aItems = mdctWindows.Items
        If blnConfirmed = False Then
            <%' Check each window to determine if it is in 'EDIT' mode and requires 
            'user confirmation.  If any window requires user confirmation, do not close any windows.%>
            For intI = 0 To UBound(aItems)
                If Not aItems(intI) Is Nothing Then
                    If Not aItems(intI).closed Then
                        strRecord = mdctWindowTable(aItems(intI).Name)
                        If Parse(strRecord,"^",2) <> "" Then
                            Set oItem = aItems(intI).document.all(Parse(strRecord,"^",2))
                            If Not oItem Is Nothing Then
                                If oItem.disabled = False Then
                                    strPages = strPages & Parse(strRecord,"^",3) & ", "
                                End If
                                Set oItem = Nothing
                            End If
                        End If
                    End If
                End If
            Next
            
            If Len(strPages) > 0 Then
                strPages = Left(strPages,Len(strPages)-2)
            End If
        End If
                
        If strPages = "" Then
            <%' Check each window if it requires a variable to be set to allow for a 
            ' close without user intervention.  Set any windows as needed and close.%>
            For intI = 0 To UBound(aItems)
                If Not aItems(intI) Is Nothing Then
                    If Not aItems(intI).closed Then
                        strRecord = mdctWindowTable(aItems(intI).Name)
                        If Parse(strRecord,"^",1) = "Y" Then
                            aItems(intI).mblnCloseClicked = True
                        End If
                        aItems(intI).close
                    End If
                End If
            Next
            ' Close any report windows
            If mdctReportWindows.Count > 0 Then
                aReportItems = mdctReportWindows.Items
                For intI = 0 To UBound(aReportItems)
                    If Not aReportItems(intI) Is Nothing Then
                        If Not aReportItems(intI).closed Then
                            aReportItems(intI).close
                        End If
                    End If
                Next
            End If
        End If
    End If
    CheckChildWindowsForClosing = strPages
End Function

Sub ManageWindows(intWindow, strAction)
    Dim intI
    Dim objWindow
    Dim objWindow2
    Dim dtmLastDate, dtmCurDate
    Dim intResponse
    Dim strWindowName
    Dim strRecord
    Dim oItem
    Dim strMessage
    Dim intReReviewWinID

    strWindowName = UCase(Form.action)
    strWindowName = Trim(Replace(strWindowName,".ASP",""))
    If strWindowName = "REREVIEWADDEDIT" Or strWindowName = "FINDREREVIEW" Then
        If Form.ReReviewTypeID.value = 1 Then
            strWindowName = "CA" & strWindowName
        End If
    End If
    Set objWindow = GetWindow(intWindow)

    Select Case strAction
        Case "Open"
            If objWindow Is Nothing Then
                Set objWindow = window.open("WindowLaunch.asp", strWindowName, mstrFeatures, False)
                mdctWindows.Add intWindow, objWindow
                If intWindow = 6 Then mstrLastPage = Form.action
            Else
                If intWindow <> 6 Then
                    ' If call is for any of the specific windows, just set focus
                    objWindow.focus
                Else
                    <%' If call is for the 6th, catch all window, check if last window opened
                    ' in 6 is same as current request.  If it is, simply set focus.  If it is
                    ' a different window, close current window and open a new one%>
                    If mstrLastPage = Form.action Then
                        objWindow.focus
                    Else
                        ' Check currently opened window for EDIT mode
                        strRecord = mdctWindowTable(objWindow.Name)
                        If Parse(strRecord,"^",2) <> "" Then
                            Set oItem = objWindow.document.all(Parse(strRecord,"^",2))
                            If Not oItem Is Nothing Then
                                If oItem.disabled = False Then
                                    strMessage = "If you proceed, changes made on the" & Space(10) & vbCrLf & _
                                                Parse(strRecord,"^",3) & " page will be lost." & vbCrLf & vbCrLf & _
                                                "Do you wish to proceed?"
                                                 
                                    intResponse = MsgBox(strMessage,vbInformation + vbYesNo,"Case Review Main Menu")
                                    
                                    If intResponse = vbNo Then
                                        objWindow.focus
                                        Exit Sub
                                    End If
                                End If
                                Set oItem = Nothing
                            End If
                        End If
                        ' If currently opened window calls window_onbeforeunload, set mblnCloseClicked to True
                        If Parse(strRecord,"^",1) = "Y" Then objWindow.mblnCloseClicked = True
                        
                        ' Use timer function to ensure window is closed before attempt to open new window begins
                        mintWindow = 6
                        mlngTimerID = window.setInterval("ReloadWindow",500)
                        objWindow.close
                        mdctWindows.Remove(6)
                    End If
                End If
            End If    
        Case "Close"
            If Not objWindow Is Nothing Then
                objWindow.close
                mdctWindows.Remove(CInt(intWindow))
            End If
        Case "EditReview" ' Called From FindCase.asp
            Form.CalledFrom.Value = "Find"
            ' Check if CaseAddEdit window is open
            Set objWindow2 = GetWindow(1)
            If objWindow2 Is Nothing Then
                ' Not open so open a new window
                Set objWindow2 = window.open("WindowLaunch.asp",  "CASEADDEDIT", mstrFeatures, False)
                mdctWindows.Add 1, objWindow2
            Else
                ' Check status of current CaseAddEdit
                If objWindow2.cmdSaveRecord.disabled = False Then
                    objWindow2.focus
                    intResponse = MsgBox("If you continue, any changes made to the currently" & vbCrLf & "loaded review will be lost." & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, "Confirm Load")
                    If intResponse = vbNo Then
                        objWindow2.focus
                        Exit Sub
                    End If
                End If
                ' Use timer function to ensure window is closed before attempt to open new window begins
                mintWindow = 1
                mlngTimerID = window.setInterval("ReloadWindow",500)
                objWindow2.mblnCloseClicked = True
                objWindow2.Close
                mdctWindows.Remove(1)
            End If
        Case "EditReReview" ' Called From FindReReview.asp
            Form.CalledFrom.Value = "Find"
            If Form.ReReviewTypeID.Value = 0 Then
                intReReviewWinID = 4
            Else
                intReReviewWinID = 7
            End If
            ' Check if CaseAddEdit window is open
            Set objWindow2 = GetWindow(intReReviewWinID)
            If objWindow2 Is Nothing Then
                ' Not open so open a new window
                If Form.ReReviewTypeID.Value = 0 Then
                    Set objWindow2 = window.open("WindowLaunch.asp",  "REREVIEWADDEDIT", mstrFeatures, False)
                Else
                    Set objWindow2 = window.open("WindowLaunch.asp",  "CARREREVIEWADDEDIT", mstrFeatures, False)
                End If
                mdctWindows.Add intReReviewWinID, objWindow2
            Else
                ' Check status of current CaseAddEdit
                If objWindow2.cmdSaveRecord.disabled = False Then
                    objWindow2.focus
                    intResponse = MsgBox("If you continue, any changes made to the currently" & vbCrLf & "loaded re-review will be lost." & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, "Confirm Load")
                    If intResponse = vbNo Then
                        objWindow2.focus
                        Exit Sub
                    End If
                End If
                ' Use timer function to ensure window is closed before attempt to open new window begins
                mintWindow = intReReviewWinID
                mlngTimerID = window.setInterval("ReloadWindow",500)
                objWindow2.mblnCloseClicked = True
                objWindow2.Close
                mdctWindows.Remove(intReReviewWinID)
            End If
    End Select
End Sub

Function ReloadWindow()
    Dim objWindow
    Dim strWindowName
 
    strWindowName = UCase(Trim(Replace(Form.action,".asp","")))
    
    window.clearInterval mlngTimerID
    Set objWindow = window.open("WindowLaunch.asp", strWindowName, mstrFeatures, False)
    mdctWindows.Add mintWindow, objWindow
    If mintWindow = 6 Then mstrLastPage = Form.action
End Function

Function GetWindow(intWindow)
    Dim intI
    Dim blnFound
    Dim aKeys
    Dim aItems
    Dim objWindow

    aKeys = mdctWindows.Keys
    aItems = mdctWindows.Items

    Set objWindow = Nothing
    blnFound = False
    For intI = 0 To UBound(aKeys)
        If CInt(aKeys(intI)) = CInt(intWindow) Then
            Set objWindow = aItems(intI)

            ' If a window was closed without cycling through Main, remove it from collection
            If objWindow.closed = True Then
                mdctWindows.Remove(aKeys(intI))
                Set objWindow = Nothing
            Else
                blnFound = True
            End If
            Exit For
        End If
    Next
    
    Set GetWindow = objWindow
End Function

Sub AddReportWindow(lngWindowID, objWindow)
    mdctReportWindows.Add lngWindowID, objWindow
End Sub

Sub ButtonMouseOver(cmdButton)
    cmdButton.style.fontWeight = "bold"
End Sub

Sub ButtonMouseOut(cmdButton)
    cmdButton.style.fontWeight = "normal"
End Sub
'-----------
Sub onclick_ProgramLabel(lblCtl)
    Dim intI
    
    For intI = 0 To UBound(maProgramIDs)
        If lblCtl.ID = document.all("lblProgram" & maProgramIDs(intI)).ID Then
            document.all("chkProgram" & maProgramIDs(intI)).checked = Not document.all("chkProgram" & maProgramIDs(intI)).checked
            Call onclick_Program(document.all("chkProgram" & maProgramIDs(intI)))
            Exit For
        End If
    Next
End Sub

Sub CheckOpenedWindows(intFrom)
    Dim objWindow
    Dim strCheckBox
    
    Set objWindow = GetWindow(2)
    If Not objWindow Is Nothing Then
        For intI = 0 To UBound(maProgramIDs)
            strCheckBox = "chkProgram" & maProgramIDs(intI)
            If document.all("chkProgram" & maProgramIDs(intI)).Checked = True And objWindow.document.all(strCheckBox).checked = False Then
                ' Check box on child page
                objWindow.document.all(strCheckBox).checked = True
            ElseIf document.all("chkProgram" & maProgramIDs(intI)).Checked = False And objWindow.document.all(strCheckBox).checked = True Then
                ' Un Check box on child page
                objWindow.document.all(strCheckBox).checked = False
            End If
        Next
    End If
End Sub

Sub onclick_Program(chkCtl)
    Call CheckOpenedWindows(0)
End Sub

Sub SecurityRoleOptions
	Dim intI
	Dim strOptions
	
	intI = 2
	strOptions = Parse("<%=gstrOptions%>", "[", intI)
	Do While strOptions <> "" 
		strOptions = Parse("<%=gstrOptions%>", "[", intI)
		Call DisplayButtons(Parse(strOptions, "]", 1))
		intI = intI + 1
	Loop
End Sub
Sub DisplayButtons(intOption)
	Select Case(intOption)
		Case 1
		    If "<%=gstrRoles%>" = "[6]" Then
    			cmdFindCase.style.left = 35
		    Else
			    cmdAddEdit.style.display = "inline" '1:Enter Case Reviews
			End If
			cmdFindCase.style.display = "inline" '1:Find an existing case review to edit.
		Case 2
			cmdViewReports.style.display = "inline" '2:View Case Review Reports
		Case 3
			cmdReReviewAddEdit.style.display = "inline" '3:Add/update Evaluations
			cmdFindReReview.style.display = "inline" '3:Find Evaluations
			cmdCARReReviewAddEdit.style.display = "inline" '3:Add/update Evaluations
			cmdFindCARReReview.style.display = "inline" '3:Find Evaluations
		Case 5
			cmdUsers.style.display = "inline" '5:Add/update the User table
		Case 6
			cmdFactors.style.display = "inline" '6:Update Causal Factors List
			cmdElements.style.display = "inline" '6:Update Elments List
		Case 7
			cmdLists.style.display = "inline" '7:Update Dropdown Lists
		Case 8
			'cmdManagers.style.display = "inline" '8:Add/Update Positions
		Case 11
			cmdReviewTypes.style.display = "inline" '11:Define Review Types
		Case 12
			cmdReportEdit.style.display = "inline" '12:Modify Reports
		Case 13
		    cmdAdminMenu.style.display = "inline" '13:System admin sub-menu
        Case 17
		    cmdArchive.style.display = "inline" '13:System admin sub-menu
		    If InStr("<%=gstrOptions%>","[3]") > 0 And InStr("<%=gstrOptions%>","[4]") > 0 Then
		        ' If user has Re-Reviewer and Employee roles, leave Archive button where it is
		    ElseIf InStr("<%=gstrOptions%>","[4]") = 0 Then
		        ' If user does not have Employee role, move Archive button where employee button is
		        cmdArchive.style.top = 100
		        cmdArchive.style.left = 545
		    ElseIf InStr("<%=gstrOptions%>","[3]") = 0 Then
		        ' If user has neither Re-Reviewer role or Employee role, move Archive button where re-review button is
		        cmdArchive.style.top = 135
		        cmdArchive.style.left = 35
		    End If
        
	End Select
	Select Case intOption
	    Case 5,6,7,8,9,10,11,12,13
	        Select Case "<% = gstrOptions %>"
	            Case "[8]", "[4][8]"
	                ' If only options are Employee and/or Position Employee, button(s) are
	                ' moved to top part of screen, do NOT display the System Admin label or
	                ' the Program List
	                lblPrograms.style.display = "none"
	            Case Else
	                If InStr("<% = gstrOptions %>","[9]") > 0 Then
        	            lblAdmin.style.display = "inline"
        	        End If
	        End Select
	End Select
End Sub
</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin="5" leftMargin="5" topMargin="5" rightMargin="5">
    
    <DIV id=Header class=DefTitleArea style="WIDTH: 737">
        <SPAN id=lblUserName class=DefLabel
            style="LEFT:10; TOP:5; WIDTH:200; COLOR:<%=gstrBorderColor%>; visibility:visible">
            <%=gstrUserName%>
        </SPAN>
        
        <SPAN id=lblAppTitle class=DefTitleText
            style="WIDTH: 737">
            <%=mstrPageTitle%>
        </SPAN>
        
        <SPAN id=lblCurrentDate class=DefLabel
            style="LEFT:525; TOP:5; WIDTH:200; TEXT-ALIGN:right; COLOR:<%=gstrBorderColor%>">
            <%=FormatDateTime(Date, vbLongdate)%>
        </SPAN>
    </DIV>
            
    <DIV id=PageFrame class=DefPageFrame style="HEIGHT: 425; WIDTH: 737; TOP: 51; visibility:hidden">

        <SPAN id=lblDatabaseStatus
            class=DefLabel
            STYLE="VISIBILITY:hidden; LEFT:20; WIDTH:200; TOP:20; TEXT-ALIGN:center">
            Accessing Database...
        </SPAN>

        <SPAN class=DefLabel 
            id=lblCaseReviewMenu
            style="FONT-SIZE:10pt; FONT-WEIGHT:bold; TOP:15; LEFT:25; WIDTH:200">
            Case Review Menu:
        </SPAN>

        <DIV id=lblPrograms
            style="TOP:50; LEFT:-1130; WIDTH:675; HEIGHT:40; BORDER-STYLE:solid; BORDER-WIDTH:1; BORDER-COLOR:silver">

            <SPAN class=DefLabel 
                id=lblSelectPrograms
                style="FONT-SIZE:10pt; TOP:-10; LEFT:10; WIDTH:260; PADDING-LEFT:5; BACKGROUND-COLOR:<%=gstrBackColor%>">
                Select the Program or Programs for Review
            </SPAN>
            
            <INPUT id=chkProgram1 type=checkbox title="Food Stamps"
                style="LEFT:25; WIDTH:20; TOP:10; HEIGHT: 20"
                onclick="onclick_Program(chkProgram1)" 
                tabIndex=16 NAME="chkProgram1">
            <SPAN id=lblProgram1 class=DefLabel style="TOP:12; LEFT:50; WIDTH:90"
                onclick="onclick_ProgramLabel(lblProgram1)">
                Food Stamps
            </SPAN>

            <INPUT id=chkProgram2 type=checkbox title="TA"
                style="LEFT:140; WIDTH:20; TOP:10; HEIGHT: 20"
                onclick="onclick_Program(chkProgram2)" 
                tabIndex=16 NAME="chkProgram2">
            <SPAN id=lblProgram2 class=DefLabel style="TOP:12; LEFT:165; WIDTH:90"
                onclick="onclick_ProgramLabel(lblProgram2)">
                TA
            </SPAN>

            <INPUT id=chkProgram3 type=checkbox title="Medical"
                onclick="onclick_Program(chkProgram3)" 
                style="LEFT:260; WIDTH:20; TOP:10; HEIGHT: 20"
                tabIndex=16 NAME="chkProgram3">
            <SPAN id=lblProgram3 class=DefLabel style="TOP:12; LEFT:285; WIDTH:90"
                onclick="onclick_ProgramLabel(lblProgram3)">
                Medical
            </SPAN>
        </DIV>
		<BUTTON id=cmdAddEdit title="Enter Case Reviews" class=DefBUTTON
            onclick="ButtonClick(cmdAddEdit)" onmouseover="ButtonMouseOver(cmdAddEdit)" onmouseout="ButtonMouseOut(cmdAddEdit)"
            style="LEFT:35; TOP:50;display:none"
            accessKey=C tabIndex=1>
            Enter <U>C</U>ase Reviews
        </BUTTON>
        <BUTTON id=cmdFindCase class=DefBUTTON title="Find an existing case review to edit." 
            onclick="ButtonClick(cmdFindCase)" onmouseover="ButtonMouseOver(cmdFindCase)" onmouseout="ButtonMouseOut(cmdFindCase)"
            style="LEFT:205; TOP:50;display:none"
            accessKey=I tabIndex=1>
            F<U>i</U>nd Case Review
        </BUTTON>
        <BUTTON id=cmdViewReports title="View Case Review Reports" class=DefBUTTON
            onclick="ButtonClick(cmdViewReports)" onmouseover="ButtonMouseOver(cmdViewReports)" onmouseout="ButtonMouseOut(cmdViewReports)"
            style="LEFT: 375; TOP: 50;display:none"
            accessKey=R tabIndex=1>
            <U>V</U>iew Reports
        </BUTTON>
        <BUTTON id=cmdArchive class=DefBUTTON title="Archive old information." 
            onclick="ButtonClick(cmdArchive)" onmouseover="ButtonMouseOver(cmdArchive)" onmouseout="ButtonMouseOut(cmdArchive)"
            style="LEFT:545; TOP:50;display:none"
            accessKey=A tabIndex=1>
            <U>A</U>rchive
        </BUTTON>
        <BUTTON id=cmdReReviewAddEdit class=DefBUTTON title="Add/update <%=gstrEvaluation%>s" 
            onclick="ButtonClick(cmdReReviewAddEdit)" onmouseover="ButtonMouseOver(cmdReReviewAddEdit)" onmouseout="ButtonMouseOut(cmdReReviewAddEdit)"
            style="LEFT: 35; TOP:85;display:none"
            accessKey=E tabIndex=1>
            Enter <%="<u>" & LEFT(gstrEvaluation, 1) & "</u>" & Mid(gstrEvaluation, 2)%>s
        </BUTTON>
        <BUTTON id=cmdFindReReview class=DefBUTTON title="Find <%=gstrEvaluation%>s" 
            onclick="ButtonClick(cmdFindReReview)" onmouseover="ButtonMouseOver(cmdFindReReview)" onmouseout="ButtonMouseOut(cmdFindReReview)"
            style="LEFT: 205; TOP:85;display:none"
            accessKey=E tabIndex=1>
            <%="<u>F</u>ind&nbsp" & gstrEvaluation%>
        </BUTTON>
        <BUTTON id="cmdCARReReviewAddEdit" class=DefBUTTON title="Add a Corrective Action Re-Review" 
            onclick="ButtonClick(cmdCARReReviewAddEdit)" onmouseover="ButtonMouseOver(cmdCARReReviewAddEdit)" onmouseout="ButtonMouseOut(cmdCARReReviewAddEdit)"
            style="LEFT: -1375; TOP: 85;display:none"
             tabIndex=1>
            Enter CAR <%="<u>" & LEFT(gstrEvaluation, 1) & "</u>" & Mid(gstrEvaluation, 2)%>s
        </BUTTON>
        <BUTTON id="cmdFindCARReReview" class=DefBUTTON title="Find Corrective Action Re-Review" 
            onclick="ButtonClick(cmdFindCARReReview)" onmouseover="ButtonMouseOver(cmdFindCARReReview)" onmouseout="ButtonMouseOut(cmdFindCARReReview)"
            style="LEFT: -1545; TOP: 85;display:none"
             tabIndex=1>
            <%="<u>F</u>ind&nbsp;CAR&nbsp;" & gstrEvaluation%>
        </BUTTON>

<%
        Dim strStyle
        strStyle="cursor:hand"
%>
        <DIV id=divToDoListHdr class=DefPageFrame style="positon:absolute;LEFT:35;WIDTH:670;border-style:none; HEIGHT:30; TOP:145">
            <SPAN class=DefLabel style="TOP:1;LEFT:0;WIDTH:668;text-align:center">
                <B>Reviews/Re-Reviews That Require Your Attention</B>
            </SPAN>
            <table style="position:absolute;top:15" width=652>
                <tr title="Click To Sort Reviews">
                    <td class=CellLabel onclick=ColClick(1) style="width:90;<%=strStyle%>">ID</td>
                    <td class=CellLabel onclick=ColClick(2) style="width:80;<%=strStyle%>">Case Number</td>
                    <td class=CellLabel onclick=ColClick(3) style="width:135;<%=strStyle%>">Case Name</td>
                    <td class=CellLabel onclick=ColClick(4) style="width:80;<%=strStyle%>">Date Entered</td>
                    <td class=CellLabel onclick=ColClick(5) style="width:100;<%=strStyle%>">Response Due</td>
                    <td class=CellLabel onclick=ColClick(6) style="width:140;<%=strStyle%>">Action</td>
                </tr>
            </table>
        </DIV>
        <DIV id=divToDoList class=DefPageFrame style="positon:absolute;LEFT:35;WIDTH:670;border-style:solid; HEIGHT:190; TOP:180">
            <IFRAME ID=fraToDoList src="Blank.html?Load=N"
                STYLE="positon:absolute;LEFT:0;WIDTH:668;HEIGHT:190;TOP:0;BORDER-style:none" FRAMEBORDER=0>
            </IFRAME>
        </DIV>

        <SPAN class=DefLabel 
            id=lblAdmin
            style="FONT-SIZE:10pt; FONT-WEIGHT:bold; TOP:130; LEFT:25; WIDTH:200;display:none">
            System Administration Menu:
        </SPAN>
        <BUTTON id=cmdUsers class=DefBUTTON title="Add/update the User table" 
            onclick="ButtonClick(cmdUsers)" onmouseover="ButtonMouseOver(cmdUsers)" onmouseout="ButtonMouseOut(cmdUsers)"
            style="LEFT:35; TOP:155;display:none"
            accessKey=U tabIndex=1><U>
            U</U>sers Logins
        </BUTTON>
        <BUTTON id=cmdElements class=DefBUTTON title="Update Actions / Screens / Questions" 
            onclick="ButtonClick(cmdElements)" onmouseover="ButtonMouseOver(cmdElements)" onmouseout="ButtonMouseOut(cmdElements)"
            style="LEFT:205; TOP:155;display:none"
            accessKey=F tabIndex=1>
            Elements
        </BUTTON>
        <BUTTON id=cmdFactors class=DefBUTTON title="Update Decisions and Field Names" 
            onclick="ButtonClick(cmdFactors)" onmouseover="ButtonMouseOver(cmdFactors)" onmouseout="ButtonMouseOut(cmdFactors)"
            style="LEFT:375; TOP:155;display:none"
            accessKey=L tabIndex=1>
            Causal Factors
        </BUTTON>
        <BUTTON id=cmdLists class=DefBUTTON title="Update Dropdown Lists" 
            onclick="ButtonClick(cmdLists)" onmouseover="ButtonMouseOver(cmdLists)" onmouseout="ButtonMouseOut(cmdLists)"
            style="LEFT:545; TOP:155;display:none"
            accessKey=L tabIndex=1>
            Dropdown <U>L</U>ists
        </BUTTON>
        <BUTTON id=cmdManagers class=DefBUTTON title="Add/Update Positions / Employees" 
            onclick="ButtonClick(cmdManagers)" onmouseover="ButtonMouseOver(cmdManagers)" onmouseout="ButtonMouseOut(cmdManagers)"
            style="LEFT:35; TOP:190;display:none"
            accessKey=P tabIndex=1>
            Upper <U>M</U>anagement
        </BUTTON>
        <BUTTON id=cmdReviewTypes class=DefBUTTON title="Define Review Types" 
            onclick="ButtonClick(cmdReviewTypes)" onmouseover="ButtonMouseOver(cmdReviewTypes)" onmouseout="ButtonMouseOut(cmdReviewTypes)"
            style="LEFT:205; TOP:190;display:none"
            accessKey=T tabIndex=1>
            Review <U>T</U>ypes
        </BUTTON>
		<BUTTON id="cmdReportEdit" class=DefBUTTON title="Modify Reports" 
            onclick="ButtonClick(cmdReportEdit)" onmouseover="ButtonMouseOver(cmdReportEdit)" onmouseout="ButtonMouseOut(cmdReportEdit)"
            style="LEFT:375; TOP:190;display:none"
            accessKey=X tabIndex=1>
            <U>R</U>eports Maintenance
        </BUTTON>
        <BUTTON id=cmdAdminMenu class=DefBUTTON title="System Admin Sub Menu" 
            onclick="ButtonClick(cmdAdminMenu)" onmouseover="ButtonMouseOver(cmdAdminMenu)" onmouseout="ButtonMouseOut(cmdAdminMenu)"
            style="LEFT:35; TOP:380;display:none"
            accessKey=O tabIndex=1>
            System <U>A</U>dmin
        </BUTTON>

        <BUTTON id=cmdAppSettings class=DefBUTTON title="Configure Settings and Options" 
            onclick="ButtonClick(cmdAppSettings)" onmouseover="ButtonMouseOver(cmdAppSettings)" onmouseout="ButtonMouseOut(cmdAppSettings)"
            style="LEFT:-1135; TOP:380;display:none"
            accessKey=O tabIndex=1>
            Application <U>O</U>ptions
        </BUTTON>
        <BUTTON id=cmdSQLQueries class=DefBUTTON title="Execute SQL Queries" 
            onclick="ButtonClick(cmdSQLQueries)" onmouseover="ButtonMouseOver(cmdSQLQueries)" onmouseout="ButtonMouseOut(cmdSQLQueries)"
            style="LEFT:-1205; TOP:380;display:none"
            accessKey=Q tabIndex=1>
            S<U>Q</U>L Queries
        </BUTTON>
        <BUTTON id=cmdLogOff class=DefBUTTON title="Log Off" 
            onclick="ButtonClick(cmdLogOff)" 
            style="LEFT:555; FONT-WEIGHT:bold; TOP:380"
            tabIndex=1>
            Log Off
        </BUTTON>
    </DIV>
</BODY>

<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="" ID=Form>
        <%Call CommonFormFields()%>
<%
' If Main is called from Logon.asp, resize the page, otherwise do not.
' ResizeScreen added to hidden form and will be used in window_onload
strResizeScreen = "N"
If Request.Form("CalledFrom") = "Logon" Then strResizeScreen = "Y"
Call WriteFormField("ResizeScreen", strResizeScreen)
Call WriteFormField("casID", "0")
Call WriteFormField("rvwID", 0)
Call WriteFormField("LastRvwID", 0)
Call WriteFormField("FormAction", "")
Call WriteFormField("ReportWindowID",0)
Call WriteFormField("ReReviewID",0)
Call WriteFormField("ReReviewTypeID",0)
Call WriteFormField("WhoCalled","")

%>
</FORM>
</HTML>
<%
gadoCon.Close
Set gadoCon = Nothing
%>
<!--#include file="IncCmnFormFields.asp"-->
