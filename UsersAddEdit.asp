<%@ LANGUAGE="VBScript" %><%
Option Explicit
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: StaffAddEdit.asp                                                '
'  Purpose: The data entry screen for maintaining the appliction's staff    '
'           (worker) table.                                                 '
'           This form is only available to admin users.                     '
' Includes:                                                                 '
'   CnnInc.asp          - Connects to the database.                         '
'   ValidUserInc.asp    - Code to lookup user ID and PW sent to the page.   '
'   DefStylesInc.asp    - Contains DHTML styles common in the application.  '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<%
'==============================================================================
' Server side actions, data operations:
'==============================================================================
Dim madoCmd
Dim madoRsLogin
Dim strHTML
Dim mstrPageTitle
Dim mstrFormAction
Dim mblnDuplicateID
Dim mintI
Dim mstrRole
Dim mstrSecurityRoles
Dim mintCnt
Dim mlngLogRecordID
Dim madoRs
Dim madCmd
Dim mstrPassword
Dim mstrNewLoginAdded
Dim mblnForceRandPWonAdd
Dim mstrPasswordRequirement
Dim mstrErrorMessage

mstrPasswordRequirement = GetAppSetting("PasswordRequirement")
If InStr(mstrPasswordRequirement,"[FRPW]") > 0 Then
    mblnForceRandPWonAdd = True
Else
    mblnForceRandPWonAdd = False
End If
mstrPageTitle = "User Record"
mstrFormAction = ReqForm("FormAction")

mblnDuplicateID = False
mlngLogRecordID = 0
mstrNewLoginAdded = "N"
Set madoRsLogin = Server.CreateObject("ADODB.Recordset")
Select Case mstrFormAction
    Case "Add"  'Adding a new user:
        mstrPageTitle = mstrPageTitle & ": Add"

    Case "Edit" 'Loading an existing user to edit:
        mlngLogRecordID = ReqForm("logRecordID")
        Call GetLoginForEdit()
        
    Case "AddSave"  'Save a new user login:
        Call SaveLoginToDb()

        If mblnDuplicateID Then 'The login already existed in the database.
            mstrPageTitle = mstrPageTitle & ": Add"
            mstrNewLoginAdded = "N"
            mstrFormAction = "AddDuplicate"
        Else 'The user login saved ok.
            'Save the assigned security roles:
            Call SaveSecurityToDb()
            'Retrieve the record back for editing:
            Call GetLoginForEdit()
        End If
		
    Case "EditSave" 'Save changes to an existing user login:
        Call SaveLoginToDb()
        If Not mblnDuplicateID Then 'The login saved ok.
            'Save the security roles for the user login:
            Call SaveSecurityToDb()
        End If
        'Retrieve the record back for editing:
        Call GetLoginForEdit()
End Select

'==============================================================================
'Server-side functions, subroutines:
'==============================================================================
Sub SaveLoginToDb()
    mstrPassword = ReqForm("logPassword")
    If mstrFormAction = "AddSave" Then
        'Adding new record:
        Set madoCmd = GetAdoCmd("spUserAdd")
        If mblnForceRandPWonAdd Then
            mstrPassword = CreatePW()
        End If
        mstrNewLoginAdded = "Y"
    ElseIf mstrFormAction = "EditSave" Then
        'Editing existing record:
        Set madoCmd = GetAdoCmd("spUserUpd")
        AddParmIn madoCmd, "@logRecordID", adInteger, 0, ReqForm("logRecordID")
    End If

    AddParmIn madoCmd, "@logUserLogin", adVarChar, 20, ReqForm("logUserLogin")
    AddParmIn madoCmd, "@logPassword", adVarchar, 60, Encrypt(mstrPassword, UCase(ReqForm("logUserLogin")))
    AddParmIn madoCmd, "@logFirstName", adVarchar, 50, ReqForm("logFirstName")
    AddParmIn madoCmd, "@logMiddleName", adVarchar, 50, ReqForm("logMiddleName")
    AddParmIn madoCmd, "@logLastName", adVarchar, 50, ReqForm("logLastName")
    'AddParmIn madoCmd, "@logAdmin", adBoolean, 0, ReqForm("logAdmin")
    'AddParmIn madoCmd, "@logQA", adBoolean, 0, ReqForm("logQA")
    AddParmIn madoCmd, "@logAlias", adInteger, 0, ReqForm("logAliasID")
    AddParmOut madoCmd, "@ReturnVal", adInteger, 0
    'Call ShowCmdParms(madoCmd) '***DEBUG
    madoCmd.Execute
    'Check to make sure the user did not attempt to create a duplicate login:
    If madoCmd.Parameters("@ReturnVal").Value = -1 Then
        mblnDuplicateID = True
        mstrErrorMessage = "MsgBox ""The user login " & ReqForm("logUserLogin") & " already exists."" & vbcrlf & vbcrlf & ""Please enter a unique user login ID."", vbInformation, ""Save User Login"""
    ElseIf madoCmd.Parameters("@ReturnVal").Value = -2 Then
        mblnDuplicateID = True
        mstrErrorMessage = "MsgBox ""An error occured changing the user login " & ReqForm("logUserLogin") & "."" & vbcrlf & vbcrlf & ""Please enter a unique user login ID."", vbInformation, ""Save User Login"""
    Else
        mlngLogRecordID = madoCmd.Parameters("@ReturnVal").Value
        mblnDuplicateID = False
    End If
    Set madoCmd = Nothing
End Sub

Function CreatePW()
    Dim intRandom, intNumbers, intULetters, intSpecial
    Dim strPassword, intLLetters
    Dim intI
    
    strPassword = ""
    intNumbers = 0 
    intULetters = 0 
    intLLetters = 0
    intSpecial = 0
    For intI = 1 To 10000
        Randomize()
        intRandom = Int(Rnd()*1000)
        If ((intRandom>=35 And intRandom<=43) Or _
            (intRandom>=91 And intRandom<=94) Or _
            intRandom=33 Or intRandom=47) And intSpecial = 0 Then
            
            strPassword = strPassword & Chr(intRandom)
            intSpecial = intSpecial + 1
        End If
        If intRandom>=48 And intRandom<=57 And intNumbers = 0 Then
            strPassword = strPassword & Chr(intRandom)
            intNumbers = intNumbers + 1
        End If
        If intRandom>=65 And intRandom<=90 And intULetters < 3 Then
            strPassword = strPassword & Chr(intRandom)
            intULetters = intULetters + 1
        End If
        If intRandom>=97 And intRandom<=122 And intLLetters < 3 Then
            strPassword = strPassword & Chr(intRandom)
            intLLetters = intLLetters + 1
        End If
        If Len(strPassword) = 8 Then Exit For
    Next
    CreatePW = strPassword
End Function

Sub SaveSecurityToDb()
    Dim strSecurityRoles
    Dim strRole
    Dim intCnt
    
    strSecurityRoles = ReqForm("SecurityRoles")
    intCnt = 1
    strRole = Parse(strSecurityRoles, ";", intCnt)
    Do While strRole <> "" And Not mblnDuplicateID 
        If mstrFormAction = "AddSave" Then
            'Adding new record:
            Set madoCmd = GetAdoCmd("spSecurityRoleAdd")
            AddParmIn madoCmd, "@logRecordID", adInteger, 0, mlngLogRecordID
            AddParmIn madoCmd, "@rolRecordID", adInteger, 0, strRole
        ElseIf mstrFormAction = "EditSave" Then
            'Editing existing record:
            Set madoCmd = GetAdoCmd("spSecurityRoleUpd")
            AddParmIn madoCmd, "@logRecordID", adInteger, 0, mlngLogRecordID
            AddParmIn madoCmd, "@rolRecordID", adInteger, 0, strRole
            AddParmIn madoCmd, "@intCnt", adInteger, 0, intCnt
        End If
		'Call ShowCmdParms(madoCmd) '***DEBUG
		madoCmd.Execute
		Set madoCmd = Nothing
		intCnt = intCnt + 1
		strRole = Parse(strSecurityRoles, ";", intCnt)
	Loop
End Sub

Sub GetLoginForEdit()
    'Retrieves a user login record for editing:
    mstrPageTitle = mstrPageTitle & ": Edit"
    Set madoCmd = GetAdoCmd("spUserGet")
        AddParmIn madoCmd, "@logRecordID", adInteger, 0, mlngLogRecordID
        AddParmIn madoCmd, "@logUserLogin", adVarchar, 60, NULL
        madoRsLogin.Open madoCmd, , adOpenForwardOnly, adLockReadOnly
    Set madoCmd = Nothing
    mstrFormAction = "Edit"
End Sub

'==============================================================================
' Server-side classes:
'==============================================================================
'None
%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="HTML 4.0">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>


<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Option Explicit
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mlngTimerIDS

Sub window_onload()
    Dim oDictObj, oOption, strAliasNames
    
    Call SizeAndCenterWindow(767, 520, True)
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>
   
    <%
        If mblnDuplicateID Then
            Response.Write vbCrLf
            Response.Write mstrErrorMessage
            Response.Write vbCrLf
        End If
    %>
    
    strAliasNames = ""
    For Each oDictObj In window.parent.opener.mdctAliasIDs
        If InStr(strAliasNames, "[" & window.parent.opener.mdctAliasIDs(oDictObj) & "]") = 0 Then
	        Set oOption = Document.createElement("OPTION")
	        oOption.Value = oDictObj
	        oOption.Text = window.parent.opener.mdctAliasIDs(oDictObj)
	        cboAlias.options.add oOption
	        Set oOption = Nothing
	        strAliasNames = strAliasNames & "[" & window.parent.opener.mdctAliasIDs(oDictObj) & "]"
	    End If
    Next

    divDisplayPassword.style.left = -1000

	Call FillControls
    Call ShowPage(True)
    txtLoginID.focus

    If "<%=mblnForceRandPWonAdd%>" = "True" Then
        If "<%=mstrFormAction%>" = "Add" Or "<%=mstrFormAction%>" = "AddDuplicate" Then
            txtPassword.value = ""
            txtConfirmPassword.value = ""
            txtPassword.disabled = True
            txtConfirmPassword.disabled = True
            txtPassword.style.left = -1000
            txtConfirmPassword.style.left = -1000
            lblPassword.style.left = -1000
            lblConfirmPassword.style.left = -1000
            divDisplayPassword.style.left = 15
            lblNewPassword.innerText = "A random password will be generated when new login is confirmed."
        End If
        If "<%=mstrNewLoginAdded%>" = "Y" Then
            divDisplayPassword.style.left = 15
            lblNewPassword.innerText = "Password for " & txtLoginID.value & " is:   " & txtPassword.value
            txtPassword.style.left = -1000
            txtConfirmPassword.style.left = -1000
            lblPassword.style.left = -1000
            lblConfirmPassword.style.left = -1000
            cmdSave.disabled = True
        End If
    End If
End Sub

<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    window.close
End Sub

Sub ShowPage(blnShow)
    If blnShow Then
        divPageFrame.style.visibility = "visible"
        divSecurityRoleFrame.style.visibility = "visible"
        lblStatus.style.visibility = "hidden"
        PageBody.style.cursor = "default"
    Else
        divPageFrame.style.visibility = "hidden"
        divSecurityRoleFrame.style.visibility = "hidden"
        lblStatus.style.visibility = "visible"
        PageBody.style.cursor = "wait"
    End If
End Sub

Sub cmdAdd_onclick()
    Call ShowPage(False)
    Form.SecurityRoles.Value = ""
    Form.FormAction.Value = "Add"
    Form.action = "UsersAddEdit.asp"
    Form.target = ""
    Form.submit
End Sub

Sub cmdSave_onclick()
    Call FillForm
    
    If Trim(txtLoginID.Value) = vbNullString Then
        MsgBox "Please enter a value for the Login ID.", vbInformation, "Save User Record"
        txtLoginID.focus
    ElseIf Len(Trim(txtLoginID.Value)) < <%=GetAppSetting("MinUserIdLen")%> Or Len(Trim(txtLoginID.Value)) > 20 Then
        MsgBox "The user's login name must be between <%=GetAppSetting("MinUserIdLen")%> and 20 characters.", vbInformation, "Save User Record"
        txtLoginID.focus
    'ElseIf Len(Trim(txtPassword.value)) < <%=GetAppSetting("MinPwLen")%> Or Len(Trim(txtPassword.value)) > 20 Then
    '    MsgBox "The user's password must be between <%=GetAppSetting("MinPwLen")%> and 20 characters.", vbInformation, "Save User Record"
    '    txtPassword.focus
    ElseIf Trim(txtPassword.value) <> Trim(Form.logPassword.Value) And Trim(txtPassword.value) <> Trim(txtConfirmPassword.value) Then
        MsgBox "The password confirmation did not match.", vbInformation, "Save User Record"
        txtConfirmPassword.focus
    ElseIf lstAssignedRoles.options.Length = 0 Then
		MsgBox "The User must be assigned at least one Security Role"
		lstSecurityRoles.focus
    ElseIf Len(Trim(txtEmpLastName.value)) = 0 Then
		MsgBox "A Last Name must be entered for the employee", vbInformation, "Save User Record"
		txtEmpLastName.focus
	Else
	    'Passed all validations
	    <%'
	    'If Form.HoldPassword.value <> Form.logPassword.value And Form.FormAction.Value <> "Add" Then
        '    Form.action = "PasswordVerification.asp"
        '    Form.Target = "SaveFrame"
        '    Form.PasswordVerified.Value = "N"
        '    Form.PasswordVerificationType.value = "ChangePW"
        '    Form.PWUserID.value = txtLoginID.value
        '    Form.NewPassword.value = txtPassword.value
        '    mlngTimerIDS = window.setInterval("CheckForCompletion",100)
        '    SaveWindow.style.left = 5
        '    divSecurityRoleFrame.style.left = -1000
        '    divPageFrame.style.left = -1000
        '    Form.Submit
        'Else
        ' %>
            If Form.FormAction.Value = "Add" Then
                Form.FormAction.Value = "AddSave"
            ElseIf Form.FormAction.Value = "Edit" Then
                Form.FormAction.Value = "EditSave"
            End If
            Form.Action = "UsersAddEdit.asp"
            Form.target = ""
            Call ShowPage(False)
            Form.Submit
        'End If
    End If
End Sub

Function CheckForCompletion()
    Dim strWarning
    If Form.PasswordVerified.Value = "Y" Then
        window.clearInterval mlngTimerIDS
        If Left(Form.PasswordMessage.value,13) = "WARNING ONLY:" Then 
            strWarning = Trim(Replace(Form.PasswordMessage.value,"WARNING ONLY:",""))
            Form.PasswordMessage.value = ""
        End If
        If Form.PasswordMessage.value = "" Then
            If strWarning <> "" Then
                MsgBox strWarning, vbInformation, "Password Expiration"
            End If

            If Form.FormAction.Value = "Add" Then
                Form.FormAction.Value = "AddSave"
            ElseIf Form.FormAction.Value = "Edit" Then
                Form.FormAction.Value = "EditSave"
            End If
            Form.Action = "UsersAddEdit.asp"
            Form.Target = ""
            Call ShowPage(False)
            Form.Submit
        Else
            MsgBox Form.PasswordMessage.value, vbInformation, "Password Change"
            txtPassword.focus
        End If
    End If
End Function 

Sub ResetFrames()
    divSecurityRoleFrame.style.left = 5
    divPageFrame.style.left = 5
End Sub

Sub cmdCancel_onclick()
    Dim intResp 

    If IsChanged() Then
        intResp = MsgBox("Information has been modified on the form." & vbCrLf & vbCrLf & "Close the form without saving changes?", vbYesNo + vbQuestion, "Close User Add/Edit")
        If intResp = vbNo Then
            Exit Sub
        End If
    End If
    lblStatus.innerText = "Returning to main menu..."
    Call ShowPage(False)
    Form.SecurityRoles.value = ""
    Form.logRecordID.value = 0
    Form.Action = "UsersSelect.asp"
    Form.target = ""
    Form.Submit
End Sub

Function IsChanged()
    Dim blnChanged

    If txtLoginID.value <> Form.logUserLogin.value Then
        blnChanged = True
    ElseIf txtPassword.value <> Form.logPassword.Value Then
        blnChanged = True
    ElseIf txtEmpLastName.value <> Form.logLastName.value Then
        blnChanged = True
    ElseIf txtEmpFirstName.value <> Form.logFirstName.value Then
        blnChanged = True
    ElseIf txtEmpMiddleName.value <> Form.logMiddleName.value Then
        blnChanged = True
    ElseIf cboAlias.value <> Form.logAliasID.value Then
        blnChanged = True
    ElseIf SaveRolesToString() <> Form.SecurityRoles.Value Then
        blnChanged = True
    Else
        blnChanged = False
    End If
    
    IsChanged = blnChanged
End Function

Sub divPageFrame_onkeydown()
    If window.event.keyCode = 13 Then
        Call cmdSave_onclick()
    ElseIf window.event.keyCode = 27 Then
        Call cmdCancel_onclick()
    End If
End Sub

Sub FillControls()
    If Form.FormAction.Value = "Edit" Or Form.FormAction.Value = "AddDuplicate" Then
        txtUserID.value = Form.logRecordID.value
        txtLoginID.value = Form.logUserLogin.value
        txtPassword.value = Form.logPassword.value
        txtEmpLastName.value = Form.logLastName.value
        txtEmpFirstName.value = Form.logFirstName.value
        txtEmpMiddleName.value = Form.logMiddleName.value
        If Trim(Form.logAliasID.value) = "" Then
            cboAlias.value = 0
        Else
            cboAlias.value = Form.logAliasID.value
        End If
        If Form.FormAction.Value = "AddDuplicate" Then
            Form.FormAction.Value = "Add"
        End If
        'Store the initial string of security roles:
        Form.SecurityRoles.value = SaveRolesToString()
    End If
End Sub

Sub FillForm()
	If txtUserID.value <> "" Then
		Form.logRecordID.value = txtUserID.value
	Else
		Form.logRecordID.value = 0
	End If
    Form.logUserLogin.value = txtLoginID.value
    Form.logPassword.value = txtPassword.value
    Form.logFirstName.value = txtEmpFirstName.value
    Form.logMiddleName.value = txtEmpMiddleName.value
    Form.logLastName.value = txtEmpLastName.value
    Form.logAliasID.value = cboAlias.value
	Form.SecurityRoles.value = SaveRolesToString()
End Sub

Function SaveRolesToString()
    Dim intI
    Dim strRoles

	'Form.logAdmin.Value = 0
	'Form.logQA.Value = 0

    strRoles = ""
    For intI = 0 To lstAssignedRoles.options.length - 1
        If strRoles <> "" Then
            strRoles = strRoles & ";"
        End If
        strRoles = strRoles & lstAssignedRoles.options(intI).Value

	    'If .Value =  3 then Administrator was selected
	    'If lstAssignedRoles.options(intI).Value = 3 Then
		'    Form.logAdmin.Value = 1
	    'ElseIf lstAssignedRoles.options(intI).Value = 4 Then
		'    Form.logQA.Value = 1
	    'End If
    Next
    SaveRolesToString = strRoles
End Function

Sub RoleMover_onclick(oButton)
    Dim oOption
    Dim iSelected
    Dim oSource
    Dim oDestin
    
    If oButton.ID = "cmdLeftToRight" Then
        Set oSource = lstSecurityRoles
        Set oDestin = lstAssignedRoles
    Else
        Set oSource = lstAssignedRoles
        Set oDestin = lstSecurityRoles
    End If

    iSelected = oSource.selectedIndex
    If iSelected < 0 Then
        Exit Sub
    End If
    cmdLeftToRight.disabled = true
    cmdRightToLeft.disabled = true
    Set oOption = document.CreateElement("OPTION")
    oOption.Value = oSource.options(iSelected).Value
    oOption.Text = oSource.options(iSelected).Text
    oDestin.options.Add(oOption)
    oDestin.selectedIndex = oDestin.options.length - 1
    oSource.options.remove(iSelected)
    If iSelected <= oSource.options.length - 1 Then
        oSource.selectedIndex = iSelected
    Else
        If oSource.options.length > 0 Then
            oSource.selectedIndex = iSelected - 1
        End If
    End If
    cmdLeftToRight.disabled = false
    cmdRightToLeft.disabled = false
End Sub
<% 
'Sub cboAlias_onclick()
'    Call cboInteliOnClick()
'End Sub
'Sub cboAlias_onfocus()
'    Call cboInteliOnFocus()
'End Sub
'Sub cboAlias_onblur()
'    Call cboInteliOnBlur()
'End Sub
'Sub cboAlias_onkeypress()
'    Call cboInteliOnKeyPress(cboAlias)
'End Sub
'Sub cboAlias_onkeydown()
'    Call cboInteliOnKeyDown(cboAlias)
'End Sub
%>
</SCRIPT>
 <!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody style="cursor:wait">
    <DIV id=Header class=DefTitleArea style="WIDTH:730; TOP:5; LEFT:5">
        <SPAN id=lblAppTitle class=DefTitleText style="WIDTH:730">
            <%=mstrPageTitle%>
        </SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>
    <% Call WriteNavigateControls(-1,30,gstrBackColor) %>
    <SPAN id="lblStatus" class=DefLabel 
        style="LEFT:25; WIDTH:300; TOP:55; FONT-SIZE:10pt; visibility:hidden">
        Accessing database...
    </SPAN>
    
    <DIV ID=SaveWindow Name=SaveWindow 
        style="position:absolute;
            TOP:46; 
            LEFT:-1000;
            WIDTH:408; 
            HEIGHT:425; 
            COLOR:black; 
            BORDER-STYLE:solid;
            BORDER-WIDTH:2;
            BORDER-COLOR:<%=gstrBorderColor%>;
            BACKGROUND-COLOR:<%=gstrBackColor%>">
        <IFRAME ID=SaveFrame Name=SaveFrame src="blank.html" style="top:0;height:220;left:0;width:400">
        </IFRAME>
        <SPAN id=lblSavingMessage class=DefLabel style="top:50;height:45;left:0;width:400"><CENTER><BIG><B></BIG>Verifying Password Strength...</B></BIG></CENTER></SPAN>
    </DIV> 
    <DIV id=divPageFrame class=DefPageFrame style="LEFT:5; HEIGHT:120; WIDTH:730; TOP:46; visibility:hidden">

        <SPAN id=lblUserID class=DefLabel style="LEFT:15; WIDTH:55; TOP:10">
            Record ID:
        </SPAN>
        <INPUT id=txtUserID TYPE=text title="Internal record ID number" 
            style="LEFT:90; WIDTH:100; TOP:10; BACKGROUND-COLOR:buttonface" 
            onfocus="CmnTxt_onfocus(txtLastName)"
            disabled=true tabIndex=1 cols=26 NAME="txtUserID">

        <SPAN id=lblLoginID class=DefLabel style="LEFT:15; WIDTH:70; TOP:35">
            User Login ID:
        </SPAN>
        <INPUT id=txtLoginID TYPE=text title="User's login ID" 
            style="LEFT:90; WIDTH:100; TOP:35" 
            onfocus="CmnTxt_onfocus(txtLoginID)"
            tabIndex=2 cols=26 NAME="txtLoginID">

        <SPAN id=lblPassword class=DefLabel style="LEFT:15; WIDTH:65; TOP:60">
            Password:
        </SPAN>
        <INPUT id=txtPassword type=password title="Enter a login password" 
            style="LEFT:90; WIDTH:100; TOP:60" 
            onfocus="CmnTxt_onfocus(txtPassword)"
            tabIndex=3 cols=26 NAME="txtPassword">

        <SPAN id=lblConfirmPassword class=DefLabel style="LEFT:15; WIDTH:100; TOP:85">
            Confirm PW:
        </SPAN>
        <INPUT id=txtConfirmPassword type=password title="Retype the password for confirmation" 
            style="LEFT:90; WIDTH:100; TOP:85" 
            onfocus="CmnTxt_onfocus(txtConfirmPassword)"
            tabIndex=4 cols=26 NAME="txtConfirmPassword">

        <DIV id=divDisplayPassword class=DefPageFrame style="position:absolute;LEFT:0; HEIGHT:70; WIDTH:310; TOP:60;border-style:none">
            <SPAN id=lblNewPassword class=DefLabel style="LEFT:5; WIDTH:280; TOP:0;font-weight:bold;text-align:center">
                  
            </SPAN>
        </DIV>
        <SPAN id=lblEmpLastName class=DefLabel style="LEFT:305; WIDTH:65; TOP:10">Last Name</SPAN>
        <INPUT id=txtEmpLastName TYPE=text tabIndex=5 
            style="LEFT:305; WIDTH:100; TOP:25;background-color:white" maxlength=50 NAME="txtEmpLastName">
        <SPAN id=lblEmpFirstName class=DefLabel style="LEFT:410; WIDTH:65; TOP:10">First Name</SPAN>
        <INPUT id=txtEmpFirstName TYPE=text tabIndex=6
            style="LEFT:410; WIDTH:100; TOP:25;background-color:white" maxlength=50 NAME="txtEmpFirstName">
        <SPAN id=lblEmpMiddleName class=DefLabel style="LEFT:515; WIDTH:65; TOP:10">Middle Name</SPAN>
        <INPUT id=txtEmpMiddleName TYPE=text tabIndex=7
            style="LEFT:515; WIDTH:100; TOP:25;background-color:white" maxlength=50 NAME="txtEmpMiddleName">

        <SPAN id=lblAlias class=DefLabel style="LEFT:305; WIDTH:65; TOP:55">
            Alias
        </SPAN>
        <SELECT id=cboAlias title="Assign this user ID to a staff person"
            style="LEFT:305; WIDTH:210; TOP:70"
            tabIndex=-1 NAME="cboAlias">
            <OPTION VALUE=0 SELECTED>
       </SELECT>
    </DIV>

	<DIV id=divSecurityRoleFrame class=DefPageFrame style="LEFT:5; HEIGHT:260; WIDTH:730; TOP:165; visibility:hidden">
        <SPAN id=lblInstructions class=DefLabel style="LEFT:500; WIDTH:200; TOP:5">
            <B>Instructions:</B>
        </SPAN>
        <SPAN id=lblInstructionsText class=DefLabel 
            style="LEFT:500; WIDTH:200; HEIGHT:75; TOP:20">
            Move items in or out of the list of classifications on the right. Click [Save] to keep your changes. Click [Close] to abandon changes.
        </SPAN>

        <SPAN id=lblJobClasses class=DefLabel style="LEFT:15; WIDTH:175; TOP:5">
            Security Roles:
        </SPAN>
        <SELECT id=lstSecurityRoles title="List of Available Security Roles"
            style="LEFT:15; WIDTH:175; TOP:25; HEIGHT:215" 
            tabIndex=1 size=13 TYPE="select-one" NAME="lstSecurityRoles">
            <%
            Set madoCmd = GetAdoCmd("spGetAllSecurityRoles")
                If mlngLogRecordID = 0 Then
                    AddParmIn madoCmd, "@logRecordID", adInteger, 0, NULL
                Else
                    AddParmIn madoCmd, "@logRecordID", adInteger, 0, mlngLogRecordID
                End If
                'Call ShowCmdParms(madCmd) '***DEBUG
            Set madoRs = GetAdoRs(madoCmd)
            Do While Not madoRs.EOF And Not madoRs.BOF
                Response.Write "<OPTION VALUE=" & madoRs.Fields("rolID").Value & ">" & madoRs.Fields("rolName").Value
                madoRs.MoveNext
            Loop
            madoRs.Close
            Set madoRs = Nothing
            Set madoCmd = Nothing
            %>
        </SELECT>

        <SPAN id=lblAssignedClasses class=DefLabel style="LEFT:305; WIDTH:175; TOP:5">
            Assigned Security Roles:
        </SPAN>
        <SELECT id=lstAssignedRoles title="List Current Security Roles"
            style="LEFT:305; WIDTH:175; TOP:25; HEIGHT:215" 
            tabIndex=1 size=13 TYPE="select-one" NAME="lstAssignedRoles">
            <%
            If mlngLogRecordID <> 0 Then
                Set madoCmd = GetAdoCmd("spGetSecurityRoles")
                    AddParmIn madoCmd, "@logRecordID", adInteger, 0, mlngLogRecordID
                    'Call ShowCmdParms(madoCmd) '***DEBUG
                Set madoRs = GetAdoRs(madoCmd)
                Do While Not madoRs.EOF And Not madoRs.BOF
                    Response.Write "<OPTION VALUE=" & madoRs.Fields("rolID").Value & ">" & madoRs.Fields("rolName").Value
                    madoRs.MoveNext
                Loop
                madoRs.Close
                Set madoRs = Nothing
                Set madoCmd = Nothing
            End If
            %>
        </SELECT>

        <BUTTON id=cmdLeftToRight class=DefBUTTON title="Add new value to the list" 
            style="LEFT:210; TOP:105; WIDTH:70; HEIGHT:20"
            onclick="RoleMover_onclick(cmdLeftToRight)"
            accessKey=R
            tabIndex=7>--&gt
        </BUTTON>
        <BUTTON id=cmdRightToLeft class=DefBUTTON title="Delete the selected value" 
            style="LEFT:210; TOP:135; WIDTH:70; HEIGHT:20"
            onclick="RoleMover_onclick(cmdRightToLeft)"
            accessKey=R
            tabIndex=7>&lt--
        </BUTTON>
       	<%
       	If mstrFormAction <> "Add" Then %>
		       <BUTTON id=cmdAdd class=DefBUTTON title="Add another user" 
		           style="LEFT:500; TOP:220; HEIGHT:20; WIDTH:70"
		           accessKey=R
		           tabIndex=9>Add
		       </BUTTON><%
		End If %>
		
		<BUTTON id=cmdSave class=DefBUTTON title="Save changes for this user record" 
		    style="LEFT:575; TOP:220; HEIGHT:20; WIDTH:70"
		    accessKey=R
		    tabIndex=9>Save
		</BUTTON>
		<BUTTON id=cmdCancel class=DefBUTTON title="Cancel the current changes and close the form" 
		    style="LEFT:650; TOP:220; HEIGHT:20; WIDTH:70"
		    accessKey=R
		    tabIndex=9>Close
		</BUTTON>
    </DIV>
</BODY>

<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="" ID=Form>
    <% 
    Call CommonFormFields()
    WriteFormField "FormAction", mstrFormAction
    WriteFormField "SecurityRoles", ReqForm("SecurityRoles")
    WriteFormField "LoginID", ReqForm("LoginID")
    WriteFormField "StaffID", ReqForm("StaffID")
    WriteFormField "LastName", ReqForm("LastName")
    WriteFormField "FirstName", ReqForm("FirstName")
    WriteFormField "MiddleName", ReqForm("MiddleName")

    Select Case mstrFormAction
        Case "Edit", "EditSave", "AddSave"
            WriteFormField "logRecordID", madoRsLogin.Fields("logRecordID").Value
            WriteFormField "logUserLogin", madoRsLogin.Fields("logUserLogin").Value
            WriteFormField "logPassword",  Decrypt(LCase(madoRsLogin.Fields("logPassword").Value), UCase(madoRsLogin.Fields("logUserLogin").Value))
            WriteFormField "HoldPassword", Decrypt(LCase(madoRsLogin.Fields("logPassword").Value), UCase(madoRsLogin.Fields("logUserLogin").Value))
            WriteFormField "logFirstName", madoRsLogin.Fields("logFirstName").Value
            WriteFormField "logMiddleName", madoRsLogin.Fields("logMiddleName").Value
            WriteFormField "logLastName", madoRsLogin.Fields("logLastName").Value
            'If madoRsLogin.Fields("logSecurityRoleID").Value = 1 Then
            '    WriteFormField "logAdmin", "True"
            'Else
            '    WriteFormField "logAdmin", "False"
            'End If
            'If madoRsLogin.Fields("logSecurityRoleID").Value = 2 Then
            '    WriteFormField "logQA", "1"
            'Else
            '    WriteFormField "logQA", "0"
            'End If
            WriteFormField "logAliasID", madoRsLogin.Fields("logAliasID").Value
        Case Else
            If mblnDuplicateID Then
                WriteFormField "logRecordID", ReqForm("logRecordID")
                WriteFormField "logUserLogin", ReqForm("logUserLogin")
                WriteFormField "logPassword", ReqForm("logPassword")
                WriteFormField "HoldPassword", ReqForm("logPassword")
                WriteFormField "logFirstName", ReqForm("logFirstName")
                WriteFormField "logMiddleName", ReqForm("logMiddleName")
                WriteFormField "logLastName", ReqForm("logLastName")
                'WriteFormField "logAdmin", ReqForm("logAdmin")
                'WriteFormField "logQA", ReqForm("logQA")
                WriteFormField "logAliasID", ReqForm("logAliasID")
            Else
                WriteFormField "logRecordID", 0
                WriteFormField "logUserLogin", ""
                WriteFormField "logPassword", ""
                WriteFormField "HoldPassword", ""
                WriteFormField "logFirstName", ""
                WriteFormField "logMiddleName", ""
                WriteFormField "logLastName", ""
                'WriteFormField "logAdmin", 0
                'WriteFormField "logQA", 0
                WriteFormField "logAliasID", 0
            End If
    End Select
    WriteFormField "PasswordVerified", ""
    WriteFormField "PasswordVerificationType", ""
    WriteFormField "PasswordMessage", ""
    WriteFormField "PWUserID", ""
    WriteFormField "NewPassword", ""

    Response.Write strHTML

    If ReqForm("FormAction") = "Edit" Then
        madoRsLogin.Close
    End If
    Set madoRsLogin = Nothing
    gadoCon.Close
    Set gadoCon = Nothing
    %>
</FORM>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncInteliType.asp"-->
<!--#include file="IncNavigateControls.asp"-->
</HTML>
