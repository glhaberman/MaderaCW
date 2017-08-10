<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: UsersSelect.asp                                                 '
'  Purpose: This screen allows admin staff to add or update application     '
'           user ID's (also known as the reviewers).                        '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
Dim adRs
Dim strSQL
Dim mstrPageTitle
Dim strResetResult
Dim mintTblWidth
Dim strPassword, strLogin
Dim mstrPasswordRequirement
Dim mblnForceRandPW

mstrPageTitle = "Add/Edit User Logins"
mstrPasswordRequirement = GetAppSetting("PasswordRequirement")
If InStr(mstrPasswordRequirement,"[FRPW]") > 0 Then
    mblnForceRandPW = True
Else
    mblnForceRandPW = False
End If

Set adRs = Server.CreateObject("ADODB.Recordset")

strDeleteResult = ""
strResetResult = ""
If ReqForm("FormAction") = "Delete" Then
	Set madoCmd = GetAdoCmd("spUserDel")
        AddParmIn madoCmd, "@UserID", adInteger, 0, ReqForm("logRecordID")
        AddParmOut madoCmd, "@UseCheck", adInteger, 0
        madoCmd.Execute
        If madoCmd.Parameters("@UseCheck").Value = 1 Then 'ID was in use, unable to delete.
            strDeleteResult = """Unable to delete the User record."" & vbcrlf & ""The record is being used in a case review."""
        ElseIf madoCmd.Parameters("@UseCheck").Value = 2 Then 'Id was a supervisor.
            strDeleteResult = """Unable to delete the User record."" & vbcrlf & ""The record is being used as a supervisor for other staff."""
        Else
            strDeleteResult = """User Login record deleted."""
        End If
    Set madoCmd = Nothing
'use spResetPassword to reset the user's password to 'password'
ElseIf ReqForm("FormAction") = "Reset" Then
    Set madoCmd = GetAdoCmd("spUserGet")
        AddParmIn madoCmd, "@logRecordID", adInteger, 0, ReqForm("logRecordID")
        AddParmIn madoCmd, "@logUserLogin", adVarchar, 60, NULL
        'Call ShowCmdParms(madoCmd) '***DEBUG
        adRs.Open madoCmd, , adOpenForwardOnly, adLockReadOnly
    Set madoCmd = Nothing
    strLogin = adRs.Fields("logUserLogin").Value
    adRs.Close

    If mblnForceRandPW = True Then
        strPassword = CreatePW()
    Else
        strPassword = "password"
    End If
	Set madoCmd = GetAdoCmd("spResetPassword")
		AddParmIn madoCmd, "@UserRecordID", adInteger, 0, ReqForm("logRecordID")
		AddParmIn madoCmd, "@Password", adVarchar,  60, Encrypt(strPassword, UCase(strLogin))
        'Call ShowCmdParms(madoCmd) '***DEBUG
		madoCmd.Execute
		strResetResult = """ User password reset to " & strPassword & """"
	Set madoCmd = Nothing
End If

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
%>
<HTML><HEAD>
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <META name=vs_targetSchema content="HTML 4.0">
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload()

    Call SizeAndCenterWindow(485, 375, True)
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>

    PageFrame.disabled = False
    
    <%
    If strDeleteResult <> "" Then
        Response.write "MsgBox " & strDeleteResult & ", vbInformation, ""Delete User Record"""
    ElseIF strResetResult <> "" Then
		Response.write "MsgBox " & strResetResult & ", vbInformation, ""Reset User Password"""
    End If
    %>
    lstUser.disabled = False
    txtLastName.value = Form.LastName.value
    txtFirstName.value = Form.FirstName.value
    txtMiddleName.value = Form.MiddleName.value
    cmdEdit.focus
End Sub

<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    mblnSetFocusToMain = False
    window.close
End Sub

Sub cmdClose_onclick()
    Call window.opener.ManageWindows(6,"Close")
End Sub

<%'If Main has not been closed, set focus back to it.%>
Sub window_onbeforeunload()
    If mblnSetFocusToMain = True Then
        window.opener.focus
    End If
End Sub

Sub lstUser_ondblclick()
    Call cmdEdit_onclick()
End Sub

Sub cmdFind_onclick()
	Form.LoginID.value = Trim(txtUserID.value)
    Form.LastName.value = Trim(txtLastName.value)
    Form.FirstName.value = Trim(txtFirstName.value)
    Form.MiddleName.value = Trim(txtMiddleName.value)
    Form.FormAction.value = "Find"
    Form.Action = "UsersSelect.asp"
    mblnSetFocusToMain = False
    Form.Submit
End Sub

Sub cmdAdd_onclick()
	Form.logRecordID.value = 0
    Form.FormAction.value = "Add"
    Form.Action = "UsersAddEdit.asp"
    mblnSetFocusToMain = False
    Form.Submit
End Sub

Sub cmdEdit_onclick()
	If Not IsItemSelected() Then
        Exit Sub
    End If
    Form.FormAction.value = "Edit"
    Form.logRecordID.value = Document.all("loginID" & Form.SelectedIndex.Value).innerText 
	Form.Action = "UsersAddEdit.asp"
    mblnSetFocusToMain = False
    Form.Submit
End Sub

Sub cmdDelete_onclick()
	If Not IsItemSelected() Then
        Exit Sub
    End If
    lstUser.disabled = True
    Form.FormAction.value = "Delete"
    Form.logRecordID.value = Document.all("loginID" & Form.SelectedIndex.Value).innerText 
    Form.Action = "UsersSelect.asp"
    mblnSetFocusToMain = False
    Form.Submit
End Sub

Sub cmdPrint_onclick()
	If IsItemSelected() Then
        Form.logRecordID.value = Document.all("loginID" & Form.SelectedIndex.Value).innerText 
    Else
        Form.logRecordID.value = ""
    End If
    Form.Action = "RptUsersPrint.asp"
    mblnSetFocusToMain = False
    Form.Submit
End Sub

Sub Gen_onkeydown
    If window.event.keyCode = 13 Then
        call cmdFind_onclick()
    End If
End Sub

Sub Gen_focus(txtBox)
    txtBox.select
End Sub

Sub cmdReset_onclick
	If Not IsItemSelected() Then
        Exit Sub
    End If
    Form.LoginID.value = Trim(txtUserID.value)
    Form.LastName.value = Trim(txtLastName.value)
    Form.FirstName.value = Trim(txtFirstName.value)
    Form.MiddleName.value = Trim(txtMiddleName.value)
	lstUser.disabled = True
    Form.FormAction.value = "Reset"
    Form.logRecordID.value = Document.all("loginID" & Form.SelectedIndex.Value).innerText 
    Form.Action = "UsersSelect.asp"
    mblnSetFocusToMain = False
    Form.Submit
End Sub

Function IsItemSelected()
    IsItemSelected = False
    
    If lstUsers.Rows.length > 1 Then
        If IsNumeric(Form.SelectedIndex.Value) Then
            If Form.SelectedIndex.Value > 0 Then
                If Trim(Document.all("loginID" & Form.SelectedIndex.Value).innerText) <> "" Then
                    IsItemSelected = True
                End If
            End If
        End If 
    End If
End Function

Sub lstResults_onkeydown()
    <%
    'This code controls the behavior in the results DIV when the Up arrow,
    'Down arrow, Home, and End keys are pressed.  This code changes the
    'selected item as the user moves up and down in the list:
    %>
    If IsNumeric(Form.SelectedIndex.Value) Then
        Select Case Window.Event.keyCode
            Case 36 'home
                Window.event.returnValue = False
                lstUsers.rows(0).scrollIntoView
                Call Result_onclick(1)
            Case 35 'end
                Window.event.returnValue = False
                Call Result_onclick(lstUsers.Rows.Length - 1)
            Case 38 'Up
                If Form.SelectedIndex.Value > 1 Then
                    Window.event.returnValue = False
                    Call Result_onclick(Form.SelectedIndex.Value - 1)
                End If
            Case 40 'Down
                If Cint(Form.SelectedIndex.Value) < Cint(lstUsers.Rows.Length - 1) Then
                    Window.event.returnValue = False
                    Call Result_onclick(Form.SelectedIndex.Value + 1)
                End If
        End Select
    End If
End Sub

Sub Result_onclick(intRow)
    Dim strRow
    If PageFrame.disabled Then
        Exit Sub
    End If
    If IsNumeric(Form.SelectedIndex.Value) Then
        strRow = "ListRow" & Form.SelectedIndex.Value
        lstUsers.Rows(strRow).className = "TableRow"
        lstUsers.Rows(strRow).cells(0).tabindex = -1
    End If

    strRow = "ListRow" & intRow
    lstUsers.Rows(strRow).className = "TableSelectedRow"
    lstUsers.Rows(strRow).cells(0).focus
    lstUsers.Rows(strRow).cells(0).tabindex = 9

    Form.SelectedIndex.Value = intRow
End Sub

Sub Result_ondblclick(intRow)
    Call Result_onclick(intRow)
    Call cmdEdit_onclick
End Sub

</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"--></HEAD>

<BODY id=PageBody bottomMargin=5 leftMargin=5 rightMargin=5>
    
    <DIV id=Header class=DefTitleArea style="WIDTH:455; HEIGHT:40">
        <SPAN id=lblAppTitle class=DefTitleText style="WIDTH:455">
            <%=mstrPageTitle%>
        </SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>
    <% Call WriteNavigateControls(-1,30,gstrBackColor) %>
            
    <DIV id=PageFrame class=DefPageFrame style="WIDTH:455px; TOP:51px; HEIGHT:255px">

        <SPAN id=lblLastName class=DefLabel style="LEFT:10px; WIDTH:300px; TOP:5px">
            Enter a name or part of a name and click [Find], or leave the name fields blank and click [Find] to see a list of all users.
        </SPAN>
        
        <SPAN id=lblUserID class=DefLabel style="LEFT:10px; WIDTH:100px; TOP:45px">
            Search by Login ID:
        </SPAN>
        
        <TEXTAREA id=txtUserID title="Enter last name to search for" 
            style="LEFT:10px; WIDTH:100px; TOP:60px" 
            onkeydown="Gen_onkeydown" onfocus="Gen_focus(txtUserID)"
            tabIndex=1 cols=26 NAME="txtUserID"></TEXTAREA>

        <SPAN id=lblOr class=DefLabel style="LEFT:120px; WIDTH:20px; TOP:45px">
            <B>Or</B>
        </SPAN>

        <SPAN id="Span1" class=DefLabel style="LEFT:150px; WIDTH:100px; TOP:45px">
            By Last Name:
        </SPAN>
        
        <TEXTAREA id=txtLastName title="Enter last name to search for"
            style="LEFT:150px; WIDTH:100px; TOP:60px" 
            onkeydown="Gen_onkeydown" onfocus="Gen_focus(txtLastName)"
            tabIndex=1 cols=26 NAME="txtLastName"></TEXTAREA>

        <SPAN id=lblFirstName class=DefLabel style="LEFT:250px; WIDTH:75px; TOP:45px">
            First Name:
        </SPAN>
        
        <TEXTAREA id=txtFirstName title="Enter first name to search for" 
            style="LEFT:250px; WIDTH:75px; TOP:60px" 
            onkeydown="Gen_onkeydown" onfocus="Gen_focus(txtFirstName)"
            tabIndex=2 cols=26 NAME="txtFirstName"></TEXTAREA>

        <SPAN id=lblMiddleName class=DefLabel style="LEFT:325px; WIDTH:40px; TOP:45px">
            Middle:
        </SPAN>

        <TEXTAREA id=txtMiddleName title="Enter middle name to search for" 
            style="LEFT:325px; WIDTH:40px; TOP:60px" 
            onkeydown="Gen_onkeydown" onfocus="Gen_focus(txtMiddleName)"
            tabIndex=3 cols=26 NAME="txtMiddleName"></TEXTAREA>

        <BUTTON id=cmdFind class=DefBUTTON title="Add new user record" 
            style="LEFT:375px; WIDTH:70px; TOP:59px; HEIGHT:20px"
            accessKey=F tabIndex=4 type=button>
            <U>F</U>ind
        </BUTTON>

        <SPAN id=lblSearchResults class=DefLabel style="LEFT:10px; WIDTH:300px; TOP:90px">
            Search Results - select a record to modify, or click [Add]:
        </SPAN>
        <DIV id=lstUser class=TableDivArea title="Results of Find"
            style="LEFT:10px; WIDTH:435px; TOP:105px; HEIGHT:110px"
            tabIndex=5 >
            <%
            If ReqForm("FormAction") = "Find" Or ReqForm("FormAction") = "Reset" Then
	            Set madoCmd = GetAdoCmd("spGetUserFind")
                    AddParmIn madoCmd, "@LoginID", adVarChar, 20, ReqIsBlank("LoginID")
                    AddParmIn madoCmd, "@LastName", adVarChar, 50, ReqIsBlank("LastName")
                    AddParmIn madoCmd, "@FirstName", adVarChar, 50, ReqIsBlank("FirstName")
                    AddParmIn madoCmd, "@MiddleName", adVarChar, 50, ReqIsBlank("MiddleName")
                    adRs.Open madoCmd, , adOpenForwardOnly, adLockReadOnly
                Set madoCmd = Nothing 
			End If
            
            intLine = 0
            intResultCnt = 1
            mintTblWidth = 415
            
            Response.Write "<Table ID=lstUsers Border=0 Rules=rows Cols=4 Width=" & mintTblWidth & " CellSpacing=0 Style=""overflow: hidden; TOP:0""> " & vbCrLf
            Response.Write "<THEAD ID=ListHeader><TR ID=HeaderRow>"
            Response.Write "<TD class=CellLabel ID=loginID title=""User's Login ID"" >Login ID</TD>" & vbCrLf
            Response.Write "<TD class=CellLabel ID=name title=""User's Name"" >Name</TD>" & vbCrLf
            Response.Write "<TD class=CellLabel ID=login title=""User's Login Name"">Login</TD>" & vbCrLf
            Response.Write "</TR></THEAD>" & vbCrLf
            Response.Write "<TBODY ID=ListBody> " & vbCrLf
            
            intLine = 1
            If ReqForm("FormAction") = "Find" Or ReqForm("FormAction") = "Reset" Then
				Do While Not adRs.EOF
				    strRowStart = "<TR ID=ListRow" & intLine & " class=TableRow onclick=Result_onclick(" & intLine & ") ondblclick=Result_ondblclick(" & intLine & ")> " & vbCrLf
				    Response.Write strRowStart & "<TD ID=loginID" & intLine & " title=""User's Login ID"" class=TableDetail>" & vbCrLf
				    Response.Write adRs.Fields(0).Value & "</TD>" & vbCrLf
				    
				    Response.Write "<TD ID=name" & intLine & " title=""User's Name"" class=TableDetail >" & vbCrLf
				    Response.Write adRs.Fields("StaffName").Value & "</TD>" & vbCrLf
				    
				    Response.Write "<TD ID=login" & intLine & " title=""User's Login Name"" class=TableDetail >"
				    Response.Write adRs.Fields("logUserLogin").Value & "</TD>" & vbCrLf
				    
				    intLine = intLine + 1
				    adRs.MoveNext 
				Loop 
            End If
            Response.Write "</TBODY> </TABLE>"
            %>
        </DIV>
             
        <BUTTON id=cmdAdd class=DefBUTTON title="Add new staff record" 
            style="LEFT:10px; WIDTH:65px; TOP:225px; HEIGHT:20px"
            accessKey=A tabIndex=6 type=button>
            <U>A</U>dd
        </BUTTON>
        <BUTTON id=cmdEdit class=DefBUTTON title="Edit the selected staff record" 
            style="LEFT:80px; WIDTH:65px; TOP:225px; HEIGHT:20px"
            accessKey=E tabIndex=7 type=button>
            <U>E</U>dit
        </BUTTON>
        <BUTTON id=cmdDelete class=DefBUTTON title="Delete the selected staff record" 
            style="LEFT:150px; WIDTH:65px; TOP:225px; HEIGHT:20px"
            accessKey=D tabIndex=8 type=button>
            <U>D</U>elete
        </BUTTON>
        <BUTTON id=cmdReset class=DefBUTTON title="Reset the password for the selected user" 
            style="LEFT:-1220; WIDTH:65; TOP:225; HEIGHT:20; visibility:visible"
            accessKey=P tabIndex=8 type=button>
            <U>R</U>eset
        </BUTTON>
        <BUTTON id=cmdPrint class=DefBUTTON title="Print a summary of the employees in the search results" 
            style="LEFT:295px; WIDTH:65px; TOP:225px; HEIGHT:20px;display:none"
            accessKey=P tabIndex=8 type=button>
            <U>P</U>rint
        </BUTTON>
        <BUTTON id=cmdClose class=DefBUTTON
            style="LEFT:380px; WIDTH:65px; TOP:225px; HEIGHT:20px"
            tabIndex=9 type=button>Close
        </BUTTON>
    </DIV>

<FORM NAME="Form" METHOD="post" STYLE="VISIBILITY: hidden" ACTION="CaseEdit.ASP" ID=Form>
    <%
    Call CommonFormFields()
    WriteFormField "LoginID", ReqForm("LoginID")
    WriteFormField "StaffID", ReqForm("StaffID")
    WriteFormField "LastName", ReqForm("LastName")
    WriteFormField "FirstName", ReqForm("FirstName")
    WriteFormField "MiddleName", ReqForm("MiddleName")
    WriteFormField "logRecordID", ReqForm("logRecordID")
    WriteFormField "SelectedIndex", ""
    WriteFormField "FormAction", ""
    gadoCon.Close
    Set gadoCon = Nothing
    %>
</FORM>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncNavigateControls.asp"-->

</BODY></HTML>
