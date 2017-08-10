<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: Logon.asp                                                       '
'  Purpose: The logon form for the application that is used to provide a    '
'           psuedo-security for the Case Review System.  It prompts for a   '
'           user ID and password and passes the values on to the next form, '
'           the main menu screen.  Server side script on the main form does '
'           the validation of the user ID and password.                     '
'==========================================================================='
'If Request.ServerVariables("SERVER_PORT_SECURE") <> "1" Then
'    Response.Write "<BR><BR>&nbsp;&nbsp;This site requires a secure connection.  Click on the link below for the secure connection.<br><br>&nbsp;&nbsp;"
'    Response.Write "<a href=""https://secure.rushmore-group.com/MaderaCW/"">https://secure.rushmore-group.com/MaderaCW/</a>"
'    Response.End
'End If
%><!--#include file="IncCnn.asp"-->
<%
Dim mstrPageTitle

'Set the title of the page:
mstrPageTitle = Trim(gstrTitle & "<br>" & gstrAppName)
%>

<HTML>
<HEAD>
    <meta name=vs_targetSchema content="HTML 4.0">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>


<SCRIPT LANGUAGE=vbscript>
Option Explicit

Sub window_onload
	If "<%=Request.ServerVariables("SERVER_NAME")%>" = "secure.rushmore-group.com" Then
        Window.ResizeTo 437, 377
    Else
        Window.ResizeTo 437, 337
    End If
    Window.MoveTo 225, 225
    
    txtUserID.focus
End Sub

Sub OnKeyDown()
    If window.event.keyCode = 13 Then
        Call cmdOk_onclick
    ElseIf window.event.keyCode = 27 Then
        Call cmdCancel_onclick
    End If
End Sub

Sub cmdCancel_onclick()
    txtUserID.value = ""
    txtPassword.value = ""
    txtNewPassword.value = ""
    txtConfirmNewPassword.value = ""

    Form.UserID.value = ""
    Form.Password.value = ""
    Form.NewPassword.value = ""
    
    window.close   
End Sub

Sub cmdOk_onclick
    If Trim(txtUserID.value) = "" Then
        Exit Sub
    End If
    If Trim(txtPassword.value) = "" Then
        Exit Sub
    End If

    If Trim(txtNewPassword.value) <> "" Then
        If txtNewPassword.value <> txtConfirmNewPassword.value Then
            MsgBox "Please confirm your new password.", vbinformation, "Change Password"
            txtNewPassword.focus
            Exit Sub
        End If
        If Len(Trim(txtNewPassword.value)) < <%=GetAppSetting("MinPwLen")%> Or Len(Trim(txtNewPassword.value)) > 20 Then
            MsgBox "Passwords must be between <%=GetAppSetting("MinPwLen")%> and 20 characters in length.", vbInformation, "Change Password"
            txtNewPassword.focus
            Exit Sub
        End If
    End If
    
    Form.UserID.value = txtUserID.value
    Form.Password.value = txtPassword.value
    Form.Newpassword.value = txtNewPassword.value
    Form.CalledFrom.value = "Logon"
    Form.Submit
End Sub

</SCRIPT>

<%'----------------------------------------------------------------------------
'  Client side include files:
'----------------------------------------------------------------------------%>
<!--#include file="IncCmnCliFunctions.asp"-->

<%'----------------------------------------------------------------------------
'  Page HTML content:
'----------------------------------------------------------------------------%>
<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5>
    
    <DIV id=Header class=DefTitleArea style="WIDTH:408; HEIGHT: 55">

        <SPAN id=lblAppTitleHiLight class=DefTitleTextHiLight
            style="WIDTH:408"><%=mstrPageTitle%>
        </SPAN>
        
        <SPAN id=lblAppTitle class=DefTitleText
            style="WIDTH:408"><%=mstrPageTitle%>
        </SPAN>

    </DIV>
            
    <DIV id=PageFrame class=DefPageFrame style="HEIGHT: 225; WIDTH:408; TOP:66">

        <SPAN id=lblPleaseLogon class=DefLabel 
            style="FONT-SIZE:12; LEFT:8; font-weight:bold; COLOR:<%=gstrTitleColor%>; WIDTH:150; TOP:5">
            Please Log On:
        </SPAN>

        <SPAN id=lblUserID class=DefLabel 
            style="LEFT:55; WIDTH:125; TOP:35">
            Enter User ID:
        </SPAN>
        
        <INPUT id=txtUserID type=text TITLE="Enter your user login ID."
            style="LEFT:185; WIDTH:125; TEXT-ALIGN:left; TOP:35" 
            onkeydown="OnKeyDown" onfocus="CmnTxt_onfocus(txtUserID)"
            tabIndex=1 rows=1 cols=50> 
        
        <SPAN id=lblPassword class=DefLabel
            style="LEFT:55; WIDTH:125; TOP:60">
            Enter Password:
        </SPAN>

        <INPUT ID=txtPassword TYPE=password TITLE="Enter your Password." 
            style="LEFT:185; WIDTH:125; HEIGHT: 18; TOP:60; TEXT-ALIGN:LEFT" 
            onkeydown="OnKeyDown" onfocus="CmnTxt_onfocus(txtPassword)"
            tabIndex=2>

        <SPAN id=lblNewPassword class=DefLabel
            style="LEFT:55; WIDTH:125; TOP:95">
            Change Password To:
        </SPAN>

        <INPUT ID=txtNewPassword TYPE=password TITLE="Enter new Password."
            style="LEFT:185; WIDTH:125; HEIGHT: 18; TOP:95" 
            onkeydown="OnKeyDown" onfocus="CmnTxt_onfocus(txtNewPassword)"
            tabIndex=3>

        <SPAN id=lblConfirmNewPassword class=DefLabel
            style="LEFT:55; WIDTH:125; TOP:120px">
            Confirm New Password:
        </SPAN>

        <INPUT ID=txtConfirmNewPassword TYPE=password TITLE="Confirm new Password." 
            style="LEFT:185; WIDTH:125; TOP:120; HEIGHT: 18"
            onkeydown="OnKeyDown" onfocus="CmnTxt_onfocus(txtConfirmNewPassword)"
            tabIndex=4>

        <BUTTON id=cmdOk class=DefBUTTON TITLE="Log On" 
            style="LEFT:200; WIDTH:90; TOP:180"
            onkeydown="OnKeyDown"
            tabIndex=5 type=button>
            OK
        </BUTTON>

        <BUTTON id=cmdCancel class=DefBUTTON TITLE="Cancel Log On" 
            style="LEFT:300; WIDTH:90; TOP:180"
            onkeydown="OnKeyDown"
            tabIndex=6 type=button>Cancel
        </BUTTON>
    </DIV>
    
    <FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="Main.ASP" ID=Form>
        <%Call CommonFormFields()%>
        <INPUT TYPE="hidden" Name="NewPassword" VALUE="" ID=NewPassword>
    </FORM>
    
</BODY>
</HTML>

<%'----------------------------------------------------------------------------
'  Server side include files:
'----------------------------------------------------------------------------%>
<!--#include file="IncCmnFormFields.asp"-->
