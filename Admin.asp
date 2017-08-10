<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%> 
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: Admin.asp                                                       '
'  Purpose: The main menu or switchboard for admin options of application.  '
'==========================================================================='
Dim mstrPageTitle   'Sets the title at the top of the form.
Dim madoCmd         'ADO command object used for this page.
Dim mstrTmp         'Temporary string holder for building prompts, etc.
Dim mstrPrgSelected 'The programs last selected by the user.
Dim mintLeft
Dim mintTop
Dim strResizeScreen	'Holds value of Screen Resize flag.
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<%
mstrPageTitle = "System Admin Functions"

%>

<HTML>
<HEAD>
    <META name=vs_targetSchema content="HTML 4.0">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Dim mblnCloseClicked
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload
    Call CheckForValidUser()
    Call SizeAndCenterWindow(767, 520, False)
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>

    mblnCloseClicked = False

    Call DisplayButtons("All")
    
    PageFrame.style.visibility = "visible"
End Sub

Sub ButtonClick(cmdButton)
    Dim strPage
    Dim intResponse
    Dim strMessage
    Dim blnClose
    
    blnClose = False
    Select Case cmdButton.id
        Case "cmdSQLQueries"
            strPage = "SQLQueries.asp"
        Case "cmdAppSettings"
            strPage = "AppOptionSelect.asp"
        Case "cmdArchive"
            strPage = "ArchiveMenu.asp"
        Case "cmdClose"
            blnClose = True
    End Select

    mblnCloseClicked = True
    If Not blnClose Then
        mblnSetFocusToMain = False
        Form.Action = strPage
        Form.submit
    Else
        mblnSetFocusToMain = True
        Call window.opener.ManageWindows(6,"Close")
    End If
End Sub

Sub window_onbeforeunload
    If Not mblnCloseClicked Then
        Call ButtonClick(cmdClose)
    End if
    If mblnSetFocusToMain = True Then
        window.opener.focus
    End If
End Sub

Sub ButtonMouseOver(cmdButton)
    cmdButton.style.fontWeight = "bold"
End Sub

Sub ButtonMouseOut(cmdButton)
    cmdButton.style.fontWeight = "normal"
End Sub

Sub DisplayButtons(intOption)
    ' Display all buttons
	cmdAppSettings.style.display = "inline" '1:Configure Settings and Options
	cmdSQLQueries.style.display = "inline" '2:Execute SQL Queries
	'cmdArchive.style.display = "inline" '3:Batch Import Maintainence
End Sub
</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin="5" leftMargin="5" topMargin="5" rightMargin="5">
    
    <DIV id=Header class=DefTitleArea style="WIDTH: 737">
        <SPAN id=lblAppTitle class=DefTitleText
            style="WIDTH: 737">
            <%=mstrPageTitle%>
        </SPAN>
        
        <SPAN id=lblCurrentDate class=DefLabel
            style="LEFT:525; TOP:5; WIDTH:200; TEXT-ALIGN:right; COLOR:<%=gstrBorderColor%>">
            <%=FormatDateTime(Date, vbLongdate)%>
        </SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>
    <% Call WriteNavigateControls(-1,30,gstrBackColor) %>   
            
    <DIV id=PageFrame class=DefPageFrame style="HEIGHT: 425; WIDTH: 737; TOP: 51; visibility:hidden">

        <SPAN id=lblDatabaseStatus
            class=DefLabel
            STYLE="VISIBILITY:hidden; LEFT:20; WIDTH:200; TOP:20; TEXT-ALIGN:center">
            Accessing Database...
        </SPAN>

        <SPAN class=DefLabel 
            id=lblCaseReviewMenu
            style="FONT-SIZE:10pt; FONT-WEIGHT:bold; TOP:15; LEFT:25; WIDTH:200">
            System Administration Menu:
        </SPAN>

        <BUTTON id=cmdAppSettings class=DefBUTTON title="Configure Settings and Options" 
            onclick="ButtonClick(cmdAppSettings)" onmouseover="ButtonMouseOver(cmdAppSettings)" onmouseout="ButtonMouseOut(cmdAppSettings)"
            style="LEFT:35; TOP:50;display:none"
            accessKey=O tabIndex=1>
            Application <U>O</U>ptions
        </BUTTON>
        <BUTTON id=cmdSQLQueries class=DefBUTTON title="Execute SQL Queries" 
            onclick="ButtonClick(cmdSQLQueries)" onmouseover="ButtonMouseOver(cmdSQLQueries)" onmouseout="ButtonMouseOut(cmdSQLQueries)"
            style="LEFT:205; TOP:50;display:none"
            accessKey=Q tabIndex=1>
            S<U>Q</U>L Queries
        </BUTTON>
        <BUTTON id=cmdArchive class=DefBUTTON title="Archive old information." 
            onclick="ButtonClick(cmdArchive)" onmouseover="ButtonMouseOver(cmdArchive)" onmouseout="ButtonMouseOut(cmdArchive)"
            style="LEFT:375; TOP:50;display:none"
            accessKey=A tabIndex=1>
            <U>A</U>rchive
        </BUTTON>
        <BUTTON id=cmdClose class=DefBUTTON title="Close" 
            onclick="ButtonClick(cmdClose)" 
            style="LEFT:555; FONT-WEIGHT:bold; TOP:380"
            tabIndex=1>
            Close
        </BUTTON>
    </DIV>
</BODY>

<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="" ID=Form>
<%
    Call CommonFormFields()
    WriteFormField "FormAction", ""
%>
</FORM>
</HTML>
<%
gadoCon.Close
Set gadoCon = Nothing
%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncNavigateControls.asp"-->
<!--#include file="IncWriteFormField.asp"-->
