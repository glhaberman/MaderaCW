<%@ LANGUAGE="VBScript" %>
<%Option Explicit%> 
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
Dim mstrPageTitle
Dim adCmd, adRs
Dim mlngTabIndex
Dim mdtmLastArchive, mstrLastUserID, mstrLastExDate

Set adRs = Server.CreateObject("ADODB.Recordset")
Set adCmd = GetAdoCmd("spArchiveGetLast")
    adRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
Set adCmd = Nothing
If adRs.RecordCount = 1 Then
    mdtmLastArchive = adRs.Fields("audDateOfAction").Value
    mstrLastUserID = adRs.Fields("audUserLogin").Value
    mstrLastExDate = adRs.Fields("audActionComments").Value
Else
    mdtmLastArchive = "Never"
    mstrLastUserID = ""
    mstrLastExDate = ""
End If
mstrPageTitle = "Move Reviews To Archive"

%>
<HTML>
<HEAD>
    <META name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE>
        <%=Trim(gstrOrgAbbr & " " & gstrAppName)%>
    </TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>
<SCRIPT LANGUAGE="vbscript">
Dim mblnCloseClicked
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload
    Call CheckForValidUser()
    Call SizeAndCenterWindow(767, 520, False)
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>

    If "<%=mdtmLastArchive%>" = "Never" Then
        'Default date of 1/1/2000 indicates archive has never ran.
        lblLastArchive.style.left = -1000
    End If
    mblnCloseClicked = False

    PageFrame.style.visibility = "visible"
End Sub

<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    mblnSetFocusToMain = False
    mblnCloseClicked = True
    window.close
End Sub

Sub cmdClose_onclick
    Form.Action = "ArchiveMenu.asp"
    Form.Submit
End Sub

Sub Gen_onkeydown
    If window.event.keyCode = 13 Then
        Call cmdFind_onclick
    ElseIf window.event.keyCode = 27 Then
        Call cmdCancel_onclick
    End If
End Sub

Sub Gen_onfocus(txtBox)
    txtBox.select
End Sub

Sub cmdBeginArchive_onclick()
    If Not IsDate(txtArchiveDate.value) Then
        MsgBox "Please enter the archive date.", vbInformation, "Begin Archive"
        txtArchiveDate.focus
        Exit Sub
    End If
    lblExecuting.innerText = "Please wait while reviews are archived.  Do not close this form."
    PageFrame.style.cursor = "wait"
    fraExecute.PageBody.style.cursor = "wait"
    fraExecute.divMessage.innerText = " Executing..."
    fraExecute.frameElement.src = "ArchiveExecute.asp?UserID=" & Form.UserID.Value & "&FormAction=Execute&ArchiveDate=" & txtArchiveDate.Value
    lblExecuting.style.visibility = "visible"
    divExecute.style.left = 50
    PageFrame.disabled = True
End Sub
Sub Date_onkeypress(ctlDate)
    If ctlDate.value = "(MM/DD/YYYY)" Then
        ctlDate.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub

Sub Date_onblur(ctlDate)
    Dim intRowID
    
    If Trim(ctlDate.value) = "(MM/DD/YYYY)" Or Trim(ctlDate.value) = "" Then
        ctlDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(ctlDate.value) Then
        MsgBox "The End Date must be a valid date - MM/DD/YYYY.", vbInformation, "Factor Maintenance"
        ctlDate.focus
        Exit Sub
    End If
    
    'If CDate(ctlDate.value) <= CDate("<%=mdtmLastArchive%>") Then
End Sub

Sub Date_onfocus(ctlDate)
    If Trim(ctlDate.value) = "" Then
        ctlDate.value = "(MM/DD/YYYY)"
    End If
    ctlDate.select
End Sub
</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->
<BODY id="PageBody" bottomMargin="5" topMargin="5" leftMargin="5" rightMargin="5">
    <DIV id=FormTitle class=DefTitleArea style="WIDTH:737">
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
    <% Call WriteNavigateControls(-1,0,gstrAltBackColor) %>

    <DIV id=PageFrame class=DefPageFrame style="HEIGHT:425; WIDTH:737; TOP:51">

        <SPAN id=lblReviewDateStart class=DefLabel style="LEFT:20; WIDTH:370; TOP:25">
            <b>Enter the archive date:</b>&nbsp&nbsp&nbsp Any reviews with a Review Date (date entered in the system) on or before the archive date will be moved into the archive.
        </SPAN>
        <INPUT id=txtArchiveDate title="Beginning Review Date" tabindex=<%=GetTabIndex%>
            style="LEFT:400; WIDTH:80; TOP:30" maxlength=10
            onkeydown="Gen_onkeydown" onblur=Date_onblur(txtArchiveDate) onkeypress=Date_onkeypress(txtArchiveDate) onfocus=Date_onfocus(txtArchiveDate) NAME="txtArchiveDate">

        <BUTTON id=cmdBeginArchive class=DefBUTTON style="LEFT:500; TOP:27" tabIndex=1>
            Begin Archive
        </BUTTON>

        <SPAN id=lblLastArchive class=DefLabel style="LEFT:20; WIDTH:470; TOP:60">
            (Last Archive Date processed was <B><%=Replace(mstrLastExDate,"Archive Date = ","")%></B> on <B><%=mdtmLastArchive%></B> by <B><%=mstrLastUserID%></B>.)

        </SPAN>
        <SPAN id=SPAN1 class=DefLabel style="LEFT:20; WIDTH:370; TOP:80">
            <b>Note:</b>&nbsp&nbsp&nbsp This action is permanent.  Once reviews are moved into the archive, they can NOT be changed.
        </SPAN>

        <BUTTON id=cmdClose class=DefBUTTON style="LEFT:555; FONT-WEIGHT:bold; TOP:380" tabIndex=1>
            Close
        </BUTTON>
        
        <SPAN id=lblExecuting style="left:30;width:400;top:125;visibility:hidden">Please wait while reviews are archived.  Do not close this form.</SPAN>
        <DIV id=divExecute class=ControlDiv style="width:600;height:150;left:-1000;top:150;z-index:101;position:absolute">
            <IFRAME id=fraExecute src="ArchiveExecute.asp" style="width:600;height:148;left:1;top:1"></IFRAME>
        </DIV>
    </DIV>
    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY: hidden"" ACTION=""Main.asp"" ID=Form>" & vbCrLf
        Call CommonFormFields()
    	WriteFormField "FormAction", ReqForm("FormAction")
    	WriteFormField "ArchiveDate", ReqForm("ArchiveDate")
    Response.Write Space(4) & "</FORM>" & vbCrLf
    Set gadoCmd = Nothing
    gadoCon.Close
    Set gadoCon = Nothing
    %>
    <BR>
    <BR>
</BODY>
</HTML>
<%
%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncNavigateControls.asp"-->
<!--#include file="IncBuildList.asp"-->
