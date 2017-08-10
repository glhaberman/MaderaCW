<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>

<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->

<%
Dim intLine
Dim mstrPageTitle
Dim mstrVisible
Dim mlngTabIndex        'Keeps track of tabindex when building controls.
Dim adRs
Dim adCmd
Dim mintTabIndex
Dim intBottomRow
Dim intButtonHeight
Dim intColumn1
Dim intColumn2
Dim intColumn3
Dim intRow1
Dim intRow2
Dim intRow3
Dim intRow4
Dim intTextBoxWidth
Dim mlngWindowID
Dim mstrReReviewType, mlngReReviewTypeID

mlngReReviewTypeID = ReqForm("ReReviewTypeID")
If ReqForm("ReReviewTypeID") = 0 Then
    mstrReReviewType = gstrEvaluation
    mlngWindowID = 5
Else
    mstrReReviewType = "CAR "
    mlngWindowID = 8
End If

mstrPageTitle = "Find " & mstrReReviewType & " For Edit"
intLine = -1 'Used to determine the number of matching results.
intBottomRow = 384
intButtonHeight = 20
intColumn1 = 80
intColumn2 = 220
intColumn3 = 280
intRow1 = 5
intRow2 = 29
intRow3 = 53
intRow4 = 77
intTextBoxWidth = 150
%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>


<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mctlStaff
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>
Dim mdctPrograms
Dim mblnCloseClicked
Dim mblnMainClosed

Sub window_onload()
    mblnCloseClicked = True
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>

    Call SizeAndCenterWindow(767, 520, True)
    Call CheckForValidUser()
    If Trim(Form.UserID.Value) = "" Then Exit Sub

    If IsNumeric(Form.SelectedIndex.Value) Then
        If CLng(Form.SelectedIndex.Value) > 0 Then
            Call Result_onclick(1)
            cmdEdit.disabled = False
            cmdPrint.disabled = False
        End If
    Else
        cmdEdit.disabled = True
        cmdPrint.disabled = True
        cmdEditWR.disabled = True
    End If

    If txtReviewer.value = "" Then txtReviewer.value = "<All>"
    If txtEvaluater.value = "" Then txtEvaluater.value = "<All>"

    PageFrame.disabled = False
    PageBody.style.cursor = "default"
    txtReReviewID.focus
End Sub

<%'If timer detects that Main has been closed, this sub will be called.%>
Sub MainClosed()
    mblnMainClosed = True
    mblnCloseClicked = True
    mblnSetFocusToMain = False
    window.close
End Sub

<%'If Main has not been closed, set focus back to it.%>
Sub window_onbeforeunload()
    If mblnSetFocusToMain = True And mblnMainClosed = False Then
        window.opener.focus
    End If
End Sub

Sub cmdCancel_onclick()
    'Return cancel:
    Call window.opener.ManageWindows(<%=mlngWindowID%>,"Close")
End Sub

Sub txtReviewID_onchange()
    txtReReviewDate.Value = ""
    txtCaseNumber.Value = ""
End Sub

Sub txtReReviewDate_onkeypress
    If txtReReviewDate.value = "(MM/DD/YYYY)" Then
        txtReReviewDate.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub
Sub txtReReviewDate_onblur
    If Trim(txtReReviewDate.value) = "(MM/DD/YYYY)" Then
        txtReReviewDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtReReviewDate.value) Then
        MsgBox "The Start Date must be a valid date - MM/DD/YYYY.", vbInformation, "Find <%=mstrReReviewType%>"
        txtReReviewDate.focus
    ElseIf IsDate(txtReReviewDate.value) Then
        If CDate(txtReReviewDate.value) < CDate("01/01/1900") Then
            MsgBox "The Start Date must be a valid date - MM/DD/YYYY." & vbCrLf & "Dates prior to 01/01/1900 are not considered valid for this entry.", vbInformation, "Find <%=mstrReReviewType%>"
            txtReReviewDate.focus
        End If
    ElseIf IsDate(txtReReviewDateEnd.value) And IsDate(txtReReviewDate.value) Then
        If CDate(txtReReviewDateEnd.value) < CDate(txtReReviewDate.value) Then
            MsgBox "The Start Date must be before the end date." & vbCrLf, vbInformation, "Find <%=mstrReReviewType%>"
            txtReReviewDate.focus
        End If
    End If
End Sub
Sub txtReReviewDate_onfocus
    If Trim(txtReReviewDate.value) = "" Then
        txtReReviewDate.value = "(MM/DD/YYYY)"
    End If
    txtReReviewDate.select
End Sub

Sub txtReReviewDate_onchange()
    txtReviewID.Value = ""
    txtCaseNumber.Value = ""
End Sub

Sub txtReReviewDateEnd_onkeypress
    If txtReReviewDateEnd.value = "(MM/DD/YYYY)" Then
        txtReReviewDateEnd.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub
Sub txtReReviewDateEnd_onblur
    If Trim(txtReReviewDateEnd.value) = "(MM/DD/YYYY)" Then
        txtReReviewDateEnd.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtReReviewDateEnd.value) Then
        MsgBox "The End Date must be a valid date - MM/DD/YYYY.", vbInformation, "Find <%=mstrReReviewType%>"
        txtReReviewDateEnd.focus
    ElseIf IsDate(txtReReviewDateEnd.value) Then
        If CDate(txtReReviewDateEnd.value) < CDate("01/01/1900") Then
            MsgBox "The End Date must be a valid date - MM/DD/YYYY." & vbCrLf & "Dates prior to 01/01/1900 are not considered valid for this entry.", vbInformation, "Find <%=mstrReReviewType%>"
            txtReReviewDateEnd.focus
        End If
    ElseIf IsDate(txtReReviewDateEnd.value) And IsDate(txtReReviewDate.value) Then
        If CDate(txtReReviewDateEnd.value) < CDate(txtReReviewDate.value) Then
            MsgBox "The End Date must be after the start date." & vbCrLf, vbInformation, "Find <%=mstrReReviewType%>"
            txtReReviewDateEnd.focus
        End If
    End If
End Sub
Sub txtReReviewDateEnd_onfocus
    If Trim(txtReReviewDateEnd.value) = "" Then
        txtReReviewDateEnd.value = "(MM/DD/YYYY)"
    End If
    txtReReviewDateEnd.select
End Sub

Sub txtCaseNumber_onchange()
    txtReReviewID.Value = ""
    txtReviewID.Value = ""
    txtReReviewDate.Value = ""
End Sub

Sub cmdFind_onclick()
    Dim blnCriteria
    Dim intResp
    Dim strParms
    Dim intI

    txtReReviewID.Value = Trim(txtReReviewID.Value)
    txtReviewID.Value = Trim(txtReviewID.Value)
    txtReReviewDate.Value = Trim(txtReReviewDate.Value)
    txtCaseNumber.Value = Trim(txtCaseNumber.Value)
    
    blnCriteria = False

    If txtReReviewID.Value <> "" Then
        If Not IsNumeric(txtReReviewID.Value) Then
            MsgBox "The <%=mstrReReviewType%> ID must be a number."               
            txtReReviewID.focus
            Exit Sub
        End If
        blnCriteria = True
    ElseIf txtReviewID.Value <> "" Then
        If Not IsNumeric(txtReviewID.Value) Then
            MsgBox "The Case Review ID must be a number."               
            txtReviewID.focus
            Exit Sub
        End If
        blnCriteria = True
    ElseIf txtReReviewDate.Value <> "" Then
        If Not IsDate(txtReReviewDate.Value) Then
            MsgBox "The Starting <%=mstrReReviewType%> Date must be a valid date."
            txtReReviewDate.focus
            Exit Sub
        End If
        blnCriteria = True
    ElseIf txtReReviewDateEnd.Value <> "" Then
        If Not IsDate(txtReReviewDateEnd.Value) Then
            MsgBox "The Ending <%=mstrReReviewType%> Date must be a valid date."
            txtReReviewDateEnd.focus
            Exit Sub
        End If
        blnCriteria = True
    ElseIf txtCaseNumber.Value <> "" Then
        If Len(txtCaseNumber.Value) > <%=GetAppSetting("MaxCaseNumberLength")%> Then
            MsgBox "The Case Number cannot be longer than <%=GetAppSetting("MaxCaseNumberLength")%> digits."
            txtCaseNumber.Value
            Exit Sub
        End If
        blnCriteria = True
    End If
    If Trim(txtEvaluater.Value) <> "<All>" Then
        blnCriteria = True
    End If    
    If Trim(txtReviewer.Value) <> "<All>" Then
        blnCriteria = True
    End If    

    If Not blnCriteria Then
        intResp = MsgBox("This may return a large number of search results.  " & vbcrlf & vbcrlf & "Do you wish to continue?",vbYesNo + vbQuestion,"Find Matching Case <%=mstrReReviewType%>s")
        If intResp = vbNo Then
            PageFrame.disabled = False
            PageFrame.style.visibility = "visible"
            Exit Sub
        End If
    End If

    Form.casID.Value = txtReviewID.Value
    Form.ReReviewDate.Value = txtReReviewDate.Value
    Form.ReReviewDateEnd.Value = txtReReviewDateEnd.Value
    Form.CaseNumber.Value = txtCaseNumber.Value
    Form.ReReviewID.Value = txtReReviewID.Value
    Form.ReReviewer.Value = txtEvaluater.Value
    Form.Reviewer.value = txtReviewer.value
    
    PageBody.style.cursor = "wait"

    strParms = "<%=gstrUserID%>^<%=gblnUserAdmin%>^<%=gblnUserQA%>"
    strParms = strParms & "^" & Form.ReReviewID.value
    strParms = strParms & "^" & Form.casID.value
    strParms = strParms & "^" & Form.ReReviewDate.value
    strParms = strParms & "^" & Form.ReReviewDateEnd.value
    strParms = strParms & "^" & Form.CaseNumber.value
    strParms = strParms & "^" & ReplaceAllWithBlank(Form.ReReviewer.value)
    strParms = strParms & "^" & Form.SortOrder.value
    strParms = strParms & "^" & ReplaceAllWithBlank(Form.Reviewer.value)
    strParms = strParms & "^<%=glngAliasPosID%>"
    
    ' Load search in IFRAME
    fraResults.frameElement.src = "FindReReviewResults.asp?Load=Y&ReReviewTypeID=<%=ReqForm("ReReviewTypeID")%>&ParmList=" & strParms
End Sub

Function ReplaceAllWithBlank(strValue)
    If strValue = "<All>" Then
        ReplaceAllWithBlank = ""
    Else
        ReplaceAllWithBlank = strValue
    End If
End Function

Sub cmdEdit_onclick()
    Call EditRecord()
End Sub

Sub EditRecord()
    If IsNull(Form.SelectedIndex.Value) Or Trim(Form.SelectedIndex.Value) = "" Then
        Exit Sub
    End If
    
    window.opener.Form.ReReviewID.Value = Form.ReReviewID.Value
    window.opener.Form.Action = "ReReviewAddEdit.asp"
    
    Call SizeAndCenterWindow(767, 520, True)

    window.opener.Form.FormAction.Value = "GetRecord"
    window.opener.Form.ReReviewTypeID.value = <%=ReqForm("ReReviewTypeID")%>
    Call window.opener.ManageWindows(<%=mlngWindowID-1%>,"EditReReview")
End Sub

Sub cmdPrint_onclick()
    Dim strReturnValue
    
    cmdPrint.disabled = True
    <%'Open the print-preview window, passing it the review ID:%>
    strReturnValue = window.showModalDialog("PrintReReview.asp?UserID=<%=gstrUserID%>&AuditRead=True&ReReviewID=" & Form.ReReviewID.Value & _
        "&PageTitle=<%=mstrPageTitle%>", , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
    cmdPrint.disabled = False
End Sub

Sub cmdClear_onclick()
    txtReReviewID.Value = ""
    txtReviewID.Value = ""
    txtReReviewDate.Value = ""
    txtReReviewDateEnd.Value = ""
    txtCaseNumber.Value = ""
    txtEvaluater.Value = "<All>"
    txtReviewer.Value = "<All>"
End Sub

Sub Gen_onkeydown
    If window.event.keyCode = 13 Then
        Call cmdFind_onclick
    ElseIf window.event.keyCode = 27 Then
        Call cmdCancel_onclick
    End If
End Sub

Sub Gen_onkeydown2(ctlFrom)
    If window.event.keyCode = 13 Then
        Call StaffLookUp(ctlFrom)
    End If
End Sub

Sub Gen_onfocus(txtBox)
    txtBox.select
End Sub

</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style="cursor:wait">
    
    <DIV id=Header class=DefTitleArea style="width:737;height:30">
		<SPAN id=lblAppTitle class=DefTitleText style="LEFT:100;WIDTH:290;text-align:right">
            <%=mstrPageTitle%>
        </SPAN>
		<SPAN id=lblAppTitle2 class=DefTitleText style="top:6;font-size:12;LEFT:400;WIDTH:200;text-align:left">
            ~ Enter Search Criteria
        </SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>

    <% Call WriteNavigateControls(5,30,gstrBackColor) %>
    
    <DIV id=PageFrame class=DefPageFrame disabled=true style="BORDER-TOP-STYLE:none;WIDTH:737; HEIGHT:410; TOP:40">
        <!-- Column 1 -->
        <SPAN id=lblReReviewID class=DefLabel style="LEFT:1; WIDTH:<%=intColumn1-5%>;TOP:<%=intRow1%>;text-align:right">
            <%=mstrReReviewType%> ID
        </SPAN>
        <INPUT id=txtReReviewID title="<%=mstrReReviewType%> ID"
            style="LEFT:<%=intColumn1%>;WIDTH:80;TOP:<%=intRow1%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReReviewID)" NAME="txtReReviewID">

        <SPAN id=lblReviewID class=DefLabel style="LEFT:1; WIDTH:<%=intColumn1-5%>;TOP:<%=intRow2%>;text-align:right">
            Review ID
        </SPAN>
        <INPUT id=txtReviewID title="Review ID"
            style="LEFT:<%=intColumn1%>;WIDTH:80;TOP:<%=intRow2%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReviewID)" NAME="txtReviewID">

        <SPAN id=lblCaseNumber class=DefLabel style="LEFT:1;WIDTH:<%=intColumn1-5%>;TOP:<%=intRow3%>;text-align:right">
            Case Number
        </SPAN>
        <INPUT id=txtCaseNumber title="Case Number"
            style="LEFT:<%=intColumn1%>; WIDTH:80;TOP:<%=intRow3%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtCaseNumber)" NAME="txtCaseNumber">

        <!-- Column 2 -->
        <SPAN id=lblRvwr class=DefLabel style="LEFT:160; WIDTH:100;TOP:<%=intRow1%>;text-align:right">
            <%=gstrEvaTitle%>
        </SPAN>
        <INPUT id=txtEvaluater 
            style="LEFT:265; WIDTH:<%=intTextBoxWidth+50%>;TOP:<%=intRow1%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtEvaluater)" NAME="txtEvaluater">

        <SPAN id=lblReviewer class=DefLabel style="LEFT:160; WIDTH:100;TOP:<%=intRow2%>;text-align:right">
            <%=gstrRvwTitle%>
        </SPAN>
        <INPUT id=txtReviewer 
            style="LEFT:265; WIDTH:<%=intTextBoxWidth+50%>;TOP:<%=intRow2%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReviewer)" NAME="txtReviewer">
        <!-- Column 3 -->
        <SPAN id=lblReReviewDate class=DefLabel style="LEFT:510; WIDTH:100; TOP:<%=intRow1%>">
            <%=mstrReReviewType%> Dates
        </SPAN>
        <SPAN id=lblReReviewDateStart class=DefLabel style="LEFT:510; WIDTH:100; TOP:22">
            From
        </SPAN>
        <INPUT id=txtReReviewDate title="Beginning <%=mstrReReviewType%> Date" tabindex=<%=GetTabIndex%>
            style="LEFT:540; WIDTH:80; TOP:22" maxlength=10
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReReviewDate)" NAME="txtReReviewDate">
        <SPAN id=lblReReviewDateEnd class=DefLabel style="LEFT:520; WIDTH:25; TOP:45">
            To
        </SPAN>
        <INPUT id=txtReReviewDateEnd title="Ending <%=mstrReReviewType%> Date" tabindex=<%=GetTabIndex%>
            style="LEFT:540; WIDTH:80; TOP:45"  maxlength=10
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReReviewDateEnd)" NAME="txtReReviewDateEnd">

        <BUTTON id=cmdFind class=DefBUTTON title="Search for matching record(s)" 
            style="LEFT:640;TOP:<%=intRow1%>;HEIGHT:<%=intButtonHeight%>;WIDTH:65" tabindex=<%=GetTabIndex%>
            accessKey=F>
            <U>F</U>ind
        </BUTTON>

        <BUTTON id=cmdClear class=DefBUTTON title="Clear all search criteria" 
            style="LEFT:640;TOP:<%=intRow2%>;HEIGHT:<%=intButtonHeight%>;WIDTH:65" tabindex=<%=GetTabIndex%>
            accessKey=C>
            <U>C</U>lear
        </BUTTON>
        
        <DIV id=lstResults class=DefPageFrame style="LEFT:0;WIDTH:736;border-left-style:none;border-right-style:solid; HEIGHT:305; TOP:75">
            <IFRAME ID=fraResults src="FindReReviewResults.asp?Load=N&ReReviewTypeID=<%=ReqForm("ReReviewTypeID")%>"
                STYLE="positon:absolute; LEFT:0; WIDTH:735; HEIGHT:303; TOP:0; BORDER-style:none" FRAMEBORDER=0>
            </IFRAME>
        </DIV>

        <BUTTON id=cmdEdit class=DefBUTTON title="Edit the selected record" 
            style="LEFT:15; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>;WIDTH:100" tabindex=<%=GetTabIndex%>
            accessKey=E>
            <U>E</U>dit <%=mstrReReviewType%>
        </BUTTON>
        <BUTTON id=cmdPrint class=DefBUTTON title="Print the selected record" 
            style="LEFT:120; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>;WIDTH:100" tabindex=<%=GetTabIndex%>
            accessKey=P>
            <U>P</U>rint <%=mstrReReviewType%>
        </BUTTON>
        <BUTTON id=cmdEditWR class=DefBUTTON title="Submit to Reports" 
            style="LEFT:-1225; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>;WIDTH:100" tabindex=<%=GetTabIndex%>
            accessKey=S>
            <U>S</U>ubmit To Reports
        </BUTTON>

        <SPAN id=lblStatus class=DefLabel style="LEFT:285; WIDTH:370; TOP:<%=intBottomRow%>; text-align:center">
            Enter search criteria and click [Find].
        </SPAN>

        <BUTTON id=cmdCancel class=DefBUTTON title="Close and return to previous" 
            style="LEFT:640; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>; WIDTH:85" tabindex=<%=GetTabIndex%>>Close
        </BUTTON>
    </DIV>

    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY:hidden"" ACTION=""CaseEdit.ASP"" ID=Form>" & vbCrLf

    Call CommonFormFields()

    If ReqForm("FormAction") = "Y" Then
        WriteFormField "casID", ReqForm("casID")
        WriteFormField "ReReviewID", ReqForm("ReReviewID")
        WriteFormField "ReReviewDate", ReqForm("ReReviewDate")
        WriteFormField "ReReviewDateEnd", ReqForm("ReReviewDateEnd")
        WriteFormField "CaseNumber", ReqForm("CaseNumber")
        WriteFormField "ReReviewer", ReqForm("ReReviewer")
        WriteFormField "Reviewer", ReqForm("ReReviewer")
    Else
        WriteFormField "casID", ""
        WriteFormField "ReReviewID", ""
        WriteFormField "ReReviewDate", ""
        WriteFormField "ReReviewDateEnd", ""
        WriteFormField "CaseNumber", ""
        WriteFormField "ReReviewer", ""
        WriteFormField "Reviewer", ""
    End If
    WriteFormField "FormAction", ""
    If intLine > 0 Then
        WriteFormField "SelectedIndex", "1"
    Else
        WriteFormField "SelectedIndex", ""
    End if
    WriteFormField "ResultsCount", intLine - 1
    WriteFormField "StaffInformation", ""
    WriteFormField "SortOrder", ""
    Response.Write Space(4) & "</FORM>"

    gadoCon.Close
    Set gadoCon = Nothing
    %>
</BODY>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncBuildList.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<!--#include file="IncNavigateControls.asp"-->
