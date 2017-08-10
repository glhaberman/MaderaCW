<%@ LANGUAGE="VBScript" %>
<%Option Explicit%> 
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
Dim mstrPageTitle
Dim adRs
Dim adCmd
Dim mintTabIndex
Dim mlngTabIndex        'Keeps track of tabindex when building controls.
Dim intBottomRow
Dim intButtonHeight
Dim intColumn1
Dim intColumn2
Dim intColumn3
Dim intRow1, intRow2, intRow3, intRow4, intI
Dim intTextBoxWidth
'Dim mdctOffices, mdctManagers
Dim mstrOffices, mstrManagers

mstrPageTitle = "Case Review Archive"

intBottomRow = 384
intButtonHeight = 20
intColumn1 = 75
intColumn2 = 170
intColumn3 = 320
intRow1 = 5
intRow2 = 29
intRow3 = 53
intRow4 = 77
intTextBoxWidth = 120

Set adRs = Server.CreateObject("ADODB.Recordset")
' Staffing
Set adCmd = GetAdoCmd("spArchiveGetStaffing")
    AddParmIn adCmd, "@RoleName", adVarchar, 50, "Manager"
    adRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
Set adCmd = Nothing
mstrManagers = "<OPTION VALUE=0 SELECTED>&ltAll&gt"
intI = 0
Do While Not adRs.EOF
    'mdctManagers.Add adRs.Fields("StaffName").value, adRs.Fields("StaffName").value
    mstrManagers = mstrManagers & "<OPTION VALUE=" & intI & ">" & adRs.Fields("StaffName").value
    intI = intI + 1
    adRs.MoveNext
Loop
adRs.Close

Set adRs = Server.CreateObject("ADODB.Recordset")
' Staffing
Set adCmd = GetAdoCmd("spArchiveGetStaffing")
    AddParmIn adCmd, "@RoleName", adVarchar, 50, "Office"
    adRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
Set adCmd = Nothing
mstrOffices = "<OPTION VALUE=0 SELECTED>&ltAll&gt"
intI = 0
Do While Not adRs.EOF
    mstrOffices = mstrOffices & "<OPTION VALUE=" & intI & ">" & adRs.Fields("StaffName").value
    intI = intI + 1
    adRs.MoveNext
Loop
%>
<HTML>
<HEAD>
    <META name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE>
        <%=Trim(gstrOrgAbbr & " " & gstrAppName)%>
    </TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
<SCRIPT LANGUAGE="vbscript">
Dim mblnCloseClicked
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload
    Call CheckForValidUser()
    Call SizeAndCenterWindow(767, 520, False)
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>

    mblnCloseClicked = False
    If <%=gblnUserAdmin%> Or <%=gblnUserQA%> Then
        cmdAddToArchive.style.display = "inline"
    Else
        cmdAddToArchive.style.display = "none"
    End If
    txtReviewer.value = "<All>"
    txtSupervisor.value = "<All>"
    txtWorker.value = "<All>"
    PageFrame.style.visibility = "visible"
End Sub

<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    mblnSetFocusToMain = True
    mblnCloseClicked = True
    window.close
End Sub

Sub cmdClose_onclick
    mblnSetFocusToMain = True
    Call window.opener.ManageWindows(6,"Close")
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

Sub Gen_onkeydown2(ctlFrom)
    If window.event.keyCode = 13 Then
        Call cmdFind_onclick
    End If
End Sub

Sub cboResponse_onblur()
    cboResponse.style.width = <%=intColumn2-intColumn1%>
End Sub
Sub cboResponse_ondblclick()
    If cboResponse.style.pixelWidth = <%=intColumn2-intColumn1%> Then
        cboResponse.style.width = <%=intColumn2-intColumn1 + 100%>
    Else
        cboResponse.style.width = <%=intColumn2-intColumn1%>
    End If
End Sub
Sub cboResponse_onchange()
    cboResponse.style.width = <%=intColumn2-intColumn1%>
End Sub
Sub lblResponse_onmouseover()
    lblResponse.style.fontWeight = "bold"
End Sub
Sub lblResponse_onmouseout()
    lblResponse.style.fontWeight = "normal"
End Sub
Sub lblResponse_ondblClick()
    Call cboResponse_ondblclick()
End Sub
Sub lblResponse_onClick()
    Call cboResponse_ondblClick()
End Sub

Sub cmdReports_onclick()
    Form.Action = "ArchiveReports.asp"
    Form.Submit
End Sub

Sub cmdAddToArchive_onclick()
    Form.Action = "Archive.asp"
    Form.Submit
End Sub

Sub cmdClear_onclick()
    txtReviewID.Value = ""
    txtReviewDate.Value = ""
    txtReviewDateEnd.Value = ""
    txtCaseNumber.Value = ""
    txtWorker.value = ""
    cboResponse.value = 0
    txtReviewer.value = "<All>"
    txtSupervisor.value = "<All>"
    txtWorker.value = "<All>"
    cboManager.value = "0"
    cboOffice.value = "0"
End Sub

Sub txtReviewer_onblur()
    If Trim(txtReviewer.value) = "" Then
        txtReviewer.value = "<All>"
    End If
End Sub

Sub txtSupervisor_onblur()
    If Trim(txtSupervisor.value) = "" Then
        txtSupervisor.value = "<All>"
    End If
End Sub

Sub txtWorker_onblur()
    If Trim(txtWorker.value) = "" Then
        txtWorker.value = "<All>"
    End If
End Sub

Sub cmdFind_onclick()
    Dim blnCriteria
    Dim intResp
    Dim strParms
    Dim intI

    Call RebuildProgramsSelected()    

    txtReviewID.Value = Trim(txtReviewID.Value)
    txtReviewDate.Value = Trim(txtReviewDate.Value)
    txtReviewDateEnd.Value = Trim(txtReviewDateEnd.Value)
    txtCaseNumber.Value = Trim(txtCaseNumber.Value)
    
    blnCriteria = False

    If txtReviewID.Value <> "" Then
        If Not IsNumeric(txtReviewID.Value) Then
            MsgBox "The Case Review ID must be a number."               
            txtReviewID.focus
            Exit Sub
        End If
        blnCriteria = True
    ElseIf txtReviewDate.Value <> "" Then
        If Not IsDate(txtReviewDate.Value) Then
            MsgBox "The Starting Review Date must be a valid date."
            txtReviewDate.focus
            Exit Sub
        End If
        blnCriteria = True
    ElseIf txtReviewDateEnd.Value <> "" Then
        If Not IsDate(txtReviewDateEnd.Value) Then
            MsgBox "The Ending Review Date must be a valid date."
            txtReviewDateEnd.focus
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
    
    If Trim(cboManager.Value) <> "0" Then
        blnCriteria = True
    End If
    If Trim(cboOffice.Value) <> "0" Then
        blnCriteria = True
    End If
    If Trim(txtSupervisor.Value) <> "<All>" Then
        blnCriteria = True
    End If
    If Trim(txtWorker.Value) <> "<All>" Then
        blnCriteria = True
    End If
    If Trim(cboResponse.Value) <> "0" Then
        blnCriteria = True
    End If
    If Trim(txtReviewer.Value) <> "<All>" Then
        blnCriteria = True
    End If

    If Not blnCriteria And Form.UseWarning.value = "Yes" Then
        intResp = MsgBox("This may return a large number of search results.  " & vbcrlf & vbcrlf & "Do you wish to continue?",vbYesNo + vbQuestion,"Find Matching Case Reviews")
        If intResp = vbNo Then
            PageFrame.disabled = False
            PageFrame.style.visibility = "visible"
            Exit Sub
        End If
    End If
    
    Call DisablePage(True)
    
    Form.rvwID.Value = txtReviewID.Value
    Form.ReviewDate.Value = txtReviewDate.Value
    Form.ReviewDateEnd.Value = txtReviewDateEnd.Value
    Form.CaseNumber.Value = txtCaseNumber.Value
    Form.WorkerName.value = txtWorker.value
    Form.Response.Value = GetComboText(cboResponse)
    Form.Reviewer.Value = txtReviewer.Value
    Form.Supervisor.Value = txtSupervisor.value
    Form.Manager.Value = GetComboText(cboManager)
    Form.Office.Value = GetComboText(cboOffice)
    Form.Director.Value = "0"
    Form.ReviewClass.Value = GetComboText(cboReviewClass)
    Form.UseWarning.Value = "Yes"

    strParms = Form.rvwID.value
    strParms = strParms & "^" & Form.ReviewDate.value
    strParms = strParms & "^" & Form.ReviewDateEnd.value
    strParms = strParms & "^" & Form.CaseNumber.value
    strParms = strParms & "^" & Form.WorkerName.value
    strParms = strParms & "^" & Form.Supervisor.value
    strParms = strParms & "^" & Form.Manager.value
    strParms = strParms & "^" & Form.Office.value
    strParms = strParms & "^" & Form.Director.value
    strParms = strParms & "^" & Form.Response.value
    strParms = strParms & "^" & Form.Reviewer.value
    strParms = strParms & "^" & Form.ProgramsSelected.value
    strParms = strParms & "^" & Form.ReviewClass.value
    strParms = strParms & "^" & Form.SortOrder.value
    
    ' Load search in IFRAME
    fraResults.frameElement.src = "ArcFindResults.asp?Load=Y&ParmList=" & strParms
End Sub

Sub DisablePage(blnVal)
    Dim strCursor

    If blnVal Then
        strCursor = "wait" 
    Else
        strCursor = "default"
    End If

    PageBody.disabled = blnVal
    PageBody.style.cursor = strCursor
    cmdFind.disabled = blnVal
    cmdClear.disabled = blnVal
    lstResults.disabled = blnVal
    lstResults.style.cursor = strCursor
    On Error Resume Next
    fraResults.lstCases.disabled = blnVal
    fraResults.lstCases.style.cursor = strCursor
    On Error Goto 0
    
    cmdAddToArchive.disabled = blnVal
    cmdPrintReview.disabled = blnVal
    cmdReports.disabled = blnVal
End Sub

Sub lblProgram_onClick(intWhich)
    If document.all("chkProgram" & intWhich).disabled = True Then Exit Sub
    document.all("chkProgram" & intWhich).checked = Not document.all("chkProgram" & intWhich).checked
End Sub

Sub RebuildProgramsSelected()
    Dim intCnt
    
    Form.ProgramsSelected.value = ""
    If Not IsNumeric(lblProgramCount.innerText) Then
        Exit Sub
    End If
    
    For intCnt = 1 To lblProgramCount.innerText
        If document.all("chkProgram" & intCnt).checked Then
            Form.ProgramsSelected.Value = Form.ProgramsSelected.Value & "[" & document.all("hidPrgInfo" & intCnt).value & "]" & "|"
        End If
    Next
End Sub

Sub cmdPrintReview_onclick()
    Dim strReturnValue

    If Not IsNumeric(Form.SelectedIndex.value) Then
        Exit Sub
    End If

    cmdPrintReview.disabled = True
    <%'Open the print-preview window, passing it the review ID:%>
    strReturnValue = window.showModalDialog("ArcPrintReview.asp?ReviewID=" & Form.rvwID.Value & _
        "&UserID=" & Form.UserID.value, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
    cmdPrintReview.disabled = False
End Sub
Sub Date_onkeypress(ctlDate)
    If ctlDate.value = "(MM/DD/YYYY)" Then
        ctlDate.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub

Sub Date_onblur(ctlDate)
    Dim intRowID
    
    If Trim(ctlDate.value) = "(MM/DD/YYYY)" Then
        ctlDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(ctlDate.value) Then
        MsgBox "The End Date must be a valid date - MM/DD/YYYY.", vbInformation, "Factor Maintenance"
        ctlDate.focus
        Exit Sub
    End If
    
    intRowID = Replace(ctlDate.ID,"txtEndDate","")
    Select Case GetButtonValue(intRowID)
        Case "End"
            If IsDate(ctlDate.value) Then
                If CDate(ctlDate.value) < CDate(document.all("tbd2R" & intRowID).innerText) Then
                    MsgBox "The End Date cannot be before the date last used, " &  FormatDateTime(document.all("tbd2R" & intRowID).innerText,2) & ".", vbInformation, "Factor Maintenance"
                    ctlDate.focus
                    Exit Sub
                Else
                    document.all("cmdLink" & intRowID).value = "&nbsp;Edit&nbsp;"
                End If
            End If
            document.all("cmdLink" & intRowID).disabled = False
            document.all("txtEndDate" & intRowID).disabled = True
        Case "Edit"
            If IsDate(ctlDate.value) Then
                If CDate(ctlDate.value) < CDate(document.all("tbd2R" & intRowID).innerText) Then
                    MsgBox "The End Date cannot be before the date last used, " &  FormatDateTime(document.all("tbd2R" & intRowID).innerText,2) & ".", vbInformation, "Factor Maintenance"
                    ctlDate.focus
                    Exit Sub
                End If
            Else
                document.all("cmdLink" & intRowID).value = "&nbsp;End&nbsp;"
            End If
            document.all("cmdLink" & intRowID).disabled = False
            document.all("txtEndDate" & intRowID).disabled = True
    End Select
End Sub

Sub Date_onfocus(ctlDate)
    If Trim(ctlDate.value) = "" Then
        ctlDate.value = "(MM/DD/YYYY)"
    End If
    ctlDate.select
End Sub
</SCRIPT>
</HEAD>

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
        <DIV id=divProgramsSelected class=DefTitleArea style="overflow:auto;left:631;top:-1;width:105;height:124">
            <SPAN id=lblProgramsSelected class=DefLabel style="LEFT:5; WIDTH:95; TOP:1;text-align:center">Programs</SPAN>
            <%
                Dim strOption
                Dim intOptionValue
                Dim intTop
                Dim strChecked
                Dim strOptions

                Set adCmd = GetAdoCmd("spArchiveLists")
                Set adRs = Server.CreateObject("ADODB.Recordset")
                AddParmIn adCmd, "@ListType", adVarchar, 255, "Program"
                'Call ShowCmdParms(adCmdPrg) '***DEBUG
                adRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
                
                intTop = 0
                intOptionValue = 0
                Do While Not adRs.EOF
                    strOption = adRs.Fields("ListItem").Value
                    strChecked = "Checked"
                    If CInt(Parse(adRs.Fields("ListItem").Value,"^",3)) >= 50 Then
                        intI = 6
                        strOption = Parse(strOption,"^",4)
                    Else
                        intI = Parse(adRs.Fields("ListItem").Value,"^",3)
                        strOption = Parse(strOption,"^",4)
                    End If
                    If strOption <> "" And InStr(strOptions,"[" & strOption & "]") = 0 Then
                        intOptionValue = intOptionValue + 1
                        intTop = intTop + 15
                        Response.Write "<INPUT type=""checkbox"" ID=chkPrg" & intOptionValue & " " & strChecked & " style=""LEFT:1;TOP:" & intTop & """ NAME=chkProgram" & intOptionValue & ">"
                        Response.Write "<SPAN id=lblProgram" & intOptionValue & " onclick=lblProgram_onClick(" & intOptionValue & ") class=DefLabel style=""LEFT:21; WIDTH:80; TOP:" & intTop & """>" & strOption & "</SPAN>"
                        Response.Write "<INPUT type=hidden id=hidPrgInfo" & intOptionValue & " value=" & intI & ">"
                        strOptions = strOptions & "[" & strOption & "]"
				    End If
                    adRs.MoveNext
                Loop
                Response.Write "<SPAN id=lblProgramCount style=""LEFT:-1000; visibility:hidden"">" & intOptionValue & "</SPAN>"
            %>
        </DIV>
        
        <!-- Column 1 -->
        <SPAN id=lblReviewID class=DefLabel style="LEFT:1; WIDTH:<%=intColumn1-5%>;TOP:<%=intRow1%>;text-align:right">
            Review ID
        </SPAN>
        <INPUT id=txtReviewID
            style="LEFT:<%=intColumn1%>;WIDTH:<%=intColumn2-intColumn1%>;TOP:<%=intRow1%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReviewID)" NAME="txtReviewID">

        <SPAN id=lblCaseNumber class=DefLabel style="LEFT:1;WIDTH:<%=intColumn1-5%>;TOP:<%=intRow2%>;text-align:right">
            Case Number
        </SPAN>
        <INPUT id=txtCaseNumber
            style="LEFT:<%=intColumn1%>; WIDTH:<%=intColumn2-intColumn1%>;TOP:<%=intRow2%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtCaseNumber)" NAME="txtCaseNumber">

        <SPAN id=lblResponse class=DefLabel style="LEFT:1;WIDTH:<%=intColumn1-5%>;TOP:<%=intRow3%>;text-align:right" unselectable=on>
            Response
        </SPAN>
        <SELECT id=cboResponse
            style="LEFT:<%=intColumn1%>; WIDTH:160;TOP:<%=intRow3%>;z-index:5000" onkeydown="Gen_onkeydown" tabIndex=<%=GetTabIndex%> NAME="cboResponse">
            <OPTION VALUE=0 SELECTED>&ltAll&gt
            <%=BuildList("WorkerResponse","",0,0,0)%>
        </SELECT>

        <!-- Column 2 -->
        <SPAN id=lblRvwr class=DefLabel style="LEFT:<%=intColumn3-65%>; WIDTH:60;TOP:<%=intRow1%>;text-align:right">
            <%=gstrRvwTitle%>
        </SPAN>
        <INPUT type="text" ID=txtReviewer NAME="txtReviewer" tabIndex=<%=GetTabIndex%> 
            onfocus="CmnTxt_onfocus(txtReviewer)"
            style="LEFT:<%=intColumn3%>; WIDTH:<%=intTextBoxWidth%>; TOP:<%=intRow1%>">

        <SPAN id=lblWrkr class=DefLabel style="LEFT:<%=intColumn3-65%>; WIDTH:60;TOP:<%=intRow2%>;text-align:right">
            <%=gstrWkrTitle%>
        </SPAN>
        <INPUT type="text" ID=txtWorker NAME="txtWorker" tabIndex=<%=GetTabIndex%> 
            onfocus="CmnTxt_onfocus(txtWorker)"
            style="LEFT:<%=intColumn3%>; WIDTH:<%=intTextBoxWidth%>; TOP:<%=intRow2%>">

        <SPAN id=lblSupervisor class=DefLabel style="LEFT:<%=intColumn3-95%>; WIDTH:90;TOP:<%=intRow3%>;text-align:right">
            <%=gstrSupTitle %>
        </SPAN>
        <INPUT type="text" ID=txtSupervisor NAME="txtSupervisor" tabIndex=<%=GetTabIndex%> 
            onfocus="CmnTxt_onfocus(txtSupervisor)"
            style="LEFT:<%=intColumn3%>; WIDTH:<%=intTextBoxWidth%>; TOP:<%=intRow3%>">

        <SPAN id=lblManager Class=DefLabel style="LEFT:<%=intColumn3-85%>; WIDTH:80;TOP:<%=intRow4-2%>;text-align:right">
            <%=gstrMgrTitle%>
        </SPAN>
        <SELECT id=cboManager
            style="LEFT:<%=intColumn3%>; WIDTH:<%=intTextBoxWidth%>;TOP:<%=intRow4-2%>;z-index:5000" onkeydown="Gen_onkeydown" tabIndex=<%=GetTabIndex%> NAME="cboManager">
            <%=mstrManagers%>
        </SELECT>

        <SPAN id=lblOffice Class=DefLabel style="LEFT:<%=intColumn3-85%>; WIDTH:80;TOP:<%=intRow4+20%>;text-align:right">
            <%=gstrOffTitle%>
        </SPAN>
        <SELECT id=cboOffice
            style="LEFT:<%=intColumn3%>; WIDTH:<%=intTextBoxWidth%>;TOP:<%=intRow4+20%>;z-index:5000" onkeydown="Gen_onkeydown" tabIndex=<%=GetTabIndex%> NAME="cboOffice">
            <%=mstrOffices%>
        </SELECT>

        <SPAN id=lblReviewClass class=DefLabel style="LEFT:1;WIDTH:<%=intColumn1-5%>;TOP:<%=intRow4%>;text-align:right" unselectable=on>
            Review Class
        </SPAN>
        <SELECT id=cboReviewClass
            style="LEFT:<%=intColumn1%>; WIDTH:160;TOP:<%=intRow4%>;z-index:5000" onkeydown="Gen_onkeydown" tabIndex=<%=GetTabIndex%> NAME="cboReviewClass">
            <OPTION VALUE=0 SELECTED>&ltAll&gt
            <%=BuildList("ReviewClass",Null,0,0,0)%>
        </SELECT>

        <!-- Column 3 -->
        <SPAN id=lblReviewDate class=DefLabel style="LEFT:450; WIDTH:100; TOP:<%=intRow1%>">
            Review Dates
        </SPAN>
        <SPAN id=lblReviewDateStart class=DefLabel style="LEFT:450; WIDTH:100; TOP:22">
            From
        </SPAN>
        <INPUT id=txtReviewDate title="Beginning Review Date" tabindex=<%=GetTabIndex%>
            style="LEFT:480; WIDTH:80; TOP:22" maxlength=10
            onkeydown="Gen_onkeydown" onblur=Date_onblur(txtReviewDate) onkeypress=Date_onkeypress(txtReviewDate) onfocus=Date_onfocus(txtReviewDate) NAME="txtReviewDate">
        <SPAN id=lblReviewDateEnd class=DefLabel style="LEFT:450; WIDTH:25; TOP:45">
            To
        </SPAN>
        <INPUT id=txtReviewDateEnd title="Ending Review Date" tabindex=<%=GetTabIndex%>
            style="LEFT:480; WIDTH:80; TOP:45"  maxlength=10
            onkeydown="txtReviewDateEnd" onblur=Date_onblur(txtReviewDateEnd) onkeypress=Date_onkeypress(txtReviewDateEnd) onfocus=Date_onfocus(txtReviewDateEnd) NAME="txtReviewDateEnd">
        
        <BUTTON id=cmdFind class=DefBUTTON title="Search for matching record(s)" 
            style="LEFT:570;TOP:<%=intRow1%>;HEIGHT:<%=intButtonHeight%>;WIDTH:55" tabindex=<%=GetTabIndex%>
            accessKey=F>
            <U>F</U>ind
        </BUTTON>

        <BUTTON id=cmdClear class=DefBUTTON title="Clear all search criteria" 
            style="LEFT:570;TOP:<%=intRow2%>;HEIGHT:<%=intButtonHeight%>;WIDTH:55" tabindex=<%=GetTabIndex%>
            accessKey=C>
            <U>C</U>lear
        </BUTTON>

        <SPAN id=lblStatus class=DefLabel style="LEFT:20; WIDTH:500; TOP:<%=intRow4 + 27%>">
            Enter search criteria and click [Find].
        </SPAN>

        <DIV id=lstResults class=DefPageFrame style="LEFT:0;WIDTH:736;border-left-style:none;border-right-style:solid; HEIGHT:260; TOP:122">
            <IFRAME ID=fraResults src="ArcFindResults.asp?Load=N"
                STYLE="positon:absolute; LEFT:0; WIDTH:735; HEIGHT:258; TOP:0; BORDER-style:none" FRAMEBORDER=0>
            </IFRAME>
        </DIV>
        
        <BUTTON id=cmdAddToArchive class=DefBUTTON style="LEFT:10; TOP:385" tabIndex=1>
            Move Reviews To Archive...
        </BUTTON>

        <BUTTON id=cmdPrintReview class=DefBUTTON disabled style="LEFT:180; TOP:385" tabIndex=1>
            Print Selected Review
        </BUTTON>

        <BUTTON id=cmdReports class=DefBUTTON style="LEFT:350; TOP:385" tabIndex=1>
            Archive Reports...
        </BUTTON>
        
        <BUTTON id=cmdClose class=DefBUTTON style="LEFT:555; FONT-WEIGHT:bold; TOP:385" tabIndex=1>
            Close
        </BUTTON>
    </DIV>
    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY: hidden"" ACTION=""Main.asp"" ID=Form>" & vbCrLf
        Call CommonFormFields()
    	WriteFormField "FormAction", ReqForm("FormAction")

        WriteFormField "rvwID", ReqForm("rvwID")
        WriteFormField "ReviewDate", ReqForm("ReviewDate")
        WriteFormField "ReviewDateEnd", ReqForm("ReviewDateEnd")
        WriteFormField "CaseNumber", ReqForm("CaseNumber")
        WriteFormField "Response", ReqForm("Response")
        WriteFormField "Reviewer", ReqForm("Reviewer")
        WriteFormField "Supervisor", ReqForm("Supervisor")
        WriteFormField "Manager", ReqForm("Manager")
        WriteFormField "Office", ReqForm("Office")
        WriteFormField "Director", ReqForm("Director")
        WriteFormField "WorkerName", ReqForm("WorkerName")
        WriteFormField "SortOrder", ""
        WriteFormField "SelectedIndex", ""
        WriteFormField "ReviewClass", ReqForm("ReviewClass")
        WriteFormField "UseWarning", "Yes"

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
