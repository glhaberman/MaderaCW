<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>

<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->

<%
Dim intLine, intI, intColID
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
Dim intRow4, intRow5
Dim intTextBoxWidth
Dim mstrProgramList
Dim mdctColumns     'Holds all columns available from the stored procedure
Dim oColumn, mstrRespWrite
Dim mstrShowColumns, strOptions, strOrder, strChecked
Dim mintMaxCaseNumLen

mintMaxCaseNumLen = GetAppSetting("MaxCaseNumberLength")

Set mdctColumns = CreateObject("Scripting.Dictionary")
Set adCmd = GetAdoCmd("spProfileSettingGet")
    AddParmIn adCmd, "@UserID", adVarChar, 20, gstrUserID
    AddParmIn adCmd, "@SettingName", adVarChar, 50, "ShowColumns"
    AddParmOut adCmd, "@SettingValue", adVarChar, 255
    'ShowCmdParms(adCmd) '***DEBUG
    adCmd.Execute
    mstrShowColumns = adCmd.Parameters("@SettingValue").Value
If IsNull(mstrShowColumns) Or mstrShowColumns = "" Then mstrShowColumns = "1^2^3^4^5^6^7^8^"

' Call stored proc and return and empty recordset to get column names
Set adRs = Server.CreateObject("ADODB.Recordset")
Set adCmd = GetAdoCmd("spReviewFind")
	AddParmIn adCmd, "@AliasID", adInteger, 20, -1
	AddParmIn adCmd, "@Admin", adBoolean, 0, 0
    AddParmIn adCmd, "@QA", adBoolean, 0, 0
    AddParmIn adCmd, "@UserID", adVarChar, 20, ""
    AddParmIn adCmd, "@casID", adInteger, 0, Null
    AddParmIn adCmd, "@casNumber", adVarChar, 20, Null
    AddParmIn adCmd, "@ReviewDate", adDBTimeStamp, 0, Null
    AddParmIn adCmd, "@ReviewDateEnd", adDBTimeStamp, 0, Null
    AddParmIn adCmd, "@WorkerName", adVarChar, 100, Null
    AddParmIn adCmd, "@Submitted", adVarchar, 1, NULL
    AddParmIn adCmd, "@Response", adInteger, 0, Null
    AddParmIn adCmd, "@Reviewer", adVarChar, 100, Null
    AddParmIn adCmd, "@PrgID", adVarchar, 255, Null
    AddParmIn adCmd, "@WorkerID", adVarchar, 20, Null
    AddParmIn adCmd, "@Supervisor", adVarchar, 100, Null
    AddParmIn adCmd, "@SupervisorID", adVarchar, 20, Null
    AddParmIn adCmd, "@ReviewClassID", adInteger, 0, Null
    Call adRs.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
Set adCmd = Nothing

strOptions = ""
mstrShowColumns = "^" & mstrShowColumns & "^"
mstrRespWrite = ""
For intI = 1 To adRs.Fields.Count
    strChecked = ""
    If InStr(mstrShowColumns,"^" & intI & "^") > 0 Then
        strChecked = "checked"
    End If
    mdctColumns.Add CInt(intI), adRs.Fields(intI-1).Name & "^" & strChecked
    mstrRespWrite = mstrRespWrite & AddColumn(intI, strChecked, adRs.Fields(intI-1).Name, intI)
Next

' Load programs to display names
Set adRs = Server.CreateObject("ADODB.Recordset")
Set adCmd = GetAdoCmd("spGetProgramList")
    AddParmIn adCmd, "@PrgID", adVarchar, 255, NULL
Call adRs.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
adRs.Sort = "prgID"
mstrProgramList = "|"
Do While Not adRs.EOF
    mstrProgramList = mstrProgramList & adRs.Fields("prgID").value & "^" & adRs.Fields("prgCode").value & "|"
    adRs.MoveNext
Loop
mstrPageTitle = "Find Case Review For Edit"
intLine = -1 'Used to determine the number of matching results.
intBottomRow = 425
intButtonHeight = 20
intColumn1 = 70
intColumn2 = 170
intColumn3 = 245
intRow1 = 5
intRow2 = 29
intRow3 = 53
intRow4 = 75
intRow5 = 105
intTextBoxWidth = 150

Function AddColumn(intRowID, strChecked, strFieldName, intOrder)
    Dim strRespWrite
    strRespWrite = strRespWrite & "<input type=checkbox id=chkColumn" & intRowID
    strRespWrite = strRespWrite & " style=""left:2;top:" & (intOrder*20) + 20 & """ " & strChecked & " />"
    strRespWrite = strRespWrite & "<span id=lblColumn" & intRowID & " class=DefLabel onclick=lblColumn_onclick(" & intRowID & ")"
    strRespWrite = strRespWrite & " style=""cursor:hand;LEFT:25;WIDTH:120;TOP:" & (intOrder*20) + 20 & """>" & strFieldName & "</span>"
    AddColumn = strRespWrite
End Function
%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>
  
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
'Dim mctlStaff
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>
Dim mdctPrograms
Dim mdctColumns
Dim mblnMainClosed
Dim mblnCloseClicked

Sub window_onload()
    Dim intI
	Dim strProgramList

    Call CheckForValidUser()
    If Trim(Form.UserID.Value) = "" Then Exit Sub
	Set mdctColumns = CreateObject("Scripting.Dictionary")
	Set mdctPrograms = CreateObject("Scripting.Dictionary")
    Set mdctPrograms = LoadDictionaryObject("<% = mstrProgramList %>")
    
    mblnSetFocusToMain = True
    mblnMainClosed = False
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>

    Call SizeAndCenterWindow(767, 520, True)
	
	txtReviewID.Value = Form.casID.Value
    txtReviewDate.Value = Form.ReviewDate.Value
    txtCaseNumber.Value = Form.CaseNumber.Value
    txtWorker.Value = Form.WorkerName.Value
    cboSubmitted.Value = Form.Submitted.Value
    cboResponse.Value = Form.Response.Value
    txtReviewer.Value = Form.Reviewer.Value
    txtSupervisor.Value = Form.Supervisor.Value

    If IsNumeric(Form.SelectedIndex.Value) Then
        If CLng(Form.SelectedIndex.Value) > 0 Then
            Call Result_onclick(1)
            cmdEdit.disabled = False
            cmdPrint.disabled = False
            cmdPrintList.disabled = False
        End If
    Else
        cmdEdit.disabled = True
        cmdPrint.disabled = True
        cmdEditWR.disabled = True
        cmdPrintList.disabled = True
    End If
    
    PageFrame.disabled = False
    PageBody.style.cursor = "default"
    txtReviewID.focus
    If txtReviewer.value = "" Then txtReviewer.value = "<All>"
    If txtWorker.value = "" Then txtWorker.value = "<All>"
    If txtSupervisor.value = "" Then txtSupervisor.value = "<All>"
    
    <%
    For Each oColumn In mdctColumns
        Response.Write "mdctColumns.Add " & oColumn & ",""" & mdctColumns(oColumn) & """" & vbCrLf
    Next
    %>
End Sub

Sub txtWorker_onblur()
    If Trim(txtWorker.value) = "" Then txtWorker.value = "<All>"
End Sub

Sub txtSupervisor_onblur()
    If Trim(txtSupervisor.value) = "" Then txtSupervisor.value = "<All>"
End Sub

Sub cmdPrintList_onclick()
    Call fraResults.PrintList()
End Sub

<%'If timer detects that Main has been closed, this sub will be called.%>
Sub MainClosed()
    mblnMainClosed = True
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
    Dim blnClose
    
    blnClose = False
    On Error Resume Next
    Call window.opener.ManageWindows(2,"Close")
    If Err.number <> 0 Then blnClose = True
    On Error Goto 0
    If blnClose Then window.close
End Sub

Sub txtReviewID_onchange()
    txtReviewDate.Value = ""
    txtCaseNumber.Value = ""
End Sub

Sub txtReviewDate_onkeypress
    If txtReviewDate.value = "(MM/DD/YYYY)" Then
        txtReviewDate.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub
Sub txtReviewDate_onblur
    If Trim(txtReviewDate.value) = "(MM/DD/YYYY)" Then
        txtReviewDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtReviewDate.value) Then
        MsgBox "The Start Date must be a valid date - MM/DD/YYYY.", vbInformation, "Find Case Review"
        txtReviewDate.focus
    ElseIf IsDate(txtReviewDate.value) Then
        If CDate(txtReviewDate.value) < CDate("01/01/1900") Then
            MsgBox "The Start Date must be a valid date - MM/DD/YYYY." & vbCrLf & "Dates prior to 01/01/1900 are not considered valid for this entry.", vbInformation, "Find Case Review"
            txtReviewDate.focus
        End If
    ElseIf IsDate(txtReviewDateEnd.value) And IsDate(txtReviewDate.value) Then
        If CDate(txtReviewDateEnd.value) < CDate(txtReviewDate.value) Then
            MsgBox "The Start Date must be before the end date." & vbCrLf, vbInformation, "Find Case Review"
            txtReviewDate.focus
        End If
    End If
End Sub
Sub txtReviewDate_onfocus
    If Trim(txtReviewDate.value) = "" Then
        txtReviewDate.value = "(MM/DD/YYYY)"
    End If
    txtReviewDate.select
End Sub

Sub txtReviewDate_onchange()
    txtReviewID.Value = ""
    txtCaseNumber.Value = ""
End Sub

Sub txtReviewDateEnd_onkeypress
    If txtReviewDateEnd.value = "(MM/DD/YYYY)" Then
        txtReviewDateEnd.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub
Sub txtReviewDateEnd_onblur
    If Trim(txtReviewDateEnd.value) = "(MM/DD/YYYY)" Then
        txtReviewDateEnd.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtReviewDateEnd.value) Then
        MsgBox "The End Date must be a valid date - MM/DD/YYYY.", vbInformation, "Find Case Review"
        txtReviewDateEnd.focus
    ElseIf IsDate(txtReviewDateEnd.value) Then
        If CDate(txtReviewDateEnd.value) < CDate("01/01/1900") Then
            MsgBox "The End Date must be a valid date - MM/DD/YYYY." & vbCrLf & "Dates prior to 01/01/1900 are not considered valid for this entry.", vbInformation, "Find Case Review"
            txtReviewDateEnd.focus
        End If
    ElseIf IsDate(txtReviewDateEnd.value) And IsDate(txtReviewDate.value) Then
        If CDate(txtReviewDateEnd.value) < CDate(txtReviewDate.value) Then
            MsgBox "The End Date must be after the start date." & vbCrLf, vbInformation, "Find Case Review"
            txtReviewDateEnd.focus
        End If
    End If
End Sub
Sub txtReviewDateEnd_onfocus
    If Trim(txtReviewDateEnd.value) = "" Then
        txtReviewDateEnd.value = "(MM/DD/YYYY)"
    End If
    txtReviewDateEnd.select
End Sub

Sub txtCaseNumber_onchange()
    txtReviewID.Value = ""
    txtReviewDate.Value = ""
End Sub

Sub CheckColumnDiv()
    If cmdColumns.value = "Save Search Columns" Then
        Call cmdColumns_onclick()
    End If
End Sub
Sub cmdFind_onclick()
    Dim blnCriteria
    Dim intResp
    Dim strParms
    Dim intI
    Dim oColumn, strShowColumns, strPrograms

    Call CheckColumnDiv() 
    txtReviewID.Value = Trim(txtReviewID.Value)
    txtReviewDate.Value = Trim(txtReviewDate.Value)
    txtCaseNumber.Value = Trim(txtCaseNumber.Value)
    
    blnCriteria = False

    strPrograms = ""
    For intI = 1 To 4
        If document.all("chkProgram" & intI).checked = True Then
            strPrograms = strPrograms & "[" & intI & "]"
        End If
    Next
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
    
    If Trim(txtWorker.Value) <> "" And txtWorker.Value <> "<All>" Then
        blnCriteria = True
    End If    
    If Trim(txtSupervisor.Value) <> "" And txtSupervisor.Value <> "<All>" Then
        blnCriteria = True
    End If    
    If Trim(cboResponse.Value) <> "0" Then
        blnCriteria = True
    End If    
    If Trim(cboReviewClass.Value) <> "0" Then
        blnCriteria = True
    End If    
    If Trim(cboSubmitted.Value) <> "0" Then
        blnCriteria = True
    End If    
    If Trim(txtReviewer.Value) <> "" And txtReviewer.Value <> "<All>" Then
        blnCriteria = True
    End If    

    If Not blnCriteria Then
        intResp = MsgBox("This may return a large number of search results.  " & vbcrlf & vbcrlf & "Do you wish to continue?",vbYesNo + vbQuestion,"Find Matching Case Reviews")
        If intResp = vbNo Then
            PageFrame.disabled = False
            PageFrame.style.visibility = "visible"
            Exit Sub
        End If
    End If

    Form.casID.Value = txtReviewID.Value
    Form.ReviewDate.Value = txtReviewDate.Value
    Form.ReviewDateEnd.Value = txtReviewDateEnd.Value
    Form.CaseNumber.Value = txtCaseNumber.Value
    Form.Submitted.Value = cboSubmitted.Value
    Form.Response.Value = cboResponse.Value
    Form.ReviewClass.value = cboReviewClass.value
    Form.Reviewer.Value = txtReviewer.Value
    Form.WorkerName.Value = txtWorker.Value
    Form.Supervisor.Value = txtSupervisor.value

    PageBody.style.cursor = "wait"

    strShowColumns = ""
    For Each oColumn In mdctColumns
        If Parse(mdctColumns(oColumn),"^",2) = "checked" Then
            strShowColumns = strShowColumns & oColumn & "^"
        End If
    Next

    strParms = "<% = gstrUserID %>^<% = gblnUserAdmin %>^<% = gblnUserQA %>"
    strParms = strParms & "^" & Form.casID.value
    strParms = strParms & "^" & Form.ReviewDate.value
    strParms = strParms & "^" & Form.CaseNumber.value
    strParms = strParms & "^" & Form.Submitted.value
    strParms = strParms & "^" & Form.Response.value
    strParms = strParms & "^" & ReplaceAllWithBlank(Form.Reviewer.value)
    strParms = strParms & "^" & Form.ProgramsSelected.value
    strParms = strParms & "^" & Form.ReviewDateEnd.value
    strParms = strParms & "^" & ReplaceAllWithBlank(Form.Supervisor.value)
    strParms = strParms & "^" & ReplaceAllWithBlank(Form.SupervisorID.value)
    strParms = strParms & "^" & ReplaceAllWithBlank(Form.WorkerName.value)
    strParms = strParms & "^" & ReplaceAllWithBlank(Form.WorkerID.value)
    strParms = strParms & "^" & Form.ReviewClass.value
    strParms = strParms & "^" & Form.SortOrder.value
    strParms = strParms & "^<% = glngAliasPosID %>"
    ' Load search in IFRAME
    fraResults.frameElement.src = "FindCaseResults.asp?Load=Y&ShowColumns=" & strShowColumns & "&ParmList=" & strParms
End Sub

Function ReplaceAllWithBlank(strValue)
    If strValue = "<All>" Then
        ReplaceAllWithBlank = ""
    ElseIf strValue = "<Left Blank>" Then
        ReplaceAllWithBlank = "*BLANK*"
    Else
        ReplaceAllWithBlank = strValue
    End If
End Function

Sub cmdEdit_onclick()
    Call CheckColumnDiv()
    Call EditRecord()
End Sub

Sub EditRecord()
    If IsNull(Form.SelectedIndex.Value) Or Trim(Form.SelectedIndex.Value) = "" Then
        Exit Sub
    End If
    
    window.opener.Form.rvwID.Value = Form.rvwID.Value
    window.opener.Form.Action = "CaseAddEdit.asp"
    
    Call SizeAndCenterWindow(767, 520, True)

    window.opener.Form.FormAction.Value = "GetRecord"
    Call window.opener.ManageWindows(1,"EditReview")
End Sub

Sub cmdPrint_onclick()
    Dim strReturnValue
    
    cmdPrint.disabled = True
    Call CheckColumnDiv()
    <%'Open the print-preview window, passing it the review ID:%>
    strReturnValue = window.showModalDialog("PrintReview.asp?AuditRead=True&UserID=<%=gstrUserID%>&ReviewID=" & Form.rvwID.Value, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
    cmdPrint.disabled = False
End Sub

Sub cmdClear_onclick()
    txtReviewID.Value = ""
    txtReviewDate.Value = ""
    txtReviewDateEnd.Value = ""
    txtCaseNumber.Value = ""
    cboSubmitted.Value = 0
    cboResponse.Value = 0
    txtReviewer.value = "<All>"
    txtWorker.value = "<All>"
    'txtWorkerID.value = "<All>"
    txtSupervisor.Value = "<All>"
    cboReviewClass.selectedIndex = 0
    Call CheckColumnDiv()
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

Sub lblProgram_onClick(intWhich)
    If document.all("chkProgram" & intWhich).disabled = True Then
        Exit Sub
    End If
    document.all("chkProgram" & intWhich).checked = Not document.all("chkProgram" & intWhich).checked
    Call RebuildProgramList()
End Sub

Sub chkProgram_onClick(intWhich)
    Call RebuildProgramList()
End Sub

Sub RebuildProgramList()
    Dim oPrg, strProgramList
    
    strProgramList = ""
    For Each oPrg In mdctPrograms
        If CInt(oPrg) < 50 Then
            If document.all("chkProgram" & oPrg).checked = True Then
                strProgramList = strProgramList & "[" & oPrg & "]"
            End If
        End If
    Next
    
    Form.ProgramsSelected.value = strProgramList
    window.opener.Form.ProgramsSelected.value = strProgramList
End Sub

Sub StaffText_OnBlur(ctlTextBox)
    If (ctlTextBox.value = "" Or InStr(ctlTextBox.value,"<") > 0) Then
        ctlTextBox.value = "<All>"
        'Set mctlStaff = ctlTextBox
    End If
End Sub

Sub lblColumn_onclick(intRowID)
    document.all("chkColumn" & intRowID).checked = Not document.all("chkColumn" & intRowID).checked
End Sub

Sub cmdSelectAll_onclick()
    Call CheckSearchBoxes(True)
End Sub

Sub cmdSelectNone_onclick()
    Call CheckSearchBoxes(False)
End Sub
Sub CheckSearchBoxes(blnCheck)
    Dim intI
    
    For intI = 1 To mdctColumns.Count
        document.all("chkColumn" & intI).checked = blnCheck
    Next
End Sub

Sub cmdColumns_onclick()
    Dim oColumn
    Dim strRecord
    Dim intI, intJ
    Dim strChecked
    
    If cmdColumns.value = "Set Search Columns" Then
        intI = 1
        For Each oColumn In mdctColumns
            strRecord = mdctColumns(oColumn)
            If Parse(strRecord,"^",2) = "checked" Then
                document.all("chkColumn" & oColumn).checked = True
            Else
                document.all("chkColumn" & oColumn).checked = False
            End If
            intI = intI + 1
        Next
        divColumns.style.left = 460
        cmdColumns.value = "Save Search Columns"
    Else
        For Each oColumn In mdctColumns
            strRecord = mdctColumns(oColumn)
            If document.all("chkColumn" & oColumn).checked = True Then
                mdctColumns(oColumn) = Parse(strRecord,"^",1) & "^checked"
            Else
                mdctColumns(oColumn) = Parse(strRecord,"^",1) & "^"
            End If
        Next
        divColumns.style.left = -1511
        cmdColumns.value = "Set Search Columns"
    End If
End Sub

Sub NavigateFix(strAction)
    If strAction = "Open" Then
        divSubmitted.style.left = <%=intColumn1%>
        cboSubmitted.style.left = -1000
    Else
        divSubmitted.style.left = -1000
        cboSubmitted.style.left = <%=intColumn1%>
    End If
End Sub
</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=0 leftMargin=5 topMargin=0 rightMargin=5 style="cursor:wait">
  
    <DIV id=Header class=DefTitleArea style="width:614;height:30;top:1">
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
  
    <DIV id=divProgramsSelected class=DefTitleArea style="top:1;left:623;width:124;height:135;border-bottom-style:none;z-index:123">
        <DIV id=divPrograms class=DefPageFrame style="VISIBILITY:visible;border-style:none; WIDTH:122; LEFT:0; TOP:0;height:134;
            BACKGROUND-COLOR: <%=gstrPageColor%>;overflow:auto">
            <SPAN id=lblProgramsSelected class=DefLabel style="LEFT:5; WIDTH:95; TOP:1">Programs Selected</SPAN>
        <%
        Dim mstrProgramNames
        Dim strOption, intOptionValue
        Dim mstrOptions
        Dim intTop
                
        intTop = 0
        mstrProgramNames = ""
        adRs.Filter = "prgID<50"
        If adRs.RecordCount > 0 Then
            adRs.MoveFirst
        End If
        Do While Not adRs.EOF
            strOption = ""
            strChecked = ""
            strOption = adRs.Fields("prgCode").Value
            intOptionValue = adRs.Fields("prgID").Value
            If InStr(ReqForm("ProgramsSelected"),"[" & intOptionValue & "]") > 0 Then
                strChecked = "Checked"
            End If
            If strOption <> "" Then
                intTop = intTop + 15
                'Set checkbox to checked and disabled, as WV is a single program app
                'strChecked = "checked disabled=true "
                Response.Write "<INPUT type=""checkbox"" ID=chkProgram" & intOptionValue & " " & strChecked & " onclick=chkProgram_onClick(" & intOptionValue & ") style=""LEFT:1;TOP:" & intTop & """ NAME=chkProgram" & intOptionValue & ">"
                Response.Write "<SPAN id=lblProgram" & intOptionValue & " title=""" & adRs.Fields("prgShortTitle").value & """ onclick=lblProgram_onClick(" & intOptionValue & ") class=DefLabel style=""LEFT:21; WIDTH:80; TOP:" & intTop & """>" & strOption & "</SPAN>"
			End If
            adRs.MoveNext
        Loop
        %>
        </DIV>
    </DIV>

    <% Call WriteNavigateControls(-2,30,gstrBackColor) %>
    <DIV id=PageFrame class=DefPageFrame disabled=true style="BORDER-TOP-STYLE:none;WIDTH:737; HEIGHT:455; TOP:31">
        <!-- Column 1 -->
        <SPAN id=lblReviewID class=DefLabel style="LEFT:1; WIDTH:<%=intColumn1-5%>;TOP:<%=intRow1%>;text-align:right">
            Review ID
        </SPAN>
        <INPUT id=txtReviewID title="Review ID"
            style="LEFT:<%=intColumn1%>;WIDTH:<%=intColumn2-intColumn1%>;TOP:<%=intRow1%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReviewID)" NAME="txtReviewID">
        <SPAN id=lblCaseNumber class=DefLabel style="LEFT:1;WIDTH:<%=intColumn1-5%>;TOP:<%=intRow2%>;text-align:right">
            Case Number
        </SPAN>
        <INPUT id=txtCaseNumber title="Case Number"
            style="LEFT:<%=intColumn1%>; WIDTH:<%=intColumn2-intColumn1%>;TOP:<%=intRow2%>" 
            maxlength=<%=mintMaxCaseNumLen%> tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtCaseNumber)" NAME="txtCaseNumber">

        <SPAN id=lblSubmitted class=DefLabel style="LEFT:1;WIDTH:<%=intColumn1-5%>;TOP:<%=intRow3%>;text-align:right">
            Submitted
        </SPAN>

		<DIV id=divSubmitted style="LEFT:-1000; WIDTH:<%=intColumn2-intColumn1%>;TOP:<%=intRow3%>;height:19;border-style:solid;border-width:1;
		    background-color:white;color:gray">
		    <DIV id=divSubmittedBtn style="LEFT:<%=intColumn2-intColumn1-18%>; WIDTH:15; TOP:0"><IMG src="downclickbutton.bmp"></DIV>
		</DIV>

        <SELECT id=cboSubmitted title="Review Submitted" tabindex=<%=GetTabIndex%>
            style="LEFT:<%=intColumn1%>; WIDTH:<%=intColumn2-intColumn1%>;TOP:<%=intRow3%>"
            onkeydown="Gen_onkeydown"
            tabIndex=<%=GetTabIndex%> NAME="cboSubmitted">
            <OPTION VALUE="0" SELECTED>&ltAll&gt
            <OPTION value="Y">Yes
            <OPTION value="N">No
        </SELECT>

        <!-- Column 2 -->
        <SPAN id=lblResponse class=DefLabel style="LEFT:<%=intColumn3-65%>;WIDTH:60;TOP:<%=intRow1%>;text-align:right">
            Response
        </SPAN>

        <SELECT id=cboResponse title="<%=gstrWkrTitle%> Response"
            style="LEFT:<%=intColumn3%>; WIDTH:<%=intTextBoxWidth%>;TOP:<%=intRow1%>" tabindex=<%=GetTabIndex%>
            onkeydown="Gen_onkeydown"
            tabIndex=<%=GetTabIndex%> NAME="cboResponse">
            <OPTION VALUE=0 SELECTED>&ltAll&gt
            <%=BuildList("WorkerResponse","",0,0,0)%>
        </SELECT>

        <SPAN id=lblReviewClass class=DefLabel style="LEFT:<%=intColumn3-75%>; WIDTH:70;TOP:<%=intRow2%>;text-align:right">
            Review Class
        </SPAN>
        <SELECT id=cboReviewClass 
            style="LEFT:<%=intColumn3%>; WIDTH:<%=intTextBoxWidth%>; TOP:<%=intRow2%>"  
            tabIndex=<%=GetTabIndex%> NAME="cboReviewClass">
            <option value=0>&ltAll&gt</option>
            <%=BuildList("ReviewClass",Null,0,0,0)%>
        </SELECT>

        <SPAN id=lblRvwr class=DefLabel style="LEFT:<%=intColumn3-65%>; WIDTH:60;TOP:<%=intRow3%>;text-align:right">
            <%=gstrRvwTitle%>
        </SPAN>
        <INPUT type="text" ID=txtReviewer NAME="txtReviewer" tabIndex=<%=GetTabIndex%> 
            onkeydown="Gen_onkeydown" onfocus="CmnTxt_onfocus(txtReviewer)"
            onblur=StaffText_OnBlur(txtReviewer)
            style="LEFT:<%=intColumn3%>; WIDTH:<%=intTextBoxWidth%>; TOP:<%=intRow3%>">

        <!-- Column 3 -->
        <SPAN id=lblReviewDate class=DefLabel style="LEFT:420; WIDTH:100; TOP:<%=intRow1%>">
            Review Dates
        </SPAN>
        <SPAN id=lblReviewDateStart class=DefLabel style="LEFT:420; WIDTH:100; TOP:22">
            From
        </SPAN>
        <INPUT id=txtReviewDate title="Beginning Review Date" tabindex=<%=GetTabIndex%>
            style="LEFT:450; WIDTH:80; TOP:22" maxlength=10
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReviewDate)" NAME="txtReviewDate">
        <SPAN id=lblReviewDateEnd class=DefLabel style="LEFT:420; WIDTH:25; TOP:45">
            To
        </SPAN>
        <INPUT id=txtReviewDateEnd title="Ending Review Date" tabindex=<%=GetTabIndex%>
            style="LEFT:450; WIDTH:80; TOP:45"  maxlength=10
            onkeydown="Gen_onkeydown" onfocus="Gen_onfocus(txtReviewDateEnd)" NAME="txtReviewDateEnd">

        <!-- Staffing Rows -->
        <DIV id=divStaffing class=DefPageFrame style="LEFT:-1; HEIGHT:25; WIDTH:737; TOP:77; border:none;background-color:transparent">
            <SPAN id=lblWorkerName class=DefLabel style="LEFT:1;WIDTH:<%=intColumn1-5%>;TOP:0;text-align:right">
                <%=gstrWkrTitle%>
            </SPAN>
            
            <INPUT type="text" ID=txtWorker NAME="txtWorker" tabIndex=<%=GetTabIndex%> 
                onfocus="CmnTxt_onfocus(txtWorker)" title="<%=gstrWkrTitle%>"
                style="LEFT:<%=intColumn1%>;WIDTH:<%=intColumn2-intColumn1%>; TOP:0">

            <SPAN id=lblSupervisorName class=DefLabel style="LEFT:<%=intColumn3-65%>;WIDTH:60;TOP:0;text-align:right">
                <%=gstrSupTitle%>
            </SPAN>
            <INPUT type="text" ID=txtSupervisor NAME="txtSupervisor" tabIndex=<%=GetTabIndex%> 
                onfocus="CmnTxt_onfocus(txtSupervisor)" 
                style="LEFT:<%=intColumn3%>;WIDTH:150;TOP:0" />

            <BUTTON id=cmdColumns class=DefBUTTON title="Customize Search Results" 
                style="LEFT:480;TOP:0;HEIGHT:20;WIDTH:127" tabindex=<%=GetTabIndex%>>
                Set Search Columns
            </BUTTON>
		</DIV>
        <DIV id=divColumns class=DefPageFrame
            style="LEFT:-1511;HEIGHT:315;WIDTH:210;TOP:105;border:single;background-color:beige;z-index:2000;overflow:auto">
            <span class=DefLabel style="top:0;left:10;width:170;text-align:center"><b>Select Columns To Display</b></span>
            <BUTTON id=cmdSelectAll class=DefBUTTON style="LEFT:10; TOP:15;HEIGHT:<%=intButtonHeight%>;WIDTH:80" tabindex=<%=GetTabIndex%>>
                Select All
            </BUTTON>
            <BUTTON id=cmdSelectNone class=DefBUTTON style="LEFT:95; TOP:15;HEIGHT:<%=intButtonHeight%>;WIDTH:80" tabindex=<%=GetTabIndex%>>
                Select None
            </BUTTON>
            <%=mstrRespWrite%>
        </DIV>
        <BUTTON id=cmdFind class=DefBUTTON title="Search for matching record(s)" 
            style="LEFT:540;TOP:<%=intRow1%>;HEIGHT:<%=intButtonHeight%>;WIDTH:65" tabindex=<%=GetTabIndex%>
            accessKey=F>
            <U>F</U>ind
        </BUTTON>

        <BUTTON id=cmdClear class=DefBUTTON title="Clear all search criteria" 
            style="LEFT:540;TOP:<%=intRow2%>;HEIGHT:<%=intButtonHeight%>;WIDTH:65" tabindex=<%=GetTabIndex%>
            accessKey=C>
            <U>C</U>lear
        </BUTTON>
        <DIV id=lstResults class=DefPageFrame style="LEFT:0;WIDTH:736;border-left-style:none;border-right-style:solid; HEIGHT:315; TOP:105;background-color:red">
            <IFRAME ID=fraResults src="FindCaseResults.asp?Load=N"
                STYLE="positon:absolute; LEFT:0; WIDTH:735; HEIGHT:315; TOP:0; BORDER-style:none" FRAMEBORDER=0>
            </IFRAME>
        </DIV>

        <BUTTON id=cmdEdit class=DefBUTTON title="Edit the selected record" 
            style="LEFT:15; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>;WIDTH:100" tabindex=<%=GetTabIndex%>
            accessKey=E>
            <U>E</U>dit
        </BUTTON>
        <BUTTON id=cmdPrint class=DefBUTTON title="Print the selected record" 
            style="LEFT:120; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>;WIDTH:100" tabindex=<%=GetTabIndex%>
            accessKey=P>
            <U>P</U>rint Review
        </BUTTON>
        <BUTTON id=cmdPrintList class=DefBUTTON title="Print the results of the search" 
            style="LEFT:225; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>;WIDTH:100" tabindex=<%=GetTabIndex%>
            accessKey=L>
            Print <U>L</U>ist
        </BUTTON>
        <BUTTON id=cmdEditWR class=DefBUTTON title="Submit to Reports" 
            style="LEFT:-1330; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>;WIDTH:100" tabindex=<%=GetTabIndex%>
            accessKey=S>
            <U>S</U>ubmit To Reports
        </BUTTON>

        <SPAN id=lblStatus class=DefLabel style="LEFT:440; WIDTH:160; TOP:<%=intBottomRow-2%>; text-align:center">
            Enter search criteria and click [Find].
        </SPAN>

        <BUTTON id=cmdCancel class=DefBUTTON title="Close and return to previous" 
            style="LEFT:640; TOP:<%=intBottomRow%>;HEIGHT:<%=intButtonHeight%>; WIDTH:75" tabindex=<%=GetTabIndex%>>Cancel
        </BUTTON>
    </DIV>
  
    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY:hidden"" ACTION=""CaseEdit.ASP"" ID=Form>" & vbCrLf

    Call CommonFormFields()

    If Instr(ReqForm("CalledFrom"), "CaseAddEdit") > 0 And ReqForm("FormAction") <> "Y" Then
        WriteFormField "LastCaseIDEdited", ReqForm("rvwID")
    Else
        WriteFormField "LastCaseIDEdited", ReqForm("LastCaseIDEdited")
    End If
    If ReqForm("FormAction") = "Y" Then
        WriteFormField "casID", ReqForm("casID")
        WriteFormField "rvwID", ReqForm("rvwID")
        WriteFormField "ReviewDate", ReqForm("ReviewDate")
        WriteFormField "ReviewDateEnd", ReqForm("ReviewDateEnd")
        WriteFormField "CaseNumber", ReqForm("CaseNumber")
        WriteFormField "WorkerID", ReqForm("WorkerID")
        WriteFormField "Submitted", ReqForm("Submitted")
        WriteFormField "Response", ReqForm("Response")
        WriteFormField "Reviewer", ReqForm("Reviewer")
        WriteFormField "Supervisor", ReqForm("Supervisor")
        WriteFormField "SupervisorID", ReqForm("SupervisorID")
        WriteFormField "WorkerName", ReqForm("WorkerName")
        WriteFormField "ReviewClass", ReqForm("ReviewClass")
    Else
        WriteFormField "casID", ""
        WriteFormField "rvwID", ""
        WriteFormField "ReviewDate", ""
        WriteFormField "ReviewDateEnd", ""
        WriteFormField "CaseNumber", ""
        WriteFormField "WorkerID", ""
        WriteFormField "Submitted", "0"
        WriteFormField "Response", "0"
        WriteFormField "Reviewer", ""
        WriteFormField "Supervisor", ""
        WriteFormField "SupervisorID", ""
        WriteFormField "WorkerName", ""
        WriteFormField "ReviewClass", ""
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