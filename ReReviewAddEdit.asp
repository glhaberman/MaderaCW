<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ReReviewAddEdit.asp                                                 '
'  Purpose: The primary data entry screen for adding re-review records     '
'           and updating existing records.                                  '
'           The form is displayed when the user clicks [Enter Re-Review Reviews] '
'           from the app main screen.                                       '
'==========================================================================='
Dim madoRs
Dim mstrPageTitle
Dim adCmd
Dim mlngReReviewID
Dim mlngReReviewTypeID
Dim madoReReview
Dim madoReReviewElms
Dim mstrReReviewElements
Dim strHTML
Dim mstrReReviewType
Dim mlngWindowID
Dim mlngReviewID
Dim intI

%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%

mlngReReviewID = ReqForm("ReReviewID")
mlngReReviewTypeID = ReqForm("ReReviewTypeID")
If ReqForm("ReReviewTypeID") = 0 Then
    mstrReReviewType = gstrEvaluation
    mlngWindowID = 4
Else
    mstrReReviewType = "Corrective Action "
    mlngWindowID = 7
End If
'Instantiate the recordset that is reused for temporary results or queries:
Set madoRs = Server.CreateObject("ADODB.Recordset")

'Set the page title:
mstrPageTitle = Trim(gstrTitle & " " & gstrAppName)

'Retrieve the Re-Review to display:
Set madoReReview = Server.CreateObject("ADODB.Recordset")
Set madoReReviewElms = Server.CreateObject("ADODB.Recordset")
Set adCmd = GetAdoCmd("spReReviewGet")
    AddParmIn adCmd, "@ReReviewID", adInteger, 0, mlngReReviewID
    AddParmIn adCmd, "@UserID", adVarChar, 20, gstrUserID
    'Call ShowCmdParms(adCmd) '***DEBUG
    madoReReview.Open adCmd, , adOpenForwardOnly, adLockReadOnly

Set madoReReviewElms = madoReReview.NextRecordset
mstrReReviewElements = ""

If madoReReview.RecordCount = 1 Then
    mlngReviewID = madoReReview("rrvOrgReviewID").Value
Else
    mlngReviewID = 0
End If
%>

<HTML>
<HEAD>
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <STYLE id=ThisPageStyles type="text/css" rel="stylesheet">
        RADIO {position: absolute}
    </STYLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
        <STYLE type="text/css">
        .DivTab
            {
            TOP:125;    
            WIDTH:144; 
            HEIGHT:20; 
            FONT-WEIGHT:bold;
            TEXT-ALIGN:center;
            OVERFLOW:visible;
            BORDER-BOTTOM-STYLE:none;
            BACKGROUND-COLOR:<%=gstrAltBackColor%>;
            Z-INDEX:100;
            CURSOR:default
            }
    </STYLE>
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Dim mblnCloseClicked
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnMainClosed      <%'Flag used througout page to determine if main has been closed or not.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>
Dim mlngTimerIDS
Dim mdctAudit
Dim mdctHistory

Sub window_onload
    Dim intI
    Dim objOption
    Dim strElm
    Dim strTxt
    Dim intPos
    
    mblnMainClosed = False
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>

    mblnCloseClicked = False
    Call SizeAndCenterWindow(767, 520, True)

    If Trim(Form.UserID.Value) = "" Then Exit Sub
    'Fill the form with values from the record that is to be edited,
    'or default values for adding a new record:
    divFindReviewForReReview.style.left = -1000
    'divFindReviewForReReviewHdr.style.left = -1000
    divReReviewEntry.style.left = -2
    If "<%=mlngReReviewID%>" = "0" Or "<%=mlngReReviewID%>" = "" Then
        cmdChangeRecord.disabled = True
        cmdDeleteRecord.disabled = True
        cmdPrint.disabled = True
    Else
        cmdChangeRecord.disabled = False
        If InStr("<%=gstrRoles%>","[1]") > 0 Then
            cmdDeleteRecord.disabled = False
        Else
            cmdDeleteRecord.disabled = True
        End If
        cmdPrint.disabled = False
    End If
    cmdSaveRecord.disabled = True
    cmdCancelEdit.disabled = True
    If InStr("<%=gstrRoles%>","[1]") > 0 Then
        cmdAddRecord.disabled = False
    Else
        cmdAddRecord.disabled = True
    End If
    cmdFindRecord.disabled = False
    ' If re-reviewer is not the original re-reviewer or an admin, cannot edit
    If Form.rrvEvaluater.value = "<% = gstrUserName %>" _
        Or ("<% = gblnUserAdmin %>" = "True") _
        Or ("<% = gblnUserQA %>" = "True") Then
        ' Allow edit
    Else
        'If user is a sup, allow edit.
        If InStr("<%=gstrRoles%>","[2]") > 0 And Form.rrvRvwSig.value = "N" Then
            cmdChangeRecord.disabled = False
        Else
            cmdChangeRecord.disabled = True
        End If
        cmdDeleteRecord.disabled = True
    End If

    Call LoadAuditDictionary(<%=mlngReReviewID%>,<%=mlngReviewID%>)
    Call FillScreen
    <%'Fill ReReview Element portion of screen with a blank record%>
    Call WriteReReviewElements(Form.ReReviewElements.value)

    'Initialy show the form with controls disabled:
    Call DisableControls(True)
    If txtEvaluationDate.value = "" Then txtEvaluationDate.value = Date
    Call divTabs_onclick(1)
    divCaseBody.style.visibility = "visible"
End Sub

<%'If timer detects that Main has been closed, this sub will be called.  If window is
  'currently not in Edit mode, simply close the window.  If window is in Edit mode,
  'do not close window, but set the mblnMainClosed flag.  This flag will cause the
  'window to be closed at the next available opportunity. %>
Sub MainClosed()
    mblnMainClosed = True
    mblnSetFocusToMain = False
    If cmdSaveRecord.disabled = True Then
        mblnCloseClicked = True
        window.close
    End If
End Sub

Sub window_onbeforeunload
    If Not mblnCloseClicked Then
        If Form.FormAction.value <> "" Then
            window.event.returnValue = "Closing the browser window will exit the application without saving." & space(10) & vbCrLf & "Please use the <Save> button to save your changes, then use" & space(10) & vbcrlf & "the <Close> button to return to the main menu." & space(10)
        End If
    End If
    If mblnSetFocusToMain = True And mblnMainClosed = False Then
        window.opener.focus
    End If
End Sub

Sub ShowDivs(strAction)
    lblFormTitle.innerText = "Case Review System ~ Enter <%=mstrReReviewType%>"
    Select Case strAction
        Case "Find"
            divFindReviewForReReview.style.left = -2
            divReReviewEntry.style.left = -1000
            txtEvaluater.style.left = -1000
            lblFormTitle.innerText = "Find Review For <%=mstrReReviewType%>"
        Case "AddRecord"
            Form.FormAction.value = "AddRecord"
            divFindReviewForReReview.style.left = -1000
            divReReviewEntry.style.left = -2
            txtEvaluater.style.left = 473
            txtEvaluationID.value = ""
            Call ClearScreen()
            txtEvaluationDate.value = Date
            cmdChangeRecord.disabled = True
            cmdSaveRecord.disabled = False
            cmdDeleteRecord.disabled = True
            cmdCancelEdit.disabled = False
            cmdPrint.disabled = False
            cmdAddRecord.disabled = True
            cmdFindRecord.disabled = True
            <%' For an ADD, default all reviewed programs to re-reviewed'
            '   to change this to default all reviewed programs to NOT be rereviewed, remove next line.%>
            Form.ProgramsReReviewed.value = Form.ProgramsReviewed.value
            Call WriteReReviewElements(Form.ReReviewElementsEdit.value)
            Call DisableControls(False)
            'Call cboResponse_onchange
        Case "Cancel"
            divFindReviewForReReview.style.left = -1000
            divReReviewEntry.style.left = -2
            txtEvaluater.style.left = 473
    End Select
End Sub

Sub ProgramRevLbl_OnClick(intProgramID)
    If document.all("chkProgramRev" & intProgramID).disabled = True Then Exit Sub
    
    document.all("chkProgramRev" & intProgramID).checked = Not document.all("chkProgramRev" & intProgramID).checked
    Call ProgramRevCtl_OnClick(intProgramID)
End Sub

Sub ProgramRevCtl_OnClick(intProgramID)
    If document.all("chkProgramRev" & intProgramID).checked = True Then
        If InStr(Form.ProgramsReReviewed.value,"[" & intProgramID & "]") = 0 Then
            Form.ProgramsReReviewed.value = Form.ProgramsReReviewed.value & "[" & intProgramID & "]"
        End If
    Else
        If InStr(Form.ProgramsReReviewed.value,"[" & intProgramID & "]") > 0 Then
            Form.ProgramsReReviewed.value = Replace(Form.ProgramsReReviewed.value,"[" & intProgramID & "]","")
        End If
    End If
    Call WriteReReviewElements(Form.ReReviewElementsEdit.value)
End Sub

Sub WriteReReviewElements(strRecords)
    Dim intI, intCtlCount
    Dim strHTML, strPrgHTML
    Dim intTop, intTop2, intPrgTop
    Dim intTabIndex
    Dim strRecord
    Dim strFactors
    Dim strCheckedC, strCheckedI
    Dim strProgram
    Dim intProgramID
    Dim strComments
    Dim strReviewType
    Dim strTabType, intTabID, intElmID, strScreen, intWidth
    Dim intExtra

    strHTML = ""
    strPrgHTML = "<SPAN id=lblProgramsReviewed class=DefLabel style=""TOP:5; LEFT:0; WIDTH:120;text-align:center""><B>Programs Reviewed</B></SPAN>"
    intTop = 0
    intTabIndex = 4
    strProgram = ""
    intPrgTop = 20
    'If strRecords = "" Then strRecords = "^^^^^^^^^^^^|"
    intCtlCount = 0
    For intI = 1 To 1000
        strRecord = Parse(strRecords,"|",intI)
        If strRecord = "" Then Exit For
        intTabID = Parse(strRecord,"^",13)
        intElmID = Parse(strRecord,"^",9)
        If Parse(strRecord,"^",6) <> strProgram Then
            strProgram = Parse(strRecord,"^",6)
            If Parse(strRecord,"^",12) <> "" Then
                strReviewType = " -- " & Parse(strRecord,"^",12)
            Else
                strReviewType = ""
            End If
            intProgramID = Parse(strRecord,"^",5)
            strChecked = ""
            If strProgram <> "" And (InStr(Form.ProgramsReReviewed.value,"[" & intProgramID & "]") > 0) Then
                If intI > 1 Then
                    intTop = intTop + 30
                End If
                strHTML = strHTML & "<SPAN class=DefLabel id=lblProgramHeader" & intProgramID
                strHTML = strHTML & " style=""LEFT:1;WIDTH:700;TOP:" & intTop & ";font-size:14""><B>" & strProgram & strReviewType & "</B></SPAN>"
                intTop = intTop + 20
                strChecked = " checked "            
            End If
            strPrgHTML = strPrgHTML & "<INPUT type=checkbox ID=chkProgramRev" & intProgramID & " onclick=ProgramRevCtl_OnClick(" & intProgramID & ")" & strChecked & " NAME=chkProgram" & intProgramID & " style=""TOP:" & intPrgTop & "; LEFT:1"">"
            strPrgHTML = strPrgHTML & "<SPAN id=lblProgramRev" & intProgramID & " onclick=ProgramRevLbl_OnClick(" & intProgramID & ") class=DefLabel style=""TOP:" & intPrgTop+1 & ";LEFT:20;WIDTH:100;cursor:hand"">" & strProgram & "</SPAN>"
            intPrgTop = intPrgTop + 15
            strTabType = ""
        End If
        If strChecked <> "" Then
            intCtlCount = intCtlCount + 1
            strHTML = strHTML & "<INPUT type=hidden ID=hidRowInfo" & intCtlCount & " value=""" & Parse(strRecord,"^",8) & "^" & intTabID & "^" & intElmID & "^" & Parse(strRecord,"^",14) & """>"
            If strTabType <> strProgram & intTabID Then
                ' Print Tab Title
                If intTabID = 3 Then intTop = intTop + 20
                strHTML = strHTML & "<SPAN class=DefLabel id=lblTabHeader" & intCtlCount
                strHTML = strHTML & " style=""LEFT:6;WIDTH:700;TOP:" & intTop & ";font-size:13""><B>" & GetTabName(intTabID) & "</B></SPAN>"
                intTop = intTop + 20
                strTabType = strProgram & intTabID
                strScreen = ""
            End If
            strCheckedC = ""
            strCheckedI = ""
            If Parse(strRecord,"^",10) = "" Then
                strCheckedC = "checked"
            Else
                If Parse(strRecord,"^",10) = "22" Or Parse(strRecord,"^",10) = "0" Then
                    strCheckedC = "checked"
                Else
                    strCheckedI = "checked"
                End If
            End If
            If intTabID = 1 Or intTabID = 3 Then
                strHTML = strHTML & "<DIV id=divElement" & intCtlCount
                strHTML = strHTML & " style=""LEFT:10;WIDTH:700;TOP:" & intTop & ";HEIGHT:120;BACKGROUND-COLOR:<%=gstrAltBackColor%>"""
                strHTML = strHTML & " tabIndex=-1>"
                intTop2 = 0
                strHTML = strHTML & "<SPAN class=DefLabel id=lblAIElm" & intElmID
                strHTML = strHTML & " style=""LEFT:6;WIDTH:700;TOP:" & intTop2 & ";font-size:12""> " & Parse(strRecord,"^",2) & " -- " & Parse(strRecord,"^",3) & "</B></SPAN>"
                intTop2 = intTop2 + 15
                strHTML = strHTML & "<SPAN id=lblReviewComments" & intCtlCount & " class=DefLabel"
                strHTML = strHTML & " style=""TOP:" & intTop2 & "; LEFT:16; WIDTH:700""><B>Review Comments:</B></SPAN>"
                intTop2 = intTop2 + 15
                strHTML = strHTML & "<TEXTAREA id=txtReviewComments" & intCtlCount & "C"
                strHTML = strHTML & " readonly style=""LEFT:16;WIDTH:680;TOP:" & intTop2 & ";HEIGHT:40;border:solid 1;background-color:#CCCC99;overflow:auto;FONT-SIZE:10pt;FONT-FAMILY: tahoma;"" tabIndex=-1 >"
                strComments = Replace(Parse(strRecord,"^",7),"[linebreak]",vbCrLf)
                strComments = CleanTextRecordParsers(CleanText(strComments,"FromDb"),"FromDb")
                strHTML = strHTML & strComments
                strHTML = strHTML & "</TEXTAREA>"
                intTop2 = intTop2 + 40
                strHTML = strHTML & "<SPAN class=DefLabel id=lblAIElm" & intElmID
                If "<%=mlngReReviewTypeID%>" = "0" Then
                    intWidth = 200
                Else
                    intWidth = 250
                End If
                strHTML = strHTML & " style=""LEFT:16;WIDTH:" & intWidth & ";TOP:" & intTop2 & ";font-size:12""><B><%=mstrReReviewType%> Status:&nbsp;&nbsp;</B> Accurate</SPAN>"
                strHTML = strHTML & "<INPUT type=radio " & strCheckedC & " name=""ReReviewStatus" & intCtlCount & """ id=optReReviewStatus" & intCtlCount & "C"
                strHTML = strHTML & " style=""top:" & intTop2 & ";left:" & intWidth-2 & """>"
                strHTML = strHTML & "<SPAN id=lblReReviewStatusInCorrect" & intCtlCount & " class=DefLabel"
                strHTML = strHTML & " style=""TOP:" & intTop2 & "; LEFT:" & intWidth+30 & "; WIDTH:70;font-size:12"">Inaccurate</SPAN>"
                strHTML = strHTML & "<INPUT type=radio " & strCheckedI & " name=""ReReviewStatus" & intCtlCount & """ id=optReReviewStatus" & intCtlCount & "I"
                strHTML = strHTML & " style=""top:" & intTop2 & ";left:" & intWidth+90 & """ tabIndex=" & intTabIndex & " >"
                intTabIndex = intTabIndex + 1
                intTop2 = intTop2 + 15
                strHTML = strHTML & "<SPAN id=lblReReviewComments" & intCtlCount & " class=DefLabel"
                strHTML = strHTML & " style=""TOP:" & intTop2 & "; LEFT:16; WIDTH:700""><B><%=mstrReReviewType%> Comments:</B></SPAN>"
                strHTML = strHTML & "<TEXTAREA id=txtReReviewComments" & intCtlCount & "C"
                strHTML = strHTML & " style=""LEFT:16;WIDTH:680;TOP:" & intTop2+15 & ";HEIGHT:40;border:solid 1;background-color:#ffffff;overflow:auto;FONT-SIZE:10pt;FONT-FAMILY: tahoma;"" tabIndex=" & intTabIndex & " >"
                strComments = Replace(Parse(strRecord,"^",11),"[linebreak]",vbCrLf)
                strComments = CleanTextRecordParsers(CleanText(strComments,"FromDb"),"FromDb")
                strHTML = strHTML & strComments
                strHTML = strHTML & "</TEXTAREA></DIV>"
                intTabIndex = intTabIndex + 1
                intTop = intTop + 145
            End If
            If intTabID = 2 Then
                strHTML = strHTML & "<DIV id=divElement" & intCtlCount
                strHTML = strHTML & " style=""LEFT:10;WIDTH:700;TOP:" & intTop & ";HEIGHT:80;BACKGROUND-COLOR:<%=gstrAltBackColor%>"""
                strHTML = strHTML & " tabIndex=-1>"
                intTop2 = 0
                intExtra = 0
                If strScreen <> Parse(strRecord,"^",2) Then
                    strHTML = strHTML & "<SPAN class=DefLabel id=lblAIElm" & intElmID
                    If Parse(strRecord,"^",5) = 6 Then
                        strHTML = strHTML & " style=""LEFT:6;WIDTH:700;TOP:" & intTop2 & ";font-size:12"">" & Parse(strRecord,"^",1) & " -- " & Parse(strRecord,"^",2) & "</SPAN>"
                    Else
                        strHTML = strHTML & " style=""LEFT:6;WIDTH:700;TOP:" & intTop2 & ";font-size:12"">" & Parse(strRecord,"^",2) & "</SPAN>"
                    End If
                    intTop2 = intTop2 + 15
                    strHTML = strHTML & "<SPAN id=lblReviewComments" & intCtlCount & " class=DefLabel"
                    strHTML = strHTML & " style=""TOP:" & intTop2 & "; LEFT:16; WIDTH:700""><B>Review Comments:</B></SPAN>"
                    intTop2 = intTop2 + 15
                    strHTML = strHTML & "<TEXTAREA id=txtReviewComments" & intCtlCount & "C"
                    strHTML = strHTML & " readonly style=""LEFT:16;WIDTH:680;TOP:" & intTop2 & ";HEIGHT:40;border:solid 1;background-color:#CCCC99;overflow:auto;FONT-SIZE:10pt;FONT-FAMILY: tahoma;"" tabIndex=-1 >"
                    strComments = Replace(Parse(strRecord,"^",7),"[linebreak]",vbCrLf)
                    strComments = CleanTextRecordParsers(CleanText(strComments,"FromDb"),"FromDb")
                    strHTML = strHTML & strComments
                    strHTML = strHTML & "</TEXTAREA>"
                    intTop2 = intTop2 + 40
                    intTop = intTop + 15
                    strScreen = Parse(strRecord,"^",2)
                    intExtra = 55
                End If
                strHTML = strHTML & "<SPAN class=DefLabel id=lblAIElm" & intElmID
                strHTML = strHTML & " style=""LEFT:16;WIDTH:700;TOP:" & intTop2 & ";font-size:12"">" & Parse(strRecord,"^",4) & " -- " & Parse(strRecord,"^",3) & "</SPAN>"
                intTop2 = intTop2 + 15
                strHTML = strHTML & "<SPAN class=DefLabel id=lblAIElm" & intElmID
                If "<%=mlngReReviewTypeID%>" = "0" Then
                    intWidth = 200
                Else
                    intWidth = 250
                End If
                strHTML = strHTML & " style=""LEFT:16;WIDTH:" & intWidth & ";TOP:" & intTop2 & ";font-size:12""><B><%=mstrReReviewType%> Status:&nbsp;&nbsp;</B> Accurate</SPAN>"
                strHTML = strHTML & "<INPUT type=radio " & strCheckedC & " name=""ReReviewStatus" & intCtlCount & """ id=optReReviewStatus" & intCtlCount & "C"
                strHTML = strHTML & " style=""top:" & intTop2 & ";left:" & intWidth-2 & """>"
                strHTML = strHTML & "<SPAN id=lblReReviewStatusInCorrect" & intCtlCount & " class=DefLabel"
                strHTML = strHTML & " style=""TOP:" & intTop2 & "; LEFT:" & intWidth+30 & "; WIDTH:70;font-size:12"">Inaccurate</SPAN>"
                strHTML = strHTML & "<INPUT type=radio " & strCheckedI & " name=""ReReviewStatus" & intCtlCount & """ id=optReReviewStatus" & intCtlCount & "I"
                strHTML = strHTML & " style=""top:" & intTop2 & ";left:" & intWidth+90 & """ tabIndex=" & intTabIndex & " >"
                intTabIndex = intTabIndex + 1
                intTop2 = intTop2 + 15
                strHTML = strHTML & "<SPAN id=lblReReviewComments" & intCtlCount & " class=DefLabel"
                strHTML = strHTML & " style=""TOP:" & intTop2 & "; LEFT:16; WIDTH:700""><B><%=mstrReReviewType%> Comments:</B></SPAN>"
                strHTML = strHTML & "<TEXTAREA id=txtReReviewComments" & intCtlCount & "C"
                strHTML = strHTML & " style=""LEFT:16;WIDTH:680;TOP:" & intTop2+15 & ";HEIGHT:40;border:solid 1;background-color:#ffffff;overflow:auto;FONT-SIZE:10pt;FONT-FAMILY: tahoma;"" tabIndex=" & intTabIndex & " >"
                strComments = Replace(Parse(strRecord,"^",11),"[linebreak]",vbCrLf)
                strComments = CleanTextRecordParsers(CleanText(strComments,"FromDb"),"FromDb")
                strHTML = strHTML & strComments
                strHTML = strHTML & "</TEXTAREA></DIV>"
                intTop = intTop + 90 + intExtra
            End If
        End If
    Next

    'Write a hidden control that contains the total number of re-reveiw controls
    strHTML = strHTML & "<INPUT Type=hidden id=hidTotalCtls value=" & intCtlCount & ">"
    divTab1.innerHTML = strHTML
    divPrograms.innerHTML = strPrgHTML
End Sub

Function GetTabName(intTabID)
    Select Case intTabID
        Case 1
            GetTabName = "Action Integrity"
        Case 2
            GetTabName = "Data Integrity"
        Case 3
            GetTabName = "Information Gathering"
    End Select
End Function

Sub cmdPrint_onclick
    Call PrintReReview(window.txtEvaluationID.value, True)
End Sub

Sub PrintReReview(lngReReviewID, blnCurrentReReview)
    Dim strReturnValue
    
    If blnCurrentReReview Then
        If Not cmdSaveRecord.disabled Then 
            If InStr(Form.FormAction.value, "Print") = 0 Then
                Form.FormAction.value = Form.FormAction.value & "Print"
            End If
            Call cmdSaveRecord_onclick
            Exit Sub
        End If
        cmdPrint.disabled = True
    End If
    <%'Open the print-preview window, passing it the review ID:%>
    strReturnValue = window.showModalDialog("PrintReReview.asp?AuditRead=False&ReReviewID=" & lngReReviewID & _
        "&UserID=<%=gstrUserID%>&PageTitle=<%=mstrPageTitle%>", , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
    cmdPrint.disabled = False
    If mblnMainClosed = True Then 
        Call CloseWindow("<%=mstrReReviewType%>",True,1)
        Exit Sub
    End If
    If txtEvaluationID.value <> "" Then
        If CLng(lngReReviewID) = CLng(txtEvaluationID.value) Then
            Call LoadAuditDictionary(lngReReviewID,Form.rrvorgReviewID.value)
            Call DisplayAuditActivity()
        End If
    End If
End Sub

Sub cmdCancelEdit_onclick
    Dim intResponse
    Dim strMessage
    
    If Form.FormAction.value = "AddRecord" Then
        strMessage = "This <%=mstrReReviewType%> has not been saved." & space(10) & vbCrlf & vbCrlF & "Are you sure you wish to Cancel?" 
    Else
        strMessage = "Any changes on this <%=mstrReReviewType%> have not been saved." & space(10) & vbCrlf & vbCrlF & "Are you sure you wish to Cancel?" 
    End If
    intResponse = MsgBox(strMessage, vbQuestion + vbYesNo, "Cancel")
    If mblnMainClosed = True Then 
        Call CloseWindow("<%=mstrReReviewType%>",True,1)
        Exit Sub
    End If
    If intResponse = vbNo Then
        Exit Sub
    End If
    Form.FormAction.Value = ""
    Call LoadAuditDictionary(Form.rrvID.value,Form.rrvorgReviewID.value)
    Call Fillscreen
    'If Len(txtEvaluationID.Value) = 0 Or txtEvaluationID.Value = 0 Then
    '    Form.ProgramsReviewed.value = ""
    'End If
    Call WriteReReviewElements(Form.ReReviewElements.value)
    Call DisableControls(True)
    cmdFindRecord.disabled = False

    If InStr("<%=gstrRoles%>","[1]") > 0 Then
        cmdAddRecord.disabled = False
    Else
        cmdAddRecord.disabled = True
    End If
    If IsNumeric(txtEvaluationID.Value) Then
        If chkSubmit.checked Then
            If <%=gblnUserAdmin%> Or <%=gblnUserQA%> Then
                cmdChangeRecord.disabled = False
                cmdDeleteRecord.disabled = False
            Else
                cmdChangeRecord.disabled = True
                cmdDeleteRecord.disabled = True
            End If
        Else
            cmdChangeRecord.disabled = False
            If InStr("<%=gstrRoles%>","[1]") > 0 Then
                cmdDeleteRecord.disabled = False
            Else
                cmdDeleteRecord.disabled = True
            End If
        End If
    Else
        ' Canceled an Add
        cmdChangeRecord.disabled = True
        cmdSaveRecord.disabled = True
        cmdDeleteRecord.disabled = True
        cmdCancelEdit.disabled = True
        cmdPrint.disabled = True
        If InStr("<%=gstrRoles%>","[1]") > 0 Then
            cmdAddRecord.disabled = False
        Else
            cmdAddRecord.disabled = True
        End If
        cmdFindRecord.disabled = False
        Call ClearOriginalLabels()
    End If
    cmdSaveRecord.disabled = true
    cmdCancelEdit.disabled = true  
    
    'Fill in the review date with the current date:
    If Trim(txtEvaluationDate.value) = "" Then
        txtEvaluationDate.value = Date
    End If
    
    'Put the user in the first control:
    If InStr("<%=gstrRoles%>","[1]") > 0 Then
        cmdAddRecord.focus
    Else
        cmdChangeRecord.focus
    End If
End Sub

Sub ClearOriginalLabels()
    lblCaseIDValue.innerText = ""
    lblReviewMonthValue.innerText = ""
    lblReviewDateValue.innerText = ""
    lblReviewClassValue.innerText = ""
    lblCaseNameValue.innerText = ""
    lblCaseNumberValue.innerText = ""
    lblReviewStatusValue.innerText = ""
    lblReviewerNameValue.innerText = ""
    lblWorkerNameValue.innerText = ""
    lblWorkerResponseValue.innerText = ""
    divPrograms.innerHTML = "<SPAN id=lblProgramsReviewed class=DefLabel style=""TOP:5; LEFT:0; WIDTH:120;text-align:center""><B>Programs Reviewed</B></SPAN>"
End Sub

Sub cmdFindRecord_onclick
    If mblnMainClosed = True Then 
        Call CloseWindow("Re Review",True,<%=mlngWindowID%>)
        Exit Sub
    End If
    window.opener.Form.CalledFrom.Value = "ReReviewAddEdit.asp"
    window.opener.Form.action = "FindReReview.asp"
    window.opener.Form.ReReviewTypeID.value = <%=mlngReReviewTypeID%>
    Call window.opener.ManageWindows(<%=mlngWindowID+1%>,"Open")
End Sub

Sub cmdAddRecord_onclick
    Call ShowDivs("Find")
End Sub

Sub PageBody_ondblclick
    If cmdChangeRecord.disabled = false Then
        If chkSubmit.checked Then
            If <%=gblnUserAdmin%> Or <%=gblnUserQA%> Then
                Call cmdChangeRecord_onclick
            End If
        Else
            Call cmdChangeRecord_onclick
        End If
    End If
End Sub

Sub cmdChangeRecord_onclick
    Call DisableControls(False)
    cmdAddRecord.disabled = True
    cmdFindRecord.disabled = True
    cmdDeleteRecord.disabled = True
    cmdChangeRecord.disabled = True
    cmdSaveRecord.disabled = False
    cmdCancelEdit.disabled = False  
    Form.ReReviewElementsEdit.value = Form.ReReviewElements.value
    Form.FormAction.value = "ChangeRecord"
    'Call cboResponse_onchange
End Sub

Sub cmdDeleteRecord_onclick()
    Dim intResp
    
    intResp = MsgBox("Delete the current record?", vbQuestion + vbYesNo, "Delete")
    If mblnMainClosed = True Then 
        Call CloseWindow("<%=mstrReReviewType%>",True,1)
        Exit Sub
    End If
    If intResp = vbYes Then
        Form.SaveCompleted.Value = "N"
        mlngTimerIDS = window.setInterval("CheckForCompletion",100)
        Form.FormAction.value = "DeleteRecord"
        Form.action = "ReReviewAddEditSave.asp"
        Form.Target = "SaveFrame"
        SaveWindow.style.left = 1
        divCaseBody.style.left = -1000
        Form.Submit
    End If
    cmdPrint.disabled = True
    If mblnMainClosed = True Then 
        Call CloseWindow("Re Review",True,<%=mlngWindowID%>)
        Exit Sub
    End If
End Sub

Sub cmdSaveRecord_onclick
    Dim intElm
    Dim intPrg
    Dim intFct
    Dim strControl
    Dim blnErrFound
    Dim blnCaseErrFound
    Dim blnIsPrgReviewed
    Dim blnIsCasReviewed
    Dim strMsg
    Dim blnValidationFailed
    Dim blnBenErrFound
    
    blnValidationFailed = False
    strMsg = "The following items must be resolved before the review can be saved:" & space(10) & vbCrLf

    If Len(Form.ProgramsReReviewed.value) < 3 Then
        <%'Form.ProgramsReReviewed.value stores the program checkboxes that have been checked.
          'If the length is less than 3 (Min=[1]), no programs were selected to be re-reviewed.%>
        strMsg = strMsg & vbCrLf & space(4) & "<%=mstrReReviewType%> Entry: At least 1 program must be selected." & space(10)
        If Not blnValidationFailed Then
            blnValidationFailed = True
        End If
    End If

    If blnValidationFailed Then
        MsgBox strMsg, vbInformation, "Save"
        If mblnMainClosed = True Then 
            Call CloseWindow("<%=mstrReReviewType%>",True,1)
            Exit Sub
        End If
        Exit Sub
    End If
    
    If chkSubmit.checked Then
        strMsg = "The following items must be resolved before the <%=mstrReReviewType%> can be submitted:" & space(10) & vbCrLf

        'If cboResponse.value = 0 Then
        '    strMsg = strMsg & vbCrLf & space(4) & "Response:  The Response must be selected." & space(10)
        ''    If Not blnValidationFailed Then
        '        cboResponse.focus
        '        blnValidationFailed = True
        '    End If
        'End If
        'If cboResponse.options(cboResponse.selectedIndex).text = "Required" Then
        '    If txtResponseDue.value = vbNullString Then
        '        strMsg = strMsg & vbCrLf & space(4) & "Response Due Date:  The month, day and year (MM/DD/YYYY) of the Response Due Date must be entered." & space(10)
        '        If Not blnValidationFailed Then
        '            txtResponseDue.focus
        '            blnValidationFailed = True
        '        End If
        '    End If
        'End If
        If blnValidationFailed Then
            MsgBox strMsg, vbInformation, "Save - Submit <%=mstrReReviewType%>"
            If mblnMainClosed = True Then 
                Call CloseWindow("<%=mstrReReviewType%>",True,1)
                Exit Sub
            End If
            Exit Sub
        End If

    End If
    Call FillForm()

    Form.SaveCompleted.Value = "N"
    mlngTimerIDS = window.setInterval("CheckForCompletion",100)
    mblnCloseClicked = True
    Form.action = "ReReviewAddEditSave.asp"
    Form.Target = "SaveFrame"
    SaveWindow.style.left = 1
    divCaseBody.style.left = -1000
    Form.Submit
    
    <%' If Main has been closed, do not allow window to remain open unless Save was called from print button.%>
    If InStr(Form.FormAction.value,"Print") = 0 Then
        If mblnMainClosed = True Then
            Call CloseWindow("<%=mstrReReviewType%>",True,2)
            Exit Sub
        End If
    End If
End Sub

Function CheckForCompletion()
    If Form.SaveCompleted.value = "Y" Then
        window.clearInterval mlngTimerIDS
        
        Call DisableControls(True)

        txtEvaluationID.value = Form.rrvID.value
        If Form.FormAction.value = "DeleteRecord" Then
            Call ClearScreen()
            Call ClearOriginalLabels()
            Form.ProgramsReviewed.value=""
            Form.ProgramsReviewedValue.value=""
            Form.ProgramsReReviewed.value=""
            Form.ProgramsReReviewedValue.value=""
            Form.ReReviewElements.value = ""
            Form.ReReviewElementsEdit.value = ""
            Call WriteReReviewElements(Form.ReReviewElements.value)
            cmdChangeRecord.disabled = True
            cmdDeleteRecord.disabled = True
        Else
            If InStr("<%=gstrRoles%>","[1]") = 0 And InStr("<%=gstrRoles%>","[2]") > 0 And chkRvwSig.checked = True Then
                cmdChangeRecord.disabled = True
            Else
                cmdChangeRecord.disabled = False
            End If
            If InStr("<%=gstrRoles%>","[1]") > 0 Then
                cmdDeleteRecord.disabled = False
            Else
                cmdDeleteRecord.disabled = True
            End If
        End If

        If InStr("<%=gstrRoles%>","[1]") > 0 Then
            cmdAddRecord.disabled = False
        Else
            cmdAddRecord.disabled = True
        End If
        cmdCancelEdit.disabled = True
        cmdSaveRecord.disabled = True
        cmdFindRecord.disabled = False
        Call LoadAuditDictionary(txtEvaluationID.value,Form.rrvorgReviewID.value)
        Call DisplayAuditActivity()
        If Form.ReReviewTypeID.value = 0 Then
            Call window.opener.LoadReviewList("REREVIEWADDEDIT")
        Else
            Call window.opener.LoadReviewList("CARREREVIEWADDEDIT")
        End If
        If InStr(Form.FormAction.value,"Print") > 0 Then
            <%'When the user clicks the Print button on the review entry screen,
            'the form is constructed to save the record first.  After posting
            'back from the save, the form will need to finish the process by 
            'recalling the Print button event code:%>

            'Insert a call to the print button event procedure:
             Call cmdPrint_onclick()
        End If
        Form.FormAction.value = ""
    End If
End Function

Sub FillForm()
    Dim intI, intJ, intCtlID
    Dim strUpdateString
    Dim blnStatusError
    Dim strReReviewElements
    Dim strRecord
    Dim strCleanComments, strBefore
    
    strUpdateString = ""
    blnStatusError = False
    
    If InStr(Form.FormAction.Value,"AddRecord") > 0 Then
        Form.rrvEvaluater.value = "<%=gstrUserName%>"
        Form.rrvDateEntered.value = txtEvaluationDate.value
    End If

    'If cboResponse.value <> Form.rrvResponseID.value Then    
    '    For intI = 0 To cboResponse.options.length - 1
    '        If CLng(Form.rrvResponseID.value) = CLng(cboResponse.options(intI).value) Then
    '            strBefore = cboResponse.options(intI).text
    '            Exit For
    '        End If
    '    Next
    '    Form.rrvResponseID.value = cboResponse.value
    '    strUpdateString = strUpdateString & "Response^" & strBefore & "^" & cboResponse.options(cboResponse.selectedIndex).text & "|"
    'End If
    'If txtResponseDue.value <> Form.rrvResponseDue.value Then    
    '    strUpdateString = strUpdateString & "Response Due^" & Form.rrvResponseDue.value & "^" & txtResponseDue.value & "|"
    '    Form.rrvResponseDue.value = txtResponseDue.value
    'End If
        
    If chkSubmit.checked Then
        If Form.rrvSubmitted.value <> "Y" Then
            Form.rrvSubmitted.value = "Y"
            If InStr(Form.FormAction.Value,"AddRecord") = 0 Then
                strUpdateString = strUpdateString & "Submitted^No^Yes|"
            End If
        End If
    Else
        If Form.rrvSubmitted.value <> "N" Then
            Form.rrvSubmitted.value = "N"
            If InStr(Form.FormAction.Value,"AddRecord") = 0 Then
                strUpdateString = strUpdateString & "Submitted^Yes^No|"
            End If
        End If
    End If
    If chkRrvSig.checked Then
        If Form.rrvRrvSig.value <> "Y" Then
            Form.rrvRrvSig.value = "Y"
            If InStr(Form.FormAction.Value,"AddRecord") = 0 Then
                strUpdateString = strUpdateString & "Re-Reviewer Signature^No^Yes|"
            End If
        End If
    Else
        If Form.rrvRrvSig.value <> "N" Then
            Form.rrvRrvSig.value = "N"
            If InStr(Form.FormAction.Value,"AddRecord") = 0 Then
                strUpdateString = strUpdateString & "Re-Reviewer Signature^Yes^No|"
            End If
        End If
    End If
    If chkRvwSig.checked Then
        If Form.rrvRvwSig.value <> "Y" Then
            Form.rrvRvwSig.value = "Y"
            If InStr(Form.FormAction.Value,"AddRecord") = 0 Then
                strUpdateString = strUpdateString & "Reviewer Signature^No^Yes|"
            End If
        End If
    Else
        If Form.rrvRvwSig.value <> "N" Then
            Form.rrvRvwSig.value = "N"
            If InStr(Form.FormAction.Value,"AddRecord") = 0 Then
                strUpdateString = strUpdateString & "Reviewer Signature^Yes^No|"
            End If
        End If
    End If
    Form.UpdateString.value = strUpdateString
    ' Update Form values of lables to keep for CANCEL purposes
    Form.ReviewMonthValue.value = lblReviewMonthValue.innerText
    Form.ReviewDateValue.value = lblReviewDateValue.innerText
    Form.ReviewClassValue.value = lblReviewClassValue.innerText
    Form.CaseNameValue.value = lblCaseNameValue.innerText
    Form.CaseNumberValue.value = lblCaseNumberValue.innerText
    Form.ReviewStatusValue.value = lblReviewStatusValue.innerText
    Form.ReviewerNameValue.value = lblReviewerNameValue.innerText
    Form.WorkerNameValue.value = lblWorkerNameValue.innerText
    Form.WorkerResponseValue.value = lblWorkerResponseValue.innerText
    Form.ProgramsReReviewedValue.value = Form.ProgramsReReviewed.value
    Form.ProgramsReviewedValue.value = Form.ProgramsReviewed.value
   
    Form.ReReviewElementsWrite.value = ""
    Form.ReReviewElements.value = ""
    strReReviewElements = Form.ReReviewElementsEdit.value
    intCtlID = 0
    For intI = 1 To 1000
        strRecord = Parse(strReReviewElements,"|",intI)
        If strRecord = "" Then Exit For
        If document.all("chkProgramRev" & Parse(strRecord,"^",5)).checked = True Then
            intCtlID = intCtlID + 1
            If document.all("optReReviewStatus" & intCtlID & "C").checked = True Then
                intStatusID = 22
            Else
                intStatusID = 23
                blnStatusError = True
            End If
            strCleanComments = CleanTextRecordParsers(CleanText(document.all("txtReReviewComments" & intCtlID & "C").value,"ToDb"),"ToDb")
            ' Build Form value that will be passed to Save page for database writes
            Form.ReReviewElementsWrite.value = Form.ReReviewElementsWrite.value & _
                document.all("hidRowInfo" & intCtlID).value & "^" & _
                intStatusID & "^" & _
                strCleanComments & "|"
            ' Rebuild Form value used to hold state on the page with values entered by user.  Only 
            ' re-review status (item 11) and re-review comments (item 12) could have changed.
            ' Rebuild first 9 items that could not have changed first
            For intJ = 1 To 9
                Form.ReReviewElements.value = Form.ReReviewElements.value & Parse(strRecord,"^",intJ) & "^"
            Next
            Form.ReReviewElements.value = Form.ReReviewElements.value & intStatusID & "^"
            Form.ReReviewElements.value = Form.ReReviewElements.value & strCleanComments & "^"
            Form.ReReviewElements.value = Form.ReReviewElements.value & Parse(strRecord,"^",12) & "^"
            Form.ReReviewElements.value = Form.ReReviewElements.value & Parse(strRecord,"^",13) & "^"
            Form.ReReviewElements.value = Form.ReReviewElements.value & Parse(strRecord,"^",14) & "|"
        Else
            'If function is not being re-reviewed, still add elements back to value in case the program is checked at a later time
            For intJ = 1 To 14
                Form.ReReviewElements.value = Form.ReReviewElements.value & Parse(strRecord,"^",intJ) & "^"
            Next
        End If
    Next
    If InStr(Form.FormAction.Value,"ChangeRecord") > 0 Then
        If Form.ReReviewElements.value = Form.ReReviewElementsEdit.value Then
            Form.ElementsChanged.value = "N"
        Else
            Form.ElementsChanged.value = "Y"
        End If
    End If
    If blnStatusError Then
        Form.rrvStatusId.Value = 23
    Else
        Form.rrvStatusID.Value = 22
    End If
End Sub

Sub cmdClose_onclick
    Dim intResp
    Dim blnClose
    
    If Form.FormAction.value <> "" And Form.FormAction.value <> "GetRecord" Then
        intResp = MsgBox("You are currently editing a record, are you sure you wish to close the edit form?", vbQuestionmark + vbYesNo, "Close Form")
        If intResp = vbYes Then
            mblnCloseClicked = True
            blnClose = True
        End If
    Else
        mblnCloseClicked = True
        blnClose = True
    End If

    If blnClose = True Then
        If mblnMainClosed = True Then
            <%' If main menu window is already closed, just close the window%>
            window.close
            Exit Sub
        End If
        Call window.opener.ManageWindows(<%=mlngWindowID%>,"Close")
    End If
End Sub

Sub txtResponseDue_onfocus
    If Trim(txtResponseDue.value) = "" Then
        txtResponseDue.value = "(MM/DD/YYYY)"
    End If
    txtResponseDue.select
End Sub

Sub txtResponseDue_onkeypress
    If txtResponseDue.value = "(MM/DD/YYYY)" Then
        txtResponseDue.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub

Sub CheckDates(ctlDate)
    If IsDate(txtEvaluationDate.value) And IsDate(txtResponseDue.value) Then
        If CDate(txtResponseDue.value) < CDate(txtEvaluationDate.value) Then
            MsgBox "The Response Due Date must greater than or equal to the Re-Review Date.", vbInformation, "Re-Review Entry"
            ctlDate.value = ""
            ctlDate.focus
        Else
            ctlDate.value = FormatDateTime(ctlDate.value,2)
        End If
    End If
End Sub
Sub txtResponseDue_onblur
    If Trim(txtResponseDue.value) = "(MM/DD/YYYY)" Then
        txtResponseDue.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtResponseDue.value) Then
        MsgBox "The Response Due Date must be a valid date - MM/DD/YYYY.", vbInformation, "Re-Review Entry"
        txtResponseDue.focus
    Else
        Call CheckDates(txtResponseDue)
    End If
End Sub

'Sub cboResponse_onchange()
'    If cboResponse.options(cboResponse.selectedIndex).text = "Required" Then
'        txtResponseDue.disabled = False
'        txtResponseDue.style.backgroundColor = "<%=gstrCtrlBackColor%>"
'    Else
'        txtResponseDue.disabled = True
'        txtResponseDue.value = ""
'        txtResponseDue.style.backgroundColor = "<%=gstrBackColor%>"
'    End If
'End Sub

Sub FillScreen()
    If Form.rrvID.value = "0" Then 
        txtEvaluationID.value = ""
    Else
        txtEvaluationID.value = Form.rrvID.value
    End If
    txtEvaluationDate.value = Form.rrvDateEntered.value
    txtEvaluater.value = Form.rrvEvaluater.value
    cboResponse.value = Form.rrvResponseID.value
    txtResponseDue.value = Form.rrvResponseDue.value
    If Form.rrvSubmitted.value = "Y" Then
        chkSubmit.checked = True
    Else
        chkSubmit.checked = False
    End If
    If Form.rrvRrvSig.value = "Y" Then
        chkRrvSig.checked = True
    Else
        chkRrvSig.checked = False
    End If
    If Form.rrvRvwSig.value = "Y" Then
        chkRvwSig.checked = True
    Else
        chkRvwSig.checked = False
    End If
    
    If Form.rrvOrgReviewID.value <> "0" Then
        lblCaseIDValue.innerText = Form.rrvOrgReviewID.value
    End If
    lblReviewMonthValue.innerText = Form.ReviewMonthValue.value
    lblReviewDateValue.innerText = Form.ReviewDateValue.value 
    lblReviewClassValue.innerText = Form.ReviewClassValue.value
    lblCaseNameValue.innerText = Form.CaseNameValue.value
    lblCaseNumberValue.innerText = Form.CaseNumberValue.value
    lblReviewStatusValue.innerText = Form.ReviewStatusValue.value
    lblReviewerNameValue.innerText = Form.ReviewerNameValue.value
    lblWorkerNameValue.innerText = Form.WorkerNameValue.value
    lblWorkerResponseValue.innerText = Form.WorkerResponseValue.value
    Form.ProgramsReviewed.value = Form.ProgramsReviewedValue.value
    Form.ProgramsReReviewed.value = Form.ProgramsReReviewedValue.value
    Call DisplayAuditActivity()
End Sub

Sub ClearScreen()
    txtEvaluater.value = "<%=gstrUserName%>"
    txtEvaluationID.value = ""
    txtEvaluationDate.value = Date
    cboResponse.value = 0
    txtResponseDue.value = ""
    chkSubmit.checked = False
    chkRvwSig.checked = False
    chkRrvSig.checked = False

    If Form.FormAction.value = "AddRecord" Then
        Call LoadAuditDictionary(0,lblCaseIDValue.innerText)
    Else
        Call LoadAuditDictionary(0,Form.rrvorgReviewID.value)
    End If
    Call DisplayAuditActivity()
End Sub 
 
Sub DisableControls(blnVal)
    Dim strBackColor, strHeaderBackColor
    Dim intI
    Dim strProgram
    
    If blnVal Then
        strBackColor = "<%=gstrPageColor%>"
        strHeaderBackColor = "<%=gstrBackColor%>"
    Else
        strBackColor = "#ffffff"
        strHeaderBackColor = "<%=gstrCtrlBackColor%>"
    End If
    If blnVal = False Then
        If chkSubmit.checked = True Then
            chkRvwSig.disabled = True
            chkRvwSig.style.backgroundColor = "<%=gstrBackColor%>"
            chkRrvSig.disabled = True
            chkRrvSig.style.backgroundColor = "<%=gstrBackColor%>"
        Else
            chkRvwSig.disabled = blnVal
            chkRvwSig.style.backgroundColor = strHeaderBackColor
            If InStr("<%=gstrRoles%>","[1]") > 0 Then
                chkRrvSig.disabled = blnVal
                chkRrvSig.style.backgroundColor = strHeaderBackColor
            Else
                chkRrvSig.disabled = True
                chkRrvSig.style.backgroundColor = "<%=gstrBackColor%>"
            End If
        End If
    Else
        chkRvwSig.disabled = blnVal
        chkRvwSig.style.backgroundColor = strHeaderBackColor
        chkRrvSig.disabled = blnVal
        chkRrvSig.style.backgroundColor = strHeaderBackColor
    End If
    If InStr("<%=gstrRoles%>","[1]") = 0 Then
        blnVal = True
        strBackColor = "<%=gstrPageColor%>"
        strHeaderBackColor = "<%=gstrBackColor%>"
    End If
    chkSubmit.disabled = blnVal
    chkSubmit.style.backgroundColor = strHeaderBackColor
    txtResponseDue.disabled = blnVal
    txtResponseDue.style.backgroundColor = strHeaderBackColor
    cboResponse.disabled = blnVal
    cboResponse.style.backgroundColor = strHeaderBackColor

    If hidTotalCtls.value > 0 Then
        For intI = 1 To hidTotalCtls.value
            document.all("optReReviewStatus" & intI & "C").disabled = blnVal
            document.all("optReReviewStatus" & intI & "I").disabled = blnVal
            document.all("txtReReviewComments" & intI & "C").disabled = blnVal
            document.all("txtReReviewComments" & intI & "C").style.backgroundColor = strBackColor
        Next
    End If
    
    For intI = 2 To 50
        strProgram = Parse(Form.ProgramsReviewed.value,"[",intI)
        If strProgram = "" Then
            Exit For
        End If
        
        strProgram = Left(strProgram,Len(strProgram)-1)
        If CInt(strProgram) >= 50 Then
            strProgram = "6"
        End If
        document.all("chkProgramRev" & strProgram).disabled = blnVal
    Next
End Sub

Function CleanText(strText, strDir)
    <%'This function is used to replace quote and double-quote characters with 
    'tokens when sending to the database, and replace the tokens with the
    'correct characters when retrieving from the database. The tokens used 
    'are {TAB}#sq# for single-quote
    '    {TAB}#dq# for double-quote%>

    If IsNull(strText) Then
        CleanText = ""
    Else 'Apostrophe Or double quotes:
        If strDir = "FromDb" Then
            strText = Replace(strText, Chr(9) & "#sq#", "'")
            strText = Replace(strText, Chr(9) & "#dq#", """")
        ElseIf strDir = "ToDb" Then
            strText = Replace(strText, "'", Chr(9) & "#sq#")
            strText = Replace(strText, """", Chr(9) & "#dq#")
        End If
        CleanText = strText
    End If
End Function
Function CleanTextRecordParsers(strText, strDir)
    <%'This function is used to replace carrot and bar characters with 
    'tokens when sending to the database, and replace the tokens with the
    'correct characters when retrieving from the database. The tokens used 
    'are {TAB}#ca# for carrot
    '    {TAB}#ba# for bar%>

    If IsNull(strText) Then
        CleanTextRecordParsers = ""
    Else 'Apostrophe Or double quotes:
        If strDir = "FromDb" Then
            strText = Replace(strText, Chr(9) & "#ca#", "^")
            strText = Replace(strText, Chr(9) & "#ba#", "|")
        ElseIf strDir = "ToDb" Then
            strText = Replace(strText, "^", Chr(9) & "#ca#")
            strText = Replace(strText, "|", Chr(9) & "#ba#")
        End If
        CleanTextRecordParsers = strText
    End If
End Function

Function ReturnNumeric(strVal)
    Dim intPos
    Dim strTmp
    
    For intPos = 1 To Len(strVal)
        If IsNumeric(Mid(strVal, intPos, 1)) Then
            strTmp = strTmp & Mid(strVal, intPos, 1)
        End If
    Next
    
    ReturnNumeric = strTmp
End Function

Sub document_onkeydown
    If window.event.keyCode = 27 Then
        If cmdCancelEdit.disabled = False Then
            Call cmdCancelEdit_onclick
        End If
    End If
End Sub

Sub Gen_onkeydown()
    If cmdChangeRecord.disabled = false Then
        If chkSubmit.checked Then
            If <%=gblnUserAdmin%> Or <%=gblnUserQA%> Then
                Call cmdChangeRecord_onclick
            End If
        Else
            Call cmdChangeRecord_onclick
        End If
    End If
End Sub

Sub FillCbo(strOrgList)
    Dim intI
    Dim oOption
    Dim strTmpList
    
    For intI = 1 To strOrgList.options.Length - 1
        strTmpList = strOrgList.options(intI).Text
        If CheckDuplicate(Parse(strTmpList, "--", 1), strOrgList, intI) = -1 Then
            strTmpList = Parse(strOrgList.options(intI).Text, "--", 1)
        Else
            strTmpList = strOrgList.options(intI).Text
        End If
        strOrgList.options(intI).Text = strTmpList
    Next
End Sub

Function CheckDuplicate(strTest, strCboName, intPos)
    Dim intI
    Dim strTmp
    Dim blnTest
    Dim intStop
    
    If intPos = strCboName.options.Length - 1 Then
        intStop = intPos
    Else
        intStop = intPos + 1
    End If
    blnTest = False
    For intI = intPos - 1 to intStop
        strTmp = Parse(strCboName.options(intI).Text, "--", 1)
        IF blnTest And StrComp(Parse(strTmp, ",", 1), Parse(strTest, ",", 1)) = 0 _
            And StrComp(Parse(strTmp, ",", 2), Parse(strTest, ",", 2)) = 0 Then
            CheckDuplicate = 1
            Exit Function
        End If
        If StrComp(Parse(strTmp, ",", 1), Parse(strTest, ",", 1)) = 0 _
            And StrComp(Parse(strTmp, ",", 2), Parse(strTest, ",", 2)) = 0 Then
            blnTest = True
        Else 
            blnTest = False
        End If
    Next
    
    CheckDuplicate = -1
End Function

Sub lblEvaluationID_onmouseover()
'   txtEvaluationID.value = document.activeElement.id
End Sub

Sub divTabs_onclick(intTab)
    Dim intI
    
    For intI = 1 To 3
        If CInt(intI) = CInt(intTab) Then
            Document.all("divTabButton" & intI).style.borderBottomStyle = "none"
            Document.all("divTabButton" & intI).style.fontWeight = "bold"
            Document.all("divTab" & intI).style.left = -2
            Document.all("divTab" & intI).style.visibility = "visible"
        Else
            Document.all("divTabButton" & intI).style.borderBottomStyle = "solid"
            Document.all("divTabButton" & intI).style.fontWeight = "normal"
            Document.all("divTab" & intI).style.left = -2000
            Document.all("divTab" & intI).style.visibility = "hidden"
        End If
    Next
End Sub
Sub divTabs_onkeydown(intTab)
End Sub

Sub RebuildProgramList(strProgramsSelected)
    Form.ProgramsSelected.value = strProgramsSelected
    window.opener.Form.ProgramsSelected.value = strProgramsSelected
End Sub

Sub LoadAuditDictionary(lngReReviewID, lngReviewID)
    Dim strURL
    strURL = "ActivityAudit.asp?Action=Read&RecordID=" & lngReReviewID & "&Table=tblReReview"
    Set mdctAudit = CreateObject("Scripting.Dictionary")
    Set mdctAudit = window.showModalDialog(strURL)
    <%'To include only the type (ReReview or CAR) being edited, pass the ReviewTypeID.  Passing -1 includes all. %>
    strURL = "ReReviewHistory.asp?ReReviewTypeID=-1&ReReviewID=" & lngReReviewID & "&ReviewID=" & lngReviewID
    Set mdctHistory = CreateObject("Scripting.Dictionary")
    Set mdctHistory = window.showModalDialog(strURL)
End Sub

Sub DisplayAuditActivity()
    Dim oRecord
    Dim strRecord
    Dim strInnerHTML, strOuterHTML
    Dim intI
    Dim strChange, strChanges, strEntryName, strBefore, strAfter
    Dim lngReReviewID
    
    intI = InStr(tblAudit.outerHTML,"<TBODY")
    If intI > 0 Then
        strOuterHTML = Left(tblAudit.outerHTML,intI-1)
        tblAudit.outerHTML = strOuterHTML & " <TBODY id=tbdAudit></TBODY></TABLE>"
    End If

    strOuterHTML = tblAudit.outerHTML
    
    For Each oRecord In mdctAudit
        strRecord = mdctAudit(oRecord)
        strChanges = Parse(strRecord,"^",5) & "!"
        For intI = 1 To 1000
            strChange = Parse(strChanges,"!",intI)
            If strChange = "" Then Exit For
            If InStr(strChange,"*") > 0 Then
                strEntryName = Parse(strChange,"*",1)
                strBefore = Parse(strChange,"*",2)
                strAfter = Parse(strChange,"*",3)
            Else
                strEntryName = strChange
                strBefore = "&nbsp;"
                strAfter = "&nbsp;"
            End If
            strInnerHTML = strInnerHTML & "<TR id=tdrAudit" & oRecord & ">" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail id=tdcAuditC0" & oRecord & " style=""width:150;text-align:center"">" & Parse(strRecord,"^",2) & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail id=tdcAuditC1" & oRecord & " style=""width:140;text-align:center"">" & Parse(strRecord,"^",3) & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail id=tdcAuditC2" & oRecord & " style=""width:140;text-align:center"">" & strEntryName & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail id=tdcAuditC3" & oRecord & " style=""width:145;text-align:center"">" & strBefore & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail id=tdcAuditC4" & oRecord & " style=""width:145;text-align:center"">" & strAfter & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "</TR>" & vbCrLf
        Next
    Next
    tblAudit.outerHTML = Replace(strOuterHTML,"<TBODY id=tbdAudit></TBODY>","<TBODY id=tbdAudit>" & strInnerHTML & "</TBODY>")
    
    'History table
    intI = InStr(tblHistory.outerHTML,"<TBODY")
    If intI > 0 Then
        strOuterHTML = Left(tblHistory.outerHTML,intI-1)
        tblHistory.outerHTML = strOuterHTML & " <TBODY id=tbdHistory></TBODY></TABLE>"
    End If

    strOuterHTML = tblHistory.outerHTML
    
    intI = 0
    strInnerHTML = ""
    lngReReviewID = txtEvaluationID.value
    If lngReReviewID = "" Then lngReReviewID = "0"
    For Each oRecord In mdctHistory
        strRecord = mdctHistory(oRecord)
        If Parse(strRecord,"^",1) = "" Or Parse(strRecord,"^",1) = "0" Then
        ElseIf CLng(Parse(strRecord,"^",1)) = CLng(lngReReviewID) Then
        Else
            strInnerHTML = strInnerHTML & "<TR id=HistoryRow" & intI & ">" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail style=""text-align:center;cursor:hand;color:blue"" onclick=HistoryRowClick(" & Parse(strRecord,"^",1) & ")><B>" & Parse(strRecord,"^",1) & "</B></TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail style=""text-align:center;"">" & Parse(strRecord,"^",2) & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail style=""text-align:center;"">" & Parse(strRecord,"^",3) & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail style=""text-align:center;"">" & Parse(strRecord,"^",4) & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail style=""text-align:center;"">" & Parse(strRecord,"^",5) & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail style=""text-align:center;"">" & Parse(strRecord,"^",6) & "</TD>" & vbCrLf
            strInnerHTML = strInnerHTML & "</TR>" & vbCrLf
            intI = intI + 1
        End If
    Next
    tblHistory.outerHTML = Replace(strOuterHTML,"<TBODY id=tbdHistory></TBODY>","<TBODY id=tbdHistory>" & strInnerHTML & "</TBODY>")
End Sub

Sub HistoryRowClick(lngReReviewID)
    Call PrintReReview(lngReReviewID,False)
End Sub

Sub cmdPrintAudit_onclick()
    Dim strReturnValue
    If Form.rrvID.Value = "" Or Form.rrvID.Value = "0" Then Exit Sub
    strReturnValue = window.showModalDialog("RptAuditHistory.asp?ReReviewTypeID=<%=mlngReReviewTypeID%>&TableName=tblReReview&RecordID=" & Form.rrvID.Value,, "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")    
    If mblnMainClosed = True Then 
        Call CloseWindow("<%=mstrReReviewType%>",True,1)
        Exit Sub
    End If
End Sub

Sub SignatureClick(intRowID)
    Dim ctlCheckBox
    
    Select Case intRowID
        Case 1, 11
            Set ctlCheckBox = chkRrvSig
        Case 2, 12
            Set ctlCheckBox = chkRvwSig
        Case 3, 13
            Set ctlCheckBox = chkSubmit
    End Select
    If intRowID < 10 Then
        If ctlCheckBox.disabled = True Then Exit Sub
        ctlCheckBox.checked = Not ctlCheckBox.checked
    End If
    
    Select Case intRowID
        Case 1, 11
            'nothing
        Case 2, 12
        Case 3, 13
            If ctlCheckBox.checked = True Then
                chkRrvSig.checked = True
                chkRrvSig.disabled = True
                chkRvwSig.checked = True
                chkRvwSig.disabled = True
            Else
                chkRrvSig.disabled = False
                chkRvwSig.disabled = False
            End If
    End Select
End Sub
</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 topMargin=5 leftMargin=5 rightMargin=5>

    <DIV id=FormTitle
        style="FONT-WEIGHT: bold; 
                COLOR: <%=gstrTitleColor%>; 
                FONT-STYLE: normal; 
                HEIGHT: 30; 
                WIDTH: 745;
                padding-top: 2;
                BACKGROUND-COLOR: <%=gstrAltBackColor%>;
                FONT-FAMILY: <%=gstrTitleFont%>;
                FONT-SIZE: <%=gstrTitleFontSmallSize%>;
                TEXT-ALIGN: center; 
                BORDER-COLOR: <%=gstrBorderColor%>;
                BORDER-STYLE: solid; 
                BORDER-WIDTH: 2">
        <span id=lblFormTitle style="width:590;left:80"><B><%= mstrPageTitle & " ~ Enter " & mstrReReviewType%></B></span>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:13;width:75">
            Navigate
        </DIV>
    </DIV>
    <% Call WriteNavigateControls(4,0,Null) %>
    
    <DIV ID=SaveWindow Name=SaveWindow 
        style="position:absolute;
            TOP:29; 
            LEFT:-1000;
            WIDTH:754; 
            HEIGHT:453; 
            COLOR:black; 
            BORDER-STYLE:solid;
            BORDER-WIDTH:2;
            BORDER-COLOR:<%=gstrBorderColor%>;
            BACKGROUND-COLOR:<%=gstrBackColor%>">
        <IFRAME ID=SaveFrame Name=SaveFrame style="top:0;height:400;left:0;width:750" src="Blank.html">
        </IFRAME>
        <SPAN id=lblSavingMessage class=DefLabel style="top:110;height:45;left:0;width:750"><CENTER><BIG><B></BIG>Saving Case Review Record...</B></BIG></CENTER></SPAN>
    </DIV> 

    <DIV id=divCaseBody
        style="OVERFLOW: hidden;
            border-style: solid;
            border-width: 2;
            border-color: <%=gstrBorderColor%>;
            TOP: 36; 
            HEIGHT: 445; 
            WIDTH: 745; 
            COLOR: black; 
            BACKGROUND-COLOR: <%=gstrBackColor%>">
     
        <SPAN id=lblEvaluationID class=DefLabel style= "TOP:6; LEFT:5;WIDTH:100;text-align:right" >
            <%=mstrReReviewType & " ID"%>
        </SPAN>

        <INPUT id=txtEvaluationID type=text 
            style="TOP:6;LEFT:110; WIDTH:55; BACKGROUND-COLOR:<%=gstrAltBackColor%>"
            readOnly NAME="txtEvaluationID">
                
        <SPAN id=lblEvaluationDate class=DefLabel style="TOP:6; LEFT:168; WIDTH:120;text-align:right">
            <%=mstrReReviewType & "&nbspDate"%>
        </SPAN>
        
        <INPUT id=txtEvaluationDate type=text 
            style="TOP:6;LEFT:298; WIDTH:75; BACKGROUND-COLOR:<%=gstrAltBackColor%>"
            readOnly NAME="txtEvaluationDate">
        
        <SPAN id=lblEvaluater class=DefLabel style="TOP:6; LEFT:400; WIDTH:70">
            <%=gstrEvaTitle%> 
        </SPAN>

        <INPUT id=txtEvaluater type=text 
            style="TOP:6;LEFT:473; WIDTH:250; BACKGROUND-COLOR:<%=gstrAltBackColor%>"
            readOnly NAME="txtEvaluater">
    
        <SPAN id=lblRrvSig class=DefLabel onkeydown="Gen_onkeydown"
            onclick=SignatureClick(1) style="cursor:hand;LEFT:15; WIDTH:130; TOP:32; TEXT-ALIGN: left">
            <%=gstrEvaTitle%> Signature
        </SPAN>
        
        <INPUT id=chkRrvSig
            title="Submit Final <%=mstrReReviewType%>"
            onclick=SignatureClick(11) 
            style="LEFT:145; WIDTH:20; TOP:30; HEIGHT: 20;BACKGROUND-COLOR:<%=gstrBackColor%>"
            tabIndex=3 type=checkbox>  
              
        <SPAN id=lblRvwSig class=DefLabel onkeydown="Gen_onkeydown"
            onclick=SignatureClick(2) style="cursor:hand;LEFT:300; WIDTH:110; TOP:32; TEXT-ALIGN: left">
            Supervisor Signature
        </SPAN>
        
        <INPUT id=chkRvwSig
            title="Submit Final <%=mstrReReviewType%>"
            onclick=SignatureClick(12) 
            style="LEFT:410; WIDTH:20; TOP:30; HEIGHT: 20;BACKGROUND-COLOR:<%=gstrBackColor%>"
            tabIndex=3 type=checkbox>    
        <SPAN id=lblSubmit class=DefLabel onkeydown="Gen_onkeydown"
            onclick=SignatureClick(3) style="cursor:hand;LEFT:515; WIDTH:100; TOP:32; TEXT-ALIGN: left">
            Submit To Reports
        </SPAN>
        
        <INPUT id=chkSubmit
            title="Submit Final <%=mstrReReviewType%>"
            onclick=SignatureClick(13) 
            style="LEFT:615; WIDTH:20; TOP:30; HEIGHT: 20;BACKGROUND-COLOR:<%=gstrBackColor%>"
            tabIndex=3 type=checkbox>    

        <SPAN id=lblResponse class=DefLabel style="LEFT:-1183; WIDTH:205; TOP:30;TEXT-ALIGN: left">
            Response
            <SELECT id=cboResponse title="Reviewer Response" style="LEFT:75; WIDTH:125;BACKGROUND-COLOR:<%=gstrBackColor%>"
                disabled NAME="cboResponse" tabindex=5>
                <OPTION VALUE=0 SELECTED>
                <%
                Dim adCmdRR
                Dim adRsRR

                Set adCmdRR = GetAdoCmd("spGetOptListValues")
                    AddParmIn adCmdRR, "@LstName", adVarChar, 50, "ReReviewResponse"
                Set adRsRR = GetAdoRs(adCmdRR)
                Do While Not adRsRR.EOF
                    Response.Write adRsRR.Fields("OptionValue").Value
                    adRsRR.MoveNext 
                Loop 
                adRsRR.Close
                Set adRsRR = Nothing
                Set adCmdRR = Nothing
                %>
            </SELECT>
        </SPAN>
        <SPAN id=lblResponseDue class=DefLabel style="TOP:30; LEFT:-1420; WIDTH:80">
            Response Due:
        </SPAN>
            
        <INPUT id=txtResponseDue type=text 
            style="TOP:30;LEFT:-1500; WIDTH:80; BACKGROUND-COLOR:<%=gstrBackColor%>"
            disabled NAME="txtResponseDue">

        <DIV id=divFindReviewForReReview class=DefPageFrame style="POSITION: absolute;LEFT:-2;z-index:2000; HEIGHT:435; WIDTH:745; TOP:0;background-color:transparent; BORDER-COLOR: <%=gstrDefButtonColor%>">
            <IFRAME ID=fraFindReviewForReReview src="FindReview.asp?Parms=<%=gstrUserID%>^<%=gblnUserAdmin%>^<%=gblnUserQA%>^<%=Request.Form("ProgramsSelected")%>^<%=glngAliasPosID%>"
                STYLE="positon:absolute; LEFT:0; WIDTH:742; HEIGHT:433; TOP:0; BORDER:none" FRAMEBORDER=0>
            </IFRAME>
        </DIV>
        <DIV id=divReReviewEntry class=DefPageFrame 
            style="POSITION: absolute;LEFT:-2; HEIGHT:350; WIDTH:745; TOP:55; BORDER-COLOR: <%=gstrDefButtonColor%>">
            <DIV id=ReviewSummary
                style="OVERFLOW:hidden;
                    Left: -2;
                    border-style: solid;
                    border-width: 2;
                    border-color: <%=gstrBorderColor%>;
                    TOP: 0; 
                    HEIGHT: 100; 
                    WIDTH: 745; 
                    COLOR: black; 
                    BACKGROUND-COLOR: <%=gstrBackColor%>">
                    
                <SPAN id=lblCaseID class=DefLabel style="TOP:5; LEFT:8; WIDTH:70"><B> Review ID:</B></SPAN>
                <SPAN id=lblCaseIDValue class=DefLabel style="TOP:5; LEFT:80; WIDTH:70"><B></B></SPAN>
                <SPAN id=lblReviewMonth class=DefLabel style="TOP:5; LEFT:148; WIDTH:85"><B> Review Month:</B></SPAN>
                <SPAN id=lblReviewMonthValue class=DefLabel style="TOP:5; LEFT:235; WIDTH:70"><B></B></SPAN>
                <SPAN id=lblReviewDate class=DefLabel style="TOP:5; LEFT:305; WIDTH:80"><B> Review Date:</B></SPAN>
                <SPAN id=lblReviewDateValue class=DefLabel style="TOP:5; LEFT:385; WIDTH:80"><B></B></SPAN>
                <SPAN id=lblCaseNumber class=DefLabel style="TOP:25; LEFT:8; WIDTH:80"><B> Case Number:</B></SPAN>
                <SPAN id=lblCaseNumberValue class=DefLabel style="TOP:25; LEFT:90; WIDTH:80"><B></B></SPAN>
                <SPAN id=lblCaseName class=DefLabel style="TOP:25; LEFT:264; WIDTH:85"><B> Case Name:</B></SPAN>
                <SPAN id=lblCaseNameValue class=DefLabel style="TOP:25; LEFT:340; WIDTH:180"><B></B></SPAN>
                <SPAN id=lblReviewClass class=DefLabel style="TOP:42; LEFT:8; WIDTH:80"><B> Review Class:</B></SPAN>
                <SPAN id=lblReviewClassValue class=DefLabel style="TOP:42; LEFT:90; WIDTH:175"><B></B></SPAN>
                <SPAN id=lblReviewStatus class=DefLabel style="TOP:42; LEFT:264; WIDTH:100"><B> Review Status:</B></SPAN>
                <SPAN id=lblReviewStatusValue class=DefLabel style="TOP:42; LEFT:355; WIDTH:70"><B></B></SPAN>
                <SPAN id=lblReviewerName class=DefLabel style="TOP:60; LEFT:8; WIDTH:80"><B> <%=gstrRvwTitle%>:</B></SPAN>
                <SPAN id=lblReviewerNameValue class=DefLabel style="TOP:60; LEFT:90; WIDTH:170"><B></B></SPAN>
                <SPAN id=lblWorkerName class=DefLabel style="TOP:77; LEFT:8; WIDTH:90"><B> <%=gstrWkrTitle%>:</B></SPAN>
                <SPAN id=lblWorkerNameValue class=DefLabel style="TOP:77; LEFT:90; WIDTH:170"><B></B></SPAN>
                <SPAN id=lblWorkerResponse class=DefLabel style="TOP:77; LEFT:264; WIDTH:120"><B> <%=gstrWkrTitle%> Response:</B></SPAN>
                <SPAN id=lblWorkerResponseValue class=DefLabel style="TOP:77; LEFT:384; WIDTH:270"><B></B></SPAN>

                <DIV id=divPrograms
                    style="OVERFLOW:auto;
                        Left: 620;
                        border-style: solid;
                        border-width: 2;
                        border-color: <%=gstrBorderColor%>;
                        TOP: -2; 
                        HEIGHT: 102; 
                        WIDTH: 125; 
                        COLOR: black; 
                        BACKGROUND-COLOR: <%=gstrBackColor%>">
                </DIV>
            </DIV>

            <DIV id=divTabButton1 class="defRectangle DivTab" style="LEFT:0;TOP:98;WIDTH:248" onclick="divTabs_onclick(1)" onkeydown="divTabs_onkeydown(1)">
                Elements
            </DIV>
            <DIV id=divTab1 class=defRectangle
                style="OVERFLOW:auto;LEFT:-2; TOP:117; WIDTH:745; HEIGHT:230; BORDER-RIGHT-STYLE:none; FONT-WEIGHT:normal; BACKGROUND-COLOR:<%=gstrAltBackColor%>">
            </DIV>
            <DIV id=divTabButton2 class="defRectangle DivTab" style="LEFT:248;TOP:98;WIDTH:248" onclick="divTabs_onclick(2)" onkeydown="divTabs_onkeydown(2)">
                History
            </DIV>
            <DIV id=divTab2 class=defRectangle 
                style="OVERFLOW:auto;LEFT:-2; TOP:117; WIDTH:745; HEIGHT:250; BORDER-RIGHT-STYLE:none; FONT-WEIGHT:normal; BACKGROUND-COLOR:<%=gstrAltBackColor%>">

                <DIV id=divReReviewHistory Class=TableDivArea style="LEFT:5; TOP:5; WIDTH:735; HEIGHT:185; 
                    OVERFLOW:auto; FONT-WEIGHT:normal;z-index:1200">
                    <TABLE id=tblAudit Border=0 Width=715 CellSpacing=0
                        Style="position:absolute;overflow: hidden; TOP:0">
                        <THEAD id=tbhAudit style="height:17">
                            <TR id=thrAudit>
                                <TD class=CellLabel style="width:150">Date Of Action</TD>
                                <TD class=CellLabel style="width:140">User</TD>
                                <TD class=CellLabel style="width:140">Entry Name</TD>
                                <TD class=CellLabel style="width:145">Value Before</TD>
                                <TD class=CellLabel style="width:145">Value After</TD>
                            </TR>
                        </THEAD>
                        <TBODY id=tbdAudit>
                        </TBODY>
                    </TABLE>
                </DIV>
                <BUTTON id=cmdPrintAudit
                    class=DefButton 
                    style="LEFT:610; WIDTH:125; TOP:195; HEIGHT: 20">
                    Print Audit History
                </BUTTON>
            </DIV>
            <DIV id=divTabButton3 class="defRectangle DivTab" style="LEFT:496;TOP:98;WIDTH:248" onclick="divTabs_onclick(3)" onkeydown="divTabs_onkeydown(3)">
                Re-Review History of Review
            </DIV>
            <DIV id=divTab3 class=defRectangle 
                style="OVERFLOW:auto;LEFT:-2; TOP:117; WIDTH:745; HEIGHT:250; BORDER-RIGHT-STYLE:none; FONT-WEIGHT:normal; BACKGROUND-COLOR:<%=gstrAltBackColor%>">
                <DIV id=divReviewHistory Class=TableDivArea style="LEFT:5; TOP:5; WIDTH:735; HEIGHT:235; 
                    OVERFLOW:auto; FONT-WEIGHT:normal;z-index:1200">
                    <TABLE id=tblHistory Border=0 Width=715 CellSpacing=0
                        Style="position:absolute;overflow: hidden; TOP:0">
                        <THEAD id=tbhHistory style="height:17">
                            <TR id=thrHistory>
                                <TD class=CellLabel style="width:50">Record ID</TD>
                                <TD class=CellLabel style="width:140">Program</TD>
                                <TD class=CellLabel style="width:140"><%=gstrEvaTitle%></TD>
                                <TD class=CellLabel style="width:145">Date Entered</TD>
                                <TD class=CellLabel style="width:120">Type</TD>
                                <TD class=CellLabel style="width:125">Submitted</TD>
                            </TR>
                        </THEAD>
                        <TBODY id=tbdHistory>
                        </TBODY>
                    </TABLE>
                </DIV>
           </DIV>
        </DIV>
        <DIV id=Buttons
            style="LEFT: -2; 
                border-style: solid;
                border-width: 2;
                border-color: <%=gstrBorderColor%>;
                TOP: 402; 
                HEIGHT: 40px; 
                WIDTH: 745; 
                BACKGROUND-COLOR: <%=gstrAltBackColor%>">

            <SPAN id=lblDatabaseStatus
                class=DefLabel
                style="VISIBILITY:hidden; LEFT:5; WIDTH:200; TOP:10; TEXT-ALIGN:center">
                Accessing Database...
            </SPAN>

            <BUTTON id=cmdFindRecord 
                class=DefButton
                title="Find <%=mstrReReviewType%> Record" 
                style="LEFT:10; WIDTH:60;  TOP:7; HEIGHT:20" 
                accesskey=F
                tabIndex=284>
                <u>F</u>ind
            </BUTTON>
            
            <BUTTON id=cmdAddRecord 
                class=DefButton
                title="Add a New <%=mstrReReviewType%> Record" 
                style="LEFT:235; WIDTH:60; TOP:7; HEIGHT:20" 
                accesskey=A
                tabIndex=285>
                <u>A</u>dd
            </BUTTON>

            <BUTTON id=cmdChangeRecord 
                class=DefButton
                title="Modify the Current <%=mstrReReviewType%> Record" 
                disabled=True
                style="LEFT:300; WIDTH:60; TOP:7; HEIGHT:20" 
                accesskey=C
                tabIndex=285>
                <u>E</u>dit
            </BUTTON>
            
            <BUTTON id=cmdDeleteRecord 
                class=DefButton
                title="Delete the Current <%=mstrReReviewType%> Record" 
                disabled=True   
                style="LEFT:365; WIDTH:60; TOP:7; HEIGHT:20" 
                accesskey=D
                tabIndex=286>
                <u>D</u>elete
            </BUTTON>

            <BUTTON id=cmdCancelEdit 
                class=DefButton
                title="Cancel Add or Change" 
                disabled=true
                style="LEFT:430; WIDTH:60; TOP:7; HEIGHT:20" 
                accesskey=L
                tabIndex=287>
                Cance<u>l</u>
            </BUTTON>

            <BUTTON id=cmdSaveRecord 
                class=DefButton
                title="Save New <%=mstrReReviewType%> Record or Changes to Current <%=mstrReReviewType%> Record" 
                disabled=true
                style="LEFT:495; WIDTH:60; TOP:7; HEIGHT:20" 
                accesskey=S
                tabIndex=288>
                <u>S</u>ave
            </BUTTON>

            <BUTTON id=cmdPrint
                class=DefButton
                title="Save and Print Current <%=mstrReReviewType%> Record"
                disabled=true 
                style="LEFT:560; WIDTH:60; TOP:7; HEIGHT: 20" 
                accesskey=P
                tabIndex=289>
                <u>P</u>rint
            </BUTTON>

            <BUTTON id=cmdClose 
                class=DefButton
                title="Close the <%=mstrReReviewType%> Form" 
                style="LEFT:655; WIDTH:60; TOP:7; HEIGHT:20" 
                tabIndex=290>
                Close
            </BUTTON>
        </DIV>
    </DIV>
    <FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="ReReviewAddEdit.ASP" ID=Form>
        <INPUT TYPE="hidden" Name="FormAction" VALUE="" ID=FormAction>
        <INPUT TYPE="hidden" Name="intCount" VALUE="0" ID=intCount>
        <INPUT TYPE="hidden" Name="ProgramsSelected" VALUE="<%=Request.Form("ProgramsSelected")%>" ID=ProgramsSelected>
        <INPUT TYPE="hidden" Name="CalledFrom" VALUE="<%=Request.Form("CalledFrom")%>" ID=CalledFrom>
        <INPUT TYPE="hidden" Name="casID" VALUE="" ID=casID>
        <INPUT TYPE="hidden" Name="UserID" value="<%=gstrUserID%>" ID=UserID>
        <INPUT TYPE="hidden" Name="Password" VALUE="<%=gstrPassword%>" ID=Password>
        <INPUT TYPE="hidden" name="rrvID" value=<%=mlngReReviewID%> id=rrvID>
        <%If mlngReReviewID > 0 And Not madoReReview.EOF Then%>
            <!--#include file="IncReRevDefEdit.asp"-->
        <%Else%>
            <!--#include file="IncReRevDefBlank.asp"-->
        <%End If%>
        <INPUT TYPE="hidden" name="SaveCompleted" value="" id=SaveCompleted>
        <INPUT TYPE="hidden" name="ReReviewElements" value="<%=mstrReReviewElements%>" id=ReReviewElements>
        <INPUT TYPE="hidden" name="ReReviewElementsEdit" value="" id=ReReviewElementsEdit>
        <INPUT TYPE="hidden" name="ReReviewElementsWrite" value="" id=ReReviewElementsWrite>
        <INPUT TYPE="hidden" name="ReReviewTypeID" value="<%=mlngReReviewTypeID%>" id=ReReviewTypeID>
        <INPUT TYPE="hidden" name="UpdateString" value="" id=UpdateString>
    </FORM>
    <P>
    <%
    Set gadoCmd = Nothing
    Set madoRs = Nothing
    gadoCon.Close
    Set gadoCon = Nothing
    %>
    <!--#include file="IncSvrFunctions.asp"-->
    <!--#include file="IncCmnCliFunctions.asp"-->
    <!--#include file="IncNavigateControls.asp"-->
    <br><br>
</BODY>
</HTML>
