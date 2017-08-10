<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: FactorAddEdit.asp                                         '
'  Purpose: The primary admin data entry screen for maintaining the causal  '
'           factors for each eligibility element.                           '
'           This form is only available to admin users.                     '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim mstrPageTitle
Dim madoRs, madoRsFactors
Dim mstrAction, strHoldName
Dim mdctFactors, mstrHidden, mdctEFLinks
Dim moDictObj
Dim mintRowID, mintTabIndex 
Dim mintReturnID, mstrMessage, mintFactorIndexID

mstrAction = ReqForm("FormAction")

If Len(mstrAction) = 0 Then
    mstrAction = "Load"
Else
End If
mstrPageTitle = "Add/Edit Elements and Causal Factors"
Select Case mstrAction
    Case "AddSave"
        Set gadoCmd = GetAdoCmd("spFactorAdd")
            AddParmIn gadoCmd, "@fctShortName", adVarChar, 250, ReqForm("FactorName")
            AddParmIn gadoCmd, "@fctLongName", adVarChar, 5000, ReqForm("FactorLongName")
            AddParmOut gadoCmd, "@fctID", adInteger, 0
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mintReturnID = gadoCmd.Parameters("@fctID").Value
            Select Case mintReturnID
                Case -1
                    mstrMessage = "Error encountered while trying to Add " & ReqForm("FactorName") & "." & vbCrLf
                Case Else
                    mstrMessage = ""
            End Select
    
            If mstrMessage = "" Then
                Call ProcessFactorConnections(gadoCmd.Parameters("@fctID").Value)
            End If
        Set gadoCmd = Nothing
    Case "EditSave"
        Set gadoCmd = GetAdoCmd("spFactorUpd")
            AddParmIn gadoCmd, "@fctID", adInteger, 0, ReqForm("FactorID")
            AddParmIn gadoCmd, "@fctShortName", adVarChar, 250, ReqForm("FactorName")
            AddParmIn gadoCmd, "@fctLongName", adVarChar, 5000, ReqForm("FactorLongName")
            AddParmOut gadoCmd, "@ReturnID", adInteger, 0 
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mintReturnID = gadoCmd.Parameters("@ReturnID").Value
            Select Case mintReturnID
                Case -1
                    mstrMessage = "Error encountered while trying to save " & ReqForm("FactorName") & "." & vbCrLf
                Case Else
                    mstrMessage = ""
            End Select
    
            If mstrMessage = "" Then
                Call ProcessFactorConnections(ReqForm("FactorID"))
            End If
        Set gadoCmd = Nothing
    Case "Delete"
        Set gadoCmd = GetAdoCmd("spFactorDel")
            AddParmIn gadoCmd, "@fctID", adInteger, 0, ReqForm("FactorID")
            AddParmOut gadoCmd, "@ReturnID", adInteger, 0
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mintReturnID = gadoCmd.Parameters("@ReturnID").Value
            Select Case mintReturnID
                Case 0
                    mstrMessage = ""
                Case -1
                    mstrMessage = "Factor could not be deleted.  It has been used on a review."
                Case -2
                    mstrMessage = "Error encountered while trying to delete Factor."
            End Select
        Set gadoCmd = Nothing
End Select

Set madoRsFactors = Server.CreateObject("ADODB.Recordset")
' Factors
Set gadoCmd = GetAdoCmd("spCausalFactorList")
    madoRsFactors.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
Set gadoCmd = Nothing

mintRowID = 0
madoRsFactors.Sort = "fctShortName"
Set mdctFactors = CreateObject("Scripting.Dictionary")
Do While Not madoRsFactors.EOF
    mdctFactors.Add CLng(madoRsFactors.Fields("fctID").value), madoRsFactors.Fields("fctShortName").Value & "^" & Replace(madoRsFactors.Fields("fctLongName").Value,Chr(13) & Chr(10),"[vbCrlf]")
'    mdctFactors.Add CLng(madoRsFactors.Fields("fctID").value), madoRsFactors.Fields("fctShortName").Value & "^" & madoRsFactors.Fields("FactorUsedLast").Value & "^" & madoRsFactors.Fields("fctLongName").Value
    madoRsFactors.MoveNext
Loop

strHoldName = ""
Set madoRsFactors = Server.CreateObject("ADODB.Recordset")
' Element / Factor Links
Set gadoCmd = GetAdoCmd("spElementFactorLinks")
    madoRsFactors.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
Set gadoCmd = Nothing
madoRsFactors.Sort = "elmShortName, Program"
Set mdctEFLinks = CreateObject("Scripting.Dictionary")
Do While Not madoRsFactors.EOF
    If strHoldName <> madoRsFactors.Fields("elmShortName").value Then
        strHoldName = madoRsFactors.Fields("elmShortName").value
        mdctEFLinks.Add strHoldName, "|"
    End If
    mdctEFLinks(strHoldName) = mdctEFLinks(strHoldName) & madoRsFactors.Fields("elmID").value & "^" & madoRsFactors.Fields("Program").value & "|"
    madoRsFactors.MoveNext
Loop
mintRowID = 0

Sub ProcessFactorConnections(intFactorID)
    Dim intI, strRecord
    
    For intI = 1 To 100
        strRecord = Parse(ReqForm("FactorList"),"|",intI)
        If strRecord = "" Then Exit For
        Set gadoCmd = GetAdoCmd("spElementFactorLinkUpd")
            AddParmIn gadoCmd, "@ElementID", adInteger, 0, Parse(strRecord,"^",1)
            AddParmIn gadoCmd, "@FactorID", adInteger, 0, intFactorID
            AddParmIn gadoCmd, "@SortOrder", adInteger, 0, Null
            AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Parse(strRecord,"^",3))
            AddParmIn gadoCmd, "@Action", adChar, 1, Parse(strRecord,"^",2)
            AddParmOut gadoCmd, "@ReturnID", adInteger, 0
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mintReturnID = gadoCmd.Parameters("@ReturnID").Value
            Select Case mintReturnID
                Case 0
                Case -1
                    mstrMessage = mstrMessage & "Link with Action/Screen ID " & Parse(strRecord,"^",1) & " could not be deleted/end dated.  It has been used on a review." & vbCrLf
                Case -2
                    mstrMessage = mstrMessage & "Error encountered while trying to delete link with Action/Screen ID " & Parse(strRecord,"^",1) & "." & vbCrLf
            End Select
    Next
End Sub

Function LocalGetTabIndex()
    LocalGetTabIndex = mintTabIndex
    mintTabIndex = CInt(mintTabIndex) + 1
End Function
%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
    <STYLE>
        .CheckBoxList
            {
            position: absolute;
            HEIGHT:292;
            WIDTH:200;
            background-color:white;
            border-style:inset;
            border-width:2    
            }
        .ControlDiv
            {
            POSITION: absolute;
            BORDER-STYLE: solid;
            BORDER-WIDTH: 1px;
            BORDER-COLOR: <%=gstrBorderColor%>;
            BACKGROUND-COLOR: <%=gstrBackColor%>;
            COLOR: black;
            HEIGHT: 400;
            LEFT: 1
            }
   </STYLE>
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Option Explicit
Dim mdctFactors
Dim mdctLinks, mdctEFLinks
Dim mdctDisplayed
Dim mintSaveCheck

Sub window_onload()
    Dim intI
    If "<%=mstrMessage%>" <> "" Then        
        MsgBox "<%=mstrMessage%>", vbOkOnly, "Case Review Maintenance"
    End If
    Set mdctFactors = CreateObject("Scripting.Dictionary")
    Set mdctLinks = CreateObject("Scripting.Dictionary")
    Set mdctEFLinks = CreateObject("Scripting.Dictionary")
    Set mdctDisplayed = CreateObject("Scripting.Dictionary")
    <%
    For Each moDictObj In mdctFactors
        Response.Write "mdctFactors.Add CLng(" & moDictObj & "), """ & mdctFactors(moDictObj) & """" & vbCrLf
    Next
    For Each moDictObj In mdctEFLinks
        Response.Write "mdctEFLinks.Add """ & moDictObj & """, """ & mdctEFLinks(moDictObj) & """" & vbCrLf
    Next
    %>
    Form.FactorIndexID.Value = 0
    Call Result_onclick(0)
    Call DisableControls("Load")
End Sub

Sub cmdSave_onclick()
    Dim strErrorMessage
    
    strErrorMessage = Validate()
    
    If strErrorMessage <> "" Then
        MsgBox strErrorMessage, vbOkOnly, "Case Review Maintenance"
        Exit Sub
    End If
    
    If Form.FormAction.Value = "Add" Then
        Form.FactorID.value = 0
        Form.FormAction.Value = "AddSave"
    ElseIf Form.FormAction.Value = "Edit" Then
        Form.FactorID.value = document.all("hidRowInfo" & Form.FactorIndexID.Value).value
        Form.FormAction.Value = "EditSave"
    End If
    Form.BuildCompleted.value = "S"
    mintSaveCheck = Window.setInterval("CheckForBuildCompleted", 100)
    divSaving.style.left = 0
    divPageFrame.style.left = -1000
End Sub

Function CheckForBuildCompleted()
    If Form.BuildCompleted.value = "Y" Then
        ' Disable timer
        window.clearInterval mintSaveCheck
        Form.submit
    ElseIf Form.BuildCompleted.value = "S" Then
        Form.BuildCompleted.value = "N"
        Call FillForm()
    End If
End Function

Sub cmdEdit_onclick()
    Call FillScreen()
    Form.FormAction.value = "Edit"
    Call DisableControls("Edit")
    txtFactorName.focus
End Sub

Sub cmdCancel_onclick()
    Call ClearScreen()
    Call FillScreen()
    Call DisableControls("Cancel")
End Sub

Sub cmdAdd_onclick()
    Call ClearScreen()
    Form.FormAction.value = "Add"
    Call DisableControls("Edit")
    txtFactorName.focus
End Sub

Sub cmdDelete_onclick()
    Dim intResp
    
    If txtLastUsed.value <> "never" Then
        Msgbox "Factor has been used in a review and cannot be deleted.", vbOkOnly, "Case Review Maintenance"
        Exit Sub
    Else
        intResp = MsgBox("Delete the Factor?", vbQuestion + vbYesNo, "Delete")
        If intResp = vbNo Then Exit Sub
    End If
    
    Form.FactorID.value = document.all("hidRowInfo" & Form.FactorIndexID.Value).value
    Form.FormAction.value = "Delete"
    Form.submit
End Sub

Function Validate()
    Dim strRecord
    Dim dtmLastUsed
    Dim strMessage
    Dim intSelectedID
    Dim intI

    strMessage = ""
    If Trim(txtFactorName.value) = "" Then
        strMessage = "Factor text cannot be blank."
        Validate = strMessage
        Exit Function
    End If
    
    If Form.FormAction.value = "Add" Then
        intSelectedID = -1
    Else
        intSelectedID = Form.FactorIndexID.value
    End If
    
    For intI = 0 To tblFactors.Rows.Length - 1
        If UCase(tblFactors.Rows(intI).Cells(0).innerText) = Trim(UCase(txtFactorName.value)) And CInt(intI) <> CInt(intSelectedID) Then
            strMessage = txtFactorName.value & " cannot be the text of this Factor.  It is identical to another Factor."
            Validate = strMessage
            Exit Function
        End If
    Next
End Function

Function CleanTextRecordParsers(strText, strDir, strType)
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
            strText = Replace(strText, Chr(9) & "#as#", "*")
            strText = Replace(strText, Chr(9) & "#ex#", "!")
            strText = Replace(strText, "[vbCrlf]", Chr(13) & Chr(10))
            If strType = "All" Then
                strText = Replace(strText, Chr(9) & "#sq#", "'")
                strText = Replace(strText, Chr(9) & "#dq#", """")
            End If
        ElseIf strDir = "ToDb" Then
            strText = Replace(strText, "^", Chr(9) & "#ca#")
            strText = Replace(strText, "|", Chr(9) & "#ba#")
            strText = Replace(strText, "*", Chr(9) & "#as#")
            strText = Replace(strText, "!", Chr(9) & "#ex#")
            If strType = "All" Then
                strText = Replace(strText, "'", Chr(9) & "#sq#")
                strText = Replace(strText, """", Chr(9) & "#dq#")
            End If
        End If
        CleanTextRecordParsers = strText
    End If
End Function

Sub FillForm()
    Form.FactorName.value = txtFactorName.value
    Form.FactorLongName.value = CleanTextRecordParsers(txtFactorLongName.value,"ToDb","All")
    Call BuildFactorList()
End Sub

Sub BuildFactorList()
    Dim intI
    Dim intElementID, strEndDate, strNewEndDate
    Dim strFactorList, strAction
    
    If hidLinkCount.value = 0 Then
        Form.FactorList.value = ""
        Form.BuildCompleted.value = "Y"
        Exit Sub
    End If

    For intI = 1 To hidLinkCount.value
        intElementID = document.all("hidLinkRow" & intI).value
        If GetButtonValue(intI) = "Remove" Or GetButtonValue(intI) = "Add" Then
            strNewEndDate = ""
        Else
            strNewEndDate = document.all("txtEndDate" & intI).value
        End If
        strAction = ""
        If GetButtonValue(intI) = "Cancel" Then
            'Delete link
            strAction = "D"
        ElseIf GetButtonValue(intI) = "Edit" And strNewEndDate <> "" Then
            'Check if end date is different from original
            strEndDate = Parse(mdctLinks(CLng(intElementID)),"^",5) 
            If strEndDate <> "" Then
                If CDate(strEndDate) <> CDate(strNewEndDate) Then
                    strAction = "A"
                End If
            Else
                strAction = "A"
            End If
            
        ElseIf GetButtonValue(intI) = "End" Then
            'Check if there was originally an end date
            strEndDate = Parse(mdctLinks(CLng(intElementID)),"^",5) 
            If strEndDate <> "" Then
                strAction = "A"
            End If
        ElseIf GetButtonValue(intI) = "Remove" Then
            'Adding a new link
            strAction = "A"
        End If
        If strAction <> "" Then
            strFactorList = strFactorList & intElementID & "^" & strAction & "^" & strNewEndDate & "|"
        End If
    Next
    Form.FactorList.value = strFactorList
    Form.BuildCompleted.value = "Y"
End Sub

Sub cmdClose_onclick()
    Call window.parent.cmdClose_onclick()
End Sub

Sub Result_onclick(intRowID)
    Dim strRow
    
    If tblFactors.Rows.length = 0 Then Exit Sub
    
    If cmdAdd.disabled = True Then 
        MsgBox "Screen is in Edit mode.  Save or Cancel current record before selecting another Factor", vbOkOnly, "Case Review Maintenance"
        Exit Sub
    End If
    
    If IsNumeric(Form.FactorIndexID.Value) Then
        strRow = "tblRow" & Form.FactorIndexID.Value
        tblFactors.Rows(strRow).className = "TableRow"
        tblFactors.Rows(strRow).cells(0).tabindex = -1
    End If

    strRow = "tblRow" & intRowID
    tblFactors.Rows(strRow).className = "TableSelectedRow"
    tblFactors.Rows(strRow).cells(0).focus
    tblFactors.Rows(strRow).cells(0).tabindex = 9

    Form.FactorIndexID.Value = intRowID
    Set mdctLinks = window.showModalDialog("ElementSort.asp?Action=FactorLinks&FactorID=" & document.all("hidRowInfo" & Form.FactorIndexID.Value).value, , "dialogWidth:210px;dialogHeight:120px;scrollbars:no;center:yes;border:thin;help:no;status:no")
    Call FillScreen()
End Sub

Sub FillScreen()
    Dim strRecord
    If tblFactors.Rows.length = 0 Then Exit Sub

    strRecord = mdctFactors(CLng(document.all("hidRowInfo" & Form.FactorIndexID.Value).value))

    txtFactorName.value = Parse(strRecord,"^",1)
    txtFactorLongName.value = CleanTextRecordParsers(Parse(strRecord,"^",2),"FromDb","All")
    txtFactorID.value = document.all("hidRowInfo" & Form.FactorIndexID.Value).value
    Call DisplayLinks()
    'cmdAvailableLinks.disabled = False
End Sub

Sub DisplayLinks()
    Dim oRow, oCell
    Dim intI, intJ
    Dim strHoldElement, strElement
    Dim oLink, strRecord, strAction, strLastUsed
    Dim strHidden

'1 TabName
'2 Program
'3 Element
'4 LinkID
'5 LinkEndDate
'6 LinkLastDate
    For intI = 0 To tblLinksBody.rows.length - 1
        tblLinksBody.deleteRow 0
    Next
    intI = 0
    intJ = 1000
    strHoldElement = ""
    strHidden = ""
    strLastUsed = "01/01/2000"

    mdctDisplayed.RemoveAll

    Set oRow = tblLinksBody.insertRow()
    Call AddHeaderRow(oRow, "<B>Current Links</B>", intJ, False)
    For Each oLink In mdctLinks
        strRecord = mdctLinks(oLink)
        strElement = Parse(strRecord,"^",3)
        If strElement <> strHoldElement Then
            intJ = intJ + 1
            Set oRow = tblLinksBody.insertRow()
            oRow.ID = "tbrELink" & intJ
            Call AddHeaderRow(oRow, strElement, intJ, True)
            strHoldElement = strElement
            'strHidden = strHidden & "<input type=hidden id=hidLinkFocus" & intJ & " value=" & intI+1 & ">"
        End If
        If Parse(strRecord,"^",6) = "" Then
            'Never used
            strAction = "Delete"
        ElseIf Parse(strRecord,"^",6) <> "" And Parse(strRecord,"^",5) = "" Then
            'Previously used, NOT end dated
            strAction = "&nbsp;End&nbsp;"
        Else
            'Ended
            strAction = "&nbsp;Edit&nbsp;"
        End If

        If Parse(strRecord,"^",6) <> "" Then
            'Keep the latest date the factor was linked with any element for display
            If CDate(Parse(strRecord,"^",6)) > CDate(strLastUsed) Then
                strLastUsed = Parse(strRecord,"^",6)
            End If
        End If
        
        intI = intI + 1
        Set oRow = tblLinksBody.insertRow()
        oRow.ID = "tbrLink" & intI
        mdctDisplayed.Add CLng(oLink), "E"
        
        Call BuildCell(oRow, Parse(strRecord,"^",2), intI, 1, 10, 235)
        If Parse(strRecord,"^",6) = "" Then
            Call BuildCell(oRow, "never", intI, 2, 1, 60)
        Else
            Call BuildCell(oRow, Parse(strRecord,"^",6), intI, 2, 1, 60)
        End If
        Call BuildCell(oRow, "<input type=text style=""width:70"" disabled onfocus=Date_onfocus(txtEndDate" & intI & ") onblur=Date_onblur(txtEndDate" & intI & ") onkeypress=Date_onkeypress(txtEndDate" & intI & ") id=txtEndDate" & intI & " value=""" & Parse(strRecord,"^",5) & """>", intI, 3, 0, 70)
        Call BuildCell(oRow, "<button disabled id=cmdLink" & intI & " style=""width:50"" onclick=cmdLink_onclick(" & intI & ")>" & strAction & "</button>", intI, 4, 0, 0)
        strHidden = strHidden & "<input type=hidden id=hidLinkRow" & intI & " value=" & oLink & ">"
    Next
    
    hidLinkCount.value = intI
    hidLinkElmCount.value = intJ
    divHoldvalues.innerHTML = strHidden
    If strLastUsed = "01/01/2000" Then strLastUsed = "never"
    txtLastUsed.value = strLastUsed
End Sub

Sub cmdAvailableLinks_onclick()
    cmdAvailableLinks.disabled = True
    Call AddAvailableLinks()
End Sub

Sub AddAvailableLinks()
    Dim strHoldElement, strRecord, strFunction, strHidden
    Dim intI, intJ, intK
    Dim oRow, oLink

    strHoldElement = ""
    intJ = hidLinkElmCount.value
    intI = hidLinkCount.value
    strHidden = divHoldvalues.innerHTML
    Set oRow = tblLinksBody.insertRow()
    Call AddHeaderRow(oRow, "<B>Available New Links</B>", intJ, False)
    For Each oLink In mdctEFLinks
        strRecord = mdctEFLinks(oLink)
        For intK = 2 To 100
            strFunction = Parse(strRecord,"|",intK)
            If strFunction = "" Then Exit For
            If Not mdctDisplayed.Exists(CLng(Parse(strFunction,"^",1))) Then
                If strHoldElement <> oLink Then
                    intJ = intJ + 1
                    Set oRow = tblLinksBody.insertRow()
                    oRow.ID = "tbrELink" & intJ
                    Call AddHeaderRow(oRow, oLink, intJ, True)
                    strHoldElement = oLink
                    'strHidden = strHidden & "<input type=hidden id=hidLinkFocus" & intJ & " value=" & intI+1 & ">"
                End If
                intI = intI + 1
                Set oRow = tblLinksBody.insertRow()
                oRow.ID = "tbrLink" & intI
                Call BuildCell(oRow, Parse(strFunction,"^",2), intI, 1, 10, 235)
                Call BuildCell(oRow, "", intI, 2, 1, 60)
                Call BuildCell(oRow, "", intI, 3, 1, 70)
                Call BuildCell(oRow, "<button id=cmdLink" & intI & " style=""width:50"" onclick=cmdLink_onclick(" & intI & ")>Add</button>", intI, 4, 0, 0)
                strHidden = strHidden & "<input type=hidden id=hidLinkRow" & intI & " value=" & Parse(strFunction,"^",1) & ">"
            End If
        Next
    Next
    divHoldvalues.innerHTML = strHidden
    hidLinkCount.value = intI
    hidLinkElmCount.value = intJ
End Sub

Sub AddHeaderRow(oRow, strText, intRowID, blnFunction)
    Call BuildCell(oRow, strText, intRowID, 1, 1, 235)
    If blnFunction Then
        Call BuildCell(oRow, "", intRowID, 2, 1, 60)
        Call BuildCell(oRow, "", intRowID, 3, 1, 70)
    Else
        If intRowID < 2000 Then
            Call BuildCell(oRow, "<B>Last Used</B>", intRowID, 2, 1, 60)
            Call BuildCell(oRow, "<B>End Date</B>", intRowID, 3, 1, 70)
        Else
            Call BuildCell(oRow, "<B>Action</B>", intRowID, 2, 1, 60)
            Call BuildCell(oRow, "", intRowID, 3, 1, 70)
        End If
    End If
    Call BuildCell(oRow, "", intRowID, 3, 1, 0)
End Sub

Sub BuildCell(oRow, strValue, intRowID, intColID, intPadLeft, intWidth)
    Dim oCell
    Set oCell = oRow.insertCell()
    oCell.ID = "tbd" & intColID & "R" & intRowID
    oCell.style.height = 20
    oCell.style.fontfamily = "Tahoma"
    oCell.style.fontsize = "8pt"
    If intColID = 2 Or (intColID > 2 And intRowID >= 1000) Then
        oCell.style.textalign = "center"
    Else
        oCell.style.textalign = "left"
    End If
    oCell.style.paddingleft = intPadLeft
    If intColID < 4 Then
        oCell.style.width = intWidth
        oCell.style.border = "1 solid"
    End If
    If intColID > 2 And intRowID < 1000 Then
        oCell.vAlign = "top"
    Else
        oCell.vAlign = "bottom"
    End If
    If intRowID < 1000 And strValue = "" Then strValue = "&nbsp;"
    oCell.innerHTML = strValue
End Sub

Sub ClearScreen()
    Dim intI
    txtFactorID.value = ""
    txtFactorName.value = ""
    txtFactorLongName.value = ""
    txtLastUsed.value = ""
    For intI = 0 To tblLinksBody.rows.length - 1
        tblLinksBody.deleteRow 0
    Next
    hidLinkCount.value = 0
    hidLinkElmCount.value = 1000
    divHoldvalues.innerHTML = ""
End Sub

Sub DisableControls(strMode)
    Select Case strMode
        Case "Edit"
            cmdDelete.disabled = True
            cmdAdd.disabled = True
            cmdEdit.disabled = True
            cmdCancel.disabled = False
            cmdSave.disabled = False
            cmdAvailableLinks.disabled = False
            txtFactorName.disabled = False
            txtFactorLongName.disabled = False
            Call DisableLinks(False)
        Case Else
            cmdDelete.disabled = False
            cmdAdd.disabled = False
            cmdEdit.disabled = False
            cmdCancel.disabled = True
            cmdSave.disabled = True
            cmdAvailableLinks.disabled = True
            txtFactorName.disabled = True
            txtFactorLongName.disabled = True
            Call DisableLinks(True)
    End Select
End Sub

Sub DisableLinks(blnVal)
    Dim intI
    
    If hidLinkCount.value > 0 Then
        For intI = 1 To hidLinkCount.value
            document.all("cmdLink" & intI).disabled = blnVal
        Next
    End If
End Sub

Sub ToggleCheckBox(oControl)
    If oControl.disabled = True Then Exit Sub
    oControl.checked = Not oControl.checked
End Sub

Function GetButtonValue(intRowID)
    If InStr(document.all("cmdLink" & intRowID).value,"End") > 0 Then
        GetButtonValue = "End"
    ElseIf InStr(document.all("cmdLink" & intRowID).value,"Edit") > 0 Then
        GetButtonValue = "Edit"
    ElseIf InStr(document.all("cmdLink" & intRowID).value,"Delete") > 0 Then
        GetButtonValue = "Delete"
    ElseIf InStr(document.all("cmdLink" & intRowID).value,"Cancel") > 0 Then
        GetButtonValue = "Cancel"
    ElseIf InStr(document.all("cmdLink" & intRowID).value,"Add") > 0 Then
        GetButtonValue = "Add"
    ElseIf InStr(document.all("cmdLink" & intRowID).value,"Remove") > 0 Then
        GetButtonValue = "Remove"
    End If
End Function

Sub cmdLink_onclick(intRowID)
    Select Case GetButtonValue(intRowID)
        Case "End"
            document.all("cmdLink" & intRowID).disabled = True
            document.all("txtEndDate" & intRowID).disabled = False
            document.all("txtEndDate" & intRowID).value = FormatDateTime(Now(),2)
            document.all("txtEndDate" & intRowID).focus
        Case "Edit"
            document.all("cmdLink" & intRowID).disabled = True
            document.all("txtEndDate" & intRowID).disabled = False
            document.all("txtEndDate" & intRowID).focus
        Case "Delete"
            document.all("tbd1R" & intRowID).style.textDecorationLineThrough = True
            document.all("cmdLink" & intRowID).value = "Cancel"
        Case "Cancel"
            document.all("tbd1R" & intRowID).style.textDecorationLineThrough = False
            document.all("cmdLink" & intRowID).value = "Delete"
        Case "Add"
            document.all("cmdLink" & intRowID).value = "Remove"
            document.all("tbd2R" & intRowID).innerText = "Add"
            document.all("tbd2R" & intRowID).style.color = "red"
        Case "Remove"
            document.all("cmdLink" & intRowID).value = "Add"
            document.all("tbd2R" & intRowID).innerText = " "
    End Select
End Sub

Sub tblLinksBody_onkeyPress()
    If hidLinkElmCount.value <= 0 Then Exit Sub
    Call GoToRowL(Chr(window.event.keyCode))
End Sub

Sub GoToRowL(strText)
    Dim intI, intButton
    strText = UCase(strText)

    For intI = 0 To tblLinksBody.rows.length - 1
        If Left(tblLinksBody.rows(intI).ID,8) = "tbrELink" Then
            If UCase(Left(tblLinksBody.rows(intI).cells(0).innerText,1)) = strText Then
                tblLinksBody.rows(intI).scrollIntoView
                Exit For
            End If
        End If
    Next
End Sub

Sub Table_onkeyPress()
    If cmdEdit.disabled = True Then Exit Sub
    Call GoToRow(Chr(window.event.keyCode))
End Sub

Sub GoToRow(strText)
    Dim intI
    strText = UCase(strText)
    
    For intI = 0 To tblFactors.rows.length - 1
        If UCase(Left(document.all("tblRow" & intI & "Cel1").innerText,1)) = strText Then
            Call Result_onclick(intI)
            Exit For
        End If
    Next
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

Sub txtFactorName_onkeypress()
    Call TextBoxOnKeyPress(window.event.keyCode,"X")
End Sub

</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY ID=PageBody style="OVERFLOW: auto; POSITION: absolute; BACKGROUND-COLOR: <%=gstrPageColor%>" 
    bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0>
    
    <DIV id=Header
        style="POSITION: absolute;BORDER-STYLE: solid;BORDER-WIDTH: 1px;BORDER-COLOR: <%=gstrBorderColor%>;
        BACKGROUND-COLOR:<%=gstrBackColor%>;COLOR: black;HEIGHT: 40;WIDTH: 730;LEFT:0;TOP:10">
       
        <SPAN id=lblAppTitle
            style="POSITION: absolute;COLOR: <%=gstrAccentColor%>;FONT-SIZE: <%=gstrTitleFontSize%>;
            FONT-FAMILY: <%=gstrTitleFont%>;
            HEIGHT: 20;TOP: 5;LEFT: 7;WIDTH: 720;TEXT-ALIGN: center;FONT-WEIGHT: bold"><%=mstrPageTitle%></SPAN>
        <SPAN id=lblAppTitleHiLight
            style="POSITION: absolute;FONT-SIZE: <%=gstrTitleFontSize%>;
            FONT-FAMILY: <%=gstrTitleFont%>;COLOR: <%=gstrTitleColor%>;
            HEIGHT: 20;TOP: 4;LEFT: 6;WIDTH: 720;TEXT-ALIGN: center;FONT-WEIGHT: bold"><%=mstrPageTitle%></SPAN>
    </DIV>

    <DIV id=divSaving Class=ControlDiv style="WIDTH:730;TOP:51;height:420;LEFT:-1000">
        <SPAN id=lblSaving class=DefLabel style="POSITION:absolute;LEFT:10;font-size:12pt;WIDTH:250;TOP:32;BACKGROUN-COLOR:<%=gstrBackColor%>">Saving Record...Please Wait...</SPAN>
    </DIV>
    <DIV id=divPageFrame Class=ControlDiv style="WIDTH:730;TOP:51;height:420;LEFT:0">
        <SPAN id=lblFactors class=DefLabel style="POSITION:absolute;LEFT:10;font-size:10pt;WIDTH:250;TOP:2;BACKGROUN-COLOR:<%=gstrBackColor%>">Elements / Factors</SPAN>
        <DIV id=divFactors class=TableDivArea style="background-color:transparent;LEFT:10; WIDTH:250; TOP:20; HEIGHT:365"
            tabIndex=-1>
		    <Table ID=tblFactors Border=0 CellSpacing=0 tabindex=<%=LocalGetTabIndex()%> Style="POSITION:absolute;overflow: auto; TOP:0;left:0;width:230">
		        <tbody>
		        <%
		            mintRowID = 0
		            For Each moDictObj In mdctFactors
		                Response.Write "<tr id=tblRow" & mintRowID & " class=TableRow style=""cursor:hand"" onkeypress=Table_onkeyPress() onclick=Result_onclick(" & mintRowID & ")>" & vbCrLf
		                Response.Write "<td id=tblRow" & mintRowID & "Cel1 class=TableDetail style=""width:220"">" & Parse(mdctFactors(moDictObj),"^",1) & "</td>" & vbCrLf
		                Response.Write "</tr>" & vbCrLf
		                mstrHidden = mstrHidden & "<input type=hidden id=hidRowInfo" & mintRowID & " value=" & moDictObj & ">" & vbCrLf
		                mintRowID = mintRowID + 1
		            Next
		        %>
		        </tbody>
		    </Table>
		</DIV>
		<%Response.Write mstrHidden%>
		<SPAN id=lblFactorID class=DefLabel style="LEFT:265; WIDTH:40; TOP:2">
            ID
            <INPUT type=text id=txtFactorID style="LEFT:1; WIDTH:40;TOP:15;TEXT-ALIGN:center"
                tabIndex=-1 disabled NAME="txtFactorID">
        </SPAN>
		<SPAN id=lblFactor class=DefLabel style="LEFT:315; WIDTH:305; TOP:2">
            Factor Text
            <INPUT type=text id=txtFactorName style="LEFT:1; WIDTH:305;TOP:15;TEXT-ALIGN:LEFT"
                tabIndex=<%=LocalGetTabIndex%> maxlength=255 disabled NAME="txtFactorName">
        </SPAN>
		<SPAN id=lblLastUsed class=DefLabel style="LEFT:625; WIDTH:70; TOP:2">
            Last Used
            <INPUT type=text id=txtLastUsed style="LEFT:1; WIDTH:70;TOP:15;TEXT-ALIGN:center"
                tabIndex=-1 disabled NAME="txtLastUsed">
        </SPAN>
		<SPAN id=lblFactorLongName class=DefLabel style="LEFT:265; WIDTH:455; TOP:45;height:50">
            Description
            <TEXTAREA id=txtFactorLongName style="LEFT:1; WIDTH:455; TOP:15; HEIGHT:45; TEXT-ALIGN:left; padding-left:3; overflow:auto"
                tabIndex=<%=LocalGetTabIndex%> NAME="txtFactorLongName"></TEXTAREA>
        </SPAN>
 
        <SPAN id=lblCurrentLinks class=DefLabel style="POSITION:absolute;LEFT:265;font-size:10pt;WIDTH:50;TOP:110;BACKGROUN-COLOR:<%=gstrBackColor%>">Links</SPAN>
        <BUTTON class=DefBUTTON id=cmdAvailableLinks style="LEFT:320;POSITION: absolute;TOP:106;WIDTH:150;height:18" tabIndex=<%=LocalGetTabIndex%>>
            Show All Available Links
        </BUTTON>
        <SPAN id=lblLastUsed class=DefLabel style="POSITION:absolute;LEFT:-1515;font-size:9pt;WIDTH:70;TOP:110;BACKGROUN-COLOR:<%=gstrBackColor%>">Last Used</SPAN>
        <SPAN id=lblEndDate class=DefLabel style="POSITION:absolute;LEFT:-1585;font-size:9pt;WIDTH:70;TOP:110;BACKGROUN-COLOR:<%=gstrBackColor%>">End Date</SPAN>
        <DIV id=divCurrentLinks class=TableDivArea style="background-color:transparent;LEFT:265; WIDTH:455; TOP:125; HEIGHT:260" tabIndex=-1>
		    <Table ID=tblLinks Border=0 CellSpacing=0 tabindex=<%=LocalGetTabIndex()%> Style="POSITION:absolute;overflow: auto; TOP:0;left:0;width:435">
		        <tbody id=tblLinksBody>
		            <tr>
		                <td id=tbCell1><BUTTON id=BUTTON1>Delete</BUTTON></td>
		            </tr>
		        </tbody>
		    </Table>
		</DIV>
        <input type=Hidden id=hidLinkCount value=0 />
        <input type=Hidden id=hidLinkElmCount value=0 />
        <DIV id=divHoldvalues></DIV>
        
        <BUTTON class=DefBUTTON id=cmdDelete disabled
            style="LEFT: 65;POSITION: absolute; TOP: 390;HEIGHT: 20;WIDTH: 70" accessKey=R
            tabIndex=<%=LocalGetTabIndex%>>Delete
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdAdd
            style="LEFT: 145;POSITION: absolute;TOP: 390;HEIGHT: 20;WIDTH: 70" accessKey=R
            tabIndex=<%=LocalGetTabIndex%>>Add
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdEdit
            style="LEFT: 225;POSITION: absolute;TOP: 390;HEIGHT: 20;WIDTH: 70" accessKey=R
            tabIndex=<%=LocalGetTabIndex%>>Edit
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdCancel disabled 
           style="LEFT: 305;POSITION: absolute;TOP: 390;HEIGHT: 20;WIDTH: 70" accessKey=R
            tabIndex=<%=LocalGetTabIndex%>>Cancel
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdSave
            disabled style="LEFT: 390;POSITION: absolute;TOP: 390;HEIGHT: 20;WIDTH: 70" accessKey=R
            tabIndex=<%=LocalGetTabIndex%>>Save
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdClose style="LEFT: 645;POSITION: absolute;
                TOP: 390;HEIGHT: 20;WIDTH: 70" accessKey=R
            tabIndex=<%=LocalGetTabIndex%>>Close
        </BUTTON>
    </DIV>
</BODY>
<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="" ID=Form>
    <%
    Call CommonFormFields()
    WriteFormField "FormAction", ""
    WriteFormField "FactorID", 0
    WriteFormField "FactorName", ""
    WriteFormField "FactorLongName", ""
    WriteFormField "FactorIndexID", mintFactorIndexID
    WriteFormField "FactorList", ""
    WriteFormField "BuildCompleted", ""
   %>
</FORM>
</HTML>
<%
gadoCon.Close
Set gadoCon = Nothing
%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
