<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: FactorAddEditAssign.asp                                         '
'  Purpose: The primary admin data entry screen for maintaining the causal  '
'           factors for each eligibility element.                           '
'           This form is only available to admin users.                     '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim mstrPageTitle
Dim madoRs, madoRsFactors
Dim mstrAction
Dim mintProgramID, mintTabID, mstrProgramName, mintElementIndexID
Dim mdctElements, mdctFactors, mstrHidden
Dim moDictObj
Dim strElementRecord, strFactorList, strElementTitle, strFactorTitle
Dim mintRowID, mintTabIndex, mstrAllFactors
Dim mintReturnID, mstrMessage

mintTabIndex = 0
mstrAction = ReqForm("FormAction")

If Len(mstrAction) = 0 Then
    mstrAction = "Load"
    mintProgramID = Request.QueryString("ProgramID")
    mstrProgramName = Request.QueryString("ProgramName")
    mintTabID = Request.QueryString("TabID")
Else
    mintProgramID = ReqForm("ProgramID")
    mstrProgramName = ReqForm("ProgramName")
    mintTabID = ReqForm("TabID")
End If
strElementTitle = GetTabInfo(mintTabID, "Element")
strFactorTitle = GetTabInfo(mintTabID, "Factor")
mstrPageTitle = "Add/Edit " & mstrProgramName & " " & GetTabInfo(mintTabID,"Name") & " " & GetTabInfo(mintTabID,"Element") & "s"

Set madoRsFactors = Server.CreateObject("ADODB.Recordset")
' Factors
Set gadoCmd = GetAdoCmd("spCausalFactorList")
    madoRsFactors.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
Set gadoCmd = Nothing

mstrAllFactors = ""
mintRowID = 0
madoRsFactors.Sort = "fctShortName"
Do While Not madoRsFactors.EOF
    mstrAllFactors = mstrAllFactors & "<SPAN id=lblFactorID" & mintRowID & " onclick=Factor_OnClick(" & mintRowID & ") ondblclick=Factor_OnDblClick(" & mintRowID & ") class=DefLabel style=""overflow:TRUNCATE;cursor:hand;LEFT:1;WIDTH:2500;TOP:" & mintRowID*16 & ";font-weight:normal"">" & madoRsFactors.Fields("fctShortName").Value & "</SPAN>" & vbCrLf
    mstrAllFactors = mstrAllFactors & "<input type=hidden id=hidFactorRowInfo" & mintRowID & " value=" & madoRsFactors.Fields("fctID").Value & ">" & vbCrLf
    madoRsFactors.MoveNext
    mintRowID = mintRowID + 1
Loop
mstrAllFactors = mstrAllFactors & "<input type=hidden id=hidFactorTotal value=" & mintRowID-1 & ">" & vbCrLf

Select Case mstrAction
    Case "AddSave"
        Set gadoCmd = GetAdoCmd("spElementAdd")
            AddParmIn gadoCmd, "@elmProgramID", adInteger, 0, mintProgramID
            AddParmIn gadoCmd, "@elmTypeID", adInteger, 0, mintTabID
            AddParmIn gadoCmd, "@elmShortName", adVarChar, 250, ReqForm("ElementName")
            AddParmIn gadoCmd, "@elmEndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
            AddParmIn gadoCmd, "@elmIncludeInFull", adBoolean, 0, ReqForm("IncludeInFull")
            AddParmOut gadoCmd, "@elmID", adInteger, 0
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mintReturnID = gadoCmd.Parameters("@elmID").Value
            Select Case mintReturnID
                Case -1
                    mstrMessage = "Error encountered while trying to Add " & ReqForm("ElementName") & "." & vbCrLf
                Case Else
                    mstrMessage = ""
            End Select
    
            If mstrMessage = "" Then
                Call ProcessFactorConnections(gadoCmd.Parameters("@elmID").Value)
            End If
        Set gadoCmd = Nothing
    Case "EditSave"
        Set gadoCmd = GetAdoCmd("spElementUpd")
            AddParmIn gadoCmd, "@elmID", adInteger, 0, ReqForm("ElementID")
            AddParmIn gadoCmd, "@elmShortName", adVarChar, 250, ReqForm("ElementName")
            AddParmIn gadoCmd, "@elmEndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
            AddParmIn gadoCmd, "@elmIncludeInFull", adBoolean, 0, ReqForm("IncludeInFull")
            AddParmOut gadoCmd, "@ReturnID", adInteger, 0 
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mintReturnID = gadoCmd.Parameters("@ReturnID").Value
            Select Case mintReturnID
                Case -1
                    mstrMessage = ReqForm("ElementName") & " could not be end dated.  It has been used on a review." & vbCrLf
                Case -2
                    mstrMessage = "Error encountered while trying to delete " & ReqForm("ElementName") & "." & vbCrLf
                Case Else
                    mstrMessage = ""
            End Select
    
            If mstrMessage = "" Then
                Call ProcessFactorConnections(ReqForm("ElementID"))
            End If
        Set gadoCmd = Nothing
    Case "Delete"
        Set gadoCmd = GetAdoCmd("spElementDel")
            AddParmIn gadoCmd, "@elmID", adInteger, 0, ReqForm("ElementID")
            AddParmOut gadoCmd, "@ReturnID", adInteger, 0
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mintReturnID = gadoCmd.Parameters("@ReturnID").Value
            Select Case mintReturnID
                Case 0
                    mstrMessage = ""
                Case -1
                    mstrMessage = strElementTitle & " could not be deleted.  It has been used on a review."
                Case -2
                    mstrMessage = "Error encountered while trying to delete " & strElementTitle
            End Select
        Set gadoCmd = Nothing
End Select
madoRsFactors.Close
' Load string values that will be converted to client side arrays
Set mdctElements = CreateObject("Scripting.Dictionary")
Set mdctFactors = CreateObject("Scripting.Dictionary")
' Elements
Set madoRsFactors = Server.CreateObject("ADODB.Recordset")
Set madoRs = Server.CreateObject("ADODB.Recordset")
' Factors
Set gadoCmd = GetAdoCmd("spElementsFactorsEdit")
    AddParmIn gadoCmd, "@ProgramID", adInteger, 0, mintProgramID
    AddParmIn gadoCmd, "@TypeID", adInteger, 0, mintTabID
    madoRsFactors.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
Set gadoCmd = Nothing

Set gadoCmd = GetAdoCmd("spGetElements")
    AddParmIn gadoCmd, "@ProgramID", adInteger, 0, mintProgramID
    AddParmIn gadoCmd, "@TypeID", adInteger, 0, mintTabID
    madoRs.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
Set gadoCmd = Nothing

madoRs.Sort = "elmSortOrder"
mintElementIndexID = 0
Do While Not madoRs.EOF
    strElementRecord = madoRs.Fields("elmShortName").Value & "^" & _
        madoRs.Fields("elmStartDate").Value & "^" & _
        madoRs.Fields("elmEndDate").Value & "^" & _
        madoRs.Fields("elmIncludeInFull").Value & "^" & _
        madoRs.Fields("elmSortOrder").Value & "^"
    strFactorList = ""
    madoRsFactors.Filter = "elmID=" & madoRs.Fields("elmID").Value
    madoRsFactors.Sort = "elfSortOrder"
    strElementRecord = strElementRecord & madoRsFactors.Fields("ElementUsedLast").Value & "^"
    Do While Not madoRsFactors.EOF
        strFactorList = strFactorList & madoRsFactors.Fields("fctID").Value & "~" & madoRsFactors.Fields("FactorUsedLast").Value & "~" & madoRsFactors.Fields("elfEndDate").Value & "*"
        madoRsFactors.MoveNext
    Loop
    If strFactorList = "~~*" Then strFactorList = ""
    mdctElements.Add CLng(madoRs.Fields("elmID").Value), strElementRecord & strFactorList
    madoRs.MoveNext
Loop
madoRs.Close
madoRsFactors.Close

mintRowID = 0

Function GetTabInfo(intTabID, strType)
    Select Case CInt(intTabID)
        Case 1
            Select Case strType
                Case "Name"
                    GetTabInfo = "Action Integrity"
                Case "Element"
                    GetTabInfo = "Action"
                Case "Factor"
                    GetTabInfo = "Decision"
            End Select
        Case 2
            Select Case strType
                Case "Name"
                    GetTabInfo = ""
                Case "Element"
                    GetTabInfo = "Element"
                Case "Factor"
                    GetTabInfo = "Causal Factor"
            End Select
        Case 3
            Select Case strType
                Case "Name"
                    GetTabInfo = "Information Gathering"
                Case "Element"
                    GetTabInfo = "Question"
                Case "Factor"
                    GetTabInfo = "Answer"
            End Select
    End Select
End Function

Function LocalGetTabIndex()
    LocalGetTabIndex = mintTabIndex
    mintTabIndex = CInt(mintTabIndex) + 1
End Function

Sub ProcessFactorConnections(intElementID)
    Dim intI, strRecord, strFactorName
    
    For intI = 1 To 100
        strRecord = Parse(ReqForm("FactorList"),"|",intI)
        If strRecord = "" Then Exit For
        Set gadoCmd = GetAdoCmd("spElementFactorLinkUpd")
            AddParmIn gadoCmd, "@ElementID", adInteger, 0, intElementID
            AddParmIn gadoCmd, "@FactorID", adInteger, 0, Parse(strRecord,"^",2)
            AddParmIn gadoCmd, "@SortOrder", adInteger, 0, intI-1
            AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Parse(strRecord,"^",3))
            AddParmIn gadoCmd, "@Action", adChar, 1, Parse(strRecord,"^",1)
            AddParmOut gadoCmd, "@ReturnID", adInteger, 0
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
            mintReturnID = gadoCmd.Parameters("@ReturnID").Value
            madoRsFactors.Filter = "fctID=" & Parse(strRecord,"^",2)
            If madoRsFactors.RecordCount=1 Then
                strFactorName = madoRsFactors.Fields("fctShortName").value
            Else
                strFactorName = "[unknown, ID=" & Parse(strRecord,"^",2) & "]"
            End If
            Select Case mintReturnID
                Case 0
                Case -1
                    mstrMessage = mstrMessage & GetTabInfo(mintTabID, "Factor") & " " & strFactorName & " - Link could not be deleted/end dated.  It has been used on a review." & vbCrLf
                Case -2
                    mstrMessage = mstrMessage & GetTabInfo(mintTabID, "Factor") & " " & strFactorName & " - Error encountered while trying to delete link." & vbCrLf
            End Select
    Next
End Sub
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
Dim mdctElements
Dim mintFactorRowID
Dim mintSelFactorRowID

Sub window_onload()
    Dim intI
    If "<%=mstrMessage%>" <> "" Then
        
        MsgBox "<%=mstrMessage%>", vbOkOnly, "Case Review Maintenance"
    End If
    Set mdctElements = CreateObject("Scripting.Dictionary")
    <%
    For Each moDictObj In mdctElements
        Response.Write "mdctElements.Add CLng(" & moDictObj & "), """ & mdctElements(moDictObj) & """" & vbCrLf
    Next
    %>
    If "<%=mintTabID%>" = "3" Then
        lblFactors.style.left = -1000
        divFactors.style.left = -1000
        lblSelFactors.style.left = -1000
        divSelFactors.style.left = -1000
        cmdMoveFacUp.style.left = -1000
        cmdMoveFacDown.style.left = -1000
    End If
    Form.ElementIndexID.Value = 0
    mintFactorRowID = -1
    mintSelFactorRowID = -1
    Call Result_onclick(0)
    window.parent.divElementsLoading.style.left = -1000
    window.parent.divElements.style.left = 0
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
        Form.ElementID.value = 0
        Form.FormAction.Value = "AddSave"
    ElseIf Form.FormAction.Value = "Edit" Then
        Form.ElementID.value = document.all("hidRowInfo" & Form.ElementIndexID.Value).value
        Form.FormAction.Value = "EditSave"
    End If
    Call FillForm()

    Form.submit
End Sub

Sub cmdEdit_onclick()
    Call DisableControls("Edit")
    Call FillScreen()
    txtElementName.focus
    Form.FormAction.value = "Edit"
End Sub

Sub cmdCancel_onclick()
    Call DisableControls("Cancel")
    Call ClearScreen()
    Call FillScreen()
End Sub

Sub cmdAdd_onclick()
    Call DisableControls("Edit")
    Call ClearScreen()
    txtStartDate.value = FormatDateTime(DateAdd("d",1,Now()),2)
    txtElementName.focus
    chkIncludeInFull.checked = False
    Form.FormAction.value = "Add"
End Sub

Sub cmdDelete_onclick()
    Dim strRecord
    Dim dtmLastUsed
    Dim intResp
    
    strRecord = mdctElements(CLng(document.all("hidRowInfo" & Form.ElementIndexID.Value).value))
    dtmLastUsed = Parse(strRecord,"^",6)
    If dtmLastUsed <> "" Then
        Msgbox "<%=strElementTitle%> has been used in a review and cannot be deleted.", vbOkOnly, "Case Review Maintenance"
        Exit Sub
    Else
        intResp = MsgBox("Delete the <%=strElementTitle%>?", vbQuestion + vbYesNo, "Delete")
        If intResp = vbNo Then Exit Sub
    End If
    
    Form.ElementID.value = document.all("hidRowInfo" & Form.ElementIndexID.Value).value
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
    If Form.FormAction.value = "Edit" And txtEndDate.value <> "" Then
        strRecord = mdctElements(CLng(document.all("hidRowInfo" & Form.ElementIndexID.Value).value))
        dtmLastUsed = Parse(strRecord,"^",6)
        
        If dtmLastUsed <> "" Then
            If CDate(txtEndDate.value) < CDate(dtmLastUsed) Then
                strMessage = "<%=strElementTitle%> was last used in a review on " & FormatDateTime(dtmLastUsed,2) & "." & vbCrLf
                strMessage = strMessage & "End Date cannot be before " & FormatDateTime(dtmLastUsed,2) & "."
                Validate = strMessage
                Exit Function
            End If
        End If
    ElseIf Form.FormAction.value = "Add" And txtEndDate.value <> "" Then
        If CDate(txtEndDate.value) <= CDate(Now()) Then
            strMessage = "End Date cannot be before " & FormatDateTime(DateAdd("d",1,Now()),2) & "."
            Validate = strMessage
            Exit Function
        End If
    End If
    
    If Trim(txtElementName.value) = "" Then
        strMessage = "<%=strElementTitle%> text cannot be blank."
        Validate = strMessage
        Exit Function
    End If
    
    If Form.FormAction.value = "Add" Then
        intSelectedID = -1
    Else
        intSelectedID = Form.ElementIndexID.value
    End If
    
    For intI = 0 To tblElements.Rows.Length - 1
        If UCase(tblElements.Rows(intI).Cells(0).innerText) = Trim(UCase(txtElementName.value)) And CInt(intI) <> CInt(intSelectedID) Then
            strMessage = txtElementName.value & " cannot be the text of this <%=strElementTitle%>.  It is identical to another <%=strElementTitle%> of this Function."
            Validate = strMessage
            Exit Function
        End If
    Next
End Function

Sub FillForm()
    Form.ElementName.value = txtElementName.value
    Form.EndDate.value = txtEndDate.value
    If chkIncludeInFull.checked = True Then
        Form.IncludeInFull.value = 1
    Else
        Form.IncludeInFull.value = 0
    End If
    Call BuildFactorList()
End Sub

Sub BuildFactorList()
    Dim dctOldFactorList, dctNewFactorList
    Dim strOldFactorList, strOldFactor
    Dim intI, strFactorList, dtmEndDate, dtmOldEndDate
    Dim oNewFactor, oOldFactor

    Set dctOldFactorList = CreateObject("Scripting.Dictionary")
    Set dctNewFactorList = CreateObject("Scripting.Dictionary")
    strOldFactorList = ""
    If Form.FormAction.value = "EditSave" Then
        strOldFactorList = Parse(mdctElements(CLng(document.all("hidRowInfo" & Form.ElementIndexID.Value).value)),"^",7)
    End If
    For intI = 1 To 100
        strOldFactor = Parse(strOldFactorList,"*",intI)
        If strOldFactor = "" Then Exit For
        dctOldFactorList.Add CLng(Parse(strOldFactor,"~",1)), Parse(strOldFactor,"~",2)
    Next
    
    If hidSelFactorTotal.value = "" Then hidSelFactorTotal.value = "-1"
    
    If CInt(hidSelFactorTotal.value) >= 0 Then
        For intI = 0 To CInt(hidSelFactorTotal.value)
            dtmEndDate = ""
            If Parse(document.all("lblSelFactorID" & intI).innerText," (Ended ",2) <> "" Then
                dtmEndDate = Parse(document.all("lblSelFactorID" & intI).innerText," (Ended ",2)
                dtmEndDate = Trim(Replace(dtmEndDate,")",""))
            End If
            dctNewFactorList.Add CLng(document.all("hidSelFactorRowInfo" & intI).value), dtmEndDate
        Next
    End If

    'dctNewFactorList contains all factors currently selected.  
    'dctOldFactorList contains all factors currently in database
    'All entries in dctNew... will be processed as adds/edits.  If any factors exist in dctOld...
    'but do not exist in dctNew..., they will be entered as deletes.
    strFactorList = ""
    For Each oNewFactor In dctNewFactorList
        'All factors added to last with Action of A - stored procedure will make INSERT or UPDATE decision
        strFactorList = strFactorList & "A^" & oNewFactor & "^" & dctNewFactorList(oNewFactor) & "|"
    Next
    For Each oOldFactor In dctOldFactorList
        If Not dctNewFactorList.Exists(oOldFactor) Then
            strFactorList = strFactorList & "D^" & oOldFactor & "^" & dctOldFactorList(oOldFactor) & "|"
        End If
    Next
    Form.FactorList.value = strFactorList
End Sub

Sub cmdClose_onclick()
    Call window.parent.ShowDiv("SubMenu")
End Sub

Sub Result_onclick(intRowID)
    Dim strRow
    
    If tblElements.Rows.length = 0 Then Exit Sub
    
    If cmdAdd.disabled = True Then 
        MsgBox "Screen is in Edit mode.  Save or Cancel current record before selecting another <%=strElementTitle%>", vbOkOnly, "Case Review Maintenance"
        Exit Sub
    End If
    
    If IsNumeric(Form.ElementIndexID.Value) Then
        strRow = "tblRow" & Form.ElementIndexID.Value
        tblElements.Rows(strRow).className = "TableRow"
        tblElements.Rows(strRow).cells(0).tabindex = -1
    End If

    strRow = "tblRow" & intRowID
    tblElements.Rows(strRow).className = "TableSelectedRow"
    tblElements.Rows(strRow).cells(0).focus
    tblElements.Rows(strRow).cells(0).tabindex = 9

    Form.ElementIndexID.Value = intRowID
    Call FillScreen()
End Sub

Sub FillScreen()
    Dim strRecord
    If tblElements.Rows.length = 0 Then Exit Sub

    strRecord = mdctElements(CLng(document.all("hidRowInfo" & Form.ElementIndexID.Value).value))

    txtElementName.value = Parse(strRecord,"^",1)
    txtStartDate.value = Parse(strRecord,"^",2)
    txtEndDate.value = Parse(strRecord,"^",3)
    If Parse(strRecord,"^",4) = "True" Then
        chkIncludeInFull.checked = True
    Else
        chkIncludeInFull.checked = False
    End If
    Call DisplayFactors(Parse(strRecord,"^",7))
End Sub

Sub ClearScreen()
    txtElementName.value = ""
    txtStartDate.value = ""
    txtEndDate.value = ""
    chkIncludeInFull.checked = False

    Call DisplayFactors("")
End Sub

Sub DisableControls(strMode)
    Select Case strMode
        Case "Edit"
            cmdDelete.disabled = True
            cmdAdd.disabled = True
            cmdEdit.disabled = True
            cmdCancel.disabled = False
            cmdSave.disabled = False
            cmdMoveFacUp.disabled = False
            cmdMoveFacDown.disabled = False
            cmdMoveElmUp.disabled = True
            cmdMoveElmDown.disabled = True
            txtElementName.disabled = False
            txtEndDate.disabled = False
            chkIncludeInFull.disabled = False
            lblIncludeInFull.style.cursor = "hand"
        Case Else
            cmdDelete.disabled = False
            cmdAdd.disabled = False
            cmdEdit.disabled = False
            cmdCancel.disabled = True
            cmdSave.disabled = True
            cmdMoveFacUp.disabled = True
            cmdMoveFacDown.disabled = True
            cmdMoveElmUp.disabled = False
            cmdMoveElmDown.disabled = False
            txtElementName.disabled = True
            txtEndDate.disabled = True
            chkIncludeInFull.disabled = True
            lblIncludeInFull.style.cursor = "auto"
    End Select
End Sub

Sub EndTheFactor(intRowID)
    Dim intI, intJ, intSelRowID
    Dim strFactor
    Dim dtmEndDate
    
    intSelRowID = hidSelFactorTotal.value

    intI = 0
    strFactor = ""
    For intJ = 0 To intSelRowID
        If CLng(document.all("hidSelFactorRowInfo" & intJ).value) = CLng(document.all("hidFactorRowInfo" & intRowID).value) Then
            If document.all("hidSelFactorLastUsed" & intJ).value <> "" Then
                dtmEndDate = InputBox("This <%=strFactorTitle%> was last used in a review on " & document.all("hidSelFactorLastUsed" & intJ).value & "." & vbCrLf & "Please enter an End Date that is on or after " & document.all("hidSelFactorLastUsed" & intJ).value & ".", "End Date", document.all("hidSelFactorLastUsed" & intJ).value)
                If dtmEndDate <> "" Then
                    If Not IsDate(dtmEndDate) Then
                        MsgBox "End Date must be a valid date.", vbOkOnly, "Case Review"
                        Call Factor_OnClick(intRowID)
                        Exit Sub
                    End If
                    If CDate(dtmEndDate) < CDate(document.all("hidSelFactorLastUsed" & intJ).value) Then
                        MsgBox "End Date cannot be before date last used.", vbOkOnly, "Case Review"
                        Exit Sub
                    End If
                    document.all("lblFactorID" & intRowID).style.fontWeight = "normal"
                    document.all("lblFactorID" & intRowID).style.color = "black"
                    dtmEndDate = " (Ended " & dtmEndDate & ")"
                End If
                strFactor = strFactor & "<SPAN id=lblSelFactorID" & intI & " onclick=SelFactor_OnClick(" & intI & ") ondblclick=SelFactor_OnDblClick(" & intI & ") class=DefLabel style=""cursor:hand;color:black;LEFT:1;WIDTH:1500;TOP:" & intI*16 & """>" & document.all("lblSelFactorID" & intJ).innerText & dtmEndDate & "</SPAN>" & vbCrLf
                strFactor = strFactor & "<input type=hidden id=hidSelFactorRowInfo" & intI & " value=" & document.all("hidSelFactorRowInfo" & intJ).value & ">" & vbCrLf
                strFactor = strFactor & "<input type=hidden id=hidSelFactorLastUsed" & intI & " value=""" & document.all("hidSelFactorLastUsed" & intJ).value & """>" & vbCrLf
                intI = intI + 1
            Else
                document.all("lblFactorID" & intRowID).style.fontWeight = "normal"
                document.all("lblFactorID" & intRowID).style.color = "black"
            End If
        Else
            strFactor = strFactor & "<SPAN id=lblSelFactorID" & intI & " onclick=SelFactor_OnClick(" & intI & ") ondblclick=SelFactor_OnDblClick(" & intI & ") class=DefLabel style=""cursor:hand;color:black;LEFT:1;WIDTH:1500;TOP:" & intI*16 & """>" & document.all("lblSelFactorID" & intJ).innerText & "</SPAN>" & vbCrLf
            strFactor = strFactor & "<input type=hidden id=hidSelFactorRowInfo" & intI & " value=" & document.all("hidSelFactorRowInfo" & intJ).value & ">" & vbCrLf
            strFactor = strFactor & "<input type=hidden id=hidSelFactorLastUsed" & intI & " value=""" & document.all("hidSelFactorLastUsed" & intJ).value & """>" & vbCrLf
            intI = intI + 1
        End If
    Next
    strFactor = strFactor & "<input type=hidden id=hidSelFactorTotal value=" & CInt(intI) - 1 & ">" & vbCrLf
    divSelFactors.innerHTML = strFactor
End Sub

Sub Factor_OnClick(intRowID)
    If cmdEdit.disabled = False Then Exit Sub
    If mintFactorRowID >= 0 Then
        document.all("lblFactorID" & mintFactorRowID).style.backgroundcolor = "<%=gstrBackColor%>"
    End If
    document.all("lblFactorID" & intRowID).style.backgroundcolor = "darkolivegreen"
    mintFactorRowID = intRowID
End Sub

Sub Factor_OnDblClick(intRowID)
    Dim intSelRowID
    
    If cmdEdit.disabled = False Then Exit Sub
    
    If document.all("lblFactorID" & intRowID).style.fontWeight = "bold" Then
        'Remove from selected list
        Call EndTheFactor(intRowID)
    Else
        'Add to selected list
        Call AddTheFactor(intRowID)
    End If
End Sub

Sub SelFactor_OnClick(intSelRowID)
    If cmdEdit.disabled = False Then Exit Sub
    If mintSelFactorRowID >= 0 Then
        document.all("lblSelFactorID" & mintSelFactorRowID).style.backgroundcolor = "<%=gstrBackColor%>"
    End If
    document.all("lblSelFactorID" & intSelRowID).style.backgroundcolor = "darkolivegreen"
    mintSelFactorRowID = intSelRowID
End Sub

Sub SelFactor_OnDblClick(intSelRowID)
    Dim intI, intFacRowID
    
    If cmdEdit.disabled = False Then Exit Sub
    
    For intI = 0 To hidFactorTotal.value
        If CLng(document.all("hidFactorRowInfo" & intI).value) = CLng(document.all("hidSelFactorRowInfo" & intSelRowID).value) Then
            intFacRowID = intI
            Exit For
        End If
    Next

    If InStr(document.all("lblSelFactorID" & intSelRowID).innerText,"(Ended") > 0 Then
        Call AddTheFactor(intFacRowID)
    Else
        Call EndTheFactor(intFacRowID)
        If CInt(hidSelFactorTotal.value) >= 0 Then
            mintSelFactorRowID = 0
            Call SelFactor_OnClick(0)
        End If
    End If
End Sub

Sub AddTheFactor(intRowID)
    Dim intSelID, intTotalSelRows, intJ, intI
    Dim dtmEndDate
    Dim strFactor
    
    intTotalSelRows = hidSelFactorTotal.value
    intSelID = -1

    For intI = 0 To intTotalSelRows
        If CLng(document.all("hidSelFactorRowInfo" & intI).value) = CLng(document.all("hidFactorRowInfo" & intRowID).value) Then
            ' If factor already exists in selected list, it has been ended
            intSelID = intI
            Exit For
        End If
    Next
    
    If intSelID = -1 Then
        document.all("lblFactorID" & intRowID).style.fontWeight = "bold"
        document.all("lblFactorID" & intRowID).style.color = "blue"
        intJ = CInt(intTotalSelRows) + 1
        strFactor = "<SPAN id=lblSelFactorID" & intJ & " onclick=SelFactor_OnClick(" & intJ & ") ondblclick=SelFactor_OnDblClick(" & intJ & ") class=DefLabel style=""cursor:hand;color:black;LEFT:1;WIDTH:1500;TOP:" & intJ*16 & """>" & document.all("lblFactorID" & intRowID).innerText & "</SPAN>" & vbCrLf
        strFactor = strFactor & "<input type=hidden id=hidSelFactorRowInfo" & intJ & " value=" & document.all("hidFactorRowInfo" & intRowID).value & ">" & vbCrLf
        strFactor = strFactor & "<input type=hidden id=hidSelFactorLastUsed" & intJ & " value="""">" & vbCrLf
        hidSelFactorTotal.value = intJ
        divSelFactors.innerHTML = divSelFactors.innerHTML & strFactor
    Else
        intI = MsgBox("<%=strFactorTitle%> has been previously ended.  If you continue, the <%=strFactorTitle%> will be un-ended." & vbCrLf & vbCrLf & "Do you wish to continue?", vbYesNo, "Case Review")
        If intI = vbYes Then
            document.all("lblFactorID" & intRowID).style.fontWeight = "bold"
            document.all("lblFactorID" & intRowID).style.color = "blue"
            intJ = CInt(intTotalSelRows) + 1
            document.all("lblSelFactorID" & intSelID).innerText = Trim(Parse(document.all("lblSelFactorID" & intSelID).innerText,"(Ended",1))
        End If
    End If
End Sub

Sub cmdMoveElmUp_onclick()
    Call MoveElement("Up")
End Sub

Sub cmdMoveElmDown_onclick()
    Call MoveElement("Down")
End Sub

Sub MoveElement(strDir)
    Dim intTempRowID, strTempText, intTempID
    Dim dctReturn, oReturn
    
    If tblElements.Rows.length = 0 Then Exit Sub
    
    If strDir = "Up" Then
        If CInt(Form.ElementIndexID.Value) = 0 Then Exit Sub
        intTempRowID = Form.ElementIndexID.Value - 1
    Else
        If CInt(Form.ElementIndexID.Value) = CInt(tblElements.Rows.length) - 1 Then Exit Sub
        intTempRowID = Form.ElementIndexID.Value + 1
    End If

    strTempText = tblElements.Rows(intTempRowID).Cells(0).innerText
    intTempID = document.all("hidRowInfo" & intTempRowID).value
    
    tblElements.Rows(intTempRowID).Cells(0).innerText = tblElements.Rows(CInt(Form.ElementIndexID.Value)).Cells(0).innerText
    document.all("hidRowInfo" & intTempRowID).value = document.all("hidRowInfo" & Form.ElementIndexID.Value).value

    tblElements.Rows(CInt(Form.ElementIndexID.Value)).Cells(0).innerText = strTempText
    document.all("hidRowInfo" & Form.ElementIndexID.Value).value = intTempID

    Call Result_onclick(intTempRowID)
    Set dctReturn = window.showModalDialog("ElementSort.asp?Action=Sort&ElmID1=" & document.all("hidRowInfo" & Form.ElementIndexID.Value).value & "&ElmID2=" & intTempID, , "dialogWidth:210px;dialogHeight:120px;scrollbars:no;center:yes;border:thin;help:no;status:no")
    If dctReturn.Exists(1) Then
        MsgBox dctReturn(1)
    End If
End Sub

Sub cmdMoveFacUp_onclick()
    Call MoveFactor("Up")
End Sub

Sub cmdMoveFacDown_onclick()
    Call MoveFactor("Down")
End Sub

Sub MoveFactor(strDir)
    Dim intTempRowID, strTempText, intTempID, dtmTempLastUsed
    
    If CInt(mintSelFactorRowID) < 0 Then Exit Sub
    
    If strDir = "Up" Then
        If CInt(mintSelFactorRowID) = 0 Then Exit Sub
        intTempRowID = mintSelFactorRowID - 1
    Else
        If CInt(mintSelFactorRowID) = CInt(hidSelFactorTotal.value) Then Exit Sub
        intTempRowID = mintSelFactorRowID + 1
    End If
    strTempText = document.all("lblSelFactorID" & intTempRowID).innerText
    intTempID = document.all("hidSelFactorRowInfo" & intTempRowID).value
    dtmTempLastUsed = document.all("hidSelFactorLastUsed" & intTempRowID).value
    
    document.all("lblSelFactorID" & intTempRowID).innerText = document.all("lblSelFactorID" & mintSelFactorRowID).innerText
    document.all("hidSelFactorRowInfo" & intTempRowID).value = document.all("hidSelFactorRowInfo" & mintSelFactorRowID).value
    document.all("hidSelFactorLastUsed" & intTempRowID).value = document.all("hidSelFactorLastUsed" & mintSelFactorRowID).value

    document.all("lblSelFactorID" & mintSelFactorRowID).innerText = strTempText
    document.all("hidSelFactorRowInfo" & mintSelFactorRowID).value = intTempID
    document.all("hidSelFactorLastUsed" & mintSelFactorRowID).value = dtmTempLastUsed
    
    Call SelFactor_OnClick(intTempRowID)
End Sub

Sub DisplayFactors(strFactors)
    Dim oDictObj
    Dim intI, strColor, strFont, intJ
    Dim strSelColor, strNotColor
    Dim strSelFactors
    Dim dctFactors, strRecord, dtmEndDate, dtmLastUsed, strName
    
    If cmdEdit.disabled = True Then 
        strSelColor = "blue"
        strNotColor = "black"
    Else
        strSelColor = "gray"
        strNotColor = "gray"
    End If
    If strFactors = "~~*" Then strFactors = ""
    strSelFactors = ""
    
    Set dctFactors = CreateObject("Scripting.Dictionary")
    For intI = 1 To 100
        strRecord = Parse(strFactors,"*",intI)
        If strRecord = "" Then Exit For
        dctFactors.Add CLng(Parse(strRecord,"~",1)),Parse(strRecord,"~",2) & "^" & Parse(strRecord,"~",3)
    Next
    
    For intI = 0 To hidFactorTotal.value
        If dctFactors.Exists(CLng(document.all("hidFactorRowInfo" & intI).value)) Then
            If dtmEndDate <> "" Then 
                strColor = strNotColor
                strFont = "normal"
            Else
                strColor = strSelColor
                strFont = "bold"
            End If
            dctFactors(CLng(document.all("hidFactorRowInfo" & intI).value)) = dctFactors(CLng(document.all("hidFactorRowInfo" & intI).value)) & "^" & document.all("lblFactorID" & intI).innerText
        Else
            strColor = strNotColor
            strFont = "normal"
        End If
        document.all("lblFactorID" & intI).style.fontWeight = strFont
        document.all("lblFactorID" & intI).style.color = strColor
    Next
    
    intJ = 0
    For Each oDictObj In dctFactors
        strRecord = dctFactors(oDictObj)
        dtmLastUsed = Parse(strRecord,"^",1)
        dtmEndDate = Parse(strRecord,"^",2)
        If dtmEndDate <> "" Then
            dtmEndDate = " (Ended " & dtmEndDate & ")"
        End If
        strName = Parse(strRecord,"^",3)
        strSelFactors = strSelFactors & "<SPAN id=lblSelFactorID" & intJ & " onclick=SelFactor_OnClick(" & intJ & ") ondblclick=SelFactor_OnDblClick(" & intJ & ") class=DefLabel style=""cursor:hand;color:" & strNotColor & ";LEFT:1;WIDTH:1500;TOP:" & intJ*16 & """>" & strName & dtmEndDate & "</SPAN>" & vbCrLf
        strSelFactors = strSelFactors & "<input type=hidden id=hidSelFactorRowInfo" & intJ & " value=" & oDictObj & ">" & vbCrLf
        strSelFactors = strSelFactors & "<input type=hidden id=hidSelFactorLastUsed" & intJ & " value=""" & dtmLastUsed & """>" & vbCrLf
        intJ = intJ + 1
    Next
    
    strSelFactors = strSelFactors & "<input type=hidden id=hidSelFactorTotal value=" & intJ-1 & ">" & vbCrLf
    
    divSelFactors.innerHTML = strSelFactors
End Sub

Sub ToggleCheckBox(oControl)
    If oControl.disabled = True Then Exit Sub
    oControl.checked = Not oControl.checked
End Sub

Sub Date_onkeypress(ctlDate)
    If ctlDate.value = "(MM/DD/YYYY)" Then
        ctlDate.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub

Sub Date_onblur(ctlDate)
    Dim intI
    
    If Trim(ctlDate.value) = "(MM/DD/YYYY)" Then
        ctlDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(ctlDate.value) Then
        MsgBox "The End Date must be a valid date - MM/DD/YYYY.", vbInformation, "Case Review Maintenance"
        ctlDate.focus
    End If
    
    If IsDate(ctlDate.value) Then
        If CDate(ctlDate.value) <= CDate(FormatDateTime(Now(),2)) Then
            MsgBox "The End Date must be after " &  FormatDateTime(Now(),2) & ".", vbInformation, "Case Review Maintenance"
            ctlDate.focus
        End If
    End If
End Sub

Sub Date_onfocus(ctlDate)
    If Trim(ctlDate.value) = "" Then
        ctlDate.value = "(MM/DD/YYYY)"
    End If
    ctlDate.select
End Sub

Sub divFactors_onkeypress()
    If CInt(hidFactorTotal.value) <= 0 Then Exit Sub
    Call GoToRow(Chr(window.event.keyCode))
End Sub

Sub GoToRow(strText)
    Dim intI
    strText = UCase(strText)
    
    For intI = 0 To hidFactorTotal.value
        If UCase(Left(document.all("lblFactorID" & intI).innerText,1)) = strText Then
            document.all("lblFactorID" & intI).focus
            Call Factor_OnClick(intI)
            Exit For
        End If
    Next
End Sub

Sub txtElementName_onkeypress()
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

    <DIV id=divSaving Class=ControlDiv style="WIDTH:730;TOP:51;height:410;LEFT:0;cursor:wait">
        <BR><BR><BR><BR><Center>Checking Existing Element / Causal Factor Links to Reviews...Please Wait</Center>
    </DIV>
    <DIV id=divPageFrame Class=ControlDiv style="WIDTH:730;TOP:51;height:420;LEFT:0">
        <SPAN id=lblElements class=DefLabel style="POSITION:absolute;LEFT:10;font-size:10pt;WIDTH:250;TOP:2;BACKGROUN-COLOR:<%=gstrBackColor%>"><%=strElementTitle%>s</SPAN>
        <DIV id=divElements class=TableDivArea style="background-color:transparent;LEFT:10; WIDTH:250; TOP:20; HEIGHT:365"
            tabIndex=-1>
		    <Table ID=tblElements Border=0 CellSpacing=0 tabindex=<%=LocalGetTabIndex()%> Style="POSITION:absolute;overflow: auto; TOP:0;left:0;width:230">
		        <tbody>
		        <%
		            mintRowID = 0
		            For Each moDictObj In mdctElements
		                Response.Write "<tr id=tblRow" & mintRowID & " class=TableRow style=""cursor:hand"" onclick=Result_onclick(" & mintRowID & ")>" & vbCrLf
		                Response.Write "<td id=tblRow" & mintRowID & "Cel1 class=TableDetail style=""width:220"">" & Parse(mdctElements(moDictObj),"^",1) & "</td>" & vbCrLf
		                Response.Write "</tr>" & vbCrLf
		                mstrHidden = mstrHidden & "<input type=hidden id=hidRowInfo" & mintRowID & " value=" & moDictObj & ">" & vbCrLf
		                mintRowID = mintRowID + 1
		            Next
		        %>
		        </tbody>
		    </Table>
		</DIV>
		<%Response.Write mstrHidden%>
        <BUTTON class=DefBUTTON id=cmdMoveElmUp disabled 
           style="LEFT:265;POSITION:absolute;TOP:150;HEIGHT:20;WIDTH:20"
            tabIndex=<%=LocalGetTabIndex%>>/\
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdMoveElmDown disabled 
           style="LEFT:265;POSITION:absolute;TOP:190;HEIGHT:20;WIDTH:20"
            tabIndex=<%=LocalGetTabIndex%>>\/
        </BUTTON>
		<SPAN id=lblElement class=DefLabel style="LEFT:300; WIDTH:420; TOP:2">
            <%=strElementTitle%> Text:
            <INPUT type=text id=txtElementName style="LEFT:1; WIDTH:420;TOP:15;TEXT-ALIGN:LEFT"
                tabIndex=<%=LocalGetTabIndex%> maxlength=255 disabled NAME="txtElementName">
        </SPAN>
		<SPAN id=lblStartDate class=DefLabel style="LEFT:300; WIDTH:100; TOP:40">
            Start Date:
            <INPUT type=text id=txtStartDate style="LEFT:1; WIDTH:70;TOP:15;TEXT-ALIGN:center"
                onkepress="Date_onkeypress(txtStartDate)" onblur="Date_onblur(txtStartDate)" onfocus="Date_onfocus(txtStartDate)"
                tabIndex=<%=LocalGetTabIndex%> maxlength=10 disabled NAME="txtStartDate">
        </SPAN>
		<SPAN id=lblEndDate class=DefLabel style="LEFT:400; WIDTH:100; TOP:40">
            End Date:
            <INPUT type=text id=txtEndDate style="LEFT:1; WIDTH:70;TOP:15;TEXT-ALIGN:center"
                onkepress="Date_onkeypress(txtEndDate)" onblur="Date_onblur(txtEndDate)" onfocus="Date_onfocus(txtEndDate)"
                tabIndex=<%=LocalGetTabIndex%> maxlength=10 disabled NAME="txtEndDate">
        </SPAN>
        <INPUT id=chkIncludeInFull type=checkbox title="Include this <%=strElementTitle%> in a Full Review"
            style="LEFT:500; WIDTH:20; TOP:50;display:none" disabled NAME="chkIncludeInFull">
        <SPAN id=lblIncludeInFull class=DefLabel style="TOP:52; LEFT:525; cursor:auto;display:none"
            onclick="ToggleCheckBox(chkIncludeInFull)">
            Include <%=strElementTitle%> in Full Review
        </SPAN>
        <SPAN id=lblFactors class=DefLabel style="TOP:75; LEFT:300;">
            <b>Available <%=strFactorTitle%>s</b>
        </SPAN>
        <DIV id=divFactors class=DefPageFrame style="Overflow:auto;top:90; LEFT:300; HEIGHT:170; WIDTH:420; BORDER-STYLE:thin">
            <%=mstrAllFactors%>
        </DIV>
        <SPAN id=lblSelFactors class=DefLabel style="TOP:265; LEFT:300;">
            <b>Selected <%=strFactorTitle%>s</b>
        </SPAN>
        <DIV id=divSelFactors class=DefPageFrame style="Overflow:auto;top:280; LEFT:300; HEIGHT:105; WIDTH:390; BORDER-STYLE:thin">
        </DIV>
        <BUTTON class=DefBUTTON id=cmdMoveFacUp disabled 
           style="LEFT:692;POSITION:absolute;TOP:300;HEIGHT:20;WIDTH:20"
            tabIndex=<%=LocalGetTabIndex%>>/\
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdMoveFacDown disabled 
           style="LEFT:692;POSITION:absolute;TOP:340;HEIGHT:20;WIDTH:20"
            tabIndex=<%=LocalGetTabIndex%>>\/
        </BUTTON>
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
    WriteFormField "ProgramID", mintProgramID
    WriteFormField "ProgramName", mstrProgramName
    WriteFormField "TabID", mintTabID
    WriteFormField "ElementID", 0
    WriteFormField "ElementName", ""
    WriteFormField "EndDate", ""
    'WriteFormField "SortOrder", ""
    WriteFormField "IncludeInFull", 1
    WriteFormField "ElementIndexID", mintElementIndexID
    WriteFormField "FactorList", ""
    'WriteFormField "ElementID", mintElementID
    'WriteFormField "ElementName", mstrElementName
    'WriteFormField "EndDate", mdtmEndDate
    'WriteFormField "IncludeInFull", mintIncludeInFull
   %>
</FORM>
</HTML>
<%
gadoCon.Close
Set gadoCon = Nothing
%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
