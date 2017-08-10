<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>

<!--#include file="IncCnn.asp"-->

<%
Dim adRsResults
Dim adCmd
Dim lngCasID
Dim dtmReviewDate
Dim dtmReviewDateEnd
Dim strCaseNumber
Dim strSubmitted
Dim lngResponse
Dim strReviewer
Dim strLoad
Dim strProgramsSelected
Dim intLine
Dim blnUserAdmin
Dim blnUserQA
Dim strParmList
Dim mstrUserID, mlngAliasID
Dim strSupervisor
Dim strSupervisorID
Dim strWorkerName
Dim strWorkerID
Dim strManager
Dim strDirector
Dim strStaffingFields
Dim mstrSortOrder
Dim mdctColumns
Dim oColumn
Dim mstrShowColumns
Dim intI
Dim lngReviewClassID

strParmList = Request.QueryString("ParmList")
strLoad = Request.QueryString("Load")
mstrShowColumns = Request.QueryString("ShowColumns")
If IsNull(mstrShowColumns) Or mstrShowColumns = "" Then mstrShowColumns = "1^2^3^4^5^6^7^8^"

mstrUserID = Parse(strParmList,"^",1)
mlngAliasID = Parse(strParmList,"^",18)
blnUserAdmin = Parse(strParmList,"^",2)
blnUserQA = Parse(strParmList,"^",3)
lngCasID = Parse(strParmList,"^",4)
If lngCasID = "" Then lngCasID = NULL
dtmReviewDate = Parse(strParmList,"^",5)
If dtmReviewDate = "" Then dtmReviewDate = NULL
strCaseNumber = Parse(strParmList,"^",6)
strSubmitted = Parse(strParmList,"^",7)
lngResponse = Parse(strParmList,"^",8)
If Len(lngResponse) = 0 Or lngResponse = "0" Then lngResponse = Null
strReviewer = Parse(strParmList,"^",9)
If Len(strReviewer) = 0 Or strReviewer = "<All>" Then strReviewer = Null
strProgramsSelected = Parse(strParmList,"^",10)
dtmReviewDateEnd = Parse(strParmList,"^",11)
If dtmReviewDateEnd = "" Then dtmReviewDateEnd = NULL
strSupervisor = Parse(strParmList,"^",12)
strSupervisorID = Parse(strParmList,"^",13)
strWorkerName = Parse(strParmList,"^",14)
strWorkerID = Parse(strParmList,"^",15)
lngReviewClassID = Parse(strParmList,"^",19)
If Len(lngReviewClassID) = 0 Or lngReviewClassID = "0" Then lngReviewClassID = Null
mstrSortOrder = Parse(strParmList,"^",20)

Set mdctColumns = CreateObject("Scripting.Dictionary")

If strLoad = "Y" Then
    If mstrUserID <> "" And strProgramsSelected <> "" Then
        'Save the current selected programs:
        Set gadoCmd = GetAdoCmd("spProfileSettingUpd")
            AddParmIn gadoCmd, "@UserID", adVarChar, 20, ClearScript(mstrUserID)
            AddParmIn gadoCmd, "@SettingName", adVarChar, 50, "ProgramsSelected"
            AddParmIn gadoCmd, "@SettingValue", adVarChar, 255, ClearScript(strProgramsSelected)
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
        Set gadoCmd = Nothing
        'Save the current selected Find Columns:
        Set gadoCmd = GetAdoCmd("spProfileSettingUpd")
            AddParmIn gadoCmd, "@UserID", adVarChar, 20, ClearScript(mstrUserID)
            AddParmIn gadoCmd, "@SettingName", adVarChar, 50, "ShowColumns"
            AddParmIn gadoCmd, "@SettingValue", adVarChar, 255, ClearScript(mstrShowColumns)
            gadoCmd.Execute
        Set gadoCmd = Nothing
    End If

    Set adRsResults = Server.CreateObject("ADODB.Recordset") 

    Set adCmd = GetAdoCmd("spReviewFind")
        AddParmIn adCmd, "@AliasID", adInteger, 0, mlngAliasID
		AddParmIn adCmd, "@Admin", adBoolean, 0, blnUserAdmin
        AddParmIn adCmd, "@QA", adBoolean, 0, blnUserQA
		AddParmIn adCmd, "@UserID", adVarChar, 20, mstrUserID
        AddParmIn adCmd, "@casID", adInteger, 0, lngCasID
        AddParmIn adCmd, "@casNumber", adVarChar, 20, strCaseNumber
        AddParmIn adCmd, "@ReviewDate", adDBTimeStamp, 0, dtmReviewDate
        AddParmIn adCmd, "@ReviewDateEnd", adDBTimeStamp, 0, dtmReviewDateEnd
        AddParmIn adCmd, "@WorkerName", adVarChar, 100, IsBlank(strWorkerName)
        If strSubmitted <> "0" Then
            AddParmIn adCmd, "@Submitted", adVarchar, 1, strSubmitted
        Else
            AddParmIn adCmd, "@Submitted", adVarchar, 1, NULL
        End If
        AddParmIn adCmd, "@Response", adInteger, 0, lngResponse
        AddParmIn adCmd, "@Reviewer", adVarChar, 100, IsBlank(strReviewer)
        AddParmIn adCmd, "@PrgID", adVarchar, 255, strProgramsSelected
        AddParmIn adCmd, "@WorkerID", adVarchar, 20, IsBlank(strWorkerID)
        AddParmIn adCmd, "@Supervisor", adVarchar, 100, IsBlank(strSupervisor)
        AddParmIn adCmd, "@SupervisorID", adVarchar, 20, IsBlank(strSupervisorID)
        AddParmIn adCmd, "@ReviewClassID", adInteger, 0, lngReviewClassID
        
        'Call ShowCmdParms(adCmd) '***DEBUG

        'Open a recordset from the query:
        Call adRsResults.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
    Set adCmd = Nothing
    
    For intI = 1 To adRsResults.Fields.Count
        mdctColumns.Add "C" & intI, adRsResults.Fields(intI-1).Name
    Next
    If adRsResults.RecordCount > 0 And mstrSortOrder <> "" Then
        adRsResults.Sort = mstrSortOrder
    End If
End If
%>
<HTML>
<HEAD>
    <meta name=vs_targetSchema content="HTML 4.0">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>


<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload()
	Dim strProgramList
	
    window.parent.FindPageBody.style.cursor = "default"
    
	If Form.LoadTable.value = "Y" Then
        If Form.ResultsCount.Value > 0 Then
            window.parent.lblStatus.innerText = "Number of reviews matching the search criteria:   " & Form.ResultsCount.Value
            If lstCases.rows.length > 0 Then
                lstCases.rows(1).cells(0).tabindex = 9
            End If
        ElseIf Form.ResultsCount.Value = 0 Then
            window.parent.lblStatus.innerText = "No reviews matched the search criteria."
        End If
    End If

    If IsNumeric(Form.SelectedIndex.Value) Then
        If CLng(Form.SelectedIndex.Value) > 0 Then
            Call Result_onclick(1)
            window.parent.cmdEdit.disabled = False
            window.parent.cmdPrint.disabled = False
            window.parent.cmdPrintList.disabled = False
        End If
    Else
        window.parent.cmdEdit.disabled = True
        window.parent.cmdPrint.disabled = True
        window.parent.cmdEditWR.disabled = True
        window.parent.cmdPrintList.disabled = True
    End If
    
    ChildPageBody.style.cursor = "default"
End Sub

Sub EditRecord()
    Dim strInfo
    
    If IsNull(Form.SelectedIndex.Value) Or Trim(Form.SelectedIndex.Value) = "" Then
        Exit Sub
    End If
    strInfo = document.all("txtRowInfo" & Form.SelectedIndex.Value).value
    Call window.parent.EditRecord(Parse(strInfo,"^",1))
End Sub

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
                lstCases.rows(0).scrollIntoView
                Call Result_onclick(1)
            Case 35 'end
                Window.event.returnValue = False
                Call Result_onclick(lstCases.Rows.Length - 1)
            Case 38 'Up
                If Form.SelectedIndex.Value > 1 Then
                    Window.event.returnValue = False
                    Call Result_onclick(Form.SelectedIndex.Value - 1)
                End If
            Case 40 'Down
                If Cint(Form.SelectedIndex.Value) < Cint(lstCases.Rows.Length - 1) Then
                    Window.event.returnValue = False
                    Call Result_onclick(Form.SelectedIndex.Value + 1)
                End If
        End Select
    End If
End Sub

Sub Result_onclick(intRow)
    Dim strRow, strInfo
    If lstCases.disabled = True Then Exit Sub
    If CInt(Form.ResultsCount.value) = 0 Then
        Exit Sub
    End If
    If IsNumeric(Form.SelectedIndex.Value) Then
        strRow = "ListRow" & Form.SelectedIndex.Value
        lstCases.Rows(strRow).className = "TableRow"
        lstCases.Rows(strRow).cells(0).tabindex = -1
    End If

    strRow = "ListRow" & intRow
    lstCases.Rows(strRow).className = "TableSelectedRow"
    lstCases.Rows(strRow).cells(0).focus
    lstCases.Rows(strRow).cells(0).tabindex = 9
    
    strInfo = document.all("txtRowInfo" & intRow).value

    If Parse(strInfo,"^",2) = "Y" Then
        window.parent.cmdEditWR.disabled = False
    Else
        window.parent.cmdEditWR.disabled = True
    End If

    Form.SelectedIndex.Value = intRow
    window.parent.form.SelectedIndex.Value = intRow
    window.parent.form.rvwID.value = Parse(strInfo,"^",1)
End Sub

Sub Result_ondblclick(intRow)
    If lstCases.disabled = True Then Exit Sub
    Call Result_onclick(intRow)
    Call EditRecord
End Sub

Sub SortResults(strColumn)
    Dim strSortOrder
    
    strColumn = Replace(strColumn,"_"," ")
    If InStr(Form.SortOrder.Value, strColumn) > 0 And InStr(Form.SortOrder.Value, "DESC") = 0 Then
        ' Reload page, sorting on selected column ASC
        strSortOrder = strColumn & " DESC"
    Else
        ' Reload page, sorting on selected column DESC
        strSortOrder = strColumn
    End If
    window.parent.Form.SortOrder.value = strSortOrder
    Call window.parent.cmdFind_onclick()
End Sub

Sub PrintList()
    Dim strURL
    Dim strSortColumn, strReturnValue, strDetails
    strDetails = "&A1=<%=mlngAliasID%>"
    strDetails = strDetails & "&A2=<%=blnUserAdmin%>"
    strDetails = strDetails & "&A3=<%=blnUserQA%>"
    strDetails = strDetails & "&A4=<%=mstrUserID%>"
    strDetails = strDetails & "&A5=<%=lngCasID%>"
    strDetails = strDetails & "&A6=<%=strCaseNumber%>"
    strDetails = strDetails & "&A7=<%=dtmReviewDate%>"
    strDetails = strDetails & "&A8=<%=dtmReviewDateEnd%>"
    strDetails = strDetails & "&A9=<%=strWorkerName%>"
    'If "<%=strSubmitted%>" <> "0" Then
        strDetails = strDetails & "&A20=<%=strSubmitted%>"
    'Else
    '    strDetails = strDetails & "&A20="
    'End If
    strDetails = strDetails & "&A10=<%=lngResponse%>"
    strDetails = strDetails & "&A11=<%=strReviewer%>"
    strDetails = strDetails & "&A12=<%=strProgramsSelected%>"
    strDetails = strDetails & "&A13=<%=strWorkerID%>"
    strDetails = strDetails & "&A14=<%=strSupervisor%>"
    strDetails = strDetails & "&A15=<%=strSupervisorID%>"
    strDetails = strDetails & "&A16=<%=strManager%>"
    strDetails = strDetails & "&A17=<%=strDirector%>"
    strDetails = strDetails & "&A18=<%=strStaffingFields%>"
    strDetails = strDetails & "&A19=<%=lngReviewClassID%>"
    strDetails = strDetails & "&ShowCols=<%=mstrShowColumns%>"
    
    strSortColumn = "A1"
    Do While strSortColumn <> ""
        strURL = "RptReportDetails.asp?DC=<%=mstrShowColumns%>&SC=" & strSortColumn & "&Rpt=Find Case Review Listing&Proc=spReviewFind" & strDetails
        strReturnValue = window.showModalDialog(strURL, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
        strSortColumn = strReturnValue
    Loop
End Sub

</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=ChildPageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style="cursor:wait">
    <DIV id=lstResults class=TableDivArea title="Results of Find"
        style="LEFT:0; WIDTH:735; TOP:0; HEIGHT:295"
        tabIndex=-1>
        <%
        Dim intResultCnt
        Dim mintTblCols
        Dim strRespWrite
        Dim strRecord
        
        intResultCnt = 1
        strRespWrite = ""
        
        For intI = 1 To 100
            If Parse(mstrShowColumns,"^",intI) = "" Then Exit For
            strRecord = mdctColumns("C" & Parse(mstrShowColumns,"^",intI))
            strRespWrite = strRespWrite & "<TD class=CellLabel ID=thc" & Parse(mstrShowColumns,"^",intI) & " onclick=SortResults(""[" & Replace(strRecord," ","_") & "]"")" 
            strRespWrite = strRespWrite & "    style=""cursor:hand"" title=""Sort by " & strRecord & """>" & strRecord & "</TD>" & vbCrLf
            mintTblCols = mintTblCols + 1
        Next
        
		Response.Write "<Table ID=lstCases Border=0 Rules=rows Cols=" & mintTblCols & " Width=" & mintTblCols*100 & " CellSpacing=0 Style=""overflow: hidden; TOP:0;left:0""> " & vbCrLf
        Response.Write "<THEAD ID=ListHeader><TR ID=HeaderRow>"
        Response.Write strRespWrite
        Response.Write "</TR><THEAD>" & vbCrLf
        Response.Write "<TBODY ID=ListBody> " & vbCrLf
        strRespWrite = ""
        If strLoad = "Y" Then
            intLine = 1
            If Not adRsResults.BOF And Not adRsResults.EOF Then
                Do While Not adRsResults.EOF
                    Response.Write "<TR ID=ListRow" & intLine & " class=TableRow onclick=Result_onclick(" & intLine & ") ondblclick=Result_ondblclick(" & intLine & ")> " & vbCrLf

                    For intI = 1 To 100
                        If Parse(mstrShowColumns,"^",intI) = "" Then Exit For
                        strRecord = mdctColumns("C" & Parse(mstrShowColumns,"^",intI))
                        Response.Write "<TD ID=tbr" & Parse(mstrShowColumns,"^",intI) & intLine & " class=TableDetail>"
                        Response.Write adRsResults.Fields(strRecord).Value & "</TD>" & vbCrLf
                    Next
                    strRespWrite = strRespWrite & "<INPUT id=txtRowInfo" & intLine & " TYPE=""hidden"" VALUE=" & _
                        adRsResults.Fields("Review ID").Value & "^" & adRsResults.Fields("Sup Signed").Value & ">"

                    If intResultCnt > 50 Then
                        Response.Flush
                        intResultcnt = 1 
                    End If
                    intLine = intLine + 1
                    intResultCnt = intResultCnt + 1
                    adRsResults.MoveNext 
                Loop 
            End If
        End If
        Response.Write "</TBODY> </TABLE>"
        Response.Write strRespWrite
        %>
    </DIV>

    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY:hidden"" ACTION=""CaseEdit.ASP"" ID=Form>" & vbCrLf

    Call CommonFormFields()

    WriteFormField "ResultsCount", intLine - 1
    WriteFormField "LoadTable", strLoad
    WriteFormField "SortOrder", mstrSortOrder
    If intLine > 1 Then
        WriteFormField "SelectedIndex", 1
    Else
        WriteFormField "SelectedIndex", ""
    End If
    
    Response.Write Space(4) & "</FORM>"

    gadoCon.Close
    Set gadoCon = Nothing
    %>
</BODY>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncSvrFunctions.asp"-->