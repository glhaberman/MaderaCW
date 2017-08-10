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
Dim strResponse
Dim strLoad
Dim strProgramsSelected
Dim intLine
Dim strParmList
Dim mstrUserID, strReviewer
Dim strSupervisor, strManager, strWorker, strDirector, strOffice
Dim mstrReviewClass
Dim intI
Dim mstrSortOrder

strParmList = Request.QueryString("ParmList")
strLoad = Request.QueryString("Load")

lngCasID = Parse(strParmList,"^",1)
If lngCasID = "" Then lngCasID = NULL

dtmReviewDate = Parse(strParmList,"^",2)
If dtmReviewDate = "" Then dtmReviewDate = NULL

dtmReviewDateEnd = Parse(strParmList,"^",3)
If dtmReviewDateEnd = "" Then dtmReviewDateEnd = NULL

strCaseNumber = Parse(strParmList,"^",4)
If Len(strCaseNumber) = 0 Then strCaseNumber = Null

strWorker = Parse(strParmList,"^",5)
If Len(strWorker) = 0 Or strWorker = "<All>" Then strWorker = Null
strSupervisor = Parse(strParmList,"^",6)
If Len(strSupervisor) = 0 Or strSupervisor = "<All>" Or strSupervisor = "0"  Then strSupervisor = Null
strManager = Parse(strParmList,"^",7)
If Len(strManager) = 0 Or strManager = "<All>" Or strManager = "0"  Then strManager = Null
strOffice = Parse(strParmList,"^",8)
If Len(strOffice) = 0 Or strOffice = "<All>" Or strOffice = "0"  Then strOffice = Null
strDirector = Parse(strParmList,"^",9)
If Len(strDirector) = 0 Or strDirector = "<All>" Or strDirector = "0"  Then strDirector = Null

strResponse = Parse(strParmList,"^",10)
If Len(strResponse) = 0 Or strResponse = "<All>" Or strResponse = "0" Then strResponse = Null

strReviewer = Parse(strParmList,"^",11)
If Len(strReviewer) = 0 Or strReviewer = "<All>" Then strReviewer = Null

strProgramsSelected = Parse(strParmList,"^",12)
mstrReviewClass = Parse(strParmList, "^", 13)
If Len(mstrReviewClass) = 0 Or mstrReviewClass = "<All>" Then mstrReviewClass = Null

mstrSortOrder = Parse(strParmList, "^", 14)

If strLoad = "Y" Then
    Set adRsResults = Server.CreateObject("ADODB.Recordset") 
    Set adCmd = GetAdoCmd("spArcReviewFind")
        AddParmIn adCmd, "@casID", adInteger, 0, lngCasID
        AddParmIn adCmd, "@casNumber", adVarChar, 20, strCaseNumber
        AddParmIn adCmd, "@ReviewDate", adDBTimeStamp, 0, dtmReviewDate
        AddParmIn adCmd, "@ReviewDateEnd", adDBTimeStamp, 0, dtmReviewDateEnd
        AddParmIn adCmd, "@Response", adVarchar, 50, strResponse
        AddParmIn adCmd, "@Reviewer", adVarchar, 100, strReviewer
        AddParmIn adCmd, "@PrgID", adVarchar, 255, strProgramsSelected
        AddParmIn adCmd, "@Worker", adVarchar, 100, strWorker
        AddParmIn adCmd, "@Supervisor", adVarchar, 100, strSupervisor
        AddParmIn adCmd, "@Manager", adVarchar, 100, strManager
        AddParmIn adCmd, "@Office", adVarchar, 100, strOffice
        AddParmIn adCmd, "@Director", adVarchar, 100, strDirector
        AddParmIn adCmd, "@ReviewClass", adVarchar, 100, mstrReviewClass
        'Call ShowCmdParms(adCmd) '***DEBUG

        'Open a recordset from the query:
        Call adRsResults.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
    Set adCmd = Nothing
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
	
    window.parent.PageBody.style.cursor = "default"
    
	If Form.LoadTable.value = "Y" Then
        If Form.ResultsCount.Value > 0 Then
            If Form.ResultsCount.Value = "250" Then
                window.parent.lblStatus.innerHTML = "<b>The search returned too many results.&nbsp Only the first 250 will be displayed.</b>"
            Else
                window.parent.lblStatus.innerText = "Number of reviews matching the search criteria:   " & Form.ResultsCount.Value
            End If
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
            window.parent.cmdPrintReview.disabled = False
        End If
    Else
        window.parent.cmdPrintReview.disabled = True
    End If
    
    Call window.parent.DisablePage(False)
    PageBody.style.cursor = "default"
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
    Dim strRow
    If CInt(Form.ResultsCount.value) = 0 Or lstCases.disabled Then
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

    Form.SelectedIndex.Value = intRow
    window.parent.form.SelectedIndex.Value = intRow
    window.parent.form.rvwID.value = Document.all("tdcLine" & intRow & "Col0").innerText
End Sub

Sub Result_ondblclick(intRow)
    Call Result_onclick(intRow)
    Call window.parent.cmdPrintReview_onclick()
End Sub

Sub SortResults(strColumn)
    Dim strSortOrder

    If lstCases.disabled Then
        Exit Sub
    End If

    strColumn = Replace(strColumn,"_"," ")
    If InStr(Form.SortOrder.Value, strColumn) > 0 And InStr(Form.SortOrder.Value, "DESC") = 0 Then
        ' Reload page, sorting on selected column ASC
        strSortOrder = strColumn & " DESC"
    Else
        ' Reload page, sorting on selected column DESC
        strSortOrder = strColumn
    End If
    window.parent.Form.SortOrder.value = strSortOrder
    window.parent.Form.UseWarning.value = "No"
    Call window.parent.cmdFind_onclick()
End Sub

</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style="cursor:wait">
    <DIV id=lstResults class=TableDivArea
        style="LEFT:0; WIDTH:735; TOP:0; HEIGHT:258"
        tabIndex=-1>
        <%
        Dim intResultCnt
        Dim mintTblWidth
        Dim strFieldName
        
        intResultCnt = 1
        mintTblWidth = 900

        If strLoad = "Y" Then
		    Response.Write "<Table ID=lstCases Border=0 Rules=rows Cols=" & adRsResults.Fields.Count & " Width=" & mintTblWidth & " CellSpacing=0 Style=""overflow: hidden; TOP:0;left:0""> " & vbCrLf
            Response.Write "<THEAD ID=ListHeader><TR ID=HeaderRow>"
            For intI = 0 To adRsResults.Fields.Count - 1 
                strFieldName = adRsResults.Fields(intI).Name
                Response.Write "<TD class=CellLabel style=""cursor:hand"" onclick=SortResults(""[" & Replace(strFieldName," ","_") & "]"") title=""Sort By " & strFieldName & """>" & strFieldName & "</TD>" & vbCrLf
            Next
            Response.Write "</TR><THEAD>" & vbCrLf
            Response.Write "<TBODY ID=ListBody> " & vbCrLf

            intLine = 1
            If Not adRsResults.BOF And Not adRsResults.EOF Then
                Do While Not adRsResults.EOF
                    Response.Write "<TR ID=ListRow" & intLine & " class=TableRow onclick=Result_onclick(" & intLine & ") ondblclick=Result_ondblclick(" & intLine & ")> " & vbCrLf
                    For intI = 0 To adRsResults.Fields.Count - 1
                        Response.Write "<TD ID=tdcLine" & intLine & "Col" & intI & " class=TableDetail>" & vbCrLf
                        Response.Write adRsResults.Fields(intI).Value & "</TD>" & vbCrLf
                    Next
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
        %>
    </DIV>

    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY:hidden"" ACTION=""CaseEdit.ASP"" ID=Form>" & vbCrLf

    Call CommonFormFields()

    WriteFormField "ResultsCount", intLine - 1
    WriteFormField "LoadTable", strLoad
    If intLine > 1 Then
        WriteFormField "SelectedIndex", 1
    Else
        WriteFormField "SelectedIndex", ""
    End If

    WriteFormField "SortOrder", mstrSortOrder

    Response.Write Space(4) & "</FORM>"

    gadoCon.Close
    Set gadoCon = Nothing
    %>
</BODY>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncSvrFunctions.asp"-->