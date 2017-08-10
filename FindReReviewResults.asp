<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>

<!--#include file="IncCnn.asp"-->

<%
Dim adRsResults
Dim adCmd
Dim lngCasID, lngReReviewID
Dim dtmReReviewDate
Dim dtmReReviewDateEnd
Dim strCaseNumber
Dim strReReviewer
Dim strLoad
Dim intLine
Dim strUserID, mlngAliasID
Dim blnUserAdmin
Dim blnUserQA
Dim strParmList
Dim mstrUserID
Dim mstrSortOrder
Dim strReviewer
Dim mintReReviewTypeID, mstrReReviewType

strParmList = Request.QueryString("ParmList")
strLoad = Request.QueryString("Load")

mstrUserID = Parse(strParmList,"^",1)
blnUserAdmin = Parse(strParmList,"^",2)
blnUserQA = Parse(strParmList,"^",3)
mlngAliasID = Parse(strParmList,"^",12)
lngReReviewID = Parse(strParmList,"^",4)
If lngReReviewID = "" Then lngReReviewID = NULL
lngCasID = Parse(strParmList,"^",5)
If lngCasID = "" Then lngCasID = NULL
dtmReReviewDate = Parse(strParmList,"^",6)
If dtmReReviewDate = "" Then dtmReReviewDate = NULL
dtmReReviewDateEnd = Parse(strParmList,"^",7)
If dtmReReviewDateEnd = "" Then dtmReReviewDateEnd = NULL
strCaseNumber = Parse(strParmList,"^",8)
strReReviewer = Parse(strParmList,"^",9)
mstrSortOrder = Parse(strParmList,"^",10)
strReviewer = Parse(strParmList,"^",11)
mintReReviewTypeID = Request.QueryString("ReReviewTypeID")

If CInt(mintReReviewTypeID) = 0 Then
    mstrReReviewType = gstrEvaluation
Else
    mstrReReviewType = "CAR"
End If

If strLoad = "Y" Then
    Set adRsResults = Server.CreateObject("ADODB.Recordset") 
    Set adCmd = GetAdoCmd("spReReviewFind")
		AddParmIn adCmd, "@AliasID", adInteger, 0, mlngAliasID
		AddParmIn adCmd, "@Admin", adBoolean, 0, blnUserAdmin
		AddParmIn adCmd, "@QA", adBoolean, 0, blnUserQA
		AddParmIn adCmd, "@UserID", adVarchar, 20, mstrUserID
		AddParmIn adCmd, "@rrvID", adInteger, 0, lngReReviewID
		AddParmIn adCmd, "@rrvOrgReviewID", adInteger, 0, IsBlank(lngCasID)
        AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(dtmReReviewDate)
        AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(dtmReReviewDateEnd)
		AddParmIn adCmd, "@rrvReReviewer", adVarchar, 100, IsBlank(strReReviewer)
		AddParmIn adCmd, "@CaseNumber", adVarchar, 20, IsBlank(strCaseNumber)
		AddParmIn adCmd, "@Reviewer", adVarchar, 100, IsBlank(strReviewer)
		AddParmIn adCmd, "@ReReviewTypeID", adInteger, 0, mintReReviewTypeID
		
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
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
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
            window.parent.lblStatus.innerText = "Number of <%=mstrReReviewType%>s matching the search criteria:   " & Form.ResultsCount.Value
            If lstCases.rows.length > 0 Then
                lstCases.rows(1).cells(0).tabindex = 9
            End If
        ElseIf Form.ResultsCount.Value = 0 Then
            window.parent.lblStatus.innerText = "No <%=mstrReReviewType%>s matched the search criteria."
        End If
    End If

    If IsNumeric(Form.SelectedIndex.Value) Then
        If CLng(Form.SelectedIndex.Value) > 0 Then
            Call Result_onclick(1)
            window.parent.cmdEdit.disabled = False
            window.parent.cmdPrint.disabled = False
        End If
    Else
        window.parent.cmdEdit.disabled = True
        window.parent.cmdPrint.disabled = True
    End If
    
    PageBody.style.cursor = "default"
End Sub

Sub EditRecord()
    If IsNull(Form.SelectedIndex.Value) Or Trim(Form.SelectedIndex.Value) = "" Then
        Exit Sub
    End If
    window.parent.Form.ReReviewID.Value = Document.all("ReReviewID" & Form.SelectedIndex.Value).innerText
    Call window.parent.EditRecord()
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

    Form.SelectedIndex.Value = intRow
    window.parent.form.SelectedIndex.Value = intRow
    window.parent.form.ReReviewID.value = Document.all("ReReviewID" & intRow).innerText
End Sub

Sub Result_ondblclick(intRow)
    Call Result_onclick(intRow)
    Call EditRecord
End Sub

Sub SortResults(strColumn)
    Dim strSortOrder
    If CInt(Form.ResultsCount.value) <= 0 Then
        Exit Sub
    End If
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
</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style="cursor:wait">
    <DIV id=lstResults class=TableDivArea title="Results of Find"
        style="LEFT:0; WIDTH:735; TOP:0; HEIGHT:303"
        tabIndex=-1>
        <%
        Dim intResultCnt
        Dim mintTblWidth
        Dim strRowStart
        Dim strValue
        
        intResultCnt = 1
        mintTblWidth = 715
		Response.Write "<Table ID=lstCases Border=0 Rules=rows Cols=7 Width=" & mintTblWidth & " CellSpacing=0 Style=""overflow: hidden; TOP:0;left:0""> " & vbCrLf
        Response.Write "<THEAD ID=ListHeader><TR ID=HeaderRow>"
        Response.Write "<TD class=CellLabel style=""cursor:hand;width:60"" onclick=SortResults(""rrvID"") ID=ReReviewID title=""Sort By " & mstrReReviewType & " ID"">" & mstrReReviewType & " ID</TD>" & vbCrLf
        Response.Write "<TD class=CellLabel style=""cursor:hand"" onclick=SortResults(""rvwID"") ID=ReviewID title=""Sort By Review ID"">Review ID</TD>" & vbCrLf
        Response.Write "<TD class=CellLabel style=""cursor:hand"" onclick=SortResults(""rvwCaseNumber"") ID=casNumber title=""Sort By Case Number"">Case#</TD>" & vbCrLf
        Response.Write "<TD class=CellLabel style=""cursor:hand"" onclick=SortResults(""rrvDateEntered"") ID=casEntered title=""Sort By Re-Review Date"">Re-Review Date</TD>" & vbCrLf
        Response.Write "<TD class=CellLabel style=""cursor:hand"" onclick=SortResults(""ReReviewerName"") ID=casReReviewer title=""Sort By " & gstrEvaTitle & """>" & gstrEvaTitle & "</TD>" & vbCrLf
        Response.Write "<TD class=CellLabel style=""cursor:hand"" onclick=SortResults(""rvwReviewerName"") ID=casReviewer title=""Sort By Original Case " & gstrRvwTitle & """>" & gstrRvwTitle & "</TD>" & vbCrLf
        Response.Write "<TD class=CellLabel style=""cursor:hand"" onclick=SortResults(""ClientName"") ID=casClient title=""Sort By Client Name"">Case Name</TD>" & vbCrLf
        Response.Write "<TD class=CellLabel style=""cursor:hand"" onclick=SortResults(""rrvSubmitted"") ID=casSubmitted title=""Sort By Submitted"">Sub?</TD>" & vbCrLf
        
        Response.Write "</TR><THEAD>" & vbCrLf
        Response.Write "<TBODY ID=ListBody> " & vbCrLf
        If strLoad = "Y" Then
            intLine = 1
            If Not adRsResults.BOF And Not adRsResults.EOF Then
                Do While Not adRsResults.EOF
                    strRowStart = "<TR ID=ListRow" & intLine & " class=TableRow onclick=Result_onclick(" & intLine & ") ondblclick=Result_ondblclick(" & intLine & ")> " & vbCrLf
                    Response.Write strRowStart & "<TD ID=ReReviewID" & intLine & " title=""" & mstrReReviewType & " ID"" class=TableDetail>" & vbCrLf
                    Response.Write adRsResults.Fields("rrvID").Value & "</TD>" & vbCrLf
                    Response.Write "<TD ID=ReviewID" & intLine & " title=""Review ID"" class=TableDetail>"
                    Response.Write adRsResults.Fields("rvwID").Value & "</TD>" & vbCrLf
                    Response.Write "<TD ID=casNumber" & intLine & " title=""Case Number"" class=TableDetail>"
                    Response.Write adRsResults.Fields("rvwCaseNumber").Value & "</TD>" & vbCrLf
                    Response.Write "<TD ID=casEntered" & intLine & " title=""Case Review Date Entered"" class=TableDetail>"
                    Response.Write FormatDateTime(adRsResults.Fields("rrvDateEntered").Value) & "</TD>" & vbCrLf
                    Response.Write "<TD ID=casReReviewer" & intLine & " title=""" & gstrEvaTitle & """ class=TableDetail>"
                    Response.Write Parse(adRsResults.Fields("ReReviewerName").Value,"--",1) & "</TD>" & vbCrLf
                    Response.Write "<TD ID=casReviewer" & intLine & " title=""" & gstrRvwTitle & """ class=TableDetail>"
                    Response.Write Parse(adRsResults.Fields("rvwReviewerName").Value,"--",1) & "</TD>" & vbCrLf
                    Response.Write "<TD ID=casClient" & intLine & " title=""Client Name"" class=TableDetail>"
                    Response.Write adRsResults.Fields("ClientName").Value & "</TD>" & vbCrLf
                    Response.Write "<TD ID=casSubmitted" & intLine & " title=""Submitted"" class=TableDetail>"
                    If adRsResults.Fields("rrvSubmitted").Value = "Y" Then
                        strValue = "Yes"
                    Else
                        strValue = "No"
                    End If
                    Response.Write strValue & "</TD>" & vbCrLf
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