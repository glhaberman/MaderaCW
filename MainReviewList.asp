<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>

<!--#include file="IncCnn.asp"-->

<%
Dim adRsResults
Dim adCmd
Dim intI, intJ, intWidth, strAlign
Dim mstrUserType
Dim strLoad, strType, strHidden
Dim mstrSortOrder, mstrUserID, strSetFocus

mstrUserType = Request.QueryString("UserType")
mstrUserID = Request.QueryString("UserID")
strLoad = Request.QueryString("Load")
mstrSortOrder = Request.QueryString("SortOrder")
strSetFocus = Request.QueryString("SetFocus")

'If strLoad = "N" Then
Set adRsResults = Server.CreateObject("ADODB.Recordset") 
Set adCmd = GetAdoCmd("spReviewFindMain")
	AddParmIn adCmd, "@UserID", adVarchar, 20, mstrUserID
    AddParmIn adCmd, "@UserLevel", adChar, 1, mstrUserType
    'Call ShowCmdParms(adCmd) '***DEBUG
    Call adRsResults.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
Set adCmd = Nothing
If adRsResults.RecordCount > 0 And mstrSortOrder <> "" Then
    adRsResults.Sort = mstrSortOrder
End If
%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="HTML 4.0">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>


<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload()
	If Form.LoadTable.value = "Y" Then
        Call window.parent.ShowReviewList(Form.ResultsCount.Value)
    End If

    If IsNumeric(Form.SelectedIndex.Value) Then
        If CLng(Form.SelectedIndex.Value) >= 0 Then
            Call Result_onclick(0)
        End If
    End If
    PageBody.style.cursor = "default"
    Select Case "<%=strSetFocus%>"
        Case "CASEADDEDIT"
            Set objWindow = window.parent.GetWindow(1)
            objWindow.focus
        Case "REREVIEWADDEDIT"
            Set objWindow = window.parent.GetWindow(4)
            objWindow.focus
        Case "CARREREVIEWADDEDIT"
            Set objWindow = window.parent.GetWindow(7)
            objWindow.focus
        Case Else
    End Select
End Sub

Sub EditRecord(intRowID)
    Dim intTypeID, intID
    
    intTypeID = Parse(Document.all("hidRowInfo" & intRowID).value,"^",1)
    intID = Parse(Document.all("hidRowInfo" & intRowID).value,"^",2)
    
    If intTypeID = "-1" Then
        window.parent.Form.rvwID.Value = intID
        window.parent.Form.Action = "CaseAddEdit.asp"
        
        window.parent.Form.FormAction.Value = "GetRecord"
        Call window.parent.ManageWindows(1,"EditReview")
    Else
        window.parent.Form.ReReviewID.Value = intID
        window.parent.Form.Action = "ReReviewAddEdit.asp"
        intWindowID = (intTypeID*3) + 4
        window.parent.Form.FormAction.Value = "GetRecord"
        window.parent.Form.ReReviewTypeID.value = intTypeID
        Call window.parent.ManageWindows(intWindowID,"EditReReview")
    End If
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
End Sub

Sub Result_ondblclick(intRow)
    Call Result_onclick(intRow)
    Call EditRecord(intRow) 'Document.all("tdcLine" & intRow & "Col0").innerText)
End Sub
</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style="cursor:wait">
    <DIV id=lstResults class=TableDivArea title="Results of Find"
        style="LEFT:0; WIDTH:667; TOP:0; HEIGHT:189"
        tabIndex=-1>
        <%
		Response.Write "<Table ID=lstCases Border=0 Rules=rows Width=650 CellSpacing=0 Style=""overflow: hidden; TOP:0;left:0""> " & vbCrLf
        Response.Write "<TBODY ID=ListBody> " & vbCrLf
        If strLoad = "Y" Then
            intI = -1
            If Not adRsResults.BOF And Not adRsResults.EOF Then
                strHidden = ""
                Do While Not adRsResults.EOF
                    intI = intI + 1
                    strHidden = strHidden & "<INPUT type=hidden id=hidRowInfo" & intI & " value=""" & adRsResults.Fields("rrvTypeID").Value & "^" & adRsResults.Fields(0).Value & """>"
                    Response.Write "<TR ID=ListRow" & intI & " class=TableRow onclick=Result_onclick(" & intI & ") ondblclick=Result_ondblclick(" & intI & ")> " & vbCrLf
                    For intJ = 0 To 5
                        strAlign = "center"
                        Select Case intJ
                            Case 0
                                intWidth = 90
                            Case 1,3
                                intWidth = 80
                            Case 4
                                intWidth = 105
                            Case 2
                                intWidth = 135
                                strAlign = "left"
                            Case 5
                                intWidth = 140
                        End Select
                        Response.Write "<TD ID=tdcLine" & intI & "Col" & intJ & " class=TableDetail style=""width:" & intWidth & ";text-align:" & strAlign & """>" & vbCrLf
                        'If intJ = 0 Then
                        '    Select Case adRsResults.Fields("rrvTypeID").Value
                        '        Case -1 'Review
                        '            Select Case adRsResults.Fields("TypeName").Value
                        '                Case "Review"
                        '                    strType = " (R)"
                        '                Case "ReviewRD"
                        '                    strType = " (R-RD)"
                        '                Case "ReviewNRD"
                        '                    strType = " (R-NRD)"
                        '            End Select
                        '        Case 0 'Re-Review
                        '            strType = " (RR)"
                        '        Case 1 'CAR
                        '            strType = " (CAR)"
                        '    End Select
                        '    Response.Write adRsResults.Fields(intJ).Value & strType & "</TD>" & vbCrLf
                        'Else
                            Response.Write adRsResults.Fields(intJ).Value & "</TD>" & vbCrLf
                        'End If
                    Next
                    adRsResults.MoveNext 
                Loop 
            End If
        End If
        Response.Write "</TBODY> </TABLE>" & vbCrLf
        Response.Write strHidden & vbCrLf
        %>
    </DIV>

    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY:hidden"" ACTION=""CaseEdit.ASP"" ID=Form>" & vbCrLf

    Call CommonFormFields()

    WriteFormField "ResultsCount", intI + 1
    WriteFormField "LoadTable", strLoad
    If intI > 0 Then
        WriteFormField "SelectedIndex", 0
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