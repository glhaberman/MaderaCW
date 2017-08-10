<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<%
    Dim intI
    Dim intJ
    Dim intWidth
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<HTML>
<HEAD>
    <META name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE>
        <%=Trim(gstrOrgAbbr & " " & gstrAppName)%>
    </TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>
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

    PageFrame.style.visibility = "visible"
    txtSQL.value = Form.SQL.value
    txtSQL.focus
End Sub

<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    mblnSetFocusToMain = False
    mblnCloseClicked = True
    window.close
End Sub

Sub window_onbeforeunload
    If Not mblnCloseClicked Then
        If Form.FormAction.value <> "" Then
            window.event.returnValue = "Closing the browser window will exit the application without saving." & space(10) & vbCrLf & "Please use the <Save> button to save your changes, then use" & space(10) & vbcrlf & "the <Close> button to return to the main menu." & space(10)
        Else
            window.event.returnValue = "Closing the browser window will exit the application." & space(10) & vbcrlf & "Please use the <Close> button to return to the main menu." & space(10)
        End If
    End If
    If mblnSetFocusToMain = True Then
        window.opener.focus
    End If
End Sub

Sub cmdExecute_onclick
    If Trim(txtSQL.value) = "" Then
        Exit Sub
    End If
    cmdExecute.disabled = True
    Form.SQL.value = txtSQL.value
    Form.FormAction.value = "Execute"
    Form.Action = "SQLQueries.asp"
    mblnCloseClicked = True
    mblnSetFocusToMain = False
    Form.Submit
End Sub

Sub cmdCopy_onclick
    Dim CtlRng
    'If the results div is not empty, copy it's contents to the clipboard:
    If divResults.children.length > 0 Then
        'A controlRange object is used to select the results div, then copy it:
        Set CtlRng = PageBody.createControlRange()
        CtlRng.AddElement(divResults)
        CtlRng.Select
        CtlRng.execCommand("Copy")
        Set CtlRng = Nothing
        'Clear the selection:
        document.selection.empty
        MsgBox "Results copied to clipboard.", ,"Copy Results"
    End If
End Sub

Sub cmdClose_onclick
    mblnCloseClicked = True
    mblnSetFocusToMain = False
    Form.Action = "Admin.asp"
    Form.Submit
End Sub
</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->
<BODY id="PageBody" bottomMargin="5" topMargin="5" leftMargin="5" rightMargin="5">
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
                BORDER-WIDTH: 2">Execute SQL Queries
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>
    <% Call WriteNavigateControls(-1,0,gstrAltBackColor) %>

    <DIV id=PageFrame
        style="OVERFLOW: hidden;
            border-style: solid;
            border-width: 2;
            border-color: <%=gstrBorderColor%>;
            TOP: 32; 
            HEIGHT: 415; 
            WIDTH: 745; 
            COLOR: black; 
            BACKGROUND-COLOR: <%=gstrBackColor%>">
            <SPAN id="lblElement" class="DefLabel" style="LEFT:10; WIDTH:185; TOP:5">
            Enter T-SQL:
        </SPAN>
            <TEXTAREA id="txtSQL" style="LEFT:10; WIDTH:720; HEIGHT:100; TOP:25; TEXT-ALIGN:left; padding-left:5; FONT-FAMILY:Courier; OVERFLOW:scroll"
                wrap="off" tabIndex="1" NAME="txtSQL"></TEXTAREA>
            <SPAN id="Span1" class="DefLabel" style="LEFT:10; WIDTH:185; TOP:135">
            Results:
        </SPAN>
        <DIV id=divResults class=TableDivArea style="LEFT:10; WIDTH:720; TOP:155; HEIGHT:200;" tabIndex=5>
            <% Call GetResults()%>
        </DIV>
        <DIV id=fraButtons
            style="LEFT: -2; 
                border-style: solid;
                border-width: 2;
                border-color: <%=gstrBorderColor%>;
                TOP: 375; 
                HEIGHT: 40px; 
                WIDTH: 745; 
                BACKGROUND-COLOR: <%=gstrAltBackColor%>">
                <SPAN id="lblDatabaseStatus" class="DefLabel" style="VISIBILITY:hidden; LEFT:5; WIDTH:200; TOP:10; TEXT-ALIGN:center">
                Accessing Database...
            </SPAN>
            <BUTTON id="cmdExecute" class="DefButton" title="Execute the query" style="LEFT:15; WIDTH:65;  TOP:7; HEIGHT:20"
                accesskey="E" tabIndex="284">
                <U>E</U>xecute
            </BUTTON>
            <BUTTON id="cmdCopy" class="DefButton" title="Copy results to clipboard" style="LEFT:85; WIDTH:65;  TOP:7; HEIGHT:20"
                accesskey="C" tabIndex="284">
                <U>C</U>opy
            </BUTTON>
            <BUTTON id="cmdClose" class="DefButton" title="Close and return to main screen" style="LEFT:655; WIDTH:65; TOP:7; HEIGHT:20"
                tabIndex="290">
                Close
            </BUTTON>
        </DIV>
    </DIV>
    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY: hidden"" ACTION=""Main.asp"" ID=Form>" & vbCrLf
        Call CommonFormFields()
    	WriteFormField "SQL", ReqForm("SQL")
    	WriteFormField "FormAction", ReqForm("FormAction")
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
Sub GetResults()
    Dim objLastErr
    Dim intCol
    Dim intWidth
    Dim adoRs
    Dim intT
    Dim strValue
    
    If Request.Form("FormAction") <> "Execute" Then
        Exit Sub
    End If

    Set adoRs = Server.CreateObject("ADODB.Recordset")
    adoRs.CursorLocation = adUseClient
    adoRs.CursorType = adOpenForwardOnly
    adoRs.LockType = adLockReadOnly
    Set gadoCmd = Server.CreateObject("ADODB.Command")
    With gadoCmd
        .ActiveConnection = gadoCon
        .CommandTimeout = 180
        .CommandType = adCmdText
        .CommandText = Request.Form("SQL")
        On Error Resume Next
        Set adoRs = .Execute
        On Error Goto 0
    End With
    
    Set objLastErr = Server.GetLastError()
    
    intT = 0
    If Trim(objLastErr.ASPCode) = "" Then
        If adoRs Is Nothing Then
            If gadoCon.Errors.Count > 0 Then
                WriteError gadoCon.Errors, intT, "E"
            Else
                WriteError "** No Results **", intT, ""
            End If
        End If
        Do While Not adoRs Is Nothing
            If adoRs.State = adStateOpen Then
                Response.Write "<TABLE ID=tblResults" & intT & " Width=700 CellSpacing=0 Style=""overflow: scroll; TOP:0"">" & vbCrLf
                Response.Write "<TBODY id=tbdResults" & intT & ">" & vbCrLf
                Response.Write "<THEAD id=thdResults" & intT & ">" & vbCrLf
                Response.Write "<TR id=thrResults" & intT & ">" & vbCrLf
                For intCol = 0 To adoRs.Fields.Count - 1
                    intWidth = Len(adoRs.Fields(intCol).Name) + 2
                    Response.Write "<TD class=CellLabel ID=thd" & intT & "Col" & intCol & " style=""COLOR: beige;background-color:darkolivegreen;width:" & intWidth & """>" & adoRs.Fields(intCol).Name & "</TD>" & vbCrLf
                Next
                Response.Write "</TR>" & vbCrLf
                Response.Write "</THEAD>" & vbCrLf
                Response.Write vbcrlf
                intI = 0
                Do While Not adoRs.EOF
                    Response.Write "<TR id=tbrT" & intT & "R" & intI & ">" & vbCrLf
                    For intCol = 0 To adoRs.Fields.Count - 1
                        intWidth = Len(adoRs.Fields(intCol).Name) + 2
                        If IsNull(adoRs.Fields(intCol).value) Then
                            strValue = "NULL"
                        Else
                            strValue = CStr(adoRs.Fields(intCol).value)
                        End If
                        Response.Write "<TD class=CellLabel ID=tbdR" & intI & "C" & intCol & " style=""text-align: left;"">" & strValue & "</TD>" & vbCrLf
                    Next
                    intI = intI + 1
                    Response.Write "</TR>" & vbCrLf
                    Response.Write vbcrlf
                    On Error Resume Next
                    adoRs.MoveNext
                    On Error Goto 0
                    If gadoCon.Errors.Count > 0 Then
                        WriteError gadoCon.Errors, intT, "E"
                        Exit Do
                    End If
                Loop
                Response.Write "</TBODY>" & vbCrLf
                Response.Write "</TABLE>" & vbCrLf
                Response.Write vbcrlf & vbcrlf
                Set adoRs = adoRs.NextRecordset
                intT = intT + 1
            Else
                If gadoCon.Errors.Count > 0 Then
                    WriteError gadoCon.Errors, intT, "E"
                Else
                    WriteError "** No Results **", intT, ""
                End If
                Set adoRs = Nothing
            End If
        Loop
    Else
        WriteError " ** Error: " & objLastErr.ASPCode & " - " & objLastErr.Description, intT, "ASP Error"
        WriteError gadoCon.Errors, intT, "ADO Errors"
    End If
End Sub

Sub WriteError(objErrors,intTableID,strHeading)
    Dim intCol
    
    If strHeading = "E" Then strHeading = "Error Message(s)"
    Response.Write "<TABLE ID=tblResults" & intTableID & " Width=700 CellSpacing=0 Style=""overflow: scroll; TOP:0"">" & vbCrLf
    Response.Write "<TBODY id=tbdResults" & intTableID & ">" & vbCrLf
    Response.Write "<THEAD id=thdResults" & intTableID & ">" & vbCrLf
    Response.Write "<TR id=thrResults" & intTableID & ">" & vbCrLf
    Response.Write "<TD class=CellLabel ID=thd" & intTableID & "Col0" & " style=""COLOR: beige;background-color:darkolivegreen;width:" & intWidth & """>" & strHeading & "</TD>" & vbCrLf
    Response.Write "</TR>" & vbCrLf
    Response.Write "</THEAD>" & vbCrLf

    If IsObject(objErrors) Then
        ' If objErrors is a collection of erros, add each error to table
        For intCol = 0 To objErrors.Count - 1
            Response.Write "<TR id=tbrResults" & intCol & ">" & vbCrLf
            Response.Write "<TD class=CellLabel ID=tbd" & intCol & "Col0" & " style=""text-align:left"">" & objErrors.Item(intCol).Description & "</TD>" & vbCrLf
            Response.Write "</TR>" & vbCrLf
        Next
    Else
        ' If objErrors is not a collection of erros, display objErrors as a string
        Response.Write "<TR id=tbrResults" & intCol & ">" & vbCrLf
        Response.Write "<TD class=CellLabel ID=tbd" & intCol & "Col0" & " style=""text-align:left"">" & objErrors & "</TD>" & vbCrLf
        Response.Write "</TR>" & vbCrLf
    End If

    Response.Write "</TBODY>" & vbCrLf
    Response.Write "</TABLE>" & vbCrLf
End Sub
%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncNavigateControls.asp"-->