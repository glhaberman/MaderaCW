<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
Dim adCmd
Dim adRs
Dim adRsElms
Dim strSQL
Dim intTop      'Used in building the spans for the question list.
Dim intQstCnt   'Counter used in building the spans for question list.
Dim strQstPrv   'Used to build spans for the question list.
Dim strOptions  'Used to build option lists for dropdown SELECTS.
Dim intElm
Dim intPrg
Dim strAction
Dim strSpaces

Dim mstrPageTitle

%>
<!--#include file="IncCnn.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName

Response.ExpiresAbsolute = Now - 5

If Trim(Request.QueryString("Program")) <> "" Then
    intPrg = Request.QueryString("Program")
    intElm = Request.QueryString("Element")
    strAction = "Find"
Else
    intPrg = ReqForm("Program")
    intElm = ReqForm("Element")
    strAction = ReqForm("FormAction")
End If
%>

<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE>Review Guide</TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <STYLE id=ThisPageStyles type="text/css" rel="stylesheet">
        
        .ReportText
            {
            PADDING-LEFT: 3;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: left;
            OVERFLOW: hidden
            }
        
        .HeaderCell
            {
            PADDING:5;
            BORDER:1 solid #C0C0C0;
            BORDER-BOTTOM:none>
            }

        .DetailCell
            {
            PADDING:5;
            BORDER:1 solid #C0C0C0;
            BORDER-TOP:none
            }
            
        .SpaceCell
            {
            WIDTH:50;
            BORDER:1 solid #C0C0C0;
            BORDER-RIGHT:none
            }
        
        .SelectedRow
            {
            COLOR:<%=gstrTitleColor%>;
            BACKGROUND-COLOR:<%=gstrPageColor%>
            }
        
    </STYLE>
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim mintSelectedQuestion    'Holds index of currently selected question.

Sub window_onload
    Dim intI
    Dim oOpt
    
    mblnDoneLoading = False
    mintSelectedQuestion = 0

    Set window.opener.oGuideWindow = window
    cboPrograms.style.visibility = "visible"
    cboElements.style.visibility = "visible"
    
    For Each oOpt In cboPrograms.options
        If Parse(oOpt.value, ":", 1) = Form.Program.Value Then
            cboPrograms.selectedIndex = oOpt.index
            Exit For
        End If
    Next

    Call cboPrograms_onchange()

    If Form.Element.Value <> "" Then
        cboElements.value = Form.Element.Value
    End If
    
    Call window_onresize
End Sub

Sub cboPrograms_onchange()
    Dim intI
    Dim oOption
    Dim intPrg
    Dim intABD

    If Trim(cboPrograms.value) = vbNullstring Then
		Exit Sub
    End If

    cboElements.options.length = Null

    Set oOption = document.createElement("OPTION")
    oOption.value = 0
    oOption.text = "<Any>"
    cboElements.options.add(oOption)

    intPrg = Parse(cboPrograms.Value, ":", 1)
    If intPrg = 14 Then
        intABD = 14
    Else
        intABD = 0
    End If
    For intI = 0 To cboElementList.options.length - 1
        If CStr(intPrg) = Parse(cboElementList.options.item(intI).Value, ":", 1) Then
            Set oOption = Document.createElement("OPTION")
            oOption.Value = Parse(cboElementList.options.item(intI).Value, ":", 2)
            oOption.Text = Parse(cboElementList.options.item(intI).Text, "^", 2)
            If oOption.Text = "ABD" Then
                If intABD = 14 Then
                    cboElements.options.add oOption
                End If
            Else
                cboElements.options.add oOption
            End If
            Set oOption = Nothing
        End If
    Next
    
    cboElements.value = 0
End Sub

Sub cmdClose_onclick
    cmdClose.disabled = True
    cmdPrint.disabled = True
    divHeader.style.visibility = "hidden"
    divFind.style.visibility = "hidden"
    window.opener.focus
    window.close
End Sub

Sub window_onbeforeunload
    window.opener.ClearGuideWindow
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub window_onbeforeprint()
    cmdPrint.style.visibility = "hidden"
    cmdClose.style.visibility = "hidden"
    cmdFind.style.visibility = "hidden"
    divFind.style.visibility = "hidden"
    window.cboElements.style.visibility = "hidden"
    window.cboPrograms.style.visibility = "hidden"
    divHeader.style.width = 630
    lblAppTitle.style.width = 630
    tblQuestions.style.width = 630
    divQuestions.style.top = "50"
End Sub

Sub window_onafterprint()
    cmdPrint.style.visibility = "visible"
    cmdClose.style.visibility = "visible"
    cmdFind.style.visibility = "visible"
    divFind.style.visibility = "visible"
    divQuestions.style.top = "210"
    window.cboElements.style.visibility = "visible"
    window.cboPrograms.style.visibility = "visible"
    Call window_onresize()
End Sub

Sub cmdFind_onclick()
    Dim oNode

    For Each oNode In PageBody.childNodes
        On Error Resume Next
            oNode.style.visibility = "hidden"
            oNode.style.cursor = "wait"
        On Error Goto 0
    Next
    lblStatus.innerText = "Searching for relevant questions..."
    lblStatus.style.visibility = "visible"
    Form.Program.value = Parse(cboPrograms.value, ":", 1)
    Form.Element.value = cboElements.value
    Form.Keywords.value = txtKeywords.value
    Form.FormAction.value = "Find"
    Form.submit
End Sub

Sub lblQuestionDetail_onclick(intQst)
    Dim strQst      'Build question ID for document.all reference.
    Dim intPrevious 'Holds the index of the previously selected question.

    intPrevious = mintSelectedQuestion

    'Do nothing if the item clicked was already selected.
    If intQst = intPrevious Then
        Exit Sub
    End If

    If intPrevious <> 0 Then
        strQst = "lblRowHdr" & intPrevious
        document.all(strQst).classname = "defLabel"
        strQst = "lblRowDet" & intPrevious
        document.all(strQst).classname = "defLabel"
    End If
    strQst = "lblRowHdr" & intQst
    document.all(strQst).classname = "defLabel SelectedRow"
    strQst = "lblRowDet" & intQst
    document.all(strQst).classname = "defLabel SelectedRow"

    mintSelectedQuestion = intQst
End Sub

Sub window_onresize()
    Dim oNode
    Dim intWidth

    intWidth = PageBody.clientWidth
    If intWidth - 25 > 220 Then
        For Each oNode In divQuestions.childNodes
            On Error Resume Next
            If Left(oNode.id,11) = "lblQuestion" Then
                oNode.style.width = intWidth - 25
            End If
            On Error Goto 0
        Next
        tblQuestions.style.width = intWidth - 25
    End If

    If intWidth - 20 >= 220 Then
        divHeader.style.width = intWidth - 20
        lblAppTitle.style.width = intWidth - 20
        divFind.style.width = intWidth - 20
    End If
    If ParseNumber(intWidth - 10 - 135) >= 80 Then
        cmdPrint.style.left = intWidth - 10 - 135
        cmdClose.style.left = intWidth - 10 - 65
    End If
End Sub

Function ParseNumber(strVal)
    Dim intI
    Dim strTemp
    
    strTemp = ""
    For intI = 1 To Len(strVal)
        If IsNumeric(Mid(strVal, intI, 1)) Then
            strTemp = strTemp & Mid(strVal, intI, 1)
        End If
    Next
    
    ParseNumber = strTemp
End Function
-->
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody style="background-color:white;overflow:scroll" bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5>
    <DIV id=divHeader
        style="BORDER-STYLE:solid; BORDER-WIDTH:1; BORDER-COLOR:#C0C0C0; HEIGHT:25; 
            WIDTH:220; TOP:5; LEFT:10; CURSOR:default;background-color:transparent">

        <SPAN id=lblAppTitle
            style="FONT-SIZE:12pt; HEIGHT:20; WIDTH:120; TOP:2; LEFT:0; TEXT-ALIGN:Center">
            <B>Case Review Guide</B>
        </SPAN>
    </DIV>

    <SPAN id=lblStatus style="visiblity:hidden;top:30;left:20;width:200;height:100"></SPAN>
    
    <DIV id=divFind
        style="HEIGHT:140; WIDTH:220; TOP:30; LEFT:10; border-style:solid;border-width:1;BORDER-COLOR:#C0C0C0">

        <SPAN id=lblPrograms class=DefLabel style="LEFT:10; WIDTH:200; TOP:5">
            <B>Select a Program</B>
            <SELECT id=cboPrograms style="LEFT:0; TOP:16; WIDTH:200; visibility:hidden" tabIndex=1 NAME="cboPrograms">
                <%Call WriteProgramList(Null,0,0,0)%>
            </SELECT>
        </SPAN>

        <SPAN id=lblElements class=DefLabel style="LEFT:10; WIDTH:200; TOP:50">
            <B>Select an Eligibility Element</B>
            <SELECT id=cboElements style="LEFT:0; TOP:16; WIDTH:200; visibility:hidden" tabIndex=1 NAME="cboElements">
                <OPTION VALUE=0 SELECTED>&ltAny&gt</option>
            </SELECT>
        </SPAN>
        
        <%'Get the list of elements and read into a hidden combobox, used to
        'populate visible element list when a program is selected.
        Set adRsElms = Server.CreateObject("ADODB.Recordset")
        Set gadoCmd = GetAdoCmd("spGetEligElemList")
            AddParmIn gadoCmd, "@ElmID", adVarChar, 50, NULL
            AddParmIn gadoCmd, "@PrgID", adVarChar, 255, NULL
            adRsElms.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
        Set gadoCmd = Nothing
        strOptions = ""
        Do While Not adRsElms.EOF
            strOptions = strOptions & "<OPTION VALUE=" & adRsElms.Fields("prgID").Value & ":" & adRsElms.Fields("elmID").Value & ">"
            strOptions = strOptions & adRsElms.Fields("elmShortTitle").Value & "^" 
            strOptions = strOptions & adRsElms.Fields("elmLongTitle").Value & "^" & adRsElms.Fields("prgElmCategory").Value
            adRsElms.MoveNext
        Loop
        adRsElms.Close
        'Hidden combobox for elements - cboElementList: 
        Response.Write "<SELECT id=cboElementList "
        Response.Write "style=""LEFT:0; TOP:0; VISIBILITY:hidden "" tabIndex=0>"
        Response.Write strOptions & "</SELECT>" & vbCrLf
        %>

        <SPAN id=lblKeywords class=DefLabel style="LEFT:10; WIDTH:200; TOP:95">
            <B>Or enter a Keyword</B>
            <INPUT type=text id=txtKeywords title="Enter keyword" 
                style="LEFT:0; TOP:16; WIDTH:200; text-align:left"
                tabIndex=1 maxlength=250 NAME="txtKeywords">
        </SPAN>
    </DIV>

    <BUTTON id=cmdFind title="Search for related questions" 
        style="LEFT:10; WIDTH:65; TOP:175; HEIGHT:20" 
        tabIndex=1>Find
    </BUTTON>
    <BUTTON id=cmdPrint title="Send report to the printer" 
        style="LEFT:95; WIDTH:65; TOP:175; HEIGHT:20" 
        tabIndex=55>Print
    </BUTTON>
    <BUTTON id=cmdClose title="Close window and return to report criteria screen" 
        style="LEFT:165; WIDTH:65; TOP:175; HEIGHT:20" 
        tabIndex=55>Close
    </BUTTON>

    <DIV id=divQuestions
        style="HEIGHT:191; TOP:210; LEFT:15">
        <%
        intQstCnt = 0
        strSpaces = "&nbsp&nbsp&nbsp&nbsp"
        If strAction = "Find" Then
            Set adCmd = Server.CreateObject("ADODB.Command")
            With adCmd
                .ActiveConnection = gadoCon
                .CommandTimeout = 180
                .CommandType = adCmdStoredProc
                .CommandText = "spFindQuestions"
                .Parameters.Append .CreateParameter("@Program", adInteger, adParamInput, 0, intPrg)
                .Parameters.Append .CreateParameter("@Element", adInteger, adParamInput, 0, intElm)
                .Parameters.Append .CreateParameter("@Keywords", adVarchar, adParamInput, 250, ReqIsBlank("Keywords"))
                'Call ShowCmdParms(adCmd) '***DEBUG
            End With

            Set adRs = Server.CreateObject("ADODB.Recordset")
            adRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
        
            Response.Write "<TABLE ID=tblQuestions Border=0 CellSpacing=0>"
            Response.Write "<TBODY>"
            Do While Not adRs.BOF And Not adRs.EOF
                intQstCnt = intQstCnt + 1
                strQstPrv = adRs.Fields(0).Value & adRs.Fields(1).Value
                Response.Write "<TR id=lblRowHdr" & intQstCnt & " class=defLabel onmouseover=lblQuestionDetail_onclick(" & intQstCnt & ") onclick=lblQuestionDetail_onclick(" & intQstCnt & ")>"
                Response.Write "<TD id=lblQuestionHeader" & intQstCnt & " class=""defLabel HeaderCell""> <b>" & adRs.Fields(0).Value & "</b> - " & adRs.Fields(1).Value & "<br></TD>"
                Response.Write "</TR>"
                Response.Write "<TR id=lblRowDet" & intQstCnt & " class=defLabel onmouseover=lblQuestionDetail_onclick(" & intQstCnt & ") onclick=lblQuestionDetail_onclick(" & intQstCnt & ")>"
                Response.Write "<TD id=lblQuestionDetail" & intQstCnt & " class=""defLabel DetailCell"">"
                If adRs.Fields(2).Value <> "" And Not IsNull(adRs.Fields(2).Value) Then
                    Response.Write strSpaces & "<b>Location:</b>" & "<br>"
                    Response.Write strSpaces & strSpaces & adRs.Fields(2).Value & "<br>"
                End If
                If adRs.Fields(4).Value <> "" And Not IsNull(adRs.Fields(4).Value) Then
                    Response.Write strSpaces & "<b>Manual Reference:</b>" & "<br>"
                    Response.Write strSpaces & strSpaces & adRs.Fields(4).Value & "<br>"
                End If
                    If Not IsNull(adRs.Fields(3).Value) And adRs.Fields(3).Value <> "" Then
                        Response.Write strSpaces & "<b>Factors:</b><br>"
                        Do While strQstPrv = adRs.Fields(0).Value & adRs.Fields(1).Value
                            Response.Write strSpaces & strSpaces & adRs.Fields(3).Value & "<br>"
                            strQstPrv = adRs.Fields(0).Value & adRs.Fields(1).Value
                            adRs.MoveNext
                            If adRs.EOF Or adRs.BOF Then
                                Exit Do
                            End If
                        Loop
                    Else
                        adRs.MoveNext
                    End If
                Response.Write "</TD><TR>"
            Loop
            Response.Write "</TBODY>"
            Response.Write "</TABLE>"
        End If
        Response.Write "<span id=lblQuestionCount style=""visibility:hidden"">" & intQstCnt & "</span>" & vbcrlf
        %>
    </DIV>
    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" TARGET=""_self"" STYLE=""VISIBILITY: hidden"" ACTION=""ReviewGuide.asp"" ID=Form1>" & vbCrLf
        Call CommonFormFields()
        WriteFormField "FormAction", ReqForm("FormAction")
        WriteFormField "Program", intPrg
        WriteFormField "Element", intElm
        WriteFormField "Keywords", ReqForm("Keywords")
    Response.Write Space(4) & "</FORM>" & vbCrLf
    %>

</BODY>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<!--#include file="IncProgramList.asp"-->