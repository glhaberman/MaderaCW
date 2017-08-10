<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: PrintReview.asp                                                 '
'  Purpose: This page is used to print a case review detail from the add or '
'           edit screens.                                                   '
' Includes:                                                                 '

'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim mlngReReviewID
Dim mstrPageTitle
Dim madoReReview
Dim madoReReviewElms, madoFactors, madoComments
Dim adCmd
Dim mstrSubmitted
Dim intCol1, intCol2,intCol3, intCol4

mlngReReviewID = Request.QueryString("ReReviewID")
mstrPageTitle = Request.QueryString("PageTitle")

Response.ExpiresAbsolute = Now - 5
'Retrieve the Re-Review to display:
Set madoReReview = Server.CreateObject("ADODB.Recordset")
Set madoReReviewElms = Server.CreateObject("ADODB.Recordset")
Set adCmd = GetAdoCmd("spReReviewGet4Print")
    AddParmIn adCmd, "@ReReviewID", adInteger, 0, mlngReReviewID
    'Call ShowCmdParms(adCmd) '***DEBUG
    madoReReview.Open adCmd, , adOpenForwardOnly, adLockReadOnly

Set madoReReviewElms = madoReReview.NextRecordset
Set madoFactors = madoReReview.NextRecordset
Set madoComments = madoReReview.NextRecordset
If madoReReview("Submitted").Value="Y" Then
    mstrSubmitted = "Yes"
Else
    mstrSubmitted = "No"
End If

%>
<HTML><HEAD>
<TITLE><%="Print " & gstrEvaluation%></TITLE>
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <!--#include file="IncDefStyles.asp"-->
    <STYLE id=ThisPageStyles type="text/css" rel="stylesheet">
        
        .ReportText
            {
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: left;
            }
        
        .ColumnHeading
            {
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: center;
            FONT-WEIGHT: bold;
            BORDER-COLOR:#C0C0C0;
            BORDER-BOTTOM-STYLE: solid;
            BORDER-BOTTOM-WIDTH: 1;
            }

        .ReportCaption
            {
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: left;
            }
            
        .ReportCaptionCenter
            {
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: center;
            FONT-WEIGHT: bold;
            }
    </STYLE>
</HEAD>

<SCRIPT LANGUAGE=vbscript>

Sub window_onload()
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub cmdClose_onclick()
    window.close
End Sub

Sub window_onbeforeprint()
    cmdPrint.style.visibility = "hidden"
    cmdClose.style.visibility = "hidden"
    cmdExport1.style.visibility = "Hidden"
End Sub

Sub window_onafterprint()
    cmdPrint.style.visibility = "visible"
    cmdClose.style.visibility = "visible"
    cmdExport1.style.visibility = "visible"
End Sub
Sub cmdExport_onclick()
    Dim CtlRng
    'If the results div is not empty, copy it's contents to the clipboard:
    If PageFrame.children.length > 0 Then
        'A controlRange object is used to select the results div, then copy it:
        Set CtlRng = PageBody.createControlRange()
        CtlRng.AddElement(PageFrame)
        CtlRng.Select
        CtlRng.execCommand("Copy")
        Set CtlRng = Nothing
        'Clear the selection:
        document.selection.empty
        MsgBox "Results copied to clipboard.", ,"Copy Results"
    End If
End Sub
</SCRIPT>

<BODY id=PageBody style="BACKGROUND-COLOR: #white; overflow:auto" 
    bottomMargin=10 
    leftMargin=10 
    topMargin=10
    rightMargin=10>

<DIV id=PageFrame
   style="LEFT:0; WIDTH:700; FONT-SIZE:12pt; TOP:0; HEIGHT:425">
        <TABLE id=tblButtons>
            <TR id=tbrButtons>
                <TD style="width:65"><INPUT TYPE="button" VALUE="Print" onClick="cmdPrint_onclick" style="width:62;height:23" ID="Button1" NAME="cmdPrint"></TD>
                <TD style="width:65"><INPUT TYPE="button" VALUE="Copy" onClick="cmdExport_onclick" style="width:62;height:23" ID="Button2" NAME="cmdExport1"></TD>
                <TD style="width:475">&nbsp;</TD>
                <TD style="width:65"><INPUT TYPE="button" VALUE="Close" onClick="cmdClose_onclick" style="width:62;height:23" ID="Button3" NAME="cmdClose"></TD>
            </TR> 
        </TABLE>
        <BR>
        <TABLE id=tblHeader>
            <TR id=tbrHeader1>
                <TD id=tbcHeader1Col1 class=ReportCaption style="width:170">&nbsp;</TD>
                <TD id=tbcHeader1Col2 class=ReportCaption style="color:#a9a9a9;width:330;text-align:center"><B><%=mstrPageTitle%></B></TD>
                <TD id=tbcHeader1Col3 class=ReportCaption style="color:#a9a9a9;width:180;text-align:right;font-size:8pt">Printed: <%=Now()%></TD>
            </TR> 
            <TR id=tbrHeader2>
                <TD id=tbcHeader2Col1 class=ReportCaption style="width:170">&nbsp;</TD>
                <TD id=tbcHeader2Col2 class=ReportCaption style="font-size:14pt;width:330;text-align:center"><B><%="Print &nbsp" & gstrEvaluation%></B></TD>
                <TD id=tbcHeader2Col3 class=ReportCaption style="width:180">&nbsp;</TD>
            </TR> 
            <TR id=tbrHeader3>
                <TD id=tbcHeader3Col1 class=ReportCaption style="width:170"><B><%=gstrEvaluation & " ID: " & madoReReview("rrvID").Value%></B></TD>
                <TD id=tbcHeader3Col2 class=ReportCaption style="width:330;text-align:center"><B><%="Date: " & madoReReview("rrvDateEntered").Value%></B></TD>
                <TD id=tbcHeader3Col3 class=ReportCaption style="width:180;text-align:right"><B><%="Status: " & madoReReview("Status").Value%></B></TD>
            </TR> 
        </TABLE>
        
        <HR>
        <TABLE id=tblReReview>
            <TR id=tbrReReview1>
                <TD id=tbcReReview1Col1 class=ReportCaption style="width:100"><B><%=gstrEvaTitle%>:<B></TD>
                <TD id=tbcReReview1Col2 class=ReportCaption style="width:235"><%=madoReReview("rrvReReviewer").Value%></TD>
                <TD id=tbcReReview1Col3 class=ReportCaption style="width:100"><B>Submitted:</B></TD>
                <TD id=tbcReReview1Col4 class=ReportCaption style="width:235"><%=mstrSubmitted%></TD>
            </TR> 
        </TABLE>
         <HR>
        <%
        intCol1 = 95
        intCol2 = 215
        intCol3 = 140
        intCol4 = 220
        %>
        <TABLE id=tblOrgReviewElms>
            <TR>
                <TD class=ReportCaption style="width:<%=intCol1%>"><B>Review ID:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol2%>"><%=madoReReview("rrvOrgReviewID").Value%></TD>
                <TD class=ReportCaption style="width:<%=intCol3%>"><B>Review Status:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol4%>"><%=madoReReview("Status").Value%></TD>
            </TR> 
            <TR>
                <TD class=ReportCaption style="width:<%=intCol1%>"><B>Review Date:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol2%>"><%=madoReReview("rvwDateEntered").Value%></TD>
                <TD class=ReportCaption style="width:<%=intCol3%>"><B>Review Month:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol4%>"><%=madoReReview("rvwMonthYear").Value%></TD>
            </TR> 
            <TR>
                <TD class=ReportCaption style="width:<%=intCol1%>"><B>Review Class:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol2%>"><%=madoReReview("ReviewClass").Value%></TD>
                <TD class=ReportCaption style="width:<%=intCol3%>"><B>Reviewer:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol4%>"><%=Parse(madoReReview("rvwReviewerName").Value,"--",1)%></TD>
            </TR> 
            <TR>
                <TD class=ReportCaption style="width:<%=intCol1%>"><B><%=gstrWkrTitle%>:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol2%>"><%=Parse(madoReReview("rvwWorkerName").Value,"--",1)%></TD>
                <TD class=ReportCaption style="width:<%=intCol3%>"><B><%=gstrWkrTitle%> Response:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol4%>"><%=madoReReview("WorkerResponse").Value%></TD>
            </TR> 
            <TR>
                <TD class=ReportCaption style="width:<%=intCol1%>"><B>Case Name:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol2%>"><%=madoReReview("CaseName").Value%></TD>
                <TD class=ReportCaption style="width:<%=intCol3%>"><B>Case Number:</B></TD>
                <TD class=ReportCaption style="width:<%=intCol4%>"><%=madoReReview("rvwCaseNumber").Value%></TD>
            </TR> 
        </TABLE>
         <HR>
        <SPAN id=lblElements
            class=ReportText
            style="LEFT:5; WIDTH:640; TEXT-ALIGN:left">
<%
            Dim intI
            Dim strComments
            
            madoReReviewElms.Filter = "rreStatusID>0"
            Do While Not madoReReviewElms.EOF
                Response.Write "<B>" & gstrEvaluation & " Program:</B>&nbsp;&nbsp;" & madoReReviewElms.Fields("prgShortTitle").Value & " -- " & madoReReviewElms.Fields("ReviewType").Value & "<BR>"
                Select Case madoReReviewElms.Fields("rreTypeID").Value
                    Case 1
                        strComments = ConvertCRLFToBR(CleanText(madoReReviewElms.Fields("rveComments").Value))
                        Response.Write "<B>Review Action:</B>&nbsp;&nbsp;" & madoReReviewElms.Fields("Element").Value & "&nbsp;&nbsp;<B>Review Status:</B>&nbsp;&nbsp;" & madoReReviewElms.Fields("ElementStatus").Value & "<BR>"
                    Case 2
                        madoComments.Filter = "rvcScreenName='" & madoReReviewElms.Fields("Element").Value & "'"
                        strComments = ""
                        If madoComments.RecordCount = 1 Then
                            strComments = ConvertCRLFToBR(CleanText(madoComments.Fields("rvcComments").Value))
                        End If
                        madoFactors.Filter = "rvfProgramID=" & madoReReviewElms.Fields("prgID").Value & " AND rvfElementID=" & madoReReviewElms.Fields("rreElementID").Value & " AND rvfFactorID=" & madoReReviewElms.Fields("rreFactorID").Value
                        If madoFactors.RecordCount = 1 Then
                            madoFactors.MoveFirst
                            Response.Write "<B>Review Element/Causal Factor:</B>&nbsp;&nbsp;" & madoReReviewElms.Fields("Element").Value & "-" & madoReReviewElms.Fields("fctShortName").Value & "&nbsp;&nbsp;<B>Review Status:</B>&nbsp;&nbsp;" & madoFactors.Fields("FactorStatus").Value & "<BR>"
                        Else
                            Response.Write "<B>Review Element/Causal Factor:</B>&nbsp;&nbsp;" & madoReReviewElms.Fields("Element").Value & "-" & madoReReviewElms.Fields("fctShortName").Value & "&nbsp;&nbsp;<B>Review Status:</B>&nbsp;&nbsp;<BR>"
                        End If
                    Case 3
                        strComments = ConvertCRLFToBR(CleanText(madoReReviewElms.Fields("rveComments").Value))
                        Response.Write "<B>Review Question:</B>&nbsp;&nbsp;" & madoReReviewElms.Fields("Element").Value & "&nbsp;&nbsp;<B>Review Status:</B>&nbsp;&nbsp;" & madoReReviewElms.Fields("ElementStatus").Value & "<BR>"
                End Select
                Response.Write "<B>Review Comments:</B>&nbsp;&nbsp;" & strComments & "<BR>"
                Response.Write "<B>" & gstrEvaluation & " Status:</B>&nbsp;&nbsp;" & madoReReviewElms.Fields("rreStatus").Value & "<BR>"
                Response.Write "<B>" & gstrEvaluation & " Comments:</B>&nbsp;&nbsp;" & ConvertCRLFToBR(CleanText(madoReReviewElms.Fields("rreComments").Value)) & "<BR>"
                Response.Write "<BR><BR>"

                madoReReviewElms.MoveNext
            Loop
%>
        </SPAN>      
    </DIV>
</BODY>
</HTML>
<%
Function CleanText(strText)
    'This function is used to replace quote and double-quote characters with 
    'tokens when sending to the database, and replace the tokens with the
    'correct characters when retrieving from the database. The tokens used 
    'are {TAB}#sq# for single-quote
    '    {TAB}#dq# for double-quote

    If IsNull(strText) Then
        CleanText = ""
    Else 'Apostrophe Or double quotes:
        strText = Replace(strText, Chr(9) & "#sq#", "'")
        strText = Replace(strText, Chr(9) & "#dq#", """")
        strText = Replace(strText, Chr(9) & "#ca#", "^")
        strText = Replace(strText, Chr(9) & "#ba#", "|")
        CleanText = strText
    End If
End Function
Function ConvertCRLFToBR(strText)
    Dim strTemp
    Dim intI
    
    strTemp = ""
    For intI = 1 To Len(strText)
        If Asc(Mid(strText, intI, 1)) = 13 Then
            strTemp = strTemp & "<BR>"
        End If
        If Asc(Mid(strText, intI, 1)) <> 10 Then
            strTemp = strTemp & Mid(strText, intI, 1)
        End If
    Next
    ConvertCRLFToBR = strTemp
End Function
%>