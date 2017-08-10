<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: RptWorkerNoRespDetail.asp                                       '
'  Purpose: This page is used to display unsubmitted reviews for a worker.  '
' Includes:                                                                 '
'   IncCnn.asp          - ADO database connection                           '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim lngReviewID, lngReReviewID, lngReReviewTypeID
Dim mstrReReviewType
Dim adRs
Dim intI
Dim mdctAudit, oAudit

lngReReviewTypeID = Request.QueryString("ReReviewTypeID")
' A lngReReviewTypeID of -1 indicates all types should be included
lngReviewID = Request.QueryString("ReviewID")
lngReReviewID = Request.QueryString("ReReviewID")
If Len(lngReReviewID) = 0 Then lngReReviewID = 0

If lngReReviewTypeID = 0 Then
    mstrReReviewType = gstrEvaluation
ElseIf lngReReviewTypeID = 1 Then
    mstrReReviewType = "CAR"
Else
    mstrReReviewType = ""
End If

Set adRs = Server.CreateObject("ADODB.Recordset")
Set gadoCmd = GetAdoCmd("spReReviewHistory")
    AddParmIn gadoCmd, "@ReviewID", adInteger, 0, lngReviewID
    'Call ShowCmdParms(gadoCmd) '***DEBUG
Set adRs = GetAdoRs(gadoCmd)

adRs.Sort = "rrvID"
Set mdctAudit = CreateObject("Scripting.Dictionary")
intI = 0
Do While Not adRs.EOF
    If adRs("RRType").Value = mstrReReviewType Or mstrReReviewType = "" Then
        mdctAudit.Add CLng(intI), adRs("rrvID").Value & "^" & _
            adRs("Program").Value & "^" & _
            adRs("ReReviewer").Value & "^" & _
            adRs("DateEntered").Value & "^" & _
            adRs("RRType").Value & "^" & _
            adRs("Submitted").Value
        intI = intI + 1
    End If
    adRs.MoveNext
Loop
%>
<HTML><HEAD>
    <TITLE><%=mstrReReviewType%> History</TITLE>
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <STYLE id=ThisPageStyles type="text/css" rel="stylesheet">
        BODY
            {
            margin:1;
            position: absolute; 
            FONT-SIZE: 10pt; 
            FONT-FAMILY: Tahoma; 
            OVERFLOW: auto; 
            BACKGROUND-COLOR: #FFFFCC
            }
    </STYLE>
    <!--#include file="IncTableStyles.asp"-->
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Option Explicit

Sub window_onload
    Dim mdctAudit
    Set mdctAudit = CreateObject("Scripting.Dictionary")
<%
    If mdctAudit.Count > 0 Then
        For Each oAudit In mdctAudit
            Response.Write "mdctAudit.Add CLng(" & oAudit & "),""" & mdctAudit(oAudit) & """" & vbCrLf
        Next
    Else
        Response.Write "mdctAudit.Add 0,""0^0^^^^"""
    End If    
%>
    window.returnvalue = mdctAudit
    window.close
End Sub

Sub cmdClose_onclick()
    window.close
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub window_onbeforeprint()
    cmdClose1.style.visibility = "hidden"
    cmdPrint1.style.visibility = "hidden"
    cmdClose2.style.visibility = "hidden"
    cmdPrint2.style.visibility = "hidden"
End Sub

Sub window_onafterprint()
    cmdClose1.style.visibility = "visible"
    cmdPrint1.style.visibility = "visible"
    cmdClose2.style.visibility = "visible"
    cmdPrint2.style.visibility = "visible"
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:white;" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
    <BUTTON id=cmdPrint1 title="Send report to the printer" 
        style="LEFT:10; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdPrint_onclick"
        tabIndex=55>Print
    </BUTTON>
     <BUTTON id=cmdExport1 title="Copy data from report to clipboard" 
        style="LEFT:95; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdExport_onclick"
        tabIndex=55>Copy
    </BUTTON>
    <BUTTON id=cmdClose1 title="Close window and return to report criteria screen" 
        style="LEFT:525; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdClose_onclick"
        tabIndex=55>Close
    </BUTTON>

    <SPAN id=lblHeading class=DefLabel
        style="FONT-SIZE:12pt; HEIGHT:20; WIDTH:590; LEFT:10;top:20;TEXT-ALIGN:center;position:absolute">
        <B><%=mstrReReviewType%> History for ID <%=lngReReviewID%></B>
    </SPAN>

    <%
    Response.Write "<DIV id=divRvwHistory Class=TableDivArea style=""LEFT:1; TOP:260; WIDTH:615; HEIGHT:155; "
    Response.Write "    OVERFLOW:auto; FONT-WEIGHT:normal;z-index:1200"">"
    Response.Write "    <TABLE id=tblReview Border=0 Rules=rows Width=570 CellSpacing=0 "
    Response.Write "        Style=""position:absolute;overflow: hidden; width:590;left:2"">"
    Response.Write "        <THEAD id=tbhReview style=""height:17"">"
    Response.Write "            <TR id=thrReview>"
    Response.Write "                <TD class=CellLabel id=thcReviewC01 style=""font-size:10pt"">" & mstrReReviewType & " ID</TD>"
    Response.Write "                <TD class=CellLabel id=thcReviewC11 style=""font-size:10pt"">Program</TD>"
    Response.Write "                <TD class=CellLabel id=thcReviewC21 style=""font-size:10pt"">" & gstrEvaTitle & "</TD>"
    Response.Write "                <TD class=CellLabel id=thcReviewC31 style=""font-size:10pt"">" & mstrReReviewType & " Date</TD>"
    Response.Write "                <TD class=CellLabel id=thcReviewC51 style=""font-size:10pt"">Type</TD>"
    Response.Write "                <TD class=CellLabel id=thcReviewC41 style=""font-size:10pt"">Submitted</TD>"
    Response.Write "            </TR>"
    Response.Write "        </THEAD>"
    Response.Write "        <TBODY id=tbdReview>"
    Dim intJ
    intJ = 2
    Do While Not adRs.EOF
        Response.Write "<TR id=tdrReview" & intJ & ">" & vbCrLf
        Response.Write "    <TD class=TableDetail id=tdcReviewC0" & intJ & " style=""font-size:10pt;text-align:center;"">a" & adRs.Fields("rrvID").Value & "</TD>" & vbCrLf
        Response.Write "    <TD class=TableDetail id=tdcReviewC1" & intJ & " style=""font-size:10pt;text-align:center;"">b" & adRs.Fields("Program").Value & "</TD>" & vbCrLf
        Response.Write "    <TD class=TableDetail id=tdcReviewC2" & intJ & " style=""font-size:10pt;text-align:center;"">c" & adRs.Fields("ReReviewer").Value & "</TD>" & vbCrLf
        Response.Write "    <TD class=TableDetail id=tdcReviewC3" & intJ & " style=""font-size:10pt;text-align:center;"">d" & adRs.Fields("DateEntered").Value & "</TD>" & vbCrLf
        Response.Write "    <TD class=TableDetail id=tdcReviewC5" & intJ & " style=""font-size:10pt;text-align:center;"">e" & adRs.Fields("RRType").Value & "</TD>" & vbCrLf
        Response.Write "    <TD class=TableDetail id=tdcReviewC4" & intJ & " style=""font-size:10pt;text-align:center;"">f" & adRs.Fields("Submitted").Value & "</TD>" & vbCrLf
        Response.Write "</TR>" & vbCrLf
        intJ = intJ + 1
        adRs.MoveNext
    Loop
    Response.Write "</TBODY>"
    Response.Write "</TABLE>"
    Response.Write "</DIV>"
    %>

    <DIV id=divFooter Class=TableDivArea style="LEFT:1; TOP:425; WIDTH:615; HEIGHT:30; 
        OVERFLOW:auto; FONT-WEIGHT:normal;z-index:1200;background-color:transparent;border-style:none">
        <TABLE id=tblButtons Width=570 CellSpacing=0 
            Style="position:absolute;overflow: hidden; width:590;left:2">
            <TBODY id=tblBtnBody>
                <TR id=tfrReview>
                    <TD><INPUT TYPE="button" VALUE="Print" onClick="cmdPrint_onclick" style="width:62;height:23" ID=cmdPrint2 NAME="cmdPrint2"></TD>
                    <TD><INPUT TYPE="button" VALUE="Copy" onClick="cmdExport_onclick" style="width:62;height:23" ID=cmdExport2 NAME="cmdExport2"></TD>
                    <TD><DIV style="width:192;height:23;background-color:transparent;">&nbsp;</DIV></TD>
                    <TD><DIV style="width:192;height:23;background-color:transparent;">&nbsp;</DIV></TD>
                    <TD><INPUT TYPE="button" VALUE="Close" onClick="cmdClose_onclick" style="width:62;height:23" ID=cmdClose2 NAME="cmdClose2"></TD>
                </TR>
            </TBODY>
        </TABLE>
    </DIV>
    <%
    Response.Write "</BODY>"
    Response.Write "</HTML>"
    adRs.Close
    %>
<!--#include file="IncCmnCliFunctions.asp"-->
