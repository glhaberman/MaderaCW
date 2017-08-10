<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
Dim adCmd
Dim adRs
Dim strSQL
Dim mstrPageTitle
Dim mstrPageHeading
Dim mintCnt
Dim oRpt
Dim mintProgramID
%>
<!--#include file="IncCnn.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
mstrPageHeading = "Elements and Causal Factors List"
If IsNumeric(Request.QueryString("ProgramID")) Then
    mintProgramID = CInt(Request.QueryString("ProgramID"))
    If mintProgramID = 0 Then
        mintProgramID = NULL
    End If
Else
    mintProgramID = NULL
End If

%>

<HTML>
<HEAD>
    <meta name=vs_targetSchema content="HTML 4.0"/>
    <TITLE><%=mstrPageHeading %></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <STYLE id=ThisPageStyles type="text/css">
        .RptDefault
            {
            FONT-SIZE: 8pt; 
            TEXT-ALIGN: center;
            background:#FFFFFF;
            border: none;
            }
        .RptDefaultLeft
            {
            FONT-SIZE: 8pt; 
            TEXT-ALIGN: left;
            background:#FFFFFF;
            border-bottom:1 solid #C0C0C0;
            border-top:1 solid #C0C0C0;
            border-left:1 solid #C0C0C0;
            border-right:1 solid #C0C0C0
            }
        .RptDefaultCenter
            {
            FONT-SIZE: 8pt; 
            TEXT-ALIGN: center;
            background:#FFFFFF;
            border: none;
            border-bottom:1 solid #C0C0C0
            }
        .RptDetailHdr
            {
            background:#DCDCDC;
            font-weight:bold;
            }
        .RptBold
            {
            font-weight:bold;
            }
        .RptBordTop
            {
            border-top:1 solid #C0C0C0;
            }
        .RptBordBot
            {
            border-bottom:1 solid #C0C0C0;
            }
        .RptBordLeft
            {
            border-left:1 solid #C0C0C0;
            }
        .RptBordRight
            {
            border-right:1 solid #C0C0C0;
            }
    </STYLE>
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim blnCloseClicked

Sub window_onload
	Call FormShow("none")
	PageBody.style.cursor = "wait"
	Call SizeAndCenterWindow(767, 520, True)
    Call FormShow("block")
    PageBody.style.cursor = "default"
    cmdPrint1.focus
End Sub

Sub cmdClose_onclick
	Window.close
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub FormShow(strVis)
	cmdPrint1.style.display = strVis
    cmdClose1.style.display = strVis
    cmdPrint2.style.display = strVis
    cmdClose2.style.display = strVis
    Header.style.display = strVis
    PageFrame.style.display = strVis
End Sub

Sub window_onbeforeprint()
    cmdPrint1.style.visibility = "hidden"
    cmdClose1.style.visibility = "hidden"
    cmdPrint2.style.visibility = "hidden"
    cmdClose2.style.visibility = "hidden"
    cmdCopy1.style.visibility = "hidden"
    cmdCopy2.style.visibility = "hidden"
End Sub

Sub window_onafterprint()
    cmdPrint1.style.visibility = "visible"
    cmdClose1.style.visibility = "visible"
    cmdPrint2.style.visibility = "visible"
    cmdClose2.style.visibility = "visible"
    cmdCopy1.style.visibility = "visible"
    cmdCopy2.style.visibility = "visible"
End Sub

Sub cmdCopy_onclick()
    Dim CtlRng

    'A controlRange object is used to select the results div, then copy it:
    Set CtlRng = PageBody.createControlRange()
    CtlRng.AddElement(Header)
    CtlRng.AddElement(PageFrame)
    CtlRng.Select
    CtlRng.execCommand("Copy")
    Set CtlRng = Nothing
    'Clear the selection:
    document.selection.empty
    MsgBox "Report copied to clipboard.", ,"Copy Report"
End Sub

Sub SizeAndCenterWindow(intX, intY, blnCenter)
    Dim intScrLeft
    Dim intScrTop
    Dim intScrAvlX
    Dim intScrAvlY
    Dim blnMove

    On Error Resume Next    
    window.resizeTo intX, intY
    
    intScrAvlX = window.screen.availWidth
    intScrAvlY = window.screen.availHeight

    If blnCenter Then
        'Center the window in the screen:
        window.moveTo (intScrAvlX - intX)/2, (intScrAvlY - intY)/2
    Else
        'Only move the window if it is not completely on screen:
        blnMove = False
        intScrLeft = Cint(window.screenLeft)
        intScrTop = Cint(window.screenTop - 30)
        If intScrLeft < 0 Then
            intScrLeft = 0
            blnMove = True
        End If
        If (intScrLeft + intX) > intScrAvlX Then
            intScrLeft = intScrLeft - Abs(intScrAvlX - (intScrLeft + intX))
            blnMove = True
        End If
        If intScrTop < 0 Then
            intScrTop = 0
            blnMove = True
        End If
        If (intScrTop + intY) > intScrAvlY Then
            intScrTop = intScrTop - Abs(intScrAVlY - (intScrTop + intY))
            blnMove = True
        End If 
        If blnMove Then
            window.moveTo intScrLeft, intScrTop
        End If
    End If

    On Error Goto 0
End Sub
-->
</SCRIPT>

<!--=== Start of Report Definition and Layout ============================= -->
<BODY id=Pagebody style="background-color:white; overflow:auto; font-size:10pt" bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5>

    <BUTTON id=cmdPrint1 title="Send report to the printer" 
        style="LEFT:20; WIDTH:65; HEIGHT:20" 
        onclick="cmdPrint_onclick"
        tabIndex=55>Print
    </BUTTON>

    <BUTTON id=cmdCopy1 title="Copy report to clipboard" 
        style="LEFT:95; WIDTH:65; HEIGHT:20" 
        onclick="cmdCopy_onclick"
        tabIndex=55>Copy
    </BUTTON>

    <BUTTON id=cmdClose1 title="Close report preview window" 
        style="LEFT:625; WIDTH:65; HEIGHT:20" 
        onclick="cmdClose_onclick"
        tabIndex=55>Close
    </BUTTON>
    
    <br style="font-size:18"/>
    <table id=Header rules=none width=700 cellspacing=0 border=0 bordercolor=#C0C0C0 bordercolorlight=#C0C0C0 bordercolordark=#C0C0C0>
        <tbody>
            <tr valign=top>
                <td style="height:18;border-bottom:1 solid #c0c0c0">
                    <SPAN id=lblOrg class=DefLabel
                        style="COLOR:#A9A9A9; FONT-SIZE:8pt; HEIGHT:20; WIDTH:648; LEFT:4; TEXT-ALIGN:left; FONT-WEIGHT:bold">
                        <%=mstrPageTitle%>
                    </SPAN>
                </td>
                <td style="height:18;border-bottom:1 solid #c0c0c0">
                    <SPAN id=lblDate class=DefLabel
                        style="COLOR:#A9A9A9; FONT-SIZE:8pt; HEIGHT:20; WIDTH:230; LEFT:408; TEXT-ALIGN:right; FONT-WEIGHT: bold">
                        <% = "Date Printed: " & FormatDateTime(Now,vbGeneralDate) %>
                    </SPAN>
                </td>
            </tr>
            <tr valign=top>
               <td colspan=2 style="height:28; text-align:center">
                    <SPAN id=lblAppTitle class=DefLabel
                        style="FONT-SIZE:14pt; HEIGHT:20; WIDTH:648; LEFT:4; TEXT-ALIGN:center">
                        <b><%=mstrPageHeading %></b>
                    </SPAN>
               </td>
            </tr>
        </tbody>
    </table>
    <%

    'Instance the helper class:
    Set oRpt = New clsReport

    'Start the HTML for the table layout.
    oRpt.WriteTblStart("PageFrame")
    
    Set adCmd = Server.CreateObject("ADODB.Command")
    With adCmd
        .ActiveConnection = gadoCon
        .CommandType = adCmdStoredProc
        .CommandText = "spRptElementFactorList"
        .Parameters.Append .CreateParameter("@ProgramID", adInteger, adParamInput, 0, mintProgramID)
    End With

    Set adRs = Server.CreateObject("ADODB.Recordset") 
    Call adRs.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)

    oRpt.WriteColumnHeaders()
    oRpt.WriteTblBodyStart()
    Do While Not adRs.EOF
        Response.Write "<TR>" & vbCrLf
        Response.Write "<TD class=""RptDefaultLeft"">" & adRs.Fields("Program").Value & "</TD>" & vbCrLf
        Response.Write "<TD class=""RptDefaultLeft"">" & adRs.Fields("Element").Value & "</TD>" & vbCrLf
        Response.Write "<TD class=""RptDefaultLeft"">" & adRs.Fields("Factor").Value & "</TD>" & vbCrLf
        Response.Write "</TR>" & vbCrLf
        Response.Flush
        adRs.movenext
    Loop
        
    oRpt.WriteTblBlankRow()
    oRpt.WriteTblBodyEnd()
    oRpt.WriteTblEnd()
    %>

    <br/>
    <BUTTON id=cmdPrint2 title="Send report to the printer" 
        style="LEFT:20; WIDTH:65; HEIGHT:20" 
        onclick="cmdPrint_onclick"
        tabIndex=55>Print
    </BUTTON>

    <BUTTON id=cmdCopy2 title="Copy report to clipboard" 
        style="LEFT:95; WIDTH:65; HEIGHT:20" 
        onclick="cmdCopy_onclick"
        tabIndex=55>Copy
    </BUTTON>

    <BUTTON id=cmdClose2 title="Close report preview window" 
        style="LEFT:625; WIDTH:65; HEIGHT:20" 
        onclick="cmdClose_onclick"
        tabIndex=55>Close
    </BUTTON>

    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY: hidden"" ACTION=""Reports.asp"" ID=Form1>" & vbCrLf
        Call CommonFormFields()

        Call ReportFormDef()
    Response.Write Space(4) & "</FORM>" & vbCrLf

    adRs.Close
    Set adRs = Nothing
    Set gadoCmd = Nothing
    gadoCon.Close
    Set gadoCon = Nothing
    %>
</BODY>
</HTML>
<%
Class clsReport
    Public Sub WriteColumnHeaders()
        'Writes out the column header row.
        Response.Write "<THEAD style=""display:table-header-group"">" & vbCrLf
        Response.Write "<TR>" & vbCrLf
        Response.Write "<TD class=""RptDefault RptDetailHdr RptBordTop RptBordBot RptBordLeft RptBordRight"">Program</TD>" & vbCrLf
        Response.Write "<TD class=""RptDefault RptDetailHdr RptBordTop RptBordBot RptBordLeft RptBordRight"">Element</TD>" & vbCrLf
        Response.Write "<TD class=""RptDefault RptDetailHdr RptBordTop RptBordBot RptBordLeft RptBordRight"">Factor</TD>" & vbCrLf
        Response.Write "</TR>" & vbCrLf
        Response.Write "</THEAD>" & vbCrLf
    End Sub

    Public Sub WriteTblStart(strID)
        Response.Write vbCrLf & "<TABLE id=""" & strID & """ rules=none cellspacing=0 border=0 "
        Response.Write "bordercolor=#C0C0C0 bordercolorlight=#C0C0C0 bordercolordark=#C0C0C0 "
        Response.Write "style=""page-break-inside:auto"">" & vbCrLf
    End Sub

    Public Sub WriteTblHeader()
        Response.Write "<THEAD>" & vbCrLf
        Response.Write "<TR valign=top>" & vbCrLf
        Response.Write "<TH width=225></TH>" & vbCrLf
        Response.Write "</TR>" & vbCrLf
        Response.Write "</THEAD>" & vbCrLf
    End Sub

    Public Sub WriteTblBodyStart()
        Response.Write "<TBODY>" & vbCrLf
    End Sub

    Public Sub WriteTblBodyEnd()
        Response.Write "</TBODY>" & vbCrLf
    End Sub
    
    Public Sub WriteTblEnd()
        Response.Write "</TABLE>" & vbCrLf
    End Sub
    
    Public Sub WriteTblBlankRow()
        Response.Write "<TR><TD colspan=6 class=""RptDefault"">&nbsp;</TD><TR>" & vbCrLf 
    End Sub

    Public Function HTMLSpace(intCount)
        Dim strIndent
        Dim intI
        strIndent = ""
        For intI = 1 To intCount
            strIndent = strIndent & "&nbsp;"
        Next
        HTMLSpace = strIndent
    End Function
End Class
%>
 
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
