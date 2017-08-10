<%
    Dim strRHReportTitle, strRHUserID, strRHProgramText, dtmRHStartDate, dtmRHEndDate
    If Request.Form("ReportTitle") <> "" Then
        strRHReportTitle = Request.Form("ReportTitle")
        strRHUserID = ReqForm("UserID")
        strRHProgramText = ReqForm("ProgramText")
        dtmRHStartDate = ReqForm("StartDate")
        dtmRHEndDate = ReqForm("EndDate")
    Else
        strRHReportTitle = mstrReportTitle
        strRHUserID = maCriteria(4)
        strRHProgramText = mstrProgramText
        dtmRHStartDate = maCriteria(5)
        dtmRHEndDate = maCriteria(6)
    End If
%>
<BODY id=Pagebody style="background-color:white" bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style=overflow:auto>

	<SPAN id=lblMessage	    
			style="COLOR:#C0C0C0; FONT-SIZE:8pt; HEIGHT:20; WIDTH:648; TOP:3; LEFT:8; TEXT-ALIGN:Left; FONT-WEIGHT:bold; visibility:hidden">
		Returning to Report Criteria Screen        
    </SPAN>

     <BUTTON id=cmdPrint1 title="Send report to the printer" 
        style="LEFT:20; WIDTH:65; TOP:10; HEIGHT:20" 
        onclick="cmdPrint_onclick"
        tabIndex=55>Print
    </BUTTON>
     <BUTTON id=cmdExport1 title="Export data from report to clipboard" 
        style="LEFT:95; WIDTH:65; TOP:10; HEIGHT:20" disabled
        onclick="cmdExport_onclick"
        tabIndex=55>Export
    </BUTTON>
    <BUTTON id=cmdClose1 title="Close window and return to report criteria screen" 
        style="LEFT:595; WIDTH:65; TOP:10; HEIGHT:20" 
        onclick="cmdClose_onclick"
        tabIndex=55>Close
    </BUTTON>
    
    <DIV id=divWaiting 
        style="BACKGROUND-COLOR: #FFFFCC;HEIGHT: 600; WIDTH: 770; TOP: 0; LEFT: -1000">
        <%
            Response.Write "<BR><BR><BR><BR>Copying <B>" & strRHReportTitle & "</B> data to clipboard, Please wait..."
            
        %>
    </DIV>

    <DIV id=Header class=RptHeader>

        <SPAN id=lblOrg class=DefLabel
            style="COLOR:#A9A9A9; FONT-SIZE:8pt; HEIGHT:20; WIDTH:300; TOP:3; LEFT:4; TEXT-ALIGN:left; FONT-WEIGHT:bold">
            <%=mstrPageTitle%>
        </SPAN>

        <SPAN id=lblDate class=DefLabel
            style="COLOR:#A9A9A9; FONT-SIZE:8pt; HEIGHT:20; WIDTH:430; TOP:3; LEFT:210; TEXT-ALIGN:right; FONT-WEIGHT: bold">
            <% = "Date Printed: " & FormatDateTime(Now,vbGeneralDate) & " (" & strRHUserID & ")" %>
        </SPAN>

        <SPAN id=lblAppTitle class=DefLabel
            style="FONT-SIZE:14pt; HEIGHT:20; WIDTH:648; TOP:23; LEFT:4; TEXT-ALIGN:center">
            <b><%=strRHReportTitle%> :</b> &nbsp &nbsp <%=strRHProgramText%>
        </SPAN>

        <SPAN id=lblReportDates class=ReportText
            style="WIDTH:350;TOP:55;LEFT:0;padding-right:0;text-align:left">
            <B>From Review Date:&nbsp</B>
            <%
            If Trim(dtmRHStartDate) = "" Then
                Response.Write "&ltAll&gt" & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>To:</B>&nbsp;&nbsp;" & "&ltAll&gt"
            Else
                Response.Write dtmRHStartDate & "&nbsp;&nbsp;<B>To:</B>&nbsp;&nbsp;" & dtmRHEndDate
            End If
            %>
        </SPAN>
    </DIV>

