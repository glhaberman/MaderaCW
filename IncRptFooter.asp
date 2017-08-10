    <br> 
    <BUTTON id=cmdPrint2 title="Send report to the printer" 
        style="LEFT:20; WIDTH:65; HEIGHT:20" 
        onclick="cmdPrint_onclick"
        tabIndex=55>Print
    </BUTTON>
    <BUTTON id=cmdExport2 title="Export data from report to clipboard" 
        style="LEFT:95; WIDTH:65; HEIGHT:20" 
        onclick="cmdExport_onclick" disabled
        tabIndex=55>Export
    </BUTTON>
    <BUTTON id=cmdClose2 title="Close window and return to report criteria screen" 
        style="LEFT:595; WIDTH:65; HEIGHT:20" 
        onclick="cmdClose_onclick"
        tabIndex=55>Close
    </BUTTON>
    </DIV>

    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" style=""VISIBILITY: hidden"" ACTION=""Reports.asp"" ID=Form>" & vbCrLf
        Call CommonFormFields()
		Call ReportFormDef()
    Response.Write Space(4) & "</FORM>" & vbCrLf

    If adRs.State <> adStateClosed Then
        adRs.Close
    End If
    Set adRs = Nothing
    gadoCon.Close
    Set gadoCon = Nothing
    %>

</BODY>
