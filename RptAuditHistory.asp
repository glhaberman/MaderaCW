<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: RptReportDetails.asp                                            '
'  Purpose: This is a generic page used by the other reports to display     '
'           a list of reviews for a particular total.                       '
' Includes:                                                                 '
'   IncCnn.asp          - ADO database connection                           '
'   IncTableStyles.asp  - styles for building HTML table                    '
'   IncCmnCliFunctions  - functions that are reused in multiple pages       '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim adRs
Dim mintJ, mintI, mintK
Dim intRecordset
Dim strHeading
Dim strColumn
Dim strTableName, strUserLogin, strAuditAction
Dim intTableRecordID, intTableTop
Dim dtmStartDate, dtmEndDate, mlngReReviewTypeID
Dim strPrintType, strRespWrite

strTableName = Request.QueryString("TableName")
strUserLogin = Request.QueryString("UserLogin")
intTableRecordID = Request.QueryString("RecordID")
dtmStartDate = Request.QueryString("StartDate")
dtmEndDate = Request.QueryString("EndDate")
strAuditAction = Request.QueryString("AuditAction")
mlngReReviewTypeID = Request.QueryString("ReReviewTypeID")
strPrintType = Request.QueryString("PrintType")

Set gadoCmd = GetAdoCmd("spActivityAuditList")
    AddParmIn gadoCmd, "@TableName", adVarChar, 50, IsBlank(strTableName)
    AddParmIn gadoCmd, "@TableRecordID", adInteger, 0, IsBlank(intTableRecordID)
    AddParmIn gadoCmd, "@UserLogin", adVarChar, 20, IsBlank(strUserLogin)
    AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(dtmStartDate)
    AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(dtmEndDate)
    AddParmIn gadoCmd, "@AuditAction", adVarChar, 100, IsBlank(strAuditAction)
    'Call ShowCmdParms(gadoCmd) '***DEBUG
Set adRs = GetAdoRs(gadoCmd)

If strPrintType = "Full" Then
    strHeading = ""
Else
    If Request.QueryString("TableName") = "tblReviews" Then
        strHeading = "Review ID " & Request.QueryString("RecordID")
    ElseIf Request.QueryString("TableName") = "tblReReview" Then
        If mlngReReviewTypeID = 0 Then
            strHeading = "Re-Review ID " & Request.QueryString("RecordID")
        Else
            strHeading = "CAR ID " & Request.QueryString("RecordID")
        End If
    End If
End If
%>
<HTML><HEAD>
    <TITLE>Review List</TITLE>
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
    cmdExport1.style.visibility = "hidden"
    cmdExport2.style.visibility = "hidden"
End Sub

Sub window_onafterprint()
    cmdClose1.style.visibility = "visible"
    cmdPrint1.style.visibility = "visible"
    cmdClose2.style.visibility = "visible"
    cmdPrint2.style.visibility = "visible"
    cmdExport1.style.visibility = "visible"
    cmdExport2.style.visibility = "visible"
End Sub

Sub cmdExport_onclick()
    Dim CtlRng
    'If the results div is not empty, copy it's contents to the clipboard:
    
    If tblReview.children.length > 0 Then
        'A controlRange object is used to select the results div, then copy it:
        Set CtlRng = PageBody.createControlRange()
        CtlRng.AddElement(tblReview)
        CtlRng.Select
        CtlRng.execCommand("Copy")
        Set CtlRng = Nothing
        'Clear the selection:
        document.selection.empty
        MsgBox "Results copied to clipboard.", ,"Copy Results"
    End If
End Sub
</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:white;" bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0>
    <BUTTON id=cmdPrint1 title="Send report to the printer" 
        style="LEFT:2; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdPrint_onclick"
        tabIndex=55>Print
    </BUTTON>
     <BUTTON id=cmdExport1 title="Export data from report to clipboard" 
        style="LEFT:87; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdExport_onclick"
        tabIndex=55>Copy
    </BUTTON>
    <BUTTON id=cmdClose1 title="Close window and return to report criteria screen" 
        style="LEFT:660; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdClose_onclick"
        tabIndex=55>Close
    </BUTTON>
    
    <SPAN id=lblHeading class=DefLabel
        style="FONT-SIZE:12pt; HEIGHT:20; WIDTH:720; TOP:35; LEFT:0; TEXT-ALIGN:center;position:absolute">
        <B>Audit History</B>
    </SPAN>
    
    <%
    If strPrintType = "Full" Then
        mintI = 0
        If strTableName <> "" Or strUserLogin <> "" Or intTableRecordID <> "" Or dtmStartDate <> "" Or dtmEndDate <> "" Or strAuditAction <> "" Then
            Response.Write "<SPAN id=lblHeading1 class=DefLabel"
            Response.Write "    style=""FONT-SIZE:10pt; HEIGHT:20; WIDTH:720; TOP:60; LEFT:0; TEXT-ALIGN:center;position:absolute"">"
            Response.Write "<B>Search Criteria</B></SPAN>"

            Response.Write "<TABLE id=tblCriteria Border=1 Width=720 CellSpacing=0 "
            Response.Write "Style=""position:absolute;overflow: hidden; TOP:75;width:720;left:5"">"
            Response.Write "    <TBODY>"
            strRespWrite = ""
            If strTableName <> "" Then
                strRespWrite = strRespWrite & AddCriteria("Table",Request.QueryString("TableDescr"),mintI)
                mintI = mintI + 1
            End If
            If strUserLogin <> "" Then
                strRespWrite = strRespWrite & AddCriteria("User ID",strUserLogin,mintI)
                mintI = mintI + 1
            End If
            If intTableRecordID <> "" Then
                strRespWrite = strRespWrite & AddCriteria("Record ID",intTableRecordID,mintI)
                mintI = mintI + 1
            End If
            If dtmStartDate <> "" Then
                strRespWrite = strRespWrite & AddCriteria("Start Date",dtmStartDate,mintI)
                mintI = mintI + 1
            End If
            If dtmEndDate <> "" Then
                strRespWrite = strRespWrite & AddCriteria("End Date",dtmEndDate,mintI)
                mintI = mintI + 1
            End If
            If strAuditAction <> "" Then
                strRespWrite = strRespWrite & AddCriteria("Action",strAuditAction,mintI)
                mintI = mintI + 1
            End If
            Response.Write strRespWrite
            If mintI = 3 Or mintI = 6 Then
            Else
                Response.Write "</TR>"
            End If
            Response.Write "    </TBODY>"
            Response.Write "</TABLE>"
        End If
        Select Case mintI
            Case 0
                intTableTop = 55
            Case 1,2,3
                intTableTop = 105
            Case 4,5,6
                intTableTop = 125
            Case 99
                intTableTop = 125
        End Select
    ElseIf strHeading <> "" Then
        Response.Write "<SPAN id=lblHeading1 class=DefLabel"
        Response.Write "    style=""FONT-SIZE:12pt; HEIGHT:20; WIDTH:720; TOP:55; LEFT:0; TEXT-ALIGN:center;position:absolute"">"
        Response.Write "<B>" & strHeading & "</B></SPAN>"
        intTableTop = 85
    Else
        intTableTop = 65
    End If
    %>
    <TABLE id=tblReview Border=0 Width=720 CellSpacing=0 
        Style="position:absolute;overflow: hidden; TOP:<%=intTableTop%>;width:720;left:5">
        <%
        Dim strRecords, strRecord, strValue, strCenter
        Dim aColumns(7,1)

        aColumns(0,0) = "Date Of Action"
        aColumns(0,1) = 120
        aColumns(1,0) = "User ID"
        aColumns(1,1) = 80
        If strPrintType = "Full" Then
            mintI = 5
            aColumns(2,0) = "Table"
            aColumns(2,1) = 80
            aColumns(3,0) = "Record ID"
            aColumns(3,1) = 80
            aColumns(4,0) = "Action"
            aColumns(4,1) = 100
        Else
            For mintI = 5 To 7
                aColumns(mintI,0) = ""
            Next
            mintI = 2
        End If
        aColumns(mintI,0) = "Entry Name"
        aColumns(mintI,1) = 80
        mintI = mintI + 1
        aColumns(mintI,0) = "Value Before"
        aColumns(mintI,1) = 80
        mintI = mintI + 1
        aColumns(mintI,0) = "Value After"
        aColumns(mintI,1) = 100
        
        Response.Write "<THEAD id=tbhReview style=""height:17"">"
        Response.Write "    <TR id=thrReview>"
        For mintI = 0 To 7
            If aColumns(mintI,0) <> "" Then
                Response.Write "<TD class=CellLabel id=thcReviewC" & mintI & " style=""font-size:8pt;width:" & aColumns(mintI,1) & """>" & aColumns(mintI,0) & "</TD>"
            End If
        Next
        Response.Write "    </TR>"
        Response.Write "</THEAD>"
        Response.Write "<TBODY id=tbdReview>"
        mintJ = 1
        Do While Not adRs.EOF
            strRecords = adRs.Fields("Changes").value
            
            For mintK = 1 To 100
                strRecord = Parse(strRecords,"|", mintK)
                If strRecord = "" And mintK > 1 Then
                    Exit For
                End If
                Response.Write "<TR id=tdrReview" & mintJ & "-0>" & vbCrLf
                Response.Write "<TD class=TableDetail id=tdcReviewC0" & mintJ & "-1 style=""font-size:8pt;text-align:center;width:150"">" & adRs.Fields("Date Of Action").value & "</TD>" & vbCrLf
                Response.Write "<TD class=TableDetail id=tdcReviewC1" & mintJ & "-1 style=""font-size:8pt;text-align:center"">" & adRs.Fields("User ID").value & "</TD>" & vbCrLf
                If strPrintType = "Full" Then
                    Response.Write "<TD class=TableDetail id=tdcReviewC1" & mintJ & "-1 style=""font-size:8pt;text-align:center"">" & adRs.Fields("Table").value & "</TD>" & vbCrLf
                    Response.Write "<TD class=TableDetail id=tdcReviewC1" & mintJ & "-1 style=""font-size:8pt;text-align:center"">" & adRs.Fields("Record ID").value & "</TD>" & vbCrLf
                    Response.Write "<TD class=TableDetail id=tdcReviewC1" & mintJ & "-1 style=""font-size:8pt;text-align:center"">" & adRs.Fields("Action").value & "</TD>" & vbCrLf
                End If
                For mintI = 1 To 3
                    strCenter = "Center"
                    If mintI = 1 Then strCenter = "left"
                    If Parse(strRecord, "^", mintI) = "" Then
                        strValue = "&nbsp;"
                    Else
                        strValue = Parse(strRecord, "^", mintI)
                    End If
                    Response.Write "<TD class=TableDetail id=tdcReviewC" & mintI + 1 & mintJ & "-" & mintK & " style=""font-size:8pt;text-align:" & strCenter & """>" & strValue & "</TD>" & vbCrLf
                Next
                Response.Write "</TR>" & vbCrLf
            Next
            mintJ = mintJ + 1
            adRs.MoveNext
        Loop
        
        Dim strType
        If mintJ > 10 Then
            strType = "button"
        Else
            strType = "hidden"
        End If
        'Blank row:
        Response.Write "</TBODY>"
        Response.Write "<TFOOT id=tblFooter>"
        Response.Write "    <TR id=tfrReviewF>"
        mintJ = 4
        If strPrintType = "Full" Then mintJ = 7
        For mintI = 0 To mintJ
            Response.Write "        <TD>&nbsp;</TD>"
        Next
        Response.Write "    </TR>"
        Response.Write "    <TR id=tfrReview>"
        Response.Write "        <TD><INPUT TYPE=""" & strType & """ VALUE=""Print"" onClick=""cmdPrint_onclick"" style=""width:62;height:23"" ID=cmdPrint2 NAME=""cmdPrint2""></TD>"
        Response.Write "        <TD><INPUT TYPE=""" & strType & """ VALUE=""Copy"" onClick=""cmdExport_onclick"" style=""width:62;height:23"" ID=cmdExport2 NAME=""cmdExport2""></TD>"
        mintJ = 2
        If strPrintType = "Full" Then mintJ = 5
        For mintI = 1 To mintJ
            Response.Write "        <TD><INPUT TYPE=""hidden"" VALUE="" style=""width:62;height:23"" ID=cmdFiller" & mintI & " NAME=""cmdFiller" & mintI & """></TD>"
        Next
        Response.Write "        <TD><INPUT TYPE=""" & strType & """ VALUE=""Close"" onClick=""cmdClose_onclick"" style=""width:62;height:23"" ID=cmdClose2 NAME=""cmdClose2""></TD>"
        Response.Write "    </TR>"
        Response.Write "</TFOOT>"
        %>
    </TABLE>
</BODY>
</HTML>
<%
    Function AddCriteria(strName, strValue, intColumn)
        Dim strReturn
        If intColumn = 0 Or intColumn = 3 Then
            strReturn = "<TR>"
        End If
        strReturn = strReturn & "<TD class=TableDetail style=""border-style:none;font-size:8pt;text-align:left;width:65""><B>" & strName & ":</B></TD>" & _
            "<TD class=TableDetail style=""border-style:none;font-size:8pt;text-align:left"">" & strValue & "</TD>"
        If intColumn = 2 Or intColumn = 5 Then
            strReturn = strReturn & "</TR>"
        End If
        AddCriteria = strReturn
    End Function
%>
<!--#include file="IncCmnCliFunctions.asp"-->
