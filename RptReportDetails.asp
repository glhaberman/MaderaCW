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
Dim adRs, adRsOther
Dim mintJ, intTop, mintI
Dim intRecordset
Dim strHeading1, strHeading2, strHeading3
Dim intLastColumn
Dim strSortColumn
Dim strSortOrder
Dim strFilter
Dim blnReReview
Dim mstrReportTitle, strValue
Dim mstrShowColumns, mstrHiddenIDs, intDDField, mstrUserID

mstrReportTitle = Request.QueryString("Rpt") & " - Case Review Listing"
strHeading1 = ""
strHeading2 = ""
strFilter = ""
' Default to first recordset and columns 0-4
intRecordset = 1
intLastColumn = 4
strSortColumn = Request.QueryString("SC")
If strSortColumn = "" Then
    strSortColumn = "A0"
End If
mstrHiddenIDs = ""
blnReReview = False
mstrUserID = Request.QueryString("AUI")
If Len(mstrUserID) = 0 Then
    mstrUserID = Request.QueryString("A4")
End If
Set gadoCmd = GetAdoCmd(Request.QueryString("Proc"))
Select Case Request.QueryString("Proc")
    Case "spRptReviewerCount"
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, Request.QueryString("AAL")
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("AUA")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("AQA")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("AUI")
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASD"))
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("AED"))
        AddParmIn gadoCmd, "@Director", adVarChar, 50, IsBlank(Request.QueryString("A6"))
        AddParmIn gadoCmd, "@Office", adVarChar, 50, IsBlank(Request.QueryString("A7"))
        AddParmIn gadoCmd, "@Manager", adVarChar, 50, IsBlank(Request.QueryString("A8"))
        AddParmIn gadoCmd, "@Reviewer", adVarChar, 50, IsBlank(Request.QueryString("A9"))
        AddParmIn gadoCmd, "@Worker", adVarChar, 50, IsBlank(Request.QueryString("A10"))
        AddParmIn gadoCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASR"))
        AddParmIn gadoCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("AER"))
        AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, Request.QueryString("DD")
        
        strHeading1 = Request.QueryString("SN")
        intLastColumn = 5
        'Call ShowCmdParms(gadoCmd) '***DEBUG
    Case "spReviewFind"
	    AddParmIn gadoCmd, "@AliasID", adInteger, 20, Request.QueryString("A1")
	    AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("A2")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("A3")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("A4")
        AddParmIn gadoCmd, "@casID", adInteger, 0, ZeroToNull(Request.QueryString("A5"))
        AddParmIn gadoCmd, "@casNumber", adVarChar, 20, IsBlank(Request.QueryString("A6"))
        AddParmIn gadoCmd, "@ReviewDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("A7"))
        AddParmIn gadoCmd, "@ReviewDateEnd", adDBTimeStamp, 0, IsBlank(Request.QueryString("A8"))
        AddParmIn gadoCmd, "@WorkerName", adVarChar, 100, IsBlank(Request.QueryString("A9"))
        If Request.QueryString("A20") <> "0" Then
            AddParmIn gadoCmd, "@Submitted", adVarchar, 1, Request.QueryString("A20")
        Else
            AddParmIn gadoCmd, "@Submitted", adVarchar, 1, NULL
        End If
        AddParmIn gadoCmd, "@Response", adInteger, 0, ZeroToNull(Request.QueryString("A10"))
        AddParmIn gadoCmd, "@Reviewer", adVarChar, 100, IsBlank(Request.QueryString("A11"))
        AddParmIn gadoCmd, "@PrgID", adVarchar, 255, IsBlank(Request.QueryString("A12"))
        AddParmIn gadoCmd, "@WorkerID", adVarchar, 20, IsBlank(Request.QueryString("A13"))
        AddParmIn gadoCmd, "@Supervisor", adVarchar, 100, IsBlank(Request.QueryString("A14"))
        AddParmIn gadoCmd, "@SupervisorID", adVarchar, 20, IsBlank(Request.QueryString("A15"))
        AddParmIn gadoCmd, "@ReviewClassID", adInteger, 0, ZeroToNull(Request.QueryString("A16"))
        strHeading1 = ""
        mstrReportTitle = "Find Case Review For Edit - Print List"
        mstrShowColumns = Request.QueryString("ShowCols")
        intLastColumn = 0
        For mintI = 1 To 100
            If Parse(mstrShowColumns,"^",mintI) = "" Then Exit For
            intLastColumn = mintI
        Next
        'Response.Write "mstrShowColumns=" & mstrShowColumns & "<br>"
        'response.End
        'Call ShowCmdParms(gadoCmd) '***DEBUG
    Case "spRptCaseErrSum"
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, Request.QueryString("AAL")
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("AUA")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("AQA")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("AUI")
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASD"))
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("AED"))
        AddParmIn gadoCmd, "@Director", adVarChar, 50, IsBlank(Request.QueryString("A6"))
        AddParmIn gadoCmd, "@Office", adVarChar, 50, IsBlank(Request.QueryString("A7"))
        If Request.QueryString("A8") = Request.QueryString("A7") Then
            AddParmIn gadoCmd, "@Manager", adVarChar, 50, Null
        Else
            AddParmIn gadoCmd, "@Manager", adVarChar, 50, IsBlank(Request.QueryString("A8"))
        End If
        AddParmIn gadoCmd, "@Supervisor", adVarChar, 50, IsBlank(Request.QueryString("A9"))
        AddParmIn gadoCmd, "@WorkerName", adVarchar, 50, IsBlank(Request.QueryString("A10"))
        AddParmIn gadoCmd, "@ReviewTypeID", adVarChar, 100, IsBlank(Request.QueryString("ART"))
        AddParmIn gadoCmd, "@ReviewClassID", adVarChar, 100, IsBlank(Request.QueryString("ARC"))
        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, IsBlank(Request.QueryString("APR"))
        AddParmIn gadoCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASR"))
        AddParmIn gadoCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("AER"))
        AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, Request.QueryString("DD")
        
        strHeading1 = Request.QueryString("SN")
        'Call ShowCmdParms(gadoCmd) '***DEBUG

    Case "spRptEmployeePerformance"
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, Request.QueryString("AAL")
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("AUA")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("AQA")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("AUI")
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASD"))
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("AED"))
        AddParmIn gadoCmd, "@WorkerName", adVarchar, 50, IsBlank(Request.QueryString("SN"))
        AddParmIn gadoCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASR"))
        AddParmIn gadoCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("AER"))
        If Request.QueryString("CIDAdd") <> "" Then
            AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, CInt(Request.QueryString("DD")) + CInt(Request.QueryString("CIDAdd"))
        Else
            AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, Request.QueryString("DD")
        End If
        AddParmIn gadoCmd, "@DrillDownKeyID", adInteger, 0, Request.QueryString("DD2")
        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Request.QueryString("PID")
        AddParmIn gadoCmd, "@TabID", adInteger, 0, Request.QueryString("TID")
        strHeading1 = Request.QueryString("SN")
        strHeading2 = Request.QueryString("SN2")
        intLastColumn = 5
        'Call ShowCmdParms(gadoCmd) '***DEBUG
    Case "spRptEligElemDet"
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, Request.QueryString("AAL")
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("AUA")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("AQA")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("AUI")
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASD"))
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("AED"))
        AddParmIn gadoCmd, "@Director", adVarChar, 50, IsBlank(Request.QueryString("A6"))
        AddParmIn gadoCmd, "@Office", adVarChar, 50, IsBlank(Request.QueryString("A7"))
        AddParmIn gadoCmd, "@Manager", adVarChar, 50, IsBlank(Request.QueryString("A8"))
        AddParmIn gadoCmd, "@Supervisor", adVarChar, 50, IsBlank(Request.QueryString("A9"))
        AddParmIn gadoCmd, "@WorkerName", adVarchar, 50, IsBlank(Request.QueryString("A10"))
        AddParmIn gadoCmd, "@ReviewTypeID", adVarChar, 255, IsBlank(Request.QueryString("ART"))
        AddParmIn gadoCmd, "@ReviewClassID", adVarChar, 100, IsBlank(Request.QueryString("ARC"))
        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Request.QueryString("APR")
        AddParmIn gadoCmd, "@ElementID", adInteger, 0, ZeroToNull(Request.QueryString("EID"))
        AddParmIn gadoCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASR"))
        AddParmIn gadoCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("AER"))
        AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, Request.QueryString("DD")
        strHeading1 = Request.QueryString("SN")
        strHeading2 = Request.QueryString("SN2")
        'Call ShowCmdParms(gadoCmd) '***DEBUG
    Case "spRptEligElemSum"
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, Request.QueryString("AAL")
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("AUA")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("AQA")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("AUI")
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASD"))
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("AED"))
        AddParmIn gadoCmd, "@Director", adVarChar, 50, IsBlank(Request.QueryString("A6"))
        AddParmIn gadoCmd, "@Office", adVarChar, 50, IsBlank(Request.QueryString("A7"))
        AddParmIn gadoCmd, "@Manager", adVarChar, 50, IsBlank(Request.QueryString("A8"))
        AddParmIn gadoCmd, "@Supervisor", adVarChar, 50, IsBlank(Request.QueryString("A9"))
        AddParmIn gadoCmd, "@WorkerName", adVarchar, 50, IsBlank(Request.QueryString("A10"))
        AddParmIn gadoCmd, "@ReviewTypeID", adVarChar, 255, IsBlank(Request.QueryString("ART"))
        AddParmIn gadoCmd, "@ReviewClassID", adVarChar, 100, IsBlank(Request.QueryString("ARC"))
        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Request.QueryString("APR")
        AddParmIn gadoCmd, "@ElementID", adInteger, 0, ZeroToNull(Request.QueryString("DD2"))
        AddParmIn gadoCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASR"))
        AddParmIn gadoCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("AER"))
        AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, Request.QueryString("DD")
        strHeading1 = Request.QueryString("SN")
        strHeading2 = Request.QueryString("SN2")
        Select Case Request.QueryString("DD")
            Case "1","5"
                intLastColumn = 4
            Case Else
                intLastColumn = 5
        End Select
        'Call ShowCmdParms(gadoCmd) '***DEBUG
    Case "spRptCausalFactor"
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, Request.QueryString("A1")
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("A2")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("A3")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("A4")
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("A5"))
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("A6"))
        AddParmIn gadoCmd, "@Director", adVarChar, 50, IsBlank(Request.QueryString("A7"))
        AddParmIn gadoCmd, "@Office", adVarChar, 50, IsBlank(Request.QueryString("A8"))
        AddParmIn gadoCmd, "@Manager", adVarChar, 50, IsBlank(Request.QueryString("A9"))
        AddParmIn gadoCmd, "@Supervisor", adVarChar, 50, IsBlank(Request.QueryString("A10"))
        AddParmIn gadoCmd, "@WorkerName", adVarchar, 50, IsBlank(Request.QueryString("A11"))
        AddParmIn gadoCmd, "@ReviewTypeID", adVarChar, 255, IsBlank(Request.QueryString("A12"))
        AddParmIn gadoCmd, "@ReviewClassID", adVarChar, 100, IsBlank(Request.QueryString("A13"))
        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Request.QueryString("A14")
        AddParmIn gadoCmd, "@ElementID", adInteger, 0, ZeroToNull(Request.QueryString("A15"))
        AddParmIn gadoCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("A16"))
        AddParmIn gadoCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("A17"))
        AddParmIn gadoCmd, "@FactorID", adInteger, 0, ZeroToNull(Request.QueryString("DD2"))
        AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, Request.QueryString("DD")
        strHeading1 = Request.QueryString("SN")
        strHeading2 = Request.QueryString("SN2")
        strHeading3 = Request.QueryString("SN3")
        intLastColumn = 5
        'Call ShowCmdParms(gadoCmd) '***DEBUG
    Case "spRptReReviewErrSum"
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, Request.QueryString("AAL")
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("AUA")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("AQA")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("AUI")
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASD"))
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("AED"))
        AddParmIn gadoCmd, "@Director", adVarChar, 50, IsBlank(Request.QueryString("A6"))
        AddParmIn gadoCmd, "@Office", adVarChar, 50, IsBlank(Request.QueryString("A7"))
        AddParmIn gadoCmd, "@Manager", adVarChar, 50, IsBlank(Request.QueryString("A8"))
        AddParmIn gadoCmd, "@Reviewer", adVarchar, 50, IsBlank(Request.QueryString("ARN"))
        AddParmIn gadoCmd, "@Supervisor", adVarChar, 50, IsBlank(Request.QueryString("A9"))
        AddParmIn gadoCmd, "@WorkerName", adVarchar, 50, IsBlank(Request.QueryString("A10"))
        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, IsBlank(Request.QueryString("APR"))
        AddParmIn gadoCmd, "@ReReviewTypeID", adInteger, 0, Request.QueryString("ARRT")
        AddParmIn gadoCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASR"))
        AddParmIn gadoCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("AER"))
        AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, Request.QueryString("DD")
        
        strHeading1 = Request.QueryString("SN")
        'Call ShowCmdParms(gadoCmd) '***DEBUG
        
        mstrUserID = Request.QueryString("AUI")
        blnReReview = True
    Case "spRptReReviewEligElemSum"
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, Request.QueryString("AAL")
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, Request.QueryString("AUA")
        AddParmIn gadoCmd, "@QA", adBoolean, 0, Request.QueryString("AQA")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, Request.QueryString("AUI")
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASD"))
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, IsBlank(Request.QueryString("AED"))
        AddParmIn gadoCmd, "@Director", adVarChar, 50, IsBlank(Request.QueryString("A6"))
        AddParmIn gadoCmd, "@Office", adVarChar, 50, IsBlank(Request.QueryString("A7"))
        AddParmIn gadoCmd, "@Manager", adVarChar, 50, IsBlank(Request.QueryString("A8"))
        AddParmIn gadoCmd, "@Reviewer", adVarchar, 50, IsBlank(Request.QueryString("A10"))
        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, IsBlank(Request.QueryString("APR"))
        AddParmIn gadoCmd, "@ElementID", adInteger, 0, IsBlank(Request.QueryString("DD2"))
        AddParmIn gadoCmd, "@FactorID", adInteger, 0, IsBlank(Request.QueryString("DD3"))
        AddParmIn gadoCmd, "@ReReviewTypeID", adInteger, 0, 0 'ReqForm("ReReviewTypeID")
        AddParmIn gadoCmd, "@StartReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("ASR"))
        AddParmIn gadoCmd, "@EndReviewMonth", adDBTimeStamp, 0, IsBlank(Request.QueryString("AER"))
        AddParmIn gadoCmd, "@DrillDownID", adInteger, 0, Request.QueryString("DD")
        
        strHeading1 = Parse(Request.QueryString("SN"),"^",1)
        strHeading1 = Replace(strHeading1,"[AMP]","&")
        Select Case Request.QueryString("DD")
            Case "1"
                strHeading2 = ""
            Case "2"
                strHeading2 = "Accurate Re-Reviews"
            Case "3"
                strHeading2 = "Inaccurate Re-Reviews"
        End Select
        'Call ShowCmdParms(gadoCmd) '***DEBUG
        
        mstrUserID = Request.QueryString("AUI")
        blnReReview = True

    Case Else
        'Generic Proc
End Select
strHeading1 = Replace(strHeading1,"[SPACE]"," ")
'Call ShowCmdParms(gadoCmd) '***DEBUG
If intRecordset = 1 Then
    Set adRs = GetAdoRs(gadoCmd)
ElseIf intRecordset = 2 Then
    Set adRsOther = GetAdoRs(gadoCmd)
    Set adRs = adRsOther.NextRecordset
End If
If Request.QueryString("Proc") = "spRptCausalFactor" Then
    If Request.QueryString("A19") <> "Y" Then
        'If "Include All Factors" is not checked, filter out causal factors that are all NA.
        adRs.Filter = "StatusID=22 Or StatusID=23"
    End If
End If
If strSortColumn <> "X" And adRs.RecordCount > 0 Then
    mintI = Mid(strSortColumn,2,Len(strSortColumn)-1)
    strSortOrder = ""
    If Left(strSortColumn,1) = "D" Then
        strSortOrder = " DESC"
    End If
    strSortOrder = "[" & adRs.Fields(CInt(mintI)).Name & "]" & strSortOrder
    adRs.Sort = strSortOrder
    adRs.MoveFirst
End If
Function ZeroToNullz(strField)
    If Not IsNumeric(strField) Then
        ZeroToNull = NULL
    Else
        ZeroToNull = CLng(strField)
    End If
End Function
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

Sub ReReviewRowClick(intRowID)
    Dim lngReviewID
    Dim strReturnValue
    
    lngReviewID = tblReview.rows("tdrReview" & intRowID).cells(3).innerHTML
    strReturnValue = window.showModalDialog("PrintReReview.asp?UserID=<%=mstrUserID%>&ReReviewID=" & lngReviewID, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
End Sub

Sub ReReviewMouseOver(intRowID)
    tblReview.rows("tdrReview" & intRowID).cells(3).style.fontWeight = "bold"
End Sub
Sub ReReviewMouseOut(intRowID)
    tblReview.rows("tdrReview" & intRowID).cells(3).style.fontWeight = "normal"
End Sub

Sub RowClick2(intRowID)
    Dim lngReviewID
    Dim strReturnValue

    lngReviewID = document.all("txtReviewID" & intRowID).value
    strReturnValue = window.showModalDialog("PrintReview.asp?UserID=<%=mstrUserID%>&ReviewID=" & lngReviewID, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
End Sub

Sub RowClick(intRowID)
    Dim lngReviewID
    Dim strReturnValue

    lngReviewID = tblReview.rows("tdrReview" & intRowID).cells(0).innerHTML
    strReturnValue = window.showModalDialog("PrintReview.asp?UserID=<%=mstrUserID%>&ReviewID=" & lngReviewID, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
End Sub
Sub CaseNumberMouseOver2(intRowID, intCellID)
    tblReview.rows("tdrReview" & intRowID).cells(intCellID).style.fontWeight = "bold"
End Sub
Sub CaseNumberMouseOut2(intRowID, intCellID)
    tblReview.rows("tdrReview" & intRowID).cells(intCellID).style.fontWeight = "normal"
End Sub


Sub CaseNumberMouseOver(intRowID)
    tblReview.rows("tdrReview" & intRowID).cells(1).style.fontWeight = "bold"
End Sub
Sub CaseNumberMouseOut(intRowID)
    tblReview.rows("tdrReview" & intRowID).cells(1).style.fontWeight = "normal"
End Sub

Sub SortResults(intColumn)
    Dim strOrderBy, strCurrent
    Dim blnDesc
    
    strCurrent = "<%=strSortColumn%>"
    If InStr(strCurrent,intColumn) > 0 Then
        If Left(strCurrent,1) = "A" Then
            strOrderBy = "D" & intColumn
        Else
            strOrderBy = "A" & intColumn
        End If
    Else
        strOrderBy = "A" & intColumn
    End If
    window.returnvalue = strOrderBy
    window.close
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
<BODY id=PageBody style="BACKGROUND-COLOR:white;" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
    <BUTTON id=cmdPrint1 title="Send report to the printer" 
        style="LEFT:10; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdPrint_onclick"
        tabIndex=55>Print
    </BUTTON>
     <BUTTON id=cmdExport1 title="Export data from report to clipboard" 
        style="LEFT:95; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdExport_onclick"
        tabIndex=55>Copy
    </BUTTON>
    <BUTTON id=cmdClose1 title="Close window and return to report criteria screen" 
        style="LEFT:495; WIDTH:65; TOP:5; HEIGHT:23;position:absolute" 
        onclick="cmdClose_onclick"
        tabIndex=55>Close
    </BUTTON>
    
    <SPAN id=lblHeading class=DefLabel
        style="FONT-SIZE:12pt; HEIGHT:20; WIDTH:640; TOP:35; LEFT:10; TEXT-ALIGN:center;position:absolute">
        <%=mstrReportTitle%>
    </SPAN>
    
    <%
    intTop = 55
    If strHeading1 <> "" Then
        Response.Write "<SPAN id=lblHeading1 class=DefLabel"
        Response.Write "    style=""FONT-SIZE:12pt; HEIGHT:20; WIDTH:640; TOP:55; LEFT:10; TEXT-ALIGN:center;position:absolute"">"
        Response.Write "<B>" & strHeading1 & "</B></SPAN>"
        intTop = 85
    End If
    If strHeading2 <> "" Then
        Response.Write "<SPAN id=lblHeading2 class=DefLabel"
        Response.Write "    style=""FONT-SIZE:12pt; HEIGHT:20; WIDTH:640; TOP:80; LEFT:10; TEXT-ALIGN:center;position:absolute"">"
        Response.Write "<B>" & strHeading2 & "</B></SPAN>"
        intTop = 110
    End If
    If strHeading3 <> "" Then
        Response.Write "<SPAN id=lblHeading3 class=DefLabel"
        Response.Write "    style=""FONT-SIZE:12pt; HEIGHT:20; WIDTH:675; TOP:105; LEFT:10; TEXT-ALIGN:center;position:absolute"">"
        Response.Write "<B>" & strHeading3 & "</B></SPAN>"
        intTop = 135
    End If
    %>
    <TABLE id=tblReview Border=0 Width=640 CellSpacing=0 
        Style="position:absolute;overflow: hidden; TOP:<%=intTop%>;width:640;left:10">
        <%
        Response.Write "<THEAD id=tbhReview style=""height:17"">"
        Response.Write "    <TR id=thrReview>"
        If Request.QueryString("Proc") = "spReviewFind" Then
            intLastColumn = 0
            For mintI = 0 To 100
                If intLastColumn < 11 Then
                    If InStr("^" & mstrShowColumns,"^" & mintI+1 & "^") > 0 Then
                        If strSortColumn <> "X" Then
                            Response.Write "        <TD class=CellLabel id=thcReviewC" & mintI & " style=""cursor:hand;font-size:8pt"""
                            Response.Write "            title=""Sort Results by " & adRs.Fields(mintI).Name & """ onclick=SortResults(" & mintI & ")>" & adRs.Fields(mintI).Name & "</TD>"
                        Else
                            Response.Write "        <TD class=CellLabel id=thcReviewC" & mintI & " style=""font-size:8pt"">" & adRs.Fields(mintI).Name & "</TD>"
                        End If
                        intLastColumn = intLastColumn + 1
                    End If
                End If
            Next
        Else
            For mintI = 0 To intLastColumn
                If strSortColumn <> "X" Then
                    Response.Write "        <TD class=CellLabel id=thcReviewC" & mintI & " style=""cursor:hand;font-size:8pt"""
                    Response.Write "            title=""Sort Results by " & adRs.Fields(mintI).Name & """ onclick=SortResults(" & mintI & ")>" & adRs.Fields(mintI).Name & "</TD>"
                Else
                    Response.Write "        <TD class=CellLabel id=thcReviewC" & mintI & " style=""font-size:8pt"">" & adRs.Fields(mintI).Name & "</TD>"
                End If
            Next
        End If
        Response.Write "    </TR>"
        Response.Write "</THEAD>"
        Response.Write "<TBODY id=tbdReview>"
        mintJ = 1
        If strFilter <> "" Then
            adRs.Filter = strFilter
        End If
        If Request.QueryString("Proc") = "spReviewFind" Then
            intDDField = -1
            If InStr("^" & mstrShowColumns,"^2^") > 0 Then
                intDDField = 1
            End If
            If InStr("^" & mstrShowColumns,"^1^") > 0 And intDDField = -1 Then
                intDDField = 0
            End If
            If intDDField = -1 Then intDDField = CInt(Parse(mstrShowColumns,"^",1)) - 1
        End If
        Do While Not adRs.EOF
            Response.Write "<TR id=tdrReview" & mintJ & " >" & vbCrLf
            If Request.QueryString("Proc") = "spReviewFind" Then
                intLastColumn = 0
                
                'response.Write "<BR>intDDField=" & intDDField & "<BR>"
                'response.End
                
                mstrHiddenIDs = mstrHiddenIDs & "<INPUT type=""hidden"" id=txtReviewID" & mintJ & " value=""" & adRs.Fields("Review ID").value & """>" & vbCrLf
                For mintI = 0 To 100
                    If intLastColumn < 11 Then
                        If InStr("^" & mstrShowColumns,"^" & mintI+1 & "^") > 0 Then
                            If mintI = intDDField Then
                                Response.Write "    <TD class=TableDetail id=tdcReviewC1" & mintJ & " style=""font-size:8pt;text-align:center;color:blue;cursor:hand"""
                                Response.Write " onmouseover=""Call CaseNumberMouseOver2(" & mintJ & "," & intLastColumn & ")"" onmouseout=""Call CaseNumberMouseOut2(" & mintJ & "," & intLastColumn & ")"" onclick=RowClick2(" & mintJ & ")>" & adRs.Fields(intDDField).Value & "</TD>" & vbCrLf
                            Else
                                If adRs.Fields(mintI).Type = 6 Then
                                    Response.Write "    <TD class=TableDetail id=tdcReviewC" & mintI & mintJ & " style=""font-size:8pt;text-align:center"">" & FormatCurrency(adRs.Fields(mintI).Value,0) & "</TD>" & vbCrLf
                                Else
                                    Response.Write "    <TD class=TableDetail id=tdcReviewC" & mintI & mintJ & " style=""font-size:8pt;text-align:center"">" & adRs.Fields(mintI).Value & "</TD>" & vbCrLf
                                End If
                            End If
                            intLastColumn = intLastColumn + 1
                        End If
                    End If
                Next
            Else
                For mintI = 0 To intLastColumn
                    If mintI = 1 Then
                        'If Request.QueryString("Proc") = "spRptIntegrityScreen" And Request.QueryString("DD") = 4 Then
    
                        'Else
                            Response.Write "    <TD class=TableDetail id=tdcReviewC1" & mintJ & " style=""font-size:8pt;text-align:center;color:blue;cursor:hand"" onmouseover=CaseNumberMouseOver(" & mintJ & ") onmouseout=CaseNumberMouseOut(" & mintJ & ") onclick=RowClick(" & mintJ & ")>" & adRs.Fields(1).Value & "</TD>" & vbCrLf
                        'End If
                    ElseIf mintI = 3 And blnReReview = True Then
                        Response.Write "    <TD class=TableDetail id=tdcReviewC3" & mintJ & " style=""width:110;font-size:10pt;text-align:center;color:blue;cursor:hand"" onmouseover=ReReviewMouseOver(" & mintJ & ") onmouseout=ReReviewMouseOut(" & mintJ & ") onclick=ReReviewRowClick(" & mintJ & ")>" & adRs.Fields(3).Value & "</TD>" & vbCrLf
                    Else
                        If adRs.Fields(mintI).Type = 6 Then
                            Response.Write "    <TD class=TableDetail id=tdcReviewC" & mintI & mintJ & " style=""font-size:8pt;text-align:center"">" & FormatCurrency(adRs.Fields(mintI).Value,0) & "</TD>" & vbCrLf
                        Else
                            Response.Write "    <TD class=TableDetail id=tdcReviewC" & mintI & mintJ & " style=""font-size:8pt;text-align:center"">" & adRs.Fields(mintI).Value & "</TD>" & vbCrLf
                        End If
                    End If
                Next
            End If
            Response.Write "</TR>" & vbCrLf
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
        %>
        </TBODY>
        <TFOOT id=tblFooter>
            <TR id=tfrReviewF>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
            </TR>
            
            <TR id=tfrReview>
                <TD><INPUT TYPE="<%=strType%>" VALUE="Print" onClick="cmdPrint_onclick" style="width:62;height:23" ID=cmdPrint2 NAME="cmdPrint2"></TD>
                <TD><INPUT TYPE="<%=strType%>" VALUE="Copy" onClick="cmdExport_onclick" style="width:62;height:23" ID=cmdExport2 NAME="cmdExport2"></TD>
                <TD><INPUT TYPE="hidden" VALUE="" style="width:62;height:23" ID=cmdFiller1 NAME="cmdFiller1"></TD>
                <TD><INPUT TYPE="hidden" VALUE="" style="width:62;height:23" ID=cmdFiller2 NAME="cmdFiller2"></TD>
                <TD><INPUT TYPE="<%=strType%>" VALUE="Close" onClick="cmdClose_onclick" style="width:62;height:23" ID=cmdClose2 NAME="cmdClose2"></TD>
            </TR>
        </TFOOT>
    </TABLE>
    <%Response.Write mstrHiddenIDs%>
</BODY>
</HTML>
<!--#include file="IncCmnCliFunctions.asp"-->
