<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: EligElemSum.asp                                                 '
'  Purpose: Displays the Eligibility Element Summary report, based on the   '
'           criteria passed to this page by the previous criteria screen.   '
'==========================================================================='
Dim mstrSQL
Dim mdblPercent 
Dim mstrPageTitle
Dim adRs
Dim adCmd
Dim blnDoBR
Dim intTotalCnt
Dim dblPercent
Dim intLeft
Dim intX
Dim intWidth
Dim intPos
Dim strName
Dim blnshowline
Dim intmiddle
Dim intright
Dim intShadeCount
Dim strColor
Dim strFont
Dim mintTotalElements
Dim mintTotalNA
Dim mintTotalCorrect
Dim strHoldName, strFirstMonth
Dim intColumnLeft
Dim intI
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName
'Retrieve the records that match the report criteria:
Set adCmd = GetAdoCmd("spRptEligElemOVTrend")
    AddParmIn adCmd, "@AliasID", adInteger, 0, glngAliasPosID
    AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
    AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
    AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
    AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, ReqIsDate("StartDate")
    AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
    AddParmIn adCmd, "@Director", adVarchar, 50, ReqIsBlank("Director")
    AddParmIn adCmd, "@Office", adVarchar, 50, ReqIsBlank("Office")
    AddParmIn adCmd, "@Manager", adVarchar, 50, ReqIsBlank("ProgramManager")
    AddParmIn adCmd, "@Supervisor", adVarchar, 50, ReqIsBlank("Supervisor")
    AddParmIn adCmd, "@WorkerName", adVarchar, 50, ReqIsBlank("Worker")
    AddParmIn adCmd, "@ReviewTypeID", adVarChar, 100, ReqIsBlank("ReviewTypeID")
    AddParmIn adCmd, "@ReviewClassID", adVarChar, 100, ReqIsBlank("ReviewClassID")
    AddParmIn adCmd, "@ProgramID", adInteger, 0, ReqZeroToNull("ProgramID")
    AddParmIn adCmd, "@ElementID", adInteger, 0, ReqZeroToNull("EligElementID")
    AddParmIn adCmd, "@StartReviewMonth", adDBTimeStamp, 0, ReqIsDate("StartReviewMonth")
    AddParmIn adCmd, "@EndReviewMonth", adDBTimeStamp, 0, ReqIsDate("EndReviewMonth")
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
%>

<HTML>
<HEAD>
    <TITLE><%=ReqForm("ReportTitle")%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncRptStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Option Explicit
Sub cmdClose_onclick()	
	Window.close
End Sub

Sub FormShow(strVis)
	cmdPrint1.style.display = strVis
    cmdClose1.style.display = strVis
    cmdPrint2.style.display = strVis
    cmdClose2.style.display = strVis
    cmdExport1.style.display = strVis
    cmdExport2.style.display = strVis
    Header.style.display = strVis
    PageFrame.style.display = strVis
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub window_onbeforeprint()
    cmdPrint1.style.visibility = "hidden"
    cmdClose1.style.visibility = "hidden"
    cmdPrint2.style.visibility = "hidden"
    cmdClose2.style.visibility = "hidden"
    cmdExport1.style.visibility = "hidden"
    cmdExport2.style.visibility = "hidden"
End Sub

Sub window_onafterprint()
    cmdPrint1.style.visibility = "visible"
    cmdClose1.style.visibility = "visible"
    cmdPrint2.style.visibility = "visible"
    cmdClose2.style.visibility = "visible"
    cmdExport1.style.visibility = "visible"
    cmdExport2.style.visibility = "visible"
End Sub
-->
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->
<!--#include file="IncRptExpParms.asp"-->
<!--#include file="IncSvrFunctions.asp"-->

<!--=== Start of Report Definition and Layout ============================= -->
<!--#include file="IncRptHeader.asp"-->
            
<DIV id=PageFrame class=RptPageFrame>
    <br>
    <%
    Dim intColumns
    Dim intElements
    Dim intRows
    Dim intColumnID
    Dim blnDone
    Dim intColWidth
    
    Response.Write "<BR><BR><BR><BR><BR><BR><BR>"
	Call WriteCriteria()
    
    blnDone = False
    If adRs.EOF Then
        Response.Write "<BR><BR>"
        Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; LEFT:0; TEXT-ALIGN:center"">"
        Response.Write " * No reviews matched the report criteria *"
        blnDone = True
    Else
		'Response.Write "<BR>"
		'Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
		'Response.Write "style=""WIDTH:640; LEFT:10; HEIGHT:40;BORDER-BOTTOM-STYLE:none"">"
		'Response.Write "</SPAN>"

        strHoldName = ""
        intLeft = 200
        strFirstMonth = ""
        If adRs.RecordCount > 0 Then
            intColumns = adRs.Fields("ColumnsNeeded").value
            intElements = adRs.Fields("ElementCount").value
        Else
            intColumns = 1
            intElements = 1
        End If
        If CInt(intColumns) > 6 Then
            intRows = 2
            If intColumns MOD 2 = 0 Then
                intColumns = CInt(intColumns) / 2
            Else
                intColumns = Int(CInt(intColumns) / 2) + 1
            End If
        Else
            intRows = 1
        End If
        intColWidth = 480/intColumns
        intColumnID = 1
        adRs.Filter = "ColumnID<" & intColumns
        Do While Not adRs.Eof
            If adRs.Fields("MonthName").value = strFirstMonth Then
                Exit Do
            End If
            If strHoldName <> adRs.Fields("MonthName").value Then
			    Response.Write "<SPAN id=lblTotalCasesHdr class=ColumnHeading "
			    Response.Write "style=""WIDTH:50; LEFT:" & intLeft & ";HEIGHT:40;BORDER-STYLE:none;overflow:auto"">"
			    Response.Write adRs.Fields("MonthName").value & "</SPAN>"
			    strHoldName = adRs.Fields("MonthName").value
			    If strFirstMonth = "" Then strFirstMonth = strHoldName
			    intLeft = intLeft + intColWidth
            End If
            adRs.MoveNext
        Loop
        Response.Write "<BR><SPAN id=lblTotalCasesHdr class=ColumnHeading "
	    Response.Write "style=""WIDTH:190; LEFT:10;HEIGHT:40;BORDER-STYLE:none;overflow:auto;text-align:left;text-valign:bottom"">"
	    Response.Write "Element</SPAN>"
	    Response.Write "<BR><HR style=""width:640"">"
		'Response.Write "<BR><BR style=""font-size:6"">"
		adRs.MoveFirst
    End If

    mintTotalElements = 0
    mintTotalNA = 0
    mintTotalCorrect = 0
    strHoldName = ""
    intLeft = 180
    Do While Not adRs.EOF
        intTotalCnt = adRs.Fields("Prg1TotalCases").Value - adRs.Fields("Prg1NotAppCnt").Value
		If strHoldName <> adRs.Fields("elmLongTitle").Value Then
			If strHoldName <> "" Then
                Response.Write "<BR>"
			End If
			Response.Write "<SPAN id=lblElement class=ManagementText "
            Response.Write "style=""WIDTH:640; LEFT:10;background:" & strColor & """></SPAN>"
            
            Response.Write "<SPAN id=lblElement class=ManagementText "
            Response.Write "style=""WIDTH:175; LEFT:10; OverFlow:hidden; Color:" & strFont & ";background:" & strColor & """>"
            Response.Write adRs.Fields("elmLongTitle").Value & "</SPAN>"
            strHoldName = adRs.Fields("elmLongTitle").Value
        End If
                    
        If intTotalCnt > 0 Then
            dblPercent = FormatNumber((adRs.Fields("Prg1CorrectCnt").Value / intTotalCnt) * 100,1)
        Else
            dblPercent = "0.0"
        End If
        
        intColumnLeft = CInt(intLeft) + (CInt(adRs.Fields("ColumnID").Value)*intColWidth)
        Response.Write "<SPAN id=lblCorrectPercentHdr class=ManagementText "
        Response.Write "style=""LEFT:" & intColumnLeft & "; width:70; text-align:right; background:" & strColor & """>"
        Response.Write dblPercent & " %</SPAN>"
        
        adRs.MoveNext
        intShadeCount = intShadeCount + 1
    Loop
    Response.Write "<BR><BR>"

'--------------- secon set of dates, if needed
	Response.Write "<BR>"
	Response.Write "<SPAN id=lblElementHdr class=ColumnHeading "
	Response.Write "style=""WIDTH:640; LEFT:10; HEIGHT:40;BORDER-BOTTOM-STYLE:none"">"
	Response.Write "</SPAN>"
    If Not blnDone Then
        strHoldName = ""
        intLeft = 200
        strFirstMonth = ""
        adRs.MoveFirst
        adRs.Filter = "ColumnID>=" & intColumns
        If adRs.RecordCount  > 0 Then
            Do While Not adRs.Eof
                If adRs.Fields("MonthName").value = strFirstMonth Then
                    Exit Do
                End If
                If strHoldName <> adRs.Fields("MonthName").value Then
			        Response.Write "<SPAN id=lblTotalCasesHdr class=ColumnHeading "
			        Response.Write "style=""WIDTH:50; LEFT:" & intLeft & ";HEIGHT:40;BORDER-BOTTOM-STYLE:none;overflow:auto"">"
			        Response.Write adRs.Fields("MonthName").value & "</SPAN>"
			        strHoldName = adRs.Fields("MonthName").value
			        If strFirstMonth = "" Then strFirstMonth = strHoldName
			        intLeft = intLeft + intColWidth
                End If
                adRs.MoveNext
            Loop
		    Response.Write "<BR><BR><BR>"
		    adRs.MoveFirst
            mintTotalElements = 0
            mintTotalNA = 0
            mintTotalCorrect = 0
            strHoldName = ""
            intLeft = 180
            Do While Not adRs.EOF
                intTotalCnt = adRs.Fields("Prg1TotalCases").Value - adRs.Fields("Prg1NotAppCnt").Value
			    If strHoldName <> adRs.Fields("elmLongTitle").Value Then
			        If strHoldName <> "" Then
                        Response.Write "<BR>"
			        End If
			        Response.Write "<SPAN id=lblElement class=ManagementText "
                    Response.Write "style=""WIDTH:640; LEFT:10;background:" & strColor & """></SPAN>"
                    
                    Response.Write "<SPAN id=lblElement class=ManagementText "
                    Response.Write "style=""WIDTH:175; LEFT:10; OverFlow:hidden; Color:" & strFont & ";background:" & strColor & """>"
                    Response.Write adRs.Fields("elmLongTitle").Value & "</SPAN>"
                    strHoldName = adRs.Fields("elmLongTitle").Value
                End If
                            
                If intTotalCnt > 0 Then
                    dblPercent = FormatNumber((adRs.Fields("Prg1CorrectCnt").Value / intTotalCnt) * 100,1)
                Else
                    dblPercent = "0.0"
                End If
                
                intColumnLeft = CInt(intLeft) + (CInt(adRs.Fields("ColumnID").Value - intColumns)*intColWidth)
                Response.Write "<SPAN id=lblCorrectPercentHdr class=ManagementText "
                Response.Write "style=""LEFT:" & intColumnLeft & "; width:70; text-align:right; background:" & strColor & """>"
                Response.Write dblPercent & " %</SPAN>"
                
                adRs.MoveNext
                intShadeCount = intShadeCount + 1
            Loop
            Response.Write "<BR><BR>"
        End If
    End If
%>
<!--#include file="IncRptFooter.asp"-->

</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->