<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: RptResponseDue.asp                                            '
'==========================================================================='
Const icWRK = 0
Const icSUP = 1
Const icMGR = 2
Const icOFF = 3
Const icDIR = 4
Const icFIN = 5
                
Dim intMglLevel
Dim adCmd
Dim adRs

Dim oCounts
Dim aDoTotals(6)
Dim aShowLevel(6)
Dim intI
Dim intJ
Dim dblPercent
Dim mstrPageTitle
Dim blnTempTot

Dim intHeader
Dim intleft
Dim intwidth
Dim strClass
Dim strAlign
Dim intReqLevel
Dim intTempReqLevel
Dim strTempName

Dim intShadeCount
Dim strRvwColor
Dim strMgrColor
Dim strBackColor
Dim strFontColor
DIM strPad 
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
mstrPageTitle = gstrLocationName & " " & gstrOrgName & " " & gstrAppName

Set adRs = Server.CreateObject("ADODB.Recordset") 
Set adCmd = GetAdoCmd("spRptUnsubmitted")
    AddParmIn adCmd, "@AliasID", adInteger, 0, glngAliasPosID
    AddParmIn adCmd, "@Admin", adBoolean, 0, gblnUserAdmin
    AddParmIn adCmd, "@QA", adBoolean, 0, gblnUserQA
    AddParmIn adCmd, "@UserID", adVarchar, 20, gstrUserID
    AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, ReqIsDate("StartDate")
    AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
    AddParmIn adCmd, "@Supervisor", adVarChar, 50, ReqIsBlank("Supervisor")
    AddParmIn adCmd, "@Director", adVarChar, 50, ReqIsBlank("Director")
    AddParmIn adCmd, "@Office", adVarChar, 50, ReqIsBlank("Office")
    AddParmIn adCmd, "@Manager", adVarChar, 50, ReqIsBlank("ProgramManager")
    AddParmIn adCmd, "@Submitted", adVarChar, 1, ReqIsBlank("Submitted")
    AddParmIn adCmd, "@StartReviewMonth", adDBTimeStamp, 0, ReqIsDate("StartReviewMonth")
    AddParmIn adCmd, "@EndReviewMonth", adDBTimeStamp, 0, ReqIsDate("EndReviewMonth")
    AddParmIn adCmd, "@Worker", adVarChar, 50, ReqIsBlank("Worker")
    AddParmOut adCmd, "@ReturnVal", adInteger, 0
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
%>

<HTML>
<HEAD>
    <TITLE><%=Request.Form("ReportTitle")%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncRptStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
    Call FormShow("none")
    PageBody.style.cursor = "wait"
    If Form.UserID.Value = "" Then
        MsgBox "User not recognized.  Logon failed, please try again.", vbinformation, "Log On"
        window.navigate "Logon.asp"
    End If
    Call SizeAndCenterWindow(767, 520, True)
    Call FormShow("")
    lblAppTitle.innerText = "<%=Request.Form("ReportTitle")%>"
    lblAppTitle.style.fontWeight = "bold"
    PageBody.style.cursor = "default"
    cmdPrint1.focus
End Sub

Sub cmdClose_onclick()    
    window.close
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

Sub ColClickEvent(intReviewID)
    Dim strReturnValue
    strReturnValue = window.showModalDialog("PrintReview.asp?ReviewID=" & intReviewID, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
End Sub

Sub ColMouseEvent(intDir, intRowID)
    If intDir = 0 Then
        ' Mouse over
        mstrWeight = document.all("lblCol2Row" & intRowID).style.fontweight
        document.all("lblCol2Row" & intRowID).style.fontweight = "bold"
    Else
        ' Mouse out
        document.all("lblCol2Row" & intRowID).style.fontweight = mstrWeight
    End If
End Sub
-->
</SCRIPT>
<!--#include file="IncRptExpParms.asp"-->
<!--#include file="IncCmnCliFunctions.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<!--#include file="IncRptHeader.asp"-->
    <DIV id=PageFrame
        style="HEIGHT:225; WIDTH:650; TOP:116; LEFT:10; FONT-SIZE:12pt; padding-top:5">

        <BR>                
        <%
        Call WriteCriteria()
        
        If adRs.EOF Then
            Response.Write "<SPAN id=lblNoResults class=ReportText style=""WIDTH:650; height:50; LEFT:10; TEXT-ALIGN:center"">"
            If adCmd.Parameters("@ReturnVal").Value > 0 Then
                Response.Write adCmd.Parameters("@ReturnVal").Value & " reviews matched the report criteria.<br>Please use more specific criteria to limit the results to less than 1000.<br>This report is not designed to display a large number of reviews."
            Else
                Response.Write " * No reviews matched the report criteria *"
            End If
            Response.Write "</SPAN>"
            Response.Write "<BR><BR><BR>"
        End If
        
        Set oCounts = New clsCounters
        
        strMgrColor = "#FFEFD5"
        
        'Initialize the flags for showing totals:
        For intI = 0 To icDir
            aDoTotals(intI) = False
        Next
        
        intReqLevel = icSUP
        
        intShadeCount = 0
        strRvwColor = "#ffffff"
        strMgrColor = "#FFEFD5"
        If intReqLevel = icDIR Then
            Call WriteColumnHeaders("")
        End If
        
        Do While Not adRs.EOF
            oCounts.CurrentName(icSUP) = adRs.Fields(oCounts.Field(icSUP)).Value
            oCounts.CurrentName(icMGR) = adRs.Fields(oCounts.Field(icMGR)).Value
            oCounts.CurrentName(icOFF) = adRs.Fields(oCounts.Field(icOff)).Value
            oCounts.CurrentName(icDIR) = adRs.Fields(oCounts.Field(icDIR)).Value
            
               'Place a break between Totals
            For intI = intReqLevel to icDIR
                If ocounts.EmployeeCount(intI) >= 1 AND oCounts.Changed(intI) AND intI >= intReqLevel AND oCounts.PreviousName(intReqLevel) <> "XXX" Then
                    Response.Write "<BR style=""FONT-SIZE:8"">"
                    Response.Write "<SPAN id=lblWkr style=""LEFT:10; WIDTH:640;HEIGHT:20;FONT-SIZE:14;BORDER-COLOR:#C0C0C0; BORDER-TOP-STYLE:double; BORDER-TOP-WIDTH:3""> </SPAN>"
                    Response.Write "<BR style=""FONT-SIZE:20"">"
                    Exit For
                End If
            Next
        
            'Reset values -- specialized for RptResponseDue
            intHeader = intReqLevel
                      
            For intI = intTempReqLevel To icDIR
                aShowLevel(intI) = False
            Next    
            oCounts.ShowLevel(icDIR) = False
            intTempReqLevel = intReqLevel
            If intTempReqLevel > icWRK Then
                'Find the highest valid management level
                Do While oCounts.CurrentName(intTempReqLevel) = oCounts.CurrentName(intTempReqLevel + 1)
                    intTempReqLevel = intTempReqLevel - 1
                Loop
            End If
            
            oCounts.ShowLevel(intTempReqLevel) = True
            
            'Do not print repeated names
            For intI = intTempReqLevel To icSUP
                If oCounts.CurrentName(intI) <> oCounts.CurrentName(intI + 1) Then
                    aShowLevel(intI) = True
                End If
            Next
            
            'reset intshadecount if upper management has changed
            If oCounts.Changed(intReqLevel) Then
                intShadeCount = 0 
            End If
            
            If intShadeCount MOD 2 = 0 Then
                strRvwColor = "#ffffff"
                strMgrColor = "#FFEFD5"
            Else 
                strRvwColor = "#FFEFD5"
                strMgrColor = "#ffffff"
            End If
                  
            For intI = icSUP To intReqLevel Step -1
                If oCounts.Changed(intI) AND oCounts.ShowLevel(intI) Then
                    If aShowLevel(intI) Then
                        If intI = intTempReqLevel Then
                            strClass = "ManagementText"
                            strBackColor = "#64865C"
                            strFontColor = "#FFEFD5"
                        Else
                            strClass = oCounts.HeaderClass(intI) 
                            strBackColor = "#ffffff"
                            strFontColor = "#000000"
                        End If
                        'Spacer
                        Response.Write "<SPAN id=lblWkr class=" & strClass & " "
                        Response.Write "style=""WIDTH:640; BackGround:" & strBackColor & """>"
                        Response.Write "</SPAN>"
                        
                        'Display the Director Name:    
                        Response.Write "<SPAN id=lblDir class=" & strClass & " "
                        Response.Write "style=""WIDTH:200; Color:" & strFontColor & "; BackGround:" & strBackColor & """>"
                        If intI < intReqLevel Then
                            Response.Write "*&nbsp" & oCounts.NameOnly(intI) & "</SPAN>"
                        Else
                            Response.Write oCounts.NameOnly(intI) & "</SPAN>"
                        End If
                        
                        If intI >= intTempReqLevel Then
                            Response.Write "<BR style=""FONT-SIZE:16"">"
                        End If
                    End If
                    
                    If intI = intHeader Then
                        Call WriteColumnHeaders("")
                    End If
                End If
            Next  
            Response.Write "<SPAN id=lblReviewID class=ReportText "
            Response.Write "style=""LEFT:10; WIDTH:640; background:" & strRvwColor & """></SPAN>"
 
             Response.Write "<SPAN id=lblCol1Row" & adRs.Fields("rvwID").Value & " class=ReportText "
            Response.Write "style=""WIDTH:60; LEFT:15; TEXT-ALIGN:center; background:" & strRvwColor & """>"
            Response.Write adRs.Fields("rvwID").Value & "</SPAN>"
            
            Response.Write "<SPAN id=lblCol2Row" & adRs.Fields("rvwID").Value & " class=ReportText "
            Response.Write "style=""WIDTH:80; LEFT:100; TEXT-ALIGN:center;cursor:hand;color:blue; background:" & strRvwColor & """"
            Response.Write " onmouseover=""Call ColMouseEvent(0," & adRs.Fields("rvwID").Value & ")"" onmouseout=""Call ColMouseEvent(1," & adRs.Fields("rvwID").Value & ")"" onclick=""Call ColClickEvent(" & adRs.Fields("rvwID").Value & ")"">"
            Response.Write adRs.Fields("rvwCaseNumber").Value & "</SPAN>"
           
            Response.Write "<SPAN id=lblWorker class=ReportText "
            Response.Write "style=""LEFT:190; WIDTH:150;overflow:hidden; TEXT-ALIGN:left; background:" & strRvwColor & """>"
            Response.Write Parse(adRs.Fields("Worker").Value, "--", 1) & "</SPAN>"

            Response.Write "<SPAN id=lblCaseName class=ReportText "
            Response.Write "style=""LEFT:335; WIDTH:135;overflow:hidden; TEXT-ALIGN:left; background:" & strRvwColor & """>"
            Response.Write adRs.Fields("CaseName").Value & "</SPAN>"
            
            Response.Write "<SPAN id=lblDueDate class=ReportText "
            Response.Write "style=""LEFT:480; TEXT-ALIGN: left; background:" & strRvwColor & """>"
            Response.Write adRs.Fields("rvwDateEntered").Value & "</SPAN>"
            
            Response.Write "<SPAN id=lblDaysEntered class=ReportText "
            Response.Write "style=""LEFT:555; WIDTH:90;TEXT-ALIGN:center; background:" & strRvwColor & """>"
            'If ReqForm("Submitted") = "1" Then
            '    Response.Write "NA</SPAN>"
            'ElseIf ReqForm("Submitted") = "2" Then
            '    Response.Write adRs.Fields("DaysUnsubmitted").Value & "</SPAN>"
            'Else
                Response.Write adRs.Fields("DaysUnsubmitted").Value & "</SPAN>"
            'End If
            Response.Write "<br>"
            
            adRs.MoveNext
            intShadeCount = intShadeCount + 1
        Loop
    Response.Write "<BR style=""FONT-SIZE:8"">"
    Response.Write "<SPAN id=lblWkr style=""LEFT:10; WIDTH:640;HEIGHT:20;FONT-SIZE:14;BORDER-COLOR:#C0C0C0; BORDER-TOP-STYLE:double; BORDER-TOP-WIDTH:3""> </SPAN>"
    Response.Write "<BR style=""FONT-SIZE:20"">"
    %>
<!--#include file="IncRptFooter.asp"-->

</HTML>
<%
Sub WriteColumnHeaders(strNameTitle) 
    Response.Write "<SPAN id=lblCaseIDHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:10; WIDTH:640; BORDER-BOTTOM-STYLE:none; background:" & strMgrColor & """></SPAN>"
    
    Response.Write "<SPAN id=lblCaseIDHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:10; BORDER-BOTTOM-STYLE:none; background:" & strMgrColor & """>"
    Response.Write "Review </SPAN>"
    
    Response.Write "<SPAN id=lblCaseNumberHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:100; BORDER-BOTTOM-STYLE:none; background:" & strMgrColor & """>"
    Response.Write "Case </SPAN>"
    
    Response.Write "<SPAN id=lblWorkerHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:190; WIDTH:150; BORDER-BOTTOM-STYLE:none; background:" & strMgrColor & """>"
    Response.Write "</SPAN>"

    Response.Write "<SPAN id=lblCaseNameHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:350; WIDTH:135; TEXT-ALIGN:LEFT; BORDER-BOTTOM-STYLE:none; background:" & strMgrColor & """>"
    Response.Write "Case </SPAN>"

    Response.Write "<SPAN id=lblDueDateHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:475; BORDER-BOTTOM-STYLE:none; background:" & strMgrColor & """>"
    Response.Write "Date </SPAN>"

    Response.Write "<SPAN id=lblDaysEnteredHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:555;WIDTH:90; BORDER-BOTTOM-STYLE:none; background:" & strMgrColor & """>"
    Response.Write "Days </SPAN>"

    Response.Write "<br style=""FONT-SIZE:10pt"">"
    
    Response.Write "<SPAN id=lblCaseIDHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:10; WIDTH:640; BORDER-TOP-STYLE:none;background:" & strMgrColor & """></SPAN>"
    
    Response.Write "<SPAN id=lblCaseIDHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:10; BORDER-TOP-STYLE:none;background:" & strMgrColor & """>"
    Response.Write "ID </SPAN>"

    Response.Write "<SPAN id=lblCaseNumberHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:100; BORDER-TOP-STYLE:none; background:" & strMgrColor & """>"
    Response.Write "Number </SPAN>"

    Response.Write "<SPAN id=lblWorkerHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:210; WIDTH:150; TEXT-ALIGN:left;BORDER-TOP-STYLE:none;background:" & strMgrColor & """>"
    Response.Write gstrWkrTitle & "</SPAN>"
    
    Response.Write "<SPAN id=lblCaseNameHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:350; WIDTH:135;TEXT-ALIGN:LEFT; BORDER-TOP-STYLE:none;background:" & strMgrColor & """>"
    Response.Write "Name </SPAN>"

    Response.Write "<SPAN id=lblDueDateHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:475; BORDER-TOP-STYLE:none;background:" & strMgrColor & """>"
    Response.Write "Entered </SPAN>"
    
    Response.Write "<SPAN id=lblDaysEnteredHdr1 class=ColumnHeading "
    Response.Write "style=""LEFT:555; WIDTH:90; BORDER-TOP-STYLE:none;background:" & strMgrColor & """>"
    Response.Write "Unsubmitted </SPAN>"
    
    Response.Write "<br>"
End Sub

                
Class clsCounters
    'Arrays hold counters.  Indexes 0-5:
    '   0 = worker,    1 = supervisor, 
    '   2 = manager,   3 = director   
    '   5 = final total
    Private aFormNames(6)       'Form Names
    Private aCurrentNames(6)   'Current Names
    Private aPreviousNames(6)  'Previous Names
    Private aEmployeeCount(6)  'Count of Employees for each level
    Private aNameChanged(6)    'Current Differs from Last
    Private aShowLevel(6)      'Keeps track of what levels are displayed
    Private aCFieldNames(6)    'Keeps track of the level's SQL field name
    Private aCTotalClass(6)    'Keeps track of the level's style class for totals
    Private aCHeaderClass(6)   'Keeps track of the level's style class for headings
    
    Private Sub Class_Initialize()
        Dim intI
       'Initialize property holders:
        aCFieldNames(0) = "Worker"
        aCFieldNames(1) = "Reviewer"
        aCFieldNames(2) = "Manager"
        aCFieldNames(3) = "Office"
        aCFieldNames(4) = "Director"
        aCFieldNames(5) = "Total"
        
        aCHeaderClass(0) = "WorkerHeading"
        aCHeaderClass(1) = "SupervisorHeading"
        aCHeaderClass(2) = "ManagerHeading"
        aCHeaderClass(3) = "OfficeHeading"
        aCHeaderClass(4) = "DirectorHeading"
        
        aCTotalClass(0) = "WorkerTotals"
        aCTotalClass(1) = "SupervisorTotals"
        aCTotalClass(2) = "ManagerTotals"
        aCTotalClass(3) = "OfficeTotals"
        aCTotalClass(4) = "DirectorTotals"
        
        aFormNames(0) = ReqForm("Worker")
        aFormNames(1) = ReqForm("Supervisor")
        aFormNames(2) = ReqForm("ProgramManager")
        aFormNames(3) = ReqForm("Office")
        aFormNames(4) = ReqForm("Director")
        aFormNames(5) = "X"
        For intI = icWRK To icFIN
            aCurrentNames(intI) = "XXX"
            aPreviousNames(intI) = "XXX"
            aEmployeeCount(intI) = 1
            aNameChanged(intI) = False
            aShowLevel(intI) = False
        Next
    End Sub

    Public Property Let CurrentName(intWhich, strVal)
        Dim intI
        
        'Move current value to holder for previous name:
        aPreviousNames(intWhich) = aCurrentNames(intWhich)
        If instr(1, strVal, "*") > 0 Then
            aCurrentNames(intWhich) = Parse(strVal, "*", 2)
        Else
            aCurrentNames(intWhich) = strVal
        End If
        'If the name is changing, set the changed flag:
        If aPreviousNames(intWhich) <> aCurrentNames(intWhich) Then
            aNameChanged(intWhich) = True
            'Reset counters for level of the name:
            aEmployeeCount(intWhich) = aEmployeeCount(intWhich) + 1
            'Reset Count of Employees in lower Management levels. 
            For intI = icWRK To intWhich - 1
                aEmployeeCount(intI) = 1
            Next
        Else
            aNameChanged(intWhich) = False
        End If
    End Property
    
    Public Property Get Field(intWhich)
        Field = aCFieldNames(intWhich)
    End Property
    
    Public Property Get HeaderClass(intWhich)
        HeaderClass = aCHeaderClass(intWhich)
    End Property
    
    Public Property Get TotalClass(intWhich)
        TotalClass = aCTotalClass(intWhich)
    End Property
    
    Public Property Get FormName(intWhich)
        FormName = aFormNames(intWhich)
    End Property
    
    Public Property Get Changed(intWhich)
        Changed = aNameChanged(intWhich)
    End Property

    Public Property Get CurrentName(intWhich)
        CurrentName = aCurrentNames(intWhich)
    End Property

    Public Property Get PreviousName(intWhich)
        PreviousName = aPreviousNames(intWhich)
    End Property
        
    Public Property Let ShowLevel(intWhich, blnVal)
        Dim intI
        Dim intBeg
        Dim intEnd
        Dim intStep
        
        intBeg = intWhich
        If blnVal Then
            intEnd = icFIN
            intStep = 1
        Else
            intEnd = icWRK
            intStep = -1
        End If
        
        For intI = intBeg To intEnd Step intStep
            aShowLevel(intI) = blnVal
        Next
    End Property
    
    Public Property Get ShowLevel(intWhich)
        ShowLevel = aShowLevel(intWhich)
    End Property
    
    Public Property Get EmployeeCount(intWhich)
        EmployeeCount = aEmployeeCount(intWhich)
    End Property
    
    Public Property Get NameOnly(intWhich)
        NameOnly = Parse(aCurrentNames(intWhich), "--", 1)
    End Property
End Class

%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormsReportDef.asp"-->
<!--#include file="IncReportPrintCrt.asp"-->
