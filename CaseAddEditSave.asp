<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: CaseAddEditSave.asp                                             '
'  Purpose: This page is used to save tblReview records                     '
'==========================================================================='
Dim madoRs              'Generic recordset reused for various tasks.
Dim mstrPageTitle       'Title used at the top of the page.
Dim mstrAction          'The action from the post-back (add, update, etc).
Dim mlngCurrentRvwID    'Holds the record ID number of the current review.
Dim madoCmdRvw          'ADO command object for updating And getting review.
Dim madoRsRvw           'ADO recordset for updating And getting a review.
Dim mintI               'Generic loop counter.
Dim mintJ               'Generic loop counter.
Dim mblnDeleteFail		'Set True if user tries to delete a Rereviewed review.
Dim mblnChangesSaved    'Set to True if page is being loaded after saving changes.
Dim mstrChangesSaved    'String to display message after saving changes.
Dim mstrComments
Dim mintDelCnt
Dim mstrDelRecord
Dim mdctElements
Dim strRecord, strElmRecord, strFacList, strFacRecord
%>
<!--#include file="IncCnn.asp"-->
<%
'==============================================================================
' Server side action:
'==============================================================================
'Instantiate the recordset that is reused for temporary results or queries:
Set madoRs = Server.CreateObject("ADODB.Recordset")
Set mdctElements = CreateObject("Scripting.Dictionary")

'Set the page title:
mstrPageTitle = Trim(gstrTitle & " " & gstrAppName)

'--------------------------------------------------------------------
' The FormAction value posted to the Request object will control
' how the page responds.  Possible values for Form Action are:
'    [blank]            Arriving here from main menu.
'    AddRecord          Postback after user clicked Add.
'    ChangeRecord       Postback after user edited existing review.
'    DeleteRecord       Postback after user clicked delete.
'    GetRecord          Arriving here after searching for a review.
'    AddRecordPrint     Postback after clicking print while adding.
'    ChangeRecordPrint  Postback after clicking print while editing.
'--------------------------------------------------------------------
mstrAction = ReqForm("FormAction")

'The form indicates that the user clicked Print will adding Or editing
'a review by appending "Print" to the end of the current action.  The
'record is saved, And when posting back the form will finish printing
'in the window_onload.  The "Print" keyword is removed here to restore
'the indicator to the previous action.
If Instr(mstrAction, "Print") Then
    'Remove the Print text from the form action
    mstrAction = Left(mstrAction, Len(mstrAction) - 5)
End If
mblnChangesSaved = False
mblnDeleteFail = False
Select Case mstrAction
    Case "AddRecord"
        'Insert the new review record:
        Set gadoCmd = GetAdoCmd("spReviewAdd")
            AddParmIn gadoCmd, "@rvwDateEntered", adDBTimeStamp, 0, ReqIsDate("rvwDateEntered")
            AddParmIn gadoCmd, "@rvwMonthYear", adVarChar, 7, ReqForm("rvwMonthYear")
            AddParmIn gadoCmd, "@rvwReviewerName", adVarChar, 100, ReqForm("rvwReviewerName")
            AddParmIn gadoCmd, "@rvwReviewClassID", adInteger, 0, ReqForm("rvwReviewClassID")
            AddParmIn gadoCmd, "@rvwWorkerID", adVarChar, 20, ReqForm("rvwWorkerID")
            AddParmIn gadoCmd, "@rvwWorkerEmpID", adVarChar, 20, ReqForm("rvwWorkerEmpID")
            AddParmIn gadoCmd, "@rvwWorkerName", adVarChar, 100, ReqForm("rvwWorkerName")
            AddParmIn gadoCmd, "@rvwSupervisorEmpID", adVarChar, 20, ReqForm("rvwSupervisorEmpID")
            AddParmIn gadoCmd, "@rvwSupervisorName", adVarChar, 100, ReqForm("rvwSupervisorName")
            AddParmIn gadoCmd, "@rvwManagerName", adVarChar, 100, ReqForm("rvwManagerName")
            AddParmIn gadoCmd, "@rvwOfficeName", adVarChar, 100, ReqForm("rvwOfficeName")
            AddParmIn gadoCmd, "@rvwCaseLastName", adVarChar, 50, ReqForm("rvwCaseLastName")
            AddParmIn gadoCmd, "@rvwCaseFirstName", adVarChar, 50, ReqForm("rvwCaseFirstName")
            AddParmIn gadoCmd, "@rvwCaseNumber", adVarChar, 25, ReqForm("rvwCaseNumber")
            AddParmIn gadoCmd, "@rvwResponseDueDate", adDBTimeStamp, 0, ReqIsDate("rvwResponseDueDate")
            AddParmIn gadoCmd, "@rvwWorkerResponseID", adInteger, 0, ReqForm("rvwWorkerResponseID")
            AddParmIn gadoCmd, "@rvwSubmitted", adVarChar, 1, ReqForm("rvwSubmitted")
            AddParmIn gadoCmd, "@rvwStatusID", adInteger, 0, ReqIsBlank("rvwStatusID")
            AddParmIn gadoCmd, "@UserID", adVarChar, 20, ReqForm("UserID")
            AddParmIn gadoCmd, "@rvwSupSig", adChar, 1, ReqForm("rvwSupSig")
            AddParmIn gadoCmd, "@rvwSupComments", adVarChar, 5000, ReqForm("rvwSupComments")
            AddParmIn gadoCmd, "@rvwWrkSig", adChar, 1, ReqForm("rvwWrkSig")
            AddParmIn gadoCmd, "@rvwWrkComments", adVarChar, 5000, ReqForm("rvwWrkComments")
            AddParmIn gadoCmd, "@rvwWorkerSigResponseID", adInteger, 0 , ReqForm("rvwWorkerSigResponseID")
            AddParmOut gadoCmd, "@NxtID", adInteger, 0
            'ShowCmdParms(gadoCmd) '***DEBUG
            On Error Resume Next
            gadoCmd.Execute
            If Err.number <> 0 Then
                WriteDbError "Add New Review - tblReview", Err.Source, Err.number, Err.Description, gadoCmd
            End If
            On Error Goto 0
            mlngCurrentRvwID = gadoCmd.Parameters("@NxtID").Value
        Set gadoCmd = Nothing
        If mlngCurrentRvwID <> -1 Then
            If ReqForm("ReviewElementData") <>  "" Then
                For mintI = 1 To 1000
                    strRecord = Parse(ReqForm("ReviewElementData"),"|",mintI)
                    If strRecord = "" Then Exit For
                    strElmRecord = Parse(strRecord,"^",4)
                    strFacList = Parse(strElmRecord,"*",4)
                    Set gadoCmd = GetAdoCmd("spReviewElementAdd")
                        AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
                        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Parse(strRecord,"^",1)
                        AddParmIn gadoCmd, "@TypeID", adInteger, 0, Parse(strRecord,"^",2)
                        AddParmIn gadoCmd, "@ElementID", adInteger, 0, Parse(strRecord,"^",3)
                        AddParmIn gadoCmd, "@StatusID", adInteger, 0, Parse(strElmRecord,"*",1)
                        AddParmIn gadoCmd, "@TimeFrameID", adInteger, 0, Parse(strElmRecord,"*",2)
                        AddParmIn gadoCmd, "@Comments", adVarchar, 5000, IsBlank(Parse(strElmRecord,"*",3))
                        'ShowCmdParms(gadoCmd) '***DEBUG
                        On Error Resume Next
                        gadoCmd.Execute
                        If Err.number <> 0 Then
                            WriteDbError "Add New Review - tblReviewsElements", Err.Source, Err.number, Err.Description, gadoCmd
                        End If
                        On Error Goto 0
                    For mintJ = 1 To 50
                        strFacRecord = Parse(strFacList,"!",mintJ)
                        If strFacRecord = "" Then Exit For
                        If Parse(strFacRecord,"~",1) <> "0" Then
                            Set gadoCmd = GetAdoCmd("spReviewFactorAdd")
                                AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
                                AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Parse(strRecord,"^",1)
                                AddParmIn gadoCmd, "@TypeID", adInteger, 0, Parse(strRecord,"^",2)
                                AddParmIn gadoCmd, "@ElementID", adInteger, 0, Parse(strRecord,"^",3)
                                AddParmIn gadoCmd, "@FactorID", adInteger, 0, Parse(strFacRecord,"~",1)
                                AddParmIn gadoCmd, "@StatusID", adInteger, 0, Parse(strFacRecord,"~",2)
                                'ShowCmdParms(gadoCmd) '***DEBUG
                                On Error Resume Next
                                gadoCmd.Execute
                                If Err.number <> 0 Then
                                    WriteDbError "Add New Review - tblReviewsFactors", Err.Source, Err.number, Err.Description, gadoCmd
                                End If
                                On Error Goto 0
                        End If
                    Next
                Next
            End If
            If ReqForm("ReviewCommentData") <>  "" Then
                For mintI = 1 To 1000
                    strRecord = Parse(ReqForm("ReviewCommentData"),"|",mintI)
                    If strRecord = "" Then Exit For
                    Set gadoCmd = GetAdoCmd("spReviewsCommentAdd")
                        AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
                        AddParmIn gadoCmd, "@ScreenName", adVarchar, 255, Parse(strRecord,"^",1)
                        AddParmIn gadoCmd, "@Comments", adVarchar, 5000, IsBlank(Parse(strRecord,"^",2))
                        'ShowCmdParms(gadoCmd) '***DEBUG
                        On Error Resume Next
                        gadoCmd.Execute
                        If Err.number <> 0 Then
                            WriteDbError "Add New Review - tblReviewComments", Err.Source, Err.number, Err.Description, gadoCmd
                        End If
                        On Error Goto 0
                Next
            End If
            If ReqForm("ReviewProgramData") <>  "" Then
                For mintI = 1 To 10
                    strRecord = Parse(ReqForm("ReviewProgramData"),"|",mintI)
                    If strRecord = "" Then Exit For
                    Set gadoCmd = GetAdoCmd("spReviewProgramAdd")
                        AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
                        AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Parse(strRecord,"^",1)
                        AddParmIn gadoCmd, "@ReviewTypeID", adInteger, 0, Parse(strRecord,"^",2)
                        'ShowCmdParms(gadoCmd) '***DEBUG
                        On Error Resume Next
                        gadoCmd.Execute
                        If Err.number <> 0 Then
                            WriteDbError "Add New Review - tblReviewsPrograms", Err.Source, Err.number, Err.Description, gadoCmd
                        End If
                        On Error Goto 0
                Next
            End If
        End If
        mblnChangesSaved = True
    Case "ChangeRecord"
        mlngCurrentRvwID = ReqForm("rvwID")
        'Only call a review update If data has changed:
        If Instr(ReqForm("Changed"), "[Case]") > 0 Then
            'Update an existing case review:
            Set gadoCmd = GetAdoCmd("spReviewUpd")
                AddParmIn gadoCmd, "@rvwID", adInteger, 0, mlngCurrentRvwID
                AddParmIn gadoCmd, "@rvwDateEntered", adDBTimeStamp, 0, ReqIsDate("rvwDateEntered")
                AddParmIn gadoCmd, "@rvwMonthYear", adVarChar, 7, ReqForm("rvwMonthYear")
                AddParmIn gadoCmd, "@rvwReviewerName", adVarChar, 100, ReqForm("rvwReviewerName")
                AddParmIn gadoCmd, "@rvwReviewClassID", adInteger, 0, ReqForm("rvwReviewClassID")
                AddParmIn gadoCmd, "@rvwWorkerID", adVarChar, 20, ReqForm("rvwWorkerID")
                AddParmIn gadoCmd, "@rvwWorkerEmpID", adVarChar, 20, ReqForm("rvwWorkerEmpID")
                AddParmIn gadoCmd, "@rvwWorkerName", adVarChar, 100, ReqForm("rvwWorkerName")
                AddParmIn gadoCmd, "@rvwSupervisorEmpID", adVarChar, 20, ReqForm("rvwSupervisorEmpID")
                AddParmIn gadoCmd, "@rvwSupervisorName", adVarChar, 100, ReqForm("rvwSupervisorName")
                AddParmIn gadoCmd, "@rvwManagerName", adVarChar, 100, ReqForm("rvwManagerName")
                AddParmIn gadoCmd, "@rvwOfficeName", adVarChar, 100, ReqForm("rvwOfficeName")
                AddParmIn gadoCmd, "@rvwCaseLastName", adVarChar, 50, ReqForm("rvwCaseLastName")
                AddParmIn gadoCmd, "@rvwCaseFirstName", adVarChar, 50, ReqForm("rvwCaseFirstName")
                AddParmIn gadoCmd, "@rvwCaseNumber", adVarChar, 25, ReqForm("rvwCaseNumber")
                AddParmIn gadoCmd, "@rvwResponseDueDate", adDBTimeStamp, 0, ReqIsDate("rvwResponseDueDate")
                AddParmIn gadoCmd, "@rvwWorkerResponseID", adInteger, 0, ReqIsNumeric("rvwWorkerResponseID")
                AddParmIn gadoCmd, "@rvwSubmitted", adVarChar, 1, ReqForm("rvwSubmitted")
                AddParmIn gadoCmd, "@rvwStatusID", adInteger, 0, ReqIsBlank("rvwStatusID")
                AddParmIn gadoCmd, "@UserID", adVarChar, 20, ReqForm("UserID")
                AddParmIn gadoCmd, "@rvwSupSig", adChar, 1, ReqForm("rvwSupSig")
                AddParmIn gadoCmd, "@rvwSupComments", adVarChar, 5000, ReqForm("rvwSupComments")
                AddParmIn gadoCmd, "@rvwWrkSig", adChar, 1, ReqForm("rvwWrkSig")
                AddParmIn gadoCmd, "@rvwWrkComments", adVarChar, 5000, ReqForm("rvwWrkComments")
                AddParmIn gadoCmd, "@rvwWorkerSigResponseID", adInteger, 0 , ReqForm("rvwWorkerSigResponseID")
                AddParmIn gadoCmd, "@UpdateString", adVarChar, 5000, ReqForm("UpdateString")
                AddParmIn gadoCmd, "@SupSubWorkerDisagree", adChar, 1, ReqForm("SupSubWorkerDisagree")
                AddParmOut gadoCmd, "@NxtID", adInteger, 0
                'ShowCmdParms(gadoCmd) '***DEBUG
                On Error Resume Next
                gadoCmd.Execute
                If Err.number <> 0 Then
                    WriteDbError "Edit Review - tblReviews", Err.Source, Err.number, Err.Description, gadoCmd
                End If
                On Error Goto 0
            Set gadoCmd = Nothing
        End If 'Instr(ReqForm("Changed"), "[Case]") > 0 

        If mlngCurrentRvwID <> -1 Then
            If ReqForm("ReviewElementData") <>  "" Then
                Set gadoCmd = GetAdoCmd("spReviewElementDel")
                    AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID
                    AddParmIn gadoCmd, "@DeleteCode", adVarchar, 20, ReqForm("DeleteCode")
                    gadoCmd.Execute
                If InStr(ReqForm("DeleteCode"),"[E]") > 0 Then
                    For mintI = 1 To 1000
                        strRecord = Parse(ReqForm("ReviewElementData"),"|",mintI)
                        If strRecord = "" Then Exit For
                        strElmRecord = Parse(strRecord,"^",4)
                        strFacList = Parse(strElmRecord,"*",4)
                        Set gadoCmd = GetAdoCmd("spReviewElementAdd")
                            AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
                            AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Parse(strRecord,"^",1)
                            AddParmIn gadoCmd, "@TypeID", adInteger, 0, Parse(strRecord,"^",2)
                            AddParmIn gadoCmd, "@ElementID", adInteger, 0, Parse(strRecord,"^",3)
                            AddParmIn gadoCmd, "@StatusID", adInteger, 0, Parse(strElmRecord,"*",1)
                            AddParmIn gadoCmd, "@TimeFrameID", adInteger, 0, Parse(strElmRecord,"*",2)
                            AddParmIn gadoCmd, "@Comments", adVarchar, 5000, IsBlank(Parse(strElmRecord,"*",3))
                            'ShowCmdParms(gadoCmd) '***DEBUG
                            On Error Resume Next
                            gadoCmd.Execute
                            If Err.number <> 0 Then
                                WriteDbError "Add New Review - tblReviewsElements", Err.Source, Err.number, Err.Description, gadoCmd
                            End If
                            On Error Goto 0
                        For mintJ = 1 To 50
                            strFacRecord = Parse(strFacList,"!",mintJ)
                            If strFacRecord = "" Then Exit For
                            If Parse(strFacRecord,"~",1) <> "0" Then
                                Set gadoCmd = GetAdoCmd("spReviewFactorAdd")
                                    AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
                                    AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Parse(strRecord,"^",1)
                                    AddParmIn gadoCmd, "@TypeID", adInteger, 0, Parse(strRecord,"^",2)
                                    AddParmIn gadoCmd, "@ElementID", adInteger, 0, Parse(strRecord,"^",3)
                                    AddParmIn gadoCmd, "@FactorID", adInteger, 0, Parse(strFacRecord,"~",1)
                                    AddParmIn gadoCmd, "@StatusID", adInteger, 0, Parse(strFacRecord,"~",2)
                                    'ShowCmdParms(gadoCmd) '***DEBUG
                                    On Error Resume Next
                                    gadoCmd.Execute
                                    If Err.number <> 0 Then
                                        WriteDbError "Add New Review - tblReviewsFactors", Err.Source, Err.number, Err.Description, gadoCmd
                                    End If
                                    On Error Goto 0
                            End If
                        Next
                    Next
                End If
            End If
            If InStr(ReqForm("DeleteCode"),"[C]") > 0 Then
                If ReqForm("ReviewCommentData") <>  "" Then
                    For mintI = 1 To 1000
                        strRecord = Parse(ReqForm("ReviewCommentData"),"|",mintI)
                        If strRecord = "" Then Exit For
                        Set gadoCmd = GetAdoCmd("spReviewsCommentAdd")
                            AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
                            AddParmIn gadoCmd, "@ScreenName", adVarchar, 255, Parse(strRecord,"^",1)
                            AddParmIn gadoCmd, "@Comments", adVarchar, 5000, IsBlank(Parse(strRecord,"^",2))
                            'ShowCmdParms(gadoCmd) '***DEBUG
                            On Error Resume Next
                            gadoCmd.Execute
                            If Err.number <> 0 Then
                                WriteDbError "Add New Review - tblReviewDIComments", Err.Source, Err.number, Err.Description, gadoCmd
                            End If
                            On Error Goto 0
                    Next
                End If
            End If
            If InStr(ReqForm("DeleteCode"),"[P]") > 0 Then
                If ReqForm("ReviewProgramData") <>  "" Then
                    For mintI = 1 To 10
                        strRecord = Parse(ReqForm("ReviewProgramData"),"|",mintI)
                        If strRecord = "" Then Exit For
                        Set gadoCmd = GetAdoCmd("spReviewProgramAdd")
                            AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
                            AddParmIn gadoCmd, "@ProgramID", adInteger, 0, Parse(strRecord,"^",1)
                            AddParmIn gadoCmd, "@ReviewTypeID", adInteger, 0, Parse(strRecord,"^",2)
                            'ShowCmdParms(gadoCmd) '***DEBUG
                            On Error Resume Next
                            gadoCmd.Execute
                            If Err.number <> 0 Then
                                WriteDbError "Add New Review - tblReviewsPrograms", Err.Source, Err.number, Err.Description, gadoCmd
                            End If
                            On Error Goto 0
                    Next
                End If
            End If
        End If
        mblnChangesSaved = True
    Case "GetRecord"
        mlngCurrentRvwID = ReqForm("rvwID")
        
    Case "DeleteRecord"
        'Delete an existing case review:
        mlngCurrentRvwID = ReqForm("rvwID")
        Set gadoCmd = GetAdoCmd("spReviewDel")
            AddParmIn gadoCmd, "@ID", adInteger, 0, mlngCurrentRvwID
            AddParmIn gadoCmd, "@UserID", adVarChar, 20, ReqForm("UserID")
            AddParmOut gadoCmd, "@NxtID", adInteger, 0
            gadoCmd.Execute
            mlngCurrentRvwID = gadoCmd.Parameters("@NxtID").Value
        Set gadoCmd = Nothing
        If mlngCurrentRvwID <> 0 Then
			mblnDeletefail = True
		End If
    Case Else
        'First time load of the page.
End Select
' If Page is being loaded after changes were saved, display message in window title
mstrChangesSaved = Trim(gstrOrgAbbr) & " " & gstrAppName
If mblnChangesSaved = True Then mstrChangesSaved = mstrChangesSaved & " [Changes Saved at " & Now() & "]"

'Determine If there is no current record, i.e. we arrived 
'here from the main menu:
If Not IsNumeric(mlngCurrentRvwID) Then
    mlngCurrentRvwID = -1
ElseIf mlngCurrentRvwID = 0 Then
    mlngCurrentRvwID = -1
End If
If mlngCurrentRvwID <> -1 Then
    'Retrieve the case to display:
    Set madoCmdRvw = GetAdoCmd("spReviewGet")
        AddParmIn madoCmdRvw, "@AliasID", adInteger, 0, ReqForm("AliasID")
        AddParmIn madoCmdRvw, "@Admin", adBoolean, 0, ReqForm("UserAdmin")
        AddParmIn madoCmdRvw, "@QA", adBoolean, 0, ReqForm("UserQA")
        AddParmIn madoCmdRvw, "@UserID", adVarChar, 20, ReqForm("UserID")
        AddParmIn madoCmdRvw, "@rvwID", adInteger, 0, mlngCurrentRvwID
    Set madoRsRvw = Server.CreateObject("ADODB.Recordset")
    Call madoRsRvw.Open(madoCmdRvw, , adOpenForwardOnly, adLockReadOnly)
    
    If madoRsRvw.EOF Or madoRsRvw.BOF Then
        'The review not found for some reason:
        mlngCurrentRvwID = -1
    End If
End If
'==============================================================================
' Server-side classes:
'==============================================================================
Sub FillElementObject()
    Dim intI
    Dim strRecord
    
    For intI = 1 To 1000
        
    Next
End Sub

Sub WriteDbError(strLocation, strSource, lngNumber, strDescription, oCmd)
    Dim strErrorMsg
    Dim intI
    Dim strReportErrorTo
    
    strErrorMsg = ""
    strReportErrorTo = ""
    
    On Error Resume Next
    strReportErrorTo = GetAppSetting("ReportErrorTo")
    On Error GoTo 0
    
    strErrorMsg = strErrorMsg & Now & "<br>"
    strErrorMsg = strErrorMsg & "Error [" & strLocation & "]:<br>"
    strErrorMsg = strErrorMsg & strSource & "<br>"
    strErrorMsg = strErrorMsg & lngNumber & " - " & strDescription & "<br><br>"
    If Not oCmd Is Nothing Then
        strErrorMsg = strErrorMsg & oCmd.CommandText & "<br>-------------------------------------------<br>" & vbCrlf
        For intI = 0 To oCmd.Parameters.Count - 1
            strErrorMsg = strErrorMsg & oCmd.Parameters.Item(intI).Name & " = "
            If IsNull(oCmd.Parameters.Item(intI).Value) Then
                strErrorMsg = strErrorMsg & "NULL"
            Else
                strErrorMsg = strErrorMsg & oCmd.Parameters.Item(intI).Value
            End If  
            strErrorMsg = strErrorMsg & "<BR>"          
        Next
    End If
    
    Response.Write "Error encountered while saving the review.&nbsp;&nbsp;Portions of this review may not have been saved.&nbsp;&nbsp;Please copy the error information below" & vbCrLf
    If strReportErrorTo <> "" Then
        Response.Write "&nbsp;" & strReportErrorTo & ".<BR><BR>" & vbCrLf
    Else
        Response.Write ".<BR><BR>" & vbCrLf
    End If
    Response.Write "<INPUT type=""button"" value=""Click To Copy Error Information"" ID=cmdCopyError NAME=""cmdCopyError""><BR><BR>" & vbCrLf
    Response.Write "<SCRIPT LANGUAGE=vbscript>" & vbCrLf
    Response.Write "Sub Window_onload()" & vbCrLf
    Response.Write "   window.parent.lblSavingMessage.style.left=-1000" & vbCrLf
    Response.Write "End Sub" & vbCrLf
    Response.Write "Sub cmdCopyError_onclick()" & vbCrLf
    Response.Write "    blnSetData = window.clipboardData.setData(""Text"",divError.InnerText)" & vbCrLf
    Response.Write "    MsgBox ""The Error Information has been copied to your clipboard. It can now be pasted into an email, Word document, etc.""" & vbCrLf
    Response.Write "End Sub" & vbCrLf
    Response.Write "Sub cmdClose_onclick()" & vbCrLf
    Response.Write "    Call window.parent.CleanClose()" & vbCrLf
    Response.Write "End Sub" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
    Response.Write "<BODY id=PageBody style=""BACKGROUND-COLOR:beige;"">" & vbCrLf
    Response.Write "<DIV id=divError>" & vbCrLf
    Response.Write strErrorMsg
    Response.Write "</DIV><BR>" & vbCrLf
    Response.Write "<INPUT type=""button"" value=""Close"" ID=cmdClose NAME=""cmdClose""><BR><BR>" & vbCrLf
    Response.Write "</BODY>" & vbCrLf
    Response.End
End Sub
%>
<HTML><HEAD>
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
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
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Option Explicit

Sub window_onload
    'window.parent.SaveWindow.style.left = -1000
    'window.parent.divCaseBody.style.left = 1

    window.parent.document.title = "<% = mstrChangesSaved %>"
    ' Update parent Form with save form.    
    window.parent.Form.rvwID.value = <% = mlngCurrentRvwID %>
    window.parent.Form.rvwDateEntered.value = Form.rvwDateEntered.value
    window.parent.Form.rvwMonthYear.value = Form.rvwMonthYear.value
    window.parent.Form.rvwReviewerName.value = Form.rvwReviewerName.value
    window.parent.Form.rvwReviewClassID.value = Form.rvwReviewClassID.value
    window.parent.Form.rvwWorkerID.value = Form.rvwWorkerID.value
    window.parent.Form.rvwWorkerName.value = Form.rvwWorkerName.value
    window.parent.Form.rvwWorkerEmpID.value = Form.rvwWorkerEmpID.value
    window.parent.Form.rvwSupervisorName.value = Form.rvwSupervisorName.value
    window.parent.Form.rvwSupervisorEmpID.value = Form.rvwSupervisorEmpID.value
    window.parent.Form.rvwManagerName.value = Form.rvwManagerName.value
    window.parent.Form.rvwOfficeName.value = Form.rvwOfficeName.value
    window.parent.Form.rvwDirectorName.value = Form.rvwDirectorName.value
    window.parent.Form.rvwCaseLastName.value = Form.rvwCaseLastName.value
    window.parent.Form.rvwCaseFirstName.value = Form.rvwCaseFirstName.value
    window.parent.Form.rvwCaseNumber.value = Form.rvwCaseNumber.value
    window.parent.Form.rvwResponseDueDate.value = Form.rvwResponseDueDate.value
    window.parent.Form.rvwWorkerResponseID.value = Form.rvwWorkerResponseID.value
    window.parent.Form.rvwSubmitted.value = Form.rvwSubmitted.value
    window.parent.Form.ReviewElementData.value = Form.ReviewElementData.value
    window.parent.Form.ReviewCommentData.value = Form.ReviewCommentData.value
    window.parent.Form.ReviewProgramData.value = Form.ReviewProgramData.value
    window.parent.Form.rvwSupSig.value = Form.rvwSupSig.value
    window.parent.Form.rvwSupComments.value = Form.rvwSupComments.value
    window.parent.Form.rvwWrkSig.value = Form.rvwWrkSig.value
    window.parent.Form.rvwWrkComments.value = Form.rvwWrkComments.value
    window.parent.Form.DeleteFail.value = "N"
    If Trim(UCase("<% = mstrAction %>")) = "DELETERECORD" Then
        If "<% = mblnDeletefail %>" = "True" Then
            window.parent.Form.DeleteFail.value = "Y"
        End If
    End If
    window.parent.Form.SaveCompleted.value = "Y"
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:<%=gstrBackColor%>; overflow:scroll" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
<BR><BR><BR>
<div id=divDisplay>
    &nbsp;
</div>
<%

'Write the HTML FORM element to hold and submit the information for the page:
Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY:hidden"" ACTION=""CaseAddEdit.ASP"" ID=Form>" & vbCrLf
WriteFormField "UpdateString", ReqForm("UpdateString")

Call CommonFormFields()

If mlngCurrentRvwID = -1 Then
    Call FormDefBlank()
Else
    Call FormDefEdit()
End If

Response.Write Space(4) & "</FORM>"
%>

</BODY>
</HTML>
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncFormDefBlank.asp"-->
<!--#include file="IncFormDefEdit.asp"-->